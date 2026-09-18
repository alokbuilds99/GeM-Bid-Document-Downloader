"""
GeM Bid Document Downloader
----------------------------
Reads bid numbers from an Excel column, searches each one on
https://bidplus.gem.gov.in/all-bids, opens the matching bid, and downloads
the bid document PDF. Logs the result (success/fail + file path) back into
an output Excel file so you can track what worked.

SETUP (run once):
    pip install selenium webdriver-manager pandas openpyxl

You also need Google Chrome installed on your machine. webdriver-manager
will auto-download the matching chromedriver, so you don't need to install
that separately.

HOW TO RUN:
    # From a plain text file - one bid number per line:
    python download_gem_bids.py --input "bids.txt"

    # From an Excel file - specify the column with bid numbers:
    python download_gem_bids.py --input "bids.xlsx" --column "Bid Number"

CONFIG:
    Edit the constants below (or pass command-line args) to match your file.
"""

import argparse
import os
import re
import time
import traceback
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

SEARCH_URL = "https://bidplus.gem.gov.in/all-bids"
WAIT_SECS = 20
DELAY_BETWEEN_BIDS = (3, 5)  # (min, max) seconds - be polite to the server


def make_safe_filename(bid_number: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]", "_", bid_number.strip())


def build_driver(download_dir: str, headless: bool = False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,1000")
    # Route downloaded PDFs straight into our folder, no "Save As" dialog
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,  # don't open PDF in Chrome viewer
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(60)
    return driver


def initial_setup_and_search(driver, bid_number: str):
    """
    ONE-TIME setup, run only for the very first bid:
      Step 1: load the all-bids page
      Step 2: set the dropdown to "Exact Search" (default is "Contains")
      Step 3: type the bid number into the search box, hit search
      Step 4: check the "Bid/RA Status" checkbox
              -> this auto-unchecks "Ongoing Bids/RA" and the page
                 re-runs the search itself using the current box value

    IMPORTANT: this checkbox + dropdown state stays in place for as long
    as we stay on this page. We must NOT reload the page (driver.get)
    again after this, or the checkbox resets and has to be re-checked -
    that's what search_next_bid() below is for.
    """
    driver.get(SEARCH_URL)
    wait = WebDriverWait(driver, WAIT_SECS)

    # --- Step 2: switch dropdown from "Contains" to "Exact Search" ---
    try:
        dropdown_btn = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[contains(@class,'dropdown-toggle') and "
                "(contains(normalize-space(.), 'Contains') or "
                "contains(normalize-space(.), 'Exact Search'))]"
            ))
        )
        if "exact" not in dropdown_btn.text.strip().lower():
            dropdown_btn.click()
            time.sleep(0.5)
            option = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//*[self::a or self::li or self::button]"
                    "[normalize-space(text())='Exact Search']"
                ))
            )
            option.click()
            time.sleep(0.5)
    except (TimeoutException, NoSuchElementException):
        print("    (warning: could not find/confirm 'Exact Search' "
              "dropdown - continuing anyway, check selector if results "
              "look wrong)")

    # --- Step 3: type bid number into the search box, hit search ---
    search_box = wait.until(
        EC.presence_of_element_located((By.ID, "searchBid"))
    )
    search_box.clear()
    search_box.send_keys(bid_number)
    search_box.send_keys(Keys.RETURN)
    time.sleep(2)  # let results (or "No data found") render

    # --- Step 4: check "Bid/RA Status" checkbox (ONE TIME ONLY) ---
    try:
        checkbox = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//*[contains(normalize-space(text()), 'Bid/RA Status')]"
                "/preceding::input[@type='checkbox'][1]"
            ))
        )
        if not checkbox.is_selected():
            checkbox.click()
            time.sleep(2.5)  # let the site's own re-search finish
    except (TimeoutException, NoSuchElementException):
        print("    (warning: could not find 'Bid/RA Status' checkbox - "
              "continuing anyway, check selector if results look wrong)")


def search_next_bid(driver, bid_number: str):
    """
    For every bid AFTER the first one: just clear the search box, type
    the new bid number, and hit search. Do NOT reload the page (no
    driver.get) - that would wipe out the "Exact Search" + "Bid/RA
    Status" state we set up once in initial_setup_and_search().
    """
    wait = WebDriverWait(driver, WAIT_SECS)
    search_box = wait.until(
        EC.presence_of_element_located((By.ID, "searchBid"))
    )
    search_box.clear()
    search_box.send_keys(bid_number)
    search_box.send_keys(Keys.RETURN)
    time.sleep(2)  # let results (or "No data found") render


def click_bid_number_link_and_wait_for_download(
    driver, bid_number: str, download_dir: str, main_window: str,
    click_timeout=20, download_timeout=30,
):
    """
    Click the "BID NO: <bid_number>" link in the results. This opens a
    NEW TAB that navigates straight to bidplus.gem.gov.in/showbidDocument/
    <internal-id> - that URL itself IS the PDF download (no extra button
    to find on the page). With our Chrome download prefs, the browser
    saves it silently into download_dir.

    Waits for the new file to land, then closes the extra tab and
    switches back to the main window. Returns the path to the downloaded
    file, or None if nothing downloaded in time.
    """
    wait = WebDriverWait(driver, click_timeout)
    before_files = set(os.listdir(download_dir))
    original_handles = set(driver.window_handles)

    bid_link = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//a[normalize-space(text())='{bid_number}']"
        ))
    )
    bid_link.click()
    time.sleep(2)  # give the new tab a moment to open and start downloading

    new_handles = set(driver.window_handles) - original_handles
    if new_handles:
        driver.switch_to.window(new_handles.pop())

    # Poll download_dir for the new file (works regardless of which tab
    # is active, since Chrome downloads go to the same folder either way)
    waited = 0
    new_file = None
    while waited < download_timeout:
        time.sleep(1)
        waited += 1
        after_files = set(os.listdir(download_dir)) - before_files
        finished = [f for f in after_files if not f.endswith(".crdownload")]
        if finished:
            new_file = finished[0]
            break

    # Clean up: close any extra tab, go back to the main window
    for handle in driver.window_handles:
        if handle != main_window:
            driver.switch_to.window(handle)
            driver.close()
    driver.switch_to.window(main_window)

    if not new_file:
        return None

    old_path = os.path.join(download_dir, new_file)
    ext = os.path.splitext(new_file)[1] or ".pdf"
    new_name = make_safe_filename(bid_number) + ext
    new_path = os.path.join(download_dir, new_name)
    counter = 1
    while os.path.exists(new_path):
        new_name = f"{make_safe_filename(bid_number)}_{counter}{ext}"
        new_path = os.path.join(download_dir, new_name)
        counter += 1
    os.rename(old_path, new_path)
    return new_path


def load_bid_numbers(input_path: str, column: str = None):
    """
    Load bid numbers from either a .txt file (one bid number per line) or
    an Excel file (.xlsx) using the given column name. Duplicates are
    removed automatically (order-preserving) so repeated lines don't
    trigger repeated downloads.
    """
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".txt":
        with open(input_path, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f]
        bid_numbers = [line for line in lines if line]  # drop blank lines

    elif ext in (".xlsx", ".xls"):
        if not column:
            raise ValueError(
                "Reading from an Excel file requires --column <header name>"
            )
        df = pd.read_excel(input_path)
        if column not in df.columns:
            raise ValueError(
                f"Column '{column}' not found. Available columns: {list(df.columns)}"
            )
        bid_numbers = df[column].dropna().astype(str).map(str.strip).tolist()

    else:
        raise ValueError(
            f"Unsupported file type '{ext}'. Use a .txt file (one bid number "
            f"per line) or an .xlsx file with --column."
        )

    # De-duplicate while preserving original order
    seen = set()
    deduped = []
    for b in bid_numbers:
        if b not in seen:
            seen.add(b)
            deduped.append(b)

    if len(deduped) < len(bid_numbers):
        print(f"(Removed {len(bid_numbers) - len(deduped)} duplicate bid "
              f"number(s) from the input.)")

    return deduped


def process_bids(input_path: str, column: str, download_dir: str,
                  output_path: str, headless: bool):
    os.makedirs(download_dir, exist_ok=True)

    bid_numbers = load_bid_numbers(input_path, column)
    print(f"Found {len(bid_numbers)} bid numbers to process.\n")

    results = []
    driver = build_driver(download_dir, headless=headless)
    main_window = driver.current_window_handle

    try:
        for i, bid_number in enumerate(bid_numbers, start=1):
            print(f"[{i}/{len(bid_numbers)}] Searching: {bid_number}")
            status = "FAILED"
            file_path = ""
            error_msg = ""

            try:
                if i == 1:
                    # First bid: full setup (page load, Exact Search,
                    # check Bid/RA Status). This state then persists.
                    initial_setup_and_search(driver, bid_number)
                else:
                    # Every bid after: just search, don't reload the page
                    # or touch the checkbox/dropdown again.
                    search_next_bid(driver, bid_number)

                saved_path = click_bid_number_link_and_wait_for_download(
                    driver, bid_number, download_dir, main_window
                )

                if saved_path:
                    status = "SUCCESS"
                    file_path = saved_path
                    print(f"    -> Saved: {saved_path}")
                else:
                    error_msg = "Bid link not found, or download did not complete in time"
                    print(f"    -> {error_msg}")

            except (TimeoutException, NoSuchElementException) as e:
                error_msg = f"Element not found / timeout: {e}"
                print(f"    -> {error_msg}")
            except WebDriverException as e:
                error_msg = f"Browser error: {e}"
                print(f"    -> {error_msg}")
            except Exception as e:
                error_msg = f"Unexpected error: {e}"
                print(f"    -> {error_msg}")
                traceback.print_exc()

            results.append({
                "Bid Number": bid_number,
                "Status": status,
                "File Path": file_path,
                "Error": error_msg,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })

            time.sleep(
                DELAY_BETWEEN_BIDS[0]
                + (DELAY_BETWEEN_BIDS[1] - DELAY_BETWEEN_BIDS[0]) * 0.5
            )

    finally:
        driver.quit()

    results_df = pd.DataFrame(results)
    results_df.to_excel(output_path, index=False)
    print(f"\nDone. Results log saved to: {output_path}")
    print(f"Success: {sum(1 for r in results if r['Status'] == 'SUCCESS')} / {len(results)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download GeM bid PDFs listed in an Excel column")
    parser.add_argument("--input", required=True, help="Path to input .txt (one bid per line) or .xlsx file")
    parser.add_argument("--column", default=None, help="Column name containing bid numbers (required only for .xlsx input)")
    parser.add_argument("--download-dir", default="downloaded_bids", help="Folder to save PDFs into")
    parser.add_argument("--output", default="bid_download_results.xlsx", help="Path for results log Excel file")
    parser.add_argument("--headless", action="store_true", help="Run Chrome headless (no visible window)")
    args = parser.parse_args()

    process_bids(
        input_path=args.input,
        column=args.column,
        download_dir=args.download_dir,
        output_path=args.output,
        headless=args.headless,
    )
