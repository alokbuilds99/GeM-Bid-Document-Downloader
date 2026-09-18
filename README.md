# GeM Bid Document Downloader

Python automation tool to search GeM bids, download bid document PDFs, and log results in Excel.

## Features
- Search GeM bids using bid numbers
- Supports TXT and Excel input
- Downloads bid PDFs using Selenium
- Logs success/failure, file path, errors, and timestamps
- Removes duplicate bid numbers
- Supports headless Chrome

## Installation

```bash
pip install -r requirements.txt
```

Google Chrome must be installed.

## Usage

From a text file:

```bash
python download_gem_bids.py --input "bids.txt"
```

From an Excel file:

```bash
python download_gem_bids.py --input "bids.xlsx" --column "Bid Number"
```

Headless mode:

```bash
python download_gem_bids.py --input "bids.xlsx" --column "Bid Number" --headless
```

## Output
- Downloaded PDFs are saved in `downloaded_bids/`
- Results are saved in `bid_download_results.xlsx`

## Disclaimer
Use this project responsibly and in accordance with GeM terms and applicable policies.
