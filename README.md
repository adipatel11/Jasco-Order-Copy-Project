# Jasco Order Automation Tool

A Python automation tool that eliminates manual data entry by scraping retail order data from a government web portal and exporting it to a formatted, ready-to-use Excel spreadsheet.

## Overview

This tool was built to automate a repetitive, time-consuming workflow: manually copying "Reserve Inventory" orders from the Mississippi Department of Revenue's TAP portal into a spreadsheet for processing. The script handles login (including two-factor authentication), navigates and paginates through order records, extracts item-level data, and writes it to Excel — all with zero manual copy-pasting.

## Features

- **Automated login + 2FA** — handles credential entry and security code prompts securely using `pwinput` (masked input)
- **Smart order filtering** — queries orders by date and status (`Reserve Inventory`) directly via the portal's filter system
- **Full pagination support** — iterates across all result pages to ensure no orders are missed
- **Excel export with formatting** — writes order data to a structured `.xlsx` file with proper fonts, date formatting, and auto-populated VLOOKUP formulas for item size lookup
- **Duplicate detection and removal** — automatically deduplicates rows before saving
- **Multi-date batch processing** — loop over multiple dates in a single session without restarting

## Tech Stack

| Tool | Purpose |
|---|---|
| Python | Core scripting language |
| Selenium | Browser automation and web scraping |
| openpyxl | Excel file creation and formatting |
| pwinput | Secure masked password input |
| ChromeDriver | Headless-compatible Chrome interface |

## Output

Each row written to `orders.xlsx` contains:

| Item # | Item Name | Size (VLOOKUP) | Reserved Qty | Order # | Date |
|---|---|---|---|---|---|

The spreadsheet includes:
- Auto-formatted date column
- VLOOKUP formula for item size using a reference data sheet
- Bold total quantity formula in the header row
- Consistent font styling across all columns

## Setup

**Prerequisites:** Python 3.11+, Google Chrome, ChromeDriver (included)

```bash
# Install dependencies
pip install -r requirements.txt

# Run the tool
python jasco.py
```

On launch, you will be prompted for:
1. Your TAP portal username and password
2. A 6-digit two-factor authentication code
3. The month and day to pull orders for

You can process multiple dates in one session.

## How It Works

1. Opens the TAP portal and logs in with credentials + 2FA
2. Navigates to the retail orders section
3. Filters orders by the provided date and `Reserve Inventory` status
4. Iterates through all result pages, clicking into each order
5. Scrapes item number, name, quantity, order ID, and date
6. Writes each item as a row in `orders.xlsx`
7. After all dates are processed, removes duplicates, applies formatting, and saves

## Skills Demonstrated

- Web automation with Selenium (dynamic waits, CSS selectors, pagination, JS execution)
- Data pipeline design: scrape → transform → export
- Excel manipulation and formula injection with openpyxl
- Robust input validation and error handling for real-world portal instability
- Process automation that saved significant manual data entry time in a production workflow
