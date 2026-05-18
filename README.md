# CongressTrader

CongressTrader is a Python research script that collects congressional stock trade disclosures, enriches ticker symbols with sector and industry data, and exports a dated Google Sheets report.

The project is designed for personal research, transparency work, and repeatable reporting from public disclosure sources.

## Features

- Collects House PTR filings from the House Clerk disclosure site
- Attempts to collect Senate PTR filings from Senate eFD/eFTS sources
- Attempts to fetch additional trade data from the Capitol Trades public API
- Parses PDF disclosure documents for stock purchases and sales
- Deduplicates trades across sources
- Enriches tickers with sector and industry data from Yahoo Finance
- Creates a Google Sheets report with separate tabs for all trades, purchases, sales, top buys, summary by ticker, and newly appearing tickers
- Can be scheduled with cron for weekly report generation

## Output

Each run creates a dated Google Sheet with tabs similar to:

- `All Trades`
- `All Purchases`
- `All Sales`
- `Top 50 Buys (Held 30d+)`
- `Summary by Ticker`
- `New On Week`

## Data Sources

CongressTrader uses publicly available sources, including:

- House Clerk financial disclosure filings
- Senate electronic financial disclosure filings, when available
- Capitol Trades public API, when available
- Yahoo Finance ticker metadata via `yfinance`

Availability and formatting of public disclosure data may change over time, so scraper behavior may require maintenance.

## Requirements

- Python 3.10 or newer recommended
- A Google account
- A Google Cloud OAuth client configured for a desktop application
- Access to Google Sheets and Google Drive APIs

Python dependencies:

```txt
gspread
yfinance
google-auth
google-auth-oauthlib
google-api-python-client
pypdf
openpyxl
```

## Setup

Install dependencies:

```bash
pip install -r requirements.txt
```

Create a Google Cloud OAuth desktop client and download the client secrets JSON file. Save it in the project folder as:

```text
congress-trades-oauth.json
```

On first run, the script opens a browser-based Google authorization flow and creates a local token file:

```text
congress-trades-token.json
```

Do not commit either of these files to git.

## Usage

Run the script:

```bash
python3 congress_trader.py
```

The script will scrape available disclosure data, create a dated Google Sheet, and print the resulting sheet URL.

## Optional Automation

The script can be scheduled with cron. For example, to run every Monday at 9:00 PM:

```cron
0 21 * * 1 /usr/bin/python3 /path/to/CongressTrader/congress_trader.py >> /path/to/CongressTrader/run.log 2>&1
```

Adjust the paths for your own machine.

## Security Notes

Never commit local credentials, OAuth tokens, generated logs, or cache folders. Recommended ignored files include:

```gitignore
.DS_Store
__pycache__/
*.pyc
_pdf_cache/
run.log
congress-trades-token.json
congress-trades-oauth.json
.env
```

If credentials are accidentally committed, revoke or rotate them immediately.

## Project Status

This is an early public release of a personal research automation script. It is not financial advice, a trading system, or an official government data product.

## Disclaimer

CongressTrader is provided for research and educational purposes only. Congressional disclosure data can contain reporting delays, amendments, formatting inconsistencies, and data quality issues. Verify important findings against the original disclosure documents.
