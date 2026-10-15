# ISC West Exhibitor Scraper

Scrapes company profiles from the ISC West exhibitor directory
(`discoverisc.com`) and writes them to `.xlsx`.

Input is a text file with one company name per line. For each name the
scraper searches the directory, opens the first `exhibitor-details` result,
and extracts the profile fields.

## Project structure

- `main.py` — scraper. Queue from `companies.txt`, threaded Botasaurus
  browsers, `write_xlsx()` output.
- `companies.txt` — one company per line. Scraping stops when the queue is
  empty; blank lines are ignored.
- `ICS West Output.xlsx` — reference sample (expected format).

## Prerequisites

- Python >= 3.13
- Brave (`/usr/bin/brave` on Linux) or edit `chrome_executable_path` in
  `main.py` (For Better Anti-detection)
- Dependencies : `botasaurus`, `botasaurus-driver`,
  `colorama`
- `openpyxl` is optional. If installed it is used for `.xlsx`; otherwise
  `main.py` falls back to a stdlib writer.

## Install

```bash
uv sync
```

## Usage

1. Put company names in `companies.txt`, one per line:

```text
Napco Security Technologies Inc.
Alarm Lock Systems, Inc.
```

1. Run:

```bash
uv run python main.py
```

1. Enter the thread count when prompted (`Enter Number Of Threading :`,
   clamped to 1–10). Each thread consumes names from the shared queue until
   it is empty.

2. Open `ICS West Output.xlsx`.

To change paths, edit in `main.py`: `COMPANIES_FILE`, `OUTPUT_XLSX`.

## Behavior notes

- Search URL per company:
  `.../exhibitor-directory.html?locale=en-US&query={company_name}`.
  The scraper polls up to ~30 s for `a[href*="exhibitor-details."]` and
  opens the first match via direct GET (click fallback).
- The directory search component is slow/flaky; a company with no rendered
  result is logged and written as a row with the name and search URL only.
- Detail pages wait for `#exhibitor-details h1.wrap-word` (~20 s) before
  extracting. Missing optional fields (brands, LinkedIn, email) are left
  empty rather than failing the row.
- Brave background/telemetry hosts are blocked via `BLOCKED_URL_PATTERNS`
  in `main.py`.
