# Website Grader

A Python 3.11 script set for grading business websites and exporting CSV reports.

## Features

- Grades a predefined list of URLs from `websites.txt`
- Finds local businesses from OpenStreetMap (Overpass) and optionally Yelp/Bing local business search, then grades discovered websites
- Uses Nominatim geocoding to convert a city name into a center point for radius-based Overpass search
- Parses HTML with `beautifulsoup4`
- If a contact page link is found on the homepage, fetches that contact page too
- Uses homepage + contact page content together for lead signals:
  - phone presence
  - email presence
  - contact signal detection
- Produces deterministic output with concurrent grading via `ThreadPoolExecutor(max_workers=8)`

## Requirements

- Python 3.11

## Setup

Install project dependencies before running the pipelines (including XLSX export via `openpyxl`):

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run: grade a static website list

```bash
python main.py
```

After running, `report.csv` and `priority_leads.csv` will be created in the project directory. Use `priority_leads.csv` for first-pass outreach.

If Google Sheets export is not configured, the run continues normally and prints:

`Google Sheets export skipped (missing GOOGLE_SHEETS_ID / GOOGLE_SERVICE_ACCOUNT_JSON)`

## Run: find and grade plumbers from OpenStreetMap

```bash
python find_and_grade.py --city "Austin, TX"
```

Optional arguments:

- `--limit` (default: `50`)
- `--radius_km` (default: `25`)
- `--terms` (default: `plumber,plumbing,drain,sewer,rooter,pipe,water heater`)
- `--overpass_timeout` (default: `180`)

Example with custom radius and terms:

```bash
python find_and_grade.py --city "Austin, TX" --limit 100 --radius_km 30
python find_and_grade.py --city "Austin, TX" --limit 100 --radius_km 30 --terms "plumbing,drain,sewer"
```

Optional lead source environment variables (default keeps OSM-only behavior):

```bash
export LEAD_SOURCES="osm"                # default
export LEAD_SOURCES="osm,yelp,bing"      # enable optional API sources
export YELP_API_KEY="your_yelp_api_key"  # optional, Yelp skipped if missing
export BING_API_KEY="your_bing_api_key"  # optional, Bing skipped if missing
```

If an API key is missing, that source is skipped and the pipeline continues without crashing.

For larger lead counts (for example `--limit 200`), tune `--radius_km` to keep Overpass queries focused and reduce timeout risk.

Example (200 leads):

```bash
python find_and_grade.py --city "Austin, TX" --limit 200 --radius_km 25
```

Outputs:

- `leads_raw.csv`: raw lead fields including source metadata (`name`, `business_name`, `website`, `website_url`, `phone`, `address`, `lat`, `lon`, `source`, `sources`, `source_listing_url`, `source_id`, plus OSM-specific fields when applicable)
- `report.csv`: website grading report for leads that have websites, with leading source columns:
  - `business_name`, `phone`, `address`, `lat`, `lon`, `source`, `sources`, `source_listing_url`, `source_id`
  - followed by grading columns (`url`, `reachable`, `https`, `status_code`, etc.)
  - includes outreach columns: `opportunity_score_0_100` and `pitch`
- `priority_leads.csv`: outreach-ready subset of actionable leads, filtered and sorted by `opportunity_score_0_100` (highest first). Start calling from this file first when planning outreach.

Lead deduping is applied before grading with this priority:

- normalized website domain,
- else normalized phone number,
- else normalized business name + address.

When duplicates are merged from multiple discovery sources, `sources` records all contributing sources (semicolon-separated).

## Input format for `main.py`

Edit `websites.txt` to include one URL per line:

```text
https://example.com
https://your-business-site.com
```

Blank lines and comment lines starting with `#` are ignored.

## Scoring rubric (0-100)

- Reachable website: +20
- Uses HTTPS in final URL: +15
- HTTP status in 200-399: +10
- Page title present: +10
- Meta description present: +10
- Phone detected (homepage/contact page): +10
- Email detected (homepage/contact page): +10
- Contact page link detected: +15

## Generate outreach email copy in Excel

If you already have `outreach_sheet.xlsx` with prioritized leads, you can append personalized outreach copy columns:

```bash
python enhance_outreach_sheet.py
```

This adds/updates `email_subject` and `email_body` columns, applies tier-based templates from `opportunity_score_0_100`, preserves existing columns, and freezes the header row.

## Optional: export CSV outputs to Google Sheets

Both pipelines can optionally export generated CSV files into a Google Spreadsheet using the official Google API client libraries (`gspread` + `google-auth`).

### 1) Create a Google Cloud service account

1. Open Google Cloud Console and create/select a project.
2. Enable **Google Sheets API** and **Google Drive API** for that project.
3. Go to **IAM & Admin > Service Accounts**, create a service account, and create a JSON key.
4. Download the key file and store it securely (for example: `./secrets/service-account.json`).

### 2) Share the target spreadsheet

- Open your target Google Sheet and click **Share**.
- Add the service account email (from the JSON key, `client_email`) as an editor.

### 3) Set environment variables

Required:

```bash
export GOOGLE_SHEETS_ID="your_spreadsheet_id"
export GOOGLE_SERVICE_ACCOUNT_JSON="./secrets/service-account.json"
```

Optional tab names (defaults shown):

```bash
export GOOGLE_SHEETS_TAB_REPORT="report"
export GOOGLE_SHEETS_TAB_PRIORITY="priority_leads"
export GOOGLE_SHEETS_TAB_LEADS_RAW="leads_raw"
```

### 4) Run the pipeline as usual

- `python main.py` exports existing `report.csv` and `priority_leads.csv` (and `leads_raw.csv` if present).
- `python find_and_grade.py --city "Austin, TX"` exports `report.csv`, `priority_leads.csv`, and `leads_raw.csv`.

For each exported CSV, the script will:
- create the tab if missing,
- clear existing tab contents,
- write all CSV rows including headers,
- freeze the header row.

A success line is printed per tab, for example:

`Exported report.csv -> tab 'report'`
