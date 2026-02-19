# Website Grader

A Python 3.11 script set for grading business websites and exporting CSV reports.

## Features

- Grades a predefined list of URLs from `websites.txt`
- Finds local businesses from OpenStreetMap (Overpass) and grades discovered websites
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

## Run: find and grade plumbers from OpenStreetMap

```bash
python find_and_grade.py --city "Austin, TX"
```

Optional arguments:

- `--limit` (default: `50`)
- `--radius_km` (default: `20`)
- `--terms` (default: `plumber,plumbing,drain,sewer,water heater,pipe,rooter`)
- `--overpass_timeout` (default: `180`)

Example with custom radius and terms:

```bash
python find_and_grade.py --city "Austin, TX" --limit 100 --radius_km 25 --terms "plumbing,drain,sewer"
```

For larger lead counts (for example `--limit 200`), tune `--radius_km` to keep Overpass queries focused and reduce timeout risk.

Example (200 leads):

```bash
python find_and_grade.py --city "Austin, TX" --limit 200 --radius_km 20
```

Outputs:

- `leads_raw.csv`: raw OSM lead fields (`name`, `website`, `phone`, `address`, `lat`, `lon`, `osm_type`, `osm_id`)
- `report.csv`: website grading report for leads that have websites, with leading source columns:
  - `business_name`, `phone`, `address`, `lat`, `lon`, `source`, `source_id`
  - followed by grading columns (`url`, `reachable`, `https`, `status_code`, etc.)
  - includes outreach columns: `opportunity_score_0_100` and `pitch`
- `priority_leads.csv`: outreach-ready subset of actionable leads, filtered and sorted by `opportunity_score_0_100` (highest first). Start calling from this file first when planning outreach.

The OSM pipeline applies basic deduping by:

- same website URL, or
- same business name + address.

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
