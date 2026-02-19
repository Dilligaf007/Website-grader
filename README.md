# Website Grader

A Python 3.11 script set for grading business websites and exporting CSV reports.

## Features

- Grades a predefined list of URLs from `websites.txt`
- Finds local businesses from OpenStreetMap (Overpass) and grades discovered websites
- Uses Nominatim geocoding to convert a city name into a search bounding box
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

After running, `report.csv` will be created in the project directory.

## Run: find and grade plumbers from OpenStreetMap

```bash
python find_and_grade.py --city "Austin, TX"
```

Optional arguments:

- `--query` (default: `plumber`)
- `--limit` (default: `50`)

Example with custom query and limit:

```bash
python find_and_grade.py --city "Seattle, WA" --query "plumber" --limit 25
```

Outputs:

- `leads_raw.csv`: raw OSM lead fields (`name`, `website`, `phone`, `address`, `lat`, `lon`, `osm_type`, `osm_id`)
- `report.csv`: website grading report for leads that have websites, with leading source columns:
  - `business_name`, `phone`, `address`, `lat`, `lon`, `source`, `source_id`
  - followed by grading columns (`url`, `reachable`, `https`, `status_code`, etc.)
- `priority_leads.csv`: filtered high-opportunity leads from `report.csv` where `reachable=True`, `has_phone=True`, URL is present, and `opportunity_score_0_100 > 0`; sorted by `opportunity_score_0_100` descending and limited to outreach-focused columns (`business_name`, `phone`, `url`, `opportunity_score_0_100`, `score_0_100`, `reasons`, `pitch`).

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
