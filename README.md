# Website Grader

A Python 3.11 script that grades a list of business websites and exports a CSV report.

## Features

- Reads URLs from `websites.txt` (one URL per line)
- Fetches each homepage with `requests` (timeout + redirect support)
- Parses HTML with `beautifulsoup4`
- If a contact page link is found on the homepage, fetches that contact page too
- Uses homepage + contact page content together for lead signals:
  - phone presence
  - email presence
  - contact signal detection
- Produces deterministic output with concurrent grading via `ThreadPoolExecutor(max_workers=8)`
- Produces `report.csv` with scores from 0 to 100 and a `reasons` column for missing items

## Requirements

- Python 3.11

## Setup

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

After running, `report.csv` will be created in the project directory.

## Input format

Edit `websites.txt` to include one URL per line:

```text
https://example.com
https://your-business-site.com
```

Blank lines and comment lines starting with `#` are ignored.

## Output columns

`report.csv` contains:

- `url`
- `reachable`
- `https`
- `status_code`
- `final_url`
- `title`
- `meta_description_present`
- `has_phone`
- `has_email`
- `has_contact_page_link`
- `contact_page_url`
- `contact_page_reachable`
- `notes`
- `reasons` (semicolon-separated missing items, such as `missing_meta_description; missing_email`)
- `score_0_100`

## Scoring rubric (0-100)

- Reachable website: +20
- Uses HTTPS in final URL: +15
- HTTP status in 200-399: +10
- Page title present: +10
- Meta description present: +10
- Phone detected (homepage/contact page): +10
- Email detected (homepage/contact page): +10
- Contact page link detected: +15
