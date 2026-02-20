import argparse
import csv
import re
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Optional, Set, Tuple

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from main import MAX_WORKERS, PRIORITY_OUTPUT_FILE, grade_website

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.openstreetmap.ru/api/interpreter",
]
USER_AGENT = "WebsiteGrader-FindAndGrade/1.0 (contact: nathan.strickland.consult@gmail.com)"
NOMINATIM_EMAIL = "nathan.strickland.consult@gmail.com"
REQUEST_DELAY_SECONDS = 1.0
OVERPASS_MAX_RETRIES = 3
OVERPASS_BACKOFF_SECONDS = 2.0
RAW_OUTPUT_FILE = "leads_raw.csv"
GRADED_OUTPUT_FILE = "report.csv"
PRIORITY_LEADS_OUTPUT_FILE = "priority_leads.csv"
OUTREACH_SHEET_OUTPUT_FILE = "outreach_sheet.xlsx"

WEBSITE_KEYS = ["website", "contact:website", "url"]
PHONE_KEYS = ["phone", "contact:phone"]
DEFAULT_TERMS = ["plumber", "plumbing", "drain", "sewer", "rooter", "pipe", "water heater"]


def clean_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def get_city_center(city: str) -> Tuple[float, float]:
    time.sleep(1)
    params = {"q": city, "format": "json", "limit": 1, "email": NOMINATIM_EMAIL}
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "en"}
    response = requests.get(
        NOMINATIM_URL,
        params=params,
        headers=headers,
        timeout=20,
    )
    if response.status_code == 403:
        request_url = requests.Request("GET", NOMINATIM_URL, params=params).prepare().url
        print("Nominatim 403: check USER_AGENT and email param")
        print(f"Requested URL: {request_url}")
    response.raise_for_status()

    payload = response.json()
    if not payload:
        raise ValueError(f"Could not geocode city: {city}")

    lat = payload[0].get("lat")
    lon = payload[0].get("lon")
    if lat is None or lon is None:
        raise ValueError(f"City geocoding did not return a valid center point: {city}")

    return float(lat), float(lon)


def build_name_regex(terms: List[str]) -> str:
    escaped_terms = [re.escape(term.strip()) for term in terms if term.strip()]
    if not escaped_terms:
        escaped_terms = [re.escape(term) for term in DEFAULT_TERMS]
    return "|".join(escaped_terms)


def build_overpass_query(lat: float, lon: float, radius_km: float, terms: List[str], overpass_timeout: int) -> str:
    name_regex = build_name_regex(terms).replace('"', '\\"')
    radius_m = int(radius_km * 1000)
    return f"""
[out:json][timeout:{overpass_timeout}];
(
  nwr(around:{radius_m},{lat},{lon})["craft"="plumber"];
  nwr(around:{radius_m},{lat},{lon})["shop"="plumbing"];
  nwr(around:{radius_m},{lat},{lon})["name"~"{name_regex}",i];
  nwr(around:{radius_m},{lat},{lon})["description"~"{name_regex}",i];
  nwr(around:{radius_m},{lat},{lon})["brand"~"{name_regex}",i];
  nwr(around:{radius_m},{lat},{lon})["operator"~"{name_regex}",i];
);
out center tags;
""".strip()


def fetch_overpass_elements(overpass_query: str, overpass_timeout: int) -> List[Dict[str, object]]:
    last_exception: Optional[Exception] = None
    for endpoint in OVERPASS_ENDPOINTS:
        for attempt in range(1, OVERPASS_MAX_RETRIES + 1):
            try:
                response = requests.post(
                    endpoint,
                    data=overpass_query,
                    headers={"User-Agent": USER_AGENT},
                    timeout=overpass_timeout + 30,
                )
                response.raise_for_status()
                return response.json().get("elements", [])
            except (requests.RequestException, ValueError) as exc:
                last_exception = exc
                if attempt < OVERPASS_MAX_RETRIES:
                    backoff_seconds = OVERPASS_BACKOFF_SECONDS * (2 ** (attempt - 1))
                    print(
                        f"Overpass request failed at {endpoint} (attempt {attempt}/{OVERPASS_MAX_RETRIES}). "
                        f"Retrying in {backoff_seconds:.1f}s..."
                    )
                    time.sleep(backoff_seconds)
                else:
                    print(f"Overpass endpoint exhausted: {endpoint}")

    if last_exception is not None:
        raise RuntimeError("All Overpass endpoints failed after retries.") from last_exception
    raise RuntimeError("All Overpass endpoints failed without a specific exception.")


def first_tag(tags: Dict[str, str], keys: List[str]) -> str:
    for key in keys:
        value = tags.get(key, "").strip()
        if value:
            return value
    return ""


def build_address(tags: Dict[str, str]) -> str:
    if tags.get("addr:full"):
        return clean_whitespace(tags["addr:full"])

    parts = [
        tags.get("addr:housenumber", "").strip(),
        tags.get("addr:street", "").strip(),
        tags.get("addr:city", "").strip(),
        tags.get("addr:state", "").strip(),
        tags.get("addr:postcode", "").strip(),
    ]
    parts = [part for part in parts if part]
    return clean_whitespace(", ".join(parts))


def parse_element(element: Dict[str, object]) -> Dict[str, object]:
    tags = element.get("tags", {}) or {}
    tags = {str(k): str(v) for k, v in tags.items()}

    lat = element.get("lat")
    lon = element.get("lon")
    center = element.get("center") or {}
    if lat is None:
        lat = center.get("lat", "")
    if lon is None:
        lon = center.get("lon", "")

    osm_type = str(element.get("type", ""))
    osm_id = element.get("id", "")

    return {
        "name": tags.get("name", "").strip(),
        "website": first_tag(tags, WEBSITE_KEYS),
        "phone": first_tag(tags, PHONE_KEYS),
        "address": build_address(tags),
        "lat": lat,
        "lon": lon,
        "osm_type": osm_type,
        "osm_id": osm_id,
    }


def dedupe_leads(leads: List[Dict[str, object]]) -> List[Dict[str, object]]:
    seen_websites: Set[str] = set()
    seen_name_address: Set[Tuple[str, str]] = set()
    deduped: List[Dict[str, object]] = []

    for lead in leads:
        website = str(lead.get("website", "")).strip().lower().rstrip("/")
        name = str(lead.get("name", "")).strip().lower()
        address = str(lead.get("address", "")).strip().lower()

        duplicate = False
        if website:
            if website in seen_websites:
                duplicate = True
            else:
                seen_websites.add(website)

        name_address = (name, address)
        if name and address:
            if name_address in seen_name_address:
                duplicate = True
            else:
                seen_name_address.add(name_address)

        if not duplicate:
            deduped.append(lead)

    return deduped


def sort_leads(leads: List[Dict[str, object]]) -> List[Dict[str, object]]:
    return sorted(
        leads,
        key=lambda lead: (
            str(lead.get("name", "")).strip().lower(),
            str(lead.get("address", "")).strip().lower(),
            str(lead.get("website", "")).strip().lower(),
            str(lead.get("osm_type", "")).strip().lower(),
            str(lead.get("osm_id", "")).strip(),
        ),
    )


def find_businesses(
    city: str,
    terms: List[str],
    limit: int,
    radius_km: float,
    overpass_timeout: int,
) -> List[Dict[str, object]]:
    lat, lon = get_city_center(city)
    time.sleep(REQUEST_DELAY_SECONDS)

    overpass_query = build_overpass_query(lat, lon, radius_km, terms, overpass_timeout)
    elements = fetch_overpass_elements(overpass_query, overpass_timeout)
    leads = [parse_element(element) for element in elements]
    deduped = dedupe_leads(leads)
    ordered = sort_leads(deduped)
    return ordered[:limit]


def parse_terms(terms_arg: str) -> List[str]:
    parsed = [term.strip() for term in terms_arg.split(",") if term.strip()]
    return parsed or DEFAULT_TERMS.copy()


def write_raw_leads(path: str, leads: List[Dict[str, object]]) -> None:
    fieldnames = ["name", "website", "phone", "address", "lat", "lon", "osm_type", "osm_id"]
    with open(path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(leads)


def build_graded_row(lead: Dict[str, object], graded: Dict[str, object]) -> Dict[str, object]:
    return {
        "business_name": lead.get("name", ""),
        "phone": lead.get("phone", ""),
        "address": lead.get("address", ""),
        "lat": lead.get("lat", ""),
        "lon": lead.get("lon", ""),
        "source": lead.get("osm_type", ""),
        "source_id": lead.get("osm_id", ""),
        **graded,
    }


def write_graded_report(path: str, rows: List[Dict[str, object]]) -> None:
    ordered_fieldnames = [
        "business_name",
        "phone",
        "address",
        "lat",
        "lon",
        "source",
        "source_id",
        "url",
        "reachable",
        "https",
        "status_code",
        "final_url",
        "title",
        "meta_description_present",
        "has_phone",
        "has_email",
        "primary_email",
        "emails",
        "has_contact_page_link",
        "contact_page_url",
        "contact_page_reachable",
        "notes",
        "reasons",
        "score_0_100",
    ]

    for row in rows:
        row.setdefault("opportunity_score_0_100", 0)
        row.setdefault("pitch", "")

    if not rows:
        fieldnames = [
            *ordered_fieldnames,
            "opportunity_score_0_100",
            "pitch",
        ]
    else:
        fieldnames = [
            *ordered_fieldnames,
            *[
                field
                for field in ("opportunity_score_0_100", "pitch")
                if field not in ordered_fieldnames
            ],
        ]

    with open(path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def is_true(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() == "true"
    return False


def as_float(value: object) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except ValueError:
            return 0.0
    return 0.0


def build_priority_row(row: Dict[str, object]) -> Dict[str, object]:
    return {
        "business_name": row.get("business_name", ""),
        "phone": row.get("phone", ""),
        "primary_email": row.get("primary_email", ""),
        "emails": row.get("emails", ""),
        "url": row.get("url", ""),
        "opportunity_score_0_100": row.get("opportunity_score_0_100", ""),
        "score_0_100": row.get("score_0_100", ""),
        "reasons": row.get("reasons", ""),
        "pitch": row.get("pitch", ""),
    }


def get_priority_rows(rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    filtered_rows = [
        row
        for row in rows
        if is_true(row.get("reachable"))
        and is_true(row.get("has_phone"))
        and str(row.get("url", "")).strip()
        and as_float(row.get("opportunity_score_0_100", 0)) > 0
    ]
    return sorted(
        filtered_rows,
        key=lambda row: as_float(row.get("opportunity_score_0_100", 0)),
        reverse=True,
    )


def normalize_rows(rows: object) -> List[Dict[str, object]]:
    if isinstance(rows, dict):
        values = list(rows.values())
        if all(isinstance(value, dict) for value in values):
            return values
        example_value = values[0] if values else None
        raise TypeError(
            "write_priority_leads expected rows to be a list of dicts (or a dict of dicts); "
            f"got dict with value type {type(example_value).__name__}. Example value: {example_value!r}"
        )

    if isinstance(rows, list):
        if all(isinstance(row, dict) for row in rows):
            return rows
        example_row = rows[0] if rows else None
        raise TypeError(
            "write_priority_leads expected rows to be a list of dicts; "
            f"got list with element type {type(example_row).__name__}. Example element: {example_row!r}"
        )

    raise TypeError(
        "write_priority_leads expected rows to be a list of dicts (or a dict of dicts); "
        f"got {type(rows).__name__}. Example value: {rows!r}"
    )


def write_priority_leads(path: str, rows: List[Dict[str, object]]) -> int:
    normalized_rows = normalize_rows(rows)
    sorted_rows = get_priority_rows(normalized_rows)

    fieldnames = [
        "business_name",
        "phone",
        "primary_email",
        "emails",
        "url",
        "opportunity_score_0_100",
        "score_0_100",
        "reasons",
        "pitch",
    ]

    with open(path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(build_priority_row(row) for row in sorted_rows)

    return len(sorted_rows)


def clean_specific_issue(reasons: object, pitch: object) -> str:
    first_reason = str(reasons or "").split(";")[0].strip()
    if first_reason:
        return first_reason.replace("_", " ")
    return str(pitch or "website improvements").strip()


def build_email_content(row: Dict[str, object]) -> Tuple[str, str]:
    score = as_float(row.get("score_0_100", 0))
    business_name = str(row.get("business_name", "your business")).strip() or "your business"
    specific_issue = clean_specific_issue(row.get("reasons", ""), row.get("pitch", ""))

    if score >= 75:
        tier = "Tier A"
    elif score >= 50:
        tier = "Tier B"
    else:
        tier = "Tier C"

    subject = f"Quick win idea for {business_name} ({tier})"
    body = (
        f"Hi {business_name} team,\n\n"
        "I took a quick look at your website and noticed one specific issue: "
        f"{specific_issue}.\n\n"
        f"You're currently in {tier} in our grading pass, and I can share a short plan to fix this quickly "
        "and improve inbound lead conversion.\n\n"
        "Would you like me to send over a 3-step recommendation?"
    )
    return subject, body


def write_outreach_sheet(path: str, prioritized_rows: List[Dict[str, object]], city_name: str) -> None:
    headers = [
        "business_name",
        "city",
        "phone",
        "primary_email",
        "emails",
        "url",
        "opportunity_score_0_100",
        "score_0_100",
        "reasons",
        "pitch",
        "email_subject",
        "email_body",
        "email_sent",
        "followup_1_sent",
        "followup_2_sent",
        "response",
        "status",
    ]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Outreach"
    sheet.append(headers)

    for row in prioritized_rows:
        email_subject, email_body = build_email_content(row)
        sheet.append(
            [
                row.get("business_name", ""),
                city_name,
                row.get("phone", ""),
                row.get("primary_email", ""),
                row.get("emails", ""),
                row.get("url", ""),
                row.get("opportunity_score_0_100", ""),
                row.get("score_0_100", ""),
                row.get("reasons", ""),
                row.get("pitch", ""),
                email_subject,
                email_body,
                "",
                "",
                "",
                "",
                "",
            ]
        )

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, _ in enumerate(headers, start=1):
        column_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in sheet[column_letter]:
            value_len = len(str(cell.value)) if cell.value is not None else 0
            if value_len > max_len:
                max_len = value_len
        sheet.column_dimensions[column_letter].width = min(max(max_len + 2, 12), 80)

    workbook.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Find and grade local businesses from OpenStreetMap.")
    parser.add_argument("--city", required=True, help="City name to search (e.g. 'Austin, TX').")
    parser.add_argument("--limit", type=int, default=50, help="Maximum number of unique leads (default: 50).")
    parser.add_argument(
        "--radius_km",
        type=float,
        default=25,
        help="Search radius around city center in kilometers (default: 25).",
    )
    parser.add_argument(
        "--terms",
        default=",".join(DEFAULT_TERMS),
        help=(
            "Comma-separated terms for case-insensitive business name matching "
            "(default: plumber,plumbing,drain,sewer,rooter,pipe,water heater)."
        ),
    )
    parser.add_argument(
        "--overpass_timeout",
        type=int,
        default=180,
        help="Overpass query timeout in seconds (default: 180).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    terms = parse_terms(args.terms)

    leads = find_businesses(
        city=args.city,
        terms=terms,
        limit=args.limit,
        radius_km=args.radius_km,
        overpass_timeout=args.overpass_timeout,
    )
    write_raw_leads(RAW_OUTPUT_FILE, leads)

    leads_with_websites = [lead for lead in leads if str(lead.get("website", "")).strip()]

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        graded_rows = list(executor.map(lambda lead: grade_website(str(lead["website"])), leads_with_websites))

    report_rows = [
        build_graded_row(lead, graded)
        for lead, graded in zip(leads_with_websites, graded_rows)
    ]
    write_graded_report(GRADED_OUTPUT_FILE, report_rows)
    priority_count = write_priority_leads(path=PRIORITY_OUTPUT_FILE, rows=report_rows)
    prioritized_rows = get_priority_rows(report_rows)
    city_name = str(args.city).split(",", maxsplit=1)[0].strip()
    write_outreach_sheet(OUTREACH_SHEET_OUTPUT_FILE, prioritized_rows, city_name)

    missing_phone_count = sum(1 for row in report_rows if not str(row.get("phone", "")).strip())

    print(f"Found {len(leads)} unique lead(s). Raw leads saved to {RAW_OUTPUT_FILE}.")
    print(
        f"Graded {len(report_rows)} lead(s) with websites. Report saved to {GRADED_OUTPUT_FILE}."
    )
    print(f"Created {PRIORITY_OUTPUT_FILE} with {priority_count} prioritized leads.")
    print(f"Created {OUTREACH_SHEET_OUTPUT_FILE} with {priority_count} prioritized leads.")
    print(f"Total graded websites: {len(report_rows)}")
    print(f"Included in {PRIORITY_OUTPUT_FILE}: {priority_count}")
    print(f"Missing phone: {missing_phone_count}")


if __name__ == "__main__":
    main()
