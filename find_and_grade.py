import argparse
import csv
import re
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Optional, Set, Tuple

import requests

from main import MAX_WORKERS, PRIORITY_OUTPUT_FILE, grade_website, write_priority_leads

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OVERPASS_URL = "https://overpass-api.de/api/interpreter"
USER_AGENT = "WebsiteGrader-FindAndGrade/1.0 (contact: nathan.strickland.consult@gmail.com)"
NOMINATIM_EMAIL = "nathan.strickland.consult@gmail.com"
REQUEST_DELAY_SECONDS = 1.0
RAW_OUTPUT_FILE = "leads_raw.csv"
GRADED_OUTPUT_FILE = "report.csv"

WEBSITE_KEYS = ["website", "contact:website", "url"]
PHONE_KEYS = ["phone", "contact:phone"]


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


def build_overpass_query(lat: float, lon: float, radius_km: float, query: str, overpass_timeout: int) -> str:
    escaped_query = query.replace('"', '\\"')
    radius_m = int(radius_km * 1000)
    return f"""
[out:json][timeout:{overpass_timeout}];
(
  nwr(around:{radius_m},{lat},{lon})["craft"="plumber"];
  nwr(around:{radius_m},{lat},{lon})["shop"="plumbing"];
  nwr(around:{radius_m},{lat},{lon})["name"~"{escaped_query}",i];
);
out center tags;
""".strip()


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


def find_businesses(
    city: str,
    query: str,
    limit: int,
    radius_km: float,
    overpass_timeout: int,
) -> List[Dict[str, object]]:
    lat, lon = get_city_center(city)
    time.sleep(REQUEST_DELAY_SECONDS)

    overpass_query = build_overpass_query(lat, lon, radius_km, query, overpass_timeout)
    response = requests.post(
        OVERPASS_URL,
        data=overpass_query,
        headers={"User-Agent": USER_AGENT},
        timeout=overpass_timeout + 30,
    )
    response.raise_for_status()

    elements = response.json().get("elements", [])
    leads = [parse_element(element) for element in elements]
    deduped = dedupe_leads(leads)
    return deduped[:limit]


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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Find and grade local businesses from OpenStreetMap.")
    parser.add_argument("--city", required=True, help="City name to search (e.g. 'Austin, TX').")
    parser.add_argument("--query", default="plumber", help="Business search term (default: plumber).")
    parser.add_argument("--limit", type=int, default=50, help="Maximum number of unique leads (default: 50).")
    parser.add_argument(
        "--radius_km",
        type=float,
        default=15,
        help="Search radius around city center in kilometers (default: 15).",
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

    leads = find_businesses(
        city=args.city,
        query=args.query,
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
    priority_count = write_priority_leads(report_rows, PRIORITY_OUTPUT_FILE)

    missing_phone_count = sum(1 for row in report_rows if not str(row.get("phone", "")).strip())

    print(f"Found {len(leads)} unique lead(s). Raw leads saved to {RAW_OUTPUT_FILE}.")
    print(
        f"Graded {len(report_rows)} lead(s) with websites. Report saved to {GRADED_OUTPUT_FILE}."
    )
    print(f"Created {PRIORITY_OUTPUT_FILE} with {priority_count} prioritized leads.")
    print(f"Total graded websites: {len(report_rows)}")
    print(f"Included in {PRIORITY_OUTPUT_FILE}: {priority_count}")
    print(f"Missing phone: {missing_phone_count}")


if __name__ == "__main__":
    main()
