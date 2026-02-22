import csv
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Optional
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from google_sheets_export import export_to_google_sheets

INPUT_FILE = "websites.txt"
OUTPUT_FILE = "report.csv"
PRIORITY_OUTPUT_FILE = "priority_leads.csv"
TIMEOUT_SECONDS = 10
MAX_WORKERS = 8
USER_AGENT = "WebsiteGrader/1.0"

# Scoring rubric (0-100):
# - Reachable website: +20
# - Uses HTTPS in final URL: +15
# - HTTP status in 200-399: +10
# - Page title present: +10
# - Meta description present: +10
# - Phone number detected (homepage + contact page): +10
# - Email address detected (homepage + contact page): +10
# - Contact page link detected: +15
# Total = 100

PHONE_PATTERN = re.compile(
    r"(?<!\d)(?:\+?1[\s.-]?)?(?:\(\d{3}\)|\d{3})[\s.-]?\d{3}[\s.-]?\d{4}(?!\d)"
)
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
DATE_LIKE_PATTERN = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b")
ZIP_PATTERN = re.compile(r"\b\d{5}(?:-\d{4})?\b")

PRIORITY_FIELDNAMES = [
    "business_name",
    "phone",
    "primary_email",
    "emails",
    "has_contact_form",
    "contact_form_url",
    "outreach_channel",
    "form_message",
    "call_script",
    "address",
    "url",
    "opportunity_score_0_100",
    "score_0_100",
    "reasons",
    "pitch",
]

CONTACT_LINK_HINTS = ["contact", "get-in-touch", "request-service", "schedule", "estimate"]


def normalize_url(url: str) -> str:
    """Ensure URL has a scheme; default to https."""
    trimmed = url.strip()
    if not trimmed:
        return ""
    if not urlparse(trimmed).scheme:
        return f"https://{trimmed}"
    return trimmed


def fetch_url(url: str) -> requests.Response:
    return requests.get(
        url,
        timeout=TIMEOUT_SECONDS,
        headers={"User-Agent": USER_AGENT},
        allow_redirects=True,
    )


def has_valid_us_phone(text_content: str) -> bool:
    for match in PHONE_PATTERN.finditer(text_content):
        candidate = match.group(0)
        digits = re.sub(r"\D", "", candidate)

        if len(digits) == 11 and digits.startswith("1"):
            digits = digits[1:]

        if len(digits) != 10:
            continue

        if DATE_LIKE_PATTERN.search(candidate) or ZIP_PATTERN.fullmatch(candidate.strip()):
            continue

        return True

    return False


def extract_emails(text: str) -> List[str]:
    seen: set[str] = set()
    emails: List[str] = []
    for match in EMAIL_PATTERN.finditer(text):
        email = match.group(0)
        normalized = email.lower()
        if normalized in seen:
            continue
        seen.add(normalized)
        emails.append(email)
    return emails


def parse_html(html: str, base_url: str) -> Dict[str, object]:
    soup = BeautifulSoup(html, "html.parser")

    title_tag = soup.find("title")
    title = title_tag.get_text(strip=True) if title_tag else ""

    meta_description_tag = soup.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
    meta_description_present = bool(
        meta_description_tag and meta_description_tag.get("content", "").strip()
    )

    text_content = soup.get_text(" ", strip=True)
    has_phone = has_valid_us_phone(text_content)
    emails = extract_emails(text_content)
    has_email = bool(emails)

    contact_page_url: Optional[str] = None
    has_contact_page_link = False

    for link in soup.find_all("a", href=True):
        href = link.get("href", "")
        anchor_text = link.get_text(" ", strip=True).lower()
        resolved_href = urljoin(base_url, href)

        if "contact" in anchor_text or "contact" in href.lower() or "contact" in resolved_href.lower():
            has_contact_page_link = True
            contact_page_url = resolved_href
            break

    if not has_contact_page_link:
        for link in soup.find_all("a", href=True):
            href = link.get("href", "")
            anchor_text = link.get_text(" ", strip=True).lower()
            resolved_href = urljoin(base_url, href)
            haystack = f"{anchor_text} {href.lower()} {resolved_href.lower()}"
            if any(hint in haystack for hint in CONTACT_LINK_HINTS):
                has_contact_page_link = True
                contact_page_url = resolved_href
                break

    return {
        "title": title,
        "meta_description_present": meta_description_present,
        "has_phone": has_phone,
        "has_email": has_email,
        "emails": emails,
        "has_contact_page_link": has_contact_page_link,
        "contact_page_url": contact_page_url or "",
    }


def is_usable_contact_form(form: BeautifulSoup) -> bool:
    descriptor_parts = [
        str(form.get("action", "")),
        str(form.get("id", "")),
        " ".join(form.get("class", [])),
    ]
    descriptor = " ".join(descriptor_parts).lower()
    if "search" in descriptor:
        return False

    textarea_exists = form.find("textarea") is not None
    email_input = form.find("input", attrs={"type": re.compile(r"^email$", re.I)})
    email_name_input = form.find("input", attrs={"name": re.compile(r"email", re.I)})
    submit_button = form.find("button", attrs={"type": re.compile(r"^submit$", re.I)})
    submit_input = form.find("input", attrs={"type": re.compile(r"^submit$", re.I)})
    has_submit = bool(submit_button or submit_input)

    relevant_inputs = form.find_all("input")
    text_like_inputs = [
        input_tag
        for input_tag in relevant_inputs
        if str(input_tag.get("type", "text")).lower() in {"", "text"}
    ]
    has_other_input_types = any(
        str(input_tag.get("type", "text")).lower() not in {"", "text"}
        for input_tag in relevant_inputs
    )

    if len(text_like_inputs) == 1 and not textarea_exists and not email_input and not email_name_input and not has_other_input_types:
        return False

    return bool(textarea_exists or email_input or email_name_input or has_submit)


def detect_contact_form(html: str) -> bool:
    soup = BeautifulSoup(html, "html.parser")
    forms = soup.find_all("form")
    if not forms:
        return False
    return any(is_usable_contact_form(form) for form in forms)


def select_outreach_channel(has_email: bool, has_contact_form: bool, has_phone: bool) -> str:
    if has_email:
        return "email"
    if has_contact_form:
        return "form"
    if has_phone:
        return "phone"
    return "none"


def merge_emails(*email_lists: List[str]) -> List[str]:
    seen: set[str] = set()
    merged: List[str] = []
    for email_list in email_lists:
        for email in email_list:
            normalized = email.lower()
            if normalized in seen:
                continue
            seen.add(normalized)
            merged.append(email)
    return merged


def calculate_score(row: Dict[str, object]) -> int:
    score = 0
    if row["reachable"]:
        score += 20
    if row["https"]:
        score += 15

    status_code = row["status_code"]
    if isinstance(status_code, int) and 200 <= status_code < 400:
        score += 10

    if row["title"]:
        score += 10
    if row["meta_description_present"]:
        score += 10
    if row["has_phone"]:
        score += 10
    if row["has_email"]:
        score += 10
    if row["has_contact_page_link"]:
        score += 15

    return score


def build_reasons(row: Dict[str, object]) -> str:
    reasons = []
    if not row["reachable"]:
        reasons.append("unreachable")
    if not row["https"]:
        reasons.append("missing_https")

    status_code = row["status_code"]
    if not (isinstance(status_code, int) and 200 <= status_code < 400):
        reasons.append("bad_status_code")

    if not row["title"]:
        reasons.append("missing_title")
    if not row["meta_description_present"]:
        reasons.append("missing_meta_description")
    if not row["has_phone"]:
        reasons.append("missing_phone")
    if not row["has_email"]:
        reasons.append("missing_email")
    if not row["has_contact_page_link"]:
        reasons.append("missing_contact_page_link")
    elif not row["contact_page_reachable"]:
        reasons.append("contact_page_unreachable")

    return "; ".join(reasons)


def calculate_opportunity_score(row: Dict[str, object]) -> int:
    opportunity = 0

    raw_score = row.get("score_0_100", 0)
    try:
        score = float(raw_score)
    except (TypeError, ValueError):
        score = 0
    has_email = bool(row.get("has_email"))
    has_contact = bool(row.get("has_contact_page_link"))
    has_meta = bool(row.get("meta_description_present"))
    has_phone = bool(row.get("has_phone"))

    # Reward missing fundamentals
    if not has_email:
        opportunity += 25
    if not has_contact:
        opportunity += 25
    if not has_meta:
        opportunity += 15
    if not has_phone:
        opportunity += 15

    # Reward weak overall site
    if score < 60:
        opportunity += 25
    elif score < 75:
        opportunity += 15
    elif score < 85:
        opportunity += 5

    # Penalize very strong sites
    if score > 90:
        opportunity -= 20

    return max(0, min(100, opportunity))


def build_pitch(row: Dict[str, object]) -> str:
    gaps: List[str] = []
    if not row["has_contact_page_link"]:
        gaps.append("a clear contact page")
    if not row["has_email"]:
        gaps.append("a visible email address")
    if not row["meta_description_present"]:
        gaps.append("an SEO-ready meta description")

    if not row["reachable"]:
        return (
            "I couldn't access your website during a quick review. "
            "I can help diagnose uptime/performance issues and improve lead capture once it's reachable."
        )

    if not gaps:
        return "Your site is in strong shape overall—small conversion-focused tweaks could still increase lead volume."

    if len(gaps) == 1:
        gap_text = gaps[0]
    else:
        gap_text = ", ".join(gaps[:-1]) + f", and {gaps[-1]}"

    return f"I noticed your website is missing {gap_text}; we can fix that quickly to improve inbound leads."


def grade_website(url: str) -> Dict[str, object]:
    normalized_url = normalize_url(url)
    result: Dict[str, object] = {
        "url": normalized_url,
        "reachable": False,
        "https": False,
        "status_code": "",
        "final_url": "",
        "title": "",
        "meta_description_present": False,
        "has_phone": False,
        "has_email": False,
        "emails": "",
        "primary_email": "",
        "has_contact_page_link": False,
        "contact_page_url": "",
        "contact_page_reachable": False,
        "has_contact_form": False,
        "contact_form_url": "",
        "outreach_channel": "none",
        "form_message": "",
        "call_script": "",
        "notes": "",
        "reasons": "",
        "score_0_100": 0,
        "opportunity_score_0_100": 0,
        "pitch": "",
    }

    if not normalized_url:
        result["notes"] = "Empty URL entry"
        result["reasons"] = build_reasons(result)
        return result

    try:
        response = fetch_url(normalized_url)
        result["reachable"] = True
        result["status_code"] = response.status_code
        result["final_url"] = response.url
        result["https"] = urlparse(response.url).scheme.lower() == "https"

        content_type = response.headers.get("Content-Type", "").lower()
        if "text/html" not in content_type:
            result["notes"] = f"Non-HTML content type: {content_type or 'unknown'}"
        else:
            homepage_data = parse_html(response.text, response.url)
            result.update(homepage_data)
            combined_emails = merge_emails(homepage_data.get("emails", []))

            contact_page_url = result["contact_page_url"]
            if contact_page_url:
                try:
                    contact_response = fetch_url(contact_page_url)
                    result["contact_page_reachable"] = True

                    contact_content_type = contact_response.headers.get("Content-Type", "").lower()
                    if "text/html" in contact_content_type:
                        result["has_contact_form"] = detect_contact_form(contact_response.text)
                        if result["has_contact_form"]:
                            result["contact_form_url"] = contact_response.url

                        contact_data = parse_html(contact_response.text, contact_response.url)
                        result["has_phone"] = bool(result["has_phone"] or contact_data["has_phone"])
                        combined_emails = merge_emails(
                            combined_emails,
                            contact_data.get("emails", []),
                        )
                        result["has_email"] = bool(combined_emails)
                        result["has_contact_page_link"] = bool(
                            result["has_contact_page_link"] or contact_data["has_contact_page_link"]
                        )
                    else:
                        note = f"Contact page non-HTML content type: {contact_content_type or 'unknown'}"
                        result["notes"] = f"{result['notes']}; {note}".strip("; ")
                except requests.exceptions.Timeout:
                    note = f"Contact page request timed out after {TIMEOUT_SECONDS}s"
                    result["notes"] = f"{result['notes']}; {note}".strip("; ")
                except requests.exceptions.RequestException as exc:
                    note = f"Contact page request failed: {exc}"
                    result["notes"] = f"{result['notes']}; {note}".strip("; ")

            result["emails"] = "; ".join(combined_emails)
            result["primary_email"] = combined_emails[0] if combined_emails else ""
            result["has_email"] = bool(combined_emails)
            result["outreach_channel"] = select_outreach_channel(
                bool(result["has_email"]),
                bool(result["has_contact_form"]),
                bool(result["has_phone"]),
            )

    except requests.exceptions.Timeout:
        result["notes"] = f"Request timed out after {TIMEOUT_SECONDS}s"
    except requests.exceptions.RequestException as exc:
        result["notes"] = f"Request failed: {exc}"
    except Exception as exc:
        result["notes"] = f"Unexpected error: {exc}"

    result["reasons"] = build_reasons(result)
    result["score_0_100"] = calculate_score(result)
    result["opportunity_score_0_100"] = calculate_opportunity_score(result)
    result["pitch"] = build_pitch(result)
    result["outreach_channel"] = select_outreach_channel(
        bool(result.get("has_email")),
        bool(result.get("has_contact_form")),
        bool(result.get("has_phone")),
    )
    return result


def run_self_test() -> None:
    sample = grade_website("https://example.com")
    assert "opportunity_score_0_100" in sample
    assert isinstance(sample["opportunity_score_0_100"], int)
    assert 0 <= sample["opportunity_score_0_100"] <= 100
    assert "pitch" in sample


def load_websites(path: str) -> List[str]:
    websites: List[str] = []
    with open(path, "r", encoding="utf-8") as file:
        for line in file:
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                websites.append(stripped)
    return websites


def write_report(rows: List[Dict[str, object]], output_path: str) -> None:
    fieldnames = [
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
        "has_contact_form",
        "contact_form_url",
        "outreach_channel",
        "form_message",
        "call_script",
        "has_contact_page_link",
        "contact_page_url",
        "contact_page_reachable",
        "notes",
        "reasons",
        "score_0_100",
        "opportunity_score_0_100",
        "pitch",
    ]

    with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def as_bool(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() == "true"
    return bool(value)


def as_float(value: object) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def write_priority_leads(rows: List[Dict[str, object]], output_path: str) -> int:
    actionable_rows: List[Dict[str, object]] = []
    for row in rows:
        if not as_bool(row.get("reachable", False)):
            continue
        if not str(row.get("url", "")).strip():
            continue

        opportunity_score = as_float(row.get("opportunity_score_0_100"))
        if opportunity_score is None or opportunity_score <= 0:
            continue

        actionable_rows.append(row)

    sorted_rows = sorted(
        actionable_rows,
        key=lambda row: as_float(row.get("opportunity_score_0_100")) or 0.0,
        reverse=True,
    )

    priority_rows = [
        {field: row.get(field, "") for field in PRIORITY_FIELDNAMES}
        for row in sorted_rows
    ]

    with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=PRIORITY_FIELDNAMES)
        writer.writeheader()
        writer.writerows(priority_rows)

    return len(priority_rows)


def main() -> None:
    websites = load_websites(INPUT_FILE)

    # executor.map preserves the input order, keeping output deterministic.
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        rows = list(executor.map(grade_website, websites))

    write_report(rows, OUTPUT_FILE)
    priority_count = write_priority_leads(rows, PRIORITY_OUTPUT_FILE)
    missing_phone_count = sum(1 for row in rows if not as_bool(row.get("has_phone", False)))

    print(f"Graded {len(rows)} website(s). Report saved to {OUTPUT_FILE}.")
    print(f"Created {PRIORITY_OUTPUT_FILE} with {priority_count} prioritized leads.")
    print(f"Total graded websites: {len(rows)}")
    print(f"Included in {PRIORITY_OUTPUT_FILE}: {priority_count}")
    print(f"Missing phone: {missing_phone_count}")

    export_to_google_sheets(
        [
            (OUTPUT_FILE, os.getenv("GOOGLE_SHEETS_TAB_REPORT", "report")),
            (PRIORITY_OUTPUT_FILE, os.getenv("GOOGLE_SHEETS_TAB_PRIORITY", "priority_leads")),
            ("leads_raw.csv", os.getenv("GOOGLE_SHEETS_TAB_LEADS_RAW", "leads_raw")),
        ]
    )


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        run_self_test()
        print("Self-test passed.")
        sys.exit(0)
    main()
