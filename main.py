import csv
import re
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Optional
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

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
    "address",
    "url",
    "opportunity_score_0_100",
    "score_0_100",
    "reasons",
    "pitch",
]


def calculate_opportunity_score(row: Dict[str, object]) -> int:
    """Estimate outreach opportunity from website gaps (0-100)."""
    if not row.get("reachable"):
        return 0

    opportunity = 0
    if not row.get("https"):
        opportunity += 15

    status_code = row.get("status_code")
    if not (isinstance(status_code, int) and 200 <= status_code < 400):
        opportunity += 10

    if not row.get("title"):
        opportunity += 20
    if not row.get("meta_description_present"):
        opportunity += 15
    if not row.get("has_phone"):
        opportunity += 15
    if not row.get("has_email"):
        opportunity += 10
    if not row.get("has_contact_page_link"):
        opportunity += 15

    return min(100, opportunity)


def build_pitch(row: Dict[str, object]) -> str:
    """Create a concise outreach pitch based on detected website gaps."""
    if not row.get("reachable"):
        return ""

    improvements: List[str] = []
    if not row.get("https"):
        improvements.append("secure HTTPS setup")
    if not row.get("title"):
        improvements.append("clear page titles")
    if not row.get("meta_description_present"):
        improvements.append("stronger search snippets")
    if not row.get("has_phone"):
        improvements.append("prominent click-to-call")
    if not row.get("has_email"):
        improvements.append("better lead capture")
    if not row.get("has_contact_page_link"):
        improvements.append("an easy-to-find contact page")

    if not improvements:
        return "Website looks strong overall; we can help improve conversion rate and local SEO performance."

    top_improvements = ", ".join(improvements[:3])
    return f"We can quickly improve {top_improvements} to help generate more qualified leads."


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
    has_email = bool(EMAIL_PATTERN.search(text_content))

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

    return {
        "title": title,
        "meta_description_present": meta_description_present,
        "has_phone": has_phone,
        "has_email": has_email,
        "has_contact_page_link": has_contact_page_link,
        "contact_page_url": contact_page_url or "",
    }


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


def calculate_opportunity_score(score_0_100: int, reasons: str, reachable: bool) -> int:
    """Estimate sales opportunity from technical score and gaps.

    Lower quality websites indicate higher opportunity, but unreachable sites are excluded.
    """
    if not reachable:
        return 0

    reason_set = {reason.strip() for reason in reasons.split(";") if reason.strip()}
    opportunity_score = max(0, 100 - score_0_100)

    # Emphasize lead-gen gaps that are easier to fix and good outreach hooks.
    if "missing_contact_page_link" in reason_set:
        opportunity_score += 10
    if "missing_phone" in reason_set:
        opportunity_score += 8
    if "missing_email" in reason_set:
        opportunity_score += 8
    if "missing_meta_description" in reason_set:
        opportunity_score += 5
    if "missing_title" in reason_set:
        opportunity_score += 4
    if "missing_https" in reason_set:
        opportunity_score += 5
    if "bad_status_code" in reason_set:
        opportunity_score += 7
    if "contact_page_unreachable" in reason_set:
        opportunity_score += 5

    return max(0, min(100, int(opportunity_score)))


def build_pitch(row: Dict[str, object]) -> str:
    if not row["reachable"]:
        return ""

    gaps: List[str] = []
    if not row["has_contact_page_link"]:
        gaps.append("a clear contact page")
    if not row["has_phone"]:
        gaps.append("a visible phone number")
    if not row["has_email"]:
        gaps.append("a visible email address")
    if not row["meta_description_present"]:
        gaps.append("an SEO-ready meta description")
    if not row["https"]:
        gaps.append("HTTPS security")

    if not gaps:
        return "Your site is in strong shape overall—small conversion-focused tweaks could still increase lead volume."

    if len(gaps) == 1:
        gap_text = gaps[0]
    else:
        gap_text = ", ".join(gaps[:-1]) + f", and {gaps[-1]}"

    return (
        f"I noticed your website is missing {gap_text}, which can reduce trust and inbound leads. "
        "We can fix these quickly to help turn more visitors into booked jobs."
    )


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
        "has_contact_page_link": False,
        "contact_page_url": "",
        "contact_page_reachable": False,
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

            contact_page_url = result["contact_page_url"]
            if contact_page_url:
                try:
                    contact_response = fetch_url(contact_page_url)
                    result["contact_page_reachable"] = True

                    contact_content_type = contact_response.headers.get("Content-Type", "").lower()
                    if "text/html" in contact_content_type:
                        contact_data = parse_html(contact_response.text, contact_response.url)
                        result["has_phone"] = bool(result["has_phone"] or contact_data["has_phone"])
                        result["has_email"] = bool(result["has_email"] or contact_data["has_email"])
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

    except requests.exceptions.Timeout:
        result["notes"] = f"Request timed out after {TIMEOUT_SECONDS}s"
    except requests.exceptions.RequestException as exc:
        result["notes"] = f"Request failed: {exc}"
    except Exception as exc:
        result["notes"] = f"Unexpected error: {exc}"

    result["score_0_100"] = calculate_score(result)
    result["opportunity_score_0_100"] = calculate_opportunity_score(result)
    result["pitch"] = build_pitch(result)
    result["reasons"] = build_reasons(result)
    result["opportunity_score_0_100"] = calculate_opportunity_score(
        int(result["score_0_100"]),
        str(result["reasons"]),
        bool(result["reachable"]),
    )
    result["pitch"] = build_pitch(result)
    return result


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


if __name__ == "__main__":
    main()
