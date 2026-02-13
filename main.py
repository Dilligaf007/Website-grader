import csv
import re
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Optional
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

INPUT_FILE = "websites.txt"
OUTPUT_FILE = "report.csv"
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
    result["reasons"] = build_reasons(result)
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
    ]

    with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    websites = load_websites(INPUT_FILE)

    # executor.map preserves the input order, keeping output deterministic.
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        rows = list(executor.map(grade_website, websites))

    write_report(rows, OUTPUT_FILE)
    print(f"Graded {len(rows)} website(s). Report saved to {OUTPUT_FILE}.")


if __name__ == "__main__":
    main()
