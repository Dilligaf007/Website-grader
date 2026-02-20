from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from typing import Dict, Optional
import xml.etree.ElementTree as ET

OUTREACH_SHEET_FILE = "outreach_sheet.xlsx"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

TIER_A_SUBJECT = "Quick website question"
TIER_B_SUBJECT_TEMPLATE = "Small idea for {business_name}"
TIER_C_SUBJECT = "Quick question"

TIER_A_BODY_TEMPLATE = """Hi {business_name},

I was reviewing plumbing companies in {city} and came across your website.

I noticed {first_reason} — that alone can reduce inbound calls, especially from mobile users.

I help local service businesses improve lead flow from their existing website without rebuilding everything.

Would you be open to a quick 10-minute call this week to see if there’s a simple fix here?

— Nathan"""

TIER_B_BODY_TEMPLATE = """Hi {business_name},

I came across your website while looking at plumbing companies in {city}.

Your site is solid overall, but there are a couple structural tweaks that could likely increase booked calls.

I work with service businesses on small conversion improvements that often move the needle quickly.

Would you be open to a brief conversation?

— Nathan"""

TIER_C_BODY_TEMPLATE = """Hi {business_name},

I came across your website while reviewing local plumbing companies.

Your website is in good shape. I specialize in optimizing already-strong service sites to increase call volume and booked jobs.

If you're open to it, I’d be happy to share 2–3 small improvements that could increase conversion rates.

Worth a quick look?

— Nathan"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Add email_subject and email_body columns to outreach_sheet.xlsx.")
    parser.add_argument("--path", default=OUTREACH_SHEET_FILE, help="Path to outreach sheet workbook.")
    parser.add_argument("--sheet", default=None, help="Optional sheet name. Defaults to active sheet.")
    return parser.parse_args()


def qn(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def col_to_index(col: str) -> int:
    value = 0
    for ch in col:
        value = value * 26 + (ord(ch) - 64)
    return value


def index_to_col(index: int) -> str:
    parts = []
    while index:
        index, rem = divmod(index - 1, 26)
        parts.append(chr(65 + rem))
    return "".join(reversed(parts))


def split_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        return "A", 1
    return m.group(1), int(m.group(2))


def clean_reason_text(reasons_value: object) -> str:
    raw = str(reasons_value or "").strip()
    if not raw:
        return "a high-impact website issue"
    first_reason = re.split(r"[;,|]", raw, maxsplit=1)[0].strip()
    first_reason = first_reason.replace("_", " ").replace("-", " ")
    first_reason = re.sub(r"\s+", " ", first_reason).strip().lower()
    replacements = {
        "missing https": "your site may not be using HTTPS consistently",
        "missing meta description": "your pages may be missing meta descriptions",
        "missing contact page link": "there may not be a clear contact page link",
        "missing phone": "your phone number may be hard to find",
        "missing email": "your email address may not be clearly visible",
        "missing title": "some pages may be missing clear page titles",
        "bad status code": "some pages may be returning weak status responses",
        "contact page unreachable": "your contact page may be difficult to reach",
        "unreachable": "your website may be intermittently unreachable",
    }
    return replacements.get(first_reason, first_reason)


def to_float(value: object) -> float:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return 0.0


def normalize_text(value: object, fallback: str) -> str:
    text = str(value or "").strip()
    return text or fallback


def infer_city(row: Dict[str, object]) -> str:
    direct_city = str(row.get("city", "")).strip()
    if direct_city:
        return direct_city
    address = str(row.get("address", "")).strip()
    parts = [part.strip() for part in address.split(",") if part.strip()]
    return parts[1] if len(parts) >= 2 else "your area"


def select_tier(score: float) -> str:
    if score >= 75:
        return "A"
    if score >= 50:
        return "B"
    return "C"


def build_message(row: Dict[str, object]) -> Dict[str, str]:
    business_name = normalize_text(row.get("business_name", ""), "there")
    city = infer_city(row)
    first_reason = clean_reason_text(row.get("reasons", ""))
    tier = select_tier(to_float(row.get("opportunity_score_0_100", 0)))
    if tier == "A":
        return {
            "email_subject": TIER_A_SUBJECT,
            "email_body": TIER_A_BODY_TEMPLATE.format(
                business_name=business_name,
                city=city,
                first_reason=first_reason,
            ),
        }
    if tier == "B":
        return {
            "email_subject": TIER_B_SUBJECT_TEMPLATE.format(business_name=business_name),
            "email_body": TIER_B_BODY_TEMPLATE.format(business_name=business_name, city=city),
        }
    return {
        "email_subject": TIER_C_SUBJECT,
        "email_body": TIER_C_BODY_TEMPLATE.format(business_name=business_name),
    }


def read_cell_value(cell: ET.Element, shared: list[str]) -> str:
    ctype = cell.get("t", "")
    if ctype == "inlineStr":
        node = cell.find(f"{qn('is')}/{qn('t')}")
        return node.text if node is not None and node.text else ""
    if ctype == "s":
        v = cell.find(qn("v"))
        if v is None or v.text is None:
            return ""
        idx = int(v.text)
        return shared[idx] if 0 <= idx < len(shared) else ""
    v = cell.find(qn("v"))
    return v.text if v is not None and v.text else ""


def set_inline_cell(cell: ET.Element, ref: str, value: str) -> None:
    cell.clear()
    cell.set("r", ref)
    cell.set("t", "inlineStr")
    is_node = ET.SubElement(cell, qn("is"))
    t_node = ET.SubElement(is_node, qn("t"))
    t_node.text = value


def load_shared_strings(content: bytes) -> list[str]:
    root = ET.fromstring(content)
    values: list[str] = []
    for si in root.findall(qn("si")):
        texts = [t.text or "" for t in si.findall(f".//{qn('t')}")]
        values.append("".join(texts))
    return values


def resolve_sheet_path(files: dict[str, bytes], sheet_name: Optional[str]) -> str:
    workbook = ET.fromstring(files["xl/workbook.xml"])
    wb_rels = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    rel_map = {
        rel.get("Id"): rel.get("Target")
        for rel in wb_rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }

    sheets = workbook.find(qn("sheets"))
    if sheets is None:
        raise ValueError("Workbook has no sheets")

    selected = None
    if sheet_name:
        for sheet in sheets.findall(qn("sheet")):
            if sheet.get("name") == sheet_name:
                selected = sheet
                break
    if selected is None:
        selected = sheets.findall(qn("sheet"))[0]

    rel_id = selected.get(f"{{{REL_NS}}}id")
    target = rel_map.get(rel_id, "")
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("worksheets/"):
        return f"xl/{target}"
    return f"xl/{target}"


def update_sheet(sheet_root: ET.Element, shared_strings: list[str]) -> None:
    sheet_data = sheet_root.find(qn("sheetData"))
    if sheet_data is None:
        raise ValueError("Worksheet has no sheetData")

    rows = sheet_data.findall(qn("row"))
    if not rows:
        raise ValueError("Worksheet has no rows")

    header_row = rows[0]
    header_map: Dict[str, int] = {}

    for cell in header_row.findall(qn("c")):
        ref = cell.get("r", "A1")
        col, _ = split_ref(ref)
        header_value = read_cell_value(cell, shared_strings).strip()
        if header_value:
            header_map[header_value] = col_to_index(col)

    max_col = max(header_map.values(), default=0)
    email_subject_col = header_map.get("email_subject")
    if email_subject_col is None:
        max_col += 1
        email_subject_col = max_col
        c = ET.SubElement(header_row, qn("c"))
        set_inline_cell(c, f"{index_to_col(email_subject_col)}1", "email_subject")

    email_body_col = header_map.get("email_body")
    if email_body_col is None:
        max_col += 1
        email_body_col = max_col
        c = ET.SubElement(header_row, qn("c"))
        set_inline_cell(c, f"{index_to_col(email_body_col)}1", "email_body")

    header_map["email_subject"] = email_subject_col
    header_map["email_body"] = email_body_col

    for row in rows[1:]:
        row_number = int(row.get("r", "0") or 0)
        if row_number <= 0:
            continue
        cells = {split_ref(c.get("r", "A1"))[0]: c for c in row.findall(qn("c"))}

        data: Dict[str, object] = {}
        for header, idx in header_map.items():
            col = index_to_col(idx)
            cell = cells.get(col)
            data[header] = read_cell_value(cell, shared_strings) if cell is not None else ""

        message = build_message(data)

        subject_col = index_to_col(email_subject_col)
        subject_ref = f"{subject_col}{row_number}"
        subject_cell = cells.get(subject_col)
        if subject_cell is None:
            subject_cell = ET.SubElement(row, qn("c"))
        set_inline_cell(subject_cell, subject_ref, message["email_subject"])

        body_col = index_to_col(email_body_col)
        body_ref = f"{body_col}{row_number}"
        body_cell = cells.get(body_col)
        if body_cell is None:
            body_cell = ET.SubElement(row, qn("c"))
        set_inline_cell(body_cell, body_ref, message["email_body"])

        sorted_cells = sorted(row.findall(qn("c")), key=lambda c: col_to_index(split_ref(c.get("r", "A1"))[0]))
        for cell in row.findall(qn("c")):
            row.remove(cell)
        for cell in sorted_cells:
            row.append(cell)

    dim = sheet_root.find(qn("dimension"))
    max_row = int(rows[-1].get("r", "1")) if rows else 1
    max_col_idx = max(header_map.values())
    if dim is None:
        dim = ET.Element(qn("dimension"))
        sheet_root.insert(0, dim)
    dim.set("ref", f"A1:{index_to_col(max_col_idx)}{max_row}")

    sheet_views = sheet_root.find(qn("sheetViews"))
    if sheet_views is None:
        sheet_views = ET.Element(qn("sheetViews"))
        sheet_root.insert(1, sheet_views)
    sheet_view = sheet_views.find(qn("sheetView"))
    if sheet_view is None:
        sheet_view = ET.SubElement(sheet_views, qn("sheetView"), {"workbookViewId": "0"})
    pane = sheet_view.find(qn("pane"))
    if pane is None:
        pane = ET.SubElement(sheet_view, qn("pane"))
    pane.attrib.update({"ySplit": "1", "topLeftCell": "A2", "activePane": "bottomLeft", "state": "frozen"})


def main() -> None:
    args = parse_args()
    xlsx_path = Path(args.path)

    with zipfile.ZipFile(xlsx_path, "r") as source:
        files = {name: source.read(name) for name in source.namelist()}

    shared_strings = []
    if "xl/sharedStrings.xml" in files:
        shared_strings = load_shared_strings(files["xl/sharedStrings.xml"])

    sheet_path = resolve_sheet_path(files, args.sheet)
    sheet_root = ET.fromstring(files[sheet_path])
    update_sheet(sheet_root, shared_strings)
    files[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    tmp_path = xlsx_path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for name, content in files.items():
            target.writestr(name, content)

    tmp_path.replace(xlsx_path)
    print("Outreach sheet updated with email_subject and email_body columns.")


if __name__ == "__main__":
    main()
