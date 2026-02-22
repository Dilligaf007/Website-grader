import csv
import os
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

GOOGLE_SHEETS_SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _read_csv_rows(csv_path: Path) -> List[List[str]]:
    with csv_path.open("r", encoding="utf-8", newline="") as file:
        reader = csv.reader(file)
        return [list(row) for row in reader]


def _upsert_worksheet(spreadsheet: object, tab_name: str, rows: Sequence[Sequence[str]]) -> None:
    from gspread.exceptions import WorksheetNotFound

    try:
        worksheet = spreadsheet.worksheet(tab_name)
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=tab_name,
            rows=max(len(rows), 1000),
            cols=max(len(rows[0]) if rows else 1, 26),
        )

    worksheet.clear()

    if rows:
        worksheet.update(values=rows, range_name="A1")
        worksheet.freeze(rows=1)

        try:
            worksheet.columns_auto_resize(0, len(rows[0]))
        except Exception:
            pass


def export_to_google_sheets(file_tab_pairs: Iterable[Tuple[str, str]]) -> None:
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_ID", "").strip()
    service_account_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

    if not spreadsheet_id or not service_account_path:
        print("Google Sheets export skipped (missing GOOGLE_SHEETS_ID / GOOGLE_SERVICE_ACCOUNT_JSON)")
        return

    try:
        import gspread
        from google.oauth2.service_account import Credentials

        credentials = Credentials.from_service_account_file(service_account_path, scopes=GOOGLE_SHEETS_SCOPE)
        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_key(spreadsheet_id)
    except Exception as exc:
        print(f"Google Sheets export skipped (setup failed: {exc})")
        return

    for csv_path, tab_name in file_tab_pairs:
        csv_file = Path(csv_path)
        if not csv_file.exists():
            continue

        try:
            rows = _read_csv_rows(csv_file)
            _upsert_worksheet(spreadsheet, tab_name, rows)
            print(f"Exported {csv_file.name} -> tab '{tab_name}'")
        except Exception as exc:
            print(f"Google Sheets export failed for {csv_file.name}: {exc}")
