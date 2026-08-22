#!/usr/bin/env python3
"""
Sync the latest scraper results to Google Sheets.

Supports two modes:
  1. FULL sync (default): clears the "Matches" sheet and rewrites all matches.
  2. INCREMENTAL sync (--incremental): appends only new matches not already
     in the sheet (deduped by Job URL).

Setup (one-time):
  1. pip install gspread google-auth
  2. Create a Google Service Account:
     - Go to https://console.cloud.google.com/iam-admin/serviceaccounts
     - Create a new service account
     - Create a JSON key and download it
     - Save it as 'google-credentials.json' in the same folder as this script
  3. Open your Google Sheet, click "Share", and add the service account's
     email (found in the JSON file as 'client_email') as an Editor.

Usage:
  # Full sync (clears + rewrites):
  python sync_to_google_sheets.py

  # Incremental sync (appends new matches only):
  python sync_to_google_sheets.py --incremental

  # Or set the sheet ID as an env var:
  export GOOGLE_SHEET_ID=1VAinuzwZfAP8IVEwcr8sLUZsNMcy1OiPmQhmhfNSvYM
  python sync_to_google_sheets.py
"""
from __future__ import annotations

import argparse
import glob
import os
import sys
from pathlib import Path

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
except ImportError:
    print("ERROR: gspread not installed. Run:  pip install gspread google-auth", file=sys.stderr)
    sys.exit(1)

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl not installed. Run:  pip install openpyxl", file=sys.stderr)
    sys.exit(1)

# The target Google Sheet ID (from the URL).
DEFAULT_SHEET_ID = "1VAinuzwZfAP8IVEwcr8sLUZsNMcy1OiPmQhmhfNSvYM"

# Scopes needed for Sheets API.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def find_credentials() -> str:
    """Find the Google service account credentials JSON file."""
    candidates = [
        "google-credentials.json",
        os.path.expanduser("~/google-credentials.json"),
        "/home/mikeysama/saasnews-job-finder/google-credentials.json",
        os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", ""),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return ""


def connect_to_sheet(sheet_id: str, creds_path: str):
    """Connect to Google Sheets and return the spreadsheet handle."""
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    gc = gspread.Client(auth=authed_session)
    gc.session = authed_session
    spreadsheet = gc.open_by_key(sheet_id)
    return spreadsheet


def read_matches_from_xlsx(xlsx_path: str) -> list[list]:
    """Read the Matches sheet from the xlsx file. Returns list of rows
    (first row = headers)."""
    wb = load_workbook(xlsx_path, read_only=True)
    if "Matches" not in wb.sheetnames:
        print(f"ERROR: 'Matches' sheet not found in {xlsx_path}", file=sys.stderr)
        print(f"Available sheets: {wb.sheetnames}", file=sys.stderr)
        sys.exit(1)
    ws = wb["Matches"]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([str(v) if v is not None else "" for v in row])
    wb.close()
    return rows


def full_sync(spreadsheet, rows: list[list]) -> None:
    """Full sync: clear the sheet and rewrite all rows."""
    try:
        worksheet = spreadsheet.worksheet("Matches")
        worksheet.clear()
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet("Matches",
                                               rows=max(len(rows), 100),
                                               cols=len(rows[0]) if rows else 20)

    # Also clear default Sheet1.
    try:
        sheet1 = spreadsheet.worksheet("Sheet1")
        sheet1.clear()
    except gspread.WorksheetNotFound:
        pass

    if not rows:
        print("No rows to write.")
        return

    print(f"Writing {len(rows)} rows (full sync)...")
    end_cell = gspread.utils.rowcol_to_a1(len(rows), len(rows[0]))
    worksheet.update(
        f"A1:{end_cell}",
        rows,
        value_input_option="RAW",
    )
    _format_header(worksheet)
    print(f"✓ Full sync complete: {len(rows) - 1} matches.")


def incremental_sync(spreadsheet, rows: list[list]) -> None:
    """Incremental sync: append only new matches (deduped by Job URL)."""
    if len(rows) <= 1:
        print("No matches to sync.")
        return

    try:
        worksheet = spreadsheet.worksheet("Matches")
    except gspread.WorksheetNotFound:
        # Create the worksheet with headers.
        worksheet = spreadsheet.add_worksheet("Matches", rows=100, cols=len(rows[0]))
        worksheet.update("A1", [rows[0]], value_input_option="RAW")
        _format_header(worksheet)
        worksheet.freeze(rows=1)
        print("Created 'Matches' worksheet with headers.")

    # Read existing Job URLs (column K = index 10 in our schema, but let's find it).
    headers = worksheet.row_values(1)
    job_url_col = None
    for i, h in enumerate(headers, 1):
        if h.strip().lower() == "job url":
            job_url_col = i
            break
    if job_url_col is None:
        print("WARNING: 'Job URL' column not found. Falling back to full sync.")
        full_sync(spreadsheet, rows)
        return

    # Read all existing job URLs.
    col_letter = gspread.utils.get_column_letter(job_url_col)
    existing_urls = set()
    try:
        existing_values = worksheet.col_values(job_url_col)[1:]  # skip header
        existing_urls = {v for v in existing_values if v}
    except Exception as e:
        print(f"WARNING: Could not read existing URLs: {e}")

    # Find new rows.
    header_row = rows[0]
    job_url_idx = None
    for i, h in enumerate(header_row):
        if h.strip().lower() == "job url":
            job_url_idx = i
            break
    if job_url_idx is None:
        print("WARNING: 'Job URL' not in xlsx headers. Falling back to full sync.")
        full_sync(spreadsheet, rows)
        return

    new_rows = []
    for row in rows[1:]:
        job_url = row[job_url_idx] if job_url_idx < len(row) else ""
        if job_url and job_url not in existing_urls:
            new_rows.append(row)
            existing_urls.add(job_url)

    if not new_rows:
        print(f"No new matches to append ({len(existing_urls)} already in sheet).")
        return

    # Append new rows.
    print(f"Appending {len(new_rows)} new matches (incremental sync)...")
    # gspread append_rows adds to the first empty row after the data.
    worksheet.append_rows(new_rows, value_input_option="RAW",
                          insert_data_option="INSERT_ROWS", table_range="A1")
    print(f"✓ Incremental sync complete: appended {len(new_rows)} new matches.")


def _format_header(worksheet) -> None:
    """Format the header row (bold white text on dark background)."""
    worksheet.format("1:1", {
        "backgroundColor": {"red": 0.12, "green": 0.16, "blue": 0.23},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "horizontalAlignment": "CENTER",
        "wrapStrategy": "WRAP",
    })
    worksheet.freeze(rows=1)


def main():
    parser = argparse.ArgumentParser(description="Sync scraper results to Google Sheets")
    parser.add_argument("--sheet-id", default=os.environ.get("GOOGLE_SHEET_ID", DEFAULT_SHEET_ID),
                        help=f"Google Sheet ID (default: {DEFAULT_SHEET_ID})")
    parser.add_argument("--xlsx", default="",
                        help="Path to xlsx file (default: latest in download/)")
    parser.add_argument("--download-dir", default="/home/mikeysama/products/saasnews-job-finder/download",
                        help="Directory containing xlsx files")
    parser.add_argument("--creds", default="/home/mikeysama/saasnews-job-finder/google-credentials.json",
                        help="Path to Google service account JSON (default: auto-detect)")
    parser.add_argument("--incremental", action="store_true",
                        help="Incremental sync: append only new matches (deduped by Job URL). "
                             "Default is full sync (clear + rewrite).")
    args = parser.parse_args()

    # Find the xlsx file.
    if args.xlsx:
        xlsx_path = args.xlsx
    else:
        latest = os.path.join(args.download_dir, "saasnews_jobs_latest.xlsx")
        if os.path.exists(latest):
            xlsx_path = latest
        else:
            files = sorted(glob.glob(os.path.join(args.download_dir, "saasnews_jobs_*.xlsx")))
            if not files:
                print(f"ERROR: No xlsx files found in {args.download_dir}", file=sys.stderr)
                sys.exit(1)
            xlsx_path = files[-1]
    print(f"Reading: {xlsx_path}")

    # Find credentials.
    creds_path = args.creds or find_credentials()
    if not creds_path:
        print("ERROR: No Google service account credentials found.", file=sys.stderr)
        print("Create one at https://console.cloud.google.com/iam-admin/serviceaccounts", file=sys.stderr)
        print("Save it as 'google-credentials.json' in this directory.", file=sys.stderr)
        sys.exit(1)
    print(f"Using credentials: {creds_path}")

    # Read matches.
    rows = read_matches_from_xlsx(xlsx_path)
    print(f"Read {len(rows) - 1} matches from xlsx.")

    if len(rows) <= 1:
        print("No matches to sync. Exiting.")
        return

    # Connect and sync.
    print(f"Connecting to Google Sheet: {args.sheet_id}")
    try:
        spreadsheet = connect_to_sheet(args.sheet_id, creds_path)
        print(f"Connected. Sheet title: {spreadsheet.title}")
    except Exception as e:
        print(f"ERROR: Could not connect to Google Sheet: {e}", file=sys.stderr)
        print("Make sure you shared the sheet with the service account email.", file=sys.stderr)
        sys.exit(1)

    if args.incremental:
        incremental_sync(spreadsheet, rows)
    else:
        full_sync(spreadsheet, rows)

    print(f"\n✓ Done. View at: https://docs.google.com/spreadsheets/d/{args.sheet_id}/edit")


if __name__ == "__main__":
    main()
