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
from datetime import datetime, timezone
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
    """Find the Google service account credentials JSON file.

    Precedence: GOOGLE_APPLICATION_CREDENTIALS env var → ./google-credentials.json
    → ~/google-credentials.json.
    """
    candidates = [
        os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", ""),
        "google-credentials.json",
        os.path.expanduser("~/google-credentials.json"),
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


# Columns the user owns in the sheet; a re-sync must never overwrite them.
USER_COLUMNS = ("applied", "feedback", "notes")


def _parse_seen(value) -> datetime:
    try:
        return datetime.fromisoformat(str(value))
    except (TypeError, ValueError):
        return datetime.min.replace(tzinfo=timezone.utc)


def merge_sheet_rows(headers: list[str], existing: list[list],
                     incoming: list[list], keep_sheet_only: bool = True) -> list[list]:
    """Merge incoming scrape rows into the sheet's current rows.

    - Deduplicated by Job URL.
    - Sorted newest-first by First Seen (latest scraped jobs at the top);
      timestamp-less rows sink to the bottom in stable order.
    - USER_COLUMNS survive from the existing sheet row when the URL matches
      (the xlsx regenerates them blank); system columns refresh from incoming.
    - keep_sheet_only=True (incremental): rows that exist only in the sheet
      are retained below the merged set. False (full sync): dropped.

    Returns rows padded/trimmed to len(headers), headers row excluded.
    """
    hmap = {h.strip().lower(): i for i, h in enumerate(headers) if h}
    url_idx = hmap.get("job url")
    seen_idx = hmap.get("first seen")
    user_idx = [hmap[c] for c in USER_COLUMNS if c in hmap]
    width = len(headers)
    if url_idx is None:
        # Cannot dedupe without a key column — fall back to incoming only.
        return [[str(v) for v in r] for r in incoming]

    def normalize(row):
        vals = [str(v) if v is not None else "" for v in row]
        return vals[:width] + [""] * (width - len(vals))

    merged: dict[str, list] = {}
    order: list[str] = []

    for row in existing:
        row = normalize(row)
        url = row[url_idx]
        if not url:
            continue
        if url not in merged:
            order.append(url)
        merged[url] = row

    for row in incoming:
        row = normalize(row)
        url = row[url_idx]
        if not url:
            continue
        prior = merged.get(url)
        if prior is not None:
            # Refresh system columns; preserve the user's tracking edits.
            for i in range(width):
                if i not in user_idx:
                    row[i] = row[i] or prior[i]
                else:
                    row[i] = prior[i] or row[i]
        else:
            order.append(url)
        merged[url] = row

    if keep_sheet_only:
        keys = [(order.index(u), u) for u in order]
    else:
        incoming_urls = {normalize(r)[url_idx] for r in incoming}
        keys = [(i, u) for i, u in enumerate(order) if u in incoming_urls]

    rows = [merged[u] for _, u in sorted(keys)]
    if seen_idx is None:
        return rows
    rows.sort(key=lambda r: _parse_seen(r[seen_idx]), reverse=True)

    # Applied jobs sink below the active feed: the top of the sheet must be
    # the fresh list. Relative order within each section is preserved.
    applied_idx = hmap.get("applied")
    if applied_idx is not None:
        active = [r for r in rows if r[applied_idx].strip().lower() != "yes"]
        done = [r for r in rows if r[applied_idx].strip().lower() == "yes"]
        if done:
            rows = active + done
    return rows


def _create_matches_sheet(spreadsheet, headers):
    try:
        worksheet = spreadsheet.worksheet("Matches")
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet("Matches", rows=100,
                                              cols=len(headers))
        worksheet.update("A1", [headers], value_input_option="RAW")
        print("Created 'Matches' worksheet with headers.")
    return worksheet


def _existing_rows(worksheet) -> list[list]:
    try:
        all_values = worksheet.get_all_values()
        return all_values[1:] if all_values else []
    except Exception as e:
        print(f"WARNING: could not read existing sheet rows: {e}")
        return []


def _write_merged(worksheet, headers: list[str], merged: list[list],
                  mode: str) -> None:
    if not merged:
        print("No matches to sync.")
        return
    worksheet.resize(rows=len(merged) + 10, cols=len(headers))
    end_cell = gspread.utils.rowcol_to_a1(len(merged) + 1, len(headers))
    worksheet.update(
        f"A1:{end_cell}",
        [headers] + merged,
        value_input_option="RAW",
    )
    _format_header(worksheet)
    _apply_dropdowns(worksheet, headers)
    print(f"✓ {mode} sync complete: {len(merged)} matches "
          f"(newest at top, Applied/Feedback/Notes preserved).")


def _apply_dropdowns(worksheet, headers: list[str]) -> None:
    """Best-effort dropdown validation for the tracking columns."""
    try:
        hmap = {h.strip().lower(): i for i, h in enumerate(headers) if h}
        validations = []
        if "applied" in hmap:
            validations.append((hmap["applied"], ["Yes", "No"]))
        if "feedback" in hmap:
            from saasnews.feedback import CONSTRAINTS
            validations.append((hmap["feedback"], list(CONSTRAINTS)))
        for col_idx, values in validations:
            col = gspread.utils.get_column_letter(col_idx + 1)
            dv = gspread.DataValidation(
                requirement_type="ONE_OF_LIST",
                strict=False,
                values=[v for v in values],
                show_custom_ui=True,
            )
            dv.add(f"{col}2:{col}10000")
            worksheet.add_data_validation(dv)
    except Exception as e:
        print(f"(Dropdown validation skipped: {e})")


def full_sync(spreadsheet, rows: list[list]) -> None:
    """Full sync: the sheet mirrors the xlsx exactly. User-owned columns are
    preserved by Job URL; rows no longer present in the xlsx are dropped."""
    headers, incoming = rows[0], rows[1:]
    worksheet = _create_matches_sheet(spreadsheet, headers)
    existing = _existing_rows(worksheet)
    merged = merge_sheet_rows(headers, existing, incoming, keep_sheet_only=False)

    # Also clear default Sheet1.
    try:
        spreadsheet.worksheet("Sheet1").clear()
    except gspread.WorksheetNotFound:
        pass

    if not merged:
        print("No rows to write.")
        return
    print(f"Merging {len(incoming)} xlsx rows against {len(existing)} sheet rows...")
    _write_merged(worksheet, headers, merged, mode="Full")


def incremental_sync(spreadsheet, rows: list[list]) -> None:
    """Incremental sync: append new matches only (deduped by Job URL),
    newest at top, preserving user tracking edits and sheet history."""
    if len(rows) <= 1:
        print("No matches to sync.")
        return
    headers, incoming = rows[0], rows[1:]
    worksheet = _create_matches_sheet(spreadsheet, headers)
    existing = _existing_rows(worksheet)

    hmap = {h.strip().lower(): i for i, h in enumerate(headers) if h}
    url_idx = hmap.get("job url")
    width = len(headers)

    def norm(row):
        vals = [str(v) if v is not None else "" for v in row]
        return vals[:width] + [""] * (width - len(vals))

    if url_idx is None:
        print("WARNING: 'Job URL' column not found; cannot dedupe.")
        return
    existing_urls = {norm(r)[url_idx] for r in existing if r}
    new_count = sum(1 for r in incoming
                    if norm(r)[url_idx] and norm(r)[url_idx] not in existing_urls)
    print(f"Merging {len(incoming)} xlsx rows against {len(existing)} sheet rows "
          f"({new_count} new)...")

    merged = merge_sheet_rows(headers, existing, incoming, keep_sheet_only=True)
    _write_merged(worksheet, headers, merged, mode="Incremental")


def _format_header(worksheet) -> None:
    """Format the header row (bold white text on dark background)."""
    worksheet.format("1:1", {
        "backgroundColor": {"red": 0.12, "green": 0.16, "blue": 0.23},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "horizontalAlignment": "CENTER",
        "wrapStrategy": "WRAP",
    })
    worksheet.freeze(rows=1)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Sync scraper results to Google Sheets")
    parser.add_argument("--sheet-id", default=os.environ.get("GOOGLE_SHEET_ID", DEFAULT_SHEET_ID),
                        help=f"Google Sheet ID (default: {DEFAULT_SHEET_ID}, "
                             "overridable via GOOGLE_SHEET_ID env var)")
    parser.add_argument("--xlsx", default="",
                        help="Path to xlsx file (default: latest in download/)")
    parser.add_argument("--download-dir", default="./download",
                        help="Directory containing xlsx files (default: ./download)")
    parser.add_argument("--creds", default="",
                        help="Path to Google service account JSON (default: auto-detect "
                             "via GOOGLE_APPLICATION_CREDENTIALS, then ./google-credentials.json, "
                             "then ~/google-credentials.json)")
    parser.add_argument("--incremental", action="store_true",
                        help="Incremental sync: append only new matches (deduped by Job URL). "
                             "Default is full sync (clear + rewrite).")
    return parser.parse_args(argv)


def main():
    args = parse_args()

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


def cli() -> None:
    """Console-script entry point (zero-arg callable)."""
    raise SystemExit(main())


if __name__ == "__main__":
    cli()
