#!/usr/bin/env python3
"""
Harvest user feedback from the tracking sheets into the learned-policy log.

Where the feedback comes from:
  - The Matches sheet of the latest xlsx (download/saasnews_jobs_latest.xlsx)
  - Optionally the live Google Sheet (--sheet-id or GOOGLE_SHEET_ID env),
    so feedback typed directly into Sheets is picked up too.

A row contributes when its `Feedback` cell names one or more constraints
(see saasnews.feedback.CONSTRAINTS — multi-select is supported: combine
values with commas). Records are appended to download/feedback.jsonl,
deduped by (job_url, constraint). The next scraper run compiles that log
into its learned policy; watch_feedback.py additionally diagnoses each new
violation and writes evidence-backed generalizations to learned.jsonl.

Usage:
  python collect_feedback.py                     # xlsx only
  python collect_feedback.py --sheet-id <ID>     # also read the Google Sheet
  python collect_feedback.py --dry-run           # report without appending

The harvest() helper is reused by watch_feedback.py for continuous polling.
"""
from __future__ import annotations

import argparse
import glob
import os
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timezone

from openpyxl import load_workbook

from saasnews.feedback import (
    CONSTRAINTS,
    append_feedback,
    dedupe_feedback,
    extract_feedback_from_rows,
    load_feedback,
)

DEFAULT_DOWNLOAD_DIR = "./download"


@dataclass
class HarvestResult:
    """Outcome of one harvest pass over all configured sources."""
    sources: list[str] = field(default_factory=list)
    candidates: int = 0
    already_logged: int = 0
    new_records: list[dict] = field(default_factory=list)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Harvest sheet feedback into the learned-policy log"
    )
    parser.add_argument("--feedback-file",
                        default=os.path.join(DEFAULT_DOWNLOAD_DIR, "feedback.jsonl"),
                        help="Append-only feedback log (default: ./download/feedback.jsonl)")
    parser.add_argument("--xlsx", default="",
                        help="Path to an xlsx file (default: latest in download-dir)")
    parser.add_argument("--download-dir", default=DEFAULT_DOWNLOAD_DIR,
                        help="Directory containing run xlsx files (default: ./download)")
    parser.add_argument("--sheet-id", default=os.environ.get("GOOGLE_SHEET_ID", ""),
                        help="Google Sheet ID to read feedback from (optional; "
                             "falls back to GOOGLE_SHEET_ID env var)")
    parser.add_argument("--creds", default="",
                        help="Google service account JSON path (default: auto-detect)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Report what would be appended without writing")
    return parser.parse_args(argv)


def find_xlsx(download_dir: str, explicit: str = "") -> str:
    if explicit:
        if not os.path.exists(explicit):
            print(f"ERROR: xlsx not found: {explicit}", file=sys.stderr)
            sys.exit(1)
        return explicit
    latest = os.path.join(download_dir, "saasnews_jobs_latest.xlsx")
    if os.path.exists(latest):
        return latest
    files = sorted(glob.glob(os.path.join(download_dir, "saasnews_jobs_*.xlsx")))
    if not files:
        print(f"ERROR: no xlsx files in {download_dir}", file=sys.stderr)
        sys.exit(1)
    return files[-1]


def rows_from_xlsx(path: str) -> tuple[list[str], list[list]]:
    wb = load_workbook(path, read_only=True)
    if "Matches" not in wb.sheetnames:
        print(f"ERROR: 'Matches' sheet missing in {path}", file=sys.stderr)
        sys.exit(1)
    ws = wb["Matches"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    headers = [str(h) if h is not None else "" for h in rows[0]]
    return headers, [list(r) for r in rows[1:]]


def rows_from_google_sheet(sheet_id: str, creds_path: str) -> tuple[list[str], list[list]]:
    """Best-effort read of the Matches sheet; raises on any failure."""
    import gspread
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession

    from sync_to_google_sheets import SCOPES

    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    gc = gspread.Client(auth=AuthorizedSession(creds))
    spreadsheet = gc.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet("Matches")
    all_values = worksheet.get_all_values()
    if not all_values:
        return [], []
    return all_values[0], all_values[1:]


def _collect_sources(download_dir: str, xlsx: str = "",
                     sheet_id: str = "", creds: str = ""
                     ) -> list[tuple[str, list[str], list[list]]]:
    """Gather (source_name, headers, rows) from every configured source.

    The local xlsx is always read; the live Google Sheet is best-effort and
    silently skipped when unavailable (no creds / offline).
    """
    sources: list[tuple[str, list[str], list[list]]] = []

    xlsx_path = find_xlsx(download_dir, xlsx)
    headers, rows = rows_from_xlsx(xlsx_path)
    sources.append((f"xlsx:{os.path.basename(xlsx_path)}", headers, rows))

    if sheet_id:
        try:
            from sync_to_google_sheets import find_credentials
            creds_path = creds or find_credentials()
            if not creds_path:
                raise RuntimeError("no service-account credentials found")
            g_headers, g_rows = rows_from_google_sheet(sheet_id, creds_path)
            sources.append((f"sheet:{sheet_id[:8]}…", g_headers, g_rows))
        except Exception as e:
            print(f"WARNING: could not read Google Sheet ({e}) — "
                  "continuing with local xlsx only.", file=sys.stderr)
    return sources


def harvest(feedback_file: str, download_dir: str = DEFAULT_DOWNLOAD_DIR,
            xlsx: str = "", sheet_id: str = "", creds: str = "",
            append: bool = True) -> HarvestResult:
    """One idempotent pass: read all sources, dedupe against the log,
    optionally append new records. Returns what was (or would be) appended —
    the exact input watch_feedback.py needs for diagnosis."""
    result = HarvestResult()
    existing = load_feedback(feedback_file)
    result.already_logged = len(existing)

    candidates: list[dict] = []
    for source, hdrs, rws in _collect_sources(download_dir, xlsx, sheet_id, creds):
        result.sources.append(source)
        found = extract_feedback_from_rows(hdrs, rws, source=source)
        for rec in found:
            rec.setdefault("created_at", datetime.now(timezone.utc).isoformat())
        candidates.extend(found)
    result.candidates = len(candidates)

    new_records = dedupe_feedback(existing, candidates)
    if new_records and append and not os.path.exists(
            os.path.dirname(feedback_file) or "."):
        os.makedirs(os.path.dirname(feedback_file), exist_ok=True)
    for rec in new_records:
        if append:
            append_feedback(feedback_file, rec)
    result.new_records = new_records
    return result


def main(argv=None) -> int:
    args = parse_args(argv)

    result = harvest(args.feedback_file, download_dir=args.download_dir,
                     xlsx=args.xlsx, sheet_id=args.sheet_id, creds=args.creds,
                     append=not args.dry_run)

    counts = Counter(r["constraint"] for r in result.new_records)
    print(f"Feedback harvest: {result.candidates} candidate row(s) from "
          f"{len(result.sources)} source(s); {result.already_logged} already logged; "
          f"{len(result.new_records)} new.")
    for constraint in CONSTRAINTS:
        if counts.get(constraint):
            print(f"  {constraint}: {counts[constraint]}")

    if not result.new_records:
        return 0
    if args.dry_run:
        print("DRY-RUN: nothing appended.")
        return 0

    print(f"✓ Appended {len(result.new_records)} record(s) to {args.feedback_file}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
