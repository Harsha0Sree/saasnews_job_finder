#!/usr/bin/env python3
"""
Watch the tracking sheets for your feedback and close the loop immediately.

Instead of waiting for the next daily scrape, this watcher:

  1. POLLS the Matches sheet (live Google Sheet when credentials exist, plus
     the local xlsx) every --interval seconds,
  2. HARVESTS new Applied/Feedback marks into download/feedback.jsonl
     (multi-select Feedback cells are split into one record per constraint),
  3. GOES BACK to each newly-flagged job posting: re-fetches its JD text and
     REPLAYS the real filters to figure out what was overlooked —
       • was the JD unavailable/too short at scrape time?
       • is there a duration demand ("24 months of experience") the years
         extractor cannot see?
       • a seniority marker missing from the tables ("SDE 2", "Experienced…")?
       • region evidence hiding in the JD body ("must be based in Texas")?
       • did the short-JD leniency branch accept a non-Python role?
  4. TWEAKS ITSELF: writes conservative, evidence-backed generalizations to
     download/learned.jsonl (title/JD patterns, region tokens, per-domain
     strictness) plus a full audit trail in download/diagnosis.jsonl. The
     next scrape compiles both logs into its learned policy — same KIND of
     mistake, never again. Un-teach anything by deleting its line.

Run it alongside your day:
    python watch_feedback.py                    # poll every 120s forever
    python watch_feedback.py --interval 60      # tighter loop
    python watch_feedback.py --once             # single harvest+diagnose pass
                                                 (used by run_daily.sh)
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone

from collect_feedback import DEFAULT_DOWNLOAD_DIR, harvest
from saasnews.diagnose import diagnose
from saasnews.feedback import (
    load_learned,
    learning_key,
)

DIAGNOSE_BUDGET_SECS = 25.0   # wall-clock cap per job re-fetch


def _feedback_path(download_dir: str) -> str:
    return os.path.join(download_dir, "feedback.jsonl")


def _learned_path(download_dir: str) -> str:
    return os.path.join(download_dir, "learned.jsonl")


def _diagnosis_path(download_dir: str) -> str:
    return os.path.join(download_dir, "diagnosis.jsonl")


def _append_jsonl(path: str, record: dict) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    line = json.dumps(record, ensure_ascii=False, default=str) + "\n"
    with open(path, "a", encoding="utf-8") as f:
        f.write(line)
        f.flush()
        os.fsync(f.fileno())


def diagnosis_key(record: dict) -> tuple[str, str]:
    """One diagnosis per (job_url, constraint): adding a second constraint on
    a job you already flagged still gets its own root-cause pass."""
    return (str(record.get("job_url", "")), str(record.get("constraint", "")))


def load_diagnosed(path: str) -> set[tuple[str, str]]:
    keys: set[tuple[str, str]] = set()
    if not os.path.exists(path):
        return keys
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if isinstance(rec, dict):
                    keys.add(diagnosis_key(rec))
    except OSError:
        pass
    return keys


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Watch sheets for feedback, diagnose misses, self-tune filters"
    )
    parser.add_argument("--interval", type=float, default=120.0,
                        help="Seconds between sheet checks (default: 120)")
    parser.add_argument("--once", action="store_true",
                        help="One harvest + diagnose pass, then exit")
    parser.add_argument("--download-dir", default=DEFAULT_DOWNLOAD_DIR,
                        help="Directory for feedback/learned/diagnosis JSONL "
                             "and the latest xlsx (default: ./download)")
    parser.add_argument("--sheet-id", default=os.environ.get("GOOGLE_SHEET_ID", ""),
                        help="Google Sheet ID to poll (optional; falls back to "
                             "GOOGLE_SHEET_ID env var)")
    parser.add_argument("--creds", default="",
                        help="Google service account JSON path (default: auto-detect)")
    parser.add_argument("--max-diagnose", type=int, default=5,
                        help="Max job postings to re-fetch per cycle (default: 5)")
    parser.add_argument("--no-diagnose", action="store_true",
                        help="Harvest only; skip re-fetching/reasoning about jobs")
    return parser.parse_args(argv)


# ---------------------------------------------------------------------------
# Diagnosis stage
# ---------------------------------------------------------------------------

def diagnose_new_records(new_records: list[dict], download_dir: str,
                         max_items: int = 5) -> dict:
    """Re-fetch + replay + learn for every not-yet-diagnosed violation."""
    from saasnews_scraper import fetch_jd_text, make_session

    diag_path = _diagnosis_path(download_dir)
    learned_path = _learned_path(download_dir)
    diagnosed = load_diagnosed(diag_path)
    learned_keys = {learning_key(e) for e in load_learned(learned_path)}

    stats = {"checked": 0, "diagnosed": 0, "skipped": 0, "learnings": 0,
             "fetch_failed": 0}
    session = make_session()

    violations = [r for r in new_records if r.get("constraint")]
    for rec in violations:
        key = diagnosis_key(rec)
        if key in diagnosed:
            stats["skipped"] += 1
            continue
        if stats["diagnosed"] >= max_items:
            break
        url = rec.get("apply_url") or rec.get("job_url") or ""
        deadline = time.monotonic() + DIAGNOSE_BUDGET_SECS
        jd_text = fetch_jd_text(session, url, deadline=deadline) if url else ""
        if not jd_text:
            stats["fetch_failed"] += 1
        result = diagnose(rec, jd_text)

        # Merge any harvested metadata the diagnoser should audit.
        merged = rec.get("fresher_basis") or ""
        _append_jsonl(diag_path, {
            "job_url": result.job_url,
            "constraint": result.constraint,
            "root_cause": result.root_cause,
            "detail": result.detail,
            "evidence": result.evidence,
            "learnings": [vars(l) for l in result.learnings],
            "fresher_basis_at_scrape": merged,
            "created_at": datetime.now(timezone.utc).isoformat(),
        })
        diagnosed.add(key)
        stats["diagnosed"] += 1

        print(f"  ✚ {result.constraint}: {result.root_cause} — {result.detail}")
        for learning in result.learnings:
            entry = {
                "kind": learning.kind,
                "value": learning.value,
                "scope": learning.scope,
                "evidence": learning.evidence,
                "job_url": result.job_url,
                "constraint": result.constraint,
                "source": "watch_feedback",
                "created_at": datetime.now(timezone.utc).isoformat(),
            }
            lkey = learning_key(entry)
            if lkey in learned_keys:
                continue
            _append_jsonl(learned_path, entry)
            learned_keys.add(lkey)
            stats["learnings"] += 1
            print(f"    ↳ learned [{learning.kind}] {learning.value}")
        stats["checked"] += 1
    return stats


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def run_cycle(args) -> None:
    stamp = datetime.now().strftime("%H:%M:%S")
    result = harvest(_feedback_path(args.download_dir),
                     download_dir=args.download_dir,
                     sheet_id=args.sheet_id, creds=args.creds)
    new_violations = sum(1 for r in result.new_records if r.get("constraint"))
    print(f"[{stamp}] harvest: {len(result.new_records)} new record(s) "
          f"({new_violations} violation(s)) from {len(result.sources)} source(s).")

    if result.new_records and not args.no_diagnose:
        stats = diagnose_new_records(result.new_records, args.download_dir,
                                     max_items=args.max_diagnose)
        print(f"  diagnosis: {stats['diagnosed']} reasoned, "
              f"{stats['skipped']} already known, {stats['learnings']} new "
              f"learning(s); next scrape applies them.")
    elif args.no_diagnose and new_violations:
        print("  (--no-diagnose: records logged; run without the flag or wait "
              "for the daily pass to reason over them.)")


def main(argv=None) -> int:
    args = parse_args(argv)
    os.makedirs(args.download_dir, exist_ok=True)
    mode = "single pass" if args.once else f"polling every {args.interval:g}s"
    print(f"Feedback watcher started ({mode}). Sources: xlsx"
          + (", Google Sheet" if args.sheet_id else "") + ". Ctrl-C to stop.")
    try:
        while True:
            try:
                run_cycle(args)
            except KeyboardInterrupt:
                raise
            except Exception as e:
                # A bad network/Sheet cycle must never kill the watcher.
                print(f"WARNING: cycle failed ({e}) — retrying next interval.",
                      file=sys.stderr)
            if args.once:
                break
            time.sleep(max(5.0, args.interval))
    except KeyboardInterrupt:
        print("\nWatcher stopped.")
    return 0


def cli() -> None:
    raise SystemExit(main())


if __name__ == "__main__":
    cli()
