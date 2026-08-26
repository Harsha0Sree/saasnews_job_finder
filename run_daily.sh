#!/usr/bin/env bash
#
# run_daily.sh — Daily entrypoint for the SaaSNews fresher job finder.
#
# This script:
#   1. Runs the scraper with --priority-recent to process the 200 newest
#      articles FIRST (ensures today's new funding rounds are never missed)
#   2. Falls back to --batch-size 500 to continue from checkpoint
#   3. Incrementally syncs new matches to Google Sheets
#
# The scraper auto-discovers all pages (no hardcoded page count).
# Resume support: if killed mid-run, just re-run this script.
#
# Configuration (env vars):
#   GOOGLE_SHEET_ID     Target Google Sheet (default below)
#   DRY_RUN=1           Print the commands instead of running them
#
# Cron setup:
#   0 9 * * * cd /path/to/saasnews-job-finder && bash run_daily.sh >> /var/log/saasnews_daily.log 2>&1

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# ---- Resolve the project interpreter (cron has no activated venv) ----
if [ -x ".venv/bin/python" ]; then
    PYTHON=".venv/bin/python"
elif command -v uv >/dev/null 2>&1; then
    echo "No .venv found — creating it with uv..."
    uv sync --extra browser
    PYTHON=".venv/bin/python"
else
    PYTHON="python3"
    echo "WARNING: using system python3 — dependencies may be missing." >&2
    echo "         Recommended: 'uv sync --extra browser' in $SCRIPT_DIR." >&2
fi

SHEET_ID="${GOOGLE_SHEET_ID:-1VAinuzwZfAP8IVEwcr8sLUZsNMcy1OiPmQhmhfNSvYM}"

run() {
    if [ "${DRY_RUN:-}" = "1" ]; then
        echo "DRY-RUN: $*"
    else
        "$@"
    fi
}

echo "============================================================"
echo "SaaSNews Daily Job Scan — $(date -u '+%Y-%m-%d %H:%M:%S UTC')"
echo "Interpreter : $PYTHON"
echo "Sheet       : $SHEET_ID"
echo "============================================================"

# ---- Step 0: Harvest user feedback + diagnose + self-tune ----
# Reads Applied/Feedback marks from the latest xlsx (and the Google Sheet when
# credentials are available) into download/feedback.jsonl BEFORE scraping.
# New violations are then diagnosed: each flagged job is re-fetched, its
# filters replayed, the missed signal isolated, and evidence-backed
# generalizations written to download/learned.jsonl — so today's run already
# benefits from yesterday's corrections AND their root causes.
echo ""
echo "[1/4] Watching feedback → diagnose → learned policy..."
FEEDBACK_ARGS=(--once --download-dir ./download)
CREDS_AVAILABLE=0
if [ -n "${GOOGLE_APPLICATION_CREDENTIALS:-}" ] || [ -f "google-credentials.json" ] || [ -f "$HOME/google-credentials.json" ]; then
    CREDS_AVAILABLE=1
    FEEDBACK_ARGS+=(--sheet-id "$SHEET_ID")
fi
run "$PYTHON" watch_feedback.py "${FEEDBACK_ARGS[@]}"
if [ "$CREDS_AVAILABLE" = "0" ]; then
    echo "  (Sheet feedback skipped — no credentials; xlsx feedback still collected.)"
fi

# ---- Step 1: Run the scraper ----
# --priority-recent 200: process the 200 newest articles FIRST (from page 1).
#                        This ensures today's new funding rounds are processed
#                        immediately, even if the checkpoint has old articles.
# --batch-size 500: after priority articles, process up to 500 more from checkpoint.
# --fresher-only: only fresher-eligible roles (0-2 years + internships)
# --concurrency 5: lower concurrency for Playwright compatibility
echo ""
echo "[2/4] Running scraper (priority=200 newest, batch=500, fresher-only)..."
run "$PYTHON" saasnews_scraper.py \
    --priority-recent 200 \
    --batch-size 500 \
    --concurrency 5 \
    --fresher-only \
    --output-dir ./download

echo ""
echo "[3/4] Syncing to Google Sheets (incremental)..."
# Only sync if service-account credentials are available.
if [ "$CREDS_AVAILABLE" = "1" ]; then
    run "$PYTHON" sync_to_google_sheets.py \
        --sheet-id "$SHEET_ID" \
        --download-dir ./download \
        --incremental
else
    echo "  ⚠ No Google credentials found — skipping Google Sheets sync."
    echo "  To enable: set GOOGLE_APPLICATION_CREDENTIALS or save the JSON key"
    echo "  as google-credentials.json in this directory."
    echo "  See README.md → 'Google Sheets sync' section."
fi

echo ""
echo "[4/4] Done — summary:"
echo "============================================================"
echo "Daily scan complete — $(date -u '+%Y-%m-%d %H:%M:%S UTC')"
echo "============================================================"
