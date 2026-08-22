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
# Cron setup:
#   0 9 * * * cd /path/to/saasnews-job-finder && bash run_daily.sh >> /var/log/saasnews_daily.log 2>&1

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "============================================================"
echo "SaaSNews Daily Job Scan — $(date -u '+%Y-%m-%d %H:%M:%S UTC')"
echo "============================================================"

# ---- Step 1: Run the scraper ----
# --priority-recent 200: process the 200 newest articles FIRST (from page 1).
#                        This ensures today's new funding rounds are processed
#                        immediately, even if the checkpoint has old articles.
# --batch-size 500: after priority articles, process up to 500 more from checkpoint.
# --fresher-only: only fresher-eligible roles (0-2 years + internships)
# --concurrency 5: lower concurrency for Playwright compatibility
echo ""
echo "[1/2] Running scraper (priority=200 newest, batch=500, fresher-only)..."
python3 saasnews_scraper.py \
    --priority-recent 200 \
    --batch-size 500 \
    --concurrency 5 \
    --fresher-only \
    --output-dir ./download

echo ""
echo "[2/2] Syncing to Google Sheets (incremental)..."
# Only sync if google-credentials.json exists.
if [ -f "google-credentials.json" ] || [ -n "${GOOGLE_APPLICATION_CREDENTIALS:-}" ]; then
    python3 sync_to_google_sheets.py \
        --sheet-id "1VAinuzwZfAP8IVEwcr8sLUZsNMcy1OiPmQhmhfNSvYM" \
        --download-dir ./download \
        --incremental
else
    echo "  ⚠ google-credentials.json not found — skipping Google Sheets sync."
    echo "  To enable: create a service account and save the JSON key as"
    echo "  google-credentials.json in this directory."
    echo "  See README.md → 'Google Sheets sync' section."
fi

echo ""
echo "============================================================"
echo "Daily scan complete — $(date -u '+%Y-%m-%d %H:%M:%S UTC')"
echo "============================================================"
