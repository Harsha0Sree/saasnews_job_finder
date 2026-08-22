# TheSaaSNews → Fresher Python-Stack Job Finder (Production-Ready)

Automatically scrapes every funding-round announcement on
[thesaasnews.com/news/](https://www.thesaasnews.com/news/), follows each
company's careers page, parses open roles, and filters for **fresher-eligible
Python-stack backend / AI-ML / full-stack / data engineering jobs** (internships
OK) located in **India** or **remote-worldwide**.

## What makes this production-ready

### 1. Bulletproof experience filter
- **Multi-signal approach**: title keywords + JD-text years extraction + seniority detection
- **10 regex patterns** for years extraction: "5+ years", "3-5 years", "minimum 5 years", "at least 4 years", "experience: 5 years", "5 years building", "should have 5 years", etc.
- **Senior title detection**: Senior, Sr., Sr, Staff, Principal, Lead, Manager, Director, VP, CTO, Architect, Founding, II, III, IV (Roman numerals)
- **Fresher title detection**: Junior, Fresher, Intern, Fellow, Apprentice, SDE-1, L1, L2, entry-level, graduate

### 2. Playwright headless browser fallback
- Many SaaS careers pages are JavaScript SPAs (React/Next.js/Vue) — `requests.get` can't render them
- When `requests.get` returns short content (<200 chars, likely JS-rendered), the scraper automatically falls back to **Playwright Chromium** to render the page
- This works for both careers page job parsing AND individual JD text extraction
- Thread-safe (uses a lock to serialize Playwright access)

### 3. All 4 article formats supported
The site has changed formats over time. All are handled:
- **Format 1 (newest)**: `<br>` separators, `<a href>` links
- **Format 2 (mid)**: `&nbsp;</p><p>` separators
- **Format 3 (older)**: `<strong>` labels, plain-text URLs
- **Format 4 (oldest)**: No Funding Details block

### 4. Never hangs, never loses progress
- Hard 90s per-company wall-clock budget
- Separate connect (6s) / read (12s) timeouts
- JSONL checkpoint after every company (atomic append + fsync)
- XLSX rewritten every 10 companies
- Resume support — re-running skips processed companies

## Quick start

```bash
# 1. Install dependencies (including Playwright)
pip install -r requirements.txt
playwright install chromium

# 2. Run tests (verify everything works — 188 tests)
python tests/run_tests.py

# 3. Run the scraper (scans ALL pages by default)
python saasnews_scraper.py --fresher-only

# 4. Sync to Google Sheets (needs google-credentials.json)
python sync_to_google_sheets.py --incremental

unzip saasnews-job-finder.zip && cd saasnews-job-finder
pip install -r requirements.txt
playwright install chromium
python tests/run_tests.py                    # 198 tests pass

# Daily scan (processes newest articles first, then resumes):
python saasnews_scraper.py --fresher-only --priority-recent 200 --batch-size 500

# Full scan (all pages, no priority):
python saasnews_scraper.py --fresher-only

# Cron job:
bash run_daily.sh
```

## CLI options

```
python saasnews_scraper.py [options]

  --news-pages N       Max pages to scan (default: 10000 = all)
  --start-page N       Start from page N (for chunked runs)
  --concurrency N      Parallel workers (default: 10; use 4-5 with Playwright)
  --fresher-only       Only fresher-eligible roles (0-2 years exp + internships)
  --no-python-stack    Disable Python-stack verification
  --batch-size N       Process at most N articles this run (0 = no limit)
  --since-days N       Only news from last N days
  --limit N            Process at most N companies (0 = no limit)
  --output-dir DIR     Where to save xlsx (default: ./download)
  --reset-checkpoint   Delete checkpoint and start fresh
  --verbose, -v        Debug logging
```

## How the experience filter works

| Title | JD available? | JD years? | Result |
|-------|---------------|-----------|--------|
| Junior Python Developer | — | — | ✅ fresher (title explicit) |
| Software Engineer Intern | — | — | ✅ fresher (title explicit) |
| Senior Software Engineer | — | — | ❌ rejected (title says Senior) |
| Sr. Forward Deployed Engineer | — | — | ❌ rejected (Sr. detected) |
| Software Engineer II | — | — | ❌ rejected (II = senior) |
| Staff Engineer | — | — | ❌ rejected (Staff) |
| Backend Engineer | Yes (300+ chars) | 0-2 years | ✅ fresher (JD verified) |
| Backend Engineer | Yes (300+ chars) | 3+ years | ❌ rejected (JD says 3+) |
| Backend Engineer | Yes (300+ chars) | no years mentioned | ✅ fresher (implicit) |
| Backend Engineer | No JD (JS-rendered, Playwright fallback) | — | ✅ fresher (implicit) |

## Resume after crash/kill

```bash
# If killed, just re-run the same command:
python saasnews_scraper.py --fresher-only

# It loads checkpoint FIRST, skips processed companies, continues from where it left off.
```

## Google Sheets sync

### One-time setup
1. Create a Google Service Account at https://console.cloud.google.com/iam-admin/serviceaccounts
2. Download the JSON key, save as `google-credentials.json`
3. Share your Google Sheet with the service account email as Editor
4. Run: `python sync_to_google_sheets.py --incremental`

### How incremental sync works
- Reads existing Job URLs from the sheet
- Appends only rows whose Job URL isn't already there
- Deduplication is automatic

## Daily cron job

```bash
crontab -e
# Add:
0 9 * * * cd /path/to/saasnews-job-finder && bash run_daily.sh >> /var/log/saasnews.log 2>&1
```

## Testing

```bash
python tests/run_tests.py
```

**188 tests** covering:
- Role matching (Python-stack only)
- Internship matching
- Fresher filter (title + JD text verification)
- Seniority detection (Sr., II, III, Roman numerals)
- Years-of-experience extraction (10 regex patterns, 33 test cases)
- Location matching (India + Remote-Worldwide)
- Checkpoint + resume
- Python-stack verification
- Article parser (all 4 formats)
- Playwright availability
- Edge cases

## Files

```
saasnews-job-finder/
├── saasnews_scraper.py           # Main scraper (Playwright + requests hybrid)
├── sync_to_google_sheets.py      # Google Sheets sync (full + incremental)
├── combine_results.py            # Merge multiple runs
├── run_daily.sh                  # Daily cron entrypoint
├── tests/
│   └── run_tests.py              # 188-test suite
├── README.md
├── requirements.txt              # Python dependencies (incl. playwright)
└── download/
    ├── saasnews_jobs_latest.xlsx
    └── checkpoint.jsonl
```

## Troubleshooting

**"No module named 'playwright'"** → `pip install playwright && playwright install chromium`

**"Playwright init failed"** → Run `playwright install chromium` to install the browser.

**"No matches found"** → The filter is strict. Try without `--fresher-only` first, or check if companies in your target pages have careers pages.

**Script seems slow** → Playwright fallback adds ~5s per JS-rendered page. Reduce `--concurrency` to 4-5 if using Playwright. Use `--batch-size 500` for chunked processing.

**Want to start fresh** → `python saasnews_scraper.py --reset-checkpoint --fresher-only`

## Supported job boards
- **Greenhouse** (US + EU, JSON API)
- **Lever** (JSON API)
- **Ashby** (JSON API)
- **Workable** (JSON API)
- **Any HTML careers page** (JSON-LD JobPosting + heuristic parsing)
- **JS-rendered SPA careers pages** (Playwright fallback)
