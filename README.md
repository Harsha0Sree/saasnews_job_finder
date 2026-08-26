# TheSaaSNews → Fresher Python-Stack Job Finder

Automatically scrapes funding-round announcements on
[thesaasnews.com/news/](https://www.thesaasnews.com/news/), follows each
company's careers page, parses open roles, and filters for **fresher-eligible
Python-stack backend / AI-ML / full-stack / data engineering jobs** (internships
OK) located in **India** or **remote-worldwide**.

## Setup (uv)

```bash
uv sync --extra browser          # runtime deps + Playwright
uv run playwright install chromium
uv sync --extra dev              # + pytest

# Run tests (deterministic, offline):
.venv/bin/python -m pytest

# Run the scraper:
uv run saasnews-scraper --fresher-only        # console script…
python saasnews_scraper.py --fresher-only     # …or directly

# Watch sheets for feedback and self-tune continuously (or --once):
uv run saasnews-watch
python watch_feedback.py

# Daily scan via the venv-aware cron entrypoint (DRY_RUN=1 to preview):
bash run_daily.sh
```

## Architecture

| Module | Responsibility |
|--------|----------------|
| `saasnews_scraper.py` | Pipeline orchestration: discovery → article parse → careers find → jobs parse → filter → xlsx |
| `saasnews/filters.py` | **Single source of truth** for job-evaluation policy: `match_role`, `match_location`, `is_fresher_role`, `extract_min_years` + regex tables. Pure functions, no I/O |
| `saasnews/feedback.py` | Closed-loop feedback: constraint taxonomy (multi-select), JSONL log, learned-policy compiler (`build_policy`) |
| `saasnews/diagnose.py` | Root-cause engine: replays the real filters against a re-fetched posting, isolates the missed signal, proposes evidence-backed learnings |
| `collect_feedback.py` | Harvests Applied/Feedback marks from xlsx + Google Sheet into `download/feedback.jsonl` |
| `watch_feedback.py` | Continuous watcher (`saasnews-watch`): polls sheets, harvests new feedback immediately, diagnoses each violation, writes learned.jsonl |
| `sync_to_google_sheets.py` | Google Sheets sync (incremental merge: newest at top, preserves your tracking columns) |
| `combine_results.py` | Merges run outputs into one master workbook using the shared filter tables |
| `tests/` | Offline deterministic pytest suite (`pytest -m network` opts into live-web checks) |

## Closed-loop feedback

Every row in the Matches feed carries three user-owned columns:

| Column | Values | Meaning |
|--------|--------|---------|
| `Applied` | Yes / No | application tracker (blank = not reviewed) |
| `Feedback` | constraint dropdown (**multi-select**: combine several with commas) | what constraint(s) this job actually violated |
| `Notes` | free text | context for future you |

Marking `Applied = Yes` does two things: the job **leaves the active feed**
(it moves to a dedicated `Applied` sheet in the xlsx, and sinks below the
fresh list in Google Sheets), and it is never re-emitted by later runs.

Constraints map 1:1 to the pipeline's real filters:
`not_fresher`, `not_python_stack`, `wrong_role`, `bad_location`,
`stale_or_closed`, `broken_link`, `duplicate_company`,
`irrelevant_company`, `other`. A cell can hold several — e.g.
`not_fresher, bad_location` — each becomes its own record (friendly
aliases like "Senior" or "Wrong stack" also work).

### The intelligent loop: watch → diagnose → self-tune

Run a watcher alongside your day and feedback takes effect immediately,
not on the next daily scrape:

```bash
python watch_feedback.py            # polls sheets every 120s (Ctrl-C to stop)
python watch_feedback.py --interval 60
```

Every cycle it:
1. **Harvests** new Applied/Feedback marks from the live Google Sheet + local xlsx.
2. **Goes back to each flagged job posting** — re-fetches its JD text and
   replays the exact same filters the pipeline runs — then isolates what was
   overlooked: JD unavailable at scrape time? a duration demand invisible to
   the years extractor ("24 months of experience")? a seniority marker
   missing from the tables ("SDE 2")? region evidence hiding in the JD body
   ("must be based in Texas")? the short-JD leniency branch accepting a
   non-Python role?
3. **Tweaks itself**, writing conservative, evidence-backed generalizations to
   `download/learned.jsonl` (title/JD patterns, region tokens, per-domain
   Python strictness) plus a full audit trail in `download/diagnosis.jsonl`.

The daily run closes the same loop in one pass (`watch_feedback.py --once`
replaces the old collect-only step), and every scrape compiles BOTH logs into
its learned policy:

```bash
bash run_daily.sh
# [1/4] watch_feedback.py --once → harvest your marks, diagnose misses, learn
# [2/4] saasnews_scraper.py     → compiles feedback.jsonl + learned.jsonl BEFORE scraping:
#         • blocked job URLs are never re-emitted
#         • flagged locations rejected exactly AND via learned region tokens
#         • companies with ≥3 flagged JOBS are skipped (multi-constraint rows
#           count once — detailed feedback is not double punishment)
#         • learned title/JD patterns tighten the fresher + role filters
#         • python_strict domains lose the short-JD leniency for the stack check
# [3/4] sync_to_google_sheets.py → newest-at-top merge, tracking columns preserved
```

Run Summary reports honest loop metrics every run: *Feedback Records Applied*,
*Learned Patterns Active* (+ per-kind breakdown), *Companies Skipped (learned
policy)*, *Known-Bad Jobs Suppressed*, plus a **Feedback Analysis** breakdown
(constraint × the pipeline signal that let the job through — e.g.
`not_fresher × no_jd_ambiguous_title: 3` tells you exactly where the fresher
filter leaks).

Learning stays auditable: every learned pattern names the job that taught it
and quotes the evidence snippet. Un-teach by deleting its line from
`download/learned.jsonl`; un-flag by deleting lines from
`download/feedback.jsonl`.

## Filtering pipeline

1. **Role match** — category patterns (Internship first), title exclusions
   (non-Python languages, sales/design roles).
2. **Location match** — India tokens or remote-worldwide; restricted-region
   tokens are checked in *both* location and title.
3. **Fresher filter** (`--fresher-only`) — multi-signal: explicit junior/intern
   titles pass, senior/staff/Sr./II titles reject, otherwise JD text is fetched
   (Playwright fallback for JS pages) and years-of-experience extracted;
   unverifiable cases accept as implicit (balanced).
4. **Python-stack verification** — title mention, else JD must mention
   Python/Django/FastAPI/PyTorch/etc.; short JD text stays lenient.

## Configuration

All paths default relative to the working directory; no machine-specific
absolute paths are baked in.

```bash
GOOGLE_SHEET_ID                  # override target sheet for sync/run_daily.sh
GOOGLE_APPLICATION_CREDENTIALS   # service-account JSON (or ./google-credentials.json)
DRY_RUN=1 bash run_daily.sh      # print commands instead of executing
```

Scraper CLI options:

```
--news-pages N       Max pages to scan (default: 0 = ALL, follow rel="next")
--start-page N       Start from page N (chunked runs)
--priority-recent N  Process N newest articles FIRST before resuming checkpoint
--batch-size N       Process at most N articles this run (0 = no limit)
--concurrency N      Parallel workers (default 5; use 4-5 with Playwright)
--limit N            Max companies this run
--since-days N       Only news from last N days
--fresher-only       Only fresher-eligible roles (0-2 years + internships)
--no-python-stack    Disable Python-stack verification
--reset-checkpoint   Delete checkpoint and start fresh
--output-dir DIR     Where to save xlsx (default: ./download)
--verbose, -v        Debug logging
```

## Resume after crash/kill

The JSONL checkpoint (`download/checkpoint.jsonl`) is appended atomically after
every company and fsynced. Re-running the same command skips processed
companies and continues. The xlsx is rebuilt from the checkpoint every 10
companies and at the end (atomic temp-file rename).

Per-company wall-clock budget is enforced across every stage (homepage,
careers probing, JD fetches) — a dead-slow site can never hold a worker much
past its budget.

## Google Sheets sync

One-time setup:
1. Create a Google Service Account, download the JSON key.
2. Point `GOOGLE_APPLICATION_CREDENTIALS` at it (or save as
   `google-credentials.json` in the project root).
3. Share your sheet with the service account email as Editor.
4. `uv run saasnews-sync --incremental`

Incremental sync reads existing Job URLs from the sheet and appends only new
rows — deduplication is automatic.

## Supported job boards

Greenhouse (US + EU) · Lever · Ashby · Workable (JSON APIs when detected) ·
any HTML careers page (JSON-LD `JobPosting` + heuristics) · JS-rendered SPA
pages (Playwright fallback).
