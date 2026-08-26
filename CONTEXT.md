# CONTEXT — saasnews-job-finder

Domain glossary and architecture notes for this automation. Read this before
touching filters or the pipeline; keep names aligned with it.

## Domain terms

- **Funding Article** (`NewsItem`): one `/news/<slug>/` page on
  thesaasnews.com announcing a funding round. Carries company name/website,
  round, date, lead investor, software category. Four historical HTML formats
  exist; the parser handles all.
- **Company**: a startup extracted from a Funding Article (must have a website).
- **Careers Page**: the company URL where open roles are listed. Found via
  homepage link heuristics or common-path probing under the caller's deadline.
- **JobPosting** (`JobPosting`): one open role parsed from a careers page —
  title, url, location, apply_url, posted_date. Sourced from hosted boards
  (Greenhouse/Lever/Ashby/Workable JSON APIs) or generic HTML/JSON-LD.
- **Match** (`JobMatch`): a JobPosting that passed role + location (+ fresher +
  python-stack when enabled) filters, enriched with funding metadata.
- **Fresher Filter**: policy deciding if a role suits 0–2 years experience.
  Balanced: explicit senior signals reject; unverifiable cases accept implicit.
- **Python-stack filter**: requires Python-stack mention in title or JD;
  lenient only when JD text is unusably short (<300 chars).
- **Checkpoint** (`download/checkpoint.jsonl`): append-only per-company run
  journal enabling resume; source of truth rebuilt into xlsx.
- **Run Summary**: workbook sheet of honest run metrics (never fabricated).
- **First Seen** (`scraped_at`): UTC timestamp stamped when a match was
  scraped; the feed sort key (newest at top in xlsx, master file, and Sheets).
- **Tracking Columns** (`Applied` / `Feedback` / `Notes`): user-owned sheet
  columns. Sync preserves them by Job URL; the xlsx regenerates them blank.
- **Feedback Record**: one line of `download/feedback.jsonl` — a job plus one
  violated constraint (`saasnews.feedback.CONSTRAINTS`, mirroring the real
  filter dimensions) and optional note/source/timestamp. Feedback cells are
  **multi-select**: comma-separated constraints each become their own record.
- **Learned Policy** (`build_policy`): compiled from feedback records before
  every scrape — blocked job URLs (exact), flagged locations (exact,
  normalized via `normalize_location`), and per-company strikes
  (≥3 flagged JOBS → company checkpointed as `skipped_policy`; multi-constraint
  rows count once). Never generalized beyond exact matches on its own;
  un-teach by deleting lines from feedback.jsonl.
- **Watcher** (`watch_feedback.py`, `saasnews-watch`): polls the sheets,
  harvests new feedback immediately, then for every new violation re-fetches
  the posting, replays `saasnews.filters`, isolates the missed signal, and
  writes evidence-backed generalizations to `download/learned.jsonl` plus an
  audit trail in `download/diagnosis.jsonl`.
- **Diagnosis** (`saasnews/diagnose.py`, pure): per-constraint root cause —
  e.g. `jd_unavailable_at_scrape`, `duration_format_missed_in_jd`,
  `seniority_word_missed_in_title`, `region_evidence_in_jd_body`,
  `lenient_short_jd_accepted`. Proposes at most conservative learnings; never
  fabricates signal it cannot quote.
- **Learnings** (`download/learned.jsonl`): kinds `title_negative`,
  `jd_negative_pattern`, `title_exclusion`, `location_token`,
  `python_strict_domain`. Compiled by `compile_learned` into Policy tables
  that extend (never replace) the built-in regex tables; applied via optional
  params on `match_role` / `match_location` / `is_fresher_role` and the
  python-strict gate in process_company. Un-teach by deleting lines.

## Architecture notes

- `saasnews/filters.py` is the **policy seam**: pure functions + regex tables,
  consumed by scraper AND combine_results. Never duplicate its tables.
- `saasnews/feedback.py` is the **feedback seam**: harvesting, dedupe, policy
  compilation live there; the scraper only consumes the compiled Policy.
- `saasnews_scraper.py` is orchestration + HTTP + Playwright + output only;
  no filter policy lives there.
- Wall-clock budgets: every stage honours the caller's `deadline`; never reset
  a budget you were given.
- Sheet ordering invariant: Matches rows are sorted by First Seen descending;
  user-owned columns always win over regenerated values on merge.
