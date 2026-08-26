"""Closed-loop feedback: user-annotated constraint violations harvested from
the tracking sheets, compiled into a learned policy applied on every run.

Flow:
  1. User opens a job from xlsx / Google Sheet and marks what was wrong in the
     `Feedback` cell — one constraint OR several combined with commas
     (e.g. ``not_fresher, bad_location``). Values come from CONSTRAINTS.
  2. collect_feedback.py / watch_feedback.py harvest those rows into an
     append-only JSONL log (`download/feedback.jsonl`) — crash-safe and
     human-editable.
  3. watch_feedback.py diagnoses every new violation: it re-fetches the job
     posting, replays the real filters, pinpoints which signal was missed,
     and records evidence-backed generalizations in `download/learned.jsonl`
     (see LEARNING_KINDS). Every learning is auditable; un-teach by deleting
     its line.
  4. Every scraper run compiles BOTH logs into a Policy:
        - blocked job URLs are never re-emitted,
        - flagged locations are rejected by match_location (exact strings
          PLUS learned region tokens),
        - companies accumulating >= strike_threshold strikes are skipped
          (one strike per flagged job, however many constraints you marked),
        - learned title/JD patterns tighten the fresher + role filters,
        - python_strict domains lose the short-JD leniency for the stack check.
  5. Run Summary reports honest suppression + learning counts, closing the loop.
"""
from __future__ import annotations

import json
import logging
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from saasnews.filters import normalize_location

log = logging.getLogger("saasnews")

# Constraint taxonomy — mirrors exactly the filter dimensions the pipeline
# enforces. One value per way a scraped job can fail your real-world bar.
CONSTRAINTS = [
    "not_fresher",          # passed fresher filter but demands experience
    "not_python_stack",     # passed stack check but isn't Python work
    "wrong_role",           # not backend/AI/data/full-stack/intern after all
    "bad_location",         # not India, or "remote" turned out restricted
    "stale_or_closed",      # posting already filled/closed
    "broken_link",          # careers/JD link unusable
    "duplicate_company",    # same company spamming near-identical roles
    "irrelevant_company",   # not a software company worth tracking
    "other",
]
_CONSTRAINT_SET = frozenset(CONSTRAINTS)

# Friendly spellings accepted inside Feedback cells (mapped to constraints).
_CONSTRAINT_ALIASES = {
    "not fresher": "not_fresher",
    "senior": "not_fresher",
    "needs experience": "not_fresher",
    "not python": "not_python_stack",
    "not python stack": "not_python_stack",
    "wrong stack": "not_python_stack",
    "wrong role": "wrong_role",
    "bad location": "bad_location",
    "wrong location": "bad_location",
    "remote us": "bad_location",
    "stale": "stale_or_closed",
    "closed": "stale_or_closed",
    "filled": "stale_or_closed",
    "broken link": "broken_link",
    "broken": "broken_link",
    "dead link": "broken_link",
    "dup": "duplicate_company",
    "duplicate": "duplicate_company",
    "irrelevant": "irrelevant_company",
}

# Cell separator for multi-select feedback (any of these splits constraints).
_FEEDBACK_SPLIT_RE = re.compile(r"[,;/|\n]+")


def parse_feedback_cell(value: str) -> tuple[list[str], list[str]]:
    """Parse one Feedback cell into (valid_constraints, unknown_tokens).

    Multi-select: several constraints may be combined in a single cell,
    separated by commas/semicolons/slashes/pipes/newlines — e.g.
    ``not_fresher, bad_location``. Order is preserved; duplicates collapse;
    friendly aliases ("Senior", "Wrong stack") map to canonical constraints;
    anything unrecognized is returned separately so the harvester can warn
    without losing the valid parts of the cell.
    """
    valid: list[str] = []
    unknown: list[str] = []
    for token in _FEEDBACK_SPLIT_RE.split(str(value or "")):
        token = token.strip().lower()
        if not token:
            continue
        token = _CONSTRAINT_ALIASES.get(token, token)
        if token in _CONSTRAINT_SET:
            if token not in valid:
                valid.append(token)
        elif token not in unknown:
            unknown.append(token)
    return valid, unknown


# Kinds of evidence-backed generalizations the diagnoser may write into
# download/learned.jsonl. Each entry: {kind, value, scope, evidence,
# job_url, constraint, created_at}. Applied at runtime by compile_learned().
LEARNING_KINDS = [
    "title_negative",       # extra FRESHER_TITLE_NEGATIVE alternative
    "jd_negative_pattern",  # extra JD-text negative regex (fresher filter)
    "title_exclusion",      # extra TITLE_EXCLUSIONS alternative (role filter)
    "location_token",       # extra restricted-region token (location filter)
    "python_strict_domain", # domain where short-JD leniency is disabled
]
_LEARNING_KIND_SET = frozenset(LEARNING_KINDS)

# Safety cap: runaway learning files cannot slow the pipeline down.
MAX_LEARNED_PER_KIND = 200


def domain_of(url: str) -> str:
    """Registrable-ish domain key for company-level learning ('www.' stripped)."""
    from urllib.parse import urlparse
    if not url:
        return ""
    netloc = urlparse(url).netloc.lower()
    return netloc[4:] if netloc.startswith("www.") else netloc


def load_feedback(path: str) -> list[dict]:
    """Load feedback records from the JSONL log; corrupt lines are skipped."""
    records: list[dict] = []
    if not os.path.exists(path):
        return records
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
                if isinstance(rec, dict) and rec.get("job_url"):
                    records.append(rec)
    except OSError as e:
        log.warning("Could not read feedback file %s: %s", path, e)
    return records


def append_feedback(path: str, record: dict) -> None:
    """Atomically append one feedback record (O_APPEND write + fsync)."""
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    line = json.dumps(record, ensure_ascii=False, default=str) + "\n"
    with open(path, "a", encoding="utf-8") as f:
        f.write(line)
        f.flush()
        os.fsync(f.fileno())


def dedupe_feedback(existing: list[dict],
                    candidates: list[dict]) -> list[dict]:
    """Drop candidates already logged (same job_url + constraint) and
    collapse duplicates within a single harvest. Order is preserved."""
    seen: set[tuple[str, str]] = {
        (str(r.get("job_url", "")), str(r.get("constraint", "")))
        for r in existing
    }
    out: list[dict] = []
    for c in candidates:
        key = (str(c.get("job_url", "")), str(c.get("constraint", "")))
        if key in seen or not key[0]:
            continue
        seen.add(key)
        out.append(c)
    return out


# ---------------------------------------------------------------------------
# Learned generalizations (evidence-backed, diagnoser-written)
# ---------------------------------------------------------------------------

def learning_key(entry: dict) -> tuple[str, str, str]:
    """Identity of a learned pattern: same kind+scope+value never duplicates.
    Value comparison is case-insensitive (everything compiles IGNORECASE)."""
    return (str(entry.get("kind", "")),
            str(entry.get("scope", "")).strip().lower(),
            str(entry.get("value", "")).strip().lower())


def load_learned(path: str) -> list[dict]:
    """Load learned-pattern entries from the JSONL log; corrupt or malformed
    lines are skipped. Same tolerance contract as load_feedback."""
    entries: list[dict] = []
    if not os.path.exists(path):
        return entries
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    entry = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if (isinstance(entry, dict)
                        and str(entry.get("kind", "")) in _LEARNING_KIND_SET
                        and str(entry.get("value", "")).strip()):
                    entries.append(entry)
    except OSError as e:
        log.warning("Could not read learned file %s: %s", path, e)
    return entries


def compile_learned(entries: list[dict]) -> dict:
    """Validate + compile raw learned entries into runtime filter tables.

    Returns a dict with compiled regex lists per kind, the combined
    word-bounded region-token regex, and strict-python domains. Invalid regex
    values are skipped loudly (they stay in the JSONL so you can fix/delete).
    """
    tables: dict = {
        "title_negatives": [],
        "jd_negatives": [],
        "title_exclusions": [],
        "region_tokens": [],
        "region_re": None,
        "python_strict_domains": set(),
        "counts": Counter(),
        "skipped": 0,
    }
    for entry in entries:
        kind = str(entry.get("kind", ""))
        value = str(entry.get("value", "")).strip()
        if kind not in _LEARNING_KIND_SET or not value:
            continue
        if tables["counts"][kind] >= MAX_LEARNED_PER_KIND:
            tables["skipped"] += 1
            continue
        tables["counts"][kind] += 1
        if kind == "python_strict_domain":
            d = domain_of(value)
            if not d:
                # Bare domains ("acme.io") are accepted as-is.
                candidate = value.strip().lower()
                if re.fullmatch(r"[a-z0-9.\-]+", candidate):
                    d = candidate[4:] if candidate.startswith("www.") else candidate
            if d:
                tables["python_strict_domains"].add(d)
            else:
                tables["counts"][kind] -= 1
                tables["skipped"] += 1
            continue
        try:
            pat = re.compile(value, re.IGNORECASE)
        except re.error as e:
            log.warning("Skipping invalid learned %s %r: %s", kind, value, e)
            tables["counts"][kind] -= 1
            tables["skipped"] += 1
            continue
        if kind == "title_negative":
            tables["title_negatives"].append(pat)
        elif kind == "jd_negative_pattern":
            tables["jd_negatives"].append(pat)
        elif kind == "title_exclusion":
            tables["title_exclusions"].append(pat)
        elif kind == "location_token":
            # Tokens are literals (word-bounded), never user-supplied regex.
            tables["region_tokens"].append(value.lower())
    if tables["region_tokens"]:
        escaped = "|".join(re.escape(t)
                           for t in sorted(set(tables.pop("region_tokens"))))
        tables["region_re"] = re.compile(rf"\b(?:{escaped})\b", re.IGNORECASE)
    return tables


def summarize_learned(entries: list[dict], top: int = 8) -> list[str]:
    """One-line-per-kind tally of active learnings for the Run Summary."""
    counts = Counter(str(e.get("kind", "")) for e in entries
                     if str(e.get("kind", "")) in _LEARNING_KIND_SET)
    return [f"{kind}: {n}" for kind, n in counts.most_common(top)]


@dataclass
class Policy:
    """Learned policy compiled from feedback records + diagnosed learnings."""
    blocked_job_urls: set[str] = field(default_factory=set)
    blocked_locations: set[str] = field(default_factory=set)  # normalized
    applied_job_urls: set[str] = field(default_factory=set)
    company_strikes: Counter = field(default_factory=Counter)
    strike_threshold: int = 3
    # Learned generalizations (from download/learned.jsonl, compiled).
    learned_title_negatives: list = field(default_factory=list)   # [re.Pattern]
    learned_jd_negatives: list = field(default_factory=list)      # [re.Pattern]
    learned_title_exclusions: list = field(default_factory=list)  # [re.Pattern]
    learned_region_re: Optional[re.Pattern] = None
    python_strict_domains: set[str] = field(default_factory=set)

    def company_suppressed(self, website_url: str) -> bool:
        d = domain_of(website_url)
        return bool(d) and self.company_strikes[d] >= self.strike_threshold

    def job_suppressed(self, job_url: str, location: str = "") -> bool:
        if job_url and job_url in self.blocked_job_urls:
            return True
        if location and normalize_location(location) in self.blocked_locations:
            return True
        return False

    def python_strict(self, website_url: str) -> bool:
        """True when this domain lost the short-JD leniency for the stack check."""
        d = domain_of(website_url)
        return bool(d) and d in self.python_strict_domains


def build_policy(records: list[dict],
                 strike_threshold: int = 3,
                 learned_entries: Optional[list[dict]] = None) -> Policy:
    """Compile raw feedback records (+ optional diagnosed learnings) into the
    learned policy.

    - Every record (violation OR Applied=Yes) blocks its exact job URL from
      re-entering the feed.
    - Only real violations add company strikes — at most ONE strike per
      flagged job, however many constraints you marked on it (detailed
      multi-select feedback must not punish a company harder than a lazy one).
    - bad_location additionally learns the exact normalized location string.
    - learned_entries (see LEARNING_KINDS / compile_learned) extend the
      fresher/role/location filters and may disable short-JD leniency for
      specific domains.
    """
    policy = Policy(strike_threshold=max(1, strike_threshold))
    struck_jobs: set[tuple[str, str]] = set()
    for rec in records:
        url = rec.get("job_url", "")
        if url:
            policy.blocked_job_urls.add(url)
            if str(rec.get("applied", "")).lower() == "yes":
                policy.applied_job_urls.add(url)
        constraint = rec.get("constraint", "")
        domain = rec.get("company_domain") or domain_of(rec.get("company_website", ""))
        if constraint == "bad_location" and rec.get("location"):
            policy.blocked_locations.add(normalize_location(rec["location"]))
        if constraint and domain and (domain, url) not in struck_jobs:
            struck_jobs.add((domain, url))
            policy.company_strikes[domain] += 1
    if learned_entries:
        tables = compile_learned(learned_entries)
        policy.learned_title_negatives = tables["title_negatives"]
        policy.learned_jd_negatives = tables["jd_negatives"]
        policy.learned_title_exclusions = tables["title_exclusions"]
        policy.learned_region_re = tables["region_re"]
        policy.python_strict_domains = tables["python_strict_domains"]
    return policy


def summarize_feedback(records: list[dict], top: int = 5) -> list[str]:
    """Analyse violations: which constraint failed, and via which pipeline
    signal (fresher_basis) the job slipped through. Rendered into the Run
    Summary so every run shows what your feedback is teaching the filters.
    Applied-only records are tracking, not violations — excluded."""
    counts: Counter = Counter()
    for r in records:
        constraint = r.get("constraint", "")
        if not constraint:
            continue
        basis = r.get("fresher_basis") or "-"
        counts[(constraint, basis)] += 1
    return [f"{c} × {b}: {n}" for (c, b), n in counts.most_common(top)]


def extract_feedback_from_rows(headers: list[str], rows: list[list],
                               source: str) -> list[dict]:
    """Harvest user tracking rows (from xlsx Matches sheet or Google Sheet).

    A row contributes records when either:
      - its Feedback cell names one or more valid constraints (multi-select:
        combine several with commas — each becomes its own record), or
      - its Applied cell is "Yes" (tracking — the job leaves the active feed).
    Records carry the pipeline's own metadata for the row (fresher basis and
    level, category, matched keyword, apply URL) so every piece of feedback
    can be re-analysed against the signal that produced it.
    """
    hmap = {h.strip().lower(): i for i, h in enumerate(headers) if h}
    needed = ("job url",)
    if any(n not in hmap for n in needed):
        return []

    def cell(row, name):
        i = hmap.get(name)
        return str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else ""

    out: list[dict] = []
    for row in rows:
        job_url = cell(row, "job url")
        if not job_url:
            continue
        applied = cell(row, "applied").lower()
        constraints, unknown = parse_feedback_cell(cell(row, "feedback"))
        for token in unknown:
            log.warning("Skipping unknown feedback value %r (valid: %s)",
                        token, ", ".join(CONSTRAINTS))
        if not constraints:
            if applied != "yes":
                continue
            constraints = [""]
        for constraint in constraints:
            out.append({
                "job_url": job_url,
                "company_website": cell(row, "company website"),
                "location": cell(row, "location"),
                "job_title": cell(row, "job title"),
                "constraint": constraint,
                "note": cell(row, "notes"),
                "applied": applied,
                # Filter metadata — why the pipeline let this job through.
                "fresher_basis": cell(row, "fresher basis"),
                "fresher_level": cell(row, "fresher level"),
                "category": cell(row, "category"),
                "matched_keyword": cell(row, "matched keyword"),
                "apply_url": cell(row, "apply url"),
                "source": source,
            })
    return out
