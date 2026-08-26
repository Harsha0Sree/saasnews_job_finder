"""Root-cause diagnosis for feedback records: WHY did this job slip through,
and what can be learned so the same KIND of mistake never happens again?

Given one harvested feedback record (saasnews.feedback) plus a fresh fetch of
the job posting's text, the diagnoser:
  1. REPLAYS the real filters (saasnews.filters — the same functions the
     pipeline runs) against the title/location/JD,
  2. DIFFS the outcome against what happened at scrape time to isolate the
     missed signal (JD unavailable then? years format not parsed? region
     token only in the JD body? lenient short-JD acceptance?),
  3. PROPOSES at most conservative, evidence-backed generalizations
     (LEARNING_KINDS in saasnews.feedback) — each carrying the exact snippet
     that justifies it, written to learned.jsonl by the caller, deletable to
     un-teach.

Pure functions only: no network, no filesystem. The caller (watch_feedback.py)
fetches pages, appends JSONL, and dedupes.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from saasnews.filters import (
    FRESHER_TITLE_NEGATIVE,
    INDIA_RE,
    PYTHON_IN_TITLE_RE,
    PYTHON_STACK_RE,
    RESTRICTED_REGION_TOKENS,
    extract_min_years,
    is_fresher_role,
    match_location,
    match_role,
)
from saasnews.feedback import domain_of

# ---------------------------------------------------------------------------
# Mid/seniority markers NOT yet covered by FRESHER_TITLE_NEGATIVE. A hit here
# on a not_fresher diagnosis becomes a "title_negative" learning.
# ---------------------------------------------------------------------------
SENIORITY_TITLE_LEXICON = [
    r"\bexperienced?\b",
    r"\bmid[\s\-]*level\b",
    r"\bmid[\s\-]*senior\b",
    r"\bsde[\s\-]*2\b",
    r"\bswe[\s\-]*2\b",
    r"\bengineer[\s\-]*2\b",
    r"\bdeveloper[\s\-]*2\b",
]

# Title tokens that make a job objectively non-target even though the generic
# role patterns accept it. A hit on a wrong_role diagnosis becomes a
# "title_exclusion" learning.
ROLE_LEAK_LEXICON = [
    r"\bqa\b",
    r"\bquality[\s\-]*assurance\b",
    r"\btest[\s\-]*engineer\b",
    r"\bsdet\b",
    r"\bsecurity\b",
    r"\bcyber[\s\-]*security\b",
    r"\bembedded\b",
    r"\bhardware\b",
    r"\belectrical\b",
    r"\bmechanical\b",
    r"\bcivil\b",
    r"\bchemical\b",
    r"\baerospace\b",
    r"\bbusiness[\s\-]*analyst\b",
    r"\btechnical[\s\-]*writer\b",
    r"\bscrum[\s\-]*master\b",
    r"\bagile[\s\-]*coach\b",
    r"\bdata[\s\-]*entry\b",
    r"\bnetwork[\s\-]*engineer\b",
    r"\bit[\s\-]*support\b",
    r"\bdesktop[\s\-]*support\b",
]

# Region candidates are extracted from JD-body sentences like
# "You must be based in Austin, Texas" or "Texas only".
_REGION_CONTEXT_RE = re.compile(
    r"(?:based\s+in|located\s+in|must\s+be\s+(?:based\s+)?in|work\s+from|"
    r"relocat\w+\s+to|candidates?\s+(?:in|from|based\s+in)|office\s+(?:is\s+)?(?:in|at)|"
    r"open\s+to\s*(?:in|:))\s*:?\s*([A-Za-z][A-Za-z .,'\-]{1,40})",
    re.IGNORECASE,
)
_REGION_ONLY_RE = re.compile(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2})\s+only\b")

# Duration demand inside a JD sentence — broader than YEARS_PATTERNS (covers
# months and prose forms the years extractor misses).
_DURATION_RE = re.compile(
    r"\b\d{1,3}\s*\+?\s*(?:years?|yrs?|months?)\b"
    r"(?:\s+of\s+(?:professional\s+|relevant\s+|hands[\s\-]*on\s+)?)?"
    r"(?:experience|exp|background|expertise|work(?:ing)?)?",
    re.IGNORECASE,
)

_LOCATION_STOPWORDS = frozenset({
    "remote", "hybrid", "onsite", "office", "worldwide", "anywhere",
    "india", "us", "usa", "uk", "eu", "united states", "united kingdom",
    "canada", "australia", "europe", "the", "our", "this", "that", "you",
    "your", "we", "they", "candidate", "candidates", "team", "company",
    "applicant", "applicants", "role", "position", "job",
})


@dataclass
class Learning:
    """One proposed generalization (kind/value per LEARNING_KINDS)."""
    kind: str
    value: str
    scope: str = "global"
    evidence: str = ""


@dataclass
class Diagnosis:
    """Verdict of replaying a feedback record against the live posting."""
    job_url: str
    constraint: str
    root_cause: str          # machine-readable code, see module docstring
    detail: str              # human-readable explanation for logs/summary
    evidence: list[str] = field(default_factory=list)
    learnings: list[Learning] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _sentences(text: str) -> list[str]:
    return [s.strip() for s in re.split(r"(?<=[.!?\n])\s+", text or "") if s.strip()]


def _snippet(text: str, limit: int = 180) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def generalize_duration_phrase(phrase: str) -> Optional[str]:
    """Turn an observed duration phrase into a conservative regex.

    '24 months of professional experience' →
        \\b\\d+\\s+months?\\s+of\\s+professional\\s+experience\\b
    Numbers become \\d+, fixed words stay literal; year/month words keep
    their optional plural so sibling postings with the same phrasing but a
    different number also match. Returns None when nothing usable remains.
    """
    plural = {"years": r"years?", "yrs": r"yrs?", "months": r"months?"}
    parts = []
    saw_number = False
    for tok in re.split(r"\s+", phrase.strip()):
        core = tok.strip(".,();:'\"")
        if not core:
            continue
        if re.fullmatch(r"\d+", core):
            saw_number = True
            parts.append(r"\d+")
        elif re.fullmatch(r"\d+\+", core):
            saw_number = True
            parts.append(r"\d+\+")
        elif re.fullmatch(r"\d+[\-–—]\d+", core):
            saw_number = True
            parts.append(r"\d+[\-–—]\d+")
        elif core.lower() in plural:
            parts.append(plural[core.lower()])
        else:
            parts.append(re.escape(core))
    if not parts or not saw_number:
        return None
    body = r"\s+".join(parts)
    return rf"\b{body}\b"


def _region_candidates(text: str) -> list[str]:
    """Extract place-name candidates from JD text ('Austin, Texas', 'Berlin')."""
    out: list[str] = []
    seen: set[str] = set()
    spans: list[str] = []
    for m in _REGION_CONTEXT_RE.finditer(text):
        spans.append(m.group(1))
    for m in _REGION_ONLY_RE.finditer(text):
        spans.append(m.group(1))
    for span in spans:
        words = [w for w in re.split(r"[,\-–—/]|\s+", span) if w]
        for n in (1, 2, 3):  # unigrams/bigrams/trigrams, e.g. "salt lake city"
            for i in range(max(0, len(words) - n + 1)):
                gram = " ".join(w.lower().strip(".,';") for w in words[i:i + n])
                gram = re.sub(r"\s+", " ", gram).strip()
                if (len(gram) < 3 or len(gram) > 30
                        or not all(c.isalpha() or c == " " for c in gram)
                        or gram in _LOCATION_STOPWORDS or gram in seen
                        or any(part in _LOCATION_STOPWORDS for part in gram.split())):
                    continue
                seen.add(gram)
                out.append(gram)
                if len(out) >= 20:
                    return out
    return out


def _region_token_covered(token: str) -> bool:
    """True when built-in tables already reject this token (nothing to learn)."""
    probe = f"{token} only"
    return bool(RESTRICTED_REGION_TOKENS.search(probe) or INDIA_RE.search(token))


def _sentence_for(text: str, needle: str) -> str:
    for s in _sentences(text):
        if needle.lower() in s.lower():
            return s
    return _snippet(text)


# ---------------------------------------------------------------------------
# Per-constraint diagnosers
# ---------------------------------------------------------------------------

def diagnose(record: dict, jd_text: str = "") -> Diagnosis:
    """Route a feedback record to its constraint's diagnoser.

    Constraints without a replayable filter dimension (stale/closed, broken
    link, duplicates, …) get an honest runtime_state verdict: the exact
    blocklist + company strikes already cover them; nothing is generalized.
    """
    constraint = str(record.get("constraint", ""))
    base = {
        "job_url": str(record.get("job_url", "")),
        "constraint": constraint,
    }
    if constraint == "not_fresher":
        return _diagnose_not_fresher(record, jd_text, **base)
    if constraint == "wrong_role":
        return _diagnose_wrong_role(record, jd_text, **base)
    if constraint == "bad_location":
        return _diagnose_bad_location(record, jd_text, **base)
    if constraint == "not_python_stack":
        return _diagnose_not_python_stack(record, jd_text, **base)
    return Diagnosis(
        **base,
        root_cause="runtime_state",
        detail=(f"'{constraint}' has no filter signal to replay; exact "
                "suppression/strikes already apply."),
    )


def _title_seniority_miss(title: str) -> Optional[tuple[str, str]]:
    """First mid/senior title marker missing from the built-in negatives."""
    if FRESHER_TITLE_NEGATIVE.search(title):
        return None  # built-ins already catch seniority here
    for pat in SENIORITY_TITLE_LEXICON:
        m = re.search(pat, title, re.IGNORECASE)
        if m:
            return pat, m.group(0)
    return None


def _diagnose_not_fresher(record: dict, jd_text: str, *,
                          job_url: str, constraint: str) -> Diagnosis:
    title = str(record.get("job_title", ""))
    basis_at_scrape = str(record.get("fresher_basis", ""))

    # Missed seniority marker in the TITLE? (e.g. "SDE 2", "Experienced …")
    miss = _title_seniority_miss(title)
    if miss:
        pat, hit = miss
        return Diagnosis(
            job_url=job_url, constraint=constraint,
            root_cause="seniority_word_missed_in_title",
            detail=(f"title marker {hit!r} signals mid/senior but is missing "
                    "from FRESHER_TITLE_NEGATIVE"),
            evidence=[_snippet(title)],
            learnings=[Learning(kind="title_negative", value=pat,
                                evidence=_snippet(title))],
        )

    now_fresher, _, _ = is_fresher_role(title, jd_text)
    if not now_fresher:
        # Replaying with today's full JD rejects → at scrape time the JD text
        # was unavailable/too short, so the balanced path accepted blindly.
        cause = ("jd_unavailable_at_scrape"
                 if basis_at_scrape in ("no_jd_ambiguous_title", "")
                 else "jd_text_insufficient_at_scrape")
        min_years = extract_min_years(jd_text)
        years_note = f"{min_years}y-min rule" if min_years is not None else "built-in seniority signals"
        ev = [_snippet(jd_text)] if jd_text else ["<no JD text available>"]
        return Diagnosis(job_url=job_url, constraint=constraint,
                         root_cause=cause,
                         detail=f"full JD now rejects via {years_note}",
                         evidence=ev)

    # Still passes: hunt for duration demands the YEARS_PATTERNS miss
    # (e.g. months-based requirements).
    for sentence in _sentences(jd_text):
        m = _DURATION_RE.search(sentence)
        if not m:
            continue
        if extract_min_years(sentence) is not None:
            continue  # existing extractor handles this phrasing
        pattern = generalize_duration_phrase(m.group(0))
        if pattern is None:
            continue
        compiled = re.compile(pattern, re.IGNORECASE)
        if not compiled.search(sentence):
            continue
        return Diagnosis(
            job_url=job_url, constraint=constraint,
            root_cause="duration_format_missed_in_jd",
            detail=f"demand phrase {m.group(0)!r} is invisible to YEARS_PATTERNS",
            evidence=[_snippet(sentence)],
            learnings=[Learning(kind="jd_negative_pattern", value=pattern,
                                evidence=_snippet(sentence))],
        )

    return Diagnosis(
        job_url=job_url, constraint=constraint,
        root_cause="no_new_signal",
        detail=("no uncovered seniority marker found; exact blocklist applies"),
        evidence=[_snippet(jd_text)] if jd_text else [],
    )


def _diagnose_wrong_role(record: dict, jd_text: str, *,
                         job_url: str, constraint: str) -> Diagnosis:
    title = str(record.get("job_title", ""))
    role = match_role(title)
    matched = f"{role[0]} via {role[1]!r} ({role[2]})" if role else "<no match now>"
    for pat in ROLE_LEAK_LEXICON:
        m = re.search(pat, title, re.IGNORECASE)
        if m:
            return Diagnosis(
                job_url=job_url, constraint=constraint,
                root_cause="role_keyword_overbroad",
                detail=(f"generic pattern accepted ({matched}); objective "
                        f"non-target marker {m.group(0)!r} in title"),
                evidence=[_snippet(title)],
                learnings=[Learning(kind="title_exclusion", value=pat,
                                    evidence=_snippet(title))],
            )
    return Diagnosis(
        job_url=job_url, constraint=constraint,
        root_cause="role_keyword_overbroad",
        detail=(f"generic pattern accepted ({matched}); no unambiguous "
                "non-target marker found — exact blocklist only"),
        evidence=[_snippet(title)],
    )


def _diagnose_bad_location(record: dict, jd_text: str, *,
                           job_url: str, constraint: str) -> Diagnosis:
    location = str(record.get("location", ""))
    title = str(record.get("job_title", ""))
    loc = match_location(location, title)
    if loc is None:
        return Diagnosis(
            job_url=job_url, constraint=constraint,
            root_cause="already_rejected_now",
            detail="current tables already reject this posting",
            evidence=[_snippet(f"{location} {title}")],
        )

    # Look for region evidence sitting in the JD BODY where the scraper only
    # sees the (clean-looking) location/title fields.
    for candidate in _region_candidates(jd_text):
        if _region_token_covered(candidate):
            continue
        sentence = _sentence_for(jd_text, candidate)
        return Diagnosis(
            job_url=job_url, constraint=constraint,
            root_cause="region_evidence_in_jd_body",
            detail=(f"accepted as {loc[0]}/{loc[1]}, but JD body mentions "
                    f"{candidate!r} which built-in region tables don't cover"),
            evidence=[_snippet(sentence)],
            learnings=[Learning(kind="location_token", value=candidate,
                                evidence=_snippet(sentence))],
        )

    cause = ("bare_remote_ambiguous"
             if loc[0] == "remote_unqualified" else "no_new_region_signal")
    return Diagnosis(
        job_url=job_url, constraint=constraint,
        root_cause=cause,
        detail=(f"accepted as {loc[0]}/{loc[1]}; no uncovered region token in "
                "JD body — exact location blocklist applies"),
        evidence=[_snippet(f"{location} | {title}")],
    )


def _diagnose_not_python_stack(record: dict, jd_text: str, *,
                               job_url: str, constraint: str) -> Diagnosis:
    title = str(record.get("job_title", ""))
    website = str(record.get("company_website", ""))
    domain = domain_of(website)

    if PYTHON_IN_TITLE_RE.search(title) or PYTHON_STACK_RE.search(jd_text):
        return Diagnosis(
            job_url=job_url, constraint=constraint,
            root_cause="posting_changed_or_refetched",
            detail="stack signal present on re-fetch — posting likely edited",
            evidence=[_snippet(title)],
        )

    # No Python anywhere now → the scrape-time acceptance came from the
    # leniency branch (short/unusable JD text was accepted without proof).
    learning = ([Learning(kind="python_strict_domain", value=domain,
                          scope="global", evidence=_snippet(title))]
                if domain else [])
    return Diagnosis(
        job_url=job_url, constraint=constraint,
        root_cause="lenient_short_jd_accepted",
        detail=("no Python-stack signal in title or JD; the <300-char leniency "
                "branch must have accepted it at scrape time"),
        evidence=[_snippet(jd_text)] if jd_text else ["<no JD text available>"],
        learnings=learning,
    )
