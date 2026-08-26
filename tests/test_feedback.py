"""Closed-loop feedback module: taxonomy, log roundtrip, policy compilation,
row harvesting, multi-select feedback cells, learned generalizations.
Pure seams only (tmp_path filesystem).
"""
import json
import re

import pytest

from saasnews.feedback import (
    CONSTRAINTS,
    LEARNING_KINDS,
    append_feedback,
    build_policy,
    compile_learned,
    dedupe_feedback,
    domain_of,
    extract_feedback_from_rows,
    learning_key,
    load_feedback,
    load_learned,
    normalize_location,
    parse_feedback_cell,
    summarize_feedback,
    summarize_learned,
)


def test_constraint_taxonomy_covers_pipeline_dimensions():
    assert len(CONSTRAINTS) == len(set(CONSTRAINTS))
    for expected in ("not_fresher", "not_python_stack", "wrong_role",
                     "bad_location", "stale_or_closed", "broken_link",
                     "duplicate_company", "irrelevant_company", "other"):
        assert expected in CONSTRAINTS


def test_append_load_roundtrip(tmp_path):
    path = str(tmp_path / "feedback.jsonl")
    rec = {"job_url": "https://x/j1", "constraint": "not_fresher"}
    append_feedback(path, rec)
    append_feedback(path, {**rec, "job_url": "https://x/j2"})
    loaded = load_feedback(path)
    assert [r["job_url"] for r in loaded] == ["https://x/j1", "https://x/j2"]


def test_load_tolerates_corrupt_lines(tmp_path):
    path = str(tmp_path / "feedback.jsonl")
    append_feedback(path, {"job_url": "https://x/j1", "constraint": "other"})
    with open(path, "a") as f:
        f.write("garbage line\n\n")
    loaded = load_feedback(path)
    assert len(loaded) == 1


def test_missing_file_loads_empty(tmp_path):
    assert load_feedback(str(tmp_path / "nope.jsonl")) == []


# ---------------------------------------------------------------------------
# Policy compilation
# ---------------------------------------------------------------------------

def test_any_constraint_blocks_exact_job_url():
    p = build_policy([{"job_url": "https://x/j1", "constraint": "stale_or_closed"}])
    assert p.job_suppressed("https://x/j1")
    assert not p.job_suppressed("https://x/j2")


def test_bad_location_learns_normalized_string():
    p = build_policy([
        {"job_url": "https://x/j1", "constraint": "bad_location",
         "location": "Remote   (US)"},
    ])
    # Normalized exact match is rejected...
    from saasnews.filters import match_location
    assert match_location("remote (us)", reject_locations=p.blocked_locations) is None
    # ...but a different location is untouched.
    assert not p.job_suppressed("https://x/j2", "Bengaluru")


def test_strikes_accumulate_per_company_and_trigger_skip():
    recs = [
        {"job_url": f"https://co.com/j{i}", "constraint": "not_fresher",
         "company_domain": "co.com"}
        for i in range(3)
    ]
    p = build_policy(recs, strike_threshold=3)
    assert p.company_strikes["co.com"] == 3
    assert p.company_suppressed("https://www.co.com/anything")

    p2 = build_policy(recs[:2], strike_threshold=3)
    assert not p2.company_suppressed("https://co.com/")


def test_strike_falls_back_to_company_website_domain():
    p = build_policy(
        [{"job_url": "https://x/j1", "constraint": "other",
          "company_website": "https://www.acme.io"}],
        strike_threshold=1,
    )
    assert p.company_suppressed("https://acme.io/careers")


@pytest.mark.parametrize(
    "url, expected",
    [
        ("https://www.acme.io/jobs", "acme.io"),
        ("http://Acme.io/", "acme.io"),
        ("", ""),
        ("not a url", ""),
    ],
)
def test_domain_of(url, expected):
    assert domain_of(url) == expected


def test_normalize_location():
    assert normalize_location("  Remote   (US) ") == "remote (us)"


# ---------------------------------------------------------------------------
# Row harvesting (xlsx / Google Sheet rows → records)
# ---------------------------------------------------------------------------

HEADERS = ["Company", "Job Title", "Job URL", "Location",
           "Company Website", "Applied", "Feedback", "Notes",
           "Fresher Basis", "Fresher Level", "Category", "Matched Keyword"]


def test_extract_harvests_valid_feedback_rows():
    rows = [
        ["Acme", "Python Dev", "https://x/j1", "Bengaluru",
         "https://acme.io", "No", "not_fresher", "asked for 5 years"],
        ["Globex", "ML Intern", "https://x/j2", "Remote (Worldwide)",
         "https://globex.io", "", "", ""],
    ]
    recs = extract_feedback_from_rows(HEADERS, rows, source="test-xlsx")
    assert len(recs) == 1
    r = recs[0]
    assert r["constraint"] == "not_fresher"
    assert r["job_url"] == "https://x/j1"
    assert r["note"] == "asked for 5 years"
    assert r["applied"] == "no"
    assert r["source"] == "test-xlsx"


def test_extract_skips_unknown_constraint_values():
    rows = [["A", "T", "https://x/j1", "L", "W", "", "salary too low", ""]]
    recs = extract_feedback_from_rows(HEADERS, rows, source="s")
    assert recs == []


def test_extract_handles_short_or_empty_rows():
    rows = [["Acme", "Dev", "https://x/j1", "Bengaluru",
             "https://acme.io", "Yes", "wrong_role"],  # short row (no notes)
            [],                                        # empty row
            ]
    recs = extract_feedback_from_rows(HEADERS, rows, source="s")
    assert len(recs) == 1
    assert recs[0]["note"] == ""


def test_extract_returns_empty_without_required_columns():
    assert extract_feedback_from_rows(["Company", "Feedback"], [], "s") == []


# ---------------------------------------------------------------------------
# Applied tracking as feedback (jobs you applied to leave the active feed)
# ---------------------------------------------------------------------------

def test_extract_harvests_applied_rows_without_feedback():
    rows = [
        ["Acme", "Python Dev", "https://x/j1", "Bengaluru",
         "https://acme.io", "Yes", "", ""],
    ]
    recs = extract_feedback_from_rows(HEADERS, rows, source="s")
    assert len(recs) == 1
    r = recs[0]
    assert r["job_url"] == "https://x/j1"
    assert r["constraint"] == ""
    assert r["applied"] == "yes"


def test_extract_enriches_with_filter_metadata():
    """Records carry WHY the pipeline let the job through — the raw material
    for analysing which signal failed."""
    rows = [
        ["Acme", "Python Dev", "https://x/j1", "Bengaluru",
         "https://acme.io", "No", "not_fresher", "asked 5 years",
         "no_jd_ambiguous_title", "implicit", "Backend", "python"],
    ]
    r = extract_feedback_from_rows(HEADERS, rows, source="s")[0]
    assert r["fresher_basis"] == "no_jd_ambiguous_title"
    assert r["fresher_level"] == "implicit"
    assert r["category"] == "Backend"
    assert r["matched_keyword"] == "python"


def test_policy_blocks_applied_urls_without_strikes():
    p = build_policy(
        [{"job_url": "https://x/j1", "constraint": "", "applied": "yes",
          "company_domain": "acme.io"}],
        strike_threshold=1,
    )
    # Applied jobs never re-enter the feed...
    assert p.job_suppressed("https://x/j1")
    assert "https://x/j1" in p.applied_job_urls
    # ...but applying is not a complaint: no company strike.
    assert p.company_strikes["acme.io"] == 0
    assert not p.company_suppressed("https://acme.io/")


def test_policy_applied_and_violation_records_coexist():
    p = build_policy([
        {"job_url": "https://x/j1", "constraint": "", "applied": "yes"},
        {"job_url": "https://x/j2", "constraint": "not_fresher",
         "company_domain": "co.com"},
    ])
    assert p.applied_job_urls == {"https://x/j1"}
    assert p.job_suppressed("https://x/j2")
    assert p.company_strikes["co.com"] == 1


# ---------------------------------------------------------------------------
# Dedupe on harvest (closed loop must not double-count)
# ---------------------------------------------------------------------------

def _rec(url, constraint):
    return {"job_url": url, "constraint": constraint}


def test_dedupe_skips_already_logged_pairs():
    existing = [_rec("https://x/j1", "not_fresher")]
    candidates = [
        _rec("https://x/j1", "not_fresher"),      # exact dup
        _rec("https://x/j1", "wrong_role"),       # same job, new constraint → kept
        _rec("https://x/j2", "not_fresher"),      # new job → kept
    ]
    result = dedupe_feedback(existing, candidates)
    assert result == [
        _rec("https://x/j1", "wrong_role"),
        _rec("https://x/j2", "not_fresher"),
    ]


def test_dedupe_collapses_duplicates_within_one_harvest():
    candidates = [
        _rec("https://x/j1", "bad_location"),
        _rec("https://x/j1", "bad_location"),
    ]
    assert dedupe_feedback([], candidates) == [_rec("https://x/j1", "bad_location")]


def test_dedupe_empty_existing_returns_all():
    cands = [_rec("https://x/j1", "other")]
    assert dedupe_feedback([], cands) == cands


# ---------------------------------------------------------------------------
# Loop analysis: which signal let the bad job through?
# ---------------------------------------------------------------------------

def test_summarize_breakdown_counts_constraint_by_basis():
    records = [
        {"constraint": "not_fresher", "fresher_basis": "no_jd_ambiguous_title"},
        {"constraint": "not_fresher", "fresher_basis": "no_jd_ambiguous_title"},
        {"constraint": "not_fresher", "fresher_basis": "jd_no_years_mentioned"},
        {"constraint": "bad_location", "fresher_basis": ""},
        {"constraint": "", "applied": "yes"},  # tracking, not a violation
    ]
    lines = summarize_feedback(records)
    assert "not_fresher × no_jd_ambiguous_title: 2" in lines
    assert "not_fresher × jd_no_years_mentioned: 1" in lines
    assert "bad_location × -: 1" in lines
    # Applied-only rows are not violations — never in the breakdown.
    assert len([l for l in lines if l.startswith(" ×")]) == 0


def test_summarize_empty_when_no_violations():
    assert summarize_feedback([]) == []
    assert summarize_feedback([{"constraint": "", "applied": "yes"}]) == []


# ---------------------------------------------------------------------------
# Multi-select Feedback cells
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "cell, expected",
    [
        ("not_fresher, bad_location", ["not_fresher", "bad_location"]),
        ("not_fresher;wrong_role", ["not_fresher", "wrong_role"]),
        ("bad_location / stale_or_closed", ["bad_location", "stale_or_closed"]),
        ("not_fresher|broken_link", ["not_fresher", "broken_link"]),
        ("not_fresher\nbad_location", ["not_fresher", "bad_location"]),
        ("not_fresher, NOT_FRESHER", ["not_fresher"]),          # dupes collapse
        ("  not_python_stack  ", ["not_python_stack"]),         # trimmed
    ],
)
def test_parse_cell_multi_select(cell, expected):
    valid, unknown = parse_feedback_cell(cell)
    assert valid == expected
    assert unknown == []


@pytest.mark.parametrize(
    "cell, expected",
    [("Senior", "not_fresher"), ("Wrong stack", "not_python_stack"),
     ("Wrong Location", "bad_location"), ("Filled", "stale_or_closed"),
     ("Dup", "duplicate_company")],
)
def test_parse_cell_aliases(cell, expected):
    valid, unknown = parse_feedback_cell(cell)
    assert valid == [expected]
    assert unknown == []


def test_parse_cell_reports_unknown_tokens_and_keeps_valid_ones():
    valid, unknown = parse_feedback_cell("not_fresher, salary too low")
    assert valid == ["not_fresher"]
    assert unknown == ["salary too low"]


def test_parse_cell_empty_is_clean():
    assert parse_feedback_cell("") == ([], [])
    assert parse_feedback_cell(None) == ([], [])


def test_extract_multi_select_row_yields_one_record_per_constraint():
    rows = [["Acme", "Python Dev", "https://x/j1", "Remote (US)",
             "https://acme.io", "", "not_fresher, bad_location",
             "both wrong", "no_jd_ambiguous_title", "implicit", "", ""]]
    recs = extract_feedback_from_rows(HEADERS, rows, source="s")
    assert [r["constraint"] for r in recs] == ["not_fresher", "bad_location"]
    for r in recs:
        assert r["job_url"] == "https://x/j1"
        assert r["fresher_basis"] == "no_jd_ambiguous_title"
        assert r["location"] == "Remote (US)"


def test_extract_mixed_valid_unknown_keeps_only_valid():
    rows = [["A", "T", "https://x/j1", "L", "W", "",
             "wrong_role, free snacks", "n", "", "", "", ""]]
    recs = extract_feedback_from_rows(HEADERS, rows, source="s")
    assert len(recs) == 1
    assert recs[0]["constraint"] == "wrong_role"


def test_extract_unknown_only_with_applied_still_tracks_job():
    rows = [["A", "T", "https://x/j1", "L", "W", "Yes", "salary too low", ""]]
    recs = extract_feedback_from_rows(HEADERS, rows, source="s")
    assert len(recs) == 1
    assert recs[0]["constraint"] == ""
    assert recs[0]["applied"] == "yes"


# ---------------------------------------------------------------------------
# Strikes stay per-JOB under multi-select (detail ≠ double punishment)
# ---------------------------------------------------------------------------

def test_multi_constraint_single_row_adds_one_strike_not_two():
    recs = extract_feedback_from_rows(
        HEADERS,
        [["Acme", "T", "https://x/j1", "L", "https://acme.io", "",
          "not_fresher, wrong_role", "", "", "", "", ""]],
        source="s")
    p = build_policy(recs, strike_threshold=2)
    assert p.company_strikes["acme.io"] == 1
    assert not p.company_suppressed("https://acme.io/x")


def test_distinct_jobs_each_contribute_a_strike():
    recs = [{"job_url": f"https://x/j{i}", "constraint": "not_fresher",
             "company_domain": "co.com"}
            for i in (1, 2)]
    p = build_policy(recs + recs[:1], strike_threshold=3)
    assert p.company_strikes["co.com"] == 2


# ---------------------------------------------------------------------------
# Learned generalizations (diagnoser-written, compile-time validated)
# ---------------------------------------------------------------------------

def _entry(kind, value, **kw):
    return {"kind": kind, "value": value, "scope": kw.get("scope", "global"),
            "evidence": kw.get("evidence", ""), "job_url": "https://x/j1"}


def test_learning_kinds_cover_all_tables():
    assert set(LEARNING_KINDS) == {
        "title_negative", "jd_negative_pattern", "title_exclusion",
        "location_token", "python_strict_domain"}


def test_load_learned_roundtrip_and_tolerates_corruption(tmp_path):
    path = str(tmp_path / "learned.jsonl")
    append_feedback(path, _entry("title_negative", r"\bsde[\s\-]*2\b"))
    with open(path, "a") as f:
        f.write("garbage\n{}\n{\"kind\": \"bogus_kind\", \"value\": \"x\"}\n")
    entries = load_learned(path)
    assert len(entries) == 1
    assert entries[0]["value"] == r"\bsde[\s\-]*2\b"


def test_compile_learned_builds_all_runtime_tables():
    tables = compile_learned([
        _entry("title_negative", r"\bsenior[\s\-]*specialist\b"),
        _entry("jd_negative_pattern", r"\b\d+\s+months?\s+experience\b"),
        _entry("title_exclusion", r"\bqa\b"),
        _entry("location_token", "texas"),
        _entry("python_strict_domain", "https://www.Acme.io/careers"),
    ])
    assert len(tables["title_negatives"]) == 1
    assert len(tables["jd_negatives"]) == 1
    assert tables["python_strict_domains"] == {"acme.io"}
    assert tables["region_re"].search("Remote (Texas)")
    assert not tables["region_re"].search("remote worldwide")


def test_compile_learned_skips_invalid_regex_loudly():
    tables = compile_learned([_entry("title_exclusion", "([unclosed")])
    assert tables["title_exclusions"] == []
    assert tables["skipped"] == 1


def test_policy_applies_learned_entries():
    p = build_policy([], learned_entries=[
        _entry("title_negative", r"\bsde[\s\-]*2\b"),
        _entry("location_token", "texas"),
        _entry("python_strict_domain", "acme.io"),
    ])
    from saasnews.filters import is_fresher_role, match_location
    # Learned title negative rejects a mid-level posting...
    assert is_fresher_role("SDE-2 Backend Engineer", "",
                           extra_title_negatives=p.learned_title_negatives)[0] is False
    # ...learned region token rejects the state built-ins miss...
    assert match_location("Remote", "Backend Engineer — Texas",
                          extra_region_re=p.learned_region_re) is None
    # ...and the strict domain loses short-JD leniency.
    assert p.python_strict("https://www.acme.io/jobs/1")
    assert not p.python_strict("https://other.io/jobs")


def test_learning_key_dedupes_same_kind_scope_value():
    a = _entry("location_token", "Texas")
    b = {"kind": "location_token", "value": "Texas", "scope": "GLOBAL"}
    c = _entry("location_token", "alberta")
    assert learning_key(a) == learning_key(b)
    assert learning_key(a) != learning_key(c)


def test_summarize_learned_counts_by_kind():
    lines = summarize_learned([
        _entry("title_negative", "a"), _entry("title_negative", "b"),
        _entry("location_token", "c")])
    assert lines[0] == "title_negative: 2"
    assert "location_token: 1" in lines
