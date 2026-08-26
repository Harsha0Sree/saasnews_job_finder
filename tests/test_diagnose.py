"""Root-cause diagnosis engine: replay → diff → evidence-backed learnings.
Pure seams only — every case runs offline against canned JD text.
"""
import re

from saasnews.diagnose import (
    Diagnosis,
    diagnose,
    generalize_duration_phrase,
)
from saasnews.feedback import build_policy, compile_learned
from saasnews.filters import is_fresher_role, match_location, match_role


def _rec(constraint, title="", location="", basis="", website="https://acme.io",
         url="https://x/j1"):
    return {"job_url": url, "constraint": constraint, "job_title": title,
            "location": location, "fresher_basis": basis,
            "company_website": website}


_FILLER = ("We are an equal opportunity employer building web services and "
           "internal tools. You will collaborate with designers and other "
           "engineers to ship features, review code, and improve reliability. "
           "Our stack serves millions of requests and the team values clean, "
           "tested code. ")


# ---------------------------------------------------------------------------
# not_fresher
# ---------------------------------------------------------------------------

def test_not_fresher_jd_unavailable_at_scrape_now_rejects():
    jd = _FILLER + "Candidates need 5+ years of experience with distributed systems."
    d = diagnose(_rec("not_fresher", title="Platform Engineer",
                      basis="no_jd_ambiguous_title"), jd)
    assert isinstance(d, Diagnosis)
    assert d.root_cause == "jd_unavailable_at_scrape"
    assert d.learnings == []


def test_not_fresher_seniority_marker_missing_from_title_tables():
    # "SDE 2" passes built-in negatives (only SDE 3+ covered) — the diagnoser
    # proposes exactly that marker as a learned title negative.
    d = diagnose(_rec("not_fresher", title="SDE 2 - Backend Engineer"), "")
    assert d.root_cause == "seniority_word_missed_in_title"
    assert len(d.learnings) == 1
    learning = d.learnings[0]
    assert learning.kind == "title_negative"
    assert re.search(learning.value, "SDE-2 Backend Engineer", re.IGNORECASE)
    # Applying the learning flips the verdict the pipeline got wrong.
    tables = compile_learned([vars(learning)])
    ok, _, _ = is_fresher_role("SDE-2 Backend Engineer", "",
                               extra_title_negatives=tables["title_negatives"])
    assert ok is False


def test_not_fresher_duration_format_invisible_to_years_extractor():
    jd = (_FILLER + "You must have at least 24 months of professional "
          "experience before joining.")
    # Sanity: the built-in extractor really cannot see this phrasing...
    from saasnews.filters import extract_min_years
    assert extract_min_years("at least 24 months of professional experience") is None
    d = diagnose(_rec("not_fresher", title="Data Engineer", basis="jd_no_years_mentioned"), jd)
    assert d.root_cause == "duration_format_missed_in_jd"
    learning = d.learnings[0]
    assert learning.kind == "jd_negative_pattern"
    assert re.search(learning.value, "36 months of professional experience", re.IGNORECASE)
    tables = compile_learned([vars(learning)])
    ok, basis, _ = is_fresher_role("Data Engineer", jd,
                                   extra_jd_negatives=tables["jd_negatives"])
    assert ok is False and basis == "rejected_seniority_jd"


def test_not_fresher_no_new_signal_stays_exact():
    d = diagnose(_rec("not_fresher", title="Python Developer",
                      basis="title_positive"), _FILLER)
    assert d.root_cause == "no_new_signal"
    assert d.learnings == []


def test_generalize_duration_phrase_shapes():
    p = generalize_duration_phrase("4-6 yrs")
    assert p == r"\b\d+[\-–—]\d+\s+yrs?\b"
    assert re.search(p, "4-6 yrs")
    assert generalize_duration_phrase("plenty of experience") is None


# ---------------------------------------------------------------------------
# wrong_role
# ---------------------------------------------------------------------------

def test_wrong_role_qa_title_learns_exclusion():
    title = "QA Automation Engineer"
    assert match_role(title) is not None  # currently slips through
    d = diagnose(_rec("wrong_role", title=title), "")
    assert d.root_cause == "role_keyword_overbroad"
    learning = d.learnings[0]
    assert learning.kind == "title_exclusion"
    tables = compile_learned([vars(learning)])
    assert match_role(title, extra_exclusions=tables["title_exclusions"]) is None
    # ...and cannot overblock legitimate roles.
    assert match_role("Backend Python Engineer",
                      extra_exclusions=tables["title_exclusions"]) is not None


def test_wrong_role_without_clear_marker_generalizes_nothing():
    d = diagnose(_rec("wrong_role", title="Platform Engineer"), "")
    assert d.root_cause == "role_keyword_overbroad"
    assert d.learnings == []


# ---------------------------------------------------------------------------
# bad_location
# ---------------------------------------------------------------------------

def test_bad_location_region_evidence_hidden_in_jd_body():
    jd = _FILLER + "This role must be based in Austin, Texas."
    d = diagnose(_rec("bad_location", title="Backend Engineer",
                      location="Remote"), jd)
    assert d.root_cause == "region_evidence_in_jd_body"
    learning = d.learnings[0]
    assert learning.kind == "location_token"
    assert learning.value == "texas"  # "austin" already covered by built-ins
    tables = compile_learned([vars(learning)])
    assert match_location("Remote", "Backend Engineer — Texas",
                          extra_region_re=tables["region_re"]) is None
    assert match_location("Remote", "Bengaluru Backend Engineer",
                          extra_region_re=tables["region_re"]) is not None


def test_bad_location_bare_remote_without_new_signal():
    d = diagnose(_rec("bad_location", title="Backend Engineer",
                      location="Remote"), _FILLER)
    assert d.root_cause == "bare_remote_ambiguous"
    assert d.learnings == []


def test_bad_location_already_rejected_by_current_tables():
    d = diagnose(_rec("bad_location", title="ML Scientist • United States • Remote",
                      location="Remote"), "")
    assert d.root_cause == "already_rejected_now"
    assert d.learnings == []


# ---------------------------------------------------------------------------
# not_python_stack
# ---------------------------------------------------------------------------

def test_python_stack_lenient_acceptance_tightens_domain():
    d = diagnose(_rec("not_python_stack", title="Platform Engineer"), _FILLER)
    assert d.root_cause == "lenient_short_jd_accepted"
    learning = d.learnings[0]
    assert learning.kind == "python_strict_domain"
    assert learning.value == "acme.io"
    policy = build_policy([], learned_entries=[{"kind": learning.kind,
                                                "value": learning.value,
                                                "scope": learning.scope}])
    assert policy.python_strict("https://www.acme.io/jobs/1")
    assert not policy.python_strict("https://globex.io/jobs")


def test_python_stack_posting_changed_on_refetch():
    d = diagnose(_rec("not_python_stack", title="Backend Engineer"),
                 _FILLER + "We primarily use Python and Django.")
    assert d.root_cause == "posting_changed_or_refetched"
    assert d.learnings == []


# ---------------------------------------------------------------------------
# Non-replayable constraints stay honest
# ---------------------------------------------------------------------------

def test_runtime_constraints_get_runtime_state_verdict():
    for constraint in ("stale_or_closed", "broken_link", "duplicate_company",
                       "irrelevant_company", "other"):
        d = diagnose(_rec(constraint), "")
        assert d.root_cause == "runtime_state"
        assert d.learnings == []
