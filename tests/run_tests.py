#!/usr/bin/env python3
"""
Test suite for the SaaSNews scraper.
Run: python3 tests/run_tests.py
"""
import json
import os
import sys
import tempfile
import time
from pathlib import Path

# Add scripts dir to path
SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
sys.path.insert(0, SCRIPTS_DIR)

# Import the modules under test
from saasnews_scraper import (
    match_role, match_location, is_fresher_role,
    extract_min_years,
    FRESHER_TITLE_POSITIVE, FRESHER_TITLE_NEGATIVE,
    FRESHER_JD_POSITIVE, FRESHER_JD_NEGATIVE,
    PYTHON_IN_TITLE_RE, PYTHON_STACK_RE,
    _load_checkpoint, _append_checkpoint, _checkpoint_path,
    _write_xlsx_from_checkpoint,
    NewsItem, JobMatch, CompanyRecord, ErrorRecord,
    ROLE_PATTERNS, TITLE_EXCLUSIONS,
    MIN_JD_TEXT_LEN,
)


# ---------------------------------------------------------------------------
# Test framework (minimal — no external deps)
# ---------------------------------------------------------------------------

passed = 0
failed = 0
failures = []

def test(name, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        print(f"  ✓ {name}")
    else:
        failed += 1
        failures.append((name, detail))
        print(f"  ✗ {name} — {detail}")

def section(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


# ---------------------------------------------------------------------------
# Test 1: Role matching (Python-stack only)
# ---------------------------------------------------------------------------

section("TEST 1: Role matching — Python-stack only")

# Should match
test("Python Backend Engineer → Backend",
     match_role("Python Backend Engineer") is not None and
     match_role("Python Backend Engineer")[0] == "Backend")

test("Backend Engineer (Python/Django) → Backend",
     match_role("Backend Engineer (Python/Django)") is not None and
     match_role("Backend Engineer (Python/Django)")[0] == "Backend")

test("AI Engineer → AI/ML",
     match_role("AI Engineer") is not None and
     match_role("AI Engineer")[0] == "AI/ML")

test("ML Engineer → AI/ML",
     match_role("ML Engineer") is not None and
     match_role("ML Engineer")[0] == "AI/ML")

test("Machine Learning Engineer → AI/ML",
     match_role("Machine Learning Engineer") is not None and
     match_role("Machine Learning Engineer")[0] == "AI/ML")

test("Data Scientist → Data",
     match_role("Data Scientist") is not None and
     match_role("Data Scientist")[0] == "Data")

test("Data Engineer → Data",
     match_role("Data Engineer") is not None and
     match_role("Data Engineer")[0] == "Data")

test("Full Stack Engineer → Full-stack",
     match_role("Full Stack Engineer") is not None and
     match_role("Full Stack Engineer")[0] == "Full-stack")

test("Software Engineer → Full-stack",
     match_role("Software Engineer") is not None and
     match_role("Software Engineer")[0] == "Full-stack")

test("Backend Engineer → Backend",
     match_role("Backend Engineer") is not None and
     match_role("Backend Engineer")[0] == "Backend")

test("SDE-1 → Full-stack",
     match_role("SDE-1") is not None and
     match_role("SDE-1")[0] == "Full-stack")

# Should NOT match (non-Python stack)
test("Java Backend Engineer → NO MATCH",
     match_role("Java Backend Engineer") is None,
     f"Got: {match_role('Java Backend Engineer')}")

test("Golang Developer → NO MATCH",
     match_role("Golang Developer") is None,
     f"Got: {match_role('Golang Developer')}")

test("Rust Developer → NO MATCH",
     match_role("Rust Developer") is None,
     f"Got: {match_role('Rust Developer')}")

test("C++ Engineer → NO MATCH",
     match_role("C++ Engineer") is None,
     f"Got: {match_role('C++ Engineer')}")

test("Kotlin Developer → NO MATCH",
     match_role("Kotlin Developer") is None,
     f"Got: {match_role('Kotlin Developer')}")

# Should NOT match (non-engineering roles)
test("Product Manager → NO MATCH",
     match_role("Product Manager") is None)

test("Sales Engineer → NO MATCH",
     match_role("Sales Engineer") is None)

test("Marketing Manager → NO MATCH",
     match_role("Marketing Manager") is None)

test("HR Manager → NO MATCH",
     match_role("HR Manager") is None)


# ---------------------------------------------------------------------------
# Test 2: Internship matching
# ---------------------------------------------------------------------------

section("TEST 2: Internship matching")

test("Software Engineer Intern → Internship",
     match_role("Software Engineer Intern") is not None and
     match_role("Software Engineer Intern")[0] == "Internship")

test("AI Engineer Intern → Internship OR AI/ML",
     match_role("AI Engineer Intern") is not None)

test("Summer Intern → Internship",
     match_role("Summer Intern") is not None and
     match_role("Summer Intern")[0] == "Internship")

test("ML Engineering Fellow → Internship",
     match_role("ML Engineering Fellow") is not None and
     match_role("ML Engineering Fellow")[0] == "Internship")

test("DevOps Apprentice → Internship",
     match_role("DevOps Apprentice") is not None and
     match_role("DevOps Apprentice")[0] == "Internship")

test("Software Engineering Co-op → Internship",
     match_role("Software Engineering Co-op") is not None and
     match_role("Software Engineering Co-op")[0] == "Internship")


# ---------------------------------------------------------------------------
# Test 3: Fresher filter
# ---------------------------------------------------------------------------

section("TEST 3: Fresher filter (0-2 years)")

# Explicit fresher (title says Junior/Fresher/Intern)
is_f, basis, level = is_fresher_role("Junior Python Developer", "")
test("Junior Python Developer → fresher=True (explicit)",
     is_f is True and level == "explicit",
     f"Got: is_f={is_f}, level={level}")

is_f, basis, level = is_fresher_role("Software Engineer Intern", "")
test("Software Engineer Intern → fresher=True (explicit)",
     is_f is True and level == "explicit",
     f"Got: is_f={is_f}, level={level}")

is_f, basis, level = is_fresher_role("Fresher - Backend Engineer", "")
test("Fresher - Backend Engineer → fresher=True (explicit)",
     is_f is True and level == "explicit",
     f"Got: is_f={is_f}, level={level}")

is_f, basis, level = is_fresher_role("SDE-1", "")
test("SDE-1 → fresher=True (explicit)",
     is_f is True and level == "explicit",
     f"Got: is_f={is_f}, level={level}")

# Ambiguous title (no seniority keyword) WITHOUT JD → now ACCEPTED (balanced).
# This is the key reliability fix: we accept ambiguous titles when JD can't be
# fetched, to avoid missing fresher roles. Senior jobs are still rejected by
# the title NEGATIVE check (Senior/Staff/Principal/Sr./II/III).
is_f, basis, level = is_fresher_role("Backend Engineer", "")
test("Backend Engineer (no JD) → fresher=True (balanced, implicit)",
     is_f is True and level == "implicit",
     f"Got: is_f={is_f}, basis={basis}, level={level}")

is_f, basis, level = is_fresher_role("AI Engineer", "")
test("AI Engineer (no JD) → fresher=True (balanced, implicit)",
     is_f is True and level == "implicit",
     f"Got: is_f={is_f}, basis={basis}, level={level}")

# NOT a fresher (senior/staff/principal)
is_f, basis, level = is_fresher_role("Senior Software Engineer", "")
test("Senior Software Engineer → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}")

is_f, basis, level = is_fresher_role("Staff Engineer", "")
test("Staff Engineer → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}")

is_f, basis, level = is_fresher_role("Principal Engineer", "")
test("Principal Engineer → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}")

is_f, basis, level = is_fresher_role("Engineering Manager", "")
test("Engineering Manager → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}")

is_f, basis, level = is_fresher_role("Founding Engineer", "")
test("Founding Engineer → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}")

# JD-text verification (use 300+ char JDs so they pass the MIN_JD_TEXT_LEN check)
long_fresher_jd = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 0-2 years of experience. Freshers welcome! "
    "You will work on backend APIs using Python, Django, and Flask. "
    "We offer mentorship and learning opportunities. Apply today!"
)
is_f, basis, level = is_fresher_role("Software Engineer", long_fresher_jd)
test("JD says 0-2 years (300+ chars) → fresher=True",
     is_f is True,
     f"Got: is_f={is_f}, basis={basis}")

long_senior_jd = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 5+ years of experience in Python. "
    "You will lead the backend team and oversee architecture. "
    "Deep expertise in distributed systems required. Apply today!"
)
is_f, basis, level = is_fresher_role("Software Engineer", long_senior_jd)
test("JD says 5+ years (300+ chars) → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}, basis={basis}")

long_mid_jd = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 3+ years of experience in Python. "
    "You will work on backend APIs using Django and Flask. "
    "Experience with PostgreSQL, Docker, and AWS is required. Apply!"
)
is_f, basis, level = is_fresher_role("Software Engineer", long_mid_jd)
test("JD says 3+ years (300+ chars) → fresher=False",
     is_f is False,
     f"Got: is_f={is_f}, basis={basis}")


# ---------------------------------------------------------------------------
# Test 3b: Seniority title detection (the bug that caused Sr. to slip through)
# ---------------------------------------------------------------------------

section("TEST 3b: Seniority title detection (regex fixes)")

# These should ALL be detected as NEGATIVE (senior) — the bug was that "Sr."
# and Roman numerals were not being caught.
senior_titles = [
    "Sr. Forward Deployed Software Engineer",
    "Sr. Software Engineer",
    "Sr Software Engineer",
    "SR. Engineer",
    "Senior Software Engineer",
    "Staff Engineer",
    "Principal Engineer",
    "Engineering Manager",
    "Head of Engineering",
    "Director of Engineering",
    "VP of Engineering",
    "Chief Technology Officer",
    "CTO",
    "Architect",
    "Founding Engineer",
    "Software Engineer II",
    "Software Engineer III",
    "Engineer IV",
    "SDE-3",
    "SDE-4",
    "SWE-3",
    "L3 Engineer",
    "L5 Engineer",
    "Level 4 Engineer",
    "5+ years experience Engineer",
    "Lead Backend Engineer",
]
for t in senior_titles:
    is_f, basis, level = is_fresher_role(t, "")
    test(f"'{t[:45]}' → REJECTED",
         is_f is False,
         f"Got: is_f={is_f}, basis={basis}")


# ---------------------------------------------------------------------------
# Test 3c: Years-of-experience extraction from JD text
# ---------------------------------------------------------------------------

section("TEST 3c: Years-of-experience extraction")

test("'5+ years' → 5", extract_min_years("Requires 5+ years of experience") == 5)
test("'5 + years' → 5", extract_min_years("Requires 5 + years of experience") == 5)
test("'5+ yrs' → 5", extract_min_years("Requires 5+ yrs of experience") == 5)
test("'5 years of experience' → 5", extract_min_years("Requires 5 years of experience") == 5)
test("'5 years experience' → 5", extract_min_years("Requires 5 years experience") == 5)
test("'minimum 5 years' → 5", extract_min_years("minimum 5 years required") == 5)
test("'minimum of 5 years' → 5", extract_min_years("minimum of 5 years required") == 5)
test("'at least 4 years' → 4", extract_min_years("at least 4 years of experience") == 4)
test("'3-5 years' → 3", extract_min_years("3-5 years of experience") == 3)
test("'3–5 years' (en-dash) → 3", extract_min_years("3–5 years of experience") == 3)
test("'3 to 5 years' → 3", extract_min_years("3 to 5 years of experience") == 3)
test("'experience: 6 years' → 6", extract_min_years("experience: 6 years in Python") == 6)
test("'No experience required' → None", extract_min_years("No experience required") is None)
test("Empty JD → None", extract_min_years("") is None)

# When multiple patterns match, return the minimum.
test("'5+ years ... 3-5 years' → 3 (min)",
     extract_min_years("Requires 5+ years. Actually 3-5 years is fine.") == 3)


# ---------------------------------------------------------------------------
# Test 3c2: Comprehensive years extraction (production-grade)
# ---------------------------------------------------------------------------

section("TEST 3c2: Comprehensive years extraction")

# Test all the new patterns
comprehensive_year_tests = [
    ("5+ years of experience", 5),
    ("5 + years of experience", 5),
    ("5+ yrs", 5),
    ("5 years of experience", 5),
    ("5 years experience", 5),
    ("5 years' experience", 5),
    ("minimum 5 years", 5),
    ("minimum of 5 years", 5),
    ("min 5 years", 5),
    ("at least 4 years", 4),
    ("atleast 4 years", 4),
    ("at-least 4 years", 4),
    ("3-5 years", 3),
    ("3–5 years", 3),  # en-dash
    ("3—5 years", 3),  # em-dash
    ("3 to 5 years", 3),
    ("require 5 years", 5),
    ("requires 5 years", 5),
    ("required: 5 years", 5),
    ("experience: 5 years", 5),
    ("experience - 5 years", 5),
    ("5 years building", 5),
    ("5 years developing", 5),
    ("5 years in Python", 5),
    ("5 years of Python", 5),
    ("5 years working", 5),
    ("should have 5 years", 5),
    ("must have 5 years", 5),
    ("7+ years", 7),
    ("10+ years", 10),
    ("2+ years", 2),
    ("0-2 years", 0),
    ("1-3 years", 1),
]
for jd_text, expected in comprehensive_year_tests:
    result = extract_min_years(jd_text)
    test(f"'{jd_text[:40]}' → {expected}",
         result == expected,
         f"Got: {result}")


# ---------------------------------------------------------------------------
# Test 3c3: Playwright availability check
# ---------------------------------------------------------------------------

section("TEST 3c3: Playwright fallback availability")

from saasnews_scraper import PLAYWRIGHT_AVAILABLE
test("Playwright is available", PLAYWRIGHT_AVAILABLE is True,
     "Install with: pip install playwright && playwright install chromium")

if PLAYWRIGHT_AVAILABLE:
    from saasnews_scraper import render_page_text_with_playwright
    # Test that Playwright can render a simple page
    text = render_page_text_with_playwright("https://example.com", wait_ms=1000)
    test("Playwright renders example.com", len(text) > 50,
         f"Got length: {len(text)}")
    test("Playwright gets 'Example Domain' text", "Example Domain" in text,
         f"Got: {text[:100]}")


# ---------------------------------------------------------------------------
# Test 3d: Conservative rejection of ambiguous titles (the key reliability fix)
# ---------------------------------------------------------------------------

section("TEST 3d: Conservative rejection (ambiguous titles without JD)")

# These titles are ambiguous (no junior/senior keyword). Without a real JD to
# verify, they should be REJECTED (not accepted as before).
ambiguous_titles = [
    "Backend Engineer",
    "AI Engineer",
    "ML Engineer",
    "Software Engineer",
    "Platform Engineer",
    "Data Scientist",
    "Performance Engineer",
    "Strategic Deployment Engineer",
    "Machine Learning Engineer",
    "Full Stack Engineer",
]
for t in ambiguous_titles:
    is_f, basis, level = is_fresher_role(t, "")  # No JD
    test(f"'{t[:35]}' (no JD) → ACCEPTED (balanced, implicit)",
         is_f is True and level == "implicit",
         f"Got: is_f={is_f}, basis={basis}")

# With a short JD (JS-rendered error page), should also accept (balanced).
short_jd = "Unable to load this role. Please refresh."
for t in ambiguous_titles[:3]:
    is_f, basis, level = is_fresher_role(t, short_jd)
    test(f"'{t[:35]}' (short JD <200 chars) → ACCEPTED (balanced)",
         is_f is True,
         f"Got: is_f={is_f}, basis={basis}")


# ---------------------------------------------------------------------------
# Test 3e: JD-verified fresher roles (realistic JD text)
# ---------------------------------------------------------------------------

section("TEST 3e: JD-verified fresher roles")

# A realistic fresher JD (300+ chars) with "0-2 years".
fresher_jd = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 0-2 years of experience in Python development. "
    "Freshers are welcome to apply! You will work on backend APIs using Django and Flask. "
    "Nice to have: knowledge of PostgreSQL, Docker, and AWS. "
    "This is an entry-level position with mentorship provided."
)
test("Fresher JD (0-2 years, 300+ chars) → fresher=True",
     is_fresher_role("Software Engineer", fresher_jd)[0] is True,
     f"Got: {is_fresher_role('Software Engineer', fresher_jd)}")

# A realistic senior JD (300+ chars) with "5+ years".
senior_jd = (
    "About the Role: We are looking for a Senior Software Engineer to join our team. "
    "Requirements: 5+ years of experience in Python development. "
    "You will lead a team of engineers and oversee the architecture of our platform. "
    "Deep expertise in distributed systems required. Proven track record of building "
    "scalable systems. Mentoring junior engineers is expected."
)
test("Senior JD (5+ years, 300+ chars) → fresher=False",
     is_fresher_role("Software Engineer", senior_jd)[0] is False,
     f"Got: {is_fresher_role('Software Engineer', senior_jd)}")

# A mid-level JD (300+ chars) with "3+ years".
mid_jd = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 3+ years of experience in Python development. "
    "You will work on backend APIs using Django and Flask. "
    "Experience with PostgreSQL, Docker, and AWS is required. "
    "We offer competitive salary and benefits."
)
test("Mid JD (3+ years, 300+ chars) → fresher=False",
     is_fresher_role("Software Engineer", mid_jd)[0] is False,
     f"Got: {is_fresher_role('Software Engineer', mid_jd)}")

# A JD with no years signal but "fresher" keyword (300+ chars).
fresher_kw_jd = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "This is an entry-level position. Freshers are encouraged to apply! "
    "You will work on backend APIs using Python, Django and Flask. "
    "We provide mentorship and learning opportunities. "
    "No prior experience is necessary — we'll teach you everything."
)
test("Fresher keyword JD (300+ chars) → fresher=True",
     is_fresher_role("Software Engineer", fresher_kw_jd)[0] is True,
     f"Got: {is_fresher_role('Software Engineer', fresher_kw_jd)}")


# ---------------------------------------------------------------------------
# Test 4: Location matching
# ---------------------------------------------------------------------------

section("TEST 4: Location matching (India or Remote-Worldwide)")

test("Bengaluru → India",
     match_location("Bengaluru") is not None and
     match_location("Bengaluru")[0] == "india")

test("Bangalore, India → India",
     match_location("Bangalore, India") is not None and
     match_location("Bangalore, India")[0] == "india")

test("Mumbai → India",
     match_location("Mumbai") is not None and
     match_location("Mumbai")[0] == "india")

test("Remote India → India",
     match_location("Remote India") is not None and
     match_location("Remote India")[0] == "india")

test("Remote (Worldwide) → remote_worldwide",
     match_location("Remote (Worldwide)") is not None and
     match_location("Remote (Worldwide)")[0] == "remote_worldwide")

test("Remote (Global) → remote_worldwide",
     match_location("Remote (Global)") is not None and
     match_location("Remote (Global)")[0] == "remote_worldwide")

test("Remote → remote_unqualified",
     match_location("Remote") is not None and
     match_location("Remote")[0] == "remote_unqualified")

# Should NOT match (US-only / UK-only / EU-only)
test("US, Remote → NO MATCH",
     match_location("US, Remote") is None,
     f"Got: {match_location('US, Remote')}")

test("Remote, United States → NO MATCH",
     match_location("Remote, United States") is None,
     f"Got: {match_location('Remote, United States')}")

test("Remote (US) → NO MATCH",
     match_location("Remote (US)") is None,
     f"Got: {match_location('Remote (US)')}")

test("Remote (UK) → NO MATCH",
     match_location("Remote (UK)") is None,
     f"Got: {match_location('Remote (UK)')}")

test("Remote (EU) → NO MATCH",
     match_location("Remote (EU)") is None,
     f"Got: {match_location('Remote (EU)')}")

test("San Francisco → NO MATCH",
     match_location("San Francisco") is None,
     f"Got: {match_location('San Francisco')}")

test("London → NO MATCH",
     match_location("London") is None,
     f"Got: {match_location('London')}")

test("Empty location → NO MATCH",
     match_location("") is None,
     f"Got: {match_location('')}")


# ---------------------------------------------------------------------------
# Test 4b: Title + location combination filter (the bug fix)
# ---------------------------------------------------------------------------

section("TEST 4b: Title + location combination filter")

# Jobs with US/EU location info in the TITLE should be rejected even if
# the location field says "Remote".
test("Title='ML Scientist • United States •', loc='Remote' → NO MATCH",
     match_location("Remote", "ML Scientist • United States •") is None,
     f"Got: {match_location('Remote', 'ML Scientist • United States •')}")

test("Title='Engineer San Francisco', loc='Remote' → NO MATCH",
     match_location("Remote", "Engineer San Francisco") is None,
     f"Got: {match_location('Remote', 'Engineer San Francisco')}")

test("Title='Backend Engineer, London', loc='Remote' → NO MATCH",
     match_location("Remote", "Backend Engineer, London") is None,
     f"Got: {match_location('Remote', 'Backend Engineer, London')}")

test("Title='Engineer (US)', loc='Remote' → NO MATCH",
     match_location("Remote", "Engineer (US)") is None)

test("Title='Engineer Berlin', loc='Remote' → NO MATCH",
     match_location("Remote", "Engineer Berlin") is None)

test("Title='Engineer Tel Aviv', loc='Remote' → NO MATCH",
     match_location("Remote", "Engineer Tel Aviv") is None)

# Jobs with India city in title should pass.
test("Title='Backend Engineer Bengaluru', loc='' → India",
     match_location("", "Backend Engineer Bengaluru") is not None and
     match_location("", "Backend Engineer Bengaluru")[0] == "india",
     f"Got: {match_location('', 'Backend Engineer Bengaluru')}")

test("Title='ML Engineer Mumbai', loc='Remote' → India (India wins over Remote)",
     match_location("Remote", "ML Engineer Mumbai") is not None and
     match_location("Remote", "ML Engineer Mumbai")[0] == "india",
     f"Got: {match_location('Remote', 'ML Engineer Mumbai')}")

# Jobs with no restricted tokens and bare "Remote" should pass.
test("Title='Backend Engineer', loc='Remote' → remote_unqualified",
     match_location("Remote", "Backend Engineer") is not None and
     match_location("Remote", "Backend Engineer")[0] == "remote_unqualified",
     f"Got: {match_location('Remote', 'Backend Engineer')}")

test("Title='AI Engineer', loc='Remote (Worldwide)' → remote_worldwide",
     match_location("Remote (Worldwide)", "AI Engineer") is not None and
     match_location("Remote (Worldwide)", "AI Engineer")[0] == "remote_worldwide",
     f"Got: {match_location('Remote (Worldwide)', 'AI Engineer')}")


# ---------------------------------------------------------------------------
# Test 5: Checkpoint + resume
# ---------------------------------------------------------------------------

section("TEST 5: Checkpoint + resume")

with tempfile.TemporaryDirectory() as tmpdir:
    # Initially no checkpoint
    ckpt = _load_checkpoint(tmpdir)
    test("Empty checkpoint on first run", len(ckpt) == 0)

    # Write 3 checkpoint records
    for i in range(3):
        rec = {
            "news_url": f"https://example.com/news/{i}/",
            "company_rec": {
                "company_name": f"Company{i}",
                "company_website": f"https://company{i}.com",
                "news_url": f"https://example.com/news/{i}/",
                "news_headline": f"Company{i} Raises Seed",
                "funding_round": "Seed",
                "funding_date": "July 2026",
                "lead_investor": "Test VC",
                "software_category": "AI",
                "careers_page_url": f"https://company{i}.com/careers",
                "careers_page_status": "found",
                "jobs_found": 5,
                "jobs_matched": 1,
                "error": "",
            },
            "matches": [
                {
                    "company_name": f"Company{i}",
                    "company_website": f"https://company{i}.com",
                    "news_url": f"https://example.com/news/{i}/",
                    "news_headline": f"Company{i} Raises Seed",
                    "funding_round": "Seed",
                    "funding_date": "July 2026",
                    "software_category": "AI",
                    "job_title": "Python Backend Engineer",
                    "job_url": f"https://company{i}.com/jobs/1",
                    "location": "Bengaluru",
                    "apply_url": f"https://company{i}.com/jobs/1",
                    "matched_keyword": "python",
                    "match_category": "Backend",
                    "confidence": "high",
                    "location_basis": "india",
                    "careers_page_url": f"https://company{i}.com/careers",
                    "posted_date": "",
                    "fresher_basis": "title_no_seniority",
                    "fresher_level": "implicit",
                }
            ],
            "error": None,
        }
        _append_checkpoint(tmpdir, rec)

    # Verify checkpoint file has 3 lines
    ckpt_path = _checkpoint_path(tmpdir)
    with open(ckpt_path) as f:
        lines = [l for l in f if l.strip()]
    test("Checkpoint file has 3 lines", len(lines) == 3, f"Got {len(lines)} lines")

    # Reload checkpoint
    ckpt = _load_checkpoint(tmpdir)
    test("Loaded 3 records from checkpoint", len(ckpt) == 3, f"Got {len(ckpt)}")

    # Verify keys match
    expected_urls = {f"https://example.com/news/{i}/" for i in range(3)}
    actual_urls = set(ckpt.keys())
    test("Checkpoint keys match", expected_urls == actual_urls,
         f"Expected: {expected_urls}, Got: {actual_urls}")

    # Verify a record's structure
    rec0 = ckpt["https://example.com/news/0/"]
    test("Record has company_rec", "company_rec" in rec0)
    test("Record has matches list", "matches" in rec0 and len(rec0["matches"]) == 1)
    test("Record company name = Company0",
         rec0["company_rec"]["company_name"] == "Company0")

    # Simulate resume: write 1 more record, verify total = 4
    rec = {
        "news_url": "https://example.com/news/3/",
        "company_rec": {"company_name": "Company3", "company_website": "https://company3.com",
                        "news_url": "https://example.com/news/3/", "news_headline": "",
                        "funding_round": "", "funding_date": "", "lead_investor": "",
                        "software_category": "", "careers_page_url": "", "careers_page_status": "not_found",
                        "jobs_found": 0, "jobs_matched": 0, "error": ""},
        "matches": [],
        "error": None,
    }
    _append_checkpoint(tmpdir, rec)
    ckpt = _load_checkpoint(tmpdir)
    test("After adding 1 more, total = 4", len(ckpt) == 4, f"Got {len(ckpt)}")

    # Test xlsx generation from checkpoint
    out_path = _write_xlsx_from_checkpoint(tmpdir, ckpt, {"news_pages": 50, "concurrency": 10})
    test("XLSX file generated", os.path.exists(out_path), f"Path: {out_path}")

    # Verify xlsx content
    from openpyxl import load_workbook
    wb = load_workbook(out_path)
    test("XLSX has Matches sheet", "Matches" in wb.sheetnames)
    test("XLSX has All Companies sheet", "All Companies" in wb.sheetnames)
    test("XLSX has Run Summary sheet", "Run Summary" in wb.sheetnames)
    ws = wb["Matches"]
    # 3 matches (Company0,1,2 each have 1 match; Company3 has 0)
    test("XLSX Matches has 3 data rows + header", ws.max_row == 4,
         f"Got {ws.max_row} rows")
    ws = wb["All Companies"]
    test("XLSX All Companies has 4 data rows + header", ws.max_row == 5,
         f"Got {ws.max_row} rows")
    wb.close()


# ---------------------------------------------------------------------------
# Test 6: Python-stack verification
# ---------------------------------------------------------------------------

section("TEST 6: Python-stack verification")

test("Title 'Python Engineer' → has_python=True",
     bool(PYTHON_IN_TITLE_RE.search("Python Engineer")) is True)

test("Title 'Backend Engineer (Django)' → has_python=True",
     bool(PYTHON_IN_TITLE_RE.search("Backend Engineer (Django)")) is True)

test("Title 'Backend Engineer' → has_python=False",
     bool(PYTHON_IN_TITLE_RE.search("Backend Engineer")) is False)

test("JD mentions 'Python, FastAPI' → Python stack",
     bool(PYTHON_STACK_RE.search("We use Python and FastAPI for our backend")) is True)

test("JD mentions 'PyTorch' → Python stack",
     bool(PYTHON_STACK_RE.search("Experience with PyTorch required")) is True)

test("JD mentions 'Java, Spring' only → NOT Python stack",
     bool(PYTHON_STACK_RE.search("We use Java and Spring Boot")) is False)

test("JD mentions 'pandas' → Python stack",
     bool(PYTHON_STACK_RE.search("Experience with pandas and numpy")) is True)

test("JD mentions 'LangChain' → Python stack",
     bool(PYTHON_STACK_RE.search("Build LLM apps with LangChain")) is True)


# ---------------------------------------------------------------------------
# Test 7: Article parser — all 4 formats
# ---------------------------------------------------------------------------

section("TEST 7: Article parser formats")

from saasnews_scraper import (
    FD_COMPANY_SITE_RE, FD_COMPANY_RE, FD_ROUND_RE, FD_DATE_RE,
    FD_LEAD_RE, FD_CATEGORY_RE,
)

# Format 1 (newest): plain text, <br> separators
fmt1 = ("Company: Azraq<br>Round: Pre-seed<br>Funding Date: July 10, 2026"
        "<br>Lead Investor: A-typical Ventures<br>"
        "Company Website: <a href=\"https://azraq.ai/?ref=thesaasnews.com\">https://azraq.ai</a><br>"
        "Software Category: FinTech")
m = FD_COMPANY_RE.search(fmt1)
test("F1 Company:", m and m.group(1).strip() == "Azraq", f"Got: {m.group(1) if m else None}")
m = FD_ROUND_RE.search(fmt1)
test("F1 Round:", m and m.group(1).strip() == "Pre-seed", f"Got: {m.group(1) if m else None}")
m = FD_DATE_RE.search(fmt1)
test("F1 Date:", m and m.group(1).strip() == "July 10, 2026", f"Got: {m.group(1) if m else None}")
m = FD_COMPANY_SITE_RE.search(fmt1)
test("F1 Website (href):", m and m.group(2) == "https://azraq.ai/?ref=thesaasnews.com",
     f"Got: {m.groups() if m else None}")

# Format 2 (mid): plain text, &nbsp;</p><p> separators
fmt2 = ("Company: Artificially Intelligent Inc.&nbsp;&nbsp;</p>"
        "<p>Raised: $100K&nbsp;&nbsp;</p><p>Round: Angel&nbsp;&nbsp;</p>"
        "<p>Funding Month: January 2024&nbsp; &nbsp;&nbsp;&nbsp;</p>"
        "<p>Lead Investors: Payment Ventures&nbsp;&nbsp;&nbsp;</p>"
        "<p>Company Website:&nbsp;<a href=\"https://www.legalyze.ai/?ref=thesaasnews.com\">https://www.legalyze.ai/</a>&nbsp;&nbsp;&nbsp;</p>"
        "<p>Software Category: Legal&nbsp;&nbsp;&nbsp;</p>")
m = FD_COMPANY_RE.search(fmt2)
test("F2 Company:", m and m.group(1).strip() == "Artificially Intelligent Inc.",
     f"Got: {m.group(1) if m else None}")
m = FD_ROUND_RE.search(fmt2)
test("F2 Round:", m and m.group(1).strip() == "Angel", f"Got: {m.group(1) if m else None}")
m = FD_DATE_RE.search(fmt2)
test("F2 Date (Funding Month):", m and m.group(1).strip() == "January 2024",
     f"Got: {m.group(1) if m else None}")
m = FD_LEAD_RE.search(fmt2)
test("F2 Lead (Investors plural):", m and m.group(1).strip() == "Payment Ventures",
     f"Got: {m.group(1) if m else None}")

# Format 3 (older): <strong> labels, plain text URL
fmt3 = ("<strong>Company:&nbsp;</strong>ToBox Ventures Pvt. Ltd."
        "<strong>Round:&nbsp;</strong>pre-Series A"
        "<strong>Funding Month:&nbsp;</strong>December 2021"
        "<strong>Lead Investors:&nbsp;</strong>GRM Foodkraft"
        "<strong>Company Website:&nbsp;</strong>https://www.gokhana.com/"
        "<strong>Software Category:&nbsp;</strong>SaaS")
m = FD_COMPANY_RE.search(fmt3)
test("F3 Company (strong):", m and m.group(1).strip() == "ToBox Ventures Pvt. Ltd.",
     f"Got: {m.group(1) if m else None}")
m = FD_ROUND_RE.search(fmt3)
test("F3 Round (strong):", m and m.group(1).strip() == "pre-Series A",
     f"Got: {m.group(1) if m else None}")
m = FD_COMPANY_SITE_RE.search(fmt3)
test("F3 Website (plain text URL):", m is not None,
     f"Got: {m.groups() if m else None}")

# Format 3b: <strong> labels, <a> link URL
fmt3b = ("<strong>Company Website:&nbsp;</strong>"
         "<a href=\"https://example.com/?ref=thesaasnews.com\">https://example.com</a>")
m = FD_COMPANY_SITE_RE.search(fmt3b)
test("F3b Website (strong + a href):", m is not None,
     f"Got: {m.groups() if m else None}")


# ---------------------------------------------------------------------------
# Test 8: Edge cases
# ---------------------------------------------------------------------------

section("TEST 8: Edge cases")

test("Empty title → role=None",
     match_role("") is None)

test("None-like title → role=None",
     match_role("   ") is None)

test("Very long title → still matches",
     match_role("Senior Backend Engineer Python Django Flask FastAPI PostgreSQL AWS Docker Kubernetes Microservices") is not None)

test("Title with special chars → matches",
     match_role("Backend Engineer (Python/Django/PostgreSQL)") is not None)

# Contradictory title: "Senior Junior Engineer" → should be rejected (senior wins)
is_f, basis, level = is_fresher_role("Senior Junior Engineer", "")
test("'Senior Junior Engineer' → fresher=False (senior wins)",
     is_f is False,
     f"Got: is_f={is_f}")


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

section("SUMMARY")
total = passed + failed
print(f"  Passed: {passed}/{total}")
print(f"  Failed: {failed}/{total}")
if failed:
    print(f"\n  FAILURES:")
    for name, detail in failures:
        print(f"    - {name}: {detail}")
    sys.exit(1)
else:
    print(f"\n  All tests passed! ✓")
    sys.exit(0)
