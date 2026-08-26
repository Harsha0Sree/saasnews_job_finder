"""Role matching, internship categorisation, fresher filter, seniority and
years-of-extraction — the pure decision logic of the scraper.

Ported 1:1 from the original tests/run_tests.py custom runner.
"""
import pytest

from saasnews_scraper import (
    FRESHER_JD_NEGATIVE,
    FRESHER_JD_POSITIVE,
    FRESHER_TITLE_NEGATIVE,
    FRESHER_TITLE_POSITIVE,
    MIN_JD_TEXT_LEN,
    PYTHON_IN_TITLE_RE,
    PYTHON_STACK_RE,
    ROLE_PATTERNS,
    TITLE_EXCLUSIONS,
    extract_min_years,
    is_fresher_role,
    match_role,
)


# ---------------------------------------------------------------------------
# Role matching (Python-stack only)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "title, category",
    [
        ("Python Backend Engineer", "Backend"),
        ("Backend Engineer (Python/Django)", "Backend"),
        ("AI Engineer", "AI/ML"),
        ("ML Engineer", "AI/ML"),
        ("Machine Learning Engineer", "AI/ML"),
        ("Data Scientist", "Data"),
        ("Data Engineer", "Data"),
        ("Full Stack Engineer", "Full-stack"),
        ("Software Engineer", "Full-stack"),
        ("Backend Engineer", "Backend"),
        ("SDE-1", "Full-stack"),
    ],
)
def test_role_match_category(title, category):
    result = match_role(title)
    assert result is not None
    assert result[0] == category


@pytest.mark.parametrize(
    "title",
    [
        "Java Backend Engineer",
        "Golang Developer",
        "Rust Developer",
        "C++ Engineer",
        "Kotlin Developer",
        # non-engineering roles
        "Product Manager",
        "Sales Engineer",
        "Marketing Manager",
        "HR Manager",
    ],
)
def test_role_rejected(title):
    assert match_role(title) is None


# ---------------------------------------------------------------------------
# Internship matching
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "title, category",
    [
        ("Software Engineer Intern", "Internship"),
        ("Summer Intern", "Internship"),
        ("ML Engineering Fellow", "Internship"),
        ("DevOps Apprentice", "Internship"),
        ("Software Engineering Co-op", "Internship"),
    ],
)
def test_internship_category(title, category):
    result = match_role(title)
    assert result is not None
    assert result[0] == category


def test_ai_engineer_intern_matches_any_category():
    assert match_role("AI Engineer Intern") is not None


# ---------------------------------------------------------------------------
# Fresher filter
# ---------------------------------------------------------------------------

def test_junior_python_developer_is_explicit_fresher():
    is_f, _, level = is_fresher_role("Junior Python Developer", "")
    assert is_f is True
    assert level == "explicit"


def test_software_engineer_intern_is_explicit_fresher():
    is_f, _, level = is_fresher_role("Software Engineer Intern", "")
    assert is_f is True
    assert level == "explicit"


def test_fresher_backend_engineer_is_explicit_fresher():
    is_f, _, level = is_fresher_role("Fresher - Backend Engineer", "")
    assert is_f is True
    assert level == "explicit"


def test_sde1_is_explicit_fresher():
    is_f, _, level = is_fresher_role("SDE-1", "")
    assert is_f is True
    assert level == "explicit"


@pytest.mark.parametrize(
    "title",
    ["Backend Engineer", "AI Engineer"],
)
def test_ambiguous_title_without_jd_accepted_as_implicit(title):
    is_f, basis, level = is_fresher_role(title, "")
    assert is_f is True
    assert level == "implicit"


LONG_FRESHER_JD = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 0-2 years of experience. Freshers welcome! "
    "You will work on backend APIs using Python, Django, and Flask. "
    "We offer mentorship and learning opportunities. Apply today!"
)
LONG_SENIOR_JD = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 5+ years of experience in Python. "
    "You will lead the backend team and oversee architecture. "
    "Deep expertise in distributed systems required. Apply today!"
)
LONG_MID_JD = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 3+ years of experience in Python. "
    "You will work on backend APIs using Django and Flask. "
    "Experience with PostgreSQL, Docker, and AWS is required. Apply!"
)


def test_jd_zero_to_two_years_accepts():
    assert is_fresher_role("Software Engineer", LONG_FRESHER_JD)[0] is True


def test_jd_five_plus_years_rejects():
    assert is_fresher_role("Software Engineer", LONG_SENIOR_JD)[0] is False


def test_jd_three_plus_years_rejects():
    assert is_fresher_role("Software Engineer", LONG_MID_JD)[0] is False


# ---------------------------------------------------------------------------
# Seniority title detection
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "title",
    [
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
    ],
)
def test_senior_titles_rejected(title):
    assert is_fresher_role(title, "")[0] is False


# ---------------------------------------------------------------------------
# Years-of-experience extraction
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "jd_text, expected",
    [
        ("Requires 5+ years of experience", 5),
        ("Requires 5 + years of experience", 5),
        ("Requires 5+ yrs of experience", 5),
        ("Requires 5 years of experience", 5),
        ("Requires 5 years experience", 5),
        ("minimum 5 years required", 5),
        ("minimum of 5 years required", 5),
        ("at least 4 years of experience", 4),
        ("3-5 years of experience", 3),
        ("3–5 years of experience", 3),  # en-dash
        ("3 to 5 years of experience", 3),
        ("experience: 6 years in Python", 6),
        ("No experience required", None),
        ("", None),
        ("Requires 5+ years. Actually 3-5 years is fine.", 3),  # min wins
        # comprehensive patterns
        ("5+ years of experience", 5),
        ("5+ yrs", 5),
        ("5 years experience", 5),
        ("5 years' experience", 5),
        ("min 5 years", 5),
        ("atleast 4 years", 4),
        ("at-least 4 years", 4),
        ("3—5 years", 3),  # em-dash
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
    ],
)
def test_extract_min_years(jd_text, expected):
    assert extract_min_years(jd_text) == expected


# ---------------------------------------------------------------------------
# JD-verified fresher roles (realistic 300+ char JDs)
# ---------------------------------------------------------------------------

FRESHER_JD = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 0-2 years of experience in Python development. "
    "Freshers are welcome to apply! You will work on backend APIs using Django and Flask. "
    "Nice to have: knowledge of PostgreSQL, Docker, and AWS. "
    "This is an entry-level position with mentorship provided."
)
SENIOR_JD = (
    "About the Role: We are looking for a Senior Software Engineer to join our team. "
    "Requirements: 5+ years of experience in Python development. "
    "You will lead a team of engineers and oversee the architecture of our platform. "
    "Deep expertise in distributed systems required. Proven track record of building "
    "scalable systems. Mentoring junior engineers is expected."
)
MID_JD = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "Requirements: 3+ years of experience in Python development. "
    "You will work on backend APIs using Django and Flask. "
    "Experience with PostgreSQL, Docker, and AWS is required. "
    "We offer competitive salary and benefits."
)
FRESHER_KW_JD = (
    "About the Role: We are looking for a Software Engineer to join our team. "
    "This is an entry-level position. Freshers are encouraged to apply! "
    "You will work on backend APIs using Python, Django and Flask. "
    "We provide mentorship and learning opportunities. "
    "No prior experience is necessary — we'll teach you everything."
)


def test_realistic_fresher_jd_accepts():
    assert is_fresher_role("Software Engineer", FRESHER_JD)[0] is True


def test_realistic_senior_jd_rejects():
    assert is_fresher_role("Software Engineer", SENIOR_JD)[0] is False


def test_realistic_mid_jd_rejects():
    assert is_fresher_role("Software Engineer", MID_JD)[0] is False


def test_fresher_keyword_jd_accepts():
    assert is_fresher_role("Software Engineer", FRESHER_KW_JD)[0] is True


# ---------------------------------------------------------------------------
# Balanced acceptance for ambiguous titles / short JS-error JDs
# ---------------------------------------------------------------------------

AMBIGUOUS_TITLES = [
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


@pytest.mark.parametrize("title", AMBIGUOUS_TITLES)
def test_ambiguous_without_jd_balanced_accept(title):
    is_f, basis, level = is_fresher_role(title, "")
    assert is_f is True
    assert level == "implicit"


SHORT_JD = "Unable to load this role. Please refresh."


@pytest.mark.parametrize("title", AMBIGUOUS_TITLES[:3])
def test_short_js_error_jd_balanced_accept(title):
    assert is_fresher_role(title, SHORT_JD)[0] is True


# ---------------------------------------------------------------------------
# Python-stack verification regexes
# ---------------------------------------------------------------------------

def test_python_in_title_positive():
    assert PYTHON_IN_TITLE_RE.search("Python Engineer")
    assert PYTHON_IN_TITLE_RE.search("Backend Engineer (Django)")


def test_python_in_title_negative():
    assert not PYTHON_IN_TITLE_RE.search("Backend Engineer")


@pytest.mark.parametrize(
    "jd",
    [
        "We use Python and FastAPI for our backend",
        "Experience with PyTorch required",
        "Experience with pandas and numpy",
        "Build LLM apps with LangChain",
    ],
)
def test_python_stack_jd_positive(jd):
    assert PYTHON_STACK_RE.search(jd)


def test_java_only_jd_not_python_stack():
    assert not PYTHON_STACK_RE.search("We use Java and Spring Boot")


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

def test_empty_title_no_role():
    assert match_role("") is None


def test_whitespace_title_no_role():
    assert match_role("   ") is None


def test_very_long_title_still_matches():
    title = ("Senior Backend Engineer Python Django Flask FastAPI "
             "PostgreSQL AWS Docker Kubernetes Microservices")
    # Seniority rejection happens in the fresher filter, not role matching.
    assert match_role(title) is not None


def test_title_with_special_chars_matches():
    assert match_role("Backend Engineer (Python/Django/PostgreSQL)") is not None


def test_contradictory_title_senior_wins():
    assert is_fresher_role("Senior Junior Engineer", "")[0] is False


def test_min_jd_text_len_constant():
    assert MIN_JD_TEXT_LEN == 200


def test_regex_collections_nonempty():
    assert ROLE_PATTERNS
    assert TITLE_EXCLUSIONS
    assert FRESHER_TITLE_POSITIVE
    assert FRESHER_TITLE_NEGATIVE
    assert FRESHER_JD_POSITIVE
    assert FRESHER_JD_NEGATIVE
