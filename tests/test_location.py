"""Location matching: India on-site OR remote-worldwide, with title+location
combination checks.

Ported 1:1 from the original tests/run_tests.py custom runner.
"""
import pytest

from saasnews_scraper import match_location


# ---------------------------------------------------------------------------
# Location matching (India or Remote-Worldwide)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "location, basis",
    [
        ("Bengaluru", "india"),
        ("Bangalore, India", "india"),
        ("Mumbai", "india"),
        ("Remote India", "india"),
        ("Remote (Worldwide)", "remote_worldwide"),
        ("Remote (Global)", "remote_worldwide"),
        ("Remote", "remote_unqualified"),
    ],
)
def test_location_accepted(location, basis):
    result = match_location(location)
    assert result is not None
    assert result[0] == basis


@pytest.mark.parametrize(
    "location",
    [
        "US, Remote",
        "Remote, United States",
        "Remote (US)",
        "Remote (UK)",
        "Remote (EU)",
        "San Francisco",
        "London",
        "",
    ],
)
def test_location_rejected(location):
    assert match_location(location) is None


# ---------------------------------------------------------------------------
# Title + location combination filter
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "location, title",
    [
        ("Remote", "ML Scientist • United States •"),
        ("Remote", "Engineer San Francisco"),
        ("Remote", "Backend Engineer, London"),
        ("Remote", "Engineer (US)"),
        ("Remote", "Engineer Berlin"),
        ("Remote", "Engineer Tel Aviv"),
    ],
)
def test_restricted_region_in_title_rejected(location, title):
    assert match_location(location, title) is None


def test_india_city_in_title_passes():
    result = match_location("", "Backend Engineer Bengaluru")
    assert result is not None
    assert result[0] == "india"


def test_india_wins_over_remote():
    result = match_location("Remote", "ML Engineer Mumbai")
    assert result is not None
    assert result[0] == "india"


def test_bare_remote_with_clean_title_passes():
    result = match_location("Remote", "Backend Engineer")
    assert result is not None
    assert result[0] == "remote_unqualified"


def test_remote_worldwide_with_clean_title_passes():
    result = match_location("Remote (Worldwide)", "AI Engineer")
    assert result is not None
    assert result[0] == "remote_worldwide"
