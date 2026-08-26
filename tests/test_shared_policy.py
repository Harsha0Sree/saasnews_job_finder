"""Shared-policy contract: the scraper pipeline (saasnews.filters) and the
master-file combiner (combine_results.location_passes) must agree on which
locations are acceptable. Both consume the same token tables from
saasnews.filters — this suite pins that agreement.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from saasnews.filters import match_location  # noqa: E402


@pytest.fixture(scope="module")
def location_passes():
    import combine_results
    return combine_results.location_passes


AGREED_ACCEPT = [
    "Bengaluru",
    "Bangalore, India",
    "Mumbai",
    "Remote India",
    "Remote (Worldwide)",
    "Remote (Global)",
]

AGREED_REJECT = [
    "Remote (US)",
    "Remote (UK)",
    "Remote (EU)",
    "US, Remote",
    "Remote, United States",
    "San Francisco",
    "London",
    "",
]


@pytest.mark.parametrize("loc", AGREED_ACCEPT)
def test_both_consumers_accept(loc, location_passes):
    assert match_location(loc) is not None
    assert location_passes(loc) is True


@pytest.mark.parametrize("loc", AGREED_REJECT)
def test_both_consumers_reject(loc, location_passes):
    assert match_location(loc) is None
    assert location_passes(loc) is False, (
        f"{loc!r}: combiner accepts what the filter rejects — "
        "policy tables have drifted"
    )


def test_single_source_of_truth():
    """The combiner must not carry its own copy of the shared tables."""
    import combine_results
    src = Path(combine_results.__file__).read_text()
    assert "RESTRICTED_REMOTE_REJECT = [" not in src
    assert "INDIA_TOKENS = [" not in src
