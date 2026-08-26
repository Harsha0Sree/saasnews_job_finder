"""Careers-page discovery, tested against a fake HTTP session (no network).

Seam under test: find_careers_page(session, website, deadline) — the public
contract is that the caller's wall-clock budget bounds ALL work, including the
common-path fallback probing. Overshoot beyond the deadline is bounded by at
most a couple of in-flight probes.
"""
from urllib.parse import urlparse

import saasnews_scraper
from saasnews_scraper import PER_COMPANY_BUDGET, find_careers_page

REQUEST_COST = 30.0  # fake-clock seconds consumed per HTTP request


class FakeClock:
    def __init__(self):
        self.now = 1000.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


class FakeResponse:
    def __init__(self, status_code=200, text="", url="https://dead.example.com/"):
        self.status_code = status_code
        self.text = text
        self.url = url


class SlowSiteSession:
    """Homepage answers 200 with no careers links; every other path 404s.
    Each request consumes REQUEST_COST fake-clock seconds."""

    def __init__(self, clock):
        self.clock = clock
        self.calls = []

    def get(self, url, timeout=None, allow_redirects=True, **kwargs):
        self.calls.append(url)
        self.clock.advance(REQUEST_COST)
        if urlparse(url).path in ("", "/"):
            return FakeResponse(text="<html><body>Welcome to our site</body></html>", url=url)
        return FakeResponse(status_code=404, text="", url=url)


def test_caller_deadline_bounds_common_path_probing(monkeypatch):
    clock = FakeClock()
    monkeypatch.setattr(saasnews_scraper.time, "monotonic", clock)
    session = SlowSiteSession(clock)

    start = clock()
    deadline = start + PER_COMPANY_BUDGET  # caller's budget: 90s

    careers_url, status = find_careers_page(
        session, "https://dead.example.com", deadline=deadline
    )

    assert careers_url == ""
    assert status == "not_found"
    # Homepage (1) + at most a couple of probes that started inside the budget.
    # The old bug reset the deadline and walked ALL 24 fallback paths (~25 calls).
    assert len(session.calls) <= 4, (
        f"expected at most 4 requests, got {len(session.calls)}: {session.calls} — "
        "find_careers_page ignored the caller's deadline"
    )
    # Total work is bounded: deadline plus at most two in-flight probes.
    assert clock() <= deadline + 2 * REQUEST_COST
