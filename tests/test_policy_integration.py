"""Policy application inside the scraping pipeline.

Seams: filters.match_location's learned reject list, and process_company's
suppression of known-bad jobs (via monkeypatched network fakes — no live HTTP).
"""
import saasnews_scraper as sc
from saasnews.feedback import build_policy
from saasnews.filters import match_location


# ---------------------------------------------------------------------------
# Learned location rejection in the pure filter seam
# ---------------------------------------------------------------------------

def test_match_location_reject_list_param_blocks_exact_normalized_match():
    blocked = {"remote (us)"}
    assert match_location("Remote  (US)", reject_locations=blocked) is None


def test_match_location_reject_list_does_not_overblock():
    blocked = {"remote (us)"}
    assert match_location("Bengaluru", reject_locations=blocked) is not None
    assert match_location("Remote (Worldwide)", reject_locations=blocked) is not None


def test_match_location_without_reject_list_unchanged():
    assert match_location("Remote") == ("remote_unqualified", "medium")
    assert match_location("Bengaluru", reject_locations=set()) is not None


# ---------------------------------------------------------------------------
# process_company suppresses policy-blocked jobs
# ---------------------------------------------------------------------------

class FakeResp:
    status_code = 200
    text = "<html></html>"
    url = "https://co.example.com/"


class FakeSession:
    def get(self, url, timeout=None, allow_redirects=True, **kw):
        return FakeResp()


def _news():
    return sc.NewsItem(news_url="https://example.com/news/1/", headline="h",
                       company_name="Co", company_website="https://co.example.com")


def _patch_pipeline(monkeypatch, jobs):
    monkeypatch.setattr(sc, "find_careers_page",
                        lambda *a, **k: ("https://x/careers", "found"))
    monkeypatch.setattr(sc, "fetch_jobs_for_careers_page",
                        lambda *a, **k: (jobs, "html"))


def _job(title, url, location="Bengaluru"):
    return sc.JobPosting(title=title, url=url, location=location,
                         apply_url=url)


def test_blocked_job_url_is_suppressed(monkeypatch):
    _patch_pipeline(monkeypatch, [
        _job("Python Backend Engineer", "https://x/j1"),
        _job("Python Backend Engineer", "https://x/j2"),
    ])
    policy = build_policy([{"job_url": "https://x/j1", "constraint": "stale_or_closed"}])
    matches, rec, err = sc.process_company(_news(), FakeSession(), policy=policy)
    assert err is None
    assert [m.job_url for m in matches] == ["https://x/j2"]
    assert rec.jobs_suppressed == 1


def test_learned_bad_location_is_suppressed(monkeypatch):
    _patch_pipeline(monkeypatch, [
        _job("Python Backend Engineer", "https://x/j1", location="remote (us)"),
    ])
    policy = build_policy([{"job_url": "https://y/other", "constraint": "bad_location",
                            "location": "Remote (US)"}])
    matches, rec, err = sc.process_company(_news(), FakeSession(), policy=policy)
    assert matches == []
    assert rec.jobs_suppressed == 1


def test_no_policy_means_no_suppression_field_changes(monkeypatch):
    _patch_pipeline(monkeypatch, [_job("Python Backend Engineer", "https://x/j1")])
    matches, rec, err = sc.process_company(_news(), FakeSession())
    assert len(matches) == 1
    assert rec.jobs_suppressed == 0
