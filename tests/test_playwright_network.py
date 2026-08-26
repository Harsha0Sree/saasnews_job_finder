"""Live-network Playwright fallback checks.

Skipped by default (see pyproject addopts). Run explicitly with:
    pytest -m network
Requires: uv sync --extra browser && playwright install chromium
"""
import pytest

from saasnews_scraper import PLAYWRIGHT_AVAILABLE

pytestmark = [
    pytest.mark.network,
    pytest.mark.skipif(not PLAYWRIGHT_AVAILABLE, reason="playwright not installed"),
]


def test_playwright_renders_example_com():
    from saasnews_scraper import render_page_text_with_playwright
    text = render_page_text_with_playwright("https://example.com", wait_ms=1000)
    assert len(text) > 50
    assert "Example Domain" in text
