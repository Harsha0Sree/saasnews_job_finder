"""Packaging contract: console entry points exist, scaffold stub is gone."""
from pathlib import Path

import pytest


def test_scraper_has_console_callable():
    import saasnews_scraper
    assert callable(getattr(saasnews_scraper, "cli", None))


def test_sync_has_console_callable():
    import sync_to_google_sheets
    assert callable(getattr(sync_to_google_sheets, "cli", None))


def test_uv_scaffold_stub_removed():
    assert not (Path(__file__).resolve().parent.parent / "main.py").exists()
