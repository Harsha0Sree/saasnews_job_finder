"""Portability contract: no machine-specific absolute paths in defaults, and
credential discovery follows a documented precedence (env → cwd → home).

Seams under test: parse_args of the scraper / sync / combine CLIs, and
find_credentials.
"""
import re
from pathlib import Path

import pytest

import combine_results
import saasnews_scraper
import sync_to_google_sheets

REPO = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------------------
# CLI defaults are portable
# ---------------------------------------------------------------------------

def test_scraper_default_output_dir_is_relative():
    args = saasnews_scraper.parse_args([])
    assert args.output_dir == "./download"


def test_sync_default_download_dir_is_relative():
    args = sync_to_google_sheets.parse_args([])
    assert args.download_dir == "./download"


def test_sync_creds_default_is_autodetect_not_hardcoded_path():
    args = sync_to_google_sheets.parse_args([])
    assert args.creds == ""  # auto-detect via find_credentials()


def test_combine_has_cli_with_portable_default():
    args = combine_results.parse_args([])
    assert args.download_dir == "./download"


# ---------------------------------------------------------------------------
# No personal absolute paths anywhere in the automation
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "relpath",
    [
        "saasnews_scraper.py",
        "sync_to_google_sheets.py",
        "combine_results.py",
        "run_daily.sh",
    ],
)
def test_no_hardcoded_home_paths(relpath):
    source = (REPO / relpath).read_text()
    assert "/home/mikeysama" not in source, (
        f"{relpath} contains a machine-specific absolute path"
    )


# ---------------------------------------------------------------------------
# Credential discovery precedence: env var → ./google-credentials.json
# ---------------------------------------------------------------------------

def test_env_var_credentials_take_priority(tmp_path, monkeypatch):
    env_file = tmp_path / "env-creds.json"
    env_file.write_text("{}")
    local = tmp_path / "google-credentials.json"
    local.write_text("{}")
    monkeypatch.setenv("GOOGLE_APPLICATION_CREDENTIALS", str(env_file))
    monkeypatch.chdir(tmp_path)
    assert sync_to_google_sheets.find_credentials() == str(env_file)


def test_local_credentials_found_in_cwd(tmp_path, monkeypatch):
    local = tmp_path / "google-credentials.json"
    local.write_text("{}")
    monkeypatch.delenv("GOOGLE_APPLICATION_CREDENTIALS", raising=False)
    monkeypatch.chdir(tmp_path)
    assert sync_to_google_sheets.find_credentials() == "google-credentials.json"


def test_missing_credentials_returns_empty(tmp_path, monkeypatch):
    monkeypatch.delenv("GOOGLE_APPLICATION_CREDENTIALS", raising=False)
    monkeypatch.chdir(tmp_path)  # empty dir, no creds anywhere
    assert sync_to_google_sheets.find_credentials() == ""
