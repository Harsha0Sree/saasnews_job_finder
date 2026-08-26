#!/usr/bin/env python3
"""
TheSaaSNews → Careers → Backend/AI/Full-stack (India or Remote-Worldwide) finder.

Pipeline:
  1. Paginate https://www.thesaasnews.com/news/page/N/  (until no rel="next")
  2. For each /news/<slug>/ article, extract the "Funding Details" block
     (Company, Website, Round, Date, Lead Investor, Software Category).
  3. For each company website, locate the careers page
     (link-text heuristics + common-path fallback).
  4. Parse the careers page for job listings. Detect hosted boards
     (Greenhouse / Lever / Ashby / Workable) and use their JSON APIs when present.
  5. Match each job against role keywords (Backend / AI-ML / Full-stack).
  6. Filter by location: India on-site OR Remote-Worldwide.
  7. Write an .xlsx workbook with three sheets: Matches | All Companies | Errors.

CLI:
  python3 saasnews_scraper.py
      [--limit N]            # only process first N discovered companies
      [--since-days N]       # only news items published in the last N days
      [--output-dir DIR]     # default: /home/z/my-project/download
      [--concurrency N]      # parallel workers, default 5
      [--news-pages N]       # max news pagination pages (safety cap), default 50
      [--verbose]
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import shutil
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Optional
from urllib.parse import urljoin, urlparse, urlunparse, parse_qs

import requests
from bs4 import BeautifulSoup

# Playwright is used as a FALLBACK for JS-rendered pages (careers pages and
# individual JD pages that return empty/error content with requests.get).
# It's optional — if not installed, the scraper falls back to requests-only.
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:  # pragma: no cover
    print("ERROR: openpyxl not installed. Run:  pip install openpyxl", file=sys.stderr)
    sys.exit(1)

# Job-evaluation policy (role/location/fresher matching) lives in the deep
# filters module; these re-exports keep the scraper's public surface stable.
from saasnews.filters import (  # noqa: E402,F401
    BARE_REMOTE,
    FRESHER_JD_NEGATIVE,
    FRESHER_JD_POSITIVE,
    FRESHER_TITLE_NEGATIVE,
    FRESHER_TITLE_POSITIVE,
    INDIA_RE,
    INDIA_TOKENS,
    MIN_JD_TEXT_LEN,
    PYTHON_IN_TITLE_RE,
    PYTHON_STACK_RE,
    REMOTE_TOKENS,
    REMOTE_WORLDWIDE_RE,
    RESTRICTED_REGION_TOKENS,
    RESTRICTED_REMOTE_RE,
    RESTRICTED_REMOTE_REJECT,
    ROLE_PATTERNS,
    TITLE_EXCLUSIONS,
    YEARS_PATTERNS,
    extract_min_years,
    is_fresher_role,
    match_location,
    match_role,
)
from saasnews.feedback import CONSTRAINTS as FEEDBACK_CONSTRAINTS  # noqa: E402,F401
from saasnews.feedback import (  # noqa: E402,F401
    Policy,
    append_feedback,
    build_policy,
    domain_of,
    extract_feedback_from_rows,
    load_feedback,
    load_learned,
    normalize_location,
    summarize_feedback,
    summarize_learned,
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
SESSION_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
TIMEOUT = 12  # seconds per HTTP request (hard)
CONNECT_TIMEOUT = 6  # seconds to establish TCP connection
READ_TIMEOUT = 12  # seconds to read response body

# Hard per-company wall-clock budget. EVERY stage (homepage, careers, JD)
# shares this budget. When exceeded, the worker abandons the company and
# moves on. This is the primary hang-prevention mechanism.
PER_COMPANY_BUDGET = 90  # seconds (increased for Playwright JD fetch fallback)
JD_FETCH_BUDGET = 60  # seconds per company for ALL JD fetches combined
JD_FETCH_MAX_JOBS = 25  # don't fetch JD for more than this many candidate jobs per company

NEWS_BASE = "https://www.thesaasnews.com/news/"

# Common careers-page paths to probe when no link is found in the homepage.
# Ordered by frequency in modern SaaS sites.
COMMON_CAREERS_PATHS = [
    "/careers", "/careers/", "/jobs", "/jobs/",
    "/about/careers", "/about-us/careers", "/about/jobs",
    "/career", "/career/",
    "/openings", "/open-roles", "/open-positions",
    "/work-with-us", "/join-us", "/join", "/join/",
    "/team", "/about/team",
    "/en/careers", "/en/jobs",
    "/life", "/life-at-", "/culture",
    "/careers/current-openings",
    "/jobs/list", "/jobs/all",
]

# Link-text heuristics for finding a careers link on the homepage.
CAREERS_LINK_TEXT_RE = re.compile(
    r"\b(careers?|jobs?|join\s+us|join\s+our\s+team|work\s+with\s+us|"
    r"open\s+roles|openings|we\s+are\s+hiring|hiring)\b",
    re.IGNORECASE,
)
CAREERS_LINK_HREF_RE = re.compile(
    r"(career|jobs?|hiring|openings|join|work-with-us)", re.IGNORECASE
)

# Hosted job-board detection patterns. When we find one of these on a careers
# page, we can hit their JSON API directly to get structured listings.
# Greenhouse supports both US (boards.greenhouse.io) and EU (boards.eu.greenhouse.io)
# instances. The embed URL pattern is `?for=<slug>` or `?board=<slug>`.
HOSTED_BOARDS = {
    # Greenhouse: boards.greenhouse.io/embed/board?board=<co>
    #             boards.eu.greenhouse.io/embed/job_board/js?for=<co>
    "greenhouse": re.compile(
        r"boards(?:\.eu)?\.greenhouse\.io/embed/(?:job_board|board)"
        r"(?:/js)?\?(?:for|board)=([a-zA-Z0-9_\-\.]+)",
        re.IGNORECASE,
    ),
    # Lever: jobs.lever.co/<co>
    "lever": re.compile(r"jobs\.lever\.co/([a-z0-9_\-]+)", re.IGNORECASE),
    # Ashby: ashbyhq.com/<co>
    "ashby": re.compile(r"ashbyhq\.com/([a-z0-9_\-]+)", re.IGNORECASE),
    # Workable: apply.workable.com/<co>
    "workable": re.compile(r"apply\.workable\.com/([a-z0-9_\-]+)", re.IGNORECASE),
}

log = logging.getLogger("saasnews")


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class NewsItem:
    """One funding-round news article from thesaasnews.com."""
    news_url: str
    headline: str
    company_name: str = ""
    company_website: str = ""
    funding_round: str = ""
    funding_date: str = ""
    lead_investor: str = ""
    software_category: str = ""
    published_at: str = ""


@dataclass
class JobMatch:
    company_name: str
    company_website: str
    news_url: str
    news_headline: str
    funding_round: str
    funding_date: str
    software_category: str
    job_title: str
    job_url: str
    location: str
    apply_url: str
    matched_keyword: str
    match_category: str
    confidence: str  # "high" | "medium" | "low"
    location_basis: str  # why we accepted the location
    careers_page_url: str
    posted_date: str = ""
    # Fresher-filter metadata (only populated when --fresher-only is used).
    fresher_basis: str = ""   # "title_positive" | "jd_positive" | "title_no_seniority" | ...
    fresher_level: str = ""   # "explicit" | "implicit" | ""
    # UTC ISO timestamp of when this match was scraped (sort key for the feed).
    scraped_at: str = ""


@dataclass
class CompanyRecord:
    """One row in the 'All Companies' sheet (every company we attempted)."""
    company_name: str
    company_website: str
    news_url: str
    news_headline: str
    funding_round: str
    funding_date: str
    lead_investor: str
    software_category: str
    careers_page_url: str = ""
    careers_page_status: str = ""  # "found" | "not_found" | "error" | "skipped_policy"
    jobs_found: int = 0
    jobs_matched: int = 0
    jobs_suppressed: int = 0  # known-bad jobs skipped via feedback policy
    error: str = ""


@dataclass
class ErrorRecord:
    company_name: str
    company_website: str
    news_url: str
    stage: str  # "news_parse" | "careers_find" | "careers_fetch" | "jobs_parse"
    error: str
    timestamp: str = ""


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)
    return s


def safe_get(session: requests.Session, url: str, retries: int = 1,
             backoff: float = 0.5, deadline: Optional[float] = None,
             **kwargs) -> Optional[requests.Response]:
    """HTTP GET with retry + exponential backoff. Returns None on failure.

    Uses separate connect/read timeouts to prevent hangs on sites that accept
    TCP but never respond. If `deadline` (monotonic) is provided, aborts early
    when the deadline is reached (used for per-company budgets).
    """
    for attempt in range(retries + 1):
        # Check deadline before each attempt.
        if deadline is not None and time.monotonic() > deadline:
            return None
        # Remaining time budget for this request's read timeout.
        remaining = None
        if deadline is not None:
            remaining = max(2.0, deadline - time.monotonic())
            read_to = min(READ_TIMEOUT, remaining)
        else:
            read_to = READ_TIMEOUT
        try:
            r = session.get(
                url,
                timeout=(CONNECT_TIMEOUT, read_to),
                allow_redirects=True,
                **kwargs,
            )
            # Retry on 429 / 5xx (transient).
            if r.status_code in (429, 500, 502, 503, 504) and attempt < retries:
                time.sleep(backoff * (attempt + 1))
                continue
            return r
        except requests.RequestException:
            if attempt < retries:
                time.sleep(backoff * (attempt + 1))
            continue
    return None


# ---------------------------------------------------------------------------
# Stage 1: discover news articles (paginate /news/page/N/)
# ---------------------------------------------------------------------------

NEWS_SLUG_RE = re.compile(r'href="(/news/[^"]+/)"', re.IGNORECASE)
REL_NEXT_RE = re.compile(r'<link[^>]+rel="next"[^>]+href="([^"]+)"', re.IGNORECASE)


def discover_news_urls(session: requests.Session, max_pages: int,
                          start_page: int = 1) -> list[str]:
    """Return a de-duplicated list of /news/<slug>/ URLs across all pages.
    Starts from page `start_page` (1-indexed) and processes up to `max_pages`
    pages total."""
    seen: dict[str, None] = {}
    # Build the starting URL based on start_page.
    if start_page <= 1:
        page_url = NEWS_BASE
    else:
        page_url = f"{NEWS_BASE}page/{start_page}/"
    pages = 0
    while page_url and pages < max_pages:
        pages += 1
        log.info("Discovering news, page %d: %s", start_page + pages - 1, page_url)
        r = safe_get(session, page_url)
        if r is None or r.status_code != 200:
            log.warning("Stopping discovery at %s (status=%s)", page_url,
                        r.status_code if r else "no-response")
            break
        html = r.text
        slugs = NEWS_SLUG_RE.findall(html)
        # Filter out pagination links like /news/page/2/
        new_on_page = 0
        for s in slugs:
            if "/page/" in s:
                continue
            full = urljoin("https://www.thesaasnews.com/", s)
            if full not in seen:
                seen[full] = None
                new_on_page += 1
        log.info("  → %d new article links (total %d)", new_on_page, len(seen))
        m = REL_NEXT_RE.search(html)
        page_url = m.group(1) if m else None
        # Be polite
        time.sleep(0.3)
    return list(seen.keys())


# ---------------------------------------------------------------------------
# Stage 2: parse a news article into a NewsItem
# ---------------------------------------------------------------------------

OG_TITLE_RE = re.compile(r'<meta\s+property="og:title"\s+content="([^"]+)"', re.IGNORECASE)
PUB_DATE_META_RE = re.compile(
    r'<meta\s+property="article:published_time"\s+content="([^"]+)"', re.IGNORECASE
)
# Funding-details field patterns (operate on raw HTML — simpler than DOM walk).
# These patterns handle FOUR different article formats across the site's history:
#
# Format 1 (newest): plain text, <br> separators
#   Company: Azraq<br>Round: Pre-seed<br>Funding Date: July 10, 2026
#
# Format 2 (mid): plain text, &nbsp;</p><p> separators
#   Company: Artificially Intelligent Inc.&nbsp;&nbsp;</p><p>Raised: $100K
#
# Format 3 (older): label in <strong>, value as plain text or <a> link
#   <strong>Company:&nbsp;</strong>ToBox Ventures Pvt. Ltd.
#   <strong>Company Website:&nbsp;</strong>https://www.gokhana.com/
#   <strong>Company Website:&nbsp;</strong><a href="...">...</a>
#
# Format 4 (oldest): no Funding Details block — article body only (no company website)
#
# The field-end lookahead matches: <br>, </p>, </strong>, &nbsp;, <strong>, $.
FD_FIELD_END = r"(?=<br>|\s*</p>|</strong>|<strong>|&nbsp;|$)"

# Company Website: handle all formats. The URL may be inside <a href="...">
# or plain text after </strong>.
FD_COMPANY_SITE_RE = re.compile(
    r'Company\s*Website:\s*(?:&nbsp;|\s|</strong>)*'
    r'(<a\s+href="([^"]+)"|https?://[^\s<"]+)',
    re.IGNORECASE,
)
FD_COMPANY_RE = re.compile(
    r'\bCompany:\s*(?:&nbsp;|\s)*(?:</strong>)?(?:&nbsp;|\s)*([^<\n&]+?)' + FD_FIELD_END,
    re.IGNORECASE,
)
FD_ROUND_RE = re.compile(
    r'\bRound:\s*(?:&nbsp;|\s)*(?:</strong>)?(?:&nbsp;|\s)*([^<\n&]+?)' + FD_FIELD_END,
    re.IGNORECASE,
)
# Date: handle both "Funding Date:" (new) and "Funding Month:" (old)
FD_DATE_RE = re.compile(
    r'\bFunding\s*(?:Date|Month):\s*(?:&nbsp;|\s)*(?:</strong>)?(?:&nbsp;|\s)*([^<\n&]+?)' + FD_FIELD_END,
    re.IGNORECASE,
)
# Lead investor: handle both "Lead Investor:" (new) and "Lead Investors:" (old)
FD_LEAD_RE = re.compile(
    r'\bLead\s*Investors?:\s*(?:&nbsp;|\s)*(?:</strong>)?(?:&nbsp;|\s)*([^<\n&]+?)' + FD_FIELD_END,
    re.IGNORECASE,
)
FD_CATEGORY_RE = re.compile(
    r'\bSoftware\s*Category:\s*(?:&nbsp;|\s)*(?:</strong>)?(?:&nbsp;|\s)*([^<\n&]+?)' + FD_FIELD_END,
    re.IGNORECASE,
)


def strip_ref(url: str) -> str:
    """Strip ?ref=thesaasnews.com etc. from a company website URL."""
    if not url:
        return ""
    p = urlparse(url)
    # Keep path & fragment, drop query (the ref param).
    return urlunparse((p.scheme, p.netloc, p.path, "", "", p.fragment))


def parse_news_article(session: requests.Session, url: str) -> Optional[NewsItem]:
    r = safe_get(session, url)
    if r is None or r.status_code != 200:
        log.warning("News fetch failed: %s (status=%s)", url,
                    r.status_code if r else "no-response")
        return None
    html = r.text

    m = OG_TITLE_RE.search(html)
    headline = m.group(1).strip() if m else ""
    m = PUB_DATE_META_RE.search(html)
    published_at = m.group(1).strip() if m else ""

    m = FD_COMPANY_SITE_RE.search(html)
    company_website = ""
    if m:
        # The URL may be in group 2 (href="...") or group 1 (the whole match
        # which is either the <a> tag or the plain text URL).
        raw_url = m.group(2) or ""
        if not raw_url:
            # Plain text URL — extract from group 1 (the whole match).
            raw_url = m.group(1) or ""
            # If it's an <a> tag, extract the href.
            if raw_url.startswith("<a"):
                href_m = re.search(r'href="([^"]+)"', raw_url)
                raw_url = href_m.group(1) if href_m else ""
        company_website = strip_ref(raw_url.strip())
    m = FD_COMPANY_RE.search(html)
    company_name = m.group(1).strip() if m else ""
    m = FD_ROUND_RE.search(html)
    funding_round = m.group(1).strip() if m else ""
    m = FD_DATE_RE.search(html)
    funding_date = m.group(1).strip() if m else ""
    m = FD_LEAD_RE.search(html)
    lead_investor = m.group(1).strip() if m else ""
    m = FD_CATEGORY_RE.search(html)
    software_category = m.group(1).strip() if m else ""

    if not company_website:
        log.debug("No company website in article: %s", url)
        return None

    return NewsItem(
        news_url=url,
        headline=headline,
        company_name=company_name or headline,
        company_website=company_website,
        funding_round=funding_round,
        funding_date=funding_date,
        lead_investor=lead_investor,
        software_category=software_category,
        published_at=published_at,
    )


# ---------------------------------------------------------------------------
# Stage 3: find the careers page for a company website
# ---------------------------------------------------------------------------

def normalize_root(url: str) -> str:
    p = urlparse(url)
    return urlunparse((p.scheme or "https", p.netloc, "/", "", "", ""))


def find_careers_page(session: requests.Session, company_website: str,
                      deadline: Optional[float] = None) -> tuple[str, str]:
    """
    Returns (careers_url, status). status ∈ {"found", "not_found"}.
    Strategy:
      1. GET the homepage. Look in <a> text/href for careers indicators.
      2. If found, follow the first match (resolved absolute URL).
      3. Else try common paths.
    """
    root = normalize_root(company_website)
    r = safe_get(session, root, deadline=deadline)
    if r is None or r.status_code != 200:
        # Try the original URL as-is (maybe it wasn't a root).
        r = safe_get(session, company_website, deadline=deadline)
        if r is None or r.status_code != 200:
            return ("", "not_found")
    base_url = r.url  # after redirects
    html = r.text
    soup = BeautifulSoup(html, "html.parser")

    # Look for <a> elements whose text or href suggests careers.
    candidates: list[tuple[int, str]] = []  # (priority, url)
    for a in soup.find_all("a", href=True):
        text = a.get_text(" ", strip=True) or ""
        href = a["href"]
        text_match = CAREERS_LINK_TEXT_RE.search(text)
        href_match = CAREERS_LINK_HREF_RE.search(href)
        if text_match and href_match:
            priority = 1
        elif text_match and ("career" in text.lower() or "job" in text.lower()
                             or "hiring" in text.lower()):
            priority = 2
        elif href_match and ("career" in href.lower() or "/jobs" in href.lower()
                             or "hiring" in href.lower()):
            priority = 3
        else:
            continue
        abs_url = urljoin(base_url, href)
        # Reject obvious socials / mailto / tel
        if abs_url.startswith(("mailto:", "tel:", "javascript:")):
            continue
        candidates.append((priority, abs_url))

    if candidates:
        candidates.sort(key=lambda x: x[0])
        return (candidates[0][1], "found")

    # Fallback: probe common paths. Honor the caller's per-company wall-clock
    # budget when provided; only fall back to a fresh budget if the caller
    # passed none. Never extend a budget that is already running out.
    if deadline is None:
        deadline = time.monotonic() + PER_COMPANY_BUDGET
    for path in COMMON_CAREERS_PATHS:
        # Don't start a probe we can't finish meaningfully (safe_get floors
        # its read timeout at 2s — require at least that much budget left).
        if time.monotonic() + 2.0 > deadline:
            break
        probe = urljoin(base_url, path)
        r2 = safe_get(session, probe, deadline=deadline)
        if r2 is None:
            continue
        if r2.status_code == 200:
            # Sanity check: must mention jobs/careers somewhere in body OR
            # contain a hosted job-board embed (Greenhouse/Lever/Ashby/Workable).
            body = r2.text[:50000].lower()
            has_careers_kw = ("career" in body or "job" in body or
                              "opening" in body or "position" in body or
                              "hiring" in body or "join" in body)
            has_job_board = ("greenhouse" in body or "lever.co" in body or
                             "ashbyhq" in body or "workable" in body or
                             "jobvite" in body or "smartrecruiters" in body or
                             "teamtailor" in body or "personio" in body or
                             "recruitee" in body or "greenhouse.io" in body)
            # Reject 404 pages that return 200 (common on SPA sites).
            is_404_page = "404" in body[:500] and "not found" in body[:500]
            if (has_careers_kw or has_job_board) and not is_404_page:
                return (r2.url, "found")
    return ("", "not_found")


# ---------------------------------------------------------------------------
# Stage 4: parse a careers page into a list of jobs
# ---------------------------------------------------------------------------

@dataclass
class JobPosting:
    title: str
    url: str
    location: str
    apply_url: str = ""
    posted_date: str = ""


def _detect_hosted_board(html: str) -> Optional[tuple[str, str]]:
    """Return (board_type, board_company_slug) if a hosted board is detected."""
    for board, pattern in HOSTED_BOARDS.items():
        m = pattern.search(html)
        if m:
            slug = m.group(1)
            # Filter out obviously-wrong captures
            if slug and len(slug) < 80 and "/" not in slug:
                return (board, slug)
    return None


def fetch_greenhouse_jobs(session: requests.Session, slug: str) -> list[JobPosting]:
    """Greenhouse public board JSON: https://boards-api.greenhouse.io/v1/boards/<slug>/jobs"""
    url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
    r = safe_get(session, url, headers={"Accept": "application/json"})
    if r is None or r.status_code != 200:
        return []
    try:
        data = r.json()
    except Exception:
        return []
    jobs: list[JobPosting] = []
    for j in data.get("jobs", []):
        title = j.get("title", "").strip()
        if not title:
            continue
        loc = j.get("location", {}).get("name", "") if isinstance(j.get("location"), dict) \
            else str(j.get("location", "") or "")
        apply_url = j.get("absolute_url") or j.get("url") or ""
        jobs.append(JobPosting(
            title=title,
            url=apply_url,
            location=loc,
            apply_url=apply_url,
            posted_date=j.get("updated_at", ""),
        ))
    return jobs


def fetch_lever_jobs(session: requests.Session, slug: str) -> list[JobPosting]:
    """Lever public board JSON: https://api.lever.co/v0/postings/<slug>?mode=json"""
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    r = safe_get(session, url, headers={"Accept": "application/json"})
    if r is None or r.status_code != 200:
        return []
    try:
        data = r.json()
    except Exception:
        return []
    jobs: list[JobPosting] = []
    for j in data:
        title = j.get("text", "").strip()
        if not title:
            continue
        loc_parts = []
        for k in ("categories", "location"):
            v = j.get(k)
            if isinstance(v, dict):
                loc_parts.append(str(v.get("location") or v.get("name") or ""))
            elif isinstance(v, str):
                loc_parts.append(v)
        loc = ", ".join(p for p in loc_parts if p)
        apply_url = j.get("applyUrl") or j.get("hostedUrl") or ""
        jobs.append(JobPosting(
            title=title,
            url=j.get("hostedUrl") or apply_url,
            location=loc,
            apply_url=apply_url,
            posted_date=j.get("createdAt", ""),
        ))
    return jobs


def fetch_ashby_jobs(session: requests.Session, slug: str) -> list[JobPosting]:
    """Ashby public board JSON POST: https://api.ashbyhq.com/posting-api/job-board?organizationName=<slug>"""
    url = "https://api.ashbyhq.com/posting-api/job-board"
    r = safe_get(session, url, params={"organizationName": slug},
                 headers={"Accept": "application/json"})
    if r is None or r.status_code != 200:
        return []
    try:
        data = r.json()
    except Exception:
        return []
    jobs: list[JobPosting] = []
    for j in data.get("jobs", []):
        title = j.get("title", "").strip()
        if not title:
            continue
        loc = j.get("locationName", "") or ""
        if j.get("locationType"):
            loc = (loc + " (" + j["locationType"] + ")").strip()
        apply_url = j.get("publicUrl") or ""
        jobs.append(JobPosting(
            title=title,
            url=apply_url,
            location=loc,
            apply_url=apply_url,
            posted_date=j.get("publishedDate", ""),
        ))
    return jobs


def fetch_workable_jobs(session: requests.Session, slug: str) -> list[JobPosting]:
    """Workable public board JSON: https://apply.workable.com/api/v1/accounts/<slug>/jobs"""
    url = f"https://apply.workable.com/api/v1/accounts/{slug}/jobs"
    r = safe_get(session, url, headers={"Accept": "application/json"})
    if r is None or r.status_code != 200:
        return []
    try:
        data = r.json()
    except Exception:
        return []
    jobs: list[JobPosting] = []
    for j in data.get("jobs", []) or data.get("results", []) or []:
        title = j.get("title", "").strip()
        if not title:
            continue
        loc = j.get("location", "") or ""
        if j.get("country"):
            loc = (loc + ", " + j["country"]).strip(", ")
        apply_url = j.get("url") or j.get("apply_url") or ""
        if apply_url and not apply_url.startswith("http"):
            apply_url = urljoin("https://apply.workable.com/", apply_url)
        jobs.append(JobPosting(
            title=title,
            url=apply_url,
            location=loc,
            apply_url=apply_url,
            posted_date=j.get("published_on", ""),
        ))
    return jobs


HOSTED_FETCHERS = {
    "greenhouse": fetch_greenhouse_jobs,
    "lever": fetch_lever_jobs,
    "ashby": fetch_ashby_jobs,
    "workable": fetch_workable_jobs,
}


def parse_generic_careers_html(html: str, base_url: str) -> list[JobPosting]:
    """
    Fallback: parse any HTML page for job postings.
    Heuristics (best-effort):
      - Look for repeated structures with a job title + location text.
      - Common patterns: <a> with text + sibling location text;
                         <li> / <div> cards with title + location;
                         JSON-LD JobPosting blocks.
    """
    jobs: list[JobPosting] = []
    soup = BeautifulSoup(html, "html.parser")

    # 1. JSON-LD JobPosting blocks (most reliable).
    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.string or ""
        if "JobPosting" not in raw:
            continue
        try:
            # JSON-LD may be a list or single object
            data = json.loads(raw)
        except Exception:
            continue
        items = data if isinstance(data, list) else [data]
        for it in items:
            if not isinstance(it, dict):
                continue
            if it.get("@type") not in ("JobPosting", ["JobPosting"]):
                continue
            title = it.get("title", "").strip()
            if not title:
                continue
            loc = ""
            jl = it.get("jobLocation") or it.get("applicantLocationRequirements")

            def _extract_addr(place: dict) -> str:
                """Extract a readable location string from a Place/PostalAddress dict."""
                if not isinstance(place, dict):
                    return str(place) if place else ""
                addr = place.get("address", place)
                if isinstance(addr, dict):
                    parts = [
                        addr.get("addressLocality") or "",
                        addr.get("addressRegion") or "",
                        addr.get("addressCountry") or addr.get("country") or "",
                    ]
                    return ", ".join(p for p in parts if p)
                return str(addr) if addr else ""

            if isinstance(jl, dict):
                loc = _extract_addr(jl)
            elif isinstance(jl, list):
                # List of Place objects — join the first 2 non-empty.
                parts = [_extract_addr(p) for p in jl if p]
                loc = " | ".join(p for p in parts if p)
            apply_url = it.get("url") or ""
            jobs.append(JobPosting(
                title=title,
                url=urljoin(base_url, apply_url) if apply_url else base_url,
                location=loc,
                apply_url=urljoin(base_url, apply_url) if apply_url else "",
                posted_date=it.get("datePosted", ""),
            ))

    if jobs:
        return jobs

    # 2. Heuristic: find <a> elements whose text looks like a job title.
    #    We then look at the parent container for a location string.
    seen_urls: set[str] = set()
    # Reject generic nav/footer links.
    NAV_TEXT_RE = re.compile(
        r"^(home|about|blog|news|contact|login|sign|privacy|terms|menu|"
        r"learn more|read more|back|next|previous|search|subscribe)$",
        re.IGNORECASE,
    )
    LOC_HINT_RE = re.compile(
        r"(remote|hybrid|on[- ]?site|india|bangalore|bengaluru|mumbai|delhi|"
        r"hyderabad|pune|chennai|gurgaon|noida|london|new york|san francisco|"
        r"berlin|tokyo|singapore|[A-Z][a-z]+,\s*[A-Z]{2})",
        re.IGNORECASE,
    )
    for a in soup.find_all("a", href=True):
        text = a.get_text(" ", strip=True)
        if not text or len(text) < 4 or len(text) > 200:
            continue
        if NAV_TEXT_RE.match(text):
            continue
        # Heuristic: title-case-ish and ends with "Engineer" / "Developer" / etc.
        if not re.search(
            r"(engineer|developer|scientist|architect|lead|manager|"
            r"specialist|analyst|programmer|consultant)\b",
            text, re.IGNORECASE,
        ):
            continue
        href = a["href"]
        abs_url = urljoin(base_url, href)
        if abs_url in seen_urls:
            continue
        # Skip anchor-only links on the same page (likely widgets).
        if abs_url.rstrip("#") == base_url.rstrip("#"):
            continue
        seen_urls.add(abs_url)
        # Clean the title: strip common CTA / metadata suffixes that often
        # appear inside the same <a> element (e.g., "Engineer San Francisco /
        # Remote Full Time Apply Now").
        title_clean = re.sub(
            r"\s+(Full[\s\-]*Time|Part[\s\-]*Time|Contract|Permanent|Temporary|"
            r"Internship|Entry[\s\-]*Level|Senior|Junior|Mid[\s\-]*Level|"
            r"Apply(\s+Now)?|Learn\s+More|Read\s+More|View\s+Job|Apply\s+→?)\b.*$",
            "", text, flags=re.IGNORECASE,
        ).strip()
        # Also strip a trailing location pattern: "<Title> <City> / Remote"
        # or "<Title> - <City>, <Country>".
        title_clean = re.sub(
            r"\s*[/·•\-|]\s*(?:[A-Z][a-zA-Z. ]+\s*,\s*)?(?:Remote|Hybrid|On[\s\-]?site).*$",
            "", title_clean,
        ).strip()
        title_clean = re.sub(
            r"\s*[/·•\-|]\s*[A-Z][a-zA-Z. ]+,\s*[A-Z]{2}.*$",
            "", title_clean,
        ).strip()
        if title_clean and len(title_clean) >= 4:
            text = title_clean
        # Try to find a location hint near this link. Walk up the DOM and look
        # at sibling elements (location is usually a sibling <span>/<div>/<p>
        # next to the title link, not buried in the parent's full text).
        loc = ""
        parent = a.parent
        for _ in range(4):
            if parent is None:
                break
            # If this parent contains MULTIPLE <a> tags, it's a list container
            # (e.g., <ul> of all jobs) — its full text would mix locations
            # across jobs. Only use direct children that contain NO nested <a>.
            multi_link = len(parent.find_all("a", href=True)) > 1
            candidates: list[str] = []
            for sib in parent.find_all(["span", "div", "p", "li", "td"], recursive=False):
                # Skip children that contain their own links (would pollute).
                if sib.find("a", href=True):
                    continue
                t = sib.get_text(" ", strip=True)
                if t:
                    candidates.append(t)
            # Only include the parent's own direct text if it's a single-job
            # container (no sibling <a> tags).
            if not multi_link:
                candidates.append(parent.get_text(" ", strip=True))
            for txt in candidates:
                if not txt:
                    continue
                m = LOC_HINT_RE.search(txt)
                if m:
                    # Extract a clean location window around the match.
                    start = max(0, m.start() - 5)
                    end = min(len(txt), m.end() + 40)
                    raw = txt[start:end].strip(" ·,|•\u2022-—")
                    # Cut at common separators to drop preceding garbage.
                    raw = re.split(r"[·•|]+", raw)[-1].strip()
                    if raw and len(raw) < 80:
                        loc = raw
                        break
            if loc:
                break
            parent = parent.parent
        jobs.append(JobPosting(
            title=text,
            url=abs_url,
            location=loc,
            apply_url=abs_url,
        ))
    return jobs


def fetch_jobs_for_careers_page(
    session: requests.Session, careers_url: str,
    deadline: Optional[float] = None,
) -> tuple[list[JobPosting], str]:
    """
    Fetch and parse the careers page. Returns (jobs, source) where source is
    "greenhouse" | "lever" | "ashby" | "workable" | "html" | "playwright" | "error".

    Strategy:
      1. Try requests.get (fast).
      2. Check for hosted board embeds (Greenhouse/Lever/Ashby/Workable) — use their JSON APIs.
      3. Parse HTML for jobs.
      4. If HTML parse finds 0 jobs AND the page looks JS-rendered, fall back to Playwright.
    """
    r = safe_get(session, careers_url, deadline=deadline)
    if r is None or r.status_code != 200:
        # Try Playwright directly if requests failed.
        if deadline and time.monotonic() > deadline:
            return ([], "error")
        pw_html = render_page_with_playwright(careers_url)
        if pw_html:
            jobs = parse_generic_careers_html(pw_html, careers_url)
            if jobs:
                return (jobs, "playwright")
        return ([], "error")
    html = r.text
    base_url = r.url

    # 1. Hosted board?
    board = _detect_hosted_board(html)
    if board:
        board_type, slug = board
        fetcher = HOSTED_FETCHERS[board_type]
        jobs = fetcher(session, slug)
        if jobs:
            return (jobs, board_type)
        # If the hosted API returns nothing, fall through to HTML parse.

    # 2. Generic HTML parse.
    jobs = parse_generic_careers_html(html, base_url)
    if jobs:
        return (jobs, "html")

    # 3. If HTML parse found 0 jobs, the page may be JS-rendered.
    #    Fall back to Playwright to render the page and re-parse.
    if deadline and time.monotonic() > deadline:
        return ([], "empty")
    pw_html = render_page_with_playwright(careers_url)
    if pw_html:
        jobs = parse_generic_careers_html(pw_html, careers_url)
        if jobs:
            return (jobs, "playwright")

    return ([], "empty")


# ---------------------------------------------------------------------------
# Fresher / 0-experience checker
# ---------------------------------------------------------------------------

# --- Playwright headless browser fallback for JS-rendered pages ---
# Many modern SaaS careers pages are SPAs (React/Next.js/Vue) that load job
# content via JavaScript. requests.get only gets the empty HTML shell.
# Playwright actually executes JS and returns the rendered DOM.

# Thread-local Playwright instances (one per worker thread) for efficiency.
import threading
_pw_storage = threading.local()
_pw_lock = threading.Lock()  # Serialize Playwright access (sync API isn't thread-safe)

def _get_playwright_page():
    """Get the shared Playwright browser page (lazy-initialized).
    Returns (playwright_ctx, browser, page) or (None, None, None) if unavailable.
    NOTE: Playwright's sync API is NOT thread-safe. All callers MUST hold _pw_lock.
    """
    if not PLAYWRIGHT_AVAILABLE:
        return (None, None, None)
    if not hasattr(_pw_storage, 'pw'):
        try:
            pw = sync_playwright().start()
            browser = pw.chromium.launch(
                headless=True,
                args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
            )
            context = browser.new_context(
                user_agent=USER_AGENT,
                viewport={'width': 1280, 'height': 800},
            )
            page = context.new_page()
            _pw_storage.pw = pw
            _pw_storage.browser = browser
            _pw_storage.context = context
            _pw_storage.page = page
        except Exception as e:
            log.debug("Playwright init failed: %s", e)
            _pw_storage.pw = None
            return (None, None, None)
    return (_pw_storage.pw, _pw_storage.browser, _pw_storage.page)


def render_page_with_playwright(url: str, wait_ms: int = 3000,
                                 timeout_ms: int = 15000) -> Optional[str]:
    """Render a JS-heavy page with Playwright and return the full HTML.
    Returns None if Playwright is unavailable or the page fails to load.
    Thread-safe: acquires a lock because Playwright sync API isn't thread-safe.
    """
    if not PLAYWRIGHT_AVAILABLE:
        return None
    with _pw_lock:
        pw, browser, page = _get_playwright_page()
        if page is None:
            return None
        try:
            page.goto(url, wait_until='networkidle', timeout=timeout_ms)
            page.wait_for_timeout(wait_ms)
            html = page.content()
            return html
        except Exception as e:
            log.debug("Playwright render failed for %s: %s", url, e)
            return None


def render_page_text_with_playwright(url: str, wait_ms: int = 3000,
                                      timeout_ms: int = 15000) -> str:
    """Render a JS-heavy page with Playwright and return visible text.
    Returns "" if unavailable or failed.
    Thread-safe: acquires a lock because Playwright sync API isn't thread-safe.
    """
    if not PLAYWRIGHT_AVAILABLE:
        return ""
    with _pw_lock:
        pw, browser, page = _get_playwright_page()
        if page is None:
            return ""
        try:
            page.goto(url, wait_until='networkidle', timeout=timeout_ms)
            page.wait_for_timeout(wait_ms)
            # Get visible text (excludes script/style automatically).
            text = page.inner_text('body')
            return text or ""
        except Exception as e:
            log.debug("Playwright text render failed for %s: %s", url, e)
            return ""


def fetch_jd_text(session: requests.Session, job_url: str,
                  max_chars: int = 10000,
                  deadline: Optional[float] = None) -> str:
    """Fetch the job-description page text (truncated).

    Strategy:
      1. Try requests.get (fast, works for server-rendered pages + Greenhouse/Lever JSON APIs).
      2. If the result is too short (< MIN_JD_TEXT_LEN, likely JS-rendered),
         fall back to Playwright headless browser (renders JS).
      3. If Playwright also fails or is unavailable, return whatever we have.

    Returns "" on complete failure. The caller falls back to title-only
    fresher detection in that case.
    """
    if not job_url or not job_url.startswith(("http://", "https://")):
        return ""

    # Step 1: Try requests.get first (fast path).
    best_text = ""
    for attempt in range(2):
        if deadline is not None and time.monotonic() > deadline:
            break
        r = safe_get(session, job_url, retries=1, deadline=deadline)
        if r is None or r.status_code != 200:
            continue
        try:
            soup = BeautifulSoup(r.text, "html.parser")
            for tag in soup(["script", "style", "noscript", "svg", "header",
                             "footer", "nav", "aside"]):
                tag.decompose()
            text = soup.get_text(" ", strip=True)
            if len(text) > len(best_text):
                best_text = text
            # If we got substantial content, no need to retry.
            if len(text) >= MIN_JD_TEXT_LEN:
                break
            time.sleep(0.3)
        except Exception:
            continue

    # If requests.get got enough content, return it.
    if len(best_text) >= MIN_JD_TEXT_LEN:
        return best_text[:max_chars]

    # Step 2: Fall back to Playwright for JS-rendered pages.
    if deadline is not None and time.monotonic() > deadline:
        return best_text[:max_chars]
    pw_text = render_page_text_with_playwright(job_url, wait_ms=3000)
    if len(pw_text) > len(best_text):
        return pw_text[:max_chars]

    return best_text[:max_chars]


# ---------------------------------------------------------------------------
# Stage 6: per-company processing
# ---------------------------------------------------------------------------

def process_company(
    news: NewsItem,
    session: requests.Session,
    fresher_only: bool = False,
    python_stack: bool = True,
    policy: Optional[Policy] = None,
) -> tuple[list[JobMatch], CompanyRecord, Optional[ErrorRecord]]:
    """Process one company: find careers, fetch jobs, filter matches.
    If fresher_only=True, additionally filter to fresher-eligible roles
    (0-2 years experience), verifying via JD-text fetch.
    If python_stack=True (default), require the title OR JD to mention a
    Python-stack technology (Python/Django/Flask/FastAPI/PyTorch/etc.).

    `policy` (learned from user feedback, saasnews.feedback) suppresses
    known-bad job URLs and flagged location strings before they reach the
    feed; suppressed counts land on the company record and Run Summary.

    A hard wall-clock deadline (PER_COMPANY_BUDGET) is enforced via the
    `deadline` param threaded into safe_get. This is the primary hang-prevention
    mechanism: no single company can block a worker for more than ~45s.
    """
    deadline = time.monotonic() + PER_COMPANY_BUDGET
    rec = CompanyRecord(
        company_name=news.company_name,
        company_website=news.company_website,
        news_url=news.news_url,
        news_headline=news.headline,
        funding_round=news.funding_round,
        funding_date=news.funding_date,
        lead_investor=news.lead_investor,
        software_category=news.software_category,
    )
    err: Optional[ErrorRecord] = None

    # 1. Find careers page.
    try:
        careers_url, status = find_careers_page(session, news.company_website,
                                                deadline=deadline)
    except Exception as e:
        rec.careers_page_status = "error"
        rec.error = f"find_careers_page: {e}"
        err = ErrorRecord(
            company_name=news.company_name,
            company_website=news.company_website,
            news_url=news.news_url,
            stage="careers_find",
            error=str(e),
            timestamp=datetime.now(timezone.utc).isoformat(),
        )
        return ([], rec, err)

    rec.careers_page_url = careers_url
    rec.careers_page_status = status
    if status != "found" or not careers_url:
        rec.error = "careers page not found"
        return ([], rec, None)

    # 2. Fetch + parse jobs.
    try:
        jobs, source = fetch_jobs_for_careers_page(session, careers_url,
                                                   deadline=deadline)
    except Exception as e:
        rec.careers_page_status = "error"
        rec.error = f"fetch_jobs: {e}"
        err = ErrorRecord(
            company_name=news.company_name,
            company_website=news.company_website,
            news_url=news.news_url,
            stage="careers_fetch",
            error=str(e),
            timestamp=datetime.now(timezone.utc).isoformat(),
        )
        return ([], rec, err)

    rec.jobs_found = len(jobs)
    if source == "error":
        rec.careers_page_status = "error"
        rec.error = "careers page returned non-200"
        return ([], rec, None)

    # 3. Match each job.
    # We fetch each candidate job's JD page ONCE and use it for both the
    # fresher check AND the Python-stack check (when enabled). The JD-fetch
    # phase has its own sub-budget (JD_FETCH_BUDGET) AND a max-jobs cap.
    # Learned tables (diagnosed feedback, saasnews.feedback) extend the
    # built-in role/location/fresher tables; python_strict domains lose the
    # short-JD leniency below.
    learned_excl = policy.learned_title_exclusions if policy else None
    learned_tneg = policy.learned_title_negatives if policy else None
    learned_jneg = policy.learned_jd_negatives if policy else None
    learned_region = policy.learned_region_re if policy else None
    strict_domain = policy.python_strict(news.company_website) if policy else False
    reject_locs = policy.blocked_locations if policy else None
    jd_deadline = min(deadline, time.monotonic() + JD_FETCH_BUDGET)
    jd_fetched = 0
    matches: list[JobMatch] = []
    jobs_suppressed = 0
    scraped_at = datetime.now(timezone.utc).isoformat()
    for job in jobs:
        # Check overall company deadline.
        if time.monotonic() > deadline:
            break

        # 0. Learned suppression (closed feedback loop): skip known-bad job
        #    URLs and user-flagged locations before spending any work.
        if policy is not None and policy.job_suppressed(job.url, job.location):
            jobs_suppressed += 1
            continue

        role = match_role(job.title, extra_exclusions=learned_excl)
        if role is None:
            continue
        category, kw, role_conf = role
        loc = match_location(job.location, job.title,
                             reject_locations=reject_locs,
                             extra_region_re=learned_region)
        if loc is None:
            continue
        basis, loc_conf = loc

        # --- Fetch JD text (once) for both fresher + Python-stack checks ---
        # We need the JD if either fresher_only or python_stack is enabled AND
        # the title alone isn't sufficient.
        title_has_python = bool(PYTHON_IN_TITLE_RE.search(job.title))
        need_jd = False
        if fresher_only:
            need_jd = True
        if python_stack and not title_has_python:
            need_jd = True

        jd_text = ""
        if need_jd:
            if jd_fetched >= JD_FETCH_MAX_JOBS:
                # Cap reached — skip JD fetch, use title-only checks.
                pass
            elif time.monotonic() < jd_deadline:
                jd_url = job.apply_url or job.url
                if jd_url:
                    jd_text = fetch_jd_text(session, jd_url,
                                            deadline=jd_deadline)
                    jd_fetched += 1

        # --- Fresher filter (if enabled) ---
        fresher_basis = ""
        fresher_level = ""
        if fresher_only:
            is_fresher, fresher_basis, fresher_level = is_fresher_role(
                job.title, jd_text,
                extra_title_negatives=learned_tneg,
                extra_jd_negatives=learned_jneg,
            )
            if not is_fresher:
                continue

        # --- Python-stack filter (if enabled) ---
        if python_stack:
            if title_has_python:
                # Title explicitly mentions Python — accept.
                python_ok = True
            elif PYTHON_STACK_RE.search(jd_text):
                # JD mentions Python stack — accept.
                python_ok = True
            elif len(jd_text) < 300 and not strict_domain:
                # JD text is too short (JS-rendered page, error message, or
                # fetch failed). Be LENIENT: Python is the dominant language
                # for backend, AI/ML, data, and internship roles at SaaS
                # startups — especially in India. Accept these and let the
                # user verify manually via the Apply URL. Domains flagged by
                # diagnosed feedback lose this leniency.
                python_ok = True
            else:
                # JD available (300+ chars) but no Python mention — reject.
                python_ok = False
            if not python_ok:
                continue

        # Combined confidence:
        if role_conf == "high" and loc_conf == "high":
            conf = "high"
        elif role_conf == "low" or loc_conf == "medium":
            conf = "medium" if role_conf != "low" else "low"
        else:
            conf = "medium"
        # Downgrade confidence for implicit fresher detection (no explicit
        # "junior" / "fresher" keyword in title).
        if fresher_only and fresher_level == "implicit" and conf == "high":
            conf = "medium"
        matches.append(JobMatch(
            company_name=news.company_name,
            company_website=news.company_website,
            news_url=news.news_url,
            news_headline=news.headline,
            funding_round=news.funding_round,
            funding_date=news.funding_date,
            software_category=news.software_category,
            job_title=job.title,
            job_url=job.url,
            location=job.location,
            apply_url=job.apply_url,
            matched_keyword=kw,
            match_category=category,
            confidence=conf,
            location_basis=basis,
            careers_page_url=careers_url,
            posted_date=job.posted_date,
            fresher_basis=fresher_basis if fresher_only else "",
            fresher_level=fresher_level if fresher_only else "",
            scraped_at=scraped_at,
        ))
    # jobs_found stays the RAW parsed count; suppression is reported
    # separately so the workbook stays honest about both.
    rec.jobs_suppressed = jobs_suppressed
    rec.jobs_matched = len(matches)
    return (matches, rec, None)


# ---------------------------------------------------------------------------
# Stage 7: Excel output
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D1D5DB"),
    right=Side(style="thin", color="D1D5DB"),
    top=Side(style="thin", color="D1D5DB"),
    bottom=Side(style="thin", color="D1D5DB"),
)
HIGH_FILL = PatternFill("solid", fgColor="D1FAE5")   # green-100
MED_FILL = PatternFill("solid", fgColor="FEF3C7")    # amber-100
LOW_FILL = PatternFill("solid", fgColor="FEE2E2")    # red-100
ERROR_FILL = PatternFill("solid", fgColor="FEE2E2")


def _write_sheet(ws, headers: list[str], rows: list[list], col_widths: list[int],
                 color_col: Optional[int] = None,
                 color_map: Optional[dict] = None) -> None:
    """Helper: write headers + rows with auto-filter and standard styling."""
    # Headers
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER
    # Data
    for r_idx, row in enumerate(rows, 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = THIN_BORDER
            if color_col is not None and color_map and c_idx == color_col + 1:
                fill = color_map.get(str(val).lower())
                if fill:
                    cell.fill = fill
    # Column widths
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # Freeze header row
    ws.freeze_panes = "A2"
    # Auto-filter
    if rows:
        last_col = get_column_letter(len(headers))
        ws.auto_filter.ref = f"A1:{last_col}{len(rows) + 1}"


def _add_tracking_validations(ws, headers: list[str], row_count: int) -> None:
    """Attach dropdown validation to the Applied / Feedback tracking columns."""
    from openpyxl.worksheet.datavalidation import DataValidation

    applied_col = get_column_letter(headers.index("Applied") + 1)
    feedback_col = get_column_letter(headers.index("Feedback") + 1)
    last = max(row_count + 1, 2)

    dv_applied = DataValidation(
        type="list", formula1='"Yes,No"', allow_blank=True,
        promptTitle="Applied?", prompt="Mark Yes once you have applied.",
    )
    ws.add_data_validation(dv_applied)
    dv_applied.add(f"{applied_col}2:{applied_col}{last}")

    constraint_list = ",".join(FEEDBACK_CONSTRAINTS)
    dv_feedback = DataValidation(
        type="list", formula1=f'"{constraint_list}"', allow_blank=True,
        # Non-strict: typed values are allowed, so several constraints can be
        # combined in one cell (multi-select), e.g. "not_fresher, bad_location".
        showErrorMessage=False,
        promptTitle="What was wrong?",
        prompt="Pick a constraint — or combine SEVERAL separated by commas "
               "(e.g. not_fresher, bad_location). The watcher harvests these "
               "and re-tunes the filters automatically.",
    )
    ws.add_data_validation(dv_feedback)
    dv_feedback.add(f"{feedback_col}2:{feedback_col}{last}")


MATCHES_HEADERS = [
    "Company", "Job Title", "Category", "Confidence", "Matched Keyword",
    "Location", "Location Basis", "Fresher Basis", "Fresher Level",
    "Apply URL", "Job URL",
    "Company Website", "Careers Page", "Funding Round", "Funding Date",
    "Software Category", "News Headline", "News URL",
    # Feed + tracking columns (user-owned: Applied / Feedback / Notes).
    "First Seen", "Applied", "Feedback", "Notes",
]

MATCHES_WIDTHS = [22, 42, 12, 12, 20, 28, 18, 22, 14, 50, 50, 28, 38, 14, 14,
                  22, 50, 50, 22, 10, 20, 32]


def _match_rows(matches: list[JobMatch]) -> list[list]:
    """Dataclass list → Matches-sheet row values (tracking columns blank)."""
    return [
        [
            m.company_name, m.job_title, m.match_category, m.confidence,
            m.matched_keyword, m.location, m.location_basis,
            m.fresher_basis, m.fresher_level,
            m.apply_url, m.job_url,
            m.company_website, m.careers_page_url, m.funding_round,
            m.funding_date, m.software_category, m.news_headline, m.news_url,
            m.scraped_at, "", "", "",
        ]
        for m in matches
    ]


def _sort_feed_rows(rows: list[list]) -> None:
    """Sort newest-scraped first ("latest scraped jobs at the top"), with
    confidence rank then company as stable tiebreakers. Timestamp-less rows
    (old checkpoints) sink to the bottom. Two passes keep the secondary
    order intact under the primary sort."""
    conf_order = {"high": 0, "medium": 1, "low": 2}
    fs_idx = MATCHES_HEADERS.index("First Seen")

    def _parse_seen(value: str) -> datetime:
        try:
            return datetime.fromisoformat(str(value))
        except (TypeError, ValueError):
            return datetime.min.replace(tzinfo=timezone.utc)

    rows.sort(key=lambda r: (conf_order.get(str(r[3]).lower(), 9),
                             str(r[0]).lower()))
    rows.sort(key=lambda r: _parse_seen(r[fs_idx]), reverse=True)


def write_workbook(
    out_path: str,
    matches: list[JobMatch],
    companies: list[CompanyRecord],
    errors: list[ErrorRecord],
    run_meta: dict,
    applied_matches: Optional[list[JobMatch]] = None,
) -> None:
    wb = Workbook()
    wb.properties.creator = "saasnews-job-finder"

    # --- Sheet 1: Matches (active feed — applied jobs excluded) ---
    ws1 = wb.active
    ws1.title = "Matches"
    headers1 = MATCHES_HEADERS
    rows1 = _match_rows(matches)
    _sort_feed_rows(rows1)
    _write_sheet(
        ws1, headers1, rows1,
        col_widths=MATCHES_WIDTHS,
        color_col=3,  # Confidence column (0-indexed) → cell column 4
        color_map={"high": HIGH_FILL, "medium": MED_FILL, "low": LOW_FILL},
    )
    # Dropdowns for the user-owned tracking columns.
    _add_tracking_validations(ws1, headers1, len(rows1))

    # --- Sheet 1b: Applied (jobs you already acted on; tracked separately) ---
    if applied_matches:
        ws_applied = wb.create_sheet("Applied")
        applied_rows = _match_rows(applied_matches)
        _sort_feed_rows(applied_rows)
        _write_sheet(
            ws_applied, headers1, applied_rows,
            col_widths=MATCHES_WIDTHS,
            color_col=3,
            color_map={"high": HIGH_FILL, "medium": MED_FILL, "low": LOW_FILL},
        )
        _add_tracking_validations(ws_applied, headers1, len(applied_rows))

    # --- Sheet 2: All Companies ---
    ws2 = wb.create_sheet("All Companies")
    headers2 = [
        "Company", "Website", "Funding Round", "Funding Date", "Lead Investor",
        "Software Category", "Careers Page", "Careers Status",
        "Jobs Found", "Jobs Matched", "Jobs Suppressed", "News Headline", "News URL", "Error",
    ]
    rows2 = []
    for c in companies:
        rows2.append([
            c.company_name, c.company_website, c.funding_round, c.funding_date,
            c.lead_investor, c.software_category, c.careers_page_url,
            c.careers_page_status, c.jobs_found, c.jobs_matched,
            c.jobs_suppressed, c.news_headline, c.news_url, c.error,
        ])
    rows2.sort(key=lambda r: r[0].lower())
    _write_sheet(
        ws2, headers2, rows2,
        col_widths=[22, 30, 16, 14, 22, 24, 38, 14, 12, 12, 14, 50, 50, 40],
    )

    # --- Sheet 3: Errors ---
    ws3 = wb.create_sheet("Errors")
    headers3 = ["Company", "Website", "News URL", "Stage", "Error", "Timestamp"]
    rows3 = [[e.company_name, e.company_website, e.news_url, e.stage, e.error,
              e.timestamp] for e in errors]
    _write_sheet(
        ws3, headers3, rows3,
        col_widths=[22, 30, 50, 16, 60, 26],
    )

    # --- Sheet 4: Run Summary ---
    ws4 = wb.create_sheet("Run Summary")
    summary_rows = [
        ["Run Timestamp", run_meta.get("timestamp", "")],
        ["Total News Articles Discovered", run_meta.get("news_total", 0)],
        ["Companies With Company Website", run_meta.get("companies_total", 0)],
        ["Companies With Careers Page Found", run_meta.get("careers_found", 0)],
        ["Companies With No Careers Page", run_meta.get("careers_not_found", 0)],
        ["Companies With Errors", run_meta.get("errors", 0)],
        ["Total Jobs Parsed", run_meta.get("jobs_parsed", 0)],
        ["Total Matching Jobs", run_meta.get("jobs_matched", 0)],
        ["  High Confidence", run_meta.get("high_conf", 0)],
        ["  Medium Confidence", run_meta.get("med_conf", 0)],
        ["  Low Confidence", run_meta.get("low_conf", 0)],
        ["Feedback Records Applied", run_meta.get("feedback_records", 0)],
        ["Learned Patterns Active", run_meta.get("learned_patterns", 0)],
        ["Companies Skipped (learned policy)", run_meta.get("companies_skipped_policy", 0)],
        ["Known-Bad Jobs Suppressed", run_meta.get("jobs_suppressed_total", 0)],
    ]
    # Feedback analysis: constraint × signal that let it through.
    for line in run_meta.get("feedback_breakdown", []):
        summary_rows.append(["  Feedback Analysis", line])
    for line in run_meta.get("learned_breakdown", []):
        summary_rows.append(["  Learned Policy", line])
    summary_rows.append(["CLI Args", json.dumps(run_meta.get("args", {}))])
    for r, (k, v) in enumerate(summary_rows, 1):
        ws4.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws4.cell(row=r, column=2, value=v)
    ws4.column_dimensions["A"].width = 38
    ws4.column_dimensions["B"].width = 80

    wb.save(out_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="TheSaaSNews → Backend/AI/Full-stack jobs (India or Remote-Worldwide)",
    )
    p.add_argument("--limit", type=int, default=0,
                   help="Max companies to process (0 = all).")
    p.add_argument("--since-days", type=int, default=0,
                   help="Only news published in last N days (0 = no filter).")
    p.add_argument("--output-dir", default="./download",
                   help="Output directory for xlsx (relative paths resolve "
                        "against the working directory).")
    p.add_argument("--concurrency", type=int, default=5,
                   help="Parallel workers for careers-page processing.")
    p.add_argument("--news-pages", type=int, default=0,
                   help="Max news pagination pages to scan (default: 0 = ALL pages, "
                        "auto-detected by following rel='next' links until the end). "
                        "Set to a specific number like 50 for a quick test.")
    p.add_argument("--start-page", type=int, default=1,
                   help="Start from this news page (1-indexed). Useful for "
                        "running the scraper in chunks across sessions.")
    p.add_argument("--priority-recent", type=int, default=0,
                   help="Process the N most recent articles FIRST (from page 1) "
                        "before resuming from checkpoint. Use this in daily cron "
                        "to ensure today's new articles are processed immediately. "
                        "Recommended: --priority-recent 100 (processes ~100 newest "
                        "articles first, then resumes from checkpoint).")
    p.add_argument("--verbose", "-v", action="store_true", help="Debug logging.")
    p.add_argument("--fresher-only", action="store_true",
                   help="Only include fresher-eligible roles (0-2 years experience). "
                        "Fetches each candidate job's JD page to verify "
                        "experience requirements. Slower but precise.")
    p.add_argument("--no-python-stack", action="store_true",
                   help="Disable Python-stack verification. By default, the scraper "
                        "requires the title OR JD to mention Python/Django/Flask/"
                        "FastAPI/PyTorch/etc. Use this flag to accept any stack.")
    p.add_argument("--reset-checkpoint", action="store_true",
                   help="Delete the checkpoint file before running. Use this when "
                        "you want a fresh start (e.g., changed filter settings).")
    p.add_argument("--batch-size", type=int, default=0,
                   help="Process at most N articles this run (0 = no limit). "
                        "Use this to process in chunks of 500: run the script "
                        "repeatedly with --batch-size 500 until all articles "
                        "are processed. Resume is automatic via checkpoint.")
    return p.parse_args(argv)


def _checkpoint_path(output_dir: str) -> str:
    """Path to the JSONL checkpoint file (one record per processed company)."""
    return os.path.join(output_dir, "checkpoint.jsonl")


def _feedback_path(output_dir: str) -> str:
    """Path to the append-only user-feedback log feeding the learned policy."""
    return os.path.join(output_dir, "feedback.jsonl")


def _load_checkpoint(output_dir: str) -> dict[str, dict]:
    """Load prior checkpoint. Returns dict keyed by news_url → record.
    Each record has: company, matches (list), company_rec (dict), error (dict|None).
    """
    path = _checkpoint_path(output_dir)
    seen: dict[str, dict] = {}
    if not os.path.exists(path):
        return seen
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                    nu = rec.get("news_url", "")
                    if nu:
                        seen[nu] = rec
                except json.JSONDecodeError:
                    continue
    except Exception as e:
        log.warning("Could not load checkpoint: %s", e)
    return seen


def _append_checkpoint(output_dir: str, record: dict) -> None:
    """Atomically append one record to the JSONL checkpoint file.
    Uses a temp file + rename for crash safety.
    """
    path = _checkpoint_path(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    line = json.dumps(record, ensure_ascii=False, default=str) + "\n"
    # Atomic append (O_APPEND on POSIX is atomic for small writes).
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(line)
            f.flush()
            os.fsync(f.fileno())
    except Exception as e:
        log.warning("Could not write checkpoint: %s", e)


def _write_xlsx_from_checkpoint(
    output_dir: str, checkpoint: dict[str, dict], args_dict: dict,
    news_total: Optional[int] = None,
    meta_extras: Optional[dict] = None,
    applied_urls: Optional[set] = None,
) -> str:
    """Rebuild the xlsx workbook from the in-memory checkpoint.
    Called periodically during the run AND at the end.
    `news_total`: real discovered-article count for the Run Summary; when
    absent, the checkpointed company count is reported instead of a guess.
    `meta_extras`: honest caller-computed metrics merged into the Run Summary.
    `applied_urls`: jobs the user already applied to (from the feedback log);
    they are excluded from the Matches feed and listed on a dedicated
    'Applied' sheet instead.
    """
    applied_urls = applied_urls or set()
    all_matches: list[JobMatch] = []
    applied_matches: list[JobMatch] = []
    all_companies: list[CompanyRecord] = []
    all_errors: list[ErrorRecord] = []
    jobs_parsed_total = 0
    for rec in checkpoint.values():
        # Reconstruct CompanyRecord
        cr = rec.get("company_rec", {})
        company_rec = CompanyRecord(
            company_name=cr.get("company_name", ""),
            company_website=cr.get("company_website", ""),
            news_url=cr.get("news_url", ""),
            news_headline=cr.get("news_headline", ""),
            funding_round=cr.get("funding_round", ""),
            funding_date=cr.get("funding_date", ""),
            lead_investor=cr.get("lead_investor", ""),
            software_category=cr.get("software_category", ""),
            careers_page_url=cr.get("careers_page_url", ""),
            careers_page_status=cr.get("careers_page_status", ""),
            jobs_found=cr.get("jobs_found", 0) or 0,
            jobs_matched=cr.get("jobs_matched", 0) or 0,
            jobs_suppressed=cr.get("jobs_suppressed", 0) or 0,
            error=cr.get("error", ""),
        )
        all_companies.append(company_rec)
        jobs_parsed_total += company_rec.jobs_found
        # Reconstruct matches
        for m_dict in rec.get("matches", []):
            m = JobMatch(
                company_name=m_dict.get("company_name", ""),
                company_website=m_dict.get("company_website", ""),
                news_url=m_dict.get("news_url", ""),
                news_headline=m_dict.get("news_headline", ""),
                funding_round=m_dict.get("funding_round", ""),
                funding_date=m_dict.get("funding_date", ""),
                software_category=m_dict.get("software_category", ""),
                job_title=m_dict.get("job_title", ""),
                job_url=m_dict.get("job_url", ""),
                location=m_dict.get("location", ""),
                apply_url=m_dict.get("apply_url", ""),
                matched_keyword=m_dict.get("matched_keyword", ""),
                match_category=m_dict.get("match_category", ""),
                confidence=m_dict.get("confidence", ""),
                location_basis=m_dict.get("location_basis", ""),
                careers_page_url=m_dict.get("careers_page_url", ""),
                posted_date=m_dict.get("posted_date", ""),
                fresher_basis=m_dict.get("fresher_basis", ""),
                fresher_level=m_dict.get("fresher_level", ""),
                scraped_at=m_dict.get("scraped_at", ""),
            )
            if m.job_url and m.job_url in applied_urls:
                applied_matches.append(m)
            else:
                all_matches.append(m)
        # Reconstruct error
        if rec.get("error"):
            err_dict = rec["error"]
            all_errors.append(ErrorRecord(
                company_name=err_dict.get("company_name", ""),
                company_website=err_dict.get("company_website", ""),
                news_url=err_dict.get("news_url", ""),
                stage=err_dict.get("stage", ""),
                error=err_dict.get("error", ""),
                timestamp=err_dict.get("timestamp", ""),
            ))
    # Write to a stable filename (not timestamped) so users/cron can find it.
    out_path = os.path.join(output_dir, "saasnews_jobs_latest.xlsx")
    careers_found = sum(1 for c in all_companies if c.careers_page_status == "found")
    careers_nf = sum(1 for c in all_companies if c.careers_page_status == "not_found")
    err_count = len(all_errors) + sum(1 for c in all_companies if c.careers_page_status == "error")
    conf_counts = {"high": 0, "medium": 0, "low": 0}
    for m in all_matches:
        conf_counts[m.confidence] = conf_counts.get(m.confidence, 0) + 1
    run_meta = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        # Honest metrics only: real count when provided, else the checkpointed
        # company count — never an estimate. (Zero is a valid count.)
        "news_total": news_total if news_total is not None else len(all_companies),
        "feedback_records": 0,
        "companies_skipped_policy": 0,
        "jobs_suppressed_total": sum(c.jobs_suppressed for c in all_companies),
        "companies_total": len(all_companies),
        "careers_found": careers_found,
        "careers_not_found": careers_nf,
        "errors": err_count,
        "jobs_parsed": jobs_parsed_total,
        "jobs_matched": len(all_matches),
        "high_conf": conf_counts.get("high", 0),
        "med_conf": conf_counts.get("medium", 0),
        "low_conf": conf_counts.get("low", 0),
        "args": args_dict,
    }
    if meta_extras:
        run_meta.update(meta_extras)
    # Write to temp file then rename (atomic).
    tmp_path = out_path + ".tmp"
    try:
        write_workbook(tmp_path, all_matches, all_companies, all_errors,
                       run_meta, applied_matches=applied_matches)
        os.replace(tmp_path, out_path)
    except Exception as e:
        log.warning("Could not write xlsx: %s", e)
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    return out_path


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )

    os.makedirs(args.output_dir, exist_ok=True)

    # ---- Load checkpoint FIRST (before any network calls) ----
    # This is critical: if the user is resuming, we want to show progress
    # immediately, not after re-discovering 994 pages.
    if args.reset_checkpoint:
        ckpt_path = _checkpoint_path(args.output_dir)
        if os.path.exists(ckpt_path):
            os.remove(ckpt_path)
            log.info("Deleted checkpoint file: %s", ckpt_path)
    checkpoint = _load_checkpoint(args.output_dir)
    if checkpoint:
        log.info("Loaded checkpoint: %d companies already processed.", len(checkpoint))

    session = make_session()

    # ---- Stage 1: discover news URLs ----
    # --news-pages 0 means "scan ALL pages" (auto-detect by following rel="next").
    max_pages = args.news_pages if args.news_pages > 0 else 100000
    if args.priority_recent:
        # In priority-recent mode, only discover the first N pages (newest).
        # This is fast and ensures today's articles are processed immediately.
        priority_pages = max(1, args.priority_recent // 12 + 1)  # ~12 articles per page
        log.info("Priority mode: discovering %d most recent pages first...", priority_pages)
        news_urls = discover_news_urls(session, max_pages=priority_pages,
                                       start_page=1)
        log.info("Discovered %d recent article URLs.", len(news_urls))
    else:
        log.info("Discovering news article URLs (all pages, auto-detected)...")
        news_urls = discover_news_urls(session, max_pages=max_pages,
                                       start_page=args.start_page)
        log.info("Discovered %d news article URLs.", len(news_urls))

    # ---- Stage 2: parse articles in PARALLEL ----
    # Filter out URLs already in checkpoint BEFORE parsing (saves time on resume).
    pending_urls = [u for u in news_urls if u not in checkpoint]
    skipped_count = len(news_urls) - len(pending_urls)
    if skipped_count:
        log.info("Skipping %d already-processed articles (resume).", skipped_count)

    # Apply --limit to pending URLs only.
    if args.limit and len(pending_urls) > args.limit:
        pending_urls = pending_urls[:args.limit]
        log.info("Limited to %d articles (--limit).", args.limit)

    # Apply --batch-size: only process N articles this run.
    batch_size = getattr(args, "batch_size", 0) or 0
    if batch_size and len(pending_urls) > batch_size:
        log.info("Batch mode: processing %d of %d pending articles (--batch-size %d).",
                 batch_size, len(pending_urls), batch_size)
        pending_urls = pending_urls[:batch_size]

    # ---- Closed feedback loop: compile learned policy BEFORE scraping ----
    feedback_path = _feedback_path(args.output_dir)
    feedback_records = load_feedback(feedback_path)
    learned_entries = load_learned(os.path.join(args.output_dir, "learned.jsonl"))
    policy = build_policy(feedback_records, learned_entries=learned_entries)
    if feedback_records:
        log.info("Feedback loop: %d records → %d blocked jobs, %d flagged "
                 "locations, %d companies with strikes.",
                 len(feedback_records), len(policy.blocked_job_urls),
                 len(policy.blocked_locations),
                 sum(1 for c in policy.company_strikes.values()
                     if c >= policy.strike_threshold))
    if learned_entries:
        log.info("Learned policy: %d diagnosed pattern(s) active (%s).",
                 len(learned_entries),
                 ", ".join(f"{k}={v}" for k, v in
                           [(l.split(": ")[0], l.split(": ")[1])
                            for l in summarize_learned(learned_entries)]))
    companies_skipped_policy = 0

    def flush_xlsx() -> str:
        return _write_xlsx_from_checkpoint(
            args.output_dir, checkpoint, args_dict,
            news_total=len(news_urls),
            meta_extras={
                "feedback_records": len(feedback_records),
                "companies_skipped_policy": companies_skipped_policy,
                "feedback_breakdown": summarize_feedback(feedback_records),
                "learned_patterns": len(learned_entries),
                "learned_breakdown": summarize_learned(learned_entries),
            },
            applied_urls=policy.applied_job_urls,
        )

    if not pending_urls:
        log.info("Nothing to do — all articles already processed.")
        flush_xlsx()
        return 0

    log.info("Parsing %d articles in parallel (concurrency=%d)...",
             len(pending_urls), min(args.concurrency, 10))
    news_items: list[NewsItem] = []

    def _parse_worker(url: str) -> Optional[NewsItem]:
        s = make_session()
        ni = parse_news_article(s, url)
        if ni is None:
            return None
        # since-days filter
        if args.since_days and ni.published_at:
            try:
                pub_dt = datetime.fromisoformat(
                    ni.published_at.replace("Z", "+00:00")
                )
                age_days = (datetime.now(timezone.utc) - pub_dt).days
                if age_days > args.since_days:
                    return None
            except Exception:
                pass
        return ni

    with ThreadPoolExecutor(max_workers=min(args.concurrency, 10)) as ex:
        futures = {ex.submit(_parse_worker, url): url for url in pending_urls}
        for i, fut in enumerate(as_completed(futures), 1):
            try:
                ni = fut.result(timeout=30)
            except Exception:
                ni = None
            if ni is not None:
                news_items.append(ni)
            if i % 100 == 0 or i == len(pending_urls):
                log.info("  Parsed %d/%d articles (%d with company website)",
                         i, len(pending_urls), len(news_items))

    log.info("Got %d companies with website.", len(news_items))

    # ---- Filter out already-processed (in case checkpoint grew during parse) ----
    pending_items = [ni for ni in news_items if ni.news_url not in checkpoint]
    if len(pending_items) < len(news_items):
        log.info("Skipping %d newly-processed articles.", len(news_items) - len(pending_items))
    log.info("Processing %d companies...", len(pending_items))

    # ---- Stages 3-5: per-company processing in parallel ----
    fresher_only = bool(getattr(args, "fresher_only", False))
    python_stack = not bool(getattr(args, "no_python_stack", False))
    args_dict = vars(args)

    # Companies with enough feedback strikes are skipped outright and
    # checkpointed as such (never refetched on later runs).
    struck = [ni for ni in pending_items
              if policy.company_suppressed(ni.company_website)]
    if struck:
        pending_items = [ni for ni in pending_items
                         if not policy.company_suppressed(ni.company_website)]
        log.info("Skipping %d companies via learned policy (feedback strikes).",
                 len(struck))
    for ni in struck:
        rec = CompanyRecord(
            company_name=ni.company_name,
            company_website=ni.company_website,
            news_url=ni.news_url,
            news_headline=ni.headline,
            funding_round=ni.funding_round,
            funding_date=ni.funding_date,
            lead_investor=ni.lead_investor,
            software_category=ni.software_category,
            careers_page_status="skipped_policy",
            error=(f"company skipped: {policy.company_strikes[domain_of(ni.company_website)]} "
                   f"feedback strikes (threshold {policy.strike_threshold})"),
        )
        ckpt_record = {
            "news_url": ni.news_url,
            "company_rec": asdict(rec),
            "matches": [],
            "error": None,
        }
        checkpoint[ni.news_url] = ckpt_record
        _append_checkpoint(args.output_dir, ckpt_record)
    companies_skipped_policy = len(struck)

    def _worker(ni: NewsItem):
        s = make_session()
        return process_company(ni, s, fresher_only=fresher_only,
                               python_stack=python_stack, policy=policy)

    FLUSH_EVERY = 10
    processed_since_flush = 0
    total_done = 0

    with ThreadPoolExecutor(max_workers=args.concurrency) as ex:
        # Process in sub-batches of 500 to avoid OOM with 12K futures.
        SUB_BATCH = 500
        for batch_start in range(0, len(pending_items), SUB_BATCH):
            batch = pending_items[batch_start:batch_start + SUB_BATCH]
            futures = {ex.submit(_worker, ni): ni for ni in batch}
            for fut in as_completed(futures):
                ni = futures[fut]
                total_done += 1
                try:
                    matches, rec, err = fut.result(timeout=PER_COMPANY_BUDGET + 30)
                except Exception as e:
                    log.warning("  [%d/%d] %s — worker error: %s",
                                total_done, len(pending_items), ni.company_name, e)
                    rec = CompanyRecord(
                        company_name=ni.company_name,
                        company_website=ni.company_website,
                        news_url=ni.news_url,
                        news_headline=ni.headline,
                        funding_round=ni.funding_round,
                        funding_date=ni.funding_date,
                        lead_investor=ni.lead_investor,
                        software_category=ni.software_category,
                        careers_page_status="error",
                        error=f"worker crash: {e}",
                    )
                    err = ErrorRecord(
                        company_name=ni.company_name,
                        company_website=ni.company_website,
                        news_url=ni.news_url,
                        stage="worker",
                        error=str(e),
                        timestamp=datetime.now(timezone.utc).isoformat(),
                    )
                    matches = []

                ckpt_record = {
                    "news_url": ni.news_url,
                    "company_rec": asdict(rec),
                    "matches": [asdict(m) for m in matches],
                    "error": asdict(err) if err else None,
                }
                checkpoint[ni.news_url] = ckpt_record
                _append_checkpoint(args.output_dir, ckpt_record)

                processed_since_flush += 1
                if matches:
                    log.info("  [%d/%d] %s — %d jobs, %d matches",
                             total_done, len(pending_items), ni.company_name,
                             rec.jobs_found, len(matches))
                elif total_done % 20 == 0:
                    log.info("  [%d/%d] %s — no matches (%s, %d jobs)",
                             total_done, len(pending_items), ni.company_name,
                             rec.careers_page_status, rec.jobs_found)

                if processed_since_flush >= FLUSH_EVERY:
                    flush_xlsx()
                    processed_since_flush = 0

    # ---- Final xlsx write ----
    out_path = flush_xlsx()
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    archive_path = os.path.join(args.output_dir, f"saasnews_jobs_{ts}.xlsx")
    try:
        shutil.copy2(out_path, archive_path)
    except Exception:
        pass

    # ---- Final summary ----
    all_companies = [r["company_rec"] for r in checkpoint.values()]
    all_matches_count = sum(len(r.get("matches", [])) for r in checkpoint.values())
    careers_found = sum(1 for c in all_companies if c.get("careers_page_status") == "found")
    err_count = sum(1 for r in checkpoint.values() if r.get("error")) + \
                sum(1 for c in all_companies if c.get("careers_page_status") == "error")
    conf_counts = {"high": 0, "medium": 0, "low": 0}
    for r in checkpoint.values():
        for m in r.get("matches", []):
            conf_counts[m.get("confidence", "low")] = conf_counts.get(m.get("confidence", "low"), 0) + 1

    log.info("=" * 60)
    log.info("DONE.  Output: %s", out_path)
    log.info("  Archive copy      : %s", archive_path)
    log.info("  Checkpoint file   : %s", _checkpoint_path(args.output_dir))
    log.info("  Total in checkpoint: %d", len(checkpoint))
    log.info("  Careers pages found: %d", careers_found)
    log.info("  Errors              : %d", err_count)
    log.info("  Matching jobs       : %d", all_matches_count)
    log.info("    high confidence   : %d", conf_counts.get("high", 0))
    log.info("    medium confidence : %d", conf_counts.get("medium", 0))
    log.info("    low confidence    : %d", conf_counts.get("low", 0))
    return 0


def cli() -> None:
    """Console-script entry point (zero-arg callable)."""
    raise SystemExit(main(sys.argv[1:]))


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
