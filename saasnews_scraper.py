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
import re
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

# Role keywords -- grouped by category. PYTHON-STACK ONLY.
# Order matters: earlier patterns take precedence for category assignment.
# We accept AI/ML, Backend (Python/Django/Flask/FastAPI), Full-stack, Data,
# and Internship variants of any of the above.
ROLE_PATTERNS: list[tuple[str, list[str]]] = [
    # Internship FIRST — so "Software Engineer Intern" is categorized as
    # Internship (not Full-stack). Interns are explicitly fresher-eligible.
    ("Internship", [
        r"\bintern\b",
        r"\binternship\b",
        r"\bintern[\s\-]*engineer",
        r"\bintern[\s\-]*developer",
        r"\bsummer[\s\-]*intern",
        r"\bco[\s\-]*op\b",
        r"\bfellow\b", r"\bfellowship\b",
        r"\bapprentice\b", r"\bapprenticeship\b",
    ]),
    ("AI/ML", [
        r"\bai[\s\-/]*engineer",
        r"\bml[\s\-/]*engineer",
        r"machine[\s\-]*learning",
        r"\bllm\b",
        r"\bnlp\b",
        r"applied[\s\-]*scientist",
        r"research[\s\-]*scientist",
        r"ai[\s\-/]*research",
        r"\bgenai\b", r"\bgen[\s\-]*ai\b",
        r"deep[\s\-]*learning",
        r"ai[\s\-/]*infrastructure",
        r"model[\s\-]*engineer",
        r"\bml[\s\-/]*ops\b",
        r"prompt[\s\-]*engineer",
    ]),
    ("Backend", [
        # Explicit backend / server-side
        r"backend", r"back[\s\-]*end",
        r"server[\s\-]*side",
        r"platform[\s\-]*engineer",
        r"\bapi\b[\s\-]*engineer",
        r"infrastructure[\s\-]*engineer",
        # Python-stack language/framework signals (STRONG match)
        r"\bpython\b",
        r"\bdjango\b",
        r"\bflask\b",
        r"\bfastapi\b",
        r"\bpytorch\b",
        r"\btensorflow\b",
        r"\bpandas\b",
        r"\bnumpy\b",
        # Generic backend role keywords
        r"systems[\s\-]*engineer",
        r"site[\s\-]*reliability",  # SRE
    ]),
    ("Data", [
        r"data[\s\-]*engineer",
        r"data[\s\-]*scientist",
        r"analytics[\s\-]*engineer",
        r"data[\s\-]*platform",
        r"data[\s\-]*infrastructure",
    ]),
    ("Full-stack", [
        r"full[\s\-]*stack", r"fullstack",
        r"software[\s\-]*engineer",
        r"software[\s\-]*developer",
        r"\bSDE\d?\b",
        r"\bSWE\d?\b",
        r"application[\s\-]*engineer",
        r"web[\s\-]*developer",
        r"\bengineer\b",  # catch-all (low confidence) — Python-stack verified via JD
    ]),
]

# Python-stack verification patterns. Used to check the JD text for Python or
# a Python framework/library. If the title doesn't explicitly mention Python,
# the JD must contain at least one of these tokens for the job to be accepted.
PYTHON_STACK_RE = re.compile(
    r"\b(python|django|flask|fastapi|pyramid|tornado|bottle|"
    r"pytorch|tensorflow|keras|scikit[\s\-]*learn|pandas|numpy|"
    r"scipy|matplotlib|jupyter|airflow|celery|"
    r"langchain|llamaindex|transformers|huggingface|"
    r"opencv|pillow|spacy|nltk|gensim)\b",
    re.IGNORECASE,
)

# Python-stack signal in TITLE (strong — skip JD check).
PYTHON_IN_TITLE_RE = re.compile(
    r"\b(python|django|flask|fastapi|pytorch|tensorflow|pandas|numpy)\b",
    re.IGNORECASE,
)

# Explicit exclusions -- if title matches, we skip (to avoid false positives).
# NOTE: "intern" / "internship" are NO LONGER excluded (user wants internships).
TITLE_EXCLUSIONS = [
    r"\bfrontend[\s\-]*only\b",
    r"\bfront[\s\-]*end\b(?!.*back)",  # frontend without backend mention
    r"\bmobile\b(?!.*back)",
    r"\bios\b(?!.*back)",
    r"\bandroid\b(?!.*back)",
    r"\bdesign\b", r"\bdesigner\b",
    r"\bmarketing\b", r"\bsales\b", r"\bgrowth\b",
    r"\bproduct[\s\-]*manager\b", r"\bpm\b",
    r"\brecruit", r"\bhr\b", r"\bpeople\b",
    r"\bfinance\b", r"\baccounting\b",
    r"\blegal\b", r"\bcompliance\b",
    r"\bsupport\b", r"\bcustomer[\s\-]*success\b",
    r"\bcontent\b", r"\bwriter\b",
    r"\bdevrel\b", r"\bdeveloper[\s\-]*relations\b",
    # Sales/consulting/revenue engineering roles (not backend/AI/full-stack)
    r"\bsales[\s\-]*engineer\b",
    r"\bsolutions[\s\-]*engineer\b",
    r"\bgtm[\s\-]*engineer\b",
    r"\brevenue[\s\-]*engineer\b",
    r"\baccount[\s\-]*engineer\b",
    r"\bimplementation[\s\-]*engineer\b",
    r"\bcustomer[\s\-]*engineer\b",
    r"\bsupport[\s\-]*engineer\b",
    r"\bpartner[\s\-]*engineer\b",
    r"\bfield[\s\-]*engineer\b",
    # Non-Python-stack language engineers (user wants Python stack only).
    # Exclude these to filter out Java/Go/Rust/C++/C#/Ruby/Elixir/Swift/Kotlin
    # specialists — even if the title contains "backend", these are explicitly
    # tied to a non-Python stack.
    r"\bjava\b(?!script)",       # Java (not JavaScript)
    r"\bgolang\b", r"\bgo[\s\-]*lang\b",
    r"\brust\b(?!.*python)",
    r"(?:^|\s)c\+\+(?:\s|$)", r"\bcpp\b",
    r"\bcsharp\b", r"\bc#[\s\-]*engineer",
    r"\bruby\b(?!.*rails.*python)",
    r"\belixir\b",
    r"\bswift\b(?!.*python)",
    r"\bkotlin\b(?!.*python)",
    r"\bphp\b(?!.*python)",
    r"\bscala\b(?!.*python)",
    r"\bclr\b",
    r"\bnet[\s\-]*core\b", r"\b\.net\b",
]

# ---------------------------------------------------------------------------
# Fresher / 0-experience filter
# ---------------------------------------------------------------------------
# A job is "fresher-eligible" if the role requires 0-2 years of experience.
#
# We use a MULTI-SIGNAL approach for maximum reliability:
#   1. Title-level POSITIVE signal (Junior/Fresher/Intern/SDE-1/L1) → fresher
#   2. Title-level NEGATIVE signal (Senior/Staff/Principal/Lead/Sr./II/III) → NOT fresher
#   3. JD-text explicit years extraction (parse "X+ years" → if X >= 3, NOT fresher)
#   4. JD-text POSITIVE signals ("0-2 years", "fresher", "entry level")
#   5. If JD unavailable (JS-rendered) AND title is ambiguous → REJECT (conservative)
#
# This is stricter than before: ambiguous titles without JD verification are
# REJECTED rather than accepted. This prevents senior jobs from slipping through.

# Title-level POSITIVE fresher signals (case-insensitive).
# Each pattern is self-contained with proper boundaries.
FRESHER_TITLE_POSITIVE = re.compile(
    r"(?:"
    r"\bfresher\b|\bfreshers\b|\bfresh[\s\-]*graduate\b|"
    r"\bgraduate[\s\-]*engineer\b|\bgraduate[\s\-]*trainee\b|"
    r"\bentry[\s\-]*level\b|\bentry[\s\-]*level[\s\-]*role\b|"
    r"\bjunior\b|\bjr\b|"
    r"\bassociate\b|\btrainee\b|\bapprentice\b|\bapprenticeship\b|"
    r"\bintern\b|\binternship\b|\bsummer[\s\-]*intern\b|\bco[\s\-]*op\b|"
    r"\bfellow\b|\bfellowship\b|"
    r"\bsde[\s\-]*1\b|\bsde[\s\-]*i\b|\bswe[\s\-]*1\b|\bswe[\s\-]*i\b|"
    r"\bl1\b|\bl2\b|\blevel[\s\-]*1\b|\blevel[\s\-]*2\b|"
    r"\bengineer[\s\-]*i\b|\bengineer[\s\-]*1\b|"
    r"\b0[\s\-]*(?:year|yr)s?\b|\b0\+\s*(?:year|yr)|"
    r"\b1[\s\-]*(?:year|yr)s?\b|\b1\+\s*(?:year|yr)"
    r")",
    re.IGNORECASE,
)

# Title-level NEGATIVE seniority signals — if present, NOT a fresher role.
# NOTE: Each alternative is self-contained. We do NOT wrap the whole group in
# \b(...) because that breaks matching for abbreviations like "Sr." (the \b
# after the group fails between "." and " ").
# NOTE: "intern" / "internship" / "fellow" / "apprentice" are NOT here.
FRESHER_TITLE_NEGATIVE = re.compile(
    r"(?:"
    r"\bsenior\b|\bsr\b|\bsr\.|\bsr\s|"
    r"\bstaff\b|\bprincipal\b|\blead\b|"
    r"\bmanager\b|\bhead[\s\-]+of\b|\bdirector\b|"
    r"\bvp\b|\bvice[\s\-]*president\b|\bchief\b|\bcto\b|"
    r"\barchitect\b|\bfounding\b|\bexpert\b|"
    r"\bindividual[\s\-]*contributor\b|\bic\b|"
    # Level 3+ (Amazon/Google style)
    r"\bsde[\s\-]*[3456789]\b|\bswe[\s\-]*[3456789]\b|"
    r"\bl[3456789]\b|\blevel[\s\-]*[3456789]\b|"
    # Roman numerals often indicate mid-senior: "Engineer II", "Engineer III"
    r"\bii\b|\biii\b|\biv\b|\bv\b|\bvi\b|"
    # Explicit years in title
    r"\b\d+\s*(?:\+)?\s*years?\s+(?:of\s+)?experience\b|"
    r"\b\d+\s*(?:\+)?\s*years?\s+(?:of\s+)?exp\b"
    r")",
    re.IGNORECASE,
)

# JD-text POSITIVE fresher signals.
FRESHER_JD_POSITIVE = re.compile(
    r"(?:"
    r"\b0[\s\-]*(?:to|\-)?\s*[012]\s*(?:years?|yrs?)\b|"
    r"\b0\s*[\+\-]\s*[12]\s*(?:years?|yrs?)\b|"
    r"\bfresher\b|\bfreshers\b|\bfresh[\s\-]*graduate\b|"
    r"\bentry[\s\-]*level\b|"
    r"\bgraduate[\s\-]*engineer\b|"
    r"\bno[\s\-]+experience[\s\-]+(?:required|necessary|needed)\b|"
    r"\bless[\s\-]+than[\s\-]+[12]\s*(?:years?|yrs?)\b|"
    r"\bup[\s\-]+to[\s\-]+[12]\s*(?:years?|yrs?)\b|"
    r"\b[01]\s*(?:to|\-)\s*[12]\s*(?:years?|yrs?)\b"
    r")",
    re.IGNORECASE,
)

# Years-of-experience extractor. Finds all "X+ years" / "X years of experience"
# patterns in the JD and returns the MINIMUM required years (or None).
# This is the most reliable signal — if the JD says "5+ years", it's NOT fresher.
YEARS_PATTERNS = [
    # "5+ years", "5 + years", "5+ yrs", "5+ years of experience"
    re.compile(r"\b(\d+)\s*\+\s*(?:years?|yrs?)", re.IGNORECASE),
    # "5 years of experience", "5 years experience", "5 years' experience"
    re.compile(r"\b(\d+)\s*(?:years?|yrs?)\s*(?:'?s?\s*)?(?:of\s+)?experience", re.IGNORECASE),
    # "minimum of 5 years", "minimum 5 years", "min 5 years"
    re.compile(r"\b(?:minimum|min(?:imum)?)\s+(?:of\s+)?(\d+)\s*(?:years?|yrs?)", re.IGNORECASE),
    # "at least 5 years", "atleast 5 years"
    re.compile(r"\bat[\s\-]*least\s+(\d+)\s*(?:years?|yrs?)", re.IGNORECASE),
    # "5-7 years", "3-5 years", "5 to 7 years" (take the lower bound)
    re.compile(r"\b(\d+)\s*[\-–—to]+\s*\d+\s*(?:years?|yrs?)", re.IGNORECASE),
    # "require 5 years", "requires 5 years", "required: 5 years"
    re.compile(r"\brequire[sd]?\s*:?\s*(\d+)\s*(?:years?|yrs?)", re.IGNORECASE),
    # "experience: 5 years", "experience - 5 years"
    re.compile(r"\bexperience\s*[:\-]\s*(\d+)\s*(?:years?|yrs?)", re.IGNORECASE),
    # "5 years building", "5 years developing", "5 years in"
    re.compile(r"\b(\d+)\s*(?:years?|yrs?)\s+(?:building|developing|in|of|working)", re.IGNORECASE),
    # "should have 5 years", "must have 5 years"
    re.compile(r"\b(?:should|must)\s+have\s+(\d+)\s*(?:years?|yrs?)", re.IGNORECASE),
    # "5+ years" without space before + (already covered, but be safe)
    re.compile(r"\b(\d+)\+\s*(?:years?|yrs?)", re.IGNORECASE),
]

# JD-text NEGATIVE signals (beyond years extraction).
FRESHER_JD_NEGATIVE = re.compile(
    r"(?:"
    r"\bsenior[\s\-]+level\b|\bsenior[\s\-]+role\b|"
    r"\bextensive[\s\-]+experience\b|"
    r"\bproven[\s\-]+track[\s\-]+record\b|"
    r"\bled[\s\-]+(?:a[\s\-]+)?team\b|"
    r"\bmentored\b|\bmentoring[\s\-]+(?:junior|engineers)\b|"
    r"\boversee\b|\bownership[\s\-]+of\b|"
    r"\bdeep[\s\-]+expertise\b|\bdeep[\s\-]+knowledge\b|"
    r"\byears[\s\-]+building\b|\byears[\s\-]+developing\b"
    r")",
    re.IGNORECASE,
)

# Minimum JD text length to be considered "real" content (not JS-rendered error).
# JS-rendered error pages like "Unable to load this role" are typically <150 chars.
# Real JD pages are usually 500+ chars. We use 200 as the threshold.
MIN_JD_TEXT_LEN = 200


# ---------------------------------------------------------------------------
# Location filter
# ---------------------------------------------------------------------------
# Indian city tokens (case-insensitive). Used to detect on-site India roles.
INDIA_TOKENS = [
    "india", "bangalore", "bengaluru", "mumbai", "delhi", "new delhi",
    "noida", "gurugram", "gurgaon", "pune", "hyderabad", "chennai",
    "kolkata", "ahmedabad", "jaipur", "kochi", "coimbatore", "indore",
    "chandigarh", "lucknow", "bhubaneswar", "trivandrum", "thiruvananthapuram",
    "visakhapatnam", "remote.*india", "india.*remote",
]

# Remote-worldwide indicators. Bare "remote" counts (with reduced confidence).
REMOTE_TOKENS = [
    r"remote\s*worldwide", r"remote\s*global", r"remote\s*\(global\)",
    r"remote\s*\(anywhere\)", r"remote\s*\(worldwide\)",
    r"anywhere", r"work\s+from\s+anywhere",
]

# Restricted-remote indicators -- these are remote but restricted to a region
# that does NOT include India, so we reject them.
RESTRICTED_REMOTE_REJECT = [
    r"remote\s*\(us\b", r"remote\s*\(united\s+states", r"remote\s*\(usa?\)",
    r"remote\s*\(uk\b", r"remote\s*\(united\s+kingdom",
    r"remote\s*\(eu\)", r"remote\s*\(europe\)", r"remote\s*\(emea\)",
    r"remote\s*\(germany\)", r"remote\s*\(france\)",
    r"remote\s*\(canada\)", r"remote\s*\(aus\w*\)",
    r"remote\s*\(latam\)", r"remote\s*\(apac\)",
    r"us\s+only", r"uk\s+only", r"eu\s+only",
    r"north\s+america\s+only", r"europe\s+only",
    r"^\s*u\.?s\.?a?\.?\s*(?:,|/|\|)?\s*(?:remote|hybrid)?\s*$",
    r"^\s*u\.?k\.?\s*(?:,|/|\|)?\s*(?:remote|hybrid)?\s*$",
    r"^\s*eu\.?\s*(?:,|/|\|)?\s*(?:remote|hybrid)?\s*$",
    r"\bus\s*,\s*remote\b", r"\buk\s*,\s*remote\b", r"\beu\s*,\s*remote\b",
    r"\bunited\s+states\s*,\s*remote\b", r"\bunited\s+kingdom\s*,\s*remote\b",
    r"\bremote\s*,\s*u\.?s\.?a?\.?\b",
    r"\bremote\s*,\s*united\s+states\b",
    r"\bremote\s*,\s*u\.?k\.?\b",
    r"\bremote\s*,\s*united\s+kingdom\b",
    r"\bremote\s*,\s*eu\.?\b",
    r"\bremote\s*,\s*europe\b",
    r"\bremote\s*,\s*canada\b",
    r"\bsan\s+francisco\s*/\s*remote\b",
    r"\bnew\s+york\s*/\s*remote\b",
]

# Bare "remote" (no qualifier) — accept but mark confidence medium.
BARE_REMOTE = re.compile(r"\bremote\b", re.IGNORECASE)

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
    careers_page_status: str = ""  # "found" | "not_found" | "error"
    jobs_found: int = 0
    jobs_matched: int = 0
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

    # Fallback: probe common paths. Honor a per-company wall-clock budget so
    # slow/dead sites don't stall the worker pool.
    deadline = time.monotonic() + PER_COMPANY_BUDGET
    for path in COMMON_CAREERS_PATHS:
        if time.monotonic() > deadline:
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
# Stage 5: role + location matching
# ---------------------------------------------------------------------------

INDIA_RE = re.compile("|".join(INDIA_TOKENS), re.IGNORECASE)
RESTRICTED_REMOTE_RE = re.compile("|".join(RESTRICTED_REMOTE_REJECT), re.IGNORECASE)
REMOTE_WORLDWIDE_RE = re.compile("|".join(REMOTE_TOKENS), re.IGNORECASE)


def match_role(title: str) -> Optional[tuple[str, str, str]]:
    """
    Return (category, matched_keyword, confidence) or None.
    confidence: "high" (exact primary keyword) | "medium" (partial) | "low" (catch-all).
    """
    if not title:
        return None
    # Apply exclusions first.
    for ex in TITLE_EXCLUSIONS:
        if re.search(ex, title, re.IGNORECASE):
            return None
    # Walk patterns in order; first hit wins.
    for category, patterns in ROLE_PATTERNS:
        for pat in patterns:
            m = re.search(pat, title, re.IGNORECASE)
            if m:
                kw = m.group(0)
                # Confidence heuristic:
                #  - high: title is short (<=60 chars) and contains the keyword
                #    as a whole word AND category is AI/ML or Backend.
                #  - medium: generic "software engineer" / "engineer" catch-all
                #    OR long titles (likely senior/lead/specialist variants)
                if category == "Full-stack" and pat in (r"engineer", r"software[\s\-]*engineer", r"software[\s\-]*developer"):
                    confidence = "low"
                elif len(title) <= 60:
                    confidence = "high"
                else:
                    confidence = "medium"
                return (category, kw, confidence)
    return None


# US/UK/EU city and country tokens — if ANY of these appear in the location
# OR title, the job is NOT remote-worldwide (it's restricted to that region).
# This prevents jobs like "Machine Learning Scientist • United States • Remote"
# from being accepted as remote-worldwide.
RESTRICTED_REGION_TOKENS = re.compile(
    r"\b("
    # US
    r"united\s+states|u\.?s\.?a?\.?|\bus\b|"
    r"san\s+francisco|new\s+york|seattle|austin|boston|chicago|"
    r"los\s+angeles|denver|atlanta|portland|washington\s+dc|"
    r"palo\s+alto|mountain\s+view|menlo\s+park|redwood\s+city|"
    r"cambridge|irvine|dallas|houston|phoenix|minneapolis|"
    r"san\s+mateo|san\s+jose|santa\s+clara|bellevue|"
    # UK
    r"united\s+kingdom|u\.?k\.?|\blondon\b|"
    # EU
    r"\beu\b|europe|european|"
    r"berlin|paris|amsterdam|dublin|madrid|barcelona|"
    r"munich|frankfurt|vienna|stockholm|copenhagen|zurich|"
    r"brussels|lisbon|prague|warsaw|helsinki|oslo|"
    # Canada
    r"canada|toronto|vancouver|montreal|"
    # Australia
    r"australia|sydney|melbourne|"
    # Other non-India
    r"tel\s+aviv|israel|singapore|tokyo|japan|"
    r"brazil|são\s+paulo|rio\s+de\s+janeiro|"
    r"mexico|mexico\s+city|"
    r"lagos|nigeria|"
    r"dubai|uae|"
    r"berlin|germany|"
    r"france|paris|"
    r"spain|madrid|"
    r"italy|rome|milan|"
    r"netherlands|amsterdam|"
    r"sweden|stockholm|"
    r"norway|oslo|"
    r"denmark|copenhagen|"
    r"finland|helsinki|"
    r"poland|warsaw|"
    r"ireland|dublin|"
    r"portugal|lisbon|"
    r"switzerland|zurich|"
    r"austria|vienna|"
    r"belgium|brussels|"
    r"czech|prague|"
    r"hungary|budapest|"
    r"romania|bucharest|"
    r"greece|athens|"
    r"turkey|istanbul|"
    r"russia|moscow|"
    r"ukraine|kyiv|"
    r"china|beijing|shanghai|shenzhen|"
    r"korea|seoul|"
    r"hong\s+kong|"
    r"taiwan|taipei|"
    r"thailand|bangkok|"
    r"vietnam|hanoi|ho\s+chi\s+minh|"
    r"indonesia|jakarta|"
    r"malaysia|kuala\s+lumpur|"
    r"philippines|manila|"
    r"argentina|buenos\s+aires|"
    r"chile|santiago|"
    r"colombia|bogota|"
    r"peru|lima|"
    r"egypt|cairo|"
    r"south\s+africa|johannesburg|cape\s+town|"
    r"kenya|nairobi|"
    r"morocco|casablanca|"
    r"saudi\s+arabia|riyadh|"
    r"qatar|doha|"
    r"kuwait|"
    r"oman|muscat|"
    r"bahrain|"
    r"jordan|amman|"
    r"lebanon|beirut|"
    r"new\s+zealand|auckland|"
    r"north\s+america|"
    r"south\s+america|"
    r"latam|"
    r"apac"  # APAC is ambiguous but usually excludes India-specific remote
    r")\b",
    re.IGNORECASE,
)


def match_location(location: str, title: str = "") -> Optional[tuple[str, str]]:
    """
    Return (basis, confidence) or None.
      basis: "india" | "remote_worldwide" | "remote_unqualified"
      confidence: "high" | "medium"

    Checks BOTH the location string AND the title for restricted-region tokens.
    This prevents jobs like "ML Scientist • United States • Remote" from being
    accepted as remote-worldwide when only the location field says "Remote".
    """
    if not location and not title:
        return None

    # Combine location + title for a comprehensive check.
    combined = f"{location} {title}".strip()
    if not combined:
        return None

    # 1. Reject restricted-remote patterns first (e.g., "Remote (US)").
    if RESTRICTED_REMOTE_RE.search(combined):
        return None

    # 2. Reject if ANY restricted-region token appears in the combined text.
    # This catches "United States" in the title even when location says "Remote".
    if RESTRICTED_REGION_TOKENS.search(combined):
        return None

    # 3. India token? (check both location and title — India cities in title
    # are a strong signal, e.g., "Backend Engineer Bengaluru")
    if INDIA_RE.search(combined):
        return ("india", "high")

    # 4. Remote-worldwide explicit?
    if REMOTE_WORLDWIDE_RE.search(combined):
        return ("remote_worldwide", "high")

    # 5. Bare "remote" (no qualifier) — only accept if NO restricted token
    # was found (already checked above). This is "remote from everywhere".
    if BARE_REMOTE.search(combined):
        return ("remote_unqualified", "medium")

    return None


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


def extract_min_years(jd_text: str) -> Optional[int]:
    """Extract the minimum required years of experience from JD text.
    Returns the minimum years found, or None if no pattern matched.
    Examples:
      "5+ years of experience" → 5
      "3-5 years" → 3
      "minimum 7 years" → 7
      "at least 4 years" → 4
    """
    if not jd_text:
        return None
    years_found = []
    for pat in YEARS_PATTERNS:
        for m in pat.finditer(jd_text):
            try:
                y = int(m.group(1))
                years_found.append(y)
            except (ValueError, IndexError):
                continue
    if not years_found:
        return None
    # Return the minimum (the lowest bar the candidate must clear).
    return min(years_found)


def is_fresher_role(title: str, jd_text: str = "") -> tuple[bool, str, str]:
    """
    Determine whether a job is fresher-eligible (0-2 years experience).

    Uses a balanced multi-signal approach:
      1. Title POSITIVE (Junior/Fresher/Intern/SDE-1) AND NOT title NEGATIVE → fresher (explicit)
      2. Title NEGATIVE (Senior/Staff/Principal/Sr./II/III) → NOT fresher
      3. JD has explicit years → if min_years >= 3, NOT fresher; if <= 2, fresher
      4. JD has POSITIVE signals ("0-2 years", "fresher", "entry level") → fresher
      5. JD has NEGATIVE signals ("extensive experience", "led a team") → NOT fresher
      6. JD fetched but no signal → ACCEPT (most JDs without years requirement are open to freshers)
      7. JD unavailable (JS-rendered) + ambiguous title → ACCEPT (better to include than exclude)

    This is BALANCED: rejects jobs that explicitly demand 3+ years or have senior
    titles, but accepts jobs where we can't verify (to avoid missing fresher roles).

    Returns (is_fresher, basis, level):
      is_fresher: True if the role is suitable for a fresher.
      basis: explanation string
      level: "explicit" | "implicit" | "rejected"
    """
    if not title:
        return (False, "rejected_no_signal", "rejected")

    title_pos = bool(FRESHER_TITLE_POSITIVE.search(title))
    title_neg = bool(FRESHER_TITLE_NEGATIVE.search(title))

    # 1. Title says BOTH fresher AND senior (e.g., "Senior Junior Engineer")
    #    → senior wins, reject.
    if title_pos and title_neg:
        return (False, "rejected_seniority_title", "rejected")

    # 2. Title POSITIVE only → fresher (explicit).
    if title_pos:
        return (True, "title_positive", "explicit")

    # 3. Title NEGATIVE only → NOT fresher.
    if title_neg:
        return (False, "rejected_seniority_title", "rejected")

    # 4. Title is ambiguous (no fresher/seniority keyword). Check JD text.
    # Only use JD text if it's long enough to be real content (not a
    # JS-rendered error page like "Unable to load this role").
    has_real_jd = bool(jd_text) and len(jd_text) >= MIN_JD_TEXT_LEN

    if has_real_jd:
        # 4a. Extract explicit years of experience from JD.
        min_years = extract_min_years(jd_text)
        if min_years is not None:
            if min_years >= 3:
                return (False, f"rejected_jd_{min_years}years", "rejected")
            else:
                return (True, f"jd_{min_years}years", "implicit")

        # 4b. JD POSITIVE signal → fresher (implicit, verified by JD).
        if FRESHER_JD_POSITIVE.search(jd_text):
            return (True, "jd_positive", "implicit")

        # 4c. JD NEGATIVE signal → NOT fresher.
        if FRESHER_JD_NEGATIVE.search(jd_text):
            return (False, "rejected_seniority_jd", "rejected")

        # 4d. JD available but no signal either way.
        # ACCEPT as implicit fresher — most JDs that don't mention years
        # are open to freshers. Better to include and let the user decide.
        return (True, "jd_no_years_mentioned", "implicit")
    else:
        # 5. No real JD text available (JS-rendered or fetch failed).
        # Title is ambiguous (no seniority keyword).
        # ACCEPT as implicit fresher — better to include potential fresher
        # roles than to exclude them. The user can manually verify via
        # the Apply URL.
        return (True, "no_jd_ambiguous_title", "implicit")


# ---------------------------------------------------------------------------
# Stage 6: per-company processing
# ---------------------------------------------------------------------------

def process_company(
    news: NewsItem,
    session: requests.Session,
    fresher_only: bool = False,
    python_stack: bool = True,
) -> tuple[list[JobMatch], CompanyRecord, Optional[ErrorRecord]]:
    """Process one company: find careers, fetch jobs, filter matches.
    If fresher_only=True, additionally filter to fresher-eligible roles
    (0-2 years experience), verifying via JD-text fetch.
    If python_stack=True (default), require the title OR JD to mention a
    Python-stack technology (Python/Django/Flask/FastAPI/PyTorch/etc.).

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
    jd_deadline = min(deadline, time.monotonic() + JD_FETCH_BUDGET)
    jd_fetched = 0
    matches: list[JobMatch] = []
    for job in jobs:
        # Check overall company deadline.
        if time.monotonic() > deadline:
            break
        role = match_role(job.title)
        if role is None:
            continue
        category, kw, role_conf = role
        loc = match_location(job.location, job.title)
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
                job.title, jd_text
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
            elif len(jd_text) < 300:
                # JD text is too short (JS-rendered page, error message, or
                # fetch failed). Be LENIENT: Python is the dominant language
                # for backend, AI/ML, data, and internship roles at SaaS
                # startups — especially in India. Accept these and let the
                # user verify manually via the Apply URL.
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
        ))
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


def write_workbook(
    out_path: str,
    matches: list[JobMatch],
    companies: list[CompanyRecord],
    errors: list[ErrorRecord],
    run_meta: dict,
) -> None:
    wb = Workbook()
    wb.properties.creator = "Z.ai"

    # --- Sheet 1: Matches ---
    ws1 = wb.active
    ws1.title = "Matches"
    headers1 = [
        "Company", "Job Title", "Category", "Confidence", "Matched Keyword",
        "Location", "Location Basis", "Fresher Basis", "Fresher Level",
        "Apply URL", "Job URL",
        "Company Website", "Careers Page", "Funding Round", "Funding Date",
        "Software Category", "News Headline", "News URL",
    ]
    rows1 = []
    for m in matches:
        rows1.append([
            m.company_name, m.job_title, m.match_category, m.confidence,
            m.matched_keyword, m.location, m.location_basis,
            m.fresher_basis, m.fresher_level,
            m.apply_url, m.job_url,
            m.company_website, m.careers_page_url, m.funding_round,
            m.funding_date, m.software_category, m.news_headline, m.news_url,
        ])
    # Sort: high → medium → low, then by company
    conf_order = {"high": 0, "medium": 1, "low": 2}
    rows1.sort(key=lambda r: (conf_order.get(str(r[3]).lower(), 9), r[0].lower()))
    _write_sheet(
        ws1, headers1, rows1,
        col_widths=[22, 42, 12, 12, 20, 28, 18, 22, 14, 50, 50, 28, 38, 14, 14, 22, 50, 50],
        color_col=3,  # Confidence column (0-indexed) → cell column 4
        color_map={"high": HIGH_FILL, "medium": MED_FILL, "low": LOW_FILL},
    )

    # --- Sheet 2: All Companies ---
    ws2 = wb.create_sheet("All Companies")
    headers2 = [
        "Company", "Website", "Funding Round", "Funding Date", "Lead Investor",
        "Software Category", "Careers Page", "Careers Status",
        "Jobs Found", "Jobs Matched", "News Headline", "News URL", "Error",
    ]
    rows2 = []
    for c in companies:
        rows2.append([
            c.company_name, c.company_website, c.funding_round, c.funding_date,
            c.lead_investor, c.software_category, c.careers_page_url,
            c.careers_page_status, c.jobs_found, c.jobs_matched,
            c.news_headline, c.news_url, c.error,
        ])
    rows2.sort(key=lambda r: r[0].lower())
    _write_sheet(
        ws2, headers2, rows2,
        col_widths=[22, 30, 16, 14, 22, 24, 38, 14, 12, 12, 50, 50, 40],
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
        ["CLI Args", json.dumps(run_meta.get("args", {}))],
    ]
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
    p.add_argument("--output-dir", default="/home/mikeysama/products/saasnews-job-finder/download",
                   help="Output directory for xlsx.")
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
    import os
    return os.path.join(output_dir, "checkpoint.jsonl")


def _load_checkpoint(output_dir: str) -> dict[str, dict]:
    """Load prior checkpoint. Returns dict keyed by news_url → record.
    Each record has: company, matches (list), company_rec (dict), error (dict|None).
    """
    import os
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
    import os
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
) -> str:
    """Rebuild the xlsx workbook from the in-memory checkpoint.
    Called periodically during the run AND at the end.
    """
    import os
    all_matches: list[JobMatch] = []
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
            error=cr.get("error", ""),
        )
        all_companies.append(company_rec)
        jobs_parsed_total += company_rec.jobs_found
        # Reconstruct matches
        for m_dict in rec.get("matches", []):
            all_matches.append(JobMatch(
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
            ))
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
        "news_total": args_dict.get("news_pages", 0) * 12,  # approx
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
    # Write to temp file then rename (atomic).
    tmp_path = out_path + ".tmp"
    try:
        write_workbook(tmp_path, all_matches, all_companies, all_errors, run_meta)
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

    import os
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

    if not pending_urls:
        log.info("Nothing to do — all articles already processed.")
        _write_xlsx_from_checkpoint(args.output_dir, checkpoint, vars(args))
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

    def _worker(ni: NewsItem):
        s = make_session()
        return process_company(ni, s, fresher_only=fresher_only,
                               python_stack=python_stack)

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
                    _write_xlsx_from_checkpoint(args.output_dir, checkpoint, args_dict)
                    processed_since_flush = 0

    # ---- Final xlsx write ----
    out_path = _write_xlsx_from_checkpoint(args.output_dir, checkpoint, args_dict)
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    archive_path = os.path.join(args.output_dir, f"saasnews_jobs_{ts}.xlsx")
    try:
        import shutil
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


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
