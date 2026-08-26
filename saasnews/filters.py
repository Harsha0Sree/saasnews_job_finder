"""Job-evaluation filters — the single source of truth for "is this job
acceptable?".

Deep module: the interface is four pure functions

    match_role(title)                -> (category, keyword, confidence) | None
    match_location(location, title)  -> (basis, confidence) | None
    is_fresher_role(title, jd_text)  -> (bool, basis, level)
    extract_min_years(jd_text)       -> int | None

plus the regex/token tables they are built from. Both the scraper pipeline and
combine_results.py consume this module; neither re-implements the policy.

No network, no filesystem, no I/O of any kind — everything here is pure and
deterministically testable.
"""
from __future__ import annotations

import re
from typing import Optional

# ---------------------------------------------------------------------------
# Role keywords -- grouped by category. PYTHON-STACK ONLY.
# Order matters: earlier patterns take precedence for category assignment.
# We accept AI/ML, Backend (Python/Django/Flask/FastAPI), Full-stack, Data,
# and Internship variants of any of the above.
# ---------------------------------------------------------------------------

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


def match_role(title: str,
               extra_exclusions: Optional[list] = None) -> Optional[tuple[str, str, str]]:
    """
    Return (category, matched_keyword, confidence) or None.
    confidence: "high" (exact primary keyword) | "medium" (partial) | "low" (catch-all).

    `extra_exclusions`: compiled regexes learned from user feedback
    (saasnews.feedback LEARNING_KINDS → "title_exclusion"); checked alongside
    the built-in TITLE_EXCLUSIONS.
    """
    if not title:
        return None
    # Apply exclusions first.
    for ex in TITLE_EXCLUSIONS:
        if re.search(ex, title, re.IGNORECASE):
            return None
    for ex in extra_exclusions or []:
        if ex.search(title):
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


def normalize_location(value: str) -> str:
    """Canonical form for exact location comparison (case/spacing-insensitive).
    Used by the learned-feedback blocklist (saasnews.feedback)."""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


INDIA_RE = re.compile("|".join(INDIA_TOKENS), re.IGNORECASE)
RESTRICTED_REMOTE_RE = re.compile("|".join(RESTRICTED_REMOTE_REJECT), re.IGNORECASE)
REMOTE_WORLDWIDE_RE = re.compile("|".join(REMOTE_TOKENS), re.IGNORECASE)


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


def match_location(location: str, title: str = "",
                   reject_locations: Optional[set[str]] = None,
                   extra_region_re: Optional[re.Pattern] = None
                   ) -> Optional[tuple[str, str]]:
    """
    Return (basis, confidence) or None.
      basis: "india" | "remote_worldwide" | "remote_unqualified"
      confidence: "high" | "medium"

    Checks BOTH the location string AND the title for restricted-region tokens.
    This prevents jobs like "ML Scientist • United States • Remote" from being
    accepted as remote-worldwide when only the location field says "Remote".

    `reject_locations`: learned exact-match blocklist (normalized strings,
    see saasnews.feedback) harvested from user feedback; checked first so
    known-bad locations never re-enter the feed.
    `extra_region_re`: compiled regex over learned region tokens
    ("location_token" learnings) — generalizes exact feedback to any string
    containing that token, in either the location or the title.
    """
    if not location and not title:
        return None

    # Combine location + title for a comprehensive check.
    combined = f"{location} {title}".strip()
    if not combined:
        return None

    # 0. Learned rejections (closed feedback loop): exact normalized match.
    if reject_locations and normalize_location(location) in reject_locations:
        return None

    # 1. Reject restricted-remote patterns first (e.g., "Remote (US)").
    if RESTRICTED_REMOTE_RE.search(combined):
        return None

    # 2. Reject if ANY restricted-region token appears in the combined text.
    # This catches "United States" in the title even when location says "Remote".
    if RESTRICTED_REGION_TOKENS.search(combined):
        return None

    # 2b. Learned region tokens (diagnosed from bad_location feedback).
    if extra_region_re is not None and extra_region_re.search(combined):
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
# Fresher / 0-experience filter
# ---------------------------------------------------------------------------
# A job is "fresher-eligible" if the role requires 0-2 years of experience.
#
# We use a MULTI-SIGNAL approach for maximum reliability:
#   1. Title-level POSITIVE signal (Junior/Fresher/Intern/SDE-1/L1) → fresher
#   2. Title-level NEGATIVE signal (Senior/Staff/Principal/Lead/Sr./II/III) → NOT fresher
#   3. JD-text explicit years extraction (parse "X+ years" → if X >= 3, NOT fresher)
#   4. JD-text POSITIVE signals ("0-2 years", "fresher", "entry level")
#   5. If JD unavailable (JS-rendered) AND title is ambiguous → ACCEPT (balanced)
#
# This is BALANCED: rejects jobs that explicitly demand 3+ years or have senior
# titles, but accepts jobs where we can't verify (to avoid missing fresher roles).

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


def is_fresher_role(title: str, jd_text: str = "",
                    extra_title_negatives: Optional[list] = None,
                    extra_jd_negatives: Optional[list] = None
                    ) -> tuple[bool, str, str]:
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

    `extra_title_negatives` / `extra_jd_negatives`: compiled regexes learned
    from diagnosed feedback (saasnews.feedback LEARNING_KINDS →
    "title_negative" / "jd_negative_pattern"); OR-ed with the built-in tables.

    Returns (is_fresher, basis, level):
      is_fresher: True if the role is suitable for a fresher.
      basis: explanation string
      level: "explicit" | "implicit" | "rejected"
    """
    if not title:
        return (False, "rejected_no_signal", "rejected")

    title_pos = bool(FRESHER_TITLE_POSITIVE.search(title))
    title_neg = (bool(FRESHER_TITLE_NEGATIVE.search(title))
                 or any(p.search(title) for p in extra_title_negatives or []))

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
        if FRESHER_JD_NEGATIVE.search(jd_text) or any(
                p.search(jd_text) for p in extra_jd_negatives or []):
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
