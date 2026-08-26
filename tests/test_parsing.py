"""Funding-details article parsing across all four historical site formats.

Ported 1:1 from the original tests/run_tests.py custom runner.
"""
from saasnews_scraper import (
    FD_COMPANY_RE,
    FD_COMPANY_SITE_RE,
    FD_CATEGORY_RE,
    FD_DATE_RE,
    FD_LEAD_RE,
    FD_ROUND_RE,
    strip_ref,
)

# Format 1 (newest): plain text, <br> separators
FMT1 = (
    'Company: Azraq<br>Round: Pre-seed<br>Funding Date: July 10, 2026'
    '<br>Lead Investor: A-typical Ventures<br>'
    'Company Website: <a href="https://azraq.ai/?ref=thesaasnews.com">https://azraq.ai</a><br>'
    "Software Category: FinTech"
)

# Format 2 (mid): plain text, &nbsp;</p><p> separators
FMT2 = (
    "Company: Artificially Intelligent Inc.&nbsp;&nbsp;</p>"
    "<p>Raised: $100K&nbsp;&nbsp;</p><p>Round: Angel&nbsp;&nbsp;</p>"
    "<p>Funding Month: January 2024&nbsp; &nbsp;&nbsp;&nbsp;</p>"
    "<p>Lead Investors: Payment Ventures&nbsp;&nbsp;&nbsp;</p>"
    '<p>Company Website:&nbsp;<a href="https://www.legalyze.ai/?ref=thesaasnews.com">https://www.legalyze.ai/</a>&nbsp;&nbsp;&nbsp;</p>'
    "<p>Software Category: Legal&nbsp;&nbsp;&nbsp;</p>"
)

# Format 3 (older): <strong> labels, plain text URL
FMT3 = (
    "<strong>Company:&nbsp;</strong>ToBox Ventures Pvt. Ltd."
    "<strong>Round:&nbsp;</strong>pre-Series A"
    "<strong>Funding Month:&nbsp;</strong>December 2021"
    "<strong>Lead Investors:&nbsp;</strong>GRM Foodkraft"
    "<strong>Company Website:&nbsp;</strong>https://www.gokhana.com/"
    "<strong>Software Category:&nbsp;</strong>SaaS"
)

# Format 3b: <strong> labels, <a> link URL
FMT3B = (
    "<strong>Company Website:&nbsp;</strong>"
    '<a href="https://example.com/?ref=thesaasnews.com">https://example.com</a>'
)


def test_f1_company():
    m = FD_COMPANY_RE.search(FMT1)
    assert m and m.group(1).strip() == "Azraq"


def test_f1_round():
    m = FD_ROUND_RE.search(FMT1)
    assert m and m.group(1).strip() == "Pre-seed"


def test_f1_date():
    m = FD_DATE_RE.search(FMT1)
    assert m and m.group(1).strip() == "July 10, 2026"


def test_f1_website_href():
    m = FD_COMPANY_SITE_RE.search(FMT1)
    assert m and m.group(2) == "https://azraq.ai/?ref=thesaasnews.com"


def test_f2_company():
    m = FD_COMPANY_RE.search(FMT2)
    assert m and m.group(1).strip() == "Artificially Intelligent Inc."


def test_f2_round():
    m = FD_ROUND_RE.search(FMT2)
    assert m and m.group(1).strip() == "Angel"


def test_f2_date_funding_month():
    m = FD_DATE_RE.search(FMT2)
    assert m and m.group(1).strip() == "January 2024"


def test_f2_lead_investors_plural():
    m = FD_LEAD_RE.search(FMT2)
    assert m and m.group(1).strip() == "Payment Ventures"


def test_f3_company_strong():
    m = FD_COMPANY_RE.search(FMT3)
    assert m and m.group(1).strip() == "ToBox Ventures Pvt. Ltd."


def test_f3_round_strong():
    m = FD_ROUND_RE.search(FMT3)
    assert m and m.group(1).strip() == "pre-Series A"


def test_f3_website_plain_text_url():
    assert FD_COMPANY_SITE_RE.search(FMT3) is not None


def test_f3b_website_strong_plus_href():
    assert FD_COMPANY_SITE_RE.search(FMT3B) is not None


def test_strip_ref_removes_query():
    assert strip_ref("https://example.com/page?ref=thesaasnews.com") == \
        "https://example.com/page"


def test_strip_ref_empty():
    assert strip_ref("") == ""
