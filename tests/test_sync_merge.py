"""Google Sheets sync merge logic (pure seam — no network).

Seam under test: sync_to_google_sheets.merge_sheet_rows — combined feed must
read newest-first, user-owned columns (Applied/Feedback/Notes) must survive
re-syncs keyed by Job URL, and incremental mode keeps sheet-only history.
"""
from datetime import datetime, timedelta, timezone

from sync_to_google_sheets import merge_sheet_rows

HEADERS = ["Company", "Job Title", "Job URL", "First Seen",
           "Applied", "Feedback", "Notes"]


def t(hour):
    base = datetime(2026, 8, 24, tzinfo=timezone.utc)
    return (base + timedelta(hours=hour)).isoformat()


def row(url, first_seen="", applied="", feedback="", notes="",
        company="Co", title="Python Backend Engineer"):
    return [company, title, url, first_seen, applied, feedback, notes]


def test_new_jobs_sorted_above_older_newest_first():
    existing = [row("https://x/old", first_seen=t(1))]
    incoming = [
        row("https://x/newest", first_seen=t(9)),
        row("https://x/mid", first_seen=t(5)),
    ]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    urls = [r[HEADERS.index("Job URL")] for r in merged]
    assert urls == ["https://x/newest", "https://x/mid", "https://x/old"]


def test_user_columns_survive_resync():
    """The xlsx regenerates Applied/Feedback as blanks; a re-sync must NOT
    wipe what the user marked in the sheet."""
    existing = [
        row("https://x/j1", first_seen=t(1), applied="Yes",
            feedback="not_fresher", notes="wants 5 years"),
    ]
    incoming = [row("https://x/j1", first_seen=t(1))]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    assert len(merged) == 1
    r = dict(zip(HEADERS, merged[0]))
    assert r["Applied"] == "Yes"
    assert r["Feedback"] == "not_fresher"
    assert r["Notes"] == "wants 5 years"


def test_system_columns_refreshed_from_incoming():
    existing = [row("https://x/j1", company="OLD NAME")]
    incoming = [row("https://x/j1", company="New Name", first_seen=t(3))]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    assert dict(zip(HEADERS, merged[0]))["Company"] == "New Name"


def test_incremental_keeps_sheet_only_history():
    existing = [row("https://x/archived", first_seen=t(1), applied="No")]
    incoming = [row("https://x/fresh", first_seen=t(9))]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    urls = {r[HEADERS.index("Job URL")] for r in merged}
    assert urls == {"https://x/archived", "https://x/fresh"}


def test_full_sync_drops_sheet_only_rows():
    existing = [row("https://x/archived")]
    incoming = [row("https://x/fresh", first_seen=t(9))]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=False)
    urls = [r[HEADERS.index("Job URL")] for r in merged]
    assert urls == ["https://x/fresh"]


def test_rows_without_timestamp_sink_to_bottom_stably():
    """Timestamp-less rows sink below timestamped ones — except Applied rows,
    which sink below everything (the top of the sheet stays the fresh list)."""
    existing = [row("https://x/a-old", applied="Yes"),
                row("https://x/b-old")]
    incoming = [row("https://x/fresh", first_seen=t(9))]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    urls = [r[HEADERS.index("Job URL")] for r in merged]
    assert urls == ["https://x/fresh", "https://x/b-old", "https://x/a-old"]


def test_short_legacy_rows_padded_and_deduped_by_url():
    """A sheet written before the tracking columns exists still merges."""
    existing = [["Co", "Python Dev", "https://x/j1"]]  # no First Seen cols
    incoming = [row("https://x/j1", first_seen=t(2)),
                row("https://x/j1", first_seen=t(2))]   # dup inside incoming
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    assert len(merged) == 1
    r = dict(zip(HEADERS, merged[0]))
    assert r["First Seen"] == t(2)
    assert len(r) == len(HEADERS)


def test_empty_sheet_gets_incoming_sorted():
    merged = merge_sheet_rows(
        HEADERS, [],
        [row("https://x/a", first_seen=t(1)), row("https://x/b", first_seen=t(5))],
        keep_sheet_only=True,
    )
    urls = [r[HEADERS.index("Job URL")] for r in merged]
    assert urls == ["https://x/b", "https://x/a"]


def test_applied_jobs_sink_below_active_feed():
    """The top of the sheet is the fresh list: jobs marked Applied=Yes move to
    the very bottom (sorted among themselves), regardless of timestamps."""
    existing = [
        row("https://x/applied-old", first_seen=t(1), applied="Yes"),
        row("https://x/mid", first_seen=t(5)),
    ]
    incoming = [
        row("https://x/fresh", first_seen=t(9)),
        row("https://x/applied-new", first_seen=t(10), applied="Yes"),
    ]
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    urls = [r[HEADERS.index("Job URL")] for r in merged]
    assert urls == ["https://x/fresh", "https://x/mid",
                    "https://x/applied-new", "https://x/applied-old"]


def test_applied_marked_only_in_sheet_still_sinks():
    """User marks Applied=Yes directly in the sheet; next incremental sync
    keeps it below the active feed."""
    existing = [
        row("https://x/j1", first_seen=t(9)),
        row("https://x/j2", first_seen=t(8), applied="Yes"),
    ]
    incoming = []
    merged = merge_sheet_rows(HEADERS, existing, incoming, keep_sheet_only=True)
    urls = [r[HEADERS.index("Job URL")] for r in merged]
    assert urls == ["https://x/j1", "https://x/j2"]
