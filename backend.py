"""
Backend for the Trip Support Intake app -- SQLite-backed request store
plus the bridge into parsing_engine.py.

Kept separate from app.py (the Streamlit UI) so the database logic can
be exercised directly in a plain Python process/tests, without needing
a live Streamlit runtime context.
"""

import json
import os
import re
import sqlite3
import tempfile
import uuid
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import parsing_engine as pe

DB_PATH = Path(__file__).parent / "requests.db"

STATUS_SUBMITTED = "Submitted"
STATUS_PROCESSING = "Processing"
STATUS_COMPLETE = "Complete"
STATUS_FAILED = "Failed"

# No placeholder names are seeded -- the salesperson list starts empty
# and is built entirely from what's added via the UI's "Manage
# salesperson list" controls.
DEFAULT_SALESPEOPLE: list[str] = []

# Weekly Report has its own roster, separate from the salespeople list --
# the people submitting weekly reports aren't the same group as the
# Trip Preference Parsing salespeople. No placeholder names are seeded,
# same as salespeople.
DEFAULT_REPORT_TEAM_MEMBERS: list[str] = []

# Weekly Report categories -- (db column name, label shown on the form
# and in the compiled report).
WEEKLY_REPORT_CATEGORIES = [
    ("highlights", "Highlights / Major Accomplishments / Wins"),
    ("opportunities", "Opportunities / Sales Engagements"),
    ("watch_items", "Watch Items / Help Needed"),
    ("upcoming_events", "Upcoming Events"),
    ("other", "Other"),
]


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------
def get_conn(db_path: Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS requests (
            id TEXT PRIMARY KEY,
            created_at TEXT,
            updated_at TEXT,
            salesperson TEXT,
            jeppesen_acct_nbr TEXT,
            payment_method TEXT,
            amount REAL,
            notes TEXT,
            onboarding_filename TEXT,
            onboarding_bytes BLOB,
            status TEXT,
            result_filename TEXT,
            result_bytes BLOB,
            error_message TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS salespeople (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL COLLATE NOCASE
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS weekly_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            submitter TEXT NOT NULL,
            period TEXT NOT NULL,
            highlights TEXT,
            opportunities TEXT,
            watch_items TEXT,
            upcoming_events TEXT,
            other TEXT,
            submitted_at TEXT,
            updated_at TEXT,
            UNIQUE(submitter, period)
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS report_recipients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL COLLATE NOCASE
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS report_team_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL COLLATE NOCASE
        )"""
    )
    _migrate_legacy_period_format(conn)
    return conn


def _migrate_legacy_period_format(conn: sqlite3.Connection) -> None:
    """One-time, idempotent data migration.

    weekly_reports.period used to store an ISO-week code (e.g.
    "2026-W36"). The canonical period identifier is now that period's
    Monday as an ISO date (e.g. "2026-08-31") -- see the period helpers
    below (period_id_from_monday / monday_from_period_id). This avoids
    ever comparing or storing periods as a formatted display label,
    keeps the identifier trivially sortable and unambiguous across
    year boundaries, and matches the "Monday start date is canonical"
    rule directly instead of only in spirit.

    Safe to call on every connection: it only rewrites rows still in
    the old "YYYY-Www" format (matched with a regex, not just "looks
    like it has a W in it") and does nothing once they're migrated.
    No rows are dropped; only the `period` value is reformatted, so
    submitted_at/updated_at/created content are untouched.
    """
    legacy_pattern = re.compile(r"^\d{4}-W\d{2}$")
    rows = conn.execute("SELECT DISTINCT period FROM weekly_reports").fetchall()
    for (old_period,) in rows:
        if not old_period or not legacy_pattern.match(old_period):
            continue
        try:
            year_str, week_str = old_period.split("-W")
            monday = date.fromisocalendar(int(year_str), int(week_str), 1)
        except (ValueError, IndexError):
            continue
        new_period = monday.isoformat()
        if new_period == old_period:
            continue
        try:
            conn.execute(
                "UPDATE weekly_reports SET period = ? WHERE period = ?",
                (new_period, old_period),
            )
        except sqlite3.IntegrityError:
            # A row for (submitter, new_period) already exists for some
            # submitter -- extremely unlikely (the new format was never
            # writable before this migration existed), but fail safe by
            # leaving the legacy row in place rather than losing data.
            continue
    conn.commit()


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# Salesperson list -- editable from the UI (Add/Edit/Delete), stored
# alongside requests.db so changes persist and are shared by everyone
# using this deployment.
# ---------------------------------------------------------------------------
def list_salespeople(db_path: Path = DB_PATH) -> list[str]:
    conn = get_conn(db_path)
    try:
        rows = conn.execute("SELECT name FROM salespeople ORDER BY name COLLATE NOCASE").fetchall()
        if not rows:
            # First run: seed the table from the default list so existing
            # deployments don't suddenly show an empty dropdown.
            for name in DEFAULT_SALESPEOPLE:
                conn.execute("INSERT OR IGNORE INTO salespeople (name) VALUES (?)", (name,))
            conn.commit()
            rows = conn.execute("SELECT name FROM salespeople ORDER BY name COLLATE NOCASE").fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def add_salesperson(name: str, db_path: Path = DB_PATH) -> None:
    name = (name or "").strip()
    if not name:
        raise ValueError("Name can't be blank.")
    conn = get_conn(db_path)
    try:
        try:
            conn.execute("INSERT INTO salespeople (name) VALUES (?)", (name,))
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"'{name}' is already in the list.")
    finally:
        conn.close()


def update_salesperson(old_name: str, new_name: str, db_path: Path = DB_PATH) -> None:
    new_name = (new_name or "").strip()
    if not new_name:
        raise ValueError("Name can't be blank.")
    conn = get_conn(db_path)
    try:
        try:
            cur = conn.execute("UPDATE salespeople SET name = ? WHERE name = ?", (new_name, old_name))
            if cur.rowcount == 0:
                raise ValueError(f"'{old_name}' was not found (it may have just been changed elsewhere).")
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"'{new_name}' is already in the list.")
    finally:
        conn.close()


def delete_salesperson(name: str, db_path: Path = DB_PATH) -> None:
    conn = get_conn(db_path)
    try:
        conn.execute("DELETE FROM salespeople WHERE name = ?", (name,))
        conn.commit()
    finally:
        conn.close()


def export_salespeople_json(db_path: Path = DB_PATH) -> str:
    """Serializes the current salesperson list to a JSON string, for the
    UI's "Download backup" button. Streamlit Community Cloud's local
    filesystem isn't guaranteed to survive a redeploy, so this file is
    the safety net: download it every so often, and restore it below if
    the list ever comes back empty unexpectedly."""
    return json.dumps({"salespeople": list_salespeople(db_path)}, indent=2)


def import_salespeople_json(json_text: str, db_path: Path = DB_PATH) -> tuple[int, int]:
    """Restores salespeople from a previously downloaded backup. Adds any
    name not already present; silently skips ones that already exist
    (case-insensitive) so restoring is safe to run more than once.
    Returns (added_count, skipped_count)."""
    try:
        data = json.loads(json_text)
        names = data["salespeople"]
        if not isinstance(names, list):
            raise TypeError
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        raise ValueError("That doesn't look like a salesperson backup file.") from e

    existing = {n.lower() for n in list_salespeople(db_path)}
    added = 0
    skipped = 0
    for name in names:
        name = (name or "").strip()
        if not name:
            continue
        if name.lower() in existing:
            skipped += 1
            continue
        add_salesperson(name, db_path)
        existing.add(name.lower())
        added += 1
    return added, skipped


# ---------------------------------------------------------------------------
# Weekly Report team roster -- editable from the UI (Add/Rename/Delete),
# same pattern as the salesperson list, but a separate table: the people
# submitting weekly reports are a different group than the Trip
# Preference Parsing salespeople.
# ---------------------------------------------------------------------------
def list_report_team_members(db_path: Path = DB_PATH) -> list[str]:
    conn = get_conn(db_path)
    try:
        rows = conn.execute("SELECT name FROM report_team_members ORDER BY name COLLATE NOCASE").fetchall()
        if not rows:
            for name in DEFAULT_REPORT_TEAM_MEMBERS:
                conn.execute("INSERT OR IGNORE INTO report_team_members (name) VALUES (?)", (name,))
            conn.commit()
            rows = conn.execute("SELECT name FROM report_team_members ORDER BY name COLLATE NOCASE").fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def add_report_team_member(name: str, db_path: Path = DB_PATH) -> None:
    name = (name or "").strip()
    if not name:
        raise ValueError("Name can't be blank.")
    conn = get_conn(db_path)
    try:
        try:
            conn.execute("INSERT INTO report_team_members (name) VALUES (?)", (name,))
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"'{name}' is already in the list.")
    finally:
        conn.close()


def update_report_team_member(old_name: str, new_name: str, db_path: Path = DB_PATH) -> None:
    new_name = (new_name or "").strip()
    if not new_name:
        raise ValueError("Name can't be blank.")
    conn = get_conn(db_path)
    try:
        try:
            cur = conn.execute("UPDATE report_team_members SET name = ? WHERE name = ?", (new_name, old_name))
            if cur.rowcount == 0:
                raise ValueError(f"'{old_name}' was not found (it may have just been changed elsewhere).")
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"'{new_name}' is already in the list.")
    finally:
        conn.close()


def delete_report_team_member(name: str, db_path: Path = DB_PATH) -> None:
    conn = get_conn(db_path)
    try:
        conn.execute("DELETE FROM report_team_members WHERE name = ?", (name,))
        conn.commit()
    finally:
        conn.close()


def export_report_team_members_json(db_path: Path = DB_PATH) -> str:
    return json.dumps({"team_members": list_report_team_members(db_path)}, indent=2)


def import_report_team_members_json(json_text: str, db_path: Path = DB_PATH) -> tuple[int, int]:
    try:
        data = json.loads(json_text)
        names = data["team_members"]
        if not isinstance(names, list):
            raise TypeError
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        raise ValueError("That doesn't look like a team-member backup file.") from e

    existing = {n.lower() for n in list_report_team_members(db_path)}
    added = 0
    skipped = 0
    for name in names:
        name = (name or "").strip()
        if not name:
            continue
        if name.lower() in existing:
            skipped += 1
            continue
        add_report_team_member(name, db_path)
        existing.add(name.lower())
        added += 1
    return added, skipped


# ---------------------------------------------------------------------------
# Weekly Report -- team members submit their update for a reporting
# period under a fixed set of categories (WEEKLY_REPORT_CATEGORIES
# above); a manager then compiles everyone's submissions for a period
# into one report, ready to email, copy, or download.
#
# PERIODS AND TIMEZONE
# ---------------------
# A reporting period is always Monday-Sunday. Its Monday (a `date`) is
# the single canonical identifier for a period everywhere in this
# module and in app.py -- see period_id_from_monday() /
# monday_from_period_id(). Formatted strings like "Aug 31-Sep 6, 2026"
# (format_period_label) or "Week of Aug 24-Aug 30, 2026" (period_label,
# kept for the existing compiled-report header/email subject) are for
# display only and are never stored or compared.
#
# "Current period", "now", and every edit-cutoff check below are all
# computed from server time in BUSINESS_TIMEZONE (see
# get_business_timezone()) -- never from a value supplied by the
# caller/browser. is_period_editable() is the single source of truth
# for the create/edit business rule and is used by both the UI (to
# show/hide the edit form) and upsert_weekly_report() (to actually
# reject a write) -- see that function's docstring for the rule.
# ---------------------------------------------------------------------------

# The business timezone used for every period/cutoff calculation.
# Override by adding BUSINESS_TIMEZONE to this app's secrets/env (any
# IANA zone name, e.g. "America/New_York") if your team isn't in
# Mountain Time. Defaults to Jeppesen ForeFlight's Englewood, CO
# headquarters timezone -- confirm this is correct for your team; it's
# an assumption, not a fact discovered from your data.
DEFAULT_BUSINESS_TIMEZONE = "America/Denver"


def get_business_timezone(secrets_getter=None) -> ZoneInfo:
    """The single configured business timezone -- see
    DEFAULT_BUSINESS_TIMEZONE above for how to override it. Falls back
    to the default if the configured name isn't a valid IANA zone."""
    name = _get_secret(secrets_getter, "BUSINESS_TIMEZONE") or DEFAULT_BUSINESS_TIMEZONE
    try:
        return ZoneInfo(name)
    except ZoneInfoNotFoundError:
        return ZoneInfo(DEFAULT_BUSINESS_TIMEZONE)


def _monday_of(d: date) -> date:
    """The Monday on or before `d` -- the start of d's Mon-Sun period."""
    return d - timedelta(days=d.weekday())  # Monday == weekday() 0


def current_period_monday(tz: ZoneInfo | None = None, now: datetime | None = None) -> date:
    """The Monday of the period containing "now" in the business
    timezone. `now` defaults to real server time -- only tests should
    pass it explicitly."""
    tz = tz or get_business_timezone()
    now = now or datetime.now(tz)
    if now.tzinfo is None:
        now = now.replace(tzinfo=tz)
    return _monday_of(now.astimezone(tz).date())


def period_id_from_monday(monday: date) -> str:
    """Canonical period identifier: the period's Monday as an ISO date
    (e.g. "2026-08-31"). This is what's stored in weekly_reports.period
    and passed to every function below that takes a `period` string."""
    return monday.isoformat()


def monday_from_period_id(period_id: str) -> date:
    """Inverse of period_id_from_monday(). Also accepts the legacy
    ISO-week identifier (e.g. "2026-W36") for any caller that still
    has one in hand -- the stored data itself is migrated by
    _migrate_legacy_period_format(), but this keeps the parser honest
    either way."""
    try:
        return date.fromisoformat(period_id)
    except ValueError:
        pass
    year_str, week_str = period_id.split("-W")
    return date.fromisocalendar(int(year_str), int(week_str), 1)


def available_periods(tz: ZoneInfo | None = None, now: datetime | None = None) -> list[date]:
    """The five Mondays for the Period dropdown: two weeks before the
    current period, one week before, the current period, one week
    after, and two weeks after -- in that order."""
    current_monday = current_period_monday(tz, now)
    return [current_monday + timedelta(weeks=offset) for offset in (-2, -1, 0, 1, 2)]


def format_period_label(monday: date, is_current: bool = False) -> str:
    """The Period dropdown's display label, e.g. 'Aug 17-23, 2026',
    'Aug 31-Sep 6, 2026' (month boundary), 'Dec 29, 2025-Jan 4, 2026'
    (year boundary), optionally suffixed ' -- Current'. Day numbers are
    never zero-padded (e.g. 'Sep 6', not 'Sep 06')."""

    def month_day(d: date) -> str:
        return f"{d.strftime('%b')} {d.day}"

    sunday = monday + timedelta(days=6)
    if monday.year != sunday.year:
        label = f"{month_day(monday)}, {monday.year}–{month_day(sunday)}, {sunday.year}"
    elif monday.month != sunday.month:
        label = f"{month_day(monday)}–{month_day(sunday)}, {sunday.year}"
    else:
        label = f"{month_day(monday)}–{sunday.day}, {sunday.year}"
    return f"{label} — Current" if is_current else label


def edit_cutoff(monday: date, tz: ZoneInfo | None = None) -> datetime:
    """The instant a period locks: 11:59:59 PM on the Monday following
    the period's Sunday (one full calendar day of grace after the
    period ends), in the business timezone."""
    tz = tz or get_business_timezone()
    following_monday = monday + timedelta(days=7)
    return datetime.combine(following_monday, time(23, 59, 59), tzinfo=tz)


def is_period_editable(monday: date, tz: ZoneInfo | None = None, now: datetime | None = None) -> bool:
    """Single source of truth for the create/edit business rule:

      - Outside the 5-period rolling window (more than 2 weeks before
        or after the current period): never editable. This is a
        defense-in-depth bound -- it rejects a crafted/stale period
        value that was never actually offered in the dropdown, even
        though the cutoff rule alone would already block old periods.
      - Current or future period (within the window): always editable.
      - Past period (within the window): editable only until
        edit_cutoff() -- Monday 11:59:59 PM the week after it ends.

    `now`/`tz` default to real server time in the business timezone --
    only tests should pass them explicitly. This function (not the
    caller's clock) is what upsert_weekly_report() checks before every
    write, so disabling a button in the UI is never the only guard.
    """
    tz = tz or get_business_timezone()
    now = now or datetime.now(tz)
    if now.tzinfo is None:
        now = now.replace(tzinfo=tz)
    now = now.astimezone(tz)

    current_monday = current_period_monday(tz, now)
    earliest = current_monday - timedelta(weeks=2)
    latest = current_monday + timedelta(weeks=2)
    if monday < earliest or monday > latest:
        return False
    if monday >= current_monday:
        return True
    return now < edit_cutoff(monday, tz)


class PeriodLockedError(RuntimeError):
    """Raised when a create/edit is rejected because its reporting
    period is no longer editable. Always raised from server-side code
    (upsert_weekly_report), using server time -- never bypassed by a
    stale browser tab or a disabled-but-not-actually-enforced button."""


def current_period(on_date: date | None = None) -> str:
    """Back-compat wrapper: returns the canonical period identifier
    (Monday ISO date) for `on_date`, or for "now" in the business
    timezone if not given. Prefer current_period_monday() directly in
    new code."""
    tz = get_business_timezone()
    monday = _monday_of(on_date) if on_date is not None else current_period_monday(tz)
    return period_id_from_monday(monday)


def period_label(period: str) -> str:
    """Turns a period identifier into 'Week of Aug 24-Aug 30, 2026' --
    used in the compiled report header and email subject line. Accepts
    both the canonical Monday-date identifier and the legacy ISO-week
    identifier (via monday_from_period_id()). For the Period dropdown's
    own label format ('Aug 17-23, 2026'), use format_period_label()
    instead."""
    try:
        monday = monday_from_period_id(period)
    except (ValueError, IndexError):
        return period
    sunday = monday + timedelta(days=6)
    if monday.month == sunday.month:
        return f"Week of {monday.strftime('%b %d')}-{sunday.strftime('%d, %Y')}"
    return f"Week of {monday.strftime('%b %d')} - {sunday.strftime('%b %d, %Y')}"


def format_timestamp(iso_utc: str | None, tz: ZoneInfo | None = None) -> str:
    """Formats a stored UTC ISO-8601 timestamp (as produced by _now())
    into the business timezone for display, e.g.
    '2026-08-31T21:05:33+00:00' -> 'Aug 31, 2026 03:05 PM'. Returns an
    em dash for a missing value, and falls back to the raw string if
    it doesn't parse (defensive -- shouldn't happen with our own
    _now() output)."""
    if not iso_utc:
        return "—"
    try:
        dt = datetime.fromisoformat(iso_utc)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        tz = tz or get_business_timezone()
        return dt.astimezone(tz).strftime("%b %d, %Y %I:%M %p")
    except ValueError:
        return iso_utc


def upsert_weekly_report(
    submitter: str,
    period: str,
    values: dict,
    db_path: Path = DB_PATH,
    tz: ZoneInfo | None = None,
    now: datetime | None = None,
) -> None:
    """Saves (or overwrites) one submitter's report for one period. Safe
    to call again for the same submitter/period -- e.g. if they come
    back to edit before the report gets compiled.

    Server-side enforcement of the edit-cutoff rule lives here, not
    just in the UI: this re-checks is_period_editable() using server
    time before writing anything, and raises PeriodLockedError if the
    period has locked since the page was loaded (e.g. a stale browser
    tab open past the Monday cutoff). `tz`/`now` are only for tests --
    real callers should never pass them, so this always uses actual
    server time.
    """
    submitter = (submitter or "").strip()
    if not submitter:
        raise ValueError("Select your name before submitting.")

    monday = monday_from_period_id(period)
    if not is_period_editable(monday, tz=tz, now=now):
        raise PeriodLockedError(
            f"The reporting period {format_period_label(monday)} is locked and "
            "can no longer be edited or created (editing closes at 11:59 PM "
            "the Monday after the period ends). Your changes were not saved."
        )

    cols = [key for key, _ in WEEKLY_REPORT_CATEGORIES]
    now = _now()
    conn = get_conn(db_path)
    try:
        existing = conn.execute(
            "SELECT id FROM weekly_reports WHERE submitter = ? AND period = ?", (submitter, period)
        ).fetchone()
        col_values = [values.get(c, "") for c in cols]
        if existing:
            set_clause = ", ".join(f"{c} = ?" for c in cols)
            conn.execute(
                f"UPDATE weekly_reports SET {set_clause}, updated_at = ? WHERE submitter = ? AND period = ?",
                (*col_values, now, submitter, period),
            )
        else:
            col_list = ", ".join(cols)
            placeholders = ", ".join("?" for _ in cols)
            conn.execute(
                f"""INSERT INTO weekly_reports (submitter, period, {col_list}, submitted_at, updated_at)
                    VALUES (?, ?, {placeholders}, ?, ?)""",
                (submitter, period, *col_values, now, now),
            )
        conn.commit()
    finally:
        conn.close()


def get_weekly_report(submitter: str, period: str, db_path: Path = DB_PATH):
    conn = get_conn(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return conn.execute(
            "SELECT * FROM weekly_reports WHERE submitter = ? AND period = ?", (submitter, period)
        ).fetchone()
    finally:
        conn.close()


def list_weekly_reports_for_period(period: str, db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    conn = get_conn(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return conn.execute(
            "SELECT * FROM weekly_reports WHERE period = ? ORDER BY submitter COLLATE NOCASE", (period,)
        ).fetchall()
    finally:
        conn.close()


def list_weekly_report_periods(db_path: Path = DB_PATH) -> list[str]:
    """Every period that has at least one submission, most recent first."""
    conn = get_conn(db_path)
    try:
        rows = conn.execute("SELECT DISTINCT period FROM weekly_reports ORDER BY period DESC").fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def _initials(name: str) -> str:
    """'James O'Dwyer' -> 'JO', 'Pete Cowley' -> 'PC'. Takes the first
    letter of each whitespace-separated word in the name, uppercased.
    Note: this can collide if two team members share the same initials
    (e.g. two "J.D." people) -- there's no dedup here since the roster
    is small and manually managed."""
    parts = [p for p in name.strip().split() if p]
    return "".join(p[0].upper() for p in parts)


def compile_weekly_report_text(period: str, db_path: Path = DB_PATH, expected_submitters: list[str] | None = None) -> str:
    """Groups every submission for a period by category into one
    plain-text report, formatted to be readable pasted directly into an
    email body. Each bullet is rendered on its own line as
    '* <text> (INITIALS)' rather than grouped under the person's full
    name, per the requested format."""
    rows = list_weekly_reports_for_period(period, db_path)
    lines = [f"WEEKLY TEAM REPORT -- {period_label(period)}", ""]

    if not rows:
        lines.append("No reports have been submitted for this period yet.")
    else:
        for key, label in WEEKLY_REPORT_CATEGORIES:
            entries = [(r["submitter"], (r[key] or "").strip()) for r in rows]
            entries = [(name, text) for name, text in entries if text]
            lines.append(label.upper())
            if entries:
                for name, text in entries:
                    initials = _initials(name)
                    # Each submission is stored as one bullet per line (from
                    # the UI's "+ Add entry" boxes) -- render every bullet
                    # as its own "* ... (INITIALS)" line.
                    bullets = [b.strip() for b in text.split("\n") if b.strip()]
                    if not bullets:
                        bullets = [text]
                    for bullet in bullets:
                        lines.append(f"* {bullet} ({initials})")
            else:
                lines.append("  (nothing reported)")
            lines.append("")

        submitted = sorted({r["submitter"] for r in rows})
        lines.append(f"Submitted by: {', '.join(submitted)}")

    if expected_submitters:
        missing = sorted(set(expected_submitters) - {r["submitter"] for r in rows})
        if missing:
            lines.append(f"Not yet submitted: {', '.join(missing)}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Weekly Report recipients -- editable from the UI, same Add/Delete +
# backup pattern as the salesperson list.
# ---------------------------------------------------------------------------
def list_recipients(db_path: Path = DB_PATH) -> list[str]:
    conn = get_conn(db_path)
    try:
        rows = conn.execute("SELECT email FROM report_recipients ORDER BY email COLLATE NOCASE").fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def add_recipient(email: str, db_path: Path = DB_PATH) -> None:
    email = (email or "").strip()
    if not email:
        raise ValueError("Email can't be blank.")
    if "@" not in email or "." not in email.rsplit("@", 1)[-1]:
        raise ValueError(f"'{email}' doesn't look like a valid email address.")
    conn = get_conn(db_path)
    try:
        try:
            conn.execute("INSERT INTO report_recipients (email) VALUES (?)", (email,))
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError(f"'{email}' is already in the list.")
    finally:
        conn.close()


def delete_recipient(email: str, db_path: Path = DB_PATH) -> None:
    conn = get_conn(db_path)
    try:
        conn.execute("DELETE FROM report_recipients WHERE email = ?", (email,))
        conn.commit()
    finally:
        conn.close()


def export_recipients_json(db_path: Path = DB_PATH) -> str:
    return json.dumps({"recipients": list_recipients(db_path)}, indent=2)


def import_recipients_json(json_text: str, db_path: Path = DB_PATH) -> tuple[int, int]:
    try:
        data = json.loads(json_text)
        emails = data["recipients"]
        if not isinstance(emails, list):
            raise TypeError
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        raise ValueError("That doesn't look like a recipients backup file.") from e

    existing = {e.lower() for e in list_recipients(db_path)}
    added = 0
    skipped = 0
    for email in emails:
        email = (email or "").strip()
        if not email:
            continue
        if email.lower() in existing:
            skipped += 1
            continue
        try:
            add_recipient(email, db_path)
            existing.add(email.lower())
            added += 1
        except ValueError:
            skipped += 1
    return added, skipped


def send_weekly_report_email(
    recipients: list[str],
    subject: str,
    body: str,
    secrets_getter=None,
) -> None:
    """Sends the compiled weekly report by email via Gmail's SMTP relay.

    secrets_getter: optional callable returning a dict-like of secrets
    (e.g. st.secrets), same injectable-dependency pattern as
    get_anthropic_client() above, so this stays testable without
    importing streamlit here.

    Required secrets:
      GMAIL_ADDRESS      -- the Gmail account emails are sent from.
      GMAIL_APP_PASSWORD -- a 16-character Gmail App Password, NOT the
                             normal account password. Generate one at
                             myaccount.google.com/apppasswords (requires
                             2-Step Verification to be turned on first).

    Recipients are validated email addresses (already checked by
    add_recipient()), so no further sanitization of the address list is
    needed here. The connection uses SMTP over SSL/TLS on port 465 with
    certificate verification -- never plaintext SMTP.
    """
    import smtplib
    import ssl
    from email.message import EmailMessage

    recipients = [r.strip() for r in (recipients or []) if r and r.strip()]
    if not recipients:
        raise ValueError("Add at least one recipient before sending.")

    sender = _get_secret(secrets_getter, "GMAIL_ADDRESS")
    app_password = _get_secret(secrets_getter, "GMAIL_APP_PASSWORD")
    if not sender or not app_password:
        raise RuntimeError(
            "Email sending isn't configured yet. Add GMAIL_ADDRESS and "
            "GMAIL_APP_PASSWORD to this app's secrets first."
        )

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=20) as server:
            server.login(sender, app_password)
            server.send_message(msg)
    except smtplib.SMTPAuthenticationError as e:
        # Never surface the password itself -- just point at what to check.
        raise RuntimeError(
            "Gmail rejected the login. Double-check GMAIL_ADDRESS and make "
            "sure GMAIL_APP_PASSWORD is a 16-character App Password (not "
            "your regular Gmail password)."
        ) from e
    except (smtplib.SMTPException, OSError) as e:
        raise RuntimeError(f"Couldn't send the email: {e}") from e


def insert_request(fields: dict, db_path: Path = DB_PATH) -> str:
    request_id = uuid.uuid4().hex[:8].upper()
    conn = get_conn(db_path)
    try:
        conn.execute(
            """INSERT INTO requests (
                id, created_at, updated_at, salesperson, jeppesen_acct_nbr,
                payment_method, amount, notes, onboarding_filename,
                onboarding_bytes, status, result_filename, result_bytes, error_message
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                request_id,
                _now(),
                _now(),
                fields["salesperson"],
                fields["jeppesen_acct_nbr"],
                fields["payment_method"],
                fields.get("amount"),
                fields.get("notes", ""),
                fields.get("onboarding_filename"),
                fields.get("onboarding_bytes"),
                STATUS_SUBMITTED if fields.get("onboarding_bytes") is None else STATUS_PROCESSING,
                None,
                None,
                None,
            ),
        )
        conn.commit()
    finally:
        conn.close()
    return request_id


def update_status(request_id: str, status: str, error_message: str | None = None, db_path: Path = DB_PATH) -> None:
    conn = get_conn(db_path)
    try:
        conn.execute(
            "UPDATE requests SET status = ?, error_message = ?, updated_at = ? WHERE id = ?",
            (status, error_message, _now(), request_id),
        )
        conn.commit()
    finally:
        conn.close()


def update_result(request_id: str, result_bytes: bytes, result_filename: str, db_path: Path = DB_PATH) -> None:
    conn = get_conn(db_path)
    try:
        conn.execute(
            "UPDATE requests SET result_bytes = ?, result_filename = ?, updated_at = ? WHERE id = ?",
            (result_bytes, result_filename, _now(), request_id),
        )
        conn.commit()
    finally:
        conn.close()


def fetch_all_requests(db_path: Path = DB_PATH) -> list[sqlite3.Row]:
    conn = get_conn(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return conn.execute("SELECT * FROM requests ORDER BY created_at DESC").fetchall()
    finally:
        conn.close()


def fetch_request(request_id: str, db_path: Path = DB_PATH):
    conn = get_conn(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return conn.execute("SELECT * FROM requests WHERE id = ?", (request_id,)).fetchone()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Parsing bridge
# ---------------------------------------------------------------------------
def _get_secret(secrets_getter, key: str):
    if secrets_getter is not None:
        try:
            val = secrets_getter().get(key)
            if val:
                return val
        except Exception:
            pass
    return os.environ.get(key)


def get_anthropic_client(secrets_getter=None):
    """secrets_getter: optional callable returning a dict-like of secrets
    (e.g. st.secrets), kept as an injectable dependency so this stays
    testable without importing streamlit here.

    Supports two auth paths, tried in this order:
      1. Direct Anthropic Console key (ANTHROPIC_API_KEY, starts with
         sk-ant-) -> talks to api.anthropic.com directly.
      2. AWS Bedrock (AWS_BEARER_TOKEN_BEDROCK + AWS_REGION) -> talks to
         Bedrock's Anthropic-compatible endpoint instead. Enterprise
         accounts that provision Claude access through AWS rather than
         a standalone Anthropic Console account will use this path.
         Optionally set BEDROCK_MODEL_ID if the default model ID isn't
         enabled for your account -- check the Bedrock console's model
         access page for the exact ID your account can call.

    Returns (client, model) -- always pass `model` through to
    pe.extract_page() rather than relying on parsing_engine.py's default,
    since Bedrock model IDs look nothing like the direct-API model name.
    """
    import anthropic

    api_key = _get_secret(secrets_getter, "ANTHROPIC_API_KEY")
    if api_key:
        return anthropic.Anthropic(api_key=api_key), pe.MODEL

    bedrock_token = _get_secret(secrets_getter, "AWS_BEARER_TOKEN_BEDROCK")
    if bedrock_token:
        region = _get_secret(secrets_getter, "AWS_REGION") or "us-east-1"
        model = _get_secret(secrets_getter, "BEDROCK_MODEL_ID") or "us.anthropic.claude-sonnet-5"
        client = anthropic.AnthropicBedrock(aws_region=region, api_key=bedrock_token)
        return client, model

    raise RuntimeError(
        "No API credentials configured. Set ANTHROPIC_API_KEY (direct Anthropic "
        "Console key), or AWS_BEARER_TOKEN_BEDROCK + AWS_REGION (AWS Bedrock key), "
        "under App settings -> Secrets."
    )


def run_parsing(request_id: str, pdf_bytes: bytes, progress_callback=None, secrets_getter=None, db_path: Path = DB_PATH, salesperson: str | None = None) -> None:
    """Runs the onboarding PDF through parsing_engine.py's pipeline and
    stores the resulting .xlsx (or the error) back on the request row.
    progress_callback(fraction: float, text: str), if given, is called
    after each page is read. `salesperson` is the name selected in the
    intake form's Salesperson dropdown -- passed through to General
    Info's "Jeppesen FF Account Exec" line instead of a placeholder."""
    try:
        client, model = get_anthropic_client(secrets_getter)
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = Path(tmp) / "onboarding.pdf"
            pdf_path.write_bytes(pdf_bytes)
            out_path = Path(tmp) / "operational_info.xlsx"

            pages = pe.render_pages(pdf_path)
            sections = []
            for i, png in enumerate(pages, start=1):
                if progress_callback is not None:
                    progress_callback(i / len(pages), f"Reading page {i} of {len(pages)}...")
                result = pe.extract_page(png, client, model=model)
                for section in result.get("sections", []):
                    section["page"] = i
                    sections.append(section)

            by_subject = pe.map_to_subjects(sections, salesperson=salesperson)
            pe.write_workbook(by_subject, out_path)
            result_bytes = out_path.read_bytes()

        update_result(request_id, result_bytes, "operational_info.xlsx", db_path)
        update_status(request_id, STATUS_COMPLETE, db_path=db_path)
    except Exception as e:
        update_status(request_id, STATUS_FAILED, error_message=str(e), db_path=db_path)
