"""
Backend for the Trip Support Intake app -- SQLite-backed request store
plus the bridge into parsing_engine.py.

Kept separate from app.py (the Streamlit UI) so the database logic can
be exercised directly in a plain Python process/tests, without needing
a live Streamlit runtime context.
"""

import json
import os
import sqlite3
import tempfile
import uuid
from datetime import date, datetime, timezone
from pathlib import Path

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
    return conn


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
# Weekly Report -- team members submit their update for the current
# period (an ISO week, e.g. "2026-W35") under a fixed set of categories
# (WEEKLY_REPORT_CATEGORIES above); a manager then compiles everyone's
# submissions for a period into one report, ready to copy into an email
# or download. Sending isn't wired up yet -- see compile_weekly_report_text().
# ---------------------------------------------------------------------------
def current_period(on_date: date | None = None) -> str:
    d = on_date or datetime.now(timezone.utc).date()
    iso_year, iso_week, _ = d.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def period_label(period: str) -> str:
    """Turns 'YYYY-Www' into a human-readable date range, e.g.
    'Week of Aug 24 - Aug 30, 2026'. Falls back to the raw string if it
    doesn't parse."""
    try:
        year_str, week_str = period.split("-W")
        monday = date.fromisocalendar(int(year_str), int(week_str), 1)
        sunday = date.fromisocalendar(int(year_str), int(week_str), 7)
        if monday.month == sunday.month:
            return f"Week of {monday.strftime('%b %d')}-{sunday.strftime('%d, %Y')}"
        return f"Week of {monday.strftime('%b %d')} - {sunday.strftime('%b %d, %Y')}"
    except (ValueError, IndexError):
        return period


def upsert_weekly_report(submitter: str, period: str, values: dict, db_path: Path = DB_PATH) -> None:
    """Saves (or overwrites) one submitter's report for one period. Safe
    to call again for the same submitter/period -- e.g. if they come
    back to edit before the report gets compiled."""
    submitter = (submitter or "").strip()
    if not submitter:
        raise ValueError("Select your name before submitting.")
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


def compile_weekly_report_text(period: str, db_path: Path = DB_PATH, expected_submitters: list[str] | None = None) -> str:
    """Groups every submission for a period by category into one
    plain-text report, formatted to be readable pasted directly into an
    email body."""
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
                    lines.append(f"  - {name}: {text}")
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
