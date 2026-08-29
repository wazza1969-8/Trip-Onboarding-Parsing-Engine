"""
Backend for the Trip Support Intake app -- SQLite-backed request store
plus the bridge into parsing_engine.py.

Kept separate from app.py (the Streamlit UI) so the database logic can
be exercised directly in a plain Python process/tests, without needing
a live Streamlit runtime context.
"""

import os
import sqlite3
import tempfile
import uuid
from datetime import datetime, timezone
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
