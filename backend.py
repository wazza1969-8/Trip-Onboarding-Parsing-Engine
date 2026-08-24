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
    return conn


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


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
def get_anthropic_client(secrets_getter=None):
    """secrets_getter: optional callable returning a dict-like of secrets
    (e.g. st.secrets), kept as an injectable dependency so this stays
    testable without importing streamlit here."""
    import anthropic

    api_key = None
    if secrets_getter is not None:
        try:
            api_key = secrets_getter().get("ANTHROPIC_API_KEY")
        except Exception:
            pass
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not configured. Set it as an environment variable, "
            "or add it under App settings -> Secrets if deployed on Streamlit Community Cloud."
        )
    return anthropic.Anthropic(api_key=api_key)


def run_parsing(request_id: str, pdf_bytes: bytes, progress_callback=None, secrets_getter=None, db_path: Path = DB_PATH) -> None:
    """Runs the onboarding PDF through parsing_engine.py's pipeline and
    stores the resulting .xlsx (or the error) back on the request row.
    progress_callback(fraction: float, text: str), if given, is called
    after each page is read."""
    try:
        client = get_anthropic_client(secrets_getter)
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = Path(tmp) / "onboarding.pdf"
            pdf_path.write_bytes(pdf_bytes)
            out_path = Path(tmp) / "operational_info.xlsx"

            pages = pe.render_pages(pdf_path)
            sections = []
            for i, png in enumerate(pages, start=1):
                if progress_callback is not None:
                    progress_callback(i / len(pages), f"Reading page {i} of {len(pages)}...")
                result = pe.extract_page(png, client)
                for section in result.get("sections", []):
                    section["page"] = i
                    sections.append(section)

            by_subject = pe.map_to_subjects(sections)
            pe.write_workbook(by_subject, out_path)
            result_bytes = out_path.read_bytes()

        update_result(request_id, result_bytes, "operational_info.xlsx", db_path)
        update_status(request_id, STATUS_COMPLETE, db_path=db_path)
    except Exception as e:
        update_status(request_id, STATUS_FAILED, error_message=str(e), db_path=db_path)
