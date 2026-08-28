"""
Backend for the desktop (local-only) Trip Support Intake app.

Same shape as backend.py (the hosted-app version), but:
  - The SQLite database lives in a per-user, OS-appropriate app-data
    folder rather than next to the script, so it works no matter where
    the packaged app is installed, and nothing is shared between users
    or machines.
  - The Anthropic API key comes from config.py (baked in at build time)
    instead of Streamlit secrets.

Kept separate from desktop_app.py (the Tkinter UI) so this logic can be
tested directly in a plain Python process.
"""

import os
import platform
import sqlite3
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path

import parsing_engine as pe

STATUS_SUBMITTED = "Submitted"
STATUS_PROCESSING = "Processing"
STATUS_COMPLETE = "Complete"
STATUS_FAILED = "Failed"


def local_app_data_dir() -> Path:
    """OS-appropriate, per-user, writable folder -- works whether the
    app is run from source or from a PyInstaller-built executable
    (which is often read-only / extracted to a temp location)."""
    system = platform.system()
    if system == "Windows":
        base = os.environ.get("APPDATA", str(Path.home()))
        folder = Path(base) / "TripSupportIntake"
    elif system == "Darwin":
        folder = Path.home() / "Library" / "Application Support" / "TripSupportIntake"
    else:
        folder = Path.home() / ".trip_support_intake"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


DB_PATH = local_app_data_dir() / "requests.db"

# Seeded into the salespeople table the first time it's empty -- after
# that, the table (editable from the UI) is the source of truth, not
# this list.
DEFAULT_SALESPEOPLE = ["Jane Smith", "John Doe", "Alex Johnson", "Morgan Lee"]


# ---------------------------------------------------------------------------
# Database helpers (local-only; nothing here talks to a network except
# the parsing call itself)
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
# Salesperson list -- editable from the UI (Add/Edit/Delete), stored in
# this user's local requests.db, same as their requests.
# ---------------------------------------------------------------------------
def list_salespeople(db_path: Path = DB_PATH) -> list[str]:
    conn = get_conn(db_path)
    try:
        rows = conn.execute("SELECT name FROM salespeople ORDER BY name COLLATE NOCASE").fetchall()
        if not rows:
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
def get_anthropic_client():
    import anthropic

    api_key = None
    try:
        import config  # the file you fill in before building
        if config.ANTHROPIC_API_KEY and config.ANTHROPIC_API_KEY != "REPLACE_WITH_YOUR_ANTHROPIC_API_KEY":
            api_key = config.ANTHROPIC_API_KEY
    except ImportError:
        pass
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "No Anthropic API key configured. Fill in ANTHROPIC_API_KEY in config.py "
            "before building, or set it as an environment variable."
        )
    return anthropic.Anthropic(api_key=api_key)


def run_parsing(request_id: str, pdf_bytes: bytes, progress_callback=None, db_path: Path = DB_PATH, salesperson: str | None = None) -> None:
    """Runs the onboarding PDF through parsing_engine.py's pipeline and
    stores the resulting .xlsx (or the error) back on the request row.
    progress_callback(fraction: float, text: str), if given, is called
    after each page is read -- call it from the main thread only if your
    GUI toolkit requires that (see desktop_app.py for how this is used
    safely from a background thread)."""
    try:
        client = get_anthropic_client()
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

            by_subject = pe.map_to_subjects(sections, salesperson=salesperson)
            pe.write_workbook(by_subject, out_path)
            result_bytes = out_path.read_bytes()

        update_result(request_id, result_bytes, "operational_info.xlsx", db_path)
        update_status(request_id, STATUS_COMPLETE, db_path=db_path)
    except Exception as e:
        update_status(request_id, STATUS_FAILED, error_message=str(e), db_path=db_path)
