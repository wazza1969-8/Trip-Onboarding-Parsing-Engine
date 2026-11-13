"""
Backend foundation for the Document Repository feature -- Phase 1 only.

Nothing in app.py or backend.py calls into this module yet (no screens,
no wiring). It is written to stand alone and be fully exercised by
test_documents_repo.py without a live Streamlit runtime context, same
as backend.py's own stated design goal.

SAME DATABASE, NEW TABLES ONLY
-------------------------------
This module targets the same production database as backend.py (the
Turso database selected by TURSO_DATABASE_URL/TURSO_AUTH_TOKEN, falling
back to a local SQLite file for dev/tests) -- see get_conn() below,
which duplicates backend.get_conn()'s connection-selection logic rather
than importing it, so this module has no dependency on backend.py's
internals and can be read/tested in isolation. ensure_schema() only
ever creates its own doc_users / doc_categories / doc_category_access /
documents / document_versions / document_blobs / doc_audit_log tables
(all CREATE TABLE IF NOT EXISTS) -- it never alters or touches
requests, salespeople, weekly_reports, report_recipients, or
report_team_members.

STORAGE
-------
Document bytes never live directly in the `documents` or
`document_versions` tables. They go through the DocumentStore
interface (put/get/delete by storage_key); the only implementation
here, DatabaseBlobStore, chunks them (~1 MB each) into document_blobs.
Every service function below goes through a DocumentStore -- none of
them execute SQL against document_blobs directly -- so swapping in a
different backend (e.g. S3) later only means writing a new
DocumentStore subclass, not touching any function in this file.

AUTHORIZATION
-------------
Every function that reads or changes document-repository data takes
the acting user's email as its first argument and checks it itself --
deny by default, no caller is trusted to have already checked. Three
roles:
  - admin: via DOC_REPO_ADMIN_EMAILS (comma-separated secret/env var,
    checked on every call, so the first admin can never be locked out
    by an empty or corrupted doc_users table) or a doc_users row with
    role='admin'. Can do everything.
  - contributor / viewer: granted per-category through
    doc_category_access (view < contribute < manage). A category with
    no access rows is visible to admins only -- there is no
    "everyone can see it" default.
Old versions of a document (anything but its current_version_id) are
downloadable by admins and contributors (contribute or manage access)
only -- never by a viewer. Archiving/restoring a document requires
manage-level access (or admin): it is a category-stewardship action,
not a day-to-day content edit the way uploading/replacing/editing
metadata is (those only need contribute).

Every mutation, every download, and every denied attempt writes a
doc_audit_log row. doc_audit_log itself is append-only on purpose --
there is deliberately no update or delete function for it anywhere in
this file. The one exception is _require_actor()'s check for a blank
actor email: that is rejected before any connection is opened, so
there is no identity to attribute an audit row to.

There is deliberately no permanent-delete function for documents --
only archive_document() / restore_document().
"""

import hashlib
import io
import os
import sqlite3
import uuid
import zipfile
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

try:
    # Optional until requirements.txt is installed with this package --
    # same fallback as backend.py. Only actually used once
    # TURSO_DATABASE_URL/TURSO_AUTH_TOKEN are configured; everything
    # falls back to local SQLite otherwise, so this module still
    # imports fine without it.
    import turso_serverless
except ImportError:  # pragma: no cover - exercised only if the package is missing
    turso_serverless = None

# Same reasoning as backend._INTEGRITY_ERRORS: every `except
# sqlite3.IntegrityError:` guarding a UNIQUE-constraint violation below
# also needs to catch the Turso driver's own IntegrityError once that
# backend is in use.
_INTEGRITY_ERRORS = (
    (sqlite3.IntegrityError, turso_serverless.IntegrityError)
    if turso_serverless is not None
    else (sqlite3.IntegrityError,)
)

DEFAULT_DB_PATH = Path(__file__).parent / "documents.db"


# ---------------------------------------------------------------------------
# Roles, access levels, upload rules
# ---------------------------------------------------------------------------
ROLE_ADMIN = "admin"
ROLE_CONTRIBUTOR = "contributor"
ROLE_VIEWER = "viewer"
VALID_ROLES = {ROLE_ADMIN, ROLE_CONTRIBUTOR, ROLE_VIEWER}

LEVEL_VIEW = "view"
LEVEL_CONTRIBUTE = "contribute"
LEVEL_MANAGE = "manage"
VALID_LEVELS = {LEVEL_VIEW, LEVEL_CONTRIBUTE, LEVEL_MANAGE}
_LEVEL_RANK = {LEVEL_VIEW: 1, LEVEL_CONTRIBUTE: 2, LEVEL_MANAGE: 3}

STATUS_ACTIVE = "active"
STATUS_ARCHIVED = "archived"

# Editable in one place, per the spec -- extend this set to allow a new
# file type. Macro-enabled/active formats (docm, xlsm, pptm, ...) are
# deliberately absent.
ALLOWED_EXTENSIONS = {"pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "png", "jpg", "jpeg"}

MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024  # 25 MB

# ~1 MB per document_blobs row -- small enough that a single chunk is a
# comfortable network round trip against Turso, large enough that a
# 25 MB file is at most 25 rows.
CHUNK_SIZE = 1024 * 1024

_MIME_TYPES = {
    "pdf": "application/pdf",
    "doc": "application/msword",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xls": "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "ppt": "application/vnd.ms-powerpoint",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
}

# Seeded by seed_default_categories() -- see its docstring for why that
# function is not wired into any call site yet.
DEFAULT_CATEGORIES = [
    "Sales & Marketing Materials",
    "Product & Service Information",
    "Pricing & Commercial Documents",
    "Contracts & Agreements",
    "Standard Operating Procedures (SOPs)",
    "Training & Reference Materials",
    "Partnership Documents",
    "Templates & Forms",
    "Other Documents",
]


class PermissionDenied(RuntimeError):
    """Raised when the acting user isn't authorized for the requested
    action. Always raised server-side, from the service function itself
    -- never left to a caller to check first."""


class ValidationError(ValueError):
    """Raised when an uploaded file fails the size, extension, filename,
    or content-signature checks in validate_upload()."""


class StorageIntegrityError(RuntimeError):
    """Raised by download_version() when the bytes read back from the
    DocumentStore don't match the size/SHA-256 recorded at upload time
    -- e.g. an interrupted write left partial chunks behind."""


# ---------------------------------------------------------------------------
# Database connection -- mirrors backend.get_conn()'s Turso-or-local
# selection exactly (see module docstring for why this is duplicated
# rather than imported).
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


def _turso_credentials():
    url = _get_secret(None, "TURSO_DATABASE_URL")
    token = _get_secret(None, "TURSO_AUTH_TOKEN")
    return url, token


def _set_row_factory(conn) -> None:
    if isinstance(conn, sqlite3.Connection):
        conn.row_factory = sqlite3.Row
    else:
        conn.row_factory = turso_serverless.Row


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# Keyed by target (Turso URL, or local file path) like backend.py's own
# _schema_ready_targets -- so CREATE TABLE IF NOT EXISTS only actually
# runs once per target per process, and a test against a temp SQLite
# file always gets its own one-time init.
_schema_ready_targets: set[str] = set()


def ensure_schema(conn) -> None:
    """Creates every table this module owns, idempotently (CREATE TABLE
    IF NOT EXISTS). Takes an already-open connection -- called once per
    target by get_conn() below, but safe to call again directly (e.g.
    from a test) against any connection."""
    conn.execute(
        """CREATE TABLE IF NOT EXISTS doc_users (
            email TEXT PRIMARY KEY,
            role TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT,
            created_by TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS doc_categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE COLLATE NOCASE,
            description TEXT,
            sort_order INTEGER NOT NULL DEFAULT 0,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT,
            created_by TEXT,
            updated_at TEXT,
            updated_by TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS doc_category_access (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER NOT NULL,
            subject TEXT NOT NULL,
            level TEXT NOT NULL,
            UNIQUE(category_id, subject)
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS documents (
            id TEXT PRIMARY KEY,
            category_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            description TEXT,
            keywords TEXT,
            status TEXT NOT NULL DEFAULT 'active',
            current_version_id TEXT,
            created_at TEXT,
            created_by TEXT,
            updated_at TEXT,
            updated_by TEXT,
            archived_at TEXT,
            archived_by TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS document_versions (
            id TEXT PRIMARY KEY,
            document_id TEXT NOT NULL,
            version_no INTEGER NOT NULL,
            filename TEXT NOT NULL,
            extension TEXT NOT NULL,
            mime_type TEXT,
            size_bytes INTEGER NOT NULL,
            sha256 TEXT NOT NULL,
            storage_key TEXT NOT NULL,
            change_note TEXT,
            uploaded_at TEXT,
            uploaded_by TEXT,
            UNIQUE(document_id, version_no)
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS document_blobs (
            storage_key TEXT NOT NULL,
            chunk_index INTEGER NOT NULL,
            data BLOB NOT NULL,
            PRIMARY KEY (storage_key, chunk_index)
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS doc_audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            actor_email TEXT,
            action TEXT NOT NULL,
            document_id TEXT,
            version_id TEXT,
            category_id INTEGER,
            detail TEXT
        )"""
    )
    conn.commit()


def get_conn(db_path: Path = DEFAULT_DB_PATH):
    """Opens a connection to the same database backend.get_conn() uses
    (Turso in production, a local SQLite file otherwise) and makes sure
    this module's own tables exist on it, once per target per process."""
    turso_url, turso_token = _turso_credentials()
    if turso_url:
        if turso_serverless is None:
            raise RuntimeError(
                "TURSO_DATABASE_URL is set, but the turso_serverless package "
                "isn't installed. Add turso_serverless to requirements.txt "
                "and redeploy."
            )
        conn = turso_serverless.connect(turso_url, auth_token=turso_token)
        conn.isolation_level = None
        cache_key = turso_url
    else:
        conn = sqlite3.connect(db_path, check_same_thread=False)
        cache_key = str(db_path)

    if cache_key not in _schema_ready_targets:
        ensure_schema(conn)
        _schema_ready_targets.add(cache_key)
    return conn


# ---------------------------------------------------------------------------
# Authorization helpers
# ---------------------------------------------------------------------------
def _normalize_email(email: str) -> str:
    return (email or "").strip().lower()


def _require_actor(actor_email: str) -> str:
    actor_email = _normalize_email(actor_email)
    if not actor_email:
        raise PermissionDenied("An actor email is required.")
    return actor_email


def _admin_emails_from_secret(secrets_getter=None) -> set[str]:
    raw = _get_secret(secrets_getter, "DOC_REPO_ADMIN_EMAILS") or ""
    return {e.strip().lower() for e in raw.split(",") if e.strip()}


def _is_admin(conn, actor_email: str, secrets_getter=None) -> bool:
    if actor_email in _admin_emails_from_secret(secrets_getter):
        return True
    row = conn.execute("SELECT role, is_active FROM doc_users WHERE email = ?", (actor_email,)).fetchone()
    return bool(row and row[0] == ROLE_ADMIN and row[1])


def is_admin(actor_email: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> bool:
    """Read-only, side-effect-free admin check for callers (e.g. a page
    that decides whether to show a tile) that only want a yes/no answer
    and must never provision anything just by asking.

    Unlike every other function in this module, this deliberately does
    NOT go through get_conn(): that would run ensure_schema() and
    silently create this module's 7 tables the first time anyone merely
    *looks* at whether to show an admin-only link. Instead it opens its
    own connection, confirms doc_users already exists before querying
    it, and never writes anything -- no audit row either, since this
    isn't a mutation or a document access, just a visibility check.

    Never raises: a missing table, a missing/unreachable database, or a
    malformed secret all simply mean "not an admin" rather than an
    error reaching the caller.
    """
    try:
        actor_email = _normalize_email(actor_email)
        if not actor_email:
            return False
        if actor_email in _admin_emails_from_secret(secrets_getter):
            return True

        turso_url, turso_token = _turso_credentials()
        if turso_url:
            if turso_serverless is None:
                return False
            conn = turso_serverless.connect(turso_url, auth_token=turso_token)
        else:
            # Read-only URI connect -- unlike a plain sqlite3.connect(),
            # this raises instead of creating db_path if it doesn't
            # exist yet, so a mere visibility check can never conjure a
            # database file into existence.
            conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        try:
            exists = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'doc_users'"
            ).fetchone()
            if not exists:
                return False
            row = conn.execute(
                "SELECT role, is_active FROM doc_users WHERE email = ?", (actor_email,)
            ).fetchone()
            return bool(row and row[0] == ROLE_ADMIN and row[1])
        finally:
            conn.close()
    except Exception:
        return False


def is_admin_via_secret(actor_email: str, secrets_getter=None) -> bool:
    """Read-only, side-effect-free (Phase 4 addition): True only if
    actor_email is listed in the DOC_REPO_ADMIN_EMAILS secret
    specifically -- never touches doc_users or opens any connection.
    Distinct from is_admin() (which also counts an active doc_users
    admin row): used by the admin UI to warn that this particular
    admin's access isn't editable there at all, since the secret isn't
    something this app can read back or change -- only a doc_users row
    is. Never raises; a malformed secret just means False.
    """
    try:
        actor_email = _normalize_email(actor_email)
        if not actor_email:
            return False
        return actor_email in _admin_emails_from_secret(secrets_getter)
    except Exception:
        return False


def _access_level(conn, actor_email: str, category_id) -> str | None:
    row = conn.execute(
        "SELECT level FROM doc_category_access WHERE category_id = ? AND subject = ?",
        (category_id, actor_email),
    ).fetchone()
    return row[0] if row else None


def _has_level(conn, actor_email: str, category_id, minimum: str) -> bool:
    level = _access_level(conn, actor_email, category_id)
    return level is not None and _LEVEL_RANK[level] >= _LEVEL_RANK[minimum]


def _authorized_for_category(conn, actor_email: str, category_id, min_level: str, secrets_getter=None) -> bool:
    return _is_admin(conn, actor_email, secrets_getter) or _has_level(conn, actor_email, category_id, min_level)


def _audit(conn, actor_email: str, action: str, document_id=None, version_id=None, category_id=None, detail=None) -> None:
    conn.execute(
        """INSERT INTO doc_audit_log (ts, actor_email, action, document_id, version_id, category_id, detail)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (_now(), actor_email, action, document_id, version_id, category_id, detail),
    )


def _deny(conn, actor_email: str, action: str, **audit_kwargs) -> None:
    """Writes a `<action>_denied` audit row, commits it, then raises
    PermissionDenied -- the single choke point every authorization
    check below funnels through, so a denial is never silently
    unlogged."""
    _audit(conn, actor_email, f"{action}_denied", **audit_kwargs)
    conn.commit()
    raise PermissionDenied(f"'{actor_email}' is not allowed to {action.replace('_', ' ')}.")


# ---------------------------------------------------------------------------
# Upload validation
# ---------------------------------------------------------------------------
_CONTROL_CHARS = {chr(c) for c in range(0x20)} | {chr(0x7F)}


def sanitize_filename(filename: str, max_length: int = 255) -> str:
    """Strips any directory path and control characters, then limits
    length -- used for both the stored filename and before extracting
    an extension, so neither can smuggle a path or odd bytes. Truncates
    from the end of the base name, not the end of the whole string, so
    a too-long name keeps its extension instead of losing it."""
    name = (filename or "").replace("\\", "/")
    name = name.rsplit("/", 1)[-1]
    name = "".join(ch for ch in name if ch not in _CONTROL_CHARS).strip()
    if not name:
        raise ValidationError("Filename can't be blank.")
    if len(name) <= max_length:
        return name
    if "." in name:
        base, ext = name.rsplit(".", 1)
        suffix = f".{ext}"
        return base[: max_length - len(suffix)] + suffix
    return name[:max_length]


def _extension_of(filename: str) -> str:
    if "." not in filename:
        return ""
    return filename.rsplit(".", 1)[-1].lower()


def _guess_mime_type(extension: str) -> str:
    return _MIME_TYPES.get(extension, "application/octet-stream")


_SIMPLE_SIGNATURES = {
    "pdf": (b"%PDF",),
    "png": (b"\x89PNG\r\n\x1a\n",),
    "jpg": (b"\xff\xd8\xff",),
    "jpeg": (b"\xff\xd8\xff",),
}
_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_OLE2_EXTENSIONS = {"doc", "xls", "ppt"}
_ZIP_OOXML_EXTENSIONS = {"docx", "xlsx", "pptx"}
_ZIP_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")


def _check_file_signature(extension: str, data: bytes) -> None:
    """Checks the file's magic bytes match its claimed extension --
    never trusts the extension alone. For docx/xlsx/pptx, also opens the
    zip central directory (never extracting any entry) to confirm it's
    a real OOXML package and reject one smuggling a VBA macro project
    under a non-macro extension."""
    if extension in _SIMPLE_SIGNATURES:
        if not any(data.startswith(sig) for sig in _SIMPLE_SIGNATURES[extension]):
            raise ValidationError(f"File content doesn't look like a .{extension} file.")
        return

    if extension in _OLE2_EXTENSIONS:
        if not data.startswith(_OLE2_SIGNATURE):
            raise ValidationError(f"File content doesn't look like a .{extension} file.")
        return

    if extension in _ZIP_OOXML_EXTENSIONS:
        if not any(data.startswith(sig) for sig in _ZIP_SIGNATURES):
            raise ValidationError(f"File content doesn't look like a .{extension} file.")
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = zf.namelist()
        except zipfile.BadZipFile:
            raise ValidationError(f"File content doesn't look like a .{extension} file.")
        if "[Content_Types].xml" not in names:
            raise ValidationError(f"File content doesn't look like a .{extension} file.")
        if any(n.lower().endswith("vbaproject.bin") for n in names):
            raise ValidationError(
                "Macro-enabled Office files aren't allowed. Save as a non-macro "
                f".{extension} file and try again."
            )
        return

    raise ValidationError(f"'.{extension}' files aren't allowed.")


def validate_upload(filename: str, data: bytes) -> tuple[str, str]:
    """Validates size, extension allowlist, filename, and content
    signature for an uploaded file. Returns (sanitized_filename,
    extension) on success; raises ValidationError describing the first
    problem found otherwise. Shared by add_document() and
    replace_document() so neither can skip a check the other runs."""
    if not data:
        raise ValidationError("The file is empty.")
    if len(data) > MAX_FILE_SIZE_BYTES:
        raise ValidationError(
            f"The file is too large ({len(data):,} bytes). The maximum is "
            f"{MAX_FILE_SIZE_BYTES:,} bytes (25 MB)."
        )
    name = sanitize_filename(filename)
    extension = _extension_of(name)
    if extension not in ALLOWED_EXTENSIONS:
        raise ValidationError(
            f"'.{extension or '(none)'}' files aren't allowed. "
            f"Allowed types: {', '.join(sorted(ALLOWED_EXTENSIONS))}."
        )
    _check_file_signature(extension, data)
    return name, extension


def _escape_like(term: str) -> str:
    """Escapes %, _, and the escape character itself so a LIKE search
    treats them as literal characters -- paired with ESCAPE '\\' at
    every call site below."""
    return term.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


# ---------------------------------------------------------------------------
# DocumentStore -- the only thing allowed to touch document_blobs
# ---------------------------------------------------------------------------
class DocumentStore:
    """Storage backend for document bytes, keyed by storage_key. Every
    service function below goes through one of these rather than
    executing SQL against document_blobs itself, so a different
    implementation (e.g. S3) can replace DatabaseBlobStore later
    without any service function changing."""

    def put(self, storage_key: str, data: bytes) -> None:
        raise NotImplementedError

    def get(self, storage_key: str) -> bytes:
        raise NotImplementedError

    def delete(self, storage_key: str) -> None:
        raise NotImplementedError


class DatabaseBlobStore(DocumentStore):
    """Stores bytes as ~CHUNK_SIZE rows in document_blobs, in the same
    database get_conn() opens by default."""

    def __init__(self, db_path: Path = DEFAULT_DB_PATH):
        self.db_path = db_path

    def put(self, storage_key: str, data: bytes) -> None:
        conn = get_conn(self.db_path)
        try:
            conn.execute("DELETE FROM document_blobs WHERE storage_key = ?", (storage_key,))
            conn.commit()
            for chunk_index, offset in enumerate(range(0, len(data), CHUNK_SIZE)):
                chunk = data[offset : offset + CHUNK_SIZE]
                conn.execute(
                    "INSERT INTO document_blobs (storage_key, chunk_index, data) VALUES (?, ?, ?)",
                    (storage_key, chunk_index, chunk),
                )
                conn.commit()
        finally:
            conn.close()

    def get(self, storage_key: str) -> bytes:
        conn = get_conn(self.db_path)
        try:
            rows = conn.execute(
                "SELECT data FROM document_blobs WHERE storage_key = ? ORDER BY chunk_index",
                (storage_key,),
            ).fetchall()
            return b"".join(bytes(r[0]) for r in rows)
        finally:
            conn.close()

    def delete(self, storage_key: str) -> None:
        conn = get_conn(self.db_path)
        try:
            conn.execute("DELETE FROM document_blobs WHERE storage_key = ?", (storage_key,))
            conn.commit()
        finally:
            conn.close()


# ---------------------------------------------------------------------------
# Categories
# ---------------------------------------------------------------------------
def seed_default_categories(db_path: Path = DEFAULT_DB_PATH, actor_email: str = "system") -> None:
    """Creates the nine categories in DEFAULT_CATEGORIES, but only if
    doc_categories is currently empty -- safe to call more than once.

    Not wired into any call site yet. It deliberately does not run
    through the normal admin check every other mutation below does:
    this is a one-time bootstrap step that, by definition, may run
    before any doc_users row or DOC_REPO_ADMIN_EMAILS secret exists.
    Whichever caller wires this in later (an init script, an admin
    "reset categories" button) is responsible for only invoking it in
    an already-authorized context.
    """
    conn = get_conn(db_path)
    try:
        if conn.execute("SELECT COUNT(*) FROM doc_categories").fetchone()[0]:
            return
        now = _now()
        for sort_order, name in enumerate(DEFAULT_CATEGORIES):
            conn.execute(
                """INSERT INTO doc_categories (name, sort_order, is_active, created_at, created_by, updated_at, updated_by)
                   VALUES (?, ?, 1, ?, ?, ?, ?)""",
                (name, sort_order, now, actor_email, now, actor_email),
            )
        conn.commit()
    finally:
        conn.close()


def create_category(actor_email: str, name: str, description: str = "", sort_order: int = 0, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> int:
    actor_email = _require_actor(actor_email)
    name = (name or "").strip()
    if not name:
        raise ValueError("Category name can't be blank.")
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "create_category", detail=name)
        now = _now()
        try:
            cur = conn.execute(
                """INSERT INTO doc_categories (name, description, sort_order, is_active, created_at, created_by, updated_at, updated_by)
                   VALUES (?, ?, ?, 1, ?, ?, ?, ?)""",
                (name, description, sort_order, now, actor_email, now, actor_email),
            )
            conn.commit()
        except _INTEGRITY_ERRORS:
            raise ValueError(f"A category named '{name}' already exists.")
        category_id = cur.lastrowid
        _audit(conn, actor_email, "create_category", category_id=category_id, detail=name)
        conn.commit()
        return category_id
    finally:
        conn.close()


def rename_category(actor_email: str, category_id: int, new_name: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    new_name = (new_name or "").strip()
    if not new_name:
        raise ValueError("Category name can't be blank.")
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "rename_category", category_id=category_id, detail=new_name)
        now = _now()
        try:
            cur = conn.execute(
                "UPDATE doc_categories SET name = ?, updated_at = ?, updated_by = ? WHERE id = ?",
                (new_name, now, actor_email, category_id),
            )
            if cur.rowcount == 0:
                raise ValueError(f"No such category: {category_id}")
            conn.commit()
        except _INTEGRITY_ERRORS:
            raise ValueError(f"A category named '{new_name}' already exists.")
        _audit(conn, actor_email, "rename_category", category_id=category_id, detail=new_name)
        conn.commit()
    finally:
        conn.close()


def reorder_categories(actor_email: str, ordered_category_ids: list[int], db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "reorder_categories", detail=str(ordered_category_ids))
        for sort_order, category_id in enumerate(ordered_category_ids):
            conn.execute("UPDATE doc_categories SET sort_order = ? WHERE id = ?", (sort_order, category_id))
        conn.commit()
        _audit(conn, actor_email, "reorder_categories", detail=str(ordered_category_ids))
        conn.commit()
    finally:
        conn.close()


def deactivate_category(actor_email: str, category_id: int, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "deactivate_category", category_id=category_id)
        now = _now()
        cur = conn.execute(
            "UPDATE doc_categories SET is_active = 0, updated_at = ?, updated_by = ? WHERE id = ?",
            (now, actor_email, category_id),
        )
        if cur.rowcount == 0:
            raise ValueError(f"No such category: {category_id}")
        conn.commit()
        _audit(conn, actor_email, "deactivate_category", category_id=category_id)
        conn.commit()
    finally:
        conn.close()


def reactivate_category(actor_email: str, category_id: int, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "reactivate_category", category_id=category_id)
        now = _now()
        cur = conn.execute(
            "UPDATE doc_categories SET is_active = 1, updated_at = ?, updated_by = ? WHERE id = ?",
            (now, actor_email, category_id),
        )
        if cur.rowcount == 0:
            raise ValueError(f"No such category: {category_id}")
        conn.commit()
        _audit(conn, actor_email, "reactivate_category", category_id=category_id)
        conn.commit()
    finally:
        conn.close()


def list_categories(actor_email: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> list[dict]:
    """Categories visible to actor_email: every category for an admin
    (active or not), or only the active categories actor_email has an
    access row on (any level) otherwise -- a category with no access
    rows is invisible to non-admins, by design."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        if _is_admin(conn, actor_email, secrets_getter):
            rows = conn.execute("SELECT * FROM doc_categories ORDER BY sort_order, name COLLATE NOCASE").fetchall()
        else:
            rows = conn.execute(
                """SELECT c.* FROM doc_categories c
                   JOIN doc_category_access a ON a.category_id = c.id
                   WHERE a.subject = ? AND c.is_active = 1
                   ORDER BY c.sort_order, c.name COLLATE NOCASE""",
                (actor_email,),
            ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_access_level(actor_email: str, category_id: int, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> str | None:
    """Read-only, side-effect-free (Phase 3 addition): actor_email's
    access level on category_id -- the literal string "admin" if
    they're an admin (full access everywhere; admins have no
    doc_category_access row of their own to report), otherwise their
    level ("view"/"contribute"/"manage") from doc_category_access, or
    None if they have no access row on this category at all. No audit
    row is written -- same as is_admin(), this is a visibility check
    used to decide what to show, not a mutation or a document access.
    """
    actor_email = _normalize_email(actor_email)
    if not actor_email:
        return None
    if is_admin(actor_email, db_path=db_path, secrets_getter=secrets_getter):
        return "admin"
    conn = get_conn(db_path)
    try:
        return _access_level(conn, actor_email, category_id)
    finally:
        conn.close()


def categories_with_min_level(actor_email: str, min_level: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> list[dict]:
    """Read-only (Phase 3 addition): the subset of
    list_categories(actor_email) where the actor has at least
    min_level access. An admin gets every category list_categories()
    already returns unfiltered -- admins aren't gated by
    doc_category_access at all, so there's no per-category level to
    compare. Useful for deciding what to show, e.g. an upload
    category dropdown (min_level=LEVEL_CONTRIBUTE) or whether a
    "show archived" control should appear at all (min_level=LEVEL_MANAGE).
    """
    categories = list_categories(actor_email, db_path=db_path, secrets_getter=secrets_getter)
    if is_admin(actor_email, db_path=db_path, secrets_getter=secrets_getter):
        return categories
    conn = get_conn(db_path)
    try:
        return [c for c in categories if _has_level(conn, actor_email, c["id"], min_level)]
    finally:
        conn.close()


def find_duplicate_by_sha256(actor_email: str, category_id: int, sha256: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> dict | None:
    """Read-only (Phase 3 addition): the active document in
    category_id whose current version has this exact sha256, if any --
    used to warn about (never to block) a likely duplicate upload.
    Requires at least view access on the category, same as any other
    read in this module; a denial is audited like every other one."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        if not _authorized_for_category(conn, actor_email, category_id, LEVEL_VIEW, secrets_getter):
            _deny(conn, actor_email, "find_duplicate_by_sha256", category_id=category_id)
        row = conn.execute(
            """SELECT d.* FROM documents d
               JOIN document_versions v ON v.id = d.current_version_id
               WHERE d.category_id = ? AND d.status = ? AND v.sha256 = ?
               LIMIT 1""",
            (category_id, STATUS_ACTIVE, sha256),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Category access
# ---------------------------------------------------------------------------
def grant_category_access(actor_email: str, category_id: int, subject_email: str, level: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    subject_email = _normalize_email(subject_email)
    if not subject_email:
        raise ValueError("Subject email can't be blank.")
    if level not in VALID_LEVELS:
        raise ValueError(f"Level must be one of {sorted(VALID_LEVELS)}.")
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "grant_category_access", category_id=category_id, detail=f"{subject_email}:{level}")
        existing = conn.execute(
            "SELECT id FROM doc_category_access WHERE category_id = ? AND subject = ?",
            (category_id, subject_email),
        ).fetchone()
        if existing:
            conn.execute("UPDATE doc_category_access SET level = ? WHERE id = ?", (level, existing[0]))
        else:
            conn.execute(
                "INSERT INTO doc_category_access (category_id, subject, level) VALUES (?, ?, ?)",
                (category_id, subject_email, level),
            )
        conn.commit()
        _audit(conn, actor_email, "grant_category_access", category_id=category_id, detail=f"{subject_email}:{level}")
        conn.commit()
    finally:
        conn.close()


def revoke_category_access(actor_email: str, category_id: int, subject_email: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    subject_email = _normalize_email(subject_email)
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "revoke_category_access", category_id=category_id, detail=subject_email)
        conn.execute(
            "DELETE FROM doc_category_access WHERE category_id = ? AND subject = ?",
            (category_id, subject_email),
        )
        conn.commit()
        _audit(conn, actor_email, "revoke_category_access", category_id=category_id, detail=subject_email)
        conn.commit()
    finally:
        conn.close()


def list_category_access(actor_email: str, category_id: int, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> list[dict]:
    """Read-only (Phase 4 addition): every doc_category_access row
    (subject + level) for category_id, admin only. No audit row for a
    successful read -- a lookup, not a mutation -- but a denial is
    audited like every other one in this module."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "list_category_access", category_id=category_id)
        rows = conn.execute(
            "SELECT * FROM doc_category_access WHERE category_id = ? ORDER BY subject", (category_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# doc_users
# ---------------------------------------------------------------------------
def create_doc_user(actor_email: str, email: str, role: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    email = _normalize_email(email)
    if not email:
        raise ValueError("Email can't be blank.")
    if role not in VALID_ROLES:
        raise ValueError(f"Role must be one of {sorted(VALID_ROLES)}.")
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "create_doc_user", detail=f"{email}:{role}")
        now = _now()
        try:
            conn.execute(
                "INSERT INTO doc_users (email, role, is_active, created_at, created_by) VALUES (?, ?, 1, ?, ?)",
                (email, role, now, actor_email),
            )
            conn.commit()
        except _INTEGRITY_ERRORS:
            raise ValueError(f"'{email}' already has a doc_users entry.")
        _audit(conn, actor_email, "create_doc_user", detail=f"{email}:{role}")
        conn.commit()
    finally:
        conn.close()


def update_doc_user_role(actor_email: str, email: str, role: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    email = _normalize_email(email)
    if role not in VALID_ROLES:
        raise ValueError(f"Role must be one of {sorted(VALID_ROLES)}.")
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "update_doc_user_role", detail=f"{email}:{role}")
        cur = conn.execute("UPDATE doc_users SET role = ? WHERE email = ?", (role, email))
        if cur.rowcount == 0:
            raise ValueError(f"No such doc_users entry: {email}")
        conn.commit()
        _audit(conn, actor_email, "update_doc_user_role", detail=f"{email}:{role}")
        conn.commit()
    finally:
        conn.close()


def deactivate_doc_user(actor_email: str, email: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    email = _normalize_email(email)
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "deactivate_doc_user", detail=email)
        cur = conn.execute("UPDATE doc_users SET is_active = 0 WHERE email = ?", (email,))
        if cur.rowcount == 0:
            raise ValueError(f"No such doc_users entry: {email}")
        conn.commit()
        _audit(conn, actor_email, "deactivate_doc_user", detail=email)
        conn.commit()
    finally:
        conn.close()


def reactivate_doc_user(actor_email: str, email: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    actor_email = _require_actor(actor_email)
    email = _normalize_email(email)
    conn = get_conn(db_path)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "reactivate_doc_user", detail=email)
        cur = conn.execute("UPDATE doc_users SET is_active = 1 WHERE email = ?", (email,))
        if cur.rowcount == 0:
            raise ValueError(f"No such doc_users entry: {email}")
        conn.commit()
        _audit(conn, actor_email, "reactivate_doc_user", detail=email)
        conn.commit()
    finally:
        conn.close()


def list_doc_users(actor_email: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> list[dict]:
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "list_doc_users")
        rows = conn.execute("SELECT * FROM doc_users ORDER BY email").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Documents
# ---------------------------------------------------------------------------
def add_document(
    actor_email: str,
    category_id: int,
    title: str,
    data: bytes,
    filename: str,
    description: str = "",
    keywords: str = "",
    change_note: str = "",
    db_path: Path = DEFAULT_DB_PATH,
    secrets_getter=None,
    store: DocumentStore | None = None,
) -> str:
    """Creates a new document with its first version. Requires contribute
    (or manage/admin) access on category_id.

    Crash safety: writes happen in this order -- blob chunks, then the
    version row, then the document row (whose current_version_id points
    at that same version from the moment it exists), then the audit
    row. If this is interrupted after the blob chunks but before the
    version row, there are orphan chunks and nothing else -- no
    document row was ever created to point at them.
    """
    actor_email = _require_actor(actor_email)
    title = (title or "").strip()
    if not title:
        raise ValueError("Title can't be blank.")
    conn = get_conn(db_path)
    try:
        if not _authorized_for_category(conn, actor_email, category_id, LEVEL_CONTRIBUTE, secrets_getter):
            _deny(conn, actor_email, "add_document", category_id=category_id, detail=title)
        if not conn.execute("SELECT 1 FROM doc_categories WHERE id = ?", (category_id,)).fetchone():
            raise ValueError(f"No such category: {category_id}")

        name, extension = validate_upload(filename, data)
        document_id = uuid.uuid4().hex
        version_id = uuid.uuid4().hex
        version_no = 1
        sha256 = hashlib.sha256(data).hexdigest()
        storage_key = f"documents/{document_id}/v{version_no}/{sha256}"
        now = _now()

        store = store or DatabaseBlobStore(db_path)
        store.put(storage_key, data)

        conn.execute(
            """INSERT INTO document_versions (
                id, document_id, version_no, filename, extension, mime_type,
                size_bytes, sha256, storage_key, change_note, uploaded_at, uploaded_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (version_id, document_id, version_no, name, extension, _guess_mime_type(extension),
             len(data), sha256, storage_key, change_note, now, actor_email),
        )
        conn.commit()

        conn.execute(
            """INSERT INTO documents (
                id, category_id, title, description, keywords, status,
                current_version_id, created_at, created_by, updated_at, updated_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (document_id, category_id, title, description, keywords, STATUS_ACTIVE,
             version_id, now, actor_email, now, actor_email),
        )
        conn.commit()

        _audit(conn, actor_email, "add_document", document_id=document_id, version_id=version_id, category_id=category_id, detail=title)
        conn.commit()
        return document_id
    finally:
        conn.close()


def replace_document(
    actor_email: str,
    document_id: str,
    data: bytes,
    filename: str,
    change_note: str = "",
    db_path: Path = DEFAULT_DB_PATH,
    secrets_getter=None,
    store: DocumentStore | None = None,
) -> str:
    """Adds a new version of an existing document without touching its
    old versions. Requires contribute (or manage/admin) access on the
    document's category. Same write order/crash-safety guarantee as
    add_document(): blob chunks, then the new version row, then
    documents.current_version_id is updated to point at it, then the
    audit row. If interrupted before the UPDATE, the document still
    points at its previous (complete) version.
    """
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        if doc is None:
            raise ValueError(f"No such document: {document_id}")
        category_id = doc["category_id"]
        if not _authorized_for_category(conn, actor_email, category_id, LEVEL_CONTRIBUTE, secrets_getter):
            _deny(conn, actor_email, "replace_document", document_id=document_id, category_id=category_id)

        name, extension = validate_upload(filename, data)
        last_version_no = conn.execute(
            "SELECT MAX(version_no) FROM document_versions WHERE document_id = ?", (document_id,)
        ).fetchone()[0] or 0
        version_no = last_version_no + 1
        version_id = uuid.uuid4().hex
        sha256 = hashlib.sha256(data).hexdigest()
        storage_key = f"documents/{document_id}/v{version_no}/{sha256}"
        now = _now()

        store = store or DatabaseBlobStore(db_path)
        store.put(storage_key, data)

        conn.execute(
            """INSERT INTO document_versions (
                id, document_id, version_no, filename, extension, mime_type,
                size_bytes, sha256, storage_key, change_note, uploaded_at, uploaded_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (version_id, document_id, version_no, name, extension, _guess_mime_type(extension),
             len(data), sha256, storage_key, change_note, now, actor_email),
        )
        conn.commit()

        conn.execute(
            "UPDATE documents SET current_version_id = ?, updated_at = ?, updated_by = ? WHERE id = ?",
            (version_id, now, actor_email, document_id),
        )
        conn.commit()

        _audit(conn, actor_email, "replace_document", document_id=document_id, version_id=version_id, category_id=category_id, detail=name)
        conn.commit()
        return version_id
    finally:
        conn.close()


def update_metadata(
    actor_email: str,
    document_id: str,
    title: str | None = None,
    description: str | None = None,
    keywords: str | None = None,
    db_path: Path = DEFAULT_DB_PATH,
    secrets_getter=None,
) -> None:
    """Updates whichever of title/description/keywords are given (None
    leaves that field unchanged). Requires contribute (or manage/admin)
    access on the document's category -- same level as uploading."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        if doc is None:
            raise ValueError(f"No such document: {document_id}")
        category_id = doc["category_id"]
        if not _authorized_for_category(conn, actor_email, category_id, LEVEL_CONTRIBUTE, secrets_getter):
            _deny(conn, actor_email, "update_metadata", document_id=document_id, category_id=category_id)

        updates = {}
        if title is not None:
            title = title.strip()
            if not title:
                raise ValueError("Title can't be blank.")
            updates["title"] = title
        if description is not None:
            updates["description"] = description
        if keywords is not None:
            updates["keywords"] = keywords
        if not updates:
            return

        now = _now()
        set_clause = ", ".join(f"{col} = ?" for col in updates) + ", updated_at = ?, updated_by = ?"
        conn.execute(
            f"UPDATE documents SET {set_clause} WHERE id = ?",
            (*updates.values(), now, actor_email, document_id),
        )
        conn.commit()
        _audit(conn, actor_email, "update_metadata", document_id=document_id, category_id=category_id, detail=", ".join(updates))
        conn.commit()
    finally:
        conn.close()


def archive_document(actor_email: str, document_id: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    """Hides a document from normal list_documents() results. Requires
    manage access (or admin) on the document's category -- a stronger
    bar than the contribute level that covers everyday uploads/edits,
    since archiving affects visibility for everyone with access to the
    category. There is no permanent-delete counterpart; see
    restore_document()."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        if doc is None:
            raise ValueError(f"No such document: {document_id}")
        category_id = doc["category_id"]
        if not _authorized_for_category(conn, actor_email, category_id, LEVEL_MANAGE, secrets_getter):
            _deny(conn, actor_email, "archive_document", document_id=document_id, category_id=category_id)
        now = _now()
        conn.execute(
            "UPDATE documents SET status = ?, archived_at = ?, archived_by = ?, updated_at = ?, updated_by = ? WHERE id = ?",
            (STATUS_ARCHIVED, now, actor_email, now, actor_email, document_id),
        )
        conn.commit()
        _audit(conn, actor_email, "archive_document", document_id=document_id, category_id=category_id)
        conn.commit()
    finally:
        conn.close()


def restore_document(actor_email: str, document_id: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> None:
    """Reverses archive_document(). Same manage-or-admin bar."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        if doc is None:
            raise ValueError(f"No such document: {document_id}")
        category_id = doc["category_id"]
        if not _authorized_for_category(conn, actor_email, category_id, LEVEL_MANAGE, secrets_getter):
            _deny(conn, actor_email, "restore_document", document_id=document_id, category_id=category_id)
        now = _now()
        conn.execute(
            "UPDATE documents SET status = ?, archived_at = NULL, archived_by = NULL, updated_at = ?, updated_by = ? WHERE id = ?",
            (STATUS_ACTIVE, now, actor_email, document_id),
        )
        conn.commit()
        _audit(conn, actor_email, "restore_document", document_id=document_id, category_id=category_id)
        conn.commit()
    finally:
        conn.close()


def list_documents(
    actor_email: str,
    category_id: int | None = None,
    search: str | None = None,
    file_type: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    status: str | None = STATUS_ACTIVE,
    page: int = 1,
    page_size: int = 50,
    db_path: Path = DEFAULT_DB_PATH,
    secrets_getter=None,
) -> list[dict]:
    """Metadata only -- joins document_versions for the current
    version's filename/extension/size/uploaded_at/uploaded_by/version_no
    (as current_extension/current_filename/current_size_bytes/
    current_uploaded_at/current_uploaded_by/current_version_no), never
    document_blobs.
    category filtering/search/file_type/date_from/date_to/status are
    all optional; status defaults to 'active' so archived documents
    stay out of the normal list (pass status=None for both, or
    'archived' for only archived). search escapes % and _ so they're
    matched literally rather than as LIKE wildcards.

    A category_id scoped to a category actor_email can't see raises
    PermissionDenied; an unscoped call instead just silently narrows to
    the categories actor_email (or everyone, if admin) can see -- that
    is normal browsing, not a denial worth logging.

    Archived documents (status='archived', or status=None which asks
    for both statuses at once) additionally require manage-level
    access on that document's category (or admin) -- view/contribute
    access only ever grants the active document. An explicit
    category_id with status='archived' and no manage access on it
    raises PermissionDenied, same as any other category_id the actor
    can't see at all; status=None instead just quietly drops archived
    rows from whichever categories are below manage level, the same
    way any other inaccessible category is quietly dropped.
    """
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        is_admin = _is_admin(conn, actor_email, secrets_getter)
        if category_id is not None and not is_admin and not _has_level(conn, actor_email, category_id, LEVEL_VIEW):
            _deny(conn, actor_email, "list_documents", category_id=category_id)
        if (
            category_id is not None and status == STATUS_ARCHIVED and not is_admin
            and not _has_level(conn, actor_email, category_id, LEVEL_MANAGE)
        ):
            _deny(conn, actor_email, "list_documents", category_id=category_id, detail="archived")

        clauses = []
        params: list = []

        if not is_admin:
            clauses.append("d.category_id IN (SELECT category_id FROM doc_category_access WHERE subject = ?)")
            params.append(actor_email)
            if status is None or status == STATUS_ARCHIVED:
                # Archived rows are only visible in categories where
                # this actor has manage-level access -- view/contribute
                # access only ever grants the active document. An
                # explicit category_id without manage access was
                # already denied above when status is exactly
                # STATUS_ARCHIVED; for status=None (both statuses at
                # once) this instead just quietly drops archived rows
                # from categories below manage level, same as any
                # other category this actor can't see.
                clauses.append(
                    "(d.status != ? OR d.category_id IN "
                    "(SELECT category_id FROM doc_category_access WHERE subject = ? AND level = ?))"
                )
                params.extend([STATUS_ARCHIVED, actor_email, LEVEL_MANAGE])
        if category_id is not None:
            clauses.append("d.category_id = ?")
            params.append(category_id)
        if status:
            clauses.append("d.status = ?")
            params.append(status)
        if search and search.strip():
            like_term = f"%{_escape_like(search.strip())}%"
            clauses.append("(d.title LIKE ? ESCAPE '\\' OR d.description LIKE ? ESCAPE '\\' OR d.keywords LIKE ? ESCAPE '\\')")
            params.extend([like_term, like_term, like_term])
        if file_type:
            clauses.append("v.extension = ?")
            params.append(file_type.lower().lstrip("."))
        if date_from:
            clauses.append("d.created_at >= ?")
            params.append(date_from)
        if date_to:
            clauses.append("d.created_at <= ?")
            params.append(date_to)

        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        page = max(page, 1)
        page_size = max(page_size, 1)
        offset = (page - 1) * page_size

        rows = conn.execute(
            f"""SELECT d.*, v.extension AS current_extension, v.filename AS current_filename,
                       v.size_bytes AS current_size_bytes, v.uploaded_at AS current_uploaded_at,
                       v.uploaded_by AS current_uploaded_by, v.version_no AS current_version_no
                FROM documents d
                LEFT JOIN document_versions v ON v.id = d.current_version_id
                {where}
                ORDER BY d.updated_at DESC, d.id DESC
                LIMIT ? OFFSET ?""",
            (*params, page_size, offset),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_document_detail(actor_email: str, document_id: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> dict | None:
    """The document row plus its full version history (metadata only --
    never reads document_blobs). Requires at least view access on the
    document's category -- or manage access (or admin) if the document
    is archived, the same rule list_documents() uses for archived
    listings: view/contribute access only ever grants an active
    document. Returns None if no such document exists."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        if doc is None:
            return None
        category_id = doc["category_id"]
        min_level = LEVEL_MANAGE if doc["status"] == STATUS_ARCHIVED else LEVEL_VIEW
        if not _authorized_for_category(conn, actor_email, category_id, min_level, secrets_getter):
            _deny(conn, actor_email, "get_document_detail", document_id=document_id, category_id=category_id)
        versions = conn.execute(
            """SELECT id, version_no, filename, extension, mime_type, size_bytes, sha256,
                      change_note, uploaded_at, uploaded_by
               FROM document_versions WHERE document_id = ? ORDER BY version_no DESC""",
            (document_id,),
        ).fetchall()
        result = dict(doc)
        result["versions"] = [dict(v) for v in versions]
        return result
    finally:
        conn.close()


def list_versions(actor_email: str, document_id: str, db_path: Path = DEFAULT_DB_PATH, secrets_getter=None) -> list[dict]:
    detail = get_document_detail(actor_email, document_id, db_path=db_path, secrets_getter=secrets_getter)
    if detail is None:
        raise ValueError(f"No such document: {document_id}")
    return detail["versions"]


def download_version(
    actor_email: str,
    document_id: str,
    version_id: str | None = None,
    db_path: Path = DEFAULT_DB_PATH,
    secrets_getter=None,
    store: DocumentStore | None = None,
) -> tuple[str, str, bytes]:
    """Returns (filename, mime_type, data) for a version -- the current
    one if version_id is omitted. Viewers may only download the
    current version; downloading any other version needs contribute
    (or manage/admin) access, same as uploading. If the document is
    archived, any download (current or old version) instead requires
    manage access (or admin) -- the same rule list_documents() and
    get_document_detail() use for archived documents; view/contribute
    access only ever grants an active document. Verifies the
    reassembled bytes' length and SHA-256 against what was recorded at
    upload time before returning them, raising StorageIntegrityError on
    a mismatch (e.g. an interrupted write left partial chunks)."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        doc = conn.execute("SELECT * FROM documents WHERE id = ?", (document_id,)).fetchone()
        if doc is None:
            raise ValueError(f"No such document: {document_id}")
        category_id = doc["category_id"]
        target_version_id = version_id or doc["current_version_id"]
        is_current = target_version_id == doc["current_version_id"]
        if doc["status"] == STATUS_ARCHIVED:
            min_level = LEVEL_MANAGE
        else:
            min_level = LEVEL_VIEW if is_current else LEVEL_CONTRIBUTE

        if not _authorized_for_category(conn, actor_email, category_id, min_level, secrets_getter):
            _deny(conn, actor_email, "download_version", document_id=document_id, version_id=target_version_id, category_id=category_id)

        version = conn.execute(
            "SELECT * FROM document_versions WHERE id = ? AND document_id = ?",
            (target_version_id, document_id),
        ).fetchone()
        if version is None:
            raise ValueError(f"No such version: {target_version_id}")

        store = store or DatabaseBlobStore(db_path)
        data = store.get(version["storage_key"])
        if len(data) != version["size_bytes"]:
            raise StorageIntegrityError(
                f"Version {target_version_id} is corrupted: expected {version['size_bytes']} bytes, got {len(data)}."
            )
        if hashlib.sha256(data).hexdigest() != version["sha256"]:
            raise StorageIntegrityError(f"Version {target_version_id} failed its integrity check (SHA-256 mismatch).")

        _audit(conn, actor_email, "download_version", document_id=document_id, version_id=target_version_id, category_id=category_id, detail=version["filename"])
        conn.commit()
        return version["filename"], version["mime_type"], data
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Audit log (read-only)
# ---------------------------------------------------------------------------
def list_audit_log(
    actor_email: str,
    actor_filter: str | None = None,
    action_filter: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    page: int = 1,
    page_size: int = 50,
    db_path: Path = DEFAULT_DB_PATH,
    secrets_getter=None,
) -> list[dict]:
    """Read-only (Phase 4 addition): doc_audit_log rows, newest first,
    admin only. actor_filter/action_filter are optional substring
    matches (LIKE, with % and _ escaped so they're matched literally,
    same as list_documents()'s search).

    date_from/date_to are calendar-day bounds (e.g. "2026-11-13"),
    both inclusive of every row anywhere in that whole day -- not a
    literal string comparison against the bound you passed. ts itself
    is stored as a full UTC timestamp (see _now()), e.g.
    "2026-11-13T21:21:23+00:00": a bare date string like "2026-11-13"
    is already a lexicographic prefix of every timestamp on that day,
    so "ts >= date_from" alone already includes that whole day with no
    further adjustment needed. The upper bound does need one: a plain
    "ts <= date_to" would exclude every row on date_to's own day
    (any real timestamp "2026-11-13T..." sorts AFTER the bare string
    "2026-11-13"), so this instead compares "ts < date_to + 1 day",
    which includes every moment of date_to's day and nothing from the
    day after.

    Reading the log is not itself audited -- same as browsing
    documents via list_documents() isn't -- but a denial is, like
    every other one in this module."""
    actor_email = _require_actor(actor_email)
    conn = get_conn(db_path)
    _set_row_factory(conn)
    try:
        if not _is_admin(conn, actor_email, secrets_getter):
            _deny(conn, actor_email, "list_audit_log")

        clauses = []
        params: list = []
        if actor_filter and actor_filter.strip():
            clauses.append("actor_email LIKE ? ESCAPE '\\'")
            params.append(f"%{_escape_like(actor_filter.strip())}%")
        if action_filter and action_filter.strip():
            clauses.append("action LIKE ? ESCAPE '\\'")
            params.append(f"%{_escape_like(action_filter.strip())}%")
        if date_from:
            clauses.append("ts >= ?")
            params.append(date_from)
        if date_to:
            clauses.append("ts < ?")
            params.append((date.fromisoformat(date_to) + timedelta(days=1)).isoformat())

        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        page = max(page, 1)
        page_size = max(page_size, 1)
        offset = (page - 1) * page_size

        rows = conn.execute(
            f"SELECT * FROM doc_audit_log {where} ORDER BY id DESC LIMIT ? OFFSET ?",
            (*params, page_size, offset),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()
