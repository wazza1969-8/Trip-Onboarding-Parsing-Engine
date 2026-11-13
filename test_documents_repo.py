"""
Automated tests for documents_repo.py (Document Repository -- Phase 1,
backend foundation only):
  - DatabaseBlobStore chunk round-trip at size boundaries
  - upload validation: size, extension allowlist, content-signature
    checks (including macro-enabled OOXML detection), filename sanitizing
  - the authorization matrix: admin, contributor (with/without access),
    viewer (with/without access), unknown user, empty actor
  - versioning: replace keeps old versions, viewers can't download old
    versions, current version is correct
  - archive/restore, including that archived items are hidden from the
    normal list
  - search wildcard escaping, filters, pagination, and that
    list_documents() never reads document_blobs
  - a doc_audit_log row for every mutation/download and every denied
    attempt
  - the DOC_REPO_ADMIN_EMAILS secret-based admin working with zero
    doc_users rows
  - is_admin(): the side-effect-free admin probe used by app.py to
    decide whether to show the Document Repository tile -- never
    creates tables, never writes audit rows, never raises

All tests use a fresh temporary SQLite file (never the production
Turso database, never a shared file) and store only fake/synthetic
file content -- no real customer data.

Run with:
    python3 test_documents_repo.py
or:
    pytest test_documents_repo.py -v
"""
import hashlib
import io
import os
import sqlite3
import tempfile
import unittest
import zipfile

import documents_repo as dr

ADMIN = "admin@example.test"
CONTRIBUTOR = "contributor@example.test"
VIEWER = "viewer@example.test"
OUTSIDER = "outsider@example.test"


def admin_secrets():
    return lambda: {"DOC_REPO_ADMIN_EMAILS": ADMIN}


def make_pdf(total_size=None):
    header = b"%PDF-1.4\n%fake pdf content for tests\n"
    if total_size is None:
        return header
    assert total_size >= len(header)
    return header + b"0" * (total_size - len(header))


def make_png():
    return b"\x89PNG\r\n\x1a\n" + b"fake png body"


def make_jpeg():
    return b"\xff\xd8\xff" + b"fake jpeg body"


def make_ole2():
    return b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"fake ole2 body"


def make_ooxml_zip(macro=False, missing_content_types=False):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        if not missing_content_types:
            zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", "<document/>")
        if macro:
            zf.writestr("word/vbaProject.bin", b"fake macro project bytes")
    return buf.getvalue()


def audit_rows(db_path, action=None):
    conn = sqlite3.connect(db_path)
    try:
        if action:
            return conn.execute(
                "SELECT * FROM doc_audit_log WHERE action = ? ORDER BY id", (action,)
            ).fetchall()
        return conn.execute("SELECT * FROM doc_audit_log ORDER BY id").fetchall()
    finally:
        conn.close()


class TempDbTestCase(unittest.TestCase):
    def setUp(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.db_path = path

    def tearDown(self):
        os.remove(self.db_path)


# ---------------------------------------------------------------------------
# DatabaseBlobStore chunk round-trip
# ---------------------------------------------------------------------------
class ChunkStoreTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.store = dr.DatabaseBlobStore(self.db_path)

    def test_round_trip_empty(self):
        self.store.put("k-empty", b"")
        self.assertEqual(self.store.get("k-empty"), b"")

    def test_round_trip_one_byte(self):
        self.store.put("k-1b", b"X")
        self.assertEqual(self.store.get("k-1b"), b"X")

    def test_round_trip_exactly_one_chunk(self):
        data = os.urandom(dr.CHUNK_SIZE)
        self.store.put("k-1chunk", data)
        self.assertEqual(self.store.get("k-1chunk"), data)
        conn = sqlite3.connect(self.db_path)
        count = conn.execute(
            "SELECT COUNT(*) FROM document_blobs WHERE storage_key = ?", ("k-1chunk",)
        ).fetchone()[0]
        conn.close()
        self.assertEqual(count, 1)

    def test_round_trip_one_chunk_plus_one_byte(self):
        data = os.urandom(dr.CHUNK_SIZE + 1)
        self.store.put("k-1chunk-plus", data)
        self.assertEqual(self.store.get("k-1chunk-plus"), data)
        conn = sqlite3.connect(self.db_path)
        count = conn.execute(
            "SELECT COUNT(*) FROM document_blobs WHERE storage_key = ?", ("k-1chunk-plus",)
        ).fetchone()[0]
        conn.close()
        self.assertEqual(count, 2)

    def test_put_overwrites_previous_chunks(self):
        self.store.put("k-overwrite", os.urandom(dr.CHUNK_SIZE + 10))
        self.store.put("k-overwrite", b"shorter")
        self.assertEqual(self.store.get("k-overwrite"), b"shorter")

    def test_delete_removes_chunks(self):
        self.store.put("k-del", b"some bytes")
        self.store.delete("k-del")
        self.assertEqual(self.store.get("k-del"), b"")


# ---------------------------------------------------------------------------
# Upload validation
# ---------------------------------------------------------------------------
class ValidationTests(unittest.TestCase):
    def test_empty_file_rejected(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("a.pdf", b"")

    def test_exactly_max_size_accepted(self):
        data = make_pdf(dr.MAX_FILE_SIZE_BYTES)
        name, ext = dr.validate_upload("a.pdf", data)
        self.assertEqual(name, "a.pdf")
        self.assertEqual(ext, "pdf")

    def test_max_size_plus_one_rejected(self):
        data = make_pdf(dr.MAX_FILE_SIZE_BYTES + 1)
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("a.pdf", data)

    def test_disallowed_extension_rejected(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("malware.exe", b"whatever bytes")

    def test_spoofed_pdf_extension_rejected(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("evil.pdf", b"this is not actually a pdf")

    def test_valid_pdf_accepted(self):
        dr.validate_upload("a.pdf", make_pdf())

    def test_valid_png_accepted(self):
        dr.validate_upload("a.png", make_png())

    def test_valid_jpeg_accepted(self):
        dr.validate_upload("a.jpg", make_jpeg())
        dr.validate_upload("a.jpeg", make_jpeg())

    def test_spoofed_png_extension_rejected(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("fake.png", b"not a png")

    def test_valid_ole2_doc_accepted(self):
        dr.validate_upload("a.doc", make_ole2())

    def test_spoofed_ole2_extension_rejected(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("fake.xls", b"not an ole2 file")

    def test_valid_docx_accepted(self):
        name, ext = dr.validate_upload("a.docx", make_ooxml_zip())
        self.assertEqual(ext, "docx")

    def test_docx_missing_content_types_rejected(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("a.docx", make_ooxml_zip(missing_content_types=True))

    def test_macro_enabled_docx_rejected(self):
        with self.assertRaises(dr.ValidationError) as ctx:
            dr.validate_upload("a.docx", make_ooxml_zip(macro=True))
        self.assertIn("Macro", str(ctx.exception))

    def test_not_a_zip_rejected_for_docx_extension(self):
        with self.assertRaises(dr.ValidationError):
            dr.validate_upload("a.docx", b"plain text, not a zip at all")

    def test_filename_sanitizing_strips_path(self):
        name, _ = dr.validate_upload("../../etc/passwd/evil.pdf", make_pdf())
        self.assertEqual(name, "evil.pdf")

    def test_filename_sanitizing_strips_control_chars(self):
        name, _ = dr.validate_upload("evil\x00\x01name.pdf", make_pdf())
        self.assertNotIn("\x00", name)
        self.assertNotIn("\x01", name)

    def test_filename_blank_after_sanitizing_raises(self):
        with self.assertRaises(dr.ValidationError):
            dr.sanitize_filename("   ")

    def test_filename_length_limited(self):
        long_name = ("a" * 300) + ".pdf"
        name, _ = dr.validate_upload(long_name, make_pdf())
        self.assertLessEqual(len(name), 255)


# ---------------------------------------------------------------------------
# Authorization matrix
# ---------------------------------------------------------------------------
class AuthorizationMatrixTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(
            ADMIN, "Test Category", db_path=self.db_path, secrets_getter=self.secrets_getter
        )

    def _grant(self, email, level):
        dr.grant_category_access(
            ADMIN, self.category_id, email, level, db_path=self.db_path, secrets_getter=self.secrets_getter
        )

    def test_admin_can_add_document(self):
        doc_id = dr.add_document(
            ADMIN, self.category_id, "Admin Doc", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertTrue(doc_id)

    def test_empty_actor_denied(self):
        before = len(audit_rows(self.db_path))
        with self.assertRaises(dr.PermissionDenied):
            dr.add_document("", self.category_id, "T", make_pdf(), "a.pdf", db_path=self.db_path)
        # No identity to attribute an audit row to -- nothing new should be logged.
        self.assertEqual(len(audit_rows(self.db_path)), before)

    def test_unknown_user_denied_and_audited(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.add_document(
                OUTSIDER, self.category_id, "T", make_pdf(), "a.pdf",
                db_path=self.db_path, secrets_getter=self.secrets_getter,
            )
        rows = audit_rows(self.db_path, "add_document_denied")
        self.assertEqual(len(rows), 1)

    def test_contributor_without_access_denied(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.add_document(
                CONTRIBUTOR, self.category_id, "T", make_pdf(), "a.pdf",
                db_path=self.db_path, secrets_getter=self.secrets_getter,
            )

    def test_contributor_with_access_can_add_document(self):
        self._grant(CONTRIBUTOR, dr.LEVEL_CONTRIBUTE)
        doc_id = dr.add_document(
            CONTRIBUTOR, self.category_id, "Contributor Doc", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertTrue(doc_id)

    def test_viewer_with_access_cannot_add_document(self):
        self._grant(VIEWER, dr.LEVEL_VIEW)
        with self.assertRaises(dr.PermissionDenied):
            dr.add_document(
                VIEWER, self.category_id, "T", make_pdf(), "a.pdf",
                db_path=self.db_path, secrets_getter=self.secrets_getter,
            )

    def test_viewer_with_access_can_list_and_download_current(self):
        self._grant(VIEWER, dr.LEVEL_VIEW)
        doc_id = dr.add_document(
            ADMIN, self.category_id, "Viewable Doc", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        docs = dr.list_documents(VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(docs), 1)
        filename, mime, data = dr.download_version(
            VIEWER, doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter
        )
        self.assertEqual(filename, "a.pdf")
        self.assertEqual(data, make_pdf())

    def test_viewer_without_access_cannot_list_specific_category(self):
        dr.add_document(
            ADMIN, self.category_id, "Hidden Doc", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        with self.assertRaises(dr.PermissionDenied):
            dr.list_documents(
                VIEWER, category_id=self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter
            )

    def test_viewer_without_access_general_browse_is_empty_not_denied(self):
        dr.add_document(
            ADMIN, self.category_id, "Hidden Doc", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        docs = dr.list_documents(VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(docs, [])

    def test_category_with_no_access_rows_invisible_to_non_admin(self):
        other = dr.create_category(ADMIN, "No Access Category", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.add_document(
            ADMIN, other, "Secret Doc", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        cats = dr.list_categories(CONTRIBUTOR, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(cats, [])

    def test_category_management_admin_only(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.rename_category(CONTRIBUTOR, self.category_id, "New Name", db_path=self.db_path, secrets_getter=self.secrets_getter)
        with self.assertRaises(dr.PermissionDenied):
            dr.deactivate_category(CONTRIBUTOR, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        with self.assertRaises(dr.PermissionDenied):
            dr.grant_category_access(CONTRIBUTOR, self.category_id, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_doc_user_management_admin_only(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.create_doc_user(CONTRIBUTOR, "new@example.test", dr.ROLE_VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.create_doc_user(ADMIN, "new@example.test", dr.ROLE_VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        with self.assertRaises(dr.PermissionDenied):
            dr.list_doc_users(CONTRIBUTOR, db_path=self.db_path, secrets_getter=self.secrets_getter)
        users = dr.list_doc_users(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(users), 1)

    def test_doc_user_role_grants_admin_rights(self):
        # role='admin' in doc_users, with no DOC_REPO_ADMIN_EMAILS match, must also work.
        no_secret_getter = lambda: {}
        dr.create_doc_user(ADMIN, "second-admin@example.test", dr.ROLE_ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        doc_id = dr.add_document(
            "second-admin@example.test", self.category_id, "T", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=no_secret_getter,
        )
        self.assertTrue(doc_id)

    def test_deactivated_doc_user_loses_admin_rights(self):
        dr.create_doc_user(ADMIN, "second-admin@example.test", dr.ROLE_ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.deactivate_doc_user(ADMIN, "second-admin@example.test", db_path=self.db_path, secrets_getter=self.secrets_getter)
        no_secret_getter = lambda: {}
        with self.assertRaises(dr.PermissionDenied):
            dr.add_document(
                "second-admin@example.test", self.category_id, "T", make_pdf(), "a.pdf",
                db_path=self.db_path, secrets_getter=no_secret_getter,
            )


# ---------------------------------------------------------------------------
# Versioning
# ---------------------------------------------------------------------------
class VersioningTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Versioned", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.doc_id = dr.add_document(
            CONTRIBUTOR, self.category_id, "Versioned Doc", make_pdf(1000), "v1.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )

    def test_replace_keeps_old_versions_and_updates_current(self):
        v2 = dr.replace_document(CONTRIBUTOR, self.doc_id, make_pdf(2000), "v2.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        v3 = dr.replace_document(CONTRIBUTOR, self.doc_id, make_pdf(3000), "v3.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)

        versions = dr.list_versions(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(versions), 3)
        version_nos = sorted(v["version_no"] for v in versions)
        self.assertEqual(version_nos, [1, 2, 3])

        detail = dr.get_document_detail(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(detail["current_version_id"], v3)

        filename, _, data = dr.download_version(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(filename, "v3.pdf")
        self.assertEqual(data, make_pdf(3000))

    def test_viewer_cannot_download_old_version(self):
        v2 = dr.replace_document(CONTRIBUTOR, self.doc_id, make_pdf(2000), "v2.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        old_version_id = [v for v in dr.list_versions(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter) if v["version_no"] == 1][0]["id"]
        with self.assertRaises(dr.PermissionDenied):
            dr.download_version(VIEWER, self.doc_id, version_id=old_version_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        # Current version is still fine for the viewer.
        dr.download_version(VIEWER, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_contributor_can_download_old_version(self):
        dr.replace_document(CONTRIBUTOR, self.doc_id, make_pdf(2000), "v2.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        old_version_id = [v for v in dr.list_versions(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter) if v["version_no"] == 1][0]["id"]
        filename, _, data = dr.download_version(CONTRIBUTOR, self.doc_id, version_id=old_version_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(filename, "v1.pdf")
        self.assertEqual(data, make_pdf(1000))

    def test_download_verifies_integrity(self):
        # Corrupt the stored bytes directly, bypassing the service layer.
        conn = sqlite3.connect(self.db_path)
        conn.execute("DELETE FROM document_blobs")
        conn.commit()
        conn.close()
        with self.assertRaises(dr.StorageIntegrityError):
            dr.download_version(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)


# ---------------------------------------------------------------------------
# Archive / restore
# ---------------------------------------------------------------------------
class ArchiveRestoreTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Archivable", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.doc_id = dr.add_document(
            ADMIN, self.category_id, "Archive Me", make_pdf(), "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )

    def test_archive_hides_from_normal_list(self):
        dr.archive_document(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        docs = dr.list_documents(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(docs, [])
        archived = dr.list_documents(ADMIN, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(archived), 1)

    def test_restore_brings_it_back(self):
        dr.archive_document(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.restore_document(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        docs = dr.list_documents(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(docs), 1)

    def test_contributor_cannot_archive(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.archive_document(CONTRIBUTOR, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_manage_level_can_archive(self):
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.archive_document(VIEWER, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        docs = dr.list_documents(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(docs, [])


# ---------------------------------------------------------------------------
# list_documents()'s archived-visibility rule: manage-level (or admin)
# only -- view/contribute access never grants archived rows.
# ---------------------------------------------------------------------------
class ArchivedListingAccessTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Archived Access", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.active_id = dr.add_document(
            ADMIN, self.category_id, "Active Doc", make_pdf(), "active.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.archived_id = dr.add_document(
            ADMIN, self.category_id, "Archived Doc", make_pdf(), "archived.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        dr.archive_document(ADMIN, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_explicit_archived_denied_for_viewer_with_category_id(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.list_documents(VIEWER, category_id=self.category_id, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_explicit_archived_denied_for_contributor_with_category_id(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.list_documents(CONTRIBUTOR, category_id=self.category_id, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_explicit_archived_allowed_for_manage_level(self):
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        docs = dr.list_documents(VIEWER, category_id=self.category_id, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual([d["id"] for d in docs], [self.archived_id])

    def test_explicit_archived_allowed_for_admin(self):
        docs = dr.list_documents(ADMIN, category_id=self.category_id, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual([d["id"] for d in docs], [self.archived_id])

    def test_mixed_status_none_hides_archived_from_viewer_but_shows_active(self):
        docs = dr.list_documents(VIEWER, category_id=self.category_id, status=None, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual([d["id"] for d in docs], [self.active_id])

    def test_mixed_status_none_hides_archived_from_contributor_but_shows_active(self):
        docs = dr.list_documents(CONTRIBUTOR, category_id=self.category_id, status=None, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual([d["id"] for d in docs], [self.active_id])

    def test_mixed_status_none_shows_both_for_manage_level(self):
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        docs = dr.list_documents(VIEWER, category_id=self.category_id, status=None, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual({d["id"] for d in docs}, {self.active_id, self.archived_id})

    def test_unscoped_archived_browse_empty_for_non_manage_user(self):
        # VIEWER has view access (not manage) on the only category with
        # an archived document -- an unscoped browse for archived
        # documents must come back empty, not an error.
        docs = dr.list_documents(VIEWER, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(docs, [])

    def test_unscoped_archived_browse_narrows_to_manage_level_categories(self):
        other_category = dr.create_category(ADMIN, "No Manage Here", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, other_category, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)
        other_doc = dr.add_document(ADMIN, other_category, "Other Doc", make_pdf(), "other.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.archive_document(ADMIN, other_doc, db_path=self.db_path, secrets_getter=self.secrets_getter)

        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        docs = dr.list_documents(VIEWER, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual([d["id"] for d in docs], [self.archived_id])

    def test_unknown_user_cannot_see_archived_anywhere(self):
        docs = dr.list_documents(OUTSIDER, status=dr.STATUS_ARCHIVED, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(docs, [])

    # -----------------------------------------------------------------
    # Phase 4 Part A: get_document_detail() / download_version() must
    # apply the same archived-requires-manage rule as list_documents(),
    # closing the "known by ID" bypass -- see each function's own
    # docstring.
    # -----------------------------------------------------------------
    def test_get_document_detail_denied_for_viewer_on_archived_doc(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.get_document_detail(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_get_document_detail_denied_for_contributor_on_archived_doc(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.get_document_detail(CONTRIBUTOR, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_get_document_detail_allowed_for_manage_level_on_archived_doc(self):
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        detail = dr.get_document_detail(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNotNone(detail)
        self.assertEqual(detail["id"], self.archived_id)

    def test_get_document_detail_allowed_for_admin_on_archived_doc(self):
        detail = dr.get_document_detail(ADMIN, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNotNone(detail)

    def test_get_document_detail_unchanged_for_viewer_on_active_doc(self):
        detail = dr.get_document_detail(VIEWER, self.active_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNotNone(detail)
        self.assertEqual(detail["id"], self.active_id)

    def test_download_version_denied_for_viewer_on_archived_doc_current_version(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.download_version(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_download_version_denied_for_contributor_on_archived_doc(self):
        # Contributors may download an ACTIVE document's old versions,
        # but archived status overrides that -- manage (or admin) only.
        with self.assertRaises(dr.PermissionDenied):
            dr.download_version(CONTRIBUTOR, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_download_version_allowed_for_manage_level_on_archived_doc(self):
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        filename, _, data = dr.download_version(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(filename, "archived.pdf")
        self.assertEqual(data, make_pdf())

    def test_download_version_allowed_for_admin_on_archived_doc(self):
        filename, _, data = dr.download_version(ADMIN, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(filename, "archived.pdf")

    def test_download_version_unchanged_for_viewer_on_active_doc(self):
        filename, _, data = dr.download_version(VIEWER, self.active_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(filename, "active.pdf")

    def test_restoring_document_restores_view_access(self):
        dr.restore_document(ADMIN, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        detail = dr.get_document_detail(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNotNone(detail)
        filename, _, _ = dr.download_version(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(filename, "archived.pdf")

    def test_denied_archived_access_does_not_change_error_message_shape(self):
        # The denial must look exactly like any other category-access
        # denial -- no wording that would let someone infer the
        # document's archived status from the error alone.
        try:
            dr.get_document_detail(VIEWER, self.archived_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
            self.fail("expected PermissionDenied")
        except dr.PermissionDenied as e:
            self.assertEqual(str(e), f"'{VIEWER}' is not allowed to get document detail.")


# ---------------------------------------------------------------------------
# Search / filters / pagination / blob isolation
# ---------------------------------------------------------------------------
class SearchFilterPaginationTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Browsable", db_path=self.db_path, secrets_getter=self.secrets_getter)

    def _add(self, title, filename="a.pdf", data=None, keywords=""):
        return dr.add_document(
            ADMIN, self.category_id, title, data or make_pdf(), filename, keywords=keywords,
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )

    def test_search_percent_literal(self):
        self._add("50% Off Report")
        self._add("Completely Unrelated")
        results = dr.list_documents(ADMIN, search="50%", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["title"], "50% Off Report")

    def test_search_underscore_literal(self):
        self._add("file_name_report")
        self._add("filexnamexreport")
        results = dr.list_documents(ADMIN, search="file_name", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["title"], "file_name_report")

    def test_search_matches_keywords(self):
        self._add("Some Title", keywords="special-keyword-xyz")
        results = dr.list_documents(ADMIN, search="special-keyword-xyz", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)

    def test_file_type_filter(self):
        self._add("PDF Doc", filename="a.pdf", data=make_pdf())
        self._add("PNG Doc", filename="a.png", data=make_png())
        results = dr.list_documents(ADMIN, file_type="png", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["title"], "PNG Doc")

    def test_current_version_upload_metadata_present(self):
        self._add("Uploaded Doc", filename="a.pdf")
        results = dr.list_documents(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["current_uploaded_by"], ADMIN)
        self.assertTrue(results[0]["current_uploaded_at"])
        self.assertEqual(results[0]["current_version_no"], 1)

    def test_current_version_upload_metadata_reflects_latest_replace(self):
        doc_id = self._add("Replaced Doc", filename="a.pdf")
        dr.grant_category_access(ADMIN, self.category_id, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.replace_document(CONTRIBUTOR, doc_id, make_pdf(2000), "b.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        results = dr.list_documents(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(results[0]["current_uploaded_by"], CONTRIBUTOR)
        self.assertEqual(results[0]["current_filename"], "b.pdf")
        self.assertEqual(results[0]["current_version_no"], 2)

    def test_date_filters_exclude_out_of_range(self):
        self._add("Today's Doc")
        future = "2999-01-01"
        results = dr.list_documents(ADMIN, date_from=future, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(results, [])
        past = "2000-01-01"
        results = dr.list_documents(ADMIN, date_from=past, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)

    def test_pagination_covers_all_without_duplicates(self):
        ids = {self._add(f"Doc {i}") for i in range(5)}
        seen = set()
        for page in (1, 2, 3):
            batch = dr.list_documents(ADMIN, page=page, page_size=2, db_path=self.db_path, secrets_getter=self.secrets_getter)
            seen.update(d["id"] for d in batch)
        self.assertEqual(seen, ids)

    def test_list_documents_never_reads_document_blobs(self):
        doc_id = self._add("Blob Isolation Doc")
        conn = sqlite3.connect(self.db_path)
        conn.execute("DELETE FROM document_blobs")
        conn.commit()
        conn.close()
        results = dr.list_documents(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["id"], doc_id)


# ---------------------------------------------------------------------------
# Audit log coverage
# ---------------------------------------------------------------------------
class AuditLogTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Audited", db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_category_mutations_are_audited(self):
        dr.rename_category(ADMIN, self.category_id, "Renamed", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.deactivate_category(ADMIN, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.reactivate_category(ADMIN, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        for action in ("create_category", "rename_category", "deactivate_category", "reactivate_category"):
            self.assertEqual(len(audit_rows(self.db_path, action)), 1, action)

    def test_access_changes_are_audited(self):
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.revoke_category_access(ADMIN, self.category_id, VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(audit_rows(self.db_path, "grant_category_access")), 1)
        self.assertEqual(len(audit_rows(self.db_path, "revoke_category_access")), 1)

    def test_user_changes_are_audited(self):
        dr.create_doc_user(ADMIN, VIEWER, dr.ROLE_VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.update_doc_user_role(ADMIN, VIEWER, dr.ROLE_CONTRIBUTOR, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.deactivate_doc_user(ADMIN, VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.reactivate_doc_user(ADMIN, VIEWER, db_path=self.db_path, secrets_getter=self.secrets_getter)
        for action in ("create_doc_user", "update_doc_user_role", "deactivate_doc_user", "reactivate_doc_user"):
            self.assertEqual(len(audit_rows(self.db_path, action)), 1, action)

    def test_document_mutations_and_download_are_audited(self):
        doc_id = dr.add_document(ADMIN, self.category_id, "T", make_pdf(), "a.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.replace_document(ADMIN, doc_id, make_pdf(2000), "b.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.update_metadata(ADMIN, doc_id, title="New Title", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.archive_document(ADMIN, doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.restore_document(ADMIN, doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.download_version(ADMIN, doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        for action in ("add_document", "replace_document", "update_metadata", "archive_document", "restore_document", "download_version"):
            self.assertEqual(len(audit_rows(self.db_path, action)), 1, action)

    def test_denied_attempts_are_audited(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.add_document(OUTSIDER, self.category_id, "T", make_pdf(), "a.pdf", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(audit_rows(self.db_path, "add_document_denied")), 1)


# ---------------------------------------------------------------------------
# Secret-based admin bootstrap
# ---------------------------------------------------------------------------
class SecretAdminBootstrapTests(TempDbTestCase):
    def test_admin_secret_works_with_zero_doc_users_rows(self):
        conn = sqlite3.connect(self.db_path)
        dr.ensure_schema(conn)
        count = conn.execute("SELECT COUNT(*) FROM doc_users").fetchone()[0]
        conn.close()
        self.assertEqual(count, 0)

        category_id = dr.create_category(ADMIN, "Bootstrap", db_path=self.db_path, secrets_getter=admin_secrets())
        self.assertTrue(category_id)


# ---------------------------------------------------------------------------
# is_admin() -- the side-effect-free probe app.py uses to decide
# whether to show the Document Repository tile.
# ---------------------------------------------------------------------------
class IsAdminTests(unittest.TestCase):
    def test_empty_actor_returns_false(self):
        self.assertFalse(dr.is_admin(""))

    def test_true_via_secret_without_touching_db(self):
        missing_path = os.path.join(tempfile.mkdtemp(), "never_created.db")
        self.assertFalse(os.path.exists(missing_path))
        self.assertTrue(dr.is_admin(ADMIN, db_path=missing_path, secrets_getter=admin_secrets()))
        # The secret matched before any connection was even attempted.
        self.assertFalse(os.path.exists(missing_path))

    def test_false_when_db_file_does_not_exist_at_all(self):
        missing_path = os.path.join(tempfile.mkdtemp(), "never_created.db")
        self.assertFalse(dr.is_admin(OUTSIDER, db_path=missing_path))
        # A visibility check must never conjure a database file into existence.
        self.assertFalse(os.path.exists(missing_path))

    def test_false_when_doc_users_table_does_not_exist(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            # A real SQLite file, but ensure_schema() was never run on it --
            # no doc_users table at all.
            sqlite3.connect(path).close()
            self.assertFalse(dr.is_admin(OUTSIDER, db_path=path))
            conn = sqlite3.connect(path)
            tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
            conn.close()
            self.assertEqual(tables, set())  # confirms is_admin created nothing
        finally:
            os.remove(path)

    def test_true_for_active_admin_doc_users_row(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            dr.create_doc_user(ADMIN, "second-admin@example.test", dr.ROLE_ADMIN, db_path=path, secrets_getter=admin_secrets())
            self.assertTrue(dr.is_admin("second-admin@example.test", db_path=path))
        finally:
            os.remove(path)

    def test_false_for_non_admin_role(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            dr.create_doc_user(ADMIN, CONTRIBUTOR, dr.ROLE_CONTRIBUTOR, db_path=path, secrets_getter=admin_secrets())
            self.assertFalse(dr.is_admin(CONTRIBUTOR, db_path=path))
        finally:
            os.remove(path)

    def test_false_for_inactive_admin(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            dr.create_doc_user(ADMIN, "second-admin@example.test", dr.ROLE_ADMIN, db_path=path, secrets_getter=admin_secrets())
            dr.deactivate_doc_user(ADMIN, "second-admin@example.test", db_path=path, secrets_getter=admin_secrets())
            self.assertFalse(dr.is_admin("second-admin@example.test", db_path=path))
        finally:
            os.remove(path)

    def test_never_writes_an_audit_row(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            dr.create_doc_user(ADMIN, "second-admin@example.test", dr.ROLE_ADMIN, db_path=path, secrets_getter=admin_secrets())
            before = len(audit_rows(path))
            dr.is_admin("second-admin@example.test", db_path=path)
            dr.is_admin(OUTSIDER, db_path=path)
            self.assertEqual(len(audit_rows(path)), before)
        finally:
            os.remove(path)

    def test_malformed_db_path_never_raises(self):
        # A directory can never be opened as a SQLite file.
        self.assertFalse(dr.is_admin(OUTSIDER, db_path=tempfile.mkdtemp()))


# ---------------------------------------------------------------------------
# Phase 3 read-only helpers: get_access_level, categories_with_min_level
# ---------------------------------------------------------------------------
class AccessHelperTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.cat_a = dr.create_category(ADMIN, "Category A", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.cat_b = dr.create_category(ADMIN, "Category B", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.cat_a, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.cat_b, CONTRIBUTOR, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_get_access_level_admin(self):
        self.assertEqual(dr.get_access_level(ADMIN, self.cat_a, db_path=self.db_path, secrets_getter=self.secrets_getter), "admin")

    def test_get_access_level_contributor(self):
        self.assertEqual(dr.get_access_level(CONTRIBUTOR, self.cat_a, db_path=self.db_path, secrets_getter=self.secrets_getter), dr.LEVEL_CONTRIBUTE)

    def test_get_access_level_none_for_no_access(self):
        self.assertIsNone(dr.get_access_level(OUTSIDER, self.cat_a, db_path=self.db_path, secrets_getter=self.secrets_getter))

    def test_get_access_level_empty_actor_returns_none(self):
        self.assertIsNone(dr.get_access_level("", self.cat_a, db_path=self.db_path, secrets_getter=self.secrets_getter))

    def test_get_access_level_writes_no_audit_row(self):
        before = len(audit_rows(self.db_path))
        dr.get_access_level(CONTRIBUTOR, self.cat_a, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.get_access_level(OUTSIDER, self.cat_a, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(audit_rows(self.db_path)), before)

    def test_categories_with_min_level_admin_sees_all(self):
        cats = dr.categories_with_min_level(ADMIN, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual({c["id"] for c in cats}, {self.cat_a, self.cat_b})

    def test_categories_with_min_level_filters_by_level(self):
        cats = dr.categories_with_min_level(CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual({c["id"] for c in cats}, {self.cat_a})

    def test_categories_with_min_level_empty_for_no_qualifying_category(self):
        cats = dr.categories_with_min_level(CONTRIBUTOR, dr.LEVEL_MANAGE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(cats, [])

    def test_categories_with_min_level_empty_for_unknown_user(self):
        cats = dr.categories_with_min_level(OUTSIDER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(cats, [])


# ---------------------------------------------------------------------------
# Phase 3 read-only helper: find_duplicate_by_sha256
# ---------------------------------------------------------------------------
class DuplicateDetectionTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Dupes", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.data = make_pdf()
        self.sha256 = hashlib.sha256(self.data).hexdigest()
        self.doc_id = dr.add_document(
            ADMIN, self.category_id, "Original", self.data, "a.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )

    def test_finds_matching_active_document(self):
        dup = dr.find_duplicate_by_sha256(ADMIN, self.category_id, self.sha256, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNotNone(dup)
        self.assertEqual(dup["id"], self.doc_id)

    def test_no_match_for_different_hash(self):
        other_hash = hashlib.sha256(b"different content").hexdigest()
        dup = dr.find_duplicate_by_sha256(ADMIN, self.category_id, other_hash, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNone(dup)

    def test_no_match_once_archived(self):
        dr.archive_document(ADMIN, self.doc_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dup = dr.find_duplicate_by_sha256(ADMIN, self.category_id, self.sha256, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNone(dup)

    def test_viewer_with_access_can_check(self):
        dup = dr.find_duplicate_by_sha256(VIEWER, self.category_id, self.sha256, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertIsNotNone(dup)

    def test_outsider_denied(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.find_duplicate_by_sha256(OUTSIDER, self.category_id, self.sha256, db_path=self.db_path, secrets_getter=self.secrets_getter)


# ---------------------------------------------------------------------------
# Phase 4 read-only helper: is_admin_via_secret
# ---------------------------------------------------------------------------
class IsAdminViaSecretTests(unittest.TestCase):
    def test_true_for_secret_listed_email(self):
        self.assertTrue(dr.is_admin_via_secret(ADMIN, secrets_getter=admin_secrets()))

    def test_false_for_non_secret_email(self):
        self.assertFalse(dr.is_admin_via_secret(OUTSIDER, secrets_getter=admin_secrets()))

    def test_false_for_doc_users_admin_not_in_secret(self):
        # is_admin() would say True for a doc_users admin row even with
        # no matching secret -- is_admin_via_secret() must say False,
        # since it only ever looks at the secret.
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            dr.create_doc_user(ADMIN, "second-admin@example.test", dr.ROLE_ADMIN, db_path=path, secrets_getter=admin_secrets())
            self.assertTrue(dr.is_admin("second-admin@example.test", db_path=path, secrets_getter=admin_secrets()))
            self.assertFalse(dr.is_admin_via_secret("second-admin@example.test", secrets_getter=admin_secrets()))
        finally:
            os.remove(path)

    def test_empty_actor_returns_false(self):
        self.assertFalse(dr.is_admin_via_secret(""))

    def test_never_raises_on_malformed_secret(self):
        self.assertFalse(dr.is_admin_via_secret(ADMIN, secrets_getter=lambda: {"DOC_REPO_ADMIN_EMAILS": None}))


# ---------------------------------------------------------------------------
# Phase 4 read-only helper: list_category_access
# ---------------------------------------------------------------------------
class ListCategoryAccessTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Access List", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, VIEWER, dr.LEVEL_VIEW, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_admin_sees_all_grants(self):
        rows = dr.list_category_access(ADMIN, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        subjects = {r["subject"]: r["level"] for r in rows}
        self.assertEqual(subjects, {CONTRIBUTOR: dr.LEVEL_CONTRIBUTE, VIEWER: dr.LEVEL_VIEW})

    def test_non_admin_denied(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.list_category_access(CONTRIBUTOR, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_denial_is_audited(self):
        before = len(audit_rows(self.db_path, "list_category_access_denied"))
        with self.assertRaises(dr.PermissionDenied):
            dr.list_category_access(VIEWER, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(audit_rows(self.db_path, "list_category_access_denied")), before + 1)

    def test_successful_read_not_audited(self):
        before = len(audit_rows(self.db_path))
        dr.list_category_access(ADMIN, self.category_id, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(audit_rows(self.db_path)), before)


# ---------------------------------------------------------------------------
# Phase 4 read-only helper: list_audit_log
# ---------------------------------------------------------------------------
class ListAuditLogTests(TempDbTestCase):
    def setUp(self):
        super().setUp()
        self.secrets_getter = admin_secrets()
        self.category_id = dr.create_category(ADMIN, "Audit List", db_path=self.db_path, secrets_getter=self.secrets_getter)
        dr.grant_category_access(ADMIN, self.category_id, CONTRIBUTOR, dr.LEVEL_CONTRIBUTE, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_non_admin_denied(self):
        with self.assertRaises(dr.PermissionDenied):
            dr.list_audit_log(CONTRIBUTOR, db_path=self.db_path, secrets_getter=self.secrets_getter)

    def test_newest_first(self):
        dr.create_category(ADMIN, "Second Category", db_path=self.db_path, secrets_getter=self.secrets_getter)
        rows = dr.list_audit_log(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertGreaterEqual(len(rows), 2)
        self.assertTrue(rows[0]["id"] > rows[1]["id"])

    def test_filter_by_action_substring(self):
        rows = dr.list_audit_log(ADMIN, action_filter="create_category", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertTrue(rows)
        self.assertTrue(all("create_category" in r["action"] for r in rows))

    def test_filter_by_actor(self):
        rows = dr.list_audit_log(ADMIN, actor_filter=CONTRIBUTOR, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(rows, [])  # CONTRIBUTOR never performed a mutation yet
        rows = dr.list_audit_log(ADMIN, actor_filter=ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertTrue(rows)

    def test_action_filter_escapes_like_wildcards(self):
        rows = dr.list_audit_log(ADMIN, action_filter="create%category", db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(rows, [])  # literal "%" must not match "_" as a wildcard

    def test_pagination(self):
        for i in range(5):
            dr.create_category(ADMIN, f"Pagination Cat {i}", db_path=self.db_path, secrets_getter=self.secrets_getter)
        page1 = dr.list_audit_log(ADMIN, page=1, page_size=2, db_path=self.db_path, secrets_getter=self.secrets_getter)
        page2 = dr.list_audit_log(ADMIN, page=2, page_size=2, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(page1), 2)
        self.assertEqual(len(page2), 2)
        self.assertNotEqual({r["id"] for r in page1}, {r["id"] for r in page2})

    def test_reading_log_is_not_itself_audited(self):
        before = len(audit_rows(self.db_path))
        dr.list_audit_log(ADMIN, db_path=self.db_path, secrets_getter=self.secrets_getter)
        self.assertEqual(len(audit_rows(self.db_path)), before)

    # -----------------------------------------------------------------
    # date_from/date_to: whole-calendar-day bounds, not a literal
    # string comparison against ts (a full UTC timestamp). ts has no
    # test-time injection point (unlike app.py's _test_clock for the
    # tile-visibility cache) -- these insert synthetic doc_audit_log
    # rows with explicit timestamps directly, same way other tests in
    # this file reach into the database with a raw sqlite3 connection
    # for setup/verification.
    # -----------------------------------------------------------------
    def _insert_audit_row(self, ts: str, action: str = "synthetic_test_action") -> None:
        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "INSERT INTO doc_audit_log (ts, actor_email, action) VALUES (?, ?, ?)",
            (ts, ADMIN, action),
        )
        conn.commit()
        conn.close()

    def test_date_from_includes_a_row_at_the_very_start_of_that_day(self):
        self._insert_audit_row("2026-11-13T00:00:00+00:00")
        rows = dr.list_audit_log(
            ADMIN, action_filter="synthetic_test_action", date_from="2026-11-13",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertEqual(len(rows), 1)

    def test_date_to_includes_a_row_at_any_time_on_that_day(self):
        # This is the exact case the bug affected: a plain "ts <=
        # date_to" excluded every row on date_to's own day, since any
        # real timestamp like "2026-11-13T23:59:59+00:00" sorts AFTER
        # the bare string "2026-11-13".
        self._insert_audit_row("2026-11-13T23:59:59+00:00")
        rows = dr.list_audit_log(
            ADMIN, action_filter="synthetic_test_action", date_to="2026-11-13",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertEqual(len(rows), 1)

    def test_date_to_includes_a_row_at_the_very_start_of_that_day_too(self):
        self._insert_audit_row("2026-11-13T00:00:00+00:00")
        rows = dr.list_audit_log(
            ADMIN, action_filter="synthetic_test_action", date_to="2026-11-13",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertEqual(len(rows), 1)

    def test_day_before_date_from_is_excluded(self):
        self._insert_audit_row("2026-11-12T23:59:59+00:00")
        rows = dr.list_audit_log(
            ADMIN, action_filter="synthetic_test_action", date_from="2026-11-13",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertEqual(rows, [])

    def test_day_after_date_to_is_excluded(self):
        self._insert_audit_row("2026-11-14T00:00:00+00:00")
        rows = dr.list_audit_log(
            ADMIN, action_filter="synthetic_test_action", date_to="2026-11-13",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertEqual(rows, [])

    def test_combined_from_and_to_same_day_includes_whole_day_only(self):
        self._insert_audit_row("2026-11-12T23:59:59+00:00")  # before
        self._insert_audit_row("2026-11-13T00:00:00+00:00")  # start of day -- in range
        self._insert_audit_row("2026-11-13T23:59:59+00:00")  # end of day -- in range
        self._insert_audit_row("2026-11-14T00:00:00+00:00")  # after
        rows = dr.list_audit_log(
            ADMIN, action_filter="synthetic_test_action", date_from="2026-11-13", date_to="2026-11-13",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        self.assertEqual(len(rows), 2)
        self.assertTrue(all(r["ts"].startswith("2026-11-13") for r in rows))


# ---------------------------------------------------------------------------
# seed_default_categories
# ---------------------------------------------------------------------------
class SeedDefaultCategoriesTests(TempDbTestCase):
    def test_seeds_nine_categories_and_is_idempotent(self):
        dr.seed_default_categories(db_path=self.db_path)
        dr.seed_default_categories(db_path=self.db_path)
        cats = dr.list_categories(ADMIN, db_path=self.db_path, secrets_getter=admin_secrets())
        self.assertEqual(len(cats), 9)
        self.assertEqual({c["name"] for c in cats}, set(dr.DEFAULT_CATEGORIES))

    def test_does_not_overwrite_existing_categories(self):
        dr.create_category(ADMIN, "Pre-existing Category", db_path=self.db_path, secrets_getter=admin_secrets())
        dr.seed_default_categories(db_path=self.db_path)
        cats = dr.list_categories(ADMIN, db_path=self.db_path, secrets_getter=admin_secrets())
        self.assertEqual(len(cats), 1)
        self.assertEqual(cats[0]["name"], "Pre-existing Category")


if __name__ == "__main__":
    unittest.main()
