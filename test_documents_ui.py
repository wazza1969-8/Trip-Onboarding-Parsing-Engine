"""
Automated tests for the Document Repository tile/screen -- Phase 2
(browse/search/download) and Phase 3 (upload, replace, edit metadata,
archive/restore, "show archived") -- both the app.py hook points and
documents_ui.py's rendering, driven through Streamlit's AppTest
harness, same pattern as test_weekly_report.py's UITests.

Phase 2 coverage: the tile hidden by default (including the "in
password mode" case, as far as AppTest can actually exercise it -- see
test_tile_hidden_when_not_authenticated_via_real_okta()'s docstring for
why that is the honest limit of what this harness can simulate here);
tile shown for an admin; the landing screen's layout unchanged when the
tile is hidden; the router resetting a stale "Document Repository" task
if permission is lost mid-session; empty states; a viewer seeing only
the categories they were granted; the document table never reading
document_blobs; a broken documents_repo import leaving the rest of the
app untouched; and the tile-visibility TTL cache in
_can_see_document_repository() (reused within 300s, recomputed after
it expires or the signed-in email changes, errors cached as False).

Phase 3 coverage: process_uploads() unit-tested directly (no widget
driving needed) for a mixed batch of good/bad files, the duplicate-
content warning, and refusal for a contributor without access or a
viewer; AppTest DOES turn out to support driving st.file_uploader (via
.set_value()), confirmed empirically before writing these -- so the
Upload tab itself, Replace file, and the archived-listing/toggle
behavior are also covered end-to-end through AppTest, not just at the
process_uploads() level.

Phase 3 fix-up coverage: _escape_markdown() neutralizes image/link/
bold/etc Markdown syntax (verified both as a pure-function unit test
and, separately, as a rendered-text assertion through the detail
panel) -- confirmed in a REAL rendered browser (not just AppTest, which
only exposes the Markdown source string, never rendered HTML) that
every one of st.markdown/st.write/st.caption/st.success/st.error/
st.warning, widget labels, and st.progress's text= all render this
Markdown subset, so a title containing
"![x](http://example.test/a.png) [click](http://example.test)" would
otherwise render as an actual image/link; two uploaded files sharing
the same name and size no longer collide on the same widget key (now
keyed by list position); a successful Replace resets that document's
file_uploader; and the duplicate-content check is cached per selected
file and only recomputed when the category changes.

Deliberately out of scope for automated proof here (reported in the
delivery notes instead, not silently skipped): the real browser
finding that even escaped text can still have a *bare* URL substring
autolinked by the underlying renderer in some of these same call sites
(not st.button/st.download_button) -- escaping defeats the explicit
image/link syntax and corrupts any such residual autolink's href (it
never resolves to the real domain), but AppTest cannot render real
HTML, so there is no way to assert "zero <a> elements" from here; the
browser-based verification already done stands in for that.

ISOLATION: app.py (and documents_ui.py) never pass an explicit db_path
to any documents_repo call -- same late-binding situation
test_weekly_report.py's own UITests docstring describes for backend.py.
So this file does the same thing that file already does: it clears
(not swaps) the real local documents.db (documents_repo.DEFAULT_DB_PATH)
before and after every test, and never touches the production Turso
database (no TURSO_DATABASE_URL secret is ever set here). Only fake
"@example.test" emails and synthetic file content are used.

Run with:
    python3 -m pytest test_documents_ui.py -v
(PYTEST_CURRENT_TEST must be set -- i.e. run under pytest, not plain
unittest -- for _RUNNING_UNDER_TEST in app.py to correctly skip the
real Okta/password gate. See app.py's own docstring for that flag.)
"""
import io
import sqlite3
import struct
import sys
import unittest
import zipfile
import zlib
from datetime import date
from unittest import mock

import documents_repo
import documents_ui

try:
    from streamlit.testing.v1 import AppTest
    _APPTEST_AVAILABLE = True
except ImportError:
    _APPTEST_AVAILABLE = False

ADMIN_EMAIL = "admin@example.test"
CONTRIBUTOR_EMAIL = "contributor@example.test"
VIEWER_EMAIL = "viewer@example.test"
OUTSIDER_EMAIL = "outsider@example.test"

PDF_BYTES = b"%PDF-1.4\n%fake pdf content for tests\n"


def make_pdf(total_size=None):
    header = b"%PDF-1.4\n%fake pdf content for tests\n"
    if total_size is None:
        return header
    return header + b"0" * (total_size - len(header))


def make_png():
    """A real, decodable 2x2 PNG -- not just magic bytes -- so tests
    that click Preview actually exercise st.image() decoding the
    bytes, not just that a button existed. Pillow is a hard streamlit
    dependency (not something this adds), so it's always available
    wherever these tests run."""
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (2, 2), color="red").save(buf, format="PNG")
    return buf.getvalue()


def make_oversized_png():
    # Magic bytes only (no need to be a real decodable image): this is
    # for the "too big to preview" path, which checks size and never
    # calls st.image() at all.
    return b"\x89PNG\r\n\x1a\n" + b"0" * (documents_ui.PREVIEW_MAX_BYTES + 1024)


def make_png_with_huge_claimed_dimensions(width=40000, height=40000):
    """A tiny, well-formed PNG (just IHDR + IEND, no pixel data) whose
    header claims enormous dimensions -- a classic "decompression
    bomb" shape. Small enough to sail under PREVIEW_MAX_BYTES, so this
    exercises the pixel-count check, not the byte-size check."""
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ihdr_chunk = struct.pack(">I", len(ihdr_data)) + b"IHDR" + ihdr_data + struct.pack(">I", zlib.crc32(b"IHDR" + ihdr_data))
    iend_chunk = struct.pack(">I", 0) + b"IEND" + struct.pack(">I", zlib.crc32(b"IEND"))
    return sig + ihdr_chunk + iend_chunk


def make_ooxml_zip(macro=False):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", "<document/>")
        if macro:
            zf.writestr("word/vbaProject.bin", b"fake macro project bytes")
    return buf.getvalue()


def admin_secrets_getter():
    return {"DOC_REPO_ADMIN_EMAILS": ADMIN_EMAIL}


_DOC_REPO_TABLES = (
    "doc_audit_log", "document_blobs", "document_versions", "documents",
    "doc_category_access", "doc_categories", "doc_users",
)


def _clear_document_repo_tables():
    conn = documents_repo.get_conn(documents_repo.DEFAULT_DB_PATH)
    for table in _DOC_REPO_TABLES:
        conn.execute(f"DELETE FROM {table}")
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# process_uploads() -- a plain function, no Streamlit calls, so these
# never touch AppTest at all; they call it directly like any other
# unit test would.
# ---------------------------------------------------------------------------
class ProcessUploadsTests(unittest.TestCase):
    def setUp(self):
        import os
        import tempfile
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.db_path = path
        self.secrets_getter = admin_secrets_getter
        self.category_id = documents_repo.create_category(
            ADMIN_EMAIL, "Uploads", db_path=self.db_path, secrets_getter=self.secrets_getter,
        )

    def tearDown(self):
        import os
        os.remove(self.db_path)

    def _upload(self, actor_email, uploads, **kwargs):
        return documents_ui.process_uploads(
            actor_email, self.category_id, uploads, db_path=self.db_path,
            secrets_getter=self.secrets_getter, **kwargs,
        )

    def test_mixed_batch_good_files_succeed_bad_files_report_specific_errors(self):
        uploads = [
            {"filename": "good.pdf", "data": make_pdf(), "title": "Good File"},
            {"filename": "empty.pdf", "data": b"", "title": "Empty File"},
            {"filename": "huge.pdf", "data": make_pdf(documents_repo.MAX_FILE_SIZE_BYTES + 1), "title": "Huge File"},
            {"filename": "spoofed.pdf", "data": b"not actually a pdf", "title": "Spoofed File"},
            {"filename": "macro.docx", "data": make_ooxml_zip(macro=True), "title": "Macro File"},
            {"filename": "bad.exe", "data": b"whatever bytes", "title": "Disallowed File"},
        ]
        results = self._upload(ADMIN_EMAIL, uploads)
        by_filename = {r["filename"]: r for r in results}

        self.assertTrue(by_filename["good.pdf"]["success"])
        self.assertIsNotNone(by_filename["good.pdf"]["document_id"])

        self.assertFalse(by_filename["empty.pdf"]["success"])
        self.assertIn("empty", by_filename["empty.pdf"]["error"].lower())

        self.assertFalse(by_filename["huge.pdf"]["success"])
        self.assertIn("large", by_filename["huge.pdf"]["error"].lower())

        self.assertFalse(by_filename["spoofed.pdf"]["success"])
        self.assertIn("doesn't look like", by_filename["spoofed.pdf"]["error"])

        self.assertFalse(by_filename["macro.docx"]["success"])
        self.assertIn("Macro", by_filename["macro.docx"]["error"])

        self.assertFalse(by_filename["bad.exe"]["success"])
        self.assertIn("aren't allowed", by_filename["bad.exe"]["error"])

        # One bad file must never stop the others -- all six were attempted.
        self.assertEqual(len(results), 6)

    def test_default_title_from_filename_when_blank(self):
        results = self._upload(ADMIN_EMAIL, [{"filename": "quarterly report.pdf", "data": make_pdf(), "title": ""}])
        self.assertEqual(results[0]["title"], "quarterly report")

    def test_duplicate_warning_is_non_blocking(self):
        data = make_pdf()
        documents_repo.add_document(
            ADMIN_EMAIL, self.category_id, "Original", data, "original.pdf",
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        results = self._upload(ADMIN_EMAIL, [{"filename": "copy.pdf", "data": data, "title": "Copy"}])
        self.assertTrue(results[0]["success"])
        self.assertIsNotNone(results[0]["warning"])
        self.assertIn("Original", results[0]["warning"])

    def test_no_warning_for_unique_content(self):
        results = self._upload(ADMIN_EMAIL, [{"filename": "unique.pdf", "data": make_pdf(), "title": "Unique"}])
        self.assertTrue(results[0]["success"])
        self.assertIsNone(results[0]["warning"])

    def test_contributor_without_access_refused(self):
        results = self._upload(CONTRIBUTOR_EMAIL, [{"filename": "a.pdf", "data": make_pdf(), "title": "A"}])
        self.assertFalse(results[0]["success"])
        self.assertIn("not allowed", results[0]["error"])

    def test_viewer_cannot_upload(self):
        documents_repo.grant_category_access(
            ADMIN_EMAIL, self.category_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW,
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        results = self._upload(VIEWER_EMAIL, [{"filename": "a.pdf", "data": make_pdf(), "title": "A"}])
        self.assertFalse(results[0]["success"])
        self.assertIn("not allowed", results[0]["error"])

    def test_contributor_with_access_succeeds(self):
        documents_repo.grant_category_access(
            ADMIN_EMAIL, self.category_id, CONTRIBUTOR_EMAIL, documents_repo.LEVEL_CONTRIBUTE,
            db_path=self.db_path, secrets_getter=self.secrets_getter,
        )
        results = self._upload(CONTRIBUTOR_EMAIL, [{"filename": "a.pdf", "data": make_pdf(), "title": "A"}])
        self.assertTrue(results[0]["success"])

    def test_progress_callback_called_once_per_file(self):
        calls = []
        uploads = [
            {"filename": "a.pdf", "data": make_pdf(), "title": "A"},
            {"filename": "b.pdf", "data": make_pdf(), "title": "B"},
        ]
        self._upload(ADMIN_EMAIL, uploads, progress_callback=lambda fraction, filename: calls.append((fraction, filename)))
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[-1][0], 1.0)


# ---------------------------------------------------------------------------
# _escape_markdown() -- a plain function, no Streamlit calls.
# ---------------------------------------------------------------------------
class MarkdownEscapingTests(unittest.TestCase):
    MALICIOUS_TEXT = "![x](http://example.test/a.png) [click](http://example.test)"

    def test_none_becomes_empty_string(self):
        self.assertEqual(documents_ui._escape_markdown(None), "")

    def test_image_and_link_syntax_characters_are_all_escaped(self):
        escaped = documents_ui._escape_markdown(self.MALICIOUS_TEXT)
        # The exact raw substrings that form the image/link constructs
        # must never survive unescaped.
        self.assertNotIn("![x](", escaped)
        self.assertNotIn("[click](", escaped)
        self.assertIn("\\!\\[x\\]\\(", escaped)
        self.assertIn("\\[click\\]\\(", escaped)

    def test_bold_marker_is_escaped(self):
        self.assertNotIn("**", documents_ui._escape_markdown("**bold**"))

    def test_backslash_is_escaped_before_other_characters(self):
        # A pre-existing backslash must itself end up escaped, not
        # interfere with (or be consumed by) escaping the character
        # that follows it -- hence escaping backslash first.
        backslash, asterisk = "\\", "*"
        escaped = documents_ui._escape_markdown(backslash + asterisk)
        self.assertEqual(escaped, backslash * 3 + asterisk)

    def test_every_listed_special_character_is_escaped(self):
        for ch in "`*_{}[]()#+-.!|>~<":
            escaped = documents_ui._escape_markdown(f"a{ch}b")
            self.assertEqual(escaped, f"a\\{ch}b", msg=f"character {ch!r} was not escaped")


@unittest.skipUnless(_APPTEST_AVAILABLE, "streamlit.testing.v1.AppTest not available")
class DocumentRepositoryUITests(unittest.TestCase):
    def setUp(self):
        _clear_document_repo_tables()

    def tearDown(self):
        _clear_document_repo_tables()

    @staticmethod
    def _at_as(actor_email, *, admin_secret=True, enabled=False):
        at = AppTest.from_file("app.py")
        if admin_secret:
            at.secrets["DOC_REPO_ADMIN_EMAILS"] = ADMIN_EMAIL
        if enabled:
            at.secrets["DOC_REPO_ENABLED"] = "true"
        if actor_email is not None:
            at.session_state["_test_actor_email"] = actor_email
        at.run(timeout=30)
        return at

    @staticmethod
    def _select_document_repository(at):
        at.button(key="select_Document Repository").click().run(timeout=30)
        return at

    @staticmethod
    def _select_admin_section(at, section_name):
        """Admin sub-sections are lazy-loaded (Phase 5 performance change):
        only the selected section's widgets exist on a given rerun, so
        tests for any section other than the default ("Categories") must
        switch to it first."""
        at.segmented_control(key="doc_repo_admin_section").set_value(section_name).run(timeout=30)
        return at

    # -----------------------------------------------------------------
    # Tile visibility
    # -----------------------------------------------------------------
    def test_tile_shown_for_an_admin(self):
        at = self._at_as(ADMIN_EMAIL)
        self.assertFalse(at.exception)
        select_keys = [b.key for b in at.button if b.key and b.key.startswith("select_")]
        self.assertIn("select_Document Repository", select_keys)

    def test_tile_hidden_for_a_non_admin(self):
        at = self._at_as(OUTSIDER_EMAIL)
        self.assertFalse(at.exception)
        select_keys = [b.key for b in at.button if b.key and b.key.startswith("select_")]
        self.assertNotIn("select_Document Repository", select_keys)

    def test_tile_hidden_when_not_authenticated_via_real_okta(self):
        """Stands in for "tile hidden in password mode".

        Under _RUNNING_UNDER_TEST, app.py's real Okta-vs-password-stopgap
        branch never executes at all -- that is the whole point of
        _RUNNING_UNDER_TEST (see its docstring in app.py): AppTest has no
        supported way to populate a real st.user OIDC identity, so the
        entire real auth block is skipped under test, password mode
        included. That means password mode specifically cannot be driven
        through AppTest here -- this is not a gap in these tests, it is
        the same pre-existing limitation that justified _RUNNING_UNDER_TEST
        in the first place.

        What IS genuinely testable, and tested here, is the thing that
        actually matters for this tile: with no test-provided actor email
        (the default), _document_repository_actor_email() returns None --
        exactly the value it would also return for a real signed-out
        visitor or someone stuck behind the password stopgap, since both
        of those also fail "real Okta sign-in confirmed". So this test
        confirms the tile is hidden in that same default state.
        """
        at = self._at_as(None)
        self.assertFalse(at.exception)
        select_keys = [b.key for b in at.button if b.key and b.key.startswith("select_")]
        self.assertNotIn("select_Document Repository", select_keys)

    def test_landing_screen_unchanged_when_tile_hidden(self):
        at = self._at_as(OUTSIDER_EMAIL)
        self.assertFalse(at.exception)
        select_keys = sorted(b.key for b in at.button if b.key and b.key.startswith("select_"))
        self.assertEqual(
            select_keys,
            sorted([
                "select_Trip Preference Parsing",
                "select_Flight Planning Parsing",
                "select_Weekly Report",
            ]),
        )

    def test_router_resets_if_permission_is_lost_mid_session(self):
        # The visibility check is now cached for 300s (see
        # _can_see_document_repository() in app.py) -- a revocation
        # inside that window is an accepted, intentional blind spot of
        # that cache, not something this test should paper over. So
        # the test's own fake clock (st.session_state["_test_clock"] --
        # see _cache_clock() in app.py) is advanced past the TTL before
        # the revoked state is expected to actually take effect; see
        # the dedicated cache tests below for that behavior in isolation.
        at = AppTest.from_file("app.py")
        at.secrets["DOC_REPO_ADMIN_EMAILS"] = ADMIN_EMAIL
        at.session_state["_test_actor_email"] = ADMIN_EMAIL
        at.session_state["_test_clock"] = 1000.0
        at.run(timeout=30)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual(at.session_state["task"], "Document Repository")

        # Access revoked mid-session: the admin secret no longer lists
        # this email -- and the fake clock has moved past the TTL.
        at.secrets["DOC_REPO_ADMIN_EMAILS"] = "someone-else@example.test"
        at.session_state["_test_clock"] = 1301.0
        at.run(timeout=30)
        self.assertFalse(at.exception)
        self.assertIsNone(at.session_state["task"])
        self.assertEqual([s.value for s in at.subheader], ["What would you like to do?"])

    # -----------------------------------------------------------------
    # Tile-visibility cache (_can_see_document_repository() in app.py)
    # -----------------------------------------------------------------
    def test_second_call_within_ttl_does_not_recall_is_admin(self):
        with mock.patch.object(documents_repo, "is_admin", wraps=documents_repo.is_admin) as mocked:
            at = AppTest.from_file("app.py")
            at.secrets["DOC_REPO_ADMIN_EMAILS"] = ADMIN_EMAIL
            at.session_state["_test_actor_email"] = ADMIN_EMAIL
            at.session_state["_test_clock"] = 1000.0
            at.run(timeout=30)
            self.assertEqual(mocked.call_count, 1)

            at.session_state["_test_clock"] = 1001.0  # 1s later -- still within the 300s TTL
            at.run(timeout=30)  # a second rerun, same email
            self.assertEqual(mocked.call_count, 1)
        self.assertFalse(at.exception)

    def test_cache_expires_after_ttl_and_recomputes(self):
        with mock.patch.object(documents_repo, "is_admin", wraps=documents_repo.is_admin) as mocked:
            at = AppTest.from_file("app.py")
            at.secrets["DOC_REPO_ADMIN_EMAILS"] = ADMIN_EMAIL
            at.session_state["_test_actor_email"] = ADMIN_EMAIL
            at.session_state["_test_clock"] = 1000.0
            at.run(timeout=30)
            self.assertEqual(mocked.call_count, 1)

            at.session_state["_test_clock"] = 1301.0  # 301s later -- past the 300s TTL
            at.run(timeout=30)
            self.assertEqual(mocked.call_count, 2)
        self.assertFalse(at.exception)

    def test_different_email_recomputes(self):
        second_admin = "second-admin@example.test"
        with mock.patch.object(documents_repo, "is_admin", wraps=documents_repo.is_admin) as mocked:
            at = AppTest.from_file("app.py")
            at.secrets["DOC_REPO_ADMIN_EMAILS"] = f"{ADMIN_EMAIL},{second_admin}"
            at.session_state["_test_actor_email"] = ADMIN_EMAIL
            at.session_state["_test_clock"] = 1000.0
            at.run(timeout=30)
            self.assertEqual(mocked.call_count, 1)

            at.session_state["_test_clock"] = 1000.0  # no time elapsed at all
            at.session_state["_test_actor_email"] = second_admin
            at.run(timeout=30)  # different email -- recomputed anyway
            self.assertEqual(mocked.call_count, 2)
        self.assertFalse(at.exception)

    def test_error_is_treated_as_false(self):
        with mock.patch.object(documents_repo, "is_admin", side_effect=RuntimeError("boom")):
            at = self._at_as(ADMIN_EMAIL)
        self.assertFalse(at.exception)
        select_keys = [b.key for b in at.button if b.key and b.key.startswith("select_")]
        self.assertNotIn("select_Document Repository", select_keys)

    def test_documents_repo_import_failure_leaves_app_working(self):
        with mock.patch.dict(sys.modules, {"documents_repo": None, "documents_ui": None}):
            at = AppTest.from_file("app.py")
            at.run(timeout=30)
        self.assertFalse(at.exception)
        select_keys = sorted(b.key for b in at.button if b.key and b.key.startswith("select_"))
        self.assertEqual(
            select_keys,
            sorted([
                "select_Trip Preference Parsing",
                "select_Flight Planning Parsing",
                "select_Weekly Report",
            ]),
        )

    # -----------------------------------------------------------------
    # Empty states
    # -----------------------------------------------------------------
    def test_no_documents_yet_empty_state_after_seeding(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertIn("No documents yet.", [i.value for i in at.info])

    def test_no_categories_available_empty_state_for_a_granted_but_unassigned_viewer(self):
        # The only way a non-admin legitimately reaches this screen at
        # all: path (b), DOC_REPO_ENABLED + an active doc_users row --
        # with zero category access grants, so list_categories()
        # legitimately comes back empty for them.
        documents_repo.create_doc_user(
            ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER,
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(VIEWER_EMAIL, enabled=True)
        self.assertFalse(at.exception)
        select_keys = [b.key for b in at.button if b.key and b.key.startswith("select_")]
        self.assertIn("select_Document Repository", select_keys)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertIn(
            "No categories are available to you yet; ask an administrator for access.",
            [i.value for i in at.info],
        )

    def test_no_documents_in_this_category_empty_state(self):
        documents_repo.create_category(ADMIN_EMAIL, "Empty Category", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.radio(key="doc_repo_selected_category").set_value(
            [o for o in at.radio(key="doc_repo_selected_category").options if o.startswith("Empty Category")][0]
        ).run(timeout=30)
        self.assertFalse(at.exception)
        self.assertIn("No documents in this category yet.", [i.value for i in at.info])

    def test_no_documents_match_your_search_empty_state(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.text_input(key="doc_repo_search").set_value("totally-unmatched-xyz").run(timeout=30)
        self.assertFalse(at.exception)
        self.assertIn("No documents match your search.", [i.value for i in at.info])

    # -----------------------------------------------------------------
    # Access scoping and data-handling guarantees
    # -----------------------------------------------------------------
    def test_viewer_sees_only_granted_categories(self):
        visible_id = documents_repo.create_category(ADMIN_EMAIL, "Visible Category", secrets_getter=admin_secrets_getter)
        documents_repo.create_category(ADMIN_EMAIL, "Hidden Category", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(
            ADMIN_EMAIL, visible_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW,
            secrets_getter=admin_secrets_getter,
        )
        documents_repo.create_doc_user(
            ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER,
            secrets_getter=admin_secrets_getter,
        )

        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        options = " | ".join(at.radio(key="doc_repo_selected_category").options)
        self.assertIn("Visible Category", options)
        self.assertNotIn("Hidden Category", options)

    def test_document_table_never_reads_blob_bytes(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Blob Isolation", secrets_getter=admin_secrets_getter)
        documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Blob Isolation Doc", PDF_BYTES, "a.pdf",
            secrets_getter=admin_secrets_getter,
        )
        conn = sqlite3.connect(documents_repo.DEFAULT_DB_PATH)
        conn.execute("DELETE FROM document_blobs")
        conn.commit()
        conn.close()

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        # Looks for the Browse table's content among whichever
        # dataframes exist, rather than assuming there is exactly one
        # on the whole page -- the Admin tab's default section
        # ("Categories") has none, but this stays robust even if that
        # default or a section's own rendering ever changes.
        all_titles = [t for df in at.dataframe for t in df.value.get("Title", [])]
        self.assertIn("Blob Isolation Doc", all_titles)

    def test_browse_pagination_shows_count_of_total(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Count Caption", secrets_getter=admin_secrets_getter)
        for i in range(3):
            documents_repo.add_document(
                ADMIN_EMAIL, cat_id, f"Doc {i}", PDF_BYTES, f"doc_{i}.pdf",
                secrets_getter=admin_secrets_getter,
            )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        captions = " ".join(c.value for c in at.caption)
        self.assertIn("Showing 1-3 of 3", captions)

    def test_browse_date_to_today_still_shows_a_document_uploaded_today(self):
        # Regression test for the date_to off-by-one on created_at
        # (same bug class as the Admin audit log's date_to): "Uploaded
        # to" defaulting to today must not exclude a document created
        # moments ago today.
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Date Filter UI", secrets_getter=admin_secrets_getter)
        documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Uploaded Today", PDF_BYTES, "today.pdf",
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.date_input(key="doc_repo_date_to").set_value(date.today()).run(timeout=30)
        self.assertFalse(at.exception)
        all_titles = [t for df in at.dataframe for t in df.value.get("Title", [])]
        self.assertIn("Uploaded Today", all_titles)

    def test_prepare_download_then_download_button_appears(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Download Test", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Downloadable Doc", PDF_BYTES, "a.pdf",
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.download_button), 0)  # not loaded before the button is pressed

        at.button(key=f"doc_repo_prepare_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.download_button), 1)

    # -----------------------------------------------------------------
    # Phase 5: on-demand Preview (goes through download_version(), so
    # authorization/archived-status/audit all apply exactly like a
    # real download)
    # -----------------------------------------------------------------
    def test_preview_button_renders_image_for_png(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Preview PNG", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Preview Doc", make_png(), "a.png",
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.image), 0)  # not loaded before Preview is pressed

        at.button(key=f"doc_repo_preview_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.image), 1)

        # Preview goes through download_version() exactly like a real
        # download -- so it's audited as one, not silently exempt.
        rows = documents_repo.list_audit_log(ADMIN_EMAIL, action_filter="download_version", secrets_getter=admin_secrets_getter, page_size=10)
        self.assertEqual(len(rows), 1)

    def test_preview_denied_for_outsider_same_as_download(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Preview Denied", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Preview Doc", make_png(), "a.png",
            secrets_getter=admin_secrets_getter,
        )
        with self.assertRaises(documents_repo.PermissionDenied):
            documents_repo.download_version(OUTSIDER_EMAIL, doc_id, secrets_getter=admin_secrets_getter)

    def test_no_preview_button_for_non_previewable_format(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Preview PDF", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "PDF Doc", PDF_BYTES, "a.pdf",
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)
        preview_keys = [b.key for b in at.button if b.key and b.key.startswith("doc_repo_preview_btn::")]
        self.assertEqual(preview_keys, [])
        # PDFs are download-only this phase (st.pdf needs an extra
        # package not in requirements.txt) -- "Prepare download" must
        # still work normally.
        self.assertEqual(len([b for b in at.button if b.key == f"doc_repo_prepare_btn::{doc_id}"]), 1)

    def test_preview_shows_message_instead_of_button_when_over_size_limit(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Preview Oversized", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Huge Image", make_oversized_png(), "big.png",
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)
        preview_keys = [b.key for b in at.button if b.key and b.key.startswith("doc_repo_preview_btn::")]
        self.assertEqual(preview_keys, [])
        captions = " ".join(c.value for c in at.caption)
        self.assertIn("use the download option below", captions)

    def test_preview_rejects_image_with_huge_claimed_dimensions(self):
        # A tiny file (well under PREVIEW_MAX_BYTES) whose PNG header
        # claims 40000x40000 -- a decompression-bomb shape. Must not
        # raise, must not render st.image, must show a friendly message
        # instead of ever decoding the (claimed) 1.6 billion pixels.
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Preview Bomb", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Bomb Image", make_png_with_huge_claimed_dimensions(), "bomb.png",
            secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)

        at.button(key=f"doc_repo_preview_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.image), 0)
        infos = " ".join(i.value for i in at.info)
        self.assertIn("too large to preview safely", infos)

    def test_previewing_a_second_document_evicts_the_first_from_cache(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Preview Cache", secrets_getter=admin_secrets_getter)
        doc_a = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Doc A", make_png(), "a.png", secrets_getter=admin_secrets_getter,
        )
        doc_b = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, "Doc B", make_png(), "b.png", secrets_getter=admin_secrets_getter,
        )
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)

        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_a).run(timeout=30)
        at.button(key=f"doc_repo_preview_btn::{doc_a}").click().run(timeout=30)
        self.assertFalse(at.exception)
        preview_keys = [k for k in at.session_state if k.startswith("_doc_repo_preview::")]
        self.assertEqual(preview_keys, [f"_doc_repo_preview::{doc_a}"])

        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_b).run(timeout=30)
        at.button(key=f"doc_repo_preview_btn::{doc_b}").click().run(timeout=30)
        self.assertFalse(at.exception)
        preview_keys = [k for k in at.session_state if k.startswith("_doc_repo_preview::")]
        self.assertEqual(preview_keys, [f"_doc_repo_preview::{doc_b}"])

    # -----------------------------------------------------------------
    # Phase 3: Upload tab
    # -----------------------------------------------------------------
    def _uploaders(self, at):
        return [f for f in at.file_uploader if f.key and f.key.startswith("doc_repo_uploader_")]

    def test_upload_tab_hidden_for_viewer(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Viewer Only", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)

        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual([t.label for t in at.tabs], ["Browse"])

    def test_upload_tab_visible_for_admin(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        # An admin also gets the Admin tab (and its sub-tabs, which
        # at.tabs flattens in alongside the top-level ones) -- see the
        # dedicated Admin-tab tests for that; this only checks Upload.
        self.assertIn("Upload", [t.label for t in at.tabs])

    def test_upload_tab_visible_for_contributor(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Contributor Cat", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, CONTRIBUTOR_EMAIL, documents_repo.LEVEL_CONTRIBUTE, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, CONTRIBUTOR_EMAIL, documents_repo.ROLE_CONTRIBUTOR, secrets_getter=admin_secrets_getter)

        at = self._at_as(CONTRIBUTOR_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual([t.label for t in at.tabs], ["Browse", "Upload"])

    def test_upload_via_ui_succeeds(self):
        documents_repo.create_category(ADMIN_EMAIL, "Upload Target", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)

        self._uploaders(at)[0].set_value((("a.pdf", PDF_BYTES, "application/pdf"),)).run(timeout=30)
        self.assertFalse(at.exception)

        at.button(key="doc_repo_upload_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        success_texts = " ".join(s.value for s in at.success)
        self.assertIn(documents_ui._escape_markdown("a.pdf"), success_texts)
        self.assertIn("uploaded as", success_texts)

    def test_upload_reports_specific_error_for_bad_file(self):
        # st.file_uploader's own type= allowlist is enforced client-side
        # by the real widget -- confirmed empirically: a .exe can't even
        # be "selected" through set_value(), it raises before reaching
        # this module at all. So the realistic bad-file case reachable
        # through the actual UI is an ALLOWED extension whose content
        # doesn't match it (caught by documents_repo's signature check,
        # not by the widget).
        documents_repo.create_category(ADMIN_EMAIL, "Upload Errors", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)

        self._uploaders(at)[0].set_value((("spoofed.pdf", b"not actually a pdf", "application/pdf"),)).run(timeout=30)
        at.button(key="doc_repo_upload_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        error_texts = " ".join(e.value for e in at.error)
        self.assertIn(documents_ui._escape_markdown("spoofed.pdf"), error_texts)
        self.assertIn("doesn't look like", error_texts)

    def test_upload_tab_shows_duplicate_warning(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Dup Warn", secrets_getter=admin_secrets_getter)
        documents_repo.add_document(ADMIN_EMAIL, cat_id, "Existing", PDF_BYTES, "existing.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_upload_category").set_value(cat_id).run(timeout=30)
        self._uploaders(at)[0].set_value((("copy.pdf", PDF_BYTES, "application/pdf"),)).run(timeout=30)
        self.assertFalse(at.exception)
        warning_texts = " ".join(w.value for w in at.warning)
        self.assertIn("Existing", warning_texts)

    # -----------------------------------------------------------------
    # Phase 3: detail panel -- Replace file, Edit details
    # -----------------------------------------------------------------
    def test_replace_creates_new_version_and_keeps_old(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Replace Test", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Replaceable", PDF_BYTES, "v1.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)

        replace_uploaders = [f for f in at.file_uploader if f.key and f.key.startswith(f"doc_repo_replace_file::{doc_id}::")]
        self.assertEqual(len(replace_uploaders), 1)
        replace_uploaders[0].set_value(("v2.pdf", make_pdf(2000), "application/pdf")).run(timeout=30)
        at.button(key=f"doc_repo_replace_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)

        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertEqual(len(detail["versions"]), 2)
        self.assertEqual({v["filename"] for v in detail["versions"]}, {"v1.pdf", "v2.pdf"})

    def test_edit_details_updates_metadata(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Edit Test", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Old Title", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        at.text_input(key=f"doc_repo_edit_title::{doc_id}").set_value("New Title").run(timeout=30)
        at.button(key=f"doc_repo_save_details_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)

        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertEqual(detail["title"], "New Title")

    # -----------------------------------------------------------------
    # Phase 3: detail panel -- Archive / Restore with confirmation
    # -----------------------------------------------------------------
    def test_archive_requires_confirmation_then_restore(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Archive Test", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "To Archive", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)

        # First click only asks for confirmation -- the document must
        # still be active.
        at.button(key=f"doc_repo_archive_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertEqual(detail["status"], documents_repo.STATUS_ACTIVE)
        warning_texts = " ".join(w.value for w in at.warning)
        self.assertIn("Archive this document", warning_texts)

        # Confirming actually archives it.
        at.button(key=f"doc_repo_confirm_archive_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertEqual(detail["status"], documents_repo.STATUS_ARCHIVED)

        # Restore brings it back.
        at.toggle(key="doc_repo_show_archived").set_value(True).run(timeout=30)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        at.button(key=f"doc_repo_restore_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertEqual(detail["status"], documents_repo.STATUS_ACTIVE)

    def test_cancel_archive_leaves_document_active(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Cancel Archive", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Not Archived", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        at.button(key=f"doc_repo_archive_btn::{doc_id}").click().run(timeout=30)
        at.button(key=f"doc_repo_cancel_archive_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertEqual(detail["status"], documents_repo.STATUS_ACTIVE)

    # -----------------------------------------------------------------
    # Phase 3: "Show archived" toggle visibility and access scoping
    # -----------------------------------------------------------------
    def test_show_archived_toggle_hidden_for_viewer(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Toggle Viewer", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)

        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.toggle), 0)

    def test_show_archived_toggle_hidden_for_contributor_without_manage(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Toggle Contributor", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, CONTRIBUTOR_EMAIL, documents_repo.LEVEL_CONTRIBUTE, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, CONTRIBUTOR_EMAIL, documents_repo.ROLE_CONTRIBUTOR, secrets_getter=admin_secrets_getter)

        at = self._at_as(CONTRIBUTOR_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.toggle), 0)

    def test_show_archived_toggle_visible_for_manage_level(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Toggle Manage", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_MANAGE, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)

        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.toggle), 1)

    def test_archived_document_hidden_from_viewer_even_with_toggle_absent(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Archived Hidden", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Archived Doc", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)
        documents_repo.archive_document(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)

        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.toggle), 0)  # no manage-level category -- no toggle at all
        # The only document in the (one) category this viewer can see
        # is archived, and they have no toggle to reveal it -- so the
        # default "All categories" browse comes back with nothing.
        self.assertIn("No documents yet.", [i.value for i in at.info])

    # -----------------------------------------------------------------
    # Fix-up 1: Markdown escaping, end-to-end through the detail panel
    # -----------------------------------------------------------------
    def test_malicious_title_not_rendered_as_image_or_link_in_detail_panel(self):
        malicious_title = "![x](http://example.test/a.png) [click](http://example.test)"
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "XSS Test", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(
            ADMIN_EMAIL, cat_id, malicious_title, PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter,
        )

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)

        markdown_texts = " ".join(m.value for m in at.markdown)
        # The raw image/link syntax must never appear verbatim -- only
        # its escaped form, which (confirmed in a real rendered
        # browser -- see this file's module docstring) renders as
        # plain text, never as an actual image or link.
        self.assertNotIn("![x](http://example.test/a.png)", markdown_texts)
        self.assertNotIn("[click](http://example.test)", markdown_texts)
        self.assertIn(documents_ui._escape_markdown(malicious_title), markdown_texts)

    def test_malicious_category_name_escaped_in_detail_panel(self):
        malicious_name = "Cat **bold** [link](http://example.test)"
        cat_id = documents_repo.create_category(ADMIN_EMAIL, malicious_name, secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Doc", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        self.assertFalse(at.exception)

        # st.write(str) dispatches to st.markdown internally -- AppTest
        # exposes it there, there is no separate at.write.
        write_texts = " ".join(m.value for m in at.markdown)
        self.assertNotIn("[link](http://example.test)", write_texts)
        self.assertIn(documents_ui._escape_markdown(malicious_name), write_texts)

    def test_malicious_upload_result_text_escaped(self):
        documents_repo.create_category(ADMIN_EMAIL, "Upload XSS", secrets_getter=admin_secrets_getter)
        malicious_name = "evil**.pdf"

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self._uploaders(at)[0].set_value(((malicious_name, PDF_BYTES, "application/pdf"),)).run(timeout=30)
        at.button(key="doc_repo_upload_btn").click().run(timeout=30)
        self.assertFalse(at.exception)

        success_texts = " ".join(s.value for s in at.success)
        self.assertNotIn("evil**.pdf", success_texts)
        self.assertIn(documents_ui._escape_markdown(malicious_name), success_texts)

    # -----------------------------------------------------------------
    # Fix-up 2: file position (not name/size) in upload title keys
    # -----------------------------------------------------------------
    def test_duplicate_name_and_size_files_do_not_collide(self):
        documents_repo.create_category(ADMIN_EMAIL, "Dup Names", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)

        data = make_pdf(2000)
        self._uploaders(at)[0].set_value((
            ("same.pdf", data, "application/pdf"),
            ("same.pdf", data, "application/pdf"),
        )).run(timeout=30)
        self.assertFalse(at.exception)  # no duplicate-widget-key crash

        title_inputs = [t for t in at.text_input if t.key and "doc_repo_upload_title::" in t.key]
        self.assertEqual(len(title_inputs), 2)

        at.button(key="doc_repo_upload_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(at.success), 2)  # both files uploaded independently

    # -----------------------------------------------------------------
    # Fix-up 3: Replace resets its file_uploader after success
    # -----------------------------------------------------------------
    def test_replace_resets_uploader_after_success(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Replace Reset", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Doc", PDF_BYTES, "v1.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)

        def _replace_uploaders():
            return [f for f in at.file_uploader if f.key and f.key.startswith(f"doc_repo_replace_file::{doc_id}::")]

        _replace_uploaders()[0].set_value(("v2.pdf", make_pdf(2000), "application/pdf")).run(timeout=30)
        at.button(key=f"doc_repo_replace_btn::{doc_id}").click().run(timeout=30)
        self.assertFalse(at.exception)

        # Re-select the document: the Replace uploader must be a fresh,
        # empty widget -- not still showing v2.pdf as selected.
        at.selectbox(key="doc_repo_selected_doc_id").set_value(doc_id).run(timeout=30)
        after = _replace_uploaders()
        self.assertEqual(len(after), 1)
        self.assertIsNone(after[0].value)

    # -----------------------------------------------------------------
    # Fix-up 4: duplicate-content check is cached per selected file
    # -----------------------------------------------------------------
    def test_duplicate_check_cached_across_unrelated_reruns(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Cache Dup", secrets_getter=admin_secrets_getter)
        documents_repo.add_document(ADMIN_EMAIL, cat_id, "Existing", PDF_BYTES, "existing.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_upload_category").set_value(cat_id).run(timeout=30)

        with mock.patch.object(documents_repo, "find_duplicate_by_sha256", wraps=documents_repo.find_duplicate_by_sha256) as mocked:
            self._uploaders(at)[0].set_value((("copy.pdf", PDF_BYTES, "application/pdf"),)).run(timeout=30)
            self.assertEqual(mocked.call_count, 1)

            # An unrelated rerun (typing in the description box) must
            # not re-hash the file or re-query the database.
            at.text_area(key="doc_repo_upload_description").set_value("some notes").run(timeout=30)
            self.assertEqual(mocked.call_count, 1)

    def test_duplicate_check_recomputes_when_category_changes(self):
        # Note: Selectbox.options in AppTest returns the *formatted*
        # display labels (category names here), not the raw category
        # ID -- set_value() needs the raw ID, so both category IDs are
        # captured directly from create_category()'s return value
        # rather than derived from .options.
        cat_a = documents_repo.create_category(ADMIN_EMAIL, "Cache Dup A", secrets_getter=admin_secrets_getter)
        cat_b = documents_repo.create_category(ADMIN_EMAIL, "Cache Dup B", secrets_getter=admin_secrets_getter)
        documents_repo.add_document(ADMIN_EMAIL, cat_a, "Existing", PDF_BYTES, "existing.pdf", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.selectbox(key="doc_repo_upload_category").set_value(cat_a).run(timeout=30)

        with mock.patch.object(documents_repo, "find_duplicate_by_sha256", wraps=documents_repo.find_duplicate_by_sha256) as mocked:
            self._uploaders(at)[0].set_value((("copy.pdf", PDF_BYTES, "application/pdf"),)).run(timeout=30)
            self.assertFalse(at.exception)
            self.assertEqual(mocked.call_count, 1)
            warning_texts = " ".join(w.value for w in at.warning)
            self.assertIn("Existing", warning_texts)

            at.selectbox(key="doc_repo_upload_category").set_value(cat_a).run(timeout=30)
            self.assertEqual(mocked.call_count, 1)  # same category, same file -- still cached

            at.selectbox(key="doc_repo_upload_category").set_value(cat_b).run(timeout=30)
            self.assertFalse(at.exception)
            self.assertEqual(mocked.call_count, 2)  # category changed -- recomputed

    # -----------------------------------------------------------------
    # Admin tab: visibility
    # -----------------------------------------------------------------
    def test_admin_tab_hidden_for_viewer(self):
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertNotIn("Admin", [t.label for t in at.tabs])

    def test_admin_tab_hidden_for_contributor(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Contrib Only", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, CONTRIBUTOR_EMAIL, documents_repo.LEVEL_CONTRIBUTE, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, CONTRIBUTOR_EMAIL, documents_repo.ROLE_CONTRIBUTOR, secrets_getter=admin_secrets_getter)
        at = self._at_as(CONTRIBUTOR_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertNotIn("Admin", [t.label for t in at.tabs])

    def test_admin_tab_hidden_even_with_manage_level_access(self):
        # Manage-level access unlocks the "Show archived" toggle, but
        # it is not adminship -- the Admin tab must still be absent.
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Manage Only", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_MANAGE, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        at = self._at_as(VIEWER_EMAIL, enabled=True)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        self.assertNotIn("Admin", [t.label for t in at.tabs])

    def test_admin_tab_absent_for_non_admin_even_with_forced_session_state(self):
        # There is no session_state flag anywhere in this module that
        # gates the Admin tab's visibility -- it is decided fresh, every
        # run, from documents_repo.is_admin(). Forcing plausible-looking
        # state before the run must have no effect.
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        at = AppTest.from_file("app.py")
        at.secrets["DOC_REPO_ADMIN_EMAILS"] = ADMIN_EMAIL
        at.secrets["DOC_REPO_ENABLED"] = "true"
        at.session_state["_test_actor_email"] = VIEWER_EMAIL
        at.session_state["_doc_repo_admin_view"] = True
        at.session_state["_doc_repo_is_admin_override"] = True
        at.session_state["is_admin"] = True
        at.run(timeout=30)
        at.button(key="select_Document Repository").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertNotIn("Admin", [t.label for t in at.tabs])

    def test_admin_tab_and_subtabs_visible_for_admin(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        labels = [t.label for t in at.tabs]
        self.assertIn("Admin", labels)
        # Sub-sections are a lazy-loaded segmented_control (Phase 5), not
        # nested st.tabs() -- every option should still be offered, even
        # though only one section's content renders at a time.
        section_options = at.segmented_control(key="doc_repo_admin_section").options
        for sub in ("Categories", "Users & Roles", "Access Grants", "Archived Records", "Audit Log"):
            self.assertIn(sub, section_options)

    # -----------------------------------------------------------------
    # Admin tab: non-admin denied at the service level regardless of UI
    # -----------------------------------------------------------------
    def test_non_admin_cannot_call_admin_functions_directly(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Service Gate", secrets_getter=admin_secrets_getter)
        calls = [
            lambda: documents_repo.create_category(VIEWER_EMAIL, "x", secrets_getter=admin_secrets_getter),
            lambda: documents_repo.rename_category(VIEWER_EMAIL, cat_id, "x", secrets_getter=admin_secrets_getter),
            lambda: documents_repo.reorder_categories(VIEWER_EMAIL, [cat_id], secrets_getter=admin_secrets_getter),
            lambda: documents_repo.deactivate_category(VIEWER_EMAIL, cat_id, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.reactivate_category(VIEWER_EMAIL, cat_id, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.create_doc_user(VIEWER_EMAIL, "x@example.test", documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.update_doc_user_role(VIEWER_EMAIL, ADMIN_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.deactivate_doc_user(VIEWER_EMAIL, ADMIN_EMAIL, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.reactivate_doc_user(VIEWER_EMAIL, ADMIN_EMAIL, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.grant_category_access(VIEWER_EMAIL, cat_id, "x@example.test", documents_repo.LEVEL_VIEW, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.revoke_category_access(VIEWER_EMAIL, cat_id, "x@example.test", secrets_getter=admin_secrets_getter),
            lambda: documents_repo.list_doc_users(VIEWER_EMAIL, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.list_category_access(VIEWER_EMAIL, cat_id, secrets_getter=admin_secrets_getter),
            lambda: documents_repo.list_audit_log(VIEWER_EMAIL, secrets_getter=admin_secrets_getter),
        ]
        for call in calls:
            with self.assertRaises(documents_repo.PermissionDenied):
                call()

    # -----------------------------------------------------------------
    # Admin tab: Categories (create/rename/reorder/deactivate/reactivate)
    # -----------------------------------------------------------------
    def test_admin_create_category_via_ui(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.text_input(key="doc_repo_admin_new_category_name").set_value("Brand New Category").run(timeout=30)
        at.button(key="doc_repo_admin_create_category_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        names = {c["name"] for c in documents_repo.list_categories(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertIn("Brand New Category", names)

    def test_admin_create_category_duplicate_name_rejected(self):
        documents_repo.create_category(ADMIN_EMAIL, "Dup Cat", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.text_input(key="doc_repo_admin_new_category_name").set_value("Dup Cat").run(timeout=30)
        at.button(key="doc_repo_admin_create_category_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertTrue(any("already exists" in e.value for e in at.error))

    def test_admin_create_category_blank_name_rejected(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.button(key="doc_repo_admin_create_category_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertTrue(len(at.error) > 0)

    def test_admin_rename_category_via_ui(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Old Name", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.text_input(key=f"doc_repo_admin_cat_rename_input::{cat_id}").set_value("New Name").run(timeout=30)
        at.button(key=f"doc_repo_admin_cat_rename_btn::{cat_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        cats = {c["id"]: c["name"] for c in documents_repo.list_categories(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertEqual(cats[cat_id], "New Name")

    def test_admin_reorder_categories_move_up(self):
        cat_a = documents_repo.create_category(ADMIN_EMAIL, "Order A", secrets_getter=admin_secrets_getter)
        cat_b = documents_repo.create_category(ADMIN_EMAIL, "Order B", secrets_getter=admin_secrets_getter)
        before = [c["id"] for c in sorted(
            documents_repo.list_categories(ADMIN_EMAIL, secrets_getter=admin_secrets_getter),
            key=lambda c: c["sort_order"],
        )]
        self.assertEqual(before.index(cat_a), before.index(cat_b) - 1)  # A immediately before B

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.button(key=f"doc_repo_admin_cat_up::{cat_b}").click().run(timeout=30)
        self.assertFalse(at.exception)

        after = [c["id"] for c in sorted(
            documents_repo.list_categories(ADMIN_EMAIL, secrets_getter=admin_secrets_getter),
            key=lambda c: c["sort_order"],
        )]
        self.assertEqual(after.index(cat_b), after.index(cat_a) - 1)  # B now immediately before A

    def test_admin_deactivate_category_hides_it_but_keeps_documents(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Deactivate Me", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Surviving Doc", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at.button(key=f"doc_repo_admin_cat_deactivate_btn::{cat_id}").click().run(timeout=30)
        self.assertFalse(at.exception)

        cats = {c["id"]: c for c in documents_repo.list_categories(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertFalse(cats[cat_id]["is_active"])
        # Hidden from the non-admin viewer who had view access on it...
        self.assertEqual(documents_repo.list_categories(VIEWER_EMAIL, secrets_getter=admin_secrets_getter), [])
        # ...but the document itself is untouched, not deleted.
        detail = documents_repo.get_document_detail(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        self.assertIsNotNone(detail)
        self.assertEqual(detail["title"], "Surviving Doc")

        # Reactivating brings the category (and the document with it)
        # back into view for that same viewer.
        at.button(key=f"doc_repo_admin_cat_reactivate_btn::{cat_id}").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(len(documents_repo.list_categories(VIEWER_EMAIL, secrets_getter=admin_secrets_getter)), 1)
        docs = documents_repo.list_documents(VIEWER_EMAIL, category_id=cat_id, secrets_getter=admin_secrets_getter)
        self.assertEqual([d["id"] for d in docs], [doc_id])

    # -----------------------------------------------------------------
    # Admin tab: Users & Roles
    # -----------------------------------------------------------------
    def test_admin_add_user_via_ui(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Users & Roles")
        at.text_input(key="doc_repo_admin_new_user_email").set_value("New.User@Example.test").run(timeout=30)
        at.selectbox(key="doc_repo_admin_new_user_role").set_value("Authorized Contributor").run(timeout=30)
        at.button(key="doc_repo_admin_add_user_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        users = {u["email"]: u for u in documents_repo.list_doc_users(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertIn("new.user@example.test", users)  # lowercased
        self.assertEqual(users["new.user@example.test"]["role"], documents_repo.ROLE_CONTRIBUTOR)

    def test_admin_add_user_invalid_email_rejected(self):
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Users & Roles")
        at.text_input(key="doc_repo_admin_new_user_email").set_value("not-an-email").run(timeout=30)
        at.button(key="doc_repo_admin_add_user_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertTrue(any("valid email" in e.value for e in at.error))
        self.assertEqual(documents_repo.list_doc_users(ADMIN_EMAIL, secrets_getter=admin_secrets_getter), [])

    def test_admin_change_user_role_via_ui(self):
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Users & Roles")
        at.selectbox(key=f"doc_repo_admin_user_role::{VIEWER_EMAIL}").set_value("Authorized Contributor").run(timeout=30)
        at.button(key=f"doc_repo_admin_user_update_role_btn::{VIEWER_EMAIL}").click().run(timeout=30)
        self.assertFalse(at.exception)
        users = {u["email"]: u for u in documents_repo.list_doc_users(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertEqual(users[VIEWER_EMAIL]["role"], documents_repo.ROLE_CONTRIBUTOR)

    def test_admin_deactivate_and_reactivate_user_via_ui(self):
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Users & Roles")
        at.button(key=f"doc_repo_admin_user_deactivate_btn::{VIEWER_EMAIL}").click().run(timeout=30)
        self.assertFalse(at.exception)
        users = {u["email"]: u for u in documents_repo.list_doc_users(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertFalse(users[VIEWER_EMAIL]["is_active"])

        at.button(key=f"doc_repo_admin_user_reactivate_btn::{VIEWER_EMAIL}").click().run(timeout=30)
        self.assertFalse(at.exception)
        users = {u["email"]: u for u in documents_repo.list_doc_users(ADMIN_EMAIL, secrets_getter=admin_secrets_getter)}
        self.assertTrue(users[VIEWER_EMAIL]["is_active"])

    def test_admin_cannot_deactivate_own_only_admin_access(self):
        # ADMIN_EMAIL here is a doc_users admin row with NO matching
        # secret entry -- deactivating/role-changing it away from admin
        # would strip their only path to admin access.
        no_secret_admin = "doc-users-admin@example.test"
        documents_repo.create_doc_user(ADMIN_EMAIL, no_secret_admin, documents_repo.ROLE_ADMIN, secrets_getter=admin_secrets_getter)

        at = AppTest.from_file("app.py")
        at.secrets["DOC_REPO_ADMIN_EMAILS"] = ADMIN_EMAIL  # no_secret_admin is NOT in the secret
        at.session_state["_test_actor_email"] = no_secret_admin
        at.run(timeout=30)
        at.button(key="select_Document Repository").click().run(timeout=30)
        at = self._select_admin_section(at, "Users & Roles")
        self.assertFalse(at.exception)

        deactivate_btn = next(b for b in at.button if b.key == f"doc_repo_admin_user_deactivate_btn::{no_secret_admin}")
        self.assertTrue(deactivate_btn.disabled)

        role_select = at.selectbox(key=f"doc_repo_admin_user_role::{no_secret_admin}")
        role_select.set_value("Viewer").run(timeout=30)
        update_btn = next(b for b in at.button if b.key == f"doc_repo_admin_user_update_role_btn::{no_secret_admin}")
        self.assertTrue(update_btn.disabled)

        # The backend itself still allows it (UI gating is convenience
        # only) -- confirming this is a UI guard, not the real boundary.
        documents_repo.deactivate_doc_user(no_secret_admin, no_secret_admin, secrets_getter=lambda: {"DOC_REPO_ADMIN_EMAILS": ADMIN_EMAIL})
        self.assertFalse(documents_repo.is_admin(no_secret_admin, secrets_getter=lambda: {"DOC_REPO_ADMIN_EMAILS": ADMIN_EMAIL}))

    def test_secret_based_admin_self_deactivate_not_blocked_and_shows_note(self):
        # ADMIN_EMAIL is admin via the secret -- changing their doc_users
        # row (if they even have one) can never remove that, so the UI
        # must not block it, and should say so.
        documents_repo.create_doc_user(ADMIN_EMAIL, ADMIN_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Users & Roles")
        self.assertFalse(at.exception)

        deactivate_btn = next(b for b in at.button if b.key == f"doc_repo_admin_user_deactivate_btn::{ADMIN_EMAIL}")
        self.assertFalse(deactivate_btn.disabled)
        captions = " ".join(c.value for c in at.caption)
        self.assertIn("DOC_REPO_ADMIN_EMAILS secret", captions)

    # -----------------------------------------------------------------
    # Admin tab: per-category access grants
    # -----------------------------------------------------------------
    def test_admin_grant_change_and_revoke_access_via_ui(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Grant Flow", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Access Grants")

        at.selectbox(key="doc_repo_admin_access_category").set_value(cat_id).run(timeout=30)
        at.text_input(key="doc_repo_admin_grant_email").set_value(VIEWER_EMAIL).run(timeout=30)
        at.selectbox(key="doc_repo_admin_grant_level").set_value(documents_repo.LEVEL_VIEW).run(timeout=30)
        at.button(key="doc_repo_admin_grant_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        grants = {g["subject"]: g["level"] for g in documents_repo.list_category_access(ADMIN_EMAIL, cat_id, secrets_getter=admin_secrets_getter)}
        self.assertEqual(grants[VIEWER_EMAIL], documents_repo.LEVEL_VIEW)

        grant_row = documents_repo.list_category_access(ADMIN_EMAIL, cat_id, secrets_getter=admin_secrets_getter)[0]
        at.selectbox(key=f"doc_repo_admin_access_level::{grant_row['id']}").set_value(documents_repo.LEVEL_MANAGE).run(timeout=30)
        at.button(key=f"doc_repo_admin_access_save_btn::{grant_row['id']}").click().run(timeout=30)
        self.assertFalse(at.exception)
        grants = {g["subject"]: g["level"] for g in documents_repo.list_category_access(ADMIN_EMAIL, cat_id, secrets_getter=admin_secrets_getter)}
        self.assertEqual(grants[VIEWER_EMAIL], documents_repo.LEVEL_MANAGE)

        at.button(key=f"doc_repo_admin_access_revoke_btn::{grant_row['id']}").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(documents_repo.list_category_access(ADMIN_EMAIL, cat_id, secrets_getter=admin_secrets_getter), [])

    # -----------------------------------------------------------------
    # Admin tab: archived records across categories, with restore
    # -----------------------------------------------------------------
    def test_admin_archived_records_view_and_restore(self):
        cat_a = documents_repo.create_category(ADMIN_EMAIL, "Archived View A", secrets_getter=admin_secrets_getter)
        cat_b = documents_repo.create_category(ADMIN_EMAIL, "Archived View B", secrets_getter=admin_secrets_getter)
        doc_a = documents_repo.add_document(ADMIN_EMAIL, cat_a, "Doc A", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)
        doc_b = documents_repo.add_document(ADMIN_EMAIL, cat_b, "Doc B", PDF_BYTES, "b.pdf", secrets_getter=admin_secrets_getter)
        documents_repo.archive_document(ADMIN_EMAIL, doc_a, secrets_getter=admin_secrets_getter)
        documents_repo.archive_document(ADMIN_EMAIL, doc_b, secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Archived Records")
        self.assertFalse(at.exception)
        restore_buttons = {b.key for b in at.button if b.key and b.key.startswith("doc_repo_admin_restore_btn::")}
        self.assertEqual(restore_buttons, {f"doc_repo_admin_restore_btn::{doc_a}", f"doc_repo_admin_restore_btn::{doc_b}"})

        at.button(key=f"doc_repo_admin_restore_btn::{doc_a}").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(documents_repo.get_document_detail(ADMIN_EMAIL, doc_a, secrets_getter=admin_secrets_getter)["status"], documents_repo.STATUS_ACTIVE)
        self.assertEqual(documents_repo.get_document_detail(ADMIN_EMAIL, doc_b, secrets_getter=admin_secrets_getter)["status"], documents_repo.STATUS_ARCHIVED)

    def test_admin_archived_records_shows_count_of_total(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Archived Count", secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Doc", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)
        documents_repo.archive_document(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Archived Records")
        self.assertFalse(at.exception)
        captions = " ".join(c.value for c in at.caption)
        self.assertIn("Showing 1-1 of 1", captions)

    # -----------------------------------------------------------------
    # Admin tab: audit log viewer
    # -----------------------------------------------------------------
    def test_admin_audit_log_shows_newest_first_and_filters(self):
        documents_repo.create_category(ADMIN_EMAIL, "Audit Cat One", secrets_getter=admin_secrets_getter)
        documents_repo.create_category(ADMIN_EMAIL, "Audit Cat Two", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Audit Log")
        self.assertFalse(at.exception)

        audit_tables = [df for df in at.dataframe if "Action" in df.value.columns]
        self.assertEqual(len(audit_tables), 1)
        actions = audit_tables[0].value["Action"].tolist()
        self.assertIn("create_category", actions)

        at.text_input(key="doc_repo_admin_audit_action_filter").set_value("create_category").run(timeout=30)
        self.assertFalse(at.exception)
        audit_tables = [df for df in at.dataframe if "Action" in df.value.columns]
        self.assertTrue(all(a == "create_category" for a in audit_tables[0].value["Action"].tolist()))

        # Count caption reflects the active filter, not the whole log.
        filtered_count = documents_repo.count_audit_log(
            ADMIN_EMAIL, action_filter="create_category", secrets_getter=admin_secrets_getter,
        )
        captions = " ".join(c.value for c in at.caption)
        self.assertIn(f"Showing 1-{filtered_count} of {filtered_count}", captions)

    def test_admin_audit_log_to_today_still_shows_todays_rows(self):
        # Regression test for the date_to off-by-one: "To" defaulting
        # to today must not exclude rows created moments ago today --
        # ts is a full timestamp, so a naive "ts <= today" would have
        # excluded everything from today.
        documents_repo.create_category(ADMIN_EMAIL, "Audit Today", secrets_getter=admin_secrets_getter)

        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        at = self._select_admin_section(at, "Audit Log")
        at.date_input(key="doc_repo_admin_audit_date_to").set_value(date.today()).run(timeout=30)
        self.assertFalse(at.exception)

        audit_tables = [df for df in at.dataframe if "Action" in df.value.columns]
        self.assertEqual(len(audit_tables), 1)
        self.assertIn("create_category", audit_tables[0].value["Action"].tolist())

    def test_admin_audit_log_survives_malicious_text_without_crashing(self):
        documents_repo.create_category(ADMIN_EMAIL, "![x](http://example.test/a.png) [click](http://example.test)", secrets_getter=admin_secrets_getter)
        at = self._at_as(ADMIN_EMAIL)
        at = self._select_document_repository(at)
        self.assertFalse(at.exception)
        # Exercise the Audit Log section specifically, since that's what
        # this test is proving survives malicious category names.
        at = self._select_admin_section(at, "Audit Log")
        self.assertFalse(at.exception)

    # -----------------------------------------------------------------
    # Audit completeness: one row per mutation, correct actor/action
    # -----------------------------------------------------------------
    def test_audit_completeness_for_admin_mutations(self):
        cat_id = documents_repo.create_category(ADMIN_EMAIL, "Audit Completeness", secrets_getter=admin_secrets_getter)
        documents_repo.rename_category(ADMIN_EMAIL, cat_id, "Renamed", secrets_getter=admin_secrets_getter)
        documents_repo.deactivate_category(ADMIN_EMAIL, cat_id, secrets_getter=admin_secrets_getter)
        documents_repo.reactivate_category(ADMIN_EMAIL, cat_id, secrets_getter=admin_secrets_getter)
        documents_repo.create_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_VIEWER, secrets_getter=admin_secrets_getter)
        documents_repo.update_doc_user_role(ADMIN_EMAIL, VIEWER_EMAIL, documents_repo.ROLE_CONTRIBUTOR, secrets_getter=admin_secrets_getter)
        documents_repo.deactivate_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, secrets_getter=admin_secrets_getter)
        documents_repo.reactivate_doc_user(ADMIN_EMAIL, VIEWER_EMAIL, secrets_getter=admin_secrets_getter)
        documents_repo.grant_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, documents_repo.LEVEL_VIEW, secrets_getter=admin_secrets_getter)
        documents_repo.revoke_category_access(ADMIN_EMAIL, cat_id, VIEWER_EMAIL, secrets_getter=admin_secrets_getter)
        doc_id = documents_repo.add_document(ADMIN_EMAIL, cat_id, "Doc", PDF_BYTES, "a.pdf", secrets_getter=admin_secrets_getter)
        documents_repo.archive_document(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)
        documents_repo.restore_document(ADMIN_EMAIL, doc_id, secrets_getter=admin_secrets_getter)

        for action in (
            "create_category", "rename_category", "deactivate_category", "reactivate_category",
            "create_doc_user", "update_doc_user_role", "deactivate_doc_user", "reactivate_doc_user",
            "grant_category_access", "revoke_category_access",
            "add_document", "archive_document", "restore_document",
        ):
            rows = documents_repo.list_audit_log(ADMIN_EMAIL, action_filter=action, secrets_getter=admin_secrets_getter, page_size=100)
            matching = [r for r in rows if r["action"] == action]
            self.assertEqual(len(matching), 1, f"expected exactly one {action!r} row, got {len(matching)}")
            self.assertEqual(matching[0]["actor_email"], ADMIN_EMAIL)


if __name__ == "__main__":
    unittest.main()
