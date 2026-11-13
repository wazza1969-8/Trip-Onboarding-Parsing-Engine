"""
Automated tests for the Document Repository tile/screen (Phase 2:
browse/search/download only) -- both the app.py hook points and
documents_ui.py's rendering, driven through Streamlit's AppTest
harness, same pattern as test_weekly_report.py's UITests.

Covers: the tile hidden by default (including the "in password mode"
case, as far as AppTest can actually exercise it -- see
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
import sqlite3
import sys
import unittest
from unittest import mock

import documents_repo

try:
    from streamlit.testing.v1 import AppTest
    _APPTEST_AVAILABLE = True
except ImportError:
    _APPTEST_AVAILABLE = False

ADMIN_EMAIL = "admin@example.test"
VIEWER_EMAIL = "viewer@example.test"
OUTSIDER_EMAIL = "outsider@example.test"

PDF_BYTES = b"%PDF-1.4\n%fake pdf content for tests\n"


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
        self.assertEqual(len(at.dataframe), 1)
        titles = at.dataframe[0].value["Title"].tolist()
        self.assertIn("Blob Isolation Doc", titles)

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


if __name__ == "__main__":
    unittest.main()
