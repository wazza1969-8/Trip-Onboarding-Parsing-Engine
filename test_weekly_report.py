"""
Automated tests for the Weekly Report reporting-period feature:
  - the timezone-aware period/cutoff business rules in backend.py
  - server-side enforcement of the edit-cutoff in upsert_weekly_report()
  - the legacy-period-format data migration
  - the shared Period dropdown in app.py (default selection, session
    persistence, empty-period display)

Run with:
    pip install pytest
    pytest test_weekly_report.py -v

or directly:
    python3 test_weekly_report.py

All backend tests pass an explicit `now`/`tz` into the functions under
test rather than relying on the real clock, so they're deterministic
regardless of when they're run.
"""
import os
import sqlite3
import tempfile
import unittest
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import backend

TZ = ZoneInfo("America/Denver")


def dt(y, m, d, h=12, mi=0, s=0):
    return datetime(y, m, d, h, mi, s, tzinfo=TZ)


# ---------------------------------------------------------------------------
# Pure period/timezone math -- no database involved.
# ---------------------------------------------------------------------------
class PeriodMathTests(unittest.TestCase):
    def test_current_period_monday_mid_week(self):
        # Aug 31, 2026 is a Monday; Sep 3 falls in the same Mon-Sun period.
        self.assertEqual(backend.current_period_monday(TZ, dt(2026, 9, 3)), date(2026, 8, 31))

    def test_current_period_monday_on_sunday(self):
        # The period's own Sunday still belongs to that period.
        self.assertEqual(backend.current_period_monday(TZ, dt(2026, 9, 6, 23, 30)), date(2026, 8, 31))

    def test_available_periods_five_rolling_weeks(self):
        periods = backend.available_periods(TZ, dt(2026, 8, 31))
        self.assertEqual(
            periods,
            [
                date(2026, 8, 17), date(2026, 8, 24), date(2026, 8, 31),
                date(2026, 9, 7), date(2026, 9, 14),
            ],
        )

    def test_period_id_roundtrip(self):
        monday = date(2026, 8, 31)
        pid = backend.period_id_from_monday(monday)
        self.assertEqual(pid, "2026-08-31")
        self.assertEqual(backend.monday_from_period_id(pid), monday)

    def test_monday_from_period_id_accepts_legacy_iso_week_format(self):
        self.assertEqual(backend.monday_from_period_id("2026-W36"), date(2026, 8, 31))

    def test_format_period_label_same_month(self):
        self.assertEqual(backend.format_period_label(date(2026, 8, 17)), "Aug 17–23, 2026")

    def test_format_period_label_month_boundary(self):
        self.assertEqual(backend.format_period_label(date(2026, 8, 31)), "Aug 31–Sep 6, 2026")

    def test_format_period_label_current_suffix(self):
        self.assertEqual(
            backend.format_period_label(date(2026, 8, 31), is_current=True),
            "Aug 31–Sep 6, 2026 — Current",
        )

    def test_format_period_label_no_leading_zero_on_day(self):
        label = backend.format_period_label(date(2026, 9, 7))
        self.assertNotIn("07", label)
        self.assertIn("Sep 7", label)

    def test_year_end_transition_label(self):
        # Dec 28, 2026 is a Monday; its Sunday (Jan 3, 2027) is next year.
        self.assertEqual(date(2026, 12, 28).weekday(), 0)
        label = backend.format_period_label(date(2026, 12, 28))
        self.assertEqual(label, "Dec 28, 2026–Jan 3, 2027")

    def test_year_end_transition_available_periods(self):
        periods = backend.available_periods(TZ, dt(2026, 12, 28))
        self.assertEqual(periods[0], date(2026, 12, 14))
        self.assertEqual(periods[-1], date(2027, 1, 11))
        self.assertEqual(periods[-1].year, 2027)

    def test_month_end_transition_available_periods(self):
        periods = backend.available_periods(TZ, dt(2026, 8, 31))
        self.assertIn(date(2026, 8, 31), periods)
        self.assertEqual(backend.format_period_label(date(2026, 8, 31)), "Aug 31–Sep 6, 2026")

    def test_dst_transition_cutoff_stays_local_2359(self):
        # Find an actual DST transition in America/Denver by scanning
        # forward through the year, instead of hardcoding a date that
        # could be wrong for a given year's DST rules.
        prev_offset = None
        transition_day = None
        d = date(2026, 1, 2)
        while d.year == 2026:
            offset = datetime(d.year, d.month, d.day, 12, tzinfo=TZ).utcoffset()
            if prev_offset is not None and offset != prev_offset:
                transition_day = d
                break
            prev_offset = offset
            d += timedelta(days=1)
        self.assertIsNotNone(transition_day, "expected a DST transition somewhere in 2026")

        # A period whose grace-day Monday cutoff lands in the same week
        # as the transition -- confirm edit_cutoff() still resolves to
        # 23:59:59 local wall-clock time (zoneinfo handles the actual
        # UTC-offset shift), not a UTC-shifted hour.
        grace_monday = transition_day - timedelta(days=transition_day.weekday())
        period_monday = grace_monday - timedelta(weeks=1)
        cutoff = backend.edit_cutoff(period_monday, TZ)
        self.assertEqual((cutoff.hour, cutoff.minute, cutoff.second), (23, 59, 59))
        self.assertEqual(cutoff.tzinfo, TZ)


# ---------------------------------------------------------------------------
# is_period_editable() -- the acceptance-criteria scenarios, by name.
# ---------------------------------------------------------------------------
class EditabilityTests(unittest.TestCase):
    def setUp(self):
        self.current_monday = date(2026, 8, 31)  # a Monday

    def test_sunday_before_period_ends_is_editable(self):
        now = dt(2026, 9, 6, 22, 0)  # Sunday night, still inside the current period
        self.assertTrue(backend.is_period_editable(self.current_monday, TZ, now))

    def test_monday_during_grace_period_is_editable(self):
        previous_monday = date(2026, 8, 24)  # period that ended Aug 30
        now = dt(2026, 8, 31, 9, 0)  # grace-day Monday morning
        self.assertTrue(backend.is_period_editable(previous_monday, TZ, now))

    def test_monday_grace_period_last_second_is_editable(self):
        previous_monday = date(2026, 8, 24)
        now = dt(2026, 8, 31, 23, 59, 58)
        self.assertTrue(backend.is_period_editable(previous_monday, TZ, now))

    def test_tuesday_after_grace_period_is_locked(self):
        previous_monday = date(2026, 8, 24)
        now = dt(2026, 9, 1, 0, 0, 1)  # just after the Monday cutoff
        self.assertFalse(backend.is_period_editable(previous_monday, TZ, now))

    def test_current_period_always_editable(self):
        for now in (dt(2026, 8, 31, 0, 1), dt(2026, 9, 3, 15, 0), dt(2026, 9, 6, 23, 59)):
            self.assertTrue(backend.is_period_editable(self.current_monday, TZ, now))

    def test_both_displayed_future_periods_editable(self):
        now = dt(2026, 8, 31, 12, 0)
        for offset in (1, 2):
            self.assertTrue(
                backend.is_period_editable(self.current_monday + timedelta(weeks=offset), TZ, now)
            )

    def test_both_displayed_past_periods_locked_on_a_normal_day(self):
        now = dt(2026, 9, 2, 12, 0)  # a Wednesday, well past any grace day
        for offset in (1, 2):
            self.assertFalse(
                backend.is_period_editable(self.current_monday - timedelta(weeks=offset), TZ, now)
            )

    def test_period_outside_rolling_window_never_editable_future(self):
        now = dt(2026, 8, 31, 12, 0)
        self.assertFalse(
            backend.is_period_editable(self.current_monday + timedelta(weeks=3), TZ, now)
        )

    def test_period_outside_rolling_window_never_editable_past(self):
        now = dt(2026, 8, 31, 12, 0)
        self.assertFalse(
            backend.is_period_editable(self.current_monday - timedelta(weeks=3), TZ, now)
        )


# ---------------------------------------------------------------------------
# Server-side write enforcement, duplicate prevention, ownership
# isolation, empty-period reads, and readability of locked reports.
# ---------------------------------------------------------------------------
class BackendWriteTests(unittest.TestCase):
    def setUp(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.db_path = path
        self.values = {key: "" for key, _ in backend.WEEKLY_REPORT_CATEGORIES}
        self.values["highlights"] = "Did a thing"

    def tearDown(self):
        os.remove(self.db_path)

    def test_current_period_write_succeeds(self):
        period_id = backend.period_id_from_monday(date(2026, 8, 31))
        backend.upsert_weekly_report(
            "Alice", period_id, self.values, db_path=self.db_path, tz=TZ, now=dt(2026, 9, 2, 10, 0)
        )
        row = backend.get_weekly_report("Alice", period_id, db_path=self.db_path)
        self.assertIsNotNone(row)
        self.assertEqual(row["highlights"], "Did a thing")

    def test_backend_rejects_write_after_cutoff_stale_browser(self):
        # Simulates a browser tab left open past the deadline: the
        # period was editable when the page loaded, but server time at
        # the moment of the actual write is past the cutoff.
        period_id = backend.period_id_from_monday(date(2026, 8, 24))
        with self.assertRaises(backend.PeriodLockedError):
            backend.upsert_weekly_report(
                "Alice", period_id, self.values, db_path=self.db_path, tz=TZ,
                now=dt(2026, 9, 1, 0, 0, 1),
            )
        self.assertIsNone(backend.get_weekly_report("Alice", period_id, db_path=self.db_path))

    def test_grace_day_write_still_succeeds(self):
        period_id = backend.period_id_from_monday(date(2026, 8, 24))
        backend.upsert_weekly_report(
            "Alice", period_id, self.values, db_path=self.db_path, tz=TZ, now=dt(2026, 8, 31, 20, 0)
        )
        self.assertIsNotNone(backend.get_weekly_report("Alice", period_id, db_path=self.db_path))

    def test_duplicate_report_prevention_updates_in_place(self):
        period_id = backend.period_id_from_monday(date(2026, 8, 31))
        now = dt(2026, 9, 1, 10, 0)
        backend.upsert_weekly_report("Alice", period_id, self.values, db_path=self.db_path, tz=TZ, now=now)
        updated_values = dict(self.values, highlights="Updated thing")
        backend.upsert_weekly_report("Alice", period_id, updated_values, db_path=self.db_path, tz=TZ, now=now)
        rows = backend.list_weekly_reports_for_period(period_id, db_path=self.db_path)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["highlights"], "Updated thing")

    def test_existing_ownership_rule_writing_as_one_submitter_does_not_touch_another(self):
        period_id = backend.period_id_from_monday(date(2026, 8, 31))
        now = dt(2026, 9, 1, 10, 0)
        backend.upsert_weekly_report("Alice", period_id, self.values, db_path=self.db_path, tz=TZ, now=now)
        bob_values = dict(self.values, highlights="Bob's own update")
        backend.upsert_weekly_report("Bob", period_id, bob_values, db_path=self.db_path, tz=TZ, now=now)

        alice_row = backend.get_weekly_report("Alice", period_id, db_path=self.db_path)
        bob_row = backend.get_weekly_report("Bob", period_id, db_path=self.db_path)
        self.assertEqual(alice_row["highlights"], "Did a thing")
        self.assertEqual(bob_row["highlights"], "Bob's own update")

        rows = backend.list_weekly_reports_for_period(period_id, db_path=self.db_path)
        self.assertEqual(len(rows), 2)

    def test_empty_period_returns_no_rows(self):
        period_id = backend.period_id_from_monday(date(2026, 9, 14))
        self.assertEqual(backend.list_weekly_reports_for_period(period_id, db_path=self.db_path), [])

    def test_locked_report_remains_readable(self):
        # "Locked" only blocks new writes -- a report submitted while a
        # period was still open must remain readable after it locks.
        period_id = backend.period_id_from_monday(date(2026, 8, 24))
        backend.upsert_weekly_report(
            "Alice", period_id, self.values, db_path=self.db_path, tz=TZ, now=dt(2026, 8, 28, 10, 0)
        )
        # "Now" has since moved well past the cutoff -- reads take no
        # `now`/editability argument at all, so this alone proves reads
        # aren't gated by the lock.
        row = backend.get_weekly_report("Alice", period_id, db_path=self.db_path)
        self.assertIsNotNone(row)
        self.assertEqual(row["highlights"], "Did a thing")


# ---------------------------------------------------------------------------
# Legacy period-format migration.
# ---------------------------------------------------------------------------
class MigrationTests(unittest.TestCase):
    def setUp(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.db_path = path

    def tearDown(self):
        os.remove(self.db_path)

    def test_legacy_iso_week_period_migrates_to_monday_date(self):
        # Simulate a pre-migration row written by an older version of
        # this app, which stored period as "YYYY-Www".
        conn = sqlite3.connect(self.db_path)
        conn.execute(
            """CREATE TABLE weekly_reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT, submitter TEXT NOT NULL,
                period TEXT NOT NULL, highlights TEXT, opportunities TEXT,
                watch_items TEXT, upcoming_events TEXT, other TEXT,
                submitted_at TEXT, updated_at TEXT, UNIQUE(submitter, period)
            )"""
        )
        conn.execute(
            "INSERT INTO weekly_reports (submitter, period, highlights, submitted_at, updated_at) "
            "VALUES ('Alice', '2026-W36', 'legacy content', "
            "'2026-09-01T00:00:00+00:00', '2026-09-01T00:00:00+00:00')"
        )
        conn.commit()
        conn.close()

        # Any normal backend call opens a connection via get_conn(),
        # which runs the migration as a side effect.
        row = backend.get_weekly_report("Alice", "2026-08-31", db_path=self.db_path)
        self.assertIsNotNone(row, "expected the legacy period to be migrated to 2026-08-31")
        self.assertEqual(row["highlights"], "legacy content")
        self.assertIsNone(backend.get_weekly_report("Alice", "2026-W36", db_path=self.db_path))

    def test_migration_is_idempotent(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute(
            """CREATE TABLE weekly_reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT, submitter TEXT NOT NULL,
                period TEXT NOT NULL, highlights TEXT, opportunities TEXT,
                watch_items TEXT, upcoming_events TEXT, other TEXT,
                submitted_at TEXT, updated_at TEXT, UNIQUE(submitter, period)
            )"""
        )
        conn.execute(
            "INSERT INTO weekly_reports (submitter, period, highlights, submitted_at, updated_at) "
            "VALUES ('Alice', '2026-W36', 'x', 'now', 'now')"
        )
        conn.commit()
        conn.close()

        backend.get_conn(self.db_path).close()
        backend.get_conn(self.db_path).close()  # run the migration again
        rows = backend.list_weekly_reports_for_period("2026-08-31", db_path=self.db_path)
        self.assertEqual(len(rows), 1)


# ---------------------------------------------------------------------------
# UI: default period selection, session persistence, empty state.
# Requires streamlit's AppTest harness.
# ---------------------------------------------------------------------------
try:
    from streamlit.testing.v1 import AppTest
    _APPTEST_AVAILABLE = True
except ImportError:
    _APPTEST_AVAILABLE = False


@unittest.skipUnless(_APPTEST_AVAILABLE, "streamlit.testing.v1.AppTest not available")
class UITests(unittest.TestCase):
    """These exercise app.py's real Streamlit script via AppTest. Every
    backend.py function's `db_path` parameter defaults to the module-
    level DB_PATH *at function-definition time* (a Python late-binding
    gotcha), so reassigning backend.DB_PATH after import does not
    redirect app.py's internal calls to a temp file -- app.py always
    calls backend functions without an explicit db_path, so it always
    hits the real default database. Isolation is done here by clearing
    the relevant tables in that real database before and after each
    test instead of swapping the file."""

    def setUp(self):
        self._clear_weekly_report_tables()

    def tearDown(self):
        self._clear_weekly_report_tables()

    @staticmethod
    def _clear_weekly_report_tables():
        conn = backend.get_conn(backend.DB_PATH)
        conn.execute("DELETE FROM weekly_reports")
        conn.execute("DELETE FROM report_team_members")
        conn.execute("DELETE FROM report_recipients")
        conn.commit()
        conn.close()

    def _open_weekly_report(self):
        at = AppTest.from_file("app.py")
        at.run(timeout=30)
        at.button(key="select_Weekly Report").click().run(timeout=30)
        return at

    def test_default_dropdown_selection_is_current_period(self):
        at = self._open_weekly_report()
        self.assertFalse(at.exception)
        period_select = at.selectbox(key="wr_selected_period")
        self.assertEqual(period_select.value, backend.current_period())

    def test_dropdown_has_exactly_five_options(self):
        at = self._open_weekly_report()
        self.assertEqual(len(at.selectbox(key="wr_selected_period").options), 5)

    def test_user_selection_persists_across_a_rerender(self):
        at = self._open_weekly_report()
        options = at.selectbox(key="wr_selected_period").options  # display labels
        at.selectbox(key="wr_selected_period").set_value(options[0]).run(timeout=30)
        selected_after_change = at.selectbox(key="wr_selected_period").value

        # An unrelated rerender (opening the team-manager panel) must
        # not reset the period selection back to the default.
        at.button(key="toggle_team_manager").click().run(timeout=30)
        self.assertFalse(at.exception)
        self.assertEqual(at.selectbox(key="wr_selected_period").value, selected_after_change)
        self.assertNotEqual(selected_after_change, backend.current_period())

    def _add_team_member(self, at, name):
        at.button(key="toggle_team_manager").click().run(timeout=30)
        at.text_input(key="new_team_member_name").set_value(name).run(timeout=30)
        at.button(key="add_team_member_btn").click().run(timeout=30)
        self.assertFalse(at.exception)
        return at

    def test_empty_period_shows_helpful_message_in_view_all_reports(self):
        at = self._open_weekly_report()
        at = self._add_team_member(at, "Test Person")
        info_texts = " ".join(i.value for i in at.info)
        self.assertIn("No reports have been submitted yet for this period.", info_texts)

    def test_locked_report_is_viewable_not_editable(self):
        at = self._open_weekly_report()
        at = self._add_team_member(at, "Locked Person")

        # Seed a report directly for the "two weeks back" period, as if
        # it had been submitted while that period was still open. That
        # period is genuinely in the past relative to real wall-clock
        # time, so it's already locked right now -- no time injection
        # needed to exercise the UI's locked/read-only path for real.
        two_weeks_back_monday = backend.available_periods()[0]
        period_id = backend.period_id_from_monday(two_weeks_back_monday)
        values = {key: "" for key, _ in backend.WEEKLY_REPORT_CATEGORIES}
        values["highlights"] = "Locked content"
        backend.upsert_weekly_report(
            "Locked Person", period_id, values,
            now=datetime.combine(
                two_weeks_back_monday, datetime.min.time(), tzinfo=backend.get_business_timezone()
            ),
        )
        self.assertFalse(backend.is_period_editable(two_weeks_back_monday))  # sanity check

        options = at.selectbox(key="wr_selected_period").options
        at.selectbox(key="wr_selected_period").set_value(options[0]).run(timeout=30)
        self.assertFalse(at.exception)

        # Submit My Report tab: read-only content + lock notice, no
        # Submit button for this locked period.
        at.selectbox(key="wr_submitter").set_value("Locked Person").run(timeout=30)
        self.assertFalse(at.exception)
        markdown_texts = " ".join(m.value for m in at.markdown)
        self.assertIn("Locked content", markdown_texts)
        info_texts = " ".join(i.value for i in at.info)
        self.assertIn("locked on", info_texts)
        self.assertFalse(any(w.key == "wr_submit_btn" for w in at.button))

        # View All Reports tab: this person's action must be "View",
        # never "Edit", once the period is locked.
        self.assertTrue(any(b.key == "wrview_view_btn_0" for b in at.button))
        self.assertFalse(any(b.key == "wrview_edit_btn_0" for b in at.button))


if __name__ == "__main__":
    unittest.main(verbosity=2)
