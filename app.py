"""
ITP Sales Hub -- combined intake + parsing app
================================================

Landing screen lets the user pick which parsing task they want to run:
  - Trip Preference Parsing: the original onboarding-form pipeline
    (Trip Support Service Preferences PDF -> Operational Information
    Excel). Fully wired to parsing_engine.py via backend.py.
  - Flight Planning Parsing: parses a completed Flight Planning
    Preferences form into the Flight Plan Template. Not wired up yet --
    see the note in that section below.
  - Weekly Report: team members submit their update for the current
    week under a fixed set of categories; a manager compiles everyone's
    submissions into one report and can send it by email with one
    click (via Gmail's SMTP relay -- see GMAIL_ADDRESS/GMAIL_APP_PASSWORD
    in secrets.toml), or copy/download it instead.

Every Trip Preference Parsing submission becomes a row in a local
SQLite database (requests.db, created next to backend.py) with a short
request ID and a status that moves through:

    Submitted -> Processing -> Complete
                             -> Failed  (with an error message + Retry)

The onboarding form upload is required -- Submit is blocked until a
file is attached, since there's nothing for the parser to run against
otherwise.

Requirements
------------
    pip install streamlit anthropic pymupdf openpyxl
    parsing_engine.py and backend.py must be in the same directory as
    this file.

Configuring the API key
------------------------
    Local run:   export ANTHROPIC_API_KEY=...
    Streamlit Community Cloud: add ANTHROPIC_API_KEY under
        App settings -> Secrets, as:  ANTHROPIC_API_KEY = "sk-ant-..."

Run locally
-----------
    streamlit run app.py

Deploying for a few users
--------------------------
    Push this file, backend.py, parsing_engine.py, and a
    requirements.txt (streamlit, anthropic, pymupdf, openpyxl) to a
    GitHub repo, then deploy free on Streamlit Community Cloud for a
    persistent URL. Note: that platform's filesystem resets on
    redeploys, so requests.db is fine for day-to-day use but isn't a
    permanent archive -- move to a hosted database if you need requests
    to survive indefinitely.
"""

import hmac
import io
import json
import os
import time
import uuid

import streamlit as st

import backend
# Not a new dependency: backend.py already imports this module, so if it
# were missing the app would not have got this far. Imported here by
# name anyway rather than reached for as backend.pe, so the Entry-Ready
# View's use of SUBJECTS / map_to_subjects is visible at the top of the
# file instead of hidden inside another module's namespace.
import parsing_engine as pe

# Document Repository (Phase 2: browse/search/download tile + screen).
# Imported in try/except on purpose -- if either module is missing or
# fails to import for any reason, the tile is simply hidden and nothing
# else in this file is affected. Not a new dependency: both modules are
# pure standard library plus streamlit, same as the rest of this app.
try:
    import documents_repo
    import documents_ui
    _DOCUMENTS_REPO_AVAILABLE = True
except Exception:
    documents_repo = None
    documents_ui = None
    _DOCUMENTS_REPO_AVAILABLE = False

# True only when the process was launched by pytest -- which is how the
# UI tests in test_weekly_report.py drive this file via AppTest. This
# can NEVER be true in the deployed app (Streamlit Cloud doesn't run
# pytest to serve it), so unlike a manually-set flag, there's no way for
# this to end up accidentally left on in production. It exists solely
# because Streamlit's AppTest harness has no supported way to simulate a
# real logged-in OIDC identity, so the auth gate below has to be skipped
# for that harness specifically -- everything else in the app is tested
# exactly as it runs for a real user.
_RUNNING_UNDER_TEST = "PYTEST_CURRENT_TEST" in os.environ

# Icons shown next to each Weekly Report category -- purely cosmetic,
# just for quick visual scanning; the underlying category keys/labels
# still come from backend.WEEKLY_REPORT_CATEGORIES.
WEEKLY_REPORT_CATEGORY_ICONS = {
    "highlights": "\U0001F31F",       # star
    "opportunities": "\U0001F91D",    # handshake
    "watch_items": "⚠️",    # warning
    "upcoming_events": "\U0001F4C5",  # calendar
    "other": "\U0001F4DD",            # memo
}


def render_editable_report_form(submitter: str, period_id: str, key_prefix: str) -> None:
    """The dynamic add/remove bullet-point editor for one submitter's
    report in one period. Shared by the "Submit My Report" tab (editing
    "your own" report, key_prefix="wr") and the "View All Reports" tab
    (a manager creating/editing someone else's report, key_prefix
    scoped per row) so both use the exact same widget behavior and the
    exact same backend.upsert_weekly_report() write path -- including
    its server-side edit-cutoff check. `key_prefix` keeps the two call
    sites' widget keys from colliding if both were ever open in the
    same session.

    Caller is responsible for only invoking this when
    backend.is_period_editable() is true for this period; the actual
    save is still re-validated server-side regardless, in case the
    period locks between page load and the Submit click.
    """
    existing = backend.get_weekly_report(submitter, period_id)
    if existing:
        st.caption("A report already exists for this period -- editing and submitting again will update it.")

    # Bullet widget state is only (re)loaded from the saved report when
    # the submitter being edited changes -- not on every rerun -- so
    # add/remove clicks don't wipe unsaved edits.
    load_marker_key = f"{key_prefix}_loaded_for_{period_id}"
    if st.session_state.get(load_marker_key) != submitter:
        for key, _ in backend.WEEKLY_REPORT_CATEGORIES:
            saved_text = existing[key] if existing and existing[key] else ""
            saved_lines = [ln.strip() for ln in saved_text.split("\n") if ln.strip()]
            new_ids = []
            for line in saved_lines:
                item_id = uuid.uuid4().hex[:8]
                st.session_state[f"{key_prefix}_{key}_{item_id}"] = line
                new_ids.append(item_id)
            st.session_state[f"{key_prefix}_{key}_ids"] = new_ids
        st.session_state[load_marker_key] = submitter

    values = {}
    for key, label in backend.WEEKLY_REPORT_CATEGORIES:
        icon = WEEKLY_REPORT_CATEGORY_ICONS.get(key, "•")
        ids_key = f"{key_prefix}_{key}_ids"
        if ids_key not in st.session_state:
            st.session_state[ids_key] = []
        item_ids = st.session_state[ids_key]

        with st.container(border=True):
            count = len(item_ids)
            st.markdown(
                f"##### {icon} {label} &nbsp; "
                f":gray[({count} {'entry' if count == 1 else 'entries'})]"
            )

            if not item_ids:
                st.caption("No entries yet -- click “+ Add entry” to add your first bullet point.")

            # Important: don't call st.rerun() from inside this loop.
            # Streamlit only keeps a widget's value if that widget
            # gets (re-)rendered during the run; rerun() aborts the
            # script immediately, so any row after the one clicked
            # would never render this pass and would lose its text.
            # So every row's widget always renders first, and the
            # delete is only actioned -- and rerun() only called --
            # once the whole loop has finished.
            pending_delete_id = None
            for item_id in list(item_ids):
                # Each row gets its own explicitly-keyed container so
                # Streamlit tracks its widgets by that key rather
                # than by position.
                with st.container(key=f"{key_prefix}_{key}_row_{item_id}"):
                    col_bullet, col_text, col_del = st.columns([0.05, 0.85, 0.10])
                    with col_bullet:
                        st.markdown(
                            "<div style='font-size:1.6rem; line-height:2.4rem; "
                            "color:#5fd0e8; text-align:center;'>&bull;</div>",
                            unsafe_allow_html=True,
                        )
                    with col_text:
                        st.text_input(
                            label,
                            label_visibility="collapsed",
                            key=f"{key_prefix}_{key}_{item_id}",
                            placeholder=f"Add a {label.split(' / ')[0].lower()} entry...",
                        )
                    with col_del:
                        if st.button("✕", key=f"{key_prefix}_{key}_del_{item_id}", help="Remove this entry"):
                            pending_delete_id = item_id

            if pending_delete_id is not None:
                st.session_state[ids_key] = [i for i in st.session_state[ids_key] if i != pending_delete_id]
                st.session_state.pop(f"{key_prefix}_{key}_{pending_delete_id}", None)
                st.rerun()

            if st.button("+ Add entry", key=f"{key_prefix}_{key}_add_btn"):
                new_id = uuid.uuid4().hex[:8]
                st.session_state[ids_key] = st.session_state[ids_key] + [new_id]
                st.session_state[f"{key_prefix}_{key}_{new_id}"] = ""
                st.rerun()

        bullet_texts = [st.session_state.get(f"{key_prefix}_{key}_{i}", "").strip() for i in st.session_state[ids_key]]
        values[key] = "\n".join(t for t in bullet_texts if t)

    if st.button("Submit report", key=f"{key_prefix}_submit_btn", type="primary"):
        try:
            backend.upsert_weekly_report(submitter, period_id, values)
            st.toast(f"Saved the report for {submitter}.")
            st.rerun()
        except backend.PeriodLockedError as e:
            st.error(str(e))


EXPORT_SHEET_NAME = "Operational Information"


def by_subject_from_workbook(data: bytes) -> tuple[dict[str, list[str]] | None, str | None]:
    """Reads the Excel file this app exports (write_workbook in
    parsing_engine.py) back into the {subject: [comment text]} shape the
    Entry-Ready View renders.

    The export is a sheet of Subject / Comments rows, and each Comments
    cell is exactly "\\n\\n".join(blocks) for that Subject -- the same string
    the view joins and shows. So nothing needs re-deriving: the cell text
    goes back in as a single block and displays character-for-character
    as exported. Because the file was written after the request's own
    account number, account exec name and High Profile Notes month were
    filled in, those come through as real values -- unlike a raw .json
    extraction, which only has what was in the PDF.

    Returns (by_subject, None) on success or (None, message) on failure.
    Read-only: nothing is written anywhere. Extra columns, extra sheets
    and blank rows are tolerated; Subjects this engine does not know are
    ignored rather than shown."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return None, "That file could not be read as an Excel workbook (.xlsx)."

    ws = wb[EXPORT_SHEET_NAME] if EXPORT_SHEET_NAME in wb.sheetnames else wb.active

    header = [str(c.value or "").strip().lower() for c in next(ws.iter_rows(min_row=1, max_row=1), [])]
    if "subject" not in header or "comments" not in header:
        return None, (
            "That workbook does not have 'Subject' and 'Comments' columns in its first "
            "row, so it is not an Operational Information export from this tool."
        )
    subj_col, comm_col = header.index("subject"), header.index("comments")

    known = {s.lower(): s for s in pe.SUBJECTS}
    by_subject: dict[str, list[str]] = {s: [] for s in pe.SUBJECTS}
    matched = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if subj_col >= len(row):
            continue
        name = str(row[subj_col] or "").strip()
        subject = known.get(name.lower())
        if subject is None:
            continue
        matched += 1
        comments = str(row[comm_col] if comm_col < len(row) and row[comm_col] is not None else "").strip()
        if comments:
            by_subject[subject].append(comments)

    if matched == 0:
        return None, (
            "None of the Subject names in that workbook match this tool's Operational "
            "Information subjects, so it is not an export from this tool."
        )
    return by_subject, None


def customer_name_from_by_subject(by_subject: dict[str, list[str]] | None) -> str:
    """The customer's name as the parser extracted it: the 'Company Name:'
    line the engine writes into the Company Details block. Returns '' when
    the document had none, so the caller can leave the label as it was
    rather than show a made-up name."""
    for block in (by_subject or {}).get("Company Details", []) or []:
        for line in str(block).splitlines():
            if line.strip().lower().startswith("company name:"):
                return line.split(":", 1)[1].strip()
    return ""


@st.cache_data(show_spinner=False)
def customer_name_from_workbook(data: bytes) -> str:
    """Same lookup for a saved Excel export. Cached by file content, so the
    dropdown does not reopen every saved workbook on each page refresh."""
    by_subject, error = by_subject_from_workbook(data)
    return "" if error else customer_name_from_by_subject(by_subject)


def render_entry_ready_view(by_subject: dict[str, list[str]], sections: list[dict],
                            key_prefix: str) -> None:
    """Renders one parsed document as the finished Operational
    Information rows -- one copy-paste block per Subject.

    The block shown for each Subject is `"\\n\\n".join(blocks)`, which is
    character-for-character what write_workbook() puts in that row's
    Comments cell. That is the point of this view: the earlier version
    listed the destination manifest's individual fields, so a Subject
    that reads as one composed paragraph in the workbook came out here
    as a dozen one-word answers that somebody then had to reassemble by
    hand. Re-deriving the prose in the UI would have been the same work
    done twice in two places, free to drift apart; instead backend.py
    hands the mapped dict straight over and this function only joins and
    displays it.

    Read-only by design. This view exists so a person can copy
    accurately, so it must never be the place a value gets quietly
    changed.

    The field-level checks are still available underneath, collapsed --
    they are what flags an unresolved critical field, which the composed
    blocks cannot show on their own."""
    blocks_by_subject = {s: (by_subject.get(s) or []) for s in pe.SUBJECTS}
    filled = [s for s in pe.SUBJECTS if blocks_by_subject[s]]

    if not filled:
        st.warning(
            "Nothing mapped to any Subject for this document. That normally means "
            "the source PDF is not a Trip Support Service Preferences form, or its "
            "headings did not match. The Excel export will be empty too."
        )
        return

    st.caption(
        f"{len(filled)} of {len(pe.SUBJECTS)} subjects have content. Each block below "
        "is exactly what the Excel export puts in that row's Comments cell — use the "
        "copy button in the top-right corner of a block rather than selecting the text."
    )

    show_empty = st.checkbox(
        "Show subjects with nothing extracted", value=False,
        key=f"{key_prefix}_show_empty",
        help="Off by default so the list matches what you actually have to paste.",
    )
    subjects = list(pe.SUBJECTS) if show_empty else filled

    for subject in subjects:
        st.markdown(f"##### {subject}")
        comments = "\n\n".join(blocks_by_subject[subject])
        if comments:
            st.code(comments, language=None)
        else:
            st.caption(
                "Nothing was extracted for this subject — it is an empty row in the "
                "workbook too. Left blank rather than guessed."
            )

    st.divider()
    if not sections:
        # Loaded from the Excel export: it holds the finished text, not the
        # raw extraction the field checks work from.
        st.caption(
            "Field-by-field checks need the raw extraction, which an Excel file does "
            "not contain. They are available for a request parsed in this browser "
            "session."
        )
        return
    # A checkbox rather than an expander: render_field_level_checks()
    # puts every field's provenance in its own expander, and Streamlit
    # refuses to nest one expander inside another.
    if st.checkbox(
        "Show field-by-field checks", value=False, key=f"{key_prefix}_show_checks",
        help="Validation only — nothing in here needs to be copied.",
    ):
        st.caption(
            "The same data broken back down into individual fields, each checked "
            "against a destination's field specification. This is for spotting a "
            "value the engine could not normalize — particularly a critical one — "
            "before you paste the blocks above."
        )
        render_field_level_checks(sections, key_prefix)


def render_field_level_checks(sections: list[dict], key_prefix: str) -> None:
    """Field-level validation of one parsed document against a
    destination manifest. Everything it shows is either a normalized
    value or a flag saying why a value could not be normalized -- and
    the original extracted text is always one click away underneath, so
    a reviewer can check the engine's reading without reopening the PDF.

    Imported lazily so app.py still starts if entry_ready.py /
    field_manifest.py / the manifests folder haven't been deployed --
    only this section breaks, rather than the whole app."""
    try:
        import entry_ready as er
        import field_manifest as fm
    except ImportError as e:
        st.error(
            "This view needs `entry_ready.py`, `field_manifest.py` and the "
            f"`manifests/` folder deployed alongside app.py. Missing: {e.name}"
        )
        return

    try:
        target_ids = fm.available_manifests()
    except Exception as e:
        st.error(f"Could not read the manifests folder: {e}")
        return

    verified, unverified, placeholder = {}, {}, {}
    for tid in target_ids:
        try:
            m = fm.load_manifest(tid)
        except fm.ManifestError as e:
            st.warning(f"Skipping an unreadable manifest: {e}")
            continue
        if not m.is_usable():
            placeholder[f"{m.target_name}  (not yet specified)"] = tid
        elif m.target_verified:
            verified[m.target_name] = tid
        else:
            unverified[m.target_name] = tid

    # Order matters: whatever lands first becomes the default selection.
    # Verified destinations first, then usable-but-unverified ones (a mock
    # or a provisional spec), then empty placeholders. Without the
    # verified/unverified split a mock manifest dropped into the folder
    # could silently become the default destination, which is exactly the
    # kind of thing nobody notices until a record is typed from it.
    # Placeholders stay listed -- their absence is information -- but they
    # can never be the default.
    labels = {**verified, **unverified, **placeholder}
    if not labels:
        st.error("No usable destination manifests could be loaded.")
        return

    choice = st.selectbox("Destination", list(labels), key=f"{key_prefix}_target")
    manifest = fm.load_manifest(labels[choice])

    if not manifest.is_usable():
        st.warning(
            f"**{manifest.target_name}** has no field specification yet, so there is "
            "nothing to map against. This is deliberate -- the layout has not been "
            "invented in its absence."
        )
        if manifest.required_before_use:
            st.caption("Needed before this destination can be used:")
            for step in manifest.required_before_use:
                st.caption(f"• {step}")
        return

    records = er.build_entry_ready_records(sections, manifest)
    summary = er.summarize(records)

    st.caption(
        f"{manifest.target_name} · manifest v{manifest.manifest_version} · "
        f"{'verified' if manifest.target_verified else 'UNVERIFIED'} against "
        f"{manifest.source_form or 'an unspecified source form'}"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Ready to enter", f"{summary.ready}/{summary.total}")
    m2.metric("Needs review", summary.needs_review)
    m3.metric("Missing required", summary.missing_required)
    m4.metric("Provisional", summary.provisional)

    if summary.blocked_on_critical:
        st.error(
            f"{summary.blocked_on_critical} critical field(s) are unresolved. "
            "Critical fields are operational and regulatory identifiers -- resolve "
            "each one against the source document before entering this record."
        )
    elif summary.needs_review:
        st.warning(f"{summary.needs_review} field(s) need a human decision before entry.")
    else:
        st.success("Every mapped field normalized cleanly. Check it against the source anyway.")

    with st.expander("Completion by criticality"):
        for level, counts in summary.by_criticality.items():
            total = counts["total"]
            pct = (counts["ready"] / total * 100) if total else 0
            st.write(
                f"**{level.replace('_', ' ').title()}** — {counts['ready']}/{total} ready "
                f"({pct:.0f}%) · {counts['needs_review']} need review · "
                f"{counts['missing_required']} missing required"
            )
        st.caption(
            "Reported separately on purpose. A blended percentage hides the case that "
            "matters: an ICAO code or customs declaration being wrong is not offset by "
            "hotel preferences being right."
        )

    show_all = st.checkbox(
        "Show fields with no value", value=False, key=f"{key_prefix}_show_all",
        help="Off by default so the list matches what you actually have to type.",
    )

    status_icon = {
        er.OK: "✅", er.AMBIGUOUS: "⚠️", er.INVALID: "⛔",
        er.MISSING: "○", er.NOT_APPLICABLE: "—",
    }

    for section in manifest.sections:
        section_records = [r for r in records if r.section_id == section.section_id]
        if not section_records:
            continue
        visible = section_records if show_all else [
            r for r in section_records if r.status != er.MISSING or r.required
        ]
        if not visible:
            continue

        st.markdown(f"##### {section.label}")
        if section.field_order_is_fixed:
            st.caption("Field order below is fixed by the destination and must be preserved.")

        for rec in visible:
            icon = status_icon.get(rec.status, "")
            flags = []
            if rec.required:
                flags.append("required")
            if rec.criticality == "critical":
                flags.append("critical")
            if rec.is_provisional:
                flags.append("provisional mapping")
            flag_text = f"  ·  _{', '.join(flags)}_" if flags else ""

            st.markdown(f"**{icon} {rec.target_label}**{flag_text}")

            if rec.status == er.OK:
                # st.code gives a copy-to-clipboard button for free, which is
                # the whole point of this view -- one click per field, no
                # selecting text out of a paragraph.
                st.code(rec.value, language=None)
            elif rec.status == er.AMBIGUOUS:
                st.warning(f"Needs a decision: {rec.reason}")
                st.code(rec.value, language=None)
            elif rec.status == er.INVALID:
                st.error(f"Will not normalize: {rec.reason}")
                if rec.original_value:
                    st.code(rec.original_value, language=None)
            elif rec.status == er.NOT_APPLICABLE:
                st.caption(rec.reason or "Not applicable.")
            else:
                st.caption(
                    ("⚠️ Required, and not answered on the source form. "
                     if rec.required else "Not answered on the source form. ")
                    + "Left blank rather than guessed."
                )

            src = " › ".join(p for p in (rec.source_section, rec.source_label) if p)
            page = f" · page {rec.source_page}" if rec.source_page else ""
            with st.expander(f"Source: {src or 'unmapped'}{page}", expanded=False):
                if rec.original_value:
                    st.text(rec.original_value)
                else:
                    st.caption("Nothing was extracted for this field.")
                if rec.is_provisional:
                    st.caption(
                        "This mapping is provisional — its destination label or its "
                        "permitted values were inferred from sample data, not taken "
                        "from a specification. Confirm before relying on it."
                    )

    if manifest.present_in_source_but_not_extracted:
        with st.expander("On the source form but deliberately not extracted"):
            st.caption(
                "These are still manual. They are not bugs — each is excluded by the "
                "engine's field whitelist, which is the control that stops unapproved "
                "fields flowing through automatically."
            )
            for item in manifest.present_in_source_but_not_extracted:
                st.write(f"**{item.get('source_label')}** ({item.get('source_section')}) — {item.get('reason')}")

    if manifest.not_covered_by_source_form:
        with st.expander("Destination sections this form does not cover"):
            for name in manifest.not_covered_by_source_form:
                st.write(f"• {name}")
            st.caption("Left blank in the workbook rather than guessed at.")


def render_report_readonly(row) -> None:
    """Read-only rendering of one submitted report -- used for locked
    (past-cutoff) reports, which must remain viewable even though
    they're no longer editable. Mirrors the editable form's container/
    icon/label layout so the visual design stays consistent, just with
    plain bulleted text instead of editable rows."""
    for key, label in backend.WEEKLY_REPORT_CATEGORIES:
        icon = WEEKLY_REPORT_CATEGORY_ICONS.get(key, "•")
        text = (row[key] if row and row[key] else "").strip()
        bullets = [b.strip() for b in text.split("\n") if b.strip()]
        with st.container(border=True):
            st.markdown(f"##### {icon} {label}")
            if bullets:
                for b in bullets:
                    st.markdown(f"- {b}")
            else:
                st.caption("Nothing reported.")

# The salesperson list itself lives in the database (backend.py's
# salespeople table) so it can be edited from the UI -- see the
# "Manage salesperson list" expander inside Trip Preference Parsing.
# backend.DEFAULT_SALESPEOPLE is only the one-time seed for a brand-new
# database.

PAYMENT_METHODS = ["Credit Card", "Ok to Invoice"]
MAX_ACCT_NBR_LEN = 20
ALLOWED_UPLOAD_TYPES = ["pdf", "doc", "docx", "png", "jpg", "jpeg"]

TASKS = ["Trip Preference Parsing", "Flight Planning Parsing", "Weekly Report"]
TASK_ICONS = {
    "Trip Preference Parsing": "\U0001F4CB",  # clipboard
    "Flight Planning Parsing": "\U0001F6EB",  # departure
    "Weekly Report": "\U0001F4CA",  # bar chart
    "Document Repository": "\U0001F4C1",  # folder
}
TASK_DESCRIPTIONS = {
    "Trip Preference Parsing": "Parse a completed Trip Support Service Preferences "
                                "onboarding form into the Operational Information spreadsheet.",
    "Flight Planning Parsing": "Parse a completed Flight Planning Preferences form "
                                "into the Flight Plan Template.",
    "Weekly Report": "Submit your weekly update, or compile everyone's "
                     "submissions into one team report.",
    "Document Repository": "Browse, search, and download documents shared with your team.",
}

# Document Repository is NOT in TASKS -- it is an admin-only tile added
# conditionally at render time (see TILES, computed just before the
# landing screen below), never a static task everyone sees.

# Tasks that are visible on the landing screen but not selectable yet --
# rendered grayed-out with a diagonal "on hold" banner instead of a
# working card. Remove a task from this set once it's ready to build.
DISABLED_TASKS = set()
DISABLED_BANNER_TEXT = "Getting It Done"

# Jeppesen ForeFlight brand: dark teal-to-navy gradient + logo, matching
# the splash-screen colors provided. Base64-embedded so no separate
# image asset needs to be uploaded alongside app.py. Also used as the
# browser tab favicon below, in place of the default Streamlit icon.
JEPPESEN_LOGO_B64 = "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"


def _decode_favicon(b64_png: str):
    """Decodes the same Jeppesen ForeFlight logo used in the page header
    into a PIL Image so it can also be the browser tab's favicon, in
    place of Streamlit's default icon."""
    import base64
    import io

    from PIL import Image

    return Image.open(io.BytesIO(base64.b64decode(b64_png)))


# The full header logo (JEPPESEN_LOGO_B64) is white-on-transparent, made
# for the dark header background -- invisible as a favicon, since browser
# tab bars are light. This is just the compass/airplane mark cropped out
# of that same logo and recolored dark navy, matching how ForeFlight's
# own site favicon looks (dark mark, visible on a light tab bar).
FAVICON_ICON_B64 = "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"

FAVICON_IMAGE = _decode_favicon(FAVICON_ICON_B64)

st.set_page_config(page_title="ITP Sales Hub", page_icon=FAVICON_IMAGE, layout="wide")

st.markdown(
    f"""
    <style>
    .stApp {{
        background: radial-gradient(circle at 12% 8%, #3a7686 0%, #234b5c 28%, #16303f 55%, #051628 100%);
        background-attachment: fixed;
    }}
    /* "wide" removes Streamlit's narrow ~730px centered column; this
    caps it at a still-generous width instead of stretching content
    edge-to-edge on large monitors. Top padding is left at Streamlit's
    own default (not overridden) so content clears the fixed toolbar
    (Share/star/GitHub icons) instead of being cut off underneath it. */
    [data-testid="stAppViewContainer"] .block-container {{
        max-width: 1200px;
        padding-left: 3rem;
        padding-right: 3rem;
        margin: 0 auto;
    }}
    h1, h2, h3, h4, h5, h6, p, label, span, div {{
        color: #eaf3f6;
    }}
    [data-testid="stCaptionContainer"] {{
        color: #a9c4cd !important;
    }}
    .hero-header {{
        display: flex;
        align-items: center;
        gap: 1.1rem;
        margin: 0.4rem 0 1.6rem 0;
        padding-bottom: 1.1rem;
        border-bottom: 1px solid rgba(255,255,255,0.15);
    }}
    .hero-logo img {{
        width: 108px;
        display: block;
    }}
    .hero-text {{
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .hero-title {{
        text-align: left;
        font-size: 1.7rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        color: #ffffff;
        margin-bottom: 0.15rem;
        line-height: 1.2;
    }}
    .hero-subtitle {{
        text-align: left;
        color: #a9c4cd;
        font-size: 0.95rem;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.18) !important;
        border-radius: 14px !important;
        backdrop-filter: blur(6px);
        transition: transform 0.15s ease, border-color 0.15s ease;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {{
        border-color: rgba(255,255,255,0.45) !important;
        transform: translateY(-2px);
    }}
    .stButton > button {{
        background: linear-gradient(135deg, #2f7688, #1b4a5c);
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.25);
        border-radius: 8px;
        font-weight: 600;
    }}
    .stButton > button:hover {{
        border-color: #ffffff;
        background: linear-gradient(135deg, #368399, #204f63);
        color: #ffffff !important;
    }}
    .stButton > button:disabled {{
        background: rgba(255,255,255,0.08) !important;
        color: rgba(255,255,255,0.45) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
    }}
    [data-testid="stExpander"] {{
        background: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.15);
        border-radius: 10px;
    }}
    input, textarea, [data-baseweb="select"] > div {{
        background-color: rgba(255,255,255,0.92) !important;
        border-radius: 6px !important;
    }}
    /* Dark text on the LIGHT input backgrounds set above. Applied to
       descendants, not just the direct child, because the broad
       `div {{ color: #eaf3f6 }}` rule earlier in this stylesheet paints
       every div near-white -- and a select's chosen value sits in a div
       nested several levels inside the component, so it rendered white
       on white. How deeply Streamlit nests that value changes between
       Streamlit versions, so this must not depend on it.

       Deliberately scoped to the CLOSED control only. The open dropdown
       is rendered by React outside this element, onto the dark panel
       from .streamlit/config.toml (secondaryBackgroundColor #123244),
       where the theme's own light text is already correct. Forcing dark
       text there makes the options invisible. */
    input, textarea,
    [data-baseweb="select"] > div,
    [data-baseweb="select"] > div div,
    [data-baseweb="select"] > div span {{
        color: #0b1c26 !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background: rgba(255,255,255,0.07) !important;
        border: 1.5px dashed rgba(255,255,255,0.3) !important;
        border-radius: 10px !important;
    }}
    </style>
    <div class="hero-header">
        <div class="hero-logo">
            <img src="data:image/png;base64,{JEPPESEN_LOGO_B64}" alt="Jeppesen ForeFlight logo">
        </div>
        <div class="hero-text">
            <div class="hero-title">ITP Sales Hub</div>
            <div class="hero-subtitle">Trip support &amp; flight planning parsing, in one place</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Authentication -- gates EVERYTHING below (including the task picker)
# behind an Okta login, so nobody outside the company can view customer
# PII or internal sales data just by visiting the URL.
#
# Fails CLOSED: if the [auth] secrets aren't fully configured, the app
# refuses to show anything (rather than quietly falling back to open
# access) until an administrator finishes setup. See DEPLOYMENT.md /
# the IT request doc for exactly what needs to go in secrets.toml.
#
# Defense in depth: Okta's own app-assignment scope is the primary
# control over who can even attempt to log in, but this also checks the
# authenticated account's email domain independently, in case that
# assignment is ever misconfigured.
# ---------------------------------------------------------------------------
ALLOWED_EMAIL_DOMAINS = ("jepp.com",)


def _auth_is_configured() -> bool:
    try:
        auth = st.secrets.get("auth")
    except Exception:
        return False
    if not auth:
        return False
    required = ("client_id", "client_secret", "server_metadata_url", "cookie_secret", "redirect_uri")
    return all(auth.get(key) for key in required)


def _interim_password() -> str | None:
    """The temporary shared-password stopgap -- used only until the real
    Okta `[auth]` secrets above are configured, at which point Okta takes
    over automatically and this is never consulted again. Weaker than
    Okta on purpose-for-now: one shared secret, no per-person identity,
    no revocation -- just enough to keep the public out while IT
    provisions the real SSO app."""
    try:
        pw = st.secrets.get("APP_PASSWORD")
    except Exception:
        return None
    return pw or None


if not _RUNNING_UNDER_TEST:
    if _auth_is_configured():
        # -- Real path: Okta SSO via Streamlit's native OIDC login. -----------
        if not st.user.is_logged_in:
            st.subheader("Sign in required")
            st.caption(
                "This app contains customer and internal sales data -- sign in with your "
                "JFF account to continue."
            )
            st.button("Log in with Okta", on_click=st.login, type="primary")
            st.stop()

        _user_email = (st.user.get("email") or "").strip().lower()
        if not any(_user_email.endswith(f"@{d}") for d in ALLOWED_EMAIL_DOMAINS):
            st.error(f"Access denied -- {st.user.get('email') or 'this account'} is not authorized for this app.")
            st.button("Log out", on_click=st.logout)
            st.stop()

        col_user, col_logout = st.columns([6, 1])
        with col_user:
            st.caption(f"Signed in as {st.user.get('name') or st.user.get('email')}")
        with col_logout:
            st.button("Log out", on_click=st.logout, key="top_logout_btn")

    else:
        _stopgap_password = _interim_password()
        if not _stopgap_password:
            st.error(
                "Login isn't configured for this app yet. An administrator needs to add "
                "the `[auth]` block (client_id, client_secret, server_metadata_url, "
                "cookie_secret, redirect_uri) under App settings -> Secrets before this "
                "app can be used."
            )
            st.stop()

        # -- Temporary stopgap: a single shared password, used only while
        # Okta is being provisioned. Automatically stops being consulted
        # the moment a valid `[auth]` block is added above -- no code
        # change needed to retire it.
        if not st.session_state.get("interim_auth_ok"):
            st.subheader("Sign in required")
            st.warning(
                "Okta SSO isn't set up for this app yet, so it's temporarily behind a "
                "shared password instead. This is a stopgap only -- everyone uses the "
                "same password, there's no per-person identity or audit trail, and it "
                "will be replaced by Okta login automatically once that's configured. "
                "Ask your team lead if you don't have the password."
            )
            _attempts = st.session_state.get("interim_auth_attempts", 0)
            if _attempts >= 5:
                st.error("Too many incorrect attempts this session. Refresh the page to try again.")
                st.stop()
            _entered = st.text_input("Password", type="password", key="interim_auth_input")
            if st.button("Enter", type="primary"):
                if hmac.compare_digest(_entered, _stopgap_password):
                    st.session_state["interim_auth_ok"] = True
                    st.session_state["interim_auth_attempts"] = 0
                    st.rerun()
                else:
                    st.session_state["interim_auth_attempts"] = _attempts + 1
                    st.error("Incorrect password.")
            st.stop()

        st.caption("Temporary shared-password access (Okta SSO not yet configured)")


# ---------------------------------------------------------------------------
# Document Repository tile visibility -- see the module-level import
# try/except above for _DOCUMENTS_REPO_AVAILABLE.
# ---------------------------------------------------------------------------
def _document_repository_actor_email() -> str | None:
    """The signed-in user's lowercased email, but ONLY when real Okta
    sign-in is confirmed -- never the shared-password stopgap, and
    never under the test harness unless a test has explicitly set
    st.session_state["_test_actor_email"] first. (AppTest has no
    supported way to populate a real st.user OIDC identity -- see
    _RUNNING_UNDER_TEST's own docstring above -- so without that
    explicit test-only override, this returns None under test, same
    as "not signed in".) Returns None in every other case.
    """
    if _RUNNING_UNDER_TEST:
        test_email = st.session_state.get("_test_actor_email")
        return (test_email or "").strip().lower() or None
    if not _auth_is_configured():
        return None  # shared-password stopgap, or auth not set up at all
    try:
        if not st.user.is_logged_in:
            return None
    except Exception:
        return None
    return (st.user.get("email") or "").strip().lower() or None


# How long a tile-visibility result is trusted before being
# recomputed -- see _can_see_document_repository() below.
_DOCUMENT_REPO_VISIBILITY_TTL_SECONDS = 300


def _cache_clock() -> float:
    """The monotonic clock _can_see_document_repository()'s TTL cache
    is measured against -- real time.monotonic(), except under the
    test harness, where a test can override it via
    st.session_state["_test_clock"] (same test-only-override pattern
    as _test_actor_email elsewhere in this file). Patching the bare
    global time.monotonic directly isn't safe here: AppTest's own
    internals (e.g. its run(timeout=...) handling) read that same
    global clock, so a global patch's side_effect sequence gets
    consumed by Streamlit's internals too, not just by this function.
    This override is scoped to exactly one call site instead.
    """
    if _RUNNING_UNDER_TEST:
        override = st.session_state.get("_test_clock")
        if override is not None:
            return override
    return time.monotonic()


def _can_see_document_repository(actor_email: str | None) -> bool:
    """Cached wrapper around _compute_can_see_document_repository()
    below -- this runs on every rerun for every user, so without
    caching, every non-admin visitor pays for a fresh
    documents_repo.is_admin() Turso round trip (two queries) just to
    render the landing screen's existing tiles.

    The result is cached in st.session_state (per-session, never
    shared across users/visitors) keyed to actor_email, and reused
    for up to _DOCUMENT_REPO_VISIBILITY_TTL_SECONDS -- a False result
    is cached too, since an ordinary user reruns this just as often as
    an admin does. Recomputed immediately if actor_email itself
    changes (e.g. a different person signs in in the same browser
    session) rather than waiting out a stale TTL for the wrong email.
    """
    if not _DOCUMENTS_REPO_AVAILABLE or not actor_email:
        return False

    cached = st.session_state.get("_doc_repo_tile_visibility_cache")
    now = _cache_clock()
    if (
        cached is not None
        and cached["email"] == actor_email
        and (now - cached["checked_at"]) < _DOCUMENT_REPO_VISIBILITY_TTL_SECONDS
    ):
        return cached["visible"]

    visible = _compute_can_see_document_repository(actor_email)
    st.session_state["_doc_repo_tile_visibility_cache"] = {
        "email": actor_email, "visible": visible, "checked_at": now,
    }
    return visible


def _compute_can_see_document_repository(actor_email: str) -> bool:
    """The actual (uncached) check -- see _can_see_document_repository()
    above for the per-session TTL cache wrapped around this. True only
    if the Document Repository imported successfully and EITHER:
      (a) actor_email is an admin -- documents_repo.is_admin(), which
          checks the DOC_REPO_ADMIN_EMAILS secret or an active
          doc_users admin row. The only path actually used today.
      (b) the DOC_REPO_ENABLED secret is exactly "true" AND actor_email
          has an active doc_users row or at least one category access
          grant. Implemented for completeness; DOC_REPO_ENABLED is
          unset by default, so this path is inert until deliberately
          turned on.
    Never raises -- any failure here just hides the tile.
    """
    if not _DOCUMENTS_REPO_AVAILABLE:
        return False
    try:
        if documents_repo.is_admin(actor_email, secrets_getter=lambda: st.secrets):
            return True
        if (st.secrets.get("DOC_REPO_ENABLED") or "").strip().lower() != "true":
            return False
        conn = documents_repo.get_conn()
        try:
            user_row = conn.execute(
                "SELECT is_active FROM doc_users WHERE email = ?", (actor_email,)
            ).fetchone()
            if user_row and user_row[0]:
                return True
            access_row = conn.execute(
                "SELECT 1 FROM doc_category_access WHERE subject = ? LIMIT 1", (actor_email,)
            ).fetchone()
            return access_row is not None
        finally:
            conn.close()
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Login splash -- plays once per browser session, right after sign-in.
#
# Streamlit does not run scripts placed in a page, so this goes in through a
# zero-height component whose only job is to hand a <script> to the parent
# page. The animation then runs in the parent page itself, not inside the
# component, which matters: Streamlit removes a component from the page on
# the next rerun (any click), and an animation living inside it would die
# mid-way and leave the splash stuck on screen. The splash removes itself
# when it finishes, when clicked, and after 10 seconds no matter what.
#
# The artwork is embedded below as a compressed WebP (about 78 KB), so there
# is no extra file to deploy. It is skipped under the pytest harness.
# ---------------------------------------------------------------------------
_SPLASH_IMAGE_B64 = """
UklGRgQvAQBXRUJQVlA4IPguAQBwBQSdASqwBCADPkkij0UioiOhJRL6kHAJCU3fhJjSRTKxOrgDDY3LtrRdZGDVH7jIfa5f+Y/v
f+g/43+B//HnbbL7X/bf8h/l/7f/8v9f8/Nf/rv96/yP+J/tv/1/ynzE/6f+X85uxf/D94fwseSfq/+v/vn+o/+P+N/////+5n+2
/6X+e91H6d/5f+e/eH6A/41/S/9D/e/85/1/8Z/////9Xv/g/13uu/dj/z/5j/S/AP+s/3b/s/5r/Vf/j5f/9Z/4P9H+9/yz/un+
l/5v+D/0n/4+gH+p/4b/0ft5///jU/+Xub/63/sf/j3Cv5//jv+1+f/zB/879wP99///pI/rH+6/+/+x/3H///+H2Pfz7/Ef+r/L
f67///+r6AP/f7Tf8A/73///9nuAf+D2Z/Ev8//iP8L+3nvq+L/uP+I/vH+Z/XfySfdv6L/Ff5f/of4LoReu/6fmh/H/vl+t/v3+
d/9P+I/ej5/8BfnF9NewL+M/z3/X/379y/7/+9XJsXQ/8nqEd7P+R/hfyT+E36T/ef6P1b/fP8x/vP8V+WH2A/rP/tf77+U31X/w
PDR/N/9T9nPgC/mP9h/7X+J93X+6/+/+k/4H7q+8X9G/0v/y/1fwMfz7+6f97/G+3V///eV+4H///7Hws/u9////iVbCQ4bgHztb
IWAeYd+ezchrQ7l/+qFbDXf/G/PZth0Qd0ywgC7cp7NsOjMrT/3E3Y91cRRQ5Nldy4ttVW8Mfrc9cFI789m2Gu9VC3n4eI4wolIz
Dym6uEY2qLNsOHoemUG57AtkXaZJ5lKB0moCachG//jfysa7/Yo2unDRSwUxSWu1FbL1eVYpWGexEx5AH/f3xJNHAkg0xpvrVDhA
FbDXf+4H5/XE4mFEoW9HdgtUcTqCa6aV5L8928QXsrUy8Et4bsi/i+ZDnwLk35DWAYiryhp/66RdYSSK43h3ICxkBBfueyOJzbXU
VLDyRQ566wbLPiBHb641jOpTtXcXe+JrWs2Lfhfz5ynvsRO3k9Wo8pywyrp8yI2PZH7yHHN21T9tWx8AM8Bj2xrc/fJPYN42a/Lu
kIWJx5qbk/dz15yT6E3WjlWKyARpF2PGdqYfyOc6Q9V1wfTfRqQnEntUpaIzD+5I3kgxS0MAo+MYjS4hLAXirPWnMzr1odXN2vZK
zG9en/jePzmnBExVz/S36zgZFPrURAND281KL9+TU6zs/hJ05vupKwf3sRsE76TEZxkh55r2fjDwP7+nMfHd0H9bt/SFQEA6mY2e
T116KfemVThW8UHu2g+gd+Vgq0oBpv6kRh2DtJvbjaq8UiRSEMeH3P0E5gnn+p+MZQNA677EaVB0J7jc/jfns2w4pcZnTKxtB7YH
tdIs2U0xP+UmDoX24Zne//r5j/Y6iFvpVdJaBrxmEabVlZmYKLCndkRuxZMfQpLxI8CbpoerNijpXrKxe8sQ9Qpmmf6/CnfcBW+j
qez8mivu2DrYURK98L4hpkRrt7G5xnGoIlLSZFCLUdDgNwxiW/r+uddT6nGfA/TyMMbNkr+fNp1qu1sMuDWiVgWUXbUbD4jAHxjj
jYh5EPbx8q5V0HYlRXFkTMnmH7AxPBLcKhhrpidW4+nrpLT0Y5mFd4NRSFJi5l5Rl/QAyFNIU4300CgP2zVNBhNFSPn1oUk3BLLh
vIC/4BFON0efhvEZ+g+qF2XUJiYGoj1hoUZ8Dl32Gtvbni9AACbYaBizDYQeOINfiinI55RvUHublk90NqEI7IVB0G/slI0SZyFw
0JRtz+ZcHJsgt9jdqENkafVA9cSsW9LRr4XhHfOqY8R43UTC9jIOut4d8qKBuns2EbDlwhJMzOFSw+ZvL+mjMq6MMQl1LuRsXjGx
NXf6iN++0Xa6jBKpgiLo2f/igbntBK6V5hqmG4+jzz8L5/ATLUzXg0JBrSiA/BgfpBDkxtT+fbOysgtqPO7ASExO3DSK2zWZ1/O8
orjWwYAklQrOTBsAtwllsMBqIFwBjxDBLw2Ak+XTUBvbuJ/llpiY/okV9Mkkir+rBf+26aEnttaW8ldkATaZNsZjlUkBnxEWiGZn
jeioYTTKuNDdZuPhVwXaLRIQ/jP/ckSItbv1L/TzG5opxCsJQy2e39nqe0NwnM2ynF8+WYO5Nq/P8GffWpuC+ZxX6ezKRiDV1AMD
/AE/7OG25m89D8EGiQRevWyqTCtFOENwlmYCe5tPKewSLAANv/iu7OeIYXjb26RbjZdLOA/LTYVVJMX0qJ2GLBVJ40Fw+8C0PlLp
Va37uFgvsjMFC2aMhcy/2a97v0Y8gZOnzwYdqzx/7alA84WHxscYchLwJdMIxMIDvUljRVfbwSa8Okn+PrskIWaMN7khOExNRQ8l
y9Ro0QrjrU4vyInrNBB6j8TqMZhIB8seXRRVjaiGK91ztj8cdlXPWV2Pqlk0JyMO3bQuswR7t6OTbuX5j/vjics1x/QpgSQAKkEC
uBDU2QdEwWcLGCELvTQoOJm9rRh6YPx/vaBDA6E68QWMVUBlzE7Y4/2j1v9CWf55aErwOKjj3FKjU61ZREjOFXOGmOjOrCLiahzk
WRdvvuNwdj/WK2d5HLErEiyrdmNDr+V5GER3xHaHXMhWrgliXKYdn6bM0dxGs2sXXi3IutpezDMyzFqr8/JtvteA4NwGq36+Cru6
mnq5ORH0JGTk3hFLoWkeH3oHXAB5yhjRfcz4tS2kT2WffQvlyETMkVghyxiR+5tx//0Y5D/uqzIPUvB84v8VTPgxytzS5KhcwCE8
myA7Q4AB/5KJ68fY+vBDcrqUV/z2seX5AFI/aA7X19k4v+sRmWW22LH/9PHUKOyTi+shnoUPMYxmsxl3zN9dnEQdCCuaVau3AFcY
UOPrVbkmSkVUwriA9Af5xwzX2NLOoSKwAiRvbwiNlgheabux9JtZCyVCJqcv57h518Ck3kuA2rrFhyNUsWpuQ/z8zaq5hhAW+5dT
ptyuoH+TxaUkyxf3QvWToFh0JDg/GLNlQBhxpokTx74dVynXC32Gvi1B6w92f0J6ZeLbURlSWoxhokfYnzkx8JhoFZrlAnpTJ6nT
PnVNVihd6NS0KgjSmnH3G9KDdQ0uJgclnny2lBDyhs4/B+idJfoB7bHEb3+dqyATnSQZmEJnjWnx0iqROg0ibuvdlEhyZrYW+2+t
Bd66O7Dz6RAwxlk7bm5nMX+mXGVBkjklenZhZi2dMiMeGr8hWvGyvXDBX8EeW0w5zGfxVocNGk7C9KD/O34uuLibSi4Tfbymb9Lc
6mLxjrhtqh6D8GImN0vrGNqI+VTZqy03Wwmp6CybH2wbSt+UNIuGZ2wC3tMu9JRcuDfCHn9a024ok5X/wr7VlRcDd2yMj9G8zXSD
hZ2w9DQaEsuL4HAvgaPt1RLRdm21y5mdqMPT5zvR3deLScaRzqUyR/7mpmiJOqkGUQ+Trpen50yLRiPq/xLHn10bpYeFwKMaTu7I
OlGGN2yNv3m/+oxHPBUEHdoaQ9vH3DXLmGsiGx4d7/m+OzuxppMZBlK/rwfzJ9L3GBv++3SWaBv7yUwAEi4Ie+k6i00bqgwUXIN9
Ak2Svl3o12ADCqovVpfa2nZQldamRaUT0JUhGZOPeBXanui+vLNGMfsqvNLgVy7bxWgKVx9VKx/zoQPLu06s8tHQ/6BXff259ZG4
Uh1Arhw/4eVn4lzsLPACWlYI4QlxuQyhLiI1zSuuLd4o1XKBP+JzR+38Su1ywJ7cSzipKg1+PI/gq+WhxUTUDIK/I5YMb2weKxa4
Qnpsw/NA0ocYZtaGxWTFHAnsb92MaTm6mpnKY8IGwqtq2Fq9PE1M/ugjSTXo8N1H47UD3veRC6EZ6oJIl0WqCNuwLkcqYzblMapA
oizLNqFoQUOmOvPBKd+K/IgSemKuvoHF578FTWmq0JFe4FEFLZD2e65SQNAZf6zBY6nU2nZ9Cos3vPoC5Brn2jWDlKgf3mh5ueTC
N9SgN6r5/Ut3HWfRjsCIxkoWdjgzmJ+Uq+PmA7X9Rmtk8/C2NZdzZC5D9h7MClzKz8RTxdRpgNq/aYiM+jI7P0ZgHXF/Bz9Kfl55
cLqPZ+js+bNsLNI0tPlNH//G2hInJ90E57/lluuNAb+hNK7eKnSKElN1oP8UHmWm3U45//0bUcun8lYfXRaNq4yD+jkvkpjrdXTb
Mu+g4kTtwyYRIiHxnuJHH9vxsB87WW/V+QZlXn4cgQR9LutrEK4anMHlUerrGtktytqEvOa2+8PBsbqMg5NSbHfRzegi2Fd1nAYt
b8g7QUfv7qu8BtjlVFmIoeTrR2l/VyRFEVQMRHWCrhze4lHp61rZsoEcLlfdvzuaqCEfQrYk//Ok07KqAbb9ET/LFFH/hqK0kS2s
hqHEgApvgUd9Jc5M4NW5cGFYnG9fBvoDVt/CpTcKxOzE7kilSoUn+ZqORqXqWKH8KVIEKli+RRt1jlyNKrXaapmL7ORLUSmVtx4w
bdsUpLwAupmwZ66ezO0CCZnb+/HfgD/ukcuZo3iwje2NyyUEZRyxmaHTqHRB2tZw9PiVcJtN074of2SeGjIhj2gxOXYBx7H4EwfT
Zvpbj1rG8Yu0RMSRDNG+80KH1f5fajfPnU8kgs7u/ec8DIAp7v4U89Ynr2XgOtWi3+NRoEbChsDqD3vBO6o6phxWvVOwmcHNP7sB
Ed3WAg5iSGSLmm6+s2PjhqNjN9cKtw4K1U9lFOuIknxeEhDx0qeHQgff3XxzlAJtOOnARsu0Zp2tac3KZWrW4RVlHghwiTKfEi16
f9ODfxBSqHXAcelvlwDxPY1DhJ3HGN9z2PcRE9hWFKqqjSWMrjgxjV7rjBpBzHnWpNo6c+lsKoD6ET5WJ8HbziRXjC2fChdjuZfQ
q46ljaYMjBgt8zZgOrLiiuAN4k7qcf1Y68XOcXGIYVQxS8DSfW5NBO2g7W3GgxB5WrL9gdLhyHzphlKIgTjvNM/fJqm2Sn4cf+Nj
duM7UX0vEFCTBqMoO0guABsrCJeGsi1adS88k5bTyFxosbz2w1ZkBGjvFJXCILj3dbo5AY/cJQ4g/78K2uCiqioTJ86n7TuEYVL1
b7hXn0hHqQIMyiEvF5xyiQKxdIJd5uU2UQ1tIzbXjs6+1cZV2dp72b3yaj6mWz4qF9dlLxFkqTUc12P/J5Gmhi/gVVIwvhIFBTPg
Z/DS99sGaJcVX/NbYkR/T61A54qWV5yA7NJ2bH9yF8Bb9/40oF3EBHvv/cqajx2Zj8TApfr/6EmfqPUL4tMOa2w/UDvl1jg4Tpfw
C2TtzzBCB4x/m6Q2WYaLD6egv3WQiqhXbtBna+jisAzx31aFMAYJP73Lip9dgf74A6P3OBDWffRAXzrgFOpvnC2//ta8QX/NHl2Z
bTduu8Zg7RBcgPiYLcJOYXdR6vWEwOXVpYom1Qt7rnvRgwj8Qfa3v5uIo3HhFYX/ycsLKVqf/17i+dT5a7Amf+1Z84wXw/vgc+Eu
8H94IBowBs5VHRYJ3fRthPYA5yQLXmp6ItsQdQ8q83sXM13ZHXB1omH/mnkMXS35bPP+mw0tSEalkENwUvQJ9ao1zMpUmfENmkmQ
GtZlbh5foTo/QfoLlDfhh5qxoR2D2mOww1QBPXpJxHAkKlGxlu5l3WIJcZ4sqtAfUppGD2B1RttPTchi8R+3JAheu2X26yiXcGwD
Nke/7w/22N/l5xmQFi4Uk/cCM12Trjy0wrGyO7HpKmRlIIKiSOpAx5CMdz4f3tMgqg7pye8XPb8fwquv2duHauRvsl0Qtz9YvvfB
5Ou+/6yrCGeZ9+FRHnx8GPT19Vo7aLPu9JpnU92roc0DXtRWoMdqT6LsGRZ4cMGprzm42kwUe7IyUmoH0nbyu3Kb98Y3vCIb/Jco
XhV//QDO9pP5KctQ6bNFWH2v8U5ENt61sgKiz0K6IQ6rCb+9k0ARl6wfdF7/awm1/EhlpUC9mxS7Wpw0UKczFaRUwLrVzdRuMHeu
mElYQgOMi4Zdo8nKgcohqzgJH+1oXfAw5h9eT8y9kmKcUIv4/QQtU/yM7Rc7l1s53eG9PFcd+z00u7XiEiBxAXluN9KWUjkFcxP8
3fXa7AHZsSAp7sa/+iYC+kUTS9n9Sq1AFnZNtxgndyE4Ho1wMVAYuJXYa5oDwc9VjVWPh8s6tzsCaAvBZ5EXsAj9NjX+fDceaiOU
R6ULCUmQlq4k+NBnatAII4SMYOUKexdywwj4bGuFnCj04GS1aI4YDT9q21E3eCAuygV1A7nqgZWWXFixyof/vS3vSAIID+zuJF5P
ZYJPo9shwgCthovK6BpWyRxVmC9ztAUzW5kmi86RFRM0B8X7hDUCARnL6xgpIbbatOt4ZC8IgMlX3oaSY2qiXEBe0pXFihaxzNUN
fCbrxE7e6OVaxNrRHJ+4ZHhf7x16KimizRO4cC/nS6XP4XWNKJPCv75mRUjSYbrfUopaFdNymBj1VKMhk325YX3P8rns2w14NVSB
a1uADWxdeNB9pXySM2Crfr/P0oiAQzX8Lqj/wZ1YIWbyMSepRiCoOKmDINgw9w7uGYhj4dEuLvzzBGMzWdJlX/QDdjhRvWCMqIgk
YREAcV9cnJQFar7/Kne2C6rYIYzTY3A4hvCNKD0tam9PXNWgd9HFOepAuk1FZi7l1Rz/3qcw07C3KyTGzTVOSAwiXQFXQYDhvqwB
6H0WXgGJ/Q5TgcF5JpeEMhYpP40ou7P9b/3fZAulEcXkCHrZzynjiM1pM2Id1nVCVSdcwGYGkH66ytwMGuMZdo+C5gVAB7/jfoEU
+xH1rK0HkdVpfLfZhn0BvTB8dmzmAPUQnn9E9LgpPvFOS8XLHfJ8Jsu+g2/rv9OnQaR3bEvk3iaUmuXEB65m4KjGZwrGi0JdryYv
SamzbxDlNuwLBX1T9UgJHAsGPKKrLAa3+nFD8QVAPALIbchTqIl5351KPBa/5dBMmfRkDYHknE6/IyC/s1BKYQcjm7uRD+AOTS3A
EtyFgMh1zYEbdy1CvOaMhN4wOv6ZBzKyz9uNmwIVvoqZ1HpQNb8CvYgzRkH1Q1penXhyGKaTx0tZY3NGUw/YV9ZQsUl78G7XLOOX
dtNS1uHvFkZSohPTBwlZd0EtZ3144PzWgN52xLAKBlRd4AWPX75rR9jAjcA2cCKUL8k/zfoXeMgp1Fn1jTlGjUPEEndF9ddw3vY3
xfVfQTcoHpe7TWA9R0sRD8YDduCKxe5uEKPxU6rTV3pdT1q7o2Kp0iRbJqXJfspEvQAWm1NjxAQiznTBpYBfS7/Y/x3oerzoDQ36
rbz4tGrRQuf4lJQYi709dgCHI3nRF1AvDOiXukLoGZ+dJR0zmPlGt62gMP6VXiq9imMhEhaiIFljygcZR99iLXzFhTK74jK4E13b
pK/4GNanUu69PtS+o4SEtzx8ZSOZh75qUJa62/cGR4Ys3NdnbjZxyes6XnSiJPcAAUVanAGf4wiksUPCONP7dVgLZH9BUlPJOeXM
qbu3HAxoxIP/oMsF8K0Mfd80vuPhwR0Jztcj+2tUjjYMnBAaQvm8L9Q0hMDncy02Ll7DQ7oStP03Xc1e5F9zd42yAVU8DNis+mqf
0Zg0yxYsW57HB1yKWvr8tbd//t50nmgME1Ht3sXOxocA6YHbLuQKTk4g4mjZ5RCVxYIyYEoXq+7fS7Kh6561zeRNCWzs3TmTrSGa
BaEfmlvSVN/Msy9XsYDwmYgWOHv09EEST+cWyXonMj6Q3Jqpgw8ePH2IqP9KbgO1fZahoxucBOIB3bRpAsTKjKeeJ5fwkI6ygLfV
uCbFJDzPn8mv//LUR3T07BpMGYgJnrKs95n0N8/Ycnwlq09ZYL31ZlIZtV77hhhZ//8oyu1ddvY+bTBKQbJv5wIhV8PXsn1/hv3B
C53Nz98aU1YMxZkVTOa1C0gOnrQ4uZZzXwhuHRR07ZnMFuAQAAmeNbBUJvCugAU27t5zt4SBRU66rf1HUVdrHygBXh5nY/HY7RAu
f/MDcTBEGi7MCzmJEuGds7h3gCFen8Rjpdikq72CPO3GdboolVNWxPkOKWVjHljjkS6cu/dN7G8s7Vb9Ck8KENlT/oMhChXoZ7cR
4FQiBYXoqHm03o3PrWObgYic/5QJrzWpr62WIjLGYjhBQZ4V47VVm/MqJ6F9fUNar+5xlnKNKNymcLPVZJm1PDDwbZfsWUJivFsR
/6A1IxeIcD66RF9PoJY3z02RHL84CE9Ni6HgeEqzEARJQaAgAScYl3hdxXv4DeKphCZ75qfS6V2ATNRv8IWXlTD9WvuXVu7M/K1L
sJTHqeLwiT2g1+aKuJH+B4IZveD6FqXmmjTip5/O4Y+Kfz9a7YxVwf5Qq9dLfGWD4dfmcu6siSLbaJT3cc4N+m5I52a4ATZaPe1l
36R9eNTEP4Q0H7MJOTbTGBfTga7sl7QGajjNGtSn41Sgpu03y6C1RnaL+f58Lr709gt+ugBOOZfK9bO8s6GFtFHFP4RG5S9fJ14J
e7EgXbDymC7H6w9iQMdY5n6MQZOYRFhnrKn0IMsy8KhVS1nlEm1LL/gSFkCto3CikwVqAuvdJDncFhPXl2PjGNqbT+qAfd9XPa44
GUbFB4OlhXYeRdmn44ZnQUGXYlg7+9jpo8XObVPLBD1yMOmHa3zITWRPNjuy9slFl/rG53lCs9JjlaTgjvDVYuVTsMuRQlMI9OLK
jpWiiBED1I6zeKcXX2TRstj7VhZ3wGeZ0UfO3ZML2hPNV9d8BvorcsDtSLvgEGFDcnr59rWbnEGm6RzCgpYSmJc8N/rjybYozkWm
6Ldc6BeXXjiEMRvoP4elB9gOd5m1zhGsUNsjP/CjdKKRcNrKOq+At+qQ3SkiJSydEJAEY8yFZqZALZrPZwFMe6GdwF3ah7RB7dza
L2tvd0xY//tcpl/U08rdx9H6D5TLohM5ZJ1/NY8QTsz9UNYjetPiKECiykN3w3bQ35CYQhjeBPDFqbDxdvMIIcXwrGv0oaBAGcUR
u+t9oXNao6fRSQuq88qWo5fdlI0anKySGA4Mn1PtlNxE7Fse9WxCtUbe+Zl0D+z8WkW/OgCty7YHsJYCpBQ3gFhers62Us/a9s3Q
hkFNn83GJoPb5dyzMIHkxcRoMKHltKzLLazsHrRFqkx6FNYiv3t/ppUl7O52CWvPXdG+Ew2COMTfYfF87hyinlm00ZXHHey+9f+m
aTkgJ67/5CQxtidBGrFff3QsG0/dpqieanccjvzqrB16/EzJXEsjKEKoHWCeMu0olnL/tiGKhWe4NJR5J0v5NHkaHS9ZPrJKU+ZQ
RNyP37o6OEjq8GxfphHu5V6RYGyI2rdtYz+HFCnEB+ThIFD931ffwXQ2oXQxxgGnAp4erMMx7s1gA3c7Djyo6kghv2yJsCKGGjh5
oGflOx+38SoQu8orrkTBzGE3FH4M2j7SyRDlDQsEbbecy1l5/rQ6LV0LbcWn0U5VTtdLz2XpjoavegBxDwORGabEPWIc4bn4hv22
DR50scleElNfmOuLB2IbeCtBwiDBewYnunYku1UXf09AXyTHSpRR6Oc5qAKNBen6MqZNvzv8JX+cuPU4P3hluD38DC8YpaNY58Ju
zmnbmVqM1lP9pJkPay0ELJAl80F8nZxSbdrMLLxEakFM1ihl7duUV+xTihMc2kmjeJBF8FF5o55dTqEFhNuXCPnSvBoOc3HMRHMP
PdBRlDg25/etJoBEw4kdE4OWGzlot2JB31S7i2miDXdeLkpbuIzBdOEqJYH0/MN1U0w4dPgXuxTxvpccbVqVDWea/4PuezruPo+4
CSee6GT0lFkGkZ6al2woXBKY2q/KqxZeiGfAttk8ySqAGEmhvZlZDH7IhKi8j0IO7uPJ4t3zPGYoWjPIzvWEHsYLMc1tTWUmCv7h
NxlpSAx1S2CWTxcZuKyiKJlTsHjiEJyjAqTyAvWg26NBqj5LnhYdWAOP9GL63Tyv9TgZgH5B5I0AUsdR7bYWm+g7jR/bicj8tmll
R7q6jyF5kuVNRXccA7bXBfzyV6Bj0p8bkSxRroQxqAphhiRtNZnAY/0bh0Kn/jbf263S0pq6v5bfSIhNWy30M+5g+WZvu5F/hjXJ
Q7YugJ+3x9xNkujUvQkMcHxsgRN8NcoFRc2lpn4tshx0XLYPck6Ja6Q4XK0sOWugg26IkwzQOtuRycXI3SgVHmsqAe5wtZc2QNKg
g04TgteiibdsYD1R6Z2nkSFkoT68lMG0kDEC2c8cVRfzkIxkxRolJQAr04A05IeerlEtrklYpbv9Yk8ChmcMjisD/7iQ25BX3RSP
rt5hmU1R2sW18Uz2w5AsHDpph6e07FgmAEn96f/P1VSp9c5pXl8Xcy+rJ5qfLfeNPg+uWjZKG3wWMwHog3j9XqeZRlXn6FBPXEkY
Sm5e6rWCeSFzo6nXnsYYiTq5WInDil4RPh+tLh1oBjRL94ui1PdT8cQt9kPzEv5YZsSt5CxrBbP9jz7IWXSQJppkmD4NkwnVwBLC
4u0VgbDdFCtd/rGm2GgYTad2YRH293Tjk/+hBJ59hPemHt+mAJwfK4IZ8RRn+vBF4lyy+5vdRlm6Gk19CoSKgIq48QqBfdXG8XXC
Tx58v7VhuqC8waMNTKZgq0ikhR7EyJMTQJ4jQk85nXltcV34JGWMTVyNd9QupA+Wldm4y5qH4vrJ744IinQx18LZILpPlhcJiZfO
3vbr5fm9mnbIR0ih4NsTkK+z8vERp1U/Zi/VT/hUMyz/PaOjad1wCSJtUt63ZxjKP9HI4x1hmRIJuvv8jrHHrc+N+MXoMQQIEpog
GswyXh4GFwLvnbMz/HgmmAF5tl6hvxCHpfJFOqN/F4siv1BPV9E8707W7k1bncAhut6HJupX+zlscMum4MkchhZeL11bOgLMEq7K
W++yU7eD0sj4U2aK6Oi10atwkYuiJjs2yGUNEjavRLRX0jU9ryfy75wUQUCyXQ7R9LMNaE/67TTibv6DtevrRiDO/Tahg0743QtV
AJT2bYJuC5Thj2MVsgAA/vOnu1+Y/Y/Bj62Khujx451Z/6hS+Sv8eDb3Jbdyfv9TSnWwNFTAF3w7prqf2ZnByLwKeERsGO0FC7uh
jUh9IEp+bpsPOOkNyvnBJ+dwCRZ1DkxKFp8QD2w8t4opRNgsU0pkcTCoHyS3YqVTtFlEFR9CbjLg42atENaGDCjXwymxNJwWm2z4
95E5hsMsLXw7zOOXW9ZAzmT6GUBcTxfHe/r4aUOfCk2KEx+xv/pcUVe628tYRa5LvM4+MK481pfM/MZosPmGKAwueWHSzklQJqvQ
A1w53GVX9kCDQHYFdwAAPasSMisABJ+EvLPA7gMaP1vrGfwi2RCUuQkSbSCb5FDgF55F8IVh1wMIuorHqJXGpzHgCRfrZW1w5U7X
SCZVaAeQ4yFKOUmVgEipri9oEJ7ET28aXe1gxfx5iRjVev/DfKlhdTLcOxGtht/25ltamY0h5YBUOYxcuD7aYaUmaA8zYFbRew1b
LleALd5iJAASDzXAFVpEBpKpSYdfm8VTivY8QZl5rwXHDYUeKQD1DOWeCIzeNfeT8Hq9EkEX/W5mOTkqIOonM3ErpA9M53TxhPm1
JD5Mve5ktmKx5xNZTRaoCFjNEVjPUWpQDL+TLT7I4ct5/oN3KoRV7aMeSPIfzivBNuFX2SxR6Aq5qgA2GcIbuDLpqUkx0OBrjHdc
fDs3i/vartMIlvlhFhzDnrOCADb7ylZaHRtVGZUbW0+/zi5afONDWnwRIGaMBwI8UZfCTZEO6tu7kNs9PylrPOc9FF4Cnk9MgzNb
+iGGH9DtXIwaCU0vKi2kavmil1bmr8Anf7qWNHU5UKVv8xRNEFMWTdiOdmktkM30Pm40oKmP3hFWfzhw3jHpfffKKL/7g4N6UGoP
G+hl1YKuat6MKpMwaX/U/yfou86l7XOHbOuv+2+T3faz1/bVZ31K7aE6RkxcxY2LXaLBWzLAdfk/iIqUeG9xwyjzWuAWTdsHmQa4
b+alTlGri7H53gFKBuw9kP/j0m30S0i3qGQrmru1hc4GIAGTJ6k27iBMgQP/teo6Q0yL+flYaQ8g2lHH49Jq5ma4pGdyGdbQuv/8
kWTnimiPI5NKQdwLX/uGg4s0FeTuhz4ZpbLZulhBciDHcmddz8YmuxI0TmNtvomOcvAJXF/Eu+PZDEm0EPRZ7E1B/n5B1uBLN4/Q
wbUvGRWtIq5TuJ7vHktGq66q0JxaZSChZtCLqcYgQ/qAcJWzGbWLObs9jRrg49mj2/XMLei5+W3l4WSablPlGf1UQ/y820uitmUd
5la3g6LLAZvBg7Eb72vrEMAMko76Jk0bokjfFFlIt+g+3DwJtm56MshQcRHEcviIvh/7tybG8HwDA754V+WtpuLbPybscuMBSz48
rjHE/A2b6Rg1K+/TyzbJjM46g9q8BJt1Nn9YdqiVmZq7IVLg64cLWIJjFD/yWZPvxTo1bYTUAmp8p+dYjnTO5gXJKsdfKFZnnwLF
ojpUujkKlchiRVdUSHOXg3IYW48Zz7Pc5Sgh7NZKl3vXug4MouOpMGTt+vW9Yqkewrr68uZVduXQ+ymgwECTCcXZO4yR8AsJbRTt
HkoXUn6xnj8VO9uiJyoZaTms0cvPcyOiQdnMi8d1nsXtbJw0ziGNxeaWd5iGheLWfji0wRJZWvR2c0m5Otg2GfdldLYziW0y8vRV
KP5aIHqJrOZ83FuP3xRoaitBePfmeB8mQ17bbd6t6cUUdUoQMbS5XLNAAY7+wYsbnNmch6G21HMxLZYnjqU1Ph0Hv75kqwADwx6e
b4SzwKIv17RU2rdvWcv+EciWLTgReEE+2afKsZHCcV42yAxvEMsA3gAwYA44nApuchnJ6XKkq1XtBJVqs4novXpaDda2cu/Dewke
RoWM8awn8lrP1ldFxaHx+WdmM/hlsaW9P+qaJ9NSpWdPdM+kL2/tqwJs5MkumtdvwjaZBApyQ91y49xkb1bQcQEfk4RTAH0kxD2m
l6LmAWdhlTPM8CAHh+dyprGtojQTuZ7VhrhpmwW4kqXn5cvsvdLm1q97EgSXfVi1xb7zjH0TgxByG23i8oTD9Avz8jpbzmCXOi0M
QY7LnySm6yjrw1+Gw9XI8Y4eoCELUsRjyAAyZg2aJa49utGPiZrt7leB7IfIAM4e99B1ZhqmQ8mxoDRPbs8rhkf/vHdsP8FsVM/L
LZtQHfPpIXHyJSp3vvojlRTVRn1gkf2IzGwHdkzPNRDm0DTOZjgzoYvIgm3cWxRFgj3s49j4+3fMQ/RuqQniVKA5wVaXL9jb90Zr
uswrPAdRetRpFGmh3dnfWP+Pg0EW8DbfuCHpk0qhFzvmKP10vor2CI5B/xAX8Y2mEzlYHC8Ke0gUshKqLB/X/wj9NwWn1jt0VTaF
vuNFx8YRY9346FgazybSma9K+n/bmZTmKfi4wah92hmb40n38H44E0H5YoBO9lNe24QM20Lg/NNYm6ujinnaNDVhrd6uTChuBm1h
0gRLfLTbcQm4/w7WMPniAMaEHNAkFLeuORtOToWcJBharHtHfOPU/9hxuAP4kTzI7+KgbUD9ZGpGHDKwVvbuHNIVbs65QebgOILK
z6ZXs+F3l565O5zZYqbUXhgxm6qfWEKLLK3wce+ikyrGEuNL9j0vTBt3Zw+uBvcHeuyXr3b2/u7O0Zq6Zf+/2+GnekAsih8Xy4Hh
6hRP+bzHJXKXkqXwis2SitjKwnQCnASD0/m7qWcilY7x8jEjVIqh4X/JriK+b54eEzn1Ry3k73by/Kj+7p7GL/yNkzXo9TEvF0zE
Vr/tutjKKwQ26naSvmsRd9Cr3pvzJ8UgILyeYEBcJKHTf7PfRyib8bYYh1PdM8xutm/i5eHVa2H92vWWMaix9M/XrGrhKw6Qshl3
kgK25sXJV+kdMcb0sP9SljmAtGxrXRV6K7ohUsMCskwZEJtawXIACIEywsKdsZkWnq0oz5inj56JB913QW8jh+36BvIlVloUiJqz
NUn4r5+7t6tgOvFFNvllF0nCmYLqcX1F4oTOXNQ0+yVcjeY/m3h6OGLJ/ZS6SYYJiwYK2LuUg/Rqyyiqtxg8NAZY9euQDFTun12X
X3VuZnUrC6NrBnvyl8Jt1MYmpMqaPn6/wwf1I3pG16mn2CFUqrKXpasdBflZcurRz0tHIUGuSIUEGFx942RidxIQeE7Gvn2F4PlT
ByFiSLXB/Yu6iSdz7fU+nk68wPlV1CPZjol9MEl6Oe5YTdtU2/STLDfsuknRC57+rTxJcl9Pkzw1Rk1BvRlhr4xpQoLFa2Vi9jx4
5xLv3frgE7VQAihe0uSBTWgb1Nz1QBcikKyocktpwtHIeSI6G22cnIro2fZPTbk/RKXNN2mZ+Ve04IkNNiTg0KBx54NM+teqmSkl
YuU9TJKLRKxZgI0TCdHMdgUt29FSfw/5o9X+jQS2zhWkJEH9rVNFB8aj+zQwJhcUflh1jijDLcbpTqDR7xDgkxUCvw2XQbLFdze4
B7ksGF57wog52CQNDMLfnyWwlUDhNTD2d6/DEU0yXDwmR0tKtgrRY5Q96PNrQPMtUly3fmmW2uLzuJr7Y0UGaypp5NoosC7TlUX/
5tJ5qshdFPRdICx+JT//uY1iK/Yb/Q2ZL1rUH6PcIwsS2rL9NIZ3v4064W0TpGyh/k3Y5BxMDQB5w2SuCDDJm0Z3ZCtEWSCjU8rO
tsGfDT73oJe3KC5TMAT84WTX3IWZBmgFEVaIan4bFk1P8p8mGChbzHxIhBKTaYMwFXnlJTpPu9mYBIAJssDr+F9eaQAwpPK+uaxs
9ijDAax4gNxmkKMkfaJ8ciXu5qg2WDNTXmowW6n5NHPrw0AKqNWhiDpl2eVz5y69ShFh319cqvjxa6+N+ukLy8yr3/AC6sWXBnmd
Xeff30k1EL3QHXlGTb4wNVIFNtW2tB8syPrmP+eLcZ995FM3MZWbd5IuAzFSpmpcwtR0NPam03fh/0cdXW+5x9n7FJ47+wcrO7PM
mwF1E50Bsoh2aq1ftRViCqVsVtdeav1pn4IH/sZes4A+Y4LLHAWiOeSKJgEnU+xgLq8Q3Cf29VYQ6It2sLkC18FtaKjUyk5KK/F+
GtK3YUP2uPCbgmn9G3uOBeI2OMLem37jzdJJlyp3zsqRoqAbWWfeuChsdd5UBTeJuK6TPK62W43i2ozHQeFilTv328mZNMC5IKVI
OB+hsRzgeFz31AvFYv844i74AW65f41yJ7wRNHTm1JqKUG2aPqhqEGtYo+cqkkwfh/MjplBFJTZpTNLQ4e27DMv9VQAn/a9rdCzb
B7pQFByHecDTFkXfhy73gBifiFykZ/GTTwea7J38EH9wTICAM1Ao3XRFjoSEt+bPYNWWm6E0M6jay7+TZ7E/FWCmCyCMJth9lsn2
n0yHuJiO1r58daTV7WCCUyNAOuIe8nlge6laD313mICF19YwFya22pEL3ECgk6s6eKZeRzu9LCoPIti2MxU75ZeMyN019NtZB1p9
YQBqqoI2M2++VO71Twkm1gVQsfA+1UKzWz0be1W2j6A94n1Olx0/UpI+N0jwyqm711zdy+uJvQ0UPNI75YIMnW4IK/ykYCnPRGRM
FFqN2TYOfzNHsVCJf4JFXaw0eai9t+plVqkNvj8qVoFHTW7TXi2L4udD4TeMuMMFrk87/pn/bHcUSlXuTR3YDrYPuD0+2Ot6mD2E
tzjGBKfN9DDhjrhZcc2vSF4f58GCZI3yv1wQPn/5ATcnqN6UcqoUDJbHz6fXTz0/9CALq+6wkxuylKTXj0y5zzxEztyaZcrT3kaP
u8/7xqlUzVef9WrjT/gDjrHlZLd/BOml94vVvrw4J/PQupz/D02tZkqUxiWiFx5WGNostasNpJgOdECiLC3O7y3pxDrk9epuPPuu
PpTYZyro9X8sb+rFC6yoWnmMOYvGMKlSk7ocGRjvJC5d6UT9ank6XsPOKWbKn4myE7BQdaNXAuPb5ZHreuUsyJ9SEoIIA7EVBmyW
Sl2PkJs5dFss8r1EQTX0r9PGpt7d51/dZ06pQuhRoASnrgdv5Axu49Car+t0Q42unExTm3eCEA2zwkeQHwCFVn4NqF/kxZ95c8Ow
pgDVzzUalnsjbIK1pjlnsVbIBqYBZkm/H5mPjMHicwBDsw9Zvbutu2xshuJwugQYmx4d/FIShPnW19m1AqpOdtPeNx3ZO0fMM4/l
jLDy0RzS9Gdj9ZjpV7CQbwDOPEsKaK02iFt0gR1NOJQV6cyCNrAWAM/AyzpzmWGNvhWtDoE6NGMKnzqmLCQbSmNwfw5juycbpSux
VC98g088C7R+KQX6DES0elsJILzlsx2VHajK6AC2QHJYnIvo8a0CnV/Wk23mNIFCYqu5cQZGm7HP4NkukRMm2qt05sgyYDSDohOn
HGRGS7mkA43Bz69UFZ7RSohkUwB0tIb38hpQAPtVDZGhgRXSiJ7MBFTvC2+CYj6EzgOUOfil/NWCYBOSqx/6H0h7/R+cZ8R49YeX
hCZAUqGW/1Q1uroiGPF35Xq8xj9ksVdCAB3UKSSZ2dKO5W4I1WVrCcbWU5zcU86HVA4grjU4tj/hbYYhdYxZz4VCdHkQuFq6IMiY
YMwJtshw19sp9o/IgNeiDSUYby+ELYQMVUn45CRefxciwtjSCQ1u7NLpT1IWXwqo4BqwdbChSRRVNVfu1gqsVPVconKa3Kl+l280
bORQh5aacHPCfJ6rEv/4UzHAxOKyy9ctok3wvWstbC50zYa8V7CZTWXpfY7ZhO0dYJoAmmDQpZZ0pyRmVHYPx7gBysOACibxoh7L
elVlZ0q6lCWwnqodi8qxmcH8ES8ANCYAtkyxKqy69xzBa+cjc8G99ZbQlZE9J5maSGu2WQ8SPxKLue9YEoL1iOdWx86Vkp0IR+ZP
Bhd+cSKAjuJBGgbAeGAXDqZUUgDVXpHAC1CTdusW86EXhl3Oc9eYRd7dm39MzDOFgoObRed3QeSKrXRFE3Sik4hwn2FS6wF/Ho7t
oraXsT1dpudO+7SQfOu8Zsy5p5SpD1AubOeKbukDGNm/fP0rqwohgi5US+JvuTOzKNZ1sxEnorUcJ4zhRY4NBNRpseushqXVX9Mi
N4fORqmm3B8o9eohKZ7tXPIqXHwH+9D1nyo4AedZoRcjt1jt2BS3bOV1guK2+Iy+l/FT2VjWsGPxmXhedYDhbf7yvLshQkX39Po8
y/Fk28L+Dt6eiP+6SoYaNmHzO9P2FiAfBa+kiNikujSexeQPlTLGS2wM7IkROEeM//FAEgXeaLuIgrNrQR3NEoAUd3gZ4aTh3tmI
jfiTHqkmhIgiuhYK778Dx/JBwe2DV5dUlGsjIKt4qxpcrzRFP0g+W6ngKGLjDp3HXBB4gmjyHReIM0UV+UU6g95vCH9DrCHiqDwT
aVPu35rQ46PfilKR61SMLXdxVirH8SyEcJVllTKMShtlhpAh4FwW9CdGAGwsJs5xLo6eii8X3VQLvdD61vRP2oAjD2cC3CVseEop
ioSuXR9UCzbWszKqr9RERsuo9KREMAp1ZiE+gi5LF91UPrDjsP+rBvO9KZftm7g6fOcRkHbfQd/ROiffHqAbjYcOlrABML5bBoZR
3Uz6cc1JlyYTRKY/cuCt9Z/JS0GST3VyXNwNk9YtYeiG4yYgaVFEAFaaQBha4xjn1ZcBu5oFxYPRhs5l+obKJJuXOFs/YnPsjvMg
iO+LUjvcZJEaMDXrdzQjMt3+jp4AlTdzrU9B0+cdhMP6z3MGru+8bSm4sF20oaPOsCD0yv8+Yfwdwwz89t9tAjPRHWR9V2pqHVen
WrlUadONW7jBNj0h7xIJRkgRDUYiZWSODv40adqaVfucx/CxI5V8ejbkC7QpuGeynLnliD06sksJD2rsAvZwZBbEglW+awgh20HI
1diGZiDBwtfKCuIG9fp7P6Vp9XOiMLlZVt1YjSPwr2X41b+j/RIzWHGW1drpiMhMWHfOkXHGoXpowkZ7X+Ofnxk3qbDgTsOX1lW5
hS8qV9GjNQmwiHIWaNkxRRn0S4uyavUlcPqjU+v/HOns4/2ANUCwZLJvBv4b8jlWSkjatfu0NSj3fIgg08rOfaE2hQGDHBACKWzT
9yFC+yg70jWkjqB+P83cuRrXol2AS+b20RxiVtV1QuufwMP4e/b5J2uQcOpiAolZ1jwp0r3pg14XWemmLHA14vghHa6Nc38DGBKX
FVgiboSuv5Z8uqpbAiuPpnzwPbXyhMXRIGtVJYGZDfiwnPBCMSxphiTTcmYntCZfaR0522RHYSLqiKZvlVNQdv8Vnbwdt2tGkpPz
am8USjSue1lNus+v/XLlP7yoFyu/x4CWrlgZ1AzLFACOn2LVwMILdqUABF6SmCLabDFJzkjQFBqbwFPaQFOYYPfmP2atwRgtN9z4
Lvp3N6WT5rexNnm900G4wWXz1Y48dzEFN3f0GU/nK5XE+nWvYTxazIkpkO+Tm2PcDoT9WrVO8Gh5VKnvjRKJy7NJZoC23661PryN
jGiT7JQz+KxboGHcp7DYt4ZMdqBB1tSnAmvAjQB4Vs7mtPsIZOTQC3uiYIpYW/7dnmK4joCfxsHkHT2sd2+yH/y7Vi47u/sT1KkS
/0k9OxYPUoEoei5AEyMhi8N7Q0VuDSA71RgLQOWaK33yV7MBFOw4wS8IFcDMembguk23dbrg8EiZnNa3xvHNZaDiuLKmKa3olalT
G9jod42zqwZ2ouWg3sdBMEzt7YHJ1Rx+P8Bly6HTv/nlWIzi1ADznBELh42TBuQshn148ymNJNJxLMmStDnnrnm+Ys/+IfojpETS
yNcDNHQQ9GsnTuUVCAt7rK/QHs2Hefx9cHLaT1YphvEK28sr4sc6XDJ1BIGJM00MqbMCYtgtA/LllSD5nIbKq6K5dK9bA/f56BsK
WTL1YYs+WSD8J67arpgPfQa4OPS5R6V22P77X1T5/B1BKjnqK71eYczBi9DT1au/VpV2UkKdZmbD9GndTfKQchMsy9iZIrF3pzmC
+5koteSmSWsm3+0nDLBpt9XMOgPSTyZg5oK6R6Xx7jOE5X6htdDZ35ADzTiAknD7kEyakd+VlYUHEv4n11J1jg2HHV1MlHV5+SCl
DY9t+bJyDqYDdLEmbZkeUGz2zv5aXeMeX/fYs1Y4tldBK4d9tJ77GLxBA068+0z9u/oeqXEf66XWbNnjeqcxbO0vFGrZIDS5aZqF
Yl/d2zaaVB7fr+FesCk2jhXFpIUteUs+xpEuZFKvc17wZLqLbSThzYJe3TKk90y6ipTEgVC7ajM84VkpV9Qdm+j7gIedgBkgs+uP
9rxO1sMbNza62uUlvLNjZYTcOxA5jGil1VJh9QjGHKq+g+0cQrs7Bgzr30PfVgJvjYBljpFGgBuVHNphwQv0NlGojAjfU+QLq6jh
KjPPBJ21PFMcgs5YIoqP0y415Pi1WEsoOYh625wOiUqWD4txdZOcREzSMf0Bos2SfsG5qsu6wtn1eNkl9PwUexi0C4f0Ab93LIxA
X/XLfvSDYuoZrxvAVRSerh5esLbahVH0ubQ3eMDiyfWVuQmQb0SzNod1f1Pf/Aa6T8C22ng1gMIY710XTAGREPj6d/hKLkI8bG1L
o2TY+JJjquJfSoy9eca894Jg4IjjwABgka3CI0bykNGsXEpkT3h7oY5NzISoganhq7NSpxXQGVDe4zsP/sLuDSwT+8eT0CybNFNs
8OpDcb+hnB3WsgBpfHa9XpVNlVb0aI4WpTygzZP4czyS08+cMvGBVICOXYhdCegNVYzMHgNyiqRlpjRKpaDiUlNXNPlU0X8Y6ak0
iwZcsCmcRD1mxJDBnuqcNtr0J/yu1ihTvJ2exiigX5qXAv/QuF0QeAIiIqbtz77gu3p+LLAxWzkD3niQzlhNkw9qVrq5aJ0UsYkh
5F9nI21oPtjzfxd1v9zIPjPv04K09rAEGnOE8aMT8olJW6To2bbqffX2lNah8T+QkZ2wxjl2xr1e0WON+1jOs8Pow2oHCZPe+Ms/
/HyZ7uVsnWKrpW5UWHqDVdq/EjokhddD0I4ZnDQ5V6svg/UuWzvgPAmietV/6FywabmoRTzz+Ret0IGlfpx70VxkjRoKbahs057Y
ZT5qkgm47Ywpmbtp5gx0wCa0Cs1uG4p/j2Z+JW7D5KfbSOHTL8yBhBFdUEF9EWzLKsdx+au/JSawMQ7Is1ap3XyVWXxTQsZ33EGO
3BpoKETNcMXxjzV991y4VPuBOgN/bWkiz6X8Y9ln/e3mtXEPYIGRTJUFdCc96/2vWuV16Q684LgmeBZCdFEQgZgjmVkbCzwzIeNI
qRoASyhAxwWpY3+wc2c45ZH5NCCGfSTbET/f7eQNwOoOeApjAbtPi94+X1hbFFpfJzMu1FJ9X+3wCJQu/VLyqwPlD2zHpNPfBBmp
5y+sZ+8Dt65BIUew91QS983Ht5aO39KMAS8TZ1f+Q0QSdHsO4CXifuNevJmFYpK19ckxTxjstE78Ii7xTC3GSakU1z8swsCn9Wnj
AJHzobBeblTydmgzg5QuZKWdOJAAZr8Qt3OtKAJWke1z8etIYhBUPdTuSnQxuSXJYf/MW6wWZCMAP0yBenLRi2wTxULZe8vHWEHv
+2/lBqSUSDRjif9DQqHXJRkOSLsAyo9mtvyjseQ6zyFOFVN/mSL637+H4f2bGJ7gpGY006z4moWoEk/sXRV7Qur/zV/WndM0/ziX
PpJ7mi8PZXINiMMgG7sBM3zudAEhnqvgC0uTzRBtzL0KNU3QhR7JEm7JLipPeyexRUHsr+KfJYTKVnuEvt16H635VVp3/J5Q4E+B
1Ua86pHKR7TeGhklcVKma+a3ua/eC81lb8l9htQqeXptfbuEo0qsTYGnZSBD/HjtmRjEVPqHydL5dtaRLMwOW4gb9/Cjeci0H0Ka
9dQaegXeLmoftHkAmBM7u+Qu6eQil1WxTztIV8r+GqHjNwJkhPTqAdamMO6OWPsXb40gI6j84B+5Tnf6iGimlqjW0zAQAAmoUp1W
X8+Z8nTBttpotFwwhOXb7lWMKY2GQGCdmS4B1GLNvZ84RrhVhx4z1J2pFImZv3fWVOxvktTNs2YgIFGn3uk7GH+xMXemYgUlL8AV
F3fejvXgFALy48c0G7kzfBJtCAaaYM2WbvN9FyklHaDClQh5uGGOB98kOoJVjQksfJ4AWLAQk0UJUiBj8AHsb2KTcnGNTem5s3Ot
2Y8kK7pVBgUo1MjCImt2dr2TseKwgLXuVocmD2b23NdpcxY9B501i9kKAV76wLbXr6eb1cxW69o99v9bJ7X1y+SrK8vPMteWx3Wq
gz4czpJSdO9IgToVtIVO2JEi/mSPb4xGDZ6cgfuY4v7X6MU8Js8VCor8TarJr3eKT3VEm43mdSAINGv0cbTpfUydR7bn5u+xJv2D
5uwmJuImfxPsphFq7REfeXeVvk/8DeMoOCQTNsIEn/c8ZfHyq9SjbdDSyYg7JHoGcX2U/eKooRhVdRGquBIsa1tt0jb2b3nytj3a
eB91U+jULPmQkoiwI5S8cdnuY3I+vvt5dTm2+2/Q27/r3UgeR5jQcIPa+JuGwCMDQjIfpb1MsuTthDnISlUZ9UkYfIlgMC2yfFg/
syPppmivpw7mD22AJN72vtRSZVAQs3xiSg9OpRygz4rMZtqD+ev1Y8FVEoOzyT2copnzcSZwGinCtoXEsZD5z6uomzj+1a415HST
tGncuRPaIHQex5MHFWV/7YViAsXdVo7CJu0ngoXstrwOLVrqU7nfW3cMlbf86aT/SJtFRU4Uy+O+q8JZy5LdmH3a6y011lZ6kMYN
z7LarIVIlrUEQhKD8J59QWn+t3izCj3owyuRos/2/bKQOmxqA554yWF3LrjsP0UYhLywJAz27K0J2NQYnoXAbm5ByNYRyDjwcqZf
XK9joGeE5w5S+HcMzkxfzVOwcAOKVx9UAPQYduPIUo8DEcSDnuu+erqYJLHdNhkGyuNLlA0Ol2nz+x4tLpxqKkntYgDzd0ZdxUyp
+yaMxkILCP8GztNmnDWCmfQoHOudJJ0Vf+EGxJHZvCOUlxgfQ6QJLPjqzhKgkbsD0Dvof52LFPVC2IyeRRpBp+Co9FCpuuAUTvKA
la+RZ+2/iomcpfQ7qAPejvru0jQeCDQm8MYx5npQTnSwEILfZj+/MikgA0QILXdKSMKZeaulMv9DUqooB0D2ZjRdRzOyQOLigl25
uJRCSZ32KNLoBrtM0OfSrehkwz3HDqFP+2rsXHw7bWmOcKreq2mtiFjAF5dGDog2nPMMT25Dae3dbXFH9BWs/V93mqXQqF9Xn7jl
a9uWvkrVx1HsKaZyZAgYVYMmKIDdFY2GFbW0kl/a2F41tlDV9F+iP9sBG3PY9i0VWLfX6xeRIgYYsFBgABpt7JK6vIJTIC73VvCe
jQpP5LLaX+lCfZPBSmmxY/rt6OJ+lS+gc7PxjOl6gMWzi8M8jIAnYOcILTOIzdnL1dQzpEElfE4N5NvOux/mZIzywrp7n7y0jYTH
T7S4kwU5c1ZCDieyZGmVXSgEZEeaLnM36rPKclSfeOuOTVHc+D6bHD642iZL8lglUMlpCmK+Kki1nVxl9aZzDA2dP8vbyTXxOpmd
H8BHEKbhgWVnoFpzfeRheIuObUyqh4zuXA3GiH7r/LiBM5FY6MMqGT2y/k++Z4wYkAYIFdObUown0kt1hsH+IZlLNV2X4zw4jP0s
Q67n00zsULHivOCzIXpsfaPAIxhnR5OpDqt6nudVcHC3JCUNObPMvEdBbgF7qFDY2sMCc/iOOWi9YCD7tdCQ04QySwYA17j1Jd5w
nd1yr4cYAS+WEWadZYKcWKW+cZkgBGF0NgMqkYF3lrOPASm6JVgYZHawmucJYReAkxH8tzHVlUwRjuELDaBOoVSSCqD3GzaG0Zii
bTydhy1sCrVNEIl5Ic/hp6eMO9N8w9+txO/63OhrElVxsPG4QLpAIX6XBM3O35Mimm/QEeGgbnak7s0mvqvF/l88YHJbQvzlM5Dy
Maezn3jH5XinZ8oD5udW70e3sjjBYLobDvROW7QrwwKayXITbBizN+eWkra7s2K14qxRqPnUvJ7pEzsbXhvrA2zqaoG2DwGdoEpY
+3A9aRw6bAXR5t2sPUSPPFK+YF5HosVpjrvgfl4EeAd3NZTmkBzsOqVYgvhc+QBihnVWUtSJDpeoGS2HViiVwNM0y7bU/esKz4QP
XWda1e8rtRZrQr0f5hvDtAFf/pHDwm9L7a+P+SNGYImwCNjnnqPN94vy+ZAFE4pCEWsWUH9gU/tB8igFxbrs4x93tTgeO0ccVXR3
xvbu5DzdFzZ6JBNB7ZuOM2ooSWMQZ1nDwLnAqry89QTe7jhHMk63WvWskh2+Tl6Y1pl+t06n0UgYnz74sRkkmJhkdqC6snH3V3ku
eC7W2B41luaUSKR4MDv6hYSbgh4nT6I/sfliTjR43X+ScAXTUH7Tp+JzVglAv/MGc5iv/L+M/w/gvgMSo4jfLB0cZTPcUvzAw2PL
SVH4QZDnkFiDw45T5iKYBDd2mNzmwWfGEn5+LPypic7nqqA742djzJgiiDXy+5/9P/vw6lW0xLWzg9/Z/vTCyR5WjEKbVOXKO/28
/Wna3Eqoc2QOYXEgcOIqnypKcPZQHE4FWZgD3VCBwDQIkkwaUDBFyoDQBgi5DvO4rFdg3aSzX5T2frc1JO0lJz5u7h6A2+WvBX8T
GTdsLxUGofaXQQiiAXTVI23swYNloyoyagPL4I29Zm4J2V8rSlv1uklXazpz2KrZsR4of/BAZmJe6wdH+2IpdXgQp4cmMo+wUV/g
JLMySc9s3zQsuB7ASt1lPhWC2DN1MlNJo7RCN5ZBKEfdoZE78wKWdDj7vmz+oRF2ND2Xm4dj8hNoeEMAaABtTEVYXjI+z8vsjnib
l8Y4Q7B8pgxEWEPrluIJg9qmXbkNDBiuV5DfJwcy6SuMiitIJXiyLP6H6MGCIJHDkkVwsiai7DhZRFZHm0jiSmPsGiUScJhR0m7X
eUGJX4quxZjkl/RVW5QxkJ7dH2XxD6AhswZRm7UOHRbR6WxFQX5YHPKC4qssIFw+O41CFfo1vZQHHzfKcqquoWhd87hbV9/w+B3Z
0jMbuEIBKECIcsArgmw7w1U7JHSO6at6jKwOzwgabOvUVZvXmZ2M9cPVjJq4THbaE43wqzw73I5u6Rk4awN2qyckNdTH45PpKYoZ
pGO1oqH6mD0R/S8py4R6iMTUe8ZodeVditL7r06mrCFytbq2jpxfnKhUbvaD7GSIp+s19h5qJvbuBrohN/JkJmBEc7W7u+8LCehn
qWQRpNJd7itsCBkr/ZDvFhZ0wcOAjP/w9G+VcT/LIMCkeZ4xWraW66jq1wlWz60G3mf2l/oSzPJMln0Ih00F8MuElKirGSAGpTZn
Y/2lBAeRK81djCrskcWq1GPQXxUfYILEywRpwuiiY+3TPEnmAlCqZThat5pYFVPJEOYQN44gn/YKIq1EzgoYNrC1PphQmVvYjOuT
dSAgAmdHdZMQoctZWjqGWkdGiEBR9qRuy1lMD/k87nYFg9Oh40eDdbi8jr/Nu33lWoiB77OCugjvBv7edDQSN05tMfZgb7wdmHAh
kXG4sbrCb3+F9jZT+K5fhTCGZlrABoiW0zMQMGp3VvF36ehABKYF+iQYtvaQ4oIhxuujSL2XewHNd9wXE4hy9l7zqxypc3Qk2S/c
dXzbvTkAI5hkiOsjbXB5MQbmHB2wrytwIEvtvWWBozfUuk/iG1vsuHjqiiRN6/NeXKfyc6ss7tVHh4hYEcBbEdwirkbzcqYqKTOO
LKBrAkqbF1mxA0FcIkyWMUdxkQHhsmX8uxVzjggExzYXILQOryf1W4ORtVm93Gk/vvDDyq7bLqBuxPSlgeUxEJF/7bx7Wh24nAGL
NbYPpLl7FygUMxUFgJUI+AUJ7Pppc6ezg5pv0CH8dsmFPO/4M3hNwL2YNo3ssy30fhe1Zgz3+Wa//1BtyTQ+6Dod+lWnchb+/dBZ
+h9b4PuDLEO2Ji/UJ0agJ9hMOtIGw5tThoc6vplxQPBFdmKYsagczg/9oa0QAiB86sfbwLjyqG3V3lg/nEJeFXw6zwX5UYHyI746
PkOhf9jiP8dar8n+TOyOru4U5n2vCXqumwJR/mp0I7DTSCQ7JNfvPCqlFTaoQjbHBYTVqHZKB43P5UdU4yhgdFnPmx1IAB/SaT5h
eAgdg5ZCzchEvfxUXPvp8Xkw8NpsS5ylvbStnma6eqRY8flG+hQl/WsYRA94I9kteecPEq+pNjWgmPxn7xBD/e/ffb8/ayGF7Ejy
Jhi9Vapmc8w2cSICQUo/FgYld7CWnDb/nQ0NAxPJ2/BuFbzWKxZGpYRU0VcvuyIuWKqAxHH6SOgubGuKywL8jgjplLdNYHFIC1Sj
1z7Bh54FV/gGF7cg8cf9g3HNGZF610iEChGWsINfjiWxLkceR/R92Ya29jFP+PDUHWXEYEaBxYHqAAAQ/sKdcRDzhni8+RyK//Tg
NOgJLOlVJ7cPHPm0lNhUFc2uA5Zw84rHQ1IplxrrHAijvweK9AKQfmeN3G9g5EorUlt0zH+5V3nypmlZNqLK3tcnvXTsAMVN+bkT
8PXekDLjtCsd6rbAWpJbO3ji2GrqDtOya1EQh/HSrOyxdecIcmNWa69K4cGDeyAQ6ZgW27b3Pb8ZE8DdUTINU0SuAYql+aai6OAa
WUuDE3FbIgp/I1m7aAtJdItryWdTd3rD3Mvi2bMQf21xPtst7NRaWtKT0wBZF9uKJEkLlOm17e/ISpBYctN1QVPLxQgQ+3+UaYlh
h4Bzcwmto+vvPIrM1t4UyCpjdvWjU9puGGzLIGE1f9LjC3uAG/8eh3oGrFrGPziOudkpvQcJeVrBe6uMMCFARY8fwUzDvPUTLmtt
qA4ESlldSSlRsrfrs7vzXYcXsw///vm+8foEQyfq4NL7I8XMMd4SAEvuW0B7Sm036Lk3gwnLnCHILzXW/2hXfaqoP8gOOEt6Bk+Q
fvbLrau+MfS2gIpUbHLlCycQUmMJNlQ9tY5lbcAefCrr7LB1BEqkGl+w7BFA6c7LZS+mIdXzev/+/cLgavKJyACp4ZMsIzL4dRQW
wX4ht6LDXkV5XFvlh/61IVg59OMg3OVD5i6vJNLVp57clOZFVB+B0cgPjvweVK6O92N6q9ecw5WnpEkI9Xi9TjKlAEK6mk6mblaN
6jpc5jvmk9of2O6c1u0nRmHEUmvSjaKZo3NqaSH+nfaTPwyzAnUQ2SCCYga8Quu08O1zi36FldwQSoTbvlrAt8Omw/nNS97lkaKA
+zpqKcxIezVGh9rMPsDBTqV+007J2vEkWwPCbYvAIIJL64wH7J46/6Xt3oWx5CMReeUZNcFt90+tlf5jxKbUeaUtn7UWVTK+hkbT
OFrkA4HbQdUr7R7i2JNZDX6x+7BWfATcDrC0BoM9O5vv7gu5MUT0kw+ACUFRE/JInDTWYO+eLHIPZII9DalVmB+0mCbUzavbjmj/
t74tzYQBMANkQawACFawE2hmUvTyJavR9vV4qA5TjzSyLKQSmuO8yMwiI4aEdMiIWGMyZFivNUjeSVQqQARk0bqkZeNC8ZWQecLH
6I7nW5i+4qD9y6uzN1k3ClHorosKejcSj9zj0vayvO6GvbVGXXTaehCyQcpG+w6PtIXPbyFpLzex6h91DQVW6VlECDJtQxVCtWo/
yzc1qjXQh+o1Rdg0qrTVlR+OGbC7+qDO7wcLaIAEBBdZxc3vky/kUOLvtfZYmMsCi52X558wgDO2vmiMn+x9Upt5xk92ChcB+jcC
b1w6ZMZdLBUO2pXys8lf1gwfKHA/JI4/R5fEmth4teveY9FuYHaGuS9tSnYUgyuwMtbrtWMVdkx7aJ0VYGFOihkquoYpSiPqS3Sp
0Sp/u0MBohpN44EGk3gnIrp0obR6QvEpx8tNluVkZmfLImgamtyE2InkcC2QHUE52qAOCN1RmafXpdJuMc+9GwsjxBlMj/Jq5XWh
B2O1aWCe5GVoeJ5fYXv3fzJSext2xTPAlp/KV1hFIb0nGzNKAIdYs9xJBgHGHnR6kMfLCk5V6tnteXHg0E4/2sU8UBJZz+kbTG7s
GWKCrCfieysZy3UVgPUaQ2F8tDGdk/H4lU/PbXYJU0uaVd5ZPDYHUA5wmLwRyoQ0+DPTbdLjBLdCEhTl+rAn2YP7w+TXx6kAQh+4
U39tsy+RnOExFfZLpnMQYywb9gOAU0LgyNzSFXDIlBxI8zOyj8dpePTM7NhvmiZPB3kY8D3QEiPv041xJbbi0AIsOWwiyXUq96Mi
w5PZb38XDVB1r3jxY5GzpbhPUViqw3Jz7U2pQjyRrIp+lK4y9DtFHdZb5EM6FLpSS66X+ACJ17YJc6C+z4Prtf9nZPqWAtSoug/R
360lKSOaUJEbirS6oPsfEPRfokN2wE7ngQAF5BnwsvKnDzzoLcMetUNFUWhqjbGAbVlcEFUVCu52lamVp5sYW9nBISoid7X2k6FJ
Se1ah3NdjmyQnxHQDoMYn3e7beR4YOTj2TtE6Abw8hv/6wffAJmznw5F2ZPaupvFqc4tefQ18pFiWg4uOZv1hItAAd2v/lPYgrnp
8FCbPThXp0GqBce50SxzflwAhjNLW5lAa75FyF+yOOdzvk+RU2b32uFjoXTq2r+IJa+tAkcoJHWZmw8cUwnW3+N4L+A7rjiLku04
SURuew7eIj/f7DlmZhyIbTx3dujzTT3nloIyIRCGg7b3WETdG8/WxlRKc82KNhd46O83HK6fFN91m/ltBUwut6cX7b2EwAG73fIK
bKVQCVbRtY1AP7P2VS8KxP3mIfGLnV4T9WfqScgQjhHbrFRprySXKLXHXjyY4ZzVGy+kglSq99oNPlNHIHXY+N2KM4CyMOFTls2D
2KxNlhubmrkTxgt0vlMf8wucGgXfkp5mjqbXOJ7liL4RAnQyyTb55HP+BfxZT/la85pbBz4w8XQwB/KoEhJByxnksc3DEjQYegqB
wzm1RUXpt46jhBFC5Xd7z//QTZ/bvdkAhSTbNuBoAXVZuAVJsJjhjq+oEZmYi3DTjoOxlOYnPFcmSgujzgyptiBD8sjs+hL/jN6t
PEJ6jWaMa9F7hrPwkukEQHbEum+S1VJIXMphMv3vWAHcR4rol90HTzvdq0SvOEBgIp08JvI8ncdtrCJHMIjejMcB06jRql5fp9j3
0Gpsxzq+3dSotXyOpgfCWMnglwVcntdNpENa2bChynTHxOFa/QAEzlwpVseMEK7JjrMLTZNjMy3q7cH3vjhFgCiH2LuaYLNZDK13
bNvsMDvIcPl7tRlpwKC3fdGeDJ5wTAMCNc1i3X8s1/hhKX2jI2vi8Imkj1GV1E5OO1UA4iachrmvtW82vYR9XpoFQAGh5OZLdko4
do+C017zR4CqeWdxk4rI8dgdQ6sm/2UGKSaHu5LUY7I/6l5zVEMamo6KRH4kXwypWHOBgZ9948XVK8kK9V6a5vfDg00sDMAWCSd2
vyICc/MLW5wk1HBOaMkBqmIjOom+bEVtPPgxLrVv/NHEafVoLcgC9A8stYTkifc9Fe1e22w+DxzAhrW4WUPYD6OPJRmtWfeeyGtv
j4aVoLYKoUP6TeYazdel0kMavldXIGlB9UA/QOiunxqPXkeDtPIH2p/11I/UkZ+PC51KlsyYppV3c7Qt3bEe7FeffTPE6ebnsJ+7
sJ2gZjRrIDSiSIUTriOdF8WBYGd4Qi5qjyUz8F55fzd8hiKZwnd5IeXr34JhcgKicf2ltvHc5WudksgHOdrM7/t4YIAKGUoFCy6s
UESn0ApX6flKM0jJ9Mx/HUdYq+Maaew0kmB/EeTs8NZ/ioOcTX7epnCvm0F1h7rWoupbsKe9ikAR8a1qa17e0WOOWbJGYiJ9jGmu
/VUpzG++7uM9yj6xcvb5YiCRYRffFbTfc72DzQjQr0Mtd4Gec6WlsTiSFzcuSRrJyyVpVFOJs1bPsSh7zYewzs6gO6en1x5gO6E0
S/0NcqaLabdm2WSa/gFHe2l09FEN/B5vkg1wea/k/2gHRjRIOo+LFcfTLlJ+09UHRE7gWlaGZ9X0BQZWN+SY8hfK4llbC0k/cW2u
bRR7peHyqlga5kNthuvSGWFlOBHReaRTxd4T9M55Q0wyjzJ3urgsBrU9EkewqJV/jPBNyYwKsS8V3Y/SN8iWG5GBkV2ZLTY0vLtM
IJHTaLhjGy2wfCk6vHJ/Be5AwyRGOHWurbZrXySIr49Tl3YKnXWs+NsvWNiiJbxpzmFAv6uDJnu4x/9Snvpekyn54ZrmmuWu4rWo
WRKrWM1By/0JOvP3pXVAshYKNeKPYE3Iq+AXeshFgp3WQhdAYihRxSUoYm9S+7woZo3sqKtUO/9VF0nDZ9LmA+igw/ny0l+7fAEe
ef0kfC/K1NjE/78FNoUAABICxxYCbdPMth3v/YgIfULCAy5YXNQ4A9akS8xXtKjpnXDJRjiVS6/3+uVx7ogEAJ6Ulx7A3L8IDhZt
r4bf9r/1q3BzbniuxJAvgcqC5wgFJftGA+79/zP0g6mAHmlFvo2OvVW4zXaxa0jNTCc33gJZbUjao7J9T5iJFj+vVW3G6GCTJWlN
YdBD6/JBrnt3w4i3Cy9kmmn7tuXe0fuuI3Uhy/dSoR6wm4iRvqnkiAUWdDGO65J13/t0O359zteQZb3fFXfz6Hz2HdTZ63Q9bXtO
f/nvQ44taf2AaJTrndoZWJ9eQ24NXqvgHsl1unJv4lcchox+eldluY0b6/QCu0ipR+hYPAwoapGlF9MCxVEg33euC5PhvPv80kOJ
eq+k8yy16uEtWwgjraqma0gSuXfNNe9KUWQPgp6GJlhpYKzgBedmMiC4krt91yn8L0Ff9F+rqvp8jystWXaPMvLdnA7dShtSjGGV
/kimJMXlyw+SaYxnsZNkK3Vor/Q3gVNHZ/EAPxaNeNqfFANLUlENfviZRn+wVshm3rICBN872dMmz04Q/PRt35qrnZ6Wd+/NwcPY
zRelDXfdCGNkJsx1y7h7f/ZiNBgr+BQ7RRN/ucFJyQ1Uj6JBs4uAR6cEdP/sA6JXNEqWz+nzZjbg9xC+wfoOtgzoys6vaIJK6vAe
0zADR1k78xkRaup4jeRnhKvN3M39dFgihhmcnkO5DzKMxuHv7S4FuNn0P0uI2L4BM/gZSUVz0ZCnuv+3wd+te1HQ4tzSXU5jzq3i
mypWeKrIwVXRQ/AMpYFdDCfjspFH8V5l3AvJq9s1BcBVw6LkELh0sX8n1zh1kcQEk7Yaa/tWyeNoTam8p89FASp8NoTGtZ8hAbDp
EsBoa5klpHCwXugpG3w+GRk6asey1oHBYVcIXEOFKAN3PGAsGC4w4u+ogImZJnKpv93bnaC2PJeRTkzaJpMCPRsYUSwpgOSpFour
4z9bmgm3iRU6GP9PHpxv9/qF/TxFZQkuYe5e1Y7XQQZAGlUZa3HwjurwnQ5gTSTisEx9mFiQNnZOaPmJoEw3kMzDAAQkx+fbNT2S
Lwsdz83vOEWcNuLV0rAgmd8z7sop8av7IRWwKhFOlNb9p2iAvVA3frgsklWu0xDds0AnnRICO7Dy61qJKum4va1POHli2fjMPWsT
KDGb3Bz8Ya2sBqPPG9pw+5npxTehRbvA+qlF7SH0RHr5f8KoYYKsP4/hVYSqKm+CBfhNwXpysPMmyzBcrsSIAbK3vHMV0cqboAAc
l5Gu0RMAnjQlvFWW3Mh2UVDNohRo+wJgpGCOi0QMcG/kzzbhc151Mcao8NAxRElShs5boVAACR4yr3bUBWtEZH3n7it6lSJ2PF5T
GpSJT3DCXlqXUcQG3qf4+5CGbIjOm/HfiBB2ZlJhIPwmx2WzgbokebGdsx3DOtjojoCmxjwD39Upzk5BeMkxSJZD/T3765WE748B
IhKVjTW9J79QwXXdrXEmqpHyBqYEwKKIHOyxLKmPN0omapw53pBMQobp8S1gMB/VPSi853w1QZq0j9M8WC6my6B8jw3UArUpL6Zw
+czSpLeM2oV0cw8d5ozmvyxd+vcA5OMtvbkTAWioGFlXlLcPUgpOf4AGU/47hTWDUMabJ8VspWfehMdhW4wcf0a+MAzBFtSK3I/3
6ysMlzAwZ4Eiz+OFB68Js6RPnWOMbkekH9G0HTMt1NtChjW8VeSm40QD88wi+Zkk4E18kCamGp4u6LxSYmZN693MOaWCmOWvILSI
y5LeGJYok6A1fZWBBtSbpy5am/eNS/ww5Uqa3cpFlvs7CrWajET5ufLECP3YCdSbzZTtuO2104hNFS1VljTFF9YKVCG8DLJfcvR2
LqpBM/UlSLqqaWTCZSPw7zYVCOHiFIaG5Bqij7mQXP0zYxN296VDooXZBtp75+OHRQFYvnnd381wvNkJ4rSMJRWjZt2wqSWEQ7Q4
fp0FJn+aCWcrDMvzUso40mPu3DrixnlDkrtNVeTXLnrsnjZADqoRRR0nsewVR4/r4WWzF7xF7lazbj73elknKdsMJ0/YZ46gndEZ
04GIvr/rAcRCHYhkRpiq1KRDtx+RL0Mppwq2LeOFjqxlTH/l6W32o6CJs8hoHokRnrugI9MoZHGz3A9tC6VIHMtDsws+U1toNLWb
1UztexASEpMsZtvtGLFBaEdHLynHKIZ7XZPB0NKzKOKYekfEOhaFb6gobp54/vDJG0Pvn+RoT3Vzx3FToQvNRFDCJeh/Sw9BCeTy
izrCBALbPgyJeHCyzII48OwZL7fAV5GeHNHPWn6Dbg4s4omqS6GXUNv3Ywr32Vvvg0OtdBvjaXuREd1bTxqoCLwWg36T1GgRuJwn
yCLp+Xcsb5wSjx6Gs5kSTpWORh4R3GFGeMDiwob5i4IH1O9FmT7FM6KuV0zuRvH0pQNi+4rzSEpcqE0bDBFOyICHtyJNdZ0o2Rkt
DEJJPFxEpv2BuH1TZWMHDE8/e0DNyvGNjZ15A+QHJDP6QGCtYgCFacPnEh5rTp4kDY73EvRsKjsivtYQfldNBwyC9y4K0mXdkJ5c
2ozaqDTp8vnZxJWs42L0F9q9JfV5F3RMslhCEDN4J4gdKVzLZrLr27izTjB+fPoZ05KuiQhkOnDvTysflyvKR2GcUMzb5W816jXf
d3AkqbUNX3s7mM5TL/1oSfdeKBggYiMnXVPsCHkWftE3vxOA+20KmBA0jGPBNLkA8sJhHmFLTOpykXAgg/mUWXMn+qeLTr18nsYc
tg+LZvX7mQ+PWtNm4vkmKLWxiKVe2c/hJmmq0Gugi3YCGlK14VRNXQmT8wcS+6pOe13CI3gA1gASdK3/UV3WLlrUQRw4S/0tj8a6
qH4u6kme3FIt9dOYZHdXlSQKKkr3H6q8M3sw55Ixx7AF3A/00SjREcDDGVsgT1xAYy+nVxGuW8drpv+U1URf5nFV2ti1PagSbPRe
r7HIZILsdPR9ndY0o2WpBry7iJWRNF8+CmP7x49582BsvOFNCeozpb2JqYN7qIGwvyUwbGqasakW6a0SzzWuSSoaNON6bW9DlmEE
OmSPFSQtkP+GMuN5e1oRjFyxzdZGqFsP2jgfGfAdN4nvRmiVnoza80K47mMF2Vb72nq/WMmLRXlZxxc5S7IRdVnZE562beghuUct
l92MstLlkQ0LlCC3Z/LYzo5x2FCZ1m9tfY/QntOR9J+V/K7op+zN3EUziY4W/94QhsRVIgeJOL+hnQPACU0CuE9/3AQtCCCvKHYv
eQIANov87a3PEWlR1NIG9a95Zg7jc6ywJubThvA5ZuDjosnDEofenwm7xzsX3IZCxOvApTLh8JxO99rfgwldFlorHGHkIy6HY6B2
tWerbWA8izWCVYnzL6/4IX7fQITML4F5YC9PT2Zu6I7vEgpEu6xVCaIx9+ZSMHfHibPrw+ynaD2PEajXPMESApvhYMEB5iMFOiMe
D6YEp6a6xofTs725FxrKWng5XNGhvp9D/GdF8SBJW6+se5aF8Ev6CQvuVJWRr6LlQuS85egAANMG2NfNnyU1j18ZQ3Ur2+aXeBmR
kMVlA+EF4bg/m2Od+ZpeOqv0oLcfM8FOX6M3fyh96Q1KsDZ8eJ/p3U/exSAxYS6gJiaEk5gk94MlSjDTcSPdQtVgXj/FFLxEQPAP
dZZ13nfs1v7/GWOWJKNQtqGR3gajgVytdw8sGI+ta83lThVHumdm1vpUysBrLSM3+Dr0vIeOvOU/XZaotsPmVoS3Mw7mUKZFBCMv
yzGicFb65AyB3fPRVNHFzR9OWKfQ1aE2Yvl1+Bft8YdsTii90mjp7De3hE0BCp/7fP7YjWJ2X00ANu4h56jzMc7t4EPkGEaXjjk1
GsZpIikPI5Oqy48wk2wdQ5kRBe5qwwo4L2ZR90gKWBg5L6ddsmJ6aF/grsMCOxLbuh72taoUtpifhdeBGc6tkTzkNNwSx0ESfSES
yProKpr/11lg33ZGwOsAF+YHkr9gELsFWo3xE4S1jzKSCAfUgdi4cjSIQM8mC456ZDK9F0mu3AVVVM4W7aial4eBMEhmw7OwVeQS
DnADMT898lToGBkPRN1Qc7wGOhtp+106PsuTdYeghqeiUujpdlqCkKiqerB/BYqZWyba+c+5hJYmmRXHYFjcPaxgToAGicHB8Ty6
A93beotZO3WSytctp8FRjyhgoBUQIqNJV05BPAVwNh4fllGVpoHliNAKW+VVtqq+BNYI2AUYSkIMmCm3BRO6O5AINxLPgVvc4IZ6
g7lFUoVmN1N7GBcuS9XbAxqH9gGoxvtu671cbIaowvdTK9fctvuIQyBjudLRSjwxQOnKjY74NkSOfLawStu0THTFP8niIqdz8OQX
Oabj2TirTGz14ws9MQHDd4A24AuxgJb6BYCHHTdmt4ITCandHlq3VvcSMavJIG7YeUbbzQyJwqIcB5PH9CT/VpAiYyV1MPoDqNdK
SUNOG+Ja3oM1RJbt6EPTTy2+QkfFzxBXQs+SDx5BcGmNmC9sL3BIt+bGc5zmXO6G5GaiqiFObj8P2E0Is8s7xpKvzrNR9zivY27W
fPbnqUnsqWjSS4FZsyottXOLGRyDB9E7h971Yg5sdWLFeDhEzYLBmFFo+HKtZrkbOxlBJDJdY4cgbzmgFaWWrqLZByDZ3OMpFeK8
4nyE6i1fl0TMTg67bZAVgUFnB8v41YU3ES8KRAuxCM4neTM8eA7MdbgGnvkzXoybwkRppVZgAhR2s3CY/6deF1KMZmFzSTz1uA3P
+ISStRoNWSTrRxpPFN9YzSG8FeYppIXncZZ9PDgxl8lXpHMe1YCiU5mA4wRxFWno2+zzKvGPlvxCoYmv4/P+Xrtq4lxuz4hHxIz8
BxngHMqyWvTTg2jQku3a9jePHpPa6rJvpG/WHdMbts7v6+Ih3VNfBpjcFujp9mkiSe9w+H9H/FwG5Leu5tVk0aG0utiYuo4Bnp+R
cbYoKY3I8YTf/spjeiSReCq0jEt8AeoJMdPguQU2aBbJ80QebCPblFhjQOUyLqq9wgI1u5evkRfbpKrmeQi/ywjgFIOE8Sh7TGHF
UI6kwOIw4vzN5T5rsAm6aPRinKZiH2BYNIsw0O35LUc8Qx+uDYflUQJOOI+p8setkB4SdRJnRHDO3BnTIux8QqQpAxQf7IGKKXdR
zfHWyrGXATAZnXJs8XOEAneYAR/f7931HrLVJs6x0UcB2Dk/PPfDxUmj4gKQ+uIjO6/7YqzUNnqJam2Wp+NONzD4YmvK5FLLTHzw
3HjXFk2QXPeFYUbTKGzpGTvSkWXINzo8mMC2IZ/PXBoQYaI+zIVJKkslZ+IhTljMOFPF9Fw2XZhJsQRRjKhkluT9GMhF/S40GG0o
pOGCNK0UESaKx+dTngVSW+LGjYDr/Ooc1tD8OOQrViplo3zEtORmYqFzagelxkPHkvZm5CA26AMQzV8ReLBN32FtXZVndR9auYxY
oPvAFPXuBWcT2rGa/jk8+7k24ePrCZZQjE2o5EzPNyF/flwjvUBXm+dgXf/cfh0ShAP5Ojs6MAeiAFZueTfzDUHxAlzZ1i2ZpNe4
k+IZ4XAa1mA+rKjxMVhzPlnkwPVQXRV+0UwdjN0fJqk0f/LaryY2hR5e4QCvphCMewTTy6ltRu6mLns8+xXTWImwbPBIOD2KqmO3
b9IeFRNIuQhGR0tEOntDhk4KfWv9yg0jPghnJvZvORC26HiqFnhMIE0DpGiemqUGQcqpRk0WsdO38cy/yqC0eUsIdOll/OpuoGzd
T58r1FwiQxrQOILNlywyEVyIkUCN/hxTqigOnJ7UKWLA5PoJkZKbNuSaUgn2nMCqldwPuOQzcYikbdHu02398L+lo8K5ll2SaKY2
7umqBMxrPVY5eiaTr8IrBN+vEsGM4kYhhfZopHU1MpHarbClBcUiiCjPo2j0mAkjvfx5h2ogE7sS8gsTlUkUaBCAA2L1BqNmH+kc
9jv5kY8qKAk+hHfaV8uhgh1uQ2GlWMKprXIL8/H5IRHh6nRX06LEWfKLrvjgZEMD23iBoBrHyTEWWAwKW0TUABMC5GLdIxh6lQxR
gVV33pobwShnqq9NT5g3BgfU+iV0N/39oFL+Z5lpeO2OFdyf8ZPPxTUw055+dAdFGdTeOXXNBj8dJz8LF8Pb8A4CS7xK8a29B5Ni
2OZKPYej2IMFL2riPrnsUKrebRXOEqTgg7WikzhL2K0U5VqNoTNsEbJcnBUPbWR3yKV+v0jWe7XYkqkRxgQQByh4Bk7kIa/pnclh
eImVbwMqKzBAVB9amfrRnC4p9h4QzZpFatx5T3GTIrk7SsUFdBc2ZOz4EPsRG4Md9W/T2kYdngLvNXCpDnDdELI5VpDTyQNAxTxI
ZAFKomlvRVVeysnK0Q+O5gz4dTmNdLXUGzB65mqfqIjfYdAtM/Q6GHl/eBKUZrb6QQ/EVQIfbI774pvzjf7P1qyX05oAMl0KPGL5
EK9eOCNXb92n5sqxrTFwa1dyUe24R075gq0ZqSM0efxekLigPf8b6zrFon5ywms6k/KKIq9Zh8Bu0WbYgQ5DXTxRxOLY/GJPNhJf
PFDNe8BYS5ZV3/Xx4w1PsCE/aPh6gjw+uTx8Cua/5pLwVUXXj9mNK0suekz0DLyRTIIAuzM5g9urbTj0sLGiWmzLOgttosP0oNAf
HIBnH0Bidn2ktnkRWYAZoLZzrgmbzp4khEeZCf8eA5HtS94eeeOqEatTUxdnu1Djp37jOqECbc+cI8CwJyTBcyl+MzTd3dOR3tb9
897kzMWX1LiZ0uXdydCWT7G3Exypas99XeSpXLDa2nFFJM545zpOHSJyzmmqog0J5enxHkFL5QJ3Mkd/J9OcsScBRTSXT0i8hvdp
/1wZj4eek0NTp088LC2uLxI4/4CK9EO/ZVyewqVe1QfZq1iHH5yEsCYnhDe3La0N7X5KvNRSQMz6ELXeI6qU4yKVa7cPFXvQSNm7
q/17Uqga1hhKLPY8+HSsyrB3OlJdce7yI/tL4ZUlI7syaYRH3SDEdAo2xn62Ixtg/FzExNtwTh201RTSgxnH2a5xQYbxkq4T+NQt
+U/1rnT3YMY5Q4QZHO5dxZenppPKkJfFKKa1sK7v3RR7nK6Xpwd8h6Ht38pGKcjCqgVlG+lKWF1E2S1AhBSSx1KoGuV30cu+kEH9
MKt4KhzIysDE14y/QSkfIECyxvyGrQCWNDqDpgEZp1W27RFwpzChhfUjgvvVnGLXbYBkUSSOMIUkkJoIzv0OO/3MQMCROaM/SyS1
tAWlzHTWk8S8bu8ejBGthh6TRF5XrqVIH/aZCj4rZtGdHz1I6UTI+Hncg9dVwP3sTRwqHSOq2s7ydls0eNCv0HvEfg2GwnzA6zUI
J14IaQzHHR734kC6RjMgNoMLjYVIpxJUvJHoiOVziMrm/TeQ9KY15O7zz+Lp/Cy7pXvOBGFfXVaorI1YhUlOmJUQfXZWTFmI1FQR
dvsn/iB1bw3T3cZflFrWcuY6seXPXjEV38O3g9I5qqdPrnLta5P4/zKNxge0YN0otr2/LWALaaNlVQqvLNYIrjDk0I4AxYJX5hno
Dl51VBO4QJfrED6wUkxhSuot3OGytktKlak8rYtojL/JtlAvCgUih0UGPwXmYMWnnRarzTaOjMRxcP9ESB1zq2lhYVm+9GP22NNn
WyLJhMjKFYvRjIbsPTNCbBPjb08sbmFYVjgbjzJYl/EAME8RMSFCpriC0rE5fFhVA+7AE0J9iXYykJhbPZan1s4hmtkbqppSdLhS
3jSQF6lcaS82/rc0fdznRrfWwGkC2dSp050wvSJLp4IqACb1Bu3XLBuD2c65KYpGiVIp0o8Qt0jGWKE9rvgIABrsHyf/I1VSy129
UJwaBtZ61ueY9jlalBXNFhBOrPfhpjvsO3KM7ev7QTMdAGUnNRkvhu5qRAfnam2SWJuvrHjvhyvM/8itdxv9ekYkVdLtN7RjKETs
fnCxRueT2usVR52EaYcj0dG3XYhXCLgPnZeNMbeg3nrsi1lzpXJ0LcROVgA98wFEshBITpDm1S0fS/ExL7oZWdK+OtQJT5tv/3h8
o7U+FHS5s7cee9KUY2txkkwdMyIqAHp9eCfc23HBl8yAbBteiLut+Ol6/G9d0XO9zfQHEIGvHv5C30QTMK+7bPIY6fXMa+rqdEUv
aGJWEjcac3/ZxyX03A3Z63DyS8yiavbBDxENhnwRtMTIKVuLAK5XvDkr0jSOp3rAEXbYXQB6T3eFF0cGN8cjd8I4+mktAA/dDVhB
c5nGXIw2CXNA58fiDjKJKE6Rl2Lb7xizFXKRJ3VJXvNQG9+mVGouPFumdu8Sp523LRphDsRuFmgDw3vCnjdGa9Ib7GuV5Rs7d0GA
dDPRBTh9co9zki9ym1yh4NaPM8qDkPVQQL4pOIIJFkb97A/Id2DwKb5xnrSjOv6rzZgKj4oxJMbadIKpJa1Gt6cbVqi2xxFHWCUB
tSQtAXj6uI+rX6qk3DJWX9VJJLEAxagPBUmqnIboiGhEM7zAkBNVwaCD0a2RK6b2id4YX9CTm7MetziKPrYfAqjSoCVLiZmN3iM5
ybrjBedowRrtALQRrWHzAPmgIz2Nw1L7zP55s7x/eeP6KUH5ZUrPjPsWovUOn0etfnY1GyptExf2rtkRvSrYaSqm5UBToUoUQN8v
ndfmlhOTvldeZRAB8jFVtvqwFxjtxsOi4mE9pUT6o3RQXQamXi88qeGWxV39K6zvErmFNXZauxK24nhPm5h06paA+Vfwl3uvLYbV
R/irkLREregt95kxTMK8s+6Goqfi3HxW/5WggrA8AuSw7G/6fdNI9up7FdFng4QbnGycS4d5OYtXdhkOgOT+aJUV87pEoINPmy31
3IXXaLd/TA4+iEcbNnfVYNMotyQS4xuULR71tTPg8Ff/wVKr9ZYcdkXzLtHwvU7/1txWpK0NVupxSusP8r19nLyrpV+ueOKtoXsI
u6HLpVCtRig1FRjmLGYyVEQIBC3nLOUnkxYyCUA2IfFW4bOd8ghrm/1An+nkEoNXAsRyGjtuMd1JPSofob4RdL7mMAH75BnQEu4M
f9q9w/xCKIFe3t1/Ra7b7VEMsfX0VqGizTSSExDPRtRLYpJBFPkfNhROyi3eIw9wcNQfSWCYA1eT48z3YTR9NWjEdIbL/8zXvnBQ
ZFWAZM0xHMDU1d8O9G2o/bMNjSaplmn3YyPL+qN2/wnaxcxEamiERx1TrzxIa4hrz1V0k9VFtLS8vDdRZs3CFClkZQ6ZiB7DWYjg
b4FTC5h4AGXMX5x1bDM/BwYHuIwWrpMsXUJilNFcJW65OuXLk0V2mm4F/W7UjJT2bnHsLFafq9KdxVFmb7/Z1DohhYYItePcGS1m
DkUpfL1oC3rX4xEcp5iAXc3ZBJVXOSGWPowBW8FPisRoI3gz9PDb4kDelcq943XEXeztkpzS0eARNgJtfWC8pIBS/tT/CrgiSyjd
AKRc6YnEXaw8tdU+sKBA7eFoet+hBm6BDKhm1gl+QAbP9FDL0yJ16nJoZ/9xlUx3UaFitBLbPpmEZJ33lZnoPuDErhnLqymwDW4T
aJscW4a23FZr7J4ViN4wHwT9JXexh6nWNg3wd9tWKEjK+ceYgYdmnsyZlj5naAKZJDKCAGOr2fBz2iCjsXuWTq5Jnc3naKdWJloJ
AjJp0/8MzEHZmA0JyEqI8Su2jffg7e560CMdU7l+8gyxvC65MkeMfj8aXi8BssbDyPdMo2DPGQTSKFoMUiEtyVj2mdz4DPC4Jhmt
CvWSCpyuDMkR7iogkoo8ak8CPr2ObhF2M2z9Db9pJiswQQO1N2qDSfUh/AYBsKVMr5Jl9RBGmXVVP+5bzmLP9rTe0nty1cF8qf8l
fcGbK9US5KZMSdi31H/C0oZoK1hmvGiUcWD3+bVann6iK/TuV7U/YjsIXpt28U+BSdbLDfAvEPXdZLCL2HqqTNgRvWKwcrO4LI8h
n9gStOOoNOjC08G7Nmp1d8Aayauk2ubuRvUoykIOi3DfMfBD5AVG+by2GvD7Xxf+YCWDpvXmNuXRyXUrjXGxeRWNmEwztm9aqpLa
YdNphsvl+npnoxKKFd0mRZWxgfKWwvsETFuqIxeN+zhDlwf7iGLpbl7lTxzn35tuz1QH3KFaEKVNTm98Y5QtPDTbbZuykiEQv9Rp
xaKNJ7/3fkt7BYRJ5ZBRU0AA+Iwv1qtlv64wr7qnEVeGNaNKhx5Z3/QC9IyQCbwDgoUaH+OrZt7wOLYq4ua1X+1CQ5h9EpZaHCpi
PDJVAHO+1REp8w67MVmVRzRNA2SEGb9oYVfGcnyOsv0Fixaad3kIEJXY4WshwVVPJabWBWGZ6PRBZrCBsX4TIodXyd9aoNPJO8x9
yL3KOFZcoahC4HIwSAOdX1aEz7P1RbvwyT+XSg347dOT3V4TVWKVGEKHdUODk//aX+qc3PkhVgpD7Umsnt2uI8HCnFqfmy8PEVjO
g1pTU/FW/NbVRqUopyRgGq8wD7A+MiqlW6xXJII6zxYjd/OZV+sXz2gmObhAV2r8gv3vCjkxca7ji2/WzTrPr/lZcJreZ+epiYT0
OvpZtHkXOBs44062PTiw0EiRpaSQpT4OuLsVV/ekOfqC+N+JyI24NQqv4chRDIpt6xv+TSpin+irz3IrVyIehEqDRTsAQs3dQxmC
nyzJ10C30DvrJhJn0v0/PU/Vhktf4tt+2C27NPJLwax8TRk0yxvArsO0AyRMfN+vFq48PSiz52+q0g/EjWsl0cGrNrbnr/y0GWmW
laSH4zkHxsYKJAvzGVOCvykY+liOuiFpRHFngdFuon9vECzZ5BBKX41CG9TozReE+NcroSVcelxrxfb1tOvXkAIQjIKz3/wD4OZY
USD3VK84XqQUpUpEf5cY9V32RtKqAJ3S3pgxbg0e76Z5vCeNqajC45R+nAzCAQuNGKE2AdxYCHwwarGhwhIJSkBsfTZgvdcj2fjI
/QWvkm/Bu8est2kfGKOhi/K0GdD6rd8RqeGwTcnGjoG8DaxgHpab+hLkR9P/zGaY69vOOl9AjSJ3OTqRp08llC01+NbeCSTCkcOJ
oy//+LyjKotrRh+1Zqs1bN9xQtgUjfFgHmyCUahoAhdz80tV3TNFUpeWOMhGisTmVRYFwy+igbGww1/ZkH90MPvr3kUJiG1DLmNr
s2KYELQqt4BYDDeDmR3/YqEvUySrfBOILrLcmv8xx3MghekzNdyCgDOD5alPZhuRceGnpMwDUZyw4fF5xQoe61NdbvdMznyDT9Dn
1Ue6mzDMmLCa6XKDCPVOxtKfYPMaVOpVECJ2d65dWi1dKMCVBpNJVAvdJGidvIwMUf8wAv9KxPXOIJx70baHXDjCI3BuwsqDzXYH
mC+yzzTXgEVhbR2catU+tHyZFVgz4j0wtuUL86EsN7p4Fb5KTrkqurDT80wTRq8nyQzwdXS01qHwmKGZmO0lAqjYdD7cIzC5selr
x3G6/wrXq800AqqkNHEC3/CCp3p+m9ha4VhsqmX6AU4EIj0dPMfiYIkPNo8ye7zTPjMxB6HoxEpao9gKxw0Wde84qogcqBXCdT42
vLBPrUdOT+O5bhEPm8wTj2/1Zo3jdZHHuS5NAqgN6AN/+VPwrUhH1qvNaEdaAICuDO45fI4P0miPsCkL89DYP1q1u835yNd2LxgK
8X9abqpgpK024BQSyjzyibaHlLzib6VzWX/WeXG6pA3DOSmKWDtaFzQ67oDCIQLcKVf0o257A3MmaaWvqnvr9TT0hYCWpdMoTpxF
Wx6qe7K9VlpZfUiMgM1CbdltiX4rr9qLNWJSb3ypqRsH/VdCrmVJsMvq1SqN3Rm7n3lSX7XMU3sAIPLRhxElZrXQgeb/S95f101t
0DJWi0OwBuqV8nwGj1FWl3ZIXGniXGREJd+VLyP2NckaV9rN5rUKP0sFqjBzckrJfImt9hrPXJu4p9Y9QyNpSLwek7+C1QYBn/+7
C9HNFo85quYD59hjmNjZeM9quEb9L/ricvkfV7IBLNhit1Su32JfHCsVdlRarCJNQhw1G/7OKizM3lDdlxrpxftaB6yG3G/WmMkk
3Kzd3HWT2nI7CMiDrUxKXpvRiNnBvJQnd9FVsnA+Va8aDNYNmZFtUuK35ZVjMa59Ar8p5AiSTPFbT7eOgxbyu+mj1RnjfVxOIGTL
EFHkU1ZkDa9ht/nMz+Ro27h1QEGtFyKnTRtQbL0YdIMGQ4iYUMZyR1iJSZsq4a9hPqbEMgpAHJ4C+X3a860x8yb01tk9WhUDKao+
HzRXJ4OPz+KtmRxCiSLHjU8HzG2y/R4QM4g131y2d2cZRsIN1VPU8oEXsLKslCuuN+PqKFpVW9Ug9pCaXhWEfgdwfpr4DgNt1OYS
kQKx72thQ6o38m52EBRgfkkGiD8k8CdQMnfwKNmLtozWpa3FTmY9wMXCeSisDkZT8MYNn6vHsCwWdUJzJOlB2kp/PZrdcPu6BAf6
FQmDfMY4HVSFeuy6c0PQkQobe3RPEXb5aqS0llrlJu1SLkvcYNdGfne/ProaPDIt62JROv46vYH9yN4yZi3Jf7sDwB5g91oIztfT
K7/WsE38Aj/MddUUv+vcL2WfblckNY0E/HRQpamooYSslAf09CiuvjtX6Ofuu93nBAg5qoEHrb5tWfDAZ+BSxc2/cpJtt7x82WER
i/pOigwVSvYzzRkZyG9xuWpArZ7cSEpf9uw5VYAZP/Nh0He48jUcWd3MwI4nODN+f53Ec989RruABoe/GFkSxJffI63DqREbPPDY
VVZQ4GA0jNQ26nqcmwI9j6cqC7RViIeYwcLPk7xJ8gxXBwuZmPTUOqgC0zkh/P4f1t6L3nefPIXd5cH44c6GyJqnytb/vN9ZB9Hy
u1zsE1Jbz6XkPOW1hCcmCybAhz+GtvbKVLsxSejAA+IDoATsPdc4RWCUKTtw61ovi5Zu4kYDq8nQQ/isSIORG4X3RUae3fbpPwEK
VgDHL8Ev+2CiCtIpz08Go4LLWJ1UK0CfUUQkDxu3TDdgXFU+/vkjFPNF1DtOF1p4T9RLdEXo2IIv3MXfQvVlf1sIvce7Tl7G8SVW
ADPEmFyqUHPp+p1vNzB6O0xWbQdRKONEDCWn1nYe1YwgWC+pdUuFGKwgswXvWPAgjdqbmprgNntPOTgCfOyED/2bbXODiCqNfS5o
4fcuGI8S6om4b4rA8NOsQIsLPGhhiu09tQl2uCl0eG0bCVfM7HrLqqLV/iorKVLU6n6m/lhtaHZPZRq4Oh+kBFe7IYedwL5zqI4f
pqoCsYG+a7P2t3FNaXiFV77lUfATckIzam34UAwV//w4cfka5xtEVNUyhxgz1Sq7syvXwGPezGEacgr29l2P55mG2f2uUsDxCixL
kylGXHdCWyybR5I7y7mREQuhxCkWSHc9FGTZX6WR0oup719tKb2NJelL7qPmrDKoEig8kidHCRl2UqQ55C1w089TZ7M02YcbQWYZ
s29WOikG3yJEfGElzyogNaHlg5njUsIxO6wW0E+iEWOukSKSKpyI9Cbncwpx5cfQpFWGMkzg04lA9eDPPpWO1tQZwxoThc5zYwrp
taS6zJOWB3oFPXSmr9OdCpLpUoaf0Wymnxhu2KMWSo/85q7oxgw1uPWn+H3+kDR0UZ3cbpPCoQr2tfO2k3+gi6sYigQMzlzDYd7G
RX5hutpOxc904jCMKIRyy54CuFea7ILo3xUr/HoeqTDyZ0+Jg3uQdQn4vrdFrZQnlHAoTq9ZT5rBKX9ow8q0ZocQphDU76mrElKH
0bKQbMv8ONVAMrXRCl/c90LAx50daeFOZrCn5mu9vakzpcN9bOJjLXMYQcEKqEbWCypJ9LprLumYDr2c/il/iUMX0jb/Wzt22g00
0BdqyA+pLe0QZnbzbxyDiIZbjgO9ibYTrHCZ7j8sq3ZBjizYfQSWteRdFx1tKylXoDmCxRu78u8un81XSHbMCa9o3WSFjZ34azE1
AOHkbGXMbZ9k9BqZpuC5qtIuakyV/v3lsavXgJADwwvNpF84olYgMce66PXKHL87mGUSGo3i6IDDpKOQPCBGG7qsIyQrUdeNvjp5
+GCXIX85v2B7A38x0t+DmNJrbe4AfDl4P8ZKb15AF1/5oUTiI3F+A558UuiztcPcgXs9aGTBi1afSsFqDnM0qRhypnSRHbIAFckt
a8F/5SSRS56xrHzQBC9Zw5oTkvoROoBcC+kwbtjEEfDhYLmQfXLlcDABOsgmB2pvPEyCzncyGDnsUT6iN6a7sSDQMZvrILDe9IYt
6H2UCNCpJMSGYooN+PDpfVbyBX8cQci6c/i6dW+E6pipoysmLwIRVlxDvph5A5xwA2DfqPmbOdms+9WN3+4drr/rxbECaQskrUyi
jZ++4z8exacUVv7aAe72MP0eOvXbdRreoYrYTmwFhIx6zoffYq7uzFodjLfrkJKlOByaFuyGTjytRkg+DYf89wDXpsc3/EOPUjJs
NAzR686MbMxsdH5f1wZuVy+Bi9v5a402iIB7RkCkuuRj1R3d434RmDHl2r1nG2/KfvY8hCsdxvFB6wzt0okYz+0hYLinCOH1bvXS
5RVYXlat+Oc0YML6hYGTe1OTXboTPW/99rfiaVdBB3OhHclsgkAt2iY3Y8VjfrYt3D3EgYpXltCYfMkFjGjMj7A0Q5HN0WniUx43
tpe7K0jAuV6ZAYPf/Z4i1DI3WAC/vp71fjLDpGWE8oX+X/Q/GH4Oq/BoCZgum3H61vM9U0SPDNPb4DxMCx5XWpwb9x2EZLVkizNN
FU5Lp21JACCNjLaTnmUd9o5IAlAe8pEQSjFl36g50S6PvfDjZ28KGJcFLCDQwo0Jq9LgdMzlARiB8cjp392hs3GhfhSNOsblKS9r
JibKK7qzpTiqgGWzSmqe6vOb4seeapGHQnOV02+XYyl6qzhxgmFLxJvA4KWNUpypogCSiGgqZAK8MX49ysqhVvbWw34SCDdIDq5K
kxShXI4MNau0EK4jmd6V6fNK/vdAGk0sR/AUC0r4iqee6qv8RF1pJJEGz5AwjnXl3ifbRR6UECEsZqvA76IqmjLQeKT23P2pmon3
GbZ7MiB83Pl4TYR0ikMFdlIOiA+EOlzX/u7bEzHqDY3cdlLOHNySFgmIfmDsdl2EdjhNt0cf44DmH1bAld57+RIKtK+9rQFlQ6Cp
5iJJc/IKi6649DArOYhJElUdjb7lrAYB09yHF92UIdwNnTkNfaTOvDTBfWvVLfN7+/ox6PdUQ0l5xP2dW0fRy0IP9vYx5l+/6Nw+
RQCX6ZoVvTbU42CHMHCPPB1Qxt7pEpLTkdm6GNxoAb4ei3BcPLeVs/YQERrRiODKXCh4lcpXgci4aJQ5s+mjOKNm6KvM+XYxw/E6
pmy7LXKztPr9fo6tV340YcaTOR+/hPoL6vRkEfvUoohv7S1l2Ws7NMxoNzwRxxSeK4RQi9cpLNgMdTdkkDOC9W5R6YGOXuLJhf/i
Ptdvgv62USMloOOyGuFrw7XqNF6xpC9TVT5MSD5oPqFP8SGK3PWSmyimO3HDGTwGHKoC/DqS50pnChel05iu8YeG+zBgejPRy7CS
buR/k1qhBR3QvI+5WWfD2An8G4ishi3avKymu2E3yJKmy++eeGaozqi71AijZtyixRGBPXbYpaLW0idL2YPPkhzVVYQFBqyPRyCS
DVa6ZJF7NXsQ9VsflQfDlY5UB9tCrMoCvWKuHujt2ZYpLphF2LiCcZ6kt0veaBl1m3+9C9kLSev8IngtwyWR7wUh5WtzrR0pjw+V
m5B06DMQJOYQr50nnu5lbUJ2CoPpo7ZwGUjaFO6p8WBsG1weBeXpOEr7Q8qSYgFfj9VYzttUGwiXFYtH0ZYC/+vrzM3+LP6Mkkv2
rMcwnc6jhbUSg8Pp4blr7+wEQOhz0+I1ufZE/MY3t3fIdrEn/29aHrxoOd1oKDeuu78Hp6IO7tcEOSY8mQv6XX2DXbV64kYWdFZh
LUCIca2huaQhzA3Pu7R/LRsVuQMkbeV84srKgFl9m3hG45z8rIoQs5s3ZWWAyCfbpBrfP5i5wzB+ftkVYi5alzoyMTOSDJNRXcTa
JOTmVXQXAgERsALEBIexI/xQtIgeH9Rkfrb6lmnTDTHIlogggs6zSN6Hl15SFFkzTaTKtnRbSy2MS9WcLSaTPjRwoeqZIWO2NbJ2
DG4zTINE9l0dcbMhwQ15J+aJzLtrXEDmyIKw+KUxsKqaNfx/9VAzzKq5T722tfRrt2s04FHdBK1O6GnV+MUe+ua2xhE4Ck8Yyg5X
ebGSPNjOD99h6lplb3rml0Emp0xQTOtWvmq+JHyWvtra9o3VI+YCoFQ91nO1FjiTWqdaFgq0tora2EPNy/O1sj9WY97jKdWvHhZJ
vGmSyRfteQqPeSDFPvgl3QP6DgGJXXPJqA5+Mto8JVz1nCD5PvgyYwI0N9eYKx2NV8IA4/oOePewXXuLYWlqEbaVUDs7o2el/XcP
A58yYJo4qMyj89DRk5ZFGpiRjOZ+kp9zS6tTosUkl2/IJTlw+gvdoZetZv+aPDFyiLqES5/6h318MwuzdBWz6v06dHzkCeKX7jme
AHN4sD55wyZ66U1qBYFMYov+516b5ZLvWbfDYJ+0U1gEjJ+SIKfOBn2ROEpY8DiJk7ozaonGVdxgx2O1fnsGYd+MCb2BlRz51qid
qLHStcMTIUQqSxADbe9PrhD7Hu/GQCROIVDEzoHy6rcniY/0wgw9H8XNrOHGFjfnx8jKSmA+2twnwFv6LHZiyVdB9COKUg6B8e4j
J0MOv8jKWIxN1mQbGeli8HZLNmOwOJrOIMiUfCD0I3TCDlPIdZKwQ5emYuI3v8g54A6houFCJqDoIOUIbolNLJjtqLyL6MY80S9y
kcE/RwnTDtgkK8FO+gJ+zRP0zHyIMIEnjrdb2Gt92Tcna7JRtTCfTyEq2Z5wN+tys2GCrrgEjNEV8GsJKLVyxoXjqDlo24F0PfXV
1R5LHHc7xcOOy7NXepgQh+VG8mu8ce6sw3Sxoi+rZFHCdSOELw9RTllrgEenv1vI9yUSh6y99eHtY4YTnIWupLOIpXGEwq3fx7BW
OXKKNo1ROuBJQGq/FewRZFiNA0cMnwt4bl7+4EyZuFCQcYUiyApy1hDt+bNKHBlqe8/+YZJg7qbsVlzjylPyzlI8zsWCYtmQl7LW
4D6F2eQ63WivMOQlqietiw4dFnvR+b2+9VV7gk19lp737GBz5iJPY3nqH/Ew/a7/Q1o48aMEQROXqwSy9Wb4z9D73Fr7gEgHIXC1
cgIWj6FAwdgf4DiNKzPfDkKn5EVGGG+TM8WmSnHBpYf1zq196Bma5QMAR9Vb+eLXYhVGRqa2Gm79jLYfRBMV4tqyPEoMG7y3zxjr
UWOyl5OXpSJqO56s/Af48FaMQjPQe8ZUfbFlPAk1WPKNzXufwE3k1vUAAqGLjIu21CWAo/8o3qiBPM5vpfegJX0DJcJAmhgpJ8wJ
njLj67cj15GMfApy2BjFhKTB6frecZg2AiFOshD0gCftpHa1XMREYSwO3rebnEg5u4XujDA9QR8SD3s/a1LCrARaU9SmUyThwlkg
rQe852ZaSDGZvNStc3DRi7VhGS/euBZSH8j9tmc8vFKSVdyxJjBwh8d7boM76v0lVj3v2x6GZzn/EdY7RtfgGOAhuZbsz+gpjU+R
fCHBxI7+XTifeKjYMc65Nscaz93NHmCoKKelXdj1dNYRBXClNBxDT80UfFe9mnZiZq7NtZsXN37Sty7s7328SlQDzYroGX01cl+G
ZU4P+nHdEZVCyqW6qCSC4u3TaPDcfCrHDc9PaOGbTpJHbCwzsuUGr/9vul06qq5NZJ63gHcAUy6im5b+KE6f2eJLFo1kLi73tfUS
nsBFv/UdeemjQiAZcv4rku1PSfdH49OLRs8eIX2uqxGJz+mFePUYMrqZX/MyAXgqwQU/acDaega9bF4P7tamOKt5/+CHVQjf1ntf
BzzUDvw7PUP7a2//OceRNwnHzdaRjSp2ncZLh4RGRTBqGNrgtcxUP2VJnNU25e/ktJvUtX63gGEyzJqC2xWnOROVtsRm+wYbFY73
IiynHS6Zo2ZSXJFUYc4fi9hsURsqdjjS4Ae7kLBo18UTFeTz9Cj4hLC+hRZNwh5AX/fk3IV8UUwZGYoRV08caUPIrwA8+MlFFvFS
Jtj1ObucJYKUA2e159le6HZVGbArISkAK/UDpEYpBQHCCRfNzPa4oGLXXASPqfLpOE4nbZDqOjS1tzsYp2YmIiaDgi8cvrf6P00e
Ebtunq4UMun7PSj+az9dGe0r5rqgpkDZJKiyPDzPpCVPoP68/IL8d33sm7tJ+kBBPALVEI5FF+gOqQyygEHa8sMeKnA3oE1Pxr4a
E7Bf2VGqk2uMzpVK4qye2cKOLlS8+XSX0P+YebfB7NO5eYAH3FFO2HHOxaEs43DE4wotSaT38Q4UbhJd5f5xDwoeTomanTqFKMAD
u1+DCZf+vl3FaUFa62G2noY05SuivoNbfHiVkBAG4OR3j4ZkR4JhzzpNRSvtFp7ykZ7h1RmWAdkfe6qvEtX4Df5SKjfuigr07QA2
TXDxjGyB6DSn9P/WUIpuEfZu9b7AlRGxRA9whCFdWD9O+V1gRmCQTgyCxDoBF8MwOvDg/674Zq73WXJlg2sPTfow+ymX+DuNJ5sY
I0Vji4nZFLIRXs0Qpe6IO2y6qfTxj/qOy9Txw5NEHKXHkbxCh/YHTXIF1hur2C5nff/lftGOTKJqGNw9TXLZE5mqLNMQgrUZBaSZ
AqMPYWSHN5xIhCIYyPCwnyWgi3KSYxBX0ALUVGpVFyOLehD/fsQhMJ/KKt4ETzwpA3Mk8qmLaT5UiAxlDkPjO625Jc2bF/CtU9aK
JLlKw0tVxEIllwmausBHjk5o1QBo8QixvLgl4H4aHTOCw5tKuzxIltnjj9NO2TiCS5HLDNPv9xpr67ORQsWyZ7fJkWQvNgnaKe8f
O9Q4KJRGSSrAL45VL29KjYT5bDb6RQg9s7MfGhcU/GyUmu8/gB7b3Q4rSu7i86UbvDAgNuDE4IccepVczKJHzoshflzav/+kcoVl
DZXQPeOHOxir6vMawDT4Ty8CIGU+hW5n3rCPmDCta1KZnqUrmfMi7XiA7lTF3Eisx5ATHid4iv0FDTqZuqrqkMC+NBewczQdg0vm
ihYWB0l2hSMYKsC/wVmzga347NWwGTp9sfjjzZNoBHkEDTgvPCSuTI5mFAKcPR1zU5hSeSWfa7VglP8UrNwph5nTzsuA5+/p1Xf0
GSaHWYwPm/EgEklQfmXeh8UtZ8CIPNxhTLOgGF7X86Wt9lU5GhpywWLcLo8xT/BI3WMV6t5P+8KuZQ9bCxjVIh5Ke/t4HsS2TR4U
tgNNxxbUwXgWYe+7Iu0Kc2NrioRm834/NhPwk2Rth2k2l5ui+fVzffLpHeWDn0JskwQcQn3RptzgABwfNEvl+1bhA+BHsTffxV1S
xqoSWOK++1qWMzg8MDB25zsFrSIatwqAmQyKZC4gaUQpWwPXJP1MabiIP8fKxym1qYU9b4IKOjfsEixwlcJSmC8WK1GeZgAQe1Ni
gbfHn6bgqjKf2JIMWpp6pSjGYkZ0Z9aU/sfKE5CQ6slzWtr0MvwBEAHwUh5sMUe9WOrt+B/aylDaEcBjGh9d1+zHbneaLs7PmPzU
Ns4KLYKrN2Jp02pAYZYAlbaFWJp0N4+1dIRzEgXI13alBWpJgmr0LYbjtkhAh+i4ksS8ZV7womWak/tGseKkb7wvMyNDoi7trMzJ
oW+H2l0xjy16wtDbjDHnXRPjM5tSdBu7vscmO2MTuLYvjPMg9YMnApNGnasJgC/rJWfCCfOpjYUrqVqUf41dkt5mOosSY7D/X9Aj
1xu9ywR6soQO2b1eCThmr8YeXNdPuFqMotGkYbJ49dYBPU85N53Z/xltkiPZ+FpEUny5gskEZX2KLn3NfVA68maDLmAhvYIg8ScT
LcBa6qb4qOra0LPo+sKsigewnuBXmlYm4BupIM5USiQeSbIMqibrqusYFQARsIkUOvmSh/GS55K/nlDQrr32l0f7ReyF+zLnqmqT
GNOy1N7v/hswplFGoPObQIT1tu8uFnUhAJex2ER+fwm1pqJFm4BMhB45zTGJPnF76L2NOSsq0xMLKf7N7JV/foiOgGsiVBiOWOcg
UipC5INcrRnobk7rZUcedawBL7Ycx8DRFvBiZlnoma4wepfdeenEBEPVYaNpA54Fo1zzkuCyYFrr3Zgniqn2hEqntuSUD4tl5Fuc
Urk8PT85MXASVv0nzYkdBPt2Tvkf4oghGFHtQOUlmixPX3BTavePk1fvYjCkO2w+A/tlhmd6GWOe+JEYi3SbH9LHQjt1KhUxktCj
PoUv5DEm0QGq9ccSvtkRFRTPgoMz4R+xV+Ahhhik14N3QsZpvJ7+a/0694aAfeS/rsSoZCT3BN5RGrzcoPCUelZmbcWQTEAOIWaW
FtJegS5BKeU6Lohia4Vo22Xn3kLYo5QUiMofTAl/wK3rfCVP11peCE9+o3NIqDa8m4BtqMhrH6U1DTADhEIFkU/AdnCHWM0pwX4f
XL1t/yjhL4CRnC03mZ7kwEQw6/YYQCQYTsDuyhTUiC5X2n7UC5uu8UkCua3adSd7jOIVoLL5UJYH+a1lIwy6qtvCywVqbPNEeyWZ
0zDZo76KOLxErDzbY6xF1ZSwcszKtuG8JgCRUhcXgGS/8Kk7CmMVit7cpuyoDF6lROPY7imSwWHP8AFLt+2rDS+q8ZczXbLLgjRe
N9cey83vaA0ICaG3wrMFmlJQRuR6ot5yEdarR7LtPRGaZJSQ8qH14lDLAbmr5PKDjSYfBbbFHDgVnsvUiyWlZJn7C1IXPU02JIgC
aJ/T5TrOdEGQbhYv6A2UFMnk8hPXVuekLZT5Sy5UVZKxgjF0AAYItK8sCQLKGlffTJzIFvQsZOpRmZTLGocwNL5jrQLRYM7ZDg+9
xPZ5wCpnTs40u7eCs2/6eN2CBGtYqFgSNHKJmb4T3Tb2Sror96b47NYy3Qem6qmhJR0fwU1pErhw5T1hZJ7VGp5Y5AzwfTPYOWrU
JuNxSJCzNNQ85ckw6fDvndSDGqE+LL+WQVg3Dd8xb3GeIvJ6o0n3R+i1mbd+t8QxRHqLx5ukJTbOiAB42t5vxksxGgE7Owl8IX3Q
aIaTASC8UpNbTOxYjKnw+rHje38W48jjeJkzFJzpxZZdZoq2K/yUwbIQ7GF+GKv2BvbsPf2nJ+lra3Pf879kBCxtRSiMTvwDLla/
BxvZFax6wTd9A3y4biU7SX9fDbhkIO3SmXnn8a1boYMfQ5Cz5wxT0+ztqACU5gw1jJYzp4rRCy9u80gZeHS9CPl1zbwUZw0YuLjB
kxUP95KPjMm+wRl5Bnb1HQ+tnuovafvjq6UYHIpWbGE22/XLOLCJVwJgF7eV+6KfOxrdBk/v0FKJXWAks5FAZQHe5b/fWyehOxrb
Xm3/1y9erj/36SQ+D8E8KlbimM+4xYUoMpsKkwAP5Ijd+w+0+0SUQHK4SIeUZZlbMHZ+reW3W9hv5ber5kXKQCBIjlOH5LlE9M6C
H0puDfjhY2cLTN0oD1la9ZR59W0HkGe0h7aa13zwSvAMAQLT2UhxAlubcoFvSS8D5pgFUwV5IRsHjcuWAkIvBeRtNuWJD9RtKFZP
L7GOcY9dewbx8fNz0OA0StxQTHzjeQ1+n+8t+g4A7rMWtth7JWYDspihnF0rppfqpBuBMdtgez9Cx13iweLbg1Q5GKqGYYBj0dpW
lVOqSSviyWdaGRYHFtyOWespMEcPrkk4y1FnOJ1qqfJ9lQHN47U2E6oGiI78IefUENnfMkUr4SfHXjWvLRYsLq6Q5rpKL0C57m3e
xaC1g0FNhqNri/b5HWgDbaopSRLvJ/9YEvcVCqCKC8hON0+V34n/2YTwr8dGOyd1xSoE6eMytRgZsZkJzXbTtmo187Me5V+Z7Pgr
mceFiBZmGh8/tbgbql1aTyaG9Yx+LQpRxZpI4d2U0sE7Ib+OxLzgdjqF1dIZJhpXMSLpJiKAhNAgt/ZJDcnbtQv8I6kfWF9sI3nT
HkhYBma36L0ohVhmj/lQuw1qN2y2SHXOHWesJctWn+9SdVYbUHCrUjssOt405fhbXdGEr2zb5Iy+GBjcGySA/ttAS+SmrZwCFnka
OFGmSVbab3dlE1L1GH8rq4zdVpZGWgjHpJuKPtbzAszdHziO4dkn1sS8p8CZHo+kJtEEnDqZdngPSd4OAfMdFCS154cQ+OC74UeK
fE+D/GPEyJrQgPWFy7xlj6WL9X/FZMFgR1FL9esF4G797o+w0tjU6nCvJklrcCigjTJqcKVfLyYFmaL0XGRexHRI4Ue0wzayCOPH
/PgSQ/9Sq5yuRV8GaY7PFIdjq12GQnt2r08MLTjeMMgNhZz7+XGvqbHWDDeoED8nHa5oTD1tJ3TnPsZ8g96OeveZyHTKdXoAVApZ
85MwjNI9QfwUUmUjFQbYhsQvI4T7uY/L1xA2jxBZxxBxXaBdNV2DXdt/ckbBENuiJR1IXWg9fRiQgddsm0tRPTOdqsB47PhdtgBe
5mT+YbPJRGzHMLhsCAgDO5gQbAVW/RHPAX9+nyXTlOHd2ai5smMJTyMxCr3E4u2wsK9VUEQs/rowHU7BSldk4f7ya3dsjUTrgep+
pEEjpZhLjhCUUhH66dLYAiM4xGSrPJGLShZl2xtOitLkebVYAAksDetnIj71nepHEw6Mq7SbL6vKw9yNRI5htblUex4Q3rgWCiTs
2s55RRZxtn5kNM66w4MEbA5AmLMXHG3OWBuucE1v4wqQh04enhosRQpAPYC9x/Fp0cnXbwyB1jZwFc8W2DJrvAKl6z6zrIkMZV6l
/0k+c3ZZwVMSoDCoSIRX2WjUujw5vbTd9hCzjVU2T6TKvClU9UyhCRXaJRpfJyDkU4anR3w9SWdilnVDFfQxkhLUvWQNPJp/Q80y
NR8HEa+FrIdLKkauBhZk+zPUVYMEAicsZtr0ettHfxX9PKEYggkBXkQM+AIXQj95Q+QzkENb/iR7HMrHWDJoYWSNDVLMkB2S/v4m
oda7i1ADymg66NKTH/BKfeC5wBXSPWJwMAPesDgSlkAycLegqOPcVuMQY4I/2Vm+R1V/shcGZbVqeyzYZZkx8iyCWyNdbybkm/mj
4mxxrZd/uy/9vb3xONXefVn1KcNO19luGHvHkqkETNMpmZzGArLh/sSxZ7KL5QU9gGpV83AbBTwTVaq8wBjpZ17tcF4VKz/1L1gC
VAIrx5aqYicsqzup6yoVlylJOhiXOpidGWgEPSYAz3ycGLz1xMMe5ksXqKhAXsKQJJ+4m+/Ah1ttOvNhB6LlhWdijup9qtyud/9M
/lWfuHwxVBHfrJ3ldQ2syXAs8YlpoxS/Yw+lRNibEn+V4uzxFh2c3xgaKYwXSmUIrMYBu0nU0hucK1j9zG633SVdRHhKiHmurw9d
4GrKptiZrJjbhzXsgM+hS8m4s/7aU8g8D6MWeOZq9ue0uj/xFlCz62lCUQEG/kChXqoHTbVJ6pCzAXC8QFThleyOUnoofrsp32qL
VMCFupz/zDM53ZubVXnvwd2w9lV8tQHNKMHB7dsNk+OSGzMGSc0YC3LmJokO2BLgFAmAjh/tlQKRSL7DQtxGopWPSJb7GgQtiaUZ
Jz+uPPk2KxrVvIHHZzF06xbm4R8fjYuKxOl0+WWisUaulY522ac7sAClKp7Ibw/DwQ4bCRAfhqSV1wjPbVokvb32RgxmbHfgVCGC
9deOAqgpUztggaWSK9bMEn4cikHBqUFTl0May3xkmDLWO9JZLE5wxYA8Ogp8c1qFH7M7SgGAOcbZzLJwxzh4egKuiDxUbB1xtdkI
qJ1m/Ahare3YcV68f+MkxGxbHfhdOCQfrDyraKf7iTZtp2MSDOScNene05hN4YljFlXu4S3jj0ttYj9UdTxpIhEgqAmDJPL6DvLC
a7KTXLJIcyERLtT0CAoKPixQSjlzn8+MpiTwr8cdF6TxP0KVE9HF+18VPvzEr0mPX/QkXpy5Q7Q5UKH1ECC+cNfLCqc9V2zqME0P
tOZS+anXTmb/9CKQCrzEtBzhXuwae+c9eIJAGMpsy3RcKlVbsemoQ4/WoD2gzmTshXJe4R1CXrh0s1nJV1TlJXZnMsz0vUc2PjeQ
ooblAyznHQd2Yf6J/o5kLmk4OciH0+D8WBbbjuyCYmxVfXypmpRHTOtlNotbykTigJnGniOWTFECsO2TfaaTyuruHwADceiSaXGp
qaKwTL8KoF3vJ7HQh7LrpJNgR5lsJ2QTNBshYf/+lWrEGiZYf74Gj6nHQKKsuYRDcbC119yevWm4TolYUI8KZNkwdzyHieg+lvjL
EcNKLXkm2wAJPnvMeEBjbxKQWZIK4pVCZnyUmOLQ5MNhT2+ijPRvmx8gMlSqgp/ubhbF8KvpJglgR9yBlNeSijqNm3idBfYm2I5K
rXUYdfZLjojUidXkQ9QFfrf60r5C5akvNuWcIlQsM6+dmWvi7Y3yq8UASFNjZIav3YuGi7FIa2ydUlFfD/Mi1F5Cq9Bfneoh0aU+
4lHCqcgEO3p+DoldQUUQw7JeQ4OpPnWh4cBvr54eEMZoHVQCh1mQzwzo7xSz3/5z/hVeUrSz72LOdOHYnrXZbRrZ0pZfULJAv66t
dJmY82xEp3KN+3Se+p+y7ZPzZVIR8GEQhWmaEILnWCJPLO6dPW0Y2vtIFiQBsjS9T4vvgHnnbHnRNG9qt1G0xs3w9xlMA3VaDhsb
rzLuKbcJcNSldmZ+FzBonJ+9TCUt24UGbXeKdTFxZ+eZQvm+Yt/0mfIgup+WsEPI95RmgoV6US4dJf1uvA+16wKvBAI1Mq5ua6u/
hOLXThND3D2eJIvxZ6lepB0nmcXyPNzeJOKIDrUrX8M9wsDrnSB4FnsyGRDzJahhmpoD4VSeGRzPdacCl7fKw90plBa3ByWOLu1i
hSAAjlPHI/rJeMXstBJ4yDpDXzDaO+RtHCh1NIgRjxdM4MlU19SWVJSh8FVJSNPMy8VguI9VkYLB9o4QEqR0DO5LQvqq/piMyxXg
axLAka97H6qU8m8lpHzgJTqQuuTKulHLq0D7bXx5jByHDoGN1TRwuD0YCbHa6T2J+wZvzOEzjSNW4NT/06urDBGw4+wlDFaU26hg
IEtgywk4wprhbNM/PRgzGXbQR62/oTldaQwj/l4xIC65cdaw8/8o/ngOzoZQ0hjmRG5LlTU4B5mk+RjBqsoG8lJH1qfTY+EJ47Q0
TDZOLlKbPID3mxAxjYNw6X9XfLpyGzb6mcT1Jf1N03mB5KqtsBRb96iMMnC3nLnAzD1ERQcPBWGcqsrpX7b1J1Ib1j4CdZduBJPs
4SooqLwzZUMPaCKqjc/ga6xqkfcm6zuhcU6OBXwgeoGZwRxBD2y9n5jzIWWxWRkh0Cmaw+YAV3rXWu00RBo1MF85VqFcI4PTR3j5
3B7ovHPJcPBADaOyd/BVwliX4RD5WxltRi8HQNa23RKB/BxQ3y00zo7e5wokAZoa5BFjZ2o4/pLChKu++nhd7FeuDRnzt0PXNPPK
2aDF00NWLhmd85H69JrANsOLcusawR8Ula0wPkHdcnEpNoDvI10tfyfFhtrIhgYszgf9Jai9jkcRxYPZxljZJwWujlD1oghHvn5q
seB5sRdjzs4kHXJhF/sK2EETjIgjeKhU8kjJw9CpdxdCz1C6K/jxu671Gfok1RtfqDjOrlmtNz63O4Qloja6G94Ntx50CTo8PMqD
TekmI0RN3MTCCGu28bpvX382ce0HUZ7D0gW9YrPdtW53xdcCNdL3t2nsDCVc+KSZdgBNufZZs0g/rsHtm2fdkVHGRgcxDmcmF2++
QmdkCSEOOYtsTWLdMhDSwTMfTIFmiHxzlSi3Nqe+6Er4s/AkjlNbcz+7PksgxNFEJUrx4SDQyQ8hqriAR6jamD+PeIUb4pUJPWHF
pILK/18FGDiR2I0WXfRHUqLNjvuKF2lBjczgCPbl2XOACrSkHMRuznZ/MUwsGLcQBLbQKZarNv4JiEesx1O3qqDkxrxevZ2O/tkI
AL9zID4gh7HhvcxBs0y83IilWMySpoAimEefGce1vnIi291Lmr2oqXLQ4hFsXftAwYjK141aJNJ2RkGJyMzHWe6TRCQXpOkMg1xz
wwUU7igBvSQ9mn7xZjlZ0RXHxfCU4g4vrGlzEaX6ma3BvZr8YfUDUFNvK7/vrePf/awCxYrk/ijwpYhVZCO1/LCYCb41kDlWbZz2
wTY9mtfxWODRMVg4QwnmMBzcEBaUPqOAMFMtXeus+9Bsf2NOiBr9lDCZly21Su3pEcqk3y5V6wBgJeQKT0yAymEKwI13Sp1BQ679
bJnLzTHJc8TENvJn8WQ8CSklzHXy+ZRAuIN/DT3dSat11HD5I1whFqpFRKqaJZRc4IrC5xi68Zp9RZNirFvLwnOj4YABSrLPmRQL
+TM/mlKGACyDVzVmqgmc0GNGBc3KdYSBi+CeEWmTgu3Z9HtMG9XC/Jmhz+ZcLq46VNq/PqRf7Rw6vHrFcgECGr30pCMw4nCg3pgj
3S8zDYSTHZLu8Dkpjl/TnfXdkht8snA+ff+McFV53gR9RmgvL5WJuak4Lu/dV5VbmrhYZ5xuTu5JdgTpdQAgMq/asjgMffc6ngTz
dQ9T9x6qufQ+b8jmI4Z+b8wN24/ary0wtw98h4lF2ohp5kcfIXJpDNqnuI/J5822Xn1QoG024mdPEf2ktK2N+H5sQOKqazkFb0AF
D2Oc3p0TvWFpEhgm1/1srTC+qYmkxh6GUxn9om1rZv9w64djlKBvjYULe65t9R5nY3EjYOTV2SvJ2M7W7QjxhXzkAQihu1yUzJcG
iNUdbN2Ce7hE8x04rRtwow9TPN/TdZ6lIHFlXp9yFYahZ89vYgJSnIzEdPRnzOS1LUEQtN2CkhaPf68YAAGAjXcHoM3vygt2R+Ez
Eci+xyUJsZgA1B4qq3DVqHk4ZrJ0Xi6bsk2aOAryPk6uVObWX1Zyj/qRXPobkbjExa5+h2hjzsmwLJ7oAdfPLEwgN/OKMsn0iVHc
qWg85B8Vn4yZFxLQb4wuNUlwtCGGZyMJKQvFV2ZmIhrStiPkQQr1hekS0YxwrieEVpyCNKcp/5ojxfrkQQ5yUPyeEuOUPLQEYO1E
RgMxoHHYjoWdYleC8Uxj+qD615b4NwgoWU9Zr4ZGWwG/Wc9mLsRdbTCbAatrzR+yKzCL7FsDJMXLKqIUV0MzCXoiIPEmtIHUvWm6
nDILf/be10Z6drDKHSIYQfGfRCKQNGHYtJrl0OrtfkoOZkCMbzqBs5g5y0vPLC6eyvSOEnfMkrUrKZ/0633M/XJlr96uLGx0vwE0
wCGLs0j4Uml/jsjvV2VhI+O49Op4dskjmZBj9jC+/veQtUdVrf9xLjy8VQM+xaBj7DejOxyeAAMnj+VnVU8UZHGr8fdesXJ2IAWZ
gGpqTcBjsFxHl6UpkYQUMGLvC/zRKGs31rmJfauE25gIRc88QeaFwIGoZ5YQASZ3nJymf9M3jyAxnPsFPk4FcX0gyYwdjpOz/L33
Ehn7G8HNbuN3xGqnAak+5n1njk1GVNH2MnUozSvUY+cYixLwjIeJuUg066fbufrUsSRN/NYar8OJJxalhfnm23t+iBZBETNjk9I/
tnaRdPBFcaybLx4w7BfFyLhUAdWEwgFCylo+mptjEr/hcRl4OvWFI90SchnIy4mtBGrF8BxHkuOOSnM5UGx7xOs3/0fwLmHV3BBS
cj9K3IhbkAd3okfxC82goBpx3BifG/g+PTan4OweqYm8rL29WuoQvPer0ARHAACmm9B/rtOBS5ENSGALdGGljE7lzhRxMIxeJnXA
BwNvm7tDDfqZWrK3O+BHarK4NM8w2qXBiRCiiwT36s5EGohJqlLq813aQv4ywjzyuMKpF/Q1X3VnsPU9NPoLMeX8yHP4Ll/BsNjN
var3H1ioX1KTSH1XGurlkoi1d2jrxG7QAVNMONhZVlajd3R8FzALS2hQD+OCQQmqjfAc6G65fBSp6pvmAg8J1vl82Kbl4fkG63JM
LnHe6BC41pl9XEGA+AIhrmgsNXCdx+oNmet6mTLpUbkwaHA35lE4/iiwHxAs7NCCPo6MdhmhfPDWdAS3u6d+WRGcDSH8Zx8AWWm+
VOZnaY/cpYMww3wde/FLY6Z9I9IfqcVxYLqycNScYKGkMODjbFEDwZU2qMp55DwsiTra9Kc46qzSCU1HufIuoirkiCc6PxkaR/Ox
R5aGpLKiLEKVulm+tvSH/Q8a3ZbKg9MNoVA7VT/GVcNaYqkaeU93oflVNGvc/QsqgpDHp8hlWXAnMRA4D4GMR+4DPQYjKARAFzH7
gwKA8SahY0o/2BTyfQv/pFe+QP46usqAp/eYpnxOMt51UQejQWoFKYFZtitgrg2oBpYioH4L0snWQCv8owAD4fCA+RU2V4pmgNYA
XjwF6WtjeCYyhUQvGCUzF3t+7Y0o1RE09b588D+EgPbls/3cuaBkQ0BKtr3hvdqqe06TTRIUvLHnZGq+v1v8391H01H62JBa/4pw
/FfLbsP997GTO61sK9rSVaVAh6u1ffmCqxrx6Bp37jcf/pzIs9EPXEfFHEmTB2uewaXFyZlRa+TmWgBj19C0IC0h+xkrtp2DHvfj
muFfKHbSadFdnv8epGF3/LweRhyUmejYHMbFBLPz/xfZ1QciTcDf2enzWKuOLKX1sSAQKniXglXCGbvjtsjqiJWSaXNEGUDG4v6w
0/pAaSOp+C0PaSeT5uvPpwQzk6tLf0qAiRLqyPnfb7jXKyfgMOMC3rWYAIh00T0MUUPTuDd0p7fQmRSbMZNeBVK5gGS90sI0UKOw
0HgsiY49cFVRkhMFG2DUmiynFL0YDSSXj2ia3fEKx5ErXtUkaDy0kKh+KrPOpWWOjmQfsnmmpPMMsdTxr14PFoRBxDPRm7dWgecz
UhEfHM+NF4z6Xz7UHvbmLXAVY+V/4Owy6W1Rswd808FqCHt8yzKhP951LnckZRppwhSMzb9e2lYHYLUH9yA63IFFhbkWW58f27Fy
eZcX6ulvkEn3e+V4A4ktfaDBGImZJw/Jo102VDKFjIwyg9Rt3NuRjg6QdPqBWpkOhlAre9qzs9SJnQYoM4FtW25iuUyPUnjGvT7o
+zoKvVOchYkJX/ZU0LEdArNbCa6DQwq7R7Z/KdqLbwfElaxVXqdYKEDsBp9KxAr9nBe5pewV50LBXAJlOxDgEasVaDZFz+tXrtuS
isSqogSndRUU987QnjReNF4PLKho5B+/JnvTAysoA1yZWGtyGNu3I/5c57BwHsmWq8cKJsyWf6P8nsWbw1hxSxvPHIJOwha5jIbK
IYs5LNlCxz3jI7FnqWF8CDh6BT0Q1jE4W8ZcCiVEDdjNcP7bedZwQoC9iYkcQhk7vKP/eoCLXqGZ3fSb2VHaRWDG85+r1wImhcOc
/b7siNi/b6aLBZDNQcP39c5Zu7qjT4ZWW84VSEDjwLaw5Vti2fGNJ+MMSfkwm9kqrXX3kQhBWl/QK+GZcVe2nyHOryNsZF9ZNyNK
yu481s7NKdSV3rbe4BC2p9hpV2DFWImEwGIhnsyBlrOhmrNq9oFPtk4nPVWunRhVBEpZ8opg9oxBSLFlirmvam4QL8An4bAw68UN
6AJr9xrYIbRIsXHuDhUokvMyyjaNa2jlZNTLOU6nyknNSFV6QNbW3hpm12+PmRLhZfkVB+Hm+FzKb1+MWKjHm0m2SaF0HBax4DCr
aAAI35C2qpsx3l+dlyZEYigiAk8kDlUkgOhcqOTui0++X90g9megp+eqOPWfS8j7Lx6LfCVjtokc5SAoBUINtt/EhTi7Qz0P60Pr
k+5TzyYu16eaT4TLhHs02TL0m6aYUatNySOYZXpX4/wQMY9jGz2+ND6+SWWdyTkKH5roOJwIdvVbRhK2CYzzaNo2TYSYvygwCBT/
LVhvblNi6th+wbAHa8NVJh0FigXZ4GxOfY37aD3XZIN1fnixDPGJd4v+fK8CMtQkf7bj/hywC8tZ/rOieNR9I/aFrEtV+gdH/ztq
T974xviGcZcS5Yg2TTpKUyepkLg9MYsrTxkup/1sIS3J7ahsxPWirxGUr4osxVfns0srs9OqROM+gnZfOplm5/GQdnReDwYV1YNh
xKOhVNS/rJSwi2b9Uf33D+Ygkej0B/QNGG+xVfQdHbyY8n+b16Kj1qJg7sA1sAPOgbXUQL0l/gI3Q/7al5gSsOed0AptTguvYZQJ
+Llq0/sW5hdojP5+tU42WZmnCHf487sfUoI52VBchU5a9fzfBkCJ0Dl6EqP7QpGZW+Ip3nOpCQmFP+nDKCL9rGwqx/rX1+hUORUu
C9N0AKXXe8gLrUefmDS4znY+mZqBDNK6HGWpY/SzFFSEXylqRyTL4Qcf2OASzbbfnuCPucY0y2TGdoPgt0Hc6u2juwyvDz5TndFr
4Tx2wr3dFSnz3JV8N8Mm5sO2rqQEqMjUJdb7PhhAWrDzIaYSz8Qo/GC8qA1An35xGx9ooLHZmXT2evKCntbS/kRodvNuzD1tpvnb
Z72hlBm3M6rNZpcDMQqLUFHcNeRuEpd8WilJ70lK6LOAHmVAPMeKKACzJUL3bai0GIOILDnh4QgpBGxp0/Zn328WOwSSFAhcxROK
9cFguejiD+7hNbNBuYOHSPiuL1oRcZXI/rJa7RkRZ2yA/e2bbUVADJF/Bfs+7jMRO3mbYdD1WE1788lo7gh0j3wyngsd4FmBwKjp
L9UFCA9b09fUdEy+w/oL3DeW3Shf13U2GLrYsf+FE8nL8gw1WTyRQA1fp6MR8DCXBjx+DA06re7I24SQkpLC8yw2ThHH55EMz4QS
AWbjTsMl2XvP11KiyzZ23pVtUwExPiRiqDqSkWFno1/PLMP3QNYzK2xparWckSVqleJc+iKnOcPLulJGU5KmcRHZNjQRuPMxj3Ca
AqsbkpboFQOq56oROHPw/q6Lb+rSUo4QBeLk+RU9zT8FGqra/Qy2acxbHfhQandF3rGDz1twtH9ant8Xege6isDvGcBhH2CEwR5n
Zm+2QKVdLu/jFOBiJRO1hlKEjN8NI5z2ZkrLhqQTFNwWWbQm39mY9OBED6sHM1ddwUDF0sieLOAYQcUJ2T/Xu4RD42B7G9FK3a7j
WR0S1hVOIuLApOHX/Gefjho2msGWmYoSb2VYTMKvTDrPPvSZ1Q9y3+4+44gvGrDtd7n+DMPpSUNOeEAsWgolx/xbfLXB5P9ZXpXJ
eTzHSPmIrUW6QgtJk0sdvQvUjMK3eEw2yOlaQiRYahPRBKbNtzvux9TlrMgiRPcFk0e83ccaHzLb1I1GVmcFm1LEFJbmpdiv/mID
9ZFR/c3BdPRiIvsMrKkAksyovXIl3Wnn4gW8vaMcVJreNSFrrDkPPmqA+1dJnJ4Kl0o1JKifEyzsV2NtHdtGt7irwCL1F06kBGuc
WIE0nUiqBLFL11ZiaK9nyfIPY2VTlOc7Fa5UCPdI/tZBtX7sMuSdg8zlxRmsYkbSd4K4OYOv2nyKFlVQszDfDoBUvoRSyhzd2qRN
9tq9zoLNTR+vFvwpIJEnJ8v8ExtU6iGfYH6y5W5sqpgDWFJbuQeea67fMWpRE2K7fMbrcc9TLaKtYdX5mu8TUiMokqe89lUEA8PA
QCqWAUz2HMEucZvYW1AkcVwgwknrgVHMSqw3u/TBbxFepvLUivV606iPxxaKeU2lcryTMW28JnLmgk6zUQOcUNqCyGcnD3pGieCr
PudFENnGf6uING6k9RuAc6bECwg8pHkgKIrUA8P5k8F8PlMbMFaHIsxs2kARyPiG/L9MosZOlgAtp9PozjB8/YdIibO8XJq2pGUA
N8wG/N3gcCeiNCU8Elm5/G0AuyWNPdPlA3+n/xapNfQ/K/gBVv8YGpSE7Jtv53qn68EktZpL79zqxkVJU0wOFkBCBRSLNZO2uYLR
tWQ796O2d5qQqAa1/kYcm+4xVPQVAIXAdGjzJXe+ZSg1rgSLjiDZBX/we/C7Yua7ooZ72gL9qYOoVkTM2jv2Fv7rrXNMqebzz6jo
Vcgst4SRAXIT8+S/58SuDPrvtJZQnNTDTPMf4TPGlbdJIJz5qC/lK2y3m3152uBKb8SRxcyKiIBDsBURpi2RmivI7UnjVFWhGEAO
xVLAUzxKTJo3EPeiaNb1QGU/32SJsm/B1l9hviCFG3BNeiVuJ6tQQ/pw8jm4GsFGspIyF/D0wkrG5POmYFVzPAjmYAgQa6YJ+5oE
nPI1zrvlzjPhIKtQ0mWricxhrvusXfxkf3Ys0C7U6fOkdyyA4j8TH5crwRPMKDbckGas5+J5Btc4g3otdMps53ZyevnrUv1GRiGf
9t21rEoX7/Ayc2zBbgCsTyOBWnFoxriDZ+zqyIU88X2dUPippWOvF+cJrZPL2U78ADfl9umneaRPdhynNtcqDt8cNWQn/mSoIP9X
6JZAbbdBdvVdEC5Ypm0Xt5IF7jv+rDPRpwwQgtmgt6djlRbFct9Rzej29UI2PkRzoQ7FYO0xVxFnCEete6NtF0BP1gobKshqnrHB
nnBnFLjQeHGd9FkZPS1cgmcA10Pf6q8avvipt7+ol0xK9FCs38PC0MBiqDXDZ2o2fVGnBFUwot0fFODtotJwXMvlk9tsdE/q4N/3
B2xza4QYN66oTnsFDrI4/tuUGNrLpPpxAIJQLFmMMw7yrdxDnhDaEZIU3MVntVbh9hq+zGI3DMg9Euhz2xoKQuPjqY21dvp/8tJr
w3Kl4BjKLO953n7aTq3gPRBUJuIRjcWWoSl7fFHm8GOIs3Sd+T94n6XgbGbuFgUGIFlS1zSaIDLq9Y4/iKKCRM7Zl57BDtKNyZ4O
gDrn/5Y0B876ol2bXrBdDdXxDrPsTnOPOsYql5dBe7ygnB40dPuPHeQquwC6owdzbyyGnacfUlvgK9Ut6A4Wz1QvRdEIB+vTlGpz
WQJwcmuRjDO9BQzcdUgHSydozWPEhYfYpqO6cFisPfD018woKnKDRNSg/tSzxaBqp82na3yZSbMt6oNcQZwA506LaB70OTeROp7h
x4GqKYTtM0xFf+3MF2iBtZNqit3Cblkc/fkcOzmf1ExvRQwxqIuDx7UGSOVB/E4ZeyFyvuhqQyWwYt4gBX8Alj3sf9zdhMwEiKeo
GI8MqX+iE0Q/u6+eGnmz7lPJnuS8/Z3vmsQc52g4b4i3ondpFq+i/JRO9ctvzf+KGzZh5wqKKujfeoj4QcPuEMX+JFRjNR5wYj0k
q+JLggmv4PZMrQmLZ8J77WinSI+eATz2d7Tepznpz1NeCaKZAki+3HWaf6eP/Ur4p4Cpk/PRDMxJMVfn+6G6H2lvLXOCVjk5nW2V
b3yNtQFvpc/wVnsajLOgvsK7TYabxjJ8wTwgnt3JvkPZ9rleZ2FZLhz33cXSD5PauIInnMgKA8zxyM27U9N72O66gqNM06oHJhSI
xWGN16yHpQC7XZ3C80hgUG+74nxhMzpZn/F+1ov3NX+e7dib/mL2moDhKKs5oNmL5Zw82sAgZuXe4hgVbzPxV4k3H1+xF4S8/LuA
7ELaBUOyXPdeVV2s2nkhgi+qAL/mRhMBBpDzwokeZXDkCBgg9EvoXx8/1pwmPh6HZ5Vgk0vVb2i+NRMCxrcwAVsmxz5Pr5HCh3eV
cFLWJ1PERvft3y7LyyzlOned+S9uS42F1bJXo4QvmE3m8uV3XU40IT1aVLcz1FRZ9wCtbtEy68ocLSWJVkGHGSC+zp3TbYYXcm9s
kHq6eiTEm+AZPTv4oWL0Oxyljnw/HoqaKksLGDZuJqrPqYpCyXC7U66awNbDJKcK6JMwpMHn6eV3H9HzaJu3WaHRgdjrd13f05Ry
i4H0Z/mREAvJJanbXoexDMStVB+aRivr4ICsJPop7F4ZcJ+YQqn5WvVBSPD5Saq8iyymOjma6zlzGWBbTNii0yR5RQ6TbDXoWaa7
XFkWzZhfJc37ylpmqA9SFZ6y3AimLmqaVtmEQ9s2c/JUZzsps9f6sBBH7HFU3hxDUxMMTeFuTBMyrjjONpvMdP3W97SRyT5NKeKk
fT6RvqO5omO6aesPmjUe11GFCTfRsfoWjIbdOjPlhDbDr5IR9WTFjN62JcmQv14bmqpEENTlDXpClgQdIxgf2h2SQFA1+0dqDlgE
rJzAxglA62ntI+3AMSSdiPE+CC8gRfJS2zeKBATmayjfaUXBGf8d4pxJydsXJdRjmAmlMQecAyHE25NKc+Uitgqyn4lHcbsWYwEG
oejA0GZgdRPWao/QCt3kvLygE47X71CEWL49tTh3az4+kwcCSmu2u2tkA5f0vlLT6bMkbSw/djIuu7bkqpX4lKYob8Yodswlp7Fx
yvlQrr/jRsyh0oHgCSYv18Uq8eBihdDlYL8hvkED4aXKPUc9rqLHt1QONDmlTVsLiGnXewr7iLpAwFAQUUYQtiqGoYQJTSLkRBHs
fT4XOobNksc3Yj6+2LdEQTkU7ZAKwctTc1DhaDhddJN6FPU8CX72z2bJ0QQDmB/D/952VMR/Fmo44WeXBprIkPQAAMDTSUUbbN98
31bXj2sbyr0smKlkf8oDw5HFkK8f6grYvIDrTM7mZgnmap47nJw3+9Ky0p3DXNzjh32NoNEN1rlOSFdJ4dai6axIxy2Opkj9X63U
W0ApQ8c0Cuk/9cJdmxc6sDmZnnmPH6s225L1hBnJ5ZrbfW+/YOx1RLp5WWealfNiD9o1QKTGdiR0jrPEpNFI4yJHQO3pp6YYwBmZ
woEnGSnQrxUq07ax0mmgxFYVncNApjM2glrlaQOmL922yXHdJKUgUCetVKrsi0uH+m7lVLUJbAeG8xKOzAzh1OSduc6pfON2xOL0
9thY52GqhSva0kRjdNegnuXxla/2z46xx0hIzRZe2RQ2FyevBsLRYNDnmfj38Li01/IeTfEPEMNOyBLYEW/VEfjJQERhoPopMuWO
wa9VlvgfIi2JLeriPjPgmHHSrRe9pa1S2hCnkFMdAao+JSXE+wznHPBzhEiKCKOOThc9PA+283ttys0M6BQmEaFiyGVV4xD3s6+k
CMTfcXTZrqWYnL4KbO6Q0iFRpJqm9SyqAoCZjV/ps2akEyxy4/XiKB2wrjhdH9bLQm0ZeQe8/yFtdxNzpKyP4RnGwtzyO7lmVUgI
+QbiIylpWkHKJddwN1zTCBVttWUM9Pc3Q4CDtOUF4j9g44x4xVX3VByXslyRcIgmGP+Wg4t8w/UKgIKBc9Byv3UJQKsknuPXBqAD
C9Mzp5NfQXHB1iHQm7EGx9Aac2jGO417ZEuep5vd6C5ytwqkvDRWNd7GLzylmvwbeajNzkSpFO6Q5VTVc68KKC+zBU9prr1XWfRQ
wDrAYMU/kHhFmBYEdfd5WcW7NJfcQsY4Q3oosNh5Z/SEyAQBjW1z3LTPn1C+bphDwmNvK+TOKpQ03PTp8YvAvxiDLuwvRtey42gc
xe2p3hR9JOwWe5C+0IODBsY2KhruMKVPKMt1tNayyRqfD3odYo88iPjAaFIgUVhbgNt2mzsKoR2dFXhGClSvMh7GQft7erKKx33r
o2XtaTo2LsGTxQR2ll1KqNUQRK8MavaUK/MRFIHY1BHG7gy4MbOeEPV6cMWihxQw9+u0VVgRyD9Fr+U2xz1thZIZ5eXOoATA+7ET
zSwJTJ4Ued7Ut1eutG/YCxunpwUL18Oo5dYgVmoXGCKtcJmgOpwN8WAdGkgmS9KIbtMclm65Lj/3+vtdVg3Aii5rCKDq3BEMSt2v
+M4kwpt19y9LEyEe3/77e9g7Gm1pV74cWwGbDKvWTzmDpB/MjFZauQw8BwfV9jAOhqQGsfcS+zTWvkUM6E3DrBSm4vb94Rq3pu4N
XHKnKeDiW3DGFfXj8IvniyMI7agSgyqx2XUszKn2C6UsWal0FefcbrDg3QG7JxtNQkacaaLeQpUK1E9vt7cQu0EGaoNtRFd8IyFJ
JsaH41wyzCvnVqcw4Ypd2dqZ2C3ak5ZmJAS4i10+TfhJQk4nqjXR/W3VmpObnuorWgtZLzSLfyuKXOirmiHrhUmMoCFbwX1e41PI
DhM11tt8IXWgvzD4QsCAW+B2Djltn6DG4vi9OedBhtWg1r2c9Rz28cSycJ/QAjoEFUP/tmhgvQFXz8MbqB7EoE2wRhPjrdwV/fPA
ENHZlBKUU0dDNwN8cAR+PEWVGPsgv3ukRKP5YZRp2JnJUyyDGgXSFhYwzROEKGd2/812I0gWpQZvPapoFgAArXNTeg3XgG/MW3wx
7cCJT7rMJnip377+9H7pajauAo5dYJczpqRsJaZOuDXVQAgQYpxeRlV9GBgAJehoT9Gd3krPxqclPzhpdm6zPgPBVZAPKkJpfA8a
cMd7vWGRjEFT1DsX3qu3HZveELU80N+DhUFDYhFatpnKXw5zc3ZK1Wwdh87Qbpzg5r1ZOSLPTN2FzbxUYTAsy5OXDlP6Qxdu1LH2
JXj6vZycZUY1Mm3RUrRKF8kES8tWMef+fjeJI1s5zcbKYgq9qvo2/as/mVhw8WR2o4X3aHRaJuTEs3mdYeqk7ucG9Z9HACjOuqkv
a1VWWRvAY99qjah669OFjlKn7V0e8OxdFyGZrIJCIv6sbTjMitiYBMrweCK9waNYfdMWBbmqXWW9T4ed0JjT7o6aO+q44m0LAOmi
3zh4jcDoy7L9CpwHtk97d/uRNqlgL7HXn/8zxhT57UPOPBj0mteO8AmdAjnE7Cwwzk7xTsdR2vC4XHtCVa0aF2AbRzYJK8mirL2F
mLk9c93okbIw7HH/fvIV8U1lFjKw4xyY0Yy1DF+dlQpmrnKdx2VaTaJc98Koy4IfOhjpjntragQrRyTkSDg2v3FJKZ+rBOhfn8m5
9AkYfxXrbIPa62nUTJ7PpolrXpycmpyRucE8q2xm1NmIQjFv2QFD+jD8PR20wzuky1p8beZ/W6yIH1iglVGWQnpQwoH4/WEstMWI
x7ht1oKIt4XeRNoL7+82k4S7loAgiDaN0QBVVJkZhCasuBTNm0fepaKb79ub6fHjMsz65laio4Vcr4cdALstHZLYIWnbkPDCIVjC
HGDrkIKisbzv44R5SZK5HhPy9L0ky5Jaja/KNvFIC9YPAfdLCPQx5SbbjN2L2qg9BK/s3OX3zHaMdwG8lXVY+SeKCsm4hPwByLC6
qz3iupNs4Zn4rm73C8upgzudmxvyOZopEazOgX7To6z5bSkl+5S4A3VnkKFeqZ2kRMtj8kWeth4lnEiEgfGbU5ttXaE80mJ4+S++
YQOICFPuV57KcBHN82mk3vUvotSN/aqHC7H4f73sC1f7LjZVUK7BQz+c9gInDKJPveCGUblTHqJw0YvwFoYRJfQbJx96mFxZLjCQ
ByJqADenRrX7I4rV9P5fg6ri77lDqF4KlZqnJtE/Ae8gXa80DYT65l+N3d7PDroYfbrQM5JBs67AFOkGP0xHhoEu03OW0U7kcysw
ohzHvWxwGgAARfTnYWXqyi/NZv+4+AXqVmeB55L7Qmhxu+n9JWQE0/z2i68IMGWohm7FhuPWOtNJL8Xk5SPWsEZetuKmpO1krk4F
yT/mFOzFln30dEf73NY/36ICKe/3vj3VDRJ45bFBaAcIzwtCBJteZzL7wZZzQRu4VP5I/ckwB/08uhp3A854M5Ea6zk/WiEFTNJE
PaiaHq78P1KD+KfDSGQE7MHwIjrEjxJeVCvLXUicesNEb41Yav08uorUXAAEippp5a0Wwj34pEyzFDxBybG9aKemSz+7BRFjgKqP
6JX0UEmvbZSrv28dzjvQSY2/d1vung6v92+1KHwY5QOO/Jo4O6n1WOauJe4xIxvbpQ0hjfWahc8dcQVWHd0Tjj4dxt6pGg4yjfJt
BmWP2v6Rgcscr/AImfUsocTLVOJK+CmjhICxaMEXnUhbk764IQ2INBXn5IkrL54N+w3jDPR17B/9tBNHcoHiWAaBsp87f9JrCZwR
OSv0lw8llfHOAHqZZ/izrupVO5HkixNIZs2we9ArfJPRqL7eE588PFJ03tHGdZNfZnathXF8CwQFfaBvMz2/D5lQToQReZUw6Wsg
HnB32xDgmsq7nT8k4SoF2jMOMkdKfcaR2uHBOknoigvWdZlGt610BO1tYDI50HQmuL/Bv0vy5ukBECXCB7NBtORb5U2WfURYf0dJ
o3pUlQuyIOIzYgXT54SPkiUc1Nwbci6/iSxxqVM4mKNW7UKf31U/pAnTTWIoouXHwj7NnkKKGJIURYMvbC8GU09ZjZrPtlAP7oN/
GuE4jv+onF+fe1n//OVRnSn3KZKmwKT/eYrWd+CQeUV2hv/WkJNW+I2uBve/vV15dOF5icYH7zDiuAHhxocwO4HbrhTi36aPty4T
hl3gAFttmHl/gkuj7q/WtVY2T/7x1cGnatTCX6bnpSngMu0cpPhPkb4e1cTqUuRTXg3hvWhBi37TUwVTOksLoqEQ+tyb6tRg0gwD
R8N3qXZquGRrfSYqDeFLZ8xdrJL/bn1qEx8LihSPR08Z7blgMA5x/b90h02aDos1kAA3LnNJAT/VrZaqsrDn1/yj32iD944Zc56F
oGEorbCNwv5CrPIrqz9bROpuwpS5JC9YfB+GGB+CkL7Dlg/hRPG64Jd26UgzdU9xy4aYzogEmbC35oVq27cqvIJ96BlTxeJtU6Fx
0cNOk20rickvo1Fil+l+ltbC8gYpEYBI9V5skwp7K2TFHdBgFzPnbHsr9cAd0gz+Vytoh1HAqtrNwlv7G39bzZhuGiqliu52bELm
7lT+C7grzrS5mN/VC6mKpuQKjVHFVEF+TPFgukC36bjK+/lC/ysPu2OqCC2HdoZdvcwqotOv34GkdyXZIwAp5fk5afEkUTxMlgHK
1BmNt6EFkh0qa6n1bB4+hbSI9ZD13/3UlYkpYSe8AqVjuwbUELsW02AwcNb3WOn286obWsY/Wr5wmmjxKcf6PeQJnEBVXp1fgrFs
ML1dFesQWxGZIBqKL2f+yS29GXNq4K5B9G55usumSRvubdo53pbM16oPmoQ7C+E2ZcxIHqAq/dBJZUd4oxbYnrWTA4/jOd1YRpkt
A4Iiqp59Sh8Y8odQr9y6L4dNJ+1EUzz+4RKdCARdIzIVOb0p1CUVlnzS+5RQxKnicP0aVedEB4FWUPP6BOA+wGhUwbC8fVU278lu
t06JE3Tb2bhBEL76KCj5EfmegLO5Nd0Vjxtde3XuAAJQouDgJtZ4UqAqmkVWrJneDqNO+6ykELI+17sGuoTq5D0vW0PF02szOFuk
o1fZnBJB0EdFs/VUYC4g515Wx4TddaMpc7o0UjlSBSa/k1ZGtTbvUryCqgIy8eObfeYhAhDI+4qdK4pVoGmRw77tH6SuiQJ0FUVI
PMJI6grv7dyV+xu3SR4JUALwgiH3ogG1nvw5qM8CtugYQBZV6qVZsp0GDBRjl58H7SJi2zxgnGTHdggDGaPX9seY+EPYCX0z2tVf
JTdNdqc+VN1PNgnwR6Q1ZIuzc94oDtzyux9lpEQiuM7rpc4IdDfQ7aEJ+w1O9gv2m1OTEV1IXvVdXfQE46H5U5PPH6vPHhD5KFMF
pEs38CMOtjM/ZBjTzEJOmxjKlp1xbnWLxeXLHKekgJ9QkXGejeUkYZ/CBxqWZJgizkD+JKZyEVp30fwJgTWlo4bylRSANZNLO1sE
N0nA0RR91MVRZ70Sq4cgUTD5HOTksugnVDYVG2GQRfg9Gez9HoejFVwPhU8zf7+TCgDhA4mV9v9UlD4xzhOjJ61Aa5rRebolHPGd
KyEmtJLU6mWfp9w86jcgplBtD5/FWS1HK5pNLrNBi3+l5MuQolJ9SI3r4pHFuQ406OrbJ7TguZ3xg1XD5CECFmti/v3lkfN8rQm8
nWKmi0lkIfN8OIE/q5pnkmajD+yS2oZJ7Y/MfE+ur2IOYsKGxKB1hcwbaPhlp+GimaU/la7gft4k7Q6rIB7dq+Tc5KooDMeeAZON
vG5ZMRMa+rET49eqg9LAXeipq3CUvjEVJZgcC6PFwyRZNVjo/Pm2ul9MyIJ4TOgH8cjfpE/pPmKl4f7xoWp/EAixpt7QaW8IVKQ7
ZbOZKOTFMd0WLW+Ad3RWglOZILAyciqCIT/CsCWwyBbiPTO4a9AH8LXfPrJ6iuYCLXU7p9tjemYA8eAcvhuaGVnyYQNlc5fhsNBc
Y5t7oq8lAVczHgADxJnWfMCVj/S+a2nFtHD9HIk/mcHnTepgHmQ9M2kbmb1/ZQN6ICDhjrp9CFnQenIwCAgZzGRokd9l7CJjfvCf
Ag9u9m8MVrjhrsluNk7YDm8/0z2Mrrd2gCbob+ku7ovXw0yT0MeGB0X6neVYtSwaVpp2DLjHepZ+wowYsZuzC1JUBj2PFA0NvN10
0+EKe42S+OcCnL21pCgTQyG+dvKBCOD6JI0QEZZl+oWt58HxqYChX+Qg/6aJ7eurXUDbdogrz+zqlm5U+XQVw04yrQARkrVvE/rb
0iOerrbq0gdsvlV8O773abBZ70UYpxg93WF4chrWCLjkSoxXpSLNVyR7xdkp9SsrZXdb4A2rLD+fcCpYiIIUysazVvSWzvpLsxt6
L7aqo/a25E7R2E/exbzTvH79JjsQYLrgQHcpeW0O4EyqmuajsvV6tQf8hVQDdfS7MqO4zJ3AMMb1T+Eef0yTqFTj4BGo1U6hLkDh
x7SKcib6+BPmJMx9sYU9ogZI0kFYICfXfK54xP8fJY/Wv7fjR+tQxr4kBynwyiUcT+xPsQLN1Cr+24lJX9uzjjuFCdjhczBEUiTz
YwJvGvknoY4bg67rukpsruph2bOM6k5e2wjujlnfcWfHtJtQSzbEa+pZoBTZa97Cgh6J8WTNOikt2tQFlXlcxCveG9vibKSf/0DN
bLbhghX/zUagfBDqqUNKz7DprjqOtEMVgXUCKd78TwzNusOmGex5RIaIOLadHxaMrK8rGSK0QHeoX4VeGkUR2ktrsMp6jo2nzme8
rVhBPND0SHg42T+U8xDu9fr+ExdvWlo8KnbrFyz45avTq5ZxlAgNvaGOXMd64I+rUZ+4CdcdQYLVIoSo1jzScSk5f454LDePnY24
OhJPn0xlWi0sZsBObO2GTvFEeKrx2SJEBjO2WsjLpBTYW/hRANaJk4fInR0oQEV8XMkFJLalVUEJa38DRipJGkbrtEC6Ak/4KkMf
gWRTXRNfgCopO+ZF9+GunsJJIjDwEc2SwCvRaPaNjBenR5skiR/Y+aSuJmKvwKUoUAePI981XuNY2dPSt+hnBDyJX0vKztneK+d9
UqT5DUAI2ELfq5188EGFG4qPvziKFby6NBhVkLr4dwxPLD2eczwKYg61SwTUYc/zXO1LQj8xyw99Z0i4Snwjt7EQiCkKdmdVZW5V
QM5lB2YflsrGlXypibDLGGrXRYMacNbEZUg/F4HgP118KyZQbWe27umqq5X94OGFQLBsERe2bLwDgKsfCePMjlWTp+bXA7BxXOCy
M+MLclBg6IEAsjYatQW7EVJVSBGVM6Wtq8F1ogcrwKMJOiVOLc96Y7Wsy9dzRQtwM+X9p36ePCOuQqu0y0mT5aa8e0GWePxMaipL
fNHg8jixX500tfGiPgXCIl0wFWE/qcCEzEHc36DcxjN7JANP6xdLzaQnbIZ8yUJOeF3a9TrZwXp0I2gkuAsRpTi9aLlybpINFpLF
SPuSD0JsdFUqIjMCBYPXF5SYseuDdC0tf5XGmZY33OyKeSJDwM75U0IbVQORloFKdSdn2F8GZjW/Cxg2fVbPV1nilT2EQGGKrArn
byKzbcqS7FzF+ndG2V47pORO935RFS+UJ9jp8irDKfBZlqr5W7KaRwPm+nqZGCgrrrupgBLlH6XnQnmvcnY9AHijmDnnLB106Jc8
IIf6RuTQkCp6V20JMHifQl0IePLHkpODO7BvFEC+i7Dxp2ipm/SdvW8BVh8Yk9gA36JkKt45sLTIeyIILKzVhRHXbosAiLKaR82N
srVX8RZ3R65ehFUFiH1wS04k4G9IElGJY1ur/m433K/ilekqmr0zgSxloO6Q3Ooc7wSVk0UqwFICufG9GaE288U8/GdYazH/+5Sg
dns4HVZvtJrYNCiGMyJo+ht2LoJJYJN+CNMooTuUrLdbbEnw0wKBJwA4BLlTXUS0Fqnvwdl7kOAmWs2ygNJyGfOqyCkXXiAAqZHx
WHVzdKlcdQqr5fMZ/rvQQqnB7Zh7tw+/j4ZJVQJ3wD5DP7L2ycGW8U3uoF/i+T1hwkgdei1TBD+pwxMbboH+oCecclu/5+XdgQtU
JcbrYmtd/jutVBTVAwfATS6Gif3QtrxDWGmZkE0KnKrFq4Do+qVodKnWP1OPKI9k3gkOsLw2k2OQyzNFyrMWoG3UPhXht4ytzY1F
U3OgTn1qTt8CVScDFhqmS15LZJB424lp/0GuFrtbY3wDL9BEAO0LpbejQ8G3Bb94qhmqR/z7+iW26NkrN0d1I44i4up3MbF5nIBr
JAtZy+3mImtAJTERNWgbZYXqAPHt3j9Yfr6hHFgTxwx79DDMyaTtyDTliCLpgxOIuvjnGkzkofnA3hXh/v19xGQJCOdyndawwTIf
3NdWjnXDMF5G0PHpuFXT3RfXZZVTEJtnMizDoXEehn8UmwLy5/qQj6IjrJY4OvmmenBWbLCGHGAkQCFC9B9k6Bs8uJRy1Ugdgld0
4J4Gzm/b+aUn5mrz0D2/iLyay/cr/CDL+gfgZ31JBwIZv/XypzBqfzUR9ptXsStpIAAODW10HWMBX8j82i/+VBLy3qdtQLcsXn5V
XDhSZAmtHI1R9nntZK+UhX8JR6wzN+cPnoW2s05IOdXRpA9MCe81xo3aiPslzPS3fTztxbSjodp8HuCyZHWDHVNhniR38q/fL9v1
/dfBumF6NkdZNIt2j+vKy7lYJsQ+PAETy4Cp3+mCa9DniK6jxGRuMIQNy9heqcfVWiP1hkdzQJWo/wQP+3J1GO7CjKbw6FFo1TVq
9hPRLr6tEV/7+0k4pb+eZagQhgimpSZiovpuZEneSI/CAcVv8KjultvxgLjteDoM9T57uT7AbO/KSyETWbP9roTJXP9LvFEbgg+W
Ye+KbaYlLSjl0pRJzN4WhSNi+W0r1wOPOjbfBuS/6hfC33DKEsdtvb6WXphEpRr3ntz/RJTlAiVp1EIR4d8mX2MlPyzZiC843zoU
GmbIPkQukoCMr3KAxcsFsEjInntm+atXSUcIc5AiRbdKjYQ99sKtyct2On1SxphcVN1QmhOT22kkR3aPXsalQffYmmbDIGGZC3DE
Wblp1/XPYSK0E8uHHDPDN5BZrq/+fsq3u9T2zV5UuswH6QpxgYUHp3SKZc6feDBwC6PRIYUwfazpAq4Z/TzEm4iEpuE2PJG3UHRe
YlShgHMnDpdzuFsY1Vkb+MlxGEcXdNrBtOYcN1aK/uhyUE50qbcBA6w0qZ3cQb7lJ99/+75FT4BdqjKEGuUeRHcfrBXqM2HswJst
d8JrAhL8V2pCpNb4vxvvRUZpOhYuumkujRgFjWrlAcUyO7Oml3iOGR+zaOGd7l2s5Ikd8JolfZOc/BxFUYRcpmwGufHF7HW5wkhs
TDllPL0jrAlW9PdxL7op0qrL4Rjucq2hXF2fK5QXsHOP+eIXLh3u6N4eO2A3KxS9fyNxH6BltNWoZKjcy3fAq/gkbClCq/GXpkBl
tJxzGQdJldT2HiYpCj4i0Nnhfm+g0ERQrENatGHnfAfMBXbkXFEmpZLIU3mEYuSOiEygaNM7GHtbrpuBsclEtzO8pyK+SoG3Z5pe
0uBNak/Dg8TczYFKWQzRensCOz2vsYHkStM4pWI2Kxl+zP44oeBn9OyZVwDw38tcviLPAX/UcYN15OLHdEqGMt+JUgOaqC1S6Zv6
Voc9lXMi9NPS99x1/0ZlBQktAjbkLmdWpurBClhldNr+UG6GSBkxUJO6jGkVvEE4DFqScds1KKS4dE9wzE/5/sZiJpjos5/Iu+Oh
0AjpK3eO8AZLvyPodnODSiY26xvmiCVNO/iM2IleqZaAmbGW+9qgR8bqIqQu3Tn9Z8pSyJk72Cn+YF04wtp+y0d6nNM+/PgrJgpA
sMvvRVn1qTSW7sr+JGr/9O76btY+nnx5qBsnIE9U1Pm3SXGuYysTfcyk8+Nkjrhul3V+mbJqI9WwtYt6HW+2SCsupvTZxVpbijHV
RJL4UAd1TZBe1Qm4POQnkCXbGrumnDK8vHWgVRrFpVDTD5inm+l9hYY8jIWT4rPPGR2mZopNvy+1QDjF9RlQUF4ZdR2oJ+2+gXQi
R4xrVsmX70ZrBgAesGLfGMs36Dastg9cfJU+c9rNDuBkgDdDwMpmigter0yYLnWX7PHcmylDMFry2EXrtXjsuH4GioociWebLDBe
B13bcQ74DOBP4asx4p1IcQEi1n59kY32WL/wX/Q2xgyhwgXOsAO2AdVq67p5v6Q3EjkP9uiTqh0eEpDsAMSznNmTU5y4FA3C1cE6
5bngHlHMntYRX5S4SY0ELVRW7FAW6LqDeHXZMDQ1fRFq6X5bU6Go38M5JawwG03HULigfoW1i3B02YLaqsH+D7X5nJvfC4bKjOmq
IsvOWlDOBpiaS9W5N14tBEUwQjzsJI6bvHgV7wSbVwlYFfccE4ocKfAMl3tZzrN5mLRvUTXA9KayECoUnoSMVzZm/lVeBotqAtlS
mdg8QRTofn5q8a/BmKimu/fD1HaDT+00VVwWjR/G7QyXfgC0N8oOU2ePJICWQLkcpqc/bienGH7J+mWdzKlx7JpwZHbTmpUaYWSR
mOWl79Gikv1PIphNgQfNYBcMieQY9s2j3oCs0slIE/z6YFA6nxqbrzOg623hAfJZa+11fhd5bJguATC4Kh0hEGVxHPPb+LC+GsdQ
kq4LmkthCjZZr2rvTShlJ9743iP5KlCrfYJXQ9QU0nxOw+LNPcUeZtDZxRaS5GNvrx0d/Lu0cjjpUM+tjjTAeedEnTU/DVTqCQav
DgU64iBKP5SItKBBmCsFF0sMVpFwqBVrWQ4XPYtE77EKm2nIpFNRR5rIJkzS7o6D1ctJnqSO83Nk/aYMBa+d3mXPSRno4afbg+Mr
W+yWpGp42pBZubnkmax/gfRzljNFpuQyHUKJ1d8Ncs0x06WSpjg1DFxnBEUvqFqWDUK2xMAEA43aa6kjg1TnNc2x++EAgufy4byb
XhgWvSP3wB/mEXdeMyosKRTl5VBZePYXbhV6YuZAOKooa1ZMzJ/m3YXhfKcDKzeW9sjr+Tu1OoKbiJ4hYuo48F9+s6l56CJ+x/CP
SlGU/xuHplSpqHb89GBHwYKj6NoJiVYnyJRrZ99CHItBfU5EBoD3ELI68CVvxJ9D8Xtu2j7WTBs19iwSpM8Q+YqOKgqvXLjq83rH
9aZfPlq6OL/mehkDm6sNsIamekgBrsbFhroMhaDGNvSIaQdtB/TECMdQ6ZMAtGxAZJ9aY7P2zpFgXZtM8YmjARwMsay0POGVhLjU
gbN6nf6VcmEbhLVXP4eWxJwDTAh0sOAJpjwsamZLsCY/Uhz9RXl473MNdepzpRmuk7hbH4mehQ/utIVBR847BT3l05kTIhkWxknk
IHZXsANz3iez2rVXtbxQO/DARGq48a3TKRXvMaA7r4NkBpDmP/moYEyGgoLf1UNTYVOh2FRuL7xipJ29RSoVP3Xpv8knLS4e1vAU
SuRwQz86BEYgsEQqHy9CYkW7+e2MdbKRegIF8ujY12zs+Q6E5CFfhMoGp/l54Ei1yqBN5mrSxgJFcnfo7TaAOyWRpunlo2BIZXi1
PLK64oAYiFZJndRg38p41Oa55FFAgj1yPut7A5Nv9qdo+I1I0Z5I8JMkRS3muUHPSsTPzKf6lqXBYCxar2zk4qT5pefYV6crTs9i
1gdat8xcMrp4gjcXUsmJNqNSkphbew5E/YIEnGRc/gmpHLKjHZ0g0aab1H5KhCHVbt2OnhqQbTWoqyi/40IuuC3lCYv2xcgYsJuR
zwZEnDWouC65e5/pamGuejxOBpRzUGdJeU/neU8gqkuJtTV/+LRw+J8eXT26vfSvBCrsUAaRkLwDoBQrYtM/3ns5mKmdELViASTt
PsxIaTXdD5IOaj6Y7Q2mbUpu01Uq/wub+VrMBk8qn/QevW3LPXFAf2tdF5BHI8wxZgNg5KBZXl54ExA7OMSlZUptFPYkdqduExX4
3ShxRU+BfFfTj5xR82Av95miz7FYztLNu1HxGtpN4l6s3PLlHI3pxsB+gtsKMV71vuKwGqiyC6Pib//RW8J4gvP0kh+3mNWcW9s1
0MPBXmb1j71Pk3melRe8wiWdi/UpStdlWBLOH4/Ck+iAtOSGVYkgONcTBChFOUGJR33R4SXImT0XUoU95gbZZ3tRIMQ9fWiPJTik
4CMgRH0jPjISUJ1Rh87GcNlt8wzO0EVjIEGPzwqlh8CfN25coN7snMkF3CRU0lbt9XQsEkD6l5G1IDxkgixl2n6Mr4TQJc6ZQnZk
TpTKpCiHppKYCV19L7Hx5mQEqpOB0nip/pxJoDl7aVminBh/2J7H2+H0fA+bCSIgGWc10Gu6vOQMnQsoVLBnYZSaRtatoxzUlyp9
lnMAQ8nQzigL9SlKBP/JgXLer0cPEpHpnl9dxm4gqr76lBeGFMiLMghLqJJQ2vYqrYJJGwnl80iIz/lQYF/8dMMkM1ZRmGrrbYUN
Fxtom8OkZkfzljiLm91fWo9tiR3+Cpeun5tfn9JRcF6kJA20zuc2EEzlQ55TIkvvrz+W8ALKaIQBL21sC5d3Y2V6f1Cd/oMPHZas
0Pl68CDKB4PTnkhFkl0bIdpEBLxUWbRh037SXJp1feJmM+P9Y2HSNWcPersyROQkHNISRAXy3LCgG6ox2fakN4dUEB/wnTPhIBQw
dIa+B8BmPoiHaEkyPQHI3a0t/42aU2Imm+mqgLR1YK/F62/nyXF6yyvr37DpKmC2W4SdJd2AbudSyJQUStYFecNdokoDSukpgq7I
zV3F9TpuNxmJNPxvxQITT6FSW/QQa0l9rDaCM3RD4tDwLTBfQcYY2KVsnMEZfJmhu/k8tjxQZG+8WzXPyVsrrHeHWSL4GSESpBG0
dEwdLYDSAhPeDt/Cox6yh+sZOzFif2LFaLWfEGgGoRYhYiJh8bvPd6Ls31HG7DGZj6OSiIwcEvAoxPr93t3FUcPiwJoQQxq6EHZL
AHLzmF765j62cYQMUin56yu5jSRg1TuBns01wIS0OmklB/Rf46slq6kKWekaaKPdDUyHxSE5mfgrpLEMr1ILN9ZOyj+3vtmFjiDr
iFcOei2eg5BqHQ1YWuzI8RGg7drf6d1CZeCQJxerIhQtuXkGdXf0ZFXO9hhshreSF0UW2Sxhu7zFjlsxKeYf2iYxqlkz540Rx2tL
4n5rMqunSKff6NGt9+94MUA9FgdPa62GXNAcTj4JEvgUyStT3E0EojrZq5pMLEK0vKWWJMXGYlg2VpdVJLu5r1JJHlmUmzRgyJ+m
4OmE+pKZSk0Estl0MtN8yDOu7o93SrvI3n2j3flxExWL4QHhvJRhz9LQ2lj0aDenUQbFqUIdvCOA8t4Xx00GRcq+yGl5lAw358nw
0cLkP+PlllIvSWmYwhIddhbhOaJi1zaHYhopHqwqTFvZ8EbtzLzz1FROjU9mwNZU10MaHotphIxzldWykJdZ6H6EawBVfwbIE2vM
OvZoRztdQLL4+BGTf9v99AGx4VIOQV0oD+ropSv8OgHQEu8eVnIqa6vbinLjIOw15yKMVPc/E+wHeRiRb1iF3i+6qA1hlO5W8K+e
OSxKWn0c1uQyl/aIHfRN2CvaS/OOQH2dzBNaK6KRQUdnjpol30FHxC3LcSPeGWQTprLBmkPhw4TYjf8cpdONQOV1euPtSjtlaVlP
aUQ8+VAj2Kn8hLhcd0UXcF9PARguDu2K7W/QicgHQtswnthUvbt2zyylZHGytEX9tn1a0x8XBXV3+XqZC8bGfGdMN14QJmDP4XST
eCYSO97AFMnfRelEIrAJ7zYTT6hXEhqlfowA7F9XBrWtGfT7MU3Cvxv4t36lsqycnW7ITz15mFHEOT7A7aDD4o3f/RLPFc47i8na
yxdPhtvnnFC/fCEsWvYKp6z6D95o798QlNZIaBDTlt+qI1WKWWGigNrXVqkAAVMsUd/gzqARVS7jhP/LM6m1mFDpBV632jfn9/ap
pVPK8SV5p5AU0TbiJa6D6M4fIH1fSX5m57dNKsmaO04mDn8hvYz1/ZUpmDhx0NkgyJuIsKh9ERi90Cma1BQ1SWmD4FonRHc9WXHM
W2XUuDTfIifw2iiUhIxw+zAUg8Qoaqc8cVTKd0cc+zrL7ydXuH7PVxPOcLW22YCCm7QeiHmWhG5LzRSELyA3k7dJ5Da/63067HwY
DVitC9UG137Ccf1qpoQUHXwYxIVXnGcCgsWikk62sNy3ucsEsIajCtX1LcOkvOAG9H1mIuVL2HK1btHPeo+242W93Am31KpzCeUw
/8+WuC4Vf/d0PMyl1kRj12c5FNPFb0Cj+PTDCvo5X9/cvR/2Kk89DhxkPIHQ2f2xAKofSylgVmZwjvRAgwwNc/MI6t44cWTdjwo4
H+Wt/QniqiPLfS1J11bBEGTxZDP5YsD3TC4bNgHhQ+4glab3TljEJQBGALQW7AhzNOlSiVxG/0ZLdTX04pCt4T0xom4dXf3U83GF
UgoHWmCKFX/9D7lwZvGVoXA0aPqY5fQGYpYhiwkIUIepV/IqGq9AsIB4YNXzYKFawt5qvtOc6qWbmqqXIKwNtiDiPaE9zc9fw+70
OZNqye1UJ9zAd/qrgh1tTzuPaJUhoxKfzCBgxqYr0Z1EdqJTLcxhNj4yJXXDMNgOYW6Rkuc3UW2hGCraoRX+B/rYfi4QhQ95iuVU
SwOZ0Spc/Wf/kFH3Q4u0R6wdsjpycee22Nd/HC163bzgEIpg949cTOyojrLe2LGNj14lL1/qz9V6dnYWyAOLZ949ZPZm/Y3XPdNn
f7v+K6mDXXspkpvCOWzb2jGzmfvXd2hYwu7xWBfgHWQpA0rA9/RWMsqwWQoEHmTe+EppsJVyKHpta7YuSMannwYk5cFQ9PesSFZI
iqZMGmesWAlAEejlr2AfcrU5ssuc1fMZK0D40gxzgXjGr9Jk4JW9s16To+iydsWZ98Wzi3+YTfsvJpM09DI5m72MOfVE1cA7m8Ex
ilA74GI8wsi4A6CPH0BFggeJOlH6oyjIBzeFrhxawofWZqqKvNCRIhsocKRjr9G0u3E7dPzMothCZ9jEb/N2jn24JP10X2M3jGsR
qi537FXKtmMySesJrAq+VxY3jZ+GCZ4IvnzYLg8DP/g60HMRKuE/tAAAzfmkLeLnypC3VqOhDxwsVotM3aiAlQQtA7aCaV5PMght
1CMLxkAEVv0mUCradvTziZWbwflTlGF5A/RzNUTKHXe3O/50RuCGV9pnxW+2GKXtGKVjxNPOV3FTG637vqpi67Pp7M3a/QFEU3iT
pFNXZZpbNImVt5Doo2fEo1NfYBdwbV0JZ6HXfPICXMbVQclXT4sfcOFJwVpglxO2ZQg2B4DFObpVZmBll49G3G+w5gvxfzY2SYIy
gixldDjI2YMDuz63kSKDPcb8zSo3VDVwfcm8ILMn+9u3JeGx1OseuePTbKzhj9wFfp+hudoezEvv24n6QYxaPOSPS0XjBAxeO3hQ
vy0mask4r4qexI95AhMNTMt0+hlmIe9qGk/PZeIXzd8mi68TDmAZZ5Z/GsNo2XRmcaQpEah7VZBlcw9M6xUukDNzJrDvpfN99QlG
uwCo2Tm8fFGxGXi0AFclRoELOMoMWKVFzAgXGNotLxCiFiepa68/J0IOl9GF2Vp3jJNvXm3fwCOE5qFRH98mXiq/6zvUkgMLQWPo
SnvV7qyJS5U+/rrIZ6Mextfczte23a6HWUjN4n/VhXYg/OSnHkORSkPKuVippFnO0rF0BgeNm1itTcdKwh1ag/gV7wKnoWRWCOZb
HiVZnPpv+lw8OMuK3lhyjOqxJlOnjp+zGwbiAkY18bP7v1rSVcrEBlBdE7GakSgHDnj1ohtAZZ97iKTZIRVZH9aOhPllZffCyd67
o1VnwEVG6tQpO8H23Ujc7RckfU4GnuRHy4cZAZYjFhBYnKNtZ23+K4ggRUWhTg23kGWPuyTgtM4sXA0rWyhpNyDou+XbJxpTbBWr
FDWq48iU6JbdQTOHWLSyJe/Wfk7q9ssnfta9kvywvE5SD0Ia3RzV4n+xQWX4cQoAw8DTThdjbnRxn9h/FAzrvpr9IUHU15UDWJqT
TVKEjxRF43iVB/AFUW2DJ7Pr9yXx0HUZJ+QkYTlMKkuHwMNvi48uA5jlr7cn5r3z9z2xjelBQGD+I6RlR2g3tA4/qPF7JguNfPrX
xaT2xikVhBPAPWwz6FeYbVZHQgef8RpDrt/M8nQXb20504TRF6DgZtfHsuKTLGNY2+xLZz4250A5la0UHxlz86izEDbLhDukQV/8
N8Ume4G30CMudmAAS+QuBm1rQDy7twgPfjUvpfd6SMBHZ8w4P0g7bmRizNuGmiVF/5ZUwv7rVZIa5aP4iTA5k6yR84JFZMLjhKOS
e0sY9OdOBUu609LZ/ExV8BQWRxm66oxZyYJCrtNomKDbx9aYcEK25+SI4GuccKthLYTLnas63k1TIrhqqKN8AWUoClQrKGn/yV9p
nspgtKSyjy79jClSCVVxOc7Yw1LvX1ACwyjB/XbSc6y3ddEauL157xJc7i2bhi63Lhe+cNdmRJQCHNfpuiKqnx+hXZsSQlmsCIXK
Gp+5zlFL9NpnidQ2xHsI9A1jokoTEKgwkulsJuf+th2kDv3OnDa/pS0dxM2Ieydj7KVxCaIIRJtx+PMcE294GjBLDx/QK7Vusjrr
GEio9f4UI16qA9VYnZCRWP5eGrm+xE//0e35zdB+QrqS3hjiqnGj4RIcsNM8ERO6mEsLbB1UWsnw+KsDl7CIpDWRqUSfYSkWj8Nm
P37NR9vcJgkLJ5mgCgDR4YnNHbirGDiUj4C8QsQDAbxyrjnSlXi28W4GTenQzr98u1RUAahVfvpeR1QNWciOBuULUkYDZIM6u826
DhXAGNZUB9KROciBotzey3lLfd4BC/SbboTL0Uue/4F4bm6ougEamyT7BMj2NfpBKVb5ezJhyB/q1sXfeXTv0G4EKqvX8RD0SFzN
btUdL+twx8898WA3k0NE7BgO3ywYJBAZyJ1MNIITD4t26XswfGbJQPddqKrn1ZAOnKIuUa8s1MTbDsOHsVDI67uw82SaRzKCrQZN
YCqFb795+S+nYhS9HBadUju/GOmUpy4MKcqilrHSI2GUnD8gon8LNL0rNqCOevHK3sXev/UuFUBVxEetryyfXgX28AfyXS+Gigw8
13YJLy5GkVCG0dNaM8cLT65a5dZ+xEzHh5lqcDpW6N9hxqRwlTBkdFgDHE9MBWfjHmitrwPT0egphGlWWD8uXJfyic9QgiLjeO4I
vTu2+OkBjgvLmsx0SYmlSaWL55t2y/2F3+BGMC7gzgS8jIRhqXwr+VLw8GH8zOkiqRNg1w5JCl+YyjdkWoyTfO7cgv4WmuFUgM0z
9ryhpx03NVf4uJi3uyctxbtF2AOhqRXyfdooD+lzZqgLVMNdJYyIHevy4vQ6gYvBPGAuGpLr8ccfP35cHnP9h72oI9xLnJQFlmnY
vWaq1vAWN3BiXTLDzir2wtjNeAVyUTCZX8V82ybBtC2oI9y5BaZnD80UC/iiAvJ843ScTztoU7pRJRcj+97BVYMd5J5O+t04v7Ca
l2rGoZZToDW0fGvsQqbDODIfYG9UnsjzyV1lK8a5j4pP9aKxu1kPhKZIaMeXnojAjK4k2cM91ucodnJWiwHLzovuIrxQXBO1GMAD
vD4pjFgj6/h+754YvVTjEnBDbWwjHRK874PUmC2NpxcSHwLvKTJ8fFLsvzovtA65TyWMSWlIMLZVeCumqa0vR8BnR0LPoO/eOu7F
/gTPgYLNJA1cgo+UyjBR5SY2WjTLGBgANMrW1bEb/Y+5XFm7deOi/45opcgFjY7QnhGtMiDVx0Ooqct++qu0GdAxNiNrk1WUfwm2
71Ayi6YeNtTw+8mvY0wSVeaZN4U/KMJ6httuzpKkMzhroyh7hKYjNatr8VRIPlFmUEBE2lF2QGP5Y+LycOc6TG/NfTWKiVSdOX5c
0pnj/gL3UAnuOUanzsjqntdBQh0AeyL0RmYU4L1lcH0kIEWqvI4WMkEDUMnxxiO8YQLACspExiWD/7W3xtauEHpw95N3LeTJDQB6
cjOaJ9sNtCBtQ24lfJ/06HVUCW7AKOBsi3jEBIBF/qyw+vVue81/r77DsXuF/XIOkA5XAt524a/DHG2s2ZXBHl2RBR1Z5nk7+cGO
m5aPzpHjr02WKwfdgw+3k5UYxkIh980sA2NacZW2lIKw9O0Axe7ab6EFjeWA945ihXe9bThkBKJUa1u5kWYrs8DocJv6qnafCGcT
QBEmd7ozOqlhy8geqHJ0/1eZaL3O7YrPmCr5c/c+KQVhVhhY5qQ4zWwwU+5HcekU+dVDw4jKi6i7WixfTS3icaiofpOhTd6ohUtG
dPa9Xx2AhglFRYGks4yhGq+F1meRkitmRUOmWdwf/ho2KIPX0yXENEft0Tzv+pxKcVxUiY+iJ/P0J34Bqaeap75H+oPMXtExITVu
NFbnP/fiDdihei/gTxy0wb4Vbvjpb26AmnD5jVApCoMvVY0CBTMmIU6FATIZ0Llnj+llTO5xticborNS88k8dXx5Ig2RyBMTYuNg
pb1z7gdeDavz3ybAUJW2ffvIji/5R076fU/2oVZhN9iwidbxfxeydkZTAph0SlSt9nORJI7kYYt7EkTSpbwZhUKM8k2Vj0h0z8ub
ROwy/B6s/adtjIU/LORGX5/qfeJasUXzClSqUcnZ6CEAoJ6oLE6Bd+udep9zQ4H44t5d1+WwyMAwM00I7F9CAhJuBM7ypkeeHa44
FiipQJ1oUsOhm4er57JKnoUFCew/M0TOVvZP3KVs9QdCKZklK0+hU0vhNcgjGEH2Moa4x4MsosxFnn6WENOnPdUplJ8elGgD8Cph
XKlhygsyGQZQBZ9ntA4RrrWjDBlj9jXnz53gCLb66aMF5UFSuuqH7PTXPcB2aWulliHmGDjYYC6S9vcLHIM+u8AJyzPjGzoPt1P/
YtvSftcQF+YCaiwj5vFFzRECVWLJkTk/Tw48piS6YLqMs1SrX9xp8HlKs9KEpbvn3Gte4Z+GzzxA930UHI4AAs+61uyX8kZzBeUb
Yo5J948jnTyH0lSJ60T0rXgZOjhDTRT4TQdYbR6XeaCIEciBBOrNayxkiX/ZrXmp+llgkgjMwpgGB2JG+2rpvKl65eLQYNr9If0e
Qbs8EYkqXy88TZEvQwJ9hWlty+2SH0QhhDbUFh6SKBhbAsYlbvoCpewgtWHXJfrOYruqs21AapFGdjixZDs0vbveIdEqWUm9w53e
+uzPYKNIh//Gx06yGd8obAQZr3rZIWOjomi2ZTfv6RskXFaUPkLwhyhofxAeJdWB5wz4YFjwXlZTTIbHavuZwO3g5u2kByYhpxop
S8wA4x1lx+wwObeRzPU2+f8HrtoJkmKZlrmx6uq8KlZUYViGosWpgzLUyIYb9KEfGuI1DsQCw4hFYf213cW6us2J+ScHDXom2Rdz
RXKLmKrW428tEjZpSTTbei3ipjZe6dLqmTOQi1AdjgLKF5vn5QWOAjLPhFb0XSzMbQfeW5NZ7PaFZsf6DcHCr+oxpXDeTiheV/VI
nZhyF5D3PWKflNqX+/mb798TEw5s7KVCJ6xlo1tKq6Ac30MQezyWLtdVLUAArd/0uWrz9rZ6OTI9YtaVVJYYt1YFUQJw8vco3Miu
SC9ikZK5uSXeoPQipL0fHl2PKw/pF27jlvEUPE57V7yn7Dq1QIIZMhUN97EDx0SO0RBLL4Fub1MCJiua5DsOazTz7sixcEujmGaN
JSst+UnmWA14fI9Dwxvcs9M3/yS71fHJHoJh9LjiYWpe78G/xyLhfCD160EFt+UfQWyracoehvQtKG95eOMKuG6o15OUcn9f6eKk
L8PjwUt0Z18XrbX7icPCUuG7oEYhpTNC3WF5m0grs5MJHsJJLQHCKbfdmx0URmzHcVX2OgOxZ6ZpkLJyDtsFu07uEpaD6+4RMMgM
EODf4C2Whw3hSH8Sq1Fpk8Vz7PEI7LDkELEkf5Uu/D7Df0hsvux6gRA3ZVqj1Jn2nPwIWvbYAfUKtTw16pSbQ3LIH7KLnGcLK1FS
qCWK8j+Ohe5f27PSFai3bG1BJkc29PEATrwrspU1C5edsoo+T3Lj6Vphu6TWhBjV2xSycimyhkp3yJA0cAIBtkCPmsRc94IMnkjZ
xBiR9/bgYFtFHI8YeYgDf7BrRZJ0gQyUanrn2mG5suECIDghlyhX5dGcZYKkZ4AkiM7RsuD7/mWP+zM5ZSZvRLDwNfdnp7/v8ETG
l0H9WObZT/dMrk2VrDnPeUeermiHtwKjdAQkWVghEKCP6PNhwR2eCP99epGL0+mNHkeg4CooukUv7FVMi7XduFb95GBrFg/atykF
aMlSOu3QQlb+2iWPilZvqQXIRaSiCL3ph3BsHuPL7Sr1Ugo1AKhhVhabNXYWXZJ4yD2yy29S9Cgf/RXWA3OhytloIHwt0ryoNBZE
q3RjXy8j0sesHTb65eiAnkAlJhTj2Ej0LPGi9unUAWd3PWJOO0aEFlq4U0K7erE7A0Sj8NHDlhQ7auUaOZKq7tYCwKXGKRCYJQbs
MtB+rQK+9jxF4QLBLybENnucwWHD++XVzKVvEeyH7yZLdxFxRVofTyGniivIwpAutXozMGj82bE7Qe7AnINrFXD7ImAIghRy81eG
QjzlsvLbsOMII2/2e5Ua8rQt2eE31aa46aUGhruVdY3a71ZYtFMwLDaUntzgZ0HDfUmnVUeVSnnsHfi+71/+dK1Y7yhe2PIHKvcz
cGk3yCB9s3Ow2B+XyLBod4B65Kwey7goGxccYSIzUzUgJ96iKHrzyiVGrjCuH5pLDwyrkheQv/qlwTLA10So55V7WOqF1xf2Jer3
BZwguwi9Kkb389RoeDelK6RxzdVRWmfomuzNToYOMcTia9wecs59ejPboyNB1w8ujkdnu0D1ZDWnig9q3W7nzsEcOGwNKCDeu7dv
qn6Ty6rOkvuPYQE0AcSDxO5Y5lO1XaWRBFmaMXiaJLvpBwnxzwyfgRbLYG84sFMA8lFqjTonqx0w033IWG5sMFCnUgEtdM3/ofCm
+701/ZvGjZsDBFxVjpqe2U63oLjkSbouGHQEX3jBa04ol9uX/Njf4ugeXkcPp1S89F8WzmqXR6RRv1pFrHRbXmFPAdDZM9BZWhVQ
IYyDg10NNxMJN7Iu3TitJYuVdDMCj80WTYBOqlJ9/7+IOUHcRxOVhNBZcfZzPlKaAP+4gLBP1negSsmrIv7xGM1Ev7r9CHh43F13
EsFmOa8R2La8BS6FHf8gEaGUuLE6hqsZFcCnfI1+CITLngVHJC0SYKx8HOX9q87b+ljdA9AMonAAAAV5ciCmjq9xrtqpJ3NUcR2l
7UQcQCPO04yTbiRhuMM36NlC4P4ldCFwVa6eqNj7dYbfP1TfSk3KK1WyghxmX5IUloXUgdTWMxGOCroC3peOKdwRbdEjbBJlaSCl
ScCoaJCvCPh54UFn774q3FUytAw/q1rXB3Olobbk7SVmP2ZUmvE4m4pxRBaeKnLdMhnYYzH5pgPoqeGrqEK3h+FhXfVOdbtzy6lh
MMXHyX2BZMWRlf9ezsNmvjR5uKkWjsLC3Pl0NZrTj4b5kGUoM1PrvxOoGe63uf1YUDWxEPrRY+5RUTNYMoxu17RmIzOkaCEE4g21
yN84qNSDTfjjgaE08BOJ6TuQoBfpO80H2motd9g5WrBOH1eEQe3Dfhw1KPQJrtQwGV9jEfhMdnJhVlGWvjPgTX5PtsMq+dTu91VR
H7DcVjgPhuoCKBu7QPl3mgrpwNszwGFf7ksBEyhb1EHiOoWQIIKWCvmEDTWeSIrw5AAdqePNxMSS0jk7kOjUMx2QZuPluuxhEtNq
w5eowxD/dfvqvyJgEdm4agNNsjI3QlY3nA3wesK+h1kxYWudMtgyuUYOaD9KsBfdGXYZRUYxQgTd301TggUwF3E52JKarVucFFL7
AXU+RNHJ/r3ClW5MtEDzDS87YDDyfIrPdISvm/VNdnjmwyOzmPDhRmkUioS+wF+o0UJ/Blao5KGkB4A5aKQiauPv616vnwIzqeEj
5u9ou4SZxT4l1Cz31bywFMrS+pplNSiRvSo070GvSU6Ng5/xD9ri+sIOseoeQiF9dyXxbg6f91CaSAmk+ytJHq0RIrzUxutfNFeW
WfpvOGF30T92jnoYdsR2FqAq9gvtLxcSmqE9SinPAW/QjP7R8iToGYXSKRzhLy4SK4TWSz2k17/YK642M3Z9kfra4NqtI+yFVbOe
WazFazsA1gE4zkmm/+1cnjdmL2gh+/DlAZMXPVtwuNJzuPhQRoeGjtdZWUtjnx9XLCmw9s9KA28LtOjgNUuF5Txe48QMJseN779J
1EuQaNPkp8VS/7ooEIe9i7oxRO8vbTi2+278MK3/gOVKJbZX29NYM8IgLwiqTpgbMR4QHxf10fQZ9zNIcmJyFiaRosuivU1pj9V3
z9//yeuqJEp2xDldi8y8IRb2yLrEIOl+tVzSwKCuN9TrGeD2TAO2HHuUmY7qeoiSKSltoS8Lznk/7f/gB+K6j2md5AvsLzg1gKQk
mCyZQMmD/4tqn22+CVBn5odWJjplW1wtcg1NjadxkYEnanjSuyxEwCeZYZs7L0dPsIap/cVv7K2inEIwLTKgmghQqBuftiR80pe4
aAI5i3y3asFxD0ktJQD//qx44K7CS3glL/bjVqyCjNM0GoQprp0UV6AwZ9yf33iA5SLVTcZii3gG6cl2mmwJ3twNgyA1d8CYZ4CM
L3z2ZM0yei5yRU10AkClylzksi3bcqw6JxfvwNYU+ulC3RGv1TLEP2xxDCSvnRwctigM5eMoc6qGoPfnStPbTulivFVQxfV2t9/M
HkHng+RCllzgIqYvp/UbwM372vXCNNgV0NtsHrkSNn9G/0hpw7NSEGDNSq1YSOQoiW/8EiEl1WwE8eq2CXV6ebFrn/13jwSjjDgX
xUk/eAiUP7IeZta3C++1HEQ1M+NA13axWSOu5ejo48w6I66p4n+TFwpQiIyiKKRHaJIqgdUuPGv94ZfZw1hBAMRrUBeyi2ij1Dvk
M2FhOb3SKoprWjTMuCI7zOqvNl/yqV+xzZRMshPecyICrOJGSXwzhDOPIs1wo1d5gl/W4an4HuCjsAM7cbgs086zthC7uBet7CfM
YWtf3Y3MU+mKznGKSlhXF6Y3bMW+W2x8ajcq6PGvnn7kPIDCG9o/yPmr8YWDgXbGpB4SrGcL5M8yJL74d+eYbmWuTkbO0J0XmHqP
lI8cwmFnqw0nMHIDlVQgsFaPvIfAr32J+oNXR4qu3FNGYXRqV5K+D5Az7cDxEku+GG5E6x9wJmomPGtmAVgkW9sReZBArsOZdUNl
BHARZMvyuFmSzvQ8Jw0eRGy5ZqjMGit7fiPRCzdhhG4FRW+YBBStoibhxkJzxVcP/oTA3y/LDz3aNimerOlmcnIraaiMvz6+8M1L
2BzRqVU6QhoPBuyuFl7IucDdmR7dB1HkLYZQuP7g8XknvZrbGB/WOB6l27s7nCqTPImfZZ5vgz+oHQAT4mq7Bq6IXM784SIrB9jX
tP0Xp3Fa3KPsRXcbTb8kNkOQhqFiQe/tbtPJ3AJLlB80AL5uxut1SM+yrTa8gmRHsTOZkChHsq0jV4KPy3KmkH7bS5n3MvKR+XdO
8EYSbQhv73823J0RXzDcqxofTYw9KjL0zaCnLcENr7tVrP8WtuFPtjn0wUqQPPwwPceqFDMAYsE5HbOwROhplvP2e0/OOoH7inGE
uPN2YA9oK0/v7qt5uPbAmymDx32cTIj8MrMJ+U/ct6mScdr0sUg2xbzpCpwDIuLYzTgiqfq+svtQvXL9k38UD/5m7dkW60g7cN9K
rbgamLNrGuntldYHOm5CJYNletMpB+prH7J2sW7KVSJZHDCC6JCIo2Oq76bXnwDa5JXY2DiGu2qTVOScTXJuKozV9LLiXiJV3/ZI
lRU3H87xBJGWrUSYutfY2gh2dl1dLhNrIsGK/FHM2lZfkWJMco3fsnOJseAxXdRGHMOZ+2/q9qvVwXgqgSmgOhW5n9h1rqxhwkyd
2TDYHGh4rhht3SUBmNUMCuhwR17psV1W5qRaBpByC0clAwn09IFM+vYoSbbEdPtrdDZnG4Yu2OBUruqF64rgMRnW5tQHrft8/uOi
2fw0I0nbdpDDHkTnA0SN3uzIgw+cvI+eSgfYXTctO39KElrTQvgoZCtcTCvMiQ2fYnoySDK+rYIfnOILovTcQ/LqTW7BPbssNyvd
4T4VdCTcMKA2tKd+MfSgw8VfjSuffmaAPFShzdxwjrj2DAz3lbjDaRTOsJjptEMPl3bj+nTPu0o/gn0ZMl+2SBeKuWAxrHlanozj
4WtG+c3Ov9PmhmQ4V1hSrhOzATWJomlIFK8v5F0VQyHoP0XlndMeCHvqtV9avCHjpoR9QLAsHA0ffLupK6xVlSKc1nWTL7N4l3u6
xt86QVH/g0K8qCWtmWvx4DbqlCpLF0lvo6mkGzL/ECrTkZUseqI/W+0z1p1Ghdr7zlvoHEq0jH6Ck6QgZhWM3lNz9XO/zH0Eq59m
pCM8TLv/dMQQ9Z0+XnKSAxyUf3hVMwv+VHkxh11hvtZ18fs75EB+t1ehsngu9MNHNTQAeg5wOcB4yNyFjILYo+gqryJpidvO+liP
RH0a2e4xxQotjOqWJJAiKsbczkD6icrQndn/Nd8DqBQ8WKSOAaz0ODIx2mgknIXwq905UQVEoL13OV/u4VVmPYNHpxm6u8GS3XAr
281LJex3iEEMwVOo0vWBzUiLiOzxfQ6WmcvlDaomSF0exs34K9lNFBQDjDSnylq5oHvShJ8IM21H+5VW3tKGaeoKkeGRQRm26FWj
ShB7xVfqUwH0Nrzxd9E7BJzwv4rMkuCyru5nkBoUm84sjE5taKkvmHeFWP8Y/hNFslYVvcKFAPTVpY4cOhlmwTOesWjqXdrdxBGt
rQfyRhZQ+jGWCXN/Aa7fER+b/flekENb0DvzSNtB3TSLB2Ko9s8PoEi9SZOI3VAG3ED8ykTxlGaSXgoDKBHRN2qezLXPLvvmyNzE
7otU5sMnjLGH8CWYe0zxiUWsBGgv6e3+XtUmi3wFE4gCcfxJbrG+Mgt6/GWWYxZ1JxPDiHmHirweJuyf5xbwqNJowz4ZudZmIQbU
vvQ0PN59d5BBgeQdhYALUi15bc8Ddmqy7P+Gn3kL9IIEbjulC09bCm7rL9Lm1SGgrWVtzXY76gFZ87R2lc1W2G9Tx+xtpuX/eDDA
zrB3B14zSJl8iUWicRKOIASO6ccV3aK64LvayL/rvTOSdCus7ik7P9vqpg5/bXXKXGbzu4gjthspeY3nWpqdctGPYQhy4zkQau3k
MiUfHys8NZq883nIO4Gji6SzoJSklY4qIE1C7J9MC13AGTufU9fuQRcmEd3NRBqxcHDjO1YlZNjzExTQeSzh3Oc2BHbvDH4ScNMH
iqSkR3MwOf1302LXv33NKDoaLkSYuTB5pykGjV988OkvAwPupt2td9ia2LiA7rlJbwK9LTeVxoz3e/yuSr2nxiFtCeHFdZTdbuLa
ccragSrhdhmA2rtQfQdQOS4ItwgIrR1rhyQF/3j6FfTHGOI1/MdSNpXRpi+/cp78RGIpEFavwW6OJXPUKHwRdg1hJx2lGNXAfhBo
Shzc4bNcQGz2XeS74BHwdsBviyRwxjKMM/m/mifUCGKmb5oQnkOHvPMHPbK39kGGCi6/xknLIVa8Xi1w2qJJL65q5xguG9QQhvDs
uGUCAOIf5SDbUUCFg/VztwknZIyD309Ib7B8n8MfhtzqfryCOc6F4KAei/UnrmpIZFfE7wyWINKMGpZj6jrPD4S7WSwijxLFugc/
JpLFrmIn4dkt005DpSaowO32IYzaegQHLTvx6eCWs6C9SnwsFp7xQMuXp2yV0ISE0ClWllv/PY/GCpnUDrUnj2rhDvfAG79yqXdK
dGQ2PcAD5rub/d9vdZxtQOiS5/KHLD9YEKmg85ptz3Qkm5urp2xp4v/IH3p2EYK0I0axaHfhTEAF8RFZgRtLRTGNiyPgGaQc9/JR
LnaofLIPWUP4OcjUlz4QRc6hHw/ZAM6UjaDphWbUCznjAde7hORPVPMGj7bRPGBfSSQvCZm1Vhgo7k+0vTODYAhGh4G8yShVgHHz
n33LMKqXw8N1UjubwYjshoFQZkCifYJ2MX+bYlryCCJuXM4LlxcNij0upt0gcffNybabD29e5IccPG+1pz4i7gtwmu1XmnNIDPHR
i0LN2o/UsdMcY9dovI2FOad808H1Yu+TvYMbI22zPJ9HCFXykcTfPUQXKk5U9YpUffr9R6SD9ejTD2PlZCBPSj6TmqIOPH6O8uky
/AITBPcHn2DqD99l/CVAC+AG9NgpPZlwAVTJdVmdVOxCiibb0eE1/vGOocGlmaEexjg0lUP/BDUTaAs3G1OTSGO2LgsSk5NoIwBF
Wmb7OiEwcZEa1mxZxA5bAZOi16xBJjANCW/hWKTk2ueLuJIBj7F1mjYmnwBvoo6v8i3eWYanM9wwewHgX4quT/3eYKXu1rYariaz
PkJlwTCQhy7ZM6rof2BT2q9zEPPpa3VItAGHRUMT3AkXAcIZlvMFeB8nGddWvA/7x5quch7JclMuRwjDePazSDuPkUa/xpufgPE2
prJwyvQNBCZAHxZgcRbnxU71O5EV4cxou3paIoP9H/rCgvfQ8qc0G0I/S+fZ3IXRygAHMclgNJMBxgqqwLkIWNjTCbevID9p6s5f
OgZY9lVsmxg3dIeQWJDyOfsDONTCwq7UAemNrFv9AwLAWk1BcbPe61h5CqULVVAl8qYi7WZjphcVxNDVqS67lyJp9N32ximpie0O
nXNNzddDTUYPEt/0ljmX1kcJqeoB5yzjgez2mg3KvFIjqcc8LAhBvz4SSgtWgYDYHHgfNm9eVbh59/394pLYxrdySjKN31e3XRd9
TTkKHWYvgRTaPLhjPnYDTyB1IY6NAILKBE/ELyKcqTOsZMrqI6Ql9XhDT82UVp2hSIfN+bhhxMGF6xhg+RJMLYMBqPsnsNWbOy4U
Qb8hCBqp3BtkaT+UftIPe2RZHsse0x/b1y1LLVsEv5Hvj+DUBU2i3+JgkPJXNPasEOGP/vvSNCj8cXc8ub455PUww1DBBbt+5uXN
CJCS1fET0Bp1EEJORVDztAGNtSQ9ezkaOUz2QIL9cOQd7iPrQqR180CtELQYoptu9YwXD5SYCTiXK/B7hXgiBsDM4Nm744CQvNnC
RXe1KGcnnq4MNBdXiJH1yylIPDYwuvRbjqaz07/iN7+Aur456NH6oZ3eg5rwQfvpzAv/DqpTLtFTVau7HPO5P97AgHNPIRWf3j9i
Rl0scsC2lfYIwBGcc2kLsuCjLmmBk39lmBp+6MwpyEiUTNlTXbfXEMBiJyTlsZcb0T/tbA7CySj1nS2ZeF0mofStLrs3Gxj1Dgfx
B4pj05AlnNUQSedzFwGaVfZgTnG9zPpZXscuYvNTUv28Migo+M7e3oQdgBLOSd80/h7CjM54JL2stQykzOElnopgxQCrhkZTjY2s
wR0K2LcemrhC54KRyV22T2gUzezCysx99ZLxdSC80vfCgK2tDF3p2j0VS2mig0jpZMrlD9IYyEAr7qsuVkXs3jEOASD87JMJwQXY
fLG/u2blGo533OXmlTDAboTaGqMRD03GowcBbwMtkoZ2jqgh0MsPd3yEp0cved/0z+TOLgKxS+/aEi8zowaVidbolQCteF74C+Hn
nDBXjHe9QvaD88B0SiwkJdhkoNKNkR6Z3wKFXfC930AS01IqDE7rEzmjZrIOJa9uEK7AdGpgpHp6Ah7tfvyGLwBlcPdc6L5gu0AI
T7oyXYmfPTKtvGQiXiZXvUig8Lxs3W3hIWHmS+nDN3t3T9SGq90Osu+Bd1VktrrcwTjSRg3B8PH1Q1GNZbzAWuwd5QmfNHZPoOyc
0n4EAdWTyfdk7kqjdaxlBip6PAPlVGaClJcb3tPhVtt2mP4Gy7ndfvCXD2JfEGjRfudGXh4Pg3tcnatx0i0PV4RrPyjXGf+aBnf+
qAdw1PG4CMJgrhwKyDFQQO4BMOfIeNjrbIsY5sowvx5KrRNFxI963rD4hLglPtO3D5be3m8bSpf7eOLr0SrrHeMzPVi3dKvhbvYO
NtqNS08IiirPBm5e/NIeZQ6bUVPknzU8LvZK27t5fbhfDotvq+9nPp6B6vlFSPRGm4SntD2SSc58cCpGQVikkTlQvhU/EGe6r4Qt
iuk/eAJAqsuYiDRsOHS1i+dfTXY60M+bajX/o8sN3FYOPrOn8nJyktdyytH4L/RU8ShnE8gDUiYwqteKT7Ov9cSB1h7EOrF8zIFO
faywza4TEJIYK+1Xxb0PjRODnbdGnJd4WVpekOGCeRQvd1Kw/kD2j+L/EJMxSuad2Y5yRftGG12B6QDpMsFUsWUKkm2FaIuCSGWC
IMw8b0NV4p6GxSqEBPUsY1bOes0kYtDuEC7CmgoA3arZUwBcD6oOK7uwAoXp/qPyFQLPZj3e4hQvSz/IrlO/X4jN+3/3w9DNGWgr
Wcx6IIK3pC3w1CW9ilcSIRt/XUmfKFu5p9gtnmkInbyi8BAsiMFWi/hKmCMQzAuK1zfHaZrFuomRtSFVKqs1D+NNdV7J4S/w60jz
6RaYP6y/lpsAptnBAwsdaUGHM5rqNOC8eGsiXPjzk1RfIe58O20M6F5mo7XJaTBGrnGtwQLKFEn5BMEC79Uv85y2NTtLmNTPk+0k
pyU7Zf5pyU37VodyMjcFEnEA4ce8Jh5Kkz3THCCGofQo5a5PMIo49JI9NjPKxfKwHzGMbwv2eHhpQmIfFWUQNzCnEZnYUTK4ZicR
eRyHpeBSSB7YCRJaZKnxkAXjTZ0A5lzCWIjhtzGTebiAD65MdJ6k9SPIwJQzZTMe4aoxlfHc8gQX0hZ2Nd/A/oEYyiMtlelt+cun
iCEpNYeSbt9GH/qpfOSI0tg1IHL3otV7B0GJTBewgSU5S+R87n55pxxqSS7wZh1SkggYvD2DeJK8bb2YRBshsJG1E5P/Ev1hyTCF
pNQ9pnGpKSyQUA0vHhDzL3fpKP9y8ylsjr+tstMR4kvAKK5obCeW+UFMjwGgMhpEMI1TQtLsFfnPCgxI0Q/q+qcNY+3/ErJO/uOJ
jg0vCQ+OEtJWGwex154uITlojYCo/Ee7HrbAgXlSUL2ojxbo2CdIzpxJICqKvdK90WZsskG24MynEEa47cBCjvvFMLFbq6Splxbz
yRdUOyVS2R011XVrmbk2Lxsqz6JPSNkEHxdLqY1QorRLdg8Uh1J6uOcGAT3leLEcS4dQPQseSgL9mxYPQ7LxwzOLxAtpVAMIHgPK
AMb1cjIPMiuypwezthU/elU1qbPWP8t0ivg9LM/4GErtFVOnQyesdILcbo7w5leur58hxaLOvaje9uz4jTmAG6U/xHlI+j6XWjnq
4k6xIeJlejHBW2uk2p5KnGt7cnQlckjGEDhYQCTCM8b/KsBx5CdPZfW4xQKxyGBmZyYThaOZym5hp7brOlUVisETrLmfL/Jx/orM
woeMz0BPzg13GnLPevCrOyQzjo9m1EdWuPtK5vQDhCVUDWtWPFaa3ULjjBlTRDcCtIqBImWbyjZHnlmSrSU3NCwgzMHg5ThG47Mx
kmtZv6yY1gW0rUc0k5QPbcNSvBKIj27LcoFmfSt01XUSjMSyO7tK7TtZAG7TuJT0wPrM06EJjDUSkFTe9LzQqsZiTIB0VlSz4JeE
1J4EaFy9VzXk/tv7kZW7nUw+ZMPB0TMMAv/1/c5emUZ97rrK1kVB19WokqLcmid7WE4y9cy5GBanA0/JXsY4Dh6hWW3BodDyrVG/
E4FZXEAebK8O72yGIlkhf5qi08t1U8xhG8JxGxM5OA8SlYUAZlZhRqOLaHAJJxHOZxhh7XE+fkfny856A7eBSLwK/Bss+vcb6Dcr
BcR4htFpDLRK+xRYCcU/JhRSyRbeLw4tFgosjRwJ15UCGX4Fzh4sGqXxUaILWRxByLKn0/z6xfLB6NJhGdeLkeh7S+68tEIw7Gor
tLWWQnFBs5Hyqg2KhMC5tGBDkZpa3OzsFAAOFA832I5uiho1Hx+vWMohnjfwXSXhyRkZYJmRS4JXXQgTgDhXDBWQ5283A/ZxNiWo
6qoREWgeQ7gxhcn8Ql+XxtSZ65r6XVHwnTGjeq1R71DF+sEbczD73SXKsrTfprzGNB73wwmu9vq+o7THJ/6Bd8grN4Ju6xEuoIB6
jUELNut3v8EGJArvF/WWV3FiS+A7csxrPeCPZ7lDmV7gLdXHT3/G7a4+ywYkPH2S0mAW1D6tt+iglRgxii/4dzXI6peBJg0q67z/
t7TiEPCtpk9vnlGC5BuIJ4Frn1mtJrGtC58NlCsJL9WllQe8SDZWLABPb7gIVlWcLJiOh+3XhsLvA+mLfTwvr4LhNICfaDRe6vmE
61Ch+jxZ6CTzsLVfn8XaJAoqwls8EXezuiSdynZVS3AWqTzeethfxTc6uyaC63e/pNezG1uIqUdgdHVzEyZWH5QN5XFvnKYZb4VO
rdF/pHe9KgBwoOrxlhULQQhQjITV5ltAwqPT67dl6oHGlqCHKMXQrwZ7YDctXXgJ89+n7TQiiIDA50uch/xaIOYkENm8FQqOBAIW
Qcr1wBgmLiYbzS5UJC/H7spjyiSGwdfjGb8D7Exo//K+xDGrlNlorB4Gc+L05BOBVGu/pi8B44l8u7fpIkUXwTbGdpon7MqaWEhT
jP5FMr7EZkX+GbddjMXit6uX4HgnEkDEnoWSfUGtY4WADUHRJKNm1KGBedsYsgoCbGiSgTbpcvgrSTckGOYJ4IhYUcyh2vSKQ0Ju
j0pY0z6+ZwZRIazmpsJbJA/slZ2dt38VIW1yeCr1lKp5fPvXWM5uRcX+He3mNsyF3fJFepWw9bn0YI5OJVB4os0TmSMgGGEelpUA
5LSKrjwUv+YnwBMW42eRJryNDC0Tp1uewuZFqagT6hC3gWqoFyDZktKTT9nTjsGR8sziN0oKEw0v+vG0FOwhXH0dDRCRHY/QVt1L
x8h855G8k/7oI0hXFgBnMHsGVfazRscqUxHWlf44eYT251Toygjw6n2lB1gpYCCItKUPsxicTMeF/3ksznStcDWdSAxbb3ckuaLZ
H944tUovJ9+t/Ofh0jhckp8gGWl9SAAfvTuAtN5bDEAB9U0RdwYxSsIYUcm3OfHiZ255PTmFsszKaYFkLy1naRv0ClGmO3uFJxPL
T10GqIVsD5lr2GCTnoYDxTPr4KQ2dVQXN4aUn3S9Esn/ou27095TdNV+w22d9loxfOfSl62rnYZuWO5iBspY0dZWCUH+4DSWCvce
4H3u2sguoCZgXuVe8isG7KhPQCmcRle6r0zfwhg9gHyb1yaApuQOP/ChmPmuq57csaWabpdKp5PcfNpfRYrf5iRGZChpGTu+wF1D
3rDi50EH614o69XVTdEEkbLD0w9XgIjDvIQXfY1f9Q+20mND8Q9MdQpZx4kDM6CT2g/mpAR6dqvztuZ1LZOzvlCS+Faum8B4hyE/
iiWXoGHcM6Xs1WlzlSBbuPNPCNlFfvHPKbRMMsI5UpTZwhbZIyvw4BAlFDJsumeLXqDFVKloXhHS2+tL+d13d+F9Th+4l+O+7+AE
TYRAbA3i4heaiCo+K1eNk8bP9f90lOWXoPaw6Wo4CYMd8UDqdQboKtysN0Z9ZhDofwVDA/3i8Mv4G3JvkHMBgZZ9EyxnLBXZIhhL
9EJQc3xU9fujYOBTs1JTk1RNcmrJ5fKGBtKsb0UYRJCMpA5WvIXMZ6TG3Qc4PY7vz3k3qE+xHZeTMC67x2q6vkIP7WT1adqT9w3a
BtZxDhLUSBhc9LREyqcbsWifHfDUNytZPLLUufuxBjXHFY/qVnpvPn3DHtTN8Dyp45qmt76+Rg8KMCU9i7DD0Nkcy8DvZZn+8wEN
4xsTnDtefg1VeG6N+87kW8nBsRYDn4t3i84v2Tp+/sNKd7GXFfAwO7pajUAkiEFgsdke7rHVOMPlJdyA5pXUUQYBAEiQyrWRHoMu
OnY9uj3j2PzuQ8e5lt0mstTrPcFanbbPD/qBDSTsND5UdByVhnqJ8D6I1Vr0jOZk8oXocjvYRvP9xsOmm1HfxoVCyqMmZV22TALq
6T6h78oO5gkA9RPldcNXE5vQjTxdN8SZXss7Jvawi0wMEGy0NtCjBERCAvOIS4cSq/krw0dqgl6R7aHKA8SYvuafEiGh9EGQUEe8
GKVnA+2luAbS3f21ogEsD6rAPl+QdQhdMlceKV7gJqrPgkFTpwctKLwc9ir8mYYsPBsV9388DIgazMP6UAC0O2bLxx0gUJJNYUJp
kmZ3+BYTQ+9feNiNAg0EtcHfoS28JTHQNwby3wBh6MK2yFq2Ppc4P4HNr6xtp20Y0Tuxf9PUGiJ419GtKelpX5pOfJ2f+iK8zn+C
HV0tJJqRXrA0PAxIEJnVLb7Uwi7enlYJXkLqX6uYoANOUJ/26aGxvKhCNUVtVU4TLApkpJBOPMp3jtXr5N1hiEEkvloO7mUS7unf
FunilpfVKWjSaETN507emxYlOdWCQd9hTLQMx3ar/tMEpoggkW/kjUNpaowDff+tN/FRHyEVIoFlLsRzDDD8Osb9rO//kSZHXxcl
Ovzbm2qK18KC4MnyjARCwgNW3grRNT60Fvqm4tebayEBfzXt3TBYvxuSuPpJka9QvuRI+YbJRm2yhGwszy0uOqhT82y9zI1F0ga7
V4OLVdQ4TVvgt5SJVB+Pq86hYc203SsuW691NGt1/7eAoMSE/VpVqIN91Q2PeRKh4iiz5hubeHqw95dk3+BMdaG6vZLJkFg2LzFX
AJ+p0Tc5iY5apHMMooBdgXtXu1/iH2K2Sy0oyLHRxOqCVESEywBhoZ1+bAWbXF2oQondPxhdgLyr5k4LCKzIwVcBIVPZ1xExTvBZ
9odbh8bVv3a4hnq+I7EP7OCxXdeZHz42NTPvsCEdXolXzuZw1Z2S/qB7AZN8ZXuhMu1XD/EwCXZR54Ko4Y8gk8iTYyoKTTbdF8LH
mx4rb92zVGTnsCWnPvkbRJNaQNIS5Ol8oxBTBHmUROa7NB6bx4t4pO9MriUgymBoBCjG74AJ4JpRPcO+b5wXMFxRPg+g2v03snXM
tpf7kSJbw/p/j3jvs0K0IlGuatrY7C3IsMYuVyJgXA7hAoUVgtpZ9C677G2jkPaZbv7IIlWM8aE/ZsVI5yirX1jXs9/8PZ3ZzUMr
kLkZH/pzy7u1Nlh4uOhQ5aYH5TxIM9cS9kD09uAjfJYv00CQiRRSkXbI5lCJh0PcJtcqbjJrtTxkynRnSNW1KezwI16dPzbfv1Ro
sIldajIZcWJnR+iq2UHSc9WCLfzNlGj90O7i+ZPtDgpvYI5fv5iSaljHjFaOLeUhvtURTv/HV7r9Exs2FenTi9N8NfKNZe/XTpFl
w9R+uiOslQA8Rp8E6B8FWZwYr3BCbzoQZIUBIVoHYjHVAj23pIOoiA9k70/VWwHo9K061psXIKH1gQGey+OmjM5mTG17Xw+Z5i3K
7BR6HIARmrE4z1BXitRjuQad3UVBo/jSSsyMT/EGyboTA0oCXrTjFIq+RFvMje6CWFeIm5SDY/r8oT6n+so5a5/aQF0nuxkJgyiW
idSH4EdcOzHBPjjDhIehjR1ZWI2Y19qvU1J2o/48g8krizdgPWQjG6GzfviYff6JFbePC6LX5/v8Utw5neKQthsP2gTZ4WzMGjXx
pTglrd1UGuZPmH9zTqy4deWwh6S7TIayVprmrHoaJsmzI5SSvrx+Dx5EBAZuBWUJDoGD9xsQvRxbw55gX2Kt+KghgD8u0OMoa4b3
znsp8nthcF3+YrbJVr0sixOAa8YV0Wal4+k6dRdyrGXvTaAnLk5AYO3r9GDVCjCkomUuTE9zy5R1yoETuwSFOFprHOvj5a5c+T7n
mN6BpSshBMEL6Rtw62w+8ryuwYB/ZuuPCs5Ov5FdCrDYJr4pHwk61pVmZ3MizuUS8obXPbB8LVSdMSeBsFtKkaT70yYbbyBEEjz0
rmFjDCPm0NflLjHivCel+cpl0v24XQ14pCbA2sL9NPKHIP4bVLaiQQ4M2sVEGr0G5l/3N4sT6buAG1n7zik8/rOLTMusM43ncRjD
561Facu+HM/SYupCY6Lt+pIRGt9HfY2iJdK+I8jOc0sF3Y+H0KfoZnO+D8nI3GF/ZVsde0xa1ST6SBks5weHDmvHpFaxnv/8MuYP
sbI4f6Es1ZmGFR6Ai87bMev2H9PE6Hu2B2qmiL9mlVf63wliGblXTBCI7eGrlgNv2fbslqYFutrHyXRU+ww77dRNaJSQf6BWmG1i
z7aa340coRqMt4htEapqKuBo5G+N7VuelkLWoABwUCS39ggKDpd84mbdYILs3+XtUu5EkXpU++HKvzum/RIp26Dl8JbnRbDTkNuD
wobf+pmH6ckkQgFCQX0tFOqx9jf72VrMXytq3l1dOkE27r0mXsqHr2yJzb1XZZSzfkymH8rXR+1yc7CQq9gSw/bINftR3Y4ZEAgm
S/cw54Pon+nB7DRioj8yRJum6OwhECCmpTc9sJKwOQAW5sTaDRM3bE1Uy8ABrIliXdonEO4kgs2Q2b/P61IFRtOm4dEkXPsOJYAI
SnwZ3ViYWOkgLH9u9x7M9iPsRvTy3ZY2dyINZizukSjOlA1Y2pxYPEkwd6Qew0qEChlmZKwuNX/zthzEUjt9n7GibLwuYlHcdPkT
gi4nEneNCKM+z2x3ki4amIFKFW0yFPLD6/4RUkR67mSeQywOKea85u/2tvMlnE+GuIcjiGJMTwWpABFm0xhGijymk5Q/nboGR3LS
+HeNlxEktRh7mixQAF3Qdrw2DEIXmvXnJt2sgszcspoNBKnBHGacH71NcYzchribAXpflojZeilekdOnYVHMpYea9OaPoM+AgQeP
eh69fWAPl+inFf0yPyE5Q4VQXxBSqNCTa+VHH4JTnnvLcb2/8lDXmbpEM39qNTdOu9NW1kWgBCH5yZdfO2FwN5cvGMkla6QLFi9z
dusLUibtZ29xBaAJtTgZZ92MR5Oo+gBAMOPD/VXRgbcB4z45asuzh5Mb8F1fLar8JSahwCWVa1uLNTAo25l4LzwN2ZPMHONroNnj
U0BhkdQ+oGocsLOtz1I1gAEZrh12SkjWe1s3zzFUQmg57SMqP7rANAlV4tioRJ6BvArAAVl1dugKYEi7lKey71xWcQ8/JQD6JzsG
7Tgm0nKQ5x7Jy9m916KAI581/ivPI6MgBeftgCrV21I53Ifl8fsgJ2u7itLdsrgLCWuqqVBvguuG7Ol7tWNDnzw0so7WNHl+vGFQ
J4TRv6IhtIIVt1uRY+FDPU7+X2fOvhZtP5n3iAJnebaJsqI4s8P3Hn3zWGW44cwRKEMz8zAxqlCPE1cKsEsdl2Nw1GtZy+y4wQoU
k6Jkc5P66pPaAUrjG335cWEqQpPgPhyuurorHd0nTvMY4u4Htjp2wdV7A9CW3mLvmdz58MIl8eVF6B8CKtZo996sVpvl7yhZvt0F
njunKcaoqnd0SoKZhU6WSoUpahF2yaEU1edH+BuNSIJjEkmBr0VeCZbhtAWNEPToaemavFJcf2t5d520nVYpOp/0pj+D5XsZX/2b
AgmNbN6ulq79thi0Pu9zZDwNFNi7oqtqMj0WmUFOTkiEpieRCDBcigxpby9R1sNr8rIFJRcWQici9W9TXpSHkm6+tHQ0VTp3j5a6
i3zsBRZUmX/F2oQAigyAdOAjqFKd6hlwusesvvD2kOMt/+5HtMuNm8h3bdw8xcZ326fBkzqXn7/DzjM/2PDwFQIq38Ae1IHTC6R4
3GzQ3drUWjxoFvDuVQLGGf1F3wdpL796/vrxycCm7QGWhlyeooZxiYDNKJlrrX6B2Vzhm/3kv05tqXKIoRCCuaPWHooUHhF07mwf
MnttvSY1pz9dlwTMwjJ7IPOjlF400pl5pNlz4PtkIFlFa9DEBKi6fG+99AFTKSqPvrnU0u4u7TYmd3rwcbs1KOY4m9qOySp1pSgm
JybplOgJb0TWF/B7FaqsqkddPVmzjZyKaFGexRizMLtBWNlBcWnW82wf6XauAW/gF3tZ9Kat5zffZCazpeUSli4PHxVYbhN0Siin
EK2XFDSzWEDTIj0Nq6NxuzqF2sSdFN1x4ulkz9qBFiF6CNvOH2FoRI3NoJKyjS818+lLtDIXTFLsReI5SANG4/WA4iz5n5/6wTn2
Vw1RjFtJAvQcJ3Q5RJDkVmYQqLNtSK8Nk08GZfoeftQDcSjpI6H/nBtivES7N8BZ8ZWmidOFUIqAvB1nxlbgoP8+O4824ujGtie3
RdQHZTU/nexxcTDpnAFPeSY2+B/rfIK9wW1xSAEzEQpSijAFI+3RLML+BVODmmqcLEf7kKeKMU8CFJHgRmKMSrcl4wMdozFWSt9g
bOJcD7bnWfUFdh755WPfOmhzkgAAMUfK04SLJeWKoGJtviUCnzaT76k4KXdSQMclJfMgY0i5U2F8cv6riaQ5fBBRJNCmDT9QfSPA
1blRxQmkCeoloEC6UgNuKj+dzPa59VeZK5dOp7kOCZy7YMbpFMEbKwoEEPUYcorFOdsnlgOD2f87gsmctsg2HCRf9XT5GLI3ien3
+zGMo4ITqJO0scDAUg89ICgq61nQcmMG6QH1DyYuhUbvjrSv7sBe7mdRlNdSPBab5EMJLNOXrhVWoggZV20BZI8JWn2P593ToooB
a1M0V9E9VVgtf0Vq6Na6JDswM8rzmF5b/cZwxC1eSfmUumn9dZhBuT53VfpmjUy+R+W5rGGqmiSnFBVgQ3h0uTr/8N4QfgJVO2cU
KZBb/bjwgz95BZBPtRixB8z8g0/FWsxQ1bBcIrE1GA+/xmunBiuxPcy6qzUwFNBUEKJYHTP1QvpIRHbqS1UQ9+TbqVFPHdVcDhS4
cKSaYt2dkcDeJZkDqb+udRSNYYWeYymCuW5UZeHF50Il5/6pXVVZxwpqm7sa3JzhCh+tqMniBptzrKrXzqmYXf2pHYg67cWNV77A
u6hWT8Hark1RbleO2lsk4Dhlj/i6yOT2onWHkI0LRUDYACHOU1EPNyd4tqTf5dKPo78uc45hb4bAnLTDGvNphM65Z+H8J1iPB+9P
FVEIwXvZ3oAhzV3tuinjdvlXG2JlV+p3At9Dkgx5IMw9AnLIZaQCrvt3dQoVtd/tiFJfv9X9gkHaKGBY+3q++JgcMYBduXVO2TlH
aBlTM2yJ7proFDv5gXDcamFLSz8xvFXz6A/lxqeNz9lHQ783XcFzDMtDlmZyHtKZeFGEdxenoCUSXAiW2az5TjhGapq0m4H4oHYp
oaUmeCQb9GS7jFrHYWDuefqK7CK7QvRUPIdTwmcPCANt7m3xaWfjtPLRnNE9iz4DE2NbB3vf/U0NUWEgfuycE/6UpsBRioWOQ6r6
mbpUAA2cVQch4WSkvtKvMO3ytLDm8t2lsG2I+ay/ydteauV/SAuxj3kruCpC0NUWcyKD9eUCUzqaS/Tw9jpuJq/4xpuC5Smds1k9
wZMAyQPFq88MGyVgojXbpkLz1/YkCkLZ/7uP4pYmecGup2igG6mmjrMagegGQsMVxn/VyChXcPsrI7balLry6DDEmjM2WlSnhOpB
y5Fa1CTbRrXMPCzTiX+jCC4br05N3EEfBoUVR1TcJvgLTZSerCx4QaRm3Sy5SdL3FeIrr9ERg2uHEPio+5Y4H2inWHH7f+ULpy2q
qXjlJHJJ9vSw87Izcyfi2B8qqS/nm4GnotMoI+MRx6U3g3PiI2gkUvqy8mmKqjyMZzbHC9EIfjn4x5yOgZfvy9k04ZQ0YaEHCcJ9
lPLDXrSj3gl6gtnZVCQCmRC7O84nyrEvRmi8sGk7WwRgWcB6ADhunSZnO0RZr6dfZfz4z942IGo2f1iJ5mx0ZfamTUHPD+B8XFzu
pH5bXqP+iGrLJ1pzhfaOiqfriHWZhKN127f3gAE+Kf4N3RRuZ5IHYzTuTAhLmS5XeIEBMobzZ99CPutJSd2psXBBXltvkoIy9Rpc
jWtPtjgOSidrhb9PAdVQJO/MQWj0D/d4SE+gJL7EHTO7VfEAVkyWcRtW0VNH7XEvPqqOE7m+RBRe6+m3lxV+IzypIbuTV2u7+Inf
HSVwyCxWrvyCe1OuWTBTisWT+oezttIBAPg8AFAdg4q06F5+dGNuutgUgDh++IhICf5rcRjv/A2Q7z9f/FPYwACgrXKjMD5IpYWa
6Wndofs9otJVSMOVSpk61o1QojY2CRlXDdvRznRSxHPH0KEGsP5ErAgvN1w90iPYCqSg3EnbYueoSRbEcWUg9V2UD30VLIV9qxk9
hBD1ArqEhoH+lgBVfk+aO29TMYTQ/+ytdQOrRW+xVP4VjaRMU+lE5BYz/6drsegWwVvLJRGJ9cudWGDiioMXpcMKxRKn87EpX5wL
79mRpHjiNN7PGhTM/84QaccwK91XFPTc1EMkOB907EXtI/a670sPRoCUcSXikFZtLdLKpc/7hCeFskpR2cO8/sNQa8nl7CsZBX44
lETeq73bpS2Jgg4ZftFsYKeaeA2mMmsCSDkkrEvkDDJWBLIWhC9kUjDcHt61/TBGnPlUO3NHnmQ5d+iuKeSclB+yQjIHzVxhpoaa
0SKPXk5+QnNnjzsz/drAdJSz7LZYVTu9vP0Hn/lnFOqEf8qynqBCwYLPAqG1cObnaMk9k1jcW1IiY7LfW5vPGn6VjAppj2WajFde
9SsN11sf69PuPognPz6xIWTe6sAmmihjAVyTCSAjpCLckWScw27+BHM1wi63gb1qWtBdIaaRW1jnoActlvBrh+BMlO7pzuDdJM+F
OSxLAiyT+edyISzLL1pRrA8F37xVGEm6Z0YtgiQ7abIT0masqW0LDsQzZyY6c6wXbtksW3NCb5DPfjKHQStqnXuzG0A4ej3ChYsZ
tlRoXEqjaPAaz6JYXuQ33t6AUdtj8THSEgoSjTkPZi3WecH3gh8gHZ9TPFxH4uJTQFIBKEb+poUwdc3xuR1BRdbwfgPWAHWfNQlU
Y3qEJQoShEA1T5D2ZAS4ufXGgd+Y9MZQQd1/yMtIAO5tgM0bnmlHlF9tddsaR5DwNaqqJDTiMdSKtTcUAHR/iAfP2LXxcjXJB13H
KTG/Md3OObrAIgqwdGKU4P8nGLlRIPbRksZ5GFRuBWq3ZgSrinnd2WI6F0+YXNGEEDQo/LeU5SqsTbTiwHzH5A9P9y+YHH97y+ZO
wrE3RbRnDoNu+EDVlB1P10NCwQaxEJ1O1b2C0tMax5XMLaiAI8aBjWkTIRI/1YX9XSH0qwfx56RkVe7wYgVlBsuKabgtRbZNGP+V
hyO0ZHNmo2pVGVOpTZqSzprZ0/wom7c1TS+GPnsd/BgORscxWuD8ekocfdZIInbRrebzD1ecnuLAIbhOgqTBzEuw+OlooxS2vHwz
DkJqWRTADdkatcZO44HlE7Kn/Vb9Fr7F8hhxmyNyIUZqKZ8D9L2aLCWnEjM33ZMrpt9dGwpwN8XcRWNI2P9Q16UEPd2dG2uVz/sP
pzsMlQxvIN7ixQ7rSFJA8L8nVM/VsR7NouaGVx2cR/Y1SMSCS11pucfzaPGg5MkNSvJlv+h/YdLLC3b8bmJZClvsWlWh0eJqyxJN
Y68cuP+GL7iRCDr4MkZyCl5AUqwbBoxCXxZum/f6TciTIFkmvSXimLxJcLfM5Cy/CYbilHLCLWpco3vMgV+4woX0AsYI1Pw2+n6A
iUWnS82IAGJOI8cXXpx5o5IuzwPI3IZLMjTzmEgA
"""

_SPLASH_JS = r"""(function () {
  var d = document, w = window;
  if (d.getElementById("itp-splash")) return;
  var SRC = "__SRC__";
  var ASSEMBLE = 1.8, HOLD = 0.6, EXPLODE = 1.4, DUR = 1.3, MAXDELAY = ASSEMBLE - DUR;
  var reduced = w.matchMedia && w.matchMedia("(prefers-reduced-motion: reduce)").matches;

  var css = d.createElement("style");
  css.id = "itp-splash-css";
  css.textContent =
    "#itp-splash{position:fixed;left:0;top:0;right:0;bottom:0;z-index:2147483000;background:#fdf5f1;" +
    "cursor:pointer;opacity:1;transition:opacity .5s ease}" +
    "#itp-splash.off{opacity:0;pointer-events:none}" +
    "#itp-splash canvas{position:absolute;left:0;top:0;display:block}" +
    "#itp-splash .flash{position:absolute;left:0;top:0;right:0;bottom:0;pointer-events:none;opacity:0;" +
    "background:radial-gradient(circle at 50% 50%,rgba(255,255,255,.95) 0%,rgba(255,214,150,.55) 22%,rgba(255,150,120,0) 62%)}" +
    "#itp-splash .flash.go{animation:itpflash .55s ease-out}" +
    "@keyframes itpflash{0%{opacity:0;transform:scale(.4)}15%{opacity:.9}100%{opacity:0;transform:scale(1.6)}}";
  d.head.appendChild(css);

  var stage = d.createElement("div");
  stage.id = "itp-splash";
  stage.title = "Click to skip";
  var cv = d.createElement("canvas");
  var flash = d.createElement("div");
  flash.className = "flash";
  stage.appendChild(cv);
  stage.appendChild(flash);
  d.body.appendChild(stage);

  var raf = 0, t0 = 0, N = 0, dpr = 1, cw = 0, ch = 0, bs = 2, cx = 0, cy = 0, exploded = false, shakeOn = false, finished = false;
  var tx, ty, sx, sy, col, al, dl, vx, vy, ph, buf, u32, ctx = cv.getContext("2d");

  function end() {
    if (finished) return;
    finished = true;
    w.cancelAnimationFrame(raf);
    cv.style.transform = "";
    stage.classList.add("off");
    setTimeout(function () {
      if (stage.parentNode) stage.parentNode.removeChild(stage);
      if (css.parentNode) css.parentNode.removeChild(css);
    }, 600);
  }
  stage.addEventListener("click", end);
  setTimeout(end, 10000);   /* failsafe: the splash can never get stuck on screen */

  function build(img) {
    dpr = Math.min(w.devicePixelRatio || 1, 2);
    var W = w.innerWidth, H = w.innerHeight;
    cw = Math.round(W * dpr); ch = Math.round(H * dpr);
    cv.width = cw; cv.height = ch; cv.style.width = W + "px"; cv.style.height = H + "px";
    var iw = Math.min(W * 0.92, 980, (H * 0.86) * 1.5), ih = iw / 1.5, ox = (W - iw) / 2, oy = (H - ih) / 2;
    cx = W * dpr / 2; cy = H * dpr / 2;

    var off = d.createElement("canvas");
    off.width = Math.round(iw); off.height = Math.round(ih);
    var o = off.getContext("2d");
    o.drawImage(img, 0, 0, off.width, off.height);
    var data = o.getImageData(0, 0, off.width, off.height).data;
    var step = Math.max(2, Math.ceil(Math.sqrt(off.width * off.height * 0.5 / 90000)));
    bs = Math.max(1, Math.round(step * dpr));
    var cap = Math.ceil((off.width / step + 1) * (off.height / step + 1));
    tx = new Float32Array(cap); ty = new Float32Array(cap); sx = new Float32Array(cap); sy = new Float32Array(cap);
    col = new Uint32Array(cap); al = new Float32Array(cap); dl = new Float32Array(cap);
    vx = new Float32Array(cap); vy = new Float32Array(cap); ph = new Float32Array(cap);
    N = 0;
    var maxD = Math.hypot(iw / 2, ih / 2) * dpr;
    for (var y = 0; y < off.height; y += step) {
      for (var x = 0; x < off.width; x += step) {
        var k = ((y + (step >> 1)) * off.width + (x + (step >> 1))) * 4;
        if (k >= data.length) continue;
        var r = data[k], g = data[k + 1], b = data[k + 2];
        var dist = Math.sqrt((r - 253) * (r - 253) + (g - 245) * (g - 245) + (b - 241) * (b - 241));
        if (dist < 9) continue;   /* background colour: no particle */
        var i = N++;
        tx[i] = (ox + x) * dpr; ty[i] = (oy + y) * dpr;
        sx[i] = Math.random() * cw; sy[i] = Math.random() * ch;
        col[i] = (r) | (g << 8) | (b << 16);
        al[i] = Math.min(1, (dist - 9) / 26);
        dl[i] = Math.random() * MAXDELAY;
        ph[i] = Math.random() * 6.283;
        var dx = tx[i] - cx, dy = ty[i] - cy;
        var ang = Math.atan2(dy, dx) + (Math.random() - 0.5) * 1.0, dd = Math.hypot(dx, dy) / maxD;
        var sp = (260 + Math.random() * 1250) * (0.55 + dd * 0.9) * dpr;
        vx[i] = Math.cos(ang) * sp;
        vy[i] = Math.sin(ang) * sp - 120 * dpr * Math.random();
      }
    }
    buf = ctx.createImageData(cw, ch);
    u32 = new Uint32Array(buf.data.buffer);
  }

  function easeOut(p) { var q = 1 - p; return 1 - q * q * q; }

  function render(t) {
    u32.fill(0);
    var G = 900 * dpr, et = t - (ASSEMBLE + HOLD);
    if (!reduced && et >= 0 && !exploded) {
      exploded = true;
      flash.classList.remove("go"); void flash.offsetWidth; flash.classList.add("go");
      shakeOn = true;
    }
    var blk = bs + 1;   /* 1px overlap so the formed picture has no gaps */
    if (!reduced && et > 0) blk = Math.max(1, Math.round((bs + 1) * (1 - 0.55 * Math.min(1, et / EXPLODE))));

    for (var i = 0; i < N; i++) {
      var x, y, a;
      if (reduced) {
        x = tx[i]; y = ty[i]; a = al[i] * Math.min(1, t * 4);
      } else if (et >= 0) {
        x = tx[i] + vx[i] * et;
        y = ty[i] + vy[i] * et + 0.5 * G * et * et;
        var f = 1 - et / EXPLODE;
        if (f <= 0) continue;
        a = al[i] * f * Math.sqrt(f);
      } else {
        var p = (t - dl[i]) / DUR;
        if (p < 0) continue;
        if (p > 1) p = 1;
        var e = easeOut(p), wob = (1 - e) * 26 * dpr;
        x = sx[i] + (tx[i] - sx[i]) * e + Math.sin(ph[i] + t * 6) * wob;
        y = sy[i] + (ty[i] - sy[i]) * e + Math.cos(ph[i] + t * 5) * wob;
        a = al[i] * Math.min(1, p * 2.6);
        if (t > ASSEMBLE) {
          var br = 1 + 0.004 * Math.sin((t - ASSEMBLE) * 9);
          x = cx + (x - cx) * br; y = cy + (y - cy) * br;
        }
      }
      var ix = x | 0, iy = y | 0;
      if (ix < 0 || iy < 0 || ix + blk > cw || iy + blk > ch) continue;
      var px = col[i] | (((a * 255) | 0) << 24);
      for (var yy = 0; yy < blk; yy++) {
        var row = (iy + yy) * cw + ix;
        for (var xx = 0; xx < blk; xx++) u32[row + xx] = px;
      }
    }
    ctx.putImageData(buf, 0, 0);

    if (!reduced && et >= 0 && et < 0.6) {   /* shockwave ring */
      ctx.save();
      ctx.lineWidth = (10 - et * 12) * dpr;
      ctx.strokeStyle = "rgba(255,120,110," + Math.max(0, 0.55 - et * 0.9).toFixed(3) + ")";
      ctx.beginPath(); ctx.arc(cx, cy, et * 1100 * dpr, 0, 6.283); ctx.stroke();
      ctx.restore();
    }
    if (shakeOn) {
      if (et < 0.4) {
        var m = (1 - et / 0.4) * 9;
        cv.style.transform = "translate(" + ((Math.random() - 0.5) * m).toFixed(1) + "px," + ((Math.random() - 0.5) * m).toFixed(1) + "px)";
      } else { cv.style.transform = ""; shakeOn = false; }
    }
  }

  function frame(now) {
    if (finished) return;
    var t = (now - t0) / 1000;
    if (reduced) { if (t > 1.6) { end(); return; } }
    else if (t > ASSEMBLE + HOLD + EXPLODE) { end(); return; }
    render(t);
    raf = w.requestAnimationFrame(frame);
  }

  var img = new Image();
  img.onload = function () {
    try {
      build(img);
      t0 = w.performance.now();
      raf = w.requestAnimationFrame(frame);
    } catch (err) { end(); }
  };
  img.onerror = end;
  img.src = SRC;
})();
"""


def show_login_splash_once() -> None:
    if st.session_state.get("_splash_shown"):
        return
    st.session_state["_splash_shown"] = True
    import streamlit.components.v1 as components
    src = "data:image/webp;base64," + _SPLASH_IMAGE_B64.replace("\n", "")
    script = _SPLASH_JS.replace("__SRC__", src)
    components.html(
        "<script>(function(){try{"
        "var p=window.parent,s=p.document.createElement('script');"
        "s.textContent=" + json.dumps(script) + ";"
        "p.document.head.appendChild(s);"
        "var f=window.frameElement,"
        "c=f&&(f.closest('[data-testid=\"stElementContainer\"]')||f.parentElement);"
        "if(c)c.style.display='none';"
        "}catch(e){}})();</script>",
        height=0,
    )


if not _RUNNING_UNDER_TEST:
    show_login_splash_once()

if "task" not in st.session_state:
    st.session_state["task"] = None

_document_repo_actor_email = _document_repository_actor_email()
_show_document_repository_tile = _can_see_document_repository(_document_repo_actor_email)

# Exactly TASKS, unchanged, when the tile is hidden -- so the landing
# screen's columns/layout are identical to before this tile existed.
TILES = TASKS + ["Document Repository"] if _show_document_repository_tile else TASKS

# ---------------------------------------------------------------------------
# Landing screen: pick a task
# ---------------------------------------------------------------------------
if st.session_state["task"] is None:
    st.subheader("What would you like to do?")
    cols = st.columns(len(TILES))
    for col, task_name in zip(cols, TILES):
        with col:
            if task_name in DISABLED_TASKS:
                st.markdown(
                    f"""
                    <div style="position:relative; overflow:hidden; border:1px solid rgba(255,255,255,0.18);
                                border-radius:14px; padding:16px; text-align:center;
                                background:rgba(255,255,255,0.03); opacity:0.6; filter:grayscale(60%);">
                        <div style="position:absolute; top:18px; right:-42px; width:170px;
                                    transform:rotate(45deg); background:#5c6b72; color:white;
                                    text-align:center; font-size:11px; font-weight:700;
                                    letter-spacing:0.4px; padding:4px 0;
                                    box-shadow:0 2px 4px rgba(0,0,0,0.35);">
                            {DISABLED_BANNER_TEXT}
                        </div>
                        <div style="font-size:2rem;">{TASK_ICONS[task_name]}</div>
                        <div style="color:#dce8ec;"><strong>{task_name}</strong></div>
                        <div style="font-size:0.85rem; color:#8fa3aa; margin-top:4px;">
                            {TASK_DESCRIPTIONS[task_name]}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.button(
                    "Select", key=f"select_{task_name}", use_container_width=True,
                    disabled=True, help="On hold for now.",
                )
            else:
                with st.container(border=True):
                    st.markdown(f"## {TASK_ICONS[task_name]}")
                    st.markdown(f"**{task_name}**")
                    st.caption(TASK_DESCRIPTIONS[task_name])
                    if st.button("Select", key=f"select_{task_name}", use_container_width=True):
                        st.session_state["task"] = task_name
                        st.rerun()
    st.stop()

# ---------------------------------------------------------------------------
# Active task header + back navigation
# ---------------------------------------------------------------------------
col_back, col_title = st.columns([1, 5])
with col_back:
    if st.button("← Back"):
        st.session_state["task"] = None
        st.rerun()
with col_title:
    st.subheader(f"{TASK_ICONS[st.session_state['task']]} {st.session_state['task']}")

task = st.session_state["task"]

# Permission can change between page loads (e.g. an admin's access is
# revoked mid-session) -- re-checked here, not just at selection time,
# so a stale session can never keep showing this screen.
if task == "Document Repository" and not _show_document_repository_tile:
    st.session_state["task"] = None
    st.rerun()

# ---------------------------------------------------------------------------
# Trip Preference Parsing -- fully wired to the existing pipeline
# ---------------------------------------------------------------------------
if task == "Trip Preference Parsing":
    tab_new, tab_requests, tab_entry = st.tabs(
        ["Submit New Request", "My Requests", "Entry-Ready View"]
    )

    with tab_new:
        # Plain widgets (not st.form) on purpose: inside a form, Streamlit
        # batches every input and only reruns the script on submit, so the
        # Amount/Notes fields wouldn't appear until after clicking Submit.
        # Outside a form, selecting "Ok to Invoice" triggers an immediate
        # rerun, and these fields open up right away as required.
        salespeople = backend.list_salespeople()

        if not salespeople:
            st.warning("No salespeople configured yet. Add one below before submitting a request.")
            salesperson = None
        else:
            salesperson = st.selectbox("Salesperson", salespeople, key="salesperson")

        # Compact toggle button (sized to its label, not the full row) so
        # this passive, occasionally-used feature stays out of the way
        # until someone actually clicks it -- only then does it expand
        # into a full-width panel with room for all the controls.
        if "show_salesperson_manager" not in st.session_state:
            st.session_state["show_salesperson_manager"] = False

        toggle_label = (
            "▾ Manage salesperson list" if st.session_state["show_salesperson_manager"]
            else "⚙️ Manage salesperson list"
        )
        if st.button(toggle_label, key="toggle_salesperson_manager"):
            st.session_state["show_salesperson_manager"] = not st.session_state["show_salesperson_manager"]

        if st.session_state["show_salesperson_manager"]:
            with st.container(border=True):
                st.caption("Changes apply immediately for everyone using this app.")

                st.markdown("**Add**")
                new_name = st.text_input("New salesperson name", key="new_salesperson_name")
                if st.button("Add", key="add_salesperson_btn"):
                    try:
                        backend.add_salesperson(new_name)
                        st.toast(f"Added '{new_name.strip()}'.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

                if salespeople:
                    st.markdown("**Rename**")
                    rename_target = st.selectbox("Salesperson", salespeople, key="rename_salesperson_select")
                    rename_new_name = st.text_input("New name", key="rename_salesperson_new_name")
                    if st.button("Save rename", key="save_rename_btn"):
                        try:
                            backend.update_salesperson(rename_target, rename_new_name)
                            st.toast(f"Renamed '{rename_target}' to '{rename_new_name.strip()}'.")
                            st.rerun()
                        except ValueError as e:
                            st.error(str(e))

                    st.markdown("**Delete**")
                    delete_target = st.selectbox("Salesperson", salespeople, key="delete_salesperson_select")
                    if st.button("Delete", key="delete_salesperson_btn"):
                        backend.delete_salesperson(delete_target)
                        st.toast(f"Deleted '{delete_target}'.")
                        st.rerun()

                st.markdown("**Backup / restore**")
                st.caption(
                    "This app's storage isn't guaranteed to survive every redeploy. "
                    "Download a backup now and then, and restore it here if the "
                    "list ever comes back empty unexpectedly."
                )
                st.download_button(
                    "Download backup (.json)",
                    data=backend.export_salespeople_json(),
                    file_name="salespeople_backup.json",
                    mime="application/json",
                    key="download_salespeople_backup",
                )
                restore_file = st.file_uploader(
                    "Restore from a backup file", type=["json"], key="restore_salespeople_file"
                )
                if restore_file is not None and st.button("Restore", key="restore_salespeople_btn"):
                    try:
                        added, skipped = backend.import_salespeople_json(restore_file.read().decode("utf-8"))
                        st.toast(f"Restored: added {added}, already present {skipped}.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

        acct_nbr = st.text_input(
            "Jeppesen Acct Nbr:",
            max_chars=MAX_ACCT_NBR_LEN,
            help=f"Max {MAX_ACCT_NBR_LEN} characters.",
            key="acct_nbr",
        )

        payment_method = st.radio("Payment Method", PAYMENT_METHODS, horizontal=True, key="payment_method")

        amount = None
        notes = ""
        if payment_method == "Ok to Invoice":
            amount = st.number_input(
                "Amount (required)", min_value=0.0, step=0.01, format="%.2f",
                help="Required when 'Ok to Invoice' is selected.",
                key="amount",
            )
            notes = st.text_area("Notes (optional)", key="notes")

        # st.file_uploader always renders as a drag-and-drop dropzone
        # (plus a Browse button) -- no extra work needed for that.
        # Required -- the parser has nothing to run against without it, so
        # both the label and a caption call that out clearly, and the
        # Submit validation below blocks the request until a file is here.
        st.markdown("##### Onboarding Form Upload &nbsp; :red[**\\* Required**]")
        st.caption("The parser can't run without this file -- upload the completed onboarding form to continue.")
        uploaded_file = st.file_uploader(
            "Drag and drop, or browse for, a PDF/Word/image of the completed onboarding form",
            type=ALLOWED_UPLOAD_TYPES,
            key="uploaded_file",
        )

        submitted = st.button("Submit")

        if submitted:
            errors = []
            if salesperson is None:
                errors.append("Add a salesperson to the list above before submitting.")
            if payment_method == "Ok to Invoice" and (amount is None or amount <= 0):
                errors.append("Amount is required and must be greater than 0 when 'Ok to Invoice' is selected.")
            if uploaded_file is None:
                errors.append("Upload the onboarding form before submitting -- it's required to run the parser.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                onboarding_bytes = uploaded_file.getvalue() if uploaded_file is not None else None
                onboarding_filename = uploaded_file.name if uploaded_file is not None else None

                request_id = backend.insert_request({
                    "salesperson": salesperson,
                    "jeppesen_acct_nbr": acct_nbr,
                    "payment_method": payment_method,
                    "amount": amount if payment_method == "Ok to Invoice" else None,
                    "notes": notes if payment_method == "Ok to Invoice" else "",
                    "onboarding_filename": onboarding_filename,
                    "onboarding_bytes": onboarding_bytes,
                })

                st.success(f"Submitted. Request ID: {request_id}")

                if onboarding_bytes is not None:
                    progress = st.empty()
                    bar = progress.progress(0.0, text="Starting...")

                    def _on_progress(fraction, text):
                        bar.progress(fraction, text=text)

                    # Keep the mapped subjects (what the Entry-Ready View
                    # displays) and the raw extraction (what its
                    # field-by-field checks work from) for this browser
                    # session only. Neither is written to the database --
                    # persisting them would need a schema change and a
                    # deliberate decision, since both contain customer PII
                    # in structured, queryable form.
                    def _keep_sections(sections, by_subject, _rid=request_id):
                        st.session_state.setdefault("parsed_sections", {})[_rid] = {
                            "sections": sections, "by_subject": by_subject,
                        }

                    backend.run_parsing(
                        request_id, onboarding_bytes,
                        progress_callback=_on_progress,
                        secrets_getter=lambda: st.secrets,
                        salesperson=salesperson,
                        sections_sink=_keep_sections,
                    )
                    progress.empty()
                    refreshed = backend.fetch_request(request_id)
                    if refreshed["status"] == backend.STATUS_COMPLETE:
                        st.success(f"Request {request_id} complete -- see it under 'My Requests'.")
                    else:
                        st.error(f"Request {request_id} failed: {refreshed['error_message']}")
                else:
                    st.info("No onboarding form was attached, so there is nothing to parse for this request.")

    with tab_requests:
        if st.button("Refresh"):
            st.rerun()

        rows = backend.fetch_all_requests()
        if not rows:
            st.caption("No requests yet.")
        for row in rows:
            status_emoji = {
                backend.STATUS_SUBMITTED: "○",
                backend.STATUS_PROCESSING: "⏳",
                backend.STATUS_COMPLETE: "✅",
                backend.STATUS_FAILED: "❌",
            }.get(row["status"], "")
            with st.expander(f"{status_emoji} {row['id']} -- {row['salesperson']} -- {row['status']} ({row['created_at']})"):
                st.write(f"**Jeppesen Acct Nbr:** {row['jeppesen_acct_nbr'] or '—'}")
                st.write(f"**Payment Method:** {row['payment_method']}")
                if row["payment_method"] == "Ok to Invoice":
                    st.write(f"**Amount:** {row['amount']}")
                    st.write(f"**Notes:** {row['notes'] or '—'}")
                st.write(f"**Onboarding Form:** {row['onboarding_filename'] or 'None uploaded'}")
                st.write(f"**Status:** {row['status']}")

                if row["status"] == backend.STATUS_COMPLETE and row["result_bytes"]:
                    st.download_button(
                        "Download Operational Information Excel",
                        data=row["result_bytes"],
                        file_name=f"{row['id']}_operational_info.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{row['id']}",
                    )
                elif row["status"] == backend.STATUS_FAILED:
                    st.error(row["error_message"] or "Unknown error.")
                    if row["onboarding_bytes"] and st.button("Retry", key=f"retry_{row['id']}"):
                        backend.update_status(row["id"], backend.STATUS_PROCESSING)
                        progress = st.empty()
                        bar = progress.progress(0.0, text="Retrying...")

                        def _on_retry_progress(fraction, text):
                            bar.progress(fraction, text=text)

                        def _keep_retry_sections(sections, by_subject, _rid=row["id"]):
                            st.session_state.setdefault("parsed_sections", {})[_rid] = {
                                "sections": sections, "by_subject": by_subject,
                            }

                        backend.run_parsing(
                            row["id"], row["onboarding_bytes"],
                            progress_callback=_on_retry_progress,
                            secrets_getter=lambda: st.secrets,
                            salesperson=row["salesperson"],
                            sections_sink=_keep_retry_sections,
                        )
                        progress.empty()
                        st.rerun()
                elif row["status"] == backend.STATUS_PROCESSING:
                    st.info("Still processing -- click Refresh above to check again.")

    with tab_entry:
        st.caption(
            "The finished Operational Information rows on screen, one copy-paste "
            "block per Subject — the same text the Excel export puts in that row's "
            "Comments cell, so nothing has to be retyped or reassembled. The Excel "
            "export is unchanged and still available under My Requests."
        )

        parsed = st.session_state.get("parsed_sections", {})

        # Completed requests already saved in the database. The finished Excel
        # for each one is stored with the request (it is what the Download
        # button under My Requests serves), so the Entry-Ready View can be
        # rebuilt from it at any time -- nothing new is stored. `rows` is the
        # list the My Requests tab loaded earlier in this same run. A request
        # parsed in this browser session is left out of this list because it
        # is already offered above with its full field-level checks.
        saved = {}
        for _row in rows:
            if (_row["status"] == backend.STATUS_COMPLETE and _row["result_bytes"]
                    and _row["id"] not in parsed):
                saved[_row["id"]] = _row

        options = (["Upload an extraction file..."]
                   + sorted(parsed, reverse=True)
                   + list(saved))

        def _document_label(option: str) -> str:
            """What the dropdown shows. The option value stays the request
            id; only the text on screen changes. The customer's name, as
            parsed from the document, comes first; if the document had no
            Company Name the label is just what it was before."""
            if option in saved:
                row = saved[option]
                name = customer_name_from_workbook(bytes(row["result_bytes"]))
                label = f"{option} -- {row['salesperson']} ({row['created_at']})"
            elif option in parsed:
                name = customer_name_from_by_subject((parsed[option] or {}).get("by_subject"))
                label = option
            else:
                return option
            return f"{name} -- {label}" if name else label

        picked = st.selectbox(
            "Parsed document", options, key="entry_source",
            format_func=_document_label,
            help="Any completed request appears here, so you can come back to it later. "
                 "You can also choose 'Upload an extraction file...' to load an Excel "
                 "export from a file.",
        )

        sections = None
        by_subject = None
        if picked == "Upload an extraction file...":
            if not parsed and not saved:
                st.info(
                    "No completed requests yet. Submit a request under 'Submit New "
                    "Request', or upload an Excel export from a previous parse below."
                )
            upload = st.file_uploader(
                "Excel export (.xlsx) or extraction file (.json)",
                type=["xlsx", "json"], key="entry_upload",
                help="Upload the Operational Information Excel file this tool exported "
                     "(from 'My Requests'). A .json extraction file is also accepted.",
            )
            if upload is not None and upload.name.lower().endswith(".xlsx"):
                by_subject, xlsx_error = by_subject_from_workbook(upload.getvalue())
                if xlsx_error:
                    st.error(xlsx_error)
                    by_subject = None
                else:
                    sections = []
                    st.success(
                        "Loaded from the Excel export. The blocks below are the exported "
                        "text exactly as it was saved, including the account number, "
                        "account exec name and High Profile Notes month."
                    )
            elif upload is not None:
                try:
                    payload = json.loads(upload.getvalue().decode("utf-8"))
                except (ValueError, UnicodeDecodeError):
                    st.error("That file could not be read as JSON.")
                    payload = None
                if payload is not None:
                    sections = payload.get("sections") if isinstance(payload, dict) else None
                    if not isinstance(sections, list):
                        st.error(
                            "That JSON has no 'sections' list, so it is not an extraction "
                            "file from this engine."
                        )
                        sections = None
                    else:
                        # An extraction file holds only what came out of the
                        # PDF. The account number, the salesperson and the
                        # High Profile Notes date come from the intake form,
                        # so a document loaded this way shows the engine's
                        # placeholders for them. Said plainly rather than
                        # silently, because those three are the lines
                        # someone would most likely paste without checking.
                        by_subject = pe.map_to_subjects(sections)
                        st.info(
                            "Loaded from a file, so the Jeppesen account number, the "
                            "account exec name and the High Profile Notes month come "
                            "out as placeholders — those three come from the intake "
                            "form, not the PDF. Pick the request from the dropdown "
                            "above instead to see the real values."
                        )
        elif picked in saved:
            by_subject, saved_error = by_subject_from_workbook(bytes(saved[picked]["result_bytes"]))
            if saved_error:
                st.error(saved_error)
                by_subject = None
            else:
                sections = []
                st.caption(
                    "Loaded from the saved Excel export for this request, exactly as "
                    "it appears under 'My Requests'."
                )
        else:
            stored = parsed.get(picked) or {}
            sections = stored.get("sections")
            by_subject = stored.get("by_subject")

        if by_subject is not None:
            render_entry_ready_view(by_subject, sections or [], key_prefix=f"entry_{picked}")

# ---------------------------------------------------------------------------
# Flight Planning Parsing -- ON HOLD. Upload UI is live for later testing,
# but wiring this up is paused for now. The target output design already
# exists as flight_plan_template_output.html (the "Edit Flight Plan
# Template" mock-up) -- once this resumes, extraction rules get defined
# for the Flight Planning Preferences source document (a different form
# than Trip Support Service Preferences, so it needs its own field
# whitelist in parsing_engine.py), and the parsed values get mapped onto
# that template's fields.
# ---------------------------------------------------------------------------
elif task == "Flight Planning Parsing":
    st.info(
        "This will parse a completed **Flight Planning Preferences** form and use "
        "it to fill out the Flight Plan Template design already built "
        "(flight_plan_template_output.html). This task is on hold for now -- the "
        "upload below is wired up for later, but extraction rules haven't been "
        "defined yet. When we pick this back up, share a sample of a completed "
        "Flight Planning Preferences form and it gets built the same careful way "
        "the Trip Preference pipeline was."
    )

    fp_uploaded_file = st.file_uploader(
        "Flight Planning Preferences form",
        type=ALLOWED_UPLOAD_TYPES,
        help="Drag and drop, or browse for, a PDF/Word/image of the completed Flight Planning Preferences form.",
        key="fp_uploaded_file",
    )

    st.button(
        "Submit", key="fp_submit", disabled=True,
        help="Disabled until the extraction rules for this document are defined.",
    )

# ---------------------------------------------------------------------------
# Weekly Report -- one shared Period selector (always the current period
# plus two weeks before/after -- see backend.available_periods()) drives
# all three tabs below:
#   - Submit My Report: create/edit your own report for the selected
#     period. Locked automatically once that period's edit cutoff has
#     passed -- see backend.is_period_editable().
#   - View All Reports: every team member's status/timestamps for the
#     selected period, with inline View/Edit/Create per person.
#   - Compile Weekly Report: groups every submission for the selected
#     period into one report, ready to email (Gmail SMTP), copy, or
#     download.
# render_editable_report_form() and render_report_readonly() below are
# shared by both the Submit tab and the View All Reports tab so the
# create/edit UI and the lock rule behave identically everywhere.
# ---------------------------------------------------------------------------
elif task == "Weekly Report":
    wr_tz = backend.get_business_timezone()
    wr_current_monday = backend.current_period_monday(wr_tz)
    wr_period_mondays = backend.available_periods(wr_tz)
    wr_period_options = [backend.period_id_from_monday(m) for m in wr_period_mondays]
    wr_current_period_id = backend.period_id_from_monday(wr_current_monday)

    # Default to the current period on first load; if a previously
    # selected period has rolled outside the 5-period window (e.g. the
    # browser tab was left open for a few weeks), fall back to current
    # rather than getting stuck on an option that no longer exists.
    if st.session_state.get("wr_selected_period") not in wr_period_options:
        st.session_state["wr_selected_period"] = wr_current_period_id

    selected_period_id = st.selectbox(
        "Reporting period",
        options=wr_period_options,
        format_func=lambda pid: backend.format_period_label(
            backend.monday_from_period_id(pid), is_current=(pid == wr_current_period_id)
        ),
        key="wr_selected_period",
    )
    selected_monday = backend.monday_from_period_id(selected_period_id)
    period_editable = backend.is_period_editable(selected_monday)

    tab_submit, tab_view, tab_compile = st.tabs(
        ["Submit My Report", "View All Reports", "Compile Weekly Report"]
    )

    with tab_submit:
        team = backend.list_report_team_members()

        # Same compact toggle pattern as "Manage salesperson list" -- kept
        # out of the way until someone actually needs to edit the roster.
        if "show_team_manager" not in st.session_state:
            st.session_state["show_team_manager"] = False

        toggle_team_label = (
            "▾ Manage team members" if st.session_state["show_team_manager"]
            else "⚙️ Manage team members"
        )
        if st.button(toggle_team_label, key="toggle_team_manager"):
            st.session_state["show_team_manager"] = not st.session_state["show_team_manager"]

        if st.session_state["show_team_manager"]:
            with st.container(border=True):
                st.caption("Changes apply immediately for everyone using this app.")

                st.markdown("**Add**")
                new_team_name = st.text_input("New team member name", key="new_team_member_name")
                if st.button("Add", key="add_team_member_btn"):
                    try:
                        backend.add_report_team_member(new_team_name)
                        st.toast(f"Added '{new_team_name.strip()}'.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

                if team:
                    st.markdown("**Rename**")
                    rename_team_target = st.selectbox("Team member", team, key="rename_team_member_select")
                    rename_team_new_name = st.text_input("New name", key="rename_team_member_new_name")
                    if st.button("Save rename", key="save_team_rename_btn"):
                        try:
                            backend.update_report_team_member(rename_team_target, rename_team_new_name)
                            st.toast(f"Renamed '{rename_team_target}' to '{rename_team_new_name.strip()}'.")
                            st.rerun()
                        except ValueError as e:
                            st.error(str(e))

                    st.markdown("**Delete**")
                    delete_team_target = st.selectbox("Team member", team, key="delete_team_member_select")
                    if st.button("Delete", key="delete_team_member_btn"):
                        backend.delete_report_team_member(delete_team_target)
                        st.toast(f"Deleted '{delete_team_target}'.")
                        st.rerun()

                st.markdown("**Backup / restore**")
                st.caption(
                    "This app's storage isn't guaranteed to survive every redeploy. "
                    "Download a backup now and then, and restore it here if the "
                    "list ever comes back empty unexpectedly."
                )
                st.download_button(
                    "Download backup (.json)",
                    data=backend.export_report_team_members_json(),
                    file_name="report_team_members_backup.json",
                    mime="application/json",
                    key="download_team_members_backup",
                )
                restore_team_file = st.file_uploader(
                    "Restore from a backup file", type=["json"], key="restore_team_members_file"
                )
                if restore_team_file is not None and st.button("Restore", key="restore_team_members_btn"):
                    try:
                        added, skipped = backend.import_report_team_members_json(restore_team_file.read().decode("utf-8"))
                        st.toast(f"Restored: added {added}, already present {skipped}.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

        if not team:
            st.warning("No team members configured yet. Add names above before submitting a report.")
        else:
            # No default selection -- starts blank every time so nobody can
            # submit under someone else's name without realizing it. Only
            # the person filling this in should ever pick their own name.
            submitter = st.selectbox(
                "Your name", team, index=None, placeholder="Select your name...", key="wr_submitter"
            )
            is_current = selected_monday == wr_current_monday
            st.caption(
                f"Reporting period: {backend.format_period_label(selected_monday, is_current=is_current)}"
            )

            if not submitter:
                st.info("Select your name above to view or submit your report for this period.")
            elif period_editable:
                render_editable_report_form(submitter, selected_period_id, key_prefix="wr")
            else:
                cutoff = backend.edit_cutoff(selected_monday)
                st.info(
                    f"This period locked on {cutoff.strftime('%b %d, %Y')} at "
                    f"{cutoff.strftime('%I:%M %p %Z')} and can no longer be edited."
                )
                existing = backend.get_weekly_report(submitter, selected_period_id)
                if existing:
                    render_report_readonly(existing)
                else:
                    st.caption(f"No report was submitted by {submitter} for this period.")

    with tab_view:
        team = backend.list_report_team_members()

        if not team:
            st.info("No team members configured yet -- add some in the Submit My Report tab.")
        else:
            rows_by_submitter = {
                r["submitter"]: r for r in backend.list_weekly_reports_for_period(selected_period_id)
            }
            submitted_count = sum(1 for name in team if name in rows_by_submitter)
            st.caption(f"{submitted_count} of {len(team)} team members submitted for this period.")
            if submitted_count == 0:
                st.info("No reports have been submitted yet for this period.")

            active = st.session_state.get("wrview_active")

            for idx, name in enumerate(team):
                row = rows_by_submitter.get(name)
                with st.container(border=True):
                    # Submitted/Updated columns widened (and Name/Status
                    # trimmed) to make room for three stacked timezone
                    # lines each, instead of the single business-timezone
                    # line shown elsewhere in the app.
                    col_name, col_status, col_submitted, col_updated, col_action = st.columns(
                        [0.20, 0.13, 0.25, 0.25, 0.17]
                    )
                    with col_name:
                        st.markdown(f"**{name}**")
                    with col_status:
                        st.markdown("🟢 Submitted" if row else "⚪ Not submitted")
                    with col_submitted:
                        submitted_lines = backend.format_timestamp_multi_tz(row["submitted_at"]) if row else ["—"]
                        for line in submitted_lines:
                            st.caption(line)
                    with col_updated:
                        updated_lines = backend.format_timestamp_multi_tz(row["updated_at"]) if row else ["—"]
                        for line in updated_lines:
                            st.caption(line)
                    with col_action:
                        if row and period_editable:
                            if st.button("Edit", key=f"wrview_edit_btn_{idx}"):
                                st.session_state["wrview_active"] = (name, "edit")
                                st.rerun()
                        elif row and not period_editable:
                            if st.button("View", key=f"wrview_view_btn_{idx}"):
                                st.session_state["wrview_active"] = (name, "view")
                                st.rerun()
                        elif not row and period_editable:
                            if st.button("Create", key=f"wrview_create_btn_{idx}"):
                                st.session_state["wrview_active"] = (name, "edit")
                                st.rerun()
                        else:
                            st.caption("Locked")

                    if active and active[0] == name:
                        _, mode = active
                        st.markdown("---")
                        if mode == "edit" and period_editable:
                            render_editable_report_form(name, selected_period_id, key_prefix=f"wrview_{idx}")
                        elif row:
                            render_report_readonly(row)
                        else:
                            st.caption("No report was submitted for this period.")
                        if st.button("Close", key=f"wrview_close_btn_{idx}"):
                            st.session_state.pop("wrview_active", None)
                            st.rerun()

    with tab_compile:
        period_to_compile = selected_period_id
        st.caption(
            f"Compiling: {backend.format_period_label(selected_monday, is_current=(selected_monday == wr_current_monday))} "
            "-- change the period at the top of the page to compile a different week."
        )

        team = backend.list_report_team_members()
        rows = backend.list_weekly_reports_for_period(period_to_compile)
        submitted_names = {r["submitter"] for r in rows}
        missing = [n for n in team if n not in submitted_names]

        if team:
            st.caption(f"{len(rows)} of {len(team)} team members submitted for this period.")
        else:
            st.caption(f"{len(rows)} submission(s) for this period.")
        if missing:
            st.warning("Not yet submitted: " + ", ".join(missing))

        # Same compact toggle pattern as "Manage salesperson list" -- kept
        # out of the way until someone actually needs to edit recipients.
        if "show_recipient_manager" not in st.session_state:
            st.session_state["show_recipient_manager"] = False

        recipients = backend.list_recipients()

        toggle_label = (
            "▾ Manage recipients" if st.session_state["show_recipient_manager"]
            else "⚙️ Manage recipients"
        )
        if st.button(toggle_label, key="toggle_recipient_manager"):
            st.session_state["show_recipient_manager"] = not st.session_state["show_recipient_manager"]

        if st.session_state["show_recipient_manager"]:
            with st.container(border=True):
                st.caption("Changes apply immediately for everyone using this app.")

                st.markdown("**Add**")
                new_email = st.text_input("Email address", key="new_recipient_email")
                if st.button("Add", key="add_recipient_btn"):
                    try:
                        backend.add_recipient(new_email)
                        st.toast(f"Added '{new_email.strip()}'.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

                if recipients:
                    st.markdown("**Delete**")
                    delete_target = st.selectbox("Recipient", recipients, key="delete_recipient_select")
                    if st.button("Delete", key="delete_recipient_btn"):
                        backend.delete_recipient(delete_target)
                        st.toast(f"Deleted '{delete_target}'.")
                        st.rerun()

                st.markdown("**Backup / restore**")
                st.caption(
                    "This app's storage isn't guaranteed to survive every redeploy. "
                    "Download a backup now and then, and restore it here if the "
                    "list ever comes back empty unexpectedly."
                )
                st.download_button(
                    "Download backup (.json)",
                    data=backend.export_recipients_json(),
                    file_name="report_recipients_backup.json",
                    mime="application/json",
                    key="download_recipients_backup",
                )
                restore_file = st.file_uploader(
                    "Restore from a backup file", type=["json"], key="restore_recipients_file"
                )
                if restore_file is not None and st.button("Restore", key="restore_recipients_btn"):
                    try:
                        added, skipped = backend.import_recipients_json(restore_file.read().decode("utf-8"))
                        st.toast(f"Restored: added {added}, already present {skipped}.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))

        st.markdown("**Recipients**")
        if recipients:
            st.code("; ".join(recipients), language=None)
        else:
            st.info("No recipients added yet -- add some above before sending this out.")

        report_text = backend.compile_weekly_report_text(period_to_compile, expected_submitters=team or None)

        st.markdown("**Compiled report**")
        st.caption("Click the copy icon in the corner to copy this into an email, or download it below.")
        st.code(report_text, language=None)

        st.download_button(
            "Download report (.txt)",
            data=report_text,
            file_name=f"weekly_report_{period_to_compile}.txt",
            mime="text/plain",
            key="wr_download_btn",
        )

        st.markdown("**Send by email**")
        gmail_configured = bool(st.secrets.get("GMAIL_ADDRESS")) and bool(st.secrets.get("GMAIL_APP_PASSWORD"))
        if not gmail_configured:
            st.info(
                "Email sending isn't configured yet. Add `GMAIL_ADDRESS` and "
                "`GMAIL_APP_PASSWORD` (a Gmail **App Password**, not your regular "
                "Gmail password) to this app's secrets, and this button will send "
                "automatically from then on. Until then, copy the report above "
                "into an email."
            )
        elif not recipients:
            st.info("Add at least one recipient above before sending.")
        else:
            if st.button("Send report now", key="wr_send_btn", type="primary"):
                subject = f"Weekly Team Report -- {backend.period_label(period_to_compile)}"
                try:
                    with st.spinner("Sending..."):
                        backend.send_weekly_report_email(
                            recipients=recipients,
                            subject=subject,
                            body=report_text,
                            secrets_getter=lambda: st.secrets,
                        )
                    st.success(f"Sent to {len(recipients)} recipient(s).")
                except (ValueError, RuntimeError) as e:
                    st.error(str(e))

# ---------------------------------------------------------------------------
# Document Repository -- Phase 2: browse/search/download only. See
# documents_ui.py for the screen itself; permission is already
# re-checked above (right after `task = st.session_state["task"]`).
# ---------------------------------------------------------------------------
elif task == "Document Repository":
    documents_ui.render_document_repository(_document_repo_actor_email)
