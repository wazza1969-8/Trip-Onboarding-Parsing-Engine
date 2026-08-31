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
    submissions into one report, ready to copy into an email or
    download. Actually sending email isn't wired up yet -- pick an
    email method (Gmail app password, company SMTP, or a transactional
    API like SendGrid) and that's a small follow-on to add.

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

import uuid

import streamlit as st

import backend

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
}
TASK_DESCRIPTIONS = {
    "Trip Preference Parsing": "Parse a completed Trip Support Service Preferences "
                                "onboarding form into the Operational Information spreadsheet.",
    "Flight Planning Parsing": "Parse a completed Flight Planning Preferences form "
                                "into the Flight Plan Template.",
    "Weekly Report": "Submit your weekly update, or compile everyone's "
                     "submissions into one team report.",
}

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
        color: #0b1c26 !important;
        border-radius: 6px !important;
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

if "task" not in st.session_state:
    st.session_state["task"] = None

# ---------------------------------------------------------------------------
# Landing screen: pick a task
# ---------------------------------------------------------------------------
if st.session_state["task"] is None:
    st.subheader("What would you like to do?")
    cols = st.columns(len(TASKS))
    for col, task_name in zip(cols, TASKS):
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

# ---------------------------------------------------------------------------
# Trip Preference Parsing -- fully wired to the existing pipeline
# ---------------------------------------------------------------------------
if task == "Trip Preference Parsing":
    tab_new, tab_requests = st.tabs(["Submit New Request", "My Requests"])

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

                    backend.run_parsing(
                        request_id, onboarding_bytes,
                        progress_callback=_on_progress,
                        secrets_getter=lambda: st.secrets,
                        salesperson=salesperson,
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

                        backend.run_parsing(
                            row["id"], row["onboarding_bytes"],
                            progress_callback=_on_retry_progress,
                            secrets_getter=lambda: st.secrets,
                            salesperson=row["salesperson"],
                        )
                        progress.empty()
                        st.rerun()
                elif row["status"] == backend.STATUS_PROCESSING:
                    st.info("Still processing -- click Refresh above to check again.")

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
# Weekly Report -- team members submit their update for the current
# period under WEEKLY_REPORT_CATEGORIES (backend.py); "Compile Weekly
# Report" groups every submission for a chosen period into one report.
# Actually emailing it out isn't wired up yet -- the compiled report is
# a copy/download only, ready to paste into an email once a sending
# method (Gmail app password / company SMTP / SendGrid etc.) is picked.
# ---------------------------------------------------------------------------
elif task == "Weekly Report":
    tab_submit, tab_compile = st.tabs(["Submit My Report", "Compile Weekly Report"])

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
            submitter = st.selectbox("Your name", team, key="wr_submitter")
            period = backend.current_period()
            st.caption(f"Reporting period: {backend.period_label(period)}")

            existing = backend.get_weekly_report(submitter, period)
            if existing:
                st.caption("You've already submitted for this period -- editing and submitting again will update it.")

            # Each category starts with zero bullet-point boxes; "+ Add
            # entry" appends one, the X removes one. Bullet widget state
            # is only (re)loaded from the saved report when the selected
            # submitter changes -- not on every rerun -- so add/remove
            # clicks don't wipe unsaved edits.
            load_marker_key = f"wr_loaded_for_{period}"
            if st.session_state.get(load_marker_key) != submitter:
                for key, _ in backend.WEEKLY_REPORT_CATEGORIES:
                    saved_text = existing[key] if existing and existing[key] else ""
                    saved_lines = [ln.strip() for ln in saved_text.split("\n") if ln.strip()]
                    new_ids = []
                    for line in saved_lines:
                        item_id = uuid.uuid4().hex[:8]
                        st.session_state[f"wr_{key}_{item_id}"] = line
                        new_ids.append(item_id)
                    st.session_state[f"wr_{key}_ids"] = new_ids
                st.session_state[load_marker_key] = submitter

            values = {}
            for key, label in backend.WEEKLY_REPORT_CATEGORIES:
                icon = WEEKLY_REPORT_CATEGORY_ICONS.get(key, "•")
                ids_key = f"wr_{key}_ids"
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
                        with st.container(key=f"wr_{key}_row_{item_id}"):
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
                                    key=f"wr_{key}_{item_id}",
                                    placeholder=f"Add a {label.split(' / ')[0].lower()} entry...",
                                )
                            with col_del:
                                if st.button("✕", key=f"wr_{key}_del_{item_id}", help="Remove this entry"):
                                    pending_delete_id = item_id

                    if pending_delete_id is not None:
                        st.session_state[ids_key] = [i for i in st.session_state[ids_key] if i != pending_delete_id]
                        st.session_state.pop(f"wr_{key}_{pending_delete_id}", None)
                        st.rerun()

                    if st.button("+ Add entry", key=f"wr_{key}_add_btn"):
                        new_id = uuid.uuid4().hex[:8]
                        st.session_state[ids_key] = st.session_state[ids_key] + [new_id]
                        st.session_state[f"wr_{key}_{new_id}"] = ""
                        st.rerun()

                bullet_texts = [st.session_state.get(f"wr_{key}_{i}", "").strip() for i in st.session_state[ids_key]]
                values[key] = "\n".join(t for t in bullet_texts if t)

            if st.button("Submit report", key="wr_submit_btn"):
                backend.upsert_weekly_report(submitter, period, values)
                st.toast(f"Saved your report for {submitter}.")
                st.rerun()

    with tab_compile:
        periods = backend.list_weekly_report_periods()
        current = backend.current_period()
        if current not in periods:
            periods = [current] + periods

        period_to_compile = st.selectbox(
            "Period", periods, index=0, key="wr_compile_period", format_func=backend.period_label,
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

        st.info(
            "Sending isn't automated yet -- copy the report above into an email to "
            "the recipients listed, or let me know which email method you want to "
            "use (Gmail app password, company SMTP, or a service like SendGrid) "
            "and this can send automatically with one click."
        )
