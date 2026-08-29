"""
Trip Support Intake + Parsing -- combined app
================================================

Landing screen lets the user pick which parsing task they want to run:
  - Trip Preference Parsing: the original onboarding-form pipeline
    (Trip Support Service Preferences PDF -> Operational Information
    Excel). Fully wired to parsing_engine.py via backend.py.
  - Flight Planning Parsing: parses a completed Flight Planning
    Preferences form into the Flight Plan Template. Not wired up yet --
    see the note in that section below.
  - Weekly Report: not yet defined -- see the note in that section below.

Every Trip Preference Parsing submission becomes a row in a local
SQLite database (requests.db, created next to backend.py) with a short
request ID and a status that moves through:

    Submitted -> Processing -> Complete
                             -> Failed  (with an error message + Retry)

If no onboarding form is attached, the request is stored as Submitted
with nothing to parse -- there's no PDF to run through the engine.

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

import streamlit as st

import backend

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
    "Weekly Report": "Generate a weekly report.",
}

st.set_page_config(page_title="Trip Support Intake", page_icon="\U0001F4CB", layout="centered")
st.title("Trip Support Intake")

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

        with st.expander("Manage salesperson list"):
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
        uploaded_file = st.file_uploader(
            "Onboarding Form Upload (optional)",
            type=ALLOWED_UPLOAD_TYPES,
            help="Drag and drop, or browse for, a PDF/Word/image of the completed onboarding form.",
            key="uploaded_file",
        )

        submitted = st.button("Submit")

        if submitted:
            errors = []
            if salesperson is None:
                errors.append("Add a salesperson to the list above before submitting.")
            if payment_method == "Ok to Invoice" and (amount is None or amount <= 0):
                errors.append("Amount is required and must be greater than 0 when 'Ok to Invoice' is selected.")

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
# Flight Planning Parsing -- upload UI is live; extraction rules for this
# document type haven't been defined yet (this is a different source
# document than the Trip Support Service Preferences form, so it needs
# its own field whitelist in parsing_engine.py before this can actually
# process anything). Share a sample of a completed Flight Planning
# Preferences form and the mapping gets built the same way the Trip
# Preference pipeline was.
# ---------------------------------------------------------------------------
elif task == "Flight Planning Parsing":
    st.info(
        "This will parse a completed **Flight Planning Preferences** form and use "
        "it to fill out the Flight Plan Template. The upload is wired up below, but "
        "the extraction rules for this document haven't been defined yet -- share a "
        "sample of a completed Flight Planning Preferences form and this gets built "
        "out the same careful way the Trip Preference pipeline was."
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
# Weekly Report -- not yet defined. Placeholder pending scope: does this
# summarize requests already processed through this app, or parse an
# uploaded weekly document?
# ---------------------------------------------------------------------------
elif task == "Weekly Report":
    st.info(
        "Weekly Report isn't defined yet -- let me know what it should cover "
        "(e.g. a summary of requests processed through this app, or parsing an "
        "uploaded weekly document) and this gets built out."
    )
