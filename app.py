"""
Trip Support Intake + Parsing -- combined app
================================================

Supersedes the standalone intake_form.py / intake_form.html for real use:
this app wires the intake form directly to parsing_engine.py (via
backend.py) so a submitted request actually gets parsed, and gives the
user a place to come back and check status, download the result, or see
why it failed.

Every submission becomes a row in a local SQLite database (requests.db,
created next to backend.py) with a short request ID and a status that
moves through:

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

SALESPEOPLE = [
    "Jane Smith",
    "John Doe",
    "Alex Johnson",
    "Morgan Lee",
]

PAYMENT_METHODS = ["Credit Card", "Ok to Invoice"]
MAX_ACCT_NBR_LEN = 20
ALLOWED_UPLOAD_TYPES = ["pdf", "doc", "docx", "png", "jpg", "jpeg"]

st.set_page_config(page_title="Trip Support Intake", page_icon="\U0001F4CB", layout="centered")
st.title("Trip Support Intake")

tab_new, tab_requests = st.tabs(["Submit New Request", "My Requests"])

with tab_new:
    with st.form("intake_form", clear_on_submit=True):
        salesperson = st.selectbox("Salesperson", SALESPEOPLE)

        acct_nbr = st.text_input(
            "Jeppesen Acct Nbr:",
            max_chars=MAX_ACCT_NBR_LEN,
            help=f"Max {MAX_ACCT_NBR_LEN} characters.",
        )

        payment_method = st.radio("Payment Method", PAYMENT_METHODS, horizontal=True)

        amount = None
        notes = ""
        if payment_method == "Ok to Invoice":
            amount = st.number_input(
                "Amount (required)", min_value=0.0, step=0.01, format="%.2f",
                help="Required when 'Ok to Invoice' is selected.",
            )
            notes = st.text_area("Notes (optional)")

        uploaded_file = st.file_uploader(
            "Onboarding Form Upload (optional)",
            type=ALLOWED_UPLOAD_TYPES,
            help="PDF, Word, or image of the completed onboarding form.",
        )

        submitted = st.form_submit_button("Submit")

    if submitted:
        errors = []
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
                    )
                    progress.empty()
                    st.rerun()
            elif row["status"] == backend.STATUS_PROCESSING:
                st.info("Still processing -- click Refresh above to check again.")
