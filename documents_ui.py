"""
Document Repository screen -- Phase 2: browse / search / download only.

No upload, replace, archive, category-management, access-management, or
audit-viewer screens here -- those are later phases. Everything below
goes through documents_repo.py's existing public functions exactly like
any other caller would: no raw SQL, and nothing here ever reads
document_blobs or loads file bytes before the user explicitly asks for
them via "Prepare download".

All document data rendered here (titles, descriptions, keywords, file
names, uploader emails) is untrusted text written by whoever uploaded
or described a document. It is shown with st.write/st.caption/
st.dataframe only -- never unsafe_allow_html -- so it is always
rendered as plain text, never interpreted as HTML/markup.

Nothing here uses st.cache_data/st.cache_resource: every query already
takes actor_email and is permission-scoped per call, so caching it
globally (shared across every visitor regardless of who they are)
would leak one person's view of the repository to another. The one
piece of caching that does happen -- per-category document counts in
the left-hand list -- is kept in st.session_state instead, which is
already private to one browser session/user, with an explicit
"Refresh" control.
"""
from datetime import datetime

import streamlit as st

import documents_repo

PAGE_SIZE = 25


# ---------------------------------------------------------------------------
# Small display helpers
# ---------------------------------------------------------------------------
def _human_size(num_bytes) -> str:
    if num_bytes is None:
        return "—"
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"  # pragma: no cover - unreachable, loop always returns first


def _format_timestamp(iso_ts) -> str:
    if not iso_ts:
        return "—"
    try:
        return datetime.fromisoformat(iso_ts).strftime("%Y-%m-%d %H:%M UTC")
    except ValueError:
        return str(iso_ts)


def _secrets_getter():
    return st.secrets


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def render_document_repository(actor_email: str) -> None:
    """Renders the whole Document Repository screen for actor_email.
    Called from app.py's task router; permission is re-checked there
    on every rerun (not just once), so by the time this runs the caller
    has already confirmed actor_email is allowed to be here at all --
    but every documents_repo call below still re-checks its own
    category-level authorization regardless, exactly as it would for
    any other caller.
    """
    secrets_getter = _secrets_getter

    # Confirms the database is reachable and (via get_conn()'s own
    # internal ensure_schema() call) that this module's tables exist,
    # before anything else tries to use them.
    try:
        conn = documents_repo.get_conn()
        conn.close()
    except Exception:
        st.error(
            "The document repository's database isn't reachable right now. "
            "Try again shortly, or contact an administrator if this continues."
        )
        return

    is_admin = documents_repo.is_admin(actor_email, secrets_getter=secrets_getter)

    if is_admin:
        try:
            existing = documents_repo.list_categories(actor_email, secrets_getter=secrets_getter)
            if not existing:
                documents_repo.seed_default_categories(actor_email=actor_email)
        except Exception:
            st.error("Couldn't set up the document repository's default categories.")
            return

    try:
        categories = documents_repo.list_categories(actor_email, secrets_getter=secrets_getter)
    except documents_repo.PermissionDenied as e:
        st.error(str(e))
        return
    except Exception:
        st.error("Couldn't load categories right now.")
        return

    if not categories:
        st.info("No categories are available to you yet; ask an administrator for access.")
        return

    categories_by_id = {c["id"]: c for c in categories}

    col_left, col_main = st.columns([1, 3])

    with col_left:
        selected_category = _render_category_list(actor_email, categories, secrets_getter)

    with col_main:
        docs, filters_active = _render_filters_and_table(
            actor_email, selected_category, secrets_getter,
        )

    st.divider()
    _render_detail_panel(actor_email, docs, categories_by_id, is_admin, secrets_getter)


# ---------------------------------------------------------------------------
# Left column: category list with counts
# ---------------------------------------------------------------------------
def _render_category_list(actor_email: str, categories: list[dict], secrets_getter) -> int | None:
    st.markdown("**Categories**")

    counts_key = f"_doc_repo_category_counts::{actor_email}"
    if st.button("↻ Refresh counts", key="doc_repo_refresh_counts_btn"):
        st.session_state.pop(counts_key, None)

    if counts_key not in st.session_state:
        counts: dict[int, int] = {}
        for cat in categories:
            try:
                rows = documents_repo.list_documents(
                    actor_email, category_id=cat["id"], page_size=10_000, secrets_getter=secrets_getter,
                )
                counts[cat["id"]] = len(rows)
            except Exception:
                counts[cat["id"]] = 0
        st.session_state[counts_key] = counts
    counts = st.session_state[counts_key]

    option_ids: list[int | None] = [None] + [c["id"] for c in categories]
    labels = {None: f"All categories ({sum(counts.values())})"}
    for cat in categories:
        labels[cat["id"]] = f"{cat['name']} ({counts.get(cat['id'], 0)})"

    return st.radio(
        "Category", option_ids, format_func=lambda cid: labels[cid],
        key="doc_repo_selected_category", label_visibility="collapsed",
    )


# ---------------------------------------------------------------------------
# Main column: search/filters + table
# ---------------------------------------------------------------------------
def _render_filters_and_table(actor_email: str, selected_category: int | None, secrets_getter):
    with st.container(border=True):
        col_search, col_type, col_from, col_to = st.columns([2, 1, 1, 1])
        with col_search:
            search = st.text_input(
                "Search", key="doc_repo_search", placeholder="Title, description, or keywords",
            )
        with col_type:
            file_type = st.selectbox(
                "File type", ["Any"] + sorted(documents_repo.ALLOWED_EXTENSIONS),
                key="doc_repo_file_type",
            )
        with col_from:
            date_from = st.date_input("Uploaded from", value=None, key="doc_repo_date_from")
        with col_to:
            date_to = st.date_input("Uploaded to", value=None, key="doc_repo_date_to")

    file_type_value = None if file_type == "Any" else file_type
    date_from_value = date_from.isoformat() if date_from else None
    date_to_value = date_to.isoformat() if date_to else None
    filters_active = bool(search.strip() or file_type_value or date_from_value or date_to_value)

    filter_signature = (selected_category, search, file_type_value, date_from_value, date_to_value)
    if st.session_state.get("_doc_repo_last_filter_sig") != filter_signature:
        st.session_state["doc_repo_page"] = 1
        st.session_state["_doc_repo_last_filter_sig"] = filter_signature
    page = st.session_state.get("doc_repo_page", 1)

    try:
        with st.spinner("Loading documents..."):
            docs = documents_repo.list_documents(
                actor_email,
                category_id=selected_category,
                search=search or None,
                file_type=file_type_value,
                date_from=date_from_value,
                date_to=date_to_value,
                page=page,
                page_size=PAGE_SIZE,
                secrets_getter=secrets_getter,
            )
    except documents_repo.PermissionDenied as e:
        st.error(str(e))
        return [], filters_active
    except Exception:
        st.error("Couldn't load documents right now.")
        return [], filters_active

    if not docs:
        if filters_active:
            st.info("No documents match your search.")
        elif selected_category is not None:
            st.info("No documents in this category yet.")
        else:
            st.info("No documents yet.")
        return [], filters_active

    table_rows = [
        {
            "Title": d["title"],
            "File name": d.get("current_filename") or "—",
            "Format": (d.get("current_extension") or "—").upper(),
            "Size": _human_size(d.get("current_size_bytes")),
            "Uploaded": _format_timestamp(d.get("current_uploaded_at")),
            "Uploader": d.get("current_uploaded_by") or "—",
            "Version": d.get("current_version_no") or "—",
        }
        for d in docs
    ]
    st.dataframe(table_rows, hide_index=True, width="stretch")

    col_prev, col_label, col_next = st.columns([1, 2, 1])
    with col_prev:
        if st.button("◀ Previous", key="doc_repo_prev_btn", disabled=(page <= 1)):
            st.session_state["doc_repo_page"] = page - 1
            st.rerun()
    with col_label:
        st.caption(f"Page {page}")
    with col_next:
        if st.button("Next ▶", key="doc_repo_next_btn", disabled=(len(docs) < PAGE_SIZE)):
            st.session_state["doc_repo_page"] = page + 1
            st.rerun()

    return docs, filters_active


# ---------------------------------------------------------------------------
# Detail panel
# ---------------------------------------------------------------------------
def _render_detail_panel(actor_email: str, docs: list[dict], categories_by_id: dict, is_admin: bool, secrets_getter) -> None:
    st.markdown("**Document detail**")
    if not docs:
        st.caption("Select a category or adjust your search to see documents here.")
        return

    doc_options = [None] + [d["id"] for d in docs]
    doc_titles = {None: "Select a document..."}
    for d in docs:
        doc_titles[d["id"]] = f"{d['title']} ({d.get('current_filename') or 'no file'})"

    selected_doc_id = st.selectbox(
        "View details for", doc_options, format_func=lambda did: doc_titles[did],
        key="doc_repo_selected_doc_id",
    )
    if selected_doc_id is None:
        return

    try:
        with st.spinner("Loading document detail..."):
            detail = documents_repo.get_document_detail(actor_email, selected_doc_id, secrets_getter=secrets_getter)
    except documents_repo.PermissionDenied as e:
        st.error(str(e))
        return
    except Exception:
        st.error("Couldn't load this document's detail right now.")
        return

    if detail is None:
        st.info("This document is no longer available.")
        return

    with st.container(border=True):
        st.markdown(f"#### {detail['title']}")
        category = categories_by_id.get(detail["category_id"])
        st.write(f"**Category:** {category['name'] if category else '—'}")
        st.write(f"**Description:** {detail.get('description') or '—'}")
        st.write(f"**Keywords:** {detail.get('keywords') or '—'}")

        st.markdown("**Version history**")
        versions = sorted(detail["versions"], key=lambda v: v["version_no"], reverse=True)
        # Old versions are only downloadable by admins and contributors
        # (documents_repo.download_version's own rule) -- a viewer sees
        # only the current version here too, so the list never shows an
        # entry they could not actually open.
        visible_versions = versions if is_admin else versions[:1]
        for v in visible_versions:
            st.caption(
                f"v{v['version_no']} -- {v['filename']} -- {_human_size(v['size_bytes'])} -- "
                f"{_format_timestamp(v['uploaded_at'])} by {v['uploaded_by']}"
                + (f" -- {v['change_note']}" if v.get("change_note") else "")
            )
        if not is_admin and len(versions) > 1:
            st.caption(f"{len(versions) - 1} earlier version(s) -- visible to administrators and contributors only.")

        prepare_key = f"_doc_repo_prepared::{selected_doc_id}"
        if st.button("Prepare download", key=f"doc_repo_prepare_btn::{selected_doc_id}"):
            try:
                with st.spinner("Preparing download..."):
                    filename, mime_type, data = documents_repo.download_version(
                        actor_email, selected_doc_id, secrets_getter=secrets_getter,
                    )
                st.session_state[prepare_key] = (filename, mime_type, data)
            except documents_repo.PermissionDenied as e:
                st.error(str(e))
            except documents_repo.StorageIntegrityError:
                st.error(
                    "This file's stored copy failed an integrity check and can't be "
                    "downloaded. Contact an administrator."
                )
            except Exception:
                st.error("Couldn't prepare this download right now.")

        if prepare_key in st.session_state:
            filename, mime_type, data = st.session_state[prepare_key]
            st.download_button(
                f"Download {filename}", data=data, file_name=filename,
                mime=mime_type or "application/octet-stream",
                key=f"doc_repo_download_btn::{selected_doc_id}",
            )
