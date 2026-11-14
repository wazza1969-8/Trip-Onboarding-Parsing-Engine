"""
Document Repository screen -- Phase 3: browse/search/download (Phase 2)
plus upload, replace, edit metadata, archive/restore, and an
admin/manage-level "show archived" toggle.

No category-management, access-management, or audit-viewer screens
here -- those are still later phases. Everything below goes through
documents_repo.py's existing public functions exactly like any other
caller would: no raw SQL, and nothing here ever reads document_blobs
or loads file bytes before the user explicitly asks for them (a
"Prepare download" click, or an actual Upload/Replace submission).

All document data rendered here (titles, category names, descriptions,
keywords, file names, change notes, uploader emails, and any of those
that might be echoed back inside a documents_repo exception message)
is untrusted text written by whoever uploaded or described a document.
unsafe_allow_html is never used -- but that alone is not enough: every
one of st.markdown/st.write/st.caption/st.success/st.error/st.warning/
st.toast, AND widget labels/help text, AND st.progress's text=, render
a Markdown subset (bold/italic/strikethrough/code/links/images),
confirmed empirically in a real browser -- so a title like
"![x](http://evil/a.png)" would otherwise render as an actual embedded
image. Every such interpolation below goes through _escape_markdown()
first. (The one exception that needs no escaping: st.dataframe's plain
cells, and any st.selectbox/st.radio option label via format_func --
neither renders Markdown.) _escape_markdown() blocks the image/link/
bold/etc. *syntax* outright; a bare URL substring can still be
separately autolinked by the underlying renderer in some of those call
sites (not st.button/st.download_button, which don't autolink at all),
but the escaping corrupts the resulting href, so it never resolves to
the attacker's real domain -- see _escape_markdown()'s own docstring.

UI-level gating here (which tab/button/toggle is shown) is a
convenience, not the security boundary -- documents_repo.py re-checks
every call's authorization itself regardless of what this module
decided to display, exactly as it does for any other caller.

Nothing here uses st.cache_data/st.cache_resource: every query already
takes actor_email and is permission-scoped per call, so caching it
globally (shared across every visitor regardless of who they are)
would leak one person's view of the repository to another. The one
piece of caching that does happen -- per-category document counts in
the left-hand list -- is kept in st.session_state instead, which is
already private to one browser session/user, with an explicit
"Refresh" control.
"""
import hashlib
import io
from datetime import datetime
from pathlib import Path

import streamlit as st
from PIL import Image

import documents_repo

PAGE_SIZE = 25

# Preview (Phase 5): on-demand only, through download_version() so the
# same authorization/archived-status/audit rules as any other download
# apply. Only these mime types get a Preview button at all; anything
# else is download-only.
#
# PDF is deliberately NOT included here: st.pdf exists on the pinned
# streamlit==1.66.1, but calling it raises StreamlitAPIException unless
# the optional `streamlit-pdf` package (the `streamlit[pdf]` extra) is
# also installed -- confirmed by actually calling it, not just checking
# hasattr(st, "pdf"). That is a requirements.txt change, which Phase 5's
# instructions say to stop and ask about rather than make unilaterally.
_PREVIEWABLE_MIME_TYPES = {"image/png", "image/jpeg"}
PREVIEW_MAX_BYTES = 10 * 1024 * 1024

# A pixel-dimension cap for image previews, independent of the byte-size
# cap above -- a small file can still claim enormous width*height in its
# header (a classic "decompression bomb" shape), which would blow up
# memory/CPU if actually decoded. Image.open() only reads the header to
# get .size, so checking this never fully decodes the image -- Pillow
# is already a hard streamlit dependency (streamlit[auth]==1.66.1 pulls
# it in for st.image itself), not a new one added here.
PREVIEW_MAX_PIXELS = 25_000_000

# Rank order for comparing a get_access_level() result against a
# minimum level -- "admin" always passes regardless of position here.
_LEVEL_ORDER = [documents_repo.LEVEL_VIEW, documents_repo.LEVEL_CONTRIBUTE, documents_repo.LEVEL_MANAGE]


def _meets_level(level: str | None, minimum: str) -> bool:
    if level == "admin":
        return True
    if level is None:
        return False
    try:
        return _LEVEL_ORDER.index(level) >= _LEVEL_ORDER.index(minimum)
    except ValueError:
        return False


# Display labels for the Admin tab's role dropdowns -- the module's own
# role values (documents_repo.ROLE_ADMIN etc.) are the actual stored
# strings; these are only what's shown on screen.
_ROLE_LABELS = {
    documents_repo.ROLE_ADMIN: "Administrator",
    documents_repo.ROLE_CONTRIBUTOR: "Authorized Contributor",
    documents_repo.ROLE_VIEWER: "Viewer",
}
_ROLE_VALUES_BY_LABEL = {label: value for value, label in _ROLE_LABELS.items()}

_ACCESS_LEVELS = [documents_repo.LEVEL_VIEW, documents_repo.LEVEL_CONTRIBUTE, documents_repo.LEVEL_MANAGE]


def _is_valid_email_format(email: str) -> bool:
    """Minimal UI-side sanity check before even attempting
    create_doc_user()/grant_category_access() -- not the real
    validation (those functions' own blank/duplicate checks are), just
    a friendlier message than a generic failure for an obviously
    malformed address."""
    email = (email or "").strip()
    return "@" in email and "." in email.rsplit("@", 1)[-1]


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


def _escape_markdown(text) -> str:
    """Escapes Markdown special characters (backslash, backtick,
    asterisk, underscore, braces, brackets, parentheses, #, +, -, .,
    !, |, >, ~, <) so untrusted text can be safely interpolated into
    any Streamlit call that renders Markdown -- confirmed empirically
    (in a real rendered browser, not just AppTest) that this covers
    not just st.markdown/st.write/st.caption/st.success/st.error/
    st.warning/st.toast, but also widget labels/help text and
    st.progress's text=. Backslash is escaped first, so escaping the
    rest doesn't get double-escaped. None becomes "".

    Known limitation: this defeats the explicit image (![alt](url))
    and link ([text](url)) *syntax*, and inline bold/italic/code/etc,
    but not a bare http(s):// URL substring getting separately
    "autolinked" by the underlying renderer in some of those call
    sites (not st.button/st.download_button, which don't autolink at
    all) -- there is no character in the required escape set that
    defeats bare-URL autolinking specifically. In practice this is not
    an open redirect: the backslashes this function inserts get
    URL-encoded into the resulting href (e.g. "example\\.test"
    becomes "example%5C.test"), so any such residual link never
    resolves to the attacker's real domain.
    """
    if text is None:
        return ""
    text = str(text)
    out = text.replace("\\", "\\\\")
    for ch in "`*_{}[]()#+-.!|>~<":
        out = out.replace(ch, f"\\{ch}")
    return out


def _default_title_from_filename(filename: str) -> str:
    """'quarterly report.pdf' -> 'quarterly report' -- just the
    prefilled suggestion in the upload form's editable title field."""
    base = (filename or "").replace("\\", "/").rsplit("/", 1)[-1]
    if "." in base:
        base = base.rsplit(".", 1)[0]
    return base or (filename or "")


# ---------------------------------------------------------------------------
# process_uploads -- pure function, no Streamlit calls, so it is
# unit-testable without driving the file_uploader widget.
# ---------------------------------------------------------------------------
def process_uploads(
    actor_email: str,
    category_id: int,
    uploads: list[dict],
    description: str = "",
    keywords: str = "",
    db_path: Path = documents_repo.DEFAULT_DB_PATH,
    secrets_getter=None,
    store=None,
    progress_callback=None,
) -> list[dict]:
    """Uploads a batch of files one at a time, through
    documents_repo.add_document() -- never raw SQL, never touching
    document_blobs directly. `uploads` is a list of
    {"filename": str, "data": bytes, "title": str} dicts (title may be
    blank, in which case _default_title_from_filename(filename) is
    used). description/keywords apply to every file in the batch.

    One bad file never stops the rest: every item gets its own result
    dict, win or lose --
        {"filename", "title", "success", "document_id", "warning", "error"}
    -- so a caller can show a full per-file result list even when some
    files failed. A failure's "error" is always add_document()'s own
    specific message (empty/oversized/disallowed type/content doesn't
    match extension/macro-enabled/not authorized) or, for anything
    unexpected, a generic message -- never a raw traceback.

    A same-category active document already sharing a file's exact
    SHA-256 is reported as a non-blocking "warning" on that file's
    result -- the upload still proceeds.

    progress_callback(fraction, filename), if given, is called after
    each file finishes (success or failure) -- this function itself
    never touches st.progress or any other Streamlit call, so it can
    be unit-tested directly.
    """
    results = []
    total = len(uploads) or 1
    for i, item in enumerate(uploads, start=1):
        filename = item.get("filename") or ""
        data = item.get("data") or b""
        title = (item.get("title") or "").strip() or _default_title_from_filename(filename)
        result = {
            "filename": filename, "title": title, "success": False,
            "document_id": None, "warning": None, "error": None,
        }

        if data:
            try:
                sha256 = hashlib.sha256(data).hexdigest()
                duplicate = documents_repo.find_duplicate_by_sha256(
                    actor_email, category_id, sha256, db_path=db_path, secrets_getter=secrets_getter,
                )
                if duplicate is not None:
                    result["warning"] = f"An active document already has this exact content: '{duplicate['title']}'."
            except documents_repo.PermissionDenied:
                pass  # the add_document() call below reports the real denial

        try:
            document_id = documents_repo.add_document(
                actor_email, category_id, title, data, filename,
                description=description, keywords=keywords,
                db_path=db_path, secrets_getter=secrets_getter, store=store,
            )
            result["success"] = True
            result["document_id"] = document_id
        except documents_repo.PermissionDenied as e:
            result["error"] = str(e)
        except documents_repo.ValidationError as e:
            result["error"] = str(e)
        except ValueError as e:
            result["error"] = str(e)
        except Exception:
            result["error"] = "Couldn't upload this file right now."

        results.append(result)
        if progress_callback is not None:
            progress_callback(i / total, filename)
    return results


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
                documents_repo.seed_default_categories(actor_email=actor_email, secrets_getter=secrets_getter)
        except Exception:
            st.error("Couldn't set up the document repository's default categories.")
            return

    try:
        categories = documents_repo.list_categories(actor_email, secrets_getter=secrets_getter)
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load categories right now.")
        return

    if not categories:
        st.info("No categories are available to you yet; ask an administrator for access.")
        return

    categories_by_id = {c["id"]: c for c in categories}

    try:
        contributable = documents_repo.categories_with_min_level(
            actor_email, documents_repo.LEVEL_CONTRIBUTE, secrets_getter=secrets_getter,
        )
    except Exception:
        contributable = []
    try:
        manageable = documents_repo.categories_with_min_level(
            actor_email, documents_repo.LEVEL_MANAGE, secrets_getter=secrets_getter,
        )
    except Exception:
        manageable = []
    can_show_archived_toggle = is_admin or bool(manageable)

    tab_names = ["Browse"]
    if contributable:
        tab_names.append("Upload")
    if is_admin:
        tab_names.append("Admin")
    tabs = dict(zip(tab_names, st.tabs(tab_names)))

    with tabs["Browse"]:
        col_left, col_main = st.columns([1, 3])

        with col_left:
            selected_category = _render_category_list(actor_email, categories, secrets_getter)

        with col_main:
            docs, filters_active = _render_filters_and_table(
                actor_email, selected_category, secrets_getter, can_show_archived_toggle,
            )

        st.divider()
        _render_detail_panel(actor_email, docs, categories_by_id, is_admin, secrets_getter)

    if "Upload" in tabs:
        with tabs["Upload"]:
            _render_upload_tab(actor_email, contributable, secrets_getter)

    if "Admin" in tabs:
        with tabs["Admin"]:
            _render_admin_tab(actor_email, secrets_getter)


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
                counts[cat["id"]] = documents_repo.count_documents(
                    actor_email, category_id=cat["id"], secrets_getter=secrets_getter,
                )
            except Exception:
                counts[cat["id"]] = 0
        st.session_state[counts_key] = counts
    counts = st.session_state[counts_key]

    option_ids: list[int | None] = [None] + [c["id"] for c in categories]
    labels = {None: f"All categories ({sum(counts.values())})"}
    for cat in categories:
        labels[cat["id"]] = f"{_escape_markdown(cat['name'])} ({counts.get(cat['id'], 0)})"

    return st.radio(
        "Category", option_ids, format_func=lambda cid: labels[cid],
        key="doc_repo_selected_category", label_visibility="collapsed",
    )


# ---------------------------------------------------------------------------
# Main column: search/filters + table
# ---------------------------------------------------------------------------
def _render_filters_and_table(actor_email: str, selected_category: int | None, secrets_getter, can_show_archived_toggle: bool):
    show_archived = False
    if can_show_archived_toggle:
        show_archived = st.toggle(
            "Show archived", key="doc_repo_show_archived",
            help="Archived documents are only shown if you have manage access (or are an admin) "
                 "on the category they're in.",
        )
    status_filter = None if show_archived else documents_repo.STATUS_ACTIVE

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

    filter_signature = (selected_category, search, file_type_value, date_from_value, date_to_value, show_archived)
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
                status=status_filter,
                page=page,
                page_size=PAGE_SIZE,
                secrets_getter=secrets_getter,
            )
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
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

    try:
        total_count = documents_repo.count_documents(
            actor_email,
            category_id=selected_category,
            search=search or None,
            file_type=file_type_value,
            date_from=date_from_value,
            date_to=date_to_value,
            status=status_filter,
            secrets_getter=secrets_getter,
        )
    except Exception:
        total_count = None

    table_rows = [
        {
            "Title": d["title"],
            "Status": "Archived" if d.get("status") == documents_repo.STATUS_ARCHIVED else "Active",
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
        if total_count is not None:
            start = (page - 1) * PAGE_SIZE + 1
            st.caption(f"Showing {start}-{start + len(docs) - 1} of {total_count}")
        else:
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
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load this document's detail right now.")
        return

    if detail is None:
        st.info("This document is no longer available.")
        return

    category_id = detail["category_id"]
    try:
        access_level = documents_repo.get_access_level(actor_email, category_id, secrets_getter=secrets_getter)
    except Exception:
        access_level = None
    can_contribute = _meets_level(access_level, documents_repo.LEVEL_CONTRIBUTE)
    can_manage = _meets_level(access_level, documents_repo.LEVEL_MANAGE)
    is_archived = detail.get("status") == documents_repo.STATUS_ARCHIVED

    with st.container(border=True):
        status_badge = " (Archived)" if is_archived else ""
        st.markdown(f"#### {_escape_markdown(detail['title'])}{status_badge}")
        category = categories_by_id.get(category_id)
        st.write(f"**Category:** {_escape_markdown(category['name']) if category else '—'}")
        st.write(f"**Description:** {_escape_markdown(detail.get('description')) if detail.get('description') else '—'}")
        st.write(f"**Keywords:** {_escape_markdown(detail.get('keywords')) if detail.get('keywords') else '—'}")

        st.markdown("**Version history**")
        versions = sorted(detail["versions"], key=lambda v: v["version_no"], reverse=True)
        # Old versions are only downloadable by admins and contributors
        # (documents_repo.download_version's own rule) -- anyone else
        # sees only the current version here too, so the list never
        # shows an entry they could not actually open.
        visible_versions = versions if can_contribute else versions[:1]
        for v in visible_versions:
            st.caption(
                f"v{v['version_no']} -- {_escape_markdown(v['filename'])} -- {_human_size(v['size_bytes'])} -- "
                f"{_format_timestamp(v['uploaded_at'])} by {_escape_markdown(v['uploaded_by'])}"
                + (f" -- {_escape_markdown(v['change_note'])}" if v.get("change_note") else "")
            )
        if not can_contribute and len(versions) > 1:
            st.caption(f"{len(versions) - 1} earlier version(s) -- visible to administrators and contributors only.")

        current_version = next((v for v in versions if v["id"] == detail.get("current_version_id")), None)
        if current_version is not None and current_version.get("mime_type") in _PREVIEWABLE_MIME_TYPES:
            if (current_version.get("size_bytes") or 0) > PREVIEW_MAX_BYTES:
                st.caption(
                    f"This file is over {_human_size(PREVIEW_MAX_BYTES)} -- "
                    "use the download option below instead of previewing it."
                )
            else:
                preview_key = f"_doc_repo_preview::{selected_doc_id}"
                if st.button("Preview", key=f"doc_repo_preview_btn::{selected_doc_id}"):
                    try:
                        with st.spinner("Loading preview..."):
                            filename, mime_type, data = documents_repo.download_version(
                                actor_email, selected_doc_id, secrets_getter=secrets_getter,
                            )
                        # Keep at most one previewed image in memory at
                        # a time -- without this, switching between
                        # documents and previewing each would pile up
                        # one full set of decoded bytes per document
                        # for the rest of the session.
                        for stale_key in [k for k in st.session_state if k.startswith("_doc_repo_preview::") and k != preview_key]:
                            del st.session_state[stale_key]
                        st.session_state[preview_key] = (filename, mime_type, data)
                    except documents_repo.PermissionDenied as e:
                        st.error(_escape_markdown(str(e)))
                    except documents_repo.StorageIntegrityError:
                        st.error(
                            "This file's stored copy failed an integrity check and can't be "
                            "previewed. Contact an administrator."
                        )
                    except Exception:
                        st.error("Couldn't load a preview right now.")

                if preview_key in st.session_state:
                    _, preview_mime_type, preview_data = st.session_state[preview_key]
                    if preview_mime_type in ("image/png", "image/jpeg"):
                        try:
                            # Image.open() is lazy -- it reads only
                            # enough of the header to report .size,
                            # never decodes pixel data -- so this check
                            # never fully decodes an image just to
                            # measure it, even a malicious one.
                            with Image.open(io.BytesIO(preview_data)) as probe:
                                width, height = probe.size
                            within_pixel_limit = (width * height) <= PREVIEW_MAX_PIXELS
                        except Exception:
                            within_pixel_limit = False

                        if not within_pixel_limit:
                            st.info(
                                "This image's dimensions are too large to preview safely -- "
                                "use the download option below instead."
                            )
                        else:
                            try:
                                st.image(preview_data)
                            except Exception:
                                st.error("Couldn't render a preview of this file. Try downloading it instead.")

        prepare_key = f"_doc_repo_prepared::{selected_doc_id}"
        if st.button("Prepare download", key=f"doc_repo_prepare_btn::{selected_doc_id}"):
            try:
                with st.spinner("Preparing download..."):
                    filename, mime_type, data = documents_repo.download_version(
                        actor_email, selected_doc_id, secrets_getter=secrets_getter,
                    )
                st.session_state[prepare_key] = (filename, mime_type, data)
            except documents_repo.PermissionDenied as e:
                st.error(_escape_markdown(str(e)))
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
                f"Download {_escape_markdown(filename)}", data=data, file_name=filename,
                mime=mime_type or "application/octet-stream",
                key=f"doc_repo_download_btn::{selected_doc_id}",
            )

        if can_contribute:
            _render_replace_and_edit(actor_email, selected_doc_id, detail, secrets_getter)
        if can_manage:
            _render_archive_restore(actor_email, selected_doc_id, is_archived, secrets_getter)


# ---------------------------------------------------------------------------
# Detail panel: Replace file / Edit details (contribute-or-higher)
# ---------------------------------------------------------------------------
def _render_replace_and_edit(actor_email: str, doc_id: str, detail: dict, secrets_getter) -> None:
    with st.expander("Replace file"):
        # A fresh key per "generation" clears the file_uploader widget
        # after a successful replace (Streamlit has no direct "clear"
        # call for it) -- same approach as the Upload tab's own
        # uploader, but keyed per-document since each document's
        # Replace uploader is independent.
        generation_key = f"_doc_repo_replace_generation::{doc_id}"
        generation = st.session_state.get(generation_key, 0)
        replace_file = st.file_uploader(
            "New file", type=sorted(documents_repo.ALLOWED_EXTENSIONS),
            key=f"doc_repo_replace_file::{doc_id}::{generation}",
        )
        change_note = st.text_input("Change note (optional)", key=f"doc_repo_replace_note::{doc_id}")
        if st.button("Replace", key=f"doc_repo_replace_btn::{doc_id}", disabled=(replace_file is None)):
            try:
                with st.spinner("Replacing..."):
                    documents_repo.replace_document(
                        actor_email, doc_id, replace_file.getvalue(), replace_file.name,
                        change_note=change_note, secrets_getter=secrets_getter,
                    )
                st.session_state[generation_key] = generation + 1
                st.toast("Replaced with a new version.")
                st.rerun()
            except documents_repo.PermissionDenied as e:
                st.error(_escape_markdown(str(e)))
            except documents_repo.ValidationError as e:
                st.error(_escape_markdown(str(e)))
            except Exception:
                st.error("Couldn't replace this file right now.")

    with st.expander("Edit details"):
        new_title = st.text_input("Title", value=detail["title"], key=f"doc_repo_edit_title::{doc_id}")
        new_description = st.text_area(
            "Description", value=detail.get("description") or "", key=f"doc_repo_edit_description::{doc_id}",
        )
        new_keywords = st.text_input(
            "Keywords", value=detail.get("keywords") or "", key=f"doc_repo_edit_keywords::{doc_id}",
        )
        if st.button("Save details", key=f"doc_repo_save_details_btn::{doc_id}"):
            try:
                documents_repo.update_metadata(
                    actor_email, doc_id, title=new_title, description=new_description,
                    keywords=new_keywords, secrets_getter=secrets_getter,
                )
                st.toast("Saved.")
                st.rerun()
            except documents_repo.PermissionDenied as e:
                st.error(_escape_markdown(str(e)))
            except ValueError as e:
                st.error(_escape_markdown(str(e)))
            except Exception:
                st.error("Couldn't save these changes right now.")


# ---------------------------------------------------------------------------
# Detail panel: Archive / Restore (manage-level or admin)
# ---------------------------------------------------------------------------
def _render_archive_restore(actor_email: str, doc_id: str, is_archived: bool, secrets_getter) -> None:
    if is_archived:
        if st.button("Restore", key=f"doc_repo_restore_btn::{doc_id}"):
            try:
                documents_repo.restore_document(actor_email, doc_id, secrets_getter=secrets_getter)
                st.toast("Restored.")
                st.rerun()
            except documents_repo.PermissionDenied as e:
                st.error(_escape_markdown(str(e)))
            except Exception:
                st.error("Couldn't restore this document right now.")
        return

    confirm_key = f"_doc_repo_confirm_archive::{doc_id}"
    if not st.session_state.get(confirm_key):
        if st.button("Archive", key=f"doc_repo_archive_btn::{doc_id}"):
            st.session_state[confirm_key] = True
            st.rerun()
        return

    st.warning("Archive this document? It will be hidden from the normal browse list, but not deleted.")
    col_yes, col_no = st.columns(2)
    with col_yes:
        if st.button("Yes, archive", key=f"doc_repo_confirm_archive_btn::{doc_id}", type="primary"):
            try:
                documents_repo.archive_document(actor_email, doc_id, secrets_getter=secrets_getter)
                st.session_state.pop(confirm_key, None)
                st.toast("Archived.")
                st.rerun()
            except documents_repo.PermissionDenied as e:
                st.session_state.pop(confirm_key, None)
                st.error(_escape_markdown(str(e)))
            except Exception:
                st.session_state.pop(confirm_key, None)
                st.error("Couldn't archive this document right now.")
    with col_no:
        if st.button("Cancel", key=f"doc_repo_cancel_archive_btn::{doc_id}"):
            st.session_state.pop(confirm_key, None)
            st.rerun()


# ---------------------------------------------------------------------------
# Upload tab -- shown only when contributable_categories is non-empty
# (see render_document_repository()).
# ---------------------------------------------------------------------------
def _render_upload_tab(actor_email: str, contributable_categories: list[dict], secrets_getter) -> None:
    category_options = {c["id"]: c["name"] for c in contributable_categories}
    selected_category_id = st.selectbox(
        "Category", list(category_options), format_func=lambda cid: category_options[cid],
        key="doc_repo_upload_category",
    )

    generation = st.session_state.get("_doc_repo_uploader_generation", 0)
    uploader_key = f"doc_repo_uploader_{generation}"
    uploaded_files = st.file_uploader(
        "Files to upload", type=sorted(documents_repo.ALLOWED_EXTENSIONS),
        accept_multiple_files=True, key=uploader_key,
    )

    description = st.text_area(
        "Description (optional -- applies to every file in this batch)",
        key="doc_repo_upload_description",
    )
    keywords = st.text_input(
        "Keywords (optional -- applies to every file in this batch)",
        key="doc_repo_upload_keywords",
    )

    # Keyed on (uploader generation, file position, name, size,
    # category) rather than re-hashing + re-querying the database on
    # every rerun -- recomputes automatically whenever any of those
    # actually change (a new file batch bumps the generation; picking
    # a different category changes the last element of the key).
    dup_cache = st.session_state.setdefault("_doc_repo_duplicate_cache", {})

    # Keyed by the file's position in the list, not its name/size --
    # two selected files that happen to share both would otherwise
    # collide on the same widget key and the same titles[...] entry.
    titles: dict[int, str] = {}
    if uploaded_files:
        st.markdown("**Titles**")
        for i, f in enumerate(uploaded_files):
            title_key = f"doc_repo_upload_title::{uploader_key}::{i}"
            titles[i] = st.text_input(
                f"Title for {_escape_markdown(f.name)}", value=_default_title_from_filename(f.name), key=title_key,
            )

            cache_key = (generation, i, f.name, f.size, selected_category_id)
            if cache_key in dup_cache:
                duplicate = dup_cache[cache_key]
            else:
                duplicate = None
                data = f.getvalue()
                if data:
                    try:
                        sha256 = hashlib.sha256(data).hexdigest()
                        duplicate = documents_repo.find_duplicate_by_sha256(
                            actor_email, selected_category_id, sha256, secrets_getter=secrets_getter,
                        )
                    except documents_repo.PermissionDenied:
                        duplicate = None  # the Upload button below reports the real denial
                dup_cache[cache_key] = duplicate

            if duplicate is not None:
                st.warning(
                    f"'{_escape_markdown(f.name)}': an active document already has this exact content "
                    f"('{_escape_markdown(duplicate['title'])}'). You can still upload it."
                )

    uploading_key = "_doc_repo_upload_in_progress"
    file_count = len(uploaded_files or [])
    upload_clicked = st.button(
        f"Upload {file_count} file(s)", key="doc_repo_upload_btn",
        disabled=(file_count == 0 or st.session_state.get(uploading_key, False)),
    )

    if upload_clicked:
        st.session_state[uploading_key] = True
        progress = st.progress(0.0, text="Starting upload...")

        def _on_progress(fraction: float, filename: str) -> None:
            escaped = _escape_markdown(filename)
            progress.progress(fraction, text=f"Uploaded {escaped}" if fraction >= 1.0 else f"Uploading {escaped}...")

        uploads = [
            {"filename": f.name, "data": f.getvalue(), "title": titles.get(i, "")}
            for i, f in enumerate(uploaded_files)
        ]
        try:
            results = process_uploads(
                actor_email, selected_category_id, uploads,
                description=description, keywords=keywords,
                secrets_getter=secrets_getter, progress_callback=_on_progress,
            )
        except Exception:
            results = [
                {"filename": u["filename"], "title": u["title"], "success": False,
                 "document_id": None, "warning": None, "error": "Couldn't process this upload batch right now."}
                for u in uploads
            ]
        progress.empty()
        st.session_state[uploading_key] = False
        # A fresh key clears the file_uploader widget on the next render
        # (Streamlit has no direct "clear" call for it) -- this also
        # clears the per-file title inputs, since their keys are
        # derived from uploader_key.
        st.session_state["_doc_repo_uploader_generation"] = generation + 1
        st.session_state["_doc_repo_last_upload_results"] = results
        st.rerun()

    last_results = st.session_state.pop("_doc_repo_last_upload_results", None)
    if last_results:
        st.markdown("**Upload results**")
        for r in last_results:
            filename = _escape_markdown(r["filename"])
            if r["success"]:
                message = f"'{filename}' uploaded as '{_escape_markdown(r['title'])}'."
                if r.get("warning"):
                    st.success(message)
                    st.caption(f"Note: {_escape_markdown(r['warning'])}")
                else:
                    st.success(message)
            else:
                st.error(f"'{filename}' failed: {_escape_markdown(r['error'])}")


# ---------------------------------------------------------------------------
# Admin tab -- shown only when documents_repo.is_admin(actor_email) is
# true (checked in render_document_repository() before this tab even
# exists). Every mutation below still goes through documents_repo.py's
# own admin-gated service functions, which re-check authorization
# themselves regardless of this tab's own visibility -- the tab is a
# convenience, not the security boundary.
# ---------------------------------------------------------------------------
def _run_admin_action(action, *args, success_message: str = "Done.", **kwargs) -> None:
    """Calls an admin-gated documents_repo mutation, showing a toast and
    rerunning on success, or a clear error on failure (PermissionDenied/
    ValueError show their own message; anything else gets a generic
    one). Factored out because the Admin tab below has a lot of
    buttons that all follow this exact shape."""
    try:
        action(*args, **kwargs)
        st.toast(success_message)
        st.rerun()
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
    except ValueError as e:
        st.error(_escape_markdown(str(e)))
    except Exception:
        st.error("Couldn't complete this action right now.")


def _render_admin_tab(actor_email: str, secrets_getter) -> None:
    # Only the selected section is rendered -- unlike a plain st.tabs()
    # (which renders every tab's content on every rerun, just hiding the
    # inactive ones with CSS), a segmented_control lets us skip the other
    # four sections' DB round trips entirely until the admin asks for them.
    sub_names = ["Categories", "Users & Roles", "Access Grants", "Archived Records", "Audit Log"]
    selected = st.segmented_control(
        "Admin section",
        sub_names,
        default=sub_names[0],
        key="doc_repo_admin_section",
        label_visibility="collapsed",
    )
    if selected is None:
        selected = sub_names[0]

    if selected == "Categories":
        _render_admin_categories(actor_email, secrets_getter)
    elif selected == "Users & Roles":
        _render_admin_users(actor_email, secrets_getter)
    elif selected == "Access Grants":
        _render_admin_access_grants(actor_email, secrets_getter)
    elif selected == "Archived Records":
        _render_admin_archived_records(actor_email, secrets_getter)
    elif selected == "Audit Log":
        _render_admin_audit_log(actor_email, secrets_getter)


# ---------------------------------------------------------------------------
# Admin: Categories (create, rename, reorder, deactivate/reactivate)
# ---------------------------------------------------------------------------
def _move_category(actor_email: str, ordered_categories: list[dict], from_index: int, to_index: int, secrets_getter) -> None:
    ids = [c["id"] for c in ordered_categories]
    ids[from_index], ids[to_index] = ids[to_index], ids[from_index]
    _run_admin_action(documents_repo.reorder_categories, actor_email, ids, secrets_getter=secrets_getter, success_message="Reordered.")


def _render_admin_categories(actor_email: str, secrets_getter) -> None:
    try:
        categories = documents_repo.list_categories(actor_email, secrets_getter=secrets_getter)
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load categories right now.")
        return
    categories = sorted(categories, key=lambda c: c["sort_order"])

    with st.container(border=True):
        st.markdown("**Create category**")
        new_name = st.text_input("Name", key="doc_repo_admin_new_category_name")
        new_description = st.text_input("Description (optional)", key="doc_repo_admin_new_category_description")
        if st.button("Create category", key="doc_repo_admin_create_category_btn"):
            _run_admin_action(
                documents_repo.create_category, actor_email, new_name,
                description=new_description, secrets_getter=secrets_getter,
                success_message="Category created.",
            )

    st.markdown("**Existing categories**")
    if not categories:
        st.caption("No categories yet.")
        return

    for i, cat in enumerate(categories):
        with st.container(border=True):
            status_label = "Active" if cat["is_active"] else "Inactive"
            st.write(f"**{_escape_markdown(cat['name'])}** ({status_label})")
            if cat.get("description"):
                st.caption(_escape_markdown(cat["description"]))

            col_rename, col_up, col_down, col_status = st.columns([3, 1, 1, 2])
            with col_rename:
                rename_key = f"doc_repo_admin_cat_rename_input::{cat['id']}"
                new_cat_name = st.text_input(
                    "Rename to", value=cat["name"], key=rename_key, label_visibility="collapsed",
                )
                if st.button(
                    "Save name", key=f"doc_repo_admin_cat_rename_btn::{cat['id']}",
                    disabled=(new_cat_name == cat["name"]),
                ):
                    _run_admin_action(
                        documents_repo.rename_category, actor_email, cat["id"], new_cat_name,
                        secrets_getter=secrets_getter, success_message="Renamed.",
                    )
            with col_up:
                if st.button("↑ Up", key=f"doc_repo_admin_cat_up::{cat['id']}", disabled=(i == 0)):
                    _move_category(actor_email, categories, i, i - 1, secrets_getter)
            with col_down:
                if st.button("↓ Down", key=f"doc_repo_admin_cat_down::{cat['id']}", disabled=(i == len(categories) - 1)):
                    _move_category(actor_email, categories, i, i + 1, secrets_getter)
            with col_status:
                if cat["is_active"]:
                    if st.button("Deactivate", key=f"doc_repo_admin_cat_deactivate_btn::{cat['id']}"):
                        _run_admin_action(
                            documents_repo.deactivate_category, actor_email, cat["id"],
                            secrets_getter=secrets_getter, success_message="Deactivated.",
                        )
                else:
                    if st.button("Reactivate", key=f"doc_repo_admin_cat_reactivate_btn::{cat['id']}"):
                        _run_admin_action(
                            documents_repo.reactivate_category, actor_email, cat["id"],
                            secrets_getter=secrets_getter, success_message="Reactivated.",
                        )


# ---------------------------------------------------------------------------
# Admin: Users & Roles
# ---------------------------------------------------------------------------
def _render_admin_users(actor_email: str, secrets_getter) -> None:
    try:
        users = documents_repo.list_doc_users(actor_email, secrets_getter=secrets_getter)
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load users right now.")
        return

    with st.container(border=True):
        st.markdown("**Add user**")
        new_email = st.text_input("Email", key="doc_repo_admin_new_user_email")
        new_role_label = st.selectbox("Role", list(_ROLE_LABELS.values()), key="doc_repo_admin_new_user_role")
        if st.button("Add user", key="doc_repo_admin_add_user_btn"):
            candidate = (new_email or "").strip().lower()
            if not _is_valid_email_format(candidate):
                st.error("That doesn't look like a valid email address.")
            else:
                _run_admin_action(
                    documents_repo.create_doc_user, actor_email, candidate, _ROLE_VALUES_BY_LABEL[new_role_label],
                    secrets_getter=secrets_getter, success_message="User added.",
                )

    st.markdown("**Existing users**")
    if not users:
        st.caption("No users yet.")
        return

    for u in users:
        email = u["email"]
        is_self = (email == actor_email)
        try:
            is_secret_admin = documents_repo.is_admin_via_secret(email, secrets_getter=secrets_getter)
        except Exception:
            is_secret_admin = False
        # Blocks a self-targeted change only when it would actually
        # strip this exact admin's only path to admin access -- not a
        # security boundary (documents_repo.py's own admin check on
        # every call is), just a guard against an easy-to-regret click.
        would_self_lockout = (
            is_self and not is_secret_admin
            and u["role"] == documents_repo.ROLE_ADMIN and u["is_active"]
        )

        with st.container(border=True):
            label = f"**{_escape_markdown(email)}**"
            if is_secret_admin:
                label += " _(also admin via the DOC\\_REPO\\_ADMIN\\_EMAILS secret)_"
            st.write(label)
            st.caption(f"Role: {_ROLE_LABELS.get(u['role'], u['role'])} -- {'Active' if u['is_active'] else 'Inactive'}")
            if is_secret_admin:
                st.caption(
                    "This email's admin access comes from the DOC_REPO_ADMIN_EMAILS secret and "
                    "can't be changed here -- the role/status below only affects its doc_users "
                    "row, which has no effect on that secret-based access."
                )

            col_role, col_status = st.columns(2)
            with col_role:
                role_options = list(_ROLE_LABELS.values())
                current_label = _ROLE_LABELS.get(u["role"], role_options[0])
                new_label = st.selectbox(
                    "Role", role_options,
                    index=role_options.index(current_label),
                    key=f"doc_repo_admin_user_role::{email}",
                )
                new_role_value = _ROLE_VALUES_BY_LABEL[new_label]
                blocks_role_change = would_self_lockout and new_role_value != documents_repo.ROLE_ADMIN
                if st.button(
                    "Update role", key=f"doc_repo_admin_user_update_role_btn::{email}",
                    disabled=(new_role_value == u["role"] or blocks_role_change),
                ):
                    _run_admin_action(
                        documents_repo.update_doc_user_role, actor_email, email, new_role_value,
                        secrets_getter=secrets_getter, success_message="Role updated.",
                    )
                if blocks_role_change:
                    st.caption("You can't remove your own admin access this way -- have another admin do it.")
            with col_status:
                if u["is_active"]:
                    if st.button(
                        "Deactivate", key=f"doc_repo_admin_user_deactivate_btn::{email}",
                        disabled=would_self_lockout,
                    ):
                        _run_admin_action(
                            documents_repo.deactivate_doc_user, actor_email, email,
                            secrets_getter=secrets_getter, success_message="Deactivated.",
                        )
                    if would_self_lockout:
                        st.caption("You can't deactivate your own only admin access this way.")
                else:
                    if st.button("Reactivate", key=f"doc_repo_admin_user_reactivate_btn::{email}"):
                        _run_admin_action(
                            documents_repo.reactivate_doc_user, actor_email, email,
                            secrets_getter=secrets_getter, success_message="Reactivated.",
                        )


# ---------------------------------------------------------------------------
# Admin: per-category access grants
# ---------------------------------------------------------------------------
def _render_admin_access_grants(actor_email: str, secrets_getter) -> None:
    try:
        categories = documents_repo.list_categories(actor_email, secrets_getter=secrets_getter)
    except Exception:
        st.error("Couldn't load categories right now.")
        return
    if not categories:
        st.caption("No categories yet -- create one in the Categories tab first.")
        return

    category_options = {c["id"]: c["name"] for c in categories}
    selected_category_id = st.selectbox(
        "Category", list(category_options), format_func=lambda cid: category_options[cid],
        key="doc_repo_admin_access_category",
    )

    try:
        grants = documents_repo.list_category_access(actor_email, selected_category_id, secrets_getter=secrets_getter)
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load access grants right now.")
        return

    with st.container(border=True):
        st.markdown("**Grant or change access**")
        grant_email = st.text_input("User email", key="doc_repo_admin_grant_email")
        grant_level = st.selectbox("Level", _ACCESS_LEVELS, key="doc_repo_admin_grant_level")
        if st.button("Grant / update access", key="doc_repo_admin_grant_btn"):
            candidate = (grant_email or "").strip().lower()
            if not _is_valid_email_format(candidate):
                st.error("That doesn't look like a valid email address.")
            else:
                _run_admin_action(
                    documents_repo.grant_category_access, actor_email, selected_category_id, candidate, grant_level,
                    secrets_getter=secrets_getter, success_message="Access granted.",
                )

    st.markdown("**Current access**")
    if not grants:
        st.caption("No one has explicit access to this category yet.")
        return

    for g in grants:
        with st.container(border=True):
            st.write(f"**{_escape_markdown(g['subject'])}** -- {g['level']}")
            col_change, col_revoke = st.columns(2)
            with col_change:
                new_level = st.selectbox(
                    "Change level", _ACCESS_LEVELS, index=_ACCESS_LEVELS.index(g["level"]),
                    key=f"doc_repo_admin_access_level::{g['id']}",
                )
                if st.button(
                    "Save level", key=f"doc_repo_admin_access_save_btn::{g['id']}",
                    disabled=(new_level == g["level"]),
                ):
                    _run_admin_action(
                        documents_repo.grant_category_access, actor_email, selected_category_id, g["subject"], new_level,
                        secrets_getter=secrets_getter, success_message="Access level updated.",
                    )
            with col_revoke:
                if st.button("Revoke", key=f"doc_repo_admin_access_revoke_btn::{g['id']}"):
                    _run_admin_action(
                        documents_repo.revoke_category_access, actor_email, selected_category_id, g["subject"],
                        secrets_getter=secrets_getter, success_message="Access revoked.",
                    )


# ---------------------------------------------------------------------------
# Admin: archived records, across every category
# ---------------------------------------------------------------------------
def _render_admin_archived_records(actor_email: str, secrets_getter) -> None:
    try:
        categories = documents_repo.list_categories(actor_email, secrets_getter=secrets_getter)
    except Exception:
        categories = []
    categories_by_id = {c["id"]: c for c in categories}

    page = st.session_state.get("_doc_repo_admin_archived_page", 1)
    try:
        with st.spinner("Loading archived documents..."):
            docs = documents_repo.list_documents(
                actor_email, status=documents_repo.STATUS_ARCHIVED, page=page, page_size=PAGE_SIZE,
                secrets_getter=secrets_getter,
            )
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load archived documents right now.")
        return

    if not docs:
        st.info("No archived documents.")
        return

    try:
        total_count = documents_repo.count_documents(
            actor_email, status=documents_repo.STATUS_ARCHIVED, secrets_getter=secrets_getter,
        )
    except Exception:
        total_count = None

    for d in docs:
        with st.container(border=True):
            category_name = categories_by_id.get(d["category_id"], {}).get("name", "—")
            st.write(f"**{_escape_markdown(d['title'])}** -- {_escape_markdown(category_name)}")
            st.caption(
                f"{_escape_markdown(d.get('current_filename') or '—')} -- "
                f"archived {_format_timestamp(d.get('archived_at'))} by {_escape_markdown(d.get('archived_by') or '—')}"
            )
            if st.button("Restore", key=f"doc_repo_admin_restore_btn::{d['id']}"):
                _run_admin_action(
                    documents_repo.restore_document, actor_email, d["id"],
                    secrets_getter=secrets_getter, success_message="Restored.",
                )

    col_prev, col_label, col_next = st.columns([1, 2, 1])
    with col_prev:
        if st.button("◀ Previous", key="doc_repo_admin_archived_prev_btn", disabled=(page <= 1)):
            st.session_state["_doc_repo_admin_archived_page"] = page - 1
            st.rerun()
    with col_label:
        if total_count is not None:
            start = (page - 1) * PAGE_SIZE + 1
            st.caption(f"Showing {start}-{start + len(docs) - 1} of {total_count}")
        else:
            st.caption(f"Page {page}")
    with col_next:
        if st.button("Next ▶", key="doc_repo_admin_archived_next_btn", disabled=(len(docs) < PAGE_SIZE)):
            st.session_state["_doc_repo_admin_archived_page"] = page + 1
            st.rerun()


# ---------------------------------------------------------------------------
# Admin: audit log viewer (read-only, newest first, paginated, filterable)
# ---------------------------------------------------------------------------
def _render_admin_audit_log(actor_email: str, secrets_getter) -> None:
    col_actor, col_action, col_from, col_to = st.columns(4)
    with col_actor:
        actor_filter = st.text_input("Actor contains", key="doc_repo_admin_audit_actor_filter")
    with col_action:
        action_filter = st.text_input("Action contains", key="doc_repo_admin_audit_action_filter")
    with col_from:
        date_from = st.date_input("From", value=None, key="doc_repo_admin_audit_date_from")
    with col_to:
        date_to = st.date_input("To", value=None, key="doc_repo_admin_audit_date_to")

    date_from_value = date_from.isoformat() if date_from else None
    date_to_value = date_to.isoformat() if date_to else None

    filter_signature = (actor_filter, action_filter, date_from_value, date_to_value)
    if st.session_state.get("_doc_repo_admin_audit_last_filter_sig") != filter_signature:
        st.session_state["_doc_repo_admin_audit_page"] = 1
        st.session_state["_doc_repo_admin_audit_last_filter_sig"] = filter_signature
    page = st.session_state.get("_doc_repo_admin_audit_page", 1)

    try:
        with st.spinner("Loading audit log..."):
            rows = documents_repo.list_audit_log(
                actor_email, actor_filter=actor_filter or None, action_filter=action_filter or None,
                date_from=date_from_value, date_to=date_to_value,
                page=page, page_size=PAGE_SIZE, secrets_getter=secrets_getter,
            )
    except documents_repo.PermissionDenied as e:
        st.error(_escape_markdown(str(e)))
        return
    except Exception:
        st.error("Couldn't load the audit log right now.")
        return

    if not rows:
        st.info("No audit log entries match these filters.")
        return

    try:
        total_count = documents_repo.count_audit_log(
            actor_email, actor_filter=actor_filter or None, action_filter=action_filter or None,
            date_from=date_from_value, date_to=date_to_value, secrets_getter=secrets_getter,
        )
    except Exception:
        total_count = None

    # st.dataframe cells are always plain text (confirmed: no Markdown
    # rendering there, unlike nearly everything else in this module) --
    # so, unlike elsewhere, these are deliberately NOT passed through
    # _escape_markdown(): doing so would just show a literal stray
    # backslash for no protective benefit.
    table_rows = [
        {
            "Time": _format_timestamp(r.get("ts")),
            "Actor": r.get("actor_email") or "—",
            "Action": r.get("action") or "—",
            "Document": r.get("document_id") or "—",
            "Category": r.get("category_id") if r.get("category_id") is not None else "—",
            "Detail": r.get("detail") or "—",
        }
        for r in rows
    ]
    st.dataframe(table_rows, hide_index=True, width="stretch")

    col_prev, col_label, col_next = st.columns([1, 2, 1])
    with col_prev:
        if st.button("◀ Previous", key="doc_repo_admin_audit_prev_btn", disabled=(page <= 1)):
            st.session_state["_doc_repo_admin_audit_page"] = page - 1
            st.rerun()
    with col_label:
        if total_count is not None:
            start = (page - 1) * PAGE_SIZE + 1
            st.caption(f"Showing {start}-{start + len(rows) - 1} of {total_count}")
        else:
            st.caption(f"Page {page}")
    with col_next:
        if st.button("Next ▶", key="doc_repo_admin_audit_next_btn", disabled=(len(rows) < PAGE_SIZE)):
            st.session_state["_doc_repo_admin_audit_page"] = page + 1
            st.rerun()
