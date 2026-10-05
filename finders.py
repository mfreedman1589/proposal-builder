"""The Audience finder and the Case study finder -- ONE implementation, two
entry points: the proposal builder (app.py, which adds its own builder and
edit controls through the hooks below) and the finders-only app
(finders_app.py, which passes none of them).

Read-only by construction: nothing in this module writes a case study, a
tag, an audience or a vault row. The only writes it can cause are the ones
its caller hands it as callbacks (`add_controls`, `row_controls`,
`on_event`) -- the builder's AND/OR/New group buttons and the vault editor
live in app.py and never reach the finders app, which isn't a matter of
hiding a button. tests/test_finders_app.py checks both: this module calls
no db write, and the finders app imports nothing from the builder.

Moved out of app.py on 2026-10-05; widget keys are unchanged, so session
state and every existing test that drives these pages by key still work.
"""
import io
from functools import partial

import streamlit as st

import db
from audience_catalog import all_categories, category_matches, load_audience_catalog
from catalog_shared import (CASE_STUDY_PRODUCT_TAGS, PPTX_MIME, VERTICALS, _valid_tags,
                            build_case_study_suggest_prompt, call_claude_audience_suggest,
                            case_study_filename, prioritize_catalog)
from claude_client import _call_claude_json

PDF_MIME = "application/pdf"


def _emit(on_event, event, detail):
    """Fire a usage hook without ever letting it break the page."""
    if on_event is None:
        return
    try:
        on_event(event, detail)
    except Exception as exc:                                     # noqa: BLE001
        print(f"[finders] usage hook failed: {type(exc).__name__}: {exc}", flush=True)


def _emit_once(on_event, event, detail, state_key):
    """Search boxes rerun on every keystroke commit; log a search only when
    what was searched actually changed, so one search is one event."""
    signature = repr(sorted(detail.items()))
    if st.session_state.get(state_key) == signature:
        return
    st.session_state[state_key] = signature
    _emit(on_event, event, detail)


# ---------------------------------------------------------------------------
# Audience finder
# ---------------------------------------------------------------------------

def render_audience_finder(vertical_key=None, *, header=None, add_controls=None,
                           existing_custom=0, on_event=None):
    """Browse/search the audience catalog, or describe a client and let
    Claude suggest segments.

    `header()` renders above the mode picker (the builder's own messages and
    booking evidence). `add_controls(cols, segment, source)` renders the
    builder's AND / OR / New group buttons into `cols` for one row -- source
    is "browse" or "suggest", which keeps their widget keys distinct.
    `existing_custom` is how many custom audiences the caller's proposal
    already holds (0 outside a proposal). `on_event(event, detail)` is the
    usage hook. None of these is passed by the finders app.

    `vertical_key` only sorts vertical-relevant categories to the front and
    seeds Claude's slice; it never filters anything out."""
    catalog = load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
    vertical_hint = vertical_key if vertical_key and vertical_key != "none" else None

    if catalog.empty:
        st.info("The audience catalog is empty -- it couldn't be loaded from Supabase and "
                "the local brochure isn't available.")
        return

    if header is not None:
        header()

    mode = st.radio("Mode", ["Browse / search", "Suggest"], horizontal=True, key="finder_mode")

    if mode == "Browse / search":
        # all_categories expands a comma-joined multi-category cell into its
        # individual values ("MOVERS, LIFESTAGE" offers both "MOVERS" and
        # "LIFESTAGE" as picks, never a combined "MOVERS, LIFESTAGE" option
        # nobody would type), and category_matches is the membership test
        # that agrees with it -- a segment picked under EITHER of its own
        # categories, matched once, never duplicated in the results below.
        cat_options = ["All"] + all_categories(catalog)
        picked_cat = st.selectbox("Category", cat_options, key="finder_category")
        search_text = st.text_input("Search by name", key="finder_search")

        filtered = catalog
        if picked_cat != "All":
            filtered = filtered[filtered["category"].apply(lambda c: category_matches(c, [picked_cat]))]
        if search_text.strip():
            filtered = filtered[filtered["segment"].str.contains(search_text.strip(), case=False, na=False)]
        total_matches = len(filtered)
        if search_text.strip() or picked_cat != "All":
            _emit_once(on_event, "audience_search",
                       {"query": search_text.strip(), "category": picked_cat,
                        "matches": total_matches}, "_finder_last_search_logged")
        filtered = prioritize_catalog(filtered, vertical_hint).head(50)
        if vertical_hint and picked_cat == "All":
            st.caption(f"{total_matches} segment(s) match; showing 50 "
                       f"(categories relevant to the selected vertical first, then by impressions)")
        else:
            st.caption(f"{total_matches} segment(s) match; showing 50 (by impressions)")

        widths = [3, 1.3, 2, 1.2, 1, 0.7, 0.7, 1] if add_controls else [3, 1.3, 2, 1.2, 1]
        labels = ["Segment", "Category", "Subcategory", "Status", "Impressions", "", "", ""]
        for col, label in zip(st.columns(widths), labels):
            col.caption(f"**{label}**")
        for _, row in filtered.iterrows():
            cols = st.columns(widths)
            cols[0].write(row["segment"])
            cols[1].write(row["category"])
            # A blank subcategory comes back from the catalog as NaN, which is
            # truthy -- `or "--"` alone rendered a literal "nan".
            sub = row["subcategory"]
            cols[2].write(sub if isinstance(sub, str) and sub.strip() else "--")
            cols[3].write("RFP" if row["rfp_selectable"] else "Custom")
            cols[4].write(f"{row['impressions']:,}")
            if add_controls:
                add_controls(cols, row["segment"], "browse")

    else:
        description = st.text_area("Describe the client or campaign", key="finder_suggest_input", height=100)
        if st.button("Suggest audiences"):
            if not description.strip():
                st.warning("Describe the client or campaign first.")
            else:
                _emit(on_event, "audience_suggest", {"description": description.strip()[:500]})
                with st.spinner("Asking Claude for audience recommendations..."):
                    recs, unmatched, error = call_claude_audience_suggest(description, vertical_hint)
                if error:
                    st.error(error)
                    st.session_state["finder_suggestions"] = None
                else:
                    st.session_state["finder_suggestions"] = recs
                    st.session_state["finder_suggestions_unmatched"] = unmatched

        suggestions = st.session_state.get("finder_suggestions")
        if suggestions:
            unmatched = st.session_state.get("finder_suggestions_unmatched") or []
            if unmatched:
                st.caption(f"Dropped {len(unmatched)} recommended name(s) not found in the catalog: "
                           + ", ".join(unmatched))

            # DISTINCT custom segments across the caller's own groups, not a
            # row count -- a custom segment reused across two groups isn't
            # double-counted (0 outside a proposal).
            rec_custom = sum(1 for r in suggestions if not rfp_map.get(r["segment"], True))
            if existing_custom + rec_custom > 1:
                st.warning(
                    f"This recommendation set includes {rec_custom} custom (non-RFP-selectable) audience(s), "
                    f"and your targeting groups already have {existing_custom} -- only one custom audience is "
                    f"allowed per campaign. Review before adding all of them.")

            widths = [2.6, 2.6, 1.3, 1, 1, 0.7, 0.7, 1] if add_controls else [2.6, 2.6, 1.3, 1, 1]
            labels = ["Segment", "Rationale", "Category", "Status", "Impressions", "", "", ""]
            for col, label in zip(st.columns(widths), labels):
                col.caption(f"**{label}**")
            for rec in suggestions:
                seg = rec["segment"]
                match = catalog[catalog["segment"] == seg]
                if match.empty:
                    continue
                cat_row = match.iloc[0]
                cols = st.columns(widths)
                cols[0].write(seg)
                cols[1].write(rec.get("rationale", ""))
                cols[2].write(cat_row["category"])
                cols[3].write("RFP" if cat_row["rfp_selectable"] else "Custom")
                cols[4].write(f"{cat_row['impressions']:,}")
                if add_controls:
                    add_controls(cols, seg, "suggest")


# ---------------------------------------------------------------------------
# Case study finder
# ---------------------------------------------------------------------------

def case_study_pdf_filename(row):
    return case_study_filename(row)[:-len(".pptx")] + ".pdf"


@st.cache_data(show_spinner=False)
def _pdf_from_images(image_paths):
    """One PDF page per rendered slide image -- what a seller attaches to an
    email. Built from the same pre-rendered images decks use, so it looks
    exactly like the case study; no PowerPoint involved."""
    from PIL import Image
    pages = [Image.open(path).convert("RGB") for path in image_paths]
    buffer = io.BytesIO()
    pages[0].save(buffer, format="PDF", save_all=True, append_images=pages[1:], resolution=200)
    return buffer.getvalue()


def case_study_pdf_bytes(row):
    """PDF bytes for a case study that has rendered slide images, else None
    (a case study still going in slide by slide has no images to build from
    yet -- the nightly merge renders them)."""
    if db.case_study_render_path(row) != "images":
        return None
    images = db.case_study_images(row["id"], tuple(row["slide_images"]))
    return _pdf_from_images(tuple(images)) if images else None


def render_case_study_download(row, key_prefix, on_event=None):
    """Download buttons for one case study -- the .pptx, and a PDF when its
    slides have been rendered -- fetched on demand.

    Deliberately not eager. st.download_button needs the file's bytes at
    render time, so drawing one per row would pull every visible case study
    out of storage (~42MiB across the vault) whether or not anyone clicked.
    Instead: a file already in the local cache -- from an earlier click, or
    because a proposal included it -- gets real download buttons straight
    away, and anything else takes one click to fetch first. Either way
    db.case_study_file caches it, so it's never fetched from Supabase twice.
    """
    slot = f"{key_prefix}_{row['id']}"
    path = db.case_study_cached_path(row["id"], row["storage_path"])

    if path is None:
        if not st.button("Get the case study", key=f"{slot}_fetch",
                         help="Fetches it from the vault, then offers it as a .pptx (and a PDF "
                              "when available) to download."):
            return
        try:
            path = db.case_study_file(row["id"], row["storage_path"])
        except Exception as exc:                                 # noqa: BLE001
            st.warning(f"Couldn't fetch that case study ({db.describe_error(exc)}).")
            return

    title = row.get("title") or row.get("filename")
    pptx_col, pdf_col = st.columns(2)
    with open(path, "rb") as handle:
        pptx_col.download_button(
            "⬇ Download .pptx", data=handle.read(), file_name=case_study_filename(row),
            mime=PPTX_MIME, key=f"{slot}_download",
            on_click=partial(_emit, on_event, "case_study_download",
                             {"case_study": title, "format": "pptx"}))
    try:
        pdf = case_study_pdf_bytes(row)
    except Exception as exc:                                     # noqa: BLE001
        print(f"[finders] PDF for {title!r} failed: {type(exc).__name__}: {exc}", flush=True)
        pdf = None
    if pdf:
        pdf_col.download_button(
            "⬇ Download PDF", data=pdf, file_name=case_study_pdf_filename(row), mime=PDF_MIME,
            key=f"{slot}_download_pdf",
            on_click=partial(_emit, on_event, "case_study_download",
                             {"case_study": title, "format": "pdf"}))


def report_source_labels(rows):
    """{report_id: "generated from a report — <period>[, <client>]"} for
    every case study `source_report_id` actually present in `rows` -- one
    fetch, not one per row. A report id a case study names but that no
    longer resolves degrades to a bare "generated from a report" rather than
    a crash or a blank line."""
    report_ids = {row["source_report_id"] for row in rows if row.get("source_report_id")}
    if not report_ids:
        return {}
    reports, warning = db.fetch_attribution_reports()
    if warning or not reports:
        return {}
    advertisers, _w = db.fetch_advertisers(active_only=False)
    by_advertiser_id = {a["id"]: a for a in (advertisers or [])}
    labels = {}
    for report in reports:
        if report["id"] not in report_ids:
            continue
        headline = (report.get("report_json") or {}).get("headline_facts") or {}
        start, end = headline.get("period_start"), headline.get("period_end")
        period = f"{start} to {end}" if (start or end) else None
        advertiser = by_advertiser_id.get(report.get("advertiser_id"))
        parts = [p for p in (period, advertiser.get("canonical_name") if advertiser else None) if p]
        labels[report["id"]] = ("📊 generated from a report — " + ", ".join(parts) if parts
                                else "📊 generated from a report")
    return labels


def render_case_study_browser(rows, *, row_controls=None, show_inactive_toggle=True,
                              on_event=None, on_rows=None):
    """Browse and filter the vault, with downloads.

    `row_controls(row, key, vertical_labels)` renders the builder's own
    editor for one case study (title/summary/tags, Save, Deactivate) inside
    its expander -- app.py passes it; the finders app doesn't, and this
    function has no write of its own. `show_inactive_toggle=False` keeps
    deactivated case studies out entirely (the finders app). `on_rows(rows)`
    lets the caller see the full list (app.py logs the render queue)."""
    vertical_labels = {v: k for k, v in VERTICALS.items() if v != "none"}
    source_labels = report_source_labels(rows)
    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        filter_verticals = st.multiselect("Vertical", list(vertical_labels),
                                          format_func=lambda v: vertical_labels[v],
                                          key="vault_filter_verticals")
    with col2:
        filter_products = st.multiselect("Product", CASE_STUDY_PRODUCT_TAGS,
                                         key="vault_filter_products")
    with col3:
        query = st.text_input("Search title or summary", key="vault_search").strip().lower()
    show_inactive = (st.checkbox("Include deactivated", key="vault_show_inactive")
                     if show_inactive_toggle else False)

    def matches(row):
        if not show_inactive and not row.get("active", True):
            return False
        if filter_verticals and not set(filter_verticals) & set(row.get("verticals") or []):
            return False
        if filter_products and not set(filter_products) & set(row.get("products") or []):
            return False
        if query and query not in f"{row.get('title') or ''} {row.get('summary') or ''}".lower():
            return False
        return True

    shown = [r for r in rows if matches(r)]
    pool = rows if show_inactive_toggle else [r for r in rows if r.get("active", True)]
    st.caption(f"{len(shown)} of {len(pool)} case studies")
    if query or filter_verticals or filter_products:
        _emit_once(on_event, "case_study_search",
                   {"query": query, "verticals": sorted(filter_verticals),
                    "products": sorted(filter_products), "matches": len(shown)},
                   "_vault_last_search_logged")
    if on_rows is not None:
        on_rows(rows)

    for row in shown:
        state = "" if row.get("active", True) else "  ·  deactivated"
        # How it goes into a deck (images vs. copied slides) is the builder's
        # concern -- shown only where the editor is (row_controls), never to
        # a finders-app seller, for whom it means nothing.
        route = ""
        if row_controls is not None:
            route = ("  ·  🖼 images" if db.case_study_render_path(row) == "images"
                     else "  ·  📄 copied")
        # Visible on the collapsed header, not just inside -- a seller
        # browsing for a pitch should see which ones come from real
        # attribution data without opening every expander (2026-09-15).
        source_marker = "  ·  📊 from a report" if row.get("source_report_id") else ""
        with st.expander(f"{row['title'] or row['filename']}{route}{state}{source_marker}",
                         expanded=False):
            st.caption(f"{row.get('summary') or '_no summary_'}")
            if row.get("source_report_id"):
                st.caption(source_labels.get(row["source_report_id"], "📊 generated from a report"))
            verticals = ", ".join(vertical_labels.get(v, v) for v in _valid_tags(
                row.get("verticals"), vertical_labels)) or "no vertical tags"
            st.caption(f"{verticals} · added by {row.get('added_by') or 'unknown'} · "
                       f"{str(row.get('date_added'))[:10]}")
            render_case_study_download(row, "browse", on_event)
            if row_controls is not None:
                row_controls(row, f"vault_{row['id']}", vertical_labels)


def render_case_study_suggest(rows, on_event=None):
    """Describe a client and Claude picks from the ACTIVE vault, with a
    one-line reason each. Read-only: recommendations and downloads only."""
    st.caption("Describe the client or campaign and Claude picks from the vault, "
               "with a one-line reason for each.")
    active = [r for r in rows if r.get("active", True)]
    description = st.text_area(
        "Client / campaign", height=110, key="cs_suggest_description",
        placeholder="Regional HVAC company, wants to drive service calls in shoulder season, "
                    "competing against national franchises")
    if st.button("Suggest case studies", type="primary"):
        if not description.strip():
            st.warning("Describe the client first.")
        elif not active:
            st.warning("No active case studies to choose from.")
        else:
            _emit(on_event, "case_study_suggest", {"description": description.strip()[:500]})
            with st.spinner("Reading the vault..."):
                result, error = _call_claude_json(
                    build_case_study_suggest_prompt(description, active),
                    label="case_study_suggest")
            if error:
                st.error(error)
            else:
                st.session_state["cs_suggestions"] = result.get("recommendations", [])

    suggestions = st.session_state.get("cs_suggestions")
    if suggestions is None:
        return
    by_id = {r["id"]: r for r in active}
    # Claude is given the exact ids, but an invented one would otherwise show
    # as a blank row -- same validation bar as audience segments.
    valid = [s for s in suggestions if s.get("id") in by_id]
    if not valid:
        st.info("Nothing in the vault fits that description well enough to recommend.")
        return
    st.success(f"{len(valid)} case study(ies) recommended:")
    for suggestion in valid:
        row = by_id[suggestion["id"]]
        st.markdown(f"**{row['title']}** — {suggestion.get('reason', '')}")
        st.caption(f"{row.get('summary') or ''}  ·  "
                   f"{', '.join(row.get('verticals') or []) or 'no vertical tags'}")
        render_case_study_download(row, "suggest", on_event)
        st.divider()
