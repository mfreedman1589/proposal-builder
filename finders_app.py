"""Premion Finder -- the finders-only entry point for DC sellers.

    streamlit run finders_app.py

Two pages and nothing else: the Audience finder and the Case study finder,
the SAME implementation the proposal builder uses (finders.py). Read-only by
construction -- search, filter, view, download. Nothing here can add, edit,
delete or re-tag anything, because none of that code is imported: this file
imports no builder module at all (app.py, assembly, the report builder),
which tests/test_finders_app.py enforces.

Its own password (FINDERS_PASSWORD), the same name picker and "Report an
issue" as the main app (reports are tagged as coming from this app), and
every search and download is logged with the seller's name for the main
app's Admin -> Usage page.
"""
import streamlit as st

st.set_page_config(page_title="Premion Finder", page_icon="🔎", layout="wide")

import app_shell                                   # noqa: E402
import db                                          # noqa: E402
import finders                                     # noqa: E402

APP_NAME = "Premion Finder"
APP_TAG = "finders"
AUDIENCE_PAGE = "Audience finder"
CASE_STUDY_PAGE = "Case study finder"
PAGES = [AUDIENCE_PAGE, CASE_STUDY_PAGE]
FEEDBACK_CATEGORIES = ["Audiences", "Case studies", "Other"]


def log_usage(event, detail):
    """Every search and download, with who did it (Admin -> Usage)."""
    db.log_finder_usage(event, app_shell.current_user(), detail, app=APP_TAG)


def capture_state(page):
    """What a "Report an issue" from this app carries: which app, which
    page, and what the seller was searching -- never the builder's state,
    which doesn't exist here."""
    return {
        "app": APP_TAG,
        "page": page,
        "build_stamp": app_shell.BUILD_STAMP,
        "user": app_shell.current_user(),
        "audience_search": st.session_state.get("finder_search"),
        "audience_category": st.session_state.get("finder_category"),
        "audience_suggest": st.session_state.get("finder_suggest_input"),
        "case_study_search": st.session_state.get("vault_search"),
        "case_study_suggest": st.session_state.get("cs_suggest_description"),
        "last_claude_failure": st.session_state.get("last_claude_failure"),
    }


def render_audience_page():
    st.header(AUDIENCE_PAGE)
    st.caption("**How to use:** search the catalog by name or category, or switch to "
               "**Suggest** and describe your client — Claude recommends the segments that "
               "fit. \"RFP\" segments can be booked directly; \"Custom\" ones need a request.")
    finders.render_audience_finder(None, on_event=log_usage)


def render_case_study_page():
    st.header(CASE_STUDY_PAGE)
    st.caption("**How to use:** filter by vertical or product, or switch to **Suggest** and "
               "describe your client. Open a case study and download it as a .pptx or a PDF "
               "to attach to an email.")
    rows, warning = db.fetch_case_studies(active_only=True)
    if warning:
        st.warning(warning)
        return
    if not rows:
        st.info("No case studies are available yet.")
        return
    browse_tab, suggest_tab = st.tabs(["Browse", "Suggest"])
    with browse_tab:
        finders.render_case_study_browser(rows, show_inactive_toggle=False, on_event=log_usage)
    with suggest_tab:
        finders.render_case_study_suggest(rows, on_event=log_usage)


def main():
    app_shell.render_dev_banner()
    if not app_shell.password_gate(
            APP_NAME, "FINDERS_PASSWORD",
            caption="Look up Premion audience segments and case studies."):
        return
    if not app_shell.identity_step(
            APP_NAME,
            "Who's using the finder? Pick your name, or add it if it's not there yet — it "
            "just helps us see what's useful."):
        return

    st.sidebar.title(APP_NAME)
    page = st.sidebar.radio("Page", PAGES, key="finders_page", label_visibility="collapsed")
    st.sidebar.divider()
    app_shell.render_identity_sidebar()
    app_shell.feedback_popover(
        capture_state, FEEDBACK_CATEGORIES, f"Finders app · {page}",
        caption="Something wrong, confusing, or missing? This attaches the page you're on "
                "and what you searched (never a screenshot).")
    st.sidebar.caption(f"Build {app_shell.BUILD_STAMP}")

    if page == AUDIENCE_PAGE:
        render_audience_page()
    else:
        render_case_study_page()


if __name__ == "__main__":
    main()
