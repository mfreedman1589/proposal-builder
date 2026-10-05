"""The app-shell building blocks both entry points share -- the proposal
builder (app.py) and the finders-only app (finders_app.py): the build stamp,
local test mode, the dev-deployment banner, and who's signed in. Moved out of
app.py (2026-10-05) so the finders app gets the same shell without importing
the builder. app.py re-imports every name here; a test that patches what
these functions read (their `st`) patches THIS module.
"""
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import streamlit as st

import db

def _read_build_stamp():
    """Short git SHA + commit time of the code THIS process is running, read
    once at import. Answers "am I running current code" at a glance --
    three separate live investigations this week each ended at "probably a
    stale process," and each cost more than this check would have. A warm
    process that predates a fix keeps reporting the SHA it started with,
    which is exactly the tell that's needed; restarting the process is what
    changes it. Never raises -- falls back to a plain label if git isn't on
    PATH or this checkout has no history, same fallback discipline as every
    other loader in this app.
    """
    try:
        repo_dir = Path(__file__).resolve().parent
        sha = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"], cwd=repo_dir,
            capture_output=True, text=True, timeout=5, check=True,
        ).stdout.strip()
        commit_time = subprocess.run(
            ["git", "log", "-1", "--format=%cI"], cwd=repo_dir,
            capture_output=True, text=True, timeout=5, check=True,
        ).stdout.strip()
        when = datetime.fromisoformat(commit_time).strftime("%Y-%m-%d %H:%M")
        return f"{sha} · {when}"
    except Exception:                                             # noqa: BLE001
        return "unknown build"


BUILD_STAMP = _read_build_stamp()


# ---------------------------------------------------------------------------
# Local test mode
#
# Skips the password gate and preselects a user so the real UI can be driven
# without anyone sharing a password. Deliberately awkward to switch on, and
# impossible to switch on in production:
#
#   * read ONLY from the OS environment -- never st.secrets, which is what a
#     deployed instance actually has, and which someone could set by mistake
#     while editing rates;
#   * refused outright when the process looks like Streamlit Cloud, so even
#     an env var set in the dashboard can't open the gate;
#   * loud on every run it's active, because a bypass nobody notices is one
#     that eventually ships.
# ---------------------------------------------------------------------------
TEST_MODE_ENV = "PROPOSAL_BUILDER_TEST_MODE"


TEST_MODE_USER = "Test Mode"


# Set by Streamlit Cloud's runtime; their presence means this is not a laptop.
_CLOUD_MARKERS = ("STREAMLIT_SHARING_MODE", "STREAMLIT_CLOUD",
                  "STREAMLIT_RUNTIME_ENV", "HOSTNAME_OVERRIDE")


def running_on_streamlit_cloud():
    if any(os.environ.get(marker) for marker in _CLOUD_MARKERS):
        return True
    # The deployed instance runs from /home/adminuser/... on Linux; a laptop
    # checkout never does.
    return sys.platform.startswith("linux") and Path.home().name == "adminuser"


def test_mode_active():
    """True only for a local process that asked for it via the environment."""
    if os.environ.get(TEST_MODE_ENV) != "1":
        return False
    if running_on_streamlit_cloud():
        # Refused rather than honoured: nothing legitimate sets this there.
        return False
    return True


# ---------------------------------------------------------------------------
# Dev deployment (dev branch, isolated from the public app -- see CLAUDE.md's
# git workflow section). Unlike TEST_MODE_ENV above, this one is MEANT to run
# on a deployed Streamlit Cloud instance -- the whole point is a second,
# separately-deployed app for daytime work -- so it reads st.secrets too, not
# just the OS environment.
# ---------------------------------------------------------------------------
DEV_MODE_ENV = "PROPOSAL_BUILDER_DEV_MODE"


def _secret_truthy(value):
    """A TOML secrets editor is free to store `DEV_MODE = "1"` as a string,
    `= 1` as an int, or `= true` as a real bool depending on how it's typed
    in -- a bare `== "1"` string check silently misses the second and third
    (found live: the first real dev deployment set it and the banner never
    showed). Accepts any of them, plus common truthy spellings, so how it
    got entered doesn't matter."""
    return str(value).strip().lower() in ("1", "true", "yes", "on")


def dev_mode_active():
    """True when this deployment is explicitly flagged as the dev/staging
    app. Checked, never assumed, from either source: a local process (the
    OS environment) or the deployed dev instance (st.secrets, since that's
    where a real Streamlit Cloud deployment's config actually lives)."""
    if _secret_truthy(os.environ.get(DEV_MODE_ENV, "")):
        return True
    try:
        return _secret_truthy(st.secrets.get(DEV_MODE_ENV, ""))
    except Exception:
        return False


def render_dev_banner():
    """A persistent, impossible-to-miss banner at the top of every page --
    the whole reason this exists is so nobody mistakes the dev deployment
    (newest UPLOADED deck/template content, per db.deck_channel) for the
    public app. Rendered unconditionally at the very top of main(), before
    even the password gate, so it shows on the login screen too."""
    if dev_mode_active():
        st.warning("🚧 **DEV** — this is the development deployment (previewing the newest "
                   "uploaded deck/template content, not what's live for the public app). "
                   f"Channel: `{db.deck_channel()}`.")


ADD_USER_OPTION = "➕ Add a name..."


def current_user():
    return st.session_state.get("current_user")


def _git_branch():
    try:
        return subprocess.run(["git", "rev-parse", "--abbrev-ref", "HEAD"],
                              cwd=Path(__file__).parent, capture_output=True, text=True,
                              timeout=5, check=True).stdout.strip()
    except Exception:                                             # noqa: BLE001
        return "unknown"


def password_gate(title, secret_name, caption=None):
    """The shared-password login for one entry point. The proposal builder
    uses APP_PASSWORD; the finders-only app uses its own FINDERS_PASSWORD, so
    a finders seller never holds the builder's password. Local test mode
    bypasses it (and says so loudly), same as before."""
    if test_mode_active():
        st.session_state["authed"] = True
        st.session_state.setdefault("current_user", TEST_MODE_USER)
        st.warning(f"⚠️ **{TEST_MODE_ENV} is on.** The password gate is bypassed and you're "
                   f"signed in as \"{TEST_MODE_USER}\". This is for local testing only — "
                   f"unset the environment variable to restore the login.")
        return True
    if st.session_state.get("authed"):
        return True

    st.title(title)
    if caption:
        st.caption(caption)
    pwd = st.text_input("Password", type="password")
    if st.button("Log in"):
        expected = st.secrets.get(secret_name)
        if expected is None:
            st.error(f"{secret_name} is not set in .streamlit/secrets.toml.")
        elif pwd == expected:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.caption(f"Build {BUILD_STAMP}")
    return False


def identity_step(title, caption):
    """Ask who's using the app, after the password -- attribution, not
    authentication. The list builds itself (anyone can add a name, deduped
    case-insensitively). If Supabase is unreachable the step is skipped
    rather than blocking."""
    # Comments, not a second docstring paragraph: a bare string below the
    # first statement is copy Streamlit's magic would render on the page.
    if test_mode_active():
        st.session_state.setdefault("current_user", TEST_MODE_USER)
        return True

    if current_user():
        return True

    names, warning = db.fetch_team_members()
    if names is None:
        st.session_state["current_user"] = None
        st.session_state["identity_skipped"] = True
        return True

    st.title(title)
    st.caption(caption)

    options = names + [ADD_USER_OPTION]
    picked = st.selectbox("Your name", options, index=None, placeholder="Choose your name",
                          key="identity_pick")

    if picked == ADD_USER_OPTION:
        new_name = st.text_input("Your name", key="identity_new_name",
                                 placeholder="First name is fine")
        if st.button("Add and continue", disabled=not new_name.strip()):
            stored, error = db.add_team_member(new_name)
            if error:
                st.error(error)
            else:
                st.session_state["current_user"] = stored
                st.rerun()
    elif picked:
        if st.button("Continue"):
            st.session_state["current_user"] = picked
            st.rerun()

    if warning:
        st.caption(warning)
    return False


def render_identity_sidebar():
    """Who's signed in, and a way to change it."""
    user = current_user()
    if user:
        st.sidebar.caption(f"Signed in as **{user}**")
    elif st.session_state.get("identity_skipped"):
        st.sidebar.caption("Not signed in — proposals won't be attributed")
    if st.sidebar.button("Switch user", use_container_width=True):
        for key in ("current_user", "identity_skipped", "identity_pick", "identity_new_name"):
            st.session_state.pop(key, None)
        st.rerun()


def feedback_popover(capture_state, categories, page, caption=None):
    """The persistent "Report an issue" control, in the sidebar.

    `capture_state(page)` returns the best-effort reproduction context the
    caller can offer (the builder's form state, or the finders app's own
    page and searches, tagged as coming from that app). `feedback_form_gen`
    moves the form's widgets to fresh keys after a successful submit, rather
    than writing into an already-instantiated widget's key, which Streamlit
    refuses outright."""
    gen = st.session_state.get("feedback_form_gen", 0)
    with st.sidebar.popover("🚩 Report an issue", use_container_width=True):
        st.caption(caption or "Something wrong, confusing, or missing? This attaches your current "
                              "page and campaign state (never a screenshot) so it's reproducible.")
        category = st.selectbox("Category", categories, key=f"feedback_category_{gen}")
        notes = st.text_area("What happened?", key=f"feedback_notes_{gen}", height=100)
        if st.button("Submit report", key=f"feedback_submit_{gen}", disabled=not notes.strip()):
            row_id, error = db.submit_feedback(
                category=category, notes=notes.strip(), page=page, state=capture_state(page),
                created_by=current_user())
            if error:
                st.error(f"Couldn't submit ({error}) — try again, or flag it directly to Matt.")
            else:
                st.session_state["feedback_form_gen"] = gen + 1
                st.session_state["feedback_just_submitted"] = True
                st.rerun()
    if st.session_state.pop("feedback_just_submitted", False):
        st.sidebar.success("Report submitted — thanks.")
