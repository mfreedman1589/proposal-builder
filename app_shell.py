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
