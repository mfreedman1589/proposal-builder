"""QA-018: choosing a main page from an Admin page always switches pages.

Muse's repro (2026-10-02): right after login, on Admin -> Feedback reports,
clicking "Build a proposal" in the sidebar left Feedback reports on screen.
The sidebar radio still showed the last main page as selected while an Admin
page was open, so clicking that same option sent no change. An Admin page
now clears the radio's selection.

Checks every Admin page against two main pages -- the one the radio used to
show by default, and one it didn't.

    python tests/test_admin_nav.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db                                       # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                      # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    from streamlit.testing.v1 import AppTest
    for admin in app.ADMIN_PAGES:
        for target in ("Build a proposal", "Proposal history"):
            at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
            at.session_state["authed"] = True
            at.session_state["current_user"] = "Nav Suite"
            at.run()
            at.session_state["goto_admin_tool"] = admin
            at.run()
            nav = next(r for r in at.radio if r.key == "nav_section")
            check(f"on {admin}: no main page shows as selected", nav.value is None, nav.value)
            nav.set_value(target).run()
            check(f"{admin} -> {target}: the page switches",
                  at.session_state["page_choice"] == target and not at.exception,
                  (at.session_state["page_choice"], at.exception))
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
