"""Render the two screens every user meets, and inventory the copy on them.

    python tests/test_entry_screens.py          # check against the inventory
    python tests/test_entry_screens.py --show   # print what's on screen

The login screen and the identity step are the first things anyone sees, and
they're the two screens no other test renders -- every suite sets `authed`
and `current_user` to skip straight past them. That gap is how several
paragraphs of internal design rationale ("Not authentication -- the password
is the gate...") came to be displayed at the top of the login screen, where
it stayed until somebody logged in and read it.

The mechanism is worth knowing: Streamlit rewrites any bare expression into
`st.write()`, and a string is exempt only as the FIRST statement of its
function. Inserting a test-mode early return above `check_identity`'s
docstring silently turned it into rendered copy. `test_no_stray_magic.py`
catches that specific shape by AST; this catches it by rendering, along with
anything else that lands on these screens -- a stray st.write, a debug
caption, a leaked warning -- because it asserts the full inventory of
visible text rather than the absence of one known string.

These screens can't be driven through a browser: the password lives in
`.streamlit/secrets.toml`, and test mode bypasses both screens rather than
showing them. AppTest runs the real script through the real Streamlit
runtime -- magic transform included -- which is what makes it able to catch
this at all.

When the copy legitimately changes, run --show and paste the new inventory in.
"""

import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
# Must be off, or both screens are skipped and this test renders nothing.
os.environ.pop("PROPOSAL_BUILDER_TEST_MODE", None)

from streamlit.testing.v1 import AppTest              # noqa: E402
import db                                             # noqa: E402

# Every string a user should see on each screen, in no particular order.
# Anything else on screen is the failure -- that's the point.
EXPECTED = {
    "login": {
        "Premion Proposal Builder",
        "Password",
        "Log in",
    },
    "identity": {
        "Premion Proposal Builder",
        "Who's using the app? This just labels the proposals you generate so the team "
        "can tell whose is whose — pick your name, or add it if it's not there yet.",
        "Your name",
        "Choose your name",
    },
}

# Prose that has no business on a login screen. A generic guard for the
# class, not just the string that got out: design notes are long, and they
# explain rather than instruct.
RATIONALE_MARKERS = ("Not authentication", "rather than blocking", "deduplicated",
                     "the password is the gate", "-- ", "e.g.")


def visible_text(at):
    """Everything the user would read, in the order it appears."""
    def group(*names):
        # ElementList doesn't support +, and an older Streamlit may not have
        # every accessor; ask for each one defensively.
        for name in names:
            for element in list(getattr(at, name, []) or []):
                yield element

    seen = []
    for element in group("markdown", "caption", "code", "text", "title", "header"):
        value = getattr(element, "value", None)
        if isinstance(value, str) and value.strip():
            seen.append(value.strip())
    for element in group("button", "text_input", "selectbox"):
        label = getattr(element, "label", None)
        if isinstance(label, str) and label.strip():
            seen.append(label.strip())
        placeholder = getattr(element, "placeholder", None)
        if isinstance(placeholder, str) and placeholder.strip():
            seen.append(placeholder.strip())
    for element in group("warning", "error", "info", "success"):
        value = getattr(element, "value", None)
        if isinstance(value, str) and value.strip():
            seen.append(value.strip())
    return seen


def render(screen):
    at = AppTest.from_file("app.py", default_timeout=120)
    if screen == "identity":
        at.session_state["authed"] = True
    at.run()
    return at


def main():
    show = "--show" in sys.argv
    # The identity step only renders when the team list loads; with Supabase
    # unreachable it's skipped by design, which would make this vacuous.
    real_fetch = db.fetch_team_members
    db.fetch_team_members = lambda *a, **k: (["Matt", "Dana"], None)

    failures = []
    try:
        for screen in ("login", "identity"):
            at = render(screen)
            if at.exception:
                print(f"  FAIL  {screen}: {at.exception}")
                failures.append(screen)
                continue
            found = visible_text(at)
            if show:
                print(f"\n=== {screen} ===")
                for line in found:
                    print(f"    {line!r}")
                continue

            unexpected = [t for t in found if t not in EXPECTED[screen]]
            missing = sorted(EXPECTED[screen] - set(found))
            rationale = [t for t in found
                         if len(t) > 120 and any(m in t for m in RATIONALE_MARKERS)]

            for text in rationale:
                print(f"  FAIL  {screen}: design rationale is rendering to users")
                print(f"          {text[:150]}...")
            for text in unexpected:
                print(f"  FAIL  {screen}: unexpected copy on screen")
                print(f"          {text[:150]}")
            for text in missing:
                print(f"  FAIL  {screen}: expected copy is gone: {text[:80]!r}")
            if not (unexpected or missing or rationale):
                print(f"  PASS  {screen}: {len(found)} element(s), all accounted for")
            else:
                failures.append(screen)
    finally:
        db.fetch_team_members = real_fetch

    if show:
        return 0
    print()
    if failures:
        print(f"{len(failures)} screen(s) show copy they shouldn't: {failures}")
        print("If the change was deliberate, run with --show and update EXPECTED.")
        return 1
    print("Both entry screens show exactly the copy they're supposed to.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
