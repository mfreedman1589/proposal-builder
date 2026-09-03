"""Client name: a real placeholder now, not a committed default -- and
Generate is blocked while it's empty. Interface audit finding, fixed
2026-09-04: "Acme Test Co" used to be `value=` on the widget, a real string
sitting in the field from the very first render, indistinguishable at a
glance from a client name someone actually typed. Nothing at Generate
checked whether it was still there -- a wrong or missing name on a
client-facing deck is the one mistake here that can't be walked back once
sent.

    python tests/test_client_name_required.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)

failures = []


def check(label, ok, detail=""):
    print(f"  {'PASS' if ok else 'FAIL'}  {label}{('  -- ' + str(detail)) if detail else ''}")
    if not ok:
        failures.append(label)


def new_app(client_name=None):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.session_state["avails_mode"] = False
    if client_name is not None:
        at.session_state["client_name"] = client_name
    return at


def generate_button(at):
    return [b for b in at.button if b.label == "Generate proposal"][0]


def main():
    print("=" * 78)
    print("SCENARIO  the field itself starts genuinely empty, not pre-filled")
    print("=" * 78)
    at = new_app()
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    name_input = [t for t in at.text_input if t.key == "client_name"][0]
    check('client_name starts as "" -- a real placeholder, not a committed default',
          name_input.value == "", repr(name_input.value))

    print("\n" + "=" * 78)
    print("SCENARIO  Generate is disabled and warned about while client name is empty")
    print("=" * 78)
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    check("Generate proposal button is disabled",
          generate_button(at).disabled, generate_button(at).disabled)
    warnings = [str(w.value) for w in at.warning]
    check("a warning names the client-name requirement",
          any("client name is required" in w.lower() for w in warnings), warnings)

    print("\n" + "=" * 78)
    print("SCENARIO  typing a real name re-enables Generate")
    print("=" * 78)
    at2 = new_app(client_name="Acme Test Co")
    at2.run()
    check("no exception", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    check("Generate proposal is enabled once a real name is set",
          not generate_button(at2).disabled, generate_button(at2).disabled)
    warnings2 = [str(w.value) for w in at2.warning]
    check("no client-name warning once it's set",
          not any("client name is required" in w.lower() for w in warnings2), warnings2)

    print("\n" + "=" * 78)
    print("SCENARIO  whitespace-only doesn't count as a real name either")
    print("=" * 78)
    at3 = new_app(client_name="   ")
    at3.run()
    check("no exception", not at3.exception,
          at3.exception[0].message[:400] if at3.exception else "")
    check("Generate stays disabled for a whitespace-only name",
          generate_button(at3).disabled, generate_button(at3).disabled)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Client name starts genuinely empty (a real placeholder, not a committed default), "
          "and Generate is blocked with a visible warning until a real name is typed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
