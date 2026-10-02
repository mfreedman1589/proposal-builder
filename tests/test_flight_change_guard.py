"""A flight change after a plan is priced (Matt's rule, 2026-10-02 --
resolves QA-005 and replaces the old block-and-dismiss guard):

- a MONTHLY option keeps its monthly budget and its total follows the
  months -- 3 months at $3,750/month is $11,250; shorten to 1 and it is
  $3,750;
- a FULL-FLIGHT option keeps its total, spread across the new dates;
- a short, non-blocking note says what happened ("Flight now 1 month --
  plan total $3,750 (was $11,250)" / "Total spend held at $11,250 across
  the new dates"). Generate is never blocked and nothing needs dismissing.

The dollar figures are worked from the fixture's own rate ($3,750/month x
months), never read back from the code under test.

    python tests/test_flight_change_guard.py
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

import app  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def rate_row(cost, impressions, cpm=30.0):
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep 2026 - Nov 2026",
           "Geo": "Washington, DC DMA", "Targeting": "Segment A",
           "Impressions": impressions, "CPM": cpm, "Type": app.ROW_TYPE_RATE, "Cost": cost}


def new_app(breakout=app.BREAKOUT_MONTHLY, dirty=True, cost=3750.0):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    # The Client name field lost its committed "Acme Test Co" default
    # (2026-09-04 -- it's a real placeholder now, and Generate is blocked
    # while it's empty), so this fixture has to supply one explicitly or
    # every "Generate is enabled" check below fails for a reason that has
    # nothing to do with what this file actually tests.
    at.session_state["client_name"] = "Acme Test Co"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)   # 3 calendar months
    at.session_state["plan_options"] = [{
        "name": "Option A", "breakout": breakout,
        # FLOW_REWORK_PLAN.md Phase 2 defect fix (2026-08-28): an option
        # whose breakout isn't locked now re-syncs to the setup band's own
        # Plan basis, which this fixture never sets (defaults to Monthly).
        # Locked here because this fixture's whole point is "a rep already
        # chose Full Flight for this option" -- exactly the deliberate case
        # the lock exists to protect from a silent band-driven override.
        "option_breakout_locked": breakout == app.BREAKOUT_FULL_FLIGHT,
        "rows": [rate_row(cost, 101902.0)],
        "dirty": [dirty], "driver": [app.DRIVER_COST], "version": 0,
    }]
    return at


def info_texts(at):
    return [str(i.value) if hasattr(i, "value") else str(i) for i in at.info]


def flight_notes(at):
    # Compared without the leading icon, which AppTest may or may not report.
    return [t.replace("🗓️", "").strip() for t in info_texts(at)
            if "Flight now" in t or "Total spend held" in t]


def generate_button(at):
    matches = [b for b in at.button if b.label == "Generate proposal"]
    return matches[0] if matches else None


def dismiss_buttons(at):
    return [b for b in at.button if "dismiss" in (b.label or "").lower()]


def err(at):
    return at.exception[0].message[:400] if at.exception else ""


def main():
    print("first render: no note, Generate enabled")
    at = new_app()
    at.run()
    check("no exception on first render", not at.exception, err(at))
    check("no flight note on a fresh plan", not flight_notes(at), info_texts(at))
    gen = generate_button(at)
    check("Generate is enabled", gen is not None and not gen.disabled)

    print("\nMonthly option, flight shortened 3 months -> 1: monthly budget held, total follows")
    at.session_state["flight_end"] = date(2026, 9, 30)
    at.run()
    check("no exception", not at.exception, err(at))
    notes = flight_notes(at)
    check("one note, Matt's wording, both totals ($3,750 now, was $11,250)",
          notes == ["Flight now 1 month — plan total $3,750 (was $11,250)."], notes)
    check("the monthly rate on the row is unchanged ($3,750/month)",
          at.session_state["plan_options"][0]["rows"][0]["Cost"] == 3750.0,
          at.session_state["plan_options"][0]["rows"][0]["Cost"])
    gen2 = generate_button(at)
    check("Generate is NOT blocked", gen2 is not None and not gen2.disabled)
    check("there is no dismiss step", not dismiss_buttons(at), [b.label for b in dismiss_buttons(at)])
    at.run()
    check("the note stays up on the next rerun (until the flight changes again)",
          flight_notes(at) == notes, flight_notes(at))

    print("\nMonthly option, flight lengthened 3 months -> 6: total follows up too")
    at2 = new_app()
    at2.run()
    at2.session_state["flight_end"] = date(2027, 2, 28)
    at2.session_state["active_months"] = [
        "Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026", "Jan 2027", "Feb 2027"]
    at2.run()
    check("note: 6 months, $22,500 (was $11,250)",
          flight_notes(at2) == ["Flight now 6 months — plan total $22,500 (was $11,250)."],
          flight_notes(at2))

    print("\nFull-Flight option: total spend held across the new dates")
    at3 = new_app(breakout=app.BREAKOUT_FULL_FLIGHT, cost=11250.0)
    at3.run()
    at3.session_state["flight_end"] = date(2026, 9, 30)
    at3.run()
    check("no exception", not at3.exception, err(at3))
    check("note: total spend held at $11,250",
          flight_notes(at3) == ["Total spend held at $11,250 across the new dates."],
          flight_notes(at3))
    check("the full-flight cost on the row is unchanged ($11,250)",
          at3.session_state["plan_options"][0]["rows"][0]["Cost"] == 11250.0)
    gen4 = generate_button(at3)
    check("Generate enabled", gen4 is not None and not gen4.disabled)

    print("\nA group-owned line starts unedited but carries money -- it still gets the note")
    at4 = new_app(dirty=False)
    at4.run()
    at4.session_state["flight_end"] = date(2026, 9, 30)
    at4.run()
    check("note for an unedited, priced Monthly line",
          flight_notes(at4) == ["Flight now 1 month — plan total $3,750 (was $11,250)."],
          flight_notes(at4))

    print("\nNothing priced: nothing to say")
    at5 = new_app(cost=0.0)
    at5.run()
    at5.session_state["flight_end"] = date(2026, 9, 30)
    at5.run()
    check("no note for a $0 plan", not flight_notes(at5), flight_notes(at5))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("All checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
