"""The setup band's flight-change guard: shortening or lengthening the
flight after real money is on a Monthly-breakout plan option must flag,
never silently rewrite -- see DECISIONS.md, "A Monthly-breakout option's
full-flight total is live, not stored". Warns (not blocks) the moment the
flight changes, names both dollar figures, and genuinely blocks Generate
until the rep dismisses it having seen them.

A Full-Flight-breakout option is the one that's SAFE from this -- its
stored Cost IS the full-flight figure directly, immune to a flight-length
change by construction -- and is asserted exempt here for exactly that
reason (an earlier draft of this guard had the two backwards).

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


def warning_texts(at):
    return [str(w.value) if hasattr(w, "value") else str(w) for w in at.warning]


def flight_warnings(at):
    return [w for w in warning_texts(at) if "flight changed" in w.lower()]


def generate_button(at):
    matches = [b for b in at.button if b.label == "Generate proposal"]
    return matches[0] if matches else None


def dismiss_button(at):
    matches = [b for b in at.button if "dismiss this warning" in b.label]
    return matches[0] if matches else None


def main():
    print("first render establishes the baseline silently, Generate is enabled")
    at = new_app()
    at.run()
    check("no exception on first render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("no flight-change warning on a fresh baseline", not flight_warnings(at), warning_texts(at))
    gen = generate_button(at)
    check("Generate button is present", gen is not None)
    check("Generate is enabled", gen is not None and not gen.disabled)

    print("\nshortening the flight after real money is on a Monthly-breakout option "
          "warns (naming both totals) and blocks Generate")
    at.session_state["flight_end"] = date(2026, 9, 30)   # 3 months -> 1 month
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    fw = flight_warnings(at)
    check("exactly one flight-change warning appeared", len(fw) == 1, warning_texts(at))
    if fw:
        w = fw[0]
        # 3750/mo x 3 months = 11,250 old full-flight total; x1 month = 3,750 new.
        check("names the old full-flight total ($11,250, 3 months)", "11,250" in w, w)
        check("names the new full-flight total ($3,750, 1 month)", "3,750" in w, w)
        check("names the option", "Option A" in w, w)
    gen2 = generate_button(at)
    check("Generate is now disabled", gen2 is not None and gen2.disabled)
    dismiss = dismiss_button(at)
    check("a dismiss button is present", dismiss is not None)

    print("\ndismissing clears the warning and re-enables Generate")
    if dismiss:
        dismiss.click().run()
    check("no exception after dismiss", not at.exception, at.exception[0].message[:400] if at.exception else "")
    check("no flight-change warning after dismiss", not flight_warnings(at), warning_texts(at))
    gen3 = generate_button(at)
    check("Generate is enabled again", gen3 is not None and not gen3.disabled)

    print("\nlengthening the flight also warns (expansion, not just shrink) -- widening "
          "flight_end alone doesn't widen n_months (a partial-overlap active_months "
          "selection is kept, not auto-grown), so this also exercises actually adding "
          "the new months as active, the same as a rep would")
    at2 = new_app()
    at2.run()
    at2.session_state["flight_end"] = date(2027, 2, 28)   # 3 months -> 6 months
    at2.session_state["active_months"] = [
        "Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026", "Jan 2027", "Feb 2027"]
    at2.run()
    check("a flight-change warning appeared on expansion too", flight_warnings(at2), warning_texts(at2))

    print("\na Full-Flight-breakout option is exempt -- its stored Cost IS the "
          "full-flight figure, immune to a flight-length change")
    at3 = new_app(breakout=app.BREAKOUT_FULL_FLIGHT, cost=11250.0)
    at3.run()
    at3.session_state["flight_end"] = date(2026, 9, 30)
    at3.run()
    check("no exception", not at3.exception, at3.exception[0].message[:400] if at3.exception else "")
    check("no flight-change warning for a Full-Flight-breakout option",
          not flight_warnings(at3), warning_texts(at3))
    gen4 = generate_button(at3)
    check("Generate stays enabled for a Full-Flight-breakout option",
          gen4 is not None and not gen4.disabled)

    print("\na clean (never-edited) row is exempt -- nothing priced to review")
    at4 = new_app(dirty=False)
    at4.run()
    at4.session_state["flight_end"] = date(2026, 9, 30)
    at4.run()
    check("no flight-change warning for an option with no priced content",
          not flight_warnings(at4), warning_texts(at4))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Shortening or lengthening the flight after real money is on a Monthly-breakout "
          "option warns (naming both dollar totals) and blocks Generate until dismissed; a "
          "Full-Flight-breakout option or an unpriced row is correctly exempt.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
