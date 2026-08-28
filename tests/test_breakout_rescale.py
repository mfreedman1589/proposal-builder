"""Flipping an option's own Monthly/Full Flight radio must rescale its rows,
not just relabel them -- found live, testing a $15K/4-month Capital Media
draft: "Full flight and monthly plan toggle came out to the same numbers."

Root cause: the radio widget only ever set `option["breakout"]`, which
`compute_plan_totals` reads to decide whether a row's stored Cost/
Impressions ALREADY represent the monthly figure or the full-flight one --
but nothing ever rescaled the row's own numbers when the toggle changed, so
the exact same stored number got reinterpreted as meaning something totally
different depending on which basis happened to be selected. For a
single-month flight (n_months=1) both interpretations coincide, which is
exactly the reported symptom; for any other flight length they'd disagree
in a genuinely wrong way, not merely a coincidentally-matching one.

Fix: `rescale_rows_for_breakout_change` (app.py), called right after the
option's own `st.radio(...)` sets `option["breakout"]` for this run,
detects an actual flip (against a stored `_breakout_basis`, defaulting to
the option's own current value so a fresh option's first render is never
treated as a change) and rescales every RATE row's Cost/Impressions by
that row's own month count -- flat fees and broadcast rows excluded
(each already has its own correct rule).

    python tests/test_breakout_rescale.py
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


# --------------------------------------------------------------------------
# Part 1 -- the pure function
# --------------------------------------------------------------------------

def rate_row(cost, impressions, cpm=30.0):
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep 2026 - Nov 2026",
           "Geo": "Washington, DC DMA", "Targeting": "Segment A",
           "Impressions": impressions, "CPM": cpm, "Type": app.ROW_TYPE_RATE, "Cost": cost}


def flat_fee_row(cost):
    return {"Tactic": "Setup Fee", "Flight": "Sep 2026 - Nov 2026",
           "Geo": "Washington, DC DMA", "Targeting": "Segment A",
           "Impressions": 0.0, "CPM": 0.0, "Type": app.ROW_TYPE_FLAT_FEE, "Cost": cost}


def test_pure_function():
    print("rescale_rows_for_breakout_change -- pure function")

    print("\nno actual flip (radio didn't change) -- rows untouched, no version bump")
    opt = app.new_plan_option("A", [rate_row(1000.0, 30000.0)])
    opt["breakout"] = app.BREAKOUT_MONTHLY
    changed = app.rescale_rows_for_breakout_change(opt, 3)
    check("no change reported on a fresh option's first render",
          changed is False, changed)
    check("Cost untouched", opt["rows"][0]["Cost"] == 1000.0, opt["rows"][0]["Cost"])

    print("\nMonthly -> Full Flight: Cost and Impressions scale UP by n_months")
    opt2 = app.new_plan_option("A", [rate_row(1000.0, 30000.0)])
    opt2["breakout"] = app.BREAKOUT_MONTHLY
    app.rescale_rows_for_breakout_change(opt2, 3)  # settle _breakout_basis
    opt2["breakout"] = app.BREAKOUT_FULL_FLIGHT  # the rep just flipped the radio
    changed2 = app.rescale_rows_for_breakout_change(opt2, 3)
    check("change reported", changed2 is True, changed2)
    check("Cost scaled to the full flight (1000 x 3)", opt2["rows"][0]["Cost"] == 3000.0,
          opt2["rows"][0]["Cost"])
    check("Impressions scaled the same way (30000 x 3)",
          opt2["rows"][0]["Impressions"] == 90000.0, opt2["rows"][0]["Impressions"])

    print("\nFull Flight -> Monthly: Cost and Impressions scale back DOWN")
    opt3 = app.new_plan_option("A", [rate_row(3000.0, 90000.0)])
    opt3["breakout"] = app.BREAKOUT_FULL_FLIGHT
    app.rescale_rows_for_breakout_change(opt3, 3)
    opt3["breakout"] = app.BREAKOUT_MONTHLY
    app.rescale_rows_for_breakout_change(opt3, 3)
    check("Cost scaled back down (3000 / 3)", opt3["rows"][0]["Cost"] == 1000.0,
          opt3["rows"][0]["Cost"])
    check("Impressions scaled back down (90000 / 3)",
          opt3["rows"][0]["Impressions"] == 30000.0, opt3["rows"][0]["Impressions"])

    print("\na single-month flight (n_months=1) -- the reported symptom's own root cause: "
          "monthly and full-flight coincide for THIS row, but only because 1x is a no-op, "
          "not because the mechanism is skipped")
    opt4 = app.new_plan_option("A", [rate_row(1000.0, 30000.0)])
    opt4["breakout"] = app.BREAKOUT_MONTHLY
    app.rescale_rows_for_breakout_change(opt4, 1)
    opt4["breakout"] = app.BREAKOUT_FULL_FLIGHT
    app.rescale_rows_for_breakout_change(opt4, 1)
    check("Cost unchanged for a 1-month flight (1000 x 1)", opt4["rows"][0]["Cost"] == 1000.0,
          opt4["rows"][0]["Cost"])

    print("\na flat fee is NEVER rescaled -- its Cost is the full-flight amount regardless of basis")
    opt5 = app.new_plan_option("A", [rate_row(1000.0, 30000.0), flat_fee_row(500.0)])
    opt5["breakout"] = app.BREAKOUT_MONTHLY
    app.rescale_rows_for_breakout_change(opt5, 3)
    opt5["breakout"] = app.BREAKOUT_FULL_FLIGHT
    app.rescale_rows_for_breakout_change(opt5, 3)
    check("the rate row scaled", opt5["rows"][0]["Cost"] == 3000.0, opt5["rows"][0]["Cost"])
    check("the flat fee did NOT scale", opt5["rows"][1]["Cost"] == 500.0, opt5["rows"][1]["Cost"])

    print("\na DIRTY (hand-edited) row still rescales -- dirty means \"don't reseed from "
          "shared defaults,\" never \"exempt from the basis this option says it's in\"")
    opt6 = app.new_plan_option("A", [rate_row(1000.0, 30000.0)])
    opt6["dirty"] = [True]
    opt6["breakout"] = app.BREAKOUT_MONTHLY
    app.rescale_rows_for_breakout_change(opt6, 3)
    opt6["breakout"] = app.BREAKOUT_FULL_FLIGHT
    app.rescale_rows_for_breakout_change(opt6, 3)
    check("the dirty row still scaled", opt6["rows"][0]["Cost"] == 3000.0, opt6["rows"][0]["Cost"])


# --------------------------------------------------------------------------
# Part 2 -- through the real render
# --------------------------------------------------------------------------

def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)   # 3 calendar months
    at.session_state["plan_options"] = [{
        "name": "Option A", "breakout": app.BREAKOUT_MONTHLY,
        "rows": [rate_row(3750.0, 101902.0)],
        "dirty": [True], "driver": [app.DRIVER_COST], "version": 0,
    }]
    return at


def ss(at, key, default=None):
    return at.session_state[key] if key in at.session_state else default


def test_full_cycle():
    print("\nfull render cycle: flipping the real radio widget rescales the real row")
    at = new_app()
    at.run()
    check("no exception on initial render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    row = ss(at, "plan_options")[0]["rows"][0]
    check("Cost unchanged on a render with no interaction", row["Cost"] == 3750.0, row["Cost"])

    print("...flip to Full Flight")
    # FLOW_REWORK_PLAN.md Phase 2 defect fix (2026-08-28): the widget's real
    # key now carries a third, per-option `_breakout_gen` segment (see
    # test_breakout_band_sync.py) so an unlocked option can be re-keyed when
    # it re-syncs to the band -- this fixture's option is unlocked and never
    # diverges from the band's own default (avails_basis unset here), so
    # `_breakout_gen` stays 0 and the key is `..._0_0_0`, not `..._0_0`.
    at.session_state["option_breakout_0_0_0"] = app.BREAKOUT_FULL_FLIGHT
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    row2 = ss(at, "plan_options")[0]["rows"][0]
    check("Cost scaled to the full flight (3750 x 3 = 11250), not left at 3750",
          row2["Cost"] == 11250.0, row2["Cost"])
    check("breakout really is Full Flight now",
          ss(at, "plan_options")[0]["breakout"] == app.BREAKOUT_FULL_FLIGHT,
          ss(at, "plan_options")[0]["breakout"])

    print("...an untouched further rerun does not keep rescaling")
    at.run()
    row3 = ss(at, "plan_options")[0]["rows"][0]
    check("Cost stays at 11250 -- not rescaled again on an untouched rerun",
          row3["Cost"] == 11250.0, row3["Cost"])

    print("...flip back to Monthly")
    at.session_state["option_breakout_0_0_0"] = app.BREAKOUT_MONTHLY
    at.run()
    row4 = ss(at, "plan_options")[0]["rows"][0]
    check("Cost scaled back down to 3750", row4["Cost"] == 3750.0, row4["Cost"])


def main():
    test_pure_function()
    test_full_cycle()
    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("An option's Monthly/Full Flight toggle now rescales its own rows when flipped, "
          "so the two bases genuinely disagree by the flight's real month count instead of "
          "silently reinterpreting the same stored number.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
