"""FLOW_REWORK_PLAN.md Phase 2 defect, found live testing the Capital Media
walkthrough with the setup band's Plan basis set to Full flight (2026-08-28):
the generated deck still showed Monthly Impressions/CPM/Monthly Cost columns
plus a "Monthly Totals" row, with a separate "Full Flight Total (4 months)"
row underneath -- as if the band's choice had been ignored outright.

Two independent, compounding bugs, both fixed here:

(a) The band's Plan basis only ever set a BRAND-NEW option's initial
    breakout. `app.py`'s per-option Breakout radio (`key=f"option_breakout_
    {gen}_{idx}"`) rendered once on the very first run -- before a rep could
    ever touch the band -- and from then on its own widget key permanently
    won over `index=`, the same "a keyed widget's session_state beats its
    value=/index= argument" trap this file documents elsewhere for the
    option-name bug. So flipping the band afterward, the ordinary way a rep
    would ever act on it, silently did nothing. Fixed: an option that hasn't
    been locked (`option_breakout_locked`) now re-syncs to the band's current
    default every run it's found diverged from it, bumping a small
    per-option `_breakout_gen` counter so the radio gets a fresh key and
    `index=` is honored again. An option whose Breakout the rep (or a draft,
    or a restored proposal) set directly locks, so the band can never
    silently override a deliberate choice.

(b) Even when the option's own breakout genuinely says Full Flight,
    `_option_payload` (the deck-payload builder) never read it: `totals_
    label` was hardcoded to "Monthly Totals" and the separate "Full Flight
    Total (N months)" footer was added whenever n_months > 1, regardless of
    breakout. Fixed: a Full-Flight-breakout option's deck payload now reads
    "Full Flight Totals" over the real full-flight figure, with no separate
    footer -- it would be the same number twice. A Monthly-breakout option's
    payload is byte-identical to before this fix.

    python tests/test_breakout_band_sync.py
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

import app  # noqa: E402
import assembly  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def g(at, key, default=None):
    try:
        return at.session_state[key]
    except (KeyError, AttributeError):
        return default


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 12, 20)  # 4 months, like Capital Media
    at.session_state["premion_streaming_tv"] = True
    return at


def flip_band_basis(at, value):
    basis_radio = [r for r in at.radio if r.key == "avails_basis"][0]
    basis_radio.set_value(value).run()


def test_unlocked_option_follows_a_later_band_flip():
    print("\nan option that has never been touched follows the band, "
          "even after it already exists")
    at = new_app()
    at.run()
    check("no exception on initial render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    opt = g(at, "plan_options")[0]
    check("Option A starts Monthly, matching the band's own default",
          opt["breakout"] == app.BREAKOUT_MONTHLY, opt["breakout"])
    check("starts unlocked", not opt.get("option_breakout_locked"), opt)

    flip_band_basis(at, app.AVAILS_BASIS_FLIGHT)
    check("no exception after flipping the band", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    opt2 = g(at, "plan_options")[0]
    check("the ALREADY-EXISTING option's breakout followed the band flip "
          "to Full Flight -- this is the real defect: it used to stay "
          "Monthly (default) here", opt2["breakout"] == app.BREAKOUT_FULL_FLIGHT,
          opt2["breakout"])
    check("still unlocked -- following the band is not the same as a rep's "
          "own choice", not opt2.get("option_breakout_locked"), opt2)

    print("...and flips back if the band flips back")
    flip_band_basis(at, app.AVAILS_BASIS_MONTHLY)
    opt3 = g(at, "plan_options")[0]
    check("follows the band back to Monthly too",
          opt3["breakout"] == app.BREAKOUT_MONTHLY, opt3["breakout"])


def test_rep_set_breakout_survives_a_later_band_flip():
    print("\na rep's own Breakout choice locks and survives a later band flip")
    at = new_app()
    at.run()
    opt = g(at, "plan_options")[0]
    check("starts Monthly", opt["breakout"] == app.BREAKOUT_MONTHLY, opt["breakout"])

    print("...rep picks Full Flight on the option's own radio directly "
          "(band is still Monthly)")
    breakout_radio = [r for r in at.radio if r.key.startswith("option_breakout_0_0")][0]
    breakout_radio.set_value(app.BREAKOUT_FULL_FLIGHT).run()
    opt2 = g(at, "plan_options")[0]
    check("the option is now Full Flight", opt2["breakout"] == app.BREAKOUT_FULL_FLIGHT,
          opt2["breakout"])
    check("and is now locked -- a direct widget change is a deliberate choice",
          bool(opt2.get("option_breakout_locked")), opt2)

    print("...band flips to Full flight and back to Monthly -- must not move the option")
    flip_band_basis(at, app.AVAILS_BASIS_FLIGHT)
    opt3 = g(at, "plan_options")[0]
    check("still Full Flight (band agrees here, so this alone doesn't prove much)",
          opt3["breakout"] == app.BREAKOUT_FULL_FLIGHT, opt3["breakout"])
    flip_band_basis(at, app.AVAILS_BASIS_MONTHLY)
    opt4 = g(at, "plan_options")[0]
    check("still Full Flight after the band moves AWAY -- this is the real proof: "
          "an unlocked option would have followed the band back to Monthly",
          opt4["breakout"] == app.BREAKOUT_FULL_FLIGHT, opt4["breakout"])


def _rate_row(cost):
    # Impressions is deliberately omitted/inconsistent here -- the real D2
    # grid's own reconciliation (driver=DRIVER_COST) recomputes it fresh from
    # Cost/CPM every render regardless of breakout (that basis interpretation
    # happens later, only in compute_plan_totals), so the only self-consistent
    # fixture is one built from Cost and CPM and left to derive its own
    # Impressions -- exactly what a rep typing a dollar figure into the grid
    # would get.
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep 2026 - Dec 2026",
            "Geo": "Washington, DC DMA", "Targeting": "Segment A",
            "Impressions": 0.0, "CPM": 30.0, "Type": app.ROW_TYPE_RATE,
            "Cost": cost}


def _capture_payload(breakout, cost):
    """Drive a real Generate with one priced option at the given breakout,
    return that option's own deck payload (`media_plan_options[0]`).

    `cost` is what a rep would type into the row's own Cost cell -- a
    MONTHLY rate under Monthly breakout, the FULL-FLIGHT figure directly
    under Full Flight breakout (CLAUDE.md: "A Full-Flight-breakout row's
    stored Cost IS the full-flight figure directly").
    """
    captured = {}
    real_personalize = assembly.personalize

    def spy(prs, fill_data):
        w = real_personalize(prs, fill_data)
        captured["fill_data"] = fill_data
        return w

    assembly.personalize = spy
    try:
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        at.session_state["market_choice"] = "DC"
        at.session_state["flight_start"] = date(2026, 9, 1)
        at.session_state["flight_end"] = date(2026, 12, 20)  # 4 months
        at.session_state["premion_streaming_tv"] = True
        at.session_state["plan_options"] = [app.new_plan_option(
            "Option A", [_rate_row(cost)],
            driver=[app.DRIVER_COST], breakout=breakout, breakout_locked=True)]
        at.run()
        gen = [b for b in at.button if b.label == "Generate proposal"]
        gen[0].click().run()
        assert not at.exception, at.exception[0].message
    finally:
        assembly.personalize = real_personalize
    return captured["fill_data"]["media_plan_options"][0]


def test_deck_payload_full_flight_totals():
    print("\ndeck payload: a Full-Flight-breakout option says \"Full Flight "
          "Totals\" over the real full-flight figure, no separate footer")
    # $15,000 IS the full-flight cost here (Full Flight breakout); at $30 CPM
    # that's 500,000 full-flight impressions, 125,000/month across 4 months.
    payload = _capture_payload(app.BREAKOUT_FULL_FLIGHT, cost=15000.0)
    check('totals_label is "Full Flight Totals", not "Monthly Totals"',
          payload["totals_label"] == "Full Flight Totals", payload["totals_label"])
    check("total_cost is the real full-flight figure the rep typed (15000), "
          "not a derived monthly one", payload["total_cost"] == "$15,000",
          payload["total_cost"])
    check("total_impressions is the full-flight figure (500,000)",
          payload["total_impressions"] == "500,000", payload["total_impressions"])
    check("no separate Full Flight Total footer -- it would be the same "
          "number as the totals row above", payload["full_flight_total"] is None,
          payload["full_flight_total"])


def test_deck_payload_monthly_byte_identical():
    print("\ndeck payload: a Monthly-breakout option is byte-identical to "
          "before this fix")
    # $3,750/month (Monthly breakout) x 4 months = $15,000 full flight, so
    # this exercises the exact same total-dollar scenario as the Full-Flight
    # case above while proving Monthly's OWN presentation is untouched.
    payload = _capture_payload(app.BREAKOUT_MONTHLY, cost=3750.0)
    check('totals_label is "Monthly Totals"',
          payload["totals_label"] == "Monthly Totals", payload["totals_label"])
    check("total_cost is the monthly figure the rep typed (3750)",
          payload["total_cost"] == "$3,750", payload["total_cost"])
    check("total_impressions is the monthly figure (125,000)",
          payload["total_impressions"] == "125,000", payload["total_impressions"])
    check("the separate Full Flight Total footer is still present (4 months, "
          "3750 x 4 = 15000, 125000 x 4 = 500000)",
          payload["full_flight_total"] == {
              "label": "Full Flight Total (4 months)",
              "impressions": "500,000", "cost": "$15,000"},
          payload["full_flight_total"])


def main():
    test_unlocked_option_follows_a_later_band_flip()
    test_rep_set_breakout_survives_a_later_band_flip()
    test_deck_payload_full_flight_totals()
    test_deck_payload_monthly_byte_identical()
    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The setup band's Plan basis now actually drives an existing option's "
          "Breakout (until a rep deliberately overrides it), and the generated "
          "deck's totals row -- not just the app's own D2 grid column headers -- "
          "reads that same breakout instead of always presenting as Monthly.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
