"""Avails ingestion, reach-based allocation, and the monthly/flight basis.

    python tests/test_avails_reach.py

Offline: no API key, no network, no PowerPoint. A recorded draft goes through
the real apply_draft_to_form, exactly as tier 1 does.

The scenario is the one the feature was specified against: three audiences
with stated avails (1,000,000 / 600,000 / 400,000 monthly), two options at
20% and 40% reach, $30 CPM, plus streaming retargeting and an NFL regular
season package. Option 1 should come out at 200K / 120K / 80K impressions
and Option 2 at double that.
"""

import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app  # noqa: E402

PASSED = FAILED = 0


def check(label, condition, detail=""):
    global PASSED, FAILED
    if condition:
        PASSED += 1
        print(f"    PASS  {label}")
    else:
        FAILED += 1
        print(f"    FAIL  {label}{('  ' + str(detail)) if detail else ''}")


AUDIENCES = [
    ("RETAIL Home Services Home Improvement", 1_000_000),
    ("DEMO Homeowner", 600_000),
    ("HH Income 100K Plus", 400_000),
]


def reach_line(segment, percent):
    return {"product": "premion_streaming_tv", "label": segment,
            "audience_track": segment, "cpm": 30,
            "allocation": {"percent_of_avails": percent, "avails_ref": segment}}


def build_draft(percent_by_option=(20, 40), with_avails=True, basis="monthly"):
    audiences = []
    for segment, avails in AUDIENCES:
        entry = {"segment": segment, "geo": "Washington, DC DMA"}
        if with_avails:
            entry["max_avails"] = avails * (3 if basis == "flight" else 1)
            entry["avails_basis"] = basis
        audiences.append(entry)
    options = []
    for n, percent in enumerate(percent_by_option, start=1):
        options.append({
            "name": f"Option {n} — {percent}% Reach",
            "total_budget": 100000, "breakout": "monthly",
            "media_plan_lines": (
                [reach_line(seg, percent) for seg, _ in AUDIENCES]
                + [{"product": "streaming_retargeting_display",
                    "allocation": {"flat_amount": 5000}},
                   {"product": "sport:nfl_reg", "allocation": {"flat_amount": 10000}}]
            ),
        })
    return {
        "client_name": "Fairmont Dermatology", "vertical": "healthcare", "market": "DC",
        "geo": "Washington, DC DMA",
        "flight_start": "2026-09-01", "flight_end": "2026-11-30",
        "agency_involved": False, "breakout": "monthly",
        "audiences": audiences, "options": options, "media_plan_lines": [],
        "sports": ["nfl_reg"], "attribution": [],
        "campaign_specs": {"goals": ["Grow new-patient volume"], "audience": ["Adults 35+"],
                           "geography": ["Washington, DC DMA"], "budget": ["$100,000"],
                           "placements": ["15s/30s CTV"], "timing": ["Sep - Nov"]},
        "unresolved": [], "unresolved_internal": [],
    }


class Stub:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply(draft):
    """Run the real apply_draft_to_form against an empty session."""
    real, stub = app.st, Stub()
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
        return dict(stub.session_state)
    finally:
        app.st = real


def avails_map(state):
    return {r["Audience"]: r[app.AVAILS_COLUMN_MONTHLY]
            for r in state.get("avails_seed_rows", [])}


def rows_of(state, option_index):
    return state["plan_options"][option_index]["rows"]


def reach_rows(state, option_index):
    names = {seg for seg, _ in AUDIENCES}
    return [r for r in rows_of(state, option_index)
            if any(r["Tactic"].endswith(name) for name in names)]


def test_avails_are_ingested():
    print("\n  Avails stated in the notes reach the table")
    state = apply(build_draft())
    got = avails_map(state)
    for segment, expected in AUDIENCES:
        check(f"{segment} = {expected:,}", got.get(segment) == expected, got.get(segment))
    check("stored as monthly", state.get("avails_basis") == app.AVAILS_BASIS_MONTHLY,
          state.get("avails_basis"))
    review = " ".join(state.get("draft_unresolved_internal", []) + state.get("draft_unresolved", []))
    check("nothing flagged about pulling avails",
          "Pull the real numbers" not in review, review)


def test_reach_impressions():
    print("\n  Reach lines price from avails, and Option 2 is double Option 1")
    state = apply(build_draft())
    # Monthly breakout, so rows hold per-month figures: 20% of the MONTHLY
    # avails. This is the number the specification named.
    expected_1 = [200_000, 120_000, 80_000]
    got_1 = [round(r["Impressions"]) for r in reach_rows(state, 0)]
    check(f"Option 1 impressions {expected_1}", got_1 == expected_1, got_1)
    expected_2 = [400_000, 240_000, 160_000]
    got_2 = [round(r["Impressions"]) for r in reach_rows(state, 1)]
    check(f"Option 2 impressions {expected_2}", got_2 == expected_2, got_2)
    check("Option 2 is exactly double Option 1",
          all(b == a * 2 for a, b in zip(got_1, got_2)), (got_1, got_2))

    # cost = impressions / 1000 * CPM, no agency markup in this scenario.
    costs = [round(r["Cost"]) for r in reach_rows(state, 0)]
    expected_costs = [round(i / 1000 * 30) for i in expected_1]
    check(f"costs follow from impressions and the $30 CPM {expected_costs}",
          costs == expected_costs, costs)
    check("the negotiated $30 CPM is on the rows",
          all(r["CPM"] == 30 for r in reach_rows(state, 0)),
          [r["CPM"] for r in reach_rows(state, 0)])
    check("reach rows drive from impressions, not cost",
          state["plan_options"][0]["driver"][:3] == [app.DRIVER_IMPRESSIONS] * 3,
          state["plan_options"][0]["driver"])


def test_other_lines_still_work():
    print("\n  The non-reach lines on the same option are untouched")
    state = apply(build_draft())
    tactics = [r["Tactic"] for r in rows_of(state, 0)]
    check("streaming retargeting line present",
          any("Retargeting" in t for t in tactics), tactics)
    check("NFL line present", any("NFL" in t for t in tactics), tactics)
    flat = [r for r in rows_of(state, 0) if "Retargeting" in r["Tactic"]][0]
    # $5,000 is a full-flight amount; a monthly breakout shows it per month.
    check("its flat allocation is $5,000 over the flight",
          round(flat["Cost"] * 3) == 5000, flat["Cost"])


def test_basis_invariance():
    """The property that makes the toggle safe."""
    print("\n  Reach is unaffected by which basis the notes used")
    monthly = apply(build_draft(basis="monthly"))
    flight = apply(build_draft(basis="flight"))
    check("a flight-basis figure is stored as its monthly equivalent",
          avails_map(flight) == avails_map(monthly), avails_map(flight))
    got_m = [round(r["Impressions"]) for r in reach_rows(monthly, 0)]
    got_f = [round(r["Impressions"]) for r in reach_rows(flight, 0)]
    check("identical impressions either way", got_m == got_f, (got_m, got_f))


def test_missing_avails_blank_and_flagged():
    print("\n  With no avails stated, the line stays but blank -- and is flagged")
    state = apply(build_draft(with_avails=False))
    check("every audience is in the table at 0",
          set(avails_map(state)) == {s for s, _ in AUDIENCES}
          and set(avails_map(state).values()) == {0}, avails_map(state))
    review = " ".join(state.get("draft_unresolved_internal", []))
    check("flagged for the seller to pull them", "avails system" in review, review)
    kept = reach_rows(state, 0)
    check("the reach lines are NOT dropped", len(kept) == 3, len(kept))
    check("their impressions are blank", all(round(r["Impressions"]) == 0 for r in kept),
          [r["Impressions"] for r in kept])
    check("their cost is blank", all(round(r["Cost"]) == 0 for r in kept),
          [r["Cost"] for r in kept])
    all_review = " ".join(state.get("draft_unresolved", []) + state.get("draft_unresolved_internal", []))
    check("and the reason is stated", "no avails" in all_review or "avails are 0" in all_review,
          all_review)


def test_percentage_over_100_is_flagged():
    print("\n  A percentage over 100 is flagged")
    state = apply(build_draft(percent_by_option=(150,)))
    review = " ".join(state.get("draft_unresolved", []) + state.get("draft_unresolved_internal", []))
    check("flagged as more than all of it", "more than all of it" in review, review)


def test_reach_and_budget_disagreement():
    print("\n  Reach wins over a budget it disagrees with, and says so")
    draft = build_draft(percent_by_option=(20,))
    # Reach comes to 400,000 impressions/month over three months at $30
    # against a stated $100,000 -- a real disagreement.
    state = apply(draft)
    review = " ".join(state.get("draft_unresolved", []) + state.get("draft_unresolved_internal", []))
    check("the disagreement is flagged", "don't agree" in review, review)
    check("and says the plan is built from the reach", "built from the reach" in review, review)
    got = [round(r["Impressions"]) for r in reach_rows(state, 0)]
    check("impressions still come from the reach", got == [200_000, 120_000, 80_000], got)


def test_budget_line_reports_derived_reach():
    print("\n  A budget line derives and reports reach without being driven by it")
    # The Plaza Motors case: the notes state a dollar budget and ask what
    # percent of the stated avails that buys -- a "flat_amount" line, not a
    # "percent_of_avails" one, even though the notes talk about reach.
    segment = "RETAIL Home Services Home Improvement"
    draft = {
        "client_name": "Plaza Motors", "vertical": "auto", "market": "DC",
        "geo": "Washington, DC DMA",
        "flight_start": "2026-09-01", "flight_end": "2026-09-30",
        "agency_involved": False, "breakout": "monthly",
        "audiences": [{"segment": segment, "geo": "Washington, DC DMA",
                       "max_avails": 1_700_000, "avails_basis": "monthly"}],
        "options": None, "total_budget": 4000,
        "media_plan_lines": [
            {"product": "premion_streaming_tv", "audience_track": segment, "cpm": 30,
             "allocation": {"flat_amount": 4000}},
        ],
        "sports": [], "attribution": [],
        "campaign_specs": {"goals": ["Drive showroom traffic"], "audience": ["Auto intenders"],
                           "geography": ["Washington, DC DMA"], "budget": ["$4,000"],
                           "placements": ["15s/30s CTV"], "timing": ["Sep"]},
        "unresolved": [], "unresolved_internal": [],
    }
    state = apply(draft)
    rows = rows_of(state, 0)
    check("exactly one plan row was produced", len(rows) == 1, rows)
    line = rows[0]
    expected_impressions = app.impressions_from_cost(4000, 30, 1.0)
    check("impressions come from the $4,000 budget at $30 CPM, not from the 1.7M avails",
          round(line["Impressions"]) == round(expected_impressions), line["Impressions"])
    check("cost is exactly the stated budget", round(line["Cost"]) == 4000, line["Cost"])
    review = " ".join(state.get("draft_unresolved", []) + state.get("draft_unresolved_internal", []))
    expected_pct = round(expected_impressions / 1_700_000 * 100.0)
    check("the derived reach percentage is reported for the reviewer, against the real avails "
          "figure, without having driven the line",
          f"{expected_pct}%" in review and "1,700,000" in review, review)


def test_unmatched_segment_keeps_its_avails():
    print("\n  An unrecognised segment with real avails keeps both")
    draft = build_draft()
    draft["audiences"].append({"segment": "Totally Made Up Segment",
                               "geo": "Washington, DC DMA",
                               "max_avails": 250_000, "avails_basis": "monthly"})
    state = apply(draft)
    got = avails_map(state)
    check("the rep's label is kept as typed", "Totally Made Up Segment" in got, list(got))
    check("with its avails attached", got.get("Totally Made Up Segment") == 250_000, got)
    review = " ".join(state.get("draft_unresolved_internal", []))
    check("and flagged as not a catalog name", "isn't a name from the audience catalog" in review,
          review)


def test_conversion_helpers():
    print("\n  Basis conversion, and no drift on an untouched row")
    check("monthly displays unchanged",
          app.avails_to_display(1_000_000, app.AVAILS_BASIS_MONTHLY, 3) == 1_000_000)
    check("flight displays x months",
          app.avails_to_display(1_000_000, app.AVAILS_BASIS_FLIGHT, 3) == 3_000_000)
    check("column header names monthly",
          app.avails_column_label(app.AVAILS_BASIS_MONTHLY, 3) == "Max Monthly Avails")
    check("column header names the flight and its length",
          app.avails_column_label(app.AVAILS_BASIS_FLIGHT, 3)
          == "Max Avails — Full Flight (3 months)",
          app.avails_column_label(app.AVAILS_BASIS_FLIGHT, 3))
    check("one month is singular",
          app.avails_column_label(app.AVAILS_BASIS_FLIGHT, 1).endswith("(1 month)"),
          app.avails_column_label(app.AVAILS_BASIS_FLIGHT, 1))

    # The drift case: a value that doesn't divide evenly by the month count,
    # flipped back and forth without being edited, must not move at all.
    stored = 333_333
    for months in (3, 7):
        value = stored
        for _ in range(5):
            shown = app.avails_to_display(value, app.AVAILS_BASIS_FLIGHT, months)
            value = app.restore_untouched_avails(
                value, shown, shown, app.AVAILS_BASIS_FLIGHT, months)
        check(f"unedited value survives 5 flips at {months} months ({stored:,})",
              value == stored, value)
    edited = app.restore_untouched_avails(
        333_333, 999_999, 1_500_000, app.AVAILS_BASIS_FLIGHT, 3)
    check("an edited value IS converted back", edited == 500_000, edited)


def test_toggle_in_the_real_form():
    """The editor round-trip, driven through the actual app.

    The helpers above are pure; this is the path that actually renders the
    table, converts for display and writes the monthly truth back. A drift
    bug would live here rather than in the arithmetic, so the invariant is
    asserted against the real widget: flip the basis repeatedly on a figure
    that does NOT divide evenly by the month count, and the stored value
    must not move at all.
    """
    print("\n  Flipping the basis in the real form changes nothing stored")
    from streamlit.testing.v1 import AppTest

    stored = [333_333, 600_000]          # 333,333 / 3 is not a whole number
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Matt"
    at.session_state["avails_seed_rows"] = [
        {"Audience": "RETAIL Home Services Home Improvement", "Geo": "Washington, DC DMA",
         app.AVAILS_COLUMN_MONTHLY: stored[0]},
        {"Audience": "DEMO Homeowner", "Geo": "Washington, DC DMA",
         app.AVAILS_COLUMN_MONTHLY: stored[1]},
    ]
    at.run()
    check("the table renders", not at.exception, at.exception)
    # AppTest's session_state proxy has no .get() -- it raises AttributeError
    # rather than returning a default.
    months = len(at.session_state["active_months"]
                 if "active_months" in at.session_state else [])
    check("the flight is more than one month (or the test proves nothing)",
          months > 1, months)

    for turn in range(4):
        at.session_state["avails_basis"] = (app.AVAILS_BASIS_FLIGHT if turn % 2 == 0
                                            else app.AVAILS_BASIS_MONTHLY)
        at.run()
        if at.exception:
            check(f"flip {turn + 1} renders", False, at.exception)
            return
    got = [r[app.AVAILS_COLUMN_MONTHLY] for r in at.session_state["avails_seed_rows"]]
    check(f"stored values survive 4 flips unchanged {stored}", got == stored, got)


def main():
    print("=" * 70)
    print("Avails ingestion, reach allocation, basis toggle")
    print("=" * 70)
    test_conversion_helpers()
    test_avails_are_ingested()
    test_reach_impressions()
    test_other_lines_still_work()
    test_basis_invariance()
    test_missing_avails_blank_and_flagged()
    test_percentage_over_100_is_flagged()
    test_reach_and_budget_disagreement()
    test_budget_line_reports_derived_reach()
    test_unmatched_segment_keeps_its_avails()
    test_toggle_in_the_real_form()
    print("\n" + "=" * 70)
    print(f"{PASSED} passed, {FAILED} failed")
    print("=" * 70)
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
