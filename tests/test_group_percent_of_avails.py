"""FLOW_REWORK_PLAN.md Phase 3, commit 8: `percent_of_avails` as a
`group_allocation` type, and the dollar-figure-wins precedence rule.

A synthetic draft (no live API call -- the same `apply_draft_to_form`
Tier 1 drives, fed a hand-built dict instead of a recorded model response)
proves the mechanism itself: `resolve_drafted_lines`'s Stage 0 needed no
change (it already keyed off `avails_by_group_id`), only the prompt and
`_synthesize_group_lines`'s already-entity-aware plumbing (commit 5) had to
carry `percent_of_avails` through. The model-behavior side of the
precedence rule (does the LIVE model actually prefer a stated dollar figure
over an inferred avails percentage) is a live, Tier-2 concern -- see
tests/test_draft_live.py's `annapolis_group_selection` fixture, extended
for exactly this in the same commit.

    python tests/test_group_percent_of_avails.py
"""
import copy
import json
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
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _Stub:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def base_draft():
    d = json.loads((REPO / "tests" / "fixtures" / "hvac_two_option.draft.json").read_text(
        encoding="utf-8"))
    d.pop("_comment", None)
    d["options"] = []
    d["media_plan_lines"] = []
    # A single-month flight avoids spread_rows_over_months' own per-month
    # rounding entirely -- this test is about the allocation TYPE, not that
    # unrelated arithmetic, so pin months to exactly 1.
    d["flight_start"] = "2026-09-01"
    d["flight_end"] = "2026-09-30"
    return d


def main():
    print("group_allocation: {'percent_of_avails': 20} prices an entity at 20% of its "
          "SELECTED avails, cost derived in Python, no dollar figure anywhere in the "
          "model's own JSON")
    group = tg.new_group(["AUTO Make Subaru"], avails_monthly=165_687, group_id="g1")

    draft = base_draft()
    draft["options"] = [{
        "name": "Option A", "total_budget": 0, "breakout": "monthly",
        "media_plan_lines": [],
        "group_selection": {"mode": "all"},
        "group_allocation": {"percent_of_avails": 20},
        "group_cpm": 30,
        "group_selection_reason": "Notes say cover 20% of the available inventory.",
    }]
    # No dollar figure ANYWHERE in this draft's own JSON -- confirmed
    # structurally, not just by construction, since a future edit to this
    # test could otherwise reintroduce one silently.
    draft_text = json.dumps(draft["options"][0])
    check("sanity: the drafted option's own JSON contains no '$' or the word 'dollar'",
          "$" not in draft_text and "dollar" not in draft_text.lower(), draft_text)

    real, stub = app.st, _Stub()
    app.st = stub
    try:
        stub.session_state["targeting_groups"] = [group]
        # FLOW_REWORK_PLAN.md Phase 5: market/flight are band inputs now, not
        # drafted outputs -- apply_draft_to_form reads them from
        # session_state rather than the draft's own flight_start/flight_end.
        # Without this preset it falls back to DEFAULT_FLIGHT_START/END (a
        # 3-month flight), silently breaking base_draft()'s own single-month
        # pin and skewing the percent-of-avails day-count math.
        stub.session_state["market_choice"] = draft.get("market", "DC")
        stub.session_state["flight_start"] = date.fromisoformat(draft["flight_start"])
        stub.session_state["flight_end"] = date.fromisoformat(draft["flight_end"])
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real

    options = stub.session_state.get("plan_options") or []
    check("exactly one plan option was written", len(options) == 1, options)
    if options:
        rows = options[0]["rows"]
        check("exactly one row (one entity, one group)", len(rows) == 1, rows)
        if rows:
            row = rows[0]
            expected_impressions = 165_687 * 0.20   # one month, monthly breakout
            check("impressions are 20% of the group's OWN avails (33,137.4), "
                  "not a dollar figure the model invented",
                  abs(float(row["Impressions"]) - expected_impressions) < 1.0,
                  (row["Impressions"], expected_impressions))
            # FLOW_REWORK_PLAN.md Phase 4b: drafted rows are always net --
            # there is no markup for the model or the draft path to apply,
            # regardless of what the agency gross-up checkbox is set to.
            expected_cost = app.cost_from_impressions(expected_impressions, 30)
            check("cost is Python's own Impressions/1000 * CPM, matching the "
                  "negotiated group_cpm (30) -- never a figure copied from the model",
                  abs(float(row["Cost"]) - expected_cost) < 0.5,
                  (row["Cost"], expected_cost))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("percent_of_avails works as a group_allocation type exactly like it already did "
          "as an ordinary line's allocation -- Python computes every dollar and impression "
          "from the model's stated percentage, never the reverse.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
