"""FLOW_REWORK_PLAN.md Phase 3, commit 10: "Show Label in plan".

Off (the default): the preview and the generated deck are byte-identical
to before this phase -- entity_prefixed_targeting is a no-op. On: a group-
owned line's Targeting text gets its entity's Label prefixed, on both the
app's own preview and the deck payload -- and never leaks into the
editable grid's own `row["Targeting"]` value, which stays exactly what a
rep typed or the app seeded.

    python tests/test_show_label.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def group_row(gid, targeting="Homeowners"):
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep - Nov", "Geo": "Denver",
            "Targeting": targeting, "Impressions": 100000.0, "CPM": 30.0,
            "Type": app.ROW_TYPE_RATE, "Cost": 3000.0, "_group_ids": [gid]}


def main():
    print("entity_prefixed_targeting: pure-function behavior")
    g_labeled = tg.new_group(["Homeowners"], group_id="g1", entity_label="Toyota of Annapolis")
    g_unlabeled = tg.new_group(["Homeowners"], group_id="g2")
    groups_by_id = {"g1": g_labeled, "g2": g_unlabeled}
    row_labeled = group_row("g1")
    row_unlabeled = group_row("g2")
    hand_typed_row = {"Tactic": "Premion Streaming TV", "Flight": "", "Geo": "", "Targeting": "Hand-typed"}

    check("toggle OFF: plain Targeting, byte-identical, regardless of entity label",
          app.entity_prefixed_targeting(row_labeled, groups_by_id, False) == "Homeowners")
    check("toggle ON, labeled entity: prefixed",
          app.entity_prefixed_targeting(row_labeled, groups_by_id, True)
          == "Toyota of Annapolis | Homeowners")
    check("toggle ON, unlabeled entity: falls through to plain Targeting",
          app.entity_prefixed_targeting(row_unlabeled, groups_by_id, True) == "Homeowners")
    check("toggle ON, no _group_ids at all (hand-typed/drafted/quick-add row): falls through",
          app.entity_prefixed_targeting(hand_typed_row, groups_by_id, True) == "Hand-typed")
    check("original row dict is untouched either way -- row['Targeting'] never mutated",
          row_labeled["Targeting"] == "Homeowners", row_labeled)

    print("\ncompute_plan_totals: toggle OFF is byte-identical to before this phase")
    rows = [group_row("g1"), group_row("g2", targeting="Business owners")]
    totals_off = app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov",
                                         groups_by_id=groups_by_id, show_entity_label=False)
    check("no show_entity_label kwarg at all also defaults off (byte-identical call sites)",
          app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov",
                                  groups_by_id=groups_by_id) == totals_off)
    targetings_off = [r["targeting"] for r in totals_off["preview_rows"]]
    check("targeting text is the plain, unprefixed text with the toggle off",
          targetings_off == ["Homeowners", "Business owners"], targetings_off)

    print("\ncompute_plan_totals: toggle ON prefixes only the labeled entity's line")
    totals_on = app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov",
                                        groups_by_id=groups_by_id, show_entity_label=True)
    targetings_on = [r["targeting"] for r in totals_on["preview_rows"]]
    check("the labeled entity's line is prefixed",
          targetings_on[0] == "Toyota of Annapolis | Homeowners", targetings_on)
    check("the unlabeled entity's line is untouched",
          targetings_on[1] == "Business owners", targetings_on)
    check("nothing else about the totals differs between on and off "
          "(same cost/impressions either way)",
          {k: v for k, v in totals_on.items() if k != "preview_rows"}
          == {k: v for k, v in totals_off.items() if k != "preview_rows"})
    check("the ORIGINAL rows list is untouched by either call -- toggling never leaks "
          "into row['Targeting']",
          rows[0]["Targeting"] == "Homeowners" and rows[1]["Targeting"] == "Business owners",
          rows)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("\"Show Label in plan\" prefixes a group-owned line's preview/deck Targeting text "
          "with its entity's Label when on, is a byte-identical no-op when off (the default), "
          "and never mutates the underlying row.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
