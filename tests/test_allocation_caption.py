"""FLOW_REWORK_PLAN.md Phase 3, commit 8, decision 12: the allocation-basis
caption (`app.describe_group_allocation`) -- a read-only line naming which
`group_allocation` fired, replacing a considered-and-rejected budget-split
COLUMN (the Cost column already IS the split; a second display of the same
number can only drift from it).

Pure function, no Streamlit -- tested directly.

    python tests/test_allocation_caption.py
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

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def group_row(gid, cost=1000.0):
    return {"Tactic": "Premion Streaming TV", "Flight": "", "Geo": "", "Targeting": "",
            "Impressions": 0.0, "CPM": 30.0, "Type": app.ROW_TYPE_RATE, "Cost": cost,
            "_group_ids": [gid]}


def main():
    g1 = tg.new_group(["A"], avails_monthly=100_000, group_id="g1")
    g2 = tg.new_group(["B"], avails_monthly=200_000, group_id="g2")
    groups_by_id = {"g1": g1, "g2": g2}
    rows = [group_row("g1"), group_row("g2")]
    dirty = [False, False]

    print("caption text for each of the five allocation types")
    cases = [
        ({"split_evenly": True}, "Budget split evenly across 2 entities"),
        ({"percent_of_total": 30}, "30% of budget applied to each entity"),
        ({"percent_of_remainder": 40}, "40% of the remaining budget applied to each entity"),
        ({"flat_amount": 5000}, "$5,000 flat applied to each entity"),
        ({"percent_of_avails": 20}, "20% of avails applied to each entity"),
    ]
    for allocation, expected_prefix in cases:
        text = app.describe_group_allocation({"group_allocation": allocation}, groups_by_id, rows, dirty)
        check(f"{next(iter(allocation))}: caption reads {expected_prefix!r}",
              text is not None and text.startswith(expected_prefix), text)

    print("\nno opt_intent / no allocation -> no caption at all (nothing to explain)")
    check("None opt_intent", app.describe_group_allocation(None, groups_by_id, rows, dirty) is None)
    check("empty group_allocation",
          app.describe_group_allocation({"group_allocation": {}}, groups_by_id, rows, dirty) is None)

    print("\nthe 'N unmerged lines' clause fires only when an entity truly has 2+ "
          "SEPARATE selected rows -- not for an ordinary one-row-per-entity plan")
    text = app.describe_group_allocation(
        {"group_allocation": {"split_evenly": True}}, groups_by_id, rows, dirty)
    check("no 'unmerged lines' clause when every entity has exactly one row",
          "unmerged" not in text, text)

    g1b = tg.new_group(["A"], avails_monthly=150_000, entity_id="g1", group_id="g1b")
    groups_by_id_shared = {"g1": g1, "g1b": g1b, "g2": g2}
    rows_shared = [group_row("g1"), group_row("g1b"), group_row("g2")]
    dirty_shared = [False, False, False]
    text_shared = app.describe_group_allocation(
        {"group_allocation": {"split_evenly": True}}, groups_by_id_shared, rows_shared, dirty_shared)
    check("2 real entities now (g1/g1b share one), not 3",
          text_shared.startswith("Budget split evenly across 2 entities"), text_shared)
    check("the clause fires for the entity with 2 unmerged rows",
          "has 2 unmerged lines -- full share applied to the first" in text_shared, text_shared)

    print("\na MERGED row (one row, two _group_ids) is already one line -- never flagged "
          "as 'unmerged'")
    merged_row = {"Tactic": "Premion Streaming TV", "Flight": "", "Geo": "", "Targeting": "",
                 "Impressions": 0.0, "CPM": 30.0, "Type": app.ROW_TYPE_RATE, "Cost": 2000.0,
                 "_group_ids": ["g1", "g1b"]}
    rows_merged = [merged_row, group_row("g2")]
    dirty_merged = [False, False]
    text_merged = app.describe_group_allocation(
        {"group_allocation": {"split_evenly": True}}, groups_by_id_shared, rows_merged, dirty_merged)
    check("no 'unmerged lines' clause -- the shared entity is already one (merged) row",
          "unmerged" not in text_merged, text_merged)

    print("\nthe caption drops its claim once any group-owned row for this option "
          "has been hand-edited (dirty)")
    dirty_one_edited = [True, False, False]
    text_dirty = app.describe_group_allocation(
        {"group_allocation": {"split_evenly": True}}, groups_by_id_shared, rows_shared, dirty_one_edited)
    check("caption is None once a group-owned row is dirty -- never asserts something "
          "the grid no longer shows",
          text_dirty is None, text_dirty)

    print("\nthe caption is never derived by reading Cost back off the rows -- changing "
          "every row's Cost changes nothing about what the caption says")
    rows_diff_cost = [group_row("g1", cost=999999.0), group_row("g1b", cost=1.0), group_row("g2", cost=0.0)]
    text_diff_cost = app.describe_group_allocation(
        {"group_allocation": {"split_evenly": True}}, groups_by_id_shared, rows_diff_cost, dirty_shared)
    check("byte-identical caption regardless of the rows' own Cost values",
          text_diff_cost == text_shared, (text_diff_cost, text_shared))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The allocation-basis caption names the allocation type in plain language, flags "
          "an entity with unmerged sibling rows, and is derived strictly from the stored "
          "allocation intent -- never from the rows' own Cost, and never once any of them "
          "has been hand-edited.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
