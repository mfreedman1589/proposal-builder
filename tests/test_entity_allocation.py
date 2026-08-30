"""FLOW_REWORK_PLAN.md Phase 3, commit 5: allocation becomes per-entity,
rows stay exactly as today.

`_synthesize_group_lines`/`_allocate_group_rows` are pure functions (no
Streamlit, no session_state) over an option/groups_by_id -- tested directly
here, no AppTest needed. Four things proved, matching the settled design:

    1. A stated split (split_evenly) divides by ENTITY count, not row/group
       count -- two rows sharing one entity plus two ordinary rows is a
       3-way split, not a 4-way one.
    2. Only the entity's REPRESENTATIVE row (the first clean, selected row
       for it) gets written back; a sibling row for the same entity is left
       exactly as it was -- never silently split, never duplicated.
    3. avails_by_group_id/entity_avails_monthly are scoped to the ROWS
       ACTUALLY ON THE PLAN, never an entity's full membership -- the
       direct regression test for the leak this whole feature exists to
       prevent (Annapolis's Subaru 5mi/10mi radius tiers: with only the
       5mi row selected, its price must come from 165,687, never the
       unselected 602,647 sibling).
    4. A single-row entity (every proposal before this phase, and the
       common case after it) behaves exactly as today -- one line, no
       behavior change.

    python tests/test_entity_allocation.py
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


def clean_row(gid, cost=0.0):
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep - Nov", "Geo": "", "Targeting": "",
            "Impressions": 0.0, "CPM": 30.0, "Type": app.ROW_TYPE_RATE, "Cost": cost,
            "_group_ids": [gid]}


def main():
    print("split_evenly divides by ENTITY count, not row count -- two rows sharing "
          "one entity plus two ordinary rows is a 3-way split")
    g_5mi = tg.new_group(["AUTO Make Subaru"], avails_monthly=165_687, entity_id="toyota")
    g_10mi = tg.new_group(["AUTO Make Subaru"], avails_monthly=602_647, entity_id="toyota")
    g_denver = tg.new_group(["Homeowners"], avails_monthly=300_000)
    g_atlanta = tg.new_group(["Homeowners"], avails_monthly=200_000)
    groups_by_id = {g["id"]: g for g in (g_5mi, g_10mi, g_denver, g_atlanta)}

    option = app.new_plan_option("Option A", [
        clean_row(g_5mi["id"]), clean_row(g_10mi["id"]),
        clean_row(g_denver["id"]), clean_row(g_atlanta["id"]),
    ])
    opt_intent = {"total_budget": 12_000.0, "group_allocation": {"split_evenly": True},
                  "group_cpm": None, "default_targeting": ""}
    unresolved = app._allocate_group_rows(option, opt_intent, groups_by_id,
                                          "Sep - Nov", "", 3)
    costs = [r["Cost"] for r in option["rows"]]
    check("no unresolved notes", not unresolved, unresolved)
    check("exactly 3 rows got a real (non-zero) share -- one per entity, not per row",
          sum(1 for c in costs if c > 0) == 3, costs)
    check("the toyota entity's 10mi row (not the representative) is untouched -- still $0, "
          "never silently split or duplicated",
          option["rows"][1]["Cost"] == 0.0, option["rows"][1])
    check("the toyota entity's 5mi row (the representative) got a real share",
          option["rows"][0]["Cost"] > 0, option["rows"][0])
    check("all three real shares are equal (an even 3-way split of $12,000 -- $4,000 each)",
          abs(option["rows"][0]["Cost"] - 4000.0) < 0.02
          and abs(option["rows"][2]["Cost"] - 4000.0) < 0.02
          and abs(option["rows"][3]["Cost"] - 4000.0) < 0.02,
          costs)

    print("\nthe critical regression: an UNSELECTED sibling's avails must never leak in. "
          "Only the 5mi row is on the plan (10mi never even appears as a row) -- its price "
          "must come from 165,687, never the unselected 602,647 sibling")
    solo_option = app.new_plan_option("Option A", [clean_row(g_5mi["id"])])
    solo_intent = {"total_budget": 0, "group_allocation": {"percent_of_avails": 20},
                  "group_cpm": None, "default_targeting": ""}
    app._allocate_group_rows(solo_option, solo_intent, groups_by_id, "Sep - Nov", "", 3)
    # 20% of 165,687 monthly x 3 months x $30 CPM / 1000 = cost; back-solve the
    # impressions to confirm they trace to 165,687, not 602,647.
    expected_impressions = 165_687 * 3 * 0.20
    check("impressions are 20% of the 5mi group's OWN avails, not the 10mi sibling's",
          abs(solo_option["rows"][0]["Impressions"] - expected_impressions) < 1.0,
          (solo_option["rows"][0]["Impressions"], expected_impressions))

    print("\na single-row entity (no grouping at all) behaves exactly as today -- one line, "
          "the plain per-row share")
    plain_option = app.new_plan_option("Option A", [
        clean_row(g_denver["id"]), clean_row(g_atlanta["id"]),
    ])
    plain_intent = {"total_budget": 10_000.0, "group_allocation": {"split_evenly": True},
                    "group_cpm": None, "default_targeting": ""}
    app._allocate_group_rows(plain_option, plain_intent, groups_by_id, "Sep - Nov", "", 3)
    check("an ordinary 2-entity split evenly divides $10,000 into $5,000 each",
          abs(plain_option["rows"][0]["Cost"] - 5000.0) < 0.02
          and abs(plain_option["rows"][1]["Cost"] - 5000.0) < 0.02,
          [r["Cost"] for r in plain_option["rows"]])

    print("\n_synthesize_group_lines directly: one line per entity, avails_by_group_id "
          "keyed only to each entity's REPRESENTATIVE id")
    lines, avails_by_group_id = app._synthesize_group_lines(
        [g_5mi["id"], g_10mi["id"], g_denver["id"]], {"split_evenly": True}, None, groups_by_id)
    check("2 lines, not 3 -- toyota's pair collapsed to one waterfall line",
          len(lines) == 2, lines)
    check("the representative id is the FIRST given id for that entity (g_5mi, not g_10mi)",
          avails_by_group_id.get(g_5mi["id"]) == 602_647
          and g_10mi["id"] not in avails_by_group_id,
          avails_by_group_id)
    check("the representative's avails figure is the MAX across the entity's given members "
          "(602,647), even though it's stamped on the 5mi id",
          avails_by_group_id[g_5mi["id"]] == 602_647, avails_by_group_id)

    print("\nplan_lines_from_groups: entity-major ordering when an entity spans several "
          "groups; byte-identical to today when nothing is grouped")
    ordered_plain = app.plan_lines_from_groups([g_denver, g_atlanta], "")
    check("no entity grouping -> pure audience-major order, unchanged from before",
          [t[2] for t in ordered_plain] == [g_denver["id"], g_atlanta["id"]], ordered_plain)

    # Wilmington-shaped: three DIFFERENT audiences sharing one entity would
    # otherwise interleave with an unrelated fourth audience under pure
    # audience-major sort.
    ug1 = tg.new_group(["College Planning Parents"], entity_id="undergrad")
    unrelated = tg.new_group(["MBA Prospects"])
    ug2 = tg.new_group(["Prospective College Students"], entity_id="undergrad")
    ug3 = tg.new_group(["Higher Education Intender"], entity_id="undergrad")
    mixed = [ug1, unrelated, ug2, ug3]
    ordered_mixed = app.plan_lines_from_groups(mixed, "")
    ordered_ids = [t[2] for t in ordered_mixed]
    ug_positions = [ordered_ids.index(g["id"]) for g in (ug1, ug2, ug3)]
    check("the three undergrad-entity groups sort ADJACENT (consecutive positions), "
          "despite three different audiences and an unrelated group interleaved in the input",
          sorted(ug_positions) == list(range(min(ug_positions), min(ug_positions) + 3)),
          (ordered_ids, ug_positions))
    check("the entity cluster anchors at ug1's own audience-major position (first-seen)",
          ordered_ids[0] == ug1["id"], ordered_ids)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Allocation divides by entity count, writes back only to each entity's "
          "representative row, and is scoped strictly to the rows actually on the plan -- "
          "never an unselected sibling's avails. Entity-major ordering clusters a spread-out "
          "entity's rows without disturbing an ungrouped proposal's existing order.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
