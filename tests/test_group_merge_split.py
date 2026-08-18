"""Merging or splitting a plan line must never alter a targeting group --
the load-bearing assertion of Phase 3 (geo_targeting_roadmap.md D, "Remaining
phases (3-9)").

    python tests/test_group_merge_split.py

Offline, no Streamlit UI driven -- `merge_plan_rows`/`split_plan_row` only
read `st.session_state["targeting_groups"]` (to render a group's label), so
`app.st` is swapped for a plain stub the same way `test_form_state.py`'s
`load_drafted_form` does for `apply_draft_to_form`. What matters is proven
directly on the `option`/`groups` objects: identity, not the rendered page.

Also proves the shared-list hazard the roadmap calls out: `copy_plan_option`
and the duplicate-line button both clone a row with a shallow `dict(r)`, so a
duplicated row's `_group_ids` is the SAME LIST OBJECT as its source until
something reassigns it. If a merge ever mutated that list in place instead
of reassigning it, duplicating a row and then merging it with something else
would corrupt the id list of every other row still holding that same object.
"""
import copy
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

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
    """Swaps in for `app.st` -- only `session_state` is touched by
    merge_plan_rows/split_plan_row, and only to read `targeting_groups`."""
    def __init__(self, groups):
        self.session_state = {"targeting_groups": groups}
        self.secrets = {}


def with_stub(groups, fn):
    real_st, app.st = app.st, _Stub(groups)
    try:
        fn()
    finally:
        app.st = real_st


def six_line_option():
    """2 audiences x 3 markets, one group per line, one product -- the shape
    `plan_lines_from_groups` + `seed_media_plan_rows` produce for a real
    multi-audience, multi-market buy. CPM=20, markup=1.0 throughout so
    Cost = Impressions/1000 * CPM exactly, with no rounding to chase."""
    audiences = ["Homeowners", "In-market for windows"]
    markets = ["Denver", "Atlanta", "Phoenix"]
    groups, rows = [], []
    for color_i, (audience, market) in enumerate(
            (a, m) for a in audiences for m in markets):
        group = tg.new_group([audience], geo_def={"kind": "text", "label": market},
                             avails_monthly=100000, color=tg.assign_color(color_i))
        groups.append(group)
        rows.append({
            "Tactic": "Premion Streaming TV", "Flight": "Jan - Mar",
            "Geo": market, "Targeting": audience,
            "Impressions": 100000.0, "CPM": 20.0,
            "Type": app.ROW_TYPE_RATE, "Cost": 2000.0,
            "_group_ids": [group["id"]],
        })
    return app.new_plan_option("Option A", rows), groups


def main():
    print("merging 2 of 6 lines")
    option, groups = six_line_option()
    original_groups = copy.deepcopy(groups)

    with_stub(groups, lambda: app.merge_plan_rows(option, [0, 1], 1.0))

    check("6 lines become 5", len(option["rows"]) == 5, len(option["rows"]))
    check("targeting_groups is untouched by the merge (deep-equal)",
          groups == original_groups, groups)
    merged = option["rows"][0]
    check("the surviving row carries both merged ids",
          app.group_ids_of(merged) == [g["id"] for g in groups[:2]], app.group_ids_of(merged))
    check("Impressions summed", merged["Impressions"] == 200000.0, merged["Impressions"])
    check("Cost summed", merged["Cost"] == 4000.0, merged["Cost"])
    check("Geo joins both groups' own labels", merged["Geo"] == "Denver, Atlanta", merged["Geo"])
    check("Targeting joins both groups' own labels",
          merged["Targeting"] == "Homeowners, Homeowners", merged["Targeting"])
    check("the merged row is marked dirty", option["dirty"][0] is True, option["dirty"][0])
    check("CPM re-derived consistent with the summed cost/impressions",
          abs(merged["CPM"] - 20.0) < 1e-6, merged["CPM"])

    print("\nsplitting the merged line back apart")
    with_stub(groups, lambda: app.split_plan_row(option, 0))

    check("back to 6 lines", len(option["rows"]) == 6, len(option["rows"]))
    check("targeting_groups is STILL untouched by the split (deep-equal)",
          groups == original_groups, groups)
    split_rows = option["rows"][0:2]
    check("each split row carries exactly one of the merged ids, in order",
          [app.group_ids_of(r) for r in split_rows] == [[groups[0]["id"]], [groups[1]["id"]]],
          [app.group_ids_of(r) for r in split_rows])
    check("split rows' Impressions sum back to the merged total",
          sum(r["Impressions"] for r in split_rows) == 200000.0,
          [r["Impressions"] for r in split_rows])
    check("split rows' Cost sums back to the merged total",
          sum(r["Cost"] for r in split_rows) == 4000.0, [r["Cost"] for r in split_rows])
    check("both split rows are dirty",
          option["dirty"][0] is True and option["dirty"][1] is True,
          (option["dirty"][0], option["dirty"][1]))
    check("split rows carry their OWN group's Geo/Targeting, not the merged joined text",
          [r["Geo"] for r in split_rows] == ["Denver", "Atlanta"] and
          all(r["Targeting"] == "Homeowners" for r in split_rows),
          [(r["Geo"], r["Targeting"]) for r in split_rows])

    print("\nsplitting a row with one id (or none) is a no-op")
    option2, groups2 = six_line_option()
    before = copy.deepcopy(option2)
    with_stub(groups2, lambda: app.split_plan_row(option2, 0))
    check("nothing changed", option2 == before, option2)

    print("\nthe shared-list hazard: a duplicated row shares its source's _group_ids "
          "object until something reassigns it")
    option3, groups3 = six_line_option()
    groups3_snapshot = copy.deepcopy(groups3)
    # Exactly what the "Duplicate line" button does: a shallow dict(r) copy.
    duplicate = dict(option3["rows"][0])
    check("the duplicate's _group_ids IS the same list object as its source",
          duplicate["_group_ids"] is option3["rows"][0]["_group_ids"])
    option3["rows"] = option3["rows"] + [duplicate]
    option3["dirty"] = option3["dirty"] + [True]
    option3["driver"] = option3["driver"] + [option3["driver"][0]]

    untouched_row = option3["rows"][1]           # not part of the merge below
    untouched_gids_obj = untouched_row["_group_ids"]

    with_stub(groups3, lambda: app.merge_plan_rows(option3, [0, 6], 1.0))

    check("merging a row with its own duplicate concatenates the SAME id twice",
          app.group_ids_of(option3["rows"][0]) == [groups3[0]["id"], groups3[0]["id"]],
          app.group_ids_of(option3["rows"][0]))
    check("the duplicated id is de-duplicated in the rendered label",
          option3["rows"][0]["Geo"] == "Denver" and
          option3["rows"][0]["Targeting"] == "Homeowners",
          (option3["rows"][0]["Geo"], option3["rows"][0]["Targeting"]))
    check("an unrelated row's _group_ids is untouched, same object, same value",
          option3["rows"][1]["_group_ids"] is untouched_gids_obj and
          option3["rows"][1]["_group_ids"] == [groups3[1]["id"]],
          option3["rows"][1]["_group_ids"])
    check("targeting_groups deep-equals its own pre-merge snapshot",
          groups3 == groups3_snapshot, groups3)

    print("\na merged row survives an UNRELATED product's toggle off, then back on")
    option4, groups4 = six_line_option()
    with_stub(groups4, lambda: app.merge_plan_rows(option4, [0, 1], 1.0))
    merged_ids_before = app.group_ids_of(option4["rows"][0])

    # A previously-seeded Dynamic Video Ads flat-fee row -- the product this
    # scenario is about to toggle off and back on, NOT the merged Premion row.
    option4["rows"].append({"Tactic": app.DYNAMIC_AD_LINE_LABEL, "Flight": "Jan - Mar",
                            "Geo": "Denver", "Targeting": app.DYNAMIC_AD_TARGETING,
                            "Impressions": 0.0, "CPM": 0.0,
                            "Type": app.ROW_TYPE_FLAT_FEE, "Cost": 500.0})
    option4["dirty"].append(False)
    option4["driver"].append(app.DRIVER_IMPRESSIONS)

    # Five "Premion Streaming TV" rows exist at this point -- the merged one
    # plus the four groups that weren't merged -- so the merged row has to be
    # picked out by its OWN ids, not by Tactic, which every one of them shares.
    def by_merged_ids(rows_):
        return [r for r in rows_ if app.group_ids_of(r) == merged_ids_before]

    previously_seeded = ["Premion Streaming TV", app.DYNAMIC_AD_LINE_LABEL]
    fresh_rows_without = [{"Tactic": "Premion Streaming TV", "Flight": "Jan - Mar",
                           "Geo": "Denver", "Targeting": "x", "Impressions": 0.0,
                           "CPM": 20.0, "Type": app.ROW_TYPE_RATE, "Cost": 0.0}]
    app._apply_product_diff(option4, fresh_rows_without, previously_seeded)
    dynamic_rows = [r for r in option4["rows"] if r.get("Tactic") == app.DYNAMIC_AD_LINE_LABEL]
    check("Dynamic Video Ads was removed", not dynamic_rows, dynamic_rows)
    check("all 5 Premion rows survived (only the unrelated product was touched)",
          len([r for r in option4["rows"] if r.get("Tactic") == "Premion Streaming TV"]) == 5,
          len(option4["rows"]))
    check("the merged Premion row survived, ids intact",
          len(by_merged_ids(option4["rows"])) == 1,
          [app.group_ids_of(r) for r in option4["rows"]])

    fresh_rows_with = fresh_rows_without + [{
        "Tactic": app.DYNAMIC_AD_LINE_LABEL, "Flight": "Jan - Mar", "Geo": "Denver",
        "Targeting": app.DYNAMIC_AD_TARGETING, "Impressions": 0.0, "CPM": 0.0,
        "Type": app.ROW_TYPE_FLAT_FEE, "Cost": app.DYNAMIC_AD_DEFAULT_FEE}]
    app._apply_product_diff(option4, fresh_rows_with, ["Premion Streaming TV"])
    dynamic_rows = [r for r in option4["rows"] if r.get("Tactic") == app.DYNAMIC_AD_LINE_LABEL]
    check("Dynamic Video Ads came back", len(dynamic_rows) == 1, dynamic_rows)
    check("the merged Premion row still has its ids -- the tactic-prefix filter "
          "never touched a row that wasn't the removed/re-added tactic",
          len(by_merged_ids(option4["rows"])) == 1,
          [app.group_ids_of(r) for r in option4["rows"]])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Merge and split change what a plan LINE says. A targeting group never "
          "moves.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
