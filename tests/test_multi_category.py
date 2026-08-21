"""Dual-category segments -- offline. MOVERS is the first live case
(AFIRST Movers Fixers and DIYers, the LIFESTAGE home-buyer/likely-to-move
segments, and the workbook's CUSTOM New Home Buyer all carry "MOVERS,
LIFESTAGE"), and the plumbing has to genuinely support it, not just store
a decorative second value. Four things asserted, each requested directly:

    1. filtering by EITHER category returns the segment
    2. it appears ONCE, not twice, when BOTH categories are in scope
    3. the Claude categorize-confirm-and-store flow persists both
    4. counts/rankings don't double-count a dual-category segment

    python tests/test_multi_category.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import pandas as pd                            # noqa: E402

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import audience_catalog as ac                  # noqa: E402
import audience_usage_import as aui            # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    print("MOVERS is genuinely dual-category in the real (local-fallback) catalog")
    catalog = ac.load_local_catalog()
    movers_rows = catalog[catalog["segment"] == "AFIRST Movers Fixers and DIYers"]
    check("the fixture segment exists and carries both categories",
          not movers_rows.empty and movers_rows.iloc[0]["category"] == "MOVERS, LIFESTAGE",
          movers_rows[["segment", "category"]].to_dict("records") if not movers_rows.empty else None)
    segment = "AFIRST Movers Fixers and DIYers"
    cell = movers_rows.iloc[0]["category"]

    print("\n1. Filtering by EITHER category returns the segment")
    check("category_matches(cell, ['MOVERS']) is True",
          ac.category_matches(cell, ["MOVERS"]), cell)
    check("category_matches(cell, ['LIFESTAGE']) is True",
          ac.category_matches(cell, ["LIFESTAGE"]), cell)
    check("category_matches(cell, ['AUTO']) is False -- not a false positive on every category",
          not ac.category_matches(cell, ["AUTO"]), cell)
    check("all_categories() offers MOVERS and LIFESTAGE as separate picks, "
          "never a combined 'MOVERS, LIFESTAGE' option nobody would select",
          "MOVERS" in ac.all_categories(catalog) and "LIFESTAGE" in ac.all_categories(catalog)
          and "MOVERS, LIFESTAGE" not in ac.all_categories(catalog), ac.all_categories(catalog))
    # End-to-end through the actual Browse filter expression app.py uses.
    filtered_by_movers = catalog[catalog["category"].apply(lambda c: ac.category_matches(c, ["MOVERS"]))]
    filtered_by_lifestage = catalog[catalog["category"].apply(lambda c: ac.category_matches(c, ["LIFESTAGE"]))]
    check("the Browse filter's own expression finds the segment under MOVERS",
          segment in set(filtered_by_movers["segment"]), None)
    check("...and under LIFESTAGE too",
          segment in set(filtered_by_lifestage["segment"]), None)

    print("\n2. Appears ONCE, not twice, when BOTH categories are in scope "
          "(a vertical mapped to LIFESTAGE and MOVERS)")
    fake_hints = {"movers_test": ["MOVERS", "LIFESTAGE"]}
    real_hints = app.VERTICAL_CATEGORY_HINTS
    app.VERTICAL_CATEGORY_HINTS = fake_hints
    try:
        prioritized = app.prioritize_catalog(catalog, "movers_test")
    finally:
        app.VERTICAL_CATEGORY_HINTS = real_hints
    occurrences = (prioritized["segment"] == segment).sum()
    check("exactly one row for the dual-category segment, not one per matching category",
          occurrences == 1, occurrences)
    check("prioritize_catalog never changes the total row count -- a partition, not a union "
          "that could double-count",
          len(prioritized) == len(catalog), (len(prioritized), len(catalog)))

    print("\n3. The Claude categorize-confirm-and-store flow persists BOTH categories")
    # Mirrors exactly what render_update_audience_usage's Activate button
    # builds from the confirm table: Category + Category2 joined, handed
    # to db.py as one confirmed_categories entry.
    confirmed = {"Some New Mover Segment": {"category": ", ".join(["MOVERS", "LIFESTAGE"])}}
    overrides = aui.load_overrides()
    for seg, c in confirmed.items():
        overrides[aui.normalize(seg)] = {"action": "categorize", "category": c["category"], "split_into": []}
    wb = aui.WorkbookImport(
        rows=[aui.UsageRow(segment="Some New Mover Segment", impressions=1000)],
        dropped_no_data_targeting=0, sheet_name="x")
    report = aui.derive_catalog_updates(wb, existing_catalog=[], overrides=overrides)
    by_seg = {u.segment: u for u in report.updates}
    check("the stored category is the comma-joined pair, exactly as confirmed",
          by_seg["Some New Mover Segment"].category == "MOVERS, LIFESTAGE",
          by_seg.get("Some New Mover Segment"))

    print("\n4. Counts/rankings don't double-count a dual-category segment")
    check("len(catalog) counts the segment once -- one row, one comma-joined cell, "
          "never two rows",
          (catalog["segment"] == segment).sum() == 1, None)
    total_matches_style_count = len(catalog[catalog["category"].apply(
        lambda c: ac.category_matches(c, ["MOVERS", "LIFESTAGE"]))])
    # Every other LIFESTAGE-only or MOVERS-only row would ALSO match this
    # OR-filter, so this isn't a count of just the dual-category segment --
    # the real check is that IT contributes exactly one to that total, not two.
    check("a filter matching both of its categories still counts the segment once",
          catalog[catalog["category"].apply(lambda c: ac.category_matches(c, ["MOVERS", "LIFESTAGE"]))]
          ["segment"].tolist().count(segment) == 1, None)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A dual-category segment is reachable from either category, counted once when both "
          "are in scope, and the Claude confirm-and-store path writes the comma-joined pair "
          "verbatim -- the mechanism proven on MOVERS' own real segments, not a synthetic "
          "stand-in alone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
