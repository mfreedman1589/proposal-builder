"""Sweep for the same stale-widget-key class bug 2 was: a keyed
`st.data_editor` whose key updates on some triggers but not on the row-count
change that actually needs it, so a row added or removed via the grid's own
"+"/"-" can outlive the key it was typed under.

Three `st.data_editor` calls exist in app.py:
    1. `usage_categorize_editor_*` -- fixed row count (one row per unresolved
       workbook segment, no `num_rows="dynamic"` at all). Never grows or
       shrinks within one upload's lifetime, so this class of bug can't
       reach it. Not vulnerable; no fix needed.
    2. `avails_editor_*` (D2) -- the original bug 2 fix: the fold-back now
       bumps `avails_version` (and reruns) whenever its own row count
       changes. Covered by tests/test_avails_grid_manual_row.py.
    3. `media_plan_editor_*_*` -- THIS FILE. Its key already updates on many
       triggers (duplicate line, merge, split, the quick-add panel, a
       product toggle), all of which bump `option["version"]` deliberately.
       But a row added or deleted through the grid's OWN "+"/"-" relied
       entirely on `reconcile_plan_rows`' `recomputed_any` return value to
       trigger that bump -- which only fires when a number actually changes
       via CPM recompute. A flat-fee row (recompute_row is a no-op for
       ROW_TYPE_FLAT_FEE) or a Cost that happens to already equal its own
       derivation added via the grid's own "+" would leave `recomputed_any`
       False, the version un-bumped, and the same stale-key risk avails had.

Fixed by having `reconcile_plan_rows` return True whenever the row COUNT
changed, independent of whether anything numeric also changed -- the same
deliberate, row-count-driven trigger the avails fix uses, rather than relying
on a numeric side effect that only correlates with it.

    python tests/test_media_plan_grid_row_count.py
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)
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


def main():
    print("a row added via the grid's own '+' forces a rerun even when nothing numeric changed")
    row0 = {"Tactic": "Premion Streaming TV", "Flight": "Sep 2026 - Nov 2026",
            "Geo": "Washington, DC DMA", "Targeting": "Segment A",
            "Impressions": 1000.0, "CPM": 30.0, "Type": app.ROW_TYPE_RATE, "Cost": 30.0}
    option = app.new_plan_option("Option A", [row0])
    # A brand-new FLAT FEE row, fully typed in by the rep -- recompute_row is
    # a documented no-op for a flat fee, so this is exactly the case that
    # produces recomputed_any=False on its own.
    new_flat_fee_row = {"Tactic": "Custom Setup Fee", "Flight": "Sep 2026 - Nov 2026",
                        "Geo": "Washington, DC DMA", "Targeting": "Segment A",
                        "Impressions": 0.0, "CPM": 0.0, "Type": app.ROW_TYPE_FLAT_FEE,
                        "Cost": 500.0}
    changed = app.reconcile_plan_rows(option, [row0, new_flat_fee_row])
    check("reconcile_plan_rows signals a change purely from the row count growing",
          changed is True, changed)
    check("the new row actually landed in option['rows']",
          len(option["rows"]) == 2 and option["rows"][1]["Tactic"] == "Custom Setup Fee",
          option["rows"])

    print("\na row deleted via the grid's own '-' forces a rerun too")
    option2 = app.new_plan_option("Option B", [row0, dict(row0, Tactic="Second Line")])
    changed2 = app.reconcile_plan_rows(option2, [row0])
    check("reconcile_plan_rows signals a change from the row count shrinking",
          changed2 is True, changed2)

    print("\nan ordinary cell edit with no row-count change still recomputes and signals correctly")
    option3 = app.new_plan_option("Option C", [dict(row0)])
    edited = [dict(row0, CPM=40.0)]  # CPM edit -- Cost should re-derive from it
    changed3 = app.reconcile_plan_rows(option3, edited)
    check("a CPM edit alone (same row count) still triggers recompute and signals True",
          changed3 is True, changed3)
    check("Cost re-derived from the new CPM",
          option3["rows"][0]["Cost"] != row0["Cost"], option3["rows"][0])

    print("\na genuinely untouched rerun (same rows, nothing edited) signals no change")
    option4 = app.new_plan_option("Option D", [dict(row0)])
    changed4 = app.reconcile_plan_rows(option4, [dict(row0)])
    check("no row-count change and no numeric change -- no forced rerun",
          changed4 is False, changed4)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The media plan grid's row-count change now forces a version bump and rerun "
          "deterministically, the same way the D2 avails grid's does, instead of relying "
          "on a numeric recompute that only correlates with it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
