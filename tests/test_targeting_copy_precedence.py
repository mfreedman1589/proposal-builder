"""A product with its own fixed Targeting copy -- Streaming Retargeting
("Retarget Exposed CTV Viewers"), every Live Sports package ("100% Live,
100% In-Game, 100% CTV") -- must keep it through every path that can touch a
plan row's Targeting cell, not just the ones that happen to seed it first.

Reported from live use (Andrea): a Streaming Retargeting line's Targeting was
getting overwritten with the audience stack. Root cause: `merge_plan_rows`
built the merged Targeting cell straight from the merged groups' own audience
labels, with no check for a product-level fixed copy at all -- the ONLY call
site that skipped the precedence `resolve_row_defaults` enforces everywhere
else. Fixed by factoring that check into `fixed_targeting_copy()` and having
`merge_plan_rows` consult it too, the same one-point-of-truth `_cell_unchanged`
already established for the D2 grid's own fold-back.

A follow-up sweep of the same function found `merge_plan_rows` also
overwrote a broadcast row's Geo unconditionally -- the SAME gap, one field
over: resolve_row_defaults holds a broadcast row's Geo via `current` for
exactly this reason (it's derived from the station's call sign, never
guessed), and a broadcast row isn't group-backed, so merging it with a
group-backed row would silently put that group's market name on it -- a
wrong DMA on a client's plan. Fixed alongside the Targeting case. split_plan_row
and quick_add_rows (the Add-lines panel) were checked in the same sweep and
already fully delegate to resolve_row_defaults -- no reimplementation gap
in either.

Four call sites asserted, matching every place `resolve_row_defaults`'s
docstring says the precedence has to hold:
    1. group re-seed        -- seed_media_plan_rows (targeting groups seeding
                                plan lines)
    2. merge                -- merge_plan_rows (the actual bug, both Targeting
                                and, found in the follow-up sweep, Geo)
    3. split                -- split_plan_row
    4. audience change      -- resolve_row_defaults's own soft-update path
                                (the shared-field re-seed uses this directly)

    python tests/test_targeting_copy_precedence.py
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
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


RETARGETING_TACTIC = "Streaming Retargeting - Display"
FLIGHT = "Sep 2026 - Nov 2026"


def _live_sports_tactic():
    """A real tactic label carrying LIVE_SPORTS_TARGETING, from whatever rate
    card this process actually loaded (live or fallback) -- never
    hardcoded, so this can't drift from the real table."""
    for label, copy in app.TARGETING_COPY_BY_LABEL.items():
        if copy == app.LIVE_SPORTS_TARGETING:
            return label
    return None


class _FakeSessionState(dict):
    def get(self, key, default=None):
        return dict.get(self, key, default)


class _FakeSt:
    def __init__(self, groups):
        self.session_state = _FakeSessionState(targeting_groups=groups)


def main():
    check("this rate card actually carries the Streaming Retargeting fixed copy",
          app.fixed_targeting_copy(RETARGETING_TACTIC) == app.STREAMING_RETARGETING_TARGETING,
          app.fixed_targeting_copy(RETARGETING_TACTIC))
    sports_tactic = _live_sports_tactic()
    check("this rate card carries at least one Live Sports package with the fixed copy",
          sports_tactic is not None, list(app.TARGETING_COPY_BY_LABEL.items())[:5])

    group_a = tg.new_group(["Segment A"], geo_def={"kind": "markets", "markets": ["Washington, DC DMA"]},
                           avails_monthly=1000)
    group_b = tg.new_group(["Segment B"], geo_def={"kind": "markets", "markets": ["Baltimore DMA"]},
                           avails_monthly=2000)

    print("1. group re-seed -- seed_media_plan_rows keeps the fixed copy for every triple")
    selections = {"products": {"streaming_retargeting": {"display": True}}}
    rows = app.seed_media_plan_rows(
        selections,
        [("Segment A", "Washington, DC DMA", group_a["id"]), ("Segment B", "Baltimore DMA", group_b["id"])],
        "Fallback Audience Stack", FLIGHT)
    check("two rows seeded, one per triple", len(rows) == 2, rows)
    check("every seeded row keeps the fixed copy, never the per-triple audience",
          all(r["Targeting"] == app.STREAMING_RETARGETING_TARGETING for r in rows), rows)

    print("\n2. merge -- the actual reported bug")
    row0 = {"Tactic": RETARGETING_TACTIC, "Flight": FLIGHT, "Geo": "Washington, DC DMA",
            "Targeting": app.STREAMING_RETARGETING_TARGETING, "Impressions": 1000.0, "CPM": 5.5,
            "Type": app.ROW_TYPE_RATE, "Cost": 100.0, "_group_ids": [group_a["id"]]}
    row1 = {"Tactic": "Premion Streaming TV", "Flight": FLIGHT, "Geo": "Baltimore DMA",
            "Targeting": "Segment B", "Impressions": 2000.0, "CPM": 30.0,
            "Type": app.ROW_TYPE_RATE, "Cost": 200.0, "_group_ids": [group_b["id"]]}
    option = app.new_plan_option("Option A", [row0, row1])
    real_st = app.st
    app.st = _FakeSt([group_a, group_b])
    try:
        app.merge_plan_rows(option, [0, 1])
    finally:
        app.st = real_st
    merged_targeting = option["rows"][0]["Targeting"]
    check("the merged survivor keeps Streaming Retargeting's fixed copy, "
          "not the merged groups' audience labels ('Segment A, Segment B')",
          merged_targeting == app.STREAMING_RETARGETING_TARGETING, merged_targeting)

    print("\n2b. merge -- a broadcast row's own schedule-derived copy survives a merge too")
    broadcast_row = {"Tactic": f"Premion {app.BROADCAST_TACTIC_MARKER}", "Flight": FLIGHT,
                     "Geo": "Washington DC DMA", "Targeting": "Prime access, 12x/week",
                     "Impressions": 500.0, "CPM": 20.0, "Type": app.ROW_TYPE_RATE, "Cost": 50.0}
    other_row = {"Tactic": "Premion Streaming TV", "Flight": FLIGHT, "Geo": "Baltimore DMA",
                "Targeting": "Segment B", "Impressions": 2000.0, "CPM": 30.0,
                "Type": app.ROW_TYPE_RATE, "Cost": 200.0, "_group_ids": [group_b["id"]]}
    bopt = app.new_plan_option("Option B", [broadcast_row, other_row])
    app.st = _FakeSt([group_b])
    try:
        app.merge_plan_rows(bopt, [0, 1])
    finally:
        app.st = real_st
    check("the merged survivor keeps the broadcast row's own Targeting",
          bopt["rows"][0]["Targeting"] == "Prime access, 12x/week", bopt["rows"][0]["Targeting"])
    check("...and its own Geo -- never the merged group's market name "
          "('Baltimore DMA'), a wrong DMA on a client's plan",
          bopt["rows"][0]["Geo"] == "Washington DC DMA", bopt["rows"][0]["Geo"])

    print("\n3. split -- each piece keeps the fixed copy")
    merged_row = {"Tactic": RETARGETING_TACTIC, "Flight": FLIGHT, "Geo": "Washington, DC DMA, Baltimore DMA",
                 "Targeting": app.STREAMING_RETARGETING_TARGETING, "Impressions": 3000.0, "CPM": 5.5,
                 "Type": app.ROW_TYPE_RATE, "Cost": 300.0, "_group_ids": [group_a["id"], group_b["id"]]}
    split_option = app.new_plan_option("Option C", [merged_row])
    app.st = _FakeSt([group_a, group_b])
    try:
        app.split_plan_row(split_option, 0)
    finally:
        app.st = real_st
    check("split produced two rows", len(split_option["rows"]) == 2, split_option["rows"])
    check("both split rows keep the fixed copy, not their own group's audience label",
          all(r["Targeting"] == app.STREAMING_RETARGETING_TARGETING for r in split_option["rows"]),
          split_option["rows"])

    print("\n4. audience change -- resolve_row_defaults's soft-update path (the shared-field re-seed)")
    defaults = app.resolve_row_defaults(
        RETARGETING_TACTIC, "Washington, DC DMA", "A brand new Campaign Specs Audience stack",
        FLIGHT, current=row0)
    check("a Campaign Specs Audience change never overwrites the fixed copy",
          defaults["Targeting"] == app.STREAMING_RETARGETING_TARGETING, defaults)

    if sports_tactic:
        print(f"\nsame precedence holds for a real Live Sports tactic ({sports_tactic!r})")
        sports_defaults = app.resolve_row_defaults(
            sports_tactic, "Washington, DC DMA", "Some other audience entirely", FLIGHT)
        check("Live Sports keeps its fixed copy too",
              sports_defaults["Targeting"] == app.LIVE_SPORTS_TARGETING, sports_defaults)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Every product carrying a fixed Targeting copy keeps it through a group re-seed, "
          "a merge, a split, and an audience change -- merge_plan_rows was the one call site "
          "that skipped the precedence resolve_row_defaults enforces everywhere else.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
