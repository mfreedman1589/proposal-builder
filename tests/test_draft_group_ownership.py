"""Drafting's side of "groups own the plan": once real targeting groups
exist, `apply_draft_to_form` never invents its own `media_plan_lines`/
`targeting_groups` for Premion Streaming TV -- it contributes budget,
allocation, products, flight and attribution, and those apply to selected
group-derived lines. The audiences[]-driven "build groups from scratch"
path stays exactly as it always has when no real groups exist.

Runs entirely offline through `apply_draft_to_form` directly (the same
`_StubSt`/`apply_draft` pattern tests/test_draft_regression.py uses) --
this is drafting's OWN behavior under test, not the reconciler or the D2
grid (see tests/test_group_plan_selection.py and
tests/test_group_plan_order_invariance.py for those).

    python tests/test_draft_group_ownership.py
"""
import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import targeting_groups as tg                  # noqa: E402
import group_scenario_fixtures as gsf          # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply_draft(draft, preset_state=None):
    """Same pattern as tests/test_draft_regression.py's own apply_draft --
    `preset_state` models a form the seller was already working in."""
    real = app.st
    stub = _StubSt()
    stub.session_state.update(preset_state or {})
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return stub.session_state


SCN = gsf.build_annapolis()
for _g in SCN["groups"]:
    _g.pop("_flight_label", None)
TEMPLATE_GROUPS = SCN["groups"]
MAKES = ["Subaru", "Hyundai", "Volvo", "Genesis"]


def all_groups(selected_ids=()):
    selected = set(selected_ids)
    return [dict(g, include_in_plan=(g["id"] in selected), include_locked=False)
           for g in TEMPLATE_GROUPS]


def ids_for(groups, make):
    return {g["id"] for g in groups if any(make.lower() in t.lower() for t in g["terms"])}


def base_preset(groups):
    return {
        "targeting_groups": groups,
        "flight_start": SCN["flight_start"], "flight_end": SCN["flight_end"],
    }


def _d(value):
    return value.isoformat() if hasattr(value, "isoformat") else str(value)


def make_draft(total_budget, group_selection=None, group_allocation=None, **overrides):
    draft = {
        "client_name": SCN["client_name"], "vertical": "auto", "market": "DC",
        "agency_involved": SCN["agency_involved"],
        "flight_start": _d(SCN["flight_start"]), "flight_end": _d(SCN["flight_end"]),
        "total_budget": total_budget, "breakout": "full_flight",
        "media_plan_lines": [],
        "group_selection": group_selection, "group_allocation": group_allocation,
        "group_selection_reason": "",
        "audiences": [], "attribution": [], "sports": [],
        "unresolved": [], "unresolved_internal": [],
    }
    draft.update(overrides)
    return draft


def group_ids_in(state):
    return {gid for opt in state["plan_options"] for r in opt["rows"] for gid in app.group_ids_of(r)}


def main():
    print("=" * 78)
    print("Groups exist -> the draft invents no lines; targeting_groups is untouched")
    print("=" * 78)
    groups = all_groups()
    preset = base_preset(groups)
    state = apply_draft(make_draft(20000, {"mode": "all"}, {"split_evenly": True}), preset)
    check("no avails_seed_rows written (never overwrites the real avails table)",
          "avails_seed_rows" not in state, list(state.keys()))
    restored = state["targeting_groups"]
    check("same 8 groups, same ids", {g["id"] for g in restored} == {g["id"] for g in groups},
          None)
    # Structurally byte-identical except include_in_plan/include_locked --
    # the direct guard on the bug that made import-then-draft destroy a
    # real import's resolved geo_def.
    by_id = {g["id"]: g for g in groups}
    mismatches = []
    for g in restored:
        orig = by_id[g["id"]]
        for key in ("terms", "op", "geo_def", "resolved_zips", "resolved_markets",
                    "avails_monthly", "color"):
            if g.get(key) != orig.get(key):
                mismatches.append((g["id"], key))
    check("every other field is structurally byte-identical to the pre-draft group",
          not mismatches, mismatches)
    check("all 8 pre-selected (mode=all)", all(g["include_in_plan"] for g in restored), restored)
    check("all 8 lines are group-owned (plan invented no lines of its own)",
          len(group_ids_in(state)) == 8, group_ids_in(state))

    print("\n" + "=" * 78)
    print("No groups -> unchanged: drafting still invents lines from scratch")
    print("=" * 78)
    state2 = apply_draft(make_draft(30000, None, None, audiences=[
        {"segment": "AUTO Intenders", "geo": "Denver"}]))
    check("targeting_groups WAS written (no real groups existed)",
          "targeting_groups" in state2, list(state2.keys()))
    check("no draft_plan_intent (only meaningful once groups own the plan)",
          "draft_plan_intent" not in state2, None)

    print("\n" + "=" * 78)
    print("Notes naming 2 of 4 makes pre-select exactly those, disclosed in unresolved")
    print("=" * 78)
    subaru_ids = ids_for(TEMPLATE_GROUPS, "Subaru")
    hyundai_ids = ids_for(TEMPLATE_GROUPS, "Hyundai")
    named_draft = make_draft(20000, {"mode": "named", "match": ["Subaru", "Hyundai"]},
                             {"split_evenly": True},
                             group_selection_reason="Notes said lead with Subaru and Hyundai.")
    state3 = apply_draft(named_draft, base_preset(all_groups()))
    selected3 = {g["id"] for g in state3["targeting_groups"] if g["include_in_plan"]}
    check("exactly Subaru + Hyundai selected", selected3 == subaru_ids | hyundai_ids, selected3)
    unresolved3 = state3.get("draft_unresolved", [])
    internal3 = state3.get("draft_unresolved_internal", [])
    check("a client-facing confirmation names the selection",
          any("Subaru and Hyundai" in u for u in unresolved3), unresolved3)
    check("no internal note falsely claims a silent select-all",
          not any("didn't say which" in n for n in internal3), internal3)
    check("Volvo/Genesis reported as available but not on the plan",
          any("Volvo" in n and "Genesis" in n for n in internal3), internal3)

    print("\n" + "=" * 78)
    print("Silent notes select everything, disclosed in unresolved_internal only")
    print("=" * 78)
    state4 = apply_draft(make_draft(20000, {"mode": "all"}, {"split_evenly": True}),
                         base_preset(all_groups()))
    unresolved4 = state4.get("draft_unresolved", [])
    internal4 = state4.get("draft_unresolved_internal", [])
    check("no client-facing claim about a specific selection",
          not any("Confirm that's what's being sold" in u for u in unresolved4), unresolved4)
    check("an internal note discloses the silent select-all",
          any("didn't say which" in n for n in internal4), internal4)

    print("\n" + "=" * 78)
    print("A rep's manual include_locked survives a redraft that names everything")
    print("=" * 78)
    locked_groups = all_groups()
    for g in locked_groups:
        if g["id"] in ids_for(TEMPLATE_GROUPS, "Genesis"):
            g["include_in_plan"], g["include_locked"] = False, True  # a rep deliberately excluded it
    state5 = apply_draft(make_draft(20000, {"mode": "all"}, {"split_evenly": True}),
                         base_preset(locked_groups))
    genesis_after = next(g for g in state5["targeting_groups"]
                         if g["id"] in ids_for(TEMPLATE_GROUPS, "Genesis"))
    check("Genesis stays excluded despite mode=all -- the rep's lock wins",
          not genesis_after["include_in_plan"], genesis_after)
    internal5 = state5.get("draft_unresolved_internal", [])
    check("kept-your-own-choice is disclosed", any("Kept your own choice" in n for n in internal5),
          internal5)
    other_selected = {g["id"] for g in state5["targeting_groups"] if g["include_in_plan"]}
    check("the other 3 makes (6 groups) ARE selected", len(other_selected) == 6, other_selected)

    print("\n" + "=" * 78)
    print("Redraft shallow-merge: an explicitly empty group_selection falls back to "
          "select-all, never select-none")
    print("=" * 78)
    # call_claude_redraft merges {**previous, **revised} -- simulate the
    # MERGED result a revision that only changed the budget would produce:
    # every other top-level key byte-identical, group_selection explicitly {}.
    previous = make_draft(20000, {"mode": "named", "match": ["Subaru"]}, {"split_evenly": True})
    revised_merged = {**previous, "total_budget": 25000, "group_selection": {}}
    state6 = apply_draft(revised_merged, base_preset(all_groups()))
    selected6 = {g["id"] for g in state6["targeting_groups"] if g["include_in_plan"]}
    check("an explicitly empty group_selection selects ALL groups, not none",
          selected6 == {g["id"] for g in TEMPLATE_GROUPS}, selected6)
    internal6 = state6.get("draft_unresolved_internal", [])
    check("disclosed as a silent select-all, same as an omitted field",
          any("didn't say which" in n for n in internal6), internal6)

    print("\n" + "=" * 78)
    print("Redraft shallow-merge: omitting the new keys keeps the previous draft's values")
    print("=" * 78)
    # call_claude_redraft itself does {**previous, **revised} -- exercise
    # THAT merge directly, then apply the result.
    revised_omits = {"total_budget": 25000}  # the model's own revision, group_selection untouched
    merged = {**previous, **revised_omits}
    check("the merge kept the previous group_selection", merged["group_selection"] == {"mode": "named", "match": ["Subaru"]}, merged.get("group_selection"))
    state7 = apply_draft(merged, base_preset(all_groups()))
    selected7 = {g["id"] for g in state7["targeting_groups"] if g["include_in_plan"]}
    check("still only Subaru selected (the merge preserved it)", selected7 == subaru_ids, selected7)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Drafting contributes budget/allocation/selection to existing groups and invents "
          "nothing when they're present; it still builds from scratch exactly as before when "
          "they're not. Pre-selection is always disclosed, routed correctly, and never "
          "overrides a rep's own locked choice.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
