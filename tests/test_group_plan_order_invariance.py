"""The acceptance test for decoupling the avails table from the media plan:
once real targeting groups exist, order stops mattering. Two importable
batches of Annapolis Cars groups (RFPID-253813's real 10mi and 5mi radius
tiers, treated as two separate imports) and one draft selection
("match": ["Subaru"], split evenly) reach the IDENTICAL end state whether
the whole avails table exists before the draft runs, or only part of it
does and the rest arrives afterward -- because `apply_draft_plan_intent_to_
new_groups` re-applies the SAME stored, raw selection criteria to whatever
groups show up later, rather than the draft baking in a fixed set of ids.

Scope, deliberately narrower than "any order at all" -- confirmed with the
user: group selection is only ever expressed once real groups exist at
draft time (the prompt only lists them, and therefore only asks about them,
when there's something to select from). A draft run before ANY avails work
exists takes the old "invent Premion lines from the notes" path, and
nothing retroactively reshapes those invented lines once a real avails
table shows up later -- see `apply_draft_plan_intent_to_new_groups`'s own
docstring and CLAUDE.md's targeting-groups section for why that's an
accepted, narrow gap rather than something this suite tries to converge.
What IS tested here is every order where real groups exist by the time
selection is expressed:

  A. both batches imported, then draft
  B. batch 1 imported, draft (selects what exists), batch 2 imported after
     (the stored intent picks up the newly-arrived match)
  C. either of the above, then a CLARIFY round changes the allocation --
     re-prices the SAME group ids, never regenerates

Plus the property the whole design rests on: `reconcile_group_plan_lines`
is idempotent -- run it twice with nothing changed and nothing moves.

Every state transformation (import, draft, redraft) runs OFFLINE through a
plain-dict `_StubSt`, the same pattern tests/test_draft_regression.py's
`apply_draft()` and group_scenario_fixtures.py's `_import_groups()` both
use -- AppTest's own `session_state` doesn't support the plain dict
operations (`.get`, wholesale copy) `apply_draft_to_form` needs, so a real
AppTest is used only once per end state, to get the REAL reconciled rows
out of the real render pipeline.

    python tests/test_group_plan_order_invariance.py
"""
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
    """apply_draft_to_form and apply_draft_plan_intent_to_new_groups only
    touch st.session_state -- a plain dict stands in, same device
    tests/test_draft_regression.py and group_scenario_fixtures.py both use,
    so this can run entirely offline."""
    def __init__(self, state):
        self.session_state = state
        self.secrets = {}


def apply_offline(state, fn, *args):
    """Run `fn(*args)` with app.st swapped for a _StubSt over `state` (a
    plain dict), returning the same dict with whatever `fn` wrote merged in."""
    real, app.st = app.st, _StubSt(state)
    try:
        fn(*args)
    finally:
        app.st = real
    return state


def run_form(state):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Order Invariance Suite"
    for key, value in state.items():
        at.session_state[key] = value
    at.run()
    return at


SCN = gsf.build_annapolis()
for _g in SCN["groups"]:
    _g.pop("_flight_label", None)
TEMPLATE_GROUPS = SCN["groups"]
TEN_MI = [dict(g, include_in_plan=False, include_locked=False)
         for g in TEMPLATE_GROUPS if g["geo_def"].get("miles") == 10.0]
FIVE_MI = [dict(g, include_in_plan=False, include_locked=False)
          for g in TEMPLATE_GROUPS if g["geo_def"].get("miles") == 5.0]
check("setup: 4 groups per radius tier", len(TEN_MI) == 4 and len(FIVE_MI) == 4,
      (len(TEN_MI), len(FIVE_MI)))


def base_state(groups, **extra):
    state = {
        "include_avails_template": True,
        "targeting_groups": [dict(g) for g in groups],
        "avails_basis": app.AVAILS_BASIS_FLIGHT,
        "flight_start": SCN["flight_start"], "flight_end": SCN["flight_end"],
        "active_months": SCN["active_months"],
        "client_name": SCN["client_name"],
        "proposal_title": "CTV/OTT Strategy",
        "vertical_choice": SCN["vertical_choice"],
        "agency_involved": SCN["agency_involved"],
        "premion_streaming_tv": True,
    }
    state.update(extra)
    return state


def premion_rows(at):
    return [r for r in at.session_state["plan_options"][0]["rows"]
           if r.get("Tactic") == "Premion Streaming TV"]


def end_state(at):
    """(selected audience labels, {label: (cost, impressions)}) -- compared
    by AUDIENCE LABEL rather than raw group id, since sequence B's batch-2
    groups are injected as fresh dicts (a real second import mints its own
    ids; labels are what has to agree)."""
    groups_by_id = {g["id"]: g for g in at.session_state["targeting_groups"]}
    selected = sorted(tg.audience_label(g) for g in groups_by_id.values() if g["include_in_plan"])
    amounts = {}
    for r in premion_rows(at):
        gid = app.group_ids_of(r)[0] if app.group_ids_of(r) else None
        label = tg.audience_label(groups_by_id[gid]) if gid in groups_by_id else r.get("Targeting")
        amounts.setdefault(label, []).append(
            (round(app._num(r["Cost"]), 2), round(app._num(r["Impressions"]), 2)))
    return selected, amounts


def subaru_draft(budget, allocation=None):
    def _d(value):
        return value.isoformat() if hasattr(value, "isoformat") else str(value)
    return {
        "client_name": SCN["client_name"], "vertical": "auto", "market": "DC",
        "agency_involved": SCN["agency_involved"],
        "flight_start": _d(SCN["flight_start"]), "flight_end": _d(SCN["flight_end"]),
        "total_budget": budget, "breakout": "full_flight",
        "media_plan_lines": [],
        "group_selection": {"mode": "named", "match": ["Subaru"]},
        "group_allocation": allocation or {"split_evenly": True},
        "group_selection_reason": "Notes said lead with Subaru.",
        "audiences": [], "attribution": [], "sports": [],
        "unresolved": [], "unresolved_internal": [],
    }


def main():
    print("=" * 78)
    print("A. Both radius tiers imported, THEN draft selects Subaru (both tiers)")
    print("=" * 78)
    state_a = base_state(TEN_MI + FIVE_MI)
    apply_offline(state_a, app.apply_draft_to_form, subaru_draft(20000))
    at_a = run_form(state_a)
    check("no exception", not at_a.exception, at_a.exception[0].message[:400] if at_a.exception else "")
    selected_a, amounts_a = end_state(at_a)
    check("both Subaru tiers selected", selected_a == ["AUTO Make Subaru", "AUTO Make Subaru"], selected_a)
    # FLOW_REWORK_PLAN.md Phase 3: Subaru's 5mi/10mi tiers are ONE entity
    # (avails_pdf_import.infer_entities' R1 rule -- same audience, same
    # radius origin, differing radius). "Split evenly" divides by ENTITY
    # count, and there's only one Subaru entity here, so the WHOLE $20,000
    # lands on its representative row; the sibling stays unpriced ($0) --
    # never $10,000 each, which would be two separate entities' worth.
    check("exactly 2 Premion rows -- one entity, full $20,000 on the "
          "representative, $0 on the sibling (never split across two rows)",
          len(premion_rows(at_a)) == 2
          and {c for vals in amounts_a.values() for c, _i in vals} == {20000.0, 0.0}, amounts_a)

    print("\n" + "=" * 78)
    print("B. Batch 1 (10mi) imported, draft selects Subaru, THEN batch 2 (5mi) arrives")
    print("=" * 78)
    state_b = base_state(TEN_MI)
    apply_offline(state_b, app.apply_draft_to_form, subaru_draft(20000))
    at_b1 = run_form(dict(state_b))
    check("no exception after the draft (batch 1 only)", not at_b1.exception,
          at_b1.exception[0].message[:400] if at_b1.exception else "")
    selected_b1, _ = end_state(at_b1)
    check("only the 10mi Subaru group is selected yet (5mi doesn't exist)",
          selected_b1 == ["AUTO Make Subaru"], selected_b1)
    check("draft_plan_intent was stored for later", bool(state_b.get("draft_plan_intent")), None)

    # Batch 2 "arrives" -- the exact function _finish_avails_import calls,
    # applied offline to the fresh 5mi groups (a real second import mints
    # its own ids for them, same as this does).
    def _second_import():
        new_groups = app.apply_draft_plan_intent_to_new_groups([dict(g) for g in FIVE_MI])
        state_b["targeting_groups"] = state_b["targeting_groups"] + new_groups
    apply_offline(state_b, _second_import)
    at_b2 = run_form(dict(state_b))
    check("no exception after batch 2 arrives", not at_b2.exception,
          at_b2.exception[0].message[:400] if at_b2.exception else "")
    selected_b2, amounts_b2 = end_state(at_b2)
    check("the stored intent picked up the newly-arrived Subaru 5mi group",
          selected_b2 == ["AUTO Make Subaru", "AUTO Make Subaru"], selected_b2)

    print("\n" + "=" * 78)
    print("A and B converge on the identical end state")
    print("=" * 78)
    check("same selected audiences", selected_a == selected_b2, (selected_a, selected_b2))
    check("same amounts per line (cost, impressions)",
          sorted(v for vals in amounts_a.values() for v in vals)
          == sorted(v for vals in amounts_b2.values() for v in vals),
          (amounts_a, amounts_b2))

    print("\n" + "=" * 78)
    print("C. Clarify-after-either: a redraft changing the allocation re-prices the "
          "SAME group ids, never regenerates them")
    print("=" * 78)
    ids_before = {gid for r in premion_rows(at_a) for gid in app.group_ids_of(r)}
    apply_offline(state_a, app.apply_draft_to_form, subaru_draft(20000, {"percent_of_total": 60}))
    at_a2 = run_form(dict(state_a))
    check("no exception after the redraft", not at_a2.exception,
          at_a2.exception[0].message[:400] if at_a2.exception else "")
    ids_after = {gid for r in premion_rows(at_a2) for gid in app.group_ids_of(r)}
    check("the SAME two group ids own the rows after the redraft -- not regenerated",
          ids_before == ids_after and len(ids_after) == 2, (ids_before, ids_after))
    new_amounts = {gid: round(app._num(r["Cost"]), 2) for r in premion_rows(at_a2)
                  for gid in app.group_ids_of(r)}
    # Same one-entity shape as scenario A: 60% of $20,000 is $12,000 for the
    # WHOLE Subaru entity, all on its representative row -- not $12,000
    # each, which would double-count one entity as two.
    check("the allocation actually changed (60% of $20,000 = $12,000 total for the "
          "one Subaru entity, on its representative row; $0 on the sibling)",
          sorted(new_amounts.values()) == [0.0, 12000.0], new_amounts)

    print("\n" + "=" * 78)
    print("The reconciler is idempotent -- run it twice with nothing changed, nothing moves")
    print("=" * 78)
    before = [(o["name"], list(o["rows"]), list(o["dirty"]), o["version"])
             for o in at_a2.session_state["plan_options"]]
    at_a2.run()
    after = [(o["name"], list(o["rows"]), list(o["dirty"]), o["version"])
            for o in at_a2.session_state["plan_options"]]
    check("plan_options byte-identical across a genuine no-op rerun", before == after, (before, after))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Order stops mattering once real groups exist: importing both tiers up front, or "
          "drafting against a partial table and letting a later import complete it, reach the "
          "identical selected-groups/plan-lines/amounts end state -- and a redraft re-prices "
          "the same lines rather than regenerating them.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
