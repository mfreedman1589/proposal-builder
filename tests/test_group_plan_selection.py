"""The include-in-plan mechanism: a targeting group contributes exactly one
Premion Streaming TV media-plan line if and only if its `include_in_plan`
flag is set, reconciled by `reconcile_group_plan_lines` (app.py). The avails
table is research by default -- importing or building a group never seeds a
plan line on its own; a rep (or a confirmed draft pre-selection) has to tick
"Plan".

    python tests/test_group_plan_selection.py

Driven against the real Annapolis Cars document (RFPID-253813, via
group_scenario_fixtures.py) -- 4 auto makes x 2 radii (10mi/5mi) = 8 real
groups, exactly the "several audiences, some sold, some shown as
opportunity" shape the three required scenarios map onto:

  1. one dealership sold, the other three shown as opportunity only
  2. five (here: four) dealerships, budget split evenly across them
  3. five (here: four) dealerships x two tiers, only one tier sold

Two properties get their own first-class checks, not just coverage-by-
implication from the scenarios: the D2 fold-back's survival of
include_in_plan/include_locked across an untouched rerun and an unrelated
cell edit (the single highest-risk mistake this whole feature can make --
tg.new_group silently losing the field on every render), and the
reconciler's idempotency (run it twice with nothing changed, assert nothing
moves) -- the property the whole order-of-operations design rests on.
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


def run_form(state):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Selection Suite"
    for key, value in state.items():
        at.session_state[key] = value
    at.run()
    return at


def ss(at, key, default=None):
    """AppTest's session_state has no `.get` -- the safe substitute used
    throughout this suite (same helper tests/test_avails_pdf_wiring.py
    defines) instead of repeating the `in` check."""
    return at.session_state[key] if key in at.session_state else default


SCN = gsf.build_annapolis()
for _g in SCN["groups"]:
    _g.pop("_flight_label", None)
# Built ONCE: apply_avails_import mints a fresh uuid per group on every
# parse, so calling build_annapolis() again per scenario would hand back a
# structurally-identical but ID-DIFFERENT set of groups, and any id
# captured from an earlier call (by_make, below) would match nothing in a
# later one. Every scenario below re-flags include_in_plan on a COPY of
# this same template instead.
TEMPLATE_GROUPS = SCN["groups"]


def annapolis_groups(selected_ids=()):
    """A copy of Annapolis's real 8 groups, every one's `include_in_plan`
    set from `selected_ids` -- default is none selected, matching what a
    fresh import actually produces now. Same group ids every call."""
    selected = set(selected_ids)
    groups = []
    for g in TEMPLATE_GROUPS:
        g = dict(g)
        g["include_in_plan"] = g["id"] in selected
        g["include_locked"] = False
        groups.append(g)
    return SCN, groups


def base_state(scn, groups, **extra):
    state = {
        "include_avails_template": True,
        "targeting_groups": groups,
        "avails_basis": app.AVAILS_BASIS_FLIGHT,
        "flight_start": scn["flight_start"], "flight_end": scn["flight_end"],
        "active_months": scn["active_months"],
        "client_name": scn["client_name"],
        "proposal_title": "CTV/OTT Strategy",
        "vertical_choice": scn["vertical_choice"],
        "agency_gross_up": scn["agency_gross_up"],
        "premion_streaming_tv": True,
    }
    state.update(extra)
    return state


def by_make(groups, make, miles=None):
    out = []
    for g in groups:
        if not any(make.lower() in t.lower() for t in g["terms"]):
            continue
        if miles is not None and g["geo_def"].get("miles") != miles:
            continue
        out.append(g)
    return out


def premion_rows(at):
    return [r for r in at.session_state["plan_options"][0]["rows"]
           if r.get("Tactic") == "Premion Streaming TV"]


def main():
    print("=" * 78)
    print("Fresh import: nothing selected, nothing on the plan")
    print("=" * 78)
    scn, groups = annapolis_groups()
    at = run_form(base_state(scn, groups))
    check("form renders", not at.exception, at.exception[0].message[:400] if at.exception else "")
    if at.exception:
        print(f"\n{len(failures)} FAILED: {failures}")
        return 1
    real = [g for g in at.session_state["targeting_groups"] if g["terms"]]
    check("all 8 real groups present", len(real) == 8, len(real))
    check("none ticked into the plan", not any(g["include_in_plan"] for g in real), real)
    check("zero Premion Streaming TV rows", len(premion_rows(at)) == 0, premion_rows(at))

    print("\n" + "=" * 78)
    print("Ticking N groups produces exactly N rows, one _group_ids each")
    print("=" * 78)
    subaru = by_make(groups, "Subaru")
    scn2, groups2 = annapolis_groups({g["id"] for g in subaru})
    at2 = run_form(base_state(scn2, groups2))
    rows2 = premion_rows(at2)
    check("2 Subaru groups ticked -> 2 Premion rows", len(rows2) == 2, rows2)
    check("each row carries exactly one group id, and they're distinct",
          all(len(app.group_ids_of(r)) == 1 for r in rows2)
          and len({app.group_ids_of(r)[0] for r in rows2}) == 2, rows2)

    print("\n" + "=" * 78)
    print("SCENARIO 1 -- one dealership sold, the other three shown as opportunity")
    print("=" * 78)
    scn3, groups3 = annapolis_groups({g["id"] for g in subaru})
    at3 = run_form(base_state(scn3, groups3))
    rows3 = premion_rows(at3)
    check("exactly 2 plan rows (Subaru's two radii)", len(rows3) == 2, rows3)
    check("both rows target Subaru", all("Subaru" in r["Targeting"] for r in rows3), rows3)
    other_ids = {g["id"] for g in groups3 if "Subaru" not in tg.audience_label(g)}
    row_ids = {gid for r in rows3 for gid in app.group_ids_of(r)}
    check("no other make's group id appears in any row", not (other_ids & row_ids), row_ids)
    check("the other 3 makes (6 groups) still carry real avails in the table",
          all(g["avails_monthly"] > 0 for g in groups3 if g["id"] in other_ids), None)

    print("\n" + "=" * 78)
    print("SCENARIO 2 -- four dealerships, budget split evenly across them")
    print("=" * 78)
    ten_mi = by_make(groups, "", miles=10.0)
    check("setup: exactly 4 ten-mile groups (one per make)", len(ten_mi) == 4, ten_mi)
    scn4, groups4 = annapolis_groups({g["id"] for g in ten_mi})
    at4 = run_form(base_state(scn4, groups4))
    rows4 = premion_rows(at4)
    check("exactly 4 plan rows", len(rows4) == 4, rows4)
    # Apply a split-evenly draft intent through the real reconciler, the same
    # way a "split evenly across the audiences in the avail" clarification
    # would: queue a realloc and rerun.
    at4.session_state["draft_plan_intent"] = {
        "source": "draft", "round": 1, "flight_label": scn4["flight_label"],
        "default_targeting": "", "n_months": scn4["n_months"],
        "options": [{"name": "Option A", "total_budget": 40000, "breakout": "Full Flight",
                     "group_allocation": {"split_evenly": True}, "group_cpm": None,
                     "other_lines": [], "selection": {"mode": "all"}}],
    }
    at4.session_state["_pending_group_realloc"] = True
    at4.run()
    check("no exception applying the allocation", not at4.exception,
          at4.exception[0].message[:400] if at4.exception else "")
    rows4b = premion_rows(at4)
    total = sum(app._num(r["Cost"]) for r in rows4b)
    check("full-flight total ties to the stated $40,000 budget (not a recomputed sum)",
          abs(total - 40000) < 1, total)
    costs = {round(app._num(r["Cost"]), 2) for r in rows4b}
    check("all four rows cost the same (split evenly)", len(costs) == 1, rows4b)

    print("\n" + "=" * 78)
    print("SCENARIO 3 -- four dealerships x two tiers, only the 10mi tier sold")
    print("=" * 78)
    scn5, groups5 = annapolis_groups({g["id"] for g in ten_mi})
    at5 = run_form(base_state(scn5, groups5))
    rows5 = premion_rows(at5)
    check("exactly 4 plan rows (the 10mi tier only)", len(rows5) == 4, rows5)
    five_mi_ids = {g["id"] for g in groups5 if g["geo_def"].get("miles") == 5.0}
    row_ids5 = {gid for r in rows5 for gid in app.group_ids_of(r)}
    check("no 5mi group id appears in any row", not (five_mi_ids & row_ids5), row_ids5)
    check("the 5mi tier is still in the avails table with real avails",
          all(g["avails_monthly"] > 0 for g in groups5 if g["id"] in five_mi_ids), None)

    print("\n" + "=" * 78)
    print("Unchecking a CLEAN group removes its line silently")
    print("=" * 78)
    scn6, groups6 = annapolis_groups({g["id"] for g in subaru})
    at6 = run_form(base_state(scn6, groups6))
    check("setup: 2 rows before unchecking", len(premion_rows(at6)) == 2, None)
    ungroups = [dict(g, include_in_plan=False) for g in at6.session_state["targeting_groups"]]
    at6.session_state["targeting_groups"] = ungroups
    at6.run()
    check("no exception", not at6.exception, at6.exception[0].message[:400] if at6.exception else "")
    check("zero rows after unchecking, no confirmation needed",
          len(premion_rows(at6)) == 0, premion_rows(at6))
    check("no pending confirmation queued", not ss(at6, "_pending_include_removal"), None)

    print("\n" + "=" * 78)
    print("Unchecking a DIRTY group interrupts -- Clear plan lines -> Keep -> Clear -> Remove")
    print("=" * 78)
    scn7, groups7 = annapolis_groups({g["id"] for g in subaru})
    at7 = run_form(base_state(scn7, groups7))
    rows7 = at7.session_state["plan_options"][0]["rows"]
    rows7[0]["Cost"] = 99999.0
    # driver=DRIVER_COST, not just dirty=True -- the grid's own bidirectional
    # recompute derives whichever of Cost/Impressions ISN'T the driver, so
    # editing Cost without also saying Cost is now the driver would have
    # this recomputed right back from the still-zero Impressions.
    at7.session_state["plan_options"][0]["driver"][0] = app.DRIVER_COST
    at7.session_state["plan_options"][0]["dirty"][0] = True
    at7.session_state["plan_options"][0]["rows"] = rows7
    at7.run()
    check("no exception after the hand-edit", not at7.exception,
          at7.exception[0].message[:400] if at7.exception else "")
    # "Clear plan lines" is a real button click -- buttons can't be driven
    # by writing session_state (Streamlit raises), same reason every other
    # button in this suite goes through at.button/.click().
    clear_btn = next(b for b in at7.button if b.label == "Clear plan lines")
    clear_btn.click().run()
    check("the edited group is still included -- the clean one was cleared, the dirty one wasn't",
          any(g["include_in_plan"] for g in at7.session_state["targeting_groups"]
              if "Subaru" in tg.audience_label(g)), at7.session_state["targeting_groups"])
    blocked = ss(at7, "_pending_include_removal")
    check("a confirmation is queued naming the edited line", bool(blocked), blocked)
    check("the edited row (Cost 99999) is still on the grid, untouched",
          any(abs(app._num(r["Cost"]) - 99999.0) < 1 for r in premion_rows(at7)), premion_rows(at7))

    # Cancel ("Keep them on the plan"): nothing changes.
    keep_btn = next(b for b in at7.button if b.label == "Keep them on the plan")
    keep_btn.click().run()
    check("Keep: the pending queue is gone, the group is still included, the row is still there",
          not ss(at7, "_pending_include_removal")
          and any(g["include_in_plan"] for g in at7.session_state["targeting_groups"]
                  if "Subaru" in tg.audience_label(g))
          and any(abs(app._num(r["Cost"]) - 99999.0) < 1 for r in premion_rows(at7)),
          (ss(at7, "_pending_include_removal"), premion_rows(at7)))

    # Clear again, then confirm ("Remove them and discard those numbers").
    clear_btn2 = next(b for b in at7.button if b.label == "Clear plan lines")
    clear_btn2.click().run()
    remove_btns = [b for b in at7.button if b.label == "Remove them and discard those numbers"]
    check("a Remove button exists to click", bool(remove_btns), None)
    if remove_btns:
        remove_btns[0].click().run()
        check("no exception confirming the removal", not at7.exception,
              at7.exception[0].message[:400] if at7.exception else "")
        check("the group is now excluded and locked",
              all((not g["include_in_plan"]) and g["include_locked"]
                  for g in at7.session_state["targeting_groups"] if "Subaru" in tg.audience_label(g)),
              at7.session_state["targeting_groups"])
        check("the edited row is gone from the plan", len(premion_rows(at7)) == 0, premion_rows(at7))

    print("\n" + "=" * 78)
    print("Non-group lines (a hand-added fee) survive Add-all and Clear untouched")
    print("=" * 78)
    scn8, groups8 = annapolis_groups()
    at8 = run_form(base_state(scn8, groups8))
    opt = at8.session_state["plan_options"][0]
    opt["rows"].append({"Tactic": "Production Fee", "Flight": scn8["flight_label"], "Geo": "n/a",
                        "Targeting": "n/a", "Impressions": 0.0, "CPM": 0.0,
                        "Type": app.ROW_TYPE_FLAT_FEE, "Cost": 850.0})
    opt["dirty"].append(True)
    opt["driver"].append(app.DRIVER_COST)
    at8.session_state["plan_options"] = [opt]
    at8.run()
    add_all_btn = next(b for b in at8.button if b.label == "Add all to plan")
    add_all_btn.click().run()
    check("Add all: the fee line survives", any(r.get("Tactic") == "Production Fee"
          for r in at8.session_state["plan_options"][0]["rows"]), None)
    check("Add all: every group is now included",
          all(g["include_in_plan"] for g in at8.session_state["targeting_groups"] if g["terms"]), None)
    check("Add all: 8 Premion rows now exist", len(premion_rows(at8)) == 8, premion_rows(at8))
    clear_btn = next(b for b in at8.button if b.label == "Clear plan lines")
    clear_btn.click().run()
    check("Clear: the fee line still survives", any(r.get("Tactic") == "Production Fee"
          for r in at8.session_state["plan_options"][0]["rows"]), None)
    check("Clear: zero Premion rows remain (none were hand-edited)",
          len(premion_rows(at8)) == 0, premion_rows(at8))

    print("\n" + "=" * 78)
    print("Fold-back survival: tick a row, untouched rerun; tick + edit an unrelated cell")
    print("=" * 78)
    scn9, groups9 = annapolis_groups({subaru[0]["id"]})
    at9 = run_form(base_state(scn9, groups9))
    ticked = next(g for g in at9.session_state["targeting_groups"] if g["id"] == subaru[0]["id"])
    check("setup: the ticked group really is included", ticked["include_in_plan"], ticked)
    at9.run()  # a genuinely untouched rerun -- nothing in state changed
    still = next(g for g in at9.session_state["targeting_groups"] if g["id"] == subaru[0]["id"])
    check("still ticked after an untouched rerun", still["include_in_plan"], still)
    check("still has exactly 1 Premion row after the untouched rerun",
          len(premion_rows(at9)) == 1, premion_rows(at9))

    scn10, groups10 = annapolis_groups({subaru[0]["id"]})
    at10 = run_form(base_state(scn10, groups10))
    edited = [dict(g) for g in at10.session_state["targeting_groups"]]
    for g in edited:
        if g["id"] == subaru[1]["id"]:  # an UNRELATED group's name, not the ticked one
            g["name"] = "Subaru 5mi (renamed)"
    at10.session_state["targeting_groups"] = edited
    at10.run()
    check("no exception after an unrelated cell edit", not at10.exception,
          at10.exception[0].message[:400] if at10.exception else "")
    still2 = next(g for g in at10.session_state["targeting_groups"] if g["id"] == subaru[0]["id"])
    check("inclusion survives an edit to a DIFFERENT row entirely", still2["include_in_plan"], still2)
    check("still has exactly 1 Premion row", len(premion_rows(at10)) == 1, premion_rows(at10))

    print("\n" + "=" * 78)
    print("The reconciler is idempotent: two consecutive runs, nothing changes")
    print("=" * 78)
    scn11, groups11 = annapolis_groups({g["id"] for g in ten_mi})
    at11 = run_form(base_state(scn11, groups11))
    before = at11.session_state["plan_options"]
    before_snapshot = [(o["name"], list(o["rows"]), list(o["dirty"]), o["version"]) for o in before]
    at11.run()  # nothing in state changed -- a genuine no-op rerun
    after = at11.session_state["plan_options"]
    after_snapshot = [(o["name"], list(o["rows"]), list(o["dirty"]), o["version"]) for o in after]
    check("plan_options unchanged (rows, dirty, version) across a no-op rerun",
          before_snapshot == after_snapshot, (before_snapshot, after_snapshot))

    print("\n" + "=" * 78)
    print("Multi-option drafting: each option's rows stay scoped to its OWN group_selection,")
    print("not every globally-included group -- the LiveWell shape (one option selling one of")
    print("many same-audience locations, another selling all of them). Checked after a real")
    print("rerun, not the moment plan_options is first written -- the bug this guards only")
    print("showed up on the NEXT render, when reconcile_group_plan_lines' own row-seeding step")
    print("(never apply_draft_to_form's) is what adds a row.")
    print("=" * 78)
    # Real LiveWell incident: "Alexandria Only" (one group, named) and "All
    # Locations" (mode=all) came out of one draft; on the very next rerun,
    # before a rep ever saw the plan, "Alexandria Only" picked up the other
    # six groups too, each at a real, visible $0. groups[0] stands in for
    # Alexandria; groups[1..6] for the other six locations "All Locations"
    # alone should own; groups[7] stands in for a group NEITHER option's
    # stored intent ever resolved -- ticked onto the shared avails table
    # after the draft ran (or simply left unaddressed by it) -- which per
    # this fix's own fallback rule should still reach every option, since
    # nobody's intent claims it and there's no per-option information to
    # exclude it with.
    solo_group, rest_groups, new_group = TEMPLATE_GROUPS[0], TEMPLATE_GROUPS[1:7], TEMPLATE_GROUPS[7]
    all_scoped_groups = TEMPLATE_GROUPS[0:7]  # what "All Locations" itself resolved -- not the 8th
    scn12, groups12 = annapolis_groups({g["id"] for g in TEMPLATE_GROUPS})  # everything ticked

    def _stub_row(group):
        return {"Tactic": "Premion Streaming TV", "Targeting": tg.audience_label(group),
                "Geo": tg.geo_label(group), "Cost": 0.0, "Impressions": 0.0, "CPM": 32.0,
                "Type": "Rate", "_group_ids": [group["id"]]}

    solo_option = app.new_plan_option("Alexandria Only", [_stub_row(solo_group)])
    all_option = app.new_plan_option(
        "All Locations", [_stub_row(g) for g in all_scoped_groups])
    draft_plan_intent = {
        "n_months": 1,
        "options": [
            {"name": "Alexandria Only", "total_budget": 6000, "breakout": "monthly",
             "group_allocation": {"split_evenly": True}, "group_cpm": None, "other_lines": [],
             "selection": {"mode": "named", "ids": [solo_group["id"]], "reason": "",
                          "matched_ids": [solo_group["id"]], "unmatched": []}},
            {"name": "All Locations", "total_budget": 28000, "breakout": "monthly",
             "group_allocation": {"split_evenly": True}, "group_cpm": None, "other_lines": [],
             "selection": {"mode": "all", "reason": "",
                          "matched_ids": [g["id"] for g in all_scoped_groups], "unmatched": []}},
        ],
    }
    at12 = run_form(base_state(
        scn12, groups12, plan_options=[solo_option, all_option],
        draft_plan_intent=draft_plan_intent))
    check("no exception", not at12.exception,
          at12.exception[0].message[:400] if at12.exception else "")
    opts_after = {o["name"]: o for o in at12.session_state["plan_options"]}
    solo_ids_after = {gid for r in opts_after["Alexandria Only"]["rows"] for gid in app.group_ids_of(r)}
    all_ids_after = {gid for r in opts_after["All Locations"]["rows"] for gid in app.group_ids_of(r)}
    check("Alexandria Only: still just its own group plus the new/unclaimed one -- "
          "NOT the other six 'All Locations' owns",
          solo_ids_after == {solo_group["id"], new_group["id"]}, solo_ids_after)
    check("the other six locations are NOT phantom rows on Alexandria Only",
          not (solo_ids_after & {g["id"] for g in rest_groups}), solo_ids_after)
    check("All Locations: its own seven plus the new/unclaimed one -- eight total",
          all_ids_after == {g["id"] for g in all_scoped_groups} | {new_group["id"]}, all_ids_after)
    check("a phantom row, if any had leaked through, would show as a real $0 on the deck "
          "table (compute_plan_totals only skips an empty Tactic, never a zero-cost row) -- "
          "so zero phantom rows is what actually keeps the deck sendable",
          len(opts_after["Alexandria Only"]["rows"]) == 2, opts_after["Alexandria Only"]["rows"])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The avails table is research; a group's include_in_plan flag alone decides its "
          "Premion Streaming TV line, reconciled the same way whether ticked one at a time, "
          "via Add all / Clear, or left untouched across any number of reruns. A multi-option "
          "draft's own per-option selection is respected too, not flattened into one shared "
          "row set.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
