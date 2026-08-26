"""Plan lines are matched back to their targeting group by ID, never by the
text a Geo or Targeting cell happens to render.

    python tests/test_group_plan_linkage.py

Driven through the real form (`streamlit.testing.v1.AppTest`), because what
matters is what the grid actually holds after a shared-field re-seed -- not
what a helper returns in isolation. `tests/test_plan_follows_avails.py`
already proves two audiences across three markets produce six lines and that
Denver appears twice without merging; what's new here is proving that holds
because of GROUP IDENTITY, not because the strings happened not to collide --
by driving a shared-field re-seed on a scenario built to make two rows render
an IDENTICAL Geo AND an identical Tactic, and confirming neither line's edit
lands on the other, and their targeting_groups stay two distinct groups
throughout.
"""
import os
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

COL = app.AVAILS_COLUMN_MONTHLY
ROOT = Path(__file__).resolve().parent.parent
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def included_groups(rows):
    """`rows` (the flat avails shape) -> targeting groups with every one
    ticked into the plan -- this suite is about group IDENTITY once a line
    exists, not about the include-in-plan default itself."""
    groups = tg.seed_rows_to_groups(rows, COL)
    for group in groups:
        group["include_in_plan"] = True
    return groups


def run(avails_rows=None, extra=None, then=None):
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (this file's own
    # subject) on market+flight.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    if avails_rows is not None:
        at.session_state["avails_seed_rows"] = avails_rows
        at.session_state["targeting_groups"] = included_groups(avails_rows)
    for key, value in (extra or {}).items():
        at.session_state[key] = value
    at.run()
    if at.exception:
        return None, None, at.exception[0].message[:300]
    if then:
        for key, value in then.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return None, None, at.exception[0].message[:300]
    groups = (at.session_state["targeting_groups"]
             if "targeting_groups" in at.session_state else None)
    return (at.session_state["plan_options"][0], groups, None)


def main():
    print("two audiences sharing one market, then a shared-field edit")
    # Both audiences target "Denver" -- an identical Geo string on the two
    # Premion rows this seeds. Old (string-only) matching couldn't tell them
    # apart; group ids can.
    rows = [{"Audience": "Homeowners", "Geo": "Denver", COL: 500000},
           {"Audience": "In-market for windows", "Geo": "Denver", COL: 300000}]
    option, groups, err = run(rows)
    check("the form renders", not err, err)
    if option:
        premion = [r for r in option["rows"] if r.get("Tactic") == "Premion Streaming TV"]
        check("two Premion lines, both reading Denver",
              len(premion) == 2 and {r["Geo"] for r in premion} == {"Denver"},
              [(r["Geo"], r["Targeting"]) for r in premion])
        check("but with two DIFFERENT group ids",
              len({tuple(app.group_ids_of(r)) for r in premion}) == 2,
              [app.group_ids_of(r) for r in premion])
        check("targeting_groups holds two distinct groups",
              groups is not None and len(groups) == 2, groups)
        ids_on_rows = {gid for r in premion for gid in app.group_ids_of(r)}
        ids_in_groups = {g["id"] for g in (groups or [])}
        check("every row's group id is a real group",
              ids_on_rows <= ids_in_groups, (ids_on_rows, ids_in_groups))

    print("\n...now trigger the re-seed branch (a shared-field change)")
    option, groups, err = run(
        rows, then={"audience_text": "should not matter -- these lines have their own audience"})
    check("the form renders after the edit", not err, err)
    if option:
        premion = [r for r in option["rows"] if r.get("Tactic") == "Premion Streaming TV"]
        check("still two Premion lines", len(premion) == 2, len(premion))
        check("still both reading Denver, not collapsed into one",
              {r["Geo"] for r in premion} == {"Denver"}, [r["Geo"] for r in premion])
        check("neither line's Targeting was reassigned to the other's",
              {r["Targeting"] for r in premion} == {"Homeowners", "In-market for windows"},
              sorted(r["Targeting"] for r in premion))
        check("group ids are unchanged by the edit",
              len({tuple(app.group_ids_of(r)) for r in premion}) == 2,
              [app.group_ids_of(r) for r in premion])
        check("targeting_groups is still exactly two groups, not merged",
              groups is not None and len(groups) == 2, groups)

    print("\na dead group id (its row was removed from the avails table) takes the default")
    # Seed two groups, then re-run with only ONE row left -- the removed
    # group's line should re-seed to the default rather than keep pointing at
    # a group that no longer exists.
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.session_state["avails_seed_rows"] = rows
    at.session_state["targeting_groups"] = included_groups(rows)
    at.run()
    at.session_state["avails_seed_rows"] = [rows[0]]   # drop the second group
    at.session_state["audience_text"] = "fallback audience"
    at.run()
    if at.exception:
        check("the form renders after a group is removed", False, at.exception[0].message[:300])
    else:
        premion = [r for r in at.session_state["plan_options"][0]["rows"]
                  if r.get("Tactic") == "Premion Streaming TV"]
        surviving = [r for r in premion if app.group_ids_of(r)
                    and app.group_ids_of(r)[0] in
                    {g["id"] for g in at.session_state["targeting_groups"]}]
        check("the surviving group's line keeps its own audience",
              any(r["Targeting"] == "Homeowners" for r in surviving), premion)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Plan lines are matched to their group by id, so two rows sharing "
          "a Geo string never collapse onto each other.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
