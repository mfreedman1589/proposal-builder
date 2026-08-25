"""The link between a targeting group's Plan checkbox and its media-plan row
must be symmetric -- found testing the real Plaza Motors avail: unchecking a
group already removes its line (reconcile_group_plan_lines, branch 1), but
deleting the LINE off the media plan grid never touched the group, so the
two disagreed and the deleted line could reappear (the very next
reconcile_group_plan_lines pass sees a still-selected group with no owning
row and adds one right back).

Fix: `unlink_deleted_group_rows` (app.py) -- the inverse of branch 1 -- and
wiring it into the media plan editor's own fold-back, right where
`reconcile_plan_rows` already detects a row-count change.

Part 1 (pure function, offline, precise): `unlink_deleted_group_rows`
directly, including the two edge cases that make it id-SET-based rather
than row-by-row -- a merged row carrying two group ids, and a duplicated
row where deleting one copy must not unselect a group still owning the
other.

Part 2 (through the real render, via the same st.data_editor monkeypatch
tests/test_avails_grid_row_deletion.py uses for the D2 grid -- AppTest can't
drive num_rows="dynamic" deletion directly): the full cycle in both
directions --
    check -> line appears -> delete line -> group unchecks (locked) ->
    the deleted line does NOT reappear on a further untouched rerun ->
    re-check -> line reappears
-- and the same starting from the checkbox (uncheck -> line disappears ->
re-check -> line reappears), confirming the pre-existing direction still
works unchanged.

    python tests/test_plan_line_group_symmetry.py
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

import streamlit as st  # noqa: E402
import pandas as pd  # noqa: E402

# Drive the media plan grid's own "-" the same way
# test_avails_grid_row_deletion.py drives the D2 grid's: monkeypatch
# st.data_editor to hand back a row set that's missing whichever row the
# test wants "deleted", keyed by the widget's own key prefix so the D2
# grid (a different data_editor call) is left alone.
_next_action = {"kind": None, "tactic": None, "done": True}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if not key.startswith("media_plan_editor") or _next_action["done"]:
        return data
    _next_action["done"] = True
    if _next_action["kind"] == "delete":
        return data[data["Tactic"] != _next_action["tactic"]].reset_index(drop=True)
    return data


st.data_editor = _fake_data_editor

import app                                     # noqa: E402
import targeting_groups as tg                  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


# --------------------------------------------------------------------------
# Part 1 -- the pure function
# --------------------------------------------------------------------------

def row(tactic, group_ids=None):
    r = {"Tactic": tactic, "Flight": "Sep 2026", "Geo": "St. Louis market",
        "Targeting": "AUTO Type Luxury", "Impressions": 1000.0, "CPM": 30.0,
        "Type": app.ROW_TYPE_RATE, "Cost": 30.0}
    if group_ids is not None:
        r["_group_ids"] = list(group_ids)
    return r


def group(gid, include_in_plan=True, include_locked=False):
    return {"id": gid, "terms": ["AUTO Type Luxury"], "op": None,
           "include_in_plan": include_in_plan, "include_locked": include_locked,
           "avails_monthly": 1000, "color": "#4C78A8", "color_locked": False}


def test_pure_function():
    print("unlink_deleted_group_rows -- pure function")

    print("\na deleted group-owned row unselects its group, locked")
    groups = [group("g1"), group("g2")]
    prev = [row("Premion Streaming TV", ["g1"]), row("Other line", ["g2"])]
    edited = [row("Other line", ["g2"])]   # g1's row deleted
    updated = app.unlink_deleted_group_rows(groups, prev, edited)
    g1 = next(g for g in updated if g["id"] == "g1")
    g2 = next(g for g in updated if g["id"] == "g2")
    check("g1 (its row was deleted) is unselected", g1["include_in_plan"] is False, g1)
    check("g1 is locked (a deliberate rep action, never silently re-ticked)",
          g1["include_locked"] is True, g1)
    check("g2 (untouched) stays selected", g2["include_in_plan"] is True, g2)

    print("\nno row-count change at all -- groups untouched (same object back)")
    groups2 = [group("g1")]
    prev2 = [row("Premion Streaming TV", ["g1"])]
    same = app.unlink_deleted_group_rows(groups2, prev2, [dict(prev2[0])])
    check("same object returned when nothing was removed", same is groups2, same)

    print("\na merged row carrying TWO group ids: deleting it unselects BOTH")
    groups3 = [group("g1"), group("g2")]
    prev3 = [row("Premion Streaming TV", ["g1", "g2"])]
    edited3 = []   # the merged row itself was deleted
    updated3 = app.unlink_deleted_group_rows(groups3, prev3, edited3)
    check("both merged groups unselected",
          all(g["include_in_plan"] is False for g in updated3), updated3)

    print("\na DUPLICATED group row: deleting one copy must not unselect a group "
          "still owning the surviving copy")
    groups4 = [group("g1")]
    prev4 = [row("Premion Streaming TV", ["g1"]), row("Premion Streaming TV", ["g1"])]
    edited4 = [row("Premion Streaming TV", ["g1"])]   # one copy deleted, one survives
    updated4 = app.unlink_deleted_group_rows(groups4, prev4, edited4)
    check("g1 stays selected -- it still owns the surviving duplicate",
          updated4 is groups4 or updated4[0]["include_in_plan"] is True, updated4)

    print("\na non-group row being deleted (no ids at all) leaves every group alone")
    groups5 = [group("g1")]
    prev5 = [row("Premion Streaming TV", ["g1"]), row("Custom Fee", None)]
    edited5 = [row("Premion Streaming TV", ["g1"])]   # the fee row deleted
    updated5 = app.unlink_deleted_group_rows(groups5, prev5, edited5)
    check("g1 untouched -- the deleted row wasn't its own",
          updated5 is groups5, updated5)


# --------------------------------------------------------------------------
# Part 2 -- through the real render
# --------------------------------------------------------------------------

def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["target_dmas"] = ["Washington, DC"]
    at.session_state["premion_streaming_tv"] = True
    return at


def ss(at, key, default=None):
    return at.session_state[key] if key in at.session_state else default


def premion_rows(at):
    rows = ss(at, "plan_options")[0]["rows"]
    return [r for r in rows if r.get("Tactic") == "Premion Streaming TV"]


def group_by_id(at, gid):
    return next(g for g in at.session_state["targeting_groups"] if g["id"] == gid)


def test_full_cycle():
    print("\nfull render cycle, both directions")
    real_group = tg.new_group(["AUTO Type Luxury"],
                              geo_def={"kind": "text", "label": "Zip list"},
                              avails_monthly=1700000, color="#4C78A8")
    gid = real_group["id"]

    at = new_app()
    at.session_state["targeting_groups"] = [real_group]
    at.run()
    check("no exception on initial render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("starts with no Premion row (include_in_plan defaults False)",
          len(premion_rows(at)) == 0, premion_rows(at))

    print("\n1. check -> line appears")
    at.session_state["_pending_group_include_all"] = True
    at.run()
    check("line appears after ticking the checkbox", len(premion_rows(at)) == 1, premion_rows(at))

    print("\n2. delete line -> group unchecks (locked), and it does not reappear")
    _next_action.update(kind="delete", tactic="Premion Streaming TV", done=False)
    at.run()
    check("the line is gone immediately after the delete", len(premion_rows(at)) == 0, premion_rows(at))
    g = group_by_id(at, gid)
    check("the group unselected itself", g["include_in_plan"] is False, g)
    check("the group is locked", g["include_locked"] is True, g)

    print("\n2b. a further untouched rerun does not resurrect the deleted line")
    at.run()
    check("still zero Premion rows after an untouched rerun -- not resurrected",
          len(premion_rows(at)) == 0, premion_rows(at))

    print("\n3. re-check -> line reappears")
    at.session_state["_pending_group_include_all"] = True
    at.run()
    check("line reappears after re-ticking", len(premion_rows(at)) == 1, premion_rows(at))

    print("\n4. the pre-existing direction still works: uncheck via the checkbox -> line disappears")
    at.session_state["_pending_group_include_clear"] = True
    at.run()
    check("line disappears after Clear plan lines", len(premion_rows(at)) == 0, premion_rows(at))

    print("\n5. re-check -> line reappears")
    at.session_state["_pending_group_include_all"] = True
    at.run()
    check("line reappears again", len(premion_rows(at)) == 1, premion_rows(at))


def main():
    test_pure_function()
    test_full_cycle()
    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The Plan checkbox and the media-plan row now agree in both directions: "
          "unchecking removes the row, and deleting the row unchecks the group -- "
          "so a deleted line can no longer reappear on its own.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
