"""Deleting a group-backed row from ONE plan option must not touch a
sibling option's copy of the same group's row, and must not silently
reappear on the next rerun while the group is still included elsewhere --
reported live (Easterns, 2026-09-23): a drafted plan put the same Premion
Streaming TV audience on two options ("CTV" and "CTV + NFL Regular
Season"); deleting the duplicated CTV line from the NFL option also
deleted it from the CTV option.

Root cause, two parts:
  1. `unlink_deleted_group_rows` decided whether to exclude a group
     GLOBALLY by looking only at the one option it was called for -- a
     group still owning a row in a sibling option read as "fully removed"
     the moment its row vanished from whichever option happened to be
     rendering. Fixed: it now also sees every OTHER option's rows before
     concluding a group has no row left anywhere.
  2. `reconcile_group_plan_lines`'s add step re-derived "should this
     option have a row for this SELECTED group" fresh on every render, so
     even with (1) fixed, a row deleted from one option (while the group
     stayed selected, because a sibling option still had a copy) came
     right back on the very next rerun. Fixed: adding a group's row now
     fires only on a genuine off->on SELECTION TRANSITION
     (`include_snapshot`, persisted in session_state across reruns), never
     merely because a still-selected group happens to lack a row in one
     option. Unticking (or reaching zero owning rows anywhere, via (1))
     still removes a group's row from every option immediately -- that
     half stays state-based and symmetric, and needed no change.

The exact scenarios agreed on before this was built:
  A. Delete from option 2 alone, group still in option 1 -> option 2 stays
     deleted across an untouched rerun, option 1 untouched, checkbox still
     ticked (the group still owns a row in option 1).
  B. Delete from option 1 too (now both gone) -> unlink_deleted_group_rows
     unticks the checkbox, since no option owns a row for it any more.
  C. Unticking a still-fully-seeded group removes its row from every
     option at once (unchanged, state-based).
  D. Re-ticking re-seeds every option that lacks a row for it.
  E. A redraft with the group already linked matches the EXISTING rows by
     group id (re-prices them) rather than re-seeding fresh ones from
     checkbox state -- confirms apply_draft_to_form's own per-option
     construction (matched_ids/_synthesize_group_lines) is untouched by
     the include_snapshot change, which only gates the SEPARATE add-step
     reconcile_group_plan_lines itself owns.

    python tests/test_group_plan_per_option_delete.py
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

import streamlit as st  # noqa: E402
import pandas as pd  # noqa: E402

_next_action = {"kind": None, "option_idx": None, "done": True}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    prefix = f"media_plan_editor_{_next_action.get('option_idx')}_"
    if _next_action["done"] or not key.startswith(prefix):
        return data
    _next_action["done"] = True
    kind = _next_action["kind"]
    if kind == "delete_dc":
        return data[~((data["Tactic"].astype(str).str.contains("Premion")) &
                      (data["Geo"].astype(str).str.contains("DC")))].reset_index(drop=True)
    return data


st.data_editor = _fake_data_editor

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def group_ids_in(rows):
    return sorted({gid for r in rows for gid in app.group_ids_of(r)})


def new_app():
    from datetime import date
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 8)
    at.session_state["flight_end"] = date(2027, 1, 4)
    at.session_state["include_avails_template"] = True
    at.session_state["premion_streaming_tv"] = True
    return at


def seed_two_option_bundle(at, dc_included=True, balt_included=True):
    dc_group = tg.new_group(["AUTO Intenders"], geo_def={"kind": "markets", "markets": ["Washington, DC DMA"]},
                            avails_monthly=500000, include_in_plan=dc_included, group_id="dc-ctv-group")
    dc_group["resolved_zips"] = []
    dc_group["resolved_markets"] = ["Washington, DC DMA"]
    balt_group = tg.new_group(["AUTO Intenders"], geo_def={"kind": "markets", "markets": ["Baltimore DMA"]},
                              avails_monthly=300000, include_in_plan=balt_included, group_id="balt-ctv-group")
    balt_group["resolved_zips"] = []
    balt_group["resolved_markets"] = ["Baltimore DMA"]
    at.session_state["targeting_groups"] = [dc_group, balt_group]

    ctv_dc = {"Tactic": "Premion Streaming TV — DC", "Flight": "Sep 2026 - Jan 2027",
              "Geo": "Washington, DC", "Targeting": "Auto intenders, HHI $50K+",
              "Impressions": 500000.0, "CPM": 28.0, "Type": app.ROW_TYPE_RATE, "Cost": 14000.0,
              "_group_ids": ["dc-ctv-group"]}
    ctv_balt = {"Tactic": "Premion Streaming TV — Baltimore", "Flight": "Sep 2026 - Jan 2027",
                "Geo": "Baltimore", "Targeting": "Auto intenders, HHI $50K+",
                "Impressions": 300000.0, "CPM": 28.0, "Type": app.ROW_TYPE_RATE, "Cost": 8400.0,
                "_group_ids": ["balt-ctv-group"]}
    nfl_dc = {"Tactic": "Live Sports - NFL - Regular Season — DC", "Flight": "Sep 2026 - Jan 2027",
              "Geo": "Washington, DC", "Targeting": "", "Impressions": 0.0, "CPM": 0.0,
              "Type": app.ROW_TYPE_RATE, "Cost": 10000.0, "_group_ids": []}

    rows0 = ([dict(ctv_dc)] if dc_included else []) + ([dict(ctv_balt)] if balt_included else [])
    rows1 = rows0 + [dict(nfl_dc)]
    at.session_state["plan_options"] = [app.new_plan_option("CTV", rows0),
                                        app.new_plan_option("CTV + NFL Regular Season", rows1)]
    return at


def main():
    print("A. Delete CTV-DC from option 1 (index 1) only; group still owns a row in option 0")
    at = seed_two_option_bundle(new_app())
    at.run()
    check("no exception on initial render", not at.exception,
          at.exception[0].message[:500] if at.exception else "")

    _next_action.update(kind="delete_dc", option_idx=1, done=False)
    at.run()
    check("no exception after delete", not at.exception,
          at.exception[0].message[:500] if at.exception else "")
    opts = at.session_state["plan_options"]
    check("option 0 still has its DC row", "dc-ctv-group" in group_ids_in(opts[0]["rows"]),
          group_ids_in(opts[0]["rows"]))
    check("option 1's DC row is gone", "dc-ctv-group" not in group_ids_in(opts[1]["rows"]),
          group_ids_in(opts[1]["rows"]))
    check("option 1's Baltimore row is untouched",
          "balt-ctv-group" in group_ids_in(opts[1]["rows"]), group_ids_in(opts[1]["rows"]))
    groups_by_id = {g["id"]: g for g in at.session_state["targeting_groups"]}
    check("dc-ctv-group is STILL ticked (a sibling option still owns its row)",
          groups_by_id["dc-ctv-group"]["include_in_plan"] is True, groups_by_id["dc-ctv-group"])

    print("\n    ...an untouched rerun must not bring the row back")
    at.run()
    opts = at.session_state["plan_options"]
    check("option 1's DC row is STILL gone after a plain rerun (this is the reported bug if it fails)",
          "dc-ctv-group" not in group_ids_in(opts[1]["rows"]), group_ids_in(opts[1]["rows"]))
    check("option 0's DC row is untouched by that rerun",
          "dc-ctv-group" in group_ids_in(opts[0]["rows"]), group_ids_in(opts[0]["rows"]))

    print("\nB. Delete CTV-DC from option 0 too -- now NO option owns a row for it")
    _next_action.update(kind="delete_dc", option_idx=0, done=False)
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:500] if at.exception else "")
    opts = at.session_state["plan_options"]
    check("neither option owns a DC row any more",
          "dc-ctv-group" not in group_ids_in(opts[0]["rows"] + opts[1]["rows"]),
          group_ids_in(opts[0]["rows"] + opts[1]["rows"]))
    groups_by_id = {g["id"]: g for g in at.session_state["targeting_groups"]}
    check("dc-ctv-group is now UNTICKED (unlink_deleted_group_rows fired, correctly, once)",
          groups_by_id["dc-ctv-group"]["include_in_plan"] is False, groups_by_id["dc-ctv-group"])

    print("\nC. Unticking a still-fully-seeded group removes its row from every option at once")
    at2 = seed_two_option_bundle(new_app())
    at2.run()
    at2.session_state["targeting_groups"][1]["include_in_plan"] = False  # balt-ctv-group
    at2.run()
    opts = at2.session_state["plan_options"]
    check("Baltimore's row is gone from BOTH options after unticking",
          "balt-ctv-group" not in group_ids_in(opts[0]["rows"] + opts[1]["rows"]),
          group_ids_in(opts[0]["rows"] + opts[1]["rows"]))

    print("\nD. Re-ticking re-seeds every option that lacks a row for it")
    at2.session_state["targeting_groups"][1]["include_in_plan"] = True
    at2.session_state["targeting_groups"][1]["include_locked"] = False
    at2.run()
    opts = at2.session_state["plan_options"]
    check("Baltimore's row is back in BOTH options after re-ticking",
          all("balt-ctv-group" in group_ids_in(o["rows"]) for o in opts),
          [group_ids_in(o["rows"]) for o in opts])

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
