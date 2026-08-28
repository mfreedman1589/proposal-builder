"""FLOW_REWORK_PLAN.md Phase 3: the D2 avails grid's "Label" column is now
the entity a row is FOR (e.g. "Toyota of Annapolis"), separate from the
renamed "Geo Label" column (the pre-existing Geo-cell override, `name`,
completely unchanged in behavior).

Four things proved here, matching CLAUDE.md's own settled decisions:
    1. Renaming an entity via the Label cell propagates to EVERY group
       sharing that row's entity_id, not just the one row edited -- and
       changes ONLY displayed text: ids, rows, avails all survive untouched.
    2. An untouched rerun never drops entity_label/entity_id/entity_locked --
       the same `_cell_unchanged` fold-back guard already proven for
       Audience/Markets/Geo Label/Color, now proven for this fifth field.
    3. "Group selected rows" (apply_pending_entity_group) assigns a fresh
       shared entity_id to the selected groups, locked.
    4. "Ungroup" (apply_pending_entity_ungroup) is the exact inverse --
       every group sharing that entity_id gets its own id back.

    python tests/test_entity_label.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import streamlit as st  # noqa: E402

# One shared mutable "what should the grid widget return next" state, same
# shape test_avails_grid_row_deletion.py/test_color_lifecycle.py use.
_next_action = {"kind": None, "done": True}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if not key.startswith("avails_editor") or _next_action["done"]:
        return data
    _next_action["done"] = True
    out = data.copy()
    if _next_action["kind"] == "rename_label":
        mask = out["gid"] == _next_action["gid"]
        out.loc[mask, "Label"] = _next_action["new_label"]
    return out


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


def sget(at, key, default=None):
    try:
        return at.session_state[key]
    except (KeyError, AttributeError):
        return default


def new_app(groups):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["targeting_groups"] = groups
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def real_groups(at):
    groups = sget(at, "targeting_groups") or []
    return [g for g in groups if g["terms"]]


def by_id(at, gid):
    return next((g for g in real_groups(at) if g["id"] == gid), None)


def run(at, kind=None, **extra):
    _next_action["kind"] = kind
    _next_action["done"] = kind is None
    _next_action.update(extra)
    at.run()
    return at


def main():
    print("renaming an entity via the Label cell propagates to every group "
          "sharing that entity_id, and changes ONLY displayed text")
    g5 = tg.new_group(["AUTO Make Subaru"], geo_def={"kind": "text", "label": ""},
                      name="5mi radius", avails_monthly=165_687, entity_id="toyota-annapolis")
    g10 = tg.new_group(["AUTO Make Subaru"], geo_def={"kind": "text", "label": ""},
                       name="10mi radius", avails_monthly=602_647, entity_id="toyota-annapolis")
    at = new_app([g5, g10])
    run(at)
    check("no exception on first render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    before = [dict(g) for g in real_groups(at)]
    check("both groups start with a blank entity_label",
          all(tg.entity_label_of(g) == "" for g in before), before)

    run(at, "rename_label", gid=g5["id"], new_label="Toyota of Annapolis")
    check("no exception after the rename", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    after_g5, after_g10 = by_id(at, g5["id"]), by_id(at, g10["id"])
    check("the edited row's entity_label is the new text",
          after_g5 is not None and after_g5["entity_label"] == "Toyota of Annapolis", after_g5)
    check("the SIBLING row (never directly edited) also picked up the new entity_label",
          after_g10 is not None and after_g10["entity_label"] == "Toyota of Annapolis", after_g10)
    check("both groups are now entity_locked (a rep touched this entity)",
          after_g5["entity_locked"] and after_g10["entity_locked"], (after_g5, after_g10))
    check("entity_id itself never moved -- renaming is not regrouping",
          after_g5["entity_id"] == "toyota-annapolis" and after_g10["entity_id"] == "toyota-annapolis",
          (after_g5["entity_id"], after_g10["entity_id"]))
    check("ids are unchanged", after_g5["id"] == g5["id"] and after_g10["id"] == g10["id"])
    check("avails are unchanged on both groups",
          after_g5["avails_monthly"] == 165_687 and after_g10["avails_monthly"] == 602_647,
          (after_g5, after_g10))

    print("\nan UNTOUCHED rerun never drops entity_label/entity_id/entity_locked "
          "(same fold-back guard as Audience/Markets/Geo Label/Color)")
    snapshot_before = [dict(g) for g in real_groups(at)]
    run(at)
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    snapshot_after = [dict(g) for g in real_groups(at)]
    check("every group survives byte-for-byte across an untouched rerun",
          snapshot_before == snapshot_after, (snapshot_before, snapshot_after))

    print("\n\"Group selected rows\" (apply_pending_entity_group) assigns a fresh "
          "shared entity_id, locked")
    ga = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Denver"},
                      avails_monthly=100_000)
    gb = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Atlanta"},
                      avails_monthly=200_000)
    at2 = new_app([ga, gb])
    run(at2)
    check("no exception", not at2.exception, at2.exception[0].message[:400] if at2.exception else "")
    check("the two fixture groups start as separate entities",
          tg.entity_id_of(by_id(at2, ga["id"])) != tg.entity_id_of(by_id(at2, gb["id"])))

    at2.session_state["_pending_entity_group"] = [ga["id"], gb["id"]]
    at2.run()
    check("no exception after grouping", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    grouped_a, grouped_b = by_id(at2, ga["id"]), by_id(at2, gb["id"])
    check("both groups now share one entity_id",
          grouped_a is not None and grouped_b is not None
          and tg.entity_id_of(grouped_a) == tg.entity_id_of(grouped_b),
          (grouped_a, grouped_b))
    check("that shared id is neither group's OWN id (a fresh id was minted)",
          tg.entity_id_of(grouped_a) not in (ga["id"], gb["id"]), tg.entity_id_of(grouped_a))
    check("both are entity_locked", grouped_a["entity_locked"] and grouped_b["entity_locked"])
    check("grouping never touches row count -- still 2 separate groups",
          len(real_groups(at2)) == 2, real_groups(at2))

    print("\n\"Ungroup\" (apply_pending_entity_ungroup) is the exact inverse -- "
          "every group sharing that entity_id gets its own id back")
    shared_id = tg.entity_id_of(grouped_a)
    at2.session_state["_pending_entity_ungroup"] = shared_id
    at2.run()
    check("no exception after ungrouping", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    ungrouped_a, ungrouped_b = by_id(at2, ga["id"]), by_id(at2, gb["id"])
    check("each group's entity_id is its own id again",
          tg.entity_id_of(ungrouped_a) == ungrouped_a["id"]
          and tg.entity_id_of(ungrouped_b) == ungrouped_b["id"],
          (ungrouped_a, ungrouped_b))
    check("they no longer share an entity_id",
          tg.entity_id_of(ungrouped_a) != tg.entity_id_of(ungrouped_b))
    check("both stay entity_locked (ungrouping is as deliberate as grouping)",
          ungrouped_a["entity_locked"] and ungrouped_b["entity_locked"])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Renaming an entity changes only displayed text, propagated to every group that "
          "shares it; an untouched rerun never drops entity state; grouping and ungrouping "
          "are explicit, symmetric rep actions that never change row count.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
