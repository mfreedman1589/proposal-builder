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
# The last dataframe the D2 grid was actually handed to DISPLAY -- captured
# regardless of _next_action, so a test can assert on what a rep would see
# (e.g. the Label column's placeholder text) without needing an edit at all.
_last_displayed = {"df": None}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if not key.startswith("avails_editor"):
        return data
    _last_displayed["df"] = data.copy()
    if _next_action["done"]:
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
import group_scenario_fixtures as gsf  # noqa: E402

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

    at2.session_state["_pending_entity_group"] = {"gids": [ga["id"], gb["id"]], "label": ""}
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
    check("neither group got a real entity_label -- left blank, no label was typed",
          grouped_a["entity_label"] == "" and grouped_b["entity_label"] == "", (grouped_a, grouped_b))

    print("\naudit follow-up (2026-09-04): a grouped-but-unlabeled entity shows a "
          "PLACEHOLDER in the D2 grid's own Label cell, not blank -- and the placeholder "
          "is never mistaken for a real edit and baked in as the entity's actual label "
          "on a later, untouched rerun")
    displayed = _last_displayed["df"]
    placeholder_cells = displayed.loc[displayed["gid"].isin([ga["id"], gb["id"]]), "Label"].tolist()
    check("both rows show the same '(grouped -- 2 rows)' placeholder, not blank",
          placeholder_cells == ["(grouped -- 2 rows)", "(grouped -- 2 rows)"], placeholder_cells)

    run(at2)  # untouched -- no _next_action queued
    check("no exception on the untouched rerun", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    still_a, still_b = by_id(at2, ga["id"]), by_id(at2, gb["id"])
    check("the placeholder text was NEVER written into entity_label -- still genuinely blank",
          still_a["entity_label"] == "" and still_b["entity_label"] == "", (still_a, still_b))
    check("still grouped, still locked -- an untouched rerun changes nothing real",
          tg.entity_id_of(still_a) == tg.entity_id_of(still_b)
          and still_a["entity_locked"] and still_b["entity_locked"], (still_a, still_b))

    print("\ntyping a real name over the placeholder renames the entity normally -- "
          "the placeholder never gets in the way of a real edit")
    run(at2, "rename_label", gid=ga["id"], new_label="Denver + Atlanta")
    check("no exception", not at2.exception, at2.exception[0].message[:400] if at2.exception else "")
    renamed_a, renamed_b = by_id(at2, ga["id"]), by_id(at2, gb["id"])
    check("both groups now carry the real typed name, not the placeholder",
          renamed_a["entity_label"] == "Denver + Atlanta"
          and renamed_b["entity_label"] == "Denver + Atlanta", (renamed_a, renamed_b))

    print("\nnaming the entity AT GROUPING TIME (the optional label field) wins outright -- "
          "one action instead of group-then-hunt-the-Label-column-then-type")
    gf = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Reston"},
                      avails_monthly=50_000)
    gg = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Ashburn"},
                      avails_monthly=60_000, entity_label="An Existing Label", entity_locked=True)
    at3 = new_app([gf, gg])
    run(at3)
    at3.session_state["_pending_entity_group"] = {
        "gids": [gf["id"], gg["id"]], "label": "Typed At Grouping Time",
    }
    at3.run()
    check("no exception", not at3.exception, at3.exception[0].message[:400] if at3.exception else "")
    named_f, named_g = by_id(at3, gf["id"]), by_id(at3, gg["id"])
    check("the typed label wins for BOTH groups, even though gg already had its own "
          "non-blank entity_label -- naming at grouping time is the rep's explicit intent",
          named_f["entity_label"] == "Typed At Grouping Time"
          and named_g["entity_label"] == "Typed At Grouping Time", (named_f, named_g))

    print("\nleaving the label blank at grouping time still falls back to the pre-existing "
          "rule -- a selected group's own non-blank entity_label wins")
    gh = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Vienna"},
                      avails_monthly=40_000)
    gi = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Herndon"},
                      avails_monthly=45_000, entity_label="Pre-Existing Label")
    at4 = new_app([gh, gi])
    run(at4)
    at4.session_state["_pending_entity_group"] = {"gids": [gh["id"], gi["id"]], "label": ""}
    at4.run()
    check("no exception", not at4.exception, at4.exception[0].message[:400] if at4.exception else "")
    fallback_h, fallback_i = by_id(at4, gh["id"]), by_id(at4, gi["id"])
    check("the pre-existing fallback still applies when nothing is typed at grouping time",
          fallback_h["entity_label"] == "Pre-Existing Label"
          and fallback_i["entity_label"] == "Pre-Existing Label", (fallback_h, fallback_i))

    print("\nthe D2 grid surfaces a one-line summary OUTSIDE the grouping expander whenever "
          "any grouping exists -- so the evidence doesn't disappear the moment a rep "
          "collapses the very expander they just used")
    captions4 = [str(c.value) if hasattr(c, "value") else str(c) for c in at4.caption]
    check("a caption names the grouping count and total rows, outside the expander",
          any("1 grouping(s) in place, 2 row(s) total" in c for c in captions4), captions4)

    print("\na LONE group (the ordinary, ungrouped case) still shows a blank Label cell -- "
          "the placeholder only ever appears for a REAL multi-row entity")
    gj = tg.new_group(["Homeowners"], geo_def={"kind": "text", "label": "Solo"},
                      avails_monthly=30_000)
    at5 = new_app([gj])
    run(at5)
    check("no exception", not at5.exception, at5.exception[0].message[:400] if at5.exception else "")
    displayed_solo = _last_displayed["df"]
    solo_cell = displayed_solo.loc[displayed_solo["gid"] == gj["id"], "Label"].tolist()
    check("blank, not a placeholder -- there's nothing to group it with",
          solo_cell == [""], solo_cell)

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

    print("\nFLOW_REWORK_PLAN.md Phase 3, commit 9: a drafted group_entities label "
          "never overwrites a rep's own locked entity")
    gc = tg.new_group(["C"], avails_monthly=50_000, group_id="gc")
    gd = tg.new_group(["D"], avails_monthly=60_000, group_id="gd",
                      entity_id="rep-set", entity_label="Rep's Own Name", entity_locked=True)
    ge = tg.new_group(["D"], avails_monthly=70_000, group_id="ge")   # same audience as gd, ungrouped
    groups_locked_case = [gc, gd, ge]
    unresolved, internal = app.apply_draft_group_entities(
        groups_locked_case,
        [{"label": "Model's Guess", "ids": ["gd", "ge"]},
         {"label": "Solo C", "ids": ["gc"]}])   # a single id -- NAMES, doesn't group (LiveWell's shape)
    check("gd (entity_locked) is completely untouched -- id, label, lock all survive",
          gd["entity_id"] == "rep-set" and gd["entity_label"] == "Rep's Own Name"
          and gd["entity_locked"] is True, gd)
    check("ge (unlocked, but gd -- its only possible partner -- is locked) stays its own entity too, "
          "since match_groups_to_selection only ever sees the ELIGIBLE (unlocked, ungrouped) list",
          tg.entity_id_of(ge) == ge["id"], ge)
    check("a single-id entry NAMES, never groups -- gc keeps its own entity_id",
          tg.entity_id_of(gc) == gc["id"], gc)
    check("but gc DOES get the label -- this is the LiveWell fix: a single already-separate "
          "row can still be named, not just merged",
          gc["entity_label"] == "Solo C", gc)
    check("gc stays unlocked -- a draft's naming is a suggestion, same as its grouping",
          gc["entity_locked"] is False, gc)
    check("an internal note discloses the single-row naming too (ids alone, no notes-quoted "
          "phrase -- a seller check, same routing rule as a multi-row group)",
          any("Solo C" in n for n in internal) and not unresolved, (unresolved, internal))

    print("\nWilmington University (RFPID-253956), real document: a drafted group_entities "
          "payload ties its 3 undergraduate audiences together -- the only mechanism that "
          "can, since commit 7's deterministic pass correctly leaves them ungrouped (no "
          "shared string). The plan still shows 5 lines -- labeling is not merging.")
    scn = gsf.build_wilmington()
    wilm_groups = [dict(g) for g in scn["groups"]]
    for g in wilm_groups:
        g.pop("_flight_label", None)
    undergrad_terms = ["LIFESTAGE College Planning Parents", "LIFESTAGE Prospective College Students",
                       "LIFESTAGE Higher Education Intender"]
    check("setup: 5 real Wilmington groups, none pre-grouped",
          len(wilm_groups) == 5
          and len({tg.entity_id_of(g) for g in wilm_groups}) == 5, wilm_groups)

    w_unresolved, w_internal = app.apply_draft_group_entities(
        wilm_groups, [{"label": "Undergraduate", "match": undergrad_terms}])
    undergrad_groups = [g for g in wilm_groups if tg.audience_label(g) in undergrad_terms]
    other_groups = [g for g in wilm_groups if tg.audience_label(g) not in undergrad_terms]
    check("all 3 undergrad audiences found and grouped",
          len(undergrad_groups) == 3, [tg.audience_label(g) for g in wilm_groups])
    check("the 3 undergrad groups now share ONE entity_id",
          len({tg.entity_id_of(g) for g in undergrad_groups}) == 1, undergrad_groups)
    check("all 3 carry the model's own label, \"Undergraduate\"",
          all(tg.entity_label_of(g) == "Undergraduate" for g in undergrad_groups), undergrad_groups)
    check("the OTHER 2 groups (Online Education; Career Employed/Education Services) are "
          "untouched -- still their own, separate entities",
          len({tg.entity_id_of(g) for g in other_groups}) == 2, other_groups)
    check("a match phrase (the notes' own words) discloses client-facing, in unresolved",
          bool(w_unresolved) and not w_internal, (w_unresolved, w_internal))

    print("\n...and the plan still shows all 5 lines -- entity LABELING never merges rows; "
          "only a rep's own merge (the existing mechanism) collapses them, to 3")
    state = gsf.build_session_state(dict(scn, groups=wilm_groups))
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    for key, value in state.items():
        at.session_state[key] = value
    at.run()
    check("no exception rendering the form with the labeled-but-unmerged groups",
          not at.exception, at.exception[0].message[:400] if at.exception else "")
    if not at.exception:
        rows = at.session_state["plan_options"][0]["rows"]
        check("still exactly 5 Premion Streaming TV lines -- labeling didn't merge anything",
              len(rows) == 5, len(rows))

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
