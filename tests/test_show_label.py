"""FLOW_REWORK_PLAN.md Phase 3, commit 10: "Show Label in plan".

Off (the default): the preview and the generated deck are byte-identical
to before this phase -- entity_prefixed_targeting is a no-op. On: a group-
owned line's Targeting text gets its entity's Label prefixed, on both the
app's own preview and the deck payload -- and never leaks into the
editable grid's own `row["Targeting"]` value, which stays exactly what a
rep typed or the app seeded.

    python tests/test_show_label.py
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
db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def ss(at, key, default=None):
    """AppTest's session_state has no `.get` -- the safe substitute used
    throughout this suite (same helper other test files define)."""
    return at.session_state[key] if key in at.session_state else default


def group_row(gid, targeting="Homeowners"):
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep - Nov", "Geo": "Denver",
            "Targeting": targeting, "Impressions": 100000.0, "CPM": 30.0,
            "Type": app.ROW_TYPE_RATE, "Cost": 3000.0, "_group_ids": [gid]}


def main():
    print("entity_prefixed_targeting: pure-function behavior")
    g_labeled = tg.new_group(["Homeowners"], group_id="g1", entity_label="Toyota of Annapolis")
    g_unlabeled = tg.new_group(["Homeowners"], group_id="g2")
    groups_by_id = {"g1": g_labeled, "g2": g_unlabeled}
    row_labeled = group_row("g1")
    row_unlabeled = group_row("g2")
    hand_typed_row = {"Tactic": "Premion Streaming TV", "Flight": "", "Geo": "", "Targeting": "Hand-typed"}

    check("toggle OFF: plain Targeting, byte-identical, regardless of entity label",
          app.entity_prefixed_targeting(row_labeled, groups_by_id, False) == "Homeowners")
    check("toggle ON, labeled entity: prefixed",
          app.entity_prefixed_targeting(row_labeled, groups_by_id, True)
          == "Toyota of Annapolis | Homeowners")
    check("toggle ON, unlabeled entity: falls through to plain Targeting",
          app.entity_prefixed_targeting(row_unlabeled, groups_by_id, True) == "Homeowners")
    check("toggle ON, no _group_ids at all (hand-typed/drafted/quick-add row): falls through",
          app.entity_prefixed_targeting(hand_typed_row, groups_by_id, True) == "Hand-typed")
    check("original row dict is untouched either way -- row['Targeting'] never mutated",
          row_labeled["Targeting"] == "Homeowners", row_labeled)

    print("\ncompute_plan_totals: toggle OFF is byte-identical to before this phase")
    rows = [group_row("g1"), group_row("g2", targeting="Business owners")]
    totals_off = app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov",
                                         groups_by_id=groups_by_id, show_entity_label=False)
    check("no show_entity_label kwarg at all also defaults off (byte-identical call sites)",
          app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov",
                                  groups_by_id=groups_by_id) == totals_off)
    targetings_off = [r["targeting"] for r in totals_off["preview_rows"]]
    check("targeting text is the plain, unprefixed text with the toggle off",
          targetings_off == ["Homeowners", "Business owners"], targetings_off)

    print("\ncompute_plan_totals: toggle ON prefixes only the labeled entity's line")
    totals_on = app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov",
                                        groups_by_id=groups_by_id, show_entity_label=True)
    targetings_on = [r["targeting"] for r in totals_on["preview_rows"]]
    check("the labeled entity's line is prefixed",
          targetings_on[0] == "Toyota of Annapolis | Homeowners", targetings_on)
    check("the unlabeled entity's line is untouched",
          targetings_on[1] == "Business owners", targetings_on)
    check("nothing else about the totals differs between on and off "
          "(same cost/impressions either way)",
          {k: v for k, v in totals_on.items() if k != "preview_rows"}
          == {k: v for k, v in totals_off.items() if k != "preview_rows"})
    check("the ORIGINAL rows list is untouched by either call -- toggling never leaks "
          "into row['Targeting']",
          rows[0]["Targeting"] == "Homeowners" and rows[1]["Targeting"] == "Business owners",
          rows)

    print("\nComputed default: 2+ ticked groups, one shared audience, 2+ distinct geo labels "
          "-- exactly LiveWell's shape -- turns the checkbox on by itself, at most once, and "
          "a rep's own later choice is never fought afterward")
    from streamlit.testing.v1 import AppTest
    from datetime import date

    def new_app():
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=180)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        at.session_state["include_avails_template"] = True
        at.session_state["premion_streaming_tv"] = True
        at.session_state["market_choice"] = "DC"
        at.session_state["flight_start"] = date(2026, 10, 1)
        at.session_state["flight_end"] = date(2026, 10, 31)
        return at

    def livewell_shaped_groups():
        return [
            tg.new_group(["LIFESTYLE Pets"], group_id="lw1", name="Alexandria",
                        include_in_plan=True, avails_monthly=700_000),
            tg.new_group(["LIFESTYLE Pets"], group_id="lw2", name="Falls Church",
                        include_in_plan=True, avails_monthly=300_000),
        ]

    at1 = new_app()
    at1.session_state["targeting_groups"] = livewell_shaped_groups()
    at1.run()
    check("no exception", not at1.exception, at1.exception[0].message[:400] if at1.exception else "")
    check("fires: show_entity_label is True with no rep action at all",
          ss(at1, "show_entity_label") is True, ss(at1, "show_entity_label"))
    check("the one-shot lock is set", ss(at1, "show_entity_label_suggested") is True)

    at1.session_state["show_entity_label"] = False
    at1.run()
    check("a rep's own uncheck afterward is respected, not re-forced back on",
          ss(at1, "show_entity_label") is False, ss(at1, "show_entity_label"))

    print("\nDoes NOT fire: fewer than 2 ticked groups")
    at2 = new_app()
    groups2 = livewell_shaped_groups()
    groups2[1]["include_in_plan"] = False   # only one actually ticked
    at2.session_state["targeting_groups"] = groups2
    at2.run()
    check("stays off -- only one group is actually on the plan",
          ss(at2, "show_entity_label") in (None, False), ss(at2, "show_entity_label"))

    print("\nDoes NOT fire: two ticked groups, but different audiences (Annapolis's shape -- "
          "the make already carries the distinction, a label would add nothing)")
    at3 = new_app()
    at3.session_state["targeting_groups"] = [
        tg.new_group(["AUTO Make Subaru"], group_id="a1", name="10mi", include_in_plan=True),
        tg.new_group(["AUTO Make Hyundai"], group_id="a2", name="10mi", include_in_plan=True),
    ]
    at3.run()
    check("stays off -- two different audiences, nothing to disambiguate with a Label",
          ss(at3, "show_entity_label") in (None, False), ss(at3, "show_entity_label"))

    print("\nDoes NOT fire: two ticked groups, same audience, but IDENTICAL geo too "
          "(Wilmington/Plaza's shape -- nothing for a Label to distinguish that Targeting "
          "doesn't already)")
    at4 = new_app()
    at4.session_state["targeting_groups"] = [
        tg.new_group(["DEMO Age 35-64"], group_id="w1", name="Same Zip List", include_in_plan=True),
        tg.new_group(["DEMO Age 35-64"], group_id="w2", name="Same Zip List", include_in_plan=True),
    ]
    at4.run()
    check("stays off -- same audience AND same geo label, nothing to distinguish",
          ss(at4, "show_entity_label") in (None, False), ss(at4, "show_entity_label"))

    print("\nshow_entity_label=True with ZERO plan rows doesn't raise -- pd.DataFrame([]) has "
          "zero COLUMNS, not just zero rows, so the grid's own Label-column insert "
          "(media_plan_df.insert(1, \"Label\", ...)) is genuinely out of bounds at position 1. "
          "Real, pre-existing bug (predates this session), only reachable once the computed "
          "default above can turn the checkbox on and a rep then unticks every group -- found "
          "live by the full test sweep, not guessed at.")
    at5 = new_app()
    at5.session_state["targeting_groups"] = livewell_shaped_groups()
    at5.run()
    check("setup: the default fired, show_entity_label is True", ss(at5, "show_entity_label") is True)
    check("setup: 2 rows on the plan", len(ss(at5, "plan_options")[0]["rows"]) == 2)
    for g in at5.session_state["targeting_groups"]:
        g["include_in_plan"] = False
    at5.run()
    check("no exception when every group is unticked while Show Label is still on",
          not at5.exception, at5.exception[0].message[:400] if at5.exception else "")
    check("zero rows on the plan now", len(ss(at5, "plan_options")[0]["rows"]) == 0,
          ss(at5, "plan_options")[0]["rows"])
    check("show_entity_label itself is untouched by unticking -- still True, the rep's own "
          "checkbox state, not something this bug or its fix should silently flip",
          ss(at5, "show_entity_label") is True)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("\"Show Label in plan\" prefixes a group-owned line's preview/deck Targeting text "
          "with its entity's Label when on, is a byte-identical no-op when off (the default), "
          "never mutates the underlying row, and turns itself on by default -- once, overridable "
          "-- exactly when the Targeting column alone can't tell same-audience rows apart.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
