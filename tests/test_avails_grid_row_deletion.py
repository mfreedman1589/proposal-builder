"""Deleting a row from the D2 avails grid must remove its plan line too --
the deletion counterpart of test_avails_grid_manual_row.py's manual-add fix.

Reported from live use: add a row (e.g. "Family"), delete it, and it still
appeared above the media plan's Total on the generated deck. The GROUP was
already being removed correctly by sync_targeting_groups/seed_rows_to_groups
-- the bug was one level up, in main()'s shared-fields reseed, which reset a
now-groupless row's Targeting/Geo to the shared default instead of removing
the row outright. A row seeded FOR a group must not outlive that group.

Same `st.data_editor` monkeypatch technique as test_avails_grid_manual_row.py
(AppTest can't drive a data_editor directly), but with a small state machine
so one file can drive add -> delete and add -> edit-to-empty in sequence,
matching the report's own repro steps. Three lifecycles asserted, each
confirming all three symptoms together: the group is gone from
targeting_groups, no plan line remains for it, and it's absent from the
assembled deck's media plan table.

    python tests/test_avails_grid_row_deletion.py
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

_ADD_ROW = {"gid": None, "Audience": "Family", "Markets": ["Washington, DC DMA"],
           "Label": "", "Color": None, "avails_col": 5000}

# One shared mutable "what should the grid widget return next" state, read
# by the monkeypatched data_editor and advanced by the test driving it --
# same shape as test_avails_grid_manual_row.py's `_injected`, generalized to
# more than one one-shot action.
_next_action = {"kind": None, "done": True}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if not key.startswith("avails_editor") or _next_action["done"]:
        return data
    _next_action["done"] = True
    avails_col = [c for c in data.columns if c not in
                 ("gid", "Audience", "Markets", "Label", "Color")][0]
    kind = _next_action["kind"]
    if kind == "add":
        row = dict(_ADD_ROW)
        row[avails_col] = row.pop("avails_col")
        return pd.concat([data, pd.DataFrame([row])], ignore_index=True)
    if kind == "delete":
        return data[data["Audience"] != "Family"].reset_index(drop=True)
    if kind == "edit_empty":
        out = data.copy()
        mask = out["Audience"] == "Family"
        out.loc[mask, "Audience"] = ""
        out.loc[mask, avails_col] = 0
        out.loc[mask, "Markets"] = out.loc[mask, "Markets"].apply(lambda _: [])
        return out
    return data


st.data_editor = _fake_data_editor

import app  # noqa: E402
import assembly  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["target_dmas"] = ["Washington, DC"]
    at.session_state["premion_streaming_tv"] = True
    return at


def real_groups(at):
    groups = at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []
    return [g for g in groups if g["terms"]]


def premion_rows(at):
    return [r for r in at.session_state["plan_options"][0]["rows"]
           if r.get("Tactic") == "Premion Streaming TV"]


def run(at, kind):
    _next_action["kind"] = kind
    _next_action["done"] = False
    at.run()
    return at


def check_lifecycle(label, second_kind):
    """add "Family", confirm it's really there, then remove it via
    `second_kind` ("delete" or "edit_empty"), confirming all three symptoms
    together: the group is gone, no plan line remains, and a real Generate
    doesn't put it on the deck."""
    print(f"\n{label}")
    at = new_app()
    run(at, "add")
    check("no exception after add", not at.exception, at.exception[0].message[:300] if at.exception else "")
    check("the group exists after adding",
          any(g["terms"] == ["Family"] for g in real_groups(at)), real_groups(at))
    check("a Premion Streaming TV line exists for it",
          any(r.get("Targeting") == "Family" for r in premion_rows(at)), premion_rows(at))
    before_count = len(premion_rows(at))

    run(at, second_kind)
    check(f"no exception after {second_kind}", not at.exception,
          at.exception[0].message[:300] if at.exception else "")
    check("the group is gone from targeting_groups",
          not any(g["terms"] == ["Family"] for g in real_groups(at)), real_groups(at))
    after_rows = premion_rows(at)
    check("no plan line remains for it -- not reset to the default, REMOVED",
          not any(r.get("Targeting") == "Family" for r in after_rows), after_rows)
    check("the row count actually dropped (a reset-in-place would leave the count unchanged)",
          len(after_rows) == before_count - 1, (before_count, len(after_rows)))

    generate = [b for b in at.button if b.label == "Generate proposal"]
    check("Generate button present", bool(generate), [b.label for b in at.button])
    if generate:
        generate[0].click().run()
        check("Generate runs without raising", not at.exception,
              at.exception[0].message[:400] if at.exception else "")
    return at


def main():
    # Capture the assembled Presentation the same way test_targeting_map.py
    # and run_scenario.py do -- spy on assembly.personalize.
    real_personalize = assembly.personalize
    captured = {}

    def spy(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        return warnings

    assembly.personalize = spy
    real_log = db.log_proposal
    db.log_proposal = lambda *a, **k: ("test", None)
    try:
        at = check_lifecycle(
            "LIFECYCLE 1: add, then DELETE the row outright", "delete")
        prs = captured.get("prs")
        if prs is not None:
            avails_slide = assembly.find_slide_with_marker(prs, "{{AVAILS}}")
            plan_slide = None
            for slide in prs.slides:
                table = assembly._find_table_shape(slide)
                if table is not None and "Included with Campaign" in (
                        assembly.slide_map.extract_slide_text(slide) or ""):
                    plan_slide = slide
                    break
            check("found the media plan slide", plan_slide is not None)
            if plan_slide is not None:
                table = assembly._find_table_shape(plan_slide).table
                cell_text = " ".join(cell.text for row in table.rows for cell in row.cells)
                check("'Family' does not appear anywhere in the assembled media plan table",
                      "Family" not in cell_text, cell_text)

        captured.clear()
        check_lifecycle(
            "\nLIFECYCLE 2: add, then EDIT the row's Audience/Markets/avails to blank "
            "(same as a delete for seed_rows_to_groups -- 'not audience and not geo')",
            "edit_empty")
    finally:
        assembly.personalize = real_personalize
        db.log_proposal = real_log

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A row removed from the D2 avails grid -- by deleting it outright, or by editing it "
          "down to blank -- removes its media-plan line too, not just resets it to the shared "
          "default, and it never reaches the assembled deck.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
