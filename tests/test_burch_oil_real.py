"""The real Burch Oil draft, replayed (local only: client data).

The notes and avails PDF a rep used (RFPID-272026) and a recorded live
draft response (2026-11-17), in the gitignored tests/fixtures/burch/. The
PDF goes through the real importer and the recorded response through the
real `apply_draft_to_form`, so this checks what the app does with what the
model actually wrote -- not a hand-written label.

What it holds the app to (Matt, 2026-11-17):
- the notes' RFPIs (4969829, 4969833) aren't on this document, and it has
  no HVAC-only row; the picks still land on exactly one Homeowner + HVAC row
  per ZIP list, and the plan prices at the notes' $11,850 a month;
- the substitution is a review row naming both RFPIs;
- a selection by ZIP list alone (what the model wrote on 2026-11-16, which
  put six rows and $32,250/month on the plan) sells nothing;
- "Show budget on Campaign Specs" starts off: the notes' "Expected totals"
  is the sum of their own lines, which is our figure -- also when the
  draft labels it the client's (what the model did on both 2026-11-16 and
  the first 2026-11-17 run).

The synthetic version that CI runs is tests/test_burch_oil_fixes.py.
"""
import copy
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TODAY", "2026-11-16")

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "burch"
NOTES_PATH = FIXTURES / "burch_notes.txt"
AVAILS_PATH = FIXTURES / "burch_avails.pdf"
DRAFT_PATH = FIXTURES / "burch.draft.json"
GROUPS_PATH = FIXTURES / "burch_groups.json"   # the import the recorded draft saw

HVAC_ROWS = ["RFPI-4968157", "RFPI-4968161"]
MONTHLY = 200_000 * 34 / 1000 + 100_000 * 34 / 1000 + 300_000 * 5.50 / 1000   # 11,850

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}" + (f"  -- {detail}" if detail != "" else ""))
        failures.append(label)


class _StubSt:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def imported_groups():
    import avails_pdf_import
    import group_scenario_fixtures as gsf
    groups, _report = gsf._import_groups(avails_pdf_import.parse_avails_pdf(str(AVAILS_PATH)))
    for g in groups:
        g["include_in_plan"] = False
    return groups


def remap_ids(draft, groups):
    """The recorded draft's row ids are the ones its own import made; a
    fresh import mints new ones -- carried across by RFPI."""
    recorded = {g["id"]: g.get("avails_rfpi") for g in json.loads(GROUPS_PATH.read_text("utf-8"))}
    now = {g.get("avails_rfpi"): g["id"] for g in groups}

    def fix(selection):
        if isinstance(selection, dict) and selection.get("ids"):
            selection["ids"] = [now.get(recorded.get(i), i) for i in selection["ids"]]

    fix(draft.get("group_selection"))
    for opt in draft.get("options") or []:
        fix(opt.get("group_selection"))
    return draft


def apply(app, draft, notes):
    groups = imported_groups()
    real = app.st
    stub = _StubSt()
    stub.session_state.update({
        "targeting_groups": groups, "market_choice": "DC",
        "flight_start": date(2027, 1, 1), "flight_end": date(2027, 12, 31),
        "avails_basis": app.AVAILS_BASIS_MONTHLY, "avails_mode": True,
        "draft_source_notes": notes, "client_name": "Burch Oil"})
    app.st = stub
    try:
        app.apply_draft_to_form(remap_ids(copy.deepcopy(draft), groups), notes=notes)
    finally:
        app.st = real
    return stub.session_state


def on_plan(state):
    return [g.get("avails_rfpi") for g in state["targeting_groups"] if g.get("include_in_plan")]


def monthly_cost(state):
    return sum(float(r["Cost"]) for r in state["plan_options"][0]["rows"])


def main():
    if not all(p.exists() for p in (NOTES_PATH, AVAILS_PATH, DRAFT_PATH, GROUPS_PATH)):
        print("SKIP: the real Burch Oil fixtures (tests/fixtures/burch/) aren't on this machine.")
        return
    import db
    db.fetch_audiences = lambda: (None, "stubbed for test isolation")
    import app
    notes = NOTES_PATH.read_text("utf-8")
    draft = json.loads(DRAFT_PATH.read_text("utf-8"))

    print("The recorded live draft, as the model wrote it")
    state = apply(app, draft, notes)
    check("exactly the two Homeowner + HVAC rows are on the plan, one per ZIP list",
          on_plan(state) == HVAC_ROWS, on_plan(state))
    check(f"the plan prices at ${MONTHLY:,.0f} a month", abs(monthly_cost(state) - MONTHLY) < 1,
          monthly_cost(state))
    rows = [str((i.get("row") or {}).get("text") or i.get("question") or "")
            for i in state.get("draft_review_items") or []]
    check("the RFPI substitution is a review row, naming both notes' RFPIs",
          any("4969829" in r for r in rows) and any("4969833" in r for r in rows), rows)
    check("'Show budget on Campaign Specs' starts off",
          state.get("show_budget_on_specs") is False, state.get("show_budget_on_specs"))
    check("Budget & Allocation has no dollars", "$" not in (state.get("budget_text") or ""),
          state.get("budget_text"))

    print("\nThe same draft, labelled the way the model did on 2026-11-16 and earlier today")
    labelled = apply(app, {**draft, "budget_source": "client",
                           "stated_budget": {"amount": 11850, "period": "month"}}, notes)
    check("'Expected totals' is the lines added up: still ours, box off",
          labelled.get("show_budget_on_specs") is False
          and labelled.get("draft_budget_source") == "seller",
          (labelled.get("draft_budget_source"), labelled.get("show_budget_on_specs")))
    check("... and Budget & Allocation has no dollars", "$" not in (labelled.get("budget_text") or ""),
          labelled.get("budget_text"))

    print("\nThe 2026-11-16 selection: the ZIP lists alone")
    by_area = apply(app, {**draft, "group_selection": {"mode": "named",
                                                       "match": ["Zip List #1", "Zip List #2"]}},
                    notes)
    check("a geography alone puts no avails row on the plan (it put six, $32,250/month)",
          on_plan(by_area) == [], on_plan(by_area))
    check("... and says so in the review",
          any("no audience" in str(i) for i in by_area.get("draft_review_items") or []),
          [i.get("kind") for i in by_area.get("draft_review_items") or []])

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("All real Burch Oil checks passed.")


if __name__ == "__main__":
    main()
