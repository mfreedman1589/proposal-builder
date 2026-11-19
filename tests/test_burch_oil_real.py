"""The real Burch Oil draft, replayed (local only: client data).

The notes a rep used (RFPID-272026) and two avails files for it, in the
gitignored tests/fixtures/burch/:

- burch_avails.pdf is the 11-16 file (ver0 of 11-16-2026): per ZIP list a
  Homeowner + HVAC row, a Homeowner row, a Homeowner + older-home row and an
  HVAC-only row -- RFPI 4969829 (Zip List #1) and 4969833 (Zip List #2), the
  ones the notes name. burch.draft.json is the live draft made against it
  (2026-11-19).
- burch_avails_11-12.pdf is the earlier 11-12 file (RFPIs 4968157-4968164),
  which has no HVAC-only row. burch_11-12.draft.json is the live draft made
  against it (2026-11-17) and burch_11-12_groups.json the import it saw.
  The 2026-11-17 regression used this file by mistake (Matt, 2026-11-19).

The PDFs go through the real importer and the recorded responses through the
real `apply_draft_to_form`, so this checks what the app does with what the
model actually wrote.

What it holds the app to (Matt, 2026-11-19):
- 11-16: exactly the two HVAC-only rows on the plan, $11,850 a month, Growth
  Areas about 89% of its avails as a soft note (no hold), nothing missing,
  "Show budget on Campaign Specs" off, and a totals caption that says the
  lines were sized by the notes' counts -- not "in proportion to max avails";
- 11-12: no audience line on the plan -- nothing stands in for an RFPI the
  avails don't carry -- and two blocking rows, one per missing RFPI, naming
  the RFPIs on file;
- a selection by ZIP list alone (the 2026-11-16 draft) sells nothing.

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
os.environ.setdefault("PROPOSAL_BUILDER_TODAY", "2026-11-19")

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "burch"
NOTES_PATH = FIXTURES / "burch_notes.txt"
AVAILS_PATH = FIXTURES / "burch_avails.pdf"              # 11-16
DRAFT_PATH = FIXTURES / "burch.draft.json"               # drafted against 11-16
OLD_AVAILS_PATH = FIXTURES / "burch_avails_11-12.pdf"
OLD_DRAFT_PATH = FIXTURES / "burch_11-12.draft.json"
OLD_GROUPS_PATH = FIXTURES / "burch_11-12_groups.json"   # the import that draft saw

HVAC_ONLY = ["RFPI-4969829", "RFPI-4969833"]
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


def imported_groups(app, path):
    """The app's own import, as the intake uploader runs it -- with the
    avails figures the plan's share checks read (the scenario helper's
    import leaves them at 0)."""
    import avails_pdf_import
    real = app.st
    stub = _StubSt()
    stub.session_state.update({
        "flight_start": date(2027, 1, 1), "flight_end": date(2027, 12, 31),
        "avails_basis": app.AVAILS_BASIS_MONTHLY, "avails_mode": True, "market_choice": "DC"})
    app.st = stub
    try:
        document = avails_pdf_import.parse_avails_pdf(str(path), path.name)
        new_groups, report = app.apply_avails_import(document)
        app._finish_avails_import(document, new_groups, report)
        groups = list(stub.session_state["targeting_groups"])
    finally:
        app.st = real
    return [g for g in groups if g.get("terms")]


def remap_ids(draft, groups, recorded_groups_path):
    """A recorded draft's row ids are the ones its own import made; a fresh
    import mints new ones -- carried across by RFPI."""
    if not recorded_groups_path.exists():
        return draft
    recorded = {g["id"]: g.get("avails_rfpi")
                for g in json.loads(recorded_groups_path.read_text("utf-8"))}
    now = {g.get("avails_rfpi"): g["id"] for g in groups}

    def fix(selection):
        if isinstance(selection, dict) and selection.get("ids"):
            selection["ids"] = [now.get(recorded.get(i), i) for i in selection["ids"]]

    fix(draft.get("group_selection"))
    for opt in draft.get("options") or []:
        fix(opt.get("group_selection"))
    return draft


def apply(app, draft, notes, avails_path, recorded_groups_path):
    groups = imported_groups(app, avails_path)
    draft = {k: v for k, v in copy.deepcopy(draft).items() if k != "_comment"}
    real = app.st
    stub = _StubSt()
    stub.session_state.update({
        "targeting_groups": groups, "market_choice": "DC",
        "flight_start": date(2027, 1, 1), "flight_end": date(2027, 12, 31),
        "avails_basis": app.AVAILS_BASIS_MONTHLY, "avails_mode": True,
        "draft_source_notes": notes, "client_name": "Burch Oil"})
    app.st = stub
    try:
        app.apply_draft_to_form(remap_ids(draft, groups, recorded_groups_path), notes=notes)
    finally:
        app.st = real
    return stub.session_state


def on_plan(state):
    return [g.get("avails_rfpi") for g in state["targeting_groups"] if g.get("include_in_plan")]


def monthly_cost(state):
    return sum(float(r["Cost"]) for r in state["plan_options"][0]["rows"])


def findings(app, state, notes):
    option = state["plan_options"][0]
    groups_by_id = {g["id"]: g for g in state["targeting_groups"]}
    totals = app.compute_plan_totals(option["rows"], option.get("breakout"), 12, "",
                                     groups_by_id=groups_by_id)
    return app.plan_sanity_findings(state["plan_options"], [totals], state["draft_plan_checks"],
                                    notes, date(2027, 1, 1), groups_by_id=groups_by_id,
                                    as_of=date(2026, 11, 19))


def main():
    needed = (NOTES_PATH, AVAILS_PATH, DRAFT_PATH, OLD_AVAILS_PATH, OLD_DRAFT_PATH)
    if not all(p.exists() for p in needed):
        print("SKIP: the real Burch Oil fixtures (tests/fixtures/burch/) aren't on this machine.")
        return
    import db
    db.fetch_audiences = lambda: (None, "stubbed for test isolation")
    import app
    notes = NOTES_PATH.read_text("utf-8")
    draft = json.loads(DRAFT_PATH.read_text("utf-8"))
    old_draft = json.loads(OLD_DRAFT_PATH.read_text("utf-8"))

    print("Burch notes + the 11-16 avails: the live draft of 2026-11-19")
    state = apply(app, draft, notes, AVAILS_PATH, FIXTURES / "_none.json")
    check("exactly the two HVAC-only rows are on the plan", on_plan(state) == HVAC_ONLY,
          on_plan(state))
    check(f"the plan prices at ${MONTHLY:,.0f} a month", abs(monthly_cost(state) - MONTHLY) < 1,
          monthly_cost(state))
    check("nothing is missing", not state["draft_plan_checks"]["missing_picks"],
          state["draft_plan_checks"]["missing_picks"])
    found = findings(app, state, notes)
    held = [f["text"] for f in found if not f.get("soft")]
    check("nothing holds Generate (no over-avails warning)", not held, held)
    growth = [f["text"] for f in found if f.get("soft") and "Growth Areas" in f["text"]]
    check("Growth Areas is about 89% of its avails, a soft note",
          len(growth) == 1 and "89%" in growth[0], [f["text"] for f in found])
    check("'Show budget on Campaign Specs' starts off",
          state.get("show_budget_on_specs") is False, state.get("show_budget_on_specs"))
    check("Budget & Allocation has no dollars", "$" not in (state.get("budget_text") or ""),
          state.get("budget_text"))
    option = state["plan_options"][0]
    caption = app.describe_group_allocation(
        app._intent_for_option(state.get("draft_plan_intent"), option),
        {g["id"]: g for g in state["targeting_groups"]}, option["rows"], option["dirty"])
    check("the totals caption names the stated counts, not max avails",
          bool(caption) and "proportion" not in caption and "200,000 impressions a month" in caption,
          caption)

    print("\nBurch notes + the 11-12 avails: the live draft of 2026-11-17")
    old = apply(app, old_draft, notes, OLD_AVAILS_PATH, OLD_GROUPS_PATH)
    check("no audience is put on the plan", on_plan(old) == [], on_plan(old))
    lines = [r.get("Tactic") for r in old["plan_options"][0]["rows"] if app.group_ids_of(r)]
    check("... and no audience line is on it", not lines, lines)
    hard = [f["text"] for f in findings(app, old, notes) if f.get("hard")]
    check("two blocking rows, one per missing RFPI",
          len(hard) == 2 and hard[0].startswith("RFPI 4969829 isn't in the uploaded avails")
          and hard[1].startswith("RFPI 4969833 isn't in the uploaded avails"), hard)
    check("... naming the RFPIs on file", all("4968157" in h and "4968164" in h for h in hard),
          hard)

    print("\nThe 2026-11-16 selection: the ZIP lists alone")
    by_area = apply(app, {**draft, "group_selection": {"mode": "named",
                                                       "match": ["Zip List #1", "Zip List #2"]}},
                    notes, AVAILS_PATH, FIXTURES / "_none.json")
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
