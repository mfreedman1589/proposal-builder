"""Draft follow-ups: every question a draft leaves open is a structured item
the rep answers with one click, and "Apply changes" applies it in code.

Checked here, end to end through the real apply_draft_to_form:
- skipping every item leaves the draft exactly as it was;
- every migrated item applies: the budget split (a model item), the
  2-segment pick, the targeting_check fixes, Great Day Washington added
  value, duplicate avails, gross-up, Total TV on/off, and avails audience
  selection (all / named / left off);
- Undo puts back everything Apply changed;
- answers persist through proposal history (Generate -> Load into form);
- no client question lands in "Before sending" (the internal list);
- every item renders as a card -- none as raw prose in a warning box.

Expected values come from the notes' own numbers and the segment names,
never from the code under test.

    python tests/test_draft_followups.py
"""
import copy
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import pandas as pd  # noqa: E402

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import draft_followups as df  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


CATALOG = pd.DataFrame([{"segment": n, "usage": u} for n, u in [
    ("DEMO Homeowner", 228), ("HH Income 150K Plus", 216), ("DEMO Age A35 Plus", 273),
    ("HH Children in Household", 137), ("HEALTH Dental Care", 40),
    ("TRAVEL Family", 50), ("TRAVEL Golf", 20),
]])
FLIGHT = (date(2026, 9, 1), date(2026, 11, 30))


class _StubSt:
    def __init__(self, state):
        self.session_state = dict(state)
        self.secrets = {}


def base_state(**extra):
    return {"market_choice": "DC", "flight_start": FLIGHT[0], "flight_end": FLIGHT[1],
            "avails_mode": False, **extra}


def apply_draft(draft, notes="", **preset):
    real_st, real_load = app.st, app.load_audience_catalog
    stub = _StubSt(base_state(**preset))
    app.st = stub
    app.load_audience_catalog = lambda: CATALOG
    try:
        app.apply_draft_to_form(copy.deepcopy(draft), notes=notes)
    finally:
        app.st, app.load_audience_catalog = real_st, real_load
    # A setup-band value the draft queued lands at the top of the next
    # render (apply_pending_draft_fields) -- before any card can be answered.
    stub.session_state.update(stub.session_state.pop("_draft_pending_fields", None) or {})
    return stub


def answer(stub, item, option_id, value=None):
    """Apply one card's answer the way "Apply changes" does."""
    real_st = app.st
    app.st = stub
    try:
        values = {item["id"]: value} if value is not None else {}
        return app.apply_review_answers(stub.session_state, {item["id"]: option_id}, values)
    finally:
        app.st = real_st


def undo(stub):
    real_st = app.st
    app.st = stub
    try:
        return app.undo_review_answers(stub.session_state)
    finally:
        app.st = real_st


def items_where(stub, text):
    return [i for i in stub.session_state.get("draft_review_items") or []
            if text.lower() in i["question"].lower()]


def rows_of(stub, option=0):
    return stub.session_state["plan_options"][option]["rows"]


def row(stub, prefix):
    return next((r for r in rows_of(stub) if r["Tactic"].endswith(prefix)), None)


PLAN_DRAFT = {
    "client_name": "QA-TEST-Brightleaf Dental", "vertical": "healthcare",
    "total_budget": 30000, "breakout": "full_flight",
    "media_plan_lines": [
        {"product": "premion_streaming_tv", "label": "Families",
         "audience_track": "Families in the area",
         "segments": ["HEALTH Dental Care", "HH Children in Household", "DEMO Age A35 Plus"],
         "allocation": {"split_evenly": True}},
        {"product": "premion_streaming_tv", "label": "Cosmetic",
         "audience_track": "Homeowners 35+ considering cosmetic work",
         "segments": ["DEMO Age A35 Plus"], "allocation": {"split_evenly": True}},
    ],
    "review_items": [
        {"type": "choice", "priority": "decision", "audience": "client",
         "question": "No split was stated between the two audiences, so the budget is divided "
                     "evenly. Which split did the client want?",
         "options": [
             {"label": "60/40 toward Families", "patch": [
                 {"op": "set_allocations", "option": "",
                  "allocations": [{"line": "Families", "allocation": {"percent_of_remainder": 60}},
                                  {"line": "Cosmetic", "allocation": {"percent_of_remainder": 40}}]}]},
             {"label": "Change the flight", "patch": [{"op": "set_flight", "months": 4}]}]},
        {"type": "choice", "audience": "client",
         "question": "Is the two-year-old spot usable, or is a new one being produced?",
         "options": []},
    ],
    "audiences": [], "attribution": [], "sports": [], "unresolved": [],
}


def check_skip_all():
    print("Skipping every item leaves the draft unchanged")
    stub = apply_draft(PLAN_DRAFT)
    before = copy.deepcopy({k: stub.session_state.get(k) for k in app.REVIEW_UNDO_KEYS})
    items = stub.session_state["draft_review_items"]
    check("the draft raised structured items", len(items) >= 3, [i["question"] for i in items])
    real_st = app.st
    app.st = stub
    try:
        summary = app.apply_review_answers(stub.session_state, {}, {})
    finally:
        app.st = real_st
    after = {k: stub.session_state.get(k) for k in app.REVIEW_UNDO_KEYS}
    check("no state an answer could touch changed", after == before)
    check("and nothing is reported applied", summary["applied"] == [], summary)
    check("every item's default is 'leave as drafted' (a None patch)",
          all(df.option_of(i, i["default"])["patch"] is None for i in items))


def check_plan_items():
    print("\nThe budget split (a model item)")
    stub = apply_draft(PLAN_DRAFT)
    split = items_where(stub, "no split was stated")
    check("the model's split question is a card, routed to the client",
          len(split) == 1 and split[0]["audience"] == "client", split)
    if split:
        labels = [o["label"] for o in split[0]["options"]]
        check("its unappliable option (a flight change) was dropped, the split kept",
              labels == ["Leave as drafted", "60/40 toward Families"], labels)
        check("drafted evenly: $15,000 each",
              [row(stub, "Families")["Cost"], row(stub, "Cosmetic")["Cost"]] == [15000, 15000],
              [r["Cost"] for r in rows_of(stub)])
        stub.session_state["plan_options"][0]["rows"].append(
            {"Tactic": "Hand-added line", "Flight": "Sep–Nov", "Geo": "DC", "Targeting": "x",
             "Impressions": 0.0, "CPM": 0.0, "Type": app.ROW_TYPE_FLAT_FEE, "Cost": 500.0})
        stub.session_state["plan_options"][0]["dirty"].append(True)
        stub.session_state["plan_options"][0]["driver"].append(app.DRIVER_COST)
        answer(stub, split[0], split[0]["options"][1]["id"])
        check("60/40 of $30,000 -> $18,000 / $12,000, re-priced in code",
              [row(stub, "Families")["Cost"], row(stub, "Cosmetic")["Cost"]] == [18000, 12000],
              [(r["Tactic"], r["Cost"]) for r in rows_of(stub)])
        check("a line added by hand since the draft is left alone",
              row(stub, "Hand-added line") is not None)
        check("the plan's per-option parallel lists stay in step",
              len(rows_of(stub)) == len(stub.session_state["plan_options"][0]["dirty"])
              == len(stub.session_state["plan_options"][0]["driver"]))
        undo(stub)
        check("Undo puts the even split back",
              [row(stub, "Families")["Cost"], row(stub, "Cosmetic")["Cost"]] == [15000, 15000])
        check("... and reopens the review", stub.session_state.get("draft_review_applied") is None)
    check("the prose question with no options is kept, for the client (shown as FYI)",
          "Is the two-year-old spot usable, or is a new one being produced?"
          in stub.session_state["draft_unresolved"], stub.session_state["draft_unresolved"])

    print("\nThe 2-segment pick")
    pick = items_where(stub, "segments combined")
    check("the Families line's 3 segments became a pick", len(pick) == 1, pick)
    if pick:
        fam = row(stub, "Families")
        check("the row carries its two kept segments (hidden field)",
              fam["_segments"] == ["HEALTH Dental Care", "HH Children in Household"], fam)
        alt = next(o for o in pick[0]["options"]
                   if o["label"] == "Use HEALTH Dental Care + DEMO Age A35 Plus")
        answer(stub, pick[0], alt["id"])
        fam = row(stub, "Families")
        check("answering swaps the row's segments",
              fam["_segments"] == ["HEALTH Dental Care", "DEMO Age A35 Plus"], fam["_segments"])
        check("... and its Targeting, which named the old pair",
              fam["Targeting"] == "HEALTH Dental Care, DEMO Age A35 Plus", fam["Targeting"])
        undo(stub)
        check("Undo restores the pair",
              row(stub, "Families")["_segments"] == ["HEALTH Dental Care", "HH Children in Household"])

    print("\nThe targeting check's fixes")
    fit = items_where(stub, "targeting says")
    cosmetic = [i for i in fit if "Cosmetic" in i["question"]]
    check("the Cosmetic line says homeowners without a homeowner segment", len(cosmetic) == 1, fit)
    if cosmetic:
        item = cosmetic[0]
        answer(stub, item, "add_segment")
        check("Add segment -> DEMO Homeowner joins the row",
              row(stub, "Cosmetic")["_segments"] == ["DEMO Age A35 Plus", "DEMO Homeowner"],
              row(stub, "Cosmetic"))
        undo(stub)
        answer(stub, item, "remove_phrase")
        check("Remove the phrase -> the description no longer promises homeowners",
              "homeowner" not in row(stub, "Cosmetic")["Targeting"].lower(),
              row(stub, "Cosmetic")["Targeting"])
        check("... and still reads as a description",
              row(stub, "Cosmetic")["Targeting"] == "35+ considering cosmetic work",
              row(stub, "Cosmetic")["Targeting"])
        undo(stub)


def check_model_item_rules():
    print("\nModel cards: segment swaps, restated drafts, repeated prose")
    swap = {"type": "choice", "audience": "internal",
            "question": "The Cosmetic line targets adults 35+. Would homeowners fit better?",
            "options": [
                {"label": "Keep it as is", "patch": []},
                {"label": "Use DEMO Homeowner", "patch": [
                    {"op": "set_segments", "option": "", "line": "Cosmetic",
                     "segments": ["DEMO Homeowner"], "audience_track": "Homeowners"}]},
                {"label": "Use a made-up segment", "patch": [
                    {"op": "set_segments", "option": "", "line": "Cosmetic",
                     "segments": ["NOT A REAL SEGMENT"]}]}]}
    draft = {**PLAN_DRAFT, "review_items": [swap],
             "unresolved_internal": ["The Cosmetic line targets adults 35+. Would homeowners "
                                     "fit better? Confirm."]}
    real_validate = app.validate_segments
    app.validate_segments = lambda names: ([n for n in names if n in set(CATALOG["segment"])],
                                           [n for n in names if n not in set(CATALOG["segment"])])
    try:
        stub = apply_draft(draft)
    finally:
        app.validate_segments = real_validate
    cards = items_where(stub, "would homeowners fit better")
    check("one card", len(cards) == 1, cards)
    if not cards:
        return
    labels = [o["label"] for o in cards[0]["options"]]
    check("the restated draft and the unknown segment are dropped; the real swap kept",
          labels == ["Leave as drafted", "Use DEMO Homeowner"], labels)
    check("the same question isn't also left in the seller's list",
          not any("would homeowners" in n.lower()
                  for n in stub.session_state["draft_unresolved_internal"]),
          stub.session_state["draft_unresolved_internal"])
    answer(stub, cards[0], cards[0]["options"][1]["id"])
    cosmetic = row(stub, "Cosmetic")
    check("answering swaps the line's segments and its description",
          cosmetic["_segments"] == ["DEMO Homeowner"] and cosmetic["Targeting"] == "Homeowners",
          cosmetic)


def check_gross_and_total_tv():
    print("\nGross-up and Total TV")
    draft = {**PLAN_DRAFT, "review_items": [], "total_budget": 34500,
             "total_budget_basis": "gross", "total_tv": True,
             "media_plan_lines": [{"product": "premion_streaming_tv",
                                   "allocation": {"percent_of_total": 100}}]}
    stub = apply_draft(draft, notes="Total TV with a WUSA9 broadcast schedule.")
    gross = items_where(stub, "gross budget")
    check("a gross budget raises a gross-up card", len(gross) == 1, gross)
    check("priced net: $30,000", round(sum(r["Cost"] for r in rows_of(stub))) == 30000,
          rows_of(stub))
    if gross:
        answer(stub, gross[0], "net")
        check("'It's a net budget' -> lines priced at the stated $34,500",
              round(sum(r["Cost"] for r in rows_of(stub))) == 34500, rows_of(stub))
        check("... and the gross-up switched off", stub.session_state["agency_gross_up"] is False)
        undo(stub)
        check("Undo -> back to $30,000 net with gross-up on",
              round(sum(r["Cost"] for r in rows_of(stub))) == 30000
              and stub.session_state["agency_gross_up"] is True)
    tv = items_where(stub, "Total TV was switched on")
    check("Total TV switched on raises a card", len(tv) == 1)
    if tv:
        answer(stub, tv[0], "off")
        check("'Turn Total TV off' -> off", stub.session_state["total_tv"] is False)
        undo(stub)
        check("Undo -> on again", stub.session_state["total_tv"] is True)


def check_gdw():
    print("\nGreat Day Washington")
    draft = {**PLAN_DRAFT, "review_items": []}
    stub = apply_draft(draft, notes="Dental practice, $30k. Include GDW.")
    gdw = [i for i in stub.session_state["draft_review_items"]
           if any(o["id"] == "added_value" for o in i["options"])]
    check("GDW at its paid default raises the added-value card", len(gdw) == 1)
    if gdw:
        answer(stub, gdw[0], "added_value")
        check("answering queues every GDW line to Added Value",
              stub.session_state.get("_pending_gdw") == {"added_value": True})
        app.apply_gdw_update(stub.session_state["plan_options"], added_value=True)
        rows = [r for r in rows_of(stub) if app.is_gdw_row(r)]
        check("... which makes the line Added Value",
              rows and rows[0]["Type"] == app.ROW_TYPE_ADDED_VALUE, rows)


def groups_fixture():
    a = tg.new_group(["TRAVEL Family"], name="Family", avails_monthly=100000, group_id="g-a")
    b = tg.new_group(["TRAVEL Golf"], name="Golf", avails_monthly=50000, group_id="g-b")
    dup = tg.new_group(["TRAVEL Golf"], name="Golf", avails_monthly=50000, group_id="g-dup")
    return [a, b, dup]


def check_avails_selection():
    print("\nAvails audience selection and duplicates")
    draft = {**PLAN_DRAFT, "review_items": [], "media_plan_lines": [],
             "group_selection": {"mode": "all"},
             "group_allocation": {"proportional_to_avails": True}}
    stub = apply_draft(draft, targeting_groups=groups_fixture(), avails_mode=True)
    every = items_where(stub, "didn't say which")
    check("'all of them are on the plan' is a card for the seller",
          len(every) == 1 and every[0]["audience"] == "internal", every)
    dups = items_where(stub, "duplicate")
    check("the duplicate avails line is a card", len(dups) == 1, dups)
    groups = {g["id"]: g for g in stub.session_state["targeting_groups"]}
    check("drafted: the duplicate is off the plan, the others on",
          groups["g-dup"]["include_in_plan"] is False and groups["g-a"]["include_in_plan"]
          and groups["g-b"]["include_in_plan"], groups)
    if dups:
        answer(stub, dups[0], "count")
        groups = {g["id"]: g for g in stub.session_state["targeting_groups"]}
        check("'count them' -> duplicates counted and the line on the plan",
              stub.session_state["avails_include_duplicates"] is True
              and groups["g-dup"]["include_in_plan"] is True, groups["g-dup"])
        check("... and the plan is re-priced next render",
              stub.session_state.get("_pending_group_realloc") is True)
        undo(stub)
        groups = {g["id"]: g for g in stub.session_state["targeting_groups"]}
        check("Undo -> duplicate off again", groups["g-dup"]["include_in_plan"] is False
              and not stub.session_state.get("avails_include_duplicates"))
    if every:
        answer(stub, every[0], "change", value=["g-a"])
        groups = {g["id"]: g for g in stub.session_state["targeting_groups"]}
        check("'Change which audiences' with Family only -> Golf comes off",
              groups["g-a"]["include_in_plan"] and not groups["g-b"]["include_in_plan"], groups)
        undo(stub)

    named = {**draft, "group_selection": {"mode": "named", "match": ["Family"]},
             "group_selection_reason": "The notes sell the family audience."}
    stub = apply_draft(named, targeting_groups=groups_fixture(), avails_mode=True)
    confirm = items_where(stub, "family audience")
    check("a named selection is a card for the CLIENT",
          len(confirm) == 1 and confirm[0]["audience"] == "client", confirm)
    left = items_where(stub, "available but not on the plan")
    check("Golf (not the duplicate) is suggested as available but not on the plan",
          len(left) == 1 and "TRAVEL Golf" in left[0]["question"], left)
    if left:
        answer(stub, left[0], "add")
        groups = {g["id"]: g for g in stub.session_state["targeting_groups"]}
        check("'Add to the plan' -> Golf on", groups["g-b"]["include_in_plan"] is True)


def check_client_routing():
    print("\nNo client question lands in 'Before sending'")
    draft = {**PLAN_DRAFT, "review_items": [], "total_budget": 10000,
             "audiences": [{"segment": "TRAVEL Family", "max_avails": 1000000,
                            "avails_basis": "monthly"}],
             "media_plan_lines": [{"product": "premion_streaming_tv",
                                   "audience_track": "TRAVEL Family",
                                   "allocation": {"percent_of_avails": 50}}]}
    stub = apply_draft(draft, avails_mode=True)
    client = stub.session_state["draft_unresolved"]
    internal = stub.session_state["draft_unresolved_internal"]
    check("reach vs. budget ('confirm which one the client agreed to') is a client question",
          any("client agreed" in n for n in client), client)
    check("... and not in the seller's list", not any("client agreed" in n for n in internal),
          internal)
    stubs = [stub, apply_draft(PLAN_DRAFT),
             apply_draft({**PLAN_DRAFT, "total_budget_basis": "gross", "total_tv": True})]
    leaks = [n for s in stubs for n in s.session_state["draft_unresolved_internal"]
             if any(w in n.lower() for w in ("client agreed", "ask the client",
                                             "confirm with the client"))]
    check("no internal note anywhere asks the client something", not leaks, leaks)
    leaks = [i["question"] for s in stubs for i in s.session_state["draft_review_items"]
             if i["audience"] == "internal" and "client agreed" in i["question"].lower()]
    check("no internal card asks the client something either", not leaks, leaks)


def check_cards_and_history():
    print("\nThrough the real page: cards, Apply, history, reopen")
    from streamlit.testing.v1 import AppTest
    stub = apply_draft(PLAN_DRAFT)
    state = {k: v for k, v in stub.session_state.items()
             if not k.startswith("_draft_pending") and k != "_product_seed_key"}
    state.update(base_state(client_name="QA-TEST-Brightleaf Dental"))
    captured = {}
    real_log = db.log_proposal

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        captured["row"] = {"id": "00000000-0000-0000-0000-000000000000",
                           "client_name": client_name, "vertical": vertical, "market": market,
                           "form_json": form_json, "deck_version_id": kwargs.get("deck_version_id"),
                           "logo_storage_path": None}
        return captured["row"]["id"], None

    db.log_proposal = spy_log
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Regression Suite"
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        check("renders", not at.exception, at.exception)
        if at.exception:
            return
        questions = [i["question"] for i in at.session_state["draft_review_items"]]
        warned = " ".join(w.value for w in at.warning)
        check("no review item is printed as raw prose in a warning box",
              not any(q[:40] in warned for q in questions), warned[:300])
        cards = [r for r in at.radio if str(r.key or "").startswith("rv_")]
        decisions = [i for i in at.session_state["draft_review_items"] if i["type"] != "fyi"]
        check("one card per decision item", len(cards) == len(decisions),
              (len(cards), len(decisions)))
        md = " ".join(m.value for m in at.markdown)
        check("the three group headings render",
              "Needs a decision" in md and "FYI" in md, md[:200])
        check("the prose question renders as an FYI card with Dismiss",
              any(b.label == "Dismiss" for b in at.button))
        split = next(r for r in cards if "60/40 toward Families" in r.options)
        split_item = items_where(stub, "no split was stated")[0]
        split.set_value(split_item["options"][1]["id"]).run()
        apply = [b for b in at.button if b.label == "Apply 1 change"]
        check("the Apply button counts the change", len(apply) == 1,
              [b.label for b in at.button if "pply" in b.label])
        if not apply:
            return
        apply[0].click().run()
        check("applies without raising", not at.exception, at.exception)
        costs = sorted(round(r["Cost"]) for r in at.session_state["plan_options"][0]["rows"])
        check("the plan is 60/40 on the page", costs == [12000, 18000], costs)
        check("the cards collapse to '1 change applied' with Undo",
              any("1 change applied" in s.value for s in at.success)
              and any(b.label == "Undo" for b in at.button))
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if generate:
            generate[0].click().run()
    finally:
        db.log_proposal = real_log
    row_logged = captured.get("row")
    check("Generate logged the proposal", row_logged is not None)
    if not row_logged:
        return
    record = row_logged["form_json"].get("draft") or {}
    check("history carries the items, answers, what was applied and the seller's list",
          record.get("review_items") and record.get("review_answers")
          and (record.get("review_applied") or {}).get("applied")
          and "unresolved_internal" in record, sorted(record))
    logged_rows = row_logged["form_json"]["plan_options"][0]["rows"]
    check("logged rows carry their segments",
          any(r.get("_segments") for r in logged_rows), logged_rows)

    reopened = _StubSt({})
    real_st, real_logo = app.st, db.proposal_logo
    app.st = reopened
    db.proposal_logo = lambda path: None
    try:
        app.rehydrate_proposal_into_form(row_logged, parent_proposal_id=row_logged["id"])
    finally:
        app.st, db.proposal_logo = real_st, real_logo
    s = reopened.session_state
    check("reopened: the answers are back",
          s.get("draft_review_answers") == record["review_answers"])
    check("reopened: '1 change applied' is back",
          len((s.get("draft_review_applied") or {}).get("applied") or []) == 1)
    check("reopened: the option keeps its drafted lines, rows their segments",
          s["plan_options"][0].get("_draft") and any(r.get("_segments")
                                                     for r in s["plan_options"][0]["rows"]))
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Regression Suite"
    for key, value in s.items():
        at.session_state[key] = value
    at.run()
    check("the reopened page renders", not at.exception, at.exception)
    check("... showing the summary, not the cards again",
          any("1 change applied" in x.value for x in at.success)
          and not [r for r in at.radio if str(r.key or "").startswith("rv_")])


def main():
    check_skip_all()
    check_plan_items()
    check_model_item_rules()
    check_gross_and_total_tv()
    check_gdw()
    check_avails_selection()
    check_client_routing()
    check_cards_and_history()
    print(f"\n{'ALL PASS' if not failures else f'{len(failures)} FAILED'}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
