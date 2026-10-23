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
from datetime import date, timedelta
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


def apply_draft(draft, notes="", catalog=None, **preset):
    real_st, real_load = app.st, app.load_audience_catalog
    stub = _StubSt(base_state(**preset))
    app.st = stub
    app.load_audience_catalog = lambda: CATALOG if catalog is None else catalog
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


def option_id(item, label):
    return next(o["id"] for o in item["options"] if o["label"] == label)


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
        check("the model's split card becomes the app's: even split drafted, 60/40 and 70/30 "
              "each way (its unappliable flight-change option gone)",
              labels == ["Keep 50/50 between Families and Cosmetic", "70/30 toward Families",
                         "60/40 toward Families", "60/40 toward Cosmetic", "70/30 toward Cosmetic"],
              labels)
        check("drafted evenly: $15,000 each",
              [row(stub, "Families")["Cost"], row(stub, "Cosmetic")["Cost"]] == [15000, 15000],
              [r["Cost"] for r in rows_of(stub)])
        stub.session_state["plan_options"][0]["rows"].append(
            {"Tactic": "Hand-added line", "Flight": "Sep–Nov", "Geo": "DC", "Targeting": "x",
             "Impressions": 0.0, "CPM": 0.0, "Type": app.ROW_TYPE_FLAT_FEE, "Cost": 500.0})
        stub.session_state["plan_options"][0]["dirty"].append(True)
        stub.session_state["plan_options"][0]["driver"].append(app.DRIVER_COST)
        answer(stub, split[0], option_id(split[0], "60/40 toward Families"))
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


def check_phase_15_items():
    print("\nPhase 1.5: negotiated CPM, missing avails, ambiguous markets, prose splits")
    rate_card = app.line_product_spec("premion_streaming_tv")[1]
    one_line = {**PLAN_DRAFT, "review_items": [],
                "media_plan_lines": [{"product": "premion_streaming_tv", "label": "Families",
                                      "audience_track": "Families", "cpm": 28,
                                      "allocation": {"percent_of_total": 100}}]}
    stub = apply_draft(one_line)
    cpm = items_where(stub, "rate card rate")
    check("a negotiated CPM is a card, not prose", len(cpm) == 1 and not any(
        "rate card rate" in n for n in stub.session_state["draft_unresolved_internal"]), cpm)
    if cpm:
        cost = row(stub, "Families")["Cost"]
        answer(stub, cpm[0], "rate_card")
        fam = row(stub, "Families")
        check(f"'Use the rate card' re-prices the line at ${rate_card:,.2f}, same budget",
              fam["CPM"] == rate_card and fam["Cost"] == cost
              and round(fam["Impressions"]) == round(cost / rate_card * 1000), fam)
    equal = {**one_line, "media_plan_lines": [dict(one_line["media_plan_lines"][0], cpm=rate_card)]}
    stub = apply_draft(equal)
    check("a drafted CPM equal to the rate card raises nothing (was '$32.00, not $32.00')",
          not items_where(stub, "rate card rate") and not any(
              "rate card rate" in n for n in stub.session_state["draft_unresolved_internal"]))

    avails = {**PLAN_DRAFT, "review_items": [],
              "audiences": [{"segment": "TRAVEL Family"}],
              "media_plan_lines": [{"product": "premion_streaming_tv",
                                    "allocation": {"percent_of_total": 100}}]}
    stub = apply_draft(avails, avails_mode=True)
    missing = items_where(stub, "No avails were given for TRAVEL Family")
    check("an audience with no avails is a missing-value card", len(missing) == 1
          and missing[0]["type"] == "missing_value", missing)
    if missing:
        answer(stub, missing[0], "enter", value=750000)
        groups = [g for g in stub.session_state["targeting_groups"]
                  if tg.audience_label(g) == "TRAVEL Family"]
        check("entering 750,000 sets the group's monthly avails",
              groups and groups[0]["avails_monthly"] == 750000, groups)
        rows = stub.session_state["avails_seed_rows"]
        check("... and the flat avails rows agree, so the next sync keeps it",
              any(750000 in r.values() for r in rows), rows)
        undo(stub)
        groups = [g for g in stub.session_state["targeting_groups"]
                  if tg.audience_label(g) == "TRAVEL Family"]
        check("Undo -> back to 0", groups and not groups[0]["avails_monthly"], groups)

    markets = {**PLAN_DRAFT, "review_items": [], "target_markets": ["Harrisburg"]}
    stub = apply_draft(markets)
    ambiguous = items_where(stub, "could be more than one market")
    check("an ambiguous market name is a choice card", len(ambiguous) == 1
          and len(ambiguous[0]["options"]) >= 3, ambiguous)
    if ambiguous:
        pick = ambiguous[0]["options"][1]
        answer(stub, ambiguous[0], pick["id"])
        check(f"picking {pick['label']} adds it to Target DMAs",
              pick["patch"][0]["label"] in (stub.session_state.get("target_dma_choice") or []),
              stub.session_state.get("target_dma_choice"))

    prose_split = {**PLAN_DRAFT, "review_items": [],
                   "unresolved_internal": ["No split was given between the Families and "
                                           "Cosmetic audiences, so they're weighted evenly. "
                                           "Confirm the weighting."]}
    stub = apply_draft(prose_split)
    split = items_where(stub, "weighted evenly")
    check("a split the model only described in prose becomes a card", len(split) == 1
          and not any("weighted evenly" in n
                      for n in stub.session_state["draft_unresolved_internal"]), split)
    if split:
        labels = [o["label"] for o in split[0]["options"]]
        check("it offers the weightings nearest the drafted 50/50, each stated",
              labels[0] == "Keep 50/50 between Families and Cosmetic"
              and "60/40 toward Families" in labels, labels)
        sixty = next(o for o in split[0]["options"] if o["label"] == "60/40 toward Families")
        answer(stub, split[0], sixty["id"])
        check("60/40 -> $18,000 / $12,000",
              [row(stub, "Families")["Cost"], row(stub, "Cosmetic")["Cost"]] == [18000, 12000],
              [(r["Tactic"], r["Cost"]) for r in rows_of(stub)])


def check_proxy_and_duplicate_prose():
    print("\nStand-in segments and duplicate avails notes")
    draft = {**PLAN_DRAFT, "review_items": [],
             "unresolved_internal": [
                 "No catalog segment matches cosmetic-dentistry patients, so DEMO Age A35 Plus "
                 "was used as the closest proxy on the Cosmetic line.",
                 "Pull real monthly avails for TRAVEL Family before sending."],
             "audiences": [{"segment": "TRAVEL Family"}]}
    real = app.segment_recommend.recommend
    app.segment_recommend.recommend = lambda catalog, text, vertical_key=None: {
        "core": [{"segment": "HEALTH Dental Care"}, {"segment": "DEMO Age A35 Plus"}],
        "modifiers": [], "custom": [], "tags": ["Health"], "primary": "Health"}
    try:
        # An avails file attached, so the audience-alternatives card (which
        # would swap the core first) doesn't apply -- this is the stand-in
        # note's own path.
        stub = apply_draft(draft, avails_mode=True, avails_documents=[{"name": "RFPID-1.pdf"}])
    finally:
        app.segment_recommend.recommend = real
    internal = stub.session_state["draft_unresolved_internal"]
    proxy = items_where(stub, "closest proxy")
    check("a stand-in segment note becomes a swap card", len(proxy) == 1
          and not any("closest proxy" in n for n in internal), (proxy, internal))
    if proxy:
        labels = [o["label"] for o in proxy[0]["options"]]
        check("it offers the recommender's segments, not the one already on the line",
              labels == ["Keep DEMO Age A35 Plus", "Use HEALTH Dental Care"], labels)
        answer(stub, proxy[0], "alt1")
        check("answering swaps the stand-in on the row",
              row(stub, "Cosmetic")["_segments"] == ["HEALTH Dental Care"], row(stub, "Cosmetic"))
    check("a 'pull avails' note for an audience that already has a missing-avails card is dropped",
          not any("Pull real monthly avails" in n for n in internal)
          and items_where(stub, "No avails were given for TRAVEL Family"), internal)


def check_reach_notes_collapse():
    print("\nPer-line reach notes fold into one card")
    notes = [f"The Premion Streaming TV line's ${c:,} reaches about {i:,} impressions -- "
             f"roughly {p}% of the 1,011,108 avails on TRAVEL Family for the flight."
             for c, i, p in ((4300, 134375, 13), (652, 20375, 2), (113, 3531, 0))]
    out = app.collapse_reach_notes(["Something else."] + notes)
    check("12 rows' worth of reach notes become one note (here 3 -> 1)", len(out) == 2, out)
    check("... listing every line", all(f"{p}% of 1,011,108" in out[1] for p in (13, 2, 0)), out)
    check("a single reach note is left as it is", app.collapse_reach_notes(notes[:1]) == notes[:1])


def check_reprice_keeps_monthly_budget():
    print("\nRe-pricing a Monthly option keeps its budget (found on the Hershey walkthrough)")
    for breakout, per_row_total in (("monthly", 10000), ("full_flight", 30000)):
        draft = {**PLAN_DRAFT, "review_items": [], "media_plan_lines": [], "breakout": breakout,
                 "group_selection": {"mode": "all"},
                 "group_allocation": {"proportional_to_avails": True}}
        stub = apply_draft(draft, targeting_groups=groups_fixture(), avails_mode=True)
        option = stub.session_state["plan_options"][0]
        before = sum(r["Cost"] for r in option["rows"])
        groups = {g["id"]: g for g in stub.session_state["targeting_groups"]}
        app._allocate_group_rows(option, app._intent_for_option(
            stub.session_state["draft_plan_intent"], option), groups, "Sep-Nov", "x", 3)
        after = sum(r["Cost"] for r in option["rows"])
        check(f"{breakout}: the drafted $30,000 shows as ${per_row_total:,} on the grid, before "
              f"and after a re-price", round(before) == round(after) == per_row_total,
              (before, after))


def check_one_card_across_options():
    print("\nThe same question on two plan options is one card")
    line = {"product": "premion_streaming_tv", "label": "Families", "audience_track": "Families",
            "cpm": 28, "allocation": {"percent_of_total": 100}}
    draft = {**PLAN_DRAFT, "review_items": [], "media_plan_lines": [], "total_budget": 0,
             "options": [{"name": "$50K Plan", "total_budget": 50000, "breakout": "full_flight",
                          "media_plan_lines": [line]},
                         {"name": "$75K Plan", "total_budget": 75000, "breakout": "full_flight",
                          "media_plan_lines": [line]}]}
    stub = apply_draft(draft)
    cpm = items_where(stub, "rate card rate")
    check("one negotiated-CPM card, not one per option", len(cpm) == 1, cpm)
    if cpm:
        answer(stub, cpm[0], "rate_card")
        rate_card = app.line_product_spec("premion_streaming_tv")[1]
        cpms = [r["CPM"] for o in stub.session_state["plan_options"] for r in o["rows"]]
        check("answering it puts both options back on the rate card", cpms == [rate_card] * 2,
              cpms)


SENIOR_CATALOG = pd.DataFrame([
    {"segment": seg, "tags": "Healthcare", "role": role, "status": "standard", "score": score,
     "line_count": lines, "category": "Healthcare", "visible": True}
    for seg, role, score, lines in [
        ("DEMO Senior Living Care", "core", 9.0, 23), ("HLTH Insurance Medicare", "core", 7.0, 14),
        ("LIFESTYLE Health Wellness", "core", 6.0, 20), ("HLTH Private Insurance", "core", 5.0, 25),
        ("DEMO Age A55 Plus", "modifier", 8.0, 60)]])
SENIOR_NOTES = ("QA-TEST call with Maple Grove, a senior living community. They want to reach "
                "seniors 55+ and adult children researching care. $20,000 over the flight.")
SENIOR_DRAFT = {**PLAN_DRAFT, "review_items": [], "total_budget": 20000,
                "media_plan_lines": [{"product": "premion_streaming_tv", "label": "Seniors",
                                      "audience_track": "Seniors 55+ researching senior living",
                                      "segments": ["HLTH Insurance Medicare", "DEMO Age A55 Plus"],
                                      "allocation": {"percent_of_total": 100}}]}


def apply_senior(draft=SENIOR_DRAFT, notes=SENIOR_NOTES, **preset):
    return apply_draft(draft, notes=notes, catalog=SENIOR_CATALOG, **preset)


def check_audience_alternatives():
    print("\nAudience alternatives (no avails file, notes name no segment)")
    stub = apply_senior()
    seniors = row(stub, "Seniors")
    check("the plan takes the recommender's primary for the core, keeping the modifier",
          seniors["_segments"] == ["DEMO Senior Living Care", "DEMO Age A55 Plus"], seniors)
    cards = items_where(stub, "use an alternative")
    check("one Suggestion card for the line", len(cards) == 1
          and cards[0]["type"] == "suggestion", cards)
    if cards:
        labels = [o["label"] for o in cards[0]["options"]]
        check("Keep primary is the default, then 2-3 alternatives, each with a why",
              labels[0].startswith("Keep primary: DEMO Senior Living Care (#1 most booked")
              and 2 <= len(labels) - 1 <= 3 and all("most booked in Healthcare" in l for l in labels),
              labels)
        total = sum(r["Cost"] for r in rows_of(stub))
        alt = cards[0]["options"][1]
        answer(stub, cards[0], alt["id"])
        seniors = row(stub, "Seniors")
        check("swapping changes the line's core segment and keeps the modifier",
              seniors["_segments"] == [alt["label"].split(" (")[0], "DEMO Age A55 Plus"], seniors)
        check("... and the totals", sum(r["Cost"] for r in rows_of(stub)) == total)
    stub = apply_senior(avails_documents=[{"name": "RFPID-1.pdf"}])
    check("an avails file attached -> no card, segments as drafted",
          not items_where(stub, "use an alternative") and row(stub, "Seniors")["_segments"]
          == ["HLTH Insurance Medicare", "DEMO Age A55 Plus"])
    proxy_note = {**SENIOR_DRAFT, "unresolved_internal": [
        "No catalog segment matches senior-living prospects exactly, so HLTH Insurance Medicare "
        "was used as the closest proxy."]}
    stub = apply_senior(draft=proxy_note)
    check("a model note about the stand-in the card just replaced is dropped (the card covers it)",
          not any("closest proxy" in n for n in stub.session_state["draft_unresolved_internal"])
          and items_where(stub, "use an alternative"), stub.session_state["draft_unresolved_internal"])
    stub = apply_senior(notes=SENIOR_NOTES + " Use HLTH Insurance Medicare for the core.")
    check("notes naming the segment -> no card, segment kept",
          not items_where(stub, "use an alternative") and row(stub, "Seniors")["_segments"][0]
          == "HLTH Insurance Medicare")


def check_unexpected_geographies():
    print("\nUnexpected geographies on the avails table")
    groups = groups_fixture()
    addon = tg.new_group(["TRAVEL Golf"], name="Philly Add-On Zip", avails_monthly=8000,
                         group_id="g-addon",
                         geo_def={"kind": "zips", "zips": ["19103"], "label": "Philly Add-On Zip"})
    draft = {**PLAN_DRAFT, "review_items": [], "media_plan_lines": [],
             "group_selection": {"mode": "all"},
             "group_allocation": {"proportional_to_avails": True}}
    stub = apply_draft(draft, notes="Travel campaign for families and golfers.",
                       targeting_groups=groups + [addon], avails_mode=True)
    cards = items_where(stub, "Philly Add-On Zip")
    check("a ZIP-list geography the notes don't mention is a Keep / Remove card",
          len(cards) == 1 and [o["label"] for o in cards[0]["options"]]
          == ["Keep it on the plan", "Remove it from the plan"], cards)
    if cards:
        answer(stub, cards[0], "remove")
        addon_after = next(g for g in stub.session_state["targeting_groups"] if g["id"] == "g-addon")
        check("Remove takes it off the plan", addon_after["include_in_plan"] is False, addon_after)
    stub = apply_draft(draft, notes="Travel campaign, include the Philly Add-On Zip list.",
                       targeting_groups=groups_fixture() + [addon], avails_mode=True)
    check("a geography the notes name raises nothing", not items_where(stub, "Philly Add-On Zip"))


def check_client_name_prefix():
    print("\nClient name: the QA-TEST- prefix isn't a mismatch")
    stub = apply_draft({**PLAN_DRAFT, "review_items": [], "client_name": "Brightleaf Dental"},
                       client_name="QA-TEST-Brightleaf Dental")
    check("no note when only the QA-TEST- prefix differs",
          not any("Client name" in n for n in stub.session_state["draft_unresolved_internal"]),
          stub.session_state["draft_unresolved_internal"])
    stub = apply_draft({**PLAN_DRAFT, "review_items": [], "client_name": "Brightleaf Ortho"},
                       client_name="QA-TEST-Brightleaf Dental")
    check("a real difference is still noted",
          any("Client name" in n for n in stub.session_state["draft_unresolved_internal"]))


def check_split_labels():
    print("\nSplit choices state the split")
    vague = {**PLAN_DRAFT, "review_items": [{
        "type": "choice", "audience": "client", "question": "Which split did the client want?",
        "options": [{"label": "More toward Families", "patch": [{
            "op": "set_allocations", "option": "",
            "allocations": [{"line": "Families", "allocation": {"percent_of_remainder": 70}},
                            {"line": "Cosmetic", "allocation": {"percent_of_remainder": 30}}]}]}]}]}
    stub = apply_draft(vague)
    card = items_where(stub, "Which split did the client want")
    labels = [o["label"] for o in card[0]["options"]] if card else []
    check("'More toward Families' is relabelled from its own patch: '70/30 toward Families'",
          "70/30 toward Families" in labels and not any("More toward" in l for l in labels), labels)

    draft = {**PLAN_DRAFT, "media_plan_lines": [], "group_selection": {"mode": "all"},
             "group_allocation": {"proportional_to_avails": True},
             "review_items": [{"type": "choice", "audience": "client",
                               "question": "No weighting was given between the audiences.",
                               "options": [{"label": "More toward Family", "patch": [{
                                   "op": "set_group_allocation", "option": "",
                                   "group_allocation": {"percent_of_remainder": 60}}]}]}]}
    groups = [g for g in groups_fixture() if g["id"] != "g-dup"]
    stub = apply_draft(draft, targeting_groups=groups, avails_mode=True)
    card = items_where(stub, "No weighting was given")
    labels = [o["label"] for o in card[0]["options"]] if card else []
    check("on an avails-table plan the split is restated by audience, numbers first",
          card and labels[0].startswith("Keep ") and all(
              "/" in l and "toward" in l or "between" in l for l in labels[1:])
          and not any("More toward" in l for l in labels), labels)
    pick = next((o for o in card[0]["options"] if o["label"] == "60/40 toward TRAVEL Family"), None) if card else None
    check("... offering '60/40 toward TRAVEL Family'", pick is not None, labels)
    if pick:
        answer(stub, card[0], pick["id"])
        option = stub.session_state["plan_options"][0]
        intent = app._intent_for_option(stub.session_state["draft_plan_intent"], option)
        groups_by_id = {g["id"]: g for g in stub.session_state["targeting_groups"]}
        app._allocate_group_rows(option, intent, groups_by_id, "Sep-Nov", "x", 3)
        by_aud = {}
        for r in option["rows"]:
            for gid in app.group_ids_of(r):
                by_aud[tg.audience_label(groups_by_id[gid])] = by_aud.get(
                    tg.audience_label(groups_by_id[gid]), 0) + r["Cost"]
        total = sum(by_aud.values())
        check("applying it re-prices 60% to TRAVEL Family, 40% to TRAVEL Golf, same total",
              round(by_aud.get("TRAVEL Family", 0) / total, 2) == 0.60 and round(total) == 30000,
              by_aud)


def check_plain_words():
    print("\nNo internal names on a card or note")
    draft = {**PLAN_DRAFT, "review_items": [], "unresolved_internal": [
        "AFIRST Travel Buffs is flagged rfp_selectable true -- confirm it can be booked.",
        "The streaming_retargeting_display line uses percent_of_remainder; confirm the weighting "
        "with the media_plan_lines owner."]}
    stub = apply_draft(draft)
    notes = stub.session_state["draft_unresolved_internal"] + stub.session_state["draft_unresolved"]
    texts = notes + [i["question"] for i in stub.session_state["draft_review_items"]] + [
        o["label"] for i in stub.session_state["draft_review_items"] for o in i["options"]]
    check("a note about RFP-selectable status is dropped",
          not any("rfp" in t.lower() for t in texts), texts)
    check("no snake_case identifier reaches the page",
          not any(__import__("re").search(r"\b[a-z]+_[a-z0-9_]+\b", t) for t in texts),
          [t for t in texts if "_" in t])
    check("a product key reads as the product",
          any("Streaming Retargeting" in t and "share of the remaining budget" in t for t in texts),
          texts)


def check_apply_button_types():
    print("\nApply is primary only once something is picked")
    from streamlit.testing.v1 import AppTest
    stub = apply_draft(PLAN_DRAFT)
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Regression Suite"
    for key, value in {**stub.session_state, **base_state(client_name="QA-TEST-x")}.items():
        at.session_state[key] = value
    at.run()
    keep = [b for b in at.button if b.label == "Keep everything as drafted"]
    check("nothing picked: 'Keep everything as drafted', secondary",
          keep and keep[0].proto.type == "secondary", [(b.label, b.proto.type) for b in keep])
    radio = next(r for r in at.radio if str(r.key or "").startswith("rv_")
                 and "60/40 toward Families" in r.options)
    split = items_where(stub, "no split was stated")[0]
    radio.set_value(split["options"][1]["id"]).run()
    apply = [b for b in at.button if b.label == "Apply 1 change"]
    check("one picked: 'Apply 1 change', primary",
          apply and apply[0].proto.type == "primary", [(b.label, b.proto.type) for b in apply])


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

    reweight = df.make_item(
        ("test", "reweight"), "choice", "decision", "client", "Which weighting?",
        [{"id": "even", "label": "Split evenly",
          "patch": {"op": "set_group_allocation", "option": stub.session_state["plan_options"][0]["name"],
                    "group_allocation": {"split_evenly": True}}}])
    stub.session_state["draft_review_items"].append(reweight)  # as a model card would be
    answer(stub, reweight, "even")
    intent = (stub.session_state.get("draft_plan_intent") or {}).get("options") or [{}]
    check("a group re-weighting card sets the option's stored allocation and queues a re-price",
          intent[0].get("group_allocation") == {"split_evenly": True}
          and stub.session_state.get("_pending_group_realloc") is True, intent)
    undo(stub)
    intent = (stub.session_state.get("draft_plan_intent") or {}).get("options") or [{}]
    check("Undo puts the drafted allocation back",
          intent[0].get("group_allocation") == {"proportional_to_avails": True}, intent)

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
        split.set_value(option_id(split_item, "60/40 toward Families")).run()
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


# --- Muse QA on 2fa3a2d: QA-022 .. QA-027 and the items filed with them -------

HVAC_NOTES = ("HVAC company in Springfield, target RETAIL Home Services HVAC and homeowners, "
              "$10,000 a month for 2 months.")
HVAC_CUSTOM = "CUSTOM HH Home Built Before 2014 AND RETAIL Home Services HVAC"
HVAC_CATALOG = pd.DataFrame([
    {"segment": seg, "tags": "Home Services", "role": role, "status": status, "score": score,
     "line_count": lines, "category": "RETAIL", "visible": True}
    for seg, role, status, score, lines in [
        (HVAC_CUSTOM, "core", "custom", 9.5, 30),
        ("RETAIL Home Services HVAC", "core", "standard", 9.0, 40),
        ("LIFESTAGE Homeowners 5 Years Plus", "core", "standard", 6.0, 20),
        ("FIN Insurance Homeowners", "core", "standard", 5.0, 12),
        ("HH Home Built Before 2014", "core", "standard", 4.0, 10),
        ("DEMO Homeowner", "modifier", "standard", 8.0, 60),
        ("CUSTOM DEMO Homeowner AND HH Income 150K Plus", "core", "custom", 2.0, 5)]])
HVAC_DRAFT = {
    "client_name": "QA-TEST-Springfield HVAC", "vertical": "home_improvement",
    "total_budget": 10000, "total_budget_period": "month", "stated_flight": {"months": 2},
    "breakout": "monthly",
    "media_plan_lines": [{"product": "premion_streaming_tv",
                          "audience_track": "Homeowners in-market for HVAC services, with a focus "
                                            "on homes built before 2014",
                          "segments": [HVAC_CUSTOM, "DEMO Homeowner"],
                          "allocation": {"percent_of_total": 100}}],
    "campaign_specs": {"audience": ["Homeowners in-market for HVAC services, with a focus on "
                                    "homes built before 2014"]},
    "audiences": [{"segment": HVAC_CUSTOM}], "attribution": [], "sports": [],
    "review_items": [], "unresolved": [],
    "unresolved_internal": ["No client name was filled in, so the notes' name was used."]}
TWO_MONTHS = (date(2026, 11, 1), date(2026, 12, 31))


def apply_hvac(draft=HVAC_DRAFT, notes=HVAC_NOTES, **preset):
    real, real_validate = app.load_audience_catalog, app.validate_segments
    names = set(HVAC_CATALOG["segment"])
    app.load_audience_catalog = lambda: HVAC_CATALOG
    app.validate_segments = lambda given: ([n for n in given if n in names],
                                           [n for n in given if n not in names])
    try:
        return apply_draft(draft, notes=notes, catalog=HVAC_CATALOG, **preset)
    finally:
        app.load_audience_catalog, app.validate_segments = real, real_validate


def option_totals(stub, months):
    """(monthly, full flight) as Section E shows them for a Monthly option."""
    monthly = sum(r["Cost"] for r in rows_of(stub))
    return round(monthly), round(monthly * months)


def check_qa022_budget_and_flight():
    print("\nQA-022: a monthly budget stays monthly; the notes' flight vs the setup band's")
    stub = apply_hvac(flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    check("'$10,000 a month for 2 months' on a 2-month flight: $10,000/mo, $20,000 total",
          option_totals(stub, 2) == (10000, 20000), option_totals(stub, 2))
    check("... and no flight card (they agree)", not items_where(stub, "The notes say 2 months"))

    stub = apply_hvac()  # the setup flight is Sep 1 - Nov 30
    check("on a 3-month setup flight the monthly figure still holds: $10,000/mo",
          option_totals(stub, 3) == (10000, 30000), option_totals(stub, 3))
    card = items_where(stub, "The notes say 2 months")
    check("a decision card: notes say 2 months / setup says Sep 1 – Nov 30",
          len(card) == 1 and "the setup flight is Sep 1 – Nov 30, 2026 (3 months)"
          in card[0]["question"] and card[0]["priority"] == "decision", card)
    if card:
        labels = [o["label"] for o in card[0]["options"]]
        check("... Keep the setup flight, or use the notes' 2 months",
              labels == ["Keep the setup flight (Sep 1 – Nov 30, 2026)",
                         "Use the notes: Sep 1 – Oct 31, 2026 (2 months)"], labels)
        answer(stub, card[0], "notes")
        state = stub.session_state
        check("applying it moves the flight to Sep 1 – Oct 31",
              (state["flight_start"], state["flight_end"]) == (date(2026, 9, 1), date(2026, 10, 31)))
        check("... and re-prices: $10,000/mo, $20,000 over 2 months",
              option_totals(stub, 2) == (10000, 20000), option_totals(stub, 2))
        undo(stub)
        check("Undo puts the setup flight and the plan back",
              (stub.session_state["flight_start"], option_totals(stub, 3))
              == (date(2026, 9, 1), (10000, 30000)))

    stub = apply_hvac(flight_start=None, flight_end=None)
    card = items_where(stub, "no flight dates are set")
    default_text = app._flight_text(*app.default_flight())
    check(f"no flight set yet: the card names the default it priced over ({default_text})",
          len(card) == 1 and f"{default_text} (3 months)" in card[0]["question"], card)
    if card:
        start = app.default_flight()[0]
        check("... and offers the notes' 2 months from the same start",
              f"Use the notes: {app._flight_text(start, app._add_months(start, 2) - timedelta(days=1))}"
              in [o["label"] for o in card[0]["options"]][1], card[0]["options"])

    print("\nThe default flight is the 1st of next month, 3 months long, never in the past")
    for today in (date(2026, 10, 23), date(2026, 12, 31), date(2027, 1, 31), date(2028, 2, 29),
                  date(2026, 11, 1), date.today()):
        start, end = app.default_flight(today)
        check(f"today {today}: {start} – {end}",
              start > today and start.day == 1 and (end + timedelta(days=1)).day == 1
              and len(app.month_list(start, end)) == 3, (start, end))
    check("... and the monthly budget still reads $10,000/mo",
          option_totals(stub, 3)[0] == 10000, option_totals(stub, 3))

    plain = {**HVAC_DRAFT, "total_budget": 20000, "total_budget_period": "campaign",
             "stated_flight": None}
    for flight, months, monthly in ((TWO_MONTHS, 2, 10000), (FLIGHT, 3, 6667)):
        stub = apply_hvac(plain, notes="HVAC company, RETAIL Home Services HVAC and homeowners. "
                                       "$20,000.", flight_start=flight[0], flight_end=flight[1])
        check(f"'$20,000' with no flight, {months}-month setup flight: $20,000 total "
              f"(${monthly:,}/mo), no flight card",
              option_totals(stub, months) == (monthly, 20000)
              and not items_where(stub, "The notes say"), option_totals(stub, months))

    # A campaign total with a stated length: the card re-prices so the total holds.
    stub = apply_hvac({**plain, "stated_flight": None},
                      notes="HVAC company, RETAIL Home Services HVAC and homeowners. $20,000 "
                            "for 2 months.")
    card = items_where(stub, "The notes say 2 months")
    check("'for 2 months' in the notes is read without the model's field", len(card) == 1, card)
    if card:
        answer(stub, card[0], "notes")
        check("applying it holds the $20,000 total over 2 months ($10,000/mo)",
              option_totals(stub, 2) == (10000, 20000), option_totals(stub, 2))


def check_qa023_unasked_criteria():
    print("\nQA-023: a custom's criterion the notes didn't ask for is never kept silently")
    stub = apply_hvac(flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    line = rows_of(stub)[0]
    check("the line's Targeting states only what the notes asked for",
          line["Targeting"] == "Homeowners in-market for HVAC services", line["Targeting"])
    check("... and so does Campaign Specs' Audience",
          stub.session_state.get("audience_text") == "Homeowners in-market for HVAC services",
          stub.session_state.get("audience_text"))
    card = items_where(stub, "This custom also targets homes built before 2014")
    check("a card: keep the custom, or switch to the plain segment",
          len(card) == 1 and [o["label"] for o in card[0]["options"]]
          == [f"Keep {HVAC_CUSTOM}", "Switch to RETAIL Home Services HVAC AND DEMO Homeowner"],
          card)
    if card:
        answer(stub, card[0], "plain")
        check("switching puts the plain segments on the line",
              rows_of(stub)[0]["_segments"] == ["RETAIL Home Services HVAC", "DEMO Homeowner"],
              rows_of(stub)[0]["_segments"])
    stub = apply_hvac({**HVAC_DRAFT, "media_plan_lines": [
        {**HVAC_DRAFT["media_plan_lines"][0], "segments": ["RETAIL Home Services HVAC",
                                                          "DEMO Homeowner"]}]},
        notes="HVAC company in Springfield reaching homeowners. $10,000 a month.",
        flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    check("the recommender's primary is never a custom with a criterion the notes didn't ask for",
          rows_of(stub)[0]["_segments"][0] != HVAC_CUSTOM, rows_of(stub)[0]["_segments"])


def check_alternatives_keep_asked_criteria():
    print("\nThe alternatives card never trades away a criterion the notes asked for")
    packaged = "CUSTOM DEMO Homeowner AND HH Income 150K Plus"
    stub = apply_hvac({**HVAC_DRAFT, "media_plan_lines": [
        {**HVAC_DRAFT["media_plan_lines"][0], "segments": [packaged],
         "audience_track": "Homeowners with $150K+ household income"}],
        "campaign_specs": {}, "audiences": []},
        notes="Heating and cooling company reaching homeowners with $150K+ household income. "
              "$10,000 a month for 2 months.",
        flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    check("a custom packaging exactly the notes' criteria stays on the line, no swap card",
          rows_of(stub)[0]["_segments"] == [packaged]
          and not items_where(stub, "use an alternative")
          and not items_where(stub, "Targeting says"),
          (rows_of(stub)[0]["_segments"], [i["question"] for i in
                                           stub.session_state["draft_review_items"]]))


def check_qa026_named_segment():
    print("\nQA-026: no audience-alternatives card when the notes name the segment")
    stub = apply_hvac(flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    check("notes name RETAIL Home Services HVAC (inside the custom) -> no alternatives card",
          not items_where(stub, "use an alternative"), [i["question"] for i in
                                                  stub.session_state["draft_review_items"]])
    check("names match however they're punctuated",
          app.notes_name_segment("RETAIL Home Services HVAC", "target retail home-services HVAC")
          and app.notes_name_segment(HVAC_CUSTOM, "RETAIL Home Services HVAC")
          and not app.notes_name_segment("DEMO Homeowner", "homeowners"))


def check_qa024_notes_name():
    print("\nQA-024: 'The notes name' only for a segment the notes name")
    stub = apply_hvac(flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    note = next((n for n in stub.session_state["draft_unresolved_internal"]
                 if "avails document" in n), "")
    check("a named segment reads 'The notes name ...'",
          note.startswith(f"The notes name {HVAC_CUSTOM}"), note)
    check("... described by the segments the line actually has (one description everywhere)",
          all(seg in note for seg in rows_of(stub)[0]["_segments"]), note)
    stub = apply_hvac({**HVAC_DRAFT, "audiences": [{"segment": "RETAIL Home Services HVAC"}]},
                      notes="Heating and cooling company in Springfield reaching homeowners. "
                            "$10,000 a month for 2 months.",
                      flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1])
    note = next((n for n in stub.session_state["draft_unresolved_internal"]
                 if "avails document" in n), "")
    check("a segment the draft picked reads 'Chosen for the audience the notes describe'",
          note.startswith("Chosen for the audience") and "The notes name" not in note, note)


def check_qa025_even_split():
    print("\nQA-025 / split cards: an even option, numbers both ways, a numeric question")
    draft = {**PLAN_DRAFT, "media_plan_lines": [], "group_selection": {"mode": "all"},
             "group_allocation": {"proportional_to_avails": True},
             "review_items": [{"type": "choice", "audience": "client",
                               "question": "They didn't say how to weight the audiences. More "
                                           "toward TRAVEL Family or more toward TRAVEL Golf?",
                               "options": [{"label": "x", "patch": [{
                                   "op": "set_group_allocation", "option": "",
                                   "group_allocation": {"percent_of_remainder": 60}}]}]}]}
    groups = [g for g in groups_fixture() if g["id"] != "g-dup"]
    stub = apply_draft(draft, targeting_groups=groups, avails_mode=True)
    card = items_where(stub, "how to weight")
    labels = [o["label"] for o in card[0]["options"]] if card else []
    check("the even split is offered", "Even split: 50/50 between TRAVEL Family and TRAVEL Golf"
          in labels, labels)
    check("... with 60/40 and 70/30 each way", {"60/40 toward TRAVEL Family",
                                                "60/40 toward TRAVEL Golf",
                                                "70/30 toward TRAVEL Golf"} <= set(labels), labels)
    question = card[0]["question"] if card else ""
    check("the question states the drafted split in numbers, never 'more toward'",
          "more toward" not in question.lower() and "Drafted at 67/33 toward TRAVEL Family"
          in question, question)
    three = groups_fixture()[:2] + [tg.new_group(["TRAVEL Spa"], name="Spa", avails_monthly=50000,
                                                 group_id="g-c")]
    stub = apply_draft(draft, targeting_groups=three, avails_mode=True)
    card = items_where(stub, "how to weight")
    labels = [o["label"] for o in card[0]["options"]] if card else []
    check("three audiences: an equal split across the three",
          any(l.startswith("Even split across the 3: 34/33/33") for l in labels), labels)


def check_undo_resets_selections():
    print("\nUndo reopens the cards on the draft's own answers")
    stub = apply_draft(PLAN_DRAFT)
    split = items_where(stub, "no split was stated")[0]
    answer(stub, split, split["options"][1]["id"])
    undo(stub)
    check("after Undo no card holds the undone answer",
          not stub.session_state.get("draft_review_answers"),
          stub.session_state.get("draft_review_answers"))


def check_qa027_small_shares():
    print("\nQA-027: a share under 1% reads '<1%', never '0%'; reach notes name the market")
    check("share_text", (app.share_text(0.35), app.share_text(0), app.share_text(13.4))
          == ("<1%", "0%", "13%"))
    note = ("The Premion Streaming TV line's $113 reaches about 3,531 impressions -- roughly <1% "
            "of the 1,011,108 avails on TRAVEL Family in Harrisburg for the flight.")
    out = app.collapse_reach_notes([note, note.replace("$113", "$4,300")])
    check("the collapsed note keeps '<1%' and the market",
          "<1% of 1,011,108 avails" in out[0] and "TRAVEL Family in Harrisburg" in out[0], out)
    draft = {**PLAN_DRAFT, "review_items": [], "media_plan_lines": [],
             "group_selection": {"mode": "all"},
             "group_allocation": {"proportional_to_avails": True}}
    groups = [dict(g, geo_def={"kind": "dma", "label": "Harrisburg"}, name="Harrisburg")
              for g in groups_fixture() if g["id"] != "g-dup"]
    stub = apply_draft(draft, targeting_groups=groups, avails_mode=True,
                       avails_seed_rows=tg.groups_to_seed_rows(groups, app.AVAILS_COLUMN_MONTHLY))
    reach = [n for n in stub.session_state["draft_unresolved_internal"]
             + stub.session_state["draft_unresolved"] if "avails" in n and "reaches" in n]
    check("each reach line names its audience and market",
          reach and all(f"{aud} in Harrisburg" in reach[0] for aud in ("TRAVEL Family",
                                                                       "TRAVEL Golf")), reach)


def check_client_name_note():
    print("\nA note saying no client name was given, with the field filled in")
    stub = apply_hvac(flight_start=TWO_MONTHS[0], flight_end=TWO_MONTHS[1],
                      client_name="QA-TEST-Springfield HVAC")
    check("is dropped", not any("client name" in n.lower()
                                for n in stub.session_state["draft_unresolved_internal"]),
          stub.session_state["draft_unresolved_internal"])
    stub = apply_hvac({**HVAC_DRAFT, "client_name": ""}, flight_start=TWO_MONTHS[0],
                      flight_end=TWO_MONTHS[1], client_name="")
    check("... and kept when the field really is empty",
          any("client name" in n.lower() for n in stub.session_state["draft_unresolved_internal"]))


def check_verticals():
    print("\nVerticals: B2B / Professional Services and Telecom/Tech")
    from catalog_shared import VERTICALS, _detect_vertical_hint
    check("both are on the list", VERTICALS.get("B2B / Professional Services") == "b2b"
          and VERTICALS.get("Telecom/Tech") == "telecom_tech")
    for notes, want in (
            ("Netmaker sells VoIP and managed IT to small law, medical and insurance offices "
             "in Northern Virginia.", "telecom_tech"),
            ("A VoIP provider targeting bank branches and credit unions.", "telecom_tech"),
            ("Staffing agency placing accountants with mid-size firms.", "b2b"),
            ("Credit union promoting auto loans.", "banking"),
            ("Family law firm in Fairfax.", "legal")):
        got = _detect_vertical_hint(notes)
        check(f"{notes[:48]}... -> {want}", got == want, got)


def check_finder_collapse():
    print("\nFinder results: normalized names, near-duplicates and typos folded")
    import segment_recommend as sr
    names = ["FOOD Beer Wine and Spirits", "DEMO Entrepreneur", "FOOD Beer Wine and Spiritsr",
             "DEMO Entrepeneur", "RETAIL Home Services Oil_HVAC", "RETAIL Home Services HVAC",
             "HH Income 150K Plus", "HH Income 200K Plus"]
    clusters = dict(sr.collapse_in_order(names))
    check("typos fold under the first spelling",
          clusters.get("FOOD Beer Wine and Spirits") == ["FOOD Beer Wine and Spiritsr"]
          and clusters.get("DEMO Entrepreneur") == ["DEMO Entrepeneur"], clusters)
    check("an underscore variant folds in too",
          "RETAIL Home Services HVAC" in clusters.get("RETAIL Home Services Oil_HVAC", []),
          clusters)
    check("different thresholds stay apart",
          "HH Income 150K Plus" in clusters and "HH Income 200K Plus" in clusters, clusters)
    check("display names lose underscores",
          sr.display_name("RETAIL Home Services Oil_HVAC") == "RETAIL Home Services Oil HVAC")


def main():
    check_skip_all()
    check_plan_items()
    check_model_item_rules()
    check_phase_15_items()
    check_proxy_and_duplicate_prose()
    check_reach_notes_collapse()
    check_reprice_keeps_monthly_budget()
    check_one_card_across_options()
    check_audience_alternatives()
    check_unexpected_geographies()
    check_client_name_prefix()
    check_split_labels()
    check_plain_words()
    check_apply_button_types()
    check_gross_and_total_tv()
    check_gdw()
    check_avails_selection()
    check_client_routing()
    check_cards_and_history()
    check_qa022_budget_and_flight()
    check_qa023_unasked_criteria()
    check_qa026_named_segment()
    check_alternatives_keep_asked_criteria()
    check_qa024_notes_name()
    check_qa025_even_split()
    check_undo_resets_selections()
    check_qa027_small_shares()
    check_client_name_note()
    check_verticals()
    check_finder_collapse()
    print(f"\n{'ALL PASS' if not failures else f'{len(failures)} FAILED'}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
