"""Great Day Washington (GDW) -- WUSA9's weekday lifestyle-show segment.

    python tests/test_great_day_washington.py

What it guards:
  * DC-originated only: offered in Section C only on a DC proposal -- never on
    Harrisburg, even targeting DC; a draft on any other proposal doesn't add
    it and says so.
  * One-time (default): $1,500 once, in the first month, OUTSIDE the Monthly
    Totals row and counted once in the Full Flight Total. Monthly: $1,500 a
    month, in Monthly Totals and multiplied across the flight, Flight cell =
    the whole flight. The Flight cell is re-derived on a toggle or a
    Monthly/Full Flight switch.
  * Added Value: the cost cell reads "Added Value", $0 goes into the totals,
    and the plan carries "Includes a Great Day Washington segment as added
    value ($1,500 value)."
  * It never triggers Total TV -- but sits on a Total TV DC plan like any line.
  * The Targeting copy is fixed: no draft or re-draft rewrites it.
  * Its two slides (gdw:overview, gdw:production) are in the deck, before the
    media plan, exactly when a GDW line is on the plan.

The deck checks build against the v13 master (TEGNA_MASTER_DECK_10-08-26_gdw.pptx,
gitignored) and SKIP without it. Offline: db.log_proposal and the logo/storage
calls are stubbed, as in test_draft_regression.py, so nothing is written to
production.
"""
import copy
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"
os.chdir(REPO)

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import slide_map                               # noqa: E402

V13 = REPO / "TEGNA_MASTER_DECK_10-08-26_gdw.pptx"
DRAFT = FIXTURES / "dental_single_option.draft.json"
FLIGHT = (date(2026, 9, 1), date(2026, 11, 30))

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


# ---------------------------------------------------------------------------
# Plan math (pure)
# ---------------------------------------------------------------------------

def _rate_row(cost):
    return {"Tactic": "Premion Streaming TV", "Flight": "Sep-Nov", "Geo": "Washington, DC",
            "Targeting": "Adults 25-54", "Impressions": cost / 32 * 1000, "CPM": 32.0,
            "Type": app.ROW_TYPE_RATE, "Cost": float(cost)}


def check_plan_math():
    print("\nPlan math: one flat cost, counted once")
    for n_months in (1, 3, 6):
        rows = [_rate_row(10000), app.gdw_row("Sep")]
        totals = app.compute_plan_totals(rows, app.BREAKOUT_MONTHLY, n_months, "Sep-Nov")
        check(f"paid GDW adds $1,500 once to a {n_months}-month flight total",
              totals["full_flight_cost"] == 10000 * n_months + 1500, totals["full_flight_cost"])
        gdw = next(r for r in totals["preview_rows"] if r["is_gdw"])
        check(f"monthly plan, {n_months} month(s): the segment sits whole in its month ($1,500, not split)",
              gdw["monthly_cost"] == 1500, gdw["monthly_cost"])
        check(f"... and one-time is left out of Monthly Totals ({n_months} month(s))",
              totals["monthly_cost"] == 10000, totals["monthly_cost"])
        check(f"... so Monthly Totals x months = Full Flight Total - one-time GDW ({n_months})",
              totals["monthly_cost"] * n_months == totals["full_flight_cost"] - 1500)

    rows = [_rate_row(30000), app.gdw_row("Sep-Nov")]
    totals = app.compute_plan_totals(rows, app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov")
    check("full-flight plan: GDW adds $1,500 once", totals["full_flight_cost"] == 31500,
          totals["full_flight_cost"])

    rows = [_rate_row(10000), app.gdw_row("Sep", added_value=True)]
    totals = app.compute_plan_totals(rows, app.BREAKOUT_MONTHLY, 3, "Sep-Nov", markup=1.15)
    print("\nPlan math: a monthly segment")
    for n_months in (1, 3):
        monthly = dict(app.gdw_row("Sep-Nov", frequency=app.GDW_MONTHLY))
        totals = app.compute_plan_totals([_rate_row(10000), monthly], app.BREAKOUT_MONTHLY,
                                         n_months, "Sep-Nov")
        check(f"monthly plan, {n_months} month(s): $1,500 a month in Monthly Totals",
              totals["monthly_cost"] == 11500, totals["monthly_cost"])
        check(f"... and multiplied across the flight",
              totals["full_flight_cost"] == 11500 * n_months, totals["full_flight_cost"])
        full = dict(app.gdw_row("Sep-Nov", cost=1500 * n_months, frequency=app.GDW_MONTHLY))
        totals = app.compute_plan_totals([_rate_row(10000 * n_months), full], app.BREAKOUT_FULL_FLIGHT,
                                         n_months, "Sep-Nov")
        check(f"full-flight plan, {n_months} month(s): stored as price x months, $1,500 a month",
              totals["full_flight_cost"] == 11500 * n_months and totals["monthly_cost"] == 11500,
              (totals["full_flight_cost"], totals["monthly_cost"]))

    one_time_full = app.compute_plan_totals([_rate_row(30000), app.gdw_row("Sep")],
                                            app.BREAKOUT_FULL_FLIGHT, 3, "Sep-Nov")
    check("full-flight plan, one-time: counted once in the flight, out of the monthly figure",
          one_time_full["full_flight_cost"] == 31500 and one_time_full["monthly_cost"] == 10000,
          (one_time_full["full_flight_cost"], one_time_full["monthly_cost"]))

    for frequency, want in ((app.GDW_ONE_TIME, "($1,500 value)"), (app.GDW_MONTHLY, "($1,500/month value)")):
        for breakout, cost in ((app.BREAKOUT_MONTHLY, 1500), (app.BREAKOUT_FULL_FLIGHT,
                                                             1500 * (3 if frequency == app.GDW_MONTHLY else 1))):
            row = app.gdw_row("Sep", cost=cost, added_value=True, frequency=frequency)
            rate = _rate_row(30000 if breakout == app.BREAKOUT_FULL_FLIGHT else 10000)
            totals = app.compute_plan_totals([rate, row], breakout, 3, "Sep-Nov")
            notes = app.added_value_notes(totals["preview_rows"])
            check(f"Added Value, {frequency}, {breakout}: $0 in both totals and the note reads {want}",
                  round(totals["monthly_cost"]) == 10000 and round(totals["full_flight_cost"]) == 30000
                  and notes == [f"Includes a Great Day Washington segment as added value {want}."],
                  (totals["monthly_cost"], totals["full_flight_cost"], notes))

    print("\nFlight cell: re-derived on a toggle or a breakout switch")
    option = app.new_plan_option("Option A", [_rate_row(10000), app.gdw_row("Sep")])
    app.set_gdw_frequency(option, app.GDW_MONTHLY, 3, "Sep", "Sep-Nov")
    gdw = option["rows"][1]
    check("one-time -> monthly: Flight becomes the whole flight, price unchanged (monthly plan)",
          (gdw["Flight"], gdw["Cost"]) == ("Sep-Nov", 1500), gdw)
    option["_breakout_basis"] = app.BREAKOUT_MONTHLY     # what main() records every render
    option["breakout"] = app.BREAKOUT_FULL_FLIGHT
    app.rescale_rows_for_breakout_change(option, 3, gdw_month="Sep", flight_label="Sep-Nov")
    check("monthly line, Monthly -> Full Flight: Cost becomes price x months, Flight the flight",
          (gdw["Flight"], gdw["Cost"]) == ("Sep-Nov", 4500), gdw)
    app.set_gdw_frequency(option, app.GDW_ONE_TIME, 3, "Sep", "Sep-Nov")
    check("monthly -> one-time in a Full-Flight option: back to one $1,500 segment in month 1",
          (gdw["Flight"], gdw["Cost"]) == ("Sep", 1500), gdw)
    gdw["Flight"] = "Oct"           # the rep moved it
    option["breakout"] = app.BREAKOUT_MONTHLY
    app.rescale_rows_for_breakout_change(option, 3, gdw_month="Sep", flight_label="Sep-Nov")
    check("one-time line, Full Flight -> Monthly: Flight re-derived to a month label, Cost unscaled",
          (gdw["Flight"], gdw["Cost"]) == ("Sep", 1500), gdw)

    rows = [_rate_row(10000), app.gdw_row("Sep", added_value=True)]
    totals = app.compute_plan_totals(rows, app.BREAKOUT_MONTHLY, 3, "Sep-Nov", markup=1.15)
    check("an Added Value line adds $0 to the flight total (gross-up on, too)",
          round(totals["full_flight_cost"]) == round(30000 * 1.15), totals["full_flight_cost"])
    check("... and $0 to the monthly total", round(totals["monthly_cost"]) == round(10000 * 1.15),
          totals["monthly_cost"])
    notes = app.added_value_notes(totals["preview_rows"])
    check("the value note under the plan reads exactly as specified",
          notes == ["Includes a Great Day Washington segment as added value ($1,500 value)."], notes)
    gdw = next(r for r in totals["preview_rows"] if r["is_gdw"])
    check("cost cell reads 'Added Value'", app.cost_cell(gdw, gdw["monthly_cost"]) == "Added Value",
          app.cost_cell(gdw, gdw["monthly_cost"]))
    check("impressions/CPM cells read '—'", app.no_figure(gdw) == "—", app.no_figure(gdw))

    edited = app.gdw_row("Sep", cost=2500, added_value=True)
    notes = app.added_value_notes(app.compute_plan_totals(
        [edited], app.BREAKOUT_MONTHLY, 3, "Sep-Nov")["preview_rows"])
    check("the rep's own value amount is what the note quotes",
          notes == ["Includes a Great Day Washington segment as added value ($2,500 value)."], notes)

    paid = app.gdw_row("Sep")
    check("a paid GDW line is never grossed up (flat, like every flat fee)",
          app.compute_plan_totals([paid], app.BREAKOUT_MONTHLY, 3, "Sep-Nov",
                                  markup=1.15)["full_flight_cost"] == 1500)

    ranges = app.flight_month_ranges(*FLIGHT)
    check("monthly flighting: the default Flight cell is the first month",
          app.gdw_first_month(ranges) == "Sep", app.gdw_first_month(ranges))
    seeded = app.seed_media_plan_rows(
        {"products": {app.GDW_PRODUCT_KEY: True}, "_premion_streaming_tv": False},
        "Washington, DC", "", "Sep-Nov", gdw_flight=app.gdw_first_month(ranges))
    check("seeded GDW row: Sep, Washington, DC, $1,500 flat, fixed copy",
          seeded == [app.gdw_row("Sep")], seeded)
    check("the seeded row's tactic and geo cells", (seeded[0]["Tactic"], seeded[0]["Geo"])
          == ("Great Day Washington (WUSA9)", "Washington, DC"), seeded[0])


# ---------------------------------------------------------------------------
# Total TV is never triggered
# ---------------------------------------------------------------------------

def check_never_total_tv():
    print("\nTotal TV: GDW never switches it on")
    check("GDW is not a line product the draft maps onto a widget (and so onto total_tv)",
          app.GDW_PRODUCT_KEY not in app.PRODUCT_TO_WIDGET_KEYS)
    sel = app.read_products_selection(
        lambda key, default=False: {"market_choice": "DC", app.GDW_PRODUCT_KEY: True}.get(key, default))
    check("GDW selected on a DC proposal leaves total_tv off", sel["total_tv"] is False
          and sel[app.GDW_PRODUCT_KEY] is True, sel)
    check("a GDW line is not a broadcast line (no gross exemption, no schedule months)",
          not app.is_broadcast_row(app.gdw_row("Sep")))
    keys = assembly.resolve_active_keys({**copy.deepcopy(assembly.SELECTIONS),
                                         "great_day_washington": True,
                                         "products": {**assembly.SELECTIONS["products"],
                                                      "total_tv": False}})
    check("resolve_active_keys: GDW adds its two slides and no Total TV key",
          assembly.GDW_SLIDE_KEYS <= keys and not any(k.startswith(("total_tv", "proposal_template_total_tv",
                                                                    "cobrand")) for k in keys),
          sorted(k for k in keys if "total" in k or "gdw" in k))


# ---------------------------------------------------------------------------
# Draft from notes
# ---------------------------------------------------------------------------

class _StubSt:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def load_draft():
    raw = json.loads(DRAFT.read_text(encoding="utf-8"))
    raw.pop("_comment", None)
    return raw


def apply_draft(draft, notes, market="DC", state=None, skip_sections=None):
    real = app.st
    stub = _StubSt()
    stub.session_state.update(state or {
        "market_choice": market, "flight_start": FLIGHT[0], "flight_end": FLIGHT[1],
        "avails_mode": False})
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft), skip_sections=skip_sections, notes=notes)
    finally:
        app.st = real
    return stub.session_state


def _gdw_rows(state):
    return [r for o in state.get("plan_options") or [] for r in o["rows"] if app.is_gdw_row(r)]


def check_notes_parsing():
    print("\nDraft from notes: detection")
    cases = [
        ("add a Great Day segment as AV", {"added_value": True, "price": None}),
        ("include GDW", {"added_value": False, "price": None}),
        ("GDW for $2,000", {"added_value": False, "price": 2000.0}),
        ("had a great day with the client", None),
        ("Great day! They want CTV across DC.", None),
        ("Throw in a Great Day Washington segment as a bonus.", {"added_value": True, "price": None}),
        ("Client wants a GDW segment, $20k CTV budget over 3 months.",
         {"added_value": False, "price": None}),
        ("Elaine would interview the owner on the show.", {"added_value": False, "price": None}),
        ("Book a lifestyle segment on WUSA9 for them.", {"added_value": False, "price": None}),
        ("monthly GDW segment as AV", {"added_value": True, "price": None}),
    ]
    for notes, want in cases:
        got = app.detect_gdw(notes)
        got = got and {"added_value": got["added_value"], "price": got["price"]}
        check(f"{notes!r} -> {want}", got == want, got)

    for notes, want in (("monthly GDW segment as AV", True), ("include GDW", False),
                        ("Add a Great Day segment every month.", True),
                        ("A recurring GDW interview.", True),
                        ("$20k monthly CTV budget across the DC DMA, and we'll add a one-off GDW.", False)):
        check(f"{notes!r} -> {'Monthly' if want else 'One-time'}",
              app.detect_gdw(notes)["monthly"] is want, app.detect_gdw(notes))

    print("\nDraft from notes: what lands on the plan")
    draft = load_draft()

    state = apply_draft(draft, "Dental practice, $30k. Add a Great Day segment as AV.")
    rows = _gdw_rows(state)
    check("'add a Great Day segment as AV' -> one GDW line, Added Value",
          len(rows) == 1 and rows[0]["Type"] == app.ROW_TYPE_ADDED_VALUE, rows)
    totals = app.compute_plan_totals(state["plan_options"][0]["rows"], app.BREAKOUT_MONTHLY, 3, "Sep-Nov")
    check("... which adds $0: the plan still ties to the stated $30,000",
          round(totals["full_flight_cost"]) == 30000, totals["full_flight_cost"])
    check("... and no Added Value question (the notes answered it)", not state.get("gdw_av_review"))
    check("... and the checkbox is on", state.get(app.GDW_PRODUCT_KEY) is True)
    check("... in the first month of a monthly plan", rows and rows[0]["Flight"] == "Sep", rows)

    state = apply_draft(draft, "Dental practice, $30k. Include GDW.")
    rows = _gdw_rows(state)
    check("'include GDW' -> One-time, $1,500 paid", len(rows) == 1 and rows[0]["Type"] == app.ROW_TYPE_FLAT_FEE
          and rows[0]["Cost"] == 1500 and not app.is_gdw_monthly(rows[0]), rows)
    check("... with the 'mark as Added Value?' review prompt queued", state.get("gdw_av_review") is True)
    totals = app.compute_plan_totals(state["plan_options"][0]["rows"], app.BREAKOUT_MONTHLY, 3, "Sep-Nov")
    check("... counted once, on top of the $30,000 media budget",
          round(totals["full_flight_cost"]) == 31500, totals["full_flight_cost"])

    state = apply_draft(draft, "Dental practice, $30k. Add a monthly GDW segment as AV.")
    rows = _gdw_rows(state)
    check("'monthly GDW segment as AV' -> Monthly + Added Value, Flight = the whole flight",
          len(rows) == 1 and app.is_gdw_monthly(rows[0]) and rows[0]["Type"] == app.ROW_TYPE_ADDED_VALUE
          and rows[0]["Flight"] != "Sep", rows)
    totals = app.compute_plan_totals(state["plan_options"][0]["rows"], app.BREAKOUT_MONTHLY, 3, "Sep-Nov")
    check("... $0 in the totals, '/month' in the note", round(totals["full_flight_cost"]) == 30000
          and app.added_value_notes(totals["preview_rows"])
          == ["Includes a Great Day Washington segment as added value ($1,500/month value)."],
          (totals["full_flight_cost"], app.added_value_notes(totals["preview_rows"])))

    state = apply_draft(draft, "Dental practice, $30k. GDW for $2,000.")
    rows = _gdw_rows(state)
    check("'GDW for $2,000' -> $2,000 paid, no review prompt",
          len(rows) == 1 and rows[0]["Cost"] == 2000 and not state.get("gdw_av_review"), rows)

    state = apply_draft(draft, "Dental practice, $30k. We had a great day with the client.")
    check("'had a great day with the client' -> no GDW", not _gdw_rows(state), _gdw_rows(state))
    check("... and the checkbox is off", not state.get(app.GDW_PRODUCT_KEY))

    state = apply_draft(draft, "Dental practice, $30k. Include GDW.", market="Harrisburg")
    internal = state.get("draft_unresolved_internal") or []
    check("GDW on a Harrisburg proposal -> not added", not _gdw_rows(state), _gdw_rows(state))
    check("... with the DC-only warning", app.GDW_DC_ONLY_NOTE in internal, internal)
    check("... and the checkbox stays off", not state.get(app.GDW_PRODUCT_KEY))

    tv_draft = dict(draft, total_tv=True)
    state = apply_draft(tv_draft, "Dental practice, $30k. Add a Great Day Washington segment on WUSA9.")
    check("'Great Day' alone never moves the draft to Total TV, even if the model says so",
          not state.get("total_tv") and not (state.get("_draft_pending_fields") or {}).get("total_tv"),
          (state.get("total_tv"), state.get("_draft_pending_fields")))
    state = apply_draft(tv_draft, "Dental practice, $30k. Keep the WUSA9 schedule going and add GDW.")
    pending = state.get("_draft_pending_fields") or {}
    check("a real broadcast mention beside GDW still turns Total TV on",
          state.get("total_tv") or pending.get("total_tv"), pending)


def check_fixed_description():
    print("\nFixed description: no draft or re-draft rewrites it")
    draft = load_draft()
    notes = "Dental practice, $30k. Include GDW."
    first = apply_draft(draft, notes)
    redraft = load_draft()
    redraft["media_plan_lines"] = redraft["media_plan_lines"] + [
        {"product": "custom_fee", "label": "Great Day Washington Segment",
         "allocation": {"flat_amount": 1500}}]
    redraft["strategy_summary"] = "Streaming TV plus a featured Great Day Washington segment."
    state = apply_draft(redraft, notes, state=first, skip_sections=set())
    rows = _gdw_rows(state)
    check("a re-draft whose model wrote its own GDW line still leaves exactly one per option",
          len(rows) == len(state["plan_options"]), [r["Tactic"] for r in rows])
    check("the GDW Targeting cell equals the constant exactly",
          all(r["Targeting"] == app.GDW_TARGETING for r in rows), [r["Targeting"] for r in rows])
    hand_edited = dict(app.gdw_row("Sep"), Targeting="Something else entirely")
    preview = app.compute_plan_totals([hand_edited], app.BREAKOUT_MONTHLY, 3, "Sep-Nov")["preview_rows"][0]
    check("the deck/preview description is the constant even if the cell was changed",
          preview["targeting"] == app.GDW_TARGETING, preview["targeting"])
    check("the constant is the exact approved text", app.GDW_TARGETING ==
          "3–4 minute featured interview segment on Great Day Washington (WUSA9, weekdays 9am or "
          "3pm), with digital copy posted on WUSA9.com/GreatDay.")
    defaults = app.resolve_row_defaults(app.GDW_TACTIC, "Harrisburg DMA", "Adults 25-54", "Sep-Nov",
                                        current=hand_edited, first_month="Sep")
    check("a shared-field re-seed restores the fixed copy and DC geo",
          defaults == {"Flight": "Sep", "Geo": "Washington, DC", "Targeting": app.GDW_TARGETING}, defaults)


# ---------------------------------------------------------------------------
# Deck assembly (v13)
# ---------------------------------------------------------------------------

def _keys(prs):
    return [slide_map.notes_key(s) if s.has_notes_slide else None for s in prs.slides]


def check_deck_assembly():
    print("\nDeck assembly: both slides exactly when GDW is on the plan")
    if not V13.exists():
        print(f"  SKIP  {V13.name} not present")
        return
    sel = copy.deepcopy(assembly.SELECTIONS)
    sel["great_day_washington"] = True
    prs, _c1, _c2 = assembly.build_presentation(str(V13), sel)
    keys = _keys(prs)
    check("GDW on: both gdw:* slides present", {"gdw:overview", "gdw:production"} <= set(keys), keys)
    # The sample selections are a Total TV DC deck, so the plan slide is that
    # variant -- whichever plan variant survived, both GDW slides precede it.
    plan_index = next((i for i, k in enumerate(keys) if str(k).startswith("proposal_template")), None)
    check("... both before the media plan slide", plan_index is not None and
          keys.index("gdw:overview") < keys.index("gdw:production") < plan_index, keys)

    sel["great_day_washington"] = False
    prs, _c1, _c2 = assembly.build_presentation(str(V13), sel)
    check("GDW off: neither slide present", not any(str(k).startswith("gdw:") for k in _keys(prs)),
          _keys(prs))
    del sel["great_day_washington"]
    prs, _c1, _c2 = assembly.build_presentation(str(V13), sel)
    check("a proposal logged before GDW existed (no key at all): neither slide",
          not any(str(k).startswith("gdw:") for k in _keys(prs)))


# ---------------------------------------------------------------------------
# The real form: picker, seeding, Generate
# ---------------------------------------------------------------------------

def run_form(state, click_av=False, generate=False):
    """Drive the real app. Returns (at, captured)."""
    from streamlit.testing.v1 import AppTest

    captured = {}
    real = (assembly.personalize, assembly._prepare_media_plan_slide, db.log_proposal,
            db.upload_proposal_logo, db.proposal_logo, db.master_deck)

    def spy_personalize(prs, fill_data):
        warnings = real[0](prs, fill_data)
        captured.update(prs=prs, fill_data=fill_data, warnings=warnings)
        return warnings

    def spy_prepare(slide, option):
        captured.setdefault("plan_slides", []).append(slide)
        return real[1](slide, option)

    assembly.personalize = spy_personalize
    assembly._prepare_media_plan_slide = spy_prepare
    db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)
    db.upload_proposal_logo = lambda *a, **k: (None, "stubbed")
    db.proposal_logo = lambda *a, **k: None
    if V13.exists():
        db.master_deck = lambda _fallback: (str(V13), None, None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Regression Suite"
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return at, captured
        if click_av:
            at.session_state["gdw_av_review"] = True
            at.session_state["draft_unresolved_internal"] = [
                "Great Day Washington ($1,500) is added on top of the drafted media budget.",
                "An unrelated review note."]
            at.run()
            buttons = [b for b in at.button if b.label == "Mark as Added Value"]
            captured["av_button_shown"] = bool(buttons)
            if buttons:
                buttons[0].click().run()
        if generate and not at.exception:
            buttons = [b for b in at.button if b.label == "Generate proposal"]
            captured["generate_shown"] = bool(buttons)
            if buttons:
                buttons[0].click().run()
    finally:
        (assembly.personalize, assembly._prepare_media_plan_slide, db.log_proposal,
         db.upload_proposal_logo, db.proposal_logo, db.master_deck) = real
    return at, captured


def _base_state(market="DC", **extra):
    return {"market_choice": market, "flight_start": FLIGHT[0], "flight_end": FLIGHT[1],
            "avails_mode": False, "client_name": "GDW Test Client", **extra}


def _plan_rows(at):
    return [r for o in at.session_state["plan_options"] for r in o["rows"]]


def _table_rows(slide):
    for shape in slide.shapes:
        if shape.has_table:
            return [[c.text_frame.text.strip() for c in row.cells] for row in shape.table.rows]
    return []


def _slide_text(slide):
    return slide_map.extract_slide_text(slide)


def check_picker():
    print("\nThe real form: product picker")
    at, _ = run_form(_base_state("DC"))
    check("DC proposal renders", not at.exception, at.exception)
    check("GDW is offered on a DC proposal", any(c.label == app.GDW_TACTIC for c in at.checkbox),
          [c.label for c in at.checkbox])

    at, _ = run_form(_base_state("Harrisburg", **{app.GDW_PRODUCT_KEY: True}))
    check("Harrisburg-only proposal renders", not at.exception, at.exception)
    check("GDW is not offered on a Harrisburg-only proposal",
          not any(c.label == app.GDW_TACTIC for c in at.checkbox))
    check("... and a stale selection puts no GDW line on its plan",
          not any(app.is_gdw_row(r) for r in _plan_rows(at)), _plan_rows(at))

    rows, _w = app.load_market_profiles()
    dc = [app.market_profile_option_label(r) for r in rows if r.get("label") == "Washington, DC"]
    if dc:
        at, _ = run_form(_base_state("Harrisburg", target_dma_choice=dc))
        check("a Harrisburg proposal TARGETING Washington, DC is still not offered GDW",
              not any(c.label == app.GDW_TACTIC for c in at.checkbox))
        state = apply_draft(load_draft(), "Dental practice, $30k. Include GDW.",
                            state={"market_choice": "Harrisburg", "target_dma_choice": dc,
                                   "flight_start": FLIGHT[0], "flight_end": FLIGHT[1],
                                   "avails_mode": False})
        check("... and a draft there doesn't add it, with the DC-only note",
              not _gdw_rows(state) and app.GDW_DC_ONLY_NOTE in (state.get("draft_unresolved_internal") or []))


def check_form_and_generate():
    print("\nThe real form: DC proposal, GDW as Added Value, Generate")
    at, cap = run_form(_base_state("DC", **{app.GDW_PRODUCT_KEY: True}), click_av=True, generate=True)
    check("renders and generates without raising", not at.exception, at.exception)
    if at.exception:
        return
    gdw = [r for r in _plan_rows(at) if app.is_gdw_row(r)]
    check("ticking GDW seeds one line, in month 1", len(gdw) == 1 and gdw[0]["Flight"] == "Sep", gdw)
    check("the 'Mark as Added Value' prompt was shown", cap.get("av_button_shown"))
    check("one click made it Added Value", gdw and gdw[0]["Type"] == app.ROW_TYPE_ADDED_VALUE, gdw)
    stale = [n for n in at.session_state["draft_unresolved_internal"]
             if str(n).startswith(app.GDW_ON_TOP_NOTE_PREFIX)]
    check("... and withdrew the 'added on top of the budget' review note", not stale, stale)
    check("Total TV is still off", not at.session_state["total_tv"])
    if not V13.exists() or "prs" not in cap:
        print(f"  SKIP  deck checks ({V13.name} not present or Generate didn't build)")
        return
    prs = cap["prs"]
    keys = _keys(prs)
    plan = cap["plan_slides"][0]
    check("the plan slide is the standard variant",
          slide_map.notes_key(plan) == "proposal_template", slide_map.notes_key(plan))
    check("no Total TV slide in the deck", not any(str(k).startswith(("total_tv", "proposal_template_total_tv"))
                                                   for k in keys), keys)
    plan_pos = [s.slide_id for s in prs.slides].index(plan.slide_id)
    check("both GDW slides before the plan slide",
          "gdw:overview" in keys and "gdw:production" in keys
          and keys.index("gdw:overview") < plan_pos and keys.index("gdw:production") < plan_pos, keys)
    table = _table_rows(plan)
    row = next((r for r in table if r and r[0].startswith(app.GDW_TACTIC)), None)
    check("the GDW line is on the plan table", row is not None, table)
    if row:
        check("its cost cell reads 'Added Value'", row[-1] == "Added Value", row)
        check("its impressions cell reads '—'", "—" in row, row)
        check("its description is the fixed copy", app.GDW_TARGETING in row, row)
        check("its geo cell is Washington, DC", "Washington, DC" in row, row)
    check("the value note is on the plan slide",
          "Includes a Great Day Washington segment as added value ($1,500 value)." in _slide_text(plan))
    option = cap["fill_data"]["media_plan_options"][0]
    check("the plan's own totals carry $0 for it", option["total_cost"] == "$0", option["total_cost"])


def check_frequency_radio():
    print("\nThe real form: the One-time / Monthly radio")
    from streamlit.testing.v1 import AppTest
    real_log = db.log_proposal
    db.log_proposal = lambda *a, **k: ("x", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Regression Suite"
        for key, value in _base_state("DC", **{app.GDW_PRODUCT_KEY: True}).items():
            at.session_state[key] = value
        at.run()
        radios = [r for r in at.radio if r.label == "Great Day Washington"]
        check("the radio renders under a plan with a GDW line, One-time by default",
              len(radios) == 1 and radios[0].value == app.GDW_ONE_TIME,
              [(r.label, r.value) for r in at.radio])
        if not radios:
            return
        radios[0].set_value(app.GDW_MONTHLY).run()
        check("switching it raises nothing", not at.exception, at.exception)
        gdw = [r for r in _plan_rows(at) if app.is_gdw_row(r)]
        ranges = app.flight_month_ranges(*FLIGHT)
        check("Monthly: the line's frequency and Flight follow (the whole flight)",
              gdw and app.is_gdw_monthly(gdw[0]) and gdw[0]["Flight"] == app.format_flight_shorthand(ranges),
              gdw)
    finally:
        db.log_proposal = real_log


def check_frequency_radio_per_option():
    """Two plan options that both carry GDW each keep their OWN One-time /
    Monthly setting: the radio renders once per option, is keyed by that
    option, and writes only that option's GDW row -- so option A can quote a
    one-time segment while option B quotes a monthly one, and each option's
    plan slide prices its own."""
    print("\nThe real form: two options with GDW -- each option's own One-time / Monthly")
    from streamlit.testing.v1 import AppTest
    real = (db.log_proposal, db.upload_proposal_logo, db.proposal_logo, db.master_deck,
            assembly._prepare_media_plan_slide)
    plan_slides = []

    def spy_prepare(slide, option):
        plan_slides.append(slide)
        return real[4](slide, option)

    db.log_proposal = lambda *a, **k: ("x", None)
    db.upload_proposal_logo = lambda *a, **k: (None, "stubbed")
    db.proposal_logo = lambda *a, **k: None
    assembly._prepare_media_plan_slide = spy_prepare
    if V13.exists():
        db.master_deck = lambda _fallback: (str(V13), None, None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Regression Suite"
        for key, value in _base_state("DC", **{app.GDW_PRODUCT_KEY: True}).items():
            at.session_state[key] = value
        at.run()
        add = [b for b in at.button if b.label == "➕ Add option"]
        check("the Add option button renders", bool(add))
        if not add:
            return
        add[0].click().run()
        options = at.session_state["plan_options"]
        radios = [r for r in at.radio if r.label == "Great Day Washington"]
        check("a copied option keeps its GDW line, and each option gets its own radio",
              len(options) == 2 and len(radios) == 2
              and all(any(app.is_gdw_row(r) for r in o["rows"]) for o in options),
              [(o["name"], [r.get("Tactic") for r in o["rows"]]) for o in options])
        if len(radios) != 2:
            return
        radios[1].set_value(app.GDW_MONTHLY).run()
        check("switching the second option's radio raises nothing", not at.exception, at.exception)
        options = at.session_state["plan_options"]
        freqs = [next(r.get(app.GDW_FREQUENCY_FIELD) or app.GDW_ONE_TIME
                      for r in o["rows"] if app.is_gdw_row(r)) for o in options]
        check("option A stays One-time while option B is Monthly",
              freqs == [app.GDW_ONE_TIME, app.GDW_MONTHLY], freqs)
        radios = [r for r in at.radio if r.label == "Great Day Washington"]
        check("each radio shows its own option's setting after the rerun",
              [r.value for r in radios] == [app.GDW_ONE_TIME, app.GDW_MONTHLY],
              [r.value for r in radios])
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not (V13.exists() and generate):
            print("  SKIP  per-option plan slide checks (master deck or Generate unavailable)")
            return
        generate[0].click().run()
        check("generates with two differently-set options", not at.exception, at.exception)
        rows_by_slide = [_table_rows(s) for s in plan_slides]
        gdw_rows = [next((r for r in rows if r and r[0].startswith(app.GDW_TACTIC)), None)
                    for rows in rows_by_slide]
        check("each option's plan slide carries its own GDW line",
              len(gdw_rows) >= 2 and all(gdw_rows[:2]), rows_by_slide)
        if len(gdw_rows) >= 2 and all(gdw_rows[:2]):
            ranges = app.flight_month_ranges(*FLIGHT)
            check("option A's Flight cell is one month; option B's is the whole flight",
                  app.format_flight_shorthand(ranges) not in gdw_rows[0]
                  and app.format_flight_shorthand(ranges) in gdw_rows[1], gdw_rows)
    finally:
        (db.log_proposal, db.upload_proposal_logo, db.proposal_logo, db.master_deck,
         assembly._prepare_media_plan_slide) = real


def check_total_tv_dc_with_gdw():
    print("\nThe real form: Total TV DC proposal with GDW -- the line rides on the Total TV plan")
    at, cap = run_form(_base_state("DC", total_tv=True, **{app.GDW_PRODUCT_KEY: True}), generate=True)
    check("renders and generates without raising", not at.exception, at.exception)
    if at.exception or not V13.exists() or "prs" not in cap:
        print("  SKIP  deck checks")
        return
    plan = cap["plan_slides"][0]
    check("the plan slide is the Total TV DC variant",
          slide_map.notes_key(plan) == "proposal_template_total_tv:dc", slide_map.notes_key(plan))
    table = _table_rows(plan)
    row = next((r for r in table if r and r[0].startswith(app.GDW_TACTIC)), None)
    check("the GDW line is on it, at $1,500", row is not None and row[-1].startswith("$1,500"), table)
    keys = _keys(cap["prs"])
    check("both GDW slides are in the Total TV deck too",
          "gdw:overview" in keys and "gdw:production" in keys, keys)


def main():
    check_plan_math()
    check_never_total_tv()
    check_notes_parsing()
    check_fixed_description()
    check_deck_assembly()
    check_picker()
    check_form_and_generate()
    check_frequency_radio()
    check_frequency_radio_per_option()
    check_total_tv_dc_with_gdw()

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
