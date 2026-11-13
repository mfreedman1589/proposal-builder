"""Draft-from-notes for a multi-option, multi-market plan (the MidPoint
review, 2026-11) -- synthetic names throughout; the real client's files stay
in the gitignored tests/fixtures/midpoint/.

The shape: one audience pulled in four DMAs (DC, Dallas, Phoenix, Atlanta),
and notes asking for Good / Better / Best --
  Good:   DC only, 500K CTV impressions + 500K retargeting display impressions
  Better: DC held at 500K, the other markets below DC in proportion to their
          avails, a retargeting pot matching the CTV impressions, ~$40-50K/mo
  Best:   Better scaled up proportionately to ~$75K/mo
over a "Jan-June" flight, broken out monthly, CTV at a negotiated $34.

Every expectation is read from the NOTES TEXT by this file's own regexes or
worked from the stated figures by hand (500K x $34 = $17,000) -- never read
back from the code under test. The drafted JSON below is the intent the
prompt asks the model for; Python does every number.

Also covered: the pre-Generate sanity checks fire on the plan the app used
to build (100% of DC's avails, a missing option), a $0 option is never
dropped silently, a yearless flight resolves to its next occurrence, the
avails date notice fires with no month in common, one market label and
plain-language targeting everywhere, the no-logo rule, and the debug bundle.

    python tests/test_multi_option_draft.py
"""
import copy
import io
import json
import os
import re
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")
os.environ.setdefault("PROPOSAL_BUILDER_TODAY", "2026-11-12")

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import avails_grid                             # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


CLIENT = "QA-TEST-Northgate Systems"
AUDIENCE = "DEMO Age M35-54, DEMO Career Business Owners Executives"
NOTES = f"""I need a plan for {CLIENT} that includes CTV primarily focused on DC with the other
markets on separate individual lines as secondary markets. Also want OTT retargeting included
at ~10% of the budget. Let's do a good, better, best set of options. First option will be just
DC market 500K CTV impressions and 500K retargeting display impressions. 2nd Option will bring
in the other markets. Keep DC at 500K then do budgets less than DC, but proportionate to the
avails for the other markets. Include a retargeting pot with a similar number of total display
impressions. The Monthly budget should be close to $40-50K/mo. Then for the Best option
increase DC and the other markets and retargeting proportionately to achieve ~$75K budget.
Use Jan-June (6 months) as the flight, broken out monthly. Include production and attribution.
CTV CPM is $34.
Meeting: November 6, 2026. Markets discussed: Washington DC (primary), Dallas, Phoenix, Atlanta.
"""

# The avails document: one audience, four DMAs, a December 2026 window.
MARKETS = [("g_dc", "washington_hagerstown", "Washington, D.C.", 14435026),
           ("g_dal", "dallas_fort_worth", "Dallas-Fort Worth", 25780995),
           ("g_phx", "phoenix_prescott", "Phoenix", 16346610),
           ("g_atl", "atlanta", "Atlanta", 26366926)]


def groups(order=None):
    rows = []
    for gid, key, label, avails in (order or MARKETS):
        rows.append({
            "id": gid, "op": "AND", "name": "", "terms": AUDIENCE.split(", "),
            "geo_def": {"kind": "markets", "markets": [key]}, "resolved_markets": [key],
            "entity_id": gid, "entity_label": label, "entity_locked": False,
            "avails_monthly": avails, "avails_full_flight": avails,
            "avails_doc_monthly": avails, "avails_doc_full_flight": avails,
            "avails_doc_start": "2026-12-01", "avails_doc_end": "2026-12-31",
            "include_in_plan": False, "include_locked": False, "color": "#4C78A8",
            "color_locked": False,
        })
    return rows


DRAFT = {
    "client_name": CLIENT, "vertical": "b2b",
    "target_markets": ["Washington, DC", "Dallas", "Phoenix", "Atlanta"],
    "primary_market": "Washington, DC", "options_requested": 3,
    "total_budget": 0, "total_budget_period": "month",
    "stated_flight": {"months": 6, "start": "2026-01-01", "end": "2026-06-30", "year_stated": False},
    "breakout": "monthly", "media_plan_lines": [],
    "audience_names": {AUDIENCE: "Men 35–54, Business Owners & Executives"},
    "options": [
        {"name": "Good", "total_budget": 0, "total_budget_period": "month", "breakout": "monthly",
         "group_selection": {"mode": "named", "match": ["Washington"]},
         "group_allocation": {"impressions": 500000, "period": "month"}, "group_cpm": 34,
         "strategy_summary": "A focused DC launch reaching executives and business owners.",
         "media_plan_lines": [{"product": "streaming_retargeting_display",
                               "allocation": {"impressions": 500000, "period": "month",
                                              "stated_share": 10}}]},
        {"name": "Better", "total_budget": 0, "total_budget_range": {"low": 40000, "high": 50000},
         "total_budget_period": "month", "breakout": "monthly",
         "group_selection": {"mode": "all"},
         "group_allocations": [{"match": ["Washington"],
                                "allocation": {"impressions": 500000, "period": "month"}}],
         "group_allocation": {"proportional_to_avails": True}, "group_cpm": 34,
         "strategy_summary": "DC stays the lead market while three secondary markets launch.",
         "media_plan_lines": [{"product": "streaming_retargeting_display",
                               "allocation": {"match_impressions": "premion_streaming_tv",
                                              "stated_share": 10}}]},
        {"name": "Best", "total_budget": 0, "total_budget_period": "month", "breakout": "monthly",
         "scale_from": {"option": "Better", "total_budget": 75000, "period": "month"},
         "group_selection": {"mode": "all"}, "group_cpm": 34,
         "strategy_summary": "More weight in all four markets, with DC still leading.",
         "media_plan_lines": []},
    ],
    "total_tv": False, "show_sov": True, "audiences": [],
    "attribution": ["web", "commercial_production"], "sports": [],
    "campaign_specs": {"goals": ["Build awareness with enterprise decision-makers."],
                       "audience": ["Men 35-54 who own or run businesses."],
                       "geography": ["Primary market: Washington, DC metropolitan area.",
                                     "Secondary markets: Dallas, Phoenix and Atlanta."],
                       "budget": [], "placements": ["Premion Streaming TV and retargeting display."],
                       "timing": ["Six months, January through June."]},
    "strategy_summary": "", "review_items": [], "unresolved": [], "unresolved_internal": [],
}


class _StubSt:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply(draft, flight=(date(2027, 1, 1), date(2027, 6, 30)), group_rows=None, notes=NOTES):
    real = app.st
    stub = _StubSt()
    stub.session_state.update({
        "targeting_groups": copy.deepcopy(groups() if group_rows is None else group_rows),
        "market_choice": "DC",
        "flight_start": flight[0], "flight_end": flight[1],
        "avails_basis": app.AVAILS_BASIS_MONTHLY,
        "avails_documents": [{"name": "synthetic_avails.pdf"}],
        "draft_source_notes": notes, "client_name": CLIENT,
    })
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft), notes=notes)
    finally:
        app.st = real
    return stub.session_state


# --- what the notes state, read by this file, not by the app ----------------
def k_figure(text):
    return float(text) * 1000


STATED_CTV = k_figure(re.search(r"(\d+)K CTV impressions", NOTES).group(1))
STATED_RT = k_figure(re.search(r"(\d+)K retargeting display impressions", NOTES).group(1))
RANGE = [k_figure(x) for x in re.search(r"\$(\d+)-(\d+)K/mo", NOTES).groups()]
BEST = k_figure(re.search(r"~\$(\d+)K budget", NOTES).group(1))
CPM = float(re.search(r"CTV CPM is \$(\d+)", NOTES).group(1))
RT_CPM = 5.50   # the rate card's display retargeting rate, stated in the target deck


def option(state, name):
    return next((o for o in state.get("plan_options") or [] if o["name"] == name), None)


def monthly_cost(opt):
    return sum(r["Cost"] for r in opt["rows"])


def media_rows(opt):
    return [r for r in opt["rows"] if "Retargeting" not in r["Tactic"]]


def rt_row(opt):
    return next(r for r in opt["rows"] if "Retargeting" in r["Tactic"])


def check_plan():
    print("\nThe three options, priced from what the notes state")
    state = apply(DRAFT)
    names = [o["name"] for o in state.get("plan_options") or []]
    check("three options, in the notes' order", names == ["Good", "Better", "Best"], names)
    good, better, best = (option(state, n) for n in ("Good", "Better", "Best"))
    if not (good and better and best):
        return state

    check("Good: DC CTV is exactly the stated 500K impressions a month",
          [r["Impressions"] for r in media_rows(good)] == [STATED_CTV],
          [(r["Geo"], r["Impressions"]) for r in media_rows(good)])
    check("Good: retargeting is the stated 500K a month (not 500K / 6)",
          rt_row(good)["Impressions"] == STATED_RT, rt_row(good)["Impressions"])
    good_cost = STATED_CTV * CPM / 1000 + STATED_RT * RT_CPM / 1000   # $17,000 + $2,750
    check(f"Good: ${good_cost:,.0f} a month", round(monthly_cost(good), 2) == good_cost,
          monthly_cost(good))
    check("Good: its retargeting runs in DC only, where Good runs",
          rt_row(good)["Geo"] == media_rows(good)[0]["Geo"], rt_row(good)["Geo"])

    dc = media_rows(better)[0]
    check("Better: DC first and held at the stated 500K", dc["Impressions"] == STATED_CTV
          and "Washington" in dc["Geo"], (dc["Geo"], dc["Impressions"]))
    others = media_rows(better)[1:]
    check("Better: three secondary markets, each below DC",
          len(others) == 3 and all(r["Impressions"] < dc["Impressions"] for r in others),
          [(r["Geo"], r["Impressions"]) for r in others])
    avails = {key: a for _g, key, _l, a in MARKETS[1:]}
    imps = [r["Impressions"] for r in others]
    ratios = [i / a for i, a in zip(imps, avails.values())]
    check("Better: secondary impressions follow their avails (within 5K rounding)",
          max(ratios) - min(ratios) < 5000 / min(avails.values()) * 2, ratios)
    check("Better: whole 5K steps", all(i % 5000 == 0 for i in imps), imps)
    check("Better: retargeting matches the CTV total",
          rt_row(better)["Impressions"] == sum(r["Impressions"] for r in media_rows(better)),
          (rt_row(better)["Impressions"], sum(r["Impressions"] for r in media_rows(better))))
    check(f"Better: inside the stated ${RANGE[0]:,.0f}-${RANGE[1]:,.0f} a month",
          RANGE[0] <= monthly_cost(better) <= RANGE[1], monthly_cost(better))
    # The target deck's own figures for this shape.
    check("Better: 500K / 240K / 155K / 245K, retargeting 1.14M, $45,030",
          [r["Impressions"] for r in media_rows(better)] == [500000, 240000, 155000, 245000]
          and rt_row(better)["Impressions"] == 1140000 and round(monthly_cost(better)) == 45030,
          ([r["Impressions"] for r in better["rows"]], monthly_cost(better)))

    check(f"Best: within 1% of the stated ${BEST:,.0f} a month",
          abs(monthly_cost(best) - BEST) <= BEST * 0.01, monthly_cost(best))
    check("Best: DC still the largest market line",
          media_rows(best)[0]["Impressions"] == max(r["Impressions"] for r in media_rows(best)),
          [(r["Geo"], r["Impressions"]) for r in media_rows(best)])
    check("Best: every line up from Better by about the same factor",
          all(abs(b["Impressions"] / a["Impressions"] - BEST / RANGE[0] * RANGE[0] / monthly_cost(better))
              < 0.02 for a, b in zip(better["rows"], best["rows"])),
          [(a["Impressions"], b["Impressions"]) for a, b in zip(better["rows"], best["rows"])])
    check("Best: 830K / 400K / 260K / 410K, retargeting 1.9M, $75,050",
          [r["Impressions"] for r in media_rows(best)] == [830000, 400000, 260000, 410000]
          and rt_row(best)["Impressions"] == 1900000 and round(monthly_cost(best)) == 75050,
          ([r["Impressions"] for r in best["rows"]], monthly_cost(best)))

    for opt in (good, better, best):
        check(f"{opt['name']}: media lines first, retargeting last",
              "Retargeting" in opt["rows"][-1]["Tactic"]
              and all("Retargeting" not in r["Tactic"] for r in opt["rows"][:-1]),
              [r["Tactic"] for r in opt["rows"]])
        check(f"{opt['name']}: priced at the stated ${CPM:g} CTV rate",
              all(r["CPM"] == CPM for r in media_rows(opt)), [r["CPM"] for r in media_rows(opt)])
    check("each option keeps its own strategy line",
          len({o.get("strategy_summary") for o in (good, better, best)}) == 3
          and "multi" not in (good.get("strategy_summary") or "").lower())

    notes = state.get("draft_unresolved_internal") or []
    check("the 10%-vs-500K disagreement goes to the seller, count wins",
          any("10%" in n and "count" in n for n in notes), notes)
    check("the range's midpoint pricing is said", any("middle of the" in n for n in notes), notes)
    check("no option dropped, none missing",
          not any("priced to $0" in n or "ask for 3 options" in n for n in notes), notes)
    return state


def check_text(state):
    print("\nOne label per market, plain-language targeting, no stray geography")
    groups_by_id = {g["id"]: g for g in state.get("targeting_groups") or []}
    better = option(state, "Better")
    geos = [r["Geo"] for r in media_rows(better)]
    check("market rows use the one client-facing label",
          geos == ["Washington, DC", "Dallas, TX", "Phoenix, AZ", "Atlanta, GA"], geos)
    check("retargeting Geo uses the same labels",
          rt_row(better)["Geo"] == ", ".join(geos), rt_row(better)["Geo"])
    names = state.get("avails_audience_names") or {}
    shown = [app.client_targeting(r, groups_by_id, names) for r in media_rows(better)]
    check("Targeting reads in plain language, no geography prefix",
          set(shown) == {"Men 35–54, Business Owners & Executives"}, shown)
    totals = app.compute_plan_totals(better["rows"], better["breakout"], 6, "",
                                     groups_by_id=groups_by_id, show_entity_label=True,
                                     audience_names=names)
    check("...even with Show Label on (the labels only restate the market)",
          {r["targeting"] for r in totals["preview_rows"] if r["matched_avails_monthly"]}
          == {"Men 35–54, Business Owners & Executives"},
          [r["targeting"] for r in totals["preview_rows"]])
    check("the catalog name humanizes without a drafted name",
          avails_grid.audience_short_name(AUDIENCE) == "Men 35–54, Business Owners Executives",
          avails_grid.audience_short_name(AUDIENCE))
    geography = state.get("geography_text") or ""
    check("Geography: the drafted lines only, no market appended again",
          geography.splitlines() == DRAFT["campaign_specs"]["geography"], geography)
    budget = state.get("budget_text") or ""
    check("Budget & Allocation states each option's markets and money, from the rows",
          budget.splitlines()[0] == "Good: Washington, DC — $19,750/month ($118,500 over 6 months)"
          and "$45,030/month" in budget and "$75,050/month" in budget
          and "proportion" not in budget.lower(), budget)
    check("target DMAs in the notes' order, primary first",
          [label.split(" (")[0] for label in state.get("target_dma_choice") or []][:1]
          and "Washington" in (state.get("target_dma_choice") or [""])[0],
          state.get("target_dma_choice"))

    results = []
    for opt in state["plan_options"]:
        results.append(app.compute_plan_totals(opt["rows"], opt["breakout"], 6, "",
                                               groups_by_id=groups_by_id))
    caption = app.avails_plan_caption(state["plan_options"], results)
    check("avails caption: a per-market share, not a summed plan figure",
          caption and "under 6%" in caption and "buys" not in caption, caption)


def check_order_from_avails():
    print("\nAn avails document in a different order still plans DC first")
    shuffled = [MARKETS[3], MARKETS[1], MARKETS[2], MARKETS[0]]
    state = apply(DRAFT, group_rows=groups(shuffled))
    better = option(state, "Better")
    check("Better's first media line is DC",
          better and "Washington" in media_rows(better)[0]["Geo"],
          better and [r["Geo"] for r in better["rows"]])
    check("then the notes' order: Dallas, Phoenix, Atlanta",
          better and [r["Geo"] for r in media_rows(better)[1:]] == ["Dallas, TX", "Phoenix, AZ", "Atlanta, GA"],
          better and [r["Geo"] for r in media_rows(better)])


def check_guards():
    print("\nThe plan the app used to build trips the pre-Generate checks")
    old = copy.deepcopy(DRAFT)
    old["options"][0]["group_allocation"] = {"percent_of_avails": 100}
    old["options"][0]["media_plan_lines"][0]["allocation"] = {"flat_amount": 2750}
    old["options"][1] = {**old["options"][1], "total_budget_range": None,
                         "group_allocations": [],
                         "media_plan_lines": [{"product": "streaming_retargeting_display",
                                               "allocation": {"percent_of_total": 10}}]}
    old["options"][2] = {**old["options"][2], "scale_from": None, "total_budget": 450000,
                         "total_budget_period": "campaign",
                         "group_allocation": {"proportional_to_avails": True},
                         "media_plan_lines": [{"product": "streaming_retargeting_display",
                                               "allocation": {"percent_of_total": 10}}]}
    for opt in old["options"]:
        opt.pop("group_allocations", None) if not opt.get("group_allocations") else None
    state = apply(old)
    names = [o["name"] for o in state["plan_options"]]
    check("Better, with no budget, isn't on the plan", "Better" not in names, names)
    checks = state.get("draft_plan_checks") or {}
    check("...and is recorded, not dropped silently", "Better" in checks.get("dropped_options", []),
          checks)
    groups_by_id = {g["id"]: g for g in state["targeting_groups"]}
    results = [app.compute_plan_totals(o["rows"], o["breakout"], 6, "", groups_by_id=groups_by_id)
               for o in state["plan_options"]]
    findings = app.plan_sanity_findings(state["plan_options"], results, checks, NOTES,
                                        date(2027, 1, 1), groups_by_id=groups_by_id,
                                        as_of=date(2026, 11, 12))
    kinds = {f["key"].split(":")[0] for f in findings}
    for kind, what in (("budget", "a monthly total past 1.5x the largest figure in the notes"),
                       ("avails", "a line at 50%+ of its avails the notes didn't ask for"),
                       ("options", "3 options asked for, 2 produced"),
                       ("dropped", "the $0 option")):
        check(f"fires: {what}", kind in kinds, [f["text"] for f in findings])
    check("Best (pure proportional) puts DC below a secondary market -> fires",
          "primary" in kinds, [f["text"] for f in findings])

    good_state = apply(DRAFT)
    gbi = {g["id"]: g for g in good_state["targeting_groups"]}
    good_results = [app.compute_plan_totals(o["rows"], o["breakout"], 6, "", groups_by_id=gbi)
                    for o in good_state["plan_options"]]
    clean = app.plan_sanity_findings(good_state["plan_options"], good_results,
                                     good_state.get("draft_plan_checks"), NOTES, date(2027, 1, 1),
                                     groups_by_id=gbi, as_of=date(2026, 11, 12))
    check("the right plan trips none of them", clean == [], [f["text"] for f in clean])
    past = app.plan_sanity_findings(good_state["plan_options"], good_results,
                                    good_state.get("draft_plan_checks"), NOTES, date(2026, 1, 1),
                                    groups_by_id=gbi, as_of=date(2026, 11, 12))
    check("a flight starting before today fires", any(f["key"].startswith("past_flight") for f in past),
          [f["text"] for f in past])
    check("stated dollar figures read ranges as both ends",
          sorted(app.stated_dollar_figures("close to $40-50K/mo, then ~$75K"))
          == [40000, 50000, 75000], app.stated_dollar_figures("close to $40-50K/mo, then ~$75K"))


def check_flight_year():
    print("\nA flight with no year is the next one coming, never the past")
    start, end = app.upcoming_flight(date(2026, 1, 1), date(2026, 6, 30), NOTES, False,
                                     as_of=date(2026, 11, 12))
    check("'Jan-June' in November 2026 -> Jan 1 - Jun 30, 2027",
          (start, end) == (date(2027, 1, 1), date(2027, 6, 30)), (start, end))
    start, end = app.upcoming_flight(date(2028, 1, 1), date(2028, 6, 30), NOTES, False,
                                     as_of=date(2026, 11, 12))
    check("a year written too far ahead comes back to 2027",
          (start, end) == (date(2027, 1, 1), date(2027, 6, 30)), (start, end))
    start, _ = app.upcoming_flight(date(2026, 1, 1), date(2026, 6, 30),
                                   "Flight January 2026 through June.", None, as_of=date(2026, 11, 12))
    check("a year the notes wrote beside the month is kept", start == date(2026, 1, 1), start)
    start, _ = app.upcoming_flight(date(2026, 12, 1), date(2026, 12, 31), "December", False,
                                   as_of=date(2026, 11, 12))
    check("next month stays this year", start == date(2026, 12, 1), start)
    state = apply(DRAFT, flight=(date(2026, 12, 1), date(2026, 12, 31)))
    pending = state.get("_draft_pending_fields") or {}
    check("drafted against a Dec 2026 band, the plan takes the notes' Jan 2027 (not Jan 2026)",
          (pending.get("flight_start"), pending.get("flight_end"))
          == (date(2027, 1, 1), date(2027, 6, 30)), pending)

    print("\nThe avails date notice fires with no month in common")
    group = groups()[0]
    ranges = app.flight_month_ranges(date(2027, 1, 1), date(2027, 6, 30))
    found = app.avails_flight_divergence(group, ranges)
    check("Dec 2026 document vs Jan-Jun 2027 plan is named",
          len(found) == 1 and found[0][0] == "The whole flight", found)
    overlap = app.avails_flight_divergence(group, app.flight_month_ranges(date(2026, 12, 1),
                                                                          date(2026, 12, 31)))
    check("an exact match still says nothing", overlap == [], overlap)


def run_pending(state):
    """What the next render does: the queued band fields, then the group
    re-price (reconcile_group_plan_lines)."""
    real = app.st
    stub = _StubSt()
    stub.session_state = state
    app.st = stub
    try:
        app.apply_pending_draft_fields()
        if state.pop("_pending_group_realloc", None):
            groups_by_id = {g["id"]: g for g in state["targeting_groups"]}
            intent = state.get("draft_plan_intent") or {}
            n = len(state.get("active_months") or []) or 6
            for opt in state["plan_options"]:
                opt_intent = next(i for i in intent["options"] if i["name"] == opt["name"])
                app._allocate_group_rows(opt, opt_intent, groups_by_id, "", "", n)
        rescale = state.pop("_pending_rescale", None)
        if rescale:
            for name in rescale:
                app.rescale_option(state, name, len(state.get("active_months") or []) or 6)
            app.overlay_hand_edits([o for o in state["plan_options"] if o["name"] in rescale],
                                   state.get("_review_hand_edits"))
        app.capture_review_plan(state)
    finally:
        app.st = real
    return state


def reapply(state, answers, values=None):
    real = app.st
    stub = _StubSt()
    stub.session_state = state
    app.st = stub
    try:
        app.reapply_review_rows(state, answers, values or {})
    finally:
        app.st = real
    return run_pending(state)


def check_flight_card_reprice():
    print("\nThe notes' flight is the default; the Flight row puts the setup flight back")
    on_notes = {o["name"]: round(monthly_cost(o)) for o in apply(DRAFT)["plan_options"]}
    state = run_pending(apply(DRAFT, flight=(date(2026, 12, 1), date(2026, 12, 31))))
    landed = {o["name"]: round(monthly_cost(o)) for o in state["plan_options"]}
    check("drafted against a one-month band, every option lands where drafting over the notes' "
          "own flight does -- no row touched", landed == on_notes, (on_notes, landed))
    check("the form holds the notes' flight", (state["flight_start"], state["flight_end"])
          == (date(2027, 1, 1), date(2027, 6, 30)), (state["flight_start"], state["flight_end"]))
    shorthand = state["plan_options"][0]["_draft"]["context"]["flight_shorthand"]
    flights = {r["Flight"] for o in state["plan_options"] for r in o["rows"]}
    check(f"every row's Flight reads the notes' flight ({shorthand})", flights == {shorthand},
          flights)
    flight_row = next(i for i in state["draft_review_items"] if i.get("kind") == "flight")
    check("the Flight row is pinned", (flight_row.get("row") or {}).get("pinned") is True)
    check("its default states the notes' dates", "Jan–Jun 2027" in
          app.draft_followups.option_of(flight_row, flight_row["default"])["label"],
          flight_row["options"])

    reapply(state, {flight_row["id"]: "setup"})
    check("choosing the setup flight moves the form to Dec 2026",
          (state["flight_start"], state["flight_end"]) == (date(2026, 12, 1), date(2026, 12, 31)),
          (state["flight_start"], state["flight_end"]))
    on_setup = {o["name"]: round(monthly_cost(o)) for o in state["plan_options"]}
    check("monthly budgets hold on the one-month flight", on_setup == on_notes, (on_notes, on_setup))
    reapply(state, {})
    back = {o["name"]: round(monthly_cost(o)) for o in state["plan_options"]}
    check("back to the notes' flight, the plan is the drafted one again",
          back == on_notes and state["flight_start"] == date(2027, 1, 1), (on_notes, back))

    rows, *_rest = app.resolve_drafted_lines(
        [{"product": "premion_streaming_tv", "allocation": {"split_evenly": True}},
         {"product": app.COMMITTED_LINE_PRODUCT, "allocation": {"flat_amount": 3000, "period": "campaign"}}],
        10000, "Jan", "DC", "")
    check("a committed line holds its cost back and adds no row",
          [(r["Tactic"], r["Cost"]) for r in rows] == [("Premion Streaming TV", 7000.0)], rows)


def rows_by_kind(state):
    out = {}
    for item in app.draft_followups.rank_rows(state.get("draft_review_items") or []):
        out.setdefault(item["kind"], []).append(item)
    return out


def check_review_rows():
    print("\nThe review shows only what the app inferred or chose between (Matt, 2026-11-13)")
    state = run_pending(apply(DRAFT, flight=(date(2026, 12, 1), date(2026, 12, 31))))
    rows = rows_by_kind(state)
    ranked = app.draft_followups.rank_rows(state["draft_review_items"])
    check("the flight row is first", ranked and ranked[0]["kind"] == "flight",
          [i["kind"] for i in ranked])
    check("rows: flight, rates, the 10%-vs-count split and Better's range -- nothing else",
          set(rows) == {"flight", "rate", "share_conflict", "budget_range"}, sorted(rows))
    rate_text = [i["row"]["text"] for i in rows.get("rate", [])]
    check("one rate row, CTV's: $34 from the notes, above the $32 rate card, every option",
          rate_text == ["$34.00 from the notes, above the $32.00 rate card."]
          and rows["rate"][0]["row"]["label"] == "CTV rate", rate_text)
    check("retargeting's fixed rate is no row", not rows.get("fixed_rate"), rows.get("fixed_rate"))
    check("Better's range is the only budget row (Best's budget is stated)",
          [i["row"]["text"].split(":")[0] for i in rows.get("budget_range", [])] == ["Better"],
          [i["row"]["text"] for i in rows.get("budget_range", [])])
    check("no row shows a dollar or line-count figure of its own",
          all("impact" not in (i["row"].get("text") or "") for i in ranked))
    check("no client questions: draft_unresolved is empty", state.get("draft_unresolved") == [],
          state.get("draft_unresolved"))
    check("every item is the seller's", all(i.get("audience") == "internal"
                                            for i in state["draft_review_items"]))

    before = {o["name"]: round(monthly_cost(o), 2) for o in state["plan_options"]}
    budget = rows["budget_range"][0]
    share = rows["share_conflict"][0]
    reapply(state, {budget["id"]: "value", share["id"]: "share"}, {budget["id"]: RANGE[0]})
    better = option(state, "Better")
    rt = rt_row(better)
    check(f"Better at ${RANGE[0]:,.0f}/mo re-prices in code to within $50 of it",
          abs(monthly_cost(better) - RANGE[0]) <= 50, monthly_cost(better))
    check("retargeting at 10% of Better's budget", abs(rt["Cost"] - RANGE[0] * 0.10) <= 1, rt)
    check("DC stays at its stated 500K in Better",
          media_rows(better)[0]["Impressions"] == STATED_CTV,
          [(r["Geo"], r["Impressions"]) for r in media_rows(better)])
    check("the other options don't move", round(monthly_cost(option(state, "Good")), 2)
          == before["Good"])
    bundle_log = app.review_bundle_log({**state, "_review_rows_shown": [{"id": "x"}]})
    changed = {c["kind"]: (c["from"], c["to"]) for c in bundle_log["changes"]}
    check("the debug bundle logs what the rep changed, from and to",
          changed.get("budget_range") == ("$45,000/mo", "$40,000/mo")
          and changed.get("share_conflict", ("", ""))[1] == "10% of budget", changed)
    reapply(state, {})
    after = {o["name"]: round(monthly_cost(o), 2) for o in state["plan_options"]}
    check("putting every row back on its default restores the drafted plan exactly",
          after == before, (before, after))

    rate = rows["rate"][0]
    reapply(state, {rate["id"]: "value"}, {rate["id"]: 30.0})
    ctv = {r["CPM"] for o in state["plan_options"] for r in media_rows(o)}
    check("the CTV rate edit reaches every option's CTV lines", ctv == {30.0}, ctv)
    reapply(state, {})

    at_card = copy.deepcopy(DRAFT)
    for opt in at_card["options"]:
        opt["group_cpm"] = 32
        for line in opt["media_plan_lines"]:
            line["cpm"] = RT_CPM
    plain = rows_by_kind(apply(at_card))
    check("rates the notes state at the rate card: still the CTV row, no retargeting row; the "
          "notes' own flight matching the band is no row",
          set(plain) == {"rate", "share_conflict", "budget_range"}, sorted(plain))
    check("... and the CTV row says it's the rate card from the notes",
          [i["row"]["text"] for i in plain.get("rate", [])]
          == ["$32.00 from the notes, the rate card."], [i["row"]["text"] for i in plain["rate"]])


def check_fixed_rates():
    print("\nRetargeting, sports and AM rates are fixed: a stated rate is flagged, not applied")
    draft = copy.deepcopy(DRAFT)
    for opt in draft["options"]:
        for line in opt["media_plan_lines"]:
            line["cpm"] = 6.0
    state = run_pending(apply(draft))
    check("every retargeting line is priced at the $5.50 rate card",
          {rt_row(o)["CPM"] for o in state["plan_options"]} == {RT_CPM},
          {o["name"]: rt_row(o)["CPM"] for o in state["plan_options"]})
    fixed = rows_by_kind(state).get("fixed_rate") or []
    check("one row flags it: 'Retargeting: notes say $6.00, rate card is $5.50'",
          len(fixed) == 1 and fixed[0]["row"]["text"].startswith(
              "Retargeting") and "notes say $6.00, rate card is $5.50" in fixed[0]["row"]["text"],
          [i["row"]["text"] for i in fixed])
    if fixed:
        check("... defaulting to the rate card",
              app.draft_followups.option_of(fixed[0], fixed[0]["default"])["label"]
              == "Rate card $5.50")
        reapply(state, {fixed[0]["id"]: "notes"})
        check("taking the notes' rate applies $6.00 to the options that stated it",
              {o["name"]: rt_row(o)["CPM"] for o in state["plan_options"]}
              == {"Good": 6.0, "Better": 6.0, "Best": RT_CPM},
              {o["name"]: rt_row(o)["CPM"] for o in state["plan_options"]})


def check_share_from_notes():
    print("\nThe ~10% is read from the notes when the draft leaves it out")
    draft = copy.deepcopy(DRAFT)
    for opt in draft["options"]:
        for line in opt["media_plan_lines"]:
            line["allocation"].pop("stated_share", None)
    state = apply(draft)
    share = rows_by_kind(state).get("share_conflict") or []
    check("the share row still shows, from the notes' '~10% of the budget'",
          len(share) == 1 and "about 10% of the budget" in share[0]["row"]["text"],
          [i["row"]["text"] for i in share])
    check("... with the share Python worked out, not the model",
          share and "(about 14% of budget)" in share[0]["row"]["text"],
          [i["row"]["text"] for i in share])
    check("notes_retargeting_share reads it", app.notes_retargeting_share(NOTES) == 10.0)
    with_field = copy.deepcopy(draft)
    with_field["retargeting_share"] = 12
    share = rows_by_kind(apply(with_field, notes=NOTES.replace("~10%", "some"))).get(
        "share_conflict") or []
    check("the draft's own retargeting_share field is used when the notes' wording isn't read",
          share and "about 12% of the budget" in share[0]["row"]["text"],
          [i["row"]["text"] for i in share])


def check_hand_edits_survive():
    print("\nA hand edit to the plan grid survives a later row change (rep edits win)")
    state = run_pending(apply(DRAFT))
    rows = rows_by_kind(state)
    budget, rate = rows["budget_range"][0], rows["rate"][0]
    reapply(state, {budget["id"]: "value"}, {budget["id"]: RANGE[0]})
    # The rep types 600,000 into Good's retargeting impressions in the grid
    # (the grid's fold-back: the row, its driver, dirty).
    good = option(state, "Good")
    index = good["rows"].index(rt_row(good))
    good["rows"][index]["Impressions"] = 600000.0
    good["driver"][index] = app.DRIVER_IMPRESSIONS
    app.recompute_row(good["rows"][index], app.DRIVER_IMPRESSIONS)
    good["dirty"][index] = True
    # ... and adds a line of their own to Best.
    best = option(state, "Best")
    best["rows"].append({"Tactic": "Hand-added fee", "Flight": "Jan–Jun 2027",
                         "Geo": "Washington, DC", "Targeting": "x", "Impressions": 0.0,
                         "CPM": 0.0, "Type": app.ROW_TYPE_FLAT_FEE, "Cost": 1000.0})
    best["dirty"].append(True)
    best["driver"].append(app.DRIVER_COST)
    # Then changes another row.
    reapply(state, {budget["id"]: "value", rate["id"]: "value"},
            {budget["id"]: RANGE[0], rate["id"]: 30.0})
    rt = rt_row(option(state, "Good"))
    check("Good's retargeting keeps the rep's 600,000 impressions", rt["Impressions"] == 600000,
          rt)
    check("... priced at its rate ($3,300)", round(rt["Cost"]) == 3300, rt)
    check("the rep's own line on Best is still there",
          any(r["Tactic"] == "Hand-added fee" for r in option(state, "Best")["rows"]),
          [r["Tactic"] for r in option(state, "Best")["rows"]])
    check("and the new row change applied too: CTV at $30",
          {r["CPM"] for o in state["plan_options"] for r in media_rows(o)
           if not app.is_flat_fee_row(r)} == {30.0},
          {o["name"]: [(r["Tactic"][:22], r["CPM"]) for r in media_rows(o)]
           for o in state["plan_options"]})
    reapply(state, {})
    check("every row back on its default: the rep's edit still stands",
          rt_row(option(state, "Good"))["Impressions"] == 600000)


def check_rescale():
    print("\nBetter's budget changes; Best was scaled from it")
    state = run_pending(apply(DRAFT))
    budget = rows_by_kind(state)["budget_range"][0]
    check("Better's budget row knows Best was scaled from it", budget["row"].get("scaled") == ["Best"],
          budget["row"])
    best_before = round(monthly_cost(option(state, "Best")))
    reapply(state, {budget["id"]: "value"}, {budget["id"]: RANGE[0]})
    check("Best isn't re-scaled on its own", round(monthly_cost(option(state, "Best"))) == best_before)
    state["draft_review_rescale"] = ["Best"]
    reapply(state, {budget["id"]: "value"}, {budget["id"]: RANGE[0]})
    best = option(state, "Best")
    check(f"'Re-scale Best' scales it from Better's new plan to its own ~${BEST:,.0f}",
          abs(monthly_cost(best) - BEST) <= 200, monthly_cost(best))
    better_dc = media_rows(option(state, "Better"))[0]["Impressions"]
    check("... keeping Better's weighting (DC first and largest)",
          media_rows(best)[0]["Impressions"] == max(r["Impressions"] for r in media_rows(best))
          and better_dc == STATED_CTV, [(r["Geo"], r["Impressions"]) for r in media_rows(best)])
    reapply(state, {})
    check("... with its lines re-weighted (DC no longer the drafted 830K)",
          media_rows(best)[0]["Impressions"] != 830000, media_rows(best)[0]["Impressions"])
    check("back on the default budget, Best re-scales to the drafted $75,050",
          round(monthly_cost(option(state, "Best"))) == best_before,
          round(monthly_cost(option(state, "Best"))))


def check_app_changes():
    print("\nApp changes to the plan are rows; a line left with no audience blocks Generate")
    draft = {"client_name": CLIENT, "total_budget": 10000, "total_budget_period": "campaign",
             "breakout": "monthly",
             "media_plan_lines": [
                 {"product": "premion_streaming_tv", "label": "Owners",
                  "segments": ["NOT A REAL SEGMENT QA"], "audience_track": "owners",
                  "allocation": {"percent_of_total": 100}}],
             "audiences": [], "attribution": [], "sports": [], "unresolved_internal": [
                 "Production costs are TBD -- confirm with the client before sending.",
                 "Used one line for the whole audience, since the notes describe one group."]}
    state = apply(draft, group_rows=[], flight=(date(2027, 1, 1), date(2027, 3, 31)),
                  notes="QA-TEST client, $10,000 for owners.")
    changed = rows_by_kind(state).get("app_change") or []
    check("the dropped segment is a row",
          any("NOT A REAL SEGMENT QA" in i["row"]["text"] and "no audience" in i["row"]["text"]
              for i in changed), [i["row"]["text"] for i in changed])
    check("its default is what was done (nothing to pick)",
          changed and len(changed[0]["options"]) == 1)
    findings = app.plan_sanity_findings(state["plan_options"],
                                        [{"monthly_cost": 0}] * len(state["plan_options"]),
                                        state.get("draft_plan_checks"), "", date(2027, 1, 1),
                                        as_of=date(2026, 11, 12))
    check("a line left with no audience blocks Generate",
          any(f["key"].startswith("no_audience:") for f in findings), [f["text"] for f in findings])
    notes = state.get("draft_model_notes") or []
    check("a to-do note ('TBD', 'confirm with the client') never reaches the review",
          not any("TBD" in n for n in notes), notes)
    check("a note of what the draft did does", any("one line" in n for n in notes), notes)


def check_no_logo():
    print("\nNo logo: the logo picture is removed, never the placeholder")
    from pptx import Presentation
    from pptx.util import Inches
    from PIL import Image
    image = io.BytesIO()
    Image.new("RGB", (8, 8), "white").save(image, format="PNG")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    image.seek(0)
    pic = slide.shapes.add_picture(image, Inches(1), Inches(1))
    pic.name = "CLIENT_LOGO"
    assembly.swap_named_picture_everywhere(prs, "CLIENT_LOGO", None)
    check("the CLIENT_LOGO picture is gone",
          not [s for s in slide.shapes if s.name == "CLIENT_LOGO"], [s.name for s in slide.shapes])


def check_dropped_option_without_budget():
    print("\nAn option that prices to $0 is a blocking item, never silent")
    draft = copy.deepcopy(DRAFT)
    draft["options"][1]["total_budget_range"] = None
    draft["options"][1]["group_allocations"] = []
    draft["options"][2]["scale_from"] = {"option": "Better", "total_budget": 75000, "period": "month"}
    state = apply(draft)
    checks = state.get("draft_plan_checks") or {}
    check("Better recorded as dropped", "Better" in (checks.get("dropped_options") or []), checks)
    check("Best, scaled from a missing option, recorded too",
          "Best" in (checks.get("dropped_options") or []), checks)
    notes = state.get("draft_unresolved_internal") or []
    check("the seller is told Generate is held", any("Generate is held" in n for n in notes), notes)


def check_debug_bundle(state):
    print("\nThe debug bundle")
    real = app.st
    stub = _StubSt()
    stub.session_state.update(state)
    stub.session_state["claude_calls"] = [{"label": "draft", "response": "{}", "stop_reason": "end_turn"}]
    app.st = stub
    db_active = db.active_deck_version
    db.active_deck_version = lambda: ({"id": 99}, None)
    real_capture = app.capture_feedback_state
    app.capture_feedback_state = lambda page: {"page": page}
    try:
        bundle = app.build_debug_bundle("manual")
        text = app.debug_bundle_bytes("manual").decode("utf-8")
    finally:
        app.st = real
        db.active_deck_version = db_active
        app.capture_feedback_state = real_capture
    for key in ("build_stamp", "deck_version_id", "user", "notes", "avails_documents",
                "targeting_groups", "model_calls", "cards", "form_before_draft", "form_after_draft",
                "form_now", "plan", "review", "unresolved_internal", "model_notes", "sanity_checks",
                "last_generate", "draft_plan_checks"):
        check(f"carries {key}", key in bundle)
    plan = bundle.get("plan") or []
    check("plan rows carry impressions, CPM, cost and % of avails",
          isinstance(plan, list) and plan and {"monthly_impressions", "cpm", "monthly_cost",
                                               "percent_of_avails"} <= set(plan[0]["rows"][0]),
          plan if not isinstance(plan, list) else plan[:1])
    check("no secrets", "ANTHROPIC" not in text and "APP_PASSWORD" not in text)
    json.loads(text)
    check("valid JSON", True)


def main():
    state = check_plan()
    check_text(state)
    check_order_from_avails()
    check_guards()
    check_flight_year()
    check_flight_card_reprice()
    check_review_rows()
    check_fixed_rates()
    check_share_from_notes()
    check_hand_edits_survive()
    check_rescale()
    check_app_changes()
    check_no_logo()
    check_dropped_option_without_budget()
    check_debug_bundle(state)
    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("All checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
