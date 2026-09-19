"""Tier 2 -- the live model-behaviour suite for ATTRIBUTION_REPORT_PLAN.md
Phase 4 (the model-facing half of the Attribution Report Builder), rewritten
for the Highlights/Takeaways rework. Costs money; run it yourself.

    ! python tests/test_attribution_draft_live.py
    ! python tests/test_attribution_draft_live.py mw          # one scenario
    ! python tests/test_attribution_draft_live.py --save

Same split as tests/test_draft_live.py (Tier 2) vs. tests/test_draft_
regression.py (Tier 1): this file asks one question -- **does the model
still connect visitor intent to the stated goal, hold the facts-only
contract, and follow the rework's new rules (threads, benchmark, trend,
plan-vs-actual, conversion definition) on the real datasets?** Assertions
are structural (a class is named, a real number is cited, no fabricated
number slips through `app.apply_attr_draft`'s own check) -- never exact
wording, which varies run to run.

Real fixtures (MW's/Cardinal's own exports) are gitignored, same convention
as test_report_assembly.py / test_wideorbit.py -- SKIPs rather than fails
when they're absent.

WAEPA now uses file "(14)" (single RFPID-258341, Mar-Jun 2026 flight), not
"(13)" -- "(13)" is a DIFFERENT real campaign (two RFPIDs, Jul-Aug 2026) that
test_report_assembly.py's/test_attribution_import.py's own multi-RFPID
regression checks are legitimately built around and must keep using; "(14)"
is the one whose numbers (0.13% attributed rate, 532 unique visitors over
2,220,083 delivered) actually match a previously-generated real report --
see ATTRIBUTION_REPORT_PLAN.md's dated section for the full arithmetic that
settled this.
"""
import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"

import streamlit as st          # noqa: E402
import app                       # noqa: E402
import attribution_import as ai  # noqa: E402
import market_lookup             # noqa: E402
import polk_import as pk         # noqa: E402
import report_assembly as ra     # noqa: E402

market_lookup.install()

ATTRIBUTION_MW = REPO / "MW attribution excel.xlsx"
DELIVERY_MW = REPO / "MW delivery.xlsx"
ATTRIBUTION_CARDINAL = REPO / "Premion Website Attribution Cardinal.xlsx"
DELIVERY_CARDINAL = REPO / "Premion OTT.xlsx"
ATTRIBUTION_WAEPA = REPO / "Premion Website Attribution and Reach Extension (14).xlsx"
POLK_DASHBOARD = REPO / "Polk Dashboard.xlsx"


class Report:
    def __init__(self):
        self.passed, self.failed, self.skipped = 0, [], []
        self.scenario = ""

    def check(self, label, ok, detail=""):
        if ok:
            self.passed += 1
            print(f"    PASS  {label}")
        else:
            self.failed.append(f"[{self.scenario}] {label}")
            print(f"    FAIL  {label}{'  -- ' + str(detail) if detail else ''}")

    def skip(self, reason):
        print(f"    SKIP  {reason}")
        self.skipped.append(reason)


def _all_draft_text(draft):
    """Every string value anywhere in the draft, joined -- used for a loose
    "did this number show up ANYWHERE" check. The model is free to place a
    real fact in whichever field is the natural home for it on a given
    run (a highlight bullet one time, the URL narrative another) -- wording
    varies run to run by design (see this file's own docstring), so a
    presence check should not hard-code which field."""
    out = []

    def walk(value):
        if isinstance(value, dict):
            for v in value.values():
                walk(v)
        elif isinstance(value, list):
            for item in value:
                walk(item)
        elif isinstance(value, str):
            out.append(value)

    walk(draft)
    return " ".join(out)


def check_facts_only(rep, draft, facts):
    kwargs, warnings = app.apply_attr_draft(draft, facts)
    # `warnings` also carries the model's own `goal_alignment_notes` -- a
    # legitimate signal, not a defect (e.g. "no data exists to quantify
    # this note"), so the facts-only check re-derives the number-
    # fabrication violations specifically rather than treating a non-empty
    # combined list as a failure.
    highlight_bullets, takeaway_bullets, _whats_next = ra.distribute_threads(draft.get("threads"))
    draft_texts = [("highlight bullet", f"{head} {detail}") for head, detail in highlight_bullets]
    draft_texts += [("takeaway bullet", f"{head} {detail}") for head, detail in takeaway_bullets]
    for key in ("attribution_headline_note", "url_headline_note", "zip_headline_note",
               "attribution_narrative", "url_intent_narrative", "zip_narrative",
               "delivery_narrative", "delivery_breakdown_narrative"):
        if draft.get(key):
            draft_texts.append((key, draft[key]))
    violations = app._attr_draft_number_violations(draft_texts, facts)
    rep.check("zero fabricated-number violations (the facts-only contract held)",
             violations == [], violations)
    if draft.get("goal_alignment_notes"):
        print(f"    ....  goal_alignment_notes: {draft['goal_alignment_notes']}")
    return kwargs, highlight_bullets, takeaway_bullets


def check_highlights_no_tile_restatement(rep, highlight_bullets, facts):
    """2026-09-08 follow-up: a highlight bullet's HEAD may not restate a
    tile value (impressions delivered, attributed unique visitors,
    attributed rate, attributed conversions) -- MW's own first pass at
    this slide put three of its four bullets' headline numbers on exactly
    these three figures, which the tiles already show two inches above."""
    headline = facts.get("headline") or {}
    forbidden = {
        f"{headline.get('delivered_impressions', 0):,}",
        f"{headline.get('attributed_unique_visitors', 0):,}",
        f"{(headline.get('attributed_rate') or 0) * 100:.2f}%",
    }
    conversions = facts.get("conversions")
    if conversions:
        forbidden.add(f"{conversions.get('attributed', 0):,}")
    heads = [head for head, _detail in highlight_bullets]
    violations = [h for h in heads if any(n in h for n in forbidden)]
    rep.check("no highlight bullet's headline restates a tile value",
             not violations, {"violations": violations, "forbidden": sorted(forbidden)})


def check_highlight_cites_fact(rep, highlight_bullets, label, count, share):
    """At least one HIGHLIGHT bullet (not just somewhere in the deck) must
    cite the goal-relevant class's own number -- Matt's own ruling: "a
    client reading one slide should get it," not only the takeaways."""
    haystack = " ".join(f"{head} {detail}" for head, detail in highlight_bullets)
    share_pct = f"{share * 100:.1f}"
    rep.check(f"a highlight bullet cites the {label} fact (count {count} or share {share_pct}%)",
             str(count) in haystack or share_pct in haystack,
             [head for head, _d in highlight_bullets])


def check_goal_intent_connection(rep, draft, goal_word, intent_class):
    """The URL slide is the point of Phase 4: with a real goal supplied,
    the narrative should name the intent class that maps to it, with a
    real number attached -- not just describe the mix in general."""
    narrative = (draft.get("url_intent_narrative") or "").lower()
    rep.check(f"names the {intent_class!r} intent class (goal: {goal_word!r})",
             intent_class.replace("_", " ") in narrative or intent_class in narrative
             or any(w in narrative for w in goal_word.split()),
             draft.get("url_intent_narrative"))
    rep.check("cites at least one real number (not empty prose)",
             any(ch.isdigit() for ch in narrative), draft.get("url_intent_narrative"))


def check_no_delivery_metric_thread(rep, highlight_bullets, takeaway_bullets):
    """The rework's own retired allowance: delivery metrics (VCR, CTV
    share, frequency, publisher/creative mix) may appear in a thread's
    finding/meaning ONLY when a stated goal names them. Neither MW's nor
    Cardinal's goals mention anything delivery-shaped, so none should
    appear in any thread text (the dedicated Delivery Recap/Breakdown
    slide's own narrative is untouched by this rule and not checked here)."""
    haystack = " ".join(f"{h} {d}" for h, d in highlight_bullets + takeaway_bullets).lower()
    forbidden = ("video completion", "vcr", "ctv share", "frequency", "publisher mix")
    violations = [w for w in forbidden if w in haystack]
    rep.check("no delivery-metric language in any thread (no goal here names one)",
             not violations, {"violations": violations, "haystack": haystack[:600]})


def check_no_benchmark_citation(rep, highlight_bullets, takeaway_bullets, facts):
    """facts.benchmark is None (below every candidate row) for all three
    real fixtures on hand -- proves the "below benchmark: silence" rule,
    not just that the field is technically absent from the schema."""
    rep.check("facts.benchmark is genuinely None for this real fixture",
             facts.get("benchmark") is None, facts.get("benchmark"))
    haystack = " ".join(f"{h} {d}" for h, d in highlight_bullets + takeaway_bullets).lower()
    rep.check("no 'benchmark' language anywhere in the drafted threads",
             "benchmark" not in haystack, haystack[:600])


def run_mw(rep, save):
    rep.scenario = "MW wrap -- in-store visits"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    for path in (ATTRIBUTION_MW, DELIVERY_MW):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_MW))
    goals = ["Drive in-store visits to Mattress Warehouse locations"]
    # "retail" -> Furniture Retail's benchmark row (5.76%/0.73%) -- MW's
    # real rate (1.363%/0.198%) sits well below it, so no citation should
    # fire; this exercises the real lookup, not a vertical-less no-op.
    facts = ra.build_facts_payload(attribution, delivery, goals=goals, vertical="retail")
    intent = ra.intent_facts(attribution)
    store_visit = next((c for c in intent["classes"] if c["intent"] == "store_visit"), None)
    rep.check("MW's own data has a real store-visit intent class to find",
             store_visit is not None and store_visit["visits"] > 0, store_visit)

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        out = FIXTURES / "attr_mw.live.json"
        import json
        out.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        print(f"    ....  saved raw response to {out.relative_to(REPO)}")

    _kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)
    check_goal_intent_connection(rep, draft, "in-store visits", "store_visit")
    if store_visit:
        haystack = _all_draft_text(draft)
        share_pct = f"{store_visit['share'] * 100:.1f}"
        rep.check(f"cites the real store-visit fact (count {store_visit['visits']} or "
                 f"share {share_pct}%) somewhere in the drafted content",
                 str(store_visit["visits"]) in haystack or share_pct in haystack,
                 haystack[:600])
        check_highlight_cites_fact(rep, highlight_bullets, "store-visit", store_visit["visits"],
                                   store_visit["share"])
    check_highlights_no_tile_restatement(rep, highlight_bullets, facts)
    check_no_delivery_metric_thread(rep, highlight_bullets, takeaway_bullets)
    check_no_benchmark_citation(rep, highlight_bullets, takeaway_bullets, facts)


def run_cardinal(rep, save):
    rep.scenario = "Cardinal monthly -- service calls"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    for path in (ATTRIBUTION_CARDINAL, DELIVERY_CARDINAL):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_CARDINAL))
    delivery = ai.parse_delivery_export(str(DELIVERY_CARDINAL))
    goals = ["Generate service calls for plumbing and HVAC work"]
    # "home_improvement" -> Home Services/Improvement's row (3.56%/0.41%) --
    # Cardinal's real rate (0.185%/0.018%) sits well below it.
    facts = ra.build_facts_payload(attribution, delivery, goals=goals, vertical="home_improvement")
    intent = ra.intent_facts(attribution)
    lead = next((c for c in intent["classes"] if c["intent"] == "lead"), None)
    rep.check("Cardinal's own data has a real lead-intent class to find",
             lead is not None and lead["visits"] > 0, lead)

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        out = FIXTURES / "attr_cardinal.live.json"
        import json
        out.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        print(f"    ....  saved raw response to {out.relative_to(REPO)}")

    _kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)
    check_goal_intent_connection(rep, draft, "service calls", "lead")
    if lead:
        haystack = _all_draft_text(draft)
        share_pct = f"{lead['share'] * 100:.1f}"
        rep.check(f"cites the real lead-intent fact (count {lead['visits']} or share "
                 f"{share_pct}%) somewhere in the drafted content",
                 str(lead["visits"]) in haystack or share_pct in haystack,
                 haystack[:600])
        check_highlight_cites_fact(rep, highlight_bullets, "lead-intent", lead["visits"], lead["share"])
    check_highlights_no_tile_restatement(rep, highlight_bullets, facts)
    check_no_delivery_metric_thread(rep, highlight_bullets, takeaway_bullets)
    check_no_benchmark_citation(rep, highlight_bullets, takeaway_bullets, facts)


def run_cardinal_trend(rep, save):
    """The account's own trend (facts.prior_periods): Cardinal's real
    current numbers, plus a hand-built EARLIER period with a visibly lower
    attributed rate/top-intent share -- does the model build a trend
    thread when the numbers genuinely moved? This only exercises the
    model-facing half (report_assembly.build_facts_payload's
    prior_periods pass-through, and the prompt's own trend rule) -- it
    does NOT touch the database; a real prior report is never logged
    against production Supabase from a test (see test_draft_regression.py's
    own stated rule)."""
    rep.scenario = "Cardinal, second report -- a real prior period with a positive trend"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    if not (ATTRIBUTION_CARDINAL.exists() and DELIVERY_CARDINAL.exists()):
        rep.skip("Cardinal fixtures not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_CARDINAL))
    delivery = ai.parse_delivery_export(str(DELIVERY_CARDINAL))
    current = ra.report_headline_facts(attribution, delivery)
    # A synthetic "first period" clearly worse than the current, real one --
    # half the attributed rate, an earlier date -- so a real trend exists
    # for the model to find, without needing a second real export on hand.
    earlier_period = dict(current)
    earlier_period["period_start"] = "2025-10-01"
    earlier_period["period_end"] = "2025-12-31"
    earlier_period["attributed_rate"] = current["attributed_rate"] / 2
    earlier_period["attributed_unique_visitors"] = current["attributed_unique_visitors"] // 2
    goals = ["Generate service calls for plumbing and HVAC work", "Improve on last quarter's performance"]
    facts = ra.build_facts_payload(attribution, delivery, goals=goals,
                                   prior_periods=[earlier_period, current])
    rep.check("facts payload carries both periods, oldest first",
             facts["prior_periods"] == [earlier_period, current], facts["prior_periods"])

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        import json
        (FIXTURES / "attr_cardinal_trend.live.json").write_text(json.dumps(draft, indent=2),
                                                                encoding="utf-8")

    _kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)
    haystack = " ".join(f"{h} {d}" for h, d in highlight_bullets + takeaway_bullets).lower()
    trend_words = ("trend", "prior", "previous", "last period", "improved", "increase", "grown",
                  "compared to", "growth", "quarter")
    rep.check("a trend thread (or trend language) appears somewhere given a real, positive "
             "prior-period comparison and a goal that asks about it",
             any(w in haystack for w in trend_words), haystack[:600])


def run_mw_optimization(rep, save):
    """The optimization engine (attribution-module-framework.md 4-6,
    ATTRIBUTION_REPORT_PLAN.md Phase 6): does a real, qualifying ZIP
    candidate become a thread action phrased as remove/reduce/reallocate,
    citing the SAME numbers report_assembly.optimization_candidates()
    computed -- never new arithmetic, never a growth recommendation?

    Same shape as run_cardinal_trend: MW's real attribution data plus a
    SYNTHETIC second period (a bare 1% baseline, marked as such) so the
    timing gate opens without a second real report on hand. Moderate level,
    every default-on dimension -- MW's real zip spread reliably produces at
    least one material outlier at this level (confirmed offline in
    tests/test_optimization_engine.py's own real-fixture pass).
    """
    rep.scenario = "MW, optimization candidates -- a real outlier becomes a subtraction-only action"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    if not ATTRIBUTION_MW.exists():
        rep.skip(f"{ATTRIBUTION_MW.name} not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_MW)) if DELIVERY_MW.exists() else None
    # SYNTHETIC prior period -- only opens the timing gate (evidence_
    # periods >= 2); no real second MW report exists on hand. Marked as
    # such rather than skipped, same ruling this file's own run_cardinal_
    # trend already applies to its synthetic earlier period.
    synthetic_prior = [{"period_start": "2026-01-01", "period_end": "2026-01-31",
                        "attributed_rate": 0.01}]
    optimizations = ra.optimization_candidates(
        attribution, "moderate", list(ra.OPTIMIZATION_DIMENSIONS_DEFAULT_ON),
        prior_periods=synthetic_prior)
    rep.check("the engine actually produced candidates on real MW data (otherwise this test "
             "proves nothing)", len(optimizations["candidates"]) > 0, optimizations["candidates"])
    if not optimizations["candidates"]:
        return

    goals = ["Drive in-store visits for the fall mattress sale",
            "Improve efficiency of underperforming zips"]
    facts = ra.build_facts_payload(attribution, delivery, goals=goals,
                                   prior_periods=synthetic_prior, optimizations=optimizations)

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        import json
        (FIXTURES / "attr_mw_optimization.live.json").write_text(json.dumps(draft, indent=2),
                                                                  encoding="utf-8")

    _kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)

    candidate_values = {str(c["value"]) for c in optimizations["candidates"]}
    watch_values = {str(c["value"]) for c in optimizations["watch_list"]}
    all_action_text = " ".join(str(t.get("action") or "") for t in (draft.get("threads") or []))
    rep.check("at least one thread's action names a real optimization candidate's value",
             any(v in all_action_text for v in candidate_values), all_action_text[:600])

    subtraction_words = ("reduce", "remove", "reallocat", "cut", "pull back", "pause", "lower")
    growth_words = ("add ", "increase ", "expand ", "grow ", "more spend", "boost ")
    action_lower = all_action_text.lower()
    rep.check("the action uses subtraction/reallocation language",
             any(w in action_lower for w in subtraction_words), all_action_text[:600])

    # Growth language is fine ELSEWHERE in the same action -- the framework's
    # own philosophy is "we remove what's clearly failing so those
    # impressions flow to what's already working" (§4), and a real run
    # confirmed the model does exactly that: naming a genuinely strong ZIP
    # from facts["zip"]["rows"] (never a candidate/watch_list value) as
    # where the freed budget goes, phrased as "Expand ... in ZIPs 27265 and
    # 28304." That's the framework working, not a violation. What subtraction-
    # only actually forbids is growth language attached to a CANDIDATE'S OWN
    # value -- checked per sentence, since the whole action routinely mixes
    # a cut clause and a reallocation-destination clause in one paragraph.
    for sentence in all_action_text.replace("\n", " ").split(". "):
        sentence_lower = sentence.lower()
        mentioned_candidates = [v for v in candidate_values if v in sentence]
        if not mentioned_candidates:
            continue
        rep.check(f"a sentence naming candidate value(s) {mentioned_candidates} never uses "
                 f"growth language for them (subtraction-only, framework §4)",
                 not any(w in sentence_lower for w in growth_words), sentence)

    # A watch_list value mentioned at all must be framed as "worth
    # watching," never as a second cut -- checked loosely (it may not be
    # mentioned at all, which is also fine) rather than requiring it.
    mentioned_watch = [v for v in watch_values if v in all_action_text]
    if mentioned_watch:
        rep.check("a mentioned watch_list value reads as 'watching', not a second cut",
                 any(w in action_lower for w in ("watch", "worth watching", "not yet")),
                 all_action_text[:600])


def run_no_goals(rep, save):
    """Edge case: no goals supplied at all. The narrative should describe
    the intent mix (naming a real class and a real number) without
    claiming it aligns to anything -- the model has nothing to align to."""
    rep.scenario = "MW, no goals -- describes the mix, claims no alignment"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    if not ATTRIBUTION_MW.exists():
        rep.skip(f"{ATTRIBUTION_MW.name} not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_MW)) if DELIVERY_MW.exists() else None
    facts = ra.build_facts_payload(attribution, delivery, goals=[])

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        out = FIXTURES / "attr_mw_no_goals.live.json"
        import json
        out.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        print(f"    ....  saved raw response to {out.relative_to(REPO)}")

    check_facts_only(rep, draft, facts)
    narrative = (draft.get("url_intent_narrative") or "").lower()
    rep.check("still cites a real number (describes the mix, doesn't go silent)",
             any(ch.isdigit() for ch in narrative), narrative)
    alignment_claims = ("against a goal", "aligns with the goal", "supports the goal",
                        "in line with the goal")
    rep.check("does NOT claim alignment to a goal that was never supplied",
             not any(phrase in narrative for phrase in alignment_claims), narrative)


def run_waepa_proposal_link(rep, save):
    """Phase 5 (proposal-linked facts) + the Highlights/Takeaways rework,
    against the REAL WAEPA proposal's own stored goals (fetched live once
    while writing this scenario, hand-carried here rather than fetched at
    test time so this file stays a pure offline-except-the-model-call test
    like its siblings -- proposal 09e61e0b-a945-4022-851d-d3c31b2acbd0,
    client_name "WAEPA", vertical "banking").

    A frequency target (3-5) and a delivery figure of 3.95 are added via a
    hand-built synthetic DeliveryExport -- no real WAEPA delivery export
    exists in this repo (confirmed by search) -- so the "delivery metric
    only when a goal names it" rule has something real to prove against:
    frequency IS named by a goal here, so (unlike MW/Cardinal above) it
    SHOULD surface in a thread. Also proves plan-vs-actual stays silent
    with the toggle off (facts.plan_vs_actual not supplied), and exercises
    it ONCE more with a synthetic plan_vs_actual dict supplied, to prove
    the model only mentions it when a goal asks (no goal here asks about
    pacing against plan, so it should NOT become a thread).
    """
    rep.scenario = "WAEPA (14), proposal-linked -- real goals, a synthetic frequency goal"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    if not ATTRIBUTION_WAEPA.exists():
        rep.skip(f"{ATTRIBUTION_WAEPA.name} not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_WAEPA))
    # The two real, stored campaign_specs.goals, plus a frequency target --
    # WAEPA's own real ask, per this account's own real media plan.
    goals = [
        "Evaluate OTT performance beyond awareness -- measure whether concentrating spend into "
        "fewer core federal markets and increasing share of voice drives more meaningful "
        "engagement signals",
        "Track engaged visits to site and application entry point, cost per engaged visit, "
        "application activity within test markets, and overall lift versus 2025 historical "
        "performance",
        "Maintain average frequency between 3 and 5 across the flight",
    ]
    synthetic_delivery = ai.DeliveryExport(
        source_name="synthetic (no real WAEPA delivery export exists)",
        delivered_impressions=attribution.delivered_impressions,
        vcr=0.98, frequency=3.95, uniques=attribution.attributed_unique_visitors * 20,
        by_geo=[("Washington, DC", int(attribution.delivered_impressions * 0.6)),
               ("Baltimore", int(attribution.delivered_impressions * 0.4))])
    facts = ra.build_facts_payload(
        attribution, synthetic_delivery, goals=goals,
        proposal_flight_label="Oct 2026 - Dec 2026",
        proposal_geography_label="Washington, DC, Baltimore",
        vertical="banking", budget=74970.0)
    rep.check("facts payload carries the proposal's own flight/geography for comparison",
             facts["proposal"] == {"flight_label": "Oct 2026 - Dec 2026",
                                   "geography_label": "Washington, DC, Baltimore"},
             facts["proposal"])
    rep.check("facts.benchmark is None -- WAEPA's real rate (0.13%/0.024%) sits well below "
             "Banking and Finance's row (4.81%/0.74%)", facts["benchmark"] is None, facts["benchmark"])
    rep.check("plan_vs_actual is None -- the toggle wasn't supplied", facts["plan_vs_actual"] is None)

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        out = FIXTURES / "attr_waepa_proposal_link.live.json"
        import json
        out.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        print(f"    ....  saved raw response to {out.relative_to(REPO)}")

    kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)
    threads = draft.get("threads") or []
    rep.check("at least 2 goal-anchored threads exist", sum(1 for t in threads
             if isinstance(t, dict) and t.get("anchor") == "goal") >= 2, threads)
    haystack_all = " ".join(f"{h} {d}" for h, d in highlight_bullets + takeaway_bullets).lower()
    rep.check("frequency (3.95, the ONE goal-named delivery metric here) surfaces in a thread",
             "3.95" in haystack_all or "frequency" in haystack_all, haystack_all[:800])
    rep.check("no benchmark language (facts.benchmark is None)",
             "benchmark" not in haystack_all, haystack_all[:800])
    plan_phrases = ("% of plan", "against plan", "vs. plan", "vs plan", "planned vs", "pacing against")
    plan_violations = [p for p in plan_phrases if p in haystack_all]
    rep.check("no plan-vs-actual language (facts.plan_vs_actual is None, no goal asks about pacing)",
             not plan_violations, (plan_violations, haystack_all[:800]))
    # Every what's-next item should trace back to SOME thread's own action
    # -- distribute_threads' own contract (no orphans, since Python derives
    # what's-next directly from the threads array).
    _hi, _ta, whats_next = ra.distribute_threads(threads)
    thread_actions = {str(t.get("action") or "").strip() for t in threads if isinstance(t, dict)}
    rep.check("every what's-next item traces to a real thread action (no orphans possible "
             "by construction, but confirm the model didn't return a stray non-dict thread)",
             all(item in thread_actions for item in whats_next), (whats_next, thread_actions))

    kwargs2, _warnings2 = app.apply_attr_draft(draft, facts)
    rep.check("apply_attr_draft's kwargs never carry a flight/geography override -- those "
             "are always Python-filled from the proposal, never the model's output",
             "flight_label" not in kwargs2 and "geography_names_override" not in kwargs2
             and "geography_label_override" not in kwargs2, sorted(kwargs2.keys()))


def run_synthetic_above_benchmark(rep, save):
    """Synthetic (not a real client) -- proves the benchmark citation rule
    actually FIRES when a campaign clears its row, cites exactly once, and
    phrases it as a comparison rather than restating the benchmark's own
    raw percentage (which would itself be caught as a facts-only violation,
    since "benchmark" is excluded from the traced-number set)."""
    rep.scenario = "Synthetic -- above-benchmark citation fires exactly once"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    # Legal's row: impression_rate 2.22%, visitor_rate 0.23%. This campaign
    # clears both, comfortably.
    facts = {
        "goals": ["Generate qualified personal injury leads"],
        "notes": "",
        "headline": {"delivered_impressions": 500000, "attributed_unique_visitors": 1500,
                    "attributed_unique_visitor_rate": 0.003, "attributed_rate": 0.035},
        "audience": {"top": None, "rows": []},
        "market": {"count": 1, "top": None, "rows": []},
        "creative": {"top": None, "rows": []},
        "breakdown": {"dimension_forced": None, "audience_available": False, "creative_available": False},
        "intent": {"classes": [
            {"intent": "lead", "label": "Lead intent", "visits": 900, "share": 0.6},
            {"intent": "other", "label": "Other pages", "visits": 600, "share": 0.4},
        ], "noise_visits": 0},
        "top_pages": [{"label": "Contact Us", "visitors": "900", "share": "60%"}],
        "zip": {"baseline_rate": 0.035, "rows": []},
        "delivery": None, "live_sports": None, "conversions": None, "budget": None, "proposal": None,
        "vertical": "legal",
        "benchmark": {"vertical": "Legal",
                      "impression_rate": {"campaign": 0.035, "cleared": True},
                      "visitor_rate": {"campaign": 0.003, "cleared": True}},
        "conversion_definition": None,
        "prior_periods": [],
        "plan_vs_actual": None,
    }
    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        import json
        (FIXTURES / "attr_synthetic_above_benchmark.live.json").write_text(
            json.dumps(draft, indent=2), encoding="utf-8")

    _kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)
    all_text = " ".join(f"{h} {d}" for h, d in highlight_bullets + takeaway_bullets)
    lower = all_text.lower()
    rep.check("cites the comparison at least once ('benchmark' named)", "benchmark" in lower, all_text)
    # "At most one report-wide" means one THREAD, not one STRING occurrence
    # -- a single benchmark thread legitimately contributes both a
    # highlight (its finding) and a takeaway (its meaning), so "benchmark"
    # can appear twice in the combined text for that one thread alone.
    # Grouped by HEAD (the same head ties a thread's highlight/takeaway
    # pair together, per distribute_threads), not by raw substring count.
    benchmark_heads = {head for head, detail in highlight_bullets + takeaway_bullets
                       if "benchmark" in head.lower() or "benchmark" in detail.lower()}
    rep.check("exactly ONE distinct thread (by head) mentions benchmark -- appearing on both "
             "highlights and takeaways for that one thread is expected, not a double-citation",
             len(benchmark_heads) == 1, benchmark_heads)
    rep.check("never quotes the benchmark row's own raw percentage (2.22% or 0.23%) -- only "
             "the campaign's own real rate (3.5%/0.30%) may appear",
             "2.22" not in all_text and "0.23%" not in all_text, all_text)
    rep.check("names 'legal' campaigns (the vertical) in the comparison",
             "legal" in lower, all_text)


def run_mw_with_polk(rep, save):
    """Phase 7 -- Polk automotive match-back, cross-paired with MW's real
    attribution file on purpose (no real attribution export exists for the
    same automotive client Polk's fixture describes; this file's own
    header rule is never to fabricate a fixture, and pairing two genuinely
    real exports from different campaigns tests the model's Polk-specific
    prompt language honestly). A goal naming "vehicle sales" should read as
    a Polk-related goal thread; the whole point of this scenario is the
    match-rate-respecting language rule: any thread citing a raw matched
    figure (44,756 households / 5 target dealer sales) must also name the
    90.49% match rate somewhere in the SAME draft, never presenting a
    matched count as if it were the campaign's complete result.
    """
    rep.scenario = "MW attribution + Polk automotive match-back"
    print(f"\n{'=' * 78}\nSCENARIO  {rep.scenario}\n{'=' * 78}")
    for path in (ATTRIBUTION_MW, POLK_DASHBOARD):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    polk = pk.parse_polk_export(str(POLK_DASHBOARD))
    goals = ["Drive incremental vehicle sales at the target dealership"]
    facts = ra.build_facts_payload(attribution, None, goals=goals, polk=polk)

    draft, error = app.call_claude_attr_draft(facts)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)
    if save:
        out = FIXTURES / "attr_mw_with_polk.live.json"
        import json
        out.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        print(f"    ....  saved raw response to {out.relative_to(REPO)}")

    _kwargs, highlight_bullets, takeaway_bullets = check_facts_only(rep, draft, facts)
    all_text = _all_draft_text(draft)
    mentions_matched_households = "44,756" in all_text or "44756" in all_text
    mentions_target_dealer_sales = ("5 target dealer" in all_text.lower()
                                    or "five target dealer" in all_text.lower())
    if mentions_matched_households or mentions_target_dealer_sales:
        rep.check("a raw Polk matched figure is cited alongside the 90.49% match rate "
                 "somewhere in the draft -- never presented as a complete total",
                 "90.49" in all_text or "90.5%" in all_text, all_text[:1200])
    else:
        print("    ....  the model didn't cite a raw Polk matched figure this run -- "
             "match-rate-language check has nothing to verify")
    check_highlights_no_tile_restatement(rep, highlight_bullets, facts)


SCENARIOS = {"mw": run_mw, "cardinal": run_cardinal, "cardinal_trend": run_cardinal_trend,
            "no_goals": run_no_goals, "waepa_proposal": run_waepa_proposal_link,
            "synthetic_benchmark": run_synthetic_above_benchmark,
            "mw_optimization": run_mw_optimization,
            "mw_polk": run_mw_with_polk}


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("only", nargs="?", help="run only the scenario starting with this")
    parser.add_argument("--save", action="store_true",
                        help="write each raw response to tests/fixtures/attr_<name>.live.json")
    args = parser.parse_args()

    if not st.secrets.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY is not set in .streamlit/secrets.toml -- this tier needs it. "
              "Tier 1 (tests/test_attribution_draft.py) runs offline.")
        return 2

    names = [n for n in SCENARIOS if not args.only or n.startswith(args.only)]
    if not names:
        print(f"No scenario matches {args.only!r}. Known: {', '.join(SCENARIOS)}")
        return 2

    print(f"Model: {app.ANTHROPIC_MODEL}   ({len(names)} live call{'s' if len(names) != 1 else ''})")
    rep = Report()
    for name in names:
        SCENARIOS[name](rep, args.save)

    print("\n" + "=" * 78)
    print(f"{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped")
    for f in rep.failed:
        print(f"  FAILED  {f}")
    if rep.failed:
        print("\nA failure here means the model's behaviour drifted from what the prompt asks "
              "for.\nRe-run once before concluding anything -- these are probabilistic, and a "
              "single\nrun is a sample, not a measurement.")
    print("=" * 78)
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
