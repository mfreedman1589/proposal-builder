"""Tier 1 -- free, offline regression for ATTRIBUTION_REPORT_PLAN.md Phase 4
(the model-facing half of the Attribution Report Builder), rewritten for the
Highlights/Takeaways rework's `threads` schema.

Same discipline as tests/test_draft_regression.py: a recorded model response
(tests/fixtures/attr_synthetic.draft.json, frozen -- treat it as read-only)
goes through the *real* `app.apply_attr_draft` (which now calls `report_
assembly.distribute_threads` internally), so this proves the parsing/
validation code, not the model's wording. The facts payload it was recorded
against is synthetic (`synthetic_facts_payload` below) rather than derived
from a real client export -- MW/Cardinal's own xlsx files are real client
data and stay gitignored (see test_report_assembly.py); freezing a fixture
from them would put real campaign numbers permanently into git history.
Re-record with:

    python -c "import json,app; from tests.test_attribution_draft import synthetic_facts_payload; \
d,e=app.call_claude_attr_draft(synthetic_facts_payload()); \
open('tests/fixtures/attr_synthetic.draft.json','w').write(json.dumps(d,indent=2))"

    python tests/test_attribution_draft.py
"""
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"

import streamlit as st  # noqa: E402
import app               # noqa: E402
import report_assembly as ra  # noqa: E402


class Report:
    def __init__(self):
        self.passed, self.failed = 0, []

    def check(self, label, ok, detail=""):
        if ok:
            print(f"  PASS  {label}")
            self.passed += 1
        else:
            print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
            self.failed.append(label)


def synthetic_facts_payload():
    """A small, entirely made-up facts payload matching
    `report_assembly.build_facts_payload`'s own shape -- not derived from
    any real export. Two markets (so `breakdown.dimension_forced` is
    "Market", the same "never overridable" case test_report_assembly.py's
    Phase 4 plumbing check exercises against a real single-market export),
    a real store-visit-shaped intent class, and rep notes describing a
    mid-flight optimization -- enough surface for the model to have
    something real to connect a goal to, and something real to enrich from
    the notes, without any of it being a real client's numbers.

    Highlights/Takeaways rework additions: a "vertical" that resolves to a
    real benchmark row (Home Services/Improvement) but sits BELOW it on
    both rates (no citation should fire); an empty "prior_periods" (a first
    report); "plan_vs_actual" absent (the toggle is off); a real
    "conversion_definition" so the model has something to name conversions
    with (not that this fixture's own headline includes conversions --
    the field is independent of whether conversions exist).
    """
    return {
        "goals": ["Drive in-store visits", "Increase brand awareness"],
        "notes": ("Optimized mid-flight by removing Wednesdays and the lowest-performing zip "
                 "codes -- shifted that budget to the strongest-performing market."),
        "headline": {"delivered_impressions": 1200000, "attributed_unique_visitors": 950,
                    "attributed_unique_visitor_rate": 0.00079, "attributed_rate": 0.0163},
        "audience": {"top": {"label": "Home Improvement Intenders", "delivered_impressions": 1200000,
                             "attributed_impressions": 19560, "attributed_rate": 0.0163},
                    "rows": [{"label": "Home Improvement Intenders", "delivered_impressions": 1200000,
                              "attributed_impressions": 19560, "attributed_rate": 0.0163}]},
        "market": {"count": 2,
                  "top": {"label": "Washington, DC", "delivered_impressions": 800000,
                          "attributed_impressions": 13500, "attributed_rate": 0.0169},
                  "rows": [{"label": "Washington, DC", "delivered_impressions": 800000,
                            "attributed_impressions": 13500, "attributed_rate": 0.0169},
                           {"label": "Richmond", "delivered_impressions": 400000,
                            "attributed_impressions": 6060, "attributed_rate": 0.01515}]},
        "creative": {"top": None, "rows": []},
        "breakdown": {"dimension_forced": "Market", "audience_available": True, "creative_available": False},
        "intent": {"classes": [
            {"intent": "consider", "label": "Product consideration", "visits": 480, "share": 0.5053},
            {"intent": "store_visit", "label": "Store visit intent", "visits": 210, "share": 0.2211},
            {"intent": "homepage", "label": "Homepage", "visits": 190, "share": 0.2},
            {"intent": "learn", "label": "Research / learning", "visits": 45, "share": 0.0474},
            {"intent": "other", "label": "Other pages", "visits": 25, "share": 0.0263},
        ], "noise_visits": 12},
        "top_pages": [
            {"label": "Store Locator", "visitors": "210", "share": "22%"},
            {"label": "Kitchen Remodel", "visitors": "300", "share": "32%"},
            {"label": "Homepage", "visitors": "190", "share": "20%"},
        ],
        "zip": {"baseline_rate": 0.0163, "rows": [
            {"zip": "20852", "area": "Washington, DC", "share": "3.1%", "rate": "2.40%",
             "multiple": "1.47x", "outperformer": True},
            {"zip": "23225", "area": "Richmond", "share": "1.4%", "rate": "1.90%",
             "multiple": "1.17x", "outperformer": True},
        ]},
        "delivery": {"delivered_impressions": 1200000, "vcr": 0.97, "frequency": 4.2, "uniques": 41000,
                    "ctv_share": 0.92,
                    "top_publishers": [{"name": "Hulu", "impressions": 300000},
                                       {"name": "Peacock", "impressions": 200000}],
                    "breakdown_applies": True,
                    "by_geo": [{"label": "Washington, DC", "impressions": 800000},
                               {"label": "Richmond", "impressions": 400000}],
                    "by_creative": [{"name": "Fall Sale :30", "impressions": 700000, "vcr": 0.98},
                                    {"name": "Fall Sale :15", "impressions": 500000, "vcr": 0.95}]},
        "vertical": "Home Services/Improvement",
        # Below the Home Services/Improvement row (3.56%/0.41%) on both
        # rates -- deliberately None, so a well-behaved response cites no
        # benchmark at all.
        "benchmark": None,
        "conversion_definition": "quote requests",
        "prior_periods": [],
        "plan_vs_actual": None,
    }


def check_facts_only_contract(rep):
    print("\nFacts-only contract -- every number in the frozen response traces to the payload")
    fixture = FIXTURES / "attr_synthetic.draft.json"
    if not fixture.exists():
        rep.check("frozen fixture present", False, f"{fixture} missing")
        return
    draft = json.loads(fixture.read_text(encoding="utf-8"))
    facts = synthetic_facts_payload()
    kwargs, warnings = app.apply_attr_draft(draft, facts)
    # `warnings` also carries the model's own `goal_alignment_notes` (a
    # separate, legitimate signal -- see check_goal_alignment_note below),
    # so the facts-only assertion checks number-fabrication specifically
    # rather than the combined list being empty.
    highlight_bullets, takeaway_bullets, _whats_next = ra.distribute_threads(draft.get("threads"))
    draft_texts = [("highlight bullet", f"{head} {detail}") for head, detail in highlight_bullets]
    draft_texts += [("takeaway bullet", f"{head} {detail}") for head, detail in takeaway_bullets]
    for key in ("attribution_headline_note", "url_headline_note", "zip_headline_note",
               "attribution_narrative", "url_intent_narrative", "zip_narrative",
               "delivery_narrative", "delivery_breakdown_narrative"):
        if draft.get(key):
            draft_texts.append((key, draft[key]))
    number_violations = app._attr_draft_number_violations(draft_texts, facts)
    rep.check("zero fabricated-number violations", number_violations == [], number_violations)
    return kwargs, draft


def check_no_benchmark_citation(rep, draft):
    print("\nBenchmark is None in this fixture -- nothing should cite it")
    highlight_bullets, takeaway_bullets, _whats_next = ra.distribute_threads(draft.get("threads"))
    haystack = " ".join(f"{h} {d}" for h, d in highlight_bullets + takeaway_bullets).lower()
    rep.check("no 'benchmark' language in the drafted threads (facts.benchmark is None here)",
             "benchmark" not in haystack, haystack[:400])


def check_pct_of_plan_is_traced_as_a_rate(rep):
    """Real bug found live 2026-09-11: `plan_vs_actual`'s own `pct_of_plan`
    values (a real, 0-1 fraction) weren't recognized as rate-shaped by
    `_ATTR_RATE_LIKE_KEYS` (only "rate"/"share"/"vcr" were listed, and
    "pct_of_plan" contains none of them) -- so a model correctly writing
    "84.6% of plan" from a REAL fact got flagged as a fabrication. Fixed by
    adding "pct" to the marker list. This is a facts-only CHECKER bug, not
    a model-drift case -- guard it permanently, offline."""
    print("\npct_of_plan is recognized as a rate-shaped key, not just rate/share/vcr")
    facts = {"plan_vs_actual": {"rows": [{"label": "x", "pct_of_plan": 0.8461539527654216}]}}
    strings, _ints = app._attr_payload_numbers(facts)
    rep.check("84.6 (one decimal) is in the allowed set", "84.6" in strings, strings)
    rep.check("85 (rounded) is in the allowed set", "85" in strings, strings)
    violations = app._attr_draft_number_violations(
        [("note", "Washington, DC delivered at 84.6% of plan.")], facts)
    rep.check("a drafted sentence citing the real pct_of_plan value is NOT flagged as fabricated",
             violations == [], violations)


def check_internal_keys_excluded_from_traced_set(rep):
    """The generalized "_internal"-prefix exclusion (2026-09-11) --
    ZIP_MIN_SHARE (the eligibility floor, not a fact about any zip) rides
    into the payload as `zip._internal_min_share_pct` and must be excluded
    from the traced set the same way "benchmark" already is, so a model
    that narrates the floor back ("both carrying more than 1% share") gets
    caught, not waved through because the number happens to be real."""
    print("\n\"_internal\"-prefixed keys are excluded from the traced-number set")
    facts = {"zip": {"rows": [], "_internal_min_share_pct": 1.0}}
    strings, _ints = app._attr_payload_numbers(facts)
    rep.check("1.0 (the internal floor) is NOT in the allowed set", "1" not in strings, strings)
    violations = app._attr_draft_number_violations(
        [("note", "Both carrying more than 1% share of attributed impressions.")], facts)
    rep.check("narrating the internal floor back is flagged as a violation",
             len(violations) == 1, violations)


def check_thread_entity_violations(rep):
    """The mechanical half of Matt's thread-coherence rule (2026-09-11): a
    thread's HEAD may only name a real market/audience/zip that its own
    FINDING actually mentions. Real case: MW's thread 3 headed
    "Raleigh-Durham Leads Markets" while the finding's own numbers named
    Greensboro-High Point-Winston-Salem instead."""
    print("\nThread entity coherence -- a head's named entity must appear in its own finding")
    facts = {
        "market": {"rows": [{"label": "Raleigh-Durham"}, {"label": "Greensboro-High Point-Winston-Salem"}]},
    }
    bad_threads = [{"head": "Raleigh-Durham Leads Markets", "anchor": "goal",
                   "finding": "Greensboro-High Point-Winston-Salem posted a 1.38% attributed rate.",
                   "meaning": "Both markets performed comparably."}]
    violations = app._thread_entity_violations(bad_threads, facts)
    rep.check("catches a head naming a market its own finding never mentions",
             violations == [("Raleigh-Durham Leads Markets", "Raleigh-Durham")], violations)

    good_threads = [{"head": "Raleigh-Durham Leads Markets", "anchor": "goal",
                    "finding": "Raleigh-Durham posted a 1.36% attributed rate, the highest of the two markets.",
                    "meaning": "Raleigh-Durham is the stronger market."}]
    rep.check("no violation when the head's entity IS in its own finding",
             app._thread_entity_violations(good_threads, facts) == [])

    hyphen_threads = [{"head": "Raleigh Leads Markets", "anchor": "goal",
                      "finding": "Raleigh-Durham posted the higher rate.", "meaning": "m"}]
    rep.check("a head naming just the first city of a hyphenated market matches",
             app._thread_entity_violations(hyphen_threads, facts) == [])


def check_goal_alignment_note(rep, draft):
    print("\ngoal_alignment_notes -- the model's own disagreement/no-data flag, not discarded")
    notes = draft.get("goal_alignment_notes") or []
    rep.check("the frozen response correctly declined to quantify the Wednesday-removal "
             "note (no day-of-week data exists in the payload) rather than fabricating one, "
             "and said so in goal_alignment_notes",
             any("wednesday" in str(n).lower() for n in notes), notes)


def check_kwargs_shape(rep, kwargs):
    print("\nKwargs shape -- ready to **-expand into report_assembly.build_report_deck")
    rep.check("highlight_bullets is a non-empty list of (head, detail) pairs, capped at 4",
             bool(kwargs["highlight_bullets"]) and len(kwargs["highlight_bullets"]) <= 4
             and all(isinstance(t, tuple) and len(t) == 2 for t in kwargs["highlight_bullets"]),
             kwargs["highlight_bullets"])
    rep.check("takeaway_bullets is a non-empty list of (head, detail) pairs, capped at 4",
             bool(kwargs["takeaway_bullets"]) and len(kwargs["takeaway_bullets"]) <= 4
             and all(isinstance(t, tuple) and len(t) == 2 for t in kwargs["takeaway_bullets"]),
             kwargs["takeaway_bullets"])
    rep.check("headline_notes carries exactly the three expected keys",
             set(kwargs["headline_notes"]) == {"attribution", "url", "zip"}, kwargs["headline_notes"])
    rep.check("narratives carries exactly the eight expected keys (v0_6 adds "
             "response_profile/ott_retargeting)",
             set(kwargs["narratives"]) == {"attribution", "response_profile", "url_intent", "zip",
                                           "delivery", "delivery_breakdown", "live_sports",
                                           "ott_retargeting"}, kwargs["narratives"])
    rep.check("live_sports narrative is None -- this fixture's facts carry no sports block",
             kwargs["narratives"]["live_sports"] is None, kwargs["narratives"]["live_sports"])
    rep.check("ott_retargeting narrative is None -- this fixture's facts carry no OTT retargeting export",
             kwargs["narratives"]["ott_retargeting"] is None, kwargs["narratives"]["ott_retargeting"])
    rep.check("neither highlight_bullets/takeaway_bullets nor any other threads-derived key "
             "leaks a raw 'threads' entry into kwargs -- only the distributed shape",
             "threads" not in kwargs, kwargs)
    # This fixture's own facts force "Market" (2 markets) -- there is no
    # judgment call for the model to make, so a well-behaved response
    # returns null here (report_assembly.pick_breakdown_dimension ignores
    # the override outright either way, but a null response is still the
    # correct thing for the model to have written).
    rep.check("breakdown_dimension_override is None (this fixture's market count forces "
             "Market, so there's nothing to override)",
             kwargs["breakdown_dimension_override"] is None, kwargs["breakdown_dimension_override"])


def check_goal_connection(rep, draft):
    print("\nThe URL slide is the point of this phase -- does it actually connect to the goal?")
    narrative = (draft.get("url_intent_narrative") or "").lower()
    rep.check("the store-visit intent class (the goal-relevant one in this fixture) is named",
             "store" in narrative, draft.get("url_intent_narrative"))
    rep.check("the real store-visit share (22.11%, from the payload) or visit count (210) is cited",
             "22.11" in narrative or "210" in narrative, draft.get("url_intent_narrative"))


def check_notes_reached_output(rep, draft):
    print("\nNotes reach the output -- the mid-flight optimization shows up SOMEWHERE, "
         "without a fabricated day-of-week number")
    _hi, takeaway_bullets, _whats_next = ra.distribute_threads(draft.get("threads"))
    haystack = " ".join([
        draft.get("delivery_narrative") or "", draft.get("delivery_breakdown_narrative") or "",
        " ".join(f"{h} {d}" for h, d in takeaway_bullets),
        " ".join(str(n) for n in draft.get("goal_alignment_notes") or []),
    ]).lower()
    rep.check("something in the drafted content reflects the mid-flight optimization",
             any(w in haystack for w in ("optimiz", "mid-flight", "reallocat", "wednesday")),
             haystack[:400])


def check_conversion_definition_used(rep, draft):
    print("\nConversion definition -- 'quote requests' should appear if a thread names a "
         "conversion at all, never the generic word alongside it")
    highlight_bullets, takeaway_bullets, whats_next = ra.distribute_threads(draft.get("threads"))
    haystack = " ".join([h for h, _d in highlight_bullets] + [d for _h, d in highlight_bullets]
                        + [h for h, _d in takeaway_bullets] + [d for _h, d in takeaway_bullets]
                        + whats_next).lower()
    if "conversion" in haystack:
        rep.check("'quote requests' (the stated conversion_definition) is used instead of "
                 "the bare word when a conversion is named",
                 "quote request" in haystack, haystack[:400])
    else:
        rep.check("no conversion mentioned at all in this fixture (facts carry none) -- "
                 "acceptable, the definition just wasn't needed", True)


if __name__ == "__main__":
    rep = Report()
    result = check_facts_only_contract(rep)
    if result:
        kwargs, draft = result
        check_kwargs_shape(rep, kwargs)
        check_goal_connection(rep, draft)
        check_notes_reached_output(rep, draft)
        check_goal_alignment_note(rep, draft)
        check_no_benchmark_citation(rep, draft)
        check_conversion_definition_used(rep, draft)
    check_pct_of_plan_is_traced_as_a_rate(rep)
    check_internal_keys_excluded_from_traced_set(rep)
    check_thread_entity_violations(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed out of {total}")
    sys.exit(1 if rep.failed else 0)
