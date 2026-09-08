"""Tier 2 -- the live model-behaviour suite for ATTRIBUTION_REPORT_PLAN.md
Phase 4 (the model-facing half of the Attribution Report Builder). Costs
money; run it yourself.

    ! python tests/test_attribution_draft_live.py
    ! python tests/test_attribution_draft_live.py mw          # one scenario
    ! python tests/test_attribution_draft_live.py --save

Same split as tests/test_draft_live.py (Tier 2) vs. tests/test_draft_
regression.py (Tier 1): this file asks one question -- **does the model
still connect visitor intent to the stated goal, and stay inside the
facts-only contract, on the real datasets?** Assertions are structural
(a class is named, a real number is cited, no fabricated number slips
through `app.apply_attr_draft`'s own check) -- never exact wording, which
varies run to run.

Real fixtures (MW's/Cardinal's own exports) are gitignored, same convention
as test_report_assembly.py / test_wideorbit.py -- SKIPs rather than fails
when they're absent.
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
import report_assembly as ra     # noqa: E402

market_lookup.install()

ATTRIBUTION_MW = REPO / "MW attribution excel.xlsx"
DELIVERY_MW = REPO / "MW delivery.xlsx"
ATTRIBUTION_CARDINAL = REPO / "Premion Website Attribution Cardinal.xlsx"
DELIVERY_CARDINAL = REPO / "Premion OTT.xlsx"


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
    draft_texts = [("highlight bullet", f"{b.get('head', '')} {b.get('detail', '')}")
                  for b in draft.get("highlight_bullets") or []]
    draft_texts += [("takeaway bullet", f"{b.get('head', '')} {b.get('detail', '')}")
                    for b in draft.get("takeaway_bullets") or []]
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
    return kwargs


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
    facts = ra.build_facts_payload(attribution, delivery, goals=goals)
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

    check_facts_only(rep, draft, facts)
    check_goal_intent_connection(rep, draft, "in-store visits", "store_visit")
    if store_visit:
        # The model may cite the raw visit count OR the formatted share --
        # both are real, payload-sourced facts about the same class, and
        # which one reads better in a sentence is a wording choice, not a
        # correctness question.
        haystack = _all_draft_text(draft)
        share_pct = f"{store_visit['share'] * 100:.1f}"
        rep.check(f"cites the real store-visit fact (count {store_visit['visits']} or "
                 f"share {share_pct}%) somewhere in the drafted content",
                 str(store_visit["visits"]) in haystack or share_pct in haystack,
                 haystack[:600])


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
    facts = ra.build_facts_payload(attribution, delivery, goals=goals)
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

    check_facts_only(rep, draft, facts)
    check_goal_intent_connection(rep, draft, "service calls", "lead")
    if lead:
        haystack = _all_draft_text(draft)
        share_pct = f"{lead['share'] * 100:.1f}"
        rep.check(f"cites the real lead-intent fact (count {lead['visits']} or share "
                 f"{share_pct}%) somewhere in the drafted content",
                 str(lead["visits"]) in haystack or share_pct in haystack,
                 haystack[:600])


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


SCENARIOS = {"mw": run_mw, "cardinal": run_cardinal, "no_goals": run_no_goals}


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
