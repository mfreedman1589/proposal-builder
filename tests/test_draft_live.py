"""Tier 2 -- the live model-behaviour suite. Costs money; run it yourself.

    ! python tests/test_draft_live.py
    ! python tests/test_draft_live.py hvac          # one fixture
    ! python tests/test_draft_live.py --save        # keep the raw responses

Two API calls per full run (one per fixture), a few cents.

This tier asks one question only: **does the model still do what the prompt
tells it to?** Everything downstream of the response -- the allocation
waterfall, the form state, the assembled deck -- is Tier 1's job
(tests/test_draft_regression.py), which is free and offline and should be
what you run while working. Run this one before shipping a change to the
prompt, the schema, the catalog, or the rate card, and after any model
version bump.

Assertions are deliberately structural. The model's wording varies run to
run and that is fine; what must not vary is that a budget range becomes two
options, that a stated rate becomes a `cpm`, that "hold at rate card" does
NOT, that a trade name resolves to the right vertical, that a year-less date
resolves forward, and that ambiguity lands in `unresolved`.

The key is read exactly the way the app reads it -- `st.secrets`, from
.streamlit/secrets.toml. It is never printed.
"""

import argparse
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"
os.chdir(REPO)

import streamlit as st                          # noqa: E402
import app                                      # noqa: E402


class Report:
    def __init__(self):
        self.passed = 0
        self.failed = []
        self.scenario = ""

    def section(self, title):
        print(f"\n  {title}")
        print(f"  {'-' * len(title)}")

    def check(self, label, ok, actual=None, expected=None):
        if ok:
            self.passed += 1
            print(f"    PASS  {label}")
        else:
            self.failed.append(f"[{self.scenario}] {label}")
            print(f"    FAIL  {label}")
            if expected is not None:
                print(f"          expected: {expected!r}")
            if actual is not None:
                print(f"          actual:   {actual!r}")
        return ok

    def equal(self, label, actual, expected):
        return self.check(label, actual == expected, actual, expected)


def lines_of(draft):
    """Every media plan line in the draft, whichever shape it came back in."""
    out = []
    for opt in draft.get("options") or []:
        out.extend(opt.get("media_plan_lines") or [])
    out.extend(draft.get("media_plan_lines") or [])
    return out


def check_hvac(rep, draft):
    rep.section("Budget range -> one option per stated figure")
    options = draft.get("options") or []
    rep.equal("two options returned", len(options), 2)
    budgets = sorted(float(o.get("total_budget") or 0) for o in options)
    rep.equal("each option carries its own stated budget", budgets, [50000.0, 75000.0])
    rep.check("no option was left without a budget",
              all(o.get("total_budget") for o in options),
              [(o.get("name"), o.get("total_budget")) for o in options])
    rep.check("options are named from the notes, not lettered",
              all((o.get("name") or "").strip() and o.get("name") not in ("A", "B", "Option A", "Option B")
                  for o in options),
              [o.get("name") for o in options])

    rep.section("Negotiated rate -> cpm; 'hold at rate card' -> no cpm")
    lines = lines_of(draft)
    premion = [l for l in lines if l.get("product") == "premion_streaming_tv"]
    sports = [l for l in lines if str(l.get("product", "")).startswith(app.SPORT_PRODUCT_PREFIX)]

    if rep.check("a Premion streaming line exists", bool(premion),
                 [l.get("product") for l in lines]):
        cpms = {float(l["cpm"]) for l in premion if l.get("cpm") not in (None, "")}
        rep.equal("the $28 negotiated rate came back as a cpm override", cpms, {28.0})

    if rep.check("a sports line exists", bool(sports), [l.get("product") for l in lines]):
        held = [l.get("cpm") for l in sports if l.get("cpm") not in (None, "")]
        rep.check("'hold the NFL at rate card' left the sports cpm absent", not held, held)
        rep.check("the sports line uses the sport: namespace, not a product key",
                  all(str(l["product"]).startswith(app.SPORT_PRODUCT_PREFIX) for l in sports),
                  [l.get("product") for l in sports])

    rep.section("Attribution read from the notes")
    attribution = draft.get("attribution") or []
    rep.check("sales attribution on", "sales" in attribution, attribution)
    rep.check("brand lift off ('I don't need a survey')",
              "brand_lift" not in attribution, attribution)


def check_dental(rep, draft):
    rep.section("Single budget -> a single plan, no scenarios invented")
    options = draft.get("options") or []
    rep.check("no options array (one plan is the normal answer)", not options,
              [o.get("name") for o in options])
    rep.equal("total_budget is the stated figure", float(draft.get("total_budget") or 0), 30000.0)
    rep.check("media_plan_lines carries the plan", bool(draft.get("media_plan_lines")),
              draft.get("media_plan_lines"))

    rep.section("Explicit exclusions honoured")
    lines = lines_of(draft)
    sports_lines = [l.get("product") for l in lines
                    if str(l.get("product", "")).startswith(app.SPORT_PRODUCT_PREFIX)]
    rep.check("no sports lines ('doesn't want to be in live sports')",
              not sports_lines, sports_lines)
    rep.check("no sports packages selected", not (draft.get("sports") or []), draft.get("sports"))
    rep.check("agency_involved false (direct client)", draft.get("agency_involved") is False,
              draft.get("agency_involved"))


SCENARIOS = {
    "hvac_two_option": {
        "title": "HVAC / budget range / negotiated rate / sports at rate card",
        "vertical": "home_improvement",
        "market": "DC",
        "checks": check_hvac,
    },
    "dental_single_option": {
        "title": "Dental / single budget / no sports / direct",
        "vertical": "healthcare",
        "market": "Harrisburg",
        "checks": check_dental,
    },
}


def check_common(rep, draft, spec):
    rep.section("Vertical, market, dates, unresolved")
    rep.equal("vertical detected from the trade name", draft.get("vertical"), spec["vertical"])
    rep.equal("market", draft.get("market"), spec["market"])

    start = app._parse_draft_date(draft.get("flight_start"))
    end = app._parse_draft_date(draft.get("flight_end"))
    rep.check("flight_start parses", start is not None, draft.get("flight_start"))
    rep.check("flight_end parses", end is not None, draft.get("flight_end"))
    if start:
        rep.check("year-less date resolved forward, not into the past",
                  start >= date.today(), str(start), f">= {date.today()}")
    if start and end:
        rep.check("flight_end follows flight_start", end >= start, f"{start} -> {end}")

    unresolved = draft.get("unresolved") or []
    rep.check("unresolved is non-empty (these notes contain real ambiguity)",
              bool(unresolved), unresolved)
    for item in unresolved:
        print(f"          - {item}")

    rep.section("Schema hygiene")
    valid_products = (set(app.PRODUCT_TO_WIDGET_KEYS) - app.NON_LINE_PRODUCT_KEYS) | {app.CUSTOM_FEE_PRODUCT}
    valid_sports = set(app.SPORTS.values())
    bad = []
    for line in lines_of(draft):
        product = str(line.get("product") or "")
        if product.startswith(app.SPORT_PRODUCT_PREFIX):
            if product[len(app.SPORT_PRODUCT_PREFIX):] not in valid_sports:
                bad.append(product)
        elif product not in valid_products:
            bad.append(product)
    rep.check("every line names a real product or sport package", not bad, bad)

    zero = [l for l in lines_of(draft)
            if (l.get("allocation") or {}).get("flat_amount") == 0]
    rep.check("no zero-dollar lines", not zero, zero)

    names = [a.get("segment") for a in (draft.get("audiences") or [])]
    if names:
        matched, unmatched = app.validate_segments(names)
        rep.check("every audience is an exact catalog name", not unmatched, unmatched)
    else:
        rep.check("audiences omitted rather than invented", True)


def run(fixture, rep, save):
    spec = SCENARIOS[fixture]
    rep.scenario = spec["title"]
    notes_path = FIXTURES / f"{fixture}.txt"
    print("\n" + "=" * 78)
    print(f"SCENARIO  {spec['title']}")
    print(f"fixture   {notes_path.name}  (live API call)")
    print("=" * 78)

    notes = notes_path.read_text(encoding="utf-8")
    draft, error = app.call_claude_draft(notes)
    if error:
        rep.check("the model returned parseable JSON", False, error)
        return
    rep.check("the model returned parseable JSON", True)

    if save:
        out = FIXTURES / f"{fixture}.live.json"
        out.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        print(f"    ....  saved raw response to {out.relative_to(REPO)}")

    check_common(rep, draft, spec)
    spec["checks"](rep, draft)


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("only", nargs="?", help="run only fixtures starting with this")
    parser.add_argument("--save", action="store_true",
                        help="write each raw response to tests/fixtures/<name>.live.json")
    args = parser.parse_args()

    if not st.secrets.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY is not set in .streamlit/secrets.toml -- "
              "this tier needs it. Tier 1 (tests/test_draft_regression.py) runs offline.")
        return 2

    fixtures = [f for f in SCENARIOS if not args.only or f.startswith(args.only)]
    if not fixtures:
        print(f"No fixture matches {args.only!r}. Known: {', '.join(SCENARIOS)}")
        return 2

    print(f"Model: {app.ANTHROPIC_MODEL}   ({len(fixtures)} live call"
          f"{'s' if len(fixtures) != 1 else ''})")

    rep = Report()
    for fixture in fixtures:
        run(fixture, rep, args.save)

    print("\n" + "=" * 78)
    print(f"{rep.passed} passed, {len(rep.failed)} failed")
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
