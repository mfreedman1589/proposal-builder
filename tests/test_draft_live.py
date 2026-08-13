"""Tier 2 -- the live model-behaviour suite. Costs money; run it yourself.

    ! python tests/test_draft_live.py
    ! python tests/test_draft_live.py hvac          # one fixture
    ! python tests/test_draft_live.py --save        # keep the raw responses

Four API calls per full run (one per fixture), a few cents.

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
resolves forward, that two audiences sharing a budget become two lines while
one audience described four ways stays one, and that ambiguity lands in one of
the two review lists. Which list is the prompt's routing rule's call (client
question vs seller check), so assertions about whether something was flagged
look in both -- and print both on failure.

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


def fee_lines(draft):
    return [l for l in lines_of(draft) if l.get("product") == app.CUSTOM_FEE_PRODUCT]


def mentions(unresolved, *words):
    """Is any unresolved item about this? Word-level, since the model's
    wording varies run to run and only the subject has to be there."""
    return [u for u in unresolved if any(w in u.lower() for w in words)]


def review_items(draft):
    """Both review lists, together.

    Which list an item lands in is the routing rule's call (client question vs
    seller check), so an assertion about whether something got flagged at all
    has to look in both -- and, just as importantly, *print* both when it
    fails. Searching one list and printing that same list is how three live
    runs reported a missing custom-audience note that was sitting in the other
    one the whole time.
    """
    return (draft.get("unresolved") or []) + (draft.get("unresolved_internal") or [])


def budgets_in(draft):
    """Every stated budget and where it lives -- top level, or one per option.

    A single plan carries it at the top level and an options array carries one
    per option. A check reading only the top level reports "0.0" for a draft
    that came back in the other shape, which says nothing about where the
    money actually was.
    """
    return {"total_budget": draft.get("total_budget"),
            "options[].total_budget": [o.get("total_budget")
                                       for o in (draft.get("options") or [])]}


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
    # The notes describe creative that swaps by county. Nothing but the
    # option's description tells the model that is dynamic_creative, and
    # without it the request reaches the proposal nowhere at all.
    rep.check("dynamic_creative on (creative 'swaps based on which county')",
              "dynamic_creative" in attribution, attribution)

    rep.section("A quoted production charge stays a production charge")
    fees = fee_lines(draft)
    production = [l for l in fees if "production" in str(l.get("label", "")).lower()]
    rep.check("the $850 is a flat-fee line labelled for commercial production",
              bool(production), [l.get("label") for l in fees])
    if production:
        amounts = {float((l.get("allocation") or {}).get("flat_amount") or 0) for l in production}
        rep.equal("at the quoted amount", amounts, {850.0})
    # Billing for production and listing it as free in the same deck is the
    # contradiction this whole rule exists to prevent.
    rep.check("commercial_production NOT also claimed as included at no charge",
              "commercial_production" not in attribution, attribution)
    # The failure this replaced: the $850 was moved onto a deliverable the
    # notes never priced.
    reassigned = [l.get("label") for l in fees
                  if "production" not in str(l.get("label", "")).lower()]
    rep.check("the quoted amount was not reassigned to another deliverable",
              not reassigned, reassigned)


def check_dental(rep, draft):
    rep.section("Single budget -> a single plan, no scenarios invented")
    options = draft.get("options") or []
    rep.check("no options array (one plan is the normal answer)", not options,
              [o.get("name") for o in options])
    rep.check("total_budget is the stated figure",
              float(draft.get("total_budget") or 0) == 30000.0, budgets_in(draft), 30000.0)
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

    rep.section("A product named without a budget still gets a line")
    # "audience targeting on top of the general streaming" names a product
    # and gives no split. Omitting it hides the request from the reviewer;
    # a flagged even split doesn't.
    targeting = [l for l in lines_of(draft)
                 if str(l.get("product", "")).startswith("audience_targeting")]
    rep.check("the requested audience targeting product has a line", bool(targeting),
              [l.get("product") for l in lines_of(draft)])
    if targeting:
        allocs = [l.get("allocation") or {} for l in targeting]
        rep.check("it carries an allocation rather than being left unpriced",
                  all(a for a in allocs), allocs)
        rep.check("the unstated split is flagged for review (either list)",
                  bool(mentions(review_items(draft),
                                "split", "evenly", "weight", "allocation")),
                  review_items(draft))


def resolved_rows(draft, budget):
    """The drafted lines run through the real resolver, so an assertion about
    "the right amounts" is about dollars on a row rather than about the
    percentages the model happened to phrase them in."""
    rows, _, _, _ = app.resolve_drafted_lines(
        lines_of(draft), budget, 1.0, "Mar 2027 - May 2027", "Washington, DC DMA", "")
    return rows


def check_ashford_two_track(rep, draft):
    """Two audiences with a stated split -> one line each, at those amounts.

    The distinction this fixture exists for: separate audiences sharing a
    budget are separate lines, and the split is the signal. Its opposite
    number is check_capital_ridge_stacked.
    """
    rep.section("Two audiences with a stated split -> a line each")
    lines = lines_of(draft)
    premion = [l for l in lines if l.get("product") == "premion_streaming_tv"]
    rep.check("two Premion streaming lines, not one blended one", len(premion) == 2,
              [(l.get("label"), l.get("audience_track")) for l in lines])

    if len(premion) == 2:
        rep.check("each line is labelled so the two are distinguishable on the plan",
                  all((l.get("label") or "").strip() for l in premion),
                  [l.get("label") for l in premion])
        rep.check("each line carries its own audience_track",
                  all((l.get("audience_track") or "").strip() for l in premion)
                  and premion[0].get("audience_track") != premion[1].get("audience_track"),
                  [l.get("audience_track") for l in premion])
        # One side is trade/designers, the other retail/consumers. Word-level,
        # since the model's phrasing varies.
        tracks = " | ".join((l.get("audience_track") or "").lower() for l in premion)
        rep.check("the trade/designer side is described as its own audience",
                  any(w in tracks for w in ("designer", "trade", "stager")), tracks)
        rep.check("the retail/consumer side is described as its own audience",
                  any(w in tracks for w in ("consumer", "retail", "shopper", "homeowner")),
                  tracks)

    rep.section("...at the amounts the notes stated")
    rep.check("total_budget is the stated figure",
              float(draft.get("total_budget") or 0) == 80000.0, budgets_in(draft), 80000.0)
    rows = resolved_rows(draft, 80000.0)
    costs = sorted(round(float(r["Cost"])) for r in rows)
    rep.equal("the 60/40 split resolves to $48,000 / $32,000", costs, [32000, 48000])


def check_capital_ridge_stacked(rep, draft):
    """One audience described by four attributes -> ONE line carrying all of
    them. The opposite number of check_ashford_two_track: same surface shape
    in the notes (several audience terms), completely different plan."""
    rep.section("One stacked audience -> one line, not one per attribute")
    lines = lines_of(draft)
    premion = [l for l in lines if l.get("product") == "premion_streaming_tv"]
    rep.check("a single Premion streaming line (the attributes are one audience)",
              len(premion) == 1,
              [(l.get("label"), l.get("audience_track")) for l in lines])
    rep.check("the four attributes were not split into separate lines",
              len(lines) == 1, [l.get("product") for l in lines])

    rep.section("...and that line carries the whole stack")
    if premion:
        track = (premion[0].get("audience_track") or "").lower()
        print(f"          audience_track: {premion[0].get('audience_track')!r}")
        attributes = {
            "age (35+)": ("35",),
            "homeowners": ("homeowner", "home owner", "own their"),
            "income": ("income", "hhi", "affluent", "$"),
            "researching cosmetic procedures": ("cosmetic", "researching", "in-market"),
        }
        missing = [name for name, words in attributes.items()
                   if not any(w in track for w in words)]
        # audience_track is the deck's Targeting column. A one-segment answer
        # under a four-attribute brief is the under-fill this rule exists to
        # stop -- it printed "DEMO Homeowner" beside campaign specs listing all
        # four.
        rep.check("every attribute from the notes reached the targeting stack",
                  not missing, f"missing: {missing}")
        rep.check("the stack is more than a single segment name",
                  len(track.split()) >= 4, track)


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
    # The two halves of the audience-to-line rule, deliberately adjacent: the
    # notes look alike (several audience terms in both) and the correct plans
    # are opposites. Testing either alone would let the model satisfy it by
    # always splitting, or by never splitting.
    "ashford_two_track": {
        "title": "Furniture retail / two audiences / stated 60-40 split",
        "vertical": "retail",
        "market": "DC",
        "checks": check_ashford_two_track,
    },
    "capital_ridge_stacked": {
        "title": "Dental / one audience described by four attributes",
        "vertical": "healthcare",
        "market": "Harrisburg",
        "checks": check_capital_ridge_stacked,
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
    # Printed too, not just searched. A failure about a note the model was
    # supposed to write is undiagnosable if half the notes it wrote are
    # invisible here -- three runs of the dental fixture failed on a missing
    # custom-audience flag before anyone could see it had gone to the other
    # list all along.
    for item in draft.get("unresolved_internal") or []:
        print(f"          - [internal] {item}")

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

        # apply_draft_to_form now caps this structurally -- extra customs are
        # moved into unresolved whatever the model returns (Tier 1's
        # custom_audience_cap scenario proves it offline). So this assertion
        # is a prompt-quality signal, not a correctness gate: a failure means
        # the model's first pass needed correcting, not that a bad draft can
        # reach the form. Worth keeping, because the segment the app keeps is
        # whichever one the model ranked first.
        catalog = app.load_audience_catalog()
        rfp = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
        custom = [n for n in names if not rfp.get(n, True)]
        rep.check(f"model's first pass respects the one-custom limit unaided "
                  f"({len(custom)} returned; the app would cap it either way)",
                  len(custom) <= 1, custom)
        if custom:
            # Either list counts. The prompt routes "double-checking a segment
            # is available" to unresolved_internal as a seller-side check,
            # which is what a custom segment needs -- so demanding it appear
            # in "unresolved" specifically was asserting against the prompt's
            # own routing rule, and failed three runs running for it.
            rep.check("the custom segment is called out (either list)",
                      bool(mentions(review_items(draft),
                                    "custom", "rfp", "selectable", custom[0].lower())),
                      review_items(draft))
        print(f"          segments: " + ", ".join(
            f"{n}{'' if rfp.get(n, True) else ' [custom]'}" for n in names))
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
