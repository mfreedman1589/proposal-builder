"""Tier 2 -- the live model-behaviour suite. Costs money; run it yourself.

    ! python tests/test_draft_live.py
    ! python tests/test_draft_live.py hvac          # one fixture
    ! python tests/test_draft_live.py --save        # keep the raw responses

Eight API calls per full run (one per fixture), a few cents.

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
    # FLOW_REWORK_PLAN.md Phase 4b: agency_involved is gone from the schema
    # -- the gross-up is a rep-only checkbox now, and the model is never
    # asked about it at all. Confirming its absence (not merely a falsy
    # value) is what proves the schema change actually reached the prompt.
    rep.check("no agency_involved field at all -- the model was never asked",
              draft.get("agency_involved") is None, draft.get("agency_involved"))

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
    rows, _, _, _, _ = app.resolve_drafted_lines(
        lines_of(draft), budget, "Mar 2027 - May 2027", "Washington, DC DMA", "")
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


def check_ridgeline_reach(rep, draft):
    """Reach-driven plan with avails stated in the notes.

    Structural only, like every tier 2 check: that the model reached for the
    reach and avails SHAPES when the notes are written in those terms. The
    arithmetic those shapes drive is Python's and is pinned offline in
    tests/test_avails_reach.py -- asserting numbers here would be paying for
    a live call to test code that never touches the model.
    """
    rep.section("Avails stated in the notes are carried, not discarded")
    audiences = draft.get("audiences") or []
    with_avails = [a for a in audiences if a.get("max_avails")]
    rep.check("at least two audiences carry the stated avails",
              len(with_avails) >= 2, [a.get("segment") for a in audiences])
    stated = {1000000, 600000, 400000}
    got = {int(a.get("max_avails") or 0) for a in with_avails}
    rep.check("the figures are the ones the notes gave, unaltered",
              got & stated == got and len(got & stated) >= 2, sorted(got), sorted(stated))
    bases = {str(a.get("avails_basis") or "").lower() for a in with_avails}
    rep.check("reported as monthly, which is what the notes said",
              bases <= {"monthly", ""}, sorted(bases))
    for audience in with_avails:
        print(f"          {audience.get('segment')}: {audience.get('max_avails'):,} "
              f"({audience.get('avails_basis') or 'unstated'})")

    rep.section("A reach brief produces reach allocations")
    options = draft.get("options") or []
    rep.check("two options, one per reach level", len(options) == 2,
              [o.get("name") for o in options])
    reach_lines = [line for option in options
                   for line in (option.get("media_plan_lines") or [])
                   if "percent_of_avails" in (line.get("allocation") or {})]
    rep.check("reach lines were used at all", bool(reach_lines), len(reach_lines))
    percents = sorted({float((l.get("allocation") or {})["percent_of_avails"])
                       for l in reach_lines})
    rep.check("at the two levels the notes named (20 and 40)",
              percents == [20.0, 40.0], percents)
    rep.check("every reach line names the audience it refers to",
              all((l.get("allocation") or {}).get("avails_ref") for l in reach_lines),
              [(l.get("allocation") or {}).get("avails_ref") for l in reach_lines])
    # The reference has to be resolvable, which means matching a segment the
    # model itself returned -- a reference to a name nowhere in the draft
    # prices at nothing.
    segments = {str(a.get("segment", "")).strip().lower() for a in audiences}
    refs = {str((l.get("allocation") or {}).get("avails_ref", "")).strip().lower()
            for l in reach_lines}
    rep.check("and every reference matches one of its own audiences",
              refs <= segments, sorted(refs), sorted(segments))

    rep.section("The rest of the brief still lands")
    rep.check("NFL regular season, not playoffs",
              "nfl_reg" in (draft.get("sports") or []), draft.get("sports"))
    products = [line.get("product") for option in options
                for line in (option.get("media_plan_lines") or [])]
    rep.check("retargeting is on the plan",
              any("retargeting" in str(p) for p in products), products)
    rep.check("the negotiated $30 CPM is on the streaming lines",
              any(l.get("cpm") == 30 for option in options
                  for l in (option.get("media_plan_lines") or [])),
              [l.get("cpm") for option in options
               for l in (option.get("media_plan_lines") or [])])



def check_summit_multi_market(rep, draft):
    """A brief naming three target markets must select all three.

    The failure this exists for is quiet: a multi-market brief that comes back
    with one target market produces a deck carrying one profile slide instead
    of three, and nothing about it looks wrong -- the deck builds, the numbers
    tie, and the two missing markets are simply absent.

    Structural, like every tier 2 check. That the names RESOLVE to real DMAs
    is Python's job and is pinned offline in tests/test_market_profiles.py;
    what the model owns is naming all three, and not confusing them with the
    originating station.
    """
    rep.section("Every named target market is picked up")
    targets = draft.get("target_markets") or []
    if isinstance(targets, str):
        targets = [targets]
    rep.check("three target markets, not one", len(targets) == 3, targets)

    flat = " | ".join(str(t).lower() for t in targets)
    for wanted in ("denver", "atlanta", "phoenix"):
        rep.check(f"{wanted} is among them", wanted in flat, targets)

    rep.section("The target markets are not the originating station")
    # FLOW_REWORK_PLAN.md Phase 5: "market" (the originating station) is a
    # band input now, never in the model's own output at all -- check_common
    # already asserts it's absent from draft. What's still worth checking
    # here is narrower and still real: DC (this scenario's band market)
    # never leaks into the target_markets list the model DOES return.
    rep.check("the originating market didn't leak into the target list",
              not any("harrisburg" in str(t).lower() for t in targets), targets)
    # The notes explicitly park a fourth market for next year. Picking it up
    # would put a market in the deck the client said to leave out.
    rep.check("the unconfirmed fourth market was not invented",
              len(targets) == 3, targets)

    rep.section("The rest of the brief still lands")
    rep.check("no live sports, as stated", not (draft.get("sports") or []),
              draft.get("sports"))
    # FLOW_REWORK_PLAN.md Phase 4b: these notes name a real agency
    # ("Agency is involved -- Meridian handles their media"), which used to
    # flip agency_involved true. The field is gone from the schema now --
    # the gross-up is a rep-only checkbox -- so the correct behaviour is
    # that the model never returns the field at all, agency mention or not.
    rep.check("no agency_involved field at all, even though the notes name an agency",
              draft.get("agency_involved") is None, draft.get("agency_involved"))


def check_plaza_motors_net_budget(rep, draft):
    """A stated budget with stated avails alongside it -- the budget must
    drive impressions and cost, with the avails figure used only to report
    what percentage of it the budget reaches.

    This is the live failure this fixture exists to catch: a plan that came
    back priced against the FULL avails figure (a percent_of_avails
    allocation) instead of the stated $4,000. (The fixture also used to
    catch agency_involved reading back true from these notes merely naming
    an agency, producing a deck marked up ×1.15 net of nothing the client
    agreed to pay -- that field is retired now, FLOW_REWORK_PLAN.md Phase
    4b, so this file just confirms the model never returns it at all,
    regardless of the notes' own net/no-commission language.)
    """
    rep.section("A stated budget with stated avails -- budget drives, avails is a ceiling")
    lines = lines_of(draft)
    premion = [l for l in lines if l.get("product") == "premion_streaming_tv"]
    rep.check("a Premion streaming line exists", bool(premion), [l.get("product") for l in lines])
    if premion:
        allocs = [(l.get("allocation") or {}) for l in premion]
        reach_allocs = [a for a in allocs if "percent_of_avails" in a]
        rep.check("NOT priced as a percent_of_avails reach line -- a budget was stated",
                  not reach_allocs, allocs)
        budget_allocs = [a for a in allocs
                         if "flat_amount" in a or "percent_of_total" in a]
        rep.check("priced from the stated budget (flat_amount or percent_of_total)",
                  bool(budget_allocs), allocs)
        cpms = {float(l["cpm"]) for l in premion if l.get("cpm") not in (None, "")}
        rep.equal("the stated $29 CPM came back as a cpm override", cpms, {29.0})

    rep.check("total_budget is the stated $4,000, not the avails-implied figure",
              float(draft.get("total_budget") or 0) == 4000.0, budgets_in(draft), 4000.0)

    rows = resolved_rows(draft, 4000.0)
    if rep.check("the plan resolves to at least one row", bool(rows), rows):
        total_cost = sum(float(r["Cost"]) for r in rows)
        rep.check("the resolved plan totals the stated $4,000, not ~$49,300 (1.7M avails at $29)",
                  abs(total_cost - 4000.0) <= 5.0, total_cost, 4000.0)
        total_impressions = sum(float(r["Impressions"]) for r in rows)
        rep.check("impressions come from the $4,000 budget (~138K at $29 CPM), not the 1.7M avails ceiling",
                  total_impressions < 500000, total_impressions)

    rep.section("The stated avails still surface, as a ceiling to report against")
    audiences = draft.get("audiences") or []
    with_avails = [a for a in audiences if a.get("max_avails")]
    rep.check("the stated 1,700,000 avails figure is carried on an audience",
              any(int(a.get("max_avails") or 0) == 1700000 for a in with_avails),
              [a.get("max_avails") for a in audiences])
    rep.check("a reach percentage against those avails is reported for review (either list)",
              bool(mentions(review_items(draft), "%", "percent", "reach")),
              review_items(draft))

    rep.section("No agency field at all, regardless of what the notes say about commission")
    # FLOW_REWORK_PLAN.md Phase 4b: these notes explicitly say Redwood
    # Creative is NOT taking a commission on this buy -- exactly the kind
    # of net/gross phrasing agency_involved used to have to classify. The
    # field is gone from the schema now; the model is never asked and must
    # never volunteer it, regardless of how explicit the notes are about it.
    rep.check("no agency_involved field at all", draft.get("agency_involved") is None,
              draft.get("agency_involved"))


def check_gross_up_verbatim(rep, draft):
    """FLOW_REWORK_PLAN.md Phase 4b's own acceptance case for the prompt
    rewrite. "Gross it up" is a PER-LINE instruction in these notes; the
    checkbox that actually grosses anything is deck-wide and rep-only, so
    the model must never do that arithmetic itself -- it transcribes the
    stated $32 net CPM exactly as given, not the $36.80 grossed figure.
    This is the live check for the double-gross bug that motivated the
    whole rebuild: a CPM rendering at $41.40 (= $36 x 1.15) because the
    model had already grossed a rate and the old per-row toggle grossed it
    again.
    """
    rep.section("A per-line gross-up instruction is transcribed, never computed")
    lines = lines_of(draft)
    premion = [l for l in lines if l.get("product") == "premion_streaming_tv"]
    rep.check("a Premion streaming line exists", bool(premion), [l.get("product") for l in lines])
    if premion:
        cpms = {float(l["cpm"]) for l in premion if l.get("cpm") not in (None, "")}
        rep.equal("the stated net $32 CPM is transcribed verbatim, never grossed to $36.80",
                  cpms, {32.0})
    rep.check("no agency_involved field at all -- the model was never asked",
              draft.get("agency_involved") is None, draft.get("agency_involved"))

    # total_budget_basis's own "genuinely ambiguous" case: "gross it up" is
    # stated about the CPM, but nothing here says the $20,000 IS the gross/
    # client-billed figure (it could just as easily be the working budget the
    # rep is planning against, separate from the rate that needs grossing).
    # This must stay "net" -- the same conclusion that shut down the old
    # agency_involved classifier for the identical reason.
    basis = draft.get("total_budget_basis")
    rep.check('total_budget_basis stays "net" (or absent) -- the notes never say the '
              "$20,000 itself is the gross/client-billed figure, only that the CPM "
              "needs grossing",
              basis in (None, "", "net"), basis)


def check_gross_budget_stated(rep, draft):
    """The Capital Media incident, as a live scenario: the notes state a
    GROSS total budget explicitly ("total gross budget"), alongside a net
    CPM that separately needs grossing -- the case check_gross_up_verbatim's
    own notes are deliberately NOT explicit enough to trigger. This is the
    phrasing variant total_budget_basis exists for.

    Model-layer only (Tier 1's check_gross_budget_basis already proves the
    Python-side conversion and the checkbox auto-tick against a synthetic
    draft): does the model label the budget "gross", and critically, does it
    still do NO arithmetic of its own -- the stated $15,000 has to survive
    verbatim, not arrive pre-divided by 1.15. A model that "helpfully"
    pre-converts the budget itself would then get divided a SECOND time by
    apply_draft_to_form, which is the double-gross bug's shape all over
    again, just moved from CPM to budget.
    """
    rep.section("A gross-labelled budget: transcribed verbatim, labelled, not computed")
    rep.check('total_budget_basis is "gross"',
              str(draft.get("total_budget_basis") or "").strip().lower() == "gross",
              draft.get("total_budget_basis"))
    rep.check("total_budget is the stated $15,000 verbatim -- NOT pre-divided by 1.15",
              float(draft.get("total_budget") or 0) == 15000.0, budgets_in(draft))

    lines = lines_of(draft)
    premion = [l for l in lines if l.get("product") == "premion_streaming_tv"]
    rep.check("a Premion streaming line exists", bool(premion), [l.get("product") for l in lines])
    if premion:
        cpms = {float(l["cpm"]) for l in premion if l.get("cpm") not in (None, "")}
        rep.equal("the stated net $32 CPM is transcribed verbatim, never grossed to $36.80",
                  cpms, {32.0})
    rep.check("no agency_involved field at all -- the model was never asked",
              draft.get("agency_involved") is None, draft.get("agency_involved"))

    rep.section("Python's own conversion (Tier 1's job; re-confirmed here against a live shape)")
    net_budget = round(15000 / app.AGENCY_MARKUP)
    rows = resolved_rows(draft, net_budget)
    if rep.check("the plan resolves to at least one row", bool(rows), rows):
        total_cost = sum(float(r["Cost"]) for r in rows)
        rep.check("priced against the converted net budget (~$13,043), not the stated $15,000",
                  abs(total_cost - net_budget) <= 5.0, total_cost, net_budget)


def check_summit_outside_linear(rep, draft):
    """A client's existing linear buy on someone else's station, in a market
    that isn't one of ours, must not read as Total TV -- that flag is for a
    broadcast schedule PREMION is running on WUSA9 (DC) or WPMT/FOX43
    (Harrisburg), not any mention of broadcast/linear TV in the notes."""
    rep.section("Outside-market linear on another vendor's station is not Total TV")
    rep.check("total_tv is false", draft.get("total_tv") is False, draft.get("total_tv"))
    rep.check("no broadcast import was invented in the media plan lines",
              not [l for l in lines_of(draft) if "broadcast" in str(l.get("product", "")).lower()],
              [l.get("product") for l in lines_of(draft)])
    # FLOW_REWORK_PLAN.md Phase 5: "market" is a band input now, never in the
    # model's own output -- check_common already asserts it's absent, which
    # makes "did the vendor's out-of-market station get read as the
    # originating market" structurally impossible rather than something to
    # keep checking for here.


def check_annapolis_group_selection(rep, draft):
    """Real groups already exist on this proposal (an avails import ran
    first, standing in for a real one) -- the model should SELECT from
    them via group_selection/group_allocation instead of inventing its own
    Premion Streaming TV media_plan_lines entries."""
    rep.section("Selects from the existing avails table instead of inventing lines")
    selection = draft.get("group_selection") or {}
    rep.check("group_selection is present", bool(selection), selection)
    rep.check("mode is named (the notes named specific audiences)",
              selection.get("mode") == "named", selection.get("mode"))
    matched_text = " ".join(str(m) for m in (selection.get("match") or [])).lower()
    rep.check("the match terms mention Subaru and Hyundai",
              "subaru" in matched_text and "hyundai" in matched_text, selection.get("match"))
    rep.check("Volvo/Genesis are NOT named in the selection",
              "volvo" not in matched_text and "genesis" not in matched_text, selection.get("match"))
    allocation = draft.get("group_allocation") or {}
    rep.check("group_allocation is exactly one recognized allocation type",
              len(allocation) == 1 and next(iter(allocation), None) in
              ("flat_amount", "percent_of_total", "percent_of_remainder", "split_evenly", "percent_of_avails"),
              allocation)
    rep.check("split evenly, matching the notes' own words", allocation.get("split_evenly") is True,
              allocation)
    # FLOW_REWORK_PLAN.md Phase 3 precedence rule: the notes also mention
    # each make's avails figure (for context only), alongside the stated
    # $24,000. A stated dollar figure must win -- this is the live failure
    # mode the rule exists to prevent: a plan sized against the mentioned
    # avails instead of what the client actually signed off on.
    rep.check("group_allocation is NOT percent_of_avails, despite avails being mentioned "
              "in the same notes -- the stated $24,000 wins",
              "percent_of_avails" not in allocation, allocation)
    reason = str(draft.get("group_selection_reason") or "")
    rep.check("group_selection_reason is a real sentence, not empty", len(reason) > 10, reason)
    rep.check("the reason is in a seller's own words, not schema keys",
              not any(k in reason for k in
                     ("group_selection", "include_in_plan", "rfp_selectable", "_group_id")),
              reason)
    rep.check("no premion_streaming_tv line was invented directly in media_plan_lines",
              not any(l.get("product") == "premion_streaming_tv" for l in lines_of(draft)),
              [l.get("product") for l in lines_of(draft)])


def _annapolis_existing_groups():
    """Real Annapolis Cars groups (RFPID-253813, via group_scenario_fixtures)
    -- what an avails import would already have put on this proposal
    before the rep pastes these notes and drafts."""
    import group_scenario_fixtures as gsf
    scn = gsf.build_annapolis()
    groups = []
    for g in scn["groups"]:
        g = dict(g)
        g.pop("_flight_label", None)
        g["include_in_plan"] = False
        groups.append(g)
    return groups


SCENARIOS = {
    # FLOW_REWORK_PLAN.md Phase 5: market/flight_start/flight_end are band
    # INPUTS now -- run() presets st.session_state from these three fields
    # before calling call_claude_draft, standing in for a rep who already
    # filled in the band having read the same notes. Each fixture's dates
    # are the same resolution a live run correctly reached under the OLD
    # (prose-inferring) system -- cross-checked against the frozen Tier 1
    # fixtures where one exists (hvac_two_option, dental_single_option) --
    # so this is "what the band already had," not a new invention.
    "hvac_two_option": {
        "title": "HVAC / budget range / negotiated rate / sports at rate card",
        "vertical": "home_improvement",
        "market": "DC",
        "flight_start": date(2026, 9, 8),
        "flight_end": date(2026, 11, 30),
        "checks": check_hvac,
    },
    "summit_multi_market": {
        "title": "Outdoor retail / three target markets / one campaign",
        "vertical": "retail",
        "market": "DC",
        "flight_start": date(2026, 10, 1),
        "flight_end": date(2026, 12, 31),
        "checks": check_summit_multi_market,
    },
    "ridgeline_reach": {
        "title": "Dermatology / reach-driven plan / avails stated in the notes",
        "vertical": "healthcare",
        "market": "DC",
        "flight_start": date(2026, 9, 1),
        "flight_end": date(2026, 11, 30),
        "checks": check_ridgeline_reach,
    },
    "dental_single_option": {
        "title": "Dental / single budget / no sports / direct",
        "vertical": "healthcare",
        "market": "Harrisburg",
        "flight_start": date(2027, 1, 1),
        "flight_end": date(2027, 3, 31),
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
        "flight_start": date(2027, 3, 1),
        "flight_end": date(2027, 5, 31),
        "checks": check_ashford_two_track,
    },
    "capital_ridge_stacked": {
        "title": "Dental / one audience described by four attributes",
        "vertical": "healthcare",
        "market": "Harrisburg",
        "flight_start": date(2026, 10, 1),
        "flight_end": date(2026, 12, 31),
        "checks": check_capital_ridge_stacked,
    },
    # The live Plaza Motors failure: a stated budget priced against the full
    # avails figure instead of the budget itself.
    "plaza_motors_net_budget": {
        "title": "Auto / stated budget with stated avails / explicit net, no commission",
        "vertical": "auto",
        "market": "DC",
        "flight_start": date(2026, 10, 1),
        "flight_end": date(2026, 12, 31),
        "checks": check_plaza_motors_net_budget,
    },
    # FLOW_REWORK_PLAN.md Phase 4b acceptance case: "gross up the net CPM"
    # is per-line; the checkbox is deck-wide and rep-only. The model must
    # transcribe, never compute.
    "gross_up_verbatim": {
        "title": "Retail / net CPM stated with a gross-it-up instruction / must transcribe, not compute",
        "vertical": "retail",
        "market": "DC",
        "flight_start": date(2026, 10, 1),
        "flight_end": date(2026, 12, 31),
        "checks": check_gross_up_verbatim,
    },
    # The Capital Media incident's phrasing variant: unlike gross_up_verbatim
    # above (ambiguous -- "gross it up" said about the CPM only), these notes
    # explicitly label the BUDGET itself as gross. total_budget_basis exists
    # to catch exactly this distinction.
    "gross_budget_stated": {
        "title": "Home improvement / explicit GROSS total budget + separately-stated net CPM",
        "vertical": "home_improvement",
        "market": "DC",
        "flight_start": date(2026, 10, 1),
        "flight_end": date(2026, 10, 31),
        "checks": check_gross_budget_stated,
    },
    "summit_outside_linear": {
        "title": "Retail / client's existing linear buy on another vendor's station",
        "vertical": "retail",
        "market": "DC",
        "flight_start": date(2027, 1, 1),
        "flight_end": date(2027, 3, 31),
        "checks": check_summit_outside_linear,
    },
    # Model-layer check for the group-selection schema (targeting groups
    # already own the plan): "existing_groups" is populated into
    # st.session_state before the call, standing in for a real avails
    # import that already ran, exactly what build_draft_prompt's own
    # `existing_groups` parameter is for.
    "annapolis_group_selection": {
        "title": "Auto / real avails table already on the proposal / notes name 2 of 4 makes",
        "vertical": "auto",
        "market": "DC",
        "flight_start": date(2026, 9, 1),
        "flight_end": date(2026, 11, 30),
        "existing_groups": _annapolis_existing_groups,
        "checks": check_annapolis_group_selection,
    },
}


def check_common(rep, draft, spec):
    rep.section("Vertical, market, dates, unresolved")
    rep.equal("vertical detected from the trade name", draft.get("vertical"), spec["vertical"])

    # FLOW_REWORK_PLAN.md Phase 5: market/flight_start/flight_end are band
    # INPUTS now, not drafted outputs -- DRAFT_JSON_SCHEMA_EXAMPLE no longer
    # has fields for them, so the model is never asked and should never
    # return them. run() presets st.session_state's own market_choice/
    # flight_start/flight_end from spec (the band, as a rep would have
    # already filled it in) before calling call_claude_draft. What used to
    # be asserted here -- the model resolving a year-less date forward, and
    # naming the right DC/Harrisburg market -- is now proven by the schema's
    # absence, not by inspecting the model's output for it.
    rep.check("market is absent from the draft (never asked for it)",
              "market" not in draft or draft.get("market") is None, draft.get("market"))
    rep.check("flight_start is absent from the draft (never asked for it)",
              "flight_start" not in draft or draft.get("flight_start") is None,
              draft.get("flight_start"))
    rep.check("flight_end is absent from the draft (never asked for it)",
              "flight_end" not in draft or draft.get("flight_end") is None,
              draft.get("flight_end"))

    unresolved = draft.get("unresolved") or []
    internal = draft.get("unresolved_internal") or []
    # The concrete "what to measure" signal from FLOW_REWORK_PLAN.md Phase 5:
    # every one of these fixtures' notes used to make the model resolve a
    # year-less or ambiguous DATE itself and flag its own resolution -- "no
    # year given," "no exact start/end date," "resolved to X as the next
    # upcoming occurrence." With the band supplying the flight directly, the
    # model never touches dates at all, so this specific phrasing pattern
    # should be structurally gone. This is a narrower net than "any mention
    # of a market" on purpose: a genuine band-vs-notes DISAGREEMENT (the
    # contradiction rule -- band wins, flagged to unresolved_internal naming
    # both) is a different, correct thing to see, not this regression.
    # A live run surfaced the real distinction this needs: the OLD regression
    # was the model computing its OWN date resolution with no reference to
    # any pre-existing frame at all ("was assumed to run through November
    # 30, 2026", "resolved to October 1 - December 31, 2026 as the next
    # upcoming occurrence"). The NEW, correct behavior can still legitimately
    # use words like "no exact" or "no end date" while describing the
    # NOTES' own vagueness -- but it does so by referencing the frame it was
    # given ("The frame has November 30 as the end date; confirm that works")
    # rather than inventing one. So: flag only an item that uses
    # date-resolution language WITHOUT ever mentioning the frame it should
    # have been handed -- that combination is what's now supposed to be
    # impossible.
    # Narrowed after a live run: bare "no specific" (and "no exact" with no
    # noun) also match phrases about geo-targeting method, avails figures,
    # etc. that have nothing to do with dates -- e.g. "no specific DMA
    # targeting method was discussed." These must name a date/year/quarter
    # explicitly to count.
    date_resolution_words = ("no year", "no exact date", "no exact start",
                             "no exact end", "no specific date", "no specific start",
                             "no specific end", "next upcoming", "was assumed to run",
                             "resolved to", "no end date", "no start date")
    frame_reference_words = ("frame", "band")
    date_items = [item for item in unresolved + internal
                 if any(w in item.lower() for w in date_resolution_words)
                 and not any(w in item.lower() for w in frame_reference_words)]
    rep.check("no date-resolution corrections that don't reference the given "
              "frame (the band supplies the flight now -- a residual mention "
              "is fine only when it's checking the frame against loose notes, "
              "never the model computing its own)",
              not date_items, date_items)

    market_items = [item for item in unresolved + internal if "market" in item.lower()
                    or "dma" in item.lower()]
    if market_items:
        print("    ....  market/DMA mentioned in review (not necessarily wrong -- could "
              "be a genuine band-vs-notes disagreement, the contradiction rule working "
              "as intended, not the old date-inference regression):")
        for item in market_items:
            print(f"    ....    - {item}")

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
    if spec.get("existing_groups"):
        st.session_state["targeting_groups"] = spec["existing_groups"]()
    # FLOW_REWORK_PLAN.md Phase 5: the band is filled in before drafting runs,
    # same order the real app now enforces -- call_claude_draft reads these
    # straight from session_state.
    st.session_state["market_choice"] = spec["market"]
    st.session_state["flight_start"] = spec["flight_start"]
    st.session_state["flight_end"] = spec["flight_end"]
    st.session_state["avails_basis"] = app.AVAILS_BASIS_MONTHLY
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
