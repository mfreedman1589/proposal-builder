"""Tier 1 -- the offline drafting regression suite.

    python tests/test_draft_regression.py            # both scenarios
    python tests/test_draft_regression.py hvac       # one, by fixture prefix
    python tests/test_draft_regression.py --keep     # leave the built decks on disk

Runs the whole pipeline below the network: a recorded model response from
tests/fixtures/*.draft.json goes through the *real* `apply_draft_to_form`,
the drafted state is loaded into the *real* form via AppTest, and Generate
runs the *real* assembly. Nothing is reimplemented here -- the point is to
exercise the same code the app runs, so this suite fails when the app would.

The one thing it deliberately does not test is the model itself. What Claude
returns for a given set of notes is Tier 2's job
(tests/test_draft_live.py), which costs money and needs a key; this tier is
free, offline, deterministic, and safe to run on every change.

Every assertion prints PASS or FAIL with the actual value, and the process
exits non-zero if anything failed.
"""

import argparse
import copy
import json
import os
import re
import sys
import zipfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"

# The app writes files relative to the working directory (cached decks, the
# local master fallback), so run as if launched from the repo root.
os.chdir(REPO)

import app                                     # noqa: E402
import assembly                                # noqa: E402
import slide_map                               # noqa: E402

TOKEN_RE = re.compile(r"\{\{[A-Z0-9_]+\}\}")


# --------------------------------------------------------------------------
# reporting
# --------------------------------------------------------------------------

class Report:
    """Pass/fail per assertion, with the actual value whenever it fails --
    the whole reason this exists is to show what drifted."""

    def __init__(self):
        self.passed = 0
        self.failed = []
        self.skipped = []
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

    def close(self, label, actual, expected, tol):
        return self.check(f"{label} (within {tol})", abs(actual - expected) <= tol,
                          actual, expected)

    def skip(self, label, why):
        self.skipped.append(f"[{self.scenario}] {label}")
        print(f"    SKIP  {label} -- {why}")


# --------------------------------------------------------------------------
# scenarios
# --------------------------------------------------------------------------

class Scenario:
    def __init__(self, name, fixture, *, vertical, agency, options, sport_key,
                 flat_fee, dynamic_creative, attribution_on, attribution_off,
                 cpm_overrides):
        self.name = name
        self.fixture = fixture
        self.vertical = vertical
        self.agency = agency
        self.options = options                  # [(name-ish, budget), ...]
        self.sport_key = sport_key              # None when the plan has no sports
        self.flat_fee = flat_fee                # expected flat-fee amount, or None
        self.dynamic_creative = dynamic_creative
        self.attribution_on = attribution_on
        self.attribution_off = attribution_off
        self.cpm_overrides = cpm_overrides      # {product_label_prefix: cpm}

    @property
    def notes_path(self):
        return FIXTURES / f"{self.fixture}.txt"

    @property
    def draft_path(self):
        return FIXTURES / f"{self.fixture}.draft.json"

    def draft(self):
        raw = json.loads(self.draft_path.read_text(encoding="utf-8"))
        raw.pop("_comment", None)
        return raw


SCENARIOS = [
    Scenario(
        "HVAC / two options / agency / sports",
        "hvac_two_option",
        vertical="home_improvement",
        agency=True,
        options=[("$50K Plan", 50000), ("$75K Plan", 75000)],
        sport_key="nfl_reg",
        flat_fee=850,
        dynamic_creative=True,
        attribution_on=["sales_attribution"],
        attribution_off=["brand_lift"],
        cpm_overrides={"Premion Streaming TV": 28.0},
    ),
    Scenario(
        "Dental / single option / direct / no sports",
        "dental_single_option",
        vertical="healthcare",
        agency=False,
        options=[(None, 30000)],
        sport_key=None,
        flat_fee=None,
        dynamic_creative=False,
        attribution_on=[],
        attribution_off=["brand_lift", "sales_attribution"],
        cpm_overrides={},
    ),
]


# --------------------------------------------------------------------------
# tier 1a -- the resolver, via the real apply_draft_to_form
# --------------------------------------------------------------------------

class _StubSt:
    """apply_draft_to_form only touches st.session_state, so a plain dict is
    enough to run it outside a Streamlit runtime."""

    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply_draft(draft):
    """Run the real apply_draft_to_form and return the session_state it
    produced -- i.e. exactly what the form is loaded with after a draft."""
    real = app.st
    stub = _StubSt()
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return stub.session_state


def n_months_of(state):
    months = app.month_list(state["flight_start"], state["flight_end"])
    return max(1, len(months))


def full_flight_cost(option, n_months):
    """Each row's own contribution, flat fees never scaled -- the same rule
    compute_plan_totals applies."""
    total = 0.0
    for row in option["rows"]:
        if app.is_flat_fee_row(row):
            total += float(row["Cost"])
        elif option["breakout"] == app.BREAKOUT_MONTHLY:
            total += float(row["Cost"]) * n_months
        else:
            total += float(row["Cost"])
    return total


def check_draft_state(rep, scn, draft, state):
    rep.section("Drafted form state")

    vertical_key = app.VERTICALS.get(state.get("vertical_choice"))
    rep.equal("vertical resolves from the trade name", vertical_key, scn.vertical)
    rep.equal("agency_involved", state.get("agency_involved"), scn.agency)

    markup = 1.15 if scn.agency else 1.0
    rep.equal("markup recorded on the form", state.get("media_plan_markup"), markup)

    options = state.get("plan_options") or []
    rep.equal("option count", len(options), len(scn.options))

    # Budgets live in the draft (the form holds resolved rows), so the tie-out
    # is row costs -> stated budget.
    drafted = app.drafted_options(draft)
    rep.equal("every drafted option carries its own total_budget",
              [o["total_budget"] for o in drafted],
              [b for _, b in scn.options])

    n_months = n_months_of(state)
    print(f"    ....  flight: {state.get('flight_start')} -> {state.get('flight_end')} "
          f"({n_months} months), markup {markup}")

    for option, (want_name, want_budget) in zip(options, scn.options):
        label = option["name"]
        if want_name is not None:
            rep.equal(f"option named {want_name!r}", label, want_name)

        rep.close(f"{label}: line costs sum to the stated budget",
                  full_flight_cost(option, n_months), float(want_budget), 1.0)

        rate_rows = [r for r in option["rows"] if not app.is_flat_fee_row(r)]
        rep.check(f"{label}: no rate line has zero impressions",
                  all(float(r["Impressions"]) > 0 for r in rate_rows),
                  [(r["Tactic"], r["Impressions"]) for r in rate_rows
                   if float(r["Impressions"]) <= 0])
        rep.check(f"{label}: no rate line has zero cost",
                  all(float(r["Cost"]) > 0 for r in rate_rows),
                  [(r["Tactic"], r["Cost"]) for r in rate_rows if float(r["Cost"]) <= 0])

        # Gross markup has to be visible in the arithmetic itself, not just in
        # the flag -- impressions = cost / (cpm * markup) * 1000.
        bad = []
        for r in rate_rows:
            want = app.impressions_from_cost(r["Cost"], r["CPM"], markup)
            if abs(float(r["Impressions"]) - want) > 1:
                bad.append((r["Tactic"], r["Impressions"], want))
        rep.check(f"{label}: {'gross' if scn.agency else 'net'} markup applied to every rate line",
                  not bad, bad)

        if scn.sport_key:
            want_label, want_cpm = app.line_product_spec(f"sport:{scn.sport_key}")
            sport_rows = [r for r in option["rows"] if r["Tactic"].startswith(want_label)]
            if rep.check(f"{label}: a {want_label} line exists", len(sport_rows) == 1,
                         [r["Tactic"] for r in option["rows"]]):
                rep.equal(f"{label}: sports CPM is the ratecard rate, not the Premion default",
                          sport_rows[0]["CPM"], want_cpm)
        else:
            sports = [r["Tactic"] for r in option["rows"] if r["Tactic"].startswith("Live Sports")]
            rep.check(f"{label}: no sports lines", not sports, sports)

        if scn.flat_fee is not None:
            fees = [r for r in option["rows"] if app.is_flat_fee_row(r)]
            if rep.check(f"{label}: a flat-fee line exists", len(fees) == 1,
                         [r["Tactic"] for r in option["rows"]]):
                rep.equal(f"{label}: flat fee amount", float(fees[0]["Cost"]), float(scn.flat_fee))
                rep.equal(f"{label}: flat fee carries no impressions",
                          float(fees[0]["Impressions"]), 0.0)

        for prefix, want_cpm in scn.cpm_overrides.items():
            rows = [r for r in option["rows"] if r["Tactic"].startswith(prefix)]
            if rep.check(f"{label}: a {prefix} line exists", bool(rows),
                         [r["Tactic"] for r in option["rows"]]):
                rep.equal(f"{label}: {prefix} uses the negotiated CPM",
                          rows[0]["CPM"], want_cpm)

    rep.section("Attribution, audiences, dates")

    for key in scn.attribution_on:
        rep.equal(f"attribution {key} on", state.get(key), True)
    for key in scn.attribution_off:
        rep.equal(f"attribution {key} off", state.get(key), False)
    rep.check("web attribution present in the draft (no toggle -- always included)",
              "web" in (draft.get("attribution") or []), draft.get("attribution"))

    rep.equal("dynamic creative product", state.get("dynamic_creative"), scn.dynamic_creative)

    matched, unmatched = app.validate_segments(
        [a["segment"] for a in (draft.get("audiences") or [])])
    rep.check("every drafted audience validates against the catalog", not unmatched, unmatched)
    if matched:
        seeded = [r["Audience"] for r in (state.get("avails_seed_rows") or [])]
        rep.check("validated audiences reached the avails table",
                  all(m in seeded for m in matched), seeded, matched)

    start = state.get("flight_start")
    rep.check("flight start is not in the past", start is not None and start >= date.today(),
              start, f">= {date.today()}")
    rep.check("flight end follows flight start",
              state.get("flight_end") is not None and state["flight_end"] >= start,
              state.get("flight_end"))

    unresolved = state.get("draft_unresolved") or []
    rep.check("unresolved is non-empty", bool(unresolved), unresolved)
    rep.check("unresolved has no duplicates",
              len(unresolved) == len(set(unresolved)),
              [u for u in unresolved if unresolved.count(u) > 1])


# --------------------------------------------------------------------------
# tier 1b -- the real form and the real assembly, via AppTest
# --------------------------------------------------------------------------

def build_deck(rep, state):
    """Load the drafted state into the real form, press Generate, and capture
    the presentation the real assembly produced.

    assembly.personalize/append_case_studies are wrapped rather than replaced:
    they run for real and hand back the objects they built. app.py holds
    `assembly` as a module reference and looks the attribute up at call time,
    so wrapping here is visible to the app even though AppTest execs app.py
    as its own module.
    """
    from streamlit.testing.v1 import AppTest

    captured = {"plan_slide_ids": [], "plan_titles": []}
    real_personalize = assembly.personalize
    real_append = assembly.append_case_studies
    real_fill_plan = assembly._fill_media_plan_slide

    def spy_fill_plan(slide, option):
        # Which slides are plan slides, straight from the code that fills
        # them. Matching on the rendered title instead would be ambiguous: a
        # single-option deck's plan title *is* the proposal title, which also
        # appears on the cover.
        result = real_fill_plan(slide, option)
        captured["plan_slide_ids"].append(slide.slide_id)
        captured["plan_titles"].append(option.get("plan_title"))
        return result

    def spy_personalize(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        captured["fill_data"] = fill_data
        captured["warnings"] = warnings
        return warnings

    def spy_append(prs, sources):
        before = {s.slide_id for s in prs.slides}
        count = real_append(prs, sources)
        captured["case_study_slide_ids"] = [s.slide_id for s in prs.slides
                                            if s.slide_id not in before]
        captured["case_study_sources"] = sources
        return count

    assembly.personalize = spy_personalize
    assembly.append_case_studies = spy_append
    assembly._fill_media_plan_slide = spy_fill_plan
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            rep.check("form renders with the drafted state", False, str(at.exception[0].value))
            return None, at
        rep.check("form renders with the drafted state", True)

        # The regression that started all this: a seed-key mismatch silently
        # rebuilds every option's rows as $0 seeds on the first render.
        rendered = at.session_state["plan_options"]
        zeroed = [(o["name"], r["Tactic"]) for o in rendered for r in o["rows"]
                  if not app.is_flat_fee_row(r) and not float(r["Cost"])]
        rep.check("drafted rows survive the form's own reseed check", not zeroed, zeroed)

        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not generate:
            rep.check("Generate button present", False, [b.label for b in at.button])
            return None, at
        generate[0].click().run()
        if at.exception:
            rep.check("Generate runs without raising", False, str(at.exception[0].value))
            return None, at
        rep.check("Generate runs without raising", True)
    finally:
        assembly.personalize = real_personalize
        assembly.append_case_studies = real_append
        assembly._fill_media_plan_slide = real_fill_plan

    return captured, at


def slide_text(slide):
    return slide_map.extract_slide_text(slide)


def check_deck(rep, scn, captured, keep):
    rep.section("Assembled deck")

    prs = captured.get("prs")
    if prs is None:
        rep.check("assembly produced a presentation", False, None)
        return

    slides = list(prs.slides)
    texts = [slide_text(s) for s in slides]
    print(f"    ....  {len(slides)} slides assembled")

    # 1. no unfilled tokens
    leftover = []
    for i, text in enumerate(texts, start=1):
        for tok in TOKEN_RE.findall(text or ""):
            leftover.append((i, tok))
    rep.check("no unfilled {{TOKEN}}s remain", not leftover, leftover[:10])

    # 2. no section dividers
    derived = slide_map.build_slide_map_from_prs(prs)
    dividers = [n for n, key in derived.items() if key == slide_map.SECTION_DIVIDER_KEY]
    rep.check("no section dividers in the deck", not dividers, dividers)

    # 3. one plan slide per option, in order
    option_payloads = captured["fill_data"]["media_plan_options"]
    by_id = {s.slide_id: i for i, s in enumerate(slides)}
    rep.equal("one plan slide per option",
              len(captured["plan_slide_ids"]), len(option_payloads))
    rep.equal("each option's own title reached its slide",
              captured["plan_titles"], [p["plan_title"] for p in option_payloads])
    flat = [by_id[sid] for sid in captured["plan_slide_ids"] if sid in by_id]
    rep.equal("every plan slide is still in the deck", len(flat), len(captured["plan_slide_ids"]))
    if flat:
        rep.check("plan slides appear in option order", flat == sorted(flat), flat)
        rep.check("plan slides are consecutive",
                  flat == list(range(flat[0], flat[0] + len(flat))), flat)
        # Each option's title belongs to its own slide and no other.
        misplaced = []
        for pos, payload in zip(flat, option_payloads):
            title = payload["plan_title"]
            if title and title not in (texts[pos] or ""):
                misplaced.append((pos + 1, title))
        rep.check("each plan slide carries its own option's title", not misplaced, misplaced)

    # 4. case studies contiguous, immediately before the first plan slide
    cs_ids = captured.get("case_study_slide_ids") or []
    if not cs_ids:
        rep.skip("case studies sit immediately before the first plan slide",
                 "none were selected (no vault match or vault unreachable)")
    elif not flat:
        rep.skip("case studies sit immediately before the first plan slide",
                 "no plan slide was located")
    else:
        by_id = {s.slide_id: i for i, s in enumerate(slides)}
        idx = sorted(by_id[sid] for sid in cs_ids if sid in by_id)
        first_plan = min(flat)
        rep.check("case study slides are contiguous",
                  idx == list(range(idx[0], idx[0] + len(idx))), idx)
        rep.equal("case studies end immediately before the first plan slide",
                  idx[-1] + 1 if idx else None, first_plan)

    # 5. the plan table clears the Included-with-Campaign graphic
    rep.check("assembly reported no overflow warnings",
              not captured.get("warnings"), captured.get("warnings"))
    measured = 0
    for i in flat:
        slide = slides[i]
        table_shape = assembly._find_table_shape(slide)
        if table_shape is None:
            continue
        bottom = assembly.table_bottom(slide)
        floor = assembly._content_floor(slide, table_shape)
        if bottom is None or floor is None:
            continue
        measured += 1
        rep.check(f"slide {i + 1}: table bottom clears the content floor "
                  f"({bottom / 914400:.2f}in <= {floor / 914400:.2f}in)",
                  bottom <= floor, f"{bottom / 914400:.3f}in", f"<= {floor / 914400:.3f}in")
    if not measured:
        rep.skip("plan table measured against the content floor", "no table shape found")

    # 6. the dynamic creative slide, when the product is on
    if scn.dynamic_creative:
        keys = set(derived.values())
        rep.check("dynamic creative slide included", "dynamic_creative" in keys,
                  sorted(k for k in keys if "dynamic" in k))

    # 7. package integrity
    out = REPO / "tests" / f"_built_{scn.fixture}.pptx"
    prs.save(str(out))
    try:
        with zipfile.ZipFile(out) as zf:
            rep.check("zip integrity", zf.testzip() is None, zf.testzip())
            names = zf.namelist()
            dupes = sorted({n for n in names if names.count(n) > 1})
            rep.check("no duplicate partnames", not dupes, dupes)
            slide_parts = [n for n in names if n.startswith("ppt/slides/slide")]
            rep.equal("slide parts match slide count", len(slide_parts), len(slides))
        print(f"    ....  wrote {out.name} ({out.stat().st_size / 1048576:.1f} MiB)")
    finally:
        if not keep:
            out.unlink(missing_ok=True)


# --------------------------------------------------------------------------

def run(scn, rep, keep):
    rep.scenario = scn.name
    print("\n" + "=" * 78)
    print(f"SCENARIO  {scn.name}")
    print(f"fixture   {scn.notes_path.name} + {scn.draft_path.name}")
    print("=" * 78)

    for path in (scn.notes_path, scn.draft_path):
        if not path.exists():
            rep.check(f"fixture {path.name} exists", False, str(path))
            return

    draft = scn.draft()
    state = apply_draft(draft)
    check_draft_state(rep, scn, draft, state)

    captured, _ = build_deck(rep, state)
    if captured:
        check_deck(rep, scn, captured, keep)


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("only", nargs="?", help="run only scenarios whose fixture name starts with this")
    parser.add_argument("--keep", action="store_true", help="keep the generated .pptx files")
    args = parser.parse_args()

    scenarios = [s for s in SCENARIOS
                 if not args.only or s.fixture.startswith(args.only)]
    if not scenarios:
        print(f"No scenario matches {args.only!r}. Known: "
              f"{', '.join(s.fixture for s in SCENARIOS)}")
        return 2

    rep = Report()
    for scn in scenarios:
        run(scn, rep, args.keep)

    print("\n" + "=" * 78)
    print(f"{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped")
    for f in rep.failed:
        print(f"  FAILED  {f}")
    for s in rep.skipped:
        print(f"  SKIPPED {s}")
    print("=" * 78)
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
