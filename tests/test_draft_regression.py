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
import tempfile
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
import db                                      # noqa: E402
import slide_map                               # noqa: E402

TOKEN_RE = re.compile(r"\{\{[A-Z0-9_]+\}\}")

# Who the suite runs as. Never written to team_members -- the identity step
# is bypassed by setting session_state directly, the same way the password
# gate is.
TEST_USER = "Regression Suite"

# Where the stubbed logo upload/download round trip keeps its bytes.
LOGO_DIR = Path(tempfile.gettempdir()) / "premion_test_logos"
LOGO_DIR.mkdir(parents=True, exist_ok=True)


def make_test_logo():
    """A small, visually unique PNG.

    Deliberately not placeholder_logo.png: the point is to prove the deck
    carries *this* image, and reusing the placeholder would make a logo that
    silently fell back indistinguishable from one that survived.
    """
    path = LOGO_DIR / "roundtrip_logo.png"
    if not path.exists():
        from PIL import Image
        Image.new("RGB", (240, 80), (0x2E, 0x86, 0xC1)).save(path)
    return str(path)


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
                 cpm_overrides, assemble=True, custom_audience_cap=None,
                 round_trip=False):
        self.name = name
        self.fixture = fixture
        # Generate -> Load from history -> regenerate, comparing the two
        # decks. Doubles the scenario's runtime, so it's opt-in per scenario
        # rather than universal.
        self.round_trip = round_trip
        # Resolver-only scenarios skip Generate: building a deck takes ~40s
        # and 20MB to re-prove invariants the other scenarios already cover.
        self.assemble = assemble
        # (expected kept segment, expected segments moved to unresolved)
        self.custom_audience_cap = custom_audience_cap
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
        # The hardest case to round-trip: two options, sports, a flat fee, a
        # negotiated CPM and case studies all have to survive rehydration.
        round_trip=True,
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
    Scenario(
        "Custom audience cap / three customs returned",
        "custom_audience_cap",
        vertical="auto",
        agency=False,
        options=[(None, 40000)],
        sport_key=None,
        flat_fee=None,
        dynamic_creative=False,
        attribution_on=[],
        attribution_off=["brand_lift", "sales_attribution"],
        cpm_overrides={},
        assemble=False,
        custom_audience_cap=(
            "AUTO Body Style Pickup and SUV",                    # first custom wins
            ["AUTO Body Style Sedan", "AUTO Body Style Minivan"],  # moved to unresolved
        ),
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
    seeded = [r["Audience"] for r in (state.get("avails_seed_rows") or [])]

    if scn.custom_audience_cap:
        # The limit is enforced in Python, not asked for in the prompt, so it
        # holds for any response -- including this fixture's deliberately
        # non-compliant one.
        want_kept, want_dropped = scn.custom_audience_cap
        catalog = app.load_audience_catalog()
        rfp = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
        customs = [s for s in seeded if not rfp.get(s, True)]
        rep.equal("exactly one custom audience survives", len(customs), 1)
        rep.equal("the first custom segment is the one kept", customs, [want_kept])
        rep.check("the dropped customs are gone from the avails table",
                  not [s for s in want_dropped if s in seeded], seeded)
        rep.check("RFP-selectable segments were untouched",
                  all(s in seeded for s in matched if rfp.get(s, True)), seeded)
        note = [u for u in (state.get("draft_unresolved") or []) if "custom" in u.lower()]
        if rep.check("the cap is explained in unresolved", bool(note),
                     state.get("draft_unresolved")):
            rep.check("it names every segment it left out",
                      all(s in note[0] for s in want_dropped), note[0], want_dropped)
            rep.check("it names the one it kept", want_kept in note[0], note[0])
            print(f"    ....  {note[0]}")
    elif matched:
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
    real_log = db.log_proposal
    real_upload_logo = db.upload_proposal_logo
    real_proposal_logo = db.proposal_logo

    def spy_upload_logo(proposal_key, data, filename):
        # Stand in for the storage round trip: write the bytes where the real
        # download cache would have put them, and hand back a path shaped
        # like a storage key. Keeps the logo assertions offline while still
        # exercising every line of app code that moves a logo around.
        target = LOGO_DIR / f"{proposal_key}_{filename}"
        target.write_bytes(data)
        captured["uploaded_logo"] = str(target)
        return f"logos/{target.name}", None

    def spy_proposal_logo(storage_path):
        return str(LOGO_DIR / Path(storage_path).name)

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        # Capture the payload instead of inserting it. Without this the suite
        # writes a real row to the production proposals table on every run --
        # it did, until this was noticed while building the round-trip case.
        # It's also exactly the form_json the History page would later read
        # back, so the round-trip gets its input for free.
        captured["log"] = {
            "id": "00000000-0000-0000-0000-000000000000",
            "client_name": client_name, "vertical": vertical, "market": market,
            "form_json": json.loads(json.dumps(db._json_safe(form_json), default=str)),
            "deck_version_id": kwargs.get("deck_version_id"),
            "output_filename": kwargs.get("output_filename"),
            "parent_proposal_id": kwargs.get("parent_proposal_id"),
            "revision_label": kwargs.get("revision_label"),
            "logo_storage_path": kwargs.get("logo_storage_path"),
            "created_by": kwargs.get("created_by"),
        }
        return "00000000-0000-0000-0000-000000000000", None

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
    db.log_proposal = spy_log
    db.upload_proposal_logo = spy_upload_logo
    db.proposal_logo = spy_proposal_logo
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        # Past the identity step. Set rather than stubbed so the attribution
        # actually flows through to the logged row, which is worth asserting.
        at.session_state["current_user"] = TEST_USER
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
        db.log_proposal = real_log
        db.upload_proposal_logo = real_upload_logo
        db.proposal_logo = real_proposal_logo

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

def rehydrate(row, revision_default=None):
    """Run the real rehydrate_proposal_into_form against a stub session,
    returning the session_state it produced -- what the form is loaded with
    after "Load into form"."""
    real = app.st
    real_logo = db.proposal_logo
    stub = _StubSt()
    app.st = stub
    db.proposal_logo = lambda storage_path: str(LOGO_DIR / Path(storage_path).name)
    try:
        notes = app.rehydrate_proposal_into_form(
            row, parent_proposal_id=row["id"], revision_default=revision_default)
    finally:
        app.st = real
        db.proposal_logo = real_logo
    return stub.session_state, notes


def deck_blobs(prs):
    """Every image blob in a presentation, for checking a specific picture
    actually made it in."""
    return {part.blob for part in prs.part.package.iter_parts()
            if getattr(part, "content_type", "").startswith("image/")}


def deck_shape(captured, prs_slides=None):
    """The invariants a round-trip has to preserve, as a comparable dict."""
    prs = captured["prs"]
    slides = list(prs.slides)
    by_id = {s.slide_id: i for i, s in enumerate(slides)}
    options = captured["fill_data"]["media_plan_options"]
    return {
        "slide_count": len(slides),
        "option_count": len(options),
        "plan_titles": list(captured["plan_titles"]),
        "plan_positions": [by_id[sid] for sid in captured["plan_slide_ids"] if sid in by_id],
        "case_study_count": len(captured.get("case_study_slide_ids") or []),
        "rows": [[(r["tactic"], r["impressions"], r["cost"]) for r in o["rows"]] for o in options],
        "totals": [(o["total_impressions"], o["total_cost"]) for o in options],
        "unfilled_tokens": sorted({t for s in slides
                                   for t in TOKEN_RE.findall(slide_text(s) or "")}),
        # Name only: the path differs between runs, what must match is which
        # image the deck was built with.
        "logo": Path(str(captured["fill_data"]["logo_path"])).name,
    }


def check_round_trip(rep, scn, first):
    """Generate -> Load from history -> regenerate untouched -> compare.

    Any field that doesn't survive rehydration shows up here as a difference
    between the two decks. This is the check that makes "Load into form"
    trustworthy: the promise is that loading and regenerating an untouched
    proposal reproduces it, and the only honest way to know is to do it.
    """
    rep.section("Round trip: generate -> load -> regenerate")

    logged = first.get("log")
    if not rep.check("the proposal was logged with a form_json payload", bool(logged),
                     None if logged else "log_proposal was never called"):
        return

    state, notes = rehydrate(logged, revision_default="Revision 2")
    for note in notes:
        print(f"    ....  rehydration note: {note}")

    # A few fields worth asserting directly, since a silent default here
    # would still produce a matching deck for the wrong reason.
    form = logged["form_json"]
    rep.equal("client name survives", state.get("client_name"), logged["client_name"])
    rep.equal("proposal title survives", state.get("proposal_title"),
              form.get("proposal_title"))
    rep.equal("agency toggle survives", state.get("agency_involved"),
              form.get("agency_involved"))
    rep.equal("option count survives", len(state.get("plan_options") or []),
              len(form.get("plan_options") or []))
    rep.check("every restored plan row is marked dirty",
              all(all(o["dirty"]) for o in state.get("plan_options") or []),
              [o["dirty"] for o in state.get("plan_options") or []])
    rep.equal("parent link is set for the regeneration",
              state.get("history_parent_id"), logged["id"])

    # --- the logo ---------------------------------------------------------
    logo_bytes = Path(make_test_logo()).read_bytes()
    rep.check("the logo was recorded on the logged row",
              bool(logged.get("logo_storage_path")), logged.get("logo_storage_path"))
    rep.check("form_json records that a logo was used",
              bool(form.get("logo_used")), form.get("logo_used"))
    rep.equal("the first deck was built with the real logo, not the placeholder",
              Path(str(first["fill_data"]["logo_path"])).name, Path(make_test_logo()).name)
    rep.check("the logo's bytes are in the first deck",
              logo_bytes in deck_blobs(first["prs"]), None,
              "the CLIENT_LOGO picture to carry the uploaded image")
    rep.equal("the stored logo is restored on load",
              state.get("restored_logo_storage_path"), logged.get("logo_storage_path"))
    rep.check("no 'logo missing' state after a successful restore",
              not state.get("restored_logo_missing"), state.get("restored_logo_missing"))
    rep.equal("revision label default is prefilled on load",
              state.get("revision_label"), "Revision 2")
    rep.equal("the proposal is attributed to the signed-in user",
              logged.get("created_by"), TEST_USER)

    # The seed key is the one that silently wipes rows when it disagrees.
    real = app.st
    stub = _StubSt()
    stub.session_state.update(state)
    app.st = stub
    try:
        derived = str(app.read_seed_selections())
    finally:
        app.st = real
    rep.equal("rehydrated seed key matches what the form will derive",
              state.get("_product_seed_key"), derived)

    second, _ = build_deck(rep, state)
    if not second or "prs" not in second:
        rep.check("the reloaded proposal regenerates", False, None)
        return
    rep.check("the reloaded proposal regenerates", True)

    before, after = deck_shape(first), deck_shape(second)
    for field in ("slide_count", "option_count", "plan_titles", "plan_positions",
                  "case_study_count", "rows", "totals", "unfilled_tokens", "logo"):
        rep.equal(f"round trip preserves {field}", after[field], before[field])
    rep.check("the logo's bytes are in the regenerated deck too",
              logo_bytes in deck_blobs(second["prs"]), None,
              "the rebuilt deck to carry the same logo image")


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

    if scn.round_trip:
        # Give the first build a logo, the way a proposal loaded from history
        # carries one. A file_uploader can't be driven from AppTest, and this
        # is the path that actually matters: the logo has to survive being
        # logged, rehydrated and used again.
        logo = make_test_logo()
        state["restored_logo_path"] = logo
        state["restored_logo_storage_path"] = f"logos/{Path(logo).name}"

    if not scn.assemble:
        print("\n    ....  resolver-only scenario -- assembly covered by the others")
        return

    captured, _ = build_deck(rep, state)
    if captured:
        check_deck(rep, scn, captured, keep)
        if scn.round_trip:
            check_round_trip(rep, scn, captured)


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
