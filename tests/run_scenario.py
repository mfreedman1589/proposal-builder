"""The reproduce-and-verify loop, in one command.

    python tests/run_scenario.py plaza_motors
    python tests/run_scenario.py plaza_motors --keep --no-restart

For a named scenario: restarts the local app process on current code (see
app_process.py -- this is what "couldn't reproduce, probably a stale
process" turns into "confirmed, and here's proof either way" instead of a
guess), builds the scenario's proposal through the real form, presses the
real Generate button, renders the resulting deck through PowerPoint, and
runs the scenario's own pass/fail checks against the rendered result.
Rendered images are saved and their path is always printed, so "go look"
never means re-running anything.

**Why AppTest, not a real browser, drives the form here.** A Streamlit
`AppTest` run executes app.py directly in a fresh Python process -- no
module-caching staleness is possible, by construction, which is exactly the
property this tool needs for the scenario-logic half of the loop. It is also
the mechanism this whole test suite already trusts for "did the real form,
the real button, the real assembly produce the right deck"
(test_group_scenarios.py, render_scenarios.py). It is NOT, on its own,
a test of the actual long-running `streamlit run` process a person's browser
talks to -- that's a genuinely different thing, and it's what app_process.py
restarting a real server (and, separately, a real browser session against
it) is for. Real browser automation (Playwright/claude-in-chrome) can't be
embedded inside this unattended script -- those tools are only available
interactively, one call at a time, to whatever is driving the conversation.
So the split is deliberate: this script is the self-sufficient, always-the-
same-answer half; a real-browser walkthrough against the server
app_process.py starts is a separate, complementary check, run by hand or by
an agent with browser tool access, not by this file.

Scenario checks live next to their builders below (`SCENARIOS`), each
returning True on success and printing what it found either way -- pass/fail
is never inferred from the absence of an exception alone, since a scenario
whose form silently accepted a wrong value is not an exception.
"""
import argparse
import copy
import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app                                       # noqa: E402
import app_process                               # noqa: E402
import assembly                                  # noqa: E402
import db                                        # noqa: E402
import deck_render                               # noqa: E402
import geo_resolver                              # noqa: E402
import targeting_map as tm                       # noqa: E402

OUT = deck_render.render_root()


def _apply_draft(draft):
    """Run a draft dict through the real apply_draft_to_form, against an
    empty session -- the same helper tests/test_avails_reach.py and
    tests/render_scenarios.py both use, so a scenario built here behaves
    exactly like one built there."""
    class _Stub:
        def __init__(self):
            self.session_state = {}
            self.secrets = {}

    real, stub = app.st, _Stub()
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
        return dict(stub.session_state)
    finally:
        app.st = real


def _generate(state, extra=None):
    """Drive the real form to a real deck. Returns (prs, fill_data,
    warnings, final_session_state) -- the session_state is AFTER Generate
    runs, so a scenario check can read draft_unresolved, targeting_groups
    as the app actually left them, etc."""
    from streamlit.testing.v1 import AppTest

    captured = {}
    real_personalize, real_log = assembly.personalize, db.log_proposal

    def spy(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        captured["fill_data"] = fill_data
        captured["warnings"] = warnings
        return warnings

    assembly.personalize = spy
    db.log_proposal = lambda *a, **k: ("run_scenario", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=900)
        for key, value in {**(state or {}), **(extra or {})}.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            raise RuntimeError(f"form render raised: {at.exception}")
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not generate:
            raise RuntimeError("no 'Generate proposal' button found")
        generate[0].click().run()
        if at.exception:
            raise RuntimeError(f"Generate raised: {at.exception}")
        # AppTest's session_state proxy can't be dict()'d wholesale (it
        # raises on internal/widget-id entries that were never really
        # "keys") -- every other scenario suite in this project pulls the
        # specific keys it needs instead, so this does too.
        captured["session_state"] = {
            key: (at.session_state[key] if key in at.session_state else default)
            for key, default in (
                ("plan_options", []), ("targeting_groups", []),
                ("draft_unresolved", []), ("draft_unresolved_internal", []))
        }
    finally:
        assembly.personalize, db.log_proposal = real_personalize, real_log
    return captured


# ---------------------------------------------------------------------------
# plaza_motors -- St. Louis auto dealer, $4,000 budget against a stated
# 1.7M-avails audience with reach asked for as a percentage, plus a resolved
# radius group so the targeting map has something to draw. Exercises, in one
# generated deck, the four things from the live Plaza Motors test: market
# lookup registration, the map reaching the targeting slide with boundaries,
# and budget (not avails) driving impressions with reach reported.
# ---------------------------------------------------------------------------
_PLAZA_SEGMENT = "AUTO Intenders"
_PLAZA_BUDGET = 4000
_PLAZA_AVAILS_MONTHLY = 1_700_000
_PLAZA_CPM = 30


def build_plaza_motors():
    draft = {
        "client_name": "Plaza Motors", "vertical": "auto", "market": "DC",
        "geo": "St. Louis, MO DMA",
        "flight_start": "2026-09-01", "flight_end": "2026-09-30",
        "agency_involved": False, "breakout": "monthly",
        "audiences": [{"segment": _PLAZA_SEGMENT, "geo": "St. Louis, MO DMA",
                       "max_avails": _PLAZA_AVAILS_MONTHLY, "avails_basis": "monthly"}],
        "options": None, "total_budget": _PLAZA_BUDGET,
        "media_plan_lines": [
            {"product": "premion_streaming_tv", "audience_track": _PLAZA_SEGMENT,
             "cpm": _PLAZA_CPM, "allocation": {"flat_amount": _PLAZA_BUDGET}},
        ],
        "sports": [], "attribution": [],
        "campaign_specs": {
            "goals": ["Drive showroom traffic for the September clearance event"],
            "audience": ["In-market auto intenders"],
            "geography": ["St. Louis, MO DMA"], "budget": [f"${_PLAZA_BUDGET:,}"],
            "placements": ["15s/30s CTV"], "timing": ["September 2026"]},
        "unresolved": [], "unresolved_internal": [],
    }
    state = _apply_draft(draft)

    # Resolve real geography for the map -- a 10mi radius around a St. Louis
    # zip, exactly like a rep would build in the geo-definition expander.
    # Data-injected rather than replayed through the Radius-mode UI widgets,
    # same deliberate shortcut group_scenario_fixtures.py documents: that
    # interaction sequence is covered elsewhere (tests/test_group_geo_
    # resolution.py), and reproducing it through AppTest widget-by-widget
    # buys nothing here but fragility.
    groups = state.get("targeting_groups") or []
    plaza_group = next((g for g in groups if _PLAZA_SEGMENT in (g.get("terms") or [])), None)
    if plaza_group is None:
        raise RuntimeError(
            f"draft didn't seed a targeting_groups entry for {_PLAZA_SEGMENT!r} -- "
            f"check apply_draft_to_form / sync_targeting_groups wiring before blaming "
            f"anything downstream")
    plaza_group["geo_def"] = {"kind": "radius", "centers": ["63101"], "miles": 10}
    result = geo_resolver.radius_to_zips(["63101"], 10)
    plaza_group["resolved_zips"] = result.resolved
    plaza_group["resolved_markets"] = sorted(
        geo_resolver.zips_to_markets(result.resolved).resolved.keys())
    state["targeting_groups"] = groups
    # sync_targeting_groups (app.py) decides "which side moved" by comparing
    # avails_seed_rows against _groups_rows_applied, the row projection it
    # last wrote FROM groups -- apply_draft_to_form writes avails_seed_rows
    # and targeting_groups together but never sets this marker (a draft
    # never has resolved geography to protect, so it's never needed it
    # before), so without it the very first render here sees "stored_rows
    # != applied_rows (None)", treats that as a real edit, and re-derives
    # targeting_groups from the flat rows -- wiping the resolved_zips just
    # set above before the map ever gets a chance to draw them. Exactly the
    # rehydration bug CLAUDE.md documents, hit here because this scenario
    # combines a draft (for the budget line) with pre-resolved geography (for
    # the map) in one step, which a live session never does in one step.
    state["_groups_rows_applied"] = [dict(r) for r in (state.get("avails_seed_rows") or [])]
    state["include_avails_template"] = True
    return _generate(state)


def check_plaza_motors(built):
    ok = True

    # 1. Market lookup actually registered -- same call app.py itself makes
    # on every run (Section A), in this same process, right after the
    # scenario's own AppTest run exercised it.
    lookup, warning = app.install_market_lookup()
    if lookup and not warning:
        print(f"  PASS  market lookup registered ({len(lookup)} counties, no warning)")
    else:
        ok = False
        print(f"  FAIL  market lookup NOT registered -- {warning}")

    # 2. The map landed on the targeting slide, with boundary outlines
    # actually drawn into the PNG (not merely "a picture exists").
    prs = built.get("prs")
    fill_data = built.get("fill_data") or {}
    map_png = (fill_data.get("avails") or {}).get("map_png")
    if not map_png:
        ok = False
        print("  FAIL  no map_png was computed for this proposal")
    else:
        print(f"  PASS  map_png computed ({len(map_png)} bytes)")
        from io import BytesIO
        from PIL import Image
        img = Image.open(BytesIO(map_png)).convert("RGB")
        colors = {c for _n, c in img.getcolors(maxcolors=img.width * img.height)}
        # app.py always renders the real Generate path's map_png with
        # dark=True (the map-variant slide's background is the deck's own
        # dark gradient, not white -- see app.py's own "dark=True, matching
        # the live Generate handler" comment), so the outline colors actually
        # drawn here are targeting_map._DARK_PALETTE's, not the light-mode
        # COUNTY_OUTLINE_COLOR/STATE_OUTLINE_COLOR module constants. Checking
        # against the light constants on a dark-rendered PNG always reads as
        # "missing" even when the outlines are genuinely there -- confirmed
        # by rendering the same group directly with dark=True and finding
        # the dark palette's own colors present.
        dark_county = tm._DARK_PALETTE["county_outline"][:3]
        dark_state = tm._DARK_PALETTE["state_outline"][:3]
        has_county = dark_county in colors
        has_state = dark_state in colors
        if has_county and has_state:
            print("  PASS  county and state outlines are actually drawn on the map")
        else:
            ok = False
            print(f"  FAIL  boundary outlines missing from the rendered map "
                  f"(county={has_county}, state={has_state})")
        avails_slide = next(
            (s for s in prs.slides if assembly._find_table_shape(s) is not None
             and assembly.targeting_map_region(s) is not None
             and "PRECISION TARGETING" in (assembly.slide_map.extract_slide_text(s) or "").upper()),
            None) if prs is not None else None
        if avails_slide is None:
            ok = False
            print("  FAIL  couldn't find the targeting/avails slide in the assembled deck")
        else:
            # A real map swaps in the map-variant slide outright (assembly.py:
            # "the map variant replaces the standard avails/targeting template
            # outright whenever a targeting map will actually be drawn"), which
            # carries no stock photo or slide-level wordmark of its own (the
            # wordmark lives on the slide MASTER) -- so the map is the ONLY
            # picture. This used to expect 3 (background + wordmark + map),
            # the right count before the map-variant slide existed to swap to;
            # see tests/test_targeting_map.py, which caught the same stale
            # expectation on itself.
            pics = [s for s in avails_slide.shapes if s.shape_type == 13]
            if len(pics) == 1:
                print(f"  PASS  map picture landed on the targeting slide ({len(pics)} pictures)")
            else:
                ok = False
                print(f"  FAIL  expected exactly 1 picture on the targeting slide (the "
                      f"map-variant slide has no stock photo or wordmark of its own), "
                      f"found {len(pics)}")

    # 3. Budget drives impressions (not the 1.7M avails), and the derived
    # reach percentage -- not a driver, just reported -- lands around 8%.
    session_state = built.get("session_state") or {}
    rows = ((session_state.get("plan_options") or [{}])[0] or {}).get("rows") or []
    plaza_row = next((r for r in rows if _PLAZA_SEGMENT in str(r.get("Tactic", ""))
                      or "Premion Streaming TV" in str(r.get("Tactic", ""))), None)
    expected_impressions = app.impressions_from_cost(_PLAZA_BUDGET, _PLAZA_CPM, 1.0)
    if plaza_row is None:
        ok = False
        print("  FAIL  couldn't find the Plaza Motors media plan row")
    else:
        cost_ok = round(float(plaza_row["Cost"])) == _PLAZA_BUDGET
        impressions_ok = round(float(plaza_row["Impressions"])) == round(expected_impressions)
        if cost_ok and impressions_ok:
            print(f"  PASS  line prices from the ${_PLAZA_BUDGET:,} budget at ${_PLAZA_CPM} CPM "
                  f"-> {plaza_row['Impressions']:,.0f} impressions (not from the 1.7M avails)")
        else:
            ok = False
            print(f"  FAIL  line did not price from budget as expected -- cost="
                  f"{plaza_row.get('Cost')}, impressions={plaza_row.get('Impressions')} "
                  f"(expected cost={_PLAZA_BUDGET}, impressions={expected_impressions:,.0f})")
    review = " ".join(
        (session_state.get("draft_unresolved") or []) +
        (session_state.get("draft_unresolved_internal") or []))
    expected_pct = round(expected_impressions / _PLAZA_AVAILS_MONTHLY * 100.0)
    if f"{expected_pct}%" in review:
        print(f"  PASS  reach reported at ~{expected_pct}% (derived, not driving the line)")
    else:
        ok = False
        print(f"  FAIL  expected a ~{expected_pct}% reach note in the review list, "
              f"got: {review or '(nothing)'}")

    return ok


SCENARIOS = {
    "plaza_motors": (build_plaza_motors, check_plaza_motors),
}


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("scenario", help="scenario name (prefix match), e.g. plaza_motors")
    parser.add_argument("--keep", action="store_true", help="leave the rendered .pptx on disk")
    parser.add_argument("--no-restart", action="store_true",
                        help="skip restarting the local app_process.py server first")
    args = parser.parse_args()

    names = [n for n in SCENARIOS if n.startswith(args.scenario)]
    if not names:
        print(f"No scenario matches {args.scenario!r}. Known: {', '.join(SCENARIOS)}")
        return 2
    name = names[0]
    builder, checker = SCENARIOS[name]

    print("=" * 78)
    print(f"REPRODUCE-AND-VERIFY  {name}")
    print("=" * 78)

    if not args.no_restart:
        print("\n-- restarting the local app process on current code --")
        ok, msg = app_process.restart()
        print(("OK  " if ok else "FAIL") + "  " + msg)
        if not ok:
            print("  (continuing anyway -- the AppTest half below doesn't depend on this "
                 "server; it's the immune-to-staleness half by construction. See this "
                 "file's own docstring.)")
    else:
        print("\n-- skipping app-process restart (--no-restart) --")

    print(f"\n-- building and generating '{name}' through the real form (AppTest) --")
    try:
        built = builder()
    except Exception as exc:                                        # noqa: BLE001
        print(f"FAIL  scenario build/generate raised: {type(exc).__name__}: {exc}")
        return 1
    for warning in built.get("warnings") or []:
        print(f"  warning: {warning}")

    print(f"\n-- rendering the generated deck --")
    OUT.mkdir(parents=True, exist_ok=True)
    scenario_dir = OUT / name
    if not deck_render.renderer_available():
        print("  SKIP -- rendering needs Windows with PowerPoint and pywin32; "
             "checks below still run against the in-memory deck.")
        saved = None
    else:
        try:
            images, saved = deck_render.render_presentation(built["prs"], scenario_dir, name=name)
            print(f"  {len(images)} slide(s) -> {scenario_dir}")
            print(f"  deck: {saved}")
        except deck_render.RenderUnavailable as exc:
            print(f"  FAIL -- PowerPoint could not open this deck: {exc}")
            return 1

    print(f"\n-- checks --")
    passed = checker(built)

    print("\n" + "=" * 78)
    if passed:
        print(f"PASS  {name} -- images at {scenario_dir}")
    else:
        print(f"FAIL  {name} -- see failures above. Images at {scenario_dir}")
    if saved is not None and not args.keep:
        try:
            saved.unlink()
        except OSError:
            pass
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
