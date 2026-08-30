"""Render the standard scenarios to PNGs so they can be looked at.

    python tests/render_scenarios.py            # all of them
    python tests/render_scenarios.py ravens     # one, by name prefix

Not a pass/fail suite -- it produces pictures. Every layout failure this
project has shipped was invisible to an assertion and obvious in a
rendering: a Comscore footer read as a programme row, a table whose declared
height cleared the floor while the rendered one didn't, per-page impressions
that contradicted their own spot counts. The assertions catch what someone
thought to measure; this is the step where you look.

Images land in deck_render.render_root()/<scenario>/ -- a local, non-synced
directory, printed on every run -- alongside the .pptx that
produced them, and are gitignored.

Needs Windows with PowerPoint (see deck_render.py); prints SKIP otherwise.
"""

import copy
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app                                       # noqa: E402
import assembly                                  # noqa: E402
import db                                        # noqa: E402
import deck_render                               # noqa: E402
import group_scenario_fixtures as gsf            # noqa: E402
import package_check                             # noqa: E402
import wideorbit                                 # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
OUT = deck_render.render_root()


def _drafted_state(fixture):
    draft = json.loads((FIXTURES / fixture).read_text(encoding="utf-8"))
    draft.pop("_comment", None)

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
    """Drive the real form to a real deck, returning the Presentation."""
    from streamlit.testing.v1 import AppTest

    captured = {}
    real_personalize, real_log = assembly.personalize, db.log_proposal

    def spy(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        captured["warnings"] = warnings
        return warnings

    assembly.personalize = spy
    db.log_proposal = lambda *a, **k: ("render", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=900)
        for key, value in {**(state or {}), **(extra or {})}.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            raise RuntimeError(at.exception)
        generate = [b for b in at.button if b.label == "Generate proposal"][0]
        generate.click().run()
        if at.exception:
            raise RuntimeError(at.exception)
    finally:
        assembly.personalize, db.log_proposal = real_personalize, real_log
    return captured


def scenario_ashford():
    """Drafted streaming plan + Total TV + a real Wide Orbit schedule."""
    state = _drafted_state("ashford_post_draft.draft.json")
    schedule = FIXTURES / "wideorbit" / "regency_planner.xls"
    if not schedule.exists():
        return None
    return _generate(state, {
        "total_tv": True,
        "broadcast_schedule": wideorbit.parse_schedule(str(schedule), schedule.name),
        "broadcast_plan_desc": "148x Commercials, Morning/Daytime ROS, 4 Weeks",
    })


def scenario_ridgeline():
    """Two options, sports, a flat fee, a negotiated CPM, case studies."""
    return _generate(_drafted_state("hvac_two_option.draft.json"))


def _schedule_only(fixture, breakout, detailed, market="DC"):
    schedule_path = FIXTURES / "wideorbit" / fixture
    if not schedule_path.exists():
        return None
    schedule = wideorbit.parse_schedule(str(schedule_path), fixture)
    master, _, _ = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    selections = copy.deepcopy(assembly.SELECTIONS)
    selections["preset"] = "standard"
    selections["market"] = market
    selections["products"] = dict(selections["products"])
    selections["products"]["total_tv"] = True
    selections["broadcast_schedule_imported"] = True
    prs, _, _ = assembly.build_presentation(master, selections)
    warnings = assembly.build_broadcast_schedule_slides(
        prs, schedule, "NFL Regular Season, Local Games", breakout, detailed)
    return {"prs": prs, "warnings": warnings}


def scenario_case_studies():
    """A standard Total TV deck with two placeholder-based case studies.

    Grafted straight from the local files rather than picked through the
    vault, so this needs no Supabase and always uses the same two decks.
    They're the ones that exposed the inherited-formatting bug: everything
    they don't state explicitly -- placeholder sizes, theme colours, theme
    fonts -- used to be reinterpreted by Premion's own master, which turned
    brand-white headings near-black. 18 of the vault's 22 case studies are
    built this way, so this is the common case rather than an awkward one.
    """
    source = REPO / "case_studies_source"
    names = ["PREMION_Case Study_Regional Furniture + Mattress Retailer.pptx",
             "PREMION_Case Study_Fine Jewelry Retailer.pptx"]
    if not all((source / name).exists() for name in names):
        return None
    master_path, _, _ = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    if master_path is None:
        return None
    import copy as _copy
    selections = _copy.deepcopy(assembly.SELECTIONS)
    selections["preset"], selections["market"] = "standard", "DC"
    selections["products"] = dict(selections["products"])
    selections["products"]["total_tv"] = True
    prs, _, _ = assembly.build_presentation(master_path, selections)
    assembly.append_case_studies(
        prs, [{"path": str(source / name), "slides": None} for name in names])
    return {"prs": prs, "warnings": []}


def build_st_louis_three_group():
    """Three targeting groups over one metro (St. Louis) with real overlap
    -- not a real client document (unlike annapolis/hershey/wilmington
    below), built directly from real zips so it needs no gitignored
    fixture and always renders. Exists to look at the targeting map's
    place-label crowding and county-fill balance at actual slide size: two
    of the three groups share a county (so the coverage-threshold /
    highest-coverage-wins fill logic in targeting_map._touched_counties
    has something real to arbitrate), and the St. Charles County cluster
    packs several named places (O'Fallon, St. Charles, St. Peters,
    Wentzville) close enough together to exercise label-overlap
    suppression rather than a roomy, uncrowded map.
    """
    all_months, flight_label = gsf._flight(date(2027, 1, 1), date(2027, 3, 31))
    rows_spec = [
        ("AUTO Intenders", [
            "63101", "63102", "63103", "63104", "63105", "63106", "63107", "63108",
            "63109", "63110", "63111", "63112", "63113", "63114", "63115", "63116",
            "63117", "63118", "63119", "63120", "63121", "63122", "63123", "63124",
            "63125", "63126", "63127", "63128", "63129", "63130", "63131", "63132",
            "63133", "63134", "63135", "63136", "63137", "63138", "63139", "63140",
        ], 1700000),
        ("Home Improvement Intenders", [
            "63017", "63021", "63025", "63026", "63038", "63040", "63043", "63044",
            "63045", "63146",
        ], 180000),
        ("Higher Education Intenders", [
            "63301", "63303", "63304", "63366", "63376", "63385",
        ], 240000),
    ]

    groups, rows = [], []
    for i, (term, zips, avails) in enumerate(rows_spec):
        geo_def, resolved_zips, markets, notes, unresolved = gsf._zips_geo(", ".join(zips))
        group = gsf._make_group([term], None, geo_def, resolved_zips, markets, avails, i)
        group["_flight_label"] = flight_label
        groups.append(group)
        rows.append(gsf._plan_row(group, "premion_streaming_tv", avails))

    return {
        "name": "St. Louis three-group", "client_name": "Plaza Motors Group",
        "flight_start": date(2027, 1, 1), "flight_end": date(2027, 3, 31),
        "active_months": all_months, "flight_label": flight_label,
        "n_months": len(all_months), "agency_gross_up": False,
        "vertical_choice": "Automotive",
        "groups": groups, "rows": rows,
    }


def _group_scenario(builder):
    """A targeting-groups scenario (tests/group_scenario_fixtures.py) driven
    through the real form -- same mechanism as every other scenario here,
    the groups/avails/plan-rows state just comes from a real avails
    document's own fixture instead of a drafted-notes fixture. See that
    module's docstring for why the groups themselves are built as data
    (already proven correct elsewhere) while geography is resolved through
    the real `app.resolve_group_geography`, and why this needs no
    `_drafted_state`/apply_draft_to_form step at all."""
    return _generate(gsf.build_session_state(builder()))


SCENARIOS = {
    "ashford_total_tv": scenario_ashford,
    "ridgeline": scenario_ridgeline,
    "case_studies": scenario_case_studies,
    "ravens_full_flight": lambda: _schedule_only(
        "ravens_campaign_schedule.xlsx", "full_flight", True),
    "regency": lambda: _schedule_only("regency_planner.xls", "full_flight", True),
    # targeting_groups_test_scenarios.md's scenarios 2-4, against the real
    # avails PDFs at the repo root (gitignored) -- see
    # tests/group_scenario_fixtures.py and tests/test_group_scenarios.py,
    # which asserts row count/ordering/geography/avails-totals/gross-markup
    # against these same documents. This is the "look at it" half; that
    # file is the "assert on it" half.
    "annapolis_cars": lambda: _group_scenario(gsf.build_annapolis),
    "visit_hershey": lambda: _group_scenario(gsf.build_hershey),
    "wilmington_university": lambda: _group_scenario(gsf.build_wilmington),
    # Not a real client document -- see build_st_louis_three_group's own
    # docstring. Three groups, real overlap, so the targeting map's place
    # labels and county fills have real crowding to show, not an
    # artificially roomy two-group map.
    "st_louis_three_group": lambda: _group_scenario(build_st_louis_three_group),
}


def main(argv):
    if not deck_render.renderer_available():
        print("SKIP -- rendering needs Windows with PowerPoint and pywin32.")
        return 0
    only = argv[1] if len(argv) > 1 else None
    names = [n for n in SCENARIOS if not only or n.startswith(only)]
    if not names:
        print(f"No scenario matches {only!r}. Known: {', '.join(SCENARIOS)}")
        return 2

    OUT.mkdir(parents=True, exist_ok=True)
    failures = []
    for name in names:
        print(f"\n=== {name} ===")
        try:
            built = SCENARIOS[name]()
        except Exception as exc:                                 # noqa: BLE001
            print(f"  FAILED to build: {type(exc).__name__}: {exc}")
            continue
        if built is None:
            print("  SKIP -- fixture not present")
            continue
        for warning in built.get("warnings") or []:
            print(f"  warning: {warning}")
        try:
            images, saved = deck_render.render_presentation(
                built["prs"], OUT / name, name=name)
        except deck_render.RenderUnavailable as exc:
            # A deck PowerPoint won't open is a FAILURE, not a skip. This is
            # the only check that catches it: python-pptx reads such a file
            # happily and every structural assertion passes.
            print(f"  FAIL -- PowerPoint could not open this deck: {exc}")
            failures.append(name)
            continue
        problems = package_check.check_package(str(saved))
        if problems:
            print(f"  FAIL -- {len(problems)} unresolved relationship(s)")
            for problem in problems[:5]:
                print(f"      {problem}")
            failures.append(name)
        print(f"  {len(images)} slide(s) -> {OUT / name}")
        print(f"  deck: {saved}")
    # Printed in full, every run: the output is deliberately outside the repo
    # (see deck_render.render_root), so "it's in tests/_rendered" would send
    # you to a stale directory.
    print(f"\nImages are under {OUT} -- open them and look.")
    if failures:
        print(f"\n{len(failures)} scenario(s) produced a deck that won't open or has "
              f"unresolved parts: {', '.join(failures)}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
