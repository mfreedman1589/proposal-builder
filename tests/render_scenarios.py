"""Render the standard scenarios to PNGs so they can be looked at.

    python tests/render_scenarios.py            # all of them
    python tests/render_scenarios.py ravens     # one, by name prefix

Not a pass/fail suite -- it produces pictures. Every layout failure this
project has shipped was invisible to an assertion and obvious in a
rendering: a Comscore footer read as a programme row, a table whose declared
height cleared the floor while the rendered one didn't, per-page impressions
that contradicted their own spot counts. The assertions catch what someone
thought to measure; this is the step where you look.

Images land in tests/_rendered/<scenario>/, alongside the .pptx that
produced them, and are gitignored.

Needs Windows with PowerPoint (see deck_render.py); prints SKIP otherwise.
"""

import copy
import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app                                       # noqa: E402
import assembly                                  # noqa: E402
import db                                        # noqa: E402
import deck_render                               # noqa: E402
import package_check                             # noqa: E402
import wideorbit                                 # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
OUT = Path(__file__).resolve().parent / "_rendered"


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


SCENARIOS = {
    "ashford_total_tv": scenario_ashford,
    "ridgeline": scenario_ridgeline,
    "ravens_full_flight": lambda: _schedule_only(
        "ravens_campaign_schedule.xlsx", "full_flight", True),
    "regency": lambda: _schedule_only("regency_planner.xls", "full_flight", True),
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
        print(f"  {len(images)} slide(s) -> {(OUT / name).relative_to(REPO)}")
        print(f"  deck: {saved.relative_to(REPO)}")
    print(f"\nImages are under {OUT.relative_to(REPO)} -- open them and look.")
    if failures:
        print(f"\n{len(failures)} scenario(s) produced a deck that won't open or has "
              f"unresolved parts: {', '.join(failures)}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
