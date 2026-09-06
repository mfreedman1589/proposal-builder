"""The Build form survives page navigation, and "New proposal" really clears.

    python tests/test_form_state.py

Two halves, deliberately at different levels.

**An AST guard.** Streamlit refuses to have the value of a button, download
button, file uploader or form set through `st.session_state` -- it raises when
the widget is created, so a snapshot carrying one of those keys doesn't
degrade, it takes the page down. That happened during development: adding
`key="logo_upload"` to the client-logo uploader so the rest of the machinery
could see it put the key straight into the snapshot, and coming back from the
Audience finder raised. Nothing about the persistence code changed to fix it;
one prefix was added to a list. A list you have to remember to extend is the
failure this project has documented over and over, so the list is checked
against the source instead of against anyone's memory.

**A behavioural half, through AppTest.** Fill the form the way a seller
would -- including a drafted two-option plan, hand-edited (dirty) rows and an
imported Wide Orbit schedule -- walk to every other page and back, and assert
the whole thing is still there. Then press New proposal and assert the form is
indistinguishable from first load.

Needs PROPOSAL_BUILDER_TEST_MODE=1 for the schedule fixture injection, which
this sets for itself.
"""

import ast
import copy
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"
os.chdir(REPO)

# Before importing app: the schedule fixture is injected through the same
# test-mode path the Wide Orbit suite uses.
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import app                                       # noqa: E402

PASSED = []
FAILED = []
SKIPPED = []


def check(label, ok, actual=None, expected=None):
    if ok:
        PASSED.append(label)
        print(f"    PASS  {label}")
    else:
        FAILED.append(label)
        print(f"    FAIL  {label}")
        if expected is not None:
            print(f"          expected: {expected!r}")
        if actual is not None:
            print(f"          actual:   {actual!r}")
    return ok


def skip(label, why):
    SKIPPED.append(label)
    print(f"    SKIP  {label} -- {why}")


def section(title):
    print(f"\n  {title}")
    print(f"  {'-' * len(title)}")


# --------------------------------------------------------------------------
# 1. the AST guard
# --------------------------------------------------------------------------

# Widgets whose value Streamlit will not let session_state set. Kept as the
# names they're called by, since that is what appears in the source.
UNSETTABLE_WIDGETS = {"button", "download_button", "file_uploader",
                      "form_submit_button", "form"}


def _representative_key(node):
    """A key this `key=` argument could produce at runtime, or None.

    A plain string is itself. An f-string contributes its constant head and
    tail with a stand-in for the interpolated middle -- `f"finder_add_{seg}"`
    becomes "finder_add_<x>" and `f"{row_id}_save"` becomes "<x>_save" -- so
    both the prefix rule and the suffix rule can be asked about a real key
    rather than reimplemented here. Building the question and letting the app
    answer it is what stops this guard drifting from the thing it guards.
    """
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        head = tail = ""
        first, last = node.values[0], node.values[-1]
        if isinstance(first, ast.Constant) and isinstance(first.value, str):
            head = first.value
        if last is not first and isinstance(last, ast.Constant) and isinstance(last.value, str):
            tail = last.value
        if head or tail:
            return f"{head}<x>{tail}"
    return None


def keyed_unsettable_widgets(path):
    """[(lineno, widget, key_prefix)] for every keyed button/uploader in a file."""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    found = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if not isinstance(func, ast.Attribute) or func.attr not in UNSETTABLE_WIDGETS:
            continue
        for kw in node.keywords:
            if kw.arg != "key":
                continue
            found.append((node.lineno, func.attr, _representative_key(kw.value)))
    return found


def check_ast_guard():
    section("Every keyed button/uploader is excluded from the snapshot")
    widgets = keyed_unsettable_widgets(REPO / "app.py")
    check(f"found keyed unsettable widgets to check ({len(widgets)})", bool(widgets),
          len(widgets))

    uncovered = []
    for lineno, widget, key in widgets:
        # A non-literal key can't be checked statically, and that would be a
        # hole in this guard rather than an inconvenience.
        bad = key is None or app._persistable(key)
        if bad:
            uncovered.append((lineno, widget, key or "<computed>"))
        print(f"    ....  app.py:{lineno} st.{widget}(key={(key or '<computed>')!r}) -> "
              f"{'NOT EXCLUDED' if bad else 'excluded'}")
    check("no keyed button or uploader would be restored into session_state",
          not uncovered, uncovered)

    # The guard has to be able to fail, or it proves nothing.
    check("a key not on the exclusion list is reported as persistable",
          app._persistable("client_name"), app._persistable("client_name"))
    check("and one on it is not",
          not app._persistable("logo_upload"), app._persistable("logo_upload"))

    section("Editors are excluded too, and for a different reason")
    # A data_editor's stored value is an edit delta, not data. Replaying an
    # `added_rows` delta over rows that already contain the addition is how a
    # grid silently grows a duplicate line.
    for key in ("media_plan_editor_0_1", "avails_editor_0_monthly"):
        check(f"{key!r} is not snapshotted", not app._persistable(key))
    for key in ("plan_options", "avails_seed_rows"):
        check(f"but {key!r} -- which rebuilds them -- is kept", app._persistable(key))


# --------------------------------------------------------------------------
# 2. the behavioural half
# --------------------------------------------------------------------------

OTHER_PAGES = ["Audience finder", "Case study finder", "Proposal history",
               "Attribution reports",
               "Add case study", "Slide vault", "Add vault slide", "Update master deck"]

# One value per kind of state the form holds, so a regression names itself.
WATCHED = [
    "client_name", "vertical_choice", "market_choice", "agency_involved",
    "preset", "tegna_positioning", "include_vertical_slides",
    "goals_text", "audience_text", "geography_text", "budget_text",
    "placements_text", "timing_text",
    "flight_start", "flight_end", "active_months", "flight_months", "custom_flighting",
    "premion_streaming_tv", "total_tv", "dynamic_creative",
    "live_sports_enabled", "selected_sports", "include_sport_viewership",
    "sales_attribution", "brand_lift", "commercial_production",
    "avails_basis", "proposal_title", "show_cpm_column", "revision_label",
    # Derived bookkeeping, watched for the same reason as the fields: a stale
    # seed key is the documented way a form silently zeroes its own rows, and
    # it has to come back from navigation AND go away on a clear.
    "_product_seed_key", "_shared_fields_key", "_seeded_tactics",
]


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=240)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Form State Suite"
    return at


def snapshot(at, keys):
    return {k: (at.session_state[k] if k in at.session_state else "<GONE>") for k in keys}


def plan_shape(at):
    """What the media plan holds, reduced to what a seller would notice."""
    if "plan_options" not in at.session_state:
        return None
    return [
        {"name": o["name"], "breakout": o["breakout"], "dirty": list(o["dirty"]),
         "rows": [{"Tactic": r["Tactic"], "Impressions": r["Impressions"],
                   "Cost": r["Cost"], "CPM": r["CPM"]} for r in o["rows"]]}
        for o in at.session_state["plan_options"]
    ]


def load_drafted_form(at):
    """Put a real drafted two-option proposal into the form.

    The recorded HVAC draft, through the real apply_draft_to_form -- the same
    fixture tier 1 uses -- so this exercises the state a draft actually
    produces rather than a hand-built imitation.

    FLOW_REWORK_PLAN.md Phase 5: market/flight/basis are band INPUTS now, not
    drafted outputs -- apply_draft_to_form reads them from session_state
    rather than writing them, and the Build page's setup gate hard-`return`s
    below the band until flight_start/flight_end are set. So this preseeds
    the band with the fixture's own market/flight (standing in for "the rep
    had already filled in the band before hitting Draft," which is the only
    order the real app allows) BEFORE calling apply_draft_to_form -- without
    it, the gate fires on every subsequent run and nothing below the band,
    including Generate's own snapshot_form_state(), ever executes.
    """
    draft = json.loads((FIXTURES / "hvac_two_option.draft.json").read_text(encoding="utf-8"))
    draft.pop("_comment", None)

    class _Stub:
        def __init__(self):
            self.session_state = {}
            self.secrets = {}

    real, stub = app.st, _Stub()
    app.st = stub
    try:
        stub.session_state["market_choice"] = draft.get("market", "DC")
        stub.session_state["flight_start"] = date.fromisoformat(draft["flight_start"])
        stub.session_state["flight_end"] = date.fromisoformat(draft["flight_end"])
        app.apply_draft_to_form(copy.deepcopy(draft))
        # FLOW_REWORK_PLAN.md Phase 6: client_name is seeded through
        # `_draft_pending_fields`, drained by `apply_pending_draft_fields` on
        # the NEXT real run, before the band's first widget -- the real app
        # never gives anything a chance to observe the queue un-applied,
        # since apply_draft_to_form's only caller reruns immediately. This
        # stub calls apply_draft_to_form directly, outside that cycle, so it
        # has to drain the queue itself or callers see an unset client_name
        # and a queue that silently overwrites whatever they set afterward
        # on the next real at.run().
        app.apply_pending_draft_fields()
    finally:
        app.st = real
    for key, value in stub.session_state.items():
        at.session_state[key] = value
    return stub.session_state


def check_navigation():
    section("A filled form survives a walk to every other page and back")

    at = new_app()
    at.run()
    if at.exception:
        check("the form renders", False, [e.value for e in at.exception])
        return
    first_load = snapshot(at, WATCHED)

    # A drafted proposal: two options, sports, a flat fee, negotiated CPMs.
    drafted = load_drafted_form(at)
    at.session_state["client_name"] = "Ridgeline Heating & Air"
    at.session_state["revision_label"] = "Revision 2"
    at.session_state["include_sport_viewership"] = True
    at.session_state["uploaded_logo"] = {"name": "ridgeline.png", "bytes": b"\x89PNG-test"}
    # Total TV is what puts the broadcast panel on the page at all, and
    # switching a product on after a draft is itself the sequence that fires
    # the product-diff branch -- worth having in the same run.
    at.session_state["total_tv"] = True
    at.run()

    # The Wide Orbit upload is injected on its OWN separate run, after the
    # fields above are already stable (round-tripped through their own
    # widgets at least once). Injecting it in the SAME run as the fields
    # above is an AppTest-only hazard, not a real one: the intake area's
    # upload-triggered st.rerun() fires before Section A ever renders in
    # that pass, and a value AppTest injected via the Session State API
    # for a widget that hasn't rendered even once yet can be reset to its
    # widget default when that happens -- confirmed by direct repro (a
    # value already stable from a prior render survives a LATER upload-
    # triggered rerun fine; only a same-run injection racing one is
    # affected). A real rep's client_name is never "freshly assigned via
    # raw external state" at the exact instant an upload reruns the page --
    # it's already sitting in a rendered widget, or set in-script by
    # apply_draft_to_form followed immediately by its OWN st.rerun(), both
    # of which persist correctly regardless of ordering.
    schedule_fixture = FIXTURES / "wideorbit" / "ravens_campaign_schedule.xlsx"
    if schedule_fixture.exists():
        at.session_state["wo_upload_path"] = str(schedule_fixture.resolve())
    at.run()

    check("the drafted form renders without error", not at.exception,
          [e.value for e in at.exception])
    check("the draft produced two plan options",
          len(at.session_state["plan_options"]) == 2,
          len(at.session_state["plan_options"]))

    # Hand-edit a row, which is what marks it dirty and what a re-seed would
    # otherwise wipe -- the state most worth proving survives.
    options = at.session_state["plan_options"]
    options[0]["rows"][0]["Cost"] = 4242.0
    options[0]["dirty"][0] = True
    at.session_state["plan_options"] = options
    at.run()

    schedule_in = at.session_state["broadcast_schedule"] if "broadcast_schedule" in at.session_state else None
    if schedule_fixture.exists():
        check("a Wide Orbit schedule was imported", schedule_in is not None)
    else:
        skip("a Wide Orbit schedule was imported", "fixture not present (gitignored)")

    filled = snapshot(at, WATCHED)
    plan_before = plan_shape(at)
    avails_before = at.session_state["avails_seed_rows"]
    review_before = at.session_state["draft_unresolved"]
    check("the form really was filled (or this proves nothing)",
          filled != first_load, None)
    check("the hand edit is on the plan", plan_before[0]["rows"][0]["Cost"] == 4242.0,
          plan_before[0]["rows"][0]["Cost"])

    for page in OTHER_PAGES:
        at.session_state["page_choice"] = page
        at.run()
        at.session_state["page_choice"] = "Build a proposal"
        at.run()

        errors = [e.value for e in at.exception]
        if not check(f"via {page}: no exception", not errors, errors):
            continue
        after = snapshot(at, WATCHED)
        differing = {k: (filled[k], after[k]) for k in WATCHED if filled[k] != after[k]}
        check(f"via {page}: every watched field survives", not differing, differing)
        check(f"via {page}: the media plan survives intact",
              plan_shape(at) == plan_before, plan_shape(at))
        check(f"via {page}: the avails table survives",
              at.session_state["avails_seed_rows"] == avails_before,
              at.session_state["avails_seed_rows"])
        check(f"via {page}: the drafted review list survives",
              at.session_state["draft_unresolved"] == review_before, None)
        check(f"via {page}: the uploaded logo survives",
              ("uploaded_logo" in at.session_state
               and at.session_state["uploaded_logo"]["name"] == "ridgeline.png"), None)
        if schedule_in is not None:
            check(f"via {page}: the Wide Orbit schedule survives",
                  ("broadcast_schedule" in at.session_state
                   and at.session_state["broadcast_schedule"] is not None), None)

    return at, first_load


def check_new_proposal(at, first_load):
    section("\"New proposal\" returns the form to first-load state")

    # The confirm step: one press arms it, and nothing is cleared yet.
    at.session_state["confirm_new_proposal"] = True
    at.run()
    check("arming the confirm does not clear anything",
          at.session_state["client_name"] == "Ridgeline Heating & Air",
          at.session_state["client_name"] if "client_name" in at.session_state else "<GONE>")

    app_state = at.session_state
    clear_labels = [b.label for b in at.button]
    check("a confirm button is offered", "Clear the form" in clear_labels, clear_labels)
    check("so is a cancel", "Cancel" in clear_labels, clear_labels)

    # Cancel leaves everything alone.
    next(b for b in at.button if b.label == "Cancel").click().run()
    check("cancelling keeps the form", app_state["client_name"] == "Ridgeline Heating & Air",
          app_state["client_name"] if "client_name" in app_state else "<GONE>")

    at.session_state["confirm_new_proposal"] = True
    at.run()
    next(b for b in at.button if b.label == "Clear the form").click().run()

    check("clearing raises nothing", not at.exception, [e.value for e in at.exception])
    after = snapshot(at, WATCHED)
    differing = {k: (first_load[k], after[k]) for k in WATCHED if first_load[k] != after[k]}
    check("every field is back to its first-load value", not differing, differing)

    # The state a "clear" is most likely to miss: drafted, dirty, derived or
    # uploaded. Note the distinction that matters -- some of these are gone
    # outright, and some are *re-seeded* by the fresh render that follows, in
    # which case "cleared" means "back to what first load produces", not
    # "absent". Asserting absence for a re-seeded key would fail on a correct
    # clear, and asserting absence for a stale one would pass on a broken one.
    gone_outright = ("ai_filled_sections", "draft_unresolved", "draft_unresolved_internal",
                     "draft_round", "draft_last_json", "draft_source_notes",
                     "uploaded_logo", "broadcast_schedule", "restored_logo_path",
                     "confirm_new_proposal")
    for key in gone_outright:
        value = at.session_state[key] if key in at.session_state else None
        check(f"{key!r} is gone", not value, value)

    # FLOW_REWORK_PLAN.md Phase 1: flight_start/flight_end were never on
    # SESSION_KEEP_ON_RESET -- clearing has always wiped them, same as
    # every other field. What changed is the CONSEQUENCE: before Phase 1 a
    # cleared flight_start fell back to a real DEFAULT_FLIGHT_START on the
    # very next render, so plan_options/avails_seed_rows re-seeded
    # immediately; now a cleared flight closes the setup band's own gate
    # (nothing below it renders, per Phase 1's own "a fresh session shows
    # only the band" acceptance criterion), so they're not re-seeded at
    # all until the band is filled in again -- they're simply gone, the
    # same as every other gated key, not reset to a blank default.
    check("plan_options is gone outright -- the gate is closed, nothing re-seeds it yet",
          "plan_options" not in at.session_state,
          at.session_state["plan_options"] if "plan_options" in at.session_state else "<GONE>")
    check("avails_seed_rows is gone outright, same reason",
          "avails_seed_rows" not in at.session_state,
          at.session_state["avails_seed_rows"] if "avails_seed_rows" in at.session_state else "<GONE>")
    # snapshot_form_state() runs at the foot of main(), after the setup
    # band's own gate check -- with the gate closed post-clear, it never
    # runs at all this pass, so the backup key is gone outright too, same
    # reason as plan_options/avails_seed_rows above, not "rebuilt empty".
    backup = (at.session_state[app.FORM_STATE_BACKUP]
              if app.FORM_STATE_BACKUP in at.session_state else None)
    check("the snapshot carries no stale drafted-review-list data (gone outright, "
          "or rebuilt empty once the band is refilled below)",
          not (backup or {}).get("draft_unresolved"), backup)

    check("the signed-in user is kept", at.session_state["current_user"] == "Form State Suite",
          at.session_state["current_user"] if "current_user" in at.session_state else "<GONE>")

    print("\nRefilling the band after a clear reaches the SAME fresh-blank state a "
          "genuinely new session would -- New Proposal isn't a dead end")
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.run()
    check("no exception once the band is refilled", not at.exception, [e.value for e in at.exception])
    reseeded = at.session_state["plan_options"]
    check("the media plan is back to one blank option",
          len(reseeded) == 1 and len(reseeded[0]["rows"]) == 1
          and float(reseeded[0]["rows"][0]["Cost"] or 0) == 0, reseeded)
    check("no row is left marked dirty", not any(any(o["dirty"]) for o in reseeded),
          [o["dirty"] for o in reseeded])
    check("the avails table is back to one blank row",
          len(at.session_state["avails_seed_rows"]) == 1
          and not at.session_state["avails_seed_rows"][0]["Audience"],
          at.session_state["avails_seed_rows"])


def main():
    print("=" * 78)
    print("Form state across page navigation")
    print("=" * 78)

    check_ast_guard()
    result = check_navigation()
    if result:
        at, first_load = result
        check_new_proposal(at, first_load)

    print("\n" + "=" * 78)
    print(f"{len(PASSED)} passed, {len(FAILED)} failed, {len(SKIPPED)} skipped")
    for f in FAILED:
        print(f"  FAILED  {f}")
    for s in SKIPPED:
        print(f"  SKIPPED {s}")
    print("=" * 78)
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
