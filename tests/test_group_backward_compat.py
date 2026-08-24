"""Scenario 5 of targeting_groups_test_scenarios.md -- backward
compatibility, "the one that matters most": generate a proposal, log it,
build it back two independent ways (Rebuild as presented, and Load into
form + Generate untouched), and assert the two decks are byte-identical.

    python tests/test_group_backward_compat.py
    python tests/test_group_backward_compat.py --keep    # leave both .pptx on disk

No pre-targeting-groups historical proposal is available to this suite (that
would need a real, years-old Supabase row), so this automates the scenario
doc's own stated fallback: "generate the same proposal twice: once via
Rebuild as presented, then Load into form and generate without touching
anything. Those two should match." The proposal itself is one of the
targeting-groups scenario fixtures (Annapolis Cars, the smallest of the
three) rather than a synthetic one, so this exercises the real property the
whole feature depends on -- `groups_to_seed_rows(seed_rows_to_groups(...))`
round-tripping -- against a real, group-heavy proposal, not a trivial one.

**What gets normalized before comparing, and why -- read this if the
comparison ever needs re-deriving:** two builds of the identical proposal,
a few seconds apart, were compared part-by-part before writing this test's
assertions (not assumed). Every one of the ~770 parts in a generated deck
came back byte-IDENTICAL in content except for one thing: the ZIP
container's own per-entry `date_time` field, which python-pptx's underlying
`zipfile` writer stamps with the wall-clock save time on every save,
regardless of whether that part's content changed. That is the ONLY
normalization this test applies -- it zeroes each entry's `date_time`
before comparing, and separately asserts the two files have the exact same
member list in the exact same order. Nothing inside any part's XML/binary
content is touched, edited, or excluded: `docProps/core.xml` (the obvious
place a real timestamp could hide) was checked directly and came back
byte-identical between the two probe builds with no normalization at all --
this app never rewrites the master deck's own created/modified properties.
If a future change ever needs a second exclusion, that's a regression in
determinism worth its own investigation, not something to paper over here.
"""
import argparse
import copy
import io
import os
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import db                                      # noqa: E402
import package_check                           # noqa: E402
import group_scenario_fixtures as gsf          # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    """rehydrate_proposal_into_form only touches st.session_state -- a plain
    dict stands in, same device tests/test_draft_regression.py uses."""

    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def generate_and_log(state):
    """Drive the real form to a real Generate, capturing exactly the row
    db.log_proposal would have inserted -- the same spy_log shape
    tests/test_draft_regression.py's build_deck uses, since that's exactly
    what a later "Load into form" or "Rebuild as presented" reads back."""
    from streamlit.testing.v1 import AppTest

    captured = {}
    real_log = db.log_proposal

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        captured["row"] = {
            "id": "00000000-0000-0000-0000-000000000000",
            "client_name": client_name, "vertical": vertical, "market": market,
            "form_json": form_json,
            "deck_version_id": kwargs.get("deck_version_id"),
            "output_filename": kwargs.get("output_filename"),
            "parent_proposal_id": kwargs.get("parent_proposal_id"),
            "revision_label": kwargs.get("revision_label"),
            "logo_storage_path": kwargs.get("logo_storage_path"),
            "created_by": kwargs.get("created_by"),
        }
        return "00000000-0000-0000-0000-000000000000", None

    db.log_proposal = spy_log
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Scenario Suite"
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return None, at
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not generate:
            return None, at
        generate[0].click().run()
        if at.exception:
            return None, at
    finally:
        db.log_proposal = real_log
    return captured.get("row"), at


def build_via_rebuild(row):
    """Path A -- "Rebuild as presented", called directly (it's a pure
    function of the logged row, no button to click)."""
    real_logo = db.proposal_logo
    db.proposal_logo = lambda storage_path: storage_path
    try:
        buffer, filename, warnings = app.rebuild_proposal_deck(row)
    finally:
        db.proposal_logo = real_logo
    return buffer, filename, warnings


def build_via_load_and_generate(row):
    """Path B -- "Load into form", then Generate untouched, exactly as a
    seller would click through it."""
    real_st, real_logo = app.st, db.proposal_logo
    stub = _StubSt()
    app.st = stub
    db.proposal_logo = lambda storage_path: storage_path
    try:
        notes = app.rehydrate_proposal_into_form(
            row, parent_proposal_id=row["id"], revision_default="Revision 2")
    finally:
        app.st = real_st
        db.proposal_logo = real_logo

    state = dict(stub.session_state)
    # rehydrate writes widget keys directly (client_name, flight_start, ...)
    # -- exactly what a real "Load into form" leaves in session_state for
    # the next render to pick up, so injecting it wholesale is the same
    # mechanism build_annapolis's own scenario uses, applied to a rehydrated
    # state instead of a fresh one.
    from streamlit.testing.v1 import AppTest
    captured = {}
    real_personalize = assembly.personalize

    def spy(prs, fill_data):
        w = real_personalize(prs, fill_data)
        captured["prs"] = prs
        return w

    assembly.personalize = spy
    db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Scenario Suite"
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return None, at, notes
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not generate:
            return None, at, notes
        generate[0].click().run()
        if at.exception:
            return None, at, notes
    finally:
        assembly.personalize = real_personalize

    prs = captured.get("prs")
    if prs is None:
        return None, at, notes
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer, at, notes


def normalized_entries(data):
    """{name: content bytes} for a .pptx's zip, dropping the one axis of
    real non-determinism (see this module's docstring): each entry's own
    `date_time`, stamped with wall-clock save time regardless of content."""
    zf = zipfile.ZipFile(io.BytesIO(data))
    return {info.filename: zf.read(info.filename) for info in zf.infolist()}, zf.namelist()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--keep", action="store_true", help="leave both rendered .pptx on disk")
    args = parser.parse_args()

    print("=" * 78)
    print("SCENARIO 5  Backward compatibility -- byte-identical round trip")
    print("=" * 78)

    scn = gsf.build_annapolis()
    state = gsf.build_session_state(scn)

    print("\nGenerate the original proposal and log it")
    row, at = generate_and_log(state)
    check("Generate runs without raising", not (at.exception if at else True),
          at.exception[0].message[:400] if at and at.exception else "")
    if not row:
        check("the proposal was logged with a form_json payload", False, "log_proposal never captured a row")
        print(f"\n{len(failures)} FAILED: {failures}")
        return 1
    check("the proposal was logged with a form_json payload", bool(row.get("form_json")))
    check("targeting_groups made it into the logged form_json (additive key)",
          bool(row["form_json"].get("targeting_groups")), None)

    print("\nPath A: Rebuild as presented")
    buf_a, filename_a, warnings_a = build_via_rebuild(row)
    for w in warnings_a:
        print(f"    ....  warning: {w}")
    check("rebuild produced a deck", buf_a is not None, warnings_a)
    if buf_a is None:
        print(f"\n{len(failures)} FAILED: {failures}")
        return 1

    print("\nPath B: Load into form, then Generate without touching anything")
    buf_b, at_b, notes = build_via_load_and_generate(row)
    for n in notes:
        print(f"    ....  rehydration note: {n}")
    check("Generate (path B) runs without raising",
          not (at_b.exception if hasattr(at_b, "exception") else True),
          at_b.exception[0].message[:400] if hasattr(at_b, "exception") and at_b.exception else "")
    check("load-and-generate produced a deck", buf_b is not None)
    if buf_b is None:
        print(f"\n{len(failures)} FAILED: {failures}")
        return 1

    print("\nCompare the two decks")
    bytes_a, bytes_b = buf_a.getvalue(), buf_b.getvalue()
    entries_a, names_a = normalized_entries(bytes_a)
    entries_b, names_b = normalized_entries(bytes_b)
    check("both decks have the exact same parts, in the exact same order",
          names_a == names_b, (len(names_a), len(names_b)))
    mismatched = [n for n in names_a if n in entries_b and entries_a[n] != entries_b[n]]
    check("every part's content is byte-identical (ZIP entry date_time excluded -- see module docstring)",
          not mismatched, mismatched[:10])
    check("raw file size matches (sanity check on the normalization itself)",
          len(bytes_a) == len(bytes_b), (len(bytes_a), len(bytes_b)))

    out_dir = Path(os.environ.get("TEMP", ".")) / "proposal_builder_round_trip"
    out_dir.mkdir(parents=True, exist_ok=True)
    path_a, path_b = out_dir / "rebuild_as_presented.pptx", out_dir / "load_and_generate.pptx"
    path_a.write_bytes(bytes_a)
    path_b.write_bytes(bytes_b)
    problems_a = package_check.check_package(str(path_a))
    problems_b = package_check.check_package(str(path_b))
    check("Path A has no unresolved relationship parts", not problems_a, problems_a[:5])
    check("Path B has no unresolved relationship parts", not problems_b, problems_b[:5])
    if not args.keep:
        path_a.unlink(missing_ok=True)
        path_b.unlink(missing_ok=True)
    else:
        print(f"\n    kept: {path_a}")
        print(f"    kept: {path_b}")

    print("\nA deliberately-blank field survives verbatim, not just a filled-in one")
    # The Timing bug's exact shape, on a different field: `or "CTV Strategy"`
    # substitutes the synthesized default for a REAL empty title too, not
    # only a genuinely-absent one. Checked directly against a copy of the
    # already-logged row rather than a hand-built one, since a synthetic
    # row is missing enough required `selections` keys that rebuild raises
    # before ever reaching this field.
    blank_row = copy.deepcopy(row)
    blank_row["form_json"]["proposal_title"] = ""
    captured_blank = {}
    real_personalize = assembly.personalize

    def spy_blank(prs, fill_data):
        captured_blank["title"] = fill_data["proposal_title"]
        return real_personalize(prs, fill_data)

    assembly.personalize = spy_blank
    try:
        app.rebuild_proposal_deck(blank_row)
    finally:
        assembly.personalize = real_personalize
    check("rebuild keeps a deliberately-blank proposal_title blank, not 'CTV Strategy'",
          captured_blank.get("title") == "", captured_blank.get("title"))

    real_st, real_logo = app.st, db.proposal_logo
    stub = _StubSt()
    app.st = stub
    db.proposal_logo = lambda storage_path: storage_path
    try:
        app.rehydrate_proposal_into_form(blank_row, parent_proposal_id=blank_row["id"])
    finally:
        app.st = real_st
        db.proposal_logo = real_logo
    check("Load into form keeps the same blank title blank",
          stub.session_state.get("proposal_title") == "", stub.session_state.get("proposal_title"))

    print("\nA proposal logged before include_in_plan existed rehydrates from its own rows")
    # Simulate a genuinely pre-feature proposal: strip include_in_plan/
    # include_locked off every group in a copy of the already-logged row,
    # the same way a real years-old proposal would arrive. Every one of
    # Annapolis's 8 groups is on this proposal's own plan (the fixture
    # includes all of them), so every group's id appears in some row's
    # _group_ids -- rehydration has to derive True for all 8 from that,
    # never fall back to the new-group default of False.
    legacy_row = copy.deepcopy(row)
    for g in legacy_row["form_json"]["targeting_groups"]:
        g.pop("include_in_plan", None)
        g.pop("include_locked", None)
    stub2 = _StubSt()
    app.st = stub2
    db.proposal_logo = lambda storage_path: storage_path
    try:
        app.rehydrate_proposal_into_form(legacy_row, parent_proposal_id=legacy_row["id"])
    finally:
        app.st = real_st
        db.proposal_logo = real_logo
    restored_groups = stub2.session_state.get("targeting_groups") or []
    check("every group came back WITH an include_in_plan key",
          all("include_in_plan" in g for g in restored_groups), restored_groups)
    check("every group derived True (all 8 were really on this proposal's plan)",
          all(g.get("include_in_plan") for g in restored_groups), restored_groups)
    check("every derived group is locked, so nothing later re-derives it",
          all(g.get("include_locked") for g in restored_groups), restored_groups)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Rebuild as presented and Load-into-form-then-Generate produce byte-identical decks "
          "(after excluding only the ZIP container's own save-time stamp) -- the round-trip "
          "property holds where it counts, not just at the data layer.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
