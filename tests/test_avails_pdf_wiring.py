"""The avails PDF importer wired into the real form -- both entry points,
through AppTest + test_mode_upload injection (the same technique
test_wideorbit-adjacent tests use for a file_uploader, which AppTest can't
drive directly). Uses the real PDFs at the repo root; SKIPs without them.

Three things asserted that the pure-module test (test_avails_pdf_import.py)
can't reach on its own, because they only exist once app.py is involved:
    1. The D2 uploader creates real, group-backed avails rows that seed real
       plan lines -- the "reaches the same state as entering the document by
       hand" claim, carried one level up from targeting_groups into the
       actual grid and media plan.
    2. Precedence: a rep-set client_name is left alone (flagged as a
       conflict), while an untouched one gets filled in from the document.
    3. The intake-area entry point (always visible now, not gated on the
       notes mentioning "avail" -- see the UX sweep, BACKLOG.md) re-uses the
       identical import path (same rfpid tracked, so uploading through
       EITHER entry point trips the same already-imported guard).

    python tests/test_avails_pdf_wiring.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402

ANNAPOLIS = REPO / "Premion Media Plan_RFPID-253813_SR&B Advertising_Annapolis Cars_1-23-2026--ver0.pdf"
LAWN_LEISURE = REPO / "Premion Media Plan_RFPID-265521_Direct - No Agency_Lawn & Leisure_7-28-2026--ver0.pdf"
HERSHEY = REPO / "Premion Media Plan_RFPID-260402_Direct - No Agency_Visit Hershey & Harrisburg_4-30-2026--ver0.pdf"

failures = []
skipped = False


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app(open_gate=True):
    """`open_gate=True` (the default, and what every D2 scenario in this
    file needs) pre-seeds the setup band's flight so D2 actually renders --
    FLOW_REWORK_PLAN.md Phase 1 gated it behind market+flight, so a truly
    untouched fresh form can no longer reach it at all. The one scenario
    that deliberately wants a fresh, flight-less form (the intake-area
    "flight lands from the document" case, below) passes `open_gate=False`
    and uses the intake entry point instead, which stayed ungated on
    purpose -- see FLOW_REWORK_PLAN.md's own "the gate must never block an
    input" principle."""
    from datetime import date
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["target_dmas"] = ["Washington, DC"]
    if open_gate:
        at.session_state["market_choice"] = "DC"
        at.session_state["flight_start"] = date(2026, 9, 1)
        at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def real_groups(at):
    groups = at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []
    return [g for g in groups if g["terms"]]


def ss(at, key, default=None):
    """AppTest's session_state has no `.get` -- this is the safe substitute
    used throughout this file instead of repeating the `in` check."""
    return at.session_state[key] if key in at.session_state else default


def main():
    global skipped
    if not (ANNAPOLIS.exists() and LAWN_LEISURE.exists() and HERSHEY.exists()):
        print("SKIP -- real avails PDFs not present (gitignored fixtures)")
        skipped = True
        return 0

    print("D2 uploader: a fresh form (client_name still at its default) picks up the document")
    at = new_app()
    at.session_state["avails_pdf_upload_path_d2"] = str(ANNAPOLIS)
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    groups = real_groups(at)
    check("8 groups created, one per audience-radius pair", len(groups) == 8, len(groups))
    check("client_name filled in from the document (was still at its default)",
          at.session_state["client_name"] == "Annapolis Cars", ss(at, "client_name"))
    # FLOW_REWORK_PLAN.md Phase 4b: the import never touches the gross-up
    # checkbox -- it's a rep-only, order-wide calculator now, not something
    # any importer or the drafting model can set. A document naming a real
    # agency is surfaced as a note instead, and the checkbox is untouched.
    check("agency_gross_up is untouched by the import -- not a widget any importer sets",
          ss(at, "agency_gross_up") in (None, False), ss(at, "agency_gross_up"))
    report = ss(at, "avails_import_report")
    check("the document's agency is surfaced as a note, not auto-applied to any checkbox",
          report and any("agency" in n.lower() and "gross-up" in n.lower()
                        for n in report["unresolved_internal"]),
          report["unresolved_internal"] if report else None)
    check("the report's parsed total matches the document's own header total",
          report and report["parsed_total"] == report["total_impressions"], report)
    check("avails_monthly was populated (non-zero) on every imported group",
          all(g["avails_monthly"] > 0 for g in groups), [g["avails_monthly"] for g in groups])

    print("\nprecedence: a rep-set client_name is left alone, flagged as a conflict, not overwritten")
    at2 = new_app()
    at2.session_state["client_name"] = "A Client The Rep Already Typed"
    at2.session_state["avails_pdf_upload_path_d2"] = str(ANNAPOLIS)
    at2.run()
    check("no exception", not at2.exception, at2.exception[0].message[:400] if at2.exception else "")
    check("the rep's own client_name survives the import untouched",
          at2.session_state["client_name"] == "A Client The Rep Already Typed",
          ss(at2, "client_name"))
    report2 = ss(at2, "avails_import_report")
    check("the conflict is reported, not silently dropped",
          report2 and any("client name" in c.lower() for c in report2["conflicts"]),
          report2["conflicts"] if report2 else None)
    check("the groups themselves still get created regardless -- a header-field conflict "
          "doesn't block the geography/audience data the document owns outright",
          len(real_groups(at2)) == 8, len(real_groups(at2)))

    print("\na document saying 'Direct - No Agency' (Lawn & Leisure) produces no agency note at "
          "all -- the checkbox is still never touched, and there's nothing to flag")
    at2b = new_app()
    at2b.session_state["avails_pdf_upload_path_d2"] = str(LAWN_LEISURE)  # Direct - No Agency
    at2b.run()
    check("no exception", not at2b.exception, at2b.exception[0].message[:400] if at2b.exception else "")
    check("agency_gross_up is still untouched",
          ss(at2b, "agency_gross_up") in (None, False), ss(at2b, "agency_gross_up"))
    report2b = ss(at2b, "avails_import_report")
    check("no agency note fires for a document that names no agency",
          report2b and not any("gross-up" in n.lower() for n in report2b["unresolved_internal"]),
          report2b["unresolved_internal"] if report2b else None)

    print("\nre-importing an already-recorded RFPID is refused, not silently duplicated")
    # A fresh AppTest instance with the RFPID pre-seeded into
    # avails_import_history -- simulating "this document was already
    # imported earlier in the session" without needing to force a second
    # real file_uploader event, which AppTest can't drive directly. (Tried
    # deleting the "already loaded this name" marker and re-running on the
    # SAME instance instead; restore_form_state() -- correctly, by design --
    # restores an absent key from its own snapshot, which silently undid
    # that deletion. That's the app protecting real session continuity
    # working as intended, not a bug this test should route around by
    # fighting it.)
    at3 = new_app()
    at3.session_state["avails_import_history"] = ["RFPID-253813"]
    at3.session_state["avails_pdf_upload_path_d2"] = str(ANNAPOLIS)
    at3.run()
    check("no exception on the upload", not at3.exception,
          at3.exception[0].message[:400] if at3.exception else "")
    check("no groups were added -- the duplicate RFPID was refused",
          len(real_groups(at3)) == 0, len(real_groups(at3)))
    check("history still has exactly the one, pre-seeded entry",
          ss(at3, "avails_import_history") == ["RFPID-253813"], ss(at3, "avails_import_history"))
    check("the refusal is a visible error, not a silent no-op",
          "already" in (ss(at3, "avails_import_error") or "").lower(),
          ss(at3, "avails_import_error"))

    print("\nintake-area entry point: always available (not gated on the notes mentioning "
          "avails -- see the UX sweep, BACKLOG.md), ungated by the Phase 1 setup band, and "
          "shares the same import path")
    at4 = new_app(open_gate=False)
    at4.session_state["draft_source_notes"] = "Client wants a CTV campaign, budget TBD."
    at4.session_state["avails_pdf_upload_path_intake"] = str(LAWN_LEISURE)
    at4.run()
    check("no exception", not at4.exception, at4.exception[0].message[:400] if at4.exception else "")
    check("the intake uploader picked up the upload and created a group -- "
          "notes saying nothing about avails doesn't gate it",
          len(real_groups(at4)) == 1, len(real_groups(at4)))
    check("flight dates filled in from Lawn & Leisure's own flight (a fresh form, still at defaults)",
          str(ss(at4, "flight_start")) == "2026-09-06", ss(at4, "flight_start"))

    print("\nVisit Hershey & Harrisburg: importing a 12-group document populates the avails "
          "table only -- zero media-plan lines until a rep ticks Plan")
    # Two separate .run() calls, matching what a real session does: open the
    # form (premion_streaming_tv on, plan_options already seeded from that),
    # THEN scroll to D2 and upload -- not the same initial session_state a
    # single-render test would use, which never has a pre-existing
    # plan_options to reconcile against.
    at6 = new_app()
    at6.session_state["premion_streaming_tv"] = True
    at6.run()
    at6.session_state["avails_pdf_upload_path_d2"] = str(HERSHEY)
    at6.run()
    check("no exception", not at6.exception, at6.exception[0].message[:400] if at6.exception else "")
    groups6 = real_groups(at6)
    check("12 groups created (2 audiences x (4 markets + 2 zip add-ons))", len(groups6) == 12, len(groups6))
    check("none of them ticked into the plan -- avails-table-only import",
          not any(g["include_in_plan"] for g in groups6), groups6)

    rows6 = ss(at6, "plan_options")[0]["rows"]
    group_ids6 = {g["id"] for g in groups6}
    check("zero Premion Streaming TV lines before anything is ticked",
          not [r for r in rows6 if set(app.group_ids_of(r)) & group_ids6], rows6)

    # "Add all to plan" (the D2 bulk control) is the fastest way to exercise
    # all 12 lines appearing -- same queue-then-apply mechanism a single
    # tick uses, just for every group at once.
    at6.session_state["_pending_group_include_all"] = True
    at6.run()
    check("no exception after Add all", not at6.exception,
          at6.exception[0].message[:400] if at6.exception else "")
    groups6b = real_groups(at6)
    check("all 12 groups now ticked", all(g["include_in_plan"] for g in groups6b), groups6b)
    rows6b = ss(at6, "plan_options")[0]["rows"]
    hershey_rows = [r for r in rows6b if set(app.group_ids_of(r)) & group_ids6]
    check("one Premion Streaming TV line per group -- 12, not 1",
          len(hershey_rows) == 12, len(hershey_rows))
    check("every one of the 12 is the Premion Streaming TV tactic",
          all(r.get("Tactic") == "Premion Streaming TV" for r in hershey_rows),
          [r.get("Tactic") for r in hershey_rows])
    # Audience-major order (plan_lines_from_groups' own ordering rule): all
    # six of audience A's rows before all six of audience B's.
    targeting_sequence = [r.get("Targeting") for r in hershey_rows]
    check("document order is audience-major -- A's six rows, then B's six",
          targeting_sequence == [targeting_sequence[0]] * 6 + [targeting_sequence[6]] * 6
          and targeting_sequence[0] != targeting_sequence[6],
          targeting_sequence)

    print("\nTargeting is never empty on a seeded Hershey line (asserted separately from the "
          "count above -- the count bug and the empty-Targeting symptom are the same root cause, "
          "but worth confirming independently in case a future fix addresses only one)")
    blank = [r for r in hershey_rows if not str(r.get("Targeting") or "").strip()]
    check("no Hershey line has empty/None Targeting", not blank, blank)
    check("every Hershey line's Targeting is one of the document's two audiences",
          all(r.get("Targeting") in {targeting_sequence[0], targeting_sequence[6]} for r in hershey_rows),
          set(targeting_sequence))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Both entry points run the same importer, apply the same precedence, and a real upload "
          "reaches the real form's targeting_groups and header fields exactly the way typing the "
          "same document in by hand would -- including seeding a media-plan line for every group, "
          "not just the first.")
    return 0


if __name__ == "__main__":
    code = main()
    if skipped:
        sys.exit(0)
    sys.exit(code)
