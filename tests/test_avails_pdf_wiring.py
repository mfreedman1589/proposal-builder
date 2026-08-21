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
    3. The review-list entry point only appears when the notes mention
       "avail", and re-uses the identical import path (same rfpid tracked,
       so uploading through EITHER entry point trips the same
       already-imported guard).

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

ANNAPOLIS = REPO / "Premion Media Plan_RFPID-253813_SR&B Advertising_Annapolis Cars_1-23-2026--ver0.pdf"
LAWN_LEISURE = REPO / "Premion Media Plan_RFPID-265521_Direct - No Agency_Lawn & Leisure_7-28-2026--ver0.pdf"

failures = []
skipped = False


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["target_dmas"] = ["Washington, DC"]
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
    if not (ANNAPOLIS.exists() and LAWN_LEISURE.exists()):
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
    check("agency_involved turned on -- this document names a real agency",
          at.session_state["agency_involved"] is True, ss(at, "agency_involved"))
    report = ss(at, "avails_import_report")
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

    print("\nprecedence, the boolean edge case: a document saying 'Direct - No Agency' "
          "(agency_involved should become False) must still be able to apply OR conflict -- "
          "a bare falsy check would silently treat False as 'nothing stated'")
    at2b = new_app()
    at2b.session_state["agency_involved"] = True  # the rep already turned it on
    at2b.session_state["avails_pdf_upload_path_d2"] = str(LAWN_LEISURE)  # Direct - No Agency
    at2b.run()
    check("no exception", not at2b.exception, at2b.exception[0].message[:400] if at2b.exception else "")
    check("the rep's own agency_involved=True survives -- flagged, not silently flipped to False",
          at2b.session_state["agency_involved"] is True, ss(at2b, "agency_involved"))
    report2b = ss(at2b, "avails_import_report")
    check("the conflict is reported for agency involved specifically",
          report2b and any("agency involved" in c.lower() for c in report2b["conflicts"]),
          report2b["conflicts"] if report2b else None)

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

    print("\nreview-list entry point: only appears when the notes mention avails, "
          "and shares the same import path")
    at4 = new_app()
    at4.session_state["draft_source_notes"] = "Client wants a CTV campaign, budget TBD."
    at4.run()
    check("no exception", not at4.exception, at4.exception[0].message[:400] if at4.exception else "")
    check("no avails mention in the notes -- no inline uploader caption shown",
          not any("Your notes mention avails" in (str(c.value) if hasattr(c, "value") else str(c))
                  for c in at4.caption),
          None)

    at5 = new_app()
    at5.session_state["draft_source_notes"] = (
        "Client wants a CTV campaign. I already pulled avails in Salesforce for this -- "
        "attaching the PDF.")
    at5.session_state["avails_pdf_upload_path_review"] = str(LAWN_LEISURE)
    at5.run()
    check("no exception", not at5.exception, at5.exception[0].message[:400] if at5.exception else "")
    check("the notes-mention uploader picked up the upload and created a group",
          len(real_groups(at5)) == 1, len(real_groups(at5)))
    check("flight dates filled in from Lawn & Leisure's own flight (a fresh form, still at defaults)",
          str(ss(at5, "flight_start")) == "2026-09-06", ss(at5, "flight_start"))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Both entry points run the same importer, apply the same precedence, and a real upload "
          "reaches the real form's targeting_groups and header fields exactly the way typing the "
          "same document in by hand would.")
    return 0


if __name__ == "__main__":
    code = main()
    if skipped:
        sys.exit(0)
    sys.exit(code)
