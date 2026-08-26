"""The intake area (UX sweep, BACKLOG.md) wired into the real form: notes
uploaded as a file (not just pasted), a Wide Orbit export uploaded inline in
the setup band once Total TV is on (FLOW_REWORK_PLAN.md Phase 1 moved this
uploader out of the always-visible intake area and gated it behind Total TV,
so "before Total TV is even on" is no longer reachable at all), and the
client logo from its intake-area location -- through AppTest +
test_mode_upload injection, the same technique test_avails_pdf_wiring.py and
test_wideorbit.py use for a file_uploader, which AppTest can't drive
directly.

Wide Orbit's own real fixture lives at tests/fixtures/wideorbit/ (gitignored,
same as test_wideorbit.py's); SKIPs cleanly without it rather than failing.

    python tests/test_intake_area.py
"""
import os
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402

WO_FIXTURE = REPO / "tests" / "fixtures" / "wideorbit" / "ravens_campaign_schedule.xlsx"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    from datetime import date
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates everything below it
    # (Section C's Wide Orbit summary panel included) on market+flight --
    # the intake area itself (notes, logo) stays ungated and doesn't need
    # this, but the Total TV/Wide Orbit scenarios below do.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def ss(at, key, default=None):
    return at.session_state[key] if key in at.session_state else default


def all_text(elements):
    """Flatten an AppTest element collection (caption, markdown, ...) to a
    single string for a simple substring check."""
    return " ".join(str(getattr(e, "value", e)) for e in elements)


def main():
    print("=" * 78)
    print("SCENARIO  meeting notes uploaded as a file populate the paste box")
    print("=" * 78)
    with tempfile.TemporaryDirectory() as tmp:
        notes_path = Path(tmp) / "discovery_call.txt"
        notes_path.write_text(
            "Client wants a CTV campaign. Budget is $50,000. Flight is September.",
            encoding="utf-8")

        at = new_app()
        at.session_state["notes_upload_path"] = str(notes_path)
        at.run()
        check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
        check("the uploaded file's text landed in the paste box",
              ss(at, "draft_notes_input", "") == notes_path.read_text(encoding="utf-8"),
              ss(at, "draft_notes_input"))
        check("no extraction error shown", not ss(at, "notes_text_error"), ss(at, "notes_text_error"))

        print("\nAN UNTOUCHED FOLLOW-UP RERUN must not re-clear or re-fill it -- "
              "the upload only seeds the box once, an edit afterward is never overwritten")
        before = ss(at, "draft_notes_input")
        at.run()
        check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
        check("the notes text survives byte-for-byte", ss(at, "draft_notes_input") == before,
              (before, ss(at, "draft_notes_input")))

        print("\nan unsupported upload extension is reported, not silently ignored")
        bad_path = Path(tmp) / "notes.xlsx"
        bad_path.write_bytes(b"not really an xlsx")
        at2 = new_app()
        at2.session_state["notes_upload_path"] = str(bad_path)
        at2.run()
        check("no exception", not at2.exception, at2.exception[0].message[:400] if at2.exception else "")
        check("an error is shown naming the file",
              bool(ss(at2, "notes_text_error")) and "notes.xlsx" in ss(at2, "notes_text_error"),
              ss(at2, "notes_text_error"))
        check("the paste box was NOT overwritten with garbage",
              not ss(at2, "draft_notes_input"), ss(at2, "draft_notes_input"))

    print("\n" + "=" * 78)
    print("SCENARIO  Wide Orbit uploads inline in the setup band, only once Total TV is on")
    print("=" * 78)
    # FLOW_REWORK_PLAN.md Phase 1 moved this uploader OUT of the always-
    # visible intake area and into the setup band itself, shown only when
    # Total TV is checked -- "upload before Total TV is on" is no longer a
    # reachable state at all (the uploader doesn't render, so a same-run
    # injection has nothing to be picked up by); this scenario now proves
    # the upload works once Total TV is on, and that a rep who checks Total
    # TV with nothing uploaded yet gets pointed at the band, not an empty
    # metrics panel.
    if not WO_FIXTURE.exists():
        print(f"  SKIP -- no fixture at {WO_FIXTURE.relative_to(REPO)}")
    else:
        print("without any upload, Total TV points back at the setup band instead of an "
              "empty metrics panel")
        at3 = new_app()
        at3.session_state["total_tv"] = True
        at3.run()
        check("no exception", not at3.exception, at3.exception[0].message[:400] if at3.exception else "")
        check("the pointer-back caption is shown, naming the Setup band",
              "No Wide Orbit schedule uploaded yet" in all_text(at3.caption)
              and "Setup" in all_text(at3.caption), all_text(at3.caption))
        check("no metrics panel -- nothing to show yet",
              not any(getattr(m, "label", None) == "Spots" for m in at3.metric), None)

        print("\nuploading through the band's own inline uploader (Total TV already on) "
              "parses immediately")
        at3.session_state["wo_upload_path"] = str(WO_FIXTURE)
        at3.run()
        check("no exception", not at3.exception, at3.exception[0].message[:400] if at3.exception else "")
        check("the schedule parsed and landed in session_state",
              ss(at3, "broadcast_schedule") is not None, ss(at3, "wo_error"))
        check("the band's own uploader confirms it loaded",
              "Loaded" in all_text(at3.caption), None)

        print("\na later rerun shows the ALREADY-parsed summary under Total TV, not an "
              "empty panel or a re-upload prompt")
        at3.run()
        check("no exception", not at3.exception, at3.exception[0].message[:400] if at3.exception else "")
        check("the metrics panel rendered (Spots is one of the st.metric labels)",
              any(getattr(m, "label", None) == "Spots" for m in at3.metric), None)
        check("no \"nothing uploaded yet\" pointer shown -- it's already there",
              "No Wide Orbit schedule uploaded yet" not in all_text(at3.caption), None)

        print("\nturning Total TV back OFF hides the band's uploader, but the already-"
              "parsed schedule is untouched (Total TV owns visibility, not the data)")
        at3.session_state["total_tv"] = False
        at3.run()
        check("no exception", not at3.exception, at3.exception[0].message[:400] if at3.exception else "")
        check("the schedule is still there", ss(at3, "broadcast_schedule") is not None, None)

    print("\n" + "=" * 78)
    print("SCENARIO  client logo, uploaded from its new intake-area location")
    print("=" * 78)
    with tempfile.TemporaryDirectory() as tmp:
        # A 1x1 PNG is enough -- this only proves the upload reaches
        # uploaded_logo from its NEW call site, not that it renders on a
        # slide (assembly.py's own tests cover that).
        logo_path = Path(tmp) / "logo.png"
        logo_path.write_bytes(
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
            b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0"
            b"\x00\x00\x03\x01\x01\x00\x18\xdd\x8d\xb0\x00\x00\x00\x00IEND\xaeB`\x82")
        at5 = new_app()
        at5.session_state["logo_upload_path"] = str(logo_path)
        at5.run()
        check("no exception", not at5.exception, at5.exception[0].message[:400] if at5.exception else "")
        uploaded = ss(at5, "uploaded_logo")
        check("the logo reached uploaded_logo from its new intake-area call site",
              uploaded is not None and uploaded.get("name") == "logo.png", uploaded)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The intake area's relocated pieces all reach the same state a rep typing or "
          "uploading directly into the old locations would have: notes-file upload seeds "
          "the paste box once and never overwrites an edit; the logo uploader works "
          "identically from its new home; and the setup band's own inline Wide Orbit "
          "uploader (Phase 1 -- gated behind Total TV, not the always-visible intake area "
          "any more) parses immediately and survives Total TV being toggled back off.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
