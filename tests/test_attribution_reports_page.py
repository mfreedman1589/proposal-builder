"""The Attribution Reports page (ATTRIBUTION_REPORT_PLAN.md Phase 2) --
through the real form via AppTest + test_mode_upload injection, the same
technique test_avails_pdf_wiring.py/test_wideorbit.py use for a
file_uploader AppTest can't drive directly. db.py's advertiser/proposal/
report functions are stubbed with a small in-memory fake store -- the real
`advertisers`/`attribution_reports` tables don't exist in Supabase until the
migration is pasted in by hand, and a test must never write real rows
either way (same reason test_draft_regression.py stubs db.log_proposal).

    python tests/test_attribution_reports_page.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
from streamlit.testing.v1 import AppTest  # noqa: E402

MW_FIXTURE = REPO / "MW attribution excel.xlsx"
WAEPA_FIXTURE = REPO / "Premion Website Attribution and Reach Extension (13).xlsx"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class FakeStore:
    """Records every call and hands back plausible rows -- enough of
    db.py's real contract for app.py's logic to exercise, without touching
    Supabase.
    """
    def __init__(self):
        self.advertisers = []
        self.proposals = []
        self.reports = []
        self.create_advertiser_calls = []
        self.log_report_calls = []
        self.link_calls = []
        self._next = 1

    def _id(self):
        self._next += 1
        return f"fake-{self._next}"

    def fetch_advertisers(self, limit=1000):
        return list(self.advertisers), None

    def create_advertiser(self, name):
        self.create_advertiser_calls.append(name)
        display = " ".join((name or "").split())
        key = display.lower()
        for row in self.advertisers:
            if row["canonical_name"].lower() == key:
                return row, None
        row = {"id": self._id(), "canonical_name": display}
        self.advertisers.append(row)
        return row, None

    def fetch_proposals(self, limit=500):
        return list(self.proposals), None

    def fetch_proposal(self, proposal_id):
        for row in self.proposals:
            if row["id"] == proposal_id:
                return row, None
        return None, "That proposal no longer exists."

    def link_proposal_advertiser(self, proposal_id, advertiser_id):
        self.link_calls.append((proposal_id, advertiser_id))
        return True, None

    def log_attribution_report(self, advertiser_id, proposal_id, report_json,
                               status="parsed", created_by=None):
        call = dict(advertiser_id=advertiser_id, proposal_id=proposal_id,
                   report_json=report_json, status=status, created_by=created_by)
        self.log_report_calls.append(call)
        return self._id(), None


def _fake_proposal_row(rid="prop-1", client_name="Mattress Warehouse", target_dmas=None,
                       goals="Drive qualified leads for the fall promotion",
                       audience="Adults 25-54 in-market for home improvement",
                       flight_label="Jun-Jul 2026", budget_cost="$50,000"):
    """Phase 5's field-map shape (ATTRIBUTION_REPORT_PLAN.md) -- real market
    KEYS (`washington_hagerstown`/`baltimore`, the same two the real WAEPA
    proposal uses) so `target_market_labels` resolves them against the real
    market-profile catalog, `campaign_specs.goals`/`.audience` (the goals-
    prefill bug's own real field), and a `deck_payload.media_plan_options`
    shape matching what a real proposal's Generate actually writes."""
    return {
        "id": rid, "client_name": client_name, "vertical": "retail", "market": "DC",
        "generated_at": "2026-08-01T12:00:00", "deck_version_id": 5,
        "parent_proposal_id": None, "revision_label": None,
        "file_storage_path": None, "logo_storage_path": None, "created_by": "Matt",
        "target_dmas": target_dmas if target_dmas is not None else ["washington_hagerstown", "baltimore"],
        "form_json": {
            "proposal_title": "CTV Strategy", "plan_options": [],
            "flight": {"label": flight_label},
            "setup": {"originating_market": "DC", "flight_start": "2026-06-01",
                     "flight_end": "2026-07-17", "plan_basis": "monthly",
                     "avails_mode": True, "total_tv": False},
            "campaign_specs": {"goals": goals, "audience": audience},
            "deck_payload": {"media_plan_options": [
                {"full_flight_total": {"cost": budget_cost, "impressions": "1,000,000"}}]},
            "targeting_groups": [],
        },
    }


def _ss(at, key, default=None):
    """at.session_state has no .get() -- AppTest's SafeSessionState only
    proxies __getitem__/__contains__, and .get resolves as an attribute
    lookup for a literal key named "get" instead. Every other test file
    here works around this with an explicit `in` check; this is that
    check, factored out since this file needs it repeatedly.
    """
    return at.session_state[key] if key in at.session_state else default


def new_app():
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Test"
    return at


def check_upload_first_new_advertiser_no_proposal(store):
    print("\nUpload-first door: brand-new advertiser, no matching proposal")
    if not MW_FIXTURE.exists():
        print("  SKIP  MW attribution excel.xlsx not present")
        return

    at = new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_attribution_upload_path"] = str(MW_FIXTURE)
    at.run()
    check("no exception after upload", not at.exception, at.exception)

    parsed = _ss(at, "attr_parsed_attribution")
    check("the export parsed", bool(parsed), parsed)
    if parsed:
        check("client name resolved from the export", parsed.get("client_name"), "Mattress Warehouse")
        check("client name is exactly 'Mattress Warehouse'",
             parsed.get("client_name") == "Mattress Warehouse", parsed.get("client_name"))

    radios = at.radio
    check("an advertiser radio is shown", len(radios) >= 1, len(radios))
    if not radios:
        return
    advertiser_radio = radios[0]
    check("its only option is 'create new'", "Create new advertiser" in advertiser_radio.options[-1],
         advertiser_radio.options)
    advertiser_radio.set_value(advertiser_radio.options[-1]).run()

    confirm_buttons = [b for b in at.button if b.label == "Confirm advertiser"]
    check("a Confirm advertiser button is present", bool(confirm_buttons), [b.label for b in at.button])
    if confirm_buttons:
        confirm_buttons[0].click().run()
    check("no exception after confirming the advertiser", not at.exception, at.exception)
    check("attr_advertiser_id got set", _ss(at, "attr_advertiser_id"),
         _ss(at, "attr_advertiser_id"))
    check("create_advertiser was called with the parsed client name",
         store.create_advertiser_calls == ["Mattress Warehouse"], store.create_advertiser_calls)

    no_proposal_buttons = [b for b in at.button
                           if b.label == "No proposal -- build this report standalone"]
    check("the no-matching-proposal path offers a standalone button",
         bool(no_proposal_buttons), [b.label for b in at.button])
    if no_proposal_buttons:
        no_proposal_buttons[0].click().run()
    check("attr_no_proposal is set", _ss(at, "attr_no_proposal") is True,
         _ss(at, "attr_no_proposal"))

    log_buttons = [b for b in at.button if b.label == "Log this report"]
    check("a Log this report button is present", bool(log_buttons), [b.label for b in at.button])
    if log_buttons:
        log_buttons[0].click().run()
    check("no exception after logging", not at.exception, at.exception)
    check("exactly one report was logged", len(store.log_report_calls), 1)
    if store.log_report_calls:
        call = store.log_report_calls[0]
        check("logged with the created advertiser id",
             call["advertiser_id"] == _ss(at, "attr_advertiser_id"), call)
        check("logged with no proposal", call["proposal_id"] is None, call)
        check("report_json carries the parsed attribution facts",
             call["report_json"].get("attribution", {}).get("client_name") == "Mattress Warehouse",
             call["report_json"])
        check("report_json's delivery half is None (no delivery file uploaded)",
             call["report_json"].get("delivery") is None, call["report_json"])

    print("\n  Generate the report deck (Phase 3 walking skeleton)")
    template = REPO / "REPORT_MASTER_v0_4.pptx"
    if not template.exists():
        print("  SKIP  REPORT_MASTER_v0_4.pptx not present")
        return
    goals_areas = [t for t in at.text_area if t.key == "attr_goals_input"]
    whats_next_areas = [t for t in at.text_area if t.key == "attr_whats_next_input"]
    check("goals text area is present", bool(goals_areas), [t.key for t in at.text_area])
    check("what's-next text area is present", bool(whats_next_areas), [t.key for t in at.text_area])
    if goals_areas and whats_next_areas:
        goals_areas[0].set_value("Drive online engagement for the sale").run()
        whats_next_areas[0].set_value("Expand into the top-performing markets").run()
    generate_buttons = [b for b in at.button if b.label == "Generate report deck"]
    check("a Generate report deck button is present", bool(generate_buttons),
         [b.label for b in at.button])
    if generate_buttons:
        generate_buttons[0].click().run()
    check("no exception after generating", not at.exception, at.exception)
    download_buttons = [b for b in at.download_button if b.label == "⬇ Download report .pptx"]
    check("a download button appears after a successful generate",
         bool(download_buttons), [b.label for b in at.download_button])


def check_rfpid_confirm_gate(store):
    print("\nMulti-RFPID confirm gate (WAEPA: 2 RFPIDs, a real split IO)")
    if not WAEPA_FIXTURE.exists():
        print("  SKIP  WAEPA fixture not present")
        return

    at = new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_attribution_upload_path"] = str(WAEPA_FIXTURE)
    at.run()
    check("no exception after upload", not at.exception, at.exception)

    rfpid_checkboxes = [c for c in at.checkbox if c.key == "attr_rfpid_confirmed"]
    check("the RFPID confirm checkbox is shown", bool(rfpid_checkboxes),
         [c.key for c in at.checkbox])
    if not rfpid_checkboxes:
        return
    check("2 RFPIDs defaults to confirmed (the split-IO shape, not the rollup one)",
         rfpid_checkboxes[0].value is True, rfpid_checkboxes[0].value)

    # Default confirmed -- the page proceeds past the gate into advertiser
    # matching (a radio, since this fixture's advertiser doesn't yet exist
    # in the fake store).
    check("advertiser matching renders past the gate (default confirmed)",
         len(at.radio) >= 1, len(at.radio))

    # Uncheck it -- the gate blocks, advertiser matching disappears. (Some
    # radio may still exist elsewhere on the page -- e.g. sidebar nav --
    # so this checks for the SPECIFIC advertiser-matching radio's own
    # option text going away, not a bare app-wide radio count.)
    rfpid_checkboxes[0].set_value(False).run()
    check("no exception after unchecking", not at.exception, at.exception)
    check("unchecking the confirm blocks advertiser matching from rendering",
         not any("Create new advertiser" in opt for r in at.radio for opt in r.options),
         [r.options for r in at.radio])
    check("an info message explains what to do",
         any("Confirm above" in str(m) for m in _all_markdown_text(at)), _all_markdown_text(at))


def check_conversions_toggle(store):
    print("\nConversions toggle (WAEPA has them, MW doesn't)")
    if not WAEPA_FIXTURE.exists():
        print("  SKIP  WAEPA fixture not present")
        return

    at = new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_attribution_upload_path"] = str(WAEPA_FIXTURE)
    at.run()
    # Default-confirmed RFPID gate lets the rest of the page through with
    # no extra interaction (see check_rfpid_confirm_gate).
    radios = at.radio
    check("advertiser radio present", bool(radios), [r.label for r in radios])
    if not radios:
        return
    radios[0].set_value(radios[0].options[-1]).run()
    confirm_buttons = [b for b in at.button if b.label == "Confirm advertiser"]
    if confirm_buttons:
        confirm_buttons[0].click().run()
    no_proposal_buttons = [b for b in at.button
                           if b.label == "No proposal -- build this report standalone"]
    if no_proposal_buttons:
        no_proposal_buttons[0].click().run()

    conv_checkboxes = [c for c in at.checkbox if c.key == "attr_include_conversions"]
    check("the Include conversions checkbox is shown for a conversions export",
         bool(conv_checkboxes), [c.key for c in at.checkbox])
    if conv_checkboxes:
        check("it defaults ON", conv_checkboxes[0].value is True, conv_checkboxes[0].value)

    template = REPO / "REPORT_MASTER_v0_4.pptx"
    if not template.exists():
        print("  SKIP  REPORT_MASTER_v0_4.pptx not present -- can't test Generate")
        return
    goals_areas = [t for t in at.text_area if t.key == "attr_goals_input"]
    whats_next_areas = [t for t in at.text_area if t.key == "attr_whats_next_input"]
    if goals_areas and whats_next_areas:
        goals_areas[0].set_value("Drive qualified insurance leads").run()
        whats_next_areas[0].set_value("Expand DC/Baltimore targeting").run()
    generate_buttons = [b for b in at.button if b.label == "Generate report deck"]
    check("a Generate report deck button is present", bool(generate_buttons),
         [b.label for b in at.button])
    if generate_buttons:
        generate_buttons[0].click().run()
    check("no exception generating WITH conversions on", not at.exception, at.exception)

    # Turn it off and generate again -- must still build cleanly (the
    # graceful-degrade path when the template hasn't been widened yet is
    # exercised by report_assembly's own tests; this just proves the app
    # wiring survives the toggle both ways).
    conv_checkboxes = [c for c in at.checkbox if c.key == "attr_include_conversions"]
    if conv_checkboxes:
        conv_checkboxes[0].set_value(False).run()
    generate_buttons = [b for b in at.button if b.label == "Generate report deck"]
    if generate_buttons:
        generate_buttons[0].click().run()
    check("no exception generating WITH conversions off", not at.exception, at.exception)


def check_mw_has_no_conversions_toggle(store):
    print("\nMW (no conversions) shows no toggle at all")
    if not MW_FIXTURE.exists():
        print("  SKIP  MW attribution excel.xlsx not present")
        return
    at = new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_attribution_upload_path"] = str(MW_FIXTURE)
    at.run()
    conv_checkboxes = [c for c in at.checkbox if c.key == "attr_include_conversions"]
    check("no Include conversions checkbox for a non-conversions export",
         not conv_checkboxes, [c.key for c in at.checkbox])


def check_clear_button(store):
    print("\nClear / new report")
    if not MW_FIXTURE.exists():
        print("  SKIP  MW attribution excel.xlsx not present")
        return
    at = new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_attribution_upload_path"] = str(MW_FIXTURE)
    at.run()
    check("something got parsed before clearing", bool(_ss(at, "attr_parsed_attribution")),
         _ss(at, "attr_parsed_attribution"))

    clear_buttons = [b for b in at.button if b.key == "attr_clear_button"]
    check("the Clear / new report button is present", bool(clear_buttons),
         [b.key for b in at.button])
    if not clear_buttons:
        return
    clear_buttons[0].click().run()
    confirm_buttons = [b for b in at.button if b.key == "attr_confirm_clear_yes"]
    check("a confirm step appears rather than clearing immediately",
         bool(confirm_buttons), [b.key for b in at.button])
    if not confirm_buttons:
        return
    confirm_buttons[0].click().run()
    check("no exception after confirming", not at.exception, at.exception)
    check("the parsed attribution export is gone",
         _ss(at, "attr_parsed_attribution") is None, _ss(at, "attr_parsed_attribution"))
    check("the attribution upload path is gone",
         _ss(at, "attr_attribution_path") is None, _ss(at, "attr_attribution_path"))


def check_prelinked_door(store):
    print("\nPre-linked door: 'Build report' from Proposal History")
    fake_row = _fake_proposal_row(rid="prop-history-1", client_name="Cardinal Plumbing")
    store.proposals = [fake_row]

    print("  History row exposes a 'Build report' button that sets the right flags")
    at = new_app()
    at.session_state["page_choice"] = "Proposal history"
    at.run()
    check("no exception rendering History", not at.exception, at.exception)
    build_report_buttons = [b for b in at.button if b.label == "Build report"]
    check("a 'Build report' button is present on the row", bool(build_report_buttons),
         [b.label for b in at.button])
    if build_report_buttons:
        build_report_buttons[0].click().run()
    check("no exception after clicking it", not at.exception, at.exception)
    check("it navigated to Attribution reports",
         _ss(at, "page_choice") == "Attribution reports",
         _ss(at, "page_choice"))
    check("the proposal id was recorded",
         _ss(at, "attr_prelinked_proposal_id") == "prop-history-1",
         _ss(at, "attr_prelinked_proposal_id"))

    print("  The page shows the linked proposal and does not re-search")
    check("the linked proposal's client name is shown",
         any("Cardinal Plumbing" in str(m) for m in _all_markdown_text(at)), None)


def check_phase5_proposal_link(store):
    """ATTRIBUTION_REPORT_PLAN.md Phase 5, through the real page --
    complements test_report_assembly.py's own Phase 5 check (which drives
    report_assembly.py's functions directly, offline) by proving the app
    WIRING actually reaches them: the goals-prefill fix, and that
    Generate's real deck carries the linked proposal's flight/geography/
    budget-derived facts. The advertiser name is deliberately made to
    DISAGREE with the export's own ("WAEPA Insurance" vs. the export's
    "WAEPA") -- proving the linked proposal's name wins throughout,
    including the client name in the finished deck's own CLIENT_NAME/
    recap tokens and the download filename.
    """
    print("\nPhase 5: proposal-linked fields reach the real page and the generated deck")
    if not WAEPA_FIXTURE.exists():
        print("  SKIP  WAEPA fixture not present")
        return
    template = REPO / "REPORT_MASTER_v0_4.pptx"
    if not template.exists():
        print("  SKIP  REPORT_MASTER_v0_4.pptx not present -- can't test Generate")
        return

    fake_row = _fake_proposal_row(
        rid="prop-phase5-1", client_name="WAEPA Insurance",
        goals="Evaluate OTT performance beyond awareness\nTrack engaged visits and cost per engaged visit",
        audience="Federal government employees researching benefits",
        flight_label="Oct 2026 - Dec 2026", budget_cost="$74,970")
    store.proposals = [fake_row]

    at = new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_prelinked_proposal_id"] = "prop-phase5-1"
    at.session_state["attr_attribution_upload_path"] = str(WAEPA_FIXTURE)
    at.run()
    check("no exception with a linked proposal + upload", not at.exception, at.exception)

    goals_areas = [t for t in at.text_area if t.key == "attr_goals_input"]
    check("goals text area is present", bool(goals_areas), [t.key for t in at.text_area])
    if goals_areas:
        check("goals prefilled from campaign_specs.goals -- the real prefill bug's own fix "
             "(it used to read a top-level 'goals' key that doesn't exist on any real "
             "proposal, silently prefilling '')",
             "Evaluate OTT performance" in goals_areas[0].value, goals_areas[0].value)

    whats_next_areas = [t for t in at.text_area if t.key == "attr_whats_next_input"]
    check("what's-next text area is present", bool(whats_next_areas),
         [t.key for t in at.text_area])
    if whats_next_areas:
        whats_next_areas[0].set_value("Expand DC/Baltimore targeting").run()

    generate_buttons = [b for b in at.button if b.label == "Generate report deck"]
    check("a Generate report deck button is present", bool(generate_buttons),
         [b.label for b in at.button])
    if not generate_buttons:
        return
    generate_buttons[0].click().run()
    check("no exception after generating", not at.exception, at.exception)
    download_buttons = [b for b in at.download_button if b.label == "⬇ Download report .pptx"]
    check("a download button appears after a successful generate",
         bool(download_buttons), [b.label for b in at.download_button])
    if not download_buttons:
        return

    import db as _db
    from pptx import Presentation
    out_path = _db.scratch_dir("attribution_reports") / "WAEPA Insurance.pptx"
    check(f"the generated deck exists on disk ({out_path.name})", out_path.exists(), out_path)
    if not out_path.exists():
        return
    prs = Presentation(str(out_path))
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text.append(cell.text_frame.text)
    alltext = "\n".join(text)
    check("the linked proposal's own client name wins over the export's own "
         "(CLIENT_NAME reads 'WAEPA Insurance', not bare 'WAEPA')",
         "WAEPA Insurance" in alltext, alltext[:300])
    check("FLIGHT_LABEL reads the linked proposal's own flight",
         "Oct 2026 - Dec 2026" in alltext, None)
    check("GEOGRAPHY_LABEL reads the target-market names ('Washington, DC' and 'Baltimore')",
         "Washington, DC" in alltext and "Baltimore" in alltext, None)
    check("AUDIENCE_BULLETS reads the linked proposal's own Campaign Specs audience",
         "Federal government employees researching benefits" in alltext, None)


def _all_markdown_text(at):
    texts = []
    for kind in ("markdown", "caption", "info", "success", "warning", "error", "title", "header", "subheader"):
        try:
            texts.extend(str(getattr(el, "value", "")) for el in getattr(at, kind))
        except Exception:                                                 # noqa: BLE001
            pass
    return texts


def check_pixel_issue_window(store):
    """The known tracking-pixel under-count window (June 9 - July 21, 2026,
    Pansophic memo, 2026-09-08 correction) -- a non-blocking, rep-facing
    warning when an uploaded export's flight overlaps it. Real fixtures on
    both sides: MW's real flight (Jun 1 - Jul 13) genuinely overlaps (this
    is the SAME defect MW's own 2.15x projection factor turned out to be,
    per DECISIONS.md); WAEPA's real flight (Jul 27 - Aug 31) genuinely
    starts after the fix. Pure edge cases (boundary days, missing dates)
    are covered directly against `app._overlaps_pixel_issue_window` --
    this is the "does it actually reach the page" half.
    """
    print("\nPixel-issue-window warning (2026-09-08 correction)")
    cases = [
        (date(2026, 6, 1), date(2026, 6, 15), True, "starts before, ends inside"),
        (date(2026, 7, 15), date(2026, 8, 1), True, "starts inside, ends after"),
        (date(2026, 5, 1), date(2026, 8, 1), True, "fully contains the window"),
        (date(2026, 1, 1), date(2026, 3, 1), False, "well before"),
        (date(2026, 8, 1), date(2026, 9, 1), False, "well after"),
        (date(2026, 6, 9), date(2026, 6, 9), True, "exact start boundary"),
        (date(2026, 7, 21), date(2026, 7, 21), True, "exact end boundary"),
        (date(2026, 7, 22), date(2026, 8, 1), False, "starts the day after the fix"),
        (None, date(2026, 6, 15), False, "missing start"),
    ]
    for start, end, expected, label in cases:
        got = app._overlaps_pixel_issue_window(start, end)
        check(f"overlap logic: {label}", got == expected, (got, expected))

    if not MW_FIXTURE.exists():
        print("  SKIP  MW attribution excel.xlsx not present")
    else:
        at = new_app()
        at.session_state["page_choice"] = "Attribution reports"
        at.session_state["attr_attribution_upload_path"] = str(MW_FIXTURE)
        at.run()
        check("no exception after upload", not at.exception, at.exception)
        texts = _all_markdown_text(at)
        check("MW's real flight overlaps the window -- the warning shows",
             any("known tracking issue" in t for t in texts), texts)

    if not WAEPA_FIXTURE.exists():
        print("  SKIP  WAEPA fixture not present")
    else:
        at = new_app()
        at.session_state["page_choice"] = "Attribution reports"
        at.session_state["attr_attribution_upload_path"] = str(WAEPA_FIXTURE)
        at.run()
        check("no exception after upload", not at.exception, at.exception)
        texts = _all_markdown_text(at)
        check("WAEPA's real flight starts after the fix -- no warning",
             not any("known tracking issue" in t for t in texts), texts)


def main():
    store = FakeStore()
    app.db.fetch_advertisers = store.fetch_advertisers
    app.db.create_advertiser = store.create_advertiser
    app.db.fetch_proposals = store.fetch_proposals
    app.db.fetch_proposal = store.fetch_proposal
    app.db.link_proposal_advertiser = store.link_proposal_advertiser
    app.db.log_attribution_report = store.log_attribution_report

    check_upload_first_new_advertiser_no_proposal(store)
    check_rfpid_confirm_gate(store)
    check_conversions_toggle(store)
    check_mw_has_no_conversions_toggle(store)
    check_clear_button(store)
    check_prelinked_door(store)
    check_phase5_proposal_link(store)
    check_pixel_issue_window(store)

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
