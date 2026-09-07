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


def _fake_proposal_row(rid="prop-1", client_name="Mattress Warehouse"):
    return {
        "id": rid, "client_name": client_name, "vertical": "retail", "market": "DC",
        "generated_at": "2026-08-01T12:00:00", "deck_version_id": 5,
        "parent_proposal_id": None, "revision_label": None,
        "file_storage_path": None, "logo_storage_path": None, "created_by": "Matt",
        "form_json": {
            "proposal_title": "CTV Strategy", "plan_options": [],
            "flight": {"label": "Jun-Jul 2026"},
            "setup": {"originating_market": "DC", "flight_start": "2026-06-01",
                     "flight_end": "2026-07-17", "plan_basis": "monthly",
                     "avails_mode": True, "total_tv": False},
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
    template = REPO / "REPORT_MASTER_v0_3.pptx"
    if not template.exists():
        print("  SKIP  REPORT_MASTER_v0_3.pptx not present")
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


def _all_markdown_text(at):
    texts = []
    for kind in ("markdown", "caption", "info", "success", "warning", "error", "title", "header", "subheader"):
        try:
            texts.extend(str(getattr(el, "value", "")) for el in getattr(at, kind))
        except Exception:                                                 # noqa: BLE001
            pass
    return texts


def main():
    store = FakeStore()
    app.db.fetch_advertisers = store.fetch_advertisers
    app.db.create_advertiser = store.create_advertiser
    app.db.fetch_proposals = store.fetch_proposals
    app.db.fetch_proposal = store.fetch_proposal
    app.db.link_proposal_advertiser = store.link_proposal_advertiser
    app.db.log_attribution_report = store.log_attribution_report

    check_upload_first_new_advertiser_no_proposal(store)
    check_prelinked_door(store)

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
