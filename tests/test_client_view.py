"""The Clients page and the Merge/rename clients admin page
(ATTRIBUTION_REPORT_PLAN.md Phase 6, items 3/4/7/8) -- through the real
pages via AppTest, db.py stubbed with a small in-memory fake store, same
convention test_attribution_reports_page.py already uses (never real rows).

    python tests/test_client_view.py
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

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def equal(label, actual, expected):
    check(label, actual == expected, f"expected {expected!r}, got {actual!r}")


class FakeStore:
    """Same shape as test_attribution_reports_page.py's own FakeStore, with
    the additions this page's read side needs (fetch_case_studies,
    update_case_study) and the merge/rename admin page's own writes."""

    def __init__(self):
        self.advertisers = []
        self.proposals = []
        self.reports = []
        self.case_studies = []
        self.set_vertical_calls = []
        self.rename_calls = []
        self.merge_calls = []
        self.tag_calls = []

    def fetch_advertisers(self, limit=1000, active_only=False):
        rows = [r for r in self.advertisers if (not active_only or r.get("active", True))]
        return list(rows), None

    def fetch_proposals(self, limit=500):
        return list(self.proposals), None

    def fetch_proposal(self, proposal_id):
        for row in self.proposals:
            if row["id"] == proposal_id:
                return row, None
        return None, "That proposal no longer exists."

    def fetch_attribution_reports(self, advertiser_id=None, limit=500):
        rows = [r for r in self.reports if not advertiser_id or r.get("advertiser_id") == advertiser_id]
        return list(rows), None

    def fetch_case_studies(self, active_only=True):
        rows = self.case_studies
        if active_only:
            rows = [r for r in rows if r.get("active", True)]
        return list(rows), None

    def update_case_study(self, case_study_id, **fields):
        self.tag_calls.append((case_study_id, fields))
        for row in self.case_studies:
            if row["id"] == case_study_id:
                row.update(fields)
                return row, None
        return None, "not found"

    def set_advertiser_vertical(self, advertiser_id, vertical):
        self.set_vertical_calls.append((advertiser_id, vertical))
        for row in self.advertisers:
            if row["id"] == advertiser_id:
                row["vertical"] = vertical
        return True, None

    def set_advertiser_name(self, advertiser_id, name):
        self.rename_calls.append((advertiser_id, name))
        display = " ".join((name or "").split())
        key = display.lower()
        for row in self.advertisers:
            if row["id"] != advertiser_id and (row.get("canonical_name") or "").lower() == key \
                    and row.get("active", True):
                return None, f"\"{display}\" already exists as another client -- merge them instead."
        for row in self.advertisers:
            if row["id"] == advertiser_id:
                row["canonical_name"] = display
                return row, None
        return None, "Advertiser not found."

    def advertiser_usage_counts(self, advertiser_id):
        return {
            "proposals": sum(1 for p in self.proposals if p.get("advertiser_id") == advertiser_id),
            "attribution_reports": sum(1 for r in self.reports if r.get("advertiser_id") == advertiser_id),
            "case_studies": sum(1 for c in self.case_studies if c.get("advertiser_id") == advertiser_id),
        }, None

    def merge_advertisers(self, source_id, target_id):
        self.merge_calls.append((source_id, target_id))
        for p in self.proposals:
            if p.get("advertiser_id") == source_id:
                p["advertiser_id"] = target_id
        for r in self.reports:
            if r.get("advertiser_id") == source_id:
                r["advertiser_id"] = target_id
        for c in self.case_studies:
            if c.get("advertiser_id") == source_id:
                c["advertiser_id"] = target_id
        for row in self.advertisers:
            if row["id"] == source_id:
                row["active"] = False
        return True, None


def _ss(at, key, default=None):
    return at.session_state[key] if key in at.session_state else default


def new_app():
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Test"
    return at


def _all_markdown_text(at):
    texts = []
    for kind in ("markdown", "caption", "info", "success", "warning", "error", "title", "header", "subheader"):
        try:
            texts.extend(str(getattr(el, "value", "")) for el in getattr(at, kind))
        except Exception:                                                 # noqa: BLE001
            pass
    return texts


def check_waepa_scoping(store):
    """The acceptance test named in the plan: the WAEPA advertiser shows
    ITS proposal, ITS reports, and nothing else's -- not Mattress
    Warehouse's, even though both are logged in the same tables."""
    print("\nClients page: one advertiser's data, nothing else's")
    waepa = {"id": "adv-waepa", "canonical_name": "WAEPA", "vertical": "banking", "active": True}
    mw = {"id": "adv-mw", "canonical_name": "Mattress Warehouse", "vertical": "retail", "active": True}
    store.advertisers = [waepa, mw]
    store.proposals = [
        {"id": "prop-waepa-1", "client_name": "WAEPA", "advertiser_id": "adv-waepa",
         "vertical": "banking", "market": "DC", "generated_at": "2026-09-01T00:00:00",
         "form_json": {"proposal_title": "WAEPA CTV", "plan_options": [],
                      "flight": {"label": "Oct-Dec 2026"},
                      "deck_payload": {"media_plan_options": [{"full_flight_total": {"cost": "$50,000", "impressions": "1,000,000"}}]}}},
        {"id": "prop-mw-1", "client_name": "Mattress Warehouse", "advertiser_id": "adv-mw",
         "vertical": "retail", "market": "DC", "generated_at": "2026-08-01T00:00:00",
         "form_json": {"proposal_title": "MW CTV", "plan_options": [],
                      "flight": {"label": "Jun-Jul 2026"},
                      "deck_payload": {"media_plan_options": [{"full_flight_total": {"cost": "$40,000", "impressions": "900,000"}}]}}},
    ]
    store.reports = [
        {"id": "rpt-waepa-1", "advertiser_id": "adv-waepa", "proposal_id": "prop-waepa-1",
         "report_json": {"headline_facts": {"period_start": "2026-10-01", "period_end": "2026-10-31",
                                            "attributed_rate": 0.014, "attributed_unique_visitors": 500,
                                            "attributed_conversions": None},
                         "report_type": "monthly", "whats_next": ["Keep going"]},
         "created_by": "Matt", "created_at": "2026-11-01T00:00:00"},
        {"id": "rpt-mw-1", "advertiser_id": "adv-mw", "proposal_id": "prop-mw-1",
         "report_json": {"headline_facts": {"period_start": "2026-07-01", "period_end": "2026-07-31",
                                            "attributed_rate": 0.02, "attributed_unique_visitors": 300,
                                            "attributed_conversions": None},
                         "report_type": "monthly", "whats_next": []},
         "created_by": "Matt", "created_at": "2026-08-01T00:00:00"},
    ]
    store.case_studies = [
        {"id": "cs-waepa-1", "title": "WAEPA case study", "advertiser_id": "adv-waepa",
         "verticals": ["banking"], "active": True},
        {"id": "cs-mw-1", "title": "MW case study", "advertiser_id": "adv-mw",
         "verticals": ["retail"], "active": True},
    ]

    at = new_app()
    at.session_state["page_choice"] = "Clients"
    at.session_state["nav_section"] = "Clients"
    at.session_state["client_view_picker"] = "WAEPA"
    at.run()
    check("no exception", not at.exception, at.exception)

    texts = _all_markdown_text(at)
    check("WAEPA's own proposal title shows", any("WAEPA CTV" in t for t in texts), None)
    check("MW's proposal title does NOT show", not any("MW CTV" in t for t in texts), None)
    check("WAEPA's own case study shows", any("WAEPA case study" in t for t in texts), None)
    check("MW's case study does NOT show", not any("MW case study" in t for t in texts), None)

    # A report row's own header (the period) lives on its st.expander LABEL,
    # not in any of _all_markdown_text's scanned element kinds.
    expander_labels = [str(e.label) for e in at.expander]
    check("WAEPA's own report period shows, on its expander header",
         any("2026-10-01" in label for label in expander_labels), expander_labels)
    check("MW's report period does NOT show", not any("2026-07-01" in label for label in expander_labels),
         expander_labels)

    # The trend table is a real st.dataframe widget -- its rows live on
    # .value (a real pandas DataFrame), not on any text element
    # _all_markdown_text scans either.
    trend_dataframes = [d for d in at.dataframe]
    check("a trend dataframe is present", bool(trend_dataframes), None)
    if trend_dataframes:
        trend_periods = list(trend_dataframes[0].value["Period"])
        check("the trend table shows WAEPA's own report period",
             any("2026-10-01" in str(p) for p in trend_periods), trend_periods)
        check("the trend table does NOT show MW's report period",
             not any("2026-07-01" in str(p) for p in trend_periods), trend_periods)


def check_vertical_write_back(store):
    print("\nVertical write-back")
    at = new_app()
    at.session_state["page_choice"] = "Clients"
    at.session_state["nav_section"] = "Clients"
    at.session_state["client_view_picker"] = "WAEPA"
    at.run()
    check("no exception", not at.exception, at.exception)

    vertical_selects = [s for s in at.selectbox if s.key and s.key.startswith("client_view_vertical_")]
    check("the vertical selectbox is present", bool(vertical_selects),
         [s.key for s in at.selectbox])
    if not vertical_selects:
        return
    vertical_selects[0].set_value("Healthcare").run()
    save_buttons = [b for b in at.button if b.key and b.key.startswith("client_view_save_vertical_")]
    check("the Save vertical button is present and enabled once the value changed",
         bool(save_buttons) and not save_buttons[0].disabled,
         [(b.key, b.disabled) for b in at.button if b.key and "vertical" in b.key])
    if not save_buttons:
        return
    save_buttons[0].click().run()
    check("no exception after saving", not at.exception, at.exception)
    check("set_advertiser_vertical was called with the WAEPA advertiser id and the new key",
         store.set_vertical_calls and store.set_vertical_calls[-1] == ("adv-waepa", "healthcare"),
         store.set_vertical_calls)


def check_trend_table_partial_coverage(store):
    """A report missing headline_facts (logged before the trend rework, or
    a legacy row) is reported as such, never silently dropped from the
    page or silently included with fabricated numbers."""
    print("\nTrend table: a report with no headline_facts is called out, not hidden")
    store.reports.append(
        {"id": "rpt-waepa-legacy", "advertiser_id": "adv-waepa", "proposal_id": None,
         "report_json": {"attribution": {}, "delivery": {}},  # no "headline_facts" key at all
         "created_by": "Matt", "created_at": "2026-06-01T00:00:00"})

    at = new_app()
    at.session_state["page_choice"] = "Clients"
    at.session_state["nav_section"] = "Clients"
    at.session_state["client_view_picker"] = "WAEPA"
    at.run()
    check("no exception", not at.exception, at.exception)
    texts = _all_markdown_text(at)
    check("the page says a report predates trend data, rather than showing nothing",
         any("before trend data" in t for t in texts), texts)


def check_merge_rename_admin_page(store):
    print("\nMerge/rename clients admin page")
    dup = {"id": "adv-waepa-dup", "canonical_name": "WAEPA TEST", "active": True}
    store.advertisers.append(dup)

    at = new_app()
    at.session_state["page_choice"] = "Merge/rename clients"
    at.run()
    check("no exception", not at.exception, at.exception)

    print("  Rename")
    rename_select = [s for s in at.selectbox if s.key == "advertiser_admin_rename_choice"]
    check("the rename client picker is present", bool(rename_select),
         [s.key for s in at.selectbox])
    if rename_select:
        rename_select[0].set_value("WAEPA TEST").run()
        rename_inputs = [t for t in at.text_input if t.key and t.key.startswith("advertiser_admin_rename_input_")]
        if rename_inputs:
            rename_inputs[0].set_value("WAEPA").run()
            rename_buttons = [b for b in at.button if b.key and b.key.startswith("advertiser_admin_rename_go_")]
            check("rename onto an existing ACTIVE name is refused, shown as an error",
                 bool(rename_buttons), [b.key for b in at.button])
            if rename_buttons:
                rename_buttons[0].click().run()
            check("no exception", not at.exception, at.exception)
            check("the refusal is shown to the rep",
                 any("merge" in str(e.value).lower() for e in at.error), [e.value for e in at.error])

    print("  Merge")
    merge_source = [s for s in at.selectbox if s.key == "advertiser_admin_merge_source"]
    merge_target = [s for s in at.selectbox if s.key == "advertiser_admin_merge_target"]
    check("source/target pickers are present", bool(merge_source) and bool(merge_target), None)
    if merge_source and merge_target:
        merge_source[0].set_value("WAEPA TEST").run()
        merge_target = [s for s in at.selectbox if s.key == "advertiser_admin_merge_target"]
        merge_target[0].set_value("WAEPA").run()
        typed_inputs = [t for t in at.text_input if t.key and t.key.startswith("advertiser_admin_merge_confirm_")]
        check("a typed-DELETE confirm field is offered", bool(typed_inputs),
             [t.key for t in at.text_input])
        if typed_inputs:
            typed_inputs[0].set_value("DELETE").run()
        merge_buttons = [b for b in at.button if b.key and b.key.startswith("advertiser_admin_merge_go_")]
        if merge_buttons:
            merge_buttons[0].click().run()
        check("no exception after merging", not at.exception, at.exception)
        check("merge_advertisers was called (WAEPA TEST -> WAEPA)",
             store.merge_calls and store.merge_calls[-1] == ("adv-waepa-dup", "adv-waepa"),
             store.merge_calls)
        source_after = next(r for r in store.advertisers if r["id"] == "adv-waepa-dup")
        check("the source is deactivated, not deleted", source_after.get("active") is False, source_after)
        check("the source row still exists", any(r["id"] == "adv-waepa-dup" for r in store.advertisers), None)


def main():
    store = FakeStore()
    app.db.fetch_advertisers = store.fetch_advertisers
    app.db.fetch_proposals = store.fetch_proposals
    app.db.fetch_proposal = store.fetch_proposal
    app.db.fetch_attribution_reports = store.fetch_attribution_reports
    app.db.fetch_case_studies = store.fetch_case_studies
    app.db.update_case_study = store.update_case_study
    app.db.set_advertiser_vertical = store.set_advertiser_vertical
    app.db.set_advertiser_name = store.set_advertiser_name
    app.db.advertiser_usage_counts = store.advertiser_usage_counts
    app.db.merge_advertisers = store.merge_advertisers

    check_waepa_scoping(store)
    check_vertical_write_back(store)
    check_trend_table_partial_coverage(store)
    check_merge_rename_admin_page(store)

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
