"""The in-app feedback loop (BACKLOG.md): a persistent "Report an issue"
popover on every page, and a "Feedback reports" admin page to
list/filter/close/export what comes in.

All of `db.submit_feedback`/`fetch_feedback`/`update_feedback_status`/
`count_open_feedback` are stubbed with an in-memory fake store -- the real
`feedback` table doesn't exist in Supabase until the schema migration in
supabase_schema.sql is pasted in by hand (DEPLOYMENT.md's own "schema
before code" rule), and even once it does, a test must never write real
rows into production, the same reason test_draft_regression.py stubs
db.log_proposal.

    python tests/test_feedback.py
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
db.active_deck_version = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402
from streamlit.testing.v1 import AppTest  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class FakeFeedbackStore:
    """An in-memory stand-in for the whole feedback table -- insert,
    filter, status update, open count -- so app.py's logic is tested
    without ever touching Supabase (which doesn't have this table until
    the migration runs, and shouldn't be written to by a test regardless).
    """
    def __init__(self):
        self.rows = []
        self._next_id = 1
        self.submit_calls = []
        self.unreachable = False

    def submit(self, category, notes, page, state, created_by=None):
        self.submit_calls.append(dict(category=category, notes=notes, page=page,
                                      state=state, created_by=created_by))
        if self.unreachable:
            return None, "Supabase isn't configured"
        row = {"id": str(self._next_id), "category": category, "notes": notes,
              "page": page, "created_by": created_by, "state_json": state,
              "status": "open", "created_at": "2026-08-23T12:00:00", "resolved_at": None}
        self._next_id += 1
        self.rows.append(row)
        return row["id"], None

    def fetch(self, status=None, category=None, limit=500):
        if self.unreachable:
            return None, "Supabase isn't configured"
        rows = self.rows
        if status:
            rows = [r for r in rows if r["status"] == status]
        if category:
            rows = [r for r in rows if r["category"] == category]
        return list(reversed(rows))[:limit], None

    def update_status(self, feedback_id, status):
        if self.unreachable:
            return False, "Supabase isn't configured"
        for row in self.rows:
            if row["id"] == feedback_id:
                row["status"] = status
        return True, None

    def count_open(self):
        if self.unreachable:
            return None, "Supabase isn't configured"
        return sum(1 for r in self.rows if r["status"] == "open"), None


def wire_store(store):
    db.submit_feedback = store.submit
    db.fetch_feedback = store.fetch
    db.update_feedback_status = store.update_status
    db.count_open_feedback = store.count_open


def new_app():
    from datetime import date
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates the Build page's own
    # body on market+flight -- the feedback popover renders on every page
    # (per this file's own subject), but the Build-page fields this file
    # also inspects need the gate open.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def main():
    print("=" * 78)
    print("SCENARIO  the popover renders on the Build page, and submitting "
          "captures the real source-of-truth session state (not a "
          "data_editor delta) and reaches db.submit_feedback with it")
    print("=" * 78)
    store = FakeFeedbackStore()
    wire_store(store)
    at = new_app()
    at.run()
    check("no exception on first render", at.exception.len == 0, at.exception)
    check("the popover control is present in the sidebar",
          any(b.key == "feedback_submit_0" for b in at.sidebar.button), "not found")
    submit_btn = [b for b in at.sidebar.button if b.key == "feedback_submit_0"][0]
    check("submit is disabled with no notes typed", submit_btn.disabled, submit_btn)

    # client_name is set through its own widget (matching a real value= would
    # otherwise trip Streamlit's "default value AND Session State API" warning);
    # targeting_groups/plan_options are plain data, not widget-backed, so
    # they're set directly, the same way a real avails import or draft would
    # populate them before a rep ever opens the popover.
    at.text_input(key="client_name").set_value("Test Feedback Client").run()
    real_group = tg.new_group(["Test"], geo_def={"kind": "zips", "zips": []})
    # D2 actually renders now (FLOW_REWORK_PLAN.md Phase 1 opened the gate
    # this test never used to get past), and its grid fold-back normalizes
    # every group with an explicit avails_full_flight -- None here, since
    # this group was never avails-imported -- where new_group() itself
    # leaves the key absent. Real, correct behavior; match it rather than
    # asserting a shape D2 never actually produces.
    real_group["avails_full_flight"] = None
    at.session_state["targeting_groups"] = [real_group]
    real_option = app.new_plan_option("Option A", [])
    at.session_state["plan_options"] = [real_option]
    at.text_area(key="feedback_notes_0").set_value("The avails row disappeared after import.")
    at.sidebar.selectbox(key="feedback_category_0").set_value("Avails")
    at.run()
    at.sidebar.button(key="feedback_submit_0").click().run()
    check("no exception on submit", at.exception.len == 0, at.exception)
    check("exactly one submit reached db.submit_feedback", len(store.submit_calls) == 1, store.submit_calls)
    call = store.submit_calls[0] if store.submit_calls else {}
    check("category carried through", call.get("category") == "Avails", call)
    check("notes carried through", call.get("notes") == "The avails row disappeared after import.", call)
    check("page carried through", call.get("page") == "Build a proposal", call)
    check("created_by carried through", call.get("created_by") == "T", call)
    state = call.get("state") or {}
    check("build_stamp captured", state.get("build_stamp") == app.BUILD_STAMP, state)
    check("client_name captured from the real form", state.get("client_name") == "Test Feedback Client", state)
    check("targeting_groups captured verbatim (the real source of truth)",
          state.get("targeting_groups") == [real_group], state)
    check("plan_options captured verbatim", state.get("plan_options") == [real_option], state)
    check("a success message is shown", "Report submitted" in " ".join(
        s.value for s in at.sidebar.success), [s.value for s in at.sidebar.success])
    check("the form moved to a fresh generation (old notes key no longer bound)",
          "feedback_notes_1" in [t.key for t in at.text_area], [t.key for t in at.text_area])

    print("\n" + "=" * 78)
    print("SCENARIO  a submit failure (backend unreachable) is shown, not raised")
    print("=" * 78)
    store2 = FakeFeedbackStore()
    store2.unreachable = True
    wire_store(store2)
    at = new_app()
    at.run()
    at.text_area(key="feedback_notes_0").set_value("Something's off with the map.")
    at.run()
    at.sidebar.button(key="feedback_submit_0").click().run()
    check("no exception even though the backend is unreachable", at.exception.len == 0, at.exception)
    check("an error is shown to the rep", any(
        "Couldn't submit" in e.value for e in at.sidebar.error), [e.value for e in at.sidebar.error])
    check("the form does NOT advance to a new generation on failure "
          "(nothing to lose by retrying the same box)",
          "feedback_notes_1" not in [t.key for t in at.text_area], [t.key for t in at.text_area])

    print("\n" + "=" * 78)
    print("SCENARIO  the Feedback reports admin page lists, filters, and "
          "closes reports")
    print("=" * 78)
    store3 = FakeFeedbackStore()
    wire_store(store3)
    store3.submit("Avails", "Zips missing from the map", "Zip/map builder",
                  {"build_stamp": "abc123", "page": "Zip/map builder"}, created_by="Andrea")
    store3.submit("Drafting", "Wrong CPM on a drafted line", "Build a proposal",
                  {"build_stamp": "abc123", "page": "Build a proposal"}, created_by="Matt")
    at = new_app()
    at.session_state["page_choice"] = "Feedback reports"
    at.run()
    check("no exception on the admin page", at.exception.len == 0, at.exception)
    check("open count shown", "2 open" in " ".join(c.value for c in at.caption), "not found")
    check("both reports render as expander headers", len(at.expander) >= 2, len(at.expander))

    print("\n" + "=" * 78)
    print("SCENARIO  filtering by category narrows the list")
    print("=" * 78)
    at.selectbox(key="feedback_category_filter").set_value("Avails").run()
    check("no exception filtering", at.exception.len == 0, at.exception)
    check("only the Avails report's text appears",
          any("Zips missing from the map" in e.label for e in at.expander), [e.label for e in at.expander])
    check("the Drafting report is filtered out",
          not any("Wrong CPM" in e.label for e in at.expander), [e.label for e in at.expander])
    at.selectbox(key="feedback_category_filter").set_value("All").run()

    print("\n" + "=" * 78)
    print("SCENARIO  marking a report closed reaches db.update_feedback_status")
    print("=" * 78)
    row_id = store3.rows[0]["id"]
    at.selectbox(key=f"fb_status_{row_id}").set_value("closed").run()
    check("no exception marking closed", at.exception.len == 0, at.exception)
    check("the store reflects the new status",
          next(r for r in store3.rows if r["id"] == row_id)["status"] == "closed", store3.rows)

    print("\n" + "=" * 78)
    print("SCENARIO  the admin page degrades gracefully when the backend "
          "is unreachable, same as every other loader in this app")
    print("=" * 78)
    store4 = FakeFeedbackStore()
    store4.unreachable = True
    wire_store(store4)
    at = new_app()
    at.session_state["page_choice"] = "Feedback reports"
    at.run()
    check("no exception when unreachable", at.exception.len == 0, at.exception)
    check("a warning is shown, not a crash",
          any("Supabase" in w.value for w in at.warning), [w.value for w in at.warning])

    print("\n" + "=" * 78)
    print("SCENARIO  markdown export format")
    print("=" * 78)
    rows = [{"category": "Avails", "notes": "Zips missing from the map",
            "page": "Zip/map builder", "created_by": "Andrea", "created_at": "2026-08-23T12:00:00",
            "state_json": {"build_stamp": "abc123", "targeting_groups": [{"id": "g1"}]}}]
    markdown = app._feedback_export_markdown(rows)
    check("category heading present", "[Avails]" in markdown, markdown)
    check("the reporter's name is in the export", "Andrea" in markdown, markdown)
    check("the description text is in the export", "Zips missing from the map" in markdown, markdown)
    check("the captured state is embedded as JSON", '"build_stamp": "abc123"' in markdown, markdown)

    # Restore real db functions so nothing downstream of this module import
    # accidentally keeps the fakes wired.
    import importlib
    importlib.reload(db)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A rep can report an issue from any page (state captured, never a screenshot -- "
          "deferred), a failed submit degrades to a visible error rather than a crash, and "
          "the Feedback reports admin page lists, filters, closes and exports reports the "
          "same offline-safe way every other admin page in this app does.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
