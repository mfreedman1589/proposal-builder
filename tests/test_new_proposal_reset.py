""""New proposal" clears everything the previous proposal left (2026-11-19).

    python tests/test_new_proposal_reset.py

Draft client A the way a rep does -- notes uploaded as a file, an avails PDF
and a logo uploaded, the Draft button pressed, a review row answered and
redraft notes typed -- then press New proposal and assert every input is
empty, on the run after the clear and on the run after that.

Why the uploads are modelled the way they are: an uploaded file lives in the
browser, attached to its widget, and is sent again on every run for as long
as that widget is on the page. Deleting the widget's session_state does not
detach it. That is the bug this guards: the clear deleted the "already
loaded" markers, the browser sent the same notes file and avails PDF again
under the same keys, and the app read them back in -- notes box, avails
table and logo all returned. AppTest can't operate a file uploader, and the
test-mode path injection lives in session_state (which the clear wipes), so
neither could see it. `FakeBrowser` stands in for the browser instead: it
keeps a file per widget key until that key leaves the page, exactly the
lifetime the real one gives it.

The model call is stubbed with a recorded draft; the avails parser returns a
synthetic one-row document (no client PDF is committed).
"""
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
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
os.environ.setdefault("PROPOSAL_BUILDER_TODAY", "2026-08-01")

import streamlit                                 # noqa: E402
import avails_pdf_import as api                  # noqa: E402
import claude_client                             # noqa: E402
import db                                        # noqa: E402

db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.log_proposal = lambda *a, **k: (True, None)

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}" + (f"  -- {detail}" if detail != "" else ""))
        failures.append(label)


class _File:
    def __init__(self, name, data):
        self.name = name
        self._data = data

    def getvalue(self):
        return self._data


class FakeBrowser:
    """Files attached to uploader widgets, by key, for as long as the
    widget stays on the page."""

    def __init__(self):
        self.files = {}
        self.rendered = set()
        self._real = streamlit.file_uploader

    def file_uploader(self, label, *args, key=None, **kwargs):
        self.rendered.add(key)
        return self.files.get(key)

    def install(self):
        streamlit.file_uploader = self.file_uploader

    def uninstall(self):
        streamlit.file_uploader = self._real

    def attach(self, prefix, file):
        keys = [k for k in self.rendered if str(k).startswith(prefix)]
        assert len(keys) == 1, (prefix, sorted(self.rendered))
        self.files[keys[0]] = file
        return keys[0]

    def run(self, at):
        """One browser interaction: the app runs (and any reruns it asks
        for), then a widget that wasn't drawn drops its file."""
        self.rendered = set()
        at.run()
        self.files = {k: v for k, v in self.files.items() if k in self.rendered}
        return at


NOTES = ("Client: QA-TEST-Client A, Rockville MD. HVAC. Flight Sep 8 - Nov 30, 2026. "
         "Budget $20,000 a month across Premion Streaming TV and retargeting.")
AVAILS_DOC = api.AvailsDocument(
    rfpid="RFPID-RESET-1", advertiser="QA-TEST-Client A", agency="No Agency",
    flight_start=date(2026, 9, 8), flight_end=date(2026, 11, 30), total_impressions=900_000,
    attribution_text="",
    groups=[api.AvailsGroup(audience_text="RETAIL Home Services HVAC", geo_kind=api.GEO_KIND_DMA,
                            geo_name="Washington, DC (Hagerstown, MD) DMA", impressions=900_000,
                            row_count=1, rfpi="RFPI-5550001")],
    source_name="client_a_avails.pdf")


def main():
    from streamlit.testing.v1 import AppTest

    draft = json.loads((FIXTURES / "hvac_two_option.draft.json").read_text(encoding="utf-8"))
    draft.pop("_comment", None)
    draft["client_name"] = "QA-TEST-Client A"
    calls = []

    def fake_model(prompt, label="draft", attempts=2, on_attempt=None):
        calls.append(label)
        return copy.deepcopy(draft), None

    claude_client._call_claude_json = fake_model
    api.parse_avails_pdf = lambda path, name=None: AVAILS_DOC

    browser = FakeBrowser()
    browser.install()
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=240)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Reset Suite"
        at.session_state["flight_start"] = date(2026, 9, 8)
        at.session_state["flight_end"] = date(2026, 11, 30)
        at.session_state["client_name"] = "QA-TEST-Client A"
        browser.run(at)
        check("the Build page renders", not at.exception, [e.value for e in at.exception])

        print("\nClient A: notes file, avails PDF and logo uploaded")
        browser.attach("notes_upload", _File("client_a_notes.txt", NOTES.encode()))
        browser.attach("avails_pdf_upload_intake", _File("client_a_avails.pdf", b"%PDF-stub"))
        browser.attach("logo_upload", _File("client_a.png", b"\x89PNG-test"))
        browser.run(at)
        check("the notes file filled the notes box",
              at.session_state["draft_notes_input"] == NOTES, at.session_state["draft_notes_input"])
        groups = [g for g in at.session_state["targeting_groups"] if g.get("terms")]
        check("the avails PDF made a targeting group", len(groups) == 1, groups)
        check("the logo was kept", "uploaded_logo" in at.session_state)

        print("\nClient A: drafted, a review row answered, redraft notes typed")
        next(b for b in at.button if b.label == "Draft proposal from notes").click()
        browser.run(at)
        check("the draft ran", calls == ["draft"] and at.session_state["draft_round"] == 1, calls)
        at.session_state["draft_review_answers"] = {"some_row": "other"}
        boxes = [t for t in at.text_area if t.key == "draft_clarifications"]
        check("the redraft box is on the page", len(boxes) == 1)
        if boxes:
            boxes[0].input("Drop the retargeting line.")
        browser.run(at)

        print("\nNew proposal -> Clear the form")
        at.session_state["confirm_new_proposal"] = True
        browser.run(at)
        next(b for b in at.button if b.label == "Clear the form").click()
        browser.run(at)
        check("clearing raises nothing", not at.exception, [e.value for e in at.exception])
        for label in ("the run after the clear", "the run after that"):
            print(f"  ({label})")
            ss = at.session_state
            check("the notes box is empty", ss["draft_notes_input"] == "", ss["draft_notes_input"])
            check("no notes file is loaded", "notes_text_loaded" not in ss)
            check("every uploader on the page is empty",
                  not any(browser.files.get(k) for k in browser.rendered),
                  sorted(k for k in browser.rendered if browser.files.get(k)))
            real = [g for g in (ss["targeting_groups"] if "targeting_groups" in ss else [])
                    if g.get("terms")]
            check("the avails table has no imported rows", not real, real)
            check("no avails file is listed",
                  not (ss["avails_documents"] if "avails_documents" in ss else None))
            check("no logo", "uploaded_logo" not in ss)
            check("the client name is empty", ss["client_name"] == "", ss["client_name"])
            left = sorted(k for k in ss.filtered_state
                          if str(k).startswith(("draft_", "_draft", "_review", "rv_"))
                          and k not in ("draft_from_notes", "draft_notes_input"))
            check("no draft state is left (draft, review answers, redraft notes)", not left, left)
            check("the model was not called again", calls == ["draft"], calls)
            browser.run(at)
    finally:
        browser.uninstall()

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("All New proposal reset checks passed.")


if __name__ == "__main__":
    main()
