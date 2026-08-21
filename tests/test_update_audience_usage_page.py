"""The "Update audience usage" admin page (app.py's
render_update_audience_usage) -- offline, via AppTest, against a small
synthetic workbook injected the same way the Wide Orbit uploader's own
tests do (`test_mode_upload`, since AppTest can't operate a real
file_uploader). Checks the page RENDERS the report correctly; activating
against real Supabase is out of scope here the same way the master-deck
upload page's own Activate step is untested offline.

db.fetch_audiences is stubbed to force the local brochure-PDF fallback --
NOT to simulate "Supabase unreachable" for its own sake, but because this
machine's .streamlit/secrets.toml points at a real, live project (checked
directly: the real "Update audience usage" page has actually been used
against it), so an unstubbed run would score "Collided" against whatever
happens to be live at test time instead of the fixed 369-segment brochure
this test's own numbers are written against. A test that silently drifts
with production data isn't testing anything.

    python tests/test_update_audience_usage_page.py
"""
import os
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation -- see module docstring")

import openpyxl                                # noqa: E402
from streamlit.testing.v1 import AppTest       # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def synthetic_workbook():
    """A tiny workbook exercising: a CLT row to exclude, a client-pattern
    row to exclude mechanically, a CUSTOM row to keep, a genuinely
    uncategorized row, and a No Data Targeting row to drop."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Segment Name", "Delivered Impressions"])
    for segment, impressions in [
        ("No Data Targeting", 999),
        ("AUTO Intenders", 1000),
        ("CLT Navy Federal Credit Union Lookalike", 2000),
        ("CUSTOM FOOD Grocery Delivery Services", 3000),
        ("Pima Medical Institute_RFPID-254492_RT", 4000),  # client-pattern (RFPID...RT)
        ("Gizmo Enthusiasts Weekly Roundup", 500),         # genuinely uncategorized
    ]:
        ws.append([segment, impressions])
    handle = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    handle.close()
    wb.save(handle.name)
    return handle.name


def main():
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=120)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["page_choice"] = "Update audience usage"
    at.session_state["usage_upload_path"] = synthetic_workbook()
    at.run()
    check("the page renders with an injected workbook, no exception",
          not at.exception, at.exception[0].message[:500] if at.exception else "")
    if at.exception:
        print(f"\n1 FAILED: {failures}")
        return 1

    metrics = {m.label: m.value for m in at.metric}
    check("all six report metrics render",
          set(metrics) == {"Stacks", "Dropped", "Gained", "Collided", "Uncategorized",
                           "Client-excluded"}, metrics)
    check("the No Data Targeting row was dropped, not counted as a stack",
          metrics.get("Dropped") == "1", metrics)
    # AUTO Intenders is already a real brochure segment (the local catalog
    # this page merges against, per the fetch_audiences stub above), so it
    # collides rather than being gained -- only the CUSTOM and Gizmo rows
    # are genuinely new (Pima is excluded, not gained).
    check("AUTO Intenders collides with the existing (brochure) catalog; "
          "CUSTOM FOOD... and Gizmo... are the two genuinely new segments",
          metrics.get("Collided") == "1" and metrics.get("Gained") == "2", metrics)
    check("the Gizmo row (no recognizable prefix) is the one uncategorized segment",
          metrics.get("Uncategorized") == "1", metrics)
    check("the Pima row (RFPID...RT) is mechanically client-excluded, not uncategorized",
          metrics.get("Client-excluded") == "1", metrics)

    body_text = "\n".join(md.value for md in at.markdown) + "\n".join(c.value for c in at.caption)
    check("the CLT exclusion expander names the excluded segment",
          "CLT Navy Federal Credit Union Lookalike" in body_text, None)
    check("the client-pattern exclusion expander names the Pima segment",
          "Pima Medical Institute" in body_text, None)
    check("the CUSTOM segment (its own second prefix resolves to FOOD) is not "
          "in the uncategorized count",
          metrics.get("Uncategorized") == "1", metrics)
    check("an Activate button is offered once the report renders",
          any(b.label == "Activate this workbook" for b in at.button), [b.label for b in at.button])

    print("\nthe Claude-categorize review section (data_editor content isn't readable via "
          "AppTest, so this checks what surrounds it: the section itself, the ask-Claude "
          "button, and the unconfirmed-count caption -- tests/test_categorize_prompt.py "
          "covers the suggestion logic itself, offline)")
    check("a 'Needs a category -- 1' subheader renders for the one unresolved component",
          any("Needs a category" in s.value and "1" in s.value for s in at.subheader),
          [s.value for s in at.subheader])
    check("an 'Ask Claude to propose categories' button is offered",
          any("Ask Claude to propose categories" in b.label for b in at.button),
          [b.label for b in at.button])
    check("with nothing confirmed yet, the unconfirmed-count caption says so",
          any("still has no category" in c.value for c in at.caption),
          [c.value for c in at.caption])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The admin page renders the parsed report -- gained/collided/uncategorized counts, "
          "the CLT exclusion list, and the uncategorized list -- from an injected workbook, "
          "before any write happens.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
