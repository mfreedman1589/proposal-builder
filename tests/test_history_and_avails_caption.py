"""Two QA round-2 fixes (Visit Hershey session, 2026-10-02).

History total: a 2-option proposal listed as $15,000 -- the SUM of a
$7,000 and an $8,000 option, though a client buys one. Matt's rule: each
option named with its own total, the single figure for one option, never
the sum; long names shortened, never dropped.

Avails caption: 103.6M available impressions beside a 781K/month plan
confused a client. One plain line under the table: "Available monthly
impressions by market — your plan buys 781,248/month."

    python tests/test_history_and_avails_caption.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

from pptx import Presentation                   # noqa: E402

import app                                      # noqa: E402
import assembly                                 # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def check_history():
    print("History total -- per option, never summed")
    two = app.history_budget_text([("CTV Plan", 3500), ("NFL Package", 4000)])
    check("two options: each named with its own total, Matt's format",
          two == "Option 1 – CTV Plan $3,500 · Option 2 – NFL Package $4,000", two)
    check("...and the $7,500 sum appears nowhere", "7,500" not in two, two)
    muse = app.history_budget_text([("Option A — $3,500/mo", 7000),
                                    ("Option B — $4,000/mo (stretch)", 8000)])
    check("Muse's proposal: $7,000 and $8,000 listed, never $15,000",
          "$7,000" in muse and "$8,000" in muse and "15,000" not in muse, muse)
    check("a name that already says 'Option' isn't prefixed twice",
          not muse.startswith("Option 1 – Option"), muse)
    long_name = app.history_budget_text([("A very long option name that goes on and on", 1000),
                                         ("Short", 2000)])
    check("a long name is shortened with an ellipsis, not dropped",
          "…" in long_name and "A very long" in long_name, long_name)
    check("one option: just its figure", app.history_budget_text([("Option A", 7000)]) == "$7,000")
    row = {"form_json": {"plan_options": [
        {"name": "Option A", "totals": {"full_flight_cost": 7000}},
        {"name": "Option B", "totals": {"full_flight_cost": 8000}}]}}
    check("the History row summary carries one figure per option",
          app._proposal_summary(row)["budgets"] == [("Option A", 7000.0), ("Option B", 8000.0)],
          app._proposal_summary(row))


def check_caption():
    print("\nAvails caption")
    options = [{"name": "Option A"}]
    results = [{"monthly_impressions": 781248.4, "full_flight_impressions": 2343745}]
    check("monthly basis: Matt's sentence, with the plan's own monthly figure",
          app.avails_plan_caption(options, results)
          == "Available monthly impressions by market — your plan buys 781,248/month.",
          app.avails_plan_caption(options, results))
    check("full-flight basis says so",
          app.avails_plan_caption(options, results, full_flight=True)
          == "Available impressions for the flight by market — your plan buys 2,343,745 for the flight.",
          app.avails_plan_caption(options, results, full_flight=True))
    two = app.avails_plan_caption([{"name": "Option A"}, {"name": "Option B"}],
                                  [{"monthly_impressions": 781248}, {"monthly_impressions": 900000}])
    check("two options: each figure named", "781,248/month (Option A) or 900,000/month (Option B)"
          in (two or ""), two)
    check("nothing priced yet -> no caption",
          app.avails_plan_caption(options, [{"monthly_impressions": 0}]) is None)

    deck = REPO / "TEGNA_MASTER_DECK_v1_1.pptx"
    if not deck.exists():
        print(f"  SKIP  {deck.name} not present")
        return
    prs = Presentation(str(deck))
    slide = assembly.find_slide_with_marker(prs, "{{AVAILS}}")
    box = assembly.add_avails_caption(slide, "Available monthly impressions by market — your "
                                             "plan buys 781,248/month.")
    table_end = assembly.table_bottom(slide)
    check("the caption sits directly under the table, inside the slide",
          box is not None and box.top >= table_end
          and box.top + box.height <= prs.slide_height, (box and box.top, table_end))
    check("it is one shape named AvailsCaption carrying the text",
          box.name == "AvailsCaption" and "781,248/month" in box.text_frame.text)
    check("no caption text -> nothing added",
          assembly.add_avails_caption(slide, None) is None)


def main():
    check_history()
    check_caption()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
