"""FLOW_REWORK_PLAN.md Phase 2: `app.format_flight_shorthand` -- the single
owner of the day-precise plan-cell/Campaign-Specs flight string.

    python tests/test_flight_shorthand.py

Pure, offline, no Streamlit runtime, no AppTest -- `app.flight_month_ranges`
builds the range list directly, exactly the shape the real form produces.

What this proves, once, before the renderer reaches any of its seven
downstream call sites (deliberately NOT wired to any of them yet -- see
FLOW_REWORK_PLAN.md's commit sequence):

1. The plan doc's own worked example, verbatim.
2. A single full month renders bare; a single partial month renders with
   its own start/end day.
3. Consecutive full months collapse to their endpoints; a skipped month's
   absence breaks a run rather than being bridged over.
4. A year boundary is never bridged, even between calendar-adjacent months,
   and years are omitted entirely unless the active months actually span
   more than one.
5. Empty input renders "", the same convention format_flight_label uses.
"""
import os
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
os.chdir(Path(__file__).resolve().parent.parent)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def ranges(*entries):
    """entries: (month_label, start, end, active=True)."""
    out = []
    for e in entries:
        month, start, end = e[0], e[1], e[2]
        active = e[3] if len(e) > 3 else True
        out.append({"month": month, "start": start, "end": end, "active": active})
    return out


def main():
    print("the plan doc's own worked example (Capital Media): Sep 21 - Dec 14")
    capital_media = app.flight_month_ranges(date(2026, 9, 21), date(2026, 12, 14))
    shorthand = app.format_flight_shorthand(capital_media)
    print(f"    rendered: {shorthand!r}")
    check("matches the doc's own example exactly",
          shorthand == "Sep 21–30, Oct–Nov, Dec 1–14", shorthand)

    print("\na 12-month flight, single calendar year")
    twelve_month = app.flight_month_ranges(date(2026, 1, 1), date(2026, 12, 31))
    shorthand12 = app.format_flight_shorthand(twelve_month)
    print(f"    rendered: {shorthand12!r}")
    check("all twelve months are full and collapse to one run, no year shown",
          shorthand12 == "Jan–Dec", shorthand12)

    print("\na single full month renders bare")
    check("bare month name", app.format_flight_shorthand(
        ranges(("Oct 2026", date(2026, 10, 1), date(2026, 10, 31)))) == "Oct")

    print("\na single partial month renders with its own start/end day")
    check("day range", app.format_flight_shorthand(
        ranges(("Dec 2026", date(2026, 12, 1), date(2026, 12, 14)))) == "Dec 1–14")

    print("\nall-full months collapse into one run")
    check("Oct-Nov collapses", app.format_flight_shorthand(
        ranges(("Oct 2026", date(2026, 10, 1), date(2026, 10, 31)),
               ("Nov 2026", date(2026, 11, 1), date(2026, 11, 30)))) == "Oct–Nov")

    print("\na skipped middle month breaks the run rather than being bridged over")
    check("Sep, Nov -- not Sep-Nov", app.format_flight_shorthand(
        ranges(("Sep 2026", date(2026, 9, 1), date(2026, 9, 30)),
               ("Oct 2026", date(2026, 10, 1), date(2026, 10, 31), False),
               ("Nov 2026", date(2026, 11, 1), date(2026, 11, 30)))) == "Sep, Nov")

    print("\na year boundary is never bridged, even between calendar-adjacent full months")
    year_boundary = app.format_flight_shorthand(
        ranges(("Nov 2026", date(2026, 11, 1), date(2026, 11, 30)),
               ("Dec 2026", date(2026, 12, 1), date(2026, 12, 31)),
               ("Jan 2027", date(2027, 1, 1), date(2027, 1, 31))))
    print(f"    rendered: {year_boundary!r}")
    check("Nov-Dec 2026, Jan 2027 -- two runs, not one bridging Dec into Jan",
          year_boundary == "Nov–Dec 2026, Jan 2027", year_boundary)

    print("\nyears are omitted entirely when everything active shares one calendar year")
    check("no year shown", app.format_flight_shorthand(
        ranges(("Dec 2026", date(2026, 12, 1), date(2026, 12, 31)))) == "Dec")

    print("\nempty input renders the empty string, same convention as format_flight_label")
    check("no active months at all", app.format_flight_shorthand(
        ranges(("Sep 2026", date(2026, 9, 1), date(2026, 9, 30), False))) == "")
    check("no ranges at all", app.format_flight_shorthand([]) == "")

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("format_flight_shorthand renders the doc's own worked example exactly, collapses "
          "consecutive full months, breaks runs at a skipped month and at a year boundary, "
          "and omits years within a single calendar year.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
