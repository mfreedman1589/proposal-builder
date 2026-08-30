"""The band-dates-change-after-customizing path: customize a per-month range
in the Custom flighting control, then change the setup band's flight dates,
and check what the control ends up showing. Reported as a suspected bug
(band dates change after customizing a month leaves the control in a wrong
state); this file is the regression guard promised for that path, whether
or not a live defect was found.

Investigation (both `AppTest` and the real running app, driven over
mcp claude-in-chrome) reproduced the EXACT reported sequence --

    1. Band flight Sep 1 - Nov 30
    2. Open the flighting control, edit September to Sep 8-30
    3. Change the band flight to Oct 1 - Dec 31

-- and two adjacent variants (customizing a month that SURVIVES the band
change instead of one that's dropped; narrowing the band within the SAME
three months, which -- unlike a month added or removed -- does not bump
`flight_months_gen`) three separate ways with no defect found:

    - `app.flight_month_ranges` (the single reconciliation owner) always
      drops a month that left the flight, re-clips a surviving month's
      stored range into the new bounds, and defaults a newly-added month --
      exactly the documented rule.
    - Streamlit's own `date_input` silently re-clamps an out-of-bounds
      session-state-backed value to the new min_value/max_value on the
      very next render, even when the widget key hasn't changed (no
      generation bump) -- confirmed against a minimal isolated widget, not
      just this app's own reconciliation, so it isn't this app's own
      correction papering over a real disagreement.
    - Driving the real running app end to end (typing dates, opening the
      expander, unchecking a month) showed the caption, the control and the
      "N active month(s)" summary all agreeing at every step.

So this suite is a regression guard for a path that had none, not a fix for
a demonstrated defect. If the defect is real, it needs a more precise
sequence than the one investigated here (see CLAUDE.md's own rule: assert
against something the code can't move, and if a value can legitimately live
in more than one place, print all of them).

FLOW_REWORK_PLAN.md Phase 4a moved this control from a Section E expander
into a "Custom flighting" checkbox in the setup band -- the reconciliation
mechanism this guards (`flight_month_ranges`, the generation bump) is
unchanged, only the widget shape and its location. The AppTest sequence
below now ticks `custom_flighting` before touching a per-month widget key,
since the rows only render, and their keys only mean anything, once it's
ticked.

    python tests/test_custom_flighting.py
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

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def by_month(ranges):
    return {r["month"]: r for r in ranges}


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    return at


def session_get(at, key, default=None):
    try:
        return at.session_state[key]
    except (KeyError, AttributeError):
        return default


def main():
    print("pure reconciliation: the exact reported sequence -- customize "
          "September (a month the band change is about to drop), then move "
          "the band to Oct 1 - Dec 31")
    stored = [
        {"month": "Sep 2026", "start": date(2026, 9, 8), "end": date(2026, 9, 30), "active": True},
        {"month": "Oct 2026", "start": date(2026, 10, 1), "end": date(2026, 10, 31), "active": True},
        {"month": "Nov 2026", "start": date(2026, 11, 1), "end": date(2026, 11, 30), "active": True},
    ]
    result = by_month(app.flight_month_ranges(date(2026, 10, 1), date(2026, 12, 31), stored))
    check("September is gone", "Sep 2026" not in result, result)
    check("October kept, re-clipped (unchanged here, full month either side)",
          result["Oct 2026"]["start"] == date(2026, 10, 1)
          and result["Oct 2026"]["end"] == date(2026, 10, 31),
          result.get("Oct 2026"))
    check("November kept, re-clipped",
          result["Nov 2026"]["start"] == date(2026, 11, 1)
          and result["Nov 2026"]["end"] == date(2026, 11, 30),
          result.get("Nov 2026"))
    check("December appears with a default range",
          result["Dec 2026"]["start"] == date(2026, 12, 1)
          and result["Dec 2026"]["end"] == date(2026, 12, 31),
          result.get("Dec 2026"))

    print("\npure reconciliation: customizing a SURVIVING month (October) "
          "instead of the dropped one -- the customization must carry "
          "through the band change untouched")
    stored2 = [
        {"month": "Sep 2026", "start": date(2026, 9, 1), "end": date(2026, 9, 30), "active": True},
        {"month": "Oct 2026", "start": date(2026, 10, 5), "end": date(2026, 10, 20), "active": True},
        {"month": "Nov 2026", "start": date(2026, 11, 1), "end": date(2026, 11, 30), "active": True},
    ]
    result2 = by_month(app.flight_month_ranges(date(2026, 10, 1), date(2026, 12, 31), stored2))
    check("October's custom 5-20 window survives byte-for-byte",
          result2["Oct 2026"]["start"] == date(2026, 10, 5)
          and result2["Oct 2026"]["end"] == date(2026, 10, 20),
          result2.get("Oct 2026"))

    print("\npure reconciliation: narrowing the band WITHIN the same three "
          "months (no month added or removed, so the generation never "
          "bumps) -- a stored range outside the new bounds must be CLAMPED, "
          "not left stale")
    stored3 = [
        {"month": "Sep 2026", "start": date(2026, 9, 8), "end": date(2026, 9, 30), "active": True},
        {"month": "Oct 2026", "start": date(2026, 10, 1), "end": date(2026, 10, 31), "active": True},
        {"month": "Nov 2026", "start": date(2026, 11, 1), "end": date(2026, 11, 30), "active": True},
    ]
    result3 = by_month(app.flight_month_ranges(date(2026, 9, 15), date(2026, 11, 30), stored3))
    check("September's stale start (Sep 8, now before the new Sep 15 band "
          "start) is clamped up to Sep 15, not left at Sep 8",
          result3["Sep 2026"]["start"] == date(2026, 9, 15), result3.get("Sep 2026"))
    check("September's end is untouched (still within bounds)",
          result3["Sep 2026"]["end"] == date(2026, 9, 30), result3.get("Sep 2026"))

    print("\nAppTest, end to end: the literal reported sequence driven "
          "through the real widgets -- customize September, then move the "
          "band to Oct 1 - Dec 31, and check both the control's own state "
          "and the band's one-line summary agree")
    at = new_app()
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.run()
    check("no exception establishing the baseline", not at.exception,
          at.exception[0].message[:400] if at.exception else "")

    fm = session_get(at, "flight_months")
    sep_idx = next(i for i, r in enumerate(fm) if r["month"].startswith("Sep"))
    gen = session_get(at, "flight_months_gen", 0)
    # FLOW_REWORK_PLAN.md Phase 4a: the per-month rows only render, and their
    # widget keys only mean anything, once "Custom flighting" is ticked --
    # the control moved into the setup band behind that checkbox.
    at.session_state["custom_flighting"] = True
    at.session_state[f"fm_start_{gen}_{sep_idx}"] = date(2026, 9, 8)
    at.run()
    check("no exception after customizing September", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("September customization landed in flight_months",
          by_month(session_get(at, "flight_months"))["Sep 2026"]["start"] == date(2026, 9, 8),
          session_get(at, "flight_months"))

    at.session_state["flight_start"] = date(2026, 10, 1)
    at.session_state["flight_end"] = date(2026, 12, 31)
    at.run()
    check("no exception after the band change", not at.exception,
          at.exception[0].message[:400] if at.exception else "")

    final = by_month(session_get(at, "flight_months"))
    check("September is gone from flight_months", "Sep 2026" not in final, final)
    check("October and November are present, at their default full-month range",
          final.get("Oct 2026", {}).get("start") == date(2026, 10, 1)
          and final.get("Nov 2026", {}).get("start") == date(2026, 11, 1),
          final)
    check("December was added at its default range",
          final.get("Dec 2026", {}).get("start") == date(2026, 12, 1)
          and final.get("Dec 2026", {}).get("end") == date(2026, 12, 31),
          final)
    check("active_months agrees with flight_months (no stale entries survive)",
          session_get(at, "active_months") == ["Oct 2026", "Nov 2026", "Dec 2026"],
          session_get(at, "active_months"))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Customizing a per-month range and then changing the setup band's "
          "flight dates reconciles correctly: a dropped month's edit "
          "disappears, a surviving month's edit is kept (re-clipped into the "
          "new bounds), and a newly-added month gets a default range -- both "
          "in the pure reconciliation function and end to end through the "
          "real app.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
