"""FLOW_REWORK_PLAN.md Phase 6: the flight control mirrored near the media
plan (Section E) -- one owner (the setup band's own flight_start/flight_end/
custom_flighting/flight_months session_state keys), two render sites. A rep
who scrolls to the bottom, sees the plan, and realizes the flight is wrong
no longer has to scroll back to the top: the mirror edits the SAME keys,
through the same queue-then-apply shape (`_pending_flight_mirror_edit` /
apply_pending_flight_mirror_edit) BAND_WIDGET_KEYS' own client_name/total_tv
fix uses, since flight_start/flight_end/custom_flighting are all band widget
keys that have already instantiated above by the time Section E renders.

    python tests/test_flight_mirror.py
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


def sget(at, key, default=None):
    try:
        return at.session_state[key]
    except (KeyError, AttributeError):
        return default


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def mirror_date_input(at, prefix):
    # flight_start_e/flight_end_e are keyed on flight_mirror_gen (a resync
    # signature over (flight_start, flight_end, custom_flighting), separate
    # from flight_months_gen -- toggling custom_flighting alone doesn't bump
    # flight_months_gen, so that alone isn't enough to keep this widget from
    # going stale. See app.py's own comment at the mirror's definition.
    gen = sget(at, "flight_mirror_gen", 0)
    matches = [w for w in at.date_input if str(w.key) == f"{prefix}_{gen}"]
    return matches[0] if matches else None


def by_month(ranges):
    return {r["month"]: r for r in (ranges or [])}


def test_mirror_reflects_band_edit():
    print("\nediting the BAND's own flight dates shows up in the mirror "
          "(collapsed label and, once opened, its own date_input values)")
    at = new_app()
    at.run()
    check("no exception on initial render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")

    expanders = [e for e in at.expander if str(e.label).startswith("📅 Flight:")]
    check("the mirror expander is present", len(expanders) == 1, [e.label for e in at.expander])
    check("collapsed label shows the band's own flight",
          "Sep 2026" in expanders[0].label and "Nov 2026" in expanders[0].label,
          expanders[0].label)

    mirror_start = mirror_date_input(at, "flight_start_e")
    mirror_end = mirror_date_input(at, "flight_end_e")
    check("mirror start matches the band", mirror_start is not None and mirror_start.value == date(2026, 9, 1),
          mirror_start.value if mirror_start else None)
    check("mirror end matches the band", mirror_end is not None and mirror_end.value == date(2026, 11, 30),
          mirror_end.value if mirror_end else None)

    band_start = [w for w in at.date_input if str(w.key) == "flight_start"][0]
    band_start.set_value(date(2026, 10, 1)).run()
    check("no exception after a band edit", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    expanders2 = [e for e in at.expander if str(e.label).startswith("📅 Flight:")]
    check("mirror's collapsed label follows the band's new flight",
          "Oct 2026" in expanders2[0].label, expanders2[0].label)
    mirror_start2 = mirror_date_input(at, "flight_start_e")
    check("mirror's own date_input follows too", mirror_start2 is not None
          and mirror_start2.value == date(2026, 10, 1), mirror_start2.value if mirror_start2 else None)


def test_mirror_edit_updates_band():
    print("\nediting the MIRROR's flight end updates the band's own "
          "flight_end (queued, applied before the band's widget renders "
          "next run -- no instantiation-order exception)")
    at = new_app()
    at.run()
    mirror_end = mirror_date_input(at, "flight_end_e")
    mirror_end.set_value(date(2026, 12, 31)).run()
    check("no exception after a mirror edit", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("the band's own flight_end followed the mirror's edit",
          sget(at, "flight_end") == date(2026, 12, 31), sget(at, "flight_end"))
    band_end = [w for w in at.date_input if str(w.key) == "flight_end"][0]
    check("the band's own widget shows the new value too",
          band_end.value == date(2026, 12, 31), band_end.value)
    check("active_months grew to include December",
          "Dec 2026" in (sget(at, "active_months") or []), sget(at, "active_months"))


def test_mirror_custom_flighting_toggle_and_row_edit():
    print("\nticking Custom flighting in the mirror updates the band's own "
          "checkbox, and editing a per-month row in the mirror updates "
          "flight_months -- reflected in the BAND's own per-month row too, "
          "not just the mirror's")
    at = new_app()
    at.run()
    mirror_custom = [w for w in at.checkbox if str(w.key).startswith("custom_flighting_e_")][0]
    mirror_custom.set_value(True).run()
    check("no exception after ticking Custom flighting in the mirror", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("the band's own custom_flighting followed", sget(at, "custom_flighting") is True,
          sget(at, "custom_flighting"))
    band_custom = [w for w in at.checkbox if str(w.key) == "custom_flighting"][0]
    check("the band's own checkbox shows True too", band_custom.value is True, band_custom.value)

    fm = sget(at, "flight_months")
    sep_idx = next(i for i, r in enumerate(fm) if r["month"].startswith("Sep"))
    mirror_gen = sget(at, "flight_months_gen", 0)
    mirror_row_start = [w for w in at.date_input if str(w.key) == f"fm_start_e_{mirror_gen}_{sep_idx}"][0]
    mirror_row_start.set_value(date(2026, 9, 10)).run()
    check("no exception after editing a mirror per-month row", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("flight_months picked up the mirror's per-row edit",
          by_month(sget(at, "flight_months"))["Sep 2026"]["start"] == date(2026, 9, 10),
          sget(at, "flight_months"))

    band_gen = sget(at, "flight_months_gen", 0)
    band_row_start = [w for w in at.date_input if str(w.key) == f"fm_start_{band_gen}_{sep_idx}"]
    check("the BAND's own per-month row (different key namespace) shows the "
          "same edit -- not stale, per the 'session_state beats value=' rule",
          band_row_start and band_row_start[0].value == date(2026, 9, 10),
          band_row_start[0].value if band_row_start else "<not found>")


def test_no_second_source_of_truth():
    print("\nthe mirror never diverges from the band across a redraft-shaped "
          "sequence: band edit, then mirror edit, then band edit again")
    at = new_app()
    at.run()
    band_start = [w for w in at.date_input if str(w.key) == "flight_start"][0]
    band_start.set_value(date(2026, 8, 1)).run()
    mirror_end = mirror_date_input(at, "flight_end_e")
    mirror_end.set_value(date(2026, 10, 31)).run()
    band_start2 = [w for w in at.date_input if str(w.key) == "flight_start"][0]
    band_start2.set_value(date(2026, 8, 15)).run()
    check("no exception across the interleaved sequence", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("final flight_start is the band's last edit",
          sget(at, "flight_start") == date(2026, 8, 15), sget(at, "flight_start"))
    check("final flight_end is the mirror's edit, carried through the "
          "later band edit untouched",
          sget(at, "flight_end") == date(2026, 10, 31), sget(at, "flight_end"))
    mirror_start_final = mirror_date_input(at, "flight_start_e")
    check("the mirror's own widget agrees with the band -- one value, not two",
          mirror_start_final is not None and mirror_start_final.value == date(2026, 8, 15),
          mirror_start_final.value if mirror_start_final else None)


def main():
    test_mirror_reflects_band_edit()
    test_mirror_edit_updates_band()
    test_mirror_custom_flighting_toggle_and_row_edit()
    test_no_second_source_of_truth()
    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The flight mirror near the media plan and the setup band's own "
          "flight controls are two render sites over the identical "
          "session_state keys -- an edit at either site is visible at both, "
          "in both directions, including per-month custom flighting rows.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
