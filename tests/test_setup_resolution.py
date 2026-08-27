"""FLOW_REWORK_PLAN.md Phase 1: the setup band's data layer --
`app.setup_snapshot` (capture) and `app.resolve_setup` (migration read).

    python tests/test_setup_resolution.py

Offline, no Streamlit runtime, no DB, no PowerPoint -- pure dict fixtures
against the two functions directly. `import app` is safe without a running
Streamlit server (see CLAUDE.md's "Headless testing": the entrypoint's
main() call is guarded by `if __name__ == "__main__"`), same device
tests/test_group_backward_compat.py already uses.

What this proves, once, offline, before any UI exists to exercise it:

1. A proposal with no "setup" key at all (every proposal logged before
   this phase) resolves its five band values from the SAME legacy fields
   `rehydrate_proposal_into_form` already reads today -- a relocation, not
   new inference -- except "avails_mode", which has no legacy field and is
   derived from whether the proposal's own groups carry any avails.
2. A proposal with a real "setup" section is read back verbatim, and wins
   even when it disagrees with what the legacy fields alone would imply
   (a rep can turn avails mode off after the fact; that has to stick).
3. `setup_snapshot` -> `resolve_setup` round-trips exactly for a freshly
   logged proposal -- the property every later phase's migration work
   depends on staying true.

Complements, rather than duplicates, `tests/test_group_backward_compat.py`,
which grounds the same "setup" key in a real generated deck (Scenario 5)
rather than a synthetic fixture.
"""
import os
import sys
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


def main():
    print("a proposal with no setup key at all (pre-Phase-1)")
    empty = app.resolve_setup({})
    check("originating market is None, not a guessed default", empty["originating_market"] is None)
    check("flight dates are None, not a guessed default",
          empty["flight_start"] is None and empty["flight_end"] is None)
    check("plan basis defaults to Monthly", empty["plan_basis"] == app.AVAILS_BASIS_MONTHLY)
    check("avails mode is False with no groups at all", empty["avails_mode"] is False)
    check("total tv is False", empty["total_tv"] is False)

    print("\na pre-Phase-1 proposal shaped like a real logged one")
    legacy_form = {
        "selections": {
            "market": "Harrisburg",
            "products": {"total_tv": True},
        },
        "flight": {"start": "2026-01-01", "end": "2026-03-31",
                   "label": "Jan - Mar", "active_months": ["2026-01", "2026-02", "2026-03"]},
        "avails_basis": app.AVAILS_BASIS_FLIGHT,
    }
    groups_with_avails = [{"id": "g1", "avails_monthly": 45000}, {"id": "g2", "avails_monthly": 0}]
    resolved = app.resolve_setup(legacy_form, groups=groups_with_avails)
    check("originating market relocated from selections.market", resolved["originating_market"] == "Harrisburg")
    check("flight dates relocated from the flight dict",
          resolved["flight_start"] == "2026-01-01" and resolved["flight_end"] == "2026-03-31")
    check("plan basis relocated from avails_basis", resolved["plan_basis"] == app.AVAILS_BASIS_FLIGHT)
    check("total tv relocated from selections.products.total_tv", resolved["total_tv"] is True)
    check("avails mode is True -- at least one group carries avails", resolved["avails_mode"] is True)

    print("\nsame legacy proposal, but none of its groups ever carried avails")
    resolved_no_avails = app.resolve_setup(legacy_form, groups=[{"id": "g1", "avails_monthly": 0}])
    check("avails mode is False when no group has any avails", resolved_no_avails["avails_mode"] is False)
    resolved_no_groups = app.resolve_setup(legacy_form, groups=None)
    check("avails mode is False when there are no groups to check at all",
          resolved_no_groups["avails_mode"] is False)

    print("\nan invalid stored avails_basis falls back to Monthly rather than propagating garbage")
    garbage_basis = dict(legacy_form, avails_basis="Quarterly")
    check("unrecognized basis string falls back to Monthly",
          app.resolve_setup(garbage_basis)["plan_basis"] == app.AVAILS_BASIS_MONTHLY)

    print("\nrow-level market is the fallback when selections.market is missing")
    no_selections_market = {"selections": {}, "flight": {}, "avails_basis": app.AVAILS_BASIS_MONTHLY}
    check("falls back to row_market", app.resolve_setup(
        no_selections_market, row_market="DC")["originating_market"] == "DC")
    check("an unrecognized market string is dropped, not passed through",
          app.resolve_setup(no_selections_market, row_market="Denver")["originating_market"] is None)

    print("\na post-Phase-1 proposal's stored setup section wins, even when it disagrees with the legacy fields")
    disagreeing_form = dict(legacy_form)
    disagreeing_form["setup"] = {
        "originating_market": "DC",              # legacy selections.market says Harrisburg
        "flight_start": "2026-04-01",             # legacy flight dict says 2026-01-01
        "flight_end": "2026-06-30",
        "plan_basis": app.AVAILS_BASIS_MONTHLY,   # legacy avails_basis says Full flight
        "avails_mode": False,                     # groups below still carry avails
        "total_tv": False,                        # legacy products.total_tv says True
    }
    resolved_stored = app.resolve_setup(disagreeing_form, groups=groups_with_avails)
    check("stored setup wins on every field, not the legacy ones underneath it",
          resolved_stored == disagreeing_form["setup"])

    print("\nsetup_snapshot -> resolve_setup round-trips for a freshly logged proposal")
    # Every one of the five now has a real widget behind it -- setup_snapshot
    # takes avails_mode straight from that widget, not inferred from groups
    # (the band's own checkbox superseded the group-inference stand-in this
    # commit's earlier draft used before the band existed).
    import datetime
    snapshot = app.setup_snapshot(
        "DC", datetime.date(2026, 7, 1), datetime.date(2026, 9, 30),
        app.AVAILS_BASIS_MONTHLY, True, True)
    fresh_form = {"setup": snapshot}
    check("resolving a freshly-snapshotted form recovers it exactly",
          app.resolve_setup(fresh_form) == snapshot)

    print("\nsetup_snapshot with avails mode and Total TV both off")
    off_snapshot = app.setup_snapshot(
        "Harrisburg", datetime.date(2026, 1, 1), datetime.date(2026, 1, 31),
        app.AVAILS_BASIS_MONTHLY, False, False)
    check("avails mode is False", off_snapshot["avails_mode"] is False)
    check("total tv is False", off_snapshot["total_tv"] is False)

    # ------------------------------------------------------------------
    # FLOW_REWORK_PLAN.md Phase 2: per-month flight ranges
    # ------------------------------------------------------------------
    D = datetime.date

    print("\nflight_month_ranges with no stored ranges: every month defaults to its own full "
          "calendar span, clipped by the flight")
    fresh = app.flight_month_ranges(D(2026, 9, 21), D(2026, 12, 14))
    check("three months", [r["month"] for r in fresh] == ["Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026"])
    check("September clipped to the flight's own start",
          fresh[0]["start"] == D(2026, 9, 21) and fresh[0]["end"] == D(2026, 9, 30))
    check("October and November are whole months",
          fresh[1]["start"] == D(2026, 10, 1) and fresh[1]["end"] == D(2026, 10, 31)
          and fresh[2]["start"] == D(2026, 11, 1) and fresh[2]["end"] == D(2026, 11, 30))
    check("December clipped to the flight's own end",
          fresh[3]["start"] == D(2026, 12, 1) and fresh[3]["end"] == D(2026, 12, 14))
    check("every month starts active", all(r["active"] for r in fresh))

    print("\na stored range survives a flight that still covers it")
    kept = app.flight_month_ranges(D(2026, 9, 21), D(2026, 12, 14), stored=[
        {"month": "Oct 2026", "start": D(2026, 10, 8), "end": D(2026, 10, 20), "active": True},
        {"month": "Nov 2026", "start": D(2026, 11, 1), "end": D(2026, 11, 30), "active": False},
    ])
    october = next(r for r in kept if r["month"] == "Oct 2026")
    november = next(r for r in kept if r["month"] == "Nov 2026")
    check("October's hand-edited range survives",
          october["start"] == D(2026, 10, 8) and october["end"] == D(2026, 10, 20))
    check("November's skip survives", november["active"] is False)
    check("September and December, never stored, still default",
          next(r for r in kept if r["month"] == "Sep 2026")["start"] == D(2026, 9, 21))

    print("\na stored range is CLAMPED, not dropped, once the flight narrows past it")
    clamped = app.flight_month_ranges(D(2026, 10, 1), D(2026, 10, 31), stored=[
        {"month": "Oct 2026", "start": D(2026, 9, 15), "end": D(2026, 11, 15), "active": True},
    ])
    check("one month, clamped into the new flight's own bounds",
          len(clamped) == 1 and clamped[0]["start"] == D(2026, 10, 1) and clamped[0]["end"] == D(2026, 10, 31))

    print("\nan empty intersection (every stored month left the range) resets to all-default, not empty")
    reset = app.flight_month_ranges(D(2026, 1, 1), D(2026, 1, 31), stored=[
        {"month": "Sep 2026", "start": D(2026, 9, 21), "end": D(2026, 9, 30), "active": True},
    ])
    check("one month, active, defaulted -- not the stale September entry",
          len(reset) == 1 and reset[0]["month"] == "Jan 2026" and reset[0]["active"] is True)

    print("\na stored list that deselects every month is treated as belonging to another flight")
    all_off = app.flight_month_ranges(D(2026, 9, 1), D(2026, 10, 31), stored=[
        {"month": "Sep 2026", "start": D(2026, 9, 1), "end": D(2026, 9, 30), "active": False},
        {"month": "Oct 2026", "start": D(2026, 10, 1), "end": D(2026, 10, 31), "active": False},
    ])
    check("resets to all-active rather than leaving nothing active",
          all(r["active"] for r in all_off))

    print("\nactive_month_labels / flight_active_days")
    check("active_month_labels lists only the active months, in order",
          app.active_month_labels(kept) == ["Sep 2026", "Oct 2026", "Dec 2026"])
    check("flight_active_days sums only active months' own day counts "
          "(Sep 10 + Oct 13 [8-20] + Dec 14, November skipped)",
          app.flight_active_days(kept) == 10 + 13 + 14)

    print("\nranges_are_customized")
    check("the untouched default is never flagged as customized",
          app.ranges_are_customized(fresh, D(2026, 9, 21), D(2026, 12, 14)) is False)
    check("a skipped month IS customized",
          app.ranges_are_customized(kept, D(2026, 9, 21), D(2026, 12, 14)) is True)

    print("\nflight_months_snapshot -> resolve_flight_months round-trips for a fresh Phase-2 proposal")
    snap = app.flight_months_snapshot(fresh)
    check("dates serialize to ISO strings", snap[0]["start"] == "2026-09-21")
    phase2_form = {"flight": {"start": "2026-09-21", "end": "2026-12-14", "month_ranges": snap}}
    revived = app.resolve_flight_months(phase2_form)
    check("round-trips to the same active labels and day count",
          app.active_month_labels(revived) == app.active_month_labels(fresh)
          and app.flight_active_days(revived) == app.flight_active_days(fresh))

    print("\nresolve_flight_months migrates a pre-Phase-2 proposal from its whole-month active_months")
    pre_phase2 = {"flight": {"start": "2026-01-01", "end": "2026-03-31",
                              "active_months": ["Jan 2026", "Mar 2026"]}}
    migrated = app.resolve_flight_months(pre_phase2)
    check("three months, February inferred as skipped from the stored whole-month list",
          app.active_month_labels(migrated) == ["Jan 2026", "Mar 2026"])
    check("February's range still defaults even though it's inactive",
          next(r for r in migrated if r["month"] == "Feb 2026")["start"] == D(2026, 2, 1))

    print("\nresolve_flight_months tolerates an unparseable stored active_months (a different vintage's "
          "label format), falling through to all-active rather than raising")
    garbage_labels = {"flight": {"start": "2026-01-01", "end": "2026-03-31",
                                  "active_months": ["2026-01", "2026-02", "2026-03"]}}
    check("falls through to all-active",
          app.active_month_labels(app.resolve_flight_months(garbage_labels))
          == ["Jan 2026", "Feb 2026", "Mar 2026"])

    print("\nresolve_flight_months with no flight dates at all")
    check("empty list, not an exception", app.resolve_flight_months({}) == [])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("resolve_setup relocates four existing fields and infers only avails_mode; "
          "a stored setup section always wins; setup_snapshot round-trips exactly.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
