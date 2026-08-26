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
    import datetime
    snapshot = app.setup_snapshot(
        "DC", datetime.date(2026, 7, 1), datetime.date(2026, 9, 30),
        app.AVAILS_BASIS_MONTHLY, True, [{"id": "g1", "avails_monthly": 12000}])
    fresh_form = {"setup": snapshot}
    check("resolving a freshly-snapshotted form recovers it exactly",
          app.resolve_setup(fresh_form, groups=[{"id": "g1", "avails_monthly": 12000}]) == snapshot)

    print("\nsetup_snapshot with no avails anywhere and Total TV off")
    off_snapshot = app.setup_snapshot(
        "Harrisburg", datetime.date(2026, 1, 1), datetime.date(2026, 1, 31),
        app.AVAILS_BASIS_MONTHLY, False, [{"id": "g1", "avails_monthly": 0}])
    check("avails mode is False", off_snapshot["avails_mode"] is False)
    check("total tv is False", off_snapshot["total_tv"] is False)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("resolve_setup relocates four existing fields and infers only avails_mode; "
          "a stored setup section always wins; setup_snapshot round-trips exactly.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
