"""A row typed by hand into the D2 avails grid must create a targeting group
and PERSIST -- exactly like one seeded from a market or an audience. Reported
from live use (Andrea): typing into a new grid row and moving on lost the
content immediately.

`st.data_editor` can't be driven through AppTest (confirmed: it's a no-op
pass-through under the test harness -- it just returns whatever `data` it was
given, since there's no simulated frontend to submit an edit). So this patches
`streamlit.data_editor` itself, globally, before the script runs: the FIRST
call for the avails grid's key returns `data` with one extra row appended (gid
blank, Audience/Markets/avails filled in -- exactly what the real widget
returns the instant a rep commits a typed row), and every later call echoes
back whatever it's given, the same pass-through AppTest already does by
default -- simulating "nothing more is edited after this."

Two things asserted, matching the report's own repro steps:
    1. The row is reflected in `targeting_groups` (and the flat
       `avails_seed_rows` projection) the SAME run it's typed -- a real group,
       not a blank one, with its Audience/Markets/avails intact.
    2. It survives an UNTOUCHED follow-up rerun byte-for-byte (the same
       regression class `test_group_builder.py` proves for the audience
       finder's own group-building path) -- proving the widget-key bump this
       fix adds (`avails_version` increments and the grid remounts fresh
       whenever the fold-back's row count changes) actually lands the new
       group rather than leaving it trapped behind a stale widget key.

    python tests/test_avails_grid_manual_row.py
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import streamlit as st  # noqa: E402

_ADDED_ROW = {
    "gid": None,
    "Audience": "HH Dwelling Single Family",
    "Markets": ["Washington, DC DMA"],
    "Label": "",
    "Color": None,
    "avails_col": 5000,
}
_injected = {"done": False}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if key.startswith("avails_editor") and not _injected["done"]:
        _injected["done"] = True
        avails_col = [c for c in data.columns if c not in
                      ("gid", "Plan", "Audience", "Markets", "Label", "Color", "Detached")][0]
        row = dict(_ADDED_ROW)
        row[avails_col] = row.pop("avails_col")
        return __import__("pandas").concat(
            [data, __import__("pandas").DataFrame([row])], ignore_index=True)
    return data


st.data_editor = _fake_data_editor

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    from datetime import date
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["target_dmas"] = ["Washington, DC"]
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (this file's own
    # subject) on market+flight -- a truly untouched fresh form no longer
    # reaches it at all.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def real_groups(at):
    groups = at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []
    return [g for g in groups if g["terms"]]


def main():
    print("baseline: avails_version after an ordinary render (market autofill bumps it once)")
    _injected["done"] = True  # suppress injection for this run -- pure pass-through
    baseline_at = new_app()
    baseline_at.run()
    baseline_version = (baseline_at.session_state["avails_version"]
                        if "avails_version" in baseline_at.session_state else 0)
    _injected["done"] = False  # re-arm for the real run below

    print("\ntyping a new row into the D2 grid and moving on")
    at = new_app()
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    check("the injected row was actually offered to the fold-back", _injected["done"])

    groups = real_groups(at)
    check("a real group now exists for the hand-typed row",
          len(groups) == 1, groups)
    if groups:
        g = groups[0]
        check("its audience is exactly what was typed, not blank",
              g["terms"] == ["HH Dwelling Single Family"], g)
        check("its markets are what was picked",
              g.get("geo_def", {}).get("markets") == ["Washington, DC DMA"], g)
        check("its avails figure is what was typed, not 0",
              g.get("avails_monthly") == 5000, g)

    seed_rows = at.session_state["avails_seed_rows"] if "avails_seed_rows" in at.session_state else []
    check("the flat avails_seed_rows projection agrees (one row carries the typed audience)",
          any(r.get("Audience") == "HH Dwelling Single Family" for r in seed_rows), seed_rows)

    print("\nthe widget key was bumped so the grid remounts fresh, not left behind a stale key")
    avails_version = at.session_state["avails_version"] if "avails_version" in at.session_state else 0
    check("avails_version advanced past the ordinary-render baseline "
          "(isolates the fold-back's OWN bump from the pre-existing market-autofill one)",
          avails_version > baseline_version, (baseline_version, avails_version))

    print("\nAN UNTOUCHED FOLLOW-UP RERUN must not lose it (same regression class as the "
          "audience-finder 2-term AND group)")
    before = [dict(g) for g in real_groups(at)]
    at.run()   # nothing further injected -- _fake_data_editor now just echoes back `data`
    check("no exception on the untouched rerun", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    after = real_groups(at)
    check("the hand-typed group survives byte-for-byte", after == before, (before, after))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A row typed directly into the D2 grid creates a real targeting group immediately, "
          "and survives an untouched rerun exactly like one seeded from a market or built "
          "through the audience finder.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
