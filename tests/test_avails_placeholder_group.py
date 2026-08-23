"""The persistent blank DC/originating-market row -- a live bug report:
"DC stays as a default empty row in the avails grid even after you upload
an avail."

Root cause: `avails_rows_for_markets([], geo_default, ...)` seeds a
placeholder row (blank Audience, a real Geo, zero avails) whenever no
target market is picked -- the D2 avails table's one blank starter row.
`seed_rows_to_groups` only ever skipped a row with BOTH Audience and Geo
blank, so this row (which has a real Geo) was promoted into a genuine,
permanent targeting group the first time D2 rendered. An avails PDF import
(`_finish_avails_import`) and the Audience finder's "start a new group"
click (`_add_segment_to_group`) both APPEND their real groups to whatever
`targeting_groups` already holds rather than replacing it -- unlike picking
a real target market, which rebuilds the list from scratch and so already
discards the placeholder for free -- so the blank row survived either one
forever, sitting in the D2 table alongside whatever was actually added.

Fix: the placeholder row (and the group it seeds) carries a `_placeholder`
marker (`app.avails_rows_for_markets` -> `targeting_groups.seed_rows_to_groups`
-> the group dict; round-tripped back by `groups_to_seed_rows`; carried
forward by the D2 grid's own fold-back loop in app.py's Section D2, the
same "unchanged cell keeps its old value" test already applied to
Audience/Markets/Label/Color). Both append sites (`_finish_avails_import`,
`_add_segment_to_group`) drop any `_placeholder` group from `existing`
before adding their real ones -- the same replacement a real target-market
pick already gets, applied to the one path that was missing it.

    python tests/test_avails_placeholder_group.py
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
import targeting_groups as tg  # noqa: E402
import avails_pdf_import as api  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StStub:
    """Same swap test_share_of_voice.py's build_option_with_known_lines uses
    to run a pure app.py function (which reads/writes st.session_state)
    outside of a live AppTest script. `rerun` is a no-op here -- state has
    already been mutated by the time these functions call it, and there's no
    real script run for it to restart."""

    def __init__(self, session_state):
        self.session_state = session_state

    def rerun(self):
        pass


def placeholder_group():
    row = {"Audience": "", "Geo": "Washington, DC DMA", app.AVAILS_COLUMN_MONTHLY: 0,
           "_placeholder": True}
    return tg.seed_rows_to_groups([row], app.AVAILS_COLUMN_MONTHLY)[0]


def main():
    print("=" * 78)
    print("SCENARIO  seed_rows_to_groups / groups_to_seed_rows carry the marker")
    print("=" * 78)
    ph_group = placeholder_group()
    check("a _placeholder row becomes a _placeholder group",
          ph_group.get("_placeholder") is True, ph_group)
    check("it still has no real audience (empty terms)", ph_group["terms"] == [], ph_group["terms"])

    real_row = {"Audience": "Homeowners", "Geo": "Denver DMA", app.AVAILS_COLUMN_MONTHLY: 500_000}
    real_group = tg.seed_rows_to_groups([real_row], app.AVAILS_COLUMN_MONTHLY)[0]
    check("an ordinary row's group carries no _placeholder key at all",
          "_placeholder" not in real_group, real_group)

    projected = tg.groups_to_seed_rows([ph_group], app.AVAILS_COLUMN_MONTHLY)[0]
    check("groups_to_seed_rows round-trips the marker onto the projected row",
          projected.get("_placeholder") is True, projected)
    projected_real = tg.groups_to_seed_rows([real_group], app.AVAILS_COLUMN_MONTHLY)[0]
    check("an ordinary group's projected row carries no _placeholder key",
          "_placeholder" not in projected_real, projected_real)

    print("\n" + "=" * 78)
    print("SCENARIO  a target-market row (real, in-progress) is never marked")
    print("=" * 78)
    market_row = {"Audience": "", "Geo": "Baltimore DMA", app.AVAILS_COLUMN_MONTHLY: 0}
    market_group = tg.seed_rows_to_groups([market_row], app.AVAILS_COLUMN_MONTHLY)[0]
    check("a blank-audience row from a REAL market pick (no marker set) stays unmarked",
          "_placeholder" not in market_group, market_group)

    print("\n" + "=" * 78)
    print("SCENARIO  an avails PDF import drops the placeholder, doesn't append past it")
    print("=" * 78)
    document = api.AvailsDocument(
        rfpid="RFPID-TEST", advertiser="Ridgeline Heating & Air", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=1_000_000, attribution_text="",
        groups=[api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_DMA,
                                geo_name="Denver DMA", impressions=1_000_000, row_count=1)],
        source_name="test.pdf")

    real, stub = app.st, _StStub({
        "targeting_groups": [ph_group],
        "active_months": ["September 2026", "October 2026", "November 2026"],
        "flight_start": "2026-09-01", "flight_end": "2026-11-30",
    })
    app.st = stub
    try:
        new_groups, report = app.apply_avails_import(document)
        app._finish_avails_import(document, new_groups, report)
        result = stub.session_state["targeting_groups"]
    finally:
        app.st = real

    check("the placeholder group is gone after the import",
          not any(g.get("_placeholder") for g in result), result)
    check("the real imported group made it in",
          any(tg.audience_label(g) == "Homeowners" for g in result), result)
    check("exactly one group remains (placeholder dropped, one real group added)",
          len(result) == 1, result)

    print("\n" + "=" * 78)
    print("SCENARIO  the Audience finder's \"start a new group\" also drops the placeholder")
    print("=" * 78)
    real, stub = app.st, _StStub({"targeting_groups": [placeholder_group()]})
    app.st = stub
    try:
        app._add_segment_to_group("Auto Intenders", "separate", "Washington, DC DMA")
        result = stub.session_state["targeting_groups"]
    finally:
        app.st = real

    check("the placeholder group is gone after adding a segment from the finder",
          not any(g.get("_placeholder") for g in result), result)
    check("the new segment's group made it in",
          any(tg.audience_label(g) == "Auto Intenders" for g in result), result)
    check("exactly one group remains", len(result) == 1, result)

    print("\n" + "=" * 78)
    print("SCENARIO  the marker survives ordinary reruns of a real form (no import yet)")
    print("=" * 78)
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.run()
    check("no exception on first render", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    for i in range(3):
        groups = list(at.session_state["targeting_groups"] or [])
        check(f"run {i + 1}: exactly one placeholder group, still marked",
              len(groups) == 1 and groups[0].get("_placeholder") is True, groups)
        at.run()

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The D2 avails table's blank starter row stays marked across ordinary reruns, "
          "and both places that ADD a real group (an avails PDF import, the Audience "
          "finder's \"start a new group\") replace it instead of leaving it to sit "
          "alongside whatever was actually added.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
