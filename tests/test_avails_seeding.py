"""Target markets seed the avails table, one row per market.

    python tests/test_avails_seeding.py

Offline, no DB. The pure helpers are asserted directly; the clean/dirty
behaviour is driven through the real form with AppTest, because that is the
part that has bitten this project repeatedly -- a rerun path silently
re-seeding rows a rep had already filled in.

The rule being pinned: market rows REPLACE the originating-market default
rather than joining it. A leftover "Washington, DC DMA" row under three
target markets is not a starting point, it is a zero-avails row that reaches
the targeting slide reading as a market with no inventory.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

ORIGINATING = "Washington, DC DMA"
COL = app.AVAILS_COLUMN_MONTHLY

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def geos(rows):
    return [r["Geo"] for r in rows]


def main():
    print("no markets -- unchanged from before this existed")
    rows = app.avails_rows_for_markets([], ORIGINATING)
    check("one row", len(rows) == 1, rows)
    check("carrying the originating market", geos(rows) == [ORIGINATING], geos(rows))
    check("blank audience and zero avails, as the table has always started",
          rows[0]["Audience"] == "" and rows[0][COL] == 0, rows[0])

    print("\none market")
    rows = app.avails_rows_for_markets(["Denver"], ORIGINATING)
    check("one row", len(rows) == 1, rows)
    check("carrying that market, not the originating one",
          geos(rows) == ["Denver"], geos(rows))
    check("the originating market is gone, not sitting alongside",
          ORIGINATING not in geos(rows), geos(rows))

    print("\nthree markets, broken out separately (the default)")
    three = ["Denver", "Atlanta", "Phoenix"]
    rows = app.avails_rows_for_markets(three, ORIGINATING)
    check("one row per market", len(rows) == 3, len(rows))
    check("in picker order", geos(rows) == three, geos(rows))
    check("no originating-market row left over",
          ORIGINATING not in geos(rows), geos(rows))
    check("each row is independently fillable",
          all(r["Audience"] == "" and r[COL] == 0 for r in rows), rows)

    print("\nthree markets, combined into one row")
    rows = app.avails_rows_for_markets(three, ORIGINATING, combine=True)
    check("a single row", len(rows) == 1, len(rows))
    check("naming every market", geos(rows) == [", ".join(three)], geos(rows))
    check("no market is dropped in the combine",
          all(m in rows[0]["Geo"] for m in three), rows[0]["Geo"])

    print("\nthe autofill only replaces what it wrote")

    class Stub:
        def __init__(self): self.session_state = {}
        def __getattr__(self, n): return lambda *a, **k: None

    real = app.st
    app.st = Stub()
    try:
        seeded = app.apply_avails_autofill(
            app.avails_rows_for_markets(["Denver"], ORIGINATING))
        check("seeds an empty form", seeded
              and geos(app.st.session_state["avails_seed_rows"]) == ["Denver"])

        changed = app.apply_avails_autofill(
            app.avails_rows_for_markets(three, ORIGINATING))
        check("follows a market change while untouched",
              changed and geos(app.st.session_state["avails_seed_rows"]) == three,
              geos(app.st.session_state["avails_seed_rows"]))

        again = app.apply_avails_autofill(
            app.avails_rows_for_markets(three, ORIGINATING))
        check("does nothing when already correct (no editor churn)", not again)

        # The rep fills in a real avails figure they went and looked up.
        app.st.session_state["avails_seed_rows"][0]["Audience"] = "Outdoor enthusiasts"
        app.st.session_state["avails_seed_rows"][0][COL] = 4_200_000
        touched = app.apply_avails_autofill(
            app.avails_rows_for_markets(["Denver", "Atlanta"], ORIGINATING))
        kept = app.st.session_state["avails_seed_rows"]
        check("never overwrites an edited row, even on a market change",
              not touched and len(kept) == 3
              and kept[0][COL] == 4_200_000, kept)
        check("and stays out of the way from then on",
              not app.apply_avails_autofill(
                  app.avails_rows_for_markets(["Boise"], ORIGINATING)))
    finally:
        app.st = real

    print("\nthrough the real form")
    import os
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest

    def run(picks, combine=None, edit=None):
        at = AppTest.from_file(str(Path(__file__).resolve().parent.parent / "app.py"),
                               default_timeout=240)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        at.session_state["include_avails_template"] = True
        if picks:
            at.session_state["target_dma_choice"] = picks
        if combine is not None:
            at.session_state["avails_combine_markets"] = combine
        if edit is not None:
            at.session_state["avails_seed_rows"] = edit
        at.run()
        if at.exception:
            return None, at.exception[0].message[:200]
        return at.session_state["avails_seed_rows"], None

    rows, err = run([])
    check("no markets: one originating row", not err and geos(rows) == [ORIGINATING],
          err or geos(rows))
    rows, err = run(["Denver"])
    check("one market: one row for it", not err and geos(rows) == ["Denver"],
          err or geos(rows))
    rows, err = run(["Denver", "Atlanta", "Phoenix"])
    check("three markets: three rows, no originating row",
          not err and geos(rows) == three, err or geos(rows))
    rows, err = run(["Denver", "Atlanta", "Phoenix"], combine=True)
    check("three markets combined: one row naming all three",
          not err and len(rows) == 1 and all(m in rows[0]["Geo"] for m in three),
          err or geos(rows))

    # An edited row must survive a market change made afterwards.
    edited = [{"Audience": "Outdoor enthusiasts", "Geo": "Denver", COL: 4_200_000}]
    rows, err = run(["Denver", "Atlanta"], edit=edited)
    check("an edited row survives a market change untouched",
          not err and rows == edited, err or rows)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Markets seed the avails table; edits are never re-seeded.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
