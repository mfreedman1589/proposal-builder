"""Geography and the media plan Geo column follow the TARGET markets.

    python tests/test_geo_defaults.py

Offline, no DB, no framework. The market rows are built locally from the
checked-in index, so this runs anywhere.

The bug being pinned: `market_label` is the ORIGINATING market -- which
Premion office sold the proposal -- and it used to seed both the Campaign
Specs Geography field and every media plan row's Geo. A campaign sold out of
DC into Denver, Atlanta and Phoenix therefore printed "Washington, DC DMA"
on the plan a client reads. The old rule also ran through `first_line()`,
which is the single-market assumption itself: against three markets it kept
Denver and silently dropped the other two.

Two things must NOT change, and both have their own case below:

  * With no target markets the behaviour is exactly what it always was --
    the originating market label. Every proposal built before target markets
    existed has to keep rendering the way it did.

  * The imported broadcast row keeps the Geo derived from its station's call
    sign, always. A DC-sold Total TV proposal targeting Denver genuinely has
    broadcast running in Washington DC DMA and streaming running in Denver,
    and those two appear on the same plan, differing. That is correct, not a
    bug to be tidied away.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402
import market_profiles as mp  # noqa: E402

ORIGINATING = "Washington, DC DMA"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def labels_for(keys, rows):
    return app.target_market_labels(keys, rows)


def main():
    rows = mp.sort_rows(mp.build_rows())
    # build_rows has no image_path (that's a stored column); the label is all
    # these helpers read, so the local rows are enough.

    print("no target markets -- unchanged from before this existed")
    none_labels = labels_for([], rows)
    check("Geography falls back to the originating market",
          app.geography_default_text(none_labels, ORIGINATING) == ORIGINATING)
    check("Geo column falls back to the originating market",
          app.geo_column_default(none_labels, "", ORIGINATING) == ORIGINATING)
    check("a Geography the rep typed still wins over the fallback",
          app.geo_column_default(none_labels, "Northern Virginia", ORIGINATING)
          == "Northern Virginia")

    print("\none target market")
    one = labels_for(["denver"], rows)
    check("resolves to the market's own label", one == ["Denver"], one)
    check("Geography is that market",
          app.geography_default_text(one, ORIGINATING) == "Denver")
    check("Geo column is that market, not the originating one",
          app.geo_column_default(one, "", ORIGINATING) == "Denver")
    check("the originating market does not leak in",
          ORIGINATING not in app.geo_column_default(one, "", ORIGINATING))

    print("\nseveral target markets")
    three = labels_for(["denver", "atlanta", "phoenix_prescott"], rows)
    check("all three resolve", len(three) == 3, three)
    check("picker order is preserved, not alphabetised",
          three[0] == "Denver" and three[1] == "Atlanta", three)
    geo_text = app.geography_default_text(three, ORIGINATING)
    check("Geography lists one market per line (it's a bullet list)",
          geo_text.count("\n") == 2, repr(geo_text))
    geo_cell = app.geo_column_default(three, geo_text, ORIGINATING)
    check("Geo column joins them on one line (it's a table cell)",
          "\n" not in geo_cell, repr(geo_cell))
    check("every market reaches the Geo cell -- none truncated",
          all(label in geo_cell for label in three), geo_cell)
    # The specific regression: first_line() over a multi-line Geography.
    check("a multi-line Geography is NOT collapsed to its first market",
          geo_cell != app.first_line(geo_text),
          f"first_line would have given {app.first_line(geo_text)!r}")

    print("\nthe rep can still override, per row and per field")
    check("a typed Geography beats the target markets -- the rep's words win",
          app.geo_column_default(three, "Denver metro + suburbs", ORIGINATING)
          == "Denver metro + suburbs")
    check("a typed multi-line Geography is joined, never truncated",
          app.geo_column_default([], "Denver metro\nBoulder", ORIGINATING)
          == "Denver metro, Boulder")
    # resolve_row_defaults is what re-seeds; a dirty row never reaches it, so
    # what matters here is that a CLEAN row takes the new default.
    clean = app.resolve_row_defaults(
        "Premion Streaming TV", geo_cell, "Adults 25-54", "Oct 1 - Dec 31")
    check("a clean streaming row takes the multi-market Geo",
          clean["Geo"] == geo_cell, clean["Geo"])

    print("\nTotal TV: broadcast and streaming Geo differ, correctly")
    # The imported broadcast row carries the Geo derived from its station's
    # call sign. resolve_row_defaults must hold it against any re-seed.
    broadcast_tactic = f"Broadcast Schedule {app.BROADCAST_TACTIC_SUFFIX}" \
        if not app.BROADCAST_TACTIC_MARKER.startswith("Broadcast") \
        else app.BROADCAST_TACTIC_MARKER
    existing_broadcast = {"Tactic": broadcast_tactic,
                          "Geo": "Washington DC DMA",
                          "Targeting": "42 spots across News and Prime"}
    check("the tactic really is recognised as the broadcast line",
          app.is_broadcast_row(existing_broadcast), broadcast_tactic)
    reseeded = app.resolve_row_defaults(
        broadcast_tactic, geo_cell, "Adults 25-54", "Oct 1 - Dec 31",
        current=existing_broadcast)
    check("broadcast keeps its call-sign Geo when the target markets change",
          reseeded["Geo"] == "Washington DC DMA", reseeded["Geo"])
    check("broadcast keeps its schedule-derived Targeting too",
          reseeded["Targeting"] == existing_broadcast["Targeting"],
          reseeded["Targeting"])
    check("...while a streaming row on the SAME plan takes the target markets",
          clean["Geo"] == geo_cell and clean["Geo"] != reseeded["Geo"],
          f"streaming={clean['Geo']!r} broadcast={reseeded['Geo']!r}")
    print(f"        one plan, two Geos: broadcast {reseeded['Geo']!r} / "
          f"streaming {clean['Geo']!r}")

    print("\nGeography autofill only replaces what it wrote")
    class Stub:
        def __init__(self): self.session_state = {}
        def __getattr__(self, n): return lambda *a, **k: None
    real = app.st
    app.st = Stub()
    try:
        app.apply_geography_autofill(ORIGINATING)
        check("applies on an empty form",
              app.st.session_state["geography_text"] == ORIGINATING)
        app.apply_geography_autofill("Denver")
        check("replaces its own earlier value when markets change",
              app.st.session_state["geography_text"] == "Denver")
        app.st.session_state["geography_text"] = "Denver metro, hand typed"
        app.apply_geography_autofill("Atlanta")
        check("never overwrites what the rep typed",
              app.st.session_state["geography_text"] == "Denver metro, hand typed",
              app.st.session_state["geography_text"])
        app.apply_geography_autofill("Phoenix")
        check("and stays out of the way from then on",
              app.st.session_state["geography_text"] == "Denver metro, hand typed")
    finally:
        app.st = real

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Geo defaults follow the target markets; broadcast holds its own.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
