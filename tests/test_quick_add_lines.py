"""The Add-lines panel builds plan rows from the avails table.

    python tests/test_quick_add_lines.py

Offline for the helpers; the panel itself is driven through the real form,
because what matters is that clicking Add puts the right rows on the right
option with the right dirty flags -- none of which the pure function can show.

The default is the CROSS PRODUCT as separate lines: one product across three
markets is three lines, two products across three markets is six. A market is
a line a client can see and a rep can move money between, so that is the shape
plans are actually built in. Combining onto one line is the exception, and has
its own toggle -- the same default and exception as the avails table.
"""
import os
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

FLIGHT = "Oct 1 - Dec 31"
STACK = "Adults 25-54, homeowners"
ORIGINATING = "Washington, DC DMA"
THREE = ["Denver", "Atlanta", "Phoenix"]

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def main():
    catalog = app._quick_add_catalog()
    labels = {label: cpm for label, cpm in catalog.values()}
    premion = "Premion Streaming TV"

    print("the catalog the panel offers")
    check("every product and sports package is offered",
          len(catalog) >= 30, len(catalog))
    check("it resolves through line_product_spec, so rates match the card",
          labels.get(premion) == app.line_product_spec("premion_streaming_tv")[1],
          labels.get(premion))
    check("the non-line product is excluded",
          not any(k in app.NON_LINE_PRODUCT_KEYS for k in catalog),
          [k for k in catalog if k in app.NON_LINE_PRODUCT_KEYS])

    print("\ncross product, as separate lines (the default)")
    rows = app.quick_add_rows([premion], [], THREE, FLIGHT, STACK, ORIGINATING)
    check("one product x three markets is three lines", len(rows) == 3, len(rows))
    check("one per market, in the order picked",
          [r["Geo"] for r in rows] == THREE, [r["Geo"] for r in rows])

    two = [premion, "Streaming Retargeting - Display"]
    rows = app.quick_add_rows(two, [], THREE, FLIGHT, STACK, ORIGINATING)
    check("two products x three markets is six lines", len(rows) == 6, len(rows))

    rows = app.quick_add_rows(two, ["Aud A", "Aud B"], THREE, FLIGHT, STACK, ORIGINATING)
    check("audiences multiply too: 2 x 2 x 3 is twelve", len(rows) == 12, len(rows))

    print("\ncombining onto one line (the exception)")
    rows = app.quick_add_rows([premion], [], THREE, FLIGHT, STACK, ORIGINATING,
                              combine=True)
    check("a single line", len(rows) == 1, len(rows))
    check("naming every market", all(m in rows[0]["Geo"] for m in THREE),
          rows[0]["Geo"])

    print("\ngenerated lines follow the existing rules")
    rows = app.quick_add_rows([premion], [], ["Denver"], FLIGHT, STACK, ORIGINATING)
    check("CPM comes from the rate card, not zero",
          rows[0]["CPM"] == labels[premion] and rows[0]["CPM"] > 0, rows[0]["CPM"])
    check("flight is the form's flight", rows[0]["Flight"] == FLIGHT)
    check("a rate row, with impressions and cost left for the rep",
          rows[0]["Type"] == app.ROW_TYPE_RATE
          and rows[0]["Impressions"] == 0.0 and rows[0]["Cost"] == 0.0, rows[0])

    # A product with its own targeting_copy must keep it -- the panel must not
    # hand it the audience stack. Same rule resolve_row_defaults enforces
    # everywhere else.
    retarget = "Streaming Retargeting - Display"
    rows = app.quick_add_rows([retarget], ["Outdoor enthusiasts"], ["Denver"],
                              FLIGHT, STACK, ORIGINATING)
    expected = app.resolve_row_defaults(retarget, "Denver", STACK, FLIGHT)["Targeting"]
    check("a product with its own targeting copy keeps it",
          rows[0]["Targeting"] == expected, rows[0]["Targeting"])
    check("...and that copy is not the picked audience",
          rows[0]["Targeting"] != "Outdoor enthusiasts", rows[0]["Targeting"])

    # A picked audience DOES reach an ordinary product's Targeting.
    rows = app.quick_add_rows([premion], ["Outdoor enthusiasts"], ["Denver"],
                              FLIGHT, STACK, ORIGINATING)
    check("an ordinary product takes the picked audience",
          rows[0]["Targeting"] == "Outdoor enthusiasts", rows[0]["Targeting"])
    rows = app.quick_add_rows([premion], [], ["Denver"], FLIGHT, STACK, ORIGINATING)
    check("with no audience picked it falls back to the Campaign Specs stack",
          rows[0]["Targeting"] == STACK, rows[0]["Targeting"])

    print("\nedges")
    check("no products means no lines -- a line has to be something",
          app.quick_add_rows([], ["A"], THREE, FLIGHT, STACK, ORIGINATING) == [])
    rows = app.quick_add_rows([premion], [], [], FLIGHT, STACK, ORIGINATING)
    check("no geo picked falls back to the form's geo default",
          len(rows) == 1 and rows[0]["Geo"] == ORIGINATING, rows)
    rows = app.quick_add_rows(["Something bespoke"], [], ["Denver"],
                              FLIGHT, STACK, ORIGINATING)
    check("free text is allowed and prices at 0 for the rep to fill in",
          len(rows) == 1 and rows[0]["CPM"] == 0.0 and rows[0]["Tactic"] == "Something bespoke",
          rows)

    print("\nthrough the real form")
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    root = Path(__file__).resolve().parent.parent

    def add(products, auds, geos, combine=False, existing=None):
        at = AppTest.from_file(str(root / "app.py"), default_timeout=300)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        at.session_state["include_avails_template"] = True
        # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (this file's
        # own subject) on market+flight.
        at.session_state["market_choice"] = "DC"
        at.session_state["flight_start"] = date(2026, 9, 1)
        at.session_state["flight_end"] = date(2026, 11, 30)
        at.session_state["avails_seed_rows"] = [
            {"Audience": "Outdoor enthusiasts", "Geo": "Denver",
             app.AVAILS_COLUMN_MONTHLY: 4_200_000},
            {"Audience": "", "Geo": "Atlanta", app.AVAILS_COLUMN_MONTHLY: 0},
        ]
        at.run()
        before = len(at.session_state["plan_options"][0]["rows"])
        at.session_state["qa_prod_0_0"] = products
        at.session_state["qa_aud_0_0"] = auds
        at.session_state["qa_geo_0_0"] = geos
        if combine:
            at.session_state["qa_combine_0_0"] = True
        at.run()
        buttons = [b for b in at.button if b.label == "Add lines"]
        if not buttons:
            return None, "no Add lines button"
        buttons[0].click().run()
        if at.exception:
            return None, at.exception[0].message[:200]
        option = at.session_state["plan_options"][0]
        return (option["rows"][before:], option["dirty"][before:]), None

    got, err = add([premion], [], ["Denver", "Atlanta"])
    check("clicking Add appends one line per market", not err and got and len(got[0]) == 2,
          err or (got and len(got[0])))
    if got:
        check("added lines are marked dirty, like a duplicated line",
              all(got[1]), got[1])
        check("and carry the rate-card CPM",
              all(r["CPM"] == labels[premion] for r in got[0]),
              [r["CPM"] for r in got[0]])

    got, err = add([premion], [], ["Denver", "Atlanta"], combine=True)
    check("the combine toggle produces one line",
          not err and got and len(got[0]) == 1, err or (got and len(got[0])))

    # The panel reads the avails TABLE, which includes market rows that have no
    # audience yet -- those are exactly what a freshly seeded table is full of.
    at = AppTest.from_file(str(root / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.session_state["avails_seed_rows"] = [
        {"Audience": "", "Geo": "Phoenix", app.AVAILS_COLUMN_MONTHLY: 0}]
    at.run()
    geo_opts = [m.options for m in at.multiselect if m.label == "Geo"]
    check("a market with no audience yet still appears in the Geo list",
          geo_opts and "Phoenix" in geo_opts[0], geo_opts)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Lines are built from the avails table, not retyped.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
