"""The targeting group data model: expression parsing/rendering, and the
backward-compatibility projection.

    python tests/test_targeting_groups.py

Offline, no Streamlit, no DB, no framework. The central thing this file
proves is the property everything else depends on: an existing flat-avails
proposal loads into groups and projects back to BYTE-IDENTICAL flat rows,
including the adversarial cases (a hand-typed audience with a comma in it, a
combined multi-market Geo string, a blank table). That property is what makes
"existing proposals load, rebuild and render identically" true without a
migration script -- there's nothing to migrate, only a projection that has to
be exact.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import targeting_groups as tg  # noqa: E402

COL = "Max Monthly Avails"   # the real constant lives in app.py; passed explicitly
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def main():
    print("expression parsing and rendering round-trip")
    g1 = tg.new_group(["DEMO Homeowner"])
    check("a single term has no operator", g1["op"] is None)
    check("expression_text of one term is unparenthesized",
          tg.expression_text(g1) == "DEMO Homeowner")
    check("audience_label of one term is the term itself",
          tg.audience_label(g1) == "DEMO Homeowner")

    g2 = tg.new_group(["DEMO Homeowner", "HH Income 150K Plus"], op="AND")
    check("expression_text is the canonical parenthesized AND form",
          tg.expression_text(g2) == "(DEMO Homeowner) AND (HH Income 150K Plus)")
    check("audience_label reads as plain language, comma-joined for AND",
          tg.audience_label(g2) == "DEMO Homeowner, HH Income 150K Plus")

    g3 = tg.new_group(["College Planning Parents", "Education Services"], op="OR")
    check("expression_text is the canonical parenthesized OR form",
          tg.expression_text(g3) == "(College Planning Parents) OR (Education Services)")
    check("audience_label reads 'or' in plain language",
          tg.audience_label(g3) == "College Planning Parents or Education Services")

    check("parse_expression round-trips a 2-term AND expression",
          tg.parse_expression(tg.expression_text(g2)) == (g2["terms"], "AND"))
    check("parse_expression round-trips a 2-term OR expression",
          tg.parse_expression(tg.expression_text(g3)) == (g3["terms"], "OR"))
    g4 = tg.new_group(["A", "B", "C"], op="AND")
    check("parse_expression handles 3+ terms",
          tg.parse_expression(tg.expression_text(g4)) == (["A", "B", "C"], "AND"))

    print("\nparse_expression is deliberately one-way and conservative")
    check("a bare single term is NOT parsed as an expression (no operator)",
          tg.parse_expression("DEMO Homeowner") is None)
    check("a lone parenthesized term with no operator is NOT parsed",
          tg.parse_expression("(DEMO Homeowner)") is None)
    check("mixed AND/OR is refused",
          tg.parse_expression("(A) AND (B) OR (C)") is None)
    check("a hand-typed audience merely containing a paren doesn't mis-parse",
          tg.parse_expression("(call first) then homeowners") is None)
    check("a hand-typed audience containing the word 'and' doesn't mis-parse",
          tg.parse_expression("Homeowners and higher income") is None)
    check("garbage trailing an otherwise-valid expression is refused",
          tg.parse_expression("(A) AND (B) garbage") is None)

    print("\ngeo_label")
    gm = tg.new_group(["X"], geo_def={"kind": "markets", "markets": ["denver", "atlanta"]})
    check("markets kind joins keys when no label lookup is given",
          tg.geo_label(gm) == "denver, atlanta")
    check("markets kind uses the label lookup when given",
          tg.geo_label(gm, label_for=lambda k: k.title()) == "Denver, Atlanta")
    gt = tg.new_group(["X"], geo_def={"kind": "text", "label": "Washington, DC DMA"})
    check("text kind is the stored label verbatim",
          tg.geo_label(gt) == "Washington, DC DMA")
    gz = tg.new_group(["X"], geo_def={"kind": "zips", "zips": ["20005", "20006"]})
    check("zips kind joins the zip list", tg.geo_label(gz) == "20005, 20006")

    print("\nthe backward-compatibility round trip -- the acceptance gate")
    adversarial_row_sets = [
        [],
        [{"Audience": "", "Geo": "", COL: 0}],
        [{"Audience": "DEMO Homeowner", "Geo": "Denver", COL: 400000}],
        # A hand-typed audience containing a comma must stay ONE term.
        [{"Audience": "Homeowners, higher income", "Geo": "Denver", COL: 100000}],
        # A combined multi-market row (avails_rows_for_markets(combine=True)).
        [{"Audience": "", "Geo": "Denver, Atlanta, Phoenix", COL: 900000}],
        # Multiple rows, some blank-audience (market-seeded default rows).
        [{"Audience": "", "Geo": "Denver", COL: 0},
         {"Audience": "", "Geo": "Atlanta", COL: 0},
         {"Audience": "DEMO Homeowner", "Geo": "Denver", COL: 500000}],
        # An audience string that happens to look parenthesized but isn't a
        # valid two-term expression (must not be split).
        [{"Audience": "(call first) then decide", "Geo": "Harrisburg", COL: 0}],
    ]
    for i, rows in enumerate(adversarial_row_sets):
        groups = tg.seed_rows_to_groups(rows, COL)
        back = tg.groups_to_seed_rows(groups, COL)
        expected = [r for r in rows
                   if str(r.get("Audience", "")).strip() or str(r.get("Geo", "")).strip()]
        check(f"round trip #{i}: {rows!r}", back == expected,
              f"got {back!r}")

    print("\nre-deriving groups from an unchanged table keeps ids stable")
    rows = [{"Audience": "DEMO Homeowner", "Geo": "Denver", COL: 500000},
            {"Audience": "AUTO Intenders", "Geo": "Atlanta", COL: 300000}]
    first = tg.seed_rows_to_groups(rows, COL)
    second = tg.seed_rows_to_groups(rows, COL, existing=first)
    check("ids survive a no-op re-derivation",
          [g["id"] for g in first] == [g["id"] for g in second],
          (first, second))
    edited = [dict(rows[0]), {"Audience": "AUTO Intenders", "Geo": "Atlanta", COL: 999999}]
    third = tg.seed_rows_to_groups(edited, COL, existing=first)
    check("an unchanged row keeps its id even when a sibling row changed",
          third[0]["id"] == first[0]["id"])
    check("a changed row gets a fresh id, not the stale one",
          third[1]["id"] != first[1]["id"])

    print("\nmigrated groups never fabricate resolved geography")
    migrated = tg.seed_rows_to_groups(
        [{"Audience": "DEMO Homeowner", "Geo": "Denver", COL: 1}], COL)
    check("resolved_markets is empty for a migrated (text) group",
          migrated[0]["resolved_markets"] == [])
    check("resolved_zips is empty for a migrated (text) group",
          migrated[0]["resolved_zips"] == [])
    check("geo_def kind is 'text'", migrated[0]["geo_def"]["kind"] == "text")

    print("\none-custom-audience count, across ALL groups")
    rfp_map = {"RFP Segment A": True, "RFP Segment B": True,
              "Custom Segment 1": False, "Custom Segment 2": False,
              "Custom Segment 3": False}
    groups = [
        tg.new_group(["Custom Segment 1"]),
        tg.new_group(["Custom Segment 2"]),
        tg.new_group(["Custom Segment 3"]),
    ]
    check("three groups, each with a DIFFERENT custom segment, counts as three",
          tg.custom_segment_count(groups, rfp_map) == 3)
    groups_same_custom = [
        tg.new_group(["Custom Segment 1"]),
        tg.new_group(["Custom Segment 1"]),
    ]
    check("the SAME custom segment reused across two groups counts once",
          tg.custom_segment_count(groups_same_custom, rfp_map) == 1)
    check("an unknown segment defaults to RFP-selectable, not custom",
          tg.custom_segment_count([tg.new_group(["Never Seen Before"])], rfp_map) == 0)
    check("an AND group with one custom term inside it still counts",
          tg.custom_segment_count(
              [tg.new_group(["RFP Segment A", "Custom Segment 1"], op="AND")], rfp_map) == 1)

    print("\ncolor assignment is deterministic")
    check("color cycles round-robin",
          [tg.assign_color(i) for i in range(len(tg.GROUP_COLORS) + 1)]
          == tg.GROUP_COLORS + [tg.GROUP_COLORS[0]])
    check("new_group without an explicit color still gets one",
          tg.new_group(["X"])["color"] in tg.GROUP_COLORS)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Expressions parse and render both ways; groups project to and "
          "from the flat row shape byte-identically.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
