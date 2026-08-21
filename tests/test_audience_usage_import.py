"""audience_usage_import.py -- offline, against a small synthetic workbook
(never the real one: `Premion OTT Audiences *.xlsx` carries Premion's real
delivery volumes and is gitignored, same as the master deck and the avails
PDFs). Every fixture below is built to exercise one specific rule from the
module's own docstring, not to resemble the real file row-for-row.

Every `derive_catalog_updates` call except the dedicated overrides section
below passes `overrides={}` explicitly -- the committed
`audience_component_overrides.csv` is a REAL, evolving registry (it grows
every time someone marks up a fresh uncategorized report), and a synthetic
fixture that happens to reuse a real component name must not silently
start behaving differently as that file changes. The overrides mechanism
itself is tested against small, self-contained synthetic registries
instead, passed explicitly.

    python tests/test_audience_usage_import.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl  # noqa: E402

import audience_usage_import as aui  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def _workbook(rows, header=("Segment Name", "Delivered Impressions"), sheet_name="Sheet1"):
    """A minimal in-memory .xlsx, written to a temp file and handed back as
    a path -- parse_workbook needs something openpyxl.load_workbook can
    open, and a BytesIO round-trip exercises the exact same code path a
    real upload's bytes would."""
    import tempfile
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(list(header))
    for row in rows:
        ws.append(list(row))
    handle = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    handle.close()
    wb.save(handle.name)
    return handle.name


def main():
    print("parse_workbook: header validation, No Data Targeting, blank rows, bad figures")
    path = _workbook([
        ("AUTO Intenders", 1000),
        ("No Data Targeting", 500),
        ("DEMO Age A25 Plus, HH Income 100K Plus", 2500),
        (None, None),  # a genuinely blank row -- real exports carry these
    ])
    wb = aui.parse_workbook(path)
    check("blank row skipped, No Data Targeting dropped, two real rows kept",
          len(wb.rows) == 2 and wb.dropped_no_data_targeting == 1,
          (len(wb.rows), wb.dropped_no_data_targeting))
    check("row order and values survive",
          wb.rows[0].segment == "AUTO Intenders" and wb.rows[0].impressions == 1000
          and wb.rows[1].impressions == 2500, wb.rows)

    bad_header = _workbook([("x", 1)], header=("Wrong Column", "Also Wrong"))
    try:
        aui.parse_workbook(bad_header)
        check("wrong header columns raise UsageImportError", False)
    except aui.UsageImportError:
        check("wrong header columns raise UsageImportError", True)

    missing_segment = _workbook([("", 100)])
    try:
        aui.parse_workbook(missing_segment)
        check("a row with no segment name raises UsageImportError", False)
    except aui.UsageImportError:
        check("a row with no segment name raises UsageImportError", True)

    bad_impressions = _workbook([("AUTO Intenders", "not a number")])
    try:
        aui.parse_workbook(bad_impressions)
        check("a non-numeric impressions figure raises UsageImportError", False)
    except aui.UsageImportError:
        check("a non-numeric impressions figure raises UsageImportError", True)

    empty = _workbook([])
    try:
        aui.parse_workbook(empty)
        check("a workbook with only No Data Targeting/blank rows raises UsageImportError", False)
    except aui.UsageImportError:
        check("an all-blank workbook raises UsageImportError", True)

    print("\nCLT exclusion")
    path = _workbook([
        ("DEMO Age A25 Plus, CLT 1P Lennar Homes Buy Sell Home", 1000),
        ("CLT Navy Federal Credit Union Lookalike", 5000),   # ONLY a CLT component
        ("AUTO Intenders", 2000),
    ])
    wb = aui.parse_workbook(path)
    report = aui.derive_catalog_updates(wb, existing_catalog=[], overrides={})
    updated_segments = {u.segment for u in report.updates}
    check("no CLT component ever reaches `updates`",
          not any(seg.upper().startswith("CLT") for seg in updated_segments), updated_segments)
    check("the CLT-only row contributes no usage row for a catalog component "
          "(DEMO Age A25 Plus and AUTO Intenders are the only two updates)",
          len(report.updates) == 2, updated_segments)
    check("excluded_clt lists both distinct CLT components, sorted",
          report.excluded_clt == sorted(["CLT 1P Lennar Homes Buy Sell Home",
                                         "CLT Navy Federal Credit Union Lookalike"]),
          report.excluded_clt)
    check("usage_rows keeps the FULL original transcript, CLT text included -- "
          "the historical record, not the catalog view",
          len(report.usage_rows) == 3
          and any("CLT" in r["segment_string"] for r in report.usage_rows), report.usage_rows)

    print("\nCLT alias check -- a stripped CLT component matching a real, non-CLT one")
    path = _workbook([
        ("CLT 1P Some Client Custom List", 100),
        ("Some Client Custom List", 50),  # deliberately the SAME text, minus the CLT prefix
    ])
    wb = aui.parse_workbook(path)
    report = aui.derive_catalog_updates(wb, existing_catalog=[], overrides={})
    check("the alias is found",
          report.clt_aliases == [("CLT 1P Some Client Custom List", "Some Client Custom List")],
          report.clt_aliases)

    path_no_alias = _workbook([
        ("CLT 1P Some Client Custom List", 100),
        ("AUTO Intenders", 50),
    ])
    wb_no_alias = aui.parse_workbook(path_no_alias)
    report_no_alias = aui.derive_catalog_updates(wb_no_alias, existing_catalog=[], overrides={})
    check("no alias reported when nothing actually matches",
          report_no_alias.clt_aliases == [], report_no_alias.clt_aliases)

    print("\ncategorization: own prefix, CUSTOM's embedded second prefix, and honest failure")
    path = _workbook([
        ("AUTO Intenders", 100),                                    # own prefix
        ("Lifestyle Green", 50),                                    # case-variant own prefix
        ("CUSTOM FOOD Grocery Delivery Services", 75),               # CUSTOM + resolvable 2nd prefix
        ("Custom_Doctors and Nurses", 10),                          # CUSTOM (underscore form) + unresolvable
        ("Gizmo Enthusiasts Weekly Roundup", 5),                    # no recognizable prefix at all,
                                                                     # and NOT a client-pattern match either
    ])
    wb = aui.parse_workbook(path)
    report = aui.derive_catalog_updates(wb, existing_catalog=[], overrides={})
    by_seg = {u.segment: u for u in report.updates}

    check("AUTO Intenders categorized AUTO, RFP-selectable False (workbook-only default)",
          by_seg["AUTO Intenders"].category == "AUTO"
          and by_seg["AUTO Intenders"].rfp_selectable is False, by_seg.get("AUTO Intenders"))
    check("a title-cased prefix ('Lifestyle Green') still resolves to LIFESTYLE",
          by_seg["Lifestyle Green"].category == "LIFESTYLE", by_seg.get("Lifestyle Green"))
    check("CUSTOM's embedded second prefix resolves to FOOD, still non-RFP-selectable",
          by_seg["CUSTOM FOOD Grocery Delivery Services"].category == "FOOD"
          and by_seg["CUSTOM FOOD Grocery Delivery Services"].rfp_selectable is False,
          by_seg.get("CUSTOM FOOD Grocery Delivery Services"))
    check("'Custom_Doctors...' (underscore form) is recognized as CUSTOM -- reported as a case "
          "variant -- even though its own second token doesn't resolve to a category",
          "Custom_Doctors and Nurses" in report.custom_case_variants
          and by_seg["Custom_Doctors and Nurses"].category == ""
          and by_seg["Custom_Doctors and Nurses"].rfp_selectable is False,
          (report.custom_case_variants, by_seg.get("Custom_Doctors and Nurses")))
    check("a component with no recognizable prefix at all is reported uncategorized, "
          "never guessed at",
          by_seg["Gizmo Enthusiasts Weekly Roundup"].category == ""
          and by_seg["Gizmo Enthusiasts Weekly Roundup"].source == "uncategorized",
          by_seg.get("Gizmo Enthusiasts Weekly Roundup"))
    check("uncategorized count and examples agree with the updates themselves",
          report.uncategorized == 2  # Custom_Doctors... and Gizmo...
          and set(report.uncategorized_examples) == {"Custom_Doctors and Nurses",
                                                      "Gizmo Enthusiasts Weekly Roundup"},
          (report.uncategorized, report.uncategorized_examples))

    print("\nmerge, don't replace: an existing catalog segment is authoritative for "
          "category/RFP-selectable, and only its usage numbers move")
    existing = [
        {"segment": "AUTO Intenders", "category": "AUTO", "rfp_selectable": True},
        {"segment": "SPORTS Golf", "category": "SPORTS", "rfp_selectable": False},  # a real
        # brochure segment that happens to be non-RFP-selectable already
    ]
    path = _workbook([
        ("auto intenders", 999),      # deliberately different case/spacing from the catalog's own
        ("SPORTS Golf", 42),
        ("Brand New Segment", 7),
    ])
    wb = aui.parse_workbook(path)
    report = aui.derive_catalog_updates(wb, existing, overrides={})
    by_seg = {u.segment: u for u in report.updates}
    check("the workbook's own casing never creates a second row -- the EXISTING catalog "
          "spelling wins, exactly the case/punctuation trap this whole feature exists to avoid",
          "AUTO Intenders" in by_seg and "auto intenders" not in by_seg, by_seg.keys())
    check("rfp_selectable is preserved from the existing catalog, not reset to the workbook default",
          by_seg["AUTO Intenders"].rfp_selectable is True
          and by_seg["SPORTS Golf"].rfp_selectable is False, by_seg)
    check("times_used/impressions are refreshed from the NEW workbook regardless",
          by_seg["AUTO Intenders"].times_used == 1 and by_seg["AUTO Intenders"].impressions == 999,
          by_seg["AUTO Intenders"])
    check("collided counts the two matches, gained counts only the net-new one",
          report.collided == 2 and report.gained == 1, (report.collided, report.gained))
    check("a net-new segment defaults to rfp_selectable=False",
          by_seg["Brand New Segment"].rfp_selectable is False, by_seg["Brand New Segment"])

    print("\nis_client_pattern: the four structural forms, prefix or suffix, and no false positives")
    for name in ["WEB RT- Trane LMG", "WEB RT-Out West Restaurant Group", "Rutter Mills _Web RT",
                "LOCATION RT - Jim Adler - Houston", "Location RT - Jim Adler Dallas",
                "Pima Medical Institute_RFPID-258042_RT", "Rutter Mills Address List"]:
        check(f"matches: {name!r}", aui.is_client_pattern(name), name)
    for name in ["CUSTOM Blue Bell Ice Cream Shoppers", "AUTO Intenders", "Custom Personal Injury "
                "Lawyer Intender", "CUSTOM Out West", "Lennar Homes Website RT"]:
        # Lennar's own "Website RT"/"Web Retargeting" phrasing is deliberately
        # OUTSIDE the pattern -- it doesn't spell "WEB RT" the way the other
        # rows do, which is exactly why it needs a manual override instead.
        check(f"does NOT match: {name!r}", not aui.is_client_pattern(name), name)

    print("\nload_overrides: reads a registry CSV into normalized-keyed dicts")
    import csv as _csv
    import tempfile as _tempfile
    handle = _tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False,
                                          newline="", encoding="utf-8")
    writer = _csv.writer(handle)
    writer.writerow(["component", "action", "category", "split_into"])
    writer.writerow(["Some Custom Thing", "categorize", "FOOD", ""])
    writer.writerow(["A OR B Compound", "split", "", "A|B"])
    handle.close()
    loaded = aui.load_overrides(handle.name)
    check("categorize row loaded, keyed normalized",
          loaded.get(aui.normalize("Some Custom Thing")) == {
              "action": "categorize", "category": "FOOD", "split_into": []}, loaded)
    check("split_into parses the pipe-delimited list",
          loaded.get(aui.normalize("A OR B Compound")) == {
              "action": "split", "category": "", "split_into": ["A", "B"]}, loaded)
    check("a missing registry file degrades to no overrides, not an error",
          aui.load_overrides("nonexistent_registry.csv") == {}, None)

    print("\napplying overrides: categorize, leave_blank, exclude_client, and split "
          "(attributed exactly like a real comma-stack)")
    overrides = {
        aui.normalize("Some Legal Thing"): {"action": "categorize", "category": "LEGAL", "split_into": []},
        aui.normalize("Some Blank Thing"): {"action": "leave_blank", "category": "", "split_into": []},
        aui.normalize("Some Client Name"): {"action": "exclude_client", "category": "", "split_into": []},
        aui.normalize("Compound A OR B"): {"action": "split", "category": "",
                                           "split_into": ["TRAVEL Part A", "AUTO Part B"]},
    }
    path = _workbook([
        ("Some Legal Thing", 100),
        ("Some Blank Thing", 200),
        ("Some Client Name", 300),
        ("Compound A OR B", 400),
        ("Compound A OR B, AUTO Intenders", 40),  # same compound in a SECOND, real stack
    ])
    wb = aui.parse_workbook(path)
    report = aui.derive_catalog_updates(wb, existing_catalog=[], overrides=overrides)
    by_seg = {u.segment: u for u in report.updates}

    check("categorize override sets the category, stays non-RFP-selectable",
          by_seg["Some Legal Thing"].category == "LEGAL"
          and by_seg["Some Legal Thing"].rfp_selectable is False
          and by_seg["Some Legal Thing"].source == "override-categorize", by_seg.get("Some Legal Thing"))
    check("leave_blank override -- blank category, but NOT counted as uncategorized "
          "(a confirmed decision, not a gap)",
          by_seg["Some Blank Thing"].category == ""
          and by_seg["Some Blank Thing"].source == "override-blank"
          and report.uncategorized == 0, (by_seg.get("Some Blank Thing"), report.uncategorized))
    check("exclude_client override drops the component from `updates` entirely",
          "Some Client Name" not in by_seg, by_seg.keys())
    check("...and is reported in excluded_client_override, not excluded_client_pattern",
          report.excluded_client_override == ["Some Client Name"]
          and report.excluded_client_pattern == [], report)
    check("split produces TWO separate components, neither the original compound text",
          "TRAVEL Part A" in by_seg and "AUTO Part B" in by_seg
          and "Compound A OR B" not in by_seg, by_seg.keys())
    check("EACH split piece gets the FULL stack count/impressions of every stack the compound "
          "appeared in -- 2 stacks, 440 impressions each, not divided between them",
          by_seg["TRAVEL Part A"].times_used == 2 and by_seg["TRAVEL Part A"].impressions == 440
          and by_seg["AUTO Part B"].times_used == 2 and by_seg["AUTO Part B"].impressions == 440,
          (by_seg.get("TRAVEL Part A"), by_seg.get("AUTO Part B")))
    check("a split piece still gets normal categorization -- AUTO Part B resolves via its own prefix",
          by_seg["AUTO Part B"].category == "AUTO", by_seg.get("AUTO Part B"))
    check("overrides_applied tallies every action actually used",
          report.overrides_applied == {"exclude_client": 1, "split": 2,
                                       "categorize": 1, "leave_blank": 1},
          report.overrides_applied)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Workbook parsing rejects malformed rows with a plain-language error, CLT is excluded "
          "entirely (and checked for aliasing) before it ever reaches a catalog row, "
          "categorization is mechanical (own prefix, or CUSTOM's own embedded second prefix) "
          "and honest about failing rather than guessing, and merging never lets the workbook's "
          "own spelling or defaults override an existing catalog segment's category or "
          "RFP-selectable status.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
