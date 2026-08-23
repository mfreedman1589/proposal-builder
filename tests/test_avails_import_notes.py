"""The avails PDF importer's unresolved-notes list -- three live bugs in one
report, all from `apply_avails_import` (app.py):

1. Every note named the WHOLE DOCUMENT ("Premion Media Plan_RFPID-...pdf"),
   never the specific group it was about. Harmless with one group; with a
   real 12-group document, two unrelated groups hitting the same failure
   read as one duplicated line, because neither said which of the 12 it
   was. Fixed: every note is prefixed by that group's own geo/audience
   label instead.

2. `apply_avails_import` never applied the calm-vs-warning split
   `geo_resolver`'s `Resolution.notes`/`.unresolved` shape exists for (see
   DECISIONS.md, the 2026-08-20 zip-messaging fix already applied to the
   interactive geo-definition expander) -- it dumped EVERY entry of a
   named-zip group's `.unresolved` as its own report line, including every
   well-formed zip that simply has no county/market on file, a case
   `.notes` already summarizes in one sentence. Fixed within that first
   pass: `.notes` is surfaced and `.unresolved` is filtered to genuinely
   malformed entries only for the named-zip branch.

3. That first fix only collapsed duplicates WITHIN one group -- the same
   calm fact still repeated once per targeting group. A real Hershey
   import (12 groups) or Wilmington import (5 groups, 303 zips since the
   continuation-page fix) turned "N zip(s) have no county on file" and "N
   zip(s) span counties in more than one market" into several near-
   identical report lines, one per group, for a fact that's true of the
   WHOLE document. Fixed: both calm facts are now reported at most once
   each, aggregated over every zip-originated group's zips combined, in
   `report["unresolved_internal"]` -- never in the rep-facing
   `report["unresolved"]` at all. Same session: the avails document's
   Attribution field stopped being reported on its own (it's still
   consumed by `_AVAILS_ATTRIBUTION_MAP` to set a Section D toggle, just
   never restated as a report line with nothing for a rep to act on).

    python tests/test_avails_import_notes.py
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
import avails_pdf_import as api  # noqa: E402
import geo_resolver  # noqa: E402
import market_lookup  # noqa: E402
market_lookup.install()  # once per process, same as install_market_lookup() in the real app

HERSHEY_PDF = REPO / ("Premion Media Plan_RFPID-260402_Direct - No Agency_"
                       "Visit Hershey & Harrisburg_4-30-2026--ver0.pdf")
WILMINGTON_PDF = REPO / ("Premion Media Plan_RFPID-253956_Direct - No Agency_"
                          "Wilmington University_1-27-2026--ver0.pdf")

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StStub:
    def __init__(self, session_state):
        self.session_state = session_state


def run_import(document):
    real, stub = app.st, _StStub({})
    app.st = stub
    try:
        return app.apply_avails_import(document)
    finally:
        app.st = real


def main():
    print("=" * 78)
    print("SCENARIO  a bulk of well-formed, unmapped zips produces ZERO "
          "rep-facing lines, one aggregate internal note")
    print("=" * 78)
    # 99900-99959: 60 distinct, well-formed zips standing in for a real add-
    # on option's bulk list. Not all of them are actually unmapped (some of
    # the 99900s are real Alaska ZCTAs) -- the true unmapped count is
    # computed via geo_resolver directly, never assumed, so this test can't
    # drift from what the zip/county data on file actually says.
    fake_zips = [str(99900 + i) for i in range(60)]
    expected_unmapped = len([z for z in fake_zips
                             if z not in geo_resolver.zips_to_counties(fake_zips).resolved])
    check("sanity: at least some of the fake range is genuinely unmapped "
          "(otherwise this scenario tests nothing)", expected_unmapped > 0, expected_unmapped)
    document = api.AvailsDocument(
        rfpid="RFPID-BULK", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=1_000_000, attribution_text="",
        groups=[api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                                geo_name="Bulk Zip Add-On", zips=fake_zips,
                                impressions=1_000_000, row_count=1)],
        source_name="test.pdf")

    new_groups, report = run_import(document)
    # Excludes the catalog-spelling note ("Homeowners" isn't an exact match
    # in this offline test's fallback catalog) -- a separate, unrelated
    # feature this suite doesn't assert on.
    zip_notes = [n for n in report["unresolved"]
                if "Bulk Zip Add-On" in n and "isn't an exact match" not in n]
    check("the group gets ZERO rep-facing lines for the calm zip fact",
          zip_notes == [], zip_notes)
    check(f"the calm fact lands once, aggregated, in unresolved_internal, "
          f"naming the real count ({expected_unmapped})",
          any(f"{expected_unmapped} zip(s) have no county on file" in n
              for n in report["unresolved_internal"]),
          report["unresolved_internal"])
    check("no single zip code from the fake list appears verbatim as its own "
          "report line anywhere (that would mean the bulk dump survived)",
          not any(any(z in n for z in fake_zips) for n in report["unresolved"]),
          report["unresolved"])
    check("the resolved group still carries every one of the 60 zips -- "
          "unmapped-to-a-market was never a drop, just quieter messaging",
          new_groups[0]["resolved_zips"] == sorted(fake_zips), new_groups[0]["resolved_zips"])

    print("\n" + "=" * 78)
    print("SCENARIO  the SAME calm fact split across two groups aggregates "
          "into ONE internal note, not two")
    print("=" * 78)
    zips_a = [str(99900 + i) for i in range(30)]
    zips_b = [str(99930 + i) for i in range(30)]
    expected_union_unmapped = len([z for z in zips_a + zips_b
                                   if z not in geo_resolver.zips_to_counties(zips_a + zips_b).resolved])
    document_split = api.AvailsDocument(
        rfpid="RFPID-SPLIT", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=200_000, attribution_text="",
        groups=[
            api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                            geo_name="Option A", zips=zips_a, impressions=100_000, row_count=1),
            api.AvailsGroup(audience_text="Auto Intenders", geo_kind=api.GEO_KIND_NAMED_ZIP,
                            geo_name="Option B", zips=zips_b, impressions=100_000, row_count=1),
        ],
        source_name="test.pdf")
    _, report_split = run_import(document_split)
    calm_fact_lines = [n for n in report_split["unresolved"]
                       if ("Option A" in n or "Option B" in n) and "isn't an exact match" not in n]
    check("neither group gets its own rep-facing calm-fact line",
          calm_fact_lines == [], calm_fact_lines)
    no_county_internal = [n for n in report_split["unresolved_internal"]
                          if "have no county on file" in n]
    check("exactly ONE internal note for the fact, not one per group",
          len(no_county_internal) == 1, report_split["unresolved_internal"])
    check(f"that one note's count is the UNION across both groups "
          f"({expected_union_unmapped}), not either group's own half",
          no_county_internal and f"{expected_union_unmapped} zip(s)" in no_county_internal[0],
          no_county_internal)

    print("\n" + "=" * 78)
    print("SCENARIO  a genuinely malformed zip still gets its own line")
    print("=" * 78)
    document2 = api.AvailsDocument(
        rfpid="RFPID-BAD", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=100_000, attribution_text="",
        groups=[api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                                geo_name="Bad Zip Option", zips=["20005", "NOT-A-ZIP"],
                                impressions=100_000, row_count=1)],
        source_name="test.pdf")
    _, report2 = run_import(document2)
    malformed_notes = [n for n in report2["unresolved"] if "Bad Zip Option" in n]
    check("the malformed entry gets its own, specific line",
          any("NOT-A-ZIP" in n for n in malformed_notes), malformed_notes)
    check("a well-formed zip (20005) is never reported as invalid",
          not any("20005" in n and "isn't a valid zip" in n for n in malformed_notes),
          malformed_notes)
    check("nothing landed in unresolved_internal for a document with no calm fact to report",
          report2["unresolved_internal"] == [], report2["unresolved_internal"])

    print("\n" + "=" * 78)
    print("SCENARIO  actionable notes are still labeled by GROUP, not the "
          "whole document")
    print("=" * 78)
    document3 = api.AvailsDocument(
        rfpid="RFPID-TWO", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=200_000, attribution_text="",
        groups=[
            api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                            geo_name="Option A", zips=["BAD-A"], impressions=100_000, row_count=1),
            api.AvailsGroup(audience_text="Auto Intenders", geo_kind=api.GEO_KIND_NAMED_ZIP,
                            geo_name="Option B", zips=["BAD-B"], impressions=100_000, row_count=1),
        ],
        source_name="a_very_long_filename_that_says_nothing_useful_twice.pdf")
    _, report3 = run_import(document3)
    check("Option A's note names Option A, not the filename",
          any("Option A" in n and "BAD-A" in n for n in report3["unresolved"]), report3["unresolved"])
    check("Option B's note names Option B, not the filename",
          any("Option B" in n and "BAD-B" in n for n in report3["unresolved"]), report3["unresolved"])
    check("neither note is bare document-filename text with no group identity",
          not any(n.startswith(document3.source_name) for n in report3["unresolved"]),
          report3["unresolved"])

    print("\n" + "=" * 78)
    print("SCENARIO  the Attribution field is still consumed (sets a "
          "Section D toggle) but never reported")
    print("=" * 78)
    document4 = api.AvailsDocument(
        rfpid="RFPID-ATTR", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=100_000, attribution_text="Innovid Reach Extension; Premion Website",
        groups=[api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_DMA,
                                geo_name="Philadelphia", impressions=100_000, row_count=1)],
        source_name="test.pdf")
    _, report4 = run_import(document4)
    check("no note anywhere quotes the document's Attribution text",
          not any("Attribution" in n for n in report4["unresolved"] + report4["unresolved_internal"]),
          report4["unresolved"] + report4["unresolved_internal"])
    check("the reach-extension toggle is still set from that same text",
          report4["field_updates"].get("linear_reach_extension") is True,
          report4["field_updates"])

    real_docs = [
        ("Visit Hershey & Harrisburg", HERSHEY_PDF, 12),
        ("Wilmington University", WILMINGTON_PDF, 5),
    ]
    for label, pdf_path, expected_groups in real_docs:
        if not pdf_path.exists():
            print(f"\nSKIP -- real document not found at {pdf_path.name}")
            continue
        print("\n" + "=" * 78)
        print(f"SCENARIO  the real {label} document's report -- before/after")
        print("=" * 78)
        document = api.parse_avails_pdf(str(pdf_path), pdf_path.name)
        _, report = run_import(document)
        # Catalog-segment-spelling notes are a separate, unrelated feature
        # (whether an audience TERM matches the catalog exactly) -- this
        # suite only asserts on the zip/attribution noise this fix targets.
        noise_lines = [n for n in report["unresolved"] if "isn't an exact match" not in n]
        check(f"{label}: {expected_groups} groups imported",
              report["n_groups"] == expected_groups, report["n_groups"])
        check(f"{label}: the impressions ground-truth line still ties out",
              report["parsed_total"] == report["total_impressions"],
              (report["parsed_total"], report["total_impressions"]))
        check(f"{label}: ZERO rep-facing zip/attribution noise lines "
              f"(got {len(noise_lines)})",
              noise_lines == [], noise_lines)
        print(f"  INFO  {label}: unresolved_internal carries "
              f"{len(report['unresolved_internal'])} aggregated fact(s), not shown to the rep:")
        for n in report["unresolved_internal"]:
            print(f"          - {n}")

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A calm zip fact (no county on file; credited to the larger of two markets) is "
          "reported at most once per import, aggregated across every group, and never in the "
          "rep-facing list at all; the Attribution field is still consumed but never restated "
          "as a report line; malformed zips and other actionable notes still get their own, "
          "correctly group-labeled line.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
