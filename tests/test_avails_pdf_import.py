"""avails_pdf_import.py against the four real Premion avails PDFs at the
repo root (gitignored -- real client pricing; reports SKIP without them,
same convention as test_wideorbit.py).

**The assertion this file leads with, per document**: every group's
impressions, summed, equals the header's own stated Total Impressions --
2,522,716 (Annapolis Cars) / 509,868 (Lawn & Leisure) / 310,800,336 (Visit
Hershey & Harrisburg) / 332,015,108 (Wilmington University) -- external
ground truth from each PDF's own "Media Plan Details" page, never a total
this module computes for itself. A parse that drops or double-counts a
row fails this immediately.

**No second, hand-typed copy of these documents to cross-check against.**
Three of these four documents used to also have a hand-transcribed
ANNAPOLIS_ROWS/HERSHEY_ROWS/WILMINGTON_ROWS in
tests/group_scenario_fixtures.py, and this file asserted the parser's
output matched that transcription exactly -- which meant two independent
typings of the same real document, with only that one cross-check standing
between them drifting apart unnoticed. group_scenario_fixtures.py now
builds its scenarios FROM this parser's own output (through the real
`app.apply_avails_import`, the same call the D2 uploader makes) rather than
a second transcription, so that comparison would just prove this module
agrees with itself. What's left here is what stays genuinely independent of
any fixture: the header-total tie for all four documents (above), and each
document's own internal structure -- group counts, geography
classification, monthly-row folding -- checked directly against what the
real PDF says, never against another file's copy of it.
tests/test_group_scenarios.py is where the resulting groups get proven
correct end to end, against the real Resolve button and
`app.resolve_group_geography`.

    python tests/test_avails_pdf_import.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import avails_pdf_import as api  # noqa: E402

LAWN_LEISURE = REPO / "Premion Media Plan_RFPID-265521_Direct - No Agency_Lawn & Leisure_7-28-2026--ver0.pdf"
ANNAPOLIS = REPO / "Premion Media Plan_RFPID-253813_SR&B Advertising_Annapolis Cars_1-23-2026--ver0.pdf"
HERSHEY = REPO / "Premion Media Plan_RFPID-260402_Direct - No Agency_Visit Hershey & Harrisburg_4-30-2026--ver0.pdf"
WILMINGTON = REPO / "Premion Media Plan_RFPID-253956_Direct - No Agency_Wilmington University_1-27-2026--ver0.pdf"
ALL_FOUR = [LAWN_LEISURE, ANNAPOLIS, HERSHEY, WILMINGTON]

failures = []
skipped = False


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    global skipped
    missing = [p.name for p in ALL_FOUR if not p.exists()]
    if missing:
        print(f"SKIP -- {len(missing)}/4 real avails PDFs not present (gitignored fixtures): {missing}")
        skipped = True
        return 0

    print("Lawn & Leisure (RFPID-265521) -- single row, no monthly breakdown, radius with NO bracketed origin")
    doc = api.parse_avails_pdf(str(LAWN_LEISURE))
    check("header total is the external ground truth", doc.total_impressions == 509868, doc.total_impressions)
    check("every group's impressions sum to the header total",
          sum(g.impressions for g in doc.groups) == doc.total_impressions,
          sum(g.impressions for g in doc.groups))
    check("advertiser", doc.advertiser == "Lawn & Leisure", doc.advertiser)
    check("agency", doc.agency == "Direct - No Agency", doc.agency)
    check("flight dates", (doc.flight_start, doc.flight_end) == (date(2026, 9, 6), date(2026, 10, 11)),
          (doc.flight_start, doc.flight_end))
    check("one group", len(doc.groups) == 1, doc.groups)
    g = doc.groups[0]
    check("classified as radius", g.geo_kind == api.GEO_KIND_RADIUS, g.geo_kind)
    check("radius miles parsed from the un-bracketed name ('10 Mile Radius Zips')",
          g.radius_miles == 10.0, g.radius_miles)
    check("no origin -- the bracket is genuinely absent on this sample, not a parse failure",
          g.radius_origin == "", g.radius_origin)
    check("audience expression parsed as two AND terms",
          g.audience_text == "(DEMO Homeowner) AND (HH Income 200K Plus)", g.audience_text)

    print("\nAnnapolis Cars (RFPID-253813) -- 8 single-flight-total rows, radius WITH bracketed origin")
    doc = api.parse_avails_pdf(str(ANNAPOLIS))
    check("header total is the external ground truth", doc.total_impressions == 2522716, doc.total_impressions)
    check("every group's impressions sum to the header total",
          sum(g.impressions for g in doc.groups) == doc.total_impressions,
          sum(g.impressions for g in doc.groups))
    check("8 groups, one per (audience, radius) pair", len(doc.groups) == 8, len(doc.groups))
    check("all 8 are classified as radius", all(g.geo_kind == api.GEO_KIND_RADIUS for g in doc.groups),
          [g.geo_kind for g in doc.groups])
    check("all 8 share the same bracketed origin", all(g.radius_origin == "21401" for g in doc.groups),
          [g.radius_origin for g in doc.groups])
    check("each audience appears at both 10mi and 5mi (4 audiences x 2 radii = 8)",
          sorted(g.radius_miles for g in doc.groups) == sorted([10.0, 5.0] * 4),
          [g.radius_miles for g in doc.groups])

    print("\nVisit Hershey & Harrisburg (RFPID-260402) -- mixed DMA + named-zip geography, "
          "two AND-audiences")
    doc = api.parse_avails_pdf(str(HERSHEY))
    check("header total is the external ground truth", doc.total_impressions == 310800336, doc.total_impressions)
    check("every group's impressions sum to the header total",
          sum(g.impressions for g in doc.groups) == doc.total_impressions,
          sum(g.impressions for g in doc.groups))
    check("12 groups (2 audiences x (4 markets + 2 zip add-ons))", len(doc.groups) == 12, len(doc.groups))
    check("advertiser survives the mid-name line wrap ('Visit Hershey &' / 'Harrisburg')",
          doc.advertiser == "Visit Hershey & Harrisburg", doc.advertiser)
    dma_groups = [g for g in doc.groups if g.geo_kind == api.GEO_KIND_DMA]
    zip_groups = [g for g in doc.groups if g.geo_kind == api.GEO_KIND_NAMED_ZIP]
    check("8 DMA groups, 4 named-zip groups", (len(dma_groups), len(zip_groups)) == (8, 4),
          (len(dma_groups), len(zip_groups)))
    check("exactly two distinct audience expressions across all 12 groups",
          len({g.audience_text for g in doc.groups}) == 2, {g.audience_text for g in doc.groups})
    # The real document's own page order is NOT strictly audience-major --
    # it interleaves (Philly-A, Baltimore-A, DC-A, Philly-B, NY-A, NY-B,
    # Baltimore-B, DC-B, ...) rather than finishing audience A before
    # starting B. The directive this importer was built against says
    # "document order" explicitly, so this parser preserves the PDF's real
    # page sequence rather than re-sorting into audience-major -- deliberate,
    # not a bug (test_group_scenarios.py checks that plan_lines_from_groups
    # un-interleaves it downstream).
    check("the real document's own order is NOT strictly audience-major (documented, not a bug)",
          [g.audience_text for g in doc.groups][:5] !=
          sorted([g.audience_text for g in doc.groups][:5]), None)
    philly_named = [g for g in zip_groups if "philly" in g.geo_name.lower()]
    ny_named = [g for g in zip_groups if g is not None and "ny" in g.geo_name.lower()
               and "philly" not in g.geo_name.lower()]
    check("2 Philly named-zip groups, 2 NY named-zip groups (one per audience, each)",
          len(philly_named) == 2 and len(ny_named) == 2, (len(philly_named), len(ny_named)))
    check("both Philly named-zip groups parsed the SAME zip list (one per audience, same geography)",
          {frozenset(g.zips) for g in philly_named} and len({frozenset(g.zips) for g in philly_named}) == 1,
          [len(g.zips) for g in philly_named])
    check("both NY named-zip groups parsed the SAME zip list",
          {frozenset(g.zips) for g in ny_named} and len({frozenset(g.zips) for g in ny_named}) == 1,
          [len(g.zips) for g in ny_named])
    check("the Philly and NY zip lists are disjoint and non-trivial (52 vs 34 zips, per the "
          "document's own two named options)",
          bool(philly_named[0].zips) and bool(ny_named[0].zips)
          and not (set(philly_named[0].zips) & set(ny_named[0].zips))
          and len(philly_named[0].zips) == 52 and len(ny_named[0].zips) == 34,
          (len(philly_named[0].zips), len(ny_named[0].zips)))

    print("\nWilmington University (RFPID-253956) -- 5 audiences x 12 monthly rows each, "
          "County Option")
    doc = api.parse_avails_pdf(str(WILMINGTON))
    check("header total is the external ground truth", doc.total_impressions == 332015108, doc.total_impressions)
    check("every group's impressions sum to the header total (60 monthly rows, summed verbatim)",
          sum(g.impressions for g in doc.groups) == doc.total_impressions,
          sum(g.impressions for g in doc.groups))
    check("5 groups, one per audience -- the monthly rows folded into the group, not kept as 60",
          len(doc.groups) == 5, len(doc.groups))
    check("every row_count is 12 (the real monthly breakdown, not a single flight total)",
          all(g.row_count == 12 for g in doc.groups), [g.row_count for g in doc.groups])
    check("all 5 classified as county", all(g.geo_kind == api.GEO_KIND_COUNTY for g in doc.groups),
          [g.geo_kind for g in doc.groups])
    check("5 distinct audience expressions (one per group, none repeated)",
          len({g.audience_text for g in doc.groups}) == 5, [g.audience_text for g in doc.groups])
    counties_by_group = [frozenset(c.strip() for c in g.county_list.split(";")) for g in doc.groups]
    check("every group parsed the SAME 10-county list (Premion's own DERIVED summary -- "
          "real, but never the resolution input; see the zips check below)",
          len(set(counties_by_group)) == 1 and len(counties_by_group[0]) == 10,
          [len(c) for c in counties_by_group])
    # A County Option block is zip-originated too, same as Zip Option --
    # confirmed directly against this real document: every County Option
    # page carries a genuine Zip Codes table (17527, 17555, 18015, 18042
    # among the real entries -- the ones a Salisbury/Lancaster/Northampton
    # gap report once flagged as "out of county," which they never were;
    # the county LIST is a lossy summary of the zips, not the other way
    # round). This is also the guard for a real, separate parser bug this
    # document exposed: _block_zip_list used to gate every page on the
    # literal string "Zip Codes", which a WRAPPED table's continuation page
    # never repeats -- silently dropping every page after the first and
    # reading only 48 of these 303 real zips. Each of Wilmington's 5 blocks
    # spans exactly 2 pages (the only blocks, across all 4 real documents on
    # hand, that span more than one page at all -- this bug was invisible
    # everywhere else purely because nothing else happens to wrap).
    zips_by_group = [frozenset(g.zips) for g in doc.groups]
    check("every group's real zip list is populated (not just the derived county summary)",
          all(len(z) > 0 for z in zips_by_group), [len(z) for z in zips_by_group])
    check("every group parsed the SAME zip list as every other (one document-wide "
          "geography, matching the county-list check above)",
          len(set(zips_by_group)) == 1, [len(z) for z in zips_by_group])
    check("the continuation page's own zips are present, not silently dropped -- "
          "303 real zips, not 48 (page 1 alone)",
          len(zips_by_group[0]) == 303, len(zips_by_group[0]))
    check("the specific zips a real gap report once flagged as \"out of county\" are "
          "genuinely present in the document's own list -- they were never a mismatch, "
          "they were on the dropped page",
          {"17527", "17555", "18015", "18042"} <= zips_by_group[0],
          zips_by_group[0] & {"17527", "17555", "18015", "18042"})

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("All four real avails PDFs parse to their own stated Total Impressions exactly, with the "
          "expected group/audience/geography structure for each -- no row dropped, none double-"
          "counted, none misclassified. tests/test_group_scenarios.py takes it from here: the same "
          "parsed groups, resolved through the real app.apply_avails_import and the real form.")
    return 0


if __name__ == "__main__":
    code = main()
    if skipped:
        sys.exit(0)
    sys.exit(code)
