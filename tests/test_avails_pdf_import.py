"""avails_pdf_import.py against the six real Premion avails PDFs at the
repo root (gitignored -- real client pricing; reports SKIP without them,
same convention as test_wideorbit.py).

**The assertion this file leads with, per document**: every group's
impressions, summed, equals the header's own stated Total Impressions --
2,522,716 (Annapolis Cars) / 509,868 (Lawn & Leisure) / 310,800,336 (Visit
Hershey & Harrisburg) / 332,015,108 (Wilmington University) / 3,321,409
(Capital Media) / 2,834,169 (Plaza Motors Group) -- external ground truth
from each PDF's own "Media Plan Details" page, never a total this module
computes for itself. A parse that drops or double-counts a row fails this
immediately.

**Also checked, per document: `AvailsGroup.periods`** -- each Product
Details row's own Start Date/End Date/Impressions, not only the summed
total (FLOW_REWORK_PLAN.md Phase 2's foundation: seeding a per-month
flighting control and per-month avails proration both need the document's
REAL periods, not a calendar-month split of the total). Every one of the
six documents on hand is one of two shapes -- confirmed for each below,
not assumed: a single-RFPI product's one period spans the document's own
overall flight exactly (Annapolis, Lawn & Leisure, Plaza Motors, and every
one of Hershey's 12 multi-geo blocks); a monthly-broken-out product's
periods are real calendar months, clipped at the first/last only when the
flight itself doesn't land on a month boundary (Capital Media does;
Wilmington's full-year flight happens not to need clipping at either end).

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
CAPITAL_MEDIA = REPO / "Premion Media Plan_RFPID-266994_Capital Media_Undisclosed Advertiser_8-25-2026--ver0.pdf"
PLAZA_MOTORS = REPO / "Premion Media Plan_RFPID-266583_TBC, Inc - Trahan, Burden & Charles_Plaza Motors Group_8-19-2026--ver0.pdf"
ALL_FOUR = [LAWN_LEISURE, ANNAPOLIS, HERSHEY, WILMINGTON]
ALL_FILES = ALL_FOUR + [CAPITAL_MEDIA, PLAZA_MOTORS]
# Gated independently below, not added to ALL_FILES -- this fixture is newer
# than the others and its absence shouldn't skip every other real-document
# check in this file for anyone who doesn't have it yet.
LIVEWELL = REPO / "Premion Media Plan_RFPID-266894_Direct - No Agency_Livewell Animal Hospital of Alexandria_8-25-2026--ver0.pdf"

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
    missing = [p.name for p in ALL_FILES if not p.exists()]
    if missing:
        print(f"SKIP -- {len(missing)}/{len(ALL_FILES)} real avails PDFs not present "
              f"(gitignored fixtures): {missing}")
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
    check("one period, spanning the document's own overall flight exactly (a single-RFPI "
          "product's Start/End Date IS the document flight, not a copy of it)",
          len(g.periods) == 1 and (g.periods[0].start, g.periods[0].end)
          == (doc.flight_start, doc.flight_end), g.periods)
    check("the period's own impressions equal the group total (nothing to sum across)",
          g.periods[0].impressions == g.impressions, (g.periods[0].impressions, g.impressions))

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
    check("every one of the 8 groups has exactly one period spanning the document's own flight",
          all(len(g.periods) == 1 and (g.periods[0].start, g.periods[0].end)
              == (doc.flight_start, doc.flight_end) for g in doc.groups),
          [g.periods for g in doc.groups])

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
    # Multi-geo, no monthly split -- every one of the 12 blocks is still the
    # single-RFPI shape (one period each), just repeated across geographies
    # rather than months. This is the shape Phase 2's per-month flighting
    # depends on NOT misreading as "no period data" -- 12 groups, 12 periods
    # total, never folded or confused with a monthly breakdown.
    check("every one of the 12 groups has exactly one period, matching the document's own flight",
          all(len(g.periods) == 1 and (g.periods[0].start, g.periods[0].end)
              == (doc.flight_start, doc.flight_end) for g in doc.groups),
          [g.periods for g in doc.groups])

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
    # A full 12-month flight (Jul 1 2026 - Jun 30 2027) that happens to align
    # exactly to calendar months -- every period is a FULL month, none
    # clipped, unlike Capital Media below. The real point of this check:
    # each audience's own monthly figures genuinely differ month to month
    # (a real document, not twelve equal twelfths of a total), which is
    # exactly why avails have to be stored per period rather than summed and
    # divided back out.
    check("every group has 12 periods, each a full calendar month, none clipped",
          all(len(g.periods) == 12 and all(p.start.day == 1 for p in g.periods)
              for g in doc.groups),
          [[(p.start, p.end) for p in g.periods] for g in doc.groups[:1]])
    check("periods span the document's own full flight, first to last",
          all(g.periods[0].start == doc.flight_start and g.periods[-1].end == doc.flight_end
              for g in doc.groups),
          [(g.periods[0].start, g.periods[-1].end) for g in doc.groups])
    check("a period's own impressions differ month to month within one group -- real "
          "monthly figures, not the total divided into twelve equal parts",
          len({p.impressions for p in doc.groups[0].periods}) > 1,
          [p.impressions for p in doc.groups[0].periods])
    check("summing one group's own periods reproduces that group's total exactly",
          all(sum(p.impressions for p in g.periods) == g.impressions for g in doc.groups),
          [(sum(p.impressions for p in g.periods), g.impressions) for g in doc.groups])

    print("\nCapital Media (RFPID-266994) -- 1 audience x 4 monthly rows, a BARE \"County "
          "Option\" with no dash-and-name suffix at all")
    doc = api.parse_avails_pdf(str(CAPITAL_MEDIA))
    check("header total is the external ground truth", doc.total_impressions == 3321409, doc.total_impressions)
    check("every group's impressions sum to the header total (4 monthly rows, summed verbatim)",
          sum(g.impressions for g in doc.groups) == doc.total_impressions,
          sum(g.impressions for g in doc.groups))
    check("1 group -- the four monthly rows folded into one, not kept as four",
          len(doc.groups) == 1, len(doc.groups))
    if doc.groups:
        g = doc.groups[0]
        check("row_count is 4 (the real monthly breakdown, not a single flight total)",
              g.row_count == 4, g.row_count)
        # This document's own "Geography Included" cell is the bare literal
        # "County Option" -- no "- <name>" suffix, unlike every other real
        # County Option document seen so far (Wilmington's own cells all
        # carry a real name). Requiring the dash misread this as a bare DMA
        # named "County Option" (matching no real market) and, worse, skipped
        # the COUNTY branch in `_parse_block` entirely -- silently leaving
        # `zips` empty even though the document plainly has a real Zip Codes
        # table. Found from a live report: the targeting map drew nothing for
        # this document.
        check("classified as county even with no dash-and-name suffix",
              g.geo_kind == api.GEO_KIND_COUNTY, g.geo_kind)
        check("geo_name is empty -- there was never a real name to extract, and this "
              "document's own Zip Codes table names it \"County Option - null\" "
              "(literally), which is not a name either",
              g.geo_name == "", repr(g.geo_name))
        check("the real zip list is populated despite the missing dash -- this is the "
              "actual fix: it's what the targeting map draws from",
              len(g.zips) == 98, len(g.zips))
        # The flight itself is 09/21-12/20 -- neither boundary lands on a
        # month edge, so the first AND last periods are both clipped
        # (Wilmington's flight above happens to align to month boundaries;
        # this document is the one that doesn't). Exact figures, not just
        # shapes: this is the real, load-bearing case Phase 2's per-month
        # avails proration is built against.
        check("4 periods, the first and last both clipped to the real flight boundary "
              "(09/21-09/30 and 12/01-12/20, neither a full calendar month)",
              [(p.start, p.end) for p in g.periods] == [
                  (date(2026, 9, 21), date(2026, 9, 30)), (date(2026, 10, 1), date(2026, 10, 31)),
                  (date(2026, 11, 1), date(2026, 11, 30)), (date(2026, 12, 1), date(2026, 12, 20))],
              [(p.start, p.end) for p in g.periods])
        check("each period's own stated impressions, exactly as the document prints them",
              [p.impressions for p in g.periods] == [364990, 1131469, 1094970, 729980],
              [p.impressions for p in g.periods])

    print("\nPlaza Motors Group (RFPID-266583) -- 2 audiences sharing one zip list, single-"
          "flight-total rows, the FLOW_REWORK_PLAN.md Phase 3 sum-vs-max example")
    doc = api.parse_avails_pdf(str(PLAZA_MOTORS))
    check("header total is the external ground truth", doc.total_impressions == 2834169, doc.total_impressions)
    check("every group's impressions sum to the header total",
          sum(g.impressions for g in doc.groups) == doc.total_impressions,
          sum(g.impressions for g in doc.groups))
    check("flight dates", (doc.flight_start, doc.flight_end) == (date(2026, 9, 17), date(2026, 9, 30)),
          (doc.flight_start, doc.flight_end))
    check("2 groups (A35-64 and M35-64, same zip list)", len(doc.groups) == 2, len(doc.groups))
    check("every group has exactly one period, matching the document's own flight",
          all(len(g.periods) == 1 and (g.periods[0].start, g.periods[0].end)
              == (doc.flight_start, doc.flight_end) for g in doc.groups),
          [g.periods for g in doc.groups])

    if LIVEWELL.exists():
        print("\nLiveWell Animal Hospital (RFPID-266894) -- infer_entity_labels: seven rows "
              "sharing one audience, told apart only by a street address in the document's own "
              "geo_name. Four extract a clean, unique city; three (all different DC-area "
              "addresses that each reduce to \"Washington\") correctly collide and stay blank "
              "rather than getting an arbitrary, wrong-looking disambiguation.")
        doc = api.parse_avails_pdf(str(LIVEWELL))
        check("7 groups", len(doc.groups) == 7, len(doc.groups))
        labels = api.infer_entity_labels(doc)
        by_name = dict(zip((g.geo_name for g in doc.groups), labels))
        expected = {
            "277 S Washington St Alexandria VA 22314 5 Mile Radius": "Alexandria",
            "1025 Broad St Falls Church VA 22046 5 Mile Radius": "Falls Church",
            "11993 Inspiration St Reston VA 20190 Mile Radius": "Reston",
            "7000 Wisconsin Ave Chevy Chase MD 20815": "Chevy Chase",
            "945 Florida Ave NW Washington DC 20001 5 Mile Radius": None,
            "1232 3rd St NE Washington DC 20002 5 Mile Radius": None,
            "7150 12th St NW Washington DC 20012 5 Mile Radius": None,
        }
        for geo_name, want in expected.items():
            check(f"{geo_name[:45]!r} -> {want!r}", by_name.get(geo_name) == want,
                  by_name.get(geo_name))
        check("exactly 4 rows got a real label, 3 stayed blank",
              sum(1 for v in labels if v) == 4 and sum(1 for v in labels if not v) == 3,
              labels)
    else:
        print("\nSKIP -- LiveWell real avails PDF not present (gitignored fixture)")

    print("\n_infer_entity_label unit checks -- the extraction rule and its collision guard, "
          "isolated from any real PDF")
    check("DMA: the geo_name itself is the label, verbatim",
          api._infer_entity_label(api.GEO_KIND_DMA, "Baltimore") == "Baltimore")
    check("radius, no address at all (Annapolis's shape) -> None",
          api._infer_entity_label(api.GEO_KIND_RADIUS, "10mi radius [21401]") is None)
    check("radius, a real address with a directional token between suffix and city",
          api._infer_entity_label(api.GEO_KIND_RADIUS,
                                  "945 Florida Ave NW Washington DC 20001 5 Mile Radius")
          == "Washington")
    check("named_zip, a real address, no radius phrase at all (Chevy Chase's shape)",
          api._infer_entity_label(api.GEO_KIND_NAMED_ZIP, "7000 Wisconsin Ave Chevy Chase MD 20815")
          == "Chevy Chase")
    check("named_zip, an identical-across-rows human label (Plaza's shape) -> None -- no "
          "street suffix at all, nothing to extract",
          api._infer_entity_label(api.GEO_KIND_NAMED_ZIP, "Plaza Motors Group L2T Campaign Zip List")
          is None)
    check("county kind is never even attempted",
          api._infer_entity_label(api.GEO_KIND_COUNTY, "277 S Washington St Alexandria VA 22314")
          is None)

    class _FakeGroup:
        def __init__(self, geo_kind, geo_name):
            self.geo_kind, self.geo_name = geo_kind, geo_name

    class _FakeDoc:
        def __init__(self, groups):
            self.groups = groups

    collision_doc = _FakeDoc([
        _FakeGroup(api.GEO_KIND_RADIUS, "1 Main St Springfield VA 22150 5 Mile Radius"),
        _FakeGroup(api.GEO_KIND_RADIUS, "2 Elm St Springfield VA 22151 5 Mile Radius"),
        _FakeGroup(api.GEO_KIND_RADIUS, "3 Oak St Reston VA 20190 5 Mile Radius"),
    ])
    collision_labels = api.infer_entity_labels(collision_doc)
    check("two DIFFERENT addresses that both reduce to \"Springfield\" collide and BOTH stay "
          "blank, not one arbitrarily kept",
          collision_labels[:2] == [None, None], collision_labels)
    check("the third, non-colliding row still gets its real label",
          collision_labels[2] == "Reston", collision_labels)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("All six real avails PDFs parse to their own stated Total Impressions exactly, with the "
          "expected group/audience/geography structure AND per-period Start Date/End Date/"
          "Impressions for each -- no row dropped, none double-counted, none misclassified, no "
          "period's own figure apportioned or inferred. tests/test_group_scenarios.py takes it "
          "from here: the same parsed groups, resolved through the real app.apply_avails_import "
          "and the real form.")
    return 0


if __name__ == "__main__":
    code = main()
    if skipped:
        sys.exit(0)
    sys.exit(code)
