"""attribution_import.py, against the real MW + Cardinal exports plus
synthetic shapes for the disambiguation rules that don't need a real file.

    python tests/test_attribution_import.py

Offline and free. The real-file checks report SKIP rather than failing when
the fixtures aren't present, same convention as test_wideorbit.py -- these
are real client exports and may deliberately not be committed. The
synthetic checks always run: they prove the tab-identification rules
(header-row matching, date-gap disambiguation, count/pct splitting,
multi-RFPID rejection) against shapes built in memory, so the suite still
covers the parser's trickiest logic even on a machine with none of the real
files.
"""
import sys
from datetime import date, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import openpyxl  # noqa: E402

import attribution_import as ai  # noqa: E402

ATTRIBUTION_MW = REPO / "MW attribution excel.xlsx"
ATTRIBUTION_CARDINAL = REPO / "Premion Website Attribution Cardinal.xlsx"
DELIVERY_CARDINAL = REPO / "Premion OTT.xlsx"
DELIVERY_MW = REPO / "MW delivery.xlsx"
ATTRIBUTION_WAEPA = REPO / "Premion Website Attribution and Reach Extension (13).xlsx"

# Three complete/attribution-only fixtures, each matched by an RFPID read out
# of the files' OWN cells rather than trusted from a filename (which is
# exactly how "Premion OTT.xlsx" spent a week mislabelled as Mattress
# Warehouse's):
#   Mattress Warehouse  RFPID-260964              MW attribution excel.xlsx + MW delivery.xlsx
#   Cardinal Plumbing   RFPID-256286              Premion Website Attribution Cardinal.xlsx
#                                                  + Premion OTT.xlsx
#   WAEPA               RFPID-265618 + RFPID-263966 (a real split-IO multi-RFPID
#                        campaign, not a lifetime rollup -- see gotcha 3)
#                                                  Premion Website Attribution and Reach
#                                                  Extension (13).xlsx, no delivery file --
#                        the fixture that carries conversions, two RFPIDs, and a real
#                        DMA market cut (Washington DC / Baltimore), none of which MW or
#                        Cardinal have.


class Report:
    def __init__(self):
        self.passed, self.failed, self.skipped = 0, [], []
        self.scenario = ""

    def section(self, title):
        print(f"\n  {title}\n  {'-' * len(title)}")

    def check(self, label, ok, actual=None, expected=None):
        if ok:
            self.passed += 1
            print(f"    PASS  {label}")
        else:
            self.failed.append(f"[{self.scenario}] {label}")
            print(f"    FAIL  {label}")
            if expected is not None:
                print(f"          expected: {expected!r}")
            if actual is not None:
                print(f"          actual:   {actual!r}")
        return ok

    def equal(self, label, actual, expected):
        return self.check(label, actual == expected, actual, expected)

    def close(self, label, actual, expected, tol):
        return self.check(f"{label} (±{tol})", abs(actual - expected) <= tol, actual, expected)

    def skip(self, reason):
        self.skipped.append(f"[{self.scenario}] {reason}")
        print(f"    SKIP  {reason}")


# ---------------------------------------------------------------------------
# Real-file checks
# ---------------------------------------------------------------------------

def check_mw_attribution(rep):
    rep.scenario = "MW attribution export"
    if not ATTRIBUTION_MW.exists():
        rep.skip(f"{ATTRIBUTION_MW.name} not present")
        return
    r = ai.parse_attribution_export(str(ATTRIBUTION_MW))

    rep.section("Headline")
    rep.equal("delivered impressions (the rate-math denominator, never the deck headline)",
             r.delivered_impressions, 2341223)
    rep.equal("attributed impressions", r.attributed_impressions, 31910)
    rep.close("attributed rate", r.attributed_rate, 0.01363, 0.0001)
    rep.equal("attributed unique visitors", r.attributed_unique_visitors, 4638)

    rep.section("Advertiser + market hint")
    rep.equal("client name", r.client_name, "Mattress Warehouse")
    rep.equal("rfpid", r.rfpid, "RFPID-260964")
    rep.equal("market hint from the WUSA pixel", r.market_hint, "DC")

    rep.section("Dimension cuts")
    rep.equal("markets", len(r.by_market), 4)
    rep.equal("zips", len(r.by_zip), 242)
    rep.equal("zip heat entries", len(r.zip_heat), 242)
    rep.equal("url rows", len(r.by_url), 554)

    rep.section("Date-tab disambiguation (gotcha 2)")
    rep.equal("daily/trailing trend is exactly 7 points", len(r.daily_trend), 7)
    daily_gaps = {(r.daily_trend[i + 1].day - r.daily_trend[i].day).days
                 for i in range(len(r.daily_trend) - 1)}
    rep.equal("daily trend points are 1 day apart", daily_gaps, {1})
    rep.equal("weekly trend is exactly 7 points", len(r.weekly_trend), 7)
    weekly_gaps = {(r.weekly_trend[i + 1].day - r.weekly_trend[i].day).days
                  for i in range(len(r.weekly_trend) - 1)}
    rep.equal("weekly trend points are 7 days apart", weekly_gaps, {7})

    rep.section("Flight span excludes the trailing/pull-date window")
    # The daily trend's own dates (Sept) are well after the weekly trend's
    # (Jun-Jul) -- confirming the daily tab reflects when the file was
    # PULLED, not the campaign's flight. Flight span must come from the
    # weekly trend alone, never include the daily one.
    rep.equal("flight start", r.flight_start, date(2026, 6, 1))
    rep.equal("flight end", r.flight_end, date(2026, 7, 13))
    rep.check("flight end excludes the September trailing window",
             r.flight_end < date(2026, 8, 1), r.flight_end)


def check_cardinal_attribution(rep):
    rep.scenario = "Cardinal attribution export"
    if not ATTRIBUTION_CARDINAL.exists():
        rep.skip(f"{ATTRIBUTION_CARDINAL.name} not present")
        return
    r = ai.parse_attribution_export(str(ATTRIBUTION_CARDINAL))

    rep.section("Headline + advertiser")
    rep.equal("delivered impressions", r.delivered_impressions, 917450)
    rep.equal("attributed impressions", r.attributed_impressions, 1696)
    rep.equal("attributed unique visitors", r.attributed_unique_visitors, 169)
    rep.equal("client name", r.client_name, "Cardinal Plumbing")
    rep.equal("rfpid", r.rfpid, "RFPID-256286")
    rep.equal("market hint (lowercase wusa in the pixel string)", r.market_hint, "DC")

    rep.section("One market, several audiences/creatives (addendum's own conditional)")
    rep.equal("markets", len(r.by_market), 1)
    rep.equal("audiences", len(r.by_audience), 5)
    rep.equal("creatives", len(r.by_creative), 3)

    rep.section("Weekly trend spans the reported period, not the pull date")
    rep.equal("weekly trend point count", len(r.weekly_trend), 14)
    rep.equal("flight start", r.flight_start, date(2026, 3, 30))
    rep.equal("flight end", r.flight_end, date(2026, 6, 29))


def check_waepa_attribution(rep):
    """The third real fixture (2026-09-08 finding): a split-IO multi-RFPID
    campaign, real conversions, and a real DMA market cut -- three shapes
    MW/Cardinal don't exercise at all. No companion delivery file for this
    one (WAEPA testing didn't include one), so the headline stays the
    attribution file's own figure."""
    rep.scenario = "WAEPA attribution export"
    if not ATTRIBUTION_WAEPA.exists():
        rep.skip(f"{ATTRIBUTION_WAEPA.name} not present")
        return
    r = ai.parse_attribution_export(str(ATTRIBUTION_WAEPA))

    rep.section("Headline + advertiser")
    rep.equal("delivered impressions", r.delivered_impressions, 609760)
    rep.equal("attributed impressions", r.attributed_impressions, 4323)
    rep.equal("attributed unique visitors", r.attributed_unique_visitors, 619)
    rep.equal("client name", r.client_name, "WAEPA")
    rep.equal("market hint (lowercase wusa in the pixel string)", r.market_hint, "DC")

    rep.section("Two RFPIDs, confirmed not rejected (a real split IO, gotcha 3)")
    rep.equal("rfpid is a '+'-joined display string", r.rfpid,
             "RFPID-265618 + RFPID-263966")
    rep.equal("rfpid_breakdown has one entry per RFPID", len(r.rfpid_breakdown), 2)
    rep.equal("first RFPID's own delivered impressions",
             r.rfpid_breakdown[0]["delivered_impressions"], 549296)
    rep.equal("second RFPID's own delivered impressions",
             r.rfpid_breakdown[1]["delivered_impressions"], 60464)
    rep.equal("the RFPIDs' own delivered impressions sum to the headline total "
             "(Premion's own aggregate tabs already combine them)",
             sum(row["delivered_impressions"] for row in r.rfpid_breakdown),
             r.delivered_impressions)
    rep.check("a plain-language multi-RFPID note is in warnings, not a raised exception",
             any("2 RFPIDs" in w for w in r.warnings), r.warnings)

    rep.section("Real conversions (attributed_conversions=37, sales_amount=0 -- a real "
               "count-without-value case)")
    rep.equal("has_conversions", r.has_conversions, True)
    rep.equal("attributed_conversions", r.attributed_conversions, 37)
    rep.equal("sales_amount", r.sales_amount, 0.0)
    rep.equal("conversions_by_url row count", len(r.conversions_by_url), 13)
    rep.equal("conversions_by_url top page", r.conversions_by_url.get("https://www.waepa.org/"), 7)
    rep.equal("per-dimension conversion_impressions is real, not always zero",
             r.by_audience[1].conversion_impressions, 89)

    rep.section("Real DMA market cut (unlike MW's/Cardinal's own market rows, these are "
               "genuine Nielsen DMA names)")
    rep.equal("markets", len(r.by_market), 2)
    market_names = sorted(m.label for m in r.by_market)
    rep.equal("market names", market_names, ["BALTIMORE", "WASHINGTON, DC (HAGRSTWN)"])


def check_no_conversions_on_mw_and_cardinal(rep):
    """MW and Cardinal both carry the per-dimension conversion COLUMNS
    (conversion_impressions/conversion_rate), always zero -- but neither
    carries the top-line "Attributed Conversions" widget tab at all. This
    is the exact distinction has_conversions exists to draw: column
    presence-with-zeros is not the same as the widget being present, and
    ONLY the widget (present AND > 0) may set has_conversions."""
    rep.scenario = "has_conversions is false on MW/Cardinal despite zero-valued conversion columns"
    for path, label in ((ATTRIBUTION_MW, "MW"), (ATTRIBUTION_CARDINAL, "Cardinal")):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            continue
        r = ai.parse_attribution_export(str(path))
        rep.equal(f"{label} has_conversions", r.has_conversions, False)
        rep.equal(f"{label} attributed_conversions", r.attributed_conversions, 0)
        if r.by_audience:
            rep.equal(f"{label} still carries the (zero) conversion_impressions column",
                     r.by_audience[0].conversion_impressions, 0)


def check_cardinal_delivery(rep):
    rep.scenario = "Cardinal delivery export"
    if not DELIVERY_CARDINAL.exists():
        rep.skip(f"{DELIVERY_CARDINAL.name} not present")
        return
    r = ai.parse_delivery_export(str(DELIVERY_CARDINAL))

    rep.section("Headline (Correction 3: this is what the deck calls 'delivered')")
    rep.equal("delivered impressions", r.delivered_impressions, 917451)
    rep.close("VCR", r.vcr, 0.985189, 0.0001)
    rep.close("frequency", r.frequency, 3.412438, 0.0001)
    rep.equal("uniques", r.uniques, 268855)

    rep.section("Count/pct string splitting (gotcha 4)")
    rep.equal("top publishers resolved", len(r.top_publishers), 10)
    rep.check("no literal TOTAL row leaked into top publishers",
             all(name.upper() != "TOTAL" for name, _, _ in r.top_publishers),
             [n for n, _, _ in r.top_publishers])
    name, count, pct = r.top_publishers[0]
    rep.equal("top publisher name", name, "Pluto TV")
    rep.equal("top publisher delivered count parsed out of the combined cell", count, 112320)
    rep.close("top publisher pct parsed out of the combined cell", pct, 12.24, 0.01)

    rep.section("Per-channel VCR is a SEPARATE tab from top publishers' share-of-total pct")
    rep.close("Pluto TV real VCR", r.channel_vcr.get("Pluto TV"), 0.98961, 0.0001)
    rep.close("A+E real VCR", r.channel_vcr.get("A+E"), 0.991523, 0.0001)

    rep.section("The delivery file's OWN daily series is real, unlike the attribution file's")
    rep.equal("daily delivery spans the full quarter, not a trailing week", len(r.daily_delivery), 91)
    rep.equal("first day", r.daily_delivery[0][0], date(2026, 4, 1))
    rep.equal("last day", r.daily_delivery[-1][0], date(2026, 6, 30))

    rep.section("Pacing")
    rep.equal("booked impressions", r.booked_impressions, 900000)
    rep.check("delivery exceeded booked (a real, reportable pacing fact)",
             r.delivered_impressions > r.booked_impressions,
             (r.delivered_impressions, r.booked_impressions))


def check_mw_delivery(rep):
    """The second real pair's delivery half, and the one that carries
    ATTRIBUTION_REPORT_PLAN.md Correction 3's own headline disagreement:
    this file says 3,104,554 delivered where MW's ATTRIBUTION file says
    2,341,223 (asserted in check_mw_attribution above). The deck must
    show the first and never the second -- proven end to end against
    these two real files in tests/test_report_assembly.py, which
    replaced a synthetic stand-in once this file was found.
    """
    rep.scenario = "MW delivery export"
    if not DELIVERY_MW.exists():
        rep.skip(f"{DELIVERY_MW.name} not present")
        return
    r = ai.parse_delivery_export(str(DELIVERY_MW))

    rep.section("Headline (Correction 3: this is what the deck calls 'delivered')")
    rep.equal("delivered impressions", r.delivered_impressions, 3104554)
    rep.close("VCR", r.vcr, 0.975405, 0.0001)
    rep.close("frequency", r.frequency, 7.861103, 0.0001)
    rep.equal("uniques", r.uniques, 394926)

    rep.section("Multi-creative -- unlike Cardinal, MW's creatives are market-specific")
    rep.equal("creatives", len(r.by_creative), 3)
    names = [c[0] for c in r.by_creative]
    rep.check("MW GEOVAST is the volume leader", names[0] == "MW GEOVAST", names)

    rep.section("Per-channel VCR comes from its own tab, not the share-of-total pct")
    rep.check("channel_vcr is populated", len(r.channel_vcr) > 0, r.channel_vcr)


# ---------------------------------------------------------------------------
# Synthetic checks -- always run, no real file needed
# ---------------------------------------------------------------------------

def _write_workbook(path, sheets):
    """sheets: {title: [row, row, ...]}, row 0 is the header."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title=title[:31])
        for row in rows:
            ws.append(row)
    wb.save(path)


def check_date_gap_disambiguation(rep, tmp_path):
    rep.scenario = "Synthetic: date-tab disambiguation by gap"
    header = ["Date", "Delivered Impressions", "Attributed Impressions",
             "Attributed Rate", "Conversion Impressions", "Conversion Impressions Rate"]
    base = date(2026, 9, 1)
    daily_rows = [header] + [[base + timedelta(days=d), 100, 1, 0.01, 0, 0.0] for d in range(7)]
    weekly_base = date(2026, 6, 1)
    weekly_rows = [header] + [[weekly_base + timedelta(weeks=i), 100, 1, 0.01, 0, 0.0] for i in range(6)]
    path = tmp_path / "date_gap.xlsx"
    _write_workbook(path, {"Untitled": daily_rows, "Untitled_1": weekly_rows})
    trend = ai._date_series(ai._index_by_header(__import__("openpyxl").load_workbook(str(path))))
    rep.equal("1-day-gap tab classifies as daily", len(trend["daily"]), 7)
    rep.equal("7-day-gap tab classifies as weekly", len(trend["weekly"]), 6)
    rep.equal("neither tab is misclassified as the other", len(trend["monthly"]), 0)


def check_ranked_tab_ignored(rep, tmp_path):
    rep.scenario = "Synthetic: ranked tab is ignored in favor of the detail tab"
    ranked = [["Attributed Rate", "Conversion Impressions Rate", "Audience Name"],
             [0.05, 0.0, "Some Segment"]]
    detail = [["Audience Name", "Delivered Impressions", "Attributed Impressions",
              "Attributed Rate", "Conversion Impressions", "Conversion Impressions Rate"],
             ["Some Segment", 1000, 50, 0.05, 0, 0.0]]
    path = tmp_path / "ranked_vs_detail.xlsx"
    _write_workbook(path, {"Ranked": ranked, "Detail": detail})
    result = ai.parse_attribution_export(str(path))
    rep.equal("the fuller detail tab is used, not the ranked one",
             len(result.by_audience), 1)
    rep.equal("delivered impressions came from the detail tab (absent from the ranked one)",
             result.by_audience[0].delivered_impressions, 1000)


def check_multi_rfpid_confirmed_not_rejected(rep, tmp_path):
    """2026-09-08: replaced the hard reject -- a real WAEPA export (2 RFPIDs,
    same audiences/creatives, overlapping dates -- a split IO) was rejected
    by the old rule, which was built against a GLS/Twin Pine-style lifetime
    rollup (180+ RFPIDs) that this parser genuinely cannot distinguish from
    a split IO on its own (no per-RFPID dates or dimension rows exist to
    tell them apart). Never refuse outright any more -- the rep pulled the
    export deliberately and decides, informed by rfpid_breakdown, via
    app.py's confirm gate."""
    rep.scenario = "Synthetic: a multi-RFPID export parses successfully, confirm data populated"
    rfpid_header = ["Rfpid", "Delivered Impressions", "Attributed Impressions",
                    "Attributed Rate", "Conversion Impressions", "Conversion Impressions Rate"]
    rows = [rfpid_header] + [[f"RFPID-{i}", 100, 1, 0.01, 0, 0.0] for i in range(3)]
    path = tmp_path / "multi_rfpid.xlsx"
    _write_workbook(path, {"ByRfpid": rows})
    result = ai.parse_attribution_export(str(path))
    rep.equal("no exception -- the export parses", len(result.rfpid_breakdown), 3)
    rep.equal("rfpid is every RFPID, '+'-joined", result.rfpid,
             "RFPID-0 + RFPID-1 + RFPID-2")
    rep.check("a plain-language note names the count, in warnings (not raised)",
             any("3 RFPIDs" in w for w in result.warnings), result.warnings)


def check_single_rfpid_breakdown_always_populated(rep, tmp_path):
    """rfpid_breakdown carries one entry even for an ordinary single-RFPID
    file -- so app.py's confirm-gate code never has to special-case
    count==1 as "missing" versus "exactly one"."""
    rep.scenario = "Synthetic: a single-RFPID export still populates rfpid_breakdown"
    rfpid_header = ["Rfpid", "Delivered Impressions", "Attributed Impressions",
                    "Attributed Rate", "Conversion Impressions", "Conversion Impressions Rate"]
    rows = [rfpid_header, ["RFPID-1", 100, 1, 0.01, 0, 0.0]]
    path = tmp_path / "single_rfpid.xlsx"
    _write_workbook(path, {"ByRfpid": rows})
    result = ai.parse_attribution_export(str(path))
    rep.equal("rfpid_breakdown has exactly one entry", len(result.rfpid_breakdown), 1)
    rep.equal("rfpid is the bare single value, no '+'", result.rfpid, "RFPID-1")
    rep.check("no multi-RFPID warning fires (this fixture's only warning, if any, is the "
             "unrelated missing-trend-tab one)",
             not any("RFPID" in w for w in result.warnings), result.warnings)


def check_conversions_widget_detection(rep, tmp_path):
    """has_conversions requires the widget tab to exist AND its value to be
    > 0 -- checking existence alone would call a genuinely zero-conversion
    export "has conversions"."""
    rep.scenario = "Synthetic: has_conversions detection (widget present AND > 0)"
    widget_present_zero = [["Attributed Conversions", "Sales Amount"], [0, 0]]
    path = tmp_path / "conversions_zero.xlsx"
    _write_workbook(path, {"Widget": widget_present_zero})
    result = ai.parse_attribution_export(str(path))
    rep.equal("widget present but zero -> has_conversions is False",
             result.has_conversions, False)

    widget_present_real = [["Attributed Conversions", "Sales Amount"], [12, 4500]]
    path2 = tmp_path / "conversions_real.xlsx"
    _write_workbook(path2, {"Widget": widget_present_real})
    result2 = ai.parse_attribution_export(str(path2))
    rep.equal("widget present and > 0 -> has_conversions is True",
             result2.has_conversions, True)
    rep.equal("sales_amount parsed", result2.sales_amount, 4500.0)

    path3 = tmp_path / "conversions_absent.xlsx"
    _write_workbook(path3, {"Nothing": [["Foo"], [1]]})
    result3 = ai.parse_attribution_export(str(path3))
    rep.equal("widget tab entirely absent -> has_conversions is False",
             result3.has_conversions, False)


def check_count_pct_split(rep):
    rep.scenario = "Synthetic: count/pct cell splitting (gotcha 4)"
    rep.equal("combined cell splits into count + pct", ai._split_count_pct("112320 - 12.24%"), (112320, 12.24))
    rep.equal("a bare number has no pct", ai._split_count_pct("917451"), (917451, None))
    rep.equal("a bare number with commas", ai._split_count_pct("1,234"), (1234, None))
    rep.equal("None is zero, no pct", ai._split_count_pct(None), (0, None))


def check_header_normalization(rep):
    rep.scenario = "Synthetic: header normalization strips BOM/NBSP and case"
    rep.equal("BOM stripped", ai._normalize_cell("﻿IMPRESSIONS BY DATA SEGMENT"),
             "impressions by data segment")
    rep.equal("mid-string NBSP stripped", ai._normalize_cell("ATTRIBUTED RATE\xa0AND CONVERSION "),
             "attributed rate and conversion")


def main():
    import tempfile
    rep = Report()
    check_mw_attribution(rep)
    check_cardinal_attribution(rep)
    check_waepa_attribution(rep)
    check_no_conversions_on_mw_and_cardinal(rep)
    check_cardinal_delivery(rep)
    check_mw_delivery(rep)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        check_date_gap_disambiguation(rep, tmp_path)
        check_ranked_tab_ignored(rep, tmp_path)
        check_multi_rfpid_confirmed_not_rejected(rep, tmp_path)
        check_single_rfpid_breakdown_always_populated(rep, tmp_path)
        check_conversions_widget_detection(rep, tmp_path)
    check_count_pct_split(rep)
    check_header_normalization(rep)

    print("\n" + "=" * 78)
    print(f"{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped")
    for f in rep.failed:
        print(f"  FAILED  {f}")
    for s in rep.skipped:
        print(f"  SKIPPED {s}")
    print("=" * 78)
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
