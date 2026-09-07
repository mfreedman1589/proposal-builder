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

# Two complete client pairs, each matched by an RFPID read out of the files'
# OWN cells rather than trusted from a filename (which is exactly how
# "Premion OTT.xlsx" spent a week mislabelled as Mattress Warehouse's):
#   Mattress Warehouse  RFPID-260964  MW attribution excel.xlsx + MW delivery.xlsx
#   Cardinal Plumbing   RFPID-256286  Premion Website Attribution Cardinal.xlsx
#                                     + Premion OTT.xlsx


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


def check_multi_rfpid_rejected(rep, tmp_path):
    rep.scenario = "Synthetic: a multi-RFPID rollup export is rejected, not summed"
    rfpid_header = ["Rfpid", "Delivered Impressions", "Attributed Impressions",
                    "Attributed Rate", "Conversion Impressions", "Conversion Impressions Rate"]
    rows = [rfpid_header] + [[f"RFPID-{i}", 100, 1, 0.01, 0, 0.0] for i in range(3)]
    path = tmp_path / "multi_rfpid.xlsx"
    _write_workbook(path, {"ByRfpid": rows})
    try:
        ai.parse_attribution_export(str(path))
        rep.check("a 3-RFPID export raises AttributionParseError", False, "no exception raised")
    except ai.AttributionParseError as exc:
        rep.check("a 3-RFPID export raises AttributionParseError", True)
        rep.check("the error names the count, not a picked/summed value",
                 "3" in str(exc), str(exc))


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
    check_cardinal_delivery(rep)
    check_mw_delivery(rep)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        check_date_gap_disambiguation(rep, tmp_path)
        check_ranked_tab_ignored(rep, tmp_path)
        check_multi_rfpid_rejected(rep, tmp_path)
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
