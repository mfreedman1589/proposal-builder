"""Wide Orbit schedule parsing, against the three real export formats.

    python tests/test_wideorbit.py

Offline and free. The fixtures are real exports, so the assertions are
against numbers a person actually sold -- which is the point: every format
disagrees about units and layout, and only a real file proves a parser
handles one.

If the fixtures aren't present the suite reports SKIP rather than failing,
since they're real client schedules and may deliberately not be committed.
"""

import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures" / "wideorbit"

import wideorbit  # noqa: E402


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

    def skip(self, label, why):
        self.skipped.append(label)
        print(f"    SKIP  {label} -- {why}")


# Every figure below is read off the export's own totals row, so a failure
# here means the parser disagrees with Wide Orbit, not with me.
CASES = [
    {
        "file": "ravens_campaign_schedule.xlsx",
        "name": "Campaign Schedule Report (.xlsx)",
        "format": "campaign_schedule_xlsx",
        "spots": 14, "cost": 79500.0, "impressions": 1785413, "reach": 65.5,
        "frequency": 1.8, "cpm": 44.53, "grps": 114.81, "demo": "A25-64",
        "station": "WUSA", "rows": 14, "grid_weeks": 21,
        "flight": (date(2026, 8, 3), date(2027, 1, 31)),
        # Impressions are already raw in this format.
        "raw_units": True,
        # A football schedule crosses New Year -- the week columns are M/D
        # with no year, so this is the case that catches a naive parser.
        "crosses_year": True,
    },
    {
        "file": "regency_planner.xls",
        "name": "Planner (.xls)",
        "format": "planner_xls",
        "spots": 148, "cost": 28000.0, "impressions": 1831600, "reach": 16.5,
        "frequency": 5.2, "cpm": 15.29, "grps": 86.4, "demo": "CS-A25+",
        "station": "WUSA", "rows": 6, "grid_weeks": 4,
        "flight": (date(2025, 5, 5), date(2025, 6, 1)),
        # 1,831.6 in the file means 1,831,600.
        "raw_units": False,
        "crosses_year": False,
        "expect_zero_rate_rows": True,
    },
    {
        "file": "regency_planner.pdf",
        "name": "Planner (.pdf)",
        "format": "planner_pdf",
        "spots": 130, "cost": 22000.0, "impressions": 2879100, "reach": 26.6,
        "frequency": 5.6, "cpm": 7.64, "grps": 2879.1, "demo": "CS-A25+",
        "station": "WUSA", "rows": 3, "grid_weeks": 3,
        "flight": (date(2026, 5, 4), date(2026, 6, 1)),
        "raw_units": False,
        "crosses_year": False,
    },
]


def check_case(rep, case):
    path = FIXTURES / case["file"]
    rep.scenario = case["name"]
    print("\n" + "=" * 78)
    print(f"FORMAT  {case['name']}  ({case['file']})")
    print("=" * 78)
    if not path.exists():
        rep.skip(case["name"], f"fixture not present at {path.relative_to(REPO)}")
        return

    schedule = wideorbit.parse_schedule(str(path), case["file"])
    summary = schedule.summary

    rep.section("Normalized shape")
    rep.equal("source format", schedule.source_format, case["format"])
    rep.equal("row count", len(schedule.rows), case["rows"])
    rep.check("every row has a station", all(r.station for r in schedule.rows),
              [r.station for r in schedule.rows])
    rep.check("every row carries the demo label",
              all(r.demo == case["demo"] for r in schedule.rows),
              sorted({r.demo for r in schedule.rows}))

    rep.section("Totals tie to the export's own totals row")
    rep.equal("total spots", summary.total_spots, case["spots"])
    rep.close("gross cost", summary.gross_cost, case["cost"], 0.5)
    rep.close("impressions", summary.impressions, case["impressions"], 2)
    rep.close("reach", summary.reach, case["reach"], 0.05)
    rep.close("frequency", summary.frequency, case["frequency"], 0.05)
    rep.close("CPM", summary.cpm, case["cpm"], 0.02)
    rep.close("GRPs", summary.grps, case["grps"], 0.05)
    rep.equal("demo label carried through, not hardcoded", summary.demo_label, case["demo"])
    rep.equal("station", summary.station, case["station"])
    rep.equal("flight dates", (summary.flight_start, summary.flight_end), case["flight"])

    rep.section("Units and week columns")
    # The 1000x trap: whichever way the file quotes it, CPM must reconcile.
    implied = summary.gross_cost / summary.impressions * 1000 if summary.impressions else 0
    rep.close("impressions reconcile with cost/CPM", implied, case["cpm"], 0.05)
    rep.check("impressions are raw, not thousands", summary.impressions > 100000,
              summary.impressions)
    row_total = sum(r.impressions for r in schedule.rows)
    rep.close("row impressions sum to the total", row_total, case["impressions"],
              max(2, case["impressions"] * 0.005))

    rep.equal("grid week count", len(schedule.grid_weeks), case["grid_weeks"])
    rep.check("grid weeks ascending and unique",
              schedule.grid_weeks == sorted(set(schedule.grid_weeks)), schedule.grid_weeks)
    rep.check("weeks-with-spots is a subset of the grid",
              set(schedule.weeks) <= set(schedule.grid_weeks),
              (schedule.weeks, schedule.grid_weeks))

    if case["crosses_year"]:
        years = sorted({w.year for w in schedule.grid_weeks})
        rep.check("week columns roll over into the next year", len(years) == 2, years)
        rep.check("weeks stay ascending across the rollover",
                  schedule.grid_weeks == sorted(schedule.grid_weeks), schedule.grid_weeks)

    if case.get("expect_zero_rate_rows"):
        zero = [r for r in schedule.rows if r.rate == 0]
        rep.check("$0.00 added-value rows are kept, not dropped", bool(zero),
                  [(r.program, r.rate) for r in schedule.rows])
        rep.check("kept $0 rows still carry spots", all(r.total_spots > 0 for r in zero),
                  [(r.program, r.total_spots) for r in zero])


def check_unit_normalization(rep):
    rep.scenario = "unit normalization"
    print("\n" + "=" * 78)
    print("UNIT NORMALIZATION")
    print("=" * 78)
    rep.section("cost/CPM decides the unit, not the format")
    n = wideorbit.normalize_impressions
    # Thousands, with the relationship available to prove it.
    rep.equal("1831.6 with $28,000 @ $15.29 -> raw", round(n(1831.6, 28000, 15.29)), 1831600)
    # Already raw, same relationship.
    rep.equal("1785413 with $79,500 @ $44.53 -> unchanged",
              round(n(1785413, 79500, 44.53)), 1785413)
    # A format that lies about its own units is still corrected.
    rep.equal("raw value in a thousands format is not multiplied again",
              round(n(1785413, 79500, 44.53, assume_thousands=True)), 1785413)
    rep.section("no relationship to check against")
    rep.equal("falls back to the format's convention", round(n(1831.6, assume_thousands=True)), 1831600)
    rep.equal("and leaves raw formats alone", round(n(1785413)), 1785413)
    rep.section("degenerate input")
    rep.equal("zero stays zero", n(0, 100, 5), 0.0)
    rep.equal("blank stays zero", n("", 100, 5), 0.0)
    rep.equal("$0.00 CPM doesn't divide by zero", round(n(122.4, 0, 0, assume_thousands=True)), 122400)


def check_failure_modes(rep):
    rep.scenario = "failure modes"
    print("\n" + "=" * 78)
    print("FAILURE MODES -- a bad file must degrade, never crash")
    print("=" * 78)
    rep.section("Unreadable input raises ScheduleParseError with a usable message")

    import tempfile
    tmp = Path(tempfile.gettempdir()) / "premion_wo_bad"
    tmp.mkdir(exist_ok=True)

    cases = [
        ("wrong extension", tmp / "notes.txt", b"not a schedule"),
        ("empty xlsx", tmp / "empty.xlsx", b"PK\x03\x04 not really a workbook"),
        ("pdf with no schedule", tmp / "junk.pdf", b"%PDF-1.4 nothing useful here"),
    ]
    for label, path, blob in cases:
        path.write_bytes(blob)
        try:
            wideorbit.parse_schedule(str(path), path.name)
            rep.check(label, False, "parsed successfully", "ScheduleParseError")
        except wideorbit.ScheduleParseError as exc:
            message = str(exc)
            rep.check(f"{label}: raises ScheduleParseError", True)
            rep.check(f"{label}: message is actionable",
                      len(message) > 40 and ("hand" in message or "isn't a file type" in message),
                      message)
        except Exception as exc:                                  # noqa: BLE001
            rep.check(f"{label}: raises ScheduleParseError, not {type(exc).__name__}",
                      False, f"{type(exc).__name__}: {exc}")


def main():
    rep = Report()
    for case in CASES:
        check_case(rep, case)
    check_unit_normalization(rep)
    check_failure_modes(rep)

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
