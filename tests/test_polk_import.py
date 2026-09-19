"""polk_import.py, against the real Polk Dashboard.xlsx fixture plus
synthetic shapes for the rules that don't need a real file.

    python tests/test_polk_import.py

Offline and free. The real-file checks report SKIP rather than failing when
the fixture isn't present (a real client export, gitignored, same convention
as test_attribution_import.py). The synthetic checks always run.
"""
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import openpyxl  # noqa: E402

import polk_import as pk  # noqa: E402

POLK_FIXTURE = REPO / "Polk Dashboard.xlsx"


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
        return self.check(f"{label} (+/-{tol})", abs(actual - expected) <= tol, actual, expected)

    def skip(self, reason):
        self.skipped.append(f"[{self.scenario}] {reason}")
        print(f"    SKIP  {reason}")


def _write_workbook(path, sheets):
    """sheets: {name: [[row1cells], [row2cells], ...]}"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


# ---------------------------------------------------------------------------
# Real-file checks
# ---------------------------------------------------------------------------

def check_real_fixture(rep):
    rep.scenario = "Polk Dashboard.xlsx"
    if not POLK_FIXTURE.exists():
        rep.skip(f"{POLK_FIXTURE.name} not present")
        return
    r = pk.parse_polk_export(str(POLK_FIXTURE))

    # Headline figures, re-verified directly against the file (ATTRIBUTION_
    # REPORT_PLAN.md Phase 7's own confirmed numbers).
    rep.equal("matched households", r.matched_households, 44756)
    rep.equal("target dealer sales", r.target_dealer_sales, 5)
    rep.close("buy rate", r.buy_rate, 0.0001117168648, 1e-9)
    rep.equal("matched impressions", r.matched_impressions, 591686)
    rep.close("match rate", r.match_rate, 0.9049, 1e-6)
    rep.close("campaign lift", r.campaign_lift, 1.5, 1e-9)
    rep.check("has_target_dealer_sales is True (5 > 0)", r.has_target_dealer_sales is True)

    rep.equal("2 target dealers", len(r.target_dealers), 2)
    rep.equal("132 all-dealer rows", len(r.all_dealers), 132)
    rep.equal("202 publisher-by-impressions rows", len(r.publisher_by_impressions), 202)
    rep.equal("202 publisher-by-sales rows", len(r.publisher_by_sales), 202)
    rep.equal("2 creatives", len(r.creative_by_impressions), 2)
    rep.equal("3 audience segments", len(r.audience_by_impressions), 3)
    rep.equal("3 make/model rows", len(r.make_model_rows), 3)
    rep.equal("no warnings on a complete real export", r.warnings, [])

    # Target Dealer(s) carries market rank vs. campaign rank -- the whole
    # point of that tab, distinct from All Dealers (campaign share only).
    top = r.target_dealers[0]
    rep.equal("first target dealer name", top["name"], "TED BRITT CHANTILLY FORD")
    rep.equal("first target dealer market rank", top["market_rank"], 47)
    rep.equal("first target dealer campaign rank", top["campaign_rank"], 17)

    # Share-of-impressions and share-of-sales are kept as two separate,
    # never-conflated numbers per dimension (module docstring).
    audience_top_by_impressions = r.top_audience
    rep.equal("top audience by impressions", audience_top_by_impressions["segment"],
             "AUTO Ford Intenders")
    rep.close("top audience's share of TOTAL matched impressions",
             audience_top_by_impressions["share"], 0.6478, 1e-9)
    sales_row = next(row for row in r.audience_by_sales if row["segment"] == "AUTO Ford Intenders")
    rep.close("the SAME segment's share of TARGET SALES impressions is a different number",
             sales_row["share"], 0.8814, 1e-9)

    rep.equal("top creative by impressions", r.top_creative["creative"],
             "TG11075TBrittChev062630")
    rep.equal("top publisher by impressions", r.top_publisher["publisher"], "Pluto TV")


def check_has_target_dealer_sales_false_on_zero(rep, tmp_path):
    """A campaign that's paid for and running but hasn't matched a sale yet
    is a valid, real report -- zero sales, not a parse failure. The six
    headline tabs are always present (a fixed dashboard shape); only the
    VALUE distinguishes real signal, the same "present-and->0" rule
    attribution_import.py's has_conversions already uses."""
    rep.scenario = "zero target dealer sales (synthetic)"
    path = tmp_path / "polk_zero_sales.xlsx"
    _write_workbook(path, {
        "Matched Households": [["Matched Households"], [1000]],
        "Target Dealer Sales": [["Target Dealer Sales"], [0]],
        "Buy Rate": [["Buy Rate"], [0.0]],
        "Matched Impressions": [["Matched Impressions"], [50000]],
        "Match Rate": [["Match Rate"], [0.85]],
        "Campaign Lift": [["Campaign Lift"], [1.0]],
    })
    r = pk.parse_polk_export(str(path))
    rep.equal("target_dealer_sales is 0", r.target_dealer_sales, 0)
    rep.check("has_target_dealer_sales is False when sales are genuinely zero",
             r.has_target_dealer_sales is False)
    rep.equal("matched_households still parses from a real, present tab",
             r.matched_households, 1000)
    rep.equal("no dealer roster tabs -> a plain-language warning, not a crash",
             len(r.warnings), 1)


def check_not_a_polk_export_raises(rep, tmp_path):
    rep.scenario = "not a Polk export (synthetic)"
    path = tmp_path / "not_polk.xlsx"
    _write_workbook(path, {"Unrelated": [["Foo", "Bar"], [1, 2]]})
    raised = False
    try:
        pk.parse_polk_export(str(path))
    except pk.PolkParseError as exc:
        raised = True
        rep.check("the error message names the missing anchor tab",
                 "Matched Households" in str(exc), str(exc))
    rep.check("a workbook with no Matched Households tab raises, rather than "
             "returning a silent all-zero result", raised)


def check_target_dealers_and_all_dealers_not_confused(rep, tmp_path):
    """Target Dealer(s) (7 columns, rank data) and All Dealers (4 columns, a
    strict PREFIX of the other's header) must never collide -- header-tuple
    matching is exact-length, so this is really a regression guard against a
    future refactor that started matching on a header PREFIX instead."""
    rep.scenario = "Target Dealer(s) vs. All Dealers header collision (synthetic)"
    path = tmp_path / "polk_dealers.xlsx"
    _write_workbook(path, {
        "Matched Households": [["Matched Households"], [10]],
        "Target Dealer(s)": [
            ["Selling Dealer Name", "Selling Dealer Addr", "New Sales", "Campaign Share",
             "Target Market Rank", "Campaign Rank", "Rank Diff"],
            ["DEALER A", "1 MAIN ST", 3, 0.1, 5, 2, 3],
        ],
        "All Dealers": [
            ["Selling Dealer Name", "Selling Dealer Addr", "New Sales", "Campaign Share"],
            ["DEALER B", "2 MAIN ST", 1, 0.02],
        ],
    })
    r = pk.parse_polk_export(str(path))
    rep.equal("exactly one target dealer parsed", len(r.target_dealers), 1)
    rep.equal("target dealer carries rank fields", r.target_dealers[0]["campaign_rank"], 2)
    rep.equal("exactly one all-dealers row parsed", len(r.all_dealers), 1)
    rep.check("All Dealers row has no rank fields (a plain dict, not the target-dealer shape)",
             "campaign_rank" not in r.all_dealers[0])
    rep.check("the two rosters didn't cross-contaminate",
             r.target_dealers[0]["name"] != r.all_dealers[0]["name"])


def main():
    rep = Report()
    check_real_fixture(rep)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        check_has_target_dealer_sales_false_on_zero(rep, tmp_path)
        check_not_a_polk_export_raises(rep, tmp_path)
        check_target_dealers_and_all_dealers_not_confused(rep, tmp_path)

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
