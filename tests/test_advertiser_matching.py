"""advertiser_matching.py -- pure, offline, no fixtures needed.

    python tests/test_advertiser_matching.py
"""
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import advertiser_matching as am  # noqa: E402


class Report:
    def __init__(self):
        self.passed, self.failed = 0, []
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


def check_normalization(rep):
    rep.scenario = "Name normalization"
    rep.equal("punctuation AND a legal suffix stripped", am.normalize_name("Cardinal Plumbing, LLC"),
             am.normalize_name("Cardinal Plumbing"))
    rep.equal("'co' inside a longer word (Cortland) is not chopped out mid-word",
             am.normalize_name("Cortland Plumbing"), "cortlandplumbing")
    rep.equal("case-insensitive", am.normalize_name("MATTRESS WAREHOUSE"),
             am.normalize_name("mattress warehouse"))
    rep.equal("empty string normalizes to empty", am.normalize_name(""), "")
    rep.equal("None normalizes to empty", am.normalize_name(None), "")


def check_exact_match_returned_alone(rep):
    rep.scenario = "An exact (normalized) match is returned alone"
    roster = [
        {"id": "1", "name": "Cardinal Plumbing"},
        {"id": "2", "name": "Cardinal Plumbing of Virginia"},
        {"id": "3", "name": "Something Unrelated"},
    ]
    results = am.find_candidates("Cardinal Plumbing, LLC", roster)
    rep.equal("exactly one candidate", len(results), 1)
    rep.equal("it's the exact match", results[0]["id"], "1")
    rep.equal("score is 1.0", results[0]["score"], 1.0)


def check_fuzzy_ranking(rep):
    rep.scenario = "Fuzzy candidates rank by similarity"
    roster = [
        {"id": "1", "name": "Mattress Warehouse"},
        {"id": "2", "name": "Warehouse Mattress Outlet"},
        {"id": "3", "name": "Completely Different Advertiser"},
    ]
    results = am.find_candidates("Mattres Warehouse", roster)  # typo'd query
    rep.check("at least one candidate returned", len(results) >= 1, len(results))
    rep.equal("closest name ranks first", results[0]["id"], "1")
    rep.check("the unrelated advertiser doesn't clear the floor",
             all(r["id"] != "3" for r in results), results)


def check_empty_roster(rep):
    rep.scenario = "Empty roster never raises"
    rep.equal("empty roster returns []", am.find_candidates("Anything", []), [])
    rep.equal("None roster returns []", am.find_candidates("Anything", None), [])


def check_market_hint_reorders_not_excludes(rep):
    rep.scenario = "Market hint reorders candidates, never excludes one"
    # Deliberately NOT identical to the query after normalization -- an
    # exact match takes the single-result path (tested above) and would
    # never exercise ranking. "& Heating" survives normalization as real
    # content, keeping both candidates fuzzy-but-not-exact.
    roster = [
        {"id": "dc", "name": "Cardinal Plumbing & Heating", "market": "DC"},
        {"id": "harrisburg", "name": "Cardinal Plumbing & Heating", "market": "Harrisburg"},
    ]
    results = am.find_candidates("Cardinal Plumbing", roster, market_hint="Harrisburg")
    rep.equal("both candidates still present despite the DC one not matching the hint",
             len(results), 2)
    rep.equal("the market-matching candidate ranks first", results[0]["id"], "harrisburg")
    rep.equal("market_match is True for the matching one", results[0]["market_match"], True)
    rep.equal("market_match is False for the non-matching one", results[1]["market_match"], False)


def check_flight_overlap_reorders(rep):
    rep.scenario = "Flight overlap reorders candidates among fuzzy ties"
    roster = [
        {"id": "old", "name": "Some Advertiser & Sons",
         "flight_start": date(2025, 1, 1), "flight_end": date(2025, 3, 31)},
        {"id": "current", "name": "Some Advertiser & Sons",
         "flight_start": date(2026, 9, 1), "flight_end": date(2026, 11, 30)},
    ]
    results = am.find_candidates(
        "Some Advertiser", roster,
        report_start=date(2026, 9, 15), report_end=date(2026, 9, 30))
    rep.equal("both candidates present", len(results), 2)
    rep.equal("the overlapping-flight candidate ranks first", results[0]["id"], "current")
    rep.equal("flight_overlap True for the current one", results[0]["flight_overlap"], True)
    rep.equal("flight_overlap False for the old one", results[1]["flight_overlap"], False)


def check_incomplete_hints_are_unknown_not_false(rep):
    rep.scenario = "A missing hint is unknown, not a mismatch"
    roster = [{"id": "1", "name": "Some Advertiser Co"}]
    results = am.find_candidates("Some Advertiser", roster, market_hint="DC",
                                 report_start=date(2026, 1, 1), report_end=None)
    rep.equal("one candidate", len(results), 1)
    rep.equal("market_match is None (candidate has no market field)",
             results[0]["market_match"], None)
    rep.equal("flight_overlap is None (report_end missing)", results[0]["flight_overlap"], None)


def main():
    rep = Report()
    check_normalization(rep)
    check_exact_match_returned_alone(rep)
    check_fuzzy_ranking(rep)
    check_empty_roster(rep)
    check_market_hint_reorders_not_excludes(rep)
    check_flight_overlap_reorders(rep)
    check_incomplete_hints_are_unknown_not_false(rep)

    print("\n" + "=" * 78)
    print(f"{rep.passed} passed, {len(rep.failed)} failed")
    for f in rep.failed:
        print(f"  FAILED  {f}")
    print("=" * 78)
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
