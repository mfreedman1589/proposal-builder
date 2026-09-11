"""Offline, free -- attribution_benchmarks.py's mapping and citation-
eligibility rule. Part of the Highlights/Takeaways rework
(ATTRIBUTION_REPORT_PLAN.md's dated section): "below benchmark: silence,
above benchmark: cite at most one rate, never the benchmark's own number."
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import attribution_benchmarks as ab  # noqa: E402


class Report:
    def __init__(self):
        self.passed, self.failed = 0, []

    def check(self, label, ok, detail=""):
        if ok:
            print(f"  PASS  {label}")
            self.passed += 1
        else:
            print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
            self.failed.append(label)


def check_mapping(rep):
    print("\nVertical -> benchmark row mapping")
    rep.check("every mapped vertical resolves to a real BENCHMARKS row",
             all(v in ab.BENCHMARKS for v in ab.VERTICAL_TO_BENCHMARK.values()))
    for unmapped in ("travel", "dining_qsr", "none", "", None):
        rep.check(f"{unmapped!r} has no benchmark row (no row, never approximated)",
                 unmapped not in ab.VERTICAL_TO_BENCHMARK)
    rep.check("Auto's row has no visitor rate (the sheet's own gap)",
             ab.BENCHMARKS["auto"]["visitor_rate"] is None)


def check_below_benchmark_is_silent(rep):
    print("\nBelow benchmark on both rates -> None")
    result = ab.benchmark_facts("legal", 0.01, 0.001)  # both well under Legal's row
    rep.check("returns None", result is None, result)


def check_unmapped_vertical_is_silent(rep):
    print("\nUnmapped vertical -> None regardless of the rate")
    result = ab.benchmark_facts("travel", 0.99, 0.99)
    rep.check("returns None even for an absurdly high rate", result is None, result)
    result_unknown = ab.benchmark_facts("unknown", 0.99, 0.99)
    rep.check("same for the literal 'unknown' vertical", result_unknown is None, result_unknown)


def check_above_benchmark_cites_only_the_cleared_rate(rep):
    print("\nAbove benchmark on ONE rate only -- only that rate is cited")
    # Legal: impression_rate 0.0222, visitor_rate 0.0023. Clear the visitor
    # rate only.
    result = ab.benchmark_facts("legal", 0.01, 0.003)
    rep.check("result is non-None", result is not None, result)
    if result:
        rep.check("names the row's label", result.get("vertical") == "Legal", result)
        rep.check("visitor_rate is cited with the real campaign figure",
                 result.get("visitor_rate") == {"campaign": 0.003, "cleared": True}, result)
        rep.check("impression_rate is NOT cited (campaign didn't clear it)",
                 "impression_rate" not in result, result)
        rep.check("the benchmark row's own raw percentages (0.0222/0.0023) never "
                 "appear anywhere in the return value",
                 0.0222 not in result.values() and 0.0023 not in result.values(), result)


def check_above_benchmark_both_rates(rep):
    print("\nAbove benchmark on BOTH rates -- both cited")
    result = ab.benchmark_facts("legal", 0.03, 0.003)
    rep.check("both rates present", result is not None
             and "impression_rate" in result and "visitor_rate" in result, result)


def check_auto_missing_visitor_rate(rep):
    print("\nAuto's missing visitor rate can't be 'cleared' -- only impression_rate can cite")
    result = ab.benchmark_facts("auto", 0.05, 0.99)  # 0.99 would clear anything, but Auto has no row
    rep.check("impression_rate cited", result is not None and "impression_rate" in result, result)
    rep.check("visitor_rate never cited for Auto, no matter the campaign's own rate",
             result is not None and "visitor_rate" not in result, result)


if __name__ == "__main__":
    rep = Report()
    check_mapping(rep)
    check_below_benchmark_is_silent(rep)
    check_unmapped_vertical_is_silent(rep)
    check_above_benchmark_cites_only_the_cleared_rate(rep)
    check_above_benchmark_both_rates(rep)
    check_auto_missing_visitor_rate(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed out of {total}")
    sys.exit(1 if rep.failed else 0)
