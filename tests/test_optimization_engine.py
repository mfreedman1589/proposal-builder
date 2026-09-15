"""report_assembly.optimization_candidates() -- the deterministic
optimization engine (attribution-module-framework.md 4-6, ATTRIBUTION_
REPORT_PLAN.md Phase 6, built after the advertiser spine). Pure function,
offline, no API calls -- Tier 1. Synthetic AttributionRow data for the
threshold/cap/timing-gate logic (constructed by hand so the exact numbers
that should and shouldn't qualify are known ahead of time), plus a pass
against whichever real fixtures happen to be on hand as a sanity check that
real data doesn't crash the engine.

    python tests/test_optimization_engine.py
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import report_assembly as ra          # noqa: E402
from attribution_import import AttributionRow, AttributionExport  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def equal(label, actual, expected):
    check(label, actual == expected, f"expected {expected!r}, got {actual!r}")


def _row(label, delivered, rate):
    return AttributionRow(label=label, delivered_impressions=delivered,
                          attributed_impressions=round(delivered * rate), attributed_rate=rate)


def _export(**dims):
    """A bare AttributionExport carrying only the dimension rows a test
    needs -- attributed_rate is the campaign baseline every dimension is
    compared against."""
    exp = AttributionExport(attributed_rate=dims.pop("attributed_rate", 0.01))
    for key, rows in dims.items():
        setattr(exp, key, rows)
    return exp


ONE_PRIOR = [{"period_start": "2026-01-01", "period_end": "2026-01-31", "attributed_rate": 0.01}]


def check_timing_gate():
    print("\nTiming gate: no evidence, no candidates")
    attribution = _export(attributed_rate=0.01, by_zip=[
        _row("20852", 5000, 0.001), _row("20853", 5000, 0.0009)])
    result = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=None)
    equal("  zero prior_periods -> zero candidates", result["candidates"], [])
    equal("  zero prior_periods -> zero watch_list", result["watch_list"], [])
    check("  timing_note explains why (first report)",
         "first report" in result["timing_note"].lower(), result["timing_note"])

    result2 = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=[])
    equal("  empty list is the same as None", result2["candidates"], [])

    result3 = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=ONE_PRIOR)
    check("  one prior period -> the gate opens", len(result3["candidates"]) > 0, result3)
    check("  timing_note now names the evidence count",
         "2 periods" in result3["timing_note"], result3["timing_note"])


def check_material_threshold():
    print("\nMaterial-swing threshold: exactly the framework's own 20% floor")
    # Baseline 1%. A value at 0.81% is a 19% relative drop -- just under the
    # floor. A value at 0.79% is a 21% relative drop -- just over it.
    attribution = _export(attributed_rate=0.01, by_market=[
        _row("Under the floor", 5000, 0.0081),
        _row("Over the floor", 5000, 0.0079),
    ])
    result = ra.optimization_candidates(attribution, "high", ["market"], prior_periods=ONE_PRIOR)
    values = {c["value"] for c in result["candidates"]}
    check("  a 19% relative drop does NOT qualify", "Under the floor" not in values, values)
    check("  a 21% relative drop DOES qualify", "Over the floor" in values, values)


def check_impression_floor():
    print("\nImpression floor: a value with too little volume is noise, not a finding")
    attribution = _export(attributed_rate=0.01, by_zip=[
        _row("Real volume", 5000, 0.001),      # clears the zip floor (200)
        _row("Too small", 50, 0.0001),          # below the zip floor
    ])
    result = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=ONE_PRIOR)
    values = {c["value"] for c in result["candidates"]}
    limited_values = {c["value"] for c in result["already_limited"]}
    check("  the real-volume outlier qualifies", "Real volume" in values, values)
    check("  the tiny-volume row is dropped outright -- not a candidate",
         "Too small" not in values, values)
    check("  and not reported as already-limited either (too small to say anything at all)",
         "Too small" not in limited_values, limited_values)


def check_already_limited_not_rediscovered():
    print("\nA value already suppressed in the plan is reported, not rediscovered as a cut")
    # Six ordinary days at ~10,000 impressions, one (Wednesday) already cut
    # to 1,000 by the media plan -- the MW "removed Wednesdays" reference
    # case (CLAUDE.md). Its rate is also bad, but that's beside the point:
    # the low VOLUME is what marks it as already-limited.
    days = [_row(d, 10000, 0.0095) for d in ("Mon", "Tue", "Thu", "Fri", "Sat", "Sun")]
    days.append(_row("Wed", 1000, 0.002))
    attribution = _export(attributed_rate=0.01, by_day_of_week=days)
    result = ra.optimization_candidates(attribution, "high", ["day_of_week"], prior_periods=ONE_PRIOR)
    candidate_values = {c["value"] for c in result["candidates"]}
    limited_values = {c["value"] for c in result["already_limited"]}
    check("  Wednesday is reported as already-limited", "Wed" in limited_values, limited_values)
    check("  Wednesday is NOT a fresh candidate (never recommended twice)",
         "Wed" not in candidate_values, candidate_values)


def check_subtraction_only_shape():
    print("\nStructurally subtraction-only: no over-performing value can appear as a candidate")
    attribution = _export(attributed_rate=0.01, by_creative=[
        _row("Strong performer", 5000, 0.03),   # well ABOVE baseline
        _row("Weak performer", 5000, 0.002),    # well BELOW baseline
    ])
    result = ra.optimization_candidates(attribution, "high", ["creative"], prior_periods=ONE_PRIOR)
    values = {c["value"] for c in result["candidates"]}
    check("  the over-performer never appears as a candidate", "Strong performer" not in values, values)
    check("  the under-performer does", "Weak performer" in values, values)
    check("  every candidate's own value_rate is below its campaign_rate (never the reverse)",
         all(c["value_rate"] < c["campaign_rate"] for c in result["candidates"]), result["candidates"])


def check_cap_is_a_ceiling_not_a_target():
    print("\nThe cap is a ceiling, never a target -- fewer real outliers than the cap allows is normal")
    attribution = _export(attributed_rate=0.01, by_market=[
        _row("The one real outlier", 5000, 0.001),
        _row("Fine performer 1", 5000, 0.0098),
        _row("Fine performer 2", 5000, 0.0102),
    ])
    result = ra.optimization_candidates(attribution, "high", ["market"], prior_periods=ONE_PRIOR)
    equal("  exactly one candidate surfaces, not padded to the cap", len(result["candidates"]), 1)
    equal("  its watch_list is empty", result["watch_list"], [])


def check_zip_cap_is_a_percentage_of_the_whole_dimension():
    print("\nThe ZIP cap is a percentage of the dimension's OWN total row count")
    # 20 real-volume zips, half genuinely below baseline by a wide, uniform
    # margin -- Moderate's 10% cap on 20 total zips is 2, so 8 of the 10
    # qualifying zips should land on the watch list, not the candidate list.
    rows = []
    for i in range(10):
        rows.append(_row(f"under-{i}", 5000, 0.001))       # all qualify
    for i in range(10):
        rows.append(_row(f"fine-{i}", 5000, 0.0099))        # none qualify
    attribution = _export(attributed_rate=0.01, by_zip=rows)
    result = ra.optimization_candidates(attribution, "moderate", ["zip"], prior_periods=ONE_PRIOR)
    equal("  cap = max(1, int(20 * 0.10)) = 2", len(result["candidates"]), 2)
    equal("  the other 8 qualifying zips land on the watch list", len(result["watch_list"]), 8)

    result_low = ra.optimization_candidates(attribution, "low", ["zip"], prior_periods=ONE_PRIOR)
    equal("  Low's flat cap of 2 applies regardless of total zip count",
         len(result_low["candidates"]), 2)

    result_high = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=ONE_PRIOR)
    equal("  High = int(20 * 0.20) = 4", len(result_high["candidates"]), 4)


def check_publisher_defaults_off():
    print("\nPublisher defaults off in OPTIMIZATION_DIMENSIONS_DEFAULT_ON (Matt's ruling, 2026-09-14)")
    check("  publisher is a real, selectable dimension",
         "publisher" in ra.OPTIMIZATION_DIMENSIONS, ra.OPTIMIZATION_DIMENSIONS)
    check("  but is NOT in the default-on set",
         "publisher" not in ra.OPTIMIZATION_DIMENSIONS_DEFAULT_ON,
         ra.OPTIMIZATION_DIMENSIONS_DEFAULT_ON)
    check("  every other dimension IS on by default",
         all(d in ra.OPTIMIZATION_DIMENSIONS_DEFAULT_ON
            for d in ra.OPTIMIZATION_DIMENSIONS if d != "publisher"),
         ra.OPTIMIZATION_DIMENSIONS_DEFAULT_ON)


def check_disabled_dimensions_are_silent():
    print("\nA dimension not enabled, or with no rows at all, contributes nothing -- never a fabricated empty finding")
    attribution = _export(attributed_rate=0.01, by_zip=[_row("20852", 5000, 0.001)],
                          by_market=[])  # market tab exists but has no rows
    result = ra.optimization_candidates(attribution, "high", ["zip", "market", "publisher"],
                                        prior_periods=ONE_PRIOR)
    dims_seen = {c["dimension"] for c in result["candidates"] + result["watch_list"]
                + result["already_limited"]}
    check("  market (empty rows) contributes nothing", "market" not in dims_seen, dims_seen)
    check("  publisher (no by_channel data on this synthetic export) contributes nothing",
         "publisher" not in dims_seen, dims_seen)
    check("  zip (enabled, has rows) does contribute", "zip" in dims_seen, dims_seen)


def check_internal_keys_never_narrated():
    print("\nInternal/eligibility values are prefixed _internal, same convention as zip._internal_min_share_pct")
    attribution = _export(attributed_rate=0.01, by_zip=[_row("20852", 5000, 0.001)])
    result = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=ONE_PRIOR)
    check("  evidence_periods lives under an _internal key, not a bare one",
         "_internal_evidence_periods" in result and "evidence_periods" not in result, result.keys())


def check_real_fixtures_dont_crash():
    print("\nReal fixtures (if present): the engine runs clean end to end")
    import attribution_import as ai
    for name, path in (("MW", REPO / "MW attribution excel.xlsx"),
                       ("Cardinal", REPO / "Premion Website Attribution Cardinal.xlsx"),
                       ("WAEPA", REPO / "Premion Website Attribution and Reach Extension (14).xlsx")):
        if not path.exists():
            print(f"  SKIP  {name} fixture not present")
            continue
        attribution = ai.parse_attribution_export(str(path))
        for level in ("low", "moderate", "high"):
            result = ra.optimization_candidates(
                attribution, level, ra.OPTIMIZATION_DIMENSIONS, prior_periods=ONE_PRIOR)
            check(f"  {name}/{level}: every candidate's rate is below baseline",
                 all(c["value_rate"] < c["campaign_rate"] for c in result["candidates"]),
                 result["candidates"])
            per_dimension_counts = {}
            for c in result["candidates"]:
                per_dimension_counts[c["dimension"]] = per_dimension_counts.get(c["dimension"], 0) + 1
            over_cap = [(dim, n, ra._optimization_cap(
                            dim, level, len(ra._optimization_dimension_rows(attribution, dim))))
                       for dim, n in per_dimension_counts.items()
                       if n > ra._optimization_cap(
                           dim, level, len(ra._optimization_dimension_rows(attribution, dim)))]
            check(f"  {name}/{level}: no dimension's candidate count exceeds its own real cap",
                 not over_cap, over_cap)


def main():
    check_timing_gate()
    check_material_threshold()
    check_impression_floor()
    check_already_limited_not_rediscovered()
    check_subtraction_only_shape()
    check_cap_is_a_ceiling_not_a_target()
    check_zip_cap_is_a_percentage_of_the_whole_dimension()
    check_publisher_defaults_off()
    check_disabled_dimensions_are_silent()
    check_internal_keys_never_narrated()
    check_real_fixtures_dont_crash()

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
