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
TWO_PRIOR = ONE_PRIOR + [{"period_start": "2026-02-01", "period_end": "2026-02-28",
                         "attributed_rate": 0.01}]


def check_timing_gate():
    print("\nTiming gate: level sets the minimum evidence periods (item 4, 2026-09-18)")
    attribution = _export(attributed_rate=0.01, by_zip=[
        _row("20852", 5000, 0.001), _row("20853", 5000, 0.0009)])

    # High opens from month 1 -- the whole point of a High client wanting
    # aggressive recommendations early. Real candidates, not forming.
    result_high = ra.optimization_candidates(attribution, "high", ["zip"], prior_periods=None)
    check("  High/month 1: the gate is already open", len(result_high["candidates"]) > 0,
         result_high)
    equal("  High/month 1: nothing is left in forming", result_high["forming"], [])
    check("  timing_note names the evidence count, not a 'first report' refusal",
         "1 period" not in result_high["timing_note"].lower()
         or "high" not in result_high["timing_note"].lower(), result_high["timing_note"])

    # Moderate needs 2 -- month 1 is all `forming`, nothing recommended.
    result_mod1 = ra.optimization_candidates(attribution, "moderate", ["zip"], prior_periods=None)
    equal("  Moderate/month 1: zero candidates", result_mod1["candidates"], [])
    equal("  Moderate/month 1: zero watch_list", result_mod1["watch_list"], [])
    check("  Moderate/month 1: both zips are forming instead", len(result_mod1["forming"]) == 2,
         result_mod1["forming"])
    check("  timing_note names Moderate's own gap",
         "moderate" in result_mod1["timing_note"].lower()
         and "forming" in result_mod1["timing_note"].lower(), result_mod1["timing_note"])

    result2 = ra.optimization_candidates(attribution, "moderate", ["zip"], prior_periods=[])
    equal("  empty list is the same as None", result2["candidates"], [])

    result_mod2 = ra.optimization_candidates(attribution, "moderate", ["zip"], prior_periods=ONE_PRIOR)
    check("  Moderate/month 2: the gate opens", len(result_mod2["candidates"]) > 0, result_mod2)
    check("  timing_note now names the evidence count",
         "2 periods" in result_mod2["timing_note"], result_mod2["timing_note"])

    # Low needs 3 -- still closed at 2 periods, open at 3.
    result_low2 = ra.optimization_candidates(attribution, "low", ["zip"], prior_periods=ONE_PRIOR)
    equal("  Low/month 2: still below its own gate -> zero candidates",
         result_low2["candidates"], [])
    result_low3 = ra.optimization_candidates(attribution, "low", ["zip"], prior_periods=TWO_PRIOR)
    check("  Low/month 3: the gate opens", len(result_low3["candidates"]) > 0, result_low3)


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

    result_low = ra.optimization_candidates(attribution, "low", ["zip"], prior_periods=TWO_PRIOR)
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


def check_none_level():
    print("\nNone level: the engine runs, but produces zero candidates")
    attribution = _export(attributed_rate=0.01, by_zip=[_row("20852", 5000, 0.001)])
    result = ra.optimization_candidates(attribution, "none", ["zip"], prior_periods=ONE_PRIOR)
    equal("  zero candidates", result["candidates"], [])
    equal("  zero watch_list", result["watch_list"], [])
    equal("  zero already_limited", result["already_limited"], [])
    check("  timing_note explains why (level, not timing)",
         "None" in result["timing_note"], result["timing_note"])
    check("  None's own note wins even with zero evidence too (level is checked first)",
         "none" in ra.optimization_candidates(
             attribution, "none", ["zip"], prior_periods=None)["timing_note"].lower(), None)


def check_describe_optimization_candidate():
    print("\ndescribe_optimization_candidate: deterministic, no model involved")
    zip_candidate = {"dimension": "zip", "value": "20852", "campaign_rate": 0.01,
                     "multiple": 10.0, "delivered_impressions": 5000}
    text = ra.describe_optimization_candidate(zip_candidate)
    check("  names the zip", "20852" in text, text)
    check("  names the real multiple", "10.0x" in text, text)
    check("  names the real campaign rate", "1.00%" in text, text)
    check("  names the real impression count", "5,000" in text, text)
    check("  uses subtraction language", any(w in text.lower() for w in ("reduce", "remove")), text)

    day_candidate = {"dimension": "day_of_week", "value": "Wed", "campaign_rate": 0.01,
                     "multiple": 2.0, "delivered_impressions": 9000}
    day_text = ra.describe_optimization_candidate(day_candidate)
    check("  expands a day-of-week abbreviation to its full name",
         "Wednesday" in day_text and "Wed " not in day_text, day_text)

    no_multiple = {"dimension": "market", "value": "Baltimore", "campaign_rate": 0.01,
                  "multiple": None, "delivered_impressions": 1000}
    check("  degrades gracefully with no multiple (rate is 0 -- division avoided upstream)",
         "below" in ra.describe_optimization_candidate(no_multiple).lower(), None)


def check_accepted_optimizations_filters_declined():
    print("\naccepted_optimizations_from_report_json: declined entries are excluded")
    report_json = {"optimizations": [
        {"dimension": "zip", "value": "A", "decision": "accepted"},
        {"dimension": "zip", "value": "B", "decision": "edited"},
        {"dimension": "zip", "value": "C", "decision": "declined"},
    ]}
    accepted = ra.accepted_optimizations_from_report_json(report_json)
    equal("  exactly the accepted + edited entries, in order",
         [e["value"] for e in accepted], ["A", "B"])
    equal("  a report with no optimizations key at all -> empty, not an error",
         ra.accepted_optimizations_from_report_json({}), [])
    equal("  None report_json -> empty, not an error",
         ra.accepted_optimizations_from_report_json(None), [])


def check_measure_optimization_effect():
    print("\nmeasure_optimization_effect: the 'did it work' before/after")
    then_dict = {"attributed_rate": 0.01, "by_zip": [
        {"label": "20852", "delivered_impressions": 5000, "attributed_impressions": 5, "attributed_rate": 0.001},
        {"label": "20853", "delivered_impressions": 5000, "attributed_impressions": 50, "attributed_rate": 0.01},
    ]}
    now_export = _export(attributed_rate=0.014, by_zip=[
        _row("20852", 300, 0.0005), _row("20853", 6000, 0.015)])
    entry = {"dimension": "zip", "value": "20852", "final_text": "Cut 20852"}
    measured = ra.measure_optimization_effect(entry, then_dict, now_export)
    equal("  delivered_impressions_then", measured["delivered_impressions_then"], 5000)
    equal("  delivered_impressions_now", measured["delivered_impressions_now"], 300)
    check("  share_then is 50% (5000 of 10000 total)", abs(measured["share_then"] - 0.5) < 1e-9,
         measured["share_then"])
    check("  share_now dropped", measured["share_now"] < measured["share_then"], measured)
    equal("  campaign_rate_then", measured["campaign_rate_then"], 0.01)
    equal("  campaign_rate_now", measured["campaign_rate_now"], 0.014)
    check("  found_now is True (the zip still has SOME delivery)", measured["found_now"], measured)

    print("  A value that dropped out of the export entirely")
    gone_entry = {"dimension": "zip", "value": "99999", "final_text": "Cut 99999"}
    gone = ra.measure_optimization_effect(gone_entry, then_dict, now_export)
    equal("  delivered_impressions_now is 0, not an error", gone["delivered_impressions_now"], 0)
    check("  found_now is False -- genuinely gone, reported as such", not gone["found_now"], gone)


def check_optimizations_in_effect_excludes_and_measures():
    print("\noptimizations_in_effect: excludes from new candidacy, measures the effect")
    prior_reports = [{"report_json": {
        "attribution": {"attributed_rate": 0.01, "by_zip": [
            {"label": "20852", "delivered_impressions": 5000, "attributed_impressions": 5, "attributed_rate": 0.001}]},
        "optimizations": [
            {"dimension": "zip", "value": "20852", "decision": "accepted", "final_text": "Cut 20852"},
            {"dimension": "zip", "value": "20853", "decision": "declined", "final_text": None},
        ]}}]
    now = _export(attributed_rate=0.014, by_zip=[
        _row("20852", 300, 0.0005), _row("20853", 6000, 0.017), _row("20854", 6000, 0.001)])
    values, facts = ra.optimizations_in_effect(prior_reports, now)
    equal("  only the ACCEPTED value is in-effect (declined excluded)",
         values, {("zip", "20852")})
    equal("  exactly one measured fact", len(facts), 1)
    equal("  it measures the accepted value", facts[0]["value"], "20852")

    print("  the excluded value never reappears as a new candidate")
    result = ra.optimization_candidates(now, "high", ["zip"],
                                        prior_periods=[{"period_start": "x", "period_end": "y",
                                                        "attributed_rate": 0.01}],
                                        in_effect_values=values)
    check("  20852 is genuinely a rate outlier here too, but is excluded",
         "20852" not in {c["value"] for c in result["candidates"] + result["watch_list"]}, result)
    check("  20854 (never in-effect, also an outlier) still surfaces normally",
         "20854" in {c["value"] for c in result["candidates"]}, result["candidates"])

    print("\noptimizations_in_effect with no prior reports")
    empty_values, empty_facts = ra.optimizations_in_effect([], now)
    equal("  empty set", empty_values, set())
    equal("  empty list", empty_facts, [])


def check_optimization_history_full_chain():
    print("\noptimization_history: the wrap's full chain, declined entries excluded, chronological")
    now = _export(attributed_rate=0.016, by_zip=[_row("20852", 100, 0.0003)])
    prior_reports = [
        {"report_json": {
            "attribution": {"attributed_rate": 0.01, "by_zip": [
                {"label": "20852", "delivered_impressions": 5000, "attributed_impressions": 5, "attributed_rate": 0.001}]},
            "headline_facts": {"period_start": "2026-02-01", "period_end": "2026-02-28"},
            "optimizations": [{"dimension": "zip", "value": "20852", "decision": "accepted",
                               "final_text": "Cut 20852"}]}},
        {"report_json": {
            "attribution": {"attributed_rate": 0.012, "by_zip": []},
            "headline_facts": {"period_start": "2026-03-01", "period_end": "2026-03-31"},
            "optimizations": [{"dimension": "market", "value": "Baltimore", "decision": "declined",
                               "final_text": None}]}},
    ]
    history = ra.optimization_history(prior_reports, now)
    equal("  only the accepted entry survives (the declined one is excluded)", len(history), 1)
    equal("  it's the 20852 cut", history[0]["value"], "20852")
    equal("  carries its own originating period", history[0]["period_start"], "2026-02-01")
    check("  measured against the WRAP's own (final) export, not an intermediate one",
         history[0]["campaign_rate_now"] == 0.016, history[0])


def check_cross_month_consistency():
    print("\nCross-month consistency (item 3, 2026-09-18) -- THE PRINCIPLE: "
         "consistency across periods is the evidence, not magnitude alone")
    # Three zips, all material-below THIS period. 20001 was ALSO under in
    # both priors (real consistency, majority 3-for-3). 20002 was OVER in
    # one prior -- a wash, excluded outright. 20003 has no history at all
    # (a brand-new zip) -- reduces to a single-observation, still a
    # candidate, exactly the old single-period behavior.
    attribution = _export(attributed_rate=0.01, by_zip=[
        _row("20001", 5000, 0.002), _row("20002", 5000, 0.002), _row("20003", 5000, 0.002)])
    prior1 = {"period_start": "2026-01-01", "period_end": "2026-01-31", "campaign_rate": 0.01,
             "zip": [{"label": "20001", "rate": 0.002}, {"label": "20002", "rate": 0.03}]}
    prior2 = {"period_start": "2026-02-01", "period_end": "2026-02-28", "campaign_rate": 0.01,
             "zip": [{"label": "20001", "rate": 0.002}, {"label": "20002", "rate": 0.002}]}
    result = ra.optimization_candidates(
        attribution, "moderate", ["zip"], prior_periods=TWO_PRIOR,
        series_period_facts=[prior1, prior2])
    values = {c["value"] for c in result["candidates"]}
    qualifying = values | {c["value"] for c in result["watch_list"]}
    check("  20001 (consistent 3-for-3) is a candidate", "20001" in values, values)
    check("  20002 (a wash -- over-average in one prior period) is excluded entirely",
         "20002" not in qualifying and "20002" not in {c["value"] for c in result["forming"]}, result)
    check("  20003 (no history at all) still qualifies, single-observation "
         "(candidate or watch_list -- only 3 total zips means a 10% cap of 1)",
         "20003" in qualifying, qualifying)
    winner = next(c for c in result["candidates"] if c["value"] == "20001")
    check("  its consistency verdict is recorded (3 of 3 periods, under)",
         winner["consistency"] == {"periods_seen": 3, "under_count": 3, "over_count": 0,
                                   "is_wash": False, "is_consistent": True},
         winner["consistency"])

    # Ranking: consistency beats raw magnitude. 20010 is a HUGE one-month
    # swing with no history; 20011 is a smaller but persistent 2-for-2.
    attribution2 = _export(attributed_rate=0.01, by_zip=[
        _row("20010", 5000, 0.0005), _row("20011", 5000, 0.0079)])
    prior_persist = {"period_start": "2026-02-01", "period_end": "2026-02-28",
                     "campaign_rate": 0.01, "zip": [{"label": "20011", "rate": 0.0079}]}
    result2 = ra.optimization_candidates(
        attribution2, "moderate", ["zip"], prior_periods=ONE_PRIOR,
        series_period_facts=[prior_persist])
    order = [c["value"] for c in result2["candidates"]]
    check("  the persistent, smaller swing (2-for-2) outranks the single wild month (1-for-1)",
         order and order[0] == "20011", order)

    # A minority (under in 1 of 3) is real signal but not yet a majority --
    # forming, not a recommendation.
    attribution3 = _export(attributed_rate=0.01, by_zip=[_row("20020", 5000, 0.002)])
    minority_priors = [
        {"period_start": "2026-01-01", "period_end": "2026-01-31", "campaign_rate": 0.01,
         "zip": [{"label": "20020", "rate": 0.0095}]},
        {"period_start": "2026-02-01", "period_end": "2026-02-28", "campaign_rate": 0.01,
         "zip": [{"label": "20020", "rate": 0.0098}]},
    ]
    result3 = ra.optimization_candidates(
        attribution3, "moderate", ["zip"], prior_periods=TWO_PRIOR,
        series_period_facts=minority_priors)
    equal("  a minority-of-periods finding never becomes a candidate", result3["candidates"], [])
    check("  it lands in forming instead, not silently dropped",
         any(c["value"] == "20020" for c in result3["forming"]), result3["forming"])


def main():
    check_timing_gate()
    check_cross_month_consistency()
    check_material_threshold()
    check_impression_floor()
    check_already_limited_not_rediscovered()
    check_subtraction_only_shape()
    check_cap_is_a_ceiling_not_a_target()
    check_zip_cap_is_a_percentage_of_the_whole_dimension()
    check_publisher_defaults_off()
    check_disabled_dimensions_are_silent()
    check_internal_keys_never_narrated()
    check_none_level()
    check_describe_optimization_candidate()
    check_accepted_optimizations_filters_declined()
    check_measure_optimization_effect()
    check_optimizations_in_effect_excludes_and_measures()
    check_optimization_history_full_chain()
    check_real_fixtures_dont_crash()

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
