"""report_assembly.zip_optimization_groups() -- the ZIP-level half of the
optimization engine, split out from optimization_candidates() in round 3
(2026-09-22 review, item 3: "we don't buy single ZIPs, we REMOVE ZIPs or
BATCH them"). Pure function, offline, no API calls -- Tier 1.

This file exists because the split left zip with NO direct pure-function
coverage of its own: test_optimization_engine.py's zip-specific checks
(timing gate, impression floor, cross-month consistency, in-effect
exclusion) were re-pointed at "market" to keep testing the FLAT engine,
which zip no longer reaches at all. Everything that's specific to the
GROUP shape -- the three tier thresholds, the minimum-group-size fold,
describe_zip_group/rebuild_zip_group -- has no flat-engine analog and is
covered only here.

    python tests/test_zip_optimization_groups.py
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


def _export(attributed_rate, rows):
    exp = AttributionExport(attributed_rate=attributed_rate)
    exp.by_zip = rows
    return exp


ONE_PRIOR = [{"period_start": "2026-01-01", "period_end": "2026-01-31", "attributed_rate": 0.01}]
TWO_PRIOR = ONE_PRIOR + [{"period_start": "2026-02-01", "period_end": "2026-02-28",
                         "attributed_rate": 0.01}]

# Three "Remove"-tier zips (0.1x baseline, well under ZIP_REMOVE_MULTIPLE),
# real volume, no history -- the standard fixture most checks below start
# from: exactly at ZIP_GROUP_MIN_SIZE, so a real group forms the moment the
# timing gate is open and nothing else excludes them.
def _three_remove_zips(baseline=0.01):
    return [_row("20001", 5000, baseline * 0.1), _row("20002", 5000, baseline * 0.1),
           _row("20003", 5000, baseline * 0.1)]


def check_tier_thresholds():
    print("\n_zip_tier: Matt's own three thresholds, exact boundaries")
    equal("  1.2x exactly -> Group A", ra._zip_tier(1.2), "a")
    equal("  1.199x -> untouched (just under Group A's own floor)", ra._zip_tier(1.199), None)
    equal("  0.8x exactly -> untouched (Group B's own floor is a STRICT <)", ra._zip_tier(0.8), None)
    equal("  0.799x -> Group B", ra._zip_tier(0.799), "b")
    equal("  0.5x exactly -> Remove", ra._zip_tier(0.5), "remove")
    equal("  0.501x -> Group B, not Remove", ra._zip_tier(0.501), "b")
    equal("  1.0x (parity) -> untouched", ra._zip_tier(1.0), None)
    equal("  None multiple (no baseline) -> untouched", ra._zip_tier(None), None)


def check_timing_gate():
    print("\nTiming gate: shares _LEVEL_GATE_MIN_PERIODS/evidence_periods with the flat engine")
    attribution = _export(0.01, _three_remove_zips())

    result_high = ra.zip_optimization_groups(attribution, "high", prior_periods=None)
    check("  High/month 1: the gate is already open -- a real Remove group forms",
         len(result_high["groups"]) == 1 and result_high["groups"][0]["tier"] == "remove",
         result_high)
    equal("  High/month 1: nothing left in forming", result_high["forming"], [])

    result_mod1 = ra.zip_optimization_groups(attribution, "moderate", prior_periods=None)
    equal("  Moderate/month 1: zero groups", result_mod1["groups"], [])
    check("  Moderate/month 1: all three zips are forming instead",
         len(result_mod1["forming"]) == 3, result_mod1["forming"])
    check("  timing_note names Moderate's own gap",
         "moderate" in result_mod1["timing_note"].lower()
         and "forming" in result_mod1["timing_note"].lower(), result_mod1["timing_note"])

    result_mod2 = ra.zip_optimization_groups(attribution, "moderate", prior_periods=ONE_PRIOR)
    check("  Moderate/month 2: the gate opens", len(result_mod2["groups"]) == 1, result_mod2)

    result_low2 = ra.zip_optimization_groups(attribution, "low", prior_periods=ONE_PRIOR)
    equal("  Low/month 2: still below its own gate -> zero groups", result_low2["groups"], [])
    result_low3 = ra.zip_optimization_groups(attribution, "low", prior_periods=TWO_PRIOR)
    check("  Low/month 3: the gate opens", len(result_low3["groups"]) == 1, result_low3)


def check_none_level():
    print("\nNone level: the engine runs, but produces zero groups")
    attribution = _export(0.01, _three_remove_zips())
    result = ra.zip_optimization_groups(attribution, "none", prior_periods=ONE_PRIOR)
    equal("  zero groups", result["groups"], [])
    equal("  zero forming", result["forming"], [])
    check("  timing_note explains why (level, not timing)", "None" in result["timing_note"],
         result["timing_note"])


def check_min_size_fold():
    print("\nZIP_GROUP_MIN_SIZE: a Remove group under the floor folds into Group B "
         "(Matt's own ruling -- never propose cutting one or two ZIPs)")
    # Two Remove-tier zips alone -- folds into "b", but "b" then also has
    # only 2, so BOTH fall through to forming, not a recommendation.
    attribution = _export(0.01, [_row("20001", 5000, 0.001), _row("20002", 5000, 0.001)])
    result = ra.zip_optimization_groups(attribution, "high", prior_periods=ONE_PRIOR)
    equal("  no group forms -- 2 is under the minimum even after folding", result["groups"], [])
    check("  both zips land in forming, not silently dropped",
         {e["value"] for e in result["forming"]} == {"20001", "20002"}, result["forming"])

    # Two Remove-tier + one real Group-B-tier zip: the fold gives Group B
    # exactly 3, which DOES clear the floor.
    attribution2 = _export(0.01, [
        _row("20001", 5000, 0.001), _row("20002", 5000, 0.001),   # Remove (0.1x)
        _row("20003", 5000, 0.007)])                                # Group B (0.7x)
    result2 = ra.zip_optimization_groups(attribution2, "high", prior_periods=ONE_PRIOR)
    equal("  exactly one group forms", len(result2["groups"]), 1)
    group = result2["groups"][0]
    equal("  it's tier 'b' (Remove's own 2 folded IN, not the reverse)", group["tier"], "b")
    equal("  all three zips are in it", set(group["zips"]), {"20001", "20002", "20003"})
    equal("  count matches", group["count"], 3)


def check_impression_floor_and_min_share():
    print("\nImpression floor + ZIP_MIN_SHARE: a low-volume zip is excluded outright, "
         "never a group member, never forming, never already_limited")
    rows = _three_remove_zips() + [_row("20004", 50, 0.001)]  # below the 200 floor
    attribution = _export(0.01, rows)
    result = ra.zip_optimization_groups(attribution, "high", prior_periods=ONE_PRIOR)
    all_seen = ({z for g in result["groups"] for z in g["zips"]}
               | {e["value"] for e in result["forming"]}
               | {e["value"] for e in result.get("already_limited") or []})
    check("  the below-floor zip appears nowhere at all", "20004" not in all_seen, all_seen)
    check("  the three real-volume zips still form their group",
         any(set(g["zips"]) == {"20001", "20002", "20003"} for g in result["groups"]), result["groups"])


def check_already_limited_not_rediscovered():
    print("\nA zip already suppressed in the plan is reported, not rediscovered as a cut")
    rows = _three_remove_zips() + [_row("20009", 1000, 0.001)]  # under, but 1/4 of peer avg
    attribution = _export(0.01, rows)
    result = ra.zip_optimization_groups(attribution, "high", prior_periods=ONE_PRIOR)
    limited_values = {e["value"] for e in result.get("already_limited") or []}
    check("  20009 is reported as already-limited", "20009" in limited_values, limited_values)
    all_group_zips = {z for g in result["groups"] for z in g["zips"]}
    check("  20009 is NOT a fresh group member (never recommended twice)",
         "20009" not in all_group_zips, all_group_zips)
    check("  the other three still form their own real group",
         any(set(g["zips"]) == {"20001", "20002", "20003"} for g in result["groups"]), result["groups"])


def check_wash_excluded_entirely():
    print("\nA zip that flips sides across periods is a WASH -- excluded from groups AND "
         "forming, not just from the group (_zip_group_consistency's own binary model: "
         "unlike the flat engine, ANY opposite-side period is a wash, no minority-that-"
         "isn't-a-wash case exists for zip)")
    rows = _three_remove_zips() + [_row("20005", 5000, 0.001)]  # Remove this period
    attribution = _export(0.01, rows)
    washed_prior = {"period_start": "2026-01-01", "period_end": "2026-01-31",
                    "campaign_rate": 0.01, "zip": [{"label": "20005", "rate": 0.02}]}  # Group A that period
    result = ra.zip_optimization_groups(attribution, "high", prior_periods=ONE_PRIOR,
                                        series_period_facts=[washed_prior])
    all_seen = ({z for g in result["groups"] for z in g["zips"]}
               | {e["value"] for e in result["forming"]})
    check("  20005 (a wash) appears nowhere -- not a group member, not forming",
         "20005" not in all_seen, all_seen)
    check("  the three consistent zips still form their group",
         any(set(g["zips"]) == {"20001", "20002", "20003"} for g in result["groups"]), result["groups"])


def check_in_effect_exclusion():
    print("\nA zip already accepted in a prior report is excluded from candidacy entirely")
    rows = _three_remove_zips() + [_row("20006", 5000, 0.001)]
    attribution = _export(0.01, rows)
    in_effect = {("zip", "20006")}
    result = ra.zip_optimization_groups(attribution, "high", prior_periods=ONE_PRIOR,
                                        in_effect_values=in_effect)
    all_seen = ({z for g in result["groups"] for z in g["zips"]}
               | {e["value"] for e in result["forming"]})
    check("  20006 (in effect) appears nowhere this month", "20006" not in all_seen, all_seen)
    check("  the three never-excluded zips still form their group",
         any(set(g["zips"]) == {"20001", "20002", "20003"} for g in result["groups"]), result["groups"])


def check_no_rows_or_no_baseline():
    print("\nNo rows, or no campaign baseline -- degrades to the empty shape, never crashes")
    empty = _export(0.01, [])
    result = ra.zip_optimization_groups(empty, "high", prior_periods=ONE_PRIOR)
    equal("  no rows -> zero groups", result["groups"], [])
    equal("  no rows -> zero forming", result["forming"], [])

    no_baseline = _export(0.0, _three_remove_zips())
    result2 = ra.zip_optimization_groups(no_baseline, "high", prior_periods=ONE_PRIOR)
    equal("  zero baseline rate -> zero groups (division avoided, not a crash)",
         result2["groups"], [])


def check_build_and_rebuild_zip_group():
    print("\n_build_zip_group / rebuild_zip_group: combined stats are impression-weighted, "
         "never an average of the individual rates")
    entries = [
        {"value": "20001", "delivered_impressions": 8000, "attributed_impressions": 80,
         "campaign_rate": 0.01},
        {"value": "20002", "delivered_impressions": 2000, "attributed_impressions": 2,
         "campaign_rate": 0.01},
    ]
    group = ra._build_zip_group("remove", entries, total_delivered=20000)
    equal("  combined_share = 10000/20000", group["combined_share"], 0.5)
    # (80+2) attributed / (8000+2000) delivered = 82/10000 = 0.0082 --
    # NOT the simple average of 0.01 and 0.001, which would be 0.0055.
    check("  combined_rate is impression-weighted, not row-averaged",
         abs(group["combined_rate"] - 0.0082) < 1e-9, group["combined_rate"])
    equal("  label", group["label"], "Remove")
    equal("  count", group["count"], 2)

    rebuilt = ra.rebuild_zip_group(group, kept_zips=["20001"])
    equal("  rebuild with one zip dropped -- zips list shrinks", rebuilt["zips"], ["20001"])
    equal("  combined_rate recomputes from the SURVIVING row alone", rebuilt["combined_rate"], 0.01)
    equal("  combined_share recomputes too (8000/20000)", rebuilt["combined_share"], 0.4)

    check("  dropping every zip returns None -- nothing left to recommend",
         ra.rebuild_zip_group(group, kept_zips=[]) is None, None)


def check_describe_zip_group():
    print("\ndescribe_zip_group: deterministic, no model involved")
    a_group = ra._build_zip_group("a", [
        {"value": "20001", "delivered_impressions": 5000, "attributed_impressions": 60,
         "campaign_rate": 0.01}], total_delivered=10000)
    text_a = ra.describe_zip_group(a_group)
    check("  Group A uses the 'Shift weight toward' verb", text_a.startswith("Shift weight toward"), text_a)
    check("  names the real zip", "20001" in text_a, text_a)

    b_group = ra._build_zip_group("b", [
        {"value": "20002", "delivered_impressions": 5000, "attributed_impressions": 30,
         "campaign_rate": 0.01}], total_delivered=10000)
    text_b = ra.describe_zip_group(b_group)
    check("  Group B uses the 'Restrict' verb", text_b.startswith("Restrict"), text_b)

    remove_group = ra._build_zip_group("remove", [
        {"value": "20003", "delivered_impressions": 5000, "attributed_impressions": 5,
         "campaign_rate": 0.01},
        {"value": "20004", "delivered_impressions": 5000, "attributed_impressions": 5,
         "campaign_rate": 0.01}], total_delivered=10000)
    text_remove = ra.describe_zip_group(remove_group)
    check("  Remove's subject is '{count} ZIPs', never 'Remove Remove (...)' -- the real bug "
         "found while building this (the verb and the tier's own label are both 'Remove')",
         "Remove Remove" not in text_remove and text_remove.startswith("Remove 2 ZIPs"), text_remove)
    check("  ends 'from the current Premion Streaming TV plan' (removal), not 'within' (a/b)",
         text_remove.rstrip(".").endswith("from the current Premion Streaming TV plan"), text_remove)
    check("  Group A/B end 'within', not 'from'",
         text_a.rstrip(".").endswith("within the current Premion Streaming TV plan")
         and text_b.rstrip(".").endswith("within the current Premion Streaming TV plan"),
         (text_a, text_b))


def check_real_fixtures_dont_crash():
    print("\nReal fixtures (if present): the group engine runs clean end to end")
    import attribution_import as ai
    for name, path in (("MW", REPO / "MW attribution excel.xlsx"),
                       ("Cardinal", REPO / "Premion Website Attribution Cardinal.xlsx"),
                       ("WAEPA", REPO / "Premion Website Attribution and Reach Extension (14).xlsx")):
        if not path.exists():
            print(f"  SKIP  {name} fixture not present")
            continue
        attribution = ai.parse_attribution_export(str(path))
        for level in ("low", "moderate", "high"):
            result = ra.zip_optimization_groups(attribution, level, prior_periods=ONE_PRIOR)
            for group in result["groups"]:
                check(f"  {name}/{level}: group '{group['tier']}' clears the minimum size",
                     group["count"] >= ra.ZIP_GROUP_MIN_SIZE, group)
                check(f"  {name}/{level}: group '{group['tier']}' tier matches its own real label",
                     group["label"] == ra._ZIP_GROUP_LABELS[group["tier"]], group)
                # describe_zip_group must never raise on real data.
                ra.describe_zip_group(group)


def main():
    check_tier_thresholds()
    check_timing_gate()
    check_none_level()
    check_min_size_fold()
    check_impression_floor_and_min_share()
    check_already_limited_not_rediscovered()
    check_wash_excluded_entirely()
    check_in_effect_exclusion()
    check_no_rows_or_no_baseline()
    check_build_and_rebuild_zip_group()
    check_describe_zip_group()
    check_real_fixtures_dont_crash()

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
