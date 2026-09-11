"""Premion Website Attribution benchmarks -- transcribed ONCE from the real
workbook (`PremionWebsiteAttribution-Benchmarks 2025 (1).xlsx`, INTERNAL USE
ONLY, gitignored like every other real-data workbook in this repo) into this
committed, pure Python data file. Never parsed at runtime -- the sheet is a
marketing spreadsheet (repeating blocks per vertical, notes in cells, a
missing unique-visitor-rate cell for Auto), not clean data, and re-parsing it
on every draft would just re-import that mess into the app.

Confirmed directly against the workbook (not a summary of it): 8 vertical
rows exist, not the 6 first assumed. Each row's "Average Attributed
(Impression) Rate" is exactly attributed impressions / delivered impressions
-- the same computation `attribution_import.AttributionExport.attributed_rate`
already is (verified against three real exports: WAEPA 0.13%, MW 1.363%,
Cardinal 0.185% -- see ATTRIBUTION_REPORT_PLAN.md's dated section for the
full arithmetic settling this). "Average Attributed Unique Visitor Rate" is
the same shape against `attributed_unique_visitor_rate` -- attributed unique
visitors over the SAME delivered-impressions denominator, not a different
one. These are two different rates by design; a campaign can clear one and
not the other.

`VERTICAL_TO_BENCHMARK` maps `app.VERTICALS`' own values (the app's internal
vertical keys) to a `BENCHMARKS` row key. Three app verticals have no row and
are deliberately absent rather than approximated to a neighbour (Retail maps
to the sheet's narrower "Furniture Retail" and Entertainment to "Entertainment
and Gambling" only because Matt confirmed those are the same audience
concept, not a neighbour-approximation): Travel & Tourism, Casual Dining &
QSR, and "None" have no benchmark at all.
"""

BENCHMARKS = {
    "auto": {
        "label": "Auto",
        "impression_rate": 0.0411,
        "visitor_rate": None,  # the sheet's own Auto row has no visitor-rate cell
        "accounts_measured": 129,
    },
    "home_services_improvement": {
        "label": "Home Services/Improvement",
        "impression_rate": 0.0356,
        "visitor_rate": 0.0041,
        "accounts_measured": 64,
    },
    "furniture_retail": {
        "label": "Furniture Retail",
        "impression_rate": 0.0576,
        "visitor_rate": 0.0073,
        "accounts_measured": 63,
    },
    "education": {
        "label": "Education",
        "impression_rate": 0.0351,
        "visitor_rate": 0.0043,
        "accounts_measured": 20,
    },
    "banking_and_finance": {
        "label": "Banking and Finance",
        "impression_rate": 0.0481,
        "visitor_rate": 0.0074,
        "accounts_measured": 27,
    },
    "legal": {
        "label": "Legal",
        "impression_rate": 0.0222,
        "visitor_rate": 0.0023,
        "accounts_measured": 25,
    },
    "healthcare": {
        "label": "Healthcare",
        "impression_rate": 0.0350,
        "visitor_rate": 0.0051,
        "accounts_measured": 25,
    },
    "entertainment_and_gambling": {
        "label": "Entertainment and Gambling",
        "impression_rate": 0.0466,
        "visitor_rate": 0.0072,
        "accounts_measured": 27,
    },
}

# app.VERTICALS' own values -> a BENCHMARKS key. Confirmed with the user
# 2026-09-10: Furniture wraps into Retail, Gambling wraps into Entertainment.
# travel/dining_qsr/none are deliberately absent -- no row, no approximation.
VERTICAL_TO_BENCHMARK = {
    "auto": "auto",
    "home_improvement": "home_services_improvement",
    "retail": "furniture_retail",
    "education": "education",
    "banking": "banking_and_finance",
    "legal": "legal",
    "healthcare": "healthcare",
    "entertainment": "entertainment_and_gambling",
}


def benchmark_facts(vertical_key, attributed_rate, attributed_unique_visitor_rate):
    """None when `vertical_key` has no benchmark row, or the campaign clears
    neither rate. Otherwise a dict naming the row's own label plus, ONLY for
    the rate(s) actually cleared, `{"campaign": <the real campaign rate>}` --
    "below benchmark: silence" per row, not per campaign, since a campaign
    can clear one rate and not the other.

    The benchmark row's OWN percentage never appears anywhere in this return
    value -- callers (`report_assembly.build_facts_payload`) must still keep
    the whole "benchmark" subtree out of the facts-only checker's traced-
    number set (see `app._attr_payload_numbers`), since a citable "cleared"
    flag next to the row's real percentage would otherwise let the model
    quote the internal number directly. This function itself simply doesn't
    return the row's raw percentages at all, as a second line of defense.
    """
    row_key = VERTICAL_TO_BENCHMARK.get(vertical_key or "")
    if row_key is None:
        return None
    row = BENCHMARKS[row_key]
    result = {"vertical": row["label"]}
    cleared_any = False
    if attributed_rate is not None and attributed_rate > row["impression_rate"]:
        result["impression_rate"] = {"campaign": attributed_rate, "cleared": True}
        cleared_any = True
    if (row["visitor_rate"] is not None and attributed_unique_visitor_rate is not None
            and attributed_unique_visitor_rate > row["visitor_rate"]):
        result["visitor_rate"] = {"campaign": attributed_unique_visitor_rate, "cleared": True}
        cleared_any = True
    return result if cleared_any else None
