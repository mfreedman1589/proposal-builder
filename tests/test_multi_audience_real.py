"""The multi-audience bank review fixes against the real exports (local only:
client data, in the gitignored tests/fixtures/multi_audience/). The synthetic
versions CI runs are tests/test_multi_audience_report.py.

    python tests/test_multi_audience_real.py

Expected values are the QA answer key's for "the multi-audience bank fixture
(2026-11)" -- the figures the rep's own corrected deck shows.
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import app  # noqa: E402
import attribution_import as ai  # noqa: E402
import report_assembly as ra  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures" / "multi_audience"
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:300]}" if detail else ""))


def main():
    paths = [FIXTURES / n for n in ("attribution.xlsx", "delivery.xlsx", "retargeting.xlsx")]
    if not all(p.exists() for p in paths):
        print("SKIP -- the multi-audience fixtures aren't present")
        return 0
    app.install_market_lookup()
    a = ai.parse_attribution_export(str(paths[0]))
    d = ai.parse_delivery_export(str(paths[1]))
    o = ai.parse_ott_retargeting_export(str(paths[2]))
    t = ra.with_targeting(a, None, d)

    print("1/9. Markets, spillover and the Geography tile")
    shown = [r.label for r in ra.campaign_market_rows(t)]
    check("all 7 markets are markets the campaign ran in", len(shown) == 7, shown)
    check("Charlottesville (126,780, the top rate) is never spillover", "CHARLOTTESVILLE" in shown)
    check("nothing is spillover", not ra.spillover_market_rows(t))
    check("the tile reads 7 markets, the table's count", ra.geography_label(t) == "7 markets",
          ra.geography_label(t))
    facts = ra.build_facts_payload(a, d, goals=["Grow deposits"], ott=o)
    check("the narrative's market count is 7", facts["market"]["count"] == 7, facts["market"]["count"])

    print("\n2/3/4. Landing pages and key pages")
    suggested = ra.suggest_landing_pages(a)
    check("the two High Interest offer pages are suggested",
          [ra._page_key(u)[1] for u in suggested] ==
          ["/high-interest-business-accounts", "/high-interest-personal-accounts"], suggested)
    lp = [{"pattern": suggested[0], "audience": a.by_audience[0].label},
          {"pattern": suggested[1], "audience": a.by_audience[1].label}]
    kp = ra.key_pages_facts(ra.with_landing_pages(a, lp))
    by = {r["group"]: r["page_visits"] for r in kp["rows"]}
    check("offers 2,670 and 975", by.get("Offer: High Interest Business Accounts") == 2670
          and by.get("Offer: High Interest Personal Accounts") == 975, by)
    check("homepage 1,344, confirmations 13", by.get("Homepage") == 1344
          and by.get("Confirmations") == 13, by)
    check("personal and business research split", by.get("Personal product research") == 860
          and by.get("Business product research") == 199, by)
    check("ready to act 411, existing customers 198 (last)", by.get("Ready to act") == 411
          and kp["rows"][-1]["group"] == "Existing customer service"
          and by["Existing customer service"] == 198, by)
    path = ra.goal_path_combined(ra.with_landing_pages(a, lp), ["Grow deposits"])
    check("the goal path leads with the landing pages: 3,645 page visits",
          path[0]["key"] == "offer" and path[0]["page_visits"] == 3645, path[0]["page_visits"])

    print("\n5. Audiences side by side")
    f = ra.audience_comparison_facts(ra.with_landing_pages(a, lp), o)
    u = {x["rfpid"]: x for x in f["units"]}
    retail, commercial = u["RFPID-266578"], u["RFPID-267661"]
    check("retail: 551,545 / 27,924 / 5.06%",
          (retail["delivered"], retail["attributed"], round(retail["rate"] * 100, 2))
          == (551_545, 27_924, 5.06))
    check("commercial: 823,824 / 29,399 / 3.57%",
          (commercial["delivered"], commercial["attributed"], round(commercial["rate"] * 100, 2))
          == (823_824, 29_399, 3.57))
    check("creative ranges 4.81–5.27% and 3.12–3.67%",
          tuple(round(x * 100, 2) for x in retail["creative_range"]) == (4.81, 5.27)
          and tuple(round(x * 100, 2) for x in commercial["creative_range"]) == (3.12, 3.67))
    check("retargeting click-through 0.116% and 0.118%",
          (round(retail["retargeting_ctr"] * 100, 3), round(commercial["retargeting_ctr"] * 100, 3))
          == (0.116, 0.118))
    check("shares: retail 40% of impressions, 49% of attributed",
          (round(retail["impression_share"] * 100), round(retail["attributed_share"] * 100)) == (40, 49))
    check("1.4x, and the own-creative caveat", "1.4x" in f["headline"] and f["creatives_separate"])

    print("\n6. Market recommendations")
    found = ra.optimization_candidates(a, "moderate", ra.OPTIMIZATION_DIMENSIONS)
    market = [e["value"] for v in found.values() if isinstance(v, list) for e in v
              if isinstance(e, dict) and e.get("dimension") == "market"]
    check("no market is cut against the blended average (no proposal on file)", not market, market)
    draft = {"threads": [{"head": "x", "finding": "x", "action": (
        "Optimize Zip Codes by concentrating spend on higher performing areas: reduce investment in "
        "Greensboro-High Point-Winston Salem and Washington, DC, both of which trail the campaign "
        "average by more than 3x and reallocate those impressions toward Norfolk and Richmond.")}]}
    check("the original takeaway is caught", len(app._market_mix_violations(draft, facts)) == 1)

    print("\n7. Export date")
    e = ra.export_date_facts(a, d)
    check("exported Nov 4, 2026; still accruing", e["label"] == "Nov 4, 2026"
          and not e["accrual_complete"], e)
    check("the last two weeks are still accruing", e["accruing_weeks"] == ["Oct 18–24", "Oct 25–31"],
          e["accruing_weeks"])

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
