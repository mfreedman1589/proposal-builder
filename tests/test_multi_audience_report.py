"""Attribution report fixes from the multi-audience bank review (2026-11).

    python tests/test_multi_audience_report.py

Synthetic exports only (CI); tests/test_multi_audience_real.py checks the
real fixture locally. One section per review item:

  1. Spillover: a market is spillover only under 2% of impressions AND
     outside the targeting -- the old 10% floor dropped the campaign's top
     market -- and the narrative's market count is the table's.
  9. The Geography tile counts places: an audience's priority tiers ("High
     Priority Regions") aren't markets.
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
from attribution_import import AttributionRow  # noqa: E402

app.install_market_lookup()
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:300]}" if detail else ""))


def row(label, delivered, rate):
    return AttributionRow(label, delivered, round(delivered * rate), rate)


# Shares: Norfolk 40%, Richmond 30%, Charlottesville 9% (the TOP rate),
# DC 15%, Harrisonburg 4.5%, Salisbury 1.5% (outside the targeting).
MARKETS = [row("NORFOLK-PORTSMTH-NEWPT NWS", 400_000, 0.055),
           row("RICHMOND-PETERSBURG", 300_000, 0.042),
           row("CHARLOTTESVILLE", 90_000, 0.081),
           row("WASHINGTON, DC (HAGRSTWN)", 150_000, 0.013),
           row("HARRISONBURG", 45_000, 0.03),
           row("SALISBURY", 15_000, 0.02)]
TARGETED_ZIPS = {"23510": 1, "23219": 1, "22901": 1, "20001": 1, "22801": 1}   # no Salisbury ZIP


def attribution(markets=MARKETS):
    total = sum(m.delivered_impressions for m in markets)
    attributed = sum(m.attributed_impressions for m in markets)
    return ai.AttributionExport(
        delivered_impressions=total, attributed_impressions=attributed,
        attributed_rate=attributed / total, attributed_unique_visitors=5_000,
        by_market=list(markets),
        by_audience=[row("FIN Banking Intenders", total, attributed / total)])


def delivery(zip_option=True, geos=None):
    return ai.DeliveryExport(
        delivered_impressions=sum(m.delivered_impressions for m in MARKETS),
        by_geo=geos or [("Central Virginia", 500_000), ("Hampton Roads", 500_000)],
        zip_option_geos=zip_option, delivery_map=dict(TARGETED_ZIPS))


def check_spillover():
    print("1. Spillover: under 2% AND outside the targeting")
    a, d = attribution(), delivery()
    targeted = ra.targeted_markets(None, d)
    check("a ZIP-option delivery file's delivered ZIPs give the targeted markets",
          targeted and "charlottesville" in targeted and not any("salisbury" in k for k in targeted),
          targeted)
    shown = [r.label for r in ra.campaign_market_rows(a, targeted)]
    check("the top-rate market at 9% is never spillover", "CHARLOTTESVILLE" in shown, shown)
    check("a 4.5% market is never spillover", "HARRISONBURG" in shown, shown)
    check("a market under 2% outside the targeting is spillover",
          [r.label for r in ra.spillover_market_rows(a, targeted)] == ["SALISBURY"])
    d_in = ai.DeliveryExport(by_geo=d.by_geo, zip_option_geos=True,
                             delivery_map={**TARGETED_ZIPS, "21801": 1})
    check("...but under 2% INSIDE the targeting is a market",
          not ra.spillover_market_rows(a, ra.targeted_markets(None, d_in)))
    check("with no targeting known, under 2% alone decides",
          [r.label for r in ra.spillover_market_rows(a, None)] == ["SALISBURY"])
    check("a linked proposal's ZIPs win over the delivery file",
          ra.targeted_markets(["21801"], d) and any("salisbury" in k
                                                    for k in ra.targeted_markets(["21801"], d)))
    check("DMA-targeted delivery (not ZIP options) gives no targeting from delivered ZIPs",
          ra.targeted_markets(None, delivery(zip_option=False)) is None)

    facts = ra.build_facts_payload(a, d, goals=["Grow deposits"])
    dim, table, _ = ra.attribution_breakdown_display(ra.with_targeting(a, None, d), delivery=d)
    check("the breakdown is by market", dim == "Market", dim)
    check("the narrative's market count is the table's row count",
          facts["market"]["count"] == len(table) == 5, (facts["market"]["count"], len(table)))
    check("the table lists every non-spillover market, the top-rate one included",
          sorted(r["label"] for r in table) == sorted(
              ra._dma_display_name(r.label) for r in ra.campaign_market_rows(a, targeted)),
          [r["label"] for r in table])
    check("the facts' spillover list is the one market", facts["market"]["spillover"] == ["Salisbury"],
          facts["market"]["spillover"])
    eight = attribution([row(f"MARKET {chr(65 + i)}", 100_000, 0.03) for i in range(8)])
    check("eight real markets all fit the table",
          len(ra.attribution_breakdown_display(eight)[1]) == 8)


def check_geography_tile():
    print("\n9. The Geography tile counts places, not targeting tiers")
    tiers = delivery(geos=[("High Priority Regions", 600_000), ("Low Priority Regions", 200_000),
                           ("Central Virginia", 160_000), ("Piedmont", 140_000)])
    check("priority tiers aren't bought places", ra.bought_geo_labels(tiers) is None,
          ra.bought_geo_labels(tiers))
    a = attribution()
    check("the tile falls back to the export's own markets, the table's count",
          ra.geography_label(ra.with_targeting(a, None, tiers)) == "5 markets",
          ra.geography_label(ra.with_targeting(a, None, tiers)))
    named = delivery(geos=[("Springfield Zips", 1), ("Bethesda Zips", 1)])
    check("named place lists still label the tile", ra.bought_geo_labels(named) is not None)


def main():
    check_spillover()
    check_geography_tile()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
