"""Live sports against the real 2026-11 exports (local only: client data, in
the gitignored tests/fixtures/live_sports/). The synthetic versions CI runs
are tests/test_live_sports_delivery.py.

    python tests/test_live_sports_real.py

Expected values are the ones the QA answer key gives for this fixture.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import attribution_import as ai  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures" / "live_sports"
DELIVERY = FIXTURES / "delivery.xlsx"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {detail}" if detail else ""))


def check_delivery():
    print("Delivery export, 2026-11 format")
    if not DELIVERY.exists():
        print(f"  SKIP  {DELIVERY.name} not present")
        return
    d = ai.parse_delivery_export(str(DELIVERY))
    ls = d.live_sports
    check("the sports block is detected, with no warning", ls is not None and not d.warnings,
          d.warnings)
    if ls is None:
        return
    check("543,301 sports impressions delivered", ls.delivered_impressions == 543_301,
          ls.delivered_impressions)
    check("98.5% sports VCR", round(ls.vcr * 100, 1) == 98.5, ls.vcr)
    check("two RFPIDs", ls.rfpids == ["RFPID-268085", "RFPID-257034"], ls.rfpids)
    check("OTT delivered is the OTT figure (601,561), not the sports one",
          d.delivered_impressions == 601_561, d.delivered_impressions)
    top = ls.games()[0]
    check("49ers v. Commanders is the top game after aggregation",
          top.event == "San Francisco 49ers - Washington Commanders"
          and top.delivered_impressions == 87_307, (top.event, top.delivered_impressions))
    check("...across ESPN and ABC", top.network == "ESPN / ABC", top.network)
    check("the 30 event rows become 18 games", (len(ls.events), len(ls.games())) == (30, 18),
          (len(ls.events), len(ls.games())))


def check_attribution():
    """The QA answer key's values for this fixture (qa/QA_RULES.md, the
    live-sports fixture: credit union, NFL, 2026-11)."""
    print("\nAttribution: sports, standard, and the two together")
    paths = [FIXTURES / n for n in ("attribution_sports.xlsx", "attribution_standard.xlsx",
                                    "delivery.xlsx")]
    if not all(p.exists() for p in paths):
        print("  SKIP  the live-sports attribution fixtures aren't present")
        return
    import report_assembly as ra
    sports = ai.parse_attribution_export(str(paths[0]), slot="sports")
    std = ai.parse_attribution_export(str(paths[1]))
    delivery = ai.parse_delivery_export(str(paths[2]))
    check("sports: 4.81% attributed rate", round(sports.attributed_rate * 100, 2) == 4.81,
          sports.attributed_rate)
    check("sports: 4,511 unique visitors", sports.attributed_unique_visitors == 4_511)
    check("the sports slot raises no multi-RFPID warning", not sports.warnings, sports.warnings)
    check("the sports file's RFPIDs match the delivery file's sports block",
          ai.sports_rfpid_warning(sports, delivery) is None)
    check("each file is in the right slot",
          ai.slot_mismatch_warning(sports, "sports") is None
          and ai.slot_mismatch_warning(std, "standard") is None)
    check("the sports label is 'Live NFL'",
          ai.sports_label_from_package(delivery.live_sports.package_type) == "Live NFL")
    comparison = ra.campaign_comparison_facts(std, ra.sports_view(sports), delivery, "Live NFL")
    check("the two flights are the same period", comparison["same_period"],
          comparison["period_note"])
    by_key = {m["key"]: m for m in comparison["measures"]}
    check("matched period: standard 8.56% attributed rate",
          by_key["attributed_rate"]["standard_display"] == "8.56%", by_key["attributed_rate"])
    check("matched period: unique visitor rate 0.81% vs sports 0.83%",
          (by_key["unique_visitor_rate"]["standard_display"],
           by_key["unique_visitor_rate"]["sports_display"]) == ("0.81%", "0.83%"),
          by_key.get("unique_visitor_rate"))
    summary = {r["metric"]: r for r in ra.campaign_summary_facts(
        std, ra.sports_view(sports), delivery, "Live NFL")["rows"]}
    check("combined delivered impressions 1,144,862",
          summary["Delivered impressions"]["combined"] == 1_144_862)
    check("unique visitors never combined", summary["Unique website visitors"]["combined"] is None)


def main():
    check_delivery()
    check_attribution()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
