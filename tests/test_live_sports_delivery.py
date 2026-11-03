"""Live-sports detection inside a delivery workbook, both export formats.

    python tests/test_live_sports_delivery.py

The 2026-11 export changed the sports block's headers (no League, "VCR"
instead of "VCR %", no Delivered Geo, geography on a "Markets" tab), and
detection failed silently: live_sports came back None with no warning. These
run against synthetic workbooks (tests/live_sports_fixtures.py) so CI covers
them; tests/test_live_sports_real.py checks the real files locally.
"""
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))

import attribution_import as ai  # noqa: E402
import live_sports_fixtures as fx  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {detail}" if detail else ""))


def main():
    tmp = Path(tempfile.mkdtemp())

    print("2026-11 format: no League, bare VCR, several RFPIDs, a Markets tab")
    d = ai.parse_delivery_export(str(fx.delivery_2026_11(tmp / "new.xlsx")))
    ls = d.live_sports
    check("the sports block is detected", ls is not None, d.warnings)
    if ls is None:
        return finish()
    check("no warning on a readable file", not d.warnings, d.warnings)
    check("sports delivered is the details total", ls.delivered_impressions == fx.NEW_SPORTS_TOTAL,
          ls.delivered_impressions)
    check("sports VCR is completed / delivered",
          abs(ls.vcr - fx.NEW_SPORTS_COMPLETED / fx.NEW_SPORTS_TOTAL) < 1e-9, ls.vcr)
    check("OTT delivered is the OTT KPI tab, not the sports one",
          d.delivered_impressions == fx.OTT_TOTAL and abs(d.vcr - fx.OTT_VCR) < 1e-9,
          (d.delivered_impressions, d.vcr))
    check("every RFPID is kept", ls.rfpids == ["RFPID-900001", "RFPID-900002"], ls.rfpids)
    check("the flight spans earliest start to latest end",
          (ls.flight_start, ls.flight_end) == (date(2026, 9, 10), date(2026, 10, 31)),
          (ls.flight_start, ls.flight_end))
    check("geography comes from the Markets tab",
          ls.delivered_geo == "Washington DC (Hagerstown MD)(511)", ls.delivered_geo)
    check("league is optional (blank, not an error)", all(e.league == "" for e in ls.events))
    check("each event row keeps its RFPID", {e.rfpid for e in ls.events} == set(ls.rfpids))

    print("\nThe two KPI tabs share a header; tab order can't swap them")
    d2 = ai.parse_delivery_export(str(fx.delivery_2026_11(tmp / "first.xlsx", sports_kpi_first=True)))
    check("OTT delivered is still the OTT figure with the sports KPI tab first",
          d2.delivered_impressions == fx.OTT_TOTAL, d2.delivered_impressions)
    check("...and sports is still the sports figure",
          d2.live_sports and d2.live_sports.delivered_impressions == fx.NEW_SPORTS_TOTAL)

    print("\nTop games aggregate across RFPIDs and networks")
    games = ls.games()
    top = games[0]
    check("a game simulcast on two networks under two RFPIDs is one row, summed",
          top.event == "Team Alpha - Team Bravo" and top.delivered_impressions == 65_000,
          (top.event, top.delivered_impressions))
    check("...listing both networks, largest first", top.network == "ESPN / ABC", top.network)
    check("...and both RFPIDs", top.rfpid == "RFPID-900001 + RFPID-900002", top.rfpid)
    check("its VCR is recomputed from the summed completes",
          abs(top.vcr - (29_400 + 19_600 + 14_700) / 65_000) < 1e-9, top.vcr)
    check("the same matchup on another date is its own game",
          any(g.event == "Team Charlie - Team Delta" and g.delivered_impressions == 45_000
              for g in games))
    generic = [g for g in games if g.event == "NFL"]
    check("generic 'NFL' rows stay as listed, never merged",
          sorted(g.delivered_impressions for g in generic) == [24_000, 25_000],
          [(g.network, g.delivered_impressions) for g in generic])
    check("aggregation keeps the total", sum(g.delivered_impressions for g in games) == fx.NEW_SPORTS_TOTAL)
    check("ranked by impressions, largest first",
          [g.delivered_impressions for g in games] ==
          sorted((g.delivered_impressions for g in games), reverse=True))

    print("\n2026-09 format: League, 'VCR %', Delivered Geo -- still read")
    d3 = ai.parse_delivery_export(str(fx.delivery_2026_09(tmp / "old.xlsx")))
    ls3 = d3.live_sports
    check("the sports block is detected", ls3 is not None, d3.warnings)
    if ls3 is not None:
        check("delivered and RFPID as before",
              ls3.delivered_impressions == fx.OLD_SPORTS_TOTAL and ls3.rfpid == "RFPID-800001",
              (ls3.delivered_impressions, ls3.rfpid))
        check("league and delivered geo still read",
              ls3.events[0].league == "College Football"
              and ls3.delivered_geo == "Washington DC (Hagerstown MD)(511)",
              (ls3.events[0].league, ls3.delivered_geo))
        check("per-event VCR read from 'VCR %'", ls3.events[0].vcr > 0.9, ls3.events[0].vcr)
        check("OTT figures untouched", d3.delivered_impressions == fx.OTT_TOTAL, d3.delivered_impressions)

    print("\nAn event tab that can't be read is reported, never dropped silently")
    d4 = ai.parse_delivery_export(str(fx.delivery_2026_11(tmp / "broken.xlsx", break_event_tab=True)))
    check("no sports block from an unreadable event tab", d4.live_sports is None)
    check("...and the rep is told, naming the tab",
          any("live sports" in w and "IMPRESSIONS BY EVENT" in w for w in d4.warnings), d4.warnings)
    check("OTT figures still parse", d4.delivered_impressions == fx.OTT_TOTAL, d4.delivered_impressions)
    return finish()


def finish():
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
