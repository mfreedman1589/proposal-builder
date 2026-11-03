"""Synthetic live-sports workbooks, shaped like the real exports but with
invented clients, RFPIDs and figures -- what CI tests the live-sports parsing
against (the real files are client data and stay local, in
tests/fixtures/live_sports/).

Two delivery formats:
  * "2026-09": the older layout -- an "IMPRESSION BY EVENT" tab with a
    League column, "VCR %" everywhere, "Delivered Geo" on the details tab.
  * "2026-11": the newer one -- no League, bare "VCR" (so the sports KPI tab
    has the OTT KPI tab's exact header), client/RFPID columns on the event
    tab, several RFPIDs, and geography on a "Markets" tab.
"""
from datetime import datetime

import openpyxl

OTT_TOTAL = 400_000
OTT_VCR = 0.981

# (rfpid, date, network, event, delivered, completed) -- one game simulcast
# on two networks and bought under two RFPIDs, and two generic "NFL" rows on
# one date (different broadcasts the export didn't name).
NEW_EVENTS = [
    ("RFPID-900001", datetime(2026, 10, 19), "ESPN", "Team Alpha - Team Bravo", 30_000, 29_400),
    ("RFPID-900001", datetime(2026, 10, 19), "ABC", "Team Alpha - Team Bravo", 20_000, 19_600),
    ("RFPID-900002", datetime(2026, 10, 19), "ESPN", "Team Alpha - Team Bravo", 15_000, 14_700),
    ("RFPID-900001", datetime(2026, 10, 18), "NBC", "Team Charlie - Team Delta", 40_000, 39_400),
    ("RFPID-900002", datetime(2026, 10, 18), "NBC", "Team Charlie - Team Delta", 5_000, 4_900),
    ("RFPID-900001", datetime(2026, 10, 25), "NBC", "NFL", 25_000, 24_800),
    ("RFPID-900001", datetime(2026, 10, 25), "NFLTV", "NFL", 24_000, 23_900),
    ("RFPID-900002", datetime(2026, 9, 28), "ESPN", "Team Echo - Team Foxtrot", 1_000, 980),
]
NEW_SPORTS_TOTAL = sum(e[4] for e in NEW_EVENTS)          # 160,000
NEW_SPORTS_COMPLETED = sum(e[5] for e in NEW_EVENTS)

OLD_EVENTS = [
    (datetime(2026, 9, 3), "College Football", "ESPN", "State U - Tech U", 20_000, 19_600),
    (datetime(2026, 9, 4), "College Football", "BTN", "North U - South U", 15_000, 14_800),
]
OLD_SPORTS_TOTAL = sum(e[4] for e in OLD_EVENTS)


def _sheet(wb, title, header, rows):
    ws = wb.create_sheet(title)
    ws.append(list(header))
    for row in rows:
        ws.append(list(row))
    return ws


def _ott_tabs(wb):
    _sheet(wb, "KPI DELIVERY", ("Delivered Impressions", "VCR"), [(OTT_TOTAL, OTT_VCR)])
    _sheet(wb, "KPI PERFORMANCE", ("Frequency", "Uniques"), [(4.2, 95_000)])
    _sheet(wb, "IMPRESSIONS BY PRODUCT PACKAGE", ("Package Type", "Delivered Impressions"),
           [("PREM TV - RON", OTT_TOTAL)])


def delivery_2026_11(path, sports_kpi_first=False, break_event_tab=False):
    """The newer format. `sports_kpi_first` puts the sports KPI tab ahead of
    the OTT one (identical headers -- the OTT figure must still be the OTT
    figure). `break_event_tab` drops a column the parser needs, so the event
    tab exists but can't be read."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sports_kpi = (("Delivered Impressions", "VCR"),
                  [(NEW_SPORTS_TOTAL, round(NEW_SPORTS_COMPLETED / NEW_SPORTS_TOTAL, 4))])
    if sports_kpi_first:
        _sheet(wb, "﻿KPI DELIVERY﻿", *sports_kpi)
    _ott_tabs(wb)
    if not sports_kpi_first:
        _sheet(wb, "﻿KPI DELIVERY﻿", *sports_kpi)
    by_rfpid = {}
    for rfpid, _d, _n, _e, delivered, completed in NEW_EVENTS:
        tot = by_rfpid.setdefault(rfpid, [0, 0])
        tot[0] += delivered
        tot[1] += completed
    _sheet(wb, "DETAILS BY CAMPAIGN",
           ("Client", "RFPID", "RFPI", "Package Type", "Flight Start Date", "Flight End Date",
            "Flight Goal", "Delivered Impressions", "Completed Impressions", "VCR"),
           [("Example Client", "RFPID-900001", "RFPI -1", "PREM TV - LIVE NFL REGULAR SEASON",
             datetime(2026, 9, 13), datetime(2026, 10, 29), 120_000, *by_rfpid["RFPID-900001"], 0.98),
            ("Example Client", "RFPID-900002", "RFPI -2", "PREM TV - LIVE NFL REGULAR SEASON",
             datetime(2026, 9, 10), datetime(2026, 10, 31), 40_000, *by_rfpid["RFPID-900002"], 0.98)])
    header = ["Client", "RFPID", "RFPI", "Package Type", "Date", "Network", "Event",
              "Delivered Impressions", "Completed Impressions", "VCR"]
    if break_event_tab:
        header[header.index("Completed Impressions")] = "Completes"
    _sheet(wb, "IMPRESSIONS BY EVENT", header,
           [("Example Client", rfpid, "RFPI -1", "PREM TV - LIVE NFL REGULAR SEASON", day, net,
             event, delivered, completed, round(completed / delivered, 4))
            for rfpid, day, net, event, delivered, completed in NEW_EVENTS])
    _sheet(wb, "IMPRESSIONS BY NETWORK", ("Network", "Delivered Impressions"),
           [("ESPN", 46_000), ("NBC", 70_000), ("ABC", 20_000), ("NFLTV", 24_000)])
    _sheet(wb, "IMPRESSIONS BY GEO", ("Markets", "Delivered Impressions"),
           [("Washington DC (Hagerstown MD)(511)", NEW_SPORTS_TOTAL)])
    wb.save(path)
    return path


def delivery_2026_09(path):
    """The older format (League column, "VCR %", "Delivered Geo")."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _ott_tabs(wb)
    completed = sum(e[5] for e in OLD_EVENTS)
    _sheet(wb, "﻿KPI DELIVERY﻿", ("Delivered Impressions", "VCR %"),
           [(OLD_SPORTS_TOTAL, round(completed / OLD_SPORTS_TOTAL, 4))])
    _sheet(wb, "DETAILS BY CAMPAIGN",
           ("Client", "RFPID", "RFPI", "Package Type", "Delivered Geo", "Flight Start Date",
            "Flight End Date", "Flight Goal", "Delivered Impressions", "Completed Impressions",
            "VCR %"),
           [("Example College", "RFPID-800001", "RFPI -9", "PREM TV - LIVE NCAAF REGULAR SEASON",
             "Washington DC (Hagerstown MD)(511)", datetime(2026, 8, 29), datetime(2026, 12, 12),
             300_000, OLD_SPORTS_TOTAL, completed, 0.98)])
    _sheet(wb, "IMPRESSION BY EVENT",
           ("Date", "League", "Network", "Event", "Delivered Impressions",
            "Completed Impressions", "VCR %"),
           [(day, league, net, event, delivered, comp, round(comp / delivered, 4))
            for day, league, net, event, delivered, comp in OLD_EVENTS])
    _sheet(wb, "IMPRESSIONS BY NETWORK", ("Network", "Delivered Impressions"),
           [("ESPN", 20_000), ("BTN", 15_000)])
    _sheet(wb, "IMPRESSIONS BY LEAGUE", ("League", "Delivered Impressions"),
           [("College Football", OLD_SPORTS_TOTAL)])
    wb.save(path)
    return path
