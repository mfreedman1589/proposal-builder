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
    _sheet(wb, "TOP 10 CHANNELS BY IMPRESSIONS", ("Channel Name", "Delivered Impressions"),
           [("Example Network", OTT_TOTAL // 2), ("Another Network", OTT_TOTAL // 2)])


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


# ---------------------------------------------------------------------------
# Attribution exports. Tab titles and headers are the real exports' own
# (format, not client data); every value is invented. A standard export has
# audience, channel and advertiser tabs; a Live Sports one has none of them
# and adds conversions/sales and conversions-by-URL tabs.
# ---------------------------------------------------------------------------
from datetime import date, timedelta  # noqa: E402

_DETAIL = ["Delivered Impressions", "Attributed Impressions", "Attributed Rate",
           "Conversion Impressions", "Conversion Impressions Rate"]
STANDARD = dict(delivered=400_000, attributed=34_000, uniques=3_300,
                start=date(2026, 8, 31), weeks=9, rfpids=["RFPID-700001"],
                devices={"smartphone": 2_000, "desktop": 1_700, "tablet": 80},
                recency={"00 - 03 DAYS": 900, "04 - 07 DAYS": 800, "08 - 11 DAYS": 800,
                         "12 - 15 DAYS": 800},
                referral={"Direct": 1_500, "Organic Search": 1_200, "Social": 100},
                pages={"https://www.example.org/": 2_500, "https://www.example.org/loans/": 700,
                       "https://www.example.org/rates/": 400,
                       "https://www.example.org/contact/": 200},
                zips={"20001": 3_000, "20002": 2_500, "20003": 2_000, "22201": 1_800,
                      "22202": 1_500},
                markets=[("WASHINGTON, DC (HAGRSTWN)", 400_000)],
                creatives=[("TG10001EXCOComm0926", 250_000), ("Example_Brand_30", 150_000)],
                conversions=0, sales=0)
SPORTS = dict(delivered=160_000, attributed=7_700, uniques=1_330,
              start=date(2026, 9, 14), weeks=7, rfpids=["RFPID-900001", "RFPID-900002"],
              devices={"smartphone": 900, "desktop": 450, "tablet": 20},
              recency={"00 - 03 DAYS": 450, "04 - 07 DAYS": 350, "08 - 11 DAYS": 300,
                       "12 - 15 DAYS": 230},
              referral={"Direct": 600, "Organic Search": 500, "Social": 60},
              pages={"https://www.example.org/": 1_000, "https://www.example.org/loans/": 300,
                     "https://www.example.org/branches/": 150,
                     "https://www.example.org/contact/": 90},
              zips={"20001": 900, "20011": 800, "20019": 700, "20782": 600, "20783": 500},
              markets=[("WASHINGTON, DC (HAGRSTWN)", 156_000), ("HARRISONBURG", 2_500),
                       ("BALTIMORE", 1_500)],
              creatives=[("TG10001EXCOComm0926", 130_000), ("Example_Team_30", 30_000)],
              conversions=0, sales=0)


def _rate(a, d):
    return round(a / d, 6) if d else 0.0


def _split(total, n):
    base = [total // n] * n
    base[-1] += total - sum(base)
    return base


def attribution_workbook(path, kind="standard", **overrides):
    """A synthetic attribution export. `kind` is "standard" or "sports"; any
    key of STANDARD/SPORTS can be overridden (start=/weeks= move the flight,
    conversions= makes them non-zero)."""
    f = dict(STANDARD if kind == "standard" else SPORTS)
    f.update(overrides)
    d, a = f["delivered"], f["attributed"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet(wb, "Untitled Widget", ["Delivered Impressions", "Attributed Impressions",
                                   "Attributed Rate"], [(d, a, _rate(a, d))])
    _sheet(wb, "Untitled Widget_1", ["Attributed Unique Visitors",
                                     "Attributed Unique Visitor Rate"],
           [(f["uniques"], _rate(f["uniques"], d))])
    if kind == "sports":
        _sheet(wb, "Untitled Widget_2", ["Attributed Conversions", "Sales Amount"],
               [(f["conversions"], f["sales"])])
    else:
        _sheet(wb, "ATTRIBUTED RATE BY AUDIENCE", ["Audience Name"] + _DETAIL,
               [("SPORTS Local Team Fans,FIN Banking Intenders", d, a, _rate(a, d), 0, 0.0)])
    creative_attr = _split(a, len(f["creatives"]))
    _sheet(wb, "ATTRIBUTED RATE BY CREATIVE", ["Creative Name"] + _DETAIL,
           [(name, dd, ca, _rate(ca, dd), 0, 0.0)
            for (name, dd), ca in zip(f["creatives"], creative_attr)])
    market_attr = [round(a * dd / d) for _, dd in f["markets"]]
    _sheet(wb, "ATTRIBUTED RATE BY MARKET", ["Market"] + _DETAIL,
           [(m, dd, ma, _rate(ma, dd), 0, 0.0) for (m, dd), ma in zip(f["markets"], market_attr)])
    _sheet(wb, "ATTRIBUTED RATE BY ZIP CODE", ["Zipcode"] + _DETAIL,
           [(z, imps * 10, imps, 0.1, 0, 0.0) for z, imps in f["zips"].items()])
    _sheet(wb, "ATTRIBUTED IMPRESSIONS BY ZIP H", ["Attributed Impressions", "Zipcode"],
           [(imps, z) for z, imps in f["zips"].items()])
    _sheet(wb, "ATTRIBUTED UNIQUE PAGE VISITS B", ["Page Url", "Attributed Unique Visitors"],
           list(f["pages"].items()))
    if kind == "sports":
        _sheet(wb, "ATTRIBUTED CONVERSIONS BY URL", ["Page Url", "Attributed Conversions"],
               [(u, 0) for u in f["pages"]])
    _sheet(wb, "RECENCY (elapsed time, in days,", ["Attributed Unique Visitors", "Recency"],
           [(v, k) for k, v in f["recency"].items()])
    _sheet(wb, "REFERRING DOMAIN", ["Attributed Unique Visitors", "Referral Domain"],
           [(v, k) for k, v in f["referral"].items()])
    if kind == "standard":
        _sheet(wb, "ATTRIBUTED RATE BY CHANNEL", ["Channel Name"] + _DETAIL,
               [("Example Network", d, a, _rate(a, d), 0, 0.0)])
    _sheet(wb, "ATTRIBUTED VISITS BY DEVICE TYP", ["Attributed Impressions", "Device Type"],
           [(v, k) for k, v in f["devices"].items()])
    # Day of week: for sports only the game days carry delivery.
    dow_days = 2 if kind == "sports" else 7
    dow_start = date(2026, 11, 2)
    _sheet(wb, "ATTRIBUTED RATE BY DAY OF WEEK ", ["Date"] + _DETAIL,
           [(dow_start + timedelta(days=i), dd, aa, _rate(aa, dd), 0, 0.0)
            for i, (dd, aa) in enumerate(zip(_split(d, dow_days), _split(a, dow_days)))])
    weeks = [f["start"] + timedelta(weeks=i) for i in range(f["weeks"])]
    _sheet(wb, "ATTRIBUTED RATE AND CONVERSIO_2", ["Date"] + _DETAIL,
           [(w, dd, aa, _rate(aa, dd), 0, 0.0)
            for w, dd, aa in zip(weeks, _split(d, len(weeks)), _split(a, len(weeks)))])
    months = sorted({date(w.year, w.month, 1) for w in weeks})
    _sheet(wb, "ATTRIBUTED RATE AND CONVERSIO_3", ["Date", "Attributed Rate",
                                                   "Conversion Impressions Rate"],
           [(m, _rate(a, d), 0.0) for m in months])
    rf_d, rf_a = _split(d, len(f["rfpids"])), _split(a, len(f["rfpids"]))
    _sheet(wb, "ATTRIBUTED RATE AND CONVERSIO_4", ["Rfpid"] + _DETAIL,
           [(r, dd, aa, _rate(aa, dd), 0, 0.0) for r, dd, aa in zip(f["rfpids"], rf_d, rf_a)])
    if kind == "standard":
        _sheet(wb, "WEBSITE ATTRIBUTION ADVERTISER ",
               ["Client", "Advertiser Pixel", "Delivered Impressions", "Attributed Impressions",
                "Attributed Unique Visitors"],
               [("Example Client - Direct", "TEGNA - WUSA - Example Client - Direct Pixel",
                 a, a, f["uniques"])])
    wb.save(path)
    return path
