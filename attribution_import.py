"""attribution_import.py -- parse a Premion Website Attribution export and
its (optional) companion Delivery export into one shape apiece.

Pure -- no Streamlit, no DB. `parse_attribution_export(path)` and
`parse_delivery_export(path)` each return their own dataclass or raise
AttributionParseError with a message aimed at a seller, the same convention
as wideorbit.py / avails_pdf_import.py / notes_file_import.py.

**Tabs are identified by header row, never by name** (ATTRIBUTION_ROADMAP_
ADDENDUM.md's own Phase 1 rule) -- the export tool emits "Untitled Widget",
"ATTRIBUTED RATE AND CONVERSIO_2" and similar auto-numbered names that carry
no meaning at all. Header text is normalized (BOM and non-breaking-space
stripped everywhere, not just at the ends -- "ATTRIBUTED RATE\xa0AND
CONVERSION " carries the NBSP mid-string -- whitespace collapsed, lowercased)
before any comparison.

**Four separate real gotchas found building this against real MW/Cardinal
exports, all load-bearing:**

1. Several dimension pairs share BYTE-IDENTICAL headers but different
   content -- a "ranked" tab (Attributed Rate, Conversion Rate, dimension
   name -- 3 columns, dimension LAST) and a "detail" tab (dimension name,
   Delivered/Attributed Impressions, Attributed Rate, Conversion Impressions/
   Rate -- 6 columns, dimension FIRST). Always take the fuller one; the
   ranked one is the same data, resorted, with less information. Since the
   dimension name column is always first in the detail tab and always last
   (or absent) in the ranked one, matching on the FULL expected header tuple
   (dimension name first) picks the right one without needing a separate
   "which is fuller" heuristic.
2. Three date-keyed tabs share the IDENTICAL header (Date | Delivered
   Impressions | Attributed Impressions | Attributed Rate | Conversion
   Impressions | Conversion Impressions Rate) and can only be told apart by
   the GAP between consecutive dates: 1-day gaps are a trailing daily
   window (the dashboard mislabels this tab "Day of Week" -- confirmed
   against two real exports it is NOT a weekday aggregate, just the most
   recent several calendar days), 7-day gaps are the weekly trend (the one
   that actually spans the reported period), ~28-31-day gaps are the
   monthly rollup. Classify each candidate independently by its own date
   deltas -- never by tab name or position.
3. An advertiser pixel can carry more than one RFPID. **Two real shapes,
   and only a rep can tell them apart -- this module no longer refuses
   either one.** A GLS/Twin Pine-style lifetime rollup (confirmed against
   two real samples, one with 180 RFPIDs, 180+ zips, years of history) and
   a WAEPA-style split IO (2 RFPIDs, 549,296 + 60,464 impressions, the
   SAME audiences/creatives, overlapping dates -- one campaign issued as
   two orders) look identical to this parser: it has no per-RFPID dates or
   per-RFPID dimension rows to tell them apart by. `result.rfpid_breakdown`
   carries every RFPID's own delivered/attributed/conversion figures (one
   entry even for an ordinary single-RFPID file, so a caller never has to
   special-case count==1); the caller decides, informed by the count and
   figures, whether to treat it as one campaign -- see `app.py`'s confirm
   gate. `result.rfpid` becomes a " + "-joined string when there's more
   than one, which nothing downstream parses structurally (confirmed --
   it's a display value only, unlike the avails-PDF importer's own RFPID,
   which IS a dedup key elsewhere).
4. The delivery export packs a count and a percentage into ONE string cell
   ("112320 - 12.24%") on several of its tabs, rather than two numeric
   columns. Parsed by `_split_count_pct`; a plain numeric cell (no "-")
   still works the same way.

**Conversions are a per-export OPTIONAL layer, not a fifth pair of MW/
Cardinal-shaped fields.** MW and Cardinal both carry the per-dimension
"Conversion Impressions"/"Conversion Impressions Rate" columns already
(`AttributionRow.conversion_impressions`/`.conversion_rate` predate WAEPA)
-- they are simply always zero there, columns present, no data. The real
signal that an export actually HAS conversions is the top-line "Attributed
Conversions"/"Sales Amount" widget: present with a real value on WAEPA
(37, $0), absent entirely on both MW and Cardinal (not present-and-zero --
genuinely missing as a tab). `has_conversions` is true only when that
widget tab exists AND its value is > 0 -- checking existence alone would
call it "present" on a hypothetical export that has the widget tab but
happened to convert zero times, which is a real, valid no-conversions
report and must read as one. `conversions_by_url` (`ATTRIBUTED CONVERSIONS
BY URL`) is the same idea, one level down -- report_assembly.py folds it
into the URL slide's own intent classes, tagged by the same
`classify_url_intent` the visits already use.

Nothing here resolves a client to a proposal, matches an advertiser name, or
draws a slide -- this module only ever reports what the export says.
"""
import re
from dataclasses import dataclass, field
from datetime import date

import openpyxl


class AttributionParseError(Exception):
    """Raised with a message intended for a seller, not a developer."""


# ---------------------------------------------------------------------------
# Header normalization + tab lookup
# ---------------------------------------------------------------------------

_INVISIBLE_RE = re.compile("[﻿\xa0]")
_WHITESPACE_RE = re.compile(r"\s+")


def _normalize_cell(value):
    text = _INVISIBLE_RE.sub(" ", str(value if value is not None else ""))
    return _WHITESPACE_RE.sub(" ", text).strip().lower()


def _sheet_header(ws):
    row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not row:
        return ()
    return tuple(_normalize_cell(c) for c in row)


def _sheet_rows(ws):
    """This sheet's data rows as plain dicts keyed by its own header."""
    header = _sheet_header(ws)
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and all(c is None for c in row):
            continue
        out.append(dict(zip(header, row)))
    return out


def _index_by_header(wb):
    """{header_tuple: [worksheet, ...]} -- more than one sheet can share a
    header, which is the whole reason this exists rather than a plain dict.
    """
    index = {}
    for ws in wb.worksheets:
        index.setdefault(_sheet_header(ws), []).append(ws)
    return index


def _find_one(index, header):
    sheets = index.get(header)
    if not sheets:
        return None
    return sheets[0]


def _clean_int(value):
    if value is None:
        return 0
    try:
        return int(round(float(str(value).replace(",", "").strip())))
    except (TypeError, ValueError):
        return 0


def _clean_float(value):
    if value is None:
        return 0.0
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


_COUNT_PCT_RE = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*(?:-\s*([\d.]+)\s*%)?\s*$")


def _split_count_pct(value):
    """"112320 - 12.24%" -> (112320, 12.24); a bare number -> (n, None).
    The delivery export's own convention -- see gotcha 4 above.
    """
    if value is None:
        return 0, None
    match = _COUNT_PCT_RE.match(str(value))
    if not match:
        return _clean_int(value), None
    count = _clean_int(match.group(1))
    pct = float(match.group(2)) if match.group(2) else None
    return count, pct


def _date_gaps(dates):
    ordered = sorted({d for d in dates if d is not None})
    if len(ordered) < 2:
        return []
    return [(ordered[i + 1] - ordered[i]).days for i in range(len(ordered) - 1)]


def _classify_date_series(dates):
    """"daily" / "weekly" / "monthly" / None -- see gotcha 2 above."""
    gaps = _date_gaps(dates)
    if not gaps:
        return None
    max_gap = max(gaps)
    if max_gap <= 1:
        return "daily"
    if 6 <= max_gap <= 8:
        return "weekly"
    if 25 <= max_gap <= 35:
        return "monthly"
    return None


# ---------------------------------------------------------------------------
# Attribution export
# ---------------------------------------------------------------------------

_HEADLINE_A = ("delivered impressions", "attributed impressions", "attributed rate")
_HEADLINE_B = ("attributed unique visitors", "attributed unique visitor rate")
_HEADLINE_CONVERSIONS = ("attributed conversions", "sales amount")
_CONVERSIONS_BY_URL = ("page url", "attributed conversions")
_DIMENSION_DETAIL_SUFFIX = ("delivered impressions", "attributed impressions",
                            "attributed rate", "conversion impressions",
                            "conversion impressions rate")
_ZIP_HEAT = ("attributed impressions", "zipcode")
_BY_URL = ("page url", "attributed unique visitors")
_BY_RECENCY = ("attributed unique visitors", "recency")
_BY_REFERRAL = ("attributed unique visitors", "referral domain")
_BY_DEVICE = ("attributed impressions", "device type")
_DATE_HEADER = ("date",) + _DIMENSION_DETAIL_SUFFIX
_BY_RFPID = ("rfpid",) + _DIMENSION_DETAIL_SUFFIX
_ADVERTISER = ("client", "advertiser pixel", "delivered impressions",
              "attributed impressions", "attributed unique visitors")

# Longest first: "wusa9" must not be shadowed by a bare "wusa" substring
# check, and this dict is also matched by simple substring search so a
# station name embedded anywhere in the pixel string is found.
_STATION_MARKETS = {"wusa9": "DC", "wusa": "DC", "wpmt": "Harrisburg", "fox43": "Harrisburg"}


@dataclass
class AttributionRow:
    label: str
    delivered_impressions: int
    attributed_impressions: int
    attributed_rate: float
    conversion_impressions: int = 0
    conversion_rate: float = 0.0


@dataclass
class DateSeriesPoint:
    day: date
    delivered_impressions: int
    attributed_impressions: int
    attributed_rate: float


@dataclass
class AttributionExport:
    source_name: str = ""
    rfpid: str = ""
    client_name: str = ""
    advertiser_pixel: str = ""
    market_hint: str = None                # "DC" / "Harrisburg" / None
    delivered_impressions: int = 0         # this file's OWN figure -- the
                                            # rate-math denominator, never
                                            # the deck's headline (see
                                            # ATTRIBUTION_REPORT_PLAN.md
                                            # Correction 3)
    attributed_impressions: int = 0
    attributed_rate: float = 0.0
    attributed_unique_visitors: int = 0
    attributed_unique_visitor_rate: float = 0.0
    by_audience: list = field(default_factory=list)     # [AttributionRow]
    by_creative: list = field(default_factory=list)
    by_market: list = field(default_factory=list)
    by_zip: list = field(default_factory=list)
    by_channel: list = field(default_factory=list)
    zip_heat: dict = field(default_factory=dict)         # {zip: attributed_impressions}
    by_url: dict = field(default_factory=dict)            # {url: unique_visitors}
    by_recency: dict = field(default_factory=dict)        # {bucket: unique_visitors}
    by_referral_domain: dict = field(default_factory=dict)
    by_device: dict = field(default_factory=dict)
    daily_trend: list = field(default_factory=list)       # [DateSeriesPoint], trailing window only
    weekly_trend: list = field(default_factory=list)      # [DateSeriesPoint], spans the report period
    monthly_trend: list = field(default_factory=list)     # [DateSeriesPoint]
    flight_start: date = None                              # derived from the trend tabs' own dates
    flight_end: date = None
    rfpid_breakdown: list = field(default_factory=list)    # [{"rfpid","delivered_impressions",
                                                            #   "attributed_impressions","attributed_rate",
                                                            #   "conversion_impressions","conversion_rate"}]
                                                            # one entry even for an ordinary single-RFPID
                                                            # file -- a caller never special-cases count==1.
                                                            # See gotcha 3 above for why this module no
                                                            # longer rejects more than one.
    attributed_conversions: int = 0        # 0 whether the widget is absent (MW/Cardinal) or
                                            # genuinely zero -- has_conversions is what tells them apart
    sales_amount: float = 0.0
    has_conversions: bool = False          # the top-line widget tab exists AND its value is > 0
    conversions_by_url: dict = field(default_factory=dict)  # {url: attributed_conversions}
    warnings: list = field(default_factory=list)


def _market_hint_from_pixel(pixel):
    lowered = (pixel or "").lower()
    for station, market in _STATION_MARKETS.items():
        if station in lowered:
            return market
    return None


def _dimension_rows(index, dimension_label):
    header = (dimension_label,) + _DIMENSION_DETAIL_SUFFIX
    ws = _find_one(index, header)
    if ws is None:
        return []
    rows = []
    for row in _sheet_rows(ws):
        rows.append(AttributionRow(
            label=str(row.get(dimension_label) or "").strip(),
            delivered_impressions=_clean_int(row.get("delivered impressions")),
            attributed_impressions=_clean_int(row.get("attributed impressions")),
            attributed_rate=_clean_float(row.get("attributed rate")),
            conversion_impressions=_clean_int(row.get("conversion impressions")),
            conversion_rate=_clean_float(row.get("conversion impressions rate")),
        ))
    return rows


def _date_series(index):
    """{"daily": [...], "weekly": [...], "monthly": [...]} from every
    candidate date-headed tab, classified independently (gotcha 2)."""
    out = {"daily": [], "weekly": [], "monthly": []}
    for ws in index.get(_DATE_HEADER, []):
        rows = _sheet_rows(ws)
        dates = [r.get("date").date() if hasattr(r.get("date"), "date") else r.get("date")
                for r in rows]
        kind = _classify_date_series(dates)
        if kind is None:
            continue
        points = []
        for row, day in zip(rows, dates):
            if day is None:
                continue
            points.append(DateSeriesPoint(
                day=day,
                delivered_impressions=_clean_int(row.get("delivered impressions")),
                attributed_impressions=_clean_int(row.get("attributed impressions")),
                attributed_rate=_clean_float(row.get("attributed rate")),
            ))
        points.sort(key=lambda p: p.day)
        if points:
            out[kind] = points
    return out


def parse_attribution_export(path, source_name=None):
    """AttributionExport, or raises AttributionParseError."""
    source_name = source_name or str(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:                                              # noqa: BLE001
        raise AttributionParseError(
            f"Couldn't open \"{source_name}\" as an Excel file: {exc}") from exc

    index = _index_by_header(wb)
    result = AttributionExport(source_name=source_name)

    ws = _find_one(index, _HEADLINE_A)
    if ws is not None:
        rows = _sheet_rows(ws)
        if rows:
            result.delivered_impressions = _clean_int(rows[0].get("delivered impressions"))
            result.attributed_impressions = _clean_int(rows[0].get("attributed impressions"))
            result.attributed_rate = _clean_float(rows[0].get("attributed rate"))
    ws = _find_one(index, _HEADLINE_B)
    if ws is not None:
        rows = _sheet_rows(ws)
        if rows:
            result.attributed_unique_visitors = _clean_int(rows[0].get("attributed unique visitors"))
            result.attributed_unique_visitor_rate = _clean_float(
                rows[0].get("attributed unique visitor rate"))

    ws = _find_one(index, _HEADLINE_CONVERSIONS)
    if ws is not None:
        rows = _sheet_rows(ws)
        if rows:
            result.attributed_conversions = _clean_int(rows[0].get("attributed conversions"))
            result.sales_amount = _clean_float(rows[0].get("sales amount"))
    result.has_conversions = ws is not None and result.attributed_conversions > 0

    result.by_audience = _dimension_rows(index, "audience name")
    result.by_creative = _dimension_rows(index, "creative name")
    result.by_market = _dimension_rows(index, "market")
    result.by_channel = _dimension_rows(index, "channel name")

    ws = _find_one(index, ("zipcode",) + _DIMENSION_DETAIL_SUFFIX)
    if ws is not None:
        for row in _sheet_rows(ws):
            zip_code = str(row.get("zipcode") or "").strip().zfill(5)
            if not zip_code:
                continue
            result.by_zip.append(AttributionRow(
                label=zip_code,
                delivered_impressions=_clean_int(row.get("delivered impressions")),
                attributed_impressions=_clean_int(row.get("attributed impressions")),
                attributed_rate=_clean_float(row.get("attributed rate")),
                conversion_impressions=_clean_int(row.get("conversion impressions")),
                conversion_rate=_clean_float(row.get("conversion impressions rate")),
            ))

    ws = _find_one(index, _ZIP_HEAT)
    if ws is not None:
        for row in _sheet_rows(ws):
            zip_code = str(row.get("zipcode") or "").strip().zfill(5)
            if zip_code:
                result.zip_heat[zip_code] = _clean_int(row.get("attributed impressions"))

    ws = _find_one(index, _BY_URL)
    if ws is not None:
        for row in _sheet_rows(ws):
            url = str(row.get("page url") or "").strip()
            if url:
                result.by_url[url] = _clean_int(row.get("attributed unique visitors"))

    ws = _find_one(index, _CONVERSIONS_BY_URL)
    if ws is not None:
        for row in _sheet_rows(ws):
            url = str(row.get("page url") or "").strip()
            if url:
                result.conversions_by_url[url] = _clean_int(row.get("attributed conversions"))

    ws = _find_one(index, _BY_RECENCY)
    if ws is not None:
        for row in _sheet_rows(ws):
            bucket = str(row.get("recency") or "").strip()
            if bucket:
                result.by_recency[bucket] = _clean_int(row.get("attributed unique visitors"))

    ws = _find_one(index, _BY_REFERRAL)
    if ws is not None:
        for row in _sheet_rows(ws):
            domain = str(row.get("referral domain") or "").strip()
            if domain:
                result.by_referral_domain[domain] = _clean_int(row.get("attributed unique visitors"))

    ws = _find_one(index, _BY_DEVICE)
    if ws is not None:
        for row in _sheet_rows(ws):
            device = str(row.get("device type") or "").strip()
            if device:
                result.by_device[device] = _clean_int(row.get("attributed impressions"))

    trend = _date_series(index)
    result.daily_trend = trend["daily"]
    result.weekly_trend = trend["weekly"]
    result.monthly_trend = trend["monthly"]
    # The daily/trailing series reflects whenever the export was PULLED, not
    # the campaign's own flight -- confirmed against two real exports, both
    # pulled well after (MW) or during a gap past (Cardinal) the window
    # their own weekly trend covers. Only the weekly/monthly series actually
    # spans the reported period (gotcha 2 above), so flight span is derived
    # from those two alone, never the daily one.
    span_days = [p.day for series in (trend["weekly"], trend["monthly"]) for p in series]
    if span_days:
        result.flight_start = min(span_days)
        result.flight_end = max(span_days)
    else:
        result.warnings.append(
            "No weekly or monthly trend tab was found to derive a flight span from -- "
            "report period will need to be entered by hand.")

    ws = _find_one(index, _BY_RFPID)
    rfpid_rows = _sheet_rows(ws) if ws is not None else []
    for row in rfpid_rows:
        result.rfpid_breakdown.append({
            "rfpid": str(row.get("rfpid") or "").strip(),
            "delivered_impressions": _clean_int(row.get("delivered impressions")),
            "attributed_impressions": _clean_int(row.get("attributed impressions")),
            "attributed_rate": _clean_float(row.get("attributed rate")),
            "conversion_impressions": _clean_int(row.get("conversion impressions")),
            "conversion_rate": _clean_float(row.get("conversion impressions rate")),
        })
    if rfpid_rows:
        result.rfpid = " + ".join(r["rfpid"] for r in result.rfpid_breakdown if r["rfpid"])
    if len(rfpid_rows) > 1:
        result.warnings.append(
            f"\"{source_name}\" covers {len(rfpid_rows)} RFPIDs, not one -- confirm this is a "
            f"single campaign (a split IO) before generating the report, not a lifetime/"
            f"advertiser rollup.")

    ws = _find_one(index, _ADVERTISER)
    if ws is not None:
        rows = [r for r in _sheet_rows(ws) if r.get("advertiser pixel")]
        if rows:
            result.client_name = str(rows[0].get("client") or "").strip()
            result.advertiser_pixel = str(rows[0].get("advertiser pixel") or "").strip()
            result.market_hint = _market_hint_from_pixel(result.advertiser_pixel)
        else:
            result.warnings.append(
                "No advertiser/pixel row was found -- client name will need to be entered by hand.")

    return result


# ---------------------------------------------------------------------------
# Delivery export (optional companion file)
# ---------------------------------------------------------------------------

_KPI_DELIVERY = ("delivered impressions", "vcr")
_KPI_PERFORMANCE = ("frequency", "uniques")
_DAILY_DELIVERY = ("date", "delivered impressions")
_CREATIVE_DELIVERY = ("creative name", "delivered impressions", "creative length",
                      "hours watched", "vcr")
_TOP_CHANNELS = ("channel name", "delivered impressions")
_CHANNEL_VCR = ("channel name", "delivered impressions", "vcr", "q1 - 25%", "q2 - 50%",
               "q3 - 75%", "complete")
_TOP_ZIPCODES = ("zipcode", "delivered impressions")
_DELIVERY_MAP = ("delivered impressions", "zipcode")
# Two DIFFERENT device tabs, and picking the wrong one answers a different
# question. "DEVICE DISTRIBUTION" (grouping) splits OTT vs Desktop vs
# Mobile-Web -- essentially always ~100% OTT, since this is a CTV product.
# "OTT DISTRIBUTION" (device category name) splits WITHIN OTT: Connected TV
# vs Mobile In-App vs Tablet In-App. The report's "% that ran on CTV
# screens" tile is the SECOND one -- on real files that's 94.2% (MW) vs
# 99.9% (Cardinal), where the first tab would have said 99.99% and 100.00%
# and told a client nothing.
_OTT_DISTRIBUTION = ("delivered impressions", "device category name")
_DAYPART = ("delivered impressions", "daypart distribution")
_TOP_GEO = ("geo", "delivered impressions")

# The daypart labels carry a sort prefix ("A.MID - 2AM", "B.2AM - 6AM") that
# exists only to order them -- it is not something a client should read.
# Stripped for display, kept for ordering (see _split_daypart_label).
_DAYPART_PREFIX_RE = re.compile(r"^\s*([A-Z])\s*\.\s*(.+)$")
# The geo values are a rep-authored zip-option label, not a DMA. On MW they
# read "Zip Option - Raleigh Market"; on Cardinal, "Zip Option - Primary Zip
# List". The shared prefix is noise in both cases; what's left is the rep's
# own words, which the Generate step lets them edit before the deck builds.
_ZIP_OPTION_PREFIX_RE = re.compile(r"^\s*zip\s+option\s*-\s*", re.I)

# What counts as a Connected TV screen in the OTT DISTRIBUTION tab. Matched
# case-insensitively against the whole category name, never a substring --
# "Mobile In-App" and "Tablet In-App" are also OTT but are not TV screens,
# which is the entire distinction this tile exists to draw.
_CTV_CATEGORY = "connected tv"


def _split_daypart_label(raw):
    """("A", "MID - 2AM") for "A.MID - 2AM"; (None, text) when there is no
    sort prefix. The prefix orders the buckets and is never displayed."""
    text = str(raw or "").strip()
    match = _DAYPART_PREFIX_RE.match(text)
    if match:
        return match.group(1), match.group(2).strip()
    return None, text


def strip_zip_option_prefix(label):
    """"Zip Option - Raleigh Market" -> "Raleigh Market". Leaves anything
    without the prefix alone, so a differently-named option survives
    untouched rather than being partially eaten."""
    return _ZIP_OPTION_PREFIX_RE.sub("", str(label or "")).strip()
_UNTITLED_FLIGHT_DETAIL = ("campaign name", "flight start date", "flight end date", "geo",
                          "booked impressions", "delivered impressions", "vcr",
                          "campaign id", "group id")


@dataclass
class DeliveryExport:
    source_name: str = ""
    delivered_impressions: int = 0     # THE headline figure when this file exists
                                        # (ATTRIBUTION_REPORT_PLAN.md Correction 3)
    vcr: float = 0.0
    frequency: float = 0.0
    uniques: int = 0
    top_publishers: list = field(default_factory=list)   # [(channel, delivered, pct_or_None)]
    channel_vcr: dict = field(default_factory=dict)       # {channel: vcr} -- a SEPARATE tab from
                                                           # top_publishers (real per-channel VCR,
                                                           # not the share-of-total pct the "Top 10
                                                           # Channels" tabs carry); missing/absent
                                                           # tab means an empty dict, never a guess
    by_creative: list = field(default_factory=list)       # [(name, delivered, length_sec, hours_watched, vcr)]
    top_zipcodes: list = field(default_factory=list)      # [(zip, delivered, pct_or_None)]
    delivery_map: dict = field(default_factory=dict)      # {zip: delivered_impressions}
    daily_delivery: list = field(default_factory=list)    # [(date, delivered_impressions)] -- a REAL
                                                           # full daily series, unlike the attribution
                                                           # file's trailing-window-only "Day of Week" tab
    by_device_category: list = field(default_factory=list)  # [(category, delivered)] within OTT
    ctv_impressions: int = 0                # Connected TV alone, 0 when the tab is absent
    by_daypart: list = field(default_factory=list)   # [(label, delivered)] in the export's own
                                                      # order, sort prefix stripped from the label
    geo_vcr: dict = field(default_factory=dict)      # {geo label: vcr} -- IMPRESSION-WEIGHTED
                                                      # across that geo's flight rows, not a
                                                      # plain mean: a geo's rows differ in size
                                                      # by 10x+ in the real files, so averaging
                                                      # the rates would let a tiny month move
                                                      # the number as much as a big one. Keyed
                                                      # by the SAME stripped label by_geo uses.
    by_geo: list = field(default_factory=list)       # [(label, delivered)] -- the rep-authored
                                                      # zip-option label, "Zip Option - " stripped.
                                                      # NOT a market/DMA: Cardinal's read "Primary
                                                      # Zip List", MW's read "Raleigh Market", and
                                                      # only the rep knows which is client-facing,
                                                      # which is why the label is editable at
                                                      # Generate rather than trusted here.
    booked_impressions: int = None
    warnings: list = field(default_factory=list)

    @property
    def ctv_share(self):
        """Connected TV as a fraction of total delivery, or None when
        either side is missing -- never 0.0, which would render as a
        confident "0.0% ran on CTV" on a slide."""
        if not self.delivered_impressions or not self.ctv_impressions:
            return None
        return self.ctv_impressions / self.delivered_impressions


def _pick_channel_tab(candidates):
    """The delivery file repeats "top channels" with an identical header --
    one carries a literal TOTAL row and no percentages, the other carries
    real percentages and no TOTAL row. Prefer the one with percentages;
    strip a TOTAL row defensively from whichever is used either way.
    """
    best = None
    for ws in candidates:
        rows = _sheet_rows(ws)
        has_pct = any(_split_count_pct(r.get("delivered impressions"))[1] is not None for r in rows)
        if has_pct:
            return rows
        if best is None:
            best = rows
    return best or []


def parse_delivery_export(path, source_name=None):
    """DeliveryExport, or raises AttributionParseError."""
    source_name = source_name or str(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:                                              # noqa: BLE001
        raise AttributionParseError(
            f"Couldn't open \"{source_name}\" as an Excel file: {exc}") from exc

    index = _index_by_header(wb)
    result = DeliveryExport(source_name=source_name)

    ws = _find_one(index, _KPI_DELIVERY)
    if ws is not None:
        rows = _sheet_rows(ws)
        if rows:
            result.delivered_impressions = _clean_int(rows[0].get("delivered impressions"))
            result.vcr = _clean_float(rows[0].get("vcr"))
    else:
        result.warnings.append("No KPI delivery tab found -- delivered impressions/VCR missing.")

    ws = _find_one(index, _KPI_PERFORMANCE)
    if ws is not None:
        rows = _sheet_rows(ws)
        if rows:
            result.frequency = _clean_float(rows[0].get("frequency"))
            result.uniques = _clean_int(rows[0].get("uniques"))

    ws = _find_one(index, _CREATIVE_DELIVERY)
    if ws is not None:
        for row in _sheet_rows(ws):
            name = str(row.get("creative name") or "").strip()
            if name:
                result.by_creative.append((
                    name,
                    _clean_int(row.get("delivered impressions")),
                    _clean_int(row.get("creative length")),
                    _clean_int(row.get("hours watched")),
                    _clean_float(row.get("vcr")),
                ))

    for row in _pick_channel_tab(index.get(_TOP_CHANNELS, [])):
        name = str(row.get("channel name") or "").strip()
        if not name or name.upper() == "TOTAL":
            continue
        count, pct = _split_count_pct(row.get("delivered impressions"))
        result.top_publishers.append((name, count, pct))

    ws = _find_one(index, _CHANNEL_VCR)
    if ws is not None:
        for row in _sheet_rows(ws):
            name = str(row.get("channel name") or "").strip()
            if name and name.upper() != "TOTAL":
                result.channel_vcr[name] = _clean_float(row.get("vcr"))

    ws = _find_one(index, _OTT_DISTRIBUTION)
    if ws is not None:
        for row in _sheet_rows(ws):
            category = str(row.get("device category name") or "").strip()
            if not category or category.upper() == "TOTAL":
                continue
            count, _pct = _split_count_pct(row.get("delivered impressions"))
            result.by_device_category.append((category, count))
            if category.lower() == _CTV_CATEGORY:
                result.ctv_impressions = count
        if not result.ctv_impressions and result.by_device_category:
            result.warnings.append(
                "The OTT distribution tab has no 'Connected TV' row -- CTV share will be "
                "left off the report rather than guessed from the other categories.")

    ws = _find_one(index, _DAYPART)
    if ws is not None:
        buckets = []
        for row in _sheet_rows(ws):
            prefix, label = _split_daypart_label(row.get("daypart distribution"))
            if not label:
                continue
            count, _pct = _split_count_pct(row.get("delivered impressions"))
            buckets.append((prefix, label, count))
        # Sorted by the export's own prefix when every bucket has one; left
        # in sheet order otherwise, rather than sorting alphabetically on a
        # label like "MID - 2AM" and putting midnight after 4PM.
        if buckets and all(p for p, _l, _c in buckets):
            buckets.sort(key=lambda b: b[0])
        result.by_daypart = [(label, count) for _p, label, count in buckets]

    ws = _find_one(index, _TOP_GEO)
    if ws is not None:
        for row in _sheet_rows(ws):
            label = strip_zip_option_prefix(row.get("geo"))
            if not label or label.upper() == "TOTAL":
                continue
            count, _pct = _split_count_pct(row.get("delivered impressions"))
            result.by_geo.append((label, count))

    ws = _find_one(index, _TOP_ZIPCODES)
    if ws is not None:
        for row in _sheet_rows(ws):
            zip_code = str(row.get("zipcode") or "").strip().zfill(5)
            if not zip_code:
                continue
            count, pct = _split_count_pct(row.get("delivered impressions"))
            result.top_zipcodes.append((zip_code, count, pct))

    ws = _find_one(index, _DELIVERY_MAP)
    if ws is not None:
        for row in _sheet_rows(ws):
            zip_code = str(row.get("zipcode") or "").strip().zfill(5)
            if zip_code:
                result.delivery_map[zip_code] = _clean_int(row.get("delivered impressions"))

    ws = _find_one(index, _DAILY_DELIVERY)
    if ws is not None:
        points = []
        for row in _sheet_rows(ws):
            day = row.get("date")
            day = day.date() if hasattr(day, "date") else day
            if day is None:
                continue
            points.append((day, _clean_int(row.get("delivered impressions"))))
        points.sort(key=lambda p: p[0])
        result.daily_delivery = points

    ws = _find_one(index, _UNTITLED_FLIGHT_DETAIL)
    if ws is not None:
        rows = [r for r in _sheet_rows(ws) if str(r.get("campaign name") or "").strip()]
        if rows:
            result.booked_impressions = sum(_clean_int(r.get("booked impressions")) for r in rows)
            # Per-geo VCR lives ONLY here -- the "TOP 10 GEO BY IMPRESSIONS"
            # tab that by_geo comes from carries impressions and nothing
            # else. Verified against both real files that this tab's
            # per-geo impressions reconcile exactly with that one
            # (MW 1,862,614 / 1,241,940; Cardinal 458,737 / 458,714), so
            # the two are describing the same split and joining them on the
            # stripped label is sound rather than a coincidence of naming.
            weighted, totals = {}, {}
            for row in rows:
                label = strip_zip_option_prefix(row.get("geo"))
                if not label:
                    continue
                delivered = _clean_int(row.get("delivered impressions"))
                if not delivered:
                    continue
                weighted[label] = weighted.get(label, 0.0) + delivered * _clean_float(row.get("vcr"))
                totals[label] = totals.get(label, 0) + delivered
            result.geo_vcr = {label: weighted[label] / totals[label]
                             for label in totals if totals[label]}

    return result
