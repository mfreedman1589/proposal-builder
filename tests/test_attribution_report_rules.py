"""The St. James attribution report review (2026-10-09) -- the report rules it
settled, each guarded here.

    python tests/test_attribution_report_rules.py

    1  The report period comes from the delivery file's flight dates, each
       flight shown and any dark gap named -- never the week-bucket labels.
    2  Recency bucket labels come from the source file; a drafted day range
       the export doesn't report is caught and removed.
    4  A drafted response claim about a dimension with no attribution data
       (daypart) is rejected after the draft.
    5  The ZIP slide's table and subtitle agree: only ZIPs at 1.2x the average
       or better on real volume, sorted by multiple; the heat map's title
       names the metric its legend shades by.
    8  Trend: the attributed rate by real-dated delivery week (dark weeks
       marked) and by flight; it qualifies on 3+ consecutive rises or a
       material flight-to-flight swing, then leads the one-sheet's bottom
       line and is always a highlight thread.
   10  One percentage per fact.
   11  Page tables show single pages with their own visitor counts, labelled
       "Visitors" -- never summed across pages, never a percentage.
   12  CPV: a tile on the highlights slide and the one-sheet with the toggle
       on; nowhere with it off. No Unique Visitor Rate tile.
   13  Recency and referral are shares only -- bar charts, a coverage
       footnote, and no count from either tab anywhere in the deck text.
   14  Page groups come from the goal's own entities plus confirmation pages;
       "Other" is never the top row.
   15  The traffic-mix table is a goal path, counts only.
   16  Confirmation pages are a standard highlight, worded "reached a
       confirmation page" -- never "converted" or "signed up".

Most checks use SYNTHETIC exports (never SKIP). Deck checks need the
gitignored report template; the real St. James pair ("Premion Website
Attribution and Reach Extension STJ.xlsx" / "Premion OTT STJ.xlsx") adds a
real-data pass and SKIPs without it.
"""
import copy
import sys
from datetime import date, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import db                                  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import market_lookup                       # noqa: E402
from pptx import Presentation              # noqa: E402

import app                                 # noqa: E402
import attribution_import as ai            # noqa: E402
import package_check                       # noqa: E402
import report_assembly as ra               # noqa: E402
import slide_map                           # noqa: E402
from attribution_import import AttributionRow, DateSeriesPoint  # noqa: E402

TEMPLATE = REPO / "REPORT_MASTER_v0_15.pptx"
STJ_ATTRIBUTION = REPO / "Premion Website Attribution and Reach Extension STJ.xlsx"
STJ_DELIVERY = REPO / "Premion OTT STJ.xlsx"
SCRATCH = REPO / "tests" / "_manual_output"
GOALS = ["Promote the Springfield and Bethesda performance clubs; the goal is new membership"]

market_lookup.install()


class Report:
    def __init__(self):
        self.passed, self.failed, self.skipped = 0, [], []

    def check(self, label, ok, detail=""):
        if ok:
            print(f"  PASS  {label}")
            self.passed += 1
        else:
            print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
            self.failed.append(label)

    def skip(self, reason):
        print(f"  SKIP  {reason}")
        self.skipped.append(reason)


# ---------------------------------------------------------------------------
# A synthetic two-flight campaign shaped like St. James: Aug 1-14 and Aug
# 23-Sep 5 2026, dark in between; 10,000 impressions a day; weekly buckets
# labelled the Monday after their first day (the export's own habit), the
# attributed rate rising every on-air week.
# ---------------------------------------------------------------------------

FLIGHT_1 = (date(2026, 8, 1), date(2026, 8, 14))
FLIGHT_2 = (date(2026, 8, 23), date(2026, 9, 5))
DAILY = 10_000


def _days(start, end):
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def make_delivery(two_flights=True):
    days = _days(*FLIGHT_1) + (_days(*FLIGHT_2) if two_flights else [])
    flights = [{"start": FLIGHT_1[0], "end": FLIGHT_1[1], "impressions": 14 * DAILY,
                "uniques": 30_000, "booked": 140_000}]
    if two_flights:
        flights.append({"start": FLIGHT_2[0], "end": FLIGHT_2[1], "impressions": 14 * DAILY,
                        "uniques": 28_000, "booked": 140_000})
    return ai.DeliveryExport(
        delivered_impressions=DAILY * len(days), vcr=0.97, frequency=4.6, uniques=58_000,
        top_publishers=[("Hulu", 100_000, 0.36), ("Pluto TV", 80_000, 0.29)],
        channel_vcr={"Hulu": 0.98, "Pluto TV": 0.97},
        daily_delivery=[(d, DAILY) for d in days],
        ctv_impressions=DAILY * len(days),          # 100% CTV -> delivered vs. booked tile
        by_daypart=[("6AM - 9AM", 100_000), ("7PM - MID", 180_000)],
        by_geo=[("Springfield Zips", 200_000), ("Bethesda Zips", 80_000)],
        geo_vcr={"Springfield Zips": 0.97, "Bethesda Zips": 0.97},
        booked_impressions=280_000,
        flights=ai.merge_flights(flights),
        creative_flights=[
            {"name": "Main Spot", "start": FLIGHT_1[0], "end": FLIGHT_1[1], "impressions": 140_000},
            {"name": "Main Spot", "start": FLIGHT_2[0], "end": FLIGHT_2[1], "impressions": 136_000},
            {"name": "Anniversary", "start": FLIGHT_2[0], "end": FLIGHT_2[1], "impressions": 4_000}],
    )


WEEKS = [  # (label Monday, delivered, attributed rate) -- Sun-Sat buckets
    (date(2026, 7, 27), 1 * DAILY, 0.005),    # Jul 26-Aug 1: Aug 1 only
    (date(2026, 8, 3), 7 * DAILY, 0.010),     # Aug 2-8
    (date(2026, 8, 10), 6 * DAILY, 0.020),    # Aug 9-15: Aug 9-14
    (date(2026, 8, 24), 7 * DAILY, 0.030),    # Aug 23-29 (Aug 16-22 dark)
    (date(2026, 8, 31), 7 * DAILY, 0.040),    # Aug 30-Sep 5
]

URLS = {
    "https://club.example.com/springfield": 900,
    "https://www.example.com/": 600,
    "https://club.example.com/": 200,
    "https://marketing.example.com/bethesda-free-trial": 400,
    "https://club.example.com/bethesda": 350,
    "https://club.example.com/join": 220,
    "https://www.example.com/memberships/purchase": 120,
    "https://club.example.com/thank-you": 40,
    "https://www.example.com/thank-you/tour-scheduled": 25,
    "https://www.example.com/account/logout/success": 6,
    "https://marketing.example.com/hs/preferences-center/en/confirm": 9,
    "https://www.example.com/academy": 150,
}


def make_attribution():
    weekly = [DateSeriesPoint(day=d, delivered_impressions=n, attributed_impressions=round(n * r),
                              attributed_rate=r) for d, n, r in WEEKS]
    delivered = sum(n for _d, n, _r in WEEKS)
    attributed = sum(round(n * r) for _d, n, r in WEEKS)
    return ai.AttributionExport(
        client_name="Example Club",
        delivered_impressions=delivered, attributed_impressions=attributed,
        attributed_rate=attributed / delivered,
        attributed_unique_visitors=12_000, attributed_unique_visitor_rate=0.0429,
        by_audience=[AttributionRow("DEMO Age A25-64,HH Income 150K Plus", delivered, attributed,
                                    attributed / delivered)],
        by_market=[AttributionRow("WASHINGTON, DC (HAGRSTWN)", 263_000, attributed - 100,
                                  (attributed - 100) / 263_000),
                   AttributionRow("BALTIMORE", 17_000, 100, 100 / 17_000)],
        by_creative=[AttributionRow("Main Spot", 276_000, attributed - 400,
                                    (attributed - 400) / 276_000),
                     AttributionRow("Anniversary", 4_000, 400, 0.1)],
        by_zip=[AttributionRow("20783", 6_000, 1_000, 1_000 / 6_000),       # 2.1% share, high
                AttributionRow("20011", 8_000, 400, 0.05),                  # 2.9% share
                AttributionRow("20019", 9_000, 150, 150 / 9_000),           # below 1.2x
                AttributionRow("20010", 1_500, 300, 0.2),                   # high rate, 0.5% share
                AttributionRow("22407", 30_000, 400, 400 / 30_000)],
        by_url=dict(URLS),
        # Distinctive counts, so check 13 can prove none of them reaches the deck.
        by_recency={"00 - 03 DAYS": 1651, "04 - 07 DAYS": 1483, "08 - 11 DAYS": 1577,
                    "12 - 15 DAYS": 2003},
        by_referral_domain={"Direct": 531, "External Website Referral": 1571,
                            "Organic Search": 973, "Paid Search / Display": 4489, "Social": 37},
        by_device={"smartphone": 900, "desktop": 100},
        by_day_of_week=[AttributionRow(day, 40_000, round(40_000 * r), r) for day, r in
                        (("Mon", 0.025), ("Tue", 0.029), ("Wed", 0.029), ("Thu", 0.03),
                         ("Fri", 0.034), ("Sat", 0.02), ("Sun", 0.026))],
        weekly_trend=weekly,
        flight_start=date(2026, 7, 27), flight_end=date(2026, 8, 31),
    )


def _slide(prs, key):
    return next(s for s in prs.slides if slide_map.notes_key(s) == key)


def _slide_text(slide):
    parts = []
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.has_text_frame:
            parts.append(shape.text_frame.text)
        if shape.has_table:
            parts += [c.text_frame.text for r in shape.table.rows for c in r.cells]
    return "\n".join(parts)


def _table(slide, name):
    shape = next(sh for sh in slide.shapes if sh.name == name)
    return [[c.text_frame.text.strip() for c in r.cells] for r in shape.table.rows]


def build_deck(attribution, delivery, name, cost_per_visit=None, **kwargs):
    SCRATCH.mkdir(exist_ok=True)
    out = SCRATCH / f"{name}.pptx"
    ra.build_report_deck(str(TEMPLATE), attribution, delivery, str(out), client_name="Example Club",
                         goals_bullets=GOALS, whats_next_bullets=["Keep going."],
                         cost_per_visit=cost_per_visit,
                         goal_keywords=ra.extract_goal_keywords(GOALS), **kwargs)
    return Presentation(out), out


def build_summary(attribution, delivery, name, threads=None, cost_per_visit=None):
    SCRATCH.mkdir(exist_ok=True)
    out = SCRATCH / f"{name}.pptx"
    ra.build_summary_slide(str(TEMPLATE), str(out), attribution=attribution, delivery=delivery,
                           client_name="Example Club", threads=threads or [],
                           accepted_optimizations=[], cost_per_visit=cost_per_visit, goals=GOALS)
    return Presentation(out), out


def _draft(**fields):
    draft = {"threads": [], "goal_alignment_notes": []}
    draft.update(fields)
    return draft


# ---------------------------------------------------------------------------
# 1 -- the report period
# ---------------------------------------------------------------------------

def check_period(rep):
    print("\n1. The report period comes from flight dates, flights and dark gap shown")
    attribution, delivery = make_attribution(), make_delivery()
    rep.check("two flights -> each shown, with the year",
             ra.report_period_label(attribution, delivery) == "Aug 1–14 · Aug 23–Sep 5, 2026",
             ra.report_period_label(attribution, delivery))
    rep.check("the dark gap is named", ra.dark_gap_note(delivery) == "Dark Aug 15–22",
             ra.dark_gap_note(delivery))
    rep.check("report_period spans the flights, not the week labels (Jul 27 - Aug 31)",
             ra.report_period(attribution, delivery) == (FLIGHT_1[0], FLIGHT_2[1]),
             ra.report_period(attribution, delivery))
    merged = ai.merge_flights([
        {"start": date(2026, 4, 1), "end": date(2026, 4, 30), "impressions": 10, "uniques": 1},
        {"start": date(2026, 5, 1), "end": date(2026, 5, 31), "impressions": 10, "uniques": 1},
        {"start": date(2026, 5, 1), "end": date(2026, 5, 31), "impressions": 5, "uniques": 1}])
    rep.check("contiguous month rows merge into one flight (no false dark gap)",
             len(merged) == 1 and merged[0]["impressions"] == 25, merged)
    rep.check("without a delivery file the export's own span is used",
             ra.report_period_label(attribution, None) == "Jul 27 - Aug 31, 2026",
             ra.report_period_label(attribution, None))
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present -- recap/summary period checks")
        return
    prs, _out = build_deck(attribution, delivery, "rules_period")
    recap = _slide_text(_slide(prs, "report:recap"))
    rep.check("the recap shows both flights", "Aug 1–14 · Aug 23–Sep 5, 2026" in recap, recap)
    rep.check("...and names the dark gap", "Dark Aug 15–22" in recap, recap)
    rep.check("...and never the week-label span", "Jul 27" not in recap, recap)
    summary, _o = build_summary(attribution, delivery, "rules_period_summary")
    text = _slide_text(summary.slides[0])
    rep.check("the one-sheet subtitle shows both flights",
             "Aug 1–14 · Aug 23–Sep 5, 2026" in text, text[:300])


# ---------------------------------------------------------------------------
# 2 -- recency bucket labels
# ---------------------------------------------------------------------------

def check_recency_labels(rep):
    print("\n2. Recency bucket labels come from the source file")
    rep.check("\"12 - 15 DAYS\" reads \"12–15 days\"",
             ra.recency_bucket_label("12 - 15 DAYS") == "12–15 days")
    rep.check("zero padding is dropped, the bounds kept",
             ra.recency_bucket_label("00 - 03 DAYS") == "0–3 days")
    attribution = make_attribution()
    facts = ra.build_facts_payload(attribution, make_delivery(), goals=GOALS)
    labels = [b["label"] for b in facts["response_profile"]["recency"]["buckets"]]
    rep.check("the model's facts carry exactly the source buckets",
             labels == ["0–3 days", "4–7 days", "8–11 days", "12–15 days"], labels)
    draft = _draft(response_profile_narrative=(
        "About 3 in 10 visits with timing data came 12-18 days after exposure. "
        "About 1 in 4 came within 0–3 days."))
    bad = app._recency_range_violations(draft, facts, attribution)
    rep.check("a drafted \"12-18 days\" (not a source bucket) is caught",
             [t for _f, t in bad] == ["12-18 days"], bad)
    cleaned = app._strip_rule_violations(copy.deepcopy(draft), facts, attribution)
    rep.check("...and its sentence removed, the real-bucket sentence kept",
             cleaned["response_profile_narrative"] == "About 1 in 4 came within 0–3 days.",
             cleaned["response_profile_narrative"])
    good = _draft(response_profile_narrative="About 3 in 10 responded 12–15 days after exposure.")
    rep.check("a source bucket range passes", not app._recency_range_violations(good, facts, attribution))


# ---------------------------------------------------------------------------
# 4 -- no claims about dimensions without attribution data
# ---------------------------------------------------------------------------

def check_unsupported_dimension_claims(rep):
    print("\n4. Post-draft fact check: no claim about a dimension with no attribution data")
    attribution = make_attribution()
    facts = ra.build_facts_payload(attribution, make_delivery(), goals=GOALS)
    draft = _draft(threads=[{
        "head": "Weight Toward Strong Areas", "anchor": "goal", "goal_ref": GOALS[0],
        "finding": "Springfield pages drew strong interest.",
        "meaning": "The clubs are resonating.",
        "action": "Continue Premion Streaming TV with targeting weighted toward the ZIP codes and "
                  "dayparts showing the highest attributed rates."}],
        delivery_narrative="Most impressions ran in the 7PM–midnight daypart.")
    hits = app._unsupported_dimension_violations(draft, facts)
    rep.check("the St. James daypart action is rejected",
             len(hits) == 1 and hits[0][2] == "daypart" and "dayparts" in hits[0][1], hits)
    rep.check("a delivery-only daypart sentence (no response claim) passes",
             all("7PM" not in sentence for _f, sentence, _d in hits), hits)
    enforced = app._enforce_draft_rules(copy.deepcopy(draft), facts, attribution)
    action = enforced["threads"][[t["head"] for t in enforced["threads"]].index(
        "Weight Toward Strong Areas")]["action"]
    rep.check("the claim never ships -- the action is removed after the retry", action is None, action)
    rep.check("...and the removal is named for review",
             any("no attribution data" in n for n in enforced.get("_review_notes") or []),
             enforced.get("_review_notes"))
    no_device = dict(facts, device=None)
    device_draft = _draft(zip_narrative="Smartphone visitors responded at the highest rate.")
    rep.check("device claims are rejected when the export has no device data",
             app._unsupported_dimension_violations(device_draft, no_device), None)
    rep.check("...and allowed when it does",
             not app._unsupported_dimension_violations(device_draft, facts), None)


# ---------------------------------------------------------------------------
# 5 -- the ZIP slide
# ---------------------------------------------------------------------------

def check_zip_slide(rep):
    print("\n5. ZIP slide: table and subtitle agree; heat-map title matches its legend")
    attribution = make_attribution()
    rows, _dropped, qualified = ra.zip_slide_rows(attribution)
    zips = [r["zip"] for r in rows]
    rep.check("only ZIPs at 1.2x+ on real volume, sorted by multiple",
             qualified and zips == ["20783", "20011"], zips)
    rep.check("a high-rate ZIP under the volume floor (20010, 0.5%) is left out",
             "20010" not in zips, zips)
    facts = ra.build_facts_payload(attribution, make_delivery(), goals=GOALS)
    rep.check("the model's zip rows are exactly the slide's rows",
             [r["zip"] for r in facts["zip"]["rows"]] == zips, facts["zip"]["rows"])
    flat = make_attribution()
    flat.by_zip = [AttributionRow("22407", 30_000, 300, 0.01)]
    _rows, _d, flat_qualified = ra.zip_slide_rows(flat)
    rep.check("with none qualifying, qualified is False (volume rows, subtitle says so)",
             not flat_qualified)
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present -- zip slide deck checks")
        return
    prs, _out = build_deck(attribution, make_delivery(), "rules_zip",
                           headline_notes={"zip": "ZIP codes above the campaign average."})
    slide = _slide(prs, "report:zip_analysis")
    table = _table(slide, "TopZipTable")
    rep.check("the table holds exactly the qualifying ZIPs, by multiple",
             [r[0] for r in table[1:]] == zips, table)
    text = _slide_text(slide)
    rep.check("the subtitle is Python's own, describing that set (a drafted note isn't used)",
             ra.ZIP_SUBTITLE_QUALIFIED in text and "above the campaign average." not in text, text)
    rep.check("the heat map's title names the attributed rate (its legend's metric)",
             ra.ZIP_MAP_HEADER in text and "ATTRIBUTED IMPRESSIONS" not in text, text)


# ---------------------------------------------------------------------------
# 8 -- trend detection
# ---------------------------------------------------------------------------

def check_trend(rep):
    print("\n8. Trend: real-dated weeks, dark weeks, flights; qualification and its uses")
    attribution, delivery = make_attribution(), make_delivery()
    weeks = ra.delivery_weeks(attribution, delivery)
    rep.check("weeks are labelled with the days they delivered on, the dark week marked",
             [(w["label"], w["dark"]) for w in weeks] == [
                 ("Aug 1", False), ("Aug 2–8", False), ("Aug 9–14", False),
                 ("Aug 16–22", True), ("Aug 23–29", False), ("Aug 30–Sep 5", False)],
             [(w["label"], w["dark"]) for w in weeks])
    trend = ra.trend_facts(attribution, delivery)
    rep.check("four consecutive rises -> qualifies on rising weeks",
             trend["consecutive_rises"] == 4 and "rising_weeks" in trend["reasons"], trend)
    flights = trend["flights"]
    rep.check("rate by flight comes from the weeks inside each flight",
             len(flights) == 2 and abs(flights[0]["attributed_rate"] - 1950 / 140_000) < 1e-9
             and abs(flights[1]["attributed_rate"] - 4900 / 140_000) < 1e-9, flights)
    rep.check("the flight swing is material, stated as a multiple",
             trend["flight_swing"]["material"] and trend["flight_swing"]["multiple"] == 2.5,
             trend["flight_swing"])
    sentence = ra.trend_sentence(trend)
    rep.check("the trend sentence carries one percentage",
             sentence and sentence.count("%") == 1, sentence)

    flat = make_attribution()
    flat.weekly_trend = [DateSeriesPoint(d, n, round(n * 0.02), 0.02) for d, n, _r in WEEKS]
    flat_trend = ra.trend_facts(flat, delivery)
    rep.check("a flat series doesn't qualify", not flat_trend["qualifies"], flat_trend["reasons"])

    one_flight = make_delivery(two_flights=False)
    short = make_attribution()
    short.weekly_trend = [DateSeriesPoint(d, n, round(n * r), r) for d, n, r in WEEKS[:3]]
    short_trend = ra.trend_facts(short, one_flight)
    rep.check("two rises in one flight doesn't qualify (needs 3+, or a flight swing)",
             not short_trend["qualifies"] and short_trend["consecutive_rises"] == 2, short_trend)

    facts = ra.build_facts_payload(attribution, delivery, goals=GOALS)
    enforced = app._enforce_draft_rules(_draft(threads=[
        {"head": "Springfield Leads", "anchor": "goal", "goal_ref": GOALS[0],
         "finding": "900 visitors viewed the Springfield page.", "meaning": "Location pages work.",
         "action": None}]), facts, attribution)
    trend_threads = [t for t in enforced["threads"] if t.get("protected") and "week" in t["finding"]]
    rep.check("a qualifying trend the draft left out is added as a protected highlight thread",
             len(trend_threads) == 1, enforced["threads"])
    rep.check("...with the keep-it-continuous action (rising, across a dark gap)",
             trend_threads and "continuously" in (trend_threads[0].get("action") or ""),
             trend_threads)
    highlights, _t, whats_next = ra.distribute_threads(enforced["threads"])
    rep.check("the trend is a highlight and drives a What's Next item",
             any(h == trend_threads[0]["head"] for h, _f in highlights)
             and any("continuously" in w for w in whats_next), (highlights, whats_next))
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present -- one-sheet bottom line check")
        return
    summary, _o = build_summary(attribution, delivery, "rules_trend_summary",
                                threads=enforced["threads"])
    bottom = next(sh.text_frame.text for sh in summary.slides[0].shapes if sh.name == "BottomLine")
    rep.check("the one-sheet bottom line leads with the trend", bottom.startswith(sentence), bottom)


# ---------------------------------------------------------------------------
# 10 -- one percentage per fact
# ---------------------------------------------------------------------------

def check_percent_discipline(rep):
    print("\n10. At most one percentage per fact")
    double = _draft(threads=[{"head": "Bethesda Free Trial", "anchor": "goal", "meaning": "m",
                              "finding": "The Bethesda Free Trial page accounted for 10% of attributed "
                                         "page visits and reached 21.7% of attributed visitors."}],
                    url_intent_narrative="Springfield (23%) led, with Bethesda at 10% (10.3%).",
                    response_profile_narrative="About 3 in 10 (29.9%) responded 12–15 days out.")
    hits = [field for field, _s in app._double_percent_violations(double)]
    rep.check("share-of-visits + share-of-visitors for one page is caught", "thread 1 finding" in hits,
             hits)
    rep.check("a percentage repeated in parentheses is caught", "url_intent_narrative" in hits, hits)
    rep.check("a share phrase restated as a percentage is caught",
             "response_profile_narrative" in hits, hits)
    fine = _draft(threads=[{"head": "Intent", "anchor": "goal", "meaning": "m",
                            "finding": "Purchase pages drew 3.7% of page visits, and lead pages 2.0%."}],
                  attribution_narrative="Flight 2 ran at 4.04%, up from 1.49% in flight 1.")
    rep.check("two different items, or a from/to comparison, pass",
             not app._double_percent_violations(fine), app._double_percent_violations(fine))
    # A live St. James draft wrote Friday's 3.43% rate as "0.03%"; every sub-1
    # token rounded to "0" and "traced" to a 0% conversion rate.
    facts = ra.build_facts_payload(make_attribution(), make_delivery(), goals=GOALS)
    bad = app._attr_draft_number_violations(
        [("x", "Friday led at 0.03%, Fri at 3.40%, the campaign at 2.42%.")], facts)
    rep.check("a mis-scaled sub-1% rate doesn't trace to the facts",
             [t for _f, t in bad] == ["0.03%"], bad)
    rep.check("...while the real rates do", not app._attr_draft_number_violations(
        [("x", "Fri at 3.40%.")], facts))
    zip_double = _draft(zip_narrative="College Park (20783) posted the highest attributed rate at "
                                      "20.98% -- 7.59x the average -- on 1.8% of attributed "
                                      "impressions. Takoma Park ran at 10.70% (3.87x).")
    rep.check("a ZIP's rate plus its impression share in one sentence is caught (live draft)",
             [f for f, _s in app._double_percent_violations(zip_double)] == ["zip_narrative"],
             app._double_percent_violations(zip_double))
    spill = _draft(threads=[{"head": "DC Leads Markets", "anchor": "goal", "meaning": "m",
                             "finding": "Washington, DC ran at 2.88%, well above Baltimore at 0.92%."},
                            {"head": "Pages", "anchor": "goal", "finding": "900 visitors viewed "
                             "Springfield.", "meaning": "Location pages work."}])
    spill_facts = ra.build_facts_payload(make_attribution(), make_delivery(), goals=GOALS)
    rep.check("Baltimore (6% of impressions) is listed as spillover",
             spill_facts["market"]["spillover"] == ["Baltimore"], spill_facts["market"]["spillover"])
    stripped = app._strip_rule_violations(copy.deepcopy(spill), spill_facts)
    rep.check("a thread about a spillover market is dropped whole, the rest kept",
             [t["head"] for t in stripped["threads"]] == ["Pages"], stripped["threads"])
    shares = _draft(response_profile_narrative="TV prompted search: about 7 in 10 visits with "
                                               "referral data came from paid or organic search.")
    rep.check("\"about 7 in 10 visits\" is a share, not a count",
             not app._recency_referral_count_violations(shares))


# ---------------------------------------------------------------------------
# 11 / 14 / 15 -- pages, page groups, goal path
# ---------------------------------------------------------------------------

def check_pages_groups_path(rep):
    print("\n11/14/15. Single-page visitor counts, goal page groups, the goal path")
    attribution = make_attribution()
    rows = ra.top_url_rows(attribution, limit=8)
    labels = [r["label"] for r in rows]
    rep.check("two sites' homepages are two rows, never summed",
             "Homepage" in labels and any("homepage" in l and l != "Homepage" for l in labels)
             and not any(r["visits"] == "800" for r in rows), labels)
    terms = [t["label"] for t in ra.goal_terms(GOALS, None, attribution)]
    rep.check("goal terms are the goal's own entities the site has pages for",
             terms[:3] == ["Springfield", "Bethesda", "Performance Club"] or
             {"Springfield", "Bethesda"} <= set(terms), terms)
    groups = ra.goal_page_groups(attribution, GOALS)["classes"]
    rep.check("page groups lead with goal entities and include confirmation pages",
             groups[0]["label"] in ("Springfield", "Bethesda")
             and any(g["intent"] == "confirmation" for g in groups), [g["label"] for g in groups])
    rep.check("\"Other\" is never the top row (always last)",
             all(g["intent"] != "other" for g in groups[:-1]), [g["label"] for g in groups])
    rep.check("each group's pages carry their own visitor counts",
             all(isinstance(p["visitors"], int) for g in groups for p in g["pages"]), None)
    other_heavy = make_attribution()
    other_heavy.by_url = {"https://ex.com/members-area/benefits": 500, "https://ex.com/news": 200,
                          "https://ex.com/": 100}
    regrouped = ra.goal_page_groups(other_heavy, ["Grow membership"])["classes"]
    rep.check("with Other over 40%, pages are re-read against the goal terms",
             regrouped[0]["label"] == "Membership" and regrouped[0]["visits"] == 500,
             [(g["label"], g["visits"]) for g in regrouped])
    path = ra.goal_path(attribution, GOALS)
    rep.check("goal path: Explored -> Started sign-up -> Purchase flow -> Reached a confirmation page",
             [p["step"] for p in path] == list(ra.GOAL_PATH_STEPS), path)
    rep.check("each step shows its top page's own count (never a step total)",
             [p["visitors"] for p in path] == [900, 400, 120, 40], path)
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present -- URL slide checks")
        return
    prs, _out = build_deck(attribution, make_delivery(), "rules_url")
    slide = _slide(prs, "report:url_report")
    path_table = _table(slide, "IntentSummaryTable")
    rep.check("the traffic-mix table is the goal path: Step | Top page | Visitors",
             path_table[0] == ["Step", "Top page", "Visitors"]
             and [r[0] for r in path_table[1:]] == list(ra.GOAL_PATH_STEPS), path_table)
    pages = _table(slide, "TopUrlTable")
    rep.check("top pages: Page | Visitors, counts only",
             pages[0] == ["Page", "Visitors"]
             and not any("%" in cell for row in pages for cell in row), pages)
    rep.check("no summed total row", not any(r[0].lower() in ("total", "all pages") for r in pages),
             pages)
    rep.check("no percentage anywhere in either table",
             not any("%" in c for r in path_table for c in r), path_table)


# ---------------------------------------------------------------------------
# 12 -- CPV
# ---------------------------------------------------------------------------

def check_cpv(rep):
    print("\n12. CPV is a tile with the toggle on, absent with it off")
    attribution, delivery = make_attribution(), make_delivery()
    cpv = ra.compute_cost_per_visit(227_040.0, attribution.attributed_unique_visitors, 0, None)
    facts_off = ra.build_facts_payload(attribution, delivery, goals=GOALS, budget=20_000.0)
    rep.check("off: no cost per visitor reaches the model, even with a budget",
             facts_off["cost_per_visit"] is None
             and facts_off["budget"]["cost_per_attributed_visit"] is None, facts_off["budget"])
    leak = _draft(zip_narrative="Cost per site visitor came to $18.92. 20783 led.")
    stripped = app._strip_rule_violations(copy.deepcopy(leak), facts_off, attribution)
    rep.check("off: a drafted CPV sentence is removed", stripped["zip_narrative"] == "20783 led.",
             stripped["zip_narrative"])
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present -- CPV tile checks")
        return
    on_deck, _o = build_deck(attribution, delivery, "rules_cpv_on", cost_per_visit=cpv)
    off_deck, _o2 = build_deck(attribution, delivery, "rules_cpv_off")
    on_summary, _s = build_summary(attribution, delivery, "rules_cpv_on_summary", cost_per_visit=cpv)
    off_summary, _s2 = build_summary(attribution, delivery, "rules_cpv_off_summary")
    hi_on = _slide_text(_slide(on_deck, "report:highlights"))
    hi_off = _slide_text(_slide(off_deck, "report:highlights"))
    sum_on, sum_off = _slide_text(on_summary.slides[0]), _slide_text(off_summary.slides[0])
    rep.check("on: the highlights slide has the tile \"$18.92 / Cost per site visitor\"",
             "$18.92" in hi_on and ra.CPV_TILE_LABEL in hi_on, hi_on[:400])
    rep.check("on: the one-sheet has the same tile", "$18.92" in sum_on and ra.CPV_TILE_LABEL in sum_on,
             sum_on[:400])
    names_on = {sh.name for sh in _slide(on_deck, "report:highlights").shapes}
    rep.check("on: no CPV text line under the tiles", "CostPerVisitNote" not in names_on, names_on)
    rep.check("off: absent from the highlights slide",
             "$18.92" not in hi_off and ra.CPV_TILE_LABEL not in hi_off, hi_off[:400])
    rep.check("off: absent from the one-sheet",
             "$18.92" not in sum_off and ra.CPV_TILE_LABEL not in sum_off, sum_off[:400])
    rep.check("no Unique Visitor Rate tile on either one-sheet",
             "Unique Visitor Rate" not in sum_on + sum_off, None)
    for label, path in (("on deck", _o), ("on summary", _s)):
        rep.check(f"{label} passes package_check", package_check.check_package(str(path)) == [],
                 package_check.check_package(str(path)))


# ---------------------------------------------------------------------------
# 13 -- recency and referral are shares only
# ---------------------------------------------------------------------------

_RECENCY_REFERRAL_COUNTS = ("1651", "1,651", "1483", "1,483", "1577", "1,577", "2003", "2,003",
                            "531", "1571", "1,571", "973", "4489", "4,489", "37",
                            "6714", "6,714", "7601", "7,601")   # the two tabs' totals too


def check_recency_referral_shares(rep):
    print("\n13. Recency and referral: shares only, bar charts, a footnote, no counts")
    attribution = make_attribution()
    facts = ra.build_facts_payload(attribution, make_delivery(), goals=GOALS)
    profile = facts["response_profile"]
    has_counts = any("visitors" in b or "total" in profile[k]
                     for k in ("recency", "referral") for b in (profile[k].get("buckets")
                                                                or profile[k].get("sources")))
    rep.check("the model's facts carry no visitor counts from either tab", not has_counts, profile)
    rep.check("search share >= 50% -> the TV-prompted-search flag is set",
             profile["referral"]["search_prompted"], profile["referral"])
    counted = _draft(response_profile_narrative=(
        "Of the 764 visitors with referral data, most came from paid search. "
        "About 3 in 10 responded 12–15 days after exposure."))
    rep.check("a count in a referral sentence is caught",
             app._recency_referral_count_violations(counted), None)
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present -- response profile slide checks")
        return
    prs, out = build_deck(attribution, make_delivery(), "rules_response")
    slide = _slide(prs, "report:response_profile")
    names = {sh.name for sh in slide.shapes}
    text = _slide_text(slide)
    rep.check("the Source/Visitors/Share table is replaced by a chart",
             "ReferralTable" not in names
             and sum(1 for sh in slide.shapes if sh.shape_type == 13) >= 2, names)
    rep.check("the coverage footnote sits under the charts", ra.RECENCY_REFERRAL_FOOTNOTE in text, text)
    rep.check("the TV-prompted-search line is in the narrative", "TV prompted search" in text, text)
    deck_text = "\n".join(_slide_text(s) for s in prs.slides)
    summary, _o = build_summary(attribution, make_delivery(), "rules_response_summary")
    deck_text += "\n" + _slide_text(summary.slides[0])
    import re
    numbers = set(re.findall(r"\b\d[\d,]*\b", deck_text))
    leaked = [n for n in _RECENCY_REFERRAL_COUNTS if n in numbers]
    rep.check("no count from either tab appears anywhere in the deck or one-sheet text",
             not leaked, leaked)


# ---------------------------------------------------------------------------
# 16 -- confirmation pages
# ---------------------------------------------------------------------------

def check_confirmation_pages(rep):
    print("\n16. Confirmation pages: a standard highlight, \"reached a confirmation page\"")
    attribution = make_attribution()
    pages = ra.confirmation_pages(attribution)
    rep.check("thank-you and tour-scheduled pages count; logout 'success' and preferences "
             "'confirm' don't", [p["visitors"] for p in pages] == [40, 25], pages)
    facts = ra.build_facts_payload(attribution, make_delivery(), goals=GOALS)
    enforced = app._enforce_draft_rules(_draft(threads=[]), facts, attribution)
    confirm = [t for t in enforced["threads"] if "confirmation page" in (t.get("finding") or "")]
    rep.check("missing from the draft -> added as a protected highlight thread",
             len(confirm) == 1 and confirm[0]["protected"], enforced["threads"])
    rep.check("...worded \"reached a confirmation page\", each page with its own count",
             confirm and "reached a confirmation page" in confirm[0]["finding"]
             and "40" in confirm[0]["finding"] and "25" in confirm[0]["finding"], confirm)
    wording = _draft(threads=[{"head": "Sign-ups", "anchor": "goal", "meaning": "Members signed up.",
                               "finding": "40 visitors converted on the thank-you page."}])
    rep.check("\"converted\" / \"signed up\" are caught with no conversion data",
             len(app._conversion_word_violations(wording, facts)) == 2,
             app._conversion_word_violations(wording, facts))
    cleaned = app._enforce_draft_rules(copy.deepcopy(wording), facts, attribution)
    text = " ".join(str(t.get(k) or "") for t in cleaned["threads"] for k in ("finding", "meaning"))
    rep.check("...and never ship", "converted" not in text and "signed up" not in text, text)


# ---------------------------------------------------------------------------
# The real St. James pair
# ---------------------------------------------------------------------------

def check_st_james(rep):
    print("\nReal St. James export (Aug 2026)")
    if not (STJ_ATTRIBUTION.exists() and STJ_DELIVERY.exists()):
        rep.skip("St. James fixtures not present")
        return
    attribution = ai.parse_attribution_export(str(STJ_ATTRIBUTION))
    delivery = ai.parse_delivery_export(str(STJ_DELIVERY))
    rep.check("period: Aug 1–14 · Aug 23–Sep 5, 2026",
             ra.report_period_label(attribution, delivery) == "Aug 1–14 · Aug 23–Sep 5, 2026",
             ra.report_period_label(attribution, delivery))
    trend = ra.trend_facts(attribution, delivery)
    rep.check("trend qualifies on both counts (4 rises; flight 2 at 2.7x)",
             trend["consecutive_rises"] == 4 and trend["flight_swing"]["multiple"] == 2.7, trend)
    comparison = ra.creative_comparison(attribution, delivery)
    rates = {c["name"]: round(c["attributed_rate"] * 100, 1) for c in comparison["creatives"]}
    rep.check("creatives compared over flight 2 only: Anniversary 8.7% vs main 3.9%, directional",
             rates.get("TSJV26302 TSJ Anniversary 30") == 8.7 and rates.get("TSJV26301_Rev") == 3.9
             and comparison["directional"], comparison)
    rep.check("Baltimore (6% spillover) isn't a breakdown -- the slide drops",
             not ra.attribution_breakdown_applies(attribution, delivery))
    rep.check("bought geos label the Geography tile",
             ra.bought_geo_labels(delivery) == ["Springfield ZIPs", "Bethesda ZIPs"],
             ra.bought_geo_labels(delivery))
    rep.check("the recency tab's own buckets: 12–15 days, never 12–18",
             "12–15 days" in [ra.recency_bucket_label(b) for b in attribution.by_recency])
    path = ra.goal_path(attribution, GOALS)
    rep.check("goal path: 1,399 -> 627 -> 192 -> 49",
             [p["visitors"] for p in path] == [1399, 627, 192, 49], path)


def main():
    rep = Report()
    check_period(rep)
    check_recency_labels(rep)
    check_unsupported_dimension_claims(rep)
    check_zip_slide(rep)
    check_trend(rep)
    check_percent_discipline(rep)
    check_pages_groups_path(rep)
    check_cpv(rep)
    check_recency_referral_shares(rep)
    check_confirmation_pages(rep)
    check_st_james(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped "
          f"out of {total}")
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
