"""Attribution report fixes from the multi-audience bank review (2026-11).

    python tests/test_multi_audience_report.py

Synthetic exports only (CI); tests/test_multi_audience_real.py checks the
real fixture locally. One section per review item:

  1. Spillover: a market is spillover only under 2% of impressions AND
     outside the targeting -- the old 10% floor dropped the campaign's top
     market -- and the narrative's market count is the table's.
  2. Campaign landing pages: suggested from an outsized share of visits; the
     rep's own lead the goal path and the key pages table.
  3. Key pages and why they matter: page groups, their visits and a fixed
     reason each matters; existing customers last and never interest.
  4. URL normalization: query strings, fragments, ".html" and a second
     domain for the same site don't split a page.
  5. Audiences side by side, from computed facts, with both caveats.
  6. A market is judged against its own audience when audiences differ --
     never cut against a blended average.
  7. Still-accruing weeks: flagged, never a decline; the export date shown,
     and the rep warned when it's too early.
  8. The key pages slide's footnote never overlaps the table or the panel.
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


# ---------------------------------------------------------------------------
# Items 2-8
# ---------------------------------------------------------------------------
import subprocess  # noqa: E402
import tempfile  # noqa: E402
from datetime import date, datetime, timedelta  # noqa: E402

from attribution_import import DateSeriesPoint  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Emu  # noqa: E402

import slide_map  # noqa: E402

TMP = Path(tempfile.mkdtemp())
SITE = "https://www.example.bank"
URLS = {
    f"{SITE}/": 1300,
    f"{SITE}/high-interest-business.html": 2600,
    "https://www.example.com/high-interest-business.html": 40,       # the old domain
    f"{SITE}/high-interest-business.html?utm_source=ott&gclid=abc": 30,
    f"{SITE}/high-interest-personal.html": 950,
    f"{SITE}/personal/checking/compare-accounts.html": 140,
    f"{SITE}/personal/savings/cds.html": 60,
    f"{SITE}/rates.html": 40,
    f"{SITE}/business/savings/compare.html": 120,
    f"{SITE}/business/commercial-credit.html": 70,
    f"{SITE}/about/locations-hours.html": 110,
    f"{SITE}/about/contact-us.html": 95,
    f"{SITE}/contact-us/thank-you.html": 9,
    f"{SITE}/personal/account-services/online-banking.html": 65,
    f"{SITE}/notice-details.html": 50,
    "https://www.example.com/login": 3,
    f"{SITE}/about/careers.html": 37,
    f"{SITE}/privacy-policy.html": 12,
}
BUSINESS = "DEMO Career Business Owners Executives"
RETAIL = "FIN Banking Intenders"


def bank_attribution(exported_on=date(2026, 12, 20)):
    """Two audiences at 3.57% and 5.06%, each with its own creatives (their
    impressions split exactly into the two audiences), two RFPIDs."""
    creatives = [row("Example Commercial Chef 30", 300_000, 0.0367),
                 row("Example Commercial Chef 15", 380_000, 0.0365),
                 row("Example Commercial Chef 15", 140_000, 0.0321),
                 row("Example Retail Mechanic 30", 250_000, 0.0520),
                 row("Example Retail Mechanic 15", 300_000, 0.0500)]
    audiences = [AttributionRow(BUSINESS, 820_000, 29_272, 0.0357),
                 AttributionRow(RETAIL, 550_000, 27_830, 0.0506)]
    weeks = [DateSeriesPoint(date(2026, 9, 14) + timedelta(weeks=i), 170_000, 7_000 + 200 * i,
                             (7_000 + 200 * i) / 170_000) for i in range(6)]
    weeks.append(DateSeriesPoint(date(2026, 10, 26), 180_000, 4_000, 4_000 / 180_000))
    total = 1_370_000
    return ai.AttributionExport(
        delivered_impressions=total, attributed_impressions=57_102,
        attributed_rate=57_102 / total, attributed_unique_visitors=4_100,
        by_audience=audiences, by_creative=creatives, by_market=list(MARKETS), by_url=dict(URLS),
        by_zip=[row("23510", 40_000, 0.06), row("23219", 30_000, 0.05), row("22901", 20_000, 0.08),
                row("20001", 15_000, 0.013)],
        zip_heat={"23510": 2_400, "23219": 1_500, "22901": 1_600, "20001": 200},
        rfpid_breakdown=[{"rfpid": "RFPID-700002", "delivered_impressions": 820_000,
                          "attributed_impressions": 29_272, "attributed_rate": 0.0357},
                         {"rfpid": "RFPID-700001", "delivered_impressions": 550_000,
                          "attributed_impressions": 27_830, "attributed_rate": 0.0506}],
        weekly_trend=weeks, flight_start=date(2026, 9, 14), flight_end=date(2026, 11, 1),
        exported_on=exported_on)


def landing(attribution):
    return ra.with_landing_pages(attribution, [
        {"pattern": "/high-interest-business.html", "audience": BUSINESS},
        {"pattern": "*/high-interest-personal*", "audience": RETAIL}])


def check_urls():
    print("\n4. URL normalization")
    key = ra._page_key
    check("the old and new domains are one site",
          key(f"{SITE}/high-interest-business.html") ==
          key("https://www.example.com/high-interest-business.html"))
    check("query strings (utm, gclid) and fragments drop",
          key(f"{SITE}/high-interest-business.html?utm_source=ott&gclid=abc#top")
          == key(f"{SITE}/high-interest-business"))
    check("a tracking leftover after a space drops",
          key("https://example.com/forgot%208601508280a74549") == key("https://example.com/forgot"))
    rows = {r["label"]: r["visitors"] for r in ra.page_rows(bank_attribution())}
    check("their visits merge into one page", rows.get("High Interest Business") == 2670, rows)


def check_landing_pages():
    print("\n2. Campaign landing pages")
    a = bank_attribution()
    check("pages with an outsized share of visits are suggested",
          [ra._page_key(u)[1] for u in ra.suggest_landing_pages(a)] ==
          ["/high-interest-business", "/high-interest-personal"], ra.suggest_landing_pages(a))
    check("the homepage is never suggested",
          not any(ra._page_key(u)[1] == "" for u in ra.suggest_landing_pages(a)))
    url = f"{SITE}/high-interest-personal.html"
    check("a full URL, a path, a * pattern and a word all match",
          all(ra.landing_page_matches(url, p) for p in
              ("https://www.example.com/high-interest-personal", "/high-interest-personal.html",
               "*/high-interest-*", "high-interest-personal")))
    check("...and a different page doesn't", not ra.landing_page_matches(url, "/high-interest-business"))
    plain = [s["key"] for s in ra.goal_path_combined(a, ["Grow deposits"])]
    path = ra.goal_path_combined(landing(a), ["Grow deposits"])
    check("with the rep's landing pages set, they're the goal path's first step",
          path[0]["key"] == "offer" and path[0]["page_visits"] == 2670 + 950, path[0])
    check("suggestions alone never reshape a goal path", "offer" not in plain, plain)


def check_key_pages():
    print("\n3. Key pages and why they matter")
    group = ra.key_page_group
    cases = {f"{SITE}/": "homepage", f"{SITE}/about/locations-hours.html": "ready",
             f"{SITE}/about/contact-us.html": "ready", f"{SITE}/contact-us/thank-you.html": "confirmation",
             f"{SITE}/personal/account-services/online-banking.html": "service",
             f"{SITE}/notice-details.html": "service", "https://x.com/login": "service",
             f"{SITE}/about/careers.html": "other", f"{SITE}/privacy-policy.html": "other",
             f"{SITE}/business/savings/compare.html": "research_business",
             f"{SITE}/personal/savings/cds.html": "research_personal",
             f"{SITE}/find-a-banker.html": "ready", f"{SITE}/open-an-account.html": "ready"}
    wrong = {u: (group(u), want) for u, want in cases.items() if group(u) != want}
    check("each page lands in its group", not wrong, wrong)
    facts = ra.key_pages_facts(landing(bank_attribution()))
    keys = [r["key"] for r in facts["rows"]]
    check("offers first, then homepage, research, ready to act, confirmations, customer service last",
          keys == ["offer", "offer", "homepage", "research_personal", "research_business", "ready",
                   "confirmation", "service"], keys)
    by = {r["key"]: r for r in facts["rows"]}
    check("unlabelled product pages count as personal when the site splits personal/business",
          by["research_personal"]["page_visits"] == 140 + 60 + 40, by["research_personal"]["page_visits"])
    check("existing customers are excluded from interest",
          facts["interest_page_visits"] == facts["total_page_visits"] - by["service"]["page_visits"]
          - facts["other_page_visits"], facts)
    check("careers and privacy are not a row", all("career" not in r["group"].lower() for r in facts["rows"]))
    check("an offer names its audience's offer",
          "Career Business Owners Executives audience's offer" in facts["rows"][0]["why"],
          facts["rows"][0]["why"])
    check("every group carries a fixed reason", all(r["why"] for r in facts["rows"]))
    check("the subtitle states the offers' share and where visitors went",
          ra.key_pages_headline(facts).startswith("The two offer pages drew 63% of all page visits"),   # 3,620 of 5,731
          ra.key_pages_headline(facts))
    one_side = ra.key_pages_facts(ra.with_landing_pages(ai.AttributionExport(
        by_url={f"{SITE}/": 10, f"{SITE}/personal/checking.html": 5}), []))
    check("research isn't split when only one side has pages",
          [r["key"] for r in one_side["rows"]] == ["homepage", "research"], one_side["rows"])


def template_v0_18():
    out = TMP / "REPORT_MASTER_v0_18.pptx"
    if not out.exists():
        subprocess.run([sys.executable, str(REPO / "build_v0_18.py"),
                        str(REPO / "REPORT_MASTER_v0_17.pptx"), str(out)], check=True,
                       capture_output=True)
    return out


def build(attribution, name, **kw):
    out = TMP / f"{name}.pptx"
    ra.build_report_deck(str(template_v0_18()), attribution, delivery(), str(out),
                         client_name="Example Bank", goals_bullets=["Grow deposits"],
                         whats_next_bullets=["Keep going"], **kw)
    prs = Presentation(str(out))
    return prs, [slide_map.notes_key(s) for s in prs.slides]


def _box(shape):
    return (int(shape.left), int(shape.top), int(shape.left) + int(shape.width),
            int(shape.top) + int(shape.height))


def check_decks():
    print("\n3/5/8. The slides")
    if not (REPO / "REPORT_MASTER_v0_17.pptx").exists():
        print("  SKIP  REPORT_MASTER_v0_17.pptx not present")
        return
    a = landing(bank_attribution())
    ott = ai.OTTRetargetingExport(campaign_lines=[
        {"name": "x_RFPID-700002", "rfpid": "RFPID-700002", "impressions": 480_000, "clicks": 567,
         "ctr": 567 / 480_000},
        {"name": "x_RFPID-700001", "rfpid": "RFPID-700001", "impressions": 342_000, "clicks": 396,
         "ctr": 396 / 342_000}])
    prs, keys = build(a, "bank", ott=ott)
    check("key pages replaces the goal-path URL slide",
          "report:key_pages" in keys and "report:url_report" not in keys, keys)
    check("two audiences add the audience comparison", "report:audience_comparison" in keys, keys)
    check("every token filled", not [sh.text_frame.text for s in prs.slides for sh in s.shapes
                                     if sh.has_text_frame and "{{" in sh.text_frame.text])
    kp = prs.slides[keys.index("report:key_pages")]
    shapes = {sh.name: sh for sh in kp.shapes}
    table, foot, panel = (_box(shapes["KeyPagesTable"]), _box(shapes["KeyPagesFootnote"]),
                          _box(shapes["Shape 7"]))
    table_bottom = int(shapes["KeyPagesTable"].top) + sum(
        int(r.height) for r in shapes["KeyPagesTable"].table.rows)
    check("8. the footnote sits under the table and above the panel -- nothing overlaps",
          table_bottom <= foot[1] and foot[3] <= panel[1], (table_bottom, foot, panel))
    _prs, keys = build(a, "compare", ott=ott, compare_locations=True)
    check("'Compare locations' keeps the goal-path URL slide",
          "report:url_report" in keys and "report:key_pages" not in keys, keys)
    one = bank_attribution()
    one.by_audience = one.by_audience[:1]
    one.rfpid_breakdown = one.rfpid_breakdown[:1]
    _prs, keys = build(one, "one")
    check("one audience: no audience comparison", "report:audience_comparison" not in keys, keys)
    recap = prs.slides[0]
    note = next((sh.text_frame.text for sh in recap.shapes if sh.name == "ExportDateNote"), "")
    check("7. the recap shows the export date", note.startswith("Data as of Dec 20, 2026."), note)


def check_audiences():
    print("\n5. Audiences side by side")
    a = landing(bank_attribution())
    ott = ai.OTTRetargetingExport(campaign_lines=[
        {"name": "x", "rfpid": "RFPID-700001", "impressions": 1000, "clicks": 1, "ctr": 0.001}])
    f = ra.audience_comparison_facts(a, ott)
    u = {x["raw"]: x for x in f["units"]}
    check("each audience's own impressions, attributed and rate",
          (u[RETAIL]["delivered"], u[RETAIL]["attributed"], u[RETAIL]["rate"]) == (550_000, 27_830, 0.0506))
    check("shares of impressions and of attributed",
          round(u[RETAIL]["impression_share"], 3) == round(550_000 / 1_370_000, 3)
          and round(u[RETAIL]["attributed_share"], 3) == round(27_830 / 57_102, 3))
    check("creatives split exactly into the audiences give each its rate range",
          u[RETAIL]["creative_range"] == (0.05, 0.052) and u[BUSINESS]["creative_range"] == (0.0321, 0.0367),
          (u[RETAIL]["creative_range"], u[BUSINESS]["creative_range"]))
    check("the matching offer page is the rep's landing page for that audience",
          u[RETAIL]["offer"] == {"page": "High Interest Personal", "page_visits": 950}, u[RETAIL]["offer"])
    check("retargeting click-through comes from the line naming the audience's RFPID",
          u[RETAIL]["retargeting_ctr"] == 0.001 and u[BUSINESS]["retargeting_ctr"] is None)
    check("each audience ran its own creative: the caveat says so",
          f["creatives_separate"] and ra.AUDIENCE_CREATIVE_CAVEAT in f["narrative"])
    check("page visits aren't split by audience: the footnote says so",
          f["footnote"] == ra.AUDIENCE_PAGES_FOOTNOTE)
    check("the headline states the rate multiple", f["headline"].startswith("Banking Intenders "
                                                                           "responded at 1.4x"), f["headline"])
    check("creatives that don't split exactly give no range (never a guess)",
          ra.assign_creatives_to_units([row("A", 10, 0.1), row("B", 7, 0.1)],
                                       [{"delivered": 9}, {"delivered": 8}]) is None)


def check_market_mix():
    print("\n6. Markets judged against their own audience")
    a = bank_attribution()
    basis = ra.market_audience_basis(a)
    check("two audiences 40% apart: the campaign average isn't a fair baseline",
          basis is not None and not basis["known"])
    found = ra.optimization_candidates(a, "moderate", ra.OPTIMIZATION_DIMENSIONS)
    market = [e for v in found.values() if isinstance(v, list) for e in v
              if isinstance(e, dict) and e.get("dimension") == "market"]
    check("with no proposal saying which audience ran where, no market is cut", not market, market)
    mapped = ra.with_market_audiences(a, {"washington_hagerstown": [ra.audience_terms(BUSINESS)]})
    basis = ra.market_audience_basis(mapped)
    check("a market one audience was aimed at is judged against that audience's rate",
          basis["baselines"]["WASHINGTON, DC (HAGRSTWN)"] == 0.0357)
    found = ra.optimization_candidates(mapped, "moderate", ra.OPTIMIZATION_DIMENSIONS)
    dc = [e for v in found.values() if isinstance(v, list) for e in v
          if isinstance(e, dict) and e.get("value") == "WASHINGTON, DC (HAGRSTWN)"]
    check("...and its candidate says so", dc and dc[0]["campaign_rate"] == 0.0357
          and dc[0]["baseline"] == "audience", dc)
    similar = bank_attribution()
    similar.by_audience = [AttributionRow(BUSINESS, 820_000, 33_000, 0.040),
                           AttributionRow(RETAIL, 550_000, 23_000, 0.042)]
    check("audiences within 20% of each other: the campaign average stays fair",
          ra.market_audience_basis(similar) is None)
    facts = ra.build_facts_payload(a, None, goals=["Grow deposits"])
    draft = {"threads": [{"head": "Optimize", "finding": "Norfolk led.", "action": (
        "Reduce investment in Washington, DC, which trails the campaign average by 3x, and "
        "reallocate toward Norfolk.")}]}
    check("a drafted cut against the campaign average is caught",
          len(app._market_mix_violations(draft, facts)) == 1)
    check("...a plain market fact isn't", not app._market_mix_violations(
        {"threads": [{"finding": "Norfolk delivered 400,000 impressions."}]}, facts))


def check_accruing():
    print("\n7. Still-accruing weeks and the export date")
    a = bank_attribution(exported_on=date(2026, 11, 4))
    f = ra.export_date_facts(a)
    check("weeks ending within 15 days of the export are still accruing",
          f["accruing_weeks"] == ["Oct 19–25", "Oct 26–Nov 1"], f["accruing_weeks"])
    check("the export is early: flight end + 15 days is after it", not f["accrual_complete"])
    check("the rep is warned, with the date to re-pull",
          "re-pull on or after Nov 16, 2026" in (ra.export_date_warning(f) or ""), ra.export_date_warning(f))
    trend = ra.trend_facts(a)
    check("an accruing week neither extends nor ends a rise: it runs to the last final week",
          trend["consecutive_rises"] == 4 and trend["rising_to_label"] == "Oct 12–18",
          (trend["consecutive_rises"], trend["rising_to_label"]))
    check("the trend facts mark accruing weeks",
          [w["still_accruing"] for w in trend["weeks"]][-2:] == [True, True])
    final = bank_attribution(exported_on=date(2026, 12, 20))
    check("a late enough export: nothing accruing, no warning",
          ra.export_date_facts(final)["accruing_weeks"] == []
          and ra.export_date_warning(ra.export_date_facts(final)) is None)
    path = TMP / "dated.xlsx"
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.append(["Delivered Impressions", "Attributed Impressions", "Attributed Rate"])
    wb.active.append([1000, 10, 0.01])
    wb.properties.created = datetime(2026, 11, 4, 15, 49)
    wb.save(path)
    check("the parser reads the export date from the workbook",
          ai.parse_attribution_export(str(path)).exported_on == date(2026, 11, 4))



def check_page():
    print("\n2/7. The page: the landing-page editor and the early-export warning")
    sys.path.insert(0, str(REPO / "tests"))
    import live_sports_fixtures as fx
    import test_attribution_reports_page as suite
    store = suite.install_store()
    at = suite.new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_attribution_upload_path"] = str(
        fx.attribution_workbook(TMP / "page_std.xlsx", "standard"))
    at.run()
    suite._confirm_advertiser(at)
    for b in at.button:
        if b.label == "No proposal -- build this report standalone":
            b.click().run()
            break
    check("the page renders", not at.exception, at.exception)
    check("a 'Campaign landing page(s)' editor is offered",
          any("Campaign landing page" in str(e.label) for e in at.expander),
          [e.label for e in at.expander])
    pages = at.session_state["attr_landing_pages"] if "attr_landing_pages" in at.session_state else []
    check("...pre-filled with the page drawing an outsized share of visits",
          [p_["pattern"] for p_ in pages] == ["https://www.example.org/loans/", "https://www.example.org/rates/"], pages)   # 18% and 10.5% of 3,800
    check("nothing is logged without Generate", not store.log_report_calls)

def main():
    check_spillover()
    check_geography_tile()
    check_urls()
    check_landing_pages()
    check_key_pages()
    check_audiences()
    check_market_mix()
    check_accruing()
    check_decks()
    check_page()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
