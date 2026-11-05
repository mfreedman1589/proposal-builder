"""The auto-dealer report path, from the auto group review (2026-11).

    python tests/test_auto_dealer_report.py

Synthetic exports and a synthetic Auto-Sales Analyst file only (CI);
tests/test_auto_dealer_real.py checks the real fixture locally. One section
per review item:

  1. The auto story is qualified traffic -> vehicles viewed -> vehicles sold;
     no confirmation/thank-you/checkout highlight, takeaway or panel.
  2. The one-sheet's sidebar is "From site visit to sale" (vehicle page
     visits = Used VDP + New VDP, never URLs containing "vehicle"), with
     values and labels that never overlap.
  3. The Watch List is never a highlight or takeaway; at most one What's
     Next idea touches it.
  4. Audience names: "NO DATA TARGETING" is "All <package> viewers";
     category prefixes go ("DEMO Spanish Speakers" -> "Spanish Speakers
     audience"); the rep's name still wins.
  5. The audience comparison carries households and frequency per line, and
     says so when frequency differs 2x or more.
  6. Group runs: the client store's own figures beside every group total; a
     store is named by its dealer, never a page title; "Client" on the
     report's own store; a group site's sales are group-wide inventory; the
     headline lead is a setting.
  7. Models normalized before ranking (F150 = F-150, F250 Super Duty =
     F-250); a vehicle's page and its /cta page are one vehicle (by VIN).
  8. No "shift creative toward used" recommendation.
  9. Differing impression totals: each slide footnotes the one it shows.
 10. Still accruing: every date within 15 days of the export date.

Follow-up (2026-11):
 A1. A Bilingual Spanish buy is ONE audience -- no by-audience breakdown,
     no audience comparison, no households/frequency.
 A2. Two or more tracked sites get "Where Visitors Went: Group Site vs.
     Store Site" and a one-sheet line, with the implied overlap.
 B1. Uploads are saved and cached by content, never by file name.
 B2. Files from two campaigns are flagged (advertiser, impressions).
 B3. No snake_case key in any generated slide.
 B4/B5. "- Direct" off client names; "Washington, DC" without "(Hagrstwn)".
 C2. "(group site)" only on a site that IS the group, not a named store.
"""
import re
import json
import os
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import subprocess  # noqa: E402

import openpyxl  # noqa: E402
from pptx import Presentation  # noqa: E402

import analyst_import as an  # noqa: E402
import app  # noqa: E402
import attribution_import as ai  # noqa: E402
import assembly  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402
from attribution_import import AttributionRow, DateSeriesPoint  # noqa: E402

failures = []
TMP = Path(tempfile.mkdtemp())
CLIENT = "Northgate Rivertown Ford"
GROUP_SITE, STORE_SITE = "northgateauto.com", "northgatefordrivertown.com"
PAGE_TITLE = "Used 2025 Honda Civic For Sale Rivertown MD"


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:400]}" if detail else ""))


def _block(viewed, sold, ltb, new, used, value, visits, sold_new, sold_used, mix):
    return {"vehicles_shopped": viewed, "vehicles_sold": sold, "look_to_book_pct": ltb,
            "look_to_book_pct_new": new, "look_to_book_pct_used": used,
            "vehicles_sold_new": sold_new, "vehicles_sold_used": sold_used,
            "higher_sell_through_condition": "USED", "est_revenue_sold": value,
            "est_pipeline_value": value * 4, "visits_total": visits, "traffic_mix": mix,
            "data_status": "ok"}


MIX = [{"category": "USED_VDP", "label": "Used VDP", "visits": 600},
       {"category": "NEW_VDP", "label": "New VDP", "visits": 150},
       {"category": "HOMEPAGE", "label": "Homepage", "visits": 100},
       {"category": "ONLINE_CONVERSIONS", "label": "Online Conversions", "visits": 5},
       {"category": "OTHER", "label": "Other", "visits": 145}]
SOLD_MODELS = [{"make": "FORD", "model": "F-150", "label": "FORD F-150", "count": 12},
               {"make": "TOYOTA", "model": "RAV4", "label": "TOYOTA RAV4", "count": 6},
               {"make": "FORD", "model": "F150", "label": "FORD F150", "count": 5},
               {"make": "FORD", "model": "F-250", "label": "FORD F-250", "count": 3},
               {"make": "FORD", "model": "F250 SUPER DUTY", "label": "FORD F250 SUPER DUTY", "count": 2}]
MISSED = [{"make": "CHEVROLET", "model": "CORVETTE", "label": "CHEVROLET CORVETTE", "visits": 30,
           "attributed_visits": 30, "est_value": 30000, "vin": "1G1YU3D65G5600001"},
          {"make": "TOYOTA", "model": "4RUNNER", "label": "TOYOTA 4RUNNER", "visits": 20,
           "attributed_visits": 20, "est_value": 25000, "vin": "JTEDU5JR1M5200002"},
          {"make": "FORD", "model": "F250 SUPER DUTY", "label": "FORD F250 SUPER DUTY", "visits": 18,
           "attributed_visits": 18, "est_value": 80000, "vin": "1FT8W2BT4TED00003"},
          {"make": "TOYOTA", "model": "4RUNNER CTA", "label": "TOYOTA 4RUNNER CTA", "visits": 14,
           "attributed_visits": 14, "est_value": 25000, "vin": "JTEDU5JR1M5200002"}]


def analyst_file():
    totals = _block(400, 100, 25.0, 15.0, 28.0, 3_400_000, 1000, 10, 90, MIX)
    totals.update({"sold_by_make_model": SOLD_MODELS, "top_sellers": SOLD_MODELS,
                   "shopped_by_make_model": [dict(m, count=m["count"] * 4) for m in SOLD_MODELS],
                   "missed_opportunities": {"count": 4, "vehicles": MISSED},
                   "unique_visitors": 500, "shopped_vehicle_status": {"available": 300, "sold": 100}})
    data = {"schema_version": 2,
            "meta": {"analysis_period": {"start": "2026-10-01", "end": "2026-10-31"},
                     "is_group": True, "site_count": 2, "inventory_scanned_at": "2026-11-05",
                     "group_name": "Auto Group (2 Sites)"},
            "sites": [{"site_id": "northgate-auto-group", "dealer_name": "Northgate Auto Group",
                       "domain": GROUP_SITE, "is_group_site": True, "data_status": "ok"},
                      {"site_id": "used-2025-honda-civic", "dealer_name": PAGE_TITLE,
                       "domain": STORE_SITE, "is_group_site": False, "data_status": "ok"}],
            "totals": totals,
            "by_site": {"northgate-auto-group": _block(300, 75, 25.0, 10.0, 26.0, 2_500_000, 600, 2,
                                                       73, MIX[:3]),
                        "used-2025-honda-civic": _block(100, 25, 25.0, 17.0, 40.0, 900_000, 400, 8,
                                                        17, MIX[:3])}}
    return an.parse_analyst_facts(json.dumps(data).encode())


URLS = {f"https://www.{STORE_SITE}/used-2024-Ford-F-150-XLT-1FTEW1EP0RFA00011": 40,
        f"https://www.{STORE_SITE}/new-2026-Ford-Bronco-Badlands-1FMEE7FP0TLA00012": 30,
        f"https://www.{GROUP_SITE}/used-2023-Toyota-RAV4-XLE-2T3P1RFV0PW000013": 25,
        f"https://www.{STORE_SITE}/": 60,
        f"https://www.{STORE_SITE}/thank-you": 4,
        f"https://www.{STORE_SITE}/checkout/confirmation": 2}
UNTARGETED, SEGMENT = "NO DATA TARGETING", "DEMO Spanish Speakers"


def attribution(exported_on=date(2026, 11, 5), segment=SEGMENT, pixels=None):
    weeks = [DateSeriesPoint(date(2026, 10, 12), 10_000, 700, 0.07),
             DateSeriesPoint(date(2026, 10, 19), 50_000, 3_600, 0.072),
             DateSeriesPoint(date(2026, 10, 26), 42_000, 3_800, 3_800 / 42_000)]
    return ai.AttributionExport(
        delivered_impressions=102_000, attributed_impressions=8_100, attributed_rate=8_100 / 102_000,
        attributed_unique_visitors=500,
        by_audience=[AttributionRow(UNTARGETED, 60_000, 5_400, 0.09),
                     AttributionRow(segment, 42_000, 2_700, 2_700 / 42_000)],
        by_pixel=list(pixels or []),
        by_url=dict(URLS), weekly_trend=weeks,
        by_market=[AttributionRow("BALTIMORE", 102_000, 8_100, 8_100 / 102_000)],
        by_zip=[AttributionRow("21201", 5_000, 1_000, 0.2), AttributionRow("21202", 6_000, 600, 0.1)],
        zip_heat={"21201": 1_000, "21202": 600},
        flight_start=date(2026, 10, 17), flight_end=date(2026, 10, 31), exported_on=exported_on)


HOMEOWNERS = "DEMO Homeowners"
RON = "PREM TV - RON"


def delivery(total=100_000, package="PREM TV - BILINGUAL SPANISH", segment="DEMO Spanish Speakers"):
    return ai.DeliveryExport(
        delivered_impressions=total, frequency=4.9, uniques=30_000, vcr=0.98,
        by_geo=[("Rivertown Zips", total)], zip_option_geos=True,
        delivery_map={"21201": 50_000, "21202": 50_000},
        packages=[(package, total)],
        segments=[("No Data Targeting", 58_000), (segment, 42_000)],
        campaign_lines=[
            {"name": f"WUSA - {CLIENT} - PREM OTT/CTV", "segment": "No Data Targeting",
             "impressions": 58_000, "households": 7_000, "frequency": 8.2},
            {"name": f"WUSA - {CLIENT} - {segment} - PREM OTT/CTV",
             "segment": segment, "impressions": 42_000, "households": 13_000,
             "frequency": 3.2}],
        flights=[{"start": date(2026, 10, 17), "end": date(2026, 10, 31), "impressions": total,
                  "uniques": 30_000, "booked": 0, "frequency": 4.9}])


def payload(analyst=None, lead=None, stores=None):
    a = ra.with_audience_names(attribution(), {}, package=ra.audience_package_phrase(delivery()))
    analyst = analyst or analyst_file()
    return a, ra.build_facts_payload(a, delivery(), goals=["Drive vehicle sales"], vertical="auto",
                                     client_name=CLIENT, analyst=analyst,
                                     analyst_client_stores=stores, store_lead=lead)


def template():
    """v0_19 (the site slide) built from the active v0_18 -- CI's stand-in
    carries v0_18's name; v0_19 is uploaded inactive until the merge."""
    src = REPO / "REPORT_MASTER_v0_18.pptx"
    if not src.exists():
        return None
    out = TMP / "REPORT_MASTER_v0_19.pptx"
    if not out.exists():
        subprocess.run([sys.executable, str(REPO / "build_v0_19.py"), str(src), str(out)],
                       check=True, capture_output=True, cwd=str(REPO))
    return out


_SNAKE_RE = re.compile(r"\b[a-z][a-z0-9]*(?:_[a-z0-9]+)+\b")


def deck_text(path):
    prs = Presentation(str(path))
    out = []
    for s in prs.slides:
        for sh in slide_map.iter_all_shapes(s.shapes):
            if sh.has_text_frame:
                out.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                out += [c.text for r in sh.table.rows for c in r.cells]
    return "\n".join(out)


BUILT = []


def scan(path):
    """B3: every deck this file builds is scanned for snake_case keys."""
    BUILT.append(path)
    return path


def drafted():
    """A draft that breaks every rule the review named."""
    return {"threads": [
        {"head": "Visitors Reached Confirmation Pages", "anchor": "goal",
         "finding": "4 confirmation-page visits -- Thankyou drew 1 visitors.", "meaning": "m1",
         "action": None, "kind": "confirmation", "protected": True},
        {"head": "Used Inventory Moving at Higher Rate", "anchor": "goal",
         "finding": "Used vehicles sold at 28.0% Look-to-Book vs. 15.0% for new.",
         "meaning": "Used inventory sees stronger movement.",
         "action": "Shift Streaming TV creative emphasis toward used vehicle inventory."},
        {"head": "Missed Opportunities Watch List", "anchor": "signal",
         "finding": "Vehicles like the 2016 Chevrolet Corvette (30 visits) are still listed.",
         "meaning": "High-interest inventory hasn't moved.",
         "action": "For each vehicle on the watch list, check the VDP for missing photos or a "
                   "'Call for Price' button."},
        {"head": "TV Sent Them Searching", "anchor": "signal",
         "finding": "About half of visits with referral data came from organic search.",
         "meaning": "The campaign sent viewers to search for the dealer.",
         "action": "Dynamic ads can feature the 2016 Chevrolet Corvette (30 visits) to the "
                   "households that viewed it.", "action_tier": 2},
        {"head": "Inventory Moving", "anchor": "signal",
         "finding": "100 of the vehicles our audience viewed have since sold.",
         "meaning": "Of those, 100 have since sold.", "action": None}]}


# ---------------------------------------------------------------------------

def check_auto_story():
    print("1. The auto story: traffic -> viewed -> sold; no online-conversion theme")
    a, facts = payload()
    check("an auto report never gets the confirmation story thread",
          not any(t.get("kind") == "confirmation"
                  for t in app._ensure_story_threads({"threads": []}, facts)["threads"]))
    non_auto = dict(facts, vertical="fitness", analyst=None,
                    confirmation_pages=[{"label": "Thank You", "visitors": 4}])
    check("...while any other report still does",
          any(t.get("kind") == "confirmation"
              for t in app._ensure_story_threads({"threads": []}, non_auto)["threads"]))
    draft = app._enforce_draft_rules(drafted(), facts, a)
    heads = [t["head"] for t in draft["threads"]]
    kinds = [t.get("kind") for t in draft["threads"]]
    check("the two auto story threads lead, in order", kinds[:2] == ["auto_traffic", "auto_sold"], kinds)
    traffic, sold = draft["threads"][0], draft["threads"][1]
    check("traffic: vehicle detail page visits are Used VDP + New VDP (750 of 1,000)",
          "(750 of 1,000) were vehicle detail pages" in traffic["finding"]
          and "500 visitors looked at 400 specific vehicles" in traffic["finding"], traffic["finding"])
    check("sold: viewed vehicles sold, est. value and Look-to-Book",
          sold["finding"].startswith("100 of those vehicles have since sold, an est. $3.4M in value "
                                     "(25.0% Look-to-Book)."), sold["finding"])
    check("no confirmation/thank-you/checkout thread survives",
          not any("confirmation" in h.lower() or "thank" in h.lower() for h in heads), heads)
    hi, ta, _wn = ra.distribute_threads(draft["threads"])
    text = " ".join(f"{h} {d}" for h, d in hi + ta).lower()
    check("...so none reaches a highlight or takeaway",
          "confirmation" not in text and "thank" not in text and "checkout" not in text, text)
    check("a drafted thread restating the sold figure gives way to Python's",
          "Inventory Moving" not in heads, heads)
    check("the rules are idempotent", app._enforce_draft_rules(json.loads(json.dumps(draft)), facts, a)
          ["threads"] == draft["threads"])


def check_one_sheet():
    print("\n2. One-sheet: 'From site visit to sale'")
    a, facts = payload()
    rows = ra.auto_sidebar_rows(facts["analyst"])
    check("sidebar rows: vehicle page visits, vehicles viewed, viewed vehicles sold, est. value sold",
          rows == [("750", "Vehicle page visits"), ("400", "Vehicles viewed"),
                   ("100", "Viewed vehicles sold"), ("$3.4M", "Est. value sold")], rows)
    if template() is None:
        print("  SKIP  REPORT_MASTER_v0_18.pptx not present")
        return
    draft = app._enforce_draft_rules(drafted(), facts, a)
    out = scan(TMP / "summary.pptx")
    ra.build_summary_slide(str(template()), str(out), attribution=a, delivery=delivery(),
                           client_name=CLIENT, threads=draft["threads"], accepted_optimizations=[],
                           vertical="auto", analyst_facts=facts["analyst"])
    slide = Presentation(str(out)).slides[0]
    shapes = {s.name: s for s in slide.shapes}
    text = {n: s.text_frame.text for n, s in shapes.items() if s.has_text_frame}
    check("headline and caption", text.get("SidebarHeadline") == ra.AUTO_SIDEBAR_HEADLINE
          and text.get("SidebarStatDetail") == ra.AUTO_SIDEBAR_CAPTION,
          (text.get("SidebarHeadline"), text.get("SidebarStatDetail")))
    check("the goal path panel is gone", "Goal path" not in " ".join(text.values()))
    pairs = [("SidebarStatValue", "SidebarStatLabel")] + [(f"SidebarSub{n}Value", f"SidebarSub{n}Label")
                                                         for n in range(1, 4)]
    overlaps = []
    for v, lab in pairs:
        value, label = shapes[v], shapes[lab]
        width = (assembly._widest_line_pt(value.text_frame, slide, 1.0) * assembly._EMU_PER_POINT
                 + value.text_frame.margin_left)
        if value.left + width > label.left or value.left + value.width > label.left:
            overlaps.append((text[v], text[lab]))
    check("no value runs into its label ($3.4M included)", not overlaps, overlaps)
    takeaways = " ".join(text.get(f"SummaryTakeaway{n}Detail", "") for n in range(1, 4)).lower()
    check("no confirmation theme in the one-sheet's takeaways",
          "confirmation" not in takeaways and "thank" not in takeaways, takeaways)


def check_watchlist():
    print("\n3. Watch List: never a highlight or takeaway; one idea at most")
    a, facts = payload()
    draft = app._enforce_draft_rules(drafted(), facts, a)
    hi, ta, wn = ra.distribute_threads(draft["threads"])
    slide_text = " ".join(f"{h} {d}" for h, d in hi + ta).lower()
    check("no watch-list highlight or takeaway",
          "watch list" not in slide_text and "missed opportunit" not in slide_text
          and "corvette" not in slide_text, slide_text)
    watch = [w for w in wn if app._mentions_watchlist(w, facts) or ra._is_listing(w)]
    check("exactly one What's Next item touches it, as an idea to consider",
          len(watch) == 1 and ra.IDEAS_HEADING in wn and wn.index(watch[0]) > wn.index(ra.IDEAS_HEADING),
          wn)
    check("no listing check (photos, 'Call for Price') survives as an action",
          not any("call for price" in w.lower() for w in wn), wn)
    texts = ra.analyst_watchlist_narrative(ra.analyst_cross_facts(an.with_store_names(
        analyst_file(), CLIENT), a))
    check("the Watch List slide's own text suggests, never directs",
          "should" not in texts.lower() and "quick check" in texts, texts)


def check_audience_names():
    print("\n4. Audience names")
    d = delivery()
    package = ra.audience_package_phrase(d)
    check("the package names the untargeted line", package == "Spanish-language", package)
    check("'NO DATA TARGETING' -> 'All Spanish-language viewers', never 'TARGETING'",
          ra.default_audience_name(UNTARGETED, package) == "All Spanish-language viewers")
    check("the prefix goes: 'DEMO Spanish Speakers' -> 'Spanish Speakers audience'",
          ra.default_audience_name(SEGMENT, package) == "Spanish Speakers audience")
    check("run of network reads 'All streaming viewers'",
          ra.default_audience_name(UNTARGETED, ra.audience_package_phrase(
              ai.DeliveryExport(packages=[("PREM TV - RON", 1)]))) == "All streaming viewers")
    a = ra.with_audience_names(attribution(), {}, package=package)
    check("audience_name uses both", [ra.audience_name(a, r.label) for r in a.by_audience]
          == ["All Spanish-language viewers", "Spanish Speakers audience"])
    plain = ai.AttributionExport(by_audience=[AttributionRow("FIN Banking Intenders", 1, 1, 1.0),
                                              AttributionRow("DEMO Career Owners", 1, 1, 1.0)])
    check("without an untargeted line, segments keep their bare names",
          [ra.audience_name(plain, r.label) for r in plain.by_audience]
          == ["Banking Intenders", "Career Owners"])
    named = ra.with_audience_names(a, {UNTARGETED: "Everyone watching Spanish TV"})
    check("a rep's name still wins", ra.audience_name(named, UNTARGETED) == "Everyone watching Spanish TV")
    ron = ra.with_targeting(attribution(segment=HOMEOWNERS), None,
                            delivery(package=RON, segment=HOMEOWNERS))
    facts = ra.build_facts_payload(ron, delivery(package=RON, segment=HOMEOWNERS), vertical="auto",
                                   client_name=CLIENT)
    check("the facts the model reads carry the new names (run of network, two real audiences)",
          [x["audience"] for x in facts["audiences_compared"]["audiences"]]
          == ["All streaming viewers", "Homeowners audience"], facts["audiences_compared"])


def check_delivery_lines():
    print("\n5. Households and frequency per line")
    path = TMP / "delivery.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IMPRESSIONS BY PRODUCT PACKAGE"
    ws.append(["Package Type", "Delivered Impressions"])
    ws.append([RON, 100_000])
    ws = wb.create_sheet("IMPRESSIONS BY DATA SEGMENT")
    ws.append(["Segment Name", "Delivered Impressions"])
    ws.append(["No Data Targeting", 58_000])
    ws.append([HOMEOWNERS, 42_000])
    ws = wb.create_sheet("DETAILS BY CAMPAIGN")
    ws.append(["Campaign Id", "Campaign Name", "Geo", "Campaign Start Date", "Campaign End Date",
               "Campaign Total Impressions", "Campaign Total Uniques", "Campaign Frequency"])
    ws.append(["TOTAL", "-", "-", "-", "-", 100_000, 20_000, "5.0"])
    ws.append(["1", "WUSA - X - PREM OTT/CTV - RFPID - 1", "Zip Option - X", "2026-10-17",
               "2026-10-31", 58_000, 7_000, "8.2"])
    ws.append(["1", "WUSA - X - DEMO Homeowners - PREM OTT/CTV - RFPID - 1", "Zip Option - X",
               "2026-10-17", "2026-10-31", 42_000, 13_000, "3.2"])
    wb.save(path)
    d = ai.parse_delivery_export(str(path))
    check("the package and segments are read", d.packages == [(RON, 100_000)]
          and [s for s, _n in d.segments] == ["No Data Targeting", HOMEOWNERS], d.packages)
    check("each campaign line has its segment, households and frequency (the untargeted one matched "
          "by impressions)",
          [(c["segment"], c["households"], c["frequency"]) for c in d.campaign_lines]
          == [("No Data Targeting", 7_000, 8.2), (HOMEOWNERS, 13_000, 3.2)],
          d.campaign_lines)
    a = ra.with_targeting(attribution(segment=HOMEOWNERS), None, d)
    facts = ra.audience_comparison_facts(a, None, CLIENT, delivery=d)
    units = {u["label"]: u for u in facts["units"]}
    check("households and frequency reach each audience",
          units["All streaming viewers"]["households"] == 7_000
          and units["Homeowners audience"]["frequency"] == 3.2, units)
    check("frequency 2x apart: the narrative says it affects per-impression rates and credits reach",
          "affects per-impression rates" in facts["narrative"]
          and "The Homeowners audience delivered the broader reach: 13,000 households" in
          facts["narrative"], facts["narrative"])
    check("...and the headline says the rate rode on frequency",
          "at 2.6x the frequency" in facts["headline"], facts["headline"])
    close = ai.DeliveryExport(campaign_lines=[dict(c, frequency=f) for c, f in
                                              zip(d.campaign_lines, (4.0, 3.2))])
    check("under 2x: no frequency sentence",
          "per-impression" not in ra.audience_comparison_facts(a, None, CLIENT, delivery=close)["narrative"])
    split = ai.DeliveryExport(campaign_lines=d.campaign_lines + [dict(d.campaign_lines[1])])
    check("an audience split across several lines gets no households (they'd overlap)",
          ra.audience_comparison_facts(a, None, CLIENT, delivery=split)["units"][1]["households"] is None)
    if template() is not None:
        out = scan(TMP / "cmp.pptx")
        ra.build_report_deck(str(template()), a, d, str(out), client_name=CLIENT,
                             goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                             vertical="auto")
        prs = Presentation(str(out))
        slide = next(s for s in prs.slides if slide_map.notes_key(s) == "report:audience_comparison")
        table = next(s for s in slide.shapes if s.name == "AudienceComparisonTable").table
        labels = [r.cells[0].text for r in table.rows]
        check("the slide's table carries Households reached and Frequency",
              "Households reached" in labels and "Frequency" in labels, labels)


def check_group():
    print("\n6. Group runs: per-store figures, store names, the Client tag")
    raw = analyst_file()
    check("a page title is never a store name",
          an.looks_like_page_title(PAGE_TITLE) and not an.looks_like_page_title("Northgate Auto Group"))
    named = an.with_store_names(raw, CLIENT)
    check("...the client's store takes the client's name",
          [s["dealer_name"] for s in named["sites"]] == ["Northgate Auto Group", CLIENT],
          named["sites"])
    check("...and its make doesn't become a franchise (no Honda from the page title)",
          an.franchise_makes(named, CLIENT) == ["FORD"], an.franchise_makes(named, CLIENT))
    check("the client store is the store, never the group site that shares its first word",
          ra.guess_client_stores(raw, CLIENT) == [STORE_SITE], ra.guess_client_stores(raw, CLIENT))
    a, facts = payload()
    stores = facts["analyst"]["client_stores"]
    check("the payload carries the client store's own figures",
          [(s["store"], s["vehicles_sold_since"]) for s in stores] == [(CLIENT, 25)], stores)
    check("...the group-site note", facts["analyst"]["group_site_note"].startswith(
        "Northgate Auto Group lists inventory from stores across the group"), facts["analyst"])
    draft = app._enforce_draft_rules(drafted(), facts, a)
    sold = draft["threads"][1]["finding"]
    check("combined lead: the group total, then the store",
          sold.index("100 of those") < sold.index(f"At {CLIENT}, 25 have sold, with used vehicles at "
                                                  f"40.0%."), sold)
    check("the Executive Summary names the store",
          draft["analyst_inventory_narrative"].startswith(f"At {CLIENT}, 25 of the 100 vehicles"),
          draft["analyst_inventory_narrative"])
    other = {"threads": [{"head": "Inventory Story", "anchor": "signal",
                          "finding": "Across the group, 100 viewed vehicles found buyers.",
                          "meaning": "m", "kind": "custom"}]}
    out = app._enforce_draft_rules(other, facts, a)
    custom = next(t for t in out["threads"] if t["head"] == "Inventory Story")
    check("any other thread citing the group total gains the store's figures",
          f"At {CLIENT}, 25 have sold" in custom["finding"], custom["finding"])
    noted = app._enforce_draft_rules(
        {"threads": [], "analyst_group_narrative": (
            f"{CLIENT} and the group site sold at similar rates. Northgate Auto Group lists "
            f"inventory from stores across the group, so its sales reflect group-wide inventory.")},
        facts, a)
    check("a drafted group narrative loses the group-site note (the footnote says it once)",
          noted["analyst_group_narrative"] == f"{CLIENT} and the group site sold at similar rates.",
          noted["analyst_group_narrative"])
    _a, store_first = payload(lead="store_first")
    sold = ra.auto_story_threads(store_first["analyst"])[1]["finding"]
    check("store_first lead: the store, then every site",
          sold.startswith(f"At {CLIENT}, 25 of the 100 vehicles our audience viewed have since sold")
          and "Across every site, 100 have since sold" in sold, sold)
    check("store_first one-sheet: the store's own vehicle page visits, viewed, sold, value",
          ra.auto_sidebar_rows(store_first["analyst"]) == [
              ("750", "Vehicle page visits"), ("100", "Vehicles viewed"),
              ("25", "Viewed vehicles sold"), ("$900K", "Est. value sold")],
          ra.auto_sidebar_rows(store_first["analyst"]))
    check("the lead is a setting with a default", app.auto_report_store_lead() in ra.STORE_LEADS
          and ra.DEFAULT_STORE_LEAD == "combined")
    if template() is None:
        return
    out_path = scan(TMP / "group.pptx")
    ra.build_report_deck(str(template()), a, delivery(), str(out_path), client_name=CLIENT,
                         goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                         vertical="auto", analyst=raw)
    prs = Presentation(str(out_path))
    slide = next(s for s in prs.slides if slide_map.notes_key(s) == "report:analyst_group")
    rows = [[c.text for c in r.cells] for r in next(
        s for s in slide.shapes if s.name == "AnalystGroupTable").table.rows][1:]
    check("scoreboard: dealer names, the group site labelled, 'Client' on the report's own store",
          sorted((r[0], r[-1]) for r in rows) == sorted([("Northgate Auto Group (group site)", ""),
                                                          (CLIENT, "Client")]), rows)
    text = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    check("the group slide says a group site's sales are group-wide inventory",
          "group-wide inventory" in text and PAGE_TITLE not in text, text)


def check_models():
    print("\n7. Models normalized; vehicles deduped by VIN")
    check("F150 / F 250Sd / F250 SUPER DUTY normalize",
          [an.normalize_model(m) for m in ("F150", "F 250Sd", "F250 SUPER DUTY", "F-150 LIGHTNING")]
          == ["F-150", "F-250", "F-250", "F-150 LIGHTNING"])
    rows = an.top_model_rows(analyst_file())
    check("ranked after merging: Ford F-150 17 first, Ford F-250 5 merged",
          [(r["label"], r["count"]) for r in rows] == [("Ford F-150", 17), ("Toyota RAV4", 6),
                                                       ("Ford F-250", 5)], rows)
    a, facts = payload()
    check("the payload's top models are merged too",
          facts["analyst"]["top_models_sold"][0] == {"model": "Ford F-150", "count": 17},
          facts["analyst"]["top_models_sold"][:2])
    missed = facts["analyst"]["missed_opportunities"]
    check("the watch list has one row per VIN, at the larger visit count, no 'Cta'",
          [m["vehicle"].split(" ", 1)[-1] if m["vehicle"][:2] == "20" else m["vehicle"] for m in missed]
          == ["Chevrolet Corvette", "Toyota 4Runner", "Ford F-250"]
          and [m["visits"] for m in missed] == [30, 20, 18], missed)


def check_no_used_shift():
    print("\n8. No 'shift creative toward used'")
    a, facts = payload()
    draft = app._enforce_draft_rules(drafted(), facts, a)
    text = json.dumps(draft).lower()
    check("no action or sentence shifts creative toward used or new",
          "toward used" not in text and "toward new" not in text, text[:300])
    check("the used-vs-new thread is gone",
          not any("used inventory" in t["head"].lower() for t in draft["threads"]))
    prompt = app.build_attr_draft_prompt(facts)
    check("the prompt no longer asks for it", "Its action shifts the Streaming TV message" not in prompt
          and "normal dealer behaviour" in prompt)


def check_basis_notes():
    print("\n9. Impression totals that differ are footnoted")
    a, d = attribution(), delivery()
    check("delivery-based note", ra.impression_basis_note(a, d, "delivery")
          == "Impressions: 100,000 delivered, per the delivery report. Attributed rates use the "
             "attribution report's own impression count, which differs.")
    check("attribution-based note", ra.impression_basis_note(a, d, "attribution")
          == "Impressions on this slide are the attribution report's own count, which differs from "
             "the 100,000 delivered in the delivery report.")
    check("...and never prints the attribution file's own total (Correction 3)",
          "102,000" not in ra.impression_basis_note(a, d, "delivery")
          + ra.impression_basis_note(a, d, "attribution"))
    check("no note when they agree or there's no delivery file",
          ra.impression_basis_note(a, delivery(102_000), "delivery") is None
          and ra.impression_basis_note(a, None, "delivery") is None)
    if template() is None:
        return
    out = scan(TMP / "basis.pptx")
    ra.build_report_deck(str(template()), a, d, str(out), client_name=CLIENT,
                         goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                         vertical="auto")
    prs = Presentation(str(out))
    notes = {slide_map.notes_key(s): next((sh.text_frame.text for sh in s.shapes
                                           if sh.name == "ImpressionBasisNote"), None) for s in prs.slides}
    check("each slide footnotes the total it shows",
          all(notes.get(k, "").startswith("Impressions: 100,000") for k in
              ("report:highlights", "report:delivery_recap"))
          and all(notes.get(k, "").startswith("Impressions on this slide are the attribution") for k in
                  ("report:zip_analysis",)),
          notes)
    slide = next(s for s in prs.slides if slide_map.notes_key(s) == "report:highlights")
    note = next(sh for sh in slide.shapes if sh.name == "ImpressionBasisNote")
    logo = next(sh for sh in slide.shapes if sh.name == "TegnaLogo")
    check("...on the logo row, clear of the logo", note.left > logo.left + logo.width
          and note.top + note.height <= prs.slide_height, (note.left, note.top))


def check_accrual():
    print("\n10. Still accruing: every date within 15 days")
    facts = ra.export_date_facts(attribution(date(2026, 11, 5)), delivery())
    check("Oct 21-31 are accruing on a Nov 5 export",
          (facts["accruing_from"], facts["accruing_to"]) == (date(2026, 10, 21), date(2026, 10, 31)),
          facts)
    check("the recap says so, dates not weeks",
          ra.export_date_note(facts) == "Data as of Nov 5, 2026. Visits from Oct 21–31 are still "
          "arriving (attributed visits keep coming in for 15 days after an ad), so late-October "
          "figures will rise.", ra.export_date_note(facts))
    late = ra.export_date_facts(attribution(date(2026, 11, 20)), delivery())
    check("an export 20 days after the flight: nothing accruing", late["accruing_from"] is None
          and ra.export_date_note(late) == "Data as of Nov 20, 2026.", late)
    early = ra.export_date_facts(attribution(date(2026, 10, 25)), delivery())
    check("an export mid-flight: the whole period so far",
          early["accruing_from"] == date(2026, 10, 17), early)


def check_spanish_merge():
    print("\nA1. A Bilingual Spanish buy is one audience")
    d = delivery()
    a = ra.with_targeting(attribution(), None, d)
    check("the two lines merge into one audience, impressions summed",
          [(r.label, r.delivered_impressions, r.attributed_impressions) for r in a.by_audience]
          == [(ra.SPANISH_AUDIENCE_NAME, 102_000, 8_100)], a.by_audience)
    check("...named 'Spanish speakers, reached through Spanish-language streaming'",
          ra.audience_name(a, a.by_audience[0].label)
          == "Spanish speakers, reached through Spanish-language streaming")
    check("no audience comparison, no households or frequency",
          not ra.audience_comparison_applies(a)
          and ra.audience_comparison_facts(a, None, CLIENT, delivery=d) is None)
    check("idempotent", ra.with_targeting(a, None, d).by_audience == a.by_audience)
    no_tab = ai.DeliveryExport(delivered_impressions=100_000)
    check("without a package tab: an untargeted line beside a Spanish segment still merges",
          len(ra.merge_spanish_audience(attribution(), no_tab).by_audience) == 1)
    check("run of network with a real segment stays two audiences",
          len(ra.with_targeting(attribution(segment=HOMEOWNERS), None,
                                delivery(package=RON, segment=HOMEOWNERS)).by_audience) == 2)
    facts = ra.build_facts_payload(a, d, vertical="auto", client_name=CLIENT)
    check("the model is told it's one audience, with nothing to compare",
          facts["audiences_compared"] is None and facts["audience_merged"]["audience"]
          == ra.SPANISH_AUDIENCE_NAME, facts.get("audience_merged"))
    if template() is None:
        return
    out = scan(TMP / "spanish.pptx")
    ra.build_report_deck(str(template()), attribution(), d, str(out), client_name=CLIENT,
                         goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                         vertical="auto")
    keys = [slide_map.notes_key(s) for s in Presentation(str(out)).slides]
    check("no attribution-by-audience and no audience comparison slide",
          "report:attribution_breakdown" not in keys and "report:audience_comparison" not in keys, keys)
    check("the recap names the one audience", ra.SPANISH_AUDIENCE_NAME in deck_text(out))


PIXELS = [{"pixel": "TEGNA - WUSA - Northgate Auto Group - Direct Pixels", "client": CLIENT,
           "attributed": 5_000, "visitors": 330},
          {"pixel": "TEGNA - WUSA - Northgate Rivertown Ford Pixels", "client": CLIENT,
           "attributed": 3_100, "visitors": 300}]


def check_sites():
    print("\nA2. Group site vs. store site")
    a = attribution(pixels=PIXELS)
    f = ra.site_facts(a, analyst_file(), CLIENT)
    sites = {s["domain"]: s for s in f["sites"]}
    check("each pixel lands on its own site's domain, page visits from that domain's pages",
          sites[GROUP_SITE]["visitors"] == 330 and sites[STORE_SITE]["visitors"] == 300
          and sites[GROUP_SITE]["page_visits"] == 25
          and sites[STORE_SITE]["page_visits"] == 40 + 30 + 60 + 4 + 2, sites)
    check("the group site is labelled by its domain; the store by its name",
          sites[GROUP_SITE]["label"] == f"{GROUP_SITE} (group site)"
          and sites[STORE_SITE]["label"] == f"{CLIENT} site", sites)
    check("overlap = per-site visitors minus the unique total (330 + 300 - 500 = 130)",
          f["overlap"] == 130 and f["overlap_rounded"] == 130, f["overlap"])
    check("headline and one-sheet line",
          f["headline"] == f"330 visitors reached {GROUP_SITE} and 300 reached the {CLIENT} site; "
                           f"about 130 visited both."
          and f["one_sheet_line"] == f"330 visitors reached {GROUP_SITE} and 300 the {CLIENT} site; "
                                     f"about 130 visited both.", (f["headline"], f["one_sheet_line"]))
    check("the narrative states the overlap as a share ('roughly 1 in 4')",
          "Roughly 1 in 4 visitors looked at both." in f["narrative"], f["narrative"])
    check("one pixel: no site comparison", ra.site_facts(attribution(pixels=PIXELS[:1]), None, CLIENT) is None)
    if template() is None:
        return
    out = scan(TMP / "sites.pptx")
    ra.build_report_deck(str(template()), a, delivery(), str(out), client_name=CLIENT,
                         goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                         vertical="auto", analyst=analyst_file())
    prs = Presentation(str(out))
    slide = next((s for s in prs.slides if slide_map.notes_key(s) == "report:site_comparison"), None)
    check("the site slide is in the deck", slide is not None)
    if slide is not None:
        rows = [[c.text for c in r.cells] for r in next(
            s for s in slide.shapes if s.name == "SiteTable").table.rows]
        check("table: each site, then the deduplicated total",
              rows[1:] == [[f"{GROUP_SITE} (group site)", "330", "25", "5,000"],
                           [f"{CLIENT} site", "300", "136", "3,100"],
                           ["Total (unique)", "500", "161", "8,100"]], rows)
        text = "\n".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        check("title, footnote and no unfilled token on the slide",
              "Where Visitors Went: Group Site vs. Store Site" in text
              and "so about 130 appear on both sites." in text and "{{" not in text, text)
    out = scan(TMP / "sites_summary.pptx")
    ra.build_summary_slide(str(template()), str(out), attribution=a, delivery=delivery(),
                           client_name=CLIENT, threads=[], accepted_optimizations=[],
                           vertical="auto", analyst_facts=None)
    check("the one-sheet carries the site line", "about 130 visited both." in deck_text(out))
    one = scan(TMP / "one_site.pptx")
    ra.build_report_deck(str(template()), attribution(), delivery(), str(one), client_name=CLIENT,
                         goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                         vertical="auto")
    check("one site: the slide is dropped",
          "report:site_comparison" not in [slide_map.notes_key(s) for s in Presentation(str(one)).slides])


class _Upload:
    def __init__(self, name, data):
        self.name, self._data = name, data

    def getvalue(self):
        return self._data


def check_uploads_by_content():
    print("\nB1. Uploads are keyed by content, never by name")
    name = "Premion_Website_Attribution_and_Reach_Extension_15.xlsx"
    first, second = _Upload(name, b"client one"), _Upload(name, b"client two")
    p1, s1 = app.save_upload(first)
    p2, s2 = app.save_upload(second)
    check("two same-named uploads land at different paths, with different signatures",
          p1 != p2 and s1 != s2 and p1.read_bytes() == b"client one" and p2.read_bytes() == b"client two")
    check("the same file twice is the same signature (no needless re-parse)",
          app.save_upload(_Upload(name, b"client one"))[1] == s1)
    real = app.attribution_import.parse_attribution_export
    app.attribution_import.parse_attribution_export = lambda path, slot=None: Path(path).read_bytes()
    try:
        shared = TMP / "same_name.xlsx"
        shared.write_bytes(b"client one")
        a1 = app._cached_parse_attribution(str(shared))
        shared.write_bytes(b"client two")
        a2 = app._cached_parse_attribution(str(shared))
    finally:
        app.attribution_import.parse_attribution_export = real
    check("a new file at the same path never gets the old file's cached parse (NWFCU, 2026-11)",
          (a1, a2) == (b"client one", b"client two"), (a1, a2))
    source = (REPO / "app.py").read_text(encoding="utf-8")
    check("every upload slot compares a content signature, never the bare name",
          not re.search(r'"attr_\w+_loaded"\) != \w+_upload\.name', source)
          and source.count("save_upload(") >= 6, source.count("save_upload("))


def check_cross_file():
    print("\nB2. Files from two campaigns are flagged")
    a, d = attribution(), delivery()
    check("matching files: nothing", ra.cross_file_mismatches(a, d, CLIENT) == [])
    other = ai.DeliveryExport(delivered_impressions=601_561, campaign_lines=[
        {"name": "WUSA - Riverbend Credit Union - Direct - PREM OTT/CTV - RFPID - 1"}])
    a.client_name = CLIENT
    msgs = ra.cross_file_mismatches(a, other, "Riverbend Credit Union")
    check("another client's delivery file: advertiser and impressions both named",
          any(f'"{CLIENT}"' in m and '"Riverbend Credit Union"' in m for m in msgs)
          and any("102,000" in m and "601,561" in m for m in msgs), msgs)
    check("the report's own client is checked too",
          any("this report is for" in m for m in msgs), msgs)
    close = ai.DeliveryExport(delivered_impressions=95_000)
    check("impressions within 10%: no flag", ra.cross_file_mismatches(a, close, CLIENT) == [])
    check("a shared generic word doesn't make two clients match",
          not ra.names_match("Northwest Federal Credit Union", "Riverbend Credit Union")
          and ra.names_match("Northgate Automotive - Direct", "Northgate Rivertown Ford"))
    source = (REPO / "app.py").read_text(encoding="utf-8")
    check("Generate is blocked until the rep confirms",
          "cross_file_mismatches(" in source and 'disabled=not _files_confirmed' in source)


def check_text_hygiene():
    print("\nB3-B5, C2. Client text hygiene and group sites")
    draft = {"threads": [{"head": "Add OTT Retargeting", "anchor": "signal", "finding": None,
                          "meaning": None, "action_tier": 2,
                          "action": "Add OTT Retargeting (streaming_retargeting_display / "
                                    "streaming_retargeting_preroll) to build frequency."}]}
    a, facts = payload()
    out = app._enforce_draft_rules(draft, facts, a)
    text = json.dumps([t.get("action") for t in out["threads"]])
    check("product keys never survive a draft", not _SNAKE_RE.search(text) and "OTT Retargeting" in text,
          text)
    check("'- Direct' comes off a client name",
          ra.clean_client_name("Riverbend Credit Union - Direct") == "Riverbend Credit Union"
          and ra.clean_client_name("Direct Mail Co") == "Direct Mail Co")
    check("market labels drop Nielsen's parenthetical",
          ra._market_display("WASHINGTON, DC (HAGRSTWN)") == "Washington, DC")
    store = {"dealer_name": "Ted Example Ford of Chantilly", "domain": "tedexamplechantilly.com",
             "is_group_site": True}
    group = {"dealer_name": "Northgate Auto Group", "domain": GROUP_SITE, "is_group_site": True}
    check("a flagged site that names a make is a store, not a group site",
          not an.is_group_site(store) and an.is_group_site(group)
          and an.group_site_label(store, "Ted Example Ford of Chantilly") == "Ted Example Ford of Chantilly")
    if template() is not None:
        titled = scan(TMP / "direct.pptx")
        a2 = attribution()
        a2.client_name = f"{CLIENT} - Direct"
        ra.build_report_deck(str(template()), a2, delivery(), str(titled), client_name=None,
                             goals_bullets=["Drive vehicle sales"], whats_next_bullets=["Keep going"],
                             vertical="auto")
        check("the title shows the client without '- Direct'",
              "- Direct" not in deck_text(titled) and CLIENT in deck_text(titled))
    leaks = {str(p): sorted(set(_SNAKE_RE.findall(deck_text(p)))) for p in BUILT if Path(p).exists()}
    check(f"no snake_case token on any slide of the {len(leaks)} decks this file built",
          leaks and not any(leaks.values()), {k: v for k, v in leaks.items() if v})
    check("...and no '(Hagrstwn)'-style market label",
          not any("Hagrstwn" in deck_text(p) for p in BUILT if Path(p).exists()))


def main():
    check_spanish_merge()
    check_sites()
    check_uploads_by_content()
    check_cross_file()
    check_auto_story()
    check_one_sheet()
    check_watchlist()
    check_audience_names()
    check_delivery_lines()
    check_group()
    check_models()
    check_no_used_shift()
    check_basis_notes()
    check_accrual()
    check_text_hygiene()
    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
