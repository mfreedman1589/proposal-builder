"""Auto-Sales Analyst facts JSON -> the report builder.

Covers analyst_import.py (parse, period gate, visitor cross-check, the slide
and payload helpers), report:url_report filled from it, and the facts
payload + facts-only checker accepting the Analyst's own numbers.

Real fixtures (gitignored client data; the real-file checks SKIP without
them): Ted Britt's August 2026 website attribution export and the Analyst's
facts export for the same group and month. Expected values are taken from
the Analyst file's own totals block and the attribution export's own
advertiser tab -- never from the helpers under test.

    python tests/test_analyst_import.py
"""
import json
import re
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import market_lookup  # noqa: E402
import package_check  # noqa: E402
from pptx import Presentation  # noqa: E402

import analyst_import as an  # noqa: E402
import app  # noqa: E402
import attribution_import as ai  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402

TEMPLATE = REPO / "REPORT_MASTER_v0_14.pptx"
ATTRIBUTION_TB930 = REPO / "Premion Website Attribution and Reach Extension TB930.xlsx"
ANALYST_TB_AUG = REPO / "auto-group-5-sites-12-07-pm-et_2026-08-01_2026-08-31_facts.json"
OUT_DIR = REPO / "tests" / "_manual_output"

market_lookup.install()

failures = []
skipped = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def skip(label):
    print(f"  SKIP  {label}")
    skipped.append(label)


def _synthetic(**overrides):
    data = {
        "schema_version": 1,
        "meta": {"group_name": "Test Group", "is_group": True, "site_count": 2,
                 "analysis_period": {"start": "2026-08-01", "end": "2026-08-31",
                                     "source": "rep_entered_report_month"},
                 "inventory_scanned_at": "2026-09-30", "lookback_days": 30},
        "sites": [{"site_id": "a", "dealer_name": "Alpha Ford", "domain": "alpha.com"},
                  {"site_id": "b", "dealer_name": "Beta Chevy", "domain": "beta.com"}],
        "totals": {
            "vehicles_shopped": 200, "vehicles_sold": 80, "look_to_book_pct": 40.0,
            "est_revenue_sold": 3_200_000, "est_pipeline_value": 8_000_000,
            "visits_total": 1000, "unique_visitors": 400,
            "traffic_mix": [
                {"category": "NEW_VDP", "label": "New VDP", "visits": 300},
                {"category": "OTHER", "label": "Other", "visits": 250},
                {"category": "USED_VDP", "label": "Used VDP", "visits": 200},
                {"category": "HOMEPAGE", "label": "Homepage", "visits": 100},
                {"category": "SERVICE", "label": "Service", "visits": 80},
                {"category": "NEW_CAR_SEARCH", "label": "New Car Search", "visits": 70},
            ],
            "sold_by_make_model": [
                {"make": "FORD", "model": "MUSTANG MACH-E", "label": "FORD MUSTANG MACH-E", "count": 7},
                {"make": "GMC", "model": "SIERRA 1500", "label": "GMC SIERRA 1500", "count": 12},
                {"make": "FORD", "model": "F-150", "label": "FORD F-150", "count": 20},
            ],
            "missed_opportunities": {"count": 1, "vehicles": [
                {"make": "FORD", "model": "BRONCO", "label": "FORD BRONCO", "visits": 44,
                 "vin": "1FMEE9BP3SLB32625"}]},
        },
        "by_site": {"b": {"visits_total": 5, "traffic_mix": []}},
    }
    data.update(overrides)
    return json.dumps(data).encode()


def _raises(source):
    try:
        an.parse_analyst_facts(source)
    except an.AnalystParseError as exc:
        return str(exc)
    return None


def check_parser():
    print("\nParser -- shape, errors a rep can act on")
    parsed = an.parse_analyst_facts(_synthetic())
    check("a valid export parses, dates kept as ISO strings",
          parsed["period_start"] == "2026-08-01" and parsed["period_end"] == "2026-08-31",
          (parsed["period_start"], parsed["period_end"]))
    check("dealer names carried from the sites list",
          [s["dealer_name"] for s in parsed["sites"]] == ["Alpha Ford", "Beta Chevy"])
    check("the parsed dict is JSON-safe (it's logged with the report)",
          json.loads(json.dumps(parsed)) == parsed)

    msg = _raises(b"{not json")
    check("unreadable JSON -> plain message pointing at the Analyst's facts download",
          msg is not None and "facts" in msg and "Traceback" not in msg, msg)
    msg = _raises(json.dumps({"hello": 1}).encode())
    check("some other JSON file -> named as not an Analyst facts export",
          msg is not None and "isn't an Auto-Sales Analyst" in msg, msg)
    msg = _raises(_synthetic(schema_version=2))
    check("a newer schema version is refused, never half-read", msg is not None and "2" in msg, msg)
    bad_period = json.loads(_synthetic())
    bad_period["meta"]["analysis_period"]["start"] = "August"
    msg = _raises(json.dumps(bad_period).encode())
    check("an unreadable analysis period is refused (the period gate depends on it)",
          msg is not None and "August" in msg, msg)
    empty = json.loads(_synthetic())
    empty["totals"] = {"visits_total": 0}
    msg = _raises(json.dumps(empty).encode())
    check("a file with no traffic or vehicle data is refused", msg is not None, msg)
    bom = b"\xef\xbb\xbf" + _synthetic()
    check("a UTF-8 BOM (Windows save) still parses", _raises(bom) is None, _raises(bom))


def check_gates():
    print("\nPeriod overlap and visitor cross-check")
    parsed = an.parse_analyst_facts(_synthetic())
    check("overlapping window (Jul 27 - Aug 31) -> used",
          an.periods_overlap(parsed, date(2026, 7, 27), date(2026, 8, 31)))
    check("one shared day at the edge counts as overlap",
          an.periods_overlap(parsed, date(2026, 8, 31), date(2026, 9, 30)))
    check("June report vs August Analyst -> not used",
          not an.periods_overlap(parsed, date(2026, 6, 1), date(2026, 6, 30)))
    check("ISO-string bounds work the same as dates (manual period entry stores strings)",
          not an.periods_overlap(parsed, "2026-06-01", "2026-06-30")
          and an.periods_overlap(parsed, "2026-08-10", "2026-08-12"))
    check("a missing report bound never blocks the file",
          an.periods_overlap(parsed, None, date(2026, 6, 30)))
    check("matching visitor counts -> no mismatch", an.visitor_count_mismatch(parsed, 400) is None)
    check("differing counts -> both named, Analyst first",
          an.visitor_count_mismatch(parsed, 390) == (400, 390), an.visitor_count_mismatch(parsed, 390))
    check("no export count -> nothing to compare", an.visitor_count_mismatch(parsed, 0) is None)

    check("a normal file (80 sold) is usable", an.unusable_reason(parsed) is None)
    zero = json.loads(_synthetic())
    zero["totals"]["vehicles_sold"] = 0
    reason = an.unusable_reason(an.parse_analyst_facts(json.dumps(zero).encode()))
    check("0 sold with vehicles shopped -> set aside, per the Sales Assist's own rule",
          reason is not None and "0 vehicles sold" in reason, reason)
    check("scan 30 days after period end (the recommended window) -> no advisory",
          an.scan_delay_days(parsed) == 30 and an.lookback_advisory(parsed) is None,
          an.scan_delay_days(parsed))
    late = json.loads(_synthetic())
    late["meta"]["inventory_scanned_at"] = "2026-11-24"
    late_parsed = an.parse_analyst_facts(json.dumps(late).encode())
    advisory = an.lookback_advisory(late_parsed)
    check("scanned 85 days after period end -> advisory naming the days (scan date, not "
          "lookback_days, decides)", advisory is not None and "85 days" in advisory, advisory)
    facts = {"analyst": an.payload_facts(late_parsed)}
    items = app.attr_actionable_review_items(facts)
    check("...and the same fact reaches Review before sending",
          any("85 days" in i for i in items), items)
    no_scan = json.loads(_synthetic())
    no_scan["meta"].pop("inventory_scanned_at")
    no_scan["meta"]["lookback_days"] = 90
    check("no scan date -> falls back to lookback_days",
          an.scan_delay_days(an.parse_analyst_facts(json.dumps(no_scan).encode())) == 90)

    check("franchise makes come from the client and dealer names",
          an.franchise_makes(parsed, "Alpha Ford Lincoln") == ["CHEVROLET", "FORD", "LINCOLN"],
          an.franchise_makes(parsed, "Alpha Ford Lincoln"))
    stafford = {"sites": [{"dealer_name": "Stafford Honda"}, {"dealer_name": "Dominion Motors"}]}
    check("short makes need a whole word (Stafford isn't Ford, Dominion isn't Mini)",
          an.franchise_makes(stafford) == ["HONDA"], an.franchise_makes(stafford))
    check("run-together domain names still match a long make (Tedbrittchevrolet)",
          an.franchise_makes({"sites": [{"dealer_name": "Tedbrittchevrolet"}]}) == ["CHEVROLET"])
    unverified = json.loads(_synthetic())
    unverified["totals"]["shopped_vehicle_status"] = {"available": 100, "sold": 80, "unverified": 20}
    payload = an.payload_facts(an.parse_analyst_facts(json.dumps(unverified).encode()))
    check("'Inventory Unavailable' vehicles ride separately, never inside the sold count",
          payload["vehicles_status_unconfirmed"] == 20 and payload["vehicles_sold_since"] == 80,
          payload)


def check_slide_helpers():
    print("\nSlide helpers -- traffic mix and top models")
    parsed = an.parse_analyst_facts(_synthetic())
    rows = an.traffic_mix_rows(parsed)
    labels = [r["label"] for r in rows]
    check("named categories largest-first, capped, remainder row last",
          labels == ["New VDP", "Used VDP", "Homepage", "Service", "All other pages"], labels)
    check("the Analyst's own 'Other' never appears beside 'All other pages'",
          "Other" not in labels, labels)
    check("the table accounts for every visit (sums to visits_total = 1000)",
          sum(r["visits"] for r in rows) == 1000, sum(r["visits"] for r in rows))
    check("visit shares sum to 100%", abs(sum(r["share"] for r in rows) - 1) < 1e-9)
    models = an.top_model_rows(parsed)
    check("models sorted by units, not file order",
          [m["count"] for m in models] == [20, 12, 7], models)
    check("labels title-cased, model codes and acronyms kept",
          [m["label"] for m in models] == ["Ford F-150", "GMC Sierra 1500", "Ford Mustang Mach-E"],
          [m["label"] for m in models])
    check("model share is of vehicles sold (20 of 80 = 25%)", abs(models[0]["share"] - 0.25) < 1e-9)
    check("VDP share counts new + used VDP visits only (500 of 1000)",
          abs(an.vdp_visit_share(parsed) - 0.5) < 1e-9, an.vdp_visit_share(parsed))

    payload = an.payload_facts(parsed)
    check("payload carries no VINs", "1FMEE9BP3SLB32625" not in json.dumps(payload))
    check("payload keeps the full, unfolded mix (Service and Other both present)",
          {m["label"] for m in payload["traffic_mix"]} >= {"Service", "Other"})
    check("payload names what sold as 'vehicles_sold_since' (the wording the prompt teaches)",
          payload["vehicles_sold_since"] == 80)


def check_real_ted_britt():
    print("\nReal Ted Britt, August 2026 -- export + Analyst for the same group and month")
    for path in (ATTRIBUTION_TB930, ANALYST_TB_AUG, TEMPLATE):
        if not path.exists():
            skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_TB930))
    analyst = an.parse_analyst_facts(str(ANALYST_TB_AUG))
    raw = json.loads(ANALYST_TB_AUG.read_text(encoding="utf-8"))

    check("export's own unique-visitor count is 2,671 (advertiser tab)",
          attribution.attributed_unique_visitors == 2671, attribution.attributed_unique_visitors)
    check("...and the Analyst's matches it, so no mismatch warning fires",
          an.visitor_count_mismatch(analyst, attribution.attributed_unique_visitors) is None)
    check("August Analyst overlaps the Jul 27 - Aug 31 report period",
          an.periods_overlap(analyst, attribution.flight_start, attribution.flight_end),
          (attribution.flight_start, attribution.flight_end))
    mix = an.traffic_mix_rows(analyst)
    check("slide's traffic table sums to the file's visits_total (8,694)",
          sum(r["visits"] for r in mix) == raw["totals"]["visits_total"] == 8694,
          sum(r["visits"] for r in mix))
    check("a truck-shop site with only visits in by_site doesn't break parsing",
          "tedbritttruckshop-group-central-site" in analyst["by_site"])

    OUT_DIR.mkdir(exist_ok=True)
    with_path = OUT_DIR / "TB930_with_analyst.pptx"
    _, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(with_path),
        client_name="Ted Britt Ford & Ted Britt Chantilly", goals_bullets=[],
        whats_next_bullets=["Analyst check: what's-next placeholder"], analyst=analyst)
    check("deck builds with no fit/column warnings", not warnings, warnings)
    prs = Presentation(str(with_path))
    url = next(s for s in prs.slides if s.has_notes_slide
               and slide_map.notes_key(s) == "report:url_report")
    texts = [sh.text_frame.text for sh in slide_map.iter_all_shapes(url.shapes) if sh.has_text_frame]
    tables = {sh.name: sh.table for sh in slide_map.iter_all_shapes(url.shapes)
              if getattr(sh, "has_table", False) and sh.has_table}
    all_text = " ".join(texts) + " ".join(
        c.text for t in tables.values() for r in t.rows for c in r.cells)
    check("no token left unfilled on the slide", "{{" not in all_text)
    check("section headers rewritten for the Analyst data",
          ra._ANALYST_MIX_HEADER in texts and ra._ANALYST_MODELS_HEADER in texts, texts)
    check("old URL-report headers gone", "VISITORS BY INTENT" not in texts, texts)
    mix_table = tables["IntentSummaryTable"]
    check("traffic table header is Page type | Visits | % of visits (no Converted column)",
          [c.text for c in mix_table.rows[0].cells] == list(ra._ANALYST_MIX_COLUMNS),
          [c.text for c in mix_table.rows[0].cells])
    check("first traffic row is New VDP, 2,818 visits (the file's largest category)",
          [c.text for c in mix_table.rows[1].cells][:2] == ["New VDP", "2,818"],
          [c.text for c in mix_table.rows[1].cells])
    model_table = tables["TopUrlTable"]
    check("models table's top row is the F-150 at 98 units (file's top_sellers[0])",
          [c.text for c in model_table.rows[1].cells][:2] == ["Ford F-150", "98"],
          [c.text for c in model_table.rows[1].cells])
    check("subtitle is the Sales Assist's approved shape, counts from the file (1,696 / 918)",
          "Of the 1,696 vehicles attributed visitors shopped, 918 have since sold." in texts,
          texts)
    check("section headers use the Analyst's own names (Traffic Mix, Top Sold Models)",
          "TRAFFIC MIX" in texts and "TOP SOLD MODELS" in texts, texts)
    check("footnote defines 'sold' the Analyst's way and names the source and month",
          any("Auto-Sales Analyst, August 2026" in t
              and "removed from the dealer's live inventory after receiving attributed traffic" in t
              for t in texts), texts)
    check("package is structurally clean", not package_check.check_package(str(with_path)),
          package_check.check_package(str(with_path)))

    without_path = OUT_DIR / "TB930_without_analyst.pptx"
    ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(without_path),
        client_name="Ted Britt Ford & Ted Britt Chantilly", goals_bullets=[],
        whats_next_bullets=["Analyst check: what's-next placeholder"])
    prs2 = Presentation(str(without_path))
    url2 = next(s for s in prs2.slides if s.has_notes_slide
                and slide_map.notes_key(s) == "report:url_report")
    texts2 = [sh.text_frame.text for sh in slide_map.iter_all_shapes(url2.shapes) if sh.has_text_frame]
    check("without the Analyst (and no vertical) the fallback URL report still builds, "
          "relabelled for visits", "TOP PAGES BY ATTRIBUTED VISITS" in texts2, texts2)

    facts = ra.build_facts_payload(attribution, None, analyst=analyst,
                                   client_name="Ted Britt Ford & Ted Britt Chantilly")
    check("facts payload carries the Analyst slice", facts["analyst"] is not None)
    check("Ted Britt's franchise makes are Ford, Chevrolet and Lincoln -- so Ford-heavy "
          "movement there is inventory mix, not audience proof",
          facts["analyst"]["franchise_makes"] == ["CHEVROLET", "FORD", "LINCOLN"],
          facts["analyst"]["franchise_makes"])
    check("scan was 30 days after period end (Sept 30 vs Aug 31) -- no timing item",
          facts["analyst"]["days_scanned_after_period_end"] == 30
          and not any("Auto-Sales Analyst" in i for i in app.attr_actionable_review_items(facts)))
    check("...and None without it", ra.build_facts_payload(attribution, None)["analyst"] is None)
    sentence = ("918 of the 1,696 vehicles attributed visitors shopped have since sold, a 54.1% "
                "look-to-book rate worth an estimated $38.1M; 57% of visits hit a vehicle "
                "detail page and the F-150 led with 98 units, with Ford at 528.")
    violations = app._attr_draft_number_violations([("thread", sentence)], facts)
    check("the facts-only checker traces every Analyst number a narrative would cite",
          not violations, violations)
    invented = app._attr_draft_number_violations(
        [("thread", "an estimated $41.2M in inventory sold")], facts)
    check("...and still flags one the file doesn't contain", bool(invented), invented)
    prompt = app.build_attr_draft_prompt(facts)
    check("the drafting prompt carries the Analyst section and its data",
          "Auto-Sales Analyst" in prompt and '"vehicles_sold_since": 918' in prompt)


REACH_FIXTURES = [
    "MW attribution excel.xlsx", "Premion Website Attribution Cardinal.xlsx",
    "Premion Website Attribution and Reach Extension (13).xlsx",
    "Premion Website Attribution and Reach Extension (15).xlsx",
    "Premion Website Attribution and Reach Extension TB.xlsx",
    "Premion Website Attribution and Reach Extension TB930.xlsx",
]


def check_fallback_url_report():
    """The fallback "Where Visitors Went" -- what every report without the
    Analyst file gets. Two fixes guarded here: (1) no share on it exceeds
    100% (the export's per-page unique-visitor counts were being summed
    into class "reach", 285% for Ted Britt's "Other pages"), and (2) an auto
    site's pages are named the Analyst's way, never by raw path."""
    print("\nFallback URL slide -- shares bounded, automotive names")
    present = [REPO / f for f in REACH_FIXTURES if (REPO / f).exists()]
    if not present:
        skip("no real attribution exports present")
    for path in present:
        attribution = ai.parse_attribution_export(str(path))
        for vertical in (None, "auto"):
            rows = (ra.intent_summary_rows(attribution, vertical=vertical)
                    + ra.top_url_rows(attribution, vertical=vertical))
            classes = ra.intent_facts(attribution, vertical=vertical)["classes"]
            worst = max([r["_visit_share_raw"] for r in rows]
                        + [r["reach"] for r in rows + classes if r.get("reach") is not None]
                        + [c["visit_share"] for c in classes])
            check(f"{path.name} ({vertical or 'no vertical'}): no share or reach over 100% "
                  f"(max {worst:.0%})", worst <= 1)
            share_sum = sum(r["_share_raw"] for r in ra.intent_summary_rows(attribution, vertical=vertical))
            check(f"{path.name} ({vertical or 'no vertical'}): Traffic Mix shares sum to 100%",
                  abs(share_sum - 1) < 1e-6, share_sum)
    try:
        ra.assert_shares_within_bounds([{"label": "Other pages", "_share_raw": 2.85}])
        raised = False
    except ValueError:
        raised = True
    check("the bounds assertion itself fires on the old 285% figure", raised)

    if not (ATTRIBUTION_TB930.exists() and ANALYST_TB_AUG.exists() and TEMPLATE.exists()):
        skip("Ted Britt August fixtures not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_TB930))
    old = {r["intent"]: r for r in ra.intent_summary_rows(attribution)}
    check("Ted Britt with no vertical: 'Other pages' is now a share of visits, well under 100% "
          "(was 285% of visitors)", old["other"]["_share_raw"] < 1, old["other"]["share"])
    raw = json.loads(ANALYST_TB_AUG.read_text(encoding="utf-8"))
    analyst_mix = {m["label"]: m["visits"] for m in raw["totals"]["traffic_mix"]}
    ours = {c["label"]: c["visits"] for c in ra.intent_facts(attribution, vertical="auto")["classes"]}
    shared = [label for label in analyst_mix if label != "Other"]
    mismatched = {label: (ours.get(label), analyst_mix[label]) for label in shared
                  if ours.get(label) != analyst_mix[label]}
    check("auto categories reproduce the Analyst file's own Traffic Mix, same names, same "
          "visit counts (New VDP 2,818, Used VDP 2,113, ...)", not mismatched, mismatched)
    check("the Analyst's 'Other' equals our Other + Finance + Trade-In (the two classes we "
          "split out of it)",
          ours.get("Other pages", 0) + ours.get("Finance / Credit App", 0)
          + ours.get("Trade-In", 0) == analyst_mix["Other"],
          (ours.get("Other pages"), ours.get("Finance / Credit App"), ours.get("Trade-In"),
           analyst_mix["Other"]))

    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / "TB930_fallback_auto.pptx"
    _, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out), client_name="Ted Britt",
        goals_bullets=[], whats_next_bullets=["fallback check"], vertical="auto")
    prs = Presentation(str(out))
    url = next(s for s in prs.slides if s.has_notes_slide
               and slide_map.notes_key(s) == "report:url_report")
    text = " ".join(
        [sh.text_frame.text for sh in slide_map.iter_all_shapes(url.shapes) if sh.has_text_frame]
        + [c.text for sh in slide_map.iter_all_shapes(url.shapes)
           if getattr(sh, "has_table", False) and sh.has_table
           for r in sh.table.rows for c in r.cells])
    check("no raw path ('Searchnew.Aspx', '.aspx') anywhere on the fallback slide",
          "searchnew" not in text.lower() and ".aspx" not in text.lower(), text[:300])
    check("vehicle rows are named as vehicles (2026 Ford F-150 Lariat)",
          "2026 Ford F-150 Lariat" in text, text[:400])
    check("no percentage over 100 on the slide",
          all(float(p) <= 100 for p in re.findall(r"(\d+(?:\.\d+)?)%", text)),
          re.findall(r"(\d+(?:\.\d+)?)%", text))
    check("fallback deck builds without fit warnings", not warnings, warnings)


def main():
    check_parser()
    check_gates()
    check_slide_helpers()
    check_real_ted_britt()
    check_fallback_url_report()
    print()
    print(f"{len(failures)} failure(s), {len(skipped)} skipped" if failures
          else f"All checks passed ({len(skipped)} skipped)")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
