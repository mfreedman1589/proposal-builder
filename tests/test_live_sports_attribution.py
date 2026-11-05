"""Live Sports website attribution in the Attribution Report Builder (2026-11).

    python tests/test_live_sports_attribution.py

Runs on synthetic workbooks (tests/live_sports_fixtures.py), so CI covers it;
tests/test_live_sports_real.py checks the real exports and the QA answer key
locally. The deck is built on template v0_17, produced here by
build_v0_17.py from v0_16 (in CI, v0_16 is the skeleton stand-in).

What it holds the feature to:
  * the slot checks: a file that looks like the other attribution type
    warns; several RFPIDs in the sports file warn only when they don't match
    the delivery file's sports block; the sports label comes from the
    package type, never a creative name;
  * unique visitors are never summed -- only delivered and attributed
    impressions and the rate built from their sums get a combined value;
  * the same-period rule: matched flights compare unique visitor rates,
    flights further apart never do, both show whole-flight attributed rates;
  * no day of week anywhere for sports, and none reaching the optimization
    engine in a sports-only report;
  * conversions/sales only when either is above zero;
  * the three report shapes (standard only, sports only, both) and
    Simple vs Extended, by slide;
  * the page: the upload slot, the label and detail controls, the
    wrong-slot and RFPID warnings, no RFPID gate for sports.
"""
import os
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import attribution_import as ai  # noqa: E402
import live_sports_fixtures as fx  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402
from pptx import Presentation  # noqa: E402

failures = []
TMP = Path(tempfile.mkdtemp())


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:300]}" if detail else ""))


def files(**sports_overrides):
    std = ai.parse_attribution_export(
        str(fx.attribution_workbook(TMP / "std.xlsx", "standard", start=date(2026, 9, 7),
                                    weeks=8)))
    name = f"sp_{abs(hash(repr(sorted(sports_overrides.items()))))}.xlsx"
    sp = ai.parse_attribution_export(
        str(fx.attribution_workbook(TMP / name, "sports", **sports_overrides)), slot="sports")
    delivery = ai.parse_delivery_export(str(fx.delivery_2026_11(TMP / "delivery.xlsx")))
    return std, sp, delivery


def check_slots():
    print("Upload slots: wrong-slot warning, RFPID check, the sports label")
    std, sp, delivery = files()
    check("a sports file in the sports slot: no warning",
          ai.slot_mismatch_warning(sp, "sports") is None)
    check("a standard file in the sports slot warns (it has audience/channel tabs)",
          "standard Website Attribution" in (ai.slot_mismatch_warning(std, "sports") or ""))
    check("a sports file in the standard slot warns",
          "Live Sports" in (ai.slot_mismatch_warning(sp, "standard") or ""))
    check("a standard file in the standard slot: no warning",
          ai.slot_mismatch_warning(std, "standard") is None)
    check("two RFPIDs in the sports slot raise no 'covers 2 RFPIDs' warning",
          not any("RFPID" in w for w in sp.warnings), sp.warnings)
    check("matching RFPIDs: no RFPID warning", ai.sports_rfpid_warning(sp, delivery) is None)
    _, odd, _ = files(rfpids=["RFPID-900001", "RFPID-912345"])
    warning = ai.sports_rfpid_warning(odd, delivery) or ""
    check("a mismatch warns, naming both sides",
          "RFPID-912345" in warning and "RFPID-900002" in warning, warning)
    check("no delivery sports block: nothing to check against",
          ai.sports_rfpid_warning(sp, None) is None)
    for package, label in (("PREM TV - LIVE NFL REGULAR SEASON", "Live NFL"),
                           ("PREM TV - LIVE NCAAF REGULAR SEASON", "Live College Football"),
                           ("PREM TV - RON", "Live Sports"),
                           ("PREM TV - LIVE NFL REGULAR SEASON + PREM TV - LIVE NBA REGULAR SEASON",
                            "Live Sports")):
        check(f"label from {package!r} is {label!r}",
              ai.sports_label_from_package(package) == label,
              ai.sports_label_from_package(package))


def check_facts():
    print("\nFacts: never summed visitors, the same-period rule, no day of week")
    std, sp, delivery = files()
    summary = ra.campaign_summary_facts(std, ra.sports_view(sp), delivery, "Live NFL")
    rows = {r["metric"]: r for r in summary["rows"]}
    check("unique visitors are never combined",
          rows["Unique website visitors"]["combined"] is None
          and rows["Unique visitor rate"]["combined"] is None)
    check("the overlap footnote is there", summary["footnote"] == ra.SPORTS_OVERLAP_FOOTNOTE)
    check("delivered impressions combine (OTT + the delivery file's sports block)",
          rows["Delivered impressions"]["combined"] == fx.OTT_TOTAL + fx.NEW_SPORTS_TOTAL,
          rows["Delivered impressions"])
    check("attributed impressions combine",
          rows["Attributed impressions"]["combined"] == std.attributed_impressions
          + sp.attributed_impressions)
    expected_rate = ((std.attributed_impressions + sp.attributed_impressions)
                     / (std.delivered_impressions + sp.delivered_impressions))
    check("combined rate = sum attributed / sum delivered",
          abs(rows["Attributed rate"]["combined"] - expected_rate) < 1e-12)
    check("no conversions row while both are zero", "Attributed conversions" not in rows)

    same = ra.campaign_comparison_facts(std, ra.sports_view(sp), delivery, "Live NFL")
    keys = [m["key"] for m in same["measures"]]
    check("matched flights count as the same period", same["same_period"], same["period_note"])
    check("...and compare unique visitor rates", "unique_visitor_rate" in keys, keys)
    check("every comparison line is written by code",
          all(isinstance(line["text"], str) and line["text"] for line in same["lines"]))
    check("no line ranks one campaign above the other",
          not any(w in line["text"].lower() for line in same["lines"]
                  for w in ("better", "worse", "outperform", "underperform")),
          [line["text"] for line in same["lines"]])

    long_std = ai.parse_attribution_export(str(fx.attribution_workbook(
        TMP / "long.xlsx", "standard", start=date(2026, 3, 30), weeks=31)))
    apart = ra.campaign_comparison_facts(long_std, ra.sports_view(sp), delivery, "Live NFL")
    keys = [m["key"] for m in apart["measures"]]
    check("flights months apart are different periods", not apart["same_period"])
    check("...never compare unique visitor rates", "unique_visitor_rate" not in keys, keys)
    check("...still compare whole-flight attributed rates", "attributed_rate" in keys)
    check("...and state each campaign's own dates",
          apart["standard_period"]["label"] in apart["period_note"]
          and apart["sports_period"]["label"] in apart["period_note"], apart["period_note"])
    check("nothing is trimmed: the attributed rate is the whole-flight one",
          next(m for m in apart["measures"] if m["key"] == "attributed_rate")["standard"]
          == long_std.attributed_rate)

    check("the parsed sports file has game-day rows (the input this guards against)",
          len(sp.by_day_of_week) > 0)
    check("sports_view empties day of week", ra.sports_view(sp).by_day_of_week == [])
    check("...without touching the parsed export", len(sp.by_day_of_week) > 0)
    only = ra.build_facts_payload(sp, delivery, goals=["Drive visits"], sports_only=True)
    check("a sports-only payload carries no day-of-week facts",
          not (only.get("response_profile") or {}).get("day_of_week"))
    found = ra.optimization_candidates(ra.sports_view(sp), "moderate", ra.OPTIMIZATION_DIMENSIONS)
    check("...and no day-of-week optimization candidates",
          not any(c.get("dimension") == "day_of_week" for v in found.values()
                  if isinstance(v, list) for c in v if isinstance(c, dict)), found)
    check("a sports-only payload's delivery headline is the sports package",
          (only.get("delivery") or {}).get("delivered_impressions") == fx.NEW_SPORTS_TOTAL)
    both = ra.build_facts_payload(std, delivery, goals=["Drive visits"], sports_attribution=sp)
    check("a both-files payload carries the summary and the comparison",
          "campaign_summary" in both and "campaign_comparison" in both)
    check("...and the sports section, without day of week",
          "live_sports_attribution" in both
          and "day_of_week" not in str(both["live_sports_attribution"]))

    facts = ra.sports_attribution_facts(ra.sports_view(sp), delivery, "Live NFL")
    check("no conversions in the facts while both are zero", "conversions" not in facts)
    _, conv, _ = files(conversions=12, sales=0)
    facts = ra.sports_attribution_facts(ra.sports_view(conv), delivery, "Live NFL")
    check("conversions above zero are shown", facts.get("conversions") == 12, facts.get("conversions"))
    check("the target market is a row; the rest is one footnote line",
          facts["market_split"]["outside"]["markets"] == 2
          and "2 other markets" in (facts["outside_market_note"] or ""), facts["outside_market_note"])


# Templates that carry the Live Sports slides, newest first: CI's stand-in is
# built under the ACTIVE template's name, which moves on at each merge.
LIVE_SPORTS_TEMPLATES = ("REPORT_MASTER_v0_18.pptx", "REPORT_MASTER_v0_17.pptx")


def template():
    """The newest template with the Live Sports slides (v0_17 onward) that's
    here, else v0_17 built from v0_16 by build_v0_17.py."""
    for name in LIVE_SPORTS_TEMPLATES:
        if (REPO / name).exists():
            return REPO / name
    out = TMP / "REPORT_MASTER_v0_17.pptx"
    if not out.exists():
        subprocess.run([sys.executable, str(REPO / "build_v0_17.py"),
                        str(REPO / "REPORT_MASTER_v0_16.pptx"), str(out)],
                       check=True, capture_output=True)
    return out


def deck(attribution, delivery, name, **kw):
    out = TMP / f"{name}.pptx"
    ra.build_report_deck(str(template()), attribution, delivery, str(out),
                         client_name="Example Client", goals_bullets=["Drive website visits"],
                         whats_next_bullets=["Keep going"], **kw)
    prs = Presentation(str(out))
    keys = [slide_map.notes_key(s) for s in prs.slides]
    text = "\n".join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
    tables = "\n".join(c.text for s in prs.slides for sh in s.shapes if sh.has_table
                       for row in sh.table.rows for c in row.cells)
    deck.by_key = {slide_map.notes_key(s): _slide_text(s) for s in prs.slides}
    return keys, text + "\n" + tables


def _slide_text(slide):
    parts = [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
    parts += [c.text for sh in slide.shapes if sh.has_table for row in sh.table.rows
              for c in row.cells]
    return "\n".join(parts)


SPORTS_SIMPLE = ["report:live_sports_attribution"]
SPORTS_EXTENDED = ["report:live_sports_attribution", "report:live_sports_response",
                   "report:live_sports_creative_geo"]


def check_decks():
    print("\nThe deck: three report shapes, Simple vs Extended")
    if not any((REPO / name).exists() for name in LIVE_SPORTS_TEMPLATES + ("REPORT_MASTER_v0_16.pptx",)):
        print("  SKIP  no report template from v0_16 on is present")
        return
    std, sp, delivery = files()

    keys, text = deck(std, delivery, "standard")
    check("standard only: no sports attribution, summary or comparison slides",
          not any(k in keys for k in SPORTS_EXTENDED + ["report:highlights_combined",
                                                        "report:campaign_comparison"]), keys)
    check("standard only: the ordinary highlights", "report:highlights" in keys)
    check("every token filled", "{{" not in text)

    keys, text = deck(std, delivery, "both_simple", sports_attribution=sp, sports_label="Live NFL")
    check("both, Simple: the side-by-side summary replaces highlights",
          "report:highlights_combined" in keys and "report:highlights" not in keys, keys)
    check("both, Simple: the sports delivery slide and ONE sports attribution slide",
          "report:live_sports" in keys
          and [k for k in keys if k in SPORTS_EXTENDED] == SPORTS_SIMPLE, keys)
    check("both: the comparison slide", "report:campaign_comparison" in keys)
    check("the overlap footnote is on the deck", ra.SPORTS_OVERLAP_FOOTNOTE in text)
    check("the sports label titles the sports slides", "Live NFL Website Response" in text)
    check("every token filled", "{{" not in text)
    simple_count = len(keys)

    keys, text = deck(std, delivery, "both_extended", sports_attribution=sp,
                      sports_label="Live NFL", sports_detail="extended")
    check("Extended adds exactly the two extended slides",
          len(keys) == simple_count + 2
          and [k for k in keys if k in SPORTS_EXTENDED] == SPORTS_EXTENDED, keys)
    sports_text = "\n".join(deck.by_key.get(k, "") for k in SPORTS_EXTENDED)
    check("no day-of-week content on any sports slide",
          sports_text and not any(w in sports_text.lower() for w in
                                  ("day of week", "monday", "sunday", "weekday")),
          sports_text[:200])
    check("every token filled", "{{" not in text)

    keys, text = deck(sp, delivery, "sports_only", sports_only=True, sports_label="Live NFL",
                      sports_detail="extended")
    check("sports only: no standard-only slides",
          not any(k in keys for k in ("report:attribution_breakdown", "report:response_profile",
                                      "report:url_report", "report:zip_analysis",
                                      "report:delivery_recap", "report:campaign_comparison",
                                      "report:highlights_combined")), keys)
    check("sports only: the sports section in full", all(k in keys for k in SPORTS_EXTENDED))
    check("sports only: no day of week anywhere in the deck", "day of week" not in text.lower())
    check("sports only: the audience defaults to the sports label", "Live NFL viewers" in text)
    check("every token filled", "{{" not in text)

    keys, text = deck(sp, None, "sports_no_delivery", sports_only=True, sports_label="Live NFL")
    check("sports attribution without a sports delivery block: no game list",
          "report:live_sports" not in keys and "report:live_sports_attribution" in keys, keys)
    check("every token filled", "{{" not in text)

    _, conv, _ = files(conversions=12, sales=0)
    _, text = deck(std, delivery, "conversions", sports_attribution=conv, sports_label="Live NFL")
    check("sports conversions above zero reach the slide", "Attributed conversions: 12" in text)
    _, text = deck(std, delivery, "no_conversions", sports_attribution=sp, sports_label="Live NFL")
    check("...and zero conversions don't", "Attributed conversions:" not in text)


def check_page():
    print("\nThe page: the slot, the controls, the warnings")
    try:
        import test_attribution_reports_page as suite
        from streamlit.testing.v1 import AppTest  # noqa: F401
    except Exception as exc:                       # noqa: BLE001
        check("the page suite imports", False, exc)
        return
    suite.install_store()
    sports = fx.attribution_workbook(TMP / "page_sports.xlsx", "sports")
    delivery = fx.delivery_2026_11(TMP / "page_delivery.xlsx")

    at = suite.new_app()
    at.session_state["page_choice"] = "Attribution reports"
    at.session_state["attr_sports_attribution_upload_path"] = str(sports)
    at.session_state["attr_delivery_upload_path"] = str(delivery)
    at.run()
    check("sports only: the page renders", not at.exception, at.exception)
    labels = [t for t in at.text_input if t.key == "attr_sports_label"]
    check("the sports label defaults from the delivery package type",
          labels and labels[0].value == "Live NFL", [t.value for t in labels])
    detail = [r for r in at.radio if r.key == "attr_sports_detail"]
    check("Simple/Extended, Simple by default", detail and detail[0].value == "Simple",
          [r.value for r in detail])
    check("no RFPID gate for the sports file",
          not [c for c in at.checkbox if c.key == "attr_rfpid_confirmed"])
    check("matching RFPIDs: no RFPID warning",
          not any("RFPIDs" in str(w.value) for w in at.warning), [w.value for w in at.warning])

    at2 = suite.new_app()
    at2.session_state["page_choice"] = "Attribution reports"
    at2.session_state["attr_sports_attribution_upload_path"] = str(
        fx.attribution_workbook(TMP / "page_wrong.xlsx", "standard"))
    at2.run()
    check("a standard file in the sports slot: the page still renders", not at2.exception,
          at2.exception)
    check("...and warns it looks like the other type",
          any("looks like a standard Website Attribution export" in str(w.value)
              for w in at2.warning), [w.value for w in at2.warning])

    at3 = suite.new_app()
    at3.session_state["page_choice"] = "Attribution reports"
    at3.session_state["attr_sports_attribution_upload_path"] = str(fx.attribution_workbook(
        TMP / "page_mismatch.xlsx", "sports", rfpids=["RFPID-900001", "RFPID-912345"]))
    at3.session_state["attr_delivery_upload_path"] = str(delivery)
    at3.run()
    check("an RFPID mismatch warns the rep",
          any("don't cover the same RFPIDs" in str(w.value) for w in at3.warning),
          [w.value for w in at3.warning])
    caption = " ".join(str(c.value) for c in at.caption)
    check("sports only: the parsed line shows the sports package, not OTT + sports",
          f"{fx.NEW_SPORTS_TOTAL:,} delivered impressions (live sports)" in caption, caption[:300])
    check("sports only: the client name comes from the delivery file's sports block",
          "Example Client" in caption, caption[:300])

    print("\nThe page, both files: Generate builds the side-by-side deck and logs both")
    store = suite.install_store()
    built = template()
    suite.app.db.report_master_deck = lambda _fallback: (str(built), 16, None)
    at4 = suite.new_app()
    at4.session_state["page_choice"] = "Attribution reports"
    at4.session_state["attr_attribution_upload_path"] = str(
        fx.attribution_workbook(TMP / "page_std.xlsx", "standard", start=date(2026, 9, 7), weeks=8))
    at4.session_state["attr_sports_attribution_upload_path"] = str(sports)
    at4.session_state["attr_delivery_upload_path"] = str(delivery)
    at4.run()
    suite._confirm_advertiser(at4)
    clicked = suite._generate_standalone(at4)
    check("Generate runs cleanly with both files", clicked and not at4.exception, at4.exception)
    if store.upload_report_file_calls:
        keys = [slide_map.notes_key(s) for s in
                Presentation(store.upload_report_file_calls[-1][1]).slides]
        check("the built deck has the side-by-side summary and the comparison",
              "report:highlights_combined" in keys and "report:campaign_comparison" in keys, keys)
    else:
        check("the built deck was stored", False, "no upload_report_file call")
    logged = store.log_report_calls[-1]["report_json"] if store.log_report_calls else {}
    check("the log carries the sports attribution file, not the delivery's sports block",
          (logged.get("sports_attribution") or {}).get("rfpid") == "RFPID-900001 + RFPID-900002"
          and "attributed_unique_visitors" in (logged.get("sports_attribution") or {}),
          {k: logged.get(k) for k in ("sports_only", "sports_label")})
    check("...flagged as a both-files report", logged.get("sports_only") is False)


def main():
    check_slots()
    check_facts()
    check_decks()
    check_page()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
