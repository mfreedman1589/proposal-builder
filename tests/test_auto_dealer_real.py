"""The auto group review fixes against the real exports (local only: client
data, in the gitignored tests/fixtures/auto_group/). The synthetic versions CI
runs are tests/test_auto_dealer_report.py.

    python tests/test_auto_dealer_real.py

Expected values are the QA answer key's for "the auto group fixture
(2026-11)" -- the figures the rep's own corrected deck shows. Store and
client names are read from the files, never written here.
"""
import os
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

from pptx import Presentation  # noqa: E402

import analyst_import as an  # noqa: E402
import app  # noqa: E402
import attribution_import as ai  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures" / "auto_group"
TEMPLATE = REPO / "REPORT_MASTER_v0_19.pptx"   # the site slide (built by build_v0_19.py)
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:400]}" if detail else ""))


def v1_draft():
    """The original tool output's threads, in shape: a confirmation
    highlight, the used-vs-new thread with its 'shift toward used' action,
    and the watch list as a highlight with a listing check."""
    return {"threads": [
        {"head": "Visitors Reached Confirmation Pages", "anchor": "goal",
         "finding": "4 confirmation-page visits -- Thankyou drew 1 visitors.",
         "meaning": "The Thankyou page drew the most confirmation-page visits.",
         "takeaway_head": "Visitors Are Completing Actions", "kind": "confirmation",
         "protected": True},
        {"head": "Vehicle Research Driving the Lot", "anchor": "goal",
         "finding": "Of the 1,321 vehicles our audience viewed, 361 have since sold -- a 27.3% "
                    "overall Look-to-Book rate -- with used vehicles moving off the lot at a higher "
                    "rate: 29.6% Look-to-Book vs. 16.7% for new.",
         "meaning": "Used inventory is seeing stronger post-view movement.",
         "takeaway_head": "Used Inventory Moving at Higher Rate",
         "action": "Shift Streaming TV creative emphasis toward used vehicle inventory to align "
                   "the message with the audience behavior."},
        {"head": "Missed Opportunities Watch List", "anchor": "signal",
         "finding": "Vehicles like the 2016 Chevrolet Corvette (37 visits) and the 2026 Ford "
                    "Mustang GT (22 visits) drew strong campaign traffic and are still listed.",
         "meaning": "The watch list flags the highest-traffic vehicles still available.",
         "action": "For each vehicle on the Missed Opportunities watch list, check the VDP for "
                   "missing photos, a 'Call for Price' button, or pricing that may be out of step."},
        {"head": "TV Sent Them Searching", "anchor": "signal",
         "finding": "About half of visits with referral data came from organic search.",
         "meaning": "The campaign sent viewers to search for the dealership.",
         "action": "Dynamic ads can feature the 2026 Ford Mustang GT (22 visits) and the 2016 "
                   "Chevrolet Corvette (37 visits) to the households that viewed them.",
         "action_tier": 2}]}


def main():
    paths = [FIXTURES / n for n in ("attribution.xlsx", "delivery.xlsx", "analyst.json")]
    if not all(p.exists() for p in paths):
        print("SKIP -- the auto group fixtures aren't present")
        return 0
    app.install_market_lookup()
    a = ai.parse_attribution_export(str(paths[0]))
    d = ai.parse_delivery_export(str(paths[1]))
    analyst = an.parse_analyst_facts(str(paths[2]))
    client = a.client_name
    package = ra.audience_package_phrase(d)
    a = ra.with_audience_names(a, {}, package=package)
    stores = ra.guess_client_stores(analyst, client)
    group = next(s for s in analyst["sites"] if s.get("is_group_site"))
    store = next(s for s in analyst["sites"] if not s.get("is_group_site"))
    facts = ra.build_facts_payload(a, d, goals=["Drive vehicle sales"], vertical="auto",
                                   client_name=client, analyst=analyst,
                                   analyst_client_stores=stores, store_lead="combined")
    af = facts["analyst"]

    print("1/2. The story and the one-sheet sidebar")
    check("vehicle page visits are Used VDP + New VDP: 3,174 (the old panel counted 141 URLs)",
          af["vdp_page_visits"] == 3174, af["vdp_page_visits"])
    check("sidebar: 3,174 / 1,321 / 361 / $12.2M",
          ra.auto_sidebar_rows(af) == [("3,174", "Vehicle page visits"), ("1,321", "Vehicles viewed"),
                                       ("361", "Viewed vehicles sold"), ("$12.2M", "Est. value sold")],
          ra.auto_sidebar_rows(af))
    draft = app._enforce_draft_rules(v1_draft(), facts, a)
    hi, ta, wn = ra.distribute_threads(draft["threads"])
    check("highlights lead with the traffic and sold threads (3,174 of 4,227; 361, est. $12.2M)",
          "(3,174 of 4,227) were vehicle detail pages: 1,330 visitors looked at 1,321" in hi[0][1]
          and hi[1][1].startswith("361 of those vehicles have since sold, an est. $12.2M in value "
                                  "(27.3% Look-to-Book)."), hi[:2])
    text = " ".join(f"{h} {x}" for h, x in hi + ta).lower()
    check("no confirmation, watch-list or used-vs-new highlight or takeaway",
          not any(w in text for w in ("confirmation", "thankyou", "watch list", "corvette",
                                      "used inventory")), text)

    print("\n3/8. What's Next")
    check("no 'shift toward used'; one watch-list idea, under Ideas to consider",
          not any("toward used" in w.lower() for w in wn)
          and sum(1 for w in wn if app._mentions_watchlist(w, facts) or ra._is_listing(w)) == 1
          and ra.IDEAS_HEADING in wn, wn)

    print("\nA1. One Spanish-language audience")
    merged = ra.with_targeting(a, None, d)
    check("the two lines are one audience: 153,205 impressions, 8.02%",
          [(r.label, r.delivered_impressions, round(r.attributed_rate * 100, 2))
           for r in merged.by_audience] == [(ra.SPANISH_AUDIENCE_NAME, 153205, 8.02)],
          merged.by_audience)
    check("no audience comparison (no households/frequency side by side)",
          facts["audiences_compared"] is None and facts["audience_merged"] is not None)

    print("\nA2. Group site vs. store site")
    sites = ra.site_facts(merged, analyst, client)
    rows = [(s["visitors"], s["page_visits"], s["attributed"]) for s in sites["sites"]]
    check("per site: 1,105 / 2,781 / 7,860 and 672 / 1,446 / 4,425; total 1,330 / 4,227 / 12,285",
          rows == [(1105, 2781, 7860), (672, 1446, 4425)]
          and (sites["total"]["visitors"], sites["total"]["page_visits"], sites["total"]["attributed"])
          == (1330, 4227, 12285), (rows, sites["total"]))
    check("overlap 447 -> 'about 450 visited both'; roughly 1 in 3",
          sites["overlap"] == 447 and "about 450 visited both" in sites["headline"]
          and "Roughly 1 in 3 visitors looked at both." in sites["narrative"],
          (sites["headline"], sites["narrative"]))
    check("the group site reads '<domain> (group site)'", sites["sites"][0]["label"]
          == f"{group['domain']} (group site)" and sites["sites"][0]["is_group"])

    print("\n6/7. Group, stores, models")
    check("the store is the client's store (by domain), never the group site",
          stores == [store["domain"]], stores)
    check("its page-title dealer name is replaced by the client's name",
          an.looks_like_page_title(store["dealer_name"])
          and [s["store"] for s in af["client_stores"]] == [client], af["client_stores"])
    check("store figures: 93 of 352 sold, est. $3,695,352, used 40.6%",
          [(s["vehicles_sold_since"], s["vehicles_viewed"], s["est_value_sold"],
            s["look_to_book_pct_used"]) for s in af["client_stores"]] == [(93, 352, 3695352, 40.6)])
    check("the sold highlight carries the store's figures",
          f"At {client}, 93 have sold, with used vehicles at 40.6%." in hi[1][1], hi[1])
    check("top sold models: Ford F-150 32 (23 + 9 merged), no 'Ford F150' row",
          an.top_model_rows(analyst)[0] == {"label": "Ford F-150", "count": 32,
                                            "share": 32 / 361}
          and not any(r["label"] == "Ford F150" for r in an.top_model_rows(analyst)),
          an.top_model_rows(analyst))
    missed = af["missed_opportunities"]
    check("watch list: 9 vehicles, the 4Runner once at 20 visits, the F-250 named F-250",
          len(missed) == 9 and sum("4Runner" in m["vehicle"] for m in missed) == 1
          and any(m["vehicle"].endswith("Ford F-250") for m in missed), [m["vehicle"] for m in missed])

    print("\n9/10. Footnotes and the export date")
    check("totals differ: 150,206 delivered vs. 153,205 attribution",
          ra.impression_basis_note(a, d, "delivery") == (
              "Impressions: 150,206 delivered, per the delivery report. Attributed rates use the "
              "attribution report's own impression count, which differs."))
    check("accruing: Oct 21-31 on a Nov 5 export",
          ra.export_date_note(ra.export_date_facts(a, d)) ==
          "Data as of Nov 5, 2026. Visits from Oct 21–31 are still arriving (attributed visits "
          "keep coming in for 15 days after an ad), so late-October figures will rise.")

    if not TEMPLATE.exists() and (REPO / "REPORT_MASTER_v0_18.pptx").exists():
        import subprocess
        subprocess.run([sys.executable, str(REPO / "build_v0_19.py"), "REPORT_MASTER_v0_18.pptx",
                        str(TEMPLATE)], check=True, capture_output=True, cwd=str(REPO))
    if not TEMPLATE.exists():
        print("  SKIP  the deck checks -- REPORT_MASTER_v0_18.pptx not present")
    else:
        kwargs, _violations = app.apply_attr_draft(draft, facts, attribution=a, client_name=client)
        out = Path(tempfile.mkdtemp()) / "auto_group.pptx"
        ra.build_report_deck(str(TEMPLATE), a, d, str(out), client_name=client,
                             goals_bullets=["Drive vehicle sales"], whats_next_bullets=wn,
                             vertical="auto", analyst=analyst, analyst_client_stores=stores,
                             **kwargs)
        prs = Presentation(str(out))
        by_key = {slide_map.notes_key(s): s for s in prs.slides}
        group_slide = by_key["report:analyst_group"]
        rows = [[c.text for c in r.cells] for r in next(
            s for s in group_slide.shapes if s.name == "AnalystGroupTable").table.rows][1:]
        check("scoreboard: group site labelled, the client's store named and tagged 'Client'",
              sorted((r[0], r[-1]) for r in rows) == sorted(
                  [(f"{group['dealer_name']} (group site)", ""), (client, "Client")]), rows)
        deck_text = " ".join(sh.text_frame.text for s in prs.slides for sh in s.shapes
                             if sh.has_text_frame)
        check("the page title never appears; group-wide inventory is noted",
              store["dealer_name"] not in deck_text and "group-wide inventory" in deck_text)
        check("the impressions footnote is on the highlights and ZIP slides",
              all(any(sh.name == "ImpressionBasisNote" for sh in by_key[k].shapes)
                  for k in ("report:highlights", "report:zip_analysis")))
        check("no by-audience or audience-comparison slide; the site slide is in",
              "report:attribution_breakdown" not in by_key and "report:audience_comparison"
              not in by_key and "report:site_comparison" in by_key, list(by_key))
        watch = by_key["report:analyst_watchlist"]
        check("the Watch List slide stays, nine rows",
              len(next(s for s in watch.shapes if s.name == "AnalystWatchlistTable").table.rows) == 10)
        summary = Path(tempfile.mkdtemp()) / "summary.pptx"
        ra.build_summary_slide(str(TEMPLATE), str(summary), attribution=a, delivery=d,
                               client_name=client, threads=draft["threads"],
                               accepted_optimizations=[], vertical="auto", analyst_facts=af)
        text = {s.name: s.text_frame.text for s in Presentation(str(summary)).slides[0].shapes
                if s.has_text_frame}
        check("one-sheet sidebar: From site visit to sale, 3,174 vehicle page visits",
              text.get("SidebarStatDetail") == "From site visit to sale"
              and text.get("SidebarStatValue") == "3,174"
              and text.get("SidebarSub3Value") == "$12.2M", text)
        check("one-sheet site line: 1,105 and 672, about 450 visited both",
              any("1,105 visitors reached" in v and "about 450 visited both" in v
                  for v in text.values()), text)

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
