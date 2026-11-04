"""The 2027 sports lineup (TEGNA sources live sports through the Premion DSP).

    python tests/test_sports_lineup.py

What it guards -- asserted against the rate table as Matt gave it
(2026-11-04), typed out below, never against the code's own table:
  * Rates, minimums and seasons for every package; Live Sports Select is
    the renamed All Live Sports; NCAAF Conferences and Golf at $61.
  * Retired packages are hidden from the picker, the drafting prompt and the
    price panel, but still known: an old proposal's line keeps its name and
    rate, a loaded proposal's retired pick stays selectable, and a draft
    naming one becomes the closest current package with a card -- never
    silently.
  * Live sports rules: a flight outside the season is a card naming the
    season; tentative dates are a note; under the order minimum or a
    geography under $2,500 net is a warning; no audience segments on a
    sports line (the Targeting is the package's copy, NCAAF Conferences
    naming its conference); NCAA packages are blocked for a sports
    betting / casino / gambling / alcohol advertiser, with pro packages
    offered and a rep override; ZIP-level targeting, a retargeting pairing
    and partial months are flagged.
  * Sports Select is a network package, not live sports: no rule applies.
  * The rate card: a table still carrying retired rows (before
    `setup_supabase.py sports` runs) keeps them off the panel, and a
    package it lacks prices at the lineup rate.
  * Plan extras: Placements gains the live creative spec once; Included with
    Campaign gains live sports reporting; Live Sports Select is co-viewing
    eligible.
  * The Attribution Report Builder's sports label maps package types to the
    2027 names, unknown to "Live Sports", and a Sports Select delivery (OTT
    tabs only) isn't read as live sports.

Offline and CI-safe: db.fetch_products and db.fetch_audiences are stubbed.
The master deck's keys are tests/test_sports_deck.py (local only).
"""
import os
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
os.chdir(REPO)

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.fetch_products = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import attribution_import as ai                # noqa: E402
import sports_lineup as sl                     # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


# The rate card as given (net CPM). key: (name, cpm, order min, season, tentative)
D = date
TABLE = {
    "all_live_sports": ("Live Sports Select", 51, 10000, None, False),
    "mlb_reg": ("MLB Regular Season", 52, 15000, (D(2027, 3, 24), D(2027, 9, 26)), False),
    "nba_reg": ("NBA Regular Season", 57, 15000, (D(2026, 10, 20), D(2027, 4, 11)), False),
    "nhl_reg": ("NHL Regular Season", 57, 15000, (D(2026, 10, 20), D(2027, 4, 10)), False),
    "ncaa_basketball": ("NCAA Men's Basketball Regular Season", 55, 15000,
                        (D(2026, 11, 1), D(2027, 3, 7)), False),
    "ncaaf_reg": ("NCAA Football Regular Season", 60, 15000, (D(2027, 8, 28), D(2027, 11, 27)), False),
    "ncaaf_conference": ("NCAA Football Conferences", 61, 15000,
                         (D(2027, 8, 28), D(2027, 11, 27)), False),
    "nfl_reg": ("NFL Regular Season", 62, 15000, (D(2027, 9, 9), D(2028, 1, 9)), True),
    "golf": ("Golf", 61, 15000, (D(2027, 1, 21), D(2027, 8, 29)), True),
    "motorsports": ("Motorsports", 59, 15000, (D(2027, 2, 21), D(2027, 11, 14)), False),
    "sports_select": ("Sports Select", 38, 0, None, False),
}
RETIRED = ["nfl_playoffs", "nfl_home_team", "nba_playoffs", "wnba_reg", "wnba_playoffs",
           "nhl_playoffs", "mlb_playoffs", "ncaaf_playoffs", "ncaaf_home_team", "pga_majors",
           "soccer_pro", "prestige_sports"]


class FakeState(dict):
    def __getattr__(self, k):
        return self[k]


def with_state(state):
    """Run app helpers that read st.session_state against a plain dict."""
    class Stub:
        session_state = state

        def __getattr__(self, name):
            return getattr(app.st, name)
    real = app.st
    app.st = Stub()
    return real


def option(rows, name="", breakout="Monthly", draft_lines=None):
    return {"name": name, "breakout": breakout, "rows": rows, "version": 0,
            "_draft": {"lines": draft_lines or []}}


def row(key, cost, geo="Washington, DC", tactic=None):
    return {"Tactic": tactic or sl.tactic(key), "Geo": geo, "Cost": cost, "Type": "Rate",
            "Targeting": sl.targeting_copy(key)}


def check_rates():
    print("\nThe lineup and rates, as given")
    check("exactly the eleven current packages", set(sl.PACKAGES) == set(TABLE), sorted(sl.PACKAGES))
    for key, (name, cpm, minimum, season, tentative) in TABLE.items():
        p = sl.PACKAGES[key]
        check(f"{key}: {name} ${cpm}, min ${minimum:,}, "
              f"{'full year' if not season else season[0].isoformat() + '..' + season[1].isoformat()}"
              f"{' tentative' if tentative else ''}",
              (p.name, p.cpm, p.order_minimum, p.season, p.tentative)
              == (name, cpm, minimum, season, tentative),
              (p.name, p.cpm, p.order_minimum, p.season, p.tentative))
    check("geo minimum $2,500 net", sl.GEO_MINIMUM == 2500)
    check("every live package is live, Sports Select isn't",
          [k for k in TABLE if not sl.is_live(k)] == ["sports_select"])
    check("NCAA flag on exactly the three NCAA packages",
          sorted(k for k, p in sl.PACKAGES.items() if p.ncaa)
          == ["ncaa_basketball", "ncaaf_conference", "ncaaf_reg"])
    check("March Madness stays out", not any("march" in p.name.lower() for p in sl.PACKAGES.values()))
    for key, (name, cpm, *_rest) in TABLE.items():
        label, rate = app.line_product_spec(f"sport:{key}")
        check(f"line_product_spec(sport:{key}) prices at ${cpm}", rate == cpm, (label, rate))
    check("tactic names", [app.line_product_spec(f"sport:{k}")[0] for k in
                           ("all_live_sports", "mlb_reg", "sports_select")]
          == ["Live Sports Select", "Live Sports - MLB Regular Season", "Sports Select"])
    import setup_supabase
    seeded = {r["key"]: r for r in setup_supabase.sports_rate_rows()}
    check("setup_supabase sports rows: current active at the table's rate",
          all(seeded[f"sport:{k}"]["active"] and seeded[f"sport:{k}"]["default_cpm"] == v[1]
              for k, v in TABLE.items()))
    check("...and every retired one inactive", all(not seeded[f"sport:{k}"]["active"] for k in RETIRED))


def check_retired():
    print("\nRetired: hidden from anything new, still known to old proposals")
    check("the twelve retired packages", sorted(sl.RETIRED) == sorted(RETIRED))
    # Confirmed by Matt, 2026-11-04: each league's playoffs -> its regular
    # season, NFL Home Team -> NFL Regular Season; the rest as first built.
    closest = {"nfl_playoffs": "nfl_reg", "nba_playoffs": "nba_reg", "nhl_playoffs": "nhl_reg",
               "mlb_playoffs": "mlb_reg", "ncaaf_playoffs": "ncaaf_reg", "nfl_home_team": "nfl_reg",
               "ncaaf_home_team": "ncaaf_conference", "pga_majors": "golf",
               "wnba_reg": "all_live_sports", "wnba_playoffs": "all_live_sports",
               "soccer_pro": "all_live_sports", "prestige_sports": "all_live_sports"}
    check("each retired package's closest current one, as confirmed",
          {k: sl.closest_current(k) for k in RETIRED} == closest,
          {k: sl.closest_current(k) for k in RETIRED})
    real = with_state(FakeState())
    try:
        options = app.sports_picker_options()
    finally:
        app.st = real
    check("picker: current packages only", [app.SPORTS[o] for o in options] == list(TABLE), options)
    retired_label = sl.picker_label("nfl_playoffs")
    real = with_state(FakeState(selected_sports=[retired_label]))
    try:
        options = app.sports_picker_options()
    finally:
        app.st = real
    check("a loaded proposal's retired pick stays an option", retired_label in options, options)
    prompt_keys = app.sports_prompt_keys()
    check("the prompt's product list is current packages only",
          prompt_keys == list(TABLE) and not set(RETIRED) & set(prompt_keys))
    label, rate = app.line_product_spec("sport:nfl_playoffs")
    check("an old NFL Playoffs line keeps its name and $66 rate",
          (label, rate) == ("Live Sports - NFL - Playoffs", 66.0), (label, rate))
    check("its Targeting copy is still the live copy",
          app.fixed_targeting_copy("Live Sports - NFL - Playoffs — DC") == sl.LIVE_TARGETING)
    check("a pre-2027 tactic name reads as its package",
          sl.key_for_tactic("Live Sports - All Live Sports") == "all_live_sports"
          and sl.key_for_tactic("Live Sports - NCAAF - Conference — SEC") == "ncaaf_conference")
    rows, _p, sports, notes, _d = app.resolve_drafted_lines(
        [{"product": "sport:nfl_playoffs", "label": "", "allocation": {"percent_of_total": 100}}],
        30000, "Sep 2027 - Nov 2027", "Washington, DC", "Adults", n_months=3)
    check("a draft naming NFL Playoffs is priced as NFL Regular Season",
          rows and rows[0]["Tactic"] == "Live Sports - NFL Regular Season" and rows[0]["CPM"] == 62.0,
          rows)
    card = [n for n in notes if getattr(n, "item", None) and "isn't sold any more" in n]
    check("...and says so on a card the rep can act on (remove the line)",
          card and any(o["id"] == "remove" for o in card[0].item["options"]), notes)
    check("...touching nfl_reg, not nfl_playoffs", sports == {"nfl_reg"}, sports)


def check_seasons():
    print("\nSeason dates")
    check("NBA inside its season", sl.season_fit("nba_reg", D(2026, 11, 1), D(2027, 1, 31)) == "inside")
    check("NFL in Nov 2026 is outside (2027 season only)",
          sl.season_fit("nfl_reg", D(2026, 11, 1), D(2026, 12, 31)) == "outside")
    check("MLB Jan-Apr 2027 is partial", sl.season_fit("mlb_reg", D(2027, 1, 1), D(2027, 4, 30)) == "partial")
    check("Live Sports Select is full year", sl.season_fit("all_live_sports", D(2026, 1, 1), D(2026, 2, 1)) is None)
    nfl_opt = option([row("nfl_reg", 10000)], draft_lines=[{"product": "sport:nfl_reg", "label": ""}])
    items, notes = app.sports_draft_review([nfl_opt], D(2026, 11, 1), D(2026, 12, 31), [])
    check("a draft outside the season gets one card", len(items) == 1, items)
    if items:
        item = items[0]
        check("...naming the season", "Sep 9, 2027 – Jan 9, 2028" in item["question"], item["question"])
        check("...for the rep, not the client", item["audience"] == "internal")
        flight = [o for o in item["options"] if o["id"] == "season"][0]["patch"][0]
        check("...offering to move the flight into the season",
              flight["op"] == "set_flight" and flight["start"] == "2027-09-09", flight)
        check("...or to remove the line",
              any(o["id"] == "remove" for o in item["options"]))
    check("...which says the dates are tentative", items and "(tentative)" in items[0]["question"])
    check("...and no second, separate tentative note", not any("tentative" in n for n in notes), notes)
    _items, notes = app.sports_draft_review([nfl_opt], D(2027, 9, 15), D(2027, 11, 30), [])
    check("in season, tentative dates are a seller note", any("tentative" in n for n in notes), notes)
    findings = app.sports_plan_findings(option([row("nba_reg", 10000)]), D(2026, 11, 1),
                                        D(2027, 1, 31), n_months=3)
    check("NBA in season, $30K over 3 months, one geo: no findings", findings == [], findings)


def check_minimums():
    print("\nMinimums")
    opt = option([row("mlb_reg", 3000)])
    f = app.sports_plan_findings(opt, D(2027, 4, 1), D(2027, 6, 30), n_months=3)
    texts = [x["text"] for x in f if x["kind"] == "minimum"]
    check("$9,000 of MLB is under its $15,000 order minimum",
          any("$9,000" in t and "$15,000" in t for t in texts), texts)
    check("...a warning the rep acts on, not a block",
          all(x["level"] == "warning" for x in f if x["kind"] == "minimum"))
    opt = option([row("all_live_sports", 4000)])
    f = app.sports_plan_findings(opt, D(2027, 4, 1), D(2027, 6, 30), n_months=3)
    check("$12,000 of Live Sports Select clears its $10,000 minimum",
          not [x for x in f if x["kind"] == "minimum"], f)
    opt = option([row("nba_reg", 6000, geo="Washington, DC, Baltimore, Salisbury")],
                 breakout="Full Flight")
    f = app.sports_plan_findings(opt, D(2026, 11, 1), D(2027, 1, 31), n_months=3,
                                 market_labels=["Washington, DC", "Baltimore", "Salisbury"])
    geo = [x["text"] for x in f if "each geography" in x["text"]]
    check("$6,000 across three markets is $2,000 a market, under $2,500",
          len(geo) == 3 and "$2,000" in geo[0], geo)


def check_targeting():
    print("\nTargeting: no audience segments on a sports line")
    rows, _p, _s, notes, _d = app.resolve_drafted_lines(
        [{"product": "sport:nba_reg", "label": "", "segments": ["DEMO Homeowner"],
          "audience_track": "Homeowners", "allocation": {"percent_of_total": 100}}],
        30000, "Nov 2026 - Jan 2027", "Washington, DC", "Adults", n_months=3)
    check("the row carries no segments", rows and not rows[0].get("_segments"), rows)
    check("its Targeting is the package's copy", rows and rows[0]["Targeting"] == sl.LIVE_TARGETING,
          rows and rows[0]["Targeting"])
    check("the dropped segment is named", any("DEMO Homeowner" in n for n in notes), notes)
    rows, *_ = app.resolve_drafted_lines(
        [{"product": "sport:ncaaf_conference", "label": "SEC", "allocation": {"percent_of_total": 100}}],
        30000, "Sep 2027 - Nov 2027", "Washington, DC", "Adults", n_months=3)
    check("NCAAF Conferences names its conference in the Targeting",
          rows and rows[0]["Targeting"].startswith("SEC games"), rows and rows[0]["Targeting"])
    f = app.sports_plan_findings(option([row("nba_reg", 10000)]), D(2026, 11, 1), D(2027, 1, 31),
                                 n_months=3, geo_kinds={"zips"})
    check("ZIP-level targeting elsewhere on the plan: a note that sports runs at DMA/state",
          any(x["kind"] == "geo" for x in f), f)
    opt = option([row("nba_reg", 10000), {"Tactic": "Streaming Retargeting - Display", "Cost": 500}])
    f = app.sports_plan_findings(opt, D(2026, 11, 1), D(2027, 1, 31), n_months=3)
    check("a retargeting line beside live sports is flagged",
          any(x["kind"] == "retargeting" for x in f), f)
    f = app.sports_plan_findings(option([row("nba_reg", 10000)]), D(2026, 11, 1), D(2027, 1, 31),
                                 n_months=3, partial_months=True)
    check("partial months (weekly flighting) are flagged", any(x["kind"] == "weekly" for x in f), f)


def check_ncaa():
    print("\nNCAA restriction")
    for text, category in (("Old Bay Brewing Company", "alcohol"),
                           ("Riverside Casino & Resort", "casino"),
                           ("BetMGM sportsbook launch", "sports betting")):
        found = sl.restricted_category(text)
        check(f"{text!r} reads as {category}", found and found[0] == category, found)
    check("a dental office isn't restricted", sl.restricted_category("Smile Dental Group") is None)
    opt = option([row("ncaa_basketball", 10000)])
    f = app.sports_plan_findings(opt, D(2026, 11, 1), D(2027, 2, 28), ["Old Bay Brewing Company"],
                                 n_months=4)
    block = [x for x in f if x["kind"] == "ncaa"]
    check("an alcohol advertiser on NCAA basketball is blocked",
          block and block[0]["level"] == "error", f)
    check("...the message offers the pro packages",
          block and "NBA Regular Season" in block[0]["text"] and "NFL Regular Season" in block[0]["text"])
    f = app.sports_plan_findings(opt, D(2026, 11, 1), D(2027, 2, 28), ["Old Bay Brewing Company"],
                                 n_months=4, ncaa_confirmed=True)
    check("the rep's override clears it", not [x for x in f if x["kind"] == "ncaa"], f)
    f = app.sports_plan_findings(option([row("nba_reg", 10000)]), D(2026, 11, 1), D(2027, 1, 31),
                                 ["Riverside Casino & Resort"], n_months=3)
    check("a casino on a pro package isn't blocked", not [x for x in f if x["kind"] == "ncaa"], f)


def check_sports_select():
    print("\nSports Select is not live sports")
    opt = option([row("sports_select", 500)])
    f = app.sports_plan_findings(opt, D(2026, 1, 1), D(2026, 1, 31), ["Old Bay Brewing Company"],
                                 n_months=1, geo_kinds={"zips"}, partial_months=True)
    check("no rule applies -- season, minimum, geo, NCAA, weekly", f == [], f)
    check("its Targeting is its own copy", sl.targeting_copy("sports_select") == sl.SPORTS_SELECT_TARGETING)
    check("it adds no live creative spec or reporting line",
          not app.plan_has_live_sports([option([row("sports_select", 500)])]))
    with tempfile.TemporaryDirectory() as tmp:
        import openpyxl
        from live_sports_fixtures import _sheet
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _sheet(wb, "KPI DELIVERY", ("Delivered Impressions", "VCR"), [(500_000, 0.97)])
        _sheet(wb, "IMPRESSIONS BY PRODUCT PACKAGE", ("Package Type", "Delivered Impressions"),
               [("PREM TV - SPORTS SELECT", 500_000)])
        _sheet(wb, "TOP 10 CHANNELS BY IMPRESSIONS", ("Channel Name", "Delivered Impressions"),
               [("Example Sports Network", 500_000)])
        path = Path(tmp) / "delivery.xlsx"
        wb.save(path)
        delivery = ai.parse_delivery_export(str(path))
    check("a Sports Select delivery (OTT tabs) isn't read as live sports",
          delivery.live_sports is None, delivery.live_sports)


def check_plan_extras():
    print("\nCreative spec, reporting, co-viewing")
    live = [option([row("nba_reg", 10000)])]
    text = app.ensure_live_creative_spec("Premium CTV inventory", app.plan_has_live_sports(live))
    check("Placements gains the :15/:30 MP4 spec", text.endswith(sl.LIVE_CREATIVE_SPEC), text)
    check("...once", app.ensure_live_creative_spec(text, True) == text)
    check("...and not without live sports", app.ensure_live_creative_spec("x", False) == "x")
    included = app.build_included_list({}, False, live_sports=True)
    check("Included with Campaign gains live sports reporting", sl.LIVE_REPORTING in included, included)
    check("...geo, impressions, VCR, channel-level",
          all(w in sl.LIVE_REPORTING for w in ("Geo", "Impressions", "VCR", "Channel")))
    check("Live Sports Select is co-viewing eligible",
          app.is_coviewing_eligible_row({"Tactic": "Live Sports Select"}))
    check("Sports Select follows the CTV rule: co-viewing eligible",
          app.is_coviewing_eligible_row({"Tactic": "Sports Select"}))
    check("...while retargeting still isn't",
          not app.is_coviewing_eligible_row({"Tactic": "Streaming Retargeting - Display"}))


def check_rate_card_transition():
    print("\nThe rate card before `setup_supabase.py sports` runs")
    old_rows = [{"key": "sport:all_live_sports", "name": "Live Sports - All Live Sports",
                 "default_cpm": 50, "targeting_copy": sl.LIVE_TARGETING, "display_group": "sports"},
                {"key": "sport:nfl_playoffs", "name": "Live Sports - NFL - Playoffs",
                 "default_cpm": 66, "targeting_copy": sl.LIVE_TARGETING, "display_group": "sports"},
                {"key": "premion_streaming_tv", "name": "Premion Streaming TV", "default_cpm": 32,
                 "targeting_copy": None, "display_group": "premion"}]
    real = db.fetch_products
    db.fetch_products = lambda: (old_rows, None)
    app.load_rate_card.clear()
    try:
        products, sport_cpm, labels, copy_map, warning = app.load_rate_card()
    finally:
        db.fetch_products = real
        app.load_rate_card.clear()
    check("a package the table lacks prices at the lineup rate",
          sport_cpm.get("motorsports") == 59 and sport_cpm.get("sports_select") == 38, sport_cpm)
    check("the name is the lineup's, whatever the table says",
          labels.get("all_live_sports") == "Live Sports Select", labels)
    check("the Sports Select tactic carries its own copy",
          copy_map.get("Sports Select") == sl.SPORTS_SELECT_TARGETING)
    keys = [k for k in sport_cpm if not sl.is_retired(k)]
    check("a retired row still active in the table stays off the panel", "nfl_playoffs" not in keys)


def check_report_label():
    print("\nAttribution Report Builder: sports label")
    for package, label in (("PREM TV - LIVE NFL REGULAR SEASON", "Live NFL"),
                           ("PREM TV - LIVE SPORTS SELECT", "Live Sports Select"),
                           ("PREM TV - LIVE ALL LIVE SPORTS", "Live Sports Select"),
                           ("PREM TV - LIVE MOTORSPORTS", "Live Motorsports"),
                           ("PREM TV - LIVE GOLF", "Live Golf"),
                           ("PREM TV - LIVE NCAA MENS BASKETBALL REGULAR SEASON",
                            "Live College Basketball"),
                           ("PREM TV - LIVE NCAAF CONFERENCES SEC", "Live College Football"),
                           ("PREM TV - LIVE NFL PLAYOFFS", "Live NFL"),
                           ("PREM TV - LIVE WNBA REGULAR SEASON", "Live WNBA"),
                           ("PREM TV - LIVE CORNHOLE", "Live Sports"),
                           ("PREM TV - RON", "Live Sports")):
        got = ai.sports_label_from_package(package)
        check(f"{package!r} -> {label!r}", got == label, got)


def check_form():
    print("\nThe real form: a brewery with NFL and NCAA basketball, Nov-Dec 2026")
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
    for key, value in {"authed": True, "current_user": "Regression Suite", "market_choice": "DC",
                       "flight_start": D(2026, 11, 1), "flight_end": D(2026, 12, 31),
                       "avails_mode": False, "client_name": "Old Bay Brewing Company",
                       "live_sports_enabled": True,
                       "selected_sports": [sl.picker_label("nfl_reg"),
                                           sl.picker_label("ncaa_basketball")]}.items():
        at.session_state[key] = value
    at.run()
    check("renders", not at.exception, at.exception)
    if at.exception:
        return
    picker = next(m for m in at.multiselect if m.label == "Sports packages")
    check("the picker offers the 2027 lineup", [app.SPORTS[o] for o in picker.options] == list(TABLE),
          picker.options)
    errors = [e.value for e in at.error]
    check("the NCAA restriction shows as an error", any("NCAA packages can't run" in e for e in errors),
          errors)
    check("the NFL season shows as a warning",
          any("Sep 9, 2027" in w.value for w in at.warning), [w.value for w in at.warning])
    generate = [b for b in at.button if b.label == "Generate proposal"]
    check("Generate is disabled", generate and generate[0].disabled)
    rows = [r for o in at.session_state["plan_options"] for r in o["rows"] if row_key(r)]
    check("the sports rows carry the package copy",
          rows and all(r["Targeting"] == sl.LIVE_TARGETING for r in rows), rows)
    override = [c for c in at.checkbox if "isn't a sports betting" in c.label]
    check("the rep can clear a false match", bool(override))
    if override:
        override[0].check().run()
        generate = [b for b in at.button if b.label == "Generate proposal"]
        check("...which lifts the block", not at.error and generate and not generate[0].disabled)


def row_key(row):
    return sl.key_for_tactic(row.get("Tactic"))


def main():
    check_rates()
    check_retired()
    check_seasons()
    check_minimums()
    check_targeting()
    check_ncaa()
    check_sports_select()
    check_plan_extras()
    check_rate_card_transition()
    check_report_label()
    check_form()
    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
