"""The Burch Oil review (2026-11-16) -- a two-ZIP-list, one-option annual HVAC
plan. Synthetic names and figures throughout; the real client's notes and
avails PDF are never committed.

The shape: an avails document with two ZIP lists ("Zip List #1", "Zip List
#2"), each with a homeowner audience and an HVAC-intender audience. The
notes name the lists ("Growth Areas", "Home Base"), buy only the HVAC rows
at stated monthly counts (200,000 and 100,000 at a $34 CPM), add a 300,000
retargeting line at the $5.50 rate card, and run January-December 2027.

What it holds the app to, each from the review:
  1. A line the notes sized themselves never holds Generate; above 75% of
     its avails it gets a soft note.
  2. Campaign Specs "Budget & Allocation" is written from the plan rows on
     a single-option plan, not "--".
  3. The avails slide shows only the audience on the plan, drops the Total
     row (one audience, or overlapping ones), never states a share over
     50%, and fits the vertical title.
  4. A case study goes in once -- one variant, the plain one without
     broadcast -- and an HVAC plan doesn't pre-select pest control.
  5. The notes' names for the ZIP lists reach the plan, the map's groups,
     the avails slide and Campaign Specs.
  6. The attribution overview drops Polk and Arrivalist for a non-auto,
     non-travel deck.

Every expected figure is worked from the notes' own numbers by hand
(200,000 x $34 / 1000 = $6,800), never read back from the code under test.

    python tests/test_burch_oil_fixes.py
"""
import copy
import io
import os
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")
os.environ.setdefault("PROPOSAL_BUILDER_TODAY", "2026-11-16")

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import avails_grid                             # noqa: E402
import case_study_select                       # noqa: E402
import slide_map                               # noqa: E402

from PIL import Image                          # noqa: E402
from pptx import Presentation                  # noqa: E402
from pptx.util import Inches                   # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


CLIENT = "QA-TEST-Ridgeway Heating"
NOTES = f"""Client: {CLIENT}, Leonardtown, MD. Vertical: home services / HVAC.
HVAC installation, repair and maintenance (furnaces, heat pumps, AC).
Goal: grow HVAC in new areas (Zip List #1) while keeping a steady presence in the home base
(Zip List #2).
Flight: Jan 1 - Dec 31, 2027 (12 months), billed monthly. One option only.
Audience: Home Services HVAC intenders only -- not the homeowner rows.
Plan lines (monthly):
- Premion Streaming TV, Zip List #1 (Growth Areas: Prince Frederick, Waldorf, Indian Head):
  200,000 impressions/month at $34 CPM.
- Premion Streaming TV, Zip List #2 (Home Base: Lexington Park, California, Leonardtown):
  100,000 impressions/month at $34 CPM.
- Streaming Retargeting - Display, both zip lists: 300,000 impressions/month at $5.50 CPM.
Expected totals: $11,850/month, $142,200 for the year.
"""
# The notes' own figures, and the plan they add up to (by hand).
CTV_CPM, RT_CPM = 34.0, 5.50
GROWTH_IMPS, HOME_IMPS, RT_IMPS = 200_000, 100_000, 300_000
MONTHLY = GROWTH_IMPS * CTV_CPM / 1000 + HOME_IMPS * CTV_CPM / 1000 + RT_IMPS * RT_CPM / 1000
MONTHS = 12

HVAC = "RETAIL Home Services HVAC"
HOMEOWNER = "DEMO Homeowner"
LISTS = {
    "Zip List #1": {"zips": ["20678", "20601", "20640"], "avails": {HOMEOWNER: 230_000, HVAC: 220_000}},
    "Zip List #2": {"zips": ["20653", "20619", "20650"], "avails": {HOMEOWNER: 180_000, HVAC: 190_000}},
}


def groups():
    rows = []
    for n, (name, spec) in enumerate(LISTS.items()):
        for m, (audience, avails) in enumerate(spec["avails"].items()):
            gid = f"g{n}{m}"
            rows.append({
                "id": gid, "op": "AND", "name": name, "terms": [audience],
                "geo_def": {"kind": "zips", "zips": list(spec["zips"])},
                "resolved_zips": list(spec["zips"]), "resolved_markets": ["washington_hagerstown"],
                "entity_id": gid, "entity_label": "", "entity_locked": False,
                "avails_monthly": avails, "avails_full_flight": avails * MONTHS,
                "avails_doc_monthly": avails, "avails_doc_full_flight": avails * MONTHS,
                "avails_doc_start": "2027-01-01", "avails_doc_end": "2027-12-31",
                "include_in_plan": False, "include_locked": False,
                "color": "#4C78A8" if audience == HVAC else "#F58518", "color_locked": False,
                "avails_source": "RFPID-000001", "avails_rfpi": f"49{n}{m}",
                "avails_geo_raw": name, "avails_geo_variant": name, "avails_geo_label": name,
            })
    return rows


ALLOCATION_STRATEGY = ("High-penetration strategy: the heaviest weight goes to the Growth Areas "
                       "to build awareness where Burch Oil is expanding, while a steady "
                       "maintenance level keeps the Home Base locked in.")
DRAFT = {
    "client_name": CLIENT, "vertical": "home_improvement", "target_markets": [],
    "total_budget": 0, "total_budget_period": "month", "breakout": "monthly",
    "stated_flight": {"months": 12, "start": "2027-01-01", "end": "2027-12-31", "year_stated": True},
    "group_selection": {"mode": "named", "match": ["Home Services HVAC"]},
    "group_allocations": [
        {"match": ["Zip List #1"], "allocation": {"impressions": GROWTH_IMPS, "period": "month"}},
        {"match": ["Zip List #2"], "allocation": {"impressions": HOME_IMPS, "period": "month"}}],
    "group_cpm": CTV_CPM,
    "media_plan_lines": [{"product": "streaming_retargeting_display",
                          "allocation": {"impressions": RT_IMPS, "period": "month"}}],
    "geo_names": [{"avails_name": "Zip List #1", "label": "Growth Areas"},
                  {"avails_name": "Zip List #2", "label": "Home Base"}],
    "options": None, "total_tv": False, "show_sov": False, "audiences": [],
    "attribution": ["web", "sales"], "sports": [],
    "campaign_specs": {
        "goals": ["Grow HVAC business in new areas while holding the home base."],
        "audience": ["Households in-market for HVAC installation, repair or maintenance."],
        "geography": ["Zip List #1: Growth areas -- Prince Frederick, Waldorf, Indian Head.",
                      "Zip List #2: Home base -- Lexington Park, California, Leonardtown."],
        "budget": [], "placements": ["Premion Streaming TV in Zip List #1 and Zip List #2.",
                                     "Streaming Retargeting Display across both."],
        "timing": ["January 1 - December 31, 2027."]},
    "strategy_summary": "", "unresolved": [], "unresolved_internal": [],
    # Seller-proposed: the notes add the plan up, they don't give a client
    # budget (2026-11-16, budget on Campaign Specs).
    "budget_source": "seller", "stated_budget": None,
    "allocation_strategy": ALLOCATION_STRATEGY,
}


class _StubSt:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply_draft(live=None):
    """`live` is a rendered session to draft into (its widget generations
    and versions), as the Draft button does on the page."""
    real = app.st
    stub = _StubSt()
    stub.session_state.update({k: live[k] for k in ("avails_version", "group_name_generation")
                               if live is not None and k in live})
    stub.session_state.update({
        "targeting_groups": groups(), "market_choice": "DC",
        "flight_start": date(2027, 1, 1), "flight_end": date(2027, 12, 31),
        "avails_basis": app.AVAILS_BASIS_MONTHLY, "avails_mode": True,
        "avails_documents": [{"name": "synthetic_avails.pdf", "sha1": "0" * 40, "lines": 4,
                              "rfpid": "RFPID-000001", "total_impressions": 820_000 * MONTHS}],
        "draft_source_notes": NOTES, "client_name": CLIENT,
    })
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(DRAFT), notes=NOTES)
    finally:
        app.st = real
    return stub.session_state


# ---------------------------------------------------------------------------
# Pure checks
# ---------------------------------------------------------------------------
def check_units():
    print("\nUnits: counts, grid totals, caption, case-study variants, names")
    check("a stated count is read from the notes",
          {200_000, 100_000, 300_000} <= app.stated_impression_figures(NOTES),
          sorted(app.stated_impression_figures(NOTES)))
    check("a dollar figure or a year is not a count",
          not ({34, 2027, 11850} & app.stated_impression_figures("$34 CPM, $11,850 in 2027")))

    def grid(spec):
        return avails_grid.build_grid(
            [{"id": f"{a}{g}", "terms": [a], "op": "AND", "name": g,
              "geo_def": {"kind": "text", "label": g}, "avails_monthly": v}
             for a, g, v in spec], lambda grp: grp["avails_monthly"])
    check("overlapping audiences in one geography: no Total row",
          avails_grid.show_total_row(grid([("A", "X", 1), ("B", "X", 2)])) is False)
    check("one audience: no Total row (it only repeats the row)",
          avails_grid.show_total_row(grid([("A", "X", 1), ("A", "Y", 2)])) is False)
    check("different audiences in different geographies: Total row kept",
          avails_grid.show_total_row(grid([("A", "X", 1), ("B", "Y", 2)])) is True)

    rows = [{"monthly_impressions": 200_000, "matched_avails_monthly": 220_000, "geo": "Growth Areas",
             "is_flat_fee": False},
            {"monthly_impressions": 100_000, "matched_avails_monthly": 190_000, "geo": "Home Base",
             "is_flat_fee": False}]
    caption = app.avails_plan_caption([{"name": "Plan"}], [{"preview_rows": rows}])
    check("the avails caption states no share above 50%", "%" not in (caption or ""), caption)
    low = [dict(r, monthly_impressions=20_000) for r in rows]
    caption_low = app.avails_plan_caption([{"name": "Plan"}], [{"preview_rows": low}])
    check("a low share is still stated", "under 11%" in (caption_low or ""), caption_low)

    body = "Client challenge results website visits broadcast premion retargeting strategy"
    facts = [(body + " total tv results", set()), (body, {"logo"}), ("A different page", set())]
    check("variants go in once -- plain without broadcast",
          case_study_select.pick_slides(facts, False, total_tv_study=True) == [0, 2])
    check("variants go in once -- branded with broadcast",
          case_study_select.pick_slides(facts, True, total_tv_study=True) == [1, 2])
    check("one vault study under two titles is one study",
          case_study_select.same_study("Local Clinic -- Brand Lift in a Top 10 Market",
                                       "Local Clinic -- Top 10 Market Brand Lift"))
    check("a plain and a Total TV row of one study are one study",
          case_study_select.same_study("HVAC Company -- Full-Funnel Home Services",
                                       "HVAC Company -- Total TV Full-Funnel Home Services"))
    check("two different companies are not",
          not case_study_select.same_study("HVAC Company -- Website Attribution",
                                           "Pest Control Company -- Website Attribution"))
    hvac = case_study_select.specialties_of(NOTES, "home_improvement")
    check("the notes make this an HVAC plan", "hvac" in hvac, hvac)
    check("pest control doesn't fit an HVAC plan",
          not case_study_select.fits_specialty("Pest Control Company -- Total TV", hvac,
                                               "home_improvement"))
    check("a study naming no specialty fits",
          case_study_select.fits_specialty("Home Services Leader -- Attribution", hvac,
                                           "home_improvement"))

    text = app.geo_names_in_text("Zip List #1: Growth areas -- Waldorf. Also Zip List #2.",
                                 {"Zip List #1": "Growth Areas", "Zip List #2": "Home Base"})
    check("Campaign Specs names follow the notes",
          text == "Growth Areas -- Waldorf. Also Home Base.", text)
    names = {"Zip List #1": "Growth Areas", "Zip List #2": "Home Base"}
    for drafted, want in (("Growth Areas (Zip List #1): Calvert County.", "Growth Areas: Calvert County."),
                          ("Growth Areas - Zip List #1 and Zip List #2 (Home Base)",
                           "Growth Areas and Home Base")):
        got = app.geo_names_in_text(drafted, names)
        check(f"either order reads the name once: {drafted!r}", got == want, got)


# ---------------------------------------------------------------------------
# The draft
# ---------------------------------------------------------------------------
def option_rows(state):
    return state["plan_options"][0]["rows"]


def check_draft(state):
    print("\nThe draft: the notes' names, the stated counts, Budget & Allocation")
    rows = option_rows(state)
    by_geo = {r["Geo"]: r for r in rows}
    ctv = [r for r in rows if r["Tactic"].startswith("Premion Streaming TV")]
    check("two CTV lines, one per ZIP list", len(ctv) == 2, [r["Tactic"] for r in rows])
    check("the plan's Geo cells use the notes' names",
          {"Growth Areas", "Home Base"} <= set(by_geo), list(by_geo))
    rt = [r for r in rows if r["Tactic"].startswith("Streaming Retargeting")]
    check("the retargeting line names both areas the client's way",
          rt and "Zip List" not in rt[0]["Geo"] and "Growth Areas" in rt[0]["Geo"],
          rt and rt[0]["Geo"])
    names = {g["name"] for g in state["targeting_groups"]}
    check("the groups (map legend, slide headers) carry the notes' names",
          names == {"Growth Areas", "Home Base"}, names)
    selected = [g for g in state["targeting_groups"] if g.get("include_in_plan")]
    check("only the HVAC rows are on the plan",
          sorted(g["terms"][0] for g in selected) == [HVAC, HVAC], [g["terms"] for g in selected])
    monthly = sum(float(r["Cost"]) for r in rows)
    check(f"the plan costs ${MONTHLY:,.0f} a month, as the notes add up",
          abs(monthly - MONTHLY) < 1, monthly)
    budget = state.get("budget_text") or ""
    check("seller-proposed: 'Show budget on Campaign Specs' starts off",
          state.get("show_budget_on_specs") is False, state.get("show_budget_on_specs"))
    check("Budget & Allocation is the drafted strategy, with no dollars",
          budget == ALLOCATION_STRATEGY and "$" not in budget, budget)
    check("... and no review row about the budget",
          not [i for i in state.get("draft_review_items") or [] if i.get("kind") == "budget_specs"],
          [i.get("kind") for i in state.get("draft_review_items") or []])
    for key in ("geography_text", "placements_text"):
        check(f"Campaign Specs {key.split('_')[0]} uses the notes' names",
              "Zip List" not in (state.get(key) or "") and "Growth Areas" in (state.get(key) or ""),
              state.get(key))


def check_sanity(state):
    print("\nThe avails check: a stated count is asked for")
    options = state["plan_options"]
    groups_by_id = {g["id"]: g for g in state["targeting_groups"]}
    results = [app.compute_plan_totals(o["rows"], o["breakout"], MONTHS, "",
                                       groups_by_id=groups_by_id) for o in options]
    findings = app.plan_sanity_findings(options, results, state.get("draft_plan_checks"), NOTES,
                                        date(2027, 1, 1), groups_by_id=groups_by_id)
    hard = [f for f in findings if not f.get("soft")]
    soft = [f for f in findings if f.get("soft")]
    check("no check holds Generate", not hard, [f["text"] for f in hard])
    # 200,000 of 220,000 = 91% (soft note); 100,000 of 190,000 = 53% (nothing).
    check("Growth Areas at 91% of its avails gets a soft note",
          any("Growth Areas" in f["text"] and "91%" in f["text"] for f in soft),
          [f["text"] for f in soft])
    check("Home Base at 53% gets nothing", not any("Home Base" in f["text"] for f in findings),
          [f["text"] for f in findings])
    rates_only = app.plan_sanity_findings(
        options, results, state.get("draft_plan_checks"),
        "CTV at $34 CPM, retargeting at $5.50 per thousand. 200,000 and 100,000 impressions.",
        date(2027, 1, 1), groups_by_id=groups_by_id)
    check("the 1.5x budget check ignores CPMs and skips notes with no budget",
          not any(f["key"].startswith("budget:") for f in rates_only),
          [f["text"] for f in rates_only])
    check("... but still fires against a stated budget",
          any(f["key"].startswith("budget:") for f in app.plan_sanity_findings(
              options, results, state.get("draft_plan_checks"), "Budget is $5,000 a month at $34 CPM.",
              date(2027, 1, 1), groups_by_id=groups_by_id)))
    unstated = app.plan_sanity_findings(options, results, state.get("draft_plan_checks"),
                                        "Grow HVAC in two areas.", date(2027, 1, 1),
                                        groups_by_id=groups_by_id)
    check("the same plan with the counts NOT in the notes still holds Generate",
          any(not f.get("soft") and "didn't ask" in f["text"] for f in unstated),
          [f["text"] for f in unstated])


# ---------------------------------------------------------------------------
# The deck
# ---------------------------------------------------------------------------
def _case_study_files(root):
    """Synthetic vault: a Total TV HVAC study in two variants (the branded
    one carries a picture its twin lacks), the same study again as a plain
    streaming row, a second Total TV study in two variants, a pest control
    study, and a residential HVAC one-pager."""
    logo = root / "logo.png"
    Image.new("RGB", (40, 20), (200, 30, 30)).save(logo)
    body = ("HVAC Company client challenge consumers navigating category pricing website "
            "attribution broadcast premion retargeting results visits strategy campaign")
    specs = {
        "cs_hvac": ("Local HVAC Company -- Total TV Full-Funnel Home Services Campaign",
                    ["total_tv", "broadcast_tv", "premion_streaming_tv"],
                    [(body + " total tv results", False), (body, True)],
                    "Local HVAC company drove website visits with Total TV."),
        "cs_hvac_dup": ("Local HVAC Company -- Full-Funnel Home Services Campaign",
                        ["premion_streaming_tv"], [(body, False)],
                        "Local HVAC company drove website visits."),
        "cs_heat": ("Heating & Cooling Contractor -- Total TV Reach Campaign",
                    ["total_tv", "broadcast_tv", "premion_streaming_tv"],
                    [("Heating cooling contractor furnace results visits " * 3 + " total tv", False),
                     ("Heating cooling contractor furnace results visits " * 3, True)],
                    "A heating and cooling contractor grew reach with Total TV."),
        "cs_pest": ("Pest Control Company -- Streaming TV with OTT Retargeting",
                    ["premion_streaming_tv"], [("Pest control results " * 5, False)],
                    "Pest control company drove site actions."),
        "cs_res": ("Residential HVAC Leader -- Website Attribution",
                   ["premion_streaming_tv"], [("Residential HVAC heating cooling results " * 4, False)],
                   "Regional HVAC company drove attributed visitors."),
    }
    rows, files = [], {}
    for n, (cid, (title, products, slides, summary)) in enumerate(specs.items()):
        prs = Presentation()
        images = []
        for i, (text, with_logo) in enumerate(slides):
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = text
            if with_logo:
                slide.shapes.add_picture(str(logo), Inches(1), Inches(5))
            png = root / f"{cid}_{i}.png"
            Image.new("RGB", (160, 90), (20 * n, 40 * i, 90)).save(png)
            images.append(str(png))
        path = root / f"{cid}.pptx"
        prs.save(path)
        files[cid] = (str(path), images)
        rows.append({"id": cid, "title": title, "summary": summary, "verticals": ["home_improvement"],
                     "products": products, "storage_path": f"{cid}.pptx",
                     "slide_images": [f"images/{cid}/slide_{i + 1:03d}.png" for i in range(len(slides))],
                     "filename": f"{cid}.pptx", "added_by": None,
                     "date_added": f"2026-10-0{9 - n}T00:00:00"})
    return rows, files


def generate(state, root):
    from streamlit.testing.v1 import AppTest

    rows, files = _case_study_files(root)
    captured = {}
    saved = {name: getattr(db, name) for name in (
        "fetch_case_studies", "case_study_file", "case_study_images", "log_proposal")}
    real_personalize = assembly.personalize
    real_append = assembly.append_case_studies

    def spy_personalize(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"], captured["fill_data"] = prs, fill_data
        return warnings

    def spy_append(prs, sources):
        captured["case_study_sources"] = sources
        return real_append(prs, sources)

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        captured["form_json"] = form_json
        return "00000000-0000-0000-0000-000000000000", None

    db.fetch_case_studies = lambda active_only=True: (copy.deepcopy(rows), None)
    db.case_study_file = lambda cid, storage_path: files[cid][0]
    db.case_study_images = lambda cid, paths: list(files[cid][1])
    db.log_proposal = spy_log
    assembly.personalize = spy_personalize
    assembly.append_case_studies = spy_append
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=900)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Scenario Suite"
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return None, at, str(at.exception[0].value)
        button = [b for b in at.button if b.label == "Generate proposal"]
        if not button:
            return None, at, "no Generate button"
        if button[0].disabled:
            return None, at, f"Generate is held: {at.session_state['_generate_blockers']}"
        button[0].click().run()
        if at.exception:
            return None, at, str(at.exception[0].value)
    finally:
        for name, fn in saved.items():
            setattr(db, name, fn)
        assembly.personalize = real_personalize
        assembly.append_case_studies = real_append
    return captured, at, None


def _key_slide(prs, prefix):
    return [s for s in prs.slides if str(slide_map.notes_key(s) or "").startswith(prefix)]


def check_deck(state):
    print("\nThe deck: avails slide, Campaign Specs, case studies, attribution overview")
    # A blank field is filled at Generate by the same rule.
    state = {**state, "budget_text": "", "_auto_budget_text": ""}
    with tempfile.TemporaryDirectory() as tmp:
        captured, at, error = generate(state, Path(tmp))
        check("Generate runs and isn't held", error is None, error)
        if error:
            return
        prs, fill = captured["prs"], captured["fill_data"]

        budget = fill["campaign_specs"]["BUDGET_BULLETS"]
        check("Campaign Specs Budget & Allocation (blank at Generate) is the strategy, no dollars",
              budget != ["--"] and budget and not any("$" in b for b in budget)
              and any("Growth Areas" in b for b in budget), budget)

        grid = fill["avails"].get("grid")
        check("the avails slide shows only the audience on the plan",
              grid and [r["key"] for r in grid["rows"]] == [HVAC],
              grid and [r["key"] for r in grid["rows"]])
        check("its columns are the notes' area names",
              grid and grid["geo_headers"] == ["Growth Areas", "Home Base"],
              grid and grid["geo_headers"])
        avails_slide = (_key_slide(prs, "targeting_avails_template") or [None])[0]
        check("the avails slide is in the deck", avails_slide is not None)
        if avails_slide is not None:
            table = next(s for s in avails_slide.shapes if s.has_table).table
            first_cells = [row.cells[0].text.strip() for row in table.rows]
            check("no Total row under one audience", "Total" not in first_cells, first_cells)
            caption = fill["avails"].get("caption") or ""
            check("the caption states no share over 50%", "%" not in caption, caption)
            title = next((s for s in slide_map.iter_all_shapes(avails_slide.shapes)
                          if s.has_text_frame and s.text_frame.text.strip() == "Home Improvement"),
                         None)
            check("the vertical title is on the slide", title is not None)
            if title is not None:
                frame = title.text_frame
                usable = (title.width - frame.margin_left - frame.margin_right) / 12700
                drawn = assembly._widest_line_pt(frame, avails_slide, 1.0)
                check("the vertical title fits its box", drawn <= usable,
                      f"{drawn:.0f}pt drawn in {usable:.0f}pt")

        sources = captured.get("case_study_sources") or []
        titles = [s.get("title") for s in sources]
        check("pest control isn't pre-selected for an HVAC plan",
              not any("Pest" in (t or "") for t in titles), titles)
        check("the HVAC study the vault holds twice goes in once",
              sum("Local HVAC Company" in (t or "") for t in titles) == 1, titles)
        hvac = next((s for s in sources if "Local HVAC Company" in (s.get("title") or "")), None)
        check("... as the plain row, with no broadcast on this plan",
              hvac and hvac["title"].startswith("Local HVAC Company -- Full-Funnel"),
              hvac and hvac.get("title"))
        heat = next((s for s in sources if "Heating & Cooling" in (s.get("title") or "")), None)
        check("a two-variant Total TV study goes in as one slide, the plain variant",
              heat and len(heat["images"]) == 1 and heat["images"][0].endswith("cs_heat_0.png"),
              heat and heat.get("images"))
        logged = {c["id"]: c.get("slides") for c in captured["form_json"]["case_studies"]}
        check("the picked slides are logged, so a rebuild inserts the same ones",
              logged.get("cs_heat") == [0] and logged.get("cs_res") == [0], logged)

        overview = _key_slide(prs, "core:attribution_overview")
        text = " ".join(slide_map.extract_slide_text(s) for s in overview)
        check("the attribution overview is in the deck", bool(overview))
        check("... without Polk or Arrivalist on a home-services deck",
              overview and "Polk" not in text and "Arrivalist" not in text, text[:200])
        check("... and still names Sales Conversion Attribution",
              "Sales Conversion Attribution" in text)


def check_rename_after_render():
    """The page renders the imported names first; the draft renames after.
    Each group's cached "Geo label" box and the D2 grid must not write the
    old name back (found walking the real page, 2026-11-16)."""
    print("\nThe rename survives a page that already showed the document's names")
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=900)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Scenario Suite"
    for key, value in {"targeting_groups": groups(), "market_choice": "DC",
                       "flight_start": date(2027, 1, 1), "flight_end": date(2027, 12, 31),
                       "avails_basis": app.AVAILS_BASIS_MONTHLY, "avails_mode": True,
                       "client_name": CLIENT}.items():
        at.session_state[key] = value
    at.run()
    live = {k: at.session_state[k] for k in ("avails_version", "group_name_generation")
            if k in at.session_state}
    drafted = apply_draft(live)
    for key, value in drafted.items():
        if not key.startswith("_draft_pending"):
            at.session_state[key] = value
    at.run()
    at.run()
    names = sorted({g["name"] for g in at.session_state["targeting_groups"]})
    check("two renders after the draft, the groups still carry the notes' names",
          names == ["Growth Areas", "Home Base"], names)


def apply_draft_with(**changes):
    global DRAFT
    saved = DRAFT
    DRAFT = {**copy.deepcopy(saved), **changes}
    try:
        return apply_draft()
    finally:
        DRAFT = saved


def check_budget_choice():
    print("\nBudget on Campaign Specs: per proposal, by whose budget it is")
    client = apply_draft_with(budget_source="client",
                              stated_budget={"amount": MONTHLY * MONTHS, "period": "year"})
    check("a client budget: the box starts on", client.get("show_budget_on_specs") is True,
          client.get("show_budget_on_specs"))
    check("... and Budget & Allocation leads with the client's figure as a fit",
          client.get("budget_text") == f"Built to your ${MONTHLY * MONTHS:,.0f} annual budget: "
                                       + ALLOCATION_STRATEGY, client.get("budget_text"))
    rfp = apply_draft_with(budget_source="rfp", stated_budget={"amount": 150000, "period": "year"})
    check("an RFP budget the notes don't state is not quoted (no figure to fit to)",
          rfp.get("show_budget_on_specs") is True and rfp.get("draft_stated_budget") is None
          and f"${MONTHLY:,.0f}/month" in (rfp.get("budget_text") or ""), rfp.get("budget_text"))
    unclear = apply_draft_with(budget_source=None)
    rows = [i for i in unclear.get("draft_review_items") or [] if i.get("kind") == "budget_specs"]
    check("an unclear source: box off, one review row",
          unclear.get("show_budget_on_specs") is False and len(rows) == 1
          and "hidden (budget source unclear)" in rows[0]["row"]["text"],
          [(i.get("kind"), (i.get("row") or {}).get("text")) for i in rows])
    check("... and no dollars on Campaign Specs", "$" not in (unclear.get("budget_text") or ""),
          unclear.get("budget_text"))
    # The rule written from the rows (no drafted sentence), on and off.
    option = {"name": "A", "breakout": "monthly", "rows": [
        {"Tactic": "Premion Streaming TV", "Geo": "Growth Areas", "Cost": 6800, "Type": "Rate"},
        {"Tactic": "Premion Streaming TV", "Geo": "Home Base", "Cost": 3400, "Type": "Rate"}]}
    off = app.budget_specs_text([option], MONTHS)
    check("from the rows: the heavier area leads, no figures",
          off.startswith("Weighted strategy: the heaviest weight goes to Growth Areas")
          and "$" not in off, off)
    on = app.budget_specs_text([option], MONTHS, show_budget=True, source="client",
                               stated={"amount": 150000, "period": "year"})
    check("on: the client's figure, the plan's own total noted where it differs",
          on.startswith("Built to your $150,000 annual budget (the plan comes to $122,400): "), on)
    two = app.budget_specs_text([option, dict(option, name="B")], MONTHS)
    check("multi-option, off: one strategy line per option, no dollars",
          two.splitlines()[0].startswith("A: Weighted") and "$" not in two, two)
    check("a drafted sentence naming a figure is not used",
          app.drafted_allocation_strategy("Heaviest weight, $6,800 a month, in Growth Areas.",
                                          [], [option]) == "")


def main():
    check_units()
    check_rename_after_render()
    check_budget_choice()
    state = apply_draft()
    check_draft(state)
    check_sanity(state)
    check_deck(state)
    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("All Burch Oil checks passed.")


if __name__ == "__main__":
    main()
