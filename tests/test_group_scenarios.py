"""Scenarios 2, 3 and 4 of targeting_groups_test_scenarios.md -- Annapolis
Cars, Visit Hershey & Harrisburg, Wilmington University -- automated against
the real avails documents (tests/group_scenario_fixtures.py transcribes
them; ground-truth totals are each document's own stated figure). Lawn &
Leisure (added 2026-08-23) is a fourth, load-bearing for the identity-
collapse no-op guard: the one real document on hand shaped as a single
group / single audience. Plaza Motors Group (FLOW_REWORK_PLAN.md Phase 3,
commit 6) is a fifth: 2 avail rows for one real-world entity, landed before
entity inference exists (commit 7) so this scenario's own merge expectation
starts as SUM and flips to MAX once that commit lands -- see
group_scenario_fixtures.py's own comment on it.

    python tests/test_group_scenarios.py                # all three, no render
    python tests/test_group_scenarios.py hershey         # one, by name prefix
    python tests/test_group_scenarios.py --render        # also render + inspect the deck
    python tests/test_group_scenarios.py --render --keep # leave the .pptx on disk too

Two tiers, like the rest of this suite:

- **State-level** (always runs, offline, ~seconds): drives the real form
  through AppTest with the fixture's groups/avails/plan rows injected as
  session_state (see group_scenario_fixtures.py's own docstring for why
  groups are built as data rather than replayed through the AND/OR/New-group
  buttons -- that mechanism is already covered elsewhere). Checks row count,
  audience-major ordering, no collapse, resolved markets/zips against the
  document's own figures, avails totals against the document's own stated
  total, and Annapolis's gross markup.
- **Deck-level** (`--render`, needs Windows + PowerPoint, ~a minute per
  scenario): presses the real Generate button, opens the resulting deck in
  PowerPoint (a deck that won't open is a FAILURE, the same rule
  render_scenarios.py uses), and checks the plan slide's table -- row count,
  clearance against the Included-with-Campaign block, no unfilled tokens,
  and the deck's own Full Flight Total against the document's ground truth.
  Images land in deck_render.render_root()/<scenario>/, same convention as
  render_scenarios.py, printed on every run.
"""
import argparse
import os
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db                                      # noqa: E402

# Forces the local brochure-PDF fallback for the audience catalog. Not a
# simulation of "Supabase unreachable" -- this machine's secrets point at a
# real, live project, and the "Update audience usage" admin page has
# actually been used against it, so an unstubbed run scores
# custom_segment_count against whatever's live at test time (real
# workbook-merged data, hundreds of components beyond the fixed 369-segment
# brochure these scenarios' expected_custom_count values were derived
# against) instead of the fixture this test is actually about. Found by
# this exact test going red the moment a real workbook was activated.
# MUST run before `import app`: app.py calls load_audience_catalog() at
# its own module scope, which would otherwise cache the live result before
# this stub ever gets a chance to apply.
db.fetch_audiences = lambda: (None, "stubbed for test isolation -- see comment above")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import package_check                           # noqa: E402
import slide_map                               # noqa: E402
import targeting_groups as tg                  # noqa: E402
import targeting_map as tm                     # noqa: E402
import group_scenario_fixtures as gsf          # noqa: E402

TOKEN_RE = re.compile(r"\{\{[A-Z0-9_]+\}\}")
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def slide_text(slide):
    return slide_map.extract_slide_text(slide)


def run_form(state):
    """Inject a scenario's session_state into a fresh AppTest and run it
    once -- the state-level checks below all read from this, before any
    Generate click, exactly the same injection mechanism
    tests/test_draft_regression.py's build_deck uses for a drafted state."""
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Scenario Suite"
    for key, value in state.items():
        at.session_state[key] = value
    at.run()
    return at


def check_state(scn, state):
    print(f"\n{'=' * 78}\nSTATE  {scn['name']}\n{'=' * 78}")
    at = run_form(state)
    check("form renders with the scenario's groups/avails/plan injected",
          not at.exception, at.exception[0].message[:400] if at.exception else "")
    if at.exception:
        return

    groups = at.session_state["targeting_groups"]
    real_groups = [g for g in groups if g.get("terms")]
    check(f"{len(scn['rows'])} groups survive the real form's own render",
          len(real_groups) == len(scn["rows"]), len(real_groups))

    # Row count + no reseed-to-zero -- the same regression build_deck guards
    # against for a drafted state: a seed-key mismatch on first render would
    # silently zero every row.
    rows = at.session_state["plan_options"][0]["rows"]
    check(f"{len(scn['rows'])} plan rows, none reseeded to $0",
          len(rows) == len(scn["rows"])
          and all(app.is_flat_fee_row(r) or float(r["Cost"]) for r in rows),
          [(r["Tactic"], r["Cost"]) for r in rows])

    # Audience-major ordering + no collapse, as a PROPERTY of the real
    # function the app itself would use to seed lines from these groups --
    # not just "I built the fixture rows in this order by hand". The
    # fixture's own `scn["groups"]` are in the REAL document's own row
    # order now (group_scenario_fixtures.py builds them straight from
    # avails_pdf_import's parse, no hand-reordering) -- Hershey's own page
    # order is NOT audience-major (see test_avails_pdf_import.py's own note
    # on this). Checked structurally rather than by reproducing
    # plan_lines_from_groups' own two-key sort here too (audience_order AND
    # geo_order, both keyed by first appearance): re-deriving the identical
    # algorithm to compare against itself would only prove the function
    # agrees with itself, not that it does what "audience-major" means --
    # every one of one audience's lines contiguous, before every one of the
    # next's, with nothing lost or merged along the way.
    triples = app.plan_lines_from_groups(groups, "", "")
    derived_order = [(a, g) for a, g, _gid in triples]
    audiences_in_order = [a for a, _g in derived_order]
    seen_audiences = []
    for audience in audiences_in_order:
        if not seen_audiences or seen_audiences[-1] != audience:
            seen_audiences.append(audience)
    check("each audience's lines are contiguous -- no audience reappears after a "
          "different one starts (the real document's own page order interleaves "
          "them; this is audience-major specifically because plan_lines_from_groups "
          "un-interleaves it)",
          len(seen_audiences) == len(set(audiences_in_order)), audiences_in_order)
    expected_pairs = {(tg.audience_label(g), tg.geo_label(g, label_for=app._market_display_name))
                      for g in scn["groups"]}
    check("every (audience, geo) pair from the document survives, none lost or invented",
          set(derived_order) == expected_pairs, (set(derived_order), expected_pairs))
    check("every (audience, geo) pair is distinct -- nothing collapsed",
          len(set(derived_order)) == len(derived_order), derived_order)

    # Avails total against the document's own stated ground truth --
    # exactly for a single-month flight (Annapolis), within the documented
    # rounding tolerance of the monthly-storage round-trip for a multi-month
    # one (see group_scenario_fixtures.py's build_hershey/build_wilmington
    # comments for why that tolerance exists and how it's bounded).
    n = scn["n_months"]
    displayed_total = sum(app.avails_to_display(g["avails_monthly"], app.AVAILS_BASIS_FLIGHT, n)
                          for g in real_groups)
    tolerance = len(real_groups) * (n // 2 + 1)
    delta = displayed_total - scn["ground_truth"]
    check(f"avails table total is the document's ground truth ({scn['ground_truth']:,}) "
          f"within the monthly-storage rounding tolerance (+/-{tolerance})",
          abs(delta) <= tolerance, f"delta={delta}")

    # Resolved geography against the document's own zip/market data.
    check_geography(scn, real_groups)

    # FLOW_REWORK_PLAN.md Phase 3's entity acceptance case.
    check_entity_behavior(scn, real_groups)

    # Custom-segment count, and the review-list warning it should (or
    # shouldn't) produce.
    catalog = app.load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
    live_count = tg.custom_segment_count(real_groups, rfp_map)
    check(f"custom_segment_count is {scn['expected_custom_count']} for the real catalog",
          live_count == scn["expected_custom_count"], live_count)
    warnings = [str(w.value) if hasattr(w, "value") else str(w) for w in at.warning]
    over_cap_warning = any("custom" in w.lower() and "targeting groups" in w.lower() for w in warnings)
    if scn["expected_custom_count"] > 1:
        check("the over-the-cap warning is shown (allow-with-a-warning, not a refusal)",
              over_cap_warning, warnings)
    else:
        check("no over-the-cap warning (only one real custom segment)",
              not over_cap_warning, warnings)

    # Gross markup -- Annapolis only (agency_involved=True there).
    if state.get("agency_involved"):
        markup = 1.15
        wrong = [r["Tactic"] for r in rows if app.row_markup(r, markup) != 1.15]
        check("every line is marked Gross (agency markup 1.15, none broadcast-exempt)",
              not wrong, wrong)

    return at


def check_entity_behavior(scn, real_groups):
    """Plaza Motors Group (RFPID-266583) is FLOW_REWORK_PLAN.md Phase 3's
    own acceptance case: 2 avail rows for one real-world entity. Landed in
    commit 6, BEFORE entity inference exists (commit 7) -- see
    group_scenario_fixtures.py's own comment on this scenario for why the
    merge expectation here is SUM, not max, until that commit lands and
    flips it.
    """
    if scn["name"] != "Plaza Motors Group":
        return
    check("Plaza imports to exactly 2 rows/groups -- entity grouping never "
          "merges rows on its own", len(real_groups) == 2, real_groups)
    if len(real_groups) != 2:
        return
    check("the two groups default to SEPARATE entity ids (nothing has grouped "
          "them yet -- entity inference lands in commit 7)",
          tg.entity_id_of(real_groups[0]) != tg.entity_id_of(real_groups[1]),
          [tg.entity_id_of(g) for g in real_groups])

    # Simulate a rep merging the two rows -- the real merge_plan_rows, not a
    # hand-built imitation (same discipline test_share_of_voice.py's own
    # build_option_with_known_lines uses).
    option = app.new_plan_option("Option A", [
        {"Tactic": "Premion Streaming TV", "Flight": "", "Geo": tg.geo_label(g), "Targeting": "",
         "Impressions": 0.0, "CPM": 30.0, "Type": app.ROW_TYPE_RATE, "Cost": 0.0,
         "_group_ids": [g["id"]]}
        for g in real_groups
    ])

    class _Stub:
        def __init__(self, groups):
            self.session_state = {"targeting_groups": groups}
            self.secrets = {}

    real_st, app.st = app.st, _Stub(real_groups)
    try:
        app.merge_plan_rows(option, [0, 1], 1.0)
    finally:
        app.st = real_st
    groups_by_id = {g["id"]: g for g in real_groups}
    merged_avails = app.matched_avails_for_row(option["rows"][0], groups_by_id)
    check(f"TODAY (before entity inference exists), merging Plaza's two rows SUMS "
          f"({gsf.PLAZA_GROUND_TRUTH:,}), since they aren't yet recognized as one "
          f"entity -- commit 7 flips this scenario to MAX ({gsf.PLAZA_ENTITY_MAX:,}) "
          f"once deterministic inference groups them",
          merged_avails == gsf.PLAZA_GROUND_TRUTH, merged_avails)


def check_geography(scn, real_groups):
    if scn["name"] == "Annapolis Cars":
        # Radius resolution is compared, not required to match exactly --
        # per the scenario doc's own framing ("the document lists both, so
        # compare"), because this app's zip-centroid/radius math and the
        # avails system's are two different implementations with no shared
        # source of truth (same caveat CLAUDE.md records for the DMA table
        # itself: a working compilation, not an authoritative one). A large
        # miss would mean the resolver is actually wrong; a couple of
        # boundary zips either way is the expected shape of "two radius
        # implementations, one real answer".
        seen_miles = set()
        for g in real_groups:
            miles = g["geo_def"].get("miles")
            expected_zips = gsf.ANNAPOLIS_ZIPS_10MI if miles == 10 else gsf.ANNAPOLIS_ZIPS_5MI
            ours = set(g["resolved_zips"])
            missing, extra = expected_zips - ours, ours - expected_zips
            overlap = len(expected_zips & ours) / len(expected_zips)
            if miles not in seen_miles:
                seen_miles.add(miles)
                print(f"    ....  {miles}mi: {len(ours)} resolved vs {len(expected_zips)} in the "
                      f"document -- missing {sorted(missing) or 'none'}, extra {sorted(extra) or 'none'}")
            check(f"{g['terms'][0]} / {miles}mi overlaps the document's own zip list by >=85% "
                  f"({len(expected_zips & ours)}/{len(expected_zips)})",
                  overlap >= 0.85, sorted(missing))
            check(f"{g['terms'][0]} / {miles}mi resolves to Baltimore (among any others)",
                  "baltimore" in g["resolved_markets"], g["resolved_markets"])
            if g["resolved_markets"] != ["baltimore"]:
                print(f"    ....  {miles}mi resolved to {g['resolved_markets']}, not Baltimore "
                      f"alone -- worth a human's judgment call, not asserted as a failure")
    elif scn["name"] == "Visit Hershey & Harrisburg":
        for g in real_groups:
            kind = g["geo_def"].get("kind")
            if kind == "markets":
                want = g["geo_def"]["markets"]
                check(f"market-mode group resolves to exactly {want}",
                      g["resolved_markets"] == want, g["resolved_markets"])
        philly_zip_groups = [g for g in real_groups
                             if g["geo_def"].get("kind") == "zips"
                             and "19601" in (g["geo_def"].get("zips") or [])]
        ny_zip_groups = [g for g in real_groups
                        if g["geo_def"].get("kind") == "zips"
                        and "07416" in (g["geo_def"].get("zips") or [])]
        check("the Philly zip add-on resolves into Philadelphia's orbit",
              all("philadelphia" in g["resolved_markets"] for g in philly_zip_groups)
              and bool(philly_zip_groups), [g["resolved_markets"] for g in philly_zip_groups])
        check("the NY zip add-on resolves into New York's orbit",
              all("new_york" in g["resolved_markets"] for g in ny_zip_groups)
              and bool(ny_zip_groups), [g["resolved_markets"] for g in ny_zip_groups])
    elif scn["name"] == "Wilmington University":
        # A "County Option" block is zip-originated, same as a Zip Option --
        # Premion resolves a rep's entered zips to county names and reports
        # BOTH back; the county summary is real but DERIVED, never the
        # resolution input (confirmed directly against the real document:
        # every County Option page carries a genuine Zip Codes table). So
        # this group resolves as `kind == "zips"` from its OWN real 303-zip
        # list, exactly like a Zip Option group -- NOT `kind == "counties"`
        # re-derived from the 10 county NAMES, which used to pull in every
        # zip in all 10 counties (335) rather than the 303 the rep actually
        # entered: real over-targeting against a signed plan, corrected the
        # same day this was found. tests/test_avails_pdf_import.py's own
        # Wilmington section is the parser-level guard for the 303-vs-48
        # continuation-page bug underneath this; this is the resolution-
        # level guard that the app actually uses the real list, not the
        # derived one.
        for g in real_groups:
            check("the county list resolves to more than one DMA, Philadelphia among them",
                  "philadelphia" in g["resolved_markets"] and len(g["resolved_markets"]) > 1,
                  g["resolved_markets"])
            check("resolved from the document's own real zip list (kind == zips), "
                  "not re-derived from the 10 county NAMES",
                  g["geo_def"].get("kind") == "zips", g["geo_def"].get("kind"))
            check("the real zip list is the document's own 303, not the county-derived 335",
                  len(g["geo_def"].get("zips") or []) == 303,
                  len(g["geo_def"].get("zips") or []))
            check("the document's own County Option label survives as the group's name "
                  "(the same convention a Zip/Radius option already gets)",
                  g.get("name") == "New Jersey PA and Delaware Counties", g.get("name"))
    elif scn["name"] == "Lawn & Leisure":
        # The identity-collapse pass (targeting_map._collapse_identical_geography,
        # 2026-08-23) only ever iterates equivalence classes of size >= 2 --
        # a single-audience document is a guaranteed structural no-op, not
        # something handled by a special case. This is the ONE real document
        # on hand shaped that way (1 group, 1 audience, "10 Mile Radius
        # Zips" with no bracketed origin -- resolved from its own zip list
        # directly, not geocoded), so it's the load-bearing real-document
        # guard for that no-op, not just a synthetic one.
        check("exactly one group survives", len(real_groups) == 1, len(real_groups))
        if real_groups:
            g = real_groups[0]
            check("resolved from the document's own zip list (no bracketed "
                  "origin to geocode from)", g["geo_def"].get("kind") == "zips",
                  g["geo_def"].get("kind"))
            entries = tm.legend_entries(tm.groups_with_zips(real_groups))
            check("legend_entries returns exactly one entry, byte-identical to "
                  "the input -- nothing to collapse, nothing enumerated",
                  entries == [(tg.audience_label(g), g.get("color"), [g])], entries)


# ---------------------------------------------------------------------------
# Deck-level: press Generate, open the real deck, inspect its plan slide(s).
# ---------------------------------------------------------------------------

def check_deck(scn, state, out_dir, keep):
    import deck_render

    print(f"\n{'=' * 78}\nDECK  {scn['name']}\n{'=' * 78}")
    if not deck_render.renderer_available():
        print("  SKIP -- rendering needs Windows with PowerPoint and pywin32.")
        return

    captured = {}
    real_personalize = assembly.personalize
    real_prepare_plan = assembly._prepare_media_plan_slide
    real_place_map = assembly.place_targeting_map
    real_log = db.log_proposal

    def spy_personalize(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        captured["fill_data"] = fill_data
        captured["warnings"] = warnings
        return warnings

    def spy_prepare_plan(slide, option):
        result = real_prepare_plan(slide, option)
        captured.setdefault("plan_slide_ids", []).append(slide.slide_id)
        return result

    def spy_place_map(slide, png_bytes):
        captured["avails_slide_id"] = slide.slide_id
        captured["map_png"] = png_bytes
        return real_place_map(slide, png_bytes)

    assembly.personalize = spy_personalize
    assembly._prepare_media_plan_slide = spy_prepare_plan
    assembly.place_targeting_map = spy_place_map
    db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)
    try:
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=900)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Scenario Suite"
        for key, value in state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            check("form renders before Generate", False, at.exception[0].message[:400])
            return
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not generate:
            check("Generate button present", False, [b.label for b in at.button])
            return
        generate[0].click().run()
        if at.exception:
            check("Generate runs without raising", False, at.exception[0].message[:600])
            return
        check("Generate runs without raising", True)
    finally:
        assembly.personalize = real_personalize
        assembly._prepare_media_plan_slide = real_prepare_plan
        assembly.place_targeting_map = real_place_map
        db.log_proposal = real_log

    prs = captured.get("prs")
    if prs is None:
        check("assembly produced a presentation", False, None)
        return

    slides = list(prs.slides)
    texts = [slide_text(s) for s in slides]

    leftover = [(i + 1, tok) for i, text in enumerate(texts) for tok in TOKEN_RE.findall(text or "")]
    check("no unfilled {{TOKEN}}s remain anywhere in the deck", not leftover, leftover[:10])

    payloads = captured["fill_data"]["media_plan_options"]
    by_id = {s.slide_id: i for i, s in enumerate(slides)}
    plan_ids = captured.get("plan_slide_ids") or []
    check("exactly one plan slide (single option)", len(plan_ids) == 1, plan_ids)
    if plan_ids and plan_ids[0] in by_id:
        idx = by_id[plan_ids[0]]
        slide = slides[idx]
        shape = assembly._find_table_shape(slide)
        check("plan slide has a table", shape is not None, None)
        if shape is not None:
            table = shape.table
            payload = payloads[0]
            expected_rows = (1 + len(payload["rows"]) + 1
                             + (1 if payload.get("full_flight_total") else 0))
            check(f"table has {expected_rows} rows (header + {len(payload['rows'])} data "
                  f"+ Monthly Totals{' + Full Flight Total' if payload.get('full_flight_total') else ''})",
                  len(table.rows) == expected_rows, len(table.rows))

            bottom = assembly.table_bottom(slide)
            floor = assembly._content_floor(slide, shape)
            clear = (floor - bottom) / 914400
            check(f"table clears the Included-with-Campaign block by {clear:+.2f}in (no overlap)",
                  bottom <= floor, clear)

            # Plan totals matching the form -- the deck's own Full Flight
            # Total (or, for a single-month flight, the Monthly Totals row,
            # which IS the full-flight total when n_months == 1) against the
            # document's published ground truth. Checked against the text
            # captured from assembly.personalize's real fill_data (what
            # actually got written into the table) AND against the ground
            # truth directly -- never against a sum this test computed.
            full_flight = payload.get("full_flight_total")
            shown = full_flight["impressions"] if full_flight else payload["total_impressions"]
            want = f"{scn['ground_truth']:,}"
            check(f"the deck's own total ({shown}) is the document's ground truth ({want})",
                  shown == want, shown)
            table_text = "\n".join(cell.text for row in table.rows for cell in row.cells)
            check("that figure is actually written on the rendered table",
                  want in table_text, want[:40])

    # The targeting map (roadmap §E): present exactly when at least one
    # group in this scenario carries resolved zips, which every one of
    # these three real documents does.
    by_id_all = {s.slide_id: i for i, s in enumerate(slides)}
    avails_slide_id = captured.get("avails_slide_id")
    check("place_targeting_map ran on the avails slide", avails_slide_id is not None)
    map_png = captured.get("map_png")
    check("a map was drawn for this scenario (it has resolved groups)", bool(map_png))
    if avails_slide_id in by_id_all:
        avails_idx = by_id_all[avails_slide_id]
        avails_pics = [s for s in slides[avails_idx].shapes if s.shape_type == 13]
        check("the map picture landed on the avails slide (3 pictures: background, "
              "PREMION wordmark, map)",
              len(avails_pics) == 3, [p.name for p in avails_pics])

    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        images, saved = deck_render.render_presentation(prs, out_dir, name=scn["name"].replace(" ", "_"))
    except deck_render.RenderUnavailable as exc:
        check("PowerPoint can open the generated deck", False, str(exc))
        return
    check("PowerPoint can open the generated deck", True)
    problems = package_check.check_package(str(saved))
    check("no unresolved relationship parts in the .pptx", not problems, problems[:5])
    print(f"    {len(images)} slide(s) -> {out_dir}")
    print(f"    deck: {saved}")

    # The clearance check above compares DECLARED geometry -- it cannot see
    # a row PowerPoint grows past what the sizer predicted, which is exactly
    # the failure mode this project already built a dedicated validator for
    # (tests/validate_text_metrics.py: "the real signal is a row reported
    # TALLER than declared"). Reused directly here rather than re-measured,
    # since a second, independent implementation of the same comparison is
    # exactly the kind of drift this suite exists to avoid elsewhere.
    import validate_text_metrics as vtm
    rows_checked, grown = vtm.check_deck(scn["name"], saved)
    check(f"no table row was rendered taller than the sizer declared "
          f"({rows_checked} sized rows checked against PowerPoint itself)",
          grown == 0, f"{grown} grew -- see the report printed above")

    # `vtm.check_deck` above only checks tables `condense_media_plan_table`
    # actually sized (its `sized_table` fingerprint is cell margins zeroed
    # by that pass) -- the AVAILS table goes through its own sizing pass,
    # `condense_avails_table`, so it gets its own, separate, COM-based
    # measurement here rather than being folded into that check. This is
    # §E's "confirm the avails table still clears its floor at every row
    # count with the map present" requirement. It used to fail here --
    # Wilmington's real audience/geo rows, long enough for the Geo column to
    # wrap, rendered the table to 12.92in on a 7.5in slide, because the
    # avails table had no row-height/font-shrink pass at all, unlike the
    # media plan table. `condense_avails_table` (assembly.py) is that pass.
    if avails_slide_id in by_id_all:
        avails_idx = by_id_all[avails_slide_id] + 1
        avails_table = assembly._find_table_shape(slides[by_id_all[avails_slide_id]])
        geo = vtm.powerpoint_geometry(str(saved))
        matches = [e for sig, e in geo.items() if sig[0] == avails_idx]
        slide_height_in = prs.slide_height / 914400
        if matches and avails_table is not None:
            rendered = matches[0][0]
            real_bottom_in = (rendered["top"] + sum(rendered["rows"])) / 72
            declared_bottom_in = (avails_table.top + avails_table.height) / 914400
            check(f"avails table's real rendered bottom "
                  f"({real_bottom_in:.2f}in, declared {declared_bottom_in:.2f}in) stays within "
                  f"the {slide_height_in:.1f}in slide -- condense_avails_table sized it",
                  real_bottom_in <= slide_height_in, (real_bottom_in, slide_height_in))
    if not keep:
        try:
            saved.unlink()
        except OSError:
            pass


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("only", nargs="?", help="run only scenarios whose name starts with this")
    parser.add_argument("--render", action="store_true", help="also Generate, open in PowerPoint, and inspect the deck")
    parser.add_argument("--keep", action="store_true", help="leave the rendered .pptx on disk")
    args = parser.parse_args()

    names = [n for n in gsf.SCENARIOS if not args.only or n.startswith(args.only)]
    if not names:
        print(f"No scenario matches {args.only!r}. Known: {', '.join(gsf.SCENARIOS)}")
        return 2

    if args.render:
        import deck_render
        out_root = deck_render.render_root()
        print(f"Rendered images will land under {out_root} -- open them and look.")

    for name in names:
        scn = gsf.SCENARIOS[name]()
        state = gsf.build_session_state(scn)
        check_state(scn, state)
        if args.render:
            import deck_render
            check_deck(scn, state, deck_render.render_root() / name, args.keep)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Row count, audience-major order, no collapse, resolved geography, avails totals "
          "and (for Annapolis) the gross markup all match the real documents.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
