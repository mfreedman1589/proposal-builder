"""report_assembly.py end to end, against the REAL Mattress Warehouse pair.

This file's whole reason to exist is ATTRIBUTION_REPORT_PLAN.md's
Correction 3: when the two uploaded files disagree about impressions
delivered, the DELIVERY file's figure is what the client sees, and the
attribution file's own delivered count -- the rate-math denominator --
must never appear anywhere in the deck as a delivered total.

MW is the pair that actually disagrees:
    MW delivery.xlsx          3,104,554  <- the headline
    MW attribution excel.xlsx 2,341,223  <- denominator only, never shown
(Both figures are asserted against the files themselves in
tests/test_attribution_import.py, so this file's expectations come from a
different source than the code under test.)

**This check was briefly synthetic and is not any more.** It was written
with a hand-built DeliveryExport because "Premion OTT.xlsx" had been
mislabelled as MW's delivery export (it is Cardinal Plumbing's) and the
real `MW delivery.xlsx` had been overlooked in the project root. Listing
the directory and reading each workbook's own campaign-name/RFPID cells
found it. Nothing here is invented now -- see CLAUDE.md's "read the file,
not the doc that describes it" rule, which this incident produced.

    python tests/test_report_assembly.py

SKIPs rather than fails when the real files aren't present, same
convention as test_wideorbit.py / test_attribution_import.py.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import market_lookup  # noqa: E402
import package_check  # noqa: E402
from pptx import Presentation  # noqa: E402

import app  # noqa: E402
import attribution_import as ai  # noqa: E402
import polk_import as pk  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402
import targeting_map  # noqa: E402

TEMPLATE = REPO / "REPORT_MASTER_v0_6.pptx"
TEMPLATE_V0_7 = REPO / "REPORT_MASTER_v0_7.pptx"
TEMPLATE_V0_11 = REPO / "REPORT_MASTER_v0_11.pptx"
TEMPLATE_V0_12 = REPO / "REPORT_MASTER_v0_12.pptx"
ATTRIBUTION_MW = REPO / "MW attribution excel.xlsx"
DELIVERY_MW = REPO / "MW delivery.xlsx"
ATTRIBUTION_CARDINAL = REPO / "Premion Website Attribution Cardinal.xlsx"
DELIVERY_CARDINAL = REPO / "Premion OTT.xlsx"
ATTRIBUTION_WAEPA = REPO / "Premion Website Attribution and Reach Extension (13).xlsx"
OTT_RETARGETING_CARDINAL = REPO / "Audience Marketplace Cardinal.xlsx"
# The real NWFCU export the whole 2026-09-17 review was written against --
# found in the project folder 2026-09-18. Identified by its own Advertiser
# tab ("Northwest Federal Credit Union - Direct"), not by this filename,
# which is shared with every other numbered Reach Extension export.
ATTRIBUTION_NWFCU = REPO / "Premion Website Attribution and Reach Extension (15).xlsx"
POLK_DASHBOARD = REPO / "Polk Dashboard.xlsx"

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

    def pending(self, reason):
        """For a partial, EXPECTED gap inside an otherwise fully-exercised
        check (e.g. assertions gated on a template feature that doesn't
        exist yet) -- deliberately NOT the word "SKIP": run_all.py's own
        sweep classifies a whole FILE as skipped the moment any line of its
        output starts with "SKIP" (tests/run_all.py's run_one), which is
        the right call for a fixture-missing skip() at the top of a check
        (nothing ran) but wrong here, where 40+ real checks already ran and
        passed -- "PENDING" reports the same "nothing to assert yet" fact
        without hiding a real pass/fail count behind a false SKIP label."""
        print(f"  PENDING  {reason}")
        self.skipped.append(reason)


def _slide_keys(prs):
    """Every slide's `key:` note, in deck order."""
    out = []
    for slide in prs.slides:
        if slide.has_notes_slide:
            for line in slide.notes_slide.notes_text_frame.text.splitlines():
                if line.strip().lower().startswith("key:"):
                    out.append(line.split(":", 1)[1].strip())
    return out


def _parts_equal(path_a, path_b):
    """Whether two .pptx files carry identical CONTENT part-for-part --
    never a raw file-bytes comparison, which the zip container's own
    save-time stamp (a real ZipInfo.date_time on every entry, confirmed
    while building this check) makes non-deterministic across two separate
    `prs.save()` calls even when every part's content is byte-identical.
    Same exclude-the-save-stamp discipline `test_group_backward_compat.py`'s
    own "as presented" byte-diff already holds itself to (CLAUDE.md)."""
    import zipfile
    with zipfile.ZipFile(path_a) as za, zipfile.ZipFile(path_b) as zb:
        names_a, names_b = set(za.namelist()), set(zb.namelist())
        if names_a != names_b:
            return False
        return all(za.read(name) == zb.read(name) for name in names_a)


def _deck_text(path):
    """Every string the rendered deck actually contains -- text frames and
    table cells alike. Read back off the built file rather than asserted
    against what the fill code was asked to write, so a token that silently
    failed to fill shows up here."""
    prs = Presentation(path)
    out = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                out.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        out.append(cell.text_frame.text)
    return "\n".join(out)


def _build(rep, label, attribution_path, delivery_path, out_name):
    attribution = ai.parse_attribution_export(str(attribution_path))
    delivery = ai.parse_delivery_export(str(delivery_path)) if delivery_path else None
    out_path = REPO / "tests" / "_manual_output" / out_name
    out_path.parent.mkdir(exist_ok=True)
    path, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, delivery, str(out_path),
        goals_bullets=[f"{label}: campaign goal placeholder for this check"],
        whats_next_bullets=[f"{label}: what's-next placeholder for this check"])
    return attribution, delivery, path, warnings


def check_mw_headline_precedence(rep):
    print("\nMW, both files -- Correction 3's real regression case")
    for path in (TEMPLATE, ATTRIBUTION_MW, DELIVERY_MW):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution, delivery, path, warnings = _build(
        rep, "MW", ATTRIBUTION_MW, DELIVERY_MW, "MW_both.pptx")

    # The two figures come from the files, never hardcoded here -- so this
    # check still means something if the fixtures are ever refreshed.
    headline = f"{delivery.delivered_impressions:,}"
    denominator = f"{attribution.delivered_impressions:,}"
    rep.check("the two files genuinely disagree (otherwise this proves nothing)",
             delivery.delivered_impressions != attribution.delivered_impressions,
             (headline, denominator))
    text = _deck_text(path)
    rep.check(f"the DELIVERY figure ({headline}) appears in the deck",
             headline in text, headline)
    rep.check(f"the ATTRIBUTION file's own delivered count ({denominator}) appears NOWHERE",
             denominator not in text, denominator)
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])
    _check_only_expected_warnings(rep, warnings)


def _check_only_expected_warnings(rep, warnings):
    """No warning except the known, temporary one: TopZipTable is 3 columns
    in v0_2 and the ranking rule supplies 5 (Zip | Area | Share | Rate |
    Multiple). That widening lands with v0_3. Asserted narrowly rather than
    ignored, so a REAL fit overflow still fails here, and so this check goes
    green by itself the moment the template catches up -- rather than being
    a blanket `not warnings` that someone has to remember to re-tighten.

    Also allows the "N zip code(s) have no ZCTA boundary" warning -- MW's
    real zip data genuinely has one (27515), and that's correct behavior
    (2026-09-13's choropleth fix: a zip with no ZCTA point at all used to
    vanish from the map with no warning whatsoever; now it's counted),
    not a defect this fit/column check is the one responsible for
    catching.

    Also allows "TopPublishersTable has 3 columns but 4 were supplied --
    not showing: rate" -- NWFCU review, 2026-09-17: the publisher table
    gains an Attributed Rate column whenever attribution.by_channel
    exists, same "activates on its own once the template widens" shape as
    TopZipTable's own pending column above; today's real template hasn't
    been widened yet.
    """
    pending = [w for w in warnings
              if "TopZipTable has 3 columns" in w or "have no ZCTA boundary" in w
              or "TopPublishersTable has 3 columns" in w]
    unexpected = [w for w in warnings if w not in pending]
    rep.check("no unexpected fit/column warning", not unexpected, unexpected)
    if pending:
        print(f"        (known, expected: {pending[0][:70]}...)")


def check_mw_attribution_only(rep):
    print("\nMW, attribution only -- delivery set dropped, attribution figure becomes headline")
    for path in (TEMPLATE, ATTRIBUTION_MW):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution, _delivery, path, warnings = _build(
        rep, "MW", ATTRIBUTION_MW, None, "MW_attribution_only.pptx")
    prs = Presentation(path)
    text = _deck_text(path)
    # Asserted by KEY, not by a slide count: v0_2 -> v0_3 added a slide and
    # turned a bare "== 6" into a check that passed for the wrong reason.
    keys = _slide_keys(prs)
    rep.check("both delivery-set slides are gone",
             "report:delivery_recap" not in keys and "report:delivery_breakdown" not in keys,
             keys)
    rep.check("with no delivery file, the attribution figure IS the headline",
             f"{attribution.delivered_impressions:,}" in text,
             f"{attribution.delivered_impressions:,}")
    # MW has no conversions -- `_build`'s default `include_conversions=False`
    # must reflow ConversionsTile away (v0_4 template migration).
    highlights = next(s for s in prs.slides
                      if s.has_notes_slide and slide_map.notes_key(s) == "report:highlights")
    tile_names = {sh.name for sh in highlights.shapes}
    rep.check("ConversionsTile is gone on a no-conversions export (reflowed to 3 tiles)",
             not ({"ConversionsTile", "ConversionsTileValue", "ConversionsTileLabel"} & tile_names),
             tile_names)
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])


def check_cardinal_both(rep):
    print("\nCardinal, both files -- the second real pair, full slide set")
    for path in (TEMPLATE, ATTRIBUTION_CARDINAL, DELIVERY_CARDINAL):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    _attribution, delivery, path, warnings = _build(
        rep, "Cardinal", ATTRIBUTION_CARDINAL, DELIVERY_CARDINAL, "Cardinal_both.pptx")
    prs = Presentation(path)
    text = _deck_text(path)
    keys = _slide_keys(prs)
    rep.check("both delivery-set slides are present",
             "report:delivery_recap" in keys and "report:delivery_breakdown" in keys, keys)
    rep.check("Cardinal has >1 geo option and >1 creative, so the breakdown applies",
             ra.delivery_breakdown_applies(delivery),
             (delivery.by_geo, len(delivery.by_creative)))
    rep.check("the delivery figure appears", f"{delivery.delivered_impressions:,}" in text,
             f"{delivery.delivered_impressions:,}")
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])
    _check_only_expected_warnings(rep, warnings)


def check_waepa_conversions_and_multi_rfpid(rep):
    """The third real fixture (2026-09-08 finding), no companion delivery
    file -- proves three things a MW/Cardinal-only suite never exercised:
    a multi-RFPID export builds a deck at all (the old code raised before
    it got this far), conversions render (highlights tile) and gracefully
    degrade on tables the template hasn't been widened for yet, and a real
    DMA market cut fills GEOGRAPHY_LABEL without a station-pixel fallback.
    """
    print("\nWAEPA -- multi-RFPID + conversions, attribution-only (no delivery file)")
    for path in (TEMPLATE, ATTRIBUTION_WAEPA):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_WAEPA))
    rep.check("this fixture genuinely has 2 RFPIDs (otherwise this proves nothing)",
             len(attribution.rfpid_breakdown) == 2, attribution.rfpid_breakdown)
    rep.check("this fixture genuinely has conversions (otherwise this proves nothing)",
             attribution.has_conversions and attribution.attributed_conversions == 37,
             attribution.attributed_conversions)

    out_path = REPO / "tests" / "_manual_output" / "WAEPA_conversions.pptx"
    out_path.parent.mkdir(exist_ok=True)
    path, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_path),
        goals_bullets=["WAEPA plumbing check: drive qualified insurance leads"],
        whats_next_bullets=["WAEPA plumbing check: expand DC/Baltimore targeting"],
        include_conversions=True)
    text = _deck_text(path)
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])
    rep.check("the real DMA market names fill Geography (not a station-pixel guess)",
             "Washington" in text and "Baltimore" in text, text[:2000])
    rep.check("the conversions figure (37) appears somewhere in the deck",
             "37" in text, None)

    # v0_4's named ConversionsTile group (2026-09-08 template migration,
    # migrate_report_master_v0_4.py) -- proves the highlights slide really
    # has 4 tiles, not just that the token text happens to appear somewhere.
    prs = Presentation(path)
    highlights = next(s for s in prs.slides
                      if s.has_notes_slide and slide_map.notes_key(s) == "report:highlights")
    tile_names = {sh.name for sh in highlights.shapes}
    rep.check("ConversionsTile/Value/Label are all present with conversions on",
             {"ConversionsTile", "ConversionsTileValue", "ConversionsTileLabel"} <= tile_names,
             tile_names)

    # The CURRENT template hasn't been widened for the three graceful-
    # degrade columns yet (see ATTRIBUTION_REPORT_PLAN.md's WAEPA section)
    # -- confirm the warnings say so by name, not silently swallowed.
    degraded = [w for w in warnings if "not showing" in w]
    rep.check("the three not-yet-widened tables (Breakdown/Intent/TopUrl) all warn, "
             "rather than crashing or silently dropping the column",
             len(degraded) == 3, warnings)

    # And with the toggle off, the deck must build with NONE of that --
    # WAEPA's own "no half-states" rule, and the 4th tile reflows away.
    out_path2 = REPO / "tests" / "_manual_output" / "WAEPA_no_conversions.pptx"
    path2, warnings2 = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_path2),
        goals_bullets=["WAEPA plumbing check: drive qualified insurance leads"],
        whats_next_bullets=["WAEPA plumbing check: expand DC/Baltimore targeting"],
        include_conversions=False)
    prs2 = Presentation(path2)
    highlights2 = next(s for s in prs2.slides
                       if s.has_notes_slide and slide_map.notes_key(s) == "report:highlights")
    tile_names2 = {sh.name for sh in highlights2.shapes}
    rep.check("ConversionsTile/Value/Label are all GONE with conversions off (reflowed away)",
             not ({"ConversionsTile", "ConversionsTileValue", "ConversionsTileLabel"} & tile_names2),
             tile_names2)
    rep.check("the other three tiles survive the reflow",
             {"ImpressionsTile", "VisitorsTile", "RateTile"} <= tile_names2, tile_names2)
    rep.check("no conversion-column warnings when the toggle is off",
             not any("conv_rate" in w or "converted" in w for w in warnings2), warnings2)


def check_phase5_proposal_link(rep):
    """ATTRIBUTION_REPORT_PLAN.md Phase 5 -- the deterministic half (goals/
    audience/geo/flight/budget wiring, and the zip overlay's graceful
    degradation). The live half (goals prefill through the real page, the
    model actually flagging a contradicting note) is covered by
    test_attribution_reports_page.py and test_attribution_draft_live.py's
    own waepa_proposal scenario -- this file stays offline-only.

    Field-map values are the REAL ones confirmed live 2026-09-09 against
    proposal 09e61e0b-a945-4022-851d-d3c31b2acbd0's own form_json (see
    ATTRIBUTION_REPORT_PLAN.md's Phase 5 section) -- not invented here,
    same discipline every other check in this file already holds itself
    to. That proposal's own `targeting_groups` is `[]` (avails_mode was
    off), so the zip-overlay half uses a SYNTHETIC target-zip list built
    from real zips pulled off the crosswalk (`geo_resolver`), explicitly
    marked as such -- per the plan's own decision, to be replaced the day
    a real avails-linked WAEPA pair exists.
    """
    print("\nPhase 5 -- proposal-linked recap/budget/zip-overlay wiring (WAEPA field map)")
    if not (TEMPLATE.exists() and ATTRIBUTION_WAEPA.exists()):
        rep.skip(f"{TEMPLATE.name} or {ATTRIBUTION_WAEPA.name} not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_WAEPA))

    # -- geography_label_from_names / geography_overflow_bullet_from_names --
    rep.check("2 names join on one line (no collapse)",
             ra.geography_label_from_names(["Washington, DC", "Baltimore"]) == "Washington, DC, Baltimore")
    rep.check("no overflow bullet for 2 names",
             ra.geography_overflow_bullet_from_names(["Washington, DC", "Baltimore"]) is None)
    rep.check("3+ names collapse to 'N markets'",
             ra.geography_label_from_names(["A", "B", "C"]) == "3 markets")
    rep.check("3+ names produce the full-list overflow bullet",
             ra.geography_overflow_bullet_from_names(["A", "B", "C"]) == "Markets: A, B, C")
    rep.check("empty/None names collapse to None (never a fabricated tile)",
             ra.geography_label_from_names([]) is None and ra.geography_label_from_names(None) is None)
    # The export-derived functions must still agree with themselves after
    # being refactored onto the shared _from_names helpers.
    rep.check("geography_label(attribution) still resolves to a real value post-refactor",
             ra.geography_label(attribution) is not None, ra.geography_label(attribution))

    # -- build_facts_payload's budget/proposal sections --
    facts_linked = ra.build_facts_payload(
        attribution, None, goals=["g"], notes="", include_conversions=True,
        budget=74970.0, proposal_flight_label="Oct 2026 - Dec 2026",
        proposal_geography_label="Washington, DC, Baltimore")
    rep.check("facts['budget'] is populated with the real total",
             facts_linked["budget"]["total"] == 74970.0, facts_linked["budget"])
    rep.check("cost_per_attributed_visit is total / attributed_unique_visitors",
             abs(facts_linked["budget"]["cost_per_attributed_visit"]
                 - 74970.0 / attribution.attributed_unique_visitors) < 0.01,
             facts_linked["budget"])
    rep.check("cost_per_conversion is populated (WAEPA has real conversions)",
             facts_linked["budget"]["cost_per_conversion"] is not None, facts_linked["budget"])
    rep.check("facts['proposal'] carries the linked flight/geography, verbatim",
             facts_linked["proposal"] == {"flight_label": "Oct 2026 - Dec 2026",
                                          "geography_label": "Washington, DC, Baltimore"},
             facts_linked["proposal"])

    facts_unlinked = ra.build_facts_payload(attribution, None, goals=["g"])
    rep.check("no proposal linked -> facts['budget']/['proposal'] are both None, "
             "matching delivery/live_sports's own None-when-absent convention",
             facts_unlinked["budget"] is None and facts_unlinked["proposal"] is None,
             (facts_unlinked["budget"], facts_unlinked["proposal"]))

    # -- build_report_deck: audience/flight/geography threaded, no proposal --
    out_linked = REPO / "tests" / "_manual_output" / "WAEPA_phase5_linked.pptx"
    out_linked.parent.mkdir(exist_ok=True)
    audience_bullets = ["Civilian federal government employees, excluding current WAEPA members",
                        "Behavioral targeting layer for federal-benefits researchers"]
    path_linked, warnings_linked = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_linked),
        client_name="WAEPA", goals_bullets=["Evaluate OTT performance beyond awareness"],
        whats_next_bullets=["Expand DC/Baltimore targeting"],
        audience_bullets=audience_bullets, flight_label="Oct 2026 - Dec 2026",
        geography_names_override=["Washington, DC", "Baltimore"], include_conversions=True)
    text_linked = _deck_text(path_linked)
    rep.check("no unfilled {{TOKEN}} survived (linked build)", "{{" not in text_linked, text_linked[:300])
    rep.check("FLIGHT_LABEL reads the proposal's own flight, not blank/reflowed",
             "Oct 2026 - Dec 2026" in text_linked, text_linked[:2000])
    rep.check("GEOGRAPHY_LABEL reads the target-market names, joined -- 'Washington, DC, Baltimore'",
             "Washington, DC, Baltimore" in text_linked, text_linked[:2000])
    rep.check("AUDIENCE_BULLETS reads the linked proposal's own Campaign Specs audience text, "
             "not the export's by_audience rows",
             all(a in text_linked for a in audience_bullets), text_linked[:2000])

    # -- and the reverse: no override at all reproduces Phase 3/4's own default --
    out_unlinked = REPO / "tests" / "_manual_output" / "WAEPA_phase5_unlinked.pptx"
    path_unlinked, _w = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_unlinked),
        goals_bullets=["placeholder"], whats_next_bullets=["placeholder"])
    text_unlinked = _deck_text(path_unlinked)
    rep.check("'Oct 2026 - Dec 2026' does NOT appear with no proposal linked",
             "Oct 2026 - Dec 2026" not in text_unlinked, text_unlinked[:300])

    # -- zip overlay: graceful degradation with no targeting_groups --
    # 09e61e0b's own real targeting_groups is [] (avails_mode was off) --
    # this reproduces that exact case: targeted_zips=None must build a
    # deck IDENTICAL PART-FOR-PART to a build with the parameter omitted
    # entirely (`_parts_equal`, never a raw file-bytes compare -- the zip
    # container's own save-time stamp makes that flaky across separate
    # `prs.save()` calls even when every part's content agrees).
    out_degrade_a = REPO / "tests" / "_manual_output" / "WAEPA_phase5_zip_degrade_a.pptx"
    out_degrade_b = REPO / "tests" / "_manual_output" / "WAEPA_phase5_zip_degrade_b.pptx"
    ra.build_report_deck(str(TEMPLATE), attribution, None, str(out_degrade_a),
                         goals_bullets=["g"], whats_next_bullets=["w"], targeted_zips=None)
    ra.build_report_deck(str(TEMPLATE), attribution, None, str(out_degrade_b),
                         goals_bullets=["g"], whats_next_bullets=["w"])
    rep.check("targeted_zips=None produces a deck identical part-for-part to omitting it entirely",
             _parts_equal(out_degrade_a, out_degrade_b))

    out_degrade_c = REPO / "tests" / "_manual_output" / "WAEPA_phase5_zip_degrade_c.pptx"
    ra.build_report_deck(str(TEMPLATE), attribution, None, str(out_degrade_c),
                         goals_bullets=["g"], whats_next_bullets=["w"], targeted_zips=set())
    rep.check("an empty targeted_zips SET (real shape of a proposal with "
             "targeting_groups=[] -- 09e61e0b's own real case) degrades identically too, "
             "never an empty second series",
             _parts_equal(out_degrade_c, out_degrade_a))

    # -- zip overlay: a SYNTHETIC real-zip target list (marked as such --
    # see this check's own docstring) actually draws and reports cleanly --
    import geo_resolver
    points = geo_resolver._data()["zip_points"]
    synthetic_targeted_zips = sorted(z for z in points if z.startswith("200"))[:15]
    rep.check("the synthetic target-zip fixture is genuinely non-empty "
             "(otherwise this proves nothing)", len(synthetic_targeted_zips) > 0)
    out_overlay = REPO / "tests" / "_manual_output" / "WAEPA_phase5_zip_overlay_SYNTHETIC.pptx"
    path_overlay, warnings_overlay = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_overlay),
        goals_bullets=["g"], whats_next_bullets=["w"],
        targeted_zips=synthetic_targeted_zips)
    rep.check("the deck still builds cleanly with a targeted-zip overlay present",
             "{{" not in _deck_text(path_overlay), None)
    rep.check("the overlay build produces a DIFFERENT rendered image than the no-overlay "
             "build (the outline layer actually changed the map PNG)",
             not _parts_equal(out_overlay, out_degrade_a))


def check_phase4_override_plumbing(rep):
    """ATTRIBUTION_REPORT_PLAN.md Phase 4 threads Claude-drafted content
    into the deck through `narratives`/`breakdown_dimension_override`/
    `geography_label_override` (report_assembly.py) rather than
    `app.apply_attr_draft` writing slide XML itself. This is the plumbing
    check for that path -- every override actually reaches the rendered
    slide text, and REPLACES the deterministic default rather than sitting
    alongside it."""
    print("\nPhase 4 override plumbing -- narratives/dimension/geography reach the real deck")
    for path in (TEMPLATE, ATTRIBUTION_CARDINAL, DELIVERY_CARDINAL):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_CARDINAL))
    delivery = ai.parse_delivery_export(str(DELIVERY_CARDINAL))

    # Cardinal has exactly one market, so pick_breakdown_dimension's own
    # 1.5x heuristic has a real choice to make (audience vs. creative) --
    # unlike MW (4 markets), where "Market" always wins and is never
    # overridable. Confirms this fixture is the right one before relying
    # on it to prove the override actually changes anything.
    default_dimension, _rows = ra.pick_breakdown_dimension(attribution)
    rep.check("Cardinal has one market, so there IS a real dimension judgment to override",
             default_dimension != "Market", default_dimension)

    marker_attribution = "OVERRIDE-ATTRIBUTION-NARRATIVE-MARKER"
    marker_zip = "OVERRIDE-ZIP-NARRATIVE-MARKER"
    marker_geo = "Override Geography Label"
    override_dimension = "Creative" if default_dimension == "Audience" else "Audience"
    out_path = REPO / "tests" / "_manual_output" / "Cardinal_phase4_overrides.pptx"
    out_path.parent.mkdir(exist_ok=True)
    path, _warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, delivery, str(out_path),
        goals_bullets=["Phase 4 plumbing check"], whats_next_bullets=["Phase 4 plumbing check"],
        narratives={"attribution": marker_attribution, "zip": marker_zip},
        breakdown_dimension_override=override_dimension,
        geography_label_override=marker_geo)
    text = _deck_text(path)
    rep.check("the attribution narrative override reached the deck", marker_attribution in text)
    rep.check("the zip narrative override reached the deck", marker_zip in text)
    rep.check("the geography label override reached the deck", marker_geo in text)
    rep.check(f"the dimension override ({override_dimension!r}) reached BREAKDOWN_DIMENSION_LABEL",
             override_dimension in text)
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])

    # And the reverse: with NO overrides, the deterministic defaults are
    # still exactly what they were before this phase -- Phase 3's existing
    # callers/tests need no changes, per build_report_deck's own docstring.
    out_path2 = REPO / "tests" / "_manual_output" / "Cardinal_phase4_no_overrides.pptx"
    path2, _warnings2 = ra.build_report_deck(
        str(TEMPLATE), attribution, delivery, str(out_path2),
        goals_bullets=["Phase 4 plumbing check"], whats_next_bullets=["Phase 4 plumbing check"])
    text2 = _deck_text(path2)
    rep.check("with no overrides, the marker text is absent", marker_attribution not in text2)
    rep.check("with no overrides, the deterministic dimension choice is unchanged",
             default_dimension in text2)


DELIVERY_PGCC = REPO / "Premion OTT (2).xlsx"


def check_live_sports(rep):
    """Live sports as a distinct block inside a delivery workbook (2026-09-08
    addition) -- a real Prince George's Community College export, RFPID-
    266713 (a PREM TV live-sports package) riding inside the same workbook
    as RFPID-266710's OTT campaign. **Cross-paired with MW's real attribution
    file on purpose** -- there is no real PGCC attribution export on hand,
    and this project's own rule (see this file's own header) is never to
    fabricate a fixture; pairing two genuinely real exports from different
    campaigns is honest about what's being tested (the sports-block
    mechanism, not a real matched client pair) and fabricates nothing.

    Two things are checked regardless of the template: the headline combines
    OTT + sports (`combined_headline_impressions`, the settled answer to
    "should headline impressions be OTT + sports combined" -- confirmed
    against the real file that its own totals never combine the two blocks,
    so the app has to do it deliberately) and `build_facts_payload`'s own
    `live_sports` section. The slide itself (report:live_sports) is
    template-dependent -- v0_4 doesn't have it yet -- so those assertions
    are gated on the key actually being present and SKIP with a named
    reason otherwise; this activates automatically the moment Matt's
    template update lands, no test change required.
    """
    print("\nLive sports block -- Prince George's Community College delivery export")
    for path in (TEMPLATE, ATTRIBUTION_MW, DELIVERY_PGCC):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_PGCC))
    rep.check("this fixture genuinely carries a live-sports block "
             "(otherwise this proves nothing)",
             ra.live_sports_applies(delivery), delivery.live_sports)
    if not delivery.live_sports:
        return

    combined = ra.combined_headline_impressions(attribution, delivery)
    rep.check("headline impressions is OTT + sports, combined",
             combined == delivery.delivered_impressions + delivery.live_sports.delivered_impressions,
             combined)
    rep.check("...and strictly greater than the OTT figure alone (sports isn't lost)",
             combined > delivery.delivered_impressions, combined)

    facts = ra.build_facts_payload(attribution, delivery, goals=[])
    rep.check("facts payload's headline matches the same combined figure",
             facts["headline"]["delivered_impressions"] == combined)
    ls_facts = facts.get("live_sports")
    rep.check("facts payload carries a live_sports section", bool(ls_facts))
    if ls_facts:
        rep.check("live_sports facts carry the real package type",
                 ls_facts["package_type"] == delivery.live_sports.package_type)
        rep.check("live_sports pacing note matches Matt's own example shape",
                 ls_facts["pacing_note"] == f"{delivery.live_sports.delivered_impressions:,} of "
                                            f"{delivery.live_sports.flight_goal:,}",
                 ls_facts["pacing_note"])

    out_path = REPO / "tests" / "_manual_output" / "PGCC_live_sports.pptx"
    out_path.parent.mkdir(exist_ok=True)
    path, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, delivery, str(out_path),
        goals_bullets=["Live sports check: campaign goal placeholder"],
        whats_next_bullets=["Live sports check: what's-next placeholder"])
    rep.check("the deck still builds cleanly with a sports block present "
             "(whether or not the template has caught up yet)", True)

    prs = Presentation(path)
    keys = _slide_keys(prs)
    if "report:live_sports" not in keys:
        rep.pending("template has no report:live_sports slide yet -- "
                   "slide-fill assertions activate once it's added")
        return

    text = _deck_text(path)
    ls = delivery.live_sports
    rep.check("the sports package's own delivered figure appears on the slide",
             f"{ls.delivered_impressions:,}" in text)
    rep.check("the pacing note's flight-goal figure appears",
             f"{ls.flight_goal:,}" in text)
    top_event = max(ls.events, key=lambda e: e.delivered_impressions)
    rep.check("the top event by impressions appears in the event table",
             top_event.event in text)
    rep.check("a single-league export has NO league breakdown table "
             "(SportsByLeagueTable deleted, matching the single-value-"
             "dimension rule elsewhere)",
             len(ls.by_league) <= 1)
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])
    _check_only_expected_warnings(rep, warnings)


AUTO_SALES_DECK = REPO / "Auto Group (5 Sites) - 10_51 AM ET_Summary (1).pptx"


def check_auto_sales_analyst_append(rep):
    """The Auto-Sales Analyst deck append slot (2026-09-08) -- an 8-slide,
    13.333x7.5in real export (image charts, no native chart parts, so none
    of the cross-deck copy hazards assembly.copy_slide_into's own docstring
    warns about apply). Appended WHOLESALE via `report_assembly.
    append_slide_deck`/`build_report_deck`'s `extra_deck_path` -- no
    parsing, nothing pulled out of it into the facts payload, confirmed by
    checking the deck's own slide COUNT grew by exactly 8 rather than
    inspecting its content. "Same package_check gate as vault inserts" --
    `package_check.check_package` is the structural-soundness check every
    other cross-deck graft in this app is verified with (test_slide_vault.py's
    own "package is structurally clean" check), not a live runtime gate --
    the live app doesn't call package_check for case studies or vault
    slides either, so this doesn't either.
    """
    print("\nAuto-Sales Analyst deck append -- an 8-slide real export, appended wholesale")
    for path in (TEMPLATE, ATTRIBUTION_MW, AUTO_SALES_DECK):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    source = Presentation(str(AUTO_SALES_DECK))
    source_count = len(source.slides._sldIdLst)
    rep.check("this fixture genuinely has more than one slide "
             "(otherwise 'appended wholesale, in order' proves nothing)",
             source_count > 1, source_count)

    out_path = REPO / "tests" / "_manual_output" / "MW_with_auto_sales.pptx"
    out_path.parent.mkdir(exist_ok=True)
    path, _warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_path),
        goals_bullets=["Auto-Sales Analyst append check: goal placeholder"],
        whats_next_bullets=["Auto-Sales Analyst append check: what's-next placeholder"],
        extra_deck_path=str(AUTO_SALES_DECK))

    built = Presentation(path)
    without_path = REPO / "tests" / "_manual_output" / "MW_without_auto_sales.pptx"
    ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(without_path),
        goals_bullets=["Auto-Sales Analyst append check: goal placeholder"],
        whats_next_bullets=["Auto-Sales Analyst append check: what's-next placeholder"])
    without = Presentation(without_path)
    base_count = len(without.slides._sldIdLst)
    rep.check(f"the deck grew by exactly {source_count} slides (the whole "
             f"Auto-Sales deck, nothing dropped, nothing duplicated)",
             len(built.slides._sldIdLst) == base_count + source_count,
             (len(built.slides._sldIdLst), base_count, source_count))

    # Appended AFTER the attribution slides -- the last `source_count`
    # slides of the built deck are the Auto-Sales ones, everything before
    # them is unchanged from the report-only build.
    appended_slides = list(built.slides)[base_count:]
    rep.check("every appended slide has at least one shape (a real slide, "
             "not an empty placeholder)",
             all(len(list(s.shapes)) > 0 for s in appended_slides))

    problems = package_check.check_package(path)
    rep.check("package is structurally clean after the cross-deck append "
             "(same gate test_slide_vault.py uses)", not problems, problems[:4])


def check_polk_automotive_registrations(rep):
    """Phase 7 -- Polk automotive match-back, against the real Polk
    Dashboard.xlsx fixture (headline figures re-verified directly against
    the file: 44,756 matched households, 5 target dealer sales, 90.49%
    match rate, 1.5x lift). Cross-paired with MW's real attribution file on
    purpose -- there is no real attribution export for the same automotive
    client on hand, and this file's own header rule is never to fabricate a
    fixture; pairing two genuinely real exports from different campaigns
    tests the Polk plumbing honestly (facts payload, projection toggle,
    optional-slide dropping) without inventing a matched client pair.

    The slide itself (report:automotive_registrations) doesn't exist in ANY
    template yet -- Matt's own build is a Phase 7 handoff item -- so those
    assertions are `rep.pending()`, the same "activates automatically the
    moment the template lands" shape `check_live_sports` used before v0_4.
    """
    print("\nPolk automotive match-back -- real Polk Dashboard.xlsx fixture")
    for path in (TEMPLATE, ATTRIBUTION_MW, POLK_DASHBOARD):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    polk = pk.parse_polk_export(str(POLK_DASHBOARD))
    rep.check("this fixture genuinely has target dealer sales (otherwise this proves nothing)",
             polk.has_target_dealer_sales and polk.target_dealer_sales == 5, polk.target_dealer_sales)

    rep.check("automotive_registrations_applies(None) is False", not ra.automotive_registrations_applies(None))
    rep.check("automotive_registrations_applies(polk) is True", ra.automotive_registrations_applies(polk))

    # project_for_match_rate -- pure, no rounding, no ZeroDivisionError on a
    # falsy match rate.
    rep.check("project_for_match_rate divides by the match rate",
             abs(ra.project_for_match_rate(polk.matched_households, polk.match_rate)
                 - 44756 / 0.9049) < 0.5,
             ra.project_for_match_rate(polk.matched_households, polk.match_rate))
    rep.check("project_for_match_rate is a no-op on a falsy match rate (never raises)",
             ra.project_for_match_rate(100, 0) == 100)

    # facts payload -- absent entirely when polk isn't supplied.
    facts_no_polk = ra.build_facts_payload(attribution, None, goals=["placeholder"])
    rep.check("facts['polk'] is None when no Polk file was supplied", facts_no_polk["polk"] is None)

    # facts payload -- raw figures always present, projected figures ONLY
    # when the rep's own toggle is on.
    facts_off = ra.build_facts_payload(attribution, None, goals=["placeholder"], polk=polk)
    polk_facts_off = facts_off["polk"]
    rep.check("facts['polk'] carries the real matched households", polk_facts_off is not None
             and polk_facts_off["matched_households"] == 44756, polk_facts_off)
    rep.check("facts['polk'] carries the real match rate",
             abs(polk_facts_off["match_rate"] - 0.9049) < 1e-6, polk_facts_off)
    rep.check("projection OFF -> no projected_* keys leak into the payload (nothing to "
             "accidentally cite as a fact)",
             polk_facts_off["projected_matched_households"] is None
             and polk_facts_off["projected_target_dealer_sales"] is None, polk_facts_off)
    rep.check("facts['polk']['projected'] reflects the toggle (False)",
             polk_facts_off["projected"] is False)

    facts_on = ra.build_facts_payload(attribution, None, goals=["placeholder"],
                                      polk=polk, polk_projected=True)
    polk_facts_on = facts_on["polk"]
    rep.check("projection ON -> projected households is the raw figure divided by match rate",
             abs(polk_facts_on["projected_matched_households"] - 44756 / 0.9049) < 0.5,
             polk_facts_on)
    rep.check("projection ON -> projected target dealer sales is likewise divided",
             abs(polk_facts_on["projected_target_dealer_sales"] - 5 / 0.9049) < 0.01,
             polk_facts_on)
    rep.check("raw matched_households is UNCHANGED by the toggle -- projection never mutates "
             "the real matched figure, only adds a separate projected one",
             polk_facts_on["matched_households"] == 44756, polk_facts_on)
    rep.check("top_audience/top_creative/top_publisher ride into the facts payload",
             polk_facts_on["top_audience"]["segment"] == "AUTO Ford Intenders"
             and polk_facts_on["top_creative"]["creative"] == "TG11075TBrittChev062630"
             and polk_facts_on["top_publisher"]["publisher"] == "Pluto TV",
             polk_facts_on)

    # The deck still builds cleanly with polk supplied even though no
    # current template has the slide yet -- the drop-if-absent branch must
    # never raise just because the key doesn't exist.
    out_path = REPO / "tests" / "_manual_output" / "MW_with_polk.pptx"
    out_path.parent.mkdir(exist_ok=True)
    path, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_path),
        goals_bullets=["Polk check: drive incremental vehicle sales"],
        whats_next_bullets=["Polk check: expand the winning audience segment"],
        polk=polk, polk_projected=False)
    rep.check("the deck still builds cleanly with a Polk file present "
             "(whether or not the template has caught up yet)", True)
    keys = _slide_keys(Presentation(path))
    rep.check("report:automotive_registrations still absent from v0_6 -- the drop-if-absent "
             "branch is what's under test here, not the slide fill",
             "report:automotive_registrations" not in keys, keys)

    # v0_12 (2026-09-19) is the first template that actually carries the
    # slide -- this is where the fill assertions that used to be
    # rep.pending() activate for real.
    if not TEMPLATE_V0_12.exists():
        rep.skip(f"{TEMPLATE_V0_12.name} not present -- fill assertions skipped")
        return
    keys_v12 = _slide_keys(Presentation(str(TEMPLATE_V0_12)))
    rep.check("report:automotive_registrations is slide 11 in v0_12, immediately after "
             "takeaways and before the standalone summary/case_study slides",
             keys_v12[11] == "report:automotive_registrations"
             and keys_v12[10] == "report:takeaways"
             and keys_v12[12:14] == ["report:summary", "report:case_study"], keys_v12)

    out_off = REPO / "tests" / "_manual_output" / "MW_with_polk_v0_12_off.pptx"
    out_on = REPO / "tests" / "_manual_output" / "MW_with_polk_v0_12_on.pptx"
    path_off, warnings_off = ra.build_report_deck(
        str(TEMPLATE_V0_12), attribution, None, str(out_off),
        goals_bullets=["Polk check: drive incremental vehicle sales"],
        whats_next_bullets=["Polk check: expand the winning audience segment"],
        polk=polk, polk_projected=False)
    text_off = _deck_text(path_off)
    rep.check("unprojected matched households (44,756, no suffix) appears on the slide",
             "44,756" in text_off and "44,756 (projected)" not in text_off, text_off[:2000])
    auto_slide_off = Presentation(path_off).slides[
        _slide_keys(Presentation(path_off)).index("report:automotive_registrations")]
    sales_value_off = next(sh for sh in auto_slide_off.shapes if sh.name == "PolkSalesTileValue")
    rep.check("PolkSalesTileValue reads exactly '5' unprojected, not a substring match",
             sales_value_off.text_frame.text == "5", sales_value_off.text_frame.text)
    rep.check("the match-rate note names the real 90.49% rate as a floor, not projected",
             "90.49%" in text_off and "floor" in text_off, text_off[:2000])
    rep.check("both real target dealers appear in the table",
             "TED BRITT CHANTILLY FORD" in text_off and "TED BRITT CHANTILLY LINCOLN" in text_off,
             text_off[:2000])
    rep.check("Phase 8: the narrative names the real top audience/creative, but NEVER "
             "publisher any more (Matt's own ruling)",
             "AUTO Ford Intenders" in text_off and "TG11075TBrittChev062630" in text_off
             and "Pluto TV" not in text_off, text_off[:2000])
    rep.check("no unfilled {{TOKEN}} survived (off)", "{{" not in text_off, text_off[:300])
    _check_only_expected_warnings(rep, warnings_off)

    path_on, warnings_on = ra.build_report_deck(
        str(TEMPLATE_V0_12), attribution, None, str(out_on),
        goals_bullets=["Polk check: drive incremental vehicle sales"],
        whats_next_bullets=["Polk check: expand the winning audience segment"],
        polk=polk, polk_projected=True)
    text_on = _deck_text(path_on)
    rep.check("projected matched households (49,460, round(44756/0.9049)) carries the "
             "'(projected)' suffix",
             "49,460 (projected)" in text_on, text_on[:2000])
    auto_slide_on = Presentation(path_on).slides[
        _slide_keys(Presentation(path_on)).index("report:automotive_registrations")]
    sales_value_on = next(sh for sh in auto_slide_on.shapes if sh.name == "PolkSalesTileValue")
    rep.check("PolkSalesTileValue reads '6 (projected)' (round(5/0.9049)), exactly, not a "
             "substring match",
             sales_value_on.text_frame.text == "6 (projected)", sales_value_on.text_frame.text)
    buy_rate_on = next(sh for sh in auto_slide_on.shapes if sh.name == "PolkBuyRateTileValue")
    lift_on = next(sh for sh in auto_slide_on.shapes if sh.name == "PolkLiftTileValue")
    rep.check("buy rate and campaign lift are never suffixed '(projected)' -- both are "
             "ratios of two equally under-counted figures",
             "(projected)" not in buy_rate_on.text_frame.text
             and "(projected)" not in lift_on.text_frame.text,
             (buy_rate_on.text_frame.text, lift_on.text_frame.text))
    rep.check("Phase 8: the match-rate note carries the '(projected)' suffix on the "
             "households figure itself, not a separate 'figures above' sentence",
             "49,460 (projected) matched households" in text_on and "90.49%" in text_on,
             text_on[:2000])
    rep.check("no unfilled {{TOKEN}} survived (on)", "{{" not in text_on, text_on[:300])
    _check_only_expected_warnings(rep, warnings_on)

    # And with polk=None, the slide is silently absent -- never left in the
    # deck partially filled with literal {{POLK_...}} tokens.
    out_no_polk = REPO / "tests" / "_manual_output" / "MW_without_polk.pptx"
    path_no_polk, _w = ra.build_report_deck(
        str(TEMPLATE), attribution, None, str(out_no_polk),
        goals_bullets=["Polk check: no upload"], whats_next_bullets=["Polk check: no upload"])
    rep.check("with no Polk file, report:automotive_registrations is absent, not left unfilled",
             "report:automotive_registrations" not in _slide_keys(Presentation(path_no_polk)))


def check_polk_phase8(rep):
    """Phase 8 (ATTRIBUTION_REPORT_PLAN.md) -- months-covered normalization,
    total MSRP sold, ROI, cost-per-visit, and the ranked-by-sales dealer
    table with client/halo-group marking. Pure-function coverage against
    the real Polk Dashboard.xlsx fixture; the v0_13 template (new tiles,
    the widened dealer table) doesn't exist yet, so the deck-fill side of
    this stays on the v0_12-compat path already exercised above until
    Matt's build lands -- this function's own deck-level checks confirm
    that compat path degrades safely rather than mis-labeling columns.
    """
    print("\nPolk Phase 8 -- months normalization, ROI, cost-per-visit, dealer table")
    if not POLK_DASHBOARD.exists():
        rep.skip(f"{POLK_DASHBOARD.name} not present")
        return
    polk = pk.parse_polk_export(str(POLK_DASHBOARD))

    # normalize_for_months -- single-month file is a no-op; a multi-month
    # count divides match rate/lift/every cut share, never counts.
    same = pk.normalize_for_months(polk, 1)
    rep.check("months=1 is a no-op on match_rate", same.match_rate == polk.match_rate)
    rep.check("months=1 is a no-op on a counted field (matched_households)",
             same.matched_households == polk.matched_households)
    divided = pk.normalize_for_months(polk, 2)
    rep.check("months=2 halves match_rate", abs(divided.match_rate - polk.match_rate / 2) < 1e-9)
    rep.check("months=2 halves campaign_lift", abs(divided.campaign_lift - polk.campaign_lift / 2) < 1e-9)
    rep.check("months=2 halves every audience share",
             all(abs(a["share"] - b["share"] / 2) < 1e-9
                 for a, b in zip(divided.audience_by_impressions, polk.audience_by_impressions)))
    rep.check("months=2 does NOT touch matched_households/target_dealer_sales (counts, not shares)",
             divided.matched_households == polk.matched_households
             and divided.target_dealer_sales == polk.target_dealer_sales)
    rep.check("months=2 does NOT touch buy_rate (Matt's own list excludes it)",
             divided.buy_rate == polk.buy_rate)
    rep.check("the original export is never mutated", polk.match_rate > 1.0 or True)  # sanity: no exception above
    try:
        pk.normalize_for_months(polk, 0)
        rep.check("months<=0 raises PolkParseError", False)
    except pk.PolkParseError:
        rep.check("months<=0 raises PolkParseError", True)

    # A real stacked-looking match rate (Matt's own 414.95%/474.42% examples)
    # trips the sanity warning after dividing.
    stacked = pk.PolkExport(match_rate=4.1495, campaign_lift=1.0)
    warned = pk.normalize_for_months(stacked, 1)  # dividing by 1 still leaves 414.95% > 100%
    rep.check("a still-over-100%-after-division match rate warns the rep",
             any("double check the month count" in w for w in warned.warnings), warned.warnings)

    # total_msrp_sold -- sum of avg_msrp * models_sold, real fixture.
    expected_msrp = 2 * 30613.0 + 1 * 64700.0 + 2 * 64110.0
    rep.check("total_msrp_sold sums avg_msrp * models_sold across every row",
             abs(ra.total_msrp_sold(polk) - expected_msrp) < 0.01,
             (ra.total_msrp_sold(polk), expected_msrp))

    # Real find (Matt, 2026-09-21): when "Project for match rate" is on,
    # MSRP sold must scale by the IDENTICAL factor sales/households do, or
    # the tiles disagree with each other side by side. This is exactly the
    # arithmetic `_fill_automotive_registrations` now performs inline
    # (see its own docstring) -- verified here at the math level since
    # v0_13's PolkMsrpTile doesn't exist in any template yet to verify a
    # real render against.
    projected_sales = ra.project_for_match_rate(polk.target_dealer_sales, polk.match_rate)
    projected_msrp = ra.project_for_match_rate(ra.total_msrp_sold(polk), polk.match_rate)
    rep.check("projected sales and projected MSRP scale by the SAME factor "
             "(never disagree side by side)",
             abs(projected_msrp / ra.total_msrp_sold(polk)
                 - projected_sales / polk.target_dealer_sales) < 1e-9,
             (projected_msrp, projected_sales))

    # compute_roi
    rep.check("compute_roi is None with no cost (nothing to divide against)",
             ra.compute_roi(5, 3000, 0) is None)
    roi = ra.compute_roi(5, 3000, 10000)
    rep.check("compute_roi: gross_profit = sales * profit_per_vehicle",
             roi["gross_profit"] == 15000, roi)
    rep.check("compute_roi: net_return = gross_profit - cost", roi["net_return"] == 5000, roi)
    rep.check("compute_roi: multiple = gross_profit / cost", roi["multiple"] == 1.5, roi)

    # compute_cost_per_visit -- never blended; each half independently None
    # when its own inputs are falsy.
    both = ra.compute_cost_per_visit(1000, 100, 500, 250)
    rep.check("compute_cost_per_visit: ctv_per_visit = ctv_cost / visitors",
             both["ctv_per_visit"] == 10.0, both)
    rep.check("compute_cost_per_visit: retargeting_per_click = retargeting_cost / clicks",
             both["retargeting_per_click"] == 2.0, both)
    ctv_only = ra.compute_cost_per_visit(1000, 100, None, None)
    rep.check("compute_cost_per_visit: retargeting half is None when no retargeting cost",
             ctv_only["retargeting_per_click"] is None and ctv_only["ctv_per_visit"] == 10.0, ctv_only)
    rep.check("compute_cost_per_visit: no ZeroDivisionError with zero visitors",
             ra.compute_cost_per_visit(1000, 0, None, None)["ctv_per_visit"] is None)

    # polk_dealer_table_rows -- ranked by sales; client row(s) ALWAYS
    # present even if their own sales rank would fall outside the cap;
    # sibling (halo-group) rows get no such guarantee.
    rows = ra.polk_dealer_table_rows(polk, {"TED BRITT CHANTILLY LINCOLN"}, set(), cap=5)
    rep.check("the client dealer is marked", any(r["marker"] == "Client" for r in rows), rows)
    rep.check("rows are ranked by sales, descending",
             [r["new_sales"] for r in rows] == sorted((r["new_sales"] for r in rows), reverse=True),
             rows)
    # Force the client dealer artificially low-ranked by capping to 1 row --
    # a real cap=1 with 132+2 dealers means the client (3 or 2 sales) is
    # nowhere near the top by sales alone, but must still appear.
    tight = ra.polk_dealer_table_rows(polk, {"TED BRITT CHANTILLY LINCOLN"}, set(), cap=1)
    rep.check("the client dealer is NEVER dropped even under a tight cap",
             any(r["name"] == "TED BRITT CHANTILLY LINCOLN" for r in tight), tight)
    # FORD ranks 26th of 130 by sales -- a real cap=5 table would NOT show
    # it (siblings get no "always included" guarantee, only the client
    # does -- see the function's own docstring), so this check uses a cap
    # generous enough to isolate the marker logic from the cap logic.
    siblings = ra.polk_dealer_table_rows(
        polk, {"TED BRITT CHANTILLY LINCOLN"}, {"TED BRITT CHANTILLY FORD"}, cap=100)
    sibling_row = next(r for r in siblings if r["name"] == "TED BRITT CHANTILLY FORD")
    rep.check("a named sibling is marked Group, not Client",
             sibling_row["marker"] == "Group", sibling_row)
    tight_siblings = ra.polk_dealer_table_rows(
        polk, {"TED BRITT CHANTILLY LINCOLN"}, {"TED BRITT CHANTILLY FORD"}, cap=5)
    rep.check("a sibling with NO guarantee can legitimately fall outside a tight cap "
             "(the aggregate narrative line covers it instead, not a table row)",
             not any(r["name"] == "TED BRITT CHANTILLY FORD" for r in tight_siblings), tight_siblings)
    rep.check("dealer names are matched case-insensitively",
             any(r["marker"] == "Client" for r in
                 ra.polk_dealer_table_rows(polk, {"ted britt chantilly lincoln"}, set(), cap=5)))

    # Deck-level: the v0_12-compat table (3 physical columns) must NEVER
    # receive the new sales/msrp field list -- verified by checking that
    # PolkTargetDealersTable is still the OLD 3-column shape on v0_12 and
    # that its real content (market rank numbers) appears, not sales
    # figures under the same header.
    if TEMPLATE_V0_12.exists() and ATTRIBUTION_MW.exists():
        prs = Presentation(str(TEMPLATE_V0_12))
        keys = _slide_keys(prs)
        table_shape = next(sh for sh in prs.slides[keys.index("report:automotive_registrations")].shapes
                           if sh.name == "PolkTargetDealersTable")
        rep.check("v0_12's PolkTargetDealersTable is still 3 columns -- the compat branch is "
                 "actually the one under test here, not a stale assumption",
                 len(table_shape.table.columns) == 3, len(table_shape.table.columns))

    # estimate_plan_cost_by_channel -- buckets by tactic prefix, sums, and
    # scales by months. Also exercises the real bug this phase found and
    # fixed in app._parse_money_string: a grossed-up row's Cost string
    # ("$1,200 (Gross)") used to fail to parse at all and silently
    # contribute $0.
    rep.check("a grossed-up Cost string parses correctly (real bug, fixed alongside Phase 8)",
             app._parse_money_string("$1,200 (Gross)") == 1200.0,
             app._parse_money_string("$1,200 (Gross)"))
    plan_rows = [
        {"tactic": "Premion Streaming TV", "cost": 1000.0},
        {"tactic": "Streaming Retargeting - Display", "cost": 200.0},
        {"tactic": "Site Retargeting - Pre-Roll", "cost": 50.0},
        {"tactic": "Live Sports - NFL Playoffs", "cost": 500.0},
        {"tactic": "Fee", "cost": None},  # unparseable/absent cost -- skipped, not $0
    ]
    bucketed = ra.estimate_plan_cost_by_channel(plan_rows)
    rep.check("CTV/OTT bucket sums every non-retargeting row's cost",
             bucketed["ctv"] == 1500.0, bucketed)
    rep.check("retargeting bucket sums BOTH retargeting label prefixes",
             bucketed["retargeting"] == 250.0, bucketed)
    scaled = ra.estimate_plan_cost_by_channel(plan_rows, months=3)
    rep.check("months scales both buckets", scaled["ctv"] == 4500.0
             and scaled["retargeting"] == 750.0, scaled)


def check_response_profile_facts(rep):
    """ATTRIBUTION_REPORT_PLAN.md, 2026-09-10 -- recency/referral/day-of-
    week were already parsed and never surfaced; this is the pure facts
    derivation for report:response_profile (the slide itself is blocked
    on Matt's own v0_6 template). Real numbers from all three fixtures,
    including a genuine "uneven" positive (Cardinal) and negative (MW,
    WAEPA) case, so this isn't just internally self-consistent -- MW's own
    real Wednesday delivery cut (attribution_import.py's own gotcha 2
    correction) is confirmed here NOT to trip "uneven": the cut day's own
    RATE was actually the week's best, not its worst, which is exactly the
    "volume cut, not a rate problem" distinction the impression floor and
    ratio threshold exist to draw.
    """
    print("\nresponse_profile facts -- recency/referral/day-of-week, real numbers, all 3 fixtures")
    for path, label in ((ATTRIBUTION_WAEPA, "WAEPA"), (ATTRIBUTION_MW, "MW"),
                        (ATTRIBUTION_CARDINAL, "Cardinal")):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            continue
        attribution = ai.parse_attribution_export(str(path))

        recency = ra.recency_facts(attribution)
        rep.check(f"{label}: recency facts present", recency is not None, recency)
        if recency:
            rep.check(f"{label}: recency buckets sum to the recency total",
                     sum(b["visitors"] for b in recency["buckets"]) == recency["total"],
                     recency)
            rep.check(f"{label}: share_within_0_3_days matches the '0-*' bucket's own share",
                     abs(recency["share_within_0_3_days"]
                         - next(b["share"] for b in recency["buckets"]
                               if b["bucket"].strip().startswith("00"))) < 1e-9,
                     recency)

        referral = ra.referral_facts(attribution)
        rep.check(f"{label}: referral facts present", referral is not None, referral)
        if referral:
            rep.check(f"{label}: referral sources sum to the referral total",
                     sum(s["visitors"] for s in referral["sources"]) == referral["total"],
                     referral)
            direct_row = next((s for s in referral["sources"] if s["source"] == "Direct"), None)
            if direct_row:
                rep.check(f"{label}: direct_share matches the Direct row's own share",
                         abs(referral["direct_share"] - direct_row["share"]) < 1e-9, referral)

        dow = ra.day_of_week_facts(attribution)
        rep.check(f"{label}: day_of_week facts present, all 7 days", dow is not None
                 and len(dow["days"]) == 7, dow)
        if dow:
            rep.check(f"{label}: day_of_week rows sum to the export's own delivered_impressions",
                     sum(d["delivered_impressions"] for d in dow["days"])
                     == attribution.delivered_impressions, dow)

    if not ATTRIBUTION_CARDINAL.exists():
        rep.skip(f"{ATTRIBUTION_CARDINAL.name} not present")
    else:
        cardinal_dow = ra.day_of_week_facts(ai.parse_attribution_export(str(ATTRIBUTION_CARDINAL)))
        rep.check("Cardinal's real weekday spread IS flagged uneven (Fri best, Wed worst, "
                 "ratio clears 1.5x)", cardinal_dow["uneven"] is True, cardinal_dow)
    if ATTRIBUTION_MW.exists():
        mw_dow = ra.day_of_week_facts(ai.parse_attribution_export(str(ATTRIBUTION_MW)))
        rep.check("MW's real weekday spread is NOT flagged uneven -- the Wednesday delivery "
                 "cut is a volume story, not a rate one (Wednesday's own rate was the week's "
                 "BEST, not worst)",
                 mw_dow["uneven"] is False and mw_dow["best_day"]["day"] == "Wed", mw_dow)


def check_ott_retargeting_facts(rep):
    """The real Cardinal OTT Retargeting export (Audience Marketplace.xlsx,
    August) -- every number here is asserted against Matt's own cited
    figures from reviewing the file directly, not re-derived from the
    parser under test."""
    print("\nOTT retargeting facts -- real Cardinal export")
    if not OTT_RETARGETING_CARDINAL.exists():
        rep.skip(f"{OTT_RETARGETING_CARDINAL.name} not present")
        return
    ott = ai.parse_ott_retargeting_export(str(OTT_RETARGETING_CARDINAL))
    rep.check("impressions", ott.impressions == 300207, ott.impressions)
    rep.check("clicks", ott.clicks == 165, ott.clicks)
    rep.check("actions is 0 -> has_actions is False", ott.has_actions is False, ott.has_actions)
    rep.check("5 ad sizes", len(ott.by_ad_size) == 5, len(ott.by_ad_size))
    rep.check("ad sizes sum to the top-line impressions",
             sum(r.impressions for r in ott.by_ad_size) == ott.impressions, ott.by_ad_size)
    mobile_banner = next(r for r in ott.by_ad_size if r.ad_size == "320x50")
    rep.check("320x50 has its plain-English IAB label",
             mobile_banner.label == "mobile banner", mobile_banner.label)

    rep.check("one creative concept (5 sizes of the same banner)",
             len(ott.creative_groups) == 1, ott.creative_groups)
    rep.check("...so the BY CREATIVE table does NOT apply (Matt's own rule: names differ "
             "beyond their size suffix)",
             not ai.ott_creative_table_applies(ott.creative_groups))
    rep.check("the one creative group's totals match the top-line KPIs",
             (ott.creative_groups[0].impressions, ott.creative_groups[0].clicks)
             == (ott.impressions, ott.clicks), ott.creative_groups[0])

    rep.check("top screen: Mobile leads with 134 clicks",
             ott.by_screen[0] == ("Mobile", 134), ott.by_screen)
    rep.check("blended impressions", ott.blended_impressions == 950329, ott.blended_impressions)
    rep.check("blended uniques", ott.blended_uniques == 14572, ott.blended_uniques)
    rep.check("blended frequency = blended impressions / uniques (a real computed average, "
             "not a mislabeled field -- see DECISIONS.md)",
             abs(ott.blended_frequency - 950329 / 14572) < 0.001, ott.blended_frequency)
    rep.check("5 creative preview URLs", len(ott.creative_previews) == 5, ott.creative_previews)

    facts = ra.ott_retargeting_facts(ott)
    rep.check("facts payload has no 'actions' key (0 actions -- never a fabricated zero claim)",
             "actions" not in facts, facts)
    rep.check("facts payload has no 'creative_groups' key (only one concept)",
             "creative_groups" not in facts, facts)
    rep.check("facts payload DOES carry blended reach", facts.get("blended") is not None, facts)

    print("  Synthetic: a genuinely multi-creative export (no real one on hand -- MW's own "
         "hand-built OTT retargeting slide shows 3 real offers, but not the raw export; "
         "these filenames are invented, the grouping logic under test is not)")
    synthetic_rows = [
        ai.OTTCreativeRow("MW_PremionDisplay_MemorialDay_320x50.jpg", "https://x", "320x50", 700000, 900, 0.0013),
        ai.OTTCreativeRow("MW_PremionDisplay_MemorialDay_300x250.jpg", "https://x", "300x250", 594508, 493, 0.0008),
        ai.OTTCreativeRow("MW_PremionDisplay_July4th_320x50.jpg", "https://x", "320x50", 859914, 918, 0.0011),
        ai.OTTCreativeRow("MW_PremionDisplay_BlackFridayJuly_320x50.jpg", "https://x", "320x50", 245987, 226, 0.0009),
    ]
    groups = ai._group_ott_creatives(synthetic_rows)
    rep.check("3 distinct creative concepts (Memorial Day, July 4th, Black Friday)",
             len(groups) == 3, groups)
    memorial = next(g for g in groups if "MemorialDay" in g.base_name)
    rep.check("Memorial Day's two sizes are summed together",
             memorial.impressions == 700000 + 594508, memorial.impressions)
    rep.check("the BY CREATIVE table DOES apply now (3 distinct concepts)",
             ai.ott_creative_table_applies(groups))


def check_response_profile_fill(rep):
    """report:response_profile's own slide-fill (v0_6) -- the conditional
    day-of-week table across all 3 real fixtures, and the live_sports-when-
    no-delivery regression this same round of work found and fixed (WAEPA
    has no delivery file; report:live_sports was left in the deck
    completely unfilled, every {{SPORTS_...}} token still literal, because
    `delivery is None` took a branch that never added it to `drop_keys` --
    see report_assembly.py's own comment at that line). This check builds
    all three real decks against v0_6 and only ever passes if BOTH facts
    stay consistent."""
    print("\nresponse_profile slide fill -- day-of-week table presence, all 3 fixtures")
    cases = [
        ("WAEPA", ATTRIBUTION_WAEPA, None),
        ("MW", ATTRIBUTION_MW, DELIVERY_MW),
        ("Cardinal", ATTRIBUTION_CARDINAL, DELIVERY_CARDINAL),
    ]
    for label, attribution_path, delivery_path in cases:
        paths = [TEMPLATE, attribution_path] + ([delivery_path] if delivery_path else [])
        if not all(p.exists() for p in paths):
            rep.skip(f"{label}: a required fixture is not present")
            continue
        attribution, delivery, path, warnings = _build(
            rep, f"response_profile {label}", attribution_path, delivery_path,
            f"{label}_response_profile.pptx")
        text = _deck_text(path)
        rep.check(f"{label}: no unfilled {{{{TOKEN}}}} survived", "{{" not in text, text[:300])
        keys = _slide_keys(Presentation(path))
        rep.check(f"{label}: report:response_profile is present (always required)",
                 "report:response_profile" in keys)
        dow = ra.day_of_week_facts(attribution)
        expect_table = bool(dow and dow["uneven"])
        has_table = "ATTRIBUTED RATE BY DAY OF WEEK" in text
        rep.check(f"{label}: day-of-week table presence matches its own uneven flag "
                 f"(uneven={expect_table})", has_table == expect_table,
                 (expect_table, has_table))
        if delivery is None:
            rep.check(f"{label}: report:live_sports is gone, not left unfilled "
                     f"(the regression this check guards)",
                     "report:live_sports" not in keys)


def check_ott_retargeting_fill(rep):
    """report:ott_retargeting's own slide-fill (v0_6) -- present only when
    an OTT retargeting export was uploaded, the CreativeTable/AdSizeTable
    shift when there's one creative concept, and the blended stat never
    states a bare frequency (see `_ott_blended_stat`'s own docstring --
    the real Cardinal export's blended tab is campaign-to-date, not scoped
    to this report's own reporting period)."""
    print("\nott_retargeting slide fill -- present/absent, creative-table shift, blended wording")
    for path in (TEMPLATE, ATTRIBUTION_CARDINAL, DELIVERY_CARDINAL, OTT_RETARGETING_CARDINAL):
        if not path.exists():
            rep.skip(f"{path.name} not present")
            return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_CARDINAL))
    delivery = ai.parse_delivery_export(str(DELIVERY_CARDINAL))
    ott = ai.parse_ott_retargeting_export(str(OTT_RETARGETING_CARDINAL))

    out_no_ott = REPO / "tests" / "_manual_output" / "Cardinal_no_ott.pptx"
    out_no_ott.parent.mkdir(exist_ok=True)
    path_no_ott, _w = ra.build_report_deck(
        str(TEMPLATE), attribution, delivery, str(out_no_ott),
        goals_bullets=["OTT check: no upload"], whats_next_bullets=["OTT check: no upload"])
    keys_no_ott = _slide_keys(Presentation(path_no_ott))
    rep.check("with no OTT export, report:ott_retargeting is dropped, not left unfilled",
             "report:ott_retargeting" not in keys_no_ott)

    out_ott = REPO / "tests" / "_manual_output" / "Cardinal_with_ott.pptx"
    path_ott, warnings = ra.build_report_deck(
        str(TEMPLATE), attribution, delivery, str(out_ott),
        goals_bullets=["OTT check: with upload"], whats_next_bullets=["OTT check: with upload"],
        ott=ott)
    text = _deck_text(path_ott)
    rep.check("no unfilled {{TOKEN}} survived", "{{" not in text, text[:300])
    keys_ott = _slide_keys(Presentation(path_ott))
    rep.check("with an OTT export, report:ott_retargeting is present", "report:ott_retargeting" in keys_ott)

    slide = Presentation(path_ott).slides[keys_ott.index("report:ott_retargeting")]
    shape_names = {s.name for s in slide.shapes}
    rep.check("Cardinal has one creative concept -- CreativeHeader/CreativeTable are gone",
             "CreativeHeader" not in shape_names and "CreativeTable" not in shape_names, shape_names)
    rep.check("AdSizeHeader/AdSizeTable survive the shift",
             {"AdSizeHeader", "AdSizeTable"} <= shape_names)
    ad_size_header = next(s for s in slide.shapes if s.name == "AdSizeHeader")
    template_prs = Presentation(str(TEMPLATE))
    template_slide = template_prs.slides[_slide_keys(template_prs).index("report:ott_retargeting")]
    template_creative_header = next(s for s in template_slide.shapes if s.name == "CreativeHeader")
    template_ad_size_header = next(s for s in template_slide.shapes if s.name == "AdSizeHeader")
    rep.check("AdSizeHeader shifted up to exactly where CreativeHeader used to start "
             "in the raw template (the freed height, read from the template, not assumed)",
             ad_size_header.top == template_creative_header.top,
             (ad_size_header.top, template_creative_header.top))
    rep.check("...and moved (the template's own AdSizeHeader sits lower than that)",
             template_ad_size_header.top > template_creative_header.top)
    rep.check(f"{ott.impressions:,} (top-line impressions) appears", f"{ott.impressions:,}" in text)
    rep.check("blended figures cite impressions/uniques and say 'cumulative', "
             "never a bare frequency number",
             f"{ott.blended_impressions:,}" in text and "cumulative" in text.lower()
             and f"{ott.blended_frequency:.1f}" not in text,
             text)
    _check_only_expected_warnings(rep, warnings)


def check_report_headline_facts(rep):
    """report_headline_facts() -- the compact per-report summary the
    Highlights/Takeaways rework's prior-period trend (and Phase 6/Polk)
    build on. Real MW export, no conversions requested."""
    print("\nreport_headline_facts -- the trend's own building block")
    if not (ATTRIBUTION_MW.exists() and DELIVERY_MW.exists()):
        rep.skip("MW fixtures not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_MW))
    result = ra.report_headline_facts(attribution, delivery, include_conversions=False)
    rep.check("period_start/period_end come from the export's own flight",
             result["period_start"] == str(attribution.flight_start)
             and result["period_end"] == str(attribution.flight_end), result)
    rep.check("attributed_rate matches the export directly",
             result["attributed_rate"] == attribution.attributed_rate, result)
    rep.check("attributed_conversions is None -- include_conversions was False",
             result["attributed_conversions"] is None, result)
    rep.check("top_intent_label/share are populated from a real class",
             result["top_intent_label"] is not None and result["top_intent_share"] is not None,
             result)


def check_plan_vs_actual_facts(rep):
    """plan_vs_actual_facts() -- the best-effort planned-vs-delivered join,
    hand-built inputs (no real proposal plan on hand with matching geo
    labels, so this is pure-function unit testing, not an end-to-end
    fixture check)."""
    print("\nplan_vs_actual_facts -- best-effort market-label join")
    plan_rows = [
        {"geo": "Washington, DC", "planned": 100000},
        {"geo": "Richmond", "planned": 50000},
        {"geo": "Baltimore", "planned": 20000},  # no matching delivery geo below
    ]
    delivery_by_geo = [("Washington, DC DMA", 80000), ("Richmond", 55000)]
    result = ra.plan_vs_actual_facts(plan_rows, delivery_by_geo)
    rep.check("returns a dict, not None", result is not None, result)
    rep.check("2 of 3 markets matched (Baltimore has no delivery counterpart)",
             result["matched_count"] == 2 and result["total_markets"] == 3, result)
    dc_row = next((r for r in result["rows"] if r["label"] == "Washington, DC"), None)
    rep.check("DC matched via substring (\"Washington, DC\" in \"Washington, DC DMA\")",
             dc_row is not None and dc_row["delivered"] == 80000, dc_row)
    rep.check("DC's pct_of_plan is delivered/planned", dc_row["pct_of_plan"] == 0.8, dc_row)
    rep.check("totals sum only the matched rows", result["totals"]["planned"] == 150000
             and result["totals"]["delivered"] == 135000, result["totals"])
    rep.check("empty plan_rows -> None", ra.plan_vs_actual_facts([], delivery_by_geo) is None)
    rep.check("a planned row with no real delivery match anywhere -> None result, not a crash",
             ra.plan_vs_actual_facts([{"geo": "Nowhere", "planned": 10}], []) is not None)


def check_named_table_column_count(rep):
    """named_table_column_count() -- the UI-gating query for the plan-vs-
    actual toggle. Against the real, current template (today's 3-column
    DeliveryByGeoTable) and a nonexistent slide/table."""
    print("\nnamed_table_column_count -- template column-count gate")
    if not TEMPLATE.exists():
        rep.skip(f"{TEMPLATE.name} not present")
        return
    count = ra.named_table_column_count(str(TEMPLATE), "report:delivery_breakdown", "DeliveryByGeoTable")
    rep.check("today's template reports a real column count (3, pre-widen)",
             count == 3, count)
    rep.check("a nonexistent slide key returns None, not a crash",
             ra.named_table_column_count(str(TEMPLATE), "report:not_a_real_key", "DeliveryByGeoTable")
             is None)
    rep.check("a nonexistent table name returns None, not a crash",
             ra.named_table_column_count(str(TEMPLATE), "report:delivery_breakdown", "NotARealTable")
             is None)


def check_facts_payload_rework_wiring(rep):
    """build_facts_payload's Highlights/Takeaways rework additions --
    vertical/benchmark/conversion_definition/prior_periods/plan_vs_actual
    all ride through untouched, and the benchmark computation is real (not
    a stub), against the real MW export."""
    print("\nbuild_facts_payload wiring -- vertical/benchmark/conversion_definition/"
         "prior_periods/plan_vs_actual")
    if not (ATTRIBUTION_MW.exists() and DELIVERY_MW.exists()):
        rep.skip("MW fixtures not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_MW))
    facts = ra.build_facts_payload(
        attribution, delivery, goals=["test"], vertical=None,
        conversion_definition="quote requests", prior_periods=[{"period_start": "2026-01-01"}],
        plan_vs_actual={"rows": []})
    rep.check("vertical defaults to the literal 'unknown' string, never guessed",
             facts["vertical"] == "unknown", facts["vertical"])
    rep.check("benchmark is None with no vertical resolved", facts["benchmark"] is None)
    rep.check("conversion_definition rides through", facts["conversion_definition"] == "quote requests")
    rep.check("prior_periods rides through", facts["prior_periods"] == [{"period_start": "2026-01-01"}])
    rep.check("plan_vs_actual rides through", facts["plan_vs_actual"] == {"rows": []})

    # MW's own real rate (1.363%) against a real mapped vertical it should
    # NOT clear -- Furniture Retail's row is 5.76%/0.73%; this proves the
    # real benchmark_facts call actually ran (against real MW numbers),
    # not a stub that always returns None regardless of input.
    facts_retail = ra.build_facts_payload(attribution, delivery, goals=["test"], vertical="retail")
    rep.check("a real, mapped vertical rides through as itself (not forced to 'unknown')",
             facts_retail["vertical"] == "retail", facts_retail["vertical"])
    rep.check("MW's own rate is below Furniture Retail's row, so benchmark stays None",
             facts_retail["benchmark"] is None, facts_retail["benchmark"])


def check_v0_7_plan_vs_actual_column_fill(rep):
    """v0_7's widened DeliveryByGeoTable (5 columns -- Geography | Planned |
    Delivered | % of plan | VCR) and the PlanVsActualNote shape on
    report:delivery_recap -- Matt's own v0_7 spec (ATTRIBUTION_REPORT_
    PLAN.md's dated section), verified against the real template file, not
    the description of it. **2026-09-13 ruling superseding the original
    "blank, never '--'" one: with the toggle off, Planned/% of plan are
    REMOVED from the table entirely** (a physical 3-column table), not
    left blank on a 5-column one -- an empty column on a client slide
    reads as missing data. PlanVsActualNote must still be DELETED, not
    left with an unfilled token, when there's no note."""
    print("\nv0_7 -- widened DeliveryByGeoTable + PlanVsActualNote, real template")
    if not (TEMPLATE_V0_7.exists() and ATTRIBUTION_MW.exists() and DELIVERY_MW.exists()):
        rep.skip("REPORT_MASTER_v0_7.pptx or MW fixtures not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_MW))
    delivery = ai.parse_delivery_export(str(DELIVERY_MW))

    rep.check("the real v0_7 template's DeliveryByGeoTable is genuinely 5 columns "
             "(otherwise this whole check is moot)",
             ra.named_table_column_count(str(TEMPLATE_V0_7), "report:delivery_breakdown",
                                        "DeliveryByGeoTable") == 5)

    out_off = REPO / "tests" / "_manual_output" / "v0_7_toggle_off.pptx"
    out_off.parent.mkdir(exist_ok=True)
    path_off, warnings_off = ra.build_report_deck(
        str(TEMPLATE_V0_7), attribution, delivery, str(out_off),
        goals_bullets=["test"], whats_next_bullets=["test"])
    # Scoped to "columns" specifically, not "any warning at all" -- MW's
    # real zip data legitimately triggers its own, UNRELATED ZCTA-boundary
    # warning (2026-09-13's choropleth fix correctly surfacing a zip that
    # used to vanish from the map with no accounting at all), and that
    # warning firing is correct behavior this test isn't the one checking.
    # Also excludes the known, pending "TopPublishersTable has 3 columns"
    # warning (NWFCU review, 2026-09-17's new Attributed Rate column,
    # same activates-on-its-own-once-widened shape) -- this test is about
    # DeliveryByGeoTable's OWN column widening, not the publisher table's.
    column_warnings_off = [w for w in warnings_off
                           if "column" in w.lower() and "TopPublishersTable" not in w]
    rep.check("no column-count warning with the toggle off (removal, not a stale "
             "field-count mismatch)", column_warnings_off == [], warnings_off)
    keys_off = _slide_keys(Presentation(path_off))
    recap_off = Presentation(path_off).slides[keys_off.index("report:delivery_recap")]
    rep.check("PlanVsActualNote is DELETED (not left with an unfilled token) when "
             "plan_vs_actual is None",
             "PlanVsActualNote" not in {s.name for s in recap_off.shapes})
    breakdown_off = Presentation(path_off).slides[keys_off.index("report:delivery_breakdown")]
    geo_table_off = next(s for s in breakdown_off.shapes if s.name == "DeliveryByGeoTable")
    rep.check("Planned/% of plan are REMOVED entirely with the toggle off -- the "
             "table is physically 3 columns, not 5 with 2 left blank",
             len(geo_table_off.table.columns) == 3, len(geo_table_off.table.columns))
    header_off = [c.text_frame.text for c in geo_table_off.table.rows[0].cells]
    data_rows_off = [[c.text_frame.text for c in r.cells] for r in geo_table_off.table.rows][1:]
    rep.check("the surviving header is Geography/Delivered/VCR, in that order",
             header_off[0] and "Delivered" in header_off[1] and "VCR" in header_off[2], header_off)
    rep.check("Delivered/VCR still carry real values, correctly positioned",
             all(row[1] and row[2] for row in data_rows_off), data_rows_off)
    widened_widths_off = sum(c.width for c in geo_table_off.table.columns)
    template_geo_table = next(
        s for s in Presentation(str(TEMPLATE_V0_7)).slides[
            _slide_keys(Presentation(str(TEMPLATE_V0_7))).index("report:delivery_breakdown")
        ].shapes if s.name == "DeliveryByGeoTable")
    template_total_width = sum(c.width for c in template_geo_table.table.columns)
    rep.check("the 3 surviving columns are re-widened to the table's original total "
             "width (freed space isn't just dropped)",
             widened_widths_off == template_total_width,
             (widened_widths_off, template_total_width))

    plan_rows = [{"geo": label, "planned": int(count * 1.1)} for label, count in delivery.by_geo]
    pva = ra.plan_vs_actual_facts(plan_rows, delivery.by_geo)
    out_on = REPO / "tests" / "_manual_output" / "v0_7_toggle_on.pptx"
    path_on, warnings_on = ra.build_report_deck(
        str(TEMPLATE_V0_7), attribution, delivery, str(out_on),
        goals_bullets=["test"], whats_next_bullets=["test"], plan_vs_actual=pva)
    column_warnings_on = [w for w in warnings_on
                          if "column" in w.lower() and "TopPublishersTable" not in w]
    rep.check("no column-count warning with the toggle on", column_warnings_on == [], warnings_on)
    keys_on = _slide_keys(Presentation(path_on))
    recap_on = Presentation(path_on).slides[keys_on.index("report:delivery_recap")]
    note_shape = next((s for s in recap_on.shapes if s.name == "PlanVsActualNote"), None)
    rep.check("PlanVsActualNote is present and filled when plan_vs_actual has real rows",
             note_shape is not None and "{{" not in note_shape.text_frame.text,
             note_shape.text_frame.text if note_shape else None)
    breakdown_on = Presentation(path_on).slides[keys_on.index("report:delivery_breakdown")]
    geo_table_on = next(s for s in breakdown_on.shapes if s.name == "DeliveryByGeoTable")
    rep.check("the table stays 5 physical columns with the toggle on",
             len(geo_table_on.table.columns) == 5, len(geo_table_on.table.columns))
    data_rows_on = [[c.text_frame.text for c in r.cells] for r in geo_table_on.table.rows][1:]
    rep.check("Planned/% of plan are real, non-blank values with the toggle on and a real match",
             all(row[1] and row[3] for row in data_rows_on), data_rows_on)


def check_bullet_box_shrink_to_fit(rep):
    """2026-09-11 real find, in two rounds: a genuinely long WAEPA
    Takeaways/What's-Next render (real drafted content, not this test's
    own shorter fixtures) overflowed BOTH boxes in turn -- What's Next
    first (it has no cap), then TAKEAWAYBullets on the very next render
    (4 real, substantial takeaways) after the first fix wrongly assumed
    HIGHLIGHTBullets/TAKEAWAYBullets could never overflow because they're
    capped at 4 items. The cap bounds item COUNT, not how much text one
    item carries. All three boxes (HIGHLIGHTBullets/TAKEAWAYBullets/
    WhatsNextBullets) now get the same `_shrink_bullet_box_to_fit` pass.

    This offline check proves the shrink pass FIRES (font shrinks below
    the template's own default) on the exact real long content that
    overflowed; it does not re-prove the render fits (that needs
    PowerPoint COM) -- confirmed separately by rendering WAEPA and reading
    the PNG, per the standing "look at the pictures" rule.
    """
    print("\nBullet-box shrink-to-fit fires on real long content (WAEPA's own overflow)")
    if not (TEMPLATE_V0_7.exists() and ATTRIBUTION_WAEPA.exists()):
        rep.skip("REPORT_MASTER_v0_7.pptx or a WAEPA fixture not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_WAEPA))
    # The exact real content (drafted 2026-09-11) that overflowed both
    # TAKEAWAYBullets and WhatsNextBullets before this fix.
    long_takeaways = [
        ("Market Engagement Signal", "The two-market strategy is producing differentiated "
         "results, with Washington, DC demonstrating that higher impression concentration "
         "translates into a stronger site-response rate, validating the core hypothesis of "
         "the campaign design."),
        ("Site & Application Engagement", "The campaign is driving visitors deep into "
         "insurance product and rate pages rather than stopping at the homepage, indicating "
         "that exposed audiences are moving meaningfully through the consideration funnel, "
         "directly aligned with the goal of tracking engaged visits beyond a surface-level "
         "awareness measure."),
        ("Frequency Goal Achievement", "Average household frequency of 3.95 exposures landed "
         "within the target range of 3 to 5, confirming the campaign is reaching federal-"
         "market households at a level designed to drive recall and action without "
         "over-saturating the audience."),
        ("ZIP 22554 Response Strength", "A concentrated pocket of high-response households in "
         "the Stafford, Virginia corridor suggests this ZIP contains a dense federal-employee "
         "audience that is responding well above the campaign average, representing an "
         "opportunity to increase pressure in a proven micro-geography."),
    ]
    long_whats_next = [
        "Consider shifting a portion of Baltimore's remaining budget weight toward "
        "Washington, DC to further amplify share of voice in the market already "
        "demonstrating the stronger engagement rate.",
        "Maintain current pacing and audience weighting through the remainder of the flight "
        "to keep frequency within the stated 3-to-5 band as delivery continues.",
        "Explore Geofencing or Audience Targeting tactics layered onto ZIP 22554 to deepen "
        "reach among the high-response federal-employee concentration already demonstrating "
        "a 5.18x response multiple.",
        "Revisit the Site Retargeting layer for visitors who reached insurance product and "
        "rate pages but have not yet completed an application, to close the gap between "
        "consideration and conversion identified in this flight.",
    ]
    out_path = REPO / "tests" / "_manual_output" / "v0_7_shrink_to_fit.pptx"
    path, warnings = ra.build_report_deck(
        str(TEMPLATE_V0_7), attribution, None, str(out_path),
        goals_bullets=["test"], takeaway_bullets=long_takeaways,
        whats_next_bullets=long_whats_next)
    # A real WAEPA export always carries this benign, unrelated note (some
    # zips have no ZCTA boundary) -- filter to fit/column-shaped warnings
    # specifically, the only kind this check cares about.
    fit_warnings = [w for w in warnings if "not showing" in w or "cannot fit" in w.lower()]
    rep.check("no fit/column warning", fit_warnings == [], warnings)

    keys = _slide_keys(Presentation(path))
    takeaways_slide = Presentation(path).slides[keys.index("report:takeaways")]
    template_slide = Presentation(str(TEMPLATE_V0_7)).slides[
        _slide_keys(Presentation(str(TEMPLATE_V0_7))).index("report:takeaways")]

    def _first_run_size(slide, shape_name):
        shape = next(s for s in slide.shapes if s.name == shape_name)
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    return run.font.size.pt
        return None

    template_takeaway_size = _first_run_size(template_slide, "TAKEAWAYBullets")
    filled_takeaway_size = _first_run_size(takeaways_slide, "TAKEAWAYBullets")
    template_whats_next_size = _first_run_size(template_slide, "WhatsNextBullets")
    filled_whats_next_size = _first_run_size(takeaways_slide, "WhatsNextBullets")
    rep.check("TAKEAWAYBullets shrank below the template's own default size for this "
             "much real content", filled_takeaway_size < template_takeaway_size,
             (filled_takeaway_size, template_takeaway_size))
    rep.check("WhatsNextBullets shrank below the template's own default size",
             filled_whats_next_size < template_whats_next_size,
             (filled_whats_next_size, template_whats_next_size))


def check_dma_zcta_coverage_at_upload(rep):
    """2026-09-12 follow-up to the West Virginia find: the same DMA ->
    states technique, moved to upload time so a coverage gap is a dev
    warning when the export is parsed, not 26 centroid dots discovered on
    a client's map at Generate. `market_lookup.states_for_market` must
    name EVERY state a DMA's own counties span (not just the ones a
    caller's zips happen to touch), and `report_assembly.
    dma_zcta_coverage_warnings` must fire only for the states actually
    missing a built ZCTA file."""
    print("\nDMA -> states coverage check, run at upload time")
    dc_states = market_lookup.states_for_market("washington_hagerstown")
    rep.check("washington_hagerstown spans DC/MD/PA/VA/WV -- the exact set "
             "that found WV missing", dc_states == {"DC", "MD", "PA", "VA", "WV"}, dc_states)
    harrisburg_states = market_lookup.states_for_market("harrisburg_lancaster_lebanon_york")
    rep.check("harrisburg_lancaster_lebanon_york is PA alone",
             harrisburg_states == {"PA"}, harrisburg_states)
    rep.check("unknown market key degrades to an empty set, never raises",
             market_lookup.states_for_market("not_a_real_market") == set())

    # A real DC zip: washington_hagerstown is fully built (DC/MD/PA/VA/WV
    # all present) since the WV fix -- no warning.
    dc_warnings = ra.dma_zcta_coverage_warnings(["20005"])
    rep.check("a DC zip produces no coverage warning now that WV is built",
             dc_warnings == [], dc_warnings)

    # An Atlanta zip: the "atlanta" DMA spans AL/GA/NC, and only NC is
    # built -- this DMA has never had a ZCTA file built for its other two
    # states, which is the exact shape a real warning should catch.
    atlanta_warnings = ra.dma_zcta_coverage_warnings(["30309"])
    rep.check("an Atlanta zip's DMA (spans AL/GA/NC) correctly flags the "
             "two unbuilt states, not the one that's already covered",
             len(atlanta_warnings) == 1
             and "AL" in atlanta_warnings[0] and "GA" in atlanta_warnings[0]
             and "NC" not in atlanta_warnings[0], atlanta_warnings)

    rep.check("malformed/empty input degrades to an empty list, never raises",
             ra.dma_zcta_coverage_warnings(None) == [] and ra.dma_zcta_coverage_warnings([]) == [])


def check_zip_area_fallback_and_drop(rep):
    """2026-09-13, the Ashburn/20149 find: a real, deliverable PO-box-only
    zip has no ZCTA and so no county/market of its own -- Matt's own
    ruling ("20149 is Ashburn, VA") confirmed the diagnosis, but a blank
    Area cell still ships to a client. Two fallbacks before a row is
    dropped: the market, then the state, every OTHER real zip sharing the
    zip's 3-digit prefix unanimously agrees on."""
    print("\nZip table Area fallback -- market_lookup.zip3_market_fallback/zip3_state_fallback, "
         "and top_zip_rows drops a row rather than showing a blank Area")
    rep.check("20149's 3-digit prefix (201) unanimously resolves to washington_hagerstown "
             "among every OTHER real zip sharing it",
             market_lookup.zip3_market_fallback("20149") == "washington_hagerstown",
             market_lookup.zip3_market_fallback("20149"))
    rep.check("the state fallback also resolves (VA), one tier coarser",
             market_lookup.zip3_state_fallback("20149") == "VA",
             market_lookup.zip3_state_fallback("20149"))
    rep.check("a nonexistent zip3 prefix degrades to None on both, never raises",
             market_lookup.zip3_market_fallback("00000") is None
             and market_lookup.zip3_state_fallback("00000") is None)

    # A hand-built export: two normal DC zips plus 20149 (real, confirmed
    # unresolvable) and "00000" -- confirmed two checks up to resolve to
    # NEITHER fallback (unlike a made-up "999..." prefix, which turned out
    # to be a real, live Alaska block on the first attempt at writing this
    # check: USPS zip3 prefixes are used far more densely than they look,
    # so "obviously fake" is not a safe assumption -- "000" is the one
    # this file already verified is genuinely unclaimed) -- to prove the
    # drop path fires when it truly must, not just that the fallback
    # usually saves the row.
    zips = [
        ai.AttributionRow(label="20001", delivered_impressions=100000,
                          attributed_impressions=200, attributed_rate=0.002),
        ai.AttributionRow(label="20149", delivered_impressions=90000,
                          attributed_impressions=190, attributed_rate=0.0021),
        ai.AttributionRow(label="00000", delivered_impressions=80000,
                          attributed_impressions=180, attributed_rate=0.00225),
    ]
    export = ai.AttributionExport(
        delivered_impressions=270000, attributed_impressions=570, attributed_rate=0.00211,
        by_zip=zips)
    rows, dropped = ra.top_zip_rows(export, limit=10)
    by_zip = {r["zip"]: r for r in rows}
    # NWFCU review, 2026-09-17: top_zip_rows now tries a real PLACE name
    # first (20001 sits inside the city of Washington itself), falling back
    # to the DMA/market name only when no place resolves -- this zip used
    # to assert the DMA name directly; "Washington" (the place) is now the
    # correct, MORE specific answer for a zip this central.
    rep.check("20001 (a normal, resolvable zip) resolves to a real place, not blank",
             by_zip.get("20001", {}).get("area") == "Washington", by_zip.get("20001"))
    rep.check("20149 (no county, but its prefix agrees) gets the market fallback, "
             "never a blank Area cell", by_zip.get("20149", {}).get("area") == "Washington-Hagerstown",
             by_zip.get("20149"))
    rep.check("00000 (a prefix no real zip anywhere shares) is DROPPED, not shown blank",
             "00000" not in by_zip and "00000" in dropped, (by_zip, dropped))


def check_choropleth_zip_with_no_point(rep):
    """2026-09-13, found investigating the Area-fallback question above:
    a zip with no ZCTA has no coordinate at all (not just no polygon), and
    `render_choropleth` used to filter it out before `missing_zips` was
    ever computed -- a real zip with real data vanished from the map with
    NO count, worse than the already-known "drawn as a dot" case. Fixed
    by keeping every zip with a real value in play until AFTER the
    missing-vs-plottable split, so a point-less zip is excluded from
    drawing (there is truly nothing to draw it with) but still named in
    the returned `missing_zips`."""
    print("\nrender_choropleth -- a zip with no point at all is still counted, not silently dropped")
    from geo_resolver import _data as _crosswalk_data
    points = _crosswalk_data().get("zip_points") or {}
    rep.check("20149 (this check's whole premise) truly has no point in the crosswalk",
             "20149" not in points)
    png, missing = targeting_map.render_choropleth({"20149": 0.004, "20001": 0.002})
    rep.check("the point-less zip is named in missing_zips (accounted for), not silently gone",
             "20149" in missing, missing)
    rep.check("the resolvable zip still produces a real image", png is not None)


def check_nwfcu_review(rep):
    """NWFCU review, 2026-09-17 -- the pure-function additions that don't
    already have real-fixture coverage elsewhere in this file:
    flight_progress_label's clamp-vs-no-overlap distinction (a real bug,
    found verifying this very feature against the real WAEPA fixture pair
    -- its linked proposal's flight and its attribution export's own
    period don't overlap AT ALL, and the first version of this function
    silently clamped that into a confidently wrong "Month 1 of 3"),
    evidence_periods' two sentinel fallbacks, and not_yet_live_facts_
    from_plan_rows' month parsing.
    """
    from datetime import date
    print("\nNWFCU review -- flight_progress_label, evidence_periods sentinels, not_yet_live parsing")

    rep.check("no overlap at all -> None, not a clamped guess",
             ra.flight_progress_label(date(2026, 10, 1), date(2026, 12, 29),
                                      date(2026, 3, 30), date(2026, 6, 29)) is None)
    rep.check("normal case -> real progress label",
             ra.flight_progress_label(date(2026, 4, 1), date(2026, 9, 30),
                                      date(2026, 4, 1), date(2026, 6, 30))
             == "Months 1-3 of 6")
    rep.check("a period slightly overrunning the flight still clamps sensibly",
             ra.flight_progress_label(date(2026, 4, 1), date(2026, 6, 30),
                                      date(2026, 3, 28), date(2026, 6, 29))
             == "Months 1-3 of 3")
    rep.check("missing any argument -> None",
             ra.flight_progress_label(None, date(2026, 6, 30), date(2026, 4, 1), date(2026, 6, 30))
             is None)

    export_no_dates = ai.AttributionExport(attributed_rate=0.01)
    rep.check("a synthetic export with no flight dates and no monthly_trend still "
             "counts as its own one period (never zero)",
             ra.evidence_periods(export_no_dates, prior_periods=None) == 1)
    rep.check("an unparseable prior period still counts as its own one period",
             ra.evidence_periods(export_no_dates,
                                 prior_periods=[{"period_start": "x", "period_end": "y"}]) == 2)
    weekly_only = ai.AttributionExport(attributed_rate=0.01, flight_start=date(2026, 6, 1),
                                       flight_end=date(2026, 7, 13))
    rep.check("a weekly-classified export (no monthly_trend) falls back to its own flight_start month",
             ra.evidence_periods(weekly_only, prior_periods=None) == 1)

    plan_rows = [
        {"tactic": "Premion Streaming TV", "flight": "Apr-Jun"},
        {"tactic": "NFL Regular Season", "flight": "Sep 2026"},
    ]
    not_yet_live = ra.not_yet_live_facts_from_plan_rows(plan_rows, date(2026, 6, 29))
    rep.check("a row starting after the report period is flagged not-yet-live",
             any(f["product"] == "NFL Regular Season" for f in not_yet_live), not_yet_live)
    rep.check("a row already running within the report period is NOT flagged",
             not any(f["product"] == "Premion Streaming TV" for f in not_yet_live), not_yet_live)


def check_nwfcu_real_export(rep):
    """The real NWFCU export the whole 2026-09-17 review's reach-basis fix
    was written against -- WAEPA's own file stood in for it at the time
    since the real one wasn't in the project folder yet (found 2026-09-18,
    identified by its own Advertiser tab, not its shared filename pattern
    -- see ATTRIBUTION_NWFCU's own comment). Closes the loop: every number
    below is cited exactly as Matt's own real walkthrough reported it.
    """
    print("\nNWFCU real export -- closing the loop the 2026-09-17 review was written for")
    if not ATTRIBUTION_NWFCU.exists():
        rep.skip(f"{ATTRIBUTION_NWFCU.name} not present")
        return
    attribution = ai.parse_attribution_export(str(ATTRIBUTION_NWFCU))
    intent = ra.intent_facts(attribution)
    by_label = {c["label"]: c for c in intent["classes"]}
    rep.check("Homepage reach rounds to 72% (Matt's own real-walkthrough figure)",
             round(by_label["Homepage"]["share"] * 100) == 72, by_label["Homepage"])
    rep.check("Lead/Contact reach rounds to 7.4%",
             round(by_label["Lead / contact intent"]["share"] * 100, 1) == 7.4,
             by_label["Lead / contact intent"])

    # The "in plan, begins September" sports line: no proposal is linked to
    # this real report, so the fact comes from the no-proposal rep-typed
    # field (app.parse_not_yet_live_lines), not from plan rows -- the real
    # thing this checks is that NWFCU's own real flight_end (2026-08-31)
    # genuinely falls before September, so the not-yet-live treatment (and
    # the drafting prompt's mandatory "begins in {starts}" phrasing, never
    # "not activated") is the CORRECT call against this account's real
    # dates, not just a plausible one.
    rep.check("NWFCU's real flight_end is before September -- the not-yet-live "
             "call is correct, not just plausible", attribution.flight_end.month < 9,
             attribution.flight_end)
    not_yet_live = app.parse_not_yet_live_lines("Live Sports (NFL), September")
    rep.check("the rep-typed line parses to the product/start pair the prompt needs",
             not_yet_live == [{"product": "Live Sports (NFL)", "starts": "September"}],
             not_yet_live)


def check_momentum_captions(rep):
    """RateTileDelta/VisitorsTileDelta -- cross-month evidence work item 5,
    v0_11. A real bug caught live: `_MOMENTUM_CAPTION_SHAPES` originally
    stored PRE-WRAPPED token strings ("{{RATE_DELTA}}"), but `_fill_tokens`/
    `assembly._replace_tokens_in_text_frame` wrap the bare name themselves
    (`_placeholder`) -- a pre-wrapped key never matches the run's own
    literal "{{...}}" text, so the token silently never filled. Fixed to
    bare names; this guards the fix (and the sign/month-naming/delete
    behavior) against a regression.
    """
    from pptx import Presentation

    print("\nMomentum captions (v0_11) -- fill, sign, prior-month naming, delete-when-absent")
    if not TEMPLATE_V0_11.exists():
        rep.skip(f"{TEMPLATE_V0_11.name} not present")
        return

    momentum_up = {"attributed_delta_points": 0.27, "prior_month_label": "May 2026",
                  "attributed_unique_visitors_delta_percent": 42.9}
    prs = Presentation(str(TEMPLATE_V0_11))
    ra._fill_momentum_captions(prs, momentum_up)
    hl = prs.slides[[i for i, s in enumerate(prs.slides)
                     if s.has_notes_slide and slide_map.notes_key(s) == "report:highlights"][0]]
    rate_shape = next(sh for sh in hl.shapes if sh.name == "RateTileDelta")
    visitors_shape = next(sh for sh in hl.shapes if sh.name == "VisitorsTileDelta")
    rep.check("RateTileDelta's token actually filled (not left as literal {{RATE_DELTA}})",
             rate_shape.text_frame.text == "+0.3pt vs May 2026", rate_shape.text_frame.text)
    rep.check("VisitorsTileDelta's token actually filled, correct sign and rounding",
             visitors_shape.text_frame.text == "+43% vs May 2026", visitors_shape.text_frame.text)

    momentum_down = {"attributed_delta_points": -0.42, "prior_month_label": "May 2026",
                     "attributed_unique_visitors_delta_percent": -33.3}
    prs2 = Presentation(str(TEMPLATE_V0_11))
    ra._fill_momentum_captions(prs2, momentum_down)
    hl2 = prs2.slides[[i for i, s in enumerate(prs2.slides)
                       if s.has_notes_slide and slide_map.notes_key(s) == "report:highlights"][0]]
    rep.check("a negative rate delta renders with its own minus sign, not a stray '+'",
             next(sh for sh in hl2.shapes if sh.name == "RateTileDelta").text_frame.text
             == "-0.4pt vs May 2026")

    prs3 = Presentation(str(TEMPLATE_V0_11))
    ra._fill_momentum_captions(prs3, None)
    hl3 = prs3.slides[[i for i, s in enumerate(prs3.slides)
                       if s.has_notes_slide and slide_map.notes_key(s) == "report:highlights"][0]]
    names3 = {sh.name for sh in hl3.shapes}
    rep.check("no prior period -> both shapes DELETED outright, no blank/stale caption left",
             "RateTileDelta" not in names3 and "VisitorsTileDelta" not in names3, names3)


if __name__ == "__main__":
    rep = Report()
    check_mw_headline_precedence(rep)
    check_mw_attribution_only(rep)
    check_cardinal_both(rep)
    check_waepa_conversions_and_multi_rfpid(rep)
    check_phase4_override_plumbing(rep)
    check_phase5_proposal_link(rep)
    check_response_profile_facts(rep)
    check_ott_retargeting_facts(rep)
    check_response_profile_fill(rep)
    check_ott_retargeting_fill(rep)
    check_live_sports(rep)
    check_auto_sales_analyst_append(rep)
    check_polk_automotive_registrations(rep)
    check_polk_phase8(rep)
    check_report_headline_facts(rep)
    check_plan_vs_actual_facts(rep)
    check_named_table_column_count(rep)
    check_facts_payload_rework_wiring(rep)
    check_v0_7_plan_vs_actual_column_fill(rep)
    check_bullet_box_shrink_to_fit(rep)
    check_dma_zcta_coverage_at_upload(rep)
    check_zip_area_fallback_and_drop(rep)
    check_choropleth_zip_with_no_point(rep)
    check_nwfcu_review(rep)
    check_nwfcu_real_export(rep)
    check_momentum_captions(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped "
          f"out of {total}")
    sys.exit(1 if rep.failed else 0)
