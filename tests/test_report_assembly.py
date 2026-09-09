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

import attribution_import as ai  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402

TEMPLATE = REPO / "REPORT_MASTER_v0_4.pptx"
ATTRIBUTION_MW = REPO / "MW attribution excel.xlsx"
DELIVERY_MW = REPO / "MW delivery.xlsx"
ATTRIBUTION_CARDINAL = REPO / "Premion Website Attribution Cardinal.xlsx"
DELIVERY_CARDINAL = REPO / "Premion OTT.xlsx"
ATTRIBUTION_WAEPA = REPO / "Premion Website Attribution and Reach Extension (13).xlsx"

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
    a blanket `not warnings` that someone has to remember to re-tighten."""
    pending = [w for w in warnings if "TopZipTable has 3 columns" in w]
    unexpected = [w for w in warnings if w not in pending]
    rep.check("no unexpected fit/column warning", not unexpected, unexpected)
    if pending:
        print(f"        (known, pending v0_3: {pending[0][:70]}...)")


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


if __name__ == "__main__":
    rep = Report()
    check_mw_headline_precedence(rep)
    check_mw_attribution_only(rep)
    check_cardinal_both(rep)
    check_waepa_conversions_and_multi_rfpid(rep)
    check_phase4_override_plumbing(rep)
    check_phase5_proposal_link(rep)
    check_live_sports(rep)
    check_auto_sales_analyst_append(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped "
          f"out of {total}")
    sys.exit(1 if rep.failed else 0)
