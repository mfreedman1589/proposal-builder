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


def _slide_keys(prs):
    """Every slide's `key:` note, in deck order."""
    out = []
    for slide in prs.slides:
        if slide.has_notes_slide:
            for line in slide.notes_slide.notes_text_frame.text.splitlines():
                if line.strip().lower().startswith("key:"):
                    out.append(line.split(":", 1)[1].strip())
    return out


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


if __name__ == "__main__":
    rep = Report()
    check_mw_headline_precedence(rep)
    check_mw_attribution_only(rep)
    check_cardinal_both(rep)
    check_waepa_conversions_and_multi_rfpid(rep)
    check_phase4_override_plumbing(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed, {len(rep.skipped)} skipped "
          f"out of {total}")
    sys.exit(1 if rep.failed else 0)
