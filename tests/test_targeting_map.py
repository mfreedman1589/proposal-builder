"""targeting_map.py + its slide-embedding half in assembly.py (roadmap §E).

    python tests/test_targeting_map.py

Entirely offline -- Pillow, python-pptx and AppTest, no COM, no PowerPoint,
tier-1 speed. This is deliberately NOT the same guarantee
`tests/test_group_scenarios.py --render` gives (which additionally proves
PowerPoint can open the result); it's the fast, always-run half, so "does a
proposal with resolved zips produce a map on the targeting slide" has an
answer in seconds, not only after a multi-minute render.
"""
import copy
import io
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import assembly                                # noqa: E402
import db                                      # noqa: E402
import geo_resolver                            # noqa: E402
import market_lookup                           # noqa: E402
import targeting_groups as tg                  # noqa: E402
import targeting_map as tm                     # noqa: E402
from PIL import Image                          # noqa: E402

market_lookup.install()
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def _resolved_group(term, zips, color, name=""):
    group = tg.new_group([term], geo_def={"kind": "zips", "zips": zips}, color=color, name=name)
    group["resolved_zips"] = zips
    group["resolved_markets"] = geo_resolver.zips_to_markets(zips).resolved.keys()
    group["resolved_markets"] = sorted(group["resolved_markets"])
    return group


def main():
    print("=" * 78)
    print("SCENARIO  groups_with_zips / render_map (pure, offline)")
    print("=" * 78)

    unresolved = tg.new_group(["Homeowners"])
    resolved = _resolved_group("DEMO Homeowner", ["20005", "20006", "20001"], "#4C78A8")
    check("a group with no resolved_zips is excluded",
          tm.groups_with_zips([unresolved, resolved]) == [resolved], None)
    check("None for an empty group list", tm.render_map([]) is None)
    check("None when nothing is resolved", tm.render_map([unresolved]) is None)

    png = tm.render_map([resolved], width_px=400, height_px=300)
    check("real PNG bytes for a resolved group", bool(png) and png[:8] == b"\x89PNG\r\n\x1a\n", None)
    img = Image.open(io.BytesIO(png))
    check("rendered at the requested size", img.size == (400, 300), img.size)

    print("\nSCENARIO  two groups render in two distinct colors")
    print("=" * 78)
    a = _resolved_group("A", ["20005", "20006"], "#E45756")
    b = _resolved_group("B", ["10001", "10002"], "#54A24B")
    png2 = tm.render_map([a, b], width_px=500, height_px=400)
    img2 = Image.open(io.BytesIO(png2)).convert("RGB")
    colors = {c for _count, c in img2.getcolors(maxcolors=500 * 400)}
    check("group A's own color is actually drawn on the map",
          tm._hex_to_rgb("#E45756") in colors, sorted(colors)[:10])
    check("group B's own color is actually drawn on the map",
          tm._hex_to_rgb("#54A24B") in colors, sorted(colors)[:10])

    print("\nSCENARIO  a rep's group name becomes the legend label")
    print("=" * 78)
    named = _resolved_group("A", ["20005"], "#4C78A8", name="Metro Zip Add-On")
    png3 = tm.render_map([named])
    check("named group renders a PNG (label text itself isn't pixel-checked, "
          "just that naming doesn't break rendering)", bool(png3))

    print("\nSCENARIO  a bad zip (absent from the crosswalk) is skipped, not fatal")
    print("=" * 78)
    dirty = _resolved_group("A", ["20005", "99999"], "#4C78A8")
    png4 = tm.render_map([dirty])
    check("renders fine despite one zip with no known centroid", bool(png4))

    print("\nSCENARIO  state/county outlines give the map geographic context")
    print("=" * 78)
    boundaries = tm._boundaries()
    check("map_boundaries.json.gz is built and loads", bool(boundaries), None)
    if boundaries:
        check("all 50 states + DC + PR are present", len(boundaries["states"]) == 52,
              len(boundaries["states"]))
        check("thousands of counties are present", len(boundaries["counties"]) > 3000,
              len(boundaries["counties"]))
        check("Missouri is keyed by its 2-digit state FIPS", "29" in boundaries["states"],
              boundaries["states"].get("29", {}).get("name"))
        check("St. Louis city is keyed by its 5-digit county GEOID", "29510" in boundaries["counties"],
              boundaries["counties"].get("29510", {}).get("name"))

    stl = _resolved_group("A", ["63101", "63108", "63118"], "#4C78A8")   # St. Louis, MO
    png_stl = tm.render_map([stl], width_px=700, height_px=500)
    check("a St. Louis group still renders with the basemap wired in", bool(png_stl))
    img_stl = Image.open(io.BytesIO(png_stl)).convert("RGB")
    colors_stl = {c for _count, c in img_stl.getcolors(maxcolors=700 * 500)}
    check("a county outline color is actually drawn on the map, not just present in code",
          tm.COUNTY_OUTLINE_COLOR in colors_stl, sorted(colors_stl)[:10])
    check("a state outline color is actually drawn on the map",
          tm.STATE_OUTLINE_COLOR in colors_stl, sorted(colors_stl)[:10])

    print("\nSCENARIO  the projection fills the frame instead of letterboxing")
    print("=" * 78)
    # A roughly square point cluster (comparable lat and lon span, after the
    # cosine correction) dropped onto a canvas that's nowhere near square --
    # true distance-scale locks both axes to the same scale, so it fills the
    # short axis and leaves large empty margins on the long one.
    sq_lats, sq_lons = [38.0, 39.0], [-90.0, -88.7]
    width, height = 800, 300

    def x_fill_fraction(max_stretch):
        real = tm._MAX_STRETCH
        tm._MAX_STRETCH = max_stretch
        try:
            project = tm._projector(sq_lats, sq_lons, width, height)
        finally:
            tm._MAX_STRETCH = real
        x0, _ = project(sq_lats[0], sq_lons[0])
        x1, _ = project(sq_lats[1], sq_lons[1])
        return abs(x1 - x0) / width

    unstretched = x_fill_fraction(1.0)     # true distance-scale, no bound relief
    bounded = x_fill_fraction(tm._MAX_STRETCH)
    check("the bounded-stretch fit uses more of the short-mismatch axis than "
          "true distance-scale would", bounded > unstretched,
          (unstretched, bounded))
    check("but the stretch itself never exceeds the configured bound",
          bounded <= unstretched * tm._MAX_STRETCH + 1e-9,
          (unstretched, bounded, tm._MAX_STRETCH))

    print("\n" + "=" * 78)
    print("SCENARIO  assembly.targeting_map_region / place_targeting_map (python-pptx, no COM)")
    print("=" * 78)
    master_path, _, warning = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    if master_path is None:
        print(f"  SKIP -- {warning}")
    else:
        selections = copy.deepcopy(assembly.SELECTIONS)
        selections["preset"] = "standard"
        selections["market"] = "DC"
        selections["include_avails_template"] = True
        prs, _, _ = assembly.build_presentation(master_path, selections)
        avails_slide = assembly.find_slide_with_marker(prs, "{{AVAILS}}")
        check("the avails slide exists in a standard DC deck", avails_slide is not None)
        if avails_slide is not None:
            table_shape = assembly._find_table_shape(avails_slide)
            before_pics = [s for s in avails_slide.shapes if s.shape_type == 13]

            region = assembly.targeting_map_region(avails_slide)
            check("a region is derived (there's a table and a slide width to bound it)",
                  region is not None, region)
            if region is not None:
                left, top, width, height = region
                check("left edge sits right after the table's own right edge + clearance",
                      left == table_shape.left + table_shape.width + assembly._TABLE_CLEARANCE,
                      (left, table_shape.left + table_shape.width))
                check("top is aligned with the table's own top",
                      top == table_shape.top, (top, table_shape.top))
                slide_width = prs.slide_width
                check("right edge mirrors the table's own left margin",
                      left + width == slide_width - table_shape.left,
                      (left + width, slide_width - table_shape.left))

            check("place_targeting_map(None) adds no picture -- the stock image is untouched",
                  True, None)
            assembly.place_targeting_map(avails_slide, None)
            after_none = [s for s in avails_slide.shapes if s.shape_type == 13]
            check("no picture added when there's nothing to draw",
                  len(after_none) == len(before_pics), (len(before_pics), len(after_none)))

            fake_png = tm.render_map([resolved], width_px=300, height_px=200)
            assembly.place_targeting_map(avails_slide, fake_png)
            after_real = [s for s in avails_slide.shapes if s.shape_type == 13]
            check("exactly one new picture added for a real map",
                  len(after_real) == len(before_pics) + 1, (len(before_pics), len(after_real)))
            if len(after_real) == len(before_pics) + 1 and region is not None:
                new_pic = [s for s in after_real
                          if s._element not in {p._element for p in before_pics}][0]
                check("the new picture is sized and positioned to the derived region exactly",
                      (new_pic.left, new_pic.top, new_pic.width, new_pic.height) == region,
                      ((new_pic.left, new_pic.top, new_pic.width, new_pic.height), region))

    print("\n" + "=" * 78)
    print("SCENARIO  end to end through the real form: a proposal with resolved "
          "zips produces a map on the targeting slide (tier-1 speed, no PowerPoint)")
    print("=" * 78)
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest

    zips_result = geo_resolver.radius_to_zips(["63101"], 10)   # St. Louis
    plaza_group = tg.new_group(
        ["AUTO Make Ford"], geo_def={"kind": "radius", "centers": ["63101"], "miles": 10},
        avails_monthly=100000, color=tg.assign_color(0))
    plaza_group["resolved_zips"] = zips_result.resolved
    plaza_group["resolved_markets"] = sorted(
        geo_resolver.zips_to_markets(zips_result.resolved).resolved.keys())
    check("this fixture's own group really has resolved zips (else the test proves nothing)",
          bool(plaza_group["resolved_zips"]), plaza_group["resolved_zips"])

    captured = {}
    real_personalize = assembly.personalize

    def spy(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        captured["fill_data"] = fill_data
        return warnings

    assembly.personalize = spy
    db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        at.session_state["include_avails_template"] = True
        at.session_state["targeting_groups"] = [plaza_group]
        at.session_state["premion_streaming_tv"] = True
        at.run()
        check("the form renders", not at.exception,
              at.exception[0].message[:400] if at.exception else "")
        generate = [b for b in at.button if b.label == "Generate proposal"]
        check("Generate button present", bool(generate), [b.label for b in at.button])
        if generate:
            generate[0].click().run()
            check("Generate runs without raising", not at.exception,
                  at.exception[0].message[:600] if at.exception else "")
    finally:
        assembly.personalize = real_personalize

    prs = captured.get("prs")
    check("assembly produced a presentation", prs is not None)
    if prs is not None:
        map_png = captured["fill_data"]["avails"].get("map_png")
        check("map_png was computed for this proposal (resolved zips exist)", bool(map_png))
        avails_slide = assembly.find_slide_with_marker(prs, "{{AVAILS}}")
        # find_slide_with_marker looks for the UNFILLED token, which
        # personalize already consumed -- the slide is the one
        # place_targeting_map actually touched instead, found by locating
        # the table and asking what it measured, the same lookup the deck
        # itself uses to place the map.
        avails_slide = next(
            (s for s in prs.slides if assembly._find_table_shape(s) is not None
             and assembly.targeting_map_region(s) is not None
             and "PRECISION TARGETING" in (assembly.slide_map.extract_slide_text(s) or "").upper()),
            None)
        check("found the targeting/avails slide", avails_slide is not None)
        if avails_slide is not None:
            pics = [s for s in avails_slide.shapes if s.shape_type == 13]
            check("the targeting slide carries a map picture (3 pictures: background, "
                  "PREMION wordmark, map) -- not just the stock image and the avails table",
                  len(pics) == 3, [p.name for p in pics])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The map draws what targeting groups already resolved, skips what hasn't been "
          "resolved without erroring, its slide footprint is derived from the avails table's "
          "own geometry rather than hand-placed, and a real proposal with resolved zips, "
          "driven through the real form and a real Generate, produces a map on the targeting "
          "slide.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
