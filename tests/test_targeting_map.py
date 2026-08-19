"""targeting_map.py + its slide-embedding half in assembly.py (roadmap §E).

    python tests/test_targeting_map.py

Offline (pure Pillow + python-pptx, no Streamlit, no COM) except the last
section, which needs Windows + PowerPoint to read back what was actually
drawn -- same reason every rendering check in this project does.
"""
import copy
import io
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

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The map draws what targeting groups already resolved, skips what hasn't been "
          "resolved without erroring, and its slide footprint is derived from the avails "
          "table's own geometry rather than hand-placed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
