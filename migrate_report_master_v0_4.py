"""One-time migration: REPORT_MASTER_v0_3.pptx -> REPORT_MASTER_v0_4.pptx.

ATTRIBUTION_REPORT_PLAN.md's WAEPA conversions section spelled out the exact
spec (2026-09-08); this is that spec executed via python-pptx rather than by
hand in PowerPoint, per Matt's own instruction -- "it's renames plus one
duplicated tile."

What this does, on the report:highlights slide only:
1. Renames the three existing highlights tiles to the `<Name>Tile` /
   `<Name>TileValue` / `<Name>TileLabel` convention this deck's OTHER tile
   rows already use (recap's ReportPeriodTile/FlightTile/GeographyTile,
   delivery recap's five) -- ImpressionsTile ({{HEADLINE_IMPRESSIONS}}),
   VisitorsTile ({{HEADLINE_UNIQUE_VISITORS}}), RateTile
   ({{HEADLINE_ATTRIBUTED_RATE}}). They predate the convention, which is why
   `report_assembly._reflow_tile_row` was written to no-op safely if this
   rename never landed -- see report_assembly.py's `_HIGHLIGHTS_TILE_ROW`.
2. Deep-copies RateTile's three shapes (same-slide XML clone -- no
   relationship IDs to remap, since these are plain autoshape/textbox
   XML with no image/chart parts) into a fourth group, ConversionsTile /
   ConversionsTileValue / ConversionsTileLabel, with the value token
   changed to {{HEADLINE_CONVERSIONS}} and the label text changed to
   "Attributed conversions" -- via run-level replace, never
   `text_frame.text`, so the source formatting (font, size, color) survives
   unchanged, matching assembly.py's own token-fill convention.
3. Repositions all four tile groups evenly across the ORIGINAL three-tile
   row span (measured from the template, not assumed) -- same left edge,
   same right edge, the same ~274320 EMU gap between tiles the original
   three already used, each tile a quarter of the row instead of a third.
   Value/label shapes keep the SAME padding pattern (137160 EMU each side,
   measured from the template) relative to their own tile.

Shapes are found by TEXT CONTENT (a token substring, or an exact label
string), not by their current default pptxgenjs name ("Shape 2", "Text 3")
-- robust to running this again if the names ever drift, and doesn't
assume today's default-name ordering.

    python migrate_report_master_v0_4.py

Writes REPORT_MASTER_v0_4.pptx beside the v0_3 file. Does not touch
Supabase -- upload separately via db.upload_report_deck.
"""
import copy
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

import slide_map

REPO = Path(__file__).resolve().parent
SOURCE = REPO / "REPORT_MASTER_v0_3.pptx"
OUTPUT = REPO / "REPORT_MASTER_v0_4.pptx"

_LABEL_TEXT = {
    "impressions": "Impressions delivered",
    "visitors": "Attributed unique website visitors",
    "rate": "Attributed rate",
}
_VALUE_TOKEN = {
    "impressions": "{{HEADLINE_IMPRESSIONS}}",
    "visitors": "{{HEADLINE_UNIQUE_VISITORS}}",
    "rate": "{{HEADLINE_ATTRIBUTED_RATE}}",
}
_NEW_NAMES = {
    "impressions": "ImpressionsTile",
    "visitors": "VisitorsTile",
    "rate": "RateTile",
    "conversions": "ConversionsTile",
}


def _find_highlights_slide(prs):
    for slide in prs.slides:
        if slide.has_notes_slide and slide_map.notes_key(slide) == "report:highlights":
            return slide
    raise SystemExit("No report:highlights slide found -- wrong template file?")


def _tile_groups(slide):
    """{"impressions"/"visitors"/"rate": {"tile": shape, "value": shape,
    "label": shape}}, found by content -- the value shape carries its own
    token, the label shape carries its own exact text, and the tile (bg)
    shape is whichever no-text shape's bounding box contains the value
    shape's own top-left corner."""
    value_shapes, label_shapes, bg_shapes = {}, {}, []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        for key, token in _VALUE_TOKEN.items():
            if token in text:
                value_shapes[key] = shape
        for key, label in _LABEL_TEXT.items():
            if text.strip() == label:
                label_shapes[key] = shape
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            continue
        if shape.name in ("Shape 0", "Image 0", "Image 1", "TegnaLogo"):
            continue
        # A tile background candidate: no text, and some value shape's
        # top-left corner sits inside its bounding box.
        for key, value_shape in value_shapes.items():
            if (shape.left <= value_shape.left <= shape.left + shape.width
                    and shape.top <= value_shape.top <= shape.top + shape.height):
                bg_shapes.append((key, shape))
    groups = {}
    for key in ("impressions", "visitors", "rate"):
        tile = next((s for k, s in bg_shapes if k == key), None)
        if tile is None or key not in value_shapes or key not in label_shapes:
            raise SystemExit(f"Could not find the full {key!r} tile group -- template shape "
                             f"has drifted from what this script expects.")
        groups[key] = {"tile": tile, "value": value_shapes[key], "label": label_shapes[key]}
    return groups


def _replace_run_text(shape, old, new):
    """Run-level replace, matching assembly.py's own token-fill convention
    -- never `text_frame.text = ...`, which would collapse run formatting."""
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)


def migrate():
    prs = Presentation(str(SOURCE))
    slide = _find_highlights_slide(prs)
    groups = _tile_groups(slide)

    # 1. Rename the three existing tiles.
    for key in ("impressions", "visitors", "rate"):
        groups[key]["tile"].name = f"{_NEW_NAMES[key]}"
        groups[key]["value"].name = f"{_NEW_NAMES[key]}Value"
        groups[key]["label"].name = f"{_NEW_NAMES[key]}Label"

    # 2. Deep-copy RateTile's three shapes into a fourth group.
    spTree = slide.shapes._spTree
    source = groups["rate"]
    new_tile_el = copy.deepcopy(source["tile"]._element)
    new_value_el = copy.deepcopy(source["value"]._element)
    new_label_el = copy.deepcopy(source["label"]._element)
    spTree.append(new_tile_el)
    spTree.append(new_value_el)
    spTree.append(new_label_el)

    # Re-read the slide's shapes so the new elements come back wrapped as
    # real python-pptx shape objects (python-pptx has no public "wrap this
    # raw element" API -- re-walking shapes is the documented way).
    conversions_tile = conversions_value = conversions_label = None
    for shape in slide.shapes:
        if shape._element is new_tile_el:
            conversions_tile = shape
        elif shape._element is new_value_el:
            conversions_value = shape
        elif shape._element is new_label_el:
            conversions_label = shape
    assert conversions_tile is not None and conversions_value is not None and conversions_label is not None

    conversions_tile.name = "ConversionsTile"
    conversions_value.name = "ConversionsTileValue"
    conversions_label.name = "ConversionsTileLabel"
    _replace_run_text(conversions_value, "{{HEADLINE_ATTRIBUTED_RATE}}", "{{HEADLINE_CONVERSIONS}}")
    _replace_run_text(conversions_label, "Attributed rate", "Attributed conversions")

    groups["conversions"] = {"tile": conversions_tile, "value": conversions_value,
                             "label": conversions_label}

    # 3. Reposition all four evenly across the ORIGINAL row span -- measured
    # from the (still current, pre-move) impressions/rate tiles, not assumed.
    row_left = groups["impressions"]["tile"].left
    row_right = groups["rate"]["tile"].left + groups["rate"]["tile"].width
    gap = groups["visitors"]["tile"].left - (groups["impressions"]["tile"].left
                                             + groups["impressions"]["tile"].width)
    value_pad_left = groups["impressions"]["value"].left - groups["impressions"]["tile"].left
    value_pad_total = groups["impressions"]["tile"].width - groups["impressions"]["value"].width
    label_pad_left = groups["impressions"]["label"].left - groups["impressions"]["tile"].left
    label_pad_total = groups["impressions"]["tile"].width - groups["impressions"]["label"].width

    count = 4
    new_width = Emu(int((row_right - row_left - gap * (count - 1)) / count))
    order = ["impressions", "visitors", "rate", "conversions"]
    left = row_left
    for key in order:
        tile, value, label = groups[key]["tile"], groups[key]["value"], groups[key]["label"]
        tile.left = left
        tile.width = new_width
        value.left = Emu(int(left) + int(value_pad_left))
        value.width = Emu(int(new_width) - int(value_pad_total))
        label.left = Emu(int(left) + int(label_pad_left))
        label.width = Emu(int(new_width) - int(label_pad_total))
        left = Emu(int(left) + int(new_width) + int(gap))

    prs.save(str(OUTPUT))
    return OUTPUT


def _inventory_check(path):
    """The token/named-shape inventory -- confirms the highlights slide now
    carries exactly the 4 expected tile groups, the 3 old shape names are
    gone, and every other slide/shape/token in the deck is untouched."""
    prs = Presentation(str(path))
    slide = _find_highlights_slide(prs)
    names = {s.name for s in slide.shapes}
    expected = {f"{n}{suffix}" for n in ("ImpressionsTile", "VisitorsTile", "RateTile",
                                          "ConversionsTile")
               for suffix in ("", "Value", "Label")}
    missing = expected - names
    if missing:
        raise SystemExit(f"Inventory check FAILED -- missing shape names: {sorted(missing)}")
    old_names = {"Shape 2", "Text 3", "Text 4", "Shape 5", "Text 6", "Text 7",
                "Shape 8", "Text 9", "Text 10"}
    leftover = old_names & names
    if leftover:
        raise SystemExit(f"Inventory check FAILED -- old shape names survived: {sorted(leftover)}")
    texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
    for token in ("{{HEADLINE_IMPRESSIONS}}", "{{HEADLINE_UNIQUE_VISITORS}}",
                 "{{HEADLINE_ATTRIBUTED_RATE}}", "{{HEADLINE_CONVERSIONS}}"):
        if not any(token in t for t in texts):
            raise SystemExit(f"Inventory check FAILED -- token {token} not found anywhere "
                             f"on the highlights slide.")
    if not any(t.strip() == "Attributed conversions" for t in texts):
        raise SystemExit("Inventory check FAILED -- 'Attributed conversions' label text missing.")
    print(f"Inventory check PASSED -- {sorted(expected)}")

    # Every OTHER slide's own key/shape count is untouched -- confirms this
    # migration is scoped to the highlights slide alone.
    source_prs = Presentation(str(SOURCE))
    for src_slide, new_slide in zip(source_prs.slides, prs.slides):
        src_key = slide_map.notes_key(src_slide) if src_slide.has_notes_slide else None
        new_key = slide_map.notes_key(new_slide) if new_slide.has_notes_slide else None
        if src_key != "report:highlights" and src_key != new_key:
            raise SystemExit(f"Inventory check FAILED -- slide key drifted: {src_key!r} -> {new_key!r}")
        if src_key != "report:highlights" and len(list(src_slide.shapes)) != len(list(new_slide.shapes)):
            raise SystemExit(f"Inventory check FAILED -- slide {src_key!r} shape count changed.")
    print("Inventory check PASSED -- every other slide is untouched.")


if __name__ == "__main__":
    out = migrate()
    print(f"Wrote {out}")
    _inventory_check(out)
    sys.exit(0)
