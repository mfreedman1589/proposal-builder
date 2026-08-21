"""condense_avails_table's growth cap AND font floor (assembly.py) --
offline, no PowerPoint needed.

Two bugs on the same table, found a commit apart from the same live deck
(Plaza Motors):

1. A live proposal (three avails rows) rendered its avails table at 3.82in
   tall: condense_avails_table shrinks to fit but, like the media plan
   table, also hands out leftover slide space evenly across rows so a short
   table doesn't sit tiny in a mostly-empty region -- and for a table this
   short, with a floor this far down the slide, that grew three ordinary
   rows into something visibly stretched. Fixed by capping a row's grown
   height at the template's OWN row height (or the row's real content need,
   if that's taller -- shrink-to-fit is unaffected) instead of handing out
   the full slack.

2. The very deck used to verify fix #1 exposed a second, opposite bug: the
   master deck's avails table declares its data row at 0.21in -- shorter
   than its own 14pt content needs, which PowerPoint silently grows past
   the declared height to draw. condense_avails_table took the declared
   height at face value, derived a 6pt ceiling from it (_max_font_for_row),
   and rendered every avails table at 6pt regardless of how much room the
   slide actually had -- a real proposal shipped at 6pt with four inches of
   slide to spare. Fixed by trusting the template's own authored font size
   (read off its runs) as the ceiling, falling back to the height-derived
   estimate only when no run carries an explicit size.

`tests/test_group_scenarios.py --render` covers real documents end-to-end
over COM; this isolates both behaviors against an artificially generous
floor, which the real fixtures don't happen to produce (their floors leave
less slack than this).
"""
import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import assembly                                # noqa: E402
import db                                      # noqa: E402
from pptx import Presentation                   # noqa: E402
from pptx.util import Emu                       # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def avails_slide_and_table(prs):
    for slide in prs.slides:
        table_shape = assembly._find_table_shape(slide)
        if table_shape is not None and len(table_shape.table.rows) >= 2:
            header = table_shape.table.rows[0].cells[0].text_frame.text.strip().lower()
            if "audience" in header or "geo" in header:
                return slide, table_shape
    return None, None


def template_row_height(master_path):
    """The avails table's per-row height as the deck's designer drew it,
    read from an untouched copy of the master -- never from a deck that's
    already been through condense_avails_table."""
    prs = Presentation(master_path)
    _, table_shape = avails_slide_and_table(prs)
    return int(table_shape.table.rows[1].height) if table_shape is not None else None


def template_font_pt(master_path):
    """The avails table's own authored font size, read the same way
    condense_avails_table itself reads it (max size across the template
    row's runs) -- what font_pt should come back as when the slide has
    plenty of room, rather than the 6pt the declared-row-height bug forced
    regardless of available space."""
    prs = Presentation(master_path)
    _, table_shape = avails_slide_and_table(prs)
    if table_shape is None:
        return None
    row = table_shape.table.rows[1]
    sizes = [run.font.size.pt for cell in row.cells
             for para in cell.text_frame.paragraphs
             for run in para.runs if run.font.size]
    return max(sizes) if sizes else None


def build_avails_deck(master_path, rows):
    selections = copy.deepcopy(assembly.SELECTIONS)
    fill_data = copy.deepcopy(assembly.FILL_DATA)
    fill_data["avails"]["rows"] = rows
    fill_data["avails"]["total_avails"] = "0"
    prs, _, _ = assembly.build_presentation(master_path, selections)
    warnings = list(assembly.personalize(prs, fill_data))
    return prs, warnings


def main():
    master_path, version_id, warning = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    if master_path is None:
        print(f"  SKIP  no master deck available ({warning})")
        return 0
    print(f"  ....  master deck version {version_id}")

    original_row_h = template_row_height(master_path)
    check("found the avails table in the untouched master", original_row_h is not None)
    if original_row_h is None:
        return 1
    print(f"  ....  template's own row height: {original_row_h / 914400:.3f}in")

    expected_font = template_font_pt(master_path)
    check("found the avails table's own authored font size", expected_font is not None)
    print(f"  ....  template's own font size: {expected_font}pt")

    three_rows = [
        {"audience": "Health & Fitness Intenders", "geo": "Washington, DC DMA", "avails": "1,250,000"},
        {"audience": "Auto Intenders", "geo": "Baltimore DMA", "avails": "480,000"},
        {"audience": "Home Improvement Intenders", "geo": "Richmond DMA", "avails": "220,000"},
    ]

    # An artificially generous floor -- far more slack than any real avails
    # slide's own layout happens to leave, isolating the growth cap itself
    # rather than depending on a specific document's geometry to trigger it.
    # Patched BEFORE personalize() runs, since that's where the deck's ONE
    # real call to condense_avails_table happens -- calling it a second time
    # afterward would size against an already-condensed table, not a fresh
    # one, which isn't how this pass is ever actually invoked.
    real_floor = assembly._content_floor
    assembly._content_floor = lambda slide, table_shape, ignore=None: (
        table_shape.top + Emu(int(6.5 * 914400)))
    try:
        prs, _ = build_avails_deck(master_path, three_rows)
    finally:
        assembly._content_floor = real_floor

    slide, table_shape = avails_slide_and_table(prs)
    check("found the avails table slide after assembly", table_shape is not None)
    if table_shape is None:
        return 1
    table = table_shape.table
    data_rows = list(table.rows)[1:]
    heights = [int(r.height) for r in data_rows]
    for i, h in enumerate(heights):
        print(f"  ....  row {i}: {h / 914400:.3f}in")

    # The font floor bug (#2): with a generous floor and nothing forcing a
    # shrink, the table should render at the template's OWN font -- not the
    # 6pt _max_font_for_row derived from the declared row height, which is
    # shorter than that same font actually needs. Read off the first data
    # cell's own run, the same way condense_avails_table reads it in.
    rendered_font = None
    first_cell = data_rows[0].cells[0]
    sizes = [run.font.size.pt for para in first_cell.text_frame.paragraphs
             for run in para.runs if run.font.size]
    if sizes:
        rendered_font = max(sizes)
    check("the table rendered at the template's own font, not shrunk to the floor",
          rendered_font == expected_font, f"got {rendered_font}pt, expected {expected_font}pt")

    # The growth cap (#1): a little slack (rounding, the header's own share)
    # is fine; a row visibly stretched to fill a generous floor is the bug
    # this guards. The cap is content-need at the CORRECT font, not the
    # template's tiny declared row height (which was only ever a valid
    # stand-in while bug #2 also held font_pt down near that same size) --
    # so the bound here is a generous, independent ceiling: two wrapped
    # lines at the template's own font plus cushion, not an exact replay of
    # condense_avails_table's own arithmetic.
    generous_ceiling = Emu(int((2 * expected_font * 1.2 / 72 + 0.15) * 914400))
    check("no data row grew past a couple of readable lines (the cap)",
          all(h <= generous_ceiling for h in heights),
          {"heights_in": [h / 914400 for h in heights], "ceiling_in": generous_ceiling / 914400})
    total_h = sum(heights) + int(table.rows[0].height)
    check("the table stays compact rather than stretching to fill the generous floor",
          total_h / 914400 < 3.5, f"{total_h / 914400:.3f}in")

    # Shrink-to-fit is unaffected: a table with too MANY rows for the space
    # still gets the overflow warning and clamps down, same as before --
    # this one runs against the REAL floor, no patching. The overflow
    # judgment (`bare` in condense_avails_table) is bare text height with
    # no cushion, by design -- see condense_avails_table's own comment --
    # so it takes a genuinely large row count to trip, not just "more than
    # a handful."
    many_rows = [dict(audience=f"Segment {i}", geo=f"Market {i} DMA", avails=str(i * 1000))
                for i in range(40)]
    _, warnings2 = build_avails_deck(master_path, many_rows)
    overflow_warnings = [w for w in warnings2 if w and "avails table" in w.lower()]
    check("shrink-to-fit still warns when a table genuinely can't fit",
          bool(overflow_warnings), warnings2)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("condense_avails_table caps growth at the template's own row height; "
          "shrink-to-fit is unaffected.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
