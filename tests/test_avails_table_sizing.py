"""condense_avails_table's growth cap (assembly.py) -- offline, no PowerPoint
needed.

A live proposal (Plaza Motors, three avails rows) rendered its avails table
at 3.82in tall: condense_avails_table shrinks to fit but, like the media
plan table, also hands out leftover slide space evenly across rows so a
short table doesn't sit tiny in a mostly-empty region -- and for a table
this short, with a floor this far down the slide, that grew three ordinary
rows into something visibly stretched. The fix caps a row's grown height at
the template's OWN row height (or the row's real content need, if that's
taller -- shrink-to-fit is unaffected) instead of handing out the full
slack. `tests/test_group_scenarios.py --render` covers real documents
end-to-end over COM; this isolates the growth cap itself against an
artificially generous floor, which the real fixtures don't happen to
produce (their floors leave less slack than this).
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
    already been through condense_avails_table, which is exactly the
    height this test asserts a small table's rows shouldn't grow past."""
    prs = Presentation(master_path)
    _, table_shape = avails_slide_and_table(prs)
    return int(table_shape.table.rows[1].height) if table_shape is not None else None


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

    # A little slack (rounding, the header's own share) is fine; a row
    # visibly stretched to fill a generous floor is the bug this guards.
    tolerance = Emu(int(0.05 * 914400))
    check("no data row grew past the template's own row height (the cap)",
          all(h <= original_row_h + tolerance for h in heights),
          [h / 914400 for h in heights])
    total_h = sum(heights) + int(table.rows[0].height)
    check("the 3-row table stays compact rather than stretching to fill the generous floor",
          total_h / 914400 < 3.0, f"{total_h / 914400:.3f}in")

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
