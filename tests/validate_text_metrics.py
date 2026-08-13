"""Check the wrap prediction against what PowerPoint actually laid out.

    python tests/validate_text_metrics.py

PowerPoint reports a row's height as max(declared height, height the content
needs). That asymmetry is the whole test, and it's easy to misread in both
directions:

  * A row PowerPoint reports at its declared height tells you the content fit.
    It does NOT tell you how many lines were drawn -- the sizer grows rows to
    fill the slide, so a one-line totals row can sit in half an inch. Deriving
    a line count from that height and calling it ground truth measures the
    sizer's own output, not the renderer's.
  * A row PowerPoint reports TALLER than declared is the real signal: the text
    didn't fit and PowerPoint grew the row to make it. That growth is exactly
    what pushes a table over the block beneath it, so it's the failure.

So the check is: no sized row may grow, and the table as a whole must still
clear its floor once PowerPoint has had its say. Rows that did grow are
reported with what the measurer predicted, since those are the cells where
the prediction can be compared against a real content height.

Needs Windows with PowerPoint (see deck_render.py); prints SKIP otherwise.
"""

import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import assembly                                  # noqa: E402
import deck_render                               # noqa: E402
import text_metrics                              # noqa: E402
from pptx import Presentation                    # noqa: E402
from pptx.util import Emu                        # noqa: E402

RENDERED = deck_render.render_root()
EMU_PER_POINT = assembly._EMU_PER_POINT

# PowerPoint's own rounding. A row reported a hair over its declared height
# hasn't grown, it's been through a float conversion.
GROWTH_TOLERANCE_PT = 1.0


def powerpoint_geometry(pptx_path):
    """{(slide, table): {"rows": [heights pt], "top": pt}} as laid out."""
    import pythoncom
    import win32com.client

    def walk(shapes):
        for shape in shapes:
            if shape.Type == 6:      # msoGroup
                for inner in walk(shape.GroupItems):
                    yield inner
            else:
                yield shape

    # Whether PowerPoint was already up decides whether we may close it
    # afterwards. Dispatch attaches to a running instance rather than starting
    # its own, so quitting one we didn't start shuts down an application
    # somebody is using -- this validator ran while a slide show of a client
    # deck was up, and an unconditional Quit() would have ended it. Same guard
    # deck_render.render_deck already carries, for the same reason.
    borrowed = deck_render.powerpoint_is_running()

    pythoncom.CoInitialize()
    powerpoint = presentation = None
    geometry = {}
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(
            str(Path(pptx_path).resolve()), ReadOnly=True, WithWindow=False)
        def signature(slide_index, table):
            """Identify a table by its shape and first row rather than by its
            position in the shape list. COM enumerates only top-level shapes
            while python-pptx recurses into groups, so the two orderings
            drift apart on any slide holding a grouped table -- and pairing
            by index then compares one table's prediction against another
            table's heights, which is how a header row with one 12pt line in
            it appeared to have been drawn at two."""
            first = " ".join(table.Cell(1, c).Shape.TextFrame.TextRange.Text.strip()
                             for c in range(1, table.Columns.Count + 1))
            return (slide_index, table.Rows.Count, table.Columns.Count,
                    " ".join(first.split()))

        for slide_index in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(slide_index)
            for shape in walk(slide.Shapes):
                if not shape.HasTable:
                    continue
                table = shape.Table
                geometry.setdefault(signature(slide_index, table), []).append({
                    "rows": [table.Rows(r).Height
                             for r in range(1, table.Rows.Count + 1)],
                    "top": shape.Top,
                })
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        finally:
            if powerpoint is not None and not borrowed:
                try:
                    powerpoint.Quit()
                except Exception:                                # noqa: BLE001
                    pass
            pythoncom.CoUninitialize()
    return geometry


def sized_table(slide, shape):
    """Is this the table the sizer actually laid out on this slide?

    Asked by looking for the sizer's own fingerprint rather than by sniffing
    slide text. Guessing from the header let in the case study slides --
    copied whole from vault decks, with their own fonts and merged cells that
    nothing here touches -- and holding this code to someone else's layout
    reported failures it had no part in. Asking whether the slide says "MEDIA
    PLAN" was the next attempt and was wrong in the other direction: the
    master's static "TV SCHEDULE PLACEHOLDER" slide says it too and carries a
    table this code never touches, so every Total TV deck reported a phantom
    grown row on it. A permanent failure nobody can act on is worse than no
    check -- it teaches people to skip the output.

    The fingerprint: condense_media_plan_table zeroes margin_top and
    margin_bottom on every cell of every table it sizes, and nothing else in
    this project does. A template table still carries the default insets
    (0.05in), and an unfilled one was never sized at all.
    """
    managed = assembly._find_table_shape(slide)
    if managed is None or managed._element is not shape._element:
        return False
    if "{{" in assembly.slide_map.extract_slide_text(slide):
        return False        # never filled, so never sized
    return all(cell.margin_top == 0 and cell.margin_bottom == 0
               for row in shape.table.rows for cell in row.cells)


def row_font(table, row_index):
    sizes = [run.font.size.pt
             for cell in table.rows[row_index].cells
             for para in cell.text_frame.paragraphs
             for run in para.runs if run.font.size]
    return max(sizes) if sizes else None


def predicted_lines(table, row_index, font_pt):
    """Lines the measurer predicts for the tallest cell in this row."""
    most = 1
    for column, cell in enumerate(table.rows[row_index].cells):
        if column >= len(table.columns):
            continue
        lines = sum(
            assembly._wrapped_lines(
                "".join(r.text for r in para.runs),
                table.columns[column].width, font_pt,
                assembly._resolve_typeface(
                    next((r.font.name for r in para.runs if r.font.name), None),
                    table),
                assembly._side_insets(cell))
            for para in cell.text_frame.paragraphs)
        most = max(most, max(1, lines))
    return most


def check_deck(name, pptx_path, declared_by_deck=None):
    prs = Presentation(str(pptx_path))
    if declared_by_deck is not None:
        declared_by_deck[name] = assembly.declared_fonts(prs)
    assembly._THEME_CACHE.clear()
    laid_out = powerpoint_geometry(pptx_path)

    rows_checked = grown = 0
    reports = []
    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape in assembly.slide_map.iter_all_shapes(slide.shapes):
            if not shape.has_table:
                continue
            table = shape.table
            assembly.register_theme_for_table(table, prs, slide=slide)
            if not sized_table(slide, shape):
                continue
            key = (slide_index, len(table.rows), len(table.columns),
                   " ".join(" ".join(c.text for c in table.rows[0].cells).split()))
            matches = laid_out.get(key) or []
            if len(matches) != 1:
                print(f"    slide {slide_index}: could not pair this table with "
                      f"PowerPoint's ({len(matches)} candidates) -- not checked")
                continue
            actual = matches[0]

            for row_index in range(min(len(table.rows), len(actual["rows"]))):
                declared = table.rows[row_index].height / EMU_PER_POINT
                drawn = actual["rows"][row_index]
                rows_checked += 1
                if drawn <= declared + GROWTH_TOLERANCE_PT:
                    continue
                grown += 1
                font_pt = row_font(table, row_index) or 0
                predicted = predicted_lines(table, row_index, font_pt) if font_pt else 0
                needed = ((drawn - 2 * assembly._DEFAULT_CELL_INSET / EMU_PER_POINT)
                          / (font_pt * assembly._LINE_SPACING)) if font_pt else 0
                text = " | ".join(" ".join(c.text.split())[:22]
                                  for c in table.rows[row_index].cells)[:86]
                reports.append(
                    f"    slide {slide_index} row {row_index} at {font_pt:g}pt: "
                    f"declared {declared:.1f}pt, PowerPoint grew it to {drawn:.1f}pt\n"
                    f"        predicted {predicted} line(s), content needed "
                    f"~{needed:.1f}\n        {text}")

            # Growth is only a problem where it costs clearance, so say what
            # it cost: the bottom the slide ends up with against the floor
            # the sizer was working to.
            bottom = actual["top"] + sum(actual["rows"])
            floor_emu = assembly._content_floor(slide, shape)
            if floor_emu is None:
                continue
            floor = (floor_emu - assembly._TABLE_CLEARANCE) / EMU_PER_POINT
            if bottom > floor + GROWTH_TOLERANCE_PT:
                reports.append(
                    f"    slide {slide_index}: table bottom {bottom / 72:.2f}in "
                    f"is past its floor of {floor / 72:.2f}in -- OVERLAP")

    print(f"\n=== {name} ===")
    print(f"  {rows_checked} sized rows; {grown} grew past their declared height")
    for report in reports[:10]:
        print(report)
    if len(reports) > 10:
        print(f"    ... and {len(reports) - 10} more")
    return rows_checked, grown


def main():
    if not deck_render.renderer_available():
        print("SKIP -- needs Windows with PowerPoint and pywin32.")
        return 0
    decks = sorted(RENDERED.glob("*/*.pptx"))
    if not decks:
        print("No rendered decks. Run tests/render_scenarios.py first.")
        return 2
    checked = grown = 0
    declared_by_deck = {}
    for deck in decks:
        try:
            a, b = check_deck(deck.parent.name, deck, declared_by_deck)
        except Exception as exc:                                 # noqa: BLE001
            print(f"\n=== {deck.parent.name} ===\n  FAILED: {type(exc).__name__}: {exc}")
            return 1
        checked, grown = checked + a, grown + b
    # What the numbers above were measured IN. A row that didn't grow proves
    # nothing if the measurement used a font the deck never draws -- that is
    # exactly how the Calibri substitution survived so long, reported as a
    # clean run every time.
    print("\n" + "=" * 70)
    declared = sorted({name for names in declared_by_deck.values() for name in names})
    if declared:
        missing = [n for n in declared if not text_metrics.is_available(n)]
        print(f"Fonts the deck declares ({len(declared)}), "
              f"{len(declared) - len(missing)} available here:")
        for name in declared:
            print(f"  {'ok ' if name not in missing else '!! '}{name}")
    else:
        print("No deck declared any embedded fonts (saved without "
              "'Embed fonts in the file').")

    print("\nFonts used for measurement:")
    for entry in text_metrics.resolution_report():
        mark = "  ok " if entry["status"] == "exact" else "  !! "
        bold = " bold" if entry["bold"] else ""
        print(f"{mark}{(entry['requested'] or '?')}{bold} -> "
              f"{entry['resolved'] or 'CHARACTER-WIDTH ESTIMATE'}"
              f"{('  [' + entry['detail'] + ']') if entry['detail'] else ''}")
    note = text_metrics.measurement_note()
    if note:
        print(f"\n  {note}")

    print("\n" + "=" * 70)
    print(f"{checked} sized rows across {len(decks)} decks: "
          f"{grown} grew past the height reserved for them.")
    print("A row that grows is a row whose text was under-measured, and that "
          "growth is what\npushes a table over the block beneath it.")
    print("=" * 70)
    return 1 if grown else 0


if __name__ == "__main__":
    sys.exit(main())
