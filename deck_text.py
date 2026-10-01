"""A generated deck's slide-by-slide text -- titles, text boxes, tables --
so a tester can check slide content without opening PowerPoint. Pure: no
Streamlit, no database. Image-only slides (an appended Analyst deck's charts)
show whatever text they carry and say when they have none."""
from pptx import Presentation

import slide_map


def deck_text(pptx_path):
    """Markdown text, one section per slide, in deck order."""
    prs = Presentation(pptx_path)
    out = []
    for number, slide in enumerate(prs.slides, start=1):
        key = slide_map.notes_key(slide) if slide.has_notes_slide else None
        blocks = []
        for shape in slide_map.iter_all_shapes(slide.shapes):
            if getattr(shape, "has_table", False) and shape.has_table:
                rows = [[cell.text.strip().replace("|", "\\|").replace("\n", " ")
                         for cell in row.cells] for row in shape.table.rows]
                if rows:
                    blocks.append("| " + " | ".join(rows[0]) + " |")
                    blocks.append("|" + "---|" * len(rows[0]))
                    blocks += ["| " + " | ".join(r) + " |" for r in rows[1:]]
                    blocks.append("")
            elif shape.has_text_frame and shape.text_frame.text.strip():
                blocks.append(shape.text_frame.text.strip())
        title = blocks[0] if blocks else "(no text on this slide)"
        header = f"### Slide {number}: {title.splitlines()[0]}"
        if key:
            header += f"  `{key}`"
        out += [header, ""] + (blocks[1:] if blocks else []) + [""]
    return "\n".join(out)
