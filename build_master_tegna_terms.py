"""Proposal master deck: replace the plan slide's terms block with TEGNA's.

Source: TEGNA Standard Advertising Terms and Conditions -- Section 3.2
(Payment), 3.3 (Credit Terms) and 9.1 (Termination: 30 days' notice, the
campaign keeps running and the advertiser owes its fees for that period).

Rewrites the "Terms & Conditions" box on every media plan slide (every slide
with `{{PLAN_TITLE}}`: the standard plan and both Total TV variants). The
standard slide carried Premion's own terms and a premionmedia.com URL; the
Total TV slides carried an older TEGNA wording and a different bit.ly link;
all three said 10 days' notice. Each box keeps its three "Label: body"
paragraphs and its run formatting -- only the text changes -- and the URL
is a real hyperlink, cloned from the template's own link run so it keeps
the link styling. `assembly._TERMS_ANCHOR` ("terms & conditions") still
finds the box, so the co-viewing footnote lands where it always did.

The new wording is longer than the old, and the box can't grow upward (the
"Approved:" signature line sits directly above it). Rendered at the old 9pt
it ran off the bottom of the standard slide and, on the 11.28in-wide Total
TV boxes, under the station logo. So the box is set to 8pt and one width,
BOX_WIDTH, which ends before the Total TV station logos (10.87in from the
left edge) -- three or four lines that clear the slide's bottom edge.

Run against the CURRENT active master (version 10 at the time of writing):
    python build_master_tegna_terms.py <active master>.pptx <out>.pptx
then upload the output on dev only (never activate from dev -- the nightly
merge does that).
"""
import copy
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.text.text import _Run
from pptx.util import Inches, Pt

import assembly
import slide_map

FONT_SIZE = Pt(8)
BOX_WIDTH = Inches(10.40)
TERMS_URL = "https://bit.ly/TegnaStandardAdTerms"
TERMS = [
    ("TEGNA Terms & Conditions: ",
     "Your purchase of advertising is subject to the TEGNA Standard Advertising "
     "Terms and Conditions (“Standard Terms”), available at: "),
    ("Payment: ",
     "Payment in full is due no later than 5 business days before the campaign "
     "start date, unless TEGNA has granted you credit terms, in which case "
     "payment is due within 30 days of invoice."),
    ("Cancellation: ",
     "You may cancel an Insertion Order at any time with 30 days’ prior "
     "written notice. Your campaign will continue to run during the notice "
     "period, and you are responsible for all fees for that period."),
]


def terms_box(slide):
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.lower()
        if "terms & conditions" in text and "cancellation" in text:
            return shape
    return None


def rewrite(slide, box):
    paragraphs = list(box.text_frame.paragraphs)
    if len(paragraphs) != len(TERMS) or any(len(p.runs) < 2 for p in paragraphs):
        raise SystemExit(f"unexpected terms box shape: {[p.text for p in paragraphs]}")

    first = paragraphs[0]
    link_template = next((r for r in first.runs if r.hyperlink.address), None)
    if link_template is None:
        raise SystemExit("terms box has no hyperlink run to clone the link styling from")
    new_link = copy.deepcopy(link_template._r)
    rpr = new_link.find(qn("a:rPr"))
    for click in rpr.findall(qn("a:hlinkClick")):   # the old rId belongs to the original
        rpr.remove(click)

    for paragraph, (label, body) in zip(paragraphs, TERMS):
        runs = paragraph.runs
        runs[0].text, runs[1].text = label, body
        for extra in runs[2:]:
            extra.hyperlink.address = None             # drops the old link's relationship
            extra._r.getparent().remove(extra._r)

    runs = first.runs
    runs[-1]._r.addnext(new_link)
    link = _Run(new_link, first)
    link.text = TERMS_URL
    link.hyperlink.address = TERMS_URL

    for paragraph in box.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = FONT_SIZE
        for end in paragraph._p.findall(qn("a:endParaRPr")):
            end.set("sz", str(int(FONT_SIZE.pt * 100)))
    box.width = BOX_WIDTH
    box.height = assembly._measured_frame_height(box, slide)


def main(src, out):
    prs = Presentation(src)
    done = 0
    for slide in prs.slides:
        shapes = list(slide_map.iter_all_shapes(slide.shapes))
        if not any(sh.has_text_frame and "{{PLAN_TITLE}}" in sh.text_frame.text for sh in shapes):
            continue
        box = terms_box(slide)
        if box is None:
            raise SystemExit(f"slide {prs.slides.index(slide) + 1}: no terms box found")
        if TERMS_URL in box.text_frame.text:
            continue                                   # idempotent: already built
        rewrite(slide, box)
        done += 1
        print(f"slide {prs.slides.index(slide) + 1} ({slide_map.notes_key(slide)}): terms replaced")
    prs.save(out)
    print(f"{done} media plan slide(s) updated -> {out}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
