"""REPORT_MASTER_v0_18 -- the multi-audience bank review (2026-11).

Run against the CURRENT active template (v0_17):
    python build_v0_18.py REPORT_MASTER_v0_17.pptx REPORT_MASTER_v0_18.pptx
then upload the output on dev only, inactive (the nightly merge activates it).

Both new slides are cloned from report:attribution_breakdown (table left or
full width, chart right, narrative panel along the bottom), so the type,
colours and tables are the template's own. Both are optional.

  * report:key_pages -- "Where Visitors Went" as "Key pages and why they
    matter": one full-width table (page group | page visits | why it
    matters), a footnote and the narrative panel. Used instead of
    report:url_report in an ordinary report; report:url_report stays for the
    Auto-Sales Analyst version and "Compare locations".
  * report:audience_comparison -- after the attribution breakdown, added when
    the export carries 2+ audiences or RFPIDs: a metrics x audiences table
    (up to three; unused audience columns are removed in code), a share of
    delivery vs. share of response chart, a footnote and the panel.
  * report:recap -- ExportDateNote under the cards: when the data was pulled,
    and whether the latest weeks are still accruing.
"""
import copy
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

SRC = sys.argv[1] if len(sys.argv) > 1 else "REPORT_MASTER_v0_17.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "REPORT_MASTER_v0_18.pptx"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
prs = Presentation(SRC)


def key_of(s): return s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
def find(key): return [s for s in prs.slides if f"key: {key}" in key_of(s)][0]


def clone_after(src, after_key, notes):
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes): sh._element.getparent().remove(sh._element)
    rid = {r.rId: new.part.relate_to(r.target_part, RT.IMAGE)
           for r in src.part.rels.values() if r.reltype == RT.IMAGE}
    for sh in src.shapes:
        el = copy.deepcopy(sh._element)
        for node in el.iter():
            for attr in ("{%s}embed" % NS["r"], "{%s}link" % NS["r"]):
                if attr in node.attrib and node.attrib[attr] in rid:
                    node.attrib[attr] = rid[node.attrib[attr]]
        new.shapes._spTree.append(el)
    new.notes_slide.notes_text_frame.text = notes
    lst = prs.slides._sldIdLst; last = list(lst)[-1]
    idx = [i for i, s in enumerate(prs.slides) if f"key: {after_key}" in key_of(s)][0]
    lst.remove(last); lst.insert(idx + 1, last)
    return new


class S:
    def __init__(self, s): self.s = s
    def get(self, n):
        for sh in self.s.shapes:
            if sh.name == n: return sh
        raise KeyError(n)
    def drop(self, *ns):
        for n in ns:
            sh = self.get(n); sh._element.getparent().remove(sh._element)
    def clone(self, n, new):
        el = copy.deepcopy(self.get(n)._element); self.s.shapes._spTree.append(el)
        sh = self.s.shapes[-1]; sh.name = new; return sh
    def rename(self, old, new): self.get(old).name = new
    def borrow(self, other, n, new):
        """A plain shape (no relationships) copied in from another slide."""
        el = copy.deepcopy(other.get(n)._element); self.s.shapes._spTree.append(el)
        sh = self.s.shapes[-1]; sh.name = new; return sh
    def to_front(self, n): self.s.shapes._spTree.append(self.get(n)._element)


def text(sh, t):
    tf = sh.text_frame
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = t
        for r in p.runs[1:]: r._r.getparent().remove(r._r)
    else: p.add_run().text = t
    for e in tf.paragraphs[1:]: e._p.getparent().remove(e._p)


def cell(c, t):
    p = c.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = t
        for r in p.runs[1:]: r._r.getparent().remove(r._r)
    elif t: p.add_run().text = t


def place(sh, x, y, w=None, h=None):
    sh.left, sh.top = Inches(x), Inches(y)
    if w is not None: sh.width = Inches(w)
    if h is not None: sh.height = Inches(h)


def recolumn(tsh, heads, widths, token, left_cols=1):
    tbl, el = tsh.table, tsh.table._tbl
    grid = el.find("a:tblGrid", NS)
    while len(grid.findall("a:gridCol", NS)) < len(heads):
        grid.append(copy.deepcopy(grid.findall("a:gridCol", NS)[-1]))
        for tr in el.findall("a:tr", NS): tr.append(copy.deepcopy(tr.findall("a:tc", NS)[-1]))
    while len(grid.findall("a:gridCol", NS)) > len(heads):
        grid.remove(grid.findall("a:gridCol", NS)[-1])
        for tr in el.findall("a:tr", NS): tr.remove(tr.findall("a:tc", NS)[-1])
    for gc, w in zip(grid.findall("a:gridCol", NS), widths): gc.set("w", str(Inches(w)))
    tsh.width = Inches(sum(widths))
    for c, h in zip(list(tbl.rows[0].cells), heads): cell(c, h)
    row = list(tbl.rows[1].cells); cell(row[0], token)
    for c in row[1:]: cell(c, "")
    for r in tbl.rows:
        for i, c in enumerate(r.cells):
            c.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT if i < left_cols else PP_ALIGN.RIGHT




url = S(find("report:url_report"))
brk_src = find("report:attribution_breakdown")

# ======================================================================
# 1. report:key_pages -- "Key pages and why they matter"
# ======================================================================
kp = S(clone_after(brk_src, "report:url_report", "key: report:key_pages"))
text(kp.get("Text 1"), "Where Visitors Went")
text(kp.get("Text 2"), "{{KEY_PAGES_HEADLINE}}")
kp.drop("ChartRegion", "ChartRegionLabel", "Text 4")
kp.rename("Text 3", "KeyPagesHeader")
text(kp.get("KeyPagesHeader"), "KEY PAGES AND WHY THEY MATTER")
place(kp.get("KeyPagesHeader"), 0.50, 2.10, 12.33, 0.26)
kp.rename("BreakdownTable", "KeyPagesTable")
recolumn(kp.get("KeyPagesTable"), ["Page group", "Page visits", "Why it matters"],
         [3.30, 1.25, 7.78], "{{KEY_PAGE_ROWS}}", left_cols=1)
table = kp.get("KeyPagesTable").table
for r in table.rows:
    r.cells[2].text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
place(kp.get("KeyPagesTable"), 0.50, 2.45)
foot = kp.borrow(url, "UrlReachFootnote", "KeyPagesFootnote")
text(foot, "{{KEY_PAGES_FOOTNOTE}}")
place(foot, 0.50, 5.38, 12.33, 0.25)
kp.rename("AttributionNarrative", "KeyPagesNarrative")
text(kp.get("KeyPagesNarrative"), "{{KEY_PAGES_NARRATIVE}}")

# ======================================================================
# 2. report:audience_comparison -- audiences side by side
# ======================================================================
au = S(clone_after(brk_src, "report:attribution_breakdown", "key: report:audience_comparison"))
text(au.get("Text 1"), "How the Audiences Compare")
text(au.get("Text 2"), "{{AUDIENCE_HEADLINE}}")
au.rename("Text 3", "AudienceTableHeader")
text(au.get("AudienceTableHeader"), "HOW THE AUDIENCES COMPARE")
au.rename("BreakdownTable", "AudienceComparisonTable")
recolumn(au.get("AudienceComparisonTable"), ["", "{{AUDIENCE_0}}", "{{AUDIENCE_1}}", "{{AUDIENCE_2}}"],
         [1.90, 1.40, 1.40, 1.40], "{{AUDIENCE_ROWS}}")
au.rename("Text 4", "AudienceChartHeader")
text(au.get("AudienceChartHeader"), "SHARE OF DELIVERY VS. SHARE OF RESPONSE")
au.rename("ChartRegion", "AudienceChartRegion")
au.rename("ChartRegionLabel", "AudienceChartRegionLabel")
text(au.get("AudienceChartRegionLabel"), "AudienceChartRegion · share of impressions vs. attributed")
foot = au.borrow(url, "UrlReachFootnote", "AudienceFootnote")
text(foot, "{{AUDIENCE_FOOTNOTE}}")
place(foot, 0.50, 5.40, 6.10, 0.25)
au.rename("AttributionNarrative", "AudienceNarrative")
text(au.get("AudienceNarrative"), "{{AUDIENCE_NARRATIVE}}")

# ======================================================================
# 3. report:recap -- the export date, under the cards ("Data as of Nov 4,
#    2026"): a report pulled before every visit is in says so.
# ======================================================================
recap = S(find("report:recap"))
note = recap.clone("PolkSalesWindowNote", "ExportDateNote")
text(note, "{{EXPORT_DATE_NOTE}}")
place(note, 0.50, 7.00, 12.33, 0.26)

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
print([key_of(x).split("\n")[0].replace("key: report:", "") for x in prs.slides])
