"""REPORT_MASTER_v0_19 -- the auto group review follow-up (2026-11).

Run against the CURRENT active template (v0_18):
    python build_v0_19.py REPORT_MASTER_v0_18.pptx REPORT_MASTER_v0_19.pptx
then upload the output on dev only, inactive (the nightly merge activates it).

  * report:site_comparison -- "Where Visitors Went: Group Site vs. Store
    Site", after the audience comparison, added when the attribution
    export's Advertiser tab lists two or more pixels (a dealer group's site
    and the store's own): a per-site table (visitors, page visits,
    attributed impressions, the deduplicated total), a footnote with the
    implied overlap, a unique-visitors bar chart and the narrative panel.
    Cloned from report:attribution_breakdown, like v0_18's two slides, so
    the type, colours and table are the template's own. Optional.
"""
import copy
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches

SRC = sys.argv[1] if len(sys.argv) > 1 else "REPORT_MASTER_v0_18.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "REPORT_MASTER_v0_19.pptx"
prs = Presentation(SRC)
# The helpers are build_v0_18.py's own (that script runs on import).
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


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
after = ("report:audience_comparison" if any("key: report:audience_comparison" in key_of(s)
                                             for s in prs.slides) else "report:attribution_breakdown")
site = S(clone_after(brk_src, after, "key: report:site_comparison"))
site.rename("Text 1", "SiteTitle")
text(site.get("SiteTitle"), "{{SITE_TITLE}}")
text(site.get("Text 2"), "{{SITE_HEADLINE}}")
site.rename("Text 3", "SiteTableHeader")
text(site.get("SiteTableHeader"), "ATTRIBUTED VISITORS BY SITE")
site.rename("BreakdownTable", "SiteTable")
recolumn(site.get("SiteTable"), ["Site", "Visitors", "Page visits", "Attr. impr."],
         [2.90, 1.00, 1.10, 1.10], "{{SITE_ROWS}}")
site.rename("Text 4", "SiteChartHeader")
text(site.get("SiteChartHeader"), "UNIQUE VISITORS BY SITE")
site.rename("ChartRegion", "SiteChartRegion")
site.rename("ChartRegionLabel", "SiteChartRegionLabel")
text(site.get("SiteChartRegionLabel"), "SiteChartRegion · unique visitors by site")
foot = site.borrow(url, "UrlReachFootnote", "SiteFootnote")
text(foot, "{{SITE_FOOTNOTE}}")
place(foot, 0.50, 3.95, 6.10, 0.30)
site.rename("AttributionNarrative", "SiteNarrative")
text(site.get("SiteNarrative"), "{{SITE_NARRATIVE}}")

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
print([key_of(x).split("\n")[0].replace("key: report:", "") for x in prs.slides])
