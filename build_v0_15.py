"""REPORT_MASTER_v0_15 — three native Auto-Sales Analyst slides.

Run against the CURRENT active template (v0_14):
    python build_v0_15.py REPORT_MASTER_v0_14.pptx REPORT_MASTER_v0_15.pptx
All three slides are cloned in-deck from report:automotive_registrations, so they
inherit the deck's header band, logo and fonts. Each carries "analyst_set: true".
"""
import copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

SRC = sys.argv[1] if len(sys.argv) > 1 else "REPORT_MASTER_v0_14.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "REPORT_MASTER_v0_15.pptx"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
MUTED = RGBColor(0x6B, 0x72, 0x80)
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
            try: sh = self.get(n); sh._element.getparent().remove(sh._element)
            except KeyError: pass
    def clone(self, n, new):
        el = copy.deepcopy(self.get(n)._element); self.s.shapes._spTree.append(el)
        sh = self.s.shapes[-1]; sh.name = new; return sh
    def rename(self, old, new): self.get(old).name = new

def text(sh, t):
    tf = sh.text_frame; tf.word_wrap = True
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

def style(sh, size, color=None, bold=None, align=None):
    for p in sh.text_frame.paragraphs:
        if align is not None: p.alignment = align
        for r in p.runs:
            r.font.size = Pt(size)
            if color is not None: r.font.color.rgb = color
            if bold is not None: r.font.bold = bold

def recolumn(tsh, heads, widths, token, left_cols=1, center_last=False):
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
    n = len(heads)
    for r in tbl.rows:
        for i, c in enumerate(r.cells):
            a = PP_ALIGN.LEFT if i < left_cols else PP_ALIGN.RIGHT
            if center_last and i == n - 1: a = PP_ALIGN.CENTER
            c.text_frame.paragraphs[0].alignment = a

POLK_TILES = ("PolkSalesTile", "PolkMsrpTile", "PolkLiftTile", "PolkRoiTile")

def tile_row(s, names, y, h=0.72, x0=0.50, span=12.33, gap=0.25, size=20):
    tw = (span - gap * (len(names) - 1)) / len(names)
    for i, n in enumerate(names):
        x = x0 + i * (tw + gap)
        place(s.get(n), x, y, tw, h)
        place(s.get(n + "Value"), x + 0.10, y + 0.05, tw - 0.20, 0.40)
        place(s.get(n + "Label"), x + 0.10, y + 0.45, tw - 0.20, 0.22)
        style(s.get(n + "Value"), size)

def footnote(s, name, token):
    fn = s.clone("PolkMatchRateNote", name); text(fn, token)
    place(fn, 0.50, 6.72, 12.33, 0.30); style(fn, 9, MUTED, False, PP_ALIGN.LEFT)

src = find("report:automotive_registrations")

# ======================================================================
# 1. report:analyst_inventory — "Inventory Movement"
# ======================================================================
inv = S(clone_after(src, "report:automotive_registrations",
                    "key: report:analyst_inventory\nanalyst_set: true"))
text(inv.get("Text 1"), "Inventory Movement")
text(inv.get("Text 2"), "{{ANALYST_INVENTORY_HEADLINE}}")

# Row 1: Revenue | Units | Sold-30+ (v2) | Pipeline — Units and the influence
# tile sit between Revenue and Pipeline so the two dollar figures never touch.
# Row 2: Look-to-Book | New | Used
plan = [("PolkMsrpTile",  "AnalystRevenueTile",   "{{ANALYST_REVENUE_SOLD}}",        "Est. value of shopped vehicles sold"),
        ("PolkSalesTile", "AnalystUnitsTile",     "{{ANALYST_UNITS_SOLD}}",          "Shopped vehicles sold"),
        ("PolkLiftTile",  "AnalystInfluenceTile", "{{ANALYST_SOLD_ABOVE_BENCHMARK}}", "Sold with 30+ campaign visits"),
        ("PolkRoiTile",   "AnalystPipelineTile",  "{{ANALYST_PIPELINE_VALUE}}",      "Pipeline value")]
for old, new, tok, lab in plan:
    for suf in ("", "Value", "Label"): inv.rename(old + suf, new + suf)
    text(inv.get(new + "Value"), tok); text(inv.get(new + "Label"), lab)
for new, tok, lab in (("AnalystLtbTile", "{{ANALYST_LTB}}", "Look-to-Book"),
                      ("AnalystLtbNewTile", "{{ANALYST_LTB_NEW}}", "Look-to-Book, new"),
                      ("AnalystLtbUsedTile", "{{ANALYST_LTB_USED}}", "Look-to-Book, used")):
    for suf in ("", "Value", "Label"): inv.clone("AnalystUnitsTile" + suf, new + suf)
    text(inv.get(new + "Value"), tok); text(inv.get(new + "Label"), lab)
tile_row(inv, ["AnalystRevenueTile", "AnalystUnitsTile", "AnalystInfluenceTile", "AnalystPipelineTile"], 2.05)
tile_row(inv, ["AnalystLtbTile", "AnalystLtbNewTile", "AnalystLtbUsedTile"], 2.87)
inv.drop("PolkMatchRateNote")  # cloned from below, so drop after footnote is made
# left: models table
inv.rename("PolkTargetDealersHeader", "AnalystModelsHeader")
text(inv.get("AnalystModelsHeader"), "SHOPPED AND SOLD BY MODEL")
place(inv.get("AnalystModelsHeader"), 0.50, 3.78, 7.70, 0.26)
inv.rename("PolkTargetDealersTable", "AnalystModelsTable")
recolumn(inv.get("AnalystModelsTable"), ["Model", "Shopped", "Sold", "Look-to-Book"],
         [3.70, 1.30, 1.30, 1.40], "{{ANALYST_MODEL_ROWS}}")
place(inv.get("AnalystModelsTable"), 0.50, 4.08)
# right: tier table + narrative
th = inv.clone("AnalystModelsHeader", "AnalystTierHeader")
text(th, "WHAT SOLD BY PRICE TIER"); place(th, 8.45, 3.78, 4.38, 0.26)
tt = inv.clone("AnalystModelsTable", "AnalystTierTable")
recolumn(tt, ["Tier", "Units sold", "% of sold", "% of shopped"],
         [1.18, 1.00, 1.00, 1.20], "{{ANALYST_TIER_ROWS}}")
place(tt, 8.45, 4.08)
inv.rename("PolkNarrative", "AnalystInventoryNarrative")
text(inv.get("AnalystInventoryNarrative"), "{{ANALYST_INVENTORY_NARRATIVE}}")
place(inv.get("AnalystInventoryNarrative"), 8.45, 5.38, 4.38, 1.28)

# ======================================================================
# 2. report:analyst_watchlist — "Missed Opportunities"
# ======================================================================
wl = S(clone_after(src, "report:analyst_inventory",
                   "key: report:analyst_watchlist\nanalyst_set: true"))
text(wl.get("Text 1"), "Missed Opportunities")
text(wl.get("Text 2"), "{{ANALYST_WATCHLIST_HEADLINE}}")
wl.rename("PolkTargetDealersHeader", "AnalystWatchlistHeader")
text(wl.get("AnalystWatchlistHeader"), "HIGH-INTEREST VEHICLES STILL ON THE LOT")
place(wl.get("AnalystWatchlistHeader"), 0.50, 2.15, 7.70, 0.26)
wl.rename("PolkTargetDealersTable", "AnalystWatchlistTable")
recolumn(wl.get("AnalystWatchlistTable"), ["Vehicle", "New/Used", "Visits", "Est. price"],
         [4.00, 1.20, 1.10, 1.40], "{{ANALYST_WATCHLIST_ROWS}}")
place(wl.get("AnalystWatchlistTable"), 0.50, 2.45)
wl.rename("PolkNarrative", "AnalystWatchlistNarrative")
text(wl.get("AnalystWatchlistNarrative"), "{{ANALYST_WATCHLIST_NARRATIVE}}")
place(wl.get("AnalystWatchlistNarrative"), 8.45, 2.15, 4.38, 4.40)

# ======================================================================
# 3. report:analyst_group — "Store Scoreboard"
# ======================================================================
gp = S(clone_after(src, "report:analyst_watchlist",
                   "key: report:analyst_group\nanalyst_set: true"))
text(gp.get("Text 1"), "Store Scoreboard")
text(gp.get("Text 2"), "{{ANALYST_GROUP_HEADLINE}}")
gp.rename("PolkTargetDealersHeader", "AnalystGroupHeader")
text(gp.get("AnalystGroupHeader"), "STORE-BY-STORE PERFORMANCE")
place(gp.get("AnalystGroupHeader"), 0.50, 2.15, 12.33, 0.26)
gp.rename("PolkTargetDealersTable", "AnalystGroupTable")
recolumn(gp.get("AnalystGroupTable"),
         ["Store", "Traffic", "VDPs", "Sold", "Look-to-Book", "Est. revenue sold", ""],
         [4.33, 1.30, 1.20, 1.20, 1.50, 1.90, 0.90], "{{ANALYST_GROUP_ROWS}}",
         center_last=True)
place(gp.get("AnalystGroupTable"), 0.50, 2.45)
gp.rename("PolkNarrative", "AnalystGroupNarrative")
text(gp.get("AnalystGroupNarrative"), "{{ANALYST_GROUP_NARRATIVE}}")
place(gp.get("AnalystGroupNarrative"), 0.50, 4.90, 12.33, 1.20)

# footnotes (cloned from the match-rate note for its text styling), then strip
# every Polk-only shape still left on the two non-inventory slides
for s_, n, tok in ((wl, "AnalystWatchlistFootnote", "{{ANALYST_WATCHLIST_FOOTNOTE}}"),
                   (gp, "AnalystGroupFootnote", "{{ANALYST_GROUP_FOOTNOTE}}")):
    footnote(s_, n, tok)
    for t in POLK_TILES: s_.drop(t, t + "Value", t + "Label")
    s_.drop("PolkMatchRateNote")

# inventory footnote: the note was already dropped there, so clone from the
# watchlist's finished footnote shape instead
el = copy.deepcopy(wl.get("AnalystWatchlistFootnote")._element)
inv.s.shapes._spTree.append(el)
f = inv.s.shapes[-1]; f.name = "AnalystFootnote"; text(f, "{{ANALYST_FOOTNOTE}}")

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
print([key_of(x).split("\n")[0].replace("key: report:", "") for x in prs.slides])
