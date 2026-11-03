"""REPORT_MASTER_v0_17 -- Live Sports website attribution (2026-11).

Run against the CURRENT active template (v0_16):
    python build_v0_17.py REPORT_MASTER_v0_16.pptx REPORT_MASTER_v0_17.pptx
then upload the output on dev only, inactive (the nightly merge activates it).

Every new slide is cloned from an existing one, so fonts, colours, tables and
tiles are the template's own. All are optional: report_assembly drops each
one a report has nothing for, and a report with no Live Sports attribution
file builds exactly as on v0_16.

  * report:live_sports -- title becomes "{{SPORTS_LABEL}} Delivery" (the
    label defaults from the package type, "Live NFL", and the rep can edit
    it).
  * report:live_sports_attribution (Simple and Extended) -- after the
    delivery slide: four tiles (attributed impressions, attributed rate,
    unique visitors, unique visitor rate), a conversions line deleted unless
    either is above zero, the top 5 pages and a "what this signals" panel.
  * report:live_sports_response (Extended) -- the weekly attributed-rate
    trend, days from ad to visit, how visitors arrived, and device. No day
    of week: for sports it is the game schedule.
  * report:live_sports_creative_geo (Extended) -- creative comparison and the
    target market, other markets as a one-line footnote.
  * report:highlights_combined -- Top Highlights with the tile row replaced
    by a side-by-side table (Targeted Streaming | sports | Combined) and the
    "visitors may overlap" footnote. Used instead of report:highlights when
    both attribution files are present.
  * report:campaign_comparison -- "How They Work Together", after the ZIP
    slide: a side-by-side measures table and a panel of code-written lines.
"""
import copy
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

SRC = sys.argv[1] if len(sys.argv) > 1 else "REPORT_MASTER_v0_16.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "REPORT_MASTER_v0_17.pptx"
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

# ======================================================================
# 0. report:live_sports -- the label in the title
# ======================================================================
live = S(find("report:live_sports"))
text(live.get("Text 1"), "{{SPORTS_LABEL}} Delivery")

# ======================================================================
# 1. report:live_sports_attribution
# ======================================================================
att = S(clone_after(live.s, "report:live_sports",
                    "key: report:live_sports_attribution\nlive_sports_attribution: simple"))
text(att.get("Text 1"), "{{SPORTS_LABEL}} Website Response")
text(att.get("Text 2"), "{{SPORTS_ATTR_SUBTITLE}}")
att.drop("SportsByLeagueHeader", "SportsByLeagueTable", "SportsRfpidCaption", "SportsEventTable")
# Four tiles in the row the delivery slide's three use.
tiles = [("SportsImpressionsTile", "SportsAttrImpressionsTile", "{{SPORTS_ATTR_IMPRESSIONS}}",
          "Attributed impressions"),
         ("SportsVcrTile", "SportsAttrRateTile", "{{SPORTS_ATTR_RATE}}", "Attributed rate"),
         ("SportsPacingTile", "SportsVisitorsTile", "{{SPORTS_VISITORS}}",
          "Unique website visitors")]
for old, new, tok, lab in tiles:
    for suf in ("", "Value", "Label"): att.rename(old + suf, new + suf)
    text(att.get(new + "Value"), tok); text(att.get(new + "Label"), lab)
for suf in ("", "Value", "Label"):
    att.clone("SportsVisitorsTile" + suf, "SportsVisitorRateTile" + suf)
text(att.get("SportsVisitorRateTileValue"), "{{SPORTS_VISITOR_RATE}}")
text(att.get("SportsVisitorRateTileLabel"), "Unique visitor rate")
width, gap = 2.895, 0.25
for i, name in enumerate(("SportsAttrImpressionsTile", "SportsAttrRateTile",
                          "SportsVisitorsTile", "SportsVisitorRateTile")):
    x = 0.50 + i * (width + gap)
    place(att.get(name), x, 2.10, width, 0.85)
    place(att.get(name + "Value"), x + 0.10, 2.18, width - 0.20, 0.48)
    place(att.get(name + "Label"), x + 0.10, 2.66, width - 0.20, 0.25)
note = att.borrow(url, "UrlReachFootnote", "SportsConversionsNote")
text(note, "{{SPORTS_CONVERSIONS_NOTE}}")
place(note, 0.50, 3.02, 12.33, 0.25)
att.rename("Text 18", "SportsTopPagesHeader")
text(att.get("SportsTopPagesHeader"), "TOP PAGES BY ATTRIBUTED VISITORS")
place(att.get("SportsTopPagesHeader"), 0.50, 3.42, 7.70, 0.26)
pages = att.borrow(url, "TopUrlTable", "SportsTopPagesTable")
recolumn(pages, ["Page / section", "Visitors"], [5.70, 2.00], "{{SPORTS_TOP_PAGE_ROWS}}")
place(pages, 0.50, 3.74)
foot = att.borrow(url, "UrlReachFootnote", "SportsPagesFootnote")
text(foot, "Each page counts its own visitors; one visitor can reach several pages.")
place(foot, 0.50, 6.70, 7.70, 0.25)
att.borrow(url, "Shape 5", "SportsAttrPanel")
place(att.get("SportsAttrPanel"), 8.45, 3.42, 4.38, 3.48)
header = att.borrow(url, "Text 6", "SportsAttrPanelHeader")
place(header, 8.70, 3.60, 3.90, 0.26)
att.rename("LiveSportsNarrative", "SportsAttrNarrative")
text(att.get("SportsAttrNarrative"), "{{SPORTS_ATTR_NARRATIVE}}")
place(att.get("SportsAttrNarrative"), 8.70, 3.98, 3.90, 2.75)
att.to_front("SportsAttrNarrative")

# ======================================================================
# 2. report:live_sports_response (Extended)
# ======================================================================
resp_src = find("report:response_profile")
resp = S(clone_after(resp_src, "report:live_sports_attribution",
                     "key: report:live_sports_response\nlive_sports_attribution: extended"))
text(resp.get("Text 1"), "{{SPORTS_LABEL}}: Response Timing & Source")
text(resp.get("Text 2"), "{{SPORTS_RESPONSE_SUBTITLE}}")
resp.drop("DayOfWeekHeader", "DayOfWeekTable", "Shape 7", "ResponseProfileNarrative")
resp.rename("Text 3", "SportsTrendHeader")
text(resp.get("SportsTrendHeader"), "ATTRIBUTED RATE BY WEEK")
resp.rename("ChartRegion", "SportsTrendRegion")
resp.rename("ChartRegionLabel", "SportsTrendRegionLabel")
text(resp.get("SportsTrendRegionLabel"), "SportsTrendRegion · attributed rate by week")
place(resp.get("SportsTrendRegion"), 0.50, 2.45, 5.97, 2.40)
place(resp.get("SportsTrendRegionLabel"), 0.50, 2.45, 5.97, 2.40)
dev_head = resp.clone("SportsTrendHeader", "SportsDeviceHeader")
text(dev_head, "VISITS BY DEVICE")
place(dev_head, 0.50, 5.02, 5.97, 0.26)
device = resp.clone("ReferralTable", "SportsDeviceTable")
recolumn(device, ["Device", "Share of visits"], [3.97, 2.00], "{{SPORTS_DEVICE_ROWS}}")
place(device, 0.50, 5.34)
resp.rename("Text 4", "SportsRecencyHeader")
text(resp.get("SportsRecencyHeader"), "DAYS FROM AD TO VISIT")
recency = resp.clone("ReferralTable", "SportsRecencyTable")
recolumn(recency, ["Days since ad", "Visitors", "Share"], [2.97, 1.60, 1.40],
         "{{SPORTS_RECENCY_ROWS}}")
place(recency, 6.87, 2.45)
ref_head = resp.clone("SportsRecencyHeader", "SportsReferralHeader")
text(ref_head, "HOW VISITORS ARRIVED")
place(ref_head, 6.87, 4.40, 5.97, 0.26)
resp.rename("ReferralTable", "SportsReferralTable")
recolumn(resp.get("SportsReferralTable"), ["Source", "Visitors", "Share"], [2.97, 1.60, 1.40],
         "{{SPORTS_REFERRAL_ROWS}}")
place(resp.get("SportsReferralTable"), 6.87, 4.72)

# ======================================================================
# 3. report:live_sports_creative_geo (Extended)
# ======================================================================
brk_src = find("report:attribution_breakdown")
cg = S(clone_after(brk_src, "report:live_sports_response",
                   "key: report:live_sports_creative_geo\nlive_sports_attribution: extended"))
text(cg.get("Text 1"), "{{SPORTS_LABEL}}: Creative & Geography")
text(cg.get("Text 2"), "{{SPORTS_CREATIVE_GEO_SUBTITLE}}")
cg.drop("ChartRegion", "ChartRegionLabel")
cg.rename("Text 3", "SportsCreativeHeader")
text(cg.get("SportsCreativeHeader"), "ATTRIBUTED RATE BY CREATIVE")
cg.rename("BreakdownTable", "SportsCreativeTable")
geo = cg.clone("SportsCreativeTable", "SportsGeoTable")
recolumn(cg.get("SportsCreativeTable"), ["Creative", "Impressions", "Attributed rate"],
         [3.10, 1.50, 1.50], "{{SPORTS_CREATIVE_ROWS}}")
place(cg.get("SportsCreativeTable"), 0.50, 2.45)
cg.rename("Text 4", "SportsGeoHeader")
text(cg.get("SportsGeoHeader"), "TARGET MARKET")
recolumn(geo, ["Market", "Impressions", "Attributed rate"], [2.97, 1.50, 1.50],
         "{{SPORTS_GEO_ROWS}}")
place(geo, 6.87, 2.45)
outside = cg.borrow(url, "UrlReachFootnote", "SportsOutsideMarketNote")
text(outside, "{{SPORTS_OUTSIDE_MARKET_NOTE}}")
place(outside, 6.87, 3.30, 5.97, 0.40)
cg.rename("AttributionNarrative", "SportsCreativeGeoNarrative")
text(cg.get("SportsCreativeGeoNarrative"), "{{SPORTS_CREATIVE_GEO_NARRATIVE}}")

# ======================================================================
# 4. report:highlights_combined
# ======================================================================
hl_src = find("report:highlights")
hc = S(clone_after(hl_src, "report:highlights", "key: report:highlights_combined"))
for t in ("ImpressionsTile", "VisitorsTile", "RateTile", "ConversionsTile"):
    hc.drop(t, t + "Value", t + "Label")
hc.drop("RateTileDelta", "VisitorsTileDelta", "CostPerVisitNote")
summary = hc.borrow(url, "TopUrlTable", "CampaignSummaryTable")
recolumn(summary, ["", "Targeted Streaming", "{{SPORTS_LABEL}}", "Combined"],
         [3.60, 2.91, 2.91, 2.91], "{{CAMPAIGN_SUMMARY_ROWS}}")
place(summary, 0.50, 2.05)
foot = hc.borrow(url, "UrlReachFootnote", "CampaignSummaryFootnote")
text(foot, "{{CAMPAIGN_SUMMARY_FOOTNOTE}}")
place(foot, 0.50, 3.92, 12.33, 0.25)
place(hc.get("Shape 11"), 0.50, 4.30, 7.70, 2.60)
place(hc.get("Text 12"), 0.80, 4.45, 7.10, 0.26)
place(hc.get("HIGHLIGHTBullets"), 0.80, 4.80, 7.10, 1.95)
place(hc.get("TrendChartHeader"), 8.45, 4.30, 4.38, 0.26)
text(hc.get("TrendChartHeader"), "TARGETED STREAMING, BY WEEK")
place(hc.get("TrendChartRegion"), 8.45, 4.62, 4.38, 2.28)
place(hc.get("TrendChartRegionLabel"), 8.45, 4.62, 4.38, 2.28)

# ======================================================================
# 5. report:campaign_comparison
# ======================================================================
cmp_ = S(clone_after(url.s, "report:zip_analysis", "key: report:campaign_comparison"))
text(cmp_.get("Text 1"), "How They Work Together")
text(cmp_.get("Text 2"), "{{COMPARISON_PERIOD_NOTE}}")
cmp_.drop("Text 4", "TopUrlTable", "UrlReachFootnote")
cmp_.rename("Text 3", "ComparisonHeader")
text(cmp_.get("ComparisonHeader"), "SIDE BY SIDE")
cmp_.rename("IntentSummaryTable", "ComparisonTable")
recolumn(cmp_.get("ComparisonTable"), ["", "Targeted Streaming", "{{SPORTS_LABEL}}"],
         [3.20, 2.00, 2.00], "{{COMPARISON_ROWS}}")
cmp_.rename("Shape 5", "ComparisonPanel")
cmp_.rename("Text 6", "ComparisonPanelHeader")
text(cmp_.get("ComparisonPanelHeader"), "HOW THEY WORK TOGETHER")
cmp_.rename("UrlIntentNarrative", "ComparisonNarrative")
text(cmp_.get("ComparisonNarrative"), "{{COMPARISON_NARRATIVE}}")
for r in cmp_.get("ComparisonNarrative").text_frame.paragraphs[0].runs: r.font.size = Pt(13)

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
print([key_of(x).split("\n")[0].replace("key: report:", "") for x in prs.slides])
