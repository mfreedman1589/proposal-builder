import copy, re
from lxml import etree
from pptx import Presentation
from pptx.util import Inches

SRC, OUT = "REPORT_MASTER_v0_8.pptx", "REPORT_MASTER_v0_9.pptx"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
A = "{%s}" % NS["a"]; P = "{%s}" % NS["p"]

# Andrea's theme, read from the source deck (both decks share it)
THEME = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "1F497D", "lt2": "EEECE1",
         "accent1": "4F81BD", "accent2": "C0504D", "accent3": "9BBB59",
         "accent4": "8064A2", "accent5": "4BACC6", "accent6": "F79646",
         "hlink": "0000FF", "folHlink": "800080"}

prs = Presentation(SRC)
layout = prs.slides[0].slide_layout

def resolve_theme(el):
    """Replace every <a:schemeClr val=X/> with the literal <a:srgbClr/>."""
    for sc in list(el.iter(A + "schemeClr")):
        val = sc.get("val")
        if val in THEME:
            rgb = etree.SubElement(sc.getparent(), A + "srgbClr"); rgb.set("val", THEME[val])
            for child in sc: rgb.append(child)
            parent = sc.getparent(); parent.replace(sc, rgb)

def import_slide(src_path, key):
    src = Presentation(src_path).slides[0]
    new = prs.slides.add_slide(layout)
    for sh in list(new.shapes): sh._element.getparent().remove(sh._element)
    # background
    bg = src._element.find(P + "cSld").find(P + "bg")
    if bg is not None:
        cSld = new._element.find(P + "cSld")
        b = copy.deepcopy(bg); resolve_theme(b); cSld.insert(0, b)
    for sh in src.shapes:
        el = copy.deepcopy(sh._element); resolve_theme(el)
        new.shapes._spTree.append(el)
    new.notes_slide.notes_text_frame.text = f"key: {key}\nstandalone: true"
    return new

def shape(s, name):
    for sh in s.shapes:
        if sh.name == name: return sh
    raise KeyError(name)

def set_text(sh, text):
    tf = sh.text_frame; p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]: r._r.getparent().remove(r._r)
    else: p.add_run().text = text
    for extra in tf.paragraphs[1:]: extra._p.getparent().remove(extra._p)

def ren(s, old, new, token=None):
    sh = shape(s, old); sh.name = new
    if token is not None: set_text(sh, token)
    return sh

def drop(s, *names):
    for n in names: sh = shape(s, n); sh._element.getparent().remove(sh._element)

# =====================================================================
# report:summary  (St. James "Campaign Performance Snapshot")
# =====================================================================
s = import_slide("/mnt/user-data/uploads/The_St__James_Attribution_1st_Flight.pptx", "report:summary")
ren(s, "Rectangle 1", "TopBar")
ren(s, "TextBox 2", "Eyebrow")                       # static: CAMPAIGN PERFORMANCE SNAPSHOT
ren(s, "TextBox 3", "SummaryClientName", "{{CLIENT_NAME}}")
ren(s, "TextBox 4", "SummarySubtitle", "{{SUMMARY_SUBTITLE}}")
for i, (rr, v, l) in enumerate([("Rounded Rectangle 5", "TextBox 6", "TextBox 7"),
                                ("Rounded Rectangle 8", "TextBox 9", "TextBox 10"),
                                ("Rounded Rectangle 11", "TextBox 12", "TextBox 13"),
                                ("Rounded Rectangle 14", "TextBox 15", "TextBox 16")], 1):
    ren(s, rr, f"SummaryTile{i}")
    ren(s, v, f"SummaryTile{i}Value", f"{{{{SUMMARY_TILE_{i}_VALUE}}}}")
    ren(s, l, f"SummaryTile{i}Label", f"{{{{SUMMARY_TILE_{i}_LABEL}}}}")
ren(s, "Rectangle 17", "SidebarPanel")
ren(s, "TextBox 18", "SidebarHeader")                # static: WHAT STOOD OUT
ren(s, "TextBox 19", "SidebarHeadline", "{{SIDEBAR_HEADLINE}}")
ren(s, "TextBox 20", "SidebarStatValue", "{{SIDEBAR_STAT_VALUE}}")
ren(s, "TextBox 21", "SidebarStatLabel", "{{SIDEBAR_STAT_LABEL}}")
ren(s, "TextBox 22", "SidebarStatDetail", "{{SIDEBAR_STAT_DETAIL}}")
for i, (v, l) in enumerate([("TextBox 23", "TextBox 24"), ("TextBox 25", "TextBox 26"), ("TextBox 27", "TextBox 28")], 1):
    ren(s, v, f"SidebarSub{i}Value", f"{{{{SIDEBAR_SUB_{i}_VALUE}}}}")
    ren(s, l, f"SidebarSub{i}Label", f"{{{{SIDEBAR_SUB_{i}_LABEL}}}}")
ren(s, "Rectangle 29", "SidebarDivider")
ren(s, "TextBox 30", "SidebarSecondHeader", "{{SIDEBAR_SECOND_HEADER}}")
ren(s, "TextBox 31", "SidebarSecondValue", "{{SIDEBAR_SECOND_VALUE}}")
ren(s, "TextBox 32", "SidebarSecondUnit", "{{SIDEBAR_SECOND_UNIT}}")
ren(s, "TextBox 33", "SidebarSecondDetail", "{{SIDEBAR_SECOND_DETAIL}}")
ren(s, "TextBox 34", "BottomLineHeader")             # static: BOTTOM LINE
ren(s, "TextBox 35", "BottomLine", "{{BOTTOM_LINE}}")
ren(s, "TextBox 36", "TakeawaysHeader")              # static: KEY TAKEAWAYS
for i, (num, head, det) in enumerate([("TextBox 37", "TextBox 38", "TextBox 39"),
                                      ("TextBox 41", "TextBox 42", "TextBox 43"),
                                      ("TextBox 45", "TextBox 46", "TextBox 47")], 1):
    ren(s, num, f"SummaryTakeaway{i}Num")            # static: 01 / 02 / 03
    ren(s, head, f"SummaryTakeaway{i}Head", f"{{{{SUMMARY_TAKEAWAY_{i}_HEAD}}}}")
    ren(s, det, f"SummaryTakeaway{i}Detail", f"{{{{SUMMARY_TAKEAWAY_{i}_DETAIL}}}}")
ren(s, "Rectangle 40", "SummaryTakeawayDivider1")
ren(s, "Rectangle 44", "SummaryTakeawayDivider2")
ren(s, "TextBox 48", "SummaryFootnote", "{{SUMMARY_FOOTNOTE}}")

# =====================================================================
# report:case_study  (Arundel)
# =====================================================================
c = import_slide("/mnt/user-data/uploads/Arundel_Bank_Premion_Streaming_2026.pptx", "report:case_study")
ren(c, "Rectangle 1", "TopBar")
ren(c, "TextBox 2", "CsEyebrow", "{{CS_EYEBROW}}")
ren(c, "TextBox 3", "CsHeadline", "{{CS_HEADLINE}}")
ren(c, "TextBox 4", "CsSubhead", "{{CS_SUBHEAD}}")
for i, (rr, v, l) in enumerate([("Rounded Rectangle 5", "TextBox 6", "TextBox 7"),
                                ("Rounded Rectangle 8", "TextBox 9", "TextBox 10"),
                                ("Rounded Rectangle 11", "TextBox 12", "TextBox 13"),
                                ("Rounded Rectangle 14", "TextBox 15", "TextBox 16"),
                                ("Rounded Rectangle 17", "TextBox 18", "TextBox 19")], 1):
    ren(c, rr, f"CsTile{i}")
    ren(c, v, f"CsTile{i}Value", f"{{{{CS_TILE_{i}_VALUE}}}}")
    ren(c, l, f"CsTile{i}Label", f"{{{{CS_TILE_{i}_LABEL}}}}")
# column 1
ren(c, "TextBox 20", "CsCol1Label", "{{CS_COL_1_LABEL}}")
ren(c, "TextBox 21", "CsCol1Headline", "{{CS_COL_1_HEADLINE}}")
ren(c, "TextBox 22", "CsCol1Body", "{{CS_COL_1_BODY}}")
# column 2 — merge the two body boxes into one
ren(c, "TextBox 24", "CsCol2Label", "{{CS_COL_2_LABEL}}")
ren(c, "TextBox 25", "CsCol2Headline", "{{CS_COL_2_HEADLINE}}")
b2 = ren(c, "TextBox 26", "CsCol2Body", "{{CS_COL_2_BODY}}")
b2b = shape(c, "TextBox 27")
b2.height = (b2b.top + b2b.height) - b2.top
drop(c, "TextBox 27")
# column 3 — body plus the "optimization opportunity" note
ren(c, "TextBox 28", "CsCol3Label", "{{CS_COL_3_LABEL}}")
ren(c, "TextBox 29", "CsCol3Headline", "{{CS_COL_3_HEADLINE}}")
ren(c, "TextBox 30", "CsCol3Body", "{{CS_COL_3_BODY}}")
ren(c, "TextBox 31", "CsCol3Note", "{{CS_COL_3_NOTE}}")
ren(c, "Rounded Rectangle 32", "CsTakeawayBand")
ren(c, "TextBox 33", "CsTakeawayHeader")             # static: CASE STUDY TAKEAWAY
ren(c, "TextBox 34", "CsTakeaway", "{{CS_TAKEAWAY}}")
ren(c, "TextBox 35", "CsSource", "{{CS_SOURCE}}")

prs.save(OUT)
print("saved", OUT, "slides", len(prs.slides))
