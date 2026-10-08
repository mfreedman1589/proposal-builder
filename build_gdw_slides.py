"""Add two restyled Great Day Washington product slides to the proposal master.

Usage:
    python build_gdw_slides.py <proposal_master_in.pptx> <GREAT_DAY_WASHINGTON_MEDIA_KIT_GROSS.pptx> <out.pptx>

Rebuilds GDW media-kit slides 3 (show overview) and 6 (production) in the
proposal master's visual language — the "Measure Sales Conversions" layout:
navy text panel on the left (eyebrow / Aptos Black title / lead / green
section labels / "+" bullets), full-bleed photo on the right, logo bottom-right.

Images are pulled from the media kit by shape name (not re-used slides), so
nothing from the media kit's layouts/masters comes along. Large photos are
downscaled to keep the master small.

Speaker-notes keys: gdw:overview, gdw:production -- the app's product-slide
convention is the bare product key with a colon sub-key (am:audience,
sport:nfl_reg), never a "product:" prefix. assembly.resolve_active_keys
switches both on whenever a Great Day Washington line is on the plan.

Run against the CURRENT master (version 12 at the time of writing), then upload
the output on dev only -- the nightly merge activates it. Refuses to run on a
deck that already carries the GDW slides, so it can't stack a second pair.
"""
import io
import sys

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import slide_map

MASTER_IN, GDW_IN, OUT = sys.argv[1:4]

NAVY = RGBColor(0x13, 0x19, 0x2C)
NAVY_DEEP = RGBColor(0x0B, 0x10, 0x24)
GREEN = RGBColor(0x66, 0xDD, 0x19)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0xB8, 0xBE, 0xD0)

PANEL_W = 6.85          # navy panel width (matches master product slides)
LEFT = 0.85             # text left edge
TEXT_W = PANEL_W - LEFT - 0.45


# ---------------------------------------------------------------- helpers
def pic_blob(prs, slide_idx, shape_name, max_w=None, jpeg=False):
    for sh in prs.slides[slide_idx].shapes:
        if sh.name == shape_name:
            blob = sh.image.blob
            if max_w is None:
                return io.BytesIO(blob), sh
            im = Image.open(io.BytesIO(blob))
            if im.width > max_w:
                im = im.resize((max_w, round(im.height * max_w / im.width)), Image.LANCZOS)
            out = io.BytesIO()
            if jpeg:
                im.convert("RGB").save(out, "JPEG", quality=85, optimize=True)
            else:
                im.save(out, "PNG", optimize=True)
            out.seek(0)
            return out, sh
    raise KeyError(f"slide {slide_idx + 1}: {shape_name}")


def blank_layout(prs):
    for name in ("Blank", "1_Blank", "3_Blank"):
        for lay in prs.slide_layouts:
            if lay.name == name:
                return lay
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


def new_slide(prs, key):
    s = prs.slides.add_slide(blank_layout(prs))
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    s.notes_slide.notes_text_frame.text = f"key: {key}"
    return s


def rect(s, name, x, y, w, h, color, transparency=None):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    r.name = name
    r.fill.solid()
    r.fill.fore_color.rgb = color
    r.line.fill.background()
    r.shadow.inherit = False
    if transparency is not None:  # 0-100
        srgb = r.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        a = srgb.makeelement(qn("a:alpha"), {"val": str(int((100 - transparency) * 1000))})
        srgb.append(a)
    return r


def gradient_panel(s, name, x, y, w, h):
    """Navy panel with the subtle top-left glow the master's panels have."""
    r = rect(s, name, x, y, w, h, NAVY)
    spPr = r._element.spPr
    spPr.remove(spPr.find(qn("a:solidFill")))
    grad = spPr.makeelement(qn("a:gradFill"), {"rotWithShape": "1"})
    gs = grad.makeelement(qn("a:gsLst"), {})
    for pos, hexc in ((0, "1C2556"), (55000, "13192C"), (100000, "0B1024")):
        g = gs.makeelement(qn("a:gs"), {"pos": str(pos)})
        c = g.makeelement(qn("a:srgbClr"), {"val": hexc})
        g.append(c)
        gs.append(g)
    grad.append(gs)
    lin = grad.makeelement(qn("a:lin"), {"ang": "3600000", "scaled": "0"})
    grad.append(lin)
    # gradFill must precede a:ln
    spPr.insert(list(spPr).index(spPr.find(qn("a:prstGeom"))) + 1, grad)
    return r


def text(s, name, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    """paras: list of dicts or list of lists of run-dicts.
    run keys: t, font, size, bold, color, spc (1/100 pt)
    para keys (when dict with 'runs'): runs, space_after, space_before, bullet, line"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    first = True
    for para in paras:
        if isinstance(para, list):
            para = {"runs": para}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = para.get("align", PP_ALIGN.LEFT)
        if "space_after" in para:
            p.space_after = Pt(para["space_after"])
        if "space_before" in para:
            p.space_before = Pt(para["space_before"])
        if "line" in para:
            p.line_spacing = para["line"]
        if para.get("bullet"):
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(Inches(0.2)))
            pPr.set("indent", str(-Inches(0.2)))
            clr = pPr.makeelement(qn("a:buClr"), {})
            clr.append(clr.makeelement(qn("a:srgbClr"), {"val": "66DD19"}))
            pPr.append(clr)
            pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Aptos"}))
            pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "+"}))
        for rd in para["runs"]:
            r = p.add_run()
            r.text = rd["t"]
            f = r.font
            f.name = rd.get("font", "Aptos")
            f.size = Pt(rd.get("size", 14))
            f.bold = rd.get("bold", False)
            f.color.rgb = rd.get("color", WHITE)
            if rd.get("spc"):
                r._r.get_or_add_rPr().set("spc", str(rd["spc"]))
    return tb


def eyebrow(s, y, right_text):
    return text(s, "Eyebrow", LEFT, y, TEXT_W, 0.25, [[
        {"t": "WUSA9 ", "font": "Aptos SemiBold", "size": 9, "bold": True, "spc": 300},
        {"t": "|", "font": "Aptos SemiBold", "size": 9, "bold": True, "spc": 300, "color": GREEN},
        {"t": f" {right_text}", "font": "Aptos SemiBold", "size": 9, "bold": True, "spc": 300},
    ]])


def label(t):
    return {"runs": [{"t": t, "font": "Aptos SemiBold", "size": 10, "bold": True,
                      "color": GREEN, "spc": 200}], "space_after": 4}


def bullet(t, size=13):
    return {"runs": [{"t": t, "font": "Aptos Light", "size": size}], "bullet": True, "space_after": 2}


def photo_right(s, stream, src_w, src_h, name, focus_x=0.5):
    """Full-bleed photo filling the right panel, cropped to fit.
    focus_x: 0 = keep left edge, 0.5 = center, 1 = keep right edge."""
    x, w, h = PANEL_W, 13.333 - PANEL_W, 7.5
    pic = s.shapes.add_picture(stream, Inches(x), 0, Inches(w), Inches(h))
    pic.name = name
    box, img = w / h, src_w / src_h
    if img > box:   # too wide -> crop sides
        c = 1 - box / img
        pic.crop_left, pic.crop_right = c * focus_x, c * (1 - focus_x)
    else:           # too tall -> crop top/bottom
        c = (1 - img / box) / 2
        pic.crop_top = pic.crop_bottom = c
    return pic


def wusa9_logo(s, stream):
    # media-kit PNG is 1920x1080 with padding; crop to the mark
    pic = s.shapes.add_picture(stream, Inches(11.55), Inches(6.88), Inches(1.15))
    pic.name = "Wusa9Logo"
    return pic


GDW_KEYS = ("gdw:overview", "gdw:production")


def move_before_media_plan(prs, n_new):
    """Move the last `n_new` slides to just before the standard media plan
    slide. Matched on the exact `proposal_template` key -- a substring match
    also hits `proposal_template_total_tv:dc/harrisburg`, and taking the last
    hit put the new slides between the two Total TV plan variants."""
    lst = prs.slides._sldIdLst
    ids = list(lst)
    target = next((i for i, sl in enumerate(prs.slides)
                   if sl.has_notes_slide and slide_map.notes_key(sl) == "proposal_template"), None)
    if target is None:
        raise SystemExit("no slide keyed proposal_template -- is this the proposal master?")
    new = ids[-n_new:]
    for el in new:
        lst.remove(el)
    for k, el in enumerate(new):
        lst.insert(target + k, el)


# ---------------------------------------------------------------- build
prs = Presentation(MASTER_IN)
gdw = Presentation(GDW_IN)
already = [k for k in (slide_map.notes_key(sl) for sl in prs.slides if sl.has_notes_slide)
           if k in GDW_KEYS]
if already:
    raise SystemExit(f"{MASTER_IN} already has the Great Day Washington slides ({', '.join(already)})")

houses, _ = pic_blob(gdw, 2, "Picture 6", max_w=2000, jpeg=True)        # row houses
host, host_sh = pic_blob(gdw, 2, "Picture 2")                           # Elaine headshot
gdw_logo, _ = pic_blob(gdw, 2, "Picture 4")                             # GDW wordmark (white)
camera, _ = pic_blob(gdw, 5, "Picture 4", max_w=2200, jpeg=True)        # camera operator
wusa9_raw, _ = pic_blob(gdw, 2, "Picture 61")                           # WUSA9 mark (359x95)

def fresh(b):
    b.seek(0)
    return io.BytesIO(b.read())

# ============ Slide A — overview (media kit slide 3) ============
a = new_slide(prs, "gdw:overview")
rect(a, "Background", 0, 0, 13.333, 7.5, NAVY)
ph = photo_right(a, fresh(houses), 2000, round(2000 * 1648 / 2930), "GdwPhoto")
rect(a, "PhotoScrim", PANEL_W, 0, 13.333 - PANEL_W, 7.5, NAVY_DEEP, transparency=45)
gradient_panel(a, "NavyPanel", 0, 0, PANEL_W, 7.5)

# right panel: show logo + host card
lg = a.shapes.add_picture(fresh(gdw_logo), Inches(PANEL_W + 0.9), Inches(0.75), Inches(4.7))
lg.name = "GdwLogo"
lg.crop_top, lg.crop_bottom = 0.28043, 0.27105                     # same crop the media kit uses
lg.height = Emu(int(Inches(4.7) * 1080 * (1 - 0.28043 - 0.27105) / 1920))
hx, hy, hw = PANEL_W + 1.67, 2.35, 3.15
frame = rect(a, "HostFrame", hx - 0.06, hy - 0.06, hw + 0.12, hw * (564 * (1 - .07564)) / (641 * (1 - .18667)) + 0.12, WHITE)
hp = a.shapes.add_picture(fresh(host), Inches(hx), Inches(hy), Inches(hw))
hp.name = "HostPhoto"
hp.crop_left, hp.crop_right, hp.crop_top, hp.crop_bottom = host_sh.crop_left, host_sh.crop_right, host_sh.crop_top, host_sh.crop_bottom
hp.height = Emu(int(Inches(hw) * (564 * (1 - .07564)) / (641 * (1 - .18667))))
cap_y = hy + hp.height / 914400 + 0.22
text(a, "HostCaption", hx - 0.3, cap_y, hw + 0.6, 0.6, [
    {"runs": [{"t": "MEET YOUR HOST", "font": "Aptos SemiBold", "size": 9, "bold": True,
               "color": GREEN, "spc": 300}], "align": PP_ALIGN.CENTER, "space_after": 2},
    {"runs": [{"t": "Elaine Espinola", "font": "Aptos", "size": 18, "bold": True}],
     "align": PP_ALIGN.CENTER},
])

# left panel copy
eyebrow(a, 1.10, "GREAT DAY WASHINGTON")
text(a, "Title", LEFT, 1.37, TEXT_W + 0.4, 1.75, [
    {"runs": [{"t": "Your Business,", "font": "Aptos Black", "size": 46, "bold": True}], "line": 0.88},
    {"runs": [{"t": "Center Stage", "font": "Aptos Black", "size": 46, "bold": True, "color": GREEN}], "line": 0.88},
])
text(a, "Lead", LEFT, 3.00, TEXT_W, 0.95, [[
    {"t": "Feature your business in a 3–4 minute interview on WUSA9's daily lifestyle show "
          "— then live on at WUSA9.com.", "size": 17},
]])
text(a, "Body", LEFT, 4.03, TEXT_W, 1.3, [
    {"runs": [{"t": "Great Day Washington is DC's daily connection to entertainment, lifestyle, "
                    "retail, home and health — Monday through Friday, year-round. Viewers tune in "
                    "for local stories and the businesses behind them.",
               "font": "Aptos Light", "size": 13}], "line": 1.05},
])
text(a, "AirsLabel", LEFT, 5.33, TEXT_W, 0.28, [label("AIRS EVERY WEEKDAY")])
for i, (big, small) in enumerate((("9 AM", "WUSA9 Mornings"), ("3 PM", "Great Day Washington"))):
    x = LEFT + i * 2.6
    text(a, f"AirTime{i + 1}", x, 5.63, 2.4, 0.62, [[
        {"t": big, "font": "Aptos Black", "size": 34, "bold": True, "color": GREEN}]])
    text(a, f"AirTime{i + 1}Label", x, 6.23, 2.4, 0.3, [[
        {"t": small, "font": "Aptos Light", "size": 12}]])
rect(a, "AirDivider", LEFT + 2.3, 5.72, 0.02, 0.78, MUTED)
wusa9_logo(a, fresh(wusa9_raw))

# ============ Slide B — production (media kit slide 6) ============
b = new_slide(prs, "gdw:production")
rect(b, "Background", 0, 0, 13.333, 7.5, NAVY)
photo_right(b, fresh(camera), 2200, round(2200 * 2854 / 4490), "GdwPhoto", focus_x=0.9)
gradient_panel(b, "NavyPanel", 0, 0, PANEL_W, 7.5)

eyebrow(b, 0.85, "GREAT DAY WASHINGTON")
text(b, "Title", LEFT, 1.12, TEXT_W + 0.4, 1.5, [
    {"runs": [{"t": "Great Day", "font": "Aptos Black", "size": 46, "bold": True}], "line": 0.88},
    {"runs": [{"t": "Production", "font": "Aptos Black", "size": 46, "bold": True, "color": GREEN}],
     "line": 0.88},
])
text(b, "Subhead", LEFT, 2.62, TEXT_W, 0.25, [[
    {"t": "YOUR SEGMENT, BUILT WITH OUR PRODUCERS", "font": "Aptos SemiBold", "size": 9,
     "bold": True, "spc": 300}]])
text(b, "Lead", LEFT, 2.98, TEXT_W, 0.7, [[
    {"t": "Work with one of our show's producers to create a customized 3–4 minute segment.",
     "size": 17}]])

col_w = (TEXT_W - 0.3) / 2
text(b, "ShootOptions", LEFT, 3.85, col_w, 1.6, [
    label("SHOOT OPTIONS"),
    bullet("Live in-studio"),
    bullet("Live or taped on-location"),
    bullet("Via Zoom"),
    bullet("Pre-shot in studio"),
])
text(b, "AfterAiring", LEFT + col_w + 0.3, 3.85, col_w, 1.6, [
    label("AFTER IT AIRS"),
    bullet("Posted at WUSA9.com/GreatDay"),
    bullet("Shared on GDW's Facebook"),
])
text(b, "Extend", LEFT, 5.3, TEXT_W, 1.3, [
    label("EXTEND YOUR SEGMENT"),
    bullet("WUSA9+ streaming impressions"),
    bullet("WUSA9.com home page takeover"),
    bullet("WUSA9.com display"),
])
text(b, "ProductionNotes", LEFT, 6.55, TEXT_W, 0.75, [
    {"runs": [{"t": "Minimum 2-week turnaround; a production meeting is encouraged. Appearances require "
                    "a signed FCC waiver and license agreement for release of the MP4 before airing. "
                    "Encore airings are not re-posted. Segment availability pending site survey.",
               "font": "Aptos Light", "size": 8, "color": MUTED}], "line": 1.0},
])
wusa9_logo(b, fresh(wusa9_raw))

move_before_media_plan(prs, 2)
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
