"""slide_inheritance.py -- make a copied slide stop depending on its deck.

A slide copied out of another .pptx renders against whatever layout, master
and theme it lands next to, and everything it *inherits* rather than states
is quietly reinterpreted: a placeholder with no explicit size takes the
destination master's default, and every `<a:schemeClr val="tx1"/>` resolves
against the destination's theme. On the furniture case studies that turned
brand-white headings near-black and filled shapes with the wrong colours,
while case studies built entirely from explicit TextBoxes came through fine
-- which is why it survived the chart-import work.

The faithful fix is to import the source layout, master and theme and point
the copied slide at them. That was built and abandoned: the resulting
package is internally consistent by every check available -- every
relationship resolves, no dangling rIds, correct content types, the master
XML byte-identical to its source apart from rIds, layout ids renumbered to
avoid collisions -- and PowerPoint refuses to open it anyway, with no more
detail than "could not open". Importing a deck's *own* master back into
itself reproduces it, so it isn't about the source. It is left unexplained
rather than papered over.

So this flattens instead: resolve what the slide would have inherited, write
it onto the shapes explicitly, and remove the `<p:ph>` element so nothing
inherits from the destination at all. Two passes, in this order:

  1. Placeholders, which pull size/position and text defaults down from the
     source layout and master.
  2. Scheme colours, resolved through the source master's `<p:clrMap>` into
     the source theme's palette and written as explicit RGB.

The order matters: step 1 copies defaults that themselves contain scheme
colour references, and step 2 has to see them.
"""

import re

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

# Run properties worth carrying down. Size and colour are the two the bug
# actually showed; the rest come along because a heading that inherits its
# weight and typeface and loses them is the same defect wearing a hat.
_RUN_ATTRS = ("sz", "b", "i", "u", "strike", "cap", "spc", "baseline")
_RUN_CHILDREN = ("solidFill", "latin", "ea", "cs", "highlight", "effectLst")

# Which of the master's txStyles a placeholder type reads from.
_STYLE_FOR_TYPE = {
    "title": "titleStyle", "ctrTitle": "titleStyle",
    "body": "bodyStyle", "subTitle": "bodyStyle", "obj": "bodyStyle",
}


def _q(tag):
    prefix, _, local = tag.partition(":")
    return f"{{{A if prefix == 'a' else P}}}{local}"


def flatten_inherited_formatting(new_slide, src_slide, relate=None):
    """Write inherited formatting onto a freshly copied slide, in place.

    `relate(src_rel) -> rId` imports a relationship of the source layout or
    master onto the new slide's part, and is required to bring across
    anything those parts draw with a picture. Without it their shapes are
    still copied, minus any image they reference.
    """
    try:
        layout = src_slide.slide_layout
        master = layout.slide_master
    except Exception:                                            # noqa: BLE001
        return
    spTree = new_slide.shapes._spTree
    _flatten_placeholders(spTree, layout, master)
    _inherit_backdrop(new_slide, src_slide, layout, master, relate)
    _resolve_scheme_colors(new_slide._element, master)
    _resolve_theme_fonts(new_slide._element, master)


# ---------------------------------------------------------------------------
# 0. the layout's and master's own content
# ---------------------------------------------------------------------------

def _first_shape_index(spTree):
    """Where shapes start: <p:nvGrpSpPr> and <p:grpSpPr> come first."""
    for index, child in enumerate(spTree):
        if child.tag not in (_q("p:nvGrpSpPr"), _q("p:grpSpPr")):
            return index
    return len(spTree)


def _inherit_backdrop(new_slide, src_slide, layout, master, relate):
    """Draw the source layout's and master's own content onto the slide.

    A slide shows three things stacked: its master's shapes, its layout's,
    and its own. Copying only the third leaves the case study without the
    header band it is titled in, without its background, and -- because the
    slide is now sitting on the destination's Blank layout -- with Premion's
    own master furniture showing through underneath it. The copied slide came
    out dark-on-dark with a "© 2020 PREMION - CONFIDENTIAL" footer that
    belongs to a different deck.

    So the backdrop is copied too, master first and layout over it, inserted
    beneath the slide's own shapes. Placeholders are skipped: an unfilled
    layout placeholder is a prompt ("Click to add title"), not content, and
    PowerPoint doesn't draw them on a slide either.
    """
    spTree = new_slide.shapes._spTree
    at = _first_shape_index(spTree)

    show_master = (layout.element.get("showMasterSp") or "1") != "0"
    for source in ([master] if show_master else []) + [layout]:
        rid_map = _relate_all(source.part, relate)
        for shape in source.element.find(_q("p:cSld")).find(_q("p:spTree")):
            if shape.tag in (_q("p:nvGrpSpPr"), _q("p:grpSpPr")):
                continue
            if shape.find(f".//{_q('p:ph')}") is not None:
                continue
            if _is_furniture(shape):
                continue
            element = _copy(shape)
            _remap(element, rid_map)
            spTree.insert(at, element)
            at += 1

    # The slide carries its whole backdrop now, so the destination master
    # must not add its own on top of it.
    new_slide._element.set("showMasterSp", "0")

    if new_slide._element.find(f"{_q('p:cSld')}/{_q('p:bg')}") is not None:
        return
    for source in (layout, master):
        bg = source.element.find(f"{_q('p:cSld')}/{_q('p:bg')}")
        if bg is None:
            continue
        element = _copy(bg)
        _remap(element, _relate_all(source.part, relate))
        new_slide._element.find(_q("p:cSld")).insert(0, element)
        return


# Slide furniture: true of the source deck, meaningless in a proposal. A
# date or slide number is about the deck it came from, and a confidentiality
# line is a claim about a different document. Most of these arrive as
# placeholders and are skipped already; this catches the decks that build
# them as plain text boxes instead.
_FURNITURE_FIELDS = ("slidenum", "datetime")
_FURNITURE_TEXT = re.compile(r"confidential|copyright|©\s*\d{4}|all rights reserved",
                             re.IGNORECASE)


def _is_furniture(shape):
    for field in shape.iter(_q("a:fld")):
        kind = (field.get("type") or "").lower()
        if any(kind.startswith(prefix) for prefix in _FURNITURE_FIELDS):
            return True
    text = "".join(node.text or "" for node in shape.iter(_q("a:t")))
    return bool(text.strip()) and bool(_FURNITURE_TEXT.search(text))


def _relate_all(part, relate):
    if relate is None:
        return {}
    mapped = {}
    for rId, rel in part.rels.items():
        new_rId = relate(rel)
        if new_rId is not None:
            mapped[rId] = new_rId
    return mapped


def _remap(element, rid_map):
    if not rid_map:
        return
    for attr in ("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                 "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
                 "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link"):
        for node in element.iter():
            value = node.get(attr)
            if value in rid_map:
                node.set(attr, rid_map[value])


# ---------------------------------------------------------------------------
# 1. placeholders
# ---------------------------------------------------------------------------

def _placeholder_index(shapes):
    """{(idx, type): element} for the placeholders of a layout or master."""
    found = {}
    for shape in shapes:
        ph = shape._element.find(f".//{_q('p:ph')}")
        if ph is None:
            continue
        found[(ph.get("idx"), ph.get("type"))] = shape._element
    return found


def _match(index, idx, ph_type):
    """The layout placeholder a slide placeholder inherits from.

    Matched on idx first: that is what PowerPoint uses, and two body
    placeholders on one layout are only told apart by it. Type is the
    fallback for the ones that carry no idx (title, usually).
    """
    for key in ((idx, ph_type), (idx, None), (None, ph_type)):
        if key in index:
            return index[key]
    for (other_idx, other_type), element in index.items():
        if idx is not None and other_idx == idx:
            return element
        if ph_type is not None and other_type == ph_type:
            return element
    return None


def _flatten_placeholders(spTree, layout, master):
    layout_index = _placeholder_index(layout.placeholders)
    master_index = _placeholder_index(master.placeholders)
    txStyles = master.part._element.find(_q("p:txStyles"))

    for shape in list(spTree):
        ph = shape.find(f".//{_q('p:ph')}")
        if ph is None:
            continue
        idx, ph_type = ph.get("idx"), ph.get("type")
        inherited = [element for element in
                     (_match(layout_index, idx, ph_type),
                      _match(master_index, idx, ph_type)) if element is not None]

        _inherit_geometry(shape, inherited)
        _inherit_body_properties(shape, inherited)
        _inherit_text(shape, inherited, txStyles, ph_type)

        # Last: with no <p:ph> the shape is an ordinary shape and inherits
        # nothing from wherever it lands. Everything above has to have run
        # first, because it is all read through this element.
        ph.getparent().remove(ph)
        _ensure_geometry(shape)


def _ensure_geometry(shape):
    """Give a de-placeholdered <p:sp> an explicit shape geometry.

    A placeholder is allowed to have no geometry -- it takes its layout's.
    Strip the <p:ph> and it becomes a shape that is neither a placeholder nor
    an autoshape, which python-pptx refuses to type at all
    ("Shape instance of unrecognized shape type"), and that breaks every walk
    over the deck's shapes rather than anything about the rendering. A plain
    rectangle is what these text placeholders already were.
    """
    # Pictures need this as much as shapes do, and more visibly: a picture
    # placeholder states only its position and takes its geometry from the
    # layout, so stripping <p:ph> left the case study's cover photo in the
    # file, at the right size, drawing nothing at all. A graphic frame (table,
    # chart) carries no spPr and is unaffected.
    if shape.tag not in (_q("p:sp"), _q("p:pic")):
        return
    spPr = shape.find(_q("p:spPr"))
    if spPr is None:
        return
    if spPr.find(_q("a:prstGeom")) is not None or spPr.find(_q("a:custGeom")) is not None:
        return
    geom = etree.SubElement(spPr, _q("a:prstGeom"))
    geom.set("prst", "rect")
    etree.SubElement(geom, _q("a:avLst"))
    # Geometry follows xfrm and precedes fill in the CT_ShapeProperties
    # sequence; appended, it would sit after the fill and be rejected.
    xfrm = spPr.find(_q("a:xfrm"))
    spPr.insert(1 if xfrm is not None else 0, geom)


def _inherit_geometry(shape, inherited):
    """A placeholder with no xfrm of its own is positioned by its layout."""
    spPr = shape.find(_q("p:spPr"))
    if spPr is None or spPr.find(_q("a:xfrm")) is not None:
        return
    for source in inherited:
        source_spPr = source.find(_q("p:spPr"))
        xfrm = source_spPr.find(_q("a:xfrm")) if source_spPr is not None else None
        if xfrm is not None:
            spPr.insert(0, _copy(xfrm))
            return


_AUTOFIT = ("normAutofit", "spAutoFit", "noAutofit")


def _inherit_body_properties(shape, inherited):
    """Anchor, insets and autofit come from the layout when unstated."""
    bodyPr = shape.find(f"{_q('p:txBody')}/{_q('a:bodyPr')}")
    if bodyPr is None:
        return
    for source in inherited:
        source_bodyPr = source.find(f"{_q('p:txBody')}/{_q('a:bodyPr')}")
        if source_bodyPr is None:
            continue
        for name, value in source_bodyPr.attrib.items():
            if name not in bodyPr.attrib:
                bodyPr.set(name, value)
        # Autofit is a child element, not an attribute, and it is the one
        # that decides whether the text fits at all: these layouts set
        # <a:normAutofit/>, so the source shrinks a long client-challenge
        # paragraph to stay in its box. Inheriting only the attributes left
        # the copy at full size, one line taller, running into the stats
        # chips below it.
        if not any(bodyPr.find(_q(f"a:{fit}")) is not None for fit in _AUTOFIT):
            for fit in _AUTOFIT:
                found = source_bodyPr.find(_q(f"a:{fit}"))
                if found is not None:
                    bodyPr.append(_copy(found))
                    break


def _level_defaults(element, level):
    """The defRPr a lstStyle gives this indent level, if any."""
    if element is None:
        return None
    lstStyle = element if element.tag == _q("a:lstStyle") else \
        element.find(f"{_q('p:txBody')}/{_q('a:lstStyle')}")
    if lstStyle is None:
        return None
    lvl = lstStyle.find(_q(f"a:lvl{level + 1}pPr"))
    return lvl.find(_q("a:defRPr")) if lvl is not None else None


def _inherit_text(shape, inherited, txStyles, ph_type):
    """Push inherited run properties onto every run that lacks them.

    A run states only what it overrides; everything else comes from the
    layout placeholder, then the master placeholder, then the master's
    txStyles for that placeholder kind. Resolved most specific first, which
    is the order PowerPoint resolves them in.
    """
    txBody = shape.find(_q("p:txBody"))
    if txBody is None:
        return
    style = None
    if txStyles is not None:
        style = txStyles.find(_q(f"p:{_STYLE_FOR_TYPE.get(ph_type, 'otherStyle')}"))

    for paragraph in txBody.findall(_q("a:p")):
        pPr = paragraph.find(_q("a:pPr"))
        level = int(pPr.get("lvl", "0")) if pPr is not None else 0
        sources = [_level_defaults(shape, level)]
        sources += [_level_defaults(source, level) for source in inherited]
        sources.append(_level_defaults(style, level))
        sources = [s for s in sources if s is not None]
        _inherit_paragraph(paragraph, shape, inherited, style, level)
        if not sources:
            continue
        for run in paragraph.findall(_q("a:r")) + paragraph.findall(_q("a:endParaRPr")):
            rPr = run.find(_q("a:rPr")) if run.tag == _q("a:r") else run
            if rPr is None:
                rPr = etree.SubElement(run, _q("a:rPr"))
                run.insert(0, rPr)
            _apply_defaults(rPr, sources)


# Paragraph-level properties inherit exactly as run properties do, and
# leaving them out is visible: the case study's client-challenge paragraph is
# justified by its layout and came back ragged-right. Bullet definitions are
# deliberately NOT carried down -- a paragraph that states no bullet is
# usually a paragraph that wants none, and inheriting one puts dots in front
# of body copy.
_PARA_ATTRS = ("algn", "marL", "marR", "indent", "defTabSz", "rtl", "lvl")
# Spacing is expressed as child elements rather than attributes, and matters
# as much as alignment: the same paragraph set at 90% line spacing by its
# layout wrapped to twelve lines instead of eleven without it, and ran out of
# its box into the stats chips below.
_PARA_CHILDREN = ("lnSpc", "spcBef", "spcAft")


def _level_ppr(element, level):
    """The lvlNpPr a lstStyle gives this indent level, if any."""
    if element is None:
        return None
    lstStyle = element if element.tag == _q("a:lstStyle") else         element.find(f"{_q('p:txBody')}/{_q('a:lstStyle')}")
    if lstStyle is None:
        return None
    return lstStyle.find(_q(f"a:lvl{level + 1}pPr"))


def _inherit_paragraph(paragraph, shape, inherited, style, level):
    sources = [_level_ppr(shape, level)]
    sources += [_level_ppr(source, level) for source in inherited]
    sources.append(_level_ppr(style, level))
    sources = [s for s in sources if s is not None]
    if not sources:
        return
    pPr = paragraph.find(_q("a:pPr"))
    if pPr is None:
        pPr = etree.Element(_q("a:pPr"))
        paragraph.insert(0, pPr)
    for source in sources:
        for name in _PARA_ATTRS:
            if name not in pPr.attrib and source.get(name) is not None:
                pPr.set(name, source.get(name))
        for name in _PARA_CHILDREN:
            if pPr.find(_q(f"a:{name}")) is not None:
                continue
            found = source.find(_q(f"a:{name}"))
            if found is not None:
                pPr.insert(0, _copy(found))


def _apply_defaults(rPr, sources):
    for source in sources:
        for name in _RUN_ATTRS:
            if name not in rPr.attrib and source.get(name) is not None:
                rPr.set(name, source.get(name))
        for child in _RUN_CHILDREN:
            if rPr.find(_q(f"a:{child}")) is not None:
                continue
            found = source.find(_q(f"a:{child}"))
            if found is not None:
                rPr.append(_copy(found))


def _copy(element):
    return etree.fromstring(etree.tostring(element))


# ---------------------------------------------------------------------------
# 2. theme colours
# ---------------------------------------------------------------------------

def _theme_part(master):
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        return master.part.part_related_by(RT.THEME)
    except Exception:                                            # noqa: BLE001
        return None


# "+mj-lt" is not a font, it is "whatever this theme calls its major latin
# font" -- so a copied slide silently restyles itself in the destination's
# typography exactly as it restyles itself in its colours. The case study's
# title inherited +mj-lt from its layout and came out in Premion's heading
# font at a different weight. Resolved from the source theme for the same
# reason and at the same time as the colours.
_THEME_FONT_REFS = {"+mj-lt": ("majorFont", "latin"), "+mn-lt": ("minorFont", "latin"),
                    "+mj-ea": ("majorFont", "ea"), "+mn-ea": ("minorFont", "ea"),
                    "+mj-cs": ("majorFont", "cs"), "+mn-cs": ("minorFont", "cs")}


def _theme_fonts(master):
    """{"+mj-lt": "Proxima Nova", ...} for the source deck's theme."""
    theme = _theme_part(master)
    if theme is None:
        return {}
    root = etree.fromstring(theme.blob)
    scheme = root.find(f".//{_q('a:fontScheme')}")
    if scheme is None:
        return {}
    fonts = {}
    for ref, (group, script) in _THEME_FONT_REFS.items():
        group_element = scheme.find(_q(f"a:{group}"))
        if group_element is None:
            continue
        element = group_element.find(_q(f"a:{script}"))
        if element is None:
            continue
        # An empty typeface is an answer, not a missing one: these themes
        # leave the East Asian and complex-script fonts unset, and leaving
        # "+mn-ea" in place would let the DESTINATION theme name one. Writing
        # the empty string back reproduces what the source deck says.
        fonts[ref] = element.get("typeface") or ""
    return fonts


def _resolve_theme_fonts(slide_element, master):
    fonts = _theme_fonts(master)
    if not fonts:
        return
    for element in slide_element.iter():
        typeface = element.get("typeface")
        if typeface in fonts:
            element.set("typeface", fonts[typeface])


def _theme_palette(master):
    """{scheme name -> RRGGBB} for the source deck's theme."""
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        theme = master.part.part_related_by(RT.THEME)
    except Exception:                                            # noqa: BLE001
        return {}
    root = etree.fromstring(theme.blob)
    scheme = root.find(f".//{_q('a:clrScheme')}")
    palette = {}
    for entry in scheme if scheme is not None else []:
        name = etree.QName(entry).localname
        srgb = entry.find(_q("a:srgbClr"))
        if srgb is not None:
            palette[name] = srgb.get("val")
            continue
        sys_color = entry.find(_q("a:sysClr"))
        if sys_color is not None:
            palette[name] = sys_color.get("lastClr") or (
                "FFFFFF" if "light" in (sys_color.get("val") or "") else "000000")
    return palette


def _resolve_scheme_colors(slide_element, master):
    """Rewrite <a:schemeClr> as explicit RGB from the source deck's theme.

    A scheme reference is a pointer into whatever theme the slide is sitting
    in, so a copied slide silently repaints itself in the destination's
    palette -- which is how white-on-brand text arrived near-black on a dark
    background. The colour map is applied first: tx1/bg1 are not colours,
    they are indirections a master is free to swap, and reading them as
    dk1/lt1 without it inverts exactly the decks that swap them.

    Child transforms (lumMod, tint, alpha) are kept -- they modify whatever
    colour they hang off, so they mean the same thing on an srgbClr.
    """
    palette = _theme_palette(master)
    if not palette:
        return
    color_map = master.part._element.find(_q("p:clrMap"))
    mapping = dict(color_map.attrib) if color_map is not None else {}

    for scheme_ref in list(slide_element.iter(_q("a:schemeClr"))):
        value = scheme_ref.get("val")
        # phClr is "whatever colour this style is applied with" -- resolving
        # it would freeze a style's placeholder to one colour.
        if value in (None, "phClr"):
            continue
        rgb = palette.get(mapping.get(value, value))
        if rgb is None:
            continue
        scheme_ref.tag = _q("a:srgbClr")
        scheme_ref.set("val", rgb)
