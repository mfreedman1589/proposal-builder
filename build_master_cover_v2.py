"""Add the personalized cover (key: client_title_v2) to the proposal master.

Usage:
    python build_master_cover_v2.py <master_in.pptx> <reference.pptx> <out.pptx>

The reference (assets/client_title_v2.pptx, one slide, gitignored like every
.pptx) is the designed cover. It is NOT copied across decks -- that path has
corrupted decks before (CLAUDE.md, "Cross-deck copying"). Instead:

1. The master's own `client_title` slide is duplicated WITHIN the master
   (assembly.duplicate_slide), directly after the original, so the new slide
   has the original's layout ("2_Blank", the same layout the reference uses)
   and its image relationships.
2. The duplicate's shape tree is rebuilt to match the reference exactly: each
   reference shape's XML -- names, positions, sizes, fonts, colours, the
   gradient, the tokens -- with every image reference resolved to the master's
   OWN image part of identical bytes (the reference was built from this
   master, so all its pictures are already here; a picture that isn't is an
   error, never a re-encoded import). The reference's two unnamed rectangles
   take the names the cover's contract gives them: COVER_GRADIENT,
   COVER_ACCENT_BAR.
3. The background (a gradient fill on <p:bg>) is set from the reference, and
   the notes say `key: client_title_v2`.

The old `client_title` slide stays, so "Rebuild as presented" reproduces a
proposal sent with it. Refuses to run on a deck that already has the v2 cover.

Verify with --check <out.pptx> <reference.pptx> (package check, slide map,
shape-by-shape comparison) and render both slides through PowerPoint
(tests/test_cover_v2.py does that). Upload the output on dev only (inactive);
the nightly merge activates it.
"""
import copy
import hashlib
import sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

import assembly
import slide_map

V2_KEY = "client_title_v2"
RENAMES = {"Rectangle 11": "COVER_GRADIENT", "Rectangle 16": "COVER_ACCENT_BAR"}
COVER_SHAPES = ("COVER_BACKGROUND", "COVER_GRADIENT", "PREMION_LOGO", "CLIENT_LOGO",
                "COVER_EYEBROW", "COVER_HEADLINE", "COVER_SUBTITLE", "COVER_ACCENT_BAR",
                "COVER_FOOTER")
_R_ATTRS = (qn("r:embed"), qn("r:link"), qn("r:id"))


def _sha(blob):
    return hashlib.sha1(blob).hexdigest()


def _slide_by_key(prs, key):
    return next((s for s in prs.slides if slide_map.notes_key(s) == key), None)


def _media_by_hash(prs):
    out = {}
    for part in prs.part.package.iter_parts():
        if str(part.partname).startswith("/ppt/media/"):
            out.setdefault(_sha(part.blob), part)
    return out


def _referenced_rids(element):
    return {el.get(attr) for el in element.iter() for attr in _R_ATTRS if el.get(attr)}


def build(master_in, reference, out):
    prs = Presentation(master_in)
    if _slide_by_key(prs, V2_KEY) is not None:
        raise SystemExit(f"{master_in} already has a {V2_KEY} slide -- nothing to do.")
    source = _slide_by_key(prs, "client_title")
    if source is None:
        raise SystemExit(f"{master_in} has no client_title slide to duplicate.")
    ref_prs = Presentation(reference)
    ref = ref_prs.slides[0]
    if slide_map.notes_key(ref) != V2_KEY:
        raise SystemExit(f"{reference}'s slide isn't labelled key: {V2_KEY}.")
    if ref.slide_layout.name != source.slide_layout.name:
        raise SystemExit(f"Layout mismatch: reference {ref.slide_layout.name!r}, "
                         f"master client_title {source.slide_layout.name!r}.")

    dup = assembly.duplicate_slide(prs, source, insert_at=assembly.slide_index(prs, source) + 1)

    # Image relationships the rebuilt tree needs, by the bytes they point at.
    media = _media_by_hash(prs)
    rid_for_hash = {}
    for rId, rel in dup.part.rels.items():
        if rel.reltype.endswith("/image"):
            rid_for_hash.setdefault(_sha(rel.target_part.blob), rId)
    remap = {}
    for rId in _referenced_rids(ref._element.cSld):
        rel = ref.part.rels[rId]
        if not rel.reltype.endswith("/image"):
            raise SystemExit(f"Reference shape references a non-image part ({rel.reltype}).")
        digest = _sha(rel.target_part.blob)
        if digest not in rid_for_hash:
            part = media.get(digest)
            if part is None:
                raise SystemExit(f"Reference image {rel.target_part.partname} isn't in the master "
                                 f"-- refusing to import it.")
            rid_for_hash[digest] = dup.part.relate_to(part, rel.reltype)
        remap[rId] = rid_for_hash[digest]

    # Rebuild the shape tree from the reference.
    tree = dup.shapes._spTree
    for child in list(tree):
        if child.tag not in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            tree.remove(child)
    for child in ref.shapes._spTree:
        if child.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        clone = copy.deepcopy(child)
        for el in clone.iter():
            for attr in _R_ATTRS:
                if el.get(attr) in remap:
                    el.set(attr, remap[el.get(attr)])
        tree.append(clone)
    for shape in dup.shapes:
        if shape.name in RENAMES:
            shape.name = RENAMES[shape.name]

    # Background: the reference's gradient, replacing whatever was there.
    csld = dup._element.cSld
    old_bg = csld.find(qn("p:bg"))
    if old_bg is not None:
        csld.remove(old_bg)
    ref_bg = ref._element.cSld.find(qn("p:bg"))
    if ref_bg is not None:
        csld.insert(0, copy.deepcopy(ref_bg))

    # Image relationships nothing on the slide uses any more.
    used = _referenced_rids(dup._element)
    for rId, rel in list(dup.part.rels.items()):
        if rel.reltype.endswith("/image") and rId not in used:
            dup.part.drop_rel(rId)

    dup.notes_slide.notes_text_frame.text = f"key: {V2_KEY}"
    prs.save(out)
    print(f"Wrote {out}: {V2_KEY} added after client_title "
          f"(slide {assembly.slide_index(prs, dup) + 1} of {len(prs.slides)}).")


def _shape_signature(shape):
    """Everything about a shape that has to match the reference: its XML with
    relationship ids replaced by the bytes they point at, and the cNvPr id
    and creationId extension (renumbered by PowerPoint, meaningless) dropped."""
    el = copy.deepcopy(shape._element)
    for node in el.iter():
        for attr in _R_ATTRS:
            rid = node.get(attr)
            if rid:
                node.set(attr, _sha(shape.part.rels[rid].target_part.blob)[:12])
        if node.tag == qn("p:cNvPr"):
            node.attrib.pop("id", None)
            if node.get("name") in RENAMES:
                node.set("name", RENAMES[node.get("name")])
    for ext in el.iter(qn("a:extLst")):
        ext.getparent().remove(ext)
    return etree.tostring(el)


def check(out, reference):
    """Problems with a built deck's v2 cover, as a list of strings."""
    import package_check
    problems = list(package_check.check_package(out) or [])
    prs = Presentation(out)
    v2 = _slide_by_key(prs, V2_KEY)
    old = _slide_by_key(prs, "client_title")
    if v2 is None:
        return problems + [f"no {V2_KEY} slide"]
    if old is None:
        problems.append("the old client_title slide is gone")
    ref = Presentation(reference).slides[0]
    names = [s.name for s in v2.shapes]
    for name in COVER_SHAPES:
        if names.count(name) != 1:
            problems.append(f"{name}: {names.count(name)} shapes")
    ref_sigs = [_shape_signature(s) for s in ref.shapes]
    new_sigs = [_shape_signature(s) for s in v2.shapes]
    if ref_sigs != new_sigs:
        for i, (a, b) in enumerate(zip(ref_sigs, new_sigs)):
            if a != b:
                problems.append(f"shape {i} ({names[i]}) differs from the reference")
        if len(ref_sigs) != len(new_sigs):
            problems.append(f"{len(new_sigs)} shapes, reference has {len(ref_sigs)}")
    bg_new = v2._element.cSld.find(qn("p:bg"))
    bg_ref = ref._element.cSld.find(qn("p:bg"))
    if (etree.tostring(bg_new) if bg_new is not None else None) != (
            etree.tostring(bg_ref) if bg_ref is not None else None):
        problems.append("background differs from the reference")
    if v2.slide_layout.name != ref.slide_layout.name:
        problems.append(f"layout {v2.slide_layout.name!r} vs reference {ref.slide_layout.name!r}")
    classified = slide_map._classify_all(prs, verbose=False)
    if not any(key == V2_KEY and source == slide_map.SOURCE_NOTES
               for key, source in classified.values()):
        problems.append(f"slide_map doesn't classify the cover as {V2_KEY} from its notes")
    if any(source == slide_map.SOURCE_CARRY_FORWARD for _key, source in classified.values()):
        problems.append("a slide resolves by carry-forward (the update page would block it)")
    return problems


if __name__ == "__main__":
    if sys.argv[1:2] == ["--check"]:
        found = check(sys.argv[2], sys.argv[3])
        print("\n".join(found) or "OK: the v2 cover matches the reference; the deck opens cleanly.")
        sys.exit(1 if found else 0)
    build(*sys.argv[1:4])
