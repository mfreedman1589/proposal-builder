"""Put the 2027 sports lineup into the proposal master.

Usage:
    python build_master_sports_2027.py <proposal_master_in.pptx> <out.pptx>

TEGNA now sources live sports inventory directly through the Premion DSP
(lineup and rates: sports_lineup.py). This build, against the CURRENT master
(version 13 at the time of writing):

- replaces every current package's `sport:<key>` slide with its 2027
  one-sheet (PREMION_*_2027.pptx, in the project folder, gitignored like
  every .pptx), at the old slide's position;
- adds `sport:motorsports` (after golf) and `sport:sports_select` (after Live
  Sports Select), which had no slide;
- replaces the sports overview ("Choose the Live, In-Game Sports", key
  `sports`) with PREMION_All_Sports_2027, fixing two errors in that file: the
  Hockey tile read "NFL Regular Season", and the Basketball and Football
  tiles ended a line on a stray "|" separator (now a line break);
- removes the retired packages' package and viewership slides -- an old
  proposal rebuilds from its own deck version, so nothing it needs is lost;
- brings the old copy into line with the 2027 lineup (Matt, 2026-11-04):
  "All Live Sports" -> "Live Sports Select" on its one-sheet and the
  overview tile, and "Why Choose PREMION for In-Game Live Sports?" no longer
  promises playoffs or championships ("every game of the season");
- slims the deck (Matt, 2026-11-04 -- v14 sat within 2 MiB of the 50 MiB
  cap): removes the two never-selectable viewership slides (World Cup, March
  Madness), stores two photos whose transparency is only a 1-pixel edge as
  JPEG (the legal vertical slide's, the sports overview's), and drops every
  layout and master no slide uses (the "Blank" layouts stay -- slide copying
  needs one);
- retags the Motorsports Viewers slide `sport_viewership:motorsports`, now
  that Motorsports is a package (it was `sports_viewership_unmapped`).

Every new slide carries its `key:` line in the speaker notes. Slides come in
through assembly.copy_slide_into, the case-study cross-deck copy. Refuses to
run on a deck that already has a `sport:motorsports` slide, so it can't
stack a second set. Upload the output on dev only -- the nightly merge
activates it.
"""
import copy
import io
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

import assembly
import slide_map
import sports_lineup

HERE = Path(__file__).resolve().parent
ONE_SHEETS = {
    "all_live_sports": "PREMION_Live_Sports_Select_2027.pptx",
    "mlb_reg": "PREMION_MLB_Regular_Season_2027.pptx",
    "nba_reg": "PREMION_NBA_Regular_Season_2027.pptx",
    "nhl_reg": "PREMION_NHL_Regular_Season_2027.pptx",
    "ncaa_basketball": "PREMION_NCAAMB_Regular_Season_2027.pptx",
    "ncaaf_reg": "PREMION_NCAAF_Regular_Season_2027.pptx",
    "ncaaf_conference": "PREMION_NCAAF_Conferences_2027.pptx",
    "nfl_reg": "PREMION_NFL_Regular_Season_2027.pptx",
    "golf": "PREMION_Golf_2027.pptx",
    "motorsports": "PREMION_Motorsports_2027.pptx",
    "sports_select": "PREMION_Sports_Select_2027.pptx",
}
OVERVIEW = "PREMION_All_Sports_2027.pptx"
# Where a package with no old slide goes: right after this one's slide.
ADD_AFTER = {"motorsports": "golf", "sports_select": "all_live_sports"}


def keys_by_index(prs):
    return dict(sorted(slide_map.build_slide_map_from_prs(prs).items()))


def index_of(prs, key):
    found = [n - 1 for n, k in keys_by_index(prs).items() if k == key]
    return found[0] if len(found) == 1 else None


def set_key(slide, key):
    """The slide's speaker notes, reduced to its key line."""
    slide.notes_slide.notes_text_frame.text = f"key: {key}"


def set_key_line(slide, key):
    """Replace only the `key:` line in existing notes, keeping the rest."""
    tf = slide.notes_slide.notes_text_frame
    lines = [ln for ln in tf.text.splitlines() if not ln.strip().lower().startswith("key:")]
    tf.text = "\n".join(lines + [f"key: {key}"]).strip()


def overview_index(prs):
    """The old overview, by its own headline. Not by its "SPORTS CONTENT
    PACKAGES" eyebrow: "Why Choose PREMION Live Sports?" carries the same one
    (a first build replaced that slide instead -- caught in the render)."""
    found = [i for i, slide in enumerate(prs.slides)
             if "Choose the Live, In-Game Sports" in slide_map.extract_slide_text(slide)]
    if len(found) != 1:
        sys.exit(f"Expected one sports overview slide, found {len(found)}.")
    return found[0]


def separator_to_break(run_el):
    """Turn a "|" run that ends a wrapped line into a line break, dropping
    the spacing run after it."""
    nxt = run_el.getnext()
    if nxt is not None and nxt.tag == qn("a:r") and not (nxt.findtext(qn("a:t")) or "").strip():
        nxt.getparent().remove(nxt)
        nxt = run_el.getnext()
    if nxt is not None and nxt.tag == qn("a:br"):
        # A break already follows (the Basketball tile): the bar just goes.
        prev = run_el.getprevious()
        run_el.getparent().remove(run_el)
        if prev is not None and prev.tag == qn("a:r"):
            t = prev.find(qn("a:t"))
            t.text = (t.text or "").rstrip()
        return
    br = run_el.makeelement(qn("a:br"), {})
    rpr = run_el.find(qn("a:rPr"))
    if rpr is not None:
        br.append(copy.deepcopy(rpr))
    run_el.addprevious(br)
    run_el.getparent().remove(run_el)
    # The run before the break keeps no trailing space.
    prev = br.getprevious()
    if prev is not None and prev.tag == qn("a:r"):
        t = prev.find(qn("a:t"))
        t.text = (t.text or "").rstrip()


def fix_overview(slide):
    """The All Sports one-sheet's two errors (see the module docstring).
    Returns what it changed, so a corrected source file is noticed."""
    changed = []
    tiles = {}
    for shape in slide.shapes:
        for sub in getattr(shape, "shapes", []):
            if sub.has_text_frame and sub.text_frame.paragraphs:
                tiles[sub.text_frame.paragraphs[0].text.strip()] = sub
    hockey = tiles.get("Hockey")
    if hockey is not None:
        for run in hockey.text_frame.paragraphs[1].runs:
            if run.text.strip() == "NFL Regular Season":
                run.text = "NHL Regular Season"
                changed.append("Hockey: NFL -> NHL Regular Season")
    for tile, cut_at in (("Basketball", 1), ("Football", 2)):
        shape = tiles.get(tile)
        if shape is None:
            continue
        bars = [r for r in shape.text_frame.paragraphs[1].runs if r.text.strip() == "|"]
        if len(bars) >= cut_at:
            separator_to_break(bars[cut_at - 1]._r)
            changed.append(f"{tile}: line-ending '|' -> line break")
    return changed


# (slide finder, old text, new text) -- each must match a single run, so the
# run's own formatting is kept; a source file already corrected is reported.
COPY_FIXES = [
    ("sport:all_live_sports", "All Live Sports Package", "Live Sports Select Package"),
    ("overview", "All Live Sports", "Live Sports Select"),
    ("why_choose", "from regular season matchups through playoffs and championship moments.",
     "from the first game through the last."),
    ("why_choose", "sports, from season to playoffs.", "sports, every game of the season."),
]


def _runs(shapes):
    for shape in shapes:
        if getattr(shape, "shape_type", None) == 6:
            yield from _runs(shape.shapes)
        elif shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                yield from paragraph.runs


def replace_in_runs(slide, old, new):
    """Replace `old` inside any single run. Returns how many runs changed."""
    changed = 0
    for run in _runs(slide.shapes):
        if old in run.text and new not in run.text:
            run.text = run.text.replace(old, new)
            changed += 1
    return changed


def fix_copy(prs):
    def find(which):
        for slide in prs.slides:
            text = slide_map.extract_slide_text(slide)
            if which == "overview" and "Choose the Live Sports Packages" in text:
                return slide
            if which == "why_choose" and "Why Choose PREMION" in text:
                return slide
            if slide_map.notes_key(slide) == which:
                return slide
        sys.exit(f"Couldn't find the {which} slide for a copy fix.")
    for which, old, new in COPY_FIXES:
        n = replace_in_runs(find(which), old, new)
        print(f"  copy ({which}): {old!r} -> {new!r}" if n else
              f"  copy ({which}): {old!r} not found -- already corrected in the source?")


# Slides whose large photo is transparent only along a 1-pixel edge -- stored
# as JPEG with that edge filled from its neighbours (approved for these two;
# the GDW host headshot has the same shape and is left alone).
EDGE_JPEG_KEYS = ("vertical:legal", "sports")
KEEP_LAYOUTS = {"blank", "1_blank", "3_blank"}


def _edge_only_alpha(image):
    if image.mode not in ("RGBA", "LA", "P"):
        return False
    alpha = image.convert("RGBA").getchannel("A")
    w, h = alpha.size
    return (alpha.getextrema()[0] < 255
            and alpha.crop((1, 1, w - 1, h - 1)).getextrema()[0] == 255)


def edge_photos_to_jpeg(prs):
    taken = {str(p.partname) for p in prs.part.package.iter_parts()}
    for slide in prs.slides:
        if slide_map.notes_key(slide) not in EDGE_JPEG_KEYS:
            continue
        for blip in slide._element.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            part = slide.part.related_part(rid) if rid else None
            if part is None or not str(part.partname).endswith(".png") or len(part.blob) < 500_000:
                continue
            image = Image.open(io.BytesIO(part.blob))
            image.load()
            if not _edge_only_alpha(image):
                continue
            rgb = image.convert("RGB")
            w, h = rgb.size
            # The edge's colour under transparency is arbitrary: copy the
            # neighbouring row/column over it before dropping the alpha.
            rgb.paste(rgb.crop((0, 1, w, 2)), (0, 0))
            rgb.paste(rgb.crop((0, h - 2, w, h - 1)), (0, h - 1))
            rgb.paste(rgb.crop((1, 0, 2, h)), (0, 0))
            rgb.paste(rgb.crop((w - 2, 0, w - 1, h)), (w - 1, 0))
            buf = io.BytesIO()
            rgb.save(buf, "JPEG", quality=85, optimize=True, progressive=True)
            before = len(part.blob)
            name, n = str(part.partname)[:-len(".png")] + ".jpeg", 1
            while name in taken:
                name, n = f"{str(part.partname)[:-len('.png')]}-{n}.jpeg", n + 1
            taken.add(name)
            part._blob = buf.getvalue()
            part.partname = PackURI(name)
            part._content_type = "image/jpeg"
            print(f"  {slide_map.notes_key(slide)}: {name.split('/')[-1]} as JPEG, "
                  f"{before / 2**20:.2f} -> {len(part.blob) / 2**20:.2f} MiB")


def drop_unused_layouts(prs):
    used = {s.slide_layout.part.partname for s in prs.slides}
    dropped = 0
    for master in prs.slide_masters:
        id_list = master._element.find(qn("p:sldLayoutIdLst"))
        for layout, entry in list(zip(master.slide_layouts, list(id_list))):
            if layout.part.partname in used or layout.name.strip().lower() in KEEP_LAYOUTS:
                continue
            master.part.drop_rel(entry.get(qn("r:id")))
            id_list.remove(entry)
            dropped += 1
    masters = prs.part._element.find(qn("p:sldMasterIdLst"))
    dropped_masters = 0
    for master, entry in list(zip(prs.slide_masters, list(masters))):
        if not len(master.slide_layouts):
            prs.part.drop_rel(entry.get(qn("r:id")))
            masters.remove(entry)
            dropped_masters += 1
    print(f"  dropped {dropped} unused layouts and {dropped_masters} unused master(s)")


def main(master_in, out):
    prs = Presentation(master_in)
    if index_of(prs, "sport:motorsports") is not None:
        sys.exit("This deck already has the 2027 sports slides.")
    cache = assembly.ImportCache(prs)
    before = len(prs.slides)

    for key, filename in ONE_SHEETS.items():
        src = HERE / filename
        old = index_of(prs, f"sport:{key}")
        if old is None:
            anchor = index_of(prs, f"sport:{ADD_AFTER[key]}")
            position = anchor + 1
        else:
            position = old
        new = assembly.copy_slide_into(str(src), 0, prs, position=position, cache=cache)
        set_key(new, f"sport:{key}")
        if old is not None:
            assembly.delete_slide(prs, old + 1)
        print(f"  sport:{key}: {'replaced' if old is not None else 'added'} at slide {position + 1}")

    at = overview_index(prs)
    new = assembly.copy_slide_into(str(HERE / OVERVIEW), 0, prs, position=at, cache=cache)
    set_key(new, "sports")
    assembly.delete_slide(prs, at + 1)
    print(f"  overview replaced at slide {at + 1}: {', '.join(fix_overview(new)) or 'no fixes needed'}")

    fix_copy(prs)

    for key in sports_lineup.RETIRED:
        for prefix in ("sport", "sport_viewership"):
            i = index_of(prs, f"{prefix}:{key}")
            if i is not None:
                assembly.delete_slide(prs, i)
                print(f"  removed {prefix}:{key}")

    # By heading, not by key: Motorsports Viewers is still "unmapped" here
    # until the retag below.
    for i in reversed([i for i, s in enumerate(prs.slides)
                       if slide_map.notes_key(s) == "sports_viewership_unmapped"
                       and ("World Cup Viewers" in slide_map.extract_slide_text(s)
                            or "March Madness" in slide_map.extract_slide_text(s))]):
        assembly.delete_slide(prs, i)
        print(f"  removed never-selectable slide {i + 1}")

    for slide in prs.slides:
        if "MOTORSPORTS VIEWERS" in slide_map.extract_slide_text(slide).upper():
            set_key_line(slide, "sport_viewership:motorsports")
            print("  Motorsports Viewers -> sport_viewership:motorsports")

    edge_photos_to_jpeg(prs)
    drop_unused_layouts(prs)
    prs.save(out)
    print(f"{before} slides -> {len(prs.slides)}; saved {out}")


if __name__ == "__main__":
    main(*sys.argv[1:3])
