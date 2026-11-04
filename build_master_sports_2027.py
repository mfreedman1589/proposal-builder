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
- retags the Motorsports Viewers slide `sport_viewership:motorsports`, now
  that Motorsports is a package (it was `sports_viewership_unmapped`).

Every new slide carries its `key:` line in the speaker notes. Slides come in
through assembly.copy_slide_into, the case-study cross-deck copy. Refuses to
run on a deck that already has a `sport:motorsports` slide, so it can't
stack a second set. Upload the output on dev only -- the nightly merge
activates it.
"""
import copy
import sys
from pathlib import Path

from pptx import Presentation
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

    for key in sports_lineup.RETIRED:
        for prefix in ("sport", "sport_viewership"):
            i = index_of(prs, f"{prefix}:{key}")
            if i is not None:
                assembly.delete_slide(prs, i)
                print(f"  removed {prefix}:{key}")

    for slide in prs.slides:
        if "MOTORSPORTS VIEWERS" in slide_map.extract_slide_text(slide).upper():
            set_key_line(slide, "sport_viewership:motorsports")
            print("  Motorsports Viewers -> sport_viewership:motorsports")

    prs.save(out)
    print(f"{before} slides -> {len(prs.slides)}; saved {out}")


if __name__ == "__main__":
    main(*sys.argv[1:3])
