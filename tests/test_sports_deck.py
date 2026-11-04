"""The proposal master's sports slides match the 2027 lineup (local only).

    python tests/test_sports_deck.py

Reads the newest local master (TEGNA_MASTER_DECK_*.pptx, gitignored) and
SKIPs when there is none, or when it predates the 2027 lineup (no
`sport:motorsports` slide). What it guards:
  * every current package has exactly one `sport:<key>` slide, keyed in its
    speaker notes, and a proposal selecting all of them gets every one;
  * no retired package has a package or viewership slide (old proposals
    rebuild from their own deck version);
  * the overview is the 2027 All Sports slide with the source file's two
    errors fixed (Hockey reads NHL; no line ends on a "|");
  * Motorsports Viewers is `sport_viewership:motorsports`;
  * the slimming held: no World Cup / March Madness slide, no unused layout;
  * the package is a valid zip (no duplicate part names -- the optimizer's
    PNG->JPEG rename once collided here) and fits the storage limit.
"""
import collections
import copy
import os
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

from pptx import Presentation                  # noqa: E402

import assembly                                # noqa: E402
import db                                      # noqa: E402
import slide_map                               # noqa: E402
import sports_lineup as sl                     # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def newest_master():
    decks = sorted(REPO.glob("TEGNA_MASTER_DECK_*.pptx"), key=lambda p: p.stat().st_mtime)
    return decks[-1] if decks else None


def main():
    path = newest_master()
    if path is None:
        print("SKIP: no local master deck")
        return 0
    prs = Presentation(str(path))
    keys = slide_map.build_slide_map_from_prs(prs)
    if "sport:motorsports" not in keys.values():
        print(f"SKIP: {path.name} predates the 2027 sports lineup")
        return 0
    print(f"Master: {path.name}")
    counts = collections.Counter(keys.values())
    notes = {n: slide_map.notes_key(s) for n, s in enumerate(prs.slides, 1)}
    for key in sl.CURRENT_KEYS:
        slides = [n for n, k in keys.items() if k == f"sport:{key}"]
        check(f"sport:{key}: one slide, keyed in its notes",
              len(slides) == 1 and notes[slides[0]] == f"sport:{key}", slides)
    for key in sl.RETIRED:
        check(f"no slide for retired {key}",
              not counts[f"sport:{key}"] and not counts[f"sport_viewership:{key}"])
    check("Motorsports Viewers is sport_viewership:motorsports",
          counts["sport_viewership:motorsports"] == 1)
    check("the never-selectable World Cup and March Madness slides are gone",
          not counts["sports_viewership_unmapped"]
          and not any(h in slide_map.extract_slide_text(s) for s in prs.slides
                      for h in ("World Cup Viewers", "March Madness")))
    check("only layouts some slide uses are kept (plus Blank)",
          all(l.part.partname in {s.slide_layout.part.partname for s in prs.slides}
              or l.name.strip().lower() in ("blank", "1_blank", "3_blank")
              for m in prs.slide_masters for l in m.slide_layouts))

    overview = [s for s in prs.slides if "Choose the Live Sports Packages"
                in slide_map.extract_slide_text(s)]
    check("the 2027 overview is in, the old one out",
          len(overview) == 1 and not any("Choose the Live, In-Game Sports"
                                         in slide_map.extract_slide_text(s) for s in prs.slides))
    if overview:
        text = slide_map.extract_slide_text(overview[0])
        check("Hockey reads NHL Regular Season", "Hockey || NHL Regular Season" in text
              or "Hockey" in text and "NFL Regular Season" not in text.split("Hockey")[1][:40], text)
        lines = [ln.strip() for shape in overview[0].shapes for sub in getattr(shape, "shapes", [])
                 if sub.has_text_frame for ln in sub.text_frame.text.replace("\x0b", "\n").split("\n")]
        check("no tile line ends on a '|'", not [ln for ln in lines if ln.endswith("|")], lines)

    with zipfile.ZipFile(path) as z:
        dupes = [n for n, c in collections.Counter(z.namelist()).items() if c > 1]
    check("no duplicate part names", not dupes, dupes)
    check("fits the storage limit", path.stat().st_size <= db.DEFAULT_STORAGE_LIMIT,
          path.stat().st_size)

    sel = copy.deepcopy(assembly.SELECTIONS)
    sel["products"]["live_sports"] = {"enabled": True, "sports": list(sl.CURRENT_KEYS),
                                      "viewership": True}
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "all_sports.pptx"
        assembly.assemble(str(path), str(out), sel, assembly.FILL_DATA)
        built = {slide_map.notes_key(s) for s in Presentation(str(out)).slides}
    missing = [k for k in sl.CURRENT_KEYS if f"sport:{k}" not in built]
    check("a proposal selecting every package gets every package slide", not missing, missing)
    check("...and the viewership slide of each package that has one",
          {"sport_viewership:nba_reg", "sport_viewership:motorsports"} <= built)

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
