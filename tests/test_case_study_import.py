"""Imported case study slides: where they land, and whether they survive the trip.

    python tests/test_case_study_import.py            # the two furniture ones
    python tests/test_case_study_import.py --sweep    # every placeholder-based one

Two bugs, both found by looking at a real deck rather than by any assertion:

  1. Case studies inserted *after* the media plan on a Total TV proposal. The
     insert point matched the condition_key "proposal_template" exactly, and
     Total TV's variants key "proposal_template_total_tv:<market>", so the
     lookup fell through to appending at the end.
  2. Copied slides lost every piece of formatting they inherited rather than
     stated, because they were attached to the destination's Blank layout:
     `<a:schemeClr>` resolved against Premion's theme and "+mj-lt" against
     Premion's fonts, so brand-white headings rendered near-black.

Both are invisible to a check that only asks whether the deck opens, which is
why this asserts the two invariants directly: nothing theme-dependent may
survive on an imported slide, and case studies must end before the plan.
18 of the vault's 22 case studies are placeholder-based, so this is most of
them, not an edge case.
"""

import copy
import os
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import assembly                                   # noqa: E402
import db                                         # noqa: E402
import deck_render                                # noqa: E402
import package_check                              # noqa: E402
import slide_map                                  # noqa: E402
import text_metrics                               # noqa: E402
from pptx import Presentation                     # noqa: E402

SOURCE = REPO / "case_studies_source"
FURNITURE = ["PREMION_Case Study_Regional Furniture + Mattress Retailer.pptx",
             "PREMION_Case Study_Fine Jewelry Retailer.pptx"]

# A reference that only means something in the deck it came from. Every one of
# these on an imported slide is a piece of formatting that will be reinterpreted
# by whatever theme the proposal happens to use.
THEME_DEPENDENT = re.compile(r'<a:schemeClr |typeface="\+m[jn]-')

failures = []


def check(label, ok, detail=""):
    print(f"  {'PASS' if ok else 'FAIL'}  {label}{('  ' + str(detail)) if detail else ''}")
    if not ok:
        failures.append(label)


def placeholder_based(path):
    try:
        prs = Presentation(str(path))
    except Exception:                                            # noqa: BLE001
        return False
    return any("<p:ph " in shape._element.xml
               for slide in prs.slides for shape in slide.shapes)


def build(case_studies, total_tv, options=1, preset="standard"):
    master_path, _, warning = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    if master_path is None:
        return None, warning
    selections = copy.deepcopy(assembly.SELECTIONS)
    selections["preset"], selections["market"] = preset, "DC"
    selections["products"] = dict(selections["products"])
    selections["products"]["total_tv"] = total_tv
    prs, _, _ = assembly.build_presentation(master_path, selections)
    position = assembly.case_study_insert_index(prs)
    copied = assembly.append_case_studies(
        prs, [{"path": str(SOURCE / name), "slides": None} for name in case_studies])
    # The exact range, from the insert point and the count -- not inferred
    # from slide text. Matching on the source deck's title misses a case
    # study's second slide whenever it is titled differently, which reads as
    # a gap in an otherwise contiguous block.
    return prs, list(range(position + 1, position + copied + 1))


def check_position(case_studies):
    print("\n--- where the case studies land ---")
    for label, total_tv, preset in (("standard", False, "standard"),
                                    ("Total TV (DC)", True, "standard"),
                                    ("extended", False, "extended")):
        prs, imported = build(case_studies, total_tv, preset=preset)
        if prs is None:
            print(f"  SKIP  {label}: master deck unavailable")
            return None
        plan = [i for i, s in enumerate(prs.slides, 1)
                if assembly._placeholder("PLAN_TITLE") in slide_map.extract_slide_text(s)
                or "TACTIC" in slide_map.extract_slide_text(s)]
        check(f"{label}: case studies end before the plan slide",
              bool(imported) and bool(plan) and max(imported) < min(plan),
              f"case studies {imported}, plan {plan}")
        check(f"{label}: case studies are contiguous",
              imported == list(range(min(imported), max(imported) + 1)), imported)
    return True


def check_formatting(prs, imported, label="furniture"):
    print(f"\n--- what survived the copy ({label}) ---")
    slides = [s for i, s in enumerate(prs.slides, 1) if i in imported]

    leftover = []
    for index, slide in zip(imported, slides):
        for shape in slide.shapes:
            hits = THEME_DEPENDENT.findall(shape._element.xml)
            if hits:
                leftover.append((index, shape.name, len(hits)))
    check("no theme-dependent colour or font references remain",
          not leftover, leftover[:4])

    stray = [(i, sh.name) for i, s in zip(imported, slides) for sh in s.shapes
             if "<p:ph " in sh._element.xml]
    check("no placeholders remain to inherit from the destination", not stray, stray[:4])

    # Overflow: a text frame taller than the shape holding it. Measured, not
    # estimated -- text_metrics knows the real fonts.
    overflowing = []
    for index, slide in zip(imported, slides):
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.height or not shape.width:
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            height = assembly._estimate_frame_height(shape.text_frame, shape.width, 1.0)
            if height > shape.height * 1.05:
                overflowing.append(
                    (index, shape.name, round(height / 914400, 2),
                     round(shape.height / 914400, 2)))
    # Reported, not asserted: the source decks' own live autofit shrinks some
    # of this text and a flattened shape can't reproduce that exactly. A
    # regression would show up as this list growing.
    print(f"  NOTE  {len(overflowing)} text frame(s) measure taller than their shape"
          f"{': ' + str(overflowing[:3]) if overflowing else ''}")
    return imported


def main():
    sweep = "--sweep" in sys.argv
    if not SOURCE.exists():
        print(f"SKIP -- {SOURCE} not present (case study decks are gitignored).")
        return 0

    case_studies = FURNITURE
    if sweep:
        case_studies = sorted(p.name for p in SOURCE.glob("*.pptx")
                              if placeholder_based(p))
        print(f"sweeping {len(case_studies)} placeholder-based case studies")
    missing = [n for n in case_studies if not (SOURCE / n).exists()]
    if missing:
        print(f"SKIP -- missing {missing}")
        return 0

    if check_position(case_studies) is None:
        return 0

    prs, imported = build(case_studies, total_tv=True)
    check_formatting(prs, imported)

    out = deck_render.render_root() / "case_study_import"
    out.mkdir(parents=True, exist_ok=True)
    path = str(out / "deck.pptx")
    prs.save(path)
    check("package is structurally clean", not package_check.check_package(path),
          package_check.check_package(path))

    if deck_render.renderer_available():
        try:
            rendered = deck_render.render_deck(path, str(out))
            check("PowerPoint opens and renders the deck", bool(rendered), len(rendered))
            print(f"  images: {out}")
        except Exception as exc:                                 # noqa: BLE001
            check("PowerPoint opens and renders the deck", False, str(exc)[:90])
    else:
        print("  SKIP  rendering (needs Windows with PowerPoint)")

    print("\n" + ("ALL PASS" if not failures else f"{len(failures)} FAILED: {failures}"))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
