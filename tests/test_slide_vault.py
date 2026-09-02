"""Slide vault grafting: where vault slides land, especially in combination
with the case-study vault they were built to sit alongside.

    python tests/test_slide_vault.py

The invariant that matters most here is ORDER, and it is the one kind of bug
this project has already shipped once for case studies (a Total TV deck's
plan-slide variant didn't match the insert-point lookup, so case studies fell
through to the end of the deck, after the media plan -- see
test_case_study_import.py's docstring). With two vaults now grafting into the
same deck at overlapping anchors, the same class of bug is easy to reintroduce
silently: a deck with slides in the wrong order still opens and generates
fine, so nothing short of checking the actual slide order catches it. This
test builds ONE deck carrying case studies AND vault slides at all three
placements (front / before_plan / appendix) and asserts the combined order.
"""

import copy
import os
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

SOURCE = REPO / "case_studies_source"
CASE_STUDY_DECK = SOURCE / "PREMION_Case Study_Regional Furniture + Mattress Retailer.pptx"
# Three different source decks stand in for three different colleagues'
# uploads, same as a real vault would hold slides from unrelated decks --
# and deliberately NOT two slides of the same deck: a real case study deck's
# two slides can restate nearly identical text in a different shape order
# (found while writing this test, on the Career Training College deck),
# which makes a text-content marker ambiguous between them.
FRONT_VAULT_DECK = SOURCE / "PREMION_Case Study_Exterior Home Remodeling Company.pptx"
BEFORE_PLAN_VAULT_DECK = SOURCE / "PREMION_Case Study_Career Training College.pptx"
APPENDIX_VAULT_DECK = SOURCE / "PREMION_Case Study_Regional Bank.pptx"

failures = []


def check(label, ok, detail=""):
    print(f"  {'PASS' if ok else 'FAIL'}  {label}{('  ' + str(detail)) if detail else ''}")
    if not ok:
        failures.append(label)


def build(total_tv=False, preset="standard"):
    master_path, _, warning = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    if master_path is None:
        return None, warning
    selections = copy.deepcopy(assembly.SELECTIONS)
    selections["preset"], selections["market"] = preset, "DC"
    selections["products"] = dict(selections["products"])
    selections["products"]["total_tv"] = total_tv
    prs, _, _ = assembly.build_presentation(master_path, selections)

    case_studies = [{"path": str(CASE_STUDY_DECK), "slides": None}]
    vault_slides = [
        {"path": str(FRONT_VAULT_DECK), "slides": [0], "placement": "front"},
        {"path": str(BEFORE_PLAN_VAULT_DECK), "slides": [0], "placement": "before_plan"},
        {"path": str(APPENDIX_VAULT_DECK), "slides": [0], "placement": "appendix"},
    ]

    # The order this is meant to mirror: app.py grafts case studies, THEN
    # vault slides, so the existing call site is untouched and the new one
    # goes directly after it.
    assembly.append_case_studies(prs, case_studies)
    assembly.append_vault_slides(prs, vault_slides)
    return prs, None


def slide_texts(prs):
    return [slide_map.extract_slide_text(s) for s in prs.slides]


def find_all(texts, needle):
    return [i for i, t in enumerate(texts, start=1) if needle in t]


def check_combined_order(label="standard"):
    print(f"\n--- combined order ({label}) ---")
    prs, warning = build()
    if prs is None:
        print(f"  SKIP  master deck unavailable: {warning}")
        return None

    texts = slide_texts(prs)
    total = len(texts)

    cover = find_all(texts, "COBRAND") or [1]
    front_vault = find_all(texts, "Exterior Home Remodeling")
    case_study = find_all(texts, "Furniture")
    before_plan_vault = find_all(texts, "Career Training College")
    appendix_vault = find_all(texts, "Regional Bank")
    plan = find_all(texts, assembly._placeholder("PLAN_TITLE")) or find_all(texts, "TACTIC")

    check("front vault slide found", bool(front_vault), front_vault)
    check("case study slide(s) found", bool(case_study), case_study)
    check("before_plan vault slide found", bool(before_plan_vault), before_plan_vault)
    check("appendix vault slide found", bool(appendix_vault), appendix_vault)
    check("plan slide found", bool(plan), plan)
    if not all([front_vault, case_study, before_plan_vault, appendix_vault, plan]):
        return False

    check("front vault slide sits immediately after the cover",
          front_vault[0] == cover[0] + 1,
          f"cover {cover}, front vault {front_vault}")

    check("before_plan vault slide comes AFTER every case study slide",
          min(before_plan_vault) > max(case_study),
          f"case study {case_study}, before_plan vault {before_plan_vault}")

    check("before_plan vault slide comes BEFORE the plan slide",
          max(before_plan_vault) < min(plan),
          f"before_plan vault {before_plan_vault}, plan {plan}")

    check("appendix vault slide is the last slide in the deck",
          appendix_vault == [total],
          f"appendix vault {appendix_vault}, deck has {total} slides")

    out = deck_render.render_root() / "slide_vault"
    out.mkdir(parents=True, exist_ok=True)
    path = str(out / "deck.pptx")
    prs.save(path)
    problems = package_check.check_package(path)
    check("package is structurally clean", not problems, problems[:4])

    return True


if __name__ == "__main__":
    check_combined_order()
    print(f"\n{len(failures)} failure(s)." if failures else "\nAll checks passed.")
    sys.exit(1 if failures else 0)
