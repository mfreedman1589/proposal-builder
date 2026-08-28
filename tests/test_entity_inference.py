"""FLOW_REWORK_PLAN.md Phase 3, commit 7: `avails_pdf_import.infer_entities`
-- deterministic, offline, no LLM, no network -- against every real avails
PDF on hand. Pure per-document partition checks: which groups the heuristic
ties together, never row/line counts (import never changes row count --
see the module's own comment on this).

    python tests/test_entity_inference.py

Real documents only, gitignored (real client pricing) -- SKIPs, same
convention as test_avails_pdf_import.py, when they aren't present.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import avails_pdf_import as api  # noqa: E402

ANNAPOLIS = REPO / "Premion Media Plan_RFPID-253813_SR&B Advertising_Annapolis Cars_1-23-2026--ver0.pdf"
PLAZA = REPO / "Premion Media Plan_RFPID-266583_TBC, Inc - Trahan, Burden & Charles_Plaza Motors Group_8-19-2026--ver0.pdf"
WILMINGTON = REPO / "Premion Media Plan_RFPID-253956_Direct - No Agency_Wilmington University_1-27-2026--ver0.pdf"
LAWN_LEISURE = REPO / "Premion Media Plan_RFPID-265521_Direct - No Agency_Lawn & Leisure_7-28-2026--ver0.pdf"
ALL_FILES = [ANNAPOLIS, PLAZA, WILMINGTON, LAWN_LEISURE]

failures = []
skipped = False


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    global skipped
    missing = [p.name for p in ALL_FILES if not p.exists()]
    if missing:
        print(f"SKIP -- {len(missing)}/{len(ALL_FILES)} real avails PDFs not present "
              f"(gitignored fixtures): {missing}")
        skipped = True
        return 0

    print("Annapolis Cars (RFPID-253813) -- R1: same audience, same radius origin, "
          "differing radius -> 4 entities (one per make, each spanning its 5mi/10mi pair)")
    doc = api.parse_avails_pdf(str(ANNAPOLIS))
    keys, notes = api.infer_entities(doc)
    check("8 groups in, 8 keys out", len(keys) == 8, keys)
    check("exactly 4 distinct entities", len(set(keys)) == 4, keys)
    check("every entity has exactly 2 members (one radius pair each)",
          all(keys.count(k) == 2 for k in set(keys)), keys)
    check("each pair really is the SAME audience term, two different radii -- "
          "never two different audiences accidentally paired",
          all(doc.groups[i].audience_text == doc.groups[j].audience_text
              for k in set(keys)
              for i, j in [[idx for idx, kk in enumerate(keys) if kk == k]]),
          [(doc.groups[i].audience_text, doc.groups[j].audience_text)
           for k in set(keys)
           for i, j in [[idx for idx, kk in enumerate(keys) if kk == k]]])
    check("no ambiguous-geography note -- every pair resolved by R1, nothing left over",
          not notes, notes)

    print("\nPlaza Motors Group (RFPID-266583) -- R2: same geography, one DEMO Age "
          "bracket narrowing another over the SAME range -> 1 entity")
    doc = api.parse_avails_pdf(str(PLAZA))
    keys, notes = api.infer_entities(doc)
    check("2 groups in, 2 keys out, both the SAME key (1 entity)",
          len(keys) == 2 and keys[0] == keys[1] and keys[0] is not None, keys)
    check("no ambiguous-geography note", not notes, notes)

    print("\nWilmington University (RFPID-253956) -- three DIFFERENT undergrad audiences "
          "sharing one geography have no shared string tying them together -- correctly "
          "left ungrouped, flagged ONCE for the whole document, never once per row")
    doc = api.parse_avails_pdf(str(WILMINGTON))
    keys, notes = api.infer_entities(doc)
    check("5 groups in, 5 keys out, ALL None -- nothing invented", keys == [None] * 5, keys)
    check("exactly ONE aggregate note for the document, not five", len(notes) == 1, notes)

    print("\nLawn & Leisure (RFPID-265521) -- a single-group document has nothing to "
          "compare against; a guaranteed structural no-op, not a special case")
    doc = api.parse_avails_pdf(str(LAWN_LEISURE))
    keys, notes = api.infer_entities(doc)
    check("1 group in, 1 key out, None", keys == [None], keys)
    check("no note", not notes, notes)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("infer_entities correctly ties Annapolis's radius pairs and Plaza's DEMO Age "
          "bracket pair together, correctly leaves Wilmington's three undergrad audiences "
          "(no shared string) and Lawn & Leisure's lone group untouched, and never emits "
          "more than one aggregate note per document.")
    return 0


if __name__ == "__main__":
    code = main()
    if skipped:
        sys.exit(0)
    sys.exit(code)
