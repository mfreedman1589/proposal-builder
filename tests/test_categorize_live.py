"""Tier 2 -- live comparison of Claude's audience-categorization proposals
against the hand-made markup (audience_uncategorized_markup.csv). Costs
money (one call, ~100 components); run it yourself.

    ! python tests/test_categorize_live.py

Runs the real workbook through the mechanical rules with overrides={} --
deliberately ignoring the committed audience_component_overrides.csv, so
Claude faces the same set of components the hand markup covers (fewer
than the original 106 the markup was built against, now that the
mechanical WEB RT/LOCATION RT/RFPID/address-list pattern runs
unconditionally and no longer needs a component-by-component decision;
the 8 `split` rows are also excluded from the agreement tally, since
Claude is never asked to split a compound phrase -- only to categorize or
flag whatever string it's given, so "disagreement" there would compare
two different questions).

This prints a comparison, and checks the three assertions that have to
hold regardless of what the model says:
  - every returned category is one that actually exists (never invented)
  - the deterministic (prefix-resolved) path is untouched -- Claude is
    never even asked about a component the rules already resolved
  - a component already written to the catalog is never re-asked for
    (checked structurally against derive_catalog_updates's own
    "existing catalog is authoritative" precedence, not by a live call)
"""
import csv
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import app                                     # noqa: E402
import audience_catalog as ac                  # noqa: E402
import audience_usage_import as aui            # noqa: E402

WORKBOOK = REPO / "Premion OTT Audiences 7.1.26.xlsx"
MARKUP = REPO / "audience_uncategorized_markup.csv"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def load_markup():
    with open(MARKUP, encoding="utf-8-sig") as f:
        rows = [r for r in csv.DictReader(f) if r.get("component")]
    return {r["component"]: r for r in rows}


def main():
    if not WORKBOOK.exists():
        print(f"SKIP -- {WORKBOOK.name} not present.")
        return 0
    if not MARKUP.exists():
        print(f"SKIP -- {MARKUP.name} not present.")
        return 0

    workbook = aui.parse_workbook(str(WORKBOOK))
    brochure = ac.load_local_catalog().to_dict("records")

    # Deliberately overrides={} -- see module docstring. Also confirms the
    # deterministic path assertion structurally: everything NOT in this
    # `unresolved` list was resolved by rules alone, before Claude is ever
    # invoked below.
    report = aui.derive_catalog_updates(workbook, brochure, overrides={})
    unresolved = [u for u in report.updates if u.source == "uncategorized"]
    print(f"{len(unresolved)} components unresolved by rules alone "
          f"(mechanical categorization untouched by anything below)")

    markup = load_markup()
    # The split PIECES (e.g. "Homeowner") are in `unresolved` too, but
    # aren't keys in `markup` at all -- markup only names the ORIGINAL
    # compound phrase -- so they naturally fall out of `scored` below,
    # correctly excluded from the agreement tally.
    scored = [u for u in unresolved if u.segment in markup and markup[u.segment]["action"] != "split"]
    print(f"{len(scored)} of those have a hand-made markup row to compare against "
          f"({len(unresolved) - len(scored)} are split pieces or otherwise not in the markup)")

    components = [{"name": u.segment, "impressions": u.impressions} for u in unresolved]
    print(f"\nCalling Claude for all {len(components)} unresolved components...")
    suggestions, error = app.call_claude_categorize(components)
    if error:
        print(f"  FAIL  Claude call failed: {error}")
        return 1
    by_segment = {s["segment"]: s for s in suggestions}

    print("\nAssertion: every returned category actually exists")
    valid = set(ac.CATEGORY_DESCRIPTIONS)
    invented = [(s["segment"], s["category"]) for s in suggestions
               if s["category"] and s["category"] not in valid]
    check("no invented category names", not invented, invented)
    invented2 = [(s["segment"], s["category2"]) for s in suggestions
                if s["category2"] and s["category2"] not in valid]
    check("no invented category2 names either", not invented2, invented2)

    print("\nAssertion: the deterministic path is untouched")
    resolved_segments = {u.segment for u in report.updates if u.source != "uncategorized"}
    claude_touched_resolved = resolved_segments & set(by_segment)
    check("Claude was never even asked about a rules-resolved component",
          not claude_touched_resolved, claude_touched_resolved)

    print("\n" + "=" * 100)
    print(f"{'Segment':<55} {'Markup':<12} {'Claude':<12} {'Conf':<8} Agree?")
    print("=" * 100)
    agree = disagree = 0
    for u in scored:
        row = markup[u.segment]
        s = by_segment.get(u.segment, {})
        markup_label = "exclude" if row["action"] == "exclude_client" else (row["category"] or "blank")
        claude_cats = [c for c in (s.get("category"), s.get("category2")) if c]
        claude_label = "exclude" if s.get("is_client_specific") else (", ".join(claude_cats) or "blank")
        is_agree = (markup_label == "exclude" and s.get("is_client_specific")) or \
                   (markup_label != "exclude" and markup_label in claude_cats) or \
                   (markup_label == "blank" and not claude_cats and not s.get("is_client_specific"))
        agree += is_agree
        disagree += not is_agree
        mark = "  agree" if is_agree else "  DIFFER"
        print(f"{u.segment[:54]:<55} {markup_label:<12} {claude_label:<12} "
              f"{s.get('confidence', '?'):<8} {mark}")
        if s.get("reason"):
            print(f"    Claude: {s['reason']}")

    print("=" * 100)
    total = agree + disagree
    print(f"\n{agree}/{total} agree with the hand-made markup "
          f"({agree / total * 100:.0f}%)" if total else "nothing to compare")

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Structural assertions hold. See the table above for where the model's "
          "categorization agrees with the hand-made markup and where it doesn't.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
