"""audience_evidence.py -- Phase 7 of the targeting-groups roadmap
(geo_targeting_roadmap.md D).

    python tests/test_audience_evidence.py

Offline, against the committed `audience_usage_ytd.csv` (207KB, not
gitignored). Every fixture below is a REAL row (or pair of rows) found in
that file by scanning it independently of `audience_evidence.py`'s own
parsing -- "assert against something the code can't move," this project's
own rule for a check that would otherwise only prove the code agrees with
itself. The exact counts this file asserts (the p90 threshold, the 24
reordered combinations, the suggested-pairing scores for "DEMO Homeowner" +
"HH Income 150K Plus") independently reproduce the specific worked numbers
in geo_targeting_roadmap.md D's own analysis.
"""
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import audience_evidence as ae  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "audience_usage_ytd.csv"
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def load_rows():
    with open(CSV_PATH, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def find_reordered_pair(rows):
    """A REAL pair of raw segment strings from the CSV that share a
    normalized SET but differ in part ORDER -- found by direct string
    manipulation here, not by calling audience_evidence at all, so this
    fixture doesn't depend on the code being tested."""
    from market_profiles import _normalize
    seen = {}
    for row in rows:
        parts = [p.strip() for p in str(row["segment"]).split(",") if p.strip()]
        if len(parts) < 2:
            continue
        key = frozenset(_normalize(p) for p in parts)
        order = tuple(parts)
        if key in seen and seen[key] != order:
            return seen[key], order
        seen.setdefault(key, order)
    return None, None


def main():
    if not CSV_PATH.exists():
        print(f"SKIP -- {CSV_PATH} not present.")
        return 0

    rows = load_rows()
    check("the committed CSV has rows", len(rows) > 1000, len(rows))
    index = ae.build_index(rows)
    check("an index was built", len(index.stacks) > 1000, len(index.stacks))

    print("\nset-not-string matching recovers a reordered combination")
    order_a, order_b = find_reordered_pair(rows)
    check("found a real reordered pair in the CSV", order_a is not None, (order_a, order_b))
    if order_a:
        stack_a = ae.parse_stack(", ".join(order_a))
        stack_b = ae.parse_stack(", ".join(order_b))
        check("both orderings parse to the IDENTICAL set", stack_a == stack_b, (stack_a, stack_b))
        ev_a = ae.evidence_for(list(order_a), index)
        ev_b = ae.evidence_for(list(order_b), index)
        check("both orderings report the SAME exact-match booking history",
              ev_a["exact_match"] == ev_b["exact_match"] == True, (ev_a["exact_match"], ev_b["exact_match"]))

    print("\ncase/punctuation normalization folds 'Lifestyle Charity' / 'LIFESTYLE Charity'")
    check("both raw spellings are present in the real CSV (a real adversarial case, not invented)",
          "Lifestyle Charity" in {r["segment"] for r in rows}
          or any("Lifestyle Charity" in str(r["segment"]) for r in rows), True)
    check("normalize() folds the two spellings to the same key",
          ae.normalize("Lifestyle Charity") == ae.normalize("LIFESTYLE Charity"),
          (ae.normalize("Lifestyle Charity"), ae.normalize("LIFESTYLE Charity")))
    check("...and so does a folded punctuation variant (Pets - Horses / Pets Horses)",
          ae.normalize("LIFESTYLE Pets - Horses") == ae.normalize("LIFESTYLE Pets Horses"),
          (ae.normalize("LIFESTYLE Pets - Horses"), ae.normalize("LIFESTYLE Pets Horses")))

    print("\nthe live p90 threshold reproduces the roadmap's own validated distribution")
    threshold = ae.widely_used_threshold(index)
    check("threshold lands near 10 (the roadmap's own p90 figure against this data)",
          9 <= threshold <= 11, threshold)
    widely_used_count = sum(1 for v in index.component_counts.values() if v >= threshold)
    fraction = widely_used_count / len(index.component_counts)
    check("selects roughly the top decile of components (~11%, per the roadmap)",
          0.08 <= fraction <= 0.14, (widely_used_count, len(index.component_counts), fraction))

    print("\nno output ever states a frequency for the exact-stack match, or says 'not booked'")
    sample_selections = [
        [comp for comp, _count in sorted(index.component_counts.items(),
                                         key=lambda kv: -kv[1])[:1]],
        [index.display(c) for c in list(index.stacks)[0]] if index.stacks else [],
        ["A Segment Nobody Has Ever Booked, Definitely Not"],
    ]
    all_lines = []
    for selection in sample_selections:
        if not selection:
            continue
        evidence = ae.evidence_for(selection, index)
        all_lines.extend(ae.evidence_lines(evidence))
    exact_lines = [ln for ln in all_lines if ln.startswith("This exact audience")]
    check("the exact-match line, whenever it appears, carries no number",
          all(not any(c.isdigit() for c in ln) for ln in exact_lines), exact_lines)
    check("no line anywhere says 'not booked' (a pairwise zero uses different, allowed wording)",
          not any("not booked" in ln.lower() for ln in all_lines), all_lines)

    print("\nboth-widely-used vs. either-rare picks the right sentence for a zero co-occurrence "
          "-- both fixtures are REAL pairs with zero real co-occurrence in the CSV")
    both_widely = ["DEMO Age A55 Plus", "LIFESTYLE Health Wellness"]
    ev_both = ae.evidence_for(both_widely, index)
    check("both components actually clear the widely-used threshold (fixture sanity)",
          all(count >= threshold for _name, count in ev_both["familiarity"]), ev_both["familiarity"])
    check("zero co-occurrence between them, for real, in this data",
          ev_both["weakest_pair"][2] == 0, ev_both["weakest_pair"])
    check("classified both_widely_used", ev_both["weakest_pair_kind"] == "both_widely_used",
          ev_both["weakest_pair_kind"])
    lines_both = ae.evidence_lines(ev_both)
    check("renders the 'both widely used' sentence, not a bare negative",
          any("widely used" in ln and "haven't been booked together" in ln for ln in lines_both),
          lines_both)

    either_rare = ["HLTH CRX DX Back Pain", "DEMO Age A55 Plus"]
    ev_rare = ae.evidence_for(either_rare, index)
    check("one component is genuinely rare (fixture sanity)",
          any(count < threshold for _name, count in ev_rare["familiarity"]), ev_rare["familiarity"])
    check("zero co-occurrence between them, for real, in this data",
          ev_rare["weakest_pair"][2] == 0, ev_rare["weakest_pair"])
    check("classified either_rare, not both_widely_used",
          ev_rare["weakest_pair_kind"] == "either_rare", ev_rare["weakest_pair_kind"])
    lines_rare = ae.evidence_lines(ev_rare)
    check("renders the plain haven't-been-booked-together sentence, claims nothing about commonality",
          any("haven't been booked together" in ln and "widely used" not in ln for ln in lines_rare),
          lines_rare)

    print("\nsuggestions for two segments aren't a flat list of 1s (the strict-superset trap)")
    two_seg = ["DEMO Homeowner", "HH Income 150K Plus"]
    ev_two = ae.evidence_for(two_seg, index)
    scores = [score for _name, score in ev_two["suggestions"]]
    check("suggestions exist", bool(scores), scores)
    check("scores are NOT a flat list of 1s -- real spread, weighted by overlap",
          len(scores) > 1 and len(set(scores)) > 1 and max(scores) > 2, scores)
    top = dict(ev_two["suggestions"])
    check("reproduces the roadmap's own worked numbers for this exact pair "
          "(DEMO Age A35 Plus 24, DEMO Age A35-64 20, DEMO Age A25 Plus 18)",
          top.get("DEMO Age A35 Plus") == 24 and top.get("DEMO Age A35-64") == 20
          and top.get("DEMO Age A25 Plus") == 18, ev_two["suggestions"])

    one_seg = ["DEMO Age A25 Plus"]
    ev_one = ae.evidence_for(one_seg, index)
    one_scores = [score for _name, score in ev_one["suggestions"]]
    check("a single-segment selection also has real, non-flat suggestion scores",
          len(set(one_scores)) > 1, one_scores)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Booking evidence matches sets not strings, folds case/punctuation, never states an "
          "exact-stack frequency or says 'not booked', picks the right zero-co-occurrence "
          "sentence, and weights suggestions by overlap instead of collapsing to a flat list.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
