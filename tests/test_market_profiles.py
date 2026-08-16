"""Invariants for the market profile row set. Offline, no framework, no DB.

    python tests/test_market_profiles.py

What's asserted here is deliberately the stuff that fails SILENTLY: a key
collision merges two markets, an unclaimed slide means a profile nobody can
reach, and a market that quietly loses `has_profile_slide` takes its warning
out of the picker with it. None of those raise on their own.

The counts are pinned against the source deck and the Nielsen list rather
than against whatever build_rows() currently returns -- a check that
recomputes its own expectation proves only that the code agrees with itself.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import market_profiles as mp  # noqa: E402

# The source deck holds 206 slides: 205 markets + 1 national roll-up.
SLIDES_IN_DECK = 206
# Nielsen defines 210 DMAs; every one is selectable, plus the roll-up.
CANONICAL_DMA_COUNT = 210

# The five Premion never authored a profile slide for. Named, not counted:
# if this set changes, someone needs to look at why rather than watch a
# number move.
NO_SLIDE = {"Honolulu", "Palm Springs", "Anchorage", "Fairbanks", "Juneau"}

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def main():
    rows = mp.build_rows()
    index = mp.load_index()

    print("row set")
    check("one row per Nielsen DMA, plus the national roll-up",
          len(rows) == CANONICAL_DMA_COUNT + 1, f"got {len(rows)}")
    check("canonical list is the expected size",
          len(mp.CANONICAL_DMAS) == CANONICAL_DMA_COUNT,
          f"got {len(mp.CANONICAL_DMAS)}")
    check("no duplicate DMA names in the canonical list",
          len(set(mp.CANONICAL_DMAS)) == len(mp.CANONICAL_DMAS))

    print("\nkeys")
    keys = [r["key"] for r in rows]
    dupes = sorted({k for k in keys if keys.count(k) > 1})
    check("every key is unique", not dupes, f"collisions: {dupes}")
    check("no key is empty", all(r["key"] for r in rows))
    check("the roll-up keeps its documented key",
          any(r["key"] == mp.NATIONAL_KEY for r in rows))
    # These four are the pairs a naive slug would merge.
    for a, b in [("Albany, GA", "Albany-Schenectady-Troy"),
                 ("Portland, OR", "Portland-Auburn")]:
        check(f"{a!r} and {b!r} keep distinct keys",
              mp.slugify(a) != mp.slugify(b))

    print("\nslide coverage")
    claimed = [r["slide_number"] for r in rows if r["has_profile_slide"]]
    check("every slide in the deck is claimed by exactly one row",
          sorted(claimed) == sorted(e["slide"] for e in index),
          f"claimed {len(claimed)} of {len(index)}")
    check("the index really holds the whole deck",
          len(index) == SLIDES_IN_DECK, f"got {len(index)}")
    check("a row with a slide has a slide number, and one without has none",
          all((r["slide_number"] is not None) == r["has_profile_slide"]
              for r in rows))

    print("\nmarkets with no profile slide")
    missing = {r["dma"] for r in mp.markets_without_a_slide(rows)}
    check("exactly the five known gaps, by name", missing == NO_SLIDE,
          f"got {sorted(missing)}")
    check("they are still selectable rows, not dropped",
          all(any(r["dma"] == name for r in rows) for name in NO_SLIDE))
    check("they carry a label a rep can read",
          all(r["label"] for r in mp.markets_without_a_slide(rows)))

    print("\nthe national roll-up")
    national = [r for r in rows if r["kind"] == "national"]
    check("there is exactly one", len(national) == 1)
    check("it has no DMA", national and national[0]["dma"] is None)
    check("it sorts last", national and national[0]["rank"] == len(rows))

    print("\nstats column")
    check("build_rows does not invent a stats value",
          all("stats" not in r for r in rows),
          "the column is deliberately unpopulated; nothing may seed it")

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print(f"all {len(rows)} rows OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
