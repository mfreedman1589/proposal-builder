"""The county -> market table, held against things it cannot move.

Offline, no framework, no network: `python tests/test_market_lookup.py`.

The table is a compilation of public sources of uncertain vintage (see
build_market_lookup.py), so what is asserted here is not "these 3,118 counties
are right" -- nothing available could establish that. It is the set of
properties that must hold for the table to be USABLE, and whose failure would
be silent:

  * every market it names is one the app already knows, and every one of the
    210 canonical markets is reachable except the one documented gap. A name
    this table invents would resolve to no market profile and no deck slide.
  * every county the resolver can return has an assignment, so a zip list
    never comes back mysteriously short.
  * independent cities are their own place. Baltimore city is not Baltimore
    County and Franklin city is not Franklin County -- and the Franklin pair
    is in two DIFFERENT markets, which is the case that proves a table
    distinguishes them rather than merely carrying both names.
  * real zips resolve to the markets a person would name for them, including
    the two this app's decks are actually built for (DC and Harrisburg).
  * the provenance block is present and complete, because a table whose origin
    can't be answered later is one nobody can responsibly replace.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import geo_resolver
import market_lookup
import market_profiles

FAILED = []


def check(label, condition, detail=None):
    if condition:
        print("  PASS  %s" % label)
    else:
        FAILED.append(label)
        print("  FAIL  %s%s" % (label, "" if detail is None else "  -- %r" % (detail,)))


def main():
    if not market_lookup.available():
        print("SKIP  market_lookup.json.gz is not built "
              "(`python build_market_lookup.py`)")
        return 0

    payload = market_lookup.load()
    by_county = payload["by_county"]
    markets = payload["markets"]
    canonical = list(market_profiles.CANONICAL_DMAS)
    canonical_keys = {market_profiles.slugify(n) for n in canonical}

    print("names: every market is one the app already knows")
    check("every market key is a canonical DMA",
          set(markets) <= canonical_keys,
          sorted(set(markets) - canonical_keys)[:5])
    check("all 210 canonical markets are named",
          len(markets) == 210, len(markets))
    check("every assigned county names a market the table defines",
          set(by_county.values()) <= set(markets),
          sorted(set(by_county.values()) - set(markets))[:5])
    check("every DMA code is distinct",
          len({m["dma_code"] for m in markets.values()}) == len(markets))

    print("\ncoverage")
    reachable = set(by_county.values())
    unreachable = sorted(set(markets) - reachable)
    # Palm Springs is carved out of Riverside County CA, which the source gives
    # whole to Los Angeles. Asserted by name so that a future table which DOES
    # cover it fails this and gets the note removed, rather than quietly
    # keeping a stale caveat.
    check("exactly one market has no counties, and it is Palm Springs",
          unreachable == ["palm_springs"], unreachable)
    check("no county is assigned to an empty string or None",
          all(v for v in by_county.values()))
    check("the table covers over 3,000 counties",
          len(by_county) > 3000, len(by_county))

    print("\nindependent cities are their own place")
    pairs = [("24005", "24510", "Baltimore"), ("51059", "51600", "Fairfax"),
             ("51067", "51620", "Franklin"), ("51159", "51760", "Richmond"),
             ("51161", "51770", "Roanoke"), ("29189", "29510", "St. Louis")]
    for county_fips, city_fips, label in pairs:
        check("%s: both the county and the city are assigned" % label,
              by_county.get(county_fips) and by_county.get(city_fips),
              (by_county.get(county_fips), by_county.get(city_fips)))
    check("Franklin County VA and Franklin city VA are in DIFFERENT markets",
          by_county.get("51067") != by_county.get("51620"),
          (by_county.get("51067"), by_county.get("51620")))
    check("...specifically Roanoke-Lynchburg and Norfolk",
          by_county.get("51067") == "roanoke_lynchburg"
          and by_county.get("51620") == "norfolk_portsmouth_newport_news",
          (by_county.get("51067"), by_county.get("51620")))
    check("Baltimore city and Baltimore County are both Baltimore",
          by_county.get("24005") == by_county.get("24510") == "baltimore")

    print("\nresolution end to end, through the real resolver")
    market_lookup.install()
    check("the resolver reports a lookup is available",
          geo_resolver.market_lookup_available())
    known = {
        "20005": "washington_hagerstown",     # Washington DC
        "22209": "washington_hagerstown",     # Arlington VA
        "21201": "baltimore",                 # Baltimore MD
        "17101": "harrisburg_lancaster_lebanon_york",   # Harrisburg PA
        "10001": "new_york",                  # Manhattan
        "60601": "chicago",                   # Chicago
        "94102": "san_francisco_oakland_san_jose",
        "06103": "hartford_new_haven",        # Hartford CT -- legacy county path
    }
    for zip_code, expected in known.items():
        res = geo_resolver.zips_to_markets([zip_code])
        check("%s resolves to %s" % (zip_code, expected),
              list(res.resolved) == [expected], (res.resolved, res.unresolved))

    res = geo_resolver.zips_to_markets(sorted(known))
    check("a mixed zip list resolves to every one of those markets",
          set(res.resolved) == set(known.values()), sorted(res.resolved))
    check("...and nothing is left unresolved", not res.unresolved, res.unresolved)

    res = geo_resolver.market_to_zips("washington_hagerstown")
    check("market_to_zips walks back to the zips",
          "20005" in res.resolved and len(res.resolved) > 100, len(res.resolved))

    print("\nprovenance is recorded")
    prov = payload.get("provenance", {})
    for field in ("retrieved", "definitions", "sources", "assignment_vintage"):
        check("provenance carries %s" % field, bool(prov.get(field)))
    check("every source names a url and what it was used for",
          prov.get("sources") and all(s.get("url") and s.get("used_for")
                                      for s in prov["sources"]),
          prov.get("sources"))
    check("the provenance says outright that the definitions are Nielsen's",
          "Nielsen" in prov.get("definitions", ""))

    print()
    if FAILED:
        print("%d FAILED: %s" % (len(FAILED), "; ".join(FAILED)))
        return 1
    print("The market lookup is complete, canonical and resolves real zips.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
