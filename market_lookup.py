"""The county -> market (DMA) table, and installing it into the geo resolver.

    python market_lookup.py            # summary, provenance and a spot check

`market_lookup.json.gz` is built by `build_market_lookup.py`, which downloads
three public sources, cross-validates them and refuses to write anything that
doesn't pass. This module is only the loading half: it reads that file and
hands it to `geo_resolver.register_market_lookup`.

Installing is EXPLICIT, and stays that way. `geo_resolver` ships with no lookup
on purpose -- county->DMA is Nielsen's intellectual property, the table in this
repo is a public secondary compilation, and a consumer should be choosing to
use it rather than finding it switched on by an import. Call `install()` from
whatever needs markets resolved; `zips_to_markets` reports itself unavailable
until something does, which is the honest state.

What the table cannot do, both recorded in its own provenance block:

  * **Palm Springs** is a sub-county DMA carved out of Riverside County CA and
    the assignment source has no rows for it at all, so no zip will ever
    resolve to it. 209 of the 210 markets are reachable.
  * **Connecticut's nine planning regions** carry no assignment, because the
    crosswalk's zip file still uses the eight legacy counties -- which ARE
    assigned, so Connecticut zips resolve correctly regardless. Only a caller
    that starts from a planning-region FIPS would notice.
"""
import gzip
import json
import sys
from functools import lru_cache
from pathlib import Path

LOOKUP_PATH = Path(__file__).resolve().parent / "market_lookup.json.gz"


class MarketLookupUnavailable(RuntimeError):
    """The built table is missing. Build it with build_market_lookup.py."""


@lru_cache(maxsize=1)
def load():
    """The whole payload: by_county, markets, provenance, validation."""
    if not LOOKUP_PATH.exists():
        raise MarketLookupUnavailable(
            f"{LOOKUP_PATH.name} is missing. Build it with "
            f"`python build_market_lookup.py`.")
    with gzip.open(LOOKUP_PATH, "rb") as fh:
        return json.loads(fh.read().decode("utf-8"))


def available():
    """Is the table on disk? Never raises -- a caller may want to degrade."""
    return LOOKUP_PATH.exists()


def install():
    """Register the table with geo_resolver. Returns the lookup, or None.

    Imported here rather than at module scope so geo_resolver keeps no
    dependency on this module in either direction.
    """
    import geo_resolver

    if not available():
        return None
    payload = load()
    lookup = geo_resolver.TableMarketLookup(
        by_county=payload["by_county"],
        name="public county->DMA compilation (build_market_lookup.py)")
    geo_resolver.register_market_lookup(lookup)
    return lookup


def market_name(market_key):
    """The canonical DMA name for a market key, or None."""
    entry = load()["markets"].get(market_key)
    return entry["name"] if entry else None


def states_for_market(market_key):
    """The set of 2-letter state abbreviations a market (DMA)'s own
    counties span -- every state it touches, not just the ones a
    particular caller's zips happen to land in. The same technique that
    found West Virginia missing from the ZCTA boundary set (2026-09-12):
    `washington_hagerstown`'s own county list resolves to {DC, MD, PA, VA,
    WV}, which is what should be checked against what's actually built,
    rather than waiting for a WV zip to show up in some future export.

    Needs `geo_resolver`'s crosswalk for the county->state step (the same
    one `build_zcta_boundaries.py`'s own `_zip_to_state` reaches into);
    returns an empty set rather than raising if it isn't built, or if
    `market_key` isn't a real market -- this is a diagnostic, not a gate.
    """
    import geo_resolver

    try:
        counties = geo_resolver._data().get("counties") or {}
    except geo_resolver.GeoDataUnavailable:
        return set()
    payload = load()
    return {state for fips, key in payload["by_county"].items()
           if key == market_key
           for state in [(counties.get(fips) or {}).get("state")]
           if state}


@lru_cache(maxsize=1)
def _zip3_fallback_tables():
    """({zip3: market_key}, {zip3: state}) for every 3-digit zip prefix
    where EVERY real zip this crosswalk can resolve agrees -- unanimous,
    not majority, since a real DMA or state boundary occasionally does
    cut through one 3-digit block, and a confident-looking wrong guess
    there is worse than falling through to a coarser tier or dropping
    the row (2026-09-13, the Ashburn/20149 find: a real, deliverable
    PO-box-only zip has no ZCTA and so no county/market of its own, but
    every OTHER real zip sharing its "201" prefix -- 45 of them --
    unanimously resolves to `washington_hagerstown`, so that market is
    almost certainly right for 20149 too). Built once from data already
    in this repo -- no new source, no network call.
    """
    import geo_resolver

    if not geo_resolver.market_lookup_available():
        return {}, {}
    data = geo_resolver._data()
    zip_counties = data.get("zip_counties") or {}
    counties = data.get("counties") or {}
    all_zips = list(zip_counties.keys())
    market_res = geo_resolver.zips_to_markets(all_zips)
    zip_market = {z: key for key, info in market_res.resolved.items() for z in info["zips"]}
    zip_state = {}
    for z, fips_list in zip_counties.items():
        state = (counties.get(fips_list[0]) or {}).get("state") if fips_list else None
        if state:
            zip_state[z] = state

    by_prefix_market, by_prefix_state = {}, {}
    for z in all_zips:
        prefix = z[:3]
        if z in zip_market:
            by_prefix_market.setdefault(prefix, set()).add(zip_market[z])
        if z in zip_state:
            by_prefix_state.setdefault(prefix, set()).add(zip_state[z])
    prefix_market = {p: next(iter(s)) for p, s in by_prefix_market.items() if len(s) == 1}
    prefix_state = {p: next(iter(s)) for p, s in by_prefix_state.items() if len(s) == 1}
    return prefix_market, prefix_state


def zip3_market_fallback(zip_code):
    """The market EVERY real, resolved zip sharing this zip's 3-digit
    prefix agrees on, or None if there's no such zip, they disagree, or
    the market lookup isn't installed. See `_zip3_fallback_tables`. A
    fallback for a zip with no county/market entry of its own -- never
    consulted for a zip that already resolved normally."""
    prefix_market, _prefix_state = _zip3_fallback_tables()
    return prefix_market.get(str(zip_code).strip().zfill(5)[:3])


def zip3_state_fallback(zip_code):
    """Same idea as `zip3_market_fallback`, one tier coarser: the state
    every real, resolved zip sharing this prefix agrees on."""
    _prefix_market, prefix_state = _zip3_fallback_tables()
    return prefix_state.get(str(zip_code).strip().zfill(5)[:3])


def provenance():
    return load()["provenance"]


def main():
    payload = load()
    prov = payload["provenance"]
    val = payload["validation"]
    print("market_lookup.json.gz")
    print("  %d counties -> %d markets" % (len(payload["by_county"]),
                                           len(payload["markets"])))
    print("  built %s by %s" % (prov["retrieved"], prov["built_by"]))
    print("\nprovenance")
    print("  " + prov["definitions"].replace(". ", ".\n  "))
    for src in prov["sources"]:
        print("\n  %s\n    %s\n    used for: %s" % (src["name"], src["repo"],
                                                    src["used_for"]))
    print("\n  vintage: " + prov["assignment_vintage"].replace(". ", ".\n  "))
    print("\nvalidation as built")
    print("  markets with counties: %s" % val["markets_with_counties"])
    print("  known gaps: %s" % (val.get("markets_known_gaps") or "none"))
    print("  counties assigned: %s of %s" % (val["counties_assigned"],
                                             val["counties_expected"]))
    print("  unassigned: %s" % (val["counties_unassigned"] or "none"))
    cc = val.get("crosscheck", {})
    if cc:
        print("  cross-check: %s of %s counties land in the paired market"
              % (cc.get("counties_in_paired_market"),
                 (cc.get("counties_in_paired_market") or 0) + (cc.get("counties_differing") or 0)))

    install()
    import geo_resolver
    res = geo_resolver.zips_to_markets(["20005", "21201", "17101"])
    print("\nspot check -- zips_to_markets(['20005', '21201', '17101'])")
    for key, entry in sorted(res.resolved.items()):
        print("  %-34s %s  (%d zip)" % (key, market_name(key), entry["zip_count"]))
    for note in res.notes:
        print("  note: " + note)
    return 0


if __name__ == "__main__":
    sys.exit(main())
