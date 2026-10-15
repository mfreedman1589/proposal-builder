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
import re
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


# The state of each DMA's first-named (principal) city, by Nielsen DMA code --
# what turns "Washington-Hagerstown" into the "Washington, DC" a client knows
# the market by. Checked against each market's own county states by
# tests/test_avails_plan_basis.py (the principal city's state must be one the
# market actually covers).
_PRINCIPAL_STATE = dict(
    (int(code), state) for code, state in zip(*[iter("""
500 ME 501 NY 502 NY 503 GA 504 PA 505 MI 506 MA 507 GA 508 PA 509 IN 510 OH 511 DC 512 MD 513 MI
514 NY 515 OH 516 PA 517 NC 518 NC 519 SC 520 GA 521 RI 522 GA 523 VT 524 GA 525 GA 526 NY 527 IN
528 FL 529 KY 530 FL 531 TN 532 NY 533 CT 534 FL 535 OH 536 OH 537 ME 538 NY 539 FL 540 MI 541 KY
542 OH 543 MA 544 VA 545 NC 546 SC 547 OH 548 FL 549 NY 550 NC 551 MI 552 ME 553 MI 554 WV 555 NY
556 VA 557 TN 558 OH 559 WV 560 NC 561 FL 563 MI 564 WV 565 NY 566 PA 567 SC 569 VA 570 SC 571 FL
573 VA 574 PA 575 TN 576 MD 577 PA 581 IN 582 IN 583 MI 584 VA 588 IN 592 FL 596 OH 597 WV 598 WV
600 TX 602 IL 603 MO 604 MO 605 KS 606 AL 609 MO 610 IL 611 MN 612 LA 613 MN 616 MO 617 WI 618 TX
619 MO 622 LA 623 TX 624 IA 625 TX 626 TX 627 TX 628 LA 630 AL 631 IA 632 KY 633 TX 634 TX 635 TX
636 TX 637 IA 638 MO 639 TN 640 TN 641 TX 642 LA 643 LA 644 LA 647 MS 648 IL 649 IN 650 OK 651 TX
652 NE 656 FL 657 TX 658 WI 659 TN 661 TX 662 TX 669 WI 670 AR 671 OK 673 MS 675 IL 676 MN 678 KS
679 IA 682 IA 686 AL 687 ND 691 AL 692 TX 693 AR 698 AL 702 WI 705 WI 709 TX 710 MS 711 MS 716 LA
717 IL 718 MS 722 NE 724 ND 725 SD 734 AR 736 KY 737 MN 740 NE 743 AK 744 HI 745 AK 746 MS 747 AK
749 TX 751 CO 752 CO 753 AZ 754 MT 755 MT 756 MT 757 ID 758 ID 759 WY 760 ID 762 MT 764 SD 765 TX
766 MT 767 WY 770 UT 771 AZ 773 CO 789 AZ 790 NM 798 MT 800 CA 801 OR 802 CA 803 CA 804 CA 807 CA
810 WA 811 NV 813 OR 819 WA 820 OR 821 OR 825 CA 828 CA 839 NV 855 CA 862 CA 866 CA 868 CA 881 WA
""".split())] * 2))
_STATE_SUFFIX_RE = re.compile(r",\s*[A-Z]{2}$")
_PAREN_RE = re.compile(r"\s*\([^)]*\)")


def client_market_name(market_key):
    """The name a client knows a DMA by: the principal city and its state,
    with Nielsen's secondary cities and any parenthetical dropped --
    "Washington-Hagerstown" -> "Washington, DC", "Harrisburg-Lancaster-
    Lebanon-York" -> "Harrisburg, PA". None for an unknown key."""
    entry = load()["markets"].get(market_key)
    if not entry:
        return None
    name = _PAREN_RE.sub("", entry["name"]).strip()
    if _STATE_SUFFIX_RE.search(name):
        return name
    head = name if name.lower().startswith("tri-cities") else name.split("-")[0].strip()
    state = _PRINCIPAL_STATE.get(entry.get("dma_code"))
    return f"{head}, {state}" if state else head


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
