"""The geo resolver, against known-good real geography.

    python tests/test_geo_resolver.py

Offline apart from one clearly-marked network check, which SKIPS rather than
fails when there's no connection -- a test that only passes on a good wifi day
is a test that gets ignored.

Expectations are real places whose geography is a matter of public record:
Arlington VA is FIPS 51013, Kansas City straddles the Missouri-Kansas line,
Washington DC is one county-equivalent. None of it is derived by re-running
the code under test.

The market functions are exercised through a FAKE lookup. That is not a
compromise -- it is the point of the pluggable design: the plumbing has to be
testable without the licensed county-to-DMA data, because the licensing may
never allow that data into this repo.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import geo_resolver as geo  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def main():
    print("normalizing what a rep actually pastes")
    for raw, want in [("20005", "20005"), (" 20005 ", "20005"),
                      ("20005-1234", "20005"), (1720, "01720"),
                      ("1720", "01720"), ("", None), ("abcde", None),
                      ("123456", None), (None, None)]:
        got = geo.normalize_zip(raw)
        check(f"{raw!r} -> {want!r}", got == want, f"got {got!r}")
    check("a pasted list splits on commas, spaces and newlines",
          geo.parse_zip_list("20005, 20006\n20007 20008") ==
          ["20005", "20006", "20007", "20008"],
          geo.parse_zip_list("20005, 20006\n20007 20008"))

    print("\nzips -> counties")
    res = geo.zips_to_counties(["22209", "20005"])
    check("Arlington's zip resolves to Arlington County (FIPS 51013)",
          res.resolved.get("22209", [None])[0] == "51013",
          res.resolved.get("22209"))
    check("a DC zip resolves to the District (FIPS 11001)",
          res.resolved.get("20005", [None])[0] == "11001",
          res.resolved.get("20005"))
    check("nothing unresolved for two real zips", not res.unresolved, res.unresolved)
    check("county names come back readable",
          geo.county_name("51013") == "Arlington County, VA", geo.county_name("51013"))

    # A zip spanning several counties keeps all of them. 10,186 of 33,791 do,
    # so dropping the extras would discard a third of the geography.
    multi = [z for z, c in geo._data()["zip_counties"].items() if len(c) > 2]
    check("zips spanning several counties keep every one",
          bool(multi) and len(geo.zips_to_counties([multi[0]]).resolved[multi[0]]) > 2,
          f"sample {multi[:1]}")

    print("\nunresolvable input is reported, never dropped")
    res = geo.zips_to_counties(["20005", "99999", "banana", "00000"])
    check("the good zip still resolves", "20005" in res.resolved, res.resolved)
    check("the bad ones come back in unresolved",
          set(res.unresolved) == {"99999", "banana", "00000"}, res.unresolved)
    check("and a note explains why", bool(res.notes), res.notes)
    check("resolved + unresolved accounts for every input",
          len(res.resolved) + len(res.unresolved) == 4)

    print("\ncounties -> zips")
    res = geo.counties_to_zips(["51013"])
    check("Arlington County yields its zips", len(res.resolved.get("51013", [])) > 5,
          len(res.resolved.get("51013", [])))
    check("22209 is among them", "22209" in res.resolved.get("51013", []))
    res = geo.counties_to_zips(["51013", "99999"])
    check("an unknown FIPS is reported, not dropped",
          res.unresolved == ["99999"] and "51013" in res.resolved, res.unresolved)

    print("\ncounty names -> FIPS (the form the avails documents use)")
    res = geo.counties_to_fips("SOMERSET NJ;BUCKS PA;NEW CASTLE DE")
    check("the semicolon-separated NAME STATE form resolves",
          len(res.resolved) == 3 and not res.unresolved, res)
    check("Bucks PA is FIPS 42017", res.resolved.get("BUCKS PA") == "42017",
          res.resolved)
    check("a type suffix is optional -- these documents omit it",
          geo.counties_to_fips("Bucks County PA").resolved.get("Bucks County PA")
          == "42017")
    check("St / Saint spellings both resolve",
          geo.counties_to_fips("ST LOUIS COUNTY MO").resolved
          and geo.counties_to_fips("Saint Louis County MO").resolved)

    # The independent-city cases. "Baltimore MD" is two different places with
    # different zips; guessing the bigger one would put money in the wrong
    # half of a metro.
    res = geo.counties_to_fips("BALTIMORE MD")
    check("an ambiguous county name is NOT guessed",
          not res.resolved and res.unresolved == ["BALTIMORE MD"], res)
    check("...and both candidates are named",
          res.notes and "Baltimore County" in res.notes[0]
          and "Baltimore city" in res.notes[0], res.notes)
    check("spelling the suffix out disambiguates it",
          geo.counties_to_fips("Baltimore County MD").resolved
          != geo.counties_to_fips("Baltimore city MD").resolved)
    res = geo.counties_to_fips("BALTIMORE MD;BUCKS PA")
    check("one ambiguous name doesn't take the others down with it",
          res.resolved.get("BUCKS PA") == "42017", res)

    res = geo.counties_to_fips("NOWHERE ZZ")
    check("an unknown county is reported, not dropped",
          res.unresolved == ["NOWHERE ZZ"] and res.notes, res)

    print("\ncounties_to_zips takes names and codes, mixed")
    res = geo.counties_to_zips("BUCKS PA")
    check("a name resolves through to zips",
          len(res.resolved.get("42017", [])) > 10, res.resolved)
    res = geo.counties_to_zips(["51013", "BUCKS PA"])
    check("codes and names together", set(res.resolved) == {"51013", "42017"},
          sorted(res.resolved))

    print("\nradius (centroid-in-circle, matching freemaptools)")
    res = geo.radius_to_zips("20005", 5)
    check("a 5-mile radius on downtown DC returns a sensible number of zips",
          20 < len(res.resolved) < 250, len(res.resolved))
    check("the centre zip is included", "20005" in res.resolved)
    small = geo.radius_to_zips("20005", 1)
    check("a smaller radius returns fewer zips",
          len(small.resolved) < len(res.resolved),
          f"{len(small.resolved)} vs {len(res.resolved)}")

    # A radius crossing a state line -- Kansas City straddles MO/KS, so a
    # radius there must return zips in both states or the maths is wrong.
    kc = geo.radius_to_zips("64106", 15)     # downtown Kansas City, MO
    states = set()
    counties = geo.zips_to_counties(kc.resolved)
    for fips_list in counties.resolved.values():
        for fips in fips_list:
            entry = geo._data()["counties"].get(fips)
            if entry:
                states.add(entry["state"])
    check("a radius over Kansas City crosses the state line into KS",
          {"MO", "KS"} <= states, sorted(states))

    multi_centre = geo.radius_to_zips(["20005", "64106"], 10)
    check("several centres union rather than replace each other",
          len(multi_centre.resolved) >
          max(len(geo.radius_to_zips("20005", 10).resolved),
              len(geo.radius_to_zips("64106", 10).resolved)),
          len(multi_centre.resolved))

    bad = geo.radius_to_zips("not-a-place", 10, geocoder=lambda _a: None)
    check("an unlocatable centre is reported, not silently empty",
          bad.unresolved == ["not-a-place"] and bad.notes, bad)
    check("a non-numeric radius is reported",
          geo.radius_to_zips("20005", "ten").unresolved == ["ten"])
    check("a zero radius says so rather than returning everything",
          geo.radius_to_zips("20005", 0).resolved == []
          and geo.radius_to_zips("20005", 0).notes)

    print("\nmarkets: unavailable until a lookup is registered")
    geo.register_market_lookup(None)
    check("no lookup means not available", not geo.market_lookup_available())
    res = geo.zips_to_markets(["20005", "22209"])
    check("zips_to_markets resolves nothing without one", res.resolved == {})
    check("...and says why, rather than looking like 'covers no markets'",
          any("licensed" in n or "lookup" in n for n in res.notes), res.notes)
    check("the zips come back as unresolved so nothing is lost",
          set(res.unresolved) >= {"20005", "22209"}, res.unresolved)
    res = geo.market_to_zips("washington_hagerstown")
    check("market_to_zips is unavailable too, and says so",
          res.resolved == [] and res.notes, res)

    print("\nmarkets: the pluggable lookup, exercised with a fake table")
    # county-keyed, which is how DMAs are really defined
    geo.register_market_lookup(geo.TableMarketLookup(by_county={
        "11001": "washington_hagerstown",     # District of Columbia
        "51013": "washington_hagerstown",     # Arlington County, VA
        "29095": "kansas_city",               # Jackson County, MO
    }, name="fake"))
    check("a registered lookup reports available", geo.market_lookup_available())
    res = geo.zips_to_markets(["20005", "22209", "64106"])
    check("a zip list spanning two markets resolves to both",
          set(res.resolved) == {"washington_hagerstown", "kansas_city"},
          sorted(res.resolved))
    check("with a zip count per market",
          res.resolved["washington_hagerstown"]["zip_count"] == 2,
          res.resolved["washington_hagerstown"])
    check("and the zips themselves listed",
          "22209" in res.resolved["washington_hagerstown"]["zips"],
          res.resolved["washington_hagerstown"]["zips"])
    res = geo.market_to_zips("washington_hagerstown")
    check("market_to_zips walks back through the counties",
          "22209" in res.resolved and "20005" in res.resolved, len(res.resolved))
    res = geo.zips_to_markets(["20005", "99999"])
    check("an unresolvable zip stays unresolved even with a lookup",
          "99999" in res.unresolved, res.unresolved)

    # A zip whose counties sit in DIFFERENT markets -- the split case. Real
    # DMAs are county-aligned, but a ZCTA is not, so this genuinely happens.
    split_zip = next((z for z, c in geo._data()["zip_counties"].items()
                      if len(c) > 1), None)
    if split_zip:
        a, b = geo._data()["zip_counties"][split_zip][:2]
        geo.register_market_lookup(geo.TableMarketLookup(
            by_county={a: "market_a", b: "market_b"}, name="split"))
        res = geo.zips_to_markets([split_zip])
        check("a zip split across two markets counts once, in the larger",
              sum(m["zip_count"] for m in res.resolved.values()) == 1,
              res.resolved)
        check("...and the split is reported rather than hidden",
              any("span" in n for n in res.notes), res.notes)

    # A zip-keyed source (how a purchased table usually arrives) wins over
    # the county path, being more specific.
    geo.register_market_lookup(geo.TableMarketLookup(
        by_county={"11001": "from_county"}, by_zip={"20005": "from_zip"}, name="both"))
    res = geo.zips_to_markets(["20005"])
    check("a zip-keyed lookup takes precedence over the county one",
          list(res.resolved) == ["from_zip"], sorted(res.resolved))
    geo.register_market_lookup(None)

    print("\nlive Census Geocoder (network -- skips if unreachable)")
    point = geo.geocode_address("1100 Wilson Blvd, Arlington, VA")
    if point is None:
        print("  SKIP  geocoder unreachable")
    else:
        lat, lon = point
        check("an address geocodes to roughly the right place",
              38.5 < lat < 39.2 and -77.5 < lon < -76.8, point)
        res = geo.radius_to_zips("1100 Wilson Blvd, Arlington, VA", 3)
        check("and a radius around an address returns zips", len(res.resolved) > 5,
              len(res.resolved))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Zips, counties and radii resolve; markets wait on a lookup.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
