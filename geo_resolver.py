"""Zip / county / market resolution. Pure functions over one data file.

No UI, no Streamlit, no database. `geo_crosswalk.json.gz` (0.5 MiB, built by
build_geo_crosswalk.py from three public-domain Census files) is loaded once
and cached.

**Every function returns a Resolution, never a bare list.** `resolved` is what
was worked out, `unresolved` is every input that couldn't be, and `notes` is
plain-language explanation. Silently returning a shorter list than it was
given is the failure mode this shape exists to prevent: a zip list that
quietly loses six zips produces an avails request for the wrong footprint, and
nothing downstream can tell.

## The market lookup is pluggable, and empty by default

county -> DMA is Nielsen's intellectual property and has no clean public
source, so this module ships without one and says so rather than guessing.
`zips_to_markets` and `market_to_zips` work the moment a lookup is registered
and report themselves unavailable until then -- they never fail silently and
never invent a market.

A lookup can come from anywhere: a purchased ZIP-by-DMA table, a file supplied
by whoever holds the Nielsen entitlement internally, or one accumulated from
imported avails documents (which state zips and markets together, so every
import teaches the table a little more). `TableMarketLookup` covers all three
-- it takes plain dicts -- and anything with the same two methods will do.

Sources are keyed on COUNTY where possible, because that is how DMAs are
actually defined; a zip-keyed source is supported too, since purchased tables
usually come that way. When both are available the zip answer wins, being the
more specific.
"""
import gzip
import json
import math
import re
import urllib.parse
import urllib.request
from collections import namedtuple
from functools import lru_cache
from pathlib import Path

CROSSWALK_PATH = Path(__file__).resolve().parent / "geo_crosswalk.json.gz"

# The `locations` endpoint, not `geographies`. All this needs is a lat/long --
# the county comes from the crosswalk, not the geocoder -- and `geographies`
# rejects a request with no `vintage` parameter as a bare HTTP 400. Since
# geocode_address swallows failures by design, that 400 surfaced as "geocoder
# unreachable" and looked like bad wifi rather than a wrong URL.
CENSUS_GEOCODER = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress"
GEOCODER_TIMEOUT = 30

EARTH_RADIUS_MILES = 3958.7613

Resolution = namedtuple("Resolution", "resolved unresolved notes")


class GeoDataUnavailable(RuntimeError):
    """The crosswalk file is missing. Rebuild it with build_geo_crosswalk.py."""


@lru_cache(maxsize=1)
def _data():
    if not CROSSWALK_PATH.exists():
        raise GeoDataUnavailable(
            f"{CROSSWALK_PATH.name} is missing. Build it with "
            f"`python build_geo_crosswalk.py`.")
    with gzip.open(CROSSWALK_PATH, "rb") as f:
        return json.loads(f.read().decode("utf-8"))


def normalize_zip(value):
    """A 5-digit zip, or None.

    Accepts what a rep actually pastes: "20005", "20005-1234", 20005 as a
    number, and the leading-zero cases a spreadsheet mangles ("1720" is
    Massachusetts, not nothing). Anything else is None and gets reported
    unresolved rather than silently coerced.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.split("-", 1)[0].strip()
    if not text.isdigit() or len(text) > 5:
        return None
    return text.zfill(5)


def parse_zip_list(text):
    """Split a pasted zip list on whatever the rep separated it with."""
    if isinstance(text, (list, tuple, set)):
        raw = list(text)
    else:
        raw = str(text or "").replace("\n", ",").replace("\t", ",").replace(" ", ",").split(",")
    return [r for r in (str(x).strip() for x in raw) if r]


# ---------------------------------------------------------------------------
# The pluggable market lookup
# ---------------------------------------------------------------------------
class TableMarketLookup:
    """A market lookup backed by plain dicts.

    `by_county` maps a 5-digit county FIPS to a market key; `by_zip` maps a
    5-digit zip. Either may be empty. Whichever source the table came from --
    purchased, internal, or accumulated from avails documents -- it arrives
    here as one of those two shapes.
    """

    def __init__(self, by_county=None, by_zip=None, name="table"):
        self.by_county = dict(by_county or {})
        self.by_zip = dict(by_zip or {})
        self.name = name

    def market_for_county(self, fips):
        return self.by_county.get(fips)

    def market_for_zip(self, zip_code):
        return self.by_zip.get(zip_code)

    def counties_for_market(self, market_key):
        return [f for f, m in self.by_county.items() if m == market_key]

    def zips_for_market(self, market_key):
        return [z for z, m in self.by_zip.items() if m == market_key]

    def __len__(self):
        return len(self.by_county) + len(self.by_zip)


_market_lookup = None


def register_market_lookup(lookup):
    """Install the county/zip -> market table. Pass None to remove it."""
    global _market_lookup
    _market_lookup = lookup


def market_lookup():
    return _market_lookup


def market_lookup_available():
    return _market_lookup is not None and len(_market_lookup) > 0


_NO_LOOKUP_NOTE = (
    "No market lookup is registered, so zips can't be resolved to DMAs. "
    "County-to-DMA is licensed data -- see geo_targeting_roadmap.md. "
    "Markets can still be picked by hand; everything else here works without it."
)


# ---------------------------------------------------------------------------
# zip <-> county
# ---------------------------------------------------------------------------
def zips_to_counties(zips):
    """Which counties a zip list falls in.

    resolved: {zip: [county FIPS, ...]}, biggest land-area part first. A zip
    spanning several counties keeps ALL of them -- 10,186 of 33,791 zips do,
    and picking only the largest would quietly discard a third of the
    geography. Callers that need one county take the first.

    `unresolved` holds two genuinely different things and only one of them
    is a note here -- malformed input (not five digits) isn't explained by
    this function at all, because the caller already lists the offending
    entries verbatim and a second explanation of the same list is the
    cry-wolf pattern this project keeps having to fix. A well-formed zip
    with no county on file -- point and PO-box zips have no Census area to
    begin with -- is the expected, common case on a real list and gets
    exactly one calm note, not a warning: it's still a real zip, still goes
    into targeting and the export, it just can't be summarized by county or
    market.
    """
    data = _data()
    table = data["zip_counties"]
    resolved, unresolved, notes = {}, [], []
    unmapped = []
    for raw in parse_zip_list(zips):
        code = normalize_zip(raw)
        if code is None:
            unresolved.append(raw)
            continue
        counties = table.get(code)
        if not counties:
            unmapped.append(code)
            unresolved.append(raw)
            continue
        resolved[code] = list(counties)
    if unmapped:
        notes.append(
            f"{len(unmapped)} zip(s) have no county on file -- expected for point "
            f"and PO-box zips, which have no Census area to match against. "
            f"They're still included in targeting and the export; they just can't "
            f"be summarized by county or market.")
    return Resolution(resolved, unresolved, notes)


def parse_county_list(text):
    """Split a county list the way the avails documents write it.

    They give `SOMERSET NJ;BUCKS PA;NEW CASTLE DE` -- semicolon-separated
    NAME STATE, no type suffix. Newlines and commas are tolerated too, since a
    rep pasting one by hand won't reproduce the semicolons exactly.
    """
    if isinstance(text, (list, tuple, set)):
        raw = list(text)
    else:
        raw = re.split(r"[;\n]", str(text or ""))
    return [part.strip() for part in raw if str(part).strip()]


def counties_to_fips(counties):
    """County FIPS for names written as NAME STATE.

    resolved: {name as given: fips}. An ambiguous name is UNRESOLVED with the
    candidates named in `notes`, never resolved by picking one -- all seven
    ambiguous names are independent-city cases ("Baltimore MD" is both
    Baltimore County and Baltimore city), which are different places with
    different zips. Spelling the suffix out ("Baltimore County MD") resolves
    cleanly.
    """
    data = _data()
    index, ambiguous = data["county_index"], data.get("county_ambiguous", {})
    resolved, unresolved, notes = {}, [], []
    for name in parse_county_list(counties):
        key = "".join(c for c in name.lower() if c.isalnum())
        if key in index:
            resolved[name] = index[key]
            continue
        unresolved.append(name)
        if key in ambiguous:
            options = ", ".join(county_name(f) for f in ambiguous[key])
            notes.append(f"{name!r} could be {options} -- say which, or spell "
                         f"the county/city suffix out.")
    plain = [n for n in unresolved
             if "".join(c for c in n.lower() if c.isalnum()) not in ambiguous]
    if plain:
        notes.append(f"{len(plain)} county name(s) weren't recognised: "
                     f"{', '.join(plain[:5])}. They should read NAME STATE, "
                     f"e.g. 'Bucks PA'.")
    return Resolution(resolved, unresolved, notes)


def counties_to_zips(county_fips):
    """Every zip falling in each county. resolved: {fips: [zips]}.

    Accepts FIPS codes or NAME STATE strings, mixed -- the avails documents
    give names, everything else here gives codes, and a caller shouldn't have
    to know which it holds.
    """
    data = _data()
    known = data["counties"]
    # Resolve any name-shaped entries to FIPS first, carrying their notes.
    wanted_raw = (county_fips if isinstance(county_fips, (list, tuple, set))
                  else parse_county_list(county_fips))
    name_notes, named = [], {}
    to_resolve = [str(x).strip() for x in wanted_raw
                  if not str(x).strip().replace("-", "").isdigit()]
    if to_resolve:
        found = counties_to_fips(to_resolve)
        named = found.resolved
        name_notes = list(found.notes)
    by_county = {}
    for code, counties in data["zip_counties"].items():
        for fips in counties:
            by_county.setdefault(fips, []).append(code)

    resolved, unresolved, notes = {}, [], list(name_notes)
    for raw in wanted_raw:
        token = str(raw).strip()
        if token in named:                       # came in as a name
            fips = named[token]
        elif token.replace("-", "").isdigit():
            fips = token.zfill(5)
        else:
            unresolved.append(token)             # already explained by name_notes
            continue
        if fips not in known and fips not in by_county:
            unresolved.append(token)
            continue
        resolved[fips] = sorted(by_county.get(fips, []))
        if not resolved[fips]:
            notes.append(f"{county_name(fips) or fips} has no zips in the crosswalk.")
    bad_codes = [u for u in unresolved if u.replace("-", "").isdigit()]
    if bad_codes:
        notes.append(f"{len(bad_codes)} county code(s) weren't recognised. "
                     f"They should be 5-digit FIPS, e.g. 51013 for Arlington County, VA.")
    return Resolution(resolved, unresolved, notes)


def county_name(fips):
    entry = _data()["counties"].get(str(fips).strip().zfill(5))
    if not entry:
        return None
    return f"{entry['name']}, {entry['state']}"


# ---------------------------------------------------------------------------
# radius
# ---------------------------------------------------------------------------
def haversine_miles(a, b):
    lat1, lon1 = math.radians(a[0]), math.radians(a[1])
    lat2, lon2 = math.radians(b[0]), math.radians(b[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 2 * EARTH_RADIUS_MILES * math.asin(math.sqrt(h))


def geocode_address(address):
    """(lat, lon) for a street address via the Census Geocoder, or None.

    Free and keyless -- no account, no rate-limit key, nothing to leak. Any
    failure returns None so the caller reports it as unresolved rather than
    taking a page down; a geocoder being slow is not a reason a rep can't
    build a proposal.
    """
    query = urllib.parse.urlencode({
        "address": address, "benchmark": "Public_AR_Current", "format": "json"})
    try:
        with urllib.request.urlopen(f"{CENSUS_GEOCODER}?{query}",
                                    timeout=GEOCODER_TIMEOUT) as response:
            payload = json.load(response)
        matches = payload["result"]["addressMatches"]
        if not matches:
            return None
        point = matches[0]["coordinates"]
        return (float(point["y"]), float(point["x"]))
    except Exception:                                          # noqa: BLE001
        return None


def resolve_center(center, geocoder=geocode_address):
    """(point, label) for a radius centre given a zip or a street address."""
    code = normalize_zip(center)
    if code is not None:
        point = _data()["zip_points"].get(code)
        return (tuple(point), code) if point else (None, str(center))
    point = geocoder(str(center))
    return (point, str(center))


def radius_to_zips(centers, miles, geocoder=geocode_address):
    """Zips whose centroid falls within `miles` of any centre.

    `centers` is one zip/address or a list of them, each measured at the same
    radius; call it once per radius when a client's locations differ.

    Centroid-in-circle is the rule, matching what freemaptools does -- the
    tool account managers use today -- so a list built here agrees with one
    they'd have built by hand. A large zip whose centroid sits just outside
    the circle is excluded even though part of it lies inside; that is the
    convention, not an oversight.
    """
    data = _data()
    points = data["zip_points"]
    try:
        radius = float(miles)
    except (TypeError, ValueError):
        return Resolution([], [str(miles)], ["The radius must be a number of miles."])
    if radius <= 0:
        return Resolution([], [], ["A radius of zero miles matches nothing."])

    wanted = centers if isinstance(centers, (list, tuple, set)) else [centers]
    resolved, unresolved, notes = set(), [], []
    used = []
    for raw in wanted:
        point, label = resolve_center(raw, geocoder)
        if point is None:
            unresolved.append(str(raw))
            continue
        used.append(label)
        for code, other in points.items():
            if haversine_miles(point, other) <= radius:
                resolved.add(code)

    if unresolved:
        notes.append(
            f"{len(unresolved)} centre(s) couldn't be located: "
            f"{', '.join(unresolved)}. A zip must be one the Census recognises; "
            f"an address needs enough detail to geocode (street, city, state).")
    if used and not resolved:
        notes.append(f"No zip centroid falls within {radius:g} miles of "
                     f"{', '.join(used)}.")
    return Resolution(sorted(resolved), unresolved, notes)


# ---------------------------------------------------------------------------
# markets -- pluggable, and honest when unavailable
# ---------------------------------------------------------------------------
def zips_to_markets(zips):
    """Which markets a zip list covers, with a zip count per market.

    resolved: {market_key: {"zips": [...], "zip_count": int}}

    Needs a registered market lookup; without one this resolves nothing and
    says why, rather than returning an empty dict that reads like "this zip
    list covers no markets".
    """
    counties = zips_to_counties(zips)
    if not market_lookup_available():
        return Resolution({}, list(counties.resolved) + counties.unresolved,
                          counties.notes + [_NO_LOOKUP_NOTE])

    lookup = market_lookup()
    resolved = {}
    no_market = []
    notes = list(counties.notes)
    split = []
    for code, fips_list in counties.resolved.items():
        market = None
        if hasattr(lookup, "market_for_zip"):
            market = lookup.market_for_zip(code)
        markets = []
        if market:
            markets = [market]
        else:
            for fips in fips_list:
                found = lookup.market_for_county(fips)
                if found and found not in markets:
                    markets.append(found)
        if not markets:
            no_market.append(code)
            continue
        # A zip spanning counties in different DMAs is real. Credit it to the
        # first (largest-area) market and say so, rather than double-counting
        # it into both and inflating every total that reads this.
        if len(markets) > 1:
            split.append(code)
        entry = resolved.setdefault(markets[0], {"zips": [], "zip_count": 0})
        entry["zips"].append(code)
        entry["zip_count"] += 1

    for entry in resolved.values():
        entry["zips"].sort()
    if split:
        notes.append(
            f"{len(split)} zip(s) span counties in more than one market and were "
            f"counted in the larger one: {', '.join(sorted(split)[:5])}"
            f"{'...' if len(split) > 5 else ''}.")
    # `no_market` -- a zip that DID match a county but that county has no
    # market on file -- is a genuinely different gap than counties.unresolved
    # (no county at all). Reported separately so the note text stays true:
    # counting county-unresolved zips in here too was the exact bug that
    # made the same zips get explained twice, once wrongly ("resolved to no
    # market" when they never resolved to a county in the first place).
    if no_market:
        notes.append(f"{len(no_market)} zip(s) matched a county but no market is on "
                      f"file for it.")
    unresolved = list(counties.unresolved) + no_market
    return Resolution(resolved, unresolved, notes)


def market_to_zips(market_key):
    """Every zip in a market. resolved: [zips]."""
    if not market_lookup_available():
        return Resolution([], [str(market_key)], [_NO_LOOKUP_NOTE])

    lookup = market_lookup()
    zips = []
    if hasattr(lookup, "zips_for_market"):
        zips = list(lookup.zips_for_market(market_key) or [])
    if not zips:
        counties = lookup.counties_for_market(market_key) if hasattr(
            lookup, "counties_for_market") else []
        if not counties:
            return Resolution([], [str(market_key)],
                              [f"No counties or zips are mapped to {market_key!r}."])
        found = counties_to_zips(counties)
        zips = sorted({z for group in found.resolved.values() for z in group})
    return Resolution(sorted(set(zips)), [], [])
