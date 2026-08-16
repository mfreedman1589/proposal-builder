"""Cross-check the market names read off the profile slides against the
canonical Nielsen DMA list.

The slide labels are Premion's own ABBREVIATED DMA names ("Cedar Rapids-Wtrlo",
"Champaign-Springfld"), so exact equality is the wrong test -- it would flag
dozens of correct readings. The test used instead is a subsequence match on
letters only: every letter of the label must appear, in order, within the
canonical name. That admits vowel-dropped abbreviations while still rejecting
a misread, because a misread introduces letters that aren't there.
"""
import json
import re
import sys
from pathlib import Path

# Read from the slides (see contact sheets), slide order = alphabetical.
EXTRACTED = """Abilene-Sweetwater
Albany-Schenectady-Troy
Albany, GA
Albuquerque-Santa Fe
Alexandria, LA
Alpena
Amarillo
Atlanta
Augusta
Austin
Bakersfield
Baltimore
Bangor
Baton Rouge
Beaumont-Port Arthur
Bend, OR
Billings
Biloxi-Gulfport
Binghamton
Birmingham
Bluefield-Beckley-Oak Hill
Boise
Boston
Bowling Green
Buffalo
Burlington-Plattsburgh
Butte-Bozeman
Casper-Riverton
Cedar Rapids-Wtrlo
Champaign-Springfld
Charleston-Huntington
Charleston, SC
Charlotte
Charlottesville
Chattanooga
Cheyenne-Scottsbluf
Chicago
Chico-Redding
Cincinnati
Clarksburg-Weston
Cleveland-Akron
Colorado Springs-Pueblo
Columbia-Jefferson City
Columbia, SC
Columbus-Tupelo-West Point
Columbus, GA
Columbus, OH
Corpus Christi
Dallas-Ft. Worth
Davenport- R. Island-Moline
Dayton
Denver
Des Moines-Ames
Detroit
Dothan
Duluth-Superior
El Paso
Elmira
Erie
Eugene
Eureka
Evansville
Fargo-Valley City
Flint-Saginaw-Bay City
Florence-Myrtle Beach
Fresno-Visalia
Ft. Myers-Naples
Ft. Smith-Fay-Sprngdl-Rgrs
Ft. Wayne
Gainesville
Glendive
Grand Junction-Montrose
Grand Rapids-Kalmzoo-B.Crk
Great Falls
Green Bay-Appleton
Greensboro-H.Point-W.Salem
Greenville-N.Bern-Washngtn
Greenvll-Spart-Ashevll-And
Greenwood-Greenville
Harlingen-Wslco-Brnsvl-Mca
Harrisburg-Lncstr-Leb-York
Harrisonburg
Hartford-New Haven
Hattiesburg-Laurel
Helena
Houston
Huntsville-Decatur
Idaho Falls-Pocatello
Indianapolis
Jackson, MS
Jackson, TN
Jacksonville
Johnstown-Altoona
Jonesboro
Joplin-Pittsburg
Kansas City
Knoxville
La Crosse-Eau Claire
Lafayette, IN
Lafayette, LA
Lake Charles
Lansing
Laredo
Las Vegas
Lexington
Lima
Lincoln-Hastings-Krny
Little Rock-Pine Bluff
Los Angeles
Louisville
Lubbock
Macon
Madison
Mankato
Marquette
Medford-Klamath Falls
Memphis
Meridian
Miami-Ft. Lauderdale
Milwaukee
Minneapolis-St. Paul
Minot-Bismarck-Dickinson
Missoula
Mobile-Pensacola
Monroe-El Dorado
Monterey-Salinas
Montgomery-Selma
Nashville
New Orleans
New York City
Norfolk-Portsmth-Newpt Nws
North Platte
Odessa-Midland
Oklahoma City
Omaha
Orlando-Daytona Bch-Melbrn
Ottumwa-Kirksville
Paducah-C.Gird-Harbg-Mt Vn
Panama City
Parkersburg
Peoria-Bloomington
Philadelphia
Phoenix
Pittsburgh
Portland-Auburn
Portland, OR
Presque Isle
Providence-New Bedford
Quincy-Hannibal-Keokuk
Raleigh-Durham
Rapid City
Reno
Richmond-Petersburg
Roanoke-Lynchburg
Rochester, NY
Rochestr-Mason City-Austin
Rockford
Sacramento-Stkton-Modesto
Salisbury
Salt Lake City
San Angelo
San Antonio
San Diego
San Francisco-Oak-San Jose
Santa Barbra-Sanmar
Savannah
Seattle-Tacoma
Sherman-Ada
Shreveport
Sioux City
Sioux Falls-Mitchell
South Bend-Elkhart
Spokane
Springfield-Holyoke
Springfield, MO
St. Joseph
St. Louis
Syracuse
Tallahassee-Thomasville
Tampa-St. Pete-Sarasota
Terre Haute
Toledo
Topeka
Traverse City-Cadillac
Tri-Cities, TN-VA
Tucson-Sierra Vista
Tulsa
Twin Falls
Tyler-Longview
Utica
Victoria
Waco-Temple-Bryan
Washington, DC
Watertown
Wausau-Rhinelander
West Palm Beach-Ft. Pierce
Wheeling-Steubenville
Wichita Falls-Lawton
Wichita-Hutchinson Plus
Wilkes Barre-Scranton
Wilmington
Yakima-Pasco-Rchlnd
Youngstown
Yuma-El Centro
Zanesville
Total U.S.""".splitlines()

# The canonical list lives in market_profiles.py -- shipping code owns it and
# this validator imports it, so there is exactly one definition of what the
# 210 Nielsen markets are. Provenance: ranks 1-200 from ustvdb.com (2024-25
# season), 201-210 from Wikipedia's media-market list, names verbatim.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from market_profiles import CANONICAL_DMAS as CANONICAL  # noqa: E402

# Not a DMA: a national roll-up slide that rides at the end of the deck.
NATIONAL = "Total U.S."

# Only unambiguous word-level short forms. Single letters are deliberately
# NOT expanded: mapping "W." to "west" turned "W.Salem" into "westsalem",
# which then failed to match "Winston Salem" -- the expansion invented
# letters that aren't in the name and reported a correct reading as a
# misread.
ABBREV = {"ft": "fort", "st": "saint", "mt": "mount"}

STATE_CODES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "HI",
    "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN",
    "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH",
    "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA",
    "WV", "WI", "WY",
}


def tokens(name):
    out = []
    for raw in re.split(r"[^A-Za-z]+", name):
        if raw:
            out.append(ABBREV.get(raw.lower(), raw.lower()))
    return out


def split_state(toks):
    """(name tokens, {state codes}) for a label like "Jackson, TN".

    The state is a QUALIFIER, not part of the name, and it is the only thing
    telling "Jackson, MS" and "Jackson, TN" apart -- stripping it from both
    sides made three pairs of markets collapse onto one DMA each.
    """
    out, states = list(toks), set()
    while len(out) > 1 and out[-1].upper() in STATE_CODES:
        states.add(out.pop().upper())
    return out, states


def is_subsequence(short, long):
    it = iter(long)
    return all(ch in it for ch in short)


def compatible(label_token, canonical_token):
    """Does this label token abbreviate this canonical component?

    ONE-DIRECTIONAL, and that is the whole point. Allowing the reverse --
    canonical token as a subsequence of the label -- let two-letter state
    components match almost anything: "Boston" resolved to "Jackson, TN"
    because "tn" is a subsequence of "bosTon". The labels are the
    abbreviations; the canonical names are not.
    """
    return is_subsequence(label_token, canonical_token)


def unmatched_against(label_tokens, canonical_tokens):
    """(label tokens with no partner, canonical tokens left over), matching
    each canonical token at most once and ignoring ORDER -- the slides list
    a market's component cities in a different order from Nielsen in at
    least two cases ("Florence-Myrtle Beach", "Minot-Bismarck-Dickinson")."""
    pool = list(canonical_tokens)
    missed = []
    for t in label_tokens:
        hit = next((x for x in pool if compatible(t, x)), None)
        if hit is None:
            missed.append(t)
        else:
            pool.remove(hit)
    return missed, pool


def match(label):
    """(canonical DMA, label tokens it could not account for).

    Best candidate = fewest unaccounted-for label tokens, then fewest
    leftover canonical components, then shortest. A non-empty second element
    means the reading matched a DMA but not completely, which is reported for
    review rather than quietly accepted -- that is where a real misread would
    surface.
    """
    lt, lstates = split_state(tokens(label))
    best = None
    for c in CANONICAL:
        ct, cstates = split_state(tokens(c))
        # A state on both sides must agree. A state on only one side is no
        # constraint: Nielsen calls the Columbus GA market "Columbus-Opelika".
        if lstates and cstates and not (lstates & cstates):
            continue
        missed, leftover = unmatched_against(lt, ct)
        score = (len(missed), len(leftover), len("".join(ct)))
        if best is None or score < best[0]:
            best = (score, c, missed)
    if best is None or best[0][0] >= len(lt):
        return None, lt                      # nothing matched at all
    return best[1], best[2]


def main():
    labels = [l.strip() for l in EXTRACTED if l.strip()]
    print(f"slides read: {len(labels)}")
    print(f"canonical DMAs: {len(CANONICAL)}")
    print()

    markets = [l for l in labels if l != NATIONAL]
    resolved, unresolved, partial, mapping = {}, [], [], []
    for i, label in enumerate(labels, start=1):
        if label == NATIONAL:
            mapping.append({"slide": i, "label": label, "dma": None,
                            "kind": "national"})
            continue
        dma, missed = match(label)
        if dma is None:
            unresolved.append((i, label))
        else:
            resolved.setdefault(dma, []).append((i, label))
            if missed:
                partial.append((i, label, dma, missed))
        mapping.append({"slide": i, "label": label, "dma": dma, "kind": "dma",
                        "unaccounted": missed})

    print(f"market slides: {len(markets)}   "
          f"resolved: {len(markets) - len(unresolved)}   "
          f"UNRESOLVED: {len(unresolved)}")
    for i, label in unresolved:
        print(f"   !! slide {i}: {label!r} matches no DMA")

    print(f"\npartial matches (a DMA, but not every word accounted for): {len(partial)}")
    for i, label, dma, missed in partial:
        print(f"   ?  slide {i}: {label!r} -> {dma!r}   unaccounted: {missed}")

    dupes = {d: v for d, v in resolved.items() if len(v) > 1}
    print(f"\ncanonical DMAs claimed by more than one slide: {len(dupes)}")
    for d, v in dupes.items():
        print(f"   !! {d!r} <- {v}")

    missing = [c for c in CANONICAL if c not in resolved]
    print(f"\nDMAs with no profile slide: {len(missing)}")
    for c in missing:
        print(f"   -  {c}")

    # The deck is alphabetical, which is a free consistency check: a misread
    # bad enough to change the leading letters would land out of sequence.
    # Compared on letters only, because the deck sorts "Albany, GA" after
    # "Albany-Schenectady-Troy" -- the comma is a qualifier, not part of the
    # sort key, and comparing raw strings flags six correct rows.
    def sort_key(name):
        return "".join(ch for ch in name.lower() if ch.isalnum())

    out_of_order = [(a, b) for a, b in zip(markets, markets[1:])
                    if sort_key(a) > sort_key(b)]
    print(f"\nalphabetical-order breaks: {len(out_of_order)}")
    for a, b in out_of_order:
        print(f"   !! {a!r} then {b!r}")

    # Default to the checked-in artifact, so re-running regenerates the file
    # the app reads rather than dropping a copy in whatever the cwd happens
    # to be. The repo root is this file's parent's parent (tests/ -> root).
    default = Path(__file__).resolve().parent.parent / "market_profiles_index.json"
    dest = Path(sys.argv[1]) if len(sys.argv) > 1 else default
    dest.write_text(json.dumps(mapping, indent=2), encoding="utf-8")
    print(f"\nindex written to {dest}")
    return 1 if (unresolved or dupes or out_of_order) else 0


if __name__ == "__main__":
    sys.exit(main())
