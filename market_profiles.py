"""The market viewer profile row set: which DMAs exist, and which have a slide.

Two facts have to be combined and they come from different places, which is
the whole reason this module exists rather than the seed being a third
checked-in file that can drift from both:

  * `market_profiles_index.json` -- what the SOURCE DECK contains. 205 market
    slides plus a national roll-up, read off renderings of the slides (they
    carry no text and no notes) and cross-checked by tests/validate_markets.py.

  * CANONICAL_DMAS -- what NIELSEN defines, in rank order. 210 markets.

The deck covers 205 of the 210. Five DMAs -- Honolulu, Palm Springs,
Anchorage, Fairbanks, Juneau -- have no slide, and they are still real
markets a rep sells into, so they are still selectable target DMAs. What
they lack is an image, and that lack is carried explicitly as
`has_profile_slide=False` rather than being left for a caller to infer from
a null path. A rep is told at selection time; nothing is silently hidden and
no control is silently dropped.

Ordering is by Nielsen rank, not alphabetical and not slide order. The deck
is alphabetical because it's a reference document; a picker is a tool, and
the markets people actually sell are at the top of the rank list.
"""
import json
import re
import unicodedata
from pathlib import Path

INDEX_PATH = Path(__file__).resolve().parent / "market_profiles_index.json"

# The key the national roll-up is addressed by. It is a selectable option --
# national advertisers are a real case and the deck ships the slide for them.
NATIONAL_KEY = "total_us"

# Nielsen DMAs in rank order, and the ONE copy of this list -- shipping code
# owns it and tests/validate_markets.py imports it from here, rather than
# each keeping its own. A hand-typed second copy was tried for exactly one
# commit and was wrong on first comparison against the validator's: seven
# real DMAs dropped, five invented, and "Fort Myers-Naples" / "Saint Louis" /
# "Columbus-Opelika" quietly misspelled. Two lists of 210 names cannot be
# kept in agreement by attention, and the failure is silent -- an invented
# market becomes a selectable option with no slide behind it.
#
# The validator still CHECKS the reading against this list; it just no longer
# holds a rival definition of it. Any edit here has to keep it passing.
CANONICAL_DMAS = [
    'New York', 'Los Angeles', 'Chicago', 'Dallas-Fort Worth',
    'Philadelphia', 'Houston', 'Atlanta', 'Washington-Hagerstown',
    'Boston-Manchester', 'San Francisco-Oakland-San Jose',
    'Tampa-St Petersburg-Sarasota', 'Phoenix-Prescott', 'Seattle-Tacoma',
    'Detroit', 'Orlando-Daytona Beach-Melbourne', 'Minneapolis-Saint Paul',
    'Denver', 'Miami-Fort Lauderdale', 'Cleveland-Akron-Canton',
    'Sacramento-Stockton-Modesto', 'Charlotte',
    'Raleigh-Durham-Fayetteville', 'Portland, OR', 'Saint Louis',
    'Indianapolis', 'Nashville', 'Pittsburgh', 'Salt Lake City',
    'Baltimore', 'San Diego', 'San Antonio', 'Hartford-New Haven',
    'Kansas City', 'Austin', 'Columbus, OH',
    'Greenville-Spartanburg-Asheville-Anderson', 'Cincinnati', 'Milwaukee',
    'West Palm Beach-Fort Pierce', 'Las Vegas', 'Jacksonville',
    'Harrisburg-Lancaster-Lebanon-York',
    'Grand Rapids-Kalamazoo-Battle Creek',
    'Norfolk-Portsmouth-Newport News', 'Birmingham-Anniston-Tuscaloosa',
    'Greensboro-High Point-Winston Salem', 'Oklahoma City',
    'Albuquerque-Santa Fe', 'Louisville', 'New Orleans', 'Memphis',
    'Providence-New Bedford', 'Fort Myers-Naples', 'Buffalo',
    'Fresno-Visalia', 'Richmond-Petersburg',
    'Mobile-Pensacola-Fort Walton Beach', 'Little Rock-Pine Bluff',
    'Wilkes Barre-Scranton-Hazleton', 'Knoxville', 'Tulsa',
    'Albany-Schenectady-Troy', 'Lexington', 'Dayton', 'Tucson-Sierra Vista',
    'Spokane', 'Des Moines-Ames', 'Green Bay-Appleton', 'Honolulu',
    'Roanoke-Lynchburg', 'Wichita-Hutchinson Plus',
    'Flint-Saginaw-Bay City', 'Omaha', 'Springfield, MO',
    'Huntsville-Decatur-Florence', 'Columbia, SC', 'Madison',
    'Portland-Auburn', 'Rochester, NY',
    'Harlingen-Weslaco-Brownsville-McAllen', 'Toledo',
    'Charleston-Huntington', 'Waco-Temple-Bryan', 'Savannah',
    'Charleston, SC', 'Chattanooga', 'Colorado Springs-Pueblo', 'Syracuse',
    'El Paso-Las Cruces', 'Paducah-Cape Girardeau-Harrisburg', 'Shreveport',
    'Champaign-Springfield-Decatur', 'Burlington-Plattsburgh',
    'Cedar Rapids-Waterloo-Iowa City-Dubuque', 'Baton Rouge',
    'Fort Smith-Fayetteville-Springdale-Rogers', 'Myrtle Beach-Florence',
    'Boise', 'Jackson, MS', 'South Bend-Elkhart', 'Tri-Cities',
    'Greenville-New Bern-Washington', 'Reno',
    'Davenport-Rock Island-Moline', 'Tallahassee-Thomasville',
    'Tyler-Longview-Lufkin-Nacogdoches', 'Lincoln-Hastings-Kearney',
    'Augusta-Aiken', 'Evansville', 'Fort Wayne', 'Sioux Falls-Mitchell',
    'Johnstown-Altoona-State College', 'Fargo-Valley City',
    'Yakima-Pasco-Richland-Kennewick', 'Springfield-Holyoke',
    'Traverse City-Cadillac', 'Lansing', 'Youngstown', 'Macon', 'Eugene',
    'Montgomery-Selma', 'Peoria-Bloomington',
    'Santa Barbara-Santa Maria-San Luis Obispo', 'Lafayette, LA',
    'Bakersfield', 'Wilmington', 'Columbus-Opelika', 'Monterey-Salinas',
    'La Crosse-Eau Claire', 'Corpus Christi', 'Salisbury', 'Amarillo',
    'Wausau-Rhinelander', 'Columbus-Tupelo-West Point-Houston',
    'Columbia-Jefferson City', 'Chico-Redding', 'Rockford',
    'Duluth-Superior', 'Medford-Klamath Falls', 'Lubbock', 'Topeka',
    'Monroe-El Dorado', 'Beaumont-Port Arthur', 'Odessa-Midland',
    'Palm Springs', 'Anchorage', 'Bismarck-Minot-Dickinson-Williston',
    'Panama City', 'Sioux City', 'Wichita Falls-Lawton', 'Joplin-Pittsburg',
    'Albany, GA', 'Rochester-Mason City-Austin', 'Erie',
    'Idaho Falls-Pocatello-Jackson', 'Bangor', 'Gainesville',
    'Biloxi-Gulfport', 'Terre Haute', 'Sherman-Ada', 'Missoula',
    'Binghamton', 'Wheeling-Steubenville', 'Yuma-El Centro', 'Billings',
    'Abilene-Sweetwater', 'Bluefield-Beckley-Oak Hill',
    'Hattiesburg-Laurel', 'Rapid City', 'Dothan', 'Utica',
    'Clarksburg-Weston', 'Harrisonburg', 'Jackson, TN',
    'Quincy-Hannibal-Keokuk', 'Charlottesville', 'Lake Charles',
    'Elmira-Corning', 'Watertown', 'Bowling Green', 'Marquette',
    'Jonesboro', 'Alexandria', 'Laredo', 'Butte-Bozeman', 'Bend',
    'Grand Junction-Montrose', 'Twin Falls', 'Lafayette, IN', 'Lima',
    'Great Falls', 'Meridian', 'Cheyenne-Scottsbluff', 'Parkersburg',
    'Greenwood-Greenville', 'Eureka', 'San Angelo', 'Casper-Riverton',
    'Mankato', 'Ottumwa-Kirksville', 'St. Joseph', 'Fairbanks',
    'Zanesville', 'Victoria', 'Helena', 'Presque Isle', 'Juneau', 'Alpena',
    'North Platte', 'Glendive',
]


def slugify(name):
    """A stable key for a market name.

    Punctuation is dropped rather than transliterated, so "Albany, GA" and
    "Portland, OR" keep their state qualifier as a word -- which is what
    keeps them distinct from "Albany-Schenectady-Troy" and "Portland-Auburn".
    """
    text = unicodedata.normalize("NFKD", name)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower().replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def load_index(path=None):
    """The slide index, as tests/validate_markets.py wrote it."""
    return json.loads(Path(path or INDEX_PATH).read_text(encoding="utf-8"))


def build_rows(index=None):
    """Every selectable target market, in Nielsen rank order.

    Returns dicts shaped for the `market_profiles` table, each carrying
    `has_profile_slide` so a caller never has to infer availability from a
    null image path.
    """
    entries = load_index() if index is None else index

    by_dma = {}
    national = None
    for entry in entries:
        if entry.get("kind") == "national":
            national = entry
        elif entry.get("dma"):
            by_dma[entry["dma"]] = entry

    rows = []
    for rank, dma in enumerate(CANONICAL_DMAS, start=1):
        entry = by_dma.get(dma)
        rows.append({
            "key": slugify(dma),
            "label": entry["label"] if entry else dma,
            "dma": dma,
            "kind": "dma",
            "rank": rank,
            "slide_number": entry["slide"] if entry else None,
            "has_profile_slide": entry is not None,
        })

    # The roll-up sorts last: it's a real option, but a rep scanning for a
    # city shouldn't have to read past it.
    rows.append({
        "key": NATIONAL_KEY,
        "label": national["label"] if national else "Total U.S.",
        "dma": None,
        "kind": "national",
        "rank": len(CANONICAL_DMAS) + 1,
        "slide_number": national["slide"] if national else None,
        "has_profile_slide": national is not None,
    })
    return rows


def markets_without_a_slide(rows=None):
    """The DMAs Premion never authored a profile slide for.

    Selectable, but the picker has to say so -- see the schema comment.
    """
    return [r for r in (rows or build_rows()) if not r["has_profile_slide"]]


if __name__ == "__main__":
    rows = build_rows()
    missing = markets_without_a_slide(rows)
    print(f"{len(rows)} selectable markets "
          f"({len(rows) - len(missing)} with a profile slide, {len(missing)} without)")
    for r in missing:
        print(f"  no slide: rank {r['rank']:>3}  {r['label']}  ({r['key']})")
