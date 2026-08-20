"""Build the city/town label data the targeting map draws place names from
(roadmap section E follow-up -- county fills without any place names to orient by
still don't read as a real map).

    python build_map_places.py [--out map_places.json.gz]

Two Census sources, joined by state+place FIPS:

- 2024 Gazetteer Places (cb-equivalent gazetteer, not a boundary file) --
  name and centroid lat/lon for every place, incorporated or not. Same
  2024 Gazetteer family build_geo_crosswalk.py already draws its ZCTA and
  county centroids from.
- Population Estimates Program, sub-est2023.csv -- annual population
  estimates for incorporated places and consolidated cities (SUMLEV 162 /
  170). The Gazetteer carries no population field of its own; this is
  where "largest place in frame" actually comes from. A different Census
  program than the Gazetteer, but the same agency and the same
  public-domain status.

Only ~60% of Gazetteer places carry a population estimate this way -- the
other ~40% are Census-designated places that were never incorporated, which
is fine here: every incorporated city or town (anything a map viewer would
recognize by name) is covered, and an unincorporated CDP with no population
figure is exactly the kind of place a label wouldn't be worth drawing
anyway. No population floor is applied on the way into the file -- keeping
every matched place costs about 320KB gzipped (still smaller than
map_boundaries.json.gz), and "biggest in frame" is a render-time decision
that depends on what's actually on the map, not one to bake in here.

Downloads cache under ../proposal-builder-assets/geo, same as
build_geo_crosswalk.py and build_map_boundaries.py. Rerun to refresh;
output is deterministic for a given pair of input files.
"""
import csv
import gzip
import json
import sys
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
CACHE = HERE.parent / "proposal-builder-assets" / "geo"
DEFAULT_OUT = HERE / "map_places.json.gz"

GAZ_PLACE_URL = ("https://www2.census.gov/geo/docs/maps-data/data/gazetteer/"
                 "2024_Gazetteer/2024_Gaz_place_national.zip")
POP_URL = ("https://www2.census.gov/programs-surveys/popest/datasets/"
          "2020-2023/cities/totals/sub-est2023.csv")

# Incorporated place, and consolidated city -- the two SUMLEVs that carry a
# place-level population estimate in this file. Everything else (counties,
# county subdivisions that aren't also places, state totals) is noise here.
POPULATION_SUMLEVS = ("162", "170")

ROUND_DP = 4


def _fetch(url, filename):
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / filename
    if not path.exists():
        print(f"  downloading {url}", file=sys.stderr)
        urllib.request.urlretrieve(url, path)
    return path


def _load_population():
    """{state+place FIPS (7 digits): latest population estimate}."""
    path = _fetch(POP_URL, "sub-est2023.csv")
    pop = {}
    with open(path, encoding="latin-1") as f:
        for row in csv.DictReader(f):
            if row["SUMLEV"] not in POPULATION_SUMLEVS:
                continue
            try:
                pop[row["STATE"] + row["PLACE"]] = int(row["POPESTIMATE2023"])
            except ValueError:
                continue
    return pop


def _load_places():
    """(name, lat, lon, geoid) for every Gazetteer place, positional fields
    read by header name so a column reorder upstream can't misalign them."""
    path = _fetch(GAZ_PLACE_URL, "2024_Gaz_place_national.zip")
    out = []
    with zipfile.ZipFile(path) as zf:
        (name,) = zf.namelist()
        with zf.open(name) as f:
            header = [h.strip() for h in f.readline().decode("utf-8").rstrip("\n").split("\t")]
            for line in f:
                parts = line.decode("utf-8").rstrip("\n").split("\t")
                row = dict(zip(header, parts))
                out.append((row["NAME"].strip(), float(row["INTPTLAT"]),
                           float(row["INTPTLONG"]), row["GEOID"].strip()))
    return out


def build(out_path=DEFAULT_OUT):
    pop = _load_population()
    places = []
    for name, lat, lon, geoid in _load_places():
        population = pop.get(geoid)
        if population is None:
            continue
        places.append({"name": name, "lat": round(lat, ROUND_DP),
                       "lon": round(lon, ROUND_DP), "pop": population})

    payload = {
        "version": 1,
        "source": "US Census Bureau 2024 Gazetteer Files (place names/centroids) joined to "
                  "the Population Estimates Program's sub-est2023.csv (place population), "
                  "by state+place FIPS. Both public domain.",
        "places": places,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with gzip.open(out_path, "wt", encoding="utf-8") as f:
        json.dump(payload, f, separators=(",", ":"))

    print(f"{len(places)} places with a population estimate")
    print(f"wrote {out_path} ({out_path.stat().st_size:,} bytes)")
    return payload


def main():
    out = Path(sys.argv[sys.argv.index("--out") + 1]) if "--out" in sys.argv else DEFAULT_OUT
    build(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
