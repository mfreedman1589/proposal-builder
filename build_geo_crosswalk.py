"""Build the compact zip/county crosswalk the geo resolver reads.

    python build_geo_crosswalk.py [--out geo_crosswalk.json.gz]

Downloads three US Government files, joins them, and writes one small gzipped
artifact. Run it again to refresh; it is deterministic, so an unchanged input
produces an unchanged output.

Sources, all public domain as works of the US Government:

  * Census 2020 ZCTA-to-County relationship file (6.5 MiB) -- which counties
    each ZCTA falls in, with the land area of each part, so a ZCTA spanning
    two counties can be assigned to the one it mostly sits in AND still
    report the others.
  * Census 2024 ZCTA gazetteer (990 KiB) -- ZCTA centroids, which is what a
    radius search measures against.
  * Census 2024 county gazetteer (138 KiB) -- county names and centroids.

Deliberately NOT included: county -> DMA. That mapping is Nielsen's
intellectual property and has no clean public source (see
geo_targeting_roadmap.md). The resolver takes it as a pluggable lookup so it
can be supplied from wherever it legitimately comes from, and this script has
no opinion about that.

A note on ZCTAs. They are not ZIP codes: they are areal approximations of
USPS delivery routes, so point-only and PO-box-only ZIPs have no ZCTA at all
and simply cannot be resolved this way. That is a real gap in the data rather
than a bug, and the resolver reports those zips as unresolved instead of
quietly dropping them.
"""
import argparse
import csv
import gzip
import io
import json
import sys
import urllib.request
import zipfile
from pathlib import Path

CACHE = Path(__file__).resolve().parent.parent / "proposal-builder-assets" / "geo"
DEFAULT_OUT = Path(__file__).resolve().parent / "geo_crosswalk.json.gz"

ZCTA_COUNTY_URL = ("https://www2.census.gov/geo/docs/maps-data/data/rel2020/"
                   "zcta520/tab20_zcta520_county20_natl.txt")
GAZ_ZCTA_URL = ("https://www2.census.gov/geo/docs/maps-data/data/gazetteer/"
                "2024_Gazetteer/2024_Gaz_zcta_national.zip")
GAZ_COUNTY_URL = ("https://www2.census.gov/geo/docs/maps-data/data/gazetteer/"
                  "2024_Gazetteer/2024_Gaz_counties_national.zip")


def _fetch(url, name):
    CACHE.mkdir(parents=True, exist_ok=True)
    target = CACHE / name
    if not target.exists():
        print(f"  downloading {name}...")
        urllib.request.urlretrieve(url, target)
    print(f"  {name}: {target.stat().st_size / 2**20:.1f} MiB")
    return target


def _gazetteer(path):
    """Rows of a zipped, tab-delimited gazetteer file."""
    with zipfile.ZipFile(path) as z:
        with z.open(z.namelist()[0]) as f:
            for row in csv.DictReader(io.TextIOWrapper(f, encoding="utf-8-sig"),
                                      delimiter="\t"):
                yield {k.strip(): (v.strip() if isinstance(v, str) else v)
                       for k, v in row.items() if k}


def build():
    print("Fetching sources:")
    rel = _fetch(ZCTA_COUNTY_URL, "zcta_county_2020.txt")
    gz_zcta = _fetch(GAZ_ZCTA_URL, "gaz_zcta.zip")
    gz_county = _fetch(GAZ_COUNTY_URL, "gaz_counties.zip")

    # --- zip -> counties, biggest land-area part first --------------------
    parts = {}
    skipped = 0
    with open(rel, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f, delimiter="|"):
            zcta = (row.get("GEOID_ZCTA5_20") or "").strip()
            fips = (row.get("GEOID_COUNTY_20") or "").strip()
            if not zcta or not fips:
                skipped += 1          # county rows with no ZCTA overlap
                continue
            try:
                land = int(row.get("AREALAND_PART") or 0)
            except ValueError:
                land = 0
            parts.setdefault(zcta, []).append((land, fips))
    zip_counties = {z: [fips for _, fips in sorted(v, reverse=True)]
                    for z, v in parts.items()}
    multi = sum(1 for v in zip_counties.values() if len(v) > 1)
    print(f"\n  zips with a county: {len(zip_counties):,}  "
          f"(spanning >1 county: {multi:,});  rows without a ZCTA skipped: {skipped:,}")

    # --- centroids ---------------------------------------------------------
    zip_points = {}
    for row in _gazetteer(gz_zcta):
        geoid = row.get("GEOID")
        if geoid:
            zip_points[geoid] = [round(float(row["INTPTLAT"]), 5),
                                 round(float(row["INTPTLONG"]), 5)]
    counties = {}
    for row in _gazetteer(gz_county):
        fips = row.get("GEOID")
        if fips:
            counties[fips] = {"name": row["NAME"], "state": row["USPS"],
                              "point": [round(float(row["INTPTLAT"]), 5),
                                        round(float(row["INTPTLONG"]), 5)]}
    print(f"  zip centroids: {len(zip_points):,}   counties: {len(counties):,}")

    missing_points = [z for z in zip_counties if z not in zip_points]
    if missing_points:
        print(f"  NOTE: {len(missing_points):,} zips have a county but no centroid "
              f"-- radius searches will report them unresolved")

    return {
        "version": 1,
        "sources": {
            "zcta_county": ZCTA_COUNTY_URL,
            "zcta_gazetteer": GAZ_ZCTA_URL,
            "county_gazetteer": GAZ_COUNTY_URL,
        },
        "zip_counties": zip_counties,
        "zip_points": zip_points,
        "counties": counties,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = ap.parse_args(argv)

    data = build()
    payload = json.dumps(data, separators=(",", ":"), sort_keys=True).encode("utf-8")
    # GzipFile rather than gzip.open, so mtime can be pinned -- otherwise the
    # artifact's bytes change on every rebuild and a "did the data change?"
    # diff is answered by the clock instead of by the data.
    with open(args.out, "wb") as raw:
        with gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as f:
            f.write(payload)
    print(f"\nwrote {args.out}  ({args.out.stat().st_size / 2**20:.2f} MiB gzipped, "
          f"{len(payload) / 2**20:.1f} MiB raw)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
