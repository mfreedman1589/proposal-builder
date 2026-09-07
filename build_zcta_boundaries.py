"""Build the ZCTA (zip-area) polygon data the attribution report's heat map
fills, packaged PER STATE.

    python build_zcta_boundaries.py                    # the default states
    python build_zcta_boundaries.py --states NC PA     # just these
    python build_zcta_boundaries.py --add GA           # keep existing, add GA
    python build_zcta_boundaries.py --list             # what's built today

Why this exists at all: `geo_crosswalk.json.gz` carries zip POINTS
(33,791 centroids) and `map_boundaries.json.gz` carries county and state
POLYGONS -- but nothing anywhere carries zip-level geometry. The report's
zip slide was therefore shading centroid dots and calling itself a heat
map. The dashboard map it replaces shades real ZCTA areas, so this is a
data gap, not a drawing problem.

**Source: cb_2020_us_zcta520_500k, not TIGER/Line.** Same Census
cartographic (pre-generalized) product family `build_map_boundaries.py`
already uses. Raw TIGER/Line ZCTA is 503 MiB; the cartographic version is
63.6 MiB and is already generalized for mapping. Public domain as a work
of the US Government, same provenance class as every other geo source
here.

**Simplification: Douglas-Peucker at 0.001 degrees**, stdlib only -- no
shapely, no GDAL, matching the existing builder's own no-new-dependency
rule. That tolerance was chosen by pixel math against real renders rather
than by feel: on the zip slide's MapRegion (5.80in at 200 DPI, ~1010px of
map once the legend column is taken), a wide campaign like Mattress
Warehouse spans 4.20 degrees of longitude -> 0.00416 deg/px, so 0.001 deg
is 0.24px. A tight single-market campaign like Cardinal Plumbing spans
0.59 degrees -> 0.001 deg is 1.73px. **The tight case is the binding one**;
anything coarser starts showing on a single-market map, which is the
commoner shape.

**Per state, in its own file, and NOT appended to map_boundaries.json.gz.**
Every targeting-map render loads the county/state basemap; only the
report's choropleth needs ZCTAs, and the proposal builder must not pay
~1.65 MiB it never reads. targeting_map loads this lazily and degrades
silently when a state is missing (centroid dots plus a warning, never a
crash).

Measured output, gzipped, at 0.001 degrees:
    DC     57 ZCTAs      3 KiB
    MD    477 ZCTAs    148 KiB
    VA    903 ZCTAs    440 KiB
    NC    853 ZCTAs    477 KiB
    PA  1,833 ZCTAs    581 KiB
    ---------------------------
    5 states          1.65 MiB      national (33,484) would be 14.7 MiB

Downloads cache under ../proposal-builder-assets/geo, same as
build_geo_crosswalk.py / build_map_boundaries.py / build_map_places.py.
"""
import argparse
import gzip
import json
import math
import sys
import urllib.request
import zipfile
from pathlib import Path

import build_map_boundaries as bmb

HERE = Path(__file__).resolve().parent
CACHE = HERE.parent / "proposal-builder-assets" / "geo"
OUT_DIR = HERE / "zcta_boundaries"

SOURCE_URL = "https://www2.census.gov/geo/tiger/GENZ2020/shp/cb_2020_us_zcta520_500k.zip"
SOURCE_FILE = "cb_2020_us_zcta520_500k.zip"
SHP_STEM = "cb_2020_us_zcta520_500k"

# The states Premion currently sells in: the DC market (DC/MD/VA), Mattress
# Warehouse's NC footprint, and Harrisburg (PA). Adding a sixth is a one-flag
# rerun -- see this module's docstring and CLAUDE.md.
DEFAULT_STATES = ("DC", "MD", "VA", "NC", "PA")

SIMPLIFY_TOLERANCE = 0.001   # degrees -- see the docstring's pixel math
ROUND_DP = bmb.ROUND_DP      # 4, matching the county/state payload exactly


def _state_of(path):
    """"zcta_NC.json.gz" -> "NC". Path.stem strips ONE suffix, so a naive
    stem.split("_")[1] yields "NC.json" -- which merely looked untidy in
    --list but would have made --add treat an existing state as a new one
    and silently rebuild every file each time."""
    return path.name[len("zcta_"):].split(".")[0]


def _fetch():
    CACHE.mkdir(parents=True, exist_ok=True)
    dest = CACHE / SOURCE_FILE
    if not dest.exists() or dest.stat().st_size < 1_000_000:
        print(f"  downloading {SOURCE_URL} ...")
        urllib.request.urlretrieve(SOURCE_URL, dest)
    print(f"  source: {dest} ({dest.stat().st_size / 1024 / 1024:.1f} MiB)")
    return dest


def simplify(points, tolerance):
    """Douglas-Peucker. Iterative rather than recursive: a real ZCTA ring
    can carry thousands of vertices, and the recursive form overflows
    Python's default stack on the worst of them (raising the limit instead
    just moves the crash to a C-level stack overflow, which is
    uncatchable)."""
    if len(points) < 3:
        return list(points)
    keep = [False] * len(points)
    keep[0] = keep[-1] = True
    stack = [(0, len(points) - 1)]
    while stack:
        start, end = stack.pop()
        if end <= start + 1:
            continue
        ax, ay = points[start]
        bx, by = points[end]
        dx, dy = bx - ax, by - ay
        denom = math.hypot(dx, dy)
        worst, index = -1.0, start
        for i in range(start + 1, end):
            px, py = points[i]
            if denom:
                dist = abs(dy * px - dx * py + bx * ay - by * ax) / denom
            else:
                dist = math.hypot(px - ax, py - ay)
            if dist > worst:
                worst, index = dist, i
        if worst > tolerance:
            keep[index] = True
            stack.append((start, index))
            stack.append((index, end))
    return [points[i] for i, k in enumerate(keep) if k]


def _zip_to_state():
    """{zip: state} from the existing crosswalk -- the ZCTA shapefile's own
    attributes carry no state at all (ZCTA5CE20/GEOID20/NAME20/LSAD20/
    ALAND20/AWATER20), so state assignment comes from the zip's county.
    A zip spanning two counties takes the first, which for state purposes
    is only ambiguous on a genuine state line -- rare, and the map draws
    the polygon either way; it only affects which file it's packaged in."""
    import geo_resolver
    data = geo_resolver._data()
    counties = data["counties"] if "counties" in data else {}
    out = {}
    for code, fips_list in data["zip_counties"].items():
        for fips in fips_list:
            state = (counties.get(fips) or {}).get("state")
            if state:
                out[code] = state
                break
    return out


def build(states, keep_existing=False):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    archive = _fetch()
    with zipfile.ZipFile(archive) as zf:
        shp = zf.read(f"{SHP_STEM}.shp")
        dbf = zf.read(f"{SHP_STEM}.dbf")
    records = list(bmb._read_dbf(dbf))
    shapes = list(bmb._read_shp_polygons(shp))
    print(f"  {len(records)} ZCTA records read")

    zip_state = _zip_to_state()
    wanted = {s.upper() for s in states}
    by_state = {}
    raw_points = simplified_points = 0
    unplaced = 0
    for record, (parts, points) in zip(records, shapes):
        code = str(record.get("ZCTA5CE20") or "").strip()
        state = zip_state.get(code)
        if state is None:
            unplaced += 1
            continue
        if state not in wanted:
            continue
        rings_out = []
        for ring in bmb._rings(parts, points):
            raw_points += len(ring)
            reduced = simplify([tuple(p) for p in ring], SIMPLIFY_TOLERANCE)
            simplified_points += len(reduced)
            if len(reduced) >= 3:
                rings_out.append([[round(x, ROUND_DP), round(y, ROUND_DP)]
                                 for x, y in reduced])
        if rings_out:
            by_state.setdefault(state, {})[code] = {"rings": rings_out}

    if raw_points:
        print(f"  simplified at {SIMPLIFY_TOLERANCE} deg: "
              f"{raw_points:,} -> {simplified_points:,} points "
              f"({simplified_points / raw_points * 100:.0f}%)")
    if unplaced:
        print(f"  {unplaced} ZCTA(s) had no county in the crosswalk and were skipped "
              f"(territories and retired codes)")

    total = 0
    for state in sorted(wanted):
        zctas = by_state.get(state, {})
        if not zctas:
            print(f"  {state}: no ZCTAs matched -- nothing written")
            continue
        payload = {
            "version": 1,
            "state": state,
            "source": f"US Census Cartographic Boundary File {SHP_STEM} "
                      f"(1:500,000), Douglas-Peucker simplified at "
                      f"{SIMPLIFY_TOLERANCE} degrees",
            "tolerance_degrees": SIMPLIFY_TOLERANCE,
            "zctas": zctas,
        }
        path = OUT_DIR / f"zcta_{state}.json.gz"
        blob = gzip.compress(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
        path.write_bytes(blob)
        total += len(blob)
        print(f"  {state}: {len(zctas):>5} ZCTAs -> {path.name} ({len(blob) / 1024:.0f} KiB)")
    print(f"  total written this run: {total / 1024:.0f} KiB")
    return by_state


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--states", nargs="+", default=None,
                        help=f"Two-letter states to build (default: {' '.join(DEFAULT_STATES)}).")
    parser.add_argument("--add", nargs="+", default=None,
                        help="Build these IN ADDITION to whatever is already on disk.")
    parser.add_argument("--list", action="store_true",
                        help="Show which states are built and their sizes, then exit.")
    args = parser.parse_args()

    if args.list:
        if not OUT_DIR.exists():
            print(f"{OUT_DIR} does not exist -- nothing built yet.")
            return 0
        files = sorted(OUT_DIR.glob("zcta_*.json.gz"))
        if not files:
            print("No ZCTA state files built yet.")
            return 0
        for path in files:
            print(f"  {_state_of(path)}  {path.stat().st_size / 1024:>7.0f} KiB")
        print(f"  total {sum(p.stat().st_size for p in files) / 1024 / 1024:.2f} MiB")
        return 0

    if args.add:
        existing = {p.stem.split("_")[1] for p in OUT_DIR.glob("zcta_*.json.gz")} \
            if OUT_DIR.exists() else set()
        states = sorted(existing | {s.upper() for s in args.add})
        print(f"== ZCTA boundaries: adding {', '.join(s.upper() for s in args.add)} "
              f"to {', '.join(sorted(existing)) or 'nothing'} ==")
    else:
        states = args.states or DEFAULT_STATES
        print(f"== ZCTA boundaries: {', '.join(s.upper() for s in states)} ==")
    build(states)
    return 0


if __name__ == "__main__":
    sys.exit(main())
