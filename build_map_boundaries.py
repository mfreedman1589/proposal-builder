"""Build the state/county outline data the targeting map draws for context
(roadmap §E -- "colored dots on white don't read as a map, a client can't
tell it's St. Louis").

    python build_map_boundaries.py [--out map_boundaries.json.gz]

Census Cartographic Boundary Files (cb_2022_us_state_20m / cb_2022_us_county_20m,
20m = 1:20,000,000 generalization -- coarse enough to stay small, plenty fine
for a slide-sized inset map), public domain as works of the US Government,
same provenance class as geo_crosswalk.json.gz's own Census sources.

No shapefile library, deliberately. The ESRI shapefile binary format has been
stable for decades and is simple enough to read directly with stdlib struct;
adding pyshp (or worse, GDAL/fiona) for a one-time build script would be a
dependency this project doesn't otherwise carry, at build time or runtime.
Only Polygon-type (shape type 5) records are handled, which is what these two
files contain -- anything else raises rather than silently producing an empty
outline.

Downloads cache under ../proposal-builder-assets/geo, same as
build_geo_crosswalk.py and build_market_lookup.py. Rerun to refresh; output is
deterministic for a given pair of input files.
"""
import gzip
import json
import struct
import sys
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
CACHE = HERE.parent / "proposal-builder-assets" / "geo"
DEFAULT_OUT = HERE / "map_boundaries.json.gz"

SOURCES = {
    "state": {
        "url": "https://www2.census.gov/geo/tiger/GENZ2022/shp/cb_2022_us_state_20m.zip",
        "file": "cb_2022_us_state_20m.zip",
        "shp_stem": "cb_2022_us_state_20m",
    },
    "county": {
        "url": "https://www2.census.gov/geo/tiger/GENZ2022/shp/cb_2022_us_county_20m.zip",
        "file": "cb_2022_us_county_20m.zip",
        "shp_stem": "cb_2022_us_county_20m",
    },
}

# ~11m at the equator -- well under the 20m generalization already baked into
# the source, so this only trims JSON text size, never visible detail.
ROUND_DP = 4


def _fetch(spec):
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / spec["file"]
    if not path.exists():
        print(f"  downloading {spec['url']}", file=sys.stderr)
        urllib.request.urlretrieve(spec["url"], path)
    return path


def _read_dbf(data):
    """Yields one {field_name: value} dict per record, in file order.

    Deleted records (leading '*') are yielded too, not skipped -- .shp and
    .dbf records are matched purely by position, so dropping one here would
    desync every record after it against the .shp file's own sequence.
    """
    header_size, record_size = struct.unpack_from("<HH", data, 8)
    n_fields = (header_size - 32 - 1) // 32
    fields = []
    for i in range(n_fields):
        off = 32 + i * 32
        raw = data[off:off + 32]
        name = raw[:11].split(b"\x00")[0].decode("ascii")
        length = raw[16]
        fields.append((name, length))
    pos = header_size
    while pos < len(data) and data[pos:pos + 1] not in (b"\x1a", b""):
        rec = data[pos:pos + record_size]
        pos += record_size
        values = {}
        off = 1  # skip the deletion flag byte
        for name, length in fields:
            values[name] = rec[off:off + length].decode("ascii", "replace").strip()
            off += length
        yield values


def _read_shp_polygons(data):
    """Yields (parts, points) per record, positional -- see _read_dbf's note.

    points is a flat list of (x, y) in file order; parts is the point-index
    each ring starts at, exactly as the shapefile spec lays them out.
    """
    pos = 100  # fixed-size main file header
    while pos < len(data):
        rec_num, content_words = struct.unpack_from(">ii", data, pos)
        pos += 8
        content = data[pos:pos + content_words * 2]
        pos += content_words * 2
        shape_type = struct.unpack_from("<i", content, 0)[0]
        if shape_type == 0:  # null shape -- a handful exist upstream
            yield [], []
            continue
        if shape_type != 5:
            raise ValueError(f"record {rec_num}: expected Polygon (5), got shape type {shape_type}")
        num_parts, num_points = struct.unpack_from("<ii", content, 36)
        parts = list(struct.unpack_from(f"<{num_parts}i", content, 44))
        pts_off = 44 + num_parts * 4
        coords = struct.unpack_from(f"<{num_points * 2}d", content, pts_off)
        points = list(zip(coords[0::2], coords[1::2]))
        yield parts, points


def _rings(parts, points):
    bounds = parts + [len(points)]
    return [points[bounds[i]:bounds[i + 1]] for i in range(len(parts))]


def _feature_bbox(rings):
    xs = [x for ring in rings for x, _ in ring]
    ys = [y for ring in rings for _, y in ring]
    return [round(min(xs), ROUND_DP), round(min(ys), ROUND_DP),
            round(max(xs), ROUND_DP), round(max(ys), ROUND_DP)]


def _load_layer(spec):
    zpath = _fetch(spec)
    with zipfile.ZipFile(zpath) as zf:
        shp = zf.read(f"{spec['shp_stem']}.shp")
        dbf = zf.read(f"{spec['shp_stem']}.dbf")
    records = list(_read_dbf(dbf))
    shapes = list(_read_shp_polygons(shp))
    if len(records) != len(shapes):
        raise ValueError(f"{spec['shp_stem']}: {len(records)} dbf records vs {len(shapes)} shp records")
    return records, shapes


def _feature(parts, points):
    rings_raw = _rings(parts, points)
    rings = [[[round(x, ROUND_DP), round(y, ROUND_DP)] for x, y in ring] for ring in rings_raw]
    return rings, _feature_bbox(rings_raw)


def build(out_path=DEFAULT_OUT):
    states = {}
    for rec, (parts, points) in zip(*_load_layer(SOURCES["state"])):
        if not points:
            continue
        rings, bbox = _feature(parts, points)
        states[rec["STATEFP"]] = {"name": rec["NAME"], "rings": rings, "bbox": bbox}

    counties = {}
    for rec, (parts, points) in zip(*_load_layer(SOURCES["county"])):
        if not points:
            continue
        rings, bbox = _feature(parts, points)
        counties[rec["GEOID"]] = {"name": rec["NAME"], "state_fp": rec["STATEFP"],
                                  "rings": rings, "bbox": bbox}

    payload = {
        "version": 1,
        "source": "US Census Bureau Cartographic Boundary Files, cb_2022_us_{state,county}_20m "
                  "(1:20,000,000 generalized), public domain.",
        "states": states,
        "counties": counties,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with gzip.open(out_path, "wt", encoding="utf-8") as f:
        json.dump(payload, f, separators=(",", ":"))

    print(f"{len(states)} states, {len(counties)} counties")
    print(f"wrote {out_path} ({out_path.stat().st_size:,} bytes)")
    return payload


def main():
    out = Path(sys.argv[sys.argv.index("--out") + 1]) if "--out" in sys.argv else DEFAULT_OUT
    build(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
