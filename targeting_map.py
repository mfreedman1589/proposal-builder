"""targeting_map.py -- draws targeting groups' resolved zips on one map,
each group in its own color, with a legend. Roadmap §E: "draws what already
exists" -- a group's `resolved_zips`/`resolved_markets`/`color` are
populated by the geo-definition expander (§C/D) before this module ever
sees them; nothing here resolves geography of its own.

Pure raster drawing with Pillow (already a pinned dependency) -- no new
package, no paid service, and nothing platform-specific, so it renders
identically on Streamlit Cloud's Linux and on a Windows dev box. Two things
worth knowing before touching this:

- The crosswalk (`geo_resolver`'s `zip_points`) has zip CENTROIDS only,
  never boundary polygons. So a group's "region" is drawn as colored dots
  at each zip's centroid, not a filled area. A convex hull would shade
  territory nobody targeted for anything but a single contiguous zip list,
  which is the uncommon case -- dishonest for the ordinary one. Geography
  reads through the natural clustering of the points themselves, now with
  real state/county outlines underneath for context (`map_boundaries.json.gz`,
  built by `build_map_boundaries.py` from public-domain Census cartographic
  boundary files -- see that script's docstring). Missing the file degrades
  to dots-on-white exactly as before rather than raising, since the whole
  map feature is optional.
- The projection auto-fits to whatever is actually being plotted (never a
  fixed CONUS frame), so a single-region group still fills the frame
  instead of sitting as a speck on an empty continent. It allows a bounded
  stretch away from true distance-scale (see `_MAX_STRETCH`) so a point
  cluster whose aspect ratio doesn't match the canvas doesn't letterbox
  into large empty margins -- the boundary outlines fill most of that
  space now anyway, but the stretch bound keeps what's left honest.
"""
import functools
import gzip
import json
import math
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import geo_resolver
import targeting_groups as tg

DOT_RADIUS = 3
LEGEND_SWATCH = 12
LEGEND_ROW_HEIGHT = 20
LEGEND_PADDING = 12
LEGEND_MIN_WIDTH = 140
LEGEND_MAX_WIDTH = 320
# On each side, beyond the plotted points' own bounding box -- so a point
# is never drawn flush against the frame edge.
MARGIN_FRACTION = 0.08
# How far the projection may deviate from true distance-scale (1.0) on
# either axis to fill more of the frame -- 1.2 means neither axis is ever
# stretched more than 20% relative to the other before falling back to
# letterboxing the rest. A local slide inset doesn't need survey accuracy;
# a wall of white space is the worse failure of the two.
_MAX_STRETCH = 1.2
DEFAULT_COLOR = tg.GROUP_COLORS[0]

BOUNDARIES_PATH = Path(__file__).resolve().parent / "map_boundaries.json.gz"
COUNTY_OUTLINE_COLOR = (214, 214, 214)
STATE_OUTLINE_COLOR = (140, 140, 140)
# How far beyond the plotted points' own bounding box (as a fraction of its
# span) to pull in surrounding counties/states for context -- generous,
# because the point is to show what's around the target area, not merely
# what it directly overlaps.
BOUNDARY_SEARCH_FRACTION = 0.6
_MIN_SEARCH_DEGREES = 0.35  # a floor for a near-point group (e.g. one zip)


@functools.lru_cache(maxsize=1)
def _boundaries():
    """The state/county outline payload, or None if it hasn't been built.

    Optional and silent on failure, unlike geo_resolver's crosswalk -- the
    whole map feature already degrades gracefully with nothing resolved, so
    a missing basemap file degrades the same way: dots on white, as before
    `build_map_boundaries.py` existed, rather than an error.
    """
    if not BOUNDARIES_PATH.exists():
        return None
    try:
        with gzip.open(BOUNDARIES_PATH, "rb") as f:
            return json.loads(f.read().decode("utf-8"))
    except Exception:                                                  # noqa: BLE001
        return None


def groups_with_zips(groups):
    """Groups worth putting on the map -- resolved, with at least one zip.
    A group still waiting on a Resolve click, or one whose real resolution
    came back empty (a radius matching nothing), contributes nothing to
    draw and is silently skipped, not an error -- the whole feature is
    optional, and "nothing to draw" is the ordinary case for a proposal
    that hasn't done any zip work.
    """
    return [g for g in (groups or []) if g.get("resolved_zips")]


def render_map(groups, width_px=900, height_px=560, background=(255, 255, 255)):
    """A map PNG (bytes) of every group's resolved zips, each in the
    group's own color, with a legend naming each one -- or None when
    there's nothing to draw. None is the caller's whole signal: leave
    whatever's already there (the stock image, on the deck slide) alone.
    Never raises for missing/bad data -- a zip absent from the crosswalk
    (there since normalize_zip already validates on the way in, but never
    assumed) is simply not plotted rather than failing the whole map.
    """
    plottable = groups_with_zips(groups)
    if not plottable:
        return None

    points = geo_resolver._data()["zip_points"]
    series = []  # (color, [(lat, lon), ...], label)
    all_lat, all_lon = [], []
    for group in plottable:
        coords = [tuple(points[z]) for z in group["resolved_zips"] if z in points]
        if not coords:
            continue
        label = tg.audience_label(group) or "(untitled)"
        series.append((group.get("color") or DEFAULT_COLOR, coords, label))
        all_lat.extend(c[0] for c in coords)
        all_lon.extend(c[1] for c in coords)
    if not series:
        return None

    legend_w = _legend_width(series)
    map_w = max(200, width_px - legend_w)

    img = Image.new("RGB", (width_px, height_px), background)
    draw = ImageDraw.Draw(img)

    project = _projector(all_lat, all_lon, map_w, height_px)

    search_box = (
        min(all_lon) - max(_MIN_SEARCH_DEGREES, (max(all_lon) - min(all_lon)) * BOUNDARY_SEARCH_FRACTION),
        min(all_lat) - max(_MIN_SEARCH_DEGREES, (max(all_lat) - min(all_lat)) * BOUNDARY_SEARCH_FRACTION),
        max(all_lon) + max(_MIN_SEARCH_DEGREES, (max(all_lon) - min(all_lon)) * BOUNDARY_SEARCH_FRACTION),
        max(all_lat) + max(_MIN_SEARCH_DEGREES, (max(all_lat) - min(all_lat)) * BOUNDARY_SEARCH_FRACTION),
    )
    counties, states = _matching_boundaries(search_box)
    for feature in counties:
        _draw_rings(draw, project, feature["rings"], COUNTY_OUTLINE_COLOR, width=1)
    for feature in states:
        _draw_rings(draw, project, feature["rings"], STATE_OUTLINE_COLOR, width=2)

    for color, coords, _label in series:
        rgb = _hex_to_rgb(color)
        for lat, lon in coords:
            x, y = project(lat, lon)
            draw.ellipse([x - DOT_RADIUS, y - DOT_RADIUS, x + DOT_RADIUS, y + DOT_RADIUS], fill=rgb)

    _draw_legend(draw, series, map_w, height_px)

    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _bbox_overlaps(bbox, search_box):
    x0, y0, x1, y1 = bbox
    sx0, sy0, sx1, sy1 = search_box
    return x0 <= sx1 and x1 >= sx0 and y0 <= sy1 and y1 >= sy0


def _matching_boundaries(search_box):
    """(counties, states) whose own bbox overlaps search_box -- empty lists
    when map_boundaries.json.gz hasn't been built, never an error."""
    data = _boundaries()
    if not data:
        return [], []
    counties = [f for f in data.get("counties", {}).values() if _bbox_overlaps(f["bbox"], search_box)]
    states = [f for f in data.get("states", {}).values() if _bbox_overlaps(f["bbox"], search_box)]
    return counties, states


def _draw_rings(draw, project, rings, color, width):
    for ring in rings:
        if len(ring) < 2:
            continue
        pixels = [project(lat, lon) for lon, lat in ring]
        draw.line(pixels + [pixels[0]], fill=color, width=width, joint="curve")


def _projector(lats, lons, width, height):
    """(lat, lon) -> (x, y) pixel. Equirectangular with a cosine-latitude
    correction (so east-west distance reads at roughly the right scale at
    the plotted points' own latitude), fit to their bounding box plus a
    margin. Scale is allowed to differ between the two axes by up to
    `_MAX_STRETCH` to fill more of the frame -- a point cluster whose aspect
    ratio doesn't match the canvas's would otherwise letterbox into large
    empty margins on one axis, all in the name of an exactness a slide inset
    doesn't need. Beyond that bound it falls back to centering the slack,
    same as before.
    """
    mean_lat = sum(lats) / len(lats)
    cos_lat = max(0.15, math.cos(math.radians(mean_lat)))
    xs = [lon * cos_lat for lon in lons]
    ys = [-lat for lat in lats]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    span_x = max(max_x - min_x, 1e-6)
    span_y = max(max_y - min_y, 1e-6)
    min_x -= span_x * MARGIN_FRACTION
    max_x += span_x * MARGIN_FRACTION
    min_y -= span_y * MARGIN_FRACTION
    max_y += span_y * MARGIN_FRACTION
    span_x, span_y = max_x - min_x, max_y - min_y

    true_scale = min(width / span_x, height / span_y)
    fill_scale_x = width / span_x
    fill_scale_y = height / span_y
    scale_x = min(fill_scale_x, true_scale * _MAX_STRETCH)
    scale_y = min(fill_scale_y, true_scale * _MAX_STRETCH)
    offset_x = (width - span_x * scale_x) / 2
    offset_y = (height - span_y * scale_y) / 2

    def project(lat, lon):
        x = (lon * cos_lat - min_x) * scale_x + offset_x
        y = (-lat - min_y) * scale_y + offset_y
        return x, y

    return project


def _hex_to_rgb(hex_color):
    text = str(hex_color or "").lstrip("#")
    if len(text) != 6:
        return _hex_to_rgb(DEFAULT_COLOR)
    try:
        return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return _hex_to_rgb(DEFAULT_COLOR)


def _legend_font():
    # Pillow's own bundled font -- no filesystem lookup, so nothing to
    # differ between a Windows dev box and Streamlit Cloud's Linux the way
    # text_metrics.py's brand-font resolution deliberately can. A map
    # legend doesn't need to match Calibri/Proxima Nova; it needs to always
    # be there.
    return ImageFont.load_default(size=13)


def _legend_width(series):
    font = _legend_font()
    widest = 0
    for _color, _coords, label in series:
        bbox = font.getbbox(label)
        widest = max(widest, bbox[2] - bbox[0])
    return min(LEGEND_MAX_WIDTH, max(LEGEND_MIN_WIDTH, widest + LEGEND_SWATCH + 3 * LEGEND_PADDING))


def _draw_legend(draw, series, x0, height):
    font = _legend_font()
    y = LEGEND_PADDING
    draw.text((x0 + LEGEND_PADDING, y), "Targeting groups", font=font, fill=(40, 40, 40))
    y += LEGEND_ROW_HEIGHT
    for color, _coords, label in series:
        rgb = _hex_to_rgb(color)
        draw.rectangle(
            [x0 + LEGEND_PADDING, y + 3, x0 + LEGEND_PADDING + LEGEND_SWATCH, y + 3 + LEGEND_SWATCH],
            fill=rgb)
        draw.text((x0 + LEGEND_PADDING + LEGEND_SWATCH + 6, y), label, font=font, fill=(20, 20, 20))
        y += LEGEND_ROW_HEIGHT
        if y > height - LEGEND_PADDING:
            break  # more groups than the legend has room for -- stop rather than overflow the frame
