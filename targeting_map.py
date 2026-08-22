"""targeting_map.py -- draws targeting groups' resolved zips on one map,
each group in its own color, with a legend. Roadmap section E: "draws what
already exists" -- a group's `resolved_zips`/`resolved_markets`/`color` are
populated by the geo-definition expander (section C/D) before this module
ever sees them; nothing here resolves geography of its own.

Pure raster drawing with Pillow (already a pinned dependency) -- no new
package, no paid service, and nothing platform-specific, so it renders
identically on Streamlit Cloud's Linux and on a Windows dev box. Several
things worth knowing before touching this:

- The crosswalk (`geo_resolver`'s `zip_points`) has zip CENTROIDS only,
  never boundary polygons, so a group's zips are drawn as dots. But its
  COUNTIES are real polygons (`map_boundaries.json.gz`), and a county
  containing a targeted zip is filled with a light tint of the group's
  color -- that's what makes this read as a shaded regional map rather
  than dots scattered on an outline. The dots stay too, since a county fill
  is coarser than the zip list itself (not every zip in a filled county is
  necessarily targeted).
- The projection auto-fits to whatever is actually being plotted (never a
  fixed CONUS frame) -- to the union of the plotted points AND the filled
  counties' own extent, so a filled county is never clipped at the frame
  edge, plus a MARGIN_FRACTION margin. It allows a bounded stretch away
  from true distance-scale (see `_MAX_STRETCH`) so a point cluster whose
  aspect ratio doesn't match the canvas doesn't letterbox into large empty
  margins.
- Place labels (largest cities in frame) come from `map_places.json.gz`
  (`build_map_places.py` -- Census Gazetteer place names/centroids joined
  to Population Estimates Program figures, since the Gazetteer alone
  carries no population field). Capped at `MAX_PLACE_LABELS`, chosen by
  population, and an overlap is suppressed rather than crowded -- a place
  that would collide with an already-placed label or marker is dropped,
  not shrunk or repositioned. Missing the file degrades to no labels
  rather than raising, same discipline as the boundary file.
- **Dark mode** (`dark=True`) renders onto a transparent canvas with light
  boundary/label colors, for compositing over the deck slide's own dark
  gradient background (see `assembly.place_targeting_map`) once the map
  gets a slide of its own with no stock photo behind it. Light mode
  (`dark=False`, the default) renders an opaque canvas, as before. Both
  modes share every other drawing decision (fills, labels, framing) --
  only the palette and canvas opacity differ.
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
# On each side, beyond the fitted extent (plotted points + filled counties)
# -- so nothing is ever drawn flush against the frame edge. 10%, not the
# previous 8%: a filled county's own boundary sits further out than its
# zips' centroids, and the margin has to clear THAT, not just the dots.
MARGIN_FRACTION = 0.10
# How far the projection may deviate from true distance-scale (1.0) on
# either axis to fill more of the frame -- 1.2 means neither axis is ever
# stretched more than 20% relative to the other before falling back to
# letterboxing the rest. A local slide inset doesn't need survey accuracy;
# a wall of white space is the worse failure of the two.
_MAX_STRETCH = 1.2
DEFAULT_COLOR = tg.GROUP_COLORS[0]

BOUNDARIES_PATH = Path(__file__).resolve().parent / "map_boundaries.json.gz"
PLACES_PATH = Path(__file__).resolve().parent / "map_places.json.gz"
# How far beyond the fitted extent (as a fraction of its span) to pull in
# surrounding counties/states for CONTEXT outlines -- generous, because the
# point is to show what's around the target area, not merely what it
# directly overlaps. Independent of MARGIN_FRACTION, which bounds the
# visible frame; this only bounds which outline features get fetched.
BOUNDARY_SEARCH_FRACTION = 0.6
_MIN_SEARCH_DEGREES = 0.35  # a floor for a near-point group (e.g. one zip)

MAX_PLACE_LABELS = 8
# A label's marker must be at least this many pixels from every
# already-placed marker, and its text box must not overlap any
# already-placed text box -- "suppress rather than crowd."
_LABEL_MIN_MARKER_GAP = 14
_LABEL_PADDING = 3

# Kept as top-level names too (not just palette entries): existing callers
# and tests reference these directly as the light mode's outline colors.
COUNTY_OUTLINE_COLOR = (214, 214, 214)
STATE_OUTLINE_COLOR = (140, 140, 140)

_LIGHT_PALETTE = {
    "county_outline": COUNTY_OUTLINE_COLOR + (255,),
    "state_outline": STATE_OUTLINE_COLOR + (255,),
    "county_fill_alpha": 70,
    "legend_title": (40, 40, 40, 255),
    "legend_text": (20, 20, 20, 255),
    "place_marker": (90, 90, 90, 255),
    "place_text": (35, 35, 35, 255),
    "place_text_halo": None,
    # A thin white break between a dot and whatever color sits directly
    # under it -- mild in light mode (fills are already a pale 70/255
    # tint there, so a dot rarely disappears), load-bearing in dark mode.
    "dot_outline": (255, 255, 255),
}
_DARK_PALETTE = {
    "county_outline": (150, 160, 185, 200),
    "state_outline": (215, 220, 235, 230),
    "legend_title": (225, 228, 238, 255),
    "legend_text": (210, 214, 226, 255),
    "place_marker": (225, 228, 238, 255),
    "place_text": (235, 238, 245, 255),
    # A dark halo behind light label text -- the slide's own background is a
    # gradient (navy fading toward black), and text with no halo reads fine
    # over the darkest part of it and washes out over the lighter part.
    "place_text_halo": (15, 18, 30, 190),
    "dot_outline": (255, 255, 255),
}


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


@functools.lru_cache(maxsize=1)
def _places():
    """The place-label payload, or None if it hasn't been built. Same
    silent-degrade discipline as `_boundaries`."""
    if not PLACES_PATH.exists():
        return None
    try:
        with gzip.open(PLACES_PATH, "rb") as f:
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


# A county is filled for a group only when at least this share of the
# county's OWN zips are among that group's targeted zips -- one targeted
# zip in a 25-zip county is not "the county is targeted," and filling the
# whole polygon from it says something the data doesn't. Chosen to catch
# the single-zip case (4% in a real example) while still filling a county
# that's genuinely mostly covered (49-57% in the same example) -- there's
# real room between those two numbers for a threshold to sit in.
_COUNTY_FILL_MIN_COVERAGE = 0.20


def _touched_counties(plottable):
    """[(fips, color), ...] -- every county at least `_COUNTY_FILL_MIN_COVERAGE`
    covered by some group's zips, each credited to whichever group covers
    the LARGEST share of it (not whichever group happened to be listed
    first): two groups can each touch a handful of zips in the same big
    county, and first-come would hand the whole polygon to whichever one
    is earlier in the list even when the other group's real presence there
    is bigger. Coverage is targeted-zips-in-county / real-zips-in-county
    (`geo_resolver.counties_to_zips`), not targeted-zips / targeted-zips-
    in-that-county -- a county the group barely touches must not look
    "fully covered" just because every one of its few zips there happens
    to be targeted.
    """
    data = _boundaries()
    if not data:
        return []

    # Per-group targeted-zip count per county, from each group's own
    # (small) zip list -- cheap, one geo_resolver call per group.
    targeted = {}   # fips -> {group_index: count}
    colors = {}
    for gi, group in enumerate(plottable):
        colors[gi] = group.get("color") or DEFAULT_COLOR
        result = geo_resolver.zips_to_counties(group["resolved_zips"])
        for fips_list in result.resolved.values():
            if not fips_list:
                continue
            fips = fips_list[0]
            if fips not in data.get("counties", {}):
                continue
            targeted.setdefault(fips, {})
            targeted[fips][gi] = targeted[fips].get(gi, 0) + 1
    if not targeted:
        return []

    # The denominator -- real zip count per touched county -- in ONE pass
    # over the crosswalk, not one `counties_to_zips` call per county (which
    # itself rebuilds a full zip->county reverse index every time it's
    # called; calling it per county made this quadratic in the crosswalk's
    # ~34,000 zips and turned a handful of touched counties into a
    # multi-minute render). Matches counties_to_zips' own rule for what
    # counts as "in" a county: every county a zip's crosswalk entry names,
    # not just its first/largest one.
    wanted = set(targeted)
    totals = {fips: 0 for fips in wanted}
    for fips_list in geo_resolver._data()["zip_counties"].values():
        for fips in fips_list:
            if fips in totals:
                totals[fips] += 1

    claimed = []
    for fips, by_group in targeted.items():
        total = totals.get(fips, 0)
        if total <= 0:
            continue
        best_gi, best_count = max(by_group.items(), key=lambda kv: kv[1])
        if (best_count / total) >= _COUNTY_FILL_MIN_COVERAGE:
            claimed.append((fips, colors[best_gi]))
    return claimed


def render_map(groups, width_px=900, height_px=560, background=(255, 255, 255), dark=False,
               label_for=None):
    """A map PNG (bytes) of every group's resolved zips, each in the
    group's own color, with filled counties, place labels, a legend -- or
    None when there's nothing to draw. None is the caller's whole signal:
    leave whatever's already there (the stock image, on the deck slide)
    alone. Never raises for missing/bad data -- a zip absent from the
    crosswalk (there since normalize_zip already validates on the way in,
    but never assumed) is simply not plotted rather than failing the whole
    map, and a missing boundaries/places file degrades to a plainer map
    rather than an error.

    `dark=True` renders onto a transparent canvas with a light palette, for
    compositing over the deck's own dark background; `background` is
    ignored in that case. `dark=False` (the default) renders the opaque
    `background` colour, as before.

    The legend is labeled by `tg.geo_label` -- the group's OWN label, the
    same one the D2 avails grid's Geo cell and Label column already use --
    never `tg.audience_label`. Two groups sharing one audience over
    different geography (an avails document listing the same audience
    against a 52-zip cluster and a separate 34-zip cluster, say) rendered
    IDENTICAL legend text under audience_label -- two colors on the map with
    no way to tell which is which. A map is a geographic view; what
    distinguishes one colored cluster from another on it is where it is,
    which is exactly what geo_label already exists to say (a rep's own
    "Label (optional)" wins first, per geo kind after that). `label_for`
    threads through to geo_label's own market-key -> display-name lookup
    (`app._market_display_name`) so a markets-kind group's legend entry
    reads "Washington, D.C." rather than the raw key; omitted, it falls
    back to the raw key exactly as geo_label itself does.
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
        label = tg.geo_label(group, label_for=label_for) or "(untitled)"
        series.append((group.get("color") or DEFAULT_COLOR, coords, label))
        all_lat.extend(c[0] for c in coords)
        all_lon.extend(c[1] for c in coords)
    if not series:
        return None

    palette = _DARK_PALETTE if dark else _LIGHT_PALETTE
    boundary_data = _boundaries()
    county_fills = _touched_counties(plottable)

    # The fitted extent is the plotted points UNION the filled counties'
    # own bboxes (lon/lat) -- a filled county reaches further than its
    # zips' centroids, and framing on the points alone would clip it.
    fit_lats, fit_lons = list(all_lat), list(all_lon)
    for fips, _color in county_fills:
        bbox = boundary_data["counties"][fips]["bbox"]  # [minlon, minlat, maxlon, maxlat]
        fit_lons.extend([bbox[0], bbox[2]])
        fit_lats.extend([bbox[1], bbox[3]])

    legend_w = _legend_width(series)
    map_w = max(200, width_px - legend_w)

    mode = "RGBA" if dark else "RGB"
    canvas_bg = (0, 0, 0, 0) if dark else background
    img = Image.new(mode, (width_px, height_px), canvas_bg)
    draw = ImageDraw.Draw(img)

    project, frame_bounds = _projector(fit_lats, fit_lons, map_w, height_px)

    search_box = (
        min(fit_lons) - max(_MIN_SEARCH_DEGREES, (max(fit_lons) - min(fit_lons)) * BOUNDARY_SEARCH_FRACTION),
        min(fit_lats) - max(_MIN_SEARCH_DEGREES, (max(fit_lats) - min(fit_lats)) * BOUNDARY_SEARCH_FRACTION),
        max(fit_lons) + max(_MIN_SEARCH_DEGREES, (max(fit_lons) - min(fit_lons)) * BOUNDARY_SEARCH_FRACTION),
        max(fit_lats) + max(_MIN_SEARCH_DEGREES, (max(fit_lats) - min(fit_lats)) * BOUNDARY_SEARCH_FRACTION),
    )
    counties, states = _matching_boundaries(search_box)

    # County fills sit UNDER everything else (outlines, dots, labels), and
    # need real alpha compositing -- drawn on a separate transparent layer,
    # then composited onto the canvas, so overlapping fills (rare, but a
    # county can be adjacent to two groups' territory) blend rather than
    # one flatly overwriting the other's alpha.
    if county_fills:
        fill_layer = Image.new("RGBA", (width_px, height_px), (0, 0, 0, 0))
        fill_draw = ImageDraw.Draw(fill_layer)
        for fips, color in county_fills:
            feature = boundary_data["counties"].get(fips)
            if not feature:
                continue
            rgb = _hex_to_rgb(color)
            if dark:
                rgb, alpha = _dark_fill_style(rgb)
            else:
                alpha = palette["county_fill_alpha"]
            for ring in feature["rings"]:
                if len(ring) < 3:
                    continue
                pixels = [project(lat, lon) for lon, lat in ring]
                fill_draw.polygon(pixels, fill=rgb + (alpha,))
        base_rgba = img.convert("RGBA") if not dark else img
        img = Image.alpha_composite(base_rgba, fill_layer)
        draw = ImageDraw.Draw(img)

    for feature in counties:
        _draw_rings(draw, project, feature["rings"], palette["county_outline"][:3], width=1)
    for feature in states:
        _draw_rings(draw, project, feature["rings"], palette["state_outline"][:3], width=2)

    # A dot drawn in its own group color can sit right on top of that same
    # group's own (lighter, alpha-blended) county fill and all but
    # disappear into it -- a light outline gives every dot a break from
    # whatever's underneath, fill or no fill, rather than trying to detect
    # per-dot whether it happens to fall inside a filled polygon.
    dot_outline = palette.get("dot_outline")
    for color, coords, _label in series:
        rgb = _hex_to_rgb(color)
        for lat, lon in coords:
            x, y = project(lat, lon)
            draw.ellipse([x - DOT_RADIUS, y - DOT_RADIUS, x + DOT_RADIUS, y + DOT_RADIUS],
                        fill=rgb, outline=dot_outline, width=1 if dot_outline else 0)

    _draw_place_labels(draw, project, frame_bounds, map_w, height_px, palette)
    _draw_legend(draw, series, map_w, height_px, palette)

    if dark and img.mode != "RGBA":
        img = img.convert("RGBA")
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
    """((lat, lon) -> (x, y) pixel, (min_lat, max_lat, min_lon, max_lon) of
    the visible frame AFTER the margin, in that order).

    Equirectangular with a cosine-latitude correction (so east-west
    distance reads at roughly the right scale at the plotted points' own
    latitude), fit to the given lat/lon extent plus MARGIN_FRACTION. Scale
    is allowed to differ between the two axes by up to `_MAX_STRETCH` to
    fill more of the frame -- a point cluster whose aspect ratio doesn't
    match the canvas's would otherwise letterbox into large empty margins
    on one axis, all in the name of an exactness a slide inset doesn't
    need. Beyond that bound it falls back to centering the slack, same as
    before. The frame bounds are returned so the caller can filter place
    labels to what's actually visible, not just what was fed in.
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

    # Invert the frame's own pixel-space corners back to lat/lon, rather
    # than re-deriving the margin math a second time -- this is exactly
    # what's visible in [0, width] x [0, height] under `project`.
    frame_min_lon = (0 - offset_x) / scale_x / cos_lat + min_x / cos_lat
    frame_max_lon = (width - offset_x) / scale_x / cos_lat + min_x / cos_lat
    frame_min_lat = -((height - offset_y) / scale_y + min_y)
    frame_max_lat = -((0 - offset_y) / scale_y + min_y)
    frame_bounds = (frame_min_lat, frame_max_lat, frame_min_lon, frame_max_lon)

    return project, frame_bounds


def _hex_to_rgb(hex_color):
    text = str(hex_color or "").lstrip("#")
    if len(text) != 6:
        return _hex_to_rgb(DEFAULT_COLOR)
    try:
        return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return _hex_to_rgb(DEFAULT_COLOR)


def _lighten(rgb, fraction):
    return tuple(round(c + (255 - c) * fraction) for c in rgb)


def _luminance(rgb):
    """0..1 perceived brightness (Rec. 709 luma) -- good enough to rank
    GROUP_COLORS' ten fixed swatches against a background, not meant as a
    real colorimetric model."""
    r, g, b = (c / 255 for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


# How far to mix a dark-mode county fill toward white before blending it
# onto the canvas -- flat across every color, unlike the alpha below.
# Lightening alone doesn't balance visual weight (a lightened orange is
# still much brighter than a lightened navy), it just keeps the darkest
# part of the gradient from turning a fill to mud; the alpha is what
# equalizes how "loud" each color reads.
_DARK_FILL_LIGHTEN = 0.30

# The luminance a representative point on the slide's own background sits
# at -- the gradient runs from accent1@50%lumMod (~L 0.09) to black (L 0),
# so this is the darker end where a fill has to hold up, not the average.
_DARK_BG_LUMINANCE = 0.06

# How much brighter, in luminance, a fill should read than the background
# it sits on -- held CONSTANT across every group color, which is the whole
# fix: navy's own luminance sits close to the background's (it "recedes"
# at a flat alpha), orange's sits far above it (it "advances"), so hitting
# the same target delta gives navy a higher alpha and orange a lower one
# instead of leaving both at one number and letting hue distance alone
# decide how loud each group's fill reads.
_DARK_FILL_TARGET_DELTA = 0.14
_DARK_FILL_ALPHA_MIN = 55
_DARK_FILL_ALPHA_MAX = 130


def _dark_fill_style(rgb):
    """(lightened_rgb, alpha) for a dark-mode county fill -- see the three
    constants above for what each one is doing and why."""
    lightened = _lighten(rgb, _DARK_FILL_LIGHTEN)
    color_luminance = _luminance(lightened)
    gap = max(color_luminance - _DARK_BG_LUMINANCE, 0.05)
    alpha = round((_DARK_FILL_TARGET_DELTA / gap) * 255)
    alpha = max(_DARK_FILL_ALPHA_MIN, min(_DARK_FILL_ALPHA_MAX, alpha))
    return lightened, alpha


def _legend_font():
    # Pillow's own bundled font -- no filesystem lookup, so nothing to
    # differ between a Windows dev box and Streamlit Cloud's Linux the way
    # text_metrics.py's brand-font resolution deliberately can. A map
    # legend doesn't need to match Calibri/Proxima Nova; it needs to always
    # be there.
    return ImageFont.load_default(size=13)


def _place_font():
    return ImageFont.load_default(size=11)


def _legend_width(series):
    font = _legend_font()
    widest = 0
    for _color, _coords, label in series:
        bbox = font.getbbox(label)
        widest = max(widest, bbox[2] - bbox[0])
    return min(LEGEND_MAX_WIDTH, max(LEGEND_MIN_WIDTH, widest + LEGEND_SWATCH + 3 * LEGEND_PADDING))


def _draw_legend(draw, series, x0, height, palette):
    font = _legend_font()
    y = LEGEND_PADDING
    draw.text((x0 + LEGEND_PADDING, y), "Targeting groups", font=font, fill=palette["legend_title"])
    y += LEGEND_ROW_HEIGHT
    for color, _coords, label in series:
        rgb = _hex_to_rgb(color)
        draw.rectangle(
            [x0 + LEGEND_PADDING, y + 3, x0 + LEGEND_PADDING + LEGEND_SWATCH, y + 3 + LEGEND_SWATCH],
            fill=rgb)
        draw.text((x0 + LEGEND_PADDING + LEGEND_SWATCH + 6, y), label, font=font, fill=palette["legend_text"])
        y += LEGEND_ROW_HEIGHT
        if y > height - LEGEND_PADDING:
            break  # more groups than the legend has room for -- stop rather than overflow the frame


def _draw_place_labels(draw, project, frame_bounds, map_w, height, palette):
    """The largest places actually in frame, up to MAX_PLACE_LABELS,
    overlaps suppressed rather than crowded: a place whose marker sits too
    close to an already-placed one, or whose text box would overlap an
    already-placed text box, is skipped outright -- not shrunk, not
    repositioned. An unreadable label is worse than one fewer label.
    """
    data = _places()
    if not data:
        return
    min_lat, max_lat, min_lon, max_lon = frame_bounds
    candidates = [p for p in data.get("places", [])
                 if min_lat <= p["lat"] <= max_lat and min_lon <= p["lon"] <= max_lon]
    candidates.sort(key=lambda p: -p["pop"])

    font = _place_font()
    marker_color = palette["place_marker"]
    text_color = palette["place_text"]
    halo = palette["place_text_halo"]

    placed_markers = []  # [(x, y), ...]
    placed_boxes = []    # [(x0, y0, x1, y1), ...]
    placed = 0
    for place in candidates:
        if placed >= MAX_PLACE_LABELS:
            break
        x, y = project(place["lat"], place["lon"])
        if not (0 <= x <= map_w and 0 <= y <= height):
            continue
        if any(math.hypot(x - px, y - py) < _LABEL_MIN_MARKER_GAP for px, py in placed_markers):
            continue
        text = place["name"]
        bbox = font.getbbox(text)
        text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        tx, ty = x + 5, y - text_h - 2
        box = (tx - _LABEL_PADDING, ty - _LABEL_PADDING,
              tx + text_w + _LABEL_PADDING, ty + text_h + _LABEL_PADDING)
        if box[2] > map_w or box[0] < 0 or box[1] < 0 or box[3] > height:
            continue
        if any(_boxes_overlap(box, other) for other in placed_boxes):
            continue

        draw.ellipse([x - 2, y - 2, x + 2, y + 2], fill=marker_color)
        if halo:
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    if dx or dy:
                        draw.text((tx + dx, ty + dy), text, font=font, fill=halo)
        draw.text((tx, ty), text, font=font, fill=text_color)

        placed_markers.append((x, y))
        placed_boxes.append(box)
        placed += 1


def _boxes_overlap(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return ax0 < bx1 and ax1 > bx0 and ay0 < by1 and ay1 > by0
