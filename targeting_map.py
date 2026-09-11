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
import re
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
# The map itself never shrinks below this -- also the hard ceiling every
# legend label is fit against below, so the legend can widen to use real
# canvas space (past LEGEND_MAX_WIDTH, when the canvas has room) without
# ever pushing text past the canvas edge. See render_map's own fitting pass.
_MIN_MAP_WIDTH = 200
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
# A three-or-more-audience overlap can't honestly be a two-color hatch --
# something would have to be dropped. Deliberately outside GROUP_COLORS (the
# closest, BAB0AC, is a warm light taupe; this is a flat neutral grey) so it
# never reads as a specific audience's own color, and cross-hatched (both
# diagonals) rather than the two-way overlap's single diagonal, so "three or
# more" is legible from the fill pattern alone, not just the legend text.
_MULTI_FILL_COLOR = "#5C5C5C"

BOUNDARIES_PATH = Path(__file__).resolve().parent / "map_boundaries.json.gz"
PLACES_PATH = Path(__file__).resolve().parent / "map_places.json.gz"
# ZCTA (zip-area) polygons, ONE FILE PER STATE, loaded lazily and only by the
# choropleth. Deliberately not folded into map_boundaries.json.gz: every
# targeting-map render loads that file, and the proposal builder must not pay
# ~1.65 MiB of zip geometry it never draws. Built by build_zcta_boundaries.py.
ZCTA_DIR = Path(__file__).resolve().parent / "zcta_boundaries"
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


@functools.lru_cache(maxsize=None)
def _zcta_state(state):
    """One state's {zip: {"rings": [...]}} payload, or None when that state
    hasn't been built. Cached per state, so a two-state campaign loads two
    files once each and a national one never loads 48 it doesn't need.

    Same silent-degrade contract as `_boundaries`/`_places`: a missing or
    unreadable file is None, and the caller falls back to centroid dots
    rather than raising. `build_zcta_boundaries.py --add <ST>` is the fix.
    """
    path = ZCTA_DIR / f"zcta_{state.upper()}.json.gz"
    if not path.exists():
        return None
    try:
        with gzip.open(path, "rb") as fh:
            return json.loads(fh.read().decode("utf-8")).get("zctas") or None
    except Exception:                                                  # noqa: BLE001
        return None


def zcta_rings_for(zips):
    """({zip: rings}, [zips with no polygon]) for the given zips.

    Loads only the states those zips actually touch -- resolved through the
    crosswalk's own zip->county->state chain, the same one
    build_zcta_boundaries.py packages by, so a zip is always looked for in
    the file it was written into.
    """
    data = geo_resolver._data()
    counties = data.get("counties") or {}
    states = {}
    for code in zips:
        for fips in data["zip_counties"].get(code, ()):
            state = (counties.get(fips) or {}).get("state")
            if state:
                states.setdefault(state, []).append(code)
                break
        else:
            states.setdefault(None, []).append(code)

    rings, missing = {}, []
    for state, codes in states.items():
        payload = _zcta_state(state) if state else None
        if not payload:
            missing.extend(codes)
            continue
        for code in codes:
            entry = payload.get(code)
            if entry and entry.get("rings"):
                rings[code] = entry["rings"]
            else:
                missing.append(code)
    return rings, missing


def groups_with_zips(groups):
    """Groups worth putting on the map -- resolved, with at least one zip.
    A group still waiting on a Resolve click, or one whose real resolution
    came back empty (a radius matching nothing), contributes nothing to
    draw and is silently skipped, not an error -- the whole feature is
    optional, and "nothing to draw" is the ordinary case for a proposal
    that hasn't done any zip work.
    """
    return [g for g in (groups or []) if g.get("resolved_zips")]


_NO_EXPORT_KINDS = frozenset({"markets", "counties"})


def exportable_zip_groups(groups):
    """Groups worth SHOWING A ZIP LIST FOR -- a narrower question than
    `groups_with_zips`. A market or county group has `resolved_zips`
    populated too (the map needs it -- see `app.resolve_group_geography`),
    but ad ops targets those by DMA/county NAME directly, so the zip list
    has no consumer and is pure noise on the page (a real market can be
    1,000+ zips). Excluded by kind, not allowlisted to "zips" alone: a
    Radius group also resolves to a real zip list, but a radius has no name
    ad ops could target by INSTEAD -- the zip list is the actual
    deliverable there, same as it is for an explicit zip list, so it stays
    exportable. Only the two kinds with a real named alternative are
    dropped.
    """
    return [g for g in (groups or [])
           if g.get("resolved_zips")
           and (g.get("geo_def") or {}).get("kind") not in _NO_EXPORT_KINDS]


def legend_entries(plottable, label_for=None):
    """[(label, color, [group, ...])] -- one entry per audience, EXCEPT an
    audience whose groups don't all share ONE color, which gets one entry
    per distinct color in play.

    This is the map's whole color model. Every group actually painted the
    audience's own reference color -- its first unlocked group's color, the
    one the cascade keeps every following group in sync with -- folds into
    ONE entry labeled by the audience alone, since that's what's true of
    the map: they're drawn identically and read as one thing. A group
    painted some OTHER color (almost always `color_locked` -- a rep's
    deliberate single-row override -- but bucketed by the color itself
    rather than trusting the flag alone, so two groups that simply HAPPEN
    to differ are never silently merged either) gets its own entry, named
    "{audience} ({its own geo_label})" -- the same geo_label the avails
    table's Label column already uses, demoted to a parenthetical since
    the audience is the primary identity a rep recognizes a color by.

    Order is first-appearance by audience, the reference-color entry before
    any distinct ones -- deterministic, and matches how a rep built the
    groups (their audience's own color was always there first).

    **A bare audience label is only unambiguous when EVERY group under it
    shares the reference color.** The moment one is broken out, "Auto
    Intenders" sitting above "Auto Intenders (Denver)" no longer says
    whether Denver is included or not -- this renders on a document a
    client signs, so the reference entry gets qualified with its OWN
    members' markets too, e.g. "Auto Intenders (DC, Richmond)". Still one
    entry for every group on the reference color; only its label changes.
    The full (unfitted) market list is built here -- pixel-width fitting
    against the actual legend column happens later, in `render_map`, the
    only place that knows the real canvas size.
    """
    order, by_audience = [], {}
    for group in plottable:
        aud = tg.audience_label(group) or "(untitled)"
        by_audience.setdefault(aud, []).append(group)
        if aud not in order:
            order.append(aud)

    entries = []
    for aud in order:
        members = by_audience[aud]
        unlocked = [g for g in members if not g.get("color_locked")]
        reference = (unlocked[0] if unlocked else members[0]).get("color") or DEFAULT_COLOR

        shared = [g for g in members if (g.get("color") or DEFAULT_COLOR) == reference]
        distinct = [g for g in members if (g.get("color") or DEFAULT_COLOR) != reference]
        shared_label = aud if not distinct else _qualified_label(aud, shared, label_for)
        entries.append((shared_label, reference, shared))
        for group in distinct:
            label = f"{aud} ({tg.geo_label(group, label_for=label_for)})"
            entries.append((label, group.get("color") or DEFAULT_COLOR, [group]))
    return _collapse_identical_geography(entries, label_for=label_for)


def _collapse_identical_geography(entries, label_for=None):
    """Merge two or more entries whose TARGETED ZIP SETS are EXACTLY
    identical (not merely overlapping) into ONE entry -- a RUNTIME check on
    the actual resolved geography, never an assumption about what an avails
    document contains. Real avails documents on hand tend to sell one
    geography against several audiences (Annapolis: all 4 AUTO Make
    audiences resolve to the exact same 20 zips, same origin/radius;
    Hershey/Wilmington: one geography per audience-major group of rows) --
    but that's evidence about a SAMPLE, not a guarantee, and the mixed
    (non-identical, partially-overlapping) case is exactly as reachable: an
    imported avail can carry one audience across 4 DMAs and a second across
    only 3 of them directly, with no hand-editing involved, and a rep can
    also reach it by hand-building groups, editing a group's markets after
    an import, or combining two avails documents in one proposal. Both
    paths are live; a 3-document sample is not a spec. The cutoff is exact
    set identity (Jaccard == 1.0) because there is no data supporting a
    fuzzy one, and anything short of exact identity keeps the EXISTING
    overlap/multi-hatch behaviour (`_touched_counties`) entirely unchanged
    -- this pass only ever REMOVES entries by merging them, never touches
    `_touched_counties`'s own coverage-threshold logic for what's left.

    Two audiences with identical geography drawn in two different colors is
    not a meaningful distinction for a client reading a map -- the WHOLE
    reason to distinguish them is already sitting in the plan table beside
    it, with impressions per row, which a map can't convey either way. So
    the merged entry does not enumerate the audiences it combines (an
    "Audience A + Audience B + Audience C + Audience D" label was tried and
    measured against a real 4-audience case: 565px against a 272px budget,
    with nothing left to trim that doesn't misrepresent who's targeted) --
    it's labeled by the GEOGRAPHY itself, the thing every merged audience
    actually has in common, using the first (plottable-order) member's own
    `tg.geo_label` -- the same label already shown per-row on the avails
    table and plan, so a rep recognizes it. A merge of ANY size still gets
    exactly one legend entry regardless of how many audiences share it,
    which is what removes the width problem outright rather than trimming
    around it.

    A single-audience map (nothing to compare against) is a guaranteed
    no-op: this only ever iterates equivalence classes of size >= 2.

    An entry with NO resolved zips at all never collapses with another,
    even another empty one -- "no geography data" is not "identical
    geography," and two audiences that both merely lack zips (every
    caller in practice pre-filters to `groups_with_zips`, but this
    function is called directly in tests too) are not the same buy.
    """
    from collections import defaultdict

    groups_of = defaultdict(list)   # zips-frozenset -> [(orig_index, entry), ...]
    for i, entry in enumerate(entries):
        _label, _color, members = entry
        zips = frozenset(z for g in members for z in (g.get("resolved_zips") or []))
        key = zips if zips else ("_empty", i)   # never collapses with another empty one
        groups_of[key].append((i, entry))

    merged = []   # (earliest_orig_index, label, color, members)
    for indexed_entries in groups_of.values():
        if len(indexed_entries) == 1:
            i, (label, color, members) = indexed_entries[0]
            merged.append((i, label, color, members))
            continue
        earliest = min(i for i, _entry in indexed_entries)
        combined_members = [g for _i, (_l, _c, members) in indexed_entries for g in members]
        first_color = min(indexed_entries, key=lambda pair: pair[0])[1][1]
        geography = tg.geo_label(combined_members[0], label_for=label_for)
        label = f"{geography} (all targeted audiences)" if geography else "All targeted audiences"
        merged.append((earliest, label, first_color, combined_members))

    merged.sort(key=lambda item: item[0])
    return [(label, color, members) for _i, label, color, members in merged]


def _qualified_label(aud, members, label_for):
    """"{aud} (m1, m2, ...)" -- every member's own geo_label, flattened to
    atomic market names (a kind:"markets"/resolved-zips geo_label already
    joins several with ", " -- split back out so a later width-fit can drop
    ONE market at a time, not one group's whole multi-market label at
    once). Unfitted; see `legend_entries`.
    """
    names = []
    for member in members:
        label = tg.geo_label(member, label_for=label_for)
        if label:
            names.extend(part.strip() for part in label.split(",") if part.strip())
    if not names:
        return aud
    return f"{aud} ({', '.join(names)})"


# A county is filled for a group only when at least this share of the
# county's OWN zips are among that group's targeted zips -- one targeted
# zip in a 25-zip county is not "the county is targeted," and filling the
# whole polygon from it says something the data doesn't. Chosen to catch
# the single-zip case (4% in a real example) while still filling a county
# that's genuinely mostly covered (49-57% in the same example) -- there's
# real room between those two numbers for a threshold to sit in.
_COUNTY_FILL_MIN_COVERAGE = 0.20


def _touched_counties(plottable, label_for=None):
    """[(fips, fill), ...] -- every county at least `_COUNTY_FILL_MIN_COVERAGE`
    covered by some LEGEND ENTRY's zips (an audience's shared-color groups
    combined, or one broken-out group -- see `legend_entries`), where
    `fill` is `("solid", color)`, `("overlap", (color_a, color_b),
    (label_a, label_b))`, or `("multi", (label_1, label_2, ...))` for three
    or more.

    Coverage is aggregated per ENTRY, not per raw group: two groups still
    sharing one audience's color are the same territory as far as a county
    fill is concerned, and must combine their zip counts rather than
    compete for credit the way two genuinely different audiences do.
    Coverage itself is targeted-zips-in-county / real-zips-in-county
    (`geo_resolver.counties_to_zips`), not targeted-zips / targeted-zips-
    in-that-county -- a county an entry barely touches must not look
    "fully covered" just because every one of its few zips there happens
    to be targeted.

    When exactly one entry clears the threshold, the county is solid in
    that entry's color, as before. When exactly TWO clear it, neither wins
    -- reported as an overlap between the two, for the caller to hatch in
    both real colors. **Three or more is a DIFFERENT fill, not a bigger
    hatch**: a two-color hatch showing only the top two by coverage would
    silently drop whoever's third, asserting (by omission) that only two
    audiences are here when a third genuinely clears the same threshold --
    reported as `("multi", (label_1, label_2, ...))` instead, every
    clearing entry named, for the caller to paint one neutral, unambiguous
    fill rather than pick which real audience to misrepresent.
    """
    data = _boundaries()
    if not data:
        return []

    entries = legend_entries(plottable, label_for=label_for)
    entry_of_group = {id(group): idx for idx, (_label, _color, members) in enumerate(entries)
                      for group in members}

    # Per-entry targeted-zip count per county, from each entry's own
    # (combined, still small) zip list -- cheap, one geo_resolver call per
    # underlying group.
    targeted = {}   # fips -> {entry_idx: count}
    for group in plottable:
        entry_idx = entry_of_group.get(id(group))
        if entry_idx is None:
            continue
        result = geo_resolver.zips_to_counties(group["resolved_zips"])
        for fips_list in result.resolved.values():
            if not fips_list:
                continue
            fips = fips_list[0]
            if fips not in data.get("counties", {}):
                continue
            targeted.setdefault(fips, {})
            targeted[fips][entry_idx] = targeted[fips].get(entry_idx, 0) + 1
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
    for fips, by_entry in targeted.items():
        total = totals.get(fips, 0)
        if total <= 0:
            continue
        covering = sorted(((idx, count) for idx, count in by_entry.items()
                           if (count / total) >= _COUNTY_FILL_MIN_COVERAGE),
                          key=lambda kv: -kv[1])
        if not covering:
            continue
        if len(covering) == 1:
            idx, _count = covering[0]
            claimed.append((fips, ("solid", entries[idx][1])))
        elif len(covering) == 2:
            idx_a, idx_b = covering[0][0], covering[1][0]
            claimed.append((fips, ("overlap",
                                   (entries[idx_a][1], entries[idx_b][1]),
                                   (entries[idx_a][0], entries[idx_b][0]))))
        else:
            claimed.append((fips, ("multi", tuple(entries[idx][0] for idx, _count in covering))))
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

    Color follows the AUDIENCE, not the group -- see `legend_entries` for
    the whole model. Every group under one audience still on that
    audience's shared color folds into one legend entry; once a rep has
    broken any group out (`color_locked`), that entry names ITS OWN
    members' markets too (never a bare, ambiguous audience name sitting
    above a qualified one), and the broken-out group gets its own entry,
    named "{audience} ({its own geo_label})". Two DIFFERENT audiences whose
    resolved geography overlaps never let one fill silently cover the other
    -- hatched in both colors with its own "Audience A + Audience B" legend
    row for exactly two; three or more is a different, neutral cross-hatch
    fill (`_fill_multi_rings`) rather than a hatch that would have to drop
    one, with one legend row naming everyone involved. Every legend label
    is fit to what the legend can actually show (`_fit_legend_label`,
    measured against the real glyphs) before the legend column is sized.
    `label_for` threads through to `tg.geo_label`'s own market-key ->
    display-name lookup (`app._market_display_name`).
    """
    plottable = groups_with_zips(groups)
    if not plottable:
        return None

    points = geo_resolver._data()["zip_points"]
    entries = legend_entries(plottable, label_for=label_for)
    series = []  # (color, [(lat, lon), ...], label) -- one per legend entry
    all_lat, all_lon = [], []
    for label, color, members in entries:
        coords = []
        for group in members:
            coords.extend(tuple(points[z]) for z in group["resolved_zips"] if z in points)
        if not coords:
            continue
        series.append((color, coords, label))
        all_lat.extend(c[0] for c in coords)
        all_lon.extend(c[1] for c in coords)
    if not series:
        return None

    palette = _DARK_PALETTE if dark else _LIGHT_PALETTE
    boundary_data = _boundaries()
    county_fills = _touched_counties(plottable, label_for=label_for)
    legend_font = _legend_font()

    # The hard ceiling every legend label is fit against: whatever's left
    # of the canvas once the map keeps its own minimum width. Never
    # SMALLER than the original fixed budget (_LEGEND_TEXT_MAX_WIDTH) --
    # a narrow canvas keeps today's behaviour -- but grows past it on a
    # wide canvas, using real available space instead of an arbitrary
    # constant. This is what makes "the legend can never exceed the
    # canvas" an actual guarantee rather than a hope: `_fit_legend_label`
    # and `_fit_combined` both fit AGAINST this number, so `legend_w`
    # (derived from what they actually returned) can only exceed it in
    # the one residual case neither can help -- a single label with
    # nothing left to shrink (no parenthetical, no siblings to summarize)
    # that's STILL wider than the whole canvas affords, which is exactly
    # the "nothing left to shrink without lying" case _fit_legend_label's
    # own docstring already accepts, one level up.
    text_ceiling = max(_LEGEND_TEXT_MAX_WIDTH,
                       width_px - _MIN_MAP_WIDTH - LEGEND_SWATCH - 3 * LEGEND_PADDING)

    overlap_entries = []  # (color_a, color_b, "Audience A + Audience B") for the legend
    multi_entries = []    # "Audience A + Audience B + Audience C" (3+) for the legend
    seen_overlap_labels, seen_multi_labels = set(), set()
    for _fips, fill in county_fills:
        if fill[0] == "overlap":
            _kind, (color_a, color_b), (label_a, label_b) = fill
            overlap_label = _fit_combined([label_a, label_b], legend_font, text_ceiling)
            if overlap_label not in seen_overlap_labels:
                seen_overlap_labels.add(overlap_label)
                overlap_entries.append((color_a, color_b, overlap_label))
        elif fill[0] == "multi":
            _kind, labels = fill
            multi_label = _fit_combined(list(labels), legend_font, text_ceiling)
            if multi_label not in seen_multi_labels:
                seen_multi_labels.add(multi_label)
                multi_entries.append(multi_label)

    # The fitted extent is the plotted points UNION the filled counties'
    # own bboxes (lon/lat) -- a filled county reaches further than its
    # zips' centroids, and framing on the points alone would clip it.
    fit_lats, fit_lons = list(all_lat), list(all_lon)
    for fips, _fill in county_fills:
        bbox = boundary_data["counties"][fips]["bbox"]  # [minlon, minlat, maxlon, maxlat]
        fit_lons.extend([bbox[0], bbox[2]])
        fit_lats.extend([bbox[1], bbox[3]])

    # Fit every remaining (non-overlap, non-multi -- already fit above)
    # legend label to what the legend can actually show BEFORE sizing the
    # legend column from them -- a qualified "Audience (m1, m2, ..., m9)"
    # label is measured against the real rendered glyphs
    # (`_fit_legend_label`, never a character count) and shortened to
    # "Audience (m1, m2, +7 more)" rather than either overflowing the frame
    # or silently vanishing past LEGEND_MAX_WIDTH's own clamp.
    series = [(color, coords, _fit_legend_label(label, legend_font, text_ceiling))
             for color, coords, label in series]

    legend_w = _legend_width(series, overlap_entries, multi_entries)
    map_w = max(_MIN_MAP_WIDTH, width_px - legend_w)

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
        for fips, fill in county_fills:
            feature = boundary_data["counties"].get(fips)
            if not feature:
                continue
            rings = [[project(lat, lon) for lon, lat in ring]
                    for ring in feature["rings"] if len(ring) >= 3]
            if not rings:
                continue
            if fill[0] == "overlap":
                _kind, (color_a, color_b), _labels = fill
                _fill_overlap_rings(fill_layer, rings, color_a, color_b, dark, palette)
            elif fill[0] == "multi":
                _fill_multi_rings(fill_layer, rings, dark, palette)
            else:
                _kind, color = fill
                rgb = _hex_to_rgb(color)
                if dark:
                    rgb, alpha = _dark_fill_style(rgb)
                else:
                    alpha = palette["county_fill_alpha"]
                for pixels in rings:
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
    _draw_legend(draw, series, map_w, height_px, palette, overlap_entries, multi_entries, img=img)

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


_OVERLAP_STRIPE_SPACING = 7
_OVERLAP_STRIPE_WIDTH = 3


def _fill_overlap_rings(fill_layer, rings, color_a, color_b, dark, palette):
    """Paint one county's (already-projected) rings as a hatch of color_a
    (a normal solid base) and color_b (diagonal stripes on top) -- two
    audiences' territory meeting here, so neither fill silently wins.

    Deliberately not an alpha blend of the two colors: blue and orange
    (say) blended read as a third, muddy hue that announces nothing about
    which two audiences are actually present. A stripe pattern keeps both
    source colors legible.

    Built as ONE flat pattern first -- every pixel is EXACTLY color_a or
    color_b, never a blend of the two -- then masked to these rings and
    alpha-composited onto `fill_layer` in a single step. Compositing
    color_b's stripes as a SEPARATE pass on top of an already-composited
    color_a base was tried and rejected: alpha-blending translucent
    color_b over translucent color_a produces a third, muddy hue at every
    stripe pixel -- exactly the "silent blend" this whole fill exists to
    avoid, just moved one layer down. `ImageDraw` replaces pixels outright
    (confirmed -- it does not blend against what's already drawn), so
    building the pattern as flat replacements first and blending against
    `fill_layer` only ONCE keeps both source colors legible.
    """
    rgb_a, rgb_b = _hex_to_rgb(color_a), _hex_to_rgb(color_b)
    if dark:
        # _dark_hatch_style, not _dark_fill_style -- a hatch is exempt from
        # the solid-fill muting rule; see that function's own docstring.
        rgb_a, alpha_a = _dark_hatch_style(rgb_a)
        rgb_b, alpha_b = _dark_hatch_style(rgb_b)
    else:
        alpha_a = alpha_b = palette["county_fill_alpha"]

    size = fill_layer.size
    width, height = size
    pattern = Image.new("RGBA", size, rgb_a + (alpha_a,))
    pattern_draw = ImageDraw.Draw(pattern)
    diag = width + height
    for offset in range(-diag, diag, _OVERLAP_STRIPE_SPACING):
        pattern_draw.line([(offset, 0), (offset + height, height)],
                          fill=rgb_b + (alpha_b,), width=_OVERLAP_STRIPE_WIDTH)

    mask = Image.new("L", size, 0)
    mask_draw = ImageDraw.Draw(mask)
    for pixels in rings:
        mask_draw.polygon(pixels, fill=255)

    masked_pattern = Image.new("RGBA", size, (0, 0, 0, 0))
    masked_pattern.paste(pattern, (0, 0), mask)
    fill_layer.alpha_composite(masked_pattern)


def _fill_multi_rings(fill_layer, rings, dark, palette):
    """Paint one county's (already-projected) rings as a neutral cross-
    hatch -- three or more DIFFERENT audiences meeting here, which a two-
    color hatch has no honest way to show without dropping one. Fixed,
    colorless (`_MULTI_FILL_COLOR`, never derived from any audience
    actually involved) and cross-hatched in BOTH diagonals, so it reads as
    "several, unspecified" at a glance and can't be mistaken for a real
    two-way overlap's single-diagonal, two-color hatch.

    Same one-flat-pattern-then-one-composite construction as
    `_fill_overlap_rings`, for the same reason: drawing the darker
    cross-hatch lines as a SEPARATE alpha-blended pass on top of an
    already-composited base washes both toward the same middling grey
    instead of reading as base-plus-lines.
    """
    base_rgb = _hex_to_rgb(_MULTI_FILL_COLOR)
    line_rgb = tuple(max(0, c - 70) for c in base_rgb)
    if dark:
        # _dark_hatch_style, not _dark_fill_style -- same exemption as
        # _fill_overlap_rings above.
        base_rgb, alpha = _dark_hatch_style(base_rgb)
        line_rgb, _line_alpha = _dark_hatch_style(line_rgb)
    else:
        alpha = palette["county_fill_alpha"]

    size = fill_layer.size
    width, height = size
    pattern = Image.new("RGBA", size, base_rgb + (alpha,))
    pattern_draw = ImageDraw.Draw(pattern)
    diag = width + height
    for offset in range(-diag, diag, _OVERLAP_STRIPE_SPACING):
        pattern_draw.line([(offset, 0), (offset + height, height)],
                          fill=line_rgb + (220,), width=_OVERLAP_STRIPE_WIDTH)
        pattern_draw.line([(offset, height), (offset + height, 0)],
                          fill=line_rgb + (220,), width=_OVERLAP_STRIPE_WIDTH)

    mask = Image.new("L", size, 0)
    mask_draw = ImageDraw.Draw(mask)
    for pixels in rings:
        mask_draw.polygon(pixels, fill=255)

    masked_pattern = Image.new("RGBA", size, (0, 0, 0, 0))
    masked_pattern.paste(pattern, (0, 0), mask)
    fill_layer.alpha_composite(masked_pattern)


def _multi_tile(size):
    """The legend's own swatch for a `multi` entry -- same cross-hatch
    language as `_fill_multi_rings`, at legend-swatch scale."""
    base_rgb = _hex_to_rgb(_MULTI_FILL_COLOR)
    line_rgb = tuple(max(0, c - 70) for c in base_rgb)
    tile = Image.new("RGB", (size, size), base_rgb)
    tile_draw = ImageDraw.Draw(tile)
    diag = size * 2
    for offset in range(-diag, diag, 5):
        tile_draw.line([(offset, 0), (offset + size, size)], fill=line_rgb, width=2)
        tile_draw.line([(offset, size), (offset + size, 0)], fill=line_rgb, width=2)
    return tile


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
    # Real bug, found live 2026-09-03 root-causing test_targeting_map.py's
    # reported multi-minute stalls: a near-point group (the common case is
    # literally ONE zip -- `all_lat`/`all_lon` then hold a single value each)
    # collapsed max_x-min_x to exactly 0, and the old `1e-6`-DEGREE floor
    # here was meant only to dodge a literal division by zero, not to stand
    # in for a real span -- it drove scale_x/scale_y to roughly 9e8 px per
    # degree. `_matching_boundaries`'s own search box is floored sanely
    # (`_MIN_SEARCH_DEGREES`, "a floor for a near-point group (e.g. one
    # zip)"), so it still correctly found the real, nearby county/state
    # boundaries -- but projecting THEIR coordinates through this insane
    # scale flung them to pixel positions in the hundreds of millions,
    # and Pillow's line rasterizer (`_draw_rings`) takes time proportional
    # to that unclipped pixel-space line length: single `draw.line` calls
    # measured at 6-10 SECONDS each, dozens of them per render, is exactly
    # the "pathologically slow, CPU-bound, not a deadlock" symptom BACKLOG.md
    # recorded without a root cause. Fixed by flooring the span at the SAME
    # real-world minimum the search box already uses, not an infinitesimal
    # epsilon -- a near-point group now renders at a sane, if maximally
    # zoomed-in, scale instead of one that makes every nearby boundary
    # explode off-canvas.
    span_x = max(max_x - min_x, _MIN_SEARCH_DEGREES * cos_lat)
    span_y = max(max_y - min_y, _MIN_SEARCH_DEGREES)
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


# A hatch is exempt from `_dark_fill_style`'s muting -- that function exists
# so a SOLID county fill sits quietly under text and doesn't compete with the
# map's own place labels; a hatch is already saying "look here" (it exists
# specifically to flag two or more audiences overlapping), so muting it the
# same way defeats the point. Solids stay muted exactly as they were.
#
# Measured directly against a real dark render before picking these:
# `_dark_fill_style` drove both stripe colors of a real overlap to the SAME
# luminance at 23-25% alpha, landing only 6.6 DeltaE apart once composited
# over the deck's dark background (down from 35.3 DeltaE at full strength --
# texture, not a legible two-color hatch). This lighten/alpha pair was
# chosen by measuring the actual composited DeltaE across the reordered
# palette's realistic pairs and picking the smallest values that keep the
# WORST case clearly separable:
#   lighten=0.30, alpha=65  (old, muted)     blue+orange DeltaE 24.1
#   lighten=0.15, alpha=200 (this)           blue+orange DeltaE 80.0
# -- comfortably past the 45 DeltaE bar the solid palette itself is held to
# (tests/test_targeting_groups.py), without full/glaring saturation on a
# slide meant to read calmly otherwise.
_DARK_HATCH_LIGHTEN = 0.15
_DARK_HATCH_ALPHA = 200


def _dark_hatch_style(rgb):
    """(lightened_rgb, alpha) for a dark-mode HATCH fill -- deliberately not
    `_dark_fill_style`; see the constants above for why and how these were
    chosen."""
    return _lighten(rgb, _DARK_HATCH_LIGHTEN), _DARK_HATCH_ALPHA


def _legend_font():
    # Pillow's own bundled font -- no filesystem lookup, so nothing to
    # differ between a Windows dev box and Streamlit Cloud's Linux the way
    # text_metrics.py's brand-font resolution deliberately can. A map
    # legend doesn't need to match Calibri/Proxima Nova; it needs to always
    # be there.
    return ImageFont.load_default(size=13)


def _place_font():
    return ImageFont.load_default(size=11)


def _legend_width(series, overlap_entries=(), multi_entries=()):
    """Every label handed in here has already been through
    `_fit_legend_label` (individually for `series`, per-side for an
    overlap/multi combination via `_fit_combined`), so the widest one is
    normally already <= `_LEGEND_TEXT_MAX_WIDTH` and this settles at or
    under `LEGEND_MAX_WIDTH` exactly as before -- no explicit clamp needed
    for the ordinary case. The one place that isn't true: a 3+-way overlap
    combining two ALREADY-qualified labels under the same broken-out
    audience can still be too wide even after each side is fit to its own
    share (there is nothing shorter left to say without inventing a kind
    of truncation -- abbreviating the audience name itself -- this was
    never asked to do). Widening the column past LEGEND_MAX_WIDTH in that
    residual case is the honest choice: a slightly wider legend beats text
    silently drawn past the canvas edge.
    """
    font = _legend_font()
    widest = 0
    labels = [label for _color, _coords, label in series]
    labels += [label for _a, _b, label in overlap_entries]
    labels += list(multi_entries)
    for label in labels:
        widest = max(widest, _text_width(font, label))
    return max(LEGEND_MIN_WIDTH, widest + LEGEND_SWATCH + 3 * LEGEND_PADDING)


def _text_width(font, text):
    bbox = font.getbbox(text)
    return bbox[2] - bbox[0]


# The pixel budget a qualified label's market list is fit against --
# derived from LEGEND_MAX_WIDTH itself (the one real width_px this app ever
# renders the deck's map at -- 900px -- is fixed, so there is no second
# "true rendered slide size" to fit against; LEGEND_MAX_WIDTH already IS
# that fit). Never grows the legend past LEGEND_MAX_WIDTH; a label this
# wide would have hit that clamp anyway.
_LEGEND_TEXT_MAX_WIDTH = LEGEND_MAX_WIDTH - LEGEND_SWATCH - 3 * LEGEND_PADDING


def _fit_legend_label(label, font, max_text_width=_LEGEND_TEXT_MAX_WIDTH):
    """`label`, shortened to fit `max_text_width` pixels in `font` -- by
    dropping trailing items from a trailing "(a, b, c)" list and saying how
    many were dropped, never by shrinking type and never by a silent
    character cut. Measured against the real rendered glyphs (`font`),
    never a character count -- a market name's own width varies too much
    for a length heuristic to mean anything at legend scale.

    A label with no such list (a bare audience name, an "A + B" overlap
    label with only two names to begin with, or a qualified label whose
    parenthetical is only ONE item -- "+1 more" would claim something was
    dropped when the single item shown WAS the whole list) is returned
    unchanged if it's still too wide -- there's nothing left to shrink
    without either lying about what's hidden or inventing a different kind
    of truncation than the one this was asked to solve. A genuinely long
    audience name is exactly this case: qualifying it can't make it fit.
    """
    if _text_width(font, label) <= max_text_width:
        return label
    match = re.match(r"^(.*) \((.+)\)$", label)
    if not match:
        return label
    prefix, inner = match.group(1), match.group(2)
    items = [item.strip() for item in inner.split(",")]
    items = [item for item in items if not re.match(r"^\+\d+ more$", item)]
    if len(items) <= 1:
        return label
    for keep in range(len(items) - 1, 0, -1):
        dropped = len(items) - keep
        candidate = f"{prefix} ({', '.join(items[:keep])}, +{dropped} more)"
        if _text_width(font, candidate) <= max_text_width:
            return candidate
    candidate = f"{prefix} (+{len(items)} more)"
    return candidate


def _fit_combined(labels, font, max_text_width=_LEGEND_TEXT_MAX_WIDTH):
    """Join `labels` (audience names, from an overlap or multi county fill)
    with " + ", fitting the WHOLE list to `max_text_width`. Module-level
    (not a closure inside `render_map`) so it's directly testable.

    Each label gets its OWN full chance to fit first -- the same budget a
    plain series entry gets, via `_fit_legend_label` -- rather than a small
    pre-divided N-way share of the budget. The old N-way pre-split was the
    actual defect behind two live reports: it gave a bare audience name (no
    parenthetical to trim, so `_fit_legend_label` can't shrink it at ANY
    budget) a share too small to matter, so it passed through UNCHANGED
    regardless -- confirmed on a real 4-audience Annapolis identity-collapse
    case, 565px through a 272px budget, completely untouched by the "fit."

    What happens next differs by count, because the two shapes mean
    different things to a client reading the legend:
    - **2 items** (a genuine two-way overlap): if the fitted pair is STILL
      too wide once each has had its own full chance, BOTH names are kept
      in full anyway. Naming the specific pair is the entire point of an
      overlap entry -- "Audience A + 1 more" is not a legitimate answer to
      "which two audiences overlap here," and a 2-item list has nothing
      shorter to say without inventing a new kind of lie. The legend column
      widens to fit it (`_legend_width` already does this, deliberately,
      for exactly this "nothing left to shrink" case) -- an honest, wider
      legend beats silently naming the wrong (or no) audience.
    - **3+ items** (a multi/neutral-hatch entry, genuinely overlapping but
      NOT identical geography -- an identical-geography set collapses to
      one solid-color entry upstream, in `_collapse_identical_geography`,
      and never reaches this at all): keep whole names, in document order,
      for as many as fit, then "+N more" -- the SAME pattern
      `_fit_legend_label`'s own parenthetical branch already uses for a
      market list, one level up, applied to a list of whole audience names
      instead of markets. Dropping to a count is legitimate here, unlike
      the 2-item case, because a 3+-way entry is already a summary rather
      than a specific pairing.
    """
    fitted = [_fit_legend_label(label, font) for label in labels]
    joined = " + ".join(fitted)
    if len(fitted) <= 2 or _text_width(font, joined) <= max_text_width:
        return joined
    for keep in range(len(fitted) - 1, 0, -1):
        candidate = " + ".join(fitted[:keep] + [f"+{len(fitted) - keep} more"])
        if _text_width(font, candidate) <= max_text_width:
            return candidate
    return joined


def _hatch_tile(rgb_a, rgb_b, size):
    """A size x size RGB tile: solid rgb_a with rgb_b diagonal stripes --
    the legend's own swatch for an overlap entry, same visual language as
    the county hatch it's explaining. Drawing on a tile exactly this size
    is what keeps the stripes from bleeding into the row above/below or the
    label beside it -- PIL clips a draw call to the image it's drawn on,
    so there's no separate clip-region math to get right.
    """
    tile = Image.new("RGB", (size, size), rgb_a)
    tile_draw = ImageDraw.Draw(tile)
    diag = size * 2
    for offset in range(-diag, diag, 5):
        tile_draw.line([(offset, 0), (offset + size, size)], fill=rgb_b, width=2)
    return tile


def _draw_legend(draw, series, x0, height, palette, overlap_entries=(), multi_entries=(), img=None):
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
            return  # more groups than the legend has room for -- stop rather than overflow the frame

    # Overlap rows come after every audience's own row -- "where two things
    # meet" reads more naturally once the two things themselves are already
    # named above it.
    for color_a, color_b, label in overlap_entries:
        x_left = x0 + LEGEND_PADDING
        y_top = y + 3
        if img is not None:
            tile = _hatch_tile(_hex_to_rgb(color_a), _hex_to_rgb(color_b), LEGEND_SWATCH)
            img.paste(tile, (x_left, y_top))
        else:
            # No Image handle to paste a tile onto (a caller drawing
            # straight onto a Draw object with no image, e.g. a future
            # test) -- a flat fallback swatch beats a missing row.
            draw.rectangle([x_left, y_top, x_left + LEGEND_SWATCH, y_top + LEGEND_SWATCH],
                          fill=_hex_to_rgb(color_a))
        draw.rectangle([x_left, y_top, x_left + LEGEND_SWATCH, y_top + LEGEND_SWATCH],
                       outline=_hex_to_rgb(color_b), width=1)
        draw.text((x0 + LEGEND_PADDING + LEGEND_SWATCH + 6, y), label, font=font, fill=palette["legend_text"])
        y += LEGEND_ROW_HEIGHT
        if y > height - LEGEND_PADDING:
            return

    # Multi (3+) rows last -- rarer than a two-way overlap, and named last
    # for the same reason overlap rows follow the audiences they involve.
    for label in multi_entries:
        x_left = x0 + LEGEND_PADDING
        y_top = y + 3
        if img is not None:
            img.paste(_multi_tile(LEGEND_SWATCH), (x_left, y_top))
        else:
            draw.rectangle([x_left, y_top, x_left + LEGEND_SWATCH, y_top + LEGEND_SWATCH],
                          fill=_hex_to_rgb(_MULTI_FILL_COLOR))
        draw.rectangle([x_left, y_top, x_left + LEGEND_SWATCH, y_top + LEGEND_SWATCH],
                       outline=_hex_to_rgb(_MULTI_FILL_COLOR), width=1)
        draw.text((x0 + LEGEND_PADDING + LEGEND_SWATCH + 6, y), label, font=font, fill=palette["legend_text"])
        y += LEGEND_ROW_HEIGHT
        if y > height - LEGEND_PADDING:
            return


def _draw_place_labels(draw, project, frame_bounds, map_w, height, palette,
                       max_labels=MAX_PLACE_LABELS):
    """The largest places actually in frame, up to `max_labels`,
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
        if placed >= max_labels:
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


# ---------------------------------------------------------------------------
# Weighted-fill choropleth -- the attribution report's zip heat map.
#
# A SIBLING of render_map, not a mode inside it. The two answer different
# questions and share almost nothing above the projection layer:
#
#   render_map   : "which group targeted where" -- color follows the
#                  AUDIENCE, one categorical color per legend entry, plus
#                  the whole overlap/multi-hatch model for two or three
#                  audiences landing in one county.
#   choropleth   : "how hard did each zip respond" -- ONE series, a value
#                  per zip, intensity carrying magnitude. No audiences, no
#                  overlaps, no per-group color.
#
# Threading a `weights=` parameter through render_map would mean branching
# around legend_entries, _touched_counties, both hatch fills and the legend
# -- i.e. most of its body -- while leaving that machinery loaded and inert,
# and putting every one of render_map's existing guards at risk for a
# feature none of them cover. The shared parts (projection, boundary
# lookup, outlines, place labels) are called from here instead.
#
# Phase 5's targeted-vs-visitor overlay composes on top of this rather than
# becoming a third mode: draw the choropleth, then stroke the targeted zips'
# outlines over it.
# ---------------------------------------------------------------------------

# Periwinkle -> navy, the report deck's own accent through its header color.
# Five steps because the bins are five; index 0 is the lightest.
CHOROPLETH_RAMP = ["#D8DAF2", "#A9AEE4", "#7178C9", "#3C46A0", "#000946"]
_CHOROPLETH_BINS = 5

# Untargeted land, tinted so the shaded ZCTAs read as FIGURE against a
# GROUND rather than floating on white. Deliberately warmer and lighter than
# any ramp step -- it must never be mistaken for a low bin, which is the one
# way a ground tint can actively mislead on a choropleth.
_CHOROPLETH_GROUND = (243, 243, 240)
# A choropleth carries no group legend, so it can afford more orientation
# labels than the targeting map's 8 -- and needs them, since a reader is
# looking for their own zip rather than a named audience.
_CHOROPLETH_PLACE_LABELS = 16

# Phase 5's targeted-vs-visitor overlay (ATTRIBUTION_REPORT_PLAN.md) --
# tab10 orange, the same bold, saturated hue GROUP_COLORS[1] already uses
# for a targeting-map audience (2026-09-08's own too-close-to-tell-apart
# fix), reused here as a single fixed outline color rather than adding a
# second ramp: this overlay names ONE thing (targeted vs. not), never
# several audiences, so there is nothing for a second palette to encode.
TARGETED_OUTLINE_COLOR = "#FF7F0E"
TARGETED_OUTLINE_WIDTH = 3
_TARGETED_LEGEND_LABEL = "Targeted zip codes"


def quantile_thresholds(values, bins=_CHOROPLETH_BINS):
    """The bin edges for a QUANTILE ramp -- equal COUNTS per bin, not equal
    value ranges.

    Measured on the two real attribution exports, and the reason this isn't
    linear: MW's 242 zips run 0% to 23.1% attributed rate against a 1.03%
    median, so five equal-width bins put **228 of 242 zips (94%) in the
    palest one** and leave two outliers dark -- a map that is one flat
    colour plus two dots, carrying no information about the 94%. The same
    split on Cardinal buries 82%. Quantile bins give 48/48/49/48/49 and
    4/4/5/4/5 respectively, so the ramp actually varies across the
    territory a client is looking at.

    Ties are why this returns thresholds rather than pre-bucketed values:
    a value equal to a threshold lands in the HIGHER bin, so a long run of
    identical rates (a real case -- many zips share 0%) collapses into one
    bin instead of being split arbitrarily across two.
    """
    ordered = sorted(v for v in values if v is not None)
    if not ordered:
        return []
    return [ordered[int(len(ordered) * k / bins)] for k in range(1, bins)]


def _bin_for(value, thresholds):
    return sum(1 for t in thresholds if value >= t)


def _choropleth_legend_labels(thresholds, value_format):
    """One label per bin, describing the RANGE it covers. Built here and
    used for BOTH the legend column's width and its drawing -- the first
    version measured a placeholder ("0.49%+") while drawing a range
    ("0.49% - 0.81%"), so the column was sized about half what it needed
    and every label rendered clipped at the canvas edge."""
    edges = [None] + list(thresholds) + [None]
    labels = []
    for index in range(len(CHOROPLETH_RAMP)):
        low, high = edges[index], edges[index + 1]
        if low is None:
            labels.append(f"under {value_format(high)}")
        elif high is None:
            labels.append(f"{value_format(low)} and up")
        else:
            labels.append(f"{value_format(low)} - {value_format(high)}")
    return labels


def render_choropleth(zip_values, width_px=900, height_px=560,
                      background=(255, 255, 255), dark=False,
                      legend_title="Attributed rate", value_format=None,
                      targeted_zips=None):
    """A choropleth PNG of `zip_values` -- {zip: numeric value} -- with each
    zip's real ZCTA AREA shaded by quantile bin.

    Returns `(png_bytes, missing_zips)`. `missing_zips` is every zip that
    had a value but no polygon (PO-box-only zips, retired ZCTAs, or a state
    nobody has run build_zcta_boundaries.py for yet). Most of these still
    have a centroid POINT and draw as a dot in their own bin colour so
    they're never silently dropped; a zip with no point either (2026-09-13
    Ashburn/20149 find -- a real, deliverable PO-box-only zip has no ZCTA
    and so no coordinate at all) genuinely can't be drawn or framed and is
    excluded from the map, but is STILL named in `missing_zips` -- the
    caller's warning is what keeps it from vanishing with no record at
    all, since the map itself has nothing to show for it. `(None, [])`
    when there is nothing at all to draw -- same leave-it-alone contract
    as render_map.

    This shades AREAS, not centroids. An earlier version shaded dots
    because nothing in this repo carried zip geometry -- `geo_crosswalk`
    has zip points and `map_boundaries` has county/state polygons, and
    neither has zip polygons. That is a data gap, now filled by
    build_zcta_boundaries.py, rather than something drawing could solve.
    A Voronoi tessellation around the centroids was explicitly rejected:
    it would look like a choropleth without being one, and a client
    finding their own zip would see the wrong shape.

    County and state outlines are drawn ON TOP of the fills, for
    orientation -- a shaded field with no visible county lines reads as
    an abstract blob rather than a map of somewhere.

    `targeted_zips` (ATTRIBUTION_REPORT_PLAN.md Phase 5, optional -- an
    iterable of zip codes) is the targeted-vs-visitor overlay: this stays
    ONE series (the module-level design note above this function explains
    why weights/audiences were rejected as a third render_map-like mode) --
    draw the choropleth exactly as always, then stroke the targeted zips'
    own polygon outlines over it, in a single fixed color never drawn from
    `zip_values` itself. The map's frame is widened to include the
    targeted zips' own geometry (not just wherever `zip_values` already
    has data), so a targeted area with zero visits is still visible rather
    than clipped out of frame. A targeted zip lacking a polygon is simply
    not outlined -- never reported in `missing_zips`, which is scoped to
    `zip_values` alone -- and the legend gains its "Targeted zip codes"
    swatch ONLY when at least one outline was actually drawn, never merely
    because `targeted_zips` was passed (a client-facing map must never
    claim a layer it didn't draw). None/empty `targeted_zips` reproduces
    the exact PNG this function always produced -- a linked proposal with
    no resolved zips (avails_mode off, or no avails import ever ran) must
    degrade to the ordinary visitor-only map, never a crash or an empty
    second series.
    """
    points = geo_resolver._data()["zip_points"]
    plotted = {}
    for code, value in (zip_values or {}).items():
        key = str(code).strip().zfill(5)
        if value is not None:
            plotted[key] = float(value)
    if not plotted:
        return None, []

    value_format = value_format or (lambda v: f"{v * 100:.2f}%")
    thresholds = quantile_thresholds(plotted.values())
    palette = _DARK_PALETTE if dark else _LIGHT_PALETTE
    legend_font = _legend_font()

    rings_by_zip, missing_no_ring = zcta_rings_for(list(plotted))
    # A real, deliverable PO-box-only zip (2026-09-13's Ashburn/20149 find)
    # has no ZCTA at all, so it's ALSO absent from `points` -- not just
    # `rings_by_zip`. That used to be filtered out of `plotted` before
    # `zcta_rings_for` ever ran, which meant it never reached `missing`
    # either: no dot, no polygon, no count, a real zip with real data just
    # gone with nothing to show for it. Split `missing_no_ring` into what
    # CAN still be drawn as a dot (has a point) and what genuinely can't
    # (no point either) -- the latter is dropped from `plotted`/framing
    # entirely (there is no coordinate to draw or frame with) but still
    # named in the combined `missing` this function returns, so the
    # caller's "N zip code(s) have no ZCTA boundary" warning accounts for
    # it the same as any other unplottable zip, rather than the count
    # silently undercounting by exactly the zips hit hardest by the gap.
    missing_with_point = [c for c in missing_no_ring if c in points]
    missing_no_point = [c for c in missing_no_ring if c not in points]
    for code in missing_no_point:
        plotted.pop(code, None)
    missing = missing_with_point + missing_no_point
    if not plotted:
        return None, missing

    # Targeted-zip resolution happens BEFORE the frame is computed, so its
    # own geometry can widen the frame -- a targeted zip with zero visits
    # must still be on the map, not clipped to whatever `zip_values` alone
    # covers. `targeted_drawn` is settled here too (which polygons actually
    # resolved), before legend sizing needs to know whether to reserve a
    # row for it.
    targeted_codes = {str(z).strip().zfill(5) for z in (targeted_zips or ())}
    if targeted_codes:
        extra_needed = targeted_codes - set(rings_by_zip)
        if extra_needed:
            extra_rings, _extra_missing = zcta_rings_for(sorted(extra_needed))
            rings_by_zip.update(extra_rings)
    targeted_drawn = {code for code in targeted_codes if code in rings_by_zip}

    # Framed on the POLYGONS, not the centroids -- a zip's area reaches
    # past its own centroid, and fitting to points alone clips the border
    # zips of the campaign in half.
    lats, lons = [], []
    for code in plotted:
        for ring in rings_by_zip.get(code, ()):
            for x, y in ring:
                lons.append(x)
                lats.append(y)
        if code not in rings_by_zip:
            lat, lon = points[code]
            lats.append(lat)
            lons.append(lon)
    for code in targeted_drawn - set(plotted):
        for ring in rings_by_zip.get(code, ()):
            for x, y in ring:
                lons.append(x)
                lats.append(y)

    legend_labels = _choropleth_legend_labels(thresholds, value_format)
    extra_legend_labels = [_TARGETED_LEGEND_LABEL] if targeted_drawn else []
    legend_w = LEGEND_SWATCH + 3 * LEGEND_PADDING + max(
        [_text_width(legend_font, legend_title)]
        + [_text_width(legend_font, label) for label in legend_labels]
        + [_text_width(legend_font, label) for label in extra_legend_labels])
    map_w = max(_MIN_MAP_WIDTH, width_px - int(legend_w))

    mode = "RGBA" if dark else "RGB"
    img = Image.new(mode, (width_px, height_px), (0, 0, 0, 0) if dark else background)
    draw = ImageDraw.Draw(img)
    project, frame_bounds = _projector(lats, lons, map_w, height_px)

    span_lon, span_lat = (max(lons) - min(lons)), (max(lats) - min(lats))
    search_box = (
        min(lons) - max(_MIN_SEARCH_DEGREES, span_lon * BOUNDARY_SEARCH_FRACTION),
        min(lats) - max(_MIN_SEARCH_DEGREES, span_lat * BOUNDARY_SEARCH_FRACTION),
        max(lons) + max(_MIN_SEARCH_DEGREES, span_lon * BOUNDARY_SEARCH_FRACTION),
        max(lats) + max(_MIN_SEARCH_DEGREES, span_lat * BOUNDARY_SEARCH_FRACTION),
    )
    counties, states = _matching_boundaries(search_box)

    # GROUND first: every county in frame gets a faint tint, so the shaded
    # ZCTAs sit on land rather than on white paper. Without it the filled
    # area reads as an abstract shape floating in a void -- the single
    # biggest gap against the basemap version of this slide.
    for feature in counties:
        for ring in feature["rings"]:
            pixels = [project(lat, lon) for lon, lat in ring]
            if len(pixels) >= 3:
                draw.polygon(pixels, fill=_CHOROPLETH_GROUND)

    # Then the data fills, then everything else over them.
    for code, value in plotted.items():
        rings = rings_by_zip.get(code)
        if not rings:
            continue
        rgb = _hex_to_rgb(CHOROPLETH_RAMP[_bin_for(value, thresholds)])
        for ring in rings:
            pixels = [project(y, x) for x, y in ring]
            if len(pixels) >= 3:
                draw.polygon(pixels, fill=rgb)

    for feature in counties:
        _draw_rings(draw, project, feature["rings"], palette["county_outline"][:3], width=1)
    for feature in states:
        _draw_rings(draw, project, feature["rings"], palette["state_outline"][:3], width=2)

    # A zip with no polygon but a real point still carries a real value --
    # drawn as a dot in its own bin colour rather than dropped, and
    # counted for the caller. `missing_no_point` codes have no coordinate
    # at all (see above) and were already excluded from `plotted`, so
    # `missing_with_point` -- not the combined `missing` -- is the safe
    # list to draw from here.
    for code in missing_with_point:
        lat, lon = points[code]
        x, y = project(lat, lon)
        rgb = _hex_to_rgb(CHOROPLETH_RAMP[_bin_for(plotted[code], thresholds)])
        draw.ellipse([x - DOT_RADIUS, y - DOT_RADIUS, x + DOT_RADIUS, y + DOT_RADIUS],
                    fill=rgb, outline=palette.get("dot_outline"))

    # Phase 5's targeted-vs-visitor overlay: stroke the targeted zips'
    # outlines OVER the choropleth (the module-level design note's own
    # phrasing), after the fills/boundaries/missing-dots are all down but
    # BEFORE place labels, so labels stay legible on top rather than
    # getting crossed by an outline.
    if targeted_drawn:
        outline_rgb = _hex_to_rgb(TARGETED_OUTLINE_COLOR)
        for code in targeted_drawn:
            _draw_rings(draw, project, rings_by_zip[code], outline_rgb, width=TARGETED_OUTLINE_WIDTH)

    # Place labels LAST and haloed. The light palette carries no halo (on a
    # targeting map, labels sit over pale 70/255 tints and read fine); here
    # they sit over solid navy fills, where unhaloed dark text is invisible.
    # A local palette copy rather than changing the shared one, so
    # render_map's own appearance is untouched.
    label_palette = dict(palette)
    if not label_palette.get("place_text_halo"):
        label_palette["place_text_halo"] = (255, 255, 255, 235)
    _draw_place_labels(draw, project, frame_bounds, map_w, height_px, label_palette,
                      max_labels=_CHOROPLETH_PLACE_LABELS)
    # The "Targeted zip codes" swatch is added ONLY when at least one
    # outline was actually drawn (`targeted_drawn`, resolved above from
    # real polygons) -- never merely because `targeted_zips` was passed,
    # per this function's own "never claim a layer it didn't draw" rule.
    extra_entries = [(_TARGETED_LEGEND_LABEL, TARGETED_OUTLINE_COLOR)] if targeted_drawn else None
    _draw_choropleth_legend(draw, map_w, legend_labels, legend_title, palette, legend_font,
                           height=height_px,
                           background=(0, 0, 0, 0) if dark else background,
                           extra_entries=extra_entries)

    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue(), missing


def _draw_choropleth_legend(draw, x0, labels, title, palette, font,
                           height=None, background=None, extra_entries=None):
    """See below -- `background`/`height` clear the legend column first.
    `extra_entries` is `[(label, hex_color), ...]` for an outline-only
    swatch drawn beneath the graded bins -- Phase 5's targeted-zip overlay
    (ATTRIBUTION_REPORT_PLAN.md), the one caller today."""
    if background is not None and height is not None:
        # Nothing clips a projected polygon to the map column, so a county
        # (ground tint) or even a dark ZCTA fill can reach under the legend
        # and sit behind its text. Harmless on this dataset, unreadable on
        # a differently-shaped one -- so the column is cleared before the
        # key is drawn rather than relying on the geography being kind.
        draw.rectangle([x0, 0, x0 + 10000, height], fill=background)
    _draw_choropleth_legend_body(draw, x0, labels, title, palette, font, extra_entries=extra_entries)


def _draw_choropleth_legend_body(draw, x0, labels, title, palette, font, extra_entries=None):
    """A graded key: one swatch per bin, labelled by the RANGE it covers --
    from the thresholds themselves rather than a bare "low..high", so a
    reader can place a specific zip's rate on the ramp instead of only
    ranking it. `labels` comes from _choropleth_legend_labels, the same
    call that sized this column.

    `extra_entries` (`[(label, hex_color), ...]`) draws below the graded
    bins as an OUTLINE-only swatch (no fill) -- visually distinct from the
    solid-fill bin swatches above, since it names a boundary drawn on the
    map, not a shaded value."""
    x = x0 + LEGEND_PADDING
    y = LEGEND_PADDING
    draw.text((x, y), title, font=font, fill=palette["legend_title"][:3])
    y += 22
    for color, label in zip(CHOROPLETH_RAMP, labels):
        draw.rectangle([x, y, x + LEGEND_SWATCH, y + LEGEND_SWATCH],
                      fill=_hex_to_rgb(color), outline=palette["county_outline"][:3])
        draw.text((x + LEGEND_SWATCH + 8, y + LEGEND_SWATCH / 2), label, font=font,
                 fill=palette["legend_text"][:3], anchor="lm")
        y += LEGEND_SWATCH + 8
    for label, color in extra_entries or ():
        draw.rectangle([x, y, x + LEGEND_SWATCH, y + LEGEND_SWATCH],
                      outline=_hex_to_rgb(color), width=2)
        draw.text((x + LEGEND_SWATCH + 8, y + LEGEND_SWATCH / 2), label, font=font,
                 fill=palette["legend_text"][:3], anchor="lm")
        y += LEGEND_SWATCH + 8
