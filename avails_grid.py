"""avails_grid.py -- one geography per real geography, the audience x
geography avails grid, and the plan checks that go with them.

Pure: no Streamlit, no DB. Works on targeting-group dicts (targeting_groups.py)
and on whatever the avails-PDF importer stamped onto them (`avails_source`,
`avails_rfpi`, `avails_geo_variant`, `avails_doc_full_flight`).

**A geography is its ZIP list or its DMA, never its name.** An avails
document can name one ZIP list two ways on two lines ("Eastern West Virginia
Delaware Southern Pennsylvania" / "Eastern WV Delaware Southern
Pennsylvania" -- Adventist Health, RFPID-269781), and two documents in one
proposal can each sell the same DMA. `geo_identity` keys on what is actually
targeted, so those collapse to one column on the slide and one region on the
map. Labels are cleaned from the document's own text (`clean_geo_label`) and
never assembled from the states the ZIPs happen to sit in.
"""
import hashlib
import re
from statistics import median

import targeting_groups as tg

DASH = "—"

_ZIP_OPTION_PREFIX_RE = re.compile(r"^\s*zip\s+option\s*[-–—]\s*", re.IGNORECASE)
_RADIUS_MILES_RE = re.compile(r"(\d+(?:\.\d+)?)\s*-?\s*(?:mi\b|mile)", re.IGNORECASE)
_STATE_ZIP_RE = re.compile(r"\b([A-Z]{2})\s+(\d{5})\b")
_STREET_SUFFIX_RE = re.compile(
    r"\b(?:St|Ave|Blvd|Dr|Rd|Ln|Way|Pl|Ct|Pkwy|Hwy|Cir|Ter|Trl|Sq)\b\.?", re.IGNORECASE)
_DIRECTIONAL_RE = re.compile(r"^(?:NE|NW|SE|SW|N|S|E|W)\b\.?\s*", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Labels
# ---------------------------------------------------------------------------
def _radius_city(text):
    """The city in "12501 Prosperity Dr Silver Spring MD 20904 50 Mile
    Radius" -- the text between the last street suffix and the trailing
    STATE ZIP. None when the address doesn't have that shape."""
    matches = list(_STATE_ZIP_RE.finditer(text))
    if not matches:
        return None
    state_zip = matches[-1]
    suffixes = list(_STREET_SUFFIX_RE.finditer(text[:state_zip.start()]))
    if suffixes:
        city = text[suffixes[-1].end():state_zip.start()]
    else:
        city = text[:state_zip.start()]
    city = _DIRECTIONAL_RE.sub("", city.strip(" .,-")).strip(" .,-")
    if not city or any(ch.isdigit() for ch in city):
        return None
    return city


def _format_miles(miles):
    return f"{int(miles)}" if float(miles).is_integer() else f"{miles:g}"


# A radius figure past this is a misread, not a radius: one real document
# prints "Reston VA 20190 Mile Radius" with no number, and the ZIP gets read
# as the mileage (LiveWell, RFPID-266894).
_MAX_PLAUSIBLE_RADIUS_MILES = 250


def _radius_street(text):
    """"945 Florida Ave NW" from "945 Florida Ave NW Washington DC 20001 5
    Mile Radius" -- the street address up to and including its suffix (and
    a trailing directional), or None."""
    matches = list(_STATE_ZIP_RE.finditer(text))
    if not matches:
        return None
    head = text[:matches[-1].start()]
    suffixes = list(_STREET_SUFFIX_RE.finditer(head))
    if not suffixes:
        return None
    street = head[:suffixes[-1].end()]
    directional = re.match(r"\s+(NE|NW|SE|SW|N|S|E|W)\b\.?", head[suffixes[-1].end():])
    if directional:
        street += " " + directional.group(1)
    return street.strip(" .,-") or None


def clean_geo_label(raw_name, radius_miles=None, with_street=False):
    """The client-facing name for one avails line's geography.

    - "Zip Option - <name>" -> "<name>" (the prefix is Salesforce plumbing).
    - A radius around a street address -> "<City> <N>-mi radius"
      ("Silver Spring 50-mi radius"); `with_street=True` puts the street
      first ("945 Florida Ave NW, Washington 5-mi radius"), for when two
      different radii would otherwise share one name.
    - Anything else (a DMA name) -> unchanged.
    """
    text = _ZIP_OPTION_PREFIX_RE.sub("", str(raw_name or "")).strip()
    if "radius" in text.lower():
        miles = radius_miles
        if miles is None or miles > _MAX_PLAUSIBLE_RADIUS_MILES:
            found = [float(m.group(1)) for m in _RADIUS_MILES_RE.finditer(text)]
            found = [m for m in found if m <= _MAX_PLAUSIBLE_RADIUS_MILES]
            miles = found[-1] if found else None
        size = f"{_format_miles(miles)}-mi radius" if miles else "radius"
        city = _radius_city(text)
        if city:
            street = _radius_street(text) if with_street else None
            return f"{street}, {city} {size}" if street else f"{city} {size}"
        bracket = re.search(r"\[([^\]]+)\]", text)
        if bracket and miles:
            return f"{bracket.group(1).strip()} {size}"
    return text


def audience_short_name(audience):
    """"LIFESTYLE Doctors and Nurses" -> "Doctors and Nurses"; "HLTH CRX DX
    Heart Conditions" -> "Heart Conditions". Catalog segment names lead with
    upper-case category codes that mean nothing to a client; they're
    dropped one leading token at a time, never past the last word. A
    multi-term stack shortens each term."""
    parts = [p.strip() for p in str(audience or "").split(",") if p.strip()]
    out = []
    for part in parts or [str(audience or "")]:
        words = part.split()
        while len(words) > 1 and words[0].isupper() and len(words[0]) >= 2 and words[0].isalpha():
            words = words[1:]
        out.append(" ".join(words))
    return ", ".join(w for w in out if w)


# ---------------------------------------------------------------------------
# Geography identity
# ---------------------------------------------------------------------------
def _digest(values):
    return hashlib.sha1(",".join(sorted(values)).encode()).hexdigest()[:12]


def geo_identity(group):
    """A stable string naming what a group actually targets: "dma:<keys>"
    for a market group, "zips:<digest>" for anything resolved to a ZIP list
    (an explicit list, a radius, a county option), else "label:<text>"."""
    geo_def = group.get("geo_def") or {}
    kind = geo_def.get("kind")
    if kind == "markets" and geo_def.get("markets"):
        return "dma:" + ",".join(sorted(geo_def["markets"]))
    zips = geo_def.get("zips") if kind == "zips" else None
    zips = zips or group.get("resolved_zips")
    if zips:
        return "zips:" + _digest(zips)
    return "label:" + tg.geo_label(group).strip().lower()


def is_dma(group):
    return geo_identity(group).startswith("dma:")


def canonical_geo_labels(groups, label_for=None):
    """{identity: label} -- one label per real geography. The shortest of
    the labels its lines currently carry, so "Eastern WV Delaware Southern
    Pennsylvania" wins over the long spelling of the same ZIP list."""
    labels = {}
    for group in groups or []:
        ident = geo_identity(group)
        label = tg.geo_label(group, label_for=label_for).strip()
        if not label:
            continue
        current = labels.get(ident)
        if current is None or len(label) < len(current):
            labels[ident] = label
    return labels


def unify_imported_geo_names(groups):
    """Give every imported line of one geography the same (shortest) name.

    Only touches a group whose `name` is still the cleaned name the import
    gave it (`avails_geo_label`) -- a rep's own rename is never overwritten.
    Mutates in place; returns the number of groups renamed."""
    by_ident = {}
    for group in groups or []:
        if group.get("avails_geo_label") and group.get("name") == group.get("avails_geo_label"):
            by_ident.setdefault(geo_identity(group), []).append(group)
    chosen = {ident: min((g["avails_geo_label"] for g in members), key=lambda s: (len(s), s))
              for ident, members in by_ident.items()}
    # Two DIFFERENT geographies that clean to one name (three DC radii, all
    # "Washington 5-mi radius" -- LiveWell) would read to a client as one
    # place. Those fall back to their street address.
    idents_by_label = {}
    for ident, label in chosen.items():
        idents_by_label.setdefault(label, []).append(ident)
    for label, idents in idents_by_label.items():
        if len(idents) < 2:
            continue
        for ident in idents:
            member = by_ident[ident][0]
            longer = clean_geo_label(member.get("avails_geo_raw") or label, with_street=True)
            chosen[ident] = longer or label
    renamed = 0
    for ident, members in by_ident.items():
        for g in members:
            if g["name"] != chosen[ident]:
                g["name"] = chosen[ident]
                g["avails_geo_label"] = chosen[ident]
                renamed += 1
    return renamed


# ---------------------------------------------------------------------------
# Duplicates
# ---------------------------------------------------------------------------
def _line_impressions(group):
    value = group.get("avails_doc_full_flight")
    if value is None:
        value = group.get("avails_full_flight")
    if value is None:
        value = group.get("avails_monthly")
    return int(value or 0)


def duplicate_groups(groups):
    """{duplicate_group_id: original_group_id} -- a line repeating an
    earlier one's audience, geography AND impressions exactly. The first
    occurrence is the original; every later one is the duplicate."""
    seen, dups = {}, {}
    for group in groups or []:
        audience = tg.audience_label(group).strip()
        if not audience:
            continue
        key = (audience, geo_identity(group), _line_impressions(group))
        if key in seen:
            dups[group["id"]] = seen[key]
        else:
            seen[key] = group["id"]
    return dups


def active_groups(groups, include_duplicates=False):
    """The groups the slide, the map and the totals use: everything with an
    audience, minus duplicate lines unless the rep has chosen to keep them."""
    dups = {} if include_duplicates else duplicate_groups(groups)
    return [g for g in (groups or [])
            if tg.audience_label(g).strip() and g["id"] not in dups]


# ---------------------------------------------------------------------------
# The grid
# ---------------------------------------------------------------------------
def needs_grid(groups):
    """True when the slide should show audience x geography: 2+ audiences
    or 2+ geographies."""
    audiences = {tg.audience_label(g) for g in groups if tg.audience_label(g).strip()}
    geos = {geo_identity(g) for g in groups if tg.audience_label(g).strip()}
    return len(audiences) >= 2 or len(geos) >= 2


def build_grid(groups, value_of, label_for=None):
    """{"audiences": [...], "geos": [(identity, label), ...],
    "cells": {(audience, identity): int}, "row_totals": {...},
    "col_totals": {...}, "total": int}. Order is first appearance. Two
    lines landing in one cell (a duplicate the rep chose to keep) sum."""
    labels = canonical_geo_labels(groups, label_for)
    audiences, geo_order, cells = [], [], {}
    for group in groups:
        audience = tg.audience_label(group).strip()
        if not audience:
            continue
        ident = geo_identity(group)
        if audience not in audiences:
            audiences.append(audience)
        if ident not in geo_order:
            geo_order.append(ident)
        cells[(audience, ident)] = cells.get((audience, ident), 0) + int(value_of(group) or 0)
    row_totals = {a: sum(v for (aa, _i), v in cells.items() if aa == a) for a in audiences}
    col_totals = {i: sum(v for (_a, ii), v in cells.items() if ii == i) for i in geo_order}
    return {
        "audiences": audiences,
        "geos": [(i, labels.get(i) or i.split(":", 1)[1]) for i in geo_order],
        "cells": cells, "row_totals": row_totals, "col_totals": col_totals,
        "total": sum(row_totals.values()),
    }


def slide_audience_name(audience, names=None):
    """What the slide calls an audience: the rep's own wording when they've
    set one (`names`, keyed by the catalog name), else the catalog name
    with its taxonomy prefixes dropped (audience_short_name)."""
    custom = str((names or {}).get(audience) or "").strip()
    return custom or audience_short_name(audience) or audience


def grid_payload(grid, fmt=lambda v: f"{v:,}", chip_colors=None, names=None):
    """The grid as plain strings for the slide -- {"geo_headers", "rows":
    [{"key", "audience", "chip", "cells": [...], "total"}], "totals": [...],
    "total"}. "—" where an audience doesn't run in a geography. "key" is
    the catalog name; "audience" is what the client reads
    (slide_audience_name)."""
    rows = []
    for audience in grid["audiences"]:
        cells = []
        for ident, _label in grid["geos"]:
            value = grid["cells"].get((audience, ident))
            cells.append(DASH if value is None else fmt(value))
        rows.append({
            "key": audience,
            "audience": slide_audience_name(audience, names),
            "chip": (chip_colors or {}).get(audience),
            "cells": cells,
            "total": fmt(grid["row_totals"][audience]),
        })
    return {
        "geo_headers": [label for _i, label in grid["geos"]],
        "rows": rows,
        "totals": [fmt(grid["col_totals"][i]) for i, _l in grid["geos"]],
        "total": fmt(grid["total"]),
    }


# ---------------------------------------------------------------------------
# Plan checks
# ---------------------------------------------------------------------------
# A ZIP with no other ZIP of its own list within this many miles is a typo
# or a paste from another plan, not part of the area. Nearest-neighbour,
# not distance from the middle: a real list can legitimately stretch from
# Delaware to southwestern Pennsylvania, but every ZIP in it still sits
# near other ZIPs in it. 02191 (Quincy, MA) in a Silver Spring list is
# ~350 miles from its nearest neighbour.
_OUT_OF_REGION_MILES = 100
# A suggested correction has to land at least this close to the list.
_SUGGESTION_MILES = 50


def _miles(a, b):
    import math
    lat1, lon1, lat2, lon2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = (math.sin((lat2 - lat1) / 2) ** 2
         + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2)
    return 3958.8 * 2 * math.asin(math.sqrt(h))


def _nearest(point, others):
    return min((_miles(point, o) for o in others), default=None)


def suspect_zips(zips, zip_points):
    """([invalid], [(zip, suggestion_or_None)]) for one ZIP list.

    Invalid: not five digits. (A well-formed ZIP missing from the crosswalk
    is usually a real PO-box ZIP -- 25402 Martinsburg -- and is not flagged.)
    Out of region: no other ZIP in the list within `_OUT_OF_REGION_MILES`.
    The suggestion is a one-digit change that lands back inside the list's
    area -- 02191 -> 22191 (Woodbridge, VA)."""
    invalid = [z for z in zips if not (len(str(z)) == 5 and str(z).isdigit())]
    located = {z: zip_points[z] for z in zips if z in zip_points}
    if len(located) < 3:
        return invalid, []
    out = []
    for z, point in located.items():
        others = [p for zz, p in located.items() if zz != z]
        if _nearest(point, others) <= _OUT_OF_REGION_MILES:
            continue
        center = (median(p[0] for p in others), median(p[1] for p in others))
        suggestion, best = None, None
        for pos in range(5):
            for digit in "0123456789":
                cand = z[:pos] + digit + z[pos + 1:]
                if cand == z or cand in located or cand not in zip_points:
                    continue
                if _nearest(zip_points[cand], others) > _SUGGESTION_MILES:
                    continue
                d = _miles(center, zip_points[cand])
                if best is None or d < best:
                    best, suggestion = d, cand
        out.append((z, suggestion))
    return invalid, out


def plan_checks(groups, zip_points, label_for=None, include_duplicates=False):
    """Plain-language warnings for the rep, grouped by kind:
    {"duplicates": [...], "zips": [...], "names": [...], "overlaps": [...],
    "excluded_zips": set()}. Never shown on the slide.

    `excluded_zips` are the out-of-region/invalid ZIPs, for the map to leave
    out of its frame (one stray Massachusetts ZIP must not zoom a DC map out
    to New England)."""
    groups = [g for g in (groups or []) if tg.audience_label(g).strip()]
    labels = canonical_geo_labels(groups, label_for)
    by_id = {g["id"]: g for g in groups}
    out = {"duplicates": [], "zips": [], "names": [], "overlaps": [], "excluded_zips": set()}

    for dup_id, orig_id in duplicate_groups(groups).items():
        dup, orig = by_id[dup_id], by_id[orig_id]
        ids = " / ".join(x for x in (orig.get("avails_rfpi"), dup.get("avails_rfpi")) if x)
        state = "included" if include_duplicates else "left out of the table and map"
        out["duplicates"].append(
            f"Duplicate line{(' ' + ids) if ids else ''}: {audience_short_name(tg.audience_label(dup))} "
            f"in {labels.get(geo_identity(dup), tg.geo_label(dup))} at "
            f"{_line_impressions(dup):,} -- same audience, geography and impressions as "
            f"another line. Currently {state}.")

    seen_idents = set()
    for group in groups:
        ident = geo_identity(group)
        if ident in seen_idents or not ident.startswith("zips:"):
            continue
        seen_idents.add(ident)
        zips = (group.get("geo_def") or {}).get("zips") or group.get("resolved_zips") or []
        invalid, far = suspect_zips(zips, zip_points)
        label = labels.get(ident, tg.geo_label(group))
        for z in invalid:
            out["zips"].append(f"{label}: \"{z}\" isn't a valid ZIP code.")
            out["excluded_zips"].add(z)
        for z, suggestion in far:
            hint = f" -- likely meant {suggestion}" if suggestion else ""
            out["zips"].append(f"{label}: ZIP {z} is far outside this area{hint}. "
                               f"Left off the map; check it with whoever built the plan.")
            out["excluded_zips"].add(z)

    variants = {}
    for group in groups:
        variant = group.get("avails_geo_variant")
        if variant:
            variants.setdefault(geo_identity(group), set()).add(variant)
    for ident, names in variants.items():
        if len(names) > 1:
            shown = labels.get(ident) or min(names, key=len)
            listed = " / ".join(f"\"{n}\"" for n in sorted(names, key=len))
            out["names"].append(f"One ZIP list is named {len(names)} ways ({listed}) -- "
                                f"shown as \"{shown}\".")

    # One line per pair of geographies, naming every audience that runs in
    # both -- not one line per audience, which said the same thing three
    # times on Adventist.
    zips_of, auds_of = {}, {}
    for group in groups:
        ident = geo_identity(group)
        zips_of.setdefault(ident, set()).update(group.get("resolved_zips") or [])
        auds_of.setdefault(ident, [])
        aud = tg.audience_label(group)
        if aud not in auds_of[ident]:
            auds_of[ident].append(aud)
    idents = list(zips_of)
    for i in range(len(idents)):
        for j in range(i + 1, len(idents)):
            a, b = idents[i], idents[j]
            shared = zips_of[a] & zips_of[b]
            both = [x for x in auds_of[a] if x in auds_of[b]]
            if not shared or not both:
                continue
            names = ", ".join(audience_short_name(x) for x in both)
            out["overlaps"].append(
                f"{labels.get(a, a)} and {labels.get(b, b)} share {len(shared):,} "
                f"ZIP{'s' if len(shared) != 1 else ''}, and {names} "
                f"{'runs' if len(both) == 1 else 'run'} in both -- those households can be "
                f"reached by two lines.")
    return out


def reconcile(groups, documents, include_duplicates=False):
    """None when the table's full-flight total ties to the documents'
    stated totals, else a sentence for the rep. Only lines that came from an
    uploaded document count; a duplicate left out is added back before
    comparing, since leaving it out is a choice, not a discrepancy."""
    if not documents:
        return None
    sources = {d["rfpid"] for d in documents}
    imported = [g for g in groups or [] if g.get("avails_source") in sources]
    if not imported:
        return None
    stated = sum(int(d.get("total_impressions") or 0) for d in documents)
    shown = sum(_line_impressions(g) for g in active_groups(imported, include_duplicates))
    left_out = 0 if include_duplicates else sum(
        _line_impressions(g) for g in imported if g["id"] in duplicate_groups(imported))
    if shown + left_out == stated:
        return None
    return (f"The avails table adds up to {shown + left_out:,} impressions, but the uploaded "
            f"document(s) state {stated:,}. Check for a line that was edited or removed.")


# ---------------------------------------------------------------------------
# Map color mode
# ---------------------------------------------------------------------------
MAP_MODE_AUTO = "Auto"
MAP_MODE_AUDIENCE = "Audience"
MAP_MODE_REGION = "Region"
MAP_MODES = [MAP_MODE_AUTO, MAP_MODE_AUDIENCE, MAP_MODE_REGION]
MAX_AUDIENCE_COMBOS = 4


def resolve_map_mode(choice, n_audiences, n_regions, n_combos):
    """Auto: 2+ audiences -> Audience, unless that would take more than
    four combination colors -> Region; one audience over 2+ regions ->
    Region; otherwise Audience (one color)."""
    if choice in (MAP_MODE_AUDIENCE, MAP_MODE_REGION):
        return choice
    if n_audiences >= 2:
        return MAP_MODE_REGION if n_combos > MAX_AUDIENCE_COMBOS else MAP_MODE_AUDIENCE
    if n_regions >= 2:
        return MAP_MODE_REGION
    return MAP_MODE_AUDIENCE
