"""
audience_catalog.py -- the canonical Premion audience-segment catalog.

The catalog lives in Supabase's `audiences` table and is read through db.py.
The brochure PDF (`PREMION_Audience Targeting (1).pdf`) is where that table
was originally derived from, and remains the local fallback for when
Supabase is unreachable -- so the parser below is still live code, not
history. `audience_segments_derived.csv` only contributes `times_used`
counts; its own `category`/`is_custom` columns are not used.

In the PDF, RFP-selectable segments are highlighted with a light-cyan
background behind their line; everything else is a "custom" (non-RFP)
segment. Detecting this requires per-line rectangle-overlap checks against
the PDF's fill shapes -- see `extract_page` below.
"""

import json
from pathlib import Path

import pandas as pd
import streamlit as st

import db

# pdfplumber is imported lazily, inside the parser. It's only needed for the
# local fallback, it pulls in a sizeable dependency tree, and on a deployed
# instance the brochure PDF isn't present at all (it's gitignored) -- so
# importing it at module scope would cost startup memory for a code path that
# can never run there.

CATALOG_COLUMNS = ["segment", "category", "subcategory", "rfp_selectable", "times_used", "impressions"]

PDF_PATH = Path(__file__).parent / "PREMION_Audience Targeting (1).pdf"
TIMES_USED_CSV_PATH = Path(__file__).parent / "audience_segments_derived.csv"

CATEGORY_PREFIXES = ["HH", "DEMO", "POL", "LIFESTAGE", "AUTO", "ENT", "FIN", "FOOD",
                     "HLTH", "LIFESTYLE", "RETAIL", "SPORTS", "TRAVEL", "AFIRST",
                     # Added for workbook-only segments -- the brochure's own 14
                     # categories above never needed these, and neither category
                     # collides with anything the brochure already has (FIN Legal
                     # Services stays FIN; LEGAL is exclusively for the custom,
                     # non-RFP-selectable personal-injury/collision-lawyer segments
                     # a rep would never find a matching brochure entry for).
                     "B2B", "LEGAL"]

# Subcategory headers spell some categories out in full ("FINANCE", "HEALTH",
# "ENTERTAINMENT") while segment rows use the abbreviated category code
# ("FIN", "HLTH", "ENT"). Maps a header's own first word to the matching
# segment-row category code so both can be looked up under the same key.
HEADER_CATEGORY_ALIASES = {
    "FINANCE": "FIN",
    "HEALTH": "HLTH",
    "ENTERTAINMENT": "ENT",
    "AUDIENCE": "AFIRST",
    "POLITICAL": "POL",
}

HIGHLIGHT_COLOR = (0.875, 0.992, 1.0)  # RFP-selectable line background
HEADER_COLOR = (0.282, 0.376, 0.973)   # subcategory header text (blue)
WHITE = (1.0, 1.0, 1.0)                # top-level category banner text
GREEN = (0.4, 0.867, 0.098)            # "AND MORE!" marker, not real content

# Fixed absolute column bands (points from page left edge). The brochure uses
# 4 columns per page at consistent positions (~42, 175, 308, 441pt starts on
# a 612pt-wide page) -- banding by absolute x0 range (rather than clustering
# every word's own x0) is what correctly keeps a multi-word line's
# continuation words in the same column as their line-starting prefix,
# instead of splitting each line into one column per word.
COLUMN_BANDS = [(0, 145), (145, 278), (278, 411), (411, 10000)]

SEGMENT_PAGES = (2, 3, 4)  # 0-indexed: pages 3-5 of the PDF


def _close(c1, c2, tol=0.02):
    if c1 is None or c2 is None:
        return False
    return all(abs(a - b) < tol for a, b in zip(c1, c2))


def _band_of(x0):
    for i, (lo, hi) in enumerate(COLUMN_BANDS):
        if lo <= x0 < hi:
            return i
    return len(COLUMN_BANDS) - 1


def _header_category(first_word):
    if first_word in CATEGORY_PREFIXES:
        return first_word
    return HEADER_CATEGORY_ALIASES.get(first_word, first_word)


def _group_lines(words):
    """Group words into (column-band, line) buckets, return list of lines
    each with its words sorted left-to-right, sorted overall by band then top."""
    buckets = {}
    for w in words:
        col = _band_of(w["x0"])
        buckets.setdefault(col, []).append(w)

    lines = []
    for col, ws in buckets.items():
        ws.sort(key=lambda w: w["top"])
        cur = []
        cur_top = None
        for w in ws:
            if cur and abs(w["top"] - cur_top) > 2.5:
                cur.sort(key=lambda w: w["x0"])
                lines.append({"col": col, "top": cur_top, "words": cur})
                cur = []
                cur_top = None
            cur.append(w)
            cur_top = w["top"] if cur_top is None else min(cur_top, w["top"])
        if cur:
            cur.sort(key=lambda w: w["x0"])
            lines.append({"col": col, "top": cur_top, "words": cur})
    lines.sort(key=lambda l: (l["col"], l["top"]))
    return lines


def _line_bbox(line):
    x0 = min(w["x0"] for w in line["words"])
    x1 = max(w["x1"] for w in line["words"])
    top = min(w["top"] for w in line["words"])
    bottom = max(w["bottom"] for w in line["words"])
    return x0, top, x1, bottom


def _overlaps(bbox, rect, tol=1.5):
    x0, top, x1, bottom = bbox
    return not (rect["x1"] < x0 - tol or rect["x0"] > x1 + tol or
                rect["bottom"] < top - tol or rect["top"] > bottom + tol)


def extract_page(page):
    """Parse one detailed-catalog page (3-5) into segment rows."""
    words = page.extract_words(extra_attrs=["size", "fontname", "non_stroking_color"])
    # The disclaimer/footnote paragraphs (~4pt) have broken/duplicate 'top'
    # values in this PDF (many words across different visual lines all
    # report the same top), which corrupts line-grouping. Segment names are
    # 6pt and subcategory headers are 5pt, so dropping anything smaller
    # removes the footnote junk without losing real content.
    words = [w for w in words if w["size"] > 4.5]
    highlight_rects = [r for r in page.rects if _close(r.get("non_stroking_color"), HIGHLIGHT_COLOR)]
    lines = _group_lines(words)

    rows = []
    # A category's subcategory list snakes across columns (e.g. AUTO MAKE
    # starts in column 0, continues unlabeled into columns 1 and 2), while
    # unrelated categories' own sections interleave with it in raw (col, top)
    # order. Tracking the current subcategory per-category (rather than as
    # one running value, or keyed by column) means each category's rows pick
    # up its own most recent header regardless of what other categories'
    # headers appeared in between.
    current_subcat = {}
    for line in lines:
        ws = line["words"]
        sizes = {round(w["size"], 1) for w in ws}
        colors = [w["non_stroking_color"] for w in ws]
        joined = " ".join(w["text"] for w in ws)

        if all(_close(c, WHITE) for c in colors):
            continue  # top-level category banner (white on dark purple)
        if any(_close(c, GREEN) for c in colors):
            continue  # "AND MORE!" marker
        if any(abs(s - 5.0) < 0.3 for s in sizes) and any(_close(c, HEADER_COLOR) for c in colors):
            header_first_word = joined.split()[0] if joined.split() else ""
            current_subcat[_header_category(header_first_word)] = joined
            continue
        if any(abs(s - 4.0) < 0.3 for s in sizes):
            continue  # tiny footnote/legend text

        first_word = joined.split()[0] if joined.split() else ""
        if first_word not in CATEGORY_PREFIXES:
            continue  # legend text, footnotes, etc -- not a segment line

        bbox = _line_bbox(line)
        rfp = any(_overlaps(bbox, r) for r in highlight_rects)
        rows.append({
            "segment": joined,
            "category": first_word,
            "subcategory": current_subcat.get(first_word, ""),
            "rfp_selectable": rfp,
        })
    return rows


def _parse_pdf_catalog() -> pd.DataFrame:
    import pdfplumber  # lazy -- see the note at the top of this module

    rows = []
    with pdfplumber.open(PDF_PATH) as pdf:
        for pidx in SEGMENT_PAGES:
            rows.extend(extract_page(pdf.pages[pidx]))

    df = pd.DataFrame(rows)
    # A handful of segment names appear twice in the brochure (once under
    # each of two adjacent subcategory blurbs) with different highlight
    # status -- a segment counts as RFP-selectable if any occurrence is.
    df = (
        df.groupby("segment", as_index=False)
        .agg({"category": "first", "subcategory": "first", "rfp_selectable": "max"})
    )
    df["rfp_selectable"] = df["rfp_selectable"].astype(bool)
    return df


def _merge_times_used(catalog: pd.DataFrame) -> pd.DataFrame:
    usage = pd.read_csv(TIMES_USED_CSV_PATH)[["segment", "times_used"]]
    merged = catalog.merge(usage, on="segment", how="left")
    merged["times_used"] = merged["times_used"].fillna(0).astype(int)
    return merged


def _sorted(catalog: pd.DataFrame) -> pd.DataFrame:
    return catalog.sort_values(["category", "subcategory", "segment"]).reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_local_catalog() -> pd.DataFrame:
    """The fallback catalog, parsed straight from the brochure PDF and merged
    with the times_used CSV.

    Returns an **empty** catalog (right columns, no rows) when the brochure
    isn't on disk, rather than raising. The PDF is gitignored, so a deployed
    instance never has it -- and this runs at import time, so raising here
    would take the whole app down at startup the first time Supabase was
    briefly unreachable, instead of degrading to a form with no audience
    picker.

    `impressions` is always 0 here -- the local fallback has no workbook
    data to draw on, only the brochure PDF and the times_used CSV, so
    ranking degrades to "everything ties" rather than raising on a missing
    column. Supabase is what actually carries impressions (see
    `_load_catalog` below); this path only exists for when Supabase itself
    is unreachable.
    """
    if not PDF_PATH.exists():
        return pd.DataFrame(columns=CATALOG_COLUMNS)
    catalog = _merge_times_used(_parse_pdf_catalog())
    catalog["impressions"] = 0
    return _sorted(catalog)


def _metrics_impressions(metrics):
    """The `impressions` key out of an `audiences.metrics` jsonb value --
    tolerant of it arriving as a dict (the normal case), a JSON string (if a
    client library ever stops auto-decoding jsonb), or missing/None."""
    if isinstance(metrics, str):
        try:
            metrics = json.loads(metrics)
        except (TypeError, ValueError):
            metrics = {}
    if not isinstance(metrics, dict):
        return 0
    try:
        return int(metrics.get("impressions") or 0)
    except (TypeError, ValueError):
        return 0


@st.cache_data(ttl=600, show_spinner=False)
def _load_catalog():
    """(catalog, warning). Cached with a TTL rather than for the process
    lifetime so a segment added in Supabase reaches a long-running session."""
    rows, warning = db.fetch_audiences()
    if rows is None:
        local = load_local_catalog()
        if local.empty:
            return local, (f"{warning}, and the local brochure PDF isn't available either. "
                           f"The audience catalog is empty for now -- avails rows can still be "
                           f"typed in by hand, but browse/suggest have nothing to offer. "
                           f"Everything else works normally.")
        return local, (f"{warning}. Using the local brochure PDF instead -- "
                       f"segments added since it was published won't appear.")
    catalog = pd.DataFrame(rows)
    # metrics is a jsonb bag (impressions today, room for a second ranking
    # metric later with no migration) -- flattened into its own column
    # BEFORE the defaulting loop below, which would otherwise blanket-zero
    # "impressions" for every row since it's never a literal key in `rows`.
    if "metrics" in catalog.columns:
        catalog["impressions"] = catalog["metrics"].apply(_metrics_impressions)
    for column in CATALOG_COLUMNS:
        if column not in catalog:
            catalog[column] = "" if column in ("category", "subcategory") else 0
    catalog["rfp_selectable"] = catalog["rfp_selectable"].fillna(False).astype(bool)
    catalog["times_used"] = catalog["times_used"].fillna(0).astype(int)
    catalog["impressions"] = catalog["impressions"].fillna(0).astype(int)
    return _sorted(catalog[CATALOG_COLUMNS]), None


def load_audience_catalog() -> pd.DataFrame:
    """The full audience catalog: segment, category, subcategory,
    rfp_selectable (bool), times_used (int)."""
    return _load_catalog()[0]


def clear_catalog_cache():
    """Drop the cached catalog -- called after an audience-usage workbook
    activation so the very next render picks up the merge instead of
    waiting out the 10-minute TTL. Same role `db.clear_deck_cache()` plays
    for the master deck."""
    _load_catalog.clear()


def catalog_warning():
    """The fallback warning for the catalog currently in use, or None when it
    came from Supabase. Kept separate from load_audience_catalog() so its many
    call sites don't all have to unpack a tuple."""
    return _load_catalog()[1]


def validate_segments(names):
    """Check a list of segment names against the catalog. Returns
    (matched, unmatched), each a list in the same order as `names`."""
    catalog = load_audience_catalog()
    valid = set(catalog["segment"])
    matched = [n for n in names if n in valid]
    unmatched = [n for n in names if n not in valid]
    return matched, unmatched
