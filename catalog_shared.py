"""Catalog and case-study logic shared by the proposal builder (app.py) and
the finders-only app (finders_app.py): the vertical tables, the audience
catalog ranking and Claude suggest prompts, and the case-study naming and
suggest prompt. Pure logic, no page rendering, and nothing from the builder
(assembly, slide_map, report builder) -- one implementation, two entry
points (2026-10-05). app.py re-imports every name here.
"""
import json
import re
from pathlib import Path

import pandas as pd
import streamlit as st

from audience_catalog import category_matches, load_audience_catalog, validate_segments
from claude_client import _call_claude_json

STREAMING_RETARGETING_TARGETING = "Retarget Exposed CTV Viewers"


VERTICALS = {
    "None": "none",
    "Education": "education",
    "Healthcare": "healthcare",
    "Retail": "retail",
    "Travel & Tourism": "travel",
    "Home Improvement": "home_improvement",
    "Banking & Finance": "banking",
    "Entertainment": "entertainment",
    "Casual Dining & QSR": "dining_qsr",
    "Automotive": "auto",
    "Legal": "legal",
}


# ---------------------------------------------------------------------------
# FALLBACK ONLY -- not the live rate card.
#
# The real product list and CPMs live in Supabase's `products` table and are
# loaded by load_rate_card() below; edit them THERE, not here. These copies
# exist purely so the app still runs (with a visible warning) when Supabase
# is unreachable. They will drift from the table over time and that is
# expected -- they are a safety net, not a source of truth.
#
# All Audience Marketplace Display tactics are $5.50 except Geofencing
# ($9.00); all AM Pre-Roll tactics are $21.00. Sports rates came from
# PREMION_Live Sports Rates.xlsx ("2026 Prem Core Live Sports" sheet, TEGNA
# Recommended Rate column).
# ---------------------------------------------------------------------------
FALLBACK_PRODUCTS = {
    "premion_streaming_tv": {"label": "Premion Streaming TV", "default_cpm": 32.00, "line_type": "premion"},
    "streaming_retargeting_display": {"label": "Streaming Retargeting - Display", "default_cpm": 5.50, "line_type": "premion", "targeting_copy": STREAMING_RETARGETING_TARGETING},
    "streaming_retargeting_preroll": {"label": "Streaming Retargeting - Pre-Roll", "default_cpm": 21.00, "line_type": "premion", "targeting_copy": STREAMING_RETARGETING_TARGETING},
    "audience_targeting_display": {"label": "Audience Targeting - Display", "default_cpm": 5.50, "line_type": "am"},
    "audience_targeting_preroll": {"label": "Audience Targeting - Pre-Roll", "default_cpm": 21.00, "line_type": "am"},
    "geofencing_display": {"label": "Geofencing - Display", "default_cpm": 9.00, "line_type": "am"},
    "geofencing_preroll": {"label": "Geofencing - Pre-Roll", "default_cpm": 21.00, "line_type": "am"},
    "site_retargeting_display": {"label": "Site Retargeting - Display", "default_cpm": 5.50, "line_type": "am"},
    "site_retargeting_preroll": {"label": "Site Retargeting - Pre-Roll", "default_cpm": 21.00, "line_type": "am"},
    "broadcast_tv": {"label": "Broadcast Schedule", "default_cpm": 5.50, "line_type": "broadcast"},
}


# Coarse category hints used to pick a relevant slice of the audience catalog
# to send Claude, based on a cheap local keyword guess at the vertical --
# Claude's own returned "vertical" field is authoritative either way, this
# just keeps the prompt from having to carry the full ~370-segment catalog.
# Also what the Audience finder's own vertical-default filter reads (a
# DEFAULT, not a restriction -- clearing it shows the whole catalog, and
# search always searches everything regardless).
#
# Data, not code, on purpose: this needs tuning as real proposals surface
# gaps, and a rep shouldn't need a code change to fix "the legal vertical
# doesn't show LEGAL first". See vertical_category_map.csv's own header for
# the rank convention.
VERTICAL_CATEGORY_MAP_PATH = Path(__file__).parent / "vertical_category_map.csv"


@st.cache_data(show_spinner=False)
def load_vertical_category_map(path=VERTICAL_CATEGORY_MAP_PATH):
    """{vertical: [category, ...]}, each list ordered by the file's own
    `rank` column (lower = shown first). Missing file degrades to an empty
    map -- every vertical falls back to `prioritize_catalog`'s own
    "no hint" behavior (times_used order, nothing pinned to the front)
    rather than raising, same policy every local-fallback loader in this
    app follows.
    """
    if not path.exists():
        return {}
    import csv as _csv
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as handle:
        for row in _csv.DictReader(handle):
            vertical = (row.get("vertical") or "").strip()
            category = (row.get("category") or "").strip()
            if not vertical or not category:
                continue
            try:
                rank = int(row.get("rank") or 0)
            except ValueError:
                rank = 0
            rows.append((vertical, category, rank))
    mapping = {}
    for vertical, category, rank in sorted(rows, key=lambda r: (r[0], r[2])):
        mapping.setdefault(vertical, []).append(category)
    return mapping


VERTICAL_CATEGORY_HINTS = load_vertical_category_map()


# What a business actually calls itself, mapped to its vertical. Discovery
# notes say "an HVAC company" or "a credit union", never "home_improvement"
# or "banking", so without these a whole category of notes drafts with no
# vertical hint at all -- which costs the audience slice its relevance.
#
# Matched on **word boundaries**, not substrings (see _mentions). That's what
# makes short names like "tire" and "spa" safe: naive `in` matching fires
# "tire" on "the entire campaign", "venue" on "revenue", "spa" on "Spanish"
# and -- a live bug this fixes -- "auto" on "automatic".
#
# Deliberately absent: apartment complexes, realtors and property management.
# They're a real category but none of the verticals below actually fits them
# (home improvement is contractors, not property sales), and a wrong hint is
# worse than none -- it sends Claude the wrong slice of the audience catalog.
# They need a Real Estate vertical, not a synonym.
VERTICAL_HINT_SYNONYMS = {
    # --- home improvement: the trades -------------------------------------
    "hvac": "home_improvement", "heating and cooling": "home_improvement",
    "air conditioning": "home_improvement", "roofing": "home_improvement",
    "roofer": "home_improvement", "plumbing": "home_improvement",
    "plumber": "home_improvement", "pest control": "home_improvement",
    "exterminator": "home_improvement", "landscaping": "home_improvement",
    "lawn care": "home_improvement", "siding": "home_improvement",
    "window replacement": "home_improvement", "replacement windows": "home_improvement",
    "gutters": "home_improvement", "remodeling": "home_improvement",
    "remodeler": "home_improvement", "home services": "home_improvement",
    "general contractor": "home_improvement", "flooring": "home_improvement",
    "kitchen and bath": "home_improvement", "garage door": "home_improvement",
    "solar": "home_improvement", "fencing": "home_improvement",
    "restoration": "home_improvement", "deck replacement": "home_improvement",
    "decking": "home_improvement", "deck builder": "home_improvement",
    # --- automotive --------------------------------------------------------
    "dealership": "auto", "car dealer": "auto", "auto dealer": "auto",
    "dealer group": "auto", "auto group": "auto", "body shop": "auto",
    "collision center": "auto", "auto repair": "auto", "tire": "auto",
    "car wash": "auto", "powersports": "auto", "rv dealer": "auto",
    # --- healthcare --------------------------------------------------------
    "hospital": "healthcare", "clinic": "healthcare", "medical": "healthcare",
    "med spa": "healthcare", "medspa": "healthcare", "dental": "healthcare",
    "dentist": "healthcare", "orthodontist": "healthcare",
    "urgent care": "healthcare", "physician": "healthcare",
    "primary care": "healthcare", "dermatology": "healthcare",
    "chiropractor": "healthcare", "optometrist": "healthcare",
    "health system": "healthcare", "surgery center": "healthcare",
    "physical therapy": "healthcare", "hearing aid": "healthcare",
    "home health": "healthcare", "senior living": "healthcare",
    "veterinary": "healthcare", "pediatric": "healthcare",
    # --- banking & finance -------------------------------------------------
    "bank": "banking", "credit union": "banking", "financial advisor": "banking",
    "wealth management": "banking", "mortgage lender": "banking",
    "insurance agency": "banking", "investment firm": "banking",
    "tax service": "banking", "accounting firm": "banking",
    # --- dining & QSR ------------------------------------------------------
    "restaurant": "dining_qsr", "qsr": "dining_qsr", "fast food": "dining_qsr",
    "pizzeria": "dining_qsr", "pizza": "dining_qsr", "cafe": "dining_qsr",
    "coffee shop": "dining_qsr", "brewery": "dining_qsr", "diner": "dining_qsr",
    "steakhouse": "dining_qsr", "taqueria": "dining_qsr", "food truck": "dining_qsr",
    "catering": "dining_qsr", "bar and grill": "dining_qsr",
    # --- retail ------------------------------------------------------------
    "furniture store": "retail", "jewelry": "retail", "jeweler": "retail",
    "mattress": "retail", "appliance store": "retail", "boutique": "retail",
    "grocery": "retail", "supermarket": "retail", "garden center": "retail",
    "sporting goods": "retail", "hardware store": "retail",
    "department store": "retail", "pharmacy": "retail",
    # --- travel & tourism --------------------------------------------------
    "hotel": "travel", "tourism": "travel", "resort": "travel", "casino": "travel",
    "cruise": "travel", "bed and breakfast": "travel", "campground": "travel",
    "visitors bureau": "travel", "convention and visitors": "travel",
    # --- entertainment -----------------------------------------------------
    "movie theater": "entertainment", "cinema": "entertainment",
    "theatre": "entertainment", "concert venue": "entertainment",
    "festival": "entertainment", "county fair": "entertainment",
    "amusement park": "entertainment", "theme park": "entertainment",
    "museum": "entertainment", "zoo": "entertainment",
    "event venue": "entertainment", "bowling": "entertainment",
    # --- education ---------------------------------------------------------
    "school": "education", "university": "education", "college": "education",
    "trade school": "education", "technical college": "education",
    "career training": "education", "academy": "education",
    "tutoring": "education", "charter school": "education",
    # --- legal -------------------------------------------------------------
    "law firm": "legal", "attorney": "legal", "lawyer": "legal",
    "personal injury": "legal", "workers comp": "legal", "law office": "legal",
    "legal services": "legal", "criminal defense": "legal",
    "family law": "legal", "estate planning": "legal", "bankruptcy": "legal",
}


# Shared between the draft-from-notes prompt and the audience-finder Suggest
# prompt so the two paths hold audience matches to the same bar -- both send
# the same catalog slice already, this keeps the matching *instruction* the
# same too, rather than letting a broader multi-field drafting task reason
# less carefully about segment precision than the finder's single-purpose one.
AUDIENCE_MATCH_GUIDANCE = (
    "When multiple catalog segments could plausibly apply, prefer the most specific one over a "
    "generic demographic proxy -- e.g. if the notes describe interest in a particular product or "
    "service (\"in-market for deposit accounts\", \"shopping for a used car\"), prefer a segment "
    "naming that specific interest over a loosely-correlated demographic segment (like a general "
    "homeowner or income segment) that merely correlates with it."
)


def _mentions(haystack, phrase):
    """Whole-word/phrase match. Substring matching misfires badly on the short
    names below -- "tire" inside "entire", "venue" inside "revenue", "auto"
    inside "automatic"."""
    return re.search(rf"\b{re.escape(phrase)}\b", haystack) is not None


# A vertical's own key that's too generic to match on. "auto" is a word in
# its own right -- "auto loans" belongs to a credit union, not a dealership --
# and the label "Automotive" carries the same meaning unambiguously.
AMBIGUOUS_VERTICAL_TERMS = frozenset({"auto"})


def _vertical_match_table():
    """{phrase: vertical} over the trade names *and* the verticals' own names
    and labels, so a newly added vertical is matchable without anyone having
    to write synonyms for it first."""
    table = dict(VERTICAL_HINT_SYNONYMS)
    for label, key in VERTICALS.items():
        if key == "none":
            continue
        for term in (key.replace("_", " "), label.lower()):
            if term not in AMBIGUOUS_VERTICAL_TERMS:
                table.setdefault(term, key)
    return table


def _detect_vertical_hint(notes):
    """Guess a vertical from discovery notes, for the audience-catalog slice.

    A *hint* only: it sorts the catalog and seeds the draft, and Claude's own
    returned vertical is authoritative either way. Returning None is fine and
    much better than returning the wrong one -- a bad hint sends the model the
    wrong slice of the catalog.

    Longest phrase wins, over one combined table rather than names-then-trades.
    That's what makes "credit union promoting auto loans" resolve to banking
    instead of the incidental "auto", and "med spa" beat "spa". A two-pass
    version that checked vertical names first got that case wrong.
    """
    low = notes.lower()
    table = _vertical_match_table()
    for phrase in sorted(table, key=len, reverse=True):
        if _mentions(low, phrase):
            return table[phrase]
    return None


def prioritize_catalog(catalog, vertical_hint):
    """Sort the catalog with a vertical's own categories first, then
    everything else, each by total delivered impressions -- the popularity
    signal ("rank by impressions, not by count": a segment booked once at
    huge volume is more relevant to surface than one booked five times at
    a trickle, which times_used alone can't distinguish). A vertical only
    *prioritizes* -- it never filters anything out, so a segment outside
    the vertical's categories is still reachable, just further down."""
    if vertical_hint and vertical_hint in VERTICAL_CATEGORY_HINTS:
        cats = VERTICAL_CATEGORY_HINTS[vertical_hint]
        # A boolean mask and its complement are a PARTITION -- every row
        # lands in exactly one side, never both -- so a dual-category
        # segment relevant under either of the vertical's categories still
        # appears exactly once in the concatenated result, not twice.
        is_relevant = catalog["category"].apply(lambda c: category_matches(c, cats))
        relevant = catalog[is_relevant].sort_values("impressions", ascending=False)
        rest = catalog[~is_relevant].sort_values("impressions", ascending=False)
        return pd.concat([relevant, rest])
    return catalog.sort_values("impressions", ascending=False)


def build_catalog_slice(vertical_hint, cap=150):
    """The slice of the catalog sent to Claude. With a vertical hint the cap
    is a real economy measure -- the hint's own categories sort to the front,
    so the cut only drops far-less-relevant segments. With no hint there's
    nothing to sort by relevance, so capping would cut arbitrarily; send the
    whole catalog instead (it's ~370 segments, well within prompt budget)."""
    catalog = load_audience_catalog()
    combined = prioritize_catalog(catalog, vertical_hint)
    if not vertical_hint or vertical_hint not in VERTICAL_CATEGORY_HINTS:
        cap = len(combined)
    sliced = combined.head(cap)
    return sliced[["segment", "category", "subcategory", "rfp_selectable",
                   "times_used", "impressions"]].to_dict("records")


def build_audience_finder_prompt(description, vertical_hint=None):
    vertical_hint = _detect_vertical_hint(description) or vertical_hint
    catalog_slice = build_catalog_slice(vertical_hint, cap=150)
    return f"""You are recommending Premion audience-targeting segments for a CTV/OTT ad campaign, based on a description of the client or campaign. Return ONLY valid JSON -- no markdown code fences, no preamble, no explanation, just the JSON object -- matching this schema:

{{"recommendations": [{{"segment": "exact catalog name", "rationale": "one-line reason this fits"}}]}}

Rules:
- "segment" must be an EXACT name from the audience catalog slice below -- do not paraphrase or invent names. If nothing in the slice fits well, return fewer recommendations rather than a poor match. {AUDIENCE_MATCH_GUIDANCE}
- Recommend at most 8 segments, ranked most-relevant first.
- "rationale" is one short sentence.

Audience catalog slice -- {len(catalog_slice)} of {len(load_audience_catalog())} total segments (JSON): {json.dumps(catalog_slice)}

Client/campaign description:
\"\"\"
{description}
\"\"\"
"""


def call_claude_audience_suggest(description, vertical_hint=None):
    """Returns (recommendations_list, unmatched_names, error_message) --
    error_message is None on success. Each recommendation is
    {"segment", "rationale"}, already validated against the catalog."""
    parsed, error = _call_claude_json(
        build_audience_finder_prompt(description, vertical_hint), label="audience_suggest")
    if error:
        return None, None, error

    raw_recs = parsed.get("recommendations", []) or []
    names = [r.get("segment", "") for r in raw_recs if r.get("segment")]
    matched, unmatched = validate_segments(names)
    recs = [r for r in raw_recs if r.get("segment") in matched]
    return recs, unmatched, None


# ---------------- Case study vault ----------------
# Coarse product tags a case study gets labelled with. The real PRODUCTS keys
# plus the two umbrella selections that aren't single products, so a tag can
# describe "this demonstrates Live Sports" without naming a package.
CASE_STUDY_PRODUCT_TAGS = list(FALLBACK_PRODUCTS) + ["live_sports", "total_tv"]


def _valid_tags(values, allowed):
    """Keep only tags that really exist, preserving order. Claude is told the
    exact lists, but a tag that doesn't match would silently never match a
    proposal either -- better dropped here than mysteriously inert later."""
    seen = []
    for value in values or []:
        if value in allowed and value not in seen:
            seen.append(value)
    return seen


def build_case_study_suggest_prompt(description, case_studies):
    catalog = [{"id": c["id"], "title": c["title"], "verticals": c.get("verticals") or [],
                "products": c.get("products") or [], "summary": c.get("summary") or ""}
               for c in case_studies]
    return f"""A Premion seller is working on a CTV/OTT campaign and wants the most relevant case studies to show the client. Return ONLY valid JSON -- no markdown fences, no preamble:

{{"recommendations": [{{"id": "the exact id from the list", "reason": "one line on why this one fits"}}]}}

Pick only case studies that genuinely help this pitch, best first, at most 5. Relevance means the client's *situation* matches -- same industry, comparable objective, or a product mix the seller is likely proposing. A case study from a different industry is still worth recommending if what it proves (a brand lift result, a first-party data match, a retargeting outcome) is what this client needs to see; say so in the reason. If nothing in the vault fits, return an empty list rather than padding it.

"reason" is shown to the seller, so write it for them: concrete and specific ("regional bank, same deposit-account objective, 20% lending lift"), never generic ("this is a relevant case study").

Case studies available (JSON): {json.dumps(catalog)}

The client / campaign:
\"\"\"
{description}
\"\"\"
"""


PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def case_study_filename(row):
    """A filename a seller can hand to a client, built from the title rather
    than the source deck's name -- those are things like
    "PREMION_Case Study_Regional_Residential_HVAC_and_Home_Services_Leader.pptx".
    Strips punctuation so it's safe on every OS."""
    base = (row.get("title") or Path(row["filename"]).stem).strip()
    # Separators become spaces before punctuation is stripped, or "CTV/OTT"
    # fuses into "CTVOTT".
    base = re.sub(r"[/&+]", " ", base)
    base = re.sub(r"[^\w\s-]", "", base)
    base = re.sub(r"[\s_-]+", "_", base).strip("_")
    return f"{(base or 'case_study')[:80]}.pptx"
