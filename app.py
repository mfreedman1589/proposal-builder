"""
app.py -- Streamlit form for the Premion Proposal Builder (build spec
section 5), wired directly to assembly.py's selection -> deletion -> fill
pipeline.

Persistent data (master deck versions, products/rates, the audience catalog,
proposal history) lives in Supabase, reached only through db.py, which falls
back to the local file / hardcoded copies below whenever it's unreachable.
"""

import io
import json
import re
import tempfile
from datetime import date, datetime
from pathlib import Path

import anthropic
import pandas as pd
import streamlit as st
from pptx import Presentation

import assembly
import db
import slide_map
from audience_catalog import catalog_warning, load_audience_catalog, validate_segments

st.set_page_config(page_title="Premion Proposal Builder", layout="wide")

ANTHROPIC_MODEL = "claude-sonnet-4-6"
ANTHROPIC_MAX_TOKENS = 2000

# The master deck normally comes from the `decks` storage bucket (whichever
# deck_versions row is active), downloaded once per session. This checked-in
# copy is the fallback for when Supabase is unreachable -- see db.py's
# fallback policy.
LOCAL_MASTER_DECK_PATH = assembly.MASTER_DECK_PATH

# Read from Supabase's `audiences` table, falling back to parsing the local
# brochure PDF. Cached either way, so this is a lookup after the first rerun.
audience_catalog = load_audience_catalog()
AUDIENCE_CATALOG_WARNING = catalog_warning()

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

SPORTS = {
    "NFL - Regular Season": "nfl_reg",
    "NFL - Playoffs": "nfl_playoffs",
    "NFL - Home Team": "nfl_home_team",
    "NBA - Regular Season": "nba_reg",
    "NBA - Playoffs": "nba_playoffs",
    "WNBA - Regular Season": "wnba_reg",
    "WNBA - Playoffs": "wnba_playoffs",
    "NHL - Regular Season": "nhl_reg",
    "NHL - Playoffs": "nhl_playoffs",
    "MLB - Regular Season": "mlb_reg",
    "MLB - Playoffs": "mlb_playoffs",
    "NCAAF - Regular Season": "ncaaf_reg",
    "NCAAF - Conference": "ncaaf_conference",
    "NCAAF - Playoffs": "ncaaf_playoffs",
    "NCAAF - Home Team": "ncaaf_home_team",
    "NCAA Basketball": "ncaa_basketball",
    "Live Golf": "golf",
    "PGA Majors": "pga_majors",
    "Soccer (Professional)": "soccer_pro",
    "Prestige Sports": "prestige_sports",
    "All Live Sports": "all_live_sports",
}
SPORT_LABEL_BY_VALUE = {v: k for k, v in SPORTS.items()}

STREAMING_RETARGETING_TARGETING = "Retarget Exposed CTV Viewers"
# Dynamic Video Ads: a one-time creative build, billed as a flat fee rather
# than a CPM. The amount is a starting point the seller edits in the grid.
DYNAMIC_AD_LINE_LABEL = "Dynamic Ad Creation"
DYNAMIC_AD_TARGETING = "Creative build + monthly refresh"
DYNAMIC_AD_DEFAULT_FEE = 850.0
LIVE_SPORTS_TARGETING = "100% Live, 100% In-Game, 100% CTV"

# A media-plan line can name a Live Sports package instead of a PRODUCTS
# entry, as "sport:<sport_key>". Sports carry their own per-package rates
# rather than one product default, so in the products table they're rows
# keyed "sport:<sport_key>" with display_group 'sports' -- without this
# namespace a drafted sports line had no way to be expressed at all and
# would reach for premion_streaming_tv (and its $32 CPM) instead of, say,
# the $66 NFL playoffs rate.
SPORT_PRODUCT_PREFIX = "sport:"

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

FALLBACK_SPORT_CPM = {
    "nfl_reg": 62.00, "nfl_playoffs": 66.00, "nfl_home_team": 85.00,
    "nba_reg": 57.00, "nba_playoffs": 63.00,
    "wnba_reg": 59.00, "wnba_playoffs": 61.00,
    "nhl_reg": 57.00, "nhl_playoffs": 63.00,
    "mlb_reg": 52.00, "mlb_playoffs": 63.00,
    "ncaaf_reg": 60.00, "ncaaf_conference": 60.00, "ncaaf_playoffs": 63.00, "ncaaf_home_team": 70.00,
    "ncaa_basketball": 55.00,
    "golf": 60.00,
    "pga_majors": 62.00,
    "soccer_pro": 57.00,
    "prestige_sports": 54.00,
    "all_live_sports": 50.00,
}
DEFAULT_SPORT_CPM = 45.00


def sport_product_label(sport_key):
    """The tactic name a Live Sports package seeds its media-plan row with."""
    return f"Live Sports - {SPORT_LABEL_BY_VALUE.get(sport_key, sport_key)}"


@st.cache_data(ttl=600, show_spinner=False)
def load_rate_card():
    """The live rate card: (products, sport_cpm, sport_labels,
    targeting_copy_by_label, warning).

    Products and Live Sports packages share one table -- a sports row is just
    one whose key carries the "sport:" prefix -- so this splits them back into
    the two shapes the rest of the app already works in. On any failure it
    returns the FALLBACK_* copies above with a warning for the caller to show;
    an unreachable rate card degrades the numbers, it doesn't stop the form.

    Cached with a TTL rather than for the process lifetime so a rate edit in
    Supabase reaches a long-running session without a restart.
    """
    rows, warning = db.fetch_products()
    if rows is None:
        return (FALLBACK_PRODUCTS, FALLBACK_SPORT_CPM,
                {key: sport_product_label(key) for key in FALLBACK_SPORT_CPM},
                _targeting_copy_map(FALLBACK_PRODUCTS, FALLBACK_SPORT_CPM),
                f"{warning}. Using the built-in fallback rate card -- CPMs may be out of date.")

    products, sport_cpm, sport_labels, targeting_copy = {}, {}, {}, {}
    for row in rows:
        key, name = row["key"], row["name"]
        cpm = float(row["default_cpm"] or 0.0)
        if key.startswith(SPORT_PRODUCT_PREFIX):
            sport_key = key[len(SPORT_PRODUCT_PREFIX):]
            sport_cpm[sport_key] = cpm
            sport_labels[sport_key] = name
        else:
            products[key] = {"label": name, "default_cpm": cpm,
                             "line_type": row.get("display_group") or ""}
        if row.get("targeting_copy"):
            targeting_copy[name] = row["targeting_copy"]
    return products, sport_cpm, sport_labels, targeting_copy, None


def _targeting_copy_map(products, sport_cpm):
    """{tactic label: fixed targeting copy} for the fallback rate card, so
    resolve_row_defaults reads the same shape either way."""
    copy_map = {spec["label"]: spec["targeting_copy"]
                for spec in products.values() if spec.get("targeting_copy")}
    copy_map.update({sport_product_label(key): LIVE_SPORTS_TARGETING for key in sport_cpm})
    return copy_map


# The live rate card, resolved once per rerun (the loader is cached, so this
# is a dict lookup after the first). PRODUCTS/SPORT_CPM keep their names
# because everything downstream reads them; what changed is where they come
# from. RATE_CARD_WARNING is surfaced in the form when it's a fallback.
PRODUCTS, SPORT_CPM, SPORT_PRODUCT_LABELS, TARGETING_COPY_BY_LABEL, RATE_CARD_WARNING = load_rate_card()

MEDIA_PLAN_FIELDS = ["Tactic", "Flight", "Geo", "Targeting", "Impressions", "CPM", "Type", "Cost"]
ROW_TYPE_RATE = "Rate"
ROW_TYPE_FLAT_FEE = "Flat Fee"
# Which side of a rate row the user last typed into. That side is the
# driver; the other is recomputed from it. A CPM edit doesn't change the
# driver, it just re-derives the other side from whichever one is driving.
DRIVER_IMPRESSIONS = "impressions"
DRIVER_COST = "cost"

BREAKOUT_MONTHLY = "Monthly (default)"
BREAKOUT_FULL_FLIGHT = "Full Flight"
BREAKOUT_MODES = [BREAKOUT_MONTHLY, BREAKOUT_FULL_FLIGHT]

# A proposal carries 1-3 media plan options (good/better/best, or several
# budget scenarios). One option is the ordinary case and renders exactly as a
# single-plan proposal always has -- no "Option A" label anywhere.
MAX_PLAN_OPTIONS = 3
DEFAULT_OPTION_NAMES = ["Option A", "Option B", "Option C"]


def bump_plan_options_generation(target=None):
    """Invalidate the per-option widget keys.

    A keyed widget's session_state entry wins over its `value=` argument, so
    once `option_name_0` exists holding "Option A", re-rendering it with
    value="Kickoff" silently keeps "Option A" -- and since the option's name
    is read back *from* the widget, the drafted name is then written into the
    option dict and lost for good. Any structural replacement of the option
    list (a draft, an add, a remove) therefore has to move the widgets to
    fresh keys. Same reasoning as the per-option editor `version`.

    Pass `target` (a dict of pending updates) to record the bump there
    instead of writing session_state directly.
    """
    store = st.session_state if target is None else target
    store["plan_options_gen"] = st.session_state.get("plan_options_gen", 0) + 1

PRESETS = {
    "Quick Pitch": "quick_pitch",
    "Standard": "standard",
    "Extended": "extended",
}

# ---------------- Draft-from-notes (Claude) ----------------
# Maps a PRODUCTS key to the session_state widget key(s) that need to be set
# True to select it in the form (some products imply turning on a parent
# toggle, e.g. any Audience Marketplace format also flips am_enabled).
PRODUCT_TO_WIDGET_KEYS = {
    "premion_streaming_tv": [("premion_streaming_tv", True)],
    "streaming_retargeting_display": [("streaming_retargeting_enabled", True), ("sr_disp", True)],
    "streaming_retargeting_preroll": [("streaming_retargeting_enabled", True), ("sr_pre", True)],
    "audience_targeting_display": [("am_enabled", True), ("am_at_disp", True)],
    "audience_targeting_preroll": [("am_enabled", True), ("am_at_pre", True)],
    "geofencing_display": [("am_enabled", True), ("am_gf_disp", True)],
    "geofencing_preroll": [("am_enabled", True), ("am_gf_pre", True)],
    "site_retargeting_display": [("am_enabled", True), ("am_srd", True)],
    "site_retargeting_preroll": [("am_enabled", True), ("am_srp", True)],
    "broadcast_tv": [("total_tv", True)],
}
# Not a PRODUCTS entry -- it's a creative option with its own flat fee, not a
# CPM product -- but a drafted plan still needs to be able to switch it on,
# and DRAFT_KEY_SECTIONS has to tag it "products" to match the section its
# widget lives in. Listed here for both, and excluded from the media-plan
# line namespace below.
PRODUCT_TO_WIDGET_KEYS["dynamic_creative"] = [("dynamic_creative", True)]

# Switched on through the draft's "attribution" array rather than as a media
# plan line: its cost is a fixed production fee the form seeds, not something
# the model should be allocating budget to.
NON_LINE_PRODUCT_KEYS = frozenset({"dynamic_creative"})

# What each attribution option means, and the phrasing in notes that maps to
# it. Rendered into the draft prompt -- a bare list of seven snake_case
# tokens left the model to infer each one's meaning from its name, and
# `dynamic_creative` was the case that proved it doesn't work: notes asking
# for creative that "swaps based on which county it's serving" are
# unmistakable to a human, but nothing in the token said so and the request
# was silently dropped from the draft entirely. ATTRIBUTION_OPTIONS is
# derived from these keys, so an option can't be added without a description.
ATTRIBUTION_DESCRIPTIONS = {
    "web": "Web attribution -- site visits and online conversions measured off a tracking pixel. "
           "Included in every campaign by default. Notes phrasing: \"want to see site traffic\", "
           "\"form fills\", \"did the web traffic move\".",
    "sales": "Sales attribution -- ad exposure matched against the client's own closed-sale "
             "records, which they upload from their CRM. Notes phrasing: \"they can share close "
             "data\", \"match it back to actual sales\", \"tie it to revenue\".",
    "brand_lift": "Brand lift study -- a survey measuring awareness and consideration lift. Notes "
                  "phrasing: \"brand lift\", \"awareness study\", \"run a survey\". Leave it out "
                  "when the notes decline it (\"I don't need a survey\").",
    "first_party": "First-party data targeting -- targeting built from the client's OWN customer "
                   "or prospect data rather than catalog segments. Notes phrasing: \"use their "
                   "customer list\", \"match our database\", \"our own CRM data for targeting\".",
    "linear_reach_ext": "Linear reach extension -- reaching viewers the client's linear/broadcast "
                        "schedule missed. Only valid alongside a Total TV / broadcast buy (the "
                        "form disables it otherwise), so include it only when the plan actually "
                        "has broadcast in it. Notes phrasing: \"extend our TV buy\", \"reach the "
                        "cord-cutters our spots miss\".",
    "commercial_production": "Commercial production provided at NO CHARGE -- appears in the deck's "
                             "Included-with-Campaign list. Notes phrasing: \"we'll produce the spot "
                             "for them\", \"production is on us\". Do NOT include this when the "
                             "notes quote a price for production -- see the flat-fee rule above.",
    "dynamic_creative": "Dynamic Video Ads -- one creative that VARIES automatically by location, "
                        "offer, inventory or audience instead of a single fixed spot. This is the "
                        "one people describe without naming it: \"swaps based on which county it's "
                        "serving\", \"shows the nearest location\", \"pulls in current inventory\", "
                        "\"different offer by market\", \"versioned by store\", \"dynamic\". "
                        "Include it whenever the notes describe creative changing by any such "
                        "variable. It carries its own one-time creative build fee automatically, "
                        "so do not also invent a custom_fee line for it.",
}

ATTRIBUTION_OPTIONS = list(ATTRIBUTION_DESCRIPTIONS)
# "web" is always included by default (build_included_list) -- there's no
# form toggle for it, so it has no entry here.
ATTRIBUTION_FIELD_MAP = {
    "sales": "sales_attribution",
    "brand_lift": "brand_lift",
    "first_party": "first_party_data",
    "linear_reach_ext": "linear_reach_extension",
    "commercial_production": "commercial_production",
    "dynamic_creative": "dynamic_creative",
}

# Coarse category hints used to pick a relevant slice of the audience catalog
# to send Claude, based on a cheap local keyword guess at the vertical --
# Claude's own returned "vertical" field is authoritative either way, this
# just keeps the prompt from having to carry the full ~370-segment catalog.
VERTICAL_CATEGORY_HINTS = {
    "healthcare": ["HLTH", "DEMO", "HH", "LIFESTAGE"],
    "retail": ["RETAIL", "LIFESTYLE", "DEMO"],
    "travel": ["TRAVEL", "LIFESTYLE", "DEMO"],
    "home_improvement": ["RETAIL", "HH", "LIFESTAGE"],
    "banking": ["FIN", "DEMO", "HH"],
    "entertainment": ["ENT", "LIFESTYLE", "DEMO"],
    "dining_qsr": ["FOOD", "LIFESTYLE", "DEMO"],
    "auto": ["AUTO", "DEMO"],
    "education": ["LIFESTAGE", "DEMO"],
    # FIN first because the catalog's one directly-legal segment ("FIN Legal
    # Services") lives there; DEMO next for the occupation targeting these
    # campaigns lean on (blue-collar, veteran), then the insurance segments
    # that personal-injury and workers'-comp work turns on.
    "legal": ["FIN", "DEMO", "HLTH", "AUTO"],
}
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
    "restoration": "home_improvement",
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

CUSTOM_FEE_PRODUCT = "custom_fee"

DRAFT_JSON_SCHEMA_EXAMPLE = """{
  "client_name": "", "vertical": "", "market": "DC|Harrisburg",
  "agency_involved": false, "spanish_campaign": false,
  "flight_start": "YYYY-MM-DD", "flight_end": "YYYY-MM-DD", "geo": "",
  "total_budget": 0,
  "breakout": "monthly|full_flight",
  "media_plan_lines": [
    {"product": "premion_streaming_tv", "label": "Commercial", "audience_track": "SMB owners and business executives", "allocation": {"percent_of_remainder": 60}},
    {"product": "premion_streaming_tv", "label": "Retail", "audience_track": "consumers in-market for deposit accounts", "allocation": {"percent_of_remainder": 40}, "cpm": 28},
    {"product": "streaming_retargeting_display", "allocation": {"percent_of_total": 10}},
    {"product": "sport:nfl_playoffs", "allocation": {"flat_amount": 40000}},
    {"product": "custom_fee", "label": "Dynamic Ad Creation", "allocation": {"flat_amount": 850}}
  ],
  "options": null,
  "audiences": [{"segment": "exact catalog name", "geo": ""}],
  "attribution": ["web", "sales", "brand_lift", "first_party", "linear_reach_ext", "commercial_production"],
  "sports": [],
  "campaign_specs": {
    "goals": [], "audience": [], "geography": [],
    "budget": [], "placements": [], "timing": []
  },
  "unresolved": ["plain-language notes about anything ambiguous or assumed"]
}"""

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


# Which form section each session_state key the draft writes belongs to --
# keyed to the *widget's* own section (the one its on_change clears), since
# that's what tells us the user has since hand-edited it. Used by the
# clarify-and-re-draft round to leave hand-edited sections alone.
DRAFT_KEY_SECTIONS = {
    "client_name": "basics", "market_choice": "basics",
    "vertical_choice": "basics", "agency_involved": "basics",
    "spanish_campaign": "attribution",
    "flight_start": "flight", "flight_end": "flight",
    "goals_text": "specs", "audience_text": "specs", "geography_text": "specs",
    "budget_text": "specs", "placements_text": "specs", "timing_text": "specs",
    "avails_seed_rows": "avails", "avails_version": "avails",
    "live_sports_enabled": "products", "selected_sports": "products",
    "plan_options": "media_plan", "media_plan_markup": "media_plan",
    "plan_options_gen": "media_plan",
    # Internal bookkeeping that only means anything alongside the rows it
    # describes -- it has to be skipped with them or it would claim rows that
    # were never written.
    "_product_seed_key": "media_plan", "_shared_fields_key": "media_plan",
}
DRAFT_KEY_SECTIONS.update({key: "attribution" for key in ATTRIBUTION_FIELD_MAP.values()})
DRAFT_KEY_SECTIONS.update({widget_key: "products"
                           for keys in PRODUCT_TO_WIDGET_KEYS.values()
                           for widget_key, _ in keys})


def check_password():
    if st.session_state.get("authed"):
        return True

    st.title("Premion Proposal Builder")
    pwd = st.text_input("Password", type="password")
    if st.button("Log in"):
        expected = st.secrets.get("APP_PASSWORD")
        if expected is None:
            st.error("APP_PASSWORD is not set in .streamlit/secrets.toml.")
        elif pwd == expected:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


ADD_USER_OPTION = "➕ Add a name..."


def current_user():
    return st.session_state.get("current_user")


def check_identity():
    """Ask who's using the app, after the shared password.

    Not authentication -- the password is the gate, this only attributes
    work. It's a separate step rather than a field on each form because it's
    answered once per session and then used in three places (proposals, case
    study uploads, attached-file notes).

    The list builds itself: anyone can add a name and it's there for everyone
    afterwards, deduplicated case-insensitively so "matt" and "Matt" can't
    become two people. If Supabase is unreachable the step is skipped
    entirely rather than blocking -- an unattributed proposal is a far better
    outcome than a seller who can't build one.
    """
    if current_user():
        return True

    names, warning = db.fetch_team_members()
    if names is None:
        st.session_state["current_user"] = None
        st.session_state["identity_skipped"] = True
        return True

    st.title("Premion Proposal Builder")
    st.caption("Who's using the app? This just labels the proposals you generate so the team "
               "can tell whose is whose — pick your name, or add it if it's not there yet.")

    options = names + [ADD_USER_OPTION]
    picked = st.selectbox("Your name", options, index=None, placeholder="Choose your name",
                          key="identity_pick")

    if picked == ADD_USER_OPTION:
        new_name = st.text_input("Your name", key="identity_new_name",
                                 placeholder="First name is fine")
        if st.button("Add and continue", disabled=not new_name.strip()):
            stored, error = db.add_team_member(new_name)
            if error:
                st.error(error)
            else:
                st.session_state["current_user"] = stored
                st.rerun()
    elif picked:
        if st.button("Continue"):
            st.session_state["current_user"] = picked
            st.rerun()

    if warning:
        st.caption(warning)
    return False


def render_identity_sidebar():
    """Who's signed in, and a way to change it."""
    user = current_user()
    if user:
        st.sidebar.caption(f"Signed in as **{user}**")
    elif st.session_state.get("identity_skipped"):
        st.sidebar.caption("Not signed in — proposals won't be attributed")
    if st.sidebar.button("Switch user", use_container_width=True):
        for key in ("current_user", "identity_skipped", "identity_pick", "identity_new_name"):
            st.session_state.pop(key, None)
        st.rerun()


def _clear_ai_section(section):
    st.session_state.get("ai_filled_sections", set()).discard(section)


def ai_section_badge(section):
    if section in st.session_state.get("ai_filled_sections", set()):
        st.caption("🤖 Some fields below were drafted from your notes -- review before generating.")


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
    everything else, each by times_used. A vertical only *prioritizes* --
    it never filters anything out, so a segment outside the vertical's
    categories is still reachable, just further down."""
    if vertical_hint and vertical_hint in VERTICAL_CATEGORY_HINTS:
        cats = VERTICAL_CATEGORY_HINTS[vertical_hint]
        relevant = catalog[catalog["category"].isin(cats)].sort_values("times_used", ascending=False)
        rest = catalog[~catalog["category"].isin(cats)].sort_values("times_used", ascending=False)
        return pd.concat([relevant, rest])
    return catalog.sort_values("times_used", ascending=False)


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
    return sliced[["segment", "category", "subcategory", "rfp_selectable", "times_used"]].to_dict("records")


def build_draft_prompt(notes):
    vertical_hint = _detect_vertical_hint(notes)
    catalog_slice = build_catalog_slice(vertical_hint)
    products_info = {k: {"label": v["label"], "default_cpm": v["default_cpm"]} for k, v in PRODUCTS.items()}
    today = date.today().isoformat()
    attribution_help = "\n".join(f'  - "{key}": {text}'
                                 for key, text in ATTRIBUTION_DESCRIPTIONS.items())

    return f"""You are drafting a first pass at a Premion CTV/OTT advertising proposal from raw meeting/discovery notes. Today's date is {today}. Return ONLY valid JSON matching the schema below -- no markdown code fences, no preamble, no explanation, just the JSON object.

Schema:
{DRAFT_JSON_SCHEMA_EXAMPLE}

"media_plan_lines" is the full media plan, expressed one entry per intended row -- not one entry per product. If the notes call for the same product run as separate lines (e.g. two different audience tracks, or a commercial vs. retail split), give each its own entry with its own "label" and "audience_track"; each becomes its own media plan row, with "label" appended to the product's own tactic name (e.g. "Premion Streaming TV — Commercial") and "audience_track" as that row's Targeting. A line with no "label" just uses the product's own name as-is.

Each line's "product" must be exactly one of:
- a product key: {list(PRODUCTS.keys())}
- a Live Sports package, written as "{SPORT_PRODUCT_PREFIX}<sport_key>" where <sport_key> is exactly one of {list(SPORTS.values())} (e.g. "{SPORT_PRODUCT_PREFIX}nfl_playoffs"). Always use this form for a sports buy -- never bill sports inventory as "premion_streaming_tv", which carries a completely different (much lower) rate.
- "{CUSTOM_FEE_PRODUCT}", for a one-time flat fee that isn't a real media-buy product (e.g. a production/creative fee) -- its "label" becomes its Tactic name directly and it must use a "flat_amount" allocation. Only use this when the notes state an actual dollar amount to charge. **Never emit a line whose amount is zero.** Things provided at no charge -- commercial production, reporting, account management -- are not media plan lines at all: put "commercial_production" in "attribution" instead and it appears in the deck's "Included with Campaign" list, which is where a client expects to see it.

**A quoted price always wins over the no-charge default.** The rule above describes what is normally free, not what is free in these notes. If the notes state a charge for commercial production ("we're doing the :30, quoted $850"), then it IS a chargeable line: emit a "{CUSTOM_FEE_PRODUCT}" line labelled for commercial production, and do NOT also put "commercial_production" in "attribution" -- billing for it and listing it as included at no charge in the same deck contradict each other in front of the client. The same applies to anything else on that list.

**Never move a quoted amount to a different deliverable than the notes attach it to.** If the notes quote $850 for commercial production, that $850 is the commercial production line -- it does not become a dynamic ad creation fee, a versioning fee, or anything else, however plausible the reassignment seems. Label the line with what the notes actually called it. If the notes are genuinely ambiguous about which deliverable a quoted amount covers, keep the wording from the notes as the label and say so in "unresolved" -- a flagged ambiguity is reviewable, a silent reassignment is not.

Most proposals are a single plan: put its rows in "media_plan_lines" and leave "options" null. Only when the notes explicitly ask for SCENARIOS to choose between -- good/better/best, tiered budgets, "show them a $50K and a $75K version" -- return "options" instead, as up to {MAX_PLAN_OPTIONS} entries:
  "options": [{{"name": "Good", "total_budget": 50000, "breakout": "monthly", "media_plan_lines": [...]}}, {{"name": "Better", "total_budget": 75000, "breakout": "monthly", "media_plan_lines": [...]}}]
Each option is a complete plan in its own right, with its own budget and its own full set of lines, and becomes its own media plan slide in the deck. "name" is what the client sees appended to the plan title, so use the notes' own words for the tier ("Good"/"Better"/"Best", "$50K Plan") rather than a generic letter. When "options" is set, leave "media_plan_lines" empty. Do NOT invent scenarios the notes didn't ask for -- one plan is the normal answer.

EVERY option MUST carry its own "total_budget" -- it is what that scenario costs, and each option's allocations are resolved independently against it. An option's "total_budget" falls back to the top-level one only if omitted, and if neither is set that option prices at $0 and gets dropped entirely, which is never a useful answer. When the notes state a BUDGET RANGE with no instruction on how to split it ("between $50K and $75K", "somewhere in the 50 to 75 range, wants to see both"), the right answer is one option per stated figure, each carrying that figure as its own "total_budget" -- e.g. a "$50K Plan" at 50000 and a "$75K Plan" at 75000. If the notes give a range but you cannot tell what the individual figures should be, put a single plan at the lower figure and say so in "unresolved"; never return lines with no budget behind them.

INCLUDE ONLY THE PRODUCTS THE NOTES ACTUALLY CALL FOR. There is no mandatory line and no default product -- "premion_streaming_tv" in particular is NOT required and must not be added just to have a baseline CTV line. A sports-only plan, an Audience-Marketplace-only plan, a retargeting-only plan, or a single-line plan are all perfectly valid proposals. If the notes describe an NFL campaign and nothing else, the correct media plan is one NFL line and nothing else. If the notes are genuinely silent about what to buy, say so in "unresolved" instead of inventing a product mix.

**But every product the notes DO ask for gets a line, even when no budget is given for it.** "She wants audience targeting on top of the general streaming" names a product; the absence of a split is a missing number, not a reason to leave the product out. Include the line with a "split_evenly" allocation so it shares what's left with the other unspecified lines, and flag the assumption in "unresolved" ("the notes didn't say how to split between X and Y -- divided evenly, confirm the intended weighting"). Omitting a requested product is the worse failure of the two: an assumption the reviewer can see gets corrected, whereas a product that never appears in the plan is invisible to them and quietly missing from the deck.

"total_budget" is always what the WHOLE campaign costs across the entire flight, never a monthly rate -- and so is every dollar amount the allocations resolve to. If the notes quote the budget per month ("$20K a month for three months"), multiply it out to the full-flight figure yourself ($60,000) and note the per-month figure in "unresolved" if it's worth flagging.

"breakout" is a separate question -- not what the plan costs, but how it's presented: "monthly" for a plan broken out month by month (the normal case, and the right answer whenever the notes talk in per-month terms at all), or "full_flight" only if the notes specifically want the flight shown as a single combined period. Python divides the campaign total across the flight for a monthly breakout; do not do that arithmetic yourself.

Each line's "allocation" has exactly one key:
- "flat_amount": this line costs exactly this many dollars.
- "percent_of_total": this line costs this percent of total_budget.
- "percent_of_remainder": this line costs this percent of whatever's left after all "flat_amount" and "percent_of_total" lines are subtracted from total_budget (percent_of_remainder entries across lines should sum to 100 if they're meant to exhaust the remainder).
- "split_evenly": this line shares equally, with every other "split_evenly" line, in whatever's left after "flat_amount"/"percent_of_total"/"percent_of_remainder" lines are all accounted for -- use this for "split evenly across N audiences/tracks" instead of trying to pre-compute a percentage yourself.
Do NOT do any arithmetic yourself beyond picking which allocation type fits each line -- Python resolves flat_amount and percent_of_total first, then percent_of_remainder, then splits whatever's left evenly across split_evenly lines, then computes every dollar amount, impression count, and markup.

Each line may also carry an optional "cpm", the rate for that line in dollars. Rates are negotiated per deal, so set it whenever the notes state a rate for that line -- "$28 CPM on the Premion line", "they're getting the streaming at 30", "we agreed $45 for the NFL inventory". Omit it and the product's rate card default applies, which is what you want whenever the notes say nothing about rate or say to hold to the rate card ("at rate card", "standard rates", "no discount"). Set it ONLY from a rate the notes actually state -- never to hit a budget or impression target, which is what the allocations are for. It is a plain number (28, not "$28" or "28 CPM"), it is the net rate before any agency markup (Python applies the markup), and it never applies to a "{CUSTOM_FEE_PRODUCT}" line, which has no rate at all. Every override is flagged for the reviewer automatically, so you do not need to mention it in "unresolved" yourself.

Rules:
- "vertical" must be exactly one of: {list(VERTICALS.values())}
- "market" must be exactly "DC" or "Harrisburg".
- "sports" entries must be exactly one of: {list(SPORTS.values())}. This drives which sports package slides go in the deck -- list every package that also appears as a "{SPORT_PRODUCT_PREFIX}" media plan line, and leave it empty when the notes call for no sports at all.
- "attribution" entries must be drawn from this list, using the exact key shown. Include every one the notes call for -- these drive real slides, real Included-with-Campaign entries and real toggles, so an option the notes ask for and you omit simply never reaches the proposal:
{attribution_help}
- "audiences[].segment" must be an EXACT name from the audience catalog slice below -- do not paraphrase or invent segment names. If nothing in the slice fits, it's fine to omit audiences or note it in "unresolved". {AUDIENCE_MATCH_GUIDANCE} When a media_plan_lines entry's "audience_track" describes the same audience as one of your "audiences" entries, use the same wording for both.
- **At most ONE non-RFP-selectable segment. This is a hard limit, not a preference.** Each catalog entry carries an "rfp_selectable" flag, and a campaign may book only one segment with "rfp_selectable": false (a "custom" segment). A draft containing two or more is invalid and cannot be used until someone removes the extras by hand. So: when two segments would serve the same purpose, take the RFP-selectable one. If several custom segments all look relevant -- which happens when a niche category's best matches are all custom -- **choose the single most important one and name the others in "unresolved" as alternatives the reviewer could swap in**, rather than returning them all. Count the custom segments in your "audiences" array before you finish; if there is more than one, cut it down. Whenever you do return a custom segment, say in "unresolved" which one it is and why no RFP-selectable segment covered it.
- Never invent a "Max Monthly Avails" number -- that field doesn't exist in this schema on purpose; avails come from a real system, not from you.
- Use "unresolved" for anything ambiguous, assumed, or not mentioned in the notes -- plain language, one item per ambiguity.
- Dates in flight_start/flight_end should be YYYY-MM-DD. If the notes give a date without a year (e.g. "September through November"), resolve it to the NEXT upcoming occurrence of that month relative to today's date -- never a date already in the past -- and flag that assumption in "unresolved" the same as any other assumption.

Available products and default CPMs (JSON): {json.dumps(products_info)}
Audience catalog slice -- {len(catalog_slice)} of {len(load_audience_catalog())} total segments (JSON): {json.dumps(catalog_slice)}

Meeting/discovery notes:
\"\"\"
{notes}
\"\"\"
"""


def _strip_markdown_fences(text):
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines)
    return text.strip()


def _parse_draft_json(raw_text):
    try:
        return json.loads(raw_text), None
    except json.JSONDecodeError:
        pass
    try:
        return json.loads(_strip_markdown_fences(raw_text)), None
    except json.JSONDecodeError as exc:
        return None, f"Claude's response wasn't valid JSON even after stripping markdown fences: {exc}"


def _call_claude_json(prompt):
    """Sends one prompt to Claude and parses the response as JSON. Returns
    (parsed_dict, error_message) -- exactly one is None."""
    api_key = st.secrets.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None, "ANTHROPIC_API_KEY is not set in .streamlit/secrets.toml."

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=ANTHROPIC_MAX_TOKENS,
            messages=[{"role": "user", "content": prompt}],
        )
        raw_text = response.content[0].text
    except Exception as exc:
        return None, f"Claude API call failed: {exc}"

    return _parse_draft_json(raw_text)


def call_claude_draft(notes):
    """Returns (draft_dict, error_message) -- exactly one is None."""
    return _call_claude_json(build_draft_prompt(notes))


def build_redraft_prompt(notes, previous_draft, clarifications):
    """A revision pass, not a fresh draft: the model gets its own previous
    JSON back plus the user's answers to the open questions, and is told to
    change only what the answers actually bear on. Starting over would churn
    parts of the draft the user already accepted."""
    return build_draft_prompt(notes) + f"""

--- REVISION PASS ---
You already produced the draft below from these same notes. The user has now answered the open questions you raised. Return a REVISED version of that JSON -- do not start over.

Your previous draft (JSON):
{json.dumps(previous_draft, indent=2)}

The user's clarifications:
\"\"\"
{clarifications}
\"\"\"

Revision rules:
- Change only what the clarifications actually bear on, plus anything else that must change to stay consistent with them (e.g. if the budget now excludes a fee, the line amounts that depend on it change too).
- A clarification naming a rate ("drop Premion to $28", "hold sports at rate card", "they negotiated the retargeting to 4.50") changes that line's "cpm" -- set it to the stated number, or REMOVE the "cpm" key entirely to go back to the rate card default. Leave the line's "allocation" alone unless the clarification also changes what it should spend: the budget and the rate are independent, and Python re-derives impressions from both.
- A clarification naming a budget ("make the second option $80K instead") changes that option's "total_budget", not its allocations.
- Keep every other field byte-identical to your previous draft. Do not re-word, re-order, or "improve" parts the clarifications didn't touch.
- Remove from "unresolved" anything the clarifications have now settled, and keep anything still genuinely open. If the clarifications raise something new and ambiguous, add it.
- The output is still the complete JSON object in the schema above -- the whole draft, revised, not a diff.
"""


def call_claude_redraft(notes, previous_draft, clarifications):
    """Returns (draft_dict, error_message) -- exactly one is None.

    The revision is merged *over* the previous draft rather than replacing
    it. The model is told to return the whole object, but a dropped field
    would otherwise silently fall back to a schema default -- and a
    disappearing `agency_involved` doesn't read as missing, it reads as
    "no agency", quietly repricing every line at net instead of gross.
    Anything the revision does return still wins.
    """
    revised, error = _call_claude_json(build_redraft_prompt(notes, previous_draft, clarifications))
    if error:
        return None, error
    return {**previous_draft, **revised}, None


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
    parsed, error = _call_claude_json(build_audience_finder_prompt(description, vertical_hint))
    if error:
        return None, None, error

    raw_recs = parsed.get("recommendations", []) or []
    names = [r.get("segment", "") for r in raw_recs if r.get("segment")]
    matched, unmatched = validate_segments(names)
    recs = [r for r in raw_recs if r.get("segment") in matched]
    return recs, unmatched, None


def _parse_draft_date(value):
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y", "%b %d, %Y", "%B %Y", "%b %Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except (ValueError, TypeError):
            continue
    return None


def _session_getter(key, default=False):
    """The default reader for the seed-selection shape: the live widget
    values. The draft handler substitutes one that prefers its own pending
    writes."""
    return st.session_state.get(key, default)


def read_products_selection(get=None):
    """The same products_selection shape main() builds from the live
    widgets, but read directly from session_state -- lets the draft handler
    compute the exact product_seed_key/shared_fields_key the next run will
    see, and lets main() and the draft handler share one source of truth.

    `get(key, default)` overrides where a widget value is read from. The
    draft handler passes one that prefers its own pending `updates`, since at
    that point session_state still holds pre-draft values. It must go through
    *this* function rather than assembling the dict itself: the key is
    compared as a `str()` of the whole structure, so a hand-rolled copy that
    merely omits a key silently fails to match even when every shared value
    agrees. That is exactly what happened when `dynamic_creative` was added
    here and not to the draft handler's duplicate -- main() saw a mismatch on
    every single draft and rebuilt each option's rows as fresh $0 seeds,
    keeping only the option names.
    """
    get = get or _session_getter
    return {
        "streaming_retargeting": {
            "enabled": get("streaming_retargeting_enabled", False),
            "display": get("sr_disp", False),
            "preroll": get("sr_pre", False),
        },
        "audience_marketplace": {
            "enabled": get("am_enabled", False),
            "audience_targeting_display": get("am_at_disp", False),
            "audience_targeting_preroll": get("am_at_pre", False),
            "geofencing_display": get("am_gf_disp", False),
            "geofencing_preroll": get("am_gf_pre", False),
            "site_retargeting_display": get("am_srd", False),
            "site_retargeting_preroll": get("am_srp", False),
        },
        "live_sports": {
            "enabled": get("live_sports_enabled", False),
            "sports": [SPORTS[s] for s in get("selected_sports", [])],
        },
        "total_tv": get("total_tv", False),
        "dynamic_creative": get("dynamic_creative", False),
    }


def read_seed_selections(get=None):
    get = get or _session_getter
    return {"products": read_products_selection(get),
            "_premion_streaming_tv": get("premion_streaming_tv", False)}


def rebuild_proposal_deck(row):
    """Rebuild a logged proposal's .pptx as it was originally presented.

    Returns (buffer, filename, warnings) -- buffer is None on failure and the
    warnings say why.

    Two deliberate choices about fidelity:

    * The deck is built from the proposal's OWN `deck_version_id`, not the
      active one. That version is almost always inactive by now; storage
      keeps every version and `deck_versions` keeps every row, so it stays
      fetchable, and `db.fetch_deck_version` doesn't filter on active.
    * The media plan comes from the stored `deck_payload` verbatim rather
      than being recomputed from the rows. Recomputing would re-price against
      today's rate card, so a CPM change since would quietly rewrite what the
      client was shown -- the opposite of the point.
    """
    form = row.get("form_json") or {}
    warnings = []

    version_id = row.get("deck_version_id")
    if version_id is None:
        master_path, _, deck_warning = db.master_deck(LOCAL_MASTER_DECK_PATH)
        warnings.append("This proposal was built from the local fallback deck, not a registered "
                        "version, so the rebuild uses whatever master deck is available now.")
        if deck_warning:
            warnings.append(deck_warning)
    else:
        version_row, error = db.fetch_deck_version(version_id)
        if error:
            return None, None, warnings + [f"Couldn't fetch deck version {version_id}: {error}"]
        try:
            master_path = db.deck_file(version_row)
        except Exception as exc:
            return None, None, warnings + [
                f"Couldn't download deck version {version_id} ({db.describe_error(exc)})."]
    if master_path is None:
        return None, None, warnings + ["There's no master deck available to rebuild from."]

    options = (form.get("deck_payload") or {}).get("media_plan_options") or []
    if not options:
        return None, None, warnings + [
            "This proposal predates stored deck payloads, so it can't be rebuilt exactly. "
            "Use \"Load into form\" and generate instead."]

    avails_rows = form.get("avails_rows") or [] or [{"audience": "", "geo": "", "avails": "0"}]
    total_avails = sum(int(str(r.get("avails", "0")).replace(",", "") or 0) for r in avails_rows)
    vertical_key = row.get("vertical") or "none"
    vertical_label = next((label for label, key in VERTICALS.items() if key == vertical_key), "")
    specs = form.get("campaign_specs") or {}

    # The placeholder is the right answer only when the proposal genuinely
    # had no logo. When it had one that can't be fetched, the rebuild still
    # proceeds -- a deck with a placeholder logo beats no deck -- but it says
    # so, because the output then differs from what the client saw.
    logo_path = "placeholder_logo.png"
    if row.get("logo_storage_path"):
        try:
            logo_path = db.proposal_logo(row["logo_storage_path"])
        except Exception as exc:
            warnings.append(f"This proposal's stored logo couldn't be fetched "
                            f"({db.describe_error(exc)}) -- rebuilt with the placeholder, so the "
                            f"cover won't match what the client saw.")
    elif form.get("logo_used"):
        warnings.append("This proposal used a logo but predates logo storage, so it was rebuilt "
                        "with the placeholder -- the cover won't match what the client saw.")

    fill_data = {
        "client_name": row.get("client_name") or "Client",
        "proposal_title": form.get("proposal_title") or "CTV Strategy",
        "logo_path": logo_path,
        "vertical_display": vertical_label if (form.get("selections") or {}).get("vertical") != "none" else "",
        "campaign_specs": {
            "GOALS_BULLETS": lines_to_bullets(specs.get("goals", "")) or ["--"],
            "AUDIENCE_BULLETS": lines_to_bullets(specs.get("audience", "")) or ["--"],
            "GEOGRAPHY_BULLETS": lines_to_bullets(specs.get("geography", "")) or ["--"],
            "BUDGET_BULLETS": lines_to_bullets(specs.get("budget", "")) or ["--"],
            "PLACEMENTS_BULLETS": lines_to_bullets(specs.get("placements", "")) or ["--"],
            "TIMING_BULLETS": lines_to_bullets(specs.get("timing", "")) or ["--"],
        },
        "avails": {"rows": avails_rows, "total_avails": f"{total_avails:,}"},
        "media_plan_options": options,
    }

    # A case study removed from the vault since must not take the rebuild
    # down with it -- the deck is still worth having without one slide.
    case_study_sources = []
    for case_study in form.get("case_studies") or []:
        stored, error = db.fetch_case_study(case_study.get("id"))
        if error or not stored:
            warnings.append(f"Case study \"{case_study.get('title')}\" is no longer in the vault -- "
                            f"rebuilt without it.")
            continue
        try:
            case_study_sources.append({"path": db.case_study_file(stored["id"], stored["storage_path"]),
                                       "slides": None, "title": stored.get("title")})
        except Exception as exc:
            warnings.append(f"Case study \"{stored.get('title')}\" couldn't be fetched "
                            f"({db.describe_error(exc)}) -- rebuilt without it.")

    try:
        prs, _, _ = assembly.build_presentation(master_path, form.get("selections") or {})
        assembly.append_case_studies(prs, case_study_sources)
        assembly.personalize(prs, fill_data)
        buffer = io.BytesIO()
        prs.save(buffer)
        buffer.seek(0)
    except Exception as exc:
        return None, None, warnings + [f"Rebuild failed: {exc}"]

    filename = row.get("output_filename") or "proposal.pptx"
    return buffer, filename, warnings


def rehydrate_proposal_into_form(row, rebuild_deck_version_id=None, parent_proposal_id=None,
                                 revision_default=None):
    """Load a logged proposal's form_json back into session_state.

    Same mechanism as the draft prefill: write every widget's key *before*
    that widget renders, then rerun. Returns a list of plain-language notes
    about anything that couldn't be restored.

    Two things this must get right, both learned the hard way elsewhere:

    1. `_product_seed_key` is built by `read_seed_selections` and never
       re-assembled here. A hand-rolled copy that merely omits a key
       mismatches what main() derives on the next run, and main() responds by
       rebuilding every option's rows from fresh $0 seeds while leaving the
       option names alone -- which reads as "the budget didn't load" and is
       almost impossible to diagnose from the symptom. That bug shipped once
       already on the draft path; the fix was one shared constructor, and
       this path uses it too.
    2. Every restored plan row is marked dirty. Rows are only "clean" in the
       sense of never-hand-touched, and a re-seed from the shared
       Audience/Geography/Flight fields overwrites clean ones -- which for a
       loaded proposal would silently discard the very numbers being loaded.
    """
    form = row.get("form_json") or {}
    selections = form.get("selections") or {}
    notes = []
    updates = {}

    updates["client_name"] = row.get("client_name") or "Client"
    updates["proposal_title"] = form.get("proposal_title") or "CTV Strategy"

    market = selections.get("market") or row.get("market")
    if market in ("DC", "Harrisburg"):
        updates["market_choice"] = market
    market_label = ("Washington, DC DMA" if market == "DC"
                    else "Harrisburg DMA" if market == "Harrisburg" else "")

    # selections["vertical"] is "none" when the vertical's own slides were
    # switched off, so the real vertical comes off the column and the toggle
    # is derived from whether the two agree.
    vertical_key = row.get("vertical") or "none"
    vertical_label = next((label for label, key in VERTICALS.items() if key == vertical_key), None)
    if vertical_label:
        updates["vertical_choice"] = vertical_label
    if vertical_key != "none":
        updates["include_vertical_slides"] = selections.get("vertical") != "none"

    preset_key = selections.get("preset")
    preset_label = next((label for label, key in PRESETS.items() if key == preset_key), None)
    if preset_label:
        updates["preset"] = preset_label

    updates["agency_involved"] = bool(form.get("agency_involved", selections.get("agency_involved")))
    updates["spanish_campaign"] = bool(selections.get("spanish_campaign"))
    updates["tegna_positioning"] = bool(selections.get("tegna_positioning"))
    updates["include_avails_template"] = bool(selections.get("include_avails_template", True))

    # --- products (Section C) -------------------------------------------
    products = selections.get("products") or {}
    sr = products.get("streaming_retargeting") or {}
    am = products.get("audience_marketplace") or {}
    sports = products.get("live_sports") or {}
    updates["premion_streaming_tv"] = bool(products.get("_premion_streaming_tv",
                                                        selections.get("_premion_streaming_tv", True)))
    updates["streaming_retargeting_enabled"] = bool(sr.get("enabled"))
    updates["sr_disp"] = bool(sr.get("display"))
    updates["sr_pre"] = bool(sr.get("preroll"))
    updates["am_enabled"] = bool(am.get("enabled"))
    updates["am_at_disp"] = bool(am.get("audience_targeting_display"))
    updates["am_at_pre"] = bool(am.get("audience_targeting_preroll"))
    updates["am_gf_disp"] = bool(am.get("geofencing_display"))
    updates["am_gf_pre"] = bool(am.get("geofencing_preroll"))
    updates["am_srd"] = bool(am.get("site_retargeting_display"))
    updates["am_srp"] = bool(am.get("site_retargeting_preroll"))
    updates["total_tv"] = bool(products.get("total_tv"))
    updates["dynamic_creative"] = bool(selections.get("dynamic_creative",
                                                      products.get("dynamic_creative")))
    updates["live_sports_enabled"] = bool(sports.get("enabled"))
    sport_keys = set(sports.get("sports") or [])
    updates["selected_sports"] = [label for label, key in SPORTS.items() if key in sport_keys]
    missing_sports = sport_keys - set(SPORTS.values())
    if missing_sports:
        notes.append(f"Sports package(s) no longer offered were dropped: {', '.join(sorted(missing_sports))}.")

    # `_premion_streaming_tv` is stored one level up in the seed key's own
    # shape; older rows may not carry it at all.
    if "_premion_streaming_tv" in selections:
        updates["premion_streaming_tv"] = bool(selections["_premion_streaming_tv"])

    # --- attribution (Section D) ----------------------------------------
    targeting = selections.get("targeting_attribution") or {}
    updates["first_party_data"] = bool(targeting.get("first_party_data"))
    updates["linear_reach_extension"] = bool(targeting.get("linear_reach_extension"))
    updates["sales_attribution"] = bool(targeting.get("sales_attribution"))
    updates["brand_lift"] = bool(targeting.get("brand_lift"))
    included = form.get("included_list") or []
    updates["commercial_production"] = any("Commercial Production" in item for item in included)

    # --- Campaign Specs copy --------------------------------------------
    specs = form.get("campaign_specs") or {}
    for field in ("goals", "audience", "geography", "budget", "placements", "timing"):
        updates[f"{field}_text"] = specs.get(field) or ""

    # --- flight ----------------------------------------------------------
    flight = form.get("flight") or {}
    start = _parse_draft_date(flight.get("start"))
    end = _parse_draft_date(flight.get("end"))
    if start:
        updates["flight_start"] = start
    if end:
        updates["flight_end"] = end
    if start and end:
        all_months = month_list(start, end)
        stored_active = [m for m in (flight.get("active_months") or []) if m in all_months]
        updates["active_months"] = stored_active or all_months
        flight_label = format_flight_label(all_months, updates["active_months"]) or "TBD"
    else:
        notes.append("Flight dates couldn't be restored -- check Section B before generating.")
        flight_label = flight.get("label") or "TBD"

    # --- avails ----------------------------------------------------------
    avails_rows = form.get("avails_rows") or []
    real_avails = [r for r in avails_rows if str(r.get("audience", "")).strip()]
    if real_avails:
        updates["avails_seed_rows"] = [
            {"Audience": r.get("audience", ""), "Geo": r.get("geo", "") or market_label,
             "Max Monthly Avails": int(str(r.get("avails", "0")).replace(",", "") or 0)}
            for r in real_avails
        ]
        updates["avails_version"] = st.session_state.get("avails_version", 0) + 1

    # --- media plan options ----------------------------------------------
    stored_options = form.get("plan_options") or []
    plan_options = []
    for stored in stored_options[:MAX_PLAN_OPTIONS]:
        rows = [dict(r) for r in (stored.get("rows") or [])]
        driver = list(stored.get("driver") or [])
        if len(driver) != len(rows):
            driver = [DRIVER_COST] * len(rows)
        option = new_plan_option(stored.get("name") or DEFAULT_OPTION_NAMES[0], rows,
                                 driver=driver,
                                 breakout=stored.get("breakout") or BREAKOUT_MONTHLY)
        # Every restored row is a deliberate value, never a seed default.
        option["dirty"] = [True] * len(rows)
        plan_options.append(option)
    if plan_options:
        updates["plan_options"] = plan_options
        updates["media_plan_markup"] = form.get("markup", 1.15 if updates["agency_involved"] else 1.0)
        # Fresh widget keys, or the existing per-option name/breakout widgets
        # would overwrite the restored ones with whatever they already hold.
        bump_plan_options_generation(updates)

    # --- case studies -----------------------------------------------------
    for key in [k for k in st.session_state if k.startswith("cs_pick_")]:
        updates[key] = False
    for case_study in form.get("case_studies") or []:
        updates[f"cs_pick_{case_study.get('id')}"] = True

    # --- the notes this proposal came from --------------------------------
    draft_info = form.get("draft") or {}
    if draft_info.get("notes"):
        updates["draft_notes_input"] = draft_info["notes"]

    # --- bookkeeping main() re-derives on the next run --------------------
    def _get(key, default=False):
        return updates.get(key, st.session_state.get(key, default))

    default_targeting = first_line(updates.get("audience_text", ""))
    default_geo = first_line(updates.get("geography_text", "")) or market_label
    updates["_product_seed_key"] = str(read_seed_selections(_get))
    updates["_shared_fields_key"] = default_targeting + "||" + default_geo + "||" + flight_label

    # --- history linkage --------------------------------------------------
    # Cleared on the next successful generate, so a loaded proposal links its
    # regeneration back to its source exactly once.
    updates["history_parent_id"] = parent_proposal_id
    updates["history_rebuild_version"] = rebuild_deck_version_id
    updates["history_loaded_from"] = row.get("id")

    # --- client logo ------------------------------------------------------
    # A file_uploader can't be prefilled, so the stored logo travels as a
    # local path the generate path uses when nothing new is uploaded. The two
    # failure modes are kept distinct on purpose: a proposal that never had a
    # logo is not the same as one whose logo can't be fetched, and only the
    # second is a problem worth flagging.
    updates["restored_logo_path"] = None
    updates["restored_logo_storage_path"] = None
    updates["restored_logo_missing"] = False
    logo_path = row.get("logo_storage_path")
    if logo_path:
        try:
            updates["restored_logo_path"] = db.proposal_logo(logo_path)
            updates["restored_logo_storage_path"] = logo_path
        except Exception as exc:
            updates["restored_logo_missing"] = True
            notes.append(f"This proposal's logo couldn't be fetched from storage "
                         f"({db.describe_error(exc)}) -- upload it again or the deck will use "
                         f"the placeholder.")
    elif (row.get("form_json") or {}).get("logo_used"):
        notes.append("This proposal predates logo storage, so its logo isn't kept -- "
                     "re-upload it if this deck needs one.")

    # --- revision label ---------------------------------------------------
    # Only prefilled on a load, where a label is actually worth having; a
    # from-scratch build leaves it blank.
    if revision_default:
        updates["revision_label"] = revision_default

    # A draft badge from a previous session would be misleading here: none of
    # this came from a draft in *this* session.
    st.session_state["ai_filled_sections"] = set()
    st.session_state["draft_unresolved"] = []
    st.session_state["draft_round"] = None
    st.session_state["draft_last_json"] = None

    for key, value in updates.items():
        st.session_state[key] = value
    return notes


def apply_draft_to_form(draft, skip_sections=None):
    """Turns a parsed Claude draft into session_state writes (applied all at
    once at the end, so a mid-processing error leaves the form untouched)
    plus an "unresolved" list shown to the user. Must be called before any
    widget in this run has rendered, followed by st.rerun().

    `skip_sections` names form sections whose writes should be dropped --
    used by the clarify-and-re-draft round to leave sections the user
    hand-edited between the two drafts exactly as they left them. Same
    principle as the media plan's per-row dirty flags, applied per section.
    """
    skip_sections = set(skip_sections or ())
    unresolved = list(draft.get("unresolved", []))
    updates = {}
    touched_sections = set()

    vertical_val = draft.get("vertical")
    vertical_label = next((label for label, key in VERTICALS.items() if key == vertical_val), None)
    if vertical_label:
        updates["vertical_choice"] = vertical_label
    elif vertical_val:
        unresolved.append(f"Vertical '{vertical_val}' not recognized -- left unchanged.")

    market_val = draft.get("market")
    if market_val in ("DC", "Harrisburg"):
        updates["market_choice"] = market_val
    elif market_val:
        unresolved.append(f"Market '{market_val}' not recognized -- left unchanged.")
    market_label = "Washington, DC DMA" if market_val == "DC" else "Harrisburg DMA" if market_val == "Harrisburg" else ""

    updates["client_name"] = draft.get("client_name") or "Client"
    # The markup that prices every plan row has to follow what the form will
    # actually hold, not what this particular draft happens to say: if
    # "basics" is being skipped, the agency toggle keeps its existing value,
    # so pricing the rows off the draft's own field would contradict the
    # toggle sitting right there on screen.
    if "basics" in skip_sections:
        agency_involved = bool(st.session_state.get("agency_involved", False))
    else:
        agency_involved = bool(draft.get("agency_involved", False))
    updates["agency_involved"] = agency_involved
    updates["spanish_campaign"] = bool(draft.get("spanish_campaign", False))
    touched_sections.add("basics")

    flight_start = _parse_draft_date(draft.get("flight_start"))
    flight_end = _parse_draft_date(draft.get("flight_end"))
    if flight_start:
        updates["flight_start"] = flight_start
    else:
        unresolved.append("Flight start date missing or unparseable -- left unchanged.")
    if flight_end:
        updates["flight_end"] = flight_end
    else:
        unresolved.append("Flight end date missing or unparseable -- left unchanged.")
    if flight_start and flight_end:
        # Match main()'s own format_flight_label(all_months, all_months) exactly
        # (byte-for-byte, including the single-month case) so the
        # _shared_fields_key we precompute below actually matches what main()
        # derives after rerun -- a mismatch would trigger main()'s own
        # reseed-on-change logic and silently discard these drafted rows.
        draft_months = month_list(flight_start, flight_end)
        flight_label = format_flight_label(draft_months, draft_months) or "TBD"
        draft_n_months = max(1, len(draft_months))
    else:
        flight_label = "TBD"
        # No usable dates -- treat the plan as a single period rather than
        # dividing a budget by a month count we don't actually know.
        draft_n_months = 1
    touched_sections.add("flight")

    geo = (draft.get("geo") or "").strip()

    specs = draft.get("campaign_specs", {}) or {}
    text_fields = (("goals", "goals_text"), ("audience", "audience_text"), ("budget", "budget_text"),
                   ("placements", "placements_text"), ("timing", "timing_text"))
    for spec_field, widget_key in text_fields:
        values = specs.get(spec_field) or []
        if values:
            updates[widget_key] = "\n".join(values)
    geo_bullets = specs.get("geography") or ([geo] if geo else [market_label] if market_label else [])
    if geo_bullets:
        updates["geography_text"] = "\n".join(geo_bullets)
    touched_sections.add("specs")

    # Mirror main()'s own default_targeting/default_geo derivation exactly
    # (first_line() over the same joined text) so downstream keys computed
    # here -- shared_fields_key in particular -- match what main() derives
    # after rerun instead of silently diverging on a bullet-list edge case.
    default_targeting = first_line(updates.get("audience_text", ""))
    geo_or_market = first_line(updates.get("geography_text", "")) or market_label

    attribution = set(draft.get("attribution", []))
    for json_key, widget_key in ATTRIBUTION_FIELD_MAP.items():
        updates[widget_key] = json_key in attribution
    touched_sections.add("attribution")

    # Collected here but applied after the media plan lines are resolved --
    # a "sport:<key>" line implies its package too, and the two sources are
    # unioned so a sports line can't end up billed without its deck slides
    # (or vice versa).
    drafted_sport_keys = []
    for s in draft.get("sports", []) or []:
        if s in SPORT_LABEL_BY_VALUE:
            drafted_sport_keys.append(s)
        else:
            unresolved.append(f"Sport '{s}' not recognized -- skipped.")

    # --- audiences ---
    audiences_in = draft.get("audiences", []) or []
    audience_names = [a.get("segment", "") for a in audiences_in if a.get("segment")]
    matched, unmatched = validate_segments(audience_names)
    if unmatched:
        unresolved.append(f"Unrecognized audience segment(s) dropped: {', '.join(unmatched)}")
    matched_audiences = [a for a in audiences_in if a.get("segment") in matched]

    # One custom (non-RFP-selectable) audience per campaign, enforced here
    # rather than asked for in the prompt. The prompt still asks -- it makes
    # the model's first pass better, and the segment it keeps is then the one
    # it judged most important -- but a prompt instruction is a probability,
    # not a guarantee: the same notes returned one custom segment on one run
    # and two on the next. Doing it in Python makes the limit hold whatever
    # comes back. The first custom segment survives (the model orders them by
    # relevance) and the rest move into unresolved as named alternatives,
    # because they're genuinely useful suggestions -- a reviewer swapping the
    # kept one out wants to know what else was considered. Nothing is
    # silently discarded.
    catalog = load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
    kept_audiences, dropped_custom, seen_custom = [], [], False
    for audience in matched_audiences:
        if rfp_map.get(audience["segment"], True):
            kept_audiences.append(audience)
            continue
        if seen_custom:
            dropped_custom.append(audience["segment"])
        else:
            seen_custom = True
            kept_audiences.append(audience)
    if dropped_custom:
        kept_custom = next(a["segment"] for a in kept_audiences
                           if not rfp_map.get(a["segment"], True))
        unresolved.append(
            f"Only one custom (non-RFP-selectable) audience is allowed per campaign, so "
            f"\"{kept_custom}\" was kept and {len(dropped_custom)} other custom segment(s) were "
            f"left out of the avails table: {', '.join(dropped_custom)}. Any of them can be "
            f"swapped in for the kept one -- but only one custom segment can ship.")
    matched_audiences = kept_audiences

    if matched_audiences:
        updates["avails_seed_rows"] = [
            {"Audience": a["segment"], "Geo": a.get("geo") or geo_or_market, "Max Monthly Avails": 0}
            for a in matched_audiences
        ]
        updates["avails_version"] = st.session_state.get("avails_version", 0) + 1
        unresolved.append("Max Monthly Avails left at 0 for drafted audiences -- pull real numbers from the avails system before finalizing.")
        touched_sections.add("avails")

    # --- budget math (no arithmetic performed by the model) ---
    # One drafted option per intended media plan slide; each resolves its own
    # lines against its own budget. A single-scenario draft is a one-option
    # proposal and renders with no option label anywhere.
    markup = 1.15 if agency_involved else 1.0
    options_in = drafted_options(draft)
    drafted_plan_options = []
    lines_valid = False
    touched_products = set()
    touched_sports = set()

    for opt_in in options_in:
        # An option that names lines but carries no budget can only resolve to
        # a grid of zeros. Every such line then trips the $0 drop below and
        # the option disappears -- so say plainly what went wrong and what to
        # do about it, rather than reporting a bare "no usable lines" for what
        # is really a missing number. A drafted plan must never come back as a
        # stated budget with nothing costed against it.
        if opt_in["lines"] and opt_in["total_budget"] <= 0:
            fixed_only = [line for line in opt_in["lines"]
                          if "flat_amount" in (line.get("allocation") or {})]
            if fixed_only:
                unresolved.append(
                    f'"{opt_in["name"]}" has no overall budget, so only its fixed-amount '
                    f"line(s) could be priced -- anything meant to share a remaining budget "
                    f"was dropped. Say what this plan should cost and re-draft.")
            else:
                unresolved.append(
                    f'"{opt_in["name"]}" came back with {len(opt_in["lines"])} media plan '
                    f"line(s) but no budget to spend on them, so nothing could be priced and "
                    f"the option was dropped. Say what this plan should cost and re-draft, or "
                    f"add the lines by hand.")

        rows, opt_products, opt_sports, opt_unresolved = resolve_drafted_lines(
            opt_in["lines"], opt_in["total_budget"], markup,
            flight_label, geo_or_market, default_targeting)
        unresolved.extend(opt_unresolved)
        touched_products |= opt_products
        touched_sports |= opt_sports
        if rows:
            lines_valid = True
            # Two independent things: the budget is always a campaign total,
            # and the breakout is how that plan is presented. Under a Monthly
            # breakout the rows hold per-month figures, so the resolved
            # full-flight amounts are divided across the flight -- which
            # keeps the full-flight total on the number the notes quoted
            # instead of multiplying it by the month count.
            breakout = opt_in["breakout"]
            if breakout == BREAKOUT_MONTHLY:
                spread_rows_over_months(rows, draft_n_months, markup)
            option = new_plan_option(opt_in["name"], rows, driver=[DRIVER_COST] * len(rows),
                                     breakout=breakout)
            option["dirty"] = [True] * len(rows)  # drafted rows are deliberate, never re-seeded away
            drafted_plan_options.append(option)

    kept_names = {o["name"] for o in drafted_plan_options}
    for opt_in in options_in:
        # Named, so a two-option draft that loses one says which. The
        # budget-specific diagnosis above already fired for the case that
        # causes this most often; this catches the rest (every line dropped
        # as an unrecognized product, an empty line list, and so on).
        if opt_in["name"] not in kept_names and opt_in["total_budget"] > 0:
            unresolved.append(f'"{opt_in["name"]}" had no usable media plan lines and was dropped.')

    # The drafted plan is authoritative over Section C: every product toggle
    # is cleared first, then only the ones the drafted lines actually use are
    # switched back on. Without this, Section C's own defaults (Premion
    # Streaming TV is checked out of the box) would survive a draft that
    # never asked for them -- an NFL-only proposal would still carry a
    # Premion Streaming TV line and its deck slides.
    if lines_valid:
        for product, widget_keys in PRODUCT_TO_WIDGET_KEYS.items():
            # A product that can't be a media plan line can't appear in
            # touched_products either, so clearing it here would switch it off
            # permanently. Dynamic Video Ads are the case: the draft enables
            # them through `attribution` (they're a creative build, not
            # impressions), and that write happens earlier in this function --
            # clearing every toggle indiscriminately silently undid it, so a
            # draft asking for dynamic ads produced a deck without the slide.
            if product in NON_LINE_PRODUCT_KEYS:
                continue
            for widget_key, _ in widget_keys:
                updates[widget_key] = False
        for p in touched_products:
            for widget_key, val in PRODUCT_TO_WIDGET_KEYS[p]:
                updates[widget_key] = val
        touched_sections.add("products")

    # Sports: union of the "sports" array and any "sport:<key>" plan lines,
    # in SPORTS' own declaration order so the multiselect reads consistently.
    all_sport_keys = set(drafted_sport_keys) | touched_sports
    if all_sport_keys:
        updates["live_sports_enabled"] = True
        updates["selected_sports"] = [label for label, key in SPORTS.items() if key in all_sport_keys]
        touched_sections.add("products")
    elif lines_valid:
        updates["live_sports_enabled"] = False
        updates["selected_sports"] = []

    if drafted_plan_options:
        updates["plan_options"] = drafted_plan_options
        updates["media_plan_markup"] = markup
        # Move the per-option name/breakout widgets to fresh keys, or the
        # drafted names would be overwritten by whatever the existing widgets
        # already hold (see bump_plan_options_generation).
        bump_plan_options_generation(updates)
        touched_sections.add("media_plan")

        # Prevent main()'s own reseed-on-mismatch logic from immediately
        # overwriting these rows on the very next run: precompute the same
        # product_seed_key/shared_fields_key it will independently derive
        # after rerun, from the *drafted* widget values in `updates` (falling
        # back to current session_state for anything the draft didn't touch)
        # -- not from session_state alone, which still holds pre-draft values
        # at this point in the function.
        def _get(key, default=False):
            return updates.get(key, st.session_state.get(key, default))

        # Built by read_seed_selections itself, never re-assembled here --
        # see its docstring for the bug a hand-rolled copy caused.
        updates["_product_seed_key"] = str(read_seed_selections(_get))
        updates["_shared_fields_key"] = default_targeting + "||" + geo_or_market + "||" + flight_label

    if skip_sections:
        preserved = sorted(s for s in skip_sections if s in touched_sections)
        if preserved:
            unresolved.append(
                "Kept your own edits to these section(s) instead of overwriting them with the "
                f"re-draft: {', '.join(preserved)}.")
        touched_sections -= skip_sections

    st.session_state["ai_filled_sections"] = touched_sections
    # Options are variants of one plan, so a per-line note (a negotiated CPM,
    # a dropped $0 line) is usually raised identically by every option. The
    # reviewer needs to read it once; anything genuinely option-specific
    # names its option and so is already distinct.
    st.session_state["draft_unresolved"] = list(dict.fromkeys(unresolved))

    for key, value in updates.items():
        if DRAFT_KEY_SECTIONS.get(key) in skip_sections:
            continue
        st.session_state[key] = value


def _round_dollar_group(raw_amounts):
    """Rounds each {index: amount} in a percentage-split group to whole
    dollars, then nudges whichever line rounded to the largest amount by
    however many cents were lost or gained in rounding the others -- so the
    group's rounded sum exactly matches its target (e.g. a 60/40 split of an
    odd remainder can't quietly end up a dollar short)."""
    if not raw_amounts:
        return {}
    target = round(sum(raw_amounts.values()))
    rounded = {i: round(a) for i, a in raw_amounts.items()}
    diff = target - sum(rounded.values())
    if diff:
        largest_i = max(rounded, key=rounded.get)
        rounded[largest_i] += diff
    return rounded


def _resolve_line_cpm(line, tactic, default_cpm):
    """(cpm, note) for one drafted line. Rates are negotiated per deal, so a
    line may carry its own "cpm" and it wins over the rate card's default.

    Always returns a note when an override is in play: a rate that didn't
    come from the rate card is exactly the kind of thing a reviewer has to
    see, since nothing downstream distinguishes a negotiated CPM from a
    standard one. A malformed or non-positive override falls back to the
    default rather than pricing the line at zero.
    """
    raw = line.get("cpm")
    if raw is None or raw == "":
        return default_cpm, None
    try:
        cpm = float(raw)
    except (TypeError, ValueError):
        return default_cpm, (f'Ignored an unreadable CPM ("{raw}") on the "{tactic}" line -- '
                             f"priced at the ${default_cpm:,.2f} rate card default instead.")
    if cpm <= 0:
        return default_cpm, (f'Ignored a ${cpm:,.2f} CPM on the "{tactic}" line -- a rate line '
                             f"can't price at zero, so the ${default_cpm:,.2f} rate card default "
                             f"was used instead.")
    return cpm, (f'"{tactic}" is priced at a negotiated ${cpm:,.2f} CPM from the notes, not the '
                 f"${default_cpm:,.2f} rate card default -- confirm the rate before sending.")


def resolve_drafted_lines(lines_in, total_budget, markup, flight_label, geo_or_market, default_targeting):
    """Turn one option's worth of drafted media_plan_lines into real media
    plan rows. Returns (rows, touched_products, touched_sports, unresolved).

    No arithmetic comes from the model -- it returns intent, this resolves it
    as a waterfall: flat_amount and percent_of_total lines come out of
    total_budget first, percent_of_remainder lines then take their share of
    what's left, and split_evenly lines divide whatever remains evenly among
    themselves.
    """
    unresolved = []
    valid_line_products = (set(PRODUCT_TO_WIDGET_KEYS) - NON_LINE_PRODUCT_KEYS) | {CUSTOM_FEE_PRODUCT}
    valid_sport_keys = set(SPORTS.values())

    lines_valid = []
    for line in lines_in or []:
        product = line.get("product") or ""
        if product.startswith(SPORT_PRODUCT_PREFIX):
            if product[len(SPORT_PRODUCT_PREFIX):] not in valid_sport_keys:
                unresolved.append(f"Media plan line for unrecognized sports package '{product}' skipped.")
                continue
        elif product not in valid_line_products:
            unresolved.append(f"Media plan line for unrecognized product '{product}' skipped.")
            continue
        lines_valid.append(line)

    resolved_amounts = {}

    # Stage 1: flat_amount and percent_of_total lines, each rounded on its
    # own -- there's no shared pool these are jointly required to exhaust,
    # unlike a percentage split of a remainder.
    for i, line in enumerate(lines_valid):
        alloc = line.get("allocation", {}) or {}
        if "flat_amount" in alloc:
            amt = float(alloc["flat_amount"] or 0)
        elif "percent_of_total" in alloc:
            amt = (float(alloc["percent_of_total"] or 0) / 100.0) * total_budget
        else:
            continue
        resolved_amounts[i] = round(amt)

    remainder = max(0, total_budget - sum(resolved_amounts.values()))

    # Stage 2: percent_of_remainder lines split the remainder -- rounded as
    # one group so a 60/40 (or any) split can't drift from the exact dollar
    # amount it's meant to divide up.
    raw_pct_remainder = {}
    for i, line in enumerate(lines_valid):
        alloc = line.get("allocation", {}) or {}
        if "percent_of_remainder" in alloc:
            raw_pct_remainder[i] = (float(alloc["percent_of_remainder"] or 0) / 100.0) * remainder
    resolved_amounts.update(_round_dollar_group(raw_pct_remainder))

    leftover = max(0, remainder - sum(resolved_amounts.get(i, 0) for i in raw_pct_remainder))

    # Stage 3: split_evenly lines share whatever's left -- same group
    # rounding, since an uneven leftover (e.g. $100 over 3 lines) would
    # otherwise round every share down and leave a dollar unaccounted for.
    split_evenly_indices = [i for i, line in enumerate(lines_valid)
                            if (line.get("allocation") or {}).get("split_evenly")]
    if split_evenly_indices:
        per_each_raw = leftover / len(split_evenly_indices)
        resolved_amounts.update(_round_dollar_group({i: per_each_raw for i in split_evenly_indices}))

    for i, line in enumerate(lines_valid):
        if i not in resolved_amounts:
            resolved_amounts[i] = 0
            unresolved.append(f"Media plan line for '{line.get('product')}' has no recognized "
                              f"allocation type -- amount left at $0.")

    rows = []
    touched_products = set()
    touched_sports = set()
    for i, line in enumerate(lines_valid):
        product = line["product"]
        label = line.get("label", "") or ""
        audience_track = line.get("audience_track", "") or ""
        amount = resolved_amounts[i]

        # A line that resolves to nothing is not a line. This showed up as a
        # "Commercial Production (:30 spot) — $0" row in a generated deck:
        # production is normally included at no charge, and the model
        # expressed that as a zero-dollar fee line rather than as an entry in
        # the Included with Campaign list. A $0 row is never meaningful for
        # any product, so it's dropped here regardless of how it arose.
        if not amount:
            unresolved.append(
                f"Dropped the \"{label or product}\" media plan line -- it worked out to $0. "
                f"Anything included at no charge belongs in the Included with Campaign list, "
                f"not the plan grid; add a line by hand if it should carry a real cost.")
            continue

        if product == CUSTOM_FEE_PRODUCT:
            # A flat fee has no rate to negotiate -- its cost is the fee.
            if line.get("cpm") not in (None, ""):
                unresolved.append(f'Ignored a CPM on the flat-fee "{label or "Flat Fee"}" line -- '
                                  f"a one-time fee is billed at its amount, not by impressions.")
            rows.append({
                "Tactic": label or "Flat Fee", "Flight": flight_label, "Geo": geo_or_market,
                "Targeting": audience_track, "Impressions": 0.0, "CPM": 0.0,
                "Type": ROW_TYPE_FLAT_FEE, "Cost": float(amount),
            })
            continue

        if product.startswith(SPORT_PRODUCT_PREFIX):
            touched_sports.add(product[len(SPORT_PRODUCT_PREFIX):])
        else:
            touched_products.add(product)
        base_label, cpm = line_product_spec(product)

        tactic = f"{base_label} — {label}" if label else base_label
        # Negotiated rate beats the rate card, and always says so.
        cpm, cpm_note = _resolve_line_cpm(line, tactic, cpm)
        if cpm_note:
            unresolved.append(cpm_note)
        rows.append({
            "Tactic": tactic, "Flight": flight_label, "Geo": geo_or_market,
            # Fall back to the same per-tactic default the form itself uses
            # (fixed copy for Streaming Retargeting / Live Sports rows,
            # otherwise the Campaign Specs Audience line) rather than always
            # reaching for the audience field.
            "Targeting": audience_track or resolve_row_defaults(
                tactic, geo_or_market, default_targeting, flight_label)["Targeting"],
            # Cost is the driver for a drafted row: the resolved allocation is
            # already whole dollars (group-rounded so the split ties exactly to
            # the budget), so deriving impressions from it keeps the plan total
            # on the number the notes asked for. Deriving cost from rounded
            # impressions instead would reintroduce the drift.
            "Impressions": impressions_from_cost(amount, cpm, markup), "CPM": cpm,
            "Type": ROW_TYPE_RATE, "Cost": float(amount),
        })

    return rows, touched_products, touched_sports, unresolved


def _drafted_breakout(value, fallback=BREAKOUT_MONTHLY):
    """Map the draft's breakout hint onto a real breakout mode. Monthly is
    the default: a media plan is normally presented month by month, and
    that's independent of the budget itself always being a campaign total."""
    if isinstance(value, str) and value.strip().lower().replace(" ", "_") in ("full_flight", "flight", "total"):
        return BREAKOUT_FULL_FLIGHT
    if isinstance(value, str) and value.strip().lower() == "monthly":
        return BREAKOUT_MONTHLY
    return fallback


def drafted_options(draft):
    """Normalize a draft's plan into a list of {name, total_budget, breakout,
    lines}.

    The model may return several named scenarios in "options", or a single
    plan as a bare "media_plan_lines" -- a single-scenario draft stays a
    one-option proposal, which renders with no option label at all.
    """
    top_budget = round(float(draft.get("total_budget") or 0))
    top_breakout = _drafted_breakout(draft.get("breakout"))
    raw_options = draft.get("options") or []
    if raw_options:
        return [
            {
                "name": (opt.get("name") or DEFAULT_OPTION_NAMES[min(i, len(DEFAULT_OPTION_NAMES) - 1)]).strip(),
                "total_budget": round(float(opt.get("total_budget") or top_budget or 0)),
                "breakout": _drafted_breakout(opt.get("breakout"), top_breakout),
                "lines": opt.get("media_plan_lines") or [],
            }
            for i, opt in enumerate(raw_options[:MAX_PLAN_OPTIONS])
        ]
    return [{"name": DEFAULT_OPTION_NAMES[0], "total_budget": top_budget,
             "breakout": top_breakout, "lines": draft.get("media_plan_lines") or []}]


def spread_rows_over_months(rows, n_months, markup):
    """Convert full-flight row amounts into per-month ones.

    resolve_drafted_lines always works in campaign totals -- a budget in the
    notes is what the whole flight costs. A Monthly-breakout option holds
    per-month figures instead, so each rate row is divided by the month
    count. The division is deliberately *not* rounded to whole dollars:
    rounding here would make month x count miss the quoted budget (a $100K
    plan over 3 months would total $99,999), and it's the full-flight total
    that has to tie to what the client was told. The Cost column displays
    whole dollars regardless.

    Flat fees are one-time costs and are never divided -- compute_plan_totals
    already treats them as full-flight amounts under either breakout.
    """
    if n_months <= 1:
        return rows
    for row in rows:
        if is_flat_fee_row(row):
            continue
        row["Cost"] = _num(row["Cost"]) / n_months
        row["Impressions"] = impressions_from_cost(row["Cost"], row["CPM"], markup)
    return rows


def _add_segment_to_avails(segment, geo, current_avails_df):
    rows = current_avails_df.to_dict("records")
    rows.append({"Audience": segment, "Geo": geo, "Max Monthly Avails": 0})
    st.session_state["avails_seed_rows"] = rows
    st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
    st.rerun()


def render_audience_finder(avails_df, market_label, vertical_key=None):
    """Section D2's 'Audience finder': browse/search the catalog, or describe
    the client/campaign and let Claude suggest segments. "Add" appends a row
    to the current avails table (audience + geo filled, avails left blank --
    those come from a real system).

    Wrapped in an expander here because it's a side tool inside a long form;
    the standalone page calls the same body without one. Same component, two
    entry points.
    """
    with st.expander("🔍 Audience finder", expanded=False):
        _audience_finder_body(avails_df, market_label, vertical_key)


def render_audience_finder_page():
    """The standalone Audience finder page.

    Reps look segments up without building a proposal, so this is the same
    finder with no avails table behind it -- which means no "Add" buttons,
    since there'd be nothing to add to. Adding stays in the proposal flow.
    """
    st.header("Audience finder")
    st.caption("Look up the audience segments Premion can target. Search the catalog directly, "
               "or describe a client and let Claude suggest the segments that fit. This page is "
               "for looking things up — to actually add segments to a proposal, use the same "
               "finder inside Section D2 on the Build page.")
    # If the rep has a vertical selected on the proposal page, use it as the
    # same soft hint it is there: it sorts relevant categories forward and
    # seeds Claude's slice. It never filters, so an unrelated lookup still
    # finds everything.
    hint = VERTICALS.get(st.session_state.get("vertical_choice", "None"), "none")
    _audience_finder_body(None, None, hint)


def _audience_finder_body(avails_df, market_label, vertical_key=None):
    """The finder itself. `avails_df is None` means standalone: show the
    catalog and the recommendations, but no Add buttons.

    Works with or without a vertical selected: `vertical_key` only sorts
    vertical-relevant categories to the front (browse) and seeds Claude's
    catalog slice when the description itself doesn't imply a vertical
    (suggest). Its absence filters nothing out."""
    catalog = load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
    vertical_hint = vertical_key if vertical_key and vertical_key != "none" else None
    can_add = avails_df is not None

    if catalog.empty:
        st.info("The audience catalog is empty -- it couldn't be loaded from Supabase and "
                "the local brochure isn't available.")
        return

    mode = st.radio("Mode", ["Browse / search", "Suggest"], horizontal=True, key="finder_mode")

    if mode == "Browse / search":
        cat_options = ["All"] + sorted(catalog["category"].unique().tolist())
        picked_cat = st.selectbox("Category", cat_options, key="finder_category")
        search_text = st.text_input("Search by name", key="finder_search")

        filtered = catalog
        if picked_cat != "All":
            filtered = filtered[filtered["category"] == picked_cat]
        if search_text.strip():
            filtered = filtered[filtered["segment"].str.contains(search_text.strip(), case=False, na=False)]
        total_matches = len(filtered)
        filtered = prioritize_catalog(filtered, vertical_hint).head(50)
        if vertical_hint and picked_cat == "All":
            st.caption(f"{total_matches} segment(s) match; showing 50 "
                       f"(categories relevant to the selected vertical first, then by times used)")
        else:
            st.caption(f"{total_matches} segment(s) match; showing 50 (by times used)")

        header_cols = st.columns([4, 1.5, 2.5, 1.5, 1, 1])
        for col, label in zip(header_cols, ["Segment", "Category", "Subcategory", "Status", "Used", ""]):
            col.caption(f"**{label}**")
        for _, row in filtered.iterrows():
            cols = st.columns([4, 1.5, 2.5, 1.5, 1, 1])
            cols[0].write(row["segment"])
            cols[1].write(row["category"])
            cols[2].write(row["subcategory"] or "--")
            cols[3].write("RFP" if row["rfp_selectable"] else "Custom")
            cols[4].write(f"{row['times_used']:,}")
            if can_add and cols[5].button("Add", key=f"finder_add_{row['segment']}"):
                _add_segment_to_avails(row["segment"], market_label, avails_df)

    else:
        description = st.text_area("Describe the client or campaign", key="finder_suggest_input", height=100)
        if st.button("Suggest audiences"):
            if not description.strip():
                st.warning("Describe the client or campaign first.")
            else:
                with st.spinner("Asking Claude for audience recommendations..."):
                    recs, unmatched, error = call_claude_audience_suggest(description, vertical_hint)
                if error:
                    st.error(error)
                    st.session_state["finder_suggestions"] = None
                else:
                    st.session_state["finder_suggestions"] = recs
                    st.session_state["finder_suggestions_unmatched"] = unmatched

        suggestions = st.session_state.get("finder_suggestions")
        if suggestions:
            unmatched = st.session_state.get("finder_suggestions_unmatched") or []
            if unmatched:
                st.caption(f"Dropped {len(unmatched)} recommended name(s) not found in the catalog: "
                           + ", ".join(unmatched))

            existing_segments = ([str(s) for s in avails_df["Audience"] if str(s).strip()]
                                 if can_add else [])
            existing_custom = sum(1 for s in existing_segments if not rfp_map.get(s, True))
            rec_custom = sum(1 for r in suggestions if not rfp_map.get(r["segment"], True))
            if existing_custom + rec_custom > 1:
                st.warning(
                    f"This recommendation set includes {rec_custom} custom (non-RFP-selectable) audience(s), "
                    f"and the avails table already has {existing_custom} -- only one custom audience is allowed "
                    f"per campaign. Review before adding all of them.")

            header_cols = st.columns([3, 3, 1.5, 1.5, 1, 1])
            for col, label in zip(header_cols, ["Segment", "Rationale", "Category", "Status", "Used", ""]):
                col.caption(f"**{label}**")
            for rec in suggestions:
                seg = rec["segment"]
                match = catalog[catalog["segment"] == seg]
                if match.empty:
                    continue
                cat_row = match.iloc[0]
                cols = st.columns([3, 3, 1.5, 1.5, 1, 1])
                cols[0].write(seg)
                cols[1].write(rec.get("rationale", ""))
                cols[2].write(cat_row["category"])
                cols[3].write("RFP" if cat_row["rfp_selectable"] else "Custom")
                cols[4].write(f"{cat_row['times_used']:,}")
                if can_add and cols[5].button("Add", key=f"finder_add_suggest_{seg}"):
                    _add_segment_to_avails(seg, market_label, avails_df)


def lines_to_bullets(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


def first_line(text):
    bullets = lines_to_bullets(text)
    return bullets[0] if bullets else ""


def month_list(start, end):
    """['Sep 2026', 'Oct 2026', ...] inclusive, for every calendar month
    between start and end (order-swapped defensively if entered backwards).
    """
    if end < start:
        start, end = end, start
    months = []
    cur = date(start.year, start.month, 1)
    last = date(end.year, end.month, 1)
    while cur <= last:
        months.append(cur.strftime("%b %Y"))
        next_month = cur.month + 1
        next_year = cur.year
        if next_month > 12:
            next_month = 1
            next_year += 1
        cur = date(next_year, next_month, 1)
    return months


def format_flight_label(all_months, active_months):
    if not active_months:
        return ""
    if active_months == all_months and len(active_months) > 1:
        return f"{active_months[0]} - {active_months[-1]}"
    return ", ".join(active_months)


def is_flat_fee_row(row):
    return str(row.get("Type", ROW_TYPE_RATE)) == ROW_TYPE_FLAT_FEE


def _num(value):
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def cost_from_impressions(impressions, cpm, markup):
    """Whole dollars. Markup applies here so the Cost column always reads as
    what the client is billed (gross when the agency toggle is on)."""
    return float(round((_num(impressions) / 1000.0) * _num(cpm) * markup))


def impressions_from_cost(cost, cpm, markup):
    """The exact inverse of cost_from_impressions -- same markup, so typing a
    cost and typing the impressions it implies land on the same row."""
    rate = _num(cpm) * markup
    return float(round((_num(cost) / rate) * 1000)) if rate else 0.0


def recompute_row(row, driver, markup):
    """Recompute whichever side of a rate row isn't driving. Mutates and
    returns the row. Flat-fee rows have neither side -- their Cost is the fee
    itself, and Impressions/CPM are ignored entirely."""
    if is_flat_fee_row(row):
        return row
    if driver == DRIVER_COST:
        row["Impressions"] = impressions_from_cost(row.get("Cost"), row.get("CPM"), markup)
    else:
        row["Cost"] = cost_from_impressions(row.get("Impressions"), row.get("CPM"), markup)
    return row


def line_product_spec(product):
    """(tactic label, CPM) for a media-plan line's product key -- the single
    lookup both the form's own row seeding and the AI draft path go through,
    so the two can't drift and a sports line can never pick up a generic
    product's CPM instead of its ratecard rate."""
    if product.startswith(SPORT_PRODUCT_PREFIX):
        sport_key = product[len(SPORT_PRODUCT_PREFIX):]
        label = SPORT_PRODUCT_LABELS.get(sport_key) or sport_product_label(sport_key)
        return label, SPORT_CPM.get(sport_key, DEFAULT_SPORT_CPM)
    # A product the form knows about but the rate card doesn't (a row not yet
    # seeded, or deactivated) still has to seed a usable line rather than
    # raising -- fall back to the built-in copy, then to a zero-rate row the
    # seller can price by hand.
    spec = PRODUCTS.get(product) or FALLBACK_PRODUCTS.get(product)
    if spec is None:
        return product, 0.0
    return spec["label"], spec["default_cpm"]


def resolve_row_defaults(tactic, default_geo, default_targeting, flight_label):
    """What a row's Flight/Geo/Targeting should be right now, given its
    Tactic name -- used both at initial seed time and to soft-update
    not-yet-edited rows when the shared form fields change.

    Most tactics take their Targeting from the Campaign Specs Audience field,
    but a few carry fixed standard copy instead (Streaming Retargeting, and
    every Live Sports package). That copy is the product's own
    `targeting_copy` from the rate card, matched by tactic-name prefix so a
    row whose name has a draft label appended ("... - Display - Commercial")
    still picks it up. Longest label first, because product labels are
    prefixes of one another.
    """
    targeting = default_targeting
    for label in sorted(TARGETING_COPY_BY_LABEL, key=len, reverse=True):
        if tactic.startswith(label):
            targeting = TARGETING_COPY_BY_LABEL[label]
            break
    return {"Flight": flight_label, "Geo": default_geo, "Targeting": targeting}


def seed_media_plan_rows(selections, market_label, default_targeting, flight_label):
    """Section C -> Section E: each selected product/format seeds a proposal
    line with its default CPM (spec section 5, Section C description).
    Targeting defaults to the Campaign Specs Audience field, except for
    Streaming Retargeting and Live Sports, which have fixed standard
    targeting copy. Geo/Flight default to the Geography/Timing-derived
    values."""
    rows = []
    products = selections["products"]

    def _row(product_key):
        """Seed one line from a rate-card product key. Everything (label,
        CPM, and whether the tactic carries fixed targeting copy) comes from
        the same rate-card lookup the AI draft path uses, so the two can't
        price or name the same product differently."""
        label, cpm = line_product_spec(product_key)
        defaults = resolve_row_defaults(label, market_label, default_targeting, flight_label)
        return {"Tactic": label, "Flight": defaults["Flight"], "Geo": defaults["Geo"],
                "Targeting": defaults["Targeting"], "Impressions": 0.0, "CPM": cpm,
                "Type": ROW_TYPE_RATE, "Cost": 0.0}

    if selections.get("_premion_streaming_tv"):
        rows.append(_row("premion_streaming_tv"))

    sr = products.get("streaming_retargeting", {})
    if sr.get("display"):
        rows.append(_row("streaming_retargeting_display"))
    if sr.get("preroll"):
        rows.append(_row("streaming_retargeting_preroll"))

    am = products.get("audience_marketplace", {})
    if am.get("enabled"):
        for flag, product_key in (
            ("audience_targeting_display", "audience_targeting_display"),
            ("audience_targeting_preroll", "audience_targeting_preroll"),
            ("geofencing_display", "geofencing_display"),
            ("geofencing_preroll", "geofencing_preroll"),
            ("site_retargeting_display", "site_retargeting_display"),
            ("site_retargeting_preroll", "site_retargeting_preroll"),
        ):
            if am.get(flag):
                rows.append(_row(product_key))

    sports = products.get("live_sports", {})
    if sports.get("enabled"):
        for sport_key in sports.get("sports", []):
            rows.append(_row(f"{SPORT_PRODUCT_PREFIX}{sport_key}"))

    if products.get("total_tv"):
        rows.append(_row("broadcast_tv"))

    # Dynamic Video Ads are a one-time production charge, not impressions --
    # a flat-fee row, seeded at the standard rate and editable like any other.
    if products.get("dynamic_creative"):
        rows.append({"Tactic": DYNAMIC_AD_LINE_LABEL, "Flight": flight_label,
                     "Geo": market_label, "Targeting": DYNAMIC_AD_TARGETING,
                     "Impressions": 0.0, "CPM": 0.0,
                     "Type": ROW_TYPE_FLAT_FEE, "Cost": float(DYNAMIC_AD_DEFAULT_FEE)})

    if not rows:
        rows.append({"Tactic": "", "Flight": flight_label, "Geo": market_label,
                     "Targeting": "", "Impressions": 0.0, "CPM": 0.0,
                     "Type": ROW_TYPE_RATE, "Cost": 0.0})

    return rows


def new_plan_option(name, rows, driver=None, breakout=BREAKOUT_MONTHLY):
    """One media plan option: its own name, line set, per-line dirty/driver
    tracking, breakout mode and editor version. Everything the single plan
    used to keep in flat session_state keys now lives per option."""
    return {
        "name": name,
        "rows": rows,
        "dirty": [False] * len(rows),
        "driver": list(driver) if driver else [DRIVER_IMPRESSIONS] * len(rows),
        "breakout": breakout,
        "version": 0,
    }


def copy_plan_option(source, name):
    """A new option cloned from an existing one -- all lines and settings, so
    the user edits the delta instead of rebuilding the plan. The clone's rows
    start dirty: they're a deliberate copy, not a fresh seed, and must not be
    silently re-seeded out from under the user."""
    return {
        "name": name,
        "rows": [dict(r) for r in source["rows"]],
        "dirty": [True] * len(source["rows"]),
        "driver": list(source["driver"]),
        "breakout": source["breakout"],
        "version": 0,
    }


def next_option_name(existing_names):
    for candidate in DEFAULT_OPTION_NAMES:
        if candidate not in existing_names:
            return candidate
    return f"Option {len(existing_names) + 1}"


def reconcile_plan_rows(option, edited_rows, markup):
    """Fold one option's edited grid back into its stored state: update each
    row's dirty flag and driver, then recompute the non-driving side of every
    rate row. Returns True if any value actually changed as a result (the
    caller re-renders once so the grid shows the recomputed numbers)."""
    prev_rows = option["rows"]
    prev_dirty = option["dirty"]
    prev_drivers = option["driver"]

    def _cell_changed(row, prev, field):
        return abs(_num(row.get(field)) - _num(prev.get(field))) > 1e-9

    new_dirty, new_drivers = [], []
    recomputed_any = False
    for i, row in enumerate(edited_rows):
        driver = prev_drivers[i] if i < len(prev_drivers) else DRIVER_IMPRESSIONS
        if i < len(prev_rows):
            prev = prev_rows[i]
            # Whichever side the user just typed into becomes the driver. A
            # CPM-only edit leaves the driver alone and re-derives the other
            # side from it, which is the whole point of tracking this.
            if _cell_changed(row, prev, "Impressions"):
                driver = DRIVER_IMPRESSIONS
            elif _cell_changed(row, prev, "Cost"):
                driver = DRIVER_COST
            changed = any(str(row.get(f, "")) != str(prev.get(f, "")) for f in MEDIA_PLAN_FIELDS)
            new_dirty.append(prev_dirty[i] or changed)
        else:
            new_dirty.append(True)  # a row added via the grid's own "+" is treated as customized

        before = (_num(row.get("Impressions")), _num(row.get("Cost")))
        recompute_row(row, driver, markup)
        if (_num(row.get("Impressions")), _num(row.get("Cost"))) != before:
            recomputed_any = True
        new_drivers.append(driver)

    option["rows"] = edited_rows
    option["dirty"] = new_dirty
    option["driver"] = new_drivers
    return recomputed_any


def compute_plan_totals(rows, breakout_mode, n_months, flight_label):
    """Per-line monthly/full-flight impressions and cost for one option, plus
    the four running totals. Both sides come straight off each row -- they
    were reconciled against each other when the grid was folded back in, so
    reading Cost here (rather than recomputing it) is what makes the preview,
    the totals and the deck all tie to what the grid shows."""
    preview_rows = []
    monthly_impressions_total = monthly_cost_total = 0.0
    flight_impressions_total = flight_cost_total = 0.0

    for row in rows:
        if not str(row.get("Tactic", "")).strip():
            continue
        flat_fee = is_flat_fee_row(row)
        if flat_fee:
            # A flat fee is a one-time full-flight cost, not a per-month rate
            # -- it must NOT scale with month count the way rate rows do, so
            # it's accumulated straight into the flight total.
            full_flight_impressions = 0.0
            full_flight_cost = _num(row.get("Cost"))
            monthly_impressions = 0.0
            monthly_cost = full_flight_cost / n_months
        else:
            entered_impressions = _num(row.get("Impressions"))
            entered_cost = _num(row.get("Cost"))
            if breakout_mode.startswith("Full Flight"):
                full_flight_impressions = entered_impressions
                full_flight_cost = entered_cost
                monthly_impressions = full_flight_impressions / n_months
                monthly_cost = full_flight_cost / n_months
            else:
                monthly_impressions = entered_impressions
                monthly_cost = entered_cost
                full_flight_impressions = monthly_impressions * n_months
                full_flight_cost = monthly_cost * n_months

        monthly_impressions_total += monthly_impressions
        monthly_cost_total += monthly_cost
        flight_impressions_total += full_flight_impressions
        flight_cost_total += full_flight_cost
        preview_rows.append({
            "tactic": str(row["Tactic"]), "flight": str(row.get("Flight", "")) or flight_label,
            "geo": str(row.get("Geo", "")), "targeting": str(row.get("Targeting", "")),
            "monthly_impressions": monthly_impressions, "monthly_cost": monthly_cost,
            "full_flight_impressions": full_flight_impressions, "full_flight_cost": full_flight_cost,
            "is_flat_fee": flat_fee,
        })

    return {
        "preview_rows": preview_rows,
        "monthly_impressions": monthly_impressions_total,
        "monthly_cost": monthly_cost_total,
        "full_flight_impressions": flight_impressions_total,
        "full_flight_cost": flight_cost_total,
    }


def option_plan_title(proposal_title, option_name, multiple_options):
    """A single-option proposal keeps the plain proposal title -- no option
    label anywhere in the deck. Only a multi-option proposal appends one."""
    return f"{proposal_title} — {option_name}" if multiple_options else proposal_title


def build_included_list(targeting, commercial_production):
    included = [
        "Dedicated Account Management Team",
        "Monthly Reporting Calls & Optimizations",
        "Dashboard Access",
        "Web Attribution (Pixel Required)",
    ]
    if commercial_production:
        included.append("Commercial Production")
    if targeting.get("sales_attribution"):
        included.append("Sales Attribution (CRM Upload Required)")
    if targeting.get("brand_lift"):
        included.append("Brand Lift Study")
    return included


# ---------------- Case study vault ----------------
# Coarse product tags a case study gets labelled with. The real PRODUCTS keys
# plus the two umbrella selections that aren't single products, so a tag can
# describe "this demonstrates Live Sports" without naming a package.
CASE_STUDY_PRODUCT_TAGS = list(FALLBACK_PRODUCTS) + ["live_sports", "total_tv"]

# The 3 most recent matching case studies are pre-checked; anything beyond
# that is opt-in. Enough to be useful, few enough that nobody ships a deck
# with nine case studies bolted on by accident.
CASE_STUDY_PRECHECK = 3


def extract_case_study_text(path, per_slide_cap=2500):
    """All of a case study's slide text, one entry per slide."""
    prs = Presentation(path)
    return [slide_map.extract_slide_text(slide)[:per_slide_cap] for slide in prs.slides]


def build_case_study_prompt(slide_texts, filename):
    return f"""You are cataloguing a Premion CTV/OTT advertising case study so sellers can find it later. Return ONLY valid JSON matching this schema -- no markdown fences, no preamble:

{{"title": "", "verticals": [], "products": [], "summary": ""}}

- "title": a short human title for this case study, in the deck's own words where possible (e.g. "Regional Bank -- Precision Banking Audience Targeting"). Do not include the word "case study".
- "verticals": every vertical this case study is genuinely relevant to, each exactly one of {list(VERTICALS.values())} (never "none"). Usually one, occasionally two when the client plausibly belongs to both. Do not stretch -- an empty list is better than a wrong tag, and a seller can add one by hand.
- "products": the Premion products this campaign actually used, each exactly one of {CASE_STUDY_PRODUCT_TAGS}. Base this on what the slides say was bought, not on what could have been.
- "summary": ONE sentence a seller can skim, naming the client type and the headline result (e.g. "Regional bank drove a 20% year-over-year lift in consumer lending with precision audience targeting and OTT retargeting."). No more than about 25 words.

Source file name: {filename}

Slide text, one entry per slide (JSON): {json.dumps(slide_texts)}
"""


def call_claude_case_study_tags(slide_texts, filename):
    return _call_claude_json(build_case_study_prompt(slide_texts, filename))


def _valid_tags(values, allowed):
    """Keep only tags that really exist, preserving order. Claude is told the
    exact lists, but a tag that doesn't match would silently never match a
    proposal either -- better dropped here than mysteriously inert later."""
    seen = []
    for value in values or []:
        if value in allowed and value not in seen:
            seen.append(value)
    return seen


def render_add_case_study():
    """Upload a case study .pptx, have Claude propose its tags, edit them,
    and store it. Anyone using the app can add one -- there are no accounts,
    hence the free-text "added by"."""
    st.header("Add case study")
    st.caption("Got a case study deck worth reusing? Upload it here and it becomes available "
               "to everyone building proposals. Claude reads the slides and proposes a title, "
               "the industries and products it speaks to, and a one-line summary — all "
               "editable before you save.")

    upload = st.file_uploader("Case study (.pptx)", type=["pptx"], key="cs_upload")
    if not upload:
        st.session_state.pop("cs_suggestion", None)
        return

    # Written to disk because both the text extraction and (on confirm) the
    # optimizer need a real file to open.
    local_path = db.scratch_dir("premion_case_study_uploads") / upload.name
    local_path.write_bytes(upload.getvalue())

    try:
        slide_texts = extract_case_study_text(str(local_path))
    except Exception as exc:
        st.error(f"Couldn't read that .pptx: {exc}")
        return
    st.caption(f"{len(slide_texts)} slide(s), "
               f"{local_path.stat().st_size / 1024 / 1024:.1f} MiB")

    if st.session_state.get("cs_suggestion_for") != upload.name:
        if st.button("Read the deck and suggest tags", type="primary"):
            with st.spinner("Reading the case study..."):
                suggestion, error = call_claude_case_study_tags(slide_texts, upload.name)
            if error:
                st.error(error)
            else:
                st.session_state["cs_suggestion"] = suggestion
                st.session_state["cs_suggestion_for"] = upload.name
                st.rerun()
        st.caption("Or fill the fields in yourself below.")

    suggestion = st.session_state.get("cs_suggestion", {}) or {}
    if suggestion:
        st.success("Claude's suggestions are filled in below -- correct anything before saving.")

    title = st.text_input("Title", value=suggestion.get("title", "") or Path(upload.name).stem)
    summary = st.text_area("One-line summary", value=suggestion.get("summary", ""), height=70)

    vertical_labels = {v: k for k, v in VERTICALS.items() if v != "none"}
    verticals = st.multiselect(
        "Verticals", list(vertical_labels), format_func=lambda v: vertical_labels[v],
        default=_valid_tags(suggestion.get("verticals"), vertical_labels),
        help="Which verticals this case study should be offered for.")
    products = st.multiselect(
        "Product tags", CASE_STUDY_PRODUCT_TAGS,
        default=_valid_tags(suggestion.get("products"), CASE_STUDY_PRODUCT_TAGS))
    # Prefilled from who's signed in, but still editable -- someone uploading
    # a colleague's case study should be able to say so.
    added_by = st.text_input("Added by",
                             value=st.session_state.get("cs_added_by") or current_user() or "",
                             placeholder="Your name")

    if not verticals:
        st.warning("With no vertical tagged this case study will only ever appear in the "
                   "off-vertical list, never pre-selected for anyone.")

    if st.button("Save to the vault", disabled=not title.strip()):
        if not added_by.strip():
            st.warning("Add your name first so the vault records who contributed this.")
            return
        st.session_state["cs_added_by"] = added_by
        with st.spinner("Optimizing and uploading..."):
            row, stats, error = db.upload_case_study(
                str(local_path), upload.name, title.strip(), verticals, products,
                summary.strip(), added_by.strip())
        if stats:
            st.caption(f"Optimized {stats['source_size'] / 1024 / 1024:.1f} MiB -> "
                       f"{stats['dst_size'] / 1024 / 1024:.1f} MiB before storing.")
        if error:
            st.error(error)
        else:
            st.success(f"Saved \"{row['title']}\" to the vault. "
                       f"It's now offered on any proposal tagged "
                       f"{', '.join(verticals) if verticals else 'no vertical'}.")
            for key in ("cs_suggestion", "cs_suggestion_for"):
                st.session_state.pop(key, None)


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

# Supabase's free tier storage allowance. Only used to express attached-file
# usage as a share of it on the History page -- everything else in this app
# is regenerable from a stored recipe and doesn't grow, so attached finals
# are the one thing worth watching.
FREE_TIER_STORAGE_BYTES = 1024 * 1024 * 1024


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


def render_case_study_download(row, key_prefix):
    """Download button for one case study, fetched on demand.

    Deliberately not eager. st.download_button needs the file's bytes at
    render time, so drawing one per row would pull every visible case study
    out of storage (~42MiB across the vault) whether or not anyone clicked.
    Instead: a file already in the local cache -- from an earlier click, or
    because a proposal included it -- gets a real download button straight
    away, and anything else takes one click to fetch first. Either way
    db.case_study_file caches it, so it's never fetched from Supabase twice.
    """
    slot = f"{key_prefix}_{row['id']}"
    path = db.case_study_cached_path(row["id"], row["storage_path"])

    if path is None:
        if not st.button("Get .pptx", key=f"{slot}_fetch",
                         help="Fetches the deck from the vault, then offers it as a download."):
            return
        try:
            path = db.case_study_file(row["id"], row["storage_path"])
        except Exception as exc:
            st.warning(f"Couldn't fetch that case study ({db.describe_error(exc)}).")
            return

    with open(path, "rb") as handle:
        st.download_button("⬇ Download .pptx", data=handle.read(),
                           file_name=case_study_filename(row), mime=PPTX_MIME,
                           key=f"{slot}_download")


def render_case_study_finder():
    """Standalone vault page: browse and edit every case study, or describe a
    client and let Claude recommend from it. Reps use this outside the
    proposal flow, so it stands on its own rather than living in an
    expander."""
    st.header("Case study finder")
    st.caption("Past Premion campaigns you can show a client as proof. Browse or search them "
               "here, or describe a client and let Claude suggest the most relevant ones. "
               "To put them *in* a proposal, use the picker just above Generate on the Build "
               "page — it pre-selects the ones matching your client's industry.")
    rows, warning = db.fetch_case_studies(active_only=False)
    if warning:
        st.warning(warning)
        return
    if not rows:
        st.info("The vault is empty. Add one from the \"Add case study\" page.")
        return

    browse_tab, suggest_tab = st.tabs(["Browse", "Suggest"])

    with browse_tab:
        _render_vault_browser(rows)
    with suggest_tab:
        _render_case_study_suggest(rows)


def _render_vault_browser(rows):
    vertical_labels = {v: k for k, v in VERTICALS.items() if v != "none"}
    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        filter_verticals = st.multiselect("Vertical", list(vertical_labels),
                                          format_func=lambda v: vertical_labels[v],
                                          key="vault_filter_verticals")
    with col2:
        filter_products = st.multiselect("Product", CASE_STUDY_PRODUCT_TAGS,
                                         key="vault_filter_products")
    with col3:
        query = st.text_input("Search title or summary", key="vault_search").strip().lower()
    show_inactive = st.checkbox("Include deactivated", key="vault_show_inactive")

    def matches(row):
        if not show_inactive and not row.get("active", True):
            return False
        if filter_verticals and not set(filter_verticals) & set(row.get("verticals") or []):
            return False
        if filter_products and not set(filter_products) & set(row.get("products") or []):
            return False
        if query and query not in f"{row.get('title') or ''} {row.get('summary') or ''}".lower():
            return False
        return True

    shown = [r for r in rows if matches(r)]
    st.caption(f"{len(shown)} of {len(rows)} case studies")

    for row in shown:
        state = "" if row.get("active", True) else "  ·  deactivated"
        with st.expander(f"{row['title'] or row['filename']}{state}", expanded=False):
            st.caption(f"{row.get('summary') or '_no summary_'}")
            st.caption(f"added by {row.get('added_by') or 'unknown'} · "
                       f"{str(row.get('date_added'))[:10]} · `{row['filename']}`")

            key = f"vault_{row['id']}"
            title = st.text_input("Title", value=row["title"] or "", key=f"{key}_title")
            summary = st.text_area("Summary", value=row.get("summary") or "", height=70,
                                   key=f"{key}_summary")
            verticals = st.multiselect(
                "Verticals", list(vertical_labels), format_func=lambda v: vertical_labels[v],
                default=_valid_tags(row.get("verticals"), vertical_labels), key=f"{key}_verticals")
            products = st.multiselect(
                "Products", CASE_STUDY_PRODUCT_TAGS,
                default=_valid_tags(row.get("products"), CASE_STUDY_PRODUCT_TAGS),
                key=f"{key}_products")

            save_col, active_col, download_col = st.columns([1, 1, 1])
            with download_col:
                render_case_study_download(row, "browse")
            with save_col:
                if st.button("Save changes", key=f"{key}_save"):
                    _, error = db.update_case_study(
                        row["id"], title=title.strip(), summary=summary.strip(),
                        verticals=verticals, products=products)
                    if error:
                        st.error(error)
                    else:
                        st.success("Saved.")
                        st.rerun()
            with active_col:
                active_now = row.get("active", True)
                label = "Deactivate" if active_now else "Reactivate"
                if st.button(label, key=f"{key}_active"):
                    _, error = db.update_case_study(row["id"], active=not active_now)
                    if error:
                        st.error(error)
                    else:
                        st.rerun()
            st.caption("Deactivated case studies stay in the vault and keep their file — "
                       "they just stop being offered on proposals.")


def _render_case_study_suggest(rows):
    st.caption("Describe the client or campaign and Claude picks from the vault, "
               "with a one-line reason for each.")
    active = [r for r in rows if r.get("active", True)]
    description = st.text_area(
        "Client / campaign", height=110, key="cs_suggest_description",
        placeholder="Regional HVAC company, wants to drive service calls in shoulder season, "
                    "competing against national franchises")
    if st.button("Suggest case studies", type="primary"):
        if not description.strip():
            st.warning("Describe the client first.")
        elif not active:
            st.warning("No active case studies to choose from.")
        else:
            with st.spinner("Reading the vault..."):
                result, error = _call_claude_json(
                    build_case_study_suggest_prompt(description, active))
            if error:
                st.error(error)
            else:
                st.session_state["cs_suggestions"] = result.get("recommendations", [])

    suggestions = st.session_state.get("cs_suggestions")
    if suggestions is None:
        return
    by_id = {r["id"]: r for r in active}
    # Claude is given the exact ids, but an invented one would otherwise show
    # as a blank row -- same validation bar as audience segments.
    valid = [s for s in suggestions if s.get("id") in by_id]
    if not valid:
        st.info("Nothing in the vault fits that description well enough to recommend.")
        return
    st.success(f"{len(valid)} case study(ies) recommended:")
    for suggestion in valid:
        row = by_id[suggestion["id"]]
        st.markdown(f"**{row['title']}** — {suggestion.get('reason', '')}")
        st.caption(f"{row.get('summary') or ''}  ·  "
                   f"{', '.join(row.get('verticals') or []) or 'no vertical tags'}")
        render_case_study_download(row, "suggest")
        st.divider()


def render_case_study_picker(vertical_key, vertical_label):
    """The generate-time checklist. Returns the selected rows, in order.

    Nothing is ever included silently: matching case studies are pre-checked
    but always visible and always overridable, and off-vertical ones are one
    expander away rather than hidden.
    """
    st.header("Case studies")
    rows, warning = db.fetch_case_studies()
    if warning:
        st.warning(f"{warning}. No case studies can be added to this deck.")
        return []
    if not rows:
        st.caption("The vault is empty -- add one from the \"Add case study\" page.")
        return []

    matching = [r for r in rows if vertical_key in (r.get("verticals") or [])]
    matching_ids = {r["id"] for r in matching}
    others = [r for r in rows if r["id"] not in matching_ids]

    # A keyed checkbox's session_state entry beats its value= argument, which
    # is exactly what we want for "pre-checked but overridable" -- the default
    # applies once and the seller's own choice sticks. But when the vertical
    # changes the pre-check set changes with it, so those keys have to be
    # dropped or they'd carry the previous vertical's answer.
    if st.session_state.get("cs_defaults_for") != vertical_key:
        for key in [k for k in st.session_state if k.startswith("cs_pick_")]:
            del st.session_state[key]
        st.session_state["cs_defaults_for"] = vertical_key

    selected = []

    def _row(case_study, default):
        label = case_study["title"] or case_study["filename"]
        picked = st.checkbox(label, value=default, key=f"cs_pick_{case_study['id']}")
        meta = [case_study["summary"] or ""]
        if case_study.get("added_by"):
            meta.append(f"added by {case_study['added_by']}")
        if case_study.get("date_added"):
            meta.append(str(case_study["date_added"])[:10])
        st.caption(" · ".join(m for m in meta if m))
        if picked:
            selected.append(case_study)

    if matching:
        st.caption(f"{len(matching)} case study(ies) tagged {vertical_label} -- "
                   f"the {min(CASE_STUDY_PRECHECK, len(matching))} most recent are pre-selected.")
        for i, case_study in enumerate(matching):
            _row(case_study, default=i < CASE_STUDY_PRECHECK)
    elif vertical_key != "none":
        st.caption(f"No case studies are tagged {vertical_label} yet.")

    with st.expander(f"Other case studies ({len(others)}) -- not tagged for this vertical", expanded=False):
        for case_study in others:
            _row(case_study, default=False)

    if selected:
        st.caption(f"{len(selected)} case study(ies) will be added at the end of the "
                   f"content, immediately before the media plan.")
    return selected


# ---------------- Master deck update ----------------
def _diff_line(entry):
    key = entry.get("key") or "?"
    return f"**{entry['number']}.** {entry['label']} — `{key}`"


def _proposal_summary(row):
    """The one-line facts the History list shows for a proposal."""
    form = row.get("form_json") or {}
    options = form.get("plan_options") or []
    total = 0.0
    for option in options:
        totals = option.get("totals") or {}
        total += float(totals.get("full_flight_cost") or 0)
    flight = (form.get("flight") or {}).get("label") or "--"
    return {
        "options": len(options),
        "budget": total,
        "flight": flight,
        "title": form.get("proposal_title") or "",
    }


def _render_proposal_row(row, siblings, index):
    """One proposal in the History list, with its actions."""
    rid = row["id"]
    summary = _proposal_summary(row)
    generated = str(row.get("generated_at") or "")[:16].replace("T", " ")
    vertical_label = next((label for label, key in VERTICALS.items()
                           if key == (row.get("vertical") or "none")), row.get("vertical") or "--")
    attached = bool(row.get("file_storage_path"))

    badges = []
    if len(siblings) > 1:
        badges.append(f"#{index + 1} of {len(siblings)} for this client")
    if row.get("parent_proposal_id"):
        badges.append("revision")
    if attached:
        badges.append("final file attached")
    if row.get("revision_label"):
        badges.append(row["revision_label"])
    if not row.get("logo_storage_path") and (row.get("form_json") or {}).get("logo_used"):
        # Distinguished from "had no logo": only this case rebuilds wrong.
        badges.append("logo not stored")

    header = (f"{generated}  ·  {summary['title'] or 'Untitled'}  ·  "
              f"{summary['options']} option(s)  ·  ${summary['budget']:,.0f}")
    if row.get("created_by"):
        header += f"  ·  {row['created_by']}"
    if badges:
        header += "   [" + " | ".join(badges) + "]"

    with st.expander(header, expanded=False):
        meta = st.columns(4)
        meta[0].caption(f"**Vertical**\n\n{vertical_label}")
        meta[1].caption(f"**Market**\n\n{row.get('market') or '--'}")
        meta[2].caption(f"**Flight**\n\n{summary['flight']}")
        meta[3].caption(f"**Deck version**\n\n"
                        f"{row.get('deck_version_id') if row.get('deck_version_id') else 'local fallback'}")

        # Editable after the fact: a label is an annotation about the row,
        # not a claim about what was generated, so it's the one field history
        # lets you change. Nothing else here is patchable.
        lcol1, lcol2 = st.columns([3, 1])
        typed_label = lcol1.text_input(
            "Revision label", value=row.get("revision_label") or "",
            key=f"hist_label_{rid}", placeholder="e.g. Revision 2, or 'the one they picked'",
            label_visibility="collapsed")
        if lcol2.button("Save label", key=f"hist_savelabel_{rid}",
                        disabled=(typed_label.strip() or None) == (row.get("revision_label") or None)):
            _, error = db.update_proposal(rid, revision_label=typed_label.strip() or None)
            if error:
                st.error(error)
            else:
                st.session_state["history_flash"] = ["Label updated."]
                st.rerun()

        if attached:
            st.info(f"A hand-edited final file is attached"
                    + (f" — {row['file_note']}" if row.get("file_note") else "")
                    + f" (attached {str(row.get('file_attached_at') or '')[:16].replace('T', ' ')}). "
                      f"This is what the client actually received; the stored recipe below "
                      f"describes the deck *before* those edits.")

        actions = st.columns(4)

        # The next build in this client's thread. Counts what's already
        # there, so loading the oldest of three still proposes "Revision 4"
        # rather than colliding with an existing label.
        next_revision = f"Revision {len(siblings) + 1}"

        if actions[0].button("Load into form", key=f"hist_load_{rid}",
                             help="Fills the Build page with everything from this proposal so "
                                  "you can change it and generate again. Uses the CURRENT "
                                  "template deck, so you get any slide updates since. Saves as "
                                  "a new revision — the original is never overwritten."):
            notes = rehydrate_proposal_into_form(row, parent_proposal_id=rid,
                                                 revision_default=next_revision)
            st.session_state["history_flash"] = (
                [f"Loaded into the form, labelled \"{next_revision}\". Review it, then Generate — "
                 f"that will log a new revision linked back to this one."] + notes)
            st.session_state["history_goto_build"] = True
            st.rerun()

        if actions[1].button("Start new from this", key=f"hist_new_{rid}",
                             help="Same as Load into form, but for a genuinely new proposal "
                                  "rather than a revision of this one — use it when an old "
                                  "proposal is a convenient starting point."):
            notes = rehydrate_proposal_into_form(row, parent_proposal_id=rid,
                                                 revision_default=next_revision)
            st.session_state["history_flash"] = (
                ["Started a new proposal from this one. It will log as a new entry linked "
                 "back to the original."] + notes)
            st.session_state["history_goto_build"] = True
            st.rerun()

        if actions[2].button("Rebuild as presented", key=f"hist_rebuild_{rid}",
                             help="Re-download this exact deck if you've lost the file. Uses "
                                  "the ORIGINAL template deck and the original numbers, so you "
                                  "get what the client saw — not what it would look like if "
                                  "you built it today."):
            st.session_state[f"hist_rebuilt_{rid}"] = True

        with actions[3]:
            if attached:
                if st.button("Download final", key=f"hist_dl_{rid}"):
                    st.session_state[f"hist_fetch_file_{rid}"] = True

        if st.session_state.get(f"hist_fetch_file_{rid}"):
            try:
                path = db.proposal_file(rid, row["file_storage_path"])
                with open(path, "rb") as handle:
                    st.download_button("⬇ Download the attached final .pptx", data=handle.read(),
                                       file_name=Path(row["file_storage_path"]).name,
                                       mime=PPTX_MIME, key=f"hist_dlbtn_{rid}")
            except Exception as exc:
                st.error(f"Couldn't fetch the attached file: {db.describe_error(exc)}")

        if st.session_state.get(f"hist_rebuilt_{rid}"):
            if attached:
                st.warning("This proposal has a hand-edited final file attached. A rebuild "
                           "reproduces the deck as it was *before* those edits, so the two will "
                           "differ — download the attached final if you want what the client saw.")
            with st.spinner("Rebuilding..."):
                buffer, filename, warnings = rebuild_proposal_deck(row)
            for message in warnings:
                st.warning(message)
            if buffer is None:
                st.error("Rebuild failed — see above.")
            else:
                st.download_button("⬇ Download the rebuilt .pptx", data=buffer,
                                   file_name=filename, mime=PPTX_MIME, key=f"hist_rbtn_{rid}")

        # --- attach a final file ------------------------------------------
        with st.popover("Attach final file"):
            st.caption("Sent the client a hand-edited version? Attach it here so the record "
                       "matches what they actually received. Without this, the saved proposal "
                       "describes the deck as the app built it — before whatever you changed "
                       "in PowerPoint.")
            upload = st.file_uploader(".pptx", type=["pptx"], key=f"hist_up_{rid}")
            note = st.text_input("Note", placeholder="final as sent 8/5 — trimmed to 40 slides",
                                 key=f"hist_note_{rid}")
            if upload is not None and st.button("Attach", key=f"hist_attach_{rid}"):
                target = db.scratch_dir("premion_proposal_finals") / upload.name
                target.write_bytes(upload.getvalue())
                # Who attached it, so a note like "trimmed to 40 slides" has
                # someone to ask about it later.
                stamped = " — ".join(filter(None, [note.strip() or None,
                                                   f"attached by {current_user()}"
                                                   if current_user() else None]))
                with st.spinner("Optimizing and uploading..."):
                    updated, stats, error = db.attach_proposal_file(
                        rid, str(target), upload.name, stamped or None)
                if error:
                    st.error(error)
                else:
                    if stats:
                        st.caption(f"Optimized {stats}")
                    st.success("Attached.")
                    st.rerun()

        # --- delete --------------------------------------------------------
        with st.popover("Delete"):
            st.caption("Hard-deletes this row and any attached file. Revisions descending from "
                       "it are kept — their link back is just cleared. There's no bulk delete: "
                       "this is for junk and test rows.")
            typed = st.text_input(f"Type DELETE to confirm", key=f"hist_del_confirm_{rid}")
            if st.button("Delete permanently", key=f"hist_del_{rid}",
                         disabled=typed.strip().upper() != "DELETE"):
                ok, error = db.delete_proposal(rid)
                if ok:
                    st.session_state["history_flash"] = [error] if error else ["Proposal deleted."]
                    st.rerun()
                else:
                    st.error(error)


def render_proposal_history():
    """The Proposal History page: browse, reuse, rebuild, attach, delete."""
    st.header("Proposal history")
    st.caption("Every proposal you generate is saved here automatically. Load one to revise it, "
               "rebuild exactly what was presented to a client, or start a new proposal from an "
               "old one as a baseline.")

    for message in st.session_state.pop("history_flash", []) or []:
        st.info(message)

    rows, warning = db.fetch_proposals()
    if warning:
        st.warning(warning)
    if rows is None:
        return
    if not rows:
        st.info("No proposals logged yet. Generate one from the Build page and it'll appear here.")
        return

    # Storage awareness: attached files are the only thing here that isn't
    # regenerable from a recipe, so they're the only thing that grows without
    # bound. Surfaced before the free tier fills, not after.
    total, count, usage_warning = db.bucket_usage()
    if not usage_warning and count:
        share = total / FREE_TIER_STORAGE_BYTES
        line = (f"Attached final files: {count} file(s), {total / 1048576:.0f} MiB "
                f"({share:.0%} of the 1GB free tier).")
        (st.warning if share > 0.75 else st.caption)(line)

    fcol1, fcol2 = st.columns([2, 1])
    search = fcol1.text_input("Search by client", key="history_search").strip().lower()
    verticals_present = sorted({r.get("vertical") for r in rows if r.get("vertical")})
    vertical_labels = ["All"] + [next((label for label, key in VERTICALS.items() if key == v), v)
                                 for v in verticals_present]
    picked = fcol2.selectbox("Vertical", vertical_labels, key="history_vertical")

    filtered = rows
    if search:
        filtered = [r for r in filtered if search in (r.get("client_name") or "").lower()]
    if picked != "All":
        want = VERTICALS.get(picked, picked)
        filtered = [r for r in filtered if r.get("vertical") == want]

    if not filtered:
        st.info("Nothing matches that filter.")
        return

    # Grouped by client so a client's history reads as a thread rather than
    # as unrelated rows scattered through one long list.
    by_client = {}
    for row in filtered:
        by_client.setdefault(row.get("client_name") or "(no client name)", []).append(row)

    st.caption(f"{len(filtered)} proposal(s) across {len(by_client)} client(s)")
    for client_name, client_rows in by_client.items():
        suffix = f" — {len(client_rows)} proposals" if len(client_rows) > 1 else ""
        st.subheader(f"{client_name}{suffix}")
        for index, row in enumerate(client_rows):
            _render_proposal_row(row, client_rows, index)


def render_update_master_deck():
    """Upload a new master deck, see what changed against the active version,
    and activate it -- but only if every slide resolves to a condition_key."""
    st.header("Update master deck")
    st.caption("For when the master template deck itself changes — new slides, new branding, "
               "updated stats. Upload it to see exactly what changed against the version in "
               "use. Nothing goes live until you activate it, and every previous version is "
               "kept, so this is reversible.")
    st.info("**Why an upload can be blocked:** the app decides which slides go in a proposal "
            "by reading a label in each slide's speaker notes. If a new or edited slide has "
            "no label, the app would have to guess where it belongs — so activation is "
            "blocked and the slides in question are listed by number. Add a `key:` line to "
            "each one's speaker notes in PowerPoint and re-upload.", icon="💡")

    active, warning = db.active_deck_version()
    if warning:
        st.warning(f"{warning}. You can still scan a deck below, but the comparison will "
                   f"treat every slide as new.")
    else:
        st.info(f"**In use:** version {active['id']} — {active['filename']} "
                f"(uploaded {str(active['uploaded_at'])[:10]})"
                + (f"\n\n{active['notes']}" if active.get("notes") else ""))

        # The whole edit loop is download → change it in PowerPoint → upload
        # again, so the download belongs on this page rather than sending
        # someone to the Supabase dashboard for the first step.
        st.caption("**To make a change:** download the active deck, edit it in PowerPoint "
                   "(remember to set `key:` in the speaker notes of any slide you add — "
                   "see `SLIDE_KEYS.md`), then upload it below.")
        try:
            deck_path = db.deck_file(active)
        except Exception as exc:
            st.warning(f"Couldn't fetch the active deck to download ({db.describe_error(exc)}).")
        else:
            with open(deck_path, "rb") as handle:
                st.download_button(
                    f"⬇ Download active master deck ({Path(deck_path).stat().st_size / 1024 / 1024:.0f} MiB)",
                    data=handle.read(), file_name=active["filename"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

    upload = st.file_uploader("New master deck (.pptx)", type=["pptx"], key="deck_upload")
    if not upload:
        return

    local_path = db.scratch_dir("premion_deck_uploads") / upload.name
    local_path.write_bytes(upload.getvalue())

    with st.spinner("Scanning the deck..."):
        try:
            new_prs = Presentation(str(local_path))
        except Exception as exc:
            st.error(f"Couldn't open that .pptx: {exc}")
            return
        old_prs = None
        if active:
            try:
                old_prs = Presentation(db.deck_file(active))
            except Exception as exc:
                st.warning(f"Couldn't fetch the active version to compare against "
                           f"({db.describe_error(exc)}). Showing the new deck's own scan only.")
        diff = slide_map.diff_decks(old_prs, new_prs)

    counts = st.columns(5)
    for column, (label, value) in zip(counts, [
        ("Slides", diff["new_count"]), ("Added", len(diff["added"])),
        ("Removed", len(diff["removed"])), ("Moved", len(diff["moved"])),
        ("Re-tagged", len(diff["retagged"])),
    ]):
        column.metric(label, value)

    # ---- the hard gate -------------------------------------------------
    if diff["unresolved"]:
        st.error(
            f"**Cannot activate: {len(diff['unresolved'])} slide(s) don't resolve to a key.**\n\n"
            f"Every slide needs a key so the builder knows when to include it. Add a "
            f"`key: <something>` line to each slide's speaker notes in PowerPoint "
            f"(View → Notes Page), save, and upload again. See `SLIDE_KEYS.md` for the "
            f"list of valid keys."
        )
        st.markdown("**These slides need a `key:` tag:**")
        for entry in diff["unresolved"]:
            st.markdown(f"- **Slide {entry['number']}** — {entry['label']}")
        st.stop()

    st.success(f"All {diff['new_count']} slides resolve to a key.")

    if diff["added"]:
        with st.expander(f"Added ({len(diff['added'])})", expanded=True):
            for entry in diff["added"]:
                st.markdown(f"- {_diff_line(entry)}")
    if diff["removed"]:
        with st.expander(f"Removed ({len(diff['removed'])})", expanded=True):
            for entry in diff["removed"]:
                st.markdown(f"- {_diff_line(entry)}")
    if diff["retagged"]:
        with st.expander(f"Re-tagged ({len(diff['retagged'])})", expanded=True):
            st.caption("Same slide, different key -- it will now be included under "
                       "different circumstances.")
            for entry in diff["retagged"]:
                st.markdown(f"- **{entry['number']}.** {entry['label']} — "
                            f"`{entry['was_key']}` → `{entry['key']}`")
    if diff["moved"]:
        with st.expander(f"Moved ({len(diff['moved'])})", expanded=False):
            st.caption("Position changed only -- harmless, since nothing depends on "
                       "slide numbering.")
            for entry in diff["moved"]:
                st.markdown(f"- **{entry['was']} → {entry['number']}.** {entry['label']}")
    st.caption(f"{diff['unchanged']} slide(s) unchanged.")

    notes = st.text_input("Notes for this version",
                          placeholder="e.g. Added PGA Majors package slide, refreshed Q3 stats")
    if st.button("Activate this deck", type="primary"):
        with st.spinner("Optimizing, uploading and activating..."):
            stamp = datetime.now().strftime("%Y%m%dT%H%M%SZ")
            row, stats, error = db.upload_deck(
                str(local_path), f"masters/{stamp}_{upload.name}",
                notes=notes.strip() or None, activate=True)
        if stats:
            st.caption(f"Optimized {stats['source_size'] / 1024 / 1024:.2f} MiB → "
                       f"{stats['dst_size'] / 1024 / 1024:.2f} MiB before storing.")
        if error:
            st.error(error)
        else:
            st.success(f"Version {row['id']} is now the active master deck. "
                       f"Previous versions are still in storage.")
            st.balloons()


def main():
    if not check_password():
        return
    if not check_identity():
        return

    # The finders are useful on their own -- a rep looking up an audience
    # segment or a case study isn't necessarily building a proposal today --
    # so each has a page as well as its embedded place in the proposal flow.
    # Both entry points call the same component; nothing is duplicated.
    st.sidebar.title("Premion")
    # "Load into form" on the History page jumps here by pre-selecting the
    # Build page, which works because the radio is keyed.
    if st.session_state.pop("history_goto_build", False):
        st.session_state["page_choice"] = "Build a proposal"
    page = st.sidebar.radio("Page", [
        "Build a proposal",
        "Proposal history",
        "Audience finder",
        "Case study finder",
        "Add case study",
        "Update master deck",
    ], label_visibility="collapsed", key="page_choice")
    st.sidebar.caption("The finders are also embedded in the proposal flow — "
                       "audiences in Section D2, case studies just before Generate.")
    st.sidebar.divider()
    render_identity_sidebar()

    standalone = {
        "Proposal history": render_proposal_history,
        "Audience finder": render_audience_finder_page,
        "Case study finder": render_case_study_finder,
        "Add case study": render_add_case_study,
        "Update master deck": render_update_master_deck,
    }
    if page in standalone:
        standalone[page]()
        return

    st.title("Premion Proposal Builder")
    st.caption("Fill in the client and campaign details below, then Generate to get a "
               "personalized PowerPoint deck. Nothing is sent anywhere — you download the "
               "file and send it yourself.")

    with st.expander("ℹ️ How this app works", expanded=False):
        st.markdown(
            "**1. Start with your notes (optional).** Paste your meeting or discovery notes "
            "into *Draft from notes* below and Claude fills in most of this form for you — "
            "client details, budget, products, audiences, dates.\n\n"
            "**2. Review anything it flagged.** A yellow box lists what was ambiguous or "
            "assumed. These are questions, not errors — answer them in *Clarify and re-draft* "
            "and it'll revise, once.\n\n"
            "**3. Adjust anything you like.** Everything is editable whether it was drafted or "
            "not. Your edits are kept — a re-draft won't overwrite a section you've touched.\n\n"
            "**4. Check the media plan.** Type either impressions or cost and the other side "
            "calculates itself. You can show up to three options side by side.\n\n"
            "**5. Generate.** You get a .pptx to download, built only from the slides your "
            "selections call for.\n\n"
            "**6. Find it again in Proposal history.** Every deck you generate is saved there. "
            "Load one back to revise it, or rebuild exactly what a client was sent.\n\n"
            "*No notes to work from?* Skip step 1 and fill the form in by hand — the notes "
            "panel is a shortcut, not a requirement."
        )

    # Every Supabase-backed loader hands back a warning instead of raising, so
    # the form always renders -- but a fallback is never silent. The master
    # deck's own warning is raised at generate time, since that's when it's
    # actually fetched.
    for warning in (RATE_CARD_WARNING, AUDIENCE_CATALOG_WARNING):
        if warning:
            st.warning(warning)

    # ---------------- Draft from notes (Claude) ----------------
    with st.expander("📝 Draft from notes (optional)", expanded=False):
        st.caption("Paste your meeting or discovery notes here — however rough — and Claude "
                   "fills in most of the form below: client details, budget and media plan, "
                   "products, audiences, flight dates. Anything it wasn't sure about is listed "
                   "for you to confirm rather than guessed at silently. **Everything it fills "
                   "in is editable, and nothing is final until you press Generate.** Whatever "
                   "you paste is saved with the proposal either way, so the next person can see "
                   "where the numbers came from.")
        notes_input = st.text_area("Meeting / discovery notes", height=180, key="draft_notes_input")
        if st.button("Draft proposal from notes"):
            if not notes_input.strip():
                st.warning("Paste some notes first.")
            else:
                with st.spinner("Drafting from notes..."):
                    draft, error = call_claude_draft(notes_input)
                if error:
                    st.error(error)
                else:
                    try:
                        apply_draft_to_form(draft)
                    except Exception as exc:
                        st.error(f"Couldn't apply the draft to the form: {exc}")
                    else:
                        # Remembered so the clarification round can send the
                        # model its own previous answer to revise, and so it
                        # knows which sections the draft originally owned.
                        st.session_state["draft_source_notes"] = notes_input
                        st.session_state["draft_last_json"] = draft
                        st.session_state["draft_round"] = 1
                        st.session_state["draft_sections_round1"] = set(
                            st.session_state.get("ai_filled_sections", set()))
                        st.rerun()

    if st.session_state.get("draft_unresolved"):
        st.warning("**Review before generating:**\n\n" +
                   "\n".join(f"- {item}" for item in st.session_state["draft_unresolved"]))

    # One clarification round: answer the open questions in plain language and
    # the model revises its own draft rather than starting over. Offered once
    # -- after the re-draft the box doesn't come back.
    if st.session_state.get("draft_round") == 1 and st.session_state.get("draft_last_json"):
        with st.container(border=True):
            st.markdown("**Clarify and re-draft**")
            st.caption("Answer the open questions above in plain language "
                       "(e.g. \"budget includes the fee, flight is 2026, use FIN General In Mkt Shopper\") "
                       "and Claude will revise the draft. Anything you've edited by hand since the first "
                       "draft is kept as you left it. One round.")
            clarifications = st.text_area("Your answers", height=110, key="draft_clarifications")
            if st.button("Re-draft with these answers"):
                if not clarifications.strip():
                    st.warning("Answer at least one of the open questions first.")
                else:
                    # Sections the first draft filled but that have since been
                    # cleared by an on_change -- i.e. the user has edited them
                    # by hand, so the re-draft must not overwrite them.
                    edited_since = (st.session_state.get("draft_sections_round1", set())
                                    - st.session_state.get("ai_filled_sections", set()))
                    with st.spinner("Re-drafting with your clarifications..."):
                        draft, error = call_claude_redraft(
                            st.session_state.get("draft_source_notes", ""),
                            st.session_state["draft_last_json"],
                            clarifications,
                        )
                    if error:
                        st.error(error)
                    else:
                        try:
                            apply_draft_to_form(draft, skip_sections=edited_since)
                        except Exception as exc:
                            st.error(f"Couldn't apply the re-draft to the form: {exc}")
                        else:
                            st.session_state["draft_last_json"] = draft
                            st.session_state["draft_round"] = 2
                            st.rerun()

    # ---------------- Section A: Client basics ----------------
    st.header("A. Client basics")
    ai_section_badge("basics")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client name", value="Acme Test Co", key="client_name",
                                     on_change=_clear_ai_section, args=("basics",))
        market_choice = st.radio("Market", ["DC", "Harrisburg"], horizontal=True, key="market_choice",
                                  on_change=_clear_ai_section, args=("basics",))
        vertical_choice = st.selectbox("Vertical", list(VERTICALS.keys()), index=0, key="vertical_choice",
                                        on_change=_clear_ai_section, args=("basics",))
        agency_involved = st.toggle("Ad agency involved? (gross markup x1.15)", value=False, key="agency_involved",
                                     on_change=_clear_ai_section, args=("basics",))
    with col2:
        logo_file = st.file_uploader("Client logo", type=["png", "jpg", "jpeg"])
        # A file_uploader can't be prefilled from session_state, so a
        # proposal loaded from History carries its stored logo as a path
        # instead: it's used unless a new file is uploaded over it, which is
        # what makes a reloaded proposal rebuild with the logo it shipped
        # with rather than silently reverting to the placeholder.
        restored_logo_path = st.session_state.get("restored_logo_path")
        if logo_file is None and restored_logo_path:
            st.caption(f"Using the logo stored with this proposal "
                       f"(`{Path(restored_logo_path).name}`). Upload one to replace it.")
        elif logo_file is None and st.session_state.get("restored_logo_missing"):
            st.warning("This proposal had a logo, but it couldn't be fetched from storage — "
                       "the placeholder will be used unless you upload one.")
        st.caption("Meeting notes go in **Draft from notes** at the top of the page — that's "
                   "the copy Claude reads, and it's kept with the proposal history. "
                   "Client-facing wording lives in Campaign Specs below.")

    vertical_key = VERTICALS[vertical_choice]
    market_label = "Washington, DC DMA" if market_choice == "DC" else "Harrisburg DMA"

    if vertical_key == "healthcare":
        st.info("Healthcare targeting (Crossix) is automatically included for the Healthcare vertical.")
    if vertical_key == "travel":
        st.info("Arrivalist destination attribution is automatically included for Travel & Tourism.")
    if vertical_key == "auto":
        st.info("Polk Audiences / Polk Signals automotive targeting is automatically included for Automotive.")

    # ---------------- Section B: Deck scope ----------------
    st.header("B. Deck scope")
    col1, col2 = st.columns(2)
    with col1:
        # Keyed so the History page can rehydrate them. A keyed widget's
        # session_state entry beats its value=/index= argument, which is
        # exactly what loading a saved proposal needs.
        preset = st.radio("Preset", list(PRESETS.keys()), index=1, horizontal=True, key="preset")
        preset_key = PRESETS[preset]
        if preset_key == "quick_pitch":
            st.caption("Quick Pitch: client title, What You Told Us, Why Premion, and one targeting/vertical slide, plus selected add-ons.")
        elif preset_key == "standard":
            st.caption("Standard: a fixed core slide set (cover, specs, intro, one Premium Content highlight, personalized targeting, core attribution, media plan) plus selected add-ons.")
        else:
            st.caption("Extended: the full core-content deck plus selected add-ons.")
    with col2:
        tegna_positioning = st.toggle("Include TEGNA media positioning slides", value=False,
                                       key="tegna_positioning")
        include_vertical_slides = True
        if vertical_key != "none":
            include_vertical_slides = st.toggle(f"Include {vertical_choice} vertical slides",
                                                 value=True, key="include_vertical_slides")
        # The personalized avails table is vertical-independent -- audiences
        # exist with or without one, so this toggle (and Section D2 below) is
        # always available. A vertical only adds the static-targeting-slide
        # fallback when the personalized table is switched off.
        include_avails_template = st.toggle("Include personalized targeting / avails table",
                                             value=True, key="include_avails_template")
        if not include_avails_template and vertical_key != "none":
            st.caption(f"The {vertical_choice} vertical's own static Precision Targeting slide will be included instead.")

    # ---------------- Section C: Products ----------------
    st.header("C. Products")
    ai_section_badge("products")
    col1, col2, col3 = st.columns(3)
    with col1:
        premion_streaming_tv = st.checkbox("Premion Streaming TV", value=True, key="premion_streaming_tv",
                                            on_change=_clear_ai_section, args=("products",))
        streaming_retargeting_enabled = st.checkbox("Streaming Retargeting", key="streaming_retargeting_enabled",
                                                     on_change=_clear_ai_section, args=("products",))
        sr_display = sr_preroll = False
        if streaming_retargeting_enabled:
            sr_display = st.checkbox("  Display", key="sr_disp", on_change=_clear_ai_section, args=("products",))
            sr_preroll = st.checkbox("  Pre-Roll", key="sr_pre", on_change=_clear_ai_section, args=("products",))
    with col2:
        am_enabled = st.checkbox("Audience Marketplace", key="am_enabled",
                                  on_change=_clear_ai_section, args=("products",))
        am_at_display = am_at_preroll = False
        am_gf_display = am_gf_preroll = False
        am_site_display = am_site_preroll = False
        if am_enabled:
            st.caption("Audience Targeting")
            am_at_display = st.checkbox("  Display", key="am_at_disp", on_change=_clear_ai_section, args=("products",))
            am_at_preroll = st.checkbox("  Pre-Roll", key="am_at_pre", on_change=_clear_ai_section, args=("products",))
            st.caption("Geofencing")
            am_gf_display = st.checkbox("  Display", key="am_gf_disp", on_change=_clear_ai_section, args=("products",))
            am_gf_preroll = st.checkbox("  Pre-Roll", key="am_gf_pre", on_change=_clear_ai_section, args=("products",))
            st.caption("Site Retargeting")
            am_site_display = st.checkbox("  Display", key="am_srd", on_change=_clear_ai_section, args=("products",))
            am_site_preroll = st.checkbox("  Pre-Roll", key="am_srp", on_change=_clear_ai_section, args=("products",))
    with col3:
        total_tv = st.checkbox("Total TV", value=False, key="total_tv",
                                on_change=_clear_ai_section, args=("products",))
        dynamic_creative = st.checkbox(
            "Dynamic Video Ads", value=False, key="dynamic_creative",
            help="Adds the Dynamic Video Ad slide and a one-time creative build fee to the plan.",
            on_change=_clear_ai_section, args=("products",))
        live_sports_enabled = st.checkbox("Live Sports", value=False, key="live_sports_enabled",
                                           on_change=_clear_ai_section, args=("products",))
        selected_sports = []
        if live_sports_enabled:
            selected_sports = st.multiselect("Sports packages", list(SPORTS.keys()), key="selected_sports",
                                              on_change=_clear_ai_section, args=("products",))

    # ---------------- Section D: Targeting & attribution ----------------
    st.header("D. Targeting & attribution")
    ai_section_badge("attribution")
    st.caption("Standard (always included): Dashboard, Reporting, Web Attribution.")
    col1, col2 = st.columns(2)
    with col1:
        spanish_campaign = st.toggle("Spanish-language campaign?", key="spanish_campaign",
                                      on_change=_clear_ai_section, args=("attribution",))
        first_party_data = st.checkbox("First-party data targeting", key="first_party_data",
                                        on_change=_clear_ai_section, args=("attribution",))
        linear_reach_extension = st.checkbox("Linear reach extension", disabled=not total_tv,
                                              help="Only available when Total TV is selected",
                                              key="linear_reach_extension",
                                              on_change=_clear_ai_section, args=("attribution",))
        if not total_tv:
            linear_reach_extension = False
    with col2:
        sales_attribution = st.checkbox("Sales attribution", key="sales_attribution",
                                         on_change=_clear_ai_section, args=("attribution",))
        brand_lift = st.checkbox("Brand lift", key="brand_lift",
                                  on_change=_clear_ai_section, args=("attribution",))
        commercial_production = st.checkbox("Commercial production", key="commercial_production",
                                             on_change=_clear_ai_section, args=("attribution",))

    # ---------------- Section D2: Audiences & avails ----------------
    avails_rows = []
    if include_avails_template:
        st.header("D2. Audiences & avails")
        ai_section_badge("avails")
        if "avails_seed_rows" not in st.session_state:
            st.session_state["avails_seed_rows"] = [{"Audience": "", "Geo": market_label, "Max Monthly Avails": 0}]
        if "avails_version" not in st.session_state:
            st.session_state["avails_version"] = 0
        default_avails = pd.DataFrame(st.session_state["avails_seed_rows"])
        avails_editor_key = f"avails_editor_{st.session_state['avails_version']}"
        avails_df = st.data_editor(default_avails, num_rows="dynamic", key=avails_editor_key, use_container_width=True,
                                    on_change=_clear_ai_section, args=("avails",))
        avails_df["Max Monthly Avails"] = avails_df["Max Monthly Avails"].fillna(0)
        total_avails_val = int(avails_df["Max Monthly Avails"].sum())
        st.caption(f"Total avails: {total_avails_val:,}")
        for _, row in avails_df.iterrows():
            if str(row["Audience"]).strip():
                avails_rows.append({
                    "audience": str(row["Audience"]),
                    "geo": str(row["Geo"]),
                    "avails": f"{int(row['Max Monthly Avails']):,}",
                })

        catalog_rfp_lookup = dict(zip(audience_catalog["segment"], audience_catalog["rfp_selectable"]))
        custom_in_table = sum(1 for r in avails_rows if not catalog_rfp_lookup.get(r["audience"], True))
        if custom_in_table > 1:
            st.warning(f"{custom_in_table} custom (non-RFP-selectable) audiences are in the table above -- "
                       f"only one is allowed per campaign. Review before generating.")

        render_audience_finder(avails_df, market_label, vertical_key)

    # ---------------- Section A2: Campaign Specs (manual copy) ----------------
    st.header("Campaign Specs copy")
    ai_section_badge("specs")
    st.caption("Typed by hand, or drafted from notes above. One bullet per line. Audience/Geography feed the media plan's Targeting/Geo defaults below.")
    spec_col1, spec_col2 = st.columns(2)
    with spec_col1:
        goals_text = st.text_area("Goals & Approach", height=90, key="goals_text",
                                   on_change=_clear_ai_section, args=("specs",))
        audience_text = st.text_area("Audience", height=90, key="audience_text",
                                      on_change=_clear_ai_section, args=("specs",))
        geography_text = st.text_area("Geography", height=90, value=market_label, key="geography_text",
                                       on_change=_clear_ai_section, args=("specs",))
    with spec_col2:
        budget_text = st.text_area("Budget & Allocation", height=90, key="budget_text",
                                    on_change=_clear_ai_section, args=("specs",))
        placements_text = st.text_area("Placements & Creative", height=90, key="placements_text",
                                        on_change=_clear_ai_section, args=("specs",))
        timing_text = st.text_area("Timing", height=90, key="timing_text",
                                    help="Narrative copy for the Campaign Specs slide. Actual flight dates for the media plan are set below.",
                                    on_change=_clear_ai_section, args=("specs",))

    default_targeting = first_line(audience_text)
    default_geo = first_line(geography_text) or market_label

    # ---------------- Section E: Proposal / media plan ----------------
    # Needed before the grid renders, not just for the preview -- the grid's
    # own Cost <-> Impressions reconciliation runs through it.
    markup = 1.15 if agency_involved else 1.0

    st.header("E. Proposal / media plan")
    ai_section_badge("media_plan")
    st.caption("Up to three plan options (good/better/best, or several budgets). "
               "Cost = Impressions/1000 x CPM, gross x1.15 if agency toggle is on.")

    st.subheader("Flight")
    ai_section_badge("flight")
    fcol1, fcol2 = st.columns(2)
    with fcol1:
        flight_start = st.date_input("Flight start", value=date(2026, 9, 1), key="flight_start",
                                      on_change=_clear_ai_section, args=("flight",))
    with fcol2:
        flight_end = st.date_input("Flight end", value=date(2026, 11, 30), key="flight_end",
                                    on_change=_clear_ai_section, args=("flight",))

    all_months = month_list(flight_start, flight_end)
    active_months = st.multiselect(
        "Active months (uncheck to skip a month -- custom flighting)",
        all_months, default=all_months, key="active_months",
    )
    # A rehydrated proposal can name months its (also rehydrated) flight
    # dates don't span if the two ever disagree -- keep only real ones, or
    # the multiselect raises on a value not in its options.
    active_months = [m for m in active_months if m in all_months]
    n_months = max(1, len(active_months))
    flight_label = format_flight_label(all_months, active_months) or "TBD"
    st.caption(f"{n_months} active month(s): {flight_label}")

    # Read back from session_state (via the same helpers the draft handler
    # uses) rather than rebuilding this dict from local variables here --
    # every widget above is keyed, so session_state already mirrors them,
    # and sharing one construction keeps main() and apply_draft_to_form from
    # ever silently drifting apart on this shape.
    seed_selections = read_seed_selections()
    products_selection = seed_selections["products"]
    product_seed_key = str(seed_selections)
    # Shared Audience/Geography/Flight fields drive a *soft* update: only
    # rows the user hasn't touched get refreshed. Rows the user has edited
    # (tracked via media_plan_dirty) keep whatever they typed -- editing a
    # shared field never silently wipes a customized line.
    shared_fields_key = default_targeting + "||" + default_geo + "||" + flight_label

    def _seed_option_rows():
        return seed_media_plan_rows(seed_selections, default_geo, default_targeting, flight_label)

    # .get() rather than [] on the two seed keys: they're written alongside
    # plan_options everywhere that sets it, but a missing key should re-seed
    # rather than raise.
    if not st.session_state.get("plan_options"):
        st.session_state["plan_options"] = [new_plan_option(DEFAULT_OPTION_NAMES[0], _seed_option_rows())]
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
    elif st.session_state.get("_product_seed_key") != product_seed_key:
        # Product selections changed -- which tactics exist is a structural
        # change, so every option's row list is rebuilt from scratch. Options
        # are variants of one product mix, so they all follow the mix.
        for opt in st.session_state["plan_options"]:
            rows = _seed_option_rows()
            opt["rows"] = rows
            opt["dirty"] = [False] * len(rows)
            opt["driver"] = [DRIVER_IMPRESSIONS] * len(rows)
            opt["version"] += 1
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
    elif st.session_state.get("_shared_fields_key") != shared_fields_key:
        for opt in st.session_state["plan_options"]:
            for row, dirty in zip(opt["rows"], opt["dirty"]):
                if not dirty:
                    row.update(resolve_row_defaults(row.get("Tactic", ""), default_geo, default_targeting, flight_label))
            opt["version"] += 1
        st.session_state["_shared_fields_key"] = shared_fields_key

    plan_options = st.session_state["plan_options"]

    # A markup change (the agency toggle) re-derives every row of every
    # option -- it moves the cost/impressions relationship itself, not just
    # one row's cells.
    markup_changed = st.session_state.get("media_plan_markup") != markup
    st.session_state["media_plan_markup"] = markup

    ocol1, ocol2, ocol3 = st.columns([1.2, 1.2, 3])
    with ocol1:
        copy_choices = ["(blank plan)"] + [o["name"] for o in plan_options]
        copy_from = st.selectbox("New option starts from", copy_choices,
                                  index=1 if plan_options else 0,
                                  disabled=len(plan_options) >= MAX_PLAN_OPTIONS)
    with ocol2:
        st.caption("")  # aligns the button with the selectbox above
        if st.button("➕ Add option", disabled=len(plan_options) >= MAX_PLAN_OPTIONS):
            name = next_option_name([o["name"] for o in plan_options])
            if copy_from == "(blank plan)":
                plan_options.append(new_plan_option(name, _seed_option_rows()))
            else:
                source = next(o for o in plan_options if o["name"] == copy_from)
                plan_options.append(copy_plan_option(source, name))
            bump_plan_options_generation()
            st.rerun()
    with ocol3:
        if len(plan_options) > 1:
            st.caption("")
            remove_pick = st.selectbox("Remove", [o["name"] for o in plan_options],
                                        label_visibility="collapsed", key="option_remove_pick")
            if st.button("🗑 Remove option"):
                st.session_state["plan_options"] = [o for o in plan_options if o["name"] != remove_pick]
                bump_plan_options_generation()
                st.rerun()

    if len(plan_options) == 1:
        st.caption("One option -- the deck renders exactly as a single-plan proposal, with no option label anywhere.")
    else:
        st.caption(f"{len(plan_options)} options -- each gets its own media plan slide, in this order, "
                   f"with its name appended to the plan title.")

    option_results = []
    tabs = st.tabs([o["name"] for o in plan_options])
    rerun_needed = markup_changed
    # Part of every per-option widget key -- see bump_plan_options_generation.
    gen = st.session_state.get("plan_options_gen", 0)

    for idx, (tab, option) in enumerate(zip(tabs, plan_options)):
        with tab:
            ncol1, ncol2 = st.columns([2, 2])
            with ncol1:
                option["name"] = st.text_input(
                    "Option name", value=option["name"], key=f"option_name_{gen}_{idx}",
                    help="Shown in the deck as part of the plan title, e.g. \"CTV Strategy — Good\".") or option["name"]
            with ncol2:
                option["breakout"] = st.radio(
                    "Breakout", BREAKOUT_MODES, horizontal=True, key=f"option_breakout_{gen}_{idx}",
                    index=BREAKOUT_MODES.index(option["breakout"]))

            breakout_mode = option["breakout"]
            basis = "Monthly" if breakout_mode.startswith("Monthly") else "Full Flight"

            edited_df = st.data_editor(
                pd.DataFrame(option["rows"]), num_rows="dynamic",
                key=f"media_plan_editor_{idx}_{option['version']}", use_container_width=True,
                column_config={
                    "Impressions": st.column_config.NumberColumn(f"Impressions ({basis})"),
                    "Type": st.column_config.SelectboxColumn(options=[ROW_TYPE_RATE, ROW_TYPE_FLAT_FEE]),
                    "Cost": st.column_config.NumberColumn(f"Cost ({basis}, $)", format="$%.0f"),
                },
            )
            st.caption("Impressions and Cost are two views of the same line -- type either one and the other is "
                       "recalculated from the CPM (gross, if the agency toggle is on). Whichever you typed last is "
                       "the one that's kept; changing the CPM re-derives the other from it. "
                       "Duplicate a line (e.g. same product, different audience), then edit the copy. "
                       "Rows you've customized won't auto-update when Audience/Geography/flight dates change above. "
                       "Set Type to Flat Fee for a one-time cost (e.g. a production fee) -- Impressions/CPM are ignored "
                       "for that row and its Cost is the full-flight amount, not multiplied by month count.")

            if reconcile_plan_rows(option, edited_df.to_dict("records"), markup):
                rerun_needed = True

            rows_now = option["rows"]
            dcol1, dcol2 = st.columns([3, 1])
            tactic_labels = [f"{i}: {row.get('Tactic', '') or '(blank)'}" for i, row in enumerate(rows_now)]
            with dcol1:
                dup_pick = st.selectbox("Line to duplicate", tactic_labels, label_visibility="collapsed",
                                         key=f"dup_pick_{idx}") if tactic_labels else None
            with dcol2:
                if st.button("Duplicate line", disabled=not tactic_labels, key=f"dup_btn_{idx}"):
                    dup_i = int(dup_pick.split(":")[0])
                    option["rows"] = rows_now + [dict(rows_now[dup_i])]
                    option["dirty"] = option["dirty"] + [True]  # a duplicate is immediately customizable
                    option["driver"] = option["driver"] + [option["driver"][dup_i]]
                    option["version"] += 1
                    st.rerun()

            totals = compute_plan_totals(option["rows"], breakout_mode, n_months, flight_label)
            option_results.append(totals)

            preview_display = pd.DataFrame([
                {
                    "tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
                    "monthly impressions": "--" if r["is_flat_fee"] else f"{int(r['monthly_impressions']):,}",
                    "monthly cost": f"${r['monthly_cost']:,.0f}",
                    "full flight impressions": "--" if r["is_flat_fee"] else f"{int(r['full_flight_impressions']):,}",
                    "full flight cost": f"${r['full_flight_cost']:,.0f}",
                }
                for r in totals["preview_rows"]
            ]) if totals["preview_rows"] else pd.DataFrame(columns=[
                "tactic", "flight", "geo", "targeting", "monthly impressions", "monthly cost",
                "full flight impressions", "full flight cost"])
            st.dataframe(preview_display, use_container_width=True)

            gross_suffix = " gross" if agency_involved else ""
            st.caption(f"Monthly totals: {int(totals['monthly_impressions']):,} impressions / "
                       f"${totals['monthly_cost']:,.0f}{gross_suffix}")
            st.caption(f"**Full Flight Total ({n_months} month{'s' if n_months != 1 else ''}): "
                       f"{int(totals['full_flight_impressions']):,} impressions / "
                       f"${totals['full_flight_cost']:,.0f}{gross_suffix}**")

    # Recomputed values live in session_state now but the grids on screen
    # still show what was typed, so re-render once. Guarded on an actual
    # change: on the next run every row already matches its own derivation,
    # nothing recomputes, and the loop ends.
    if rerun_needed:
        for opt in plan_options:
            opt["version"] += 1
        st.rerun()

    if len(plan_options) > 1:
        st.markdown("**All options**")
        st.dataframe(pd.DataFrame([
            {"option": o["name"], "lines": len(t["preview_rows"]), "breakout": o["breakout"],
             "monthly cost": f"${t['monthly_cost']:,.0f}",
             "full flight cost": f"${t['full_flight_cost']:,.0f}"}
            for o, t in zip(plan_options, option_results)
        ]), use_container_width=True)

    included_list = build_included_list(
        {"sales_attribution": sales_attribution, "brand_lift": brand_lift},
        commercial_production,
    )
    st.caption("Included with Campaign: " + ", ".join(included_list))

    # ---------------- Case studies ----------------
    selected_case_studies = render_case_study_picker(vertical_key, vertical_choice)

    # ---------------- Generate ----------------
    st.header("Generate")
    tcol1, tcol2 = st.columns([2, 1])
    with tcol1:
        proposal_title = st.text_input("Proposal title (appears on cover + media plan)",
                                        value="Total TV Strategy" if total_tv else "CTV Strategy",
                                        key="proposal_title")
    with tcol2:
        # History-only annotation: it labels the row, never the deck. Loading
        # a proposal prefills a thread-position default, since that's the
        # case where a label is actually worth having.
        revision_label = st.text_input(
            "Revision label (optional)", key="revision_label",
            placeholder="Revision 2",
            help="Shown on the Proposal history page to tell this build apart from "
                 "others for the same client. It never appears in the deck.")

    if st.button("Generate proposal", type="primary"):
        selections = {
            "preset": preset_key,
            "market": market_choice,
            "vertical": vertical_key if include_vertical_slides else "none",
            "agency_involved": agency_involved,
            "spanish_campaign": spanish_campaign,
            "dynamic_creative": dynamic_creative,
            "tegna_positioning": tegna_positioning,
            "include_avails_template": include_avails_template and bool(avails_rows),
            "products": products_selection,
        }
        selections["targeting_attribution"] = {
            "first_party_data": first_party_data,
            "linear_reach_extension": linear_reach_extension,
            "sales_attribution": sales_attribution,
            "brand_lift": brand_lift,
        }

        if not avails_rows:
            avails_rows_final = [{"audience": "", "geo": market_label, "avails": "0"}]
            total_avails_str = "0"
        else:
            avails_rows_final = avails_rows
            total_avails_str = f"{sum(int(r['avails'].replace(',', '')) for r in avails_rows):,}"

        gross_note = " (Gross)" if agency_involved else ""
        multiple_options = len(plan_options) > 1

        def _option_payload(option, totals):
            rows = [
                {"tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
                 "impressions": "--" if r["is_flat_fee"] else f"{int(r['monthly_impressions']):,}",
                 "cost": f"${r['monthly_cost']:,.0f}{gross_note}"}
                for r in totals["preview_rows"]
            ] or [{"tactic": "", "flight": flight_label, "geo": market_label, "targeting": "",
                   "impressions": "0", "cost": "$0"}]

            full_flight_total = None
            if n_months > 1 and totals["preview_rows"]:
                full_flight_total = {
                    "label": f"Full Flight Total ({n_months} months)",
                    "impressions": f"{int(totals['full_flight_impressions']):,}",
                    "cost": f"${totals['full_flight_cost']:,.0f}{gross_note}",
                }

            return {
                "plan_title": option_plan_title(proposal_title, option["name"], multiple_options),
                "rows": rows,
                "totals_label": "Monthly Totals",
                "total_impressions": f"{int(totals['monthly_impressions']):,}",
                "total_cost": f"${totals['monthly_cost']:,.0f}{gross_note}",
                "full_flight_total": full_flight_total,
                "included_list": included_list,
            }

        option_payloads = [_option_payload(o, t) for o, t in zip(plan_options, option_results)]

        fill_data = {
            "client_name": client_name or "Client",
            "proposal_title": proposal_title,
            # Precedence: a freshly uploaded file, then the logo restored
            # with a loaded proposal, then the placeholder.
            "logo_path": (io.BytesIO(logo_file.getvalue()) if logo_file
                          else restored_logo_path or "placeholder_logo.png"),
            "vertical_display": vertical_choice if vertical_key != "none" else "",
            "campaign_specs": {
                "GOALS_BULLETS": lines_to_bullets(goals_text) or ["--"],
                "AUDIENCE_BULLETS": lines_to_bullets(audience_text) or ["--"],
                "GEOGRAPHY_BULLETS": lines_to_bullets(geography_text) or ["--"],
                "BUDGET_BULLETS": lines_to_bullets(budget_text) or ["--"],
                "PLACEMENTS_BULLETS": lines_to_bullets(placements_text) or ["--"],
                "TIMING_BULLETS": lines_to_bullets(timing_text) or [flight_label],
            },
            "avails": {
                "rows": avails_rows_final,
                "total_avails": total_avails_str,
            },
            # One entry per option, in tab order. assembly.personalize clones
            # the media plan template once per extra option and fills each
            # independently; a single-option list renders exactly as before.
            "media_plan_options": option_payloads,
        }

        master_path, deck_version_id, deck_warning = db.master_deck(LOCAL_MASTER_DECK_PATH)
        if master_path is None:
            # No deck in Supabase and none on disk -- there is nothing to
            # build from, so stop here with the explanation rather than
            # letting python-pptx raise on a path that doesn't exist.
            st.error(deck_warning)
            st.stop()
        if deck_warning:
            st.warning(deck_warning)

        # Fetched before assembly so a vault problem surfaces as its own
        # message rather than as a failure part-way through building a deck.
        case_study_sources, case_study_errors = [], []
        for case_study in selected_case_studies:
            try:
                case_study_sources.append(
                    {"path": db.case_study_file(case_study["id"], case_study["storage_path"]),
                     "slides": None, "title": case_study["title"]})
            except Exception as exc:
                case_study_errors.append(f"{case_study['title']}: {db.describe_error(exc)}")
        for message in case_study_errors:
            st.warning(f"Case study left out -- couldn't fetch it. {message}")

        with st.spinner("Assembling deck..."):
            try:
                prs, original_count, kept_count = assembly.build_presentation(master_path, selections)
                # Before personalize, deliberately -- see
                # case_study_insert_index: the plan slide is found by its
                # {{PLAN_TITLE}} token (which personalize consumes), and
                # personalize clones it per extra option directly after the
                # original, so inserting here is what puts case studies ahead
                # of the first option rather than between options.
                case_study_slides = assembly.append_case_studies(prs, case_study_sources)
                assembly.personalize(prs, fill_data)
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)
            except Exception as exc:
                st.error(f"Assembly failed: {exc}")
                raise

        extra_option_slides = len(option_payloads) - 1
        extras = []
        if extra_option_slides:
            extras.append(f"{extra_option_slides} extra media plan slide"
                          f"{'s' if extra_option_slides != 1 else ''} for options B/C")
        if case_study_slides:
            extras.append(f"{case_study_slides} case study slide"
                          f"{'s' if case_study_slides != 1 else ''} from "
                          f"{len(case_study_sources)} case stud"
                          f"{'ies' if len(case_study_sources) != 1 else 'y'}")
        st.success(f"Assembled {kept_count + extra_option_slides + case_study_slides} of "
                   f"{original_count} slides"
                   + (f" (including {' and '.join(extras)})." if extras else "."))

        output_filename = f"{(client_name or 'client').replace(' ', '_')}_proposal.pptx"

        # Store the logo so a rebuild is faithful. An uploaded file isn't
        # part of form_json, so without this "Rebuild as presented" quietly
        # fell back to the placeholder. A reloaded proposal keeps pointing at
        # the blob it already has rather than re-uploading identical bytes.
        logo_storage_path = None
        if logo_file is not None:
            logo_storage_path, logo_error = db.upload_proposal_logo(
                f"{(client_name or 'client').replace(' ', '_')}", logo_file.getvalue(),
                logo_file.name)
            if logo_error:
                st.caption(f"⚠️ The logo couldn't be stored ({logo_error}) — this deck is fine, "
                           f"but rebuilding it later will fall back to the placeholder.")
        elif restored_logo_path:
            logo_storage_path = st.session_state.get("restored_logo_storage_path")

        # Log the whole form state, every option included. Capture only --
        # nothing reads it back yet, so a failed write is worth a caption but
        # must not cost the seller the deck they just generated.
        _, log_error = db.log_proposal(
            client_name=client_name,
            vertical=vertical_key,
            market=market_choice,
            form_json={
                "proposal_title": proposal_title,
                "selections": selections,
                "agency_involved": agency_involved,
                "markup": markup,
                "flight": {"start": str(flight_start), "end": str(flight_end),
                           "label": flight_label, "active_months": [str(m) for m in active_months]},
                "campaign_specs": {
                    "goals": goals_text, "audience": audience_text, "geography": geography_text,
                    "budget": budget_text, "placements": placements_text, "timing": timing_text,
                },
                "avails_rows": avails_rows_final,
                "included_list": included_list,
                "plan_options": [
                    {"name": option["name"], "breakout": option["breakout"],
                     "rows": option["rows"], "driver": option["driver"],
                     "totals": {k: v for k, v in totals.items() if k != "preview_rows"}}
                    for option, totals in zip(plan_options, option_results)
                ],
                "deck_payload": {"media_plan_options": option_payloads},
                "case_studies": [{"id": c["id"], "title": c["title"]}
                                 for c in selected_case_studies],
                # Whether a logo was used at all, independent of whether
                # storing it succeeded -- that's what tells a later rebuild
                # "the placeholder is wrong here" versus "there was none".
                "logo_used": bool(logo_file or restored_logo_path),
                # The notes themselves, not just the round number. They're the
                # only record of *why* a proposal looks the way it does, and
                # they were previously typed into the form and thrown away.
                "draft": {"round": st.session_state.get("draft_round"),
                          "notes": st.session_state.get("draft_notes_input") or None,
                          "clarifications": st.session_state.get("draft_clarifications") or None,
                          "unresolved": st.session_state.get("draft_unresolved")},
            },
            deck_version_id=deck_version_id,
            output_filename=output_filename,
            # History is append-only: a proposal loaded from the History page
            # logs a NEW row linked back to its source rather than updating
            # it. What a client was actually sent stays as it was logged.
            parent_proposal_id=st.session_state.get("history_parent_id"),
            revision_label=(revision_label or "").strip() or None,
            logo_storage_path=logo_storage_path,
            created_by=current_user(),
        )
        if log_error:
            st.caption(f"⚠️ Proposal history not recorded: {log_error}")
        elif st.session_state.get("history_parent_id"):
            st.caption("Logged as a new revision, linked to the proposal it was loaded from.")
        # Consumed once: a second generate from the same loaded state is a
        # sibling revision of the original, not a child of itself.
        st.session_state["history_parent_id"] = None

        st.download_button(
            "Download .pptx",
            data=buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


# Streamlit execs the entrypoint script as "__main__", so the app still runs
# normally under `streamlit run app.py`. The guard is what lets a test import
# this module for its pure helpers without rendering the whole form.
if __name__ == "__main__":
    main()
