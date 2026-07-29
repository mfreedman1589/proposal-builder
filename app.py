"""
app.py -- Phase 1 Streamlit form for the Premion Proposal Builder (build spec
section 5), wired directly to assembly.py's selection -> deletion -> fill
pipeline. No Supabase, no Claude API yet: the product list/CPMs are a
hardcoded dict below, and Campaign Specs bullets are typed in by hand.
"""

import io
import json
from datetime import date, datetime

import anthropic
import pandas as pd
import streamlit as st

import assembly
from audience_catalog import load_audience_catalog, validate_segments

st.set_page_config(page_title="Premion Proposal Builder", layout="wide")

ANTHROPIC_MODEL = "claude-sonnet-4-6"
ANTHROPIC_MAX_TOKENS = 2000

MASTER_DECK_PATH = assembly.MASTER_DECK_PATH

# Cached at startup -- both the PDF parse and the times_used CSV merge only
# need to happen once per process. Not consumed yet: the "Draft from notes"
# and "Audience finder" features that read this catalog land in follow-up
# commits.
audience_catalog = load_audience_catalog()

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

# Hardcoded product list + CPM defaults (spec section 3 "products" table,
# no Supabase yet). line_type mirrors the DB schema's premion/am/broadcast.
# All Audience Marketplace Display tactics are $5.50 except Geofencing
# ($9.00); all AM Pre-Roll tactics are $21.00.
PRODUCTS = {
    "premion_streaming_tv": {"label": "Premion Streaming TV", "default_cpm": 32.00, "line_type": "premion"},
    "streaming_retargeting_display": {"label": "Streaming Retargeting - Display", "default_cpm": 5.50, "line_type": "premion"},
    "streaming_retargeting_preroll": {"label": "Streaming Retargeting - Pre-Roll", "default_cpm": 21.00, "line_type": "premion"},
    "audience_targeting_display": {"label": "Audience Targeting - Display", "default_cpm": 5.50, "line_type": "am"},
    "audience_targeting_preroll": {"label": "Audience Targeting - Pre-Roll", "default_cpm": 21.00, "line_type": "am"},
    "geofencing_display": {"label": "Geofencing - Display", "default_cpm": 9.00, "line_type": "am"},
    "geofencing_preroll": {"label": "Geofencing - Pre-Roll", "default_cpm": 21.00, "line_type": "am"},
    "site_retargeting_display": {"label": "Site Retargeting - Display", "default_cpm": 5.50, "line_type": "am"},
    "site_retargeting_preroll": {"label": "Site Retargeting - Pre-Roll", "default_cpm": 21.00, "line_type": "am"},
    "broadcast_tv": {"label": "Broadcast Schedule", "default_cpm": 5.50, "line_type": "broadcast"},
}

# Per-sport net CPMs from PREMION_Live Sports Rates.xlsx ("2026 Prem Core
# Live Sports" sheet, TEGNA Recommended Rate column). Falls back to a flat
# default for any sport key not on the ratecard.
SPORT_CPM = {
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

STREAMING_RETARGETING_TARGETING = "Retarget Exposed CTV Viewers"
LIVE_SPORTS_TARGETING = "100% Live, 100% In-Game, 100% CTV"
MEDIA_PLAN_FIELDS = ["Tactic", "Flight", "Geo", "Targeting", "Impressions", "CPM", "Type", "Flat Cost"]
ROW_TYPE_RATE = "Rate"
ROW_TYPE_FLAT_FEE = "Flat Fee"

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

SPORT_LABEL_BY_VALUE = {v: k for k, v in SPORTS.items()}

ATTRIBUTION_OPTIONS = ["web", "sales", "brand_lift", "first_party", "linear_reach_ext", "commercial_production"]
# "web" is always included by default (build_included_list) -- there's no
# form toggle for it, so it has no entry here.
ATTRIBUTION_FIELD_MAP = {
    "sales": "sales_attribution",
    "brand_lift": "brand_lift",
    "first_party": "first_party_data",
    "linear_reach_ext": "linear_reach_extension",
    "commercial_production": "commercial_production",
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
}
VERTICAL_HINT_SYNONYMS = {
    "bank": "banking", "hospital": "healthcare", "clinic": "healthcare", "medical": "healthcare",
    "car dealer": "auto", "dealership": "auto", "auto dealer": "auto",
    "restaurant": "dining_qsr", "qsr": "dining_qsr", "fast food": "dining_qsr",
    "hotel": "travel", "tourism": "travel", "resort": "travel",
    "school": "education", "university": "education", "college": "education",
}

CUSTOM_FEE_PRODUCT = "custom_fee"

DRAFT_JSON_SCHEMA_EXAMPLE = """{
  "client_name": "", "vertical": "", "market": "DC|Harrisburg",
  "agency_involved": false, "spanish_campaign": false,
  "flight_start": "YYYY-MM-DD", "flight_end": "YYYY-MM-DD", "geo": "",
  "total_budget": 0,
  "media_plan_lines": [
    {"product": "premion_streaming_tv", "label": "Commercial", "audience_track": "SMB owners and business executives", "allocation": {"percent_of_remainder": 60}},
    {"product": "premion_streaming_tv", "label": "Retail", "audience_track": "consumers in-market for deposit accounts", "allocation": {"percent_of_remainder": 40}},
    {"product": "streaming_retargeting_display", "allocation": {"percent_of_total": 10}},
    {"product": "custom_fee", "label": "Dynamic Ad Creation", "allocation": {"flat_amount": 850}}
  ],
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


def _clear_ai_section(section):
    st.session_state.get("ai_filled_sections", set()).discard(section)


def ai_section_badge(section):
    if section in st.session_state.get("ai_filled_sections", set()):
        st.caption("🤖 Some fields below were drafted from your notes -- review before generating.")


def _detect_vertical_hint(notes):
    low = notes.lower()
    for label, key in VERTICALS.items():
        if key == "none":
            continue
        if key.replace("_", " ") in low or label.lower() in low:
            return key
    for word, key in VERTICAL_HINT_SYNONYMS.items():
        if word in low:
            return key
    return None


def build_catalog_slice(vertical_hint, cap=150):
    catalog = load_audience_catalog()
    if vertical_hint and vertical_hint in VERTICAL_CATEGORY_HINTS:
        cats = VERTICAL_CATEGORY_HINTS[vertical_hint]
        relevant = catalog[catalog["category"].isin(cats)].sort_values("times_used", ascending=False)
        rest = catalog[~catalog["category"].isin(cats)].sort_values("times_used", ascending=False)
        combined = pd.concat([relevant, rest])
    else:
        combined = catalog.sort_values("times_used", ascending=False)
    sliced = combined.head(cap)
    return sliced[["segment", "category", "subcategory", "rfp_selectable", "times_used"]].to_dict("records")


def build_draft_prompt(notes):
    vertical_hint = _detect_vertical_hint(notes)
    catalog_slice = build_catalog_slice(vertical_hint)
    products_info = {k: {"label": v["label"], "default_cpm": v["default_cpm"]} for k, v in PRODUCTS.items()}
    today = date.today().isoformat()

    return f"""You are drafting a first pass at a Premion CTV/OTT advertising proposal from raw meeting/discovery notes. Today's date is {today}. Return ONLY valid JSON matching the schema below -- no markdown code fences, no preamble, no explanation, just the JSON object.

Schema:
{DRAFT_JSON_SCHEMA_EXAMPLE}

"media_plan_lines" is the full media plan, expressed one entry per intended row -- not one entry per product. If the notes call for the same product run as separate lines (e.g. two different audience tracks, or a commercial vs. retail split), give each its own entry with its own "label" and "audience_track"; each becomes its own media plan row, with "label" appended to the product's own tactic name (e.g. "Premion Streaming TV — Commercial") and "audience_track" as that row's Targeting. A line with no "label" just uses the product's own name as-is. "product" must be exactly one of: {list(PRODUCTS.keys())}, or the special value "{CUSTOM_FEE_PRODUCT}" for a one-time flat fee that isn't a real media-buy product (e.g. a production/creative fee) -- a "{CUSTOM_FEE_PRODUCT}" line's "label" becomes its Tactic name directly and it must use a "flat_amount" allocation.

Each line's "allocation" has exactly one key:
- "flat_amount": this line costs exactly this many dollars.
- "percent_of_total": this line costs this percent of total_budget.
- "percent_of_remainder": this line costs this percent of whatever's left after all "flat_amount" and "percent_of_total" lines are subtracted from total_budget (percent_of_remainder entries across lines should sum to 100 if they're meant to exhaust the remainder).
- "split_evenly": this line shares equally, with every other "split_evenly" line, in whatever's left after "flat_amount"/"percent_of_total"/"percent_of_remainder" lines are all accounted for -- use this for "split evenly across N audiences/tracks" instead of trying to pre-compute a percentage yourself.
Do NOT do any arithmetic yourself beyond picking which allocation type fits each line -- Python resolves flat_amount and percent_of_total first, then percent_of_remainder, then splits whatever's left evenly across split_evenly lines, then computes every dollar amount, impression count, and markup.

Rules:
- "vertical" must be exactly one of: {list(VERTICALS.values())}
- "market" must be exactly "DC" or "Harrisburg".
- "sports" entries must be exactly one of: {list(SPORTS.values())}
- "attribution" entries must be exactly one of: {ATTRIBUTION_OPTIONS}
- "audiences[].segment" must be an EXACT name from the audience catalog slice below -- do not paraphrase or invent segment names. If nothing in the slice fits, it's fine to omit audiences or note it in "unresolved". {AUDIENCE_MATCH_GUIDANCE} When a media_plan_lines entry's "audience_track" describes the same audience as one of your "audiences" entries, use the same wording for both.
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


def build_audience_finder_prompt(description):
    vertical_hint = _detect_vertical_hint(description)
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


def call_claude_audience_suggest(description):
    """Returns (recommendations_list, unmatched_names, error_message) --
    error_message is None on success. Each recommendation is
    {"segment", "rationale"}, already validated against the catalog."""
    parsed, error = _call_claude_json(build_audience_finder_prompt(description))
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


def read_products_selection():
    """The same products_selection shape main() builds from the live
    widgets, but read directly from session_state -- lets the draft handler
    compute the exact product_seed_key/shared_fields_key the next run will
    see, and lets main() and the draft handler share one source of truth."""
    return {
        "streaming_retargeting": {
            "enabled": st.session_state.get("streaming_retargeting_enabled", False),
            "display": st.session_state.get("sr_disp", False),
            "preroll": st.session_state.get("sr_pre", False),
        },
        "audience_marketplace": {
            "enabled": st.session_state.get("am_enabled", False),
            "audience_targeting_display": st.session_state.get("am_at_disp", False),
            "audience_targeting_preroll": st.session_state.get("am_at_pre", False),
            "geofencing_display": st.session_state.get("am_gf_disp", False),
            "geofencing_preroll": st.session_state.get("am_gf_pre", False),
            "site_retargeting_display": st.session_state.get("am_srd", False),
            "site_retargeting_preroll": st.session_state.get("am_srp", False),
        },
        "live_sports": {
            "enabled": st.session_state.get("live_sports_enabled", False),
            "sports": [SPORTS[s] for s in st.session_state.get("selected_sports", [])],
        },
        "total_tv": st.session_state.get("total_tv", False),
    }


def read_seed_selections():
    return {"products": read_products_selection(), "_premion_streaming_tv": st.session_state.get("premion_streaming_tv", False)}


def apply_draft_to_form(draft):
    """Turns a parsed Claude draft into session_state writes (applied all at
    once at the end, so a mid-processing error leaves the form untouched)
    plus an "unresolved" list shown to the user. Must be called before any
    widget in this run has rendered, followed by st.rerun()."""
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
    else:
        flight_label = "TBD"
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

    sports_in = draft.get("sports", []) or []
    sport_labels = []
    for s in sports_in:
        label = SPORT_LABEL_BY_VALUE.get(s)
        if label:
            sport_labels.append(label)
        else:
            unresolved.append(f"Sport '{s}' not recognized -- skipped.")
    if sport_labels:
        updates["live_sports_enabled"] = True
        updates["selected_sports"] = sport_labels
        touched_sections.add("products")

    # --- audiences ---
    audiences_in = draft.get("audiences", []) or []
    audience_names = [a.get("segment", "") for a in audiences_in if a.get("segment")]
    matched, unmatched = validate_segments(audience_names)
    if unmatched:
        unresolved.append(f"Unrecognized audience segment(s) dropped: {', '.join(unmatched)}")
    matched_audiences = [a for a in audiences_in if a.get("segment") in matched]

    catalog = load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
    custom_count = sum(1 for a in matched_audiences if not rfp_map.get(a["segment"], True))
    if custom_count > 1:
        unresolved.append(
            f"{custom_count} custom (non-RFP-selectable) audiences were drafted, but only one is allowed "
            f"per campaign -- review the avails table before generating.")

    if matched_audiences:
        updates["avails_seed_rows"] = [
            {"Audience": a["segment"], "Geo": a.get("geo") or geo_or_market, "Max Monthly Avails": 0}
            for a in matched_audiences
        ]
        updates["avails_version"] = st.session_state.get("avails_version", 0) + 1
        unresolved.append("Max Monthly Avails left at 0 for drafted audiences -- pull real numbers from the avails system before finalizing.")
        touched_sections.add("avails")

    # --- budget math (no arithmetic performed by the model) ---
    # media_plan_lines is one entry per intended media-plan row (not one per
    # product) -- the same product can appear on multiple lines with
    # different labels/audience tracks (e.g. a Commercial vs. Retail split),
    # each becoming its own row. Resolution is a waterfall, matching the
    # order described to the model: flat_amount and percent_of_total lines
    # are subtracted from total_budget first: percent_of_remainder lines then
    # take their share of what's left; split_evenly lines divide whatever
    # remains after that evenly among themselves.
    total_budget = float(draft.get("total_budget") or 0)
    valid_line_products = set(PRODUCT_TO_WIDGET_KEYS) | {CUSTOM_FEE_PRODUCT}
    lines_in = draft.get("media_plan_lines", []) or []
    lines_valid = []
    for line in lines_in:
        product = line.get("product")
        if product not in valid_line_products:
            unresolved.append(f"Media plan line for unrecognized product '{product}' skipped.")
            continue
        lines_valid.append(line)

    markup = 1.15 if agency_involved else 1.0

    def _impressions_for(amount, cpm):
        return round((amount / (cpm * markup)) * 1000) if cpm else 0

    resolved_amounts = {}
    flat_and_pct_total_sum = 0.0
    for i, line in enumerate(lines_valid):
        alloc = line.get("allocation", {}) or {}
        if "flat_amount" in alloc:
            amt = float(alloc["flat_amount"] or 0)
        elif "percent_of_total" in alloc:
            amt = (float(alloc["percent_of_total"] or 0) / 100.0) * total_budget
        else:
            continue
        resolved_amounts[i] = amt
        flat_and_pct_total_sum += amt

    remainder = max(0.0, total_budget - flat_and_pct_total_sum)
    pct_remainder_sum = 0.0
    for i, line in enumerate(lines_valid):
        alloc = line.get("allocation", {}) or {}
        if "percent_of_remainder" in alloc:
            amt = (float(alloc["percent_of_remainder"] or 0) / 100.0) * remainder
            resolved_amounts[i] = amt
            pct_remainder_sum += amt

    leftover = max(0.0, remainder - pct_remainder_sum)
    split_evenly_indices = [i for i, line in enumerate(lines_valid)
                            if (line.get("allocation") or {}).get("split_evenly")]
    if split_evenly_indices:
        per_each = leftover / len(split_evenly_indices)
        for i in split_evenly_indices:
            resolved_amounts[i] = per_each

    for i, line in enumerate(lines_valid):
        if i not in resolved_amounts:
            resolved_amounts[i] = 0.0
            unresolved.append(f"Media plan line for '{line.get('product')}' has no recognized "
                               f"allocation type -- amount left at $0.")

    media_plan_rows = []
    touched_products = set()
    for i, line in enumerate(lines_valid):
        product = line["product"]
        label = line.get("label", "") or ""
        audience_track = line.get("audience_track", "") or ""
        amount = resolved_amounts[i]

        if product == CUSTOM_FEE_PRODUCT:
            media_plan_rows.append({
                "Tactic": label or "Flat Fee", "Flight": flight_label, "Geo": geo_or_market,
                "Targeting": audience_track, "Impressions": 0, "CPM": 0,
                "Type": "Flat Fee", "Flat Cost": round(amount, 2),
            })
            continue

        touched_products.add(product)
        cpm = PRODUCTS[product]["default_cpm"]
        base_label = PRODUCTS[product]["label"]
        tactic = f"{base_label} — {label}" if label else base_label
        media_plan_rows.append({
            "Tactic": tactic, "Flight": flight_label, "Geo": geo_or_market,
            "Targeting": audience_track or default_targeting,
            "Impressions": _impressions_for(amount, cpm), "CPM": cpm,
            "Type": "Rate", "Flat Cost": 0.0,
        })

    for p in touched_products:
        for widget_key, val in PRODUCT_TO_WIDGET_KEYS[p]:
            updates[widget_key] = val
    if touched_products:
        touched_sections.add("products")

    if media_plan_rows:
        updates["media_plan_rows"] = media_plan_rows
        updates["media_plan_dirty"] = [True] * len(media_plan_rows)
        updates["media_plan_version"] = st.session_state.get("media_plan_version", 0) + 1

        # Prevent main()'s own reseed-on-mismatch logic from immediately
        # overwriting these rows on the very next run: precompute the same
        # product_seed_key/shared_fields_key it will independently derive
        # after rerun, from the *drafted* widget values in `updates` (falling
        # back to current session_state for anything the draft didn't touch)
        # -- not from session_state alone, which still holds pre-draft values
        # at this point in the function.
        def _get(key, default=False):
            return updates.get(key, st.session_state.get(key, default))

        seed_products_selection = {
            "streaming_retargeting": {
                "enabled": _get("streaming_retargeting_enabled"),
                "display": _get("sr_disp"),
                "preroll": _get("sr_pre"),
            },
            "audience_marketplace": {
                "enabled": _get("am_enabled"),
                "audience_targeting_display": _get("am_at_disp"),
                "audience_targeting_preroll": _get("am_at_pre"),
                "geofencing_display": _get("am_gf_disp"),
                "geofencing_preroll": _get("am_gf_pre"),
                "site_retargeting_display": _get("am_srd"),
                "site_retargeting_preroll": _get("am_srp"),
            },
            "live_sports": {
                "enabled": _get("live_sports_enabled"),
                "sports": [SPORTS[s] for s in _get("selected_sports", [])],
            },
            "total_tv": _get("total_tv"),
        }
        updates["_product_seed_key"] = str({"products": seed_products_selection, "_premion_streaming_tv": _get("premion_streaming_tv")})
        updates["_shared_fields_key"] = default_targeting + "||" + geo_or_market + "||" + flight_label

    updates["ai_filled_sections"] = touched_sections
    updates["draft_unresolved"] = unresolved

    for key, value in updates.items():
        st.session_state[key] = value


def _add_segment_to_avails(segment, geo, current_avails_df):
    rows = current_avails_df.to_dict("records")
    rows.append({"Audience": segment, "Geo": geo, "Max Monthly Avails": 0})
    st.session_state["avails_seed_rows"] = rows
    st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
    st.rerun()


def render_audience_finder(avails_df, market_label):
    """Section D2's 'Audience finder' expander: browse/search the catalog,
    or describe the client/campaign and let Claude suggest segments. Either
    way, "Add" appends a row to the current avails table (audience + geo
    filled, avails left blank -- those come from a real system)."""
    catalog = load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))

    with st.expander("🔍 Audience finder", expanded=False):
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
            filtered = filtered.sort_values("times_used", ascending=False).head(50)
            st.caption(f"{len(filtered)} segment(s) shown (top 50 by times used)")

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
                if cols[5].button("Add", key=f"finder_add_{row['segment']}"):
                    _add_segment_to_avails(row["segment"], market_label, avails_df)

        else:
            description = st.text_area("Describe the client or campaign", key="finder_suggest_input", height=100)
            if st.button("Suggest audiences"):
                if not description.strip():
                    st.warning("Describe the client or campaign first.")
                else:
                    with st.spinner("Asking Claude for audience recommendations..."):
                        recs, unmatched, error = call_claude_audience_suggest(description)
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

                existing_segments = [str(s) for s in avails_df["Audience"] if str(s).strip()]
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
                    if cols[5].button("Add", key=f"finder_add_suggest_{seg}"):
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


def resolve_row_defaults(tactic, default_geo, default_targeting, flight_label):
    """What a row's Flight/Geo/Targeting should be right now, given its
    Tactic name -- used both at initial seed time and to soft-update
    not-yet-edited rows when the shared form fields change."""
    if tactic.startswith("Streaming Retargeting"):
        targeting = STREAMING_RETARGETING_TARGETING
    elif tactic.startswith("Live Sports"):
        targeting = LIVE_SPORTS_TARGETING
    else:
        targeting = default_targeting
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

    def _row(label, cpm, targeting):
        return {"Tactic": label, "Flight": flight_label, "Geo": market_label,
                "Targeting": targeting, "Impressions": 0, "CPM": cpm,
                "Type": ROW_TYPE_RATE, "Flat Cost": 0.0}

    if selections.get("_premion_streaming_tv"):
        p = PRODUCTS["premion_streaming_tv"]
        rows.append(_row(p["label"], p["default_cpm"], default_targeting))

    sr = products.get("streaming_retargeting", {})
    if sr.get("display"):
        p = PRODUCTS["streaming_retargeting_display"]
        rows.append(_row(p["label"], p["default_cpm"], STREAMING_RETARGETING_TARGETING))
    if sr.get("preroll"):
        p = PRODUCTS["streaming_retargeting_preroll"]
        rows.append(_row(p["label"], p["default_cpm"], STREAMING_RETARGETING_TARGETING))

    am = products.get("audience_marketplace", {})
    if am.get("enabled"):
        if am.get("audience_targeting_display"):
            p = PRODUCTS["audience_targeting_display"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("audience_targeting_preroll"):
            p = PRODUCTS["audience_targeting_preroll"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("geofencing_display"):
            p = PRODUCTS["geofencing_display"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("geofencing_preroll"):
            p = PRODUCTS["geofencing_preroll"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("site_retargeting_display"):
            p = PRODUCTS["site_retargeting_display"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("site_retargeting_preroll"):
            p = PRODUCTS["site_retargeting_preroll"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))

    sports = products.get("live_sports", {})
    if sports.get("enabled"):
        for sport_key in sports.get("sports", []):
            label = next((k for k, v in SPORTS.items() if v == sport_key), sport_key)
            cpm = SPORT_CPM.get(sport_key, DEFAULT_SPORT_CPM)
            rows.append(_row(f"Live Sports - {label}", cpm, LIVE_SPORTS_TARGETING))

    if products.get("total_tv"):
        p = PRODUCTS["broadcast_tv"]
        rows.append(_row(p["label"], p["default_cpm"], default_targeting))

    if not rows:
        rows.append(_row("", 0.0, ""))

    return rows


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


def main():
    if not check_password():
        return

    st.title("Premion Proposal Builder")
    st.caption("Phase 1 -- hardcoded product/CPM list, manually-typed Campaign Specs copy. No Supabase yet.")

    # ---------------- Draft from notes (Claude) ----------------
    with st.expander("📝 Draft from notes (optional)", expanded=False):
        st.caption("Paste meeting or discovery notes below. Claude drafts a first pass at the form below -- "
                   "review everything before generating, nothing here is final.")
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
                        st.rerun()

    if st.session_state.get("draft_unresolved"):
        st.warning("**Review before generating:**\n\n" +
                   "\n".join(f"- {item}" for item in st.session_state["draft_unresolved"]))

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
        discovery_notes = st.text_area("Discovery notes", height=100, help="Reference only -- superseded by the Draft from notes panel above.")

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
        preset = st.radio("Preset", list(PRESETS.keys()), index=1, horizontal=True)
        preset_key = PRESETS[preset]
        if preset_key == "quick_pitch":
            st.caption("Quick Pitch: client title, What You Told Us, Why Premion, and one targeting/vertical slide, plus selected add-ons.")
        elif preset_key == "standard":
            st.caption("Standard: a fixed core slide set (cover, specs, intro, one Premium Content highlight, personalized targeting, core attribution, media plan) plus selected add-ons.")
        else:
            st.caption("Extended: the full core-content deck plus selected add-ons.")
    with col2:
        tegna_positioning = st.toggle("Include TEGNA media positioning slides", value=True)
        include_vertical_slides = True
        include_avails_template = False
        if vertical_key != "none":
            include_vertical_slides = st.toggle(f"Include {vertical_choice} vertical slides", value=True)
            include_avails_template = st.toggle("Include personalized targeting / avails table", value=True)
            if not include_avails_template:
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
    if vertical_key != "none" and include_avails_template:
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

        render_audience_finder(avails_df, market_label)

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
    st.header("E. Proposal / media plan")
    st.caption("Phase 1 supports a single plan option (Option A). Cost = Impressions/1000 x CPM, gross x1.15 if agency toggle is on.")

    st.subheader("Flight & breakout")
    ai_section_badge("flight")
    fcol1, fcol2, fcol3 = st.columns([1, 1, 1])
    with fcol1:
        flight_start = st.date_input("Flight start", value=date(2026, 9, 1), key="flight_start",
                                      on_change=_clear_ai_section, args=("flight",))
    with fcol2:
        flight_end = st.date_input("Flight end", value=date(2026, 11, 30), key="flight_end",
                                    on_change=_clear_ai_section, args=("flight",))
    with fcol3:
        breakout_mode = st.radio("Breakout", ["Monthly (default)", "Full Flight"], horizontal=True)

    all_months = month_list(flight_start, flight_end)
    active_months = st.multiselect(
        "Active months (uncheck to skip a month -- custom flighting)",
        all_months, default=all_months,
    )
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

    if "media_plan_version" not in st.session_state:
        st.session_state["media_plan_version"] = 0

    if "media_plan_rows" not in st.session_state:
        st.session_state["media_plan_rows"] = seed_media_plan_rows(
            seed_selections, default_geo, default_targeting, flight_label)
        st.session_state["media_plan_dirty"] = [False] * len(st.session_state["media_plan_rows"])
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
    elif st.session_state["_product_seed_key"] != product_seed_key:
        # Product selections changed -- which tactics exist is a structural
        # change, so the row list itself is rebuilt from scratch.
        st.session_state["media_plan_rows"] = seed_media_plan_rows(
            seed_selections, default_geo, default_targeting, flight_label)
        st.session_state["media_plan_dirty"] = [False] * len(st.session_state["media_plan_rows"])
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
        st.session_state["media_plan_version"] += 1
    elif st.session_state["_shared_fields_key"] != shared_fields_key:
        for row, dirty in zip(st.session_state["media_plan_rows"], st.session_state["media_plan_dirty"]):
            if not dirty:
                row.update(resolve_row_defaults(row.get("Tactic", ""), default_geo, default_targeting, flight_label))
        st.session_state["_shared_fields_key"] = shared_fields_key
        st.session_state["media_plan_version"] += 1

    plan_df = pd.DataFrame(st.session_state["media_plan_rows"])
    editor_key = f"media_plan_editor_{st.session_state['media_plan_version']}"
    impressions_label = "Impressions (Monthly)" if breakout_mode.startswith("Monthly") else "Impressions (Full Flight)"
    edited_plan_df = st.data_editor(
        plan_df, num_rows="dynamic", key=editor_key, use_container_width=True,
        column_config={
            "Impressions": st.column_config.NumberColumn(impressions_label),
            "Type": st.column_config.SelectboxColumn(options=[ROW_TYPE_RATE, ROW_TYPE_FLAT_FEE]),
            "Flat Cost": st.column_config.NumberColumn("Flat Cost ($)", format="$%.2f"),
        },
    )
    st.caption("Duplicate a line (e.g. same product, different audience/impressions), then edit the copy. "
               "Rows you've customized won't auto-update when Audience/Geography/flight dates change above. "
               "Set Type to Flat Fee for a one-time cost (e.g. a production fee) -- Impressions/CPM are ignored "
               "for that row and its Flat Cost is the full-flight amount, not multiplied by month count.")

    # Reconcile edits: diff against the pre-render snapshot to update dirty
    # flags, then persist both back to session_state as the new baseline.
    prev_rows = st.session_state["media_plan_rows"]
    prev_dirty = st.session_state["media_plan_dirty"]
    new_rows = edited_plan_df.to_dict("records")
    new_dirty = []
    for i, row in enumerate(new_rows):
        if i < len(prev_rows):
            changed = any(str(row.get(f, "")) != str(prev_rows[i].get(f, "")) for f in MEDIA_PLAN_FIELDS)
            new_dirty.append(prev_dirty[i] or changed)
        else:
            new_dirty.append(True)  # a row added via the grid's own "+" is treated as customized
    st.session_state["media_plan_rows"] = new_rows
    st.session_state["media_plan_dirty"] = new_dirty

    dcol1, dcol2 = st.columns([3, 1])
    tactic_labels = [f"{i}: {row.get('Tactic', '') or '(blank)'}" for i, row in enumerate(new_rows)]
    with dcol1:
        dup_pick = st.selectbox("Line to duplicate", tactic_labels, label_visibility="collapsed") if tactic_labels else None
    with dcol2:
        if st.button("Duplicate line", disabled=not tactic_labels):
            idx = int(dup_pick.split(":")[0])
            st.session_state["media_plan_rows"] = new_rows + [dict(new_rows[idx])]
            st.session_state["media_plan_dirty"] = new_dirty + [True]  # a duplicate is immediately customizable
            st.session_state["media_plan_version"] += 1
            st.rerun()

    markup = 1.15 if agency_involved else 1.0
    preview_rows = []
    monthly_total_impressions = 0.0
    monthly_total_cost = 0.0
    full_flight_total_impressions = 0.0
    full_flight_total_cost = 0.0
    for _, row in edited_plan_df.iterrows():
        if not str(row.get("Tactic", "")).strip():
            continue
        is_flat_fee = str(row.get("Type", ROW_TYPE_RATE)) == ROW_TYPE_FLAT_FEE
        if is_flat_fee:
            # A flat fee is a one-time full-flight cost, not a per-month rate
            # -- it must NOT scale with month count the way rate-based rows
            # do, so it's excluded from the monthly_total*n_months shortcut
            # below and instead accumulated directly into the flight total.
            full_flight_impressions = 0.0
            full_flight_cost = float(row.get("Flat Cost") or 0)
            monthly_impressions = 0.0
            monthly_cost = full_flight_cost / n_months
        else:
            entered_impressions = float(row.get("Impressions") or 0)
            cpm = float(row.get("CPM") or 0)
            if breakout_mode.startswith("Full Flight"):
                full_flight_impressions = entered_impressions
                full_flight_cost = (full_flight_impressions / 1000.0) * cpm * markup
                monthly_impressions = full_flight_impressions / n_months
                monthly_cost = full_flight_cost / n_months
            else:
                monthly_impressions = entered_impressions
                monthly_cost = (monthly_impressions / 1000.0) * cpm * markup
                full_flight_impressions = monthly_impressions * n_months
                full_flight_cost = monthly_cost * n_months

        monthly_total_impressions += monthly_impressions
        monthly_total_cost += monthly_cost
        full_flight_total_impressions += full_flight_impressions
        full_flight_total_cost += full_flight_cost
        preview_rows.append({
            "tactic": str(row["Tactic"]), "flight": str(row.get("Flight", "")) or flight_label,
            "geo": str(row.get("Geo", "")), "targeting": str(row.get("Targeting", "")),
            "monthly_impressions": monthly_impressions, "monthly_cost": monthly_cost,
            "full_flight_impressions": full_flight_impressions, "full_flight_cost": full_flight_cost,
            "is_flat_fee": is_flat_fee,
        })

    preview_display = pd.DataFrame([
        {
            "tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
            "monthly impressions": "--" if r["is_flat_fee"] else f"{int(r['monthly_impressions']):,}",
            "monthly cost": f"${r['monthly_cost']:,.0f}",
            "full flight impressions": "--" if r["is_flat_fee"] else f"{int(r['full_flight_impressions']):,}",
            "full flight cost": f"${r['full_flight_cost']:,.0f}",
        }
        for r in preview_rows
    ]) if preview_rows else pd.DataFrame(columns=[
        "tactic", "flight", "geo", "targeting", "monthly impressions", "monthly cost",
        "full flight impressions", "full flight cost"])
    st.dataframe(preview_display, use_container_width=True)

    gross_suffix = " gross" if agency_involved else ""
    st.caption(f"Monthly totals: {int(monthly_total_impressions):,} impressions / ${monthly_total_cost:,.0f}{gross_suffix}")
    st.caption(f"**Full Flight Total ({n_months} month{'s' if n_months != 1 else ''}): "
               f"{int(full_flight_total_impressions):,} impressions / ${full_flight_total_cost:,.0f}{gross_suffix}**")

    included_list = build_included_list(
        {"sales_attribution": sales_attribution, "brand_lift": brand_lift},
        commercial_production,
    )
    st.caption("Included with Campaign: " + ", ".join(included_list))

    # ---------------- Generate ----------------
    st.header("Generate")
    proposal_title = st.text_input("Proposal title (appears on cover + media plan)",
                                    value="Total TV Strategy" if total_tv else "CTV Strategy")

    if st.button("Generate proposal", type="primary"):
        selections = {
            "preset": preset_key,
            "market": market_choice,
            "vertical": vertical_key if include_vertical_slides else "none",
            "agency_involved": agency_involved,
            "spanish_campaign": spanish_campaign,
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

        media_plan_rows_final = [
            {"tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
             "impressions": "--" if r["is_flat_fee"] else f"{int(r['monthly_impressions']):,}",
             "cost": f"${r['monthly_cost']:,.0f}" + (" (Gross)" if agency_involved else "")}
            for r in preview_rows
        ] or [{"tactic": "", "flight": flight_label, "geo": market_label, "targeting": "", "impressions": "0", "cost": "$0"}]

        full_flight_total = None
        if n_months > 1 and preview_rows:
            full_flight_total = {
                "label": f"Full Flight Total ({n_months} months)",
                "impressions": f"{int(full_flight_total_impressions):,}",
                "cost": f"${full_flight_total_cost:,.0f}" + (" (Gross)" if agency_involved else ""),
            }

        fill_data = {
            "client_name": client_name or "Client",
            "proposal_title": proposal_title,
            "logo_path": io.BytesIO(logo_file.getvalue()) if logo_file else "placeholder_logo.png",
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
            "media_plan": {
                "plan_title": proposal_title,
                "rows": media_plan_rows_final,
                "totals_label": "Monthly Totals",
                "total_impressions": f"{int(monthly_total_impressions):,}",
                "total_cost": f"${monthly_total_cost:,.0f}" + (" (Gross)" if agency_involved else ""),
                "full_flight_total": full_flight_total,
                "included_list": included_list,
            },
        }

        with st.spinner("Assembling deck..."):
            try:
                prs, original_count, kept_count = assembly.build_presentation(MASTER_DECK_PATH, selections)
                assembly.personalize(prs, fill_data)
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)
            except Exception as exc:
                st.error(f"Assembly failed: {exc}")
                raise

        st.success(f"Assembled {kept_count} of {original_count} slides.")
        st.download_button(
            "Download .pptx",
            data=buffer,
            file_name=f"{(client_name or 'client').replace(' ', '_')}_proposal.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


main()
