"""
app.py -- Streamlit form for the Premion Proposal Builder (build spec
section 5), wired directly to assembly.py's selection -> deletion -> fill
pipeline.

Persistent data (master deck versions, products/rates, the audience catalog,
proposal history) lives in Supabase, reached only through db.py, which falls
back to the local file / hardcoded copies below whenever it's unreachable.
"""

import contextlib
import io
import sys
import os
import json
import re
import subprocess
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import anthropic
import pandas as pd
import streamlit as st
from pptx import Presentation

import assembly
import audience_evidence
import audience_usage_import
import avails_pdf_import
import text_metrics
import db
import geo_resolver
import market_lookup
import market_profiles
import slide_map
import targeting_groups as tg
import targeting_map
import wideorbit
from audience_catalog import (CATEGORY_DESCRIPTIONS, all_categories, catalog_warning,
                              category_matches, clear_catalog_cache, load_audience_catalog,
                              validate_segments)

st.set_page_config(page_title="Premion Proposal Builder", layout="wide")


def _read_build_stamp():
    """Short git SHA + commit time of the code THIS process is running, read
    once at import. Answers "am I running current code" at a glance --
    three separate live investigations this week each ended at "probably a
    stale process," and each cost more than this check would have. A warm
    process that predates a fix keeps reporting the SHA it started with,
    which is exactly the tell that's needed; restarting the process is what
    changes it. Never raises -- falls back to a plain label if git isn't on
    PATH or this checkout has no history, same fallback discipline as every
    other loader in this app.
    """
    try:
        repo_dir = Path(__file__).resolve().parent
        sha = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"], cwd=repo_dir,
            capture_output=True, text=True, timeout=5, check=True,
        ).stdout.strip()
        commit_time = subprocess.run(
            ["git", "log", "-1", "--format=%cI"], cwd=repo_dir,
            capture_output=True, text=True, timeout=5, check=True,
        ).stdout.strip()
        when = datetime.fromisoformat(commit_time).strftime("%Y-%m-%d %H:%M")
        return f"{sha} · {when}"
    except Exception:                                             # noqa: BLE001
        return "unknown build"


BUILD_STAMP = _read_build_stamp()

ANTHROPIC_MODEL = "claude-sonnet-4-6"

# Sized against the LARGEST realistic draft, not the typical one. The old
# ceiling was 2000, set when a draft was a handful of lines and one review
# list -- below what a real proposal now needs, so the model was being cut
# off mid-JSON. Measured worst case: two options of six lines each, three
# audiences carrying avails, full Campaign Specs, and both review lists at
# their 8-item cap comes to ~12,400 characters of pretty-printed JSON, or
# roughly 3,500 output tokens. 16000 leaves ~4.5x headroom on that.
#
# It is also the ceiling for a NON-STREAMING request: past roughly this
# size the SDK starts refusing non-streaming calls it estimates will exceed
# the HTTP timeout. Raising this further means switching these calls to
# client.messages.stream() + get_final_message(), not just editing the
# number. Sonnet 4.6 itself allows up to 128K output.
ANTHROPIC_MAX_TOKENS = 16000

# Where a record of every Claude call goes. The draft path failed live with
# "Claude's response wasn't valid JSON even after stripping markdown fences:
# Expecting value: line 1 column 1 (char 0)" -- char 0 means the text was
# EMPTY, so the fence-stripping the message blamed was never the problem and
# the message pointed at the wrong thing. Nothing was recorded, so a failure
# that cost a live API call and a long wait told us nothing at all. Every
# call now leaves a line behind whether it worked or not.
CLAUDE_LOG_PATH = Path(tempfile.gettempdir()) / "proposal_builder_claude_calls.log"

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
def load_market_profiles():
    """(rows, warning) -- every selectable target DMA, in Nielsen rank order.

    The fallback is unusually complete: market_profiles.build_rows() derives
    the whole set from a file that IS in the repo, so an unreachable Supabase
    costs the profile IMAGES and nothing else -- every market is still
    offered and a proposal can still record the one it targets. Both paths go
    through the same sort_rows, so a rep sees one order either way.
    """
    rows, warning = db.fetch_market_profiles()
    if rows is None:
        return (market_profiles.sort_rows(market_profiles.build_rows()),
                f"{warning}. Using the built-in market list -- "
                "profile slides won't be available.")
    return market_profiles.sort_rows(rows), None


# The picker's "no target market" option. A proposal is allowed not to have
# one -- every proposal logged before this existed doesn't -- so nothing is
# selected by default and no market is assumed on the seller's behalf.


def market_profile_option_label(row):
    """How one market reads in the picker.

    A market whose profile slide was never authored says so HERE, at the
    moment of choosing, rather than the rep picking it, generating, and
    working out from an absence why no profile appeared. Five markets are in
    this state (Honolulu, Palm Springs, Anchorage, Fairbanks, Juneau) and
    they stay selectable: they're real markets, and Premion not having drawn
    a slide is not a reason the app can't sell into them.
    """
    label = row.get("label") or row.get("dma") or row.get("key")
    return label if row.get("image_path") else f"{label}  (no profile slide)"


def target_market_labels(target_dmas, profiles):
    """Display labels for the selected target markets, in picker order."""
    by_key = {row.get("key"): row for row in (profiles or [])}
    return [by_key[key].get("label") for key in (target_dmas or []) if key in by_key]


def geography_default_text(target_labels, originating_label):
    """What the Campaign Specs Geography field defaults to.

    One market per line, because that field is a bullet list on the slide.
    With no target markets selected it falls back to the originating market
    label, which is exactly what every proposal built before target markets
    existed shows -- so nothing changes for them.
    """
    return "\n".join(target_labels) if target_labels else (originating_label or "")


def geo_column_default(target_labels, geography_text, originating_label):
    """What a media plan row's Geo cell defaults to.

    Comma-joined on one line: it's a table cell, not a bullet list.

    The Geography field still governs, which keeps the rep's own words in
    charge: it is auto-filled FROM the target markets (see
    apply_geography_autofill), so in the ordinary case the two say the same
    thing -- and when a rep overrides it with "Denver metro only", the plan
    follows them rather than the raw market list. Target labels are the
    fallback for the case where Geography is somehow empty.

    What is gone is `first_line()`. That was the single-market assumption
    itself: against a three-market Geography it kept "Denver" and silently
    dropped Atlanta and Phoenix -- the same class of truncation that once put
    one audience attribute on a plan beside a specs slide listing four. The
    lines are joined instead. A rep who writes prose across several lines
    gets a longer cell rather than a quietly truncated one, which is the
    right way round: an over-full cell is visible and editable, a missing
    market looks deliberate.

    This is only the DEFAULT. The rep edits any row's Geo directly and the
    dirty-row rule then protects it: resolve_row_defaults only re-seeds rows
    nobody has touched. And the imported broadcast row is held regardless --
    its Geo comes from the station call sign, so a DC-sold proposal targeting
    Denver correctly reads broadcast in Washington DC DMA and streaming in
    Denver on the same plan.
    """
    lines = [line.strip() for line in str(geography_text or "").splitlines()
             if line.strip()]
    if lines:
        return ", ".join(lines)
    if target_labels:
        return ", ".join(target_labels)
    return originating_label or ""


def avails_rows_for_markets(target_labels, geo_default, combine=False):
    """The avails rows a target-market selection seeds.

    One row PER MARKET is the default, because that is how these proposals are
    actually built: a rep pulls avails for each market separately and the
    board wants to see them side by side. Combining several markets into one
    line is real but less common, which is what `combine` is for.

    The market rows REPLACE the originating-market default rather than sitting
    beside it. A leftover "Washington, DC DMA" row under three target markets
    isn't a starting point, it's a row someone has to notice and delete -- and
    a zero-avails row that reaches the targeting slide reads as a market with
    no inventory.

    With no markets selected this is exactly the single blank row the table
    has always started with, so nothing changes for a proposal that names no
    target market.
    """
    if not target_labels:
        return [{"Audience": "", "Geo": geo_default, AVAILS_COLUMN_MONTHLY: 0}]
    if combine:
        return [{"Audience": "", "Geo": ", ".join(target_labels),
                 AVAILS_COLUMN_MONTHLY: 0}]
    return [{"Audience": "", "Geo": label, AVAILS_COLUMN_MONTHLY: 0}
            for label in target_labels]


def plan_lines_from_avails(avails_table, fallback_geo, fallback_audience=""):
    """(audience, geo) for every avails row -- one plan line each, always.

    The hierarchy is AUDIENCE over GEO. Homeowners in Denver and
    in-market-for-windows in Denver are two lines, not one: they carry
    separate avails, and they can carry separate budgets and separate
    negotiated CPMs. Collapsing rows that share a market would make those
    inexpressible, and would silently discard an avails figure the rep went
    and pulled.

    So nothing is deduplicated. Two audiences across three markets give six
    avails rows and six plan lines.

    Ordered AUDIENCE-MAJOR -- audience 1 across every geo, then audience 2
    across every geo -- because that is how a plan reads to a client: one
    audience's buy laid out across its markets, then the next. Both audience
    and geo keep first-appearance order within that, so the plan tracks the
    order the rep built the table in rather than an alphabetical one nobody
    chose.

    Note a combined AND audience ("homeowners with income $50K+") is ONE
    audience here and therefore one line. It is a single audience defined
    more precisely, not two audiences sharing a geo -- which is exactly the
    distinction the audience builder draws when it offers AND against
    "add as a separate group".
    """
    pairs = []
    for row in avails_table or []:
        geo = str(row.get("Geo", "") or "").strip()
        audience = str(row.get("Audience", "") or "").strip()
        if not geo and not audience:
            continue
        pairs.append((audience, geo or fallback_geo))
    if not pairs:
        return [(fallback_audience, fallback_geo)]

    audience_order, geo_order = {}, {}
    for audience, geo in pairs:
        audience_order.setdefault(audience, len(audience_order))
        geo_order.setdefault(geo, len(geo_order))
    return sorted(pairs, key=lambda pair: (audience_order[pair[0]],
                                           geo_order[pair[1]]))


def plan_lines_from_groups(groups, fallback_geo, fallback_audience=""):
    """(audience, geo, group_id) for every targeting group -- one plan line
    each, always, audience-major -- the group-aware sibling of
    `plan_lines_from_avails`, which stays as the flat-row entry point for
    everything that doesn't know about groups yet.

    Same ordering algorithm, and for the same reason: nothing here assumes
    the group LIST already arrived in audience-major order. It usually will
    (the audience builder creates groups in that order), but a migrated
    proposal's groups come from whatever order its flat avails table
    happened to be in, so this re-derives the order from first appearance
    rather than trusting the list.
    """
    triples = []
    for group in (groups or []):
        audience = tg.audience_label(group)
        geo = tg.geo_label(group, label_for=_market_display_name).strip() or fallback_geo
        # A group with no real audience (a market-only row from before any
        # audience was typed) carries no group_id here on purpose. Its line
        # has no audience identity to hold onto -- resolve_row_defaults must
        # keep tracking the CURRENT Campaign Specs default forever, exactly
        # as a blank-Audience flat row always has (valid_audiences below
        # excludes "" for the same reason). Stamping a stable id on it would
        # freeze its Targeting at whatever it happened to be on first seed.
        group_id = group.get("id") if audience else None
        triples.append((audience, geo, group_id))
    if not triples:
        return [(fallback_audience, fallback_geo, None)]

    audience_order, geo_order = {}, {}
    for audience, geo, _group_id in triples:
        audience_order.setdefault(audience, len(audience_order))
        geo_order.setdefault(geo, len(geo_order))
    return sorted(triples, key=lambda t: (audience_order[t[0]], geo_order[t[1]]))


def group_ids_of(row):
    """The group ids a plan row carries, or [] when it carries none -- a
    legacy row, a broadcast row, or one seeded before targeting groups
    existed. `isinstance(v, list)`, never truthiness or `pd.isna()`: a row
    missing the key entirely, or one whose value came back through a
    data_editor as a bare float NaN, must read as "no ids", and `pd.isna()`
    raises on an actual list rather than reporting False."""
    value = row.get("_group_ids")
    return list(value) if isinstance(value, list) else []


def apply_avails_autofill(rows):
    """Seed the avails rows, and keep them in step until the rep edits them.

    Returns True when it actually changed something, so the caller can bump
    the editor version -- a data_editor handed new rows under the same key
    keeps showing the old ones.

    Same clean/dirty discipline as apply_geography_autofill and as the media
    plan grid: the last value THIS function wrote is remembered, and the rows
    are replaced only while they still match it. The moment the rep edits a
    row -- or a draft or a loaded proposal writes real avails -- the two stop
    matching and this never touches them again. That matters more here than
    for Geography, because an avails figure is something a rep went and looked
    up, and silently replacing it with a zero would be worse than useless.
    """
    current = st.session_state.get("avails_seed_rows")
    applied = st.session_state.get("_avails_autofill")
    if current is not None and current != applied:
        return False                      # edited, drafted or loaded -- leave it
    if current == rows:
        return False                      # already right; don't churn the editor
    st.session_state["avails_seed_rows"] = [dict(r) for r in rows]
    st.session_state["_avails_autofill"] = [dict(r) for r in rows]
    return True


def sync_targeting_groups():
    """Keep `targeting_groups` and `avails_seed_rows` in agreement, and
    decide which side moved since this last ran.

    `targeting_groups` is the source of truth once anything here has written
    it, but four things still write `avails_seed_rows` directly and know
    nothing about groups: a draft, a rehydrated proposal, the Audience
    finder's Add, and every existing avails test -- none of those are being
    taught about groups in this phase. So the same clean/dirty marker
    `apply_avails_autofill` uses decides which side changed: if the flat rows
    differ from the projection THIS function last wrote, something upstream
    of groups wrote them, and groups are re-derived from the rows (keeping
    the id of any group whose projected row is unchanged, via
    `seed_rows_to_groups`'s own `existing=` matching). Otherwise the rows are
    re-projected from the groups, so an edit made through a future
    groups-aware UI still reaches everything that reads the flat shape.

    Call before anything needs `targeting_groups` to be current. In this
    phase that's just ahead of building `form_json`, so a saved proposal's
    new `targeting_groups` key agrees with the `avails_rows` key it's already
    saving. A later phase that renders a groups-aware UI will need this
    called earlier too -- before that UI reads groups, and before Section A's
    market picker once a group can autofill it.
    """
    stored_rows = st.session_state.get("avails_seed_rows")
    groups = st.session_state.get("targeting_groups")
    applied_rows = st.session_state.get("_groups_rows_applied")

    if stored_rows is not None and stored_rows != applied_rows:
        new_groups = tg.seed_rows_to_groups(stored_rows, AVAILS_COLUMN_MONTHLY, existing=groups)
        st.session_state["targeting_groups"] = new_groups
        st.session_state["_groups_rows_applied"] = [dict(r) for r in stored_rows]
        return

    if groups is not None:
        projected = tg.groups_to_seed_rows(groups, AVAILS_COLUMN_MONTHLY)
        if projected != stored_rows:
            st.session_state["avails_seed_rows"] = projected
            st.session_state["_groups_rows_applied"] = [dict(r) for r in projected]


def apply_geography_autofill(default_text):
    """Keep Geography in step with the target markets until the rep edits it.

    `geography_text` is a keyed widget, so its `value=` argument is ignored
    the moment session_state holds anything -- which means a market picked
    after the first render would never reach the field. Writing it directly
    is the only way, and writing it unconditionally would erase whatever the
    rep typed.

    So the last value THIS function applied is remembered, and the field is
    replaced only while it still matches -- the same "clean vs dirty"
    distinction the media plan grid draws per row, applied to one text area.
    Once the rep edits it, or a draft or a loaded proposal writes real content
    into it, the two stop matching and this never touches it again.

    Must run before the widget is instantiated: Streamlit raises on writing a
    widget's key afterwards.
    """
    current = st.session_state.get("geography_text")
    applied = st.session_state.get("_geography_autofill")
    if current is None or not str(current).strip() or current == applied:
        st.session_state["geography_text"] = default_text
        st.session_state["_geography_autofill"] = default_text


def target_dma_list(selections, row=None):
    """The target DMA keys a stored proposal carries, in order.

    ONE place decides this, because three callers need the answer -- the
    rehydration path, rebuild-as-presented, and the generate path reading its
    own selections back -- and a proposal that rebuilt with different markets
    than it loaded with would be a different deck from the one the client got.

    Reads the list shape first and falls back to the pre-multi-select scalar,
    so a proposal logged before this shipped still loads with its one market
    rather than none. Order is preserved as stored: it is what drives slide
    order, not a presentation detail.
    """
    for source in (selections or {}, row or {}):
        values = source.get("target_dmas")
        if values:
            return [v for v in values if v]
        single = source.get("target_dma")
        if single:
            return [single]
    return []


@st.cache_resource
def install_market_lookup():
    """Register the county->DMA table with geo_resolver, once per process.
    Returns (lookup, warning) -- the same convention every other loader in
    this app follows.

    The first thing in the whole app that actually calls
    `market_lookup.install()` -- `geo_resolver` ships with no lookup
    registered until something does, on purpose (county->DMA is Nielsen's
    intellectual property, see market_lookup.py), so this is the one place
    that opts the app in. `@st.cache_resource` makes it idempotent across
    reruns within a process without re-reading and re-registering the table
    on every script execution.

    **Surfaced now, not silent.** A missing/unreadable table used to degrade
    all the way down to `geo_resolver`'s own per-resolve-click note ("No
    market lookup is registered") -- technically honest, but easy to read as
    the resolver doing its job rather than the table never having loaded at
    all, and invisible until a rep happened to click Resolve. A live report
    hit exactly this with every market auto-resolution dead and no clear
    signal why. Every failure mode this function can actually hit is
    distinguished in the warning text on purpose, because "why" has three
    different real answers: the file genuinely isn't in this checkout/deploy
    (`market_lookup.available()` False), it's present but failed to parse
    (`install()` raised -- corruption, a bad gzip member, anything), or it
    loaded but came back empty (a build defect, not a delivery one).
    """
    if not market_lookup.available():
        return None, (
            f"{market_lookup.LOOKUP_PATH.name} isn't present in this deployment -- every "
            f"market auto-resolution (radius/counties/zips -> DMA) is unavailable until it "
            f"is. It's a small, committed file (not gitignored); if it's missing here, this "
            f"process is very likely running a stale build -- reboot the app before assuming "
            f"a code problem.")
    try:
        lookup = market_lookup.install()
    except Exception as exc:                                     # noqa: BLE001
        return None, (f"Market lookup failed to load ({db.describe_error(exc)}) -- market "
                      f"auto-resolution is unavailable.")
    if not lookup or len(lookup) == 0:
        return None, "Market lookup loaded but is empty -- market auto-resolution is unavailable."
    return lookup, None


def _market_display_name(market_key):
    """The canonical DMA name for a resolved market key, or the key itself
    when the table isn't available -- never raises, so a group resolved
    while the lookup was installed still displays sensibly if it's ever
    missing later (a different process, a stripped-down environment)."""
    try:
        return market_lookup.market_name(market_key) or market_key
    except market_lookup.MarketLookupUnavailable:
        return market_key


def apply_group_markets_autofill(profiles):
    """Add newly-resolved targeting-group markets to Section A's target DMA
    picker -- MONOTONE ADD-ONLY, deliberately not the replace-while-clean
    idiom `apply_geography_autofill` uses. That idiom is wrong here: it
    would either stop applying the moment a rep removes one autofilled
    market, or fight them by putting it straight back on the next rerun.

    `_group_markets_applied` accumulates every market key this autofill has
    EVER contributed, across the whole session -- it never shrinks, even
    once a market is removed from the picker afterward. Each run adds only
    the DIFFERENCE between what's resolved across every group right now and
    what has already been contributed, so a rep's removal is never
    re-applied -- not on the next run, and not by a LATER, different group
    resolving to that same market.

    Must run before `market_profile_picker` renders: Streamlit raises on
    writing a widget's state after it's instantiated, and this writes
    `target_dma_choice` directly.
    """
    groups = st.session_state.get("targeting_groups") or []
    resolved_now = {m for g in groups for m in (g.get("resolved_markets") or [])}
    already_applied = st.session_state.get("_group_markets_applied") or set()
    new_markets = resolved_now - already_applied
    st.session_state["_group_markets_applied"] = already_applied | resolved_now
    if not new_markets:
        return False

    by_key = {r.get("key"): r for r in (profiles or [])}
    new_labels = [market_profile_option_label(by_key[key])
                 for key in sorted(new_markets) if key in by_key]
    current = list(st.session_state.get("target_dma_choice") or [])
    added = [label for label in new_labels if label not in current]
    if not added:
        return False
    st.session_state["target_dma_choice"] = current + added
    return True


GEO_MODE_MARKETS = "Markets"
GEO_MODE_COUNTIES = "Counties"
GEO_MODE_ZIPS = "Zips"
GEO_MODE_RADIUS = "Radius"
GEO_MODES = [GEO_MODE_MARKETS, GEO_MODE_COUNTIES, GEO_MODE_ZIPS, GEO_MODE_RADIUS]


def parse_radius_centers(text, default_miles):
    """[(center, effective_miles), ...] from a textarea, one center per
    line -- a zip or a street address, blank lines dropped.

    A line may end in ", <number>" to override the default radius for that
    line alone ("1100 Wilson Blvd, Arlington VA, 25"). Recognized only when
    the text after the LAST comma parses as a plain number -- an ordinary
    address that merely contains a comma ("1100 Wilson Blvd, Arlington VA")
    is left whole and uses the default, which is what a rep typing it
    expects. `rpartition` rather than `split` for the same reason: only the
    final comma can possibly be an override, an address may have several.
    """
    entries = []
    for raw_line in str(text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        head, sep, tail = line.rpartition(",")
        miles = default_miles
        if sep and head.strip():
            try:
                miles = float(tail.strip())
                line = head.strip()
            except ValueError:
                pass
        entries.append((line, miles))
    return entries


def resolve_group_geography(mode, markets_picked=None, counties_text="", zips_text="",
                            radius_centers_text="", radius_miles=""):
    """(geo_def, resolved_zips, resolved_markets, notes, unresolved) for one
    Resolve click on a group's geo-definition expander -- the ONE place that
    decides what each of the four modes means, so app.py's UI code never
    calls geo_resolver directly and every mode's notes/unresolved entries go
    through the same reporting shape.

    Markets needs no resolver call at all: a directly-picked market IS its
    own resolution. The other three each call straight into the existing
    geo_resolver functions and finish with zips_to_markets for
    resolved_markets, per geo_targeting_roadmap.md D's Phase 6 plan --
    Counties through counties_to_fips + counties_to_zips, Zips through
    parse_zip_list, Radius through radius_to_zips. Every `Resolution.notes`
    and unresolved entry from every step is concatenated and returned, never
    dropped -- that shape exists precisely so nothing is lost silently.

    Radius resolves EACH center with its own `radius_to_zips` call (a list
    of one), not one batched call across all of them -- the network cost is
    identical either way (radius_to_zips geocodes one center at a time
    internally regardless of how many are in the list), and calling it per
    center is what makes a per-center report possible at all: how many zips
    that center alone contributed, or that it specifically is the one that
    didn't geocode. The union is deduplicated by construction (`set`), so
    two stores' overlapping circles never double-count a zip. A center
    override survives in `geo_def["centers"]` as `{"center", "miles"}"`;
    an un-overridden one stays a bare string, which is also exactly what
    every geo_def stored before this feature already looks like -- one
    center, one string, the group's own `miles` -- so an old proposal's
    radius group is unaffected.
    """
    if mode == GEO_MODE_MARKETS:
        keys = [str(k).strip() for k in (markets_picked or []) if str(k).strip()]
        return {"kind": "markets", "markets": keys}, [], keys, [], []

    notes, unresolved = [], []
    if mode == GEO_MODE_COUNTIES:
        fips_result = geo_resolver.counties_to_fips(counties_text)
        zips_result = geo_resolver.counties_to_zips(list(fips_result.resolved.values()))
        zips = sorted({z for group in zips_result.resolved.values() for z in group})
        geo_def = {"kind": "counties", "counties": list(fips_result.resolved.keys())}
        notes = list(fips_result.notes) + list(zips_result.notes)
        unresolved = list(fips_result.unresolved) + [str(u) for u in zips_result.unresolved]
    elif mode == GEO_MODE_ZIPS:
        zips = geo_resolver.parse_zip_list(zips_text)
        geo_def = {"kind": "zips", "zips": zips}
    elif mode == GEO_MODE_RADIUS:
        try:
            default_miles = float(radius_miles)
        except (TypeError, ValueError):
            return None, [], [], ["The default radius must be a number of miles."], [str(radius_miles)]
        entries = parse_radius_centers(radius_centers_text, default_miles)
        if not entries:
            return None, [], [], ["Enter at least one center (a zip or an address), one per line."], []

        union_zips = set()
        per_center_notes = []
        for center, miles in entries:
            result = geo_resolver.radius_to_zips([center], miles)
            if result.unresolved:
                unresolved.append(center)
                per_center_notes.append(f"{center}: couldn't be located.")
                continue
            union_zips |= set(result.resolved)
            override = f" (custom {miles:g}mi)" if miles != default_miles else ""
            per_center_notes.append(
                f"{center}{override}: {len(result.resolved)} zip(s) within {miles:g}mi.")
        zips = sorted(union_zips)

        custom_count = sum(1 for _, m in entries if m != default_miles)
        notes = [f"{len(entries)} location(s), default {default_miles:g}mi"
                + (f" ({custom_count} with a custom radius)" if custom_count else "")
                + f", {len(zips)} unique zip(s)."] + per_center_notes

        centers_field = [center if miles == default_miles else {"center": center, "miles": miles}
                         for center, miles in entries]
        geo_def = {"kind": "radius", "centers": centers_field, "miles": default_miles}
    else:
        return None, [], [], [f"Unknown geo mode {mode!r}."], []

    market_result = geo_resolver.zips_to_markets(zips)
    notes = notes + list(market_result.notes)
    unresolved = unresolved + list(market_result.unresolved)
    resolved_markets = sorted(market_result.resolved.keys())
    return geo_def, zips, resolved_markets, notes, unresolved


def render_group_geo_expander(group):
    """One targeting group's geo-definition panel: pick a mode, resolve it
    through `resolve_group_geography`, and write the result straight onto
    THIS group in `targeting_groups` -- reassigned, never mutated in place,
    the same discipline `_add_segment_to_group` and `merge_plan_rows` use.
    Never touches any other group.

    Separate from the D2 grid's own Markets cell (Phase 4), which stays the
    quick, direct way to pick markets by hand -- this is for when a rep has
    geography in counties, zips or a radius instead, and needs it resolved
    to markets rather than typed as one.
    """
    gid = group["id"]
    label = tg.audience_label(group) or "(untitled)"
    current_markets = group.get("resolved_markets") or []
    summary = tg.geo_label(group, label_for=_market_display_name)
    with st.expander(f"📍 Geography: {label}", expanded=False):
        if summary:
            st.caption(f"Currently: {summary}"
                       + (f" -- {len(current_markets)} market(s) resolved"
                          if current_markets else " -- not yet resolved to markets"))
        # A rep's own label always wins in tg.geo_label -- the plan table's
        # Geo cell, the targeting slide, and every other reader of it. Plain
        # widget state, not the D2 grid's fold-back dance: unlike a
        # data_editor cell, this key's own persisted value already IS the
        # source of truth, so it just needs writing onto the group whenever
        # it changes, same reassign-not-mutate discipline as Resolve below.
        #
        # The D2 grid's own Label cell (Phase 4-adjacent) writes this SAME
        # `name` field. `group_name_generation` is part of the key for the
        # same reason `option_name_{gen}_{idx}` carries one: once
        # `geo_name_{gid}` exists in session_state, `value=` is ignored on
        # every later rerun, so a grid edit landing in `group["name"]`
        # without a fresh key would be invisible here -- and this widget's
        # own stale cached text would then overwrite the grid's edit right
        # back, via the `if name != ...` write-back below.
        name_gen = st.session_state.get("group_name_generation", 0)
        name = st.text_input(
            "Label (optional)", value=group.get("name") or "", key=f"geo_name_{gid}_{name_gen}",
            placeholder="e.g. Philly Zip Add-On",
            help="Shown on the plan table's Geo column and the targeting slide "
                 "instead of a derived summary. A Zips or Radius group with many "
                 "entries especially benefits -- the raw list never belongs on a "
                 "client-facing table regardless of whether it fits.")
        if name != (group.get("name") or ""):
            groups = [dict(g) for g in (st.session_state.get("targeting_groups") or [])]
            for g in groups:
                if g["id"] == gid:
                    g["name"] = name
                    break
            st.session_state["targeting_groups"] = groups
        mode = st.radio("Mode", GEO_MODES, horizontal=True, key=f"geo_mode_{gid}",
                        label_visibility="collapsed")

        markets_picked, counties_text, zips_text, radius_centers_text, radius_miles = None, "", "", "", ""
        if mode == GEO_MODE_MARKETS:
            catalog = market_lookup.load().get("markets", {}) if market_lookup.available() else {}
            name_to_key = {entry.get("name", key): key for key, entry in catalog.items()}
            picked_names = st.multiselect(
                "Markets", sorted(name_to_key), key=f"geo_markets_{gid}",
                help="Picked directly -- no resolution needed, this IS the group's geography.")
            markets_picked = [name_to_key[n] for n in picked_names if n in name_to_key]
        elif mode == GEO_MODE_COUNTIES:
            counties_text = st.text_area(
                "Counties", key=f"geo_counties_{gid}", height=70,
                placeholder="Somerset NJ; Bucks PA; New Castle DE",
                help="NAME STATE, semicolon or newline separated -- the same way the avails "
                     "documents write them.")
        elif mode == GEO_MODE_ZIPS:
            zips_text = st.text_area(
                "Zips", key=f"geo_zips_{gid}", height=70,
                placeholder="20005, 20006, 20007",
                help="Paste a zip list, however it's separated.")
        else:
            rcol1, rcol2 = st.columns([3, 1])
            with rcol1:
                radius_centers_text = st.text_area(
                    "Centers (one per line -- zip or street address)",
                    key=f"geo_radius_centers_{gid}", height=100,
                    placeholder="20005\n1100 Wilson Blvd, Arlington VA\n"
                                "1100 Wilson Blvd, Arlington VA, 25",
                    help="One location per line. A line ending in \", <number>\" uses that "
                         "many miles for that line alone instead of the default -- "
                         "\"1100 Wilson Blvd, Arlington VA, 25\" resolves that one store at "
                         "25 miles even if every other line is using 10.")
            with rcol2:
                radius_miles = st.text_input(
                    "Default miles", key=f"geo_radius_miles_{gid}", placeholder="25")

        pending = st.session_state.pop(f"_geo_result_{gid}", None)
        if pending:
            geo_notes, geo_unresolved = pending
            for note in geo_notes:
                st.caption(f"ℹ️ {note}")
            # A well-formed zip that just has no county/market on file (point
            # and PO-box zips) is expected and already explained calmly by
            # geo_notes above -- it stays a real, targetable zip. The yellow
            # warning is reserved for input that isn't even a five-digit
            # number (a bad address, an unrecognized county name, a typo),
            # which is the genuinely actionable case.
            malformed = [u for u in geo_unresolved if geo_resolver.normalize_zip(u) is None]
            if malformed:
                st.warning(f"{len(malformed)} entr{'y' if len(malformed) == 1 else 'ies'} "
                           f"couldn't be resolved: {', '.join(str(u) for u in malformed[:10])}"
                           f"{'...' if len(malformed) > 10 else ''}")

        if st.button("Resolve", key=f"geo_resolve_btn_{gid}"):
            # Radius is the one mode that can hit a real network geocoder,
            # once per address line, sequentially -- a bare zip resolves
            # locally and instantly, but a page of street addresses is a
            # page of blocking HTTP calls. The spinner is the honest signal
            # that this click can take several seconds with 15-20 lines of
            # addresses; it is NOT a promise to parallelize those calls into
            # the Census Geocoder's free, keyless, unrate-limited-in-name-
            # only API -- see parse_radius_centers/resolve_group_geography.
            spinner = st.spinner("Resolving...") if mode == GEO_MODE_RADIUS else contextlib.nullcontext()
            with spinner:
                geo_def, resolved_zips, resolved_markets, geo_notes, geo_unresolved = resolve_group_geography(
                    mode, markets_picked=markets_picked, counties_text=counties_text,
                    zips_text=zips_text, radius_centers_text=radius_centers_text,
                    radius_miles=radius_miles)
            if geo_def is not None:
                groups = [dict(g) for g in (st.session_state.get("targeting_groups") or [])]
                for g in groups:
                    if g["id"] == gid:
                        g["geo_def"] = geo_def
                        g["resolved_zips"] = resolved_zips
                        g["resolved_markets"] = resolved_markets
                        break
                st.session_state["targeting_groups"] = groups
                st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
            st.session_state[f"_geo_result_{gid}"] = (geo_notes, geo_unresolved)
            st.rerun()

        if group.get("resolved_zips") and st.button(
                "🗺️ View on map / export zips", key=f"geo_goto_map_{gid}"):
            st.session_state["goto_zip_map_builder"] = True
            st.rerun()


def market_profile_picker(profiles, warning):
    """Target DMAs + whether to include their profile slides.

    Returns (target_keys, include_slides) -- a LIST, in the order the rep
    selected them, which is the order the profile slides appear in the deck.
    Reps routinely run one campaign across several markets, and the deck was
    built for it: the national roll-up slide is replaced by one profile slide
    per selected market.

    Independent of the originating market above it, which drives station
    branding, the Total TV variants and the broadcast DMA -- a DC-originated
    proposal can target any set of DMAs in the country.
    """
    if warning:
        st.warning(warning)

    by_label = {market_profile_option_label(r): r for r in profiles}
    chosen_labels = st.multiselect(
        "Target DMAs", list(by_label),
        key="target_dma_choice", on_change=_clear_ai_section, args=("basics",),
        placeholder="Search markets...",
        help="The markets this campaign is aimed at -- type to filter. "
             "Separate from the Market above, which is the station the "
             "proposal comes from. Profile slides appear in the order you "
             "pick them.")

    rows = [by_label[label] for label in chosen_labels if label in by_label]
    if not rows:
        return [], False

    with_slide = [r for r in rows if r.get("image_path")]
    without = [r for r in rows if not r.get("image_path")]

    # The toggle is NOT dropped when nothing selected has a slide -- it stays
    # put, disabled, next to the reason. A control that silently disappears
    # reads as the feature being broken, and leaves the rep guessing which of
    # the two it was.
    include = st.checkbox(
        "Include the market viewer profile slides", value=bool(with_slide),
        disabled=not with_slide, key="include_market_profile",
        help=None if with_slide else
        "Premion's profile deck has no slide for any of the markets selected, "
        "so there's nothing to include. They're still valid targets.")

    # How many slides this actually adds, said plainly. A rep selecting a
    # dozen markets is making a deliberate choice and it isn't the app's
    # place to veto it -- but they should be able to see the deck growing
    # without counting the chips themselves.
    if include and with_slide:
        st.caption(f"{len(with_slide)} market profile slide"
                   f"{'s' if len(with_slide) != 1 else ''} will replace the "
                   f"national Total U.S. slide.")
    if without:
        names = ", ".join(r.get("label") for r in without)
        st.caption(f":grey[No viewer profile slide exists for {names} -- "
                   f"{'they' if len(without) > 1 else 'it'} will still be "
                   f"targeted, just without a profile slide.]")

    return [r.get("key") for r in rows], bool(include and with_slide)


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

AUDIENCE_USAGE_CSV_PATH = Path(__file__).parent / "audience_usage_ytd.csv"


@st.cache_resource(show_spinner=False)
def load_audience_index():
    """The `audience_evidence.AudienceIndex` the booking-evidence panel reads
    from -- Supabase's `audience_usage` table first, the committed
    `audience_usage_ytd.csv` when that's unreachable, same (DB, local
    fallback) convention as `load_rate_card`/`load_audience_catalog`.
    `cache_resource` rather than `cache_data`: the index holds Counters, not
    a DataFrame, and is read-only for the life of the process, so one shared
    object is exactly right, the same choice `install_market_lookup` makes
    for `market_lookup`.
    """
    rows, warning = db.fetch_audience_usage()
    if rows is None:
        try:
            rows = pd.read_csv(AUDIENCE_USAGE_CSV_PATH).to_dict("records")
        except Exception:
            rows = []
    return audience_evidence.build_index(rows)

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

# --- Avails basis ---------------------------------------------------------
# Avails are ALWAYS stored monthly, whatever basis the table is showing.
# The avails system reports monthly numbers, so monthly is the one basis
# that never needed converting to get here, and keeping storage on it means
# a reach percentage, a rebuild and a rehydration all read the same figure.
# The toggle changes what's DISPLAYED (in the form and on the slide) and the
# column header that names it, so a number in front of a client is never
# ambiguous about which basis it is.
AVAILS_BASIS_MONTHLY = "Monthly (default)"
AVAILS_BASIS_FLIGHT = "Full flight"
AVAILS_COLUMN_MONTHLY = "Max Monthly Avails"


def avails_column_label(basis, n_months):
    """The column header, which has to name the basis it's showing."""
    if basis == AVAILS_BASIS_FLIGHT:
        months = max(1, int(n_months or 1))
        return f"Max Avails — Full Flight ({months} month{'s' if months != 1 else ''})"
    return AVAILS_COLUMN_MONTHLY


def avails_to_display(monthly, basis, n_months):
    """Stored monthly value -> what the chosen basis shows."""
    try:
        monthly = int(monthly or 0)
    except (TypeError, ValueError):
        return 0
    if basis == AVAILS_BASIS_FLIGHT:
        return monthly * max(1, int(n_months or 1))
    return monthly


def avails_from_display(shown, basis, n_months):
    """What the user typed in the chosen basis -> the monthly value to store."""
    try:
        shown = int(float(str(shown).replace(",", "") or 0))
    except (TypeError, ValueError):
        return 0
    if basis == AVAILS_BASIS_FLIGHT:
        return int(round(shown / max(1, int(n_months or 1))))
    return shown


def _cell_unchanged(shown_before, shown_after):
    """The one fold-back test every D2 grid column shares: is this cell
    showing exactly what it showed before this run, or did the user really
    edit it.

    A targeting group's underlying structure often can't be recovered from
    what the grid DISPLAYS -- a multi-term audience renders as
    `audience_label`'s plain, comma-joined text, not the form its own parser
    recognizes; a resolved geo_def renders as market names, not the
    counties/zips/radius that produced them; an avails figure renders in
    whichever basis is on screen, not the monthly figure that's stored. Any
    fold-back that re-derives structure from that display text UNCONDITIONALLY
    -- on every rerun, not just an actual edit -- silently destroys the real
    data the moment the grid merely redraws it, with no edit at all. This
    was a real bug, once: Phase 4 collapsed a 2-term AND group into one
    verbatim term this way. Every fold-back in this file now goes through
    this one test, so a fourth field (the map's color hook, e.g.) gets the
    fix by construction instead of needing the hazard rediscovered.

    `shown_before is None` (a brand-new row, no prior group) is never
    "unchanged" -- there is nothing to preserve.
    """
    return shown_before is not None and shown_after == shown_before


def restore_untouched_avails(stored_monthly, shown_before, shown_after, basis, n_months):
    """The monthly value to keep for one row, given what it displayed before
    the user saw it and what it displays now.

    **A row the user did not touch is never round-tripped.** Converting to
    the flight basis and back is monthly -> monthly*N -> monthly*N/N, which
    is only exact while N divides cleanly; flipping the toggle twice on an
    odd number would otherwise walk the stored figure a unit at a time, and
    the drift would land in a number the client is quoted. So a row whose
    displayed value is unchanged (`_cell_unchanged`) keeps its stored value
    byte for byte, and only a row that was actually edited is converted back.

    `shown_before=None` (a brand-new row, no prior group) must reach
    `_cell_unchanged` as None, not as 0 -- collapsing it here would make a
    freshly-typed 0 read as "unchanged" and silently keep it at 0 forever,
    defeating `_cell_unchanged`'s own documented contract that a brand-new
    row is never unchanged.
    """
    try:
        shown_after_num = int(float(str(shown_after).replace(",", "") or 0))
    except (TypeError, ValueError):
        return avails_from_display(shown_after, basis, n_months)
    shown_before_num = None if shown_before is None else int(shown_before or 0)
    if _cell_unchanged(shown_before_num, shown_after_num):
        return int(stored_monthly or 0)
    return avails_from_display(shown_after, basis, n_months)
BREAKOUT_MODES = [BREAKOUT_MONTHLY, BREAKOUT_FULL_FLIGHT]

# The flight the form opens on. Constants rather than literals in the widget
# calls because apply_draft_to_form has to predict the month count main() will
# derive *before* those widgets have rendered -- on the very first run there is
# no session_state to read, so the only honest answer is the same default the
# widget is about to use. Two copies of these dates would put the draft's
# arithmetic and the deck's totals back out of step, which is the bug this
# whole shape exists to prevent.
DEFAULT_FLIGHT_START = date(2026, 9, 1)
DEFAULT_FLIGHT_END = date(2026, 11, 30)

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
  "target_markets": [],
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
  "total_tv": false,
  "audiences": [{"segment": "exact catalog name", "geo": "", "max_avails": 0, "avails_basis": "monthly"}],
  "attribution": ["web", "sales", "brand_lift", "first_party", "linear_reach_ext", "commercial_production"],
  "sports": [],
  "campaign_specs": {
    "goals": [], "audience": [], "geography": [],
    "budget": [], "placements": [], "timing": []
  },
  "unresolved": ["What you assumed, in one sentence. What to confirm, in one sentence."],
  "unresolved_internal": ["Same, for checks the seller does rather than asks the client."]
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
    "flight_start": "flight", "flight_end": "flight", "active_months": "flight",
    "goals_text": "specs", "audience_text": "specs", "geography_text": "specs",
    "budget_text": "specs", "placements_text": "specs", "timing_text": "specs",
    "avails_seed_rows": "avails", "avails_version": "avails",
    # The basis rides with the rows it describes: a re-draft that rewrote the
    # avails but left the basis alone (or the reverse) would put one
    # audience's monthly figure under a full-flight header.
    "avails_basis": "avails",
    # targeting_groups is the projection's source once anything has written
    # it, so it has to be skipped alongside avails_seed_rows or a preserved
    # "avails" section would keep the old rows while a re-draft silently
    # overwrote the groups they're supposed to agree with.
    "targeting_groups": "avails",
    "live_sports_enabled": "products", "selected_sports": "products",
    "include_sport_viewership": "products",
    "plan_options": "media_plan", "media_plan_markup": "media_plan",
    "plan_options_gen": "media_plan",
    # Internal bookkeeping that only means anything alongside the rows it
    # describes -- it has to be skipped with them or it would claim rows that
    # were never written.
    "_product_seed_key": "media_plan", "_shared_fields_key": "media_plan",
}
# ---------------------------------------------------------------------------
# Keeping the Build form alive across page navigation
#
# Streamlit garbage-collects the session_state entry of any *keyed widget*
# that wasn't rendered during a run, so switching to another page destroys the
# whole form. Measured: 111 keys gone -- client name, vertical, market, flight
# dates, every product toggle, every Campaign Specs field -- while the 14
# non-widget keys (plan_options, avails_seed_rows, broadcast_schedule, the
# drafted review lists) survive untouched. That is why the finders were
# unusable mid-proposal: they are the one thing a seller navigates away to
# use, which is the workflow they were built for.
#
# The fix is to mirror widget values into an ordinary session_state entry,
# which is not collected, and put them back before the widgets render again.
# **Restoring only keys that are ABSENT** is what makes it safe: a live value
# is never overwritten by a stale snapshot, so this can't fight the form, a
# draft, or a rehydration -- all of which write widget keys directly and are
# already subject to the same instantiation rules (see rehydrate_proposal_into_form).
FORM_STATE_BACKUP = "_form_state_backup"

# Keys a snapshot must never carry, for three separate reasons.
#
# 1. Streamlit REFUSES to have the value of a button, download button, file
#    uploader or form set through session_state -- it raises
#    StreamlitAPIException when the widget is created. Restoring one of these
#    doesn't degrade, it takes the page down. tests/test_form_state.py walks
#    the AST for keyed widgets of those types and fails if a new one isn't
#    covered here, because remembering is not a strategy.
# 2. A data_editor's stored value is its *edit delta*, not its data. The grids
#    are rebuilt from plan_options / avails_seed_rows, which survive on their
#    own, so restoring a delta is unnecessary -- and an `added_rows` delta
#    replayed on top of rows that already contain the addition duplicates it.
# 3. page_choice is the navigation itself: restoring it from a snapshot would
#    fight the sidebar it came from.
NON_PERSISTABLE_PREFIXES = (
    # buttons and uploaders -- Streamlit raises on these
    "wo_upload", "wo_clear", "dup_btn_", "cs_upload",
    "deck_upload", "logo_upload", "usage_upload", "avails_pdf_upload_",
    # The Audience finder's AND/OR/New-group buttons (Phase 5 of the
    # targeting-groups roadmap, geo_targeting_roadmap.md D) -- replaced the
    # old single "Add" button (finder_add_). The mode radio, category
    # selectbox, search box and suggest textarea are all settable and
    # persist like any other.
    "finder_and_", "finder_or_", "finder_new_",
    # The Add-lines panel's button. Its three multiselects and its combine
    # checkbox are all settable and DO persist -- only the button can't, and
    # missing it took every return-from-another-page down with
    # "Values for the widget with key 'qa_add_1_0' cannot be set using
    # st.session_state". Caught by test_form_state rather than by review,
    # which is the whole reason that guard walks the AST for keyed widgets.
    "qa_add_",
    # Merge/split on the media plan grid (Phase 3 of the targeting-groups
    # roadmap, geo_targeting_roadmap.md D). Only the buttons -- the merge
    # multiselect and the split-target selectbox are ordinary settable
    # widgets and persist like any other.
    "tg_btn_",
    # The per-group geo-definition expander's Resolve button (Phase 6). The
    # mode radio and its mode-specific inputs are ordinary settable widgets
    # and persist like any other.
    "geo_resolve_btn_",
    # Same expander's "View on map" link (roadmap §E) and the Zip/map
    # builder page's own per-group download buttons -- the page's text_area
    # zip lists are ordinary settable widgets and persist like any other.
    "geo_goto_map_", "map_zipdl_",
    # The whole History page. Not only its buttons: a page the seller visited
    # before coming here leaves its widget state behind, and 96 of that page's
    # keys were measured riding along in a Build snapshot. They aren't Build's
    # to carry, and one of them being a button is all it takes.
    "hist_",
    # editors -- reconstructed from state that already survives
    "media_plan_editor_", "avails_editor_", "usage_categorize_editor_",
)

# The same rule for widgets keyed by what they act on rather than by what they
# are -- `f"{case_study_id}_save"`. A prefix list cannot see these.
NON_PERSISTABLE_SUFFIXES = ("_fetch", "_download", "_save", "_active")

# About the session rather than about the proposal, so they outlive a "New
# proposal" too.
SESSION_KEEP_ON_RESET = frozenset({
    "page_choice", "authed", "current_user", "identity_skipped",
})
SESSION_SCOPED_KEYS = SESSION_KEEP_ON_RESET | {FORM_STATE_BACKUP}


def _persistable(key):
    key = str(key)
    return (key not in SESSION_SCOPED_KEYS
            and not key.startswith(NON_PERSISTABLE_PREFIXES)
            and not key.endswith(NON_PERSISTABLE_SUFFIXES))


def snapshot_form_state():
    """Mirror the form into a key Streamlit won't collect.

    Taken after every widget on the page has been instantiated, and before
    Generate does any work -- so a snapshot exists whatever Generate then does
    or raises.
    """
    st.session_state[FORM_STATE_BACKUP] = {
        key: st.session_state[key] for key in list(st.session_state.keys())
        if _persistable(key)}


def restore_form_state():
    """Put back anything Streamlit collected while another page was showing."""
    for key, value in (st.session_state.get(FORM_STATE_BACKUP) or {}).items():
        if key not in st.session_state:
            st.session_state[key] = value


def clear_proposal_state():
    """Reset the Build form to first-load state.

    A deny-list, not an allow-list: everything goes except the handful of keys
    that describe the session rather than the proposal. Listing what to *clear*
    is the shape of bug this project keeps paying for -- it silently fails to
    cover whatever was added last, and the symptom here would be a "new"
    proposal quietly carrying the previous one's dirty rows, drafted badges,
    seed keys or uploaded schedule.
    """
    for key in list(st.session_state.keys()):
        if key not in SESSION_KEEP_ON_RESET:
            del st.session_state[key]


def render_new_proposal_button():
    """Arm the confirm. The confirm itself renders full width below the
    title -- see render_new_proposal_confirm."""
    if st.button("New proposal", use_container_width=True,
                 help="Clears this form completely — every field, the media plan, "
                      "the avails table, any drafted copy, and any uploaded logo or "
                      "Wide Orbit schedule."):
        st.session_state["confirm_new_proposal"] = True
        st.rerun()


def render_new_proposal_confirm():
    """The confirm step for New proposal, across the full page width.

    Not a `st.dialog`, and that was measured rather than assumed. A modal
    looks better and was built first, but its lifecycle doesn't survive
    either of the two things this needs. In the browser, dismissing it with
    the "x" leaves the arming flag set, so it reopens immediately and Cancel
    becomes the only way out -- a trap. Clearing the flag inside the dialog
    fixes that and breaks the other half: `AppTest` doesn't re-execute a
    dialog's fragment, so the buttons inside it are unreachable on the run
    that clicks them, and a destructive action would ship untested. Inline
    costs a little polish and keeps both.

    Full width rather than beside the title because there is a quarter of the
    page up there, which wrapped the confirm button onto two lines -- a
    destructive action offered in a cramped corner reads like a mis-click
    waiting to happen.
    """
    if not st.session_state.get("confirm_new_proposal"):
        return
    with st.container(border=True):
        st.markdown("**Start a new proposal?**")
        st.write("This clears the whole form — every field, the media plan and its "
                 "options, the avails table, anything drafted from notes, and any logo "
                 "or Wide Orbit schedule you've uploaded.")
        st.caption("Proposals you've already generated are safe — they stay in "
                   "Proposal history.")
        clear, cancel, _ = st.columns([1, 1, 3])
        if clear.button("Clear the form", type="primary", use_container_width=True):
            clear_proposal_state()
            st.rerun()
        if cancel.button("Cancel", use_container_width=True):
            st.session_state["confirm_new_proposal"] = False
            st.rerun()


DRAFT_KEY_SECTIONS.update({key: "attribution" for key in ATTRIBUTION_FIELD_MAP.values()})
# One key per vertical rather than one shared key: the checkbox says something
# different in each ("Polk New Car Sales Attribution" vs "Arrivalist
# Destination Attribution"), and a single key would carry an answer given
# about Polk over to a travel proposal.
DRAFT_KEY_SECTIONS.update({f"vertical_attribution_{v}": "attribution"
                           for v in assembly.VERTICAL_ATTRIBUTION})
DRAFT_KEY_SECTIONS.update({widget_key: "products"
                           for keys in PRODUCT_TO_WIDGET_KEYS.values()
                           for widget_key, _ in keys})


# ---------------------------------------------------------------------------
# Local test mode
#
# Skips the password gate and preselects a user so the real UI can be driven
# without anyone sharing a password. Deliberately awkward to switch on, and
# impossible to switch on in production:
#
#   * read ONLY from the OS environment -- never st.secrets, which is what a
#     deployed instance actually has, and which someone could set by mistake
#     while editing rates;
#   * refused outright when the process looks like Streamlit Cloud, so even
#     an env var set in the dashboard can't open the gate;
#   * loud on every run it's active, because a bypass nobody notices is one
#     that eventually ships.
# ---------------------------------------------------------------------------
TEST_MODE_ENV = "PROPOSAL_BUILDER_TEST_MODE"
TEST_MODE_USER = "Test Mode"

# Set by Streamlit Cloud's runtime; their presence means this is not a laptop.
_CLOUD_MARKERS = ("STREAMLIT_SHARING_MODE", "STREAMLIT_CLOUD",
                  "STREAMLIT_RUNTIME_ENV", "HOSTNAME_OVERRIDE")


def running_on_streamlit_cloud():
    if any(os.environ.get(marker) for marker in _CLOUD_MARKERS):
        return True
    # The deployed instance runs from /home/adminuser/... on Linux; a laptop
    # checkout never does.
    return sys.platform.startswith("linux") and Path.home().name == "adminuser"


def test_mode_active():
    """True only for a local process that asked for it via the environment."""
    if os.environ.get(TEST_MODE_ENV) != "1":
        return False
    if running_on_streamlit_cloud():
        # Refused rather than honoured: nothing legitimate sets this there.
        return False
    return True


class _InjectedUpload:
    """Stands in for Streamlit's UploadedFile: name plus getvalue()."""

    def __init__(self, path):
        self._path = Path(path)
        self.name = self._path.name

    def getvalue(self):
        return self._path.read_bytes()


def test_mode_upload(state_key):
    """An injected file for `state_key`, or None.

    Only in test mode, and only from session_state -- there is no widget and
    no secret involved, so a deployed instance has no way to reach it even if
    the key were somehow set.
    """
    if not test_mode_active():
        return None
    path = st.session_state.get(state_key)
    if not path or not Path(path).exists():
        return None
    return _InjectedUpload(path)


def check_password():
    if test_mode_active():
        st.session_state["authed"] = True
        st.session_state.setdefault("current_user", TEST_MODE_USER)
        st.warning(f"⚠️ **{TEST_MODE_ENV} is on.** The password gate is bypassed and you're "
                   f"signed in as \"{TEST_MODE_USER}\". This is for local testing only — "
                   f"unset the environment variable to restore the login.")
        return True
    return _check_password()


def _check_password():
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
    st.caption(f"Build {BUILD_STAMP}")
    return False


ADD_USER_OPTION = "➕ Add a name..."


def current_user():
    return st.session_state.get("current_user")


def check_identity():
    """Ask who's using the app, after the shared password."""
    # Not authentication -- the password is the gate, this only attributes
    # work. It's a separate step rather than a field on each form because
    # it's answered once per session and then used in three places
    # (proposals, case study uploads, attached-file notes).
    #
    # The list builds itself: anyone can add a name and it's there for
    # everyone afterwards, deduplicated case-insensitively so "matt" and
    # "Matt" can't become two people. If Supabase is unreachable the step is
    # skipped entirely rather than blocking -- an unattributed proposal is a
    # far better outcome than a seller who can't build one.
    #
    # Kept as comments rather than a second docstring paragraph on purpose:
    # this text shipped to every user who logged in. A bare string is only
    # exempt from Streamlit's magic when it is the FIRST statement in the
    # function, and the test-mode block below was later inserted above it --
    # which silently demoted the docstring to an expression that magic
    # rewrote into st.write(). A comment cannot be rendered by anything.
    if test_mode_active():
        st.session_state.setdefault("current_user", TEST_MODE_USER)
        return True

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

**One line per audience the budget is split between -- but one line for one audience, however many attributes describe it.** These read almost the same in notes and are completely different plans:
- **Separate audiences sharing the budget get separate lines**, each with its own allocation and its own "label" (e.g. "Premion Streaming TV — Commercial" and "Premion Streaming TV — Retail"). The signals are an explicit split ("60/40 between small business and consumer", "split evenly across all four audiences") or tracks the client names and talks about as distinct things.
- **One audience defined by several attributes stays ONE line**, with every attribute stacked into that line's "audience_track". "Homeowners, adults 35+, $150K+ household income" is a single group described three ways, not three audiences -- splitting it invents a budget division the client never asked for and pads the plan with rows that all reach the same people.
- **Several audiences named with no split given: still one line per audience**, each with a "split_evenly" allocation, and flag the assumption ("the notes named three audiences but didn't say how to weight them -- divided evenly, confirm the intended split"). Same principle as a requested product with no budget behind it: an assumption the reviewer can see gets corrected, an audience that never appears is invisible to them.
- **When it is genuinely unclear** whether the notes describe one stacked audience or several distinct ones, pick the reading the notes best support, build the plan that way, and say which reading you took. Do not hedge by doing both.

"audience_track" is that line's FULL targeting stack, not a single segment name. It renders directly into the deck's Targeting column -- the widest column on the media plan table -- and is what the client reads to understand who that line reaches. Give it every attribute that defines the audience, in the notes' own terms: "Adults 35+, homeowners, $150K+ income, researching cosmetic procedures", not "DEMO Homeowner". A one-segment answer under a four-attribute brief makes the plan look thinner than the campaign specs printed beside it.

Each line's "product" must be exactly one of:
- a product key: {list(PRODUCTS.keys())}
- a Live Sports package, written as "{SPORT_PRODUCT_PREFIX}<sport_key>" where <sport_key> is exactly one of {list(SPORTS.values())} (e.g. "{SPORT_PRODUCT_PREFIX}nfl_playoffs"). Always use this form for a sports buy -- never bill sports inventory as "premion_streaming_tv", which carries a completely different (much lower) rate.
- "{CUSTOM_FEE_PRODUCT}", for a one-time flat fee that isn't a real media-buy product (e.g. a production/creative fee) -- its "label" becomes its Tactic name directly and it must use a "flat_amount" allocation. Only use this when the notes state an actual dollar amount to charge. **Never emit a line whose amount is zero.** Things provided at no charge -- commercial production, reporting, account management -- are not media plan lines at all: put "commercial_production" in "attribution" instead and it appears in the deck's "Included with Campaign" list, which is where a client expects to see it.

**A quoted price always wins over the no-charge default.** The rule above describes what is normally free, not what is free in these notes. If the notes state a charge for commercial production ("we're doing the :30, quoted $850"), then it IS a chargeable line: emit a "{CUSTOM_FEE_PRODUCT}" line labelled for commercial production, and do NOT also put "commercial_production" in "attribution" -- billing for it and listing it as included at no charge in the same deck contradict each other in front of the client. The same applies to anything else on that list.

**Never move a quoted amount to a different deliverable than the notes attach it to.** If the notes quote $850 for commercial production, that $850 is the commercial production line -- it does not become a dynamic ad creation fee, a versioning fee, or anything else, however plausible the reassignment seems. Label the line with what the notes actually called it. If the notes are genuinely ambiguous about which deliverable a quoted amount covers, keep the wording from the notes as the label and say so -- a flagged ambiguity is reviewable, a silent reassignment is not.

Most proposals are a single plan: put its rows in "media_plan_lines" and leave "options" null. Only when the notes explicitly ask for SCENARIOS to choose between -- good/better/best, tiered budgets, "show them a $50K and a $75K version" -- return "options" instead, as up to {MAX_PLAN_OPTIONS} entries:
  "options": [{{"name": "Good", "total_budget": 50000, "breakout": "monthly", "media_plan_lines": [...]}}, {{"name": "Better", "total_budget": 75000, "breakout": "monthly", "media_plan_lines": [...]}}]
Each option is a complete plan in its own right, with its own budget and its own full set of lines, and becomes its own media plan slide in the deck. "name" is what the client sees appended to the plan title, so use the notes' own words for the tier ("Good"/"Better"/"Best", "$50K Plan") rather than a generic letter. When "options" is set, leave "media_plan_lines" empty. Do NOT invent scenarios the notes didn't ask for -- one plan is the normal answer.

EVERY option MUST carry its own "total_budget" -- it is what that scenario costs, and each option's allocations are resolved independently against it. An option's "total_budget" falls back to the top-level one only if omitted, and if neither is set that option prices at $0 and gets dropped entirely, which is never a useful answer. When the notes state a BUDGET RANGE with no instruction on how to split it ("between $50K and $75K", "somewhere in the 50 to 75 range, wants to see both"), the right answer is one option per stated figure, each carrying that figure as its own "total_budget" -- e.g. a "$50K Plan" at 50000 and a "$75K Plan" at 75000. If the notes give a range but you cannot tell what the individual figures should be, put a single plan at the lower figure and say so; never return lines with no budget behind them.

Set "total_tv" to true ONLY when the notes describe a broadcast schedule PREMION ITSELF is running, on one of our own two stations -- WUSA9 in DC, or WPMT/FOX43 in Harrisburg -- alongside the streaming plan: "keep the WUSA schedule going", "broadcast plan attached", "Total TV", "we're also running spots on FOX43", an existing station buy on one of those two being continued or added to. Total TV switches the deck to its co-branded station template and opens the panel where the seller uploads the Wide Orbit schedule, so getting it from the notes saves them a step and, more importantly, means they turn it on BEFORE building the plan rather than after.

**A client's own linear TV, on any station that isn't WUSA9 or WPMT/FOX43, is not Total TV -- it's background, not a Premion broadcast component**, however the notes describe it: "they're continuing their existing broadcast buy in Denver", "client also runs spots on the local Fox affiliate through their own agency", "keeping their linear schedule going alongside this." The presence of the words "broadcast" or "linear" is not the signal -- WHICH STATION, and whether Premion is the one running it, is. Do NOT set it for a streaming-only campaign, for a client's own linear buy on someone else's station or through a different vendor, or for a market other than DC/Harrisburg. Do NOT invent broadcast lines in "media_plan_lines" -- the schedule is imported from a real Wide Orbit export, not drafted. Flag that Total TV was switched on and that the schedule still needs uploading.

INCLUDE ONLY THE PRODUCTS THE NOTES ACTUALLY CALL FOR. There is no mandatory line and no default product -- "premion_streaming_tv" in particular is NOT required and must not be added just to have a baseline CTV line. A sports-only plan, an Audience-Marketplace-only plan, a retargeting-only plan, or a single-line plan are all perfectly valid proposals. If the notes describe an NFL campaign and nothing else, the correct media plan is one NFL line and nothing else. If the notes are genuinely silent about what to buy, say so instead of inventing a product mix.

**But every product the notes DO ask for gets a line, even when no budget is given for it.** "She wants audience targeting on top of the general streaming" names a product; the absence of a split is a missing number, not a reason to leave the product out. Include the line with a "split_evenly" allocation so it shares what's left with the other unspecified lines, and flag the assumption ("the notes didn't say how to split between X and Y -- divided evenly, confirm the intended weighting"). Omitting a requested product is the worse failure of the two: an assumption the reviewer can see gets corrected, whereas a product that never appears in the plan is invisible to them and quietly missing from the deck.

"total_budget" is always what the WHOLE campaign costs across the entire flight, never a monthly rate -- and so is every dollar amount the allocations resolve to. If the notes quote the budget per month ("$20K a month for three months"), multiply it out to the full-flight figure yourself ($60,000) and note the per-month figure if it's worth flagging.

"breakout" is a separate question -- not what the plan costs, but how it's presented: "monthly" for a plan broken out month by month (the normal case, and the right answer whenever the notes talk in per-month terms at all), or "full_flight" only if the notes specifically want the flight shown as a single combined period. Python divides the campaign total across the flight for a monthly breakout; do not do that arithmetic yourself.

Each line's "allocation" has exactly one key:
- "flat_amount": this line costs exactly this many dollars.
- "percent_of_total": this line costs this percent of total_budget.
- "percent_of_remainder": this line costs this percent of whatever's left after all "flat_amount" and "percent_of_total" lines are subtracted from total_budget (percent_of_remainder entries across lines should sum to 100 if they're meant to exhaust the remainder).
- "split_evenly": this line shares equally, with every other "split_evenly" line, in whatever's left after "flat_amount"/"percent_of_total"/"percent_of_remainder" lines are all accounted for -- use this for "split evenly across N audiences/tracks" instead of trying to pre-compute a percentage yourself.
- "percent_of_avails": this line buys this percent of one audience's available impressions -- a reach TARGET, not a budget. Use it only when the reach percentage itself is the thing being bought -- "reach 20% of the available audience", "one option at 20% and one at 40%", "40% penetration against the home-services segment", "spend whatever it takes to hit 25%". Pair it with "avails_ref", naming the audience it refers to -- use the exact same wording as that "audiences" entry's "segment" so the two can be matched up. State only the percentage and which audience; Python multiplies it out against the real avails figure, applies the CPM and the markup, and works out the cost. A reach line needs no budget: if the notes give both a reach percentage and a budget and the arithmetic disagrees, the reach wins and the difference is flagged for the reviewer.
  **Do NOT use "percent_of_avails" just because the notes mention a dollar figure alongside an avails number and ask what percentage that buys.** "$4,000 against the 1.7M avails on that segment, tell her what percent that reaches" states a BUDGET ($4,000) and asks for the resulting reach to be reported -- it is not asking to spend whatever it takes to hit a percentage. That is a "flat_amount" (or "percent_of_total") line like any other; Python derives the reach percentage that budget buys from the real avails figure and reports it automatically, so no field is needed to make that happen. The question that tells the two apart: which number must come out exactly as stated -- a reach percentage that spend is built to hit, or a dollar figure that reach is only measured against afterward? The second case is far more common than the first.
Do NOT do any arithmetic yourself beyond picking which allocation type fits each line -- Python resolves flat_amount and percent_of_total first, then percent_of_remainder, then splits whatever's left evenly across split_evenly lines, then computes every dollar amount, impression count, and markup.

Each line may also carry an optional "cpm", the rate for that line in dollars. Rates are negotiated per deal, so set it whenever the notes state a rate for that line -- "$28 CPM on the Premion line", "they're getting the streaming at 30", "we agreed $45 for the NFL inventory". Omit it and the product's rate card default applies, which is what you want whenever the notes say nothing about rate or say to hold to the rate card ("at rate card", "standard rates", "no discount"). Set it ONLY from a rate the notes actually state -- never to hit a budget or impression target, which is what the allocations are for. Write it as a plain JSON number -- 28 -- leaving the currency symbol and the word CPM to the app. It is the net rate before any agency markup (Python applies the markup), and it never applies to a "{CUSTOM_FEE_PRODUCT}" line, which has no rate at all. Every override is flagged for the reviewer automatically, so you do not need to mention it yourself.

Rules:
- **A stated dollar budget always drives cost and impressions; a stated avails/reach figure next to it is only a ceiling to report reach against, never a spending target.** When the notes give BOTH a budget and an avails number for the same line, price it as "flat_amount" (or "percent_of_total") against the budget -- never "percent_of_avails". "$4,000 against the 1.7M avails on that segment, what percent does that reach" is a $4,000 budget with a reach question attached, not a request to spend whatever it takes to hit a percentage; Python computes and reports that percentage automatically once the line is priced from the budget. Reserve "percent_of_avails" for the rarer case where the notes state a reach PERCENTAGE as the thing being bought -- see the allocation type below for exactly how to tell the two apart. A dollar figure is a budget even when an avails number sits in the very same sentence.
- **"agency_involved" is whether an agency is taking a commission on THIS buy, not whether the word "agency" appears anywhere in the notes.** Set it true only when the notes say an agency is placing this buy, marking it up, or otherwise taking a cut of it -- "their agency is placing this, bill gross", "agency commission applies". Set it false whenever the notes say to keep rates net, no agency markup, no commission, or that the client is buying direct -- even when an agency is named elsewhere in the notes (a different line of business, a different media type, or named and then explicitly excluded from this buy). An agency's name appearing in the notes is not evidence of a commission on this deal; an explicit net/no-commission instruction always wins over that inference.
- "vertical" must be exactly one of: {list(VERTICALS.values())}
- "market" must be exactly "DC" or "Harrisburg". This is the ORIGINATING station -- which Premion office the proposal comes from -- and it is not where the campaign runs.
- "target_markets" is where the campaign is AIMED: a list of DMA or city names exactly as the notes give them ("Denver", "Atlanta", "Washington DC", "the Bay Area"). List EVERY market the notes name, in the order they are named -- a brief that says "Denver, Atlanta and Phoenix" produces three entries, not one. The app matches each name against the real Nielsen DMA list and reports anything it can't place, so give the name as written rather than guessing at an official spelling. Leave it empty when the notes name no target market at all; do not fall back to the originating market, which is a different thing and is already captured above.
- "sports" entries must be exactly one of: {list(SPORTS.values())}. This drives which sports package slides go in the deck -- list every package that also appears as a "{SPORT_PRODUCT_PREFIX}" media plan line, and leave it empty when the notes call for no sports at all.
- "attribution" entries must be drawn from this list, using the exact key shown. Include every one the notes call for -- these drive real slides, real Included-with-Campaign entries and real toggles, so an option the notes ask for and you omit simply never reaches the proposal:
{attribution_help}
- "audiences[].segment" must be an EXACT name from the audience catalog slice below -- do not paraphrase or invent segment names. If nothing in the slice fits, it's fine to omit audiences or note it. {AUDIENCE_MATCH_GUIDANCE} When a media_plan_lines entry's "audience_track" describes the same audience as one of your "audiences" entries, use the same wording for both.
- **At most ONE non-RFP-selectable segment. This is a hard limit, not a preference.** Each catalog entry carries an "rfp_selectable" flag, and a campaign may book only one segment with "rfp_selectable": false (a "custom" segment). A draft containing two or more is invalid and cannot be used until someone removes the extras by hand. So: when two segments would serve the same purpose, take the RFP-selectable one. If several custom segments all look relevant -- which happens when a niche category's best matches are all custom -- **choose the single most important one and name the others as alternatives the reviewer could swap in**, rather than returning them all. Count the custom segments in your "audiences" array before you finish; if there is more than one, cut it down. Whenever you do return a custom segment, say which one it is and why no RFP-selectable segment covered it -- and say it in the seller's words: **"a custom audience", or "not directly selectable in Salesforce"**, never by naming the catalog field the flag lives in. That flag is data for you to read, not vocabulary to repeat. Always mention a custom segment in one of the two lists: it is the only thing telling the seller this audience can't just be booked through Salesforce RFP.
- "audiences[].max_avails" is the audience's available impressions, and **it comes from the notes or it stays out**. Set it ONLY when the notes give a number for that specific audience -- "about a million a month on the home-services segment", "600K avails", "we can get 400,000 impressions against that one". Never estimate one, never carry a figure across from a different audience, and never work one back from a budget, a CPM or an impression goal: a made-up avails number looks exactly like a real one to the person sending the proposal, and the whole point of the field is that it came from the avails system. Leave it out entirely when the notes are silent -- an audience with no stated avails is flagged automatically for the seller to pull the real number, which is the correct outcome and needs no note from you.
- "audiences[].avails_basis" says which basis that number was given in: "monthly" for a per-month figure (what the avails system reports, and the default), "flight" for a whole-campaign total ("1.5 million over the three months"). Report the basis the notes used and let Python convert -- do not divide or multiply it yourself. If the notes give a number without saying which it is, use "monthly" and flag the assumption.
- "unresolved" is a list a salesperson reads before they send the proposal. **Write it for them, not for a developer.** Rules, all of which matter:
  * **Two sentences per item, maximum.** What you assumed, then what to confirm. Nothing else -- no reasoning, no justification, no explanation of how the app works.
  * **Name every thing the way a seller would say it out loud.** An audience that can't be booked straight through Salesforce RFP is "a custom audience", or "not directly selectable in Salesforce"; a one-time charge is "the production fee line"; a sports package is "NFL Regular Season"; a share of what's left of the budget is "the remaining budget"; a measurement option is "the brand lift study". Schema keys, catalog flags and product codes do their job inside the JSON fields themselves and have no business in a sentence a salesperson reads. If you can't put a thing in a seller's own words, leave it out of the list.
  * **At most 8 items, and fewer is better.** Merge anything related into one item -- three questions about the flight dates are one item about the flight dates. If you have more than 8, you're flagging things that don't need flagging.
  * Put anything the CLIENT has to answer in "unresolved". Put checks the SELLER does on their own -- pulling real avails numbers, confirming a rate internally, double-checking a segment is available, uploading a file -- in "unresolved_internal". Same rules apply to both. **This is the only rule that decides which list something goes in.** Everywhere else in these instructions that tells you to flag, note or say something, it means "write it as an item" and this rule alone picks the list -- so route by who has to act on it, never by where the wording of some other rule happens to point.
  * Write plain sentences, the way you would say them to a colleague: "No end date was given, so the flight runs through November 30. Confirm the end date." State the assumption, then what to confirm, and stop there.
- Dates in flight_start/flight_end should be YYYY-MM-DD. If the notes give a date without a year (e.g. "September through November"), resolve it to the NEXT upcoming occurrence of that month relative to today's date -- never a date already in the past -- and flag that assumption the same as any other assumption.

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


def log_claude_call(record):
    """Record one call, to the console and to a file. Never raises.

    Both, deliberately: the console is what's visible in Streamlit Cloud's
    log viewer, the file is what survives locally after the tab is closed.
    A logging failure must never be what takes a draft down, so every error
    here is swallowed -- the draft is the point, the log is the evidence.
    """
    line = json.dumps(record, default=str)
    print(f"[claude] {line}")
    try:
        with open(CLAUDE_LOG_PATH, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass


def _response_text_and_facts(response):
    """(raw_text, facts) for a Claude response.

    Joins EVERY text block rather than reading content[0].text: a response
    whose first block isn't text -- or which has no blocks at all, which is
    what a pre-output refusal returns -- would otherwise raise IndexError or
    AttributeError inside the try block and be reported as "Claude API call
    failed", hiding what actually happened.
    """
    blocks = list(getattr(response, "content", None) or [])
    raw_text = "".join(getattr(b, "text", "") for b in blocks
                       if getattr(b, "type", None) == "text")
    usage = getattr(response, "usage", None)
    stop_reason = getattr(response, "stop_reason", None)
    facts = {
        "stop_reason": stop_reason,
        "blocks": [getattr(b, "type", "?") for b in blocks],
        "chars": len(raw_text),
        "input_tokens": getattr(usage, "input_tokens", None),
        "output_tokens": getattr(usage, "output_tokens", None),
        "max_tokens": ANTHROPIC_MAX_TOKENS,
        "request_id": getattr(response, "_request_id", None),
        "head": raw_text[:200],
        "tail": raw_text[-200:],
    }
    # stop_details is populated only on a refusal and is None otherwise --
    # read it without guarding and it's an AttributeError on every ordinary
    # response.
    if stop_reason == "refusal":
        details = getattr(response, "stop_details", None)
        facts["refusal_category"] = getattr(details, "category", None)
    return raw_text, facts


def interpret_claude_response(response, label):
    """(parsed, error, retryable) for one raw response.

    Split out from the network call so the failure paths are testable
    without an API key: the truncation and empty-response branches are
    exactly the ones that cost a live run to discover, and they should not
    need another one to re-check.

    stop_reason is consulted BEFORE the JSON is parsed. Truncated JSON fails
    to parse, so a max_tokens cut-off used to be reported as a parse error --
    which points the reader at the response's formatting instead of at its
    length, and is why the original failure was mis-diagnosed as a
    fence-stripping problem.
    """
    raw_text, facts = _response_text_and_facts(response)
    facts["label"] = label

    if facts["stop_reason"] == "max_tokens":
        log_claude_call({**facts, "outcome": "truncated"})
        return None, (
            f"The draft was too long to finish -- Claude hit the {ANTHROPIC_MAX_TOKENS:,}-token "
            f"response limit and the JSON was cut off mid-way. Try fewer plan options, fewer "
            f"lines per option, or shorter notes. If this keeps happening on a plan that's "
            f"genuinely this big, the limit itself needs raising (ANTHROPIC_MAX_TOKENS in "
            f"app.py)."), True

    if facts["stop_reason"] == "refusal":
        log_claude_call({**facts, "outcome": "refusal"})
        return None, (
            "Claude declined to answer this request. Re-word the notes and try again; if they "
            "contain nothing unusual, this is worth reporting."), False

    if not raw_text.strip():
        log_claude_call({**facts, "outcome": "empty"})
        return None, (
            f"Claude returned an empty response (stop reason: {facts['stop_reason']}). Nothing "
            f"was wrong with the notes -- try again, and if it repeats, the log at "
            f"{CLAUDE_LOG_PATH} has the details."), True

    parsed, error = _parse_draft_json(raw_text)
    log_claude_call({**facts, "outcome": "ok" if parsed is not None else "unparseable"})
    if parsed is None:
        return None, error, True
    return parsed, None, False


def _call_claude_json(prompt, label="draft", attempts=2):
    """Sends one prompt to Claude and parses the response as JSON. Returns
    (parsed_dict, error_message) -- exactly one is None.

    Retries once on an empty, truncated or unparseable response before
    surfacing anything: all three are transient often enough that making a
    seller re-paste their notes and wait again is the wrong first move. An
    API-level failure (auth, network, rate limit) is not retried here -- the
    SDK already retries those itself.
    """
    api_key = st.secrets.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None, "ANTHROPIC_API_KEY is not set in .streamlit/secrets.toml."

    client = anthropic.Anthropic(api_key=api_key)
    error = "Claude was not called."
    for attempt in range(1, attempts + 1):
        try:
            response = client.messages.create(
                model=ANTHROPIC_MODEL,
                max_tokens=ANTHROPIC_MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}],
            )
        except Exception as exc:                                 # noqa: BLE001
            log_claude_call({"label": label, "attempt": attempt, "outcome": "api_error",
                             "error": f"{type(exc).__name__}: {exc}",
                             "prompt_chars": len(prompt)})
            return None, f"Claude API call failed: {exc}"

        parsed, error, retryable = interpret_claude_response(
            response, f"{label} (attempt {attempt} of {attempts})")
        if parsed is not None:
            return parsed, None
        if not retryable or attempt == attempts:
            break
    return None, error


def call_claude_draft(notes):
    """Returns (draft_dict, error_message) -- exactly one is None."""
    return _call_claude_json(build_draft_prompt(notes), label="draft")


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
- Remove from both lists anything the clarifications have now settled, and keep anything still genuinely open. If the clarifications raise something new and ambiguous, add it.
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
    revised, error = _call_claude_json(
        build_redraft_prompt(notes, previous_draft, clarifications), label="redraft")
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
    parsed, error = _call_claude_json(
        build_audience_finder_prompt(description, vertical_hint), label="audience_suggest")
    if error:
        return None, None, error

    raw_recs = parsed.get("recommendations", []) or []
    names = [r.get("segment", "") for r in raw_recs if r.get("segment")]
    matched, unmatched = validate_segments(names)
    recs = [r for r in raw_recs if r.get("segment") in matched]
    return recs, unmatched, None


def build_categorize_prompt(components):
    """`components` is [{"name", "impressions"}, ...] -- the workbook
    components rules genuinely can't resolve (no recognizable prefix of
    their own, and not a CUSTOM with a resolvable second one). Category
    descriptions are findability language (see audience_catalog.
    CATEGORY_DESCRIPTIONS), the same text a rep would see -- not a
    taxonomy definition, so the model reasons about "where would a seller
    look for this" the same way it's asked to reason about it."""
    category_lines = "\n".join(f"- {cat}: {desc}"
                               for cat, desc in sorted(CATEGORY_DESCRIPTIONS.items()))
    items = "\n".join(f"{i}. {c['name']} (impressions to date: {c['impressions']:,})"
                      for i, c in enumerate(components))
    return f"""You are categorizing Premion CTV/OTT audience-targeting segments for a media seller's catalog, so a rep can FIND a segment by browsing a category -- not writing a formal taxonomy. Return ONLY valid JSON -- no markdown fences, no preamble -- matching this schema:

{{"components": [{{"index": 0, "category": "one category name from the list below, or null", "category2": "a SECOND category name, or null", "is_client_specific": false, "confidence": "high", "reason": "one short sentence"}}]}}

Categories (use EXACTLY one of these names -- never invent, abbreviate, or combine one):
{category_lines}

Rules:
- "category" must be exactly one of the names above, or null if none genuinely fits (that's a legitimate answer, not a failure).
- "category2" is null unless the segment truly belongs in a second category too -- most components should have only one. Findability beats taxonomic purity: if a rep working either category would reasonably expect to find this segment there, name it; don't stretch to force a second one that isn't real.
- "is_client_specific": true if this reads as ONE ADVERTISER'S OWN retargeting pool, address list, or a specific client/campaign/vendor name rather than a reusable audience anyone could target (e.g. "CUSTOM City of Mesa", "Gallery Furniture Conquesting", a named business's own website visitors) -- judge this the same way you'd judge whether a stranger's ad campaign should be able to reuse this exact segment. When true, "category" is usually null (a client-specific segment doesn't need a findability category -- it's being excluded from the shared catalog, not filed under one).
- "confidence" is your own honest confidence in the "category"/"category2" call: "high", "medium", or "low". A null category with is_client_specific=false is still worth a confidence -- it says how sure you are that NOTHING fits.
- "reason" is one short sentence a media seller would find useful -- plain language, not a taxonomy justification.
- Return exactly one entry per numbered component below, "index" matching its number, in any order.

Components ({len(components)}):
{items}
"""


def call_claude_categorize(components):
    """`components` is [{{"name", "impressions"}}, ...]. Returns
    (suggestions, error) -- error_message is None on success. Each
    suggestion is {{"segment", "category", "category2", "is_client_specific",
    "confidence", "reason"}}, ALREADY VALIDATED: a "category"/"category2"
    Claude returns that isn't one of the real category names is dropped to
    "" rather than trusted, so an invented category can never reach the
    catalog. Suggest, don't decide -- this is a proposal for the admin page
    to show and the rep to confirm or correct, never written anywhere on
    its own.
    """
    if not components:
        return [], None
    parsed, error = _call_claude_json(build_categorize_prompt(components), label="categorize_audience")
    if error:
        return None, error

    valid_categories = set(CATEGORY_DESCRIPTIONS)
    by_index = {i: c["name"] for i, c in enumerate(components)}
    suggestions = []
    for entry in parsed.get("components", []) or []:
        index = entry.get("index")
        name = by_index.get(index)
        if name is None:
            continue  # an index Claude invented or skipped -- ignore rather than guess which component it meant
        category = entry.get("category")
        category = category if category in valid_categories else ""
        category2 = entry.get("category2")
        category2 = category2 if (category2 in valid_categories and category2 != category) else ""
        suggestions.append({
            "segment": name,
            "category": category,
            "category2": category2,
            "is_client_specific": bool(entry.get("is_client_specific")),
            "confidence": entry.get("confidence") or "",
            "reason": entry.get("reason") or "",
        })
    return suggestions, None


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
            # One deck-wide opt-in, not one per sport. It adds no media plan
            # line, so it can never change what the seed key's diff decides --
            # it rides here rather than in its own field only because
            # `products` is the shape assembly and the stored proposal both
            # read the sports selection out of.
            "viewership": get("include_sport_viewership", False),
        },
        "total_tv": get("total_tv", False),
        "dynamic_creative": get("dynamic_creative", False),
    }


def read_seed_selections(get=None):
    get = get or _session_getter
    schedule = get("broadcast_schedule", None)
    return {"products": read_products_selection(get),
            "_premion_streaming_tv": get("premion_streaming_tv", False),
            # Importing or removing a Wide Orbit schedule changes which lines
            # exist, so it belongs in the seed signature exactly like a
            # product toggle. Reduced to a stable fingerprint rather than the
            # object itself, which has no meaningful repr.
            "_broadcast": (f"{schedule.summary.station}|{schedule.summary.total_spots}|"
                           f"{schedule.summary.gross_cost}" if schedule else None)}


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
    # "avails" is what the client was shown, in whatever basis was on screen
    # at the time, so a rebuild reproduces it verbatim. Proposals logged
    # before the basis toggle existed carry no label and were all monthly,
    # which is exactly what the default gives them.
    total_avails = sum(int(str(r.get("avails", "0")).replace(",", "") or 0) for r in avails_rows)
    avails_label = form.get("avails_label") or AVAILS_COLUMN_MONTHLY
    vertical_key = row.get("vertical") or "none"
    vertical_label = next((label for label, key in VERTICALS.items() if key == vertical_key), "")
    specs = form.get("campaign_specs") or {}
    # The stored flight label, not a recomputed one -- same "verbatim, never
    # recomputed" rule as the media plan itself. This is Generate's own
    # fallback for an empty Timing narrative (main(), the live TIMING_BULLETS
    # line); rebuild used to fall back to "--" instead, a second, drifted
    # construction of the same fallback that made a rebuilt deck read
    # differently from the one the client actually received whenever Timing
    # was left blank. Caught by tests/test_group_backward_compat.py.
    stored_flight_label = (form.get("flight") or {}).get("label") or "--"

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
        # `.get(key, default)`, not `or default`: the live Generate handler
        # stores whatever the widget held, verbatim, including a
        # deliberately-blank title -- `or` would substitute the synthesized
        # default for that real, empty value too, same drift TIMING_BULLETS
        # had. The synthesized default is only correct when the KEY ITSELF
        # is absent (a proposal logged before this field existed).
        "proposal_title": form.get("proposal_title", "CTV Strategy"),
        "logo_path": logo_path,
        "vertical_display": vertical_label if (form.get("selections") or {}).get("vertical") != "none" else "",
        "campaign_specs": {
            "GOALS_BULLETS": lines_to_bullets(specs.get("goals", "")) or ["--"],
            "AUDIENCE_BULLETS": lines_to_bullets(specs.get("audience", "")) or ["--"],
            "GEOGRAPHY_BULLETS": lines_to_bullets(specs.get("geography", "")) or ["--"],
            "BUDGET_BULLETS": lines_to_bullets(specs.get("budget", "")) or ["--"],
            "PLACEMENTS_BULLETS": lines_to_bullets(specs.get("placements", "")) or ["--"],
            "TIMING_BULLETS": lines_to_bullets(specs.get("timing", "")) or [stored_flight_label],
        },
        "avails": {"rows": avails_rows, "total_avails": f"{total_avails:,}",
                   "label": avails_label,
                   # Regenerated from the stored groups, not stored as bytes:
                   # rendering is a pure, deterministic function of
                   # resolved_zips/color (no randomness, no timestamp in the
                   # PNG), so re-rendering the SAME stored data reproduces
                   # the SAME image byte-for-byte -- verbatim in substance,
                   # exactly like every other rebuilt field, without paying
                   # to store a picture that's fully implied by data already
                   # kept. None (and therefore no picture at all) for any
                   # proposal logged before this key existed.
                   #
                   # dark=True, matching the live Generate handler exactly
                   # (see its own map_png comment) -- a resolved proposal
                   # always keeps the map-variant slide (dark gradient, no
                   # stock photo; `rebuild_selections`/`selections` below
                   # reuse the ORIGINAL `targeting_map_present`, so that
                   # choice is preserved too), and painting an OPAQUE
                   # light-mode map onto that dark background is not a
                   # cosmetic mismatch -- it's a rebuild that visibly
                   # differs from what the client was actually sent. Found
                   # by tests/test_group_backward_compat.py: this was the
                   # ONE part that didn't come back byte-identical.
                   "map_png": targeting_map.render_map(form.get("targeting_groups") or [], dark=True,
                                                       label_for=_market_display_name)},
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
            case_study_sources.append(case_study_source(stored))
        except Exception as exc:
            warnings.append(f"Case study \"{stored.get('title')}\" couldn't be fetched "
                            f"({db.describe_error(exc)}) -- rebuilt without it.")

    # The markets this proposal was built for, resolved the same way generate
    # resolves them. A rebuild that swapped in different profile slides -- or
    # dropped back to the national one -- would not be "as presented".
    rebuild_selections = form.get("selections") or {}
    profile_paths, profile_warnings = market_profile_images(
        target_dma_list(rebuild_selections, row),
        rebuild_selections.get("include_market_profile"))
    warnings.extend(profile_warnings)

    try:
        prs, _, _ = assembly.build_presentation(master_path, rebuild_selections)
        assembly.replace_market_profile_slides(prs, profile_paths)
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
    # Same `.get(key, default)` fix as rebuild_proposal_deck, for the same
    # reason: a deliberately-blank title has to survive "Load into form"
    # verbatim, or a regenerate-without-touching-anything stops matching
    # the original the moment the title was ever cleared.
    updates["proposal_title"] = form.get("proposal_title", "CTV Strategy")

    market = selections.get("market") or row.get("market")
    if market in ("DC", "Harrisburg"):
        updates["market_choice"] = market
    market_label = ("Washington, DC DMA" if market == "DC"
                    else "Harrisburg DMA" if market == "Harrisburg" else "")

    # Target DMAs are restored by KEY and re-labelled from the current market
    # list, not from whatever labels were stored: a market whose profile slide
    # has been added or removed since reads differently now, and a
    # multiselect's values have to match the options it is actually showing or
    # Streamlit drops the ones that don't -- silently, which would look like
    # the proposal had targeted fewer markets than it did.
    #
    # ORDER IS PRESERVED, because it is what drives slide order. Restoring a
    # three-market proposal with its slides in a different sequence would be a
    # different deck from the one the client was sent.
    stored_dmas = target_dma_list(selections, row)
    # Hoisted: the Geography/Geo defaults below need the same labels, and a
    # proposal that reloaded with different Geo than it was built with would
    # not be the proposal that was sent.
    restored_profiles = []
    if stored_dmas:
        restored_profiles, _ = load_market_profiles()
        profiles = restored_profiles
        by_key = {r.get("key"): r for r in profiles}
        include_stored = bool(selections.get("include_market_profile"))
        restored, gone = [], []
        for key in stored_dmas:
            current = by_key.get(key)
            if current is None:
                gone.append(key)
                continue
            restored.append(market_profile_option_label(current))
        if restored:
            updates["target_dma_choice"] = restored
        # A market can lose its slide between the original build and now.
        # Restoring include=True in that state would tick a box against
        # slides that don't exist, so it's the AND of both facts.
        updates["include_market_profile"] = bool(
            include_stored
            and any(by_key.get(k, {}).get("image_path") for k in stored_dmas))
        if gone:
            notes.append(
                f"{len(gone)} target DMA(s) this proposal was built for are no "
                f"longer in the market list ({', '.join(gone)}), so they haven't "
                "been restored -- re-pick them before generating if they still "
                "need a market profile.")
        # Deliberately NOT one note per slide-less market. Rehydration can
        # only see the market list as it stands now, so it cannot tell a
        # market that LOST its slide from one that never had a slide to lose
        # -- and five of them never did. Saying "there's no profile slide for
        # it now" about Honolulu describes a change that didn't happen. The
        # restored picker already labels each one "(no profile slide)", which
        # is the true statement and is where the rep is looking anyway.
        #
        # What IS worth saying is the case where the rebuild will differ from
        # what the client was sent: profiles were included, and nothing in
        # the set can supply one any more, so the deck falls back to the
        # national slide.
        if include_stored and restored and not any(
                by_key.get(k, {}).get("image_path") for k in stored_dmas):
            notes.append(
                "This proposal included market viewer profile slides, but none "
                "of its markets has one now -- the rebuild will carry the "
                "national Total U.S. slide instead.")

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
    # Absent means the proposal predates the toggle and so was built with the
    # viewership slides in -- the same default resolve_active_keys applies, so
    # loading a row and regenerating it reproduces the deck it describes.
    updates["include_sport_viewership"] = bool(sports.get("viewership", True))
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

    # Only the checkbox this proposal's own vertical would have shown. A row
    # logged before the checkbox existed carries None, and the checkbox's own
    # default (checked) is the right answer for it: those decks all had the
    # vertical's attribution slide in, since it rode in with the vertical.
    va_spec = assembly.vertical_attribution_spec(vertical_key)
    if va_spec is not None and selections.get("vertical_attribution") is not None:
        updates[f"vertical_attribution_{vertical_key}"] = bool(selections["vertical_attribution"])

    # --- Campaign Specs copy --------------------------------------------
    specs = form.get("campaign_specs") or {}
    for field in ("goals", "audience", "geography", "budget", "placements", "timing"):
        updates[f"{field}_text"] = specs.get(field) or ""

    # The one geo default for this loaded proposal, derived exactly as main()
    # will after the rerun. Used by the avails rows below and by
    # _shared_fields_key further down, so a reloaded proposal can't show a
    # different Geo in the avails table than on its own media plan.
    rehydrated_geo = geo_column_default(
        target_market_labels(stored_dmas, restored_profiles),
        updates.get("geography_text", ""), market_label)

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
        # Prefer the stored monthly figure over the displayed string: the
        # string is in whatever basis the table was showing, and a full-flight
        # one loaded straight into a monthly column would multiply the
        # audience's avails by the month count. Rows logged before the basis
        # toggle existed have no monthly figure and were monthly anyway, so
        # the fallback is exact for them.
        updates["avails_seed_rows"] = [
            {"Audience": r.get("audience", ""), "Geo": r.get("geo", "") or rehydrated_geo,
             AVAILS_COLUMN_MONTHLY: int(r.get("avails_monthly") if r.get("avails_monthly") is not None
                                        else (str(r.get("avails", "0")).replace(",", "") or 0))}
            for r in real_avails
        ]
        updates["avails_version"] = st.session_state.get("avails_version", 0) + 1
    stored_basis = form.get("avails_basis")
    if stored_basis in (AVAILS_BASIS_MONTHLY, AVAILS_BASIS_FLIGHT):
        updates["avails_basis"] = stored_basis

    # A proposal saved after targeting groups shipped carries them directly;
    # one logged before that migrates its flat avails rows into groups here,
    # once, on load -- the "flat rows migrate into groups on load" rule.
    # Never re-derived from the ORIGINAL form's avails_rows (which predates
    # any edit this rehydration made above) -- the rows just written into
    # updates["avails_seed_rows"] are the ones groups must agree with.
    stored_groups = form.get("targeting_groups")
    if stored_groups is not None:
        updates["targeting_groups"] = stored_groups
    elif "avails_seed_rows" in updates:
        updates["targeting_groups"] = tg.seed_rows_to_groups(
            updates["avails_seed_rows"], AVAILS_COLUMN_MONTHLY)
    # sync_targeting_groups decides "which side moved" by comparing
    # avails_seed_rows against _groups_rows_applied, the row projection it
    # last wrote FROM groups. Rehydration writes both keys together, in
    # agreement by construction -- but _groups_rows_applied otherwise stays
    # whatever an EARLIER proposal in this session left it (or unset, for a
    # fresh one), so the very next sync_targeting_groups call saw a
    # "mismatch" that was never real and re-derived targeting_groups from
    # the flat rows via seed_rows_to_groups -- correct for that function's
    # actual job of migrating a pre-groups proposal, which invents no
    # resolved geography, but wrong here: for an already-group-shaped
    # proposal it silently downgraded a real radius/zips geo_def to
    # kind:text and wiped resolved_zips/resolved_markets. Invisible until
    # something actually depended on them surviving -- the targeting map
    # (tests/test_group_backward_compat.py caught a rehydrated-then-
    # regenerated deck missing its map entirely, because there was nothing
    # left resolved to draw). Setting this here keeps the two in agreement
    # from the very first render, so sync_targeting_groups has nothing to
    # "fix" that wasn't broken.
    if "avails_seed_rows" in updates:
        updates["_groups_rows_applied"] = [dict(r) for r in updates["avails_seed_rows"]]

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

    default_targeting = audience_stack(updates.get("audience_text", ""))
    # Must derive the same way main() will after the rerun, or _shared_fields_key
    # disagrees and every restored row is re-seeded over the numbers being
    # loaded. Restored rows are marked dirty, which protects the row VALUES --
    # but the key still has to match or the grid churns for nothing.
    default_geo = rehydrated_geo
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
    st.session_state["draft_unresolved_internal"] = []
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
    # Two lists on purpose. `unresolved` is what the client has to settle;
    # `internal` is what the seller checks before sending. They're shown as
    # separate sections because mixing them made a list where nothing looked
    # actionable -- every note this function appends is an internal check.
    unresolved = list(draft.get("unresolved", []))
    internal = list(draft.get("unresolved_internal", []) or [])
    updates = {}
    touched_sections = set()

    vertical_val = draft.get("vertical")
    vertical_label = next((label for label, key in VERTICALS.items() if key == vertical_val), None)
    if vertical_label:
        updates["vertical_choice"] = vertical_label
    elif vertical_val:
        internal.append(f"Vertical '{vertical_val}' not recognized -- left unchanged.")

    market_val = draft.get("market")
    if market_val in ("DC", "Harrisburg"):
        updates["market_choice"] = market_val
    elif market_val:
        internal.append(f"Market '{market_val}' not recognized -- left unchanged.")
    market_label = "Washington, DC DMA" if market_val == "DC" else "Harrisburg DMA" if market_val == "Harrisburg" else ""

    # Target markets: EVERY market the notes name, not the first one. A brief
    # saying "Denver, Atlanta and Phoenix" is a three-market campaign and the
    # deck carries a profile slide for each; taking only the head of that list
    # would silently ship a one-market proposal against a three-market brief.
    target_market_names = draft.get("target_markets") or []
    if isinstance(target_market_names, str):      # a model returning one string
        target_market_names = [target_market_names]
    # The plain market names, used below to default Geography and the plan's
    # Geo column -- kept separate from the picker's option strings, which
    # carry the "(no profile slide)" suffix and must never reach a slide.
    draft_geo_labels = []
    if target_market_names:
        profiles, _ = load_market_profiles()
        chosen_labels, seen = [], set()
        for name in target_market_names:
            key, candidates = market_profiles.match_market(name, profiles)
            if key is None:
                if candidates:
                    options = ", ".join(
                        next((r.get("label") for r in profiles if r.get("key") == c), c)
                        for c in candidates[:4])
                    internal.append(
                        f"\"{name}\" could be more than one market ({options}) -- "
                        f"pick the right one in Target DMAs.")
                else:
                    internal.append(
                        f"\"{name}\" didn't match a Nielsen DMA, so it isn't "
                        f"selected as a target market -- add it by hand if it's real.")
                continue
            if key in seen:      # the notes named the same market twice
                continue
            seen.add(key)
            row = next(r for r in profiles if r.get("key") == key)
            chosen_labels.append(market_profile_option_label(row))
            draft_geo_labels.append(row.get("label"))
        if chosen_labels:
            updates["target_dma_choice"] = chosen_labels
            updates["include_market_profile"] = any(
                r.get("image_path") for r in profiles
                if r.get("key") in seen)

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
        internal.append("Flight start date missing or unparseable -- left unchanged.")
    if flight_end:
        updates["flight_end"] = flight_end
    else:
        internal.append("Flight end date missing or unparseable -- left unchanged.")
    # The month count that spreads this budget must be the one main() derives
    # after the rerun -- it is what totals the plan, labels the totals row and
    # divides a flat fee across months. Two ways they used to diverge, both of
    # which shipped a plan whose allocation and totals disagreed:
    #
    #   1. A draft set new dates but never set `active_months`, so the months
    #      picked for the PREVIOUS flight survived -- and main() only discards
    #      a stored selection when it doesn't overlap the new range at all, so
    #      a partial overlap kept a stale subset. An Oct-Dec draft landing on
    #      the form's default Sep-Nov flight left main() on 2 months while the
    #      allocation had used 3: $45,000 spread over three months, totalled
    #      over two, and rendered as "Full Flight Total (2 months)  $30,400".
    #   2. No usable dates (or a skipped "flight" section) left the form's own
    #      flight standing, which is rarely the one month assumed here.
    #
    # Writing the drafted months explicitly -- exactly as the rehydration path
    # does -- makes the two agree by construction; where the draft has no say
    # over the flight, the form's own months are read instead of invented.
    if flight_start and flight_end and "flight" not in skip_sections:
        # Match main()'s own format_flight_label(all_months, all_months) exactly
        # (byte-for-byte, including the single-month case) so the
        # _shared_fields_key we precompute below actually matches what main()
        # derives after rerun -- a mismatch would trigger main()'s own
        # reseed-on-change logic and silently discard these drafted rows.
        draft_months = month_list(flight_start, flight_end)
        updates["active_months"] = draft_months
        all_flight_months = draft_months
    else:
        all_flight_months, draft_months = form_flight_months()
    flight_label = format_flight_label(all_flight_months, draft_months) or "TBD"
    draft_n_months = max(1, len(draft_months))
    touched_sections.add("flight")

    geo = (draft.get("geo") or "").strip()

    specs = draft.get("campaign_specs", {}) or {}
    text_fields = (("goals", "goals_text"), ("audience", "audience_text"), ("budget", "budget_text"),
                   ("placements", "placements_text"), ("timing", "timing_text"))
    for spec_field, widget_key in text_fields:
        values = specs.get(spec_field) or []
        if values:
            updates[widget_key] = "\n".join(values)
    # Geography the notes actually stated wins outright; then the markets this
    # draft resolved; then the originating market label, which is what a
    # proposal naming no target market has always shown.
    geo_bullets = (specs.get("geography")
                   or ([geo] if geo else None)
                   or draft_geo_labels
                   or ([market_label] if market_label else []))
    if geo_bullets:
        updates["geography_text"] = "\n".join(geo_bullets)
    touched_sections.add("specs")

    # Mirror main()'s own default_targeting/default_geo derivation exactly
    # (the same helpers over the same joined text) so downstream keys computed
    # here -- shared_fields_key in particular -- match what main() derives
    # after rerun instead of silently diverging on a bullet-list edge case.
    default_targeting = audience_stack(updates.get("audience_text", ""))
    geo_or_market = geo_column_default(
        draft_geo_labels, updates.get("geography_text", ""), market_label)

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
            internal.append(f"Sport '{s}' not recognized -- skipped.")

    # --- audiences ---
    audiences_in = draft.get("audiences", []) or []
    audience_names = [a.get("segment", "") for a in audiences_in if a.get("segment")]
    matched, unmatched = validate_segments(audience_names)
    # No prompt or schema change asks the model for this, but if a drafted
    # segment ever IS the canonical "(A) AND (B)" form (targeting_groups.py's
    # own expression syntax), validate its individual terms instead of the
    # whole string -- otherwise a real two-term audience is dropped whole as
    # one unrecognized name. No committed fixture contains that string
    # (parse_expression only recognizes it, never produces it from prose), so
    # every frozen fixture's behavior here is unchanged.
    if unmatched:
        still_unmatched = []
        for name in unmatched:
            parsed = tg.parse_expression(name)
            if parsed:
                term_matched, term_unmatched = validate_segments(parsed[0])
                if parsed[0] and not term_unmatched:
                    matched.append(name)
                    continue
            still_unmatched.append(name)
        unmatched = still_unmatched
    if unmatched:
        internal.append(f"Unrecognized audience segment(s) dropped: {', '.join(unmatched)}")
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
        internal.append(
            f"Only one custom audience can run per campaign, so {kept_custom} was kept and "
            f"{len(dropped_custom)} other(s) left out: {', '.join(dropped_custom)}. Swap one in "
            f"if you'd rather use it.")
    matched_audiences = kept_audiences

    # Avails the notes actually supplied. The old rule was "never fabricate
    # avails nobody supplied", which had been implemented as "never accept
    # avails at all" -- so a rep who read the numbers out of the avails
    # system and put them in the notes still got a table of zeroes and a
    # chore. Numbers stated in the notes are ingested; a missing one is left
    # blank and flagged, exactly as before.
    unmatched_with_avails = [a for a in audiences_in
                             if a.get("segment") not in matched
                             and _drafted_avails(a, draft_n_months)[0] is not None]
    if matched_audiences or unmatched_with_avails:
        seed_rows, missing, basis_assumed = [], [], []
        for audience in matched_audiences:
            monthly, note = _drafted_avails(audience, draft_n_months)
            seed_rows.append({"Audience": audience["segment"],
                              "Geo": audience.get("geo") or geo_or_market,
                              AVAILS_COLUMN_MONTHLY: monthly or 0})
            if monthly is None:
                missing.append(audience["segment"])
            elif note:
                basis_assumed.append(audience["segment"])
        # An audience the catalog didn't recognise still had a real number
        # attached to it. Dropping the row silently would throw that number
        # away; the rep's label is kept as typed so they can correct the
        # segment without re-keying the figure.
        for audience in unmatched_with_avails:
            monthly, note = _drafted_avails(audience, draft_n_months)
            label = str(audience.get("segment") or "").strip()
            seed_rows.append({"Audience": label,
                              "Geo": audience.get("geo") or geo_or_market,
                              AVAILS_COLUMN_MONTHLY: monthly or 0})
            internal.append(
                f"\"{label}\" isn't a name from the audience catalog, but the notes gave avails for it, "
                f"so it's in the table with the label as written. Pick the matching catalog segment "
                f"before sending.")
        updates["avails_seed_rows"] = seed_rows
        updates["avails_version"] = st.session_state.get("avails_version", 0) + 1
        # Stored monthly whatever basis the notes used, so the table's own
        # toggle decides how it's shown.
        updates["avails_basis"] = AVAILS_BASIS_MONTHLY
        # A drafted audience becomes a group the same way any other seed row
        # does: one drafted segment -> one single-term group -> one plan
        # line; a canonical "(A) AND (B)" segment (see above) becomes a real
        # two-term group instead of two separate lines. Nothing else about
        # drafting changes -- percent_of_avails/avails_ref still resolve
        # through the same avails_lookup figure as today.
        updates["targeting_groups"] = tg.seed_rows_to_groups(seed_rows, AVAILS_COLUMN_MONTHLY)
        if missing:
            internal.append(
                f"No avails were given for {', '.join(missing)}, so the table shows 0. Pull the real "
                f"numbers from the avails system before sending.")
        if basis_assumed:
            internal.append(
                f"The notes gave avails for {', '.join(basis_assumed)} without saying whether they were "
                f"monthly or for the whole flight. They've been read as monthly -- check that's right.")
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
                internal.append(
                    f'"{opt_in["name"]}" has no overall budget, so only its fixed-amount '
                    f"line(s) could be priced -- anything meant to share a remaining budget "
                    f"was dropped. Say what this plan should cost and re-draft.")
            else:
                internal.append(
                    f'"{opt_in["name"]}" came back with {len(opt_in["lines"])} media plan '
                    f"line(s) but no budget to spend on them, so nothing could be priced and "
                    f"the option was dropped. Say what this plan should cost and re-draft, or "
                    f"add the lines by hand.")

        rows, opt_products, opt_sports, opt_unresolved, opt_drivers = resolve_drafted_lines(
            opt_in["lines"], opt_in["total_budget"], markup,
            flight_label, geo_or_market, default_targeting,
            # The avails this same draft just put in the table, so a reach
            # line can price against them in the same pass. Read from the
            # pending updates rather than session_state: the table hasn't
            # been written yet at this point in the draft.
            avails_by_name=avails_lookup(updates.get("avails_seed_rows")),
            n_months=draft_n_months)
        internal.extend(opt_unresolved)
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
            option = new_plan_option(opt_in["name"], rows, driver=opt_drivers,
                                     breakout=breakout)
            option["dirty"] = [True] * len(rows)  # drafted rows are deliberate, never re-seeded away
            drafted_plan_options.append(option)

    # A drafted plan sitting alongside an imported schedule: the broadcast
    # line's pricing rule is the opposite of every other line's, so say so
    # rather than leaving the reviewer to notice the totals don't move when
    # they flip the agency toggle.
    if st.session_state.get("broadcast_schedule") and agency_involved:
        internal.append(
            "The broadcast schedule line uses its Wide Orbit cost as quoted and is not marked "
            "up by the x1.15 agency uplift -- broadcast is already gross. Every other line on "
            "this plan is marked up as usual.")

    kept_names = {o["name"] for o in drafted_plan_options}
    for opt_in in options_in:
        # Named, so a two-option draft that loses one says which. The
        # budget-specific diagnosis above already fired for the case that
        # causes this most often; this catches the rest (every line dropped
        # as an unrecognized product, an empty line list, and so on).
        if opt_in["name"] not in kept_names and opt_in["total_budget"] > 0:
            internal.append(f'"{opt_in["name"]}" had no usable media plan lines and was dropped.')

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

    # Total TV comes from the draft's own flag rather than from a media plan
    # line, because a broadcast schedule is imported from Wide Orbit rather
    # than drafted -- turning it on here is what lets a seller draft first
    # and upload the schedule second, which is the order they work in.
    if bool(draft.get("total_tv")):
        # Set after the product sweep above, which clears every toggle:
        # Total TV isn't expressible as a media plan line, so it would
        # otherwise be switched straight back off.
        updates["total_tv"] = True
        internal.append(
            "Total TV was switched on because the notes mention a broadcast plan. Upload the "
            "Wide Orbit schedule in the Broadcast schedule panel to add it to the media plan.")

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
            internal.append(
                "Kept your own edits to these section(s) instead of overwriting them with the "
                f"re-draft: {', '.join(preserved)}.")
        touched_sections -= skip_sections

    st.session_state["ai_filled_sections"] = touched_sections
    # Options are variants of one plan, so a per-line note (a negotiated CPM,
    # a dropped $0 line) is usually raised identically by every option. The
    # reviewer needs to read it once; anything genuinely option-specific
    # names its option and so is already distinct.
    st.session_state["draft_unresolved"] = list(dict.fromkeys(unresolved))
    st.session_state["draft_unresolved_internal"] = list(dict.fromkeys(internal))

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
    return cpm, (f'{tactic} is priced at ${cpm:,.2f}, not the ${default_cpm:,.2f} rate card rate. '
                 f"Confirm the negotiated rate.")


def _drafted_avails(audience, n_months=1):
    """(monthly_avails, basis_was_assumed) for one drafted audience entry.

    Returns (None, False) when the notes gave no number -- which is a
    different thing from zero, and has to stay different: zero is a real
    answer the seller can read as "no inventory", while absent means nobody
    has looked yet. Only the absent case is chased up.

    The model reports the basis; the conversion happens here, because asking
    it to divide by a month count is asking it to do arithmetic, which is the
    one thing this whole path is built not to do.
    """
    if not isinstance(audience, dict):
        return None, False
    raw = audience.get("max_avails")
    if raw in (None, "", False):
        return None, False
    try:
        value = int(float(str(raw).replace(",", "").replace("$", "")))
    except (TypeError, ValueError):
        return None, False
    if value < 0:
        return None, False
    basis = str(audience.get("avails_basis") or "").strip().lower()
    if basis in ("flight", "full_flight", "full flight", "campaign", "total"):
        return int(round(value / max(1, int(n_months or 1)))), False
    # Anything else -- including nothing at all -- is read as monthly, which
    # is what the avails system reports. Silence is flagged, not guessed at.
    return value, basis not in ("monthly", "month", "per_month", "per month")


def avails_lookup(seed_rows):
    """{lowercased audience name: monthly avails} for resolving avails_ref.

    Matched case- and whitespace-insensitively because the reference is the
    model repeating a segment name back, and an exact-string match would fail
    on a stray capital.
    """
    lookup = {}
    for row in seed_rows or []:
        name = str(row.get("Audience", "") or "").strip().lower()
        if name:
            lookup[name] = int(row.get(AVAILS_COLUMN_MONTHLY, 0) or 0)
    return lookup


def resolve_drafted_lines(lines_in, total_budget, markup, flight_label, geo_or_market, default_targeting,
                          avails_by_name=None, n_months=1):
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
    # Impressions a reach line fixed directly. A reach line is the one kind
    # whose impressions are the input and whose cost is derived, so it can't
    # go through the cost-driven path the other stages share.
    reach_impressions = {}
    avails_by_name = avails_by_name or {}
    months = max(1, int(n_months or 1))

    # Stage 0: percent_of_avails -- reach, not budget. Resolved first so its
    # cost is known before the remainder is worked out, exactly as a flat
    # amount is.
    #
    # **Full-flight impressions, because that is what this function returns**
    # (the caller divides by the month count for a monthly breakout). Avails
    # are stored monthly, so the flight figure is monthly x months x percent
    # -- which is the same as the percentage of the flight's own avails, and
    # is why the basis toggle can't move this number.
    for i, line in enumerate(lines_valid):
        alloc = line.get("allocation", {}) or {}
        if "percent_of_avails" not in alloc:
            continue
        label = line.get("label") or line.get("product")
        try:
            percent = float(alloc.get("percent_of_avails") or 0)
        except (TypeError, ValueError):
            percent = 0.0
        ref = str(alloc.get("avails_ref") or line.get("audience_track") or "").strip()
        monthly = avails_by_name.get(ref.lower())
        if percent > 100:
            unresolved.append(
                f"The {label} line asks for {percent:g}% of an audience, which is more than all of it. "
                f"It's been left as written -- correct the percentage before sending.")
        if not ref or monthly is None:
            unresolved.append(
                f"The {label} line is meant to reach {percent:g}% of "
                f"{ref or 'an audience'}, but there are no avails for it, so its impressions and cost "
                f"are blank. Fill the avails in and the line will price itself.")
            resolved_amounts[i] = 0
            reach_impressions[i] = 0.0
            continue
        if not monthly:
            unresolved.append(
                f"The {label} line is meant to reach {percent:g}% of {ref}, whose avails are 0, so it "
                f"prices at nothing. Pull the real avails before sending.")
        impressions = monthly * months * (percent / 100.0)
        reach_impressions[i] = impressions
        _, line_cpm = line_product_spec(line["product"])
        line_cpm, _ = _resolve_line_cpm(line, label, line_cpm)
        resolved_amounts[i] = round(cost_from_impressions(impressions, line_cpm, markup))

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
    drivers = []
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
        # A reach line is exempt: when its avails haven't been pulled yet it
        # prices at nothing, and dropping it would throw away the one thing
        # the notes did specify -- which audience and what share of it. It
        # stays on the plan with blanks and a flag, so filling the avails in
        # makes it price itself.
        if not amount and i not in reach_impressions:
            unresolved.append(
                f"The {label or product} line worked out to $0 and was dropped. Add it by hand if it "
                f"should carry a cost, or leave it in the Included with Campaign list.")
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
            # A row and its driver are a pair. This branch appended the row and
            # returned without the driver, so every drafted plan carrying a
            # flat fee came out with its parallel lists one short -- latent
            # until something walked them together, at which point
            # _apply_product_diff raised IndexError and took the page down. A
            # fee's cost IS its specification, so it drives from cost.
            drivers.append(DRIVER_COST)
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
        # Cost is the driver for a drafted row: the resolved allocation is
        # already whole dollars (group-rounded so the split ties exactly to
        # the budget), so deriving impressions from it keeps the plan total
        # on the number the notes asked for. Deriving cost from rounded
        # impressions instead would reintroduce the drift.
        # A reach line's impressions ARE the specification, so they are
        # used as given and the cost follows from them. Every other line
        # is the other way round: its dollar allocation is already
        # group-rounded to tie to the budget, so deriving impressions
        # from it preserves the total.
        line_impressions = (reach_impressions[i] if i in reach_impressions
                            else impressions_from_cost(amount, cpm, markup))
        rows.append({
            "Tactic": tactic, "Flight": flight_label, "Geo": geo_or_market,
            # Fall back to the same per-tactic default the form itself uses
            # (fixed copy for Streaming Retargeting / Live Sports rows,
            # otherwise the Campaign Specs Audience line) rather than always
            # reaching for the audience field.
            "Targeting": audience_track or resolve_row_defaults(
                tactic, geo_or_market, default_targeting, flight_label)["Targeting"],
            "Impressions": line_impressions,
            "CPM": cpm,
            "Type": ROW_TYPE_RATE, "Cost": float(amount),
        })
        # A budget-driven line still gets its reach reported when the notes
        # gave avails for the audience it targets -- "$4,000 against 1.7M
        # avails, what percent does that reach" names a budget, not a reach
        # target, so the line is priced from cost like any other, and this is
        # the only place that percentage against the real avails figure gets
        # computed and shown to the reviewer.
        if i not in reach_impressions:
            track_ref = str(audience_track or "").strip()
            monthly_avails = avails_by_name.get(track_ref.lower())
            if track_ref and monthly_avails:
                flight_avails = monthly_avails * months
                reach_pct = (line_impressions / flight_avails) * 100.0
                unresolved.append(
                    f"The {tactic} line's ${amount:,.0f} reaches about {line_impressions:,.0f} "
                    f"impressions -- roughly {reach_pct:.0f}% of the {flight_avails:,.0f} "
                    f"avails on {track_ref} for the flight.")
        drivers.append(DRIVER_IMPRESSIONS if i in reach_impressions else DRIVER_COST)

    # A plan quoted in reach doesn't have to add up to a budget the notes
    # also mentioned, and when the two disagree the reach is what was
    # actually asked for. Say so rather than silently picking one.
    if reach_impressions and total_budget:
        planned = sum(resolved_amounts.get(i, 0) for i in range(len(lines_valid)))
        if planned and abs(planned - total_budget) >= max(1, round(total_budget * 0.01)):
            unresolved.append(
                f"The notes gave both a reach target and a ${total_budget:,.0f} budget, and they don't "
                f"agree -- the reach comes to ${planned:,.0f}. The plan is built from the reach. "
                f"Confirm which one the client agreed to.")

    return rows, touched_products, touched_sports, unresolved, drivers


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


def _open_builder_group(groups):
    """The group AND/OR extends right now, or None when there isn't one --
    nothing built yet this session, or the group `_builder_open_group_id`
    points at was since deleted or edited away in the Phase 4 grid. Falling
    back to None in that case is deliberate: the next click just starts a
    fresh group instead of raising or reviving a dead one."""
    open_id = st.session_state.get("_builder_open_group_id")
    if not open_id:
        return None
    return next((g for g in groups if g["id"] == open_id), None)


def _add_segment_to_group(segment, action, geo_default):
    """Extend or start a targeting group from the Audience finder's builder
    -- see "The audience half of a group is BUILT" in
    geo_targeting_roadmap.md D for the design this implements.

    `action` is "and", "or" or "separate". AND/OR extend the currently OPEN
    group (the one the last click built or extended); "separate", or a click
    with no open group yet, starts a brand new one-term group and makes IT
    the open one. Mixing AND and OR within one group is refused with a plain
    sentence -- an expression is terms joined by ONE operator, never a
    general boolean tree -- and nothing about the refused click is applied.

    Rewrites `targeting_groups` directly, the same as the Phase 4 grid does,
    and for the same reason: `sync_targeting_groups` projects it down to
    `avails_seed_rows` on the next call, so nothing else has to know groups
    exist. Always reassigns -- a new list, a new/updated group dict -- never
    mutates a group in place.
    """
    groups = list(st.session_state.get("targeting_groups") or [])
    open_group = _open_builder_group(groups)

    if action == "separate" or open_group is None:
        new = tg.new_group([segment], geo_def={"kind": "text", "label": geo_default or ""},
                           color=tg.assign_color(len(groups)))
        groups = groups + [new]
        st.session_state["_builder_open_group_id"] = new["id"]
        queue_group_seed([new["id"]])
    else:
        wanted_op = "AND" if action == "and" else "OR"
        if open_group.get("op") and open_group["op"] != wanted_op:
            st.session_state["_builder_refused"] = (
                f"This group is already \"{tg.audience_label(open_group)}\" joined by "
                f"{open_group['op']} -- an audience is joined by ONE operator. Use \"New "
                f"group\" to start a separate one with \"{segment}\" instead.")
            st.rerun()
            return
        updated = dict(open_group)
        updated["terms"] = list(open_group["terms"]) + [segment]
        updated["op"] = wanted_op if len(updated["terms"]) > 1 else None
        groups = [updated if g["id"] == open_group["id"] else g for g in groups]

    st.session_state["targeting_groups"] = groups
    st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1

    # The one-custom rule is allow-with-a-warning here, not a refusal --
    # decided during phase-1/2 planning. Set for THIS click's own feedback;
    # render_review_list also recomputes this live from targeting_groups on
    # every run, so the same overage is visible at generate time even if
    # this toast scrolls past unread.
    catalog = load_audience_catalog()
    rfp_map = dict(zip(catalog["segment"], catalog["rfp_selectable"]))
    count = tg.custom_segment_count(groups, rfp_map)
    if count > 1:
        st.session_state["_builder_warning"] = (
            f"{count} custom (non-RFP-selectable) audiences are now in play across your "
            f"targeting groups -- only one is allowed per campaign. Review before generating.")

    st.rerun()


# ---------------------------------------------------------------------------
# The avails PDF importer (geo_targeting_roadmap.md F)
# ---------------------------------------------------------------------------
# Attribution products this app can safely infer from Salesforce's own
# Attribution text -- deliberately ONE mapping, not a guessed keyword list,
# because "Premion Website" is already a standard included item with no
# toggle of its own, and nothing else in the four real samples checked names
# a product this app's OTHER toggles (first_party_data, sales_attribution,
# brand_lift, commercial_production) could be confidently matched to. See
# module docstring section below: this is a floor, and an unmapped mention is
# surfaced for the seller to check by hand rather than silently ignored.
_AVAILS_ATTRIBUTION_MAP = {"reach extension": "linear_reach_extension"}


def _avails_import_field(key, new_value, default_value, is_stated=None):
    """Apply-or-conflict for one header field, following the precedence
    every importer in this app already uses: rep edits > this document >
    a draft > defaults. "Safe to apply" means the widget still shows either
    the untouched default or the value THIS SAME IMPORTER last wrote (a
    re-import of the same or a corrected document); anything else means a
    rep (or a draft) has since set something of their own, which an avails
    document -- a real, external, but SECOND-HAND source for these
    particular fields -- must never silently overwrite.

    Returns ("apply", value) or ("conflict", current_value, new_value).
    `is_stated` defaults to `bool(new_value)` -- right for a string/date
    field, where empty genuinely means "the document didn't say" -- but is
    passed explicitly for a BOOLEAN field like agency_involved, where False
    is a real, meaningful answer ("Direct - No Agency" on the document),
    not an absent one; without this, a document correctly saying "no
    agency" could never be applied OR flagged as a conflict, because a bare
    `if not new_value` would read False as "nothing to say" the same way it
    reads an empty string.
    """
    if is_stated is None:
        is_stated = bool(new_value)
    if not is_stated:
        return ("skip", None)
    written = st.session_state.get("_avails_import_written", {})
    current = st.session_state.get(key)
    if current in (default_value, "", None, written.get(key)):
        return ("apply", new_value)
    if current == new_value:
        return ("skip", None)
    return ("conflict", current, new_value)


def apply_avails_import(document):
    """One parsed AvailsDocument -> real targeting_groups plus a report of
    what was and wasn't applied. Never touches session_state's widget keys
    directly for anything OTHER than what it decides to apply -- the caller
    reruns, so a widget already rendered this pass keeps showing its current
    value regardless; this only sets up what the NEXT render shows.

    Geography and audience are resolved here, in ONE place, through the
    SAME functions everything else in this app already uses --
    `resolve_group_geography` (never geo_resolver directly) and the
    catalog's own segment set (never a second notion of what a valid
    segment is) -- so an imported group is indistinguishable from one built
    by hand through the finder or the geo-definition expander.
    """
    unresolved = []
    new_groups = []
    catalog = load_audience_catalog()
    valid_segments = set(catalog["segment"])
    profiles, _ = load_market_profiles()
    existing = st.session_state.get("targeting_groups") or []

    for i, g in enumerate(document.groups):
        terms, op = tg.terms_from_audience_text(g.audience_text)
        for term in terms:
            if term not in valid_segments:
                unresolved.append(
                    f"\"{term}\" (from {document.source_name}) isn't an exact match in the "
                    f"audience catalog -- kept as a custom segment; check the spelling against "
                    f"the catalog if it should have matched.")

        if g.geo_kind == avails_pdf_import.GEO_KIND_DMA:
            key, candidates = market_profiles.match_market(g.geo_name, profiles)
            if key is None:
                if candidates:
                    options = ", ".join(
                        next((r.get("label") for r in profiles if r.get("key") == c), c)
                        for c in candidates[:4])
                    unresolved.append(
                        f"\"{g.geo_name}\" (from {document.source_name}) could be more than one "
                        f"market ({options}) -- resolve it by hand in that group's Markets cell.")
                else:
                    unresolved.append(
                        f"\"{g.geo_name}\" (from {document.source_name}) isn't a recognized "
                        f"market -- resolve it by hand in that group's Markets cell.")
                geo_def, zips, markets, notes = {"kind": "text", "label": g.geo_name}, [], [], []
            else:
                geo_def, zips, markets, notes, geo_unresolved = resolve_group_geography(
                    GEO_MODE_MARKETS, markets_picked=[key])
                unresolved += [f"{document.source_name}: {n}" for n in geo_unresolved]
        elif g.geo_kind == avails_pdf_import.GEO_KIND_COUNTY:
            geo_def, zips, markets, notes, geo_unresolved = resolve_group_geography(
                GEO_MODE_COUNTIES, counties_text=g.county_list)
            unresolved += [f"{document.source_name}: {n}" for n in geo_unresolved]
        elif g.geo_kind == avails_pdf_import.GEO_KIND_RADIUS and g.radius_origin:
            # A real origin exists -- resolve through the SAME Radius mode a
            # rep would use by hand, so this group is byte-identical to one
            # built that way (geo_targeting_roadmap.md's own fixture for
            # this exact document does this, and is what "reaches the same
            # state as entering the document by hand" is checked against).
            geo_def, zips, markets, notes, geo_unresolved = resolve_group_geography(
                GEO_MODE_RADIUS, radius_centers_text=g.radius_origin, radius_miles=g.radius_miles)
            unresolved += [f"{document.source_name}: {n}" for n in geo_unresolved]
        else:
            # Named zip option, or a radius with NO bracketed origin (one
            # real sample has exactly this) -- there's no center to resolve
            # a radius from, so this is the same fallback a rep is forced
            # into by hand: the document's own zip list, entered as Zips.
            zips = geo_resolver.parse_zip_list(",".join(g.zips))
            geo_def = {"kind": "zips", "zips": zips}
            market_result = geo_resolver.zips_to_markets(zips)
            markets = sorted(market_result.resolved.keys())
            notes = list(market_result.notes)
            unresolved += [f"{document.source_name}: {u}" for u in market_result.unresolved]

        group = tg.new_group(terms, op=op, geo_def=geo_def,
                             avails_monthly=0,  # set below, once the flight (and so n_months) is known
                             color=tg.assign_color(len(existing) + len(new_groups)))
        group["resolved_zips"] = zips
        group["resolved_markets"] = markets
        group["_avails_import_impressions"] = g.impressions   # full-flight; see caller
        new_groups.append(group)

    agency_involved = bool(document.agency) and "no agency" not in document.agency.lower()
    field_updates = {}
    conflicts = []
    for key, value, default, is_stated in (
        ("client_name", document.advertiser, "Acme Test Co", None),
        ("agency_involved", agency_involved, False, bool(document.agency)),
        ("flight_start", document.flight_start, DEFAULT_FLIGHT_START, None),
        ("flight_end", document.flight_end, DEFAULT_FLIGHT_END, None),
    ):
        outcome, *rest = _avails_import_field(key, value, default, is_stated)
        if outcome == "apply":
            field_updates[key] = rest[0]
        elif outcome == "conflict":
            conflicts.append(
                f"{key.replace('_', ' ')} is already set to {rest[0]!r}; the avails document "
                f"says {rest[1]!r}. Left as-is -- update it by hand if the document is right.")

    attribution_lower = document.attribution_text.lower()
    for phrase, toggle_key in _AVAILS_ATTRIBUTION_MAP.items():
        if phrase in attribution_lower and not st.session_state.get(toggle_key):
            field_updates[toggle_key] = True
    if document.attribution_text:
        unresolved.append(
            f"{document.source_name} lists Attribution: \"{document.attribution_text}\" -- "
            f"an avail's attribution is typically a SUBSET of what's actually being sold; "
            f"confirm the Section D toggles cover everything, don't just match the avail.")

    return new_groups, {
        "unresolved": unresolved, "field_updates": field_updates, "conflicts": conflicts,
        "rfpid": document.rfpid, "n_groups": len(new_groups),
        "total_impressions": document.total_impressions,
        "parsed_total": sum(g.impressions for g in document.groups),
    }


def apply_pending_avails_import_fields():
    """Apply whatever `_finish_avails_import` queued last run, before ANY
    widget on this page is instantiated -- the same "keyed widget's own
    session_state wins over value=" rule as everywhere else in this file,
    but sharper here: `st.session_state[key] = value` for a key ALREADY
    bound to a widget instantiated earlier in the SAME run raises outright
    (confirmed -- this is exactly how the D2 entry point's own writes to
    `client_name` were first found breaking, since Section A's widget
    renders before D2 ever runs). Queuing the write and applying it here,
    before Section A's first widget, is what makes it safe regardless of
    which entry point (D2, deep in the form; the review list, near the top)
    queued it.
    """
    pending = st.session_state.pop("_avails_import_pending_fields", None)
    if not pending:
        return
    for key, value in pending.items():
        st.session_state[key] = value


def _finish_avails_import(document, new_groups, report):
    """Commit an apply_avails_import() result to session_state: queue the
    header-field updates for next run (see apply_pending_avails_import_fields
    for why this can't write them directly), spread each new group's
    full-flight impressions into avails_monthly against the CURRENT flight,
    append the groups, and record the report for the uploader to display.

    Flight dates are read AS THEY ARE RIGHT NOW for the avails_monthly
    spread, deliberately not the just-queued flight update -- that new
    flight won't be live until the field-update queue is applied, next run,
    same as the rest of this function's queued writes. A flight-date
    conflict (the rep already set different ones) already left the OLD
    flight in place, which is exactly what should drive this spread in that
    case, so reading current state here is correct either way, not just
    the common one.
    """
    written = dict(st.session_state.get("_avails_import_written", {}))
    written.update(report["field_updates"])
    st.session_state["_avails_import_written"] = written
    if report["field_updates"]:
        st.session_state["_avails_import_pending_fields"] = report["field_updates"]

    _, active_months = form_flight_months()
    n_months = max(1, len(active_months))
    for group in new_groups:
        full_flight = group.pop("_avails_import_impressions", 0)
        group["avails_monthly"] = int(round(full_flight / n_months))

    st.session_state["targeting_groups"] = (st.session_state.get("targeting_groups") or []) + new_groups
    queue_group_seed([g["id"] for g in new_groups])
    st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
    history = list(st.session_state.get("avails_import_history") or [])
    history.append(report["rfpid"])
    st.session_state["avails_import_history"] = history
    st.session_state["avails_import_report"] = report
    st.session_state["avails_import_error"] = None


def render_avails_pdf_uploader(key_suffix, prompt):
    """One avails-PDF upload widget, shared verbatim by both entry points
    (beside the D2 avails table, and inline in the review list) -- same
    import, same precedence, same report, per geo_targeting_roadmap.md F's
    own "second entry point, not a second implementation" requirement.

    `key_suffix` keeps the two entry points' widget keys distinct (a rep
    could conceivably want to use either); `prompt` is the copy shown above
    the uploader, which is the only thing that actually differs between them.
    """
    upload = st.file_uploader("Avails PDF (from Salesforce)", type=["pdf"],
                              key=f"avails_pdf_upload_{key_suffix}", help=prompt)
    injected = test_mode_upload(f"avails_pdf_upload_path_{key_suffix}")
    if injected is not None:
        upload = injected

    loaded_key = f"avails_pdf_loaded_{key_suffix}"
    if upload is not None and st.session_state.get(loaded_key) != upload.name:
        target = db.scratch_dir("premion_avails_uploads") / upload.name
        target.write_bytes(upload.getvalue())
        try:
            document = avails_pdf_import.parse_avails_pdf(str(target), upload.name)
        except avails_pdf_import.AvailsParseError as exc:
            st.session_state["avails_import_error"] = str(exc)
            st.session_state["avails_import_report"] = None
        else:
            if document.rfpid and document.rfpid in (st.session_state.get("avails_import_history") or []):
                st.session_state["avails_import_error"] = (
                    f"{document.rfpid} was already imported this session. Re-uploading it would "
                    f"add duplicate targeting groups -- remove the ones already on the table first "
                    f"if you meant to re-import.")
            else:
                new_groups, report = apply_avails_import(document)
                _finish_avails_import(document, new_groups, report)
        st.session_state[loaded_key] = upload.name
        st.rerun()

    error = st.session_state.get("avails_import_error")
    if error:
        st.error(error)

    report = st.session_state.get("avails_import_report")
    if report:
        st.caption(f"✅ {report['n_groups']} targeting group(s) added from {report['rfpid'] or 'the upload'} "
                   f"-- {report['parsed_total']:,} impressions "
                   f"({'matches' if report['parsed_total'] == report['total_impressions'] else 'does NOT match'} "
                   f"the document's own stated total of {report['total_impressions']:,}).")
        for note in report["conflicts"]:
            st.warning(note)
        for note in report["unresolved"]:
            st.caption(f"ℹ️ {note}")


def render_audience_finder(avails_df, geo_default, vertical_key=None):
    """Section D2's 'Audience finder': browse/search the catalog, or describe
    the client/campaign and let Claude suggest segments. Each result offers
    AND / OR / "New group" instead of one Add button -- see
    "The audience half of a group is BUILT" in geo_targeting_roadmap.md D.

    Wrapped in an expander here because it's a side tool inside a long form;
    the standalone page calls the same body without one. Same component, two
    entry points.
    """
    with st.expander("🔍 Audience finder", expanded=False):
        _audience_finder_body(avails_df, geo_default, vertical_key)


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


def render_zip_map_builder_page():
    """The Zip/map builder (roadmap §E) -- its own page, like Audience
    finder, reading the CURRENT proposal's `targeting_groups` directly
    rather than taking its own input. "Draws what already exists": a group
    only shows up here once the geo-definition expander (Section D2 on the
    Build page) has resolved it, and nothing entered on this page writes
    back to a group -- there's nothing to enter. Also reachable mid-build
    from a group's own geo expander (`_goto_zip_map_builder`).

    Outputs both a picture and the resolved zip list -- the list is what a
    planner actually pulls avails against, so it gets equal billing with
    the map, not a footnote under it: copyable (a plain text_area) and
    exportable (a download button) per group.
    """
    st.header("Zip/map builder")
    st.caption("Every targeting group that's been resolved to real zips (Section D2's geo "
               "expander, any mode) renders here on one map, each in its own color, with the "
               "same zip list a planning avails pull needs. Nothing on this page changes a "
               "group -- resolve geography in Section D2, then come back to look or export.")

    groups = st.session_state.get("targeting_groups") or []
    plottable = targeting_map.groups_with_zips(groups)
    if not plottable:
        st.info("No targeting group has been resolved to real zips yet. Go to **Build a "
                 "proposal → Section D2 → Geography** for a group, pick a mode, and Resolve -- "
                 "it shows up here as soon as it has zips.")
        return

    png = targeting_map.render_map(plottable, width_px=1000, height_px=620,
                                   label_for=_market_display_name)
    if png:
        st.image(png, use_container_width=True)
        st.caption("This is exactly the picture that goes onto the targeting slide when you "
                   "Generate -- not a separate preview that might disagree with it.")

    st.divider()
    st.subheader("Resolved zips, per group")
    for group in plottable:
        label = tg.audience_label(group) or "(untitled)"
        zips = group.get("resolved_zips") or []
        with st.expander(f"{label} -- {len(zips):,} zip(s)", expanded=False):
            zip_text = "\n".join(zips)
            st.text_area("Zip list", value=zip_text, height=140, key=f"map_ziptext_{group['id']}",
                        label_visibility="collapsed")
            st.download_button(
                "Download as .txt", data=zip_text, file_name=f"{label.replace(' ', '_')}_zips.txt",
                mime="text/plain", key=f"map_zipdl_{group['id']}")


def _audience_finder_body(avails_df, geo_default, vertical_key=None):
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

    if can_add:
        # Feedback from the LAST click, shown once here rather than inline
        # at the button -- the click that set these triggered a rerun, so by
        # the time this renders again it's the previous action's result, not
        # a stale message left over from an earlier one.
        refused = st.session_state.pop("_builder_refused", None)
        if refused:
            st.error(refused)
        builder_warning = st.session_state.pop("_builder_warning", None)
        if builder_warning:
            st.warning(builder_warning)

        open_group = _open_builder_group(st.session_state.get("targeting_groups") or [])
        if open_group:
            st.info(f"Building: **{tg.audience_label(open_group)}** -- AND/OR adds to it, "
                    f"\"New group\" starts a separate one.")
            # Booking evidence, phase 7 of the targeting-groups roadmap
            # ("Show booking evidence while building" in geo_targeting_roadmap.md
            # D): evidence for the group as it stands right now, so a rep sees
            # it change with every AND/OR click, not just once at the end.
            # `evidence_lines` already carries the settled panel order (exact
            # match only when true, component familiarity, the weakest pair,
            # suggested pairings) and the exact wording rules, so this just
            # renders what it returns.
            open_terms = open_group.get("terms") or []
            if open_terms:
                lines = audience_evidence.evidence_lines(
                    audience_evidence.evidence_for(open_terms, load_audience_index()))
                if lines:
                    with st.container(border=True):
                        st.caption("**Booking evidence**")
                        for line in lines:
                            st.caption(line)

    mode = st.radio("Mode", ["Browse / search", "Suggest"], horizontal=True, key="finder_mode")

    if mode == "Browse / search":
        # all_categories expands a comma-joined multi-category cell into its
        # individual values ("MOVERS, LIFESTAGE" offers both "MOVERS" and
        # "LIFESTAGE" as picks, never a combined "MOVERS, LIFESTAGE" option
        # nobody would type), and category_matches is the membership test
        # that agrees with it -- a segment picked under EITHER of its own
        # categories, matched once, never duplicated in the results below.
        cat_options = ["All"] + all_categories(catalog)
        picked_cat = st.selectbox("Category", cat_options, key="finder_category")
        search_text = st.text_input("Search by name", key="finder_search")

        filtered = catalog
        if picked_cat != "All":
            filtered = filtered[filtered["category"].apply(lambda c: category_matches(c, [picked_cat]))]
        if search_text.strip():
            filtered = filtered[filtered["segment"].str.contains(search_text.strip(), case=False, na=False)]
        total_matches = len(filtered)
        filtered = prioritize_catalog(filtered, vertical_hint).head(50)
        if vertical_hint and picked_cat == "All":
            st.caption(f"{total_matches} segment(s) match; showing 50 "
                       f"(categories relevant to the selected vertical first, then by impressions)")
        else:
            st.caption(f"{total_matches} segment(s) match; showing 50 (by impressions)")

        header_cols = st.columns([3, 1.3, 2, 1.2, 1, 0.7, 0.7, 1])
        for col, label in zip(header_cols,
                              ["Segment", "Category", "Subcategory", "Status", "Impressions", "", "", ""]):
            col.caption(f"**{label}**")
        for _, row in filtered.iterrows():
            cols = st.columns([3, 1.3, 2, 1.2, 1, 0.7, 0.7, 1])
            cols[0].write(row["segment"])
            cols[1].write(row["category"])
            cols[2].write(row["subcategory"] or "--")
            cols[3].write("RFP" if row["rfp_selectable"] else "Custom")
            cols[4].write(f"{row['impressions']:,}")
            if can_add:
                seg = row["segment"]
                if cols[5].button("AND", key=f"finder_and_{seg}",
                                  help="Narrow the currently open group with this segment"):
                    _add_segment_to_group(seg, "and", geo_default)
                if cols[6].button("OR", key=f"finder_or_{seg}",
                                  help="Add this as an alternative on the currently open group"):
                    _add_segment_to_group(seg, "or", geo_default)
                if cols[7].button("New group", key=f"finder_new_{seg}",
                                  help="Add as a separate targeting group, with its own avails "
                                       "row and its own campaign line"):
                    _add_segment_to_group(seg, "separate", geo_default)

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

            # Same lookup Phase 4's D2 warning and render_review_list use --
            # DISTINCT segments across every group, not a row count, so a
            # custom segment reused across two groups isn't double-counted.
            existing_custom = (tg.custom_segment_count(st.session_state.get("targeting_groups") or [], rfp_map)
                               if can_add else 0)
            rec_custom = sum(1 for r in suggestions if not rfp_map.get(r["segment"], True))
            if existing_custom + rec_custom > 1:
                st.warning(
                    f"This recommendation set includes {rec_custom} custom (non-RFP-selectable) audience(s), "
                    f"and your targeting groups already have {existing_custom} -- only one custom audience is "
                    f"allowed per campaign. Review before adding all of them.")

            header_cols = st.columns([2.6, 2.6, 1.3, 1, 1, 0.7, 0.7, 1])
            for col, label in zip(header_cols,
                                  ["Segment", "Rationale", "Category", "Status", "Impressions", "", "", ""]):
                col.caption(f"**{label}**")
            for rec in suggestions:
                seg = rec["segment"]
                match = catalog[catalog["segment"] == seg]
                if match.empty:
                    continue
                cat_row = match.iloc[0]
                cols = st.columns([2.6, 2.6, 1.3, 1, 1, 0.7, 0.7, 1])
                cols[0].write(seg)
                cols[1].write(rec.get("rationale", ""))
                cols[2].write(cat_row["category"])
                cols[3].write("RFP" if cat_row["rfp_selectable"] else "Custom")
                cols[4].write(f"{cat_row['impressions']:,}")
                if can_add:
                    if cols[5].button("AND", key=f"finder_and_suggest_{seg}",
                                      help="Narrow the currently open group with this segment"):
                        _add_segment_to_group(seg, "and", geo_default)
                    if cols[6].button("OR", key=f"finder_or_suggest_{seg}",
                                      help="Add this as an alternative on the currently open group"):
                        _add_segment_to_group(seg, "or", geo_default)
                    if cols[7].button("New group", key=f"finder_new_suggest_{seg}",
                                      help="Add as a separate targeting group, with its own avails "
                                           "row and its own campaign line"):
                        _add_segment_to_group(seg, "separate", geo_default)


def lines_to_bullets(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


def first_line(text):
    bullets = lines_to_bullets(text)
    return bullets[0] if bullets else ""


def audience_stack(text):
    """The Campaign Specs Audience field as one line's Targeting copy.

    Every bullet, joined -- not first_line(). An audience is normally defined
    by several attributes at once ("adults 35+", "homeowners", "higher
    income", "researching cosmetic procedures"), and taking only the first
    put one attribute in the Targeting column while the Campaign Specs slide
    beside it listed all four. Silent truncation is the worse failure of
    the two available: an over-full cell is visible and a rep can trim it,
    whereas a missing attribute looks exactly like a deliberate choice.

    Geography deliberately still uses first_line() -- the Geo column is one
    market, not a stack.
    """
    return ", ".join(lines_to_bullets(text))


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


def form_flight_months():
    """(all_months, active_months) as main() will derive them from whatever the
    form currently holds.

    A mirror of main()'s own derivation, for callers that run *before* the
    flight widgets have rendered and need to know the month count those widgets
    are going to produce. The month count that spreads a budget has to be the
    month count that totals it, and anything computing its own is guessing.
    Falls back to the widget defaults when session_state is still empty, since
    that is what the widgets themselves will fall back to.
    """
    start = st.session_state.get("flight_start") or DEFAULT_FLIGHT_START
    end = st.session_state.get("flight_end") or DEFAULT_FLIGHT_END
    if not (isinstance(start, date) and isinstance(end, date)):
        return [], []
    all_months = month_list(start, end)
    stored = st.session_state.get("active_months")
    if stored is None:
        return all_months, all_months
    # Same rule as main(): an empty intersection means the stored choice
    # belongs to a different flight and the whole range is reselected, while a
    # partial overlap is real custom flighting and is kept.
    kept = [m for m in stored if m in all_months]
    return all_months, (kept or all_months)


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


# ---------------------------------------------------------------------------
# The broadcast line (Total TV)
# ---------------------------------------------------------------------------
BROADCAST_TACTIC_SUFFIX = "Broadcast Schedule (:15/:30 Broadcast Video)"
# How a broadcast row is recognized after the fact -- by its tactic name,
# which survives rehydration from history, a draft, and the grid's own
# editing. Renaming the tactic past this marker forfeits the gross exemption
# below, which is why the row carries a visible caption saying so.
#
# Deliberately the full suffix, not just "Broadcast Schedule": that shorter
# string is also the `broadcast_tv` product's own label, and matching it
# would exempt a hand-entered rate-card broadcast line from the markup too.
# The gross rule is about Wide Orbit's already-gross costs, not about the
# word "broadcast".
BROADCAST_TACTIC_MARKER = BROADCAST_TACTIC_SUFFIX

# The rate-card product whose seeded line an imported schedule replaces --
# they describe the same buy, and leaving both would double-count it.
BROADCAST_PRODUCT_LABEL = "Broadcast Schedule"

# Longest first: "FOX43" must win over a bare "FOX" if one is ever added.
STATION_MARKETS = {"WUSA9": "DC", "WUSA": "DC", "WPMT": "Harrisburg", "FOX43": "Harrisburg"}
MARKET_GEO = {"DC": "Washington DC DMA", "Harrisburg": "Harrisburg DMA"}

# **Settled policy, not a default.** A Wide Orbit schedule's cost is already
# gross -- the station quotes it that way and the agency commission is inside
# it, as the Planner PDF's own "Total Cost / Agency Commission @ 15% / Net
# Cost" block spells out. Running the x1.15 agency markup over it would bill
# the commission twice. So the broadcast line uses the WO cost verbatim and
# is excluded from the markup permanently, whatever the agency toggle says.
BROADCAST_EXCLUDED_FROM_AGENCY_MARKUP = True


def is_broadcast_row(row):
    return BROADCAST_TACTIC_MARKER in str(row.get("Tactic", "") or "")


def row_markup(row, markup):
    """The markup that applies to one row. Every derivation goes through
    this, so the broadcast exemption can't be missed at one call site."""
    if BROADCAST_EXCLUDED_FROM_AGENCY_MARKUP and is_broadcast_row(row):
        return 1.0
    return markup


def station_market(station):
    """(market, warning) for a station call sign. Never guesses.

    An unknown station means the geo and the deck's branding would both be
    wrong, and a wrong DMA on a media plan is the kind of error a client
    notices -- so it's reported rather than defaulted.
    """
    name = (station or "").upper().strip()
    for call_sign in sorted(STATION_MARKETS, key=len, reverse=True):
        if call_sign in name:
            return STATION_MARKETS[call_sign], None
    return None, (f"This schedule is for station \"{station or 'unknown'}\", which isn't one of "
                  f"the stations this app knows ({', '.join(sorted(STATION_MARKETS))}). Set the "
                  f"market and geo on the broadcast line by hand before generating.")


def broadcast_targeting_copy(schedule, description, monthly, n_months):
    """The Targeting cell for the broadcast line.

    Built from the schedule itself, with the rep's own description first when
    they've written one -- they know what the buy is *for* in a way the
    export doesn't say.
    """
    parts = []
    if description and description.strip():
        parts.append(description.strip())
    else:
        spots = schedule.summary.total_spots
        if monthly and n_months > 1:
            parts.append(f"{round(spots / n_months):,}x Commercials Per Month")
        else:
            parts.append(f"{spots:,}x Commercials")
        programs = [r.program for r in schedule.rows if r.program]
        if programs:
            parts.append(", ".join(dict.fromkeys(programs))[:80])
        weeks = len(schedule.grid_weeks)
        if weeks:
            parts.append(f"{weeks} Week{'s' if weeks != 1 else ''}")
    return ", ".join(parts)


def broadcast_row_for(schedule, description, geo_default, monthly, n_months, flight_label):
    """The media plan row for an imported schedule. (row, warning).

    Impressions and cost come from the Wide Orbit summary rather than being
    re-derived, and the CPM is computed from them -- so the media plan and
    the schedule slides are reading the same three numbers, which is exactly
    the agreement a client would check.
    """
    summary = schedule.summary
    market, warning = station_market(summary.station)
    geo = MARKET_GEO.get(market) or geo_default

    # **Full flight is always the Wide Orbit totals, verbatim.** Monthly
    # divides by the months the SCHEDULE runs in, worked out from the week
    # start dates -- not by the plan's month count, which is a different
    # span entirely and gave a $28,000 four-week May buy a $9,333 "monthly"
    # figure against a three-month plan flight it had nothing to do with.
    impressions = summary.impressions
    cost = summary.gross_cost
    schedule_months = schedule.active_month_count()
    if monthly:
        if schedule_months:
            impressions /= schedule_months
            cost /= schedule_months
        elif n_months > 1:
            # No resolvable week dates: an even split is the only option
            # left, and the caller flags it rather than presenting it as
            # though it were derived.
            impressions /= n_months
            cost /= n_months

    station = (summary.station or "Broadcast").upper()
    row = {
        "Tactic": f"{station} {BROADCAST_TACTIC_SUFFIX}",
        "Flight": flight_label,
        "Geo": geo,
        "Targeting": broadcast_targeting_copy(schedule, description, monthly, n_months),
        "Impressions": float(round(impressions)),
        # Derived, not taken from the export: WO's CPM is rounded to cents and
        # rounding it back out would leave the plan a few dollars off the
        # schedule slide.
        "CPM": (cost / impressions * 1000) if impressions else 0.0,
        "Type": ROW_TYPE_RATE,
        "Cost": float(round(cost)),
    }
    return row, warning


def recompute_row(row, driver, markup):
    """Recompute whichever side of a rate row isn't driving. Mutates and
    returns the row. Flat-fee rows have neither side -- their Cost is the fee
    itself, and Impressions/CPM are ignored entirely."""
    if is_flat_fee_row(row):
        return row
    # The one place markup is applied to a rate row, so the broadcast
    # exemption is enforced here rather than at each caller.
    effective = row_markup(row, markup)
    if driver == DRIVER_COST:
        row["Impressions"] = impressions_from_cost(row.get("Cost"), row.get("CPM"), effective)
    else:
        row["Cost"] = cost_from_impressions(row.get("Impressions"), row.get("CPM"), effective)
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


def fixed_targeting_copy(tactic):
    """The product's own fixed Targeting copy for this tactic name, matched
    by prefix (longest first, since product labels are prefixes of one
    another) -- or None when the tactic has no fixed copy and falls through
    to the audience stack.

    The single point of truth `resolve_row_defaults` uses for priority (1) of
    its own three-source order; `merge_plan_rows` uses it too, so a Streaming
    Retargeting or Live Sports row keeps its fixed copy through a merge
    exactly like it does through a shared-field re-seed or a split, instead
    of merge alone building the cell from the merged groups' audience labels
    and silently overwriting it -- the bug this function exists to close off
    at every call site, not just the one it was first found in.
    """
    tactic = str(tactic or "")
    for label in sorted(TARGETING_COPY_BY_LABEL, key=len, reverse=True):
        if tactic.startswith(label):
            return TARGETING_COPY_BY_LABEL[label]
    return None


def resolve_row_defaults(tactic, default_geo, default_targeting, flight_label,
                         current=None):
    """What a row's Flight/Geo/Targeting should be right now, given its
    Tactic name -- used both at initial seed time and to soft-update
    not-yet-edited rows when the shared form fields change.

    Targeting has three sources, and which one applies is a property of the
    LINE TYPE rather than one global rule:

      1. The product's own `targeting_copy` from the rate card, when it has
         one (Streaming Retargeting, every Live Sports package). Matched by
         tactic-name prefix so a row whose name has a draft label appended
         ("... - Display - Commercial") still picks it up, longest label
         first because product labels are prefixes of one another.
      2. The row's OWN copy, kept as-is, when it was derived from something
         this function can't reproduce -- the imported broadcast line, whose
         Targeting is built from the Wide Orbit schedule (commercial count,
         programs/dayparts, cadence).
      3. Otherwise the Campaign Specs Audience stack.

    (2) is why `current` exists. Broadcast is bought by program and daypart,
    not by audience segment, and the audience stack overwrote the
    schedule-derived summary on any shared-field edit -- the line came out
    reading "Homeowners with higher household income (100K+)..." on a real
    demo deck. Its Geo is held for the same reason and it is the more
    dangerous of the two: that value is derived from the station's call sign
    (WUSA -> Washington DC DMA) and is deliberately never guessed, so
    replacing it with the form's market would put a wrong DMA on a client's
    media plan whenever the station and the selected market disagree.
    """
    targeting = fixed_targeting_copy(tactic) or default_targeting
    geo = default_geo
    if current is not None and is_broadcast_row({"Tactic": tactic}):
        targeting = current.get("Targeting") or targeting
        geo = current.get("Geo") or geo
    return {"Flight": flight_label, "Geo": geo, "Targeting": targeting}


def quick_add_rows(product_labels, audiences, geos, flight_label,
                   default_targeting, geo_default, combine=False):
    """Build media plan rows from the Add-lines panel's picks.

    The CROSS PRODUCT, as separate lines, is the default: one product across
    three markets is three lines, two products across three markets is six.
    That is how these plans are actually built and priced -- a market is a
    line a client can see and a rep can move money between. `combine` folds
    the geos onto one line per product, for the case where several markets
    sell as a single campaign line, matching the avails table's own toggle.

    Everything else about a row comes from the paths that already own it:
    line_product_spec for the label and rate-card CPM (so the panel, Section
    C's own seeding and the AI draft can't price the same product three
    different ways), and resolve_row_defaults for Flight/Geo/Targeting --
    which is also what keeps a product with its own targeting_copy (Streaming
    Retargeting, every sports package) from being handed the audience stack.

    An empty product list yields nothing: a line has to be something before
    it can be anywhere.
    """
    if not product_labels:
        return []

    geo_values = list(geos) if geos else []
    if combine and geo_values:
        geo_values = [", ".join(geo_values)]
    if not geo_values:
        geo_values = [geo_default]
    audience_values = list(audiences) if audiences else [None]

    rows = []
    for label in product_labels:
        cpm = quick_add_cpm(label)
        for audience in audience_values:
            for geo in geo_values:
                defaults = resolve_row_defaults(
                    label, geo, audience or default_targeting, flight_label)
                rows.append({
                    "Tactic": label, "Flight": defaults["Flight"],
                    "Geo": defaults["Geo"], "Targeting": defaults["Targeting"],
                    "Impressions": 0.0, "CPM": cpm,
                    "Type": ROW_TYPE_RATE, "Cost": 0.0,
                })
    return rows


def quick_add_cpm(product_label):
    """The rate-card CPM for a label the panel offered.

    The panel lists product NAMES (what a rep recognises) while the rate card
    is keyed by product key, so this walks back. A label typed as free text
    matches nothing and prices at 0.0, which is correct and visible -- the rep
    types the rate in, exactly as they would for any product not on the card.
    """
    for key, (label, cpm) in _quick_add_catalog().items():   # noqa: B007
        if label == product_label:
            return cpm
    return 0.0


@st.cache_data(ttl=600, show_spinner=False)
def _quick_add_catalog():
    """{product_key: (label, cpm)} for everything the panel can offer.

    Reads the live rate card, so a rate edited in Supabase reaches the panel
    on the same TTL as everywhere else. Excludes the keys that aren't media
    lines (dynamic_creative is a flat build fee with no CPM, and a drafted
    line naming it would find no rate and produce the $0 row this app already
    has a rule against).
    """
    products, sport_cpm, _sport_labels, _, _ = load_rate_card()
    keys = [k for k in products if k not in NON_LINE_PRODUCT_KEYS]
    keys += [f"{SPORT_PRODUCT_PREFIX}{k}" for k in sport_cpm]
    # Through line_product_spec, never off the raw dicts: it is the one lookup
    # that resolves a label and a rate for any key, sports included, and going
    # around it is how the same product ends up named or priced two ways.
    return {key: line_product_spec(key) for key in keys}


def seed_media_plan_rows(selections, geo_default, default_targeting, flight_label):
    """Section C -> Section E: each selected product/format seeds a proposal
    line with its default CPM (spec section 5, Section C description).
    Targeting defaults to the Campaign Specs Audience field, except for
    Streaming Retargeting and Live Sports, which have fixed standard
    targeting copy. Geo/Flight default to the Geography/Timing-derived
    values.

    `geo_default` may be ONE geo, a LIST of geos, a list of (audience, geo)
    pairs, or a list of (audience, geo, group_id) triples -- and a list
    produces one line per entry per product. That list comes from the avails
    table (see plan_lines_from_avails) or from targeting groups (see
    plan_lines_from_groups), so the plan mirrors it exactly: one plan line
    per avails row, audience-major, nothing merged.

    A triple's group_id is stamped onto the row as `_group_ids: [group_id]`
    -- a hidden field the media plan editor never shows, used only to match a
    clean row back to its group when a shared field changes (never string
    matching, which is what let two rows sharing a Geo string collapse onto
    each other). A pair (no group_id) or a scalar geo produces a row with NO
    `_group_ids` key at all, not one set to None -- so a plan seeded before
    targeting groups existed, or from anything that still calls this with
    the older shapes, is byte-identical to what it always was.

    Before this the plan always took the joined geo string, so a three-market
    proposal came back as a single line reading "Philadelphia, Atlanta" no
    matter what the avails table said -- a default written when the avails
    table had no grouping to follow.

    There is deliberately no second toggle here. The avails table IS the
    grouping, and merge/split on plan lines is the independent control that
    handles the mixed cases a toggle can't express.

    A row's own audience becomes its Targeting, so two audiences in one market
    are two distinguishable lines. The product's own targeting_copy still
    wins where it has one (Streaming Retargeting, every sports package) --
    resolve_row_defaults settles that, exactly as it does everywhere else.
    """
    rows = []
    products = selections["products"]

    entries = geo_default if isinstance(geo_default, (list, tuple)) else [geo_default]
    pairs = []
    for entry in entries:
        group_id = None
        if isinstance(entry, (list, tuple)) and len(entry) == 3:
            audience, geo, group_id = entry
        elif isinstance(entry, (list, tuple)) and len(entry) == 2:
            audience, geo = entry
        else:
            audience, geo = "", entry
        if str(geo or "").strip():
            pairs.append((str(audience or "").strip(), geo, group_id))
    if not pairs:
        pairs = [("", geo_default, None)]

    def _row(product_key):
        """Seed one line per avails entry from a rate-card product key.

        Everything (label, CPM, and whether the tactic carries fixed
        targeting copy) comes from the same rate-card lookup the AI draft
        path uses, so the two can't price or name the same product
        differently. Returns a LIST -- one row per entry.
        """
        label, cpm = line_product_spec(product_key)
        out = []
        for audience, geo, group_id in pairs:
            defaults = resolve_row_defaults(
                label, geo, audience or default_targeting, flight_label)
            row = {"Tactic": label, "Flight": defaults["Flight"],
                  "Geo": defaults["Geo"], "Targeting": defaults["Targeting"],
                  "Impressions": 0.0, "CPM": cpm,
                  "Type": ROW_TYPE_RATE, "Cost": 0.0}
            if group_id is not None:
                row["_group_ids"] = [group_id]
            out.append(row)
        return out

    if selections.get("_premion_streaming_tv"):
        rows.extend(_row("premion_streaming_tv"))

    sr = products.get("streaming_retargeting", {})
    if sr.get("display"):
        rows.extend(_row("streaming_retargeting_display"))
    if sr.get("preroll"):
        rows.extend(_row("streaming_retargeting_preroll"))

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
                rows.extend(_row(product_key))

    sports = products.get("live_sports", {})
    if sports.get("enabled"):
        for sport_key in sports.get("sports", []):
            rows.extend(_row(f"{SPORT_PRODUCT_PREFIX}{sport_key}"))

    if products.get("total_tv"):
        rows.extend(_row("broadcast_tv"))

    # Dynamic Video Ads are a one-time production charge, not impressions --
    # a flat-fee row, seeded at the standard rate and editable like any other.
    if products.get("dynamic_creative"):
        rows.append({"Tactic": DYNAMIC_AD_LINE_LABEL, "Flight": flight_label,
                     "Geo": pairs[0][1], "Targeting": DYNAMIC_AD_TARGETING,
                     "Impressions": 0.0, "CPM": 0.0,
                     "Type": ROW_TYPE_FLAT_FEE, "Cost": float(DYNAMIC_AD_DEFAULT_FEE)})

    if not rows:
        rows.append({"Tactic": "", "Flight": flight_label, "Geo": pairs[0][1],
                     "Targeting": "", "Impressions": 0.0, "CPM": 0.0,
                     "Type": ROW_TYPE_RATE, "Cost": 0.0})

    return rows


def row_belongs_to_tactic(tactic, seeded_tactic):
    """Does this row come from that seeded product?

    Prefix match, because a drafted row appends its own label to the
    product's name -- "Premion Streaming TV — Commercial" belongs to
    "Premion Streaming TV" and has to disappear with it when the product is
    unticked.
    """
    tactic = str(tactic or "").strip()
    return tactic == seeded_tactic or tactic.startswith(seeded_tactic + " ")


def _add_missing_rows(option, fresh_rows):
    """Append any `fresh_rows` entry not already represented on the option's
    grid -- the add-only half of `_apply_product_diff`'s reseed (a product
    selection changed: remove-by-tactic, then this).

    Grouped by tactic first, because the two cases need different rules:

    - **Exactly one fresh row for this tactic** (the ordinary case: one
      product, no multi-group fan-out) keeps the old, pre-groups rule --
      covered by ANY existing row under that tactic, prefix-matched via
      row_belongs_to_tactic, grouped or not. This has to stay string-based:
      `resolve_drafted_lines` (the AI-draft path) never stamps `_group_ids`
      onto the rows it builds, even though `apply_draft_to_form` creates a
      real, id-bearing targeting group for the same audience alongside it
      -- so a drafted row and "its" group's id are never linked on the row
      itself. Matching by id here would read a drafted line as uncovered
      and silently duplicate it (found the hard way: fixing the bug below
      this way first broke every drafted-budget regression test, adding a
      second, zero-priced "Premion Streaming TV" line next to the real one).
    - **More than one fresh row sharing a tactic** (a real multi-group
      product -- an avails PDF import, or several Audience-finder groups)
      is matched by GROUP ID, never by tactic-label presence: two rows can
      share one tactic label without being the same line, so "a row with
      this tactic already exists" is not "this group's row already
      exists." Matching by tactic alone is the bug this whole function
      exists to fix -- adding a second (or eleventh) targeting group to a
      form whose product was already selected produced no new plan line at
      all, because the one existing "Premion Streaming TV" row read as
      "this tactic is already seeded" for every group that shares it.
      Confirmed on both group-creating paths that hit this: the avails PDF
      importer (11 groups in one import) and the Audience finder's own
      "New group" button (one at a time). A pre-existing UNGROUPED row for
      this same tactic (a stray single-line default from before any group
      existed, or a drafted line) is left exactly as it is, coexisting
      alongside the new group rows rather than being silently guessed to
      already stand in for one specific group -- that guess is exactly the
      kind of silent, hard-to-spot wrong content a visible extra row is
      safer than.
    """
    fresh_by_tactic = {}
    for row in fresh_rows:
        fresh_by_tactic.setdefault(row["Tactic"], []).append(row)

    for tactic, rows in fresh_by_tactic.items():
        if len(rows) == 1:
            already = any(row_belongs_to_tactic(r.get("Tactic"), tactic) for r in option["rows"])
            if already:
                continue
            option["rows"].append(dict(rows[0]))
            option["dirty"].append(False)
            option["driver"].append(DRIVER_IMPRESSIONS)
            continue

        existing_group_ids = {gid for r in option["rows"]
                              if row_belongs_to_tactic(r.get("Tactic"), tactic)
                              for gid in group_ids_of(r)}
        for row in rows:
            ids = group_ids_of(row)
            if ids and any(gid in existing_group_ids for gid in ids):
                continue
            option["rows"].append(dict(row))
            option["dirty"].append(False)
            option["driver"].append(DRIVER_IMPRESSIONS)
            existing_group_ids.update(ids)


def queue_group_seed(group_ids):
    """Mark one or more freshly-created targeting group ids as needing their
    own media-plan line -- called directly by whatever created them (the
    avails PDF importer's `_finish_avails_import`, the Audience finder's
    `_add_segment_to_group`, in its "start a new group" branch) rather than
    inferred from a generic "something about the groups changed" comparison.

    That distinction is load-bearing, not stylistic. A generic comparison
    (e.g. re-deriving "new" as `valid_group_ids - previously_seen_ids` on
    every shared-fields change) cannot tell a real avails/finder group --
    which DOES want its own dedicated line -- from a drafted audience, which
    also gets a real, id-bearing targeting group but is deliberately
    represented by a coarser line the model itself wrote
    (`resolve_drafted_lines`, never one line per group). Tried the generic
    version first: it duplicated every drafted line the moment ANY shared
    field changed after a draft, because a draft's audiences always look
    like "new" groups to a comparison that only looks at ids. Only the two
    functions that actually mean "this group wants its own line" call this.
    """
    ids = [g for g in (group_ids or []) if g]
    if not ids:
        return
    pending = list(st.session_state.get("_pending_group_seed_ids") or [])
    st.session_state["_pending_group_seed_ids"] = pending + ids


def seed_pending_group_rows(plan_options, seed_option_rows):
    """Consume `queue_group_seed`'s queue: for each plan option, add exactly
    one row per (pending group id, currently-selected product) pair not
    already on the grid -- strict id matching, no `_add_missing_rows`
    single-row fallback, because every row this touches is, by construction,
    for a group `queue_group_seed` explicitly named. `seed_option_rows` is
    main()'s own `_seed_option_rows` closure, so a pending group's line comes
    from the exact same product/CPM/targeting resolution every other seeded
    row does.
    """
    pending_ids = set(st.session_state.pop("_pending_group_seed_ids", None) or [])
    if not pending_ids:
        return
    for option in plan_options:
        existing_ids = {gid for row in option["rows"] for gid in group_ids_of(row)}
        added = False
        for row in seed_option_rows(option["breakout"]):
            ids = [g for g in group_ids_of(row) if g in pending_ids]
            if not ids or any(g in existing_ids for g in ids):
                continue
            option["rows"].append(dict(row))
            option["dirty"].append(False)
            option["driver"].append(DRIVER_IMPRESSIONS)
            existing_ids.update(ids)
            added = True
        if added:
            option["version"] += 1


def _apply_product_diff(option, fresh_rows, previously_seeded):
    """Add rows for newly-selected products, drop rows for deselected ones,
    and leave every other row alone.

    Mutates the option in place, keeping its parallel dirty/driver lists in
    step. `previously_seeded` is what the last selection produced; anything
    in it that the current selection no longer produces has been deselected.
    """
    fresh_tactics = {row["Tactic"] for row in fresh_rows}
    removed = [t for t in previously_seeded if t not in fresh_tactics]

    kept = [i for i, row in enumerate(option["rows"])
            if not any(row_belongs_to_tactic(row.get("Tactic"), t) for t in removed)]
    option["rows"] = [option["rows"][i] for i in kept]
    option["dirty"] = [option["dirty"][i] for i in kept]
    option["driver"] = [option["driver"][i] for i in kept]

    _add_missing_rows(option, fresh_rows)


def new_plan_option(name, rows, driver=None, breakout=BREAKOUT_MONTHLY):
    """One media plan option: its own name, line set, per-line dirty/driver
    tracking, breakout mode and editor version. Everything the single plan
    used to keep in flat session_state keys now lives per option."""
    # `dirty` and `driver` are indexed in step with `rows` by everything that
    # touches an option, so a supplied driver list is padded or trimmed to
    # match rather than trusted. Belt and braces on top of the caller that got
    # this wrong (the flat-fee branch of resolve_drafted_lines): the invariant
    # matters at every construction site, so it is enforced at the only one.
    driver = list(driver or [])[:len(rows)]
    driver += [DRIVER_IMPRESSIONS] * (len(rows) - len(driver))
    return {
        "name": name,
        "rows": rows,
        "dirty": [False] * len(rows),
        "driver": driver,
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


def merge_plan_rows(option, indexes, markup):
    """Combine two or more of an option's plan lines into one, summing
    Impressions/Cost and carrying every merged line's targeting-group
    identity forward onto the surviving row (the lowest of `indexes`).

    This NEVER reads OR writes `st.session_state["targeting_groups"]` beyond
    looking a group up by id to render its label -- merging changes what a
    plan LINE says, never what a GROUP is. That's the assertion, not an
    implementation detail: `tests/test_group_merge_split.py` checks
    `targeting_groups` is deep-equal before and after.

    `_group_ids` is always REASSIGNED to a new list on the surviving row,
    never mutated in place, because `copy_plan_option` and the duplicate-line
    button both do a shallow `dict(r)` -- a row cloned that way still shares
    its source row's `_group_ids` LIST OBJECT until something reassigns it,
    so mutating in place here could silently rewrite an unrelated row's ids.
    """
    indexes = sorted(set(int(i) for i in indexes))
    if len(indexes) < 2:
        return
    rows, dirty, driver = option["rows"], option["dirty"], option["driver"]
    survivor_i = indexes[0]
    selected = [rows[i] for i in indexes]

    merged_ids = []
    for row in selected:
        merged_ids += group_ids_of(row)

    groups_by_id = {g["id"]: g for g in (st.session_state.get("targeting_groups") or [])}
    # A dead id (its group was deleted) or a repeat (a row already merged
    # once) drops out of the LABEL join, but `_group_ids` below stays the
    # literal concatenation regardless -- the ids are the row's real backing,
    # the label is just today's rendering of them.
    referenced = [groups_by_id[gid] for gid in dict.fromkeys(merged_ids) if gid in groups_by_id]
    if referenced:
        new_geo = ", ".join(tg.geo_label(g, label_for=_market_display_name) for g in referenced)
        new_targeting = ", ".join(tg.audience_label(g) for g in referenced)
    else:
        # Nothing selected is group-backed (a legacy row, a hand-typed line)
        # -- fall back to the rows' own text rather than blanking the cell.
        new_geo = ", ".join(dict.fromkeys(
            s for s in (str(r.get("Geo", "")).strip() for r in selected) if s))
        new_targeting = ", ".join(dict.fromkeys(
            s for s in (str(r.get("Targeting", "")).strip() for r in selected) if s))

    new_impressions = sum(_num(r.get("Impressions")) for r in selected)
    new_cost = sum(_num(r.get("Cost")) for r in selected)

    survivor = dict(rows[survivor_i])
    # Same three-source precedence resolve_row_defaults applies everywhere
    # else: a product with its own fixed Targeting copy (Streaming
    # Retargeting, every Live Sports package) keeps it through a merge, never
    # the merged groups' audience labels -- this was the actual bug (a
    # Streaming Retargeting line's "Retarget Exposed CTV Viewers" fell
    # through to the audience stack the moment it was merged with anything).
    # A broadcast row's own Geo AND Targeting are held as-is the same way,
    # for the same reason resolve_row_defaults holds both via `current` --
    # Geo is the more dangerous of the two to get wrong here: it's derived
    # from the station's call sign, never guessed, and a broadcast row isn't
    # group-backed, so merging it with a group-backed row would otherwise
    # overwrite it with that group's market name, a wrong DMA on a client's
    # plan the moment a rep merges a broadcast line with anything else.
    survivor_is_broadcast = is_broadcast_row(survivor)
    if survivor_is_broadcast:
        pass  # Geo untouched -- schedule-derived, never the merged groups' markets
    else:
        survivor["Geo"] = new_geo
    fixed = fixed_targeting_copy(survivor.get("Tactic", ""))
    if fixed:
        survivor["Targeting"] = fixed
    elif survivor_is_broadcast:
        survivor["Targeting"] = survivor.get("Targeting") or new_targeting
    else:
        survivor["Targeting"] = new_targeting
    survivor["Impressions"] = new_impressions
    survivor["Cost"] = new_cost
    # Re-derive CPM from the summed pair rather than averaging the merged
    # rows' own CPMs -- the exact inverse of cost_from_impressions, so the
    # merged row is internally consistent the instant it's created.
    effective = row_markup(survivor, markup)
    survivor["CPM"] = ((new_cost * 1000.0) / (new_impressions * effective)
                       if new_impressions and effective else 0.0)
    if merged_ids:
        survivor["_group_ids"] = merged_ids
    else:
        survivor.pop("_group_ids", None)

    drop = set(indexes[1:])
    new_rows, new_dirty, new_driver = [], [], []
    for i, row in enumerate(rows):
        if i in drop:
            continue
        if i == survivor_i:
            new_rows.append(survivor)
            new_dirty.append(True)          # a merge is immediately customized
            new_driver.append(driver[i])
        else:
            new_rows.append(row)
            new_dirty.append(dirty[i])
            new_driver.append(driver[i])

    option["rows"] = new_rows
    option["dirty"] = new_dirty
    option["driver"] = new_driver


def split_plan_row(option, index):
    """Reverse a merge: reproduce one row per group id the row at `index`
    carries, each through `resolve_row_defaults` with its own group's own
    label -- the group-aware sibling of `seed_media_plan_rows`' per-triple
    seeding, run on demand for one existing row instead of at seed time.

    Never touches `st.session_state["targeting_groups"]`, same invariant as
    `merge_plan_rows`, the other direction -- a group is only ever READ here,
    to render its label.

    Impressions/Cost are split EVENLY across the new rows: the merge that
    produced this row summed them and did not keep each part's own original
    share, so a merge immediately followed by a split reproduces the same
    full-flight total, not necessarily the row-by-row numbers from before the
    merge ever happened. A row with one id or none has nothing to split into
    and is left alone -- the "Split line" button only appears on a row with
    more than one, but this stays safe if ever called otherwise.
    """
    rows, dirty, driver = option["rows"], option["dirty"], option["driver"]
    row = rows[index]
    ids = group_ids_of(row)
    if len(ids) <= 1:
        return

    groups_by_id = {g["id"]: g for g in (st.session_state.get("targeting_groups") or [])}
    n = len(ids)
    tactic = row.get("Tactic", "")
    flight = row.get("Flight", "")
    cpm = row.get("CPM", 0.0)
    row_type = row.get("Type", ROW_TYPE_RATE)
    share_impressions = _num(row.get("Impressions")) / n
    share_cost = _num(row.get("Cost")) / n
    this_driver = driver[index]

    new_rows = []
    for group_id in ids:
        group = groups_by_id.get(group_id)
        # A dead id (its group was deleted since the merge) still gets its
        # own row -- it just can't recover a label, so it falls back to the
        # empty string, the same honesty `geo_label`'s own text fallback uses.
        audience = tg.audience_label(group) if group else ""
        geo = tg.geo_label(group, label_for=_market_display_name) if group else ""
        defaults = resolve_row_defaults(tactic, geo, audience, flight)
        new_rows.append({
            "Tactic": tactic, "Flight": defaults["Flight"], "Geo": defaults["Geo"],
            "Targeting": defaults["Targeting"], "Impressions": share_impressions,
            "CPM": cpm, "Type": row_type, "Cost": share_cost,
            "_group_ids": [group_id],
        })

    option["rows"] = rows[:index] + new_rows + rows[index + 1:]
    option["dirty"] = dirty[:index] + [True] * n + dirty[index + 1:]
    option["driver"] = driver[:index] + [this_driver] * n + driver[index + 1:]


def next_option_name(existing_names):
    for candidate in DEFAULT_OPTION_NAMES:
        if candidate not in existing_names:
            return candidate
    return f"Option {len(existing_names) + 1}"


def reconcile_plan_rows(option, edited_rows, markup):
    """Fold one option's edited grid back into its stored state: update each
    row's dirty flag and driver, then recompute the non-driving side of every
    rate row. Returns True if any value actually changed as a result, OR the
    row count changed -- either way the caller bumps the option's `version`
    and re-renders.

    The row-count check is deliberate, not incidental: a row added or removed
    via the grid's own "+"/"-" (num_rows="dynamic") has to force a fresh
    widget key on the very next render the same way avails_version does for
    the D2 grid, or the widget's own in-progress edit delta can outlive the
    stale key it was created under and never actually reach `option["rows"]`
    to stay. Relying on `recomputed_any` alone would only bump by accident,
    whenever the new row happens to also change a number through CPM
    recompute -- true for the common case (a rep types a real dollar figure),
    false for a flat-fee row or a Cost that happens to already match its
    derivation, which is exactly the gap this class of bug hides in.
    """
    prev_rows = option["rows"]
    prev_dirty = option["dirty"]
    prev_drivers = option["driver"]
    row_count_changed = len(edited_rows) != len(prev_rows)

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
    return recomputed_any or row_count_changed


def compute_plan_totals(rows, breakout_mode, n_months, flight_label,
                        broadcast_months=None):
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
            # A broadcast line scales by the months its own schedule runs in,
            # not the plan's. A four-week May buy inside a three-month plan
            # flight is one month of broadcast, and multiplying its monthly
            # figure by three would have invented $56,000 of spend that no
            # station is going to run.
            row_months = (broadcast_months if broadcast_months and is_broadcast_row(row)
                          else n_months)
            row_months = max(1, int(row_months))
            if breakout_mode.startswith("Full Flight"):
                full_flight_impressions = entered_impressions
                full_flight_cost = entered_cost
                monthly_impressions = full_flight_impressions / row_months
                monthly_cost = full_flight_cost / row_months
            else:
                monthly_impressions = entered_impressions
                monthly_cost = entered_cost
                full_flight_impressions = monthly_impressions * row_months
                full_flight_cost = monthly_cost * row_months

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
            "cpm": _num(row.get("CPM")),
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


def build_included_list(targeting, commercial_production, vertical_attribution_label=None):
    """The "Included with Campaign" list printed on the media plan slide.

    `vertical_attribution_label` is the vertical's own branded measurement
    product (Polk Signals, Arrivalist) and stands in for the generic sales
    attribution line rather than joining it -- an automotive client is told
    they're getting new-car sales attribution, not "Sales Attribution (CRM
    Upload Required)", which describes a different mechanism entirely.
    """
    included = [
        "Dedicated Account Management Team",
        "Monthly Reporting Calls & Optimizations",
        "Dashboard Access",
        "Web Attribution (Pixel Required)",
    ]
    if commercial_production:
        included.append("Commercial Production")
    if vertical_attribution_label:
        included.append(vertical_attribution_label)
    elif targeting.get("sales_attribution"):
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
    return _call_claude_json(build_case_study_prompt(slide_texts, filename),
                             label="case_study_tags")


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
            # Said plainly rather than left to be discovered: it works right
            # now, via the copied-slide path, which is a little less faithful
            # than the rendered one. The queue is the vault list itself --
            # anything without images is pending, so there's nothing separate
            # to keep in sync.
            pending, _ = db.fetch_case_studies(active_only=False)
            waiting = [r for r in (pending or []) if db.case_study_render_path(r) == "copy"]
            st.warning(
                "**Needs slide images for full fidelity.** Until then it goes into decks "
                "as copied slides, which render slightly less faithfully — a long "
                "paragraph can clip. Rendering needs PowerPoint and the brand font, so it "
                "can't happen here.\n\n"
                f"Run this locally when you get a chance — {len(waiting)} case "
                f"stud{'y is' if len(waiting) == 1 else 'ies are'} waiting:\n\n"
                "```\npython render_case_study_images.py --pending\n```")
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


def market_profile_images(target_dmas, include_slides):
    """(image_paths, warnings) for the selected markets, in picker order.

    One resolver for both graft sites -- generate and rebuild-as-presented --
    for the same reason case_study_source is one function: a rebuild that
    resolved its markets differently from the build would produce a different
    deck from the one the client was sent, which is the exact thing
    "as presented" promises it won't.

    Order is the caller's order and is preserved: it is the order the profile
    slides appear in the deck.

    A market with no authored profile slide is skipped SILENTLY and on
    purpose. It is a known, expected state -- five DMAs are in it, the picker
    labels each one at the point of choosing, and the market is still a valid
    target. Warning again here would fire on every build for a condition the
    rep has already been told about and deliberately accepted, which is how a
    warning channel stops being read.
    """
    if not include_slides or not target_dmas:
        return [], []

    profiles, _ = load_market_profiles()
    by_key = {row.get("key"): row for row in profiles}
    paths, warnings = [], []
    for key in target_dmas:
        row = by_key.get(key)
        if row is None:
            warnings.append(f"Target market \"{key}\" is no longer in the market "
                            f"list -- built without its profile slide.")
            continue
        if not row.get("image_path"):
            continue
        path = db.market_profile_image(key, row["image_path"])
        if path:
            paths.append(path)
        else:
            warnings.append(f"Couldn't fetch the {row.get('label')} viewer profile "
                            f"slide -- built without it.")
    return paths, warnings


def case_study_source(row):
    """How this case study goes into a deck: rendered images, or the .pptx.

    One resolver for both graft sites (generate, and rebuild-as-presented) so
    a proposal and a rebuild can't take different routes for the same case
    study. Images win when they exist: copying slide XML between decks has
    produced five distinct corruption bugs here and cannot reproduce
    PowerPoint's live autofit, so flattened text can clip. Falling back to
    the copy path is normal, not an error -- a case study has images only
    once render_case_study_images.py has been run for it locally.
    """
    if db.case_study_render_path(row) == "images":
        images = db.case_study_images(row["id"], tuple(row["slide_images"]))
        return {"images": images, "title": row.get("title")}
    return {"path": db.case_study_file(row["id"], row["storage_path"]),
            "slides": None, "title": row.get("title")}


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

    # Which route each one takes into a deck. Worth surfacing rather than
    # leaving implicit: a case study without rendered slides still works, but
    # it goes in as copied XML, which is the path that can clip text.
    pending = [r for r in rows if db.case_study_render_path(r) == "copy"]
    if pending:
        st.info(f"**{len(rows) - len(pending)} of {len(rows)} render from images.** "
                f"The other {len(pending)} are copied slide-by-slide, which is a little "
                f"less faithful — long paragraphs can clip. Run "
                f"`python render_case_study_images.py --pending` locally to fix that "
                f"(it needs PowerPoint and the brand font, so it can't run here).")
    else:
        st.success(f"All {len(rows)} case studies render from pre-rendered images.")

    for row in shown:
        state = "" if row.get("active", True) else "  ·  deactivated"
        route = "🖼 images" if db.case_study_render_path(row) == "images" else "📄 copied"
        with st.expander(f"{row['title'] or row['filename']}  ·  {route}{state}",
                         expanded=False):
            st.caption(f"{row.get('summary') or '_no summary_'}")
            st.caption(f"added by {row.get('added_by') or 'unknown'} · "
                       f"{str(row.get('date_added'))[:10]} · `{row['filename']}`")
            if db.case_study_render_path(row) == "images":
                st.caption(f"Renders from {len(row['slide_images'])} pre-rendered slide "
                           f"image(s) at {row.get('image_width') or '?'}px — pixel-faithful "
                           f"to the source deck, but the text isn't selectable.")
            else:
                st.caption("Goes into a deck as copied slides. Run "
                           "`python render_case_study_images.py --pending` locally to "
                           "render images for it.")

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
                    build_case_study_suggest_prompt(description, active),
                    label="case_study_suggest")
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


BREAKOUT_FULL_FLIGHT_LABEL = "Full flight (default)"
BREAKOUT_MONTHLY_LABEL = "One slide per month"
VIEW_WEEKLY_LABEL = "Week by week (default)"
VIEW_TOTALS_LABEL = "Totals only"


def escape_markdown_money(text):
    """Keep dollar amounts readable in Streamlit markdown.

    Streamlit treats `$...$` as inline LaTeX, so a review item mentioning
    "$850" and "$50,000" has everything between the two signs swallowed and
    re-rendered as maths -- which is where the mangled "`850" came from. An
    escaped dollar sign renders literally.
    """
    return str(text).replace("$", r"\$")


def render_review_list():
    """The "Review before generating" panel, in two sections.

    Split because the two kinds of item need different people: questions
    only the client can settle, and checks the seller does before sending.
    Mixing them made a list where nothing looked actionable.
    """
    client_items = list(st.session_state.get("draft_unresolved") or [])
    internal_items = list(st.session_state.get("draft_unresolved_internal") or [])

    # Live, not stored -- recomputed every run from whatever targeting_groups
    # holds RIGHT NOW, so a custom-audience overage reaches the seller here,
    # at generate time, regardless of whether it came from a draft, a hand
    # edit in the Phase 4 grid, or the Phase 5 finder's own click-time toast
    # (which shows once and can scroll past unread). One computation, three
    # sources -- see tg.custom_segment_count's own docstring for why it
    # counts distinct segments across every group rather than per group.
    groups = st.session_state.get("targeting_groups") or []
    rfp_map = dict(zip(audience_catalog["segment"], audience_catalog["rfp_selectable"]))
    custom_count = tg.custom_segment_count(groups, rfp_map)
    if custom_count > 1:
        internal_items = internal_items + [
            f"{custom_count} custom (non-RFP-selectable) audiences are in play across your "
            f"targeting groups -- only one is allowed per campaign."]

    if not client_items and not internal_items:
        return

    lines = []
    if client_items:
        lines.append("**Confirm with the client**")
        lines += [f"- {escape_markdown_money(item)}" for item in client_items]
    if internal_items:
        if lines:
            lines.append("")
        lines.append("**Before sending**")
        lines += [f"- {escape_markdown_money(item)}" for item in internal_items]
    st.warning("\n".join(lines))


def _month_sort(label):
    """Sort "May 2025"-style labels chronologically, not alphabetically."""
    try:
        return datetime.strptime(label, "%b %Y")
    except ValueError:
        return datetime.min


def abbreviate_count(value):
    """1,831,600 -> 1.83M. For st.metric, which truncates rather than wraps."""
    value = float(value or 0)
    for limit, suffix in ((1_000_000_000, "B"), (1_000_000, "M"), (1_000, "K")):
        if abs(value) >= limit:
            scaled = value / limit
            return f"{scaled:.2f}{suffix}" if abs(scaled) < 10 else f"{scaled:.1f}{suffix}"
    return f"{value:,.0f}"


def abbreviate_money(value):
    """$28,000 -> $28.0K."""
    return f"${abbreviate_count(value)}"


def render_broadcast_schedule_import():
    """Import a Wide Orbit schedule and choose how it's shown.

    Sits inside Section C under Total TV because that's the only context it
    means anything in. Parsing happens on upload rather than at generate
    time, so a file the reader can't handle is reported while the seller is
    still looking at the uploader -- and it degrades to manual entry of the
    summary numbers rather than blocking the proposal.
    """
    with st.expander("📺 Broadcast schedule (from Wide Orbit)", expanded=False):
        st.caption("Upload the Wide Orbit export for this buy — the Campaign Schedule Report "
                   "(.xlsx), or a Planner (.xls / .pdf). The schedule becomes its own slide, "
                   "and the broadcast line is added to your media plan automatically.")
        upload = st.file_uploader("Wide Orbit export", type=["xlsx", "xls", "pdf"],
                                  key="wo_upload")
        # AppTest can't operate a file_uploader, so in test mode a path may
        # be handed in instead. The real widget path below is untouched --
        # this only supplies the same (name, bytes) a click would have.
        injected = test_mode_upload("wo_upload_path")
        if injected is not None:
            upload = injected

        if upload is not None and st.session_state.get("wo_loaded_name") != upload.name:
            target = db.scratch_dir("premion_wo_uploads") / upload.name
            target.write_bytes(upload.getvalue())
            try:
                schedule = wideorbit.parse_schedule(str(target), upload.name)
            except wideorbit.ScheduleParseError as exc:
                st.session_state["broadcast_schedule"] = None
                st.session_state["wo_error"] = str(exc)
            else:
                st.session_state["broadcast_schedule"] = schedule
                st.session_state["wo_error"] = None
            st.session_state["wo_loaded_name"] = upload.name
            st.rerun()

        error = st.session_state.get("wo_error")
        if error:
            st.error(error)
            st.caption("You can still build the proposal — fill the broadcast line in by hand "
                       "on the media plan below.")

        schedule = st.session_state.get("broadcast_schedule")
        if schedule:
            s = schedule.summary
            # Abbreviated, because st.metric truncates rather than wraps and
            # five of them across the panel leaves each one narrow -- a
            # cost rendered "$2..." is worse than useless. Full precision
            # goes in the caption directly beneath.
            cols = st.columns(5)
            cols[0].metric("Spots", abbreviate_count(s.total_spots))
            cols[1].metric("Gross", abbreviate_money(s.gross_cost))
            cols[2].metric("Imps", abbreviate_count(s.impressions))
            cols[3].metric("Reach", f"{s.reach:.1f}" if s.reach else "--")
            cols[4].metric("Freq", f"{s.frequency:.1f}" if s.frequency else "--")
            st.caption(f"**{s.total_spots:,} commercials · ${s.gross_cost:,.0f} gross · "
                       f"{s.impressions:,.0f} {s.demo_label} impressions**")
            st.caption(f"{s.station or 'Station'} · {s.flight_start} to {s.flight_end} · "
                       f"{len(schedule.grid_weeks)} weeks · {len(schedule.rows)} programs · "
                       f"read from {schedule.source_name}")
            for note in schedule.notes:
                st.caption(f"ℹ️ {note}")

            st.text_area(
                "Broadcast plan description", key="broadcast_plan_desc", height=70,
                placeholder="e.g. 210x Commercials Per Month, Morning News Mon-Tue, 3 Weeks/Month",
                help="Appears on the schedule slide, in your own words.")
            bcol1, bcol2 = st.columns(2)
            bcol1.radio("Breakout", [BREAKOUT_FULL_FLIGHT_LABEL, BREAKOUT_MONTHLY_LABEL],
                        key="broadcast_breakout",
                        help="Full flight shows the whole schedule; monthly gives one slide "
                             "per calendar month.")
            bcol2.radio("Detail", [VIEW_WEEKLY_LABEL, VIEW_TOTALS_LABEL], key="broadcast_view",
                        help="Week by week shows a column per week; totals only shows one "
                             "Total Spots column per program.")

            # Informational, not a nudge to change the default: week-by-week
            # is the right view for most schedules, and a long season one is
            # genuinely a lot of information. This just makes the cost of the
            # current settings visible before Generate rather than after.
            breakout, detailed = broadcast_display_options()
            slides = assembly.estimate_schedule_slide_count(schedule, breakout, detailed)
            if slides > 2:
                alternative = assembly.estimate_schedule_slide_count(schedule, breakout, False)
                if detailed and alternative < slides:
                    st.caption(f"ℹ️ These settings produce **{slides} schedule slides**. "
                               f"Switch Detail to *{VIEW_TOTALS_LABEL}* for "
                               f"{'a single slide' if alternative == 1 else f'{alternative}'}.")
                else:
                    st.caption(f"ℹ️ These settings produce **{slides} schedule slides**.")

            if st.button("Remove this schedule", key="wo_clear"):
                for key in ("broadcast_schedule", "wo_error", "wo_loaded_name"):
                    st.session_state.pop(key, None)
                st.rerun()


def broadcast_display_options():
    """(breakout, detailed) for the schedule slides, from the rep's picks."""
    breakout = ("monthly" if st.session_state.get("broadcast_breakout") == BREAKOUT_MONTHLY_LABEL
                else "full_flight")
    detailed = st.session_state.get("broadcast_view", VIEW_WEEKLY_LABEL) != VIEW_TOTALS_LABEL
    return breakout, detailed


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
            injected = test_mode_upload(f"hist_up_path_{rid}")
            if injected is not None:
                upload = injected
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


def _week_start(value):
    """The Monday of the week a timestamp falls in, as a date."""
    stamp = _parse_draft_date(str(value)[:10])
    if stamp is None:
        return None
    return stamp - timedelta(days=stamp.weekday())


def render_usage_stats(rows, case_studies):
    """Adoption, in one screen.

    Deliberately read-only and computed from what's already logged -- no new
    tracking, no events table. The question it answers is "is anyone actually
    using this, and how", not "what is everyone doing"; anything past a
    single screen would be analytics theater.

    Every count here has to survive an empty database, since a brand new
    deployment renders this page before anything exists.
    """
    st.subheader("Usage")
    if not rows:
        st.caption("No proposals generated yet — this fills in as the team uses the app.")
        return

    drafted = sum(1 for r in rows if ((r.get("form_json") or {}).get("draft") or {}).get("notes"))
    revisions = sum(1 for r in rows if r.get("parent_proposal_id"))
    top = st.columns(4)
    top[0].metric("Proposals", len(rows))
    top[1].metric("From notes", f"{drafted}",
                  help="Drafted from meeting notes rather than filled in by hand.")
    top[2].metric("By hand", f"{len(rows) - drafted}")
    top[3].metric("Case studies", len(case_studies or []))

    # Last 8 weeks, including the quiet ones -- a gap is the finding.
    this_week = _week_start(date.today().isoformat())
    if this_week:
        weeks = [this_week - timedelta(weeks=offset) for offset in range(7, -1, -1)]
        counts = {week: 0 for week in weeks}
        for row in rows:
            week = _week_start(row.get("generated_at"))
            if week in counts:
                counts[week] += 1
        st.caption("**Proposals per week** (last 8 weeks)")
        st.bar_chart(pd.DataFrame({"week": [w.strftime("%b %d") for w in weeks],
                                   "proposals": [counts[w] for w in weeks]})
                     .set_index("week"), height=180)

    def _breakdown(label, key, mapper=None):
        tally = {}
        for row in rows:
            value = row.get(key) or None
            if mapper:
                value = mapper(value)
            tally[value or "—"] = tally.get(value or "—", 0) + 1
        frame = (pd.DataFrame({label: list(tally), "proposals": list(tally.values())})
                 .sort_values("proposals", ascending=False).set_index(label))
        return frame

    scol1, scol2, scol3 = st.columns(3)
    with scol1:
        st.caption("**By vertical**")
        st.dataframe(_breakdown("Vertical", "vertical",
                                lambda v: next((label for label, key in VERTICALS.items()
                                                if key == v), v)),
                     use_container_width=True, height=240)
    with scol2:
        st.caption("**By market**")
        st.dataframe(_breakdown("Market", "market"), use_container_width=True, height=240)
    with scol3:
        # "—" is every proposal generated before identity existed, which is
        # worth seeing rather than hiding.
        st.caption("**By person**")
        st.dataframe(_breakdown("Person", "created_by"), use_container_width=True, height=240)

    if revisions:
        st.caption(f"{revisions} of these are revisions of an earlier proposal.")


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

    with st.expander("📊 Usage stats", expanded=False):
        case_studies, cs_warning = db.fetch_case_studies(active_only=False)
        render_usage_stats(rows, [] if cs_warning else case_studies)

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

    fcol1, fcol2, fcol3 = st.columns([2, 1, 1])
    search = fcol1.text_input("Search by client", key="history_search").strip().lower()
    verticals_present = sorted({r.get("vertical") for r in rows if r.get("vertical")})
    vertical_labels = ["All"] + [next((label for label, key in VERTICALS.items() if key == v), v)
                                 for v in verticals_present]
    picked = fcol2.selectbox("Vertical", vertical_labels, key="history_vertical")
    picked_market = fcol3.selectbox("Market", ["All", "DC", "Harrisburg"], key="history_market")

    filtered = rows
    if search:
        filtered = [r for r in filtered if search in (r.get("client_name") or "").lower()]
    if picked != "All":
        want = VERTICALS.get(picked, picked)
        filtered = [r for r in filtered if r.get("vertical") == want]
    if picked_market != "All":
        filtered = [r for r in filtered if r.get("market") == picked_market]

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


def render_update_audience_usage():
    """Upload a refreshed YTD audience-usage workbook, see exactly what it
    would add to the catalog, and activate it -- the deck_versions pattern
    applied to audience data. Merge, never replace: the brochure stays
    authoritative for category/RFP-selectable on anything it already has
    (`targeting_groups.custom_segment_count` depends on that), and the
    workbook only ever refreshes usage numbers on those and adds net-new
    segments, `rfp_selectable=False` by default. CLT (client first-party
    data) and one advertiser's own retargeting/address pool (mechanically
    detected, or from `audience_component_overrides.csv` when it isn't)
    are filtered out entirely, not just hidden -- see
    `audience_usage_import`'s own docstring.
    """
    st.header("Update audience usage")
    st.caption("For when Matt sends a refreshed YTD workbook (Segment Name / Delivered "
               "Impressions, one row per booked stack). Upload it to see exactly what it "
               "would add before anything goes live. Nothing goes live until you activate "
               "it, and every previous workbook is kept, so this is reversible.")

    active, warning = db.active_audience_usage_version()
    if warning:
        st.warning(f"{warning}. You can still scan a workbook below.")
    else:
        st.info(f"**In use:** version {active['id']} — {active['filename']} "
                f"(uploaded {str(active['uploaded_at'])[:10]})"
                + (f"\n\n{active['notes']}" if active.get("notes") else ""))

    upload = st.file_uploader("New audience usage workbook (.xlsx)", type=["xlsx"], key="usage_upload")
    # AppTest can't operate a file_uploader, so in test mode a path may be
    # handed in instead -- same device the Wide Orbit uploader uses.
    injected = test_mode_upload("usage_upload_path")
    if injected is not None:
        upload = injected
    if not upload:
        return

    local_path = db.scratch_dir("premion_usage_uploads") / upload.name
    local_path.write_bytes(upload.getvalue())

    with st.spinner("Parsing the workbook..."):
        try:
            workbook = audience_usage_import.parse_workbook(str(local_path))
        except audience_usage_import.UsageImportError as exc:
            st.error(f"Couldn't read that workbook: {exc}")
            return
        existing_catalog = load_audience_catalog().to_dict("records")
        report = audience_usage_import.derive_catalog_updates(workbook, existing_catalog)

    counts = st.columns(6)
    for column, (label, value) in zip(counts, [
        ("Stacks", len(workbook.rows)), ("Dropped", workbook.dropped_no_data_targeting),
        ("Gained", report.gained), ("Collided", report.collided),
        ("Uncategorized", report.uncategorized),
        ("Client-excluded", len(report.excluded_client_pattern) + len(report.excluded_client_override)),
    ]):
        column.metric(label, value)
    st.caption(f"{report.collided} existing catalog segment(s) get refreshed usage numbers "
               f"only -- category and RFP-selectable status untouched, brochure-authoritative.")

    if report.excluded_clt:
        with st.expander(f"Excluded as CLT, client first-party data ({len(report.excluded_clt)})",
                         expanded=False):
            st.caption("Never selectable for anyone else -- filtered out of the catalog and the "
                       "booking-evidence index entirely, not just hidden in the UI.")
            for name in report.excluded_clt:
                st.markdown(f"- {name}")
            if report.clt_aliases:
                st.error(f"**{len(report.clt_aliases)} CLT component(s) also match a non-CLT "
                         f"segment once the CLT tag is stripped off -- review before activating:**")
                for clt_raw, matched in report.clt_aliases:
                    st.markdown(f"- `{clt_raw}` matches **{matched}**")
            else:
                st.caption("No CLT component's stripped text matches a non-CLT segment -- "
                           "checked, not assumed.")

    client_excluded_total = len(report.excluded_client_pattern) + len(report.excluded_client_override)
    if client_excluded_total:
        with st.expander(f"Excluded as one advertiser's own retargeting/address pool "
                         f"({client_excluded_total})", expanded=False):
            st.caption("Same treatment as CLT -- never selectable for anyone else, dropped from "
                       "the catalog entirely.")
            if report.excluded_client_pattern:
                st.markdown(f"**Caught automatically** (WEB RT / LOCATION RT / "
                            f"`..._RFPID-..._RT` / an address-list suffix) -- "
                            f"{len(report.excluded_client_pattern)}:")
                for name in report.excluded_client_pattern:
                    st.markdown(f"- {name}")
            if report.excluded_client_override:
                st.markdown(f"**From the manual registry** (no structural marker to catch "
                            f"automatically) -- {len(report.excluded_client_override)}:")
                for name in report.excluded_client_override:
                    st.markdown(f"- {name}")

    if report.overrides_applied:
        with st.expander(f"Manual overrides applied ({sum(report.overrides_applied.values())})",
                         expanded=False):
            st.caption("From `audience_component_overrides.csv` -- decisions a human made because "
                       "they couldn't be derived mechanically.")
            for action, count in sorted(report.overrides_applied.items()):
                st.markdown(f"- **{action}**: {count}")

    unresolved = [u for u in report.updates if u.source == "uncategorized"]
    confirmed_categories = {}
    excluded_by_confirm = set()
    if unresolved:
        st.subheader(f"Needs a category -- {len(unresolved)}")
        st.caption("No recognizable prefix of its own (rules handle everything else "
                   "deterministically -- this step never touches those). Ask Claude for a "
                   "starting proposal, then confirm or correct every row before activating. "
                   "What you confirm here is what gets stored -- once a component is written "
                   "with a category, a later upload never re-derives or re-asks for it.")

        suggest_key = f"usage_categorize_for_{upload.name}"
        if st.session_state.get(suggest_key) is None:
            if st.button(f"Ask Claude to propose categories for these {len(unresolved)}"):
                with st.spinner("Asking Claude..."):
                    components = [{"name": u.segment, "impressions": u.impressions} for u in unresolved]
                    suggestions, error = call_claude_categorize(components)
                if error:
                    st.error(error)
                else:
                    st.session_state[suggest_key] = {s["segment"]: s for s in suggestions}
                    st.rerun()
            st.caption("Or fill categories in yourself in the table below -- Claude's "
                       "proposal is a starting point, not a requirement.")

        suggestions_by_segment = st.session_state.get(suggest_key) or {}
        category_options = [""] + sorted(CATEGORY_DESCRIPTIONS)
        rows = []
        for u in unresolved:
            s = suggestions_by_segment.get(u.segment, {})
            rows.append({
                "Segment": u.segment,
                "Impressions": u.impressions,
                "Category": s.get("category", ""),
                "Category 2": s.get("category2", ""),
                "Exclude (client-specific)": bool(s.get("is_client_specific")),
                "Confidence": s.get("confidence", ""),
                "Reason": s.get("reason", ""),
            })
        editor_key = f"usage_categorize_editor_{upload.name}"
        edited = st.data_editor(
            pd.DataFrame(rows), key=editor_key, use_container_width=True, hide_index=True,
            disabled=["Segment", "Impressions", "Confidence", "Reason"],
            column_config={
                "Impressions": st.column_config.NumberColumn(format="%d"),
                "Category": st.column_config.SelectboxColumn(options=category_options),
                "Category 2": st.column_config.SelectboxColumn(options=category_options),
            },
        )
        for _, row in edited.iterrows():
            if row["Exclude (client-specific)"]:
                excluded_by_confirm.add(row["Segment"])
            else:
                cats = [c for c in (row["Category"], row["Category 2"]) if c]
                confirmed_categories[row["Segment"]] = ", ".join(cats)
        still_unconfirmed = sum(1 for name in confirmed_categories
                                if not confirmed_categories[name] and name not in excluded_by_confirm)
        if still_unconfirmed:
            st.caption(f"{still_unconfirmed} still has no category and isn't marked "
                       f"client-specific -- activating will keep it uncategorized, same as "
                       f"today, rather than block on it.")

    if report.custom_case_variants:
        with st.expander(f"CUSTOM-prefixed with unusual casing ({len(report.custom_case_variants)})",
                         expanded=False):
            st.caption("Matched case-insensitively and kept non-RFP-selectable the same as every "
                       "other CUSTOM segment -- listed since \"CUSTOM\" alone undercounts them.")
            for name in report.custom_case_variants:
                st.markdown(f"- {name}")

    notes = st.text_input("Notes for this version",
                          placeholder="e.g. Q3 refresh through 9/30")
    if st.button("Activate this workbook", type="primary"):
        confirmed_for_upload = {name: {"category": cat} for name, cat in confirmed_categories.items() if cat}
        confirmed_for_upload.update({name: {"exclude": True} for name in excluded_by_confirm})
        with st.spinner("Uploading and rebuilding the catalog..."):
            stamp = datetime.now().strftime("%Y%m%dT%H%M%SZ")
            row, error = db.upload_audience_usage_workbook(
                str(local_path), f"usage/{stamp}_{upload.name}",
                notes=notes.strip() or None, activate=True,
                confirmed_categories=confirmed_for_upload)
        if error:
            st.error(error)
        else:
            clear_catalog_cache()
            load_audience_index.clear()
            st.session_state.pop(f"usage_categorize_for_{upload.name}", None)
            confirmed_note = (f" {len(confirmed_for_upload)} reviewed categorization(s) are "
                              f"now stored -- a later upload won't ask again for these."
                              if confirmed_for_upload else "")
            st.success(f"Version {row['id']} is now the active audience usage workbook -- "
                       f"the catalog and booking-evidence panel are rebuilt.{confirmed_note} "
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
    # A group's own geo expander can jump here directly (roadmap §E), same
    # keyed-radio mechanic as history_goto_build above.
    if st.session_state.pop("goto_zip_map_builder", False):
        st.session_state["page_choice"] = "Zip/map builder"
    page = st.sidebar.radio("Page", [
        "Build a proposal",
        "Proposal history",
        "Audience finder",
        "Zip/map builder",
        "Case study finder",
        "Add case study",
        "Update master deck",
        "Update audience usage",
    ], label_visibility="collapsed", key="page_choice")
    st.sidebar.caption("The finders are also embedded in the proposal flow — "
                       "audiences in Section D2, case studies just before Generate.")
    st.sidebar.divider()
    render_identity_sidebar()
    st.sidebar.caption(f"Build {BUILD_STAMP}")

    standalone = {
        "Proposal history": render_proposal_history,
        "Audience finder": render_audience_finder_page,
        "Zip/map builder": render_zip_map_builder_page,
        "Case study finder": render_case_study_finder,
        "Add case study": render_add_case_study,
        "Update master deck": render_update_master_deck,
        "Update audience usage": render_update_audience_usage,
    }
    if page in standalone:
        standalone[page]()
        return

    # Before any widget on this page is instantiated, and only on this page:
    # a keyed widget that didn't render while the seller was on another page
    # has had its value collected, and this is the one moment it can be put
    # back. Restoring on every page instead would inject Build keys into runs
    # that don't own them, for no gain.
    restore_form_state()
    apply_pending_avails_import_fields()

    heading, new_proposal = st.columns([4, 1], vertical_alignment="bottom")
    with heading:
        st.title("Premion Proposal Builder")
    with new_proposal:
        render_new_proposal_button()
    render_new_proposal_confirm()
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

    render_review_list()

    # A rep's notes mentioning avails is the common real flow: they have the
    # Salesforce document in hand while drafting. A simple, deterministic
    # keyword check on the notes actually pasted in -- not a Claude-classified
    # signal -- so this doesn't depend on a live model call to verify; the
    # richer three-way "figures stated / avails mentioned / not mentioned"
    # signal geo_targeting_roadmap.md F describes is a fast-follow, not this.
    source_notes = st.session_state.get("draft_source_notes", "") or ""
    if "avail" in source_notes.lower() and not st.session_state.get("avails_import_history"):
        with st.container(border=True):
            st.caption("📄 Your notes mention avails you've pulled. Upload the document to fill "
                       "in the table and zip targeting -- Audience/Geography and the media plan "
                       "below will pick it up the same way they pick up anything typed by hand.")
            render_avails_pdf_uploader("review", "The avails PDF your notes referenced.")

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
        market_profile_rows, market_profile_warning = load_market_profiles()
        # Ahead of the picker, not only at D2/E's own call sites: a group
        # resolved earlier in THIS run (or a prior one) has to be reflected
        # before target_dma_choice renders, and the autofill below reads
        # targeting_groups to do it.
        _, market_lookup_warning = install_market_lookup()
        if market_lookup_warning:
            st.warning(f"⚠️ {market_lookup_warning}")
        sync_targeting_groups()
        apply_group_markets_autofill(market_profile_rows)
        target_dmas, include_market_profile = market_profile_picker(
            market_profile_rows, market_profile_warning)
        vertical_choice = st.selectbox("Vertical", list(VERTICALS.keys()), index=0, key="vertical_choice",
                                        on_change=_clear_ai_section, args=("basics",))
        agency_involved = st.toggle("Ad agency involved? (gross markup x1.15)", value=False, key="agency_involved",
                                     on_change=_clear_ai_section, args=("basics",))
    with col2:
        logo_file = st.file_uploader("Client logo", type=["png", "jpg", "jpeg"],
                                      key="logo_upload")
        injected_logo = test_mode_upload("logo_upload_path")
        if injected_logo is not None:
            logo_file = injected_logo
        # An uploader's own value is collected the moment the seller opens
        # another page, and it is one of the widgets Streamlit refuses to let
        # session_state write back -- so the bytes are kept beside it instead.
        # Without this, walking over to the Audience finder and back silently
        # reverted the cover to the placeholder, which is exactly the failure
        # "logo_used" exists to distinguish elsewhere.
        if logo_file is not None:
            st.session_state["uploaded_logo"] = {
                "name": logo_file.name, "bytes": logo_file.getvalue()}
        uploaded_logo = st.session_state.get("uploaded_logo")

        # A file_uploader can't be prefilled from session_state, so a
        # proposal loaded from History carries its stored logo as a path
        # instead: it's used unless a new file is uploaded over it, which is
        # what makes a reloaded proposal rebuild with the logo it shipped
        # with rather than silently reverting to the placeholder.
        restored_logo_path = st.session_state.get("restored_logo_path")
        if logo_file is None and uploaded_logo:
            drop, keep = st.columns([1, 1])
            drop.caption(f"Using **{uploaded_logo['name']}**. Upload another to replace it.")
            if keep.button("Remove logo", use_container_width=True):
                st.session_state.pop("uploaded_logo", None)
                st.rerun()
        elif logo_file is None and restored_logo_path:
            st.caption(f"Using the logo stored with this proposal "
                       f"(`{Path(restored_logo_path).name}`). Upload one to replace it.")
        elif logo_file is None and st.session_state.get("restored_logo_missing"):
            st.warning("This proposal had a logo, but it couldn't be fetched from storage — "
                       "the placeholder will be used unless you upload one.")
        st.caption("Meeting notes go in **Draft from notes** at the top of the page — that's "
                   "the copy Claude reads, and it's kept with the proposal history. "
                   "Client-facing wording lives in Campaign Specs below.")

    vertical_key = VERTICALS[vertical_choice]
    # The ORIGINATING market's label. Still what the audience finder shows and
    # still the fallback everywhere below -- but no longer the answer to
    # "where does this campaign run", which is what the target markets say.
    market_label = "Washington, DC DMA" if market_choice == "DC" else "Harrisburg DMA"
    target_labels = target_market_labels(target_dmas, market_profile_rows)
    # Before the Campaign Specs widgets render, or Streamlit raises.
    apply_geography_autofill(geography_default_text(target_labels, market_label))
    # ONE geo default, computed once and used by all three surfaces that show
    # it: the avails table (D2), the media plan's Geo column (Section E) and
    # the Campaign Specs Geography field it is derived from. They describe the
    # same fact, and a deck saying "Denver" in one place and "Washington, DC
    # DMA" in another is the kind of contradiction a client notices before
    # anyone here does.
    #
    # Read from session_state rather than from the Geography widget's return
    # value because D2 renders BEFORE Campaign Specs -- the widget doesn't
    # exist yet at that point, but apply_geography_autofill has already put
    # the value there.
    default_geo = geo_column_default(
        target_labels, st.session_state.get("geography_text", ""), market_label)

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
        if total_tv:
            # Indented directly beneath its own checkbox, so it reads as part
            # of Total TV rather than as a panel floating between two
            # unrelated products.
            _, nested = st.columns([0.05, 0.95])
            with nested:
                render_broadcast_schedule_import()
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
            # Off by default: a package slide is what was asked for, and the
            # matching MRI viewership chart used to come along with it, so a
            # four-sport proposal pulled nine slides. One toggle for the whole
            # deck rather than one per sport -- showing the stats for two of
            # four packages reads as an omission, not a choice.
            st.checkbox("Include viewership stats", value=False,
                        key="include_sport_viewership",
                        help="Adds the audience/viewership stats slide for every selected "
                             "sport, plus the Live Sports Viewers overview.",
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
        # Two verticals sell their own named measurement product instead of
        # the generic CRM sales attribution. Driven by the vertical alone, not
        # by the Sales attribution checkbox beside it: it's a different product
        # with its own slide, and a client buying Polk Signals is told so
        # whether or not a CRM upload is also on the table. Unchecked, the
        # generic line and slide behave exactly as they do for every other
        # vertical.
        va_spec = assembly.vertical_attribution_spec(vertical_key)
        vertical_attribution = None
        if va_spec:
            vertical_attribution = st.checkbox(
                va_spec["label"], value=True,
                key=f"vertical_attribution_{vertical_key}",
                help=f"Includes the {va_spec['label']} slide and lists it "
                     f"under \"Included with Campaign\" in place of the generic sales "
                     f"attribution line.",
                on_change=_clear_ai_section, args=("attribution",))

    # ---------------- Section D2: Audiences & avails ----------------
    avails_rows = []
    # Defaults for the case where the avails table isn't shown at all -- the
    # generate block below reads these unconditionally.
    avails_basis = AVAILS_BASIS_MONTHLY
    avails_label = AVAILS_COLUMN_MONTHLY
    avails_months = 1
    if include_avails_template:
        st.header("D2. Audiences & avails")
        ai_section_badge("avails")
        if "avails_version" not in st.session_state:
            st.session_state["avails_version"] = 0

        # Several markets normally sell as separate lines, so they get a row
        # each. The toggle is only worth showing when there's something to
        # combine; below two markets it would be a control that does nothing.
        combine_markets = False
        if len(target_labels) > 1:
            combine_markets = st.checkbox(
                "Combine markets into one row", key="avails_combine_markets",
                help="Off (the default) gives each target market its own avails "
                     "row and its own campaign line, which is how most of these "
                     "are built. On puts them all on one row, for when several "
                     "markets sell as a single campaign line.")

        # Seeds one row per market, replacing the originating-market default.
        # Only ever replaces rows this has written before -- see
        # apply_avails_autofill. Off/on is now also the fan-out control at
        # GROUP-creation time: seed_rows_to_groups (via sync below) turns N
        # single-market rows into N one-market groups, and one combined row
        # into a single group spanning all of them -- no separate group-side
        # logic needed for that, the existing flat-row projection already
        # produces the right shape.
        if apply_avails_autofill(
                avails_rows_for_markets(target_labels, default_geo, combine_markets)):
            st.session_state["avails_version"] += 1

        # The month count comes from form_flight_months() rather than from
        # main()'s own n_months, which isn't computed until Section E further
        # down -- the same reason apply_draft_to_form uses it.
        _, avails_active_months = form_flight_months()
        avails_months = max(1, len(avails_active_months))
        avails_basis = st.radio(
            "Avails basis", [AVAILS_BASIS_MONTHLY, AVAILS_BASIS_FLIGHT],
            horizontal=True, key="avails_basis",
            help="How the numbers below are shown and printed on the targeting slide. "
                 "They're always stored monthly, so switching back and forth changes "
                 "nothing but the presentation.")
        avails_label = avails_column_label(avails_basis, avails_months)

        # A group edit made just above (the market autofill) has to reach
        # the table this run, not next -- sync BEFORE reading groups, not
        # only in Section E. (Phase 6 adds a still-earlier call, in Section
        # A ahead of the market picker; this is the D2 one that call's own
        # comment refers to as already existing.)
        sync_targeting_groups()
        groups = st.session_state.get("targeting_groups") or []

        def _group_markets(group):
            """The Markets cell for one group -- its own picked markets when
            it has them; else the market NAMES it resolved to via the Phase 6
            geo-definition expander's Counties/Zips/Radius modes; else its
            free-text geo label as a single chip, so a group migrated from a
            flat row or seeded before this phase still shows something
            editable rather than a blank cell. A PURE function of the group
            (deterministic, no I/O beyond a name lookup) -- the fold-back
            below depends on that to detect an untouched cell."""
            geo_def = group.get("geo_def") or {}
            if geo_def.get("kind") == "markets":
                return list(geo_def.get("markets") or [])
            resolved = group.get("resolved_markets") or []
            if resolved:
                return [_market_display_name(m) for m in sorted(resolved)]
            label = tg.geo_label(group, label_for=_market_display_name)
            return [label] if label else []

        shown_before_by_gid = {
            group["id"]: avails_to_display(group.get("avails_monthly", 0), avails_basis, avails_months)
            for group in groups
        }
        default_rows = []
        for group in groups:
            row = {"gid": group["id"], "Audience": tg.audience_label(group),
                  "Markets": _group_markets(group),
                  "Label": tg.geo_label(group, label_for=_market_display_name),
                  "Color": group.get("color", "")}
            row[avails_label] = shown_before_by_gid[group["id"]]
            default_rows.append(row)
        default_avails = (pd.DataFrame(default_rows) if default_rows
                          else pd.DataFrame(columns=["gid", "Audience", "Markets", "Label", avails_label, "Color"]))
        # The basis is part of the editor key: it renames a column, and a
        # data_editor handed a different schema under the same key keeps the
        # old one.
        avails_editor_key = (f"avails_editor_{st.session_state['avails_version']}"
                             f"_{'flight' if avails_basis == AVAILS_BASIS_FLIGHT else 'monthly'}")
        market_options = sorted({m for group in groups for m in _group_markets(group)} | set(target_labels))
        avails_df = st.data_editor(
            default_avails, num_rows="dynamic", key=avails_editor_key, use_container_width=True,
            on_change=_clear_ai_section, args=("avails",),
            column_config={
                # Hidden, not shown to the rep -- the group id a row is
                # backed by, the same round-tripping hidden-column mechanic
                # `_group_ids` uses on the media plan grid. `None` hides it
                # but still returns it through an edit, a delete or a new row.
                "gid": None,
                "Markets": st.column_config.MultiselectColumn(
                    "Markets", options=market_options, accept_new_options=True,
                    help="One row can span several markets -- add more than one here "
                         "for a group that sells as a single campaign line."),
                # Same field as the geo-definition expander's "Label
                # (optional)" text input -- this is just a second place to
                # see and edit it, since several Zips-mode rows resolving to
                # the same market otherwise all show as identical "Saint
                # Louis" Markets chips with no way to tell them apart at a
                # glance. Defaults to the auto-derived summary (market plus
                # zip count) exactly like the plan table's own Geo cell.
                "Label": st.column_config.TextColumn(
                    "Label", help="What the plan table's Geo column and the targeting slide show "
                                  "for this line -- market plus zip count by default. Edit to relabel; "
                                  "clear it to go back to the derived summary."),
                # Hidden until the targeting map (Prompt E) exists to show
                # it -- a raw hex string is noise in a rep-facing table with
                # no map to key it against. Same hidden-column mechanic as
                # "gid" above: the value stays on the group and still
                # round-trips through an edit, a delete or a new row: only
                # the COLUMN comes back with the map.
                "Color": None,
            },
        )
        avails_df[avails_label] = avails_df[avails_label].fillna(0)
        # The MultiselectColumn new-row trap: a row added via the grid's own
        # "+" comes back with Markets = None, not [] -- normalize on read,
        # same discipline group_ids_of uses for _group_ids.
        avails_df["Markets"] = avails_df["Markets"].apply(lambda v: list(v) if isinstance(v, list) else [])
        total_avails_val = int(avails_df[avails_label].sum())
        st.caption(f"Total avails ({'full flight' if avails_basis == AVAILS_BASIS_FLIGHT else 'monthly'}): "
                   f"{total_avails_val:,}")

        # Fold back into groups, keyed on gid -- NOT position, which a merge
        # of add/delete/reorder can't be trusted to preserve. A row missing
        # its gid (added via the grid's own "+") gets a fresh one from
        # new_group; a group whose gid isn't among the returned rows was
        # deleted and simply isn't carried into new_groups.
        groups_by_gid = {group["id"]: group for group in groups}
        new_groups = []
        for position, (_, row) in enumerate(avails_df.iterrows()):
            gid = row.get("gid")
            gid = gid if isinstance(gid, str) and gid.strip() else None
            prior = groups_by_gid.get(gid)
            # `_cell_unchanged` (see its docstring, right above
            # restore_untouched_avails): the Audience cell can only DISPLAY a
            # multi-term group as `audience_label`'s plain, comma-joined
            # text, not the canonical form `terms_from_audience_text`
            # recognizes, so only a real edit re-derives terms/op -- an
            # untouched cell keeps them byte-for-byte.
            audience_text = str(row["Audience"] or "").strip()
            audience_unchanged = prior is not None and _cell_unchanged(
                tg.audience_label(prior), audience_text)
            if audience_unchanged:
                terms, op = prior["terms"], prior["op"]
            else:
                terms, op = tg.terms_from_audience_text(audience_text)

            # Same test, on Markets/geo_def -- and a worse hazard if it were
            # skipped, because Phase 6's geo-definition expander (not this
            # grid) is what populates resolved_zips/resolved_markets via real
            # geo resolution, and this grid can only DISPLAY the result
            # (market names for a counties/zips/radius-kind group) through
            # `_group_markets`'s fallback, never the geo_def that produced
            # it. Only a real edit to the Markets cell re-derives geo_def;
            # an untouched cell keeps geo_def AND both resolved fields
            # byte-for-byte.
            markets_now = row["Markets"]
            markets_unchanged = prior is not None and _cell_unchanged(_group_markets(prior), markets_now)
            if markets_unchanged:
                geo_def = prior.get("geo_def")
                resolved_zips = prior.get("resolved_zips") or []
                resolved_markets = prior.get("resolved_markets") or []
            elif markets_now:
                geo_def = {"kind": "markets", "markets": markets_now}
                resolved_zips, resolved_markets = [], []
            else:
                geo_def = {"kind": "text", "label": ""}
                resolved_zips, resolved_markets = [], []

            # Same fold-back test as Audience/Markets above: the cell can
            # only DISPLAY the group's CURRENT resolved label (market plus
            # zip count, or an existing override) -- it never round-trips
            # geo_def/resolved_markets on its own, so an untouched cell must
            # keep whatever name (possibly none) it already had rather than
            # writing today's derived text in as a permanent override, which
            # would freeze it against every future market/zip change.
            label_text = str(row.get("Label", "") or "").strip()
            label_unchanged = prior is not None and _cell_unchanged(
                tg.geo_label(prior, label_for=_market_display_name), label_text)
            name = (prior.get("name", "") if prior else "") if label_unchanged else label_text

            monthly = restore_untouched_avails(
                prior.get("avails_monthly", 0) if prior else 0,
                shown_before_by_gid.get(gid), row[avails_label], avails_basis, avails_months)
            built = tg.new_group(
                terms, op=op, geo_def=geo_def,
                name=name,
                avails_monthly=monthly,
                color=(prior.get("color") if prior else tg.assign_color(position)),
                group_id=gid,
            )
            built["resolved_zips"] = resolved_zips
            built["resolved_markets"] = resolved_markets
            new_groups.append(built)
        # A real Label-cell edit changes group["name"], the same field the
        # geo-definition expander's own "Label (optional)" text_input writes
        # -- and that widget's session_state, once created, beats a fresh
        # `value=` on every later rerun (the same trap `bump_plan_options_
        # generation` exists for). Folding a grid edit into `name` without
        # also moving the expander widget to a new key would have the
        # expander's stale cached text silently overwrite the grid's edit
        # the moment this same run reaches `render_group_geo_expander`.
        # Bumped only when a name actually changed here -- an expander edit
        # from the PRIOR run is already reflected in `prior`/`groups` by the
        # time this loop runs, so it compares equal and never bumps, which
        # is what keeps that widget's own state authoritative for its own
        # edits.
        if any((groups_by_gid.get(g["id"], {}).get("name") or "") != (g.get("name") or "")
               for g in new_groups):
            st.session_state["group_name_generation"] = st.session_state.get("group_name_generation", 0) + 1
        # Never avails_seed_rows directly here -- sync_targeting_groups (next
        # called in Section E) projects THIS write down to the flat shape,
        # the mirror image of how a flat-row change used to flow up into
        # groups. Writing both here would race the projection.
        st.session_state["targeting_groups"] = new_groups
        # A row the grid's own "+" just added exists ONLY as this run's
        # in-progress edit delta, tracked by the widget under today's
        # avails_editor_key -- the same "a keyed widget's session_state
        # entry beats its value= argument" trap bump_plan_options_generation
        # exists for, here for st.data_editor instead of a plain widget: the
        # newly-assigned gid is real in `targeting_groups` from this line
        # down, but the widget has no way to learn its own delta just became
        # a real backing row, so it keeps re-offering (or silently drops) the
        # same in-progress content forever under the OLD key. A row deleted
        # via the grid's own "-" is the same hazard the other way. Bumping
        # the generation on any row-count change -- never on an ordinary
        # cell edit, which must NOT force a fresh widget -- moves the editor
        # to a new key next run, so it re-mounts from the fresh baseline
        # (the new group, with its real gid) instead of a stale delta.
        if len(new_groups) != len(groups):
            st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
            st.rerun()

        for group in new_groups:
            label = tg.audience_label(group)
            if label.strip():
                avails_rows.append({
                    "audience": label,
                    "geo": tg.geo_label(group, label_for=_market_display_name),
                    # What the client sees, in the basis on screen -- read
                    # back off the SAME group, not re-derived, so this can't
                    # disagree with what the grid just showed.
                    "avails": f"{avails_to_display(group['avails_monthly'], avails_basis, avails_months):,}",
                    "avails_monthly": group["avails_monthly"],
                })

        catalog_rfp_lookup = dict(zip(audience_catalog["segment"], audience_catalog["rfp_selectable"]))
        custom_count = tg.custom_segment_count(new_groups, catalog_rfp_lookup)
        if custom_count > 1:
            st.warning(f"{custom_count} custom (non-RFP-selectable) audiences are in play across your "
                       f"targeting groups above -- only one is allowed per campaign. Review before generating.")

        # One geo-definition expander per real group -- Counties/Zips/Radius
        # resolution, beside the grid's own quick Markets cell rather than
        # replacing it. A market-only group (no audience yet) still gets one:
        # geography can be defined before the audience is picked.
        for group in new_groups:
            render_group_geo_expander(group)

        # The finder builds groups directly now (Phase 5), so it needs the
        # SAME geo default the table around it seeds new groups with --
        # otherwise a segment added through the finder starts a group with a
        # different default geo from the row above it.
        render_audience_finder(avails_df, default_geo, vertical_key)

        with st.expander("📄 Import an avails PDF (from Salesforce)", expanded=False):
            st.caption("Upload the Premion avails export for this buy -- one targeting group "
                       "is added per audience/geography pair, client/agency/flight are filled in "
                       "where they're not already set, and Section D's attribution toggles get "
                       "a floor from what the document lists (never a ceiling -- nothing already "
                       "on is ever turned off).")
            render_avails_pdf_uploader("d2", "The avails PDF you pulled from Salesforce for this buy.")

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
        geography_text = st.text_area("Geography", height=90, key="geography_text",
                                       on_change=_clear_ai_section, args=("specs",))
    with spec_col2:
        budget_text = st.text_area("Budget & Allocation", height=90, key="budget_text",
                                    on_change=_clear_ai_section, args=("specs",))
        placements_text = st.text_area("Placements & Creative", height=90, key="placements_text",
                                        on_change=_clear_ai_section, args=("specs",))
        timing_text = st.text_area("Timing", height=90, key="timing_text",
                                    help="Narrative copy for the Campaign Specs slide. Actual flight dates for the media plan are set below.",
                                    on_change=_clear_ai_section, args=("specs",))

    default_targeting = audience_stack(audience_text)
    # default_geo was computed up in Section A, from the same session_state
    # value this widget just returned -- recomputing it here would be a second
    # definition of one fact, which is the bug this whole change is about.

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
        flight_start = st.date_input("Flight start", value=DEFAULT_FLIGHT_START, key="flight_start",
                                      on_change=_clear_ai_section, args=("flight",))
    with fcol2:
        flight_end = st.date_input("Flight end", value=DEFAULT_FLIGHT_END, key="flight_end",
                                    on_change=_clear_ai_section, args=("flight",))

    all_months = month_list(flight_start, flight_end)
    # A keyed multiselect ignores its `default=` once session_state holds a
    # value, so months chosen for the OLD flight survive a change of dates.
    # When the flight moves somewhere else entirely they're all invalid, the
    # selection filters down to nothing, and the flight label becomes "TBD" --
    # which is what a drafted proposal hit, since a draft sets new dates over
    # whatever the form was showing. An empty intersection means the stored
    # choice is about a different flight, so the whole new range is selected;
    # a partial overlap is a real custom-flighting choice and is kept.
    stored_months = st.session_state.get("active_months")
    if stored_months is not None and not [m for m in stored_months if m in all_months]:
        st.session_state["active_months"] = all_months
    active_months = st.multiselect(
        "Active months (uncheck to skip a month -- custom flighting)",
        all_months, default=all_months, key="active_months",
    )
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
    # The plan follows the GROUPS' grouping now, not the flat avails table
    # directly -- sync first, so a group edited earlier in this same run
    # (D2 already rendered above) is what seeds from, not a stale value from
    # before this run's edits. plan_lines_from_groups falls back to the flat
    # row grouping automatically: sync_targeting_groups keeps targeting_groups
    # derived from avails_seed_rows whenever nothing group-aware has written
    # it directly, so this is exactly today's grouping for every proposal
    # that predates targeting groups.
    sync_targeting_groups()
    plan_lines = plan_lines_from_groups(
        st.session_state.get("targeting_groups"), default_geo, default_targeting)
    shared_fields_key = (default_targeting + "||"
                         + " / ".join(f"{a}@{g}" for a, g, _gid in plan_lines)
                         + "||" + flight_label)

    # An imported schedule adds one more line to every option. It's part of
    # the seed key (via read_seed_selections) so importing or removing a
    # schedule re-seeds exactly the way ticking a product does.
    schedule = st.session_state.get("broadcast_schedule")
    broadcast_warning = None

    def _broadcast_row_for_option(breakout):
        """The broadcast line in one option's own basis.

        **The schedule slides' breakout has no say here.** That toggle is
        about how the schedule grid is laid out; this is about what basis the
        media plan is quoted in, and they're different questions about
        different slides. The plan's own Monthly/Full Flight radio decides.
        """
        if not schedule:
            return None, None
        # market_label, NOT default_geo, and deliberately. This argument is
        # only reached when the station's call sign isn't recognised, and a
        # broadcast line runs where its STATION is -- never where the campaign
        # is targeted. Handing it the target markets would put "Denver" on a
        # line carrying a Washington station's spots.
        return broadcast_row_for(
            schedule, st.session_state.get("broadcast_plan_desc", ""), market_label,
            breakout == BREAKOUT_MONTHLY, n_months, flight_label)

    if schedule:
        _, broadcast_warning = _broadcast_row_for_option(BREAKOUT_MONTHLY)

    def _seed_option_rows(breakout=BREAKOUT_MONTHLY):
        rows = seed_media_plan_rows(seed_selections, plan_lines, default_targeting, flight_label)
        broadcast_row, _ = _broadcast_row_for_option(breakout)
        if broadcast_row is not None:
            # The imported schedule IS the broadcast buy, so it replaces the
            # rate-card line Total TV seeds rather than sitting beside it --
            # two broadcast lines would double-count the same spots.
            rows = [r for r in rows
                    if str(r.get("Tactic", "")).strip() != BROADCAST_PRODUCT_LABEL]
            rows.append(dict(broadcast_row))
        return rows

    # .get() rather than [] on the two seed keys: they're written alongside
    # plan_options everywhere that sets it, but a missing key should re-seed
    # rather than raise.
    if not st.session_state.get("plan_options"):
        seeded = _seed_option_rows()
        st.session_state["plan_options"] = [new_plan_option(DEFAULT_OPTION_NAMES[0], seeded)]
        st.session_state["_seeded_tactics"] = [r["Tactic"] for r in seeded]
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
    elif st.session_state.get("_product_seed_key") != product_seed_key:
        # Product selections changed. **Diff, never rebuild.**
        #
        # This used to reseed every option's rows from scratch, which meant
        # ticking any Section C box -- or importing a schedule -- silently
        # zeroed a drafted plan: allocations, negotiated CPMs and all. The
        # rule now is that a product toggle only ever touches rows belonging
        # to the product that changed. Everything else, drafted or typed or
        # untouched, is left exactly as it is.
        for opt in st.session_state["plan_options"]:
            fresh = _seed_option_rows(opt["breakout"])
            _apply_product_diff(opt, fresh, st.session_state.get("_seeded_tactics") or [])
            opt["version"] += 1
        st.session_state["_seeded_tactics"] = [r["Tactic"] for r in _seed_option_rows()]
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
    elif st.session_state.get("_shared_fields_key") != shared_fields_key:
        # A clean row KEEPS its own Geo when that Geo is still one the plan
        # covers. Handing every row the single default here would collapse
        # three per-market lines back onto one joined string the moment the
        # audience or the flight changed -- undoing the grouping this seeding
        # exists to honour, and looking for all the world like a bug in the
        # avails table. A row whose Geo is no longer on the plan (its market
        # was deselected) takes the default, which is what re-seeding is for.
        # Both halves of a line's identity are preserved, for the same
        # reason. Handing every clean row the single default would collapse
        # six audience-by-market lines back onto one -- undoing the grouping
        # and looking exactly like a bug in the avails table. A row keeps its
        # own Geo and its own Targeting while each is still one the plan
        # covers, and takes the default only when that market or audience has
        # gone from the table, which is what re-seeding is for.
        valid_geos = {geo for _, geo, _gid in plan_lines}
        valid_audiences = {aud for aud, _, _gid in plan_lines if aud}
        valid_group_ids = {gid for _, _, gid in plan_lines if gid}
        for opt in st.session_state["plan_options"]:
            # A CLEAN, group-backed row whose group is now COMPLETELY gone
            # (every id it carries, not just some -- a merged row can
            # partially survive, handled below exactly as before) is
            # removed outright rather than reset to the shared default.
            # This is the deletion counterpart of queue_group_seed/
            # seed_pending_group_rows above: a row seeded FOR a group must
            # not outlive that group just because nothing else about the
            # plan changed. Reset-to-default (the old, only, behavior) left
            # a deleted avails row's plan line on the grid forever, zombied
            # into a blank/default-targeted "Premion Streaming TV" line
            # that still counted toward the total and still made the deck.
            # An edit protects a row from every kind of reseed, deletion
            # included -- the same dirty check every other branch here uses.
            kept = [i for i, (row, dirty) in enumerate(zip(opt["rows"], opt["dirty"]))
                   if dirty or not group_ids_of(row)
                   or any(gid in valid_group_ids for gid in group_ids_of(row))]
            if len(kept) != len(opt["rows"]):
                opt["rows"] = [opt["rows"][i] for i in kept]
                opt["dirty"] = [opt["dirty"][i] for i in kept]
                opt["driver"] = [opt["driver"][i] for i in kept]
            for row, dirty in zip(opt["rows"], opt["dirty"]):
                if not dirty:
                    ids = group_ids_of(row)
                    if ids:
                        # Matched by GROUP ID, never by string -- two rows
                        # that happen to render the same Geo text (two
                        # audiences in the same market) must never re-seed
                        # onto each other just because their text matches.
                        # All ids still live -> the row keeps its own
                        # Geo/Targeting untouched; a dead id (its group was
                        # removed from the campaign) takes the default.
                        if all(gid in valid_group_ids for gid in ids):
                            row_geo, row_audience = row.get("Geo"), row.get("Targeting")
                        else:
                            row_geo, row_audience = plan_lines[0][1], default_targeting
                    else:
                        # No ids at all -- a proposal loaded before targeting
                        # groups existed, or a broadcast row (which is never
                        # group-backed and is protected below regardless, via
                        # current=row). Exactly today's string matching.
                        row_geo = (row.get("Geo") if row.get("Geo") in valid_geos
                                   else plan_lines[0][1])
                        row_audience = (row.get("Targeting")
                                        if row.get("Targeting") in valid_audiences
                                        else default_targeting)
                    row.update(resolve_row_defaults(
                        row.get("Tactic", ""), row_geo, row_audience,
                        flight_label, current=row))
            opt["version"] += 1
        st.session_state["_shared_fields_key"] = shared_fields_key

    # A group `queue_group_seed` marked this run (a fresh avails-PDF import,
    # or the Audience finder's "New group" button) gets its own line now,
    # independent of whichever branch above did or didn't fire -- see
    # queue_group_seed's docstring for why this can't just be folded into
    # the shared-fields branch above.
    seed_pending_group_rows(st.session_state["plan_options"], _seed_option_rows)

    # An option's own Monthly/Full Flight choice changes what basis its rows
    # are quoted in, and the broadcast line is derived from a fixed set of
    # Wide Orbit totals rather than typed -- so it follows that choice, on
    # any option the seller hasn't hand-edited. The schedule slides' own
    # breakout is a separate question and deliberately has no effect here.
    if schedule:
        for opt in st.session_state["plan_options"]:
            if opt.get("_broadcast_basis") == opt["breakout"]:
                continue
            fresh, _ = _broadcast_row_for_option(opt["breakout"])
            if fresh is not None:
                for index, row in enumerate(opt["rows"]):
                    if is_broadcast_row(row) and not opt["dirty"][index]:
                        row.update({k: fresh[k] for k in ("Impressions", "Cost", "CPM")})
                        opt["version"] += 1
            opt["_broadcast_basis"] = opt["breakout"]

    plan_options = st.session_state["plan_options"]

    # A markup change (the agency toggle) re-derives every row of every
    # option -- it moves the cost/impressions relationship itself, not just
    # one row's cells.
    markup_changed = st.session_state.get("media_plan_markup") != markup
    st.session_state["media_plan_markup"] = markup

    if broadcast_warning:
        st.warning(broadcast_warning)
    if schedule:
        # The two flights are set independently -- the plan's on the form,
        # the schedule's by Wide Orbit -- and nothing forces them to agree.
        # Often they shouldn't: a four-week broadcast burst inside a longer
        # streaming campaign is a normal buy. So this names both spans and
        # leaves the seller to decide, rather than correcting either.
        schedule_months = {week.strftime("%b %Y") for week in schedule.grid_weeks}
        if schedule_months and not schedule_months <= set(all_months):
            st.warning(
                f"This schedule runs **{', '.join(sorted(schedule_months, key=_month_sort))}** "
                f"but the proposal's flight is **{flight_label}**. The broadcast line is priced "
                f"off the schedule's own dates, so the two don't have to match — just check "
                f"they're meant to differ.")
    if schedule and BROADCAST_EXCLUDED_FROM_AGENCY_MARKUP:
        st.caption(
            f"📺 The **{BROADCAST_TACTIC_MARKER}** line uses the Wide Orbit cost exactly as "
            f"quoted — broadcast is already gross, so it is never marked up by the ×1.15 "
            f"agency uplift"
            + (", even though the agency toggle is on." if agency_involved else ".")
            + f" Renaming that line so it no longer says \"{BROADCAST_TACTIC_MARKER}\" would "
              f"put it back under the markup.")

    # The Targeting column is per line, and every line starts from the same
    # Campaign Specs Audience field -- which is right for a plan whose lines
    # all reach one audience, and wrong the moment a plan splits a budget
    # across distinct tracks (the Commercial/Retail case). Nothing can detect
    # that from the form, so it's said out loud instead of guessed at.
    st.caption(
        "🎯 **Targeting** is per line. Rows seed it from the Campaign Specs Audience "
        "field (Live Sports and Retargeting carry their own standard copy), so if this "
        "plan splits its budget across distinct audiences, edit each row to the "
        "audience that row actually reaches.")

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

            # --- Add lines, from what the avails table currently holds ----
            # Deliberately thin: it reads the avails rows on every run rather
            # than keeping a copy. When targeting groups land (they own this
            # data next) the source changes and the interaction doesn't.
            # Read the avails TABLE, not `avails_rows`: that list is the
            # deck-facing one and drops rows with no audience yet -- which is
            # exactly what a freshly market-seeded table is full of, so the
            # markets a rep just picked would be missing from the Geo list.
            avails_table = st.session_state.get("avails_seed_rows") or []
            avail_audiences = [a for a in dict.fromkeys(
                str(r.get("Audience", "")).strip() for r in avails_table) if a]
            avail_geos = [g for g in dict.fromkeys(
                str(r.get("Geo", "")).strip() for r in avails_table) if g]
            catalog_labels = sorted({label for label, _ in _quick_add_catalog().values()})

            with st.expander("Add lines", expanded=False):
                qa1, qa2, qa3 = st.columns(3)
                with qa1:
                    pick_products = st.multiselect(
                        "Product", catalog_labels, key=f"qa_prod_{gen}_{idx}",
                        accept_new_options=True, placeholder="Search products...")
                with qa2:
                    pick_audiences = st.multiselect(
                        "Audience", avail_audiences, key=f"qa_aud_{gen}_{idx}",
                        accept_new_options=True,
                        placeholder="From the avails table, or type your own")
                with qa3:
                    pick_geos = st.multiselect(
                        "Geo", avail_geos, key=f"qa_geo_{gen}_{idx}",
                        accept_new_options=True,
                        placeholder="From the avails table, or type your own")
                qa_combine = False
                if len(pick_geos) > 1:
                    qa_combine = st.checkbox(
                        "Combine into one line", key=f"qa_combine_{gen}_{idx}",
                        help="Off (the default) gives each market its own line. "
                             "On puts them on a single line, for when several "
                             "markets sell as one campaign line.")
                preview = quick_add_rows(
                    pick_products, pick_audiences, pick_geos, flight_label,
                    default_targeting, default_geo, qa_combine)
                if preview:
                    st.caption(f"Adds {len(preview)} line"
                               f"{'s' if len(preview) != 1 else ''}.")
                if st.button("Add lines", disabled=not preview,
                             key=f"qa_add_{gen}_{idx}"):
                    # Dirty on arrival, like a duplicated line: the rep chose
                    # this audience and this market, and a shared-field
                    # re-seed overwriting that choice is the bug the dirty
                    # flag exists for.
                    option["rows"] = option["rows"] + preview
                    option["dirty"] = option["dirty"] + [True] * len(preview)
                    option["version"] += 1
                    st.rerun()

            # Editing a line is a pick rather than a retype. Options carry the
            # values already ON the grid as well as the avails ones, so
            # nothing existing becomes unselectable -- including the imported
            # broadcast row's call-sign Geo, which comes from neither list.
            grid_geos = [g for g in dict.fromkeys(
                list(avail_geos) + [str(r.get("Geo", "")) for r in option["rows"]]) if g]
            grid_audiences = [a for a in dict.fromkeys(
                list(avail_audiences) + [str(r.get("Targeting", "")) for r in option["rows"]]) if a]

            edited_df = st.data_editor(
                pd.DataFrame(option["rows"]), num_rows="dynamic",
                key=f"media_plan_editor_{idx}_{option['version']}", use_container_width=True,
                column_config={
                    "Impressions": st.column_config.NumberColumn(f"Impressions ({basis})"),
                    "Type": st.column_config.SelectboxColumn(options=[ROW_TYPE_RATE, ROW_TYPE_FLAT_FEE]),
                    "Cost": st.column_config.NumberColumn(f"Cost ({basis}, $)", format="$%.0f"),
                    "Geo": st.column_config.SelectboxColumn("Geo", options=grid_geos),
                    "Targeting": st.column_config.SelectboxColumn(
                        "Targeting", options=grid_audiences),
                    # Hidden, not shown to the rep: the group id(s) a row is
                    # backed by, used only to match a clean row back to its
                    # group when a shared field changes. `None` hides the
                    # column but still round-trips its value through an edit,
                    # a delete or a newly-added row (verified against this
                    # Streamlit version) -- unlike leaving it out of
                    # column_config entirely, which would render it as a
                    # raw, visible, editable list column.
                    "_group_ids": None,
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

            # Merge two or more lines into one (e.g. the same product sold to
            # two audiences that turned out to price the same way), and split
            # a merged line back apart. Neither touches
            # st.session_state["targeting_groups"] -- see merge_plan_rows /
            # split_plan_row for why that's the assertion, not a detail.
            mcol1, mcol2 = st.columns([3, 1])
            with mcol1:
                merge_pick = st.multiselect(
                    "Lines to merge", tactic_labels, label_visibility="collapsed",
                    key=f"tg_merge_pick_{idx}", placeholder="Pick two or more lines to merge")
            with mcol2:
                if st.button("Merge lines", disabled=len(merge_pick) < 2, key=f"tg_btn_merge_{idx}"):
                    merge_plan_rows(option, [int(p.split(":")[0]) for p in merge_pick], markup)
                    option["version"] += 1
                    st.rerun()

            splittable = [i for i, row in enumerate(rows_now) if len(group_ids_of(row)) > 1]
            split_labels = [tactic_labels[i] for i in splittable]
            if split_labels:
                scol1, scol2 = st.columns([3, 1])
                with scol1:
                    split_pick = st.selectbox(
                        "Line to split", split_labels, label_visibility="collapsed",
                        key=f"tg_split_pick_{idx}")
                with scol2:
                    if st.button("Split line", key=f"tg_btn_split_{idx}"):
                        split_plan_row(option, int(split_pick.split(":")[0]))
                        option["version"] += 1
                        st.rerun()
                st.caption("Splits impressions and cost evenly across the new lines -- if this line was "
                           "merged from parts of different sizes, splitting won't recover their original "
                           "amounts.")

            totals = compute_plan_totals(
                option["rows"], breakout_mode, n_months, flight_label,
                broadcast_months=(schedule.active_month_count() if schedule else None))
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
        va_spec["label"] if vertical_attribution else None,
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
    show_cpm_column = st.checkbox(
        "Show CPM column on the media plan", value=True, key="show_cpm_column",
        help="Adds a CPM column between Impressions and Cost. Turn it off for a client who "
             "shouldn't see rates broken out.")
    with tcol2:
        # History-only annotation: it labels the row, never the deck. Loading
        # a proposal prefills a thread-position default, since that's the
        # case where a label is actually worth having.
        revision_label = st.text_input(
            "Revision label (optional)", key="revision_label",
            placeholder="Revision 2",
            help="Shown on the Proposal history page to tell this build apart from "
                 "others for the same client. It never appears in the deck.")

    # Every widget on the page now exists, so this is the complete form. Taken
    # here rather than at the foot of the function so that whatever Generate
    # does below -- rerun, exception, a long assembly -- a snapshot of what the
    # seller typed is already safe.
    snapshot_form_state()

    if st.button("Generate proposal", type="primary"):
        selections = {
            "preset": preset_key,
            "market": market_choice,
            # The DMAs the campaign is aimed at, in the order the rep picked
            # them, which is the order their profile slides appear. Ordered,
            # so a list rather than a set. Independent of "market" above,
            # which is the originating station -- a DC proposal can target
            # any set of markets in the country.
            "target_dmas": target_dmas,
            "include_market_profile": include_market_profile,
            "vertical": vertical_key if include_vertical_slides else "none",
            "agency_involved": agency_involved,
            "spanish_campaign": spanish_campaign,
            "dynamic_creative": dynamic_creative,
            "tegna_positioning": tegna_positioning,
            # Swaps the master's static schedule placeholder for a generated
            # grid. False until a Wide Orbit export has actually been read,
            # so a Total TV deck with no schedule behaves exactly as it does
            # today.
            "broadcast_schedule_imported": bool(st.session_state.get("broadcast_schedule")),
            "include_avails_template": include_avails_template and bool(avails_rows),
            "products": products_selection,
            # None for every vertical that has no branded attribution product,
            # which is also what a proposal logged before this existed carries
            # -- build_presentation reads that as "leave the slides alone".
            "vertical_attribution": vertical_attribution,
            # Which avails/targeting template slide build_presentation keeps
            # -- the map variant (no stock photo, more room for the map) when
            # a targeting map will actually be drawn, the standard one
            # otherwise. Computed the same way the map itself is (see
            # map_png below), so the two can't disagree about whether
            # there's a map to show.
            "targeting_map_present": bool(targeting_map.groups_with_zips(
                st.session_state.get("targeting_groups") or [])),
        }
        selections["targeting_attribution"] = {
            "first_party_data": first_party_data,
            "linear_reach_extension": linear_reach_extension,
            "sales_attribution": sales_attribution,
            "brand_lift": brand_lift,
        }

        # Groups are reconciled here, right before they're saved, rather than
        # earlier in the render -- nothing in this phase's UI reads them yet,
        # so the only thing that has to be true is that a saved proposal's
        # targeting_groups agrees with the avails_rows it's saved alongside.
        sync_targeting_groups()

        if not avails_rows:
            avails_rows_final = [{"audience": "", "geo": default_geo,
                                  "avails": "0", "avails_monthly": 0}]
            total_avails_str = "0"
        else:
            avails_rows_final = avails_rows
            # Totalled in the basis on screen, so the table's own total and
            # its rows can't disagree about which basis they're in.
            total_avails_str = f"{sum(int(r['avails'].replace(',', '')) for r in avails_rows):,}"

        gross_note = " (Gross)" if agency_involved else ""
        multiple_options = len(plan_options) > 1

        def _option_payload(option, totals):
            rows = [
                {"tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
                 "impressions": "--" if r["is_flat_fee"] else f"{int(r['monthly_impressions']):,}",
                 "cost": f"${r['monthly_cost']:,.0f}{gross_note}",
                 # A flat fee has no rate, so "--" rather than a misleading $0.
                 "cpm": "--" if r["is_flat_fee"] else f"${_num(r.get('cpm')):,.2f}"}
                for r in totals["preview_rows"]
            ] or [{"tactic": "", "flight": flight_label, "geo": default_geo, "targeting": "",
                   "impressions": "0", "cost": "$0"}]

            full_flight_total = None
            if n_months > 1 and totals["preview_rows"]:
                full_flight_total = {
                    "label": f"Full Flight Total ({n_months} months)",
                    "impressions": f"{int(totals['full_flight_impressions']):,}",
                    "cost": f"${totals['full_flight_cost']:,.0f}{gross_note}",
                }

            # Blended, not averaged: the plan's own cost over its own
            # impressions, which is what a client would compute.
            #
            # Rate lines only, on BOTH sides of the division. A flat fee has a
            # cost but no impressions by definition, so counting its dollars
            # against the media lines' impressions inflates the rate into
            # something no line actually carries -- a $45,000 plan with a
            # $1,200 production fee showed $19.12 against a real media rate of
            # $18.36. Whatever this cell says, a client will divide the media
            # cost by the media impressions and expect to land on it.
            rate_cost = sum(r["monthly_cost"] for r in totals["preview_rows"]
                            if not r["is_flat_fee"])
            rate_impressions = sum(r["monthly_impressions"] for r in totals["preview_rows"]
                                   if not r["is_flat_fee"])
            blended = rate_cost / rate_impressions * 1000 if rate_impressions else 0
            return {
                "show_cpm": show_cpm_column,
                "total_cpm": f"${blended:,.2f}" if blended else "--",
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
            # Precedence: the logo uploaded in this session (which outlives
            # the uploader widget), then the one restored with a loaded
            # proposal, then the placeholder.
            "logo_path": (io.BytesIO(uploaded_logo["bytes"]) if uploaded_logo
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
                # The slide's own column header. It is literal text in the
                # template, not a token, so assembly rewrites it -- a figure
                # in front of a client must never be ambiguous about whether
                # it is monthly or the whole flight.
                "label": avails_label,
                # None whenever no targeting group has been resolved to real
                # zips yet -- assembly.place_targeting_map does nothing at
                # all in that case, and targeting_map_present above is False,
                # so the deck keeps the standard (photo) template exactly as
                # it is today. dark=True: the map template's own background
                # is the deck's dark gradient (no stock photo behind it, see
                # roadmap section E), not the white the Zip/map builder
                # page's own preview renders against -- that page calls
                # render_map itself, separately, for its own light-background
                # display.
                "map_png": targeting_map.render_map(
                    st.session_state.get("targeting_groups") or [], dark=True,
                    label_for=_market_display_name),
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
                case_study_sources.append(case_study_source(case_study))
            except Exception as exc:
                case_study_errors.append(f"{case_study['title']}: {db.describe_error(exc)}")
        for message in case_study_errors:
            st.warning(f"Case study left out -- couldn't fetch it. {message}")

        # Same reason as the case studies above: fetched before assembly so a
        # storage problem reads as its own message rather than as a failure
        # part-way through building a deck.
        market_profile_paths, market_profile_errors = market_profile_images(
            target_dmas, include_market_profile)
        for message in market_profile_errors:
            st.warning(message)

        with st.spinner("Assembling deck..."):
            try:
                prs, original_count, kept_count = assembly.build_presentation(master_path, selections)
                # Replace-and-expand: the national Total U.S. profile slide
                # becomes one slide per target market, in picker order. Runs
                # before the case studies so the index they anchor to is
                # re-derived from the token afterwards rather than going stale.
                assembly.replace_market_profile_slides(prs, market_profile_paths)
                # Before personalize, deliberately -- see
                # case_study_insert_index: the plan slide is found by its
                # {{PLAN_TITLE}} token (which personalize consumes), and
                # personalize clones it per extra option directly after the
                # original, so inserting here is what puts case studies ahead
                # of the first option rather than between options.
                case_study_slides = assembly.append_case_studies(prs, case_study_sources)
                # Before personalize: the schedule slides fill their own
                # tokens, and personalize's deck-wide pass would otherwise run
                # over the template's placeholders before they're cloned per
                # page.
                schedule_warnings = []
                schedule = st.session_state.get("broadcast_schedule")
                if schedule:
                    breakout, detailed = broadcast_display_options()
                    schedule_warnings = assembly.build_broadcast_schedule_slides(
                        prs, schedule, st.session_state.get("broadcast_plan_desc", ""),
                        breakout=breakout, detailed=detailed)
                # personalize reports overflow it couldn't shrink away (the
                # media plan table, Campaign Specs copy). These were being
                # discarded, so a deck could ship with a table running into
                # the graphic below it and nothing said so.
                #
                # Cleared first so the caption below describes THIS deck. The
                # log accumulates across a session otherwise, and a note about
                # a font some earlier proposal used is a note about nothing.
                assembly.reset_width_calibration_log()
                layout_warnings = assembly.personalize(prs, fill_data) or []
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)
            except Exception as exc:
                st.error(f"Assembly failed: {exc}")
                raise

        for message in schedule_warnings + layout_warnings:
            st.warning(message)

        # An environment note rather than a problem with this proposal: it is
        # true of every deck this machine builds. Shown as a caption so it is
        # visible without competing with the layout warnings above -- it
        # matters because the deployed instance has neither the brand fonts
        # nor Calibri, estimates text width instead, and so sizes tables (and
        # decides whether to compress the "Included with Campaign" band)
        # slightly differently from a machine that has the fonts.
        # Two different facts, so two notes rather than one. measurement_note
        # says a substitution happened; width_calibration_note says the
        # correction applied for it was not measured against that face, which
        # is what decides whether the client-name title can be trusted to sit
        # inside its box. A log line nobody opens is the phantom-slide-37
        # problem in a different costume, so it goes on the page.
        for note in (text_metrics.measurement_note(),
                     assembly.width_calibration_note()):
            if note:
                st.caption(note)

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
        if uploaded_logo:
            logo_storage_path, logo_error = db.upload_proposal_logo(
                f"{(client_name or 'client').replace(' ', '_')}", uploaded_logo["bytes"],
                uploaded_logo["name"])
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
            target_dmas=target_dmas,
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
                # Stored so a rebuild renders in the basis the client
                # actually saw. The label is stored rather than
                # recomputed, so a rebuild can never disagree with the
                # month count the original was built from.
                "avails_basis": avails_basis,
                "avails_label": avails_label,
                # Additive: a new key, ignored by anything that predates
                # targeting groups. avails_rows above is still what every
                # existing reader (rebuild-as-presented, avails_lookup, the
                # deck payload) consumes -- this is purely so a rebuilt-into
                # form has groups to show; nothing about an old proposal's
                # stored, rendered or rebuilt output depends on this key.
                "targeting_groups": st.session_state.get("targeting_groups") or [],
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
                "logo_used": bool(uploaded_logo or restored_logo_path),
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
