"""
app.py -- Streamlit form for the Proposal Builder (build spec
section 5), wired directly to assembly.py's selection -> deletion -> fill
pipeline.

Persistent data (master deck versions, products/rates, the audience catalog,
proposal history) lives in Supabase, reached only through db.py, which falls
back to the local file / hardcoded copies below whenever it's unreachable.
"""

import contextlib
import io
import logging
import sys
import os
import json
import re
import subprocess
import tempfile
import uuid
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
import notes_file_import
import slide_map
import targeting_groups as tg
import targeting_map
import wideorbit
from audience_catalog import (CATEGORY_DESCRIPTIONS, all_categories, catalog_warning,
                              category_matches, clear_catalog_cache, load_audience_catalog,
                              validate_segments)

st.set_page_config(page_title="Proposal Builder", layout="wide")


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

# An assistant-turn prefill ("{" to force the response to continue directly
# into JSON) was tried here and abandoned: ANTHROPIC_MODEL rejects it
# outright with a 400 ("This model does not support assistant message
# prefill. The conversation must end with a user message."), confirmed
# live. Don't re-attempt it without checking that constraint against
# whatever model is current at the time -- the defense against a preamble
# instead of JSON is `_extract_largest_json_object` below, not this.

# Where a record of every Claude call goes. The draft path failed live with
# "Claude's response wasn't valid JSON even after stripping markdown fences:
# Expecting value: line 1 column 1 (char 0)" -- char 0 means the text was
# EMPTY, so the fence-stripping the message blamed was never the problem and
# the message pointed at the wrong thing. Nothing was recorded, so a failure
# that cost a live API call and a long wait told us nothing at all. Every
# call now leaves a line behind whether it worked or not.
CLAUDE_LOG_PATH = Path(tempfile.gettempdir()) / "proposal_builder_claude_calls.log"

_LOG = logging.getLogger(__name__)

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
    target market. That row is marked `_placeholder` -- it has a real Geo
    (the originating market) but no audience and no avails, which used to be
    enough for `seed_rows_to_groups` to promote it into a genuine,
    permanent targeting group the first time D2 rendered, before a rep had
    entered anything at all. Once created it never went away: an avails PDF
    import appends its real groups to `existing` rather than replacing it,
    so the blank originating-market row sat in the D2 table forever,
    alongside whatever was actually imported. `_placeholder` is never set on
    a per-market row (target_labels non-empty) -- "a market picked, no
    audience typed yet" is a real, intentional in-progress state elsewhere
    in this file (see plan_lines_from_groups), not a placeholder to discard.
    """
    if not target_labels:
        return [{"Audience": "", "Geo": geo_default, AVAILS_COLUMN_MONTHLY: 0,
                 "_placeholder": True}]
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

    FLOW_REWORK_PLAN.md Phase 3: when an entity spans more than one of the
    given groups, its members sort ADJACENT -- at the position its first-
    seen member's own audience would occupy alone -- rather than scattering
    to wherever each member's individual audience happens to sort.
    Wilmington's three undergraduate audiences, grouped into one entity by
    a draft's `group_entities`, would otherwise interleave with every other
    audience on the document; Annapolis doesn't even need this (an entity's
    5mi/10mi pair already shares one audience, so pure audience-major
    already clusters them) but the general case does. A group whose entity
    has no OTHER member among `groups` sorts exactly as before -- this is a
    no-op for every proposal with no entity grouping in play, including
    everything before this phase.
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
        triples.append((audience, geo, group_id, group))
    if not triples:
        return [(fallback_audience, fallback_geo, None)]

    audience_order, geo_order = {}, {}
    for audience, geo, _group_id, _group in triples:
        audience_order.setdefault(audience, len(audience_order))
        geo_order.setdefault(geo, len(geo_order))

    cluster_order = {}
    for eid, members in tg.entities_of([t[3] for t in triples]).items():
        if len(members) > 1:
            cluster_order[eid] = audience_order[tg.audience_label(members[0])]

    def sort_key(t):
        audience, geo, _group_id, group = t
        primary = cluster_order.get(tg.entity_id_of(group), audience_order[audience])
        return (primary, geo_order[geo])

    ordered = sorted(triples, key=sort_key)
    return [(a, g, gid) for a, g, gid, _group in ordered]


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
        # A directly-picked market IS its own resolution -- geo_def/resolved_markets
        # need no resolver call. resolved_zips DOES, though: the map draws from
        # resolved_zips alone (targeting_map.groups_with_zips), so a market-mode
        # group that skipped this used to be silently invisible on the map --
        # ~96% of a real document's impressions, on the Hershey & Harrisburg
        # scenario that found this. One definition of "resolved" for every mode.
        # (The zip EXPORT is a separate question -- see render_zip_map_builder_page's
        # own geo_def["kind"] == "zips" filter, which deliberately does NOT show
        # these: ad ops targets a DMA by name, not by a zip list nobody consumes.)
        union_zips = set()
        notes, unresolved = [], []
        for key in keys:
            result = geo_resolver.market_to_zips(key)
            union_zips |= set(result.resolved)
            notes.extend(result.notes)
            unresolved.extend(result.unresolved)
        return {"kind": "markets", "markets": keys}, sorted(union_zips), keys, notes, unresolved

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


def _geo_panel_body(group):
    """The controls for ONE targeting group's geo-definition -- mode picker,
    Resolve, results -- with NO expander of its own. Split out of
    `render_group_geo_expander` so several groups sharing one audience can
    be stacked inside ONE outer expander (Streamlit doesn't allow nesting
    an expander inside another) -- see the D2 call site, which groups
    geography expanders by audience rather than rendering one per group
    (a real 12-row Hershey import used to render 12 near-identical "📍
    Geography: ..." headers differing only in which market followed the
    audience name; live feedback, 2026-08-23).

    Resolves through `resolve_group_geography`, and writes the result
    straight onto THIS group in `targeting_groups` -- reassigned, never
    mutated in place, the same discipline `_add_segment_to_group` and
    `merge_plan_rows` use. Never touches any other group.

    Separate from the D2 grid's own Markets cell (Phase 4), which stays the
    quick, direct way to pick markets by hand -- this is for when a rep has
    geography in counties, zips or a radius instead, and needs it resolved
    to markets rather than typed as one.
    """
    gid = group["id"]
    current_markets = group.get("resolved_markets") or []
    summary = tg.geo_label(group, label_for=_market_display_name)
    # This is now a mini-section header WITHIN a shared audience expander
    # (several of these can be stacked in one), not the expander's own
    # title -- bold rather than a caption, so it still reads as a break
    # between groups when there's more than one.
    if summary:
        st.markdown(f"**{summary}**"
                   + (f" -- {len(current_markets)} market(s) resolved"
                      if current_markets else " -- not yet resolved to markets"))
    else:
        st.caption("Not yet resolved to markets")
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
        # FLOW_REWORK_PLAN.md Phase 3: renamed from "Label (optional)" so it
        # can't be confused with the D2 grid's separate entity Label column
        # (the real-world thing a row is FOR) -- this field is, and always
        # was, purely the Geo cell's own override.
        "Geo label (optional)", value=group.get("name") or "", key=f"geo_name_{gid}_{name_gen}",
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


def render_group_geo_expander(group):
    """One targeting group's geo panel in its OWN expander -- for a group
    with no audience yet (there's nothing to group it under). A group that
    DOES have an audience is instead rendered by the D2 call site's own
    per-audience expander, stacking `_geo_panel_body` calls directly --
    Streamlit doesn't allow nesting an expander inside another one, so this
    wrapper and that call site are mutually exclusive paths to the same
    body, never both for the same group.
    """
    label = tg.audience_label(group) or "(untitled)"
    summary = tg.geo_label(group, label_for=_market_display_name)
    header_text = f"{label} -- {summary}" if summary else label
    with st.expander(f"📍 Geography: {header_text}", expanded=False):
        _geo_panel_body(group)


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


# ---------------------------------------------------------------------------
# FALLBACK ONLY -- not the live setting.
#
# The real multiplier and citation live in Supabase's `app_settings` table
# (key='coviewing') and are loaded by load_coviewing_settings() below; edit
# them THERE, not here. This copy exists purely so the app still runs (with a
# visible warning) when Supabase is unreachable.
#
# Sourced from TVision's "State of Streaming 2025" report (Jan.-Dec. 2024
# viewing data) and Nielsen 2025: TVision's P2+ measurement of major
# streaming apps shows viewers-per-viewing-household generally ranging from
# about 1.3-1.7; Nielsen reports 47% of U.S. TV viewing happens with more
# than one person watching. 1.4x sits inside that TVision range rather than
# at or above it -- see BACKLOG.md's Co-viewing item for why that matters on
# a document a client signs.
# ---------------------------------------------------------------------------
FALLBACK_COVIEWING = {
    "multiplier": 1.4,
    "source": "TVision, State of Streaming 2025; Nielsen, 2025",
    "study_date": "2025 (Jan.-Dec. 2024 viewing data)",
    "footnote": "Person-level exposure estimates apply a 1.4x CTV co-viewing "
                "factor (TVision State of Streaming 2025; Nielsen 2025). "
                "Actual co-viewing varies by app, content and household.",
}


@st.cache_data(ttl=600, show_spinner=False)
def load_coviewing_settings():
    """The live co-viewing coefficient: (settings dict, warning).

    Same shape as load_rate_card() -- a TTL cache so an edit in Supabase
    reaches a long-running session without a restart, and a fallback that
    degrades the number rather than breaking the form when Supabase is
    unreachable or the row doesn't exist yet (e.g. before `setup_supabase.py
    settings` has ever been run against a project).
    """
    value, warning = db.fetch_setting("coviewing")
    if value is None:
        return FALLBACK_COVIEWING, f"{warning}. Using the built-in co-viewing factor."
    return value, None


COVIEWING_SETTINGS, COVIEWING_WARNING = load_coviewing_settings()
# Console only, never a rep-facing caption -- the fallback already resolves
# this cleanly (a built-in multiplier that's the same number the table
# would hold once seeded), so there is nothing here for a rep to act on,
# and the old caption showed a raw Supabase payload
# ("Could not find the table 'public.app_settings' in the schema cache...")
# on every render of the co-viewing toggle. Logged once, at import, not
# per-render -- this is a fixed fact about the deployment, not something
# that changes across reruns.
if COVIEWING_WARNING:
    print(f"[coviewing] {COVIEWING_WARNING}")

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


def default_breakout_for_basis(avails_basis):
    """FLOW_REWORK_PLAN.md Phase 2: a brand-new plan option's Breakout
    defaults to the setup band's own Plan basis rather than a hardcoded
    Monthly -- the waterfall's core promise is that a decision made once at
    the top is inherited below, not re-asked. Only a *new* option reads
    this; copying an existing option keeps that option's own breakout."""
    return BREAKOUT_FULL_FLIGHT if avails_basis == AVAILS_BASIS_FLIGHT else BREAKOUT_MONTHLY

# The D2 avails table's Color column reads/writes one of these labels
# rather than a raw hex string. Streamlit has no ColorColumn (a real color
# picker embedded in a data_editor cell isn't a thing this version of
# Streamlit -- or, as far as a live search turned up while building this,
# any released version -- offers; it's a long-open feature request, not
# something declined here). A SelectboxColumn constrained to
# `tg.GROUP_COLORS` is the closest real substitute, and it fits this app's
# own design better than a free RGB picker would anyway: colors were
# always "a fixed, ordered, colorblind-legible-ish palette," never
# arbitrary. The emoji is a rough visual cue, not an exact color match --
# Unicode has no dedicated teal/mauve/pink/grey CIRCLE glyph, so those four
# borrow the closest-reading alternative available. The actual hex drawn
# on the map and in the legend always comes from `tg.GROUP_COLORS`
# straight, never from the emoji.
_COLOR_SWATCH_LABELS = {
    "#4C78A8": "\U0001F535 Blue",
    "#FF7F0E": "\U0001F7E0 Orange",
    "#54A24B": "\U0001F7E2 Green",
    "#FF9DA6": "\U0001F338 Pink",
    "#EECA3B": "\U0001F7E1 Yellow",
    "#D62728": "\U0001F534 Red",
    "#72B7B2": "\U0001F537 Teal",
    "#9D755D": "\U0001F7E4 Brown",
    "#B279A2": "\U0001F7E3 Mauve",
    "#BAB0AC": "\U000026AA Grey",
}
_COLOR_LABEL_TO_HEX = {label: hexcode for hexcode, label in _COLOR_SWATCH_LABELS.items()}
# Superseded hexes (the 2026-08-23 orange/red swap, see
# targeting_groups.GROUP_COLORS), kept OUT of _COLOR_SWATCH_LABELS itself --
# that dict's values() populate the D2 grid's Color dropdown
# (SelectboxColumn options), and a superseded hex sharing a label with its
# replacement would show "Orange" twice in that list. This is consulted only
# as a fallback, so a proposal saved before the swap still displays a real
# name for a color it was actually given, rather than falling back to Blue.
_LEGACY_COLOR_SWATCH_LABELS = {
    "#F58518": "\U0001F7E0 Orange",
    "#E45756": "\U0001F534 Red",
}


def _color_swatch_label(hexcode):
    hexcode = str(hexcode or "").upper()
    if hexcode in _COLOR_SWATCH_LABELS:
        return _COLOR_SWATCH_LABELS[hexcode]
    if hexcode in _LEGACY_COLOR_SWATCH_LABELS:
        return _LEGACY_COLOR_SWATCH_LABELS[hexcode]
    return next(iter(_COLOR_SWATCH_LABELS.values()))


def _color_from_swatch_label(label):
    return _COLOR_LABEL_TO_HEX.get(str(label or ""), tg.GROUP_COLORS[0])


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

CUSTOM_FEE_PRODUCT = "custom_fee"

DRAFT_JSON_SCHEMA_EXAMPLE = """{
  "client_name": "", "vertical": "",
  "target_markets": [],
  "spanish_campaign": false,
  "geo": "",
  "total_budget": 0,
  "total_budget_basis": "net|gross",
  "breakout": "monthly|full_flight",
  "media_plan_lines": [
    {"product": "premion_streaming_tv", "label": "Commercial", "audience_track": "SMB owners and business executives", "allocation": {"percent_of_remainder": 60}},
    {"product": "premion_streaming_tv", "label": "Retail", "audience_track": "consumers in-market for deposit accounts", "allocation": {"percent_of_remainder": 40}, "cpm": 28},
    {"product": "streaming_retargeting_display", "allocation": {"percent_of_total": 10}},
    {"product": "sport:nfl_playoffs", "allocation": {"flat_amount": 40000}},
    {"product": "custom_fee", "label": "Dynamic Ad Creation", "allocation": {"flat_amount": 850}}
  ],
  "group_selection": null,
  "group_allocation": null,
  "group_cpm": null,
  "group_selection_reason": "",
  "group_entities": null,
  "options": null,
  "total_tv": false,
  "show_sov": false,
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


# Every session_state key bound to a widget that renders inside the setup
# band (from st.header("Setup") down through the Draft-from-notes block),
# in other words above the point in the SAME run where apply_draft_to_form's
# own write happens. A widget in this set has always already instantiated
# by the time the Draft button's handler runs -- writing to its key directly
# raises "cannot be modified after the widget ... is instantiated" (found
# live, twice now: client_name, then total_tv, both fixed as one-off patches
# before this set existed). apply_draft_to_form's own final write loop
# checks every key it's about to write against this set and defers a match
# into `_draft_pending_fields` instead (apply_pending_draft_fields, applied
# at the very top of the NEXT run, before this band's first widget) -- so a
# future band addition, or a future field this function starts writing,
# can't reintroduce the same crash unnoticed. Section A and everything below
# it render AFTER the band and after the gate, so nothing there needs to be
# in this set -- confirmed for target_dma_choice specifically (see
# apply_draft_to_form's own comment on that write) and true by construction
# for the rest: the Draft button's own `st.rerun()` cuts a run off before
# anything past the band is ever reached.
BAND_WIDGET_KEYS = frozenset({
    "avails_mode", "client_name", "market_choice", "avails_basis", "total_tv",
    "flight_start", "flight_end", "custom_flighting", "draft_from_notes",
    "notes_upload", "draft_notes_input", "draft_clarifications",
})

# Which form section each session_state key the draft writes belongs to --
# keyed to the *widget's* own section (the one its on_change clears), since
# that's what tells us the user has since hand-edited it. Used by the
# clarify-and-re-draft round to leave hand-edited sections alone.
DRAFT_KEY_SECTIONS = {
    # FLOW_REWORK_PLAN.md Phase 5: market_choice, flight_start/flight_end/
    # active_months and avails_basis (Plan basis) are band INPUTS now, never
    # drafted outputs -- apply_draft_to_form no longer writes any of them, so
    # they carry no entry here. Phase 6: client_name and target_dma_choice
    # joined them -- both are seed-if-empty only now (queued for client_name,
    # direct for target_dma_choice; see apply_draft_to_form), never touched
    # once real, so neither ever lands in `updates` and neither needs a
    # skip_sections entry here either. vertical_choice is still drafted,
    # unchanged.
    "vertical_choice": "basics",
    "spanish_campaign": "attribution",
    "goals_text": "specs", "audience_text": "specs", "geography_text": "specs",
    "budget_text": "specs", "placements_text": "specs", "timing_text": "specs",
    "avails_seed_rows": "avails", "avails_version": "avails",
    # targeting_groups is the projection's source once anything has written
    # it, so it has to be skipped alongside avails_seed_rows or a preserved
    # "avails" section would keep the old rows while a re-draft silently
    # overwrote the groups they're supposed to agree with.
    "targeting_groups": "avails",
    "live_sports_enabled": "products", "selected_sports": "products",
    "include_sport_viewership": "products",
    "plan_options": "media_plan",
    "plan_options_gen": "media_plan", "show_sov": "media_plan",
    # A gross-labelled budget's auto-tick rides with the rows it priced --
    # a re-draft that's skipping "media_plan" because the rep already
    # hand-edited the plan (which includes hand-toggling this box) must
    # leave their choice alone too, not silently re-tick it.
    "agency_gross_up": "media_plan",
    # Internal bookkeeping that only means anything alongside the rows it
    # describes -- it has to be skipped with them or it would claim rows that
    # were never written.
    "_product_seed_key": "media_plan", "_shared_fields_key": "media_plan",
    # The budget/allocation/selection a draft contributed when targeting
    # groups already own the plan -- rides with the rows it priced, same
    # reasoning as _product_seed_key/_shared_fields_key above.
    "draft_plan_intent": "media_plan",
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
    "wo_upload", "wo_clear", "dup_btn_", "cs_upload", "vault_upload",
    "deck_upload", "logo_upload", "usage_upload", "avails_pdf_upload_",
    # The intake area's notes-file uploader (UX sweep, BACKLOG.md).
    "notes_upload",
    # The D2 avails table's "Apply a color to a whole audience" buttons, one
    # per (audience, detached color) pair actually in play -- not one per
    # group; a real 12-row Hershey import rendered 12 near-identical buttons
    # before this was scoped down (live feedback, 2026-08-23). Caught by
    # this test on the first real run after being added -- exactly the gap
    # this guard exists for.
    "apply_color_",
    # FLOW_REWORK_PLAN.md Phase 3: the D2 entity-grouping expander's
    # "Group selected rows" button and its per-entity "Ungroup" buttons --
    # same hazard, same fix as "Apply to all" right above.
    "entity_group_apply", "ungroup_",
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
    # The D2 avails divergence panel's "Adjust to plan dates" / "Use
    # document figure" buttons, one pair per group with a real divergence
    # (FLOW_REWORK_PLAN.md Phase 2) -- both keyed per group id.
    "avails_adjust_", "avails_use_doc_",
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
    # The "Report an issue" popover's Submit button (rendered on every
    # page, including Build). Its category selectbox and notes text_area
    # are ordinary settable widgets, but since the popover is unconditional
    # they're never garbage-collected in the first place -- there's nothing
    # for the snapshot to carry for them either way.
    "feedback_submit_",
    # The D2 avails table's "Add all to plan" / "Clear plan lines" buttons,
    # and the confirmation panel's Remove/Keep pair (the uncheck-with-edits
    # guard) -- all four render on the Build page like everything else here.
    "avails_include_all_", "avails_include_clear_",
    "avails_include_confirm_remove_", "avails_include_confirm_keep_",
    # The setup band's "Match avails flighting" button (FLOW_REWORK_PLAN.md
    # Phase 4a) -- one key, band-level, not per-group like the D2 pair above.
    "match_avails_flighting",
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

    st.title("Proposal Builder")
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

    st.title("Proposal Builder")
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


# ---------------------------------------------------------------------------
# In-app feedback -- a persistent "Report an issue" popover on every page
# (BACKLOG.md). Category, free-text notes, and a best-effort state
# snapshot -- never a screenshot, which is a deferred fast-follow, not
# part of this pass.
# ---------------------------------------------------------------------------
FEEDBACK_CATEGORIES = ["Avails", "Proposal", "Map", "Audiences", "Drafting", "Deck output", "Other"]


def capture_feedback_state(page):
    """Best-effort reproduction context for one feedback report.

    Not a full Generate-time payload -- that's assembled from many local
    variables deep inside the Generate handler, not reconstructable from a
    sidebar popover reachable on every page. This captures the Build page's
    own raw, reassign-not-mutate source-of-truth state instead
    (`targeting_groups`, `plan_options`, the scalar Section A/flight
    fields) -- deliberately never a `data_editor`'s own widget state (this
    file's own documented trap: a delta is an edit, not data) and never the
    derived, markup-applied numbers Generate computes, which are
    recomputable from these same inputs. Whatever's absent (a rep reporting
    from a page other than Build never populated these) is just missing
    from the dict, not faked.

    `last_claude_failure` (see `log_claude_call`) rides along the same
    way -- present only when the most recent Claude call actually failed,
    absent otherwise, never a stale failure from earlier in the session.
    It's truncated model output, not proposal content, so it stays out of
    every other export path (the deck, the proposal record) -- this popover
    is the one place it's meant to surface.
    """
    active, _ = db.active_deck_version()
    flight_start = st.session_state.get("flight_start")
    flight_end = st.session_state.get("flight_end")
    return {
        "page": page,
        "build_stamp": BUILD_STAMP,
        "user": current_user(),
        "active_deck_version": (active or {}).get("id"),
        "client_name": st.session_state.get("client_name"),
        "vertical_choice": st.session_state.get("vertical_choice"),
        "market_choice": st.session_state.get("market_choice"),
        "agency_gross_up": st.session_state.get("agency_gross_up"),
        "proposal_title": st.session_state.get("proposal_title"),
        "flight_start": str(flight_start) if flight_start else None,
        "flight_end": str(flight_end) if flight_end else None,
        "targeting_groups": st.session_state.get("targeting_groups"),
        "plan_options": st.session_state.get("plan_options"),
        "draft_source_notes": st.session_state.get("draft_source_notes"),
        "draft_unresolved": st.session_state.get("draft_unresolved"),
        "draft_unresolved_internal": st.session_state.get("draft_unresolved_internal"),
        "last_claude_failure": st.session_state.get("last_claude_failure"),
    }


def render_feedback_popover():
    """The persistent "Report an issue" control, on every page's sidebar.

    A popover, not a page of its own -- reporting a bug shouldn't cost the
    rep their place in the form. `feedback_form_gen` moves the form's
    widgets to fresh keys after a successful submit (the same device as
    `option_name_{gen}_{idx}` elsewhere in this file) rather than writing
    into an already-instantiated widget's own session_state key, which
    Streamlit refuses outright.
    """
    gen = st.session_state.get("feedback_form_gen", 0)
    with st.sidebar.popover("🚩 Report an issue", use_container_width=True):
        st.caption("Something wrong, confusing, or missing? This attaches your current "
                   "page and campaign state (never a screenshot) so it's reproducible.")
        category = st.selectbox("Category", FEEDBACK_CATEGORIES, key=f"feedback_category_{gen}")
        notes = st.text_area("What happened?", key=f"feedback_notes_{gen}", height=100)
        if st.button("Submit report", key=f"feedback_submit_{gen}", disabled=not notes.strip()):
            page = st.session_state.get("page_choice") or "Build a proposal"
            state = capture_feedback_state(page)
            row_id, error = db.submit_feedback(
                category=category, notes=notes.strip(), page=page, state=state,
                created_by=current_user())
            if error:
                st.error(f"Couldn't submit ({error}) — try again, or flag it directly to Matt.")
            else:
                st.session_state["feedback_form_gen"] = gen + 1
                st.session_state["feedback_just_submitted"] = True
                st.rerun()
    if st.session_state.pop("feedback_just_submitted", False):
        st.sidebar.success("Report submitted — thanks.")


def _feedback_export_markdown(rows):
    """One markdown bug report per row -- description, captured state, and
    a reproduction pointer, the format meant to be pasted straight into
    Claude Code (BACKLOG.md)."""
    sections = []
    for row in rows:
        created = (row.get("created_at") or "")[:16].replace("T", " ")
        state = row.get("state_json") or {}
        lines = [
            f"## [{row.get('category') or 'Other'}] {created}",
            "",
            f"**Reported by:** {row.get('created_by') or 'unknown'}  ",
            f"**Page:** {row.get('page') or '—'}  ",
            f"**Build:** {state.get('build_stamp') or '—'}",
            "",
            "**Description:**",
            "",
            row.get("notes") or "",
            "",
            "**Captured state:**",
            "",
            "```json",
            json.dumps(state, indent=2, default=str),
            "```",
        ]
        sections.append("\n".join(lines))
    return "\n\n---\n\n".join(sections) + "\n"


def render_feedback_admin_page():
    """Feedback reports -- newest first, filterable by status/category,
    exportable as a markdown bug report. No access restriction beyond the
    shared password, same as every other admin page here -- this app has
    no user accounts to gate on.
    """
    st.title("Feedback reports")
    open_count, warning = db.count_open_feedback()
    if warning:
        st.warning(f"⚠️ {warning}")
    elif open_count:
        st.caption(f"{open_count} open report{'s' if open_count != 1 else ''}.")
    else:
        st.caption("No open reports.")

    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        status_filter = st.selectbox("Status", ["open", "closed", "All"], key="feedback_status_filter")
    with filter_col2:
        category_filter = st.selectbox("Category", ["All"] + FEEDBACK_CATEGORIES,
                                       key="feedback_category_filter")

    rows, warning = db.fetch_feedback(
        status=None if status_filter == "All" else status_filter,
        category=None if category_filter == "All" else category_filter)
    if warning:
        st.warning(f"⚠️ {warning}")
        return
    if not rows:
        st.caption("No feedback reports match this filter.")
        return

    selected_ids = []
    for row in rows:
        created = (row.get("created_at") or "")[:16].replace("T", " ")
        summary = row.get("notes") or ""
        summary = summary if len(summary) <= 80 else summary[:80] + "…"
        with st.expander(f"{'🟢' if row.get('status') == 'open' else '⚪'} "
                         f"[{row.get('category') or 'Other'}] {summary} — {created}"):
            st.write(row.get("notes") or "")
            state = row.get("state_json") or {}
            st.caption(f"Page: {row.get('page') or '—'} · Reported by: "
                       f"{row.get('created_by') or 'unknown'} · Build: "
                       f"{state.get('build_stamp') or '—'}")
            with st.expander("Captured state", expanded=False):
                st.json(state)
            pick_col, status_col = st.columns([1, 2])
            with pick_col:
                if st.checkbox("Include in export", key=f"fb_pick_{row['id']}"):
                    selected_ids.append(row["id"])
            with status_col:
                new_status = st.selectbox(
                    "Status", ["open", "closed"],
                    index=0 if row.get("status") == "open" else 1,
                    key=f"fb_status_{row['id']}", label_visibility="collapsed")
                if new_status != row.get("status"):
                    ok, error = db.update_feedback_status(row["id"], new_status)
                    if error:
                        st.error(error)
                    else:
                        st.rerun()

    if selected_ids:
        selected_rows = [r for r in rows if r["id"] in selected_ids]
        st.download_button("Export selected as markdown",
                           data=_feedback_export_markdown(selected_rows),
                           file_name="feedback_reports.md", mime="text/markdown")


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


def build_draft_prompt(notes, existing_groups=None, *, market_choice=None,
                       flight_start=None, flight_end=None, avails_basis=None):
    """`market_choice`/`flight_start`/`flight_end`/`avails_basis` are the
    setup band's own current values (FLOW_REWORK_PLAN.md Phase 5) -- given
    to the model as context it must not revise, never asked for as schema
    fields (DRAFT_JSON_SCHEMA_EXAMPLE carries no "market"/"flight_start"/
    "flight_end" any more). All four default to None so a caller with no
    band context yet (a test calling this directly, or -- in principle -- a
    rep who somehow reaches the button before the band gate would apply)
    gets a prompt with no frame section at all, rather than one claiming a
    market or flight that isn't real."""
    vertical_hint = _detect_vertical_hint(notes)
    catalog_slice = build_catalog_slice(vertical_hint)
    products_info = {k: {"label": v["label"], "default_cpm": v["default_cpm"]} for k, v in PRODUCTS.items()}
    today = date.today().isoformat()
    attribution_help = "\n".join(f'  - "{key}": {text}'
                                 for key, text in ATTRIBUTION_DESCRIPTIONS.items())

    # Only when real groups already exist on this proposal (an avails PDF
    # already imported, or a rep already built some by hand) does the
    # prompt gain this section at all -- a no-groups draft's prompt (and
    # every frozen tier-1 fixture) stays byte-identical to before. Real
    # groups only: a `_placeholder` row (D2's blank starter) or a
    # term-less market-only row has nothing for the model to select.
    real_groups = [g for g in (existing_groups or []) if g.get("id") and g.get("terms")]
    group_section = ""
    if real_groups:
        group_catalog = json.dumps([
            {"id": g["id"], "audience": tg.audience_label(g),
             "geo": tg.geo_label(g, label_for=_market_display_name),
             "monthly_avails": g.get("avails_monthly", 0),
             # FLOW_REWORK_PLAN.md Phase 3: which real-world thing (entity)
             # this row already belongs to, and its label if it has one --
             # so the model can see what's already grouped rather than
             # treating every row as its own thing. Two rows sharing an
             # "entity" value here are already ONE real-world thing; a
             # stated split divides by how many DISTINCT entity values are
             # selected, never by row count.
             "entity": tg.entity_id_of(g), "label": tg.entity_label_of(g)}
            for g in real_groups
        ])
        group_section = f"""

The avails table for this buy already lists {len(real_groups)} audience/geography combination(s), with the id the app uses for each (JSON): {group_catalog}

Sell from THIS LIST, not from a new plan you invent -- these rows are the real, resolved audiences and geographies already on the campaign, and every Premion Streaming TV dollar in this buy is one of them, priced by an allocation rather than written as its own media_plan_lines entry:
- "group_selection" names which of those rows this plan sells. Use {{"mode": "named", "ids": [...]}} with the exact "id" values when the notes single specific rows out by name; use {{"mode": "named", "match": ["Subaru", "10 mile"]}} with the notes' own words when you can tell which rows they mean but not their ids; use {{"mode": "all"}} when the notes describe the whole buy without singling any row out. "match" lets the app re-apply the same choice to rows added later, so use the audience or geography language the notes themselves used.
- "group_allocation" is ONE allocation, applied per ENTITY, not per row -- two selected rows sharing one "entity" value count as ONE share, not two, so "split evenly across the four stores" divides by how many distinct entities are selected, however many rows sit behind them. Five shapes, and the model states ONLY the split -- Python always computes every dollar and impression from it, never the model:
  - {{"split_evenly": true}} -- a stated total divided evenly across the selected entities.
  - {{"percent_of_total": N}} / {{"percent_of_remainder": N}} -- a stated share of the budget (or of what's left after other shares).
  - {{"flat_amount": N}} -- a stated per-entity dollar rate, TRANSCRIBED from the notes, never a figure you worked out yourself.
  - {{"percent_of_avails": N}} -- a stated percentage of the row's own avails ("cover 20% of the available inventory") -- no reference needed, the selection already names the row; Python multiplies it out against the real avails figure.
  Precedence when the notes could support more than one reading for the same entity: a STATED DOLLAR FIGURE always wins over an inferred avails percentage -- the dollar figure is what the client actually agreed to.
- "group_cpm" is the rate, in dollars, for EVERY selected row -- set it whenever the notes state a rate for this streaming buy ("$29 CPM", "they're getting the streaming at 30"), the same rule as a media plan line's own "cpm" below. Omit it and the rate card default applies. This is the ONLY way a negotiated rate reaches a group-selection sale -- a "cpm" on a "premion_streaming_tv" media_plan_lines entry is ignored here, since that product is not written as its own line once groups own the plan (see below).
- "group_selection_reason" is one plain sentence a salesperson reads: which rows you put on the plan and what in the notes told you that.
- "media_plan_lines" then covers only the products that are NOT part of that streaming selection -- retargeting, Audience Marketplace, sports packages, one-time fees. The Premion Streaming TV line for each selected row comes from "group_allocation"/"group_cpm" instead, so do not also write a "premion_streaming_tv" entry in "media_plan_lines" for it.

"group_entities" (top-level, ONE list for the whole proposal, never per-option -- these entries describe the campaign's own avails table, not any one plan) does TWO different jobs, both strictly from what the notes actually state -- never invent or guess a name for a row the notes don't name:
- **Naming several rows that are really the SAME real-world thing** -- {{"label": "Toyota of Annapolis", "match": ["Subaru"]}} or {{"label": "Undergraduate", "ids": [...]}}, base each entry on the notes' own words tying specific rows together ("split the $15K evenly across the four stores" naming four rows as one store each is exactly this).
- **Naming ONE row that is already its own, correctly separate thing** -- just as real a use of this field, and easy to miss: several avails rows sharing one audience but differing only by geography (LiveWell Animal Hospital: seven locations, all "LIFESTYLE Pets", one per address) are NOT the same entity and must not be merged -- but each one still deserves its own real name when the notes give it one, e.g. {{"label": "Alexandria", "match": ["277 S Washington"]}} for the Alexandria row alone. Without this, a client reads the exact same "LIFESTYLE Pets" Targeting text on every row and has no way to tell which line is which location. Use the notes' own name for the place ("our Alexandria location", "the Falls Church store") -- match it to the row by whatever the avails table's own audience/geo text shares with it (a street name, a city, a landmark). **If the notes don't give you a name for a specific row -- including when several rows would extract the exact same name from the avails table alone (e.g. three different DC-area radii all just called "Washington DC") -- leave that row out of "group_entities" entirely.** A blank label invites a rep to type the real name in; a wrong or arbitrarily-disambiguated one (making up "DC-North"/"DC-South" to tell two identical-looking rows apart) gets sent to a client as if it were real.
Leaving a row out of every "group_entities" entry is always safe and is the default -- for both jobs above.

Each entry in "options" carries its own "group_selection"/"group_allocation"/"group_cpm"/"group_selection_reason" too, the same way it carries its own "total_budget"."""

    # FLOW_REWORK_PLAN.md Phase 5: the setup band is filled in BEFORE
    # drafting runs now, not inferred from the notes afterward -- state its
    # values as given facts rather than asking the model to extract them.
    # This is the actual accuracy gain the phase exists for; dropping the
    # schema fields alone would just make the model silently drop
    # information it used to at least attempt. Omitted entirely (not a
    # placeholder sentence) when the band isn't filled in yet, since a
    # rep can reach the notes box and this prompt before typing dates.
    frame_section = ""
    if market_choice in ("DC", "Harrisburg") and isinstance(flight_start, date) and isinstance(flight_end, date):
        frame_months = len(month_list(flight_start, flight_end))
        frame_section = f"""

This proposal's frame is already set and is not yours to fill in: originating market {market_choice}, flight {flight_start.isoformat()} to {flight_end.isoformat()} ({frame_months} month{'s' if frame_months != 1 else ''}), plan basis {avails_basis or AVAILS_BASIS_MONTHLY}. Do not return a market or flight date of any kind -- there is no field for them. If the notes state a date or market that disagrees with this frame, do not resolve the disagreement yourself: name it in "unresolved_internal" quoting both what the notes said and what the frame already has, and leave the frame as the frame. Read the notes only for what's still open: budget, audiences, products, allocation."""

    return f"""You are drafting a first pass at a Premion CTV/OTT advertising proposal from raw meeting/discovery notes. Today's date is {today}.{frame_section} Return ONLY valid JSON matching the schema below -- no markdown code fences, no preamble, no explanation, just the JSON object.

Schema:
{DRAFT_JSON_SCHEMA_EXAMPLE}

"media_plan_lines" is the full media plan, expressed one entry per intended row -- not one entry per product. If the notes call for the same product run as separate lines (e.g. two different audience tracks, or a commercial vs. retail split), give each its own entry with its own "label" and "audience_track"; each becomes its own media plan row, with "label" appended to the product's own tactic name (e.g. "Premion Streaming TV — Commercial") and "audience_track" as that row's Targeting. A line with no "label" just uses the product's own name as-is.{group_section}

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
  "options": [{{"name": "Good", "total_budget": 50000, "breakout": "monthly", "media_plan_lines": [...]}}, {{"name": "Better", "total_budget": 75000, "total_budget_basis": "gross", "breakout": "monthly", "media_plan_lines": [...]}}]
Each option is a complete plan in its own right, with its own budget and its own full set of lines, and becomes its own media plan slide in the deck. "name" is what the client sees appended to the plan title, so use the notes' own words for the tier ("Good"/"Better"/"Best", "$50K Plan") rather than a generic letter. When "options" is set, leave "media_plan_lines" empty. Do NOT invent scenarios the notes didn't ask for -- one plan is the normal answer.

EVERY option MUST carry its own "total_budget" -- it is what that scenario costs, and each option's allocations are resolved independently against it. An option's "total_budget" falls back to the top-level one only if omitted, and if neither is set that option prices at $0 and gets dropped entirely, which is never a useful answer. When the notes state a BUDGET RANGE with no instruction on how to split it ("between $50K and $75K", "somewhere in the 50 to 75 range, wants to see both"), the right answer is one option per stated figure, each carrying that figure as its own "total_budget" -- e.g. a "$50K Plan" at 50000 and a "$75K Plan" at 75000. If the notes give a range but you cannot tell what the individual figures should be, put a single plan at the lower figure and say so; never return lines with no budget behind them.

**When the notes explicitly ask YOU to determine or propose the investment level -- as opposed to simply not mentioning a budget at all -- estimate a real number, never $0.** "Build the plan around whatever investment makes sense for the audience," "no budget was set, price it at a level that would be effective," "propose what you think is reasonable" are an instruction to price the plan, not a reason to leave it unpriced -- the notes are handing you a judgment call, not describing an open question. A $0 plan is never a valid answer to this kind of note: every line resolves to $0, the deck ships with a zeroed-out media plan on every row, and there is nothing left for a rep to send. Ground the estimate in the audience's own avails from the avails table above -- a figure landing roughly in the 10-20% range of the audience's monthly avails at the applicable rate-card CPM is a defensible starting point for a single, modest-reach buy; scale it up for a broader footprint (more locations, more markets, more audiences sharing the buy). State your estimate as that option's "total_budget" like any other figure, and name it plainly in "unresolved" as a seller-proposed number that needs confirming before it goes to the client -- the same way you'd flag an assumed 50/50 split or a rate-card default, not as a missing input.

Set "total_tv" to true ONLY when the notes describe a broadcast schedule PREMION ITSELF is running, on one of our own two stations -- WUSA9 in DC, or WPMT/FOX43 in Harrisburg -- alongside the streaming plan: "keep the WUSA schedule going", "broadcast plan attached", "Total TV", "we're also running spots on FOX43", an existing station buy on one of those two being continued or added to. Total TV switches the deck to its co-branded station template and opens the panel where the seller uploads the Wide Orbit schedule, so getting it from the notes saves them a step and, more importantly, means they turn it on BEFORE building the plan rather than after.

**A client's own linear TV, on any station that isn't WUSA9 or WPMT/FOX43, is not Total TV -- it's background, not a Premion broadcast component**, however the notes describe it: "they're continuing their existing broadcast buy in Denver", "client also runs spots on the local Fox affiliate through their own agency", "keeping their linear schedule going alongside this." The presence of the words "broadcast" or "linear" is not the signal -- WHICH STATION, and whether Premion is the one running it, is. Do NOT set it for a streaming-only campaign, for a client's own linear buy on someone else's station or through a different vendor, or for a market other than DC/Harrisburg. Do NOT invent broadcast lines in "media_plan_lines" -- the schedule is imported from a real Wide Orbit export, not drafted. Flag that Total TV was switched on and that the schedule still needs uploading.

INCLUDE ONLY THE PRODUCTS THE NOTES ACTUALLY CALL FOR. There is no mandatory line and no default product -- "premion_streaming_tv" in particular is NOT required and must not be added just to have a baseline CTV line. A sports-only plan, an Audience-Marketplace-only plan, a retargeting-only plan, or a single-line plan are all perfectly valid proposals. If the notes describe an NFL campaign and nothing else, the correct media plan is one NFL line and nothing else. If the notes are genuinely silent about what to buy, say so instead of inventing a product mix.

**But every product the notes DO ask for gets a line, even when no budget is given for it.** "She wants audience targeting on top of the general streaming" names a product; the absence of a split is a missing number, not a reason to leave the product out. Include the line with a "split_evenly" allocation so it shares what's left with the other unspecified lines, and flag the assumption ("the notes didn't say how to split between X and Y -- divided evenly, confirm the intended weighting"). Omitting a requested product is the worse failure of the two: an assumption the reviewer can see gets corrected, whereas a product that never appears in the plan is invisible to them and quietly missing from the deck.

"total_budget" is always what the WHOLE campaign costs across the entire flight, never a monthly rate -- and so is every dollar amount the allocations resolve to. If the notes quote the budget per month ("$20K a month for three months"), multiply it out to the full-flight figure yourself ($60,000) and note the per-month figure if it's worth flagging.

"total_budget_basis" says which side of the agency markup "total_budget" is stated on: "net" (the default -- what the media plan lines are priced against, same convention as "cpm") or "gross" -- ONLY when the notes EXPLICITLY say the budget is what the client is billed, is inclusive of agency commission, or otherwise names it as the grossed figure, e.g. "$15K total gross budget, net CPM is $32," "the client's all-in number is $50K," "that budget includes our commission." Set it exactly like any other transcription -- read the label the notes actually used, write it down, do no arithmetic: Python divides the stated figure by the agency markup to get the net amount each line is priced against, and switches the gross-up checkbox on so the deck lands back on the stated gross figure. **The genuinely ambiguous case -- an agency or commission is mentioned, but nothing says which side of it the budget number sits on -- stays "net."** Guessing "gross" from ambiguous phrasing is exactly the kind of classification this app got burned by once already (see CLAUDE.md's agency gross-up history); only an explicit label earns the conversion.

"breakout" is a separate question -- not what the plan costs, but how it's presented: "monthly" for a plan broken out month by month (the normal case, and the right answer whenever the notes talk in per-month terms at all), or "full_flight" only if the notes specifically want the flight shown as a single combined period. Python divides the campaign total across the flight for a monthly breakout; do not do that arithmetic yourself.

Each line's "allocation" has exactly one key:
- "flat_amount": this line costs exactly this many dollars.
- "percent_of_total": this line costs this percent of total_budget.
- "percent_of_remainder": this line costs this percent of whatever's left after all "flat_amount" and "percent_of_total" lines are subtracted from total_budget (percent_of_remainder entries across lines should sum to 100 if they're meant to exhaust the remainder).
- "split_evenly": this line shares equally, with every other "split_evenly" line, in whatever's left after "flat_amount"/"percent_of_total"/"percent_of_remainder" lines are all accounted for -- use this for "split evenly across N audiences/tracks" instead of trying to pre-compute a percentage yourself.
- "percent_of_avails": this line buys this percent of one audience's available impressions -- a reach TARGET, not a budget. Use it only when the reach percentage itself is the thing being bought -- "reach 20% of the available audience", "one option at 20% and one at 40%", "40% penetration against the home-services segment", "spend whatever it takes to hit 25%". Pair it with "avails_ref", naming the audience it refers to -- use the exact same wording as that "audiences" entry's "segment" so the two can be matched up. State only the percentage and which audience; Python multiplies it out against the real avails figure and the CPM, and works out the cost. A reach line needs no budget: if the notes give both a reach percentage and a budget and the arithmetic disagrees, the reach wins and the difference is flagged for the reviewer.
  **Do NOT use "percent_of_avails" just because the notes mention a dollar figure alongside an avails number and ask what percentage that buys.** "$4,000 against the 1.7M avails on that segment, tell her what percent that reaches" states a BUDGET ($4,000) and asks for the resulting reach to be reported -- it is not asking to spend whatever it takes to hit a percentage. That is a "flat_amount" (or "percent_of_total") line like any other; Python derives the reach percentage that budget buys from the real avails figure and reports it automatically, so no field is needed to make that happen. The question that tells the two apart: which number must come out exactly as stated -- a reach percentage that spend is built to hit, or a dollar figure that reach is only measured against afterward? The second case is far more common than the first.
Do NOT do any arithmetic yourself beyond picking which allocation type fits each line -- Python resolves flat_amount and percent_of_total first, then percent_of_remainder, then splits whatever's left evenly across split_evenly lines, then computes every dollar amount and impression count.

Each line may also carry an optional "cpm", the rate for that line in dollars. Rates are negotiated per deal, so set it whenever the notes state a rate for that line -- "$28 CPM on the Premion line", "they're getting the streaming at 30", "we agreed $45 for the NFL inventory". Write down exactly the rate the notes state, as stated, and nothing more -- even when the notes also say to gross it up ("the net CPM is $32, gross it up"), write the rate exactly as given (32, not 36.80). Grossing up is a rep action taken once for the whole order, after the plan is drafted; it is never computed here, and never per line. Omit "cpm" entirely and the product's rate card default applies, which is what you want whenever the notes say nothing about rate or say to hold to the rate card ("at rate card", "standard rates", "no discount"). Set it ONLY from a rate the notes actually state -- never to hit a budget or impression target, which is what the allocations are for. Write it as a plain JSON number -- 28 -- leaving the currency symbol and the word CPM to the app. It never applies to a "{CUSTOM_FEE_PRODUCT}" line, which has no rate at all. Every override is flagged for the reviewer automatically, so you do not need to mention it yourself.

Rules:
- **A stated dollar budget always drives cost and impressions; a stated avails/reach figure next to it is only a ceiling to report reach against, never a spending target.** When the notes give BOTH a budget and an avails number for the same line, price it as "flat_amount" (or "percent_of_total") against the budget -- never "percent_of_avails". "$4,000 against the 1.7M avails on that segment, what percent does that reach" is a $4,000 budget with a reach question attached, not a request to spend whatever it takes to hit a percentage; Python computes and reports that percentage automatically once the line is priced from the budget. Reserve "percent_of_avails" for the rarer case where the notes state a reach PERCENTAGE as the thing being bought -- see the allocation type below for exactly how to tell the two apart. A dollar figure is a budget even when an avails number sits in the very same sentence.
- "vertical" must be exactly one of: {list(VERTICALS.values())}{f' -- a trade term in the notes suggests "{vertical_hint}" (used only to pick which audience segments are shown above); confirm it actually fits before returning it, or return a different one if the notes point elsewhere' if vertical_hint else ""}
- "target_markets" is where the campaign is AIMED: a list of DMA or city names exactly as the notes give them ("Denver", "Atlanta", "Washington DC", "the Bay Area"). List EVERY market the notes name, in the order they are named -- a brief that says "Denver, Atlanta and Phoenix" produces three entries, not one. The app matches each name against the real Nielsen DMA list and reports anything it can't place, so give the name as written rather than guessing at an official spelling. Leave it empty when the notes name no target market at all; do not fall back to the originating market, which is a different thing (stated in the frame above, not something you return).
- "sports" entries must be exactly one of: {list(SPORTS.values())}. This drives which sports package slides go in the deck -- list every package that also appears as a "{SPORT_PRODUCT_PREFIX}" media plan line, and leave it empty when the notes call for no sports at all.
- "attribution" entries must be drawn from this list, using the exact key shown. Include every one the notes call for -- these drive real slides, real Included-with-Campaign entries and real toggles, so an option the notes ask for and you omit simply never reaches the proposal:
{attribution_help}
- "show_sov" turns on "% of avails" (share of voice) on the media plan -- set it true whenever the notes ask to show SOV, share of voice, or what percent of the available inventory a line reaches ("include the SOV", "show what percent of avails this buys"). It only ever affects display for a line whose targeting already traces back to a real avails figure, so turning it on is always safe to do whenever asked.
- "audiences[].segment" must be an EXACT name from the audience catalog slice below -- do not paraphrase or invent segment names. If nothing in the slice fits, it's fine to omit audiences or note it. {AUDIENCE_MATCH_GUIDANCE} When a media_plan_lines entry's "audience_track" describes the same audience as one of your "audiences" entries, use the same wording for both.
- **At most ONE non-RFP-selectable segment. This is a hard limit, not a preference.** Each catalog entry carries an "rfp_selectable" flag, and a campaign may book only one segment with "rfp_selectable": false (a "custom" segment). A draft containing two or more is invalid and cannot be used until someone removes the extras by hand. So: when two segments would serve the same purpose, take the RFP-selectable one. If several custom segments all look relevant -- which happens when a niche category's best matches are all custom -- **choose the single most important one and name the others as alternatives the reviewer could swap in**, rather than returning them all. Count the custom segments in your "audiences" array before you finish; if there is more than one, cut it down. Whenever you do return a custom segment, say which one it is and why no RFP-selectable segment covered it -- and say it in the seller's words: **"a custom audience", or "not directly selectable in Salesforce"**, never by naming the catalog field the flag lives in. That flag is data for you to read, not vocabulary to repeat. Always mention a custom segment in one of the two lists: it is the only thing telling the seller this audience can't just be booked through Salesforce RFP.
- "audiences[].max_avails" is the audience's available impressions, and **it comes from the notes or it stays out**. Set it ONLY when the notes give a number for that specific audience -- "about a million a month on the home-services segment", "600K avails", "we can get 400,000 impressions against that one". Never estimate one, never carry a figure across from a different audience, and never work one back from a budget, a CPM or an impression goal: a made-up avails number looks exactly like a real one to the person sending the proposal, and the whole point of the field is that it came from the avails system. Leave it out entirely when the notes are silent -- an audience with no stated avails is flagged automatically for the seller to pull the real number, which is the correct outcome and needs no note from you.
- "audiences[].avails_basis" says which basis that number was given in: "monthly" for a per-month figure (what the avails system reports, and the default), "flight" for a whole-campaign total ("1.5 million over the three months"). Report the basis the notes used and let Python convert -- do not divide or multiply it yourself. If the notes give a number without saying which it is, use "monthly" and flag the assumption.
- "unresolved" is a list a salesperson reads before they send the proposal. **Write it for them, not for a developer.** Rules, all of which matter:
  * **Two sentences per item, maximum.** What you assumed, then what to confirm. Nothing else -- no reasoning, no justification, no explanation of how the app works.
  * **Name every thing the way a seller would say it out loud.** An audience that can't be booked straight through Salesforce RFP is "a custom audience", or "not directly selectable in Salesforce"; a one-time charge is "the production fee line"; a sports package is "NFL Regular Season"; a share of what's left of the budget is "the remaining budget"; a measurement option is "the brand lift study". Schema keys, catalog flags and product codes do their job inside the JSON fields themselves and have no business in a sentence a salesperson reads. If you can't put a thing in a seller's own words, leave it out of the list.
  * **At most 8 items, and fewer is better.** Merge anything related into one item -- three questions about the flight dates are one item about the flight dates. If you have more than 8, you're flagging things that don't need flagging.
  * Put anything the CLIENT has to answer in "unresolved". Put checks the SELLER does on their own -- pulling real avails numbers, confirming a rate internally, double-checking a segment is available, uploading a file, deciding which internal category or vertical a client gets filed under -- in "unresolved_internal". A client is never asked an internal-bookkeeping question (which vertical, which internal tag) -- that's a seller/system decision even when it happens to be framed as an open question in your own reasoning. Same rules apply to both. **This is the only rule that decides which list something goes in.** Everywhere else in these instructions that tells you to flag, note or say something, it means "write it as an item" and this rule alone picks the list -- so route by who has to act on it, never by where the wording of some other rule happens to point.
  * Write plain sentences, the way you would say them to a colleague: "No split was stated between the two audiences, so the budget was divided evenly. Confirm the intended split." State the assumption, then what to confirm, and stop there.

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


def _extract_largest_json_object(text):
    """The LARGEST balanced {...} object anywhere in `text`, respecting
    string literals so a brace inside a quoted value doesn't miscount --
    or None if no balanced object is found.

    Not the FIRST one -- that was tried and it silently corrupted a real
    draft. A preamble routinely quotes a small schema fragment verbatim
    while reasoning about a choice ("I'll use group_selection: {"mode":
    "all"}, since the notes describe the whole buy"), and that fragment is
    itself complete, valid JSON. Taking the first balanced object grabbed
    exactly that two-word fragment and returned it as the entire draft --
    no error, no crash, just 41,000 characters of a real LiveWell plan
    silently replaced by `{"mode": "all"}`, discovered only by reading the
    output rather than checking that it parsed. The real answer is
    overwhelmingly the largest object in the response, whether the noise
    around it is a preamble, trailing commentary, or both, so this scans
    every top-level object in the text and returns the longest."""
    candidates = []
    n = len(text)
    i = 0
    while i < n:
        if text[i] != "{":
            i += 1
            continue
        start = i
        depth = 0
        in_string = False
        escape = False
        matched_end = None
        j = i
        while j < n:
            ch = text[j]
            if in_string:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == '"':
                    in_string = False
            else:
                if ch == '"':
                    in_string = True
                elif ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        matched_end = j
                        break
            j += 1
        if matched_end is not None:
            candidates.append(text[start:matched_end + 1])
            i = matched_end + 1
        else:
            # This "{" never closes -- don't skip past it to j (== n),
            # which would abandon the scan; a later "{" might still start
            # a real, well-formed object.
            i = start + 1
    return max(candidates, key=len) if candidates else None


def _parse_draft_json(raw_text):
    try:
        return json.loads(raw_text), None
    except json.JSONDecodeError as exc:
        first_error = exc
    try:
        return json.loads(_strip_markdown_fences(raw_text)), None
    except json.JSONDecodeError:
        pass
    extracted = _extract_largest_json_object(raw_text)
    if extracted is not None:
        try:
            return json.loads(extracted), None
        except json.JSONDecodeError:
            pass
    return None, f"Claude's response wasn't valid JSON even after stripping markdown fences: {first_error}"


_CLAUDE_FAILURE_OUTCOMES = {"truncated", "empty", "unparseable", "refusal", "api_error"}


def log_claude_call(record):
    """Record one call, to the console, to a file, and (on a failure only)
    into session_state. Never raises.

    Console + file, deliberately: the console is what's visible in
    Streamlit Cloud's log viewer, the file is what survives locally after
    the tab is closed. A logging failure must never be what takes a draft
    down, so every error here is swallowed -- the draft is the point, the
    log is the evidence.

    Neither of those is reachable from where a rep actually is, though --
    Streamlit Cloud's log viewer is a developer tool, and the local file
    doesn't exist on the machine a rep is using at all (the LiveWell
    incident: nobody could see what Claude actually returned, and the
    on-screen message pointed at a path that was never going to be there).
    session_state["last_claude_failure"] is the fix -- `stop_reason` plus
    the first 200 characters of what Claude returned, exactly enough to
    diagnose without either an unbounded transcript or anything
    client-facing, which `capture_feedback_state` folds into a rep's
    "Report an issue" so it's readable from the admin page. Overwritten by
    every call (success clears it) rather than accumulated, so it always
    describes the failure that JUST happened, never a stale one from
    earlier in the session.
    """
    line = json.dumps(record, default=str)
    print(f"[claude] {line}")
    try:
        with open(CLAUDE_LOG_PATH, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass
    try:
        outcome = record.get("outcome")
        if outcome in _CLAUDE_FAILURE_OUTCOMES:
            head = (record.get("head") or record.get("error") or "")[:200]
            st.session_state["last_claude_failure"] = {
                "label": record.get("label"),
                "outcome": outcome,
                "stop_reason": record.get("stop_reason"),
                "head": head,
            }
        elif outcome == "ok":
            st.session_state.pop("last_claude_failure", None)
    except Exception:
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
    # Every message returned below is what a REP sees on screen, so it says
    # what happened in plain terms and what to do next -- retry, then
    # escalate via "Report an issue" -- and nothing dev-facing: no file
    # paths, no session keys, no function/constant names. The one incident
    # that made this the rule: a message once pointed a rep at
    # "%TEMP%/proposal_builder_claude_calls.log", a path that doesn't exist
    # on the machine she was actually using (Streamlit Cloud), and named
    # "ANTHROPIC_MAX_TOKENS in app.py" as if she could edit it. The real
    # diagnostic detail (stop_reason, the response's own head/tail) still
    # goes to `log_claude_call` on every branch below, same as always --
    # it's just not printed into the string a rep reads. See DECISIONS.md,
    # the LiveWell incident, and `capture_feedback_state`'s
    # `last_claude_failure`, which is where that detail actually surfaces.
    raw_text, facts = _response_text_and_facts(response)
    facts["label"] = label

    if facts["stop_reason"] == "max_tokens":
        log_claude_call({**facts, "outcome": "truncated"})
        return None, (
            f"The draft got too long to finish -- it hit Claude's {ANTHROPIC_MAX_TOKENS:,}-token "
            f"response limit and was cut off partway through. Try fewer plan options, fewer "
            f"lines per option, or shorter notes, then draft again. If a plan this size keeps "
            f"happening, use Report an issue and I'll take a look."), True

    if facts["stop_reason"] == "refusal":
        log_claude_call({**facts, "outcome": "refusal"})
        return None, (
            "Claude declined to answer this request. Re-word the notes and try again -- if "
            "they contain nothing unusual, use Report an issue and I'll take a look."), False

    if not raw_text.strip():
        log_claude_call({**facts, "outcome": "empty"})
        return None, (
            "The draft didn't come back in a usable form. Try again -- if it happens twice, "
            "use Report an issue and I'll take a look."), True

    parsed, parse_error = _parse_draft_json(raw_text)
    log_claude_call({**facts, "outcome": "ok" if parsed is not None else "unparseable"})
    if parsed is None:
        return None, (
            "The draft didn't come back in a usable form. Try again -- if it happens twice, "
            "use Report an issue and I'll take a look."), True
    return parsed, None, False


def _call_claude_json(prompt, label="draft", attempts=2, on_attempt=None):
    """Sends one prompt to Claude and parses the response as JSON. Returns
    (parsed_dict, error_message) -- exactly one is None.

    The original bug here (see DECISIONS.md, the LiveWell incident): a
    genuinely open-ended, judgment-heavy notes set made the model narrate
    ("I need to analyze these notes carefully...") instead of returning raw
    JSON, on BOTH attempts, because the retry resent a byte-identical
    prompt and had no reason to behave differently the second time. An
    assistant-turn prefill was tried as the structural fix and abandoned --
    ANTHROPIC_MODEL rejects it with a 400 (see the comment above
    `ANTHROPIC_MAX_TOKENS`) -- so the real defense is
    `_extract_largest_json_object` (`_parse_draft_json`'s fallback, which
    salvages the JSON even when a preamble gets through) plus the corrective
    retry below, not anything at the network-call level.

    Retries once on an empty, truncated or unparseable response before
    surfacing anything -- but the retry is no longer that identical resend.
    It appends a corrective instruction to the prompt ("return ONLY the
    JSON object...") so attempt 2 is actually a different request, not a
    re-roll of the same one. That correction is deliberately generic, not
    the previous attempt's own error text -- `interpret_claude_response`'s
    return value is rep-facing now (plain language, no dev detail), and
    feeding a rep-facing sentence back into the model as if it were
    technical guidance would be both useless to the model and a way for
    on-screen wording to leak into the next request's prompt. An API-level
    failure (auth, network, rate limit) is not retried here -- the SDK
    already retries those itself.

    `on_attempt(attempt, attempts)`, if given, fires before each attempt's
    network call -- the caller's hook for telling a rep what's happening
    instead of a spinner that just sits there (attempt 1 needs no comment;
    attempt 2 means the first pass didn't come back clean, which is worth
    saying).
    """
    api_key = st.secrets.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None, "ANTHROPIC_API_KEY is not set in .streamlit/secrets.toml."

    client = anthropic.Anthropic(api_key=api_key)
    error = "Claude was not called."
    attempt_prompt = prompt
    for attempt in range(1, attempts + 1):
        if on_attempt:
            on_attempt(attempt, attempts)
        try:
            response = client.messages.create(
                model=ANTHROPIC_MODEL,
                max_tokens=ANTHROPIC_MAX_TOKENS,
                messages=[{"role": "user", "content": attempt_prompt}],
            )
        except Exception as exc:                                 # noqa: BLE001
            log_claude_call({"label": label, "attempt": attempt, "outcome": "api_error",
                             "error": f"{type(exc).__name__}: {exc}",
                             "prompt_chars": len(attempt_prompt)})
            return None, ("Claude couldn't be reached. Try again in a moment -- if it keeps "
                          "failing, use Report an issue and I'll take a look.")

        parsed, error, retryable = interpret_claude_response(
            response, f"{label} (attempt {attempt} of {attempts})")
        if parsed is not None:
            return parsed, None
        if not retryable or attempt == attempts:
            break
        attempt_prompt = prompt + (
            "\n\nYour previous response could not be used. This time, return ONLY the JSON "
            "object -- nothing before the opening brace, nothing after the closing one, no "
            "reasoning or commentary.")
    return None, error


def _draft_attempt_status_updater(status):
    """`on_attempt` closure for an `st.status` element -- so a slow draft
    tells a rep what's happening instead of a spinner that just sits there.
    Attempt 1 keeps the status's initial label (the normal, single-attempt
    case); attempt 2+ means the previous attempt didn't come back as clean
    JSON, which is worth saying rather than leaving the rep to guess why a
    'quick' draft is taking a second pass."""
    def _update(attempt, attempts):
        if attempt > 1:
            status.update(label=f"First pass didn't come back as clean JSON -- "
                                 f"retrying (attempt {attempt} of {attempts})...")
    return _update


def call_claude_draft(notes, on_attempt=None):
    """Returns (draft_dict, error_message) -- exactly one is None.

    Reads `targeting_groups` from session_state so the prompt can tell the
    model to sell FROM real, already-resolved rows (an avails import that
    ran before this draft) rather than inventing its own -- see
    build_draft_prompt's `existing_groups` parameter. Also reads the setup
    band's own market_choice/flight_start/flight_end/avails_basis
    (FLOW_REWORK_PLAN.md Phase 5) and passes them through as given context,
    the same shape as existing_groups -- this function is where session_state
    gets read, build_draft_prompt itself stays a pure function of its
    arguments.

    `on_attempt` passes straight through to `_call_claude_json` -- see its
    docstring.
    """
    existing_groups = st.session_state.get("targeting_groups")
    return _call_claude_json(build_draft_prompt(
        notes, existing_groups,
        market_choice=st.session_state.get("market_choice"),
        flight_start=st.session_state.get("flight_start"),
        flight_end=st.session_state.get("flight_end"),
        avails_basis=st.session_state.get("avails_basis")), label="draft", on_attempt=on_attempt)


def build_redraft_prompt(notes, previous_draft, clarifications, existing_groups=None, *,
                         market_choice=None, flight_start=None, flight_end=None,
                         avails_basis=None):
    """A revision pass, not a fresh draft: the model gets its own previous
    JSON back plus the user's answers to the open questions, and is told to
    change only what the answers actually bear on. Starting over would churn
    parts of the draft the user already accepted.

    market_choice/flight_start/flight_end/avails_basis pass straight through
    to build_draft_prompt (FLOW_REWORK_PLAN.md Phase 5) -- the band's frame
    is still given context on a redraft, same as the first draft."""
    return build_draft_prompt(notes, existing_groups, market_choice=market_choice,
                              flight_start=flight_start, flight_end=flight_end,
                              avails_basis=avails_basis) + f"""

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


def call_claude_redraft(notes, previous_draft, clarifications, on_attempt=None):
    """Returns (draft_dict, error_message) -- exactly one is None.

    The revision is merged *over* the previous draft rather than replacing
    it. The model is told to return the whole object, but a dropped field
    would otherwise silently fall back to a schema default rather than to
    "unchanged" -- the incident that settled this: a re-draft once dropped
    the (then-existing) `agency_involved` field, and a missing boolean
    doesn't read as "no opinion", it reads as false, which quietly repriced
    every line at net instead of gross (see DECISIONS.md). That specific
    field is gone now (FLOW_REWORK_PLAN.md Phase 4b -- the gross-up is a
    rep-only checkbox, never drafted), but the general failure mode isn't:
    a dropped `cpm` override reads as "use the rate card" rather than "keep
    what I had", for exactly the same reason. Anything the revision does
    return still wins. Same reasoning covers "group_selection"/
    "group_allocation": read `apply_draft_to_form`'s `match_groups_to_selection`
    docstring for why an explicitly EMPTY `group_selection` is treated as
    absent rather than "select nothing".
    """
    existing_groups = st.session_state.get("targeting_groups")
    revised, error = _call_claude_json(
        build_redraft_prompt(
            notes, previous_draft, clarifications, existing_groups,
            market_choice=st.session_state.get("market_choice"),
            flight_start=st.session_state.get("flight_start"),
            flight_end=st.session_state.get("flight_end"),
            avails_basis=st.session_state.get("avails_basis")), label="redraft", on_attempt=on_attempt)
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


# --- Setup band (FLOW_REWORK_PLAN.md Phase 1) -----------------------------
# Five decisions that cascade -- originating market, flight dates, plan
# basis, avails mode, Total TV -- get ONE owner each: `setup_snapshot`
# captures them at Generate time (from wherever their widgets currently
# live -- Phase 1's own UI commit relocates the widgets, not this
# function), and `resolve_setup` reads them back for ANY proposal, old or
# new. Nothing downstream should read market_choice/flight_start/
# avails_basis/total_tv out of a stored form_json directly -- go through
# one of these two instead, the same "single owner" discipline `_market_
# display_name`/`geo_column_default` already apply to originating-market
# and geo text.
def setup_snapshot(market_choice, flight_start, flight_end, avails_basis, avails_mode, total_tv):
    """The setup band's five values, captured at Generate time straight off
    their own widgets -- every one of the five now has a real widget behind
    it (the band, built in this same commit).

    Additive form_json key ("setup"): a proposal logged before this exists
    has no such section, which is exactly the case `resolve_setup` handles.
    """
    return {
        "originating_market": market_choice,
        "flight_start": str(flight_start) if flight_start else None,
        "flight_end": str(flight_end) if flight_end else None,
        "plan_basis": avails_basis,
        "avails_mode": bool(avails_mode),
        "total_tv": bool(total_tv),
    }


def resolve_setup(form, groups=None, row_market=None):
    """The setup band's five values for ANY proposal, old or new -- the
    migration-read half of `setup_snapshot`.

    A proposal logged after Phase 1 carries its own "setup" section
    verbatim, which is authoritative: a rep may have deliberately set the
    band to something the legacy fields alone wouldn't imply (avails mode
    off on a proposal whose groups still carry priced avails from before
    it was toggled off, say).

    A pre-rework proposal has no "setup" key. Four of the five values are
    then a straight relocation of reads that already exist elsewhere in
    this file (`rehydrate_proposal_into_form`'s `selections.get("market")
    or row.get("market")`, its `flight` dict read, and `form.get(
    "avails_basis")`) -- not new inference. Only "avails_mode" has no
    legacy field to fall back to, because the pre-band form never asked
    the question: it defaults True whenever the proposal's OWN groups
    (already migrated from `avails_rows`/`targeting_groups` by
    `seed_rows_to_groups` before this is called) carry any avails -- a
    proposal actually built from an avails document always has some group
    with avails_monthly > 0, and a hand-typed, avails-free proposal never
    does.
    """
    stored = form.get("setup")
    if isinstance(stored, dict):
        return {
            "originating_market": stored.get("originating_market"),
            "flight_start": stored.get("flight_start"),
            "flight_end": stored.get("flight_end"),
            "plan_basis": stored.get("plan_basis") or AVAILS_BASIS_MONTHLY,
            "avails_mode": bool(stored.get("avails_mode", True)),
            "total_tv": bool(stored.get("total_tv", False)),
        }

    selections = form.get("selections") or {}
    flight = form.get("flight") or {}
    plan_basis = form.get("avails_basis")
    if plan_basis not in (AVAILS_BASIS_MONTHLY, AVAILS_BASIS_FLIGHT):
        plan_basis = AVAILS_BASIS_MONTHLY
    market = selections.get("market") or row_market
    return {
        "originating_market": market if market in ("DC", "Harrisburg") else None,
        "flight_start": flight.get("start"),
        "flight_end": flight.get("end"),
        "plan_basis": plan_basis,
        "avails_mode": any(int(g.get("avails_monthly") or 0) > 0 for g in (groups or [])),
        "total_tv": bool((selections.get("products") or {}).get("total_tv")),
    }


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
    # The stored flight text, not a recomputed one -- same "verbatim, never
    # recomputed" rule as the media plan itself. This is Generate's own
    # fallback for an empty Timing narrative (main(), the live TIMING_BULLETS
    # line); rebuild used to fall back to "--" instead, a second, drifted
    # construction of the same fallback that made a rebuilt deck read
    # differently from the one the client actually received whenever Timing
    # was left blank. Caught by tests/test_group_backward_compat.py.
    #
    # Prefers the day-precise "shorthand" (FLOW_REWORK_PLAN.md Phase 2) over
    # the whole-month "label", falling back to the label for a proposal
    # logged before Phase 2 ever stored a shorthand -- never recomputed from
    # the flight dates, which is exactly the drift that made this fallback
    # disagree with Generate's in the first place.
    stored_flight = form.get("flight") or {}
    stored_flight_display = stored_flight.get("shorthand") or stored_flight.get("label") or "--"

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
            "TIMING_BULLETS": lines_to_bullets(specs.get("timing", "")) or [stored_flight_display],
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

    # form.get("vault_slides") or [] -- the dict-default form, so a proposal
    # logged before this feature existed (no key at all) rebuilds exactly as
    # it always did, never picking up vault slides nobody chose.
    vault_slide_sources = []
    for vault_slide in form.get("vault_slides") or []:
        stored, error = db.fetch_slide_vault_entry(vault_slide.get("id"))
        if error or not stored:
            warnings.append(f"Vault slide \"{vault_slide.get('title')}\" is no longer in the "
                            f"vault -- rebuilt without it.")
            continue
        try:
            placement = vault_slide.get("placement") or assembly.VAULT_PLACEMENT_BEFORE_PLAN
            vault_slide_sources.append(vault_slide_source(stored, placement))
        except Exception as exc:
            warnings.append(f"Vault slide \"{stored.get('title')}\" couldn't be fetched "
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
        assembly.append_vault_slides(prs, vault_slide_sources)
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

    # FLOW_REWORK_PLAN.md Phase 4b: "agency_gross_up" being present in the
    # stored form_json means this proposal was logged AFTER the checkbox
    # rebuild -- its rows are already net, and the field reads straight
    # through. Its ABSENCE means an older proposal, whose rows were stored
    # GROSS whenever the old per-row toggle was on (form["markup"] == 1.15,
    # the old toggle's own effective-markup value) -- `_legacy_gross_migration`
    # is read again below, alongside the plan_options loop, which is where
    # the one-time Cost migration actually happens (see the comment there).
    # Impressions need no migration either way: old impressions were already
    # cost/(cpm*1.15)*1000, and net cost over the same CPM gives the
    # identical number.
    if "agency_gross_up" in form:
        updates["agency_gross_up"] = bool(form.get("agency_gross_up"))
        _legacy_gross_migration = False
    else:
        _legacy_gross_migration = form.get("markup") == 1.15
        updates["agency_gross_up"] = _legacy_gross_migration
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
        # resolve_flight_months reads this same form dict -- a post-Phase-2
        # proposal's own stored month_ranges verbatim, or migrated defaults
        # from its legacy whole-month active_months for a pre-Phase-2 one.
        # Written back into flight_months explicitly: whoever replaces
        # state owns all of it.
        flight_ranges = resolve_flight_months(form) or flight_month_ranges(start, end)
        all_months = month_list(start, end)
        updates["active_months"] = active_month_labels(flight_ranges) or all_months
        updates["flight_months"] = flight_months_snapshot(flight_ranges)
        flight_label = format_flight_label(all_months, updates["active_months"]) or "TBD"
        flight_shorthand = format_flight_shorthand(flight_ranges) or flight_label
    else:
        notes.append("Flight dates couldn't be restored -- check Section B before generating.")
        flight_label = flight.get("label") or "TBD"
        flight_shorthand = flight.get("shorthand") or flight_label

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
        # A group logged before include_in_plan existed has no such key.
        # Missing there means "whatever its rows already say" -- the legacy
        # value has to reproduce old behavior, not the new default (same
        # shape as sports.get("viewership", True)): derived from whether the
        # group's id appears in any stored option's own _group_ids, and
        # locked, so nothing later (a re-sync, a re-draft, a fresh import)
        # re-derives it out from under a proposal that already shipped. A
        # group missing the field on a genuinely NEW proposal is a
        # different question entirely and stays False via new_group's own
        # default -- this branch only fires for a proposal logged before
        # this feature existed at all.
        if any("include_in_plan" not in g for g in updates["targeting_groups"]):
            plan_group_ids = {gid for opt in (form.get("plan_options") or [])
                              for prow in (opt.get("rows") or [])
                              for gid in group_ids_of(prow)}
            updates["targeting_groups"] = [
                g if "include_in_plan" in g else
                dict(g, include_in_plan=(g.get("id") in plan_group_ids), include_locked=True)
                for g in updates["targeting_groups"]]
    elif "avails_seed_rows" in updates:
        # Older still: no targeting_groups key at all, only flat avails
        # rows -- a proposal from before groups existed, when every avails
        # row drove a plan line automatically (there was no "research only"
        # concept yet). Migrating it starts every group included and
        # locked, the historically faithful legacy value, not today's
        # opt-in default.
        updates["targeting_groups"] = [
            dict(g, include_in_plan=True, include_locked=True)
            for g in tg.seed_rows_to_groups(updates["avails_seed_rows"], AVAILS_COLUMN_MONTHLY)]
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
        if _legacy_gross_migration:
            # FLOW_REWORK_PLAN.md Phase 4b migration: this proposal's rows
            # were stored GROSS (the old per-row toggle applied x1.15 to
            # Cost directly). The new checkbox grosses at display time only
            # -- the grid holds net -- so loading this unchanged would gross
            # it a SECOND time the instant agency_gross_up (set True above)
            # renders. Divide back to net, once, here. Flat-fee and
            # broadcast rows were never grossed in the old system either
            # (a flat fee's Cost never passed through the markup math; a
            # broadcast row was always exempt) so both are left alone --
            # dividing them would be the actual corruption this guards
            # against, not a fix.
            for row in rows:
                if not is_flat_fee_row(row) and not is_broadcast_row(row):
                    row["Cost"] = _num(row.get("Cost")) / 1.15
        driver = list(stored.get("driver") or [])
        if len(driver) != len(rows):
            driver = [DRIVER_COST] * len(rows)
        option = new_plan_option(stored.get("name") or DEFAULT_OPTION_NAMES[0], rows,
                                 driver=driver,
                                 breakout=stored.get("breakout") or BREAKOUT_MONTHLY,
                                 # A restored breakout is a deliberate stored
                                 # value, same as every other restored field
                                 # here -- it must never get silently synced
                                 # to whatever today's band happens to say.
                                 breakout_locked=True)
        # Every restored row is a deliberate value, never a seed default.
        option["dirty"] = [True] * len(rows)
        plan_options.append(option)
    if plan_options:
        updates["plan_options"] = plan_options
        # Fresh widget keys, or the existing per-option name/breakout widgets
        # would overwrite the restored ones with whatever they already hold.
        bump_plan_options_generation(updates)

    # --- case studies -----------------------------------------------------
    for key in [k for k in st.session_state if k.startswith("cs_pick_")]:
        updates[key] = False
    for case_study in form.get("case_studies") or []:
        updates[f"cs_pick_{case_study.get('id')}"] = True

    # --- vault slides -------------------------------------------------------
    for key in [k for k in st.session_state if k.startswith("vault_pick_")]:
        updates[key] = False
    for vault_slide in form.get("vault_slides") or []:
        vault_id = vault_slide.get("id")
        updates[f"vault_pick_{vault_id}"] = True
        updates[f"vault_place_{vault_id}"] = (
            vault_slide.get("placement") or assembly.VAULT_PLACEMENT_BEFORE_PLAN)

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
    updates["_shared_fields_key"] = (default_targeting + "||" + default_geo
                                     + "||" + flight_label + "||" + flight_shorthand)

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

    # Groups own the plan (see reconcile_group_plan_lines / CLAUDE.md's "The
    # media plan and the form") the moment a REAL group already exists --
    # audience-bearing, not a `_placeholder` starter row. Computed once, from
    # session_state as it stood BEFORE this draft (never from `updates`,
    # which this function hasn't started writing yet): an avails import that
    # ran before this draft is exactly the case that must not be overwritten.
    groups_own_plan = any(
        g.get("terms") and not g.get("_placeholder")
        for g in (st.session_state.get("targeting_groups") or []))

    vertical_val = draft.get("vertical")
    vertical_label = next((label for label, key in VERTICALS.items() if key == vertical_val), None)
    if vertical_label:
        updates["vertical_choice"] = vertical_label
    elif vertical_val:
        internal.append(f"Vertical '{vertical_val}' not recognized -- left unchanged.")

    # FLOW_REWORK_PLAN.md Phase 5: originating market is a band INPUT now,
    # never a drafted output -- the model is no longer asked for it at all
    # (DRAFT_JSON_SCHEMA_EXAMPLE dropped "market"), and this reads the band's
    # own current value instead of writing one. The band's setup gate
    # guarantees market_choice is "DC"/"Harrisburg" by the time anything
    # below it can render, but the draft button itself lives inside the band
    # and isn't gated (typing notes and clicking Draft is never blocked) --
    # so a rep who drafts before picking a market gets market_label="" here,
    # same as the pre-gate state always produced before this phase existed.
    market_choice_now = st.session_state.get("market_choice")
    market_label = ("Washington, DC DMA" if market_choice_now == "DC"
                    else "Harrisburg DMA" if market_choice_now == "Harrisburg" else "")

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
    # FLOW_REWORK_PLAN.md Phase 6: plain labels for avails-resolved markets
    # still present in Target DMAs, folded into Geography text below
    # regardless of whether target_dma_choice itself got written this call
    # -- initialized here so it exists even when target_market_names is
    # empty (that path never touches Target DMAs at all, so there's nothing
    # to report).
    avails_union_geo_labels = []
    # Target DMAs is an AVAIL-FILLS-FACTS field, same tier as client_name --
    # seeded from the notes only when nothing has claimed it yet (no avails
    # import, no rep pick); once real, drafting never writes it again, only
    # flags a disagreement. This is what makes the old ledger bug (import
    # resolving Harrisburg, draft naming Denver, Harrisburg silently
    # dropped) structurally unreachable rather than guarded against: there
    # is no longer a second writer for _group_markets_applied's own
    # monotone-add contract to race, so the union machinery that used to
    # live here is gone. The one exception is self-serve, notes-only
    # proposals with no avails import at all -- there Target DMAs has no
    # other way to fill, which is exactly the case the empty-check below
    # still covers.
    current_target_dma_choice = st.session_state.get("target_dma_choice") or []
    if target_market_names:
        profiles, _ = load_market_profiles()
        by_key = {r.get("key"): r for r in profiles}
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
            # Geography TEXT still needs an avails-resolved market
            # represented even when the model's own geography prose never
            # mentioned it -- independent of whether Target DMAs itself gets
            # written below. Same walk over _group_markets_applied as
            # before, just no longer feeding a target_dma_choice write.
            applied_keys = st.session_state.get("_group_markets_applied") or set()
            for key in applied_keys:
                row = by_key.get(key)
                if row is None:
                    continue
                option_label = market_profile_option_label(row)
                if option_label not in current_target_dma_choice:
                    continue
                plain_label = row.get("label")
                if plain_label and plain_label not in avails_union_geo_labels:
                    avails_union_geo_labels.append(plain_label)

            if not current_target_dma_choice:
                # Nothing has claimed Target DMAs yet -- a self-serve,
                # notes-only proposal with no avails import. Seed straight
                # from the notes; no union needed, since there is no
                # existing list for a second writer to have raced.
                updates["target_dma_choice"] = chosen_labels
                final_keys = {r.get("key") for r in profiles
                             if market_profile_option_label(r) in chosen_labels}
                updates["include_market_profile"] = any(
                    r.get("image_path") for r in profiles
                    if r.get("key") in final_keys)
            else:
                missing = [label for label in chosen_labels
                          if label not in current_target_dma_choice]
                if missing:
                    internal.append(
                        f"The notes mention {', '.join(missing)}, not in Target DMAs -- "
                        f"add {'it' if len(missing) == 1 else 'them'} by hand if that's real.")

    # client_name is the other avail-fills-facts field, and the one with a
    # real widget-instantiation hazard: it renders in the setup band, above
    # the Draft button, so writing it here (inside `updates`, applied
    # synchronously below, in this same run) raises exactly like the bug
    # this replaces. Seeding it has to go through the same queue-then-apply
    # shape the avails importer already uses for the same field
    # (`_draft_pending_fields` / apply_pending_draft_fields, called at the
    # top of the NEXT run, before the band's first widget). "Empty" reuses
    # apply_avails_import's own sentinel for this exact field ("Acme Test
    # Co", the untouched default) rather than inventing a second notion of
    # empty.
    draft_client_name = (draft.get("client_name") or "").strip()
    current_client_name = st.session_state.get("client_name")
    if draft_client_name:
        if current_client_name in (None, "", "Acme Test Co"):
            pending_fields = dict(st.session_state.get("_draft_pending_fields") or {})
            pending_fields["client_name"] = draft_client_name
            st.session_state["_draft_pending_fields"] = pending_fields
        elif current_client_name != draft_client_name:
            internal.append(
                f"The notes name the client as \"{draft_client_name}\" -- the Client name "
                f"field already says \"{current_client_name}\". Left as-is -- update it by "
                f"hand if the notes are right.")

    # FLOW_REWORK_PLAN.md Phase 4b: the agency gross-up is a rep-only,
    # order-wide calculator now, not a drafted field -- the model transcribes
    # whatever rate the notes state, gross or net, and never touches this
    # checkbox. There is nothing to write here, and no skip-sections special
    # case is needed: the checkbox simply keeps whatever the rep already had
    # it set to, exactly like any other widget a draft doesn't mention.
    updates["spanish_campaign"] = bool(draft.get("spanish_campaign", False))
    touched_sections.add("basics")

    # "show_sov" (share of voice / % of avails on the media plan) had no
    # field at all before this -- a rep asking to "include the SOV" in the
    # notes could only ever be flagged in unresolved, never actually turned
    # on, since there was nowhere for the model to put that intent. Always
    # safe to set: it only ever affects DISPLAY, gated per-line on a real
    # avails match (see the checkbox's own help text), never what's sold.
    updates["show_sov"] = bool(draft.get("show_sov", False))

    # FLOW_REWORK_PLAN.md Phase 5: the flight is a band INPUT now, never a
    # drafted output. This used to reconcile a DRAFTED flight against the
    # form's own -- the real incident that machinery fixed (an Oct-Dec draft
    # landing on a stale Sep-Nov month selection, $45,000 spread over three
    # months but totalled over two) required two competing flights to exist
    # in the first place. There is only one now: the band's. So this always
    # takes what used to be the fallback branch, unconditionally -- not
    # merely safe, but simpler, and the incident is unreachable by
    # construction rather than guarded against (see DECISIONS.md). If the
    # notes state a date that conflicts with the band's, the prompt tells
    # the model to flag it in unresolved_internal naming both, never to
    # return a flight_start/flight_end field -- there is no longer a schema
    # slot for one.
    current_start = st.session_state.get("flight_start") or DEFAULT_FLIGHT_START
    current_end = st.session_state.get("flight_end") or DEFAULT_FLIGHT_END
    if isinstance(current_start, date) and isinstance(current_end, date):
        draft_ranges = flight_month_ranges(current_start, current_end,
                                           st.session_state.get("flight_months"))
    else:
        draft_ranges = []
    all_flight_months, draft_months = form_flight_months()
    flight_label = format_flight_label(all_flight_months, draft_months) or "TBD"
    flight_shorthand = format_flight_shorthand(draft_ranges) or flight_label
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
    geo_bullets = list(specs.get("geography")
                       or ([geo] if geo else None)
                       or draft_geo_labels
                       or ([market_label] if market_label else []))
    # FLOW_REWORK_PLAN.md Phase 5 ledger fix, continued: an avails-derived
    # market has to survive here too, even when the model's own geography
    # prose (specs.geography / geo) took precedence and never mentioned it
    # -- otherwise the picker is correct but the Geography text a client
    # actually sees still silently drops the market the avails resolved.
    for label in avails_union_geo_labels:
        if label not in geo_bullets:
            geo_bullets.append(label)
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
    if groups_own_plan:
        # Real groups already own this proposal's avails table -- drafting
        # never overwrites it (this is the fix for the bug that made
        # import-then-draft and draft-then-import diverge: a draft used to
        # rebuild targeting_groups from scratch here even when a real
        # import already resolved real geo_def/resolved_zips for it).
        # match_groups_to_selection (in the media-plan block below) is what
        # actually connects a drafted audience to a real, existing group;
        # this only reports what the model wanted to add that ISN'T already
        # on the table, so nothing is silently invented OR silently lost.
        real_groups = [g for g in (st.session_state.get("targeting_groups") or [])
                      if g.get("terms") and not g.get("_placeholder")]
        known_labels = {tg.audience_label(g).strip().lower() for g in real_groups}
        unknown = sorted({str(a.get("segment") or "").strip()
                          for a in matched_audiences + unmatched_with_avails
                          if str(a.get("segment") or "").strip().lower() not in known_labels})
        if unknown:
            internal.append(
                f"The notes mention {', '.join(unknown)}, which {'is' if len(unknown) == 1 else 'are'} "
                f"not on the avails table -- add {'it' if len(unknown) == 1 else 'them'} in D2 and tick "
                f"Plan if {'it' if len(unknown) == 1 else 'they'} should be sold.")
    elif matched_audiences or unmatched_with_avails:
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
        # FLOW_REWORK_PLAN.md Phase 5: Plan basis is a band INPUT, never a
        # drafted output -- this used to force avails_basis to Monthly
        # unconditionally, silently overwriting whatever the rep had
        # already set in the band. That was a real, live bug (see
        # DECISIONS.md): every draft that seeded its own avails rows reset
        # Plan basis to Monthly regardless of the rep's own choice, which is
        # likely why a rep-set Full Flight basis kept "not sticking." The
        # avails figures below are always stored monthly internally
        # regardless (unchanged) -- only the DISPLAY basis was ever wrong to
        # touch, and now nothing here touches it at all.
        # A drafted audience becomes a group the same way any other seed row
        # does: one drafted segment -> one single-term group -> one plan
        # line; a canonical "(A) AND (B)" segment (see above) becomes a real
        # two-term group instead of two separate lines. Nothing else about
        # drafting changes -- percent_of_avails/avails_ref still resolve
        # through the same avails_lookup figure as today. (Only reached when
        # NOT groups_own_plan -- see above -- "drafting only creates lines
        # from scratch when there are no groups".)
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
    # proposal and renders with no option label anywhere. Every row prices
    # at its NET rate now (FLOW_REWORK_PLAN.md Phase 4b) -- there is no
    # markup to thread through this any more.
    options_in = drafted_options(draft)
    drafted_plan_options = []
    lines_valid = False
    touched_products = set()
    touched_sports = set()
    # (option name, stated gross figure, converted net figure) for every KEPT
    # option whose total_budget_basis was "gross" -- populated only once an
    # option actually survives into drafted_plan_options (below), so a
    # dropped option's conversion never triggers the checkbox for nothing.
    gross_budget_options = []
    # What a BRAND-NEW, untouched option would default to right now, reading
    # the band's Plan basis exactly like the per-option radio widget's own
    # resync does (default_breakout_for_basis). Used below to decide whether
    # a drafted option's breakout is a genuine deliberate signal from the
    # notes (locks, same as a rep's own radio click) or just the model
    # landing on the same default the band would have given it anyway (stays
    # unlocked, so it keeps mirroring the band -- see the incident this
    # fixes in DECISIONS.md).
    band_default_breakout = default_breakout_for_basis(st.session_state.get("avails_basis"))

    # groups_own_plan: the working copy of targeting_groups this draft may
    # pre-select into, and the running list of per-option intent -- both
    # written to `updates` once, after the loop, never per-option (several
    # options can select the same group without racing each other).
    real_groups = ([g for g in (st.session_state.get("targeting_groups") or [])
                    if g.get("terms") and not g.get("_placeholder")] if groups_own_plan else [])
    groups_by_id = {g["id"]: g for g in real_groups}
    pending_groups = ([dict(g) for g in (st.session_state.get("targeting_groups") or [])]
                      if groups_own_plan else [])
    intent_options = []

    for opt_in in options_in:
        # The one arithmetic step total_budget_basis authorizes: a budget the
        # notes explicitly labelled gross is divided by AGENCY_MARKUP here,
        # once, before anything downstream (the $0 check, resolve_drafted_
        # lines, intent_options) ever sees it -- every one of them then
        # operates on the same net figure they always have, with no idea a
        # conversion happened. gross_budget_stated is the ORIGINAL figure,
        # kept only to name it in the internal note once this option is
        # confirmed kept, below.
        gross_budget_stated = None
        if opt_in.get("total_budget_basis") == "gross" and opt_in["total_budget"] > 0:
            gross_budget_stated = opt_in["total_budget"]
            opt_in["total_budget"] = round(gross_budget_stated / AGENCY_MARKUP)

        # An option that names lines but carries no budget can only resolve to
        # a grid of zeros. Every such line then trips the $0 drop below and
        # the option disappears -- so say plainly what went wrong and what to
        # do about it, rather than reporting a bare "no usable lines" for what
        # is really a missing number. A drafted plan must never come back as a
        # stated budget with nothing costed against it.
        if (opt_in["lines"] or opt_in.get("group_selection")) and opt_in["total_budget"] <= 0:
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

        lines_for_waterfall = opt_in["lines"]
        matched_ids = []
        avails_by_group_id = None
        if groups_own_plan:
            # A model that emits its own Premion Streaming TV line despite
            # the prompt's instructions is never silently discarded -- its
            # allocation/cpm is adopted as the group allocation (unless the
            # option already named one explicitly), and it's reported
            # either way, since nothing downstream else expects a
            # "premion_streaming_tv" entry once groups own this plan.
            non_premion, premion_extra = [], []
            for line in opt_in["lines"]:
                (premion_extra if line.get("product") == GROUP_LINE_PRODUCT_KEY
                 else non_premion).append(line)
            group_allocation = opt_in.get("group_allocation")
            group_cpm = None
            if premion_extra:
                first = premion_extra[0]
                if group_allocation is None:
                    group_allocation = first.get("allocation")
                if first.get("cpm") not in (None, ""):
                    group_cpm = first.get("cpm")
                internal.append(
                    f'"{opt_in["name"]}" drafted a Premion Streaming TV line directly even '
                    f"though targeting groups already own this proposal's avails table -- "
                    f"used it as the group allocation instead of adding it as its own line.")
            elif opt_in.get("group_cpm") not in (None, ""):
                group_cpm = opt_in.get("group_cpm")

            matched_ids, sel_unmatched, sel_mode = match_groups_to_selection(
                real_groups, opt_in.get("group_selection") or {})
            for needle in sel_unmatched:
                unresolved.append(
                    f'"{opt_in["name"]}" wanted to sell "{needle}", which doesn\'t match anything '
                    f"on the avails table -- pick the row by hand if it should be included.")

            # Pre-selection never overrides a rep's own choice.
            lockable_ids = [gid for gid in matched_ids
                            if not groups_by_id.get(gid, {}).get("include_locked")]
            skipped_locked = [gid for gid in matched_ids if gid not in lockable_ids]
            for g in pending_groups:
                if g["id"] in lockable_ids:
                    g["include_in_plan"] = True

            # Mandatory disclosure, routed by mode -- never collapsed. Only
            # for options that actually matched something: an option with
            # no real selection at all has nothing to disclose beyond the
            # unmatched-phrase notes above.
            if matched_ids:
                if sel_mode == "named":
                    reason = (opt_in.get("group_selection_reason")
                             or f'"{opt_in["name"]}" sells the audiences the notes named.')
                    unresolved.append(reason + " Confirm that's what's being sold.")
                else:
                    internal.append(
                        f'"{opt_in["name"]}": the notes didn\'t say which of the {len(real_groups)} '
                        f"audience(s) on the avails table to sell, so all of them are on the "
                        f"plan. Untick any that are opportunity rather than this buy.")
            left_off = [g for g in real_groups if g["id"] not in matched_ids]
            if left_off:
                internal.append(
                    f'"{opt_in["name"]}" -- available but not on the plan: '
                    + ", ".join(tg.audience_label(g) for g in left_off) + ".")
            if skipped_locked:
                kept_labels = [tg.audience_label(groups_by_id[gid]) for gid in skipped_locked]
                internal.append(f"Kept your own choice on {', '.join(kept_labels)} -- a draft "
                                f"never overrides a plan inclusion you already set.")

            group_lines, avails_by_group_id = _synthesize_group_lines(
                matched_ids, group_allocation, group_cpm, groups_by_id)
            lines_for_waterfall = non_premion + group_lines
            # The RAW selection criteria, not the resolved matched_ids --
            # apply_draft_plan_intent_to_new_groups re-runs
            # match_groups_to_selection against whatever groups exist LATER
            # (a second avails import), and a stale id from THIS batch would
            # never match a different one. `match` phrases and "all" both
            # re-resolve correctly against a new batch on their own terms;
            # exact `ids` correctly apply to nothing outside the batch they
            # named. `matched_ids`/`unmatched` ride along too, purely as a
            # record of what THIS draft resolved -- never read back by the
            # matcher.
            stored_selection = dict(opt_in.get("group_selection") or {}) or {"mode": "all"}
            stored_selection.setdefault("mode", sel_mode)
            intent_options.append({
                "name": opt_in["name"], "total_budget": opt_in["total_budget"],
                "breakout": opt_in["breakout"],
                "group_allocation": group_allocation or {"split_evenly": True},
                "group_cpm": group_cpm, "other_lines": non_premion,
                "selection": {**stored_selection,
                             "reason": opt_in.get("group_selection_reason") or "",
                             "matched_ids": matched_ids, "unmatched": sel_unmatched},
            })

        # Avails for a `percent_of_avails` line: the avails THIS draft just
        # wrote (groups_own_plan is False here, so `updates` holds them) --
        # or, once groups own the plan, the real avails table already on
        # the form, since this draft wrote no avails rows of its own.
        avails_source = (st.session_state.get("avails_seed_rows") if groups_own_plan
                         else updates.get("avails_seed_rows"))
        rows, opt_products, opt_sports, opt_unresolved, opt_drivers = resolve_drafted_lines(
            lines_for_waterfall, opt_in["total_budget"],
            flight_shorthand, geo_or_market, default_targeting,
            avails_by_name=avails_lookup(avails_source),
            n_months=draft_n_months, avails_by_group_id=avails_by_group_id)
        # FLOW_REWORK_PLAN.md Phase 3: expand each entity's single resolved
        # row back out to one row per selected group id -- see
        # _expand_entity_group_rows' own docstring for why this is needed
        # (the waterfall is correctly entity-collapsed; the FINAL row set
        # must not be).
        rows, opt_drivers = _expand_entity_group_rows(
            rows, opt_drivers, matched_ids, groups_by_id, default_targeting, flight_shorthand)
        internal.extend(opt_unresolved)
        touched_products |= opt_products
        touched_sports |= opt_sports
        if matched_ids:
            touched_products = touched_products | {GROUP_LINE_PRODUCT_KEY}
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
                spread_rows_over_months(rows, draft_n_months)
            option = new_plan_option(opt_in["name"], rows, driver=opt_drivers,
                                     breakout=breakout,
                                     # Locked only when the drafted breakout
                                     # actually DIVERGES from what the band's
                                     # own Plan basis would have given a
                                     # brand-new option anyway -- the same
                                     # "did this genuinely change" test the
                                     # rep's own Breakout radio uses to decide
                                     # whether ITS click deserves a lock.
                                     # Unconditionally locking every drafted
                                     # option (the old behaviour) treated the
                                     # model's default fallback -- "monthly"
                                     # whenever the notes don't address
                                     # presentation at all, which is most
                                     # drafts -- as though it were an
                                     # explicit, deliberate choice, and that
                                     # froze the option out of the band mirror
                                     # forever: flipping the band's Plan basis
                                     # after drafting moved the avails table
                                     # but silently left the plan behind. A
                                     # draft that DOES say something real
                                     # ("show it as one combined flight
                                     # total") still locks, exactly as before
                                     # -- see DECISIONS.md.
                                     breakout_locked=(breakout != band_default_breakout))
            # A group-derived row starts CLEAN -- it's owned by its group's
            # include_in_plan flag (reconcile_group_plan_lines), not
            # protected by dirty, and only becomes dirty the moment a rep
            # actually hand-edits it (reconcile_plan_rows's own sticky-dirty
            # compare). Every other drafted row is dirty=True exactly as
            # before -- deliberate, never re-seeded away.
            option["dirty"] = [False if group_ids_of(r) else True for r in rows]
            drafted_plan_options.append(option)
            if gross_budget_stated is not None:
                gross_budget_options.append(
                    (opt_in["name"], gross_budget_stated, opt_in["total_budget"]))

    if groups_own_plan:
        # FLOW_REWORK_PLAN.md Phase 3: entity naming against the WORKING
        # COPY (pending_groups), never the live session_state groups
        # directly -- this function's own "all updates applied together or
        # not at all" guarantee (see its docstring) depends on nothing
        # being mutated in place before the single write below.
        raw_group_entities = draft.get("group_entities") or []
        entity_unresolved, entity_internal = apply_draft_group_entities(
            pending_groups, raw_group_entities)
        unresolved.extend(entity_unresolved)
        internal.extend(entity_internal)
        updates["targeting_groups"] = pending_groups
        updates["draft_plan_intent"] = {
            "source": "draft", "round": (st.session_state.get("draft_round") or 0) + 1,
            "flight_label": flight_label, "default_targeting": default_targeting,
            "n_months": draft_n_months, "options": intent_options,
            # RAW selection criteria, never resolved ids -- same reason
            # each option's own "selection" stores matched_ids/unmatched
            # rather than trusting them to still mean anything later:
            # apply_draft_plan_intent_to_new_groups re-runs this against
            # whatever groups a LATER import produces.
            "group_entities": raw_group_entities,
        }

    # FLOW_REWORK_PLAN.md Phase 4b: the broadcast exemption note used to fire
    # here, at draft time, gated on the drafted agency_involved field -- that
    # field no longer exists (the gross-up is a rep-only checkbox the model
    # never touches), and the live grid caption near the checkbox itself
    # (see BROADCAST_EXCLUDED_FROM_AGENCY_MARKUP's call site in main()) says
    # the same thing whenever a schedule exists, so there's nothing left for
    # a draft-time note to add.

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
        # Move the per-option name/breakout widgets to fresh keys, or the
        # drafted names would be overwritten by whatever the existing widgets
        # already hold (see bump_plan_options_generation).
        bump_plan_options_generation(updates)
        touched_sections.add("media_plan")

        # total_budget_basis == "gross": the checkbox is switched ON, not
        # just flagged. The stored rows are already net (divided above), so
        # leaving the checkbox off would show the rep a number nobody asked
        # for -- the notes stated a gross figure, and unticked the deck
        # reads net. Ticked is the correct starting state for a budget
        # that's already been converted; the rep un-ticks it in one click if
        # that's wrong, same as any other drafted default. Safe to write
        # directly (not through _draft_pending_fields): "agency_gross_up"
        # renders in Section E, beside the plan table near Generate, which
        # is well after this function's own call site in the script (the
        # Draft button lives in the setup band) -- confirmed against
        # BAND_WIDGET_KEYS, which is exactly the set that would force this
        # through the deferred-write queue, and doesn't contain this key.
        if gross_budget_options:
            updates["agency_gross_up"] = True
            for opt_name, gross, net in gross_budget_options:
                internal.append(
                    f'"{opt_name}" states a ${gross:,.0f} gross budget against a net rate, so its '
                    f'plan lines are priced net (${net:,.0f}) and "Apply agency gross-up" was '
                    f"switched on to bring the deck back to ${gross:,.0f}. Untick it if that's not "
                    f"what's intended.")

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
        updates["_shared_fields_key"] = (default_targeting + "||" + geo_or_market
                                         + "||" + flight_label + "||" + flight_shorthand)

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

    # Any key this function is about to write that's ALSO a band widget
    # (BAND_WIDGET_KEYS -- see its own comment) can't be written here: that
    # widget already instantiated earlier in this same run. Deferred into
    # the SAME queue client_name's own seed-if-empty write already uses
    # (`_draft_pending_fields` / apply_pending_draft_fields, applied at the
    # top of the next run) -- not a second mechanism, the same one, so a
    # future band addition is caught here automatically instead of needing
    # its own one-off fix the next time this crashes on a different field.
    # Found live: total_tv (PRODUCT_TO_WIDGET_KEYS' "broadcast_tv" entry, plus
    # the explicit block above) writes here on every draft with a valid
    # plan, and used to raise exactly like client_name once did -- silently
    # aborting this loop partway through, which is also why a drafted
    # plan_options/timing_text write (inserted into `updates` AFTER
    # total_tv) could go missing even though nothing was wrong with either
    # of them on their own.
    pending_band_fields = {}
    for key, value in updates.items():
        if DRAFT_KEY_SECTIONS.get(key) in skip_sections:
            continue
        if key in BAND_WIDGET_KEYS:
            pending_band_fields[key] = value
            continue
        st.session_state[key] = value
    if pending_band_fields:
        existing_pending = dict(st.session_state.get("_draft_pending_fields") or {})
        existing_pending.update(pending_band_fields)
        st.session_state["_draft_pending_fields"] = existing_pending


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


def resolve_drafted_lines(lines_in, total_budget, flight_label, geo_or_market, default_targeting,
                          avails_by_name=None, n_months=1, avails_by_group_id=None):
    """Turn one option's worth of drafted media_plan_lines into real media
    plan rows. Returns (rows, touched_products, touched_sports, unresolved, drivers).

    No arithmetic comes from the model -- it returns intent, this resolves it
    as a waterfall: flat_amount and percent_of_total lines come out of
    total_budget first, percent_of_remainder lines then take their share of
    what's left, and split_evenly lines divide whatever remains evenly among
    themselves.

    `avails_by_group_id` and a line's own `_group_id`/`_geo` are all optional
    and exist only for `_allocate_group_rows` (a targeting group's plan
    line, synthesized as one of these "lines" so it goes through the same
    waterfall a drafted line does): `avails_by_group_id`, when the line
    carries `_group_id`, is consulted BEFORE `avails_by_name` for
    `percent_of_avails` -- necessary because several groups can share one
    audience name (Hershey has 6), and a name-keyed lookup would silently
    grab the wrong group's figure. `_geo` overrides the row's Geo with that
    group's own resolved geography instead of the option's single shared
    `geo_or_market` -- several selected groups on one option usually target
    different places. A rate-type row with `_group_id` set gets `_group_ids:
    [that id]` stamped onto it, the same hidden field `seed_media_plan_rows`
    stamps on every other group-backed row. No drafted line the model
    itself writes ever carries `_group_id`/`_geo` -- only
    `_allocate_group_rows` constructs one -- so every existing frozen
    `.draft.json` fixture goes through this function byte-identically to
    before.
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
        group_id = line.get("_group_id")
        if group_id and avails_by_group_id and group_id in avails_by_group_id:
            monthly = avails_by_group_id[group_id]
        else:
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
        resolved_amounts[i] = round(cost_from_impressions(impressions, line_cpm))

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
                            else impressions_from_cost(amount, cpm))
        # `_geo`, like `_group_id`, is only ever set by _allocate_group_rows
        # -- a synthesized group line carries its OWN group's resolved
        # geography rather than the option's single shared geo_or_market,
        # since several selected groups on one option can (and usually do)
        # target different places. No model-written line ever sets it, so
        # this is a no-op for every existing draft.
        row = {
            "Tactic": tactic, "Flight": flight_label, "Geo": line.get("_geo") or geo_or_market,
            # Fall back to the same per-tactic default the form itself uses
            # (fixed copy for Streaming Retargeting / Live Sports rows,
            # otherwise the Campaign Specs Audience line) rather than always
            # reaching for the audience field.
            "Targeting": audience_track or resolve_row_defaults(
                tactic, geo_or_market, default_targeting, flight_label)["Targeting"],
            "Impressions": line_impressions,
            "CPM": cpm,
            "Type": ROW_TYPE_RATE, "Cost": float(amount),
        }
        if line.get("_group_id"):
            row["_group_ids"] = [line["_group_id"]]
        rows.append(row)
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


def _drafted_budget_basis(value, fallback="net"):
    """"net" (the default -- byte-identical to every draft before this field
    existed, and the right answer whenever the notes are silent or genuinely
    ambiguous about which side of the agency markup a budget sits on) or
    "gross", from the model's own exact wording. Only a literal "gross"
    converts anything; anything else, including a value the model didn't
    understand, falls back rather than guessing -- see the prompt's own
    "genuinely ambiguous case stays net" rule."""
    if isinstance(value, str) and value.strip().lower() == "gross":
        return "gross"
    return fallback


def drafted_options(draft):
    """Normalize a draft's plan into a list of {name, total_budget,
    total_budget_basis, breakout, lines}.

    The model may return several named scenarios in "options", or a single
    plan as a bare "media_plan_lines" -- a single-scenario draft stays a
    one-option proposal, which renders with no option label at all.
    """
    top_budget = round(float(draft.get("total_budget") or 0))
    top_basis = _drafted_budget_basis(draft.get("total_budget_basis"))
    top_breakout = _drafted_breakout(draft.get("breakout"))
    top_selection = draft.get("group_selection")
    top_allocation = draft.get("group_allocation")
    top_cpm = draft.get("group_cpm")
    top_reason = draft.get("group_selection_reason") or ""
    raw_options = draft.get("options") or []
    if raw_options:
        return [
            {
                "name": (opt.get("name") or DEFAULT_OPTION_NAMES[min(i, len(DEFAULT_OPTION_NAMES) - 1)]).strip(),
                "total_budget": round(float(opt.get("total_budget") or top_budget or 0)),
                # Same fallback rule as total_budget itself: an option that
                # states its own budget without repeating the basis inherits
                # the top-level one, not "net" outright -- a "Better" option
                # silent on basis shouldn't quietly stop being gross just
                # because "Good" was the one that spelled it out.
                "total_budget_basis": _drafted_budget_basis(
                    opt.get("total_budget_basis"), top_basis),
                "breakout": _drafted_breakout(opt.get("breakout"), top_breakout),
                "lines": opt.get("media_plan_lines") or [],
                # Falls back to the top-level value only when the OPTION
                # itself omits the key entirely -- same rule total_budget
                # already follows. group_selection is never merged field-by-
                # field with the top-level one; an option that names its own
                # (even an empty {}) is authoritative for that option.
                "group_selection": (opt.get("group_selection") if "group_selection" in opt
                                    else top_selection),
                "group_allocation": (opt.get("group_allocation") if "group_allocation" in opt
                                     else top_allocation),
                # Same fallback rule as group_allocation -- a negotiated rate
                # is per-option just like the allocation it prices.
                "group_cpm": (opt.get("group_cpm") if "group_cpm" in opt else top_cpm),
                "group_selection_reason": opt.get("group_selection_reason") or top_reason,
            }
            for i, opt in enumerate(raw_options[:MAX_PLAN_OPTIONS])
        ]
    return [{"name": DEFAULT_OPTION_NAMES[0], "total_budget": top_budget,
             "total_budget_basis": top_basis,
             "breakout": top_breakout, "lines": draft.get("media_plan_lines") or [],
             "group_selection": top_selection, "group_allocation": top_allocation,
             "group_cpm": top_cpm, "group_selection_reason": top_reason}]


def spread_rows_over_months(rows, n_months):
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
        row["Impressions"] = impressions_from_cost(row["Cost"], row["CPM"])
    return rows


def _color_for_new_group(groups, terms, op):
    """The color a brand-new group should start with -- its audience's
    current shared color, when that audience already has one, else the
    next round-robin swatch. Always paired with `color_locked=False`: a
    fresh group always starts ON the cascade; it only leaves if a rep
    later edits its own Color cell directly.

    The round-robin slot is keyed to how many DISTINCT AUDIENCES already
    exist in `groups`, never a group/row position -- that used to be a
    `position` argument every caller passed as a plain row count. On the
    real Hershey & Harrisburg document (live report, 2026-08-23) the second
    audience's first row lands at group index 3, which the OLD code handed
    GROUP_COLORS[3] regardless of what the first audience got -- and on a
    document whose rows happen to interleave differently, that same
    row-position bug hands two audiences GROUP_COLORS[1] and [4]: literal
    orange and red, no color math needed to reproduce it, just a different
    ordering of the same two audiences' rows. Two audiences always deserve
    the two colors furthest apart in the palette, not whatever slots their
    row positions happen to land on.
    """
    audience = tg.audience_label({"terms": terms, "op": op})
    cascade = tg.audience_cascade_color(groups, audience)
    if cascade is not None:
        return cascade
    known_audiences = {tg.audience_label(g) for g in (groups or [])
                       if tg.audience_label(g).strip()}
    return tg.assign_color(len(known_audiences))


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
        # Drop D2's placeholder group (the blank starter row, seeded when no
        # target market was picked -- see avails_rows_for_markets) rather
        # than adding this real one alongside it. The same append-doesn't-
        # replace gap _finish_avails_import had: a rep who never touches
        # target markets and adds their first audience straight from the
        # finder would otherwise keep a permanently blank row in D2 forever.
        groups = [g for g in groups if not g.get("_placeholder")]
        new = tg.new_group([segment], geo_def={"kind": "text", "label": geo_default or ""},
                           color=_color_for_new_group(groups, [segment], None))
        groups = groups + [new]
        st.session_state["_builder_open_group_id"] = new["id"]
        # No automatic plan line -- the avails table is research by default,
        # same as every other group-creating path now (an import, a hand-
        # typed D2 row). A one-time pointer at the checkbox that actually
        # puts it on the plan, so a rep who's used to the old auto-add isn't
        # left wondering where the line went.
        st.session_state["_builder_note"] = (
            "Added to the avails table. Tick **Plan** in D2 to put it on the media plan.")
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


def _avails_import_field(key, new_value, default_value):
    """Apply-or-conflict for one header field, following the precedence
    every importer in this app already uses: rep edits > this document >
    a draft > defaults. "Safe to apply" means the widget still shows either
    the untouched default or the value THIS SAME IMPORTER last wrote (a
    re-import of the same or a corrected document); anything else means a
    rep (or a draft) has since set something of their own, which an avails
    document -- a real, external, but SECOND-HAND source for these
    particular fields -- must never silently overwrite.

    Returns ("apply", value) or ("conflict", current_value, new_value).
    A field counts as stated when `new_value` is truthy -- right for every
    field this is called on today, all strings/dates. (FLOW_REWORK_PLAN.md
    Phase 4b retired this function's one boolean field, agency_involved,
    along with the `is_stated` override that field alone needed -- a
    document correctly saying "Direct - No Agency" used to be a real,
    meaningful False that a bare `if not new_value` would otherwise have
    read as "nothing to say." If a future boolean field needs the same
    treatment, bring the override back rather than reaching for a truthy
    check on a value where False is a real answer.)
    """
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


_CALM_GEO_NOTE_MARKERS = ("have no county on file", "no market is on file",
                          "counted in the larger one")


def _actionable_geo_notes(notes):
    """Drop the `zips_to_markets` facts that are correct and expected on
    every real document -- a well-formed zip with no county on file at all;
    one that matched a county but that county has no market on file (a
    separate gap in the data, same "nothing for a rep to act on" bucket);
    or a zip credited to the larger of two markets it spans -- from a note
    list that's about to be shown to a rep. `apply_avails_import` reports
    each, at most once, aggregated over the WHOLE import (see its own zip
    pool), never per group: a real Wilmington County Option group alone
    carries 303 zips, and repeating any of these sentences once per
    targeting group (12 on a real Hershey import) is exactly the noise this
    exists to cut.
    """
    return [n for n in notes if not any(marker in n for marker in _CALM_GEO_NOTE_MARKERS)]


def _resolve_zip_originated_geo(zips_list, group_label):
    """(geo_def, zips, markets, message_lines) for a group whose geography
    was specified in the DOCUMENT as an explicit zip list -- a Zip Option,
    a Radius option with no bracketed origin to re-geocode from, or a
    County Option (Premion's own system resolves a rep's entered zips to
    county names and reports both back; the county summary is real but
    DERIVED and never the resolution input -- see AvailsGroup.county_list
    and avails_pdf_import._parse_block). One shared implementation so all
    three kinds get the identical calm-vs-warning treatment rather than
    three near-copies drifting apart.

    The SAME calm-vs-warning split the geo-definition expander already
    applies to a hand-typed zip list (DECISIONS.md, the 2026-08-20
    zip-messaging fix) -- most of zips_to_markets' `.unresolved` is
    well-formed zips with no county/market on file, already summarized in
    `.notes` ("N zip(s) matched a county but no market is on file"). Only a
    genuinely malformed entry (not even a 5-digit zip) gets its own line
    here; the calm summary itself is filtered out entirely (see
    `_actionable_geo_notes`) -- `apply_avails_import` reports that fact
    once, aggregated across every zip-originated group in the document,
    not once per group.
    """
    zips = geo_resolver.parse_zip_list(",".join(zips_list))
    geo_def = {"kind": "zips", "zips": zips}
    market_result = geo_resolver.zips_to_markets(zips)
    markets = sorted(market_result.resolved.keys())
    malformed = [u for u in market_result.unresolved
                if geo_resolver.normalize_zip(u) is None]
    lines = [f"{group_label}: {n}" for n in _actionable_geo_notes(market_result.notes)]
    lines += [f"{group_label}: {u!r} isn't a valid zip code" for u in malformed]
    return geo_def, zips, markets, lines


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
    unresolved_internal = []
    zip_note_pool = []   # every zip-originated group's zips, for ONE aggregate calm-note pass below
    new_groups = []
    catalog = load_audience_catalog()
    valid_segments = set(catalog["segment"])
    profiles, _ = load_market_profiles()
    existing = st.session_state.get("targeting_groups") or []

    for i, g in enumerate(document.groups):
        # Identifies THIS group, not the document -- every note in this loop
        # used to say "(from <the whole document's filename>)", which is the
        # same long string on all twelve of a real document's groups: pure
        # noise (the report is only ever shown for the ONE document a rep
        # just uploaded, so "which document" was never the missing
        # information), and worse, indistinguishable when two DIFFERENT
        # groups hit the same failure -- a real Hershey import surfaced two
        # unrelated "can't resolve this zip" notes that read as one
        # duplicated line because both only named the document, never which
        # of its 12 groups either one was about. Geo comes before audience:
        # several groups sharing one audience across markets is the normal
        # shape an avails document takes, so the geo option is usually the
        # part that actually distinguishes them.
        group_label = f"{g.geo_name or g.geo_kind} / {g.audience_text}"
        terms, op = tg.terms_from_audience_text(g.audience_text)
        for term in terms:
            if term not in valid_segments:
                unresolved.append(
                    f"\"{term}\" ({group_label}) isn't an exact match in the "
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
                        f"\"{g.geo_name}\" ({group_label}) could be more than one "
                        f"market ({options}) -- resolve it by hand in that group's Markets cell.")
                else:
                    unresolved.append(
                        f"\"{g.geo_name}\" ({group_label}) isn't a recognized "
                        f"market -- resolve it by hand in that group's Markets cell.")
                geo_def, zips, markets, notes = {"kind": "text", "label": g.geo_name}, [], [], []
            else:
                geo_def, zips, markets, notes, geo_unresolved = resolve_group_geography(
                    GEO_MODE_MARKETS, markets_picked=[key])
                unresolved += [f"{group_label}: {n}" for n in _actionable_geo_notes(notes)]
                unresolved += [f"{group_label}: {n}" for n in geo_unresolved]
        elif g.geo_kind == avails_pdf_import.GEO_KIND_COUNTY:
            # NOT GEO_MODE_COUNTIES -- that resolves from COUNTY NAMES,
            # which re-derives every zip in every named county (Wilmington:
            # 335) rather than the rep's actual, smaller, authoritative zip
            # list (303) -- a real over-targeting bug against a signed
            # plan, found the same day as this fix. `g.county_list` is
            # Premion's own derived summary -- parsed and kept on the
            # AvailsGroup for whoever reads the document directly, but
            # never fed into resolution; `g.zips` is the real geography,
            # and the group's own name below still comes from the
            # option's OWN label (g.geo_name, e.g. "New Jersey PA and
            # Delaware Counties"), same as a Zip Option already gets --
            # not the raw county list either.
            geo_def, zips, markets, lines = _resolve_zip_originated_geo(g.zips, group_label)
            unresolved += lines
            zip_note_pool.extend(zips)
        elif g.geo_kind == avails_pdf_import.GEO_KIND_RADIUS and g.radius_origin:
            # A real origin exists -- resolve through the SAME Radius mode a
            # rep would use by hand, so this group is byte-identical to one
            # built that way (geo_targeting_roadmap.md's own fixture for
            # this exact document does this, and is what "reaches the same
            # state as entering the document by hand" is checked against).
            geo_def, zips, markets, notes, geo_unresolved = resolve_group_geography(
                GEO_MODE_RADIUS, radius_centers_text=g.radius_origin, radius_miles=g.radius_miles)
            unresolved += [f"{group_label}: {n}" for n in _actionable_geo_notes(notes)]
            unresolved += [f"{group_label}: {n}" for n in geo_unresolved]
            zip_note_pool.extend(zips)
        else:
            # Named zip option, or a radius with NO bracketed origin (one
            # real sample has exactly this) -- there's no center to resolve
            # a radius from, so this is the same fallback a rep is forced
            # into by hand: the document's own zip list, entered as Zips.
            # Same shared helper as County Option above -- see its
            # docstring for the calm-vs-warning messaging this gets.
            geo_def, zips, markets, lines = _resolve_zip_originated_geo(g.zips, group_label)
            unresolved += lines
            zip_note_pool.extend(zips)

        # The document's own name for a Zip/County/Radius option ("Philly Zip
        # Add-On") is what a rep recognizes the buy by -- a market-plus-count
        # summary derived AFTER resolving loses that entirely, and is why
        # five options against one audience rendered as five identical
        # geo-expander headers. A bare DMA has no such name to mirror (its
        # geo_name is just the market name itself, which the derived summary
        # already produces via label_for's real display name); leaving it
        # unset there keeps that path's nicer, resolved-market casing.
        imported_name = (g.geo_name if g.geo_kind != avails_pdf_import.GEO_KIND_DMA else "")
        # Several geographies against one audience is exactly what an avails
        # document does (Hershey: one audience, four DMAs plus a zip add-on)
        # -- they cascade onto ONE shared color by default, same as any other
        # audience, rather than each claiming its own round-robin swatch.
        group = tg.new_group(terms, op=op, geo_def=geo_def, name=imported_name,
                             avails_monthly=0,  # set below, once the flight (and so n_months) is known
                             color=_color_for_new_group(existing + new_groups, terms, op))
        group["resolved_zips"] = zips
        group["resolved_markets"] = markets
        group["_avails_import_impressions"] = g.impressions   # full-flight; see caller
        new_groups.append(group)

    # The calm zips_to_markets facts, ONE aggregate pass over every
    # zip-originated group's zips combined -- not the sum of each group's
    # own (differently-worded, differently-counted) sentence. A real
    # Wilmington import puts 303 zips across its groups through this; the
    # union is deduplicated by construction (`set`), so a zip appearing in
    # more than one group is still counted once.
    if zip_note_pool:
        aggregate = geo_resolver.zips_to_markets(sorted(set(zip_note_pool)))
        unresolved_internal.extend(
            n for n in aggregate.notes
            if any(marker in n for marker in _CALM_GEO_NOTE_MARKERS))

    # FLOW_REWORK_PLAN.md Phase 3: deterministic entity inference -- provable
    # containment only (avails_pdf_import.infer_entities' own module comment
    # has the two rules and what's verified against real documents). This
    # assigns entity_id ONLY -- it never touches row count, _group_ids or
    # include_in_plan; row count stays 1:1 with document.groups regardless.
    # entity_locked stays False (a rep or a later draft can still adjust
    # freely) -- only a rep's own rename/group action ever locks. Ambiguous
    # cases (neither rule fires, e.g. Wilmington's undergrad audiences)
    # route to unresolved_internal as ONE aggregate note, never one per row.
    entity_keys, entity_notes = avails_pdf_import.infer_entities(document)
    unresolved_internal.extend(entity_notes)
    fresh_entity_ids = {}
    for group, key in zip(new_groups, entity_keys):
        if key is None or group.get("entity_locked"):
            continue
        if key not in fresh_entity_ids:
            fresh_entity_ids[key] = uuid.uuid4().hex[:8]
        group["entity_id"] = fresh_entity_ids[key]

    # Deterministic entity LABELING, a separate and narrower pass -- only
    # for a group that stayed its own, singleton entity above (a merged
    # group's label is a different, pre-existing, unrelated gap -- see
    # infer_entity_labels' own module comment). LiveWell is why this
    # exists: seven locations sharing one audience, told apart only by a
    # real street address in the document's own geo_name, which used to
    # reach the D2 grid as a blank Label -- the Targeting column read
    # identically seven times, and a client could only tell rows apart by
    # reading zip lists. Never guessed past what the document's own text
    # supports: infer_entity_labels already drops anything that would
    # collide (three different DC-area radii all reducing to "Washington"
    # stay blank, not arbitrarily disambiguated) -- so nothing here needs
    # to re-check that. A suggestion, not a lock: entity_locked stays
    # False, same as the entity_id assignment just above.
    entity_labels = avails_pdf_import.infer_entity_labels(document)
    for group, key, label in zip(new_groups, entity_keys, entity_labels):
        if key is not None or not label or group.get("entity_locked"):
            continue
        group["entity_label"] = label

    field_updates = {}
    conflicts = []
    for key, value, default in (
        ("client_name", document.advertiser, "Acme Test Co"),
        ("flight_start", document.flight_start, DEFAULT_FLIGHT_START),
        ("flight_end", document.flight_end, DEFAULT_FLIGHT_END),
    ):
        outcome, *rest = _avails_import_field(key, value, default)
        if outcome == "apply":
            field_updates[key] = rest[0]
        elif outcome == "conflict":
            conflicts.append(
                f"{key.replace('_', ' ')} is already set to {rest[0]!r}; the avails document "
                f"says {rest[1]!r}. Left as-is -- update it by hand if the document is right.")

    # FLOW_REWORK_PLAN.md Phase 4b: the document's Agency field used to
    # auto-tick the (now-retired) per-row agency toggle. The new gross-up
    # checkbox is a rep-only calculator with no auto-tick from anywhere --
    # the model doesn't set it and neither does an import -- so a document
    # naming an agency is surfaced as a note instead, for the rep to act on
    # with the checkbox themselves.
    if document.agency and "no agency" not in document.agency.lower():
        unresolved_internal.append(
            f"The avails document names an agency ({document.agency!r}). Tick "
            f"\"Apply agency gross-up (×1.15)\" by the plan if a commission applies "
            f"to this buy.")

    # The document's Attribution field is matched against
    # _AVAILS_ATTRIBUTION_MAP as a genuine input (it can turn a Section D
    # toggle on), but is never reported on its own -- Section D's toggles
    # are the rep's call either way, and a raw restatement of the
    # document's Attribution text added a report line with nothing for a
    # rep to act on.
    attribution_lower = document.attribution_text.lower()
    for phrase, toggle_key in _AVAILS_ATTRIBUTION_MAP.items():
        if phrase in attribution_lower and not st.session_state.get(toggle_key):
            field_updates[toggle_key] = True

    return new_groups, {
        "unresolved": unresolved, "unresolved_internal": unresolved_internal,
        "field_updates": field_updates, "conflicts": conflicts,
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


def apply_pending_flight_match_avails():
    """Apply a band "Match avails flighting" click queued last run, before
    ANY widget on this page is instantiated -- same reason and same shape as
    apply_pending_avails_import_fields: flight_start/flight_end are widgets
    already instantiated once the band renders, so writing them from inside
    the band's own button handler raises.

    This moves the PLAN's flight to the avails DOCUMENT's own window -- the
    opposite direction from D2's per-group "Adjust to plan dates" button,
    which re-scopes one avails GROUP to the plan's dates instead
    (FLOW_REWORK_PLAN.md Phase 4a, (D)). Clearing flight_months makes the
    per-month ranges re-derive as plain defaults against the new bounds
    rather than carrying over stale customization from the old flight, and
    the generation bump is what gives the per-month widgets fresh keys for
    that re-derivation -- same rule flight_month_ranges' own reconciliation
    already depends on elsewhere.
    """
    pending = st.session_state.pop("_pending_flight_match_avails", None)
    if not pending:
        return
    st.session_state["flight_start"] = _parse_iso_date(pending["start"])
    st.session_state["flight_end"] = _parse_iso_date(pending["end"])
    st.session_state.pop("flight_months", None)
    bump_flight_months_generation()


def apply_pending_draft_fields():
    """Apply a band-input seed value drafting queued last run (client_name
    today, the only one that needs this -- see apply_draft_to_form), before
    ANY widget on this page is instantiated. Same reason and same shape as
    apply_pending_avails_import_fields: client_name renders in the setup
    band, above the Draft button, so writing it from inside the button's own
    handler raises -- the exact bug this queue exists to avoid, and the
    reorder alone doesn't fix it, since Draft-from-notes is now the LAST
    thing in the band and still sits below the widget it would be writing.
    """
    pending = st.session_state.pop("_draft_pending_fields", None)
    if not pending:
        return
    for key, value in pending.items():
        st.session_state[key] = value


def apply_pending_flight_mirror_edit():
    """Apply an edit made through the flight control mirrored near the media
    plan (Section E), queued the same "queue now, apply before ANY widget
    renders this run" shape as apply_pending_flight_match_avails -- for the
    identical reason: flight_start/flight_end/custom_flighting are all setup
    -band widget keys that have already instantiated by the time Section E
    renders, so writing any of them from there directly raises. One owner
    (the band's own session_state keys), two render sites -- this is what
    keeps them in sync in both directions without a second copy of the
    flight ever existing. `flight_months` itself never rides in this queue
    -- it isn't a widget key, so the mirror writes it directly, the same way
    the band's own per-month block already does.
    """
    pending = st.session_state.pop("_pending_flight_mirror_edit", None)
    if not pending:
        return
    for key in ("flight_start", "flight_end", "custom_flighting"):
        if key in pending:
            st.session_state[key] = pending[key]


def apply_pending_color_cascade():
    """Push a color queued by the D2 avails table's "apply to audience"
    button (see its expander, right below the grid) onto EVERY group
    sharing that audience -- INCLUDING one already broken out
    (`color_locked`), re-attaching it to the cascade.

    That inclusion is deliberate, not an oversight the lock flag would
    otherwise prevent: a single-row Color edit detaches permanently by
    design (nothing else in this file un-detaches a row on its own), so
    this button is the ONLY way back from an accidental one. Re-attaching
    means clearing the flag too, not just matching the color -- a row
    holding the audience's own current color while still marked
    `color_locked` would keep showing the D2 grid's own "Detached" marker
    for no reason a rep could see, and would still refuse the NEXT
    cascade push the moment the audience's color changes again.

    Queued rather than applied on the click itself, and applied here
    before the grid below reads `targeting_groups` this run -- the same
    "queue now, apply before anything downstream reads it" shape
    `apply_pending_avails_import_fields` uses, for the same reason: the
    button lives inside the same rerun that would otherwise read the OLD
    color.
    """
    pending = st.session_state.pop("_pending_color_cascade", None)
    if not pending:
        return
    audience, color = pending
    groups = st.session_state.get("targeting_groups") or []
    updated = []
    changed = False
    for group in groups:
        if tg.audience_label(group) == audience:
            if group.get("color") != color or group.get("color_locked"):
                group = dict(group)
                group["color"] = color
                group["color_locked"] = False
                changed = True
        updated.append(group)
    if changed:
        st.session_state["targeting_groups"] = updated


def apply_pending_entity_group():
    """Pop the D2 "Group selected rows into one entity" queue (a list of
    group ids, from the grouping expander below the grid) and assign them
    all one fresh, shared `entity_id`, locked -- applied before the grid
    below reads `targeting_groups` this run, the same "queue now, apply
    before anything downstream reads it" shape `apply_pending_color_cascade`
    uses.

    FLOW_REWORK_PLAN.md Phase 3: a rep's own explicit action, never
    automatic inference -- grouping/ungrouping is deliberately NOT
    inferred from two rows' Label text happening to match (see
    targeting_groups.py's module docstring). Whichever selected group
    already carries a non-blank `entity_label` wins it for the whole new
    shared entity (first one found, in the order given); if none does, the
    new entity stays unlabeled until a rep types one into the D2 Label
    column.
    """
    pending = st.session_state.pop("_pending_entity_group", None)
    if not pending:
        return
    gids = set(pending)
    groups = st.session_state.get("targeting_groups") or []
    shared_id = uuid.uuid4().hex[:8]
    shared_label = ""
    for group in groups:
        if group["id"] in gids and tg.entity_label_of(group):
            shared_label = tg.entity_label_of(group)
            break
    updated = []
    changed = False
    for group in groups:
        if group["id"] in gids:
            group = dict(group)
            group["entity_id"] = shared_id
            group["entity_label"] = shared_label
            group["entity_locked"] = True
            changed = True
        updated.append(group)
    if changed:
        st.session_state["targeting_groups"] = updated


def apply_pending_entity_ungroup():
    """Pop the D2 "Ungroup" queue (one entity_id) and hand every group
    currently sharing it back its OWN id as its entity_id -- the exact
    inverse of `apply_pending_entity_group`. Locked, same as grouping,
    since ungrouping is just as much a deliberate rep action as grouping
    was -- nothing automatic (a later re-import, a re-draft) may regroup
    these again on its own."""
    entity_id = st.session_state.pop("_pending_entity_ungroup", None)
    if not entity_id:
        return
    groups = st.session_state.get("targeting_groups") or []
    updated = []
    changed = False
    for group in groups:
        if tg.entity_id_of(group) == entity_id:
            group = dict(group)
            group["entity_id"] = group["id"]
            group["entity_locked"] = True
            changed = True
        updated.append(group)
    if changed:
        st.session_state["targeting_groups"] = updated


def apply_pending_group_include_all():
    """Pop the D2 "Add all to plan" queue and tick every real group, locked
    -- applied before the grid below reads `targeting_groups` this run, the
    same "queue now, apply before anything downstream reads it" shape
    `apply_pending_color_cascade`/`apply_pending_avails_import_fields` use.
    """
    if not st.session_state.pop("_pending_group_include_all", False):
        return
    groups = st.session_state.get("targeting_groups") or []
    updated = [dict(g, include_in_plan=True, include_locked=True) if g.get("id") else g
              for g in groups]
    st.session_state["targeting_groups"] = updated
    st.session_state["_pending_group_realloc"] = True


def apply_pending_group_include_clear():
    """Pop the D2 "Clear plan lines" queue. A group whose plan line has
    been hand-edited is never cleared silently -- it's queued into
    `_pending_include_removal` for the confirmation panel instead (the same
    mechanism a single unchecked row uses, not a second one), and every
    other included group is cleared immediately.
    """
    if not st.session_state.pop("_pending_group_include_clear", False):
        return
    groups = st.session_state.get("targeting_groups") or []
    plan_options = st.session_state.get("plan_options") or []
    blocked, updated = {}, []
    for group in groups:
        if group.get("id") and group.get("include_in_plan"):
            rows = include_removal_blocked(group["id"], plan_options)
            if rows:
                blocked[group["id"]] = rows
                updated.append(group)
                continue
            group = dict(group, include_in_plan=False, include_locked=True)
        updated.append(group)
    st.session_state["targeting_groups"] = updated
    if blocked:
        existing_pending = st.session_state.get("_pending_include_removal") or {}
        st.session_state["_pending_include_removal"] = {**existing_pending, **blocked}
    st.session_state["_pending_group_realloc"] = True


def apply_pending_include_confirm():
    """Pop the confirmation panel's Remove/Keep queue -- same "queue now,
    apply before anything downstream reads it" shape as
    apply_pending_group_include_all/_clear, rather than mutating state and
    calling st.rerun() directly inside the button's own `if st.button(...)`
    block. Keeping every D2 action on one mechanism isn't just consistency
    for its own sake: a button that reads AND writes `_pending_include_
    removal` in the same script pass -- as the direct version did -- reads
    it BEFORE the D2 fold-back above has run this render and writes it
    AFTER, so the fold-back's own (empty, this render) merge into that key
    was already stale by the time the button's write landed, and the
    now-cleared key came back on the very next render looking untouched.
    """
    action = st.session_state.pop("_pending_include_confirm", None)
    if not action:
        return
    pending_removal = st.session_state.get("_pending_include_removal") or {}
    if action == "remove" and pending_removal:
        updated = []
        for group in st.session_state.get("targeting_groups") or []:
            if group.get("id") in pending_removal:
                group = dict(group, include_in_plan=False, include_locked=True)
            updated.append(group)
        st.session_state["targeting_groups"] = updated
        st.session_state["_confirmed_group_row_removals"] = list(pending_removal.keys())
        st.session_state["_pending_group_realloc"] = True
        # The grid's own checkbox state just changed under the hood -- bump
        # so it remounts under a fresh key instead of a real data_editor
        # showing a cached, now-stale checked box (AppTest's own mock is a
        # plain pass-through and wouldn't catch this, but the real widget
        # would -- the same "a keyed widget's value beats value=" trap
        # every other D2/media-plan mutation in this file bumps a version
        # for).
        st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
    # "keep" (or "remove" with nothing actually pending any more) has
    # nothing else to do -- the group is already still included, since the
    # fold-back refused the uncheck in the first place.
    st.session_state.pop("_pending_include_removal", None)


def match_groups_to_selection(groups, selection):
    """(matched_ids, unmatched_needles, mode) -- resolve a draft's
    "group_selection" (mode/ids/match) against REAL targeting groups (real =
    a group id and at least one audience term; never a `_placeholder`).

    Exact ids first -- only ever meaningful for import-then-draft, where the
    draft actually saw real ids to echo back. Then case-insensitive
    substring matching against the SAME three strings
    `plan_lines_from_groups` already derives per group -- `tg.audience_label`,
    `tg.geo_label`, and the group's own `name` -- so a needle like "Subaru"
    or "10 mile" matches whatever it plausibly describes. `geo_def["kind"]`
    is never read, which is what makes this provably geo-kind-agnostic (the
    same selection logic behaves identically whether a group is DMA-based,
    zip-based or radius-based).

    "mode" governs disclosure routing, not just matching: "all" means the
    notes were silent and every real group is included (a seller check,
    `unresolved_internal`); "named" means the notes named specific audiences
    (a client confirmation, `unresolved`) -- returned even when a `match`
    phrase matched nothing, so the caller can still flag the phrase. An
    empty/missing selection while real groups exist is "all", never "select
    nothing" -- the redraft shallow-merge safety this exists for.
    """
    selection = selection or {}
    real_groups = [g for g in (groups or []) if g.get("id") and g.get("terms")]
    if not real_groups:
        return [], [], "all"

    mode = str(selection.get("mode") or "").strip().lower()
    ids_in = [str(i) for i in (selection.get("ids") or []) if i]
    match_in = [str(m).strip() for m in (selection.get("match") or []) if str(m).strip()]

    if mode == "all" or (mode != "named" and not ids_in and not match_in):
        return [g["id"] for g in real_groups], [], "all"

    real_ids = {g["id"] for g in real_groups}
    matched = {i for i in ids_in if i in real_ids}
    # FLOW_REWORK_PLAN.md Phase 3: entity_label_of joins the same haystack --
    # a match phrase can now also name the entity a row already belongs to
    # (e.g. "Toyota of Annapolis"), not just its audience/geo/name.
    haystacks = {g["id"]: " | ".join(filter(None, [
        tg.audience_label(g), tg.geo_label(g, label_for=_market_display_name), g.get("name", ""),
        tg.entity_label_of(g)])).lower()
        for g in real_groups}
    unmatched = []
    for needle in match_in:
        n = needle.lower()
        hit = [gid for gid, hay in haystacks.items() if n in hay]
        if hit:
            matched.update(hit)
        else:
            unmatched.append(needle)
    # Audience-major-ish -- real_groups' own order, not an alphabetical sort
    # of ids -- so a synthesized line list built from this reads the same
    # order every other group-derived row list does.
    ordered = [g["id"] for g in real_groups if g["id"] in matched]
    return ordered, unmatched, "named"


def apply_draft_group_entities(groups, group_entities):
    """FLOW_REWORK_PLAN.md Phase 3: the Label's "sync point" job -- a draft
    names entities among `groups` via `group_entities`
    (`[{"label": ..., "match": [...] | "ids": [...]}]`), the only mechanism
    that can tie together groups with no shared structural fact a
    deterministic pass (`avails_pdf_import.infer_entities`) could find --
    Wilmington's three differently-audienced undergraduate groups being the
    real case this exists for.

    Resolved per entry through the SAME `match_groups_to_selection` every
    `group_selection` entry already uses (its own haystack now includes
    `entity_label_of`, so a draft can also name an entity by a label
    already on the table). Mutates `groups`' own dicts IN PLACE and returns
    `(unresolved, unresolved_internal)` -- so a caller that wants atomic
    "all updates applied together or not at all" semantics (`apply_draft_to_
    form`) must call this against a WORKING COPY it's already about to
    write back wholesale (`pending_groups`), never against session_state's
    live objects directly.

    Only ever applied to a group that is BOTH un-locked and currently its
    own, singleton entity -- a rep's own grouping/rename and commit 7's
    deterministic inference both outrank a draft's naming; this never
    reshapes either. Never invents a label: an entry with a blank "label"
    is skipped outright.

    Two matched shapes, not one -- LiveWell found the gap in the second.
    Two or more matched ids MERGE (a fresh shared entity_id, same as
    before): rows that are really one real-world thing, told apart only by
    the model recognizing them as such. Exactly one matched id NAMES,
    never merges: the row already IS its own entity by construction
    (entity_id already defaults to its own id) -- only entity_label
    changes. This is the case the original design missed entirely: seven
    LiveWell locations sharing one audience are already correctly
    SEPARATE, and needed naming, not merging, so a client reading the
    Targeting column doesn't see the identical text seven times. Naming
    one row is not "an entity of one" in any structural sense -- nothing
    about entity_id, allocation-by-entity-count, or the map's audience-
    keyed color changes; a label is display text, independent of how many
    rows the entity spans.

    Disclosure follows the existing routing rule: a "match" phrase (the
    notes' own words) is a client-facing claim (`unresolved`); "ids" alone
    (found without a phrase to cite) is a seller-only check
    (`unresolved_internal`).
    """
    if not group_entities:
        return [], []
    buckets = tg.entities_of(groups)
    eligible_ids = {members[0]["id"] for members in buckets.values() if len(members) == 1}
    eligible = [g for g in groups if g["id"] in eligible_ids and not g.get("entity_locked")]
    by_id = {g["id"]: g for g in groups}

    unresolved, internal = [], []
    for entry in (group_entities or []):
        label = str((entry or {}).get("label") or "").strip()
        if not label:
            continue
        match_phrases = [str(m) for m in (entry.get("match") or []) if str(m).strip()]
        selection = {"mode": "named", "ids": entry.get("ids") or [], "match": match_phrases}
        matched_ids, unmatched, _mode = match_groups_to_selection(eligible, selection)
        note_lines = []
        if len(matched_ids) >= 2:
            shared_id = uuid.uuid4().hex[:8]
            for gid in matched_ids:
                by_id[gid]["entity_id"] = shared_id
                by_id[gid]["entity_label"] = label
            names = ", ".join(tg.audience_label(by_id[gid]) for gid in matched_ids)
            note_lines.append(f'Grouped {names} as one entity, "{label}".')
        elif len(matched_ids) == 1:
            # Naming, not merging -- see this function's own docstring.
            gid = matched_ids[0]
            by_id[gid]["entity_label"] = label
            note_lines.append(f'Labeled {tg.audience_label(by_id[gid])} ("{tg.geo_label(by_id[gid])}") "{label}".')
        if unmatched:
            note_lines.append(
                f'Couldn\'t tie "{", ".join(unmatched)}" to any avails row for the '
                f'"{label}" entity -- group them by hand below the avails table if '
                f"they belong together.")
        if note_lines:
            (unresolved if match_phrases else internal).extend(note_lines)
    return unresolved, internal


def apply_draft_plan_intent_to_new_groups(new_groups):
    """The draft-then-import order: if a draft already contributed a
    selection/allocation intent for this proposal before these groups
    existed, apply its selection criteria to them now -- the mechanism that
    makes draft-then-import converge with import-then-draft PROVIDED real
    groups already existed at draft time (drafting always contributes an
    intent then, never invented lines -- see groups_own_plan in
    apply_draft_to_form). A suggestion, never a rep decision:
    `include_locked` stays False, so a rep's own tick afterward still wins
    over it (same as every other pre-selection path).

    A draft that ran with NO groups at all is a narrower, accepted gap, not
    something this function tries to fix: the prompt only asks the model
    about group selection when there's something real to select from, so
    that draft took the old path and invented its own Premion Streaming TV
    line(s) straight from the notes. There is no `draft_plan_intent` to
    apply here, and nothing retroactively reshapes those invented lines to
    match whatever this import just produced -- the normal workflow already
    pulls avails before drafting, so this is a rep working out of order, not
    a case worth a rebuild-and-carry mechanism. The one thing owed here is
    visibility: a plain note, not a silent mismatch.
    """
    intent = st.session_state.get("draft_plan_intent")
    if not intent:
        if new_groups and st.session_state.get("draft_last_json"):
            internal = list(st.session_state.get("draft_unresolved_internal") or [])
            internal.append(
                "The media plan's Premion Streaming TV line(s) came from the drafted notes, not "
                "from this avails table -- the draft ran before anything was imported. Import "
                "the avails PDF before drafting next time to sell straight from it; for this "
                "proposal, tick Plan on the rows that belong on the plan and adjust by hand.")
            st.session_state["draft_unresolved_internal"] = list(dict.fromkeys(internal))
        return new_groups
    if not new_groups:
        return new_groups
    matched_ids = set()
    for opt_intent in (intent.get("options") or []):
        ids, _unmatched, _mode = match_groups_to_selection(new_groups, opt_intent.get("selection") or {})
        matched_ids.update(ids)

    result = [dict(g, include_in_plan=True) if g["id"] in matched_ids else g for g in new_groups]

    # FLOW_REWORK_PLAN.md Phase 3: the same re-application this function
    # already does for group_selection, for group_entities -- raw
    # selection criteria stored on the intent, re-resolved against
    # whatever groups THIS import actually produced. Runs even when
    # matched_ids is empty (a draft can name entities without also having
    # selected anything for the plan).
    entity_unresolved, entity_internal = apply_draft_group_entities(
        result, intent.get("group_entities") or [])

    internal = list(st.session_state.get("draft_unresolved_internal") or [])
    if matched_ids:
        st.session_state["_pending_group_realloc"] = True
        internal.append(
            f"{len(matched_ids)} of the {len(new_groups)} audience(s) just imported were pre-ticked "
            f"into the plan, from the selection your last draft made. Review before generating.")
    # Both lists land here, not split -- this function is a seller-side,
    # import-time check throughout; it never writes draft_unresolved.
    internal.extend(entity_unresolved)
    internal.extend(entity_internal)
    if internal != (st.session_state.get("draft_unresolved_internal") or []):
        st.session_state["draft_unresolved_internal"] = list(dict.fromkeys(internal))
    return result


def avails_daily_rate(full_flight_impressions, avail_start, avail_end):
    """The avail document's own per-day rate: impressions divided by the
    number of days in the DOCUMENT'S OWN stated flight -- the portable unit
    an avails figure reduces to, since the plan line those avails feed may
    end up running a flight WIDER, NARROWER or otherwise different from the
    specific document that pulled the number (a rep sets a longer campaign
    flight than the one avail pull covers, or later edits the flight dates).
    Falls back to treating the whole figure as already a one-day rate when
    the document's own dates are missing -- a date-parse gap elsewhere, not
    a reason to guess at a day count.
    """
    if not (isinstance(avail_start, date) and isinstance(avail_end, date)):
        return full_flight_impressions
    days = (avail_end - avail_start).days + 1
    return full_flight_impressions / max(1, days)


def avails_monthly_from_daily_rate(daily_rate, plan_start, plan_end):
    """(avails_monthly, n_months) for the flight the PLAN LINE actually
    runs -- never the avail document's own flight, which may not be the
    same one. Full-flight avails is the daily rate times the plan flight's
    own day count (a wider plan flight means more total avails, not the
    same total spread thinner); avails_monthly is that same full-flight
    total divided by the calendar months the PLAN flight touches, which is
    exactly what avails_to_display's `monthly * n_months` convention
    already expects everywhere else avails are shown or multiplied out.

    When the plan flight and the document's own flight are IDENTICAL --
    the common case -- this reduces exactly to `full_flight / n_months`,
    the same number the previous fix already produced; it only diverges
    (correctly) once the two flights differ. See
    tests/test_avails_import_proration.py, which asserts both.
    """
    if not (isinstance(plan_start, date) and isinstance(plan_end, date)):
        return int(round(daily_rate)), 1
    plan_days = (plan_end - plan_start).days + 1
    n_months = max(1, len(month_list(plan_start, plan_end)))
    full_flight = daily_rate * max(1, plan_days)
    return int(round(full_flight / n_months)), n_months


def avails_full_flight_from_daily_rate(daily_rate, plan_start, plan_end):
    """The EXACT full-flight avails total for the flight the plan line
    actually runs -- daily rate times that flight's own day count, computed
    directly rather than via `avails_monthly * n_months`.

    That indirect route cannot be exact whenever a flight spans months of
    different lengths: `avails_monthly` is one rounded integer reused for
    every month, so multiplying it back out drifts from the true total by
    up to a few impressions (a 4-month flight rounds its monthly figure
    once and compounds that rounding four times). A document broken into
    several real monthly exception rows (Capital Media RFPID-266994: four
    rows, 9/21-9/30 + Oct + Nov + 12/1-12/20, summing to 3,321,409) is
    exactly the shape where this matters -- the SOV denominator on the
    slide has to tie to the document's own stated total, not land a few
    impressions off because September and December are partial months.

    Stored alongside `avails_monthly` (never replacing it -- the D2 table's
    own Monthly/Full-Flight toggle still reads avails_monthly, unchanged);
    `matched_avails_full_flight_for_row` is the one consumer that prefers
    this exact figure over the approximate multiply-out, for the same
    reason avails_monthly itself is frozen at import time rather than
    live-recomputed: a later flight edit can make it stale, and that's the
    same accepted tradeoff avails_monthly already makes.
    """
    if not (isinstance(plan_start, date) and isinstance(plan_end, date)):
        return int(round(daily_rate))
    plan_days = (plan_end - plan_start).days + 1
    return int(round(daily_rate * max(1, plan_days)))


def matched_avails_full_flight_for_row(row, groups_by_id, row_months):
    """The EXACT full-flight avails a plan line's own targeting backs, or
    None when it has no match -- the full-flight counterpart of
    `matched_avails_for_row`, preferring each matched group's own
    `avails_full_flight` (see `avails_full_flight_from_daily_rate`) over
    `avails_monthly * row_months`, which rounds. A group with no stored
    `avails_full_flight` (hand-typed or drafted avails, which never had a
    document flight to derive a day-precise total from) falls back to the
    approximate multiply-out for its own share.

    FLOW_REWORK_PLAN.md Phase 3: ids are bucketed by `tg.entity_id_of` first
    -- MAX within a bucket (two groups of the same real-world entity,
    e.g. a 5mi and a 10mi radius tier, overlap households; summing them
    invents reach that was never there), **sum** across buckets (a line
    merged from Denver + Atlanta, two different entities, still reports
    their combined avails). Entity grouping alone never causes this --
    `_group_ids` only ever holds more than one id here because a rep's own
    `merge_plan_rows` put them there; see that function and
    targeting_groups.py's own module docstring.
    """
    ids = [gid for gid in group_ids_of(row) if gid in groups_by_id]
    if not ids:
        return None

    def full_flight_value(group):
        exact = group.get("avails_full_flight")
        if exact is not None:
            return int(exact)
        return int(group.get("avails_monthly") or 0) * row_months

    buckets = tg.entities_of([groups_by_id[gid] for gid in ids])
    return sum(max(full_flight_value(g) for g in bucket) for bucket in buckets.values())


def _daterange_text(start, end):
    """'Dec 1-14', or 'Nov 28 - Dec 3' across a month boundary -- the plain-
    language form both the divergence notice and the D2 Avail dates column
    use, so a rep reads the same shape in both places."""
    if start.year == end.year and start.month == end.month:
        return f"{start.strftime('%b')} {start.day}-{end.day}"
    return f"{start.strftime('%b')} {start.day} - {end.strftime('%b')} {end.day}"


def avails_flight_divergence(group, flight_ranges):
    """Per-month mismatches between what a group's avails DOCUMENT actually
    covers and the plan's own active per-month ranges for the SAME calendar
    month -- FLOW_REWORK_PLAN.md Phase 2's "hold the quote, name the
    divergence" rule, reversing this app's earlier behaviour of silently
    re-prorating an uploaded avail against whatever flight happened to be
    in effect.

    The document's own implied per-month ranges are derived the same way
    the plan's own are -- `flight_month_ranges(doc_start, doc_end)`,
    clipping the document's overall window to each calendar month it
    touches -- rather than storing the avails-PDF parser's own per-row
    `AvailsPeriod` list on the group. A real document's periods ARE exactly
    this clipping (confirmed against Capital Media's four rows: 9/21-9/30,
    Oct, Nov, 12/1-12/20 -- precisely what clipping 9/21-12/20 to each
    calendar month produces), so this is lossless for every document on
    hand, and avoids keeping a second, nested per-group data structure in
    `targeting_groups` at all. (A raw list-of-dicts value nested inside a
    `targeting_groups` entry was tried and found, live against the real
    12-group Hershey document, to break Streamlit's OWN widget-identity
    bookkeeping on a later rerun -- reproducible with no divergence UI, no
    "Avail dates" column and no plan action involved at all. This
    derivation sidesteps the whole class of problem rather than working
    around it.)

    Returns a list of (month_label, plan_start, plan_end, doc_start,
    doc_end) tuples, one per month that exists in BOTH the plan's active
    ranges and the document's own implied months and where the two
    disagree. A group with no stored document window (hand-typed avails)
    always returns []. Purely informational -- never applied automatically;
    see `apply_pending_avails_plan_adjust` for the rep-triggered, reversible
    action this makes possible.
    """
    doc_start = _parse_iso_date(group.get("avails_doc_start"))
    doc_end = _parse_iso_date(group.get("avails_doc_end"))
    if not (doc_start and doc_end):
        return []
    doc_ranges = flight_month_ranges(doc_start, doc_end)
    by_month = {r["month"]: (r["start"], r["end"]) for r in doc_ranges}
    mismatches = []
    for entry in (flight_ranges or []):
        if not entry.get("active"):
            continue
        doc_range = by_month.get(entry.get("month"))
        if doc_range is None:
            continue
        if (entry.get("start"), entry.get("end")) != doc_range:
            mismatches.append((entry["month"], entry["start"], entry["end"],
                               doc_range[0], doc_range[1]))
    return mismatches


def format_avails_divergence_note(group_label, mismatches):
    """One plain-language sentence naming every month a group's avails
    document diverges from the plan's own dates -- 'Dec 1-14' vs 'Dec
    1-20', the shape a rep can act on directly. None when there's nothing
    to say.
    """
    if not mismatches:
        return None
    parts = [f"{month}: plan runs {_daterange_text(p_start, p_end)}, the avails document "
            f"covers {_daterange_text(d_start, d_end)}"
            for month, p_start, p_end, d_start, d_end in mismatches]
    return f"**{group_label}** -- " + "; ".join(parts) + "."


def _finish_avails_import(document, new_groups, report):
    """Commit an apply_avails_import() result to session_state: queue the
    header-field updates for next run (see apply_pending_avails_import_fields
    for why this can't write them directly), freeze each new group's avails
    figures to the DOCUMENT's own stated numbers (see
    `_apply_imported_avails_groups`), append the groups, and record the
    report for the uploader to display.

    Any `_placeholder` group already in `targeting_groups` (D2's one blank
    starter row, seeded when no target market was picked -- see
    avails_rows_for_markets/seed_rows_to_groups) is dropped before the real
    imported groups are appended. Appending is exactly the operation that
    doesn't get the "a real pick replaces the placeholder" treatment picking
    a target market gets for free (that path rebuilds `targeting_groups`
    from scratch instead of adding to it) -- without this, the blank row
    survived an import forever, sitting in the D2 table alongside whatever
    was actually imported.

    FLOW_REWORK_PLAN.md Phase 2: runs immediately regardless of whether the
    setup band's flight is set yet. Before this, an avails figure had to be
    reduced against SOME plan flight at import time, so an upload arriving
    before the band's dates existed had to be parked until they did
    (`_avails_import_pending_groups`, `apply_pending_avails_import_groups`
    -- both retired in this commit). Freezing a group's figures to the
    document's OWN dates removes that dependency entirely: the plan's flight
    is never consulted at import, so there is nothing left to wait for.
    """
    written = dict(st.session_state.get("_avails_import_written", {}))
    written.update(report["field_updates"])
    st.session_state["_avails_import_written"] = written
    if report["field_updates"]:
        st.session_state["_avails_import_pending_fields"] = report["field_updates"]

    history = list(st.session_state.get("avails_import_history") or [])
    history.append(report["rfpid"])
    st.session_state["avails_import_history"] = history

    _apply_imported_avails_groups(document, new_groups)

    st.session_state["avails_import_report"] = report
    st.session_state["avails_import_error"] = None


def _apply_imported_avails_groups(document, new_groups):
    """The freeze-and-append half of an avails-PDF import.

    FLOW_REWORK_PLAN.md Phase 2, reversing this app's earlier behaviour: an
    uploaded avail is a real quote from a real document, and it holds at
    that quote until a rep EXPLICITLY asks to re-scope it to the plan's own
    dates (`apply_pending_avails_plan_adjust`) -- never automatically, never
    silently. The plan's flight is not read anywhere in this function.

    `avails_full_flight` is the document's own stated total, verbatim --
    not round-tripped through any rate math, so there is zero rounding risk
    between what the document says and what gets stored. `avails_monthly`
    is that same total divided by the calendar months the DOCUMENT's own
    flight spans (via the existing avails_daily_rate/avails_monthly_from_
    daily_rate reduction, called with the document's dates on BOTH sides --
    reusing the existing math rather than a second formula, per "this
    already exists; it now reads from one place"). The document's own
    `avails_doc_start`/`avails_doc_end` are all `avails_flight_divergence`
    needs -- it derives the document's implied per-month ranges the same
    way the plan's own are derived, rather than this function storing the
    avails-PDF parser's per-row `AvailsPeriod` list on the group (tried and
    found, live against the real Hershey document, to break Streamlit's own
    widget-identity bookkeeping on a later rerun -- see
    `avails_flight_divergence`'s own docstring).

    `avails_doc_monthly`/`avails_doc_full_flight` duplicate the frozen
    figures under their own names so `apply_pending_avails_plan_adjust`'s
    "Use document figure" action can restore them exactly even after an
    adjustment has overwritten `avails_monthly`/`avails_full_flight`.
    """
    doc_start, doc_end = document.flight_start, document.flight_end
    for group in new_groups:
        full_flight = group.pop("_avails_import_impressions", 0)
        if isinstance(doc_start, date) and isinstance(doc_end, date):
            daily_rate = avails_daily_rate(full_flight, doc_start, doc_end)
            monthly, _ = avails_monthly_from_daily_rate(daily_rate, doc_start, doc_end)
        else:
            monthly = full_flight
        group["avails_monthly"] = monthly
        group["avails_full_flight"] = full_flight
        group["avails_doc_monthly"] = monthly
        group["avails_doc_full_flight"] = full_flight
        group["avails_adjusted_to_plan"] = False
        group["avails_doc_start"] = str(doc_start) if doc_start else None
        group["avails_doc_end"] = str(doc_end) if doc_end else None

    existing = [g for g in (st.session_state.get("targeting_groups") or [])
               if not g.get("_placeholder")]
    # No automatic plan line -- an import populates the avails table only;
    # every new group arrives include_in_plan=False (new_group's own
    # default), real inventory a rep can show without it silently becoming
    # a billed line. If a draft already contributed a selection/allocation
    # intent for this proposal (the draft-then-import order), apply it to
    # these newly-imported groups now -- see apply_draft_plan_intent_to_new_groups.
    new_groups = apply_draft_plan_intent_to_new_groups(new_groups)
    st.session_state["targeting_groups"] = existing + new_groups
    st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1


def apply_pending_avails_plan_adjust():
    """Apply a D2 "Adjust to plan dates" / "Use document figure" click.

    Queued the same way every other D2 action is (CLAUDE.md: queue, then
    apply before D2 reads groups this run) -- FLOW_REWORK_PLAN.md Phase 2's
    reversible, rep's-call escape hatch out of the frozen document figure.
    Called once per run, before D2 reads `targeting_groups`.

    Adjusting re-derives the group's own daily rate from its FROZEN
    `avails_doc_full_flight`/`avails_doc_start`/`avails_doc_end` -- never
    the live `avails_monthly`/`avails_full_flight`, which may already be an
    earlier adjustment; re-deriving from an already-adjusted figure would
    compound -- and reapplies it across the PLAN's own ACTIVE window
    (`flight_active_days`/`active_month_labels`, commit 1's single owner),
    which is day-and-skip-aware in a way a bare start/end pair can't be.
    Reverting restores the frozen document figures exactly, byte-for-byte.
    """
    pending = st.session_state.pop("_pending_avails_plan_adjust", None)
    if not pending:
        return
    groups = st.session_state.get("targeting_groups") or []
    flight_start = st.session_state.get("flight_start")
    flight_end = st.session_state.get("flight_end")
    ranges = (flight_month_ranges(flight_start, flight_end, st.session_state.get("flight_months"))
             if isinstance(flight_start, date) and isinstance(flight_end, date) else [])
    active_days = flight_active_days(ranges)
    active_n_months = max(1, len(active_month_labels(ranges)))
    for group in groups:
        if group["id"] not in pending:
            continue
        if not pending[group["id"]]:
            group["avails_monthly"] = group.get("avails_doc_monthly", group.get("avails_monthly"))
            group["avails_full_flight"] = group.get("avails_doc_full_flight", group.get("avails_full_flight"))
            group["avails_adjusted_to_plan"] = False
            continue
        doc_start = _parse_iso_date(group.get("avails_doc_start"))
        doc_end = _parse_iso_date(group.get("avails_doc_end"))
        doc_full_flight = group.get("avails_doc_full_flight")
        if not (doc_start and doc_end and doc_full_flight and active_days):
            continue
        daily_rate = avails_daily_rate(doc_full_flight, doc_start, doc_end)
        group["avails_monthly"] = int(round(daily_rate * active_days / active_n_months))
        group["avails_full_flight"] = int(round(daily_rate * active_days))
        group["avails_adjusted_to_plan"] = True


def render_logo_upload():
    """The client-logo uploader -- lives in the intake area now (UX sweep,
    BACKLOG.md), not Section A, alongside the other "things a rep has in
    hand." Behavior is unchanged from when this sat in Section A's own
    column; only the call site moved.
    """
    logo_file = st.file_uploader("🖼️ Client logo", type=["png", "jpg", "jpeg"],
                                  key="logo_upload",
                                  help="The client's logo, placed on the proposal's cover slide.")
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


def render_avails_pdf_uploader(key_suffix, prompt):
    """One avails-PDF upload widget, shared verbatim by both entry points
    (the intake area at the top of the page, and beside the D2 avails
    table) -- same import, same precedence, same report, per
    geo_targeting_roadmap.md F's own "second entry point, not a second
    implementation" requirement.

    `key_suffix` keeps the two entry points' widget keys distinct (a rep
    could conceivably want to use either); `prompt` is the copy shown above
    the uploader, which is the only thing that actually differs between them.
    """
    upload = st.file_uploader("📄 Avails PDF (from Salesforce)", type=["pdf"],
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
        # FLOW_REWORK_PLAN.md Phase 2: no more "parsed, pending the flight"
        # state to report -- a group's figures freeze to the document's own
        # dates, never the plan's, so an import always finishes immediately.
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
    # Deliberately NOT `plottable` -- the map wants every resolved group
    # (market, county, zips, radius all draw dots), but a market or county
    # group's zip list has no consumer: ad ops targets those by DMA/county
    # NAME, not a list of however many thousand zips make it up. Exporting
    # that list is just noise, and on a real document it's a lot of noise --
    # on the Hershey & Harrisburg scenario, showing `plottable` here instead
    # would have grown this from 4 entries to 12, 8 of them an unused
    # 200-1,000+ line zip dump. See targeting_map.exportable_zip_groups.
    for group in targeting_map.exportable_zip_groups(groups):
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
        builder_note = st.session_state.pop("_builder_note", None)
        if builder_note:
            st.caption(builder_note)

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


# --- Per-month flight ranges (FLOW_REWORK_PLAN.md Phase 2) ----------------
# The flight's month model gets ONE owner. Before this, "which months does
# the flight run in" was derived twice -- `form_flight_months` and an inline
# copy inside main() -- each re-implementing the same empty/partial-overlap
# rule by hand, and neither could express a CLIPPED month: a flight running
# 9/21-12/14 was modelled as four whole calendar months, so a plan cell said
# "Sep 2026" for ten days of September.
#
# A range entry is {"month": "Sep 2026", "start": date, "end": date,
# "active": bool}. `month` keeps month_list's exact "%b %Y" format, because
# it is the join key for active_months, for _month_sort, and for the Wide
# Orbit coverage warning's set comparison.
#
# The list is stored in an ORDINARY session_state key, never in the widgets
# themselves: Streamlit garbage-collects the session_state of any keyed
# widget not rendered during a run, and these rows are dynamic and live
# inside an expander. Widgets mirror the list and fold back into it only on
# a real edit -- the same discipline the D2 avails grid already uses.
def _month_bounds(label):
    """(first_day, last_day) of the calendar month a "%b %Y" label names,
    or (None, None) when it doesn't parse -- a stored label from a
    different vintage is a reason to fall back, never to raise."""
    try:
        first = datetime.strptime(label, "%b %Y").date()
    except (ValueError, TypeError):
        return None, None
    if first.month == 12:
        next_first = date(first.year + 1, 1, 1)
    else:
        next_first = date(first.year, first.month + 1, 1)
    return first, next_first - timedelta(days=1)


def default_month_range(label, flight_start, flight_end):
    """One month's default entry: the full calendar month, clipped by the
    flight's own bounds, active. The default state of a per-month range is
    therefore identical to having no per-month ranges at all, which is what
    makes turning the control on a no-op until a rep actually edits a row.
    """
    first, last = _month_bounds(label)
    if first is None:
        return {"month": label, "start": flight_start, "end": flight_end, "active": True}
    return {
        "month": label,
        "start": max(first, flight_start) if flight_start else first,
        "end": min(last, flight_end) if flight_end else last,
        "active": True,
    }


def flight_month_ranges(flight_start, flight_end, stored=None):
    """The canonical per-month range list for a flight -- THE single owner.

    Reconciliation against `stored` is deliberately identical in EFFECT to
    the rule the old active-months multiselect used, which CLAUDE.md records
    as an incident rather than a preference: a month that still exists keeps
    its stored range (re-clipped to the flight, so a stale range is CLAMPED
    rather than dropped -- a rep's edit survives a flight nudge); a month
    with no stored entry gets the default; a stored entry whose month has
    left the range is dropped; and if NOTHING survives, the whole thing
    resets to defaults, because an empty intersection means the stored
    choice belongs to a different flight entirely, while a partial overlap
    is real custom flighting and is kept.
    """
    if not (isinstance(flight_start, date) and isinstance(flight_end, date)):
        return []
    if flight_end < flight_start:
        flight_start, flight_end = flight_end, flight_start
    labels = month_list(flight_start, flight_end)
    by_label = {}
    for entry in (stored or []):
        if isinstance(entry, dict) and entry.get("month") in labels:
            by_label[entry["month"]] = entry

    ranges = []
    for label in labels:
        base = default_month_range(label, flight_start, flight_end)
        prior = by_label.get(label)
        if prior is None:
            ranges.append(base)
            continue
        start = prior.get("start") if isinstance(prior.get("start"), date) else base["start"]
        end = prior.get("end") if isinstance(prior.get("end"), date) else base["end"]
        # Clamp into BOTH the flight's bounds and the month's own -- a
        # stored range is re-clipped, never trusted blindly, because the
        # flight it was entered against may since have moved. Clamping
        # before any date_input renders is also what stops Streamlit
        # raising on an out-of-range value.
        start = min(max(start, base["start"]), base["end"])
        end = min(max(end, base["start"]), base["end"])
        if end < start:
            start, end = base["start"], base["end"]
        ranges.append({"month": label, "start": start, "end": end,
                       "active": bool(prior.get("active", True))})

    if not any(r["active"] for r in ranges):
        # Every month deselected is not a state a rep can reach through the
        # control (it refuses the last one), so reaching it means the stored
        # list belongs to another flight.
        return [default_month_range(label, flight_start, flight_end) for label in labels]
    return ranges


def active_month_labels(ranges):
    return [r["month"] for r in (ranges or []) if r.get("active")]


def flight_active_days(ranges):
    """Total inclusive days the flight actually runs, summed over the ACTIVE
    months' own ranges -- the day count an avails daily rate is re-applied
    across when a rep asks for that, and the honest denominator for a flight
    with a clipped first or last month."""
    total = 0
    for entry in (ranges or []):
        if not entry.get("active"):
            continue
        start, end = entry.get("start"), entry.get("end")
        if isinstance(start, date) and isinstance(end, date) and end >= start:
            total += (end - start).days + 1
    return total


def ranges_are_customized(ranges, flight_start, flight_end):
    """True when `ranges` says something the plain band flight doesn't --
    a skipped month, or a month clipped tighter than the flight already
    implies. Derived rather than stored: a `custom_flighting` flag would be
    a second owner for a fact the ranges themselves already carry, and the
    two could disagree."""
    default = flight_month_ranges(flight_start, flight_end)
    if len(default) != len(ranges or []):
        return True
    for base, entry in zip(default, ranges):
        if (entry.get("month") != base["month"] or not entry.get("active")
                or entry.get("start") != base["start"] or entry.get("end") != base["end"]):
            return True
    return False


def flight_months_snapshot(ranges):
    """The per-month ranges as form_json stores them -- ISO strings, so the
    round-trip through JSON is lossless and a stored proposal stays readable
    without a date parser on the other side."""
    out = []
    for entry in (ranges or []):
        start, end = entry.get("start"), entry.get("end")
        out.append({
            "month": entry.get("month"),
            "start": str(start) if isinstance(start, date) else None,
            "end": str(end) if isinstance(end, date) else None,
            "active": bool(entry.get("active", True)),
        })
    return out


def _parse_iso_date(value):
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def resolve_flight_months(form):
    """The per-month ranges for ANY proposal, old or new -- the migration-read
    half of `flight_months_snapshot`, and the exact mirror of what
    `resolve_setup` does for the band's own five values.

    A proposal logged after Phase 2 carries "flight"."month_ranges" and that
    is authoritative, reconciled against its own flight dates.

    A pre-Phase-2 proposal has no such key, only the whole-month
    "active_months" list. Every month then gets its default clipped range
    and `active` comes from that list -- a straight relocation of the read
    already in `rehydrate_proposal_into_form`, not new inference. That is
    what makes a pre-rework proposal reproduce byte-identical output: its
    plan rows still carry their own stored whole-month Flight text, and
    rehydration marks restored rows dirty so nothing re-seeds over them.

    Never raises. A label from a different vintage ("2026-01" rather than
    "Jan 2026", as one real test fixture carries) simply fails to match and
    falls through to the all-active default, the same honesty the empty-
    intersection rule already applies.
    """
    flight = form.get("flight") or {}
    start = _parse_iso_date(flight.get("start"))
    end = _parse_iso_date(flight.get("end"))
    if not (start and end):
        return []

    stored = flight.get("month_ranges")
    if isinstance(stored, list) and stored:
        revived = []
        for entry in stored:
            if not isinstance(entry, dict):
                continue
            revived.append({
                "month": entry.get("month"),
                "start": _parse_iso_date(entry.get("start")),
                "end": _parse_iso_date(entry.get("end")),
                "active": bool(entry.get("active", True)),
            })
        return flight_month_ranges(start, end, revived)

    ranges = flight_month_ranges(start, end)
    stored_active = flight.get("active_months")
    if isinstance(stored_active, list) and stored_active:
        wanted = {str(m) for m in stored_active}
        # Only honour the stored selection when it names months this flight
        # actually has -- otherwise it belongs to a different flight, and
        # the empty-intersection rule says take the whole range.
        if any(r["month"] in wanted for r in ranges):
            for entry in ranges:
                entry["active"] = entry["month"] in wanted
    return ranges


def format_flight_shorthand(ranges):
    """The plan-cell / Campaign-Specs day-precise flight string --
    FLOW_REWORK_PLAN.md Phase 2's single owner, so the plan slide, the
    on-screen grid and Campaign Specs never reimplement this per surface.

    Only ACTIVE months participate; a skipped month's absence breaks a run
    rather than being bridged over. "Full" means the full CALENDAR month,
    not merely "as clipped by the flight" -- the plan doc's own worked
    example (Sep 8-30, Oct-Nov, Dec 1-14) renders September with days even
    though day 8 is where the flight itself starts, which only makes sense
    under a calendar-month reading of "full": a full month renders bare
    ("Oct"), a partial one with its own start/end day ("Sep 8-30"),
    consecutive full months collapse to their endpoints ("Oct-Nov"), and
    runs are comma-joined.

    A run never crosses a calendar-year boundary, even when the months are
    otherwise adjacent (Dec into Jan) -- collapsing across one would leave
    a bare "Jan" with no way to tell which year it names once the string
    already spans more than one. Years are omitted entirely unless the
    active months span more than one calendar year; a bare "Jan" is
    unambiguous within a single-year flight.
    """
    entries = []
    for r in (ranges or []):
        if not r.get("active"):
            continue
        first, last = _month_bounds(r.get("month"))
        if first is None:
            continue
        entries.append({
            "year": first.year, "month_num": first.month,
            "abbrev": r["month"].split(" ")[0],
            "full": (r.get("start") == first and r.get("end") == last),
            "start": r.get("start"), "end": r.get("end"),
        })
    if not entries:
        return ""

    multi_year = len({e["year"] for e in entries}) > 1

    def year_suffix(entry):
        return f" {entry['year']}" if multi_year else ""

    runs = []
    i = 0
    while i < len(entries):
        e = entries[i]
        if not e["full"]:
            start_day = e["start"].day if isinstance(e["start"], date) else "?"
            end_day = e["end"].day if isinstance(e["end"], date) else "?"
            # A single-day "month" (a short broadcast schedule's last real
            # week can land here) renders as one day, not the degenerate
            # "Jun 1-1" a bare start-end range would give it.
            days = str(start_day) if start_day == end_day else f"{start_day}–{end_day}"
            runs.append(f"{e['abbrev']} {days}{year_suffix(e)}")
            i += 1
            continue
        j = i
        while (j + 1 < len(entries) and entries[j + 1]["full"]
               and entries[j + 1]["year"] == entries[j]["year"]
               and entries[j + 1]["month_num"] == entries[j]["month_num"] + 1):
            j += 1
        run = entries[i:j + 1]
        text = run[0]["abbrev"] if len(run) == 1 else f"{run[0]['abbrev']}–{run[-1]['abbrev']}"
        runs.append(text + year_suffix(run[-1]))
        i = j + 1

    return ", ".join(runs)


def bump_flight_months_generation(target=None):
    """Invalidate the per-month range widget keys -- the same trap
    `bump_plan_options_generation` guards against, for the same reason.
    Once `fm_start_0_3` exists holding a September date, moving the flight
    from Sep-Dec to Oct-Jan leaves index 3 -- now January -- rendering a
    September date, and the fold-back would write it right back into the
    January row. Any change to the MONTH SET (not an edit to one month's
    own range, which should NOT bump this) needs fresh keys.
    """
    store = st.session_state if target is None else target
    store["flight_months_gen"] = st.session_state.get("flight_months_gen", 0) + 1


def form_flight_months():
    """(all_months, active_months) as main() will derive them from whatever the
    form currently holds.

    A mirror of main()'s own derivation, for callers that run *before* the
    per-month flight widgets have rendered this run and need to know the
    month count those widgets are going to produce. The month count that
    spreads a budget has to be the month count that totals it, and anything
    computing its own is guessing. Falls back to the widget defaults when
    session_state is still empty, since that is what the widgets themselves
    will fall back to.

    Collapses onto `flight_month_ranges` -- the single owner -- rather than
    re-implementing its own copy of the reconciliation rule, which is what
    this function and main()'s own derivation used to do independently
    before FLOW_REWORK_PLAN.md Phase 2.
    """
    start = st.session_state.get("flight_start") or DEFAULT_FLIGHT_START
    end = st.session_state.get("flight_end") or DEFAULT_FLIGHT_END
    if not (isinstance(start, date) and isinstance(end, date)):
        return [], []
    stored = st.session_state.get("flight_months")
    if stored:
        ranges = flight_month_ranges(start, end, stored)
    else:
        # No per-month ranges rendered yet this session -- reconstruct from
        # whatever the whole-month `active_months` holds (a pre-Phase-2
        # session, or the very first run before Section E has ever
        # written flight_months), the same fallback resolve_flight_months
        # uses for a stored form.
        ranges = flight_month_ranges(start, end)
        active_labels = st.session_state.get("active_months")
        if isinstance(active_labels, list) and active_labels:
            wanted = set(active_labels)
            if any(r["month"] in wanted for r in ranges):
                for r in ranges:
                    r["active"] = r["month"] in wanted
    return month_list(start, end), active_month_labels(ranges)


def is_flat_fee_row(row):
    return str(row.get("Type", ROW_TYPE_RATE)) == ROW_TYPE_FLAT_FEE


# Household-level CTV/streaming inventory a viewer could plausibly co-view --
# confirmed with the user as exactly these two, nothing else. `line_type`
# can't make this distinction (Premion Streaming TV and Streaming Retargeting
# are both "premion") and neither can "has fixed targeting_copy" (every Live
# Sports package has one too, same as Streaming Retargeting) -- an explicit
# allowlist, matched by tactic-name prefix the same way fixed_targeting_copy
# matches (longest label first isn't needed here since neither prefix is a
# prefix of the other, but the match itself must be a prefix, not equality,
# since a drafted or merged label can have text appended).
COVIEWING_ELIGIBLE_TACTIC_PREFIXES = ("Premion Streaming TV", "Live Sports - ")


def is_coviewing_eligible_row(row):
    """True for a Premion Streaming TV or Live Sports line -- never
    Streaming Retargeting, broadcast, AM audio, or a flat fee (which has no
    impressions to project a co-viewing figure from in the first place)."""
    tactic = str(row.get("Tactic", "") or "")
    return any(tactic.startswith(prefix) for prefix in COVIEWING_ELIGIBLE_TACTIC_PREFIXES)


def _num(value):
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def cost_from_impressions(impressions, cpm):
    """Whole dollars, at the NET rate. FLOW_REWORK_PLAN.md Phase 4b: the
    agency gross-up no longer touches this at all -- it's a rep-only,
    order-wide multiplier applied once, at display/deck time, in
    `compute_plan_totals`. This function (and its inverse below) never knew
    about it in the first place."""
    return float(round((_num(impressions) / 1000.0) * _num(cpm)))


def impressions_from_cost(cost, cpm):
    """The exact inverse of cost_from_impressions -- so typing a cost and
    typing the impressions it implies land on the same row."""
    rate = _num(cpm)
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

# The one ratio this app ever grosses by -- named so the checkbox's own
# multiply (compute_plan_totals, via row_markup) and a drafted gross budget's
# divide (resolve_drafted_lines' caller, apply_draft_to_form) can't drift
# apart into two different numbers for what's supposed to be one inverse
# operation. Was a bare 1.15 literal at both sites; still is at a few
# display-only caption strings, which don't compute anything and so aren't
# a drift risk.
AGENCY_MARKUP = 1.15

# **Settled policy, not a default.** A Wide Orbit schedule's cost is already
# gross -- the station quotes it that way and the agency commission is inside
# it, as the Planner PDF's own "Total Cost / Agency Commission @ 15% / Net
# Cost" block spells out. Running the x1.15 agency markup over it would bill
# the commission twice. So the broadcast line uses the WO cost verbatim and
# is excluded from the markup permanently, whatever the agency gross-up
# checkbox says.
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

    # FLOW_REWORK_PLAN.md Phase 2's broadcast exemption: the schedule's OWN
    # dates, exactly as Geo is never re-seeded from the plan's market. Falls
    # back to the plan's own flight text only when the export gave no
    # parseable dates at all (a Planner export headers by position, and
    # Schedule.weeks can come back empty) -- rendered through the same
    # shorthand renderer the plan cell uses, so both Flight columns read in
    # one voice.
    if isinstance(summary.flight_start, date) and isinstance(summary.flight_end, date):
        broadcast_flight = format_flight_shorthand(
            flight_month_ranges(summary.flight_start, summary.flight_end)) or flight_label
    else:
        broadcast_flight = flight_label

    station = (summary.station or "Broadcast").upper()
    row = {
        "Tactic": f"{station} {BROADCAST_TACTIC_SUFFIX}",
        "Flight": broadcast_flight,
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


def recompute_row(row, driver):
    """Recompute whichever side of a rate row isn't driving. Mutates and
    returns the row. Flat-fee rows have neither side -- their Cost is the fee
    itself, and Impressions/CPM are ignored entirely.

    FLOW_REWORK_PLAN.md Phase 4b: every row on the grid is NET now -- the
    agency gross-up stopped being a per-row concern the moment it became a
    rep-only, order-wide multiplier applied once at display/deck time in
    `compute_plan_totals` (via `row_markup`, which still enforces the
    broadcast exemption there). There is nothing for this function to gross
    any more.
    """
    if is_flat_fee_row(row):
        return row
    if driver == DRIVER_COST:
        row["Impressions"] = impressions_from_cost(row.get("Cost"), row.get("CPM"))
    else:
        row["Cost"] = cost_from_impressions(row.get("Impressions"), row.get("CPM"))
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

    Flight is held the same way, for the same reason (FLOW_REWORK_PLAN.md
    Phase 2's broadcast exemption): the imported schedule's Flight is the
    Wide Orbit schedule's own real span, seeded by `broadcast_row_for`, and
    a shared-field edit re-stamping it with the plan's own flight text would
    put a span on the row the schedule never actually ran.
    """
    targeting = fixed_targeting_copy(tactic) or default_targeting
    geo = default_geo
    flight = flight_label
    if current is not None and is_broadcast_row({"Tactic": tactic}):
        targeting = current.get("Targeting") or targeting
        geo = current.get("Geo") or geo
        flight = current.get("Flight") or flight
    return {"Flight": flight, "Geo": geo, "Targeting": targeting}


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


# ---------------------------------------------------------------------------
# Targeting groups own their Premion Streaming TV plan line, exclusively --
# see CLAUDE.md's "The media plan and the form" for the design this
# implements. A group's `include_in_plan` flag (targeting_groups.py) is the
# ONLY thing that decides whether it has a plan line; the avails table
# itself is research, never an implicit seed list. Every other product
# (retargeting, AM, sports, fees, broadcast) is untouched by any of this.
# ---------------------------------------------------------------------------
GROUP_LINE_PRODUCT_KEY = "premion_streaming_tv"


def is_group_line(row):
    """A Premion Streaming TV row carrying at least one targeting-group id --
    the ownership predicate `reconcile_group_plan_lines` uses to know which
    rows on a media plan option it, and only it, may add or remove. Every
    other row -- retargeting, AM, sports, a fee, a broadcast line, or an
    UNGROUPED Premion row (a proposal from before targeting groups existed,
    or a form that has never touched groups at all) -- belongs to someone
    else and is never touched here.
    """
    label, _cpm = line_product_spec(GROUP_LINE_PRODUCT_KEY)
    return bool(group_ids_of(row)) and row_belongs_to_tactic(row.get("Tactic"), label)


def include_removal_blocked(gid, plan_options):
    """[{option, tactic, geo, impressions, cost}, ...] naming every DIRTY
    group-owned row that unchecking `gid` would discard -- empty means the
    uncheck is safe to apply with no interruption. A row carrying several
    ids (a `merge_plan_rows` product) is blocked by `gid` only if `gid` is
    one of them AND the row has been hand-edited; an untouched row is never
    a reason to interrupt.
    """
    blocked = []
    for option in plan_options or []:
        for row, dirty in zip(option.get("rows", []), option.get("dirty", [])):
            if not dirty or not is_group_line(row) or gid not in group_ids_of(row):
                continue
            blocked.append({"option": option.get("name", ""), "tactic": row.get("Tactic", ""),
                            "geo": row.get("Geo", ""), "impressions": row.get("Impressions", 0),
                            "cost": row.get("Cost", 0)})
    return blocked


def selected_group_triples(groups, fallback_geo=""):
    """(audience, geo, group_id) for `include_in_plan` groups only, in the
    SAME audience-major order `plan_lines_from_groups` derives for the full
    set -- filtered down to selection rather than re-implemented, so a
    selected group's geo/audience resolution can never drift from what
    "every group" mode already computes (branch C, above).
    """
    included_ids = {g["id"] for g in (groups or []) if g.get("include_in_plan") and g.get("id")}
    if not included_ids:
        return []
    return [t for t in plan_lines_from_groups(groups, fallback_geo) if t[2] in included_ids]


def _group_alloc_line(group, allocation, cpm):
    """One synthesized media-plan "line" for a targeting group, ready to
    fold into `resolve_drafted_lines`'s waterfall alongside a drafted
    option's other lines -- the shared building block for a fresh draft's
    group-derived lines (`apply_draft_to_form`) and a later re-allocation
    over already-selected ones (`_allocate_group_rows`), so there is
    exactly one place that turns "a group plus an allocation" into
    something the waterfall understands.
    """
    line = {"product": GROUP_LINE_PRODUCT_KEY, "audience_track": tg.audience_label(group),
            "allocation": dict(allocation or {"split_evenly": True}),
            "_group_id": group["id"],
            "_geo": tg.geo_label(group, label_for=_market_display_name).strip()}
    if cpm not in (None, ""):
        line["cpm"] = cpm
    return line


def _synthesize_group_lines(ids, allocation, cpm, groups_by_id):
    """One waterfall line per ENTITY among the given ids (FLOW_REWORK_PLAN.md
    Phase 3), not one per id -- so a stated split ("split evenly across the
    four stores") divides by how many real-world things are selected, never
    by how many avail rows happen to back them (a store with two radius
    tiers both selected is still one share). Ids with no matching group (a
    stale or dropped one) are skipped rather than raising, same as before.

    The representative id for each entity is the FIRST of `ids`, in the
    order given, whose group belongs to that entity -- stable, never
    dict-iteration-order-dependent, and never entity-wide: only ids that
    are actually IN `ids` (the caller's own selected/matched set) are ever
    considered, exactly `tg.entities_of`'s own scoping rule.

    Returns `(lines, avails_by_group_id)` -- the second keys ONLY each
    entity's representative id, to `tg.entity_avails_monthly` over that
    entity's members among the GIVEN ids only (max, not sum, and never
    reaching outside what was actually selected)."""
    ordered_groups = [groups_by_id[gid] for gid in ids if gid in groups_by_id]
    buckets = tg.entities_of(ordered_groups)
    lines, avails_by_group_id = [], {}
    for gid in ids:
        group = groups_by_id.get(gid)
        if group is None:
            continue
        members = buckets.get(tg.entity_id_of(group))
        if members is None or members[0]["id"] != gid:
            continue   # not this entity's representative (first-seen) member
        lines.append(_group_alloc_line(group, allocation, cpm))
        avails_by_group_id[gid] = tg.entity_avails_monthly(members)
    return lines, avails_by_group_id


def _expand_entity_group_rows(rows, drivers, matched_ids, groups_by_id,
                              default_targeting, flight_label):
    """FLOW_REWORK_PLAN.md Phase 3: `resolve_drafted_lines`' rows, for a
    FRESH draft, are built straight from `_synthesize_group_lines`'
    entity-collapsed lines -- one row per ENTITY among `matched_ids`, not
    one per group id. That's exactly right for computing a single dollar/
    impression figure per entity, but wrong as the FINAL row set: "one
    ticked row is one plan line, always" (decision 10) applies to a fresh
    draft too, the same as it does to `_allocate_group_rows`' re-pricing
    path.

    Expands each entity's single resolved row back out to one row per id
    in `matched_ids` sharing its entity -- the representative id keeps the
    computed row exactly as resolved; every sibling gets a fresh, UNPRICED
    ($0) row, seeded the same way `seed_media_plan_rows` always seeds a
    brand-new group row. Mirrors `_allocate_group_rows`' own "representative
    gets the share, siblings are left as they are" rule -- for a fresh
    draft, "as they are" is simply unpriced, since there's no prior row to
    leave alone.
    """
    if not matched_ids:
        return rows, drivers
    siblings_by_entity = {}
    for gid in matched_ids:
        group = groups_by_id.get(gid)
        if group is not None:
            siblings_by_entity.setdefault(tg.entity_id_of(group), []).append(gid)

    out_rows, out_drivers = [], []
    for row, driver in zip(rows, drivers):
        out_rows.append(row)
        out_drivers.append(driver)
        ids = group_ids_of(row)
        if len(ids) != 1 or ids[0] not in groups_by_id:
            continue
        rep_gid = ids[0]
        entity_id = tg.entity_id_of(groups_by_id[rep_gid])
        for sibling_gid in siblings_by_entity.get(entity_id, []):
            if sibling_gid == rep_gid:
                continue
            sibling_group = groups_by_id[sibling_gid]
            geo = tg.geo_label(sibling_group, label_for=_market_display_name).strip()
            only_premion = {"products": {}, "_premion_streaming_tv": True}
            entry_list = [(tg.audience_label(sibling_group), geo, sibling_gid)]
            sibling_row = seed_media_plan_rows(
                only_premion, entry_list, default_targeting, flight_label)[0]
            out_rows.append(sibling_row)
            out_drivers.append(DRIVER_IMPRESSIONS)
    return out_rows, out_drivers


def _allocate_group_rows(option, opt_intent, groups_by_id, flight_label,
                         default_targeting, n_months):
    """Re-price this option's CLEAN, group-owned rows against one allocation
    instruction (`opt_intent["group_allocation"]`/`group_cpm`) -- "split
    evenly across the audiences in the avail" becomes THIS, applied to
    whichever rows are currently selected and untouched, never a reason to
    regenerate rows from scratch.

    Reuses `resolve_drafted_lines`'s waterfall rather than reimplementing
    it: every row NOT being reallocated (a dirty group row a rep edited, or
    any non-group row already on the grid -- retargeting, AM, sports, fees,
    broadcast) is fed back in as its own `flat_amount` line at its CURRENT
    cost, so its money is committed inside the waterfall instead of being
    pre-subtracted -- which is what keeps `percent_of_total` a percent of
    the REAL total while `split_evenly`/`percent_of_remainder` divide only
    what's genuinely left.

    FLOW_REWORK_PLAN.md Phase 3: `_synthesize_group_lines` turns the clean,
    selected group ids into ONE waterfall line per ENTITY, not one per row
    -- so a stated split divides by entity count. Only that entity's
    REPRESENTATIVE row (the first clean, selected row for it) is ever
    written back; every other clean row sharing that entity is deliberately
    left untouched, keeping whatever it already had -- never silently split
    or duplicated. (This is what commit 8's allocation-basis caption exists
    to surface, since an unpriced sibling row next to a priced one would
    otherwise look like a mistake.) Only the output rows carrying
    `_group_ids` are ever written back, and only onto their OWN row
    (matched by group id) -- Cost/Impressions/CPM alone, so a clean row's
    Tactic/Geo/Targeting keep coming from the same resolution
    `seed_group_row` already gave it.

    Returns a list of plain-language notes.
    """
    committed_rows, clean_gids = [], []
    for row, dirty in zip(option["rows"], option["dirty"]):
        if is_group_line(row):
            gid = group_ids_of(row)[0]
            group = groups_by_id.get(gid)
            if dirty or group is None:
                committed_rows.append(row)
                continue
            clean_gids.append(gid)
        else:
            committed_rows.append(row)

    if not clean_gids:
        return []

    lines, avails_by_group_id = _synthesize_group_lines(
        clean_gids, opt_intent.get("group_allocation"), opt_intent.get("group_cpm"), groups_by_id)
    for row in committed_rows:
        lines.append({"product": "_committed", "allocation": {"flat_amount": _num(row.get("Cost"))}})

    rows, _touched_products, _touched_sports, unresolved, drivers = resolve_drafted_lines(
        lines, opt_intent.get("total_budget") or 0, flight_label,
        opt_intent.get("default_targeting") or default_targeting, default_targeting,
        avails_by_name=None, n_months=n_months, avails_by_group_id=avails_by_group_id)

    by_gid = {group_ids_of(row)[0]: (row, driver) for row, driver in zip(rows, drivers)
             if group_ids_of(row)}
    by_gid_in_option = {group_ids_of(r)[0]: i for i, r in enumerate(option["rows"])
                        if is_group_line(r)}
    for gid, (new_row, driver) in by_gid.items():
        idx = by_gid_in_option.get(gid)
        if idx is None:
            continue
        option["rows"][idx]["Cost"] = new_row["Cost"]
        option["rows"][idx]["Impressions"] = new_row["Impressions"]
        option["rows"][idx]["CPM"] = new_row["CPM"]
        option["driver"][idx] = driver
    return unresolved


def describe_group_allocation(opt_intent, groups_by_id, rows, dirty):
    """FLOW_REWORK_PLAN.md Phase 3, decision 12: a read-only caption naming
    the allocation `opt_intent["group_allocation"]` actually used for this
    option's group-owned lines -- the answer to "what did the draft
    assume?" that a new budget-split COLUMN would have duplicated (the Cost
    column already IS the split; a second display of the same number can
    only drift from it). Derived STRICTLY from the stored allocation intent
    -- never recomputed or inferred from the rows' own Cost figures -- and
    returns None the moment any of this option's group-owned rows has been
    hand-edited (dirty): a hand-edited row's cost is the rep's now, not the
    draft's, and this must never assert something false about a number it
    no longer produced.
    """
    if not opt_intent:
        return None
    allocation = opt_intent.get("group_allocation") or {}
    if not allocation:
        return None

    group_rows = [r for r in rows if is_group_line(r)]
    if any(d for r, d in zip(rows, dirty) if is_group_line(r)):
        return None

    entity_gids = {}
    for row in group_rows:
        for gid in group_ids_of(row):
            if gid in groups_by_id:
                entity_gids.setdefault(tg.entity_id_of(groups_by_id[gid]), []).append(gid)
    if not entity_gids:
        return None
    n_entities = len(entity_gids)

    alloc_type, value = next(iter(allocation.items()))
    entity_word = "entity" if n_entities == 1 else "entities"
    if alloc_type == "split_evenly":
        basis = f"Budget split evenly across {n_entities} {entity_word}"
    elif alloc_type == "percent_of_total":
        basis = f"{value}% of budget applied to each entity"
    elif alloc_type == "percent_of_remainder":
        basis = f"{value}% of the remaining budget applied to each entity"
    elif alloc_type == "flat_amount":
        try:
            basis = f"${float(value):,.0f} flat applied to each entity"
        except (TypeError, ValueError):
            basis = f"{value} flat applied to each entity"
    elif alloc_type == "percent_of_avails":
        basis = f"{value}% of avails applied to each entity"
    else:
        basis = "Budget allocation applied"

    # An entity with 2+ SEPARATE rows (never a merged row -- that's already
    # one line) got its whole computed share on the first; the rest are
    # left exactly as they were (commit 5's own rule) -- flagged here so an
    # unpriced sibling next to a priced one doesn't read as a mistake.
    flags = []
    for eid, gids in entity_gids.items():
        rows_for_entity = [r for r in group_rows if any(g in gids for g in group_ids_of(r))]
        if len(rows_for_entity) > 1:
            rep_group = groups_by_id[gids[0]]
            label = tg.entity_label_of(rep_group) or tg.audience_label(rep_group)
            flags.append(f"{label} has {len(rows_for_entity)} unmerged lines -- "
                        f"full share applied to the first.")

    return " ".join([basis + "."] + flags)


def reconcile_group_plan_lines(plan_options, groups, *, seed_group_row, fallback_geo="",
                               intent=None, realloc=False, confirmed_removals=(),
                               premion_selected=True, flight_label="",
                               default_targeting="", n_months=1):
    """The single owner of every Premion Streaming TV row that carries a
    targeting-group id. Called exactly once per run (see main()) -- never
    more than once, and idempotent when called with no change to its
    inputs: the same `groups`/`intent`/`realloc`/`confirmed_removals` leaves
    every option's rows, dirty flags, driver list and version untouched.
    That property is what makes import-then-draft, draft-then-import and
    clarify-after-either converge on the same end state regardless of
    order -- there is no path-dependent sequence to get wrong, only "what
    does the current state say" answered fresh every time.

    Returns a list of plain-language notes (a partially-deselected merged
    row kept on the plan, a re-allocation over a hand-edited row) for the
    caller to fold into whatever review list is showing.
    """
    real_groups = [g for g in (groups or []) if g.get("id")]
    groups_by_id = {g["id"]: g for g in real_groups}
    has_real_groups = any(g.get("terms") for g in real_groups)
    selected = selected_group_triples(real_groups, fallback_geo) if premion_selected else []
    selected_ids = {gid for _a, _g, gid in selected}
    confirmed = set(confirmed_removals or [])
    notes = []
    premion_label, _cpm = line_product_spec(GROUP_LINE_PRODUCT_KEY)

    # A multi-option draft resolves each option's OWN group_selection into
    # `matched_ids` at apply time (build_draft_prompt's "group_selection"
    # section) and stores it in `intent["options"][i]["selection"]
    # ["matched_ids"]` -- but that scoping used to reach only the
    # re-allocation step (below), never the row-seeding step, so a rerun
    # right after the draft added every OTHER option's groups back onto
    # every option, at a real, visible $0 (found on a real LiveWell draft:
    # two options, one meant to sell one of seven locations, ended up with
    # all seven rows, six of them zeroed). `union_matched_ids` is every id
    # ANY option's stored intent claimed; a group in that union but absent
    # from THIS option's own `matched_ids` was deliberately left off THIS
    # option by the draft, so it's excluded from this option's seeding.
    # A group in neither -- never claimed by anyone's intent, e.g. ticked
    # onto the shared avails table by a rep after the draft ran, or the
    # proposal was never drafted at all -- carries no per-option
    # information, and defaults to every option, exactly like before this
    # fix. That's what makes an option with no stored intent at all (a
    # hand-built plan, or one added by hand after a draft) fall through to
    # today's untouched global behavior: its own contribution to the
    # exclusion set is empty, so nothing about it can ever be excluded, and
    # nothing else's intent can carve anything out of ITS candidate set
    # either, because that only happens per-option, below.
    intent_options = (intent or {}).get("options") or []
    union_matched_ids = set()
    for opt_intent in intent_options:
        union_matched_ids |= set((opt_intent.get("selection") or {}).get("matched_ids") or [])

    for option in plan_options or []:
        changed = False
        opt_intent = _intent_for_option(intent, option)
        opt_selection = (opt_intent or {}).get("selection") or {}
        # `matched_ids` present (even as an empty list) means this option
        # went through a real draft resolution -- absent means no stored
        # intent at all, and excluded_for_option must stay empty so this
        # option falls all the way through to the original global set.
        if "matched_ids" in opt_selection:
            own_matched_ids = set(opt_selection.get("matched_ids") or [])
            excluded_for_option = union_matched_ids - own_matched_ids
        else:
            excluded_for_option = set()
        option_selected = [(a, g, gid) for a, g, gid in selected if gid not in excluded_for_option]

        # 1. Remove rows whose every group id is deselected or gone. A row
        # carrying SEVERAL ids (merge_plan_rows) survives if any of them is
        # still selected -- mirroring branch C's own partial-survival rule
        # for a group that was deleted entirely, not merely unticked.
        kept_rows, kept_dirty, kept_driver = [], [], []
        for row, dirty, driver in zip(option["rows"], option["dirty"], option["driver"]):
            if not is_group_line(row):
                kept_rows.append(row); kept_dirty.append(dirty); kept_driver.append(driver)
                continue
            ids = group_ids_of(row)
            live = [gid for gid in ids if gid in selected_ids]
            if live:
                if len(live) < len(ids) and dirty:
                    notes.append(
                        f"{row.get('Tactic', premion_label)} ({row.get('Geo', '')}) still covers an "
                        f"audience you unticked in the avails table -- edit or split the line if that's "
                        f"not what you want.")
                kept_rows.append(row); kept_dirty.append(dirty); kept_driver.append(driver)
                continue
            # No id on this row is still selected.
            if not dirty or any(gid in confirmed for gid in ids):
                changed = True
                continue                       # a clean removal, or a confirmed one
            # Dirty and not confirmed -- kept. The D2 fold-back is what
            # actually stops an unconfirmed uncheck from ever reaching here
            # (it queues a confirmation instead of writing it); this is the
            # belt to that braces, for any other caller of this function.
            notes.append(
                f"Kept {row.get('Tactic', premion_label)} ({row.get('Geo', '')}) on the plan -- it's "
                f"been edited and its audience was unticked without confirming the removal.")
            kept_rows.append(row); kept_dirty.append(dirty); kept_driver.append(driver)
        if changed:
            option["rows"], option["dirty"], option["driver"] = kept_rows, kept_dirty, kept_driver

        # 2. Add one row per newly-selected group with no owning row yet --
        # scoped to THIS option's own draft intent (option_selected) rather
        # than every globally-included group, so a sibling option's own
        # selection doesn't leak its rows in here at a phantom $0.
        owned_ids = {gid for row in option["rows"] if is_group_line(row) for gid in group_ids_of(row)}
        to_add = [(a, g, gid) for a, g, gid in option_selected if gid not in owned_ids]
        for _audience, _geo, gid in to_add:
            option["rows"].append(seed_group_row(groups_by_id[gid], option["breakout"]))
            option["dirty"].append(False)
            option["driver"].append(DRIVER_IMPRESSIONS)
            changed = True

        # 3. No real groups at all -> today's single ungrouped Premion row,
        # exactly as `seed_media_plan_rows` has always produced for a form
        # that has never touched targeting groups.
        if not has_real_groups and premion_selected:
            has_ungrouped_premion = any(
                row_belongs_to_tactic(r.get("Tactic"), premion_label) and not group_ids_of(r)
                for r in option["rows"])
            if not has_ungrouped_premion:
                option["rows"].append(seed_group_row(None, option["breakout"]))
                option["dirty"].append(False)
                option["driver"].append(DRIVER_IMPRESSIONS)
                changed = True

        # 4. Re-price selected, CLEAN group rows against the stored intent.
        if realloc and intent:
            opt_intent = _intent_for_option(intent, option)
            if opt_intent:
                realloc_notes = _allocate_group_rows(
                    option, opt_intent, groups_by_id, flight_label,
                    default_targeting, n_months)
                if realloc_notes:
                    notes.extend(realloc_notes)
                changed = True

        if changed:
            option["version"] += 1

    return notes


def _intent_for_option(intent, option):
    """The stored `draft_plan_intent` entry for THIS option, matched by
    name (options are drafted and stored in the same order they're
    rendered) -- or the first entry when there's only one, since a
    single-option plan's option may have been renamed after the draft."""
    options = (intent or {}).get("options") or []
    if not options:
        return None
    for opt_intent in options:
        if opt_intent.get("name") == option.get("name"):
            return opt_intent
    return options[0] if len(options) == 1 else None


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


def new_plan_option(name, rows, driver=None, breakout=BREAKOUT_MONTHLY, breakout_locked=False):
    """One media plan option: its own name, line set, per-line dirty/driver
    tracking, breakout mode and editor version. Everything the single plan
    used to keep in flat session_state keys now lives per option.

    `breakout_locked` (default False, same shape as `color_locked`/
    `include_locked` elsewhere in this app): False means this option's
    Breakout keeps following the setup band's own Plan basis as it changes,
    which is what a genuinely blank, just-created option should do. A caller
    that's giving this option an explicit, deliberate breakout -- a
    rehydrated/restored proposal, a draft -- passes True so the band can
    never silently override it. See the per-option Breakout radio in
    main() for the other half: it's what actually re-syncs an unlocked
    option and sets this flag the moment a rep changes the radio by hand."""
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
        "option_breakout_locked": breakout_locked,
        "version": 0,
    }


def copy_plan_option(source, name):
    """A new option cloned from an existing one -- all lines and settings, so
    the user edits the delta instead of rebuilding the plan. The clone's rows
    start dirty: they're a deliberate copy, not a fresh seed, and must not be
    silently re-seeded out from under the user.

    Locked, regardless of the source's own lock state: copying is itself a
    one-time, deliberate decision to carry these settings over, not an
    instruction to keep tracking the band on the source's behalf forever
    after -- same "keeps that option's own breakout" rule this function's
    docstring already stated, just also applied to the new lock flag."""
    return {
        "name": name,
        "rows": [dict(r) for r in source["rows"]],
        "dirty": [True] * len(source["rows"]),
        "driver": list(source["driver"]),
        "breakout": source["breakout"],
        "option_breakout_locked": True,
        "version": 0,
    }


def merge_plan_rows(option, indexes):
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
    # merged row is internally consistent the instant it's created. Cost is
    # always NET on the grid now (FLOW_REWORK_PLAN.md Phase 4b), so there is
    # no markup to divide back out here -- that's exactly why row_markup no
    # longer has a call site in this function.
    survivor["CPM"] = ((new_cost * 1000.0) / new_impressions
                       if new_impressions else 0.0)
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


def unlink_deleted_group_rows(groups, prev_rows, edited_rows):
    """The inverse of reconcile_group_plan_lines' own add/remove: a rep
    deleting a group-owned row directly off the media plan grid (its own
    "-") must unselect that row's group too, or the very next
    reconcile_group_plan_lines pass -- which only knows "a selected group
    with no owning row gets one added back" -- resurrects the row the rep
    just deleted. Unchecking a group already removes its row (branch 1 of
    reconcile_group_plan_lines); this is the missing other direction, found
    live: the two disagreed and the deleted line kept reappearing.

    Compares the SET of group ids owning any row before vs. after the edit
    (not row-by-row -- a merged row can carry several ids, and a duplicated
    row can leave a group still owning ANOTHER row after just one copy is
    deleted, which must not unselect it). Every id that owned a row before
    and owns none after gets `include_in_plan=False, include_locked=True`
    -- locked for the same reason an unchecked box locks: this was the
    rep's own deliberate action, never silently re-ticked by a later draft
    or reconciliation pass.

    Returns the SAME `groups` object, untouched, when nothing was actually
    removed -- so a caller can tell "nothing to do" apart from "wrote a new
    list" with a plain `is` check, the same idiom apply_avails_autofill's
    caller uses for its own no-churn return.
    """
    prev_ids = set()
    for row in prev_rows:
        prev_ids.update(group_ids_of(row))
    if not prev_ids:
        return groups
    remaining_ids = set()
    for row in edited_rows:
        remaining_ids.update(group_ids_of(row))
    removed_ids = prev_ids - remaining_ids
    if not removed_ids:
        return groups
    return [dict(g, include_in_plan=False, include_locked=True)
           if g.get("id") in removed_ids else g
           for g in (groups or [])]


def rescale_rows_for_breakout_change(option, n_months, broadcast_months=None):
    """Keep an option's own rows meaning the same thing when the rep flips
    its Monthly/Full Flight radio -- found live, testing the Capital Media
    avail: the toggle changed nothing about the stored numbers, only how
    `compute_plan_totals` INTERPRETS them, so the same stored Cost/
    Impressions read as a monthly figure one moment and the whole flight's
    total the next. For anything but a single-month flight that produces
    two genuinely different (both wrong) totals; for a single-month flight
    it happens to produce the SAME number either way, which is the "Full
    flight and monthly plan toggle came out to the same numbers" symptom
    reported directly.

    Detects an actual flip by comparing against `option["_breakout_basis"]`
    (absent -- a freshly created option -- defaults to the option's own
    current breakout, so a brand-new option's first render is never treated
    as a change) and, on a genuine flip, rescales every RATE row's Cost and
    Impressions by that row's own month count -- broadcast rows are excluded
    (already followed by their own mechanism, from a fixed Wide Orbit total
    rather than a proportional rescale) and flat fees are excluded (never
    scaled by month count, by design, the same rule spread_rows_over_months
    and compute_plan_totals already follow). Applies to EVERY row regardless
    of its dirty flag -- dirty only means "don't silently reseed this row's
    Audience/Geo/CPM from the shared defaults," never "this row's own number
    is exempt from meaning what the basis label says it means."

    Must be called AFTER the option's own `st.radio(...)` widget has already
    set `option["breakout"]` to this run's value -- calling it earlier reads
    the PREVIOUS render's value and can never observe a flip on the same run
    the rep actually clicked it.
    """
    prev_basis = option.get("_breakout_basis", option["breakout"])
    # Always written, even when nothing changed THIS call -- otherwise an
    # absent key keeps re-deriving from whatever `option["breakout"]"
    # CURRENTLY says (the default above), which is exactly wrong the moment
    # this function is first called AFTER a flip: with no persisted prior
    # value, the default reads the just-changed CURRENT breakout back to
    # itself and the flip is never detected at all. This is a real bug this
    # function's own tests caught, not a hypothetical.
    option["_breakout_basis"] = option["breakout"]
    if prev_basis == option["breakout"]:
        return False
    was_full_flight = prev_basis.startswith("Full Flight")
    is_full_flight = option["breakout"].startswith("Full Flight")
    for row in option["rows"]:
        if is_flat_fee_row(row) or is_broadcast_row(row):
            continue
        row_months = (broadcast_months if broadcast_months and is_broadcast_row(row)
                     else n_months)
        row_months = max(1, int(row_months))
        if was_full_flight and not is_full_flight:
            row["Cost"] = _num(row.get("Cost")) / row_months
            row["Impressions"] = _num(row.get("Impressions")) / row_months
        elif not was_full_flight and is_full_flight:
            row["Cost"] = _num(row.get("Cost")) * row_months
            row["Impressions"] = _num(row.get("Impressions")) * row_months
    return True


def reconcile_plan_rows(option, edited_rows):
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
        recompute_row(row, driver)
        if (_num(row.get("Impressions")), _num(row.get("Cost"))) != before:
            recomputed_any = True
        new_drivers.append(driver)

    option["rows"] = edited_rows
    option["dirty"] = new_dirty
    option["driver"] = new_drivers
    return recomputed_any or row_count_changed


def matched_avails_for_row(row, groups_by_id):
    """The avails figure a plan line's OWN targeting actually backs, or None
    when it has no such match.

    A line matches only through `_group_ids` (`group_ids_of`) -- the same
    join merge/split/the map legend already use, never a text match on
    Targeting/Geo. Zero matching ids (a drafted line, a quick-add line, a
    hand-typed line, a stale id whose group was deleted) is unmatched, full
    stop -- None, not 0, so a caller can tell "no match" apart from "matched
    a group with no avails pulled yet" even though both mean no percentage.

    FLOW_REWORK_PLAN.md Phase 3: two or more ids (a line built by merging
    several avails-backed lines into one, via `merge_plan_rows` -- entity
    grouping alone never puts more than one id on a row) are bucketed by
    `tg.entity_id_of` first. Two groups of the SAME real-world entity (e.g.
    a 5mi and a 10mi radius tier for one dealership) **max** -- overlapping
    households, not additional reach, per the Annapolis/Plaza Motors/
    Wilmington evidence in the plan doc. Two DIFFERENT entities (a line
    merged from Denver + Atlanta) still **sum**, exactly as before -- a rep
    merging two unrelated markets into one line wants their combined
    avails, not neither market alone.
    """
    ids = [gid for gid in group_ids_of(row) if gid in groups_by_id]
    if not ids:
        return None
    buckets = tg.entities_of([groups_by_id[gid] for gid in ids])
    return sum(tg.entity_avails_monthly(bucket) for bucket in buckets.values())


def entity_prefixed_targeting(row, groups_by_id, show_label):
    """FLOW_REWORK_PLAN.md Phase 3, "Show Label in plan": prefixes a
    group-owned row's Targeting text with its entity's label, e.g.
    "Toyota of Annapolis | Subaru intenders, 10-mile radius" -- so a
    dealer-group client sees which store each line belongs to. (The plan
    doc's own worked example renders the entity name in bold; this ships
    the plain-text prefix only -- true run-level bold on part of a table
    cell needs XML-level formatting work `fill_table_rows` doesn't do
    today, out of scope for this phase. See DECISIONS.md.)

    Falls through to the plain Targeting text UNCHANGED when the toggle
    is off, the row isn't group-backed, or its entity has no label -- a
    no-op by default and for every proposal before this phase."""
    targeting = str(row.get("Targeting", ""))
    if not show_label:
        return targeting
    ids = group_ids_of(row)
    if not ids or not groups_by_id:
        return targeting
    group = groups_by_id.get(ids[0])
    if group is None:
        return targeting
    label = tg.entity_label_of(group)
    if not label:
        return targeting
    return f"{label} | {targeting}" if targeting else label


def compute_plan_totals(rows, breakout_mode, n_months, flight_label,
                        broadcast_months=None, groups_by_id=None,
                        coviewing_multiplier=None, markup=1.0, show_entity_label=False):
    """Per-line monthly/full-flight impressions and cost for one option, plus
    the four running totals.

    FLOW_REWORK_PLAN.md Phase 4b: `row["Cost"]`/`row["CPM"]` are always NET
    now -- the agency gross-up is a rep-only, order-wide checkbox applied
    exactly once, HERE, at display/deck time, never on the grid itself. This
    is the only place the ×1.15 is ever applied to a dollar figure. `markup`
    grosses each row's own cost AND `preview_rows["cpm"]` (via `row_markup`,
    so the broadcast exemption still applies -- a Wide Orbit cost is already
    gross and is never marked up regardless of the checkbox). Impressions
    are never touched by markup, by construction: they come straight off the
    row exactly as entered. `markup=1.0` (the default, and always the
    effective value when the checkbox is off) is a no-op, so every existing
    caller that doesn't pass it keeps getting the net figures back
    unchanged. A flat-fee row is exempt from the gross-up the same way it
    was under the old per-row toggle (its Cost is never passed through
    `cost_from_impressions`/`impressions_from_cost` at all) -- unchanged by
    this phase, not a new exemption.

    `groups_by_id`, when given, additionally resolves each line's own
    matched avails (see `matched_avails_for_row`) at both bases, using that
    line's own row_months -- a broadcast line's own schedule length, not the
    plan's -- so a merged or broadcast line's full-flight percentage divides
    by the flight length that line actually runs in. The full-flight side
    goes through `matched_avails_full_flight_for_row`, which prefers a
    group's own exact `avails_full_flight` over `monthly * row_months` --
    see that function for why the multiply-out alone can't tie an
    avails-PDF import's SOV denominator to the document's stated total
    once the flight spans months of different lengths.

    `coviewing_multiplier`, when given, additionally stamps each ELIGIBLE
    line (`is_coviewing_eligible_row`) with the additional person-level
    impressions beyond its own household figure -- None for every
    ineligible or flat-fee line, never a fabricated number."""
    preview_rows = []
    monthly_impressions_total = monthly_cost_total = 0.0
    flight_impressions_total = flight_cost_total = 0.0

    for row in rows:
        if not str(row.get("Tactic", "")).strip():
            continue
        flat_fee = is_flat_fee_row(row)
        # The one markup lookup per row, reused below for both Cost and CPM
        # so the two can never disagree about which rate they're quoting.
        effective = row_markup(row, markup)
        row_avails_monthly = None if flat_fee else matched_avails_for_row(row, groups_by_id or {})
        row_avails_full_flight = None
        coviewing_eligible = not flat_fee and is_coviewing_eligible_row(row)
        coviewing_additional_monthly = None
        if flat_fee:
            # A flat fee is a one-time full-flight cost, not a per-month rate
            # -- it must NOT scale with month count the way rate rows do, so
            # it's accumulated straight into the flight total. Never grossed,
            # same as before this phase: a flat fee's stored Cost has never
            # passed through cost_from_impressions/impressions_from_cost, so
            # there is no markup relationship to apply here either.
            full_flight_impressions = 0.0
            full_flight_cost = _num(row.get("Cost"))
            monthly_impressions = 0.0
            monthly_cost = full_flight_cost / n_months
        else:
            entered_impressions = _num(row.get("Impressions"))
            entered_cost = _num(row.get("Cost")) * effective
            # A broadcast line scales by the months its own schedule runs in,
            # not the plan's. A four-week May buy inside a three-month plan
            # flight is one month of broadcast, and multiplying its monthly
            # figure by three would have invented $56,000 of spend that no
            # station is going to run. Its matched avails scale the same way.
            row_months = (broadcast_months if broadcast_months and is_broadcast_row(row)
                          else n_months)
            row_months = max(1, int(row_months))
            if row_avails_monthly is not None:
                row_avails_full_flight = matched_avails_full_flight_for_row(
                    row, groups_by_id or {}, row_months)
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

        if coviewing_eligible and coviewing_multiplier:
            coviewing_additional_monthly = round(monthly_impressions * (coviewing_multiplier - 1))

        monthly_impressions_total += monthly_impressions
        monthly_cost_total += monthly_cost
        flight_impressions_total += full_flight_impressions
        flight_cost_total += full_flight_cost
        preview_rows.append({
            "tactic": str(row["Tactic"]), "flight": str(row.get("Flight", "")) or flight_label,
            "geo": str(row.get("Geo", "")),
            "targeting": entity_prefixed_targeting(row, groups_by_id, show_entity_label),
            "monthly_impressions": monthly_impressions, "monthly_cost": monthly_cost,
            "full_flight_impressions": full_flight_impressions, "full_flight_cost": full_flight_cost,
            "matched_avails_monthly": row_avails_monthly,
            "matched_avails_full_flight": row_avails_full_flight,
            "coviewing_eligible": coviewing_eligible,
            "coviewing_additional_monthly": coviewing_additional_monthly,
            "is_flat_fee": flat_fee,
            "cpm": _num(row.get("CPM")) * effective,
        })

    return {
        "preview_rows": preview_rows,
        "monthly_impressions": monthly_impressions_total,
        "monthly_cost": monthly_cost_total,
        "full_flight_impressions": flight_impressions_total,
        "full_flight_cost": flight_cost_total,
    }


def effective_cpm_with_coviewing(preview_rows, multiplier):
    """The plan's blended CPM once eligible (CTV/Sports) lines' impressions
    are boosted by the co-viewing multiplier -- for the footnote's trailing
    clause, never a table column (see BACKLOG.md's Co-viewing item).

    Cost is untouched; only an ELIGIBLE rate line's own impressions are
    boosted before summing -- a broadcast or AM line sold on household
    impressions still costs the same per household impression, so blending
    its unboosted impressions in with a boosted CTV line's would understate
    the CTV line's own effective rate and overstate the mixed line's. Flat
    fees are excluded from both sides, same as the plain blended CPM
    (`_option_payload`'s own `blended` calculation) they're compared against.
    """
    rate_rows = [r for r in preview_rows if not r["is_flat_fee"]]
    cost = sum(r["monthly_cost"] for r in rate_rows)
    impressions = sum(
        r["monthly_impressions"] * (multiplier if r["coviewing_eligible"] else 1)
        for r in rate_rows)
    return cost / impressions * 1000 if impressions else 0.0


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


_SHAPE_KIND_LABELS = {
    "chart": ("chart", "charts"), "table": ("table", "tables"),
    "picture": ("picture", "pictures"), "text": ("text box", "text boxes"),
}


def shape_kind_summary(slide):
    """"1 chart, 2 pictures" -- a cheap structural fingerprint for a slide
    extract_slide_text has nothing to show for. The slide-vault picker has
    no thumbnail (no PowerPoint at contribute time, same constraint as
    rendering everywhere else in this app), so a chart- or picture-heavy
    slide previews as blank text; this is the free signal available without
    one, not a substitute for actually opening "Full text" or the deck.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    counts = {}
    for shape in slide_map.iter_all_shapes(slide.shapes):
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                continue  # children are yielded separately by iter_all_shapes
        except NotImplementedError:
            pass
        if getattr(shape, "has_chart", False) and shape.has_chart:
            kind = "chart"
        elif getattr(shape, "has_table", False) and shape.has_table:
            kind = "table"
        else:
            try:
                is_picture = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            except NotImplementedError:
                is_picture = False
            if is_picture:
                kind = "picture"
            elif shape.has_text_frame and shape.text_frame.text.strip():
                kind = "text"
            else:
                continue  # an empty placeholder, a line, decoration -- not worth counting
        counts[kind] = counts.get(kind, 0) + 1

    parts = []
    for kind in ("chart", "table", "picture", "text"):
        n = counts.get(kind, 0)
        if n:
            singular, plural = _SHAPE_KIND_LABELS[kind]
            parts.append(f"{n} {singular if n == 1 else plural}")
    return ", ".join(parts)


def extract_case_study_shape_summary(path):
    """One shape_kind_summary() per slide, same shape as extract_case_study_text
    (one entry per slide, same order) so a caller can zip them together."""
    prs = Presentation(path)
    return [shape_kind_summary(slide) for slide in prs.slides]


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
        print(f"[case study] couldn't read {local_path.name}: {type(exc).__name__}: {exc}")
        st.error("Couldn't read that file as a .pptx -- check that it's a valid PowerPoint "
                 "file, not currently open elsewhere, and try again.")
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


# ---------------- Slide vault ----------------
# Colleagues add individual slides they like -- a chart, a capabilities page,
# a research stat -- separate from both the master deck and the case-study
# vault. Deliberately never pre-selected into a proposal the way a matching
# case study is: a vault slide is a colleague's own favourite, not vetted
# proof, so auto-adding one to a client deck is a bigger risk than
# auto-adding a case study. See CLAUDE.md / BACKLOG.md "Slide vault".

VAULT_PLACEMENT_LABELS = {
    assembly.VAULT_PLACEMENT_FRONT: "Front of the deck (right after the cover)",
    assembly.VAULT_PLACEMENT_BEFORE_PLAN: "Before the media plan (with the case studies)",
    assembly.VAULT_PLACEMENT_APPENDIX: "Appendix (end of the deck)",
}

# Looser than verticals/products on purpose -- what KIND of slide this is
# matters more for a vault slide than for a case study, but the taxonomy is a
# guess. It's one text column, so getting it wrong later is a data edit, not
# a migration (see supabase_schema.sql Stage 12).
VAULT_PURPOSE_TAGS = ["capabilities", "research", "creative", "pricing", "other"]


def build_vault_slide_prompt(picked_texts, filename):
    """picked_texts: [(slide_index, text), ...] -- only the ticked slides,
    since a rep has already decided which slides matter; Claude tags them,
    it doesn't pick them."""
    slides_json = [{"index": i, "text": text[:2500]} for i, text in picked_texts]
    return f"""You are cataloguing individual PowerPoint slides a Premion seller wants to reuse in future proposals -- not a whole deck, just these specific slides. Return ONLY valid JSON matching this schema -- no markdown fences, no preamble:

{{"slides": [{{"index": 0, "title": "", "summary": "", "verticals": [], "products": [], "purpose": ""}}]}}

One entry per slide, in the same order as the input, each "index" copied exactly from it.

- "title": a short human title for this ONE slide (e.g. "Q3 Streaming Growth Chart"), in the deck's own words where possible.
- "summary": ONE sentence a seller can skim to know what this slide shows. No more than about 20 words.
- "verticals": every vertical this slide is genuinely relevant to, each exactly one of {list(VERTICALS.values())} (never "none"). An empty list is fine -- plenty of slides are vertical-agnostic. Do not stretch.
- "products": the Premion products this slide speaks to, each exactly one of {CASE_STUDY_PRODUCT_TAGS}. An empty list is fine.
- "purpose": exactly one of {VAULT_PURPOSE_TAGS} -- what KIND of slide this is, not what it's about.

Source file name: {filename}

Slides (JSON): {json.dumps(slides_json)}
"""


def call_claude_vault_slide_tags(picked_texts, filename):
    return _call_claude_json(build_vault_slide_prompt(picked_texts, filename),
                             label="vault_slide_tags")


def render_add_vault_slide():
    """Upload a .pptx, tick the individual slides worth keeping, have Claude
    propose tags for each, edit them, and store them.

    Each ticked slide becomes its own vault entry, but the source deck is
    uploaded once and shared by every entry picked from it (db.upload_
    slide_vault) -- the deck is never split apart, since extracting one
    slide into its own .pptx is the cross-deck copy problem that produced
    five distinct corruption bugs on the case-study vault (see CLAUDE.md).
    """
    st.header("Add vault slide")
    st.caption("Got a slide worth reusing -- a chart, a capabilities page, a research stat -- "
               "even if the rest of the deck around it isn't? Upload it, tick the slide(s) "
               "worth keeping, and each one becomes its own entry everyone building a "
               "proposal can pull in.")

    upload = st.file_uploader("Source deck (.pptx)", type=["pptx"], key="vault_upload")
    if not upload:
        st.session_state.pop("vault_suggestion", None)
        st.session_state.pop("vault_suggestion_for", None)
        return

    local_path = db.scratch_dir("premion_vault_slide_uploads") / upload.name
    local_path.write_bytes(upload.getvalue())

    try:
        slide_texts = extract_case_study_text(str(local_path))
        shape_summaries = extract_case_study_shape_summary(str(local_path))
    except Exception as exc:
        print(f"[slide vault] couldn't read {local_path.name}: {type(exc).__name__}: {exc}")
        st.error("Couldn't read that file as a .pptx -- check that it's a valid PowerPoint "
                 "file, not currently open elsewhere, and try again.")
        return
    st.caption(f"{len(slide_texts)} slide(s), "
               f"{local_path.stat().st_size / 1024 / 1024:.1f} MiB")

    st.write("**Tick the slides worth keeping.** There's no thumbnail here -- each slide is "
             "shown as its extracted text plus a shape count (\"1 chart, 2 pictures\"), so a "
             "chart- or image-heavy slide still gives you something to go on even with little "
             "or no text. Open \"Full text\" if that's still not enough to tell.")

    ticked = []
    for i, text in enumerate(slide_texts):
        text_preview = text[:120].replace("\n", " | ")
        shapes = shape_summaries[i]
        if text_preview and shapes:
            preview = f"{text_preview}  —  ({shapes})"
        elif text_preview:
            preview = text_preview
        elif shapes:
            preview = f"(no extractable text — {shapes})"
        else:
            preview = "(no extractable text or shapes found)"
        picked = st.checkbox(f"Slide {i + 1}: {preview}",
                             key=f"vault_slide_pick_{upload.name}_{i}")
        with st.expander(f"Full text of slide {i + 1}", expanded=False):
            st.text(text or "(no extractable text)")
        if picked:
            ticked.append(i)

    if not ticked:
        st.caption("Tick at least one slide above to continue.")
        return

    suggestion_key = f"{upload.name}:{tuple(ticked)}"
    if st.session_state.get("vault_suggestion_for") != suggestion_key:
        if st.button("Read the ticked slides and suggest tags", type="primary"):
            with st.spinner("Reading the slide(s)..."):
                picked_texts = [(i, slide_texts[i]) for i in ticked]
                result, error = call_claude_vault_slide_tags(picked_texts, upload.name)
            if error:
                st.error(error)
            else:
                st.session_state["vault_suggestion"] = {
                    s["index"]: s for s in (result.get("slides") or []) if "index" in s}
                st.session_state["vault_suggestion_for"] = suggestion_key
                st.rerun()
        st.caption("Or fill the fields in yourself below.")

    suggestions = st.session_state.get("vault_suggestion") or {}
    if suggestions:
        st.success("Claude's suggestions are filled in below -- correct anything before saving.")

    vertical_labels = {v: k for k, v in VERTICALS.items() if v != "none"}
    added_by = st.text_input("Added by",
                             value=st.session_state.get("vault_added_by") or current_user() or "",
                             placeholder="Your name")

    picks = []
    for i in ticked:
        suggestion = suggestions.get(i, {}) or {}
        st.markdown(f"---\n**Slide {i + 1}**")
        title = st.text_input(
            "Title", key=f"vault_title_{i}",
            value=suggestion.get("title", "") or f"{Path(upload.name).stem} -- slide {i + 1}")
        summary = st.text_area("One-line summary", value=suggestion.get("summary", ""),
                               height=68, key=f"vault_summary_{i}")
        verticals = st.multiselect(
            "Verticals", list(vertical_labels), format_func=lambda v: vertical_labels[v],
            default=_valid_tags(suggestion.get("verticals"), vertical_labels),
            key=f"vault_verticals_{i}")
        products = st.multiselect(
            "Product tags", CASE_STUDY_PRODUCT_TAGS,
            default=_valid_tags(suggestion.get("products"), CASE_STUDY_PRODUCT_TAGS),
            key=f"vault_products_{i}")
        purpose_default = suggestion.get("purpose")
        purpose = st.selectbox(
            "Slide purpose", VAULT_PURPOSE_TAGS, key=f"vault_purpose_{i}",
            index=VAULT_PURPOSE_TAGS.index(purpose_default)
            if purpose_default in VAULT_PURPOSE_TAGS else len(VAULT_PURPOSE_TAGS) - 1)
        placement = st.selectbox(
            "Suggested placement (a rep can change this per proposal)",
            list(VAULT_PLACEMENT_LABELS), format_func=lambda p: VAULT_PLACEMENT_LABELS[p],
            index=list(VAULT_PLACEMENT_LABELS).index(assembly.VAULT_PLACEMENT_BEFORE_PLAN),
            key=f"vault_placement_{i}")
        picks.append({"slide_index": i, "title": title.strip(), "summary": summary.strip(),
                      "verticals": verticals, "products": products, "purpose": purpose,
                      "placement": placement})

    if st.button("Save to the vault", disabled=not all(p["title"] for p in picks)):
        if not added_by.strip():
            st.warning("Add your name first so the vault records who contributed this.")
            return
        st.session_state["vault_added_by"] = added_by
        with st.spinner("Optimizing and uploading..."):
            rows, stats, error = db.upload_slide_vault(
                str(local_path), upload.name, picks, added_by.strip())
        if stats:
            st.caption(f"Optimized {stats['source_size'] / 1024 / 1024:.1f} MiB -> "
                       f"{stats['dst_size'] / 1024 / 1024:.1f} MiB before storing.")
        if error:
            st.error(error)
        else:
            st.success(f"Saved {len(rows)} slide(s) to the vault.")
            # Same "it works right now, a little less faithfully" note as
            # the case-study vault -- see render_add_case_study.
            pending, _ = db.fetch_slide_vault(active_only=False)
            waiting = [r for r in (pending or []) if db.slide_vault_render_path(r) == "copy"]
            st.warning(
                "**Needs an image for full fidelity.** Until then it goes into decks as a "
                "copied slide, which renders slightly less faithfully. Rendering needs "
                "PowerPoint and the brand font, so it can't happen here.\n\n"
                f"Run this locally when you get a chance — {len(waiting)} vault "
                f"slide{'s are' if len(waiting) != 1 else ' is'} waiting:\n\n"
                "```\npython render_slide_vault_images.py --pending\n```")
            for key in ("vault_suggestion", "vault_suggestion_for"):
                st.session_state.pop(key, None)


def vault_slide_source(row, placement):
    """How this vault slide goes into a deck: a rendered image, or the
    source .pptx to copy one slide out of. One resolver for both graft
    sites, the same reason case_study_source is one function.

    `placement` is passed in rather than read off the row, deliberately --
    it's the REP'S chosen placement for THIS proposal (from the picker at
    generate time, or the placement recorded in a stored proposal's
    form_json at rebuild time), never the vault entry's own stored default.
    Reading the row's default here would make a rebuild drift from what was
    actually presented the moment someone edited the vault entry afterward.
    """
    if db.slide_vault_render_path(row) == "image":
        image = db.slide_vault_image(row["id"], row["slide_image"])
        return {"images": [image], "title": row.get("title"), "placement": placement}
    return {"path": db.slide_vault_file(row["storage_path"]), "slides": [row["slide_index"]],
            "title": row.get("title"), "placement": placement}


def vault_slide_filename(row):
    """A filename a seller can hand to a client, built from the title --
    same sanitization case_study_filename uses."""
    base = (row.get("title") or Path(row["filename"]).stem).strip()
    base = re.sub(r"[/&+]", " ", base)
    base = re.sub(r"[^\w\s-]", "", base)
    base = re.sub(r"[\s_-]+", "_", base).strip("_")
    return f"{(base or 'vault_slide')[:80]}.pptx"


def render_vault_slide_download(row, key_prefix):
    """Download button for one vault slide's SOURCE DECK, fetched on demand.

    Labelled explicitly as the source deck, not "this slide" -- what's
    stored is the whole upload (see db.upload_slide_vault), so a download
    hands over every slide the colleague uploaded, not just the picked one.
    Same lazy-fetch discipline as render_case_study_download: never eager,
    or every visible row in the browser pulls a deck out of storage.
    """
    slot = f"{key_prefix}_{row['id']}"
    path = db.slide_vault_cached_path(row["storage_path"])

    if path is None:
        if not st.button("Get source deck", key=f"{slot}_fetch",
                         help="Fetches the source deck from the vault, then offers it as a "
                              "download. Contains every slide in the original upload, not "
                              "just this one."):
            return
        try:
            path = db.slide_vault_file(row["storage_path"])
        except Exception as exc:
            st.warning(f"Couldn't fetch that deck ({db.describe_error(exc)}).")
            return

    with open(path, "rb") as handle:
        st.download_button(
            f"⬇ Download source deck (slide {row['slide_index'] + 1} of it)",
            data=handle.read(), file_name=vault_slide_filename(row), mime=PPTX_MIME,
            key=f"{slot}_download")


def render_vault_slide_finder():
    """Standalone vault page: browse and edit every vault slide."""
    st.header("Slide vault")
    st.caption("Individual slides colleagues have added -- charts, capabilities pages, "
               "research, whatever's worth reusing. To put one *in* a proposal, use the "
               "picker just above Generate on the Build page.")
    rows, warning = db.fetch_slide_vault(active_only=False)
    if warning:
        st.warning(warning)
        return
    if not rows:
        st.info("The slide vault is empty. Add one from the \"Add vault slide\" page.")
        return
    _render_vault_slide_browser(rows)


def _render_vault_slide_browser(rows):
    vertical_labels = {v: k for k, v in VERTICALS.items() if v != "none"}
    col1, col2, col3, col4 = st.columns([2, 2, 2, 3])
    with col1:
        filter_verticals = st.multiselect("Vertical", list(vertical_labels),
                                          format_func=lambda v: vertical_labels[v],
                                          key="slide_vault_filter_verticals")
    with col2:
        filter_products = st.multiselect("Product", CASE_STUDY_PRODUCT_TAGS,
                                         key="slide_vault_filter_products")
    with col3:
        filter_purpose = st.multiselect("Purpose", VAULT_PURPOSE_TAGS,
                                        key="slide_vault_filter_purpose")
    with col4:
        query = st.text_input("Search title or summary", key="slide_vault_search").strip().lower()
    show_inactive = st.checkbox("Include deactivated", key="slide_vault_show_inactive")

    def matches(row):
        if not show_inactive and not row.get("active", True):
            return False
        if filter_verticals and not set(filter_verticals) & set(row.get("verticals") or []):
            return False
        if filter_products and not set(filter_products) & set(row.get("products") or []):
            return False
        if filter_purpose and row.get("purpose") not in filter_purpose:
            return False
        if query and query not in f"{row.get('title') or ''} {row.get('summary') or ''}".lower():
            return False
        return True

    shown = [r for r in rows if matches(r)]
    st.caption(f"{len(shown)} of {len(rows)} vault slide(s)")

    for row in shown:
        title = row.get("title") or f"{row['filename']} -- slide {row['slide_index'] + 1}"
        with st.expander(title, expanded=False):
            edit_title = st.text_input("Title", value=row.get("title") or "",
                                       key=f"slide_vault_edit_title_{row['id']}")
            edit_summary = st.text_area("Summary", value=row.get("summary") or "",
                                        height=68, key=f"slide_vault_edit_summary_{row['id']}")
            edit_verticals = st.multiselect(
                "Verticals", list(vertical_labels), format_func=lambda v: vertical_labels[v],
                default=_valid_tags(row.get("verticals"), vertical_labels),
                key=f"slide_vault_edit_verticals_{row['id']}")
            edit_products = st.multiselect(
                "Products", CASE_STUDY_PRODUCT_TAGS,
                default=_valid_tags(row.get("products"), CASE_STUDY_PRODUCT_TAGS),
                key=f"slide_vault_edit_products_{row['id']}")
            edit_purpose = st.selectbox(
                "Purpose", VAULT_PURPOSE_TAGS, key=f"slide_vault_edit_purpose_{row['id']}",
                index=VAULT_PURPOSE_TAGS.index(row["purpose"])
                if row.get("purpose") in VAULT_PURPOSE_TAGS else len(VAULT_PURPOSE_TAGS) - 1)
            edit_placement = st.selectbox(
                "Default placement", list(VAULT_PLACEMENT_LABELS),
                format_func=lambda p: VAULT_PLACEMENT_LABELS[p],
                key=f"slide_vault_edit_placement_{row['id']}",
                index=list(VAULT_PLACEMENT_LABELS).index(
                    row["placement"] if row.get("placement") in VAULT_PLACEMENT_LABELS
                    else assembly.VAULT_PLACEMENT_BEFORE_PLAN))
            edit_active = st.checkbox("Active (offered on new proposals)",
                                      value=row.get("active", True),
                                      key=f"slide_vault_edit_active_{row['id']}")
            meta = [f"slide {row['slide_index'] + 1} of {row['filename']}"]
            if row.get("added_by"):
                meta.append(f"added by {row['added_by']}")
            if row.get("date_added"):
                meta.append(str(row["date_added"])[:10])
            meta.append("has an image" if db.slide_vault_render_path(row) == "image"
                        else "pending image render")
            st.caption(" · ".join(meta))

            col_save, col_download = st.columns([1, 2])
            with col_save:
                # Suffix, not prefix -- NON_PERSISTABLE_SUFFIXES matches
                # keys built as "{whatever}_save", not "save_{whatever}".
                if st.button("Save changes", key=f"slide_vault_{row['id']}_save"):
                    _, error = db.update_slide_vault_entry(
                        row["id"], title=edit_title.strip(), summary=edit_summary.strip(),
                        verticals=edit_verticals, products=edit_products,
                        purpose=edit_purpose, placement=edit_placement, active=edit_active)
                    if error:
                        st.error(error)
                    else:
                        st.success("Saved.")
                        st.rerun()
            with col_download:
                render_vault_slide_download(row, key_prefix="slide_vault_browse")


def render_vault_slide_picker(vertical_key, vertical_label):
    """The generate-time checklist for the slide vault. Returns the selected
    rows, each with its own resolved "placement" key.

    Unlike render_case_study_picker, NOTHING is pre-checked -- a vault slide
    is a colleague's own favourite, not vetted proof, so there is no
    pre-check default to protect against a vertical change and therefore no
    key-purge machinery to mirror.
    """
    st.header("Vault slides")
    rows, warning = db.fetch_slide_vault()
    if warning:
        st.warning(f"{warning}. No vault slides can be added to this deck.")
        return []
    if not rows:
        st.caption("The slide vault is empty -- add one from the \"Add vault slide\" page.")
        return []

    matching = [r for r in rows if vertical_key in (r.get("verticals") or [])]
    matching_ids = {r["id"] for r in matching}
    others = [r for r in rows if r["id"] not in matching_ids]

    selected = []

    def _row(row):
        label = row.get("title") or f"{row['filename']} -- slide {row['slide_index'] + 1}"
        picked = st.checkbox(label, value=False, key=f"vault_pick_{row['id']}")
        meta = [row.get("summary") or ""]
        if row.get("added_by"):
            meta.append(f"added by {row['added_by']}")
        st.caption(" · ".join(m for m in meta if m))
        if picked:
            default_placement = (row.get("placement")
                                 if row.get("placement") in VAULT_PLACEMENT_LABELS
                                 else assembly.VAULT_PLACEMENT_BEFORE_PLAN)
            placement = st.selectbox(
                "Where in the deck", list(VAULT_PLACEMENT_LABELS),
                format_func=lambda p: VAULT_PLACEMENT_LABELS[p],
                index=list(VAULT_PLACEMENT_LABELS).index(default_placement),
                key=f"vault_place_{row['id']}")
            selected.append({**row, "placement": placement})

    if matching:
        st.caption(f"{len(matching)} vault slide(s) tagged {vertical_label}.")
        for row in matching:
            _row(row)
    elif vertical_key != "none":
        st.caption(f"No vault slides are tagged {vertical_label} yet.")

    with st.expander(f"Other vault slides ({len(others)}) -- not tagged for this vertical",
                     expanded=False):
        for row in others:
            _row(row)

    if selected:
        st.caption(f"{len(selected)} vault slide(s) will be added where you chose above.")
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


def render_wide_orbit_upload():
    """The Wide Orbit upload-and-parse step, alone -- lives in the intake
    area now (UX sweep, BACKLOG.md), not inside Section C, so a rep can drop
    the file in as soon as they have it rather than hunting for Total TV
    first. Parsing happens on upload rather than at generate time, so a file
    the reader can't handle is reported while the seller is still looking at
    the uploader -- and it degrades to manual entry of the summary numbers
    rather than blocking the proposal.

    The CONFIGURATION this schedule feeds (breakout, description, the
    metrics themselves) stays in `render_wide_orbit_summary`, under Total TV
    in Section C -- that's the only context those numbers mean anything in,
    which is exactly why this split exists instead of moving the whole
    panel: only the upload was ever "a thing a rep has in hand" before
    they've even gotten to Products.
    """
    upload = st.file_uploader("📺 Wide Orbit export", type=["xlsx", "xls", "pdf"],
                              key="wo_upload",
                              help="The Wide Orbit export for this buy — the Campaign Schedule "
                                   "Report (.xlsx), or a Planner (.xls / .pdf). The schedule "
                                   "becomes its own slide, and the broadcast line is added to "
                                   "your media plan automatically. Configure it under Total TV "
                                   "in Products once that's turned on.")
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
        st.caption(f"✅ Loaded — {s.total_spots:,} commercials, ${s.gross_cost:,.0f} gross, "
                   f"read from {schedule.source_name}.")


def render_wide_orbit_summary():
    """The Wide Orbit configuration panel -- metrics, description, breakout
    -- shown under Total TV, the only place these numbers mean anything.
    Reads what `render_wide_orbit_upload` (inline in the setup band, shown
    only once Total TV is checked -- FLOW_REWORK_PLAN.md Phase 1 moved it
    out of the intake area) already parsed; points back up there instead of
    showing an empty panel when nothing's been uploaded yet.
    """
    schedule = st.session_state.get("broadcast_schedule")
    if not schedule:
        if st.session_state.get("wo_error"):
            st.caption("The Wide Orbit file uploaded in the **Setup** band couldn't be read — "
                       "see the error up there, or fill the broadcast line in by hand below.")
        else:
            st.caption("No Wide Orbit schedule uploaded yet — upload one in the **Setup** "
                       "band at the top of the page, or fill the broadcast line in by hand below.")
        return

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
            print(f"[master deck] couldn't open {local_path.name}: {type(exc).__name__}: {exc}")
            st.error("Couldn't open that file as a .pptx -- check that it's a valid, unopened "
                     "PowerPoint file and try again.")
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
            print(f"[usage import] couldn't read {local_path.name}: {exc}")
            st.error("Couldn't read that workbook -- check that it's the expected export "
                     "format and try again.")
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
        "Slide vault",
        "Add vault slide",
        "Update master deck",
        "Update audience usage",
        "Feedback reports",
    ], label_visibility="collapsed", key="page_choice")
    st.sidebar.caption("The finders are also embedded in the proposal flow — "
                       "audiences in Section D2, case studies and vault slides just "
                       "before Generate.")
    st.sidebar.divider()
    render_identity_sidebar()
    # On every page, not just Build -- a rep can hit something worth
    # flagging anywhere in the app.
    render_feedback_popover()
    st.sidebar.caption(f"Build {BUILD_STAMP}")

    standalone = {
        "Proposal history": render_proposal_history,
        "Audience finder": render_audience_finder_page,
        "Zip/map builder": render_zip_map_builder_page,
        "Case study finder": render_case_study_finder,
        "Add case study": render_add_case_study,
        "Slide vault": render_vault_slide_finder,
        "Add vault slide": render_add_vault_slide,
        "Update master deck": render_update_master_deck,
        "Update audience usage": render_update_audience_usage,
        "Feedback reports": render_feedback_admin_page,
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
    # Ahead of EVERYTHING that can resolve a zip -- in particular the
    # intake avails uploader a few lines below, which Phase 6 moved to the
    # very top of the page. install_market_lookup() used to sit down in
    # Section A, which was fine when D2/Section A were the only places
    # avails ever got resolved; once the intake uploader started running
    # ahead of Section A, every avails PDF imported through it (the
    # primary, unconditional entry point since the UX sweep) resolved its
    # zips with NO market lookup registered yet -- geo_resolver.
    # zips_to_markets came back empty every time, silently, on every real
    # document imported this way. Confirmed live: the same PDF through the
    # D2 entry point (which sits after this call) resolved correctly; the
    # only difference was which side of this line it ran on. See
    # DECISIONS.md for the full incident and what else reads
    # resolved_markets. `@st.cache_resource`-idempotent and depends on
    # nothing computed between here and its old spot, so moving it here is
    # a pure reordering.
    _, market_lookup_warning = install_market_lookup()
    if market_lookup_warning:
        st.warning(f"⚠️ {market_lookup_warning}")
    apply_pending_avails_import_fields()
    apply_pending_flight_match_avails()
    apply_pending_draft_fields()
    apply_pending_flight_mirror_edit()

    heading, new_proposal = st.columns([4, 1], vertical_alignment="bottom")
    with heading:
        st.title("Proposal Builder")
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

    # ---------------- Setup band (FLOW_REWORK_PLAN.md Phase 1, restructured
    # in Phase 5, reordered in Phase 6) ----------------
    # Phase 6 puts the avails document FIRST, ahead of every band widget it
    # can fill. An avails document is a document -- it STATES client name,
    # originating market, flight dates and target markets, verbatim,
    # regardless of what the notes say. Importing it before those widgets
    # render is what lets `apply_pending_avails_import_fields` (called at
    # the very top of main(), before this band's first widget) seed them as
    # legal SEED VALUES rather than fighting a widget that's already
    # instantiated -- the same "queue now, apply before anything downstream
    # reads it" shape used everywhere else in this file. The setup band
    # below becomes VALIDATION of what the document supplied, not blank
    # data entry, which is the point:
    #   1. Working from an avails document -- reveals its own uploader
    #      directly beneath it when ticked, filling what it can before
    #      anything else renders
    #   2. Client name, originating market, Plan basis, Total TV (left
    #      column) / Flight start/end + Custom flighting (right column,
    #      unchanged mechanism from Phase 4a)
    #   3. Draft from notes -- LAST. By the time a rep can click Draft,
    #      market/flight/basis are already band inputs (Phase 5, unchanged)
    #      and client name/target markets are already seeded from the avail
    #      if one was uploaded -- drafting supplies INTENT (vertical, which
    #      avails segments to tick, budget, allocation, products), never
    #      re-derives a fact the avail or the rep already supplied. See
    #      apply_draft_to_form's own client_name/target_dma_choice handling
    #      for the seed-if-empty rule this enables.
    # The gate below is UNCHANGED from Phase 1 -- still market + flight
    # dates only, still a hard `return` -- it just now sits below a band
    # that contains more. "Never block an input" holds unconditionally:
    # none of the three toggles gates any OTHER widget, only its own reveal.
    st.header("🧭 Setup")
    st.caption("These drive everything below. Nothing below renders until the "
               "originating market and flight dates are set.")

    avails_mode = st.checkbox(
        "Working from an avails document", value=True, key="avails_mode",
        help="On (the default): the plan is built from the avails table below. Off: build "
             "the media plan directly, same as a proposal with no avails at all.")
    if avails_mode:
        # An avails PDF's own figures are frozen to the DOCUMENT's own
        # dates (Phase 2), so an upload before the band's flight is even
        # set resolves immediately, with nothing left to defer -- and
        # rendering this uploader before client name/market/flight below is
        # what lets a fresh import seed them as this run's SEED VALUES.
        render_avails_pdf_uploader("intake", "The avails PDF you pulled for this buy.")

    band_col1, band_col2 = st.columns(2)
    with band_col1:
        # No ai_section_badge and no on_change here any more -- client_name
        # is a band INPUT now (an avail seeds it, a rep types it), never a
        # drafted write, so it carries no "basics" fill status of its own to
        # show or to clear. See apply_draft_to_form's seed-if-empty handling.
        client_name = st.text_input("Client name", value="Acme Test Co", key="client_name")
        render_logo_upload()
        # Read back rather than returned: Generate, far below, needs both
        # regardless of whether this run touched the uploader at all (a
        # restored proposal's logo, or one uploaded on an earlier run, both
        # live in session_state already).
        uploaded_logo = st.session_state.get("uploaded_logo")
        restored_logo_path = st.session_state.get("restored_logo_path")

        market_choice = st.radio("Originating market", ["DC", "Harrisburg"], horizontal=True,
                                  key="market_choice", on_change=_clear_ai_section, args=("basics",),
                                  help="Which Premion office this proposal comes from -- not where "
                                       "the campaign runs. Target markets are set separately below.")
        avails_basis = st.radio(
            "Plan basis", [AVAILS_BASIS_MONTHLY, AVAILS_BASIS_FLIGHT], horizontal=True,
            key="avails_basis",
            help="How avails and the plan are shown and printed on the targeting slide. "
                 "They're always stored monthly, so switching back and forth changes nothing "
                 "but the presentation.")

        total_tv = st.checkbox("Total TV", value=False, key="total_tv",
                                on_change=_clear_ai_section, args=("products",))
        if total_tv:
            render_wide_orbit_upload()

    # flight_label/flight_shorthand/active_months/n_months are read far below
    # (Section E's grid, the flight-change guard key, the deck payload) --
    # computed once here and left as plain locals for the rest of main() to
    # read. custom_flighting/base_ranges/gen are computed in band_col2 too
    # (the checkbox that decides them renders there), then reused in the
    # full-width block below that actually draws the per-month rows.
    all_months = []
    flight_label = "TBD"
    flight_shorthand = "TBD"
    active_months = []
    n_months = 1
    custom_flighting = False
    base_ranges = []
    gen = st.session_state.get("flight_months_gen", 0)
    with band_col2:
        fcol1, fcol2 = st.columns(2)
        with fcol1:
            flight_start = st.date_input("Flight start", value=None, key="flight_start",
                                          on_change=_clear_ai_section, args=("flight",))
        with fcol2:
            flight_end = st.date_input("Flight end", value=None, key="flight_end",
                                        on_change=_clear_ai_section, args=("flight",))
        ai_section_badge("flight")

        if flight_start and flight_end:
            all_months = month_list(flight_start, flight_end)
            # FLOW_REWORK_PLAN.md Phase 2: one control does both jobs a rep
            # used to need two mechanisms for -- skipping a whole month, or
            # clipping one to part of its span. `flight_month_ranges` is the
            # single owner of the reconciliation (clamp a stale range rather
            # than dropping it; reset to all-active on an empty intersection
            # -- see its own docstring), so this block only has to render it
            # and fold edits back.
            if (st.session_state.get("_flight_month_set") is not None
                    and st.session_state["_flight_month_set"] != all_months):
                bump_flight_months_generation()
            st.session_state["_flight_month_set"] = all_months
            gen = st.session_state.get("flight_months_gen", 0)

            base_ranges = flight_month_ranges(flight_start, flight_end,
                                               st.session_state.get("flight_months"))
            # A rehydrated or drafted proposal can arrive with real custom
            # ranges already in flight_months before this checkbox has ever
            # rendered -- reflect that on first sight rather than showing
            # the box unticked over data that disagrees with it. Only on
            # first sight: once the key exists, a rep's own tick/untick is
            # never fought.
            if "custom_flighting" not in st.session_state and ranges_are_customized(
                    base_ranges, flight_start, flight_end):
                st.session_state["custom_flighting"] = True
            custom_flighting = st.checkbox(
                "Custom flighting -- skip or clip individual months",
                value=False, key="custom_flighting")

    # ---------------- Per-month flighting, full width (FLOW_REWORK_PLAN.md
    # Phase 4a) -- breaks out below both band columns rather than being
    # cramped into band_col2's own half width. Moved here (from Section E,
    # three sections away) so a divergence between the plan's flight and an
    # uploaded avails document's own dates -- surfaced right below -- is
    # known right where the dates are set. Phase 2's one-line band summary
    # caption is GONE, not left alongside this -- it existed only because
    # the control lived elsewhere, and two representations of one fact
    # drift.
    if flight_start and flight_end:
        if custom_flighting:
            st.caption("Each month defaults to its own full calendar span, clipped by the "
                       "flight above. Uncheck a month to drop it from the plan, or edit its "
                       "dates to run it for only part of the month.")
            edited_ranges = []
            for i, entry in enumerate(base_ranges):
                bounds = default_month_range(entry["month"], flight_start, flight_end)
                rcol1, rcol2, rcol3 = st.columns([1.6, 1.2, 1.2])
                with rcol1:
                    row_active = st.checkbox(entry["month"], value=entry["active"],
                                             key=f"fm_active_{gen}_{i}")
                with rcol2:
                    row_start = st.date_input(
                        f"{entry['month']} start", value=entry["start"],
                        min_value=bounds["start"], max_value=bounds["end"],
                        key=f"fm_start_{gen}_{i}", disabled=not row_active,
                        label_visibility="collapsed")
                with rcol3:
                    row_end = st.date_input(
                        f"{entry['month']} end", value=entry["end"],
                        min_value=bounds["start"], max_value=bounds["end"],
                        key=f"fm_end_{gen}_{i}", disabled=not row_active,
                        label_visibility="collapsed")
                if row_end < row_start:
                    # Refuse rather than swap -- a silently swapped range is
                    # a data change a rep can't see. Reverts to this row's
                    # own last-good values instead.
                    st.warning(f"{entry['month']}: end date can't be before start -- left unchanged.")
                    row_start, row_end = entry["start"], entry["end"]
                edited_ranges.append({"month": entry["month"], "start": row_start,
                                      "end": row_end, "active": row_active})

            if not any(r["active"] for r in edited_ranges):
                st.warning("At least one month has to stay active -- reverting to the full flight.")
                edited_ranges = flight_month_ranges(flight_start, flight_end)
                bump_flight_months_generation()
        else:
            # Unticking resets to the plain flight rather than merely hiding
            # the rows underneath a closed checkbox -- keeping stale custom
            # ranges behind a hidden control would be invisible state on a
            # client-facing number. Bumping the generation here (only on a
            # REAL transition away from custom ranges, never on a fresh,
            # never-customized mount) means the fm_* widgets get fresh keys
            # if the rep ticks the box again later, instead of resurrecting
            # whatever they'd previously typed.
            edited_ranges = flight_month_ranges(flight_start, flight_end)
            _prev_ranges = st.session_state.get("flight_months")
            if _prev_ranges is not None and _prev_ranges != edited_ranges:
                bump_flight_months_generation()

        st.session_state["flight_months"] = edited_ranges
        active_months = active_month_labels(edited_ranges)
        st.session_state["active_months"] = active_months
        n_months = max(1, len(active_months))
        # flight_label stays whole-month and is the IDENTITY string --
        # _shared_fields_key, the flight-change guard key and the Wide Orbit
        # coverage warning all key off it, and CLAUDE.md documents what
        # changing that string costs (a mismatched key silently reseeds
        # every option's rows to $0). flight_shorthand is the day-precise
        # DISPLAY string for the plan cell, Campaign Specs and the deck --
        # the two do different jobs and are deliberately never the same
        # variable.
        flight_label = format_flight_label(all_months, active_months) or "TBD"
        flight_shorthand = format_flight_shorthand(edited_ranges) or flight_label
        st.caption(f"{n_months} active month(s): {flight_label}")

        # ---- Avails-flighting divergence, aggregated (Phase 4a (D)/(E)) ----
        # Phase 2's avails_flight_divergence is the ONE mechanism for "plan
        # dates vs avail document dates" -- this is the band's own rendering
        # of it, never a second notice. D2 keeps its per-group detail
        # (Adjust to plan dates / Use document figure) exactly where it is;
        # this is the aggregate, one-click-fix view that belongs where the
        # dates are actually set. Direction is the opposite of D2's button:
        # D2 re-scopes an avails GROUP to the plan; this moves the PLAN to
        # the avails document(s).
        _groups_for_divergence = st.session_state.get("targeting_groups") or []
        _diverging_groups = []
        _doc_starts, _doc_ends = [], []
        for _g in _groups_for_divergence:
            _ds = _parse_iso_date(_g.get("avails_doc_start"))
            _de = _parse_iso_date(_g.get("avails_doc_end"))
            if _ds and _de:
                _doc_starts.append(_ds)
                _doc_ends.append(_de)
            if avails_flight_divergence(_g, edited_ranges):
                _diverging_groups.append(_g)
        if _diverging_groups and _doc_starts and _doc_ends:
            _window_start, _window_end = min(_doc_starts), max(_doc_ends)
            _docs_disagree = len(set(_doc_starts)) > 1 or len(set(_doc_ends)) > 1
            mcol1, mcol2 = st.columns([5, 2])
            with mcol1:
                st.warning(
                    f"{len(_diverging_groups)} avails group(s) run outside the plan's own "
                    f"dates -- the avails document(s) span **{_window_start:%b %d, %Y} – "
                    f"{_window_end:%b %d, %Y}**, the plan runs **{flight_start:%b %d, %Y} – "
                    f"{flight_end:%b %d, %Y}**."
                    + (" Documents disagree with each other on their own span; this button "
                       "uses their combined range." if _docs_disagree else "")
                    + " See the avails table below for which group(s) and month(s).")
            with mcol2:
                if st.button("Match avails flighting", key="match_avails_flighting"):
                    st.session_state["_pending_flight_match_avails"] = {
                        "start": _window_start.isoformat(), "end": _window_end.isoformat()}
                    st.rerun()

    # ---- 3. Draft from notes -- last (Phase 6) ----
    # Every fact the model would otherwise have to infer or extract is
    # already established by the time a rep can reach this: market, flight
    # and plan basis are band inputs (Phase 5); client name and target
    # markets are already seeded from an avails import, if one happened
    # above. Drafting's only job left is turning notes into a plan --
    # vertical, which avails segments to tick, budget, allocation, products.
    # Still ungated, same as every other band toggle: a rep can paste notes
    # and click Draft before market/flight are set, exactly as before.
    draft_from_notes = st.checkbox(
        "Draft from notes", value=True, key="draft_from_notes",
        help="On (the default): paste or upload meeting notes and Claude drafts a first "
             "pass at the whole form. Off: self-serve -- fill in everything below by hand, "
             "no notes box shown.")
    if draft_from_notes:
        st.caption("Paste your meeting or discovery notes here — however rough — and "
                   "Claude fills in most of the form below: budget and media plan, products, "
                   "audiences. The client name, market, flight, plan basis and target markets "
                   "above are given to it as already decided, not guessed at from the notes -- "
                   "an avails document you uploaded above wins any disagreement. Anything it "
                   "wasn't sure about is listed for you to confirm rather than guessed at "
                   "silently. **Everything it fills in is editable, and nothing is final until "
                   "you press Generate.**")
        notes_upload = st.file_uploader("Upload notes instead (.txt, .pdf, .docx)",
                                        type=["txt", "pdf", "docx"], key="notes_upload")
        injected_notes = test_mode_upload("notes_upload_path")
        if injected_notes is not None:
            notes_upload = injected_notes
        # Companion keys deliberately do NOT start with "notes_upload"
        # (unlike the uploader widget itself) -- that prefix is in
        # NON_PERSISTABLE_PREFIXES because Streamlit raises on restoring
        # a file_uploader, and a prefix match would silently take these
        # two plain, ordinary, settable keys down with it, the same trap
        # "wo_loaded_name" (vs. "wo_upload") already sidesteps by naming
        # convention.
        if notes_upload is not None and st.session_state.get("notes_text_loaded") != notes_upload.name:
            target = db.scratch_dir("premion_notes_uploads") / notes_upload.name
            target.write_bytes(notes_upload.getvalue())
            text, extract_error = notes_file_import.extract_notes_text(str(target), notes_upload.name)
            if extract_error:
                st.session_state["notes_text_error"] = extract_error
            else:
                # Written before the text_area below is instantiated
                # THIS run -- a keyed widget's own session_state wins
                # over value= once it exists, so this only ever takes
                # effect on the run that follows a NEW upload (guarded
                # by notes_text_loaded above), never silently
                # overwriting an edit made after the fact.
                st.session_state["notes_text_error"] = None
                st.session_state["draft_notes_input"] = text
            st.session_state["notes_text_loaded"] = notes_upload.name
            # Safe to write a widget-owned key (draft_notes_input) and
            # immediately rerun: st.rerun() builds its RerunData with
            # widget_states=None, so the compact/reassert cycle in
            # on_script_will_rerun() (which is what can shadow a value
            # with a stale widget snapshot) never runs for an internal
            # rerun -- only a rerun carrying real frontend widget data
            # does. That cycle already ran once for THIS pass, before
            # this code executed, so there's no second one left to
            # shadow this write before the text_area below re-registers.
            # See DECISIONS.md's Streamlit mechanics section (verified
            # against Streamlit 1.60.0) for the full trace, including
            # why this doesn't generalize to every st.rerun().
            st.rerun()

        notes_upload_error = st.session_state.get("notes_text_error")
        if notes_upload_error:
            st.error(notes_upload_error)

        notes_input = st.text_area("Meeting / discovery notes", height=180, key="draft_notes_input")
        if st.button("Draft proposal from notes"):
            if not notes_input.strip():
                st.warning("Paste or upload some notes first.")
            else:
                status = st.status(
                    "Drafting from notes -- usually well under a minute, longer for "
                    "open-ended or multi-scenario notes...", expanded=False)
                draft, error = call_claude_draft(
                    notes_input, on_attempt=_draft_attempt_status_updater(status))
                if error:
                    status.update(label="Drafting failed", state="error")
                    st.error(error)
                else:
                    status.update(label="Drafted", state="complete")
                    try:
                        apply_draft_to_form(draft)
                    except Exception as exc:
                        print(f"[draft] apply_draft_to_form failed: {type(exc).__name__}: {exc}")
                        st.error("Couldn't apply the draft to the form. Try drafting again -- "
                                 "if it keeps happening, use Report an issue and I'll take a look.")
                    else:
                        # Remembered so the clarification round can send
                        # the model its own previous answer to revise,
                        # and so it knows which sections the draft
                        # originally owned.
                        st.session_state["draft_source_notes"] = notes_input
                        st.session_state["draft_last_json"] = draft
                        st.session_state["draft_round"] = 1
                        st.session_state["draft_sections_round1"] = set(
                            st.session_state.get("ai_filled_sections", set()))
                        st.rerun()

    # Review list and the clarify round render regardless of the toggle
    # above -- a rep who drafts, then unticks "Draft from notes" to
    # finish the form by hand, still needs to see what the draft left
    # open. Both already no-op when there's nothing to show.
    render_review_list()
    if st.session_state.get("draft_round") == 1 and st.session_state.get("draft_last_json"):
        with st.container(border=True):
            st.markdown("**Clarify and re-draft**")
            st.caption("Answer the open questions above in plain language "
                       "(e.g. \"budget includes the fee, use FIN General In Mkt Shopper\") "
                       "and Claude will revise the draft. Anything you've edited by hand "
                       "since the first draft is kept as you left it. One round.")
            clarifications = st.text_area("Your answers", height=110, key="draft_clarifications")
            if st.button("Re-draft with these answers"):
                if not clarifications.strip():
                    st.warning("Answer at least one of the open questions first.")
                else:
                    # Sections the first draft filled but that have
                    # since been cleared by an on_change -- i.e. the
                    # user has edited them by hand, so the re-draft must
                    # not overwrite them.
                    edited_since = (st.session_state.get("draft_sections_round1", set())
                                    - st.session_state.get("ai_filled_sections", set()))
                    status = st.status(
                        "Re-drafting with your clarifications -- usually well under a "
                        "minute...", expanded=False)
                    draft, error = call_claude_redraft(
                        st.session_state.get("draft_source_notes", ""),
                        st.session_state["draft_last_json"],
                        clarifications,
                        on_attempt=_draft_attempt_status_updater(status),
                    )
                    if error:
                        status.update(label="Re-drafting failed", state="error")
                        st.error(error)
                    else:
                        status.update(label="Re-drafted", state="complete")
                        try:
                            apply_draft_to_form(draft, skip_sections=edited_since)
                        except Exception as exc:
                            print(f"[draft] apply_draft_to_form (redraft) failed: "
                                  f"{type(exc).__name__}: {exc}")
                            st.error("Couldn't apply the re-draft to the form. Try again -- if "
                                     "it keeps happening, use Report an issue and I'll take a look.")
                        else:
                            st.session_state["draft_last_json"] = draft
                            st.session_state["draft_round"] = 2
                            st.rerun()

    _setup_missing = []
    if market_choice not in ("DC", "Harrisburg"):
        _setup_missing.append("the originating market")
    if not (flight_start and flight_end):
        _setup_missing.append("the flight dates")
    if _setup_missing:
        st.info(f"Set {' and '.join(_setup_missing)} above to continue.")
        return

    # ---------------- Section A: Client basics ----------------
    # FLOW_REWORK_PLAN.md Phase 5: client_name moved into the setup band
    # (rendered above, before the gate) alongside originating market. Phase 6
    # went further -- client_name is a band INPUT now, seeded by an avails
    # import or typed by hand, never a drafted write (see
    # apply_draft_to_form) -- so this section is genuinely target markets +
    # vertical only, and the badge below describes exactly those two fields.
    st.header("A. Client basics")
    ai_section_badge("basics")
    market_profile_rows, market_profile_warning = load_market_profiles()
    # install_market_lookup() itself now runs at the very top of main() --
    # see that call site's own comment for why (the intake avails uploader,
    # ABOVE this section since Phase 6, needs it registered before IT
    # resolves any zips). sync/autofill stay here, unchanged: they depend on
    # market_profile_rows (just loaded above) and belong immediately before
    # target_dma_choice renders, which is still true regardless of where
    # the lookup table itself gets installed.
    sync_targeting_groups()
    apply_group_markets_autofill(market_profile_rows)
    target_dmas, include_market_profile = market_profile_picker(
        market_profile_rows, market_profile_warning)
    vertical_choice = st.selectbox("Vertical", list(VERTICALS.keys()), index=0, key="vertical_choice",
                                    on_change=_clear_ai_section, args=("basics",))
    # The agency gross-up moved to a rep-only checkbox beside the plan table
    # (FLOW_REWORK_PLAN.md Phase 4b) -- it's no longer a Client basics field,
    # and drafting never sets it, so it no longer needs an ai_section_badge
    # or a "basics"-section on_change either.

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
        # The Total TV checkbox itself now lives in the setup band (Phase
        # 1) -- `total_tv` is already set from there by the time this run
        # reaches here. This is just its configuration panel, which is
        # still only meaningful in the context of the rest of Products.
        if total_tv:
            st.caption("Total TV is on (see the setup band above).")
            # Indented directly beneath its own caption, so it reads as
            # part of Total TV rather than as a panel floating between two
            # unrelated products.
            _, nested = st.columns([0.05, 0.95])
            with nested:
                render_wide_orbit_summary()
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
    # avails_basis is the setup band's "Plan basis" control -- already set,
    # unconditionally, before this code runs. avails_label/avails_months
    # are computed here regardless of whether D2 itself is shown, since the
    # generate block below reads them unconditionally.
    _, avails_active_months_default = form_flight_months()
    avails_months = max(1, len(avails_active_months_default))
    avails_label = avails_column_label(avails_basis, avails_months)
    if include_avails_template:
        st.header("D2. Audiences & avails")
        ai_section_badge("avails")
        if "avails_version" not in st.session_state:
            st.session_state["avails_version"] = 0

        # FLOW_REWORK_PLAN.md Phase 3: a local Monthly/Full-flight override
        # for THIS section only, independent of the setup band's own Plan
        # basis -- safe because Phase 2 made dates authoritative, so the two
        # views are two renderings of one underlying set of numbers, not two
        # independent sets. `avails_section_basis` is None (follow the band,
        # the default -- ticking this control changes nothing until a rep
        # actually picks a different basis) or one of the same
        # AVAILS_BASIS_* constants the band itself uses. Conversion runs
        # through the same avails_to_display/avails_from_display every other
        # basis-aware read/write already uses -- no new conversion code.
        # Deliberately scoped to THIS grid's own display, editor key, caption
        # and fold-back: `avails_label`/`avails_basis` (the band's own,
        # computed above) keep feeding form_json and the deck/targeting-slide
        # payload below unchanged -- a client-facing deck reflects the plan's
        # stated basis, never a rep's momentary "let me peek at this the
        # other way" toggle.
        avails_section_basis = st.radio(
            "This section's basis", [None, AVAILS_BASIS_MONTHLY, AVAILS_BASIS_FLIGHT],
            format_func=lambda v: f"Follow the band ({avails_basis})" if v is None else v,
            horizontal=True, key="avails_section_basis",
            help="Changes how this table displays and how avails are typed here -- never "
                 "what's stored, and never the plan basis (set in the setup band above) that "
                 "the generated deck actually uses.")
        avails_basis_effective = avails_section_basis or avails_basis
        avails_label_section = avails_column_label(avails_basis_effective, avails_months)

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

        # A queued "apply to audience" color click (see the expander below
        # the grid) has to land before the table below reads groups this
        # run, same reason the market autofill syncs before reading groups.
        apply_pending_color_cascade()
        # Same reason, for "Add all to plan"/"Clear plan lines"/the Remove-
        # or-Keep confirmation (see the buttons below the grid).
        apply_pending_group_include_all()
        apply_pending_group_include_clear()
        apply_pending_include_confirm()
        # Same reason, for the divergence panel's "Adjust to plan dates" /
        # "Use document figure" buttons below (FLOW_REWORK_PLAN.md Phase 2).
        apply_pending_avails_plan_adjust()
        # Same reason, for the entity grouping expander's "Group selected
        # rows"/"Ungroup" buttons below (FLOW_REWORK_PLAN.md Phase 3).
        apply_pending_entity_group()
        apply_pending_entity_ungroup()

        # A group edit made just above (the market autofill) has to reach
        # the table this run, not next -- sync BEFORE reading groups, not
        # only in Section E. (Phase 6 adds a still-earlier call, in Section
        # A ahead of the market picker; this is the D2 one that call's own
        # comment refers to as already existing.)
        sync_targeting_groups()
        groups = st.session_state.get("targeting_groups") or []

        # FLOW_REWORK_PLAN.md Phase 2: the uploaded avail is a real quote and
        # holds at it until a rep explicitly says otherwise -- this panel is
        # where the PER-GROUP divergence detail surfaces (the setup band, up
        # above, now owns the aggregate one-click "Match avails flighting"
        # view of the same fact -- Phase 4a). flight_months is computed in
        # the band, earlier in this same run, so there is no render lag here
        # any more. Only a group with real periods (an avails-PDF import)
        # can diverge at all; a hand-typed or drafted avails row has none
        # and is silently skipped.
        if flight_start and flight_end:
            _divergence_ranges = flight_month_ranges(
                flight_start, flight_end, st.session_state.get("flight_months"))
            for group in groups:
                mismatches = avails_flight_divergence(group, _divergence_ranges)
                if not mismatches:
                    continue
                glabel = f"{tg.audience_label(group)} / {tg.geo_label(group, label_for=_market_display_name)}"
                note = format_avails_divergence_note(glabel, mismatches)
                dcol1, dcol2 = st.columns([5, 2])
                with dcol1:
                    adjusted = bool(group.get("avails_adjusted_to_plan"))
                    st.warning(note + (" Currently adjusted to the plan's own dates."
                                       if adjusted else
                                       " Showing the avails document's own figure."))
                with dcol2:
                    # Both buttons render on EVERY pass, toggling `disabled=`
                    # rather than which one appears -- a widget key that's
                    # sometimes rendered and sometimes skipped between runs
                    # (this group flips between "adjusted" and "not" across
                    # reruns) trips Streamlit's own "can't set a widget's
                    # value via session_state" policy the moment it comes
                    # back after a run where it was absent. Found live
                    # writing the adjust-then-revert test.
                    # Same reason every other D2 queued action calls
                    # st.rerun() immediately (see "Add all to plan" above):
                    # apply_pending_avails_plan_adjust() runs BEFORE this
                    # panel, at the top of D2 -- queuing here and waiting
                    # for the natural next rerun would show a stale figure
                    # for one render, and in AppTest a single click().run()
                    # would show no change at all.
                    if st.button("Adjust to plan dates", key=f"avails_adjust_{group['id']}",
                                 disabled=adjusted):
                        pending = dict(st.session_state.get("_pending_avails_plan_adjust") or {})
                        pending[group["id"]] = True
                        st.session_state["_pending_avails_plan_adjust"] = pending
                        st.rerun()
                    if st.button("Use document figure", key=f"avails_use_doc_{group['id']}",
                                 disabled=not adjusted):
                        pending = dict(st.session_state.get("_pending_avails_plan_adjust") or {})
                        pending[group["id"]] = False
                        st.session_state["_pending_avails_plan_adjust"] = pending
                        st.rerun()

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
            group["id"]: avails_to_display(group.get("avails_monthly", 0), avails_basis_effective, avails_months)
            for group in groups
        }
        # Same "what did this cell show before the edit" capture as avails
        # itself, for the same reason -- the fold-back below can only tell a
        # real tick/untick apart from an untouched cell by comparing against
        # what was actually shown, not by re-deriving from `groups` (which
        # this SAME run may have already changed via apply_pending_color_
        # cascade or the market autofill above).
        include_shown_before_by_gid = {
            group["id"]: bool(group.get("include_in_plan")) for group in groups
        }
        default_rows = []
        for group in groups:
            row = {"gid": group["id"], "Plan": bool(group.get("include_in_plan")),
                  "Audience": tg.audience_label(group),
                  "Markets": _group_markets(group),
                  # FLOW_REWORK_PLAN.md Phase 3: "Label" is now the ENTITY
                  # this row is for (e.g. "Toyota of Annapolis") -- the
                  # column that used to carry this name (the Geo-cell
                  # override, `group["name"]`) is renamed "Geo Label" right
                  # below, UNCHANGED in behavior. The two are deliberately
                  # separate fields: an entity can span two avails rows with
                  # two different Geo Labels (a 5mi and a 10mi radius tier
                  # for one dealership), so reusing one field for both jobs
                  # doesn't survive that real case -- see
                  # targeting_groups.py's module docstring.
                  "Label": tg.entity_label_of(group),
                  "Geo Label": tg.geo_label(group, label_for=_market_display_name),
                  "Color": _color_swatch_label(group.get("color")),
                  # Purely informational -- derived fresh from color_locked
                  # every render, never read back in the fold-back below.
                  # Without this a rep has no way to tell, just by looking
                  # at the table, WHY a row didn't move when they pushed a
                  # new color out to the rest of its audience.
                  "Detached": "\U0001F512" if group.get("color_locked") else "",
                  # Purely informational, same as Detached -- the avails
                  # document's own quoted window, never round-tripped by the
                  # fold-back below. A rep can't judge a divergence warning
                  # above without seeing both date sets (FLOW_REWORK_PLAN.md
                  # Phase 2). Blank for a hand-typed/drafted row with no
                  # document behind it.
                  "Avail dates": (_daterange_text(_parse_iso_date(group["avails_doc_start"]),
                                                  _parse_iso_date(group["avails_doc_end"]))
                                 if group.get("avails_doc_start") and group.get("avails_doc_end")
                                 else "")}
            row[avails_label_section] = shown_before_by_gid[group["id"]]
            default_rows.append(row)

        # Sorting is a VIEW onto `groups`, never a rewrite of it -- the real
        # order (and so the plan lines a group seeds) stays exactly what
        # `groups` already holds; only what's DISPLAYED here is reordered.
        # `num_rows="dynamic"` (needed for the grid's own +/- row add/delete)
        # disables st.data_editor's native header-click sort outright, so
        # this is a small control instead, defaulting to audience-major then
        # Label -- the same hierarchy the plan table and the map legend use.
        # FLOW_REWORK_PLAN.md Phase 3: "Label" here is the entity column
        # (see default_rows above); "Geo Label" is the renamed former Label
        # column, given its own explicit sort options so neither is stranded.
        avails_sort_options = {
            "Audience, then Label (default)": None,
            "Audience (A -> Z)": ("Audience", False),
            "Label (A -> Z)": ("Label", False),
            "Geo Label (A -> Z)": ("Geo Label", False),
            f"{avails_label_section} (High -> Low)": (avails_label_section, True),
            f"{avails_label_section} (Low -> High)": (avails_label_section, False),
        }
        sort_choice = st.selectbox(
            "Sort table", list(avails_sort_options), key="avails_sort_choice",
            help="Changes how the rows are shown here only. The groups' real order -- and "
                 "the order of the plan lines and targeting slide they seed -- never changes.")
        sort_spec = avails_sort_options[sort_choice]
        display_rows = list(default_rows)
        if sort_spec is None:
            display_rows.sort(key=lambda r: (str(r["Audience"]), str(r["Label"])))
        else:
            sort_col, descending = sort_spec
            display_rows.sort(key=lambda r: r[sort_col], reverse=descending)

        default_avails = (pd.DataFrame(display_rows) if display_rows
                          else pd.DataFrame(columns=["gid", "Plan", "Audience", "Markets", "Label",
                                                     "Geo Label", avails_label_section, "Color",
                                                     "Detached", "Avail dates"]))
        # The basis and the sort choice are both part of the editor key: a
        # data_editor handed a differently-ordered (or differently-schemaed)
        # frame under the SAME key keeps rendering its own prior value
        # instead -- the same "a keyed widget's session_state entry beats
        # its value= argument" trap as every other replaceable-state widget
        # in this file. Remounting on a sort change is safe because nothing
        # is lost: the real data lives in `targeting_groups`, re-derived
        # into `display_rows` fresh on every run.
        sort_key_part = list(avails_sort_options).index(sort_choice)
        avails_editor_key = (f"avails_editor_{st.session_state['avails_version']}"
                             f"_{'flight' if avails_basis_effective == AVAILS_BASIS_FLIGHT else 'monthly'}"
                             f"_sort{sort_key_part}")
        market_options = sorted({m for group in groups for m in _group_markets(group)} | set(target_labels))
        avails_df = st.data_editor(
            default_avails, num_rows="dynamic", key=avails_editor_key, use_container_width=True,
            on_change=_clear_ai_section, args=("avails",),
            disabled=["Detached", "Avail dates"],
            column_config={
                # Hidden, not shown to the rep -- the group id a row is
                # backed by, the same round-tripping hidden-column mechanic
                # `_group_ids` uses on the media plan grid. `None` hides it
                # but still returns it through an edit, a delete or a new row.
                "gid": None,
                # The whole point of this change: the avails table is
                # research by default, and a group contributes a Premion
                # Streaming TV line only once this is ticked. First visible
                # column, so it reads as the question every other column's
                # answer depends on.
                "Plan": st.column_config.CheckboxColumn(
                    "Plan", default=False,
                    help="Tick the audiences this proposal is actually selling. Each ticked "
                         "row becomes one Premion Streaming TV line on the media plan below. "
                         "Everything on this table stays on the targeting slide either way -- "
                         "an unticked row is inventory you're showing, not a line you're "
                         "quoting."),
                "Markets": st.column_config.MultiselectColumn(
                    "Markets", options=market_options, accept_new_options=True,
                    help="One row can span several markets -- add more than one here "
                         "for a group that sells as a single campaign line."),
                # FLOW_REWORK_PLAN.md Phase 3: the entity this row is FOR --
                # e.g. "Toyota of Annapolis" -- distinct from the Geo Label
                # column right below. Several rows can share one entity
                # (grouped via the expander under the table); editing this
                # cell renames the WHOLE shared entity, never just this one
                # row -- see the fold-back below. Blank is the default and
                # valid state for a single-entity proposal; nothing downstream
                # behaves differently until a rep actually groups rows.
                "Label": st.column_config.TextColumn(
                    "Label", help="The real-world thing this row is for (e.g. \"Toyota of "
                                  "Annapolis\") -- optional. Several rows can share one entity "
                                  "(group them below the table); editing this renames the whole "
                                  "shared entity, not just this row. Leave blank for an ordinary, "
                                  "single-entity proposal -- nothing changes until you group rows."),
                # The renamed former "Label" column -- same field
                # (`group["name"]`), same job as always: the Geo cell
                # override. Same text input as the geo-definition expander's
                # "Geo label (optional)" field -- this is just a second place
                # to see and edit it, since several Zips-mode rows resolving
                # to the same market otherwise all show as identical "Saint
                # Louis" Markets chips with no way to tell them apart at a
                # glance. Defaults to the auto-derived summary (market plus
                # zip count) exactly like the plan table's own Geo cell.
                "Geo Label": st.column_config.TextColumn(
                    "Geo Label", help="What the plan table's Geo column and the targeting slide "
                                  "show for this line -- market plus zip count by default. Edit to "
                                  "relabel; clear it to go back to the derived summary."),
                # A real swatch would be `st.column_config.ColorColumn`, but
                # no released Streamlit version has one (a long-open feature
                # request, not something declined here) -- a SelectboxColumn
                # over the app's own fixed, ordered palette is the closest
                # real substitute, and true to how colors already work here
                # (never an arbitrary RGB pick). Picking a color here is
                # always a SINGLE-ROW edit -- see the fold-back below -- it
                # detaches this one group from its audience's shared color
                # the same way editing Label detaches a geo override.
                "Color": st.column_config.SelectboxColumn(
                    "Color", options=list(_COLOR_SWATCH_LABELS.values()),
                    help="Every group under one audience shares a color by default. Picking one "
                         "here breaks just THIS row out on its own -- it won't move again even if "
                         "the audience's shared color changes later. Use \"Apply a color to a whole "
                         "audience\" below the table to push a color back out to every row still "
                         "following the shared one."),
                # Read-only (see `disabled=` above) -- a plain fact about
                # THIS row (does it follow its audience's color right now),
                # never an input. Column header is the marker itself, not a
                # word, so it reads at a glance rather than competing with
                # Audience/Label for space.
                "Detached": st.column_config.TextColumn(
                    "\U0001F512", help="This row's color was set by hand and won't move if the "
                                       "audience's shared color changes -- \"Apply a color to a whole "
                                       "audience\" below is the only way to bring it back in.",
                    width="small"),
                # Read-only (see `disabled=` above) -- the avails document's
                # own quoted window, FLOW_REWORK_PLAN.md Phase 2. Blank for a
                # hand-typed/drafted row with no document behind it.
                "Avail dates": st.column_config.TextColumn(
                    "Avail dates", help="The window this row's avails figure was actually quoted "
                                        "for, straight off the avails document -- blank for a "
                                        "hand-typed or drafted row. A mismatch against the plan's "
                                        "own dates shows as a warning above, with an option to "
                                        "adjust this figure to the plan's dates instead."),
            },
        )
        avails_df[avails_label_section] = avails_df[avails_label_section].fillna(0)
        # The MultiselectColumn new-row trap: a row added via the grid's own
        # "+" comes back with Markets = None, not [] -- normalize on read,
        # same discipline group_ids_of uses for _group_ids.
        avails_df["Markets"] = avails_df["Markets"].apply(lambda v: list(v) if isinstance(v, list) else [])
        # Same trap, boolean-flavored: a brand-new row's CheckboxColumn comes
        # back as None too, not False.
        avails_df["Plan"] = avails_df["Plan"].fillna(False).astype(bool)
        total_avails_val = int(avails_df[avails_label_section].sum())
        st.caption(f"Total avails ({'full flight' if avails_basis_effective == AVAILS_BASIS_FLIGHT else 'monthly'}): "
                   f"{total_avails_val:,}")

        # Fold back into groups, keyed on gid -- NOT position, which a merge
        # of add/delete/reorder can't be trusted to preserve. A row missing
        # its gid (added via the grid's own "+") gets a fresh one from
        # new_group; a group whose gid isn't among the returned rows was
        # deleted and simply isn't carried into new_groups.
        groups_by_gid = {group["id"]: group for group in groups}
        new_groups = []
        # An uncheck that would discard a hand-edited plan line is queued
        # here for the confirmation panel below the grid, never applied and
        # then undone -- see include_removal_blocked.
        pending_blocked = {}
        # FLOW_REWORK_PLAN.md Phase 3: {entity_id: new_label} for any REAL
        # entity-Label edit found this loop -- applied to every group
        # sharing that entity_id in the propagation pass right after this
        # loop, since a rename touches the whole shared entity, not just the
        # one row a rep happened to edit.
        entity_renames = {}
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
            # FLOW_REWORK_PLAN.md Phase 3: this is the "Geo Label" cell now
            # (renamed from "Label" -- see default_rows above), still the
            # exact same field (`name`) and exact same logic, untouched.
            label_text = str(row.get("Geo Label", "") or "").strip()
            label_unchanged = prior is not None and _cell_unchanged(
                tg.geo_label(prior, label_for=_market_display_name), label_text)
            name = (prior.get("name", "") if prior else "") if label_unchanged else label_text

            # FLOW_REWORK_PLAN.md Phase 3: a fifth fold-back field, on the
            # NEW "Label" cell -- the entity this row is FOR, a separate
            # field from `name`/Geo Label above (see default_rows' own
            # comment). Same test, same shape: an unchanged cell keeps
            # whatever entity_label/entity_id/entity_locked this group
            # already had, byte-for-byte; a real edit is a RENAME of the
            # whole shared entity, not just this one row -- `entity_renames`
            # queues it here and the propagation pass right after this loop
            # (before `new_groups` is written to session_state) pushes the
            # new text onto every group sharing this row's `entity_id`.
            # `entity_id` itself is never touched here -- renaming must never
            # regroup; grouping/ungrouping is its own explicit action below
            # the table.
            entity_label_text = str(row.get("Label", "") or "").strip()
            entity_label_unchanged = prior is not None and _cell_unchanged(
                tg.entity_label_of(prior), entity_label_text)
            if entity_label_unchanged:
                entity_id = prior.get("entity_id") or (prior.get("id") if prior else None)
                entity_label = prior.get("entity_label", "")
                entity_locked = bool(prior.get("entity_locked"))
            else:
                entity_id = (prior.get("entity_id") or prior.get("id")) if prior else None
                entity_label = entity_label_text
                entity_locked = True
                if prior is not None and entity_id:
                    entity_renames[entity_id] = entity_label

            # Same fold-back test a third time, on Color: the cell can only
            # DISPLAY a color as one of the fixed swatch labels, never round-
            # trip whether this group is following its audience's cascade or
            # was deliberately broken out of it -- an untouched cell keeps
            # both the color AND the lock byte-for-byte. A REAL edit is
            # always single-row and always locks (see the SelectboxColumn's
            # own help text): a plain pick from this dropdown is exactly the
            # "deliberate, permanent" action the Label column's override
            # already models, never a hint to touch any other row.
            color_label_now = str(row.get("Color", "") or "").strip()
            color_unchanged = prior is not None and _cell_unchanged(
                _color_swatch_label(prior.get("color")), color_label_now)
            if color_unchanged:
                color = prior.get("color")
                color_locked = bool(prior.get("color_locked"))
            elif prior is not None:
                color = _color_from_swatch_label(color_label_now)
                color_locked = True
            else:
                # Brand-new row (the grid's own "+") -- starts on its
                # audience's current shared color, same as a new group built
                # anywhere else in the app. `groups + new_groups`, not just
                # `groups`: two new rows for the SAME new audience added in
                # one batch (before a rerun) must land on the same color,
                # not each roll its own round-robin swatch.
                color = _color_for_new_group(groups + new_groups, terms, op)
                color_locked = False

            # Same fold-back test a fifth time, on Plan: the cell can only
            # DISPLAY whether a group is on the plan right now, never round-
            # trip WHO decided that -- an untouched cell carries BOTH
            # include_in_plan and include_locked forward byte-for-byte. A
            # real edit always locks, the same "deliberate, permanent" rule
            # Color's own SelectboxColumn already models -- and this is the
            # highest-risk line in the whole change: skipping it would wipe
            # every group's inclusion on every single render.
            include_now = bool(row.get("Plan"))
            include_unchanged = prior is not None and _cell_unchanged(
                include_shown_before_by_gid.get(gid), include_now)
            if include_unchanged:
                include_in_plan = bool(prior.get("include_in_plan"))
                include_locked = bool(prior.get("include_locked"))
            elif prior is not None:
                include_in_plan, include_locked = include_now, True
            else:
                # Brand-new row (the grid's own "+", or a rep typing a fresh
                # row and ticking Plan in the SAME submit) -- there's no
                # "before" to compare against, so whatever the checkbox
                # shows right now IS the rep's real choice. Left unticked
                # (the common case: audience/geo typed first, Plan touched
                # later or not at all), this is exactly today's safe
                # default and stays UNLOCKED, so a later "Add all" or a
                # draft's pre-selection can still pick it up automatically;
                # ticked in the same submit is as deliberate as any other
                # edit and locks the same way.
                include_in_plan, include_locked = include_now, include_now

            # Unchecking a group whose plan line has been hand-edited is
            # never silent. Refuse the uncheck here (the group stays
            # included) and queue it for the confirmation panel below the
            # grid to resolve -- applying the removal and then undoing it
            # after the fact would mean the edited line briefly didn't
            # exist, exactly the "discard first, ask later" shape this
            # mechanism exists to avoid. A group that was never included, or
            # whose row was never edited, is never queued -- see
            # include_removal_blocked.
            if prior is not None and prior.get("include_in_plan") and not include_in_plan:
                blocked = include_removal_blocked(gid, st.session_state.get("plan_options") or [])
                if blocked:
                    pending_blocked[gid] = blocked
                    include_in_plan, include_locked = True, bool(prior.get("include_locked"))

            monthly = restore_untouched_avails(
                prior.get("avails_monthly", 0) if prior else 0,
                shown_before_by_gid.get(gid), row[avails_label_section], avails_basis_effective, avails_months)
            built = tg.new_group(
                terms, op=op, geo_def=geo_def,
                name=name,
                avails_monthly=monthly,
                color=color, color_locked=color_locked,
                group_id=gid,
                include_in_plan=include_in_plan, include_locked=include_locked,
                entity_id=entity_id, entity_label=entity_label, entity_locked=entity_locked,
            )
            built["resolved_zips"] = resolved_zips
            built["resolved_markets"] = resolved_markets
            # Same fold-back test as the avails figure itself, one field
            # over: `tg.new_group` has no notion of `avails_full_flight` (an
            # avails-PDF import's exact daily-rate total -- see
            # avails_full_flight_from_daily_rate), so building a fresh group
            # here on every render silently dropped it on the very next
            # rerun after every real import, which is why it read back as
            # None the moment the D2 grid itself rendered. There is no cell
            # for this field to compare "before" vs "after" against -- it's
            # never shown -- so the only sound signal is whether the Max
            # Monthly Avails CELL itself was actually edited: an untouched
            # cell carries the exact total forward byte for byte; a real
            # edit means the rep just overrode the imported figure with
            # their own number, so the precise total no longer describes
            # anything real and is dropped, falling back to the same
            # monthly*n_months approximation an ordinary hand-typed row
            # already uses.
            shown_before_raw = shown_before_by_gid.get(gid)
            shown_before_num = None if shown_before_raw is None else int(shown_before_raw or 0)
            shown_after_num = int(float(str(row[avails_label_section]).replace(",", "") or 0))
            avails_cell_unchanged = prior is not None and _cell_unchanged(
                shown_before_num, shown_after_num)
            built["avails_full_flight"] = (
                prior.get("avails_full_flight") if avails_cell_unchanged else None)
            # The document-quote metadata (FLOW_REWORK_PLAN.md Phase 2) rides
            # the SAME fold-back test as avails_full_flight above -- a hand-
            # edit that overrides the imported figure means this row is now
            # an ordinary typed number, not a document's quote to hold or
            # adjust, so all of it drops together rather than leaving a
            # frozen "document" figure sitting behind a number the rep just
            # typed over it.
            if avails_cell_unchanged and prior is not None:
                for _doc_field in ("avails_doc_start", "avails_doc_end", "avails_doc_monthly",
                                   "avails_doc_full_flight", "avails_adjusted_to_plan"):
                    if _doc_field in prior:
                        built[_doc_field] = prior[_doc_field]
            # Same fold-back test a fourth time, on the D2 placeholder marker
            # (avails_rows_for_markets/seed_rows_to_groups): the cell can
            # only DISPLAY a blank Audience, never distinguish "still the
            # untouched starter row" from "a rep cleared a real audience
            # back to blank" -- so an unchanged Audience cell carries the
            # marker forward from `prior`, and a REAL edit (a rep typed
            # something in) clears it. Without this, `tg.new_group` (called
            # directly here, not through seed_rows_to_groups) never sets the
            # marker at all, and it was gone the instant this grid rendered
            # once -- before an avails import or a finder "Add" ever got a
            # chance to see it and drop it.
            if audience_unchanged and prior is not None and prior.get("_placeholder"):
                built["_placeholder"] = True
            new_groups.append(built)
        # `new_groups` is in whatever order the display sort above put it
        # in -- restore the groups' own stored order before this becomes the
        # new `targeting_groups`, so sorting the table never reorders the
        # real list (a brand-new row, added via the grid's own "+", has no
        # prior position and keeps its place among the other new rows,
        # appended after every pre-existing group).
        _original_position = {group["id"]: i for i, group in enumerate(groups)}
        new_groups.sort(key=lambda grp: _original_position.get(grp["id"], len(_original_position)))
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
        # FLOW_REWORK_PLAN.md Phase 3: propagate any real entity-Label rename
        # (queued above, keyed by entity_id) onto EVERY group sharing that
        # entity_id -- including ones this render's loop already visited and
        # left alone because THEIR OWN Label cell looked unchanged. Without
        # this, renaming a shared entity via just one of its rows would only
        # ever stick to that one row; every sibling would keep showing the
        # old text until its own row happened to be edited too.
        if entity_renames:
            for g in new_groups:
                eid = tg.entity_id_of(g)
                if eid in entity_renames:
                    g["entity_label"] = entity_renames[eid]
                    g["entity_locked"] = True
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
        if pending_blocked:
            existing_pending = st.session_state.get("_pending_include_removal") or {}
            st.session_state["_pending_include_removal"] = {**existing_pending, **pending_blocked}
        if pending_blocked or len(new_groups) != len(groups):
            st.session_state["avails_version"] = st.session_state.get("avails_version", 0) + 1
            st.rerun()

        # Bulk alternatives to ticking one row at a time. Both go through
        # the SAME queue-then-apply-next-run shape as a single Plan edit
        # (apply_pending_group_include_all/_clear, called before this
        # section reads groups) rather than mutating `new_groups` in place
        # here -- this run has already built the grid and the plan from the
        # PRE-click state, so a same-run mutation would show a stale grid
        # for one render.
        include_count = sum(1 for g in new_groups if g.get("include_in_plan"))
        bcol1, bcol2, bcol3 = st.columns([2, 2, 5])
        with bcol1:
            if st.button("Add all to plan",
                         key=f"avails_include_all_{st.session_state['avails_version']}",
                         disabled=not new_groups):
                st.session_state["_pending_group_include_all"] = True
                st.rerun()
        with bcol2:
            if st.button("Clear plan lines",
                         key=f"avails_include_clear_{st.session_state['avails_version']}",
                         disabled=not include_count):
                st.session_state["_pending_group_include_clear"] = True
                st.rerun()
        with bcol3:
            st.caption(f"{include_count} of {len(new_groups)} audience(s) on the plan.")

        # The uncheck-with-edits confirmation -- one panel covering every
        # group a grid submit or "Clear plan lines" just tried to take off
        # the plan and couldn't, because a rep's own numbers were on the
        # line. Never applied silently; Confirm and Cancel are both here.
        pending_removal = st.session_state.get("_pending_include_removal")
        if pending_removal:
            groups_by_gid_now = {g["id"]: g for g in new_groups}
            st.warning("Taking these audiences off the plan discards lines you've edited:")
            for gid, blocked_rows in pending_removal.items():
                group = groups_by_gid_now.get(gid)
                audience_name = tg.audience_label(group) if group else "(removed audience)"
                for item in blocked_rows:
                    opt_suffix = f" -- {item['option']}" if item.get("option") else ""
                    st.caption(f"**{audience_name}**{opt_suffix}: {item['tactic']} ({item['geo']}), "
                              f"{item['impressions']:,.0f} impressions / ${item['cost']:,.0f}")
            rcol1, rcol2, _rcol3 = st.columns([2, 2, 5])
            with rcol1:
                if st.button("Remove them and discard those numbers",
                             key=f"avails_include_confirm_remove_{st.session_state['avails_version']}"):
                    st.session_state["_pending_include_confirm"] = "remove"
                    st.rerun()
            with rcol2:
                if st.button("Keep them on the plan",
                             key=f"avails_include_confirm_keep_{st.session_state['avails_version']}"):
                    st.session_state["_pending_include_confirm"] = "keep"
                    st.rerun()

        for group in new_groups:
            label = tg.audience_label(group)
            if label.strip():
                avails_rows.append({
                    "audience": label,
                    "geo": tg.geo_label(group, label_for=_market_display_name),
                    # What the client sees -- deliberately the BAND's own
                    # basis (`avails_basis`), not this section's own display
                    # override (`avails_basis_effective`, FLOW_REWORK_PLAN.md
                    # Phase 3): a client-facing deck reflects the plan's
                    # stated basis, never a rep's momentary "let me peek at
                    # this the other way" toggle on the D2 grid above. The
                    # two agree whenever the section override is left at its
                    # default (None, follow the band) -- only a rep who's
                    # deliberately switched THIS section's own view sees them
                    # diverge, and only in the app, never in the deck.
                    "avails": f"{avails_to_display(group['avails_monthly'], avails_basis, avails_months):,}",
                    "avails_monthly": group["avails_monthly"],
                })

        catalog_rfp_lookup = dict(zip(audience_catalog["segment"], audience_catalog["rfp_selectable"]))
        custom_count = tg.custom_segment_count(new_groups, catalog_rfp_lookup)
        if custom_count > 1:
            st.warning(f"{custom_count} custom (non-RFP-selectable) audiences are in play across your "
                       f"targeting groups above -- only one is allowed per campaign. Review before generating.")

        # The explicit cascade action -- the real substitute for the
        # mockup's per-row icon, since a data_editor cell can't host a
        # button of its own. Reclaims every detached row in an audience
        # back onto that audience's OWN shared/cascade color in one click
        # -- not a row's own (detached) color pushed OUT, which would
        # PROMOTE an accidental edit into the new shared default instead of
        # undoing it. `apply_pending_color_cascade` re-attaches (clears
        # `color_locked`) every group it touches, so this is the intended
        # "undo an accidental detach" action, and the shared color is the
        # only color that actually undoes one.
        #
        # One control per AUDIENCE, not one per group -- the old version
        # rendered one button per member of every audience with >1 group,
        # REGARDLESS of whether anything was actually detached, which is
        # nonsense: a rep who imported a real 12-row Hershey document (2
        # audiences x 6 geographies each, nothing broken out yet) saw 12
        # near-identical "Apply to all" buttons before this was scoped down
        # (live feedback, 2026-08-23). Only an audience with >=1 DETACHED
        # row gets a button, and the whole expander is hidden outright when
        # nothing anywhere is detached -- there's nothing to reclaim.
        audiences_seen = {}
        for group in new_groups:
            aud = tg.audience_label(group)
            if aud.strip():
                audiences_seen.setdefault(aud, []).append(group)
        cascade_targets = []   # (audience, shared_color, member_count, detached_count)
        for aud, members in audiences_seen.items():
            detached_count = sum(1 for g in members if g.get("color_locked"))
            if detached_count == 0:
                continue
            shared_color = tg.audience_cascade_color(members, aud) or tg.GROUP_COLORS[0]
            cascade_targets.append((aud, shared_color, len(members), detached_count))

        if cascade_targets:
            with st.expander("🎨 Apply a color to a whole audience", expanded=False):
                st.caption("A color picked in the table above changes only that one row and "
                           "detaches it from its audience's shared color. Reclaim every "
                           "detached row in an audience back onto its shared color in one "
                           "click below -- a row still following the shared color is never "
                           "touched.")
                for aud, shared_color, member_count, detached_count in cascade_targets:
                    cols = st.columns([5, 2])
                    cols[0].markdown(
                        f"{_color_swatch_label(shared_color)} &nbsp; **{aud}** "
                        f"&mdash; {detached_count} of {member_count} group(s) detached")
                    if cols[1].button("Apply to all", key=f"apply_color_{aud}"):
                        st.session_state["_pending_color_cascade"] = (aud, shared_color)
                        st.rerun()

        # FLOW_REWORK_PLAN.md Phase 3: grouping rows into one entity is a
        # rep's own explicit action, never inferred from Label text
        # happening to match -- see targeting_groups.py's module docstring.
        # This NEVER combines rows/lines -- each stays its own avails row
        # and its own plan line until a rep merges it by hand (commit 3's
        # existing merge_plan_rows); grouping only changes how a stated
        # budget divides (by entity count, not row count) and how a later
        # merge aggregates avails (max within the entity, not sum).
        entity_option_to_gid = {}
        for group in new_groups:
            aud = tg.audience_label(group)
            geo = tg.geo_label(group, label_for=_market_display_name)
            option_label = f"{aud} / {geo}" if (aud or geo) else group["id"]
            if option_label in entity_option_to_gid:
                # A real label collision (two rows can render identically) --
                # disambiguated for the multiselect's own internal key only;
                # a rep picks by the text either way.
                option_label = f"{option_label} ({group['id']})"
            entity_option_to_gid[option_label] = group["id"]

        with st.expander("🔗 Group rows into one entity", expanded=False):
            st.caption("Several avails rows for one real-world thing -- two radius tiers "
                       "for one dealership, two locations of the same brand -- can share "
                       "one entity, so a stated budget divides by ENTITY count, not row "
                       "count, and a later merge maxes their avails instead of summing "
                       "them. This never combines the rows themselves.")
            picked = st.multiselect("Rows to group", list(entity_option_to_gid),
                                    key="entity_group_picker")
            if st.button("Group selected rows", key="entity_group_apply",
                        disabled=len(picked) < 2):
                st.session_state["_pending_entity_group"] = [entity_option_to_gid[p] for p in picked]
                st.rerun()

            grouped_entities = {eid: members for eid, members in tg.entities_of(new_groups).items()
                               if len(members) > 1}
            if grouped_entities:
                st.caption("Existing groupings:")
                for eid, members in grouped_entities.items():
                    entity_display = tg.entity_label_of(members[0]) or "(unlabeled entity)"
                    cols = st.columns([5, 2])
                    cols[0].markdown(f"**{entity_display}** &mdash; {len(members)} rows")
                    if cols[1].button("Ungroup", key=f"ungroup_{eid}"):
                        st.session_state["_pending_entity_ungroup"] = eid
                        st.rerun()

        # One geo-definition expander per AUDIENCE, not per group --
        # Counties/Zips/Radius resolution, beside the grid's own quick
        # Markets cell rather than replacing it. A real 12-row Hershey
        # import (one audience, six geographies) used to render 12
        # near-identical "📍 Geography: ..." headers differing only in
        # which market followed the audience name (live feedback,
        # 2026-08-23); grouping by audience turns that into two expanders,
        # one per audience, each stacking every one of its geographies.
        # `_geo_panel_body` (no expander of its own) is what makes this
        # possible -- Streamlit refuses to nest an expander inside another.
        #
        # A group with no audience yet (geography can be defined before the
        # audience is picked) has nothing to group it under, so it keeps
        # its own standalone expander via `render_group_geo_expander`.
        groups_by_audience = {}
        ungrouped_geo = []
        for group in new_groups:
            aud = tg.audience_label(group)
            if aud.strip():
                groups_by_audience.setdefault(aud, []).append(group)
            else:
                ungrouped_geo.append(group)

        for aud, members in groups_by_audience.items():
            header = f"📍 Geography: {aud}"
            if len(members) > 1:
                header += f" -- {len(members)} geograph{'y' if len(members) == 1 else 'ies'}"
            with st.expander(header, expanded=False):
                for i, group in enumerate(members):
                    if i > 0:
                        st.divider()
                    _geo_panel_body(group)

        for group in ungrouped_geo:
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
    # The gross-up checkbox itself renders further down, immediately above
    # the editable grid (moved there from "near Generate" -- see DECISIONS.md
    # -- since it's the control that changes what a rep types into that
    # grid's CPM/Cost cells) -- but `markup` is needed here, before any
    # option's grid renders, since compute_plan_totals (called per option
    # below) reads it. Reading a keyed widget's own session_state ahead of
    # instantiation is fine; only writing it isn't, so this is safe even
    # though the checkbox hasn't rendered yet THIS run -- session_state
    # already holds whatever it was last set to (including a click that
    # triggered this very rerun).
    agency_gross_up = bool(st.session_state.get("agency_gross_up", False))
    markup = AGENCY_MARKUP if agency_gross_up else 1.0

    st.header("E. Proposal / media plan")
    ai_section_badge("media_plan")
    st.caption("Up to three plan options (good/better/best, or several budgets). "
               "Cost = Impressions/1000 x CPM, net -- see the agency gross-up "
               "checkbox just above the grid for a x1.15 order-wide multiplier.")

    # Flight dates and per-month flighting still live entirely in the setup
    # band's own session_state keys (FLOW_REWORK_PLAN.md Phase 4a) -- ONE
    # owner. What follows is a SECOND RENDER SITE for those same keys, not a
    # second control: a rep who gets to the bottom, sees the plan, and
    # realizes the flight is wrong used to have to scroll back to the top
    # (Phase 6) -- this lets them fix it here, and the band above picks up
    # the change on the very next render, same as if they'd edited it up
    # there. flight_start/flight_end/custom_flighting are all band widget
    # keys that already instantiated above in this SAME run, so an edit here
    # can't write them directly (the exact error client_name/total_tv raised
    # -- see BAND_WIDGET_KEYS); it's queued instead
    # (apply_pending_flight_mirror_edit, applied at the very top of the next
    # run, before the band's first widget) and reruns. `flight_months`
    # itself is NOT a widget key (the band's own per-month block already
    # writes it directly, unconditionally, every render), so a per-month
    # edit here writes it the same way -- the only extra step is bumping
    # `flight_months_gen` when it actually changes, which is what keeps the
    # band's own per-month widgets (a DIFFERENT key namespace) from showing
    # a stale value under the "a keyed widget's session_state beats value="
    # rule this file is careful about everywhere else.
    #
    # The top row (start/end/custom-flighting) needs its OWN resync
    # mechanism, separate from `flight_months_gen` -- found live, building
    # this: toggling "Custom flighting" alone (no date change) never bumps
    # `flight_months_gen` (existing, correct band behavior -- the per-month
    # ranges don't change just because the checkbox did), so the mirror's
    # own checkbox kept its stale session_state from an EARLIER render under
    # the same key, and read back its own OLD value as if the rep had just
    # (re-)unchecked it here -- silently writing the band's real
    # custom_flighting=True back to False. Same "a keyed widget's session_
    # state beats value=" trap, one level up: `bump_plan_options_generation`
    # /the per-option Breakout resync (Section E, above) solve the identical
    # problem by comparing against the CANONICAL value every render and
    # re-keying on any mismatch, not just on a mismatch this render's own
    # widget produced -- `_flight_mirror_signature` is that same comparison
    # for this mirror's three top-row widgets.
    mirror_signature = (flight_start, flight_end, custom_flighting)
    if st.session_state.get("_flight_mirror_signature") != mirror_signature:
        st.session_state["flight_mirror_gen"] = st.session_state.get("flight_mirror_gen", 0) + 1
        st.session_state["_flight_mirror_signature"] = mirror_signature
    mirror_top_gen = st.session_state["flight_mirror_gen"]
    mirror_row_gen = st.session_state.get("flight_months_gen", 0)
    with st.expander(
            f"📅 Flight: {n_months} active month(s), {flight_label}"
            + (" -- custom flighting" if custom_flighting else ""), expanded=False):
        st.caption("Same flight as the setup band above -- edit here or up there, both stay "
                   "in sync.")
        mcol1, mcol2 = st.columns(2)
        with mcol1:
            mirror_start = st.date_input("Flight start", value=flight_start,
                                          key=f"flight_start_e_{mirror_top_gen}")
        with mcol2:
            mirror_end = st.date_input("Flight end", value=flight_end,
                                        key=f"flight_end_e_{mirror_top_gen}")
        mirror_custom = st.checkbox(
            "Custom flighting -- skip or clip individual months",
            value=custom_flighting, key=f"custom_flighting_e_{mirror_top_gen}")

        mirror_pending = {}
        if mirror_start != flight_start or mirror_end != flight_end:
            mirror_pending["flight_start"] = mirror_start
            mirror_pending["flight_end"] = mirror_end
        if mirror_custom != custom_flighting:
            mirror_pending["custom_flighting"] = mirror_custom

        # Gated on mirror_custom/mirror_start/mirror_end (this render's OWN
        # widget returns), not the band's possibly-stale locals, the same
        # way the band's own block gates on its just-returned
        # `custom_flighting` -- toggling the checkbox above reveals or hides
        # these rows in the SAME render.
        if mirror_start and mirror_end and mirror_custom:
            mirror_base_ranges = flight_month_ranges(
                mirror_start, mirror_end, st.session_state.get("flight_months"))
            mirror_edited_ranges = []
            for i, entry in enumerate(mirror_base_ranges):
                bounds = default_month_range(entry["month"], mirror_start, mirror_end)
                ercol1, ercol2, ercol3 = st.columns([1.6, 1.2, 1.2])
                with ercol1:
                    row_active = st.checkbox(entry["month"], value=entry["active"],
                                             key=f"fm_active_e_{mirror_row_gen}_{i}")
                with ercol2:
                    row_start = st.date_input(
                        f"{entry['month']} start", value=entry["start"],
                        min_value=bounds["start"], max_value=bounds["end"],
                        key=f"fm_start_e_{mirror_row_gen}_{i}", disabled=not row_active,
                        label_visibility="collapsed")
                with ercol3:
                    row_end = st.date_input(
                        f"{entry['month']} end", value=entry["end"],
                        min_value=bounds["start"], max_value=bounds["end"],
                        key=f"fm_end_e_{mirror_row_gen}_{i}", disabled=not row_active,
                        label_visibility="collapsed")
                if row_end < row_start:
                    st.warning(f"{entry['month']}: end date can't be before start -- left unchanged.")
                    row_start, row_end = entry["start"], entry["end"]
                mirror_edited_ranges.append({"month": entry["month"], "start": row_start,
                                             "end": row_end, "active": row_active})
            if not any(r["active"] for r in mirror_edited_ranges):
                st.warning("At least one month has to stay active -- reverting to the full flight.")
                mirror_edited_ranges = flight_month_ranges(mirror_start, mirror_end)
            if mirror_edited_ranges != mirror_base_ranges:
                st.session_state["flight_months"] = mirror_edited_ranges
                bump_flight_months_generation()
                # A plain session_state mutation doesn't repaint the band's
                # OWN per-month rows, already drawn earlier in this SAME
                # run, before this write happened -- they'd show one stale
                # render otherwise, under the pre-bump key. Forcing the
                # rerun here (mirroring the queued-edit path below) is what
                # keeps this a real "one owner, two render sites" mirror
                # rather than one that's briefly a second source of truth.
                st.rerun()

        if mirror_pending:
            st.session_state["_pending_flight_mirror_edit"] = mirror_pending
            st.rerun()

    # Deck-wide display options, grouped together because both are the same
    # kind of thing: off-by-default, opt-in-per-proposal toggles that change
    # what shows on the media plan preview and the generated slide, never
    # what's actually sold. Neither is a product selection (Section C) even
    # though co-viewing only affects Premion Streaming TV and Live Sports
    # lines -- it's a way of presenting those lines, not a line itself.
    st.subheader("Options")
    ocol1, ocol2, ocol3 = st.columns(3)
    with ocol1:
        # Share of voice: one plan LINE's impressions against ITS OWN matching
        # avails line, never a deck-wide aggregate -- a line only has a match
        # when its _group_ids (the same join merge/split/the map legend already
        # use to tie a plan row back to a targeting group) resolve to a real
        # group, so an AI-drafted line, a quick-add line (picked by audience/geo
        # TEXT, never string-matched to an avails row -- see quick_add_rows) and
        # a legacy or hand-typed line all correctly show no percentage rather
        # than a guessed one. A merged line (2+ ids, from merge_plan_rows
        # combining several avails-backed lines into one) sums the avails of
        # every id it still carries, the same summing merge already does for
        # that line's Impressions and Cost -- a rep who merges "Homeowners,
        # Denver" and "Homeowners, Atlanta" into one line wants the combined
        # avails on the bottom of that line's percentage, not no percentage at
        # all.
        show_sov = st.checkbox(
            "Show % of avails (share of voice) on matched lines",
            value=False, key="show_sov",
            help="Adds \"(X% of avails)\" after a line's impressions, both here "
                 "and on the generated slide -- only for a line whose targeting "
                 "traces back to a real avails figure. A line with no such match "
                 "(drafted, quick-added, hand-typed) shows no percentage.")
    with ocol2:
        # Eligible lines are Premion Streaming TV and Live Sports only
        # (COVIEWING_ELIGIBLE_TACTIC_PREFIXES) -- every other line shows a
        # dash regardless of this toggle.
        show_coviewing = st.checkbox(
            "Show co-viewing (estimated person-level exposure)", value=False,
            key="show_coviewing",
            help="Adds a Monthly Coviewing column (the additional impressions beyond "
                 "household, at the configured multiplier) to Premion Streaming TV and "
                 "Live Sports lines, plus a small citation on the slide. Every other "
                 "line -- Streaming Retargeting, broadcast, AM, flat fees -- shows a dash.")
        if show_coviewing:
            # The slide's own citation is condensed to fit the real space it
            # lands in (see the coviewing_footnote comment further down) -- the
            # full citation is shown here instead, where there's no such limit.
            st.caption(f":grey[{COVIEWING_SETTINGS['footnote']}]")
    with ocol3:
        # FLOW_REWORK_PLAN.md Phase 3: off by default, and a no-op for every
        # proposal with no entity grouping in play (Lawn & Leisure included) --
        # see entity_prefixed_targeting's own docstring for what it actually
        # changes (the Targeting text only, never the underlying data).
        #
        # A computed default, not a forced state (LiveWell: seven ticked
        # rows, one audience, seven different locations -- exactly the
        # shape where the Targeting column alone reads identically on
        # every row). Fires AT MOST ONCE ever for a proposal
        # (show_entity_label_suggested is the lock, same one-shot shape
        # as every other queued suggestion in this file), the first
        # render on which the ticked groups actually qualify -- checked
        # fresh every render (cheap: audience_label/geo_label are already
        # computed per group, no I/O), because which groups are TICKED
        # can change via several different actions (an avails import's
        # own "Add all to plan", a draft's pre-selection, a single D2
        # tick), not just import. After it fires once, this code never
        # touches the key again, so a rep's own later choice -- leaving
        # it on, or turning it back off -- is never fought.
        if not st.session_state.get("show_entity_label_suggested"):
            ticked = [g for g in (st.session_state.get("targeting_groups") or [])
                     if g.get("include_in_plan")]
            if len(ticked) >= 2:
                audiences = {tg.audience_label(g) for g in ticked}
                geo_labels = {tg.geo_label(g, label_for=_market_display_name) for g in ticked}
                if len(audiences) == 1 and len(geo_labels) >= 2:
                    st.session_state["show_entity_label"] = True
                st.session_state["show_entity_label_suggested"] = True
        show_entity_label = st.checkbox(
            "Show Label in plan", value=False, key="show_entity_label",
            help="Prefixes a group-owned line's Targeting text with its entity's "
                 "Label (e.g. \"Toyota of Annapolis\") on both this preview and the "
                 "generated slide -- useful for a dealer group or franchise where a "
                 "client needs to see which store each line belongs to. A line with "
                 "no entity Label, or from a single-entity proposal, is unaffected.")

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
    # flight_shorthand rides along in the key (not just flight_label) so
    # editing one month's own range -- without changing the month SET --
    # still re-seeds clean rows' Flight cells. All three _shared_fields_key
    # construction sites (here, rehydration, draft application) must move
    # together, or a mismatch silently reseeds every option's rows to $0.
    shared_fields_key = (default_targeting + "||"
                         + " / ".join(f"{a}@{g}" for a, g, _gid in plan_lines)
                         + "||" + flight_label + "||" + flight_shorthand)

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
        # Premion Streaming TV is owned entirely by reconcile_group_plan_lines
        # now, via a group's own include_in_plan flag -- never seeded here.
        # A copy of seed_selections with that one flag forced off is what
        # every OTHER product (retargeting, AM, sports, fees, broadcast)
        # still gets built from, so _seeded_tactics below never contains
        # "Premion Streaming TV" and the ordinary product-toggle branches
        # can never add or remove a group's row.
        #
        # `default_geo` (the single flat campaign default), not `plan_lines`
        # (the per-group fan-out): retargeting/AM/sports/fees were never
        # conceptually "one line per targeting group" -- they inherited the
        # fan-out only because this closure used to seed Premion the same
        # way. One line per selected product, campaign-wide, is what they
        # produced before targeting groups existed and what they produce
        # again now that Premion has its own explicit, ticked ownership --
        # a rep with 12 avails-table audiences and Retargeting on gets ONE
        # Retargeting line, not 12, whether or not any of those 12 are
        # actually on the media plan.
        non_group_selections = dict(seed_selections, _premion_streaming_tv=False)
        rows = seed_media_plan_rows(non_group_selections, default_geo, default_targeting, flight_shorthand)
        # seed_media_plan_rows falls back to one blank placeholder row when
        # NOTHING was selected -- but Premion is never in `rows` here, so
        # that blank row would fire even when Premion IS selected and about
        # to get its own real row from reconcile_group_plan_lines. Drop it
        # in that case: the reconciler owns "nothing at all is selected"
        # too (its own ungrouped-fallback branch), so there's exactly one
        # blank/default row either way, never two.
        if (len(rows) == 1 and not rows[0].get("Tactic")
                and seed_selections.get("_premion_streaming_tv")):
            rows = []
        broadcast_row, _ = _broadcast_row_for_option(breakout)
        if broadcast_row is not None:
            # The imported schedule IS the broadcast buy, so it replaces the
            # rate-card line Total TV seeds rather than sitting beside it --
            # two broadcast lines would double-count the same spots.
            rows = [r for r in rows
                    if str(r.get("Tactic", "")).strip() != BROADCAST_PRODUCT_LABEL]
            rows.append(dict(broadcast_row))
        return rows

    def _seed_group_row(group, breakout):
        """One freshly-seeded Premion Streaming TV row for a single selected
        group, or the ungrouped single-row fallback when `group` is None --
        thin wrapper over seed_media_plan_rows so a group's line comes from
        the exact same product/CPM/targeting resolution every other seeded
        row does.
        """
        if group is not None:
            geo = tg.geo_label(group, label_for=_market_display_name).strip() or default_geo
            entry_list = [(tg.audience_label(group), geo, group["id"])]
        else:
            entry_list = [default_geo]
        only_premion = {"products": {}, "_premion_streaming_tv": True}
        return seed_media_plan_rows(only_premion, entry_list, default_targeting, flight_shorthand)[0]

    # .get() rather than [] on the two seed keys: they're written alongside
    # plan_options everywhere that sets it, but a missing key should re-seed
    # rather than raise.
    if not st.session_state.get("plan_options"):
        default_breakout = default_breakout_for_basis(st.session_state.get("avails_basis"))
        seeded = _seed_option_rows(default_breakout)
        st.session_state["plan_options"] = [new_plan_option(
            DEFAULT_OPTION_NAMES[0], seeded, breakout=default_breakout)]
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
            # This is "the group was DELETED" -- a different question from
            # "the group was unticked", which reconcile_group_plan_lines
            # (called further down, after this branch) owns exclusively for
            # Premion Streaming TV rows; this branch still applies to every
            # OTHER group-backed row (retargeting/AM/sports fan-out) and
            # remains a harmless no-op re-check for a Premion row the
            # reconciler already removed. A row seeded FOR a group must not
            # outlive that group just because nothing else about the plan
            # changed. Reset-to-default (the old, only, behavior) left
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
                        flight_shorthand, current=row))
            opt["version"] += 1
        st.session_state["_shared_fields_key"] = shared_fields_key

    # The single reconciliation pass for every Premion Streaming TV group
    # row -- add/remove by include_in_plan, re-price against a stored draft
    # intent when asked -- independent of whichever product-toggle branch
    # above did or didn't fire. See reconcile_group_plan_lines's own
    # docstring for why idempotence here is what makes order of operations
    # (import vs. draft vs. clarify, in any sequence) converge.
    group_plan_notes = reconcile_group_plan_lines(
        st.session_state["plan_options"], st.session_state.get("targeting_groups") or [],
        seed_group_row=_seed_group_row, fallback_geo=default_geo,
        intent=st.session_state.get("draft_plan_intent"),
        realloc=bool(st.session_state.pop("_pending_group_realloc", False)),
        confirmed_removals=st.session_state.pop("_confirmed_group_row_removals", ()),
        premion_selected=bool(seed_selections.get("_premion_streaming_tv")),
        flight_label=flight_label,
        default_targeting=default_targeting, n_months=n_months)
    if group_plan_notes:
        st.session_state["_group_plan_notes"] = group_plan_notes

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

    if broadcast_warning:
        st.warning(broadcast_warning)
    if schedule:
        # The two flights are set independently -- the plan's on the form,
        # the schedule's by Wide Orbit -- and nothing forces them to agree.
        # Often they shouldn't: a four-week broadcast burst inside a longer
        # streaming campaign is a normal buy. So this names both spans and
        # leaves the seller to decide, rather than correcting either.
        #
        # Compared against ACTIVE months only (FLOW_REWORK_PLAN.md Phase 2
        # ruling, decided explicitly, not left to the doc's own ambiguous
        # wording): a skipped month is one the plan isn't buying, so a
        # schedule running inside it is exactly the mismatch this warning
        # exists to catch, not noise to suppress. Comparing against every
        # calendar month the flight SPANS (including a skipped one) would
        # bury that real warning among noise about months nobody's buying --
        # and noise in a warning is how real warnings get ignored.
        schedule_months = {week.strftime("%b %Y") for week in schedule.grid_weeks}
        if schedule_months and not schedule_months <= set(active_months):
            st.warning(
                f"This schedule runs **{', '.join(sorted(schedule_months, key=_month_sort))}** "
                f"but the proposal's flight is **{flight_label}**. The broadcast line is priced "
                f"off the schedule's own dates, so the two don't have to match — just check "
                f"they're meant to differ.")
    if schedule and BROADCAST_EXCLUDED_FROM_AGENCY_MARKUP:
        st.caption(
            f"📺 The **{BROADCAST_TACTIC_MARKER}** line uses the Wide Orbit cost exactly as "
            f"quoted — broadcast is already gross, so it is never marked up by the ×1.15 "
            f"agency gross-up"
            + (", even though the gross-up checkbox above the plan is on." if agency_gross_up else ".")
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
                default_breakout = default_breakout_for_basis(st.session_state.get("avails_basis"))
                plan_options.append(new_plan_option(
                    name, _seed_option_rows(default_breakout), breakout=default_breakout))
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

    # show_sov (the "Options" section above, under Flight) gates the suffix
    # below: blank unless a line's _group_ids (the same join merge/split/the
    # map legend already use to tie a plan row back to a targeting group)
    # resolve to a real group with a real avails figure -- an AI-drafted
    # line, a quick-add line, or a hand-typed line all correctly show no
    # percentage rather than a guessed one. A merged line (2+ ids, from
    # merge_plan_rows combining several avails-backed lines into one) sums
    # the avails of every id it still carries, the same summing merge
    # already does for that line's Impressions and Cost.
    groups_by_id = {g["id"]: g for g in (st.session_state.get("targeting_groups") or [])}

    def _sov_suffix(impressions, avails):
        # " (X% of avails)" when the toggle is on and there's a real,
        # matched avails figure to divide into -- blank otherwise, never a
        # bogus 0%/inf% and never a guess at an unmatched line's avails.
        if not show_sov or not avails:
            return ""
        return f" ({impressions / avails * 100.0:.0f}% of avails)"

    # ---------------- Agency gross-up (FLOW_REWORK_PLAN.md Phase 4b) ------
    # A rep-only, order-wide calculator -- never touched by drafting, never
    # auto-ticked by an avails-PDF import (drafted or not: total_budget_basis
    # can tick it automatically, but the model still never computes anything
    # -- see DECISIONS.md). Every row on the grid below is NET; this
    # multiplies CPM and cost by x1.15 at display/deck time only (in
    # compute_plan_totals, via row_markup), for every line except broadcast
    # (BROADCAST_EXCLUDED_FROM_AGENCY_MARKUP -- a Wide Orbit cost is already
    # gross). Impressions never move. "A number gets grossed exactly once,
    # by whoever grossed it first" -- the model never grosses (it
    # transcribes whatever rate the notes state, gross or net, verbatim), so
    # this checkbox is the ONLY place a number is ever grossed, and it
    # applies to the whole order at once -- there is no per-line exception
    # besides broadcast. Placed HERE, immediately above the editable grid
    # rather than after it (and after the deck-preview table, and after
    # Included with Campaign) -- this is the control that changes what a rep
    # types into that grid's CPM/Cost cells, so it belongs beside them, not
    # buried several sections below where the numbers it affects have
    # already scrolled out of view. One checkbox for the whole order
    # regardless of how many plan options exist, so it renders once, here,
    # before the per-option tabs -- not inside the tabs loop, which would
    # try to instantiate the same widget key more than once.
    agency_gross_up = st.checkbox(
        "Apply agency gross-up (×1.15)", value=False, key="agency_gross_up",
        help="Multiplies every non-broadcast line's CPM and cost by x1.15 -- never "
             "impressions. Off by default. If the notes already state a grossed "
             "figure (e.g. \"gross up the $32 net CPM\" -> $36.80 entered on the "
             "line), leave this off -- the rate on the line IS the grossed one "
             "already, and ticking this would gross it again. Tick it when the "
             "whole order should go gross, which is the only case there is -- one "
             "line grossing and another not is not a real buy.")

    option_results = []
    tabs = st.tabs([o["name"] for o in plan_options])
    # FLOW_REWORK_PLAN.md Phase 4b: a markup change used to force this --
    # the old toggle mutated row["Cost"]/row["Impressions"] in place, and a
    # data_editor widget needed a fresh key to show the new numbers. The
    # gross-up no longer touches a stored row at all (it's applied fresh,
    # every render, inside compute_plan_totals), so there's nothing stale
    # left for a rerun to fix here.
    rerun_needed = False
    # Part of every per-option widget key -- see bump_plan_options_generation.
    gen = st.session_state.get("plan_options_gen", 0)

    # FLOW_REWORK_PLAN.md Phase 1 / DECISIONS.md: a MONTHLY-breakout row's
    # stored Cost is the rate the rep typed; compute_plan_totals DERIVES its
    # full-flight total by multiplying that rate by the row's own month
    # count (n_months) -- so anything that changes n_months after real money
    # is on the plan silently moves the quoted full-flight budget with no
    # ROW ever "rewritten" for the ordinary dirty-row protection to catch.
    # (A Full-Flight-breakout row's stored Cost IS the full-flight figure
    # directly, monthly derived by dividing -- immune to this by
    # construction.) Keyed on (flight_label, n_months) -- the same two
    # values main() already derives from active_months -- NOT on the band's
    # raw flight_start/flight_end: widening flight_end alone doesn't widen
    # n_months (a keyed multiselect keeps a still-valid partial selection,
    # per the documented rule below), so a raw-date comparison would have
    # missed exactly the expansion case this guard exists to catch, and
    # would have flagged date edits that never touched a priced dollar
    # figure at all. `_priced_flight_baseline` is the (flight_label,
    # n_months) as of the last time this was reviewed and dismissed;
    # `flight_change_deltas` collects (option name, old total, new total)
    # for every Monthly-breakout option with real priced content whose
    # total actually moved, populated inside the loop right after each
    # option's own totals are computed.
    flight_change_deltas = []
    current_flight_key = (flight_label, n_months)
    priced_flight_baseline = st.session_state.get("_priced_flight_baseline")

    for idx, (tab, option) in enumerate(zip(tabs, plan_options)):
        with tab:
            ncol1, ncol2 = st.columns([2, 2])
            with ncol1:
                option["name"] = st.text_input(
                    "Option name", value=option["name"], key=f"option_name_{gen}_{idx}",
                    help="Shown in the deck as part of the plan title, e.g. \"CTV Strategy — Good\".") or option["name"]
            with ncol2:
                # Real bug, fixed 2026-08-28 (FLOW_REWORK_PLAN.md Phase 2):
                # the band's Plan basis only ever set a BRAND-NEW option's
                # initial breakout -- once this radio had rendered once
                # (which happens on the very first run, before a rep can
                # act on the band at all), its own widget key permanently
                # won over `index=` on every later run, the same "a keyed
                # widget's session_state beats its value=/index= argument"
                # trap documented elsewhere in this file for the option-name
                # bug. So flipping the band afterward -- the ordinary way a
                # rep would ever touch this -- silently did nothing.
                #
                # An option that hasn't been locked re-syncs to the band's
                # CURRENT default every run it's found diverged from it,
                # bumping a small per-option `_breakout_gen` counter so the
                # radio gets a fresh key and `index=` is honored again --
                # same device as `bump_plan_options_generation`, scoped to
                # just this one widget so it never disturbs this option's
                # other widgets (name, rows, ...). A rep who picks a
                # DIFFERENT Breakout for this option by hand locks it
                # (`option_breakout_locked`), so the band can never silently
                # override a deliberate choice again -- detected by
                # comparing the widget's return value against what
                # `option["breakout"]` held immediately before this render.
                # That comparison can't false-positive on a same-run resync:
                # a resync sets the value AND bumps the key together, so the
                # freshly-keyed widget's own return is guaranteed to agree
                # with what was just set (nothing was there before).
                band_default_breakout = default_breakout_for_basis(
                    st.session_state.get("avails_basis"))
                if (not option.get("option_breakout_locked")
                        and option["breakout"] != band_default_breakout):
                    option["breakout"] = band_default_breakout
                    option["_breakout_gen"] = option.get("_breakout_gen", 0) + 1
                pre_widget_breakout = option["breakout"]
                option["breakout"] = st.radio(
                    "Breakout", BREAKOUT_MODES, horizontal=True,
                    key=f"option_breakout_{gen}_{idx}_{option.get('_breakout_gen', 0)}",
                    index=BREAKOUT_MODES.index(option["breakout"]))
                if option["breakout"] != pre_widget_breakout:
                    option["option_breakout_locked"] = True

            if rescale_rows_for_breakout_change(
                    option, n_months, schedule.active_month_count() if schedule else None):
                option["version"] += 1
                rerun_needed = True

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
                    pick_products, pick_audiences, pick_geos, flight_shorthand,
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

            media_plan_df = pd.DataFrame(option["rows"])
            # FLOW_REWORK_PLAN.md Phase 3: "Show Label in plan" adds a
            # disabled, informational Label column to the EDITABLE grid too
            # -- read-only, since the Targeting SelectboxColumn is a fixed
            # option list and folding a Label prefix into it would pollute
            # those options and get folded back into row["Targeting"] on
            # the very next edit (see entity_prefixed_targeting's own
            # docstring for why the preview/deck get the prefix a
            # different way). Popped from edited_records below before
            # reconcile_plan_rows ever sees it.
            if show_entity_label and not media_plan_df.empty:
                # `pd.DataFrame([])` (zero plan rows -- every group unticked,
                # or a brand-new option) has zero COLUMNS, not just zero
                # rows, so `.insert(1, ...)` is genuinely out of bounds
                # ("loc must be an integer between -0 and 0") -- a real,
                # pre-existing bug this phase's own code already had,
                # unreachable before "Show Label in plan" gained a computed
                # default that can now land on True for exactly this shape
                # (found live: unticking every group on an already-labeled
                # option). Nothing to label on an empty grid anyway --
                # skipping is the correct behavior here, not just the safe
                # one. The Label entry in column_config below is harmless
                # when the column itself doesn't exist.
                media_plan_df.insert(1, "Label", [
                    tg.entity_label_of(groups_by_id[group_ids_of(row)[0]])
                    if group_ids_of(row) and group_ids_of(row)[0] in groups_by_id else ""
                    for row in option["rows"]
                ])
            edited_df = st.data_editor(
                media_plan_df, num_rows="dynamic",
                key=f"media_plan_editor_{idx}_{option['version']}", use_container_width=True,
                column_config={
                    "Impressions": st.column_config.NumberColumn(f"Impressions ({basis})"),
                    "Type": st.column_config.SelectboxColumn(options=[ROW_TYPE_RATE, ROW_TYPE_FLAT_FEE]),
                    "Cost": st.column_config.NumberColumn(f"Cost ({basis}, $)", format="$%.0f"),
                    "Geo": st.column_config.SelectboxColumn("Geo", options=grid_geos),
                    "Targeting": st.column_config.SelectboxColumn(
                        "Targeting", options=grid_audiences),
                    **({"Label": st.column_config.TextColumn(
                        "Label", disabled=True,
                        help="The entity this line belongs to (set in the D2 avails table "
                             "above) -- read-only here.")}
                       if show_entity_label else {}),
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

            edited_records = edited_df.to_dict("records")
            if show_entity_label:
                for _r in edited_records:
                    _r.pop("Label", None)
            # Must run BEFORE reconcile_plan_rows overwrites option["rows"]
            # (it's the "prev" side of the comparison), and must land in
            # session_state THIS run -- reconcile_group_plan_lines already
            # ran earlier this same render, but a forced rerun (row count
            # changed, right below) re-enters it on the very next internal
            # pass, and that pass is what would otherwise resurrect the row
            # this rep just deleted.
            updated_groups = unlink_deleted_group_rows(
                st.session_state.get("targeting_groups") or [], option["rows"], edited_records)
            if updated_groups is not st.session_state.get("targeting_groups"):
                st.session_state["targeting_groups"] = updated_groups

            if reconcile_plan_rows(option, edited_records):
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
                    merge_plan_rows(option, [int(p.split(":")[0]) for p in merge_pick])
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
                option["rows"], breakout_mode, n_months, flight_shorthand,
                broadcast_months=(schedule.active_month_count() if schedule else None),
                groups_by_id=groups_by_id,
                coviewing_multiplier=(COVIEWING_SETTINGS.get("multiplier") if show_coviewing else None),
                markup=markup, show_entity_label=show_entity_label)
            option_results.append(totals)

            # Only MONTHLY-breakout options are exposed to this failure mode.
            # compute_plan_totals treats a Full-Flight row's stored Cost as
            # the full-flight figure directly (monthly is DERIVED by
            # dividing) -- immune to a flight-length change by construction.
            # A Monthly row's stored Cost is the rate (full-flight is
            # DERIVED by multiplying by row_months), which is exactly what
            # silently moves when n_months changes. See DECISIONS.md.
            if (breakout_mode == BREAKOUT_MONTHLY and priced_flight_baseline
                    and priced_flight_baseline != current_flight_key
                    and any(dirty and _num(row.get("Cost"))
                            for row, dirty in zip(option["rows"], option["dirty"]))):
                old_flight_label, old_n_months = priced_flight_baseline
                old_totals = compute_plan_totals(
                    option["rows"], breakout_mode, old_n_months, old_flight_label,
                    broadcast_months=(schedule.active_month_count() if schedule else None),
                    groups_by_id=groups_by_id,
                    coviewing_multiplier=(COVIEWING_SETTINGS.get("multiplier") if show_coviewing else None),
                    markup=markup)
                if abs(old_totals["full_flight_cost"] - totals["full_flight_cost"]) > 0.01:
                    flight_change_deltas.append(
                        (option["name"], old_totals["full_flight_cost"], totals["full_flight_cost"]))

            preview_columns = [
                "tactic", "flight", "geo", "targeting", "cpm", "monthly impressions", "monthly cost",
                "full flight impressions", "full flight cost"]
            if show_coviewing:
                preview_columns.insert(preview_columns.index("monthly cost"), "monthly coviewing")
            preview_display = pd.DataFrame([
                {
                    "tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
                    # The number this whole table exists to surface: r["cpm"]
                    # is compute_plan_totals' own row_markup-applied figure,
                    # already grossed when the checkbox above is on and net
                    # when it's off -- the same toggle-driven value the deck
                    # itself shows (_option_payload's own "cpm", below), just
                    # finally visible here too instead of only after
                    # Generate. One rate for the whole line regardless of
                    # basis, so it sits once, not duplicated per column the
                    # way impressions/cost are.
                    "cpm": "--" if r["is_flat_fee"] else f"${r['cpm']:,.2f}",
                    "monthly impressions": ("--" if r["is_flat_fee"] else
                                             f"{int(r['monthly_impressions']):,}"
                                             + _sov_suffix(r['monthly_impressions'], r['matched_avails_monthly'])),
                    **({"monthly coviewing": ("--" if r["coviewing_additional_monthly"] is None
                                               else f"+{r['coviewing_additional_monthly']:,}")}
                       if show_coviewing else {}),
                    "monthly cost": f"${r['monthly_cost']:,.0f}",
                    "full flight impressions": ("--" if r["is_flat_fee"] else
                                                 f"{int(r['full_flight_impressions']):,}"
                                                 + _sov_suffix(r['full_flight_impressions'], r['matched_avails_full_flight'])),
                    "full flight cost": f"${r['full_flight_cost']:,.0f}",
                }
                for r in totals["preview_rows"]
            ]) if totals["preview_rows"] else pd.DataFrame(columns=preview_columns)
            # Not a second copy of the grid above -- this is compute_plan_totals'
            # own preview_rows, the SAME source _option_payload (below) draws the
            # deck's media plan table from: it's grossed (if the checkbox is on,
            # where the editable grid above is always net), shows both monthly
            # AND full-flight side by side (the grid only labels/edits whichever
            # one the option's breakout is set to), and carries the SOV/co-viewing
            # suffixes the deck shows. Labelled so a rep isn't left guessing which
            # of two similar-looking tables is the real one.
            st.caption("**How this renders in the deck** -- both bases, with the "
                       "agency gross-up applied if it's on. The grid above is what "
                       "you edit; it's always net, one basis at a time.")
            st.dataframe(preview_display, use_container_width=True)

            gross_suffix = " gross" if agency_gross_up else ""
            st.caption(f"Monthly totals: {int(totals['monthly_impressions']):,} impressions / "
                       f"${totals['monthly_cost']:,.0f}{gross_suffix}")
            st.caption(f"**Full Flight Total ({n_months} month{'s' if n_months != 1 else ''}): "
                       f"{int(totals['full_flight_impressions']):,} impressions / "
                       f"${totals['full_flight_cost']:,.0f}{gross_suffix}**")

            # FLOW_REWORK_PLAN.md Phase 3, decision 12: what the draft
            # assumed, read-only -- see describe_group_allocation's own
            # docstring for why this is a caption, not a second editable
            # column.
            allocation_caption = describe_group_allocation(
                _intent_for_option(st.session_state.get("draft_plan_intent"), option),
                groups_by_id, option["rows"], option["dirty"])
            if allocation_caption:
                st.caption(f"\U0001F4CB {allocation_caption}")

    # Recomputed values live in session_state now but the grids on screen
    # still show what was typed, so re-render once. Guarded on an actual
    # change: on the next run every row already matches its own derivation,
    # nothing recomputes, and the loop ends.
    if rerun_needed:
        for opt in plan_options:
            opt["version"] += 1
        st.rerun()

    # Resolve the flight-change baseline now that every option's totals (and
    # flight_change_deltas, if any) are known. `flight_change_blocking`
    # (read by the Generate button further down) is a plain local, not
    # session_state -- it only ever needs to be right for THIS run.
    flight_change_blocking = False
    if priced_flight_baseline is None:
        st.session_state["_priced_flight_baseline"] = current_flight_key
    elif priced_flight_baseline != current_flight_key:
        if flight_change_deltas:
            flight_change_blocking = True
            old_flight_label, old_n_months = priced_flight_baseline
            delta_text = "; ".join(
                f"**{name}**: ${old:,.0f} → ${new:,.0f}"
                for name, old, new in flight_change_deltas)
            st.warning(
                f"⚠️ The active flight changed from {old_flight_label} ({old_n_months} "
                f"month{'s' if old_n_months != 1 else ''}) to {flight_label} ({n_months} "
                f"month{'s' if n_months != 1 else ''}). Because the Full Flight Total is "
                f"computed from the flight length, not stored per row, this moves the "
                f"quoted total on priced line(s) below: {delta_text}. Review the media "
                f"plan, then dismiss this to confirm before generating.")
            if st.button("I've reviewed the new totals — dismiss this warning"):
                st.session_state["_priced_flight_baseline"] = current_flight_key
                st.rerun()
        else:
            # The flight changed, but nothing priced/Full-Flight was
            # affected -- nothing to review, so this was never a pending
            # change to begin with.
            st.session_state["_priced_flight_baseline"] = current_flight_key

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

    if agency_gross_up and option_results:
        _grossed_total = sum(t["full_flight_cost"] for t in option_results)
        st.caption(f"Grossed full-flight total: ${_grossed_total:,.0f} across "
                   f"{len(option_results)} option{'s' if len(option_results) != 1 else ''} "
                   f"(broadcast, if any, stays at its own Wide Orbit cost).")

    # ---------------- Case studies ----------------
    selected_case_studies = render_case_study_picker(vertical_key, vertical_choice)

    # ---------------- Vault slides ----------------
    selected_vault_slides = render_vault_slide_picker(vertical_key, vertical_choice)

    # ---------------- Generate ----------------
    st.header("Generate")
    # FLOW_REWORK_PLAN.md Phase 5: replaces the withdrawn "satisfied
    # declarations" gate idea -- non-blocking, at Generate, where it's free,
    # rather than a hard block on the band. avails_mode/total_tv ticked
    # doesn't mean a document is coming: hand-typing every D2 row with no
    # PDF, or pricing broadcast off the rate card with no Wide Orbit
    # schedule, are both real, first-class paths (see DECISIONS.md), so a
    # gate that required an upload would lock a rep out of their own normal
    # workflow. These two warnings exist only to catch "forgot to pull it,"
    # never to demand something that was never promised.
    if avails_mode:
        _has_real_avails_group = any(
            g.get("terms") and not g.get("_placeholder")
            for g in (st.session_state.get("targeting_groups") or []))
        if not _has_real_avails_group:
            st.warning("⚠️ \"Working from an avails document\" is checked, but the avails "
                       "table is empty. Pull the avails, or uncheck the toggle.")
    if total_tv and not st.session_state.get("broadcast_schedule"):
        st.warning("⚠️ Total TV is on with no Wide Orbit schedule uploaded — the broadcast "
                   "line will use the rate-card default rather than real spot data.")
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

    if flight_change_blocking:
        st.caption("Generate is disabled until the flight-change warning above is dismissed.")
    if st.button("Generate proposal", type="primary", disabled=flight_change_blocking):
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
            "agency_gross_up": agency_gross_up,
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

        gross_note = " (Gross)" if agency_gross_up else ""
        multiple_options = len(plan_options) > 1

        def _option_payload(option, totals):
            rows = [
                {"tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
                 "impressions": ("--" if r["is_flat_fee"] else
                                 f"{int(r['monthly_impressions']):,}"
                                 + _sov_suffix(r['monthly_impressions'], r['matched_avails_monthly'])),
                 "coviewing": ("--" if r["coviewing_additional_monthly"] is None
                               else f"+{r['coviewing_additional_monthly']:,}"),
                 "cost": f"${r['monthly_cost']:,.0f}{gross_note}",
                 # r["cpm"] is already grossed (compute_plan_totals applies
                 # row_markup) when the agency toggle is on -- shows the
                 # same effective, all-in rate the Cost column's own
                 # "(Gross)" note is describing, instead of the net
                 # rate-card number sitting next to a cost that disagrees
                 # with it. A flat fee has no rate, so "--" rather than a misleading $0.
                 "cpm": "--" if r["is_flat_fee"] else f"${_num(r.get('cpm')):,.2f}"}
                for r in totals["preview_rows"]
            ] or [{"tactic": "", "flight": flight_shorthand, "geo": default_geo, "targeting": "",
                   "impressions": "0", "coviewing": "--", "cost": "$0"}]

            # Real bug, fixed 2026-08-28 (FLOW_REWORK_PLAN.md Phase 2): this
            # never read the option's own breakout at all -- a genuinely
            # Full-Flight-breakout option's deck still said "Monthly Totals"
            # over the derived monthly-equivalent number, with the actual
            # full-flight figure the rep typed relegated to a second,
            # separate footer row underneath. The D2/media-plan grid's own
            # column headers already read this same breakout (see the
            # `basis` local above); the deck's totals row and footer now do
            # too. Byte-identical to before this fix whenever breakout is
            # Monthly (the only case this ever ran before).
            is_full_flight_breakout = option["breakout"].startswith("Full Flight")
            full_flight_total = None
            if n_months > 1 and totals["preview_rows"] and not is_full_flight_breakout:
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

            # Only when the toggle is on AND this option actually has
            # something the multiplier applies to -- no dead footnote on a
            # plan with no CTV/Sports lines in it.
            #
            # Deliberately NOT the settings record's full `footnote` text
            # verbatim -- measured against the real master deck, the plan
            # slide's existing Terms & Conditions box (where this lands, see
            # assembly.add_coviewing_footnote) already sits within ~0.2in of
            # the slide's bottom edge before this adds anything, so the
            # on-slide clause is condensed to fit that real, measured space:
            # multiplier + source + the computed effective CPM. The fuller
            # citation (the settings record's own `footnote` field) is shown
            # in the app instead, next to the checkbox above, where there's
            # no such constraint.
            #
            # No "Co-viewing: " prefix here -- assembly.add_coviewing_footnote
            # already supplies that as the bold run label, matching the box's
            # own "Label: body" paragraphs (Terms & Conditions:, Payment:,
            # Cancellation:). Prefixing it here too doubled the label on the
            # real rendered slide, caught by looking at the render, not by
            # any assertion.
            coviewing_footnote = None
            if show_coviewing and any(r["coviewing_eligible"] for r in totals["preview_rows"]):
                effective_cpm = effective_cpm_with_coviewing(
                    totals["preview_rows"], COVIEWING_SETTINGS["multiplier"])
                coviewing_footnote = (
                    f"{COVIEWING_SETTINGS['multiplier']:g}x factor applied "
                    f"({COVIEWING_SETTINGS['source']}) -- effective CPM at estimated "
                    f"exposure ${effective_cpm:,.2f}.")

            return {
                "show_cpm": show_cpm_column,
                "show_coviewing": show_coviewing,
                "total_cpm": f"${blended:,.2f}" if blended else "--",
                "plan_title": option_plan_title(proposal_title, option["name"], multiple_options),
                "rows": rows,
                "totals_label": "Full Flight Totals" if is_full_flight_breakout else "Monthly Totals",
                "total_impressions": f"{int(totals['full_flight_impressions'] if is_full_flight_breakout else totals['monthly_impressions']):,}",
                "total_cost": f"${(totals['full_flight_cost'] if is_full_flight_breakout else totals['monthly_cost']):,.0f}{gross_note}",
                "full_flight_total": full_flight_total,
                "coviewing_footnote": coviewing_footnote,
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
                "TIMING_BULLETS": lines_to_bullets(timing_text) or [flight_shorthand],
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

        vault_slide_sources, vault_slide_errors = [], []
        for vault_slide in selected_vault_slides:
            try:
                vault_slide_sources.append(vault_slide_source(vault_slide, vault_slide["placement"]))
            except Exception as exc:
                vault_slide_errors.append(f"{vault_slide['title']}: {db.describe_error(exc)}")
        for message in vault_slide_errors:
            st.warning(f"Vault slide left out -- couldn't fetch it. {message}")

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
                # After case studies (so "before_plan" vault slides land
                # after them, not interleaved -- see
                # assembly.vault_slide_insert_index) and, like them, before
                # personalize.
                vault_slide_count = assembly.append_vault_slides(prs, vault_slide_sources)
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
                print(f"[assembly] failed: {type(exc).__name__}: {exc}")
                st.error("Something went wrong assembling the deck. Try Generate again -- if "
                         "it keeps happening, use Report an issue and I'll take a look.")
                raise

        for message in schedule_warnings + layout_warnings:
            st.warning(message)

        # An environment note rather than a problem with THIS proposal: it is
        # true of every deck this machine builds, and it's aimed at whoever
        # maintains this app, not at the rep who just generated a deck -- it
        # matters because the deployed instance has neither the brand fonts
        # nor Calibri, estimates text width instead, and so sizes tables (and
        # decides whether to compress the "Included with Campaign" band)
        # slightly differently from a machine that has the fonts. Surfacing
        # it as a page caption on every proposal put a paragraph of font
        # diagnostics in front of every rep on the deployed instance, where
        # it's true on literally every generate -- so it's logged in full
        # always (a log line nobody opens is the phantom-slide-37 problem in
        # a different costume) and shown on the page only in test mode, for
        # the one person it's actually written for. The behaviour it
        # describes is unchanged; only who sees the paragraph is.
        # Two different facts, so two notes rather than one. measurement_note
        # says a substitution happened; width_calibration_note says the
        # correction applied for it was not measured against that face, which
        # is what decides whether the client-name title can be trusted to sit
        # inside its box.
        for note in (text_metrics.measurement_note(),
                     assembly.width_calibration_note()):
            if note:
                _LOG.info(note)
                if test_mode_active():
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
        if vault_slide_count:
            extras.append(f"{vault_slide_count} vault slide"
                          f"{'s' if vault_slide_count != 1 else ''}")
        st.success(f"Assembled {kept_count + extra_option_slides + case_study_slides + vault_slide_count} "
                   f"of {original_count} slides"
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
                # "agency_gross_up" is the authoritative field from this phase
                # on -- its PRESENCE is what rehydrate_proposal_into_form uses
                # to tell a Phase-4b-or-later proposal (rows already stored
                # net) from an older one (rows stored gross whenever the old
                # per-row toggle was on, needing a one-time migration). Kept
                # alongside "markup" for a human reading the raw JSON, and
                # because DECISIONS.md's own incident record cites it.
                "agency_gross_up": agency_gross_up,
                "markup": markup,
                "flight": {"start": str(flight_start), "end": str(flight_end),
                           "label": flight_label, "active_months": [str(m) for m in active_months],
                           # Additive (FLOW_REWORK_PLAN.md Phase 2): a proposal
                           # logged before this exists has neither key, which is
                           # exactly the case resolve_flight_months/rebuild's
                           # stored_flight_display fallback handle.
                           "shorthand": flight_shorthand,
                           "month_ranges": flight_months_snapshot(edited_ranges)},
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
                # FLOW_REWORK_PLAN.md Phase 1 -- see setup_snapshot's own
                # docstring. Additive: a proposal reloaded by code that
                # predates this key simply doesn't look for it.
                "setup": setup_snapshot(
                    market_choice, flight_start, flight_end, avails_basis, avails_mode, total_tv),
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
                "vault_slides": [{"id": v["id"], "title": v["title"], "placement": v["placement"]}
                                 for v in selected_vault_slides],
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
