"""
db.py -- the single point of Supabase access for the whole app.

Nothing else imports the supabase client: every table read, storage
download and insert goes through a function in here. That's what makes the
fallback policy enforceable in one place rather than scattered across
call sites.

**Fallback policy.** Supabase is where the data lives, but the app has to
keep working when it isn't reachable (offline, broken URL, project paused,
local dev with no secrets). So every loader returns a
``(value, warning)`` pair:

    value    -- the real data when Supabase answered, otherwise the local
                fallback (the checked-in deck file / the hardcoded dicts).
    warning  -- ``None`` on success, otherwise a plain-language sentence the
                caller shows with ``st.warning`` so a fallback is never
                silent.

Never raise out of a loader for an unreachable backend; a *caller* deciding
it can't proceed is a different matter. Writes (``log_proposal`` and the
admin paths) return ``(ok, error)`` instead -- there's no local fallback for
a write, but a failed write must not take a generated deck down with it.
"""

import io
import json
import os
import tempfile
import time
from datetime import date, datetime, timezone
from pathlib import Path

import streamlit as st

try:
    from supabase import create_client
except ImportError:  # package not installed -- every loader falls back
    create_client = None

try:
    from supabase.client import ClientOptions
except ImportError:  # pragma: no cover -- older/newer client layouts
    ClientOptions = None

DECKS_BUCKET = "decks"
CASE_STUDIES_BUCKET = "case_studies"
# Individual slides colleagues want to reuse -- the .pptx a slide was picked
# out of lives here, same bucket regardless of how many slides were picked
# from it. See "# Slide vault" below.
SLIDE_VAULT_BUCKET = "slide_vault"
# Hand-edited final decks attached to a history row. The exception to
# storing-the-recipe: everything else in this app is regenerable from
# form_json, these are not, which is exactly why they're worth keeping --
# and why the History page surfaces this bucket's total usage.
PROPOSAL_FILES_BUCKET = "proposal_files"
# Market viewer profile slides, one JPEG per DMA. Its own bucket rather than
# a corner of case_studies: this set is replaced wholesale when Premion
# reissues the deck, where the vault accretes one case study at a time, and
# keeping them apart makes the quota reading mean something.
MARKET_PROFILES_BUCKET = "market_profiles"
# The raw YTD audience-usage workbook Matt uploads periodically -- one
# object per upload (see upload_audience_usage_workbook), the deck_versions
# pattern applied to a much smaller file.
AUDIENCE_USAGE_BUCKET = "audience_usage_workbooks"
# The attribution report master template (REPORT_MASTER_v0_2.pptx) -- a
# second, separate deck from the proposal master, own table
# (report_deck_versions) and own bucket, so a proposal deck and a report
# deck never collide in one "active" flag. See ATTRIBUTION_REPORT_PLAN.md
# Phase 3 / DECISIONS.md's Attribution Report Builder section.
REPORT_DECKS_BUCKET = "report_decks"

# What a single storage object is allowed to be. A bucket can carry its own
# file_size_limit, but the *project* has a global ceiling on top of it that
# no API exposes -- so this is the assumed limit whenever the bucket doesn't
# state one. 50MiB is the default for a project that hasn't raised it
# (Supabase dashboard -> Storage -> Settings -> Upload file size limit).
DEFAULT_STORAGE_LIMIT = 50 * 1024 * 1024

# What ensure_bucket asks for when creating the decks bucket. The project
# ceiling clamps this, and creation is retried without it if the project
# rejects the request outright.
PREFERRED_DECK_LIMIT = 500 * 1024 * 1024

# PostgREST calls are small and must fail fast, so a broken URL surfaces as a
# warning in a second or two rather than hanging the form. Storage moves the
# ~110MB master deck, so it gets a much longer leash.
POSTGREST_TIMEOUT = 15
STORAGE_TIMEOUT = 600

_DECK_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_decks"
_REPORT_DECK_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_report_decks"
_CASE_STUDY_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_case_studies"
_SLIDE_VAULT_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_slide_vault"
_PROPOSAL_FILE_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_proposal_files"
# Not scratch_dir(): these paths are handed out by @st.cache_resource, and
# scratch_dir sweeps files older than a few hours -- which would delete one
# out from under a live cache entry. Same reasoning as the two caches above.
_MARKET_PROFILE_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_market_profiles"

# Keyed on (url, key) rather than @st.cache_resource so that editing
# secrets.toml (e.g. temporarily breaking the URL to exercise the fallback)
# takes effect on the next rerun instead of being pinned by a cache.
_clients = {}


def _secret(name):
    """Secrets come from .streamlit/secrets.toml under the app, or the
    environment when a helper script runs outside Streamlit."""
    try:
        value = st.secrets.get(name)
    except Exception:
        value = None
    return value or os.environ.get(name)


def is_configured():
    return bool(create_client and _secret("SUPABASE_URL") and _secret("SUPABASE_SERVICE_KEY"))


def get_client():
    """The shared Supabase client, or None if the package or secrets are
    missing. Constructing a client does no I/O, so this never blocks -- an
    unreachable project only shows up when a call is made."""
    url, key = _secret("SUPABASE_URL"), _secret("SUPABASE_SERVICE_KEY")
    if not (create_client and url and key):
        return None
    cache_key = (url, key)
    if cache_key not in _clients:
        if ClientOptions is not None:
            options = ClientOptions(
                postgrest_client_timeout=POSTGREST_TIMEOUT,
                storage_client_timeout=STORAGE_TIMEOUT,
            )
            _clients[cache_key] = create_client(url, key, options=options)
        else:  # pragma: no cover
            _clients[cache_key] = create_client(url, key)
    return _clients[cache_key]


def describe_error(exc):
    """One-line form of an exception for a user-facing warning -- Supabase
    errors stringify to multi-line JSON blobs that read badly in st.warning."""
    text = " ".join(str(exc).split())
    return text[:200] + ("..." if len(text) > 200 else "")


# ---------------------------------------------------------------------------
# Master deck versions (Stage 1)
# ---------------------------------------------------------------------------
def active_deck_version():
    """(row, warning) for the single active deck_versions row."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = client.table("deck_versions").select("*").eq("active", True).limit(1).execute()
    except Exception as exc:
        return None, f"Couldn't reach Supabase ({describe_error(exc)})"
    rows = result.data or []
    if not rows:
        return None, "No active master deck version is registered in Supabase"
    return rows[0], None


@st.cache_resource(show_spinner="Fetching the master deck...")
def _deck_file_for_version(version_id, storage_path):
    """Local path to one deck version's .pptx, downloading it if needed.

    Cached on the version id, so a session downloads the ~110MB deck once
    rather than once per generate. The on-disk copy is keyed on the version
    id too, so a *new* process reuses it instead of re-downloading, and
    activating a new version can never be served a stale file. Downloads land
    on a .part file first so an interrupted one can't be mistaken for a
    complete deck later.
    """
    target = _DECK_CACHE_DIR / f"v{version_id}_{Path(storage_path).name}"
    if target.exists() and target.stat().st_size > 0:
        return str(target)

    blob = get_client().storage.from_(DECKS_BUCKET).download(storage_path)
    _DECK_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


def master_deck(local_fallback_path):
    """(path, deck_version_id, warning) -- where to read the master deck from.

    Falls back to the checked-in local file whenever Supabase can't supply
    one, with deck_version_id None so a proposal logged against it doesn't
    claim a version it didn't use.

    **path is None when there is no deck at all.** The .pptx is gitignored,
    so a deployed instance has no local copy to fall back to -- if Supabase
    is also unreachable there is genuinely nothing to build from, and the
    caller has to say so rather than handing python-pptx a path that doesn't
    exist and surfacing a bare PackageNotFoundError.
    """
    row, warning = active_deck_version()
    if row is None:
        return _fallback_deck(local_fallback_path, warning)
    try:
        return _deck_file_for_version(row["id"], row["storage_path"]), row["id"], None
    except Exception as exc:
        return _fallback_deck(
            local_fallback_path,
            f"Couldn't download master deck version {row['id']} from Supabase "
            f"({describe_error(exc)})")


def _fallback_deck(local_fallback_path, warning):
    if local_fallback_path and Path(local_fallback_path).exists():
        return local_fallback_path, None, f"{warning}. Using the local master deck file instead."
    return None, None, (
        f"{warning}, and there's no local master deck to fall back on. "
        f"A proposal can't be generated until Supabase is reachable again — "
        f"nothing else you've filled in is lost, so try again in a moment."
    )


def active_report_deck_version():
    """(row, warning) for the single active report_deck_versions row --
    active_deck_version()'s exact equivalent for the attribution report
    master."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = client.table("report_deck_versions").select("*").eq("active", True).limit(1).execute()
    except Exception as exc:
        return None, f"Couldn't reach Supabase ({describe_error(exc)})"
    rows = result.data or []
    if not rows:
        return None, "No active report deck version is registered in Supabase"
    return rows[0], None


@st.cache_resource(show_spinner="Fetching the report master deck...")
def _report_deck_file_for_version(version_id, storage_path):
    """_deck_file_for_version's exact equivalent, against REPORT_DECKS_BUCKET
    and its own cache directory -- this deck is a few hundred KB, not ~110MB,
    but the same "download once per version id, reuse across sessions and
    processes" reasoning still applies."""
    target = _REPORT_DECK_CACHE_DIR / f"v{version_id}_{Path(storage_path).name}"
    if target.exists() and target.stat().st_size > 0:
        return str(target)

    blob = get_client().storage.from_(REPORT_DECKS_BUCKET).download(storage_path)
    _REPORT_DECK_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


def report_master_deck(local_fallback_path):
    """(path, report_deck_version_id, warning) -- master_deck()'s exact
    equivalent for the attribution report template. Same contract: path is
    None only when Supabase is unreachable AND no local fallback exists."""
    row, warning = active_report_deck_version()
    if row is None:
        return _fallback_deck(local_fallback_path, warning)
    try:
        return _report_deck_file_for_version(row["id"], row["storage_path"]), row["id"], None
    except Exception as exc:
        return _fallback_deck(
            local_fallback_path,
            f"Couldn't download report deck version {row['id']} from Supabase "
            f"({describe_error(exc)})")


def clear_report_deck_cache():
    """report_master_deck()'s equivalent of clear_deck_cache() -- called
    after activating a new report deck version."""
    _report_deck_file_for_version.clear()


def scratch_dir(name, max_age_hours=6):
    """A temp directory for a page's uploads, with stale files swept out.

    Uploaded decks are ~44MiB each and nothing else ever deletes them, so on
    a long-lived instance with a small disk they'd accumulate one copy per
    distinct file anyone ever uploaded. Deliberately *not* used for the deck
    and case-study download caches: those are handed out by @st.cache_resource
    as paths, so deleting one underneath a live cache entry would break the
    next generate.
    """
    directory = Path(tempfile.gettempdir()) / name
    directory.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - max_age_hours * 3600
    for entry in directory.iterdir():
        try:
            if entry.is_file() and entry.stat().st_mtime < cutoff:
                entry.unlink()
        except OSError:
            pass  # another session is mid-upload, or the file just went away
    return directory


def deck_file(version_row):
    """Local path to a specific deck version's .pptx, downloading if needed.

    Same cache as the generate path, so opening a version on the update page
    doesn't re-download a deck this session already has.
    """
    return _deck_file_for_version(version_row["id"], version_row["storage_path"])


def clear_deck_cache():
    """Drop the cached deck download -- called after activating a new version
    so the next generate picks it up without a restart."""
    _deck_file_for_version.clear()


# ---------------------------------------------------------------------------
# Products / rate card (Stage 2)
# ---------------------------------------------------------------------------
def fetch_products():
    """(rows, warning) -- active product rows in sort order.

    rows is None, never [], whenever the caller should fall back: from the
    form's point of view an unreachable project and an empty table are the
    same thing -- there's no rate card to price with either way.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = (client.table("products").select("*")
                  .eq("active", True).order("sort_order").execute())
    except Exception as exc:
        return None, f"Couldn't load the rate card from Supabase ({describe_error(exc)})"
    rows = result.data or []
    if not rows:
        return None, "The products table in Supabase is empty"
    return rows, None


def upsert_products(rows):
    """Insert or update products by their `key`. Returns (count, error)."""
    client = get_client()
    if client is None:
        return 0, "Supabase isn't configured"
    try:
        result = client.table("products").upsert(rows, on_conflict="key").execute()
        return len(result.data or []), None
    except Exception as exc:
        return 0, describe_error(exc)


# ---------------------------------------------------------------------------
# App settings -- one generic table for small, admin-editable config records
# (the co-viewing coefficient today; a second setting is a new row, not a
# migration, same jsonb-bag philosophy as audiences.metrics)
# ---------------------------------------------------------------------------
def fetch_setting(key):
    """(value_dict, warning) -- the jsonb `value` for one settings row, or
    None (never {}) whenever the caller should fall back."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = client.table("app_settings").select("value").eq("key", key).execute()
    except Exception as exc:
        return None, f"Couldn't load the '{key}' setting from Supabase ({describe_error(exc)})"
    rows = result.data or []
    if not rows:
        return None, f"No '{key}' row in app_settings"
    return rows[0].get("value") or {}, None


def upsert_setting(key, value):
    """Insert or update one settings row by its `key`. Returns (ok, error)."""
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        client.table("app_settings").upsert(
            {"key": key, "value": value}, on_conflict="key").execute()
        return True, None
    except Exception as exc:
        return False, describe_error(exc)


# ---------------------------------------------------------------------------
# Audience catalog (Stage 3)
# ---------------------------------------------------------------------------
# The catalog is ~370 rows today but PostgREST caps a select at 1000 by
# default, so it's paged -- a silently truncated catalog would drop real
# segments from the finder and from what the model is allowed to pick.
PAGE_SIZE = 1000


def _fetch_all(client, table, columns="*", order_by=None):
    """Every row of a table, paged past PostgREST's default row limit."""
    rows, offset = [], 0
    while True:
        query = client.table(table).select(columns)
        if order_by:
            query = query.order(order_by)
        page = query.range(offset, offset + PAGE_SIZE - 1).execute().data or []
        rows.extend(page)
        if len(page) < PAGE_SIZE:
            return rows
        offset += PAGE_SIZE


def fetch_audiences():
    """(rows, warning) -- the active audience catalog. rows is None whenever
    the caller should fall back to parsing the local brochure PDF."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        rows = _fetch_all(client, "audiences", order_by="segment")
    except Exception as exc:
        return None, f"Couldn't load the audience catalog from Supabase ({describe_error(exc)})"
    rows = [row for row in rows if row.get("active", True)]
    if not rows:
        return None, "The audiences table in Supabase is empty"
    return rows, None


def fetch_audience_usage():
    """(rows, warning) -- the year-to-date audience booking log
    (`segment_string`, `delivered_impressions`, `is_custom`), for
    audience_evidence.py's booking-evidence panel. Same `(rows, warning)`
    convention as `fetch_audiences`: rows is None whenever the caller should
    fall back to the local `audience_usage_ytd.csv` instead, paged past
    PostgREST's default row limit like every other catalog-sized table.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        rows = _fetch_all(client, "audience_usage")
    except Exception as exc:
        return None, f"Couldn't load audience usage from Supabase ({describe_error(exc)})"
    if not rows:
        return None, "The audience_usage table in Supabase is empty"
    return rows, None


def upsert_audiences(rows, batch_size=500):
    """Insert or update catalog segments by `segment`. Returns (count, error)."""
    client = get_client()
    if client is None:
        return 0, "Supabase isn't configured"
    written = 0
    try:
        for start in range(0, len(rows), batch_size):
            batch = rows[start:start + batch_size]
            result = client.table("audiences").upsert(batch, on_conflict="segment").execute()
            written += len(result.data or [])
    except Exception as exc:
        return written, describe_error(exc)
    return written, None


def replace_audience_usage(rows, batch_size=500):
    """Replace the whole year-to-date usage log.

    Delete-then-insert rather than upsert: the log has no natural key (the
    same segment string can legitimately appear more than once) so re-running
    the seed would otherwise pile up duplicates.
    """
    client = get_client()
    if client is None:
        return 0, "Supabase isn't configured"
    written = 0
    try:
        client.table("audience_usage").delete().neq("id", 0).execute()
        for start in range(0, len(rows), batch_size):
            batch = rows[start:start + batch_size]
            result = client.table("audience_usage").insert(batch).execute()
            written += len(result.data or [])
    except Exception as exc:
        return written, describe_error(exc)
    return written, None


# ---------------------------------------------------------------------------
# Audience usage workbook versions (Stage 9) -- the deck_versions pattern
# applied to the YTD usage workbook Matt uploads periodically.
# ---------------------------------------------------------------------------
def list_audience_usage_versions():
    """(rows, warning) -- every registered usage-workbook version, newest first."""
    client = get_client()
    if client is None:
        return [], "Supabase isn't configured"
    try:
        result = (client.table("audience_usage_versions").select("*")
                  .order("uploaded_at", desc=True).execute())
        return result.data or [], None
    except Exception as exc:
        return [], f"Couldn't reach Supabase ({describe_error(exc)})"


def active_audience_usage_version():
    """(row, warning) for the single active audience_usage_versions row."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = client.table("audience_usage_versions").select("*").eq("active", True).limit(1).execute()
    except Exception as exc:
        return None, f"Couldn't reach Supabase ({describe_error(exc)})"
    rows = result.data or []
    if not rows:
        return None, "No active audience usage workbook is registered in Supabase"
    return rows[0], None


def upload_audience_usage_workbook(local_path, storage_path, notes=None, activate=True,
                                   confirmed_categories=None):
    """Upload a usage workbook (.xlsx) and register it as a version.

    Returns (row, error). Same two-step shape as upload_deck: the storage
    object goes up first, then a version row (active=False) -- only with
    activate=True does a separate step make it live, which is also where
    the actual audience_usage/audiences rebuild happens (see
    activate_audience_usage_version), not here. A registered-but-inactive
    version costs nothing and can be activated later without re-uploading.

    `confirmed_categories` passes through to activate_audience_usage_version
    -- see its own docstring; this is what a rep confirmed on the admin
    page for whatever the deterministic rules and the static registry
    couldn't resolve on their own.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        with open(local_path, "rb") as handle:
            client.storage.from_(AUDIENCE_USAGE_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                 "upsert": "true"},
            )
        row = {"storage_path": storage_path, "filename": Path(local_path).name,
               "notes": notes, "active": False}
        inserted = client.table("audience_usage_versions").insert(row).execute().data[0]
    except Exception as exc:
        return None, describe_error(exc)

    if activate:
        ok, error = activate_audience_usage_version(inserted["id"], confirmed_categories=confirmed_categories)
        if not ok:
            return inserted, error
        inserted["active"] = True
    return inserted, None


def activate_audience_usage_version(version_id, confirmed_categories=None):
    """Make one usage-workbook version the active one -- and actually
    REBUILD from it: download the stored .xlsx, re-parse it, merge the
    result into `audiences` (brochure-authoritative, workbook-additive --
    see audience_usage_import.derive_catalog_updates) and replace
    `audience_usage` wholesale. Unlike activate_deck_version, "activation"
    here can't be a bare flag flip -- the deck is served as-is from
    storage, but this data feeds two derived tables with no other way to
    pick up a refresh.

    `confirmed_categories` is {segment: {"category": "...", "exclude": bool}}
    -- what a rep confirmed on the admin page's Claude-suggestion review for
    components the deterministic rules and the committed registry both left
    unresolved. Merged OVER audience_usage_import.load_overrides() (the
    committed registry) as one-shot, session-specific overrides for this
    activation only: nothing here is written back to the registry file,
    because it doesn't need to be -- a confirmed category gets upserted
    into `audiences` below, and the NEXT workbook that mentions this same
    component will see it via the "existing catalog is authoritative"
    branch (derive_catalog_updates's own precedence) and never re-derive or
    re-ask for it again. This is also why a confirmed category survives a
    correction: whatever a rep confirms is exactly what gets stored.

    Clears the previous active row LAST, only once the rebuild itself has
    succeeded -- a parse or write failure midway leaves the previously
    active version's data as the last known-good state rather than
    flipping the flag to a version whose rebuild never completed.
    """
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        result = (client.table("audience_usage_versions").select("*")
                  .eq("id", version_id).limit(1).execute())
    except Exception as exc:
        return False, f"Couldn't reach Supabase ({describe_error(exc)})"
    rows = result.data or []
    if not rows:
        return False, f"No audience usage version {version_id} is registered"
    version_row = rows[0]

    import pandas as pd                             # lazy -- only needed for this reassignment step
    import audience_catalog                        # lazy -- avoids a module-level cycle with db.py
    import audience_usage_import as usage_import   # lazy -- openpyxl, only needed here

    try:
        blob = client.storage.from_(AUDIENCE_USAGE_BUCKET).download(version_row["storage_path"])
    except Exception as exc:
        return False, f"Couldn't download workbook version {version_id} ({describe_error(exc)})"
    try:
        workbook = usage_import.parse_workbook(io.BytesIO(blob))
    except usage_import.UsageImportError as exc:
        return False, f"Stored workbook version {version_id} no longer parses ({exc})"

    existing, _warning = fetch_audiences()
    # Category REASSIGNMENT (a segment that already has SOME category, but
    # not the most useful one for findability -- audience_category_
    # reassignments.csv) applies here too, not just at display time: without
    # this, a workbook that re-touches a reassigned brochure segment would
    # read its STALE category off Supabase as "existing" and re-persist it,
    # undoing the reassignment the very next time this segment is active.
    existing_df = audience_catalog.apply_category_reassignments(pd.DataFrame(existing or []))
    existing = existing_df.to_dict("records") if not existing_df.empty else []

    overrides = usage_import.load_overrides()
    for segment, confirmed in (confirmed_categories or {}).items():
        if confirmed.get("exclude"):
            overrides[usage_import.normalize(segment)] = {
                "action": "exclude_client", "category": "", "split_into": []}
        elif confirmed.get("category"):
            overrides[usage_import.normalize(segment)] = {
                "action": "categorize", "category": confirmed["category"], "split_into": []}
    report = usage_import.derive_catalog_updates(workbook, existing, overrides=overrides)

    catalog_rows = [
        {"segment": u.segment, "category": u.category, "rfp_selectable": u.rfp_selectable,
         "times_used": u.times_used, "metrics": {"impressions": u.impressions}, "active": True}
        for u in report.updates
    ]
    count, error = upsert_audiences(catalog_rows)
    if error:
        return False, f"Catalog merge failed after writing {count} row(s): {error}"

    count, error = replace_audience_usage(report.usage_rows)
    if error:
        return False, f"Usage log replace failed after the catalog merge succeeded: {error}"

    try:
        client.table("audience_usage_versions").update({"active": False}).eq("active", True).execute()
        client.table("audience_usage_versions").update({"active": True}).eq("id", version_id).execute()
    except Exception as exc:
        return False, f"Rebuild succeeded but activating version {version_id} failed: {describe_error(exc)}"
    return True, None


# ---------------------------------------------------------------------------
# Case study vault
# ---------------------------------------------------------------------------
def fetch_case_studies(active_only=True):
    """(rows, warning) -- every case study, newest first.

    Unlike the rate card and the catalog there's no local fallback here: a
    case study only exists in the vault. An empty list is a legitimate
    answer (nobody has added one yet), so this returns [] with a warning
    only when Supabase itself couldn't be reached.
    """
    client = get_client()
    if client is None:
        return [], "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        query = client.table("case_studies").select("*")
        if active_only:
            query = query.eq("active", True)
        result = query.order("date_added", desc=True).execute()
    except Exception as exc:
        return [], f"Couldn't load case studies from Supabase ({describe_error(exc)})"
    return result.data or [], None


def upload_case_study(local_path, filename, title, verticals, products,
                      summary, added_by, optimize=True):
    """Store a case study .pptx and register it. Returns (row, stats, error).

    Optimized and size-gated exactly like a master deck -- these are ordinary
    PowerPoint exports, so they carry the same oversized PNGs, and the
    storage ceiling applies to them just the same.
    """
    client = get_client()
    if client is None:
        return None, None, "Supabase isn't configured"

    upload_path, stats, error = (local_path, None, None)
    if optimize:
        upload_path, stats, error = prepare_deck_for_upload(
            local_path, limit=storage_limit_bytes(CASE_STUDIES_BUCKET))
        if error:
            return None, stats, error

    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    storage_path = f"{stamp}_{filename}"
    try:
        with open(upload_path, "rb") as handle:
            client.storage.from_(CASE_STUDIES_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                 "upsert": "true"},
            )
        row = {"filename": filename, "storage_path": storage_path, "title": title,
               "verticals": list(verticals), "products": list(products),
               "summary": summary, "added_by": added_by, "active": True}
        inserted = client.table("case_studies").insert(row).execute().data[0]
    except Exception as exc:
        return None, stats, describe_error(exc)
    finally:
        if optimize and upload_path and upload_path != local_path:
            try:
                os.unlink(upload_path)
            except OSError:
                pass
    return inserted, stats, None


def fetch_case_study(case_study_id):
    """(row, error) for one case study, active or not.

    Rebuilding an old proposal needs case studies it referenced even if
    they've since been deactivated -- deactivation stops one being *offered*,
    it doesn't rewrite what a client was already sent.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        result = client.table("case_studies").select("*").eq("id", case_study_id).limit(1).execute()
    except Exception as exc:
        return None, describe_error(exc)
    rows = result.data or []
    return (rows[0], None) if rows else (None, "No longer in the vault")


def update_case_study(case_study_id, **fields):
    """Patch one case study's metadata. Returns (row, error).

    Tags are corrected here rather than by re-uploading: a mis-tagged
    vertical is a one-word fix, and making someone re-upload a 40MB deck to
    change it guarantees nobody bothers.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    allowed = {"title", "verticals", "products", "summary", "active", "added_by"}
    payload = {k: v for k, v in fields.items() if k in allowed}
    if not payload:
        return None, "Nothing to update"
    try:
        result = client.table("case_studies").update(payload).eq("id", case_study_id).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0], None


def case_study_cached_path(case_study_id, storage_path):
    """The local path if this case study has already been fetched, else None.

    Lets a caller offer an instant download without triggering a fetch. That
    matters on the vault browser: rendering a download button needs the file's
    bytes, and doing that for every visible row would pull the whole vault
    (22 files, ~42MiB) out of storage just to draw buttons nobody clicked.
    """
    target = _CASE_STUDY_CACHE_DIR / f"{case_study_id}_{Path(storage_path).name}"
    return str(target) if target.exists() and target.stat().st_size > 0 else None


@st.cache_resource(show_spinner="Fetching case study...")
def case_study_file(case_study_id, storage_path):
    """Local path to one case study's .pptx, cached on its id for the same
    reason the master deck is -- assembly needs a real file on disk to open,
    and a proposal can pull in several."""
    target = _CASE_STUDY_CACHE_DIR / f"{case_study_id}_{Path(storage_path).name}"
    if target.exists() and target.stat().st_size > 0:
        return str(target)

    blob = get_client().storage.from_(CASE_STUDIES_BUCKET).download(storage_path)
    _CASE_STUDY_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


def upload_case_study_images(case_study_id, image_paths, width):
    """Store a case study's pre-rendered slides. Returns (count, error).

    Generated locally and uploaded, never rendered on demand: rendering needs
    PowerPoint and the Proxima Nova brand font, and the deployed app has
    neither -- it would quietly substitute fonts and produce images that
    aren't what the case study looks like. See render_case_study_images.py.

    Ordering is the slide order and is carried by the array, not by the file
    names, so a re-run that produces a different count can't interleave with
    what was there before -- the row is replaced wholesale.
    """
    client = get_client()
    if client is None:
        return 0, "Supabase isn't configured"
    stored = []
    try:
        for index, path in enumerate(image_paths, start=1):
            key = f"images/{case_study_id}/slide_{index:03d}{Path(path).suffix}"
            blob = Path(path).read_bytes()
            client.storage.from_(CASE_STUDIES_BUCKET).upload(
                key, blob, {"content-type": "image/jpeg", "upsert": "true"})
            stored.append(key)
        client.table("case_studies").update({
            "slide_images": stored,
            "image_width": int(width),
            "images_generated_at": datetime.now(timezone.utc).isoformat(),
        }).eq("id", case_study_id).execute()
    except Exception as exc:                                     # noqa: BLE001
        return 0, describe_error(exc)
    return len(stored), None


@st.cache_resource(show_spinner="Fetching case study slides...")
def case_study_images(case_study_id, storage_paths):
    """Local paths to one case study's rendered slides, in order.

    `storage_paths` is a tuple so it can be a cache key -- a list isn't
    hashable, and this is fetched once per session per case study exactly
    like the .pptx is.
    """
    local = []
    directory = _CASE_STUDY_CACHE_DIR / "images" / str(case_study_id)
    directory.mkdir(parents=True, exist_ok=True)
    for key in storage_paths:
        target = directory / Path(key).name
        if not (target.exists() and target.stat().st_size > 0):
            blob = get_client().storage.from_(CASE_STUDIES_BUCKET).download(key)
            partial = target.with_suffix(target.suffix + ".part")
            partial.write_bytes(blob)
            partial.replace(target)
        local.append(str(target))
    return local


def case_study_render_path(case_study):
    """Which path a case study will take into a deck: "images" or "copy".

    One place decides it, so the vault browser's coverage column and the
    generator can't disagree about what a proposal will actually contain.
    """
    return "images" if (case_study or {}).get("slide_images") else "copy"


# ---------------------------------------------------------------------------
# Slide vault
#
# The case study vault again, one slide at a time: a contributor uploads a
# .pptx and picks individual slides out of it, so several rows can share one
# storage_path (slide_index tells them apart). The deck is stored whole and
# never split apart -- extracting a single slide into its own .pptx is
# exactly the cross-deck copy problem that produced five distinct corruption
# bugs on the case-study vault (CLAUDE.md), so a vault slide is grafted the
# same way a case study is, by index, out of the deck it always lived in.
# ---------------------------------------------------------------------------
def fetch_slide_vault(active_only=True):
    """(rows, warning) -- every vault slide, newest first.

    Same no-local-fallback shape as fetch_case_studies: an empty vault is a
    legitimate answer, only an unreachable Supabase gets a warning.
    """
    client = get_client()
    if client is None:
        return [], "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        query = client.table("slide_vault").select("*")
        if active_only:
            query = query.eq("active", True)
        result = query.order("date_added", desc=True).execute()
    except Exception as exc:
        return [], f"Couldn't load the slide vault from Supabase ({describe_error(exc)})"
    return result.data or [], None


def upload_slide_vault(local_path, filename, picks, added_by, optimize=True):
    """Store a .pptx once and register one row per picked slide.

    `picks` is a list of dicts, each carrying slide_index, title, verticals,
    products, purpose, placement, summary for one ticked slide. Returns
    (rows, stats, error) -- rows is [] on failure, and nothing is written if
    optimizing/uploading the deck itself fails, same all-or-nothing shape
    upload_case_study uses.
    """
    client = get_client()
    if client is None:
        return [], None, "Supabase isn't configured"
    if not picks:
        return [], None, "No slides were picked"

    upload_path, stats, error = (local_path, None, None)
    if optimize:
        upload_path, stats, error = prepare_deck_for_upload(
            local_path, limit=storage_limit_bytes(SLIDE_VAULT_BUCKET))
        if error:
            return [], stats, error

    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    storage_path = f"{stamp}_{filename}"
    try:
        with open(upload_path, "rb") as handle:
            client.storage.from_(SLIDE_VAULT_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                 "upsert": "true"},
            )
        rows = [{
            "filename": filename, "storage_path": storage_path,
            "slide_index": int(pick["slide_index"]), "title": pick.get("title"),
            "verticals": list(pick.get("verticals") or []),
            "products": list(pick.get("products") or []),
            "purpose": pick.get("purpose"),
            "placement": pick.get("placement") or "before_plan",
            "summary": pick.get("summary"), "added_by": added_by, "active": True,
        } for pick in picks]
        inserted = client.table("slide_vault").insert(rows).execute().data
    except Exception as exc:
        return [], stats, describe_error(exc)
    finally:
        if optimize and upload_path and upload_path != local_path:
            try:
                os.unlink(upload_path)
            except OSError:
                pass
    return inserted or [], stats, None


def fetch_slide_vault_entry(entry_id):
    """(row, error) for one vault slide, active or not -- ignores `active`
    for the same reason fetch_case_study does: rebuilding an old proposal
    needs a slide it referenced even if it's since been deactivated."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        result = client.table("slide_vault").select("*").eq("id", entry_id).limit(1).execute()
    except Exception as exc:
        return None, describe_error(exc)
    rows = result.data or []
    return (rows[0], None) if rows else (None, "No longer in the vault")


def update_slide_vault_entry(entry_id, **fields):
    """Patch one vault slide's metadata. Returns (row, error)."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    allowed = {"title", "verticals", "products", "purpose", "placement",
               "summary", "active", "added_by"}
    payload = {k: v for k, v in fields.items() if k in allowed}
    if not payload:
        return None, "Nothing to update"
    try:
        result = client.table("slide_vault").update(payload).eq("id", entry_id).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0], None


def delete_slide_vault_entry(entry_id):
    """Hard-delete one slide_vault row. Returns (ok, error).

    Unlike case studies -- which have no hard delete at all, only
    `update_slide_vault_entry(..., active=False)` -- this exists because
    deactivation doesn't cover "this shouldn't be in a shared library at
    all" (a test upload, a wrong file, something off-brand). But a hard
    delete is a real risk case studies' design deliberately avoids:
    `fetch_slide_vault_entry`'s own "ignore active, rebuild needs it" escape
    hatch (the thing that lets "Rebuild as presented" reproduce a proposal
    that used a since-deactivated slide) only works because the ROW still
    exists. A hard-deleted row breaks that outright for any proposal that
    used it. So this checks EVERY logged proposal's own `vault_slides` list
    first (client-side, same shape as the History page's own usage
    reads -- no JSONB containment query anywhere else in this app to
    match) and refuses, unchanged, if any proposal references this id. A
    deck that can't be rebuilt is worse than a vault with a stale entry in
    it, so refusing is the safe default -- deactivate instead for anything
    that's actually been used.

    The underlying storage file is removed only if no OTHER slide_vault row
    still shares its `storage_path` -- several rows can come from one
    uploaded deck, picked slide by slide, and one row's delete must never
    take a sibling's still-referenced file with it.
    """
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    row, error = fetch_slide_vault_entry(entry_id)
    if error:
        return False, error

    proposals, perror = fetch_proposals(limit=10000)
    if perror:
        return False, f"Couldn't confirm this slide isn't in use ({perror}) -- not deleted."
    using = [p.get("client_name") or "(no client name)" for p in (proposals or [])
            if any(v.get("id") == entry_id
                  for v in (p.get("form_json") or {}).get("vault_slides") or [])]
    if using:
        names = ", ".join(using[:5]) + ("..." if len(using) > 5 else "")
        return False, (f"Can't delete -- used by {len(using)} logged proposal(s) ({names}). "
                       f"Deactivate it instead so it stops being offered; rebuilding those "
                       f"proposals still needs this row to exist.")

    storage_path = row.get("storage_path")
    try:
        client.table("slide_vault").delete().eq("id", entry_id).execute()
    except Exception as exc:
        return False, describe_error(exc)

    if storage_path:
        siblings, serror = fetch_slide_vault(active_only=False)
        still_shared = bool(siblings) and any(
            s.get("storage_path") == storage_path for s in siblings)
        if not still_shared:
            try:
                client.storage.from_(SLIDE_VAULT_BUCKET).remove([storage_path])
            except Exception as exc:
                return True, (f"The vault entry was deleted, but the stored file couldn't be "
                             f"removed ({describe_error(exc)}). It's now orphaned.")
    return True, None


def slide_vault_cached_path(storage_path):
    """The local path if this source deck has already been fetched, else
    None -- keyed on storage_path, not an entry id, so several rows sharing
    one upload share the cache too."""
    target = _SLIDE_VAULT_CACHE_DIR / Path(storage_path).name
    return str(target) if target.exists() and target.stat().st_size > 0 else None


@st.cache_resource(show_spinner="Fetching vault slide's source deck...")
def slide_vault_file(storage_path):
    """Local path to a vault slide's source .pptx, cached on storage_path so
    several rows picked from the same upload fetch it once."""
    target = _SLIDE_VAULT_CACHE_DIR / Path(storage_path).name
    if target.exists() and target.stat().st_size > 0:
        return str(target)

    blob = get_client().storage.from_(SLIDE_VAULT_BUCKET).download(storage_path)
    _SLIDE_VAULT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


def upload_slide_vault_image(entry_id, image_path, width):
    """Store one vault slide's pre-rendered image. Returns (ok, error).

    Singular, unlike upload_case_study_images: a vault row is already one
    slide, so there's one image, not an ordered array.
    """
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        key = f"images/{entry_id}{Path(image_path).suffix}"
        blob = Path(image_path).read_bytes()
        client.storage.from_(SLIDE_VAULT_BUCKET).upload(
            key, blob, {"content-type": "image/jpeg", "upsert": "true"})
        client.table("slide_vault").update({
            "slide_image": key,
            "image_width": int(width),
            "image_generated_at": datetime.now(timezone.utc).isoformat(),
        }).eq("id", entry_id).execute()
    except Exception as exc:                                      # noqa: BLE001
        return False, describe_error(exc)
    return True, None


@st.cache_resource(show_spinner="Fetching vault slide...")
def slide_vault_image(entry_id, storage_key):
    """Local path to one vault slide's rendered image."""
    directory = _SLIDE_VAULT_CACHE_DIR / "images"
    directory.mkdir(parents=True, exist_ok=True)
    target = directory / f"{entry_id}{Path(storage_key).suffix}"
    if not (target.exists() and target.stat().st_size > 0):
        blob = get_client().storage.from_(SLIDE_VAULT_BUCKET).download(storage_key)
        partial = target.with_suffix(target.suffix + ".part")
        partial.write_bytes(blob)
        partial.replace(target)
    return str(target)


def slide_vault_render_path(row):
    """Which path a vault slide will take into a deck: "image" or "copy".

    One place decides it, the counterpart of case_study_render_path, so the
    vault browser's coverage column and the generator can't disagree about
    what a proposal will actually contain.
    """
    return "image" if (row or {}).get("slide_image") else "copy"


# ---------------------------------------------------------------------------
# Proposal history (Stage 4)
# ---------------------------------------------------------------------------
def _json_default(obj):
    if hasattr(obj, "item"):  # numpy scalar out of the media-plan grid
        return obj.item()
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    return str(obj)


def _json_safe(value):
    """Coerce a payload into something jsonb will accept.

    The media plan rows come back from a Streamlit data editor carrying numpy
    scalars, and empty cells arrive as NaN -- which json.dumps happily writes
    as a bare `NaN` literal that isn't valid JSON and that Postgres rejects.
    Both are normalized here rather than at the call site, so any future
    caller gets the same treatment.
    """
    if isinstance(value, float) and (value != value or value in (float("inf"), float("-inf"))):
        return None
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(v) for v in value]
    if isinstance(value, (str, int, bool)) or value is None:
        return value
    if isinstance(value, float):
        return value
    return _json_safe(json.loads(json.dumps(value, default=_json_default)))


def log_proposal(client_name, vertical, market, form_json,
                 deck_version_id=None, output_filename=None,
                 parent_proposal_id=None, revision_label=None, logo_storage_path=None,
                 created_by=None, target_dmas=None):
    """Record one generated proposal. Returns (row_id, error).

    Always an INSERT, never an update: history is append-only, so
    regenerating a proposal loaded from the History page writes a new row
    carrying parent_proposal_id back to its source rather than overwriting
    what a client was actually sent.

    A failed log must never take a successfully generated deck down with it,
    so this reports rather than raises and the caller shows it as a caption,
    not an error.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    row = {
        "client_name": client_name,
        "vertical": vertical,
        "market": market,
        "form_json": _json_safe(form_json),
        "deck_version_id": deck_version_id,
        "output_filename": output_filename,
        "parent_proposal_id": parent_proposal_id,
        "revision_label": revision_label,
        "logo_storage_path": logo_storage_path,
        "created_by": created_by,
        # The DMAs targeted, as market_profiles keys, in the order the rep
        # picked them -- which is the order their profile slides appear in
        # the deck, so the ordering is data rather than presentation. Its own
        # column as well as being in form_json, so "which markets are we
        # selling into" is a containment query rather than a scan of every
        # stored form. Empty is the honest value for every proposal logged
        # before target DMAs existed.
        #
        # The older scalar `target_dma` column is frozen and deliberately NOT
        # written here: it was backfilled into this array once, and two
        # copies of one fact drifting apart is a bug this project has already
        # paid for more than once.
        "target_dmas": list(target_dmas or []),
    }
    try:
        result = client.table("proposals").insert(row).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0].get("id"), None


def fetch_proposals(limit=500):
    """(rows, warning) for the History page, newest first.

    Returns None rather than [] when Supabase can't answer, so the caller can
    tell "no history yet" from "no backend" -- the same distinction every
    other loader here draws.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = (client.table("proposals").select("*")
                  .order("generated_at", desc=True).limit(limit).execute())
    except Exception as exc:
        return None, f"Couldn't load proposal history ({describe_error(exc)})"
    return result.data or [], None


def fetch_proposal(proposal_id):
    """(row, error) for one proposal, form_json included."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        result = client.table("proposals").select("*").eq("id", proposal_id).limit(1).execute()
    except Exception as exc:
        return None, describe_error(exc)
    rows = result.data or []
    if not rows:
        return None, "That proposal no longer exists."
    return rows[0], None


def fetch_deck_version(version_id):
    """(row, error) for one deck version, active or not.

    "Rebuild as presented" needs the deck the proposal was originally built
    from, which is by definition usually NOT the active one. Storage keeps
    every version and deck_versions keeps every row, so an old version stays
    fetchable indefinitely.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        result = client.table("deck_versions").select("*").eq("id", version_id).limit(1).execute()
    except Exception as exc:
        return None, describe_error(exc)
    rows = result.data or []
    if not rows:
        return None, f"Deck version {version_id} is no longer registered."
    return rows[0], None


def delete_proposal(proposal_id):
    """Hard-delete one history row and its attached file. Returns (ok, error).

    Deliberately not a soft delete and deliberately not offered in bulk: this
    exists for junk and test rows, and anything reachable in bulk is a way to
    lose real history by accident. Descendants are NOT cascaded -- the
    parent_proposal_id FK is ON DELETE SET NULL, so deleting a test row that a
    real proposal happens to descend from orphans the link instead of taking
    the real one with it.
    """
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    row, error = fetch_proposal(proposal_id)
    if error:
        return False, error
    # Both blobs this row owns: the attached final and the stored logo.
    owned = [row.get(key) for key in ("file_storage_path", "logo_storage_path") if row.get(key)]
    try:
        client.table("proposals").delete().eq("id", proposal_id).execute()
    except Exception as exc:
        return False, describe_error(exc)
    if owned:
        # The row is already gone; a failed object delete leaves an orphan
        # blob, which costs quota but breaks nothing, so it's reported rather
        # than raised.
        try:
            client.storage.from_(PROPOSAL_FILES_BUCKET).remove(owned)
        except Exception as exc:
            return True, (f"The history row was deleted, but {len(owned)} stored file(s) couldn't "
                          f"be removed ({describe_error(exc)}). They're now orphaned.")
    return True, None


def attach_proposal_file(proposal_id, local_path, filename, note=None):
    """Attach a hand-edited final .pptx to a history row. (row, stats, error).

    Optimized and size-gated through the same path as a master deck or case
    study -- it's an ordinary PowerPoint export carrying the same oversized
    PNGs, and the storage ceiling applies identically.
    """
    client = get_client()
    if client is None:
        return None, None, "Supabase isn't configured"

    upload_path, stats, error = prepare_deck_for_upload(
        local_path, limit=storage_limit_bytes(PROPOSAL_FILES_BUCKET))
    if error:
        return None, stats, error

    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    storage_path = f"{proposal_id}/{stamp}_{filename}"
    try:
        with open(upload_path, "rb") as handle:
            client.storage.from_(PROPOSAL_FILES_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                 "upsert": "true"},
            )
        payload = {
            "file_storage_path": storage_path,
            "file_attached_at": datetime.now(timezone.utc).isoformat(),
            "file_note": note,
        }
        result = client.table("proposals").update(payload).eq("id", proposal_id).execute()
    except Exception as exc:
        return None, stats, describe_error(exc)
    finally:
        if upload_path and upload_path != local_path:
            try:
                os.unlink(upload_path)
            except OSError:
                pass
    return (result.data or [{}])[0], stats, None


@st.cache_resource(show_spinner="Fetching attached file...")
def proposal_file(proposal_id, storage_path):
    """Local path to a row's attached final .pptx, cached on its id."""
    target = _PROPOSAL_FILE_CACHE_DIR / f"{proposal_id}_{Path(storage_path).name}"
    if target.exists() and target.stat().st_size > 0:
        return str(target)

    blob = get_client().storage.from_(PROPOSAL_FILES_BUCKET).download(storage_path)
    _PROPOSAL_FILE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


def team_member_key(name):
    """The deduplication key for a person's name: trimmed, case-folded, and
    with runs of internal whitespace collapsed. "matt", "Matt" and " Matt "
    are one person."""
    return " ".join((name or "").split()).lower()


def fetch_team_members(active_only=True):
    """(names, warning) -- the identity picker's options, alphabetical.

    Returns None rather than [] when Supabase can't answer, so the caller can
    tell "nobody has signed in yet" from "no backend", same as every other
    loader here.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        query = client.table("team_members").select("name, active").order("name")
        if active_only:
            query = query.eq("active", True)
        result = query.execute()
    except Exception as exc:
        return None, f"Couldn't load the team list ({describe_error(exc)})"
    return [row["name"] for row in (result.data or [])], None


def add_team_member(name):
    """Add a person to the team list. Returns (stored_name, error).

    Deduplicated on name_key, so adding a name that already exists in any
    capitalisation is a no-op that returns the EXISTING display name rather
    than an error or a second row -- two people called Matt is the failure
    mode worth designing against, and someone re-typing their own name
    shouldn't look like a mistake.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    display = " ".join((name or "").split())
    if not display:
        return None, "Enter a name first."
    key = team_member_key(display)
    try:
        existing = (client.table("team_members").select("name")
                    .eq("name_key", key).limit(1).execute())
        if existing.data:
            return existing.data[0]["name"], None
        client.table("team_members").insert(
            {"name": display, "name_key": key, "active": True}).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return display, None


def update_proposal(proposal_id, **fields):
    """Patch one history row's editable metadata. Returns (row, error).

    Deliberately narrow: only the label and the logo pointer. History is
    append-only, so nothing that describes *what was generated* -- form_json,
    deck_version_id, generated_at -- is patchable here. A revision label is
    an annotation about a row, not a claim about what the client received,
    which is why it's the one thing editable after the fact.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    allowed = {"revision_label", "logo_storage_path", "file_note"}
    payload = {k: v for k, v in fields.items() if k in allowed}
    if not payload:
        return None, "Nothing to update"
    try:
        result = client.table("proposals").update(payload).eq("id", proposal_id).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0], None


def upload_proposal_logo(proposal_key, data, filename):
    """Store a client logo for a proposal. Returns (storage_path, error).

    Images, not decks, so this skips prepare_deck_for_upload -- a logo is a
    few KB and there's nothing to optimize. Uploaded under a `logos/` prefix
    so bucket_usage's two-level walk still reports them, and so attached
    finals and logos are distinguishable in the storage browser.

    Failure is reported, never raised: a logo that didn't upload must not
    take down a deck that generated fine.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    safe = Path(filename or "logo.png").name
    storage_path = f"logos/{proposal_key}_{stamp}_{safe}"
    content_type = "image/png" if safe.lower().endswith(".png") else "image/jpeg"
    try:
        client.storage.from_(PROPOSAL_FILES_BUCKET).upload(
            storage_path, data, {"content-type": content_type, "upsert": "true"})
    except Exception as exc:
        return None, describe_error(exc)
    return storage_path, None


@st.cache_resource(show_spinner=False)
def proposal_logo(storage_path):
    """Local path to a stored logo, downloading it once per session.

    Keyed on the storage path alone -- a logo is immutable once written (each
    upload gets its own timestamped key), so there's no version to invalidate
    against.
    """
    target = _PROPOSAL_FILE_CACHE_DIR / f"logo_{Path(storage_path).name}"
    if target.exists() and target.stat().st_size > 0:
        return str(target)
    blob = get_client().storage.from_(PROPOSAL_FILES_BUCKET).download(storage_path)
    _PROPOSAL_FILE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


# Storage's list() caps at 100 objects per call, a DIFFERENT default from
# PostgREST's 1000 for tables -- so a bucket is paged too. Found when the
# market profile bucket's 206 images reported as 100 objects and half the
# real size: the count was wrong, the upload was fine. It matters because
# bucket_usage exists to notice the free tier's 1GB *before* it's reached,
# and a usage figure that silently stops counting at 100 files per prefix
# reports least when there's most to report.
STORAGE_PAGE_SIZE = 1000


def _list_objects(client, bucket, path=""):
    """Every object under one prefix, paged past storage's 100-row default."""
    entries, offset = [], 0
    while True:
        page = client.storage.from_(bucket).list(
            path, {"limit": STORAGE_PAGE_SIZE, "offset": offset}) or []
        entries.extend(page)
        if len(page) < STORAGE_PAGE_SIZE:
            return entries
        offset += STORAGE_PAGE_SIZE


def bucket_usage(bucket=PROPOSAL_FILES_BUCKET):
    """(total_bytes, file_count, warning) for one bucket.

    Attached files are the only thing in this app that isn't regenerable from
    a recipe, so they're the only thing that grows without bound. Surfaced on
    the History page so the free tier's 1GB is noticed before it's reached
    rather than after.
    """
    client = get_client()
    if client is None:
        return 0, 0, "Supabase isn't configured"
    total = count = 0
    try:
        # Files are stored under one prefix per proposal id, so this is a
        # two-level walk rather than a flat list.
        for entry in _list_objects(client, bucket) or []:
            name = entry.get("name")
            if not name:
                continue
            meta = entry.get("metadata") or {}
            if meta.get("size") is not None:
                total += int(meta["size"])
                count += 1
                continue
            for child in _list_objects(client, bucket, name) or []:
                child_meta = child.get("metadata") or {}
                if child_meta.get("size") is not None:
                    total += int(child_meta["size"])
                    count += 1
    except Exception as exc:
        return 0, 0, f"Couldn't measure {bucket} usage ({describe_error(exc)})"
    return total, count, None


# ---------------------------------------------------------------------------
# In-app feedback (BACKLOG.md's "Report an issue" popover)
# ---------------------------------------------------------------------------
# A few-KB state snapshot, same shape as proposals.form_json -- no bucket
# needed. status is a plain string (not a boolean) so a third state
# ("wontfix") is a value, not a migration.
def submit_feedback(category, notes, page, state, created_by=None):
    """Record one feedback report. Returns (row_id, error).

    `state` is app.capture_feedback_state()'s raw Python dict -- JSON-safed
    here, same as log_proposal does for form_json, so the caller never has
    to think about date/set serialization.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    row = {
        "category": category,
        "notes": notes,
        "page": page,
        "created_by": created_by,
        "state_json": _json_safe(state or {}),
    }
    try:
        result = client.table("feedback").insert(row).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0].get("id"), None


def fetch_feedback(status=None, category=None, limit=500):
    """(rows, warning) for the Feedback reports admin page, newest first.

    Returns None (not []) when Supabase can't answer, so the caller can
    tell "no reports yet" from "no backend" -- the same distinction every
    other loader in this file draws.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        query = client.table("feedback").select("*").order("created_at", desc=True).limit(limit)
        if status:
            query = query.eq("status", status)
        if category:
            query = query.eq("category", category)
        result = query.execute()
    except Exception as exc:
        return None, f"Couldn't load feedback reports ({describe_error(exc)})"
    return result.data or [], None


def update_feedback_status(feedback_id, status):
    """Mark one report open/closed (or any other status value). Returns
    (ok, error). resolved_at is cleared on reopen rather than left stale
    from a prior close."""
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    row = {"status": status,
           "resolved_at": datetime.now(timezone.utc).isoformat() if status != "open" else None}
    try:
        client.table("feedback").update(row).eq("id", feedback_id).execute()
        return True, None
    except Exception as exc:
        return False, describe_error(exc)


def count_open_feedback():
    """(count, warning) -- cheap count for the sidebar badge."""
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        result = client.table("feedback").select("id", count="exact").eq("status", "open").execute()
    except Exception as exc:
        return None, describe_error(exc)
    return result.count or 0, None


# ---------------------------------------------------------------------------
# Advertisers (Stage 13) + Attribution reports (Stage 14)
# ATTRIBUTION_REPORT_PLAN.md Phase 2 -- the canonical client spine and the
# report rows that join to it.
# ---------------------------------------------------------------------------

def advertiser_name_key(name):
    """The dedup key for an advertiser row: trimmed, case-folded, internal
    whitespace collapsed -- exactly team_member_key's own convention,
    applied to a company name instead of a person's. Kept separate from
    advertiser_matching.normalize_name (which also strips legal suffixes
    and all punctuation for FUZZY scoring): this is the exact-dedup key
    stored in the database, deliberately closer to the name as typed.
    """
    return " ".join((name or "").split()).lower()


def fetch_advertisers(limit=1000):
    """(rows, warning) -- every advertiser, for a caller to rank against
    with advertiser_matching.find_candidates. None (not []) when Supabase
    can't answer, so "no advertisers yet" and "no backend" stay distinct.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        result = (client.table("advertisers").select("*")
                  .order("canonical_name").limit(limit).execute())
    except Exception as exc:
        return None, f"Couldn't load advertisers ({describe_error(exc)})"
    return result.data or [], None


def create_advertiser(name):
    """Add a new advertiser. Returns (row, error).

    Deduplicated on name_key -- "creating" an advertiser that already
    exists (in any capitalisation/whitespace) returns the EXISTING row
    rather than a second one or an error, same convention as
    add_team_member: a rep confirming a match they already confirmed once
    shouldn't fork the spine this table exists to be.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    display = " ".join((name or "").split())
    if not display:
        return None, "Enter an advertiser name first."
    key = advertiser_name_key(display)
    try:
        existing = (client.table("advertisers").select("*")
                    .eq("name_key", key).limit(1).execute())
        if existing.data:
            return existing.data[0], None
        result = client.table("advertisers").insert(
            {"canonical_name": display, "name_key": key}).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0], None


def set_advertiser_vertical(advertiser_id, vertical):
    """Best-effort write of an advertiser's own `vertical` column
    (Highlights/Takeaways rework, Stage 16 DDL) -- so a LATER standalone
    (no-proposal-linked) report can resolve the same advertiser's vertical
    without asking the rep again. Returns (ok, error); a failure here is a
    caption, never a blocker, same convention as `log_attribution_report`.
    Requires the Stage 16 DDL (`advertisers.vertical`) to have been pasted
    into Supabase -- degrades to an error string, not a crash, if it
    hasn't (the column simply doesn't exist yet on an older schema)."""
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    if not advertiser_id or not vertical:
        return False, "No advertiser or vertical to save"
    try:
        client.table("advertisers").update({"vertical": vertical}).eq("id", advertiser_id).execute()
        return True, None
    except Exception as exc:
        return False, describe_error(exc)


def link_proposal_advertiser(proposal_id, advertiser_id):
    """Set a proposal's advertiser_id after the fact (a rep confirming a
    match for a proposal logged before this table existed, or correcting
    one). Returns (ok, error)."""
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        client.table("proposals").update(
            {"advertiser_id": advertiser_id}).eq("id", proposal_id).execute()
        return True, None
    except Exception as exc:
        return False, describe_error(exc)


def log_attribution_report(advertiser_id, proposal_id, report_json, status="parsed",
                           created_by=None):
    """Record one attribution report. Returns (row_id, error).

    Always an INSERT, never an update -- same append-only convention as
    log_proposal, so re-parsing or re-confirming an export never silently
    overwrites an earlier report row.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    row = {
        "advertiser_id": advertiser_id,
        "proposal_id": proposal_id,
        "report_json": _json_safe(report_json or {}),
        "status": status,
        "created_by": created_by,
    }
    try:
        result = client.table("attribution_reports").insert(row).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0].get("id"), None


def fetch_attribution_reports(advertiser_id=None, limit=500):
    """(rows, warning) newest first, optionally scoped to one advertiser --
    the same None-vs-[] distinction every other loader here draws.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        query = (client.table("attribution_reports").select("*")
                .order("created_at", desc=True).limit(limit))
        if advertiser_id:
            query = query.eq("advertiser_id", advertiser_id)
        result = query.execute()
    except Exception as exc:
        return None, f"Couldn't load attribution reports ({describe_error(exc)})"
    return result.data or [], None


# ---------------------------------------------------------------------------
# Market viewer profiles (Stage 7)
# ---------------------------------------------------------------------------
# The fallback here is unusually good: market_profiles.build_rows() derives
# the whole selectable set from a file that IS in the repo, so an unreachable
# Supabase costs the profile IMAGES and nothing else -- every DMA is still
# offered and a proposal can still name the one it targets. That's why this
# loader returns None (fall back) rather than raising, like every other one.
def fetch_market_profiles():
    """(rows, warning) -- selectable target markets.

    None means fall back to market_profiles.build_rows(); an unreachable
    project and an empty table are the same thing to the picker.

    Ordering is deliberately NOT done here. Nielsen rank isn't a column, and
    the caller sorts both these rows and the local fallback rows through the
    one market_profiles.sort_rows(), so the two paths can't disagree about
    what order a rep sees. `key` only makes the fetch deterministic.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured (no SUPABASE_URL / SUPABASE_SERVICE_KEY)"
    try:
        rows = _fetch_all(client, "market_profiles", order_by="key")
    except Exception as exc:
        return None, f"Couldn't load market profiles from Supabase ({describe_error(exc)})"
    rows = [r for r in rows if r.get("active", True)]
    if not rows:
        return None, "The market_profiles table in Supabase is empty"
    return rows, None


def upsert_market_profiles(rows):
    """Insert or update market profiles by `key`. Returns (count, error).

    Upsert rather than replace, deliberately: image_path and stats are
    populated by separate passes, and a re-seed that wiped them would mean
    re-uploading 35MiB of images to fix a typo in a label. The seed carries
    identity and ordering only -- it never sends image or stats columns, so
    it cannot null them.
    """
    client = get_client()
    if client is None:
        return 0, "Supabase isn't configured"
    try:
        result = client.table("market_profiles").upsert(
            rows, on_conflict="key").execute()
        return len(result.data or []), None
    except Exception as exc:                                     # noqa: BLE001
        return 0, describe_error(exc)


def upload_market_profile_images(pairs, width):
    """Store the profile slides. `pairs` is [(key, local_path), ...].

    Returns (count, error). Each image is keyed by the market's own slug, so
    a re-upload replaces in place and the object name says which market it
    is when someone is looking at the bucket rather than the table.
    """
    client = get_client()
    if client is None:
        return 0, "Supabase isn't configured"
    stored = 0
    try:
        for key, path in pairs:
            object_key = f"markets/{key}{Path(path).suffix}"
            client.storage.from_(MARKET_PROFILES_BUCKET).upload(
                object_key, Path(path).read_bytes(),
                {"content-type": "image/jpeg", "upsert": "true"})
            client.table("market_profiles").update({
                "image_path": object_key,
                "image_width": int(width),
                "images_generated_at": datetime.now(timezone.utc).isoformat(),
            }).eq("key", key).execute()
            stored += 1
    except Exception as exc:                                     # noqa: BLE001
        return stored, describe_error(exc)
    return stored, None


@st.cache_resource(show_spinner="Fetching the market profile...")
def market_profile_image(key, storage_path):
    """A local path to one market's profile slide, or None.

    Cached per market for the session, like the case study slides -- a rep
    comparing markets shouldn't re-download one they've already looked at.
    Returns None rather than raising when the market has no slide or the
    fetch fails: the caller has something to say about that (see the picker),
    and a missing profile must never be able to take a proposal down.
    """
    if not storage_path:
        return None
    _MARKET_PROFILE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    target = _MARKET_PROFILE_CACHE_DIR / f"{key}{Path(storage_path).suffix}"
    if target.exists() and target.stat().st_size > 0:
        return str(target)
    try:
        blob = get_client().storage.from_(MARKET_PROFILES_BUCKET).download(storage_path)
    except Exception:                                            # noqa: BLE001
        return None
    partial = target.with_suffix(target.suffix + ".part")
    partial.write_bytes(blob)
    partial.replace(target)
    return str(target)


# ---------------------------------------------------------------------------
# Admin / setup helpers. Used by setup_supabase.py and the admin pages, not
# on the ordinary form path.
# ---------------------------------------------------------------------------
def ensure_bucket(name, file_size_limit=None):
    """Create a private storage bucket if it doesn't already exist.

    file_size_limit is per-bucket; note the *project's* own global upload
    limit still applies on top of it, and the master deck is ~110MB -- an
    upload rejected as too large means that project-level limit needs raising
    (Supabase dashboard -> Storage -> Settings), not this argument.
    """
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        existing = {b.name if hasattr(b, "name") else b["name"] for b in client.storage.list_buckets()}
        if name in existing:
            return True, None
        try:
            options = {"public": False}
            if file_size_limit:
                options["file_size_limit"] = file_size_limit
            client.storage.create_bucket(name, options=options)
        except Exception as exc:
            # A per-bucket limit above the project's own ceiling is rejected
            # outright (413), which would otherwise leave the bucket
            # uncreated. Fall back to creating it with no explicit limit, so
            # it simply inherits whatever the project allows.
            if not file_size_limit or not _is_too_large(exc):
                raise
            client.storage.create_bucket(name, options={"public": False})
        return True, None
    except Exception as exc:
        return False, describe_error(exc)


def _is_too_large(exc):
    text = str(exc).lower()
    return "413" in text or "exceeded the maximum allowed size" in text or "too large" in text


def storage_limit_bytes(bucket=DECKS_BUCKET):
    """The largest object this bucket will accept.

    Prefers the bucket's own file_size_limit; falls back to
    DEFAULT_STORAGE_LIMIT when it doesn't state one, since the project-level
    ceiling that actually applies isn't exposed by any API.
    """
    client = get_client()
    if client is None:
        return DEFAULT_STORAGE_LIMIT
    try:
        info = client.storage.get_bucket(bucket)
    except Exception:
        return DEFAULT_STORAGE_LIMIT
    limit = getattr(info, "file_size_limit", None)
    if limit is None and isinstance(info, dict):
        limit = info.get("file_size_limit")
    return int(limit) if limit else DEFAULT_STORAGE_LIMIT


def _mib(size):
    return f"{size / 1024 / 1024:.2f} MiB"


def prepare_deck_for_upload(local_path, limit=None):
    """Optimize a deck ahead of storing it, and refuse it if it still
    won't fit. Returns (path_to_upload, stats, error).

    Every upload route goes through here, so a master deck can't reach
    storage un-optimized: the raw v1.1 master is ~95MiB against a 50MiB
    ceiling, and PowerPoint saves photographic content as PNG by default, so
    a hand-uploaded deck would routinely be twice the size it needs to be.

    The optimizer is idempotent (it leaves existing JPEGs alone), so a deck
    that has already been through here re-uploads byte-identically rather
    than losing a little quality on every round trip.

    On failure nothing is uploaded and no version row is written, so the
    currently active deck is left exactly as it was. `stats` is returned
    either way, so the caller can say how far off it was.
    """
    import optimize_deck as optimizer

    limit = storage_limit_bytes() if limit is None else limit
    source_size = Path(local_path).stat().st_size

    handle = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    handle.close()
    try:
        stats = optimizer.optimize_deck(local_path, handle.name, verbose=False)
    except Exception as exc:
        os.unlink(handle.name)
        return None, None, f"Couldn't optimize the deck before uploading it ({describe_error(exc)})"

    stats["source_size"] = source_size
    if stats["dst_size"] > limit:
        os.unlink(handle.name)
        # The optimizer is idempotent, so a deck that has already been
        # through it comes back the same size. Saying "down from 47.75 MiB"
        # about a deck that is still 47.75 MiB reads as a bug in the
        # reporting rather than as the fact it is, so the two cases are
        # worded apart.
        if source_size - stats["dst_size"] > 0.1 * 1024 * 1024:
            preamble = (f"The optimized deck is {_mib(stats['dst_size'])} "
                        f"(down from {_mib(source_size)}), still over this project's "
                        f"{_mib(limit)} storage limit.")
        else:
            preamble = (f"The deck is {_mib(stats['dst_size'])}, over this project's "
                        f"{_mib(limit)} storage limit. It was already optimized, so "
                        f"re-running the optimizer recovered nothing further at these settings.")
        return None, stats, (
            f"{preamble} Nothing was uploaded and the active master deck is unchanged. "
            f"Either re-run optimize_deck.py with a lower --quality (or a smaller --max-dim) "
            f"and upload that, or raise the limit in the Supabase dashboard under "
            f"Storage -> Settings -> Upload file size limit."
        )
    return handle.name, stats, None


def upload_deck(local_path, storage_path, notes=None, activate=True, optimize=True):
    """Upload a .pptx into the decks bucket and register it as a version.

    Returns (row, stats, error). The deck is optimized first unless
    optimize=False, and is rejected before anything is written if it still
    exceeds the storage limit -- see prepare_deck_for_upload.

    With activate=True the new row becomes the single active version -- the
    previous one is cleared first, because the partial unique index allows
    only one active row at a time.
    """
    client = get_client()
    if client is None:
        return None, None, "Supabase isn't configured"

    upload_path, stats, error = (local_path, None, None)
    if optimize:
        upload_path, stats, error = prepare_deck_for_upload(local_path)
        if error:
            return None, stats, error

    try:
        with open(upload_path, "rb") as handle:
            client.storage.from_(DECKS_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                 "upsert": "true"},
            )
        # filename is the deck's own name, not the temp file the optimized
        # copy happened to live in.
        row = {"storage_path": storage_path, "filename": Path(local_path).name,
               "notes": notes, "active": False}
        inserted = client.table("deck_versions").insert(row).execute().data[0]
    except Exception as exc:
        return None, stats, describe_error(exc)
    finally:
        if optimize and upload_path and upload_path != local_path:
            try:
                os.unlink(upload_path)
            except OSError:
                pass

    if activate:
        ok, error = activate_deck_version(inserted["id"])
        if not ok:
            return inserted, stats, error
        inserted["active"] = True
    return inserted, stats, None


def activate_deck_version(version_id):
    """Make one version the active one. Clears the previous active row first
    -- deck_versions_single_active would reject two active rows."""
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        client.table("deck_versions").update({"active": False}).eq("active", True).execute()
        client.table("deck_versions").update({"active": True}).eq("id", version_id).execute()
    except Exception as exc:
        return False, describe_error(exc)
    clear_deck_cache()
    return True, None


def upload_report_deck(local_path, storage_path, notes=None, activate=True):
    """Upload a .pptx into the report_decks bucket and register it as a
    version -- upload_deck()'s equivalent for the attribution report
    master. No optimize_deck pass: that machinery is tuned for the
    proposal master's ~95MiB of photographic PNGs, and the report
    template is a few hundred KB of small brand-asset PNGs, nowhere near
    the storage ceiling. Still refused outright if it somehow exceeds the
    bucket's own limit, same as every other upload route -- nothing
    uploaded, nothing written, on failure.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    size = Path(local_path).stat().st_size
    limit = storage_limit_bytes(REPORT_DECKS_BUCKET)
    if size > limit:
        return None, (f"{_mib(size)} exceeds this project's {_mib(limit)} storage limit "
                      f"for the report_decks bucket -- nothing uploaded.")
    try:
        with open(local_path, "rb") as handle:
            client.storage.from_(REPORT_DECKS_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                 "upsert": "true"},
            )
        row = {"storage_path": storage_path, "filename": Path(local_path).name,
               "notes": notes, "active": False}
        inserted = client.table("report_deck_versions").insert(row).execute().data[0]
    except Exception as exc:
        return None, describe_error(exc)
    if activate:
        ok, error = activate_report_deck_version(inserted["id"])
        if not ok:
            return inserted, error
        inserted["active"] = True
    return inserted, None


def activate_report_deck_version(version_id):
    """activate_deck_version()'s equivalent for report_deck_versions."""
    client = get_client()
    if client is None:
        return False, "Supabase isn't configured"
    try:
        client.table("report_deck_versions").update({"active": False}).eq("active", True).execute()
        client.table("report_deck_versions").update({"active": True}).eq("id", version_id).execute()
    except Exception as exc:
        return False, describe_error(exc)
    clear_report_deck_cache()
    return True, None


def list_report_deck_versions():
    """(rows, warning) -- every registered report deck version, newest
    first. list_deck_versions()'s equivalent."""
    client = get_client()
    if client is None:
        return [], "Supabase isn't configured"
    try:
        result = (client.table("report_deck_versions").select("*")
                  .order("uploaded_at", desc=True).execute())
        return result.data or [], None
    except Exception as exc:
        return [], f"Couldn't reach Supabase ({describe_error(exc)})"


def list_deck_versions():
    """(rows, warning) -- every registered deck version, newest first."""
    client = get_client()
    if client is None:
        return [], "Supabase isn't configured"
    try:
        result = (client.table("deck_versions").select("*")
                  .order("uploaded_at", desc=True).execute())
        return result.data or [], None
    except Exception as exc:
        return [], f"Couldn't reach Supabase ({describe_error(exc)})"
