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
# Hand-edited final decks attached to a history row. The exception to
# storing-the-recipe: everything else in this app is regenerable from
# form_json, these are not, which is exactly why they're worth keeping --
# and why the History page surfaces this bucket's total usage.
PROPOSAL_FILES_BUCKET = "proposal_files"

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
_CASE_STUDY_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_case_studies"
_PROPOSAL_FILE_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_proposal_files"

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
                 parent_proposal_id=None, revision_label=None):
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
    storage_path = row.get("file_storage_path")
    try:
        client.table("proposals").delete().eq("id", proposal_id).execute()
    except Exception as exc:
        return False, describe_error(exc)
    if storage_path:
        # The row is already gone; a failed object delete leaves an orphan
        # blob, which costs quota but breaks nothing, so it's reported rather
        # than raised.
        try:
            client.storage.from_(PROPOSAL_FILES_BUCKET).remove([storage_path])
        except Exception as exc:
            return True, (f"The history row was deleted, but its attached file couldn't be "
                          f"removed from storage ({describe_error(exc)}). It's now orphaned.")
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
        for entry in client.storage.from_(bucket).list() or []:
            name = entry.get("name")
            if not name:
                continue
            meta = entry.get("metadata") or {}
            if meta.get("size") is not None:
                total += int(meta["size"])
                count += 1
                continue
            for child in client.storage.from_(bucket).list(name) or []:
                child_meta = child.get("metadata") or {}
                if child_meta.get("size") is not None:
                    total += int(child_meta["size"])
                    count += 1
    except Exception as exc:
        return 0, 0, f"Couldn't measure {bucket} usage ({describe_error(exc)})"
    return total, count, None


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
