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
    """
    row, warning = active_deck_version()
    if row is None:
        return local_fallback_path, None, f"{warning}. Using the local master deck file instead."
    try:
        return _deck_file_for_version(row["id"], row["storage_path"]), row["id"], None
    except Exception as exc:
        return (local_fallback_path, None,
                f"Couldn't download master deck version {row['id']} from Supabase "
                f"({describe_error(exc)}). Using the local master deck file instead.")


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
                 deck_version_id=None, output_filename=None):
    """Record one generated proposal. Returns (row_id, error).

    Capture only -- nothing reads this back yet. A failed log must never take
    a successfully generated deck down with it, so this reports rather than
    raises and the caller shows it as a caption, not an error.
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
    }
    try:
        result = client.table("proposals").insert(row).execute()
    except Exception as exc:
        return None, describe_error(exc)
    return (result.data or [{}])[0].get("id"), None


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
