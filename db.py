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

import os
import tempfile
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

# PostgREST calls are small and must fail fast, so a broken URL surfaces as a
# warning in a second or two rather than hanging the form. Storage moves the
# ~110MB master deck, so it gets a much longer leash.
POSTGREST_TIMEOUT = 15
STORAGE_TIMEOUT = 600

_DECK_CACHE_DIR = Path(tempfile.gettempdir()) / "premion_proposal_builder_decks"

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
        if name not in existing:
            options = {"public": False}
            if file_size_limit:
                options["file_size_limit"] = file_size_limit
            client.storage.create_bucket(name, options=options)
        return True, None
    except Exception as exc:
        return False, describe_error(exc)


def upload_deck(local_path, storage_path, notes=None, activate=True):
    """Upload a .pptx into the decks bucket and register it as a version.

    Returns (row, error). With activate=True the new row becomes the single
    active version -- the previous one is cleared first, because the partial
    unique index allows only one active row at a time.
    """
    client = get_client()
    if client is None:
        return None, "Supabase isn't configured"
    try:
        with open(local_path, "rb") as handle:
            client.storage.from_(DECKS_BUCKET).upload(
                storage_path, handle.read(),
                {"content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                 "upsert": "true"},
            )
        row = {"storage_path": storage_path, "filename": Path(local_path).name,
               "notes": notes, "active": False}
        inserted = client.table("deck_versions").insert(row).execute().data[0]
    except Exception as exc:
        return None, describe_error(exc)

    if activate:
        ok, error = activate_deck_version(inserted["id"])
        if not ok:
            return inserted, error
        inserted["active"] = True
    return inserted, None


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
