"""
setup_supabase.py -- one-time bootstrap of the Supabase backend.

Run the DDL in `supabase_schema.sql` first (paste it into the Supabase SQL
editor -- the service key can't issue DDL over PostgREST), then run this to
create the storage buckets and load the seed data:

    python setup_supabase.py all          # everything below, in order
    python setup_supabase.py deck         # buckets + master deck as version 1

Every step is idempotent: it checks for what it would create and skips it,
so re-running is safe and reports "already present" rather than duplicating.
Reads SUPABASE_URL / SUPABASE_SERVICE_KEY through db.py, i.e. from
.streamlit/secrets.toml.
"""

import sys
from datetime import datetime, timezone
from pathlib import Path

import db

MASTER_DECK_LOCAL = Path(__file__).parent / "TEGNA_MASTER_DECK_v1_1.pptx"

# Generous per-bucket ceiling for the ~110MB master deck. The project-level
# upload limit applies on top of this and may also need raising.
DECK_SIZE_LIMIT = 500 * 1024 * 1024


def _fail(message):
    print(f"  FAILED: {message}")
    return False


def setup_deck():
    """Create the decks bucket and register the local master deck as the
    active version, unless a version is already registered."""
    print("== Stage 1: storage bucket + master deck ==")
    ok, error = db.ensure_bucket(db.DECKS_BUCKET, file_size_limit=DECK_SIZE_LIMIT)
    if not ok:
        return _fail(f"couldn't create the '{db.DECKS_BUCKET}' bucket: {error}")
    print(f"  bucket '{db.DECKS_BUCKET}' ready")

    existing, warning = db.list_deck_versions()
    if warning:
        return _fail(warning)
    if existing:
        active = [row for row in existing if row["active"]]
        print(f"  {len(existing)} deck version(s) already registered "
              f"(active: {active[0]['filename'] if active else 'none'}) -- skipping upload")
        return True

    if not MASTER_DECK_LOCAL.exists():
        return _fail(f"local master deck not found at {MASTER_DECK_LOCAL}")

    size_mb = MASTER_DECK_LOCAL.stat().st_size / (1024 * 1024)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    storage_path = f"masters/{stamp}_{MASTER_DECK_LOCAL.name}"
    print(f"  uploading {MASTER_DECK_LOCAL.name} ({size_mb:.0f}MB) -> {storage_path} ...")

    row, error = db.upload_deck(
        str(MASTER_DECK_LOCAL), storage_path,
        notes="Initial import of the checked-in master deck (v1.1, 119 slides).",
        activate=True,
    )
    if error:
        hint = ""
        if "413" in str(error) or "too large" in str(error).lower():
            hint = ("\n  (the project's global upload limit is below the deck size -- raise it in "
                    "Supabase -> Storage -> Settings -> Upload file size limit)")
        return _fail(f"{error}{hint}")
    print(f"  registered as deck version {row['id']} (active)")
    return True


STEPS = {
    "deck": setup_deck,
}


def main(argv):
    if not db.is_configured():
        print("SUPABASE_URL / SUPABASE_SERVICE_KEY not found -- check .streamlit/secrets.toml.")
        return 1

    requested = argv[1:] or ["all"]
    names = list(STEPS) if requested == ["all"] else requested
    unknown = [name for name in names if name not in STEPS]
    if unknown:
        print(f"Unknown step(s): {', '.join(unknown)}. Known: {', '.join(STEPS)}, all")
        return 2

    ok = True
    for name in names:
        ok = STEPS[name]() and ok
    print("\nDone." if ok else "\nFinished with errors -- see above.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
