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
import market_profiles

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

    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    storage_path = f"masters/{stamp}_{MASTER_DECK_LOCAL.name}"
    print(f"  optimizing and uploading {MASTER_DECK_LOCAL.name} "
          f"({MASTER_DECK_LOCAL.stat().st_size / 1024 / 1024:.2f} MiB) -> {storage_path} ...")

    row, stats, error = db.upload_deck(
        str(MASTER_DECK_LOCAL), storage_path,
        notes="Initial import of the checked-in master deck (v1.1, 119 slides).",
        activate=True,
    )
    if stats:
        print(f"  optimized {stats['source_size'] / 1024 / 1024:.2f} MiB -> "
              f"{stats['dst_size'] / 1024 / 1024:.2f} MiB "
              f"({stats['converted']} of {stats['images']} images re-encoded)")
    if error:
        return _fail(error)
    print(f"  registered as deck version {row['id']} (active)")
    return True


def setup_products():
    """Seed the products table from app.py's fallback rate card, Live Sports
    packages included.

    Upserted on `key`, so re-running refreshes CPMs from the built-in copies
    rather than duplicating rows -- which also means running this *after*
    editing rates in Supabase would overwrite those edits with the older
    hardcoded values. It's a bootstrap, not a sync.
    """
    print("== Stage 2: products / rate card ==")
    import app  # module body is import-safe; only main() is behind the guard

    rows = []
    for order, (key, spec) in enumerate(app.FALLBACK_PRODUCTS.items()):
        rows.append({
            "key": key,
            "name": spec["label"],
            "line_type": "rate",
            "default_cpm": spec["default_cpm"],
            "targeting_copy": spec.get("targeting_copy"),
            "display_group": spec["line_type"],
            "active": True,
            "sort_order": order,
        })

    # Sports are products too -- one row per package, keyed "sport:<key>",
    # carrying the ratecard rate that a generic product default would get
    # wrong by tens of dollars per thousand.
    for order, (sport_key, cpm) in enumerate(app.FALLBACK_SPORT_CPM.items(), start=100):
        rows.append({
            "key": f"{app.SPORT_PRODUCT_PREFIX}{sport_key}",
            "name": app.sport_product_label(sport_key),
            "line_type": "rate",
            "default_cpm": cpm,
            "targeting_copy": app.LIVE_SPORTS_TARGETING,
            "display_group": "sports",
            "active": True,
            "sort_order": order,
        })

    count, error = db.upsert_products(rows)
    if error:
        return _fail(f"couldn't seed products: {error}")
    print(f"  seeded {count} products "
          f"({len(app.FALLBACK_PRODUCTS)} media products + {len(app.FALLBACK_SPORT_CPM)} sports packages)")
    return True


def setup_audiences():
    """Seed `audiences` from the brochure-PDF catalog and `audience_usage`
    from the year-to-date delivery CSV."""
    print("== Stage 3: audience catalog ==")
    import pandas as pd

    from audience_catalog import load_local_catalog

    catalog = load_local_catalog()
    rows = [{"segment": row.segment, "category": row.category, "subcategory": row.subcategory,
             "rfp_selectable": bool(row.rfp_selectable), "times_used": int(row.times_used),
             "active": True}
            for row in catalog.itertuples()]
    count, error = db.upsert_audiences(rows)
    if error:
        return _fail(f"couldn't seed audiences ({count} written before failing): {error}")
    print(f"  seeded {count} segments "
          f"({sum(r['rfp_selectable'] for r in rows)} RFP-selectable)")

    usage_csv = Path(__file__).parent / "audience_usage_ytd.csv"
    if not usage_csv.exists():
        print(f"  skipping audience_usage -- {usage_csv.name} not found")
        return True

    # A booked audience is often a *combination* of catalog segments rather
    # than one catalog name, which is exactly what makes it custom: the same
    # rule the form applies to the avails table (anything not an
    # RFP-selectable catalog segment counts as the campaign's one custom
    # audience).
    rfp_segments = set(catalog.loc[catalog["rfp_selectable"], "segment"])
    usage = pd.read_csv(usage_csv)
    usage_rows = [{"segment_string": str(row.segment),
                   "delivered_impressions": int(row.impressions),
                   "is_custom": str(row.segment) not in rfp_segments}
                  for row in usage.itertuples()]
    count, error = db.replace_audience_usage(usage_rows)
    if error:
        return _fail(f"couldn't seed audience_usage ({count} written before failing): {error}")
    print(f"  seeded {count} usage rows "
          f"({sum(r['is_custom'] for r in usage_rows)} custom)")
    return True


def setup_case_studies():
    """Create the case study bucket. Nothing is seeded -- the vault fills up
    through the app's own "Add case study" page."""
    print("== Case study vault ==")
    ok, error = db.ensure_bucket(db.CASE_STUDIES_BUCKET, file_size_limit=DECK_SIZE_LIMIT)
    if not ok:
        return _fail(f"couldn't create the '{db.CASE_STUDIES_BUCKET}' bucket: {error}")
    rows, warning = db.fetch_case_studies()
    if warning:
        return _fail(warning)
    print(f"  bucket '{db.CASE_STUDIES_BUCKET}' ready; {len(rows)} case study(ies) in the vault")
    return True


def setup_proposal_files():
    """Create the bucket that holds hand-edited final decks attached to
    history rows. Nothing is seeded -- it fills up through the History page's
    "Attach final file" action.

    The columns these files are recorded against (file_storage_path,
    file_attached_at, file_note) are DDL and live in supabase_schema.sql --
    run that in the SQL editor first, since the service key can't issue DDL
    over PostgREST.
    """
    print("== Proposal files ==")
    ok, error = db.ensure_bucket(db.PROPOSAL_FILES_BUCKET, file_size_limit=DECK_SIZE_LIMIT)
    if not ok:
        return _fail(f"couldn't create the '{db.PROPOSAL_FILES_BUCKET}' bucket: {error}")
    total, count, warning = db.bucket_usage(db.PROPOSAL_FILES_BUCKET)
    if warning:
        return _fail(warning)
    print(f"  bucket '{db.PROPOSAL_FILES_BUCKET}' ready; "
          f"{count} attached file(s), {total / 1024 / 1024:.1f} MiB used")
    return True


def setup_market_profiles():
    """Create the market profile bucket and seed one row per selectable market.

    The seed is derived from market_profiles_index.json, which is in the repo,
    so this step needs no local renders and can be re-run anywhere. It writes
    identity and ordering ONLY -- never image_path, image_width or stats --
    so re-running after the images are uploaded can't null them. Uploading
    the images is a separate, local step (upload_market_profiles.py), because
    the renders are 35MiB and aren't in the repo.

    All 210 Nielsen DMAs get a row, including the five the source deck has no
    slide for (Honolulu, Palm Springs, Anchorage, Fairbanks, Juneau). They're
    real markets and stay selectable; the picker says the profile is missing
    rather than the market silently disappearing.
    """
    print("== Market viewer profiles ==")
    ok, error = db.ensure_bucket(db.MARKET_PROFILES_BUCKET,
                                 file_size_limit=DECK_SIZE_LIMIT)
    if not ok:
        return _fail(f"couldn't create the '{db.MARKET_PROFILES_BUCKET}' bucket: {error}")
    print(f"  bucket '{db.MARKET_PROFILES_BUCKET}' ready")

    rows = market_profiles.build_rows()
    seed = [{
        "key": row["key"],
        "label": row["label"],
        "dma": row["dma"],
        "kind": row["kind"],
        "slide_number": row["slide_number"],
        "active": True,
    } for row in rows]

    count, error = db.upsert_market_profiles(seed)
    if error:
        return _fail(f"couldn't seed market profiles: {error}")

    without = market_profiles.markets_without_a_slide(rows)
    print(f"  {count} market(s) seeded; {len(rows) - len(without)} with a profile "
          f"slide, {len(without)} without ({', '.join(r['label'] for r in without)})")
    return True


STEPS = {
    "deck": setup_deck,
    "products": setup_products,
    "audiences": setup_audiences,
    "case_studies": setup_case_studies,
    "proposal_files": setup_proposal_files,
    "market_profiles": setup_market_profiles,
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
