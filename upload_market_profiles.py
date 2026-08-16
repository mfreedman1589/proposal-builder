"""Upload the rendered market profile slides to Supabase storage.

Run locally, not on the deployed instance -- the renders aren't in the repo
(35MiB of JPEG, and the source deck is gitignored like every other .pptx):

    python upload_market_profiles.py --list        # what would be uploaded
    python upload_market_profiles.py --pending     # only markets with no image yet
    python upload_market_profiles.py --all         # re-upload everything

Images come from the archive written when the deck was rendered, which lives
outside the repo AND outside %TEMP% -- Storage Sense reaps temp directories,
and re-rendering 206 slides through PowerPoint COM is the slow step this
whole pipeline exists to avoid repeating. Override with --source.

The mapping from file to market is by SLIDE NUMBER, taken from
market_profiles_index.json, never by sorting filenames next to sorted market
names: the deck is alphabetical by Premion's own abbreviated labels and the
canonical DMA names sort differently, so a filename-order pairing would put
a handful of markets' profiles under the wrong DMA. That is the kind of
error a client notices and nobody else does.
"""
import argparse
import sys
from pathlib import Path

import db
import market_profiles

DEFAULT_SOURCE = Path(r"C:\projects\proposal-builder-assets\market_profiles\jpeg")

# Matches the renders and render_case_study_images.DEFAULT_WIDTH. Recorded on
# each row so a later re-render at a different resolution is identifiable.
IMAGE_WIDTH = 2560


def pairs_for(source, rows, only_pending, remote):
    """[(key, path), ...] for markets whose slide exists on disk.

    A market with no slide in the source deck is skipped silently here --
    it's expected, it's named in the seed output, and the picker is what
    tells a rep about it.
    """
    pairs, missing = [], []
    for row in rows:
        if row["slide_number"] is None:
            continue
        if only_pending and remote.get(row["key"]):
            continue
        path = source / f"slide_{row['slide_number']:03d}.jpeg"
        if path.exists():
            pairs.append((row["key"], path))
        else:
            missing.append((row["key"], path.name))
    return pairs, missing


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--list", action="store_true",
                       help="report what would be uploaded, change nothing")
    group.add_argument("--pending", action="store_true",
                       help="only markets that have no image stored yet")
    group.add_argument("--all", action="store_true",
                       help="re-upload every market's slide")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    args = parser.parse_args(argv)

    if not db.is_configured():
        print("SUPABASE_URL / SUPABASE_SERVICE_KEY not found -- check .streamlit/secrets.toml.")
        return 2
    if not args.source.is_dir():
        print(f"No renders at {args.source}\n"
              "Point --source at the JPEG archive, or re-render the deck.")
        return 2

    stored, warning = db.fetch_market_profiles()
    if stored is None:
        print(f"Can't read the market_profiles table: {warning}\n"
              "Run `python setup_supabase.py market_profiles` first.")
        return 2
    remote = {r["key"]: r.get("image_path") for r in stored}

    rows = market_profiles.sort_rows(stored)
    only_pending = args.pending or args.list
    pairs, missing = pairs_for(args.source, rows, only_pending, remote)

    if missing:
        print(f"WARNING: {len(missing)} market(s) have a slide number but no "
              f"rendered file -- these will be skipped:")
        for key, name in missing[:10]:
            print(f"   {key}: expected {name}")

    already = sum(1 for v in remote.values() if v)
    print(f"{len(stored)} market(s) in the table; {already} already have an image.")

    if args.list:
        print(f"\n{len(pairs)} would be uploaded from {args.source}:")
        for key, path in pairs[:15]:
            print(f"   {key:<34} <- {path.name}")
        if len(pairs) > 15:
            print(f"   ... and {len(pairs) - 15} more")
        return 0

    if not pairs:
        print("Nothing to upload.")
        return 0

    total = sum(p.stat().st_size for _, p in pairs)
    print(f"Uploading {len(pairs)} image(s), {total / 2**20:.1f} MiB...")
    count, error = db.upload_market_profile_images(pairs, IMAGE_WIDTH)
    if error:
        print(f"  FAILED after {count} upload(s): {error}")
        return 1
    print(f"  {count} uploaded.")

    used, objects, warn = db.bucket_usage(db.MARKET_PROFILES_BUCKET)
    if not warn:
        print(f"  bucket '{db.MARKET_PROFILES_BUCKET}': "
              f"{objects} object(s), {used / 2**20:.1f} MiB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
