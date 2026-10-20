"""Derive a ranked audience-usage list, write the reviewable tag CSV, and
optionally upload it as a new INACTIVE usage version.

    python ingest_ranked_usage.py "IMPRESSIONS BY DATA SEGMENT 10.12.26.xlsx"
    python ingest_ranked_usage.py "IMPRESSIONS BY DATA SEGMENT 10.12.26.xlsx" --upload --notes "..."

Same derivation the admin page's "Update audience usage" runs
(segment_catalog.derive_from_workbook). The upload never activates -- the
nightly merge does that (CLAUDE.md, "The merge routine"). The CSV
(`audience_segment_tags.csv`) is the review artifact: one row per segment
with its status, role, tags, score and top pairing. Corrections go in the
admin page's "Review tags & roles" tab, not in this file -- it's
regenerated on every ingest.
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import segment_catalog

REVIEW_CSV = Path(__file__).parent / "audience_segment_tags.csv"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("workbook")
    parser.add_argument("--upload", action="store_true",
                        help="register it as a new, inactive usage version in Supabase")
    parser.add_argument("--replace-version", type=int, default=None,
                        help="re-store the derived catalog of this existing, INACTIVE version "
                             "(after a tagging-rule change) instead of uploading a new one")
    parser.add_argument("--notes", default=None)
    parser.add_argument("--json", default=None, help="also write the derived catalog here")
    args = parser.parse_args(argv)

    import db  # lazy -- deriving and writing the CSV needs no backend

    existing, warning = db.fetch_audiences()
    if existing is None:
        print(f"Note: {warning}; deriving without the current catalog's segment names.")
    names = [row["segment"] for row in existing or []]
    path = Path(args.workbook)
    derived = segment_catalog.derive_from_workbook(str(path), existing_segments=names,
                                                   source_filename=path.name)
    segment_catalog.write_review_csv(derived, REVIEW_CSV)
    print(json.dumps(derived["summary"], indent=1))
    print(f"Review CSV: {REVIEW_CSV}")
    if args.json:
        Path(args.json).write_text(json.dumps(derived), encoding="utf-8")

    if args.replace_version is not None:
        ok, error = db.replace_segment_catalog(args.replace_version, derived)
        if not ok:
            print(f"Replace failed: {error}")
            return 1
        print(f"Re-stored the derived catalog for usage version {args.replace_version} (still inactive).")
    elif args.upload:
        stamp = datetime.now().strftime("%Y%m%dT%H%M%SZ")
        row, error = db.upload_ranked_usage_workbook(
            str(path), f"usage/{stamp}_{path.name}", derived, notes=args.notes)
        if error:
            print(f"Upload failed: {error}")
            return 1
        print(f"Uploaded as usage version {row['id']} (inactive -- the nightly merge activates it).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
