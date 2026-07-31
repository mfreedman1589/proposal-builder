"""
bulk_load_case_studies.py -- load a folder of case study decks into the vault.

Runs every .pptx through exactly the code path the "Add case study" page uses
-- extract slide text, have Claude propose title/verticals/products/summary,
validate the tags against the real lists, optimize, upload, insert -- so a
bulk load and a hand upload can't diverge.

Already-loaded decks are skipped by filename, so re-running only picks up
what's new.

    python bulk_load_case_studies.py                       # load case_studies_source/
    python bulk_load_case_studies.py --dir some/folder
    python bulk_load_case_studies.py --dry-run             # tag only, store nothing
    python bulk_load_case_studies.py --list                # print the vault and exit
"""

import argparse
import sys
from pathlib import Path

import app
import db

SOURCE_DIR = Path(__file__).parent / "case_studies_source"
ADDED_BY = "bulk load"


def vault_by_filename():
    rows, warning = db.fetch_case_studies(active_only=False)
    if warning:
        return None, warning
    return {row["filename"]: row for row in rows}, None


def print_vault():
    rows, warning = db.fetch_case_studies(active_only=False)
    if warning:
        print(f"  couldn't read the vault: {warning}")
        return
    print(f"\n{'=' * 100}\nVAULT -- {len(rows)} case studies\n{'=' * 100}")
    for row in sorted(rows, key=lambda r: (r.get("verticals") or ["~"])[0]):
        state = "" if row.get("active", True) else "  [DEACTIVATED]"
        print(f"\n  {row['title']}{state}")
        print(f"    verticals : {', '.join(row.get('verticals') or []) or '(none)'}")
        print(f"    products  : {', '.join(row.get('products') or []) or '(none)'}")
        print(f"    summary   : {(row.get('summary') or '')[:110]}")
        print(f"    file      : {row['filename']}  ·  added by {row.get('added_by') or '?'}")


def load_one(path, dry_run=False):
    """Tag and store one deck. Returns (row_or_suggestion, error)."""
    slide_texts = app.extract_case_study_text(str(path))
    suggestion, error = app.call_claude_case_study_tags(slide_texts, path.name)
    if error:
        return None, error

    vertical_keys = [v for v in app.VERTICALS.values() if v != "none"]
    verticals = app._valid_tags(suggestion.get("verticals"), vertical_keys)
    products = app._valid_tags(suggestion.get("products"), app.CASE_STUDY_PRODUCT_TAGS)
    title = (suggestion.get("title") or path.stem).strip()
    summary = (suggestion.get("summary") or "").strip()

    if dry_run:
        return {"title": title, "verticals": verticals, "products": products,
                "summary": summary, "raw": suggestion}, None

    row, stats, error = db.upload_case_study(
        str(path), path.name, title, verticals, products, summary, added_by=ADDED_BY)
    if error:
        return None, error
    row["_stats"] = stats
    return row, None


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--dir", default=str(SOURCE_DIR))
    parser.add_argument("--dry-run", action="store_true",
                        help="tag every deck and print the result, store nothing")
    parser.add_argument("--list", action="store_true", help="print the vault and exit")
    args = parser.parse_args(argv)

    if not db.is_configured():
        print("SUPABASE_URL / SUPABASE_SERVICE_KEY not found.")
        return 1

    if args.list:
        print_vault()
        return 0

    source = Path(args.dir)
    if not source.exists():
        print(f"No such folder: {source}")
        return 1

    existing, warning = vault_by_filename()
    if warning:
        print(f"Couldn't read the vault: {warning}")
        return 1

    decks = sorted(p for p in source.rglob("*.pptx") if not p.name.startswith("~$"))
    print(f"Found {len(decks)} .pptx under {source.name}; "
          f"{len(existing)} already in the vault.\n")

    loaded, skipped, failed = [], [], []
    for path in decks:
        if path.name in existing:
            skipped.append(path.name)
            print(f"  skip   {path.name}  (already in the vault)")
            continue

        print(f"  load   {path.name} ...", flush=True)
        result, error = load_one(path, dry_run=args.dry_run)
        if error:
            failed.append((path.name, error))
            print(f"         FAILED: {error}")
            continue

        loaded.append((path.name, result))
        stats = result.get("_stats") if not args.dry_run else None
        size = (f"  [{stats['source_size'] / 1024 / 1024:.1f} -> "
                f"{stats['dst_size'] / 1024 / 1024:.1f} MiB]" if stats else "")
        print(f"         {result['title']}")
        print(f"         verticals={result['verticals']} products={result['products']}{size}")

    print(f"\n{'-' * 100}")
    print(f"loaded {len(loaded)}, skipped {len(skipped)}, failed {len(failed)}")
    for name, error in failed:
        print(f"  FAILED {name}: {error}")

    if not args.dry_run:
        print_vault()
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
