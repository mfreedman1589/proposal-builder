"""Render vault slides to images and store them alongside their source deck.

    python render_slide_vault_images.py --pending     # only ones without an image
    python render_slide_vault_images.py --all         # re-render everything
    python render_slide_vault_images.py --list        # what's covered, what isn't
    python render_slide_vault_images.py --pending --dry-run

Or add a second line to render_pending.bat, which already does the case-study
--pending run and waits so you can read the summary.

Same reasoning as render_case_study_images.py, one level down: rendering
needs PowerPoint and the Proxima Nova brand font, which Streamlit Cloud has
neither of, so this is a local, human-run step. A rendered slide has no
relationships, theme, layout, placeholders, autofit or embedded parts for
cross-deck copying to get wrong.

The one real difference from the case-study script: several vault rows can
share one uploaded deck (one row per picked slide, not per upload -- see
db.upload_slide_vault), so targets are grouped by storage_path and each
source deck is opened and rendered through PowerPoint ONCE, however many
rows came from it -- only the picked slide_index's own image is kept per row.

Safe to run unattended (see --quiet and --skip-if-busy, and CLAUDE.md for the
Task Scheduler recipe).
"""

import argparse
import datetime
import sys
from collections import defaultdict
from pathlib import Path

import db
import deck_render

# Same three exit codes as render_case_study_images.py.
EXIT_OK, EXIT_FAILED, EXIT_UNAVAILABLE = 0, 1, 2

DEFAULT_WIDTH = 2560
JPEG_QUALITY = 88

_log_file = None


def log(message=""):
    print(message)
    if _log_file is not None:
        _log_file.write(message + "\n")
        _log_file.flush()


def log_path():
    return deck_render.render_root() / "render_slide_vault_images.log"


def to_jpeg(png_path):
    from PIL import Image

    jpeg_path = Path(str(png_path).rsplit(".", 1)[0] + ".jpg")
    with Image.open(png_path) as image:
        image.convert("RGB").save(jpeg_path, "JPEG", quality=JPEG_QUALITY,
                                  optimize=True, progressive=True)
    png_path = Path(png_path)
    if png_path.exists():
        png_path.unlink()
    return jpeg_path


def render_group(storage_path, rows, width, dry_run=False):
    """Render every slide of ONE source deck once, then pick out each row's
    own slide_index. Returns [(row, count_or_0, error_or_None), ...].
    """
    sample = rows[0]
    try:
        source = db.slide_vault_file(storage_path)
    except Exception as exc:                                     # noqa: BLE001
        return [(row, 0, f"couldn't fetch the .pptx ({exc})") for row in rows]

    out = deck_render.render_root() / "slide_vault_images" / storage_path.replace("/", "_")
    out.mkdir(parents=True, exist_ok=True)
    try:
        images = [to_jpeg(path) for path in
                  deck_render.render_deck(source, str(out), width=width)]
    except deck_render.RenderUnavailable as exc:
        return [(row, 0, str(exc)) for row in rows]
    except Exception as exc:                                     # noqa: BLE001
        return [(row, 0, f"{type(exc).__name__}: {exc}") for row in rows]

    results = []
    for row in rows:
        index = row["slide_index"]
        if index < 0 or index >= len(images):
            results.append((row, 0, f"slide index {index} out of range "
                                    f"({len(images)} slide(s) in the deck)"))
            continue
        image = images[index]
        if dry_run:
            log(f"      would upload slide {index + 1} ({image.stat().st_size / 1048576:.2f} MiB)")
            results.append((row, 1, None))
            continue
        ok, error = db.upload_slide_vault_image(row["id"], str(image), width)
        results.append((row, 1 if ok else 0, error))
    return results


def show_coverage(rows):
    covered = [r for r in rows if db.slide_vault_render_path(r) == "image"]
    log(f"{len(covered)} of {len(rows)} vault slide(s) have a rendered image.")
    log("")
    for row in rows:
        route = db.slide_vault_render_path(row)
        mark = "image " if route == "image" else "COPY  "
        title = row.get("title") or f"{row['filename']} -- slide {row['slide_index'] + 1}"
        log(f"  {mark}  {title[:52]:52}"
            f"{'' if route == 'image' else '   <-- needs generating'}")


def main(argv=None):
    global _log_file

    parser = argparse.ArgumentParser(description="Render slide-vault slides to images.")
    parser.add_argument("--all", action="store_true",
                        help="re-render every vault slide, replacing existing images")
    parser.add_argument("--pending", action="store_true",
                        help="only vault slides that have no image yet")
    parser.add_argument("--list", action="store_true", help="show coverage and exit")
    parser.add_argument("--dry-run", action="store_true",
                        help="render and measure, upload nothing")
    parser.add_argument("--width", type=int, default=DEFAULT_WIDTH)
    parser.add_argument("--quiet", action="store_true",
                        help="for scheduled runs: say nothing when there's nothing to do")
    parser.add_argument("--skip-if-busy", action="store_true",
                        help="for scheduled runs: do nothing if PowerPoint is already open")
    args = parser.parse_args(argv)

    path = log_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    _log_file = open(path, "a", encoding="utf-8")
    log("")
    log(f"=== {datetime.datetime.now():%Y-%m-%d %H:%M:%S} "
        f"{'(scheduled)' if args.quiet else ''} ===")

    try:
        if args.skip_if_busy and deck_render.powerpoint_is_running():
            log("PowerPoint is already open -- skipping this run, will try again next time.")
            return EXIT_OK

        rows, warning = db.fetch_slide_vault(active_only=False)
        if warning:
            log(f"FAILED: {warning}")
            return EXIT_UNAVAILABLE
        if not rows:
            log("No slides in the vault.")
            return EXIT_OK

        if args.list or not (args.all or args.pending):
            show_coverage(rows)
            if not (args.all or args.pending):
                log("")
                log("Nothing rendered. Pass --pending (or --all) to generate.")
            return EXIT_OK

        targets = rows if args.all else [r for r in rows
                                         if db.slide_vault_render_path(r) == "copy"]
        if not targets:
            log("Every vault slide already has a rendered image. Nothing to do.")
            return EXIT_OK

        if not deck_render.renderer_available():
            log("FAILED: rendering needs Windows with PowerPoint and pywin32.")
            return EXIT_UNAVAILABLE

        groups = defaultdict(list)
        for row in targets:
            groups[row["storage_path"]].append(row)

        log(f"Rendering {len(targets)} vault slide(s) from {len(groups)} source deck(s) "
            f"at {args.width}px{' (dry run)' if args.dry_run else ''}...")
        log("")
        failures, stored = [], 0
        for index, (storage_path, rows_in_group) in enumerate(groups.items(), start=1):
            name = Path(storage_path).name[:52]
            log(f"  [{index}/{len(groups)}] {name} ({len(rows_in_group)} slide(s))")
            for row, count, error in render_group(storage_path, rows_in_group, args.width, args.dry_run):
                title = (row.get("title") or f"slide {row['slide_index'] + 1}")[:52]
                if error:
                    log(f"      FAILED [{title}]: {error}")
                    failures.append(title)
                else:
                    stored += count
                    log(f"      [{title}] {'rendered' if args.dry_run else 'stored'}")

        log("")
        log(f"SUMMARY: {len(targets) - len(failures)} of {len(targets)} vault slide(s) done, "
            f"{stored} image(s) {'measured' if args.dry_run else 'uploaded'}.")
        if failures:
            log(f"FAILED ({len(failures)}): {', '.join(failures)}")
            return EXIT_FAILED
        remaining, _ = db.fetch_slide_vault(active_only=False)
        left = [r for r in (remaining or []) if db.slide_vault_render_path(r) == "copy"]
        log(f"{len(left)} vault slide(s) still pending.")
        return EXIT_OK
    finally:
        log(f"--- log: {path}")
        if _log_file is not None:
            _log_file.close()
            _log_file = None


if __name__ == "__main__":
    sys.exit(main())
