"""Render case study slides to images and store them alongside the .pptx.

    python render_case_study_images.py --pending     # only ones without images
    python render_case_study_images.py --all         # re-render everything
    python render_case_study_images.py --list        # what's covered, what isn't
    python render_case_study_images.py --pending --dry-run

Or double-click `render_pending.bat`, which does the --pending run and waits
so you can read the summary.

Why this is a local script and not part of the app: rendering needs
PowerPoint and the Proxima Nova brand font. Streamlit Cloud has neither, so
rendering there would substitute fonts and produce images that aren't what
the case study looks like. The app therefore *consumes* images and never
makes them; this is the step a human runs.

Why images at all: copying a slide's XML between decks has produced five
distinct corruption bugs in this project, and it cannot reproduce
PowerPoint's live autofit -- flattened text renders a line taller than its
box and clips, on a slide a client sees. A rendered slide has no
relationships, theme, layout, placeholders, autofit or embedded parts, so
none of that is possible. It is also *smaller*: copying drags in the source
deck's full-resolution photography, where a rendered slide is one flat JPEG
(the two furniture case studies cost 7.83 MiB copied and 1.44 MiB rendered).

The tradeoff, which is permanent: an image slide is not selectable,
searchable or editable text. Fixing a typo means fixing the source deck and
re-running this.

Safe to run unattended (see --quiet and --skip-if-busy, and CLAUDE.md for
the Task Scheduler recipe): it renders only what's missing, refuses to touch
a PowerPoint somebody is already using, never closes an instance it didn't
start, exits non-zero on failure, and appends everything to a log.
"""

import argparse
import datetime
import sys
from pathlib import Path

import db
import deck_render

# Exit codes, so a scheduled run can be judged without reading the log:
#   0  rendered something, or there was nothing to do
#   1  at least one case study failed
#   2  couldn't run at all (no PowerPoint, no Supabase)
EXIT_OK, EXIT_FAILED, EXIT_UNAVAILABLE = 0, 1, 2

# 2560px across a 13.33in slide is ~192dpi -- enough to hold up zoomed in and
# on a printed leave-behind. 1920 (~144dpi) was measured against it and saves
# only 0.47 MiB on a deck already 29 MiB from the master, which is not a
# trade worth making for visibly softer type at full screen.
DEFAULT_WIDTH = 2560

# JPEG, not PNG. These slides are photographic and opaque; the same reasoning
# as optimize_deck.py, and the difference is several MiB per case study.
JPEG_QUALITY = 88

_log_file = None


def log(message=""):
    """Print, and append to the log file when one is open.

    An unattended run leaves nothing on screen, so anything worth knowing
    afterwards has to be on disk -- including the runs that did nothing,
    since "did it even fire?" is the first question asked of a scheduled job.
    """
    print(message)
    if _log_file is not None:
        _log_file.write(message + "\n")
        _log_file.flush()


def log_path():
    return deck_render.render_root() / "render_case_study_images.log"


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


def render_one(case_study, width, dry_run=False):
    """Render one case study's slides and upload them. Returns (count, error)."""
    try:
        source = db.case_study_file(case_study["id"], case_study["storage_path"])
    except Exception as exc:                                     # noqa: BLE001
        return 0, f"couldn't fetch the .pptx ({exc})"

    out = deck_render.render_root() / "case_study_images" / str(case_study["id"])
    out.mkdir(parents=True, exist_ok=True)
    try:
        images = [to_jpeg(path) for path in
                  deck_render.render_deck(source, str(out), width=width)]
    except deck_render.RenderUnavailable as exc:
        return 0, str(exc)
    except Exception as exc:                                     # noqa: BLE001
        # An unattended run must not die on one bad deck; the others are
        # still worth rendering and the failure is named in the log.
        return 0, f"{type(exc).__name__}: {exc}"
    if not images:
        return 0, "PowerPoint rendered no slides"
    if dry_run:
        total = sum(p.stat().st_size for p in images)
        log(f"      would upload {len(images)} image(s), {total / 1048576:.2f} MiB")
        return len(images), None
    return db.upload_case_study_images(case_study["id"], images, width)


def show_coverage(rows):
    covered = [r for r in rows if db.case_study_render_path(r) == "images"]
    log(f"{len(covered)} of {len(rows)} case studies have rendered slides.")
    log("")
    for row in rows:
        route = db.case_study_render_path(row)
        mark = "images" if route == "images" else "COPY  "
        count = len(row.get("slide_images") or [])
        log(f"  {mark}  {(row.get('title') or row['filename'])[:52]:52} "
            f"{count or '-':>3} slide(s)"
            f"{'' if route == 'images' else '   <-- needs generating'}")


def main(argv=None):
    global _log_file

    parser = argparse.ArgumentParser(description="Render case study slides to images.")
    parser.add_argument("--all", action="store_true",
                        help="re-render every case study, replacing existing images")
    parser.add_argument("--pending", action="store_true",
                        help="only case studies that have no images yet")
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
        # Asked before anything else, and before touching COM: automating a
        # PowerPoint somebody is using steals focus, and a modal dialog in
        # that window blocks every call this would make.
        if args.skip_if_busy and deck_render.powerpoint_is_running():
            log("PowerPoint is already open -- skipping this run, will try again next time.")
            return EXIT_OK

        rows, warning = db.fetch_case_studies(active_only=False)
        if warning:
            log(f"FAILED: {warning}")
            return EXIT_UNAVAILABLE
        if not rows:
            log("No case studies in the vault.")
            return EXIT_OK

        if args.list or not (args.all or args.pending):
            show_coverage(rows)
            if not (args.all or args.pending):
                log("")
                log("Nothing rendered. Pass --pending (or --all) to generate.")
            return EXIT_OK

        targets = rows if args.all else [r for r in rows
                                         if db.case_study_render_path(r) == "copy"]
        if not targets:
            # The quiet path exists for exactly this: a scheduled job that
            # finds nothing to do should leave no trace but a log line.
            log("Every case study already has rendered slides. Nothing to do.")
            return EXIT_OK

        if not deck_render.renderer_available():
            log("FAILED: rendering needs Windows with PowerPoint and pywin32.")
            return EXIT_UNAVAILABLE

        log(f"Rendering {len(targets)} case study(ies) at {args.width}px"
            f"{' (dry run)' if args.dry_run else ''}...")
        log("")
        failures, stored = [], 0
        for index, row in enumerate(targets, start=1):
            name = (row.get("title") or row.get("filename") or "")[:52]
            log(f"  [{index}/{len(targets)}] {name}")
            count, error = render_one(row, args.width, args.dry_run)
            if error:
                log(f"      FAILED: {error}")
                failures.append(name)
            else:
                stored += count
                log(f"      {count} slide(s) {'rendered' if args.dry_run else 'stored'}")

        log("")
        log(f"SUMMARY: {len(targets) - len(failures)} of {len(targets)} case "
            f"stud{'y' if len(targets) == 1 else 'ies'} done, {stored} slide image(s) "
            f"{'measured' if args.dry_run else 'uploaded'}.")
        if failures:
            log(f"FAILED ({len(failures)}): {', '.join(failures)}")
            return EXIT_FAILED
        remaining, _ = db.fetch_case_studies(active_only=False)
        left = [r for r in (remaining or []) if db.case_study_render_path(r) == "copy"]
        log(f"{len(left)} case stud{'y' if len(left) == 1 else 'ies'} still pending.")
        return EXIT_OK
    finally:
        log(f"--- log: {path}")
        if _log_file is not None:
            _log_file.close()
            _log_file = None


if __name__ == "__main__":
    sys.exit(main())
