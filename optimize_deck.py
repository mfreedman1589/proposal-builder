"""
optimize_deck.py -- shrink a master deck's embedded images without touching
its content.

The v1.1 master is ~95MiB, of which ~81MB is PNG. Almost none of that is
resolution: nothing in the deck exceeds 2000px and half the images are under
220px. It's format -- photographic content saved as PNG, which stores it
losslessly and enormously. Re-encoding those as JPEG takes the whole deck to
~48MiB, under the 50MiB storage ceiling, with only the largest ~6% of images
touched by the resolution cap at all.

What it does, part by part:

  opaque PNG     -> JPEG (quality/progressive), part renamed to .jpeg
  raster + alpha -> PNG re-encoded with optimize=True, extension unchanged
                    (flattening these to JPEG would fill every transparent
                    pixel with black, so alpha is the hard line)
  existing JPEG  -> untouched unless it needs downscaling; see below
  .emf / .svg    -> untouched (vector; re-encoding would rasterize them)
  .fntdata       -> untouched (embedded fonts; dropping them changes how the
                    deck renders on a machine without them installed)

**This is idempotent, and deliberately so.** Re-running it on its own output
re-encodes nothing and produces a byte-identical file. That matters because
db.upload_deck runs this on every master deck upload, and a deck can make the
round trip repeatedly -- download the active master, edit a slide, upload it
again. Without the leave-existing-JPEGs-alone rule, each trip re-compressed
~85 of 255 images for a measured 0.0MB of gain, so quality would erode a
little every time for nothing.

Renaming a part means rewriting every relationship that points at it. Image
references in slide XML are part-local rIds, so only the `.rels` files carry
the actual filename -- that plus a `[Content_Types].xml` Default for jpeg is
the whole of the bookkeeping.

Nothing is ever written in place: the output is a new file, and any image
whose "optimized" form comes out larger than the original keeps its original
bytes.

    python optimize_deck.py TEGNA_MASTER_DECK_v1_1.pptx
    python optimize_deck.py in.pptx -o out.pptx --max-dim 1200 --quality 82

Verify the result with `--verify`, which re-derives the slide map from both
decks and refuses to call it a success unless every slide still resolves to
the same condition_key it did before.
"""

import argparse
import io
import re
import sys
import zipfile
from pathlib import Path

from PIL import Image

# Defaults shared by the CLI and the upload path in db.py, so a deck stored
# through the app is encoded exactly like one checked by hand. q72 is the
# setting the v1.1 master was reviewed and signed off at; 1600px is ~120dpi
# across a 13.33in slide and touches only the largest ~6% of its images.
DEFAULT_QUALITY = 72
DEFAULT_MAX_DIM = 1600

MEDIA_PREFIX = "ppt/media/"
RASTER_EXTENSIONS = {".png", ".jpg", ".jpeg"}
CONTENT_TYPES = "[Content_Types].xml"
JPEG_DEFAULT = '<Default Extension="jpeg" ContentType="image/jpeg"/>'


def _has_alpha(image):
    """True if any pixel is actually non-opaque -- not merely that the mode
    has an alpha channel. Plenty of these PNGs carry a fully-opaque alpha
    channel they never use, and those convert to JPEG safely."""
    if image.mode not in ("RGBA", "LA", "PA") and not (
        image.mode == "P" and "transparency" in image.info
    ):
        return False
    return image.convert("RGBA").getchannel("A").getextrema()[0] < 255


def optimize_image(raw, max_dim, quality, source_suffix=""):
    """(new_bytes, new_extension) for one media part, or None to keep it as
    it is. Returns None whenever re-encoding wouldn't actually save
    anything, so a well-optimized image is never degraded for nothing."""
    try:
        image = Image.open(io.BytesIO(raw))
        image.load()
    except Exception:
        return None  # not something Pillow reads -- leave it alone

    resized = False
    if max_dim and max(image.size) > max_dim:
        scale = max_dim / max(image.size)
        new_size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
        image = image.resize(new_size, Image.LANCZOS)
        resized = True

    # An image that is already a JPEG has already been through lossy
    # compression, so re-encoding it costs quality every time and buys
    # essentially nothing -- measured at 0.0MB across a whole deck, while
    # still replacing 85 of 255 parts. That matters because a deck can make
    # this round trip repeatedly (download the active master, edit a slide,
    # upload it again), and the loss would compound on every pass. The one
    # exception is an image being downscaled, where re-encoding is
    # unavoidable anyway.
    if not resized and source_suffix.lower() in (".jpg", ".jpeg"):
        return None

    buffer = io.BytesIO()
    if _has_alpha(image):
        image.convert("RGBA").save(buffer, "PNG", optimize=True)
        extension = ".png"
    else:
        image.convert("RGB").save(buffer, "JPEG", quality=quality, optimize=True, progressive=True)
        extension = ".jpeg"

    encoded = buffer.getvalue()
    if len(encoded) >= len(raw):
        return None
    return encoded, extension


def _rewrite_rels(xml_bytes, renames):
    """Point every relationship at its part's new filename.

    Matched on the full basename with a boundary, because image names are
    prefixes of each other -- a plain replace of "image33.png" would corrupt
    "image335.png" as well.
    """
    text = xml_bytes.decode("utf-8")
    for old_name, new_name in renames.items():
        text = re.sub(rf"(?<=/){re.escape(old_name)}(?=\")", new_name, text)
    return text.encode("utf-8")


def _ensure_jpeg_default(xml_bytes):
    """A part renamed to .jpeg needs a content type. The deck already has
    JPEGs so the Default is normally there -- this only covers a deck that
    happened to have none."""
    text = xml_bytes.decode("utf-8")
    if 'Extension="jpeg"' in text or "Extension='jpeg'" in text:
        return xml_bytes
    return text.replace("</Types>", JPEG_DEFAULT + "</Types>", 1).encode("utf-8")


def optimize_deck(src_path, dst_path, max_dim=DEFAULT_MAX_DIM, quality=DEFAULT_QUALITY, verbose=True):
    """Rewrite src_path into dst_path with its raster media re-encoded.
    Returns a stats dict."""
    src_zip = zipfile.ZipFile(src_path)
    entries = src_zip.infolist()

    optimized = {}   # partname -> new bytes
    renames = {}     # old basename -> new basename
    stats = {"images": 0, "converted": 0, "before": 0, "after": 0}

    for entry in entries:
        name = entry.filename
        if not name.startswith(MEDIA_PREFIX):
            continue
        if Path(name).suffix.lower() not in RASTER_EXTENSIONS:
            continue

        raw = src_zip.read(name)
        stats["images"] += 1
        stats["before"] += len(raw)

        result = optimize_image(raw, max_dim, quality, Path(name).suffix)
        if result is None:
            stats["after"] += len(raw)
            continue

        encoded, extension = result
        stats["after"] += len(encoded)
        stats["converted"] += 1

        new_name = str(Path(name).with_suffix(extension)).replace("\\", "/")
        optimized[name] = (new_name, encoded)
        if new_name != name:
            renames[Path(name).name] = Path(new_name).name

    if verbose:
        saved = stats["before"] - stats["after"]
        print(f"  {stats['images']} raster parts, {stats['converted']} re-encoded, "
              f"{stats['before'] / 1e6:.1f}MB -> {stats['after'] / 1e6:.1f}MB "
              f"(saved {saved / 1e6:.1f}MB)")

    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as out:
        for entry in entries:
            name = entry.filename
            data = src_zip.read(name)

            if name in optimized:
                name, data = optimized[name]
            elif name.endswith(".rels") and renames:
                data = _rewrite_rels(data, renames)
            elif name == CONTENT_TYPES:
                data = _ensure_jpeg_default(data)

            # Everything is deflated, including media. Storing media instead
            # looks reasonable -- JPEG and PNG are already compressed, so
            # deflate gains nothing on them -- but this deck is also full of
            # SVG and EMF, which are text and vector and compress hugely.
            # Storing those made the "optimized" file *larger* than a deck
            # PowerPoint or python-pptx had saved (43.9MiB in, 47.8MiB out),
            # which is both embarrassing and dangerous next to a 50MiB
            # ceiling. Deflate costs a little CPU on a once-per-upload
            # operation and never loses.
            out.writestr(zipfile.ZipInfo(name, date_time=entry.date_time), data,
                         compress_type=zipfile.ZIP_DEFLATED)

    src_zip.close()
    stats["src_size"] = Path(src_path).stat().st_size
    stats["dst_size"] = Path(dst_path).stat().st_size
    return stats


def verify(src_path, dst_path):
    """Confirm the optimized deck still opens, still has every slide, and
    still classifies identically. A deck that lost a slide or stopped
    resolving to the same condition_keys is worse than useless -- assembly
    would silently mis-select from it."""
    import slide_map

    print("\nVerifying...")
    try:
        before = slide_map.build_slide_map(src_path)
        after = slide_map.build_slide_map(dst_path)
    except Exception as exc:
        print(f"  FAILED: the optimized deck wouldn't open ({exc})")
        return False

    ok = True
    if len(before) != len(after):
        print(f"  FAILED: slide count changed, {len(before)} -> {len(after)}")
        ok = False
    else:
        print(f"  slide count unchanged ({len(after)})")

    changed = [n for n in before if before.get(n) != after.get(n)]
    if changed:
        print(f"  FAILED: {len(changed)} slide(s) classify differently now: {changed[:10]}")
        ok = False
    else:
        print("  every slide resolves to the same condition_key as before")

    unresolved = [n for n, key in after.items() if key is None]
    if unresolved:
        print(f"  FAILED: {len(unresolved)} slide(s) don't resolve at all: {unresolved}")
        ok = False
    else:
        print(f"  slide map resolves 100% ({len(after)}/{len(after)})")

    return ok


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("source", help="master deck .pptx to optimize")
    parser.add_argument("-o", "--output", help="output path (default: <source>_optimized.pptx)")
    parser.add_argument("--max-dim", type=int, default=DEFAULT_MAX_DIM,
                        help=f"downscale any image whose longest side exceeds this (default {DEFAULT_MAX_DIM}, "
                             "~120dpi across a 13.33in slide). 0 disables downscaling.")
    parser.add_argument("--quality", type=int, default=DEFAULT_QUALITY,
                        help=f"JPEG quality for opaque images (default {DEFAULT_QUALITY})")
    parser.add_argument("--verify", action="store_true",
                        help="re-derive the slide map from both decks and compare")
    args = parser.parse_args(argv)

    source = Path(args.source)
    if not source.exists():
        print(f"No such deck: {source}")
        return 1
    output = Path(args.output) if args.output else source.with_name(source.stem + "_optimized.pptx")

    print(f"Optimizing {source.name} (max-dim {args.max_dim or 'none'}, JPEG q{args.quality})")
    stats = optimize_deck(str(source), str(output), args.max_dim, args.quality)
    print(f"  {source.name}: {stats['src_size'] / 1e6:.1f}MB")
    print(f"  {output.name}: {stats['dst_size'] / 1e6:.1f}MB "
          f"({100 * (1 - stats['dst_size'] / stats['src_size']):.0f}% smaller)")

    if args.verify and not verify(str(source), str(output)):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
