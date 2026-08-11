"""text_metrics.py -- measure how many lines a string wraps to.

Replaces an average-character-width estimate. That estimate had to err high
(guess low and a table renders on top of the block beneath it), so it
reserved height for two lines where PowerPoint drew one, and the font search
shrank the type further than it needed to. Measuring the actual glyphs
removes the guess in both directions.

Widths come from PIL against the real TTF. Loading a font at `size=points`
makes PIL's units points, since a point is a pixel at 72dpi -- so
`getlength()` returns the width in the same units the slide is laid out in.

Everything degrades: no PIL, no font file, an unknown typeface, and it falls
back to the old ratio estimate rather than failing. A deck still builds on a
machine with no fonts to measure.
"""

import functools
import os
import sys

# Only used when measurement isn't available. Deliberately generous.
FALLBACK_CHAR_WIDTH_RATIO = 0.5

# A fraction of a line, not a fraction of every string: enough to absorb
# kerning and the renderer's own rounding without inflating long cells.
SAFETY_POINTS = 2.0

_FONT_DIRS = [
    os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts"),
    "/usr/share/fonts", "/Library/Fonts", os.path.expanduser("~/.fonts"),
]

# What Office reaches for when a deck names a font the machine doesn't have.
_SUBSTITUTE = "calibri"


@functools.lru_cache(maxsize=1)
def _font_index():
    """{typeface name -> file} built by reading the fonts on this machine.

    Discovered rather than hardcoded, which is not a refinement: a fixed
    list of filenames had no entry for Proxima Nova -- the brand font this
    deck actually uses -- so every measurement quietly fell back to Calibri,
    a substantially narrower face, and under-measured the whole table. A
    list you have to remember to extend fails silently the first time a
    deck uses a font nobody added, and under-measurement is the direction
    that overlaps.

    Both "family style" and the bare family are registered, so a run naming
    "Proxima Nova Extrabold" and one naming "Proxima Nova" both resolve.
    First writer wins, so Regular claims the bare family name.
    """
    try:
        from PIL import ImageFont
    except ImportError:
        return {}
    index = {}
    for directory in _FONT_DIRS:
        if not directory or not os.path.isdir(directory):
            continue
        try:
            entries = sorted(os.listdir(directory))
        except OSError:
            continue
        for entry in entries:
            if not entry.lower().endswith((".ttf", ".otf")):
                continue
            path = os.path.join(directory, entry)
            try:
                family, style = ImageFont.truetype(path, 12).getname()
            except Exception:                                    # noqa: BLE001
                continue
            if not family:
                continue
            family = family.strip()
            style = (style or "").strip()
            index.setdefault(f"{family} {style}".strip().lower(), path)
            # The bare family has to mean Regular. Taking whichever weight
            # was read first makes it alphabetical, and "Proxima Nova" then
            # measures as Black -- wider than the text will draw, so the
            # table shrinks for room it didn't need.
            if style.lower() in ("", "regular", "book", "roman"):
                index[family.lower()] = path
            else:
                index.setdefault(family.lower(), path)
    return index


@functools.lru_cache(maxsize=256)
def font_path(typeface, bold=False):
    """A font file for this typeface, or None.

    Substitutes rather than giving up: an absent font is what PowerPoint
    substitutes too, so measuring the substitute is measuring what will
    actually be drawn. A bold run is measured in the bold face when the
    machine has one -- bold is wider, and wider is the direction that wraps.
    """
    index = _font_index()
    name = (typeface or "").strip().lower()
    candidates = []
    if name:
        if bold:
            candidates += [f"{name} bold", f"{name} extrabold", f"{name} semibold"]
        candidates.append(name)
    candidates += [f"{_SUBSTITUTE} bold"] if bold else []
    candidates.append(_SUBSTITUTE)
    for candidate in candidates:
        path = index.get(candidate)
        if path:
            return path
    return None


# Always measure at this size and scale the answer down, never at the size
# actually asked for. PIL takes an integer pixel size and quantizes glyph
# advances to whole pixels, so measuring 6pt text at 6px rounds every
# character to the nearest point and rounds DOWN in aggregate -- a 44
# character cell measured 115pt against the 119.5pt it had, "fit" on that
# arithmetic, and PowerPoint wrapped it to two lines. At 512 the quantum is
# under a hundredth of a point at any size a slide uses. It also means one
# font object per typeface rather than one per typeface and size.
_REFERENCE_SIZE = 512


@functools.lru_cache(maxsize=64)
def _pil_font(path):
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    try:
        return ImageFont.truetype(path, _REFERENCE_SIZE)
    except OSError:
        return None


def measurement_available(typeface=None):
    return _pil_font(font_path(typeface)) is not None


def text_width_points(text, typeface, size_pt, bold=False):
    """Width of `text` in points, or None when it can't be measured."""
    path = font_path(typeface, bold)
    if path is None:
        return None
    font = _pil_font(path)
    if font is None:
        return None
    return font.getlength(text) * (size_pt / _REFERENCE_SIZE)


def wrapped_lines(text, available_points, typeface, size_pt, bold=False):
    """How many lines `text` takes in `available_points` of width.

    Greedy word wrap, the same way a renderer does it -- a long word that
    doesn't fit on its own occupies its line and overflows rather than
    forcing an extra one, which is what PowerPoint does too.
    """
    text = " ".join((text or "").split())
    if not text:
        return 1
    usable = max(1.0, available_points - SAFETY_POINTS)

    width = lambda s: text_width_points(s, typeface, size_pt, bold)
    if width("M") is None:
        chars_per_line = max(1.0, usable / (size_pt * FALLBACK_CHAR_WIDTH_RATIO))
        return max(1, -(-len(text) // int(chars_per_line)))

    if width(text) <= usable:
        return 1

    lines, current = 1, ""
    for word in text.split(" "):
        candidate = f"{current} {word}".strip()
        if current and width(candidate) > usable:
            lines += 1
            current = word
        else:
            current = candidate
    return lines
