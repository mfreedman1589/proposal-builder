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
import logging
import os
import sys

_LOG = logging.getLogger(__name__)

# Only used when measurement isn't available. Deliberately generous.
FALLBACK_CHAR_WIDTH_RATIO = 0.5

# Every typeface this process has been asked to measure, and what it actually
# got: {(typeface, bold): {...}} with a status of "exact", "substituted" or
# "unavailable".
#
# This exists because a silent substitution has now caused the same bug twice.
# First a hardcoded name->filename map had no entry for Proxima Nova, so every
# measurement quietly became Calibri. That was fixed by discovering fonts
# instead of listing them -- and the bug came straight back by another route,
# because assembly resolved "+mn-lt" against the wrong slide master and asked
# for Calibri BY NAME. Both times the measurement was wrong, narrower than the
# truth, and nothing said so. A substitution is now recorded, logged once, and
# reported by anything that can put it in front of a person.
_RESOLUTIONS = {}

# Typefaces that fell all the way through to the character-ratio estimate.
# Tracked separately because it is an environment fact rather than a font
# fact: it is what the deployed Linux instance does, having neither the brand
# font nor Calibri, and it makes that instance lay out tables differently from
# a local Windows one.
_ESTIMATED = set()

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

# Family names that already denote bold or heavier. Asking for bold on top of
# one of these is a no-op -- "Proxima Nova Extrabold" IS the bold face -- so
# resolving it to itself is an exact match, not a substitution. Without this
# the log cried wolf on every deck ("asked for Proxima Nova Extrabold bold,
# using Proxima Nova Extrabold, no bold face; measured regular"), which is
# the failure this reporting exists to prevent: a channel nobody reads.
# "semibold" is deliberately absent -- it is lighter than bold, so falling
# back to it really is measuring something the renderer may embolden.
_ALREADY_BOLD = {"bold", "extrabold", "ultrabold", "black", "heavy"}


def _name_denotes_bold(name):
    """Whether a family name already carries a bold-or-heavier weight.

    Matched on whole words, not substrings: "semibold" contains "bold" and a
    naive `in` would silently treat it as already-bold. Same trap as the
    vertical hints, where "auto" fired on "automatic".
    """
    return any(token in _ALREADY_BOLD
               for token in (name or "").replace("-", " ").split())


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


def _record(typeface, bold, path, status, detail=""):
    """Remember what a typeface resolved to, and say so once if it isn't
    what was asked for."""
    key = ((typeface or "").strip().lower(), bool(bold))
    entry = _RESOLUTIONS.get(key)
    if entry is None:
        entry = {"requested": typeface, "bold": bool(bold), "path": path,
                 "status": status, "detail": detail,
                 "resolved": _family_of(path) if path else None}
        _RESOLUTIONS[key] = entry
        if status != "exact":
            _LOG.warning(
                "font substitution: asked for %r%s, using %s (%s)",
                typeface, " bold" if bold else "",
                entry["resolved"] or "a character-width estimate", detail or status)
    return entry


@functools.lru_cache(maxsize=256)
def _family_of(path):
    """The family/style a font file actually declares."""
    try:
        from PIL import ImageFont
        family, style = ImageFont.truetype(path, 12).getname()
        return f"{family} {style}".strip()
    except Exception:                                            # noqa: BLE001
        return os.path.basename(path)


def resolve_font(typeface, bold=False):
    """(path, status) for this typeface. Never substitutes silently.

    Substituting is still the right behaviour -- an absent font is what
    PowerPoint substitutes too, so measuring the substitute is measuring what
    will actually be drawn. What was wrong before was doing it without a
    word: see _RESOLUTIONS.

    A bold run resolves to the bold face when the machine has one, since bold
    is wider and wider is the direction that wraps. Falling back to the
    regular weight is reported as a substitution rather than a match,
    precisely because it measures narrower than what gets drawn.
    """
    index = _font_index()
    name = (typeface or "").strip().lower()
    if name:
        if bold:
            for suffix in ("bold", "extrabold", "semibold"):
                path = index.get(f"{name} {suffix}")
                if path:
                    return _record(typeface, bold, path, "exact")
            path = index.get(name)
            if path:
                if _name_denotes_bold(name):
                    return _record(typeface, bold, path, "exact")
                return _record(typeface, bold, path, "substituted",
                               "no bold face on this machine; measured regular")
        else:
            path = index.get(name)
            if path:
                return _record(typeface, bold, path, "exact")
    for candidate in ([f"{_SUBSTITUTE} bold"] if bold else []) + [_SUBSTITUTE]:
        path = index.get(candidate)
        if path:
            return _record(typeface, bold, path, "substituted",
                           f"{typeface!r} is not installed here")
    return _record(typeface, bold, None, "unavailable",
                   f"neither {typeface!r} nor {_SUBSTITUTE} is installed here")


@functools.lru_cache(maxsize=256)
def font_path(typeface, bold=False):
    """A font file for this typeface, or None."""
    return resolve_font(typeface, bold)["path"]


def is_available(typeface, bold=False):
    """Whether this exact typeface can be measured here.

    Deliberately does NOT record a resolution: it's for reporting on the
    fonts a DECK declares, which routinely include faces no table ever uses,
    and logging those as substitutions would bury the ones that matter.
    """
    index = _font_index()
    name = (typeface or "").strip().lower()
    if not name:
        return False
    if bold and any(index.get(f"{name} {s}") for s in ("bold", "extrabold", "semibold")):
        return True
    return bool(index.get(name))


def resolution_report():
    """Every typeface asked for this process and what was actually used,
    imperfect resolutions first."""
    return sorted(_RESOLUTIONS.values(),
                  key=lambda e: (e["status"] == "exact", (e["requested"] or "").lower()))


def imperfect_resolutions():
    return [e for e in _RESOLUTIONS.values() if e["status"] != "exact"]


def measurement_note():
    """One plain sentence about anything that wasn't measured as drawn, or
    None when every typeface resolved exactly.

    Aimed at whoever is looking at the deck. The estimate case is called out
    separately from the substitution case because it is the one that makes a
    deck built on the deployed instance lay out differently from the same
    deck built locally -- the estimate is deliberately generous, so it
    reserves more room, shrinks type sooner, and so compresses the Included
    band on plans that would have kept it.
    """
    if _ESTIMATED:
        names = ", ".join(sorted(n for n in _ESTIMATED if n))
        return (f"No font files were available to measure {names} on this machine, so "
                f"table text was estimated from average character width rather than "
                f"measured. The estimate errs wide, so tables may use smaller type and "
                f"compress the \"Included with Campaign\" band sooner than the same "
                f"proposal built on a machine that has the brand fonts installed.")
    imperfect = imperfect_resolutions()
    if imperfect:
        # Each one carries its own reason. Without that, a weight fallback
        # renders as "Proxima Nova Light -> Proxima Nova Light", which reads
        # as a bug in the reporting rather than as the real finding it is:
        # the deck asks for a bold face the machine hasn't got, so bold text
        # is being measured in the regular weight -- narrower than it draws,
        # which is the direction that overflows.
        parts = sorted(
            f"{e['requested']}{' bold' if e['bold'] else ''} "
            f"({e['detail'] or 'measured as ' + str(e['resolved'])})"
            for e in imperfect)
        return ("Some table text was measured in a font other than the one it will be "
                "drawn in: " + "; ".join(parts) + ". The table may be sized slightly "
                "differently from how it renders.")
    return None


def reset_resolution_log():
    """Forget what has been resolved so far. For tests that want to assert on
    a single build's worth of resolutions."""
    _RESOLUTIONS.clear()
    _ESTIMATED.clear()
    font_path.cache_clear()


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
        # Recorded, not silent: this is the deployed instance's normal path
        # and it lays tables out differently from a machine with the fonts.
        _ESTIMATED.add(typeface)
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
