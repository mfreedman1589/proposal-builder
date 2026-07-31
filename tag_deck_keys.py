"""
tag_deck_keys.py -- stamp each slide's condition_key into its speaker notes.

Derives the slide map the usual way (text anchors, see slide_map.py) and
writes the result back into every slide's notes as a line of the form

    key: vertical:healthcare

Once a deck is tagged, slide_map reads those lines in preference to the
anchor rules, which means the map stops depending on the deck's prose. Retitle
a slide, reword a heading, reorder the deck -- the key travels with the slide.
An untagged slide still resolves by anchors exactly as before, so a deck can
be half-tagged without anything breaking.

The label lives in the *speaker notes*, so it is invisible in the deck as
presented, and it is a plain line of text anyone can read and correct in
PowerPoint (View -> Notes Page).

Existing notes are preserved: the key line is appended as its own paragraph,
and re-running replaces only that line rather than accumulating duplicates or
flattening whatever else the notes say.

    python tag_deck_keys.py TEGNA_MASTER_DECK_v1_1_optimized_q72.pptx
    python tag_deck_keys.py in.pptx -o out.pptx --report slide_keys.md

Never writes in place -- the output is a separate file.
"""

import argparse
import sys
from pathlib import Path

from pptx import Presentation

import slide_map

# One definition of the label format, shared with the reader in slide_map, so
# the writer and the reader can't drift apart.
KEY_LINE = slide_map.NOTES_KEY_LINE


def write_notes_key(slide, key):
    """Set this slide's key line, leaving every other note paragraph alone."""
    text_frame = slide.notes_slide.notes_text_frame

    # Drop any key lines already there, so re-running is idempotent rather
    # than additive.
    for paragraph in list(text_frame.paragraphs):
        if KEY_LINE.match(paragraph.text):
            paragraph._p.getparent().remove(paragraph._p)

    # Never assign text_frame.text -- it would collapse the seller's own
    # notes into one unformatted run. A new paragraph touches nothing else.
    if text_frame.paragraphs and text_frame.paragraphs[-1].text.strip():
        paragraph = text_frame.add_paragraph()
    else:
        paragraph = text_frame.paragraphs[-1] if text_frame.paragraphs else text_frame.add_paragraph()
    paragraph.text = f"key: {key}"


def slide_title(slide, max_length=95):
    """A human label for the reference list.

    The title placeholder when there is one. Many slides in this deck are
    built entirely from free-floating text boxes with no title placeholder at
    all, though, and their first text frame is often a stray bullet rather
    than the heading -- so the fallback stitches the first few fragments
    together, which is enough to recognise a slide by.
    """
    title_shape = None
    try:
        title_shape = slide.shapes.title
    except (AttributeError, ValueError):
        pass
    if title_shape is not None and title_shape.has_text_frame:
        text = " ".join(title_shape.text_frame.text.split())
        if text:
            return text[:max_length]

    # No title placeholder, which is the norm in this deck -- most slides are
    # built from free-floating text boxes. Shape order is no help (the first
    # text box is routinely a stray bullet), but type size is: on a
    # viewership slide the heading is 54pt against 11pt bullets. Pick the
    # biggest explicitly-sized text, and fall back to document order only
    # when nothing carries an explicit size.
    biggest, biggest_size, first = None, 0, None
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        text = " ".join(shape.text_frame.text.split())
        if not text:
            continue
        if first is None:
            first = text
        sizes = [run.font.size.pt
                 for para in shape.text_frame.paragraphs
                 for run in para.runs if run.font.size]
        if sizes and max(sizes) > biggest_size:
            biggest, biggest_size = text, max(sizes)
    return (biggest or first or "(no text)")[:max_length]


def tag_deck(src_path, dst_path, report_path=None):
    prs = Presentation(src_path)
    derived = slide_map.build_slide_map_from_prs(prs)

    rows = []
    tagged = 0
    for number, slide in enumerate(prs.slides, start=1):
        key = derived.get(number)
        if key is None:
            print(f"  slide {number} does not resolve -- left untagged")
            rows.append((number, slide_title(slide), None))
            continue
        write_notes_key(slide, key)
        tagged += 1
        rows.append((number, slide_title(slide), key))

    prs.save(dst_path)
    print(f"  tagged {tagged} of {len(rows)} slides -> {dst_path}")

    if report_path:
        write_report(rows, report_path, Path(src_path).name)
        print(f"  reference list -> {report_path}")
    return rows


def write_report(rows, report_path, deck_name):
    """A flat slide -> title -> key list, for eyeballing the labels."""
    width = max((len(row[2] or "") for row in rows), default=3)
    lines = [
        f"# Master deck slide keys -- {deck_name}",
        "",
        "Generated by `tag_deck_keys.py`. Each slide's `key:` line is stored in its own",
        "speaker notes, so this file is a convenience copy, not the source of truth --",
        "to correct one, edit the notes in PowerPoint (View -> Notes Page) and re-run",
        "`python slide_map.py <deck>` to confirm.",
        "",
        f"{len(rows)} slides, "
        f"{len(sorted({row[2] for row in rows if row[2]}))} distinct keys.",
        "",
        "| # | key | title |",
        "|---:|---|---|",
    ]
    for number, title, key in rows:
        label = key or "**UNRESOLVED**"
        clean = title.replace("|", "\\|")[:90]
        lines.append(f"| {number} | `{label}` | {clean} |")

    counts = {}
    for _, _, key in rows:
        counts[key or "(unresolved)"] = counts.get(key or "(unresolved)", 0) + 1
    lines += ["", "## Keys by frequency", "", "| key | slides |", "|---|---:|"]
    for key, count in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        lines.append(f"| `{key}` | {count} |")

    Path(report_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("source", help="deck to tag")
    parser.add_argument("-o", "--output", help="output path (default: <source>_tagged.pptx)")
    parser.add_argument("--report", help="write a slide -> title -> key reference list here")
    args = parser.parse_args(argv)

    source = Path(args.source)
    if not source.exists():
        print(f"No such deck: {source}")
        return 1
    output = Path(args.output) if args.output else source.with_name(source.stem + "_tagged.pptx")

    print(f"Tagging {source.name}")
    tag_deck(str(source), str(output), args.report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
