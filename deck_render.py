"""deck_render.py -- render a generated .pptx to one PNG per slide.

Test tooling, not part of the app. It exists because every layout failure
this project has shipped was invisible to an assertion and obvious in a
picture: a Comscore footer read as a programme row, a table whose declared
height cleared the floor while the rendered one didn't, per-page impressions
that contradicted their own spot counts. Geometry checks confirm what you
thought to measure; a rendering shows what's actually there.

Uses PowerPoint itself through COM rather than LibreOffice. Not a preference
-- LibreOffice is blocked by policy on this machine -- but it turns out to be
the better answer anyway: this is the same renderer the client will open the
deck in, so what it draws is what they see, including font substitution and
autofit behaviour a converter would approximate.

    python deck_render.py some_deck.pptx out_dir/

Requires Windows with PowerPoint and `pywin32`. `render_deck` raises
`RenderUnavailable` anywhere else, and callers are expected to skip rather
than fail -- a machine without PowerPoint can still run every other suite.
"""

import os
import tempfile
import shutil
import sys
import time
from pathlib import Path


class RenderUnavailable(RuntimeError):
    """No renderer on this machine. Callers should skip, not fail."""


def renderer_available():
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


# Where rendered decks and images go. Deliberately NOT inside the repo: the
# project lives in a OneDrive-synced folder, and OneDrive treats a generated
# .pptx there as co-authored -- PowerPoint then raises a modal "resolve
# conflict" prompt, which blocks COM automation entirely (every subsequent
# call fails with "Cannot perform this action with a modal dialog displayed")
# and needs a human to dismiss it. Same family as the file-lock problems this
# renderer already works around. Override with PROPOSAL_BUILDER_RENDER_DIR.
RENDER_DIR_ENV = "PROPOSAL_BUILDER_RENDER_DIR"


def render_root():
    """Local, non-synced directory to write rendered output into."""
    configured = os.environ.get(RENDER_DIR_ENV)
    if configured:
        return Path(configured)
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return Path(base) / "Temp" / "proposal-builder-renders" if os.environ.get("LOCALAPPDATA")         else Path(base) / "proposal-builder-renders"


def render_deck(pptx_path, out_dir, width=1600):
    """Render every slide to PNG. Returns the list of image paths, in order.

    The deck is opened without a window and closed again whatever happens --
    an orphaned PowerPoint process holds a file lock that makes the next run
    fail with a bare PackageNotFoundError, which is a genuinely confusing way
    to discover you left one running.
    """
    if not renderer_available():
        raise RenderUnavailable(
            "Rendering needs Windows with PowerPoint and pywin32 installed.")

    import pythoncom
    import win32com.client

    source = Path(pptx_path).resolve()
    target = Path(out_dir).resolve()
    # Clear the old images individually rather than removing the directory:
    # this repo lives in a OneDrive folder, and sync holds a handle on a
    # directory it's mid-upload, which turns rmtree into PermissionError.
    target.mkdir(parents=True, exist_ok=True)
    for stale in target.glob("slide_*.png"):
        try:
            stale.unlink()
        except OSError:
            pass

    pythoncom.CoInitialize()
    powerpoint = presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # ReadOnly and WithWindow=False: the deck is a build artefact, and a
        # visible window steals focus from whoever is using the machine.
        # One retry: OneDrive can still hold a freshly written 25MB deck when
        # PowerPoint first reaches for it, which surfaces as a flat "could
        # not open the file" on a deck that is perfectly well formed.
        for attempt in range(2):
            try:
                presentation = powerpoint.Presentations.Open(
                    str(source), ReadOnly=True, WithWindow=False)
                break
            except Exception:                                    # noqa: BLE001
                if attempt:
                    raise
                time.sleep(3)
        height = int(width * presentation.PageSetup.SlideHeight
                     / presentation.PageSetup.SlideWidth)
        for index in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(index)
            slide.Export(str(target / f"slide_{index:03d}.png"), "PNG", width, height)
    except Exception as exc:                                     # noqa: BLE001
        raise RenderUnavailable(f"PowerPoint couldn't render this deck: {exc}") from exc
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        finally:
            try:
                if powerpoint is not None:
                    powerpoint.Quit()
            except Exception:                                    # noqa: BLE001
                pass
            pythoncom.CoUninitialize()

    return sorted(target.glob("slide_*.png"))


def render_presentation(prs, out_dir, name="deck", width=1600):
    """Same, for an in-memory Presentation -- saves it beside the images so
    the .pptx that produced a suspect slide is always at hand."""
    target = Path(out_dir).resolve()
    target.mkdir(parents=True, exist_ok=True)
    saved = target / f"{name}.pptx"
    prs.save(str(saved))
    return render_deck(saved, target / f"{name}_slides"), saved


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    out = argv[2] if len(argv) > 2 else "rendered"
    try:
        images = render_deck(argv[1], out)
    except RenderUnavailable as exc:
        print(f"SKIP -- {exc}")
        return 0
    print(f"{len(images)} slide(s) -> {Path(out).resolve()}")
    for image in images:
        print("  ", image.name)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
