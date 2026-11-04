"""optimize_deck never writes two parts under one name.

    python tests/test_optimize_deck.py

A re-encoded PNG is renamed to .jpeg, and the 2027 sports master had
image1.png beside an existing image1.jpeg: zipfile wrote both entries (only a
UserWarning), and one image silently replaced the other. The rename now
picks a free name. Also checks the optimizer stays idempotent.
"""
import collections
import io
import os
import random
import sys
import tempfile
import warnings
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

from PIL import Image                          # noqa: E402
from pptx import Presentation                  # noqa: E402
from pptx.opc.packuri import PackURI           # noqa: E402
from pptx.util import Inches                   # noqa: E402

import optimize_deck                           # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def noisy(fmt, seed):
    """An opaque, photo-like image -- PNG of it is big, so it gets re-encoded."""
    rnd = random.Random(seed)
    img = Image.new("RGB", (400, 300))
    img.putdata([(rnd.randrange(256), rnd.randrange(256), rnd.randrange(256))
                 for _ in range(400 * 300)])
    buf = io.BytesIO()
    img.save(buf, fmt)
    return buf.getvalue()


def main():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        jpg = slide.shapes.add_picture(io.BytesIO(noisy("JPEG", 1)), 0, 0, Inches(2))
        png = slide.shapes.add_picture(io.BytesIO(noisy("PNG", 2)), Inches(3), 0, Inches(2))
        # The collision: image1.jpeg beside image1.png, which re-encodes to
        # image1.jpeg.
        slide.part.related_part(jpg._element.blip_rId).partname = PackURI("/ppt/media/image1.jpeg")
        slide.part.related_part(png._element.blip_rId).partname = PackURI("/ppt/media/image1.png")
        src, out, again = tmp / "in.pptx", tmp / "out.pptx", tmp / "again.pptx"
        prs.save(src)

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            stats = optimize_deck.optimize_deck(str(src), str(out), verbose=False)
        check("the PNG was re-encoded (the case under test happened)", stats["converted"] == 1, stats)
        with zipfile.ZipFile(out) as z:
            names = z.namelist()
        dupes = [n for n, c in collections.Counter(names).items() if c > 1]
        check("no duplicate part names", not dupes, dupes)
        check("no zipfile duplicate-name warning",
              not [w for w in caught if "Duplicate name" in str(w.message)])
        pics = [s for s in Presentation(str(out)).slides[0].shapes if s.shape_type == 13]
        blobs = {p.image.blob for p in pics}
        check("both pictures still resolve, to two different images", len(pics) == 2 and len(blobs) == 2)
        stats = optimize_deck.optimize_deck(str(out), str(again), verbose=False)
        check("a second pass converts nothing (idempotent)", stats["converted"] == 0, stats)

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
