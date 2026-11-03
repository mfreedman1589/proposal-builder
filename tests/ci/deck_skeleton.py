"""Stand-in master decks for CI: the real decks' structure with none of their
content, so the deck tests can run where the real .pptx files can't go.

    python tests/ci/deck_skeleton.py check     # locally: re-extract if a deck changed
    python tests/ci/deck_skeleton.py extract   # locally: re-extract unconditionally
    python tests/ci/deck_skeleton.py build     # in CI, before the sweep

`extract` reads each deck in DECKS -- the ACTIVE master and report template,
downloaded from Supabase, plus any pinned local file a test names -- and writes
tests/ci/decks/<role>.json.gz, recording in decks/manifest.json which version
each skeleton came from. `check` compares that record with the active versions
and, when one differs, says so, re-extracts it and exits 1: the skeletons are
committed, so CI keeps testing the old layout until the new ones are pushed.
`run_all.py --local-only` runs it first, which puts it in the nightly merge.

A skeleton is
every slide, layout, master, theme, notes page and chart part, with

  * text kept only where it is already published -- a token ({{CLIENT_NAME}}),
    a run or paragraph that appears verbatim in a committed .py/.md file, or an
    anchor string the code itself matches on. Everything else becomes filler of
    the same length (letters for letters, so text measurement stays close).
  * speaker notes reduced to their `key:` lines (SLIDE_KEYS.md);
  * every image replaced by a blank of the same format, embedded font files,
    comments, authors, alt text, custom XML and add-ins dropped, and every
    explicit colour turned to a grey of the same lightness.

So the slide keys, layouts, placeholders, tables, geometry and theme fonts the
tests rely on are the real ones, and nothing about the deck is published that
the repo didn't already contain. `build` turns each skeleton back into a .pptx
at the file names the tests and db.master_deck's fallback look for.
"""
import ast
import gzip
import io
import json
import re
import struct
import subprocess
import sys
import zipfile
from pathlib import Path, PurePosixPath

from lxml import etree

REPO = Path(__file__).resolve().parents[2]
OUT_DIR = Path(__file__).resolve().parent / "decks"

# role -> where its source deck comes from, plus extra file names `build`
# writes it under. An "active" role follows the version Supabase marks active
# (so `check` can tell when the skeleton is stale) and is also written as the
# local fallback the code opens when Supabase is unreachable, which is how CI
# reaches it (assembly.MASTER_DECK_PATH). A "pinned" role is a fixed historical
# template a test names by file.
DECKS = {
    "master": {"active": "deck_versions", "also": ["TEGNA_MASTER_DECK_v1_1.pptx"]},
    "report": {"active": "report_deck_versions", "also": []},
    "report_v0_15": {"pinned": "REPORT_MASTER_v0_15.pptx", "also": []},
}
MANIFEST = OUT_DIR / "manifest.json"
# Code whose string literals are anchors matched against slide text.
ANCHOR_SOURCES = ["slide_map.py", "assembly.py", "report_assembly.py"]

DROP_PREFIXES = ("customXml/", "docProps/custom.xml", "ppt/authors.xml", "ppt/comments/",
                 "ppt/commentAuthors.xml", "ppt/fonts/", "ppt/webextensions/",
                 "ppt/revisionInfo.xml", "ppt/changesInfos/")
TEXT_PARTS = re.compile(r"ppt/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|charts)/[^/]+\.xml$")
COLOUR_PARTS = re.compile(r"ppt/(slides|slideLayouts|slideMasters|notesMasters|theme|charts)/[^/]+\.xml$")

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
TOKEN = re.compile(r"\{\{[A-Z0-9_]+\}\}")
FILLER = "loremipsumdolorsitametconsecteturadipiscingelitseddoeiusmodtempor"

APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>Microsoft Office PowerPoint</Application></Properties>')
CORE_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>CI stand-in deck</dc:title>'
            '</cp:coreProperties>')


# ---------------------------------------------------------------- extract

def _corpus():
    names = subprocess.run(["git", "ls-files", "*.py", "*.md"], cwd=REPO, capture_output=True,
                           text=True, check=True).stdout.split()
    return "\n".join((REPO / n).read_text(encoding="utf-8", errors="replace") for n in names
                     if (REPO / n).is_file() and "tests/ci/" not in n)


def _anchors():
    found = set()
    for name in ANCHOR_SOURCES:
        for node in ast.walk(ast.parse((REPO / name).read_text(encoding="utf-8"))):
            if (isinstance(node, ast.Constant) and isinstance(node.value, str)
                    and 3 <= len(node.value) <= 80 and re.search(r"[A-Za-z]", node.value)
                    and "\n" not in node.value):
                found.add(node.value)
    return sorted(found, key=len, reverse=True)


class Sanitizer:
    def __init__(self):
        self.corpus = _corpus()
        self.anchor_re = re.compile("|".join(re.escape(a) for a in _anchors()), re.IGNORECASE)
        self.kept = self.filled = 0

    def known(self, text):
        s = text.strip()
        return (not s or not re.search(r"[A-Za-z0-9]", s) or s in self.corpus)

    def fill(self, text):
        keep = [False] * len(text)
        for pattern in (TOKEN, self.anchor_re):
            for m in pattern.finditer(text):
                keep[m.start():m.end()] = [True] * (m.end() - m.start())
        out, i = [], 0
        for ch, k in zip(text, keep):
            if k or not ch.isalnum():
                out.append(ch)
            elif ch.isdigit():
                out.append("5")
            else:
                f = FILLER[i % len(FILLER)]
                i += 1
                out.append(f.upper() if ch.isupper() else f)
        return "".join(out)

    def text_runs(self, root, notes):
        for para in root.iter(f"{{{NS['a']}}}p"):
            runs = para.findall(".//a:t", NS)
            joined = "".join(r.text or "" for r in runs)
            if notes:
                if "key:" not in joined.lower():
                    for r in runs:
                        r.text = ""
                continue
            if self.known(joined):
                self.kept += 1
                continue
            for r in runs:
                if r.text and not self.known(r.text):
                    r.text = self.fill(r.text)
                    self.filled += 1
        for v in root.iterfind(".//c:strCache//c:v", NS):
            if v.text and not self.known(v.text):
                v.text = self.fill(v.text)
        for i, v in enumerate(root.iterfind(".//c:numCache//c:v", NS), 1):
            v.text = str(i)


def _grey(hex6):
    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    y = round(0.299 * r + 0.587 * g + 0.114 * b)
    return f"{y:02X}" * 3


def _resolve(base, target):
    parts = []
    for piece in (PurePosixPath(base).parent / target).parts:
        if piece == "..":
            parts.pop()
        elif piece != ".":
            parts.append(piece)
    return "/".join(parts)


def _rels_source(rels_name):
    # ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
    p = PurePosixPath(rels_name)
    return str(p.parent.parent / p.name[:-len(".rels")]) if p.parent.name == "_rels" else ""


def _dropped(name):
    return name.startswith(DROP_PREFIXES)


def extract_parts(src, sanitizer):
    z = zipfile.ZipFile(src)
    names = z.namelist()
    parts = []
    for name in names:
        if _dropped(name):
            continue
        data = z.read(name)
        if name.startswith("ppt/media/"):
            parts.append([name, "media", PurePosixPath(name).suffix.lower().lstrip(".")])
            continue
        if name.startswith("ppt/embeddings/"):
            parts.append([name, "blank", PurePosixPath(name).suffix.lower().lstrip(".")])
            continue
        if name == "docProps/app.xml":
            parts.append([name, "xml", APP_XML])
            continue
        if name == "docProps/core.xml":
            parts.append([name, "xml", CORE_XML])
            continue
        if not name.endswith((".xml", ".rels")):
            raise SystemExit(f"{Path(src).name}: unexpected part {name} -- decide how to sanitize it")
        root = etree.fromstring(data)
        if name.endswith(".rels"):
            source = _rels_source(name)
            for rel in list(root):
                target = rel.get("Target", "")
                if rel.get("TargetMode") == "External":
                    continue
                if _dropped(_resolve(source or "x", target) if source else target.lstrip("/")):
                    root.remove(rel)
        elif name == "[Content_Types].xml":
            for o in root.findall("ct:Override", NS):
                if _dropped(o.get("PartName", "").lstrip("/")):
                    root.remove(o)
        elif name == "ppt/presentation.xml":
            # The font NAMES stay (assembly.declared_fonts reads them); the
            # embedded files and the references to them go.
            for face in root.findall("p:embeddedFontLst/p:embeddedFont", NS):
                for child in list(face):
                    if etree.QName(child).localname != "font":
                        face.remove(child)
            for lst in root.findall("p:custDataLst", NS):
                root.remove(lst)
        if TEXT_PARTS.match(name):
            sanitizer.text_runs(root, notes=name.startswith("ppt/notesSlides/"))
            for el in root.iter():
                for attr in ("descr", "title"):
                    if attr in el.attrib and el.tag.endswith("}cNvPr"):
                        del el.attrib[attr]
            for el in root.iter():
                if etree.QName(el).localname in ("commentRel", "custDataLst"):
                    el.getparent().remove(el)
        if COLOUR_PARTS.match(name):
            for el in root.iter(f"{{{NS['a']}}}srgbClr"):
                if re.fullmatch(r"[0-9A-Fa-f]{6}", el.get("val", "")):
                    el.set("val", _grey(el.get("val")))
        if name == "ppt/presentation.xml":
            for el in root.iter():
                if etree.QName(el).localname == "section" and el.get("name"):
                    if not sanitizer.known(el.get("name")):
                        el.set("name", sanitizer.fill(el.get("name")))
        parts.append([name, "xml", etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                                  standalone=True).decode("utf-8")])
    return parts


# ---------------------------------------------------------------- build

def _blank(ext):
    if ext in ("png", "jpeg", "jpg", "gif", "bmp", "tif", "tiff"):
        from PIL import Image
        buf = io.BytesIO()
        Image.new("RGB", (8, 8), (200, 200, 200)).save(
            buf, {"jpg": "JPEG", "jpeg": "JPEG", "tif": "TIFF"}.get(ext, ext.upper()))
        return buf.getvalue()
    if ext == "svg":
        return b'<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"/>'
    if ext == "emf":
        # EMR_HEADER (108 bytes) + EMR_EOF (20 bytes): an empty, valid metafile.
        header = struct.pack("<II4i4i4sIIIHHIII2i2iIII2i",
                             1, 108, 0, 0, 8, 8, 0, 0, 200, 200, b" EMF", 0x10000, 128, 2,
                             1, 0, 0, 0, 0, 1024, 768, 320, 240, 0, 0, 0, 320000, 240000)
        return header + struct.pack("<IIIII", 14, 20, 0, 16, 20)
    if ext == "xlsx":
        import openpyxl
        buf = io.BytesIO()
        openpyxl.Workbook().save(buf)
        return buf.getvalue()
    return b"\0" * 16          # wdp and other formats nothing here decodes


def build():
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    for role, record in manifest.items():
        skeleton = OUT_DIR / f"{role}.json.gz"
        parts = json.loads(gzip.decompress(skeleton.read_bytes()))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, kind, payload in parts:
                z.writestr(name, payload.encode("utf-8") if kind == "xml" else _blank(payload))
        for target in record["targets"]:
            (REPO / target).write_bytes(buf.getvalue())
            print(f"built {target} from {skeleton.name}")


# ---------------------------------------------------------------- versions

def _active_row(table):
    """The active row of a deck_versions-shaped table -- always the ACTIVE one,
    whatever DECK_CHANNEL says, since that is what the public app serves."""
    sys.path.insert(0, str(REPO))
    import db
    client = db.get_client()
    if client is None:
        raise SystemExit("Supabase isn't configured here, so the active deck versions "
                         "can't be checked -- run this where the app's secrets are.")
    rows = client.table(table).select("*").eq("active", True).limit(1).execute().data or []
    if not rows:
        raise SystemExit(f"No active row in {table}.")
    return rows[0]


def _source(role):
    """(manifest record, local path) for one role's source deck as it is now."""
    spec = DECKS[role]
    if "pinned" in spec:
        path = REPO / spec["pinned"]
        if not path.exists():
            raise SystemExit(f"{spec['pinned']} isn't here -- it's the source of the "
                             f"{role} skeleton.")
        return {"source": spec["pinned"], "targets": [spec["pinned"]] + spec["also"]}, path
    row = _active_row(spec["active"])
    import db
    fetch = (db._deck_file_for_version if spec["active"] == "deck_versions"
             else db._report_deck_file_for_version)
    path = Path(fetch(row["id"], row["storage_path"]))
    record = {"source": spec["active"], "version_id": row["id"], "filename": row["filename"],
              "storage_path": row["storage_path"],
              "targets": [row["filename"]] + [t for t in spec["also"] if t != row["filename"]]}
    return record, path


def _manifest():
    return json.loads(MANIFEST.read_text(encoding="utf-8")) if MANIFEST.exists() else {}


def _describe(record):
    if not record:
        return "nothing (no skeleton yet)"
    if "version_id" in record:
        return f"version {record['version_id']} ({record['filename']})"
    return record["source"]


def extract(roles):
    OUT_DIR.mkdir(exist_ok=True)
    manifest = _manifest()
    sanitizer = Sanitizer()
    for role in roles:
        record, path = _source(role)
        parts = extract_parts(path, sanitizer)
        out = OUT_DIR / f"{role}.json.gz"
        out.write_bytes(gzip.compress(json.dumps(parts).encode("utf-8"), 9, mtime=0))
        manifest[role] = record
        print(f"{out.name}: {len(parts)} parts, {out.stat().st_size // 1024} KiB, "
              f"from {_describe(record)}")
    MANIFEST.write_text(json.dumps({r: manifest[r] for r in DECKS if r in manifest},
                                   indent=2) + "\n", encoding="utf-8")
    print(f"paragraphs kept verbatim {sanitizer.kept}, runs filled {sanitizer.filled}")


def stale_roles():
    """[(role, recorded, current)] for each skeleton that no longer matches its
    source: an active role whose version moved, or a role never extracted."""
    manifest = _manifest()
    stale = []
    for role, spec in DECKS.items():
        recorded = manifest.get(role)
        missing = not (OUT_DIR / f"{role}.json.gz").exists()
        if "active" in spec:
            row = _active_row(spec["active"])
            if missing or not recorded or recorded.get("version_id") != row["id"]:
                stale.append((role, recorded, {"version_id": row["id"],
                                               "filename": row["filename"]}))
        elif missing or not recorded:
            stale.append((role, recorded, {"source": spec["pinned"]}))
    return stale


def check():
    """0 when every skeleton matches the deck it stands in for; otherwise
    re-extract the stale ones, say what to do next, and return 1."""
    stale = stale_roles()
    if not stale:
        for role, record in _manifest().items():
            print(f"CI stand-in {role}: {_describe(record)} -- current")
        return 0
    for role, recorded, current in stale:
        print(f"CI stand-in {role} is STALE: extracted from {_describe(recorded)}, "
              f"but the source is now {_describe(current)}. Re-extracting.")
    extract([role for role, _, _ in stale])
    print("\nCI has been testing the old layout. Commit tests/ci/decks/, push to dev, "
          "and wait for CI to pass on it before merging.")
    return 1


def main():
    if sys.argv[1:] == ["extract"]:
        extract(list(DECKS))
    elif sys.argv[1:] == ["build"]:
        build()
    elif sys.argv[1:] == ["check"]:
        sys.exit(check())
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main()
