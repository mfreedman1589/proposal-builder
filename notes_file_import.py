"""notes_file_import.py -- read an uploaded meeting-notes file into plain
text, for the intake area's "upload notes" slot (roadmap: UX sweep,
BACKLOG.md).

Pure -- no Streamlit, no DB, no network -- same shape as
`avails_pdf_import.py`/`wideorbit.py`: one function, called with a path and
the upload's own filename, returning either real text or a message aimed at
the seller who uploaded the file, never raising.

Three formats, by extension:
  - `.txt` -- decoded as UTF-8, with bad bytes replaced rather than raised on
    (a rep's notes file is not a place to be strict about encoding).
  - `.pdf` -- `pdfplumber`, already a pinned dependency (the audience-catalog
    fallback uses it too): `extract_text()` per page, joined with blank
    lines.
  - `.docx` -- `python-docx`'s own `Document`, paragraphs joined with
    newlines. The one format-specific dependency this module adds.

Deliberately not trying to be a general document-text-extraction library --
three extensions, three small branches. A format this doesn't recognize, or a
file that fails to open as the format its extension claims, is reported, not
guessed at.
"""
from pathlib import Path

import pdfplumber
from docx import Document


def extract_notes_text(path, filename):
    """(text, error) -- exactly one is not None. `filename` (not `path`,
    which is often a scratch-dir temp name) is what the error message and
    the extension check are based on, since that's what the rep actually
    named the file.
    """
    suffix = Path(filename).suffix.lower()
    try:
        if suffix == ".txt":
            return _extract_txt(path), None
        if suffix == ".pdf":
            return _extract_pdf(path), None
        if suffix == ".docx":
            return _extract_docx(path), None
    except Exception as exc:                                              # noqa: BLE001
        return None, (f"Couldn't read \"{filename}\" as a {suffix} file: {exc}. "
                      f"Paste the notes directly instead.")
    return None, (f"\"{filename}\" isn't a format this can read (.txt, .pdf and .docx "
                  f"are supported). Paste the notes directly instead.")


def _extract_txt(path):
    # read_bytes -- never text mode -- so encoding is decided here, not by
    # whatever newline translation open() would otherwise apply, and a
    # Windows-authored file's \r\n is normalized to \n like every other
    # source of notes text in this app (paste, PDF, docx) already is.
    raw = Path(path).read_bytes().decode("utf-8", errors="replace")
    return raw.replace("\r\n", "\n").replace("\r", "\n")


def _extract_pdf(path):
    with pdfplumber.open(path) as pdf:
        pages = [page.extract_text() or "" for page in pdf.pages]
    return "\n\n".join(pages).strip()


def _extract_docx(path):
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs).strip()
