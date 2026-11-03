"""Stand-ins for what a local run reads from gitignored files or Supabase, built
at runtime from data the repo already commits -- so a CI run exercises the same
code paths without anything new being published.

Installed by tests/ci/sitecustomize.py, CI only. Each one patches a module just
after it is first imported; a test that stubs the same name itself still wins,
since it assigns after the import.

  * audience_catalog: the brochure PDF (gitignored) is the offline fallback.
    The stand-in feeds load_local_catalog() the standard segments of
    audience_segment_tags.csv instead -- names and taxonomy prefixes only, the
    way the brochure supplies them. No subcategory, and RFP-selectable is the
    safe default (False), since only the brochure knows it.
  * db.fetch_market_profiles: the offline fallback has no image paths, so every
    market reads "(no profile slide)". The stand-in returns the same built-in
    rows with a placeholder path for each market that has a slide in the deck
    (market_profiles_index.json), which is what the live table holds.
"""
import csv
import importlib.abc
import importlib.machinery
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def _catalog_rows():
    import audience_catalog
    prefixes = set(audience_catalog.CATEGORY_PREFIXES)
    rows = []
    with open(REPO / "audience_segment_tags.csv", encoding="utf-8", newline="") as fh:
        for row in csv.DictReader(fh):
            first = row["segment"].split()[0] if row["segment"].split() else ""
            if row["status"] == "standard" and first in prefixes:
                rows.append({"segment": row["segment"], "category": first,
                             "subcategory": "", "rfp_selectable": False})
    return rows


def _patch_audience_catalog(module):
    import pandas as pd
    module.PDF_PATH = Path(__file__)          # exists, so the fallback runs
    module._parse_pdf_catalog = lambda: pd.DataFrame(_catalog_rows())


def _patch_db(module):
    def fetch_market_profiles():
        import market_profiles
        rows = market_profiles.build_rows()
        for row in rows:
            if row.get("has_profile_slide"):
                row["image_path"] = f"ci-stand-in/{row['key']}.png"
        return rows, None
    module.fetch_market_profiles = fetch_market_profiles


PATCHES = {"audience_catalog": _patch_audience_catalog, "db": _patch_db}


class _Loader(importlib.abc.Loader):
    def __init__(self, real, patch):
        self.real, self.patch = real, patch

    def create_module(self, spec):
        return self.real.create_module(spec)

    def exec_module(self, module):
        self.real.exec_module(module)
        self.patch(module)


class _Finder(importlib.abc.MetaPathFinder):
    def find_spec(self, name, path=None, target=None):
        if name not in PATCHES:
            return None
        spec = importlib.machinery.PathFinder.find_spec(name, path)
        if spec is not None and spec.loader is not None:
            spec.loader = _Loader(spec.loader, PATCHES[name])
        return spec


def install():
    sys.meta_path.insert(0, _Finder())
