"""An uploaded deck/report template is INACTIVE unless activation is asked for.

    python tests/test_upload_default_inactive.py

db.upload_deck and db.upload_report_deck default to activate=False: a
dev-time upload must never go live by omitting an argument. Activation is the
nightly merge's job (activate_*_version) or the admin "Activate this deck"
button's, and every caller that means to activate passes activate=True.

Offline: a fake in-memory client stands in for Supabase, so nothing is ever
written to production.
"""
import os
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import db  # noqa: E402

failures = []


def check(label, condition, detail=""):
    print(f"  {'PASS' if condition else 'FAIL'}  {label}{'' if condition else '  -- ' + str(detail)}")
    if not condition:
        failures.append(label)


class _Query:
    def __init__(self, store, table):
        self.store, self.table = store, table

    def insert(self, row):
        row = dict(row, id=len(self.store["rows"]) + 1)
        self.store["rows"].append((self.table, row))
        self._data = [row]
        return self

    def execute(self):
        return type("Result", (), {"data": self._data})()


class _Bucket:
    def __init__(self, store, name):
        self.store, self.name = store, name

    def upload(self, path, data, options):
        self.store["uploads"].append((self.name, path))


class _FakeClient:
    def __init__(self):
        self.store = {"rows": [], "uploads": []}
        self.storage = type("Storage", (), {"from_": lambda _s, name: _Bucket(self.store, name)})()

    def table(self, name):
        return _Query(self.store, name)


def run(fn, *args, **kwargs):
    fake = _FakeClient()
    activated = []
    saved = (db.get_client, db.activate_deck_version, db.activate_report_deck_version,
             db.storage_limit_bytes)
    db.get_client = lambda: fake
    db.activate_deck_version = lambda vid: (activated.append(("deck", vid)), (True, None))[1]
    db.activate_report_deck_version = lambda vid: (activated.append(("report", vid)), (True, None))[1]
    db.storage_limit_bytes = lambda *_a, **_k: 10 ** 12
    try:
        result = fn(*args, **kwargs)
    finally:
        (db.get_client, db.activate_deck_version, db.activate_report_deck_version,
         db.storage_limit_bytes) = saved
    return result, fake.store, activated


def main():
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as handle:
        handle.write(b"not really a deck")
    path = handle.name
    try:
        print("\nupload_deck")
        (row, _stats, error), store, activated = run(db.upload_deck, path, "masters/x.pptx",
                                                     optimize=False)
        check("uploads and registers a version", error is None and store["rows"], error)
        check("left inactive by default", row and row["active"] is False and not activated,
              (row, activated))
        (row, _stats, error), _store, activated = run(db.upload_deck, path, "masters/x.pptx",
                                                      optimize=False, activate=True)
        check("activate=True still activates", activated == [("deck", 1)] and row["active"], activated)

        print("\nupload_report_deck")
        (row, error), store, activated = run(db.upload_report_deck, path, "masters/r.pptx")
        check("uploads and registers a version", error is None and store["rows"], error)
        check("left inactive by default", row and row["active"] is False and not activated,
              (row, activated))
        (row, error), _store, activated = run(db.upload_report_deck, path, "masters/r.pptx",
                                              activate=True)
        check("activate=True still activates", activated == [("report", 1)] and row["active"],
              activated)

        print("\ncallers that mean to activate say so")
        for name in ("app.py", "setup_supabase.py"):
            text = (REPO / name).read_text(encoding="utf-8")
            calls = text.count("db.upload_deck(") + text.count("db.upload_report_deck(")
            explicit = text.count("activate=True")
            check(f"{name}: every upload call passes activate=True ({calls} call(s))",
                  calls and explicit >= calls, (calls, explicit))
    finally:
        os.unlink(path)

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
