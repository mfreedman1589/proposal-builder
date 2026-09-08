"""db.delete_slide_vault_entry -- the hard delete Matt asked for alongside
deactivation, and the two ways it can go wrong if built carelessly:

    python tests/test_slide_vault_delete.py

Deactivation (update_slide_vault_entry(..., active=False)) already covers
"retire this slide" -- it doesn't cover "this shouldn't be in a shared
library at all" (a test upload, a wrong file). A genuine hard delete is a
real risk this codebase has specifically engineered around elsewhere:
fetch_case_study/fetch_slide_vault_entry both ignore the `active` flag so
"Rebuild as presented" can reproduce a proposal that used a since-
deactivated slide -- which only works because the row still exists. Delete
this row out from under a proposal that used it and rebuild breaks. So this
checks every logged proposal's own form_json["vault_slides"] first and
refuses, unchanged, if any of them reference the id.

The second hazard is the shared file: several slide_vault rows can point at
one storage_path (one uploaded deck, picked slide by slide -- see
test_slide_vault.py). Deleting one row must never delete the file while a
sibling row still needs it.

Everything here is stubbed against fakes -- no real Supabase table is
touched, same reason test_feedback.py stubs its own store.
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class FakeTable:
    def __init__(self, store, name):
        self.store, self.name = store, name
        self._filters = {}
        self._op = None
        self._payload = None

    def select(self, *_a, **_k):
        self._op = "select"
        return self

    def delete(self):
        self._op = "delete"
        return self

    def eq(self, field, value):
        self._filters[field] = value
        return self

    def limit(self, *_a, **_k):
        return self

    def order(self, *_a, **_k):
        return self

    def execute(self):
        rows = self.store.tables.get(self.name, [])
        matched = [r for r in rows
                  if all(r.get(k) == v for k, v in self._filters.items())]
        if self._op == "delete":
            remaining = [r for r in rows if r not in matched]
            self.store.tables[self.name] = remaining
            self.store.deleted.append((self.name, self._filters.get("id")))
        return _Result(matched)


class _Result:
    def __init__(self, data):
        self.data = data


class FakeStorageBucket:
    def __init__(self, store, bucket):
        self.store, self.bucket = store, bucket

    def remove(self, paths):
        self.store.removed_files.append((self.bucket, list(paths)))


class FakeStorage:
    def __init__(self, store):
        self.store = store

    def from_(self, bucket):
        return FakeStorageBucket(self.store, bucket)


class FakeClient:
    def __init__(self, store):
        self.store = store
        self.storage = FakeStorage(store)

    def table(self, name):
        return FakeTable(self.store, name)


class FakeStore:
    """Holds slide_vault rows directly (so FakeTable can see them) and a
    separate `proposals` list db.fetch_proposals is stubbed to read from --
    proposals aren't a table this delete touches directly, only the
    usage-check reads them."""
    def __init__(self, slide_vault_rows, proposals):
        self.tables = {"slide_vault": list(slide_vault_rows)}
        self.proposals = proposals
        self.deleted = []
        self.removed_files = []


def wire(store):
    db.get_client = lambda: FakeClient(store)
    db.fetch_proposals = lambda limit=10000: (store.proposals, None)


def main():
    print("=" * 78)
    print("SCENARIO  refuses to delete a vault slide a logged proposal still uses")
    print("=" * 78)
    store = FakeStore(
        slide_vault_rows=[{"id": "v1", "storage_path": "deck1.pptx", "title": "Chart A"}],
        proposals=[{"client_name": "Acme",
                   "form_json": {"vault_slides": [{"id": "v1", "title": "Chart A"}]}}],
    )
    wire(store)
    ok, error = db.delete_slide_vault_entry("v1")
    check("delete refused", ok is False)
    check("error names the using proposal", error and "Acme" in error, error)
    check("row was NOT removed from the table",
          any(r["id"] == "v1" for r in store.tables["slide_vault"]))
    check("no storage file was touched", store.removed_files == [], store.removed_files)

    print()
    print("=" * 78)
    print("SCENARIO  deletes cleanly and removes the file when no sibling shares it")
    print("=" * 78)
    store = FakeStore(
        slide_vault_rows=[{"id": "v1", "storage_path": "deck1.pptx", "title": "Chart A"}],
        proposals=[],
    )
    wire(store)
    ok, error = db.delete_slide_vault_entry("v1")
    check("delete succeeded", ok is True, error)
    check("row removed from the table",
          not any(r["id"] == "v1" for r in store.tables["slide_vault"]))
    check("storage file was removed",
          store.removed_files == [("slide_vault", ["deck1.pptx"])], store.removed_files)

    print()
    print("=" * 78)
    print("SCENARIO  a sibling row sharing storage_path keeps the file alive")
    print("=" * 78)
    store = FakeStore(
        slide_vault_rows=[
            {"id": "v1", "storage_path": "deck1.pptx", "title": "Chart A"},
            {"id": "v2", "storage_path": "deck1.pptx", "title": "Chart B"},
        ],
        proposals=[],
    )
    wire(store)
    ok, error = db.delete_slide_vault_entry("v1")
    check("delete succeeded", ok is True, error)
    check("only the target row was removed",
          [r["id"] for r in store.tables["slide_vault"]] == ["v2"])
    check("the shared file was NOT removed", store.removed_files == [], store.removed_files)

    print()
    print("=" * 78)
    print("SCENARIO  a nonexistent id fails cleanly")
    print("=" * 78)
    store = FakeStore(slide_vault_rows=[], proposals=[])
    wire(store)
    ok, error = db.delete_slide_vault_entry("ghost")
    check("delete refused for a missing row", ok is False)
    check("error says so", bool(error), error)


if __name__ == "__main__":
    main()
    print(f"\n{len(failures)} failure(s)." if failures else "\nAll checks passed.")
    sys.exit(1 if failures else 0)
