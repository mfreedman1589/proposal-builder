"""db.py's new advertisers/attribution_reports functions (Stage 13/14,
ATTRIBUTION_REPORT_PLAN.md Phase 2), against a minimal fake Supabase client
-- exercises the REAL db.py code, not a reimplementation of it, the same
distinction test_feedback.py draws for why it stubs whole functions instead
(there, app.py's own logic was under test; here, db.py's query construction
is). The real `advertisers`/`attribution_reports` tables don't exist in
Supabase until the migration in supabase_schema.sql is pasted in by hand, so
this must never touch a real client either way.

    python tests/test_advertisers_db.py
"""
import sys
import uuid
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import db  # noqa: E402


class _FakeResult:
    def __init__(self, data, count=None):
        self.data = data
        self.count = count


class _FakeQuery:
    def __init__(self, table, op, payload=None):
        self.table, self.op, self.payload = table, op, payload
        self.filters = []
        self.order_col, self.order_desc = None, False
        self.limit_n = None
        self.count_mode = None

    def eq(self, col, val):
        self.filters.append((col, val))
        return self

    def order(self, col, desc=False):
        self.order_col, self.order_desc = col, desc
        return self

    def limit(self, n):
        self.limit_n = n
        return self

    def _filtered(self):
        rows = self.table.rows
        for col, val in self.filters:
            rows = [r for r in rows if r.get(col) == val]
        return rows

    def execute(self):
        if self.op == "select":
            rows = self._filtered()
            if self.count_mode:
                return _FakeResult(rows, count=len(rows))
            if self.order_col:
                rows = sorted(rows, key=lambda r: r.get(self.order_col) or "",
                             reverse=self.order_desc)
            if self.limit_n:
                rows = rows[:self.limit_n]
            return _FakeResult(list(rows))
        if self.op == "insert":
            new_rows = self.payload if isinstance(self.payload, list) else [self.payload]
            inserted = []
            for raw in new_rows:
                row = dict(raw)
                row.setdefault("id", str(uuid.uuid4()))
                row.setdefault("created_at", f"2026-01-01T00:00:{len(self.table.rows):02d}")
                if self.table.name == "advertisers":
                    # Models supabase_schema.sql's `active boolean not null
                    # default true` (Stage 17) -- a real insert never omits
                    # this column, and db.create_advertiser relies on that
                    # DEFAULT rather than setting it itself.
                    row.setdefault("active", True)
                self.table.rows.append(row)
                inserted.append(row)
            return _FakeResult(inserted)
        if self.op == "update":
            matched = self._filtered()
            for row in matched:
                row.update(self.payload)
            return _FakeResult(matched)
        if self.op == "delete":
            matched = self._filtered()
            for row in matched:
                self.table.rows.remove(row)
            return _FakeResult(matched)
        raise NotImplementedError(self.op)


class _FakeTable:
    def __init__(self, name=""):
        self.name = name
        self.rows = []

    def select(self, *args, count=None):
        query = _FakeQuery(self, "select")
        query.count_mode = count
        return query

    def insert(self, payload):
        return _FakeQuery(self, "insert", payload)

    def update(self, payload):
        return _FakeQuery(self, "update", payload)

    def delete(self):
        return _FakeQuery(self, "delete")


class _FakeStorageBucket:
    """In-memory stand-in for one Supabase storage bucket -- just enough
    for upload_report_file/delete_attribution_report's own cleanup to
    exercise real db.py code without touching real storage."""

    def __init__(self):
        self.objects = {}

    def upload(self, path, data, options=None):
        self.objects[path] = data
        return {"path": path}

    def download(self, path):
        return self.objects[path]

    def remove(self, paths):
        for path in paths:
            self.objects.pop(path, None)
        return []


class _FakeStorage:
    def __init__(self):
        self.buckets = {}

    def from_(self, name):
        return self.buckets.setdefault(name, _FakeStorageBucket())


class FakeSupabaseClient:
    def __init__(self):
        self._tables = {}
        self.storage = _FakeStorage()

    def table(self, name):
        return self._tables.setdefault(name, _FakeTable(name))


failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def equal(label, actual, expected):
    check(label, actual == expected, f"expected {expected!r}, got {actual!r}")


def main():
    fake = FakeSupabaseClient()
    real_get_client = db.get_client
    db.get_client = lambda: fake
    try:
        print("advertiser_name_key")
        equal("  trims + case-folds + collapses whitespace",
             db.advertiser_name_key("  Cardinal   Plumbing  "), "cardinal plumbing")

        print("create_advertiser")
        row1, error1 = db.create_advertiser("Cardinal Plumbing")
        check("  first create succeeds with no error", error1 is None, error1)
        check("  row carries the display name as typed",
             row1 is not None and row1.get("canonical_name") == "Cardinal Plumbing", row1)

        row2, error2 = db.create_advertiser("  cardinal plumbing  ")
        check("  re-creating under a different case/whitespace is a no-op", error2 is None, error2)
        equal("  it returns the SAME row, not a new one", row2.get("id"), row1.get("id"))
        equal("  exactly one row exists in the fake table",
             len(fake.table("advertisers").rows), 1)

        row3, error3 = db.create_advertiser("Mattress Warehouse")
        check("  a genuinely different name creates a second row", error3 is None, error3)
        check("  it's a distinct id from the first", row3.get("id") != row1.get("id"), row3)

        blank, blank_error = db.create_advertiser("   ")
        check("  a blank name is refused, not silently created", blank is None and blank_error, blank_error)

        print("fetch_advertisers")
        rows, warning = db.fetch_advertisers()
        check("  no warning", warning is None, warning)
        equal("  both real advertisers present", len(rows), 2)
        equal("  alphabetical by canonical_name",
             [r["canonical_name"] for r in rows], ["Cardinal Plumbing", "Mattress Warehouse"])

        print("link_proposal_advertiser")
        fake.table("proposals").rows.append({"id": "prop-1", "client_name": "Cardinal Plumbing",
                                             "advertiser_id": None})
        ok, error = db.link_proposal_advertiser("prop-1", row1["id"])
        check("  ok", ok and error is None, error)
        equal("  the proposal row now carries the advertiser id",
             fake.table("proposals").rows[0]["advertiser_id"], row1["id"])

        print("log_attribution_report + fetch_attribution_reports")
        rid1, err1 = db.log_attribution_report(row1["id"], "prop-1", {"delivered": 917451},
                                               created_by="Matt")
        check("  first report logs cleanly", err1 is None, err1)
        rid2, err2 = db.log_attribution_report(row3["id"], None, {"delivered": 2341223})
        check("  second report (no-proposal mode: proposal_id=None) logs cleanly", err2 is None, err2)
        check("  the two reports got distinct ids", rid1 != rid2, (rid1, rid2))

        all_reports, warning = db.fetch_attribution_reports()
        equal("  both reports returned with no scope", len(all_reports), 2)

        scoped, warning = db.fetch_attribution_reports(advertiser_id=row1["id"])
        equal("  scoping to one advertiser returns only its report", len(scoped), 1)
        equal("  it's the Cardinal one", scoped[0]["id"], rid1)
        equal("  report_json round-trips", scoped[0]["report_json"], {"delivered": 917451})
        equal("  default status is 'parsed'", scoped[0]["status"], "parsed")

        # --- Stage 17: active_only, rename, merge, report files, delete ---
        print("fetch_advertisers(active_only=True) -- every row starts active")
        active_rows, warning = db.fetch_advertisers(active_only=True)
        check("  no warning", warning is None, warning)
        equal("  both rows are active by default", len(active_rows), 2)

        print("set_advertiser_name")
        renamed, rename_error = db.set_advertiser_name(row3["id"], "  Mattress Warehouse Inc  ")
        check("  rename succeeds", rename_error is None, rename_error)
        equal("  canonical_name updated (whitespace collapsed)",
             renamed.get("canonical_name"), "Mattress Warehouse Inc")
        equal("  name_key re-derived", renamed.get("name_key"), "mattress warehouse inc")

        collide_row, collide_error = db.set_advertiser_name(row3["id"], "Cardinal Plumbing")
        check("  renaming onto an existing ACTIVE name_key is refused",
             collide_row is None and collide_error, collide_error)
        check("  the refusal names the fix", "merge" in (collide_error or "").lower(), collide_error)

        same_row, same_error = db.set_advertiser_name(row3["id"], "Mattress Warehouse Inc")
        check("  renaming a row onto its OWN current name is not a collision with itself",
             same_error is None, same_error)

        print("advertiser_usage_counts")
        fake.table("attribution_reports").rows.append(
            {"id": "extra-report", "advertiser_id": row3["id"], "proposal_id": None,
             "report_json": {}, "status": "parsed"})
        counts, counts_error = db.advertiser_usage_counts(row3["id"])
        check("  no error", counts_error is None, counts_error)
        equal("  proposals count", counts.get("proposals"), 0)
        equal("  attribution_reports count", counts.get("attribution_reports"), 2)
        equal("  case_studies count (none tagged yet)", counts.get("case_studies"), 0)

        print("merge_advertisers")
        source_id, target_id = row3["id"], row1["id"]
        # A real merge that actually moves something in all THREE tables --
        # the live 2026-09-15 WAEPA TEST -> WAEPA merge moved nothing (WAEPA
        # TEST had no proposals/reports/case studies attached), so that
        # merge alone never exercised the repoint or the "deactivated but
        # still resolvable" guarantee. Manufacturing a second live merge
        # just to prove it was explicitly ruled out; this fake-client test
        # is where that proof belongs instead.
        fake.table("proposals").rows.append(
            {"id": "prop-source-only", "client_name": "Mattress Warehouse Inc",
             "advertiser_id": source_id})
        fake.table("case_studies").rows.append(
            {"id": "cs-source-only", "title": "Mattress Warehouse case study",
             "advertiser_id": source_id})
        pre_merge_target_reports = len([r for r in fake.table("attribution_reports").rows
                                        if r["advertiser_id"] == target_id])

        ok, merge_error = db.merge_advertisers(source_id, target_id)
        check("  merge succeeds", ok and merge_error is None, merge_error)

        equal("  the proposal that pointed at the source now points at the target",
             next(r for r in fake.table("proposals").rows
                 if r["id"] == "prop-source-only")["advertiser_id"], target_id)
        equal("  the case study that pointed at the source now points at the target",
             next(r for r in fake.table("case_studies").rows
                 if r["id"] == "cs-source-only")["advertiser_id"], target_id)
        merged_reports = [r for r in fake.table("attribution_reports").rows
                          if r["advertiser_id"] == target_id]
        equal("  every report that pointed at the source now points at the target too",
             len(merged_reports), pre_merge_target_reports + 2)
        check("  nothing in any of the three tables still points at the source id",
             not any(r.get("advertiser_id") == source_id
                    for table_name in ("proposals", "attribution_reports", "case_studies")
                    for r in fake.table(table_name).rows),
             True)

        source_after = next(r for r in fake.table("advertisers").rows if r["id"] == source_id)
        equal("  the source row is deactivated, not deleted", source_after.get("active"), False)

        active_after, _ = db.fetch_advertisers(active_only=True)
        check("  the merged-away row no longer shows up in an active-only fetch",
             all(r["id"] != source_id for r in active_after), True)
        # The "deactivated but still resolvable by id" guarantee, proven
        # through the real db.py function -- not by reading the fake table
        # directly -- since that's what a rebuild/render actually calls.
        all_after, all_warning = db.fetch_advertisers(active_only=False)
        check("  db.fetch_advertisers(active_only=False) has no warning", all_warning is None, all_warning)
        resolved_source = next((r for r in all_after if r["id"] == source_id), None)
        check("  the deactivated advertiser still resolves by id through the real function",
             resolved_source is not None, all_after)
        check("  and its own canonical_name is still intact (not blanked)",
             resolved_source is not None and resolved_source.get("canonical_name"), resolved_source)

        print("upload_report_file / report_file (prepare_deck_for_upload stubbed)")
        real_prepare = db.prepare_deck_for_upload
        db.prepare_deck_for_upload = lambda local_path, limit=None: (local_path, {"note": "stub"}, None)
        try:
            tmp_pptx = Path(__file__).resolve().parent / "_fake_report_for_upload.pptx"
            tmp_pptx.write_bytes(b"not a real pptx, just bytes for the fake storage layer")
            try:
                updated_row, stats, upload_error = db.upload_report_file(rid1, str(tmp_pptx), "report.pptx")
                check("  upload succeeds", upload_error is None, upload_error)
                equal("  storage_path recorded on the row",
                     updated_row.get("storage_path"), f"{rid1}/report.pptx")
                stored_row = next(r for r in fake.table("attribution_reports").rows if r["id"] == rid1)
                equal("  and on the underlying table row too",
                     stored_row.get("storage_path"), f"{rid1}/report.pptx")

                local_path = db.report_file(rid1, updated_row["storage_path"])
                check("  report_file returns a real local path",
                     Path(local_path).exists(), local_path)
                equal("  round-tripped bytes match what was uploaded",
                     Path(local_path).read_bytes(), tmp_pptx.read_bytes())
            finally:
                tmp_pptx.unlink(missing_ok=True)
        finally:
            db.prepare_deck_for_upload = real_prepare

        print("delete_attribution_report")
        fake.table("case_studies").insert(
            {"id": "cs-1", "title": "Uses the report", "source_report_id": rid1}).execute()
        blocked_ok, blocked_error = db.delete_attribution_report(rid1)
        check("  refused while a case study references it",
             not blocked_ok and blocked_error, blocked_error)
        check("  the refusal names the case study",
             "Uses the report" in (blocked_error or ""), blocked_error)
        still_there = any(r["id"] == rid1 for r in fake.table("attribution_reports").rows)
        check("  the report row is untouched", still_there, still_there)

        deletable_ok, deletable_error = db.delete_attribution_report(rid2)
        check("  a report nothing references deletes cleanly",
             deletable_ok and deletable_error is None, deletable_error)
        still_gone = any(r["id"] == rid2 for r in fake.table("attribution_reports").rows)
        check("  the row is actually gone", not still_gone, still_gone)

        print("Unreachable Supabase -- None, never an exception")
        db.get_client = lambda: None
        rows, warning = db.fetch_advertisers()
        check("  fetch_advertisers returns (None, warning)", rows is None and warning, warning)
        row, error = db.create_advertiser("Anyone")
        check("  create_advertiser returns (None, warning)", row is None and error, error)
        rid, error = db.log_attribution_report(None, None, {})
        check("  log_attribution_report returns (None, warning)", rid is None and error, error)
    finally:
        db.get_client = real_get_client

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
