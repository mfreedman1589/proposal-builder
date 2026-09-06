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
                self.table.rows.append(row)
                inserted.append(row)
            return _FakeResult(inserted)
        if self.op == "update":
            matched = self._filtered()
            for row in matched:
                row.update(self.payload)
            return _FakeResult(matched)
        raise NotImplementedError(self.op)


class _FakeTable:
    def __init__(self):
        self.rows = []

    def select(self, *args, count=None):
        query = _FakeQuery(self, "select")
        query.count_mode = count
        return query

    def insert(self, payload):
        return _FakeQuery(self, "insert", payload)

    def update(self, payload):
        return _FakeQuery(self, "update", payload)


class FakeSupabaseClient:
    def __init__(self):
        self._tables = {}

    def table(self, name):
        return self._tables.setdefault(name, _FakeTable())


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
