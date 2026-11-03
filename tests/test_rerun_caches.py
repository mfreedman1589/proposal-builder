"""The three caches that keep a Build-page rerun cheap, and the rule each one
has to keep: the answer is exactly what it was uncached.

    python tests/test_rerun_caches.py

* db's read cache (case studies, vault slides, team list): a read is reused
  for READ_CACHE_SECONDS, a write through db clears its table at once, a
  failed read is never stored, and only the real client is cached -- a fake
  client a test installs is read every time.
* segment_recommend.collapse_in_order: cached on the exact list, a fresh copy
  each call.
* report_assembly._nearest_place: the latitude-ordered search picks the same
  place at the same distance as scanning every place.

Measured 2026-11-03: the finder's duplicate check was ~3.5s of a ~5.5s
Build-page rerun, and zip_place_name ~1s of a report-deck build. Everything
here is stubbed -- no real Supabase table is touched.
"""
import os
import random
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
import geo_resolver  # noqa: E402
import report_assembly  # noqa: E402
import segment_recommend  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {detail}" if detail else ""))


class FakeQuery:
    def __init__(self, store, table):
        self.store, self.table = store, table
        self.op = "select"
        self.payload = None

    def select(self, *a, **k):
        return self

    def eq(self, *a, **k):
        return self

    def order(self, *a, **k):
        return self

    def limit(self, *a, **k):
        return self

    def insert(self, payload):
        self.op, self.payload = "insert", payload
        return self

    def execute(self):
        if self.store.fail:
            raise RuntimeError("backend unreachable")
        rows = self.store.tables.setdefault(self.table, [])
        if self.op == "insert":
            rows.extend(self.payload if isinstance(self.payload, list) else [self.payload])
            return type("R", (), {"data": [self.payload]})()
        self.store.reads += 1
        return type("R", (), {"data": [dict(r) for r in rows]})()


class FakeClient:
    def __init__(self):
        self.tables, self.reads, self.fail = {}, 0, False

    def table(self, name):
        return FakeQuery(self, name)


def check_read_cache():
    print("\ndb read cache")
    original_get_client = db.get_client
    real = FakeClient()
    real.tables["case_studies"] = [{"id": 1, "title": "A"}]
    real.tables["team_members"] = [{"name": "Ann", "active": True}]
    db._clients[("test-url", "test-key")] = real
    db.get_client = lambda: real
    db.forget_reads()
    try:
        first, _ = db.fetch_case_studies()
        second, _ = db.fetch_case_studies()
        check("a second read within the window reuses the first", real.reads == 1, real.reads)
        check("and returns the same rows", first == second == [{"id": 1, "title": "A"}], second)
        second[0]["title"] = "changed by a caller"
        third, _ = db.fetch_case_studies()
        check("a caller changing its copy can't change the stored one",
              third[0]["title"] == "A", third)

        db.fetch_case_studies(active_only=False)
        check("active-only and all-rows reads are cached separately", real.reads == 2, real.reads)

        @db._clears_reads("case_studies")
        def write():
            real.tables["case_studies"].append({"id": 2, "title": "B"})
        write()
        after, _ = db.fetch_case_studies()
        check("a write through db shows on the very next read",
              [r["id"] for r in after] == [1, 2] and real.reads == 3, (after, real.reads))

        writes = ["upload_case_study", "update_case_study", "upload_case_study_images",
                  "upload_slide_vault", "update_slide_vault_entry", "delete_slide_vault_entry",
                  "upload_slide_vault_image", "add_team_member"]
        unwrapped = [name for name in writes if not hasattr(getattr(db, name), "__wrapped__")]
        check("every db write to a cached table clears it", not unwrapped, unwrapped)

        db.forget_reads()
        real.fail = True
        rows, warning = db.fetch_case_studies()
        check("a failed read is reported, as before", rows == [] and warning, warning)
        real.fail = False
        reads_before = real.reads
        rows, warning = db.fetch_case_studies()
        check("and is never stored: the next read goes to the backend",
              real.reads == reads_before + 1 and rows and warning is None, (real.reads, warning))

        db.READ_CACHE_SECONDS, saved = -1, db.READ_CACHE_SECONDS
        try:
            db.forget_reads()
            reads_before = real.reads
            db.fetch_case_studies()
            db.fetch_case_studies()
            check("an expired entry is read again", real.reads == reads_before + 2, real.reads)
        finally:
            db.READ_CACHE_SECONDS = saved

        fake = FakeClient()
        fake.tables["case_studies"] = [{"id": 9}]
        db.get_client = lambda: fake
        db.fetch_case_studies()
        db.fetch_case_studies()
        check("a client db didn't build (a test's fake) is never cached", fake.reads == 2, fake.reads)
    finally:
        db.get_client = original_get_client
        db._clients.pop(("test-url", "test-key"), None)
        db.forget_reads()


def check_collapse_cache():
    print("\nsegment_recommend.collapse_in_order")
    names = ["AUTO Body Style Sedan", "AUTO Body Style Sedans", "DEMO Homeowner",
             "LIFESTYLE Wine and Spiritsr", "LIFESTYLE Wine and Spirits"]
    first = segment_recommend.collapse_in_order(names)
    first[0][1].append("tampered")
    second = segment_recommend.collapse_in_order(names)
    check("a caller changing its result can't change the cached one",
          "tampered" not in second[0][1], second)
    check("near-duplicates and typos still fold under their first member",
          second == [("AUTO Body Style Sedan", ["AUTO Body Style Sedans"]),
                     ("DEMO Homeowner", []),
                     ("LIFESTYLE Wine and Spiritsr", ["LIFESTYLE Wine and Spirits"])], second)
    check("a different list is a different answer",
          segment_recommend.collapse_in_order(names[2:]) ==
          [("DEMO Homeowner", []), ("LIFESTYLE Wine and Spiritsr", ["LIFESTYLE Wine and Spirits"])])


def check_nearest_place():
    print("\nreport_assembly._nearest_place")
    points = geo_resolver._data().get("zip_points") or {}
    places = report_assembly._places_for_zip_lookup()
    if not points or not places:
        print("  SKIP  zip centroids or the places file aren't available")
        return

    def scan(point):
        best, best_dist = None, None
        for place in places:
            dist = geo_resolver.haversine_miles(point, (place["lat"], place["lon"]))
            if best_dist is None or dist < best_dist:
                best, best_dist = place, dist
        return best, best_dist

    zips = sorted(points)
    by_lat = sorted(zips, key=lambda z: points[z][0])
    by_lon = sorted(zips, key=lambda z: points[z][1])
    sample = random.Random(7).sample(zips, 40) + by_lat[:3] + by_lat[-3:] + by_lon[:3] + by_lon[-3:]
    wrong = [z for z in sample if scan(points[z]) != report_assembly._nearest_place(tuple(points[z]))]
    check(f"same place, same distance as a full scan ({len(sample)} zips, extremes included)",
          not wrong, wrong[:5])


def main():
    check_read_cache()
    check_collapse_cache()
    check_nearest_place()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
