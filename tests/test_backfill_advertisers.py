"""backfill_advertisers.py against a stubbed db module -- never real rows,
same rule test_feedback.py/test_draft_regression.py already follow for a
script that can write to production. Covers: grouping by normalized name
(so spelling variants join one advertiser), the placeholder-name carve-out
("Client" is reported, never turned into a real advertiser), an
already-linked proposal is left untouched, dry-run writes nothing, and a
second --write run is a no-op for everything the first one reached.

    python tests/test_backfill_advertisers.py
"""
import contextlib
import io
import sys
import uuid
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import db  # noqa: E402
import advertiser_matching  # noqa: E402
import backfill_advertisers  # noqa: E402


failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def equal(label, actual, expected):
    check(label, actual == expected, f"expected {expected!r}, got {actual!r}")


class FakeStore:
    """A plain in-memory stand-in for the db.py functions backfill_
    advertisers.py calls -- no fake Supabase query layer needed here, since
    this script only ever calls db's already-tested public functions, never
    builds its own queries."""

    def __init__(self):
        self.proposals = []
        self.advertisers = []

    def fetch_proposals(self, limit=500):
        return list(self.proposals), None

    def fetch_advertisers(self, limit=1000, active_only=False):
        rows = [r for r in self.advertisers if (not active_only or r.get("active", True))]
        return rows, None

    def create_advertiser(self, name):
        display = " ".join((name or "").split())
        if not display:
            return None, "Enter an advertiser name first."
        key = db.advertiser_name_key(display)
        for row in self.advertisers:
            if row["name_key"] == key:
                return row, None
        row = {"id": str(uuid.uuid4()), "canonical_name": display, "name_key": key, "active": True}
        self.advertisers.append(row)
        return row, None

    def link_proposal_advertiser(self, proposal_id, advertiser_id):
        for row in self.proposals:
            if row["id"] == proposal_id:
                row["advertiser_id"] = advertiser_id
                return True, None
        return False, "proposal not found"

    def is_configured(self):
        return True


def make_proposal(client_name, advertiser_id=None, generated_at="2026-01-01T00:00:00"):
    return {"id": str(uuid.uuid4()), "client_name": client_name,
            "advertiser_id": advertiser_id, "generated_at": generated_at}


def run_main(store, argv):
    """Runs backfill_advertisers.main(argv) against `store`, with db's own
    functions monkeypatched onto it for the duration. Returns the captured
    stdout (the script is print-heavy; tests assert on state, but capturing
    keeps the test's own PASS/FAIL lines legible)."""
    patched = ("fetch_proposals", "fetch_advertisers", "create_advertiser",
              "link_proposal_advertiser", "is_configured")
    saved = {name: getattr(db, name) for name in patched}
    for name in patched:
        setattr(db, name, getattr(store, name))
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exit_code = backfill_advertisers.main(argv)
    finally:
        for name in patched:
            setattr(db, name, saved[name])
    return exit_code, buf.getvalue()


def main():
    print("group_proposals")
    p_cardinal_1 = make_proposal("Cardinal Plumbing")
    p_cardinal_2 = make_proposal("Cardinal Plumbing, LLC")
    p_mw = make_proposal("Mattress Warehouse")
    p_client_1 = make_proposal("Client")
    p_client_2 = make_proposal("Client")
    p_blank = make_proposal("   ")
    rows = [p_cardinal_1, p_cardinal_2, p_mw, p_client_1, p_client_2, p_blank]
    groups, unplaceable = backfill_advertisers.group_proposals(rows)

    cardinal_key = advertiser_matching.normalize_name("Cardinal Plumbing")
    check("  'Cardinal Plumbing' and 'Cardinal Plumbing, LLC' collapse into one group",
         groups.get(cardinal_key) == [p_cardinal_1, p_cardinal_2], groups.get(cardinal_key))
    mw_key = advertiser_matching.normalize_name("Mattress Warehouse")
    equal("  Mattress Warehouse is its own group", groups.get(mw_key), [p_mw])
    equal("  exactly two real groups", len(groups), 2)
    equal("  'Client' x2 and the blank name are all unplaceable, never a group",
         {row["id"] for row in unplaceable}, {p_client_1["id"], p_client_2["id"], p_blank["id"]})

    print("dry run writes nothing")
    store = FakeStore()
    store.proposals = [make_proposal("Cardinal Plumbing"), make_proposal("Cardinal Plumbing LLC"),
                       make_proposal("WAEPA", advertiser_id="already-linked")]
    exit_code, output = run_main(store, [])
    equal("  exit code 0", exit_code, 0)
    equal("  no advertisers created", store.advertisers, [])
    check("  no proposal.advertiser_id written",
         all(row["advertiser_id"] in (None, "already-linked") for row in store.proposals),
         store.proposals)
    check("  output says nothing was written",
         "nothing written" in output.lower() or "dry run" in output.lower(), output[:200])

    print("--write actually links, and skips an already-linked proposal")
    store2 = FakeStore()
    p1 = make_proposal("Cardinal Plumbing")
    p2 = make_proposal("Cardinal Plumbing LLC")
    p_already = make_proposal("WAEPA", advertiser_id="pre-existing-id")
    store2.proposals = [p1, p2, p_already]
    exit_code, output = run_main(store2, ["--write"])
    equal("  exit code 0", exit_code, 0)
    equal("  exactly one advertiser created for the Cardinal variants", len(store2.advertisers), 1)
    equal("  both Cardinal proposals now share that advertiser's id",
         p1["advertiser_id"], store2.advertisers[0]["id"])
    equal("  and so does the second one", p2["advertiser_id"], store2.advertisers[0]["id"])
    equal("  the already-linked proposal (WAEPA) was never touched",
         p_already["advertiser_id"], "pre-existing-id")

    print("a second --write run is a no-op for what the first one reached")
    exit_code2, output2 = run_main(store2, ["--write"])
    equal("  exit code 0", exit_code2, 0)
    equal("  still exactly one advertiser (no duplicate created)", len(store2.advertisers), 1)
    check("  '0 to place' -- nothing left for this run to do",
         "0 to place" in output2, output2[:300])

    print("a new, still-unlinked proposal is picked up by a later run")
    p3 = make_proposal("Cardinal Plumbing Inc")
    store2.proposals.append(p3)
    exit_code3, output3 = run_main(store2, ["--write"])
    equal("  exit code 0", exit_code3, 0)
    equal("  it joins the SAME existing advertiser, not a new one", len(store2.advertisers), 1)
    equal("  it's now linked", p3["advertiser_id"], store2.advertisers[0]["id"])

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
