"""backfill_advertiser_verticals.py against a stubbed db module -- never
real rows, same rule test_backfill_advertisers.py already follows.

    python tests/test_backfill_advertiser_verticals.py
"""
import contextlib
import io
import sys
import uuid
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import db  # noqa: E402
import backfill_advertiser_verticals  # noqa: E402


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
    def __init__(self):
        self.advertisers = []
        self.proposals = []

    def fetch_advertisers(self, limit=1000, active_only=False):
        rows = [r for r in self.advertisers if (not active_only or r.get("active", True))]
        return rows, None

    def fetch_proposals(self, limit=500):
        return list(self.proposals), None

    def derive_advertiser_vertical_from_proposals(self, advertiser_id):
        """A plain-list stand-in for db.py's real function -- that
        function's own correctness (the disagree/already_set/none-
        exclusion rules) is already covered against a real fake-Supabase-
        query-client in tests/test_advertisers_db.py; this file tests the
        SCRIPT's own orchestration (looping, dry-run vs --write, what gets
        printed), which only needs this to behave observably the same way,
        not to BE the same code."""
        advertiser = next((a for a in self.advertisers if a["id"] == advertiser_id), None)
        if advertiser is None:
            return None, "Advertiser not found."
        if advertiser.get("vertical"):
            return "already_set", None
        verticals = {p.get("vertical") for p in self.proposals
                    if p.get("advertiser_id") == advertiser_id and p.get("vertical")
                    and p.get("vertical") != "none"}
        if not verticals:
            return "no_proposals", None
        if len(verticals) > 1:
            return "disagree", None
        advertiser["vertical"] = next(iter(verticals))
        return "set", None

    def is_configured(self):
        return True


def make_advertiser(name, vertical=None):
    return {"id": str(uuid.uuid4()), "canonical_name": name, "vertical": vertical, "active": True}


def make_proposal(advertiser_id, vertical):
    return {"id": str(uuid.uuid4()), "advertiser_id": advertiser_id, "vertical": vertical}


def run_main(store, argv):
    patched = ("fetch_advertisers", "fetch_proposals",
              "derive_advertiser_vertical_from_proposals", "is_configured")
    saved = {name: getattr(db, name) for name in patched}
    for name in patched:
        setattr(db, name, getattr(store, name))
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exit_code = backfill_advertiser_verticals.main(argv)
    finally:
        for name in patched:
            setattr(db, name, saved[name])
    return exit_code, buf.getvalue()


def main():
    print("dry run: agree -> would set, disagree -> left alone, already-set -> untouched")
    store = FakeStore()
    agree = make_advertiser("Agrees")
    disagree = make_advertiser("Disagrees")
    already = make_advertiser("Already Set", vertical="banking")
    none_only = make_advertiser("Only None Picks")
    store.advertisers = [agree, disagree, already, none_only]
    store.proposals = [
        make_proposal(agree["id"], "retail"), make_proposal(agree["id"], "retail"),
        make_proposal(disagree["id"], "retail"), make_proposal(disagree["id"], "healthcare"),
        make_proposal(already["id"], "travel"),  # already has one -- must never be touched
        make_proposal(none_only["id"], "none"), make_proposal(none_only["id"], "none"),
    ]
    exit_code, output = run_main(store, [])
    equal("exit code 0", exit_code, 0)
    check("agree shows what it would derive to", "Agrees\" -> retail" in output, output)
    check("disagree is named as disagreeing", "Disagrees" in output and "disagree" in output.lower(), output)
    check("the 'none' sentinel is NOT signal -- reported as no real proposals",
         "Only None Picks" in output and "no linked proposals" in output, output)
    check("already-set advertiser isn't even listed (nothing to derive)",
         "Already Set" not in output, output)
    check("dry run writes nothing", "nothing written" in output.lower(), output)
    equal("no advertiser actually changed", [a.get("vertical") for a in store.advertisers],
         [None, None, "banking", None])

    print("\n--write actually derives")
    store2 = FakeStore()
    agree2 = make_advertiser("Agrees2")
    store2.advertisers = [agree2]
    store2.proposals = [make_proposal(agree2["id"], "healthcare"),
                        make_proposal(agree2["id"], "healthcare")]
    exit_code, output = run_main(store2, ["--write"])
    equal("exit code 0", exit_code, 0)
    equal("the vertical was actually written", agree2.get("vertical"), "healthcare")

    print("\na second --write run is a no-op (already set, never overwritten)")
    agree2["vertical"] = "healthcare"  # reflect the write from above for this second pass
    exit_code2, output2 = run_main(store2, ["--write"])
    equal("exit code 0", exit_code2, 0)
    check("no advertiser left to derive (already has one)",
         "0 would be set" in output2, output2)

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
