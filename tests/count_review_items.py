"""Count draft follow-ups by type, from saved draft responses.

Runs every saved draft through the real apply_draft_to_form (offline, no API
call) and tallies what a rep would see: structured cards by type, priority,
audience and where they came from, plus the prose notes that fit no card.

Sources, both optional:
- the frozen Tier 1 responses (tests/fixtures/*.draft.json), and
- Tier 2 responses saved with `python tests/test_draft_live.py --save`
  (tests/fixtures/*.live.json, gitignored), run against the same band the
  live scenario used.

    python tests/count_review_items.py
"""
import collections
import copy
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures"


class _StubSt:
    def __init__(self, state):
        self.session_state = dict(state)
        self.secrets = {}


def kind_of(item):
    """A short name for what an item is about, for the tally."""
    if item.get("source") == "model":
        ops = {op.get("op") for o in item["options"] for op in o.get("patch") or []}
        return "model: " + ", ".join(sorted(ops))
    if item.get("source") == "targeting_check":
        return item.get("check", "targeting_check")
    ops = {op.get("op") for o in item["options"] for op in o.get("patch") or []}
    keys = {op.get("key") for o in item["options"] for op in o.get("patch") or []}
    if "gdw_added_value" in ops:
        return "GDW added value"
    if "total_tv" in keys:
        return "Total TV"
    if "agency_gross_up" in keys:
        return "gross-up"
    if "avails_include_duplicates" in keys:
        return "duplicate avails"
    if "set_plan_groups" in ops:
        return "avails audience selection"
    if "include_groups" in ops:
        return "available, not on the plan"
    if "set_segments" in ops:
        return "stand-in segment swap" if item.get("priority") == "suggestion" else "2-segment pick"
    return ", ".join(sorted(ops)) or "other"


def run_one(draft, notes, band):
    real = app.st
    stub = _StubSt(band)
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft), notes=notes)
    finally:
        app.st = real
    return stub.session_state


def sources():
    live = {}
    try:
        import test_draft_live
        live = test_draft_live.SCENARIOS
    except Exception as exc:  # the live module needs nothing but imports
        print(f"(live scenarios unavailable: {exc})")
    for path in sorted(FIXTURES.glob("*.draft.json")):
        name = path.name[:-len(".draft.json")]
        notes = FIXTURES / f"{name}.txt"
        yield ("tier1", name, json.loads(path.read_text(encoding="utf-8")),
               notes.read_text(encoding="utf-8") if notes.exists() else "",
               {"market_choice": "DC", "flight_start": date(2026, 9, 1),
                "flight_end": date(2026, 11, 30), "avails_mode": False})
    for path in sorted(FIXTURES.glob("*.live.json")):
        name = path.name[:-len(".live.json")]
        spec = live.get(name)
        if not spec:
            continue
        band = {"market_choice": spec["market"], "flight_start": spec["flight_start"],
                "flight_end": spec["flight_end"], "avails_mode": True}
        if spec.get("existing_groups"):
            band["targeting_groups"] = spec["existing_groups"]()
        yield ("tier2", name, json.loads(path.read_text(encoding="utf-8")),
               (FIXTURES / f"{name}.txt").read_text(encoding="utf-8"), band)


def main():
    tally = collections.Counter()
    by_kind = collections.Counter()
    drafts = collections.Counter()
    per_draft = []
    for tier, name, draft, notes, band in sources():
        state = run_one(draft, notes, band)
        items = state.get("draft_review_items") or []
        client = state.get("draft_unresolved") or []
        internal = state.get("draft_unresolved_internal") or []
        drafts[tier] += 1
        for item in items:
            tally[(item["type"], item["priority"], item["audience"])] += 1
            by_kind[kind_of(item)] += 1
        tally[("fyi (prose)", "fyi", "client")] += len(client)
        tally[("fyi (prose)", "fyi", "internal")] += len(internal)
        per_draft.append((tier, name, len(items), len(client), len(internal)))

    print(f"Drafts: {drafts['tier1']} Tier 1 (frozen), {drafts['tier2']} Tier 2 (saved live)\n")
    print(f"{'draft':38} {'cards':>5} {'prose: client':>14} {'prose: internal':>16}")
    for tier, name, n_items, n_client, n_internal in per_draft:
        print(f"{tier + ' ' + name:38} {n_items:>5} {n_client:>14} {n_internal:>16}")
    print(f"\n{'type':14} {'priority':10} {'audience':9} {'count':>5}")
    for (type_, priority, audience), count in sorted(tally.items()):
        print(f"{type_:14} {priority:10} {audience:9} {count:>5}")
    print("\nCards by what they're about:")
    for kind, count in by_kind.most_common():
        print(f"  {count:>3}  {kind}")


if __name__ == "__main__":
    main()
