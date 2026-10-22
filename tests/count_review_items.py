"""What share of the plan-changing follow-ups a draft raises are cards.

The measure (Matt, 2026-10-22): of the open questions whose answer would
change the plan -- a split, a budget, a rate, a segment, an avails figure,
a market or geography, a line -- how many reach the rep as a card they can
answer in one click, rather than a sentence to act on by hand. Target ~100%.
Open questions with no plan change behind them (creative, measurement, the
internal vertical, flight dates set in the setup band, the client name, a
product the notes didn't ask for) are FYI by nature and don't count.

Every saved draft is run through the real apply_draft_to_form (offline, no
API call). Sources, both optional:
- the frozen Tier 1 responses (tests/fixtures/*.draft.json), and
- Tier 2 responses saved with `python tests/test_draft_live.py --save`
  (tests/fixtures/*.live.json, gitignored), run against the band the live
  scenario used.

Prints the percentage and every plan-changing note still in prose, so the
next migration is a list, not a guess.

    python tests/count_review_items.py
"""
import collections
import copy
import re
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

import db  # noqa: E402,F401 -- the real catalog (read-only): a fallback catalog
# would report real segments as "unrecognized" and inflate the prose count.

import app  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures"


class _StubSt:
    def __init__(self, state):
        self.session_state = dict(state)
        self.secrets = {}


# What a plan-changing note is about -- the plan changes the app can apply.
PLAN_TOPICS = re.compile(
    r"\b(split|weight\w*|divided|evenly|allocation|budget|cpm|rate card|negotiated|"
    r"segment|catalog|proxy|closest|stand-?in|custom audience|avails|market|dma|zip|"
    r"geograph\w*|county|counties|radius|line)\b", re.I)
# ...and what isn't, however it's worded: no plan change behind it.
NOT_PLAN = re.compile(
    r"\b(creative|spot|production|video|attribution|measurement|brand lift|pixel|tracking|"
    r"report\w*|vertical|category|flight|dates?|start|end date|client name|retargeting|"
    r"call tracking)\b", re.I)


def plan_changing(note):
    return bool(PLAN_TOPICS.search(note)) and not NOT_PLAN.search(note)


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
    cards, prose_plan, prose_other = 0, [], 0
    by_kind = collections.Counter()
    drafts = collections.Counter()
    for tier, name, draft, notes, band in sources():
        state = run_one(draft, notes, band)
        drafts[tier] += 1
        for item in state.get("draft_review_items") or []:
            if item["type"] != "fyi":
                cards += 1
                by_kind[kind_of(item)] += 1
        for note in (state.get("draft_unresolved") or []) + (state.get("draft_unresolved_internal")
                                                              or []):
            if plan_changing(note):
                prose_plan.append((tier, name, note))
            else:
                prose_other += 1

    total = cards + len(prose_plan)
    pct = 100.0 * cards / total if total else 100.0
    print(f"Drafts: {drafts['tier1']} Tier 1 (frozen), {drafts['tier2']} Tier 2 (saved live)\n")
    print(f"Plan-changing follow-ups that are cards: {cards} of {total} ({pct:.0f}%)")
    print(f"(Not counted: {prose_other} FYI notes with no plan change behind them.)\n")
    print("Cards by what they're about:")
    for kind, count in by_kind.most_common():
        print(f"  {count:>3}  {kind}")
    print(f"\nPlan-changing notes still in prose ({len(prose_plan)}):")
    for tier, name, note in prose_plan:
        print(f"  [{tier} {name}] {note}")


if __name__ == "__main__":
    main()
