"""QA-015: a second draft never loses an avails group the first draft had.

Muse's exact repro (build 654ea56, "QA-TEST-Lawn & Leisure", the Lawn &
Leisure avails PDF): upload the avail, Draft #1 on clean notes, then replace
the notes and Draft #2 with no Clarify answers in between. Draft #2 removed
the avails group "DEMO Homeowner, HH Income 200K Plus" (254,934 monthly
avails) from the plan and an invented "HH Income 100K Plus" line (0 avails)
took its place, while the checklist claimed the 200K group "is not on the
avails table" -- the QA-014 false absence, which likely led the model to
drop the group and invent one.

Two layers checked here:
- the false-absence note can't come back (QA-014's term matching);
- even if a model's selection drops the group and invents a lookalike, the
  group is kept, the invented line re-pointed at it (budget kept), and a
  review note says so -- unless the notes explicitly rule the group out.

Draft #1 is a real recorded model response for this document
(tests/fixtures/qa015/draft1_live.json); Draft #2's notes are Muse's,
verbatim (notes_draft2.txt). Draft #2's response is built to reproduce the
swap Muse saw -- today's model, given today's prompt, kept the group in a
live check, so the guard is tested against the failure itself.

    python tests/test_redraft_keeps_avail_group.py
"""
import copy
import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import group_scenario_fixtures as gsf          # noqa: E402

FIX = REPO / "tests" / "fixtures" / "qa015"
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    def __init__(self, state):
        self.session_state = state
        self.secrets = {}


def apply(state, draft, notes):
    real = app.st
    app.st = _StubSt(state)
    try:
        app.apply_draft_to_form(copy.deepcopy(draft), notes=notes)
        app.apply_pending_draft_fields()
    finally:
        app.st = real
    return state


def on_plan(state, gid):
    return any(gid in app.group_ids_of(r) for o in state["plan_options"] for r in o["rows"])


def swap_draft(base):
    """Draft #2 as Muse saw it: the avails group left out of the selection,
    the budget put on an invented 100K line instead."""
    d = copy.deepcopy(base)
    d["group_selection"] = {"mode": "match", "match": ["HH Income 100K Plus"]}
    d["audiences"] = [{"segment": "HH Income 100K Plus", "geo": "10 Mile Radius Zips"}]
    d["media_plan_lines"] = [{"product": "premion_streaming_tv", "label": "Homeowners",
                              "audience_track": "HH Income 100K Plus",
                              "allocation": {"percent_of_total": 100}}]
    d["unresolved_internal"] = ["HH Income 200K Plus is not on the avails table -- add it in D2."]
    return d


def main():
    if not gsf.LAWN_LEISURE_PDF.exists():
        print(f"  SKIP  {gsf.LAWN_LEISURE_PDF.name} not present")
        return 0
    scn = gsf.build_lawn_leisure()
    groups = [dict(g, include_in_plan=False, include_locked=False) for g in scn["groups"]]
    for g in groups:
        g.pop("_flight_label", None)
    gid = groups[0]["id"]
    state = {"targeting_groups": groups, "flight_start": scn["flight_start"],
             "flight_end": scn["flight_end"], "market_choice": "DC", "avails_mode": True,
             "client_name": "QA-TEST-Lawn & Leisure"}
    draft1 = json.loads((FIX / "draft1_live.json").read_text(encoding="utf-8"))
    notes2 = (FIX / "notes_draft2.txt").read_text(encoding="utf-8")

    print("Draft #1 (real recorded response) puts the avails group on the plan")
    apply(state, draft1, "Lawn & Leisure, higher-income homeowners within 10 miles, ~$3,500/mo.")
    check("the 254,934-avails group is on the plan after draft #1", on_plan(state, gid))

    print("\nDraft #2 with Muse's notes, the model dropping the group for an invented lookalike line")
    apply(state, swap_draft(draft1), notes2)
    rows = [r for o in state["plan_options"] for r in o["rows"]]
    internal = state.get("draft_unresolved_internal") or []
    check("the avails group is still on the plan, its budget with it",
          any(gid in app.group_ids_of(r) and float(r.get("Cost") or 0) > 0 for r in rows),
          [(r.get("Targeting"), r.get("Cost"), app.group_ids_of(r)) for r in rows])
    check("no plan line targets the invented 'HH Income 100K Plus'",
          not any("100K" in str(r.get("Targeting")) for r in rows), [r.get("Targeting") for r in rows])
    check("no zero-avail '100K Plus' group is left on the avails table",
          not any("100K" in " ".join(g.get("terms") or []) for g in state["targeting_groups"]),
          [g.get("terms") for g in state["targeting_groups"]])
    check("the false 'HH Income 200K Plus is not on the avails table' note is gone (QA-014)",
          not any("200K" in n and "not on the avails table" in n for n in internal), internal)

    print("\nDraft #2 whose plan genuinely leaves the group out (budget moved to another product)")
    state_b = {"targeting_groups": [dict(g, include_in_plan=False) for g in groups],
               "flight_start": scn["flight_start"], "flight_end": scn["flight_end"],
               "market_choice": "DC", "avails_mode": True, "client_name": "QA-TEST-Lawn & Leisure"}
    apply(state_b, draft1, "higher-income homeowners")
    lost = copy.deepcopy(draft1)
    lost["group_selection"] = {"mode": "match", "match": ["HH Income 100K Plus"]}
    lost["audiences"] = [{"segment": "HH Income 100K Plus", "geo": "10 Mile Radius Zips"}]
    lost["media_plan_lines"] = [{"product": "streaming_retargeting_display",
                                 "allocation": {"percent_of_total": 100}}]
    apply(state_b, lost, notes2)
    internal_b = state_b.get("draft_unresolved_internal") or []
    check("the avails group is put back on the plan",
          any(g["id"] == gid and g.get("include_in_plan") for g in state_b["targeting_groups"]),
          [(g["id"], g.get("include_in_plan")) for g in state_b["targeting_groups"]])
    check("a review note says it was put back and why",
          any("back on the plan" in n and "200K" in n for n in internal_b), internal_b)

    print("\nThe guard itself (keep_avail_groups_on_plan)")
    group = state["targeting_groups"][0]
    prev = [{"name": "A", "rows": [{"Tactic": "Premion Streaming TV", "_group_ids": [gid],
                                    "Targeting": "DEMO Homeowner, HH Income 200K Plus", "Cost": 3500}]}]
    invented = [{"name": "A", "rows": [{"Tactic": "Premion Streaming TV", "Targeting": "HH Income 100K Plus",
                                        "Geo": "Washington, DC", "Cost": 3500}]}]
    off = [dict(group, include_in_plan=False)]
    g2, o2, n2 = app.keep_avail_groups_on_plan([group], prev, off, invented, notes2)
    row = o2[0]["rows"][0]
    check("an invented zero-avail lookalike line is re-pointed at the avails group, $3,500 kept",
          app.group_ids_of(row) == [gid] and "200K" in row["Targeting"] and row["Cost"] == 3500, row)
    check("...the group is ticked onto the plan and a note names both audiences",
          g2[0].get("include_in_plan") and n2 and "Kept the avails audience" in n2[0]
          and "HH Income 100K Plus" in n2[0], (g2, n2))
    _g3, o3, n3 = app.keep_avail_groups_on_plan(
        [group], prev, off, invented, notes2 + " Actually drop the HH Income 200K Plus row, go broader.")
    check("notes that explicitly rule the group out are respected (nothing changed, no note)",
          not n3 and app.group_ids_of(o3[0]["rows"][0]) == [], (n3, o3))
    unrelated = [{"name": "A", "rows": [{"Tactic": "Premion Streaming TV", "Targeting": "AUTO Ford Intenders",
                                         "Cost": 3500}]}]
    _g4, o4, n4 = app.keep_avail_groups_on_plan([group], prev, off, unrelated, notes2)
    check("an unrelated new line is left alone; the group is put back beside it instead",
          o4[0]["rows"][0]["Targeting"] == "AUTO Ford Intenders" and n4 and "back on the plan" in n4[0],
          (o4, n4))

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
