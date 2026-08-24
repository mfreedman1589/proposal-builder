"""Retargeting/Audience Marketplace/sports/fees seed ONE line for the whole
campaign, never one per targeting group. They inherited the per-group
fan-out only because `_seed_option_rows` used to share `plan_lines` (the
per-group geo_default) with Premion Streaming TV; once Premion is owned
entirely by `reconcile_group_plan_lines`, there is no remaining reason for
the other products to fan out at all -- a 12-group avails table with
Retargeting on used to produce 12 Retargeting rows regardless of how many
of those 12 groups were actually selected into the CTV plan, directly
contradicting "the avails table is research."

    python tests/test_product_fanout_fix.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def run_form(state):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "Fan-out Suite"
    for key, value in state.items():
        at.session_state[key] = value
    at.run()
    return at


def main():
    COL = app.AVAILS_COLUMN_MONTHLY
    rows = [{"Audience": a, "Geo": "Denver", COL: 100000} for a in
            ["Homeowners", "In-market for windows", "Pet owners"]]
    groups = app.tg.seed_rows_to_groups(rows, COL)
    for g in groups:
        g["include_in_plan"] = False  # avails only -- nothing sold yet

    print("=" * 78)
    print("3 groups, Retargeting Display on, Premion off -> ONE Retargeting line")
    print("=" * 78)
    at = run_form({
        "include_avails_template": True,
        "avails_seed_rows": rows,
        "targeting_groups": [dict(g) for g in groups],
        "premion_streaming_tv": False,
        "streaming_retargeting_enabled": True,
        "sr_disp": True,
    })
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    plan_rows = at.session_state["plan_options"][0]["rows"]
    retarg = [r for r in plan_rows if r.get("Tactic", "").startswith("Streaming Retargeting")]
    check("exactly one Retargeting line, not three", len(retarg) == 1, retarg)
    check("it carries no _group_ids -- ungrouped, campaign-wide",
          not app.group_ids_of(retarg[0]) if retarg else False, retarg)

    print("\n" + "=" * 78)
    print("Same 3 groups, only ONE ticked into the CTV plan -- Retargeting still one line")
    print("=" * 78)
    ticked = [dict(g, include_in_plan=(g["terms"] == ["Homeowners"])) for g in groups]
    at2 = run_form({
        "include_avails_template": True,
        "avails_seed_rows": rows,
        "targeting_groups": ticked,
        "premion_streaming_tv": True,
        "streaming_retargeting_enabled": True,
        "sr_disp": True,
    })
    check("no exception", not at2.exception, at2.exception[0].message[:400] if at2.exception else "")
    plan_rows2 = at2.session_state["plan_options"][0]["rows"]
    retarg2 = [r for r in plan_rows2 if r.get("Tactic", "").startswith("Streaming Retargeting")]
    premion2 = [r for r in plan_rows2 if r.get("Tactic") == "Premion Streaming TV"]
    check("still exactly one Retargeting line (unaffected by CTV selection)",
          len(retarg2) == 1, retarg2)
    check("exactly one Premion line -- the one ticked group, and only it",
          len(premion2) == 1 and premion2[0]["Targeting"] == "Homeowners", premion2)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Retargeting/AM/sports/fees seed one line for the whole campaign, independent of "
          "how many targeting groups exist or how many are ticked into the CTV plan -- Premion "
          "Streaming TV alone fans out per selected group, via its own separate ownership.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
