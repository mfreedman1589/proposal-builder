"""A drafted option with no Premion Streaming TV product of its own must
never get avails-table group rows synthesized into it, regardless of what
its own "group_selection" says -- found live (Easterns, 2026-09-23): a
sports-only "NFL Regular Season" option, drafted alongside a "CTV Plan"
option that legitimately sells from the avails table, kept getting two $0
Premion Streaming TV rows it had no product for.

Root cause: `match_groups_to_selection`'s own "empty/missing selection ->
sell everything" default is correct for a SINGLE-option plan (nothing else
it could mean) but wrong the moment a second, unrelated option exists on
the same draft and also says nothing -- it inherited the same "sell
everything" default purely because nothing else stopped it. `apply_draft_
to_form` now gates the whole group-selection resolution per option on
`has_streaming_tv_product` (a redrafted-in Premion line, a stated
group_allocation, or a stated group_cpm -- any real signal this option
prices Streaming TV lines at all): when none of those fired, group_
selection resolves to nothing for that option -- and, symmetrically, is
REJECTED even when the model states one anyway (it shouldn't, but nothing
stops it), rather than synthesizing rows into an option with no product to
attach them to. `match_groups_to_selection` also gained a real {"mode":
"none"} -- previously indistinguishable from an empty selection, which
always meant "all".

Runs entirely offline through `apply_draft_to_form` directly (the same
`_StubSt` pattern test_draft_group_ownership.py/test_draft_regression.py
use).

    python tests/test_group_selection_product_gate.py
"""
import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply_draft(draft, preset_state=None):
    real = app.st
    stub = _StubSt()
    stub.session_state.update(preset_state or {})
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return stub.session_state


def existing_groups():
    dc = tg.new_group(["AUTO Intenders"], geo_def={"kind": "markets", "markets": ["Washington, DC DMA"]},
                      avails_monthly=500000, group_id="dc-auto")
    balt = tg.new_group(["AUTO Intenders"], geo_def={"kind": "markets", "markets": ["Baltimore DMA"]},
                        avails_monthly=300000, group_id="balt-auto")
    for g in (dc, balt):
        g["resolved_zips"] = []
        g["resolved_markets"] = [tg.geo_label(g)]
    return [dc, balt]


def group_ids_in(rows):
    return sorted({gid for r in rows for gid in app.group_ids_of(r)})


def base_draft(nfl_selection):
    """CTV Plan legitimately sells the avails table (a real group_allocation);
    NFL Regular Season is sports-only and carries `nfl_selection` for its own
    "group_selection" -- the thing under test. `nfl_selection is None` omits
    the key entirely (drafted_options' own "opt.get(...) if key in opt else
    top-level" rule needs real absence, not a present-but-None value, to
    exercise the "entirely omitted" case honestly)."""
    nfl_option = {
        "name": "NFL Regular Season", "total_budget": 15000,
        "media_plan_lines": [
            {"product": "sport:nfl_reg", "label": "DC",
             "allocation": {"flat_amount": 10000}},
            {"product": "sport:nfl_reg", "label": "Baltimore",
             "allocation": {"flat_amount": 5000}},
        ],
    }
    if nfl_selection is not None:
        nfl_option["group_selection"] = nfl_selection
    return {
        "client_name": "Easterns", "vertical": "auto",
        "options": [
            {
                "name": "CTV Plan", "total_budget": 0,
                "group_allocation": {"percent_of_avails": 5},
                "group_cpm": 28,
                "media_plan_lines": [],
            },
            nfl_option,
        ],
    }


def main():
    print("A. NFL option's group_selection entirely omitted (the reported shape)")
    draft = base_draft(None)
    state = apply_draft(draft, {"targeting_groups": existing_groups()})
    opts = state["plan_options"]
    ctv = next(o for o in opts if o["name"] == "CTV Plan")
    nfl = next(o for o in opts if o["name"] == "NFL Regular Season")
    check("CTV Plan got both avails-table groups",
         group_ids_in(ctv["rows"]) == ["balt-auto", "dc-auto"], group_ids_in(ctv["rows"]))
    check("NFL Regular Season got NO group rows",
         group_ids_in(nfl["rows"]) == [], group_ids_in(nfl["rows"]))
    check("NFL Regular Season still has its own 2 sports rows",
         len(nfl["rows"]) == 2, [r.get("Tactic") for r in nfl["rows"]])
    check("a review note explains why nothing was added",
         any("doesn't include Premion Streaming TV" in n
             for n in state.get("draft_unresolved_internal", [])),
         state.get("draft_unresolved_internal"))

    print("\nB. NFL option's group_selection explicitly {\"mode\": \"all\"} anyway (the model shouldn't, but might)")
    draft = base_draft({"mode": "all"})
    state = apply_draft(draft, {"targeting_groups": existing_groups()})
    opts = state["plan_options"]
    nfl = next(o for o in opts if o["name"] == "NFL Regular Season")
    check("NFL Regular Season still got NO group rows despite an explicit selection",
         group_ids_in(nfl["rows"]) == [], group_ids_in(nfl["rows"]))
    check("the rejection note fires for an explicit-but-irrelevant selection too",
         any("named avails-table audiences to sell, but this option has no" in n
             for n in state.get("draft_unresolved_internal", [])),
         state.get("draft_unresolved_internal"))

    print("\nC. NFL option's group_selection explicitly {\"mode\": \"none\"} (the belt-and-braces prompt rule)")
    draft = base_draft({"mode": "none"})
    state = apply_draft(draft, {"targeting_groups": existing_groups()})
    nfl = next(o for o in state["plan_options"] if o["name"] == "NFL Regular Season")
    check("an explicit mode:none also resolves to no group rows",
         group_ids_in(nfl["rows"]) == [], group_ids_in(nfl["rows"]))

    print("\nD. match_groups_to_selection itself: mode:none is a real, distinct mode (offline unit check)")
    real_groups = existing_groups()
    ids, unmatched, mode = app.match_groups_to_selection(real_groups, {"mode": "none"})
    check("mode:none resolves to zero ids", ids == [], ids)
    check("mode:none reports its own mode, not 'all'", mode == "none", mode)
    ids2, _, mode2 = app.match_groups_to_selection(real_groups, {})
    check("an empty selection, called directly, is still 'all' (unchanged for a single-option caller)",
         mode2 == "all" and len(ids2) == 2, (mode2, ids2))

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
