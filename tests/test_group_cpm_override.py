"""A negotiated CPM stated in the notes must survive drafting even when
targeting groups already own the plan (an avails PDF already imported) --
found testing the real Plaza Motors avail (RFPID-266583): the notes stated
"$29 CPM" but the generated line priced at $32, the rate-card default for
Premion Streaming TV.

Root cause, confirmed directly (not guessed): the model DID extract the rate
correctly on a fresh, no-groups draft (a live call reproduced this: the
drafted media_plan_lines entry carried "cpm": 29). The loss happened
downstream, in a genuine schema/prompt gap -- once real targeting groups
already own the avails table, the prompt's own "group_section" tells the
model to sell via "group_selection"/"group_allocation" instead of writing a
"premion_streaming_tv" media_plan_lines entry, but never told it there was
any field at all for a rate on that path. And even where the plumbing
already half-existed (`apply_draft_to_form` already read
`opt_in.get("group_cpm")`), `drafted_options()` never copied a "group_cpm"
key out of the raw draft into the per-option dict it returns -- so even a
model that guessed the field name right had it silently dropped before
`apply_draft_to_form` ever saw it. The fix is three-part: the schema example
and the group-selection prompt text now document "group_cpm", and
`drafted_options()` carries it through (top-level and per-option, same
fallback rule as "group_allocation").

This is NOT the "a rate-card reseed on product toggle" or "the avails-import
path" or "resolve_row_defaults" causes the bug report asked to rule out --
none of those touch this row at all, confirmed by tracing
`_group_alloc_line`/`_resolve_line_cpm` directly. It is a fourth thing: the
model was never given anywhere to put the number.

    python tests/test_group_cpm_override.py
"""
import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import targeting_groups as tg                  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    """Same technique tests/test_draft_regression.py uses: apply_draft_to_form
    only touches st.session_state, so a plain dict stands in for Streamlit."""

    def __init__(self, preset=None):
        self.session_state = dict(preset or {})
        self.secrets = {}


def apply_draft(draft, preset_state):
    real = app.st
    stub = _StubSt(preset_state)
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return stub.session_state


def two_real_groups():
    return [
        tg.new_group(["DEMO Age A35-64", "AUTO Type Luxury"], op="AND",
                    geo_def={"kind": "text", "label": "Zip list provided by agency"},
                    avails_monthly=1707337, color="#4C78A8"),
        tg.new_group(["AUTO Type Luxury", "DEMO Age M35-64"], op="AND",
                    geo_def={"kind": "text", "label": "Zip list provided by agency"},
                    avails_monthly=1126832, color="#4C78A8"),
    ]


def base_draft(groups, **overrides):
    draft = {
        "client_name": "Plaza Motors", "vertical": "auto", "market": "DC",
        "flight_start": "2026-09-17", "flight_end": "2026-09-30",
        "total_budget": 4000, "breakout": "full_flight",
        "media_plan_lines": [],
        "group_selection": {"mode": "named", "ids": [groups[0]["id"]]},
        "group_allocation": {"flat_amount": 4000},
        "group_selection_reason": "Sells the A35-64 x Luxury row per the notes.",
        "audiences": [], "attribution": [], "sports": [],
        "campaign_specs": {}, "unresolved": [], "unresolved_internal": [],
    }
    draft.update(overrides)
    return draft


def premion_row(state):
    for opt in state.get("plan_options") or []:
        for row in opt["rows"]:
            if row.get("Tactic") == "Premion Streaming TV":
                return row
    return None


def check_proposal_history_survival(drafted_state, groups):
    """A negotiated CPM must survive proposal history, not just the initial
    draft. Two paths, two different guarantees:

    1. "Load into form" (`rehydrate_proposal_into_form`) copies
       `form_json["plan_options"]`'s rows -- and therefore each row's own
       "CPM" -- VERBATIM into session_state; there is no CPM-specific
       re-derivation anywhere in that path (confirmed by reading it, not
       assumed). This test proves that stays true: a $29 row logged into
       form_json comes back as $29, not recomputed against the $32 rate
       card default.
    2. "Rebuild as presented" (`rebuild_proposal_deck`) is a STRONGER, by-
       design guarantee documented directly in its own docstring: it reads
       `form_json["deck_payload"]["media_plan_options"]` verbatim and never
       recomputes it, specifically so "a CPM change since would quietly
       rewrite what the client was shown" can't happen. That guarantee has
       no CPM-specific logic to regress -- it is the ABSENCE of any
       recompute step -- so it is not re-tested here with a full deck
       build; the code path itself is the evidence, and this file's own
       root-cause tracing already confirmed nothing upstream of it can
       reintroduce the rate-card default once a row's CPM is set.
    """
    print("\nProposal history: a $29 row survives \"Load into form\"")
    form_json = {
        "targeting_groups": groups,
        "plan_options": [
            {"name": "Option A", "breakout": "full_flight",
             "rows": drafted_state["plan_options"][0]["rows"],
             "driver": drafted_state["plan_options"][0]["driver"]},
        ],
        "avails_basis": app.AVAILS_BASIS_MONTHLY,
        "markup": 1.0,
        "flight_start": "2026-09-17", "flight_end": "2026-09-30",
    }
    stored_row = {
        "id": "test-proposal-id", "client_name": "Plaza Motors",
        "form_json": form_json, "market": "DC",
    }

    real = app.st
    stub = _StubSt()
    app.st = stub
    try:
        app.rehydrate_proposal_into_form(stored_row, parent_proposal_id=stored_row["id"])
    finally:
        app.st = real

    restored = premion_row(stub.session_state)
    check("the Premion Streaming TV row survived rehydration",
          restored is not None, stub.session_state.get("plan_options"))
    if restored:
        check("its CPM is still the negotiated $29, not re-derived to $32",
              restored["CPM"] == 29.0, restored["CPM"])
        check("its Cost is unchanged too (a full verbatim row copy, not a partial one)",
              restored["Cost"] == 4000.0, restored["Cost"])


def main():
    groups = two_real_groups()

    print("A drafted \"group_cpm\" survives onto the group-selection line")
    draft = base_draft(groups, group_cpm=29)
    state = apply_draft(draft, {"targeting_groups": groups})
    row = premion_row(state)
    check("a Premion Streaming TV row was created", row is not None, state.get("plan_options"))
    if row:
        check("CPM is the negotiated $29, not the $32 rate card default",
              row["CPM"] == 29.0, row["CPM"])
    check("the override announces itself in unresolved_internal, per the existing rule",
          any("29.00" in u and "32.00" in u for u in state.get("draft_unresolved_internal") or []),
          state.get("draft_unresolved_internal"))

    print("\nNo \"group_cpm\" in the draft -- rate card default applies, with no false override note")
    draft2 = base_draft(groups)  # no group_cpm key at all
    state2 = apply_draft(draft2, {"targeting_groups": groups})
    row2 = premion_row(state2)
    check("CPM falls back to the $32 rate card default", row2 and row2["CPM"] == 32.0,
          row2["CPM"] if row2 else None)
    check("no negotiated-rate note is fabricated when nothing was negotiated",
          not any("rate card rate" in u for u in state2.get("draft_unresolved_internal") or []),
          state2.get("draft_unresolved_internal"))

    print("\nA per-option \"group_cpm\" overrides a top-level one, same fallback rule as group_allocation")
    draft3 = base_draft(groups, group_cpm=99)
    draft3["options"] = [dict(
        name="Option A", total_budget=4000, breakout="full_flight",
        media_plan_lines=[], group_selection={"mode": "named", "ids": [groups[0]["id"]]},
        group_allocation={"flat_amount": 4000}, group_cpm=29,
        group_selection_reason="per-option override")]
    del draft3["group_selection"], draft3["group_allocation"], draft3["group_selection_reason"]
    state3 = apply_draft(draft3, {"targeting_groups": groups})
    row3 = premion_row(state3)
    check("the option's own group_cpm (29) wins over the top-level one (99)",
          row3 and row3["CPM"] == 29.0, row3["CPM"] if row3 else None)

    check_proposal_history_survival(state, groups)

    print("\ndrafted_options() itself carries group_cpm through, top-level and per-option")
    d = {"total_budget": 1000, "group_cpm": 28}
    opts = app.drafted_options(d)
    check("bare media_plan_lines draft: top-level group_cpm carried onto the one option",
          opts[0]["group_cpm"] == 28, opts[0])
    d2 = {"total_budget": 1000, "group_cpm": 28,
         "options": [{"name": "A", "total_budget": 500},
                    {"name": "B", "total_budget": 500, "group_cpm": 45}]}
    opts2 = app.drafted_options(d2)
    check("option A falls back to the top-level group_cpm", opts2[0]["group_cpm"] == 28, opts2[0])
    check("option B's own group_cpm wins", opts2[1]["group_cpm"] == 45, opts2[1])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A negotiated CPM now has somewhere to go once targeting groups own the plan, "
          "and it flows through to the same row the rate-card default would have priced.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
