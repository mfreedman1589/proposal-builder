"""Two more "the model had nowhere to put that intent" schema gaps, found
testing a live draft against the Capital Media avail (RFPID-266994) with a
Homeowner-style prompt asking to "Include the SOV" for a "deck replacement"
client -- the same class of bug as `group_cpm` (tests/test_group_cpm_
override.py): a rep's request was correct and unambiguous, but the draft
schema gave the model no field to act on it, so it could only ever land in
`unresolved` as a question instead of actually happening.

1. **"show_sov"** -- there was no field for turning on "% of avails" (share
   of voice) at all. A live call confirmed the model correctly RECOGNIZED
   the request ("The notes mention including SOV... confirm what SOV figure
   to include") but had nothing to set. Added "show_sov" to the schema and
   the prompt, wired into `apply_draft_to_form` the same way
   "spanish_campaign" already is.

2. **Vertical hint never reaches the model.** `_detect_vertical_hint` (used
   only to pick which audience segments the model sees) had no "deck
   replacement" entry, AND even a correctly-detected hint was never told to
   the model at all -- `vertical_hint` only ever shaped the catalog slice,
   never appeared in the prompt text. A live call for "a deck replacement
   client" returned `"vertical": "none"`. Fixed both: added "deck
   replacement"/"decking"/"deck builder" to VERTICAL_HINT_SYNONYMS, and the
   prompt's own "vertical" instruction now surfaces the detected hint as a
   suggestion (never authoritative -- the model still decides, and can
   override it when the notes point elsewhere).

Both are offline/pure checks -- no live API call needed to prove the schema
carries the field and the prompt surfaces it; tests/test_draft_live.py is
where a live call would confirm the MODEL actually uses it.

    python tests/test_draft_schema_gaps.py
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

import app  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    def __init__(self, preset=None):
        self.session_state = dict(preset or {})
        self.secrets = {}


def apply_draft(draft, preset_state=None):
    real = app.st
    stub = _StubSt(preset_state)
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return stub.session_state


def base_draft(**overrides):
    draft = {
        "client_name": "Test Co", "vertical": "none", "market": "DC",
        "agency_involved": False, "flight_start": "2026-09-01", "flight_end": "2026-09-30",
        "total_budget": 5000, "breakout": "monthly",
        "media_plan_lines": [{"product": "premion_streaming_tv",
                              "allocation": {"percent_of_total": 100}}],
        "audiences": [], "attribution": [], "sports": [],
        "campaign_specs": {}, "unresolved": [], "unresolved_internal": [],
    }
    draft.update(overrides)
    return draft


def check_show_sov_field():
    print("\"show_sov\" -- the draft schema/prompt/apply_draft_to_form wiring")

    check("\"show_sov\" is documented in the schema example",
          '"show_sov"' in app.DRAFT_JSON_SCHEMA_EXAMPLE, None)
    check("the prompt explains what \"show_sov\" does",
          "show_sov" in app.build_draft_prompt("test notes"), None)

    state_on = apply_draft(base_draft(show_sov=True))
    check("show_sov=true in the draft sets the session_state toggle on",
          state_on.get("show_sov") is True, state_on.get("show_sov"))

    state_off = apply_draft(base_draft())  # no show_sov key at all
    check("omitting show_sov defaults to off, never left as some stale prior value",
          state_off.get("show_sov") is False, state_off.get("show_sov"))

    state_false = apply_draft(base_draft(show_sov=False))
    check("show_sov=false explicitly is also off",
          state_false.get("show_sov") is False, state_false.get("show_sov"))


def check_vertical_hint_reaches_model():
    print("\nVertical hint: detection covers \"deck replacement\", and the "
          "prompt actually tells the model about it")

    check("\"deck replacement\" resolves to home_improvement",
          app._detect_vertical_hint("They're a deck replacement company in Northern VA.")
          == "home_improvement", app._detect_vertical_hint("deck replacement company"))
    check("\"decking\" resolves to home_improvement too",
          app._detect_vertical_hint("A decking and patio contractor.") == "home_improvement",
          app._detect_vertical_hint("A decking and patio contractor."))

    prompt_with_hint = app.build_draft_prompt("They're a deck replacement company wanting a CTV plan.")
    check("the prompt surfaces the detected hint to the model, not just the catalog slice",
          "home_improvement" in prompt_with_hint.split("Schema:")[0]
          or 'suggests "home_improvement"' in prompt_with_hint,
          None)
    check("the hint is phrased as a suggestion, not an instruction to return it unconditionally",
          "confirm it actually fits" in prompt_with_hint, None)

    prompt_no_hint = app.build_draft_prompt("Generic notes naming no trade at all.")
    check("no vertical hint sentence appears when nothing in the notes suggests one",
          "a trade term in the notes suggests" not in prompt_no_hint, None)


def main():
    check_show_sov_field()
    check_vertical_hint_reaches_model()

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A rep asking to \"include the SOV\" or naming a deck-replacement trade now has "
          "somewhere for that intent to land, the same fix shape as the group_cpm gap.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
