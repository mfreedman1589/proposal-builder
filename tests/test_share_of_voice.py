"""% of avails (share of voice) on the media plan -- BACKLOG.md.

Per-LINE, never a deck-wide aggregate: a plan row's impressions divided by
the avails of the targeting group(s) it actually traces back to via
`_group_ids` (`group_ids_of` / `matched_avails_for_row` in app.py) -- the
same join merge/split/the map legend already use, never a text match on
Targeting/Geo. A line with no such match (drafted, quick-added, hand-typed,
a stale id) shows no percentage. A merged line (2+ ids, from
`merge_plan_rows` combining several avails-backed lines into one) sums the
avails of every id it still carries -- confirmed with the user this is the
intended behavior (e.g. one avails document with one audience broken into
several markets, combined into a single plan line, should report its
percentage against the combined avails of those markets, not no percentage).

Gated behind a "Show % of avails" checkbox that defaults off.

Two independent AppTest instances (OFF, ON) rather than one instance
toggled mid-test: `show_sov` is a widget-keyed checkbox, and this project's
own DECISIONS.md entry on the AppTest session_state/rerun artifact is
precisely about the hazard of overriding an ALREADY-RENDERED widget's key
via `at.session_state[...]` in the same run as other state changes. Setting
it (or leaving it untouched) before that instance's very first `.run()`
avoids the question entirely, matching every other override in this file.

    python tests/test_share_of_voice.py
"""
import os
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

sys.path.insert(0, str(REPO / "tests"))
from test_form_state import load_drafted_form  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    return at


class _StStub:
    """Lets merge_plan_rows (which reads st.session_state for group labels)
    run outside of a live AppTest script, exactly the swap
    test_avails_reach.py's own `apply()` uses for the same reason."""

    def __init__(self, session_state):
        self.session_state = session_state


def build_option_with_known_lines():
    """One option, four lines covering every match shape this feature has to
    tell apart: a clean 1:1 match, no match at all, a merged (2-id) line, and
    a match against a group with no avails pulled yet (0). Returns the
    option plus the three real targeting groups and the expected percentages
    to check the merged and matched lines against.
    """
    group_a = tg.new_group(["Audience A"], avails_monthly=500_000)
    group_b1 = tg.new_group(["Audience B, Denver"], avails_monthly=300_000)
    group_b2 = tg.new_group(["Audience B, Atlanta"], avails_monthly=200_000)
    group_zero = tg.new_group(["Audience Z"], avails_monthly=0)
    groups = [group_a, group_b1, group_b2, group_zero]

    def row(tactic_suffix, impressions, group_ids=None):
        r = {"Tactic": f"Streaming TV {tactic_suffix}", "Flight": "", "Geo": "Washington, DC DMA",
             "Targeting": tactic_suffix, "Impressions": str(impressions), "Cost": "5000", "CPM": "30.00"}
        if group_ids is not None:
            r["_group_ids"] = group_ids
        return r

    matched_row = row("matched", 180_000, [group_a["id"]])
    unmatched_row = row("unmatched", 100_000)  # no _group_ids key at all
    b1_row = row("b1", 90_000, [group_b1["id"]])
    b2_row = row("b2", 60_000, [group_b2["id"]])
    zero_row = row("zero-avails", 50_000, [group_zero["id"]])

    option = app.new_plan_option(
        "Option A", [matched_row, unmatched_row, b1_row, b2_row, zero_row])

    # Merge the two Audience B lines (indices 2, 3) into one -- the real
    # merge_plan_rows, not a hand-built imitation, so the resulting
    # _group_ids is whatever that function actually produces.
    real_st, app.st = app.st, _StStub({"targeting_groups": groups})
    try:
        app.merge_plan_rows(option, [2, 3])
    finally:
        app.st = real_st

    expected = {
        "matched": round(180_000 / 500_000 * 100),
        "merged": round((90_000 + 60_000) / (300_000 + 200_000) * 100),
    }
    return option, groups, expected


def seed_known_state(at, show_sov):
    """A drafted baseline (so Generate has everything else it needs), with
    the avails and the plan replaced by known lines covering every match
    shape. Everything here is set before this instance's first `.run()` --
    no widget for any of these keys exists yet, so there's nothing to
    shadow.
    """
    load_drafted_form(at)

    option, groups, expected = build_option_with_known_lines()
    rows = tg.groups_to_seed_rows(groups, app.AVAILS_COLUMN_MONTHLY)
    at.session_state["targeting_groups"] = groups
    at.session_state["avails_seed_rows"] = rows
    # Tells sync_targeting_groups() the two sides already agree, so Generate
    # doesn't re-derive groups from rows and overwrite this override.
    at.session_state["_groups_rows_applied"] = [dict(r) for r in rows]
    at.session_state["plan_options"] = [option]

    if show_sov:
        at.session_state["show_sov"] = True

    return expected


def generate(at):
    real_personalize = app.assembly.personalize
    captured = {}

    def spy_personalize(prs, fill_data):
        captured["fill_data"] = fill_data
        return real_personalize(prs, fill_data)

    app.assembly.personalize = spy_personalize
    try:
        buttons = [b for b in at.button if b.label == "Generate proposal"]
        if not buttons:
            return None, at
        buttons[0].click().run()
    finally:
        app.assembly.personalize = real_personalize
    return captured.get("fill_data"), at


def sov_suffix_present(text):
    return bool(re.search(r"\(\d+% of avails\)", str(text)))


def rows_by_suffix(fill_data):
    return {r["tactic"]: r["impressions"] for r in fill_data["media_plan_options"][0]["rows"]}


def main():
    print("=" * 78)
    print("SCENARIO  % of avails is OFF by default -- no line gets a suffix")
    print("=" * 78)
    at_off = new_app()
    seed_known_state(at_off, show_sov=False)
    at_off.run()
    check("no exception after seeding known lines", not at_off.exception,
          at_off.exception[0].message[:400] if at_off.exception else "")

    fill_data, at_off = generate(at_off)
    check("Generate runs without raising", not at_off.exception,
          at_off.exception[0].message[:400] if at_off.exception else "")
    check("fill_data was captured", fill_data is not None)
    if fill_data is not None:
        impressions_by_tactic = rows_by_suffix(fill_data)
        for tactic, text in impressions_by_tactic.items():
            check(f"no suffix on {tactic!r} when the toggle is off",
                  not sov_suffix_present(text), text)

    print("\n" + "=" * 78)
    print("SCENARIO  % of avails ON -- matched, merged, unmatched and zero-avails lines")
    print("=" * 78)
    at_on = new_app()
    expected = seed_known_state(at_on, show_sov=True)
    at_on.run()
    check("no exception after seeding known lines", not at_on.exception,
          at_on.exception[0].message[:400] if at_on.exception else "")

    fill_data, at_on = generate(at_on)
    check("Generate runs without raising", not at_on.exception,
          at_on.exception[0].message[:400] if at_on.exception else "")
    check("fill_data was captured", fill_data is not None)
    if fill_data is not None:
        by_tactic = rows_by_suffix(fill_data)
        matched_text = by_tactic["Streaming TV matched"]
        unmatched_text = by_tactic["Streaming TV unmatched"]
        zero_text = by_tactic["Streaming TV zero-avails"]
        # merge_plan_rows keeps the survivor's own Tactic untouched -- the
        # merged line (b1 + b2) is still tagged "b1", the lower index.
        merged_text = by_tactic["Streaming TV b1"]

        check(f"a 1:1 matched line carries the ({expected['matched']}% of avails) suffix",
              f"({expected['matched']}% of avails)" in matched_text, matched_text)
        check("an unmatched line (no _group_ids) carries no suffix",
              not sov_suffix_present(unmatched_text), unmatched_text)
        check(f"a merged (2-id) line sums the merged groups' avails -> "
              f"({expected['merged']}% of avails)",
              f"({expected['merged']}% of avails)" in merged_text, merged_text)
        check("a line matched to a group with no avails pulled yet (0) carries no suffix",
              not sov_suffix_present(zero_text), zero_text)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("% of avails stays off unless a rep opts in; when it's on, a line only shows a "
          "percentage when its own _group_ids resolve to a real, non-zero avails figure -- "
          "summed across a merged line's ids, blank for an unmatched or zero-avails line -- "
          "and there is no deck-wide aggregate anywhere.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
