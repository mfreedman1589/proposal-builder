"""Phases 4-5 of the targeting-groups roadmap (geo_targeting_roadmap.md D):
the grouped avails table (D2) and the Audience finder's AND/OR/New-group
builder.

    python tests/test_group_builder.py

Driven through the real form with AppTest -- `st.data_editor` can't be
driven directly in this Streamlit version (confirmed: AppTest exposes no
`.data_editor` accessor), which is why every existing avails/plan test in
this suite also drives edits by writing session_state and re-running rather
than touching the grid widget itself. This file follows the same
convention, and additionally exercises what only the REAL script path can
show: what an ordinary, un-edited rerun does to state the grid displays.

The load-bearing regression this file exists for: the D2 grid can only
DISPLAY a multi-term group through `audience_label`'s plain, comma-joined
text ("A, B"), which is not the canonical form `terms_from_audience_text`
recognizes. A fold-back that re-derives terms from that displayed text on
EVERY run -- not just an actual edit -- collapses a 2-term AND group into
one verbatim term the moment the grid merely redraws it, with no edit at
all. This was a real bug caught by exactly this test while building Phase 4;
`main()` asserts an untouched extra rerun leaves a built group unchanged.
"""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    return at


def groups_of(at):
    return at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []


def search(at, text):
    widget = [w for w in at.text_input if str(w.key) == "finder_search"][0]
    widget.set_value(text)
    at.run()


def click(at, prefix, segment):
    widget = [w for w in at.button if str(w.key) == f"{prefix}{segment}"][0]
    widget.click()
    at.run()


def warnings_text(at):
    return [str(w.value) if hasattr(w, "value") else str(w) for w in at.warning]


def main():
    print("first click starts a one-term group and opens it")
    at = new_app()
    at.run()
    check("the form renders", not at.exception, at.exception[0].message[:300] if at.exception else "")
    search(at, "AFIRST Foodies")
    click(at, "finder_and_", "AFIRST Foodies")
    check("no exception", not at.exception, at.exception[0].message[:300] if at.exception else "")
    real_groups = [g for g in groups_of(at) if g["terms"]]
    check("one real group, one term, no operator",
          len(real_groups) == 1 and real_groups[0]["terms"] == ["AFIRST Foodies"]
          and real_groups[0]["op"] is None, real_groups)

    print("\nAND extends the open group into a 2-term AND")
    search(at, "AFIRST Fitness Fanatics")
    click(at, "finder_and_", "AFIRST Fitness Fanatics")
    check("no exception", not at.exception, at.exception[0].message[:300] if at.exception else "")
    real_groups = [g for g in groups_of(at) if g["terms"]]
    check("still one group, now AND of both terms",
          len(real_groups) == 1 and real_groups[0]["op"] == "AND"
          and real_groups[0]["terms"] == ["AFIRST Foodies", "AFIRST Fitness Fanatics"], real_groups)

    print("\nTHE REGRESSION: an untouched rerun must not collapse the AND group")
    at.run()   # nothing clicked, nothing edited -- D2's editor just redraws
    check("no exception on the extra rerun", not at.exception,
          at.exception[0].message[:300] if at.exception else "")
    real_groups = [g for g in groups_of(at) if g["terms"]]
    check("the 2-term AND group survives byte-for-byte",
          len(real_groups) == 1 and real_groups[0]["op"] == "AND"
          and real_groups[0]["terms"] == ["AFIRST Foodies", "AFIRST Fitness Fanatics"], real_groups)

    print("\nmixing OR onto an open AND group is refused, and the group is untouched")
    before = [dict(g) for g in groups_of(at)]
    search(at, "AFIRST Gamers and Tech Junkies")
    click(at, "finder_or_", "AFIRST Gamers and Tech Junkies")
    check("no exception", not at.exception, at.exception[0].message[:300] if at.exception else "")
    check("groups are exactly what they were before the refused click",
          groups_of(at) == before, groups_of(at))
    check("a refusal message is shown, naming the operator already in force",
          any("already" in w and "AND" in w for w in
              ([str(x.value) if hasattr(x, "value") else str(x) for x in at.error])),
          [str(x.value) if hasattr(x, "value") else str(x) for x in at.error])

    print("\nthe refusal message shows once, not on a later unrelated rerun")
    at.run()
    check("no leftover refusal banner",
          not any("already" in (str(x.value) if hasattr(x, "value") else str(x)) for x in at.error),
          [str(x.value) if hasattr(x, "value") else str(x) for x in at.error])

    print("\n\"New group\" commits the open group and starts a separate one")
    search(at, "AFIRST Gamers and Tech Junkies")
    click(at, "finder_new_", "AFIRST Gamers and Tech Junkies")
    check("no exception", not at.exception, at.exception[0].message[:300] if at.exception else "")
    real_groups = [g for g in groups_of(at) if g["terms"]]
    check("two real groups now: the AND stack, untouched, plus the new single-term one",
          len(real_groups) == 2
          and any(g["op"] == "AND" and g["terms"] == ["AFIRST Foodies", "AFIRST Fitness Fanatics"]
                  for g in real_groups)
          and any(g["terms"] == ["AFIRST Gamers and Tech Junkies"] and g["op"] is None
                  for g in real_groups),
          real_groups)

    print("\nOR works the same way, symmetric to AND")
    at2 = new_app()
    at2.run()
    search(at2, "AUTO Body Style Compact")
    click(at2, "finder_new_", "AUTO Body Style Compact")
    search(at2, "AUTO Body Style Minivan")
    click(at2, "finder_or_", "AUTO Body Style Minivan")
    check("no exception", not at2.exception, at2.exception[0].message[:300] if at2.exception else "")
    real_groups = [g for g in groups_of(at2) if g["terms"]]
    target = [g for g in real_groups if "AUTO Body Style Compact" in g["terms"]]
    check("OR joins both terms on one group",
          len(target) == 1 and target[0]["op"] == "OR"
          and set(target[0]["terms"]) == {"AUTO Body Style Compact", "AUTO Body Style Minivan"},
          real_groups)

    print("\nmixing AND onto an open OR group is refused too (symmetric)")
    before2 = [dict(g) for g in groups_of(at2)]
    search(at2, "AUTO Body Style Pickup and SUV")
    click(at2, "finder_and_", "AUTO Body Style Pickup and SUV")
    check("groups unchanged by the refused AND", groups_of(at2) == before2, groups_of(at2))

    print("\nthe one-custom rule reaches the generate-time review list, not just the click")
    # Both AUTO Body Style segments above are non-RFP-selectable, so the OR
    # group alone already carries 2 distinct customs.
    cat = app.load_audience_catalog()
    rfp_map = dict(zip(cat["segment"], cat["rfp_selectable"]))
    live_count = tg.custom_segment_count(groups_of(at2), rfp_map)
    check("2 distinct customs are in play", live_count == 2, live_count)
    at2.run()   # render_review_list recomputes live every run, top of page
    check("no exception", not at2.exception, at2.exception[0].message[:300] if at2.exception else "")
    review = warnings_text(at2)
    hit = [w for w in review if "custom" in w.lower() and "targeting groups" in w.lower()]
    check("the review-before-generating panel names the overage",
          bool(hit) and "2 custom" in hit[0], review)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The grid displays a group's plain-language label, but an untouched rerun never "
          "re-derives its structure from that text -- and the one-custom rule is visible at "
          "generate time, not only in the toast at the moment of the click.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
