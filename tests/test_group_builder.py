"""Phases 4-5 of the targeting-groups roadmap (geo_targeting_roadmap.md D):
the grouped avails table (D2) and the Audience finder's AND/OR/New-group
builder -- plus the booking-evidence panel (Phase 7's data layer wired into
this same builder), against the real committed audience_usage_ytd.csv.

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

import db  # noqa: E402

# Forces the local brochure-PDF fallback for the audience catalog rather than
# scoring the one-custom-audience count (below) against whatever's live in
# Supabase at test time -- same fix, same reason, as
# tests/test_group_scenarios.py and tests/test_draft_regression.py.
# MUST run before `import app`: app.py calls load_audience_catalog() at its
# own module scope, which would otherwise cache the live result first.
db.fetch_audiences = lambda: (None, "stubbed for test isolation -- see comment above")

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
    from datetime import date
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (where the
    # Audience finder this file drives lives) on market+flight.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def groups_of(at):
    return at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []


def search(at, text):
    # [-1], not [0]: once the booking-evidence panel (Phase 7) renders a
    # variable number of caption lines ahead of this widget in the SAME
    # block, an interrupted run (the one a click's st.rerun() aborts
    # mid-script) can leave a stale, disconnected "finder_search" proxy
    # ahead of the real one in AppTest's element list -- confirmed empirically
    # (different object ids, no at.exception either way, so the live app
    # never actually double-renders the widget; this is an AppTest artifact
    # of the interrupted-run/rerun pair, not a real bug). The live widget is
    # always the last one returned.
    widget = [w for w in at.text_input if str(w.key) == "finder_search"][-1]
    widget.set_value(text)
    at.run()


def click(at, prefix, segment):
    widget = [w for w in at.button if str(w.key) == f"{prefix}{segment}"][-1]
    widget.click()
    at.run()


def warnings_text(at):
    return [str(w.value) if hasattr(w, "value") else str(w) for w in at.warning]


def captions(at):
    return [str(c.value) if hasattr(c, "value") else str(c) for c in at.caption]


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

    print("\na third AND is refused -- an audience line combines at most 2 segments")
    before_and = [dict(g) for g in groups_of(at)]
    search(at, "AFIRST Gamers and Tech Junkies")
    click(at, "finder_and_", "AFIRST Gamers and Tech Junkies")
    check("no exception", not at.exception, at.exception[0].message[:300] if at.exception else "")
    check("the 2-term AND group is untouched by the refused third AND",
          groups_of(at) == before_and, groups_of(at))
    check("the refusal names the 2-segment limit",
          any("at most 2" in (str(x.value) if hasattr(x, "value") else str(x)) for x in at.error),
          [str(x.value) if hasattr(x, "value") else str(x) for x in at.error])
    at.run()

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

    print("\nOR across lines is unlimited -- a third OR term is accepted")
    search(at2, "AUTO Body Style Pickup and SUV")
    click(at2, "finder_or_", "AUTO Body Style Pickup and SUV")
    target = [g for g in groups_of(at2) if "AUTO Body Style Compact" in (g["terms"] or [])]
    check("the OR group now has 3 terms", target and len(target[0]["terms"]) == 3, target)
    at2.run()
    review = warnings_text(at2)
    check("an OR group of 3 raises no 2-segment warning",
          not any("at most 2" in w for w in review), review)

    print("\na hand-built AND of 3 (e.g. typed into D2) is flagged before generating")
    over = tg.new_group(["DEMO Homeowner", "HH Income 150K Plus", "HH Home Built Before 2010"],
                        op="AND")
    check("and_limit_violations finds it", tg.and_limit_violations([over]) == [over])
    check("the message names the line",
          "DEMO Homeowner" in (tg.and_limit_message([over]) or ""), tg.and_limit_message([over]))
    check("a 2-term AND and a 3-term OR are fine",
          not tg.and_limit_violations([
              tg.new_group(["A", "B"], op="AND"), tg.new_group(["A", "B", "C"], op="OR")]))

    print("\nbooking evidence renders while building, and only while building")
    at3 = new_app()
    at3.run()
    check("no booking-evidence panel before anything is open",
          not any("Booking evidence" in c for c in captions(at3)), captions(at3))
    search(at3, "DEMO Homeowner")
    click(at3, "finder_and_", "DEMO Homeowner")
    caps = captions(at3)
    check("evidence panel appears after the first term",
          any("Booking evidence" in c for c in caps), caps)
    # Wording follows the impressions-based familiarity signal ("rank by
    # impressions, not by count") -- "appears in N booked stacks" was the
    # stack-count wording this line carried before that change.
    check("component familiarity is shown for a single term",
          any("DEMO Homeowner has delivered" in c and "impressions" in c for c in caps), caps)
    search(at3, "HH Income 150K Plus")
    click(at3, "finder_and_", "HH Income 150K Plus")
    caps = captions(at3)
    # Asserted against the real committed audience_usage_ytd.csv, not
    # re-derived through the module under test -- this is the exact worked
    # pair from geo_targeting_roadmap.md D's own validation.
    check("the weakest-pair line states the real co-occurrence count",
          any("DEMO Homeowner and HH Income 150K Plus have been booked together 13 times." == c
              for c in caps), caps)
    check("no pairing is suggested once the AND line has 2 segments (the pairing fills slot 2)",
          not any(c.startswith("Often paired with what's selected:") for c in caps), caps)

    print("\na 'New group' click adds to the avails table only -- no automatic plan line "
          "(the avails table is research; ticking Plan is what puts a group on the media "
          "plan) -- and ticking a SECOND group's Plan box doesn't disturb the first's line, "
          "the same regression concern test_avails_pdf_wiring.py's Hershey scenario covers "
          "for an import, reproduced here through the finder's own button one group at a time")
    at4 = new_app()
    at4.session_state["premion_streaming_tv"] = True
    at4.run()
    search(at4, "DEMO Homeowner")
    click(at4, "finder_new_", "DEMO Homeowner")
    real1 = [g for g in groups_of(at4) if g["terms"]]
    rows1 = at4.session_state["plan_options"][0]["rows"]
    premion1 = [r for r in rows1 if set(app.group_ids_of(r)) & {real1[0]["id"]}]
    check("no automatic Premion Streaming TV line for a freshly-added group",
          len(premion1) == 0, rows1)
    check("the group itself exists, just not included in the plan yet",
          len(real1) == 1 and not real1[0]["include_in_plan"], real1)

    # Simulate ticking "Plan" for the first group -- the same thing the D2
    # grid's checkbox column does under the hood (data_editor can't be
    # driven directly, same as every other avails/plan test in this file).
    groups = groups_of(at4)
    for g in groups:
        if g["id"] == real1[0]["id"]:
            g["include_in_plan"] = True
    at4.session_state["targeting_groups"] = groups
    at4.run()
    rows1b = at4.session_state["plan_options"][0]["rows"]
    premion1b = [r for r in rows1b if set(app.group_ids_of(r)) & {real1[0]["id"]}]
    check("ticking Plan gives the first group its own Premion Streaming TV line",
          len(premion1b) == 1 and premion1b[0]["Targeting"] == "DEMO Homeowner", rows1b)

    search(at4, "HH Income 150K Plus")
    click(at4, "finder_new_", "HH Income 150K Plus")
    check("no exception", not at4.exception, at4.exception[0].message[:300] if at4.exception else "")
    real2 = [g for g in groups_of(at4) if g["terms"]]
    check("two real groups now", len(real2) == 2, real2)
    rows2 = at4.session_state["plan_options"][0]["rows"]
    second_group_id = next(g["id"] for g in real2 if g["terms"] == ["HH Income 150K Plus"])
    check("the SECOND group has no line yet either -- added to avails, not the plan",
          not any(set(app.group_ids_of(r)) & {second_group_id} for r in rows2), rows2)
    check("the first group's line is still there, untouched by adding a second group",
          any(set(app.group_ids_of(r)) & {real1[0]["id"]} for r in rows2), rows2)

    groups2 = groups_of(at4)
    for g in groups2:
        if g["id"] == second_group_id:
            g["include_in_plan"] = True
    at4.session_state["targeting_groups"] = groups2
    at4.run()
    rows3 = at4.session_state["plan_options"][0]["rows"]
    premion3_first = [r for r in rows3 if set(app.group_ids_of(r)) & {real1[0]["id"]}]
    premion3_second = [r for r in rows3 if set(app.group_ids_of(r)) & {second_group_id}]
    check("the SECOND group also gets its own Premion Streaming TV line once ticked -- not "
          "silently absent", len(premion3_second) == 1
          and premion3_second[0]["Targeting"] == "HH Income 150K Plus", rows3)
    check("the first group's line is STILL there too, untouched by the second tick",
          len(premion3_first) == 1 and premion3_first[0]["Targeting"] == "DEMO Homeowner", rows3)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The grid displays a group's plain-language label, but an untouched rerun never "
          "re-derives its structure from that text -- an AND line stops at 2 segments while OR "
          "is unlimited, and the booking-evidence panel tracks the group being built, live, "
          "against the real usage log.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
