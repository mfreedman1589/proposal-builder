"""Post-draft check: a plan line's Targeting text against its segments.

1. An attribute the Targeting description names (homeowner, income, age,
   home age, children, gender) that no segment on the line covers -- a
   CUSTOM segment's packaged criteria included -- becomes a review item with
   two fixes: add the matching segment (only when the line has room under
   the 2-segment AND rule) or remove the phrase from the description.
2. A threshold the notes state that a same-kind segment on the line
   contradicts (notes "$150K+", segment "$200K Plus") becomes a review item
   suggesting the closest catalog segment.

Both are structured items (id, type, priority, question, options with
patches, default) for the interactive draft follow-ups. Expected values come
from the notes and the segment names themselves, never from the checker.

    python tests/test_targeting_check.py
"""
import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import pandas as pd  # noqa: E402

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import targeting_check as tc  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


# Real segment names (the ranked usage list), small enough to reason about.
CATALOG = [{"segment": name, "usage": usage} for name, usage in [
    ("DEMO Homeowner", 228), ("HH Income 100K Plus", 284), ("HH Income 150K Plus", 216),
    ("HH Income 200K Plus", 57), ("HH Income Less than 75K", 41),
    ("DEMO Age A35 Plus", 273), ("DEMO Age A25-54", 155), ("DEMO Age F35-64", 60),
    ("HH Home Built Before 2000", 17), ("HH Children in Household", 137),
    ("HEALTH Dental Care", 40), ("RETAIL Home Services HVAC", 30),
    ("CUSTOM DEMO Homeowner AND HH Income 150K Plus", 25),
    ("CUSTOM HH Income 200K Plus AND HH Home Built Before 2000", 5),
]]

DENTAL_NOTES = (REPO / "tests/fixtures/dental_single_option.txt").read_text(encoding="utf-8")
HVAC_NOTES = ("Ridgeline HVAC. Northern Virginia homeowners, household income $150K+, "
              "replacement systems. $40,000 total, Premion Streaming TV only.")


def items_of(check_name, items):
    return [i for i in items if i["check"] == check_name]


def check_item_shape(item):
    keys = {"id", "type", "priority", "question", "options", "default"}
    check(f"item {item['id']} carries {sorted(keys)}", keys <= set(item), sorted(item))
    check("its default is an option, and skipping it changes nothing (patch None)",
          any(o["id"] == item["default"] and o["patch"] is None for o in item["options"]),
          item["options"])


def check_dental_homeowners():
    print("Dental: Targeting says homeowners, no homeowner segment on the line")
    line = {"tactic": "Premion Streaming TV", "segments": ["HEALTH Dental Care"],
            "audience_track": "Adult homeowners considering cosmetic dental work"}
    items = items_of("targeting_text_uncovered",
                     tc.targeting_review_items([line], DENTAL_NOTES, CATALOG, "Plan"))
    check("flagged once", len(items) == 1, items)
    if not items:
        return
    item = items[0]
    check_item_shape(item)
    check("it's a decision", item["priority"] == "decision" and item["type"] == "choice", item)
    by_id = {o["id"]: o for o in item["options"]}
    check("fix 1: add DEMO Homeowner (the line has one segment, room for a second)",
          by_id.get("add_segment", {}).get("patch") == {
              "op": "add_segment", "option": "Plan", "tactic": "Premion Streaming TV",
              "segment": "DEMO Homeowner"}, by_id.get("add_segment"))
    check("fix 2: remove the phrase from the description",
          (by_id.get("remove_phrase", {}).get("patch") or {}).get("phrase") == "homeowners",
          by_id.get("remove_phrase"))
    check("the question names the line and the phrase",
          "Premion Streaming TV" in item["question"] and '"homeowners"' in item["question"],
          item["question"])

    print("A full line (2 segments) offers only removing the phrase")
    full = dict(line, segments=["HEALTH Dental Care", "DEMO Age A35 Plus"])
    items = items_of("targeting_text_uncovered",
                     tc.targeting_review_items([full], "", CATALOG, "Plan"))
    ids = [o["id"] for o in items[0]["options"]] if items else []
    check("no add-segment fix past the AND limit", "add_segment" not in ids and
          "remove_phrase" in ids, ids)
    check("the AND limit here is the app's own", tc.MAX_AND_SEGMENTS == tg.MAX_AND_SEGMENTS)


def check_hvac_threshold():
    print("HVAC: notes say $150K+, the line uses a $200K custom")
    line = {"tactic": "Premion Streaming TV",
            "segments": ["CUSTOM HH Income 200K Plus AND HH Home Built Before 2000"],
            "audience_track": "Homeowners in older homes with $200K+ household income"}
    items = tc.targeting_review_items([line], HVAC_NOTES, CATALOG, "Plan")
    swaps = items_of("notes_threshold_mismatch", items)
    check("the income mismatch is flagged", len(swaps) == 1, items)
    if swaps:
        item = swaps[0]
        check_item_shape(item)
        swap = next((o for o in item["options"] if o["id"] == "swap_segment"), None)
        to = (swap or {}).get("patch", {}).get("to") or []
        check("suggested swap is $150K (what the notes said)", "HH Income 150K Plus" in to, to)
        check("the custom's home-age criterion rides along as the second segment",
              "HH Home Built Before 2000" in to and len(to) <= tg.MAX_AND_SEGMENTS, to)
        check("the patch replaces the custom it found",
              swap["patch"]["from"] == line["segments"][0], swap)
        check("the question quotes both thresholds",
              "$150K+" in item["question"] and "$200K+" in item["question"], item["question"])
    check("homeowners is also flagged: the custom doesn't target homeowners",
          any("homeowners" in i["question"].lower()
              for i in items_of("targeting_text_uncovered", items)), items)


def check_matching_line():
    print("A line whose segments match its text and the notes raises nothing")
    line = {"tactic": "Premion Streaming TV",
            "segments": ["CUSTOM DEMO Homeowner AND HH Income 150K Plus"],
            "audience_track": "Homeowners with $150K+ household income"}
    items = tc.targeting_review_items([line], HVAC_NOTES, CATALOG, "Plan")
    check("no items", items == [], items)
    women = {"tactic": "P", "segments": ["DEMO Age F35-64"], "audience_track": "Women 35-64"}
    check("women 35-64 is covered by F35-64",
          tc.targeting_review_items([women], "", CATALOG) == [],
          tc.targeting_review_items([women], "", CATALOG))
    retargeting = {"tactic": "Streaming Retargeting", "audience_track": "Homeowners who saw the spot"}
    check("a line that names no segments isn't checked",
          tc.targeting_review_items([retargeting], HVAC_NOTES, CATALOG) == [])


def check_parsing():
    print("Reading attributes")
    kinds = {a["kind"] for a in tc.text_attributes(
        "Homeowners with $150K+ household income, adults 35-64, homes built before 2000, moms")}
    check("text: homeowner, income, age, home age, children, gender",
          kinds == {"homeowner", "income", "age", "home_built", "children", "gender"}, kinds)
    custom = tc.segment_attributes("CUSTOM HH Income 200K Plus AND HH Home Built Before 2000")
    check("a custom covers each criterion it packages",
          {"income", "home_built"} == {a["kind"] for a in custom}, custom)
    under = tc.text_attributes("household income under $75K")
    check("'under $75K' is a ceiling", under and under[0]["max"] == 75000 and
          under[0]["min"] is None, under)
    check("a year in the text isn't read as an age",
          not tc.text_attributes("Running January through March 2026"))


def check_through_apply_draft():
    print("\nThrough apply_draft_to_form: items land in draft_review_items, as prose in review")
    draft = {
        "client_name": "QA-TEST-Brightleaf Family Dental", "vertical": "healthcare",
        "total_budget": 30000, "breakout": "full_flight",
        "media_plan_lines": [{"product": "premion_streaming_tv", "label": "Families",
                              "audience_track": "Adult homeowners considering cosmetic dental work",
                              "segments": ["HEALTH Dental Care"],
                              "allocation": {"percent_of_total": 100}}],
        "audiences": [], "attribution": [], "sports": [], "unresolved": [],
    }
    real_st, real_load = app.st, app.load_audience_catalog
    stub = type("S", (), {})()
    stub.session_state = {"flight_start": pd.Timestamp("2027-01-01").date(),
                          "flight_end": pd.Timestamp("2027-03-31").date(), "avails_mode": False}
    stub.secrets = {}
    app.st = stub
    app.load_audience_catalog = lambda: pd.DataFrame(CATALOG)
    try:
        app.apply_draft_to_form(copy.deepcopy(draft), notes=DENTAL_NOTES)
        state = stub.session_state
        items = state.get("draft_review_items") or []
        check("one targeting item stored, structured", len(items) == 1 and
              items[0]["check"] == "targeting_text_uncovered", items)
        if items:
            patch = next(o["patch"] for o in items[0]["options"] if o["id"] == "add_segment")[0]
            tactic = state["plan_options"][0]["rows"][0]["Tactic"]
            check("its patch names the row's real Tactic", patch["tactic"] == tactic,
                  (patch, tactic))
        check("the draft's own notes aren't touched by it",
              not any("homeowners" in n for n in (state.get("draft_unresolved_internal") or [])),
              state.get("draft_unresolved_internal"))

        # A re-draft that keeps the rep's own plan keeps the items raised on it.
        app.apply_draft_to_form(copy.deepcopy(dict(draft, media_plan_lines=[])),
                                skip_sections={"media_plan"}, notes=DENTAL_NOTES)
        check("kept across a re-draft that preserved the plan",
              stub.session_state.get("draft_review_items") == items)
    finally:
        app.st, app.load_audience_catalog = real_st, real_load


def main():
    check_parsing()
    check_dental_homeowners()
    check_hvac_threshold()
    check_matching_line()
    check_through_apply_draft()
    print(f"\n{'ALL PASS' if not failures else f'{len(failures)} FAILED'}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
