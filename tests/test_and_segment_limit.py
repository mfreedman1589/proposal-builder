"""An audience line ANDs at most two segments (Matt, 2026-10-20).

Geography doesn't count; a CUSTOM segment counts as one, since that's what
customs are for; OR across separate lines is unlimited. Drafting asks the
model to stay inside it and to reach for a custom combo first, and
`resolve_drafted_lines` enforces it whatever comes back: a line naming
three or more keeps its first two and lists the rest in the review notes.
The finder's AND/OR builder refuses a third AND (tests/test_group_builder.py
drives that through the real page); `targeting_groups.and_limit_violations`
is the shared check behind the D2 and review-list warnings.

    python tests/test_and_segment_limit.py
"""
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
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def resolve(lines, budget=30000):
    rows, _products, _sports, unresolved, _drivers = app.resolve_drafted_lines(
        lines, budget, "Oct–Dec 2026", "Washington, DC", "Homeowners", n_months=3)
    return rows, unresolved


def line(label, segments, track, share):
    return {"product": "premion_streaming_tv", "label": label, "audience_track": track,
            "segments": segments, "allocation": {"percent_of_total": share}}


def main():
    print("A drafted line with 3 ANDed segments is trimmed to 2")
    rows, notes = resolve([line("HVAC", ["RETAIL Home Services HVAC", "DEMO Homeowner",
                                         "HH Income 100K Plus"],
                                "HVAC shoppers who own their home, $100K+ income", 100)])
    targeting = rows[0]["Targeting"] if rows else ""
    check("one row", len(rows) == 1, rows)
    check("its Targeting names exactly the first two segments",
          targeting == "RETAIL Home Services HVAC, DEMO Homeowner", targeting)
    note = [n for n in notes if "at most 2" in n]
    check("the dropped segment is named in the review notes",
          note and "HH Income 100K Plus" in note[0], notes)

    print("\nHome age + income + homeowner: a custom combo plus one segment is within the limit")
    rows, notes = resolve([line("HVAC", ["CUSTOM HH Home Built Before 2014 AND RETAIL Home Services HVAC",
                                         "HH Income 100K Plus"],
                                "Older homes in-market for HVAC, $100K+ income", 100)])
    check("kept as drafted", rows and rows[0]["Targeting"] == "Older homes in-market for HVAC, $100K+ income",
          rows)
    check("no trim note", not [n for n in notes if "at most 2" in n], notes)

    print("\nHome age + income + homeowner as three separate segments: trimmed to 2")
    rows, notes = resolve([line("HVAC", ["HH Home Built Before 2014", "HH Income 100K Plus",
                                         "DEMO Homeowner"], "Older homes, $100K+, homeowners", 100)])
    check("never more than 2 segments on the line",
          rows and len(rows[0]["Targeting"].split(", ")) == 2, rows)
    check("the third is in the review notes",
          any("DEMO Homeowner" in n and "at most 2" in n for n in notes), notes)

    print("\nA segment already inside a custom on the same line isn't counted twice")
    rows, notes = resolve([line("HVAC", ["CUSTOM HH Income 200K Plus AND HH Home Built Before 2000",
                                         "HH Home Built Before 2000", "DEMO Homeowner"],
                                "Homeowners in older homes, $200K+", 100)])
    check("custom + its own component + one more is 2 segments, not 3",
          not [n for n in notes if "at most 2" in n], notes)
    check("Targeting as drafted", rows and rows[0]["Targeting"] == "Homeowners in older homes, $200K+",
          rows)

    print("\nMultiple separate lines are allowed -- OR across lines is unlimited")
    rows, notes = resolve([
        line("Road Trips", ["TRAVEL Road Trips"], "Road trippers", 25),
        line("Families", ["TRAVEL Family", "HH Children in Household"], "Families who travel", 25),
        line("Outdoor", ["TRAVEL Outdoor Adventure"], "Outdoor adventurers", 25),
        line("Weekend", ["TRAVEL Weekend"], "Weekend getaways", 25),
    ])
    check("four lines, four rows", len(rows) == 4, [r["Tactic"] for r in rows])
    check("no trim note on any of them", not [n for n in notes if "at most 2" in n], notes)

    print("\nA line with no segments listed (every frozen fixture) is untouched")
    rows, notes = resolve([{"product": "premion_streaming_tv", "audience_track":
                            "Adults 35+, homeowners, $150K+ income", "allocation": {"percent_of_total": 100}}])
    check("Targeting as drafted", rows and rows[0]["Targeting"] == "Adults 35+, homeowners, $150K+ income",
          rows)

    print("\nThe shared group check")
    check("MAX_AND_SEGMENTS is 2", tg.MAX_AND_SEGMENTS == 2)
    over = tg.new_group(["A", "B", "C"], op="AND")
    check("a 3-term AND is a violation", tg.and_limit_violations([over]) == [over])
    check("a 2-term AND, a 3-term OR and a single custom are fine",
          not tg.and_limit_violations([tg.new_group(["A", "B"], op="AND"),
                                       tg.new_group(["A", "B", "C"], op="OR"),
                                       tg.new_group(["CUSTOM Homeowner AND Auto or Home Insurance"])]))

    print("\nThe draft prompt states the rule and drops the one-custom limit")
    prompt = app.build_draft_prompt("HVAC company, homeowners with older homes and $100K+ income.")
    check("asks for at most two segments per line", "at most TWO catalog segments" in prompt)
    check("a custom segment counts as one", "CUSTOM segment counts as one" in prompt)
    check("no one-custom limit", "At most ONE non-RFP" not in prompt)
    check("the schema shows a per-line segments list", '"segments": ["exact catalog name"]' in prompt)

    print(f"\n{'FAILED: ' + str(len(failures)) if failures else 'All checks passed.'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
