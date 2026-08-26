"""The media plan mirrors the avails table's grouping.

    python tests/test_plan_follows_avails.py

The bug: the plan's Geo default always joined the target markets, so three
markets broken out separately in the avails table still produced ONE plan line
reading "Philadelphia, Atlanta". That default was written before the avails
table had a separate-vs-combined choice to follow.

There is deliberately no second toggle for the plan. The avails grouping is
the default; merge and split on plan lines is the independent control, and it
covers the mixed cases a toggle can't express. Until that lands, editing a
row's Geo and deleting a row is the crude version.

Driven through the real form, because what is being asserted is what the grid
actually holds after a rerun -- not what a helper returns in isolation. A
rerun path quietly re-seeding rows is a bug this project has shipped twice.
"""
import os
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

COL = app.AVAILS_COLUMN_MONTHLY
ROOT = Path(__file__).resolve().parent.parent
THREE = ["Philadelphia", "Atlanta", "Denver"]

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + detail if detail else ''}")
        failures.append(label)


def avails(*geos, audience="Homeowners"):
    return [{"Audience": audience, "Geo": g, COL: 1_000_000} for g in geos]


def included_groups(rows):
    """`rows` (the flat avails shape) -> targeting groups with every one
    ticked into the plan. This suite is about what the PLAN does with the
    avails table's GROUPING once a line exists -- not about the
    include-in-plan default itself, which is test_group_plan_selection.py's
    job -- so every test here starts from "everything is on the plan",
    matching what this suite always asserted before groups owned their own
    inclusion flag.
    """
    groups = app.tg.seed_rows_to_groups(rows, COL)
    for group in groups:
        group["include_in_plan"] = True
    return groups


def run(avails_rows=None, extra=None, then=None):
    """Render the form, optionally change something, and return option 0."""
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (this file's own
    # subject) on market+flight.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    if avails_rows is not None:
        at.session_state["avails_seed_rows"] = avails_rows
        at.session_state["targeting_groups"] = included_groups(avails_rows)
    for key, value in (extra or {}).items():
        at.session_state[key] = value
    at.run()
    if at.exception:
        return None, at.exception[0].message[:250]
    if then:
        for key, value in then.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return None, at.exception[0].message[:250]
    return at.session_state["plan_options"][0], None


def geos_of(option, tactic="Premion Streaming TV"):
    return [r["Geo"] for r in option["rows"] if r.get("Tactic") == tactic]


def main():
    print("the helper that decides the plan's lines")
    geos = lambda pairs: [g for _, g in pairs]
    check("separate avails rows give one line each",
          geos(app.plan_lines_from_avails(avails(*THREE), "DC DMA")) == THREE)
    check("a combined avails row gives one joined geo",
          geos(app.plan_lines_from_avails(
              avails("Philadelphia, Atlanta"), "DC DMA")) == ["Philadelphia, Atlanta"])
    check("an empty table falls back to the single default",
          geos(app.plan_lines_from_avails([], "DC DMA")) == ["DC DMA"])
    check("blank rows are ignored rather than seeding a blank line",
          geos(app.plan_lines_from_avails(
              [{"Geo": "", "Audience": ""}, {"Geo": "Denver"}], "DC")) == ["Denver"])

    print("\naudience > geo: nothing that shares a market is merged")
    two_by_three = (avails(*THREE, audience="Homeowners")
                    + avails(*THREE, audience="In-market for windows"))
    pairs = app.plan_lines_from_avails(two_by_three, "DC DMA")
    check("2 audiences x 3 markets is six lines", len(pairs) == 6, len(pairs))
    check("ordered audience-major, each audience across every market",
          pairs == [("Homeowners", g) for g in THREE]
                 + [("In-market for windows", g) for g in THREE], pairs)
    check("the same market under two audiences is never collapsed",
          [g for _, g in pairs].count("Denver") == 2, pairs)

    print("\na combined AND audience stays one audience")
    combined = avails("Denver", audience="Homeowners with income $50K+")
    check("one line, not two",
          len(app.plan_lines_from_avails(combined, "DC DMA")) == 1)
    check("and it keeps the whole audience as written",
          app.plan_lines_from_avails(combined, "DC")[0][0]
          == "Homeowners with income $50K+")

    print("\nthree markets separate -> three lines")
    option, err = run(avails(*THREE))
    check("the form renders", not err, err)
    if option:
        got = geos_of(option)
        check("one Premion line per market", len(got) == 3, got)
        check("in avails-table order", got == THREE, got)
        check("no joined-geo line survives",
              not any("," in g for g in got), got)

    print("\n2 audiences x 3 markets -> 6 lines, through the real form")
    option, err = run(avails(*THREE, audience="Homeowners")
                      + avails(*THREE, audience="In-market for windows"))
    check("the form renders", not err, err)
    if option:
        premion = [r for r in option["rows"]
                   if r.get("Tactic") == "Premion Streaming TV"]
        check("six Premion lines, one per avails row", len(premion) == 6,
              len(premion))
        check("audience-major: the first three are one audience",
              len({r["Targeting"] for r in premion[:3]}) == 1,
              [r["Targeting"] for r in premion[:3]])
        check("...and the next three are the other",
              len({r["Targeting"] for r in premion[3:]}) == 1
              and premion[0]["Targeting"] != premion[3]["Targeting"],
              [r["Targeting"] for r in premion[3:]])
        check("each audience runs across every market",
              [r["Geo"] for r in premion[:3]] == THREE
              and [r["Geo"] for r in premion[3:]] == THREE,
              [r["Geo"] for r in premion])
        denver = [r for r in premion if r["Geo"] == "Denver"]
        check("Denver appears twice, once per audience, never merged",
              len(denver) == 2
              and denver[0]["Targeting"] != denver[1]["Targeting"], denver)
        check("each line carries its own audience as its Targeting",
              {r["Targeting"] for r in premion}
              == {"Homeowners", "In-market for windows"},
              sorted({r["Targeting"] for r in premion}))

    print("\nthree markets combined -> one line with the joined geo")
    option, err = run(avails("Philadelphia, Atlanta, Denver"))
    check("the form renders", not err, err)
    if option:
        got = geos_of(option)
        check("a single Premion line", len(got) == 1, got)
        check("carrying every market", got and all(m in got[0] for m in THREE), got)

    print("\nan edited line's Geo survives a market change")
    # Seed three markets, hand-edit one line's Geo, then change a shared field
    # (the audience) -- which is what triggers the re-seed path.
    option, err = run(
        avails(*THREE),
        then={"audience_text": "Adults 25-54, homeowners, in-market"})
    check("the form renders after the shared-field change", not err, err)
    if option:
        got = geos_of(option)
        check("per-market lines are NOT collapsed onto one joined geo",
              len(got) == 3 and not any("," in g for g in got), got)
        check("and each keeps its own market", sorted(got) == sorted(THREE), got)

    # An explicitly dirty row must keep a hand-typed Geo the plan doesn't know.
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (this file's own
    # subject) on market+flight.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.session_state["avails_seed_rows"] = avails(*THREE)
    at.session_state["targeting_groups"] = included_groups(avails(*THREE))
    at.run()
    opt = at.session_state["plan_options"][0]
    if opt["rows"]:
        opt["rows"][0]["Geo"] = "Philadelphia metro only"
        opt["dirty"][0] = True
        at.session_state["audience_text"] = "Adults 25-54, homeowners"
        at.run()
        after = at.session_state["plan_options"][0]["rows"][0]["Geo"]
        check("a hand-edited Geo is never re-seeded, even off-plan",
              after == "Philadelphia metro only", after)

    print("\nTotal TV: broadcast Geo still differs from streaming")
    # resolve_row_defaults holds the broadcast row's call-sign Geo against any
    # re-seed. Asserted here as well as in test_geo_defaults, because this
    # change moved what the re-seed passes in.
    broadcast_tactic = app.BROADCAST_TACTIC_MARKER
    existing = {"Tactic": broadcast_tactic, "Geo": "Washington DC DMA",
                "Targeting": "42 spots across News and Prime"}
    check("the tactic is recognised as broadcast", app.is_broadcast_row(existing))
    held = app.resolve_row_defaults(broadcast_tactic, "Philadelphia",
                                    "Homeowners", "Oct", current=existing)
    check("broadcast keeps its call-sign DMA when the plan is per-market",
          held["Geo"] == "Washington DC DMA", held["Geo"])
    streaming = app.resolve_row_defaults("Premion Streaming TV", "Philadelphia",
                                         "Homeowners", "Oct")
    check("...while a streaming line on the same plan takes its market",
          streaming["Geo"] == "Philadelphia" and streaming["Geo"] != held["Geo"],
          f"streaming={streaming['Geo']!r} broadcast={held['Geo']!r}")

    print("\nCampaign Specs is unchanged by any of this")
    check("Geography still lists markets one per line",
          app.geography_default_text(THREE, "DC DMA") == "\n".join(THREE),
          app.geography_default_text(THREE, "DC DMA"))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The plan follows the avails table's grouping.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
