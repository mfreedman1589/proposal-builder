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


def run(avails_rows=None, extra=None, then=None):
    """Render the form, optionally change something, and return option 0."""
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    if avails_rows is not None:
        at.session_state["avails_seed_rows"] = avails_rows
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
    print("the helper that decides the plan's grouping")
    check("separate avails rows give one geo each",
          app.plan_geos_from_avails(avails(*THREE), "DC DMA") == THREE)
    check("a combined avails row gives one joined geo",
          app.plan_geos_from_avails(avails("Philadelphia, Atlanta"), "DC DMA")
          == ["Philadelphia, Atlanta"])
    check("an empty table falls back to the single default",
          app.plan_geos_from_avails([], "DC DMA") == ["DC DMA"])
    check("blank geos are ignored rather than seeding a blank line",
          app.plan_geos_from_avails([{"Geo": ""}, {"Geo": "Denver"}], "DC")
          == ["Denver"])
    check("the same market twice (two audiences) is still one place to buy",
          app.plan_geos_from_avails(
              avails("Denver") + avails("Denver", audience="In-market"), "DC")
          == ["Denver"])

    print("\nthree markets separate -> three lines")
    option, err = run(avails(*THREE))
    check("the form renders", not err, err)
    if option:
        got = geos_of(option)
        check("one Premion line per market", len(got) == 3, got)
        check("in avails-table order", got == THREE, got)
        check("no joined-geo line survives",
              not any("," in g for g in got), got)

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
    at.session_state["avails_seed_rows"] = avails(*THREE)
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
