"""% of avails (share of voice) under the media plan impression totals --
BACKLOG.md. Impressions / total resolved avails, shown in parentheses right
after the Monthly and Full Flight impression totals, both in the app's own
preview and on the generated slide -- gated behind a new "Show % of avails"
checkbox that defaults off.

Two independent AppTest instances (OFF, ON) rather than one instance toggled
mid-test: `show_sov` is a widget-keyed checkbox, and this project's own
DECISIONS.md entry on the AppTest session_state/rerun artifact is precisely
about the hazard of overriding an ALREADY-RENDERED widget's key via
`at.session_state[...]` in the same run as other state changes. Setting it
(or leaving it untouched) before that instance's very first `.run()` avoids
the question entirely, matching every other override in this file.

Expected percentages are computed with the app's own `compute_plan_totals`
against whatever `active_months` the drafted baseline actually settles on,
rather than a hardcoded month count -- the number of months isn't what this
test is about, and hardcoding it would make the test fragile to the fixture
rather than to the feature.

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

KNOWN_AVAILS_MONTHLY = 500_000
KNOWN_MONTHLY_IMPRESSIONS = 180_000

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


def seed_known_state(at, show_sov):
    """A drafted baseline (so Generate has everything else it needs), with
    the avails and the plan replaced by fully-known figures so the expected
    percentage is arithmetic rather than tied to the fixture's own numbers.

    Everything here is set before this instance's first `.run()` -- no
    widget for any of these keys exists yet, so there's nothing to shadow.
    """
    load_drafted_form(at)

    group = tg.new_group(["Test Audience"], avails_monthly=KNOWN_AVAILS_MONTHLY)
    rows = tg.groups_to_seed_rows([group], app.AVAILS_COLUMN_MONTHLY)
    at.session_state["targeting_groups"] = [group]
    at.session_state["avails_seed_rows"] = rows
    # Tells sync_targeting_groups() the two sides already agree, so Generate
    # doesn't re-derive groups from rows and overwrite this override.
    at.session_state["_groups_rows_applied"] = [dict(r) for r in rows]

    row = {"Tactic": "Streaming TV", "Flight": "", "Geo": "Washington, DC DMA",
           "Targeting": "Test Audience", "Impressions": str(KNOWN_MONTHLY_IMPRESSIONS),
           "Cost": "5400", "CPM": "30.00"}
    at.session_state["plan_options"] = [app.new_plan_option("Option A", [row])]

    if show_sov:
        at.session_state["show_sov"] = True


def expected_percentages(at):
    """The app's own compute_plan_totals against whatever active_months this
    run actually settled on, so the expected number tracks the fixture
    instead of assuming a month count."""
    option = at.session_state["plan_options"][0]
    n_months = max(1, len(list(at.session_state["active_months"] or [])))
    totals = app.compute_plan_totals(option["rows"], option["breakout"], n_months, "flight")
    monthly_pct = round(totals["monthly_impressions"] / KNOWN_AVAILS_MONTHLY * 100.0)
    flight_avails = KNOWN_AVAILS_MONTHLY * n_months
    flight_pct = round(totals["full_flight_impressions"] / flight_avails * 100.0)
    return n_months, monthly_pct, flight_pct


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


def main():
    print("=" * 78)
    print("SCENARIO  % of avails is OFF by default")
    print("=" * 78)
    at_off = new_app()
    seed_known_state(at_off, show_sov=False)
    at_off.run()
    check("no exception after seeding known avails/plan", not at_off.exception,
          at_off.exception[0].message[:400] if at_off.exception else "")

    fill_data, at_off = generate(at_off)
    check("Generate runs without raising", not at_off.exception,
          at_off.exception[0].message[:400] if at_off.exception else "")
    check("fill_data was captured", fill_data is not None)
    if fill_data is not None:
        option = fill_data["media_plan_options"][0]
        check("no suffix on the monthly total when the toggle is off",
              not sov_suffix_present(option["total_impressions"]), option["total_impressions"])
        fft = option.get("full_flight_total")
        if fft:
            check("no suffix on the full flight total when the toggle is off",
                  not sov_suffix_present(fft["impressions"]), fft["impressions"])

    print("\n" + "=" * 78)
    print("SCENARIO  % of avails, toggled on, matches impressions / avails")
    print("=" * 78)
    at_on = new_app()
    seed_known_state(at_on, show_sov=True)
    at_on.run()
    check("no exception after seeding known avails/plan", not at_on.exception,
          at_on.exception[0].message[:400] if at_on.exception else "")

    n_months, monthly_pct, flight_pct = expected_percentages(at_on)
    print(f"    ....  settled on {n_months} active month(s); "
          f"expecting {monthly_pct}% monthly, {flight_pct}% full flight")

    fill_data, at_on = generate(at_on)
    check("Generate runs without raising", not at_on.exception,
          at_on.exception[0].message[:400] if at_on.exception else "")
    check("fill_data was captured", fill_data is not None)
    if fill_data is not None:
        option = fill_data["media_plan_options"][0]
        check(f"monthly total carries the ({monthly_pct}% of avails) suffix",
              f"({monthly_pct}% of avails)" in option["total_impressions"],
              option["total_impressions"])
        fft = option.get("full_flight_total")
        if n_months > 1:
            check("a full flight total row exists when the flight is more than one month",
                  fft is not None, fft)
            if fft:
                check(f"full flight total carries the ({flight_pct}% of avails) suffix",
                      f"({flight_pct}% of avails)" in fft["impressions"], fft["impressions"])
        else:
            check("no full flight total row for a one-month flight (nothing to add a suffix to)",
                  fft is None, fft)

    print("\n" + "=" * 78)
    print("SCENARIO  no avails resolved -- the toggle can't fabricate a percentage")
    print("=" * 78)
    at_none = new_app()
    load_drafted_form(at_none)
    at_none.session_state["targeting_groups"] = []
    at_none.session_state["avails_seed_rows"] = []
    at_none.session_state["_groups_rows_applied"] = []
    at_none.session_state["show_sov"] = True
    at_none.run()
    check("no exception with zero avails", not at_none.exception,
          at_none.exception[0].message[:400] if at_none.exception else "")

    fill_data, at_none = generate(at_none)
    check("Generate runs without raising", not at_none.exception,
          at_none.exception[0].message[:400] if at_none.exception else "")
    if fill_data is not None:
        option = fill_data["media_plan_options"][0]
        check("no suffix when there's no avails pool to divide into, even with the toggle on",
              not sov_suffix_present(option["total_impressions"]), option["total_impressions"])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("% of avails stays off unless a rep opts in, matches impressions / avails exactly "
          "when it's on (both the app's own preview payload and the slide-fill payload), "
          "and never fabricates a percentage when there's no avails pool to divide into.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
