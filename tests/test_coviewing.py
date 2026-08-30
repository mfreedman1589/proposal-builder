"""Co-viewing coefficient (BACKLOG.md): an opt-in projection of person-level
exposure alongside household impressions on CTV lines.

Eligibility is an explicit allowlist (`app.is_coviewing_eligible_row`) --
Premion Streaming TV and Live Sports packages only, confirmed with the user.
Everything else (Streaming Retargeting, broadcast, a flat fee) shows a dash,
never a guessed number. A merged/aggregate figure is deliberately NOT
computed here the way % SOV's is -- per-row "+N" and one plan-wide effective
CPM (in the footnote, never a table column) are the whole feature.

Two independent AppTest instances (OFF, ON), same reasoning as
test_share_of_voice.py: `show_coviewing` is a widget-keyed checkbox, and
setting it before an instance's first `.run()` (rather than mid-session)
sidesteps the AppTest-only session_state/rerun hazard documented in
DECISIONS.md entirely.

    python tests/test_coviewing.py
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
db.fetch_setting = lambda key: (None, "stubbed -- use the fallback")

import app  # noqa: E402
import assembly  # noqa: E402

sys.path.insert(0, str(REPO / "tests"))
from test_form_state import load_drafted_form  # noqa: E402

MULTIPLIER = app.FALLBACK_COVIEWING["multiplier"]  # 1.4, confirmed with the user

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


def build_option():
    """One option covering every eligibility case: a Premion Streaming TV
    line and a Live Sports line (both eligible), a Streaming Retargeting
    line and a broadcast line (both ineligible, never string-matched --
    excluded purely by not matching either allowed prefix), and a flat fee
    (no impressions to project from). Monthly breakout, so each row's own
    Impressions is its monthly figure directly -- no month-count dependency
    to control for.
    """
    ctv_row = {"Tactic": "Premion Streaming TV", "Flight": "", "Geo": "Washington, DC DMA",
               "Targeting": "Homeowners 35+", "Impressions": "200000", "Cost": "6400", "CPM": "32.00"}
    sports_row = {"Tactic": app.sport_product_label("nfl_reg"), "Flight": "", "Geo": "Washington, DC DMA",
                  "Targeting": app.LIVE_SPORTS_TARGETING, "Impressions": "100000", "Cost": "5000", "CPM": "50.00"}
    retargeting_row = {"Tactic": "Streaming Retargeting - Display", "Flight": "", "Geo": "Washington, DC DMA",
                       "Targeting": "Retarget Exposed CTV Viewers", "Impressions": "50000",
                       "Cost": "275", "CPM": "5.50"}
    broadcast_row = {"Tactic": f"Cable News {app.BROADCAST_TACTIC_MARKER}", "Flight": "", "Geo": "Washington, DC DMA",
                     "Targeting": "Adults 25-54", "Impressions": "80000", "Cost": "440", "CPM": "5.50"}
    flat_fee_row = {"Tactic": "Commercial Production (:30 spot)", "Flight": "", "Geo": "Washington, DC DMA",
                    "Targeting": "", "Impressions": "0", "Cost": "1200", "CPM": "0",
                    "Type": app.ROW_TYPE_FLAT_FEE}
    rows = [ctv_row, sports_row, retargeting_row, broadcast_row, flat_fee_row]
    # Driven from Cost, not the default Impressions -- an Impressions-driven
    # row gets its Cost RECOMPUTED from Impressions x CPM the moment the
    # grid reconciles it. Driving from Cost instead keeps the hand-set
    # dollar figures above exactly what this test's own expected-value math
    # assumes. (FLOW_REWORK_PLAN.md Phase 4b: recompute_row no longer has a
    # markup term at all -- every row is net on the grid regardless of the
    # gross-up checkbox -- but the Cost-vs-Impressions driver hazard this
    # guards against is unrelated to that and still real.)
    return app.new_plan_option("Option A", rows, driver=[app.DRIVER_COST] * len(rows))


def expected_effective_cpm():
    """The same blended-across-eligible-and-ineligible-lines math
    effective_cpm_with_coviewing does, computed independently here from the
    row figures above -- not by calling the function under test."""
    # (monthly_cost, monthly_impressions, eligible)
    rate_rows = [(6400, 200_000, True), (5000, 100_000, True),
                 (275, 50_000, False), (440, 80_000, False)]
    cost = sum(c for c, _, _ in rate_rows)
    impressions = sum(i * (MULTIPLIER if elig else 1) for _, i, elig in rate_rows)
    return cost / impressions * 1000


def expected_footnote_body():
    """The EXACT string app.py's coviewing_footnote should compose --
    multiplier, source, and the effective CPM -- matched with == rather
    than a handful of substring checks.

    Substring checks (`"1.4x" in footnote`) are exactly the assertion shape
    that let the real doubled-"Co-viewing:"-label bug through: `"X" in text`
    is also true of "X X". An exact match on the whole composed string is
    the only way a second, redundant copy of any one piece (the label, the
    multiplier, the source, the CPM) can't slip past this test the way it
    slipped past a plain presence check the first time.
    """
    expected_cpm = expected_effective_cpm()
    return (f"{MULTIPLIER:g}x factor applied ({app.FALLBACK_COVIEWING['source']}) "
            f"-- effective CPM at estimated exposure ${expected_cpm:,.2f}.")


def seed_known_state(at, show_coviewing):
    """A drafted baseline (so Generate has everything else it needs), with
    the plan replaced by the known, hand-built option above. Everything here
    is set before this instance's first `.run()` -- no widget for any of
    these keys exists yet, so there's nothing to shadow.
    """
    load_drafted_form(at)
    # Forced regardless of what the drafted fixture set: the gross-up
    # checkbox now only affects compute_plan_totals' DISPLAY figures (via
    # row_markup), never the stored row itself -- but expected_effective_cpm()
    # above hand-computes its own totals at markup 1.0, so this stays
    # explicit rather than relying on the checkbox's own default.
    at.session_state["agency_gross_up"] = False
    at.session_state["plan_options"] = [build_option()]
    if show_coviewing:
        at.session_state["show_coviewing"] = True


def generate(at):
    """Click Generate through a spy on BOTH assembly.personalize (to capture
    the fill_data payload) and assembly._finish_media_plan_slide (to capture
    the real, filled slide -- the real deck is still built either way, the
    spies only observe)."""
    real_personalize = assembly.personalize
    real_finish = assembly._finish_media_plan_slide
    captured = {"slides": []}

    def spy_personalize(prs, fill_data):
        captured["fill_data"] = fill_data
        return real_personalize(prs, fill_data)

    def spy_finish(prepared, compressed):
        result = real_finish(prepared, compressed)
        captured["slides"].append(prepared["slide"])
        return result

    assembly.personalize = spy_personalize
    assembly._finish_media_plan_slide = spy_finish
    try:
        buttons = [b for b in at.button if b.label == "Generate proposal"]
        if not buttons:
            return None, at
        buttons[0].click().run()
    finally:
        assembly.personalize = real_personalize
        assembly._finish_media_plan_slide = real_finish
    return captured, at


def rows_by_tactic(fill_data):
    return {r["tactic"]: r for r in fill_data["media_plan_options"][0]["rows"]}


def main():
    print("=" * 78)
    print("SCENARIO  co-viewing is OFF by default -- no column, no footnote")
    print("=" * 78)
    at_off = new_app()
    seed_known_state(at_off, show_coviewing=False)
    at_off.run()
    check("no exception after seeding the known plan", not at_off.exception,
          at_off.exception[0].message[:400] if at_off.exception else "")

    captured, at_off = generate(at_off)
    check("Generate runs without raising", not at_off.exception,
          at_off.exception[0].message[:400] if at_off.exception else "")
    fill_data = captured.get("fill_data") if captured else None
    check("fill_data was captured", fill_data is not None)
    if fill_data is not None:
        option = fill_data["media_plan_options"][0]
        check("show_coviewing is False in the payload", option["show_coviewing"] is False)
        check("no coviewing footnote", option.get("coviewing_footnote") is None)
        by_tactic = rows_by_tactic(fill_data)
        for tactic, row in by_tactic.items():
            check(f"{tactic!r} carries no coviewing figure when the toggle is off",
                  row["coviewing"] == "--", row["coviewing"])
    if captured and captured["slides"]:
        table = assembly._find_table_shape(captured["slides"][0]).table
        header_texts = [cell.text_frame.text for cell in table.rows[0].cells]
        check("no Monthly Coviewing column on the real generated slide",
              "MONTHLY COVIEWING" not in header_texts, header_texts)
        terms = assembly._text_shape_containing(captured["slides"][0], assembly._TERMS_ANCHOR)
        check("the Terms & Conditions box has no Co-viewing line added",
              terms is None or "Co-viewing:" not in terms.text_frame.text,
              terms.text_frame.text if terms else None)

    print("\n" + "=" * 78)
    print("SCENARIO  co-viewing ON -- eligible/ineligible lines, footnote, real slide")
    print("=" * 78)
    at_on = new_app()
    seed_known_state(at_on, show_coviewing=True)
    at_on.run()
    check("no exception after seeding the known plan", not at_on.exception,
          at_on.exception[0].message[:400] if at_on.exception else "")

    captured, at_on = generate(at_on)
    check("Generate runs without raising", not at_on.exception,
          at_on.exception[0].message[:400] if at_on.exception else "")
    fill_data = captured.get("fill_data") if captured else None
    check("fill_data was captured", fill_data is not None)

    if fill_data is not None:
        option = fill_data["media_plan_options"][0]
        check("show_coviewing is True in the payload", option["show_coviewing"] is True)
        by_tactic = rows_by_tactic(fill_data)

        check("Premion Streaming TV gets +80,000 (200,000 x 0.4)",
              by_tactic["Premion Streaming TV"]["coviewing"] == "+80,000",
              by_tactic["Premion Streaming TV"]["coviewing"])
        sports_tactic = app.sport_product_label("nfl_reg")
        check("the Live Sports package gets +40,000 (100,000 x 0.4)",
              by_tactic[sports_tactic]["coviewing"] == "+40,000",
              by_tactic[sports_tactic]["coviewing"])
        check("Streaming Retargeting gets no coviewing figure",
              by_tactic["Streaming Retargeting - Display"]["coviewing"] == "--",
              by_tactic["Streaming Retargeting - Display"]["coviewing"])
        broadcast_tactic = f"Cable News {app.BROADCAST_TACTIC_MARKER}"
        check("the broadcast line gets no coviewing figure",
              by_tactic[broadcast_tactic]["coviewing"] == "--",
              by_tactic[broadcast_tactic]["coviewing"])
        check("the flat fee gets no coviewing figure",
              by_tactic["Commercial Production (:30 spot)"]["coviewing"] == "--",
              by_tactic["Commercial Production (:30 spot)"]["coviewing"])

        footnote = option.get("coviewing_footnote")
        check("a coviewing footnote exists (the option has eligible lines)", bool(footnote), footnote)
        # Exact match, not three substring checks -- see expected_footnote_body's
        # own docstring for why "in" is exactly the wrong tool here.
        check("the footnote exactly matches multiplier + source + effective CPM, "
              "with no extra or duplicated text",
              footnote == expected_footnote_body(), (footnote, expected_footnote_body()))

    if captured and captured["slides"]:
        table = assembly._find_table_shape(captured["slides"][0]).table
        header_texts = [cell.text_frame.text for cell in table.rows[0].cells]
        check("the real generated slide has a Monthly Coviewing column",
              "MONTHLY COVIEWING" in header_texts, header_texts)
        check("Impressions sits immediately left of Monthly Coviewing",
              "IMPRESSIONS" in header_texts[header_texts.index("MONTHLY COVIEWING") - 1].upper(),
              header_texts)
        terms = assembly._text_shape_containing(captured["slides"][0], assembly._TERMS_ANCHOR)
        # Exact suffix match against "Co-viewing: " + the same expected body
        # checked above -- not `"Co-viewing:" in text`. A plain substring
        # check is exactly the assertion shape that let the real doubled-
        # label bug through ("Co-viewing: Co-viewing: ..." still contains
        # "Co-viewing:"); an exact match on the whole appended paragraph
        # catches a doubled label, a doubled body, or either supplied twice
        # by two different producers -- not just one of those shapes.
        expected_paragraph = f"Co-viewing: {expected_footnote_body()}"
        check("the Terms & Conditions box's last paragraph is exactly the expected "
              "Co-viewing line -- not missing, not doubled, not appended twice",
              terms is not None and terms.text_frame.text.endswith(expected_paragraph)
              and terms.text_frame.text.count("Co-viewing:") == 1,
              terms.text_frame.text if terms else None)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Co-viewing stays off unless a rep opts in; when it's on, only Premion Streaming TV "
          "and Live Sports lines get a projected additional-impressions figure, the effective "
          "CPM in the footnote correctly blends eligible and ineligible lines, and the real "
          "generated slide carries both the new column and the citation.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
