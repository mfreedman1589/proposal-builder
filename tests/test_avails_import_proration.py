"""An avails-PDF group's own stated impressions must never be prorated
against an unrelated month count -- found testing the real Plaza Motors
avail (RFPID-266583, a 14-day partial-month flight: 9/17-9/30).

**First cut of this fix (superseded, kept here for the record):**
`_finish_avails_import` computed `avails_monthly` as `full_flight /
n_months`, where `n_months` came from `form_flight_months()` -- the CURRENT
form flight, read *before* the document's own flight_start/flight_end land.
On a fresh form that's just DEFAULT_FLIGHT_START..DEFAULT_FLIGHT_END (Sep-
Nov, 3 calendar months), unrelated to the document's real 14-day flight, so
1,707,337 was divided by 3 and 569,112 got baked in permanently. Fixing
`n_months` to come from the EFFECTIVE (about-to-land) flight made the
reported case pass -- but only because that proposal's plan flight ended up
IDENTICAL to the document's own. Dividing the raw full-flight total by a
month count, with no day-level term at all, is still wrong the moment the
plan's actual flight differs from the document's -- a wider plan flight
covering the same document doesn't mean "the same total spread across more
months," it means more total avails, since the inventory keeps generating
at whatever rate the document's own flight measured.

**The real fix: a daily rate.** `avails_daily_rate` divides the document's
own full-flight impressions by the number of days in the DOCUMENT's own
stated flight (1,707,337 / 14 for the A35-64 line). `avails_monthly_from_
daily_rate` then reapplies that rate to however many days the PLAN LINE
actually runs -- never the document's own day count -- and divides by the
calendar months the PLAN flight touches, matching the `monthly * n_months`
convention every other avails consumer already assumes. When the plan
flight and the document's own flight are identical (the reported case),
this reduces to exactly the same number the first-cut fix produced -- see
`check_same_flight` below. It only diverges, correctly, once the two
flights differ -- see `check_flight_mismatch`, the case the first cut could
never have handled, because it only fixed the flight the DIVISION used,
never taught the app that the total itself has to scale with the plan's own
day count.

Rule this guards: an avails figure reduces to a per-day rate scoped to the
document's own flight, and that rate applies to whatever flight the plan
line actually runs -- never a coincidence of month counts.

**Second real document, a different shape: Capital Media (RFPID-266994)**
carries FOUR monthly exception rows for one audience/geography (9/21-9/30,
October, November, 12/1-12/20) instead of one row spanning the whole
flight -- `avails_pdf_import.py`'s own parser already sums these into one
group's `impressions` total (1 group, `row_count=4`), verbatim, never
re-derived -- confirmed directly, not assumed. `avails_daily_rate` divides
that summed total by the document's OWN flight span (9/21-12/20, 91 days),
which recovers the real per-day rate (36,499) exactly here because the
document happens to be internally consistent, but the formula never assumes
that -- it only ever reads the total and the total day count, regardless of
how many rows contributed to either.

That document is also why `avails_monthly` alone isn't enough for an EXACT
full-flight tie-out once a flight spans months of different lengths:
storing one rounded monthly figure and multiplying it back out by n_months
drifts by a few impressions (September/December here are partial months,
October/November aren't). `avails_full_flight_from_daily_rate` /
`matched_avails_full_flight_for_row` compute the full-flight figure
directly from the daily rate instead, so the SOV denominator on the slide
ties to the document's own 3,321,409 exactly -- see
`check_capital_media_multi_row`.

Uses the real PDFs at the repo root (gitignored); SKIPs without them.

    python tests/test_avails_import_proration.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                 # noqa: E402
import avails_pdf_import   # noqa: E402

PLAZA_MOTORS = REPO / ("Premion Media Plan_RFPID-266583_TBC, Inc - Trahan, "
                       "Burden & Charles_Plaza Motors Group_8-19-2026--ver0.pdf")
CAPITAL_MEDIA = REPO / "Premion Media Plan_RFPID-266994_Capital Media_Undisclosed Advertiser_8-25-2026--ver0.pdf"

# The document's own stated, ground-truth figures (Product Summary table) --
# never anything the app itself recomputes.
A3564_LUXURY_IMPRESSIONS = 1707337
M3564_LUXURY_IMPRESSIONS = 1126832
# The document's own stated flight -- 14 days, the divisor for the daily rate.
AVAIL_FLIGHT_START = date(2026, 9, 17)
AVAIL_FLIGHT_END = date(2026, 9, 30)

# Capital Media's four monthly exception rows, verbatim from the document's
# own Product Details table -- ground truth for "the parser summed these
# right," never recomputed.
CAPITAL_MEDIA_ROWS = [
    (date(2026, 9, 21), date(2026, 9, 30), 364990),
    (date(2026, 10, 1), date(2026, 10, 31), 1131469),
    (date(2026, 11, 1), date(2026, 11, 30), 1094970),
    (date(2026, 12, 1), date(2026, 12, 20), 729980),
]
CAPITAL_MEDIA_TOTAL = 3321409
CAPITAL_MEDIA_FLIGHT_START = date(2026, 9, 21)
CAPITAL_MEDIA_FLIGHT_END = date(2026, 12, 20)
CAPITAL_MEDIA_DAILY_RATE = 36499.0

failures = []
skipped = False


class _StubSt:
    """Same technique tests/test_draft_regression.py uses for
    apply_draft_to_form: apply_avails_import/_finish_avails_import only
    touch st.session_state, so a plain dict stands in for Streamlit -- lets
    this test pre-seed an ALREADY-DIFFERENT flight (the "conflict" state a
    real rep-set flight would leave in place) without needing to fight
    _avails_import_field's own precedence rules just to get there."""

    def __init__(self, preset=None):
        self.session_state = dict(preset or {})
        self.secrets = {}


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


def ss(at, key, default=None):
    return at.session_state[key] if key in at.session_state else default


def real_groups(at):
    groups = ss(at, "targeting_groups", [])
    return [g for g in groups if g["terms"]]


def check_daily_rate_pure():
    """The raw formula, offline, no import/session involved -- reports the
    daily rate explicitly so it's checkable against the document by hand:
    1,707,337 / 14 and 1,126,832 / 14."""
    print("avails_daily_rate / avails_monthly_from_daily_rate -- pure formula")
    rate_a = app.avails_daily_rate(A3564_LUXURY_IMPRESSIONS, AVAIL_FLIGHT_START, AVAIL_FLIGHT_END)
    rate_m = app.avails_daily_rate(M3564_LUXURY_IMPRESSIONS, AVAIL_FLIGHT_START, AVAIL_FLIGHT_END)
    print(f"    ....  A35-64 x Luxury daily rate: {A3564_LUXURY_IMPRESSIONS:,} / 14 days = {rate_a:,.4f} / day")
    print(f"    ....  M35-64 x Luxury daily rate: {M3564_LUXURY_IMPRESSIONS:,} / 14 days = {rate_m:,.4f} / day")
    check("A35-64 daily rate matches hand math (1,707,337 / 14)",
          abs(rate_a - 121952.642857) < 0.01, rate_a)
    check("M35-64 daily rate matches hand math (1,126,832 / 14)",
          abs(rate_m - 80488.0) < 0.01, rate_m)

    # Same-flight: plan flight IDENTICAL to the avail's own -- must reduce to
    # the document's stated figure exactly, the same as the first-cut fix.
    same_monthly, same_n_months = app.avails_monthly_from_daily_rate(
        rate_a, AVAIL_FLIGHT_START, AVAIL_FLIGHT_END)
    check("same-flight: avails_monthly is the document's own stated figure",
          same_monthly == A3564_LUXURY_IMPRESSIONS, same_monthly)
    check("same-flight: touches exactly 1 calendar month", same_n_months == 1, same_n_months)

    # Divergent flight: plan sells the WHOLE of September (9/1-9/30, 30 days),
    # wider than the avail's own 14-day pull. Full-flight avails must scale
    # UP proportionally to the wider window, not stay pinned to the
    # document's own 1,707,337 / 1,126,832.
    wide_start, wide_end = date(2026, 9, 1), date(2026, 9, 30)
    wide_a, wide_n_months = app.avails_monthly_from_daily_rate(rate_a, wide_start, wide_end)
    wide_m, _ = app.avails_monthly_from_daily_rate(rate_m, wide_start, wide_end)
    print(f"    ....  plan flight widened to 9/1-9/30 (30 days): "
          f"A35-64 -> {rate_a:,.4f} x 30 = {rate_a * 30:,.2f} -> {wide_a:,}")
    check("flight-mismatch: A35-64 avails_monthly scales to the WIDER plan flight "
          "(daily rate x 30 days), not the document's own 1,707,337",
          wide_a == 3658579, wide_a)
    check("flight-mismatch: M35-64 avails_monthly scales the same way",
          wide_m == 2414640, wide_m)
    check("flight-mismatch: still 1 calendar month (both dates fall in September)",
          wide_n_months == 1, wide_n_months)
    check("flight-mismatch: full-flight basis (1 month) ties to monthly*1 -- "
          "internally consistent with avails_to_display's own convention",
          app.avails_to_display(wide_a, app.AVAILS_BASIS_FLIGHT, wide_n_months) == wide_a,
          wide_a)


def check_same_flight():
    """The reported live scenario: a fresh form, no flight set yet, imports
    Plaza Motors -- the document's own 9/17-9/30 flight lands on the form
    untouched, so the plan flight and the avail's own flight end up
    IDENTICAL. avails_monthly must be the document's stated figure exactly,
    both bases, and SOV must divide into the real number.

    Uploaded through the INTAKE entry point, not D2 -- FLOW_REWORK_PLAN.md
    Phase 1 gated D2 behind the setup band's own market+flight, so "a fresh
    form, no flight set yet" can no longer reach D2 at all; intake stayed
    deliberately ungated for exactly this reason. This also now exercises
    Phase 1's own parking mechanism (_finish_avails_import defers proration
    when the flight isn't known yet; apply_pending_avails_import_groups
    finishes it once the document's own dates land) -- AppTest's `.run()`
    settles the whole upload -> park -> field-apply -> gate-open -> un-park
    cascade in one call, so the assertions below see the same converged
    end state regardless of that intermediate detour."""
    print("Same-flight case: a fresh form (no flight set yet) imports the Plaza Motors avail "
          "through the intake entry point")
    at = new_app()
    at.session_state["avails_pdf_upload_path_intake"] = str(PLAZA_MOTORS)
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")

    groups = real_groups(at)
    check("2 groups created (A35-64 x Luxury, M35-64 x Luxury)", len(groups) == 2, len(groups))

    a3564 = next((g for g in groups if "DEMO Age A35-64" in g["terms"]), None)
    m3564 = next((g for g in groups if "DEMO Age M35-64" in g["terms"]), None)
    check("A35-64 x Luxury group found", a3564 is not None, [g["terms"] for g in groups])
    check("M35-64 x Luxury group found", m3564 is not None, [g["terms"] for g in groups])

    if a3564:
        check("A35-64 x Luxury avails_monthly ties to the document's own stated figure "
              "-- NOT prorated against the form's unrelated default flight",
              a3564["avails_monthly"] == A3564_LUXURY_IMPRESSIONS,
              a3564["avails_monthly"])
    if m3564:
        check("M35-64 x Luxury avails_monthly ties to the document's own stated figure",
              m3564["avails_monthly"] == M3564_LUXURY_IMPRESSIONS,
              m3564["avails_monthly"])

    check("the document's own flight landed on the form (single calendar month)",
          str(ss(at, "flight_start")) == "2026-09-17" and str(ss(at, "flight_end")) == "2026-09-30",
          (ss(at, "flight_start"), ss(at, "flight_end")))

    # Full-flight basis must show the SAME number, since the document's own
    # flight is exactly one calendar month -- never monthly * some other
    # month count.
    if a3564:
        full_flight = app.avails_to_display(a3564["avails_monthly"], app.AVAILS_BASIS_FLIGHT, 1)
        check("Max Avails (Full Flight, 1 month) shows the same stated figure",
              full_flight == A3564_LUXURY_IMPRESSIONS, full_flight)

    print("\nTicking the A35-64 group onto the plan prices its % of avails against the "
          "real 1,707,337 -- not the prorated 569,112 the first-cut fix's coincidence relied on")
    if a3564:
        at.session_state["_pending_group_include_all"] = True
        at.run()
        check("no exception after Add all", not at.exception,
              at.exception[0].message[:400] if at.exception else "")
        rows = ss(at, "plan_options")[0]["rows"]
        a3564_row = next((r for r in rows if a3564["id"] in app.group_ids_of(r)), None)
        check("a plan line exists for the A35-64 group", a3564_row is not None, rows)
        if a3564_row:
            groups_by_id = {g["id"]: g for g in real_groups(at)}
            matched = app.matched_avails_for_row(a3564_row, groups_by_id)
            check("the row's matched avails is the document's real stated figure",
                  matched == A3564_LUXURY_IMPRESSIONS, matched)


def check_flight_mismatch():
    """The case the first-cut fix could not have handled: the PLAN's actual
    flight (9/1-9/30, a rep-set 30-day September flight, simulating a real
    _avails_import_field CONFLICT -- the rep already had a flight set,
    different from the document's own, so the document's dates never
    landed) is WIDER than the avail document's own 14-day pull (9/17-9/30).

    Drives the real `apply_avails_import`/`_finish_avails_import` pair
    (the actual import pipeline, not a hand-rolled substitute) against a
    stubbed session_state pre-seeded with the already-different flight --
    the same technique tests/test_draft_regression.py uses for
    apply_draft_to_form, chosen here because reproducing this exact
    combination through the real widget precedence path is incidental to
    what this test is actually about (the day-count math), not the subject
    of it.
    """
    print("\nFlight-mismatch case: the plan's own flight (9/1-9/30, 30 days) is WIDER "
          "than the avail document's own 14-day pull (9/17-9/30)")
    document = avails_pdf_import.parse_avails_pdf(str(PLAZA_MOTORS))
    check("parsed 2 groups", len(document.groups) == 2, len(document.groups))
    check("document's own flight is the 14-day pull",
          document.flight_start == AVAIL_FLIGHT_START and document.flight_end == AVAIL_FLIGHT_END,
          (document.flight_start, document.flight_end))

    real_st = app.st
    stub = _StubSt({
        # The rep already has a real, different flight set -- 9/1-9/30 --
        # so _avails_import_field's own precedence rules (rep-set values
        # beat the document) would leave these exactly as they are; this
        # pre-seeds that OUTCOME directly rather than re-deriving it.
        "flight_start": date(2026, 9, 1), "flight_end": date(2026, 9, 30),
        "targeting_groups": [],
    })
    app.st = stub
    try:
        new_groups, report = app.apply_avails_import(document)
        check("2 new groups resolved", len(new_groups) == 2, len(new_groups))
        # Simulates the conflict outcome directly: no flight keys queued,
        # so _finish_avails_import falls back to the (already-different)
        # session flight -- exactly what a real conflict leaves in place.
        report["field_updates"].pop("flight_start", None)
        report["field_updates"].pop("flight_end", None)
        app._finish_avails_import(document, new_groups, report)
        groups = stub.session_state["targeting_groups"]
    finally:
        app.st = real_st

    check("the pre-seeded flight was left alone (9/1-9/30), not overwritten "
          "by the document's own 9/17-9/30",
          stub.session_state["flight_start"] == date(2026, 9, 1)
          and stub.session_state["flight_end"] == date(2026, 9, 30),
          (stub.session_state.get("flight_start"), stub.session_state.get("flight_end")))

    a3564 = next((g for g in groups if "DEMO Age A35-64" in g["terms"]), None)
    m3564 = next((g for g in groups if "DEMO Age M35-64" in g["terms"]), None)
    check("A35-64 x Luxury group found", a3564 is not None, [g.get("terms") for g in groups])
    check("M35-64 x Luxury group found", m3564 is not None, [g.get("terms") for g in groups])

    if a3564:
        rate = A3564_LUXURY_IMPRESSIONS / 14
        expected = round(rate * 30)   # 30 days in the plan's own September flight
        print(f"    ....  A35-64 x Luxury: daily rate {rate:,.4f} x 30 plan-flight days "
              f"= {rate * 30:,.2f} -> stored as {expected:,}")
        check("A35-64 x Luxury avails_monthly scales to the WIDER plan flight, "
              "not the document's own 1,707,337",
              a3564["avails_monthly"] == expected, a3564["avails_monthly"])
        check("...and that figure is NOT the document's own stated total "
              "(the two flights genuinely differ)",
              a3564["avails_monthly"] != A3564_LUXURY_IMPRESSIONS, a3564["avails_monthly"])
    if m3564:
        rate = M3564_LUXURY_IMPRESSIONS / 14
        expected = round(rate * 30)
        print(f"    ....  M35-64 x Luxury: daily rate {rate:,.4f} x 30 plan-flight days "
              f"= {rate * 30:,.2f} -> stored as {expected:,}")
        check("M35-64 x Luxury avails_monthly scales the same way",
              m3564["avails_monthly"] == expected, m3564["avails_monthly"])


def check_capital_media_multi_row():
    """Second real document, a different shape: FOUR monthly exception rows
    for one audience/geography (9/21-9/30, October, November, 12/1-12/20)
    instead of one row spanning the whole flight. Confirms three things the
    Plaza Motors (single-row) fixture can't reach on its own:

    1. The parser recognizes the four rows as ONE group, not four --
       `row_count` proves it summed them rather than, say, only reading the
       first.
    2. The summed total, and the daily rate derived from it, are exact --
       not an average of the four rows' own individual rates (which would
       coincidentally give the same number here, since this document's
       rate happens to be constant, but the formula never relies on that),
       not a double-count, not a recompute from a single row.
    3. The full-flight (SOV) figure ties to the document's own 3,321,409
       EXACTLY, via `avails_full_flight` -- not the few-impressions-off
       figure `avails_monthly * n_months` alone would produce once
       September and December are partial calendar months.
    """
    print("\nCapital Media (RFPID-266994): four monthly exception rows for one "
          "audience/geography, not one row spanning the whole flight")
    document = avails_pdf_import.parse_avails_pdf(str(CAPITAL_MEDIA))
    check("parsed as exactly 1 group (the four rows collapsed into one, not four)",
          len(document.groups) == 1, len(document.groups))
    if not document.groups:
        return
    group = document.groups[0]
    check("the group's row_count is 4 -- proof the four rows were actually summed, "
          "not just the first one read",
          group.row_count == 4, group.row_count)
    check("the four rows' own stated impressions sum to the group's total",
          sum(imp for _s, _e, imp in CAPITAL_MEDIA_ROWS) == CAPITAL_MEDIA_TOTAL,
          [imp for _s, _e, imp in CAPITAL_MEDIA_ROWS])
    check("the parsed group's impressions equal the document's own stated total, exactly",
          group.impressions == CAPITAL_MEDIA_TOTAL, group.impressions)
    check("the document's own flight is the full 9/21-12/20 span",
          document.flight_start == CAPITAL_MEDIA_FLIGHT_START
          and document.flight_end == CAPITAL_MEDIA_FLIGHT_END,
          (document.flight_start, document.flight_end))

    daily_rate = app.avails_daily_rate(group.impressions, document.flight_start, document.flight_end)
    print(f"    ....  daily rate: {group.impressions:,} / 91 days = {daily_rate:,.4f} / day "
          f"(document's four rows each imply {CAPITAL_MEDIA_DAILY_RATE:,.0f}/day)")
    check("the derived daily rate matches every one of the four rows' own implied rate "
          "(not an average masking real per-row variation -- this document has none)",
          abs(daily_rate - CAPITAL_MEDIA_DAILY_RATE) < 0.01, daily_rate)
    for start, end, imp in CAPITAL_MEDIA_ROWS:
        row_days = (end - start).days + 1
        check(f"row {start}-{end} ({row_days}d, {imp:,} imp) implies the same daily rate",
              abs(imp / row_days - CAPITAL_MEDIA_DAILY_RATE) < 0.01, imp / row_days)

    print("\nSame-flight import: a fresh form imports Capital Media directly, through intake "
          "(D2 is gated behind the setup band's flight -- see check_same_flight's docstring)")
    at = new_app()
    at.session_state["avails_pdf_upload_path_intake"] = str(CAPITAL_MEDIA)
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    groups = real_groups(at)
    check("1 group created", len(groups) == 1, len(groups))
    if not groups:
        return
    g = groups[0]
    n_months = max(1, len(app.month_list(CAPITAL_MEDIA_FLIGHT_START, CAPITAL_MEDIA_FLIGHT_END)))
    check("document's own flight touches 4 calendar months (Sep, Oct, Nov, Dec)",
          n_months == 4, n_months)
    check("avails_full_flight is the document's own stated total, EXACTLY -- no "
          "distortion from September/December being partial calendar months",
          g.get("avails_full_flight") == CAPITAL_MEDIA_TOTAL, g.get("avails_full_flight"))
    approx_full_flight = g["avails_monthly"] * n_months
    print(f"    ....  avails_monthly {g['avails_monthly']:,} x {n_months} months = "
          f"{approx_full_flight:,} (the OLD multiply-out -- off by "
          f"{abs(approx_full_flight - CAPITAL_MEDIA_TOTAL)}, which is exactly why "
          f"avails_full_flight is stored and preferred instead)")

    print("\nSOV: ticking the group onto the plan, full-flight avails on the row "
          "ties to 3,321,409 exactly, via matched_avails_full_flight_for_row")
    at.session_state["_pending_group_include_all"] = True
    at.run()
    check("no exception after Add all", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    rows = ss(at, "plan_options")[0]["rows"]
    row = next((r for r in rows if g["id"] in app.group_ids_of(r)), None)
    check("a plan line exists for the group", row is not None, rows)
    if row:
        groups_by_id = {gr["id"]: gr for gr in real_groups(at)}
        exact_full_flight = app.matched_avails_full_flight_for_row(row, groups_by_id, n_months)
        check("matched_avails_full_flight_for_row ties to the document's stated "
              "total exactly -- the SOV denominator a client-facing slide would show",
              exact_full_flight == CAPITAL_MEDIA_TOTAL, exact_full_flight)
        totals = app.compute_plan_totals(
            [row], app.BREAKOUT_MONTHLY, n_months, "Sep 2026 - Dec 2026",
            groups_by_id=groups_by_id)
        preview_row = totals["preview_rows"][0]
        check("compute_plan_totals' own preview_rows carries the same exact figure "
              "(the actual code path the deck and app preview both read)",
              preview_row["matched_avails_full_flight"] == CAPITAL_MEDIA_TOTAL,
              preview_row["matched_avails_full_flight"])


def check_fold_back_survival():
    """`avails_full_flight` must survive an untouched D2 rerun byte-for-
    byte, and must be DROPPED (falling back to the approximate multiply-
    out) the moment a rep actually edits the Max Monthly Avails cell -- the
    same "an unchanged cell keeps its value, a real edit re-derives"
    discipline the D2 grid's Audience/Markets/Label/Color/Plan columns
    already follow (CLAUDE.md's `_cell_unchanged` convention).

    Found and fixed in the same round as this file's other checks: the D2
    grid's own fold-back rebuilds every group fresh via `tg.new_group` on
    EVERY render, which has no notion of this field at all -- so a real
    import's exact total read back as None the instant the D2 grid itself
    rendered even once, before any rep edit at all. Caught here so it can't
    come back. Uses the same `st.data_editor` monkeypatch technique
    tests/test_avails_grid_manual_row.py established (AppTest can't drive
    a data_editor directly).
    """
    import streamlit as st_module

    print("\nFold-back survival: avails_full_flight must not be silently "
          "dropped by the D2 grid's own re-render")

    real_data_editor = st_module.data_editor
    state = {"mode": None}

    def fake_data_editor(data, *args, **kwargs):
        key = str(kwargs.get("key", ""))
        if not key.startswith("avails_editor") or state["mode"] != "edit":
            return data
        avails_col = [c for c in data.columns if c not in
                     ("gid", "Plan", "Audience", "Markets", "Label", "Color", "Detached")][0]
        out = data.copy()
        out.loc[0, avails_col] = 999999
        state["mode"] = "done"
        return out

    st_module.data_editor = fake_data_editor
    try:
        at = new_app()
        # Intake, not D2 -- see check_same_flight's docstring. The gate
        # opens as soon as the document's own dates land (this run), so the
        # D2 grid this test actually exercises still renders normally on
        # every SUBSEQUENT run below.
        at.session_state["avails_pdf_upload_path_intake"] = str(CAPITAL_MEDIA)
        at.run()
        groups = real_groups(at)
        check("1 group created", len(groups) == 1, len(groups))
        if not groups:
            return
        check("avails_full_flight set immediately after import",
              groups[0].get("avails_full_flight") == CAPITAL_MEDIA_TOTAL,
              groups[0].get("avails_full_flight"))

        print("    ....  untouched rerun")
        at.run()
        groups2 = real_groups(at)
        check("avails_full_flight survives an untouched rerun byte-for-byte",
              groups2[0].get("avails_full_flight") == CAPITAL_MEDIA_TOTAL,
              groups2[0].get("avails_full_flight"))

        print("    ....  a genuine edit to the Max Monthly Avails cell")
        state["mode"] = "edit"
        at.run()
        check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
        groups3 = real_groups(at)
        check("avails_full_flight is dropped once the rep hand-edits the avails figure "
              "-- the exact total no longer describes what's actually stored",
              groups3[0].get("avails_full_flight") is None, groups3[0].get("avails_full_flight"))
        check("avails_monthly reflects the rep's own typed figure",
              groups3[0].get("avails_monthly") == 999999, groups3[0].get("avails_monthly"))
    finally:
        st_module.data_editor = real_data_editor


def main():
    global skipped
    if not PLAZA_MOTORS.exists() or not CAPITAL_MEDIA.exists():
        print("SKIP -- real avails PDFs not present (gitignored fixtures)")
        skipped = True
        return 0

    check_daily_rate_pure()
    check_same_flight()
    check_flight_mismatch()
    check_capital_media_multi_row()
    check_fold_back_survival()

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("An avails-PDF group's impressions reduce to a daily rate scoped to the "
          "document's own flight, then reapply to whatever flight the plan line "
          "actually runs -- correct whether the two flights are identical (the "
          "reported case) or genuinely differ (the case the first fix couldn't "
          "have handled).")
    return 0


if __name__ == "__main__":
    code = main()
    if skipped:
        sys.exit(0)
    sys.exit(code)
