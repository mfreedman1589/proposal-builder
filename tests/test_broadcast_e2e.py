"""End to end: a real Wide Orbit file through the real form to a real deck,
asserting the media plan and the schedule slides agree on the numbers."""
import os
import sys
from datetime import date

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
os.chdir(REPO)

from streamlit.testing.v1 import AppTest       # noqa: E402
import app, assembly, db, slide_map, wideorbit  # noqa: E402

FIXTURE = os.path.join("tests", "fixtures", "wideorbit", "regency_planner.xls")
if not os.path.exists(FIXTURE):
    print(f"SKIP -- {FIXTURE} not present (real client schedules are gitignored).")
    sys.exit(0)
schedule = wideorbit.parse_schedule(FIXTURE, "regency_planner.xls")

captured = {}
real_personalize, real_log = assembly.personalize, db.log_proposal


def spy(prs, fill_data):
    captured["prs"] = prs
    captured["fill"] = fill_data
    return real_personalize(prs, fill_data)


assembly.personalize = spy
db.log_proposal = lambda *a, **k: ("test", None)

failures = []


def check(label, ok, detail=""):
    print(f"  {'PASS' if ok else 'FAIL'}  {label}{('  ' + str(detail)) if detail else ''}")
    if not ok:
        failures.append(label)


try:
    at = AppTest.from_file("app.py", default_timeout=600)
    for key, value in {
        "authed": True, "current_user": "Matt", "total_tv": True, "market_choice": "DC",
        # FLOW_REWORK_PLAN.md Phase 1: the setup band gates everything below
        # it on market+flight -- broadcast rows are exempt from the band's
        # flight once the form renders (their own schedule dates are used
        # verbatim), but the form has to render at all first.
        "flight_start": date(2026, 9, 1), "flight_end": date(2026, 11, 30),
        "broadcast_schedule": schedule, "agency_involved": True,
        "broadcast_plan_desc": "210x Commercials Per Month, Morning News Mon-Tue",
    }.items():
        at.session_state[key] = value
    at.run()
    assert not at.exception, at.exception

    rows = at.session_state["plan_options"][0]["rows"]
    broadcast = [r for r in rows if app.is_broadcast_row(r)]
    months = len(at.session_state["active_months"])
    s = schedule.summary

    print("\n--- the seeded broadcast line ---")
    check("exactly one broadcast line", len(broadcast) == 1, len(broadcast))
    check("the rate-card Broadcast Schedule line was replaced, not duplicated",
          not any(str(r.get("Tactic", "")).strip() == "Broadcast Schedule" for r in rows))
    row = broadcast[0]
    print(f"       tactic : {row['Tactic']}")
    print(f"       geo    : {row['Geo']}")
    print(f"       target : {row['Targeting']}")
    print(f"       monthly: {row['Impressions']:,.0f} imps  ${row['Cost']:,.0f}  "
          f"CPM ${row['CPM']:.2f}")
    check("tactic names the station and format",
          row["Tactic"] == f"WUSA {app.BROADCAST_TACTIC_SUFFIX}")
    check("geo mapped from the station", row["Geo"] == "Washington DC DMA")
    check("targeting uses the rep's description",
          "210x Commercials Per Month" in row["Targeting"])

    print("\n--- the gross rule (agency toggle is ON) ---")
    check("markup for this row is 1.0, not 1.15", app.row_markup(row, 1.15) == 1.0)
    check("other rows still get 1.15",
          all(app.row_markup(r, 1.15) == 1.15 for r in rows if not app.is_broadcast_row(r)))
    check("cost x markup is untouched: cpm*imps/1000 == cost",
          abs(row["CPM"] * row["Impressions"] / 1000 - row["Cost"]) < 1.0)

    # The multiplier is the months the SCHEDULE runs in, not the plan's. A
    # four-week May buy is one month of broadcast however long the streaming
    # campaign around it is; scaling by the plan's month count is what turned
    # $28,000 into $84,000 and gave a $9,333 "monthly" figure to a schedule
    # that has nothing to do with those months.
    sched_months = max(1, schedule.active_month_count())
    print(f"\n--- agreement with Wide Orbit (x{sched_months} schedule month(s); "
          f"the plan's own flight is {months}) ---")
    check(f"impressions {row['Impressions'] * sched_months:,.0f} == WO {s.impressions:,.0f}",
          abs(row["Impressions"] * sched_months - s.impressions) <= sched_months)
    check(f"cost ${row['Cost'] * sched_months:,.0f} == WO ${s.gross_cost:,.0f}",
          abs(row["Cost"] * sched_months - s.gross_cost) <= sched_months)
    check(f"CPM ${row['CPM']:.2f} == WO ${s.cpm:.2f}", abs(row["CPM"] - s.cpm) < 0.02)

    generate = [b for b in at.button if b.label == "Generate proposal"][0]
    generate.click().run()
    assert not at.exception, at.exception

    prs = captured["prs"]
    slides = [x for x in prs.slides
              if "BROADCAST TV | MEDIA PLAN" in slide_map.extract_slide_text(x).upper()]
    print("\n--- the generated deck ---")
    check("schedule slides were produced", len(slides) >= 1, len(slides))
    final = slide_map.extract_slide_text(slides[-1])
    check("schedule slide carries the flight's spot total", f"{s.total_spots:,}" in final)
    check("schedule slide carries the gross cost", f"${s.gross_cost:,.0f}" in final)
    check("schedule slide carries reach and frequency",
          f"{s.reach:.1f}" in final and f"{s.frequency:.1f}" in final)
    # Humanized for the slide ("CS-A25+" -> "Adults 25+"), but still derived
    # from the parse rather than hardcoded -- so the check is that what's on
    # the slide is what this schedule's own demo expands to.
    expected_demo = wideorbit.humanize_demo(s.demo_label)
    check(f"schedule slide names this schedule's demo ({s.demo_label!r} -> {expected_demo!r})",
          expected_demo in final)

    payload = captured["fill"]["media_plan_options"][0]
    line = [r for r in payload["rows"] if app.BROADCAST_TACTIC_MARKER in r["tactic"]]
    check("the deck's media plan carries the broadcast line", len(line) == 1)
    if line:
        print(f"       {line[0]}")
        deck_imps = int(line[0]["impressions"].replace(",", ""))
        deck_cost = float(line[0]["cost"].replace("$", "").replace(",", "").split()[0])
        check("deck's media plan impressions == the grid's", deck_imps == int(row["Impressions"]))
        check("deck's media plan cost == the grid's", abs(deck_cost - row["Cost"]) < 1)
        check("deck plan x schedule months == the schedule slide's total",
              abs(deck_imps * sched_months - s.impressions) <= sched_months)
    # --- the two breakouts are separate questions -------------------------
    # The schedule slides' breakout lays out the grid; the media plan's own
    # breakout decides what basis the plan is quoted in. Crossing them is the
    # kind of bug that shows a client a monthly plan against a full-flight
    # schedule and expects them not to notice.
    print("\n--- schedule breakout must not move media plan numbers ---")

    def plan_numbers(schedule_breakout, plan_breakout):
        run = AppTest.from_file("app.py", default_timeout=600)
        for key, value in {
            "authed": True, "current_user": "Matt", "total_tv": True, "market_choice": "DC",
        # FLOW_REWORK_PLAN.md Phase 1: the setup band gates everything below
        # it on market+flight -- broadcast rows are exempt from the band's
        # flight once the form renders (their own schedule dates are used
        # verbatim), but the form has to render at all first.
        "flight_start": date(2026, 9, 1), "flight_end": date(2026, 11, 30),
            "broadcast_schedule": schedule, "agency_involved": True,
            "broadcast_breakout": schedule_breakout,
        }.items():
            run.session_state[key] = value
        run.run()
        run.session_state["plan_options"][0]["breakout"] = plan_breakout
        run.run()
        line = [r for r in run.session_state["plan_options"][0]["rows"]
                if app.is_broadcast_row(r)][0]
        return round(line["Impressions"]), round(line["Cost"])

    months = len(at.session_state["active_months"])
    # Divided by the schedule's own months, and full flight is the WO totals
    # verbatim -- never monthly x the plan's month count.
    monthly_expected = (round(s.impressions / sched_months), round(s.gross_cost / sched_months))
    flight_expected = (round(s.impressions), round(s.gross_cost))
    for schedule_breakout in (app.BREAKOUT_FULL_FLIGHT_LABEL, app.BREAKOUT_MONTHLY_LABEL):
        got = plan_numbers(schedule_breakout, app.BREAKOUT_MONTHLY)
        check(f"monthly plan is WO/{sched_months} whatever the schedule shows ({schedule_breakout})",
              got == monthly_expected, got)
        got = plan_numbers(schedule_breakout, app.BREAKOUT_FULL_FLIGHT)
        check(f"full-flight plan is WO totals whatever the schedule shows ({schedule_breakout})",
              got == flight_expected, got)
    # --- the broadcast line's Targeting is schedule-derived, and STAYS so ---
    #
    # Broadcast is bought by program and daypart, not by audience segment.
    # resolve_row_defaults used to re-seed every clean row's Targeting from
    # the Campaign Specs Audience stack on any shared-field edit, which
    # overwrote the schedule-derived summary -- a real demo deck went out
    # with the broadcast line reading "Homeowners with higher household
    # income (100K+)...". The edit is what triggers it, so the assertion has
    # to make one rather than just reading the freshly seeded row.
    print()
    print("--- broadcast Targeting is schedule-derived, not the audience stack ---")
    audience = "\n".join([
        "Homeowners with higher household income (100K+)",
        "Adults 35-64",
        "In-market for home services",
    ])
    run = AppTest.from_file("app.py", default_timeout=600)
    for key, value in {
        "authed": True, "current_user": "Matt", "total_tv": True, "market_choice": "DC",
        # FLOW_REWORK_PLAN.md Phase 1: the setup band gates everything below
        # it on market+flight -- broadcast rows are exempt from the band's
        # flight once the form renders (their own schedule dates are used
        # verbatim), but the form has to render at all first.
        "flight_start": date(2026, 9, 1), "flight_end": date(2026, 11, 30),
        "broadcast_schedule": schedule, "agency_involved": True,
        "broadcast_plan_desc": "210x Commercials Per Month, Morning News Mon-Tue",
    }.items():
        run.session_state[key] = value
    run.run()
    # A COPY. session_state hands back the live row dict, and the re-seed
    # mutates it in place -- holding the reference made "unchanged by the
    # edit" compare the row against itself and pass no matter what happened.
    before = dict([r for r in run.session_state["plan_options"][0]["rows"]
                   if app.is_broadcast_row(r)][0])
    stack = app.audience_stack(audience)
    # Now edit the shared Audience field, which is what re-seeds clean rows.
    run.session_state["audience_text"] = audience
    run.run()
    after = [r for r in run.session_state["plan_options"][0]["rows"]
             if app.is_broadcast_row(r)][0]

    check("the audience stack really would have been a different string",
          stack and stack != before["Targeting"], stack)
    check("a streaming line DID take the new audience stack (the edit landed)",
          any(r["Targeting"] == stack for r in run.session_state["plan_options"][0]["rows"]
              if not app.is_broadcast_row(r) and r["Tactic"].startswith("Premion")),
          [r["Targeting"] for r in run.session_state["plan_options"][0]["rows"]])
    check("broadcast Targeting is unchanged by the audience edit",
          after["Targeting"] == before["Targeting"], after["Targeting"])
    check("broadcast Targeting is not the audience stack",
          after["Targeting"] != stack, after["Targeting"])
    check("broadcast Targeting is the schedule-derived copy",
          after["Targeting"] == "210x Commercials Per Month, Morning News Mon-Tue",
          after["Targeting"])
    # Same rule, and the more dangerous half: the Geo comes from the
    # station's call sign and must not be replaced by the form's market.
    check("broadcast Geo is still station-derived",
          after["Geo"] == before["Geo"] == "Washington DC DMA", after["Geo"])

finally:
    assembly.personalize, db.log_proposal = real_personalize, real_log

print("\n" + ("ALL AGREE" if not failures else f"{len(failures)} FAILED: {failures}"))
sys.exit(1 if failures else 0)
