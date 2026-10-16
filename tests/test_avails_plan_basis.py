"""Avails are the ceiling, the request is the plan (2026-10-15, Adventist
Health follow-up).

An avails PDF is a pull of MAXIMUM available inventory: it belongs on the
avails slide ("Max Monthly Avails"), never as the media plan's planned
impressions. The plan comes from the request -- budget / CPM, or what the
notes state -- and uses the max only when the notes ask for it ("max out
the avails"). A plan line above its audience x geography's max avails is
warned about in the app ("Plan exceeds available inventory").

Also here: client-facing DMA names ("Washington, DC", never Nielsen's
"Washington-Hagerstown"), checked across all 210 markets.

The expectations come from things the code can't move: the stated $20,000
budget and the rate card's CPM for the budget case, the avails figures the
PDFs themselves state for the max-out case, each market's own county states
for the DMA names.

The Adventist PDFs carry real client pricing and are gitignored; the parts
that need them report SKIP without them.

    python tests/test_avails_plan_basis.py
"""
import copy
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import market_lookup  # noqa: E402
import targeting_groups as tg  # noqa: E402
import test_avails_multi_file as amf  # noqa: E402

market_lookup.install()
failures = []
BUDGET = 20_000


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def dma_names():
    print("\nclient-facing DMA names")
    check("Washington-Hagerstown reads Washington, DC",
          market_lookup.client_market_name("washington_hagerstown") == "Washington, DC")
    check("Harrisburg-Lancaster-Lebanon-York reads Harrisburg, PA",
          market_lookup.client_market_name("harrisburg_lancaster_lebanon_york") == "Harrisburg, PA")
    check("a name that already carries its state is kept",
          market_lookup.client_market_name("charleston_sc") == "Charleston, SC")
    check("the app's display name is the client name",
          app._market_display_name("washington_hagerstown") == "Washington, DC")
    markets = market_lookup.load()["markets"]
    wrong = []
    for key, entry in markets.items():
        name = market_lookup.client_market_name(key)
        state = name.rsplit(", ", 1)[-1] if ", " in name else None
        covered = market_lookup.states_for_market(key)
        if not state or "-" in name.replace("Tri-Cities", "") or (covered and state not in covered):
            wrong.append((key, name, sorted(covered)))
    check(f"all {len(markets)} DMAs: principal city, a state the market covers, no secondary city",
          not wrong, wrong[:5])
    real_load = market_lookup.load
    market_lookup.load = lambda: {"markets": {"x": {"name": "Washington, DC (Hagerstown)",
                                                     "dma_code": 511}}}
    try:
        check("a parenthetical secondary city is dropped",
              market_lookup.client_market_name("x") == "Washington, DC",
              market_lookup.client_market_name("x"))
    finally:
        market_lookup.load = real_load


def apply_draft(draft, preset):
    real = app.st
    stub = amf._StubSt(preset)
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return stub.session_state


def make_draft(groups, total_budget, allocation, breakout="full_flight"):
    start, end = groups[0]["avails_doc_start"], groups[0]["avails_doc_end"]
    return {
        "client_name": "Adventist Health", "vertical": "healthcare", "market": "DC",
        "flight_start": start, "flight_end": end,
        "total_budget": total_budget, "breakout": breakout,
        "media_plan_lines": [], "group_selection": {"mode": "all"},
        "group_allocation": allocation, "group_selection_reason": "",
        "audiences": [], "attribution": [], "sports": [],
        "unresolved": [], "unresolved_internal": [],
    }


def group_rows(state):
    return [r for opt in state["plan_options"] for r in opt["rows"] if app.group_ids_of(r)]


def draft_checks():
    groups, _docs = amf.import_both()
    from datetime import date
    preset = {"targeting_groups": groups,
              "flight_start": date.fromisoformat(groups[0]["avails_doc_start"]),
              "flight_end": date.fromisoformat(groups[0]["avails_doc_end"])}
    _label, cpm = app.line_product_spec(app.GROUP_LINE_PRODUCT_KEY)

    print(f"\navails attached, a ${BUDGET:,} budget: impressions are budget / CPM (${cpm})")
    state = apply_draft(make_draft(groups, BUDGET, {"split_evenly": True}), preset)
    rows = group_rows(state)
    total_cost = sum(app._num(r["Cost"]) for r in rows)
    total_imps = sum(app._num(r["Impressions"]) for r in rows)
    check("the plan spends the stated budget", abs(total_cost - BUDGET) < 1, total_cost)
    check("planned impressions = budget / CPM x 1000",
          abs(total_imps - BUDGET / cpm * 1000) <= len(rows), (total_imps, BUDGET / cpm * 1000))
    avails_total = sum(amf.full_flight(g) for g in groups)
    check("not the avails (10M+ available, a few hundred thousand planned)",
          total_imps < avails_total / 10, (total_imps, avails_total))
    check("no line's impressions equal its own avails",
          not any(app._num(r["Impressions"]) and app._num(r["Impressions"]) == amf.full_flight(g)
                  for r in rows for g in groups if g["id"] in app.group_ids_of(r)))

    print(f"\nno split stated: ${BUDGET:,} in proportion to each line's max avails")
    import avails_grid as ag
    dup_ids = set(ag.duplicate_groups(groups))
    check("Adventist has exactly one duplicate line", len(dup_ids) == 1, dup_ids)
    state = apply_draft(make_draft(groups, BUDGET, {"proportional_to_avails": True}), preset)
    rows = group_rows(state)
    by_id = {g["id"]: g for g in groups}
    check("8 lines, not 9 -- the duplicate is left off", len(rows) == 8, len(rows))
    check("the duplicate is not ticked onto the plan",
          not any(g["include_in_plan"] for g in state["targeting_groups"] if g["id"] in dup_ids))
    costs = [app._num(r["Cost"]) for r in rows]
    check(f"costs add to exactly ${BUDGET:,}", sum(costs) == BUDGET, sum(costs))
    check("every cost is whole dollars", all(c == int(c) for c in costs), costs)
    check("every line plans whole impressions",
          all(app._num(r["Impressions"]) == int(app._num(r["Impressions"])) for r in rows))
    weights = [app._num(by_id[app.group_ids_of(r)[0]]["avails_monthly"]) for r in rows]
    raw = [BUDGET * w / sum(weights) for w in weights]
    check("each line's share follows its avails (within the $1 rounding)",
          all(abs(c - x) < 1 for c, x in zip(costs, raw)),
          list(zip(costs, [round(x, 2) for x in raw])))
    largest = max(range(len(raw)), key=lambda i: raw[i])
    check("the rounding remainder sits on the largest line",
          all(c == round(x) for i, (c, x) in enumerate(zip(costs, raw)) if i != largest),
          list(zip(costs, [round(x, 2) for x in raw])))
    check("each line's impressions = its cost / CPM, to the impression",
          all(abs(app._num(r["Impressions"]) - app._num(r["Cost"]) / cpm * 1000) <= 1 for r in rows))
    check(f"impressions add to exactly what ${BUDGET:,} buys ({BUDGET / cpm * 1000:,.0f})",
          sum(app._num(r["Impressions"]) for r in rows) == round(BUDGET / cpm * 1000),
          sum(app._num(r["Impressions"]) for r in rows))
    for r, c in zip(rows, costs):
        print(f"        {r['Targeting'][:34]:34} {r['Geo'][:28]:28} ${c:>8,.0f}  "
              f"{int(app._num(r['Impressions'])):>9,}")

    print("\nthe notes state a split: it wins")
    state = apply_draft(make_draft(groups, BUDGET, {"split_evenly": True}), preset)
    costs = [app._num(r["Cost"]) for r in group_rows(state)]
    check("split evenly when the notes say so", max(costs) - min(costs) <= 1, costs)

    print("\nnotes say \"max out the avails\": the plan uses the avails")
    state = apply_draft(make_draft(groups, 0, {"percent_of_avails": 100}, breakout="monthly"), preset)
    rows = [r for r in group_rows(state) if app._num(r["Impressions"])]
    by_id = {g["id"]: g for g in groups}
    mismatched = []
    for r in rows:
        members = [by_id[gid] for gid in app.group_ids_of(r)]
        expected = tg.entity_avails_monthly(members)
        if abs(app._num(r["Impressions"]) - expected) > 1:
            mismatched.append((r["Targeting"], r["Impressions"], expected))
    check("priced lines carry their audience x geography's full monthly avails",
          rows and not mismatched, mismatched or "no priced lines")

    print("\nthe prompt says when the max applies")
    prompt = app.build_draft_prompt("notes", existing_groups=groups) \
        if "existing_groups" in app.build_draft_prompt.__code__.co_varnames else ""
    check("avails described as the maximum, used only on an explicit ask",
          "MAXIMUM inventory available" in prompt and "max out the avails" in prompt)


def name_checks():
    print("\nslide audience names: humanized, rep-editable")
    import avails_grid as ag
    check("taxonomy prefixes dropped",
          [ag.slide_audience_name(a) for a in ("LIFESTYLE Doctors and Nurses",
                                                "HLTH CRX DX Heart Conditions", "HLTH CRX Neurology")]
          == ["Doctors and Nurses", "Heart Conditions", "Neurology"])
    check("the rep's own wording wins",
          ag.slide_audience_name("HLTH CRX Neurology", {"HLTH CRX Neurology": "Brain health"})
          == "Brain health")


def overrun_checks():
    print("\nplan above avails -> warning")
    row = {"tactic": "Premion Streaming TV", "geo": "Washington, DC", "targeting": "HLTH Doctors",
           "is_flat_fee": False, "monthly_impressions": 900_000, "full_flight_impressions": 900_000,
           "matched_avails_monthly": 600_000, "matched_avails_full_flight": 600_000}
    under = dict(row, monthly_impressions=500_000, full_flight_impressions=500_000)
    unmatched = dict(row, matched_avails_monthly=None, matched_avails_full_flight=None)
    check("a line above its avails is named", len(app.plan_inventory_overruns([row])) == 1)
    check("a line within its avails, or with no matched avails, is not",
          not app.plan_inventory_overruns([under, unmatched]))

    at = amf.new_app()
    at.session_state["avails_pdf_upload_path_intake"] = str(amf.DOC_A)
    at.run()
    add_all = [b for b in at.button if str(b.key).startswith("avails_include_all_")]
    if add_all:
        add_all[0].click().run()
    options = at.session_state["plan_options"]
    for opt in options:
        for r in opt["rows"]:
            if app.group_ids_of(r):
                r["Impressions"], r["Cost"] = 50_000_000.0, 50_000_000.0 * r["CPM"] / 1000
                break
        opt["version"] += 1
    at.session_state["plan_options"] = options
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:600] if at.exception else "")
    warnings = [w.value for w in at.warning if "Plan exceeds available inventory" in w.value]
    check("the app warns: Plan exceeds available inventory", len(warnings) == 1, warnings)


def main():
    dma_names()
    name_checks()
    if amf.DOC_A and amf.DOC_B and amf.DOC_A.exists() and amf.DOC_B.exists():
        draft_checks()
        overrun_checks()
    else:
        print("SKIP -- the Adventist Health avails PDFs aren't present (gitignored fixtures)")
    print(f"\n{'FAILED: ' + str(len(failures)) if failures else 'ALL PASSED'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
