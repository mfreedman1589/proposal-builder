"""The second half of tests/test_group_geo_resolution.py (split 2026-11-03 so
each file finishes well inside a 10-minute foreground call): Scenario 0's
exact inputs, the Section-A market autofill and a rep's removal surviving it,
the same ledger from drafting's side, the broadcast row's station-derived
Geo, and install_market_lookup reporting WHY it failed.

    python tests/test_group_geo_autofill.py

Driven through the real form with AppTest. The helpers, and the stub that
keeps the one Census geocoder call off the network, come from
test_group_geo_resolution.py, imported rather than copied.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_group_geo_resolution import *  # noqa: E402,F401,F403 -- helpers + network stub
from test_group_geo_resolution import failures  # noqa: E402


def main():
    print("Scenario 0's own exact inputs (targeting_groups_test_scenarios.md), not just an "
          "equivalent case -- the multi-county semicolon list and the 21401 radius weren't "
          "covered above with THESE inputs, and the Zips unresolved-99999 case above used "
          "different numbers than the scenario doc names")
    at3b = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid3b = real_group_id(at3b)
    set_mode(at3b, gid3b, "Zips")
    [w for w in at3b.text_area if str(w.key) == f"geo_zips_{gid3b}"][0].set_value("20120, 20147, 20148, 99999")
    at3b.run()
    resolve(at3b, gid3b)
    check("no exception", not at3b.exception, at3b.exception[0].message[:300] if at3b.exception else "")
    g3b = group_by_id(at3b, gid3b)
    check("the three real zips resolve", set(g3b["resolved_zips"]) >= {"20120", "20147", "20148"},
          g3b["resolved_zips"])
    check("99999 is reported as unresolved, not silently dropped",
          "99999" in g3b["resolved_zips"], g3b["resolved_zips"])
    # Same calm-note-not-warning behavior as the check above -- 99999 here
    # is still just a well-formed, unmapped zip.
    captions3b = [str(c.value) if hasattr(c, "value") else str(c) for c in at3b.caption]
    check("99999 is explained in a calm note", any("1 zip(s) have no county on file" in c
          for c in captions3b), captions3b)
    warnings3b_final = [str(w.value) if hasattr(w, "value") else str(w) for w in at3b.warning]
    check("...and no yellow warning names 99999",
          not any("99999" in w for w in warnings3b_final), warnings3b_final)

    at3c = new_app([{"Audience": "Homeowners", "Geo": "Somerset NJ", COL: 100000}])
    gid3c = real_group_id(at3c)
    set_mode(at3c, gid3c, "Counties")
    [w for w in at3c.text_area if str(w.key) == f"geo_counties_{gid3c}"][0].set_value(
        "NEW CASTLE DE; CHESTER PA")
    at3c.run()
    resolve(at3c, gid3c)
    check("no exception", not at3c.exception, at3c.exception[0].message[:300] if at3c.exception else "")
    g3c = group_by_id(at3c, gid3c)
    check("the semicolon-separated NAME STATE list (no suffixes) resolves both counties",
          g3c["geo_def"] == {"kind": "counties", "counties": ["NEW CASTLE DE", "CHESTER PA"]},
          g3c["geo_def"])
    check("resolves to real zips, then to markets",
          len(g3c["resolved_zips"]) > 0 and len(g3c["resolved_markets"]) > 0,
          (len(g3c["resolved_zips"]), g3c["resolved_markets"]))
    check("Philadelphia is among the resolved markets (both counties sit in its orbit)",
          "philadelphia" in g3c["resolved_markets"], g3c["resolved_markets"])

    at3d = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid3d = real_group_id(at3d)
    set_mode(at3d, gid3d, "Radius")
    [w for w in at3d.text_area if str(w.key) == f"geo_radius_centers_{gid3d}"][0].set_value("21401")
    at3d.run()
    [w for w in at3d.text_input if str(w.key) == f"geo_radius_miles_{gid3d}"][0].set_value("10")
    at3d.run()
    resolve(at3d, gid3d)
    check("no exception", not at3d.exception, at3d.exception[0].message[:300] if at3d.exception else "")
    g3d = group_by_id(at3d, gid3d)
    check("10mi around 21401 (Annapolis) resolves to a real zip list",
          len(g3d["resolved_zips"]) > 0, len(g3d["resolved_zips"]))
    check("resolves to Baltimore -- the market a person would name for Annapolis",
          "baltimore" in g3d["resolved_markets"], g3d["resolved_markets"])

    print("\nthe autofill adds a resolved market to Section A, and a rep's removal "
          "survives a further, different group resolving to it again")
    at5 = new_app([{"Audience": "Homeowners", "Geo": "Somerset NJ", COL: 100000}])

    def target_dmas(a):
        return list([w for w in a.multiselect if str(w.key) == "target_dma_choice"][0].value)

    check("no target DMAs before any resolution", target_dmas(at5) == [], target_dmas(at5))
    gid5 = real_group_id(at5)
    set_mode(at5, gid5, "Counties")
    [w for w in at5.text_area if str(w.key) == f"geo_counties_{gid5}"][0].set_value("Somerset NJ")
    at5.run()
    resolve(at5, gid5)
    after_resolve = target_dmas(at5)
    check("New York and Philadelphia were autofilled into Section A",
          {"New York City", "Philadelphia"} <= set(after_resolve), after_resolve)

    picker5 = [w for w in at5.multiselect if str(w.key) == "target_dma_choice"][0]
    picker5.set_value([v for v in after_resolve if "New York" not in v])
    at5.run()
    check("the rep's removal of New York took",
          "New York City" not in target_dmas(at5), target_dmas(at5))

    search = [w for w in at5.text_input if str(w.key) == "finder_search"][0]
    search.set_value("AFIRST Foodies")
    at5.run()
    [w for w in at5.button if str(w.key) == "finder_new_AFIRST Foodies"][0].click()
    at5.run()
    gid5b = real_group_id(at5, exclude={gid5})
    set_mode(at5, gid5b, "Zips")
    [w for w in at5.text_area if str(w.key) == f"geo_zips_{gid5b}"][0].set_value("10001")
    at5.run()
    resolve(at5, gid5b)
    check("no exception", not at5.exception, at5.exception[0].message[:300] if at5.exception else "")
    final = target_dmas(at5)
    check("New York does NOT come back even though a second group resolved to it "
          "(monotone add-only, not replace-while-clean)",
          "New York City" not in final, final)
    check("Philadelphia (never removed) is still there", "Philadelphia" in final, final)

    print("\nthe SAME ledger, from drafting's side, Phase 6: Target DMAs is now "
          "seed-if-empty only -- apply_group_markets_autofill's monotone-add union "
          "(above) still applies to a resolved group, but once Target DMAs is "
          "non-empty (Philadelphia, from the avails resolution), drafting never "
          "writes it again at all, not even a union. A notes-named market not "
          "already there (Denver) is flagged in unresolved_internal instead -- "
          "the old union-with-_group_markets_applied machinery this test used to "
          "exercise is gone outright, which is what makes the ledger bug it "
          "guarded against unreachable rather than guarded against.")
    draft_preset = {
        "market_choice": at5.session_state["market_choice"],
        "flight_start": at5.session_state["flight_start"],
        "flight_end": at5.session_state["flight_end"],
        "target_dma_choice": final,
        "_group_markets_applied": set(at5.session_state["_group_markets_applied"]),
    }
    draft = {
        "client_name": "Test Co", "vertical": "none", "market": "DC",
        "flight_start": "2026-09-01", "flight_end": "2026-09-30",
        "total_budget": 5000, "breakout": "monthly",
        "target_markets": ["Denver"],
        "media_plan_lines": [{"product": "premion_streaming_tv",
                              "allocation": {"percent_of_total": 100}}],
        "audiences": [], "attribution": [], "sports": [],
        "campaign_specs": {}, "unresolved": [], "unresolved_internal": [],
    }
    drafted_state = apply_draft(draft, preset_state=draft_preset)
    drafted_dmas = drafted_state.get("target_dma_choice") or []
    check("Denver (the newly drafted market) is NOT added -- Target DMAs was "
          "already non-empty (Philadelphia), so drafting only flags it now",
          not any("Denver" in v for v in drafted_dmas), drafted_dmas)
    # Flagged as a seller-facing follow-up card now ("Add Denver"), with the
    # prose list as before for anything that isn't one (draft_followups).
    internal_notes = (list(drafted_state.get("draft_unresolved_internal") or [])
                      + [i["question"] for i in drafted_state.get("draft_review_items") or []
                         if i.get("audience") == "internal"])
    check("Denver is flagged for the seller instead of silently added",
          any("Denver" in n and "Target DMAs" in n for n in internal_notes), internal_notes)
    check("Philadelphia (avails-resolved, never removed) still there -- drafting "
          "never touched a non-empty Target DMAs at all",
          "Philadelphia" in drafted_dmas, drafted_dmas)
    check("New York stays gone -- the rep's earlier removal is still respected",
          "New York City" not in drafted_dmas, drafted_dmas)
    geography_text = drafted_state.get("geography_text") or ""
    check("Campaign Specs Geography names both markets too",
          "Denver" in geography_text and "Philadelphia" in geography_text,
          geography_text)

    print("\nbroadcast row Geo stays station-derived throughout -- re-asserted, the invariant "
          "most likely to regress from a Section-A ordering mistake")
    fixture = ROOT / "tests" / "fixtures" / "wideorbit" / "regency_planner.xls"
    if not fixture.exists():
        print(f"  SKIP  {fixture} not present (real client schedules are gitignored).")
    else:
        import wideorbit
        from streamlit.testing.v1 import AppTest
        schedule = wideorbit.parse_schedule(str(fixture), fixture.name)
        at6 = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
        at6.session_state["authed"] = True
        at6.session_state["current_user"] = "T"
        at6.session_state["total_tv"] = True
        at6.session_state["market_choice"] = "DC"
        at6.session_state["flight_start"] = date(2026, 9, 1)
        at6.session_state["flight_end"] = date(2026, 11, 30)
        at6.session_state["broadcast_schedule"] = schedule
        at6.session_state["include_avails_template"] = True
        at6.session_state["avails_seed_rows"] = [{"Audience": "Homeowners", "Geo": "Somerset NJ", COL: 100000}]
        at6.run()
        check("no exception", not at6.exception, at6.exception[0].message[:300] if at6.exception else "")
        gid6 = real_group_id(at6)
        set_mode(at6, gid6, "Counties")
        [w for w in at6.text_area if str(w.key) == f"geo_counties_{gid6}"][0].set_value("Somerset NJ")
        at6.run()
        resolve(at6, gid6)
        check("no exception after resolving alongside a broadcast import",
              not at6.exception, at6.exception[0].message[:300] if at6.exception else "")
        rows = at6.session_state["plan_options"][0]["rows"]
        broadcast = [r for r in rows if app.is_broadcast_row(r)]
        check("exactly one broadcast line", len(broadcast) == 1, len(broadcast))
        check("broadcast Geo is still station-derived, untouched by any group resolution",
              broadcast[0]["Geo"] == "Washington DC DMA", broadcast[0]["Geo"])

    print("\ninstall_market_lookup surfaces WHY it failed, not just that market "
          "auto-resolution is silently dead")
    # A live report hit exactly this: every market auto-resolution dead, no
    # signal why beyond a per-resolve-click note that reads like the
    # resolver doing its job. Three failure modes, each with its own text --
    # monkeypatched at the market_lookup module boundary rather than by
    # touching the real committed file, and install_market_lookup.clear()
    # between cases since it's an @st.cache_resource function (one cached
    # result per process otherwise).
    import market_lookup as ml
    real_available, real_install = ml.available, ml.install

    def restore():
        ml.available, ml.install = real_available, real_install
        app.install_market_lookup.clear()

    try:
        app.install_market_lookup.clear()
        ml.available = lambda: False
        lookup, warning = app.install_market_lookup()
        check("file genuinely missing: a specific, actionable warning",
              lookup is None and warning is not None and "isn't present" in warning, warning)

        app.install_market_lookup.clear()
        ml.available = lambda: True
        ml.install = lambda: (_ for _ in ()).throw(OSError("simulated corrupt gzip"))
        lookup, warning = app.install_market_lookup()
        check("install() raising: caught and reported, not an uncaught exception",
              lookup is None and warning is not None and "failed to load" in warning
              and "simulated corrupt gzip" in warning, warning)

        app.install_market_lookup.clear()
        ml.available = lambda: True
        ml.install = lambda: geo_resolver.TableMarketLookup(by_county={}, name="empty")
        lookup, warning = app.install_market_lookup()
        check("loaded but empty: reported as empty, not silently treated as healthy",
              lookup is None and warning is not None and "empty" in warning, warning)

        app.install_market_lookup.clear()
        ml.available, ml.install = real_available, real_install
        lookup, warning = app.install_market_lookup()
        check("the real, healthy table: no warning at all",
              lookup is not None and len(lookup) > 0 and warning is None, warning)
    finally:
        restore()


    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Scenario 0's own inputs resolve as the scenario doc says, the Section-A autofill is "
          "monotone add-only, and a broadcast row's Geo is never touched by any of it. A "
          "missing, corrupt or empty market lookup is reported specifically, not left to "
          "degrade silently into a per-resolve-click note.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
