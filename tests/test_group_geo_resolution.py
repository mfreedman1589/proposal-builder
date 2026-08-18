"""Phase 6 of the targeting-groups roadmap (geo_targeting_roadmap.md D): the
per-group geo-definition expander and the Section-A market autofill it
feeds.

    python tests/test_group_geo_resolution.py

Driven through the real form with AppTest, same reason as every other group
test in this suite -- `st.data_editor` can't be driven directly, and what
matters here is what a real Resolve click actually writes into
`targeting_groups`, not what a helper returns in isolation.
"""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
COL = app.AVAILS_COLUMN_MONTHLY
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app(avails_rows):
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["avails_seed_rows"] = avails_rows
    at.run()
    return at


def real_group_id(at, exclude=()):
    return next(g["id"] for g in at.session_state["targeting_groups"]
               if g.get("terms") and g["id"] not in exclude)


def set_mode(at, gid, mode):
    [w for w in at.radio if str(w.key) == f"geo_mode_{gid}"][0].set_value(mode)
    at.run()


def resolve(at, gid):
    [w for w in at.button if str(w.key) == f"geo_resolve_btn_{gid}"][0].click()
    at.run()


def group_by_id(at, gid):
    return next(g for g in at.session_state["targeting_groups"] if g["id"] == gid)


def main():
    print("Counties mode resolves, through counties_to_fips + counties_to_zips + zips_to_markets")
    at = new_app([{"Audience": "Homeowners", "Geo": "Somerset NJ", COL: 100000}])
    check("the form renders", not at.exception, at.exception[0].message[:300] if at.exception else "")
    gid = real_group_id(at)
    set_mode(at, gid, "Counties")
    [w for w in at.text_area if str(w.key) == f"geo_counties_{gid}"][0].set_value("Somerset NJ")
    at.run()
    resolve(at, gid)
    check("no exception", not at.exception, at.exception[0].message[:300] if at.exception else "")
    g = group_by_id(at, gid)
    check("geo_def is kind:counties", g["geo_def"] == {"kind": "counties", "counties": ["Somerset NJ"]}, g["geo_def"])
    check("resolved to real zips", len(g["resolved_zips"]) > 0, g["resolved_zips"])
    check("resolved to New York and Philadelphia (the roadmap's own cross-check example)",
          set(g["resolved_markets"]) == {"new_york", "philadelphia"}, g["resolved_markets"])

    print("\nan untouched rerun does not disturb the resolution (the fold-back fix)")
    before = dict(g)
    at.run()
    after = group_by_id(at, gid)
    check("geo_def/resolved_zips/resolved_markets survive byte-for-byte",
          after["geo_def"] == before["geo_def"] and after["resolved_zips"] == before["resolved_zips"]
          and after["resolved_markets"] == before["resolved_markets"], after)

    print("\nZips mode resolves, and reports an unresolvable zip rather than dropping it")
    at2 = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid2 = real_group_id(at2)
    set_mode(at2, gid2, "Zips")
    [w for w in at2.text_area if str(w.key) == f"geo_zips_{gid2}"][0].set_value("80202, 80203, 99999")
    at2.run()
    resolve(at2, gid2)
    check("no exception", not at2.exception, at2.exception[0].message[:300] if at2.exception else "")
    g2 = group_by_id(at2, gid2)
    check("all three zips kept in resolved_zips -- nothing silently dropped",
          g2["resolved_zips"] == ["80202", "80203", "99999"], g2["resolved_zips"])
    check("resolves to Denver", g2["resolved_markets"] == ["denver"], g2["resolved_markets"])
    unresolved_warnings = [str(w.value) if hasattr(w, "value") else str(w) for w in at2.warning]
    check("the unresolvable zip surfaces in the UI, not silently dropped",
          any("99999" in w for w in unresolved_warnings), unresolved_warnings)

    print("\nRadius mode resolves through the Census geocoder / zip centroid")
    at3 = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid3 = real_group_id(at3)
    set_mode(at3, gid3, "Radius")
    [w for w in at3.text_input if str(w.key) == f"geo_radius_center_{gid3}"][0].set_value("20005")
    at3.run()
    [w for w in at3.text_input if str(w.key) == f"geo_radius_miles_{gid3}"][0].set_value("10")
    at3.run()
    resolve(at3, gid3)
    check("no exception", not at3.exception, at3.exception[0].message[:300] if at3.exception else "")
    g3 = group_by_id(at3, gid3)
    check("resolves to Washington-Hagerstown around a DC zip",
          g3["resolved_markets"] == ["washington_hagerstown"], g3["resolved_markets"])
    check("resolved_zips is non-trivial", len(g3["resolved_zips"]) > 10, len(g3["resolved_zips"]))

    print("\nMarkets mode needs no resolver call -- a direct pick IS its own resolution")
    at4 = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid4 = real_group_id(at4)
    picker = [w for w in at4.multiselect if str(w.key) == f"geo_markets_{gid4}"][0]
    check("real market names are offered, not raw keys",
          "Denver" in picker.options and "Atlanta" in picker.options, picker.options[:5])
    picker.set_value(["Atlanta", "Denver"])
    at4.run()
    resolve(at4, gid4)
    check("no exception", not at4.exception, at4.exception[0].message[:300] if at4.exception else "")
    g4 = group_by_id(at4, gid4)
    check("geo_def carries the market KEYS, not the display names",
          g4["geo_def"] == {"kind": "markets", "markets": ["atlanta", "denver"]}, g4["geo_def"])
    check("resolved_markets matches the picked keys",
          g4["resolved_markets"] == ["atlanta", "denver"], g4["resolved_markets"])

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

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Each geo kind resolves through geo_resolver, unresolved entries surface rather than "
          "vanish, the Section-A autofill is monotone add-only, and a broadcast row's Geo is "
          "never touched by any of it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
