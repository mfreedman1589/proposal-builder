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
import geo_resolver  # noqa: E402

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
    [w for w in at3.text_area if str(w.key) == f"geo_radius_centers_{gid3}"][0].set_value("20005")
    at3.run()
    [w for w in at3.text_input if str(w.key) == f"geo_radius_miles_{gid3}"][0].set_value("10")
    at3.run()
    resolve(at3, gid3)
    check("no exception", not at3.exception, at3.exception[0].message[:300] if at3.exception else "")
    g3 = group_by_id(at3, gid3)
    check("resolves to Washington-Hagerstown around a DC zip",
          g3["resolved_markets"] == ["washington_hagerstown"], g3["resolved_markets"])
    check("resolved_zips is non-trivial", len(g3["resolved_zips"]) > 10, len(g3["resolved_zips"]))

    print("\nRadius mode with 20 centers: a deduplicated union, one bad address named "
          "while the other 19 still resolve, and a per-line radius override")
    # 18 plain DC-area zips (instant, local -- no network call) plus one
    # zip with a ", <miles>" override and one deliberately bad address --
    # 20 lines total, 19 of which are real. 20005 and 20006 are adjacent
    # enough that their 10mi circles overlap heavily on purpose, so the
    # union coming back smaller than the sum of their individual counts is
    # the actual proof overlapping radii don't double-count a zip.
    zips_pool = ["20005", "20006", "20001", "20002", "20003", "20004", "20007", "20008",
                "20009", "20010", "20011", "20012", "20015", "20016", "20017", "20018",
                "20019", "20020"]
    bad_address = "NOT-A-REAL-ADDRESS-XYZQQQ-999999"
    lines = "\n".join(zips_pool + ["20024, 2", bad_address])
    at3e = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid3e = real_group_id(at3e)
    set_mode(at3e, gid3e, "Radius")
    [w for w in at3e.text_area if str(w.key) == f"geo_radius_centers_{gid3e}"][0].set_value(lines)
    at3e.run()
    [w for w in at3e.text_input if str(w.key) == f"geo_radius_miles_{gid3e}"][0].set_value("10")
    at3e.run()
    resolve(at3e, gid3e)
    check("no exception resolving 20 centers", not at3e.exception,
          at3e.exception[0].message[:300] if at3e.exception else "")
    g3e = group_by_id(at3e, gid3e)
    check("all 20 lines parsed as centers, in geo_def",
          len(g3e["geo_def"]["centers"]) == 20, len(g3e["geo_def"]["centers"]))
    warnings3e = [str(w.value) if hasattr(w, "value") else str(w) for w in at3e.warning]
    check("the bad address is reported BY NAME, not just a count",
          any(bad_address in w for w in warnings3e), warnings3e)
    check("exactly the one bad address is unresolved -- the other 19 are not swept in with it",
          any(f"1 entr" in w and bad_address in w for w in warnings3e), warnings3e)
    check("the union is deduplicated: smaller than the naive sum of all 20 centers' own counts "
          "(20005 and 20006 alone overlap on purpose)",
          len(g3e["resolved_zips"]) < 139 * 1.5, len(g3e["resolved_zips"]))
    check("still a real, non-trivial union", len(g3e["resolved_zips"]) > 100, len(g3e["resolved_zips"]))
    check("resolves to Washington-Hagerstown, same as any one of these DC zips alone",
          g3e["resolved_markets"] == ["washington_hagerstown"], g3e["resolved_markets"])

    print("\na per-line radius override applies only to its own line")
    override_center = "20024"
    other_center = "20005"
    lines_override_only = "\n".join([other_center, f"{override_center}, 2"])
    at3f = new_app([{"Audience": "Homeowners", "Geo": "Denver", COL: 100000}])
    gid3f = real_group_id(at3f)
    set_mode(at3f, gid3f, "Radius")
    [w for w in at3f.text_area if str(w.key) == f"geo_radius_centers_{gid3f}"][0].set_value(
        lines_override_only)
    at3f.run()
    [w for w in at3f.text_input if str(w.key) == f"geo_radius_miles_{gid3f}"][0].set_value("10")
    at3f.run()
    resolve(at3f, gid3f)
    check("no exception", not at3f.exception, at3f.exception[0].message[:300] if at3f.exception else "")
    g3f = group_by_id(at3f, gid3f)
    check("geo_def keeps the un-overridden center as a plain string (byte-identical to before "
          "this feature, for any proposal that never used an override)",
          g3f["geo_def"]["centers"][0] == other_center, g3f["geo_def"]["centers"])
    check("geo_def keeps the overridden center as its own {center, miles}",
          g3f["geo_def"]["centers"][1] == {"center": override_center, "miles": 2.0},
          g3f["geo_def"]["centers"])
    check("the group's own default miles is untouched by the per-line override",
          g3f["geo_def"]["miles"] == 10.0, g3f["geo_def"]["miles"])
    # Independent confirmation the override actually took effect in the
    # RESOLUTION, not just in the stored geo_def: 20024 alone at 2mi is a
    # small, specific zip count (27, confirmed directly against geo_resolver
    # while building this test) -- resolving the two-line group and the
    # override-only line separately and comparing their union proves the
    # override line contributed its 2mi answer, not the group's 10mi default.
    solo_override = geo_resolver.radius_to_zips([override_center], 2)
    solo_default = geo_resolver.radius_to_zips([override_center], 10)
    check("the override radius (2mi) genuinely differs from the default (10mi) for this center "
          "-- otherwise this test would prove nothing",
          len(solo_override.resolved) < len(solo_default.resolved), (
              len(solo_override.resolved), len(solo_default.resolved)))
    other_solo = geo_resolver.radius_to_zips([other_center], 10)
    expected_union = set(solo_override.resolved) | set(other_solo.resolved)
    check("the union matches 20005 at the DEFAULT 10mi plus 20024 at its OWN 2mi override "
          "-- not 20024 at 10mi, which would silently ignore the override",
          set(g3f["resolved_zips"]) == expected_union, (
              len(g3f["resolved_zips"]), len(expected_union)))

    print("\nScenario 0's own exact inputs (targeting_groups_test_scenarios.md), not just an "
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
    warnings3b = [str(w.value) if hasattr(w, "value") else str(w) for w in at3b.warning]
    check("99999 surfaces in the UI warning", any("99999" in w for w in warnings3b), warnings3b)

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
    print("Each geo kind resolves through geo_resolver, unresolved entries surface rather than "
          "vanish, the Section-A autofill is monotone add-only, and a broadcast row's Geo is "
          "never touched by any of it. A missing, corrupt or empty market lookup is reported "
          "specifically, not left to degrade silently into a per-resolve-click note.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
