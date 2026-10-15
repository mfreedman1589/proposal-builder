"""Several avails files in one proposal, one geography per real geography,
the audience x geography grid, the coverage map and its color control, and
the plan checks -- against the real Adventist Health pair (RFPID-269771 +
RFPID-269781, 2026-10-15).

What the two documents say, read off the PDFs themselves (not off anything
this app computes):
  - RFPID-269771 states Total Impressions 8,075,336; RFPID-269781 states
    2,066,550 -- 10,141,886 together.
  - RFPI-4943549 and RFPI-4943552 are the same line twice: HLTH CRX
    Neurology, the Silver Spring 50 Mile Radius ZIP option, 313,680 each.
    Without the repeat: 9,828,206.
  - Three geographies: a ZIP list around 12501 Prosperity Dr, Silver Spring
    (with 02191 -- Quincy, MA -- in it, almost certainly 22191), the
    Washington, D.C. DMA, and one ZIP list printed as both "Eastern West
    Virginia Delaware Southern Pennsylvania" and "Eastern WV Delaware
    Southern Pennsylvania".
  - Neurology has no DMA line.

The PDFs carry real client pricing and are gitignored; this file reports
SKIP without them.

    python tests/test_avails_multi_file.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import assembly  # noqa: E402
import avails_grid as ag  # noqa: E402
import avails_pdf_import as api  # noqa: E402
import geo_resolver  # noqa: E402
import market_lookup  # noqa: E402
import targeting_map as tm  # noqa: E402

market_lookup.install()

DOC_A = next(REPO.glob("Premion Media Plan_RFPID-269771_*Adventist Health*.pdf"), None)
DOC_B = next(REPO.glob("Premion Media Plan_RFPID-269781_*Adventist Health*.pdf"), None)
STATED_A, STATED_B = 8_075_336, 2_066_550
WITH_DUPLICATE, WITHOUT_DUPLICATE = 10_141_886, 9_828_206
SILVER_SPRING = "Silver Spring 50-mi radius"
EASTERN = "Eastern WV Delaware Southern Pennsylvania"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    def __init__(self, state=None):
        self.session_state = dict(state or {})
        self.secrets = {}


def import_both():
    """Both documents through the real parser and the real
    app.apply_avails_import + freeze, as the uploader runs them."""
    real = app.st
    groups, docs = [], []
    try:
        for path in (DOC_A, DOC_B):
            document = api.parse_avails_pdf(str(path), path.name)
            app.st = _StubSt({"targeting_groups": groups})
            new_groups, _report = app.apply_avails_import(document)
            app._freeze_avails_import_figures(new_groups, document)
            groups = groups + new_groups
            docs.append({"rfpid": document.rfpid, "total_impressions": document.total_impressions})
            ag.unify_imported_geo_names(groups)
    finally:
        app.st = real
    return groups, docs


def full_flight(group):
    return int(group.get("avails_full_flight") or 0)


def pure_checks():
    print("parsing: stated totals, the line ids, the street number that isn't a ZIP")
    a = api.parse_avails_pdf(str(DOC_A), DOC_A.name)
    b = api.parse_avails_pdf(str(DOC_B), DOC_B.name)
    check("RFPID-269771 parses to its stated 8,075,336",
          a.total_impressions == STATED_A and sum(g.impressions for g in a.groups) == STATED_A)
    check("RFPID-269781 parses to its stated 2,066,550",
          b.total_impressions == STATED_B and sum(g.impressions for g in b.groups) == STATED_B)
    check("each line carries its RFPI", [g.rfpi for g in a.groups][2] == "RFPI-4943549",
          [g.rfpi for g in a.groups])
    radius = a.groups[0]
    check("12501 (the street number in '12501 Prosperity Dr') is not read as a ZIP",
          "12501" not in radius.zips and "20904" in radius.zips, len(radius.zips))

    print("\nlabels: cleaned from the document's words, never from ZIP states")
    check("radius address -> 'Silver Spring 50-mi radius'",
          ag.clean_geo_label("12501 Prosperity Dr Silver Spring MD 20904 50 Mile Radius", 50.0)
          == SILVER_SPRING)
    check("'Zip Option -' is stripped",
          ag.clean_geo_label("Zip Option - Eastern WV Delaware Southern Pennsylvania") == EASTERN)
    check("a DMA name is left as-is", ag.clean_geo_label("Washington, D.C.") == "Washington, D.C.")

    groups, docs = import_both()
    check("9 lines across both files", len(groups) == 9, len(groups))
    eastern = [g for g in groups if g.get("avails_source") == "RFPID-269781"]
    check("the ZIP list named two ways shows one (shortest) name on every line",
          {g["name"] for g in eastern} == {EASTERN}, {g["name"] for g in eastern})
    check("...and the two spellings are still on record for the check",
          {g["avails_geo_variant"] for g in eastern}
          == {EASTERN, "Eastern West Virginia Delaware Southern Pennsylvania"})
    check("the two lines of one ZIP list are one geography",
          len({ag.geo_identity(g) for g in eastern}) == 1)

    print("\nduplicates and totals")
    dups = ag.duplicate_groups(groups)
    by_id = {g["id"]: g for g in groups}
    check("exactly one duplicate: RFPI-4943552 repeats RFPI-4943549",
          len(dups) == 1 and [(by_id[d]["avails_rfpi"], by_id[o]["avails_rfpi"])
                              for d, o in dups.items()] == [("RFPI-4943552", "RFPI-4943549")], dups)
    for include, expected in ((False, WITHOUT_DUPLICATE), (True, WITH_DUPLICATE)):
        grid = ag.build_grid(ag.active_groups(groups, include), full_flight)
        check(f"grid total {expected:,} with the duplicate {'included' if include else 'excluded'}",
              grid["total"] == expected, grid["total"])
        check("...and it reconciles to the documents' stated totals",
              ag.reconcile(groups, docs, include) is None, ag.reconcile(groups, docs, include))
    check("both files' stated totals add to 10,141,886",
          sum(d["total_impressions"] for d in docs) == WITH_DUPLICATE)
    tampered = [dict(g) for g in groups]
    tampered[0]["avails_doc_full_flight"] = 1
    check("a total that doesn't tie is reported", ag.reconcile(tampered, docs, False) is not None)

    print("\nthe grid: rows are audiences, columns are geographies + Total, never one row per line")
    active = ag.active_groups(groups, False)
    payload = ag.grid_payload(ag.build_grid(active, full_flight, app._market_display_name))
    check("3 audiences x 3 geographies", len(payload["rows"]) == 3 and len(payload["geo_headers"]) == 3,
          (len(payload["rows"]), payload["geo_headers"]))
    check("geography columns carry the clean labels",
          payload["geo_headers"][0] == SILVER_SPRING and payload["geo_headers"][2] == EASTERN,
          payload["geo_headers"])
    neuro = next(r for r in payload["rows"] if "Neurology" in r["audience"])
    check("'—' where Neurology doesn't run (the DMA)", neuro["cells"][1] == ag.DASH, neuro["cells"])
    check("column totals add to the grand total",
          sum(int(t.replace(",", "")) for t in payload["totals"]) == WITHOUT_DUPLICATE)

    print("\nplan checks")
    checks = ag.plan_checks(groups, geo_resolver._data()["zip_points"], app._market_display_name)
    check("duplicate named by both RFPIs",
          any("RFPI-4943549" in x and "RFPI-4943552" in x for x in checks["duplicates"]),
          checks["duplicates"])
    check("02191 flagged, with 22191 suggested",
          any("02191" in x and "22191" in x for x in checks["zips"]), checks["zips"])
    check("only 02191 is flagged (real PO-box and outlying ZIPs are not)",
          checks["excluded_zips"] == {"02191"}, checks["excluded_zips"])
    check("the two names for one ZIP list are reported", len(checks["names"]) == 1, checks["names"])
    check("overlapping geographies for the same audience are reported",
          any(SILVER_SPRING in x and "Doctors and Nurses" in x for x in checks["overlaps"]),
          checks["overlaps"])

    print("\nthe map: Audience mode keys only the colors drawn, never by a place")
    layers = tm.coverage_layers(active, ag.MAP_MODE_AUDIENCE, app._market_display_name,
                                checks["excluded_zips"])
    labels = [e["label"] for e in layers["entries"]]
    check("Adventist draws 2 colors: 'All 3 audiences' and 'Doctors + Heart'",
          labels == ["All 3 audiences", "Doctors + Heart"], labels)
    geo_names = set(ag.canonical_geo_labels(groups, app._market_display_name).values())
    check("no key entry is a geography name", not (set(labels) & geo_names), labels)
    drawn = set(layers["zip_key"].values()) | set(layers["county_key"].values())
    check("every key entry is a color actually on the map, and vice versa",
          drawn == {e["key"] for e in layers["entries"]})
    check("02191 isn't drawn (it would zoom the map to New England)",
          "02191" not in layers["zip_key"])
    combos = len(layers["entries"])
    mode = ag.resolve_map_mode(ag.MAP_MODE_AUTO, 3, 3, combos)
    check("Auto picks Audience for Adventist", mode == ag.MAP_MODE_AUDIENCE, mode)

    one_audience = [g for g in active if "Doctors and Nurses" in g["terms"][0]]
    region = tm.coverage_layers(one_audience, ag.MAP_MODE_REGION, app._market_display_name)
    check("1 audience x 3 regions: Auto picks Region",
          ag.resolve_map_mode(ag.MAP_MODE_AUTO, 1, 3, 1) == ag.MAP_MODE_REGION)
    check("...with 3 colors keyed by region", len(region["entries"]) == 3,
          [e["label"] for e in region["entries"]])
    check("more than 4 combination colors: Auto falls back to Region",
          ag.resolve_map_mode(ag.MAP_MODE_AUTO, 4, 5, 5) == ag.MAP_MODE_REGION)
    same_area = [g for g in active if g.get("avails_source") == "RFPID-269781"]
    single = tm.coverage_layers(same_area, ag.MAP_MODE_AUDIENCE, app._market_display_name)
    check("every audience over the same area -> one color keyed 'All audiences'",
          [e["label"] for e in single["entries"]] == ["All audiences"],
          [e["label"] for e in single["entries"]])

    png_a, entries_a = tm.render_coverage_map(active, ag.MAP_MODE_AUDIENCE,
                                              label_for=app._market_display_name)
    png_r, entries_r = tm.render_coverage_map(active, ag.MAP_MODE_REGION,
                                              label_for=app._market_display_name)
    check("switching mode re-renders the map and the key",
          png_a and png_r and png_a != png_r
          and [e["label"] for e in entries_a] != [e["label"] for e in entries_r])
    custom_key = entries_a[0]["key"]
    png_c, entries_c = tm.render_coverage_map(active, ag.MAP_MODE_AUDIENCE,
                                              colors={custom_key: "#123456"},
                                              label_for=app._market_display_name)
    check("a custom key color is drawn", png_c != png_a and entries_c[0]["color"] == "#123456")

    print("\nthe slide: the grid replaces one-row-per-line")
    fill = {"client_name": "Adventist Health", "proposal_title": "", "vertical_display": "Healthcare",
            "logo_path": "placeholder_logo.png",
            "avails": {"rows": [{"audience": "x", "geo": "y", "avails": "1"}],
                       "total_avails": payload["total"], "label": "Max Monthly Avails",
                       "map_png": png_a, "grid": payload},
            "media_plan_options": [{}]}
    master = REPO / "TEGNA_MASTER_DECK_10-08-26_gdw.pptx"
    if master.exists():
        prs, _warnings = assembly.build_avails_only_deck(str(master), fill, map_present=True)
        table = next(s.table for s in prs.slides[0].shapes if s.has_table)
        texts = [[c.text for c in r.cells] for r in table.rows]
        check("slide table is 5 columns: Audience, 3 geographies, Total",
              len(table.columns) == 5 and texts[0][1] == SILVER_SPRING
              and texts[0][4] == "Max Monthly Avails", texts[0])
        check("one row per audience plus header and Total", len(table.rows) == 5, len(table.rows))
        check("Total row reads 9,828,206", texts[-1][-1] == "9,828,206", texts[-1])
        check("no geography header is assembled from ZIP states",
              not any(h in ("MD", "VA", "PA", "WV", "DE", "DC") for h in texts[0]), texts[0])
    else:
        print("  SKIP  slide build -- master deck not present")


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["target_dmas"] = ["Washington, DC"]
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 11, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.session_state["client_name"] = "Adventist Health"
    return at


def ss(at, key, default=None):
    return at.session_state[key] if key in at.session_state else default


def app_checks():
    print("\nthe real app: two uploads, both listed")
    at = new_app()
    at.session_state["avails_pdf_upload_path_intake"] = str(DOC_A)
    at.run()
    at.session_state["avails_pdf_upload_path_intake"] = str(DOC_B)
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:600] if at.exception else "")
    docs = ss(at, "avails_documents") or []
    check("both files listed", [d["rfpid"] for d in docs] == ["RFPID-269771", "RFPID-269781"], docs)
    shown = " ".join(m.value for m in at.markdown)
    check("each listed with its line count and total",
          "RFPID-269771** — 6 lines, 8,075,336" in shown and "RFPID-269781** — 3 lines, 2,066,550" in shown)
    captions = " ".join(c.value for c in at.caption)
    check("combined total 10,141,886", "Combined: **10,141,886** impressions across 2 files" in captions)
    check("a Remove button per file",
          {b.key for b in at.button if str(b.key).startswith("avails_doc_remove_")}
          == {"avails_doc_remove_RFPID-269771", "avails_doc_remove_RFPID-269781"})

    toggle = [c for c in at.checkbox if c.key == "avails_include_duplicates"]
    check("duplicate toggle offered, off by default, naming both totals",
          toggle and not toggle[0].value and "10,141,886" in toggle[0].label
          and "9,828,206" in toggle[0].label, [c.label for c in toggle])
    warnings = " ".join(w.value for w in at.warning)
    for needle, what in (("RFPI-4943549", "duplicate"), ("02191", "stray ZIP"),
                         ("22191", "suggested ZIP"), ("named 2 ways", "two names"),
                         ("share", "overlapping geographies")):
        check(f"plan check fires: {what}", needle in warnings, warnings[:300])
    pickers = [p.label for p in at.color_picker]
    check("Auto colors by audience: key swatches are the 2 combinations",
          pickers == ["All 3 audiences", "Doctors + Heart"], pickers)

    if toggle:
        toggle[0].check().run()
        pickers_with = [p.label for p in at.color_picker]
        check("counting the duplicate changes nothing on the map (Neurology still in Silver Spring)",
              pickers_with == ["All 3 audiences", "Doctors + Heart"], pickers_with)
        at.checkbox(key="avails_include_duplicates").uncheck().run()

    at.radio(key="map_color_mode").set_value("Region").run()
    pickers = [p.label for p in at.color_picker]
    check("switching to Region re-keys the swatches by region",
          pickers == [SILVER_SPRING, "Washington, DC", EASTERN], pickers)
    first_region = at.color_picker[0]
    first_region.set_value("#123456").run()
    colors = ss(at, "map_colors") or {}
    check("a picked color is saved", "#123456" in [str(c).lower() for c in colors.values()], colors)

    print("\nGenerate, then reopen: mode and custom colors survive")
    captured = {}
    real_log = db.log_proposal

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        captured["row"] = {"id": "00000000-0000-0000-0000-000000000000", "client_name": client_name,
                           "vertical": vertical, "market": market, "form_json": form_json,
                           "deck_version_id": kwargs.get("deck_version_id"),
                           "output_filename": kwargs.get("output_filename"),
                           "parent_proposal_id": None, "revision_label": None,
                           "logo_storage_path": None, "created_by": None}
        return "00000000-0000-0000-0000-000000000000", None

    db.log_proposal = spy_log
    try:
        generate = [b for b in at.button if b.label == "Generate proposal"]
        check("Generate is offered", bool(generate))
        if generate:
            generate[0].click().run()
    finally:
        db.log_proposal = real_log
    check("Generate ran clean", not at.exception, at.exception[0].message[:600] if at.exception else "")
    row = captured.get("row")
    if row:
        form = row["form_json"]
        check("logged: the grid", form.get("avails_grid") and form["avails_grid"]["total"] == "9,828,206",
              (form.get("avails_grid") or {}).get("total"))
        check("test mode, client name without QA-TEST-: logged flagged as a test",
              form.get("is_test") is True, form.get("is_test"))
        check("a flagged proposal is hidden from Proposal history",
              db.is_test_proposal(row) and not db.is_test_proposal({"form_json": {}}))
        check("logged: mode and colors",
              form["map_color"]["choice"] == "Region"
              and "#123456" in [str(c).lower() for c in form["map_color"]["colors"].values()],
              form.get("map_color"))
        real = app.st
        stub = _StubSt()
        app.st = stub
        try:
            app.rehydrate_proposal_into_form(row, parent_proposal_id=row["id"],
                                             revision_default="Revision 2")
        finally:
            app.st = real
        state = stub.session_state
        check("reopened: Region is still the choice", state.get("map_color_mode") == "Region",
              state.get("map_color_mode"))
        check("reopened: the custom color is still there",
              "#123456" in [str(c).lower() for c in (state.get("map_colors") or {}).values()],
              state.get("map_colors"))
        check("reopened: both files still listed",
              [d["rfpid"] for d in state.get("avails_documents") or []]
              == ["RFPID-269771", "RFPID-269781"])
        real_logo = db.proposal_logo
        db.proposal_logo = lambda path: path
        try:
            buffer, _name, rebuild_warnings = app.rebuild_proposal_deck(row)
        finally:
            db.proposal_logo = real_logo
        check("rebuild-as-presented draws a deck", buffer is not None, rebuild_warnings)
        check("rebuild redraws the same coverage map", app._rebuild_map_png(form) is not None)

    print("\nremoving a file whose audiences are on the plan: confirm, then they come off")
    add_all = [b for b in at.button if str(b.key).startswith("avails_include_all_")]
    check("Add all to plan offered", bool(add_all))
    if add_all:
        add_all[0].click().run()

    def plan_ids():
        return {gid for opt in ss(at, "plan_options") or [] for r in opt["rows"]
                for gid in app.group_ids_of(r)}
    groups_b = {g["id"] for g in ss(at, "targeting_groups") or []
                if g.get("avails_source") == "RFPID-269781"}
    check("RFPID-269781's 3 lines are on the plan", groups_b and groups_b <= plan_ids(),
          (groups_b, plan_ids()))
    remove = at.button(key="avails_doc_remove_RFPID-269781")
    check("Remove stays enabled while the file's lines are on the plan", not remove.disabled)
    remove.click().run()
    warnings = " ".join(w.value for w in at.warning)
    check('asks first: "This will remove 3 audiences from the plan"',
          "This will remove 3 audiences from the plan" in warnings, warnings[:300])
    check("nothing removed before confirming",
          len(ss(at, "avails_documents") or []) == 2 and groups_b <= plan_ids())
    at.button(key="avails_doc_remove_no_RFPID-269781").click().run()
    check("Cancel keeps everything", len(ss(at, "avails_documents") or []) == 2
          and "This will remove" not in " ".join(w.value for w in at.warning),
          (len(ss(at, "avails_documents") or []), [w.value[:80] for w in at.warning]))
    at.button(key="avails_doc_remove_RFPID-269781").click().run()
    at.button(key="avails_doc_remove_yes_RFPID-269781").click().run()
    check("no exception", not at.exception, at.exception[0].message[:600] if at.exception else "")
    check("its plan lines are gone", not (groups_b & plan_ids()), plan_ids())
    groups = [g for g in (ss(at, "targeting_groups") or []) if g["terms"]]
    check("6 lines left, none from the removed file",
          len(groups) == 6 and not any(g.get("avails_source") == "RFPID-269781" for g in groups),
          len(groups))
    check("the list shows one file", [d["rfpid"] for d in ss(at, "avails_documents") or []]
          == ["RFPID-269771"])


def main():
    if not (DOC_A and DOC_B and DOC_A.exists() and DOC_B.exists()):
        print("SKIP -- the Adventist Health avails PDFs aren't present (gitignored fixtures)")
        return 0
    pure_checks()
    app_checks()
    print(f"\n{'FAILED: ' + str(len(failures)) if failures else 'ALL PASSED'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
