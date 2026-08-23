"""Shared fixture data for the three real-document scenarios in
targeting_groups_test_scenarios.md: Scenario 2 (Annapolis Cars), Scenario 3
(Visit Hershey & Harrisburg), Scenario 4 (Wilmington University).

Parsed from the three real Premion avails PDFs sitting at the repo root
(gitignored -- they carry real client pricing, same reason the Wide Orbit
fixtures are gitignored) through the real `avails_pdf_import.parse_avails_pdf`
and `app.apply_avails_import` -- the SAME two functions the D2 uploader
itself calls, not a second, hand-typed copy of the same facts. That used to
be exactly that: two independent transcriptions of one document (this
module's own literals, cross-checked in tests/test_avails_pdf_import.py
against the parser's output) -- a real duplication risk, since a document
re-read or a future parser change could drift the two apart with no signal
beyond that one cross-check noticing. Only the four `*_GROUND_TRUTH` figures
below stay hand-copied: each PDF's own "Total Impressions" line on its own
page, external to every parser this app owns, which is exactly what a
ground truth has to be -- asserted against the parsed document's own total
immediately below each constant, so a header-parsing regression is caught
here too, not just in test_avails_pdf_import.py. Every caller should assert
against these directly rather than against `sum(row avails)`, which only
proves this module's own parse agrees with itself.

**Design choice, worth stating plainly:** groups are built through
`app.apply_avails_import` (audience terms via `tg.terms_from_audience_text`,
geography via `app.resolve_group_geography`, DMA/catalog matching via
`market_profiles.match_market` -- the exact call the D2 uploader makes, with
a stubbed `st.session_state` standing in for a running app) rather than
replayed through the ~80 AND/OR/New-group button clicks, or a real
file-upload event, it would take to build 8-12 groups through AppTest one
segment/click at a time. That mechanism is already covered elsewhere
(tests/test_group_builder.py, tests/test_avails_pdf_wiring.py); replaying it
here would spend a lot of runtime re-proving something this suite already
knows. What these scenarios exist to exercise is geography and avails
arithmetic end to end through the real form, against real documents. Avails
are written directly onto each group's `avails_monthly` the same way
`_finish_avails_import` derives it (full-flight impressions / month count),
matching every other avails-editing test in this suite (`st.data_editor`
can't be driven directly, so a data_editor-backed edit is always injected as
state).

A real, honest finding surfaced while building this: several of these real
segment names are not exact matches in the currently-loaded audience
catalog (`AUTO Make Genesis Intender`, the `DEMO Age A21-44`/`A25 Plus`
brackets, `LIFESTAGE Prospective College Students`, `DEMO Career Employed`,
`LIFESTAGE Higher Education Intender` -- either absent, or in Wilmington's
`LIFESTAGE College Planning Parents` case, off by a pluralization the
catalog spells `...Parent`). That's a catalog-vintage question, not a
targeting-groups bug, and `custom_segment_count` treats an unmapped name as
RFP-selectable by default (`rfp_map.get(term, True)`), so it doesn't
silently invent a warning either way -- see ANNAPOLIS_EXPECTED_CUSTOM_COUNT
below for what it actually does with the segments that ARE mapped.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import app                                     # noqa: E402
import avails_pdf_import as api                # noqa: E402
import market_lookup                           # noqa: E402
import targeting_groups as tg                  # noqa: E402

# Registers the county->DMA table with geo_resolver directly (the same call
# app.install_market_lookup() makes from inside a running script) -- needed
# here because this module resolves geography OUTSIDE any Streamlit run.
market_lookup.install()

COL = app.AVAILS_COLUMN_MONTHLY

ANNAPOLIS_PDF = REPO / "Premion Media Plan_RFPID-253813_SR&B Advertising_Annapolis Cars_1-23-2026--ver0.pdf"
HERSHEY_PDF = REPO / "Premion Media Plan_RFPID-260402_Direct - No Agency_Visit Hershey & Harrisburg_4-30-2026--ver0.pdf"
WILMINGTON_PDF = REPO / "Premion Media Plan_RFPID-253956_Direct - No Agency_Wilmington University_1-27-2026--ver0.pdf"

# Transcribed verbatim from RFPID-253813's own "Zip Codes" tables -- NOT
# produced by avails_pdf_import (a bracketed-origin radius group carries no
# zip list of its own to parse; the app resolves 10mi/5mi from the origin
# itself), so this stays an independent, hand-copied check of the app's OWN
# radius resolution against the document's stated zips, compared (not
# required to match exactly) in test_group_scenarios.py.
ANNAPOLIS_ZIPS_10MI = {"20765", "20776", "21012", "21032", "21035", "21037", "21054",
                       "21056", "21108", "21114", "21122", "21140", "21146", "21401",
                       "21402", "21403", "21405", "21409"}
ANNAPOLIS_ZIPS_5MI = {"21012", "21032", "21140", "21401", "21402", "21403", "21405"}


def _flight(start, end):
    all_months = app.month_list(start, end)
    label = app.format_flight_label(all_months, all_months)
    return all_months, label


class _StubSt:
    """apply_avails_import only reads st.session_state (to number a new
    group's color against whatever's already on the campaign) -- a plain
    dict stands in for a running app, the same substitution every other
    module-level test in this suite makes for app.st, so this can resolve
    real geography OUTSIDE any Streamlit run."""
    def __init__(self):
        self.session_state = {}


def _import_groups(document):
    """new_groups exactly as the real D2 uploader would produce them --
    audience terms, geography and market matching all resolved through
    app.apply_avails_import itself, never re-derived here."""
    real, app.st = app.st, _StubSt()
    try:
        new_groups, report = app.apply_avails_import(document)
    finally:
        app.st = real
    return new_groups, report


def _zips_geo(zips_text):
    """Still needed by render_scenarios.py's synthetic St. Louis three-group
    scenario (build_st_louis_three_group) -- that one isn't a real document
    with its own PDF to parse, just real zips grouped by hand to give the
    targeting map something with real geographic overlap to draw, so it
    resolves geography directly through app.resolve_group_geography rather
    than through the importer."""
    return app.resolve_group_geography(app.GEO_MODE_ZIPS, zips_text=zips_text)


def _make_group(terms, op, geo_def, resolved_zips, resolved_markets, avails_monthly, index):
    """Same reason as `_zips_geo` -- render_scenarios.py's synthetic
    scenario builds groups directly, since it has no document/importer to
    build them from."""
    group = tg.new_group(terms, op=op, geo_def=geo_def, avails_monthly=avails_monthly,
                         color=tg.assign_color(index))
    group["resolved_zips"] = resolved_zips
    group["resolved_markets"] = resolved_markets
    return group


def _plan_row(group, product_key, impressions, markup):
    """One media-plan row for one group, built through the real
    line_product_spec/resolve_row_defaults so the tactic label, CPM and
    Flight/Geo/Targeting defaults come from the exact functions the app
    itself uses -- only Impressions/Cost/Type are this fixture's own,
    driven from the document's own full-flight impressions for that group
    (this app has no "impressions = avails" automation; a rep always types
    the number, so a row's Impressions equal to its group's own document
    figure IS how a rep would replicate an already-booked document like
    these).
    """
    label, cpm = app.line_product_spec(product_key)
    audience = tg.audience_label(group)
    geo = tg.geo_label(group, label_for=app._market_display_name)
    defaults = app.resolve_row_defaults(label, geo, audience, group["_flight_label"])
    cost = app.cost_from_impressions(impressions, cpm, markup)
    return {
        "Tactic": label, "Flight": defaults["Flight"], "Geo": defaults["Geo"],
        "Targeting": defaults["Targeting"], "Impressions": float(impressions),
        "CPM": cpm, "Type": app.ROW_TYPE_RATE, "Cost": cost,
        "_group_ids": [group["id"]],
    }


def _build(pdf_path, ground_truth, vertical_choice, expected_custom_count):
    """Shared body for all three scenarios: parse the real PDF, resolve its
    groups through the real importer, spread each group's full-flight
    impressions into avails_monthly the way _finish_avails_import does, and
    build one plan row per group from the document's own per-group figure.
    `agency_involved` and the markup it drives (1.15 gross / 1.0 net -- the
    same `1.15 if agency_involved else 1.0` main() itself uses) are both
    DERIVED from the document's own Agency field, never a second hardcoded
    per-scenario flag.
    """
    document = api.parse_avails_pdf(str(pdf_path))
    assert document.total_impressions == ground_truth, (
        f"{pdf_path.name}'s own stated Total Impressions no longer matches this module's "
        f"hand-copied ground truth -- either the document changed or the header parser "
        f"regressed; check tests/test_avails_pdf_import.py before assuming either.")

    new_groups, report = _import_groups(document)
    all_months, flight_label = _flight(document.flight_start, document.flight_end)
    n_months = max(1, len(all_months))
    agency_involved = bool(document.agency) and "no agency" not in document.agency.lower()
    markup = 1.15 if agency_involved else 1.0

    groups, rows = [], []
    for group, src in zip(new_groups, document.groups):
        full_flight = src.impressions
        group["avails_monthly"] = int(round(full_flight / n_months))
        group["_flight_label"] = flight_label
        groups.append(group)
        rows.append(_plan_row(group, "premion_streaming_tv", full_flight, markup))

    return {
        "name": document.advertiser, "client_name": document.advertiser,
        "flight_start": document.flight_start, "flight_end": document.flight_end,
        "active_months": all_months, "flight_label": flight_label,
        "n_months": n_months, "agency_involved": agency_involved,
        "vertical_choice": vertical_choice,
        "groups": groups, "rows": rows,
        "ground_truth": ground_truth,
        "expected_custom_count": expected_custom_count,
        "import_report": report,
    }


# ===========================================================================
# Scenario 2 -- Annapolis Cars (RFPID-253813)
# ===========================================================================
ANNAPOLIS_GROUND_TRUTH = 2522716
# Of the four AUTO Make segments, three (Subaru/Hyundai/Volvo) are real,
# mapped, non-RFP-selectable catalog segments; Genesis Intender isn't in the
# catalog at all and so defaults to RFP-selectable (see module docstring).
# So the real, correct behavior for this real document is 3 distinct
# customs -- over the cap -- which the app allows with a warning, not a
# refusal.
ANNAPOLIS_EXPECTED_CUSTOM_COUNT = 3


def build_annapolis():
    return _build(ANNAPOLIS_PDF, ANNAPOLIS_GROUND_TRUTH, "Automotive", ANNAPOLIS_EXPECTED_CUSTOM_COUNT)


# ===========================================================================
# Scenario 3 -- Visit Hershey & Harrisburg (RFPID-260402)
# ===========================================================================
HERSHEY_GROUND_TRUTH = 310800336
# TRAVEL Family is the only mapped, non-RFP-selectable segment in this
# proposal (see module docstring for the unmapped ones) -- one distinct
# custom, under the cap, no warning expected.
HERSHEY_EXPECTED_CUSTOM_COUNT = 1


def build_hershey():
    return _build(HERSHEY_PDF, HERSHEY_GROUND_TRUTH, "Travel & Tourism", HERSHEY_EXPECTED_CUSTOM_COUNT)


# ===========================================================================
# Scenario 4 -- Wilmington University (RFPID-253956)
# ===========================================================================
WILMINGTON_GROUND_TRUTH = 332015108
# LIFESTYLE Online Education is the only mapped, non-RFP-selectable segment
# here -- one distinct custom, under the cap, no warning expected.
WILMINGTON_EXPECTED_CUSTOM_COUNT = 1


def build_wilmington():
    return _build(WILMINGTON_PDF, WILMINGTON_GROUND_TRUTH, "Education", WILMINGTON_EXPECTED_CUSTOM_COUNT)


# ===========================================================================
# Scenario 6 -- Lawn & Leisure (RFPID-265521)
#
# Added for the Hershey/2026-08-23 targeting-map fixes: the ONE real
# document on hand with a single group / single audience, load-bearing for
# the identity-collapse no-op guard (an audience-collapse pass has nothing
# to compare against with only one audience -- confirmed structurally, not
# just by construction, against a genuine real document) and for the
# multi-doc palette/export guards. "10 Mile Radius Zips" with NO bracketed
# origin -- the one real sample of that shape -- so this resolves via the
# document's own zip list directly (GEO_MODE_ZIPS fallback), not a
# geocode, unlike Annapolis's bracketed-origin radius groups.
# ===========================================================================
LAWN_LEISURE_PDF = REPO / "Premion Media Plan_RFPID-265521_Direct - No Agency_Lawn & Leisure_7-28-2026--ver0.pdf"
LAWN_LEISURE_GROUND_TRUTH = 509868
# Both DEMO Homeowner and HH Income 200K Plus are real, mapped, RFP-
# selectable catalog segments -- zero customs, no warning expected.
LAWN_LEISURE_EXPECTED_CUSTOM_COUNT = 0


def build_lawn_leisure():
    return _build(LAWN_LEISURE_PDF, LAWN_LEISURE_GROUND_TRUTH, "Home Improvement",
                  LAWN_LEISURE_EXPECTED_CUSTOM_COUNT)


def build_session_state(scenario):
    """A scenario dict (from build_annapolis/build_hershey/build_wilmington)
    -> the session_state dict a caller injects into a fresh AppTest before
    its first `.run()`, exactly like tests/test_draft_regression.py's
    build_deck does for a drafted state.
    """
    for group in scenario["groups"]:
        group.pop("_flight_label", None)
    option = app.new_plan_option(
        app.DEFAULT_OPTION_NAMES[0], scenario["rows"],
        driver=[app.DRIVER_IMPRESSIONS] * len(scenario["rows"]),
        breakout=app.BREAKOUT_FULL_FLIGHT)
    return {
        "include_avails_template": True,
        "targeting_groups": scenario["groups"],
        "avails_basis": app.AVAILS_BASIS_FLIGHT,
        "flight_start": scenario["flight_start"], "flight_end": scenario["flight_end"],
        "active_months": scenario["active_months"],
        "client_name": scenario["client_name"],
        # Distinct from client_name on purpose -- the plan slide's title
        # concatenates client name and proposal title, and a proposal_title
        # equal to the client name rendered as a visibly doubled "Annapolis
        # Cars Annapolis Cars", caught by actually looking at a rendered
        # slide. A rep would never type the client's own name here.
        "proposal_title": "CTV/OTT Strategy",
        "vertical_choice": scenario["vertical_choice"],
        "agency_involved": scenario["agency_involved"],
        "premion_streaming_tv": True,
        "plan_options": [option],
    }


SCENARIOS = {
    "annapolis": build_annapolis,
    "hershey": build_hershey,
    "wilmington": build_wilmington,
    "lawn_leisure": build_lawn_leisure,
}
