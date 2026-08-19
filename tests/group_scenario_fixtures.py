"""Shared fixture data for the three real-document scenarios in
targeting_groups_test_scenarios.md: Scenario 2 (Annapolis Cars), Scenario 3
(Visit Hershey & Harrisburg), Scenario 4 (Wilmington University).

The audiences, geos and avails figures below are transcribed verbatim from
the three real Premion avails PDFs sitting at the repo root (gitignored --
they carry real client pricing, same reason the Wide Orbit fixtures are
gitignored). The three GROUND_TRUTH totals are each PDF's own stated "Total
Impressions" figure -- external ground truth, never a sum this module
computes -- and every caller should assert against them directly rather
than against `sum(row avails)`, which only proves this module's own
transcription agrees with itself.

**Design choice, worth stating plainly:** groups are built as real
`targeting_groups` data (via `targeting_groups.new_group`, the exact shape
`app._add_segment_to_group` produces -- already proven correct by
tests/test_group_builder.py) rather than replayed through the ~80 AND/OR/
New-group button clicks it would take to build 8-12 groups one segment at a
time through AppTest. That mechanism is already covered; replaying it here
would spend a lot of runtime re-proving something this suite already knows.
What's NOT already covered, and what these scenarios exist to exercise, is
geography: every group's `geo_def`/`resolved_zips`/`resolved_markets` is
produced by calling `app.resolve_group_geography` -- the exact function the
real Resolve button calls, wired straight into `geo_resolver.py` -- never by
computing or guessing zip/radius/county membership here. Avails are written
directly onto each group's `avails_monthly`, matching every other
avails-editing test in this suite (`st.data_editor` can't be driven
directly, so a data_editor-backed edit is always injected as state).

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
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import app                                     # noqa: E402
import market_lookup                           # noqa: E402
import targeting_groups as tg                  # noqa: E402

# Registers the county->DMA table with geo_resolver directly (the same call
# app.install_market_lookup() makes from inside a running script) -- needed
# here because this module resolves geography OUTSIDE any Streamlit run.
market_lookup.install()

COL = app.AVAILS_COLUMN_MONTHLY


def _flight(start, end):
    all_months = app.month_list(start, end)
    label = app.format_flight_label(all_months, all_months)
    return all_months, label


def _radius_geo(center, miles):
    geo_def, zips, markets, notes, unresolved = app.resolve_group_geography(
        app.GEO_MODE_RADIUS, radius_centers_text=str(center), radius_miles=miles)
    return geo_def, zips, markets, notes, unresolved


def _markets_geo(market_keys):
    return app.resolve_group_geography(app.GEO_MODE_MARKETS, markets_picked=market_keys)


def _zips_geo(zips_text):
    return app.resolve_group_geography(app.GEO_MODE_ZIPS, zips_text=zips_text)


def _counties_geo(counties_text):
    return app.resolve_group_geography(app.GEO_MODE_COUNTIES, counties_text=counties_text)


def _make_group(terms, op, geo_def, resolved_zips, resolved_markets, avails_monthly, index):
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
    driven from the group's avails figure (this app has no "impressions =
    avails" automation; a rep always types the number, so a row's
    Impressions equal to its group's avails figure IS how a rep would
    replicate an already-booked document like these).
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


# ===========================================================================
# Scenario 2 -- Annapolis Cars (RFPID-253813)
# ===========================================================================
ANNAPOLIS_GROUND_TRUTH = 2522716
ANNAPOLIS_FLIGHT_START, ANNAPOLIS_FLIGHT_END = date(2026, 3, 1), date(2026, 3, 31)
ANNAPOLIS_CENTER = "21401"
# Transcribed verbatim from RFPID-253813's own "Zip Codes" tables.
ANNAPOLIS_ZIPS_10MI = {"20765", "20776", "21012", "21032", "21035", "21037", "21054",
                       "21056", "21108", "21114", "21122", "21140", "21146", "21401",
                       "21402", "21403", "21405", "21409"}
ANNAPOLIS_ZIPS_5MI = {"21012", "21032", "21140", "21401", "21402", "21403", "21405"}
# (audience term, radius miles, avails) -- in the document's own row order.
ANNAPOLIS_ROWS = [
    ("AUTO Make Subaru", 10, 602647),
    ("AUTO Make Subaru", 5, 165687),
    ("AUTO Make Hyundai", 10, 669089),
    ("AUTO Make Hyundai", 5, 184798),
    ("AUTO Make Volvo", 10, 279310),
    ("AUTO Make Volvo", 5, 84637),
    ("AUTO Make Genesis Intender", 10, 418144),
    ("AUTO Make Genesis Intender", 5, 118404),
]
assert sum(r[2] for r in ANNAPOLIS_ROWS) == ANNAPOLIS_GROUND_TRUTH, \
    "transcription of RFPID-253813 no longer sums to its own stated total"
# Of the four AUTO Make segments, three (Subaru/Hyundai/Volvo) are real,
# mapped, non-RFP-selectable catalog segments; Genesis Intender isn't in the
# catalog at all and so defaults to RFP-selectable (see module docstring).
# So the real, correct behavior for this real document is 3 distinct
# customs -- over the cap -- which the app allows with a warning, not a
# refusal.
ANNAPOLIS_EXPECTED_CUSTOM_COUNT = 3


def build_annapolis():
    all_months, flight_label = _flight(ANNAPOLIS_FLIGHT_START, ANNAPOLIS_FLIGHT_END)
    geo_10mi = _radius_geo(ANNAPOLIS_CENTER, 10)
    geo_5mi = _radius_geo(ANNAPOLIS_CENTER, 5)
    geo_by_miles = {10: geo_10mi, 5: geo_5mi}

    groups, rows, geo_notes = [], [], []
    for i, (term, miles, avails) in enumerate(ANNAPOLIS_ROWS):
        geo_def, zips, markets, notes, unresolved = geo_by_miles[miles]
        geo_notes.append((term, miles, notes, unresolved))
        group = _make_group([term], None, geo_def, zips, markets, avails, i)
        group["_flight_label"] = flight_label
        groups.append(group)
        rows.append(_plan_row(group, "premion_streaming_tv", avails, markup=1.15))

    return {
        "name": "Annapolis Cars", "client_name": "Annapolis Cars",
        "flight_start": ANNAPOLIS_FLIGHT_START, "flight_end": ANNAPOLIS_FLIGHT_END,
        "active_months": all_months, "flight_label": flight_label,
        "n_months": len(all_months), "agency_involved": True,
        "vertical_choice": "Automotive",
        "groups": groups, "geo_notes": geo_notes,
        "rows": rows, "ground_truth": ANNAPOLIS_GROUND_TRUTH,
        "expected_custom_count": ANNAPOLIS_EXPECTED_CUSTOM_COUNT,
    }


# ===========================================================================
# Scenario 3 -- Visit Hershey & Harrisburg (RFPID-260402)
# ===========================================================================
HERSHEY_GROUND_TRUTH = 310800336
HERSHEY_FLIGHT_START, HERSHEY_FLIGHT_END = date(2026, 5, 25), date(2026, 7, 5)

HERSHEY_MARKETS = {
    "Philadelphia": "philadelphia", "Baltimore": "baltimore",
    "Washington, D.C.": "washington_hagerstown", "New York": "new_york",
}
# Verbatim from RFPID-260402's zip-code tables.
HERSHEY_PHILLY_ZIPS = (
    "19601,19602,19604,19605,19606,19607,19608,19609,19508,19510,19512,19518,19522,19526,19533,19540,"
    "18101,18102,18103,18104,18106,18109,18015,18031,18034,18052,18062,18069,18017,18018,18020,18042,"
    "18045,18064,18067,18072,18083,18091,18201,18210,18229,18235,18240,18244,18255,08023,08038,08067,"
    "08069,08070,08072,08318"
)
HERSHEY_NY_ZIPS = (
    "07416,07418,07419,07422,07428,07439,07460,07461,07821,07822,07826,07848,07823,07825,07832,07838,"
    "07840,07844,07863,07865,07882,08525,08551,08559,08801,08822,08825,08826,08829,08833,08848,08867,"
    "08887,08889"
)
# (audience key "A"/"B", geo kind, geo value, avails) -- document row order.
HERSHEY_ROWS = [
    ("A", "market", "Philadelphia", 38426724),
    ("A", "market", "Baltimore", 14635026),
    ("A", "market", "Washington, D.C.", 25264764),
    ("A", "market", "New York", 66349836),
    ("A", "zips_philly", None, 5829978),
    ("A", "zips_ny", None, 1011108),
    ("B", "market", "Philadelphia", 40732230),
    ("B", "market", "Baltimore", 15489096),
    ("B", "market", "Washington, D.C.", 26736864),
    ("B", "market", "New York", 70003248),
    ("B", "zips_philly", None, 5397126),
    ("B", "zips_ny", None, 924336),
]
assert sum(r[3] for r in HERSHEY_ROWS) == HERSHEY_GROUND_TRUTH, \
    "transcription of RFPID-260402 no longer sums to its own stated total"
HERSHEY_AUDIENCE_TERMS = {
    "A": (["AFIRST Travel Buffs and Sightseers", "DEMO Age A21-44"], "AND"),
    "B": (["TRAVEL Family", "DEMO Age A25 Plus"], "AND"),
}
# TRAVEL Family is the only mapped, non-RFP-selectable segment in this
# proposal (see module docstring for the unmapped ones) -- one distinct
# custom, under the cap, no warning expected.
HERSHEY_EXPECTED_CUSTOM_COUNT = 1


def build_hershey():
    all_months, flight_label = _flight(HERSHEY_FLIGHT_START, HERSHEY_FLIGHT_END)
    zips_geo_philly = _zips_geo(HERSHEY_PHILLY_ZIPS)
    zips_geo_ny = _zips_geo(HERSHEY_NY_ZIPS)
    market_geo_cache = {}

    n_months = len(all_months)
    groups, rows, geo_notes = [], [], []
    for i, (aud_key, kind, geo_value, avails) in enumerate(HERSHEY_ROWS):
        terms, op = HERSHEY_AUDIENCE_TERMS[aud_key]
        if kind == "market":
            market_key = HERSHEY_MARKETS[geo_value]
            if market_key not in market_geo_cache:
                market_geo_cache[market_key] = _markets_geo([market_key])
            geo_def, zips, markets, notes, unresolved = market_geo_cache[market_key]
        elif kind == "zips_philly":
            geo_def, zips, markets, notes, unresolved = zips_geo_philly
        else:
            geo_def, zips, markets, notes, unresolved = zips_geo_ny
        geo_notes.append((aud_key, kind, geo_value, notes, unresolved))
        # The GROUP stores a true monthly figure (the "Max Monthly Avails"
        # column's own convention -- see restore_untouched_avails/
        # avails_from_display in app.py), rounded from the document's
        # full-flight number the same way a rep typing that number under the
        # Full-flight basis toggle would produce it. The PLAN ROW keeps the
        # document's exact full-flight figure, unrounded -- Full-Flight
        # breakout reads a row's Impressions as-is, with no division at all,
        # which is why the deck's own Full Flight Total ties to the
        # published ground truth exactly while the avails table's round-trip
        # carries a small, bounded rounding delta (see build_session_state's
        # caller for the tolerance and why it exists).
        group = _make_group(terms, op, geo_def, zips, markets, round(avails / n_months), i)
        group["_flight_label"] = flight_label
        groups.append(group)
        rows.append(_plan_row(group, "premion_streaming_tv", avails, markup=1.0))

    return {
        "name": "Visit Hershey & Harrisburg", "client_name": "Visit Hershey & Harrisburg",
        "flight_start": HERSHEY_FLIGHT_START, "flight_end": HERSHEY_FLIGHT_END,
        "active_months": all_months, "flight_label": flight_label,
        "n_months": len(all_months), "agency_involved": False,
        "vertical_choice": "Travel & Tourism",
        "groups": groups, "geo_notes": geo_notes,
        "rows": rows, "ground_truth": HERSHEY_GROUND_TRUTH,
        "expected_custom_count": HERSHEY_EXPECTED_CUSTOM_COUNT,
    }


# ===========================================================================
# Scenario 4 -- Wilmington University (RFPID-253956)
# ===========================================================================
WILMINGTON_GROUND_TRUTH = 332015108
WILMINGTON_FLIGHT_START, WILMINGTON_FLIGHT_END = date(2026, 7, 1), date(2027, 6, 30)
WILMINGTON_COUNTIES = ("NEW CASTLE DE; KENT DE; SUSSEX DE; CHESTER PA; PHILADELPHIA PA; "
                       "DELAWARE PA; BUCKS PA; SALEM NJ; GLOUCESTER NJ; CAMDEN NJ")
# (audience terms, op, avails full-flight) -- document row order.
WILMINGTON_ROWS = [
    (["LIFESTAGE College Planning Parents"], None, 54005249),
    (["LIFESTAGE Prospective College Students"], None, 77173481),
    (["LIFESTYLE Online Education"], None, 73412846),
    (["DEMO Career Employed", "LIFESTAGE Education Services"], "AND", 52725960),
    (["LIFESTAGE Higher Education Intender"], None, 74697572),
]
assert sum(r[2] for r in WILMINGTON_ROWS) == WILMINGTON_GROUND_TRUTH, \
    "transcription of RFPID-253956 no longer sums to its own stated total"
# LIFESTYLE Online Education is the only mapped, non-RFP-selectable segment
# here -- one distinct custom, under the cap, no warning expected.
WILMINGTON_EXPECTED_CUSTOM_COUNT = 1


def build_wilmington():
    all_months, flight_label = _flight(WILMINGTON_FLIGHT_START, WILMINGTON_FLIGHT_END)
    geo_def, zips, markets, notes, unresolved = _counties_geo(WILMINGTON_COUNTIES)

    n_months = len(all_months)
    groups, rows = [], []
    for i, (terms, op, avails) in enumerate(WILMINGTON_ROWS):
        # See build_hershey's comment: the group stores a rounded true
        # monthly figure, the plan row keeps the document's exact
        # full-flight number.
        group = _make_group(terms, op, geo_def, zips, markets, round(avails / n_months), i)
        group["_flight_label"] = flight_label
        groups.append(group)
        rows.append(_plan_row(group, "premion_streaming_tv", avails, markup=1.0))

    return {
        "name": "Wilmington University", "client_name": "Wilmington University",
        "flight_start": WILMINGTON_FLIGHT_START, "flight_end": WILMINGTON_FLIGHT_END,
        "active_months": all_months, "flight_label": flight_label,
        "n_months": len(all_months), "agency_involved": False,
        "vertical_choice": "Education",
        "groups": groups, "geo_notes": [("all 5 groups", "counties", WILMINGTON_COUNTIES, notes, unresolved)],
        "rows": rows, "ground_truth": WILMINGTON_GROUND_TRUTH,
        "expected_custom_count": WILMINGTON_EXPECTED_CUSTOM_COUNT,
    }


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
}
