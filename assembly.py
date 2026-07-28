"""
assembly.py -- standalone deck assembly engine (build spec sections 2 & 7).

Proves the core mechanic only: open the master deck, resolve a hardcoded
selection dictionary into a set of slide numbers to keep, delete everything
else by editing <p:sldIdLst> directly, and save a valid .pptx.

No Streamlit, no Supabase, no personalization fill yet -- this only proves
selection -> deletion -> valid file open-able in PowerPoint.
"""

from pptx import Presentation

MASTER_DECK_PATH = "TEGNA MASTER DECK2.pptx"
OUTPUT_PATH = "test.pptx"

# ---------------------------------------------------------------------------
# Slide map: 1-indexed master-deck slide number -> condition_key
# Transcribed from build spec section 4. Slides 45-46, 49-50 (duplicate
# leftovers) and 120-123 (old static proposal examples, to be replaced by a
# generated proposal slide per section 7 item 3) are intentionally left out
# of this map -- they are dropped unconditionally, matching the section 9
# phase-1 cleanup notes.
# ---------------------------------------------------------------------------
SLIDE_MAP = {}


def _assign(start, end, key):
    for n in range(start, end + 1):
        SLIDE_MAP[n] = key


_assign(1, 4, "always")            # 1 = client title slide (template TBD, section 7 item 1)
_assign(5, 8, "full_deck")
_assign(9, 9, "spanish")
_assign(10, 11, "full_deck")
_assign(12, 12, "vertical:healthcare")
_assign(13, 13, "first_party")
_assign(14, 14, "streaming_retargeting")
_assign(15, 20, "full_deck")
_assign(21, 21, "linear_reach_ext")
_assign(22, 23, "brand_lift")
_assign(24, 24, "sales_attribution")
_assign(25, 25, "case_study:retail")
_assign(26, 26, "vertical:travel")
_assign(27, 28, "full_deck")
_assign(29, 29, "any_vertical")
_assign(30, 32, "vertical:education")
_assign(33, 35, "vertical:healthcare")
_assign(36, 37, "vertical:retail")
_assign(38, 40, "vertical:travel")
_assign(41, 42, "vertical:home_improvement")
_assign(43, 44, "vertical:banking")
# 45-46 duplicate healthcare -- dropped, not mapped
_assign(47, 48, "vertical:entertainment")
# 49-50 duplicate education -- dropped, not mapped
_assign(51, 52, "vertical:dining_qsr")
_assign(53, 55, "vertical:auto")
_assign(56, 59, "tegna_positioning")
_assign(60, 63, "am")
SLIDE_MAP[64] = "am:retargeting"
SLIDE_MAP[65] = "am:audience"
SLIDE_MAP[66] = "am:geofencing"
_assign(67, 68, "am")
_assign(69, 73, "sports")
# 74-94: per-sport viewership slides. Exact viewership<->package pairing is
# flagged as an open item in spec section 10; treated coarsely here as one
# block gated on "any sport selected" until that pairing is finalized.
_assign(74, 94, "sports_viewership")
# 95-114: per-sport package slides, individually keyed from slide titles.
SLIDE_MAP[95] = "sport:nhl_reg"
SLIDE_MAP[96] = "sport:nhl_playoffs"
SLIDE_MAP[97] = "sport:wnba_playoffs"
SLIDE_MAP[98] = "sport:wnba_reg"
SLIDE_MAP[99] = "sport:ncaa_basketball"
SLIDE_MAP[100] = "sport:nba_playoffs"
SLIDE_MAP[101] = "sport:nba_reg"
SLIDE_MAP[102] = "sport:soccer_pro"
SLIDE_MAP[103] = "sport:nfl_playoffs"
SLIDE_MAP[104] = "sport:nfl_home_team"
SLIDE_MAP[105] = "sport:nfl_reg"
SLIDE_MAP[106] = "sport:ncaaf_playoffs"
SLIDE_MAP[107] = "sport:ncaaf_home_team"
SLIDE_MAP[108] = "sport:ncaaf_conference"
SLIDE_MAP[109] = "sport:ncaaf_reg"
SLIDE_MAP[110] = "sport:prestige_sports"
SLIDE_MAP[111] = "sport:golf_pga"
SLIDE_MAP[112] = "sport:all_live_sports"
SLIDE_MAP[113] = "sport:mlb_playoffs"
SLIDE_MAP[114] = "sport:mlb_reg"
_assign(115, 116, "total_tv")
SLIDE_MAP[117] = "total_tv:dc"
SLIDE_MAP[118] = "total_tv:harrisburg"
SLIDE_MAP[119] = "total_tv"
# 120-123 old static proposal section -- dropped, replaced by a generated
# proposal slide per section 7 item 3 (not built yet)

# ---------------------------------------------------------------------------
# Hardcoded selection dictionary, standing in for a form submission (section 5)
# ---------------------------------------------------------------------------
SELECTIONS = {
    "preset": "full_proposal",       # 'full_proposal' | 'quick_pitch'
    "market": "DC",                  # 'DC' | 'Harrisburg'
    "vertical": "healthcare",
    "agency_involved": False,
    "spanish_campaign": False,
    "tegna_positioning": True,
    "products": {
        "streaming_retargeting": True,
        "audience_marketplace": {
            "enabled": True,
            "audience_targeting": False,
            "geofencing": True,
            "site_retargeting_display": False,
            "site_retargeting_preroll": False,
        },
        "live_sports": {
            "enabled": True,
            "sports": ["nfl_playoffs", "nba_reg"],
        },
        "total_tv": True,
    },
    "targeting_attribution": {
        "first_party_data": True,
        "linear_reach_extension": True,
        "sales_attribution": True,
        "brand_lift": True,
    },
}


def resolve_active_keys(selections):
    """Turn the selection dictionary into the set of active condition_keys."""
    active = {"always"}

    if selections["preset"] == "full_proposal":
        active.add("full_deck")

    vertical = selections.get("vertical")
    if vertical and vertical != "none":
        active.add(f"vertical:{vertical}")
        active.add("any_vertical")

    if selections["spanish_campaign"]:
        active.add("spanish")

    if selections["tegna_positioning"]:
        active.add("tegna_positioning")

    products = selections["products"]

    if products.get("streaming_retargeting"):
        active.add("streaming_retargeting")

    am = products.get("audience_marketplace", {})
    if am.get("enabled"):
        active.add("am")
        if am.get("audience_targeting"):
            active.add("am:audience")
        if am.get("geofencing"):
            active.add("am:geofencing")
        if am.get("site_retargeting_display") or am.get("site_retargeting_preroll"):
            active.add("am:retargeting")

    sports = products.get("live_sports", {})
    if sports.get("enabled") and sports.get("sports"):
        active.add("sports")
        active.add("sports_viewership")
        for sport_key in sports["sports"]:
            active.add(f"sport:{sport_key}")

    if products.get("total_tv"):
        active.add("total_tv")
        market_key = "dc" if selections["market"] == "DC" else "harrisburg"
        active.add(f"total_tv:{market_key}")

    targeting = selections["targeting_attribution"]
    if targeting.get("first_party_data"):
        active.add("first_party")
    if targeting.get("linear_reach_extension") and products.get("total_tv"):
        active.add("linear_reach_ext")
    if targeting.get("sales_attribution"):
        active.add("sales_attribution")
    if targeting.get("brand_lift"):
        active.add("brand_lift")

    return active


def slides_to_keep(slide_map, active_keys):
    return sorted(n for n, key in slide_map.items() if key in active_keys)


def delete_slide(prs, slide_index):
    """Remove one slide (0-indexed) from the presentation.

    Editing <p:sldIdLst> directly, per spec section 2 -- python-pptx has no
    built-in slide-deletion API. Dropping the relationship as well as the
    <p:sldId> entry means the slide part becomes unreachable from the
    presentation part's relationship graph; python-pptx only serializes
    parts it can still reach on save, so the orphaned slide (and anything
    only it referenced, e.g. its notes slide or unique images) is cleaned up
    automatically on save -- no manual part surgery needed.
    """
    sldIdLst = prs.slides._sldIdLst
    sldId = list(sldIdLst)[slide_index]
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)


def assemble(master_path, output_path, selections):
    prs = Presentation(master_path)
    original_count = len(prs.slides._sldIdLst)

    active_keys = resolve_active_keys(selections)
    keep_numbers = set(slides_to_keep(SLIDE_MAP, active_keys))

    # Delete back-to-front so earlier indices don't shift under us.
    for slide_number in range(original_count, 0, -1):
        if slide_number not in keep_numbers:
            delete_slide(prs, slide_number - 1)

    prs.save(output_path)
    return original_count, len(keep_numbers)


if __name__ == "__main__":
    original_count, kept_count = assemble(MASTER_DECK_PATH, OUTPUT_PATH, SELECTIONS)
    print(f"Master deck: {original_count} slides")
    print(f"Kept: {kept_count} slides -> saved to {OUTPUT_PATH}")
