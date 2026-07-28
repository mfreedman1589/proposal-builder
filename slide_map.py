"""
slide_map.py -- scan a master deck and derive slide_number -> condition_key.

Per master_v1_1_changelog.md: "Don't hand-patch offsets -- regenerate
SLIDE_MAP by scanning v1.1 (extract per-slide text, match section anchors)."
This is also called out as reusable infrastructure for the Vault versioning
flow (build spec section 8), so it's written as a standalone, deck-agnostic
utility rather than a one-off script: point it at any deck built on the same
content anchors and it re-derives the map without touching slide numbers.

Approach: extract concatenated text per slide, then run each slide's text
through an ordered set of anchor rules (most specific first). A few rules
mark section-divider slides and update a "current section" carry-forward
default; that default is what genuinely blank/undifferentiated slides fall
back to. Every other slide is resolved by a directly-matched anchor.
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)


def extract_slide_text(slide):
    """Concatenate all run/cell text on a slide into one fingerprint string."""
    parts = []
    for shape in iter_all_shapes(slide.shapes):
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t:
                parts.append(t)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    t = cell.text_frame.text.strip()
                    if t:
                        parts.append(t)
    return " || ".join(parts)


def _contains_all(text, *substrings):
    return all(s in text for s in substrings)


def _classify_sport_package(text):
    """Disambiguate which sport a 'PREMION | SPORTS PACKAGES' slide covers.

    Ordered so that more specific families (WNBA before NBA, Home Team
    before Regular/Playoffs) are checked first -- several of these strings
    are substrings of each other (e.g. "WNBA" contains "NBA").
    """
    rules = [
        (("NHL", "Regular"), "sport:nhl_reg"),
        (("NHL", "Playoffs"), "sport:nhl_playoffs"),
        (("WNBA", "Playoffs"), "sport:wnba_playoffs"),
        (("WNBA", "Regular"), "sport:wnba_reg"),
        (("NCAA", "Basketball"), "sport:ncaa_basketball"),
        (("NBA", "Playoffs"), "sport:nba_playoffs"),
        (("NBA", "Regular"), "sport:nba_reg"),
        # "Home Team Game Inventory Only" is the actual home-team-package
        # marker; bare "Home Team" also appears in regular-season slides'
        # coverage bullet ("League-Wide & Home Team Games"), so that alone
        # would misclassify them.
        (("Home Team Game Inventory", "NFL"), "sport:nfl_home_team"),
        (("NFL", "Playoffs"), "sport:nfl_playoffs"),
        (("NFL", "Regular"), "sport:nfl_reg"),
        (("Home Team Game Inventory", "NCAA"), "sport:ncaaf_home_team"),
        (("NCAA Football", "Playoffs"), "sport:ncaaf_playoffs"),
        (("NCAA Football", "Conferences"), "sport:ncaaf_conference"),
        (("NCAA", "Football", "Regular"), "sport:ncaaf_reg"),
        (("MLB", "Playoffs"), "sport:mlb_playoffs"),
        (("MLB", "Regular"), "sport:mlb_reg"),
        (("Professional", "Soccer"), "sport:soccer_pro"),
        (("Professional", "Golf"), "sport:golf_pga"),
        (("Prestige", "Sports"), "sport:prestige_sports"),
        (("All Live", "Sports"), "sport:all_live_sports"),
    ]
    for needles, key in rules:
        if _contains_all(text, *needles):
            return key
    return None


VERTICAL_ANCHORS = [
    ("PREMION + RETAIL", "vertical:retail"),
    ("PREMION + TRAVEL", "vertical:travel"),
    ("PREMION + HOME IMPROVEMENT", "vertical:home_improvement"),
    ("PREMION + BANKING", "vertical:banking"),
    ("PREMION + ENTERTAINMENT", "vertical:entertainment"),
    ("PREMION + EDUCATION", "vertical:education"),
    ("PREMION + CASUAL DINING", "vertical:dining_qsr"),
    ("PREMION + AUTOMOTIVE", "vertical:auto"),
    ("POLK AUDIENCES", "vertical:auto"),
    ("POLK SIGNALS", "vertical:auto"),
]

# Exact-text section dividers: (literal text, this slide's own key, the
# carry-forward default to adopt afterward). The three "transitional" content
# dividers (Premium Content / Precision Targeting / Attribution+Measurement)
# get their own key so they can be dropped from every preset, while slides
# that follow them still fall back to "full_deck" as before.
SECTION_DIVIDERS = [
    ("Content || Premium || © PREMION 2024", "transitional_divider", "full_deck"),
    ("Targeting || Precision || © PREMION 2024", "transitional_divider", "full_deck"),
    ("Measurement || Attribution +", "transitional_divider", "full_deck"),
    ("Specialties || Vertical", "any_vertical", "any_vertical"),
    ("Media Group || TEGNA", "tegna_positioning", "tegna_positioning"),
    ("Marketplace || Audience", "am", "am"),
    ("Audience Marketplace", "am", "am"),
    ("Live Sports", "sports", "sports"),
    ("Total TV", "total_tv", "total_tv"),
    ("Proposal Slides", "proposal_divider", "proposal_divider"),
]


def _classify_sport_viewership(upper_text):
    """Disambiguate which sport an 'Ad-Supported Streaming TV Viewer' slide
    covers, mirroring _classify_sport_package. Order matters: WNBA before
    NBA (substring), NCAAF before bare NCAA, etc. A handful of viewership
    slides (bare "Golf Viewers", "Motorsports Viewers", "World Cup Viewers",
    "March Madness") have no 1:1 package counterpart in this deck version
    and are intentionally left unmatched -- they're never selectable so
    never included.
    """
    rules = [
        (("WNBA", "PLAYOFF"), "sport_viewership:wnba_playoffs"),
        (("WNBA",), "sport_viewership:wnba_reg"),
        (("NBA", "PLAYOFF"), "sport_viewership:nba_playoffs"),
        (("NBA",), "sport_viewership:nba_reg"),
        (("NFL", "PLAYOFF"), "sport_viewership:nfl_playoffs"),
        (("NFL",), "sport_viewership:nfl_reg"),
        (("NHL", "PLAYOFF"), "sport_viewership:nhl_playoffs"),
        (("NHL",), "sport_viewership:nhl_reg"),
        (("MLB", "PLAYOFF"), "sport_viewership:mlb_playoffs"),
        (("MLB",), "sport_viewership:mlb_reg"),
        # "CONF" alone also matches "confidential" in every slide's MRI
        # footnote boilerplate; "CONF." (with the period, as in "Conf.
        # Viewers") does not.
        (("NCAAF", "CONF."), "sport_viewership:ncaaf_conference"),
        (("NCAAF", "PLAYOFF"), "sport_viewership:ncaaf_playoffs"),
        (("NCAAF",), "sport_viewership:ncaaf_reg"),
        (("NCAA BASKETBALL",), "sport_viewership:ncaa_basketball"),
        (("PGA",), "sport_viewership:golf_pga"),
        (("PRESTIGE",), "sport_viewership:prestige_sports"),
        # "SOCCER VIEWERS" (the heading, two words together) rather than
        # bare "SOCCER" -- the World Cup slide's own footnote lists
        # "Soccer-World Cup" among included leagues, which would otherwise
        # misclassify it as the Soccer slide.
        (("SOCCER VIEWERS",), "sport_viewership:soccer_pro"),
    ]
    for needles, key in rules:
        if _contains_all(upper_text, *needles):
            return key
    return None


def _classify_slide_raw(text, current_default):
    """Return (condition_key, new_default) for one slide's text."""
    upper = text.upper()

    # --- Priority 1: template token slides -----------------------------
    if "{{PROPOSAL_TITLE}}" in text:
        return "client_title", current_default
    if "{{GOALS_BULLETS}}" in text:
        return "campaign_specs", current_default
    if "{{AVAILS}}" in text and "{{VERTICAL}}" in text:
        return "targeting_avails_template", current_default
    if "{{PLAN_TITLE}}" in text:
        return "proposal_template", current_default

    # --- Priority 2: distinctive single-slide markers -------------------
    if "BILINGUAL-SPANISH" in upper:
        return "spanish", current_default
    if "CROSSIX" in upper or "PREMION + HEALTHCARE" in upper:
        return "vertical:healthcare", current_default
    if "1ST PARTY DATA TARGETING" in upper:
        return "first_party", current_default
    if "STREAMING TV RETARGETING" in upper:
        return "streaming_retargeting", current_default
    if "Dynamic Video Ad" in text:
        return "dynamic_creative", current_default
    if "Linear TV + PREMION" in text:
        return "linear_reach_ext", current_default
    if "PREMION BRAND LIFT STUDY" in upper or "Brand Lift Campaign Results" in text:
        return "brand_lift", current_default
    if "CRM data" in text:
        return "sales_attribution", current_default
    if "Regional Furniture Store" in text:
        return "case_study:retail", current_default
    if "ADVANCED ATTRIBUTION INSIGHTS" in upper:
        return "vertical:travel", current_default

    # --- Priority 2b: generic per-vertical anchor -----------------------
    for needle, key in VERTICAL_ANCHORS:
        if needle in upper:
            return key, current_default

    # --- Priority 3: section dividers + subsection markers --------------
    stripped = text.strip()
    for divider_text, key, default_key in SECTION_DIVIDERS:
        if stripped == divider_text:
            return key, default_key

    if "puts your brand alongside trusted journalism" in text:
        return "tegna_positioning", "tegna_positioning"
    if "product solution suite was designed" in text:
        return "tegna_positioning", "tegna_positioning"
    if "AWARENESS || ENGAGEMENT || CONVERSION" in text:
        return "tegna_positioning", "tegna_positioning"
    if "utilizes precision audience targeting, geofencing and retargeting" in text:
        return "am", "am"
    if "AUDIENCE MARKETPLACE ///" in text:
        return "am", "am"
    if "HOW WE DO IT" in upper and "RETARGETING" in upper:
        return "am:retargeting", current_default
    if "HOW WE DO IT" in upper and "AUDIENCE TARGETING" in upper:
        return "am:audience", current_default
    if "HOW WE DO IT" in upper and "GEOFENCING" in upper:
        return "am:geofencing", current_default
    if "Formats & Screens" in text:
        return "am", current_default
    if "HOW WE DO IT" in upper and "REPORTING" in upper:
        return "am", current_default
    if "Reach Sports Fans" in text:
        return "sports", "sports"
    if "SPORTS CONTENT PACKAGES" in upper:
        return "sports", "sports"

    # --- Priority 4: sports viewership / package blocks ------------------
    if "LIVE SPORTS VIEWERSHIP" in upper:
        if "Live Sports Viewers" in text:
            # the one generic overview slide, not tied to a specific sport
            return "sports_viewership_intro", current_default
        sv_key = _classify_sport_viewership(upper)
        return (sv_key or "sports_viewership_unmapped"), current_default
    if "SPORTS PACKAGES" in upper:
        sport_key = _classify_sport_package(text)
        if sport_key:
            return sport_key, current_default

    # --- Priority 5: Total TV specifics ----------------------------------
    if "TV SCHEDULE PLACEHOLDER" in upper:
        return "total_tv", "total_tv"
    if "Total TV is Still King" in text:
        return "total_tv", "total_tv"
    if "WUSA + PREMION" in text:
        return "total_tv:dc", current_default
    if "WPMT + PREMION" in text:
        return "total_tv:harrisburg", current_default

    # --- Fallback: carry-forward ambient section -------------------------
    return current_default, current_default


def classify_slide(text, current_default):
    """Wraps _classify_slide_raw, appending ':targeting' to any 'vertical:X'
    result whose slide text reads as the vertical's own "PRECISION
    TARGETING" page (as opposed to its stats/attribution/case-study
    slides). This lets the caller swap that one slide out in favor of the
    personalized targeting_avails_template, per the vertical specialty
    targeting/avails mutual-exclusion rule.
    """
    key, new_default = _classify_slide_raw(text, current_default)
    if key and key.startswith("vertical:") and ":targeting" not in key:
        upper = text.upper()
        if "PRECISION" in upper and "TARGETING" in upper:
            key = key + ":targeting"
    return key, new_default


def build_slide_map_from_prs(prs):
    """Scan an already-open Presentation and return {slide_number: condition_key}.

    Slides that resolve to None are printed as warnings -- they need a new
    anchor rule added above (should not happen for a deck built on the same
    content conventions as v1.1).
    """
    result = {}
    current_default = "always"
    unresolved = []

    for i, slide in enumerate(prs.slides, start=1):
        text = extract_slide_text(slide)
        key, current_default = classify_slide(text, current_default)
        result[i] = key
        if key is None:
            unresolved.append((i, text[:80]))

    if unresolved:
        print(f"WARNING: {len(unresolved)} slide(s) could not be classified:")
        for n, snippet in unresolved:
            print(f"  slide {n}: {snippet!r}")

    return result


def build_slide_map(pptx_path):
    """Scan pptx_path (opening it fresh) and return {slide_number: condition_key}."""
    return build_slide_map_from_prs(Presentation(pptx_path))


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "TEGNA_MASTER_DECK_v1_1.pptx"
    result = build_slide_map(path)
    for n in sorted(result):
        print(n, "->", result[n])
