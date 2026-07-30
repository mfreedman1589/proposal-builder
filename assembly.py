"""
assembly.py -- standalone deck assembly engine (build spec sections 2 & 7),
updated for master deck v1.1 (see master_v1_1_changelog.md).

Deletes unselected slides, then fills the retained template slides using the
{{TOKEN}} convention: find runs whose text contains a {{TOKEN}}, replace via
run.text assignment (never text_frame.text, which collapses formatting).
Multi-item content (bullets, table rows) is filled by deep-copying the
tokenized paragraph/row element and inserting clones as siblings.

No Streamlit, no Supabase yet -- still just proving the assembly + fill
mechanics end to end on the real deck.
"""

import copy

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.util import Emu, Pt

import slide_map

# Relationship-reference attributes (r:embed, r:id, r:link) live in this
# namespace. They're part-local, so copied shape XML has to be rewritten.
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

MASTER_DECK_PATH = "TEGNA_MASTER_DECK_v1_1.pptx"
OUTPUT_PATH = "test.pptx"

# "Standard" preset (default): an exact, hand-picked slide-number allowlist
# from the current master deck, plus whatever add-ons (products/vertical/
# sports/etc.) are separately selected. Independent of condition_key
# resolution -- these numbers are forced in regardless of preset-driven key
# exclusions (e.g. slide 7 is 'full_deck', which "standard" otherwise drops).
STANDARD_CORE_SLIDES = frozenset({1, 2, 3, 4, 5, 7, 13, 19, 21, 22, 118})


# ---------------------------------------------------------------------------
# Hardcoded selection dictionary, standing in for a form submission (section 5)
# ---------------------------------------------------------------------------
SELECTIONS = {
    "preset": "extended",       # 'quick_pitch' | 'standard' | 'extended'
    "market": "DC",                  # 'DC' | 'Harrisburg'
    "vertical": "healthcare",
    "agency_involved": False,
    "spanish_campaign": False,
    "tegna_positioning": True,
    "include_avails_template": True,
    "products": {
        "streaming_retargeting": {"enabled": True, "display": True, "preroll": False},
        "audience_marketplace": {
            "enabled": True,
            "audience_targeting_display": False,
            "audience_targeting_preroll": False,
            "geofencing_display": True,
            "geofencing_preroll": False,
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


# ---------------------------------------------------------------------------
# Personalization payload -- what the Claude-written copy + form fields would
# produce (spec section 6/7). Hardcoded here since there's no Streamlit/Claude
# wiring yet.
# ---------------------------------------------------------------------------
FILL_DATA = {
    "client_name": "Acme Test Co",
    "proposal_title": "CTV Strategy",
    "logo_path": "placeholder_logo.png",
    "vertical_display": "Healthcare",
    "campaign_specs": {
        "GOALS_BULLETS": [
            "Grow new-patient volume across the DC market",
            "Build measurable brand lift ahead of open enrollment",
        ],
        "AUDIENCE_BULLETS": [
            "Health & Fitness intenders",
            "Crossix condition-specific segments",
        ],
        "GEOGRAPHY_BULLETS": ["Washington, DC DMA"],
        "BUDGET_BULLETS": ["$85,000 monthly gross"],
        "PLACEMENTS_BULLETS": [
            "15s/30s CTV creative",
            "Dynamic Video Ad refresh monthly",
        ],
        "TIMING_BULLETS": ["Flight: 9/1/26 - 11/30/26"],
    },
    "avails": {
        "rows": [
            {"audience": "Health & Fitness Intenders", "geo": "Washington, DC DMA", "avails": "1,250,000"},
            {"audience": "Health & Fitness Intenders", "geo": "Baltimore DMA", "avails": "480,000"},
        ],
        "total_avails": "1,730,000",
    },
    "media_plan": {
        "plan_title": "CTV Strategy",
        "rows": [
            {"tactic": "Premion Streaming TV", "flight": "9/1-11/30", "geo": "DC DMA",
             "targeting": "Health & Fitness Intenders", "impressions": "1,000,000", "cost": "$35,000"},
            {"tactic": "Streaming TV Retargeting", "flight": "9/1-11/30", "geo": "DC DMA",
             "targeting": "Site Visitors", "impressions": "250,000", "cost": "$8,750"},
        ],
        "totals_label": "Monthly Totals",
        "total_impressions": "1,250,000",
        "total_cost": "$43,750",
        "included_list": [
            "Dedicated Account Management Team",
            "Monthly Reporting Calls & Optimizations",
            "Dashboard Access",
            "Web Attribution (Pixel Required)",
        ],
    },
}


def resolve_active_keys(selections):
    """Turn the selection dictionary into the set of active condition_keys.

    Note: 'transitional_divider' (Premium Content / Precision Targeting /
    Attribution+Measurement section-title cards) is never added here -- it's
    dropped from every preset.
    """
    active = {"always", "client_title", "campaign_specs", "proposal_divider", "proposal_template"}

    if selections["preset"] == "extended":
        active.add("full_deck")
    # 'standard' intentionally does NOT add full_deck -- its base slide set
    # comes from the literal STANDARD_CORE_SLIDES allowlist instead (applied
    # in build_presentation), same as 'quick_pitch'.

    # The personalized avails table stands on its own -- audiences exist with
    # or without a vertical, so this is keyed off the toggle alone. A vertical
    # only supplies the *fallback* (its own static "PRECISION TARGETING"
    # slide) for when the personalized table is switched off; with no vertical
    # selected there simply is no fallback and neither slide appears.
    if selections.get("include_avails_template"):
        active.add("targeting_avails_template")

    vertical = selections.get("vertical")
    if vertical and vertical != "none":
        active.add(f"vertical:{vertical}")
        active.add("any_vertical")
        if not selections.get("include_avails_template"):
            active.add(f"vertical:{vertical}:targeting")

    if selections["spanish_campaign"]:
        active.add("spanish")

    if selections["tegna_positioning"]:
        active.add("tegna_positioning")

    products = selections["products"]

    if products.get("streaming_retargeting", {}).get("enabled"):
        active.add("streaming_retargeting")

    am = products.get("audience_marketplace", {})
    if am.get("enabled"):
        active.add("am")
        if am.get("audience_targeting_display") or am.get("audience_targeting_preroll"):
            active.add("am:audience")
        if am.get("geofencing_display") or am.get("geofencing_preroll"):
            active.add("am:geofencing")
        if am.get("site_retargeting_display") or am.get("site_retargeting_preroll"):
            active.add("am:retargeting")

    sports = products.get("live_sports", {})
    if sports.get("enabled") and sports.get("sports"):
        active.add("sports")
        active.add("sports_viewership_intro")
        for sport_key in sports["sports"]:
            active.add(f"sport:{sport_key}")
            # Only the viewership slide for the specific sport(s) picked --
            # not the whole viewership block.
            active.add(f"sport_viewership:{sport_key}")

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


def slides_to_keep(deck_slide_map, active_keys):
    return sorted(n for n, key in deck_slide_map.items() if key in active_keys)


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


def reorder_vertical_stats_slide(prs):
    """Move the selected vertical's industry-expertise/stats slide (the one
    headed "PREMION VERTICAL EXPERTISE:") to position 3, right after the
    client title and Campaign Specs slides. Only one vertical can be
    selected at a time, so any kept slide carrying that heading is
    unambiguously the right one -- no need to match it to a specific
    vertical name.
    """
    sldIdLst = prs.slides._sldIdLst
    sldId_elements = list(sldIdLst)
    target_sldId = None
    for slide, sldId in zip(prs.slides, sldId_elements):
        if "PREMION VERTICAL EXPERTISE:" in slide_map.extract_slide_text(slide).upper():
            target_sldId = sldId
            break
    if target_sldId is None:
        return
    sldIdLst.remove(target_sldId)
    sldIdLst.insert(2, target_sldId)


def _remap_relationship_ids(element, rid_map):
    """Rewrite every r:embed/r:id/r:link in a copied subtree to the new
    part's equivalent rId. Without this, a duplicated slide's pictures point
    at relationship IDs that mean something different (or nothing) on the
    part they now live in -- which is how a cloned slide ends up rendering
    the wrong image, or none.
    """
    for el in element.iter():
        for attr_name, value in list(el.attrib.items()):
            if attr_name.startswith("{" + _R_NS + "}") and value in rid_map:
                el.set(attr_name, rid_map[value])


def _next_free_slide_partname(prs):
    """The lowest /ppt/slides/slideN.xml not already taken by a reachable
    part.

    python-pptx names a newly added slide "slide{len(sldIdLst)+1}.xml", which
    assumes slides are numbered contiguously from 1. After build_presentation
    has deleted most of the master deck that is false -- the ~50 retained
    slides keep their original numbering from a 119-slide deck, so the "next"
    name collides with a slide that still exists. The collision is silent
    until save, where it surfaces only as a zipfile "Duplicate name" warning
    and one of the two slides overwrites the other in the package.
    """
    used = {str(part.partname) for part in prs.part.package.iter_parts()}
    n = 1
    while f"/ppt/slides/slide{n}.xml" in used:
        n += 1
    return PackURI(f"/ppt/slides/slide{n}.xml")


def duplicate_slide(prs, source_slide, insert_at=None):
    """Copy one slide (all shapes plus the relationships they reference) into
    a new slide, optionally moved to a specific position. Returns the new
    slide.

    python-pptx has no slide-copy API. The pieces that matter: a new slide is
    added from the *same layout* (so inherited placeholders/theme match), the
    layout's auto-cloned placeholders are then dropped because the copied
    shape tree is already complete, and each of the source slide's own
    relationships gets an equivalent on the new part with the copied XML
    rewritten to the new rIds.

    The notes-slide relationship is deliberately not copied -- a notes slide
    belongs to exactly one slide, so pointing two slides at the same notes
    part would be malformed.
    """
    # Claim a free partname *before* add_slide picks a colliding one, then
    # rename immediately -- relationship targets are serialized as paths
    # derived from partname at save time, so renaming now is safe.
    free_partname = _next_free_slide_partname(prs)
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    new_slide.part.partname = free_partname

    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)

    rid_map = {}
    for rId, rel in source_slide.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            rid_map[rId] = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rid_map[rId] = new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)

    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        _remap_relationship_ids(new_el, rid_map)
        new_slide.shapes._spTree.append(new_el)

    if insert_at is not None:
        sldIdLst = prs.slides._sldIdLst
        new_sldId = list(sldIdLst)[-1]  # add_slide appends
        sldIdLst.remove(new_sldId)
        sldIdLst.insert(insert_at, new_sldId)

    return new_slide


def slide_index(prs, slide):
    for i, s in enumerate(prs.slides):
        if s is slide or s.part is slide.part:
            return i
    return None


def build_presentation(master_path, selections):
    """Open the master deck and delete unselected slides. Returns the
    in-memory Presentation (not yet saved) plus slide counts, so callers can
    run personalization fill on the retained template slides before saving.
    """
    prs = Presentation(master_path)
    deck_slide_map = slide_map.build_slide_map_from_prs(prs)
    original_count = len(prs.slides._sldIdLst)

    active_keys = resolve_active_keys(selections)
    keep_numbers = set(slides_to_keep(deck_slide_map, active_keys))
    if selections["preset"] == "standard":
        keep_numbers |= (STANDARD_CORE_SLIDES & set(deck_slide_map))

    # Targeting/avails mutual exclusion applies globally, regardless of how
    # a slide ended up in keep_numbers -- a condition_key match or, for
    # "standard", its literal slide-number allowlist (which always includes
    # the personalized avails template's slide number). The personalized
    # template always wins over the vertical's static Precision Targeting
    # slide whenever both would otherwise be present.
    vertical = selections.get("vertical")
    if vertical and vertical != "none":
        avails_present = any(deck_slide_map.get(n) == "targeting_avails_template" for n in keep_numbers)
        if avails_present:
            static_targeting_key = f"vertical:{vertical}:targeting"
            keep_numbers = {n for n in keep_numbers if deck_slide_map.get(n) != static_targeting_key}

    # Delete back-to-front so earlier indices don't shift under us.
    for slide_number in range(original_count, 0, -1):
        if slide_number not in keep_numbers:
            delete_slide(prs, slide_number - 1)

    reorder_vertical_stats_slide(prs)

    return prs, original_count, len(keep_numbers)


# ---------------------------------------------------------------------------
# Token fill convention (master_v1_1_changelog.md "Fill convention"):
# find runs whose text contains a {{TOKEN}}, replace via run.text assignment.
# Never assign text_frame.text. Multi-item content (bullets, table rows) is
# filled by deep-copying the tokenized paragraph/row element and inserting
# clones as siblings, then filling each clone.
# ---------------------------------------------------------------------------
def _placeholder(token):
    return "{{" + token + "}}"


def _replace_tokens_in_text_frame(text_frame, values):
    for para in text_frame.paragraphs:
        for run in para.runs:
            for token, value in values.items():
                ph = _placeholder(token)
                if ph in run.text:
                    run.text = run.text.replace(ph, value)


def _fill_simple_tokens_in_slide(slide, values):
    """Single-value token replacement across one slide's shapes (including
    nested groups) and table cells."""
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.has_text_frame:
            _replace_tokens_in_text_frame(shape.text_frame, values)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    _replace_tokens_in_text_frame(cell.text_frame, values)


def fill_all_simple_tokens(prs, values):
    """Single-value token replacement across every slide in the deck."""
    for slide in prs.slides:
        _fill_simple_tokens_in_slide(slide, values)


def find_slide_with_marker(prs, marker):
    for slide in prs.slides:
        if marker in slide_map.extract_slide_text(slide):
            return slide
    return None


def _find_text_frame_with_token(shapes, token):
    ph = _placeholder(token)
    for shape in slide_map.iter_all_shapes(shapes):
        if shape.has_text_frame and ph in shape.text_frame.text:
            return shape.text_frame
    return None


def fill_bullet_list(text_frame, token, items):
    """Fill a multi-bullet section: first item goes in place via run.text;
    each additional item is a deep-copied clone of the tokenized paragraph,
    inserted as a following sibling and then filled the same way.
    """
    ph = _placeholder(token)
    target_para = None
    for para in text_frame.paragraphs:
        if any(ph in run.text for run in para.runs):
            target_para = para
            break
    if target_para is None:
        raise RuntimeError(f"token {ph} not found in text frame")

    template_p = copy.deepcopy(target_para._p)  # pristine, placeholder intact

    for run in target_para.runs:
        if ph in run.text:
            run.text = run.text.replace(ph, items[0])
            break

    last_p = target_para._p
    for extra_item in items[1:]:
        new_p = copy.deepcopy(template_p)
        last_p.addnext(new_p)
        last_p = new_p
        for run in target_para.__class__(new_p, target_para._parent).runs:
            if ph in run.text:
                run.text = run.text.replace(ph, extra_item)
                break


def fill_bullet_list_in_slide(slide, token, items):
    text_frame = _find_text_frame_with_token(slide.shapes, token)
    if text_frame is None:
        raise RuntimeError(f"token {_placeholder(token)} not found on slide")
    fill_bullet_list(text_frame, token, items)


def _find_table_shape(slide):
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.has_table:
            return shape
    return None


def clone_table_row(table, after_tr):
    new_tr = copy.deepcopy(after_tr)
    after_tr.addnext(new_tr)
    return new_tr


def fill_table_rows(slide, template_row_index, rows, field_to_token):
    """Fill a repeating table section: rows[0] fills the template row in
    place; each additional row is a deep-copied clone of the template <a:tr>,
    inserted directly after the previous one to preserve order.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        raise RuntimeError("no table found on slide")
    table = table_shape.table
    tbl = table._tbl

    template_tr = tbl.tr_lst[template_row_index]
    trs = [template_tr]
    last_tr = template_tr
    for _ in rows[1:]:
        new_tr = clone_table_row(table, last_tr)
        trs.append(new_tr)
        last_tr = new_tr

    for tr, row_data in zip(trs, rows):
        row_index = tbl.tr_lst.index(tr)
        row = table.rows[row_index]
        values = {token: row_data[field] for field, token in field_to_token.items()}
        for cell in row.cells:
            _replace_tokens_in_text_frame(cell.text_frame, values)


def _find_shape_by_name(shapes, name):
    for shape in slide_map.iter_all_shapes(shapes):
        if shape.name == name:
            return shape
    return None


def condense_media_plan_table(slide, num_data_rows, extra_total_rows=0):
    """Shrink cloned media-plan data rows (and font size, if needed) so the
    table never grows down into the "every advantage" graphic below it.
    Only touches anything if the rows at their natural (template) height
    would actually overlap -- small row counts are left at template size.
    """
    if num_data_rows <= 1 and extra_total_rows == 0:
        return

    table_shape = _find_table_shape(slide)
    floor_shape = _find_shape_by_name(slide.shapes, "Picture 12")
    if table_shape is None or floor_shape is None:
        return

    table = table_shape.table
    data_rows = list(table.rows)[1:-1]
    header_h = table.rows[0].height
    totals_h = table.rows[len(table.rows) - 1].height
    original_row_h = data_rows[0].height

    margin = Emu(50000)
    available = floor_shape.top - margin - table_shape.top
    natural_total = header_h + original_row_h * num_data_rows + totals_h * (1 + extra_total_rows)
    if natural_total <= available:
        return  # fits at template size already -- leave untouched

    budget_for_data = available - header_h - totals_h * (1 + extra_total_rows)
    min_row_h = Emu(280000)
    row_h = max(min_row_h, Emu(int(budget_for_data / num_data_rows)))

    for row in data_rows:
        row.height = row_h

    if row_h <= Emu(320000):
        font_pt = 8
    elif row_h <= Emu(420000):
        font_pt = 10
    else:
        font_pt = None

    if font_pt is not None:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(font_pt)


def add_full_flight_total_row(slide, label, impressions, cost):
    """Append one more row below the (already-filled) monthly totals row,
    showing the full-flight grand total -- a plain clone-and-overwrite since
    the totals row's {{TOKENS}} are already gone by the time this runs.
    """
    table_shape = _find_table_shape(slide)
    table = table_shape.table
    tbl = table._tbl

    totals_tr = tbl.tr_lst[-1]
    new_tr = clone_table_row(table, totals_tr)
    new_index = tbl.tr_lst.index(new_tr)
    cells = list(table.rows[new_index].cells)

    for cell, value in ((cells[0], label), (cells[4], impressions), (cells[5], cost)):
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.text = value
                break
            break


def swap_named_picture_everywhere(prs, shape_name, image_path):
    """Image placeholder swap: replace the image of every picture shape with
    the given name, keeping its existing position/size untouched."""
    for slide in prs.slides:
        for shape in slide_map.iter_all_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.name == shape_name:
                image_part, rId = shape.part.get_or_add_image_part(image_path)
                shape._element.blipFill.blip.rEmbed = rId


def _fill_media_plan_slide(slide, option):
    """Fill one media plan slide from one plan option. Its own scalar tokens
    are filled here rather than by the deck-wide pass, because with several
    options each slide's title/totals differ."""
    _fill_simple_tokens_in_slide(slide, {
        "PLAN_TITLE": option["plan_title"],
        "TOTALS_LABEL": option["totals_label"],
        "TOTAL_IMPRESSIONS": option["total_impressions"],
        "TOTAL_COST": option["total_cost"],
    })

    plan_rows = option["rows"]
    full_flight_total = option.get("full_flight_total")
    fill_table_rows(
        slide,
        template_row_index=1,
        rows=plan_rows,
        field_to_token={
            "tactic": "TACTIC", "flight": "FLIGHT", "geo": "GEO",
            "targeting": "TARGETING", "impressions": "IMPRESSIONS", "cost": "COST",
        },
    )
    condense_media_plan_table(slide, len(plan_rows), extra_total_rows=1 if full_flight_total else 0)
    if full_flight_total:
        add_full_flight_total_row(
            slide, full_flight_total["label"],
            full_flight_total["impressions"], full_flight_total["cost"],
        )
    fill_bullet_list_in_slide(slide, "INCLUDED_LIST", option["included_list"])


def media_plan_options(fill_data):
    """One entry per plan option. A single-option proposal is just a
    one-element list, and renders identically to how it always has."""
    options = fill_data.get("media_plan_options")
    if options:
        return options
    return [fill_data["media_plan"]]


def personalize(prs, fill_data):
    specs_slide = find_slide_with_marker(prs, "{{GOALS_BULLETS}}")
    avails_slide = find_slide_with_marker(prs, "{{AVAILS}}")
    plan_slide = find_slide_with_marker(prs, "{{TACTIC}}")
    options = media_plan_options(fill_data)

    # Clone the media plan template once per extra option, before any of the
    # {{TOKEN}}s are consumed -- each clone needs a pristine copy to fill.
    # They're inserted consecutively so the options read A, B, C in place of
    # where the single plan slide sits today.
    plan_slides = [plan_slide] if plan_slide is not None else []
    if plan_slide is not None and len(options) > 1:
        base_index = slide_index(prs, plan_slide)
        for offset in range(1, len(options)):
            plan_slides.append(duplicate_slide(prs, plan_slide, insert_at=base_index + offset))

    # Deck-wide scalars only. The media plan's own tokens are filled per
    # slide below, since they differ per option.
    fill_all_simple_tokens(prs, {
        "CLIENT_NAME": fill_data["client_name"],
        "PROPOSAL_TITLE": fill_data["proposal_title"],
        "VERTICAL": fill_data["vertical_display"],
        "TOTAL_AVAILS": fill_data["avails"]["total_avails"],
    })

    if specs_slide is not None:
        for section_token, bullets in fill_data["campaign_specs"].items():
            fill_bullet_list_in_slide(specs_slide, section_token, bullets)

    if avails_slide is not None:
        fill_table_rows(
            avails_slide,
            template_row_index=1,
            rows=fill_data["avails"]["rows"],
            field_to_token={"audience": "AUDIENCE", "geo": "GEO", "avails": "AVAILS"},
        )

    for slide, option in zip(plan_slides, options):
        _fill_media_plan_slide(slide, option)

    swap_named_picture_everywhere(prs, "CLIENT_LOGO", fill_data["logo_path"])


def assemble(master_path, output_path, selections, fill_data):
    prs, original_count, kept_count = build_presentation(master_path, selections)
    personalize(prs, fill_data)
    prs.save(output_path)
    return original_count, kept_count


if __name__ == "__main__":
    original_count, kept_count = assemble(MASTER_DECK_PATH, OUTPUT_PATH, SELECTIONS, FILL_DATA)
    print(f"Master deck: {original_count} slides")
    print(f"Kept: {kept_count} slides + personalized templates -> saved to {OUTPUT_PATH}")
