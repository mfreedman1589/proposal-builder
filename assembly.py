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
import hashlib
import math

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import slide_map

# Relationship-reference attributes (r:embed, r:id, r:link) live in this
# namespace. They're part-local, so copied shape XML has to be rewritten.
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

MASTER_DECK_PATH = "TEGNA_MASTER_DECK_v1_1.pptx"
OUTPUT_PATH = "test.pptx"

# The core content slides: one Premium Content highlight, the attribution
# overview, detailed reporting, and web attribution. Included by "standard"
# and "extended", left out of "quick_pitch".
#
# This used to be a literal slide-number allowlist ({1, 2, 3, 4, 5, 7, 13,
# 19, 21, 22, 118}), which silently assumed the master deck's numbering never
# moved. It does move -- that's the entire premise of the deck update page --
# so the selection is by condition_key like everything else. The numbers it
# used to carry decomposed into: 1/2/3/4/5 (already covered by the
# client_title / campaign_specs / always keys every preset includes), 7/19/21/
# 22 (these four, which needed their own anchors in slide_map since they'd
# otherwise fall through to the eleven-slide "full_deck" default), 13 (the
# avails template, see STANDARD_FORCED_KEYS) and 118 (the "Proposal Slides"
# divider, now dropped from every proposal anyway).
CORE_CONTENT_KEYS = frozenset({
    "core:premium_content",
    "core:attribution_overview",
    "core:reporting",
    "core:web_attribution",
})

# "Standard" pulls the personalized targeting/avails template in even when
# the form's own toggle is off -- preserved from the old allowlist, which
# listed slide 13 unconditionally. The targeting/avails mutual exclusion in
# build_presentation then drops the vertical's static Precision Targeting
# slide, exactly as it did before.
STANDARD_FORCED_KEYS = frozenset({"targeting_avails_template"})


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


# Total TV market-variant key prefixes. One place, because the selection
# logic, the mutual-exclusion sweep and the tests all have to agree on them.
COBRAND_TITLE_PREFIX = "client_title_cobrand:"
TOTAL_TV_PLAN_PREFIX = "proposal_template_total_tv:"
BROADCAST_SCHEDULE_PREFIX = "broadcast_schedule_template:"

# The master's own schedule slide is a static placeholder. Once a real
# schedule is imported the generated grid replaces it, and it's identified by
# this marker rather than by key -- it shares `total_tv:dc` with the WUSA
# pitch slide, so dropping by key would take the pitch slide with it.
SCHEDULE_PLACEHOLDER_MARKER = "TV SCHEDULE PLACEHOLDER"


def market_suffix(market):
    """`DC` -> `dc`, anything else -> `harrisburg`. The market half of every
    Total TV variant key."""
    return "dc" if market == "DC" else "harrisburg"


def resolve_active_keys(selections):
    """Turn the selection dictionary into the set of active condition_keys.

    Note: SECTION_DIVIDER_KEY is never added here, and build_presentation
    drops those slides outright -- no section-title card appears in a
    generated proposal under any preset.
    """
    active = {"always", "client_title", "campaign_specs", "proposal_divider", "proposal_template"}

    # 'standard' deliberately does not add full_deck -- it takes only the four
    # core content slides on top of the always-on ones. 'extended' takes both,
    # since those four are no longer part of full_deck now that they carry
    # their own keys.
    if selections["preset"] in ("standard", "extended"):
        active |= CORE_CONTENT_KEYS
    if selections["preset"] == "extended":
        active.add("full_deck")
    if selections["preset"] == "standard":
        active |= STANDARD_FORCED_KEYS

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

    # Dynamic Video Ads. The slide has always carried a `dynamic_creative`
    # key but nothing ever switched it on, so it could never appear in a
    # generated deck -- see the orphaned-keys note in SLIDE_KEYS.md.
    if selections.get("dynamic_creative"):
        active.add("dynamic_creative")

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
        market_key = market_suffix(selections["market"])
        active.add(f"total_tv:{market_key}")
        # Total TV drives the whole deck's branding -- there's no separate
        # toggle. The market's co-brand cover replaces the standard client
        # title and the market's Total TV template replaces the standard plan
        # slide; build_presentation drops the two standard slides so only one
        # of each survives.
        active.add(f"{COBRAND_TITLE_PREFIX}{market_key}")
        active.add(f"{TOTAL_TV_PLAN_PREFIX}{market_key}")
        # Only once a real Wide Orbit schedule has been read: without one
        # there's nothing to fill the grid with, so the master's existing
        # placeholder slide stays exactly as it is today.
        if selections.get("broadcast_schedule_imported"):
            active.add(f"{BROADCAST_SCHEDULE_PREFIX}{market_key}")

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


# ---------------------------------------------------------------------------
# Cross-deck slide copying (case study vault)
#
# duplicate_slide above copies a slide within one presentation, where every
# part it references already exists in the package. Pulling a slide out of a
# *different* .pptx is a bigger job: none of the parts it points at exist in
# the destination yet, so they have to be imported too, and their partnames
# and relationship ids renumbered to ones that are free on arrival.
# ---------------------------------------------------------------------------
_XML_MEDIA_SUFFIX = ".xml"


def _blank_layout(prs):
    """The emptiest layout in the destination deck.

    A slide copied from another deck can't bring its own layout with it, and
    inheriting a layout full of placeholders would stamp stray title/body
    boxes underneath the copied shapes. "Blank" is the conventional name;
    falling back to whichever layout has the fewest placeholders keeps this
    working on a deck that renamed it.
    """
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


class ImportCache:
    """Shared bookkeeping for a run of copy_slide_into calls.

    Two jobs, both of which need to span calls rather than live inside one:

    `parts` maps a source part (and, for leaf binaries, its content digest)
    to the destination part standing in for it, so the same image is stored
    once no matter how many slides or how many *source decks* reference it.
    Every case study deck carries its own copy of the same PREMION logos and
    icons, so a per-call cache deduplicates within a slide and then imports
    the identical bytes again for the next one.

    `used` is every partname already claimed. It can't be re-derived from
    the package on each allocation because a part that's been created but
    isn't related to anything yet is not reachable via iter_parts().

    Seeded from the destination's own media, so a case study sharing an
    asset with the master deck reuses the master's copy rather than adding a
    second one.
    """

    def __init__(self, dst_prs):
        self.parts = {}
        self.used = set()
        for part in dst_prs.part.package.iter_parts():
            partname = str(part.partname)
            self.used.add(partname)
            if partname.startswith("/ppt/media/"):
                self.parts[(part.content_type, hashlib.sha1(part.blob).hexdigest())] = part


def _next_free_partname(used, template_partname):
    """A free partname in the same family as template_partname.

    "/ppt/media/image7.svg" yields the lowest "/ppt/media/imageN.svg" not
    already taken -- the destination has its own image1..N and the numbering
    of the two decks has nothing to do with each other. `used` is carried
    across the whole import because a part that's been created but not yet
    related to anything isn't reachable via iter_parts() yet.
    """
    text = str(template_partname)
    directory, _, filename = text.rpartition("/")
    stem, dot, extension = filename.partition(".")
    base = stem.rstrip("0123456789") or "part"
    n = 1
    while f"{directory}/{base}{n}{dot}{extension}" in used:
        n += 1
    partname = PackURI(f"{directory}/{base}{n}{dot}{extension}")
    used.add(str(partname))
    return partname


def _import_part(src_part, dst_package, cache):
    """Copy one part (and everything it references) into dst_package.

    Deliberately not get_or_add_image_part: that re-encodes through
    python-pptx's Image class, which only understands raster formats and
    fails outright on the EMF and SVG parts these decks are full of. Copying
    the blob verbatim keeps any content type intact.

    Recursion is what makes charts work -- a chart part isn't a leaf, it
    carries its own relationships to an embedded workbook plus colour and
    style parts, and importing only the media would leave a chart pointing
    at nothing.
    """
    identity = id(src_part)
    if identity in cache.parts:
        return cache.parts[identity]

    blob = src_part.blob
    child_rels = [(rId, rel) for rId, rel in src_part.rels.items()
                  if rel.reltype != RT.NOTES_SLIDE]

    # Leaf binary parts (images, embedded workbooks) are deduplicated by
    # content: one physical copy however many slides point at it. Parts with
    # relationships of their own are not, since two identical blobs could
    # still resolve their rIds to different targets.
    digest = None
    if not child_rels:
        digest = (src_part.content_type, hashlib.sha1(blob).hexdigest())
        if digest in cache.parts:
            cache.parts[identity] = cache.parts[digest]
            return cache.parts[digest]

    new_part = Part(_next_free_partname(cache.used, src_part.partname),
                    src_part.content_type, dst_package, blob)
    cache.parts[identity] = new_part
    if digest is not None:
        cache.parts[digest] = new_part

    rid_map = {}
    for rId, rel in child_rels:
        if rel.is_external:
            rid_map[rId] = new_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rid_map[rId] = new_part.rels.get_or_add(
                rel.reltype, _import_part(rel.target_part, dst_package, cache))

    # An XML part addresses its own relationships by rId in its markup (a
    # chart's <c:externalData r:id="..."/>, say), and those ids were just
    # reassigned. Binary parts have no such references, so they're left as
    # the bytes they arrived as.
    if rid_map and str(src_part.partname).endswith(_XML_MEDIA_SUFFIX):
        root = etree.fromstring(blob)
        _remap_relationship_ids(root, rid_map)
        new_part._blob = etree.tostring(root, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)
    return new_part


def copy_slide_into(src_pptx_path, slide_index_in_src, dst_prs, position=None, cache=None):
    """Copy one slide out of another .pptx into dst_prs. Returns the new slide.

    Opens src_pptx_path, takes slide `slide_index_in_src` (0-based), and
    rebuilds it on a blank layout in the destination: background, shape tree,
    and every part the shapes reference, with all relationship ids rewritten
    to the destination's own. `position` is a 0-based index in the
    destination's slide order; None appends.

    Pass one ImportCache through a run of calls so shared assets are stored
    once across all of them; omitting it is correct but stores a fresh copy
    of every asset per call.
    """
    src_prs = Presentation(src_pptx_path)
    src_slide = src_prs.slides[slide_index_in_src]

    # Same collision hazard as duplicate_slide: after build_presentation has
    # deleted most of the master, the retained slides keep their original
    # numbering, so python-pptx's "slide{count+1}.xml" lands on a slide that
    # still exists and one silently overwrites the other on save.
    free_partname = _next_free_slide_partname(dst_prs)
    new_slide = dst_prs.slides.add_slide(_blank_layout(dst_prs))
    new_slide.part.partname = free_partname

    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)

    if cache is None:
        cache = ImportCache(dst_prs)

    rid_map = {}
    for rId, rel in src_slide.part.rels.items():
        # The new slide has its own layout, and a notes slide belongs to
        # exactly one slide -- sharing one between two is malformed.
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            rid_map[rId] = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rid_map[rId] = new_slide.part.rels.get_or_add(
                rel.reltype, _import_part(rel.target_part, dst_prs.part.package, cache))

    # Background lives on <p:cSld>, outside the shape tree, so a slide with a
    # full-bleed colour or picture background loses it unless it's copied
    # separately.
    src_bg = src_slide._element.cSld.find(qn("p:bg"))
    if src_bg is not None:
        new_bg = copy.deepcopy(src_bg)
        _remap_relationship_ids(new_bg, rid_map)
        new_slide._element.cSld.insert(0, new_bg)

    for shape in src_slide.shapes:
        new_element = copy.deepcopy(shape._element)
        _remap_relationship_ids(new_element, rid_map)
        new_slide.shapes._spTree.append(new_element)

    if position is not None:
        sldIdLst = dst_prs.slides._sldIdLst
        new_sldId = list(sldIdLst)[-1]  # add_slide appends
        sldIdLst.remove(new_sldId)
        sldIdLst.insert(position, new_sldId)

    return new_slide


def case_study_insert_index(prs):
    """Where case study slides belong in an already-assembled deck (0-based).

    Immediately before the media plan slide: case studies are the proof that
    closes the argument, so they sit at the end of the content, after the
    attribution slides, and the plan is the next thing the client sees.

    Anchored to the plan slide itself rather than to the "Proposal Slides"
    divider, so they stay put on a preset that drops the divider. **This must
    run before personalize()** -- for two independent reasons. The map is
    re-derived by scanning, and the plan slide is identified by its
    {{PLAN_TITLE}} token, which personalize consumes; and personalize clones
    the plan slide once per extra option *directly after* the original, so
    inserting ahead of the original here is what puts case studies before the
    first option rather than between options.

    Section dividers are dropped from every proposal, so in practice the plan
    slide is the anchor. The divider fallback is kept anyway rather than
    deleted: it costs one line, and a future master deck that reintroduces a
    divider (or a preset that somehow retains one) should degrade to "before
    the proposal section" instead of silently dumping case studies at the end
    of the deck, after the plan.
    """
    assembled = slide_map.build_slide_map_from_prs(prs)
    for wanted in ("proposal_template", slide_map.SECTION_DIVIDER_KEY):
        for n, key in sorted(assembled.items()):
            if key == wanted:
                return n - 1  # 1-based position N -> 0-based index before it
    return len(assembled)


def append_case_studies(prs, case_studies):
    """Graft every selected case study's slides into the assembled deck.

    `case_studies` is an ordered list of {"path": ..., "slides": [...]}, where
    "slides" is a list of 0-based slide indices to take (None for all). One
    ImportCache spans the whole run so shared assets -- every one of these
    decks carries the same PREMION logos -- are stored once.
    """
    if not case_studies:
        return 0

    cache = ImportCache(prs)
    position = case_study_insert_index(prs)
    copied = 0

    for case_study in case_studies:
        source = Presentation(case_study["path"])
        indices = case_study.get("slides")
        if indices is None:
            indices = range(len(source.slides._sldIdLst))
        for index in indices:
            copy_slide_into(case_study["path"], index, prs, position=position, cache=cache)
            position += 1
            copied += 1
    return copied


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

    # Targeting/avails mutual exclusion applies globally, regardless of how a
    # slide ended up in keep_numbers -- a toggle, or "standard" forcing the
    # avails template in. The personalized template always wins over the
    # vertical's static Precision Targeting slide whenever both would
    # otherwise be present.
    vertical = selections.get("vertical")
    if vertical and vertical != "none":
        avails_present = any(deck_slide_map.get(n) == "targeting_avails_template" for n in keep_numbers)
        if avails_present:
            static_targeting_key = f"vertical:{vertical}:targeting"
            keep_numbers = {n for n in keep_numbers if deck_slide_map.get(n) != static_targeting_key}

    # Total TV replaces two standard slides rather than adding beside them.
    # Done here rather than by omitting keys in resolve_active_keys for the
    # same reason as the avails rule above: it then holds however a slide got
    # into keep_numbers, including a preset that forces one in. Two covers or
    # two plan templates in one deck would be a visible, embarrassing bug --
    # and personalize() finds the plan slide by token, so a second one would
    # also silently take the fill.
    if selections["products"].get("total_tv"):
        cobrand_present = any(str(deck_slide_map.get(n, "")).startswith(COBRAND_TITLE_PREFIX)
                              for n in keep_numbers)
        if cobrand_present:
            keep_numbers = {n for n in keep_numbers if deck_slide_map.get(n) != "client_title"}
        tt_plan_present = any(str(deck_slide_map.get(n, "")).startswith(TOTAL_TV_PLAN_PREFIX)
                              for n in keep_numbers)
        if tt_plan_present:
            keep_numbers = {n for n in keep_numbers if deck_slide_map.get(n) != "proposal_template"}

        # A generated schedule grid supersedes the static placeholder.
        if selections.get("broadcast_schedule_imported") and any(
                str(deck_slide_map.get(n, "")).startswith(BROADCAST_SCHEDULE_PREFIX)
                for n in keep_numbers):
            keep_numbers = {
                n for n in keep_numbers
                if SCHEDULE_PLACEHOLDER_MARKER not in
                slide_map.extract_slide_text(prs.slides[n - 1]).upper()}

    # A Total TV variant that somehow survived for the wrong market would put
    # a competitor station's branding in front of a client, so both variants
    # are swept unconditionally rather than trusted to have been selected
    # correctly upstream.
    wanted_suffix = market_suffix(selections["market"])
    for prefix in (COBRAND_TITLE_PREFIX, TOTAL_TV_PLAN_PREFIX, BROADCAST_SCHEDULE_PREFIX):
        keep_numbers = {
            n for n in keep_numbers
            if not (str(deck_slide_map.get(n, "")).startswith(prefix)
                    and deck_slide_map.get(n) != f"{prefix}{wanted_suffix}")}

    # Section dividers never appear in a generated proposal, in any preset.
    # Kept as a post-processing sweep rather than folded into
    # resolve_active_keys so that it holds no matter how a slide got into
    # keep_numbers -- a carry-forward default, a forced key, or some future
    # rule. Back when "standard" was a literal slide-number allowlist this
    # was load-bearing (the list contained slide 118, the "Proposal Slides"
    # divider); it's belt-and-braces now, and cheap.
    keep_numbers = {n for n in keep_numbers
                    if deck_slide_map.get(n) != slide_map.SECTION_DIVIDER_KEY}

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


# A table row is only as short as its text allows: PowerPoint grows a row
# whose contents don't fit, so setting a small height without also shrinking
# the font achieves nothing. Roughly, a line of N-point text occupies N * 1.2
# points, plus the cell's top and bottom insets.
_LINE_SPACING = 1.2
_EMU_PER_POINT = 12700
_DEFAULT_CELL_INSET = Emu(45720)  # 0.05in, python-pptx's default top/bottom
_MIN_TABLE_FONT_PT = 6


def _content_floor(slide, table_shape):
    """The top of the highest shape sitting below the table and overlapping it
    horizontally -- i.e. the first thing the table would collide with.

    Derived from the slide rather than by shape name. The previous version
    looked up "Picture 12" (the "One Solution" graphic at 4.70in) and missed
    that the "Included with Campaign:" heading starts higher at 4.63in, so
    the table was sized against a ceiling 0.07in below the one that actually
    mattered -- and any deck that renamed or re-layered that picture lost the
    check entirely.
    """
    table_bottom_start = table_shape.top
    table_left, table_right = table_shape.left, table_shape.left + table_shape.width

    floor = None
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape is table_shape or shape.top is None or shape.left is None:
            continue
        if shape.top <= table_bottom_start:
            continue  # above the table, or the table itself
        # only things that actually sit under the table's own columns matter
        if shape.left + (shape.width or 0) <= table_left or shape.left >= table_right:
            continue
        floor = shape.top if floor is None else min(floor, shape.top)
    return floor


def _max_font_for_row(row_height_emu):
    """The largest whole point size whose line box fits in a row of this
    height, so the row height we ask for is one PowerPoint can honour."""
    usable = row_height_emu - 2 * _DEFAULT_CELL_INSET
    return max(_MIN_TABLE_FONT_PT, int(usable / (_LINE_SPACING * _EMU_PER_POINT)))


def condense_media_plan_table(slide, num_data_rows, extra_total_rows=0):
    """Shrink cloned media-plan rows (and their font) so the table can't grow
    into the "Included with Campaign" block or the graphic beneath it.

    Only touches anything if the rows at their natural (template) height would
    actually overlap -- small row counts are left exactly as the template has
    them.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    floor = _content_floor(slide, table_shape)
    if floor is None:
        return None

    table = table_shape.table
    data_rows = list(table.rows)[1:-1]
    if not data_rows:
        return None
    header_h = table.rows[0].height
    totals_h = table.rows[len(table.rows) - 1].height
    original_row_h = data_rows[0].height

    margin = Emu(50000)
    available = floor - margin - table_shape.top
    natural_total = header_h + original_row_h * num_data_rows + totals_h * (1 + extra_total_rows)
    if natural_total <= available:
        return None  # fits at template size already

    # The smallest a row can be and still hold readable text.
    min_row_h = Emu(int(_MIN_TABLE_FONT_PT * _LINE_SPACING * _EMU_PER_POINT) + 2 * _DEFAULT_CELL_INSET)
    total_rows = 1 + num_data_rows + (1 + extra_total_rows)

    # Everything that isn't a data row is fixed overhead, and at high row
    # counts the data can't fit inside what's left of it -- so the header and
    # totals rows shrink too, down to the same floor as the data rows. An
    # earlier version clamped them at half their template height instead,
    # which left 12 rows plus a full-flight row overflowing by 0.21in.
    fixed = header_h + totals_h * (1 + extra_total_rows)
    if available - fixed < min_row_h * num_data_rows:
        spare = available - min_row_h * num_data_rows
        share = max(min_row_h, Emu(int(spare / (1 + 1 + extra_total_rows))) if spare > 0 else min_row_h)
        header_h = totals_h = min(header_h, share)
        table.rows[0].height = header_h
        # Only the template's own totals row exists yet; the full-flight row is
        # cloned from it afterwards and inherits whatever we set here.
        table.rows[len(table.rows) - 1].height = totals_h
        fixed = header_h + totals_h * (1 + extra_total_rows)

    row_h = max(min_row_h, Emu(int((available - fixed) / num_data_rows)))
    for row in data_rows:
        row.height = row_h

    # Past this point the table physically cannot fit: every row is already at
    # the minimum readable height. Say so rather than shipping an overlap.
    if min_row_h * total_rows > available:
        overflow_warning = (
            f"The media plan has {num_data_rows} lines, which is more than the slide can "
            f"hold without running into the \"Included with Campaign\" block "
            f"(about {int(available / min_row_h) - 2 - extra_total_rows} lines is the limit). "
            f"Split it across plan options, or combine some lines.")
    else:
        overflow_warning = None

    # Derived from the row height rather than picked off a coarse ladder, so
    # the text is always guaranteed to fit the row it's in.
    font_pt = min(_max_font_for_row(row_h), _max_font_for_row(min(header_h, totals_h)))
    for row in table.rows:
        for cell in row.cells:
            cell.margin_top = Emu(0)
            cell.margin_bottom = Emu(0)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size is None or run.font.size.pt > font_pt:
                        run.font.size = Pt(font_pt)
    return overflow_warning


# ---------------------------------------------------------------------------
# Broadcast schedule slides (Feature C)
#
# The template ships with three week columns; a real schedule has anywhere
# from one to twenty-one. Columns are cloned from the styled ones already in
# the template rather than built from scratch, for exactly the reason rows
# are: a hand-built <a:gridCol>/<a:tc> carries none of the theme's borders,
# fills or margins, and the result looks like a different deck.
# ---------------------------------------------------------------------------

# Where the template's own week columns sit, and what surrounds them.
WEEK_COL_START = 6           # first {{WKn}} column
TEMPLATE_WEEK_COLS = 3       # how many the template ships with
TRAILING_COLS = 2            # Total # and (000), which always follow the weeks

# Past this many week columns the grid stops being readable at any font size
# that fits the row height, so a detailed full-flight view paginates instead.
MAX_WEEK_COLUMNS = 10


def _grid_cols(table):
    return list(table._tbl.tblGrid.findall(qn("a:gridCol")))


def _row_cells(table):
    """The <a:tc> elements of each row, in document order."""
    return [list(tr.findall(qn("a:tc"))) for tr in table._tbl.findall(qn("a:tr"))]


def clone_table_column(table, source_index):
    """Duplicate one column, inserting the copy immediately after it.

    Clones the <a:gridCol> and every row's <a:tc> at that index, so the copy
    inherits the template's borders, fills and cell margins. Building the XML
    by hand instead produces a column that is subtly and visibly not part of
    the table.
    """
    grid = table._tbl.tblGrid
    cols = _grid_cols(table)
    source_col = cols[source_index]
    grid.insert(list(grid).index(source_col) + 1, copy.deepcopy(source_col))

    for tr in table._tbl.findall(qn("a:tr")):
        cells = list(tr.findall(qn("a:tc")))
        source_cell = cells[source_index]
        tr.insert(list(tr).index(source_cell) + 1, copy.deepcopy(source_cell))


def remove_table_column(table, index):
    grid = table._tbl.tblGrid
    cols = _grid_cols(table)
    grid.remove(cols[index])
    for tr in table._tbl.findall(qn("a:tr")):
        cells = list(tr.findall(qn("a:tc")))
        tr.remove(cells[index])


def set_week_column_count(table, count):
    """Grow or shrink the week columns to `count`, keeping the table's total
    width -- the surrounding columns are fixed, so the week block absorbs the
    difference and every column stays legible rather than the table growing
    off the slide."""
    current = len(_grid_cols(table)) - WEEK_COL_START - TRAILING_COLS
    # python-pptx's column collection doesn't slice.
    week_block = sum(table.columns[i].width
                     for i in range(WEEK_COL_START, WEEK_COL_START + current))

    while current < count:
        clone_table_column(table, WEEK_COL_START)
        current += 1
    while current > count:
        remove_table_column(table, WEEK_COL_START)
        current -= 1

    if count:
        each = int(week_block / count)
        for index in range(WEEK_COL_START, WEEK_COL_START + count):
            table.columns[index].width = Emu(each)
    return count


def _set_cell(table, row, col, text):
    """Write a cell's text without flattening its run formatting."""
    cell = table.cell(row, col)
    paragraphs = cell.text_frame.paragraphs
    if paragraphs and paragraphs[0].runs:
        paragraphs[0].runs[0].text = text
        for extra in paragraphs[0].runs[1:]:
            extra.text = ""
        for para in paragraphs[1:]:
            for run in para.runs:
                run.text = ""
    else:
        cell.text_frame.text = text


def week_column_labels(weeks, style):
    """Header text per week column.

    A Campaign Schedule Report labels its columns with real dates, so the
    deck shows them; a Planner export labels them by position, so the deck
    says "Wk N" rather than inventing a precision the source didn't have.
    """
    if style == "date":
        return [f"{week.month}/{week.day:02d}" for week in weeks]
    return [str(index) for index in range(1, len(weeks) + 1)]


def _page_weeks(weeks, breakout, detailed):
    """[(label, [weeks])] -- one entry per schedule slide."""
    if not detailed or not weeks:
        return [("", list(weeks))]
    if breakout == "monthly":
        pages, current, key = [], [], None
        for week in weeks:
            if (week.year, week.month) != key:
                if current:
                    pages.append(current)
                current, key = [], (week.year, week.month)
            current.append(week)
        if current:
            pages.append(current)
        return [(f"{p[0]:%B %Y}", p) for p in pages]
    if len(weeks) <= MAX_WEEK_COLUMNS:
        return [("", list(weeks))]
    pages = [weeks[i:i + MAX_WEEK_COLUMNS] for i in range(0, len(weeks), MAX_WEEK_COLUMNS)]
    labels, start = [], 1
    for page in pages:
        labels.append(f"Weeks {start}–{start + len(page) - 1}")
        start += len(page)
    return list(zip(labels, pages))


def build_broadcast_schedule_slides(prs, schedule, plan_description="",
                                    breakout="full_flight", detailed=True):
    """Turn one parsed Wide Orbit schedule into slides. Returns warnings.

    The template slide is cloned once per page and then removed, so a
    schedule that needs three pages produces three slides in place of the one
    template -- the same duplicate-then-fill approach the media plan uses for
    multiple options.
    """
    template = find_slide_with_marker(prs, _placeholder("BROADCAST_PLAN_DESC"))
    if template is None:
        return []

    weeks = schedule.grid_weeks
    pages = _page_weeks(weeks, breakout, detailed)
    warnings = []

    # A detailed full flight past the column cap can't be shown week by week
    # at a readable size. Falling back to totals-only keeps the numbers
    # rather than shipping an unreadable grid, and says so.
    if detailed and breakout != "monthly" and len(weeks) > MAX_WEEK_COLUMNS and len(pages) == 1:
        detailed = False
        pages = [("", list(weeks))]
        warnings.append(
            f"This schedule runs {len(weeks)} weeks, too many to show week by week on one "
            f"slide, so it's shown as totals only. Switch to a Monthly breakout to keep the "
            f"weekly detail.")

    base_index = slide_index(prs, template)
    made = []
    for offset, (label, page_weeks) in enumerate(pages):
        slide = duplicate_slide(prs, template, insert_at=base_index + offset)
        made.append(slide)
        page_label = label or ("Monthly Broadcast Plan:" if breakout == "monthly"
                               else "Broadcast Plan:")
        overflow = _fill_broadcast_slide(
            slide, schedule, page_weeks, detailed=detailed,
            plan_label=("Monthly Broadcast Plan:" if breakout == "monthly" else page_label),
            plan_description=plan_description,
            is_last=(offset == len(pages) - 1), all_weeks=weeks)
        if overflow:
            warnings.append(overflow)

    delete_slide(prs, slide_index(prs, template))
    return warnings


def _fill_broadcast_slide(slide, schedule, weeks, detailed, plan_label,
                          plan_description, is_last, all_weeks):
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    table = table_shape.table

    count = set_week_column_count(table, len(weeks) if detailed else 0)
    labels = week_column_labels(weeks, getattr(schedule, "week_header_style", "index")) \
        if detailed else []

    # Two header rows: the template's "Wk" band and the numbers beneath it.
    for offset, text in enumerate(labels):
        col = WEEK_COL_START + offset
        _set_cell(table, 0, col, "" if getattr(schedule, "week_header_style", "") == "date" else "Wk")
        _set_cell(table, 1, col, text)

    # Every program row appears on every page. Filtering to rows with spots in
    # this page's weeks would make the programme list change from page to
    # page, so a reader comparing two pages of one schedule would see rows
    # appear and vanish -- and a page whose weeks are all empty (a hiatus
    # week, which a football schedule has) would come out with no rows at all.
    rows = list(schedule.rows)

    template_tr = table._tbl.findall(qn("a:tr"))[2]
    for _ in range(len(rows) - 1):
        clone_table_row(table, template_tr)

    for index, row in enumerate(rows):
        r = 2 + index
        _set_cell(table, r, 0, row.station)
        _set_cell(table, r, 1, row.time)
        _set_cell(table, r, 2, row.days)
        _set_cell(table, r, 3, row.program)
        _set_cell(table, r, 4, row.length)
        _set_cell(table, r, 5, f"${row.rate:,.0f}")
        for offset, week in enumerate(weeks if detailed else []):
            _set_cell(table, r, WEEK_COL_START + offset,
                      str(row.spots_per_week.get(week, "") or ""))
        spots = sum(row.spots_per_week.get(w, 0) for w in weeks) if detailed else row.total_spots
        _set_cell(table, r, WEEK_COL_START + count, str(spots))
        _set_cell(table, r, WEEK_COL_START + count + 1, f"{row.impressions / 1000:,.1f}")

    totals_r = 2 + len(rows)
    page_spots = sum(sum(r.spots_per_week.get(w, 0) for w in weeks) for r in rows) if detailed \
        else sum(r.total_spots for r in rows)
    page_cost = sum(r.rate * sum(r.spots_per_week.get(w, 0) for w in weeks) for r in rows) if detailed \
        else sum(r.cost for r in rows)
    page_impressions = (schedule.summary.impressions if not detailed or len(weeks) == len(all_weeks)
                        else sum(r.impressions * (sum(r.spots_per_week.get(w, 0) for w in weeks)
                                                  / r.total_spots) for r in rows if r.total_spots))

    # Every page totals itself; the last one also carries the flight's grand
    # totals, so a multi-page schedule adds up on the page a reader stops at.
    _set_cell(table, totals_r, 0, "FLIGHT TOTALS" if is_last else f"{plan_label} totals")
    _set_cell(table, totals_r, 5, f"${(schedule.summary.gross_cost if is_last else page_cost):,.0f}")
    for offset, week in enumerate(weeks if detailed else []):
        _set_cell(table, totals_r, WEEK_COL_START + offset,
                  str(sum(r.spots_per_week.get(week, 0) for r in rows) or ""))
    _set_cell(table, totals_r, WEEK_COL_START + count,
              str(schedule.summary.total_spots if is_last else page_spots))
    _set_cell(table, totals_r, WEEK_COL_START + count + 1,
              f"{(schedule.summary.impressions if is_last else page_impressions) / 1000:,.1f}")

    summary = schedule.summary
    demo = summary.demo_label or "Adults"
    _fill_simple_tokens_in_slide(slide, {
        "BROADCAST_PLAN_LABEL": plan_label,
        "BROADCAST_PLAN_DESC": plan_description or "",
        "DEMO_LABEL": f"{demo} (000)",
        "TOTALS_LABEL": "FLIGHT TOTALS" if is_last else f"{plan_label} totals",
    })

    lines = [
        f"{summary.impressions:,.0f} {demo} impressions",
        f"{summary.total_spots:,} commercials",
        f"${summary.gross_cost:,.0f} gross",
    ]
    if summary.reach:
        lines.append(f"{summary.reach:.1f} reach")
    if summary.frequency:
        lines.append(f"{summary.frequency:.1f} frequency")
    try:
        fill_bullet_list_in_slide(slide, "BROADCAST_SUMMARY", lines)
    except RuntimeError:
        pass

    # Same layout-floor verification the media plan table gets: derive the
    # font from the final row height and confirm the table clears whatever
    # sits below it. Only the wording differs -- a seller reading "the media
    # plan has too many lines" on a broadcast schedule slide would go looking
    # in the wrong place.
    if condense_media_plan_table(slide, len(rows), extra_total_rows=0):
        # Deliberately not suggesting a different breakout or view: program
        # rows repeat on every page, so neither Monthly nor totals-only
        # changes the row count. The only thing that helps is fewer programs.
        return (f"This schedule's {len(rows)} programs don't fit on the slide even at the "
                f"smallest readable size -- about 11 fit. The table will run into the summary "
                f"below it. Trim the schedule in Wide Orbit, or import it in two parts.")
    return None


def table_bottom(slide):
    """Where the table actually ends, for verifying it cleared the content
    below it. python-pptx reports a table's height as the sum of its row
    heights, which is what we just set."""
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    return table_shape.top + sum(row.height for row in table_shape.table.rows)


# Campaign Specs auto-fit. python-pptx can't measure rendered text, and
# PowerPoint's own <a:normAutofit> only recalculates when the file is *edited*
# -- a generated deck that's merely opened and presented keeps whatever scale
# is stored. So the size is computed here and written onto the runs directly,
# which renders identically everywhere.
_CHAR_WIDTH_RATIO = 0.5   # average glyph advance as a fraction of point size
_SPECS_MIN_SCALE = 0.6    # never shrink past this; truncate instead
_SPECS_STEPS = (1.0, 0.92, 0.85, 0.78, 0.72, 0.66, 0.6)


def _estimate_frame_height(text_frame, width_emu, scale):
    """Approximate rendered height of a text frame at a given font scale.

    An estimate, deliberately: exact metrics need the actual font, which isn't
    available here. It's tuned to over- rather than under-estimate, since
    guessing high costs a little whitespace and guessing low costs an overlap.
    """
    usable_pt = max(1.0, (width_emu - 2 * Emu(91440)) / _EMU_PER_POINT)
    total_pt = 0.0
    for para in text_frame.paragraphs:
        sizes = [run.font.size.pt for run in para.runs if run.font.size]
        base_pt = (max(sizes) if sizes else 18.0) * scale
        text = "".join(run.text for run in para.runs)
        chars_per_line = max(1.0, usable_pt / (base_pt * _CHAR_WIDTH_RATIO))
        lines = max(1, math.ceil(len(text) / chars_per_line))
        total_pt += lines * base_pt * _LINE_SPACING
        if para.space_before is not None:
            total_pt += para.space_before.pt
        if para.space_after is not None:
            total_pt += para.space_after.pt
    return Emu(int(total_pt * _EMU_PER_POINT))


def _apply_font_scale(text_frame, scale):
    for para in text_frame.paragraphs:
        for run in para.runs:
            if run.font.size is not None:
                run.font.size = Pt(round(run.font.size.pt * scale, 1))


def fit_text_frame(text_frame, available_emu, width_emu):
    """Shrink a text frame's type until it fits, in steps, down to a floor.

    Returns (scale_applied, fits). fits=False means even the floor size
    overflows and the caller should shorten the copy -- the point of the floor
    is that shrinking past it produces a slide nobody can read, which is worse
    than an honest warning.
    """
    for scale in _SPECS_STEPS:
        if _estimate_frame_height(text_frame, width_emu, scale) <= available_emu:
            if scale < 1.0:
                _apply_font_scale(text_frame, scale)
            return scale, True
    _apply_font_scale(text_frame, _SPECS_MIN_SCALE)
    return _SPECS_MIN_SCALE, False


def fit_campaign_specs(slide):
    """Auto-fit the Campaign Specs body. Returns (scale, fits).

    The body is one tall frame holding all six labelled sections, and the
    template gives it no autofit at all -- drafted copy simply ran off the
    slide and over the client-name block below it.
    """
    frame_shape = None
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.has_text_frame and "Campaign Specs" not in shape.text_frame.text:
            text = shape.text_frame.text
            if text.count("\n") >= 3 and len(text) > 40:
                if frame_shape is None or shape.height > frame_shape.height:
                    frame_shape = shape
    if frame_shape is None:
        return 1.0, True

    floor = _content_floor(slide, frame_shape)
    bottom_limit = floor if floor is not None else Emu(int(7.5 * 914400))
    available = bottom_limit - Emu(50000) - frame_shape.top
    return fit_text_frame(frame_shape.text_frame, available, frame_shape.width)


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
    overflow = condense_media_plan_table(
        slide, len(plan_rows), extra_total_rows=1 if full_flight_total else 0)
    if full_flight_total:
        add_full_flight_total_row(
            slide, full_flight_total["label"],
            full_flight_total["impressions"], full_flight_total["cost"],
        )
    fill_bullet_list_in_slide(slide, "INCLUDED_LIST", option["included_list"])
    return overflow


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

    warnings = []
    if specs_slide is not None:
        for section_token, bullets in fill_data["campaign_specs"].items():
            fill_bullet_list_in_slide(specs_slide, section_token, bullets)
        # After filling, not before -- the bullets are what make it overflow.
        scale, fits = fit_campaign_specs(specs_slide)
        if not fits:
            warnings.append(
                "The Campaign Specs copy is too long for the slide even at the smallest "
                "readable type size. It has been set at "
                f"{int(scale * 100)}% and will still run over -- shorten the longest "
                "sections (Goals, Audience or Placements) and generate again.")

    if avails_slide is not None:
        fill_table_rows(
            avails_slide,
            template_row_index=1,
            rows=fill_data["avails"]["rows"],
            field_to_token={"audience": "AUDIENCE", "geo": "GEO", "avails": "AVAILS"},
        )

    for slide, option in zip(plan_slides, options):
        overflow = _fill_media_plan_slide(slide, option)
        if overflow:
            name = option.get("plan_title", "the media plan")
            warnings.append(f"{name}: {overflow}")

    swap_named_picture_everywhere(prs, "CLIENT_LOGO", fill_data["logo_path"])
    return warnings


def assemble(master_path, output_path, selections, fill_data):
    prs, original_count, kept_count = build_presentation(master_path, selections)
    for warning in personalize(prs, fill_data):
        print(f"WARNING: {warning}")
    prs.save(output_path)
    return original_count, kept_count


if __name__ == "__main__":
    original_count, kept_count = assemble(MASTER_DECK_PATH, OUTPUT_PATH, SELECTIONS, FILL_DATA)
    print(f"Master deck: {original_count} slides")
    print(f"Kept: {kept_count} slides + personalized templates -> saved to {OUTPUT_PATH}")
