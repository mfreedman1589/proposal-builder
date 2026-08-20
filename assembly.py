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
import functools
import hashlib
import io
import math
import re

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import slide_inheritance
import slide_map
import text_metrics
import wideorbit

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
            "viewership": False,
        },
        "total_tv": True,
    },
    # None means "this proposal predates the vertical-attribution checkbox",
    # which is what a rebuild of an older row carries -- see build_presentation.
    "vertical_attribution": None,
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


# Two verticals sell a named, branded measurement product instead of the
# generic CRM sales attribution: Polk Signals for automotive, Arrivalist for
# travel. Each has its own slide in the master, and each gets a checkbox in
# the attribution section (checked by default) that controls both the slide
# and the "Included with Campaign" line.
#
# `marker`, not a condition_key: both slides carry their vertical's own key
# (`vertical:auto` / `vertical:travel`), shared with that vertical's stats and
# precision-targeting slides, so a key can't single them out -- the same
# situation SCHEDULE_PLACEHOLDER_MARKER exists for. Each marker is the slide's
# own section footer, which is unique across the deck (verified on master
# version 6: one slide each) and, unlike a heading, isn't what gets rewritten
# when somebody retitles a slide.
#
# `replaces_sales_slide` is the difference between the two, and it isn't
# arbitrary. Polk Signals IS new-car sales attribution, so showing it beside
# the generic "Measure Sales Conversions" slide says the same thing twice.
# Arrivalist measures visits to a destination rather than closed sales, so the
# generic slide still has something of its own to say next to it.
VERTICAL_ATTRIBUTION = {
    "auto": {
        "label": "Polk New Car Sales Attribution",
        "marker": "PREMION + AUTOMOTIVE | ATTRIBUTION + MEASUREMENT",
        "replaces_sales_slide": True,
    },
    "travel": {
        "label": "Arrivalist Destination Attribution",
        "marker": "PREMION + TRAVEL & TOURISM | ATTRIBUTION + MEASUREMENT",
        "replaces_sales_slide": False,
    },
}


def vertical_attribution_spec(vertical):
    """The vertical's own branded attribution product, or None."""
    if not vertical or vertical == "none":
        return None
    return VERTICAL_ATTRIBUTION.get(vertical)


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
    # `market_profile` is unconditional, exactly as it was when that slide
    # carried the `always` key. Tagging it so the assembler can FIND it must
    # not change whether it SHIPS: the national Total U.S. profile is in every
    # deck today, and a preset quietly losing it would be a regression paid
    # for by a rep who never asked for one. What the tag buys is the ability
    # to replace it (see replace_market_profile_slides), not to drop it.
    active = {"always", "client_title", "campaign_specs", "proposal_divider",
              "proposal_template", "market_profile"}

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
        # Package slides only by default. Each selected sport used to bring
        # its MRI viewership chart along automatically, so a four-sport
        # proposal pulled nine slides where four were asked for. Viewership is
        # now one deck-wide opt-in covering every selected sport -- deliberately
        # not per sport, since a deck showing the stats for two of its four
        # packages reads as an omission rather than a choice.
        #
        # Absent means True, and only for a proposal logged before the toggle
        # existed: those were all built with the viewership slides in, and
        # "Rebuild as presented" has to reproduce what the client received. The
        # form always writes the flag explicitly, so this default never
        # governs a new proposal.
        show_viewership = sports.get("viewership", True)
        if show_viewership:
            active.add("sports_viewership_intro")
        for sport_key in sports["sports"]:
            active.add(f"sport:{sport_key}")
            # Only the viewership slide for the specific sport(s) picked --
            # not the whole viewership block.
            if show_viewership:
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


MARKET_PROFILE_KEY = "market_profile"


def _full_bleed_picture(slide):
    """The picture that IS this slide, or None.

    The profile slides are a single full-bleed metafile, so this is the
    largest picture by area rather than a shape looked up by name -- the
    name on the real deck is "Picture 4", which also occurs on the
    healthcare targeting slide. Matching by name there would have swapped a
    market profile into a Crossix graphic, which is exactly the silent,
    client-visible corruption this project keeps designing against.
    """
    pictures = [s for s in slide.shapes
                if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if not pictures:
        return None
    return max(pictures, key=lambda s: (s.width or 0) * (s.height or 0))


def replace_market_profile_slides(prs, image_paths):
    """Replace the national profile slide with one slide per target market.

    `image_paths` is one local image per selected market, in the order the
    rep picked them, and that is the order they appear in the deck. Returns
    the number of slides the deck now carries for this (0 when nothing was
    replaced).

    Replace-and-EXPAND: the first market takes over the existing slide and
    each additional market gets a duplicate inserted directly after it, so a
    three-market proposal carries three profile slides where the master has
    one. Passing an empty list is a no-op and leaves the national Total U.S.
    slide exactly as the master has it -- which is the right default and the
    behaviour every deck had before target markets existed.

    Located by condition_key rather than by position. The slide carries
    `key: market_profile` in its own speaker notes and retained slides keep
    their notes through build_presentation, so this survives the deck being
    re-cut; nothing here depends on it being slide 5.
    """
    if not image_paths:
        return 0

    deck_map = slide_map.build_slide_map_from_prs(prs)
    targets = sorted(n for n, key in deck_map.items() if key == MARKET_PROFILE_KEY)
    if not targets:
        return 0

    base_number = targets[0]
    base_slide = prs.slides[base_number - 1]

    picture = _full_bleed_picture(base_slide)
    if picture is None:
        return 0
    _swap_picture_image(picture, image_paths[0])

    # Each extra market is inserted immediately after the one before it, so
    # picker order survives into the deck. insert_at is a 0-based index into
    # sldIdLst, and base_number is 1-based, so the first insert lands at
    # base_number (i.e. directly after the base slide).
    for offset, path in enumerate(image_paths[1:], start=0):
        clone = duplicate_slide(prs, base_slide, insert_at=base_number + offset)
        clone_picture = _full_bleed_picture(clone)
        if clone_picture is not None:
            _swap_picture_image(clone_picture, path)
        # duplicate_slide deliberately doesn't carry the notes part across (a
        # notes slide belongs to exactly one slide), so a clone arrives with
        # no `key:` label and re-scanning the built deck reports it as an
        # unresolved slide that inherited its neighbour's section. Re-apply
        # the label -- these ARE market profile slides, and a deck that can't
        # describe itself puts noise in the one warning channel that is
        # supposed to mean something.
        clone.notes_slide.notes_text_frame.text = f"key: {MARKET_PROFILE_KEY}"

    return len(image_paths)


def _swap_picture_image(picture, image_path):
    """Point one picture shape at a new image, keeping position and size."""
    image_part, rId = picture.part.get_or_add_image_part(image_path)
    picture._element.blipFill.blip.rEmbed = rId


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
        # Source presentations are retained for the cache's lifetime. Without
        # this their parts are collected between calls and CPython recycles
        # the addresses, which is fatal when anything is keyed on identity.
        self.sources = []
        for part in dst_prs.part.package.iter_parts():
            partname = str(part.partname)
            self.used.add(partname)
            if _is_shareable_media(part):
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


def _is_shareable_media(part):
    """Can two slides safely point at one copy of this part?

    Only /ppt/media/ -- images, video, the vector logos these decks are full
    of. A chart's embedded workbook lives in /ppt/embeddings/ and must be
    owned by exactly one chart; colour and style parts live in /ppt/charts/
    and are cheap enough not to be worth the risk.
    """
    return str(part.partname).startswith("/ppt/media/")


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
    # Keyed on (package, partname), NOT id(src_part).
    #
    # copy_slide_into opens its source Presentation locally, so every part of
    # it becomes garbage the moment the call returns -- and CPython reuses
    # those addresses. A later source deck's style part could then land on
    # the same id() as a freed workbook, and the cache would hand back the
    # workbook: chart4 came out with two embedded workbooks and no style
    # part. It needed four case studies before an address happened to be
    # recycled, every relationship still resolved, and PowerPoint refused the
    # file with no more detail than "could not open". The cache also keeps
    # each source package alive (see ImportCache.sources) so the identity
    # can't churn underneath it.
    identity = (id(src_part.package), str(src_part.partname))
    if identity in cache.parts:
        return cache.parts[identity]

    blob = src_part.blob
    child_rels = [(rId, rel) for rId, rel in src_part.rels.items()
                  if rel.reltype != RT.NOTES_SLIDE]

    # Deduplication is deliberately limited to **media**, and that limit is
    # load-bearing rather than cautious.
    #
    # It used to cover any leaf part -- anything without relationships of its
    # own -- which swept in a chart's embedded workbook and its colour and
    # style parts. Two charts with identical data then shared one workbook,
    # and PowerPoint refuses to open a deck where that happens: a chart's
    # externalData is the copy it edits, so two charts pointing at one part
    # is a contradiction, not an economy. python-pptx opens such a file
    # happily, every relationship resolves, no partname is duplicated -- and
    # PowerPoint says only "could not open the file". It cost a real case
    # study, and it reproduced by importing a single chart slide *twice*.
    #
    # Images are genuinely shareable and are the reason the cache exists at
    # all: every case study deck ships its own copy of the same PREMION
    # logos. Everything else now gets its own part, which is what the
    # "fresh partname per imported part" rule was supposed to mean.
    digest = None
    if not child_rels and _is_shareable_media(src_part):
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
    if cache is None:
        cache = ImportCache(dst_prs)
    # Held so the source package outlives this call -- see ImportCache.
    cache.sources.append(src_prs)

    free_partname = _next_free_slide_partname(dst_prs)
    new_slide = dst_prs.slides.add_slide(_blank_layout(dst_prs))
    new_slide.part.partname = free_partname

    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)


    rid_map = {}
    for rId, rel in src_slide.part.rels.items():
        # The layout is handled above; a notes slide belongs to exactly one
        # slide, and sharing one between two is malformed.
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

    # The slide is now sitting on the destination's Blank layout, so anything
    # it inherited rather than stated would resolve against the wrong master
    # and theme. Write those inherited properties onto the shapes while the
    # source deck is still open to be read from.
    def relate(rel):
        """Bring one of the layout's or master's relationships onto this slide."""
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.SLIDE_MASTER, RT.NOTES_SLIDE, RT.THEME):
            return None
        if rel.is_external:
            return new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        return new_slide.part.rels.get_or_add(
            rel.reltype, _import_part(rel.target_part, dst_prs.part.package, cache))

    slide_inheritance.flatten_inherited_formatting(new_slide, src_slide, relate)
    _fit_flattened_text(new_slide)

    if position is not None:
        sldIdLst = dst_prs.slides._sldIdLst
        new_sldId = list(sldIdLst)[-1]  # add_slide appends
        sldIdLst.remove(new_sldId)
        sldIdLst.insert(position, new_sldId)

    return new_slide


# No margin. A 0.92 factor was tried to make up for PowerPoint's autofit
# also reducing line spacing, and it shrank the case study titles below the
# size the source deck renders them at while still not saving the one
# paragraph it was aimed at -- trading a visible defect for a different one.
# What survives is a genuine residual: a paragraph the source's live autofit
# shrinks can still come out one line tall here, and the render check below
# reports it rather than the deck hiding it.
_FLATTENED_FIT_MARGIN = 1.0


def _fit_flattened_text(slide):
    """Do the shrinking PowerPoint will no longer do for these shapes.

    A placeholder set to <a:normAutofit> is shrunk by PowerPoint to stay in
    its box. Flattening strips <p:ph>, so the shape stops being a placeholder
    and stops being autofitted -- and PowerPoint only recalculates autofit on
    *edit* anyway, so a generated deck that is merely opened keeps whatever
    size is stored. The case study's client-challenge paragraph came out one
    line taller than its box and ran into the stats chips beneath it.

    So the shrink is computed here, with the same measured fit the Campaign
    Specs block uses, and only where the source actually asked for it --
    text the source deck intended to overflow is left overflowing.
    """
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if not shape.has_text_frame or shape.height is None or shape.width is None:
            continue
        bodyPr = shape.text_frame._txBody.find(qn("a:bodyPr"))
        if bodyPr is None or bodyPr.find(qn("a:normAutofit")) is None:
            continue
        insets = (_inset(bodyPr, "tIns", _DEFAULT_CELL_INSET)
                  + _inset(bodyPr, "bIns", _DEFAULT_CELL_INSET))
        width = shape.width - (_inset(bodyPr, "lIns", _DEFAULT_SIDE_INSET)
                               + _inset(bodyPr, "rIns", _DEFAULT_SIDE_INSET))
        # Judged against slightly less than the box really has. PowerPoint's
        # own autofit reduces line spacing as well as type size, so its idea
        # of "fits" is tighter than a pure height estimate -- and this text
        # measured as fitting by a hair while the renderer drew it a line
        # past the bottom, into the stats chips.
        available = int((shape.height - insets) * _FLATTENED_FIT_MARGIN)
        if available > 0 and width > 0:
            fit_text_frame(shape.text_frame, available, width)


def _inset(bodyPr, name, default):
    value = bodyPr.get(name)
    return int(value) if value is not None else int(default)


def append_case_study_images(prs, image_paths, position=None):
    """Insert pre-rendered case study slides as full-bleed pictures.

    The alternative to copying a slide's XML across decks, and immune to
    everything that has gone wrong there: a picture has no relationships to
    remap, no theme or layout to resolve against, no placeholders to inherit
    from, no autofit for PowerPoint to recompute, and no embedded parts to
    collide. What the source deck rendered is what the client sees, exactly.

    The cost is real and permanent: the text is a picture. Nobody can select
    it, search it, or fix a typo in the deck -- see CLAUDE.md.
    """
    inserted = 0
    for path in image_paths:
        partname = _next_free_slide_partname(prs)
        slide = prs.slides.add_slide(_blank_layout(prs))
        slide.part.partname = partname
        for shape in list(slide.shapes):
            shape._element.getparent().remove(shape._element)
        slide.shapes.add_picture(str(path), 0, 0, prs.slide_width, prs.slide_height)
        # The destination master would otherwise draw its footer over the
        # top of a full-bleed image.
        slide._element.set("showMasterSp", "0")
        if position is not None:
            sldIdLst = prs.slides._sldIdLst
            appended = list(sldIdLst)[-1]
            sldIdLst.remove(appended)
            sldIdLst.insert(position + inserted, appended)
        inserted += 1
    return inserted


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
    # Found by the token every plan template carries, not by condition_key.
    # The key was "proposal_template" exactly, which stopped being the only
    # plan slide the moment Total TV shipped: its market variants key
    # `proposal_template_total_tv:dc` and friends, matched nothing, and the
    # whole lookup fell through to appending at the end -- so a Total TV deck
    # put its case studies *after* the media plan, which is the one place
    # they argue for nothing. Any future variant does the same thing to a
    # list of key names, whereas {{PLAN_TITLE}} is what makes a slide the
    # plan slide in the first place: personalize finds it the same way.
    plan = find_slide_with_marker(prs, _placeholder("PLAN_TITLE"))
    if plan is not None:
        found = slide_index(prs, plan)
        if found is not None:
            return found

    assembled = slide_map.build_slide_map_from_prs(prs)
    for n, key in sorted(assembled.items()):
        if key == "proposal_template" or key.startswith(TOTAL_TV_PLAN_PREFIX):
            return n - 1  # 1-based position N -> 0-based index before it
    # Section dividers are dropped from every proposal, so this is a
    # degradation path rather than a live one -- see the docstring.
    for n, key in sorted(assembled.items()):
        if key == slide_map.SECTION_DIVIDER_KEY:
            return n - 1
    return len(assembled)


def append_case_studies(prs, case_studies):
    """Graft every selected case study's slides into the assembled deck.

    `case_studies` is an ordered list of entries that are one of two things:

      {"images": [path, ...]}  -- pre-rendered slides, inserted as full-bleed
                                  pictures. Preferred, and immune to every
                                  cross-deck copying hazard below.
      {"path": ..., "slides": [...]}  -- the .pptx to copy slides out of,
                                  "slides" being 0-based indices (None = all).

    Both routes are handled here rather than at the call sites so the insert
    position and ordering are worked out once; a proposal can mix the two
    freely, since a case study only has images once someone has run
    render_case_study_images.py for it.

    One ImportCache spans the whole run so shared assets -- every one of
    these decks carries the same PREMION logos -- are stored once.
    """
    if not case_studies:
        return 0

    cache = ImportCache(prs)
    position = case_study_insert_index(prs)
    copied = 0

    for case_study in case_studies:
        images = case_study.get("images")
        if images:
            inserted = append_case_study_images(prs, images, position)
            position += inserted
            copied += inserted
            continue
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

    # The vertical's own branded attribution product (Polk Signals for auto,
    # Arrivalist for travel). Its slide rides into the deck on the vertical's
    # own condition_key, so it can only be included or excluded by finding it
    # -- hence the marker, and hence a sweep here rather than a key in
    # resolve_active_keys.
    #
    # `None` means the proposal predates the checkbox: leave both the slide and
    # the generic sales slide exactly as they were, so rebuilding an older
    # proposal reproduces what the client was actually sent.
    va_spec = vertical_attribution_spec(vertical)
    va_choice = selections.get("vertical_attribution")
    if va_spec is not None and va_choice is not None:
        marker = va_spec["marker"].upper()
        va_slides = {n for n in keep_numbers
                     if marker in slide_map.extract_slide_text(prs.slides[n - 1]).upper()}
        if not va_choice:
            keep_numbers -= va_slides
        elif va_slides and va_spec["replaces_sales_slide"]:
            # Only when the slide is actually there -- with the vertical's own
            # slides switched off there is nothing to replace it with, and
            # dropping the generic one would leave the deck with no sales
            # attribution slide at all.
            keep_numbers = {n for n in keep_numbers
                            if deck_slide_map.get(n) != "sales_attribution"}

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
    shape = _find_shape_with_token(shapes, token)
    return shape.text_frame if shape is not None else None


def _find_shape_with_token(shapes, token):
    """The shape carrying a token, rather than just its text frame.

    Sizing needs the shape: a frame knows its insets but not how wide the box
    it sits in is. Callers that want to measure a token's slot have to find it
    *before* the token is consumed, the same way the broadcast summary block
    is captured ahead of fill_bullet_list_in_slide.
    """
    ph = _placeholder(token)
    for shape in slide_map.iter_all_shapes(shapes):
        if shape.has_text_frame and ph in shape.text_frame.text:
            return shape
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


def strip_bullets(text_frame):
    """Remove the bullet glyph and its hanging indent from every paragraph.

    For the compressed Included band, which is one joined line rather than a
    list: a single bullet in front of "Dedicated Account Management Team,
    Monthly Reporting Calls & Optimizations, ..." reads as a formatting
    mistake, and this is a slide clients sign. The uncompressed band keeps
    its bullets, because there it really is a list.

    `<a:buNone/>` has to be stated rather than the existing bullet merely
    deleted -- the paragraph inherits one from the layout, so removing
    `buChar` alone lets the inherited bullet through. It also has to be
    inserted in schema order (bullet properties come before `tabLst`,
    `defRPr` and `extLst`); appending it blindly produces a pPr PowerPoint
    rejects, which is the sort of file that opens nowhere and reports
    nothing useful.
    """
    for para in text_frame.paragraphs:
        pPr = para._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buNone", "a:buBlip"):
            for element in pPr.findall(qn(tag)):
                pPr.remove(element)
        pPr.insert_element_before(
            pPr.makeelement(qn("a:buNone"), {}), "a:tabLst", "a:defRPr", "a:extLst")
        pPr.set("marL", "0")
        pPr.set("indent", "0")


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
            # Unconditionally, not only when the condense pass runs: a small
            # plan never reaches condense, and the inherited paragraph
            # padding is what made even a short table sit too low.
            _tighten_cell_paragraphs(cell)


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
_DEFAULT_SIDE_INSET = Emu(91440)  # 0.1in, the default left/right -- twice that
_MIN_TABLE_FONT_PT = 6
_TABLE_CLEARANCE = Emu(228600)   # 0.25in

# Two different floors, and conflating them is what left long plans at 6pt.
# _MIN_TABLE_FONT_PT is the absolute floor: below it the table cannot fit at
# all and condense_media_plan_table warns. This one is the size below which a
# plan stops being readable across a conference table -- not a failure, but
# the point at which it's worth spending layout elsewhere on the slide to buy
# the type back. Measured on the real deck: the font ladder runs 12 -> 10 ->
# 8 -> 7 -> 6, so 9pt itself never occurs and this reads as "must not go
# below 10". It is deliberately expressed as the readable size rather than as
# a row count, because the row count that reaches it moves with the content:
# changing only the Flight column from "Sep 1 - Nov 30" to "9/1 - 11/30"
# moved a six-row plan from 8pt to 12pt.
_COMFORTABLE_TABLE_FONT_PT = 9

# Slack reserved on top of a row's text. Empirical, and larger than the text
# strictly needs: PowerPoint draws these table rows taller than their content
# accounts for -- a 9pt single-line row in the schedule table comes back 0.3in
# tall however little is in it, where the same row in a table built from
# scratch honours 0.15in. Nothing in the cell XML explains the difference, so
# the allowance stays rather than being tuned down to a number that only looks
# right on paper. It is deliberately NOT used to decide whether a table fits
# -- see _row_text_height -- because reserving generously and refusing
# generously are different things, and the second one turns into a warning
# about a plan PowerPoint would have drawn comfortably.
_ROW_CUSHION = 2 * _DEFAULT_CELL_INSET   # 0.1in


def _floor_below(slide, top, left, right, exclude=()):
    """The top of the highest shape sitting below `top` and overlapping the
    horizontal span [left, right) -- i.e. the first thing something filling
    that span down to `top` would collide with. The bounds-based core
    `_content_floor` (a table's own ceiling) and `map_slide_region` (the
    targeting map's footprint, which has no table of its own to measure
    from) both derive their floor through here, rather than each walking
    the shape tree its own way.

    `exclude` is elements to skip by identity -- comparing the underlying
    XML element, never the shape wrapper itself: python-pptx builds a fresh
    proxy object every time a shape tree is walked, so `is` never matches a
    shape obtained from an earlier call.
    """
    excluded_elements = {getattr(s, "_element", s) for s in exclude}
    floor = None
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if (shape._element in excluded_elements
                or shape.top is None or shape.left is None):
            continue
        if shape.top <= top:
            continue  # above the region, or the region's own anchor
        # only things that actually sit under the region's own span matter
        if shape.left + (shape.width or 0) <= left or shape.left >= right:
            continue
        floor = shape.top if floor is None else min(floor, shape.top)
    return floor


def _content_floor(slide, table_shape, ignore=None):
    """The top of the highest shape sitting below the table and overlapping it
    horizontally -- i.e. the first thing the table would collide with.

    Derived from the slide rather than by shape name. The previous version
    looked up "Picture 12" (the "One Solution" graphic at 4.70in) and missed
    that the "Included with Campaign:" heading starts higher at 4.63in, so
    the table was sized against a ceiling 0.07in below the one that actually
    mattered -- and any deck that renamed or re-layered that picture lost the
    check entirely.
    """
    exclude = [table_shape] + ([ignore] if ignore is not None else [])
    return _floor_below(slide, table_shape.top, table_shape.left,
                        table_shape.left + table_shape.width, exclude=exclude)


def _max_font_for_row(row_height_emu):
    """The largest whole point size whose line box fits in a row of this
    height, so the row height we ask for is one PowerPoint can honour."""
    usable = row_height_emu - _ROW_CUSHION
    return max(_MIN_TABLE_FONT_PT, int(usable / (_LINE_SPACING * _EMU_PER_POINT)))


def min_table_row_height():
    """The shortest a row can be and still hold readable text."""
    return Emu(int(_MIN_TABLE_FONT_PT * _LINE_SPACING * _EMU_PER_POINT) + _ROW_CUSHION)


def table_row_capacity(slide, header_rows=1, totals_rows=1, ignore=None):
    """How many data rows fit above whatever sits below the table.

    Computed from the slide's own geometry so a caller can decide to
    paginate *before* filling anything -- once the rows are cloned in it's
    too late to find out they don't fit.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    floor = _content_floor(slide, table_shape, ignore=ignore)
    if floor is None:
        return None
    available = floor - _TABLE_CLEARANCE - table_shape.top
    total = int(available / min_table_row_height())
    return max(0, total - header_rows - totals_rows)


@functools.lru_cache(maxsize=8)
def _theme_typefaces(theme_xml):
    """(major, minor) latin typefaces from a theme part's XML."""
    major = re.search(r"<a:majorFont>\s*<a:latin typeface=\"([^\"]*)\"", theme_xml)
    minor = re.search(r"<a:minorFont>\s*<a:latin typeface=\"([^\"]*)\"", theme_xml)
    return (major.group(1) if major else None, minor.group(1) if minor else None)


def _resolve_typeface(run_name, table):
    """The typeface a run actually renders in.

    Runs mostly carry the theme placeholders +mj-lt / +mn-lt rather than a
    name, so measuring needs the theme resolved -- Calibri Light is
    measurably narrower than Calibri and the header rows use it.
    """
    major = minor = None
    entry = _THEME_CACHE.get(id(table))
    # Confirm the identity rather than trusting the address: a recycled id()
    # belongs to a different table, and taking its fonts is how a deck gets
    # measured in a typeface it never uses.
    if entry is not None and entry[0] is table._tbl:
        major, minor = entry[1]
    if run_name in (None, "", "+mn-lt"):
        return minor or "Calibri"
    if run_name == "+mj-lt":
        return major or "Calibri Light"
    return run_name


_THEME_CACHE = {}


def declared_fonts(prs):
    """The typefaces the deck itself says it uses.

    Read from presentation.xml's <p:embeddedFontLst>, so this is deck-driven
    rather than tied to a font name written down in here: when the brand font
    changes, a re-saved deck declares the new one and every check that reads
    this follows it with no code change. Empty for a deck saved without
    "Embed fonts in the file".

    Note the embedded data itself is NOT usable for measurement -- PowerPoint
    writes each face as EOT 2.2 with the TTCOMPRESSED flag (MicroType Express),
    which PIL cannot open and which has no practical Python decompressor. So
    the list tells us what the deck NEEDS; whether we can measure it is still
    a question about the machine.
    """
    try:
        xml = prs.part.blob.decode("utf-8", "ignore")
    except Exception:                                            # noqa: BLE001
        return []
    block = re.search(r"<p:embeddedFontLst>.*?</p:embeddedFontLst>", xml, re.S)
    if block is None:
        return []
    return sorted(set(re.findall(r'typeface="([^"]+)"', block.group(0))))


def _theme_part_for_slide(slide):
    """The theme a slide's runs actually resolve against.

    Its OWN master's, reached through its layout -- not the presentation's
    first master. The master deck carries 21 masters and master 0 is the only
    one whose minor font is Calibri; every content slide uses one of the
    Proxima Nova masters. Resolving against master 0 therefore measured the
    entire deck in a font it never draws.
    """
    try:
        return slide.part.slide_layout.slide_master.part.part_related_by(RT.THEME)
    except Exception:                                            # noqa: BLE001
        return None


def register_theme_for_table(table, prs, slide=None):
    """Tell the measurer which theme a table's runs resolve against.

    Pass the slide the table is on. Without one this falls back to the first
    master, which is right only for a single-master deck -- see
    _theme_part_for_slide for why that fallback under-measures this one.

    The cache retains the table's XML element alongside the value. Keying on
    id() alone is the bug this project has already hit once (ImportCache, on
    freed part addresses being recycled): _THEME_CACHE never evicts, so a
    collected table's address could be reused by a later one and hand it
    another deck's fonts.
    """
    theme = _theme_part_for_slide(slide) if slide is not None else None
    if theme is None:
        try:
            theme = prs.slide_masters[0].part.part_related_by(RT.THEME)
        except Exception:                                        # noqa: BLE001
            _THEME_CACHE.pop(id(table), None)
            return
    _THEME_CACHE[id(table)] = (
        table._tbl, _theme_typefaces(theme.blob.decode("utf-8", "ignore")))


def _wrapped_lines(text, column_width_emu, font_pt, typeface=None,
                   side_insets_emu=2 * _DEFAULT_SIDE_INSET, bold=False):
    """How many lines this text takes in a column that wide.

    Measured against the real font rather than estimated from an average
    character width. The estimate had to err high -- guess low and the table
    renders on top of the block beneath it -- so it reserved height for two
    lines where one was drawn, and the font search shrank the type further
    than it had to. It also wasn't reliably conservative: at 12pt it
    predicted two lines for a string that measures three.
    """
    text = (text or "").strip()
    if not text:
        return 1
    usable_pt = max(1.0, (column_width_emu - side_insets_emu) / _EMU_PER_POINT)
    return text_metrics.wrapped_lines(text, usable_pt, typeface or "Calibri", font_pt, bold)


def _side_insets(cell):
    """How much of a column's width the cell's own padding takes.

    Not the same constant as the vertical one, and getting them confused
    over-states every column by about 11 points: PowerPoint's default side
    inset is 0.1in per side against 0.05in top and bottom, and only the
    vertical pair is zeroed here (they cost row height; the horizontal pair
    is the padding that keeps text off the cell border). A cell returns None
    for an inset it inherits rather than sets.
    """
    left = cell.margin_left
    right = cell.margin_right
    return ((_DEFAULT_SIDE_INSET if left is None else left)
            + (_DEFAULT_SIDE_INSET if right is None else right))


def _lines_in_row(table, row_index, font_pt):
    """Rendered lines in this row's tallest cell, at this font size."""
    most = 1
    row = table.rows[row_index]
    for column, cell in enumerate(row.cells):
        if column >= len(table.columns):
            continue
        lines = 0
        for para in cell.text_frame.paragraphs:
            runs = list(para.runs)
            typeface = _resolve_typeface(
                next((r.font.name for r in runs if r.font.name), None), table)
            lines += _wrapped_lines("".join(r.text for r in runs),
                                    table.columns[column].width, font_pt, typeface,
                                    _side_insets(cell),
                                    any(r.font.bold for r in runs))
        most = max(most, max(1, lines))
    return most


def _max_lines_in_rows(table, header_rows, font_pt):
    """The tallest data row, in rendered lines, at this font size."""
    most = 1
    for index, row in enumerate(table.rows):
        if index < header_rows:
            continue
        for column, cell in enumerate(row.cells):
            if column >= len(table.columns):
                continue
            lines = 0
            for para in cell.text_frame.paragraphs:
                runs = list(para.runs)
                typeface = _resolve_typeface(runs[0].font.name if runs else None, table)
                lines += _wrapped_lines("".join(r.text for r in runs),
                                        table.columns[column].width, font_pt, typeface,
                                        _side_insets(cell))
            lines = max(1, lines)
            most = max(most, lines)
    return most


# Text columns a media plan table can rebalance, and the least each may keep.
# Flight is the donor: dates are short and a fixed width, so the space it
# doesn't need is exactly what Targeting does.
_MIN_COLUMN_WIDTH = Emu(int(0.85 * 914400))


def balance_text_columns(table, header_rows=1, protect_last=2):
    """Redistribute width between text columns according to what's in them.

    The template's proportions are a guess made before anyone knew what the
    cells would say. Flight always reads "Jan 2027 - Mar 2027" and needs
    about an inch; Targeting carries a sentence and wraps to three lines in
    the space left over. Nothing here changes the table's total width -- it
    only moves slack from the columns that have it to the ones that don't.
    """
    text_columns = list(range(0, len(table.columns) - protect_last))
    if len(text_columns) < 2:
        return

    demand = {}
    for column in text_columns:
        longest = 0
        for index, row in enumerate(table.rows):
            if index < header_rows:
                continue
            longest = max(longest, len(" ".join(row.cells[column].text.split())))
        demand[column] = max(longest, 4)

    total = sum(table.columns[c].width for c in text_columns)
    floor = _MIN_COLUMN_WIDTH * len(text_columns)
    if total <= floor:
        return
    spare = total - floor
    weight = sum(demand.values())
    for column in text_columns:
        table.columns[column].width = Emu(
            int(_MIN_COLUMN_WIDTH + spare * demand[column] / weight))


def _row_height_for(lines, font_pt):
    """Height to reserve for a row showing `lines` lines of `font_pt` text."""
    return Emu(int(lines * font_pt * _LINE_SPACING * _EMU_PER_POINT) + _ROW_CUSHION)


def _row_text_height(lines, font_pt):
    """Height the text alone occupies -- the cushion left off.

    What "can this table fit at all" has to be judged against. Judging it
    against the reserved height instead declared a 12-line plan too tall for
    a slide that PowerPoint then drew it on with a quarter inch to spare,
    because twelve rows of cushion is an inch of space that isn't really
    required by anything.
    """
    return Emu(int(lines * font_pt * _LINE_SPACING * _EMU_PER_POINT))


def condense_media_plan_table(slide, num_data_rows, extra_total_rows=0, header_rows=1,
                              extra_total_label=None, template_metrics=None):
    """Shrink cloned media-plan rows (and their font) so the table can't grow
    into the "Included with Campaign" block or the graphic beneath it.

    Only touches anything if the rows at their natural (template) height would
    actually overlap -- small row counts are left exactly as the template has
    them.

    `template_metrics` is an in/out dict that makes a SECOND call correct.
    This function reads the template's own row heights to decide the starting
    font and how much of the slide is fixed overhead, and it overwrites those
    heights -- so calling it twice would take its own first-pass output for
    the template and lock in the smaller font. Pass the same dict both times:
    the first call records the template heights, the second reuses them. It
    also records the font settled on, which is how a caller learns whether
    the plan came out readable (see _COMFORTABLE_TABLE_FONT_PT) without this
    function having to grow a second return value that every existing caller
    would have to unpack.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    floor = _content_floor(slide, table_shape)
    if floor is None:
        return None

    table = table_shape.table
    # Resolve the theme so runs carrying +mj-lt / +mn-lt are measured against
    # the typeface they'll actually render in.
    try:
        register_theme_for_table(
            table, slide.part.package.presentation_part.presentation, slide=slide)
    except Exception:                                            # noqa: BLE001
        pass
    # Trim every cell first, including the header and totals rows -- those
    # are filled by the simple-token pass rather than fill_table_rows, so
    # they keep their surplus paragraphs and would otherwise set the
    # per-row line count for the whole table on their own.
    for row in table.rows:
        for cell in row.cells:
            _drop_surplus_paragraphs(cell)
    data_rows = list(table.rows)[header_rows:-1]
    if not data_rows:
        return None
    # Every header row is fixed overhead, not just the first -- the broadcast
    # schedule table has two (the "Wk" band and the labels beneath it), and
    # counting one left it short by a row's worth of height.
    if template_metrics is None:
        template_metrics = {}
    if "template_heads" not in template_metrics:
        template_metrics["template_heads"] = [table.rows[i].height
                                              for i in range(header_rows)]
        template_metrics["totals_h"] = table.rows[len(table.rows) - 1].height
        template_metrics["original_row_h"] = data_rows[0].height
        # The template's OWN type size, read before this pass overwrites it.
        # It is the real ceiling on the font -- _max_font_for_row derives ~47pt
        # from the template's row height, which is no ceiling at all, and what
        # actually held the plan table at 12pt was _tighten_cell_paragraphs
        # refusing to enlarge a run. That refusal reads the run's current size,
        # so on a second pass it compares against the first pass's output and
        # pins the table to whatever the tightest pass chose -- a plan given
        # more room stayed at 8pt while the search had already worked out it
        # could have 13. Recorded here, it survives both passes.
        sizes = [run.font.size.pt
                 for row in data_rows for cell in row.cells
                 for para in cell.text_frame.paragraphs
                 for run in para.runs if run.font.size]
        template_metrics["template_run_pt"] = max(sizes) if sizes else None
    template_heads = template_metrics["template_heads"]
    totals_h = template_metrics["totals_h"]
    original_row_h = template_metrics["original_row_h"]

    # Clearance the table must leave below itself. Was ~0.05in, which is
    # arithmetically "fits" and visually touching -- a seven-line plan cleared
    # the graphic beneath it by half a millimetre. A quarter inch reads as
    # deliberate space instead, and makes the condense pass engage a little
    # earlier rather than only once something already overlaps.
    margin = _TABLE_CLEARANCE
    available = floor - margin - table_shape.top
    # Measured against what the table will actually DRAW, not against its
    # declared row heights. There's deliberately no early return for a short
    # table: three rows fit on paper and still overlapped, because one
    # Targeting cell wrapped to four lines in a row declared at 0.33in.
    template_font = _max_font_for_row(original_row_h)
    template_run_pt = template_metrics.get("template_run_pt")
    if template_run_pt:
        template_font = min(template_font, int(template_run_pt))
    # No early return. A table that fits is still sized here, because "fits"
    # and "uses the space well" are different things: a three-line plan left
    # at template height sat in the top third of the slide with the rest
    # empty, and the schedule grid did the same. Rows grow into the space
    # available and shrink when there isn't any -- capped at the template's
    # own font so a one-line plan doesn't balloon into a poster.

    # The smallest a row can be and still hold readable text.
    min_row_h = min_table_row_height()
    # A row is as tall as its tallest cell, and a cell with two paragraphs
    # is two lines whatever height the row asks for. Counting them keeps the
    # declared height honest -- the previous model assumed one line per row
    # and under-reported by exactly the surplus paragraphs above.
    # min_row_h stays ONE line at the smallest readable size -- the absolute
    # floor for a row. Scaling it by the wrapped line count was tried and is
    # wrong: it inflates the floor for every table, so a schedule with long
    # programme names could no longer fit at any size and warned instead.
    # Wrapping is handled where it belongs, by reducing the font until the
    # text fits the row (below).

    # What each row needs at a given size, measured row by row. One height for
    # every data row is the thing that was wrong here: a row is as tall as its
    # own tallest cell, so giving a row that wraps to two lines the same share
    # as one that doesn't means PowerPoint grows it past what was reserved --
    # the overlap arriving after the arithmetic said it fit. A "Fairfax,
    # Loudoun and Prince William counties, VA" Geo cell overran its column by
    # a tenth of a point and cost the slide a row.
    def needs(size_pt, index):
        return max(min_row_h, _row_height_for(_lines_in_row(table, index, size_pt), size_pt))

    def extra_totals_height(size_pt, foot):
        """Height to reserve for the full-flight row, which doesn't exist yet.

        It's cloned from the totals row after this runs, so reserving the
        totals row's own height assumes the clone says something the same
        length -- and it doesn't: "Full Flight Total (3 months)" wraps where
        "Monthly Totals" doesn't, and the clone then grew past the space set
        aside for it. Measuring the label the caller is going to use closes
        that gap; without one, the old assumption is still the best guess.
        """
        if not extra_total_label:
            return foot
        cell = table.rows[len(table.rows) - 1].cells[0]
        lines = _wrapped_lines(extra_total_label, table.columns[0].width, size_pt,
                               _resolve_typeface(None, table), _side_insets(cell))
        return max(foot, _row_height_for(lines, size_pt))

    def layout(size_pt):
        """(fixed overhead, per-data-row heights) at this size.

        Everything that isn't a data row is overhead, and at high row counts
        the data can't fit inside what's left of it -- so the header and
        totals rows give up their template height and fall back to what their
        own text needs. They keep it while there's room, because a table
        whose header is the same height as its rows reads as a grid rather
        than a plan. An earlier version clamped them at half their template
        height, which left 12 rows plus a full-flight row overflowing by
        0.21in; a later one shrank them on a one-line-per-row estimate, which
        is the same mistake in the other direction.
        """
        rows = [needs(size_pt, i) for i in range(header_rows, len(table.rows) - 1)]
        head_needs = [needs(size_pt, i) for i in range(header_rows)]
        foot_need = needs(size_pt, len(table.rows) - 1)

        # While there's room, keep the template's own proportions -- but
        # never below what the header's text needs. Trusting the template
        # height outright is what left a "Wk / Total / Adults 25-64 (000)"
        # header declared at one line and drawn at two: the header wraps for
        # exactly the same reasons a data row does, and it isn't exempt from
        # being measured just because the template had an opinion about it.
        roomy_heads = [max(template_heads[i], head_needs[i]) for i in range(header_rows)]
        roomy_foot = max(totals_h, foot_need)
        roomy = (sum(roomy_heads) + roomy_foot
                 + extra_totals_height(size_pt, roomy_foot) * extra_total_rows)
        if roomy + sum(rows) <= available:
            return roomy, rows, (roomy_heads, roomy_foot)

        tight = (sum(head_needs) + foot_need
                 + extra_totals_height(size_pt, foot_need) * extra_total_rows)
        return tight, rows, (head_needs, foot_need)

    # Never larger than the template intended, however much room there is.
    # Reduced until every row's text actually fits; smaller type wraps to
    # fewer lines, so this converges.
    font_pt = template_font
    while font_pt > _MIN_TABLE_FONT_PT:
        fixed, needed, tightened = layout(font_pt)
        if fixed + sum(needed) <= available:
            break
        font_pt -= 1
    fixed, needed, tightened = layout(font_pt)
    template_metrics["font_pt"] = font_pt

    heads, foot = tightened
    for index in range(header_rows):
        table.rows[index].height = heads[index]
    # Only the template's own totals row exists yet; the full-flight row
    # is cloned from it afterwards and inherits whatever we set here.
    table.rows[len(table.rows) - 1].height = foot

    budget = available - fixed
    # Past this point the table physically cannot fit: every row is already at
    # the minimum readable height and the text still doesn't go in. Say so
    # rather than shipping an overlap -- and still clamp the rows to the space
    # there is, so an unavoidable overflow is as small as it can be instead of
    # as large as the text wants.
    if sum(needed) > budget:
        share = max(_row_text_height(1, font_pt), Emu(int(budget / len(needed))))
        needed = [min(height, share) for height in needed]
    bare = sum(_row_text_height(_lines_in_row(table, index, font_pt), font_pt)
               for index in range(header_rows, len(table.rows) - 1))
    if bare > budget:
        fits = max(0, int(budget / _row_text_height(1, font_pt)))
        # Deliberately does NOT suggest splitting the plan across options.
        # Options are alternatives the client chooses between, not pages of one
        # plan: splitting a plan in half presents it as a choice between the
        # halves, and each option's totals row would then show half the
        # campaign. That advice shipped here for a while and would have
        # produced a misleading deck for anyone who followed it.
        overflow_warning = (
            f"The media plan has {num_data_rows} lines, which is more than the slide can "
            f"hold without running into the \"Included with Campaign\" block "
            f"(about {fits} lines is the limit, fewer if any of them wrap). "
            f"Combine lines that share a tactic, or shorten the Targeting text.")
    else:
        overflow_warning = None

    # Rows take the space that's there: whatever is left over after every row
    # has what it needs is handed out evenly. Capping their HEIGHT at what the
    # template font needs left a three-line plan sitting in the top third of
    # the slide in 7pt type with room going spare -- the cap belongs on the
    # font, not the height, so a short table reads at a comfortable size and
    # a long one still shrinks.
    slack = budget - sum(needed)
    share = Emu(int(slack / len(needed))) if slack > 0 else Emu(0)
    for row, height in zip(data_rows, needed):
        row.height = Emu(int(height) + int(share))
    for row in table.rows:
        for cell in row.cells:
            cell.margin_top = Emu(0)
            cell.margin_bottom = Emu(0)
            # force: the size may need to go UP on a second pass, and the
            # shrink-only rule can't tell "the template said 12" from "the
            # last pass chose 8". font_pt is capped at the template's own
            # size above, so this can never enlarge past the template.
            _tighten_cell_paragraphs(cell, font_pt, force=True)
    return overflow_warning


def condense_avails_table(slide, num_data_rows, header_rows=1, template_metrics=None):
    """Shrink the avails table's rows (and their font) so it can't run past
    whatever sits below it, the same problem condense_media_plan_table
    solves for the media plan -- but for a table with no totals/footer row
    to treat specially, so the layout here is the plain version of that
    one's. A real proposal (many audience/geo combinations, real client)
    rendered this table to 12.92in on a 7.5in slide with nothing here to
    stop it; the media plan table has always had this pass and the avails
    table never did.

    `template_metrics` follows the same in/out contract as
    condense_media_plan_table -- pass the same dict across repeated calls so
    a second pass reuses the template's own heights instead of the first
    pass's already-shrunk output.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    floor = _content_floor(slide, table_shape)
    if floor is None:
        return None

    table = table_shape.table
    try:
        register_theme_for_table(
            table, slide.part.package.presentation_part.presentation, slide=slide)
    except Exception:                                            # noqa: BLE001
        pass
    for row in table.rows:
        for cell in row.cells:
            _drop_surplus_paragraphs(cell)
    data_rows = list(table.rows)[header_rows:]
    if not data_rows:
        return None

    if template_metrics is None:
        template_metrics = {}
    if "template_heads" not in template_metrics:
        template_metrics["template_heads"] = [table.rows[i].height for i in range(header_rows)]
        template_metrics["original_row_h"] = data_rows[0].height
        sizes = [run.font.size.pt
                 for row in data_rows for cell in row.cells
                 for para in cell.text_frame.paragraphs
                 for run in para.runs if run.font.size]
        template_metrics["template_run_pt"] = max(sizes) if sizes else None
    template_heads = template_metrics["template_heads"]
    original_row_h = template_metrics["original_row_h"]

    margin = _TABLE_CLEARANCE
    available = floor - margin - table_shape.top

    template_font = _max_font_for_row(original_row_h)
    template_run_pt = template_metrics.get("template_run_pt")
    if template_run_pt:
        template_font = min(template_font, int(template_run_pt))

    min_row_h = min_table_row_height()

    def needs(size_pt, index):
        return max(min_row_h, _row_height_for(_lines_in_row(table, index, size_pt), size_pt))

    def layout(size_pt):
        rows = [needs(size_pt, i) for i in range(header_rows, len(table.rows))]
        head_needs = [needs(size_pt, i) for i in range(header_rows)]
        roomy_heads = [max(template_heads[i], head_needs[i]) for i in range(header_rows)]
        roomy = sum(roomy_heads)
        if roomy + sum(rows) <= available:
            return roomy, rows, roomy_heads
        return sum(head_needs), rows, head_needs

    font_pt = template_font
    while font_pt > _MIN_TABLE_FONT_PT:
        fixed, needed, tightened = layout(font_pt)
        if fixed + sum(needed) <= available:
            break
        font_pt -= 1
    fixed, needed, tightened = layout(font_pt)
    template_metrics["font_pt"] = font_pt

    heads = tightened
    for index in range(header_rows):
        table.rows[index].height = heads[index]

    budget = available - fixed
    # Past this point the table physically cannot fit -- same graceful
    # degradation as the media plan: clamp to what's there instead of
    # shipping the overlap, and say so.
    if sum(needed) > budget:
        share = max(_row_text_height(1, font_pt), Emu(int(budget / len(needed))))
        needed = [min(height, share) for height in needed]
    bare = sum(_row_text_height(_lines_in_row(table, index, font_pt), font_pt)
               for index in range(header_rows, len(table.rows)))
    if bare > budget:
        fits = max(0, int(budget / _row_text_height(1, font_pt)))
        overflow_warning = (
            f"The avails table has {num_data_rows} rows, which is more than the slide can "
            f"hold without running past the bottom of the slide "
            f"(about {fits} rows is the limit, fewer if any wrap). "
            f"Combine rows that share an audience and geography, or shorten the Geo text.")
    else:
        overflow_warning = None

    slack = budget - sum(needed)
    share = Emu(int(slack / len(needed))) if slack > 0 else Emu(0)
    for row, height in zip(data_rows, needed):
        row.height = Emu(int(height) + int(share))
    for row in table.rows:
        for cell in row.cells:
            cell.margin_top = Emu(0)
            cell.margin_bottom = Emu(0)
            _tighten_cell_paragraphs(cell, font_pt, force=True)
    return overflow_warning


def _drop_surplus_paragraphs(cell):
    """Remove empty paragraphs beyond the cell's content.

    Blanking a run isn't enough: an empty paragraph still renders a full
    line of height, so a Geo cell carrying three paragraphs with two empty
    occupied three lines' worth of row. That is the gap between the declared
    table height -- which python-pptx computes from row heights -- and what
    PowerPoint actually renders, and it's why the floor check passed at
    3.74in against a 4.63in floor while the real table ran past it.

    The first paragraph is always kept, empty or not, so a genuinely blank
    cell keeps its shape.
    """
    body = cell.text_frame._txBody
    paragraphs = body.findall(qn("a:p"))
    keep_to = 0
    for index, para in enumerate(paragraphs):
        if "".join(node.text or "" for node in para.iter(qn("a:t"))).strip():
            keep_to = index
    for para in paragraphs[keep_to + 1:]:
        body.remove(para)


def _tighten_cell_paragraphs(cell, font_pt=None, force=False):
    """Remove the vertical padding a table cell's paragraphs carry.

    Setting a small row height achieves nothing on its own: PowerPoint grows
    a row to fit its content, and a paragraph's space_before/space_after and
    line spacing are part of that content's height. The template's cells
    inherit both from the theme, which is why generated rows came out far
    taller than the height they were given and pushed the table into the
    graphic below it.
    """
    _drop_surplus_paragraphs(cell)
    for para in cell.text_frame.paragraphs:
        para.space_before = Pt(0)
        para.space_after = Pt(0)
        para.line_spacing = 1.0
        for run in para.runs:
            if font_pt is not None and (force or run.font.size is None
                                        or run.font.size.pt > font_pt):
                run.font.size = Pt(font_pt)
        # The paragraph's end mark carries its own size, and PowerPoint sizes
        # the line by the largest thing in it -- the end mark included. A cell
        # whose runs were shrunk to 12pt but whose end mark still said 14pt
        # drew a taller line than anything visible in it justified, which is
        # a row growing past its reserved height for text that isn't there.
        end = para._p.find(qn("a:endParaRPr"))
        if end is not None and font_pt is not None:
            size = end.get("sz")
            if force or size is None or int(size) > int(font_pt * 100):
                end.set("sz", str(int(font_pt * 100)))


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

# A week column must hold a header like "10/05"; a donor column must stay
# wide enough to be worth having.
MIN_WEEK_COL_WIDTH = Emu(329184)   # 0.36in
MIN_DONOR_WIDTH = Emu(548640)      # 0.60in

# What the shipped template yields, for the UI's advisory page count only.
SCHEDULE_ROWS_CONTINUATION = 20
SCHEDULE_ROWS_FINAL_PAGE = 11


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


def set_week_column_count(table, count, week_start=None):
    """Grow or shrink the week columns to `count`, keeping the table's total
    width -- the surrounding columns are fixed, so the week block absorbs the
    difference and every column stays legible rather than the table growing
    off the slide."""
    week_start = WEEK_COL_START if week_start is None else week_start
    current = len(_grid_cols(table)) - week_start - TRAILING_COLS
    # python-pptx's column collection doesn't slice.
    week_block = sum(table.columns[i].width
                     for i in range(week_start, week_start + current))

    while current < count:
        clone_table_column(table, week_start)
        current += 1
    while current > count:
        remove_table_column(table, week_start)
        current -= 1

    if count:
        # A week column has to hold a header like "10/05", which needs about
        # 0.36in at the sizes these tables use. Dividing the template's week
        # block by ten gives 0.19in and the headers wrap or clip -- so when
        # the block is too narrow, the shortfall is taken from the widest
        # text columns (Program Name, then Time), which have slack a date
        # column doesn't.
        each = int(week_block / count)
        if each < MIN_WEEK_COL_WIDTH:
            needed = MIN_WEEK_COL_WIDTH * count - week_block
            # Days first (often empty -- a Campaign Schedule Report folds the
            # day into Day/Time), then Time, and Program Name last: it holds
            # the longest strings and is the one column a reader actually
            # needs to read.
            for donor in ([2, 1, 3] if week_start == WEEK_COL_START else [1, 2]):
                if needed <= 0:
                    break
                available = table.columns[donor].width - MIN_DONOR_WIDTH
                if available <= 0:
                    continue
                take = min(available, needed)
                table.columns[donor].width = Emu(int(table.columns[donor].width - take))
                needed -= take
            week_block = week_block + (MIN_WEEK_COL_WIDTH * count - week_block - max(0, needed))
            each = max(int(week_block / count), int(MIN_WEEK_COL_WIDTH * 0.8))
        for index in range(week_start, week_start + count):
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


# Acronyms that must survive title-casing a shouty program name.
_PROGRAM_ACRONYMS = {"NFL", "NBA", "MLB", "NHL", "NCAA", "PGA", "WNBA", "ROS", "TV",
                     "CTV", "OTT", "AM", "PM", "US", "USA", "SEC", "ACC", "MNF", "SNF",
                     "TNF", "NY", "LA", "DC", "SF", "KC", "TB", "NE"}
# Trailing codes Wide Orbit appends for its own routing -- "(RIO)", "(SYS)".
# Only a trailing all-caps parenthetical is stripped; a real parenthetical
# like "(Sunday)" stays.
_TRAILING_CODE = re.compile(r"\s*\([A-Z0-9][A-Z0-9 /&.-]*\)\s*$")


def clean_program_name(name):
    """Make a Wide Orbit program name presentable without losing it.

    Two things only: drop the trailing station/system code, and title-case a
    name that arrived shouting. A name that already has mixed case was
    written by a person and is left exactly as it is -- title-casing
    "Morning/Daytime ROS" would produce "Morning/Daytime Ros", which is
    worse than the problem.
    """
    text = _TRAILING_CODE.sub("", (name or "").strip())
    letters = [c for c in text if c.isalpha()]
    if not letters or text != text.upper():
        return text
    words = []
    for word in text.split():
        stripped = word.strip(":,.")
        if stripped in _PROGRAM_ACRONYMS:
            words.append(word)
        elif word == "@":
            words.append(word)
        else:
            words.append(word.title())
    return " ".join(words)


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


def estimate_schedule_slide_count(schedule, breakout="full_flight", detailed=True):
    """How many slides the current settings would produce.

    Advisory only, for the panel's "switch to totals only" hint -- the real
    build measures the template rather than trusting these figures. The
    capacities are the ones the shipped template actually yields (20 rows on
    a continuation page, 11 on the page that keeps the summary block).
    """
    weeks = schedule.grid_weeks
    pages = _page_weeks(weeks, breakout, detailed)
    if detailed and breakout != "monthly" and len(weeks) > MAX_WEEK_COLUMNS and len(pages) == 1:
        pages = [("", list(weeks))]
    total = 0
    for index, _ in enumerate(pages):
        capacity = (SCHEDULE_ROWS_FINAL_PAGE if index == len(pages) - 1
                    else SCHEDULE_ROWS_CONTINUATION)
        rows = max(1, len(schedule.rows))
        total += max(1, -(-rows // capacity)) if capacity else 1
    return total


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

    # Three tiers, cheapest first, so a long schedule only pays for what it
    # actually needs:
    #
    #   1. The summary block describes the whole schedule, so it belongs on
    #      the final page only. Dropping it from continuation pages raises
    #      their floor by its full height -- worth about nine rows here.
    #   2. condense_media_plan_table then shrinks row height and font to fit,
    #      down to the legibility floor, exactly as the media plan does.
    #   3. Only if both leave rows over do programs split across pages, with
    #      weeks staying the primary split. Splitting rows first would
    #      multiply against the week pages and turn one schedule into six
    #      slides.
    # Continuation pages lose the summary block and so have the higher
    # capacity; the final page keeps it and has the lower one.
    capacity_continuation = _capacity_without_summary(prs, template)
    capacity_final = table_row_capacity(template, header_rows=2, totals_rows=1)
    rows = list(schedule.rows)

    plan_label = "Monthly Broadcast Plan:" if breakout == "monthly" else "Broadcast Plan:"
    base_index = slide_index(prs, template)
    plan = []
    for week_offset, (week_label, page_weeks) in enumerate(pages):
        is_last_week_page = week_offset == len(pages) - 1
        chunks = _chunk_rows(rows, capacity_continuation, capacity_final, is_last_week_page)
        for chunk_offset, chunk in enumerate(chunks):
            is_last_row_page = chunk_offset == len(chunks) - 1
            label_parts = [p for p in (week_label,) if p]
            if len(chunks) > 1:
                start = sum(len(c) for c in chunks[:chunk_offset]) + 1
                label_parts.append(f"Programs {start}–{start + len(chunk) - 1}")
            plan.append({
                "weeks": page_weeks, "rows": chunk,
                "label": ", ".join(label_parts) or plan_label,
                # The label heads the description. On a single-page schedule
                # with nothing typed in the description box it was heading
                # empty space -- "Broadcast Plan:" with a blank line under
                # it. Suppressed there, but NOT when the label is a real page
                # identifier ("Jan 2027", "Programs 1-11"), which says which
                # page you are looking at and is worth having with or without
                # a description.
                "show_label": bool(label_parts) or bool((plan_description or "").strip()),
                # Totals close out each week group, on its final row page.
                "totals": True,
                # A row-split page that isn't closing its week group carries
                # a page subtotal, not the group's.
                "subtotal_only": not is_last_row_page,
                "final": is_last_week_page and is_last_row_page,
            })

    for offset, page in enumerate(plan):
        slide = duplicate_slide(prs, template, insert_at=base_index + offset)
        overflow = _fill_broadcast_slide(
            slide, schedule, page["weeks"], detailed=detailed,
            plan_label=page["label"], plan_description=plan_description,
            rows=page["rows"], show_totals=page["totals"], is_last=page["final"],
            subtotal_only=page["subtotal_only"], all_weeks=weeks,
            show_label=page["show_label"])
        if overflow:
            warnings.append(overflow)

    delete_slide(prs, slide_index(prs, template))
    return warnings


def _summary_shape(slide):
    return _find_text_frame_with_token(slide.shapes, "BROADCAST_SUMMARY")


def _summary_shape_object(slide):
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.has_text_frame and _placeholder("BROADCAST_SUMMARY") in shape.text_frame.text:
            return shape
    return None


def _capacity_without_summary(prs, template):
    """Row capacity on a continuation page, where the summary block is gone.

    Measured rather than assumed: removing the summary raises the floor to
    whatever sits below it, which on this template is worth roughly nine
    extra rows. Getting this wrong is what turns a schedule into six slides
    instead of four -- the first attempt measured the floor with the summary
    still present, so continuation pages split rows they had ample room for.
    """
    return table_row_capacity(template, header_rows=2, totals_rows=1,
                              ignore=_summary_shape_object(template))


def _chunk_rows(rows, capacity_continuation, capacity_final, is_last_week_page):
    """Split programs across pages only when they genuinely don't fit."""
    capacity = capacity_final if is_last_week_page else capacity_continuation
    if not capacity or len(rows) <= capacity:
        return [rows]
    return [rows[i:i + capacity] for i in range(0, len(rows), capacity)]


def _fill_broadcast_slide(slide, schedule, weeks, detailed, plan_label,
                          plan_description, is_last, all_weeks,
                          rows=None, show_totals=True, subtotal_only=False,
                          show_label=True):
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    table = table_shape.table

    # Tier 1: the summary describes the whole schedule, so it goes on the
    # final page only. Removing it from continuation pages is also what
    # raises their floor, so this has to happen before anything is measured.
    if not is_last:
        summary_frame = _summary_shape(slide)
        if summary_frame is not None:
            element = summary_frame._txBody.getparent()
            element.getparent().remove(element)

    # The Days column is empty in a Campaign Schedule Report, which folds the
    # day into Day/Time. An always-blank column costs an inch that the week
    # columns badly need, so it's removed when nothing in the schedule fills
    # it -- and every column index after it shifts left by one.
    source_rows = list(rows if rows is not None else schedule.rows)
    drop_days = not any((r.days or "").strip() for r in schedule.rows)
    if drop_days:
        # Give the width to Program Name rather than letting the table shrink
        # -- it's the column that runs out of room first, and the table has a
        # fixed footprint on the slide.
        freed = table.columns[2].width
        remove_table_column(table, 2)
        table.columns[2].width = Emu(int(table.columns[2].width + freed))
    week_start = WEEK_COL_START - (1 if drop_days else 0)

    count = set_week_column_count(table, len(weeks) if detailed else 0, week_start)
    labels = week_column_labels(weeks, getattr(schedule, "week_header_style", "index")) \
        if detailed else []

    # Two header rows: the template's "Wk" band and the numbers beneath it.
    for offset, text in enumerate(labels):
        col = week_start + offset
        # "Wk" over "5/05" reads as "week of", and keeps the band from being
        # blank where Total and the demo column carry labels.
        _set_cell(table, 0, col, "Wk")
        _set_cell(table, 1, col, text)

    # Every program row appears on every page unless the caller has split
    # them across pages. Filtering to rows with spots in this page's weeks
    # would make the programme list change from page to page, and a page
    # whose weeks are all empty -- a hiatus week, which a football schedule
    # has -- would come out with no rows at all.
    rows = source_rows

    template_tr = table._tbl.findall(qn("a:tr"))[2]
    for _ in range(len(rows) - 1):
        clone_table_row(table, template_tr)

    for index, row in enumerate(rows):
        r = 2 + index
        col = 0
        _set_cell(table, r, col, row.station); col += 1
        _set_cell(table, r, col, row.time); col += 1
        if not drop_days:
            _set_cell(table, r, col, row.days); col += 1
        _set_cell(table, r, col, clean_program_name(row.program)); col += 1
        _set_cell(table, r, col, row.length); col += 1
        _set_cell(table, r, col, f"${row.rate:,.0f}")
        for offset, week in enumerate(weeks if detailed else []):
            _set_cell(table, r, week_start + offset,
                      str(row.spots_per_week.get(week, "") or ""))
        spots = sum(row.spots_per_week.get(w, 0) for w in weeks) if detailed else row.total_spots
        _set_cell(table, r, week_start + count, str(spots))
        # Scaled to this page's spots, for the same reason the spot count is.
        # Showing the flight's whole audience beside a page spot count of 0
        # read as a contradiction -- and on a paginated schedule the columns
        # would have summed to several times the real total.
        share = (spots / row.total_spots) if row.total_spots else 0
        _set_cell(table, r, week_start + count + 1,
                  f"{row.impressions * share / 1000:,.1f}")

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
    # Every page closes with a totals row -- a row-split page used to simply
    # stop after its last program, which reads as an unfinished table. The
    # label is what keeps a page figure from being mistaken for the flight's:
    # only the final page says FLIGHT TOTALS.
    if subtotal_only:
        label = "Page subtotal"
    elif is_last:
        label = "FLIGHT TOTALS"
    else:
        label = f"{plan_label} subtotal"
    grand = is_last and not subtotal_only
    _set_cell(table, totals_r, 0, label)
    cost_col = 5 - (1 if drop_days else 0)
    _set_cell(table, totals_r, cost_col,
              f"${(schedule.summary.gross_cost if grand else page_cost):,.0f}")
    for offset, week in enumerate(weeks if detailed else []):
        _set_cell(table, totals_r, week_start + offset,
                  str(sum(r.spots_per_week.get(week, 0) for r in rows) or ""))
    _set_cell(table, totals_r, week_start + count,
              str(schedule.summary.total_spots if grand else page_spots))
    _set_cell(table, totals_r, week_start + count + 1,
              f"{(schedule.summary.impressions if grand else page_impressions) / 1000:,.1f}")

    summary = schedule.summary
    # Humanized for the slide; the parse still supplies it, and anything that
    # can't be expanded comes back untouched.
    demo = wideorbit.humanize_demo(summary.demo_label) or "Adults"
    _fill_simple_tokens_in_slide(slide, {
        # Suppressed rather than left heading an empty description. The
        # totals rows below still use plan_label, because "... totals" is a
        # row label rather than a heading and reads fine either way.
        "BROADCAST_PLAN_LABEL": plan_label if show_label else "",
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
    # Grabbed before the fill, which consumes the {{BROADCAST_SUMMARY}}
    # token the shape is found by -- afterwards there's nothing left to
    # identify it with.
    summary_shape = _summary_shape_object(slide) if is_last else None
    if is_last:
        try:
            fill_bullet_list_in_slide(slide, "BROADCAST_SUMMARY", lines)
        except RuntimeError:
            pass

    # Same layout-floor verification the media plan table gets: derive the
    # font from the final row height and confirm the table clears whatever
    # sits below it. Only the wording differs -- a seller reading "the media
    # plan has too many lines" on a broadcast schedule slide would go looking
    # in the wrong place.
    overflow = condense_media_plan_table(slide, len(rows), extra_total_rows=0,
                                        header_rows=2)
    # After sizing, so it follows where the table really ended rather than
    # where the template guessed it would.
    place_summary_below_table(slide, summary_shape)
    if overflow:
        # Deliberately not suggesting a different breakout or view: program
        # rows repeat on every page, so neither Monthly nor totals-only
        # changes the row count. The only thing that helps is fewer programs.
        return (f"This schedule's {len(rows)} programs don't fit on the slide even at the "
                f"smallest readable size -- about 11 fit. The table will run into the summary "
                f"below it. Trim the schedule in Wide Orbit, or import it in two parts.")
    return None


def place_summary_below_table(slide, summary_shape, gap=_TABLE_CLEARANCE):
    """Sit the summary block under where the table actually ends.

    The template puts it at a fixed height chosen before anyone knew how many
    programs the schedule would have, and the table is sized to fit above it
    -- which works only for as long as the summary is the thing the sizer
    finds beneath the table. It stops working the moment the table's floor
    resolves to something lower (an empty footer shape, say): the table grows
    straight through the summary, and a page of programs renders on top of
    the flight's own totals.

    Anchoring it to the table's real bottom removes that dependency, and
    tightens a short schedule at the same time -- a four-program page used to
    leave the summary marooned an inch below the last row. It only ever moves
    as far as the slide allows; a table that has genuinely overrun is already
    reported by the caller, and shoving the summary off the bottom edge would
    hide the evidence rather than fix it.
    """
    bottom = table_bottom(slide)
    if bottom is None or summary_shape is None or summary_shape.top is None:
        return
    try:
        slide_height = slide.part.package.presentation_part.presentation.slide_height
    except Exception:                                            # noqa: BLE001
        return
    lowest = slide_height - (summary_shape.height or 0) - gap
    summary_shape.top = Emu(int(max(0, min(bottom + gap, lowest))))


# The plan slide is a signable document: it carries "Approved: ____ Date: ___"
# and the Premion terms. That signature is only meaningful on a page that also
# states the totals, what's included and the terms -- so the Included band may
# be made SMALLER to give the table room, and may be moved, but it may never
# leave the page the signature is on. Every plan slide is signable (a
# multi-option deck gives the client one signature line per option, and they
# sign the option they pick), which is why there is no version of this that
# lifts the band off "the long ones".
_SIGNATURE_ANCHOR = "approved:"
_INCLUDED_HEADING_ANCHOR = "included with campaign"


def _text_shape_containing(slide, needle):
    for shape in slide_map.iter_all_shapes(slide.shapes):
        try:
            if shape.has_text_frame and needle in shape.text_frame.text.lower():
                return shape
        except Exception:                                        # noqa: BLE001
            continue
    return None


def included_band_shapes(slide):
    """(heading, list frame, [decorative pictures]) of the Included band.

    Found by text anchor and geometry rather than by shape name, for the same
    reason _content_floor stopped looking up "Picture 12": a deck that renames
    or re-layers a shape would silently lose the handle, and this code moves
    things around by that handle. The pictures are whatever sits between the
    heading and the signature line -- an icon beside the list and the "One
    Solution" graphic on the right, both decorative. The list frame is found
    by its own token, so this has to run BEFORE fill_bullet_list_in_slide
    consumes it.
    """
    heading = _text_shape_containing(slide, _INCLUDED_HEADING_ANCHOR)
    # The SHAPE, not the text frame _find_text_frame_with_token would give
    # back -- this band gets moved and resized, and a TextFrame has no
    # geometry to move.
    listing = _text_shape_containing(slide, _placeholder("INCLUDED_LIST").lower())
    signature = _text_shape_containing(slide, _SIGNATURE_ANCHOR)
    if heading is None or listing is None:
        return None, None, []
    ceiling = heading.top
    floor = signature.top if signature is not None else None
    pictures = []
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE or shape.top is None:
            continue
        if shape.top < ceiling or (floor is not None and shape.top >= floor):
            continue
        pictures.append(shape)
    return heading, listing, pictures


def compress_included_band(slide, included_list):
    """Drop the band's decorative art so the plan table can have the space.

    The band is ~1.84in of a 7.5in slide and the table only gets 2.90in,
    which is what drives a long plan down to 6pt type. Almost all of that
    height is artwork: the content is the list itself, which the app already
    renders elsewhere as one comma-separated line. Removing the pictures and
    setting the list inline keeps everything the signature attests to on the
    page and returns most of the height.

    The band is also PARKED at its lowest legal position here, immediately
    above the signature line. That ordering is the whole trick: _content_floor
    reads the topmost shape below the table, so a band whose heading is still
    sitting at 4.63in raises no floor however much art was deleted from
    underneath it -- deleting the pictures alone changed nothing at all. The
    caller sizes the table against this parked position and then floats the
    band back up to the table's real bottom (place_included_band_below_table),
    the same order place_summary_below_table runs in.

    Returns True if it changed anything, so the caller can require that every
    plan slide in the deck compressed or none of them do.
    """
    heading, listing, pictures = included_band_shapes(slide)
    if heading is None or listing is None:
        return False
    for picture in pictures:
        element = picture._element
        element.getparent().remove(element)
    # Filled here rather than with the uncompressed lists, because the parked
    # position depends on how tall the compressed text actually is.
    fill_bullet_list_in_slide(slide, "INCLUDED_LIST", [", ".join(included_list)])
    strip_bullets(listing.text_frame)
    listing.height = Emu(int(_estimate_frame_height(
        listing.text_frame, listing.width, 1.0)))
    signature = _text_shape_containing(slide, _SIGNATURE_ANCHOR)
    if signature is not None:
        offset = listing.top - heading.top
        band_height = offset + (listing.height or 0)
        top = int(max(0, signature.top - _TABLE_CLEARANCE - band_height))
        heading.top = Emu(top)
        listing.top = Emu(top + offset)
    return True


def place_included_band_below_table(slide, heading, listing, gap=_TABLE_CLEARANCE):
    """Sit the compressed band under where the table actually ends.

    Same manoeuvre as place_summary_below_table and for the same reason: the
    band's template position was chosen before anyone knew how many lines the
    plan would have. It is clamped so it can never overrun the signature line
    -- the table having genuinely overrun is already reported by the caller,
    and pushing the band through the signature would hide that rather than
    fix it.
    """
    bottom = table_bottom(slide)
    if bottom is None or heading is None or listing is None:
        return
    signature = _text_shape_containing(slide, _SIGNATURE_ANCHOR)
    offset = listing.top - heading.top       # keep the band's internal spacing
    band_height = offset + (listing.height or 0)
    lowest = ((signature.top - gap - band_height) if signature is not None
              else heading.top)
    top = int(max(0, min(bottom + gap, lowest)))
    heading.top = Emu(top)
    listing.top = Emu(top + offset)


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
# The floor was 0.6, which on this template's 20pt body is 12pt -- and a real
# drafted Campaign Specs panel (19 bullets, ~1,200 characters) still needed
# 7.5in in a 6.5in space at that size, so it overflowed. 0.5 is 10pt, which is
# ordinary body copy on a 13.3in slide and buys the ~1in that case needs.
# Past that the panel genuinely has too much in it and the caller warns.
_SPECS_MIN_SCALE = 0.5
_SPECS_STEPS = (1.0, 0.92, 0.85, 0.78, 0.72, 0.66, 0.6, 0.55, 0.5)


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


# The Campaign Specs client-name title. Its own ladder rather than
# _SPECS_STEPS': the two are independent decisions about different frames, and
# sharing the constant would mean a change made for the body silently retunes
# the title. The floor is 0.4 of the template's 42pt -- 16.8pt, still well
# above the 10pt the specs body may be driven to, and past it a client's own
# name is being shrunk into something nobody reads across a room. The steps
# bunch up towards the bottom deliberately: a coarse 0.5 floor missed
# "Chesapeake Regional Medical Center" by two points and warned about a name
# that had room to fit. That name is also what sets the floor -- in the real
# Aptos Black it needs 18.0pt to clear the box, and there is no point having a
# ladder whose last rung a 34-character name can't reach.
_TITLE_STEPS = (1.0, 0.92, 0.85, 0.78, 0.72, 0.66, 0.6, 0.55, 0.51, 0.47, 0.43, 0.4)
_TITLE_MIN_SCALE = _TITLE_STEPS[-1]


def _shape_typeface(run, slide):
    """The typeface a run on an ordinary (non-table) shape renders in.

    Same job as _resolve_typeface, which can only answer for a table -- it
    reads the theme out of a per-table cache keyed on the table itself. A
    plain text frame resolves +mn-lt / +mj-lt against its own slide's master,
    for exactly the reason _theme_part_for_slide exists: this deck has 21
    masters and only master 0's minor font is Calibri.
    """
    name = run.font.name
    if name not in (None, "", "+mn-lt", "+mj-lt"):
        return name
    theme = _theme_part_for_slide(slide)
    major = minor = None
    if theme is not None:
        major, minor = _theme_typefaces(theme.blob.decode("utf-8", "ignore"))
    if name == "+mj-lt":
        return major or "Calibri Light"
    return minor or "Calibri"


# How much wider a face we could NOT measure actually draws.
#
# Why any of this exists: the Campaign Specs title asks for Aptos Black, which
# is on none of these machines -- it arrives with the deck's own embedded font
# data, which is EOT with the TTCOMPRESSED flag and which PIL cannot open. So
# text_metrics falls back to Calibri Bold while PowerPoint renders the real
# thing, and the first version of this fit trusted that measurement, reported
# the name fitted, and PowerPoint drew it straight through the edge of the
# panel: the under-measurement failure this project has now hit three times,
# arriving by a third route.
#
# These are per-face MEASUREMENTS, not one constant extrapolated everywhere.
# Each value is `max(real width / substituted width)` over six client-name
# strings (including a deliberately adversarial all-caps one, which is the
# glyph mix that produces the worst ratios), plus a 3% cushion for sampling
# error, rounded up to 0.01. Per-face spread across those six strings was
# 0.026-0.119, so a face's ratio is a real property of the face rather than an
# artifact of the string it was measured with.
#
# The installed faces were measured directly through text_metrics; Aptos Black
# had to be measured through PowerPoint's own TextRange.BoundWidth inside the
# master deck, since it renders only from the embedded font data. A ratio
# transfers between machines -- it describes the face, not the box -- so these
# hold on an instance where the face is missing, which is the only time they
# are consulted.
_SUBSTITUTE_WIDTH_RATIOS = {
    "verdana": 1.41,
    "arial black": 1.39,
    "georgia": 1.32,
    "tahoma": 1.26,
    "segoe ui black": 1.24,
    "aptos black": 1.20,
    "franklin gothic heavy": 1.19,
    "proxima nova black": 1.19,
    "proxima nova extrabold": 1.18,
    "trebuchet ms": 1.17,
    "impact": 1.03,
}

# An unmeasured face gets the worst ratio observed across every face above,
# not the value that happens to be right for this deck's title. Over-reserving
# shrinks a title slightly; under-reserving draws a client's name through the
# edge of the panel, and this project has already settled which way to err.
_UNCALIBRATED_SUBSTITUTE_RATIO = max(_SUBSTITUTE_WIDTH_RATIOS.values())

# The character-ratio estimate is a different mechanism, not a measurement of
# any font, so the table above does not transfer to it -- it has to be
# calibrated against the estimate itself. This is the deployed instance's
# normal path (no brand fonts, no Calibri, nothing to measure) and it is much
# looser: real/estimate ranges 0.816-1.525 across the same faces and strings,
# because the estimate ignores glyph mix entirely.
_ESTIMATE_WIDTH_RATIOS = {
    "aptos black": 1.36,
}
_UNCALIBRATED_ESTIMATE_RATIO = 1.58

# A run whose family is present but whose BOLD face is missing is measured in
# that same family's regular weight -- not in Calibri -- so none of the above
# applies. Measured previously and recorded in CLAUDE.md: Proxima Nova Light
# bold draws about 2% NARROWER than the regular weight, so the fallback
# already over-measures, which is the safe direction. Inflating it would
# shrink type for room nothing needs.
_SAME_FAMILY_WEIGHT_RATIO = 1.0

_UNCALIBRATED_FACES = set()


def uncalibrated_width_faces():
    """Faces this process reserved width for without a measurement of their own.

    Read by the caller so it can put this in front of a person. A silent
    extrapolation is the same class of failure as a silent substitution, and
    the project has paid for that one three times; a log line nobody opens is
    the phantom-slide-37 problem in a different costume.
    """
    return sorted(f for f in _UNCALIBRATED_FACES if f)


def reset_width_calibration_log():
    """Forget which faces were extrapolated. Called per build so the caption
    describes this deck rather than accumulating across a session."""
    _UNCALIBRATED_FACES.clear()


def width_calibration_note():
    """One plain sentence about any face whose width was extrapolated, or None.

    Deliberately separate from text_metrics.measurement_note(), which reports
    that a substitution HAPPENED. This reports that the correction applied for
    it was not measured against that face -- a different fact, and the one that
    decides whether a title can be trusted to sit inside its box.
    """
    faces = uncalibrated_width_faces()
    if not faces:
        return None
    return (f"The client-name title was sized for {', '.join(faces)} using the widest "
            f"correction measured on any font ({int((_UNCALIBRATED_SUBSTITUTE_RATIO - 1) * 100)}%), "
            f"rather than one measured on that face. Type may be smaller than it needs to "
            f"be; if a title still runs past the panel, that face needs measuring.")


def _width_ratio_for(typeface, entry, estimated):
    """The correction to apply to a measurement of `typeface` that wasn't made
    in `typeface`. Records the face when no measurement of it exists."""
    name = (typeface or "").strip().lower()
    table = _ESTIMATE_WIDTH_RATIOS if estimated else _SUBSTITUTE_WIDTH_RATIOS
    if name in table:
        return table[name]
    # Same family, wrong weight: measured in the face's own regular cut.
    resolved = ((entry or {}).get("resolved") or "").strip().lower()
    if not estimated and name and resolved.startswith(name):
        return _SAME_FAMILY_WEIGHT_RATIO
    _UNCALIBRATED_FACES.add(typeface)
    return _UNCALIBRATED_ESTIMATE_RATIO if estimated else _UNCALIBRATED_SUBSTITUTE_RATIO


def _width_correction(typeface, bold):
    """How much wider the drawn face is than whatever text_metrics measured.

    1.0 when the face itself was measured. One shared entry point so the two
    callers -- a width, and a wrap-line count -- can't drift onto different
    corrections for the same run.
    """
    entry = text_metrics.resolve_font(typeface, bold)
    if entry["status"] == "exact":
        return 1.0
    # Determined the same way _text_width_pt determines it, rather than from
    # the recorded status: a face whose file is present but unreadable also
    # lands on the estimate.
    estimated = text_metrics.text_width_points("M", typeface, 12, bold) is None
    return _width_ratio_for(typeface, entry, estimated=estimated)


def _text_width_pt(text, typeface, size_pt, bold):
    """Predicted RENDERED width in points -- not what was measured.

    Three paths, and they need different corrections: measured in the face
    itself (trust it), measured in Calibri because the face is missing (the
    per-face ratio table), or estimated from average character width because
    there is nothing to measure at all (its own, much looser table). Erring
    wide throughout, since under-reserving is what draws a name off the slide.
    """
    measured = text_metrics.text_width_points(text, typeface, size_pt, bold)
    if measured is None:
        measured = len(text or "") * size_pt * text_metrics.FALLBACK_CHAR_WIDTH_RATIO
    return measured * _width_correction(typeface, bold)


def _widest_line_pt(text_frame, slide, scale):
    """The longest paragraph's rendered width at this font scale."""
    widest = 0.0
    for para in text_frame.paragraphs:
        total = 0.0
        for run in para.runs:
            size = (run.font.size.pt if run.font.size else 18.0) * scale
            total += _text_width_pt(run.text, _shape_typeface(run, slide), size,
                                    bool(run.font.bold))
        widest = max(widest, total)
    return widest


# A title may wrap to this many lines before it is genuinely too long. Two,
# because the panel has room above the name for exactly one more line without
# disturbing anything (see _reflow_wrapped_title) and because a client name
# broken over three lines stops reading as a name.
_TITLE_MAX_LINES = 2


def _run_face(run, slide):
    """(typeface, bold) for a run, resolved against its own slide's theme."""
    return _shape_typeface(run, slide), bool(run.font.bold)


def _title_line_count(text_frame, slide, scale, usable_pt):
    """How many lines this title wraps to in `usable_pt` of width.

    The available width is DEFLATED by the face's width correction rather than
    the measurement being inflated: a face `r` times wider than the one we can
    measure fits as much in `usable/r` of measured width as it really will in
    `usable`. That reuses text_metrics' own greedy wrap -- the same algorithm a
    renderer uses -- instead of reimplementing wrapping here and letting the
    two drift.
    """
    lines = 0
    for para in text_frame.paragraphs:
        runs = list(para.runs)
        if not runs:
            continue
        typeface, bold = _run_face(runs[0], slide)
        size = (runs[0].font.size.pt if runs[0].font.size else 18.0) * scale
        effective = usable_pt / _width_correction(typeface, bold)
        lines += text_metrics.wrapped_lines(
            "".join(r.text for r in runs), effective, typeface, size, bold)
    return max(1, lines)


def _longest_word_pt(text_frame, slide, scale):
    """The widest single word, which wrapping cannot help with.

    Greedy wrap gives a word that doesn't fit a line of its own and lets it
    overflow, exactly as PowerPoint does -- so a line count of 2 is not on its
    own proof that nothing runs past the edge.
    """
    widest = 0.0
    for para in text_frame.paragraphs:
        for run in para.runs:
            typeface, bold = _run_face(run, slide)
            size = (run.font.size.pt if run.font.size else 18.0) * scale
            for word in (run.text or "").split():
                widest = max(widest, _text_width_pt(word, typeface, size, bold))
    return widest


def _reflow_wrapped_title(shape, original_top, template_pt, lines):
    """Move a now-multi-line title up so its LAST line sits where the single
    line did, and give the shape the height its text needs.

    The name has to keep its relationship to the "Campaign Specs" line beneath
    it -- that gap is the template's design and is only 46pt, nothing like
    enough for a second line. What there is, is room *above*: the title starts
    5.13in down a full-height panel with nothing above it at all. So the block
    grows upward from a fixed bottom instead of downward into the heading.

    Anchored to where the TEMPLATE's own 42pt line ends rather than to where
    this (smaller) one would, so a wrapped title sits exactly where an
    unwrapped one does. For one line at the template size the arithmetic is
    the identity, which is what keeps the single-line path untouched.
    """
    frame = shape.text_frame
    top_inset = _DEFAULT_CELL_INSET if frame.margin_top is None else frame.margin_top
    bottom_inset = _DEFAULT_CELL_INSET if frame.margin_bottom is None else frame.margin_bottom
    final_pt = max((run.font.size.pt for para in frame.paragraphs
                    for run in para.runs if run.font.size), default=template_pt)

    template_block = int(template_pt * _LINE_SPACING * _EMU_PER_POINT)
    text_block = int(lines * final_pt * _LINE_SPACING * _EMU_PER_POINT)
    bottom_target = original_top + top_inset + template_block

    shape.top = Emu(max(0, bottom_target - text_block - top_inset))
    shape.height = Emu(max(int(shape.height), text_block + top_inset + bottom_inset))


def fit_no_wrap_title(shape, slide):
    """Fit a title into its own box: shrink first, then wrap to two lines.

    Returns (scale, fits), the same contract as fit_text_frame.

    Width, not height, is the constraint: the Campaign Specs client name sits
    in a text box with `wrap="none"` and `<a:spAutoFit/>`, so PowerPoint
    neither wraps a long name nor -- on a generated file that is opened rather
    than edited -- recomputes the box around it. A name wider than the box
    simply draws straight out of the dark panel it's set on and across the
    slide. "Ridgeline Heating & Air" measures 407pt against 321pt of usable
    box at the template's 42pt, so this is the ordinary case, not an edge one.

    Two phases, and the first is unchanged from the single-line version:
    anything that fits on one line is sized exactly as it was before, still on
    one line, with its box untouched. Only a name that would otherwise have
    been held at the floor and reported as overflowing reaches phase two,
    where wrapping is switched on and the ladder is walked again against a
    two-line budget. That is what a 44-character legal name needs: at the
    floor it drew 405pt into a 359pt box, and it was a real client's name
    rather than a contrived one, so shrinking-and-warning wasn't good enough.
    """
    frame = shape.text_frame
    usable = max(1.0, (shape.width - frame.margin_left - frame.margin_right)
                 / _EMU_PER_POINT)
    original_top = shape.top
    template_pt = max((run.font.size.pt for para in frame.paragraphs
                       for run in para.runs if run.font.size), default=18.0)

    # --- phase 1: one line, box untouched --------------------------------
    for scale in _TITLE_STEPS:
        if _widest_line_pt(frame, slide, scale) <= usable:
            if scale < 1.0:
                _apply_font_scale(frame, scale)
            return scale, True

    # --- phase 2: let it wrap ---------------------------------------------
    frame.word_wrap = True
    for scale in _TITLE_STEPS:
        lines = _title_line_count(frame, slide, scale, usable)
        # A word too wide for a line of its own overflows however many lines
        # it is allowed, so both conditions have to hold.
        if lines <= _TITLE_MAX_LINES and _longest_word_pt(frame, slide, scale) <= usable:
            _apply_font_scale(frame, scale)
            _reflow_wrapped_title(shape, original_top, template_pt, lines)
            return scale, True

    # Even wrapped, even at the floor -- a name long enough to need a third
    # line. Take the floor and say so, but reflow for the lines it ACTUALLY
    # takes rather than for the two it was allowed: the extra line has to go
    # somewhere, and upward into empty panel is strictly better than downward
    # over the "Campaign Specs" heading. That isn't hiding the overrun -- the
    # warning still fires, and it's the same reasoning as the schedule
    # summary, which is placed below the real table bottom rather than the one
    # it was supposed to have.
    _apply_font_scale(frame, _TITLE_MIN_SCALE)
    _reflow_wrapped_title(shape, original_top, template_pt,
                          _title_line_count(frame, slide, 1.0, usable))
    return _TITLE_MIN_SCALE, False


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


def _clear_cell_to_bare_paragraph(cell):
    """Empty a cell all the way to a run-less paragraph (<a:endParaRPr> only)
    -- the shape PowerPoint treats as one line, rather than a paragraph
    holding a run whose text happens to be empty, which it does not. See
    add_full_flight_total_row for the incident this exists to prevent from
    recurring elsewhere.
    """
    body = cell.text_frame._txBody
    for para in body.findall(qn("a:p")):
        run = para.find(qn("a:r"))
        if run is None:
            continue
        rPr = run.find(qn("a:rPr"))
        if rPr is not None:
            end_pr = copy.deepcopy(rPr)
            end_pr.tag = qn("a:endParaRPr")
            para.append(end_pr)
        para.remove(run)


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

    # Counted from the RIGHT, not hardcoded. Cost is always the last column
    # and Impressions two before it, but an optional CPM column sits between
    # them -- with fixed indices 4 and 5 the cost landed in the CPM column
    # and the real cost cell kept the monthly figure it was cloned from,
    # which is visible on the slide as a CPM of "$98,000".
    last = len(cells) - 1
    cpm_index = last - 1 if len(cells) >= 7 else None
    impressions_index = last - (2 if cpm_index is not None else 1)

    values = [(cells[0], label), (cells[impressions_index], impressions), (cells[last], cost)]
    for cell, value in values:
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.text = value
                break
            break
    if cpm_index is not None:
        # A blended CPM across a full flight of mixed rate and flat-fee lines
        # isn't a rate anyone quotes, so the cell is cleared rather than
        # filled with something that looks authoritative -- and cleared all
        # the way. `run.text = ""` (the same path the other cells above take)
        # leaves the <a:r> element itself in the paragraph, just with empty
        # <a:t/>, which is a DIFFERENT shape from a cell that was never
        # touched at all (bare <a:endParaRPr>, no run) -- every other blank
        # cell in this cloned row is the latter. PowerPoint measures the two
        # differently: an empty run at 6pt bold Proxima Nova Light reported a
        # BoundHeight of three lines (confirmed via COM, isolated cell by
        # cell) where a bare endParaRPr paragraph reports one, and that
        # invisible cell was what grew the Hershey scenario's Full Flight
        # Total row from 14.4pt to 21.6pt -- an overlap with the
        # Included-with-Campaign block for a row with nothing visibly wrong
        # in it. Clearing to the run-less shape matches what every other
        # blank cell already renders as.
        _clear_cell_to_bare_paragraph(cells[cpm_index])

    # Claim the height this row's own text needs, rather than keeping the one
    # it was cloned with. The sizer reserved space for this label (it's passed
    # the same string), but the clone carries the totals row's height, and a
    # declared height shorter than the content is exactly what PowerPoint
    # silently grows -- into the block below the table.
    row = table.rows[new_index]
    font_pt = next((run.font.size.pt
                    for cell in cells
                    for para in cell.text_frame.paragraphs
                    for run in para.runs if run.font.size), None)
    if font_pt:
        row.height = max(row.height, _row_height_for(
            _lines_in_row(table, new_index, font_pt), font_pt))


def swap_named_picture_everywhere(prs, shape_name, image_path):
    """Image placeholder swap: replace the image of every picture shape with
    the given name, keeping its existing position/size untouched."""
    for slide in prs.slides:
        for shape in slide_map.iter_all_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.name == shape_name:
                image_part, rId = shape.part.get_or_add_image_part(image_path)
                shape._element.blipFill.blip.rEmbed = rId


def add_cpm_column(slide, rows, totals_cpm):
    """Insert a CPM column before Cost on a media plan table.

    Cloned from the Cost column rather than built, for the same reason the
    schedule's week columns are: a hand-made <a:gridCol>/<a:tc> inherits none
    of the theme's borders or fills. The width comes off the two widest text
    columns so the table keeps its footprint.

    `rows` is one display string per data row ("--" for a flat fee), and
    `totals_cpm` the blended figure for the totals row.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return False
    table = table_shape.table
    cost_col = len(table.columns) - 1

    clone_table_column(table, cost_col)
    # The clone lands to the right of Cost; the new CPM column is the
    # left-hand one of the pair, so Cost stays last where readers expect it.
    cpm_col = cost_col

    width = table.columns[cpm_col].width
    needed = width
    for donor in sorted(range(cost_col), key=lambda i: table.columns[i].width, reverse=True):
        if needed <= 0:
            break
        available = table.columns[donor].width - MIN_DONOR_WIDTH
        if available <= 0:
            continue
        take = min(available, needed)
        table.columns[donor].width = Emu(int(table.columns[donor].width - take))
        needed -= take

    _set_cell(table, 0, cpm_col, "CPM")
    for index, text in enumerate(rows):
        _set_cell(table, 1 + index, cpm_col, text)
    _set_cell(table, len(table.rows) - 1, cpm_col, totals_cpm)
    for row in table.rows:
        for cell in row.cells:
            _drop_surplus_paragraphs(cell)
    return True


def _prepare_media_plan_slide(slide, option):
    """Everything up to sizing: scalar tokens, rows, the CPM column, columns.

    Split from sizing because whether the Included band has to be compressed
    is a decision for the WHOLE DECK, not for one slide: no option's plan can
    be sized until every option's plan has been measured, or a three-option
    deck ships one plan slide with the decorative graphics and the next
    without. The scalar tokens are filled here rather than by the deck-wide
    pass because with several options each slide's title and totals differ.
    """
    _fill_simple_tokens_in_slide(slide, {
        "PLAN_TITLE": option["plan_title"],
        "TOTALS_LABEL": option["totals_label"],
        "TOTAL_IMPRESSIONS": option["total_impressions"],
        "TOTAL_COST": option["total_cost"],
    })

    plan_rows = option["rows"]
    fill_table_rows(
        slide,
        template_row_index=1,
        rows=plan_rows,
        field_to_token={
            "tactic": "TACTIC", "flight": "FLIGHT", "geo": "GEO",
            "targeting": "TARGETING", "impressions": "IMPRESSIONS", "cost": "COST",
        },
    )
    # Before condensing, so the extra column is part of what gets measured.
    if option.get("show_cpm"):
        add_cpm_column(slide, [r.get("cpm", "--") for r in plan_rows],
                       option.get("total_cpm", "--"))
    # After the CPM column exists, so the rebalance accounts for it.
    balance_text_columns(_find_table_shape(slide).table, header_rows=1, protect_last=3
                         if option.get("show_cpm") else 2)
    # Captured now: the list frame is found by its own {{INCLUDED_LIST}}
    # token, which fill_bullet_list_in_slide consumes at the end.
    heading, listing, _ = included_band_shapes(slide)
    return {"slide": slide, "option": option, "heading": heading,
            "listing": listing, "metrics": {}}


def _size_media_plan_slide(prepared):
    """Fit the table to whatever room the slide currently offers."""
    option = prepared["option"]
    full_flight_total = option.get("full_flight_total")
    return condense_media_plan_table(
        prepared["slide"], len(option["rows"]),
        extra_total_rows=1 if full_flight_total else 0,
        extra_total_label=(full_flight_total or {}).get("label"),
        template_metrics=prepared["metrics"])


def _finish_media_plan_slide(prepared, compressed):
    """The parts that must follow the final sizing pass.

    The full-flight row is cloned from the totals row and inherits the height
    set during sizing, and the Included list consumes the token the band was
    found by -- so both have to come after the table has stopped moving.
    """
    slide, option = prepared["slide"], prepared["option"]
    full_flight_total = option.get("full_flight_total")
    if full_flight_total:
        add_full_flight_total_row(
            slide, full_flight_total["label"],
            full_flight_total["impressions"], full_flight_total["cost"],
        )
    if compressed:
        # Already filled and sized by compress_included_band -- all that's
        # left is to float it up from its parked position to wherever the
        # table actually ended, so a plan that didn't need every row of the
        # room it was given doesn't leave a gap.
        place_included_band_below_table(
            slide, prepared["heading"], prepared["listing"])
    else:
        fill_bullet_list_in_slide(slide, "INCLUDED_LIST", option["included_list"])


# The avails table's value column is literal text in the template, not a
# {{TOKEN}}, so it is found by what it says. Anchored on "Avails" rather than
# the full string: the header is the one cell in that row carrying the word,
# and matching the whole phrase would silently do nothing if anyone ever
# retitled it.
AVAILS_HEADER_ANCHOR = "avails"


def set_avails_column_label(slide, label):
    """Rename the avails table's value column to name its basis.

    A number on a client-facing slide must say whether it is monthly or the
    whole flight; the two differ by the month count, and a reader has no way
    to tell them apart from the figure alone.

    Written run by run rather than through `text_frame.text`, which would
    collapse the header's formatting onto the paragraph -- the same rule the
    {{TOKEN}} fill follows.
    """
    if not label:
        return False
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return False
    for cell in table_shape.table.rows[0].cells:
        if AVAILS_HEADER_ANCHOR not in cell.text.lower():
            continue
        runs = [run for para in cell.text_frame.paragraphs for run in para.runs]
        if not runs:
            continue
        runs[0].text = label
        for extra in runs[1:]:
            extra.text = ""
        return True
    return False


def targeting_map_region(slide):
    """(left, top, width, height), all Emu -- the footprint of the stock
    image beside the avails table: the room to the table's right, down to
    whatever sits below (today, the small PREMION wordmark near the
    bottom-right). Derived from the slide's own geometry through
    `_floor_below`, the exact machinery `_content_floor` uses for the media
    plan table -- never a hand-placed rectangle, so a master-deck edit that
    moves the wordmark or widens the table is picked up automatically
    rather than needing this updated by hand. None when there's no table on
    the slide to measure from, or no slide-width available to bound the
    right edge against.
    """
    table_shape = _find_table_shape(slide)
    if table_shape is None:
        return None
    try:
        slide_width = slide.part.package.presentation_part.presentation.slide_width
    except Exception:                                              # noqa: BLE001
        return None

    left = table_shape.left + table_shape.width + _TABLE_CLEARANCE
    # Mirrors the table's own left margin on the right side, rather than a
    # hardcoded margin -- a symmetric layout is the reasonable default for a
    # region with no shape of its own to measure a margin from.
    right = slide_width - table_shape.left
    top = table_shape.top
    if right <= left:
        return None

    floor = _floor_below(slide, top, left, right, exclude=[table_shape])
    bottom = (floor - _TABLE_CLEARANCE) if floor is not None else (
        slide.part.package.presentation_part.presentation.slide_height - table_shape.top)
    if bottom <= top:
        return None
    return Emu(int(left)), Emu(int(top)), Emu(int(right - left)), Emu(int(bottom - top))


def place_targeting_map(slide, png_bytes):
    """Draws the targeting map into the stock-image area beside the avails
    table. `png_bytes` is None whenever no targeting group has been
    resolved to real zips yet -- the whole feature is optional, so this
    does nothing at all in that case, and the stock background photo shows
    through exactly as it does today. Sized to `targeting_map_region`'s own
    footprint, never hand-placed, and added fresh each call rather than
    tracked -- personalize() runs this once per deck, so there is nothing
    to collide with.
    """
    if not png_bytes:
        return
    region = targeting_map_region(slide)
    if region is None:
        return
    left, top, width, height = region
    slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, width, height)


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

    # Captured before fill_all_simple_tokens consumes the token it's found by.
    # Same ordering constraint as the broadcast summary shape.
    specs_title_shape = (_find_shape_with_token(specs_slide.shapes, "CLIENT_NAME")
                         if specs_slide is not None else None)

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
    if specs_title_shape is not None:
        title_scale, title_fits = fit_no_wrap_title(specs_title_shape, specs_slide)
        if not title_fits:
            warnings.append(
                f"\"{fill_data['client_name']}\" is too long for the Campaign Specs title "
                f"box even at the smallest readable size. It has been set at "
                f"{int(title_scale * 100)}% and will still run past the panel -- use a "
                f"shorter form of the name (the trading name rather than the full legal "
                f"one) and generate again.")

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
        # Before the rows: set_avails_column_label finds the header by its
        # own text, and filling the rows first doesn't disturb it -- but
        # doing the header first keeps the "find it by what it says" step
        # away from anything this function has already rewritten.
        set_avails_column_label(avails_slide, fill_data["avails"].get("label"))
        fill_table_rows(
            avails_slide,
            template_row_index=1,
            rows=fill_data["avails"]["rows"],
            field_to_token={"audience": "AUDIENCE", "geo": "GEO", "avails": "AVAILS"},
        )
        # After the rows, not before -- same ordering as the media plan's own
        # condense pass, and for the same reason: the rows are what make it
        # overflow. Unlike the media plan there's no second pass here (no
        # deck-wide band to compress and retry against), so one call settles it.
        avails_overflow = condense_avails_table(avails_slide, len(fill_data["avails"]["rows"]))
        if avails_overflow:
            warnings.append(avails_overflow)
        # None whenever no targeting group has been resolved to real zips --
        # place_targeting_map does nothing in that case, and the stock
        # background photo shows through the region exactly as it does today.
        place_targeting_map(avails_slide, fill_data["avails"].get("map_png"))

    # Three phases, because compressing the Included band is a whole-deck
    # decision. Size every option's plan first; if ANY of them came out below
    # the comfortable floor, compress the band on ALL of them and size again
    # with the room that frees. Deciding per slide would give a three-option
    # deck one plan with the decorative graphics and the next without.
    prepared = [_prepare_media_plan_slide(slide, option)
                for slide, option in zip(plan_slides, options)]
    overflows = [_size_media_plan_slide(entry) for entry in prepared]
    cramped = any(entry["metrics"].get("font_pt", _COMFORTABLE_TABLE_FONT_PT)
                  < _COMFORTABLE_TABLE_FONT_PT for entry in prepared)
    compressed = False
    if cramped:
        # All or nothing: if the band won't compress on every plan slide,
        # compress none of them rather than ship a deck that disagrees with
        # itself about what its own plan pages look like.
        compressed = all(compress_included_band(
            entry["slide"], entry["option"]["included_list"]) for entry in prepared)
        if compressed:
            # The same template_metrics dict goes back in, so this pass sizes
            # against the template's row heights rather than against its own
            # first-pass output. Only this pass's warnings are real -- a
            # first-pass overflow may have just been given the room to fix it.
            overflows = [_size_media_plan_slide(entry) for entry in prepared]

    for entry, overflow in zip(prepared, overflows):
        _finish_media_plan_slide(entry, compressed)
        if overflow:
            name = entry["option"].get("plan_title", "the media plan")
            warnings.append(f"{name}: {overflow}")

    swap_named_picture_everywhere(prs, "CLIENT_LOGO", fill_data["logo_path"])
    # Deliberately NOT appended to `warnings`: whether the type was measured
    # in the font it will be drawn in is a property of the machine, not of
    # this proposal, and it is true of every deck the deployed instance
    # builds. Putting a permanent condition in the channel that means "this
    # deck has a layout problem" is how a warning stops being read -- the same
    # mistake the validator's phantom slide-37 failure was. The caller reads
    # text_metrics.measurement_note() and shows it as a note instead.
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
