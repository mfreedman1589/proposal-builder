"""report_assembly.py -- fills the report master deck (REPORT_MASTER_vN.pptx,
CLAUDE.md/REPORT_MASTER_README.md name the version currently active -- never
hardcode it here, same rule as the proposal master deck) from a parsed
attribution/delivery export pair into a client-ready attribution report.

Pure deck-editing (python-pptx + report_charts) -- no Streamlit, no DB. The
caller (app.py) supplies the template path, the already-parsed
AttributionExport/DeliveryExport dataclasses (attribution_import.py), and
whatever content genuinely can't be derived from the export alone (goals,
what's-next -- see ATTRIBUTION_REPORT_PLAN.md's own token-source notes).
**No-proposal mode only** (Phase 3-4 of that plan) -- the proposal-linked
content path (goals/audience/geo from a real form_json, the targeted-vs-
visitor zip overlay) is Phase 5, and is not this module's job yet.

Slide map is a bare `key:` scan, not the proposal master's anchor/carry-
forward machinery -- this is a second, much smaller deck where every slide
already carries its own key (REPORT_MASTER_README.md), so there is nothing
to fall back to and nothing to carry forward.

Fill convention matches assembly.py's own rule exactly: run-level replace
only (never `text_frame.text`, which collapses run formatting), and multi-
item content is filled by deep-copying the tokenized paragraph/row and
inserting clones as siblings. This module reuses assembly.py's primitives
(`fill_bullet_list`, `clone_table_row`, `_replace_tokens_in_text_frame`,
`_tighten_cell_paragraphs`, `_find_shape_by_name`, `delete_slide`) rather
than re-implementing the same XML-cloning logic a second time -- assembly.py
is a pure python-pptx module already, not Streamlit-coupled, so importing it
here carries no extra weight.

**Every expected token must resolve to a real value before Generate.**
`MissingTokenError` names the slide/token/reason rather than the caller
ever shipping a literal "{{TOKEN}}" or a blank cell into a client-facing
deck. Two tokens have no export-derived signal at all -- `GOALS_BULLETS`
and `WHATS_NEXT_BULLETS` -- and are REQUIRED, non-defaulted parameters for
exactly that reason: a rep types them (Phase 3) or Claude drafts them from
notes (Phase 4's own drafting schema); this module never invents them.
`FLIGHT_LABEL` is the one token allowed to resolve to an empty string
(REPORT_MASTER_README.md's own documented exception -- "delete that tile's
shapes or leave the label with an empty value") -- Phase 3 always leaves it
blank, since deriving a real campaign flight separate from the report
period needs either a linked proposal (Phase 5) or the delivery file's own
per-campaign flight-detail tab, which -- see DECISIONS.md's attribution-
report-builder section -- turned out to carry more than one campaign's
worth of rows in the one real delivery fixture on hand, so extracting a
single flight from it isn't a Phase 3 problem with a clean answer yet.
"""
import io
import re
import types
from datetime import date, timedelta
from urllib.parse import urlparse

from pptx import Presentation
from pptx.util import Emu

import assembly
import attribution_benchmarks
import geo_resolver
import market_lookup
import report_charts
import slide_map
import targeting_map
from attribution_import import AttributionRow

DELIVERY_SET_RE = re.compile(r"^\s*delivery_set\s*:\s*true\s*$", re.IGNORECASE)
# Same marker-scan shape as delivery_set (ATTRIBUTION_REPORT_PLAN.md Phase 6,
# the one-slide-summary/case-study work, v0_9): a slide carrying this line
# is a standalone deliverable -- report:summary, report:case_study -- never
# part of the normal multi-slide report build. Driven off the deck itself,
# not a hardcoded key list, same reason delivery_set is: a future standalone
# slide needs no code change here to be dropped correctly.
STANDALONE_RE = re.compile(r"^\s*standalone\s*:\s*true\s*$", re.IGNORECASE)

# Longest-first, mirroring attribution_import._STATION_MARKETS -- a plain-
# language fallback for a station-pixel hint when the export's own by_market
# breakdown (the reliable source, see below) isn't populated.
_STATION_MARKET_NAMES = {"DC": "Washington, DC", "Harrisburg": "Harrisburg, PA"}


class MissingTokenError(Exception):
    """Raised naming the slide, token and reason -- never a silent blank or
    a literal "{{TOKEN}}" left in a client-facing deck."""


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def _int(n):
    return f"{int(round(n or 0)):,}"


def _pct(fraction, decimals=2):
    return f"{(fraction or 0.0) * 100:.{decimals}f}%"


def _money(amount):
    return f"${amount:,.0f}"


def _date_range_label(start, end):
    """"June 2026" for a whole calendar month; "Jun 1 - Jul 13, 2026"
    otherwise. Never invents a period the export didn't state -- returns
    None when either date is missing, which the caller turns into a loud
    MissingTokenError rather than a silently blank slide.

    Formats the day number by hand (`str(d.day)`, no leading zero) rather
    than strftime's `%-d`/`%#d` -- those flags are platform-specific
    (glibc/BSD vs. MSVCRT) and this app runs on Windows.
    """
    if not start or not end:
        return None
    next_month_first = date(end.year + (end.month == 12), (end.month % 12) + 1, 1)
    last_day_of_month = (next_month_first - timedelta(days=1)).day
    if start.year == end.year and start.month == end.month and start.day == 1 and end.day == last_day_of_month:
        return start.strftime("%B %Y")
    start_label = f"{start.strftime('%b')} {start.day}"
    end_label = f"{end.strftime('%b')} {end.day}, {end.year}"
    if start.year != end.year:
        start_label += f", {start.year}"
    return f"{start_label} - {end_label}"


# ---------------------------------------------------------------------------
# Slide map -- a bare key: scan, no anchors, no carry-forward
# ---------------------------------------------------------------------------

def _slide_by_key(prs):
    """{condition_key: slide_index} -- the last slide wins a duplicate key,
    matching slide_map.notes_key's own "last key: line wins" rule one level
    up. Every slide in this template carries its own key (README), so an
    untagged slide is a template defect, not a case this handles."""
    out = {}
    for index, slide in enumerate(prs.slides):
        key = slide_map.notes_key(slide)
        if key:
            out[key] = index
    return out


def _is_delivery_set(slide):
    if not slide.has_notes_slide:
        return False
    return any(DELIVERY_SET_RE.match(line)
              for line in slide.notes_slide.notes_text_frame.text.splitlines())


def _is_standalone(slide):
    if not slide.has_notes_slide:
        return False
    return any(STANDALONE_RE.match(line)
              for line in slide.notes_slide.notes_text_frame.text.splitlines())


_ATTRIBUTION_ROW_LIST_FIELDS = ("by_audience", "by_creative", "by_market", "by_zip",
                                "by_channel", "by_day_of_week")


def rehydrate_attribution(attribution_dict):
    """A duck-typed stand-in for a real `attribution_import.
    AttributionExport`, built from the STORED flattened dict (`dataclasses.
    asdict()`'s own shape -- `report_json["attribution"]`, the same shape a
    freshly-parsed export flattens to at upload time). Gives every *_facts()/
    top_*_rows() function in this module the same attribute access a
    freshly-parsed export has, so they can be reused UNCHANGED for the one-
    slide-summary/case-study "regenerate later, no second draft call" path
    (ATTRIBUTION_REPORT_PLAN.md Phase 6) -- this is what makes that promise
    real rather than needing a second, dict-native copy of every fact
    function this module already has.

    Only the row-list fields need reconstructing into real `AttributionRow`
    instances (their own `.label`/`.delivered_impressions`/etc. attribute
    access, which a plain dict doesn't have) -- everything else this module
    actually reads (`by_url`, `conversions_by_url`, `by_recency`, `by_
    referral_domain`, `by_device`, scalar fields) is already a plain value
    or dict either way, which `types.SimpleNamespace` alone covers.

    Known, accepted gap: `daily_trend`/`weekly_trend`/`monthly_trend`
    (`DateSeriesPoint` lists) pass through as plain dicts, unreconstructed --
    no summary/case-study token reads them today. Reconstruct them the same
    way if one ever does.

    `flight_start`/`flight_end` need the same treatment for a different
    reason: `db._json_safe` (the thing that actually makes `dataclasses.
    asdict()`'s output jsonb-safe at `log_attribution_report` time, per that
    call site's own comment) round-trips a `date` through `date.isoformat()`
    -- so what comes BACK out of storage is a plain "2026-08-03" string, not
    a `date`. Every caller that formats a period (`_date_range_label`,
    `report_headline_facts`) calls `.year`/`.strftime` on these two fields,
    so leaving them as strings would raise the moment a regenerated summary
    tried to render its own subtitle. Reconstructed here, once, rather than
    taught to every caller.
    """
    data = dict(attribution_dict or {})
    for field_name in _ATTRIBUTION_ROW_LIST_FIELDS:
        data[field_name] = [AttributionRow(**row) for row in (data.get(field_name) or [])]
    for field_name in ("flight_start", "flight_end"):
        value = data.get(field_name)
        if isinstance(value, str):
            try:
                data[field_name] = date.fromisoformat(value[:10])
            except ValueError:
                data[field_name] = None
    return types.SimpleNamespace(**data)


def rehydrate_delivery(delivery_dict):
    """`rehydrate_attribution`'s counterpart for a stored `report_json[
    "delivery"]` dict, or None when no delivery file was ever uploaded for
    this report (same None-means-absent convention the rest of this module
    uses). No row-list fields need reconstructing -- every list field on
    `DeliveryExport` this module actually reads (`top_publishers`, `by_
    creative`) is a list of plain tuples already, not a dataclass row.

    Known, accepted gap: `live_sports`, if present, stays a plain dict
    rather than a real `LiveSportsDelivery` -- no summary/case-study token
    reads it today.
    """
    if delivery_dict is None:
        return None
    return types.SimpleNamespace(**delivery_dict)


# ---------------------------------------------------------------------------
# Facts derived from the parsed export(s) alone -- no proposal, no DB
# ---------------------------------------------------------------------------

_GEOGRAPHY_TILE_MAX_MARKETS = 2  # a tile holds one short value -- 2026-09-06 correction


def _resolved_market_names(attribution):
    """The export's own real market names, by_market first (confirmed
    against the real MW file that the station-pixel `market_hint` can be
    flatly wrong for where a campaign actually attributed -- MW's real
    by_market rows are all North Carolina/Virginia DMAs, despite a
    DC-station pixel), falling back to a zip->market resolution. None
    when neither source has anything -- `market_hint` is a SEPARATE,
    later fallback in `geography_label` itself, not folded in here, since
    it's a single hint string, never a list to apply the 2026-09-06
    2-or-fewer rule to."""
    if attribution.by_market:
        labels = [row.label.title() for row in attribution.by_market if row.label]
        if labels:
            return labels
    if attribution.by_zip:
        zips = [row.label for row in attribution.by_zip]
        resolution = geo_resolver.zips_to_markets(zips)
        if resolution.resolved:
            return sorted((market_lookup.market_name(k) if market_lookup.available() else None) or k
                         for k in resolution.resolved)
    return None


def geography_label_from_names(names):
    """The Geography tile's one short value, given an already-resolved list
    of market names -- the shared collapse rule (2026-09-06: a tile holds
    one short value -- if a value would run past one line, it's the wrong
    container, not a fit problem) both `geography_label` (export-derived
    names) and Phase 5's proposal-linked path (`target_market_labels`-
    derived names) apply. Up to 2 real names are joined; 3 or more
    collapses to "N markets" (the full list belongs in
    `geography_overflow_bullet_from_names` instead -- never a truncated
    "X, Y, +N more" squeezed into the tile). None for an empty/falsy
    `names` -- the caller decides what to fall back to (a station-pixel
    hint for the export path; nothing at all for Phase 5, which omits the
    override entirely rather than passing an empty list)."""
    if not names:
        return None
    if len(names) > _GEOGRAPHY_TILE_MAX_MARKETS:
        return f"{len(names)} markets"
    return ", ".join(names)


def geography_overflow_bullet_from_names(names):
    """The full market list, as its own Audience-bullets line -- produced
    only when `geography_label_from_names` would collapse to "N markets"
    (more than `_GEOGRAPHY_TILE_MAX_MARKETS`). None when there's nothing to
    overflow. Shared by `geography_overflow_bullet` and Phase 5's
    proposal-linked path, same reasoning as `geography_label_from_names`."""
    if names and len(names) > _GEOGRAPHY_TILE_MAX_MARKETS:
        return "Markets: " + ", ".join(names)
    return None


def geography_label(attribution):
    """The Geography tile's one short value, from the EXPORT's own
    resolved market names, falling back to the station-pixel hint. See
    `geography_label_from_names` for the collapse rule itself."""
    label = geography_label_from_names(_resolved_market_names(attribution))
    if label:
        return label
    if attribution.market_hint:
        return _STATION_MARKET_NAMES.get(attribution.market_hint, attribution.market_hint)
    return None


def geography_overflow_bullet(attribution):
    """The export-derived overflow bullet -- see
    `geography_overflow_bullet_from_names`."""
    return geography_overflow_bullet_from_names(_resolved_market_names(attribution))


def _bucket_url(url):
    """A raw URL -> a client-readable label. Never the raw URL itself (up
    to 186 characters in a real export, meaningless to a client) --
    ATTRIBUTION_REPORT_PLAN.md's own rule for this slide. The second path
    segment is used when the first is a generic wrapper ("collections",
    "products", ...), so "/collections/beautyrest" reads as "Beautyrest"
    rather than every product page collapsing into one "Collections"
    bucket."""
    segments = [s for s in urlparse(url).path.split("/") if s]
    if not segments:
        return "Homepage"
    generic_wrappers = {"collections", "products", "category", "categories", "pages"}
    label_segment = segments[1] if segments[0].lower() in generic_wrappers and len(segments) > 1 else segments[0]
    return label_segment.replace("-", " ").replace("_", " ").title()


def top_url_rows(attribution, limit=8, include_conversions=False):
    """`include_conversions` (default off, explicit opt-in -- WAEPA's own
    "no half-states" rule) adds a "converted" key, bucketed the SAME way
    visits are (`_bucket_url`), from `attribution.conversions_by_url`."""
    buckets = {}
    for url, visitors in (attribution.by_url or {}).items():
        buckets[_bucket_url(url)] = buckets.get(_bucket_url(url), 0) + int(visitors or 0)
    conv_buckets = {}
    if include_conversions:
        for url, conversions in (attribution.conversions_by_url or {}).items():
            b = _bucket_url(url)
            conv_buckets[b] = conv_buckets.get(b, 0) + int(conversions or 0)
    ordered = sorted(buckets.items(), key=lambda kv: kv[1], reverse=True)[:limit]
    total = sum(buckets.values()) or 1
    out = []
    for label, visitors in ordered:
        row = {"label": label, "visitors": _int(visitors),
              "share": f"{visitors / total * 100:.0f}%"}
        if include_conversions:
            row["converted"] = _int(conv_buckets.get(label, 0))
        out.append(row)
    return out


# ---------------------------------------------------------------------------
# URL intent classification -- the deterministic half of the URL slide.
#
# Python tags each URL with an intent class from its path, and ONLY that:
# it counts, it never interprets. The model is handed these aggregates plus
# the campaign's goals and asked to connect the two (Phase 4) -- it quotes
# these numbers verbatim and computes nothing, the same facts-payload
# contract every other Claude call in this app already follows.
#
# Patterns are matched against each path SEGMENT (so "/collections/beautyrest"
# is tagged on "collections"), longest-and-most-specific class first, and a
# trailing "*" makes a prefix match -- which is what "/find-a-*" in the
# settled spec needs, and what catches "thank-you-offers" alongside
# "thank-you".
#
# Extended from the two real exports beyond the starter taxonomy, per the
# instruction to do so: MW contributes `store-locator` (its real
# foot-traffic path, 667 visits), `collections` and `search`; Cardinal
# contributes `thank-you*` (conversion confirmation pages -- the strongest
# completed-lead signal either file carries), `service-areas`, `offers`,
# `maintenance-plan` and `shop`. `careers` is deliberately mapped to
# `other`, not `lead`: a job seeker is not a campaign outcome, and letting
# it score as a lead would inflate exactly the number a client cares most
# about.
# ---------------------------------------------------------------------------

URL_INTENT_CLASSES = ("store_visit", "lead", "purchase", "consider", "learn",
                      "homepage", "other")

URL_INTENT_LABELS = {
    "store_visit": "Store visit intent",
    "lead": "Lead / contact intent",
    "purchase": "Purchase intent",
    "consider": "Product consideration",
    "learn": "Research / learning",
    "homepage": "Homepage",
    "other": "Other pages",
}

# Order matters: the first class whose pattern matches a segment wins, so
# the more committed intents are checked before the broader ones.
URL_INTENT_PATTERNS = (
    ("store_visit", ("locations", "location", "stores", "store", "store-locator",
                     "storelocator", "hours", "directions", "find-a-*", "find-us",
                     "visit-us")),
    ("lead", ("contact", "contact-us", "quote", "schedule", "book", "booking",
              "request", "appointment", "estimate", "consultation", "thank-you*",
              "get-started")),
    ("purchase", ("cart", "checkout", "financing", "finance", "apply", "buy",
                  "order", "purchase", "payment")),
    ("consider", ("products", "product", "collections", "collection", "catalog",
                  "inventory", "shop", "services", "service", "service-areas",
                  "search", "offers", "specials", "deals", "maintenance-plan",
                  "compare", "pricing")),
    ("learn", ("blog", "blogs", "about", "about-us", "faq", "faqs", "reviews",
               "why-*", "guide", "guides", "resources", "news", "how-*",
               "*-guide", "*-information")),
)

# Not pages a visitor meaningfully "went to" -- Shopify emits per-session
# tracking-pixel URLs (MW's export carries several, e.g.
# "/web-pixels@0d88c59a.../..."). Counting them as destinations would put
# machine noise in a client-facing table. Dropped before any tallying, and
# reported as a count so the drop is never silent.
_URL_NOISE_PREFIXES = ("web-pixels@", "cdn-cgi", "wpm@", ".well-known")


def _segment_matches(segment, pattern):
    if pattern.startswith("*") and pattern.endswith("*"):
        return pattern.strip("*") in segment
    if pattern.startswith("*"):
        return segment.endswith(pattern[1:])
    if pattern.endswith("*"):
        return segment.startswith(pattern[:-1])
    return segment == pattern


def classify_url_intent(url):
    """One of URL_INTENT_CLASSES for a raw URL, from its path alone.
    "other" is a real answer, not a failure -- a client-specific service
    page ("/plumbing", "/hvac") genuinely cannot be told from a blog
    category by pattern, and guessing would be worse than saying so."""
    segments = [s.lower() for s in urlparse(str(url or "")).path.split("/") if s]
    if not segments:
        # Its own class, not "other" -- added after checking the real data,
        # which is the whole reason the spec said to extend the taxonomy from
        # the exports. The homepage is 92% of everything MW would otherwise
        # have filed under "Other pages" (1,688 of 1,837 visits) and 66% of
        # Cardinal's (131 of 197). Folding the single largest destination on
        # the site into a bucket labelled "Other" on a client-facing slide is
        # actively misleading, and it was hiding how little genuinely
        # unclassifiable traffic there is: splitting it out takes real
        # "other" from 25% to 2% on MW and 62% to 21% on Cardinal.
        return "homepage"
    for intent, patterns in URL_INTENT_PATTERNS:
        for segment in segments:
            for pattern in patterns:
                if _segment_matches(segment, pattern):
                    return intent
    return "other"


def _is_noise_url(url):
    segments = [s.lower() for s in urlparse(str(url or "")).path.split("/") if s]
    return any(seg.startswith(p) for seg in segments for p in _URL_NOISE_PREFIXES)


def intent_summary_rows(attribution, include_conversions=False):
    """[{intent, label, visits, share, _visits_raw}] -- one row per intent
    class that actually occurs, biggest first, plus the counts the model is
    given to reason about. Never emits a class with zero visits: a table row
    reading "Purchase intent 0 0%" invites a client question whose answer is
    "that category doesn't apply to your site."

    `include_conversions` (default off, explicit opt-in) adds a "converted"
    key per row, tallied from `attribution.conversions_by_url` and folded
    into "Other" in EXACT lockstep with whichever visit-classes fold there
    -- a class's conversions can't survive in the table under a row its own
    visits no longer have, or the two columns would silently disagree about
    which classes exist.
    """
    tallies = {}
    conv_tallies = {}
    noise = 0
    for url, visitors in (attribution.by_url or {}).items():
        count = int(visitors or 0)
        if _is_noise_url(url):
            noise += count
            continue
        tallies[classify_url_intent(url)] = tallies.get(classify_url_intent(url), 0) + count
    if include_conversions:
        for url, conversions in (attribution.conversions_by_url or {}).items():
            if _is_noise_url(url):
                continue
            intent = classify_url_intent(url)
            conv_tallies[intent] = conv_tallies.get(intent, 0) + int(conversions or 0)
    total = sum(tallies.values()) or 1
    # Anything that would round to "0%" is folded into Other rather than
    # given its own row. Two reasons, both about the client reading it: a
    # row saying "Purchase intent  5  0%" invites a question whose honest
    # answer is "that is statistically nothing", and the fold is what keeps
    # this inside the 1-6 rows the template's IntentSummaryTable expects
    # (MW otherwise produces 7). The visits are NOT discarded -- they move
    # into Other, so the column still totals 100%.
    #
    # `intent_facts()` below deliberately does NOT fold: the model reasoning
    # about goals gets every class at full precision, because "8 lead-intent
    # visits" can be worth a sentence even when it is not worth a table row.
    folded, other_extra = {}, 0
    folded_conv, other_conv_extra = {}, 0
    for intent, count in tallies.items():
        if not count:
            continue
        if intent != "other" and round(count / total * 100) < 1:
            other_extra += count
            other_conv_extra += conv_tallies.get(intent, 0)
        else:
            folded[intent] = folded.get(intent, 0) + count
            folded_conv[intent] = folded_conv.get(intent, 0) + conv_tallies.get(intent, 0)
    if other_extra:
        folded["other"] = folded.get("other", 0) + other_extra
        folded_conv["other"] = folded_conv.get("other", 0) + other_conv_extra
    rows = []
    for intent, count in folded.items():
        if not count:
            continue
        row = {"intent": intent,
              "label": URL_INTENT_LABELS[intent],
              "visits": _int(count),
              "share": f"{count / total * 100:.0f}%",
              "_visits_raw": count,
              "_share_raw": count / total}
        if include_conversions:
            row["converted"] = _int(folded_conv.get(intent, 0))
        rows.append(row)
    rows.sort(key=lambda r: -r["_visits_raw"])
    return rows


def intent_facts(attribution, include_conversions=False):
    """The intent half of the facts payload Phase 4 hands the model, next to
    the campaign's goals. Deliberately numbers only -- no wording, no claim
    about alignment. `noise_visits` is reported rather than hidden so a
    later reader can tell a small total from a filtered one.

    `include_conversions` (default off, explicit opt-in -- never inferred
    from `attribution.has_conversions` here, since the caller,
    `build_facts_payload`, is where the rep's own toggle decides whether
    the model sees conversions at all) adds a "conversions" key per class,
    at the same full, unfolded precision as "visits" -- the model
    connecting an 8-conversion class to a goal is exactly this slide's job.
    """
    tallies = {}
    conv_tallies = {}
    for url, visitors in (attribution.by_url or {}).items():
        if _is_noise_url(url):
            continue
        intent = classify_url_intent(url)
        tallies[intent] = tallies.get(intent, 0) + int(visitors or 0)
    if include_conversions:
        for url, conversions in (attribution.conversions_by_url or {}).items():
            if _is_noise_url(url):
                continue
            intent = classify_url_intent(url)
            conv_tallies[intent] = conv_tallies.get(intent, 0) + int(conversions or 0)
    total = sum(tallies.values()) or 1
    ordered = sorted(((i, c) for i, c in tallies.items() if c), key=lambda kv: -kv[1])
    classes = [{"intent": intent, "label": URL_INTENT_LABELS[intent],
               "visits": count, "share": count / total}
              for intent, count in ordered]
    if include_conversions:
        for entry in classes:
            entry["conversions"] = conv_tallies.get(entry["intent"], 0)
    return {
        "classes": classes,
        "noise_visits": sum(int(v or 0) for u, v in (attribution.by_url or {}).items()
                            if _is_noise_url(u)),
    }


# A zip must carry at least this share of campaign impressions to be called
# an outperformer. 2026-09-06 ruling, and the reason it exists: the
# dashboard sorts zips by attributed impressions, so a zip can lead on
# volume at a perfectly average rate; sort by RATE instead and the top of
# the list is a zip with 40 impressions and one visit. Neither is "where
# the campaign worked best." Requiring above-average rate AND real volume
# is. Measured against both real datasets at 1.0%: MW yields 12 eligible
# of 242 zips, Cardinal 7 of 22 -- neither below five, so the floor trims
# the long tail without emptying a small campaign. (Cardinal is not
# actually limited by this floor at all -- dropping it to 0.25% still
# yields 7, because its binding constraint is the above-average-rate test.)
ZIP_MIN_SHARE = 1.0


def top_zip_rows(attribution, limit=10, include_conversions=False):
    """(rows, dropped_zips) -- "where the campaign worked best", not
    "biggest" and not "highest rate". See ZIP_MIN_SHARE above for why
    either of those alone is the wrong list.

    Eligible = attributed rate strictly above the campaign average AND
    impression share >= ZIP_MIN_SHARE. Eligible zips rank by attributed
    impressions, descending, capped at `limit`. If fewer than `limit`
    qualify, the remainder is filled from the next-highest by attributed
    impressions REGARDLESS of rate, each carrying `outperformer: False`
    so the narrative can't describe a filler row as one -- the row is
    still worth showing (it is real volume) but it did not beat the
    average, and the deck must not imply it did.

    A zip whose own area can't be resolved (no market lookup installed,
    or the zip itself has no county/market -- a real, deliverable
    PO-box-only zip has no ZCTA at all, 2026-09-13's Ashburn/20149 find)
    tries two coarser fallbacks (`market_lookup.zip3_market_fallback`/
    `zip3_state_fallback` -- see their own docstrings) before being
    DROPPED from `rows` outright and named in `dropped_zips` instead: a
    blank Area cell ships to a client either way, and "20149 | | 0.4%"
    teaches nothing while still advertising the hole. The caller decides
    what to do with `dropped_zips` (`_fill_zip_analysis` logs it as a dev
    warning); this function never guesses a place it can't support.

    `include_conversions` (default off, explicit opt-in) adds a
    "conversions" key (the zip's own `AttributionRow.conversion_impressions`)
    -- WAEPA's own explicit ruling: this table is already five columns
    wide and does NOT gain a sixth for conversions; this is for the facts
    payload/narrative to cite instead, never the deck's own zip table.
    """
    zips_all = list(attribution.by_zip or [])
    if not zips_all:
        return [], []
    # Denominator: the file's own headline delivered count, falling back to
    # the sum of the zip rows. On both real datasets these are identical, so
    # the choice is invisible today -- the fallback exists for an export
    # whose zip tab is a partial cut of total delivery, where dividing by
    # the (smaller) zip sum would overstate every share.
    denominator = attribution.delivered_impressions or sum(
        z.delivered_impressions for z in zips_all) or 0
    baseline = attribution.attributed_rate or 0.0

    def share_pct(row):
        return (row.delivered_impressions / denominator * 100) if denominator else 0.0

    by_volume = sorted(zips_all, key=lambda r: r.attributed_impressions, reverse=True)
    eligible = [r for r in by_volume
               if baseline and r.attributed_rate > baseline and share_pct(r) >= ZIP_MIN_SHARE]
    chosen = [(r, True) for r in eligible[:limit]]
    if len(chosen) < limit:
        already = {id(r) for r, _ in chosen}
        for row in by_volume:
            if len(chosen) >= limit:
                break
            if id(row) not in already:
                chosen.append((row, False))

    resolution = geo_resolver.zips_to_markets([r.label for r, _ in chosen]) if chosen else None
    zip_to_market = {}
    if resolution and resolution.resolved:
        for market_key, info in resolution.resolved.items():
            name = (market_lookup.market_name(market_key) if market_lookup.available() else None) or market_key
            for z in info["zips"]:
                zip_to_market[z] = name

    # A blank Area cell ships to a client either way -- "20149 | | 0.4%"
    # teaches nothing and just as notices the hole (2026-09-13, the
    # Ashburn/WAEPA find: a real, deliverable PO-box-only zip has no
    # ZCTA and so no county/market of its own). Two coarser fallbacks,
    # tried in order, before the row is dropped outright:
    #   1. the market every OTHER real, resolved zip sharing this zip's
    #      3-digit prefix agrees on (market_lookup.zip3_market_fallback)
    #      -- a PO-box zip still sits inside a normal USPS geographic
    #      numbering block, so its neighbors' unanimous market is almost
    #      always right for it too.
    #   2. failing that, the STATE those same neighbors agree on
    #      (market_lookup.zip3_state_fallback) -- coarser, but still a
    #      real place, never a guess across a boundary (both fallbacks
    #      require UNANIMOUS agreement among real zips sharing the
    #      prefix, not a majority -- a confident-looking wrong answer at
    #      an actual DMA/state line is worse than falling through).
    # A genuinely unresolvable zip (no listed fallback state has been
    # observed in practice) is dropped, not shown with no place --
    # logged by the caller (`_fill_zip_analysis`) as a dev warning
    # naming the zip, never a rep-facing one. No backfill to keep the
    # table at `limit` rows -- a dropped row just isn't replaced today.
    out, dropped_zips = [], []
    for row, outperformer in chosen:
        area = zip_to_market.get(row.label, "")
        if not area and market_lookup.available():
            fallback_market = market_lookup.zip3_market_fallback(row.label)
            if fallback_market:
                area = market_lookup.market_name(fallback_market) or fallback_market
            else:
                area = market_lookup.zip3_state_fallback(row.label) or ""
        if not area:
            dropped_zips.append(row.label)
            continue
        multiple = f"{row.attributed_rate / baseline:.2f}x" if baseline else "--"
        entry = {
            "zip": row.label,
            "area": area,
            "share": f"{share_pct(row):.1f}%",
            "rate": _pct(row.attributed_rate),
            "multiple": multiple,
            "outperformer": outperformer,
        }
        if include_conversions:
            entry["conversions"] = row.conversion_impressions
        out.append(entry)
    return out, dropped_zips


# ---------------------------------------------------------------------------
# Recency / referral / day-of-week (ATTRIBUTION_REPORT_PLAN.md, 2026-09-10) --
# all three were already parsed (attribution_import.py) and never surfaced
# anywhere: no slide, no facts. report:response_profile is the new slide;
# these are the pure, testable-now facts derivation it and the drafting
# prompt's highlight/takeaway rules will both draw on. Slide-fill code is
# blocked on Matt's own v0_6 template -- this module only computes facts.
# ---------------------------------------------------------------------------

_RECENCY_BUCKET_RE = re.compile(r"(\d+)\s*-\s*(\d+)")


def recency_facts(attribution):
    """{"total", "buckets": [{"bucket", "visitors", "share"}, ...],
    "share_within_0_3_days"} from `attribution.by_recency`. The "0-3 days"
    share is picked out by parsing each bucket's own LOW bound (never a
    hardcoded label string like "00 - 03 DAYS" -- a real export's own
    spacing/padding isn't guaranteed) and summing whichever bucket(s) start
    at 0; there is exactly one on every real export checked. None when the
    export carries no recency tab at all (nothing to divide into)."""
    buckets = attribution.by_recency or {}
    total = sum(buckets.values())
    if not total:
        return None
    rows = []
    immediate = 0
    for label, visitors in buckets.items():
        rows.append({"bucket": label, "visitors": visitors, "share": visitors / total})
        match = _RECENCY_BUCKET_RE.search(label)
        if match and int(match.group(1)) == 0:
            immediate += visitors
    return {"total": total, "buckets": rows, "share_within_0_3_days": immediate / total}


def referral_facts(attribution):
    """{"total", "sources": [{"source", "visitors", "share"}, ...],
    "direct_share"} from `attribution.by_referral_domain`. `direct_share`
    is its own top-level key (not just another row) because Matt's own
    interpretation rule treats Direct as the strongest single signal --
    the model should be able to cite it without hunting the row list.
    None when the export carries no referral tab."""
    sources = attribution.by_referral_domain or {}
    total = sum(sources.values())
    if not total:
        return None
    rows = [{"source": s, "visitors": v, "share": v / total} for s, v in sources.items()]
    return {"total": total, "sources": rows,
            "direct_share": sources.get("Direct", 0) / total}


# "Uneven" is a NAMED, code-level definition (Matt's own instruction) so
# the same threshold governs both the facts payload's own "uneven" flag
# (what the drafting prompt is told it may claim a weekday pattern from)
# and, later, the response_profile slide's own day-of-week table --
# CONDITIONAL, deleted with its header when the week is flat. A day below
# the impression floor is excluded from the max/min comparison entirely --
# a near-zero-volume day's rate is noise (one attributed visit on a
# thousand-impression day can swing a rate wildly), not a real pattern,
# and letting it drive "uneven" would make the flag fire on volume
# starvation rather than a genuine weekday effect.
_DAY_OF_WEEK_UNEVEN_RATIO = 1.5
_DAY_OF_WEEK_MIN_IMPRESSIONS_FLOOR = 1000


def day_of_week_facts(attribution):
    """{"days": [{"day", "attributed_rate", "delivered_impressions"}, ...],
    "uneven", "best_day", "worst_day"} from `attribution.by_day_of_week`
    (attribution_import.py's own 2026-09-10 correction -- see that
    module's gotcha 2 for why this is a real weekday aggregate, not a
    trailing window). "uneven" is true when the best/worst day's
    attributed-rate RATIO clears `_DAY_OF_WEEK_UNEVEN_RATIO`, considering
    only days that clear `_DAY_OF_WEEK_MIN_IMPRESSIONS_FLOOR`. None when
    the export has no day_of_week tab, or fewer than two days clear the
    floor (nothing to compare)."""
    rows = attribution.by_day_of_week or []
    eligible = [r for r in rows if r.delivered_impressions >= _DAY_OF_WEEK_MIN_IMPRESSIONS_FLOOR]
    if len(eligible) < 2:
        return None
    best = max(eligible, key=lambda r: r.attributed_rate)
    worst = min(eligible, key=lambda r: r.attributed_rate)
    uneven = bool(worst.attributed_rate) and (best.attributed_rate / worst.attributed_rate
                                              ) >= _DAY_OF_WEEK_UNEVEN_RATIO
    return {
        "days": [{"day": r.label, "attributed_rate": r.attributed_rate,
                  "delivered_impressions": r.delivered_impressions} for r in rows],
        "uneven": uneven,
        "best_day": {"day": best.label, "attributed_rate": best.attributed_rate},
        "worst_day": {"day": worst.label, "attributed_rate": worst.attributed_rate},
    }


def response_profile_facts(attribution):
    """{"recency", "referral", "day_of_week"} -- the whole report:
    response_profile facts bundle, each key None-when-absent per this
    file's own convention. A single grouped call so `build_facts_payload`
    doesn't need three separate optional keys at its own top level."""
    return {
        "recency": recency_facts(attribution),
        "referral": referral_facts(attribution),
        "day_of_week": day_of_week_facts(attribution),
    }


def device_split_facts(attribution):
    """{"device", "share", "count"} for the top device type by ATTRIBUTED
    impressions, from `attribution.by_device` ({device_type: attributed_
    impressions}) -- the one-slide summary's own sidebar second block
    (ATTRIBUTION_REPORT_PLAN.md Phase 6), deliberately generic rather than
    goal-specific (device mix isn't a thing a campaign goal targets). None
    when the export carries no device breakdown at all, or every device's
    count is zero."""
    by_device = attribution.by_device or {}
    total = sum(int(v or 0) for v in by_device.values())
    if not total:
        return None
    top_device, top_count = max(by_device.items(), key=lambda kv: int(kv[1] or 0))
    return {"device": top_device, "share": int(top_count or 0) / total, "count": int(top_count or 0)}


# ---------------------------------------------------------------------------
# OTT Retargeting (Audience Marketplace export) -- a THIRD, separate export
# from a THIRD dashboard (attribution_import.parse_ott_retargeting_export),
# never part of the delivery set. Pure facts derivation only, same reason
# as response_profile above.
# ---------------------------------------------------------------------------

def ott_retargeting_facts(ott):
    """{"impressions", "clicks", "ctr", "actions" (optional), "by_ad_size",
    "creative_groups" (optional), "top_screen", "blended" (optional)} from
    an `attribution_import.OTTRetargetingExport`, or None when no OTT
    retargeting export was uploaded -- mirrors `build_facts_payload`'s own
    "delivery"/"live_sports" None-when-absent convention. "actions" is
    present only when > 0 (the same present-and->0 rule `has_conversions`
    already follows -- Cardinal's own real Actions column is 0, and a
    fact the model could turn into a false claim never enters the
    payload). "creative_groups" is present only when there's more than
    one distinct creative CONCEPT (Cardinal: one creative in five sizes,
    so this key is absent for it) -- the same "shown only when creative
    names differ beyond their size suffix" rule the slide itself follows,
    computed once here rather than re-derived by the drafting prompt.

    "blended" carries impressions and uniques only, never "frequency" --
    confirmed against the real Cardinal export that the "PREMION + OTT
    RETARGETING" tab's blended figures are campaign-to-date, not scoped to
    the export's own reporting period: every OTHER tab in that file (the
    top-line KPIs, the daily CAMPAIGN SUMMARY, all four KPI-shaped tabs)
    sums to exactly 300,207 impressions for August alone, but blended
    impressions is 950,329 -- 3.16x that, with no in-file component that
    explains the gap -- while the file's own Pacing Report tab shows the
    underlying campaign spans March 2026 - February 2027, comfortably wide
    enough to account for the difference as accumulation since campaign
    start. There is no field on the blended tab itself marking its own
    period, so this can't be re-derived automatically for a future export;
    a bare "65 frequency" on a one-month report answers a question nobody
    asked (Matt, reviewing WAEPA, 2026-09-10) -- impressions and uniques
    are still real, standalone facts either way, so they stay in the
    payload; frequency, which is meaningless without knowing what span it
    was computed over, does not.
    """
    if ott is None:
        return None
    facts = {
        "impressions": ott.impressions,
        "clicks": ott.clicks,
        "ctr": ott.ctr,
        "by_ad_size": [{"ad_size": r.ad_size, "label": r.label, "impressions": r.impressions,
                        "clicks": r.clicks, "ctr": r.ctr} for r in ott.by_ad_size],
        "top_screen": [{"screen": s, "clicks": c} for s, c in ott.by_screen] or None,
    }
    if ott.has_actions:
        facts["actions"] = ott.actions
    if len(ott.creative_groups) > 1:
        facts["creative_groups"] = [{"name": g.base_name, "impressions": g.impressions,
                                     "clicks": g.clicks, "ctr": g.ctr}
                                    for g in ott.creative_groups]
    if ott.blended_impressions:
        facts["blended"] = {"impressions": ott.blended_impressions,
                            "uniques": ott.blended_uniques,
                            "frequency": ott.blended_frequency}
    return facts


# ---------------------------------------------------------------------------
# Optimization engine (attribution-module-framework.md §4-6; ATTRIBUTION_
# REPORT_PLAN.md Phase 6, built after the advertiser spine landed). Pure
# derivation only, same reason as every other *_facts function in this file
# -- no DB, no Streamlit. Deterministic Python identifies candidates; the
# model turns a qualifying one into a thread's own "action" (never a
# parallel list -- see app.build_attr_draft_prompt's own optimization
# rules). Every number here rides the SAME facts-only tracing every other
# section of this payload does (app._attr_payload_numbers walks the whole
# dict); only a key prefixed "_internal" is exempt, matching "zip"."_
# internal_min_share_pct"'s own convention -- an eligibility RULE, never a
# fact about any one candidate.
# ---------------------------------------------------------------------------

OPTIMIZATION_DIMENSIONS = ("zip", "market", "creative", "day_of_week", "publisher")

# Publisher defaults OFF (Matt's ruling, 2026-09-14): the framework's own
# seasonality caveat ("a sports network looks weak until the playoffs
# start") makes a publisher-level cut unreliable enough that a rep should
# turn this on deliberately rather than have it silently recommend against
# a publisher about to turn around. Every other dimension defaults on.
OPTIMIZATION_DIMENSIONS_DEFAULT_ON = frozenset({"zip", "market", "creative", "day_of_week"})

# The material-swing floor, reused verbatim from the threads work's own rule
# (app.py's drafting prompt: "a 5x ZIP, a 2x market gap... are swings; a
# 0.02-point difference never is") -- one definition of "material" across
# this app, not a second one invented for this engine.
_OPT_MATERIAL_RELATIVE_DIFF = 0.20   # >= 20% relative difference on a rate

# A value with fewer delivered impressions than this floor is noise, not a
# finding -- day_of_week_facts's own _DAY_OF_WEEK_MIN_IMPRESSIONS_FLOOR
# idea, generalized across dimensions. Zips run far smaller than days/
# markets/creatives/publishers by nature, so they get their own, lower
# floor (matches ZIP_MIN_SHARE's own real-volume reasoning above).
_OPT_MIN_IMPRESSIONS_FLOOR = {"zip": 200, "market": 1000, "creative": 1000,
                              "day_of_week": 1000, "publisher": 1000}

# A value already running at less than this share of its dimension's own
# peer-average impressions was a MEDIA-BUY choice, not a performance
# result -- MW's "removed Wednesdays" is the reference case (CLAUDE.md).
# An underperforming value this suppressed already is reported as ALREADY
# limited, never re-discovered as a new cut.
_OPT_ALREADY_LIMITED_IMPRESSION_RATIO = 0.5

# Low/Moderate/High cap how many candidates SURFACE, never a target
# (framework §6: "If only three ZIPs genuinely stand out under a 20% cap,
# we cut three"). The cap is a ceiling on the dimension's own TOTAL row
# count, not on however many already qualify. ZIPs get the framework's own
# stated percentages, since that's the one dimension with enough real
# cardinality for a percentage to mean anything; every other dimension
# (day of week maxes at 7 rows, market/creative/publisher are typically a
# handful) gets a flat count cap instead -- 10-20% of 7 days is not a
# usable number.
_OPT_ZIP_CAP_PCT = {"moderate": 0.10, "high": 0.20}
_OPT_ZIP_LOW_CAP = 2
_OPT_SMALL_DIMENSION_CAP = {"low": 1, "moderate": 2, "high": 3}


def _optimization_dimension_rows(attribution, dimension):
    """The real AttributionRow list backing one optimization dimension, or
    [] when the export doesn't carry it (e.g. no channel-name tab)."""
    return {
        "zip": attribution.by_zip,
        "market": attribution.by_market,
        "creative": attribution.by_creative,
        "day_of_week": attribution.by_day_of_week,
        "publisher": attribution.by_channel,
    }.get(dimension, [])


def _optimization_cap(dimension, level, total_count):
    """The maximum number of candidates this dimension/level may surface,
    given how many rows the dimension actually has -- see the module-level
    comment above for why zip alone gets a percentage."""
    if dimension == "zip":
        if level == "low":
            return _OPT_ZIP_LOW_CAP
        return max(1, int(total_count * _OPT_ZIP_CAP_PCT[level]))
    return _OPT_SMALL_DIMENSION_CAP[level]


def optimization_candidates(attribution, level, dimensions, prior_periods=None,
                            in_effect_values=None):
    """`facts["optimizations"]` -- candidates to recommend REDUCING or
    REMOVING, ranked worst-first and capped per dimension/level, plus the
    rest of what qualified (`watch_list`) and what was already suppressed
    in the media plan rather than newly discovered (`already_limited`).

    Optimize by subtraction, structurally, not by prompt instruction alone
    (framework §4): every candidate here is a value whose OWN attributed
    rate sits materially BELOW the campaign baseline -- there is no
    over-performing-value candidate type in this structure at all, so "add
    more of X" has no data to build itself from. The drafting prompt still
    tells the model to phrase actions as remove/reduce (belt and suspenders
    against it reaching for an over-performer named elsewhere in the
    payload), but the candidate list itself cannot contain a growth
    recommendation.

    Timing (framework §5) is enforced, not merely suggested: `evidence_
    periods` (this report plus every prior one logged for the same
    advertiser) below 2 means a first report, and the framework is explicit
    -- "Month 1 -- report only. No optimizations." -- so every list here is
    empty and `timing_note` says why, regardless of level/dimensions. From
    the second report on, candidates compute normally; `timing_note` still
    names the evidence count so the model (and the rep) can frame an early
    read as exactly that.

    `level == "none"` (the optimization-sequence work, ATTRIBUTION_REPORT_
    PLAN.md Phase 6): a client who wants reporting only. The candidate/
    watch_list/already_limited loop below never runs -- "no candidates
    render, no checklist" -- but this function is still called every time
    (never skipped by the caller), because the SEPARATE in-effect
    measurement (`optimizations_in_effect`/`optimization_history`, below)
    is unconditional: continuity on a PAST accepted optimization doesn't
    stop mattering just because new recommendations are off this month.

    `in_effect_values` ({(dimension, value), ...}, optional) excludes a
    value already accepted in a PRIOR report from candidacy here -- it's
    being measured (did it work), not re-discovered as if new. Same
    exclusion shape as `already_limited`, a different reason.

    A dimension the rep didn't enable, or one the export has no rows for
    (e.g. a single-market campaign has nothing to say about "market"), is
    silently absent from every list -- never a fabricated empty finding.
    """
    evidence_periods = len(prior_periods or []) + 1
    result = {
        "level": level,
        "dimensions_enabled": sorted(d for d in dimensions if d in OPTIMIZATION_DIMENSIONS),
        "_internal_evidence_periods": evidence_periods,
        "candidates": [],
        "watch_list": [],
        "already_limited": [],
    }
    if evidence_periods < 2:
        result["timing_note"] = ("First report for this account -- no optimization "
                                 "recommendations yet. The framework calls for waiting "
                                 "until a trend is established (typically the second or "
                                 "third report) before recommending a cut.")
        return result
    if level == "none":
        result["timing_note"] = ("Optimization level is set to None for this client -- "
                                 "reporting only. No new recommendations this month.")
        return result
    result["timing_note"] = (f"{evidence_periods} periods of evidence for this account "
                             f"(this report plus {evidence_periods - 1} prior)." +
                             (" Still early -- keep recommendations to the most significant "
                              "outliers only." if evidence_periods == 2 else ""))

    exclude = in_effect_values or set()
    baseline_rate = attribution.attributed_rate
    for dimension in result["dimensions_enabled"]:
        rows = _optimization_dimension_rows(attribution, dimension)
        if not rows:
            continue
        floor = _OPT_MIN_IMPRESSIONS_FLOOR[dimension]
        eligible = [r for r in rows if r.delivered_impressions >= floor
                   and (dimension, r.label) not in exclude]
        if not eligible:
            continue
        peer_avg_impressions = sum(r.delivered_impressions for r in eligible) / len(eligible)

        qualifying, limited = [], []
        for row in eligible:
            if not baseline_rate or row.attributed_rate >= baseline_rate:
                continue
            relative_diff = (baseline_rate - row.attributed_rate) / baseline_rate
            if relative_diff < _OPT_MATERIAL_RELATIVE_DIFF:
                continue
            entry = {
                "dimension": dimension, "value": row.label, "metric": "attributed_rate",
                "value_rate": row.attributed_rate, "campaign_rate": baseline_rate,
                "multiple": round(baseline_rate / row.attributed_rate, 1) if row.attributed_rate else None,
                "delivered_impressions": row.delivered_impressions,
            }
            if row.delivered_impressions < peer_avg_impressions * _OPT_ALREADY_LIMITED_IMPRESSION_RATIO:
                limited.append(entry)
            else:
                qualifying.append(entry)

        qualifying.sort(key=lambda e: e["value_rate"])
        cap = _optimization_cap(dimension, level, len(rows))
        result["candidates"].extend(qualifying[:cap])
        result["watch_list"].extend(qualifying[cap:])
        result["already_limited"].extend(limited)

    return result


# ---------------------------------------------------------------------------
# Accept/edit/decline, and the memory between monthly reports (ATTRIBUTION_
# REPORT_PLAN.md Phase 6, the optimization-sequence work). Pure derivation
# only, same reason as the rest of this file.
# ---------------------------------------------------------------------------

_OPT_DIMENSION_SUBJECT = {
    "zip": "ZIP {value}", "market": "the {value} market",
    "creative": 'the "{value}" creative', "day_of_week": "{value}",
    "publisher": "{value}",
}
_OPT_DIMENSION_PLURAL = {"zip": "ZIPs", "market": "markets", "creative": "creatives",
                         "day_of_week": "days", "publisher": "publishers"}
_OPT_DAY_NAMES = {"Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday", "Thu": "Thursday",
                  "Fri": "Friday", "Sat": "Saturday", "Sun": "Sunday"}


def describe_optimization_candidate(candidate):
    """A deterministic, Python-composed recommendation sentence for one
    candidate -- what a rep reviews on the accept/edit/decline checklist
    BEFORE any model call happens. Never the model's job: the rep has to
    be able to read and edit this pre-draft, and a candidate declined here
    must never reach the model at all (never merely be told not to use it).
    """
    dimension = candidate["dimension"]
    raw_value = candidate["value"]
    value = _OPT_DAY_NAMES.get(raw_value, raw_value) if dimension == "day_of_week" else raw_value
    subject = _OPT_DIMENSION_SUBJECT.get(dimension, "{value}").format(value=value)
    multiple = candidate.get("multiple")
    magnitude = f"{multiple}x below" if multiple else "well below"
    rate_pct = candidate["campaign_rate"] * 100
    impressions = candidate["delivered_impressions"]
    plural = _OPT_DIMENSION_PLURAL.get(dimension, dimension + "s")
    return (f"Reduce or remove delivery in {subject} ({magnitude} the campaign's "
            f"{rate_pct:.2f}% attributed rate, {impressions:,} impressions) and "
            f"reallocate to stronger-performing {plural}.")


def accepted_optimizations_from_report_json(report_json):
    """Every accepted-or-edited entry from a prior report's own stored
    `report_json["optimizations"]` log -- declined ones are excluded, they
    were never acted on and carry nothing to measure."""
    entries = (report_json or {}).get("optimizations") or []
    return [e for e in entries if e.get("decision") in ("accepted", "edited")]


_DIMENSION_TO_ATTRIBUTION_KEY = {
    "zip": "by_zip", "market": "by_market", "creative": "by_creative",
    "day_of_week": "by_day_of_week", "publisher": "by_channel",
}


def _find_dimension_row_in_dict(attribution_dict, dimension, value):
    key = _DIMENSION_TO_ATTRIBUTION_KEY.get(dimension)
    for row in (attribution_dict or {}).get(key) or []:
        if row.get("label") == value:
            return row
    return None


def _find_dimension_row_in_export(attribution, dimension, value):
    for row in _optimization_dimension_rows(attribution, dimension):
        if row.label == value:
            return row
    return None


def measure_optimization_effect(entry, then_attribution_dict, now_attribution):
    """Before/after for one accepted optimization -- "did it work," the
    most persuasive sentence a later report can carry. `entry` is a stored
    `report_json["optimizations"]` entry (accepted or edited); `then_
    attribution_dict` is the ORIGINATING report's own stored `report_
    json["attribution"]` dict (the export as it looked when the cut was
    decided -- a plain dict, `dataclasses.asdict()`'s own shape, never
    re-parsed from a file that may no longer exist locally); `now_
    attribution` is the CURRENT export's real `AttributionExport`.

    A value that no longer appears in the current export at all (delivery
    there may have gone to genuine zero) is reported as such via
    `found_now=False` and zero impressions -- never silently skipped, since
    "it's gone" is itself the finding for a real cut that worked.
    """
    dimension, value = entry["dimension"], entry["value"]
    then_row = _find_dimension_row_in_dict(then_attribution_dict, dimension, value)
    now_row = _find_dimension_row_in_export(now_attribution, dimension, value)
    then_key = _DIMENSION_TO_ATTRIBUTION_KEY.get(dimension)
    then_total = sum((r.get("delivered_impressions") or 0)
                     for r in (then_attribution_dict or {}).get(then_key) or [])
    now_total = sum(r.delivered_impressions
                    for r in _optimization_dimension_rows(now_attribution, dimension))
    then_impressions = (then_row or {}).get("delivered_impressions") or 0
    now_impressions = now_row.delivered_impressions if now_row else 0
    return {
        "dimension": dimension, "value": value,
        "final_text": entry.get("final_text"),
        "delivered_impressions_then": then_impressions,
        "delivered_impressions_now": now_impressions,
        "share_then": (then_impressions / then_total) if then_total else None,
        "share_now": (now_impressions / now_total) if now_total else None,
        "campaign_rate_then": (then_attribution_dict or {}).get("attributed_rate"),
        "campaign_rate_now": now_attribution.attributed_rate,
        "found_now": now_row is not None,
    }


def optimizations_in_effect(prior_reports, now_attribution):
    """(in_effect_values, in_effect_facts) for an ORDINARY (non-wrap)
    report. `prior_reports` is every prior logged report for this
    advertiser, OLDEST FIRST (app.py's own `prior_periods` convention) --
    only the MOST RECENT one's own accepted optimizations are "in effect"
    here; a report is a snapshot against last month, not the whole history
    (that's `optimization_history`, the wrap's own job, below).

    `in_effect_values` is a `{(dimension, value), ...}` set for
    `optimization_candidates()`'s own exclusion -- a value already acted on
    is measured, never re-discovered as a new candidate. `in_effect_facts`
    is the measured list this function computes unconditionally (see
    `optimization_candidates`'s own "level == none" docstring paragraph for
    why this never gates on the current report's level).
    """
    if not prior_reports:
        return set(), []
    latest_report_json = prior_reports[-1].get("report_json") or {}
    accepted = accepted_optimizations_from_report_json(latest_report_json)
    then_dict = latest_report_json.get("attribution")
    facts = [measure_optimization_effect(entry, then_dict, now_attribution) for entry in accepted]
    values = {(e["dimension"], e["value"]) for e in accepted}
    return values, facts


def optimization_history(prior_reports, now_attribution):
    """`facts["optimization_history"]` for a WRAP report -- every accepted
    optimization across EVERY prior report, in chronological order, each
    measured against the wrap's own (final) export -- the full journey,
    not just the last step. `prior_reports` oldest-first."""
    history = []
    for report_row in prior_reports:
        report_json = report_row.get("report_json") or {}
        headline = report_json.get("headline_facts") or {}
        then_dict = report_json.get("attribution")
        for entry in accepted_optimizations_from_report_json(report_json):
            measured = measure_optimization_effect(entry, then_dict, now_attribution)
            measured["period_start"] = headline.get("period_start")
            measured["period_end"] = headline.get("period_end")
            history.append(measured)
    return history


def pick_breakdown_dimension(attribution, dimension_override=None):
    """(dimension_label, rows) -- the ONE breakdown table this slide shows.
    Market wins whenever the export has more than one, and is never
    overridable -- there is nothing to judge when the data itself already
    has more than one market to show. Otherwise `dimension_override`
    ("Audience"/"Creative"), when it names a dimension this export actually
    has rows for, is Phase 4's Claude-driven judgment call replacing the
    heuristic below -- an override naming a dimension with nothing to show
    is ignored rather than fabricating rows that don't exist. With no
    override (Claude unreachable, or a rep hasn't drafted), falls back to
    the deterministic stand-in: creative only pre-empts audience when it's
    genuinely the standout -- at least two creatives ran AND the leader's
    attributed rate beats the runner-up's by 50%+.
    """
    if len(attribution.by_market) > 1:
        return "Market", attribution.by_market
    if dimension_override == "Audience" and attribution.by_audience:
        return "Audience", attribution.by_audience
    if dimension_override == "Creative" and len(attribution.by_creative) >= 2:
        return "Creative", attribution.by_creative
    creatives = sorted(attribution.by_creative, key=lambda r: r.attributed_rate, reverse=True)
    if len(creatives) >= 2 and creatives[0].attributed_rate >= creatives[1].attributed_rate * 1.5:
        return "Creative", creatives
    if attribution.by_audience:
        return "Audience", attribution.by_audience
    if creatives:
        return "Creative", creatives
    return "Market", attribution.by_market


def breakdown_rows(rows, limit=6):
    """`_attributed_raw` rides alongside the formatted `attributed` string so
    a caller building a chart from these SAME rows can't drift out of sync
    with what the table shows -- the ordering/limiting happens exactly
    once, here, not re-derived a second time from the unsorted input.

    `conv_rate` is always computed (cheap, and every `AttributionRow`
    already carries `conversion_rate`, real or a harmless zero) -- whether
    it's actually shown is the caller's call, made by whether "conv_rate"
    is in the `fields` list handed to `_fill_named_table`, not by anything
    computed here."""
    ordered = sorted(rows, key=lambda r: r.attributed_impressions, reverse=True)[:limit]
    return [{"label": r.label, "delivered": _int(r.delivered_impressions),
            "attributed": _int(r.attributed_impressions), "rate": _pct(r.attributed_rate),
            "conv_rate": _pct(r.conversion_rate),
            "_attributed_raw": r.attributed_impressions}
           for r in ordered]


def _row_fact(row):
    """An AttributionRow reduced to the plain numbers a facts payload can
    hand the model -- label plus every raw figure, never a pre-formatted
    string, so the model quotes a number and Python's own _int/_pct
    formatters (the same ones every deterministic slide uses) are what
    actually put it on the deck."""
    return {"label": row.label, "delivered_impressions": row.delivered_impressions,
            "attributed_impressions": row.attributed_impressions,
            "attributed_rate": row.attributed_rate}


def build_facts_payload(attribution, delivery, *, goals=None, notes=None, include_conversions=False,
                        budget=None, proposal_flight_label=None, proposal_geography_label=None,
                        ott=None, vertical=None, conversion_definition=None, prior_periods=None,
                        plan_vs_actual=None, optimizations=None, optimization_history_facts=None):
    """The complete, Python-computed facts payload Phase 4 hands the model,
    alongside the campaign goals and any rep notes -- app.py's
    `build_attribution_prompt` needs nothing else. Every value here is a raw
    number/string derived straight from the parsed export(s), or one of this
    module's own row/table helpers (`top_url_rows`, `top_zip_rows`,
    `intent_facts`, `breakdown_rows`) -- never a number invented for the
    model. The facts-only contract (ATTRIBUTION_REPORT_PLAN.md Phase 4):
    the model SELECTS which of these to mention and PHRASES the sentence;
    it does not compute a number that isn't sitting in this dict somewhere.

    `include_conversions` is the caller's call (same contract as
    `build_report_deck`'s own parameter of the same name) -- pass True only
    when BOTH `attribution.has_conversions` and the rep's own toggle are
    true. When False, "conversions" is None and every other section
    (`intent`, `top_pages`, `zip`) carries no conversion figures at all --
    the model must never even see them to accidentally mention, matching
    WAEPA's "no half-states" rule.

    `intent` is `intent_facts()`'s full, unfolded precision -- never
    `intent_summary_rows()`'s folded table -- because the URL slide's whole
    point is connecting a class the folded table would round away (an 8-visit
    lead-intent class) to a stated goal. `market`/`audience`/`creative` each
    carry every row, not just `breakdown_rows`'s capped/sorted table, so the
    model can cite a segment that isn't in the slide's own table too (e.g. a
    highlight bullet naming the top audience segment while the breakdown
    slide happens to be showing markets).

    `goals`/`notes` ride along in the same dict purely so one object is the
    complete input to one prompt call -- they are echoed back exactly as
    given, never interpreted here. `notes` is optional and may be empty:
    Phase 4's design rule is that a report can always be drafted from goals
    and computed facts alone; notes only add rep-supplied context on top.

    `headline`.`delivered_impressions` is OTT + live sports COMBINED
    whenever a sports block is present (`combined_headline_impressions`) --
    the one number the model is told is "impressions delivered." `live_sports`
    (None when the delivery file carries no sports block) carries the
    sports-only figures separately, so a drafted `live_sports_narrative` can
    still name the sports-specific numbers without recomputing anything.

    `budget` (ATTRIBUTION_REPORT_PLAN.md Phase 5) is the linked proposal's
    OWN full-flight total, in dollars -- there is no recap slot for it (a
    client already knows what they spent); it exists purely so Python can
    compute `facts["budget"]["cost_per_attributed_visit"]`/
    `["cost_per_conversion"]`, the derived figures WAEPA's own notes asked
    the reporting to show, as facts the model may cite like any other.
    `cost_per_conversion` is None unless `include_conversions` and the
    export genuinely has conversions -- the same "no half-states" rule the
    conversions layer already follows. `facts["budget"]` is None (never a
    half-filled dict) with no proposal linked, matching `delivery`/
    `live_sports`'s own None-when-absent convention.

    `proposal_flight_label`/`proposal_geography_label` (also Phase 5) ride
    into `facts["proposal"]` purely so the model can compare a rep's own
    note against the LINKED PROPOSAL's stated flight/geography, not just
    the export's computed facts -- the same "notes never override a fact,
    the disagreement is named in goal_alignment_notes instead" precedence
    rule this prompt already applies to every other fact here. Neither
    token is filled from the model's own kwargs; FLIGHT_LABEL/GEOGRAPHY_LABEL
    are always Python-derived (see `build_report_deck`'s own
    `geography_names_override`) -- this section exists only so a
    contradiction can be NOTICED, never so one could be introduced.

    `facts["response_profile"]` (`response_profile_facts`, ATTRIBUTION_
    REPORT_PLAN.md, 2026-09-10) is unconditional -- recency/referral/day-
    of-week were already parsed and never surfaced; this is what lets the
    existing highlight/takeaway rules cite them even before the new
    report:response_profile slide exists to show them directly.

    `ott` (also 2026-09-10, optional) is an already-parsed `attribution_
    import.OTTRetargetingExport` from a THIRD, separate export -- never
    part of the attribution/delivery pair. `facts["ott_retargeting"]` is
    None when it's not supplied, matching `delivery`/`live_sports`'s own
    convention.

    Highlights/Takeaways rework additions (all optional, all None/absent by
    default -- app.py supplies them once it has them):

    `vertical` (a plain string, the app's own vertical label or None) is
    stored as `facts["vertical"]` -- literally `"unknown"` when None, never
    guessed. `facts["benchmark"]` is computed from it here via
    `attribution_benchmarks.benchmark_facts` against this campaign's own
    `attributed_rate`/`attributed_unique_visitor_rate` -- callers never
    import `attribution_benchmarks` themselves. **The whole "benchmark"
    subtree must stay out of `app._attr_payload_numbers`'s traced-number
    set** -- the benchmark row's own percentage is context, never a
    citable campaign fact.

    `conversion_definition` (a rep-typed string, or None) rides straight
    into `facts["conversion_definition"]` -- what the export's conversion
    figure actually means for this client ("application starts"), so the
    model can name it instead of the generic word.

    `prior_periods` (a list of this same module's own `report_headline_
    facts()` dicts, oldest-first, or None) rides straight into
    `facts["prior_periods"]` -- the account's own trend, pulled from
    previously logged reports for the same advertiser (app.py's job; this
    module never touches the database).

    `plan_vs_actual` (`plan_vs_actual_facts()`'s own return, or None) rides
    straight into `facts["plan_vs_actual"]` -- gated entirely by the rep's
    own toggle in app.py, since this is never a thread on its own unless a
    stated goal names it (same gating as delivery metrics, below).

    `optimizations` (`optimization_candidates()`'s own return, or None) --
    the deterministic optimization engine (ATTRIBUTION_REPORT_PLAN.md Phase
    6). Computed in app.py from the rep's own level/dimension toggles and
    handed in here unchanged (already filtered to accepted/edited-only, and
    already carrying an `"in_effect"` key of `optimizations_in_effect()`'s
    own measured list, when app.py added one); None when the rep hasn't
    enabled it for this report. See `optimization_candidates`'s own
    docstring for the shape.

    `optimization_history_facts` (`optimization_history()`'s own return, or
    None) -- the WRAP-only full chain of every accepted optimization across
    the flight, each measured against THIS report's own export. Rides into
    `facts["optimization_history"]` unchanged; None for an ordinary
    (non-wrap) report, which uses `optimizations["in_effect"]` instead (one
    prior month, not the whole flight).
    """
    dimension, _rows = pick_breakdown_dimension(attribution)
    facts = {
        "goals": list(goals or []),
        "notes": (notes or "").strip(),
        "headline": {
            "delivered_impressions": combined_headline_impressions(attribution, delivery),
            "attributed_unique_visitors": attribution.attributed_unique_visitors,
            "attributed_unique_visitor_rate": attribution.attributed_unique_visitor_rate,
            "attributed_rate": attribution.attributed_rate,
        },
        "audience": {
            "top": (_row_fact(max(attribution.by_audience, key=lambda r: r.attributed_impressions))
                   if attribution.by_audience else None),
            "rows": [_row_fact(r) for r in attribution.by_audience],
        },
        "market": {
            "count": len(attribution.by_market),
            "top": (_row_fact(max(attribution.by_market, key=lambda r: r.attributed_impressions))
                   if attribution.by_market else None),
            "rows": [_row_fact(r) for r in attribution.by_market],
        },
        "creative": {
            "top": (_row_fact(max(attribution.by_creative, key=lambda r: r.attributed_rate))
                   if attribution.by_creative else None),
            "rows": [_row_fact(r) for r in attribution.by_creative],
        },
        "breakdown": {
            # None (not a value) when there's a genuine judgment call to
            # make -- Phase 4's own "breakdown_dimension" schema field only
            # means anything in that case. Forced when the export already
            # has more than one market: there is nothing to judge.
            "dimension_forced": dimension if len(attribution.by_market) > 1 else None,
            "audience_available": bool(attribution.by_audience),
            "creative_available": len(attribution.by_creative) >= 2,
        },
        "intent": intent_facts(attribution, include_conversions=include_conversions),
        "top_pages": top_url_rows(attribution, limit=8, include_conversions=include_conversions),
        "zip": {
            "baseline_rate": attribution.attributed_rate,
            # dropped_zips ignored here -- the model reads only real,
            # placed rows; the deck-side caller is what logs the drop.
            "rows": top_zip_rows(attribution, include_conversions=include_conversions)[0],
            # The ELIGIBILITY floor, not a fact about any zip -- ZIP_MIN_SHARE
            # decides which zips reached "rows" at all; it must never be
            # narrated back as if it were a real per-zip observation ("both
            # carrying more than 1% share" restates the selection rule, not
            # a finding). "_internal" keys are excluded from the facts-only
            # checker's traced-number set the same way "benchmark" is -- see
            # app._attr_payload_numbers.
            "_internal_min_share_pct": ZIP_MIN_SHARE,
        },
        "delivery": None,
        "live_sports": None,
        "conversions": ({
            "attributed": attribution.attributed_conversions,
            "sales_amount": attribution.sales_amount or None,
            "rate_of_attributed_impressions": (
                attribution.attributed_conversions / attribution.attributed_impressions
                if attribution.attributed_impressions else None),
        } if include_conversions and attribution.has_conversions else None),
        "budget": ({
            "total": budget,
            "cost_per_attributed_visit": (
                budget / attribution.attributed_unique_visitors
                if attribution.attributed_unique_visitors else None),
            "cost_per_conversion": (
                budget / attribution.attributed_conversions
                if include_conversions and attribution.has_conversions
                and attribution.attributed_conversions else None),
        } if budget else None),
        "proposal": ({
            "flight_label": proposal_flight_label or None,
            "geography_label": proposal_geography_label or None,
        } if (proposal_flight_label or proposal_geography_label) else None),
        "vertical": vertical or "unknown",
        "benchmark": attribution_benchmarks.benchmark_facts(
            vertical, attribution.attributed_rate, attribution.attributed_unique_visitor_rate),
        "conversion_definition": (conversion_definition or None),
        "prior_periods": list(prior_periods or []),
        "plan_vs_actual": plan_vs_actual,
        "optimizations": optimizations,
        "optimization_history": optimization_history_facts,
    }
    if delivery is not None:
        facts["delivery"] = {
            "delivered_impressions": delivery.delivered_impressions,
            "vcr": delivery.vcr,
            "frequency": delivery.frequency,
            "uniques": delivery.uniques,
            "ctv_share": delivery.ctv_share,
            "top_publishers": [{"name": name, "impressions": count}
                               for name, count, _pct in delivery.top_publishers],
            "breakdown_applies": delivery_breakdown_applies(delivery),
            "by_geo": [{"label": label, "impressions": count} for label, count in (delivery.by_geo or [])],
            "by_creative": [{"name": name, "impressions": count, "vcr": vcr}
                            for name, count, _length, _hours, vcr in (delivery.by_creative or [])],
        }
        if live_sports_applies(delivery):
            ls = delivery.live_sports
            top_events = sorted(ls.events, key=lambda e: e.delivered_impressions, reverse=True)
            facts["live_sports"] = {
                "package_type": ls.package_type,
                "delivered_geo": ls.delivered_geo,
                "delivered_impressions": ls.delivered_impressions,
                "vcr": ls.vcr,
                "flight_goal": ls.flight_goal,
                "pacing_note": f"{ls.delivered_impressions:,} of {ls.flight_goal:,}",
                "event_count": len(ls.events),
                "top_events": [{"event": e.event, "network": e.network,
                               "impressions": e.delivered_impressions}
                              for e in top_events[:5]],
                "by_network": [{"network": n, "impressions": c} for n, c in ls.by_network],
            }
    facts["response_profile"] = response_profile_facts(attribution)
    facts["ott_retargeting"] = ott_retargeting_facts(ott)
    return facts


# Highlights/Takeaways rework's own cap -- goal threads never lose a slot to
# a signal thread. Matches this slide's own long-standing "up to 4" rule.
_THREAD_SLIDE_CAP = 4


def distribute_threads(threads):
    """(highlight_bullets, takeaway_bullets, whats_next_bullets) -- Python's
    own deterministic distribution of the model's `threads` array (the
    Highlights/Takeaways rework, replacing three independently-drafted model
    outputs with one structure Python fans out). Each item in `threads` is
    `{"head", "anchor": "goal"|"signal", "goal_ref", "finding", "meaning",
    "action"}`; the model supplies goal threads FIRST, in its own inferred
    priority order, then up to two signal threads -- that ordering is what
    this function treats as the priority, not a separate field.

    **One selection, both slides (2026-09-12 real find, WAEPA):** the
    surviving threads are chosen ONCE, capped at `_THREAD_SLIDE_CAP` (4),
    and BOTH slides render from that same set -- never two independent
    caps. The earlier version built the highlight pool and the takeaway
    pool separately (filtering to "has a finding" / "has a meaning" BEFORE
    capping) and capped each on its own; when one thread happened to lack a
    finding and a different thread happened to lack a meaning, the two caps
    trimmed different threads off the end, and Highlights/Takeaways showed
    two DIFFERENT sets of four -- a goal thread (frequency) landed on
    Takeaways only while a signal thread (direct visits) landed on
    Highlights only, arguing different points on the two slides that are
    supposed to open and close the same story. Now: rank once (goal
    threads first in listed order, then signal threads in listed order),
    trim signal threads off the end first until at most 4 remain (a goal
    thread never loses its slot to a signal thread), and every surviving
    thread contributes a highlight (when it has a finding) AND a takeaway
    (when it has a meaning) from that one shared set. The only permitted
    asymmetry is the one designed in: a thread with a meaning but no
    finding is closing-only (takeaway, no highlight) -- never the reverse,
    since a highlight is only added here when a takeaway is added for the
    same thread too, which is what makes "every highlight has a takeaway"
    a mechanical guarantee rather than a hope about model compliance.

    - A takeaway is `(head, meaning)` for every SURVIVING thread whose
      `meaning` is non-empty. A thread with no meaning contributes nothing
      at all -- malformed model output, not a place to guess one.
    - A highlight is `(head, finding)` for every surviving thread that got
      a takeaway AND whose `finding` is non-empty -- a thread with no
      finding (a closing-only signal, e.g. "direct visits at 31%") simply
      doesn't contribute one.
    - A what's-next item is `action` for every surviving thread that got a
      takeaway and has one -- no cap of its own beyond the shared one
      ("what's-next takes every action" whose own thread made Takeaways;
      see the no-orphan-actions note below).

    Malformed input (not a list, or a thread missing `head`/`meaning`
    outright) degrades rather than raises -- this is model output, and a
    thread `apply_attr_draft` can't use is one fewer bullet, not a crash.
    """
    threads = [t for t in (threads or []) if isinstance(t, dict) and str(t.get("head") or "").strip()]
    goal_threads = [t for t in threads if t.get("anchor") == "goal"]
    signal_threads = [t for t in threads if t.get("anchor") != "goal"]
    ordered = goal_threads + signal_threads
    is_goal_flags = [True] * len(goal_threads) + [False] * len(signal_threads)

    # One cap, applied once, to the ranked thread list itself -- not to a
    # derived pairs list -- so both slides read from identical survivors.
    survivors = list(zip(ordered, is_goal_flags))
    while len(survivors) > _THREAD_SLIDE_CAP:
        # Drop the LAST signal thread; if none remain, the cap is being
        # asked to drop a goal thread -- shouldn't happen given the
        # prompt's own 3-4 goal + up to 2 signal guidance, but falls back
        # to dropping from the end rather than raising.
        signal_indices = [i for i, (_t, is_goal) in enumerate(survivors) if not is_goal]
        drop_index = signal_indices[-1] if signal_indices else len(survivors) - 1
        survivors.pop(drop_index)
    survivor_threads = [t for t, _is_goal in survivors]

    highlight_bullets, takeaway_triples = [], []
    for thread in survivor_threads:
        head = str(thread.get("head") or "").strip()
        finding = str(thread.get("finding") or "").strip()
        meaning = str(thread.get("meaning") or "").strip()
        action = str(thread.get("action") or "").strip()
        if not meaning:
            continue
        takeaway_triples.append((head, meaning, action))
        if finding:
            highlight_bullets.append((head, finding))

    takeaway_bullets = [(head, meaning) for head, meaning, _action in takeaway_triples]
    # No orphan actions (2026-09-11 finding): an action belongs to what's-
    # next ONLY when its own thread's meaning made the takeaways slide --
    # `takeaway_triples` already IS that set, since both are built from the
    # same `survivor_threads` pass above.
    whats_next = [action for _head, _meaning, action in takeaway_triples if action]
    return highlight_bullets, takeaway_bullets, whats_next


def report_headline_facts(attribution, delivery, include_conversions=False):
    """The compact per-report summary logged alongside every attribution
    report (`db.log_attribution_report`'s `report_json["headline_facts"]`)
    -- the foundation both this rework's own prior-period trend and Phase 6
    (Polk) build on. Deliberately small: a handful of scalars, not the raw
    export.

    `period_start`/`period_end` come from the export's OWN flight dates
    (`attribution.flight_start`/`flight_end`), never the report's upload
    date -- trend has to be ordered by when the campaign ran, not when the
    report was generated. `attributed_conversions` is None (never 0) unless
    both `include_conversions` and `attribution.has_conversions` are true --
    the same "no half-states" rule every other conversions surface in this
    module already follows. `top_intent_label`/`top_intent_share` come from
    `intent_facts`'s own top class (None when there are no classified visits
    at all).
    """
    intent = intent_facts(attribution)
    classes = intent.get("classes") or []
    top = classes[0] if classes else None
    return {
        "period_start": str(attribution.flight_start) if attribution.flight_start else None,
        "period_end": str(attribution.flight_end) if attribution.flight_end else None,
        "attributed_rate": attribution.attributed_rate,
        "attributed_unique_visitors": attribution.attributed_unique_visitors,
        "attributed_conversions": (attribution.attributed_conversions
                                  if include_conversions and attribution.has_conversions else None),
        "top_intent_label": top["label"] if top else None,
        "top_intent_share": top["share"] if top else None,
    }


def _normalize_geo_label(label):
    return " ".join(str(label or "").split()).lower()


def _geo_labels_match(plan_label, delivery_label):
    a, b = _normalize_geo_label(plan_label), _normalize_geo_label(delivery_label)
    if not a or not b:
        return False
    return a == b or a in b or b in a


def plan_vs_actual_facts(plan_rows, delivery_by_geo):
    """Best-effort planned-vs-delivered join, by normalized market label --
    `plan_rows` (a linked proposal's own media-plan rows, `{"geo",
    "impressions", ...}`, `app.linked_proposal_report_fields`'s `plan_rows`)
    against `delivery_by_geo` (`[(label, delivered_impressions), ...]`,
    `DeliveryExport.by_geo`).

    Market-label matching is inherently fuzzy -- a rep-typed/derived plan Geo
    string ("Washington, DC, Baltimore") and the Premion export's own DMA
    name ("Washington, DC") were never designed to agree byte-for-byte -- so
    this is a normalized-fold, substring-either-direction match, not an exact
    one, and a plan row that matches nothing simply isn't included in `rows`
    (it still counts against `total_markets`, just not `matched_count`).

    Returns None when `plan_rows` is empty (nothing to compare). Otherwise
    `{"rows": [{"label", "planned", "delivered", "pct_of_plan"}, ...matched
    only...], "matched_count", "total_markets", "totals": {"planned",
    "delivered", "pct_of_plan"}}`. `pct_of_plan` is None when `planned` is 0
    (never a fabricated infinity/zero).
    """
    plan_rows = plan_rows or []
    if not plan_rows:
        return None
    delivery_by_geo = list(delivery_by_geo or [])
    rows = []
    for plan_row in plan_rows:
        planned = plan_row.get("planned")
        if planned is None:
            continue
        label = plan_row.get("geo") or ""
        matched = [count for d_label, count in delivery_by_geo if _geo_labels_match(label, d_label)]
        if not matched:
            continue
        delivered = sum(matched)
        rows.append({
            "label": label,
            "planned": planned,
            "delivered": delivered,
            "pct_of_plan": (delivered / planned if planned else None),
        })
    total_planned = sum(r["planned"] for r in rows)
    total_delivered = sum(r["delivered"] for r in rows)
    return {
        "rows": rows,
        "matched_count": len(rows),
        "total_markets": len(plan_rows),
        "totals": {
            "planned": total_planned,
            "delivered": total_delivered,
            "pct_of_plan": (total_delivered / total_planned if total_planned else None),
        },
    }


def named_table_column_count(template_path, slide_key, table_name):
    """The real column count of `table_name` on the slide tagged `slide_key`
    in the report template at `template_path` -- lets a caller gate a UI
    control on whether the template has actually been widened yet (the same
    guard shape TopZipTable's own 5-column widen used), rather than offering
    a toggle that changes nothing on the rendered slide. None if the slide
    or table isn't present at all (an older template, or a typo'd name) --
    never raises, since this is a UI-gating query, not a fill operation."""
    prs = Presentation(template_path)
    keys = _slide_by_key(prs)
    if slide_key not in keys:
        return None
    shape = _shape_or_none(prs.slides[keys[slide_key]], table_name)
    if shape is None or not shape.has_table:
        return None
    return len(shape.table.columns)


def default_highlight_bullets(attribution, delivery):
    """Up to 4 (head, detail) pairs, grounded strictly in computed facts --
    deliberately plain, not client-ready prose. This is the FALLBACK used
    when no Claude draft is available (Claude unreachable, or a rep hasn't
    drafted) -- ATTRIBUTION_REPORT_PLAN.md Phase 4's `apply_attribution_
    draft` supplies `highlight_bullets` straight to `build_report_deck`
    instead, in the normal case."""
    out = []
    if attribution.by_audience:
        top = max(attribution.by_audience, key=lambda r: r.attributed_impressions)
        out.append((f"{_pct(top.attributed_rate)} attributed rate",
                    f"led by {top.label}, the top-performing audience segment."))
    if attribution.by_market and len(attribution.by_market) > 1:
        top = max(attribution.by_market, key=lambda r: r.attributed_impressions)
        out.append((f"{top.label.title()} led all markets",
                    f"with {_int(top.attributed_impressions)} attributed impressions."))
    out.append((f"{_int(attribution.attributed_unique_visitors)} unique visitors",
               f"attributed back to the campaign ({_pct(attribution.attributed_unique_visitor_rate)} of delivery)."))
    if delivery is not None:
        out.append((f"{_pct(delivery.vcr, 1)} video completion rate",
                    f"across {_int(delivery.delivered_impressions)} delivered impressions."))
    elif attribution.by_creative:
        top = max(attribution.by_creative, key=lambda r: r.attributed_rate)
        out.append((f"{top.label} was the top creative",
                    f"at a {_pct(top.attributed_rate)} attributed rate."))
    return out[:4]


def default_takeaway_bullets(attribution, delivery):
    """Same fallback role as default_highlight_bullets -- used only when no
    Claude draft is available."""
    out = [("Attribution confirmed real site engagement",
           f"{_int(attribution.attributed_impressions)} attributed impressions across the flight.")]
    dimension, rows = pick_breakdown_dimension(attribution)
    if rows:
        top = max(rows, key=lambda r: r.attributed_impressions)
        out.append((f"{dimension} concentration", f"{top.label} drove the strongest response."))
    if delivery is not None and delivery.top_publishers:
        top_pub = max(delivery.top_publishers, key=lambda p: p[1])
        out.append((f"{top_pub[0]} was the leading publisher",
                    f"by delivered impressions among {len(delivery.top_publishers)} publishers."))
    return out[:4]


# ---------------------------------------------------------------------------
# Deck assembly
# ---------------------------------------------------------------------------

def build_report_deck(template_path, attribution, delivery, output_path, *,
                      client_name=None, report_title="Website Attribution Report",
                      goals_bullets, whats_next_bullets, flight_label="",
                      audience_bullets=None, highlight_bullets=None,
                      takeaway_bullets=None, headline_notes=None,
                      narratives=None, breakdown_dimension_override=None,
                      geography_label_override=None, geography_names_override=None,
                      targeted_zips=None, include_conversions=False,
                      extra_deck_path=None, ott=None, plan_vs_actual=None):
    """Fill the report master deck at `template_path` and save to
    `output_path`. Returns (output_path, warnings) -- warnings is a list of plain-language strings
    from any table that still overflows its measured floor even after the
    row cap and the shrink-to-fit pass (see `_fill_delivery_recap`'s own
    STOPGAP-turned-real-fix note); empty in the ordinary case.

    This module never fetches a linked proposal itself -- app.py (Phase 5,
    ATTRIBUTION_REPORT_PLAN.md) reads `form_json` and hands the results
    through as keyword arguments here, same as it always has for goals/
    audience/what's-next. `flight_label` defaults to "" (no automatic
    campaign flight separate from the report period without a linked
    proposal) -- a linked proposal passes a real one from its own flight.
    `geography_names_override` (a list of market names, e.g. from
    `target_market_labels`) is Phase 5's geography source -- prefer it over
    `geography_label_override` (a plain string) when both are available; it
    also drives the recap's overflow-bullet collapse for 3+ markets, which
    a bare string override can't (see `geography_label_from_names`/
    `geography_overflow_bullet_from_names`). Pass None (never an empty
    list) when there's nothing to override with, so the export-derived
    default still applies. `targeted_zips` (an iterable of zip codes, e.g.
    the union of a linked proposal's `targeting_groups[].resolved_zips`) is
    Phase 5's targeted-vs-visitor zip overlay -- None (default) draws
    exactly the visitor-only choropleth Phase 3 always drew; a proposal
    with no resolved zips (avails_mode off, or no avails import ever ran --
    the common case for a hand-typed proposal) must produce an identical
    map, never a crash or an empty second series. See
    `targeting_map.render_choropleth`'s own docstring for the overlay
    mechanism.

    `extra_deck_path` (2026-09-08 -- the Auto-Sales Analyst integration) is
    an optional .pptx whose slides are appended WHOLESALE, in order, after
    every slide this function itself filled -- no parsing, no facts
    extracted from it, nothing absorbed into the payload; it rides along
    exactly as uploaded. Deliberately the LAST step, after every token fill
    above, so appended slides can never shift a `_slide_by_key` lookup this
    function still needs to make.

    `goals_bullets`/`whats_next_bullets` are REQUIRED -- no export signal
    produces them, so this raises MissingTokenError rather than defaulting
    if either is empty; a rep types them, or Claude drafts them from notes
    (ATTRIBUTION_REPORT_PLAN.md Phase 4, `app.apply_attribution_draft`).
    `audience_bullets`/`highlight_bullets`/`takeaway_bullets` default to
    this module's own computed facts when not overridden; `headline_notes`
    (a dict of the three *_HEADLINE_NOTE tokens, keyed "attribution"/"url"/
    "zip") and `narratives` (a dict of narrative-sentence tokens, keyed
    "attribution"/"delivery"/"delivery_breakdown"/"url_intent"/"zip"/
    "live_sports" -- the last only meaningful when the delivery file
    carries a sports block -- plus "response_profile" and
    "ott_retargeting", v0_6) likewise default to a plain computed sentence
    per slide when not supplied. `ott` (an `attribution_import.
    OTTRetargetingExport`, default None) drives report:ott_retargeting the
    same way `delivery` drives the delivery-set slides -- None drops the
    slide outright when the template has it; a real export fills it. `breakdown_dimension_override` ("Audience"/"Creative") is
    `pick_breakdown_dimension`'s own override parameter, threaded through
    unchanged. `include_conversions` (default False -- WAEPA's own "no
    half-states" rule) fills the highlights slide's fourth tile, adds a
    conv_rate/converted column to the breakdown and URL-report tables, and
    is the caller's responsibility to set true only when BOTH
    `attribution.has_conversions` and the rep's own toggle are true --
    this module never reads `has_conversions` itself, so a caller can
    always override it (a test forcing the column on against a template
    that doesn't have conversions data, for instance).

    `plan_vs_actual` (Highlights/Takeaways rework, `plan_vs_actual_facts()`'s
    own return or None) is the caller's own toggle -- app.py is responsible
    for gating it on the template actually having room (`named_table_
    column_count`) before ever passing a non-None value; this module just
    fills whatever it's handed, gracefully degrading via `_fill_named_
    table`'s own column-count check if it's passed against a narrower
    template anyway (a test doing so deliberately, say).
    """
    if not goals_bullets:
        raise MissingTokenError("report:recap/GOALS_BULLETS: no goals were supplied -- "
                                "type them in, or draft from notes (Phase 4).")
    if not whats_next_bullets:
        raise MissingTokenError("report:takeaways/WHATS_NEXT_BULLETS: no what's-next items "
                                "were supplied -- type them in, or draft from notes (Phase 4).")

    prs = Presentation(template_path)
    keys = _slide_by_key(prs)
    for required in ("report:recap", "report:highlights", "report:attribution_breakdown",
                    "report:response_profile", "report:url_report", "report:zip_analysis",
                    "report:takeaways"):
        if required not in keys:
            raise MissingTokenError(f"template is missing a slide tagged key: {required}")
    for required in ("report:delivery_recap", "report:delivery_breakdown"):
        if required not in keys:
            raise MissingTokenError(f"template is missing a slide tagged key: {required}")

    # Standalone slides (v0_9: report:summary, report:case_study) are never
    # part of the normal multi-slide report -- they're their OWN separate
    # single-slide deliverables (build_single_slide), built on demand from
    # the same stored threads/facts. Dropped unconditionally, same highest-
    # index-first reasoning the delivery-set drop below uses.
    standalone_indices = sorted((i for i, slide in enumerate(prs.slides) if _is_standalone(slide)),
                                reverse=True)
    for index in standalone_indices:
        assembly.delete_slide(prs, index)
    if standalone_indices:
        keys = _slide_by_key(prs)  # indices shifted

    # The delivery SET -- every slide whose notes carry `delivery_set: true`
    # -- is dropped wholesale when no delivery file was uploaded. Driven off
    # the marker in the deck rather than a hardcoded key list, so a third
    # delivery slide added to a future template needs no code change here.
    # Dropped highest-index-first: deleting a slide shifts every index after
    # it, and collecting the indices up front then deleting low-to-high
    # would delete the wrong slides.
    # Every delivery-set slide is dropped by ONE rule -- a per-slide
    # applicability check, called uniformly whether `delivery` is None or
    # not. `delivery_breakdown_applies`/`live_sports_applies` already treat
    # `delivery is None` as "doesn't apply" on their own, so no separate
    # branch is needed for that case at all. This used to be TWO
    # independently-hardcoded lists -- one for `delivery is None`, one for
    # real delivery -- and that shape is exactly what caused a real bug:
    # report:live_sports was added to the second list when it was built
    # and never carried over to the first, so a no-delivery report left
    # the slide in the deck completely unfilled, every {{SPORTS_...}}
    # token still literal (caught 2026-09-10 rendering a real WAEPA
    # report). Matt's own ruling on the fix: any list of slide keys that
    # exists in two places is stale the moment the first one grows -- so
    # there is now exactly one list, this dict, and a fourth delivery-set
    # slide needs one new entry here, never two.
    delivery_set_applies = {
        "report:delivery_recap": lambda d: d is not None,
        "report:delivery_breakdown": delivery_breakdown_applies,
        "report:live_sports": live_sports_applies,
    }
    drop_keys = [key for key, applies in delivery_set_applies.items() if not applies(delivery)]
    # report:live_sports is OPTIONAL in the template -- it doesn't exist at
    # all in v0_4 and earlier, and most delivery files never carry a sports
    # block even once the template does. Only ever act on it when it's
    # actually present, the one exception to every other key in this list
    # being required (checked above).
    drop_keys = [k for k in drop_keys if k != "report:live_sports" or "report:live_sports" in keys]
    for key in drop_keys:
        if not _is_delivery_set(prs.slides[keys[key]]):
            raise MissingTokenError(
                f"{key} is missing its `delivery_set: true` notes line -- "
                f"can't tell whether it's safe to drop.")
    for index in sorted((keys[k] for k in drop_keys), reverse=True):
        assembly.delete_slide(prs, index)
    if drop_keys:
        keys = _slide_by_key(prs)  # indices shifted

    # report:ott_retargeting is independent of the delivery set -- driven
    # entirely by whether an OTT retargeting export was uploaded at all,
    # per Matt's own instruction that this slide is explicitly NOT
    # delivery_set. Optional in the template the same way report:live_sports
    # is (absent from every template before v0_6), so only ever act on it
    # when it's actually present.
    if ott is None and "report:ott_retargeting" in keys:
        assembly.delete_slide(prs, keys["report:ott_retargeting"])
        keys = _slide_by_key(prs)  # indices shifted

    narratives = narratives or {}
    _fill_recap(prs.slides[keys["report:recap"]], attribution, client_name, report_title,
               goals_bullets, audience_bullets, flight_label,
               geography_label_override=geography_label_override,
               geography_names_override=geography_names_override)
    _fill_highlights(prs.slides[keys["report:highlights"]], attribution, delivery, highlight_bullets,
                     include_conversions=include_conversions)
    warnings = []
    if delivery is not None:
        warnings += _fill_delivery_recap(prs.slides[keys["report:delivery_recap"]], delivery,
                                         narrative_override=narratives.get("delivery"),
                                         plan_vs_actual=plan_vs_actual)
    if delivery is not None and delivery_breakdown_applies(delivery):
        warnings += _fill_delivery_breakdown(
            prs.slides[keys["report:delivery_breakdown"]], delivery,
            narrative_override=narratives.get("delivery_breakdown"),
            plan_vs_actual=plan_vs_actual)
    if delivery is not None and live_sports_applies(delivery) and "report:live_sports" in keys:
        warnings += _fill_live_sports(
            prs.slides[keys["report:live_sports"]], delivery.live_sports,
            narrative_override=narratives.get("live_sports"))
    warnings += _fill_attribution_breakdown(
        prs.slides[keys["report:attribution_breakdown"]], attribution,
        (headline_notes or {}).get("attribution"),
        narrative_override=narratives.get("attribution"),
        dimension_override=breakdown_dimension_override,
        include_conversions=include_conversions)
    warnings += _fill_response_profile(prs.slides[keys["report:response_profile"]], attribution,
                                       narrative_override=narratives.get("response_profile"))
    warnings += _fill_url_report(prs.slides[keys["report:url_report"]], attribution,
                                 (headline_notes or {}).get("url"),
                                 narrative_override=narratives.get("url_intent"),
                                 include_conversions=include_conversions)
    warnings += _fill_zip_analysis(prs.slides[keys["report:zip_analysis"]], attribution,
                                   (headline_notes or {}).get("zip"),
                                   narrative_override=narratives.get("zip"),
                                   targeted_zips=targeted_zips)
    if ott is not None and "report:ott_retargeting" in keys:
        warnings += _fill_ott_retargeting(prs.slides[keys["report:ott_retargeting"]], ott,
                                          narrative_override=narratives.get("ott_retargeting"))
    _fill_takeaways(prs.slides[keys["report:takeaways"]], attribution, delivery,
                   takeaway_bullets, whats_next_bullets)

    if extra_deck_path:
        append_slide_deck(prs, extra_deck_path)

    prs.save(output_path)
    return output_path, [w for w in warnings if w]


def build_single_slide(template_path, key, output_path, fill_fn):
    """Keep only the slide tagged `key`, delete every other slide, fill the
    survivor via `fill_fn(slide)`, save. `assembly.build_avails_only_deck`'s
    exact shape (one standalone slide, kept and filled, nothing else in the
    package) -- used here for report:summary and report:case_study (v0_9),
    both tagged `standalone: true` so the ordinary multi-slide report build
    drops them, and both built on their own through this function instead.

    `fill_fn` takes the one surviving slide and does the actual token-
    filling -- kept as a plain callback rather than this function knowing
    about summary/case-study tokens itself, so a third standalone slide
    (a future case-study variant, framework's own "brand lift, Arrivalist,
    sales" list) needs only its own fill function, never a change here.
    """
    prs = Presentation(template_path)
    keys = _slide_by_key(prs)
    if key not in keys:
        raise MissingTokenError(f"template is missing a slide tagged key: {key}")
    keep_index = keys[key]
    for index in range(len(prs.slides._sldIdLst) - 1, -1, -1):
        if index != keep_index:
            assembly.delete_slide(prs, index)
    fill_fn(prs.slides[0])
    prs.save(output_path)
    return output_path


def append_slide_deck(prs, path):
    """Append every slide of the .pptx at `path`, in order, to the END of
    `prs` -- an external deck (the Auto-Sales Analyst's own summary, image
    charts, no native chart parts -- confirmed against a real 8-slide
    export) grafted in wholesale, no picking, no placement choice. Reuses
    assembly.py's own cross-deck copy primitives (`copy_slide_into`/
    `ImportCache`) -- the same mechanism case studies and slide vault
    entries already use in the proposal builder, and the same reason:
    rId remapping and part importing across two different python-pptx
    packages is real, load-bearing work, not something to reimplement a
    third time. Returns the number of slides appended.
    """
    cache = assembly.ImportCache(prs)
    source = Presentation(path)
    position = len(prs.slides._sldIdLst)
    count = len(source.slides._sldIdLst)
    for index in range(count):
        assembly.copy_slide_into(path, index, prs, position=position, cache=cache)
        position += 1
    return count


def _shape(slide, name):
    found = assembly._find_shape_by_name(slide.shapes, name)
    if found is None:
        raise MissingTokenError(f"named shape '{name}' not found on this slide")
    return found


def _delete_named_shapes(slide, *names):
    """Remove named shapes outright -- used when a tile's value has
    nothing to show (FLIGHT_LABEL blank in Phase 3), mirroring ChartRegion/
    MapRegion's own "delete both shapes when there's nothing to place"
    convention rather than leaving an empty-but-visible box."""
    for name in names:
        shape = assembly._find_shape_by_name(slide.shapes, name)
        if shape is not None:
            shape._element.getparent().remove(shape._element)


# The recap's three tiles, left to right. Generic list, not Flight-specific
# -- _reflow_tile_row below treats any subset going blank the same way, per
# Matt's 2026-09-06 ruling, even though FLIGHT_LABEL is the only one that
# can actually be blank today (REPORT_PERIOD_LABEL/GEOGRAPHY_LABEL both
# raise MissingTokenError rather than resolving to nothing).
_RECAP_TILE_ROW = ("ReportPeriodTile", "FlightTile", "GeographyTile")
# The delivery recap's own five-tile row (v0_3). Same mechanism, different
# names -- which is the point of `tile_names` being a parameter rather than
# this module knowing about one hardcoded row.
_DELIVERY_TILE_ROW = ("DeliveredTile", "VcrTile", "FrequencyTile", "UniquesTile",
                      "CtvShareTile")
# The highlights slide's own tile row (WAEPA conversions layer) -- SAME
# mechanism, and requires the template's existing three tiles to gain these
# names too (they predate this convention -- see ATTRIBUTION_REPORT_PLAN.md's
# conversions section for the exact rename spec). Until that rename lands,
# `_shape_or_none` finds none of these four by name, `_reflow_tile_row`'s own
# `if not present: return` makes the reflow call below a safe no-op, and the
# three original (unnamed) tiles keep filling exactly as they did before this
# was added -- this module never assumes the template has caught up.
_HIGHLIGHTS_TILE_ROW = ("ImpressionsTile", "VisitorsTile", "RateTile", "ConversionsTile")


def _reflow_tile_row(slide, blank_tiles, tile_names=_RECAP_TILE_ROW):
    """When one or more of `_RECAP_TILE_ROW`'s tiles has nothing to show,
    redistribute the SURVIVING tiles evenly across the row's original span
    -- 2026-09-06 ruling: a deleted tile leaves a gap, not a hole. Two
    survivors become half-width each (three become a third each, and so
    on), with the SAME gap preserved between them that the template's own
    tiles use -- measured from the template, never assumed, so a future
    template with different spacing still reflows correctly. Each
    survivor's Value/Label children are resized in lockstep, preserving
    THEIR OWN side padding relative to their tile (also measured, not
    assumed). Deletes the blank tile(s)' own three shapes each afterward,
    since their original geometry is what "the row's span" is measured
    against, and reflowing has to happen before they're gone.

    Generic over how many/which tiles are blank -- not Flight-specific,
    even though only `FLIGHT_LABEL` can actually be blank today.
    """
    present = [(name, _shape_or_none(slide, name)) for name in tile_names]
    present = [(name, shape) for name, shape in present if shape is not None]
    if not present:
        return
    survivors = [(name, shape) for name, shape in present if name not in blank_tiles]
    if len(survivors) == len(present):
        return  # nothing blank -- nothing to reflow

    ordered = sorted(present, key=lambda t: t[1].left)
    row_left = ordered[0][1].left
    row_right = max(shape.left + shape.width for _n, shape in present)
    gap = (ordered[1][1].left - (ordered[0][1].left + ordered[0][1].width)
          if len(ordered) > 1 else Emu(0))

    if survivors:
        ordered_survivors = sorted(survivors, key=lambda t: t[1].left)
        count = len(ordered_survivors)
        new_width = Emu(int((row_right - row_left - gap * (count - 1)) / count))
        left = row_left
        for name, tile_shape in ordered_survivors:
            value_shape = _shape_or_none(slide, f"{name}Value")
            label_shape = _shape_or_none(slide, f"{name}Label")
            # Measured from the ORIGINAL tile position, before it moves.
            value_pad_left = (value_shape.left - tile_shape.left) if value_shape else Emu(0)
            value_pad_total = (tile_shape.width - value_shape.width) if value_shape else Emu(0)
            label_pad_left = (label_shape.left - tile_shape.left) if label_shape else Emu(0)
            label_pad_total = (tile_shape.width - label_shape.width) if label_shape else Emu(0)
            tile_shape.left = left
            tile_shape.width = new_width
            if value_shape is not None:
                value_shape.left = Emu(int(left) + int(value_pad_left))
                value_shape.width = Emu(int(new_width) - int(value_pad_total))
            if label_shape is not None:
                label_shape.left = Emu(int(left) + int(label_pad_left))
                label_shape.width = Emu(int(new_width) - int(label_pad_total))
            left = Emu(int(left) + int(new_width) + int(gap))

    for name in blank_tiles:
        _delete_named_shapes(slide, name, f"{name}Value", f"{name}Label")


def _shape_or_none(slide, name):
    return assembly._find_shape_by_name(slide.shapes, name)


def _fill_tokens(slide, values):
    """Sweep every text frame on the slide -- including table cells, since
    BREAKDOWN_DIMENSION_LABEL sits in BreakdownTable's own HEADER row, not
    a separate text box."""
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape.has_text_frame:
            assembly._replace_tokens_in_text_frame(shape.text_frame, values)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    assembly._replace_tokens_in_text_frame(cell.text_frame, values)


def _set_cell_text(cell, value):
    """Write into a template table cell that carries NO token -- every named
    table here has exactly one token, in its data row's first cell; every
    other cell in that row is a genuinely empty paragraph (no runs at all,
    confirmed against the real template), so there's no existing run to
    replace via the usual token-swap path. Writes into the first existing
    run if the clone already has one (a cell cloned from a row that already
    got a value written into it), else adds one -- never `cell.text = ...`,
    which would blow away whatever paragraph-level formatting the template
    cell carries."""
    para = cell.text_frame.paragraphs[0]
    if para.runs:
        para.runs[0].text = str(value)
        for extra in para.runs[1:]:
            extra.text = ""
    else:
        para.add_run().text = str(value)


def _fill_named_table(slide, table_name, token, rows, fields, full_fields=None):
    """Fill a repeating table section by NAME -- required here because the
    delivery recap slide carries two tables (TopPublishersTable,
    CreativeTable), so "the first table on the slide"
    (assembly._find_table_shape's own rule, and assembly.fill_table_rows
    which relies on it) can't disambiguate. Row shape matches the real
    template exactly (REPORT_MASTER_README.md): the data row's first cell
    carries `token`; every other cell is empty and is written into
    directly, positionally, via `fields[1:]`.

    `rows` is a list of dicts; `fields` is the column order, `fields[0]`
    naming which dict key fills the tokened first cell.

    `full_fields` (2026-09-13 ruling, superseding the earlier "blank, never
    '--'" one): the table's own MAXIMAL column-to-field mapping, in real
    physical order -- which is not always `fields` with optional ones
    appended at the end (DeliveryByGeoTable interleaves Planned/% of plan
    between the two required columns). When the real template still has
    exactly `len(full_fields)` columns, any field present in `full_fields`
    but absent from `fields` names a column with nothing to show for THIS
    report, and it is REMOVED ENTIRELY -- gridCol and every row's tc
    deleted (`assembly.remove_table_column`), freed width given to column
    0 (the label column), same precedent as `assembly._fill_broadcast_
    slide`'s own empty-Days-column removal. An empty column on a client
    slide reads as missing data, not as "nothing to show here" -- a
    3-column table renders as 3 columns; a template carrying 5 is a
    capacity, not a promise. A template that DOESN'T match `full_fields`'s
    column count (an older, narrower one) is left alone here and falls
    through to the graceful column_warning below unchanged, so a stray
    pre-widened template stays exactly as safe as it always was.

    2026-09-06: after filling, runs the SAME measured shrink-to-fit pass
    the proposal master's avails table uses (`assembly.condense_avails_
    table`, given this table by name since it can't rely on "the only
    table on the slide") -- the row COUNT is a fixed design cap chosen by
    the caller (`_fill_delivery_recap`'s PUBLISHER_ROW_CAP/CREATIVE_ROW_CAP),
    and this is the safety net underneath it, not a substitute for it.
    Returns the overflow warning string (or None) if even the minimum font
    still doesn't fit.
    """
    table_shape = _shape(slide, table_name)
    table = table_shape.table
    tbl = table._tbl
    if full_fields and len(table.columns) == len(full_fields):
        drop_indices = sorted(
            (i for i, f in enumerate(full_fields) if f not in fields), reverse=True)
        if drop_indices:
            freed = sum(int(table.columns[i].width) for i in drop_indices)
            for i in drop_indices:
                assembly.remove_table_column(table, i)
            table.columns[0].width = Emu(int(table.columns[0].width) + freed)
    # A field list longer than the table has columns would silently drop the
    # tail (zip() stops at the shorter side) -- which is exactly how a
    # client-facing table loses a column nobody notices is missing. Report
    # it instead. This is expected and benign while a template revision is
    # in flight (the 5-column TopZipTable arrives with v0_3; today's
    # template has 3), which is why it warns rather than raising -- a
    # half-filled table is still a better interim state than no deck.
    column_warning = None
    if len(fields) > len(table.columns):
        dropped = ", ".join(fields[len(table.columns):])
        column_warning = (
            f"{table_name} has {len(table.columns)} columns but {len(fields)} were supplied -- "
            f"not showing: {dropped}. (Expected until the template revision that widens it "
            f"lands; see ATTRIBUTION_REPORT_PLAN.md.)")
    template_tr = tbl.tr_lst[1]  # header row is row 0, template data row is row 1
    trs = [template_tr]
    last_tr = template_tr
    for _ in rows[1:]:
        new_tr = assembly.clone_table_row(table, last_tr)
        trs.append(new_tr)
        last_tr = new_tr
    for tr, row_data in zip(trs, rows):
        row_index = tbl.tr_lst.index(tr)
        cells = list(table.rows[row_index].cells)
        assembly._replace_tokens_in_text_frame(cells[0].text_frame, {token: str(row_data[fields[0]])})
        assembly._tighten_cell_paragraphs(cells[0])
        for cell, field in zip(cells[1:], fields[1:]):
            _set_cell_text(cell, row_data[field])
            assembly._tighten_cell_paragraphs(cell)
    fit_warning = assembly.condense_avails_table(slide, len(rows), table_shape=table_shape)
    # Both are real, independent facts about this table; neither should mask
    # the other, so the caller gets whichever fired (or both).
    return [w for w in (column_warning, fit_warning) if w]


def _fill_bullets(slide, token, items):
    text_frame = assembly._find_text_frame_with_token(slide.shapes, token)
    if text_frame is None:
        raise MissingTokenError(f"token {{{{{token}}}}} not found on this slide")
    assembly.fill_bullet_list(text_frame, token, [str(i) for i in items])


def _fill_head_detail_bullets(slide, box_name, items, max_items):
    """items: [(head, detail), ...], up to max_items -- extra template
    paragraphs beyond len(items) are deleted whole, per README's own
    "delete unused paragraphs whole" rule for HIGHLIGHTBullets/
    TAKEAWAYBullets."""
    box = _shape(slide, box_name)
    tf = box.text_frame
    paragraphs = list(tf.paragraphs)
    if len(items) > len(paragraphs):
        raise MissingTokenError(f"{box_name} has {len(paragraphs)} template paragraphs, "
                                f"but {len(items)} items were supplied")
    for para, (head, detail) in zip(paragraphs, items):
        for run in para.runs:
            # lambda replacements -- re.sub treats a plain string replacement
            # as a backreference template ("\1" etc.), which a rep-typed
            # head/detail sentence could contain by accident.
            run.text = re.sub(r"\{\{\w+_HEAD_\d\}\}", lambda _m: head, run.text)
            run.text = re.sub(r"\{\{\w+_DETAIL_\d\}\}", lambda _m: detail, run.text)
    for para in paragraphs[len(items):]:
        para._p.getparent().remove(para._p)


def _fill_recap(slide, attribution, client_name, report_title, goals_bullets, audience_bullets,
                flight_label, geography_label_override=None, geography_names_override=None):
    period = _date_range_label(attribution.flight_start, attribution.flight_end)
    if period is None:
        raise MissingTokenError("report:recap/REPORT_PERIOD_LABEL: the export carries no "
                                "weekly/monthly trend tab to derive a period from -- "
                                f"warnings: {attribution.warnings}")
    # `geography_names_override` (Phase 5's target-market-derived list) wins
    # over `geography_label_override` (a bare string, Phase 4's own
    # plumbing) which wins over the export's own derivation -- see
    # `build_report_deck`'s docstring for why a list is preferred when both
    # are available (it also drives the overflow bullet below).
    if geography_names_override:
        geo = geography_label_from_names(geography_names_override)
    else:
        geo = geography_label_override or geography_label(attribution)
    if geo is None:
        raise MissingTokenError("report:recap/GEOGRAPHY_LABEL: no by-market breakdown, "
                                "station-pixel hint or zip data to derive a market from")
    name = client_name or attribution.client_name
    if not name:
        raise MissingTokenError("report:recap/CLIENT_NAME: not in the export and none supplied "
                                "-- confirm the advertiser first")
    audiences = list(audience_bullets or [row.label for row in attribution.by_audience[:4]])
    # 2026-09-06: when the Geography tile collapsed a 3+-market list down to
    # "N markets," the full list isn't lost -- it lands here, as its own
    # Audience-bullets line, never as a truncated join back in the tile.
    # Drawn from the SAME name source as the tile itself -- the override
    # list when one was given, never a stale export-derived overflow next
    # to an overridden tile.
    overflow = (geography_overflow_bullet_from_names(geography_names_override)
               if geography_names_override else geography_overflow_bullet(attribution))
    if overflow:
        audiences.append(overflow)
    if not audiences:
        raise MissingTokenError("report:recap/AUDIENCE_BULLETS: the export has no audience "
                                "breakdown and none was supplied")
    _fill_tokens(slide, {
        "CLIENT_NAME": name, "REPORT_TITLE": report_title,
        "REPORT_PERIOD_LABEL": period, "GEOGRAPHY_LABEL": geo,
    })
    if flight_label:
        _fill_tokens(slide, {"FLIGHT_LABEL": flight_label})
    else:
        # 2026-09-06: a blank value with the tile still drawn read as a
        # broken box, not an intentional omission. Reflow, don't just
        # delete -- Matt's ruling: the survivors redistribute evenly across
        # the row's original span rather than leaving a gap, since Flight
        # is blank on EVERY Phase 3 report (this is the default look, not
        # an edge case) until Phase 5 links a proposal.
        _reflow_tile_row(slide, {"FlightTile"})
    _fill_bullets(slide, "GOALS_BULLETS", goals_bullets)
    _fill_bullets(slide, "AUDIENCE_BULLETS", audiences)


def _fill_highlights(slide, attribution, delivery, highlight_bullets, include_conversions=False):
    """`include_conversions` fills a fourth tile, {{HEADLINE_CONVERSIONS}}
    -- the count, with the sales amount appended only when it's genuinely
    > 0 (WAEPA's own $0 case is real: a count without a value, and a tile
    reading "37 ($0)" would be a confident wrong claim). When it's False
    (conversions absent, or the rep's own toggle turned them off), the
    fourth tile is reflowed away -- see _HIGHLIGHTS_TILE_ROW's own note on
    why that reflow is a safe no-op on a template that hasn't been updated
    to name its tiles yet.

    HEADLINE_IMPRESSIONS is OTT + live sports COMBINED when a sports block
    is present -- see combined_headline_impressions's own docstring for the
    2026-09-08 finding this settles (a client who bought both bought one
    campaign; sports stays visibly broken out on its own report:live_sports
    slide, never silently folded in with nothing to show for it)."""
    headline_impressions = combined_headline_impressions(attribution, delivery)
    tokens = {
        "HEADLINE_IMPRESSIONS": _int(headline_impressions),
        "HEADLINE_UNIQUE_VISITORS": _int(attribution.attributed_unique_visitors),
        "HEADLINE_ATTRIBUTED_RATE": _pct(attribution.attributed_rate),
    }
    if include_conversions:
        value = _int(attribution.attributed_conversions)
        if attribution.sales_amount:
            value = f"{value} ({_money(attribution.sales_amount)})"
        tokens["HEADLINE_CONVERSIONS"] = value
    _fill_tokens(slide, tokens)
    if not include_conversions:
        _reflow_tile_row(slide, {"ConversionsTile"}, tile_names=_HIGHLIGHTS_TILE_ROW)
    items = highlight_bullets or default_highlight_bullets(attribution, delivery)
    _fill_head_detail_bullets(slide, "HIGHLIGHTBullets", items, max_items=4)
    _shrink_bullet_box_to_fit(slide, "HIGHLIGHTBullets")


TOP_PUBLISHERS_ROW_CAP = 5  # 2026-09-06: a design rule (what a client reads), not a fit guess
TOP_CREATIVES_ROW_CAP = 3


def _fill_delivery_recap(slide, delivery, narrative_override=None, plan_vs_actual=None):
    """v0_3: five tiles (CTV share joins the four originals), the publisher
    table alone in the left column, and the chart region now carrying the
    DAYPART breakdown -- the publisher bar chart was dropped because it and
    the publisher table said the same thing, and the table carries VCR too.
    CreativeTable moved off this slide to report:delivery_breakdown.

    `plan_vs_actual` (Highlights/Takeaways rework, `plan_vs_actual_facts()`'s
    own return or None) fills v0_7's named `PlanVsActualNote` shape
    (`{{PLAN_VS_ACTUAL_NOTE}}`, sitting under `DeliveryNarrative`) with a
    one-line planned-vs-delivered summary -- the fallback for markets that
    didn't match the geo table's own row-level augmentation (see
    `_fill_delivery_breakdown`). **The shape is DELETED, not left with an
    unfilled token, whenever there's no note to show** (`plan_vs_actual` is
    None, or genuinely has nothing to say) -- Matt's own v0_7 instruction,
    so a blank line never renders. A template older than v0_7 simply
    doesn't have this shape at all, so `_delete_named_shapes` no-ops on it.
    """
    _fill_tokens(slide, {
        "DELIVERED_IMPRESSIONS": _int(delivery.delivered_impressions),
        "VCR": _pct(delivery.vcr, 1),
        "FREQUENCY": f"{delivery.frequency:.1f}",
        "UNIQUES": _int(delivery.uniques),
        "DELIVERY_NARRATIVE": narrative_override or (
            f"Delivered {_int(delivery.delivered_impressions)} impressions "
            f"at a {_pct(delivery.vcr, 1)} completion rate."),
    })
    if plan_vs_actual:
        totals = plan_vs_actual["totals"]
        pct = (f"{totals['pct_of_plan'] * 100:.0f}% of plan"
              if totals["pct_of_plan"] is not None else "plan not stated")
        _fill_tokens(slide, {
            "PLAN_VS_ACTUAL_NOTE": (
                f"Planned vs. delivered: {_int(totals['planned'])} planned, "
                f"{_int(totals['delivered'])} delivered ({pct}) across "
                f"{plan_vs_actual['matched_count']} of {plan_vs_actual['total_markets']} "
                f"markets matched."),
        })
    else:
        _delete_named_shapes(slide, "PlanVsActualNote")
    # ctv_share is None (never 0.0) when the export has no OTT-distribution
    # tab -- a tile reading "0.0%" would be a confident wrong claim, so the
    # whole tile is removed and the remaining four reflow across the row,
    # exactly as the Flight tile does on the recap slide.
    if delivery.ctv_share is not None:
        _fill_tokens(slide, {"CTV_SHARE": _pct(delivery.ctv_share, 1)})
    else:
        _reflow_tile_row(slide, {"CtvShareTile"}, tile_names=_DELIVERY_TILE_ROW)

    publishers = sorted(delivery.top_publishers, key=lambda p: p[1],
                        reverse=True)[:TOP_PUBLISHERS_ROW_CAP]
    warnings = _fill_named_table(
        slide, "TopPublishersTable", "TOP_PUBLISHERS_ROWS",
        [{"name": name, "impressions": _int(count),
          "vcr": _pct(delivery.channel_vcr[name], 1) if name in delivery.channel_vcr else "--"}
         for name, count, _pct_share in publishers],
        ["name", "impressions", "vcr"])

    region = _shape(slide, "ChartRegion")
    label_shape = _shape(slide, "ChartRegionLabel")
    png = report_charts.render_bar_chart(
        [label for label, _count in delivery.by_daypart],
        [count for _label, count in delivery.by_daypart],
        region.width, region.height,
        value_labels=[_int(count) for _label, count in delivery.by_daypart])
    _place_image(slide, region, label_shape, png)
    return warnings


def _fill_delivery_breakdown(slide, delivery, narrative_override=None, plan_vs_actual=None):
    """report:delivery_breakdown (v0_3) -- two CONDITIONAL tables stacked in
    the left column, each with its own named header, plus a VCR-by-creative
    chart in its own column.

    A dimension with one value or none is not a breakdown, so its table and
    its header are deleted together. No reflow is needed or wanted: the
    template stacks them, so a surviving table simply has space beneath it
    (Matt's own layout note for v0_3). The caller decides whether this
    slide exists at all -- see `delivery_breakdown_applies`.

    `plan_vs_actual` (Highlights/Takeaways rework, `plan_vs_actual_facts()`'s
    own return or None) supplies DeliveryByGeoTable's `planned`/`pct_of_plan`
    values, matched per row by the SAME normalized-label match
    `plan_vs_actual_facts` already used. **Superseded 2026-09-13: Planned/
    % of plan are no longer left blank when the toggle is off -- those two
    columns are REMOVED from the table entirely** (`_fill_named_table`'s
    `full_fields` mechanism), because an empty column on a client slide
    reads as missing data, not as "nothing to show here." `geo_fields` is
    the 5-field list only when `plan_vs_actual` is truthy; otherwise the
    plain 3-field list, with `full_fields` always the 5-field list so
    `_fill_named_table` can tell which two columns to drop -- but only
    when the real template still HAS 5 columns, so a pre-v0_7 3-column
    template (which structurally can't show them regardless of the toggle)
    still falls through to the old graceful column_warning unchanged.
    """
    geo = list(delivery.by_geo or [])
    creatives = sorted(delivery.by_creative, key=lambda c: c[1],
                       reverse=True)[:TOP_CREATIVES_ROW_CAP]
    total = delivery.delivered_impressions or sum(c for _l, c in geo) or 0

    dimensions = []
    warnings = []
    if len(geo) > 1:
        dimensions.append("geography")
        # v0_7 widened DeliveryByGeoTable to 5 columns -- Geography |
        # Planned | Delivered | % of plan | VCR. Matt's 2026-09-13 ruling:
        # with no plan data (toggle off), Planned/% of plan are REMOVED,
        # not left blank -- `full_fields` below is what lets
        # `_fill_named_table` know which two columns to drop, and only
        # does so when the real template still has all 5 (a pre-v0_7
        # 3-column template can't show them regardless of the toggle, and
        # falls through to the old graceful column_warning unchanged).
        geo_fields = (["label", "planned", "impressions", "pct_of_plan", "vcr"] if plan_vs_actual
                     else ["label", "impressions", "vcr"])
        plan_by_label = {}
        if plan_vs_actual:
            for row in plan_vs_actual["rows"]:
                for label, _count in geo:
                    if _geo_labels_match(row["label"], label):
                        plan_by_label[label] = row
        warnings += _fill_named_table(
            slide, "DeliveryByGeoTable", "DELIVERY_BY_GEO_ROWS",
            [{"label": label,
              "impressions": _int(count),
              # Real, impression-weighted VCR from the flight-detail tab --
              # the geo tab itself carries impressions only. "--" only when
              # that tab genuinely can't supply one, never a borrowed number.
              "vcr": (_pct(delivery.geo_vcr[label], 1)
                      if label in delivery.geo_vcr else "--"),
              "planned": (_int(plan_by_label[label]["planned"]) if label in plan_by_label else ""),
              "pct_of_plan": (_pct(plan_by_label[label]["pct_of_plan"], 0)
                             if label in plan_by_label and plan_by_label[label]["pct_of_plan"] is not None
                             else "")}
             for label, count in sorted(geo, key=lambda g: -g[1])],
            geo_fields,
            full_fields=["label", "planned", "impressions", "pct_of_plan", "vcr"])
    else:
        _delete_named_shapes(slide, "DeliveryByGeoTable", "DeliveryByGeoHeader")

    if len(delivery.by_creative or []) > 1:
        dimensions.append("creative")
        warnings += _fill_named_table(
            slide, "DeliveryByCreativeTable", "DELIVERY_BY_CREATIVE_ROWS",
            [{"label": name, "impressions": _int(count), "vcr": _pct(vcr, 1)}
             for name, count, _length, _hours, vcr in creatives],
            ["label", "impressions", "vcr"])
    else:
        _delete_named_shapes(slide, "DeliveryByCreativeTable", "DeliveryByCreativeHeader")

    lead_geo = max(geo, key=lambda g: g[1]) if geo else None
    note = "How delivery split across " + " and ".join(dimensions) + "."
    narrative_bits = []
    if lead_geo and total:
        narrative_bits.append(
            f"{lead_geo[0]} took {lead_geo[1] / total * 100:.0f}% of delivery")
    if creatives:
        best = max(creatives, key=lambda c: c[4])
        narrative_bits.append(f"{best[0]} completed at {_pct(best[4], 1)}")
    _fill_tokens(slide, {
        "DELIVERY_BREAKDOWN_NOTE": note,
        "DELIVERY_BREAKDOWN_NARRATIVE": narrative_override or (
            "; ".join(narrative_bits) + "." if narrative_bits else ""),
    })

    region = _shape(slide, "ChartRegion")
    label_shape = _shape(slide, "ChartRegionLabel")
    png = report_charts.render_bar_chart(
        [name for name, _c, _l, _h, _v in creatives],
        [vcr for _n, _c, _l, _h, vcr in creatives],
        region.width, region.height,
        value_labels=[_pct(vcr, 1) for _n, _c, _l, _h, vcr in creatives])
    _place_image(slide, region, label_shape, png)
    return warnings


def delivery_breakdown_applies(delivery):
    """Whether report:delivery_breakdown has anything to say. One geo option
    and one creative is not a breakdown -- the slide is dropped entirely
    rather than shown with two one-row tables, the same
    show-what-matters rule the attribution slide's single-dimension
    choice already follows."""
    if delivery is None:
        return False
    return len(delivery.by_geo or []) > 1 or len(delivery.by_creative or []) > 1


LIVE_SPORTS_EVENT_ROW_CAP = 10  # 2026-09-08: a design rule (what a client reads), same
                                # reasoning as TOP_PUBLISHERS_ROW_CAP/TOP_CREATIVES_ROW_CAP
_LIVE_SPORTS_TILE_ROW = ("SportsImpressionsTile", "SportsVcrTile", "SportsPacingTile")


def live_sports_applies(delivery):
    """Whether report:live_sports has anything to show -- a live-sports
    package detected inside the delivery workbook (attribution_import.py's
    own _SPORTS_EVENT detection). Most delivery exports don't carry one, so
    this is False far more often than True."""
    return delivery is not None and delivery.live_sports is not None


def combined_headline_impressions(attribution, delivery):
    """OTT + live sports, combined -- a real WAEPA-adjacent finding
    (Prince George's Community College, 2026-09-08): the export's own totals
    NEVER combine the two blocks (confirmed against the real file's own
    DETAILS BY FLIGHT/KPI DELIVERY totals, both OTT-only), because Premion's
    dashboard treats them as two separate campaigns. But a client who bought
    both OTT and a live-sports package bought one combined flight, and the
    Highlights tile is the ONE number a client remembers -- reporting only
    the OTT half of what they paid for undercounts the campaign. Sports
    stays visibly BROKEN OUT on its own report:live_sports slide (its own
    tiles, its own event table) rather than folded silently into the
    Highlights number with nothing to show for it -- "combined, with sports
    broken out" is the settled shape, not a plain sum with no trace of where
    the sports share went. `DeliveryExport.delivered_impressions` itself is
    never mutated to include sports -- every other caller of that field
    (delivery recap, delivery breakdown, facts payload's own "delivery"
    section) keeps meaning "OTT only," unchanged."""
    base = delivery.delivered_impressions if delivery is not None else attribution.delivered_impressions
    sports = delivery.live_sports.delivered_impressions if live_sports_applies(delivery) else 0
    return base + sports


def _fill_live_sports(slide, live_sports, narrative_override=None):
    """report:live_sports -- a live-sports package riding inside the same
    delivery workbook as the OTT figures (attribution_import.py's own
    _SPORTS_EVENT detection note has the full story: a real Prince George's
    Community College export, RFPID-266713, a PREM TV package alongside the
    OTT campaign's RFPID-266710).

    Three tiles (delivered, VCR, pacing against the flight goal -- "55,674
    of 625,000", the exact shape Matt's own spec asked for) and ONE table:
    the game-level event breakdown, top N by impressions
    (LIVE_SPORTS_EVENT_ROW_CAP), with a final rolled-up row summing EVERY
    event (not just the shown top N) so the true total is always visible --
    "a network rollup line," read literally as one summary line rather than
    a whole second table, since the slide's own spec described one table,
    not two. An optional league breakdown (SportsByLeagueTable) is shown
    only when more than one league actually ran; almost every real package
    is single-league, in which case it would just repeat the pacing tile's
    own total, so it's deleted rather than shown as a one-row table -- same
    show-what-matters rule _fill_delivery_breakdown already follows for a
    single-value dimension.
    """
    _fill_tokens(slide, {
        "SPORTS_IMPRESSIONS": _int(live_sports.delivered_impressions),
        "SPORTS_VCR": _pct(live_sports.vcr, 1),
        "SPORTS_PACING": f"{_int(live_sports.delivered_impressions)} of {_int(live_sports.flight_goal)}",
        "SPORTS_PACKAGE_TYPE": live_sports.package_type,
        "SPORTS_RFPID": live_sports.rfpid,
        "SPORTS_GEO": live_sports.delivered_geo,
    })

    top_events = sorted(live_sports.events, key=lambda e: e.delivered_impressions,
                        reverse=True)[:LIVE_SPORTS_EVENT_ROW_CAP]
    event_rows = [{
        "date": f"{e.day.strftime('%b')} {e.day.day}", "event": e.event, "network": e.network,
        "impressions": _int(e.delivered_impressions), "vcr": _pct(e.vcr, 1),
    } for e in top_events]
    total_impressions = sum(e.delivered_impressions for e in live_sports.events)
    total_completed = sum(e.completed_impressions for e in live_sports.events)
    total_vcr = total_completed / total_impressions if total_impressions else 0.0
    event_rows.append({
        "date": "", "event": f"All {len(live_sports.events)} events", "network": "",
        "impressions": _int(total_impressions), "vcr": _pct(total_vcr, 1),
    })
    warnings = _fill_named_table(
        slide, "SportsEventTable", "SPORTS_EVENT_ROWS", event_rows,
        ["date", "event", "network", "impressions", "vcr"])

    if len(live_sports.by_league) > 1:
        warnings += _fill_named_table(
            slide, "SportsByLeagueTable", "SPORTS_BY_LEAGUE_ROWS",
            [{"label": name, "impressions": _int(count)}
             for name, count in sorted(live_sports.by_league, key=lambda l: -l[1])],
            ["label", "impressions"])
    else:
        _delete_named_shapes(slide, "SportsByLeagueTable", "SportsByLeagueHeader")

    top_event = (max(live_sports.events, key=lambda e: e.delivered_impressions)
                if live_sports.events else None)
    top_network = (max(live_sports.by_network, key=lambda n: n[1])
                  if live_sports.by_network else None)
    narrative_bits = []
    if top_event:
        narrative_bits.append(f"{top_event.event} led all events with "
                              f"{_int(top_event.delivered_impressions)} impressions")
    if top_network:
        narrative_bits.append(f"{top_network[0]} was the leading network")
    _fill_tokens(slide, {
        "LIVE_SPORTS_NARRATIVE": narrative_override or (
            "; ".join(narrative_bits) + "." if narrative_bits else ""),
    })
    return warnings


def _fill_attribution_breakdown(slide, attribution, headline_note, narrative_override=None,
                                dimension_override=None, include_conversions=False):
    dimension, rows = pick_breakdown_dimension(attribution, dimension_override)
    table_rows = breakdown_rows(rows)
    _fill_tokens(slide, {
        "ATTRIBUTION_HEADLINE_NOTE": headline_note or (
            f"Attribution broke out by {dimension.lower()} -- "
            f"{_pct(attribution.attributed_rate)} overall attributed rate."),
        "BREAKDOWN_DIMENSION_LABEL": dimension,
        "ATTRIBUTION_NARRATIVE": narrative_override or (
            f"{table_rows[0]['label']} led all {dimension.lower()}s at a "
            f"{table_rows[0]['rate']} attributed rate." if table_rows else ""),
    })
    # A fifth column, "conv_rate" -- present only when `include_conversions`
    # is True. When it's False, `full_fields` tells `_fill_named_table` to
    # REMOVE the column outright (2026-09-13 ruling) rather than leave it
    # blank, on a real v0_8 template; an older, narrower one falls through
    # to the graceful column_warning unchanged.
    breakdown_fields = ["label", "delivered", "attributed", "rate"]
    if include_conversions:
        breakdown_fields.append("conv_rate")
    warnings = _fill_named_table(slide, "BreakdownTable", "BREAKDOWN_ROWS", table_rows,
                                breakdown_fields,
                                full_fields=["label", "delivered", "attributed", "rate", "conv_rate"])
    region = _shape(slide, "ChartRegion")
    label_shape = _shape(slide, "ChartRegionLabel")
    png = report_charts.render_bar_chart(
        [r["label"] for r in table_rows], [r["_attributed_raw"] for r in table_rows],
        region.width, region.height, value_labels=[r["rate"] for r in table_rows])
    _place_image(slide, region, label_shape, png)
    return warnings


# v0_3's url_report is laid out for a fixed 4 intent rows and 7 URL rows,
# and does NOT auto-reflow (Matt's own note). Both real datasets produce 5-6
# intent rows after folding, so this fires on every real report.
_INTENT_ROWS_IN_TEMPLATE = 4
_URL_ROWS_IN_TEMPLATE = 7
_URL_ROWS_FLOOR = 4          # never trim the actual subject of the slide below this
_INTENT_ROWS_CAP = 6         # the template's own stated 1-6 expectation


def _header_above(slide, table_shape):
    """The section header sitting immediately above a table -- found by
    GEOMETRY, not by name or by matching its text. v0_3 leaves these headers
    generically named ("Text 4"), and matching on the literal
    "TOP PAGES BY ATTRIBUTED VISITS" would silently stop working the first
    time someone rewords it in PowerPoint. The lowest shape that starts
    above the table and overlaps it horizontally is unambiguous here.
    """
    best = None
    left, right = table_shape.left, table_shape.left + table_shape.width
    for shape in slide_map.iter_all_shapes(slide.shapes):
        if shape is table_shape or shape.top is None or shape.left is None:
            continue
        if shape.top >= table_shape.top:
            continue
        if shape.left + (shape.width or 0) <= left or shape.left >= right:
            continue
        if best is None or shape.top > best.top:
            best = shape
    return best


def _url_intent_narrative(intent_rows, url_rows):
    """Two facts about where visitors went, stated without inventing a link
    between them and without saying the same thing twice.

    Three real shapes this has to survive, all seen in the two real
    renders: the leading intent class is a category ("Product
    consideration") and the top page belongs to a different class
    (MW -- Homepage); the leading class IS the homepage (Cardinal), where
    "landed on homepage pages" is broken English and naming Homepage twice
    is redundant; and no intent data at all.
    """
    if not intent_rows:
        return ""
    lead = intent_rows[0]
    if lead["intent"] == "homepage":
        first = f"{lead['share']} of attributed visits landed on the homepage"
    else:
        first = (f"{lead['share']} of attributed visits landed on "
                f"{lead['label'].lower()} pages")
    # The second fact names the most-visited page the first sentence has not
    # already named -- otherwise Cardinal reads "landed on the homepage.
    # Homepage was the single most-visited destination", which is one fact
    # wearing two sentences.
    named = "homepage" if lead["intent"] == "homepage" else None
    nxt = next((r for r in url_rows if r["label"].strip().lower() != named), None)
    if nxt is None:
        return first + "."
    if lead["intent"] == "homepage":
        return (f"{first}, and {nxt['label']} was the most-visited page beyond it "
               f"({nxt['share']}).")
    return f"{first}. {nxt['label']} was the single most-visited destination ({nxt['share']})."


def _response_profile_narrative(recency, referral, day_of_week):
    """Deterministic fallback when no Claude narrative was supplied -- the
    response-profile equivalent of `_url_intent_narrative`. Names the
    strongest signal available: the direct-visit share (the highest-
    confidence read on its own), the recency split, and the day-of-week
    pattern only when it actually cleared the uneven threshold -- never
    manufactured from a flat week.
    """
    parts = []
    if referral and referral.get("direct_share"):
        parts.append(f"{_pct(referral['direct_share'], 1)} of attributed visits arrived direct.")
    if recency:
        parts.append(f"{_pct(recency['share_within_0_3_days'], 1)} of attributed visitors "
                     f"responded within 3 days of exposure.")
    if day_of_week and day_of_week.get("uneven"):
        parts.append(f"{day_of_week['best_day']['day']} had the strongest attributed rate, "
                     f"{day_of_week['worst_day']['day']} the weakest.")
    return " ".join(parts)


def _fill_response_profile(slide, attribution, narrative_override=None):
    """report:response_profile (v0_6) -- always present, between
    attribution_breakdown and url_report. Left column is a recency bar
    chart (ChartRegion) plus a one-line stat; right column is the
    referral-source table; a conditional day-of-week table sits under the
    chart, shown only when the week is genuinely uneven (`day_of_week_
    facts`'s own named-threshold gate) -- deleted along with its header
    otherwise, the same "delete both, nothing reflows into the gap" rule
    ChartRegion/MapRegion already use elsewhere. The chart region could
    grow into the freed space when the table's gone; it doesn't yet -- a
    v2 polish per Matt, not required for launch.
    """
    recency = recency_facts(attribution)
    referral = referral_facts(attribution)
    day_of_week = day_of_week_facts(attribution)
    if recency is None:
        raise MissingTokenError("report:response_profile: the export has no recency breakdown")
    if referral is None:
        raise MissingTokenError("report:response_profile: the export has no referral breakdown")

    _fill_tokens(slide, {
        "RECENCY_STAT": (f"{_pct(recency['share_within_0_3_days'], 1)} of attributed visitors "
                        f"responded within 3 days of exposure."),
        "RESPONSE_PROFILE_NARRATIVE": narrative_override or _response_profile_narrative(
            recency, referral, day_of_week),
    })

    region = _shape(slide, "ChartRegion")
    label_shape = _shape(slide, "ChartRegionLabel")
    buckets = recency["buckets"]
    png = report_charts.render_bar_chart(
        [b["bucket"] for b in buckets], [b["visitors"] for b in buckets],
        region.width, region.height,
        value_labels=[_pct(b["share"], 1) for b in buckets])
    _place_image(slide, region, label_shape, png)

    referral_rows = sorted(referral["sources"], key=lambda r: r["visitors"], reverse=True)
    warnings = _fill_named_table(slide, "ReferralTable", "REFERRAL_ROWS", [
        {"source": r["source"], "visitors": _int(r["visitors"]), "share": _pct(r["share"], 1)}
        for r in referral_rows
    ], ["source", "visitors", "share"])

    if day_of_week is not None and day_of_week["uneven"]:
        warnings += _fill_named_table(slide, "DayOfWeekTable", "DAY_OF_WEEK_ROWS", [
            {"day": d["day"], "rate": _pct(d["attributed_rate"], 2)}
            for d in day_of_week["days"]
        ], ["day", "rate"])
    else:
        _delete_named_shapes(slide, "DayOfWeekHeader", "DayOfWeekTable")
    return warnings


def _fill_url_report(slide, attribution, headline_note, narrative_override=None,
                     include_conversions=False):
    intent_rows = intent_summary_rows(
        attribution, include_conversions=include_conversions)[:_INTENT_ROWS_CAP]
    url_rows = top_url_rows(attribution, limit=_URL_ROWS_IN_TEMPLATE,
                            include_conversions=include_conversions)
    if not url_rows:
        raise MissingTokenError("report:url_report: the export has no URL breakdown")

    # Row height is READ from the template rather than assumed -- the two
    # tables happen to share 0.300in today, and a future retheme that changes
    # it must move this shift with it automatically.
    intent_table = _shape(slide, "IntentSummaryTable")
    row_height = intent_table.table.rows[1].height
    overflow = max(0, len(intent_rows) - _INTENT_ROWS_IN_TEMPLATE)
    if overflow:
        url_table = _shape(slide, "TopUrlTable")
        header = _header_above(slide, url_table)
        shift = int(row_height) * overflow
        url_table.top = Emu(int(url_table.top) + shift)
        if header is not None:
            header.top = Emu(int(header.top) + shift)
        # Trim the URL table by exactly what the intent table borrowed, so
        # the pair still ends where the template's own block ended -- but
        # never below the floor, since the URL list is what this slide is
        # actually about.
        url_rows = url_rows[:max(_URL_ROWS_FLOOR, _URL_ROWS_IN_TEMPLATE - overflow)]

    lead = intent_rows[0] if intent_rows else None
    _fill_tokens(slide, {
        "URL_HEADLINE_NOTE": headline_note or "Top pages by attributed unique visitors.",
        # Falls back to the deterministic mix-only sentence (never claiming
        # goal alignment) when no Claude draft supplied its own narrative --
        # see app.apply_attribution_draft.
        "URL_INTENT_NARRATIVE": narrative_override or _url_intent_narrative(intent_rows, url_rows),
    })
    # "converted" is a fourth column on each table, present only when
    # `include_conversions` is True -- REMOVED entirely otherwise
    # (2026-09-13 ruling), same `full_fields` pattern as BreakdownTable's
    # conv_rate column above.
    intent_fields = ["label", "visits", "share"]
    url_fields = ["label", "visitors", "share"]
    if include_conversions:
        intent_fields.append("converted")
        url_fields.append("converted")
    warnings = _fill_named_table(slide, "IntentSummaryTable", "INTENT_SUMMARY_ROWS",
                                 intent_rows, intent_fields,
                                 full_fields=["label", "visits", "share", "converted"])
    warnings += _fill_named_table(slide, "TopUrlTable", "TOP_URL_ROWS", url_rows,
                                  url_fields,
                                  full_fields=["label", "visitors", "share", "converted"])
    return warnings


def dma_zcta_coverage_warnings(zips):
    """Dev-facing, run at UPLOAD time (2026-09-12 follow-up to the West
    Virginia find) -- the same DMA -> states technique that found WV
    missing, moved earlier so a coverage gap is a warning at upload, not
    26 centroid dots discovered on a client's map at Generate.

    `zips` (an export's own real zip codes, e.g. `attribution.by_zip`'s
    labels) resolve to their DMA(s) via `geo_resolver.zips_to_markets`.
    For each DMA actually present, EVERY state its own counties span
    (`market_lookup.states_for_market` -- not just the states THESE
    particular zips happen to touch) must already have a built ZCTA file,
    or a later zip from the same DMA repeats the identical silent gap.
    Degrades to an empty list, never raises, when the market lookup isn't
    installed -- this is a diagnostic, not a gate on anything.
    """
    if not geo_resolver.market_lookup_available():
        return []
    resolution = geo_resolver.zips_to_markets([str(z) for z in (zips or []) if z])
    warnings = []
    for market_key in sorted(resolution.resolved):
        states = market_lookup.states_for_market(market_key)
        missing = sorted(s for s in states
                         if not (targeting_map.ZCTA_DIR / f"zcta_{s}.json.gz").exists())
        if missing:
            name = market_lookup.market_name(market_key) or market_key
            warnings.append(
                f"{name} reaches {', '.join(missing)}, which "
                f"{'has' if len(missing) == 1 else 'have'} no ZCTA boundary file built -- "
                f"run build_zcta_boundaries.py --add {' '.join(missing)}.")
    return warnings


def _fill_zip_analysis(slide, attribution, headline_note, narrative_override=None, targeted_zips=None):
    rows, dropped_zips = top_zip_rows(attribution)
    if not rows:
        raise MissingTokenError("report:zip_analysis: the export has no zip-code breakdown")
    outperformers = [r for r in rows if r["outperformer"]]
    # The narrative describes ONLY rows that actually beat the campaign
    # average -- a backfilled row is real volume but did not outperform,
    # and calling it a leader is exactly what `outperformer` exists to
    # prevent. With no outperformer at all, say so plainly rather than
    # promoting the biggest row into one.
    if outperformers:
        lead = outperformers[0]
        narrative = (f"{lead['zip']} delivered the strongest response at "
                     f"{lead['rate']} ({lead['multiple']} the campaign average), "
                     f"with {len(outperformers)} zip "
                     f"{'code' if len(outperformers) == 1 else 'codes'} beating the "
                     f"average on meaningful volume.")
    else:
        narrative = ("No zip code combined an above-average attributed rate with "
                     "meaningful delivery volume; the table shows the highest-volume "
                     "zip codes instead.")
    _fill_tokens(slide, {
        "ZIP_HEADLINE_NOTE": headline_note or (
            "Zip codes beating the campaign's average attributed rate on meaningful volume."),
        "ZIP_NARRATIVE": narrative_override or narrative,
    })
    warnings = _fill_named_table(slide, "TopZipTable", "TOP_ZIP_ROWS", rows,
                                ["zip", "area", "share", "rate", "multiple"])
    if dropped_zips:
        # Dev-facing (2026-09-13 Ashburn/20149 ruling): a client never sees
        # "20149 | | 0.4%" -- the row is dropped from the table entirely,
        # but a rep/dev still needs to know it happened and which zip(s),
        # same "never silent" discipline as the map's own missing-polygon
        # warning just below.
        warnings.append(
            f"{len(dropped_zips)} zip code(s) qualified for the zip table but had no "
            f"resolvable area (no county on file, and no market/state agreement among "
            f"other zips sharing their prefix either) and were dropped rather than "
            f"shown with a blank Area cell: {', '.join(sorted(dropped_zips))}.")
    region = _shape(slide, "MapRegion")
    label_shape = _shape(slide, "MapRegionLabel")
    # Attributed RATE is the weight, not impressions: the slide's question
    # is where the campaign worked hardest, not where it delivered most --
    # the same distinction the table's own ranking rule draws.
    png, missing_polygons = report_charts.render_zip_map(
        {r.label: r.attributed_rate for r in attribution.by_zip
         if r.delivered_impressions > 0},
        region.width, region.height, targeted_zips=targeted_zips)
    _place_image(slide, region, label_shape, png)
    if missing_polygons:
        # Reported, never silent -- same discipline as the Shopify pixel
        # drop on the URL slide. Most of these draw as centroid dots
        # rather than shaded areas; a zip with no point at all (2026-09-13
        # Ashburn/20149 find -- a real, deliverable PO-box-only zip has no
        # ZCTA and so no coordinate either) can't be drawn at all and is
        # omitted from the map entirely -- named here either way, so a
        # WHOLE STATE missing its ZCTA file (`build_zcta_boundaries.py
        # --add <ST>`) is visible as a big number rather than looking like
        # a slightly sparse map, and a genuinely unplaceable zip doesn't
        # just vanish with nothing to show for it.
        warnings.append(
            f"{len(missing_polygons)} zip code(s) have no ZCTA boundary -- most render "
            f"as points rather than shaded areas, and any with no resolvable location "
            f"at all are omitted from the map entirely "
            f"(e.g. {', '.join(sorted(missing_polygons)[:4])}). PO-box-only zips and "
            f"retired ZCTAs are expected; a large count usually means a state hasn't "
            f"been built -- see build_zcta_boundaries.py --add.")
    return warnings


def _ott_screen_stat(by_screen):
    if not by_screen:
        return "--"
    total = sum(clicks for _screen, clicks in by_screen)
    if not total:
        return "--"
    top_screen, top_clicks = max(by_screen, key=lambda pair: pair[1])
    return (f"{top_screen} led with {_int(top_clicks)} of {_int(total)} total clicks "
           f"({_pct(top_clicks / total, 0)}).")


def _ott_blended_stat(ott):
    """None when the export has no blended tab at all -- the caller deletes
    BlendedHeader/BlendedStat outright in that case, same as every other
    "nothing to show" shape pair in this module. Never states a frequency
    -- see `ott_retargeting_facts`'s own docstring: the real Cardinal
    export's blended tab is campaign-to-date, not scoped to this export's
    reporting period, and a bare "65" would misstate a months-long
    accumulation as a one-period finding. States impressions/uniques and
    says "cumulative" in the same sentence instead of a caveat the reader
    has to infer.
    """
    if not ott.blended_impressions:
        return None
    return (f"{_int(ott.blended_impressions)} blended CTV + display impressions reached "
           f"{_int(ott.blended_uniques)} unique visitors (cumulative since campaign start).")


def _fill_ott_retargeting(slide, ott, narrative_override=None):
    """report:ott_retargeting (v0_6) -- present only when an OTT
    retargeting export was uploaded; the caller decides that via
    `build_report_deck`'s `ott` parameter, never the delivery-file
    presence check every other conditional slide in this module uses --
    per Matt, this slide is explicitly NOT part of the delivery set.

    The creative table and its header are deleted together, with
    AdSizeHeader/AdSizeTable shifted up by the exact freed height (read
    from the template, not assumed), whenever there's only one creative
    CONCEPT to show -- the same "shift the block below up to close the
    gap" idea `_fill_url_report`'s intent-table overflow uses, run in the
    opposite direction (closing a gap instead of opening one).
    """
    _fill_tokens(slide, {
        "OTT_IMPRESSIONS": _int(ott.impressions),
        "OTT_CLICKS": _int(ott.clicks),
        "OTT_CTR": _pct(ott.ctr, 2),
        "OTT_RETARGETING_NARRATIVE": narrative_override or (
            f"The retargeting campaign delivered {_int(ott.impressions)} display "
            f"impressions at a {_pct(ott.ctr, 2)} click-through rate."),
    })

    warnings = []
    if len(ott.creative_groups) > 1:
        warnings += _fill_named_table(slide, "CreativeTable", "OTT_CREATIVE_ROWS", [
            {"creative": g.base_name, "impressions": _int(g.impressions),
             "clicks": _int(g.clicks), "ctr": _pct(g.ctr, 2)}
            for g in ott.creative_groups
        ], ["creative", "impressions", "clicks", "ctr"])
    else:
        creative_header = _shape(slide, "CreativeHeader")
        ad_size_header = _shape(slide, "AdSizeHeader")
        ad_size_table = _shape(slide, "AdSizeTable")
        shift = int(ad_size_header.top) - int(creative_header.top)
        _delete_named_shapes(slide, "CreativeHeader", "CreativeTable")
        ad_size_header.top = Emu(int(ad_size_header.top) - shift)
        ad_size_table.top = Emu(int(ad_size_table.top) - shift)

    warnings += _fill_named_table(slide, "AdSizeTable", "OTT_AD_SIZE_ROWS", [
        {"unit": f"{r.ad_size} {r.label}", "impressions": _int(r.impressions),
         "clicks": _int(r.clicks), "ctr": _pct(r.ctr, 2)}
        for r in ott.by_ad_size
    ], ["unit", "impressions", "clicks", "ctr"])

    _fill_tokens(slide, {"OTT_SCREEN_STAT": _ott_screen_stat(ott.by_screen)})

    blended = _ott_blended_stat(ott)
    if blended:
        _fill_tokens(slide, {"OTT_BLENDED_STAT": blended})
    else:
        _delete_named_shapes(slide, "BlendedHeader", "BlendedStat")

    return warnings


def _fill_takeaways(slide, attribution, delivery, takeaway_bullets, whats_next_bullets):
    items = takeaway_bullets or default_takeaway_bullets(attribution, delivery)
    _fill_head_detail_bullets(slide, "TAKEAWAYBullets", items, max_items=4)
    _fill_bullets(slide, "WHATS_NEXT_BULLETS", whats_next_bullets)
    _shrink_bullet_box_to_fit(slide, "TAKEAWAYBullets")
    _shrink_bullet_box_to_fit(slide, "WhatsNextBullets")


# ---------------------------------------------------------------------------
# report:summary / report:case_study (v0_9, ATTRIBUTION_REPORT_PLAN.md Phase
# 6's "one-slide summary and case study") -- both STANDALONE (see
# `_is_standalone` above), built one at a time via `build_single_slide`,
# never part of the normal multi-slide report. Both render from EXACTLY the
# same stored `threads`/`attribution`/`delivery` a full report build already
# used -- `rehydrate_attribution`/`rehydrate_delivery` above are what make
# "regenerate later from a logged report, no second draft call" true rather
# than aspirational: every fact function below is called identically whether
# it's handed a freshly-parsed export or a rehydrated stored one.
# ---------------------------------------------------------------------------

def _lead_thread_head(threads):
    """The framing line for a panel that leads with the model's own
    priority -- the first GOAL thread's `head` when one exists (goal
    threads always sort first, `distribute_threads`'s own ordering rule),
    falling back to the first thread of any kind. None when `threads` has
    nothing usable at all."""
    threads = [t for t in (threads or []) if isinstance(t, dict) and str(t.get("head") or "").strip()]
    if not threads:
        return None
    goal_threads = [t for t in threads if t.get("anchor") == "goal"]
    return (goal_threads[0] if goal_threads else threads[0])["head"].strip()


def _lead_thread_meaning(threads):
    """The synthesis line for BOTTOM_LINE/CS_TAKEAWAY -- the first GOAL
    thread's own `meaning` (what a goal thread's finding actually means for
    the campaign, `distribute_threads`'s own field), falling back to the
    first thread of any kind that has one. None when no surviving thread
    carries a meaning at all."""
    threads = [t for t in (threads or []) if isinstance(t, dict) and str(t.get("meaning") or "").strip()]
    if not threads:
        return None
    goal_threads = [t for t in threads if t.get("anchor") == "goal"]
    return (goal_threads[0] if goal_threads else threads[0])["meaning"].strip()


def _accepted_optimization_texts(accepted_optimizations):
    """The rep-approved wording of every accepted-or-edited optimization
    this period -- `entry["final_text"]` from `accepted_optimizations_
    from_report_json`'s own return (declined/watch-list candidates are
    already excluded there, never re-filtered here). Both artifacts read
    from this ONE list so a declined candidate structurally cannot surface
    in either -- there is no separate path that could re-derive one."""
    return [e["final_text"] for e in (accepted_optimizations or []) if e.get("final_text")]


def accepted_optimizations_clause(accepted_optimizations):
    """A trailing clause for BOTTOM_LINE naming this period's accepted
    optimization(s), or "" when none were accepted -- appended to, never
    replacing, the thread-derived synthesis sentence."""
    texts = _accepted_optimization_texts(accepted_optimizations)
    if not texts:
        return ""
    if len(texts) == 1:
        return f" This period also applied one optimization: {texts[0]}"
    return f" This period also applied {len(texts)} optimizations, including: {texts[0]}"


def _optimization_col_note(accepted_optimizations):
    """CS_COL_3_NOTE's own compact form of the same accepted-only list --
    None (shape deleted) when nothing was accepted this period."""
    texts = _accepted_optimization_texts(accepted_optimizations)
    return f"Optimization applied: {texts[0]}" if texts else None


# A case study is white-labeled by DEFAULT (opt-in "Name the client," never
# assumed) -- "a case study is a claim to a prospect," and a prospect's own
# claim can't carry a competitor's name. The noun here is deliberately BARE
# (no article) -- used once, in CS_EYEBROW's own "{NOUN} CASE STUDY" phrase,
# which reads correctly without one ("REGIONAL BANK CASE STUDY").  Verticals
# with no entry fall back to the generic "Premion client" rather than
# guessing.
_WHITE_LABEL_VERTICAL_NOUN = {
    "auto": "auto dealer", "banking": "regional bank", "healthcare": "healthcare provider",
    "retail": "retailer", "home_improvement": "home improvement retailer",
    "education": "education client", "travel": "travel brand",
    "entertainment": "entertainment venue", "dining_qsr": "QSR brand",
    "legal": "law firm", "real_estate": "real estate company",
}


def white_label_vertical_noun(vertical):
    return _WHITE_LABEL_VERTICAL_NOUN.get(vertical, "Premion client")


_WHITE_LABEL_STOPWORDS = {
    "the", "and", "of", "for", "inc", "llc", "co", "corp", "corporation",
    "company", "group", "holdings", "partners", "national", "regional",
}


def white_label_text(text, client_name):
    """Strip the client's identity from `text`, case-insensitively,
    replacing each match with "the campaign" -- a literal-string redaction
    pass, deliberately not NLP.

    Two passes, not one: the FULL name phrase first (the safe, precise
    match this always did), then each individual word of it that is 4+
    characters and not a generic connector (`_WHITE_LABEL_STOPWORDS`) --
    added 2026-09-15 after a real leak: a drafted thread's own `meaning`
    referred to "Cardinal Plumbing" as just "Cardinal," which the full-
    phrase-only pass let straight through onto a client-facing slide (a
    real vault-insertion verification caught it, not a synthetic test).
    Natural writing routinely shortens a multi-word client name to its
    distinctive word ("Cardinal", "Mattress Warehouse" -> "MW" in this
    very codebase's own shorthand) -- a full-phrase match alone can't
    catch that.

    This trades a small amount of over-redaction for the thing that
    actually matters here: a generic industry word inside the client's own
    name ("Plumbing", "Warehouse", "Bank") can occasionally also get
    swapped out of an unrelated, legitimate sentence -- accepted, since
    the alternative is a real identity leak in front of a prospect, and
    the word list already excludes the most common short connectors. A
    single-word name (e.g. "WAEPA") is unaffected -- the second pass has
    nothing left to add once the first pass already covers it.

    A blank `text` or `client_name` passes through unchanged.

    A trailing mechanical cleanup collapses a doubled "the the" -- a real,
    visible artifact caught live (2026-09-15, walking the Cardinal Plumbing
    case study through the app): "...moving qualified prospects into the
    Cardinal services funnel" became "...into the the campaign services
    funnel" once "Cardinal" was swapped for "the campaign" right after an
    article the original sentence already had. This is a plain string fix
    (never grammar-aware beyond this one specific, mechanical collision),
    same "literal, not NLP" discipline as the rest of this function.
    """
    name = str(client_name or "").strip()
    if not text or not name:
        return text
    out = re.sub(re.escape(name), "the campaign", text, flags=re.IGNORECASE)
    for word in name.split():
        stripped = word.strip(".,")
        if len(stripped) < 4 or stripped.lower() in _WHITE_LABEL_STOPWORDS:
            continue
        out = re.sub(r"\b" + re.escape(stripped) + r"\b", "the campaign", out, flags=re.IGNORECASE)
    return re.sub(r"\bthe\s+the\s+campaign\b", "the campaign", out, flags=re.IGNORECASE)


def _fill_summary_takeaways(slide, highlight_bullets):
    """SummaryTakeaway{1,2,3}Head/Detail -- the first three surviving
    threads' own (head, finding) pairs, in the SAME rank `distribute_
    threads` already produced for the full deck's own Highlights slide
    (never re-ranked here). Fewer than 3 deletes the trailing row(s) and
    the divider(s) BETWEEN surviving rows outright -- a vertical list, not
    a tile row, so nothing reflows to fill the gap; the remaining rows
    already sit in their own fixed template slots."""
    pairs = list(highlight_bullets or [])[:3]
    values = {}
    for n in range(1, 4):
        if n <= len(pairs):
            head, detail = pairs[n - 1]
            values[f"SUMMARY_TAKEAWAY_{n}_HEAD"] = head
            values[f"SUMMARY_TAKEAWAY_{n}_DETAIL"] = detail
        else:
            _delete_named_shapes(slide, f"SummaryTakeaway{n}Num",
                                 f"SummaryTakeaway{n}Head", f"SummaryTakeaway{n}Detail")
    if len(pairs) < 3:
        _delete_named_shapes(slide, "SummaryTakeawayDivider2")
    if len(pairs) < 2:
        _delete_named_shapes(slide, "SummaryTakeawayDivider1")
    _fill_tokens(slide, values)


def fill_summary_slide(slide, attribution, delivery, client_name, threads,
                       accepted_optimizations, include_conversions=False, footnote=None):
    """report:summary (v0_9) -- the whole one-slide summary. `threads` is
    the model's own stored `draft["threads"]` array (the SAME shape
    `distribute_threads` already fans out for the full deck) -- no new
    drafting happens here, ever; that's what makes regenerating this later
    from a logged report a pure Python re-render. `accepted_optimizations`
    is `accepted_optimizations_from_report_json(report_json)`'s own
    return -- declined/watch-list candidates are excluded upstream, by
    construction, so they cannot reach BOTTOM_LINE. `footnote` is a plain
    string the CALLER supplies (e.g. app.py's own pixel-issue-window
    check) -- this module has no Streamlit/app-layer knowledge of that
    rule, so it never re-derives it; None (default) deletes the shape.
    """
    highlight_bullets, _takeaways, _whats_next = distribute_threads(threads)

    period = _date_range_label(attribution.flight_start, attribution.flight_end)
    subtitle = f"Website Attribution · {period}" if period else "Website Attribution"
    name = client_name or attribution.client_name
    if not name:
        raise MissingTokenError("report:summary/CLIENT_NAME: not in the export and none supplied")

    values = {"CLIENT_NAME": name, "SUMMARY_SUBTITLE": subtitle}

    # Tiles 1-3 are fixed; tile 4 is conversions when the rep's own toggle
    # and the export both support it, else the unique-visitor rate -- the
    # same "no half-states" discipline `_fill_highlights` already applies
    # to its own fourth tile.
    values["SUMMARY_TILE_1_VALUE"] = _int(combined_headline_impressions(attribution, delivery))
    values["SUMMARY_TILE_1_LABEL"] = "Impressions Delivered"
    values["SUMMARY_TILE_2_VALUE"] = _pct(attribution.attributed_rate)
    values["SUMMARY_TILE_2_LABEL"] = "Attributed Rate"
    values["SUMMARY_TILE_3_VALUE"] = _int(attribution.attributed_unique_visitors)
    values["SUMMARY_TILE_3_LABEL"] = "Unique Visitors"
    if include_conversions and getattr(attribution, "has_conversions", False):
        values["SUMMARY_TILE_4_VALUE"] = _int(attribution.attributed_conversions)
        values["SUMMARY_TILE_4_LABEL"] = "Conversions"
    else:
        values["SUMMARY_TILE_4_VALUE"] = _pct(attribution.attributed_unique_visitor_rate)
        values["SUMMARY_TILE_4_LABEL"] = "Unique Visitor Rate"

    # Sidebar -- framed by the model's own lead thread; the numbers
    # underneath are the export's own top intent classes, ranked by share
    # (a deterministic stand-in for true goal-matched selection -- see
    # ATTRIBUTION_REPORT_PLAN.md's own note on this simplification).
    values["SIDEBAR_HEADLINE"] = _lead_thread_head(threads) or "What Stood Out"
    intent_rows = intent_summary_rows(attribution)
    if intent_rows:
        top = intent_rows[0]
        values["SIDEBAR_STAT_VALUE"] = top["share"]
        values["SIDEBAR_STAT_LABEL"] = top["label"]
        values["SIDEBAR_STAT_DETAIL"] = f"{top['visits']} attributed visits"
    else:
        _delete_named_shapes(slide, "SidebarStatValue", "SidebarStatLabel", "SidebarStatDetail")
    subs = intent_rows[1:4]
    for n in range(1, 4):
        if n <= len(subs):
            row = subs[n - 1]
            values[f"SIDEBAR_SUB_{n}_VALUE"] = row["share"]
            values[f"SIDEBAR_SUB_{n}_LABEL"] = row["label"]
        else:
            _delete_named_shapes(slide, f"SidebarSub{n}Value", f"SidebarSub{n}Label")

    device = device_split_facts(attribution)
    if device:
        values["SIDEBAR_SECOND_HEADER"] = "Device Split"
        values["SIDEBAR_SECOND_VALUE"] = f"{device['share'] * 100:.0f}"
        values["SIDEBAR_SECOND_UNIT"] = "%"
        values["SIDEBAR_SECOND_DETAIL"] = f"of attributed impressions came from {device['device']}"
    else:
        _delete_named_shapes(slide, "SidebarSecondHeader", "SidebarSecondValue",
                             "SidebarSecondUnit", "SidebarSecondDetail", "SidebarDivider")

    bottom_line = (_lead_thread_meaning(threads)
                  or "This period's campaign performance is summarized above.")
    bottom_line += accepted_optimizations_clause(accepted_optimizations)
    values["BOTTOM_LINE"] = bottom_line

    if footnote:
        values["SUMMARY_FOOTNOTE"] = footnote
    else:
        _delete_named_shapes(slide, "SummaryFootnote")

    _fill_tokens(slide, values)
    _fill_summary_takeaways(slide, highlight_bullets)


_CS_TILE_ROW = ("CsTile1", "CsTile2", "CsTile3", "CsTile4", "CsTile5")


def _case_study_zip_narrative(rows):
    """(headline, body) for CS_COL_3, from `top_zip_rows`'s own rows --
    the SAME "outperformer" rule the full zip_analysis slide uses (rate
    strictly above campaign average AND real volume), condensed to the
    single strongest one. (None, None) when nothing in the table actually
    beat the average -- the case study leads with a real outperformer or
    it doesn't run this column's own headline claim at all."""
    outperformers = [r for r in (rows or []) if r["outperformer"]]
    if not outperformers:
        return None, None
    lead = outperformers[0]
    headline = f"{lead['area']} led at {lead['rate']}"
    body = (f"ZIP {lead['zip']} ({lead['area']}) delivered the strongest response at "
           f"{lead['rate']} attributed -- {lead['multiple']} the campaign average.")
    return headline, body


def fill_case_study_slide(slide, attribution, delivery, threads, accepted_optimizations,
                          client_name, vertical=None, white_label=True):
    """report:case_study (v0_9) -- the vault-bound case study. Facts-only
    contract applies with full force here: every number is Python-computed
    from the same export a full report already used, never re-derived by a
    model. White-labeled by DEFAULT (`white_label=True`) -- "Name the
    client" is an explicit, opt-in caller choice (app.py), never assumed.

    White-label mode changes CS_EYEBROW/CS_HEADLINE structurally (not a
    redaction of the named-mode text -- genuinely different content), then
    sweeps every OTHER composed string through `white_label_text` as a
    safety net, since a thread's own `finding`/`meaning` (model output,
    drawn from real notes) can mention the client by name even when this
    function's own Python composition never does.
    """
    name = client_name or attribution.client_name
    if not name:
        raise MissingTokenError("report:case_study/CS_HEADLINE: not in the export and none supplied")
    period = _date_range_label(attribution.flight_start, attribution.flight_end)

    def short_date(d):
        return f"{d.month}/{d.day}/{d.strftime('%y')}" if d else None
    date_range = (f"{short_date(attribution.flight_start)} - {short_date(attribution.flight_end)}"
                 if attribution.flight_start and attribution.flight_end else None)

    values = {}
    if white_label:
        values["CS_EYEBROW"] = (f"{white_label_vertical_noun(vertical).upper()} CASE STUDY"
                                + (f" | {date_range}" if date_range else ""))
        # Falls back to a fact-only claim rather than raising when there's
        # no drafted thread at all (a report generated before a narrative
        # existed, or one whose draft call failed) -- same "degrade, don't
        # block" rule fill_summary_slide's own BOTTOM_LINE/SIDEBAR_HEADLINE
        # already follow. Found via a test scenario using an empty stub
        # draft (2026-09-15): raising here meant "Create case study" had no
        # working path at all for such a report.
        values["CS_HEADLINE"] = (_lead_thread_head(threads)
                                 or f"{_pct(attribution.attributed_rate)} Attributed Response Rate")
    else:
        values["CS_EYEBROW"] = "STREAMING OTT" + (f" | {date_range}" if date_range else "")
        values["CS_HEADLINE"] = name

    subhead_parts = [
        f"{_int(combined_headline_impressions(attribution, delivery))} impressions delivered",
        f"{_int(attribution.attributed_unique_visitors)} attributed visitors",
    ]
    if period:
        subhead_parts.append(period)
    values["CS_SUBHEAD"] = " · ".join(subhead_parts)

    # Tiles 1/2 are fixed; 3 (cost per visitor) is conditional on a linked
    # proposal's own budget; 4/5 are the export's top two intent classes --
    # the same deterministic ranking the summary slide's sidebar uses.
    values["CS_TILE_1_VALUE"] = _pct(attribution.attributed_rate)
    values["CS_TILE_1_LABEL"] = "Attributed Rate"
    values["CS_TILE_2_VALUE"] = _int(attribution.attributed_unique_visitors)
    values["CS_TILE_2_LABEL"] = "Unique Visitors"

    blank_tiles = set()
    budget = getattr(attribution, "_case_study_budget", None)  # see build_case_study_slide
    if budget and attribution.attributed_unique_visitors:
        values["CS_TILE_3_VALUE"] = _money(budget / attribution.attributed_unique_visitors)
        values["CS_TILE_3_LABEL"] = "Cost Per Visitor"
    else:
        blank_tiles.add("CsTile3")

    intent_rows = intent_summary_rows(attribution)
    for offset, tile_n in ((0, 4), (1, 5)):
        if offset < len(intent_rows):
            row = intent_rows[offset]
            values[f"CS_TILE_{tile_n}_VALUE"] = row["share"]
            values[f"CS_TILE_{tile_n}_LABEL"] = row["label"]
        else:
            blank_tiles.add(f"CsTile{tile_n}")

    # Column 1 -- where visitors went. Reuses the SAME deterministic
    # narrative the full url_report slide falls back to when no Claude
    # narrative override was supplied, so this column never invents a
    # second voice for the same fact.
    url_rows = top_url_rows(attribution)
    if not intent_rows or not url_rows:
        raise MissingTokenError("report:case_study/CS_COL_1: the export has no URL/intent "
                                "breakdown to build this column from")
    top_intent = intent_rows[0]
    values["CS_COL_1_LABEL"] = "WHERE THEY WENT"
    values["CS_COL_1_HEADLINE"] = f"{top_intent['share']} {top_intent['label']}"
    values["CS_COL_1_BODY"] = _url_intent_narrative(intent_rows, url_rows)

    # Column 2 -- how visitors responded. Same reuse of the response-
    # profile slide's own deterministic fallback narrative.
    recency = recency_facts(attribution)
    referral = referral_facts(attribution)
    day_of_week = day_of_week_facts(attribution)
    if recency is None or referral is None:
        raise MissingTokenError("report:case_study/CS_COL_2: the export has no recency/referral "
                                "breakdown to build this column from")
    values["CS_COL_2_LABEL"] = "HOW THEY RESPONDED"
    if referral.get("direct_share"):
        values["CS_COL_2_HEADLINE"] = f"{_pct(referral['direct_share'], 1)} Direct"
    else:
        values["CS_COL_2_HEADLINE"] = f"{_pct(recency['share_within_0_3_days'], 1)} Responded Fast"
    values["CS_COL_2_BODY"] = _response_profile_narrative(recency, referral, day_of_week)

    # Column 3 -- where it worked best, plus an accepted-optimization note
    # (accepted-only, by construction -- see _accepted_optimization_texts).
    zip_rows, _dropped = top_zip_rows(attribution)
    col3_headline, col3_body = _case_study_zip_narrative(zip_rows)
    if not col3_headline:
        raise MissingTokenError("report:case_study/CS_COL_3: no zip code beat the campaign "
                                "average on meaningful volume to lead this column with")
    values["CS_COL_3_LABEL"] = "WHERE IT WORKED BEST"
    values["CS_COL_3_HEADLINE"] = col3_headline
    values["CS_COL_3_BODY"] = col3_body
    note = _optimization_col_note(accepted_optimizations)
    if note:
        values["CS_COL_3_NOTE"] = note
    else:
        _delete_named_shapes(slide, "CsCol3Note")

    takeaway = _lead_thread_meaning(threads) or "This campaign delivered measurable results."
    values["CS_TAKEAWAY"] = takeaway
    values["CS_SOURCE"] = f"Source: Premion Website Attribution data{f', {period}' if period else ''}."

    if white_label:
        # CS_EYEBROW is pure Python composition (vertical noun + date) and
        # never carries model text -- the only key skipped. CS_HEADLINE is
        # deliberately NOT skipped: it's a model-drafted thread `head`
        # (real 2026-09-15 finding, testing this against a thread whose own
        # head named the client -- "Mattress Warehouse saw strong response"
        # shipped unredacted the first time this was written, since the
        # branch above treats CS_HEADLINE as "already the white-label
        # content" when it's really just "not the bare client name," which
        # isn't the same guarantee).
        for key in list(values.keys()):
            if key == "CS_EYEBROW":
                continue
            values[key] = white_label_text(values[key], name)

    _fill_tokens(slide, values)
    if blank_tiles:
        _reflow_tile_row(slide, blank_tiles, tile_names=_CS_TILE_ROW)


def build_summary_slide(template_path, output_path, *, attribution, delivery, client_name,
                        threads, accepted_optimizations, include_conversions=False, footnote=None):
    """The one-slide summary deliverable -- report:summary, isolated via
    `build_single_slide`. Never makes an API call; `threads` is always a
    STORED value (a fresh report's own draft, or a logged report's `report_
    json["draft"]["threads"]` on the regenerate-later path)."""
    return build_single_slide(template_path, "report:summary", output_path, lambda slide: (
        fill_summary_slide(slide, attribution, delivery, client_name, threads,
                          accepted_optimizations, include_conversions=include_conversions,
                          footnote=footnote)))


def build_case_study_slide(template_path, output_path, *, attribution, delivery, threads,
                           accepted_optimizations, client_name, vertical=None,
                           white_label=True, budget=None):
    """The vault-bound case-study deliverable -- report:case_study,
    isolated via `build_single_slide`. `budget` (a linked proposal's own
    full-flight total, or None) drives CS_TILE_3's conditional cost-per-
    visitor -- threaded through as a private attribute on `attribution`
    rather than a new `fill_case_study_slide` parameter of its own, since
    it's the one value on this slide that comes from neither the export
    nor the threads and every other value-source on this slide already
    hangs off `attribution`/`delivery`."""
    attribution = types.SimpleNamespace(**vars(attribution))
    attribution._case_study_budget = budget
    return build_single_slide(template_path, "report:case_study", output_path, lambda slide: (
        fill_case_study_slide(slide, attribution, delivery, threads, accepted_optimizations,
                             client_name, vertical=vertical, white_label=white_label)))


def _shrink_bullet_box_to_fit(slide, shape_name):
    """Measured shrink-to-fit for a bullet box, by real shape NAME --
    HIGHLIGHTBullets, TAKEAWAYBullets, WhatsNextBullets. Real finds
    (2026-09-11, rendering WAEPA), in two rounds: first What's Next
    overflowed (it has no cap -- the rework's own "every action" rule
    means a multi-goal report routinely produces more items than a single-
    goal one does); the first fix assumed HIGHLIGHTBullets/TAKEAWAYBullets
    were safe because they're capped at 4 and the template's own paragraph
    slots are sized for that many -- WRONG, a second render (4 real,
    substantial takeaways, not the earlier draft's shorter ones) overflowed
    TAKEAWAYBullets too. The cap bounds ITEM COUNT, not the text each item
    actually carries -- a real takeaway can run several lines. All three
    boxes get the same treatment now, not just the one that overflowed
    first.

    Reuses `assembly.fit_text_frame` -- the SAME measured shrink Campaign
    Specs and the case-study flatten pass already use -- rather than
    building a second one. Deliberately unconditional (not gated on the
    shape carrying its own `<a:normAutofit>`, the way `assembly.py`'s own
    `_apply_flattened_fit` is): none of these three boxes were authored
    with autofit at all, so gating on it would just never fire.
    """
    shape = assembly._find_shape_by_name(slide.shapes, shape_name)
    if shape is None or shape.height is None or shape.width is None:
        return
    body_pr = shape.text_frame._txBody.find(assembly.qn("a:bodyPr"))
    t_ins = assembly._inset(body_pr, "tIns", assembly._DEFAULT_CELL_INSET) if body_pr is not None else assembly._DEFAULT_CELL_INSET
    b_ins = assembly._inset(body_pr, "bIns", assembly._DEFAULT_CELL_INSET) if body_pr is not None else assembly._DEFAULT_CELL_INSET
    l_ins = assembly._inset(body_pr, "lIns", assembly._DEFAULT_SIDE_INSET) if body_pr is not None else assembly._DEFAULT_SIDE_INSET
    r_ins = assembly._inset(body_pr, "rIns", assembly._DEFAULT_SIDE_INSET) if body_pr is not None else assembly._DEFAULT_SIDE_INSET
    available = int((shape.height - t_ins - b_ins) * assembly._FLATTENED_FIT_MARGIN)
    width = shape.width - l_ins - r_ins
    if available > 0 and width > 0:
        assembly.fit_text_frame(shape.text_frame, available, width)


def _place_image(slide, region_shape, label_shape, png_bytes):
    """Place png_bytes into region_shape's footprint, then delete BOTH the
    region rectangle and its label -- README's own two-shape convention.
    Does nothing (leaving both template shapes in place) when there's
    nothing to plot, matching assembly.place_targeting_map's own
    None-does-nothing rule -- a slide with no chart data at all is a
    template defect worth seeing, not one this silently patches over."""
    if not png_bytes:
        return
    left, top, width, height = region_shape.left, region_shape.top, region_shape.width, region_shape.height
    slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, width, height)
    region_shape._element.getparent().remove(region_shape._element)
    label_shape._element.getparent().remove(label_shape._element)
