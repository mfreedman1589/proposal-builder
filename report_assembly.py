"""report_assembly.py -- fills REPORT_MASTER_v0_2.pptx from a parsed
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
from datetime import date, timedelta
from urllib.parse import urlparse

from pptx import Presentation
from pptx.util import Emu

import assembly
import geo_resolver
import market_lookup
import report_charts
import slide_map

DELIVERY_SET_RE = re.compile(r"^\s*delivery_set\s*:\s*true\s*$", re.IGNORECASE)

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


def geography_label(attribution):
    """The Geography tile's one short value. 2026-09-06 correction: a tile
    holds one short value -- if a value would run past one line, it's the
    wrong container, not a fit problem. Up to 2 real market names are
    listed; 3 or more collapses to "N markets" and the full list moves to
    the recap's Audience bullets instead (`geography_overflow_bullet`) --
    never a truncated "X, Y, +N more" squeezed into the tile."""
    names = _resolved_market_names(attribution)
    if names:
        if len(names) > _GEOGRAPHY_TILE_MAX_MARKETS:
            return f"{len(names)} markets"
        return ", ".join(names)
    if attribution.market_hint:
        return _STATION_MARKET_NAMES.get(attribution.market_hint, attribution.market_hint)
    return None


def geography_overflow_bullet(attribution):
    """The full market list, as its own Audience-bullets line -- produced
    only when `geography_label` collapsed to "N markets" (more than
    `_GEOGRAPHY_TILE_MAX_MARKETS`). None when there's nothing to overflow."""
    names = _resolved_market_names(attribution)
    if names and len(names) > _GEOGRAPHY_TILE_MAX_MARKETS:
        return "Markets: " + ", ".join(names)
    return None


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


def top_url_rows(attribution, limit=8):
    buckets = {}
    for url, visitors in (attribution.by_url or {}).items():
        buckets[_bucket_url(url)] = buckets.get(_bucket_url(url), 0) + int(visitors or 0)
    ordered = sorted(buckets.items(), key=lambda kv: kv[1], reverse=True)[:limit]
    total = sum(buckets.values()) or 1
    return [{"label": label, "visitors": _int(visitors),
            "share": f"{visitors / total * 100:.0f}%"} for label, visitors in ordered]


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


def intent_summary_rows(attribution):
    """[{intent, label, visits, share, _visits_raw}] -- one row per intent
    class that actually occurs, biggest first, plus the counts the model is
    given to reason about. Never emits a class with zero visits: a table row
    reading "Purchase intent 0 0%" invites a client question whose answer is
    "that category doesn't apply to your site."
    """
    tallies = {}
    noise = 0
    for url, visitors in (attribution.by_url or {}).items():
        count = int(visitors or 0)
        if _is_noise_url(url):
            noise += count
            continue
        tallies[classify_url_intent(url)] = tallies.get(classify_url_intent(url), 0) + count
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
    for intent, count in tallies.items():
        if not count:
            continue
        if intent != "other" and round(count / total * 100) < 1:
            other_extra += count
        else:
            folded[intent] = folded.get(intent, 0) + count
    if other_extra:
        folded["other"] = folded.get("other", 0) + other_extra
    rows = [{"intent": intent,
            "label": URL_INTENT_LABELS[intent],
            "visits": _int(count),
            "share": f"{count / total * 100:.0f}%",
            "_visits_raw": count,
            "_share_raw": count / total}
           for intent, count in folded.items() if count]
    rows.sort(key=lambda r: -r["_visits_raw"])
    return rows


def intent_facts(attribution):
    """The intent half of the facts payload Phase 4 hands the model, next to
    the campaign's goals. Deliberately numbers only -- no wording, no claim
    about alignment. `noise_visits` is reported rather than hidden so a
    later reader can tell a small total from a filtered one."""
    tallies = {}
    for url, visitors in (attribution.by_url or {}).items():
        if _is_noise_url(url):
            continue
        intent = classify_url_intent(url)
        tallies[intent] = tallies.get(intent, 0) + int(visitors or 0)
    total = sum(tallies.values()) or 1
    ordered = sorted(((i, c) for i, c in tallies.items() if c), key=lambda kv: -kv[1])
    return {
        "classes": [{"intent": intent, "label": URL_INTENT_LABELS[intent],
                     "visits": count, "share": count / total}
                    for intent, count in ordered],
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


def top_zip_rows(attribution, limit=10):
    """The zip table's rows -- "where the campaign worked best", not
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

    The area label is left blank (never a raw market key) when no market
    lookup is installed, rather than guessing.
    """
    zips_all = list(attribution.by_zip or [])
    if not zips_all:
        return []
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

    out = []
    for row, outperformer in chosen:
        multiple = f"{row.attributed_rate / baseline:.2f}x" if baseline else "--"
        out.append({
            "zip": row.label,
            "area": zip_to_market.get(row.label, ""),
            "share": f"{share_pct(row):.1f}%",
            "rate": _pct(row.attributed_rate),
            "multiple": multiple,
            "outperformer": outperformer,
        })
    return out


def pick_breakdown_dimension(attribution):
    """(dimension_label, rows) -- the ONE breakdown table this slide shows.
    Market wins whenever the export has more than one; else audience.
    Creative only pre-empts audience when it's genuinely the standout --
    at least two creatives ran AND the leader's attributed rate beats the
    runner-up's by 50%+ -- a deterministic stand-in for the judgment call
    REPORT_MASTER_README.md leaves to "genuinely the story"; Phase 4's
    Claude-driven synthesis is the natural place to replace this heuristic
    with real narrative judgment, not this module's job to perfect now.
    """
    if len(attribution.by_market) > 1:
        return "Market", attribution.by_market
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
    once, here, not re-derived a second time from the unsorted input."""
    ordered = sorted(rows, key=lambda r: r.attributed_impressions, reverse=True)[:limit]
    return [{"label": r.label, "delivered": _int(r.delivered_impressions),
            "attributed": _int(r.attributed_impressions), "rate": _pct(r.attributed_rate),
            "_attributed_raw": r.attributed_impressions}
           for r in ordered]


def default_highlight_bullets(attribution, delivery):
    """Up to 4 (head, detail) pairs, grounded strictly in computed facts --
    deliberately plain, not client-ready prose. Phase 4 (ATTRIBUTION_REPORT_
    PLAN.md) replaces this with Claude synthesis over the same facts this
    function already computes; this is the swappable placeholder that
    keeps Phase 3's deck-assembly path exercised end to end in the
    meantime."""
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
    """Same swappable-placeholder shape as default_highlight_bullets --
    Phase 4 replaces this with Claude's own takeaways synthesis."""
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
                      takeaway_bullets=None, headline_notes=None):
    """Fill REPORT_MASTER_v0_2.pptx and save to output_path. Returns
    (output_path, warnings) -- warnings is a list of plain-language strings
    from any table that still overflows its measured floor even after the
    row cap and the shrink-to-fit pass (see `_fill_delivery_recap`'s own
    STOPGAP-turned-real-fix note); empty in the ordinary case.

    No-proposal mode only (Phase 3-4) -- every content decision below either
    comes straight off the parsed export or is a required keyword argument;
    there is no silent proposal-linked path here yet (that's Phase 5).
    `flight_label` defaults to "" (Phase 3 never derives a real campaign
    flight separate from the report period) -- Phase 5 passes a real one
    from a linked proposal's own flight.

    `goals_bullets`/`whats_next_bullets` are REQUIRED -- no export signal
    produces them, so this raises MissingTokenError rather than defaulting
    if either is empty. `audience_bullets`/`highlight_bullets`/
    `takeaway_bullets` default to this module's own computed facts when not
    overridden; `headline_notes` (a dict of the three *_HEADLINE_NOTE
    tokens, keyed "attribution"/"url"/"zip") likewise defaults to a plain
    computed sentence per slide when not supplied.
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
                    "report:url_report", "report:zip_analysis", "report:takeaways"):
        if required not in keys:
            raise MissingTokenError(f"template is missing a slide tagged key: {required}")
    for required in ("report:delivery_recap", "report:delivery_breakdown"):
        if required not in keys:
            raise MissingTokenError(f"template is missing a slide tagged key: {required}")

    # The delivery SET -- every slide whose notes carry `delivery_set: true`
    # -- is dropped wholesale when no delivery file was uploaded. Driven off
    # the marker in the deck rather than a hardcoded key list, so a third
    # delivery slide added to a future template needs no code change here.
    # Dropped highest-index-first: deleting a slide shifts every index after
    # it, and collecting the indices up front then deleting low-to-high
    # would delete the wrong slides.
    drop_keys = []
    if delivery is None:
        drop_keys = ["report:delivery_recap", "report:delivery_breakdown"]
    elif not delivery_breakdown_applies(delivery):
        # A delivery file with one geo option and one creative has no
        # breakdown to show -- drop that slide alone, keep the recap.
        drop_keys = ["report:delivery_breakdown"]
    for key in drop_keys:
        if not _is_delivery_set(prs.slides[keys[key]]):
            raise MissingTokenError(
                f"{key} is missing its `delivery_set: true` notes line -- "
                f"can't tell whether it's safe to drop.")
    for index in sorted((keys[k] for k in drop_keys), reverse=True):
        assembly.delete_slide(prs, index)
    if drop_keys:
        keys = _slide_by_key(prs)  # indices shifted

    _fill_recap(prs.slides[keys["report:recap"]], attribution, client_name, report_title,
               goals_bullets, audience_bullets, flight_label)
    _fill_highlights(prs.slides[keys["report:highlights"]], attribution, delivery, highlight_bullets)
    warnings = []
    if delivery is not None:
        warnings += _fill_delivery_recap(prs.slides[keys["report:delivery_recap"]], delivery)
    if delivery is not None and delivery_breakdown_applies(delivery):
        warnings += _fill_delivery_breakdown(
            prs.slides[keys["report:delivery_breakdown"]], delivery)
    warnings += _fill_attribution_breakdown(prs.slides[keys["report:attribution_breakdown"]], attribution,
                                            (headline_notes or {}).get("attribution"))
    warnings += _fill_url_report(prs.slides[keys["report:url_report"]], attribution,
                                 (headline_notes or {}).get("url"))
    warnings += _fill_zip_analysis(prs.slides[keys["report:zip_analysis"]], attribution,
                                   (headline_notes or {}).get("zip"))
    _fill_takeaways(prs.slides[keys["report:takeaways"]], attribution, delivery,
                   takeaway_bullets, whats_next_bullets)

    prs.save(output_path)
    return output_path, [w for w in warnings if w]


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


def _fill_named_table(slide, table_name, token, rows, fields):
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
                flight_label):
    period = _date_range_label(attribution.flight_start, attribution.flight_end)
    if period is None:
        raise MissingTokenError("report:recap/REPORT_PERIOD_LABEL: the export carries no "
                                "weekly/monthly trend tab to derive a period from -- "
                                f"warnings: {attribution.warnings}")
    geo = geography_label(attribution)
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
    overflow = geography_overflow_bullet(attribution)
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


def _fill_highlights(slide, attribution, delivery, highlight_bullets):
    headline_impressions = delivery.delivered_impressions if delivery else attribution.delivered_impressions
    _fill_tokens(slide, {
        "HEADLINE_IMPRESSIONS": _int(headline_impressions),
        "HEADLINE_UNIQUE_VISITORS": _int(attribution.attributed_unique_visitors),
        "HEADLINE_ATTRIBUTED_RATE": _pct(attribution.attributed_rate),
    })
    items = highlight_bullets or default_highlight_bullets(attribution, delivery)
    _fill_head_detail_bullets(slide, "HIGHLIGHTBullets", items, max_items=4)


TOP_PUBLISHERS_ROW_CAP = 5  # 2026-09-06: a design rule (what a client reads), not a fit guess
TOP_CREATIVES_ROW_CAP = 3


def _fill_delivery_recap(slide, delivery):
    """v0_3: five tiles (CTV share joins the four originals), the publisher
    table alone in the left column, and the chart region now carrying the
    DAYPART breakdown -- the publisher bar chart was dropped because it and
    the publisher table said the same thing, and the table carries VCR too.
    CreativeTable moved off this slide to report:delivery_breakdown.
    """
    _fill_tokens(slide, {
        "DELIVERED_IMPRESSIONS": _int(delivery.delivered_impressions),
        "VCR": _pct(delivery.vcr, 1),
        "FREQUENCY": f"{delivery.frequency:.1f}",
        "UNIQUES": _int(delivery.uniques),
        "DELIVERY_NARRATIVE": (f"Delivered {_int(delivery.delivered_impressions)} impressions "
                               f"at a {_pct(delivery.vcr, 1)} completion rate."),
    })
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


def _fill_delivery_breakdown(slide, delivery):
    """report:delivery_breakdown (v0_3) -- two CONDITIONAL tables stacked in
    the left column, each with its own named header, plus a VCR-by-creative
    chart in its own column.

    A dimension with one value or none is not a breakdown, so its table and
    its header are deleted together. No reflow is needed or wanted: the
    template stacks them, so a surviving table simply has space beneath it
    (Matt's own layout note for v0_3). The caller decides whether this
    slide exists at all -- see `delivery_breakdown_applies`.
    """
    geo = list(delivery.by_geo or [])
    creatives = sorted(delivery.by_creative, key=lambda c: c[1],
                       reverse=True)[:TOP_CREATIVES_ROW_CAP]
    total = delivery.delivered_impressions or sum(c for _l, c in geo) or 0

    dimensions = []
    warnings = []
    if len(geo) > 1:
        dimensions.append("geography")
        warnings += _fill_named_table(
            slide, "DeliveryByGeoTable", "DELIVERY_BY_GEO_ROWS",
            [{"label": label,
              "impressions": _int(count),
              # Real, impression-weighted VCR from the flight-detail tab --
              # the geo tab itself carries impressions only. "--" only when
              # that tab genuinely can't supply one, never a borrowed number.
              "vcr": (_pct(delivery.geo_vcr[label], 1)
                      if label in delivery.geo_vcr else "--")}
             for label, count in sorted(geo, key=lambda g: -g[1])],
            ["label", "impressions", "vcr"])
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
        "DELIVERY_BREAKDOWN_NARRATIVE": ("; ".join(narrative_bits) + "."
                                         if narrative_bits else ""),
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


def _fill_attribution_breakdown(slide, attribution, headline_note):
    dimension, rows = pick_breakdown_dimension(attribution)
    table_rows = breakdown_rows(rows)
    _fill_tokens(slide, {
        "ATTRIBUTION_HEADLINE_NOTE": headline_note or (
            f"Attribution broke out by {dimension.lower()} -- "
            f"{_pct(attribution.attributed_rate)} overall attributed rate."),
        "BREAKDOWN_DIMENSION_LABEL": dimension,
        "ATTRIBUTION_NARRATIVE": (
            f"{table_rows[0]['label']} led all {dimension.lower()}s at a "
            f"{table_rows[0]['rate']} attributed rate." if table_rows else ""),
    })
    warnings = _fill_named_table(slide, "BreakdownTable", "BREAKDOWN_ROWS", table_rows,
                                ["label", "delivered", "attributed", "rate"])
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


def _fill_url_report(slide, attribution, headline_note):
    intent_rows = intent_summary_rows(attribution)[:_INTENT_ROWS_CAP]
    url_rows = top_url_rows(attribution, limit=_URL_ROWS_IN_TEMPLATE)
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
        # Phase 4 replaces this with the model connecting these same
        # aggregates to the campaign's stated goals. Until then it states
        # the mix without claiming alignment to a goal nobody has supplied.
        "URL_INTENT_NARRATIVE": _url_intent_narrative(intent_rows, url_rows),
    })
    warnings = _fill_named_table(slide, "IntentSummaryTable", "INTENT_SUMMARY_ROWS",
                                 intent_rows, ["label", "visits", "share"])
    warnings += _fill_named_table(slide, "TopUrlTable", "TOP_URL_ROWS", url_rows,
                                  ["label", "visitors", "share"])
    return warnings


def _fill_zip_analysis(slide, attribution, headline_note):
    rows = top_zip_rows(attribution)
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
        "ZIP_NARRATIVE": narrative,
    })
    warnings = _fill_named_table(slide, "TopZipTable", "TOP_ZIP_ROWS", rows,
                                ["zip", "area", "share", "rate", "multiple"])
    region = _shape(slide, "MapRegion")
    label_shape = _shape(slide, "MapRegionLabel")
    # Attributed RATE is the weight, not impressions: the slide's question
    # is where the campaign worked hardest, not where it delivered most --
    # the same distinction the table's own ranking rule draws.
    png, missing_polygons = report_charts.render_zip_map(
        {r.label: r.attributed_rate for r in attribution.by_zip
         if r.delivered_impressions > 0},
        region.width, region.height)
    _place_image(slide, region, label_shape, png)
    if missing_polygons:
        # Reported, never silent -- same discipline as the Shopify pixel
        # drop on the URL slide. These are drawn as centroid dots rather
        # than omitted, so the map still accounts for every zip; the
        # warning exists so a WHOLE STATE missing its ZCTA file (the
        # `build_zcta_boundaries.py --add <ST>` case) is visible as a big
        # number rather than looking like a slightly sparse map.
        warnings.append(
            f"{len(missing_polygons)} zip code(s) have no ZCTA boundary and are shown "
            f"as points rather than shaded areas "
            f"(e.g. {', '.join(sorted(missing_polygons)[:4])}). PO-box-only zips and "
            f"retired ZCTAs are expected; a large count usually means a state hasn't "
            f"been built -- see build_zcta_boundaries.py --add.")
    return warnings


def _fill_takeaways(slide, attribution, delivery, takeaway_bullets, whats_next_bullets):
    items = takeaway_bullets or default_takeaway_bullets(attribution, delivery)
    _fill_head_detail_bullets(slide, "TAKEAWAYBullets", items, max_items=4)
    _fill_bullets(slide, "WHATS_NEXT_BULLETS", whats_next_bullets)


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
