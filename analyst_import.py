"""Auto-Sales Analyst facts JSON -- parse, gate and summarize.

The Analyst (a separate repo) exports one JSON per report: which vehicles the
campaign's attributed website visitors shopped, how many of those have since
left the dealers' live inventory, and where on the sites those visits went.
This module reads it into plain dicts; it never touches Streamlit or the
database, so tests can import it directly.

Field names are the Analyst's own (it shipped first; this repo adapts to it).
See ATTRIBUTION_REPORT_PLAN.md item 7 for the full shape and matching rules.
"""
import json
import re
from datetime import date

# v2 is additive only (new keys, no renamed or redefined ones), so a v1
# reader stays correct on it; v2's extras are used when present.
SUPPORTED_SCHEMA_VERSIONS = (1, 2)

# 4 named rows + the "All other pages" remainder leaves the models table
# below it at full type size on the v0_14 slide (5 named rows shrank it).
TRAFFIC_MIX_ROW_CAP = 4
TOP_MODEL_ROW_CAP = 5
PAYLOAD_MODEL_CAP = 10
PAYLOAD_MISSED_CAP = 5

_VDP_CATEGORIES = ("NEW_VDP", "USED_VDP")


class AnalystParseError(Exception):
    """A plain-language, rep-facing reason the file can't be used."""


def _load(source):
    if isinstance(source, (bytes, bytearray)):
        text = bytes(source).decode("utf-8-sig")
    else:
        with open(source, encoding="utf-8-sig") as fh:
            text = fh.read()
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise AnalystParseError(
            "This file isn't readable as an Auto-Sales Analyst facts export -- re-export the "
            "facts JSON from the Analyst and upload that file.") from exc


def _iso_date(value, what):
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError) as exc:
        raise AnalystParseError(
            f"The Analyst file's {what} isn't a date the app can read ({value!r}) -- "
            f"re-export it from the Analyst.") from exc


def parse_analyst_facts(source):
    """`source` is a path or raw bytes. Returns a JSON-safe dict (dates as ISO
    strings, so it can sit in session_state and a logged report unchanged).
    Raises AnalystParseError with a message a rep can act on."""
    data = _load(source)
    if not isinstance(data, dict) or "totals" not in data or "meta" not in data:
        raise AnalystParseError(
            "This JSON isn't an Auto-Sales Analyst facts export (no totals/meta section) -- "
            "make sure it's the \"facts\" download from the Analyst.")
    version = data.get("schema_version")
    if version not in SUPPORTED_SCHEMA_VERSIONS:
        raise AnalystParseError(
            f"This Analyst file uses a newer format (version {version}) than the report "
            f"builder understands yet. Report an issue so it can be updated.")
    meta = data.get("meta") or {}
    period = meta.get("analysis_period") or {}
    start = _iso_date(period.get("start"), "analysis start date")
    end = _iso_date(period.get("end"), "analysis end date")
    totals = data.get("totals") or {}
    if not totals.get("traffic_mix") and not totals.get("vehicles_shopped"):
        raise AnalystParseError(
            "The Analyst file has no traffic or vehicle data in it -- re-run the Analyst for "
            "this report's month and export again.")
    return {
        "schema_version": version,
        "group_name": meta.get("group_name") or meta.get("report_id") or "",
        "is_group": bool(meta.get("is_group")),
        "period_start": start.isoformat(),
        "period_end": end.isoformat(),
        "inventory_scanned_at": meta.get("inventory_scanned_at"),
        "lookback_days": meta.get("lookback_days"),
        "influence_benchmark_visits": meta.get("influence_benchmark_visits"),
        "sites": [{"site_id": s.get("site_id"), "dealer_name": s.get("dealer_name") or "",
                   "domain": s.get("domain") or "",
                   "is_group_site": s.get("is_group_site")}
                  for s in (data.get("sites") or []) if isinstance(s, dict)],
        "totals": totals,
        "by_site": data.get("by_site") or {},
    }


def _as_date(value):
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def periods_overlap(analyst, report_start, report_end):
    """True when the Analyst's own analysis period shares at least one day
    with the report's. Bounds may be dates or ISO strings. Either report
    bound missing -> True: nothing to contradict."""
    report_start, report_end = _as_date(report_start), _as_date(report_end)
    if report_start is None or report_end is None:
        return True
    start = date.fromisoformat(analyst["period_start"])
    end = date.fromisoformat(analyst["period_end"])
    return start <= report_end and report_start <= end


# The Analyst's own guidance (Sales Assist) is to run it 15 and/or 30 days
# after month end; inventory turns over within ~3 months, so a much later
# scan counts ordinary turnover as "sold". 60 = well past the recommended
# 30-day run, still short of the 3-month turnover point.
LOOKBACK_ADVISORY_DAYS = 60


def unusable_reason(analyst):
    """A rep-facing reason not to use this file at all, or None. The Sales
    Assist's own rule: a report returning 0 sold (with vehicles shopped)
    usually means the scanner couldn't read the dealer's site -- don't
    present it."""
    totals = analyst.get("totals") or {}
    if (totals.get("vehicles_shopped") or 0) > 0 and not totals.get("vehicles_sold"):
        return ("The Auto-Sales Analyst file shows 0 vehicles sold, which usually means it "
                "couldn't read the dealer's website -- it won't be used. Re-run the Analyst "
                "(its Dealer Inspire Fix, if needed) and upload the new file.")
    return None


def scan_delay_days(analyst):
    """Days between the end of the Analyst's analysis period and the date it
    checked the dealer sites (`inventory_scanned_at`), or `lookback_days`
    when the scan date is missing. None when neither is known."""
    scanned = _as_date(analyst.get("inventory_scanned_at"))
    end = _as_date(analyst.get("period_end"))
    if scanned and end:
        return (scanned - end).days
    days = analyst.get("lookback_days")
    return int(days) if isinstance(days, (int, float)) else None


def lookback_advisory(analyst):
    """A rep-facing advisory when the Analyst ran long after the month
    ended, or None. The file is still used -- the rep decides."""
    days = scan_delay_days(analyst)
    if days is not None and days > LOOKBACK_ADVISORY_DAYS:
        return (f"The Auto-Sales Analyst checked the dealer sites {days} days after the "
                f"period ended. Its own guidance is 15-30 days: the longer the wait, the more "
                f"ordinary inventory turnover shows up as \"sold\". Consider re-running it closer "
                f"to month end before sending.")
    return None


# Makes as they appear in dealer/client names, mapped to the Analyst's own
# canonical make. Only makes named in the dealership's own name count as its
# franchise -- a Ford store's used lot carrying a Honda doesn't make it a
# Honda franchise.
_MAKE_ALIASES = {
    "ford": "FORD", "lincoln": "LINCOLN", "chevrolet": "CHEVROLET", "chevy": "CHEVROLET",
    "gmc": "GMC", "buick": "BUICK", "cadillac": "CADILLAC", "toyota": "TOYOTA",
    "lexus": "LEXUS", "honda": "HONDA", "acura": "ACURA", "nissan": "NISSAN",
    "infiniti": "INFINITI", "hyundai": "HYUNDAI", "genesis": "GENESIS", "kia": "KIA",
    "subaru": "SUBARU", "mazda": "MAZDA", "volkswagen": "VOLKSWAGEN", "vw": "VOLKSWAGEN",
    "audi": "AUDI", "bmw": "BMW", "mini": "MINI", "mercedes": "MERCEDES-BENZ",
    "mercedes-benz": "MERCEDES-BENZ", "volvo": "VOLVO", "jeep": "JEEP", "ram": "RAM",
    "dodge": "DODGE", "chrysler": "CHRYSLER", "cdjr": "JEEP", "porsche": "PORSCHE",
    "jaguar": "JAGUAR", "land rover": "LAND ROVER", "mitsubishi": "MITSUBISHI",
    "tesla": "TESLA",
}


def franchise_makes(analyst, client_name=None):
    """Makes named in the client's or dealer sites' own names, sorted --
    Ted Britt Ford / Tedbritt Chevrolet / Tedbritt Lincoln Of Chantilly ->
    CHEVROLET, FORD, LINCOLN. Longer makes run together from a domain
    ("Tedbrittchevrolet") still match; short ones need a whole word."""
    names = [client_name or ""] + [s.get("dealer_name") or "" for s in analyst.get("sites") or []]
    text = " ".join(names).lower()
    found = set()
    for alias, make in _MAKE_ALIASES.items():
        # Short makes need a whole word ("Stafford" isn't Ford, "Dominion"
        # isn't Mini); longer ones may be run together from a domain.
        if len(alias) <= 4:
            if re.search(rf"\b{re.escape(alias)}\b", text):
                found.add(make)
        elif alias in text:
            found.add(make)
    return sorted(found)


def visitor_count_mismatch(analyst, attributed_unique_visitors):
    """(analyst_count, export_count) when both exist and disagree, else None.
    The Analyst reads the same Premion page-visit data the attribution export
    does (Ted Britt, Aug 2026: 2,671 on both), so a mismatch usually means the
    two files are for different campaigns or periods."""
    analyst_count = (analyst.get("totals") or {}).get("unique_visitors")
    if not analyst_count or not attributed_unique_visitors:
        return None
    if int(analyst_count) == int(attributed_unique_visitors):
        return None
    return int(analyst_count), int(attributed_unique_visitors)


_KEEP_UPPER = {"GMC", "BMW", "RAM", "KIA", "GLE", "GLC", "GLS", "SUV", "RS", "GT", "SC", "SS",
               "ST", "XL", "XLT", "LT", "LTZ", "RST", "AMG", "EV", "HD", "SD", "RTR"}


def _title_label(label):
    """The Analyst's canonical labels are uppercase ("FORD F-150"); a client
    table reads better in title case. Model codes with digits ("F-150",
    "CX-70") and known acronyms stay as written."""
    def fix(part, next_word):
        if not part or any(ch.isdigit() for ch in part) or part in _KEEP_UPPER:
            return part
        # A short letter code before a model number stays a code: "GX 460",
        # and "GX-460" (the next hyphen part, not the next word).
        if part.isalpha() and len(part) <= 3 and next_word[:1].isdigit():
            return part.upper()
        return part.capitalize()
    words = str(label).split()
    out = []
    for i, word in enumerate(words):
        parts = word.split("-")
        following = words[i + 1] if i + 1 < len(words) else ""
        out.append("-".join(fix(p, parts[j + 1] if j + 1 < len(parts) else following)
                            for j, p in enumerate(parts)))
    return " ".join(out)


def store_label(site):
    """A store's name for a client-facing slide: the Analyst's dealer_name,
    unless that is only its domain slug run together ("Tedbritttruckshop"
    for tedbritttruckshop.com, Ted Britt Aug 2026) -- then the domain,
    which at least reads as an address rather than a typo."""
    name = (site or {}).get("dealer_name") or ""
    domain = (site or {}).get("domain") or ""
    stem = domain.split(".")[0].lower()
    if domain and (not name or (" " not in name.strip() and name.strip().lower() == stem)):
        return domain
    return name or domain


def traffic_mix_rows(analyst, cap=TRAFFIC_MIX_ROW_CAP):
    """[{label, visits, share}] -- where attributed visits went on the sites,
    largest first. Shares are of total visits, so they sum to 100%. Beyond
    `cap` -- and always for the Analyst's own "OTHER" category, so the table
    never shows "Other" and "All other pages" side by side -- the remaining
    categories fold into one trailing "All other pages" row, so the table
    still accounts for every visit."""
    totals = analyst.get("totals") or {}
    mix = sorted((m for m in totals.get("traffic_mix") or [] if m.get("visits")),
                 key=lambda m: m["visits"], reverse=True)
    total = totals.get("visits_total") or sum(m["visits"] for m in mix)
    if not mix or not total:
        return []
    named = [m for m in mix if m.get("category") != "OTHER"]
    shown = named[:cap]
    rest = [m for m in mix if m not in shown]
    rows = [{"label": m.get("label") or m.get("category", ""), "visits": m["visits"]} for m in shown]
    if rest:
        rows.append({"label": "All other pages", "visits": sum(m["visits"] for m in rest)})
    for row in rows:
        row["share"] = row["visits"] / total
    return rows


def top_model_rows(analyst, cap=TOP_MODEL_ROW_CAP):
    """[{label, count, share}] -- the models attributed visitors shopped that
    have since sold, most units first. `share` is of all shopped-then-sold
    units."""
    totals = analyst.get("totals") or {}
    models = totals.get("sold_by_make_model") or totals.get("top_sellers") or []
    models = sorted((m for m in models if m.get("count")), key=lambda m: m["count"], reverse=True)
    sold = totals.get("vehicles_sold") or 0
    return [{"label": _title_label(m.get("label") or f"{m.get('make', '')} {m.get('model', '')}"),
             "count": m["count"], "share": (m["count"] / sold) if sold else None}
            for m in models[:cap]]


def vdp_visit_share(analyst):
    """Share of attributed visits that landed on a vehicle detail page, or
    None when the mix carries no VDP categories."""
    totals = analyst.get("totals") or {}
    mix = totals.get("traffic_mix") or []
    total = totals.get("visits_total") or sum(m.get("visits") or 0 for m in mix)
    vdp = sum(m.get("visits") or 0 for m in mix if m.get("category") in _VDP_CATEGORIES)
    return (vdp / total) if total and vdp else None


def model_family_key(model):
    """A model's family for matching shopped pages to sold vehicles: its
    first word, alphanumerics only, joining a short letter code to its
    number ("F-150"/"F 150 Lariat" -> F150, "GX-460"/"Gx 460" -> GX460,
    "SILVERADO 2500 HD" -> SILVERADO). Make + model FAMILY, never trim
    (ATTRIBUTION_REPORT_PLAN.md item 7)."""
    tokens = str(model or "").upper().replace("-", " ").split()
    if not tokens:
        return ""
    if len(tokens) > 1 and tokens[0].isalpha() and len(tokens[0]) <= 3 and tokens[1][:1].isdigit():
        return re.sub(r"[^A-Z0-9]", "", tokens[0] + tokens[1])
    return re.sub(r"[^A-Z0-9]", "", tokens[0])


def material_gap(high, low, floor=0.2):
    """The drafting prompt's own material-swing floor (20% relative)."""
    return bool(high) and low is not None and (high - low) / high >= floor


def payload_facts(analyst, client_name=None):
    """The compact, model-facing slice of the Analyst file for
    `report_assembly.build_facts_payload`. Totals only -- per-site figures
    wait for multi-dealer group reports (Phase 9); dealer names ride along so
    a thread can say "across the group's five sites". VINs are left out: no
    client sentence needs one, and every digit run in the payload widens the
    facts-only checker's allowed set."""
    totals = analyst.get("totals") or {}
    mix_total = totals.get("visits_total") or sum(
        m.get("visits") or 0 for m in totals.get("traffic_mix") or [])
    mix = [{"label": m.get("label") or m.get("category", ""), "visits": m["visits"],
            "share": (m["visits"] / mix_total) if mix_total else None}
           for m in sorted((m for m in totals.get("traffic_mix") or [] if m.get("visits")),
                           key=lambda m: m["visits"], reverse=True)]
    missed = (totals.get("missed_opportunities") or {}).get("vehicles") or []
    models = sorted(totals.get("sold_by_make_model") or totals.get("top_sellers") or [],
                    key=lambda m: m.get("count") or 0, reverse=True)
    return {
        "period_start": analyst["period_start"],
        "period_end": analyst["period_end"],
        "site_count": len(analyst.get("sites") or []),
        "dealer_names": [s["dealer_name"] for s in analyst.get("sites") or [] if s.get("dealer_name")],
        "lookback_days": analyst.get("lookback_days"),
        "inventory_scanned_at": analyst.get("inventory_scanned_at"),
        "days_scanned_after_period_end": scan_delay_days(analyst),
        "franchise_makes": franchise_makes(analyst, client_name),
        "unique_visitors": totals.get("unique_visitors"),
        "visits_total": totals.get("visits_total"),
        "vdp_visit_share": vdp_visit_share(analyst),
        "traffic_mix": [{"label": r["label"], "visits": r["visits"], "visit_share": r["share"]}
                        for r in mix],
        # "Viewed", never "shopped" (positioning, 2026-10-01): the JSON's own
        # field names stay as the Analyst shipped them; only these model-
        # facing names change, since the model echoes whatever it reads.
        "vehicles_viewed": totals.get("vehicles_shopped"),
        "vehicles_sold_since": totals.get("vehicles_sold"),
        "vehicles_sold_since_new": totals.get("vehicles_sold_new"),
        "vehicles_sold_since_used": totals.get("vehicles_sold_used"),
        "vehicles_still_available": (totals.get("shopped_vehicle_status") or {}).get("available"),
        "vehicles_status_unconfirmed": (totals.get("shopped_vehicle_status") or {}).get("unverified"),
        "look_to_book_pct": totals.get("look_to_book_pct"),
        "look_to_book_pct_new": totals.get("look_to_book_pct_new"),
        "look_to_book_pct_used": totals.get("look_to_book_pct_used"),
        "est_value_sold": totals.get("est_revenue_sold"),
        "est_total_value_viewed": totals.get("est_pipeline_value"),
        "sold_by_make": [{"make": _title_label(m.get("make", "")), "count": m.get("count")}
                         for m in (totals.get("sold_by_make") or [])[:5]],
        "top_models_sold": [{"model": _title_label(m.get("label", "")), "count": m.get("count")}
                            for m in models[:PAYLOAD_MODEL_CAP]],
        "sold_by_price_tier": [{"tier": t.get("label"), "count": t.get("count")}
                               for t in totals.get("sold_by_price_tier") or []],
        "missed_opportunities": [{"vehicle": _title_label(v.get("label", "")),
                                  "visits": v.get("attributed_visits", v.get("visits"))}
                                 for v in missed[:PAYLOAD_MISSED_CAP]],
        # Deliberately absent: sold_above_benchmark, sold_high_influence, the
        # benchmark/threshold fields and look_to_book_by_visit_band. Look-to-
        # Book is flat across visit counts in the real data (58.6/50.7/56.4/
        # 54.4%), so no visit-count influence claim is supportable; the bands
        # are kept in period_facts (`visit_band_facts`) for testing that later.
    }


def visit_band_facts(analyst, min_viewed=30):
    """v2's look_to_book_by_visit_band, for the logged report's period_facts
    only -- never shown, never cited. Bands with fewer than `min_viewed`
    vehicles are dropped as too thin to compare."""
    bands = (analyst.get("totals") or {}).get("look_to_book_by_visit_band") or []
    return [{"band": b.get("band"), "viewed": b.get("viewed"), "sold": b.get("sold"),
             "look_to_book_pct": b.get("look_to_book_pct")}
            for b in bands if (b.get("viewed") or 0) >= min_viewed]


def site_domains(analyst):
    """{site_id: domain}. Stores are identified by DOMAIN everywhere a store
    is joined or remembered -- site_id and by_site keys are slugs of the
    display name and change when a name is edited."""
    return {s.get("site_id"): s.get("domain") for s in analyst.get("sites") or []}


