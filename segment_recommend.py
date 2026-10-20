"""segment_recommend.py -- deterministic audience recommendations from the
ranked, tagged catalog (segment_catalog.py, overlaid by audience_catalog).

Pure: takes catalog records, returns plain dicts. No Streamlit, no DB, no
model call -- the Audience finder shows these directly, and draft-from-notes
and the finder's Suggest mode hand them to Claude as the starting point.

The rules (2026-10-20, refined the same day):
- Recommend CORE segments ranked by score WITHIN the matched vertical tag,
  never across the whole catalog. At most 6 picks.
- Near-duplicates collapse to the highest-scoring one ("Business Owner
  General" / "Business Owners Executives" / "... Small" is one pick); the
  rest are listed as "similar".
- "Often paired with" shows a modifier only when the pair shares 3+ lines,
  and never an income cap ("Less than ...") beside a B2B segment.
- A modifier is never recommended on its own unless the request asks for a
  demographic or income target.
- B2B is slotted: one business-owner/size segment, one decision maker
  matching the product (IT for VoIP/IT, facilities for building services,
  HR for staffing/benefits, insurance for commercial insurance), one
  industry or occupation segment per target industry named -- the nearest
  available segment, labeled "nearest match", when the catalog has none for
  that industry. Remaining B2B picks must match the request's own terms.
- Proven custom recipes for the vertical are listed separately, labeled.
"""
import re

MAX_PICKS = 6
MIN_PAIRING_LINES = 3

TAG_FOR_VERTICAL = {
    "healthcare": "Healthcare", "education": "Education", "auto": "Auto",
    "retail": "Retail", "legal": "Legal", "home_improvement": "Home Services",
    "banking": "Financial", "travel": "Tourism/Travel",
    "entertainment": "Entertainment/Events", "dining_qsr": "Restaurants/Food",
}

# Phrases that make a request B2B, matched as whole words, lowercase.
B2B_SYNONYMS = [
    "b2b", "smb", "smbs", "small business", "small businesses", "business owner",
    "business owners", "businesses", "decision maker", "decision makers", "office",
    "offices", "commercial", "contractor", "contractors", "facilities", "facility",
    "manufacturing", "manufacturers", "voip", "staffing", "payroll", "managed services",
    "it services", "office products", "office supplies", "construction supply",
    "construction supplies", "enterprise", "employers", "commercial insurance",
    "business insurance", "business services",
]
# Matched case-sensitively in the ORIGINAL text -- "it" is a pronoun.
B2B_ACRONYMS = ["IT", "HR"]

# Words in a request (notes or a search) that point at a vertical tag. When
# several match, the one mentioned FIRST leads -- "a casino hotel" is an
# entertainment request first, a travel one second.
TAG_WORDS = {
    "Healthcare": ["healthcare", "health care", "medical", "health", "hospital", "hospitals",
                   "clinic", "doctor", "dental", "dentist", "urgent care"],
    "Education": ["education", "school", "college", "university", "enrollment"],
    "Auto": ["automotive", "dealership", "car dealer", "auto dealer"],
    "Retail": ["retail", "retailer", "store", "shop"],
    "Legal": ["legal", "lawyer", "attorney", "law firm", "personal injury"],
    "Home Services": ["home services", "home improvement", "hvac", "roofing", "plumbing",
                      "remodeling"],
    "Financial": ["financial", "bank", "banking", "credit union", "finance"],
    "B2B": ["b2b"],
    "Entertainment/Events": ["entertainment", "casino", "casinos", "concert", "concerts",
                             "venue", "theater", "theatre", "festival", "festivals", "events",
                             "comedy", "museum", "nightlife"],
    "Tourism/Travel": ["tourism", "travel", "cvb", "destination", "visitors", "visitor",
                       "getaway", "getaways", "vacation", "hotel", "resort"],
    "Restaurants/Food": ["restaurant", "restaurants", "qsr", "bar", "brewery", "dining",
                         "food", "pizza"],
    "Sports": ["sports", "golf", "team", "league", "stadium"],
    "Outdoors/Recreation": ["outdoors", "outdoor", "hunting", "fishing", "camping", "boating"],
    "Fitness/Wellness": ["gym", "fitness", "club membership", "yoga", "wellness"],
    "Insurance": ["insurance"],
    "Real Estate/Movers": ["realtor", "realtors", "apartments", "apartment", "builder",
                           "real estate", "home builder", "movers"],
    "Political/Advocacy": ["candidate", "campaign for office", "ballot", "election", "voters",
                           "political"],
    "Military/Government": ["military", "veterans", "government"],
    "Nonprofit/Charity": ["donor", "donors", "fundraising", "nonprofit", "non profit",
                          "charity"],
    "Pets": ["pet", "pets", "veterinary", "vet clinic"],
    "Telecom/Tech": ["isp", "wireless", "cable", "internet", "broadband", "fiber"],
}

# Decision-maker function by what's being sold.
FUNCTION_TERMS = [
    ("it", ["voip", "phone system", "phone systems", "telecom", "network", "networking",
            "software", "cloud", "cyber", "cybersecurity", "managed services", "computer",
            "computers", "it services", "internet", "fiber", "broadband", "data center",
            "saas", "tech support"], ["IT"]),
    ("hr", ["staffing", "benefits", "payroll", "recruiting", "recruitment", "hiring",
            "human resources", "peo", "employee", "employees", "workforce", "temp agency"],
     ["HR"]),
    ("facilities", ["janitorial", "cleaning", "building services", "facility", "facilities",
                    "maintenance", "commercial hvac", "security systems", "pest control",
                    "property management", "commercial roofing", "office furniture"], []),
    ("insurance", ["commercial insurance", "business insurance", "workers comp",
                   "liability insurance", "insurance"], []),
]

# Target industries a B2B request can name: (request words, words a segment
# must contain, nearest-match fallbacks when the catalog has none).
INDUSTRY_TERMS = [
    (["law", "legal", "attorney", "attorneys", "lawyer", "lawyers", "law firm", "law firms"],
     ["law"], []),
    (["medical", "healthcare", "health care", "doctor", "doctors", "dental", "dentist",
      "clinic", "clinics", "physician", "physicians", "hospital", "hospitals"],
     ["healthcare", "medical"], []),
    (["insurance agency", "insurance agencies", "insurance agents", "insurance offices",
      "insurance"], ["insurance agents", "insurance"], []),
    (["financial", "accounting", "accountant", "accountants", "cpa", "bank", "banks"],
     ["financial", "finance", "accounting"],
     ["DEMO Career White Collar Worker", "FIN Business Owner Small"]),
    (["construction", "contractor", "contractors", "builders", "trades"],
     ["construction", "contractor"], []),
    (["real estate", "realtor", "realtors", "property", "properties"],
     ["real estate", "property"], []),
    (["manufacturing", "manufacturer", "manufacturers", "factory", "plant"],
     ["manufacturing"], []),
    (["agriculture", "farm", "farms", "farming", "agricultural"], ["agriculture"], []),
    (["restaurant", "restaurants", "hospitality", "hotel", "hotels"],
     ["hospitality", "food service"], []),
    (["architect", "architects", "architecture"], ["architects"], []),
    (["engineering", "engineer", "engineers"], ["engineer", "engineering"], []),
    (["government", "municipal", "public sector"], ["government", "govt"], []),
    (["school", "schools", "education", "teachers"], ["teachers", "educators"], []),
    (["trucking", "logistics", "fleet", "fleets"], ["truck", "commercial vehicle"], []),
    (["oil", "gas", "energy"], ["oil and gas"], []),
]
# When an industry has no segment of its own and no specific fallback.
GENERIC_INDUSTRY_FALLBACK = ["DEMO Career White Collar Worker"]

# A request that asks for a demographic or income target, by modifier kind.
MODIFIER_REQUESTS = [
    ("income", r"\b(income|hhi|affluent|wealthy|high[- ]earning|\$?\d{2,3}k)\b",
     r"\b(income|hhi)\b"),
    ("age", r"\b(ages?|aged|years? old|\d{2}\s*(-|to|–)\s*\d{2}|\d{2}\s*\+|\d{2} plus|"
            r"seniors?|millennials?|gen z|boomers?)\b", r"\bage\b"),
    ("gender", r"\b(women|woman|female|females|men|male|males|moms?|dads?)\b",
     r"\b(age [fm]\d|female|male|women|men)"),
    ("homeowner", r"\b(homeowners?|renters?)\b", r"\b(homeowner|homeowners|renter|renters)\b"),
    ("children", r"\b(kids|children|families|parents)\b", r"\b(children|child|kids)\b"),
    ("ethnicity", r"\b(hispanic|latino|latina|spanish|african american|black|asian|"
                  r"bilingual)\b",
     r"\b(ethnicity|hispanic|spanish|african american|asian|bilingual)\b"),
    ("education", r"\b(college educated|degree|graduates?)\b", r"\beducation\b"),
]

_INCOME_CAP_RE = re.compile(r"\b(less than|under|below)\b", re.IGNORECASE)

_STOP = {"for", "and", "the", "with", "who", "that", "are", "our", "their", "small", "large",
         "local", "near", "area", "people", "customers", "clients", "client", "campaign",
         "target", "targeting", "audience", "audiences", "segments", "segment", "want",
         "wants", "looking", "reach", "offices", "office", "business", "businesses",
         "services", "service", "company", "companies", "selling", "sells", "sell", "from",
         "into", "this", "they", "them", "promote", "promoting", "drive", "family"}

# Words that don't distinguish one segment from another when comparing
# names for near-duplicates: the taxonomy prefixes and filler.
_SIMILARITY_IGNORE = {"demo", "career", "fin", "hh", "hlth", "auto", "retail", "lifestyle",
                      "lifestage", "ent", "sports", "sport", "food", "travel", "afirst",
                      "custom", "pol", "and", "or", "the", "of", "in", "to", "a", "b2b",
                      "general", "only", "expanded", "all"}


def _low_words(text):
    return " " + re.sub(r"[^a-z0-9$+]+", " ", (text or "").lower()).strip() + " "


def _has(words, phrase):
    return f" {phrase} " in words


def _first_position(words, phrases):
    hits = [words.find(f" {p} ") for p in phrases if _has(words, p)]
    return min(hits) if hits else None


def _tags(cell):
    if isinstance(cell, (list, tuple)):
        return list(cell)
    return [t.strip() for t in str(cell or "").split(",") if t.strip()]


def _rows(catalog):
    """Visible catalog records, whether handed a DataFrame or a list."""
    records = catalog.to_dict("records") if hasattr(catalog, "to_dict") else list(catalog)
    out = []
    for r in records:
        if r.get("visible") is False:
            continue
        if r.get("status") not in (None, "standard", "custom"):
            continue
        r = dict(r)
        r["tags"] = _tags(r.get("tags"))
        r["score"] = float(r.get("score") or 0.0)
        r["_words"] = _low_words(r.get("segment", ""))
        out.append(r)
    return out


def is_b2b_request(text):
    words = _low_words(text)
    if any(_has(words, p) for p in B2B_SYNONYMS):
        return True
    return any(re.search(rf"\b{a}\b", text or "") for a in B2B_ACRONYMS)


def request_tags(text, vertical_key=None):
    """Vertical tags a request points at, most specific first. B2B leads
    whenever the request reads B2B -- a VoIP seller pitching law offices is
    a B2B campaign, the offices are its target industries. Otherwise the
    tag the request mentions first leads, and a detected vertical
    (`vertical_key`) only joins after the request's own words."""
    words = _low_words(text)
    found = []
    for i, (tag, phrases) in enumerate(TAG_WORDS.items()):
        position = _first_position(words, phrases)
        if position is not None:
            found.append((position, i, tag))
    tags = [tag for _p, _i, tag in sorted(found)]
    if is_b2b_request(text):
        tags = ["B2B"] + [t for t in tags if t != "B2B"]
    mapped = TAG_FOR_VERTICAL.get(vertical_key or "")
    if mapped and mapped not in tags:
        tags.append(mapped)
    return tags


def requested_modifier_kinds(text):
    low = (text or "").lower()
    return [kind for kind, ask, _match in MODIFIER_REQUESTS if re.search(ask, low)]


def _query_terms(text):
    return [w for w in _low_words(text).split() if len(w) >= 4 and w not in _STOP]


def _relevance(row, terms):
    return sum(1 for t in terms if _has(row["_words"], t))


# ---------------------------------------------------------------------------
# Near-duplicates
# ---------------------------------------------------------------------------

def _name_tokens(segment):
    tokens = set()
    for word in _low_words(segment).split():
        if word in _SIMILARITY_IGNORE:
            continue
        if len(word) > 3 and word.endswith("s"):
            word = word[:-1]
        tokens.add(word)
    return tokens


def near_duplicates(a, b):
    """Two segment names that say the same thing: one's distinguishing
    words contain the other's (two words at least, so a lone "home" isn't
    a match for everything), or they share more than half of them."""
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return False
    small, large = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    if len(small) >= 2 and small <= large:
        return True
    return len(ta & tb) / len(ta | tb) > 0.5


def collapse(rows):
    """Rows grouped into near-duplicate clusters, each led by its
    highest-scoring member: [(lead row, [similar rows...]), ...] in score
    order. A row joins the first cluster any of whose members it
    duplicates."""
    clusters = []
    for row in sorted(rows, key=lambda r: (-r["score"], r["segment"])):
        for lead, members in clusters:
            if any(near_duplicates(row["segment"], m["segment"]) for m in [lead] + members):
                members.append(row)
                break
        else:
            clusters.append((row, []))
    return clusters


# ---------------------------------------------------------------------------
# Pairings and records
# ---------------------------------------------------------------------------

def _pairing(row):
    """(modifier, line count) to show beside a pick, or (None, 0): the most
    common pairing that shares 3+ lines, skipping an income cap beside a
    B2B segment (a business owner "Less than 100K" reads as a contradiction,
    and it came from a consumer stack that happened to include one)."""
    b2b = "B2B" in row.get("tags", []) or row.get("b2b_group")
    for pairing in row.get("pairings") or []:
        name, count = (pairing[0], int(pairing[1])) if isinstance(pairing, (list, tuple)) \
            else (str(pairing), 0)
        if count < MIN_PAIRING_LINES:
            continue
        if b2b and _INCOME_CAP_RE.search(name):
            continue
        return name, count
    return None, 0


def _rec(row, slot, reason=None, similar=(), nearest=False):
    pairing, count = _pairing(row)
    return {
        "segment": row["segment"], "slot": slot, "role": row.get("role") or "core",
        "status": row.get("status") or "standard", "tags": row["tags"],
        "b2b_group": row.get("b2b_group"), "score": round(row["score"], 3),
        "line_count": int(row.get("line_count") or 0), "category": row.get("category") or "",
        "rfp_selectable": bool(row.get("rfp_selectable")),
        "pairing": pairing, "pairing_count": count, "reason": reason,
        "similar": [s["segment"] for s in similar], "nearest_match": nearest,
    }


def pairing_note(rec):
    """The optional-add line shown under a recommendation."""
    if not rec.get("pairing"):
        return ""
    return f"often paired with {rec['pairing']}"


# ---------------------------------------------------------------------------
# Recommendation
# ---------------------------------------------------------------------------

def _b2b_core(rows, text, limit):
    words = _low_words(text)
    terms = _query_terms(text)
    b2b = [r for r in rows if "B2B" in r["tags"] and r.get("status") == "standard"
           and r.get("role") != "modifier"]
    clusters = collapse(b2b)
    leads = [lead for lead, _m in clusters]
    similar_of = {lead["segment"]: members for lead, members in clusters}
    collapsed_away = {m["segment"] for _lead, members in clusters for m in members}
    picked, used = [], set()

    def take(candidates, slot, reason, nearest=False):
        for r in candidates:
            if r["segment"] not in used:
                used.add(r["segment"])
                picked.append(_rec(r, slot, reason, similar_of.get(r["segment"], ()), nearest))
                return True
        return False

    # 1. Business owner / size.
    size = sorted((r for r in leads if r.get("b2b_group") == "size_owner"),
                  key=lambda r: -r["score"])
    take(size, "Business owner / size", "business-owner segment")

    # 2. Decision maker for the product being sold.
    function = None
    for name, phrases, acronyms in FUNCTION_TERMS:
        if any(_has(words, p) for p in phrases) or any(
                re.search(rf"\b{a}\b", text or "") for a in acronyms):
            function = name
            break
    makers = [r for r in leads if r.get("b2b_group") == "decision_maker"]
    matching = sorted((r for r in makers if function and r.get("b2b_function") == function),
                      key=lambda r: -r["score"])
    fallback = sorted([r for r in makers if r.get("b2b_function") == "executive"] or makers,
                      key=lambda r: -r["score"])
    if matching:
        label = function.upper() if function in ("it", "hr") else function
        take(matching, "Decision maker", f"{label} decision makers")
    else:
        take(fallback, "Decision maker", "business decision makers")

    # 3. One industry/occupation segment per target industry named, or the
    # nearest available segment when the catalog has none.
    by_name = {r["segment"]: r for r in rows if r.get("status") == "standard"}
    for asks, segment_words, fallbacks in INDUSTRY_TERMS:
        if not any(_has(words, a) for a in asks):
            continue
        candidates = sorted((r for r in leads if r.get("b2b_group") in ("industry", "occupation")
                             and any(_has(r["_words"], w) for w in segment_words)),
                            key=lambda r: -r["score"])
        if take(candidates, f"Target industry: {asks[0]}", f"named industry: {asks[0]}"):
            continue
        nearest = [by_name[n] for n in fallbacks + GENERIC_INDUSTRY_FALLBACK
                   if n in by_name and n not in collapsed_away]
        take(nearest, f"Target industry: {asks[0]} (nearest match)",
             f"no {asks[0]} segment in the catalog; nearest available", nearest=True)

    # The rest of the tag -- only picks that match the request's own words.
    rest = [r for r in leads if r["segment"] not in used and _relevance(r, terms) > 0]
    rest.sort(key=lambda r: (-_relevance(r, terms), -r["score"]))
    for r in rest:
        if len(picked) >= limit:
            break
        take([r], "Also in B2B", "matches the request")
    return picked[:limit]


def _in_tag_core(rows, tag, terms, limit):
    in_tag = [r for r in rows if tag in r["tags"] and r.get("status") == "standard"
              and r.get("role") != "modifier"]
    clusters = collapse(in_tag)
    clusters.sort(key=lambda c: (-_relevance(c[0], terms), -c[0]["score"], c[0]["segment"]))
    return [_rec(lead, tag, None, members) for lead, members in clusters[:limit]]


def _modifiers_for(rows, kinds, used, per_kind=1):
    out = []
    mods = [r for r in rows if r.get("role") == "modifier" and r.get("status") == "standard"]
    for kind, _ask, match in MODIFIER_REQUESTS:
        if kind not in kinds:
            continue
        hits = sorted((r for r in mods if re.search(match, r["_words"]) and r["segment"] not in used),
                      key=lambda r: -r["score"])[:per_kind]
        for r in hits:
            used.add(r["segment"])
            out.append(_rec(r, "Requested demographic", f"the request asks for a {kind} target"))
    return out


def recommend(catalog, text, vertical_key=None, limit=MAX_PICKS, custom_limit=3):
    """{"tags", "primary", "core", "modifiers", "custom"} for a request.
    `core` is ranked within the first matched tag, at most `limit` picks;
    with no tag at all it's empty -- ranking across every segment is
    exactly what this replaces."""
    rows = _rows(catalog)
    tags = request_tags(text, vertical_key)
    terms = _query_terms(text)
    result = {"tags": tags, "primary": tags[0] if tags else None, "core": [],
              "modifiers": [], "custom": []}
    if not tags:
        return result
    primary = tags[0]
    if primary == "B2B":
        core = _b2b_core(rows, text, limit)
    else:
        core = _in_tag_core(rows, primary, terms, limit)
    used = {r["segment"] for r in core}
    result["core"] = core
    result["modifiers"] = _modifiers_for(rows, requested_modifier_kinds(text), used)
    customs = [r for r in rows if r.get("status") == "custom" and primary in r["tags"]]
    customs.sort(key=lambda r: (-_relevance(r, terms), -r["score"], r["segment"]))
    result["custom"] = [_rec(r, "Proven custom recipe") for r in customs[:custom_limit]]
    return result


def flat_list(result):
    """Core first, then any requested modifiers, then custom recipes --
    the order a reader should meet them in."""
    return result["core"] + result["modifiers"] + result["custom"]
