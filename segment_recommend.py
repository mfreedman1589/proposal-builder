"""segment_recommend.py -- deterministic audience recommendations from the
ranked, tagged catalog (segment_catalog.py, overlaid by audience_catalog).

Pure: takes catalog records, returns plain dicts. No Streamlit, no DB, no
model call -- the Audience finder shows these directly, and draft-from-notes
and the finder's Suggest mode hand them to Claude as the starting point.

The rules (2026-10-20):
- Recommend CORE segments ranked by score WITHIN the matched vertical tag,
  never across the whole catalog.
- Offer each one's most common modifier pairing as an optional add ("often
  paired with HH Income 100K Plus").
- A modifier is never recommended on its own unless the request asks for a
  demographic or income target -- and even then it comes after the core
  segments.
- B2B is slotted: one business-owner/size segment, one decision-maker
  segment matching the product (IT for VoIP/IT, facilities for building
  services, HR for staffing/benefits, insurance for commercial insurance),
  and one industry or occupation segment per target industry the request
  names. The rest of the B2B tag fills in by score.
- Proven custom recipes for the vertical are listed separately, labeled.
"""
import re

TAG_FOR_VERTICAL = {
    "healthcare": "Healthcare", "education": "Education", "auto": "Auto",
    "retail": "Retail", "legal": "Legal", "home_improvement": "Home Services",
    "banking": "Financial",
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

# Words that name a vertical tag directly ("show me healthcare segments").
TAG_WORDS = {
    "Healthcare": ["healthcare", "health care", "medical", "health"],
    "Education": ["education", "school", "college", "university"],
    "Auto": ["automotive", "dealership", "car dealer"],
    "Retail": ["retail", "retailer", "store"],
    "Legal": ["legal", "lawyer", "attorney", "law firm", "personal injury"],
    "Home Services": ["home services", "home improvement", "hvac", "roofing", "plumbing",
                      "contractor"],
    "Financial": ["financial", "bank", "banking", "credit union", "finance"],
    "B2B": ["b2b"],
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

# Target industries a B2B request can name -> words a segment must contain.
INDUSTRY_TERMS = [
    (["law", "legal", "attorney", "attorneys", "lawyer", "lawyers", "law firm", "law firms"],
     ["law"]),
    (["medical", "healthcare", "health care", "doctor", "doctors", "dental", "dentist",
      "clinic", "clinics", "physician", "physicians", "hospital", "hospitals"],
     ["healthcare", "medical"]),
    (["insurance agency", "insurance agencies", "insurance agents", "insurance offices",
      "insurance"], ["insurance agents", "insurance"]),
    (["financial", "accounting", "accountant", "accountants", "cpa", "bank", "banks"],
     ["financial", "finance", "accounting"]),
    (["construction", "contractor", "contractors", "builders", "trades"],
     ["construction", "contractor"]),
    (["real estate", "realtor", "realtors", "property", "properties"],
     ["real estate", "property"]),
    (["manufacturing", "manufacturer", "manufacturers", "factory", "plant"],
     ["manufacturing"]),
    (["agriculture", "farm", "farms", "farming", "agricultural"], ["agriculture"]),
    (["restaurant", "restaurants", "hospitality", "hotel", "hotels"],
     ["hospitality", "food service"]),
    (["architect", "architects", "architecture"], ["architects"]),
    (["engineering", "engineer", "engineers"], ["engineer", "engineering"]),
    (["government", "municipal", "public sector"], ["government", "govt"]),
    (["school", "schools", "education", "teachers"], ["teachers", "educators"]),
    (["trucking", "logistics", "fleet", "fleets"], ["truck", "commercial vehicle"]),
    (["oil", "gas", "energy"], ["oil and gas"]),
]

# A request that asks for a demographic or income target, by modifier kind.
MODIFIER_REQUESTS = [
    ("income", r"\b(income|hhi|affluent|wealthy|high[- ]earning|\$?\d{2,3}k)\b",
     r"\b(income|hhi)\b"),
    ("age", r"\b(ages?|aged|years? old|\d{2}\s*(-|to|–)\s*\d{2}|\d{2}\s*\+|\d{2} plus|"
            r"seniors?|millennials?|gen z|boomers?)\b", r"\bage\b"),
    ("gender", r"\b(women|woman|female|females|men|male|males|moms?|dads?)\b",
     r"\b(age [fm]\d|female|male|women|men)"),
    ("homeowner", r"\b(homeowners?|renters?)\b", r"\b(homeowner|homeowners|renter|renters)\b"),
    ("children", r"\b(kids|children|families|family|parents)\b", r"\b(children|child|kids)\b"),
    ("ethnicity", r"\b(hispanic|latino|latina|spanish|african american|black|asian|"
                  r"bilingual)\b",
     r"\b(ethnicity|hispanic|spanish|african american|asian|bilingual)\b"),
    ("education", r"\b(college educated|degree|graduates?)\b", r"\beducation\b"),
]

_STOP = {"for", "and", "the", "with", "who", "that", "are", "our", "their", "small", "large",
         "local", "near", "area", "people", "customers", "clients", "client", "campaign",
         "target", "targeting", "audience", "audiences", "segments", "segment", "want",
         "wants", "looking", "reach", "offices", "office", "business", "businesses",
         "services", "service", "company", "companies", "selling", "sells", "sell"}


def _low_words(text):
    return " " + re.sub(r"[^a-z0-9$+]+", " ", (text or "").lower()).strip() + " "


def _has(words, phrase):
    return f" {phrase} " in words


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


def _tags(cell):
    if isinstance(cell, (list, tuple)):
        return list(cell)
    return [t.strip() for t in str(cell or "").split(",") if t.strip()]


def is_b2b_request(text):
    words = _low_words(text)
    if any(_has(words, p) for p in B2B_SYNONYMS):
        return True
    return any(re.search(rf"\b{a}\b", text or "") for a in B2B_ACRONYMS)


def request_tags(text, vertical_key=None):
    """Vertical tags a request points at, most specific first. B2B leads
    whenever the request reads B2B -- a VoIP seller pitching law offices is
    a B2B campaign, the offices are its target industries."""
    words = _low_words(text)
    tags = []
    if is_b2b_request(text):
        tags.append("B2B")
    for tag, phrases in TAG_WORDS.items():
        if tag not in tags and any(_has(words, p) for p in phrases):
            tags.append(tag)
    mapped = TAG_FOR_VERTICAL.get(vertical_key or "")
    if mapped and mapped not in tags:
        tags.insert(1 if tags[:1] == ["B2B"] else 0, mapped)
    return tags


def requested_modifier_kinds(text):
    low = (text or "").lower()
    return [kind for kind, ask, _match in MODIFIER_REQUESTS if re.search(ask, low)]


def _query_terms(text):
    return [w for w in _low_words(text).split() if len(w) >= 4 and w not in _STOP]


def _relevance(row, terms):
    return sum(1 for t in terms if _has(row["_words"], t))


def _pairing(row):
    pairings = row.get("pairings") or []
    if not pairings:
        return None, 0
    first = pairings[0]
    return (first[0], int(first[1])) if isinstance(first, (list, tuple)) else (str(first), 0)


def _rec(row, slot, reason=None):
    pairing, count = _pairing(row)
    return {
        "segment": row["segment"], "slot": slot, "role": row.get("role") or "core",
        "status": row.get("status") or "standard", "tags": row["tags"],
        "b2b_group": row.get("b2b_group"), "score": round(row["score"], 3),
        "line_count": int(row.get("line_count") or 0), "category": row.get("category") or "",
        "rfp_selectable": bool(row.get("rfp_selectable")),
        "pairing": pairing, "pairing_count": count, "reason": reason,
    }


def pairing_note(rec):
    """The optional-add line shown under a recommendation."""
    if not rec.get("pairing"):
        return ""
    return f"often paired with {rec['pairing']}"


def _ranked(rows, terms=()):
    return sorted(rows, key=lambda r: (-_relevance(r, terms), -r["score"], r["segment"]))


def _b2b_core(rows, text, limit):
    words = _low_words(text)
    b2b = [r for r in rows if "B2B" in r["tags"] and r.get("status") == "standard"
           and r.get("role") != "modifier"]
    picked, used = [], set()

    def take(candidates, slot, reason):
        for r in candidates:
            if r["segment"] not in used:
                used.add(r["segment"])
                picked.append(_rec(r, slot, reason))
                return True
        return False

    # 1. Business owner / size -- "small" in the request prefers a small-
    # business segment.
    size = [r for r in b2b if r.get("b2b_group") == "size_owner"]
    wants_small = any(_has(words, w) for w in ("small", "smb", "smbs", "startup", "startups"))
    size.sort(key=lambda r: (-(wants_small and _has(r["_words"], "small")), -r["score"]))
    take(size, "Business owner / size", "business-owner segment")

    # 2. Decision maker for the product being sold.
    function = None
    for name, phrases, acronyms in FUNCTION_TERMS:
        if any(_has(words, p) for p in phrases) or any(
                re.search(rf"\b{a}\b", text or "") for a in acronyms):
            function = name
            break
    makers = [r for r in b2b if r.get("b2b_group") == "decision_maker"]
    matching = [r for r in makers if function and r.get("b2b_function") == function]
    fallback = [r for r in makers if r.get("b2b_function") == "executive"] or makers
    take(sorted(matching, key=lambda r: -r["score"]) or sorted(fallback, key=lambda r: -r["score"]),
         "Decision maker", f"{function.upper() if function in ('it', 'hr') else function} "
                           f"decision makers" if matching else "business decision makers")

    # 3. One industry/occupation segment per target industry named.
    for asks, segment_words in INDUSTRY_TERMS:
        if not any(_has(words, a) for a in asks):
            continue
        candidates = [r for r in b2b if r.get("b2b_group") in ("industry", "occupation")
                      and any(_has(r["_words"], w) for w in segment_words)]
        take(sorted(candidates, key=lambda r: -r["score"]), "Target industry",
             f"named industry: {asks[0]}")

    # The rest of the tag by score -- owners, decision makers and industries
    # before occupations the request didn't name.
    rest = [r for r in b2b if r["segment"] not in used]
    rest.sort(key=lambda r: (r.get("b2b_group") == "occupation", -r["score"]))
    for r in rest:
        if len(picked) >= limit:
            break
        used.add(r["segment"])
        picked.append(_rec(r, "Also in B2B", None))
    return picked[:limit]


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


def recommend(catalog, text, vertical_key=None, limit=10, custom_limit=3):
    """{"tags", "core", "modifiers", "custom"} for a request. `core` is
    ranked within the first matched tag; with no tag at all it's empty --
    ranking across every segment is exactly what this replaces."""
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
        in_tag = [r for r in rows if primary in r["tags"] and r.get("status") == "standard"
                  and r.get("role") != "modifier"]
        core = [_rec(r, primary, None) for r in _ranked(in_tag, terms)[:limit]]
    used = {r["segment"] for r in core}
    result["core"] = core
    result["modifiers"] = _modifiers_for(rows, requested_modifier_kinds(text), used)
    customs = [r for r in rows if r.get("status") == "custom" and primary in r["tags"]]
    result["custom"] = [_rec(r, "Proven custom recipe", None)
                        for r in _ranked(customs, terms)[:custom_limit]]
    return result


def flat_list(result):
    """Core first, then any requested modifiers, then custom recipes --
    the order a reader should meet them in."""
    return result["core"] + result["modifiers"] + result["custom"]
