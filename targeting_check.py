"""Post-draft check: does a plan line's Targeting description match the
segments it actually buys?

Two checks, both pure (no DB, no Streamlit):

1. `uncovered_attribute_items` -- an attribute the line's Targeting text
   names (homeowner, an income threshold, an age range, home age, children,
   gender) that no segment on the line covers, a CUSTOM segment's packaged
   criteria included. The client reads the Targeting column as a promise;
   "Homeowners, adults 35+" over a line that buys only an age segment
   promises something the plan doesn't target.
2. `threshold_mismatch_items` -- a threshold the notes state (income, age,
   home built before X) that a same-kind segment on the line contradicts
   (notes "$150K+", segment "$200K Plus"), with the closest catalog segment
   offered as the swap.

Both return structured review items rather than prose, in the shape the
interactive draft follow-ups work will consume (BACKLOG.md): an id, a type,
a priority, a short question, options each carrying a patch, and a default
(what happens if the rep skips it -- always the draft as it stands). Nothing
here applies a patch; a patch only describes the state change.

Attributes are read from segment NAMES by rule (the catalog's own naming
grammar -- "HH Income 150K Plus", "DEMO Age F35-64", "HH Home Built Before
2000"), never by a model call.
"""

import hashlib
import re

# A line ANDs at most this many segments -- targeting_groups.MAX_AND_SEGMENTS,
# repeated rather than imported so this module stays dependency-free; the
# test asserts the two agree.
MAX_AND_SEGMENTS = 2

ATTRIBUTE_LABELS = {
    "homeowner": "homeowners",
    "income": "household income",
    "age": "age",
    "home_built": "home age",
    "children": "children in the household",
    "gender": "gender",
}

_GENDER_WORDS = {"F": "women", "M": "men"}


def _money(value):
    """'150' + 'k' / '150,000' -> 150000."""
    text = value.replace(",", "").replace("$", "").strip().lower()
    mult = 1
    if text.endswith("k"):
        text, mult = text[:-1], 1000
    elif text.endswith("m"):
        text, mult = text[:-1], 1_000_000
    try:
        number = float(text) * mult
    except ValueError:
        return None
    # "150" next to "income" with no unit is thousands in this trade.
    if number < 1000:
        number *= 1000
    return int(number)


def _attr(kind, **values):
    return {"kind": kind, **values}


# --- reading a segment name ------------------------------------------------

_SEG_INCOME = re.compile(
    r"income\s+(?:(less than|under)\s+)?\$?(\d[\d,]*\s*k?)\s*(plus|\+)?", re.I)
_SEG_AGE = re.compile(r"\bage\s+([AFMW])(\d{2})(?:\s*-\s*(\d{2})|\s*(?:plus|\+))", re.I)
_SEG_BUILT = re.compile(r"built\s+before\s+(\d{4})", re.I)


def segment_attributes(name):
    """Every attribute a catalog segment name covers. A CUSTOM segment is
    read the same way, so each criterion it packages counts."""
    text = str(name or "")
    found = []
    if re.search(r"home\s*owner", text, re.I):
        found.append(_attr("homeowner"))
    for m in _SEG_INCOME.finditer(text):
        amount = _money(m.group(2))
        if amount is None:
            continue
        if m.group(1):
            found.append(_attr("income", min=None, max=amount))
        else:
            found.append(_attr("income", min=amount, max=None))
    for m in _SEG_AGE.finditer(text):
        sex = m.group(1).upper()
        low = int(m.group(2))
        high = int(m.group(3)) if m.group(3) else None
        found.append(_attr("age", min=low, max=high))
        if sex in ("F", "W"):
            found.append(_attr("gender", value="F"))
        elif sex == "M":
            found.append(_attr("gender", value="M"))
    for m in _SEG_BUILT.finditer(text):
        found.append(_attr("home_built", before=int(m.group(1))))
    if (re.search(r"\bchild(ren)?\b|\bkids?\b|\bparents?\b|\bmoms?\b|\bdads?\b|\bfamil(y|ies)\b",
                  text, re.I)
            and not re.search(r"\bno children\b|\bcharity\b", text, re.I)):
        found.append(_attr("children"))
    if re.search(r"\b(women|female|moms?)\b", text, re.I):
        found.append(_attr("gender", value="F"))
    if re.search(r"\b(men|male|dads?)\b", text, re.I):
        found.append(_attr("gender", value="M"))
    return found


# --- reading prose (a Targeting description, or the notes) -----------------

_TXT_MONEY = r"\$?\s?(\d{2,3}(?:,000)?\s?[kK]?|\d{1,3}(?:,\d{3})+)"
_TXT_INCOME_PATTERNS = [
    # "$150K+ household income", "$150K+ HHI", "$150,000 or more in income"
    (re.compile(_TXT_MONEY + r"\s*(\+|plus|or more|and up|and above)?\s*"
                r"(?:in\s+)?(?:annual\s+)?(?:household\s+income|hh\s*income|hhi|income)\b", re.I),
     "after"),
    # "household income $150K+", "income over $100K", "HHI under $75K", "earning $150K+"
    (re.compile(r"(?:household\s+income|hh\s*income|hhi|income|earning|earn|making)\s*"
                r"(?:of\s+)?(over|above|at least|more than|greater than|under|below|less than|"
                r"up to)?\s*(?:of\s+)?" + _TXT_MONEY + r"\s*(\+|plus|or more|and up|and above)?",
                re.I),
     "before"),
]
_UPPER_WORDS = {"under", "below", "less than", "up to"}

_TXT_AGE_PATTERNS = [
    # "adults 35+", "ages 25-54", "aged 35 to 64", "women 25-54", "A35+"
    re.compile(r"\b(?:adults?|ages?|aged|women|men|females?|males?|people|consumers|"
               r"homeowners|parents|moms|dads)\s+(\d{2})\s*(?:(\+|plus|and older|and over|or older)"
               r"|\s*(?:-|–|to)\s*(\d{2}))", re.I),
    re.compile(r"\b[AFMW](\d{2})\s*(?:(\+)|-(\d{2}))(?=\W|$)"),
    # "35-64 year olds", "35+ year olds"
    re.compile(r"\b(\d{2})\s*(?:(\+)|\s*(?:-|–|to)\s*(\d{2}))\s*(?:year[- ]olds?|years old|yo)\b",
               re.I),
    # "over 55", "55 and older"
    re.compile(r"\b(?:adults?|people|consumers)\s+(?:over|aged over)\s+(\d{2})()()", re.I),
]

_TXT_BUILT = re.compile(r"\b(?:built|constructed)\s+(?:before|prior to|pre-?)\s*(\d{4})\b"
                        r"|\bpre-?(\d{4})\s+homes?\b|\bolder than\s+(\d{4})\b", re.I)


def _income_from_text(text):
    found = []
    for pattern, _where in _TXT_INCOME_PATTERNS:
        for m in pattern.finditer(text):
            groups = m.groups()
            if pattern is _TXT_INCOME_PATTERNS[0][0]:
                amount_text, word = groups[0], None
            else:
                word, amount_text = groups[0], groups[1]
            amount = _money(amount_text)
            if amount is None or amount < 10000:
                continue
            upper = bool(word) and word.lower() in _UPPER_WORDS
            attr = (_attr("income", min=None, max=amount) if upper
                    else _attr("income", min=amount, max=None))
            attr["phrase"] = m.group(0).strip()
            found.append(attr)
    return _dedupe(found)


def _age_from_text(text):
    found = []
    for pattern in _TXT_AGE_PATTERNS:
        for m in pattern.finditer(text):
            low = int(m.group(1))
            high = int(m.group(3)) if m.group(3) else None
            if not 13 <= low <= 99 or (high is not None and not low < high <= 99):
                continue
            attr = _attr("age", min=low, max=high)
            attr["phrase"] = m.group(0).strip()
            found.append(attr)
    return _dedupe(found)


def _dedupe(attrs):
    seen, out = set(), []
    for a in attrs:
        key = tuple(sorted((k, v) for k, v in a.items() if k != "phrase"))
        if key not in seen:
            seen.add(key)
            out.append(a)
    return out


def text_attributes(text):
    """Every attribute a prose description names, each with the `phrase`
    that named it (what a "remove it from the description" fix removes)."""
    text = str(text or "")
    found = []
    m = re.search(r"\bhome\s*-?\s*owners?(?:hip)?\b|\bowner[- ]occupied\b", text, re.I)
    if m:
        found.append({**_attr("homeowner"), "phrase": m.group(0)})
    found += _income_from_text(text)
    found += _age_from_text(text)
    for m in _TXT_BUILT.finditer(text):
        year = next(int(g) for g in m.groups() if g)
        found.append({**_attr("home_built", before=year), "phrase": m.group(0)})
    m = re.search(r"\b(?:families|family|parents|moms|dads|children|kids)\b", text, re.I)
    if m and not re.search(r"\bno (?:children|kids)\b|\bfamily[- ]owned\b|\bfamily dentistry\b",
                           text, re.I):
        found.append({**_attr("children"), "phrase": m.group(0)})
    m = re.search(r"\b(women|female|moms)\b", text, re.I)
    if m:
        found.append({**_attr("gender", value="F"), "phrase": m.group(0)})
    m = re.search(r"\b(men|male|dads)\b", text, re.I)
    if m:
        found.append({**_attr("gender", value="M"), "phrase": m.group(0)})
    return found


# --- matching --------------------------------------------------------------

def _covers(segment_attrs, attr):
    """Does any of a segment's attributes cover this text attribute? Kind is
    what coverage means here; a threshold that disagrees is check 2's job."""
    for s in segment_attrs:
        if s["kind"] != attr["kind"]:
            continue
        if attr["kind"] == "gender" and s.get("value") != attr.get("value"):
            continue
        return True
    return False


def _same_threshold(a, b):
    if a["kind"] != b["kind"]:
        return False
    if a["kind"] == "home_built":
        return a["before"] == b["before"]
    return a.get("min") == b.get("min") and a.get("max") == b.get("max")


def _same_direction(a, b):
    if a["kind"] == "income":
        return (a.get("min") is None) == (b.get("min") is None)
    return True


def _describe(attr):
    kind = attr["kind"]
    if kind == "income":
        if attr.get("max") is not None:
            return f"household income under ${attr['max'] // 1000:,}K"
        return f"household income ${attr['min'] // 1000:,}K+"
    if kind == "age":
        return (f"ages {attr['min']}-{attr['max']}" if attr.get("max") is not None
                else f"ages {attr['min']}+")
    if kind == "home_built":
        return f"homes built before {attr['before']}"
    if kind == "gender":
        return _GENDER_WORDS.get(attr.get("value"), "gender")
    return ATTRIBUTE_LABELS[kind]


def _distance(want, have):
    if want["kind"] == "home_built":
        return abs(want["before"] - have["before"])
    if want["kind"] == "income":
        side = "max" if want.get("max") is not None else "min"
        return abs((want.get(side) or 0) - (have.get(side) or 0))
    return (abs((want.get("min") or 0) - (have.get("min") or 0))
            + abs((want.get("max") or 100) - (have.get("max") or 100)))


def _catalog_rows(catalog):
    """[(name, usage)] from a DataFrame, a list of dicts or a list of names."""
    if catalog is None:
        return []
    if hasattr(catalog, "to_dict"):
        catalog = catalog.to_dict("records")
    rows = []
    for row in catalog:
        if isinstance(row, str):
            rows.append((row, 0))
            continue
        if row.get("visible") is False or row.get("hidden"):
            continue
        name = str(row.get("segment") or "").strip()
        if not name:
            continue
        usage = row.get("score") or row.get("usage") or row.get("impressions") or 0
        try:
            usage = float(usage)
        except (TypeError, ValueError):
            usage = 0.0
        rows.append((name, usage))
    return rows


def _best_segment_for(attr, catalog_rows, exclude=(), also_cover=()):
    """The catalog segment closest to `attr`: same kind (and direction), an
    exact threshold first, then the nearest one; among equals, one that also
    covers `also_cover` (a custom's other criteria) wins, then usage."""
    best = None
    for name, usage in catalog_rows:
        if name in exclude:
            continue
        attrs = segment_attributes(name)
        for have in attrs:
            if have["kind"] != attr["kind"]:
                continue
            if attr["kind"] == "gender" and have.get("value") != attr.get("value"):
                continue
            if attr["kind"] in ("income", "age", "home_built"):
                if not ("min" in attr or "before" in attr):
                    # A bare mention ("household income") -- any threshold.
                    distance = 0
                elif not _same_direction(attr, have):
                    continue
                else:
                    distance = _distance(attr, have)
            else:
                distance = 0
            covers_rest = sum(1 for other in also_cover if _covers(attrs, other))
            # Fewer criteria is better when nothing else is asked for: a plain
            # "DEMO Homeowner" beats a custom that also packages income.
            extra = len(attrs) - 1 - covers_rest
            key = (distance, -covers_rest, extra, -usage, name)
            if best is None or key < best[0]:
                best = (key, name)
    return best[1] if best else None


def _item_id(*parts):
    return "tgt-" + hashlib.sha1("|".join(str(p) for p in parts).encode()).hexdigest()[:10]


def _line_name(line):
    return line.get("tactic") or line.get("label") or line.get("product") or "plan"


def uncovered_attribute_items(lines, catalog, option_name=""):
    """Check 1. `lines`: dicts carrying `segments` (the line's kept
    segments), `audience_track` (its Targeting text) and `tactic`/`label`
    (how the rep finds it on the grid). A line with no `segments` key at
    all isn't checked -- a retargeting/sports line, or a draft from before
    lines named their segments."""
    rows = _catalog_rows(catalog)
    items = []
    for line in lines or []:
        if "segments" not in line:
            continue
        segments = [s for s in (line.get("segments") or []) if str(s).strip()]
        text = str(line.get("audience_track") or "")
        if not text.strip():
            continue
        seg_attrs = [a for s in segments for a in segment_attributes(s)]
        for attr in text_attributes(text):
            if _covers(seg_attrs, attr):
                continue
            phrase = attr["phrase"]
            label = _describe(attr) if attr["kind"] != "children" else ATTRIBUTE_LABELS["children"]
            tactic = _line_name(line)
            options = []
            room = len(segments) < MAX_AND_SEGMENTS
            suggestion = _best_segment_for(attr, rows, exclude=segments) if room else None
            if suggestion:
                options.append({
                    "id": "add_segment",
                    "label": f"Add {suggestion}",
                    "patch": {"op": "add_segment", "option": option_name, "tactic": tactic,
                              "segment": suggestion},
                })
            options.append({
                "id": "remove_phrase",
                "label": f'Remove "{phrase}" from the description',
                "patch": {"op": "remove_phrase", "option": option_name, "tactic": tactic,
                          "field": "Targeting", "phrase": phrase},
            })
            options.append({"id": "keep", "label": "Leave it as drafted", "patch": None})
            why = ("" if room or suggestion else
                   f" The line already combines {MAX_AND_SEGMENTS} segments, so there's no room "
                   f"to add one.")
            items.append({
                "id": _item_id("uncovered", option_name, tactic, attr["kind"], phrase),
                "check": "targeting_text_uncovered",
                "type": "choice",
                "priority": "decision",
                "question": (f'The {tactic} line\'s Targeting says "{phrase}", but none of its '
                             f"segments target {label}.{why}"),
                "options": options,
                "default": "keep",
                "context": {"option": option_name, "tactic": tactic, "segments": segments,
                            "attribute": {k: v for k, v in attr.items() if k != "phrase"},
                            "phrase": phrase},
            })
    return items


def notes_thresholds(notes):
    """The income, age and home-age thresholds the notes state."""
    return [a for a in text_attributes(notes)
            if a["kind"] in ("income", "age", "home_built")]


def threshold_mismatch_items(lines, notes, catalog, option_name=""):
    """Check 2. For each same-kind threshold on a line's segments that no
    threshold in the notes agrees with, a swap to the closest catalog
    segment. Silent when the notes state no threshold of that kind."""
    stated = notes_thresholds(notes)
    if not stated:
        return []
    rows = _catalog_rows(catalog)
    items = []
    for line in lines or []:
        if "segments" not in line:
            continue
        segments = [s for s in (line.get("segments") or []) if str(s).strip()]
        tactic = _line_name(line)
        for segment in segments:
            seg_attrs = segment_attributes(segment)
            for have in seg_attrs:
                if have["kind"] not in ("income", "age", "home_built"):
                    continue
                wanted = [w for w in stated
                          if w["kind"] == have["kind"] and _same_direction(w, have)]
                if not wanted or any(_same_threshold(w, have) for w in wanted):
                    continue
                want = min(wanted, key=lambda w: _distance(w, have))
                # A custom's other criteria (home age beside its income) are
                # kept if the catalog has a segment that keeps them.
                others = [a for a in seg_attrs if a is not have]
                swap = _best_segment_for(want, rows, exclude=segments, also_cover=others)
                if swap and not any(_same_threshold(want, a) for a in segment_attributes(swap)):
                    swap = None
                dropped = [a for a in others if swap and not _covers(segment_attributes(swap), a)]
                # Carry a criterion the swap loses as its own segment when the
                # line still has room under the AND limit.
                replacement = [swap] if swap else []
                carried = []
                for lost in dropped:
                    extra = _best_segment_for(lost, rows, exclude=segments + replacement)
                    if (extra and any(_same_threshold(lost, a) or lost["kind"] not in
                                      ("income", "age", "home_built")
                                      for a in segment_attributes(extra))
                            and len(segments) - 1 + len(replacement) + 1 <= MAX_AND_SEGMENTS):
                        replacement.append(extra)
                        carried.append(lost)
                dropped = [a for a in dropped if a not in carried]
                options = []
                if swap:
                    options.append({
                        "id": "swap_segment",
                        "label": "Swap to " + " AND ".join(replacement),
                        "patch": {"op": "replace_segment", "option": option_name, "tactic": tactic,
                                  "from": segment, "to": replacement},
                    })
                options.append({"id": "keep", "label": f"Keep {segment}", "patch": None})
                note = ""
                if dropped:
                    note = (" The swap doesn't carry "
                            + ", ".join(_describe(a) for a in dropped)
                            + f" from {segment}.")
                elif not swap:
                    note = f" The catalog has no segment at {_describe(want)}."
                items.append({
                    "id": _item_id("threshold", option_name, tactic, segment, have["kind"]),
                    "check": "notes_threshold_mismatch",
                    "type": "suggestion" if swap else "fyi",
                    "priority": "suggestion" if swap else "fyi",
                    "question": (f"The notes ask for {_describe(want)}, but the {tactic} line "
                                 f"uses {segment} ({_describe(have)}).{note}"),
                    "options": options,
                    "default": "keep",
                    "context": {"option": option_name, "tactic": tactic, "segment": segment,
                                "notes_value": {k: v for k, v in want.items() if k != "phrase"},
                                "segment_value": have, "suggested": replacement or None},
                })
    return items


# --- criteria the notes never asked for (QA-023) ----------------------------
# A criterion only counts as asked for when the notes state it. Age is also
# asked for by a word that implies it ("seniors", "retirees", "students").
_AGE_IMPLIED = re.compile(r"\b(seniors?|retirees?|retired|older adults|elderly|boomers?|"
                          r"millennials?|gen ?z|young adults?|teens?|students?|empty[- ]nesters?)\b",
                          re.I)


def _asked_kinds(notes):
    kinds = {a["kind"] for a in text_attributes(notes)}
    if _AGE_IMPLIED.search(str(notes or "")):
        kinds.add("age")
    return kinds


def unasked_criteria(segment, notes):
    """The criteria a segment targets that the notes never asked for --
    "homes built before 2014" on a custom the notes only asked for as HVAC
    + homeowners. Gender rides along with an age segment's letter (F35-64)
    and children with a family segment's name, so only income, age and home
    age are read here: the criteria a client would notice were invented."""
    asked = _asked_kinds(notes)
    seen, out = set(), []
    for attr in segment_attributes(segment):
        if attr["kind"] in ("income", "age", "home_built") and attr["kind"] not in asked:
            key = _describe(attr)
            if key not in seen:
                seen.add(key)
                out.append(attr)
    return out


def _components(segment):
    """The catalog names a CUSTOM segment packages ("CUSTOM A AND B" -> A, B)."""
    inner = re.sub(r"^custom\s+", "", str(segment or "").strip(), flags=re.I)
    return [p.strip() for p in re.split(r"\s+(?:AND|&)\s+|\s*,\s*", inner) if p.strip()]


def custom_criteria_items(lines, notes, catalog, option_name=""):
    """Check 3. A CUSTOM segment that packages a criterion the notes didn't
    ask for is never kept silently: a card offers to keep it or to switch to
    the plain segment(s) it was built from, minus the unasked criterion."""
    known = {name.lower(): name for name, _usage in _catalog_rows(catalog)}
    items = []
    for line in lines or []:
        if "segments" not in line:
            continue
        segments = [s for s in (line.get("segments") or []) if str(s).strip()]
        tactic = _line_name(line)
        for segment in segments:
            if not re.match(r"^custom\b", segment, re.I):
                continue
            extras = unasked_criteria(segment, notes)
            if not extras:
                continue
            extra_kinds = {a["kind"] for a in extras}
            plain = [known[c.lower()] for c in _components(segment)
                     if c.lower() in known
                     and not extra_kinds & {a["kind"] for a in segment_attributes(c)}]
            replacement = []
            for seg in segments:
                for new in (plain if seg == segment else [seg]):
                    if new not in replacement:
                        replacement.append(new)
            replacement = replacement[:MAX_AND_SEGMENTS]
            what = ", ".join(_describe(a) for a in extras)
            options = [{"id": "keep", "label": f"Keep {segment}", "patch": None}]
            if plain:
                options.append({
                    "id": "plain",
                    "label": "Switch to " + " AND ".join(replacement),
                    "patch": {"op": "replace_segment", "option": option_name, "tactic": tactic,
                              "from": segment, "to": plain[:MAX_AND_SEGMENTS]},
                })
            switch = f" Keep it, or switch to {' AND '.join(replacement)}?" if plain else ""
            items.append({
                "id": _item_id("custom_extra", option_name, tactic, segment),
                "check": "custom_unasked_criteria",
                "type": "choice" if plain else "fyi",
                "priority": "decision" if plain else "fyi",
                "question": (f"The {tactic} line uses {segment}. This custom also targets {what}, "
                             f"which the notes didn't ask for.{switch}"),
                "options": options,
                "default": "keep",
                "context": {"option": option_name, "tactic": tactic, "segment": segment,
                            "unasked": [_describe(a) for a in extras],
                            "plain": replacement if plain else None},
            })
    return items


_CLAUSE_LEFT = (", ", "; ", " with ", " who ", " whose ", " that ", " in ", " and ")
_CLAUSE_RIGHT = (", ", "; ", ". ", ")", " in ", " who ", " with ", " and ")


def _remove_clause(text, phrase):
    """`text` without the clause holding `phrase`: from the nearest clause
    start before it (", with a focus on") to the next clause end. Falls back
    to the phrase alone when the clause is the whole text."""
    low = text.lower()
    at = low.find(phrase.lower())
    if at < 0:
        return text
    end = at + len(phrase)
    left = max((low.rfind(sep, 0, at) for sep in _CLAUSE_LEFT), default=-1)
    rights = [i for i in (low.find(sep, end) for sep in _CLAUSE_RIGHT) if i >= 0]
    right = min(rights) if rights else len(text)
    if left <= 0:
        out = text[:at] + text[end:]
    else:
        out = text[:left] + text[right:]
    out = re.sub(r"\s{2,}", " ", out)
    out = re.sub(r"\s+([,.;)])", r"\1", out)
    out = re.sub(r"\(\s*\)", "", out).strip(" ,;-")
    out = re.sub(r"\s+(with|and|in|who|that|whose)$", "", out, flags=re.I)
    out = re.sub(r"^(who are|who|that are|with|and|in)\s+", "", out, flags=re.I)
    return out[:1].upper() + out[1:] if out else out


def strip_unasked_criteria(text, notes):
    """Specs copy and a line's Targeting state only criteria the notes asked
    for (QA-023): an income, age or home-age criterion the notes never
    stated is taken out of `text`. Returns (text, [phrases removed])."""
    asked = _asked_kinds(notes)
    out, removed = str(text or ""), []
    for attr in text_attributes(out):
        if attr["kind"] in ("income", "age", "home_built") and attr["kind"] not in asked:
            new = _remove_clause(out, attr["phrase"])
            if new and new != out:
                out = new
                removed.append(attr["phrase"])
    return out, removed


def targeting_review_items(lines, notes, catalog, option_name=""):
    """All three checks, deduplicated by id."""
    items = (uncovered_attribute_items(lines, catalog, option_name)
             + threshold_mismatch_items(lines, notes, catalog, option_name)
             + custom_criteria_items(lines, notes, catalog, option_name))
    seen, out = set(), []
    for item in items:
        if item["id"] not in seen:
            seen.add(item["id"])
            out.append(item)
    return out


def review_item_text(item):
    """One line for the existing prose review list, until the interactive
    follow-ups render items as cards: the question plus its fixes."""
    fixes = [o["label"] for o in item.get("options") or [] if o.get("patch")]
    if not fixes:
        return item["question"]
    return item["question"] + " Fix: " + " -- or -- ".join(fixes) + "."
