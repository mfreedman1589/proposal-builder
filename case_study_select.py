"""Which case studies, and which of their slides, go into a proposal.

Pure -- no Streamlit, no database (Burch Oil, 2026-11-16). Three rules:

1. **A study goes in once.** Most vault decks hold the same study twice in
   two layouts, and every Total TV study holds a plain slide plus a
   TOTAL tv-branded one. Those are variants, not pages: `pick_slides`
   keeps one -- the TOTAL tv-branded slide when the plan has broadcast,
   the plain one when it doesn't. Slides are variants when their words
   overlap heavily (measured on the real vault: variants 0.63-1.0,
   the one genuinely two-page study 0.03).
2. **Two vault rows can be one study** (the vault holds "Local Healthcare
   Organization -- Brand Preference Lift in a Top 10 Market" and "... --
   Top 10 Market Brand Preference Lift"): `same_study` compares the
   subject before the dash and the descriptor's words, ignoring channel
   words (Total TV, Streaming TV) so a plain and a Total TV row match.
3. **A specialty narrows a vertical.** An HVAC plan doesn't pre-select a
   pest control study. `specialties_of` reads a study's title and summary,
   and the plan's notes, against `SPECIALTIES`; a study with no specialty
   ("where tagged") is never filtered out.
"""
import hashlib
import re

# The TOTAL tv logo the branded variant of every Total TV study carries
# (identical image in all five on 2026-11-16). The fallback below catches a
# re-exported logo: within a Total TV study, the slide with a picture its
# twin lacks is the branded one.
TOTAL_TV_LOGO_SHA1 = {"1da9cd8568a8"}

VARIANT_SIMILARITY = 0.5

# Words that name a channel, not the study.
_CHANNEL_WORDS = {"total", "tv", "streaming", "ott", "ctv", "premion", "with", "and", "the",
                  "a", "an", "in", "of", "for", "via", "to", "case", "study"}

# Specialties within a vertical. Whole-word, case-insensitive; a study or a
# plan can have several ("HVAC & Plumbing").
SPECIALTIES = {
    "home_improvement": {
        "hvac": ["hvac", "heating", "air conditioning", "furnace", "furnaces", "heat pump",
                 "heat pumps", "cooling"],
        "plumbing": ["plumbing", "plumber", "plumbers", "water heater"],
        "pest_control": ["pest control", "pest", "exterminator", "termite", "termites"],
        "roofing": ["roofing", "roofer", "roof", "roofs"],
        "remodeling": ["remodeling", "remodel", "renovation", "siding", "windows", "kitchen",
                       "bath"],
        "landscaping": ["landscaping", "lawn", "lawn care", "landscape"],
        "electrical": ["electrician", "electrical"],
        "security": ["home security", "alarm", "alarms"],
        "flooring": ["flooring", "carpet"],
        "solar": ["solar"],
    },
}


def _words(text):
    return set(re.findall(r"[a-z]{4,}", str(text or "").lower()))


def similarity(a, b):
    a, b = _words(a), _words(b)
    return len(a & b) / max(1, len(a | b))


def _picture_hashes(slide):
    out = set()
    for shape in _all_shapes(slide.shapes):
        try:
            blob = shape.image.blob
        except Exception:
            continue
        out.add(hashlib.sha1(blob).hexdigest()[:12])
    return out


def _all_shapes(shapes):
    for shape in shapes:
        yield shape
        if getattr(shape, "shape_type", None) == 6:   # group
            yield from _all_shapes(shape.shapes)


def _slide_text(slide):
    return " ".join(s.text_frame.text for s in _all_shapes(slide.shapes) if s.has_text_frame)


def variant_groups(texts):
    """[[slide index, ...], ...]: slides grouped as variants of one page,
    in deck order. A slide joins the first group whose first slide it
    overlaps by VARIANT_SIMILARITY."""
    groups = []
    for i, text in enumerate(texts):
        for group in groups:
            if similarity(texts[group[0]], text) >= VARIANT_SIMILARITY:
                group.append(i)
                break
        else:
            groups.append([i])
    return groups


def slide_facts(prs):
    """[(text, picture hashes)] per slide of a case study deck."""
    return [(_slide_text(s), _picture_hashes(s)) for s in prs.slides]


def pick_slides(facts, has_broadcast, total_tv_study=False):
    """0-based slide indices to insert: one per page. Among variants, the
    TOTAL tv-branded slide when the plan has broadcast, else the plain one;
    a study with no branded variant keeps its first layout."""
    texts = [t for t, _h in facts]
    picked = []
    for group in variant_groups(texts):
        if len(group) == 1:
            picked.append(group[0])
            continue
        branded = [i for i in group if facts[i][1] & TOTAL_TV_LOGO_SHA1]
        if not branded and total_tv_study:
            # A slide carrying a picture its twin lacks.
            branded = [i for i in group
                       if any(facts[i][1] - facts[j][1] for j in group if j != i)]
        plain = [i for i in group if i not in branded]
        if has_broadcast and branded:
            picked.append(branded[0])
        else:
            picked.append((plain or group)[0])
    return sorted(picked)


def _split_title(title):
    parts = re.split(r"\s+(?:--|–|—|-)\s+|\s*[–—]\s*|\s+�\s+", str(title or ""), maxsplit=1)
    subject = parts[0]
    descriptor = parts[1] if len(parts) > 1 else ""
    return subject, descriptor


def _key_words(text):
    return {w for w in re.findall(r"[a-z0-9]+", str(text or "").lower()) if w not in _CHANNEL_WORDS}


def same_study(a_title, b_title):
    """True when two vault titles name one study (see module docstring)."""
    a_subject, a_desc = _split_title(a_title)
    b_subject, b_desc = _split_title(b_title)
    if _key_words(a_subject) != _key_words(b_subject) or not _key_words(a_subject):
        return False
    a, b = _key_words(a_desc), _key_words(b_desc)
    if not a and not b:
        return True
    return len(a & b) / max(1, len(a | b)) >= 0.5


def dedupe(rows, has_broadcast, is_total_tv=lambda row: False):
    """Rows with each study once, in order. Of two rows for one study, the
    one matching the plan's broadcast keeps the first one's place."""
    kept = []
    for row in rows or []:
        twin = next((i for i, k in enumerate(kept)
                     if same_study(k.get("title"), row.get("title"))), None)
        if twin is None:
            kept.append(row)
        elif is_total_tv(row) == has_broadcast and is_total_tv(kept[twin]) != has_broadcast:
            kept[twin] = row
    return kept


def specialties_of(text, vertical):
    """The vertical's specialties a text names, as a set of keys."""
    table = SPECIALTIES.get(vertical) or {}
    text = str(text or "").lower()
    found = set()
    for key, terms in table.items():
        if any(re.search(r"\b" + re.escape(term) + r"\b", text) for term in terms):
            found.add(key)
    return found


def fits_specialty(study_text, plan_specialties, vertical):
    """A study fits the plan unless both name specialties and none match.
    An untagged study, or a plan whose specialty isn't known, always fits."""
    study = specialties_of(study_text, vertical)
    if not study or not plan_specialties:
        return True
    return bool(study & set(plan_specialties))
