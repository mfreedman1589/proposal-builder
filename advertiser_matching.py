"""advertiser_matching.py -- fuzzy-with-confirmation matching between a
free-text advertiser/client name (from an attribution export, or typed on
the Build page) and a roster of existing candidates (advertisers, or
proposals' own client_name values).

Pure -- no Streamlit, no DB, no network. Nothing here ever silently picks a
match on the caller's behalf: every candidate that clears the score floor is
returned, ranked, for a human to confirm. The one exception is a byte-exact
match (after normalization), which is returned alone -- there is nothing
left to disambiguate, the same way `match_market` in market_profiles.py
returns a bare key rather than a candidate list when a name is unambiguous.

Reused across three call sites (ATTRIBUTION_REPORT_PLAN.md Phase 2): the
attribution-report upload-first door's advertiser match, the "Build report
from proposal" door's after-the-fact verification, and (later) client_name
entry on the Build-a-proposal page.
"""
import difflib
import re

_ALNUM_RE = re.compile(r"[^a-z0-9]+")

# Company-name noise a DMA-name normalizer doesn't have to deal with (this
# is why this isn't just a reuse of market_profiles._normalize). Word-
# boundary matched so "Company" survives (no boundary lands after "co"
# mid-word) while a standalone "Co"/"LLC"/etc. is dropped.
_LEGAL_SUFFIX_RE = re.compile(
    r"\b(llc|inc|incorporated|corp|corporation|co|company|ltd|limited|group|holdings)\b\.?",
    re.IGNORECASE)


def normalize_name(name):
    """Legal-entity suffixes dropped, then letters and digits only,
    lowercased -- the same "strip what carries no matching signal" idiom
    market_profiles._normalize uses for market names, extended here with
    the noise that's specific to company names. Without the suffix strip,
    "Cardinal Plumbing" (an export's own client name) would never exactly
    match "Cardinal Plumbing, LLC" (a proposal's client_name as typed) --
    a real, expected mismatch this module exists to catch.
    """
    text = _LEGAL_SUFFIX_RE.sub("", (name or "").lower())
    return _ALNUM_RE.sub("", text)


def name_score(a, b):
    """0.0-1.0 similarity between two raw names, via difflib on their
    normalized forms. 0.0 whenever either name is empty after
    normalization -- an empty string is not a match for anything.
    """
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def _overlaps(a_start, a_end, b_start, b_end):
    """True/False, or None when either range is incomplete -- "unknown,"
    never treated as "no overlap." Both ranges are inclusive.
    """
    if None in (a_start, a_end, b_start, b_end):
        return None
    return a_start <= b_end and b_start <= a_end


def _bool_or_none(hint, value):
    """True/False when both sides are known, else None ("unknown") --
    never False just because the HINT was never given, and never False just
    because the CANDIDATE has no value for this field either. Only a real
    hint compared against a real value ever produces a mismatch.
    """
    if hint is None or value is None:
        return None
    return hint == value


def find_candidates(query_name, roster, *, market_hint=None,
                    report_start=None, report_end=None,
                    min_score=0.55, max_results=5):
    """Rank `roster` against `query_name`.

    `roster`: a list of dicts, each at least {"id", "name"}, optionally
    carrying "market" and "flight_start"/"flight_end" (dates or None).

    Returns a list of dicts -- {"id", "name", "score", "market_match",
    "flight_overlap"} -- most-likely first, capped at `max_results`.
    `market_match`/`flight_overlap` are True/False/None (None = the hint or
    the candidate's own data wasn't available to compare). Neither hint
    ever EXCLUDES a candidate that cleared the score floor -- they only
    reorder among real candidates, because a station/market mismatch or a
    report window narrower than a stored flight can be genuine (a rep in
    the wrong originating market; a monthly report inside a longer flight)
    rather than proof of the wrong proposal.

    An exact match (normalized) is returned ALONE, regardless of
    `min_score`/`max_results` -- there is nothing left to disambiguate.
    Never raises; an empty or all-below-floor roster returns [].
    """
    if not roster:
        return []

    scored = []
    for entry in roster:
        score = name_score(query_name, entry.get("name"))
        if score >= 0.999:
            return [{"id": entry["id"], "name": entry.get("name"), "score": 1.0,
                    "market_match": None, "flight_overlap": None}]
        if score >= min_score:
            scored.append((score, entry))

    if not scored:
        return []

    def sort_key(item):
        score, entry = item
        market_match = _bool_or_none(market_hint, entry.get("market"))
        overlap = _overlaps(report_start, report_end,
                            entry.get("flight_start"), entry.get("flight_end"))
        # Neither hint is a hard filter -- both only break ties among
        # candidates that already cleared the score floor. `market_match`/
        # `overlap` being None (unknown) sorts identically to False here,
        # which is intentional: an unknown hint should never look better
        # than a real mismatch, only equal to one.
        return (market_match is not True, overlap is not True, -score)

    scored.sort(key=sort_key)
    results = []
    for score, entry in scored[:max_results]:
        market_match = _bool_or_none(market_hint, entry.get("market"))
        overlap = _overlaps(report_start, report_end,
                            entry.get("flight_start"), entry.get("flight_end"))
        results.append({"id": entry["id"], "name": entry.get("name"), "score": round(score, 4),
                        "market_match": market_match, "flight_overlap": overlap})
    return results
