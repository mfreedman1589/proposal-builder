"""Booking evidence for the audience builder (Phase 7 of the targeting-groups
roadmap, geo_targeting_roadmap.md D -- read "Show booking evidence while
building" and "Finding: the comma in audience_usage is AND" there for the
full design and the analysis this module's shape comes from).

Pure -- no Streamlit, no DB, no framework. Reads whatever rows
`db.fetch_audience_usage()` handed back or the local
`audience_usage_ytd.csv` fallback loaded; the two disagree on column names
(`segment_string`/`delivered_impressions` in the DB, `segment`/`impressions`
locally) and `build_index` accepts either shape, or a mix.

Five things settled by the analysis, load-bearing for everything below:

- **The comma is AND, confirmed not assumed** -- 250 rows spell out "and",
  121 spell out "or", always INSIDE one component; the comma is reserved for
  the stack. `parse_stack` splits on it accordingly.
- **Match as a SET, not a string** -- 24 stored combinations appear in more
  than one part order, so string/tuple matching would miss a stack that has
  in fact been booked.
- **Exact-stack counts are nearly useless as a frequency signal** -- almost
  every distinct stack was booked exactly once, twice at most, so
  `evidence_lines` never states a count for the exact-match line, and never
  says "not booked before" (absence is the normal case; flagging it trains
  people to ignore the panel).
- **Component familiarity is the strong signal** -- how many DISTINCT
  stacks a component appears in, which is why `build_index` counts stacks a
  component is IN, not rows it appears on (a stack booked twice must not
  double-count).
- **"Widely used" is the live 90th percentile of the component-count
  distribution, never a frozen constant** -- `widely_used_threshold`
  recomputes it from whatever index it's handed.

**CLT and client-retargeting exclusion.** A `CLT`-prefixed component is a
client's own first-party data list; a `WEB RT`/`LOCATION RT`/
`..._RFPID-..._RT`/address-list component is the same thing without the
CLT prefix -- one advertiser's own retargeting pool. Neither is ever
selectable for anyone else, and neither is ever surfaced here: not in
`component_counts`/`component_impressions`, not in a pair, not as a
suggested pairing. `_split_parts` is the one place this is enforced,
shared by `parse_stack` and `build_index` so it can't drift between them;
a stack that also has real components keeps those (only the excluded part
of it drops out), and a stack that was ONLY an excluded component
contributes nothing at all, the same as a blank row. A handful of
one-off client names carry no structural marker at all and can't be
caught this way -- those are excluded from the CATALOG (via
`audience_component_overrides.csv`, see `audience_usage_import.py`) but
not from this index, since reading that registry here would pull in
db.py/streamlit through audience_catalog, against this module's own
design.

**Two different popularity signals, on purpose.** `familiarity` (in
`evidence_for`) reports total delivered IMPRESSIONS -- "rank by
impressions, not by count" applies here too, since a component that moved
huge volume once is more relevant than one booked five times at a trickle.
Everything about PRECEDENT -- the weakest pair, `widely_used_threshold`,
suggested pairings -- stays on `component_counts` (distinct stacks), since
those questions are "has this combination been booked before," not "how
much volume did it move."
"""
from collections import Counter
import re

from market_profiles import _normalize as normalize

WIDELY_USED_PERCENTILE = 90

_CLT_RE = re.compile(r"^CLT(?=[_\s]|$)", re.IGNORECASE)
# One advertiser's own retargeting/address-list pool, structurally
# recognizable without a CLT prefix -- same rule and same regex as
# `audience_usage_import.is_client_pattern` (see its own comment for why
# this is a search, not an anchored match, and the false-positive check
# against the real workbook). A handful of one-off client names with no
# structural marker at all can't be caught this way and stay OUT of the
# catalog via `audience_component_overrides.csv` instead -- that registry
# isn't consulted here, since importing it would pull in db.py/streamlit
# through audience_catalog, which is exactly what this module's own "no
# Streamlit, no DB" design avoids.
_CLIENT_PATTERN_RE = re.compile(
    r"(WEB\s*RT|LOCATION\s*RT|_RFPID-[^_]*_RT|ADDRESS\s*LIST)", re.IGNORECASE)


class AudienceIndex:
    """Everything `evidence_for` needs, built once from the usage rows.

    `stacks` is the set of DISTINCT booked stacks (a stack booked twice is
    one entry, per the "component familiarity" finding above), with any
    CLT component already dropped out of every stack.
    `component_counts` is {normalized component: number of distinct stacks
    containing it} -- the PRECEDENT signal (weakest pair, widely-used
    threshold, suggested pairings all read this one). `component_impressions`
    is {normalized component: total delivered impressions across every row
    containing it} -- the POPULARITY signal `evidence_for`'s familiarity
    line reports; unlike component_counts this is summed over ROWS, not
    deduped stacks, since a component that happens to book the identical
    stack twice really did deliver twice the impressions. `pair_counts` is
    {frozenset({a, b}): number of distinct stacks containing BOTH} -- built
    sparsely, only for pairs actually observed together, since stacks are
    small. `display_names` maps a normalized component back to its most
    common raw spelling, for rendering -- several raw spellings can fold to
    the same normalized key ("Lifestyle Charity" / "LIFESTYLE Charity"),
    and the most common one reads as the "real" name.
    """

    def __init__(self, stacks, component_counts, component_impressions, pair_counts, display_names):
        self.stacks = stacks
        self.component_counts = component_counts
        self.component_impressions = component_impressions
        self.pair_counts = pair_counts
        self.display_names = display_names

    def display(self, component):
        return self.display_names.get(component, component)


def _is_clt(raw_component):
    """A client's own first-party data list -- never selectable for anyone
    else, and never counted toward familiarity, precedent, or a suggested
    pairing. Same detection rule `audience_usage_import.is_clt` uses; kept
    as its own regex rather than an import, since this module has no
    DB/openpyxl dependency and the two modules otherwise don't know about
    each other."""
    return bool(_CLT_RE.match(raw_component.strip()))


def _is_client_pattern(raw_component):
    """One of the four structural markers of a single advertiser's own
    retargeting/address-list pool -- same treatment as CLT, for the same
    reason. See `_CLIENT_PATTERN_RE`'s own comment for what this does and
    doesn't catch."""
    return bool(_CLIENT_PATTERN_RE.search(raw_component))


def _split_parts(segment_string):
    """The raw, non-empty parts of one stack string, with CLT and
    client-retargeting-pattern components excluded -- the shared core
    `parse_stack` and `build_index` both use, so exclusion can't drift
    between them. A stack that was ONLY an excluded component ends up with
    no parts at all, same as a genuinely blank row."""
    parts = [p.strip() for p in str(segment_string or "").split(",") if p.strip()]
    return [p for p in parts if not _is_clt(p) and not _is_client_pattern(p)]


def parse_stack(segment_string):
    """The SET of normalized components in one booked audience string.

    The comma is confirmed AND (see this module's docstring), and matching
    has to be a SET -- 24 stored combinations appear in more than one part
    order, so a tuple or the raw string would miss a stack that has, in
    fact, been booked before.
    """
    return frozenset(normalize(p) for p in _split_parts(segment_string) if normalize(p))


def _usage_row_text(row):
    """The raw segment string from either column-name convention."""
    if "segment_string" in row:
        return row.get("segment_string")
    return row.get("segment")


def _usage_row_impressions(row):
    """The impressions figure from either column-name convention -- 0 for
    a row that carries none, so a malformed/missing figure degrades the
    popularity signal rather than raising mid-index-build."""
    value = row.get("delivered_impressions")
    if value is None:
        value = row.get("impressions")
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def build_index(rows):
    """An `AudienceIndex` from usage rows in either shape."""
    stacks = set()
    display_votes = {}
    component_impressions = Counter()
    for row in rows or []:
        text = _usage_row_text(row)
        impressions = _usage_row_impressions(row)
        parts = _split_parts(text)
        stack = frozenset(normalize(p) for p in parts if normalize(p))
        if not stack:
            continue
        stacks.add(stack)
        for raw in parts:
            key = normalize(raw)
            if not key:
                continue
            display_votes.setdefault(key, Counter())[raw.strip()] += 1
            component_impressions[key] += impressions

    component_counts = Counter()
    pair_counts = Counter()
    for stack in stacks:
        for component in stack:
            component_counts[component] += 1
        ordered = sorted(stack)
        for i, a in enumerate(ordered):
            for b in ordered[i + 1:]:
                pair_counts[frozenset((a, b))] += 1

    display_names = {key: votes.most_common(1)[0][0] for key, votes in display_votes.items()}
    return AudienceIndex(stacks, component_counts, component_impressions, pair_counts, display_names)


def widely_used_threshold(index):
    """The live 90th percentile of the component-count distribution --
    recomputed from `index` every call, never frozen. Linear-interpolated
    (the standard definition), matching the validation in
    geo_targeting_roadmap.md D: median 2, p75 5, p80 6, p90 ~10, p95 19
    against the committed CSV, selecting roughly the top decile of
    components as "widely used."

    Returns 0 for an empty index -- nothing is "widely used" against no
    data, and a caller comparing a real count against 0 always gets True,
    which is the honest answer when there's nothing to compare to.
    """
    values = sorted(index.component_counts.values())
    if not values:
        return 0
    if len(values) == 1:
        return values[0]
    position = (WIDELY_USED_PERCENTILE / 100) * (len(values) - 1)
    lo = int(position)
    hi = min(lo + 1, len(values) - 1)
    frac = position - lo
    return values[lo] * (1 - frac) + values[hi] * frac


def _pair_count(index, a, b):
    return index.pair_counts.get(frozenset((a, b)), 0)


def suggested_pairings(index, selected, limit=8):
    """Components that most often accompany the current selection, weighted
    by OVERLAP rather than requiring a strict superset.

    The strict version -- stacks that are supersets of every selected
    segment -- is excellent for one segment and collapses to a flat list of
    1x the moment a second is added, because few stacks are supersets of a
    specific pair (geo_targeting_roadmap.md D's own diagnosis). Weighting
    each candidate by how many of the SELECTED segments its stack shares
    restores real signal: a stack sharing 2 of 2 selected segments counts
    twice as much toward its other components as a stack sharing only 1.

    `selected` is normalized keys already. Returns [(normalized, score), ...]
    sorted by score descending, excluding anything already selected.
    """
    selected_set = set(selected)
    if not selected_set:
        return []
    scores = Counter()
    for stack in index.stacks:
        overlap = len(stack & selected_set)
        if not overlap:
            continue
        for component in stack - selected_set:
            scores[component] += overlap
    return scores.most_common(limit)


def evidence_for(components, index):
    """The full evidence picture for a set of chosen segment names (display
    text, as the builder holds them -- normalized internally).

    Returns a dict:
      exact_match: bool -- has this exact SET been booked before.
      familiarity: [(display name, total delivered impressions), ...], in
        the order `components` was given -- the POPULARITY signal ("rank
        by impressions, not by count"); precedent questions below stay on
        distinct-stack counts instead, see this module's own docstring.
      weakest_pair: None, or (display a, display b, count) for the LOWEST
        co-occurrence pair among `components` -- with 3+ components, only
        the weakest pair is surfaced (the design's own rule: the other
        pairs don't need narrating).
      weakest_pair_kind: None, "both_widely_used", or "either_rare" -- which
        sentence a zero co-occurrence is allowed to claim, derived from
        whether BOTH of the weakest pair clear the widely-used threshold.
      suggestions: [(display name, weighted score), ...], overlap-weighted,
        excluding anything already selected.

    CLT and client-retargeting-pattern terms are filtered out of
    `components` itself before anything else, on the same rule
    `_split_parts` uses -- neither can reach here through the catalog
    (both are excluded there entirely), but a hand-typed custom audience
    could in principle match one, and the point of keeping them out of the
    evidence index is that they never turn up in ANY evidence output, not
    just the ones sourced from the index. A selection that was ONLY such a
    term degrades to an empty one (no exact match, no familiarity, no
    pairing to narrate).
    """
    components = [c for c in components if c and not _is_clt(c) and not _is_client_pattern(c)]
    normalized = [normalize(c) for c in components]
    threshold = widely_used_threshold(index)

    exact_match = frozenset(n for n in normalized if n) in index.stacks

    familiarity = [(comp, index.component_impressions.get(norm, 0))
                  for comp, norm in zip(components, normalized)]

    weakest_pair, weakest_kind = None, None
    if len(normalized) >= 2:
        pairs = [(normalized[i], normalized[j])
                for i in range(len(normalized)) for j in range(i + 1, len(normalized))]
        counted = [(a, b, _pair_count(index, a, b)) for a, b in pairs]
        a, b, count = min(counted, key=lambda p: p[2])
        weakest_pair = (index.display(a), index.display(b), count)
        if count == 0:
            a_count = index.component_counts.get(a, 0)
            b_count = index.component_counts.get(b, 0)
            weakest_kind = ("both_widely_used" if a_count >= threshold and b_count >= threshold
                            else "either_rare")

    suggestions = [(index.display(comp), score)
                  for comp, score in suggested_pairings(index, normalized)]

    return {
        "exact_match": exact_match,
        "familiarity": familiarity,
        "weakest_pair": weakest_pair,
        "weakest_pair_kind": weakest_kind,
        "suggestions": suggestions,
    }


def evidence_lines(evidence):
    """Plain-language lines for `evidence`, in the settled panel order:
    exact match (only when true, never a count, never "not booked before"),
    component familiarity, the weakest pair's sentence, then suggested
    pairings. Nearest-neighbour stacks (item 5 of the design) are
    deliberately not built -- "skip rather than delay," per the roadmap.
    """
    lines = []
    if evidence["exact_match"]:
        lines.append("This exact audience has been booked before.")

    for name, impressions in evidence["familiarity"]:
        lines.append(f"{name} has delivered {impressions:,} impressions.")

    weakest = evidence["weakest_pair"]
    if weakest:
        a, b, count = weakest
        if count > 0:
            times_word = "time" if count == 1 else "times"
            lines.append(f"{a} and {b} have been booked together {count} {times_word}.")
        elif evidence["weakest_pair_kind"] == "both_widely_used":
            lines.append(f"Both {a} and {b} are widely used, but haven't been booked together before.")
        else:
            lines.append(f"{a} and {b} haven't been booked together before.")

    if evidence["suggestions"]:
        names = ", ".join(name for name, _score in evidence["suggestions"][:5])
        lines.append(f"Often paired with what's selected: {names}.")

    return lines
