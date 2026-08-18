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
"""
from collections import Counter

from market_profiles import _normalize as normalize

WIDELY_USED_PERCENTILE = 90


class AudienceIndex:
    """Everything `evidence_for` needs, built once from the usage rows.

    `stacks` is the set of DISTINCT booked stacks (a stack booked twice is
    one entry, per the "component familiarity" finding above).
    `component_counts` is {normalized component: number of distinct stacks
    containing it}. `pair_counts` is {frozenset({a, b}): number of distinct
    stacks containing BOTH} -- built sparsely, only for pairs actually
    observed together, since stacks are small. `display_names` maps a
    normalized component back to its most common raw spelling, for
    rendering -- several raw spellings can fold to the same normalized key
    ("Lifestyle Charity" / "LIFESTYLE Charity"), and the most common one
    reads as the "real" name.
    """

    def __init__(self, stacks, component_counts, pair_counts, display_names):
        self.stacks = stacks
        self.component_counts = component_counts
        self.pair_counts = pair_counts
        self.display_names = display_names

    def display(self, component):
        return self.display_names.get(component, component)


def parse_stack(segment_string):
    """The SET of normalized components in one booked audience string.

    The comma is confirmed AND (see this module's docstring), and matching
    has to be a SET -- 24 stored combinations appear in more than one part
    order, so a tuple or the raw string would miss a stack that has, in
    fact, been booked before.
    """
    parts = [p.strip() for p in str(segment_string or "").split(",") if p.strip()]
    return frozenset(normalize(p) for p in parts if normalize(p))


def _usage_row_text(row):
    """The raw segment string from either column-name convention."""
    if "segment_string" in row:
        return row.get("segment_string")
    return row.get("segment")


def build_index(rows):
    """An `AudienceIndex` from usage rows in either shape."""
    stacks = set()
    raw_by_normalized = Counter()          # normalized -> Counter(raw spelling -> count)
    display_votes = {}
    for row in rows or []:
        text = _usage_row_text(row)
        parts = [p.strip() for p in str(text or "").split(",") if p.strip()]
        stack = frozenset(normalize(p) for p in parts if normalize(p))
        if not stack:
            continue
        stacks.add(stack)
        for raw in parts:
            key = normalize(raw)
            if not key:
                continue
            display_votes.setdefault(key, Counter())[raw.strip()] += 1

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
    return AudienceIndex(stacks, component_counts, pair_counts, display_names)


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
      familiarity: [(display name, distinct-stack count), ...], in the
        order `components` was given.
      weakest_pair: None, or (display a, display b, count) for the LOWEST
        co-occurrence pair among `components` -- with 3+ components, only
        the weakest pair is surfaced (the design's own rule: the other
        pairs don't need narrating).
      weakest_pair_kind: None, "both_widely_used", or "either_rare" -- which
        sentence a zero co-occurrence is allowed to claim, derived from
        whether BOTH of the weakest pair clear the widely-used threshold.
      suggestions: [(display name, weighted score), ...], overlap-weighted,
        excluding anything already selected.
    """
    normalized = [normalize(c) for c in components]
    threshold = widely_used_threshold(index)

    exact_match = frozenset(n for n in normalized if n) in index.stacks

    familiarity = [(comp, index.component_counts.get(norm, 0))
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

    for name, count in evidence["familiarity"]:
        stacks_word = "stack" if count == 1 else "stacks"
        lines.append(f"{name} appears in {count} booked {stacks_word}.")

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
