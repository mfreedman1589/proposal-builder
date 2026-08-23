"""The targeting group: the shared object behind audience and geo targeting.

No Streamlit, no DB -- pure functions over plain dicts, the same convention
as `geo_resolver.py`, so this is testable offline and has no import cycle
with `app.py` (which is the only caller).

A targeting group is a name, an audience expression (one or more exact
catalog segment names joined by a single AND or OR), a geo definition
(markets / counties / zips / radius / free text), the geography that
definition resolves to, one avails figure for the whole expression, and a
color reserved for the future map. **One group is one avails row is one plan
line, by default** -- "several markets combined into one line" is a group
whose geo_def spans several markets, not several groups; "broken out
separately" is one group per audience-times-single-market pair, created that
way at build time. Merge and split (in app.py) operate on PLAN LINES that
reference group ids; they never change a group.

**Backward compatibility is one property, not a migration script.**
`avails_seed_rows` -- the flat {"Audience","Geo",avails_column} shape every
proposal before this feature stored -- keeps being what `app.py` persists and
what every existing reader (rebuild-as-presented, the drafting reach
calculation, the deck payload) consumes. `groups_to_seed_rows` projects
groups down to that shape; `seed_rows_to_groups` is its inverse, used once at
load time to give an old proposal groups to show. The whole compatibility
story is that these two are inverses:

    groups_to_seed_rows(seed_rows_to_groups(rows, COL), COL) == rows

for every flat row list `rows` -- including a hand-typed audience, one that
happens to contain a comma, and a combined multi-market Geo string. Nothing
about an old proposal's stored, rendered or rebuilt output depends on groups
existing at all.
"""
import re
import uuid

# ---------------------------------------------------------------------------
# Color -- drives the D2 avails table's Color column and the targeting map
# (targeting_map.py). A fixed, ordered palette; assignment is round-robin by
# creation order so it's deterministic and never looked up by name.
#
# Orange/red were `#F58518`/`#E45756` -- reported as too close to tell apart
# on the map (live feedback, 2026-08-23). Both were desaturated, similarly
# light warm hues (the "red" was closer to salmon than a true red), which is
# exactly the pairing that reads as one color at small map-dot sizes. Swapped
# for tab10's true orange/red (`#FF7F0E`/`#D62728`), a bolder and more
# saturated pair with real hue and lightness separation. Every OTHER color
# stays byte-identical: `color` is stored on a group as this literal hex
# string, not re-derived from an index, so a proposal saved before this
# change keeps whatever hex it was actually given -- this only changes what
# a NEW group gets assigned.
# ---------------------------------------------------------------------------
GROUP_COLORS = [
    "#4C78A8", "#FF7F0E", "#54A24B", "#B279A2", "#D62728",
    "#72B7B2", "#EECA3B", "#FF9DA6", "#9D755D", "#BAB0AC",
]


def assign_color(index):
    return GROUP_COLORS[index % len(GROUP_COLORS)]


def audience_cascade_color(groups, audience):
    """The color an audience's groups currently share, or None when the
    audience doesn't exist yet or every one of its groups has been broken
    out (`color_locked`). The one place that answers "what color is THIS
    audience right now" -- a new group for an existing audience inherits
    this; a group that's still following it is never distinguished from
    its siblings on the map or in the legend (see targeting_map.py).

    Deliberately reads whichever UNLOCKED group comes first rather than
    asserting every unlocked group already agrees -- callers that change a
    shared color re-derive it fresh, and a caller mid-cascade-write hasn't
    necessarily reached every sibling yet when this is consulted.
    """
    for group in (groups or []):
        if audience_label(group) == audience and not group.get("color_locked"):
            return group.get("color")
    return None


def new_group(terms, op=None, geo_def=None, name="", avails_monthly=0,
             avails_basis_assumed=False, color=None, color_locked=False, group_id=None):
    """One targeting group. `terms` is the audience expression's parts, in
    build order; `op` is ignored (forced to None) when there's at most one
    term, since a single term has no operator to disagree about.

    `color_locked=False` (the default) means this group follows its
    audience's shared color -- assigned by whoever creates it (usually via
    `audience_cascade_color`, falling back to `assign_color` for an
    audience seen for the first time), and kept in sync by app.py's own
    cascade whenever a rep changes the audience's color from any sibling
    row. `color_locked=True` means a rep deliberately broke this ONE
    group's color out (a single-row edit in the avails table's Color
    column) -- it keeps whatever color it's given from then on, even if
    the audience's shared color later changes, and the map/legend show it
    as its own distinguishable entry rather than folding it into the
    audience's.
    """
    terms = [str(t).strip() for t in (terms or []) if str(t).strip()]
    if len(terms) <= 1:
        op = None
    return {
        "id": group_id or uuid.uuid4().hex[:8],
        "name": str(name or ""),
        "op": op,
        "terms": terms,
        "geo_def": geo_def or {"kind": "text", "label": ""},
        "resolved_zips": [],
        "resolved_markets": [],
        "avails_monthly": int(avails_monthly or 0),
        "avails_basis_assumed": bool(avails_basis_assumed),
        "color": color or GROUP_COLORS[0],
        "color_locked": bool(color_locked),
    }


# ---------------------------------------------------------------------------
# Two renderers, one parser (geo_targeting_roadmap.md section F: "the two
# representations must agree -- one parser, one renderer"). The canonical
# form is parenthesized terms joined by AND/OR, the same shape the Salesforce
# avails documents use; the plain form is what a client reads.
# ---------------------------------------------------------------------------
_PAREN_TERM = re.compile(r"\(([^()]+)\)")
_OP_THEN_PAREN = re.compile(r"^\s+(AND|OR)\s+(?=\()")


def expression_text(group):
    """The canonical form: '(A) AND (B)', or a bare term when there's only
    one -- a single term is equivalent to an unparenthesized one."""
    terms = group.get("terms") or []
    if len(terms) <= 1:
        return terms[0] if terms else ""
    op = group.get("op") or "AND"
    return f" {op} ".join(f"({t})" for t in terms)


def audience_label(group):
    """The plain-language form -- what a client reads. Never boolean syntax:
    AND reads as a comma-joined stack (the same convention `audience_stack()`
    already uses for the Targeting column), OR reads as "or"."""
    terms = group.get("terms") or []
    if not terms:
        return str(group.get("name") or "")
    if len(terms) == 1:
        return terms[0]
    joiner = ", " if (group.get("op") or "AND") == "AND" else " or "
    return joiner.join(terms)


def parse_expression(text):
    """(terms, op) if `text` is the canonical '(A) AND (B) [...]' form, or a
    single bracketed term '(A)' with no operator, else None.

    A LONE bracketed term is accepted (as `([A], None)`) as a strict
    superset of what this parser accepted before: `expression_text` itself
    never PRODUCES that shape for a single-term group (it renders bare, no
    parens), so nothing in this app round-trips differently for accepting
    it now. It exists because a source OUTSIDE this app can reasonably wrap
    even a single term -- Premion's own Salesforce avails export always
    does, e.g. "(AUTO Make Subaru)" with no AND/OR (avails_pdf_import.py) --
    and that document's own audience text should parse through this SAME
    one parser rather than needing a second, parallel one.

    Two or more terms still REQUIRE an operator between them -- mixed
    AND/OR is rejected, and a hand-typed audience that merely contains a
    parenthesis or the word "and" without ever closing back out to this
    exact shape is never mis-split into bogus catalog terms it doesn't
    resolve to. A group's expression is terms joined by ONE operator, never
    a general boolean tree.
    """
    text = str(text or "").strip()
    if not text.startswith("(") or not text.endswith(")"):
        return None
    terms, pos, n, op = [], 0, len(text), None
    while pos < n:
        m = _PAREN_TERM.match(text, pos)
        if not m:
            return None
        terms.append(m.group(1).strip())
        pos = m.end()
        if pos == n:
            break
        om = _OP_THEN_PAREN.match(text[pos:])
        if not om:
            return None
        this_op = om.group(1)
        if op is not None and this_op != op:
            return None
        op = this_op
        pos += om.end()
    if op is None and len(terms) > 1:
        return None
    return terms, op


def terms_from_audience_text(text):
    """(terms, op) for one plain audience string -- the canonical
    parenthesized form when the text IS exactly that, else the whole string
    as ONE VERBATIM TERM, never split on a comma a rep happened to type.

    The one place free-typed audience text becomes group terms, used both by
    the flat-row migration (`seed_rows_to_groups`) and by app.py's grouped
    avails-table fold-back (Phase 4) -- so a hand-typed audience is parsed
    the same way regardless of which editor it came through.
    """
    text = str(text or "").strip()
    parsed = parse_expression(text)
    if parsed:
        return parsed
    return ([text] if text else [], None)


def _count_word(n, noun):
    return f"{n} {noun}{'s' if n != 1 else ''}"


def geo_label(group, label_for=None):
    """The Geo cell text -- what a rep and a client see. **This is a plan
    table cell, not the place for a zip list** -- a client-facing document
    is wrong content regardless of whether it technically fits, which a
    50-zip Zips-mode group proved concretely (geo_targeting_roadmap.md D).
    The full zip list stays on the group (`resolved_zips`/`geo_def["zips"]`)
    for wherever zips actually belong -- the targeting slide, a planning
    export -- never here.

    The rep's own label (`group["name"]`, set in the geo-definition
    expander) always wins, for every kind. Absent that:

    - **markets**: the market names (or raw keys with no `label_for`).
    - **counties**: the county list verbatim -- short enough as-is (at most
      a handful of `NAME STATE` entries; see the roadmap for why this one
      wasn't touched).
    - **zips**: the resolved market(s) plus a zip count ("Philadelphia —
      52 zips"), or just the count before resolution -- never the zips
      themselves. The market name is always humanized through `label_for`
      when the caller supplies one -- never the raw market key/slug, which
      is internal plumbing and reached client-facing text more than once.
    - **radius**: `Nmi radius of <address>` for one or two centers (the
      common single-store case stays exactly as readable as before); for
      more, `Nmi radius of K locations` -- a rep building a 20-store radius
      group hits the identical wrong-content problem the zip list did.
    - **text** (a flat-row migration): the stored label verbatim, unchanged
      -- this is the one branch `groups_to_seed_rows`'s round-trip
      guarantee depends on staying byte-exact.

    `label_for` is an optional market-key -> display-label lookup (wired in
    once geo resolution lands); until then the raw key is used. This is
    deliberately the same projection `avails_rows_for_markets(combine=True)`
    already produces for a combined-market row, so nothing downstream of
    the Geo column changes for the kinds this docstring doesn't call out.
    """
    name = str(group.get("name") or "").strip()
    if name:
        return name
    geo_def = group.get("geo_def") or {}
    kind = geo_def.get("kind")
    if kind == "markets":
        keys = geo_def.get("markets") or []
        labels = [(label_for(k) if label_for else None) or k for k in keys]
        return ", ".join(labels)
    if kind == "counties":
        return ", ".join(geo_def.get("counties") or [])
    if kind == "zips":
        zips = geo_def.get("zips") or []
        resolved = group.get("resolved_markets") or []
        count_word = _count_word(len(zips), "zip")
        if not resolved:
            return count_word
        area = ", ".join((label_for(m) if label_for else None) or m for m in sorted(resolved))
        return f"{area} — {count_word}"
    if kind == "radius":
        centers = geo_def.get("centers") or []
        miles = geo_def.get("miles")
        if not centers:
            return ""
        if len(centers) <= 2:
            names = [c.get("center", "") if isinstance(c, dict) else str(c) for c in centers]
            return f"{miles}mi radius of {', '.join(names)}"
        return f"{miles}mi radius of {_count_word(len(centers), 'location')}"
    return str(geo_def.get("label", "") or "")


# ---------------------------------------------------------------------------
# The compatibility projection
# ---------------------------------------------------------------------------
def groups_to_seed_rows(groups, avails_column, label_for=None):
    """Groups -> the flat {"Audience","Geo",avails_column} shape.

    The one place a group becomes the legacy row shape, so `avails_lookup`,
    the deck payload, and rebuild-as-presented keep working without knowing
    groups exist. `avails_column` is passed explicitly (the caller's real
    `AVAILS_COLUMN_MONTHLY`/basis-labelled column name) rather than
    duplicated as a second literal here, so the two can't drift.

    A `_placeholder` group's row carries the marker forward too -- dropping
    it here would make this projection disagree with what
    `avails_rows_for_markets` computes fresh every run, and
    `apply_avails_autofill`'s clean/dirty check (a plain equality test)
    would then read the untouched placeholder as "edited by something else"
    and stop replacing it the moment a rep actually picks a target market.
    """
    rows = []
    for group in (groups or []):
        row = {"Audience": audience_label(group), "Geo": geo_label(group, label_for),
               avails_column: int(group.get("avails_monthly") or 0)}
        if group.get("_placeholder"):
            row["_placeholder"] = True
        rows.append(row)
    return rows


def seed_rows_to_groups(rows, avails_column, existing=None):
    """Flat rows -> groups. The load-time migration backward compatibility
    depends on.

    Never invents resolved geography: a flat row's Geo is a rendered label,
    not a verified market list, so it becomes `geo_def={"kind":"text",
    "label": geo}` and `resolved_markets`/`resolved_zips` stay empty --
    exactly as honest as the row it came from, no more. The audience string
    becomes ONE VERBATIM TERM (never split on a comma a rep happened to type)
    unless it's already the canonical parenthesized expression form.

    `existing`, when given, is the group list this row list was last
    projected from. A row whose (Audience, Geo, avails) is unchanged from a
    not-yet-consumed existing group keeps that group's id (and its name and
    color), so re-deriving groups from an untouched table doesn't reassign
    ids -- and therefore doesn't orphan any plan line's group_ids -- on every
    rerun. Each existing group is matched to at most one row.

    A row marked `_placeholder` (app.py's `avails_rows_for_markets`, the
    seeded row for "no target market picked yet") still becomes a group --
    that's what gives D2 its one blank starter row on a fresh proposal, same
    as it always has -- but the GROUP carries the marker forward too. A
    caller that's about to ADD groups on top of what's already there (the
    avails PDF importer's `_finish_avails_import`) is expected to drop any
    `_placeholder` group from `existing` first, the same way picking a real
    target market already replaces this row outright rather than sitting
    beside it (see avails_rows_for_markets) -- appending is exactly the
    operation that skips that replacement, which is how a blank placeholder
    group used to survive forever alongside real imported ones. A
    market-only row from an ACTUAL target-market pick never carries this
    marker -- that one is a genuine in-progress state, not a placeholder.
    """
    prior_by_row = {}
    for group in (existing or []):
        key = (audience_label(group), geo_label(group), int(group.get("avails_monthly") or 0))
        prior_by_row.setdefault(key, []).append(group)

    groups = []
    for index, row in enumerate(rows or []):
        audience = str(row.get("Audience", "") or "").strip()
        geo = str(row.get("Geo", "") or "").strip()
        avails_monthly = int(row.get(avails_column, 0) or 0)
        if not audience and not geo:
            continue

        bucket = prior_by_row.get((audience, geo, avails_monthly))
        prior = bucket.pop(0) if bucket else None

        terms, op = terms_from_audience_text(audience)

        group = new_group(
            terms, op=op, geo_def={"kind": "text", "label": geo},
            name=(prior.get("name", "") if prior else ""),
            avails_monthly=avails_monthly,
            color=(prior.get("color") if prior else assign_color(index)),
            group_id=(prior.get("id") if prior else None),
        )
        if row.get("_placeholder"):
            group["_placeholder"] = True
        groups.append(group)
    return groups


def custom_segment_count(groups, rfp_map):
    """How many DISTINCT non-RFP-selectable segments are in play across
    every group in the campaign -- not per group. The same custom segment
    reused across two groups (e.g. one audience split across two markets) is
    still one Salesforce request and counts once; two different custom
    segments in two different groups count as two, which is the whole point
    of counting across groups rather than within one.
    """
    customs = set()
    for group in (groups or []):
        for term in group.get("terms") or []:
            if not rfp_map.get(term, True):
                customs.add(term)
    return len(customs)
