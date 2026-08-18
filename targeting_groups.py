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
# Color, reserved for the future map (Prompt E). Design hook only -- nothing
# here renders a map. A fixed, ordered, colorblind-legible-ish palette;
# assignment is round-robin by creation order so it's deterministic and never
# looked up by name.
# ---------------------------------------------------------------------------
GROUP_COLORS = [
    "#4C78A8", "#F58518", "#54A24B", "#B279A2", "#E45756",
    "#72B7B2", "#EECA3B", "#FF9DA6", "#9D755D", "#BAB0AC",
]


def assign_color(index):
    return GROUP_COLORS[index % len(GROUP_COLORS)]


def new_group(terms, op=None, geo_def=None, name="", avails_monthly=0,
             avails_basis_assumed=False, color=None, group_id=None):
    """One targeting group. `terms` is the audience expression's parts, in
    build order; `op` is ignored (forced to None) when there's at most one
    term, since a single term has no operator to disagree about."""
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
    """(terms, op) if `text` is exactly the canonical '(A) AND (B) [...]'
    form, else None.

    Parses ONLY that form, and only when it names at least one operator -- a
    bare '(X)' with no AND/OR is deliberately left unparsed (there is no
    group expression this system would ever render as a single parenthesized
    term with no operator; `expression_text` never produces one). That keeps
    this one-way: a hand-typed audience that merely contains a parenthesis or
    the word "and" is never mis-split into bogus catalog terms it doesn't
    resolve to. Mixed AND/OR is rejected -- a group's expression is terms
    joined by ONE operator, never a general boolean tree.
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
    if op is None:
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


def geo_label(group, label_for=None):
    """The Geo cell text -- what a rep and a client see.

    `label_for` is an optional market-key -> display-label lookup (wired in
    once geo resolution lands); until then, and always for a "text" geo_def,
    the stored label is used verbatim. This is deliberately the same
    projection `avails_rows_for_markets(combine=True)` already produces for a
    combined-market row, so nothing downstream of the Geo column changes.
    """
    geo_def = group.get("geo_def") or {}
    kind = geo_def.get("kind")
    if kind == "markets":
        keys = geo_def.get("markets") or []
        labels = [(label_for(k) if label_for else None) or k for k in keys]
        return ", ".join(labels)
    if kind == "counties":
        return ", ".join(geo_def.get("counties") or [])
    if kind == "zips":
        return ", ".join(geo_def.get("zips") or [])
    if kind == "radius":
        centers = geo_def.get("centers") or []
        miles = geo_def.get("miles")
        if not centers:
            return ""
        return f"{miles}mi radius of {', '.join(str(c) for c in centers)}"
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
    """
    return [
        {"Audience": audience_label(group), "Geo": geo_label(group, label_for),
         avails_column: int(group.get("avails_monthly") or 0)}
        for group in (groups or [])
    ]


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

        groups.append(new_group(
            terms, op=op, geo_def={"kind": "text", "label": geo},
            name=(prior.get("name", "") if prior else ""),
            avails_monthly=avails_monthly,
            color=(prior.get("color") if prior else assign_color(index)),
            group_id=(prior.get("id") if prior else None),
        ))
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
