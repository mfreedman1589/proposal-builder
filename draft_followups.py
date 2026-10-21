"""Draft follow-ups: every question a draft leaves open, as a structured item.

After draft-from-notes, each follow-up is an item the rep answers with one
click (or skips), and "Apply changes" applies the chosen patches in code --
no redraft. This module is the item shape and its validation, pure (no DB,
no Streamlit); app.py owns rendering and applying.

An item:
    {"id", "type", "priority", "audience", "question", "options", "default",
     "source", "context"?}
- type: confirm | choice | suggestion | missing_value | fyi
- priority: decision | suggestion | fyi (the card group it renders under)
- audience: "client" (ask the client) or "internal" (the seller checks)
- options: [{"id", "label", "patch", "input"?, "choices"?}] -- `patch` is a
  list of operations, or None for "leave as drafted". An option with an
  `input` ("number" or "groups") takes a value from the rep, substituted
  wherever the patch says "$value".
- default: the option id that applies when the rep skips the item. Always
  the draft as it stands (a None patch), so skipping everything changes
  nothing.

The model writes its own items in the draft JSON ("review_items") using the
model-facing ops below; anything that doesn't validate falls back to a prose
note, shown as FYI.
"""

import copy
import hashlib
import json

TYPES = ("confirm", "choice", "suggestion", "missing_value", "fyi")
PRIORITIES = ("decision", "suggestion", "fyi")
AUDIENCES = ("client", "internal")
KEEP = "keep"
VALUE = "$value"

PRIORITY_HEADINGS = {"decision": "Needs a decision", "suggestion": "Suggestions", "fyi": "FYI"}
AUDIENCE_LABELS = {"client": "Ask the client", "internal": "For you"}

# Operations the model may put in a patch. Each changes a drafted option
# that the app can re-price from its stored lines -- never arithmetic the
# model does itself.
MODEL_OPS = {
    "set_allocations",       # {option, allocations: [{line, allocation}]}
    "set_budget",            # {option, total_budget}
    "set_cpm",               # {option, line, cpm}
    "remove_line",           # {option, line}
    "set_group_allocation",  # {option, group_allocation}
    "set_segments",          # {option, line, segments: [1-2 names], audience_track?}
}
# Operations only the app's own items use.
APP_OPS = {
    "include_groups",        # {ids_on: [...], ids_off: [...]}
    "set_plan_groups",       # {ids: [...] | "$value", among: [...]}
    "set_state",             # {key, value} -- STATE_KEYS only
    "gdw_added_value",       # {added_value: bool}
    "add_segment",           # {option, tactic, segment}
    "remove_phrase",         # {option, tactic, phrase}
    "replace_segment",       # {option, tactic, from, to: [...]}
}
ALL_OPS = MODEL_OPS | APP_OPS
# Form values a patch may set directly -- each a setup-band or plan toggle
# with no other state hanging off it that the app doesn't re-derive itself.
STATE_KEYS = {"total_tv", "agency_gross_up", "avails_include_duplicates"}

# An audience line ANDs at most two segments (targeting_groups.MAX_AND_SEGMENTS).
MAX_AND_SEGMENTS = 2

ALLOCATION_KEYS = {"flat_amount", "percent_of_total", "percent_of_remainder", "split_evenly",
                   "percent_of_avails", "proportional_to_avails"}


def item_id(*parts):
    raw = json.dumps(parts, sort_keys=True, default=str)
    return "rv-" + hashlib.sha1(raw.encode()).hexdigest()[:10]


def _as_patch(patch):
    if patch is None:
        return None
    if isinstance(patch, dict):
        return [patch]
    return list(patch)


def make_item(key, type_, priority, audience, question, options, default=KEEP,
              keep_label="Leave as drafted", source="app", context=None):
    """Build an item. `options` need only `label` and `patch` (an `id` is
    derived when absent). A "leave as drafted" option is added when no
    option carries the default id, so every item can be skipped."""
    opts = []
    for i, opt in enumerate(options or []):
        opt = dict(opt)
        opt.setdefault("id", f"opt{i + 1}")
        opt["patch"] = _as_patch(opt.get("patch"))
        opts.append(opt)
    if not any(o["id"] == default for o in opts):
        opts.insert(0, {"id": default, "label": keep_label, "patch": None})
    item = {
        "id": key if str(key).startswith("rv-") else item_id(key),
        "type": type_ if type_ in TYPES else "fyi",
        "priority": priority if priority in PRIORITIES else "fyi",
        "audience": audience if audience in AUDIENCES else "internal",
        "question": str(question).strip(),
        "options": opts,
        "default": default,
        "source": source,
    }
    if context:
        item["context"] = context
    return item


def fyi_item(text, audience="internal", source="app"):
    """A prose note that fits no structured shape: shown as FYI, dismissable."""
    return make_item(("fyi", audience, str(text)), "fyi", "fyi", audience, str(text), [],
                     keep_label="Dismiss", source=source)


def from_targeting_check(item, audience="internal"):
    """targeting_check's items already carry the shape; give them the
    audience and patch-list form the rest of the system uses."""
    out = copy.deepcopy(item)
    out["audience"] = audience
    out["source"] = "targeting_check"
    out["type"] = "fyi" if out.get("type") not in TYPES else out["type"]
    for opt in out.get("options") or []:
        opt["patch"] = _as_patch(opt.get("patch"))
    return out


# --- the model's own items --------------------------------------------------

def _valid_allocation(value):
    return (isinstance(value, dict) and len(value) == 1
            and next(iter(value)) in ALLOCATION_KEYS)


def _num(value):
    try:
        return float(str(value).replace(",", "").replace("$", ""))
    except (TypeError, ValueError):
        return None


def validate_op(op, option_names, lines_by_option, group_owned, valid_segment=None):
    """One model op, normalized -- or None when it can't be applied.
    `lines_by_option`: {option name: set of line references (label, lower
    case, or product)}. `group_owned`: option names whose Premion lines are
    owned by targeting groups, where only budget and group allocation can
    change."""
    if not isinstance(op, dict) or op.get("op") not in MODEL_OPS:
        return None
    op = dict(op)
    name = op.get("option")
    if name not in option_names:
        if len(option_names) == 1:
            op["option"] = name = next(iter(option_names))
        else:
            return None
    lines = lines_by_option.get(name) or set()

    def line_ok(ref):
        return str(ref or "").strip().lower() in lines

    kind = op["op"]
    if kind == "set_budget":
        value = op.get("total_budget")
        if value != VALUE:
            value = _num(value)
            if value is None or value <= 0:
                return None
        op["total_budget"] = value
        return op
    if kind == "set_group_allocation":
        if name not in group_owned or not _valid_allocation(op.get("group_allocation")):
            return None
        return op
    if name in group_owned:
        return None
    if kind == "set_allocations":
        allocations = op.get("allocations")
        if not isinstance(allocations, list) or not allocations:
            return None
        if not all(isinstance(a, dict) and line_ok(a.get("line"))
                   and _valid_allocation(a.get("allocation")) for a in allocations):
            return None
        return op
    if kind == "set_cpm":
        cpm = op.get("cpm")
        if cpm != VALUE:
            cpm = _num(cpm)
            if cpm is None or cpm <= 0:
                return None
        op["cpm"] = cpm
        return op if line_ok(op.get("line")) else None
    if kind == "remove_line":
        return op if line_ok(op.get("line")) else None
    if kind == "set_segments":
        segments = [str(x).strip() for x in op.get("segments") or [] if str(x).strip()]
        if not line_ok(op.get("line")) or not 1 <= len(segments) <= MAX_AND_SEGMENTS:
            return None
        if valid_segment is not None and not all(valid_segment(x) for x in segments):
            return None
        op["segments"] = segments
        return op
    return None


def normalize_model_items(raw_items, option_names, lines_by_option, group_owned=(),
                          valid_segment=None):
    """The model's "review_items" -> (items, prose_client, prose_internal).
    An option whose patch doesn't validate is dropped; an option with an
    empty patch is the drafted value restated (the app adds "Leave as
    drafted" itself) and is dropped quietly; an item left with no applicable
    option becomes prose for its audience (shown as FYI)."""
    items, prose_client, prose_internal = [], [], []
    option_names = set(option_names or ())
    group_owned = set(group_owned or ())
    for raw in raw_items or []:
        if not isinstance(raw, dict):
            continue
        question = str(raw.get("question") or "").strip()
        if not question:
            continue
        audience = raw.get("audience") if raw.get("audience") in AUDIENCES else "internal"
        options = []
        for opt in raw.get("options") or []:
            if not isinstance(opt, dict) or not str(opt.get("label") or "").strip():
                continue
            ops = _as_patch(opt.get("patch")) or []
            valid = [validate_op(op, option_names, lines_by_option, group_owned, valid_segment)
                     for op in ops]
            if not ops or any(v is None for v in valid):
                continue
            entry = {"label": str(opt["label"]).strip(), "patch": valid}
            if any(VALUE in json.dumps(v) for v in valid):
                entry["input"] = "number"
            options.append(entry)
        if not options:
            (prose_client if audience == "client" else prose_internal).append(question)
            continue
        type_ = raw.get("type") if raw.get("type") in TYPES and raw.get("type") != "fyi" else "choice"
        priority = raw.get("priority") if raw.get("priority") in ("decision", "suggestion") else (
            "suggestion" if type_ == "suggestion" else "decision")
        items.append(make_item(("model", audience, question), type_, priority, audience,
                               question, options, source="model",
                               context={"target": raw.get("target")}))
    return items, prose_client, prose_internal


# --- answers ----------------------------------------------------------------

def drop_repeated_prose(notes, items):
    """Prose notes minus any that restate a card's question -- an open
    question is a card or a list item, never both."""
    def key(text):
        return " ".join(str(text).lower().split())[:60]
    asked = {key(i["question"]) for i in items or []}
    return [n for n in notes or [] if key(n) not in asked]


def by_priority(items):
    """{priority: [items]} in display order, empty groups left out."""
    out = {}
    for priority in PRIORITIES:
        group = [i for i in items or [] if i.get("priority") == priority]
        if group:
            out[priority] = group
    return out


def option_of(item, option_id):
    return next((o for o in item.get("options") or [] if o["id"] == option_id), None)


def substitute(patch, value):
    """A patch with every "$value" replaced by the rep's input."""
    def walk(node):
        if node == VALUE:
            return value
        if isinstance(node, dict):
            return {k: walk(v) for k, v in node.items()}
        if isinstance(node, list):
            return [walk(v) for v in node]
        return node
    return walk(copy.deepcopy(patch))


def chosen_patches(items, answers, values=None):
    """[(item, option, ops)] for every item answered with something other
    than its default. An option needing input with none given is skipped."""
    values = values or {}
    out = []
    for item in items or []:
        answer = (answers or {}).get(item["id"], item["default"])
        if answer == item["default"]:
            continue
        option = option_of(item, answer)
        if option is None or not option.get("patch"):
            continue
        ops = option["patch"]
        if option.get("input"):
            value = values.get(item["id"])
            if value in (None, "", []):
                continue
            ops = substitute(ops, value)
        out.append((item, option, ops))
    return out
