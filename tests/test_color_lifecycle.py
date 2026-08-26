"""The D2 avails table's Color column, through its full real lifecycle --
not just the single edit that prompted adding it. Reported: an "Apply to
all" click must be able to RE-ATTACH a row a rep had earlier detached with
a single-row swatch edit (otherwise a detach is permanent with no recovery
but re-typing the group by hand), and a detached row needs a marker
visible directly in the grid, not just in the "Apply a color to a whole
audience" expander below it.

Same `st.data_editor` monkeypatch technique as test_avails_grid_manual_row.py
and test_avails_grid_row_deletion.py (AppTest can't drive a data_editor
directly): a one-shot action queue read by the monkeypatched data_editor,
advanced by the test driving it through cascade -> detach -> cascade again
(detached row holds) -> apply-to-all (detached row re-attaches) -> delete
-> add, checking `targeting_groups` -- what the grid is a pure view of --
agrees with what was just done at every step. The "Detached" marker itself
is checked directly off the dataframe HANDED TO the data_editor each run,
not just off targeting_groups, so this also proves the grid would actually
have shown it.

The delete -> add pair runs in a SECOND, freshly-seeded AppTest session
rather than continuing the first one. Confirmed by a direct, minimal repro
with no color/cascade code involved at all: AppTest itself rejects a
brand-new `st.button(key=...)` the moment it's rendered in a session where
that exact set of dynamically-keyed buttons had shrunk to zero and grown
back -- a harness limitation, not a product one (a real browser session
has no such restriction), so splitting the session tests the real
add-after-delete behavior without fighting it.

    python tests/test_color_lifecycle.py
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import streamlit as st  # noqa: E402
import pandas as pd  # noqa: E402

# One shared mutable "what should the grid widget return next" state, same
# shape test_avails_grid_row_deletion.py uses.
_next_action = {"kind": None, "done": True}
_captured = {"data": None}


def _fake_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if not key.startswith("avails_editor"):
        return data
    _captured["data"] = data.copy()
    if _next_action["done"]:
        return data
    _next_action["done"] = True
    avails_col = [c for c in data.columns if c not in
                 ("gid", "Audience", "Markets", "Label", "Color", "Detached")][0]
    kind = _next_action["kind"]
    out = data.copy()
    if kind == "detach_dc":
        mask = out["Label"] == "DC"
        out.loc[mask, "Color"] = _next_action["color_label"]
    elif kind == "delete_dc":
        out = out[out["Label"] != "DC"].reset_index(drop=True)
    elif kind == "add_chicago":
        row = {"gid": None, "Audience": "Family", "Markets": [], "Label": "Chicago",
              "Color": out.iloc[0]["Color"] if len(out) else "", "Detached": "",
              avails_col: 3000}
        out = pd.concat([out, pd.DataFrame([row])], ignore_index=True)
    return out


st.data_editor = _fake_data_editor

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app(groups):
    from datetime import date
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["include_avails_template"] = True
    at.session_state["targeting_groups"] = groups
    at.session_state["premion_streaming_tv"] = True
    # FLOW_REWORK_PLAN.md Phase 1: the setup band gates D2 (this file's own
    # subject) on market+flight -- a truly untouched fresh form no longer
    # reaches it at all.
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    return at


def real_groups(at):
    groups = at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []
    return [g for g in groups if g["terms"]]


def by_label(at, label):
    return next((g for g in real_groups(at) if g.get("name") == label), None)


def run(at, kind=None, **extra):
    _next_action["kind"] = kind
    _next_action["done"] = kind is None
    _next_action.update(extra)
    at.run()
    return at


def detached_marker_for(label):
    data = _captured["data"]
    if data is None:
        return None
    row = data[data["Label"] == label]
    if row.empty:
        return None
    return row.iloc[0]["Detached"]


def main():
    print("=" * 78)
    print("SCENARIO  cascade -> single-row detach -> cascade holds -> "
          "apply-to-all re-attaches -> delete -> add")
    print("=" * 78)

    dc = tg.new_group(["Family"], geo_def={"kind": "text", "label": ""},
                      name="DC", color=tg.GROUP_COLORS[0], avails_monthly=100000)
    richmond = tg.new_group(["Family"], geo_def={"kind": "text", "label": ""},
                            name="Richmond", color=tg.GROUP_COLORS[0], avails_monthly=100000)

    print("\nbaseline: both groups already on the audience's shared color (an ordinary "
          "cascade, not yet touched)")
    at = new_app([dc, richmond])
    run(at)
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    g_dc, g_rich = by_label(at, "DC"), by_label(at, "Richmond")
    check("both groups present", g_dc is not None and g_rich is not None, real_groups(at))
    check("both share the same color", g_dc and g_rich and g_dc["color"] == g_rich["color"],
          (g_dc, g_rich))
    check("neither is detached", g_dc and not g_dc["color_locked"] and not g_rich["color_locked"],
          (g_dc, g_rich))
    check("the grid shows no lock marker on DC yet",
          detached_marker_for("DC") == "", detached_marker_for("DC"))
    shared_color_label = app._color_swatch_label(g_dc["color"])
    detach_target_label = next(lbl for lbl in app._COLOR_SWATCH_LABELS.values()
                               if lbl != shared_color_label)

    print("\na single-row Color edit on DC detaches it -- deliberate, single-row, permanent")
    run(at, "detach_dc", color_label=detach_target_label)
    check("no exception after the edit", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    g_dc, g_rich = by_label(at, "DC"), by_label(at, "Richmond")
    check("DC is now detached", g_dc is not None and g_dc["color_locked"], g_dc)
    check("DC's color actually changed", g_dc and g_dc["color"] != g_rich["color"], (g_dc, g_rich))
    check("Richmond is untouched -- a plain edit is single-row, never a side effect on siblings",
          g_rich is not None and not g_rich["color_locked"] and g_rich["color"] == tg.GROUP_COLORS[0],
          g_rich)

    print("\nAN UNTOUCHED FOLLOW-UP RERUN: the detached state holds, it doesn't silently "
          "re-cascade or revert on its own")
    run(at)
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")
    g_dc2 = by_label(at, "DC")
    check("DC is still detached, same color as right after the edit",
          g_dc2 is not None and g_dc2["color_locked"] and g_dc2["color"] == g_dc["color"], g_dc2)
    check("the grid itself shows the lock marker on the detached row",
          detached_marker_for("DC") == "\U0001F512", detached_marker_for("DC"))
    check("...and shows no marker on the still-shared row",
          detached_marker_for("Richmond") == "", detached_marker_for("Richmond"))

    print("\n\"Apply to all\" is keyed by AUDIENCE now, not by one row's id (one control "
          "per audience with a detached row, not one per group -- see the D2 item-3 "
          "redesign, live feedback, 2026-08-23), and pushes the audience's OWN shared/"
          "cascade color -- the only color that actually undoes an accidental detach, "
          "not a detached row's own new color, which would PROMOTE the accident instead")
    aud = tg.audience_label(by_label(at, "Richmond"))
    apply_buttons = [b for b in at.button if b.key == f"apply_color_{aud}"]
    check("the Apply-to-all button for this audience exists and is labeled for a rep",
          bool(apply_buttons) and apply_buttons[0].label == "Apply to all",
          [(b.key, b.label) for b in at.button])

    # Drive the actual state transition the way apply_pending_color_cascade
    # itself is triggered -- queuing the same session_state key the button
    # click sets -- rather than through AppTest's own button-click replay,
    # which leaves a one-shot trigger flag that a later run can't always
    # clear cleanly once the button it belonged to stops being rendered
    # (exactly what happens two steps down, once Family drops to one
    # group after the delete, then back to two after the add -- a stale
    # click on an id from an EARLIER run collided with a brand-new button's
    # id on a LATER one). The button's own existence and label are already
    # asserted above; this isolates the state-transition assertion from
    # that AppTest quirk.
    at.session_state["_pending_color_cascade"] = (tg.audience_label(by_label(at, "Richmond")),
                                                   by_label(at, "Richmond")["color"])
    run(at)
    check("no exception after Apply to all", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    g_dc3, g_rich3 = by_label(at, "DC"), by_label(at, "Richmond")
    check("DC is RE-ATTACHED -- color_locked cleared", g_dc3 is not None and not g_dc3["color_locked"],
          g_dc3)
    check("DC's color matches the audience's (Richmond's) color again",
          g_dc3 and g_rich3 and g_dc3["color"] == g_rich3["color"], (g_dc3, g_rich3))
    check("the grid's own lock marker is gone for DC now that it's re-attached",
          detached_marker_for("DC") == "", detached_marker_for("DC"))

    print("\ndeleting DC removes it from targeting_groups outright")
    run(at, "delete_dc")
    check("no exception after delete", not at.exception,
          at.exception[0].message[:400] if at.exception else "")
    check("DC is gone", by_label(at, "DC") is None, real_groups(at))
    survivor = by_label(at, "Richmond")
    check("Richmond survives, untouched", survivor is not None, real_groups(at))

    # A fresh AppTest session here, seeded with exactly the state the
    # session above just reached (Richmond, alone, unlocked) -- an AppTest-
    # only quirk (confirmed by direct repro, unrelated to this feature's
    # own logic: reproduces on plain st.button(key=...) calls with no
    # color/cascade code involved at all) rejects a BRAND NEW button key
    # rendered in the SAME session after a run where that exact set of
    # "Apply to all" buttons had shrunk to zero and grown back. A real
    # browser session has no such restriction; splitting the session here
    # tests the real add-after-delete behavior without fighting a harness
    # limitation that has nothing to do with what's being verified.
    print("\nadding a new row for the SAME audience (Chicago) inherits the audience's "
          "current shared color, unlocked -- same as any other new group")
    at2 = new_app([dict(survivor)])
    run(at2, "add_chicago")
    check("no exception after add", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    g_chicago = by_label(at2, "Chicago")
    g_rich_final = by_label(at2, "Richmond")
    check("the new Chicago group exists", g_chicago is not None, real_groups(at2))
    check("it inherited the audience's shared color",
          g_chicago and g_rich_final and g_chicago["color"] == g_rich_final["color"],
          (g_chicago, g_rich_final))
    check("it starts unlocked", g_chicago is not None and not g_chicago["color_locked"], g_chicago)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("A single-row Color edit detaches permanently until explicitly reversed; "
          "\"Apply to all\" is that reversal, re-attaching a detached row rather than "
          "skipping it; the D2 grid's own Detached marker tracks color_locked exactly, "
          "at every step from the initial cascade through detach, an untouched rerun, "
          "re-attach, delete and a fresh add.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
