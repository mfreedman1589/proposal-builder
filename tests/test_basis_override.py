"""FLOW_REWORK_PLAN.md Phase 3, commit 1: a per-section Monthly/Full-flight
basis override for the D2 avails table, independent of the setup band's own
Plan basis -- safe because Phase 2 made dates authoritative (the two views
are two renderings of one underlying set of numbers, not two independent
sets).

`avails_section_basis` (None = follow the band, the default) changes what
THIS grid displays, its column header, its editor key and its own
edit/fold-back conversion -- and nothing else. In particular:
  - the underlying stored `avails_monthly` on the group never changes just
    because the section's own display basis changed;
  - the deck/targeting-slide payload (and `form_json`) keep reading the
    BAND's own basis, never this section's override -- a client-facing deck
    reflects the plan's stated basis, not a rep's momentary "let me peek at
    this the other way" toggle.

    python tests/test_basis_override.py
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)

import streamlit as st  # noqa: E402

# Captures whatever DataFrame the D2 avails grid was handed, keyed by its own
# widget key (which itself encodes the effective basis + sort choice -- see
# app.py's `avails_editor_key`) -- pure pass-through otherwise, the same
# no-op behavior AppTest's own un-monkeypatched st.data_editor already has.
_captured_frames = {}


def _capturing_data_editor(data, *args, **kwargs):
    key = str(kwargs.get("key", ""))
    if key.startswith("avails_editor"):
        _captured_frames[key] = data.copy()
    return data


st.data_editor = _capturing_data_editor

import app  # noqa: E402
import targeting_groups as tg  # noqa: E402

sys.path.insert(0, str(REPO / "tests"))
from test_form_state import load_drafted_form  # noqa: E402

failures = []


def sget(at, key, default=None):
    """AppTest's session_state has no working .get() (it treats "get" as a
    key lookup itself) -- this is the safe substitute every call site here
    uses instead."""
    try:
        return at.session_state[key]
    except (KeyError, AttributeError):
        return default


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    return at


def seed_known_state(at, section_basis):
    """A drafted baseline (so Generate has everything else it needs), with
    the avails table replaced by ONE group at a known avails_monthly figure
    -- easy to hand-check against both bases. Everything here is set before
    this instance's first `.run()`, matching test_share_of_voice.py's own
    discipline for widget-shadowing safety.
    """
    load_drafted_form(at)

    group = tg.new_group(["Homeowners"], avails_monthly=6000, group_id="g1")
    rows = tg.groups_to_seed_rows([group], app.AVAILS_COLUMN_MONTHLY)
    at.session_state["targeting_groups"] = [group]
    at.session_state["avails_seed_rows"] = rows
    at.session_state["_groups_rows_applied"] = [dict(r) for r in rows]
    # The band's own basis stays Monthly regardless of what the fixture drafted
    # -- this test is specifically about the section overriding AWAY from the
    # band, so the band must start at the default to make that visible.
    at.session_state["avails_basis"] = app.AVAILS_BASIS_MONTHLY
    if section_basis is not None:
        at.session_state["avails_section_basis"] = section_basis
    return group


def generate(at):
    real_personalize = app.assembly.personalize
    captured = {}

    def spy_personalize(prs, fill_data):
        captured["fill_data"] = fill_data
        return real_personalize(prs, fill_data)

    app.assembly.personalize = spy_personalize
    try:
        buttons = [b for b in at.button if b.label == "Generate proposal"]
        if not buttons:
            return None
        buttons[0].click().run()
    finally:
        app.assembly.personalize = real_personalize
    return captured.get("fill_data")


def avails_editor_frame(at):
    matches = [df for key, df in _captured_frames.items() if key.startswith("avails_editor")]
    return matches[-1] if matches else None


def avails_column(df):
    # FLOW_REWORK_PLAN.md Phase 3 added a "Geo Label" column alongside the
    # entity "Label" one -- both excluded here, same reason
    # test_avails_grid_manual_row.py's own fix documents.
    for col in df.columns:
        if col not in ("gid", "Plan", "Audience", "Markets", "Label", "Geo Label", "Color",
                       "Detached", "Avail dates"):
            return col
    return None


def main():
    print("baseline: section basis left at default (None, follow the band -- Monthly)")
    _captured_frames.clear()
    at = new_app()
    group = seed_known_state(at, section_basis=None)
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")

    df = avails_editor_frame(at)
    check("the grid was rendered", df is not None)
    if df is not None:
        col = avails_column(df)
        check("column header is the plain monthly label when following the band",
              col == app.AVAILS_COLUMN_MONTHLY, col)
        check("displayed figure equals the stored monthly value",
              int(df.iloc[0][col]) == 6000, df.iloc[0].to_dict() if len(df) else None)

    print("\nswitching ONLY this section's basis to Full flight -- the band itself is untouched")
    n_months = max(1, len(sget(at, "active_months") or []))
    _captured_frames.clear()
    at2 = new_app()
    seed_known_state(at2, section_basis=app.AVAILS_BASIS_FLIGHT)
    at2.run()
    check("no exception", not at2.exception, at2.exception[0].message[:400] if at2.exception else "")

    df2 = avails_editor_frame(at2)
    expected_label = app.avails_column_label(app.AVAILS_BASIS_FLIGHT, n_months)
    if df2 is not None:
        col2 = avails_column(df2)
        check("column header names Full Flight and the month count",
              col2 == expected_label, (col2, expected_label))
        check("displayed figure is the monthly figure multiplied out over the flight",
              int(df2.iloc[0][col2]) == 6000 * n_months,
              df2.iloc[0].to_dict() if len(df2) else None)

    stored_groups = sget(at2, "targeting_groups") or []
    check("the stored avails_monthly is untouched by the section's own display basis",
          stored_groups and stored_groups[0].get("avails_monthly") == 6000, stored_groups)

    print("\nband stays Monthly regardless of the section override "
          "(band radio's own session_state key)")
    check("band's own avails_basis session key is still Monthly",
          sget(at2, "avails_basis") == app.AVAILS_BASIS_MONTHLY,
          sget(at2, "avails_basis"))

    print("\nthe deck/targeting-slide payload and form_json read the BAND's basis, "
          "never this section's override")
    fill_data = generate(at2)
    check("Generate produced a captured fill_data payload", fill_data is not None)
    if fill_data:
        avails_payload = fill_data.get("avails") or {}
        check("payload's avails LABEL is the band's plain monthly label, not Full Flight",
              avails_payload.get("label") == app.AVAILS_COLUMN_MONTHLY,
              avails_payload.get("label"))
        payload_rows = avails_payload.get("rows") or []
        check("payload's avails FIGURE is the band's (monthly) figure, not the section's "
              "full-flight one",
              payload_rows and payload_rows[0].get("avails") == "6,000",
              payload_rows)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The D2 section's own Monthly/Full-flight override changes only this grid's own "
          "display, header and editor key -- the stored avails_monthly, the band's own basis, "
          "and the deck/targeting-slide payload all stay exactly what the band says.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
