"""avails_mode ("Working from an avails document") is authoritative, not
advisory -- 2026-09-08 fix, found on a real WAEPA proposal.

    python tests/test_avails_mode_authoritative.py

Before this fix, `avails_mode` controlled only whether the avails-PDF
uploader was shown -- nothing else read it. THREE independent things
ignored the toggle being off (the third found only after regenerating a
real WAEPA proposal against the first two fixes -- it had real target
markets, which the "dental_single_option" fixture used for (a)/(b) below
does not, so it exercised a path those two checks couldn't reach):

  (a) `apply_draft_to_form` seeded a targeting group from the notes' own
      drafted audiences regardless of avails_mode, with the ORIGINATING
      market as its Geo rather than the target markets.
  (b) `assembly.build_presentation` could include the personalized
      targeting/avails slide even with the toggle off, because "standard"
      (`STANDARD_FORCED_KEYS`) forces `targeting_avails_template` in
      unconditionally -- so the slide rendered with a single zero-avails
      placeholder row. A client must never see "0".
  (c) D2's OWN market-driven blank-row seeding (`avails_rows_for_markets`,
      called whenever D2 rendered at all -- gated only by the SEPARATE
      "Include personalized targeting / avails table" toggle, never by
      avails_mode) created one empty-audience, zero-avails group PER
      TARGET MARKET regardless of avails_mode. D2 itself is now gated on
      `include_avails_template and avails_mode`.

Reuses the committed "dental_single_option" fixture (healthcare vertical,
one drafted audience -- "HLTH Preventative Care Dental" -- with no stated
avails figure, which is exactly the shape that triggers the seeding path)
rather than authoring a new one; this suite is about the avails_mode gate,
not about drafting content, so a live Claude call isn't needed.

Offline. Stubs db.log_proposal/upload_proposal_logo/proposal_logo the same
way test_draft_regression.py does, so nothing here ever writes to
production.
"""
import copy
import json
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
FIXTURES = Path(__file__).resolve().parent / "fixtures"
os.chdir(REPO)

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import slide_map                               # noqa: E402

DRAFT = FIXTURES / "dental_single_option.draft.json"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StubSt:
    """Same technique test_draft_regression.py uses: apply_draft_to_form
    only touches st.session_state, so a plain dict runs it outside a real
    Streamlit runtime."""
    def __init__(self):
        self.session_state = {}
        self.secrets = {}


def apply_draft(draft, preset_state):
    """Returns (session_state, unresolved, internal) -- apply_draft_to_form
    writes the two review lists into session_state itself
    ("draft_unresolved"/"draft_unresolved_internal"), it doesn't return
    them."""
    real = app.st
    stub = _StubSt()
    stub.session_state.update(preset_state)
    app.st = stub
    try:
        app.apply_draft_to_form(copy.deepcopy(draft))
    finally:
        app.st = real
    return (stub.session_state, stub.session_state.get("draft_unresolved", []),
           stub.session_state.get("draft_unresolved_internal", []))


def load_draft():
    raw = json.loads(DRAFT.read_text(encoding="utf-8"))
    raw.pop("_comment", None)
    return raw


def band_preset(draft, avails_mode):
    return {
        "market_choice": draft.get("market") or "DC",
        "flight_start": date.fromisoformat(draft["flight_start"]),
        "flight_end": date.fromisoformat(draft["flight_end"]),
        "avails_mode": avails_mode,
    }


# ---------------------------------------------------------------------------
# (a) apply_draft_to_form: no group is seeded when avails_mode is off
# ---------------------------------------------------------------------------

def check_no_group_seeded_when_off():
    print("\napply_draft_to_form: avails_mode OFF seeds no targeting group")
    draft = load_draft()
    check("fixture genuinely names a real audience (otherwise this proves nothing)",
         bool(draft.get("audiences")), draft.get("audiences"))

    state, _unresolved, internal = apply_draft(draft, band_preset(draft, avails_mode=False))
    groups = state.get("targeting_groups") or []
    real_groups = [g for g in groups if g.get("terms") and not g.get("_placeholder")]
    check("no REAL targeting group exists", not real_groups, groups)
    check("no _placeholder group either (nothing was seeded at all)",
         not groups, groups)
    check("avails_seed_rows was not written",
         "avails_seed_rows" not in state, state.get("avails_seed_rows"))
    check("the audience name is still named, in unresolved_internal (not silently dropped)",
         any("HLTH Preventative Care Dental" in n for n in internal), internal)


def check_group_still_seeded_when_on():
    print("\napply_draft_to_form: avails_mode ON -- unchanged, a group IS seeded "
         "(today's intended behavior, per Matt's own paired test)")
    draft = load_draft()
    state, _unresolved, _internal = apply_draft(draft, band_preset(draft, avails_mode=True))
    groups = state.get("targeting_groups") or []
    real_groups = [g for g in groups if g.get("terms") and not g.get("_placeholder")]
    check("a real targeting group WAS seeded from the drafted audience",
         bool(real_groups), groups)


# ---------------------------------------------------------------------------
# (b) assembly.build_presentation: the slide is dropped when avails_mode off
# ---------------------------------------------------------------------------

def _selections_with(avails_mode, include_avails_template=True):
    sel = copy.deepcopy(assembly.SELECTIONS)
    sel["preset"] = "standard"   # the exact shape that forces the slide in
    sel["vertical"] = "healthcare"
    sel["include_avails_template"] = include_avails_template
    sel["avails_mode"] = avails_mode
    return sel


def _has_targeting_avails_slide(prs):
    dsm = slide_map.build_slide_map_from_prs(prs)
    return any(key in assembly.TARGETING_AVAILS_KEYS for key in dsm.values())


def check_slide_dropped_when_avails_mode_off():
    print("\nbuild_presentation: avails_mode OFF drops targeting_avails_template "
         "even though 'standard' forces it in")
    master = REPO / "TEGNA_MASTER_DECK_v1_1.pptx"
    if not master.exists():
        print(f"  SKIP  {master.name} not present")
        return

    on_sel = _selections_with(avails_mode=True)
    prs_on, _c1, _c2 = assembly.build_presentation(str(master), on_sel)
    check("sanity: with avails_mode ON, 'standard' DOES force the slide in "
         "(otherwise this proves nothing)",
         _has_targeting_avails_slide(prs_on), None)

    off_sel = _selections_with(avails_mode=False)
    prs_off, _c1, _c2 = assembly.build_presentation(str(master), off_sel)
    check("with avails_mode OFF, the slide is gone despite 'standard'",
         not _has_targeting_avails_slide(prs_off), None)


def check_missing_avails_mode_key_defaults_true():
    print("\nbuild_presentation: a selections dict with NO 'avails_mode' key at all "
         "(every caller/test that predates this) is unaffected")
    master = REPO / "TEGNA_MASTER_DECK_v1_1.pptx"
    if not master.exists():
        print(f"  SKIP  {master.name} not present")
        return
    sel = _selections_with(avails_mode=True)
    del sel["avails_mode"]
    prs, _c1, _c2 = assembly.build_presentation(str(master), sel)
    check("the slide still appears (missing key defaults to True, old behavior)",
         _has_targeting_avails_slide(prs), None)


# ---------------------------------------------------------------------------
# (c) D2's own market-driven blank-row seeding, through the REAL form
# ---------------------------------------------------------------------------

def check_no_d2_seeding_when_avails_mode_off():
    print("\nThe real form: picking target markets with avails_mode OFF seeds "
         "no D2 placeholder groups (the WAEPA gap -- needs REAL target markets, "
         "which the dental fixture above doesn't have)")
    from streamlit.testing.v1 import AppTest

    rows, _warning = app.load_market_profiles()
    two = [r for r in rows if r.get("key") in ("denver", "atlanta")]
    check("fixture sanity: two real market rows resolved", len(two) == 2, two)
    if len(two) != 2:
        return
    labels = [app.market_profile_option_label(r) for r in two]

    real_log = db.log_proposal
    db.log_proposal = lambda *a, **k: ("test", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Regression Suite"
        at.session_state["market_choice"] = "DC"
        at.session_state["flight_start"] = date(2026, 10, 1)
        at.session_state["flight_end"] = date(2026, 12, 31)
        at.session_state["avails_mode"] = False
        at.session_state["target_dma_choice"] = labels
        at.run()
        check("no exception with target markets picked and avails_mode off",
             not at.exception, at.exception)
        if at.exception:
            return
        groups = at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []
        check("no targeting groups exist at all (D2 never ran to seed them)",
             not groups, groups)
        d2_headers = [h.value for h in at.header if "D2" in str(h.value)]
        check("the D2 header itself doesn't render", not d2_headers, d2_headers)
    finally:
        db.log_proposal = real_log


def check_d2_still_seeds_when_avails_mode_on():
    print("\nThe real form: the SAME two target markets with avails_mode ON -- "
         "unchanged, D2 seeds one blank row per market (today's intended "
         "behavior, per Matt's own paired test)")
    from streamlit.testing.v1 import AppTest

    rows, _warning = app.load_market_profiles()
    two = [r for r in rows if r.get("key") in ("denver", "atlanta")]
    if len(two) != 2:
        check("fixture sanity: two real market rows resolved", False, two)
        return
    labels = [app.market_profile_option_label(r) for r in two]

    real_log = db.log_proposal
    db.log_proposal = lambda *a, **k: ("test", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Regression Suite"
        at.session_state["market_choice"] = "DC"
        at.session_state["flight_start"] = date(2026, 10, 1)
        at.session_state["flight_end"] = date(2026, 12, 31)
        at.session_state["avails_mode"] = True
        at.session_state["target_dma_choice"] = labels
        at.run()
        check("no exception with target markets picked and avails_mode on",
             not at.exception, at.exception)
        if at.exception:
            return
        groups = at.session_state["targeting_groups"] if "targeting_groups" in at.session_state else []
        check("one blank-audience group per target market IS seeded",
             len(groups) == 2, groups)
        d2_headers = [h.value for h in at.header if "D2" in str(h.value)]
        check("the D2 header renders", bool(d2_headers), d2_headers)
    finally:
        db.log_proposal = real_log


def main():
    check_no_group_seeded_when_off()
    check_group_still_seeded_when_on()
    check_slide_dropped_when_avails_mode_off()
    check_missing_avails_mode_key_defaults_true()
    check_no_d2_seeding_when_avails_mode_off()
    check_d2_still_seeds_when_avails_mode_on()

    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
