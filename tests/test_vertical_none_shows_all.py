"""A vertical narrows the case-study/vault-slide pickers, it doesn't gate
them -- interface audit finding, fixed 2026-09-04.

With Vertical left at its default ("None"), `matching` used to be empty by
construction (nothing is tagged "none"), so every real case study/vault slide
sat behind an always-collapsed "Other ... -- not tagged for this vertical"
expander. A rep working a proposal before picking a vertical -- the ordinary
first-few-minutes state of a real proposal -- saw ZERO visible case studies
despite a populated vault (22 real ones in production). Fixed: with no
vertical, every row renders directly, same as a real "matching" row would,
just none of them pre-checked.

    python tests/test_vertical_none_shows_all.py
"""
import os
import sys
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")
db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)

failures = []


def check(label, ok, detail=""):
    print(f"  {'PASS' if ok else 'FAIL'}  {label}{('  -- ' + str(detail)) if detail else ''}")
    if not ok:
        failures.append(label)


def new_app():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
    at.session_state["authed"] = True
    at.session_state["current_user"] = "T"
    at.session_state["market_choice"] = "DC"
    at.session_state["flight_start"] = date(2026, 9, 1)
    at.session_state["flight_end"] = date(2026, 11, 30)
    at.session_state["avails_mode"] = False  # no real avails table needed for this check
    return at


def main():
    print("=" * 78)
    print("SCENARIO  Vertical left at its default (\"None\") -- case studies and "
          "vault slides show unfiltered, not hidden behind a collapsed expander")
    print("=" * 78)
    at = new_app()
    at.run()
    check("no exception", not at.exception, at.exception[0].message[:400] if at.exception else "")

    real_case_studies, cs_warning = db.fetch_case_studies()
    real_vault_slides, vault_warning = db.fetch_slide_vault()
    if cs_warning or not real_case_studies:
        print(f"\nSKIP -- no case studies reachable ({cs_warning}); can't prove rows are visible")
        skipped_cs = True
    else:
        skipped_cs = False
    if vault_warning or not real_vault_slides:
        print(f"\nSKIP -- no vault slides reachable ({vault_warning}); can't prove rows are visible")
        skipped_vault = True
    else:
        skipped_vault = False

    if not skipped_cs:
        cs_checkboxes = [c for c in at.checkbox if c.key and c.key.startswith("cs_pick_")]
        check(f"every real case study ({len(real_case_studies)}) has a visible, top-level "
              f"checkbox -- none hidden inside a collapsed expander",
              len(cs_checkboxes) == len(real_case_studies),
              (len(cs_checkboxes), len(real_case_studies)))
        check("none of them are pre-checked -- there's no vertical to have matched",
              cs_checkboxes and not any(c.value for c in cs_checkboxes),
              [c.value for c in cs_checkboxes])
        cs_other_expanders = [e for e in at.expander if e.label.startswith("Other case studies")]
        check("no leftover \"Other case studies\" expander -- everything is already shown, "
              "nothing left to bucket separately",
              not cs_other_expanders, [e.label for e in cs_other_expanders])

    if not skipped_vault:
        vault_checkboxes = [c for c in at.checkbox if c.key and c.key.startswith("vault_pick_")]
        check(f"every real vault slide ({len(real_vault_slides)}) has a visible, top-level "
              f"checkbox",
              len(vault_checkboxes) == len(real_vault_slides),
              (len(vault_checkboxes), len(real_vault_slides)))
        vault_other_expanders = [e for e in at.expander if e.label.startswith("Other vault slides")]
        check("no leftover \"Other vault slides\" expander",
              not vault_other_expanders, [e.label for e in vault_other_expanders])

    print("\n" + "=" * 78)
    print("SCENARIO  a REAL vertical still narrows -- matching pre-checked and visible, "
          "off-vertical ones still tucked behind their own \"Other\" expander (unchanged)")
    print("=" * 78)
    at2 = new_app()
    at2.session_state["vertical_choice"] = "Automotive"
    at2.run()
    check("no exception", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    if not skipped_cs:
        matching_now = [r for r in real_case_studies if "auto" in (r.get("verticals") or [])]
        off_vertical_now = [r for r in real_case_studies if "auto" not in (r.get("verticals") or [])]
        cs_other_expanders2 = [e for e in at2.expander if e.label.startswith("Other case studies")]
        if off_vertical_now:
            check("the \"Other\" expander is BACK once a real vertical is set and some case "
                  "studies don't match it -- this is the narrowing behavior, not something this "
                  "fix should have removed",
                  bool(cs_other_expanders2), [e.label for e in at2.expander])
        if matching_now:
            check("at least one matching case study is pre-checked (unchanged pre-check behavior)",
                  any(c.value for c in at2.checkbox
                      if c.key and c.key.startswith("cs_pick_")), None)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Vertical narrows the case-study/vault-slide pickers; it no longer gates them shut "
          "when it's left at its default.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
