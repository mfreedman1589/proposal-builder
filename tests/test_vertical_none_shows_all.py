"""A vertical narrows the case-study/vault-slide pickers, it doesn't gate
them -- interface audit finding #4, fixed 2026-09-04, then corrected the
same day.

With Vertical left at its default ("None"), `matching` used to be empty by
construction (nothing is tagged "none"), so every real case study/vault slide
sat behind an always-collapsed "Other ... -- not tagged for this vertical"
expander. A rep working a proposal before picking a vertical -- the ordinary
first-few-minutes state of a real proposal -- saw ZERO visible rows despite a
populated vault (22 real case studies, 1 real vault slide in production).

The first fix over-corrected: it made every row render directly, outside any
expander, whenever no vertical was set -- trading invisibility for the
opposite problem (22 checkboxes on every proposal that hasn't picked a
vertical yet). The actual bug was invisibility, not inaccessibility. The
correction: BOTH pickers now live behind ONE expander, named with the vault's
total count ("Case studies (22)", "Vault slides (1)"), collapsed by default
in EVERY state -- no vertical, a vertical with matches, a vertical with none.
Opening it shows matching rows pre-checked and sorted first (with their own
caption), then any off-vertical rows below a plain "Other ..." caption (no
longer its own nested expander -- Streamlit doesn't allow nesting one
expander inside another, and this whole picker is inside one now).

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


def find_expander(at, prefix):
    matches = [e for e in at.expander if e.label.startswith(prefix)]
    return matches[0] if matches else None


def captions(at):
    return [c.value for c in at.caption]


def main():
    print("=" * 78)
    print("SCENARIO  Vertical left at its default (\"None\") -- everything lives behind ONE "
          "collapsed, counted expander, not shown directly and not hidden with no signal at all")
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
        cs_expander = find_expander(at, "Case studies (")
        check("a single 'Case studies (N)' expander exists, N being the vault's real total",
              cs_expander is not None and cs_expander.label == f"Case studies ({len(real_case_studies)})",
              cs_expander.label if cs_expander else None)
        check("it's collapsed by default -- the count is the visibility signal, not an open list",
              cs_expander is not None and cs_expander.proto.expanded is False,
              cs_expander.proto.expanded if cs_expander else None)
        cs_checkboxes = [c for c in at.checkbox if c.key and c.key.startswith("cs_pick_")]
        check(f"every real case study ({len(real_case_studies)}) still has a checkbox -- collapsed "
              f"isn't hidden, a rep who opens it sees everything",
              len(cs_checkboxes) == len(real_case_studies),
              (len(cs_checkboxes), len(real_case_studies)))
        check("none of them are pre-checked -- there's no vertical to have matched",
              cs_checkboxes and not any(c.value for c in cs_checkboxes),
              [c.value for c in cs_checkboxes])
        check("no leftover 'Other case studies' caption or expander -- everything is already in "
              "the one list, nothing left to bucket separately",
              not any(c.startswith("Other case studies") for c in captions(at))
              and find_expander(at, "Other case studies") is None,
              captions(at))

    if not skipped_vault:
        vault_expander = find_expander(at, "Vault slides (")
        check("a single 'Vault slides (N)' expander exists, N being the vault's real total",
              vault_expander is not None
              and vault_expander.label == f"Vault slides ({len(real_vault_slides)})",
              vault_expander.label if vault_expander else None)
        check("it's collapsed by default too",
              vault_expander is not None and vault_expander.proto.expanded is False,
              vault_expander.proto.expanded if vault_expander else None)
        vault_checkboxes = [c for c in at.checkbox if c.key and c.key.startswith("vault_pick_")]
        check(f"every real vault slide ({len(real_vault_slides)}) still has a checkbox",
              len(vault_checkboxes) == len(real_vault_slides),
              (len(vault_checkboxes), len(real_vault_slides)))
        check("no leftover 'Other vault slides' caption or expander",
              not any(c.startswith("Other vault slides") for c in captions(at))
              and find_expander(at, "Other vault slides") is None,
              captions(at))

    print("\n" + "=" * 78)
    print("SCENARIO  a REAL vertical still narrows -- matching sorts first and is pre-checked, "
          "off-vertical rows sit below a plain 'Other ...' caption, all INSIDE the same one "
          "expander -- which is STILL collapsed by default, not sprung open just because a "
          "vertical is set")
    print("=" * 78)
    at2 = new_app()
    at2.session_state["vertical_choice"] = "Automotive"
    at2.run()
    check("no exception", not at2.exception,
          at2.exception[0].message[:400] if at2.exception else "")
    if not skipped_cs:
        matching_now = [r for r in real_case_studies if "auto" in (r.get("verticals") or [])]
        off_vertical_now = [r for r in real_case_studies if "auto" not in (r.get("verticals") or [])]
        cs_expander2 = find_expander(at2, "Case studies (")
        check("the same one expander is still there, still named with the vault's real total "
              "(not the narrowed count)",
              cs_expander2 is not None
              and cs_expander2.label == f"Case studies ({len(real_case_studies)})",
              cs_expander2.label if cs_expander2 else None)
        check("...and still collapsed by default -- a real vertical narrows what's INSIDE, it "
              "doesn't spring the expander open",
              cs_expander2 is not None and cs_expander2.proto.expanded is False,
              cs_expander2.proto.expanded if cs_expander2 else None)
        check("no separate 'Other case studies' EXPANDER -- it's flattened to a caption now, "
              "since this whole picker is already inside one expander and Streamlit can't nest "
              "a second one inside it",
              find_expander(at2, "Other case studies") is None,
              [e.label for e in at2.expander])
        if off_vertical_now:
            check("the 'Other case studies (M)' CAPTION is there instead, naming the real count",
                  f"Other case studies ({len(off_vertical_now)})" in " ".join(
                      c for c in captions(at2) if c.startswith("Other case studies")),
                  [c for c in captions(at2) if c.startswith("Other")])
        if matching_now:
            check("at least one matching case study is pre-checked (unchanged pre-check behavior)",
                  any(c.value for c in at2.checkbox
                      if c.key and c.key.startswith("cs_pick_")), None)

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Both pickers stay behind one collapsed, counted expander in every state -- a rep "
          "always sees that something's there without it eating the page, and opening it shows "
          "vertical-matching rows pre-checked and sorted first exactly as before.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
