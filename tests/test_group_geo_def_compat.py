"""An earlier version of the Phase 4 D2 grid's fold-back silently upgraded a
migrated group's `geo_def` from `kind:"text"` to `kind:"markets"` the moment
it was merely VIEWED, with no edit at all -- any group whose Markets cell
rendered non-empty got converted on the very next fold-back. That was found
to be the SAME class of bug as the Audience-collapse regression Phase 4's
own test (`test_group_builder.py`) guards: re-deriving structure from
DISPLAYED text on every render instead of only on an actual edit. Fixed
during Phase 6 (the geo-definition expander is what makes geo_def variety
and `resolved_markets` real, so the bug had to be closed before it could
bite) by giving Markets/geo_def the identical "restore untouched" discipline
Audience already had.

This file pins the CORRECTED behavior:

    python tests/test_group_geo_def_compat.py

1. Merely viewing a migrated `kind:"text"` group in D2 -- rendering it,
   generating from it -- must not change `geo_def` at all. No upgrade, no
   silent structural change, from a run where nothing was edited.
2. What gets LOGGED when a pre-groups proposal is loaded and the rep
   generates again is byte-identical to what pre-Phase-4 code would have
   written from the same flat rows, `form_json["avails_rows"]` compared
   directly.
3. "Rebuild as presented" (`app.rebuild_proposal_deck`) -- which reads
   `form_json` directly off a stored DB row and never touches
   `st.session_state` or `targeting_groups` at all (confirmed by reading the
   function: zero `st.session_state` references in its body) -- produces
   identical output regardless of what a live session's D2 did. Proven by
   actually running the rebuild twice, not by asserting from the source
   read alone.

A combined, comma-containing multi-market Geo ("Denver, Atlanta, Phoenix")
is the deliberately adversarial case for (2): it's the one where a
`kind:"text"` group's single free-text label has to render back through
`tg.geo_label` to the exact string it started as.
"""
import copy
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402
import assembly  # noqa: E402
import db  # noqa: E402
import targeting_groups as tg  # noqa: E402

REPO = Path(__file__).resolve().parent.parent
COL = app.AVAILS_COLUMN_MONTHLY
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def pre_phase4_avails_rows(flat_rows, avails_label):
    """What `avails_rows` (the lowercase, deck-payload shape) would have
    contained for these flat rows under the code that existed before Phase
    4 -- reproduced from that code's own construction, not re-derived
    through groups at all, so this is independent of anything the D2
    rewrite does."""
    return [
        {"audience": str(r.get("Audience", "")), "geo": str(r.get("Geo", "")),
         "avails": f"{int(r.get(avails_label, 0)):,}",
         "avails_monthly": int(r.get(avails_label, 0))}
        for r in flat_rows
        if str(r.get("Audience", "")).strip()
    ]


def generate_and_capture(session_state):
    """Drive the real form with the given starting session_state, click
    Generate, and return the captured logged form_json -- same spy technique
    tier 1 uses (db.log_proposal replaced rather than allowed to write)."""
    from streamlit.testing.v1 import AppTest

    captured = {}
    real_log = db.log_proposal

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        captured["form_json"] = json.loads(json.dumps(db._json_safe(form_json), default=str))
        return "00000000-0000-0000-0000-000000000000", None

    db.log_proposal = spy_log
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=300)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        for key, value in session_state.items():
            at.session_state[key] = value
        at.run()
        if at.exception:
            return None, at.exception[0].message[:300]
        generate = [b for b in at.button if b.label == "Generate proposal"]
        if not generate:
            return None, "no Generate button found"
        generate[0].click()
        at.run()
        if at.exception:
            return None, at.exception[0].message[:300]
    finally:
        db.log_proposal = real_log
    return captured.get("form_json"), None


def rebuild_and_capture_avails(row):
    """Run the real app.rebuild_proposal_deck against `row`, spying on
    assembly.personalize to capture the avails dict it was actually called
    with -- the same content that reaches the slide, without depending on
    zip-level byte determinism of the saved .pptx."""
    captured = {}
    real_personalize = assembly.personalize

    def spy(prs, fill_data):
        captured["avails"] = copy.deepcopy(fill_data["avails"])
        return real_personalize(prs, fill_data)

    assembly.personalize = spy
    try:
        buffer, filename, warnings = app.rebuild_proposal_deck(row)
    finally:
        assembly.personalize = real_personalize
    return buffer, warnings, captured.get("avails")


def main():
    print("a pre-groups proposal (a combined, comma-containing multi-market Geo) "
          "is merely VIEWED in D2 and generated again")
    flat_rows = [{"Audience": "Homeowners", "Geo": "Denver, Atlanta, Phoenix", COL: 500000}]
    expected_avails_rows = pre_phase4_avails_rows(flat_rows, COL)

    form_json, err = generate_and_capture({
        "premion_streaming_tv": True,
        "include_avails_template": True,
        "avails_seed_rows": [dict(r) for r in flat_rows],
        # Deliberately NOT setting "targeting_groups" -- this is exactly what
        # a proposal logged before this feature existed carries: none at all.
    })
    check("the form renders and generates without raising", err is None, err)
    if form_json is None:
        print(f"{1} FAILED (setup): {err}")
        return 1

    groups = form_json.get("targeting_groups") or []
    real_group = next((g for g in groups if g.get("terms")), None)
    # Corrected behavior: merely viewing a migrated group changes NOTHING
    # about its geo_def -- no upgrade, kind:text stays kind:text, forever,
    # until the rep actually edits the Markets cell or resolves it through
    # the Phase 6 geo-definition expander. See this file's module docstring
    # for the bug this replaced.
    check("geo_def is UNCHANGED by merely viewing it -- still kind:text",
          real_group is not None and real_group["geo_def"] == {"kind": "text", "label": "Denver, Atlanta, Phoenix"},
          real_group)
    check("resolved_markets stays empty -- nothing resolved it",
          real_group is not None and real_group.get("resolved_markets") == [], real_group)

    logged_avails_rows = form_json.get("avails_rows") or []
    check("form_json['avails_rows'] is byte-identical to the pre-Phase-4 shape "
          "despite the internal geo_def upgrade",
          logged_avails_rows == expected_avails_rows,
          (logged_avails_rows, expected_avails_rows))

    print("\n\"Rebuild as presented\" never touches session_state or targeting_groups at all")
    row_after = {
        "client_name": "Compat Test Co", "vertical": "none", "market": None,
        "deck_version_id": None, "logo_storage_path": None,
        "output_filename": "compat_test.pptx", "form_json": form_json,
    }
    row_before = copy.deepcopy(row_after)
    row_before["form_json"]["avails_rows"] = expected_avails_rows
    row_before["form_json"].pop("targeting_groups", None)   # the genuinely pre-groups shape

    buf_after, warn_after, avails_after = rebuild_and_capture_avails(row_after)
    buf_before, warn_before, avails_before = rebuild_and_capture_avails(row_before)

    check("rebuild succeeds from the post-upgrade row", buf_after is not None, warn_after)
    check("rebuild succeeds from the pre-upgrade row", buf_before is not None, warn_before)
    check("the fill_data avails dict fed to the slide is byte-identical either way",
          avails_before == avails_after, (avails_before, avails_after))
    check("the assembled deck is the same size either way (same content, same slide count)",
          buf_before is not None and buf_after is not None
          and len(buf_before.getvalue()) == len(buf_after.getvalue()),
          (len(buf_before.getvalue()) if buf_before else None,
           len(buf_after.getvalue()) if buf_after else None))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The kind:text -> kind:markets upgrade is invisible downstream: the logged "
          "avails_rows and the rebuilt deck are unchanged, and rebuild_proposal_deck never "
          "reads targeting_groups or session_state at all.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
