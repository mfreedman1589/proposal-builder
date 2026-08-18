"""The Phase 4 D2 grid silently upgrades a migrated group's `geo_def` from
`kind:"text"` to `kind:"markets"` the moment it's viewed (any group whose
Markets cell renders non-empty gets converted on the very next fold-back --
see the comment on that conversion in app.py's D2 block). This file confirms
that upgrade never changes what a proposal LOOKS LIKE, at either of the two
places that matters:

    python tests/test_group_geo_def_compat.py

1. What gets LOGGED when a pre-groups proposal is loaded, D2 renders (the
   upgrade fires), and the rep generates again -- `form_json["avails_rows"]`
   must be byte-identical to what the OLD, pre-Phase-4 code would have
   written from the same flat rows.
2. "Rebuild as presented" (`app.rebuild_proposal_deck`) -- which reads
   `form_json` directly off a stored DB row and never touches
   `st.session_state` or `targeting_groups` at all (confirmed by reading the
   function: zero `st.session_state` references in its body) -- produces the
   identical `fill_data["avails"]` whether the stored row's `avails_rows`
   came from the pre-upgrade or the post-upgrade path. Proven by actually
   running the rebuild twice, not by asserting from the source read alone.

A combined, comma-containing multi-market Geo ("Denver, Atlanta, Phoenix")
is the deliberately adversarial case: it's the one where `kind:"text"`'s
single free-text label and `kind:"markets"`'s list-of-three both have to
render back to the exact same string through `tg.geo_label` for this
property to hold.
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
          "goes through the live D2 upgrade on generate")
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
    check("the upgrade actually fired -- geo_def is now kind:markets, not kind:text",
          real_group is not None and real_group["geo_def"].get("kind") == "markets",
          real_group)
    # _group_markets' fallback for a migrated kind:text group wraps the WHOLE
    # label as ONE chip -- it never splits on a comma a rep (or the market
    # autofill) happened to join with, same "never guess at free text"
    # discipline seed_rows_to_groups already uses for Audience. So the
    # upgrade produces ONE market entry that still CONTAINS the commas, not
    # three separate ones -- which is exactly what makes tg.geo_label render
    # it back to the identical string below.
    check("...as ONE chip holding the whole joined string, not split into three",
          real_group is not None
          and real_group["geo_def"].get("markets") == ["Denver, Atlanta, Phoenix"],
          real_group)

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
