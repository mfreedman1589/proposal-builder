"""The personalized cover (client_title_v2, 2026-11-16).

Builds real proposals through the real form and a real Generate, against the
newest local master that carries the v2 cover (TEGNA_MASTER_DECK_*.pptx,
gitignored -- build one with build_master_cover_v2.py), and checks what
PowerPoint actually DRAWS -- never the declared box heights, which a
spAutoFit box doesn't recompute on a generated file:

  * Burch-shaped and MidPoint-shaped plans (the synthetic stand-ins
    tests/test_burch_oil_fixes.py and tests/test_multi_option_draft.py use)
    each produce a cover whose two headline lines, subtitle and footer sit
    inside their own boxes, on the lines they should, without running into
    the next block.
  * The app's own lines: the eyebrow (STREAMING TV / TOTAL TV STRATEGY and
    the flight's first year), places and flight, "Prepared for".
  * A logo lands inside the dashed box's bounds, aspect kept, left/top
    aligned, and the box is gone; with no logo the box is gone and every
    text shape has moved up by the box's height + 0.2in.
  * The old cover never appears beside the new one, and a proposal logged
    without the cover selection (every one before this) still builds on the
    old cover.
  * Notes too thin for a headline: "{Client}" / "Streaming TV Strategy" and
    a Cover row in the review.

    python tests/test_cover_v2.py
"""
import copy
import io
import os
import sys
import tempfile
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")
os.environ.setdefault("PROPOSAL_BUILDER_TODAY", "2026-11-16")

import db                                      # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app                                     # noqa: E402
import assembly                                # noqa: E402
import deck_render                             # noqa: E402
import slide_map                               # noqa: E402
import test_burch_oil_fixes as burch           # noqa: E402
import test_multi_option_draft as midpoint     # noqa: E402

from PIL import Image                          # noqa: E402
from pptx import Presentation                  # noqa: E402
from pptx.util import Inches                   # noqa: E402

failures = []
EMU_PER_PT = 12700


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def v2_master():
    """The newest local master with the v2 cover, or None."""
    for path in sorted(REPO.glob("TEGNA_MASTER_DECK_*.pptx"), key=lambda p: p.stat().st_mtime,
                       reverse=True):
        prs = Presentation(path)
        if any(slide_map.notes_key(s) == assembly.COVER_V2_KEY for s in prs.slides):
            return path
    return None


def logo_png(path):
    """A wide synthetic logo (4:1), so fitting is width-bound."""
    Image.new("RGB", (800, 200), (30, 120, 200)).save(path)
    return path


def generate(state, master, logo=None):
    """A real Generate through the form, on `master`. Returns (prs, form_json,
    error)."""
    from streamlit.testing.v1 import AppTest
    captured = {}
    real = {name: getattr(db, name) for name in ("master_deck", "log_proposal",
                                                 "fetch_case_studies")}
    real_personalize = assembly.personalize

    def spy_personalize(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"], captured["fill"] = prs, fill_data
        return warnings

    def spy_log(client_name, vertical, market, form_json, **kwargs):
        captured["form_json"] = form_json
        return "00000000-0000-0000-0000-000000000000", None

    db.master_deck = lambda fallback: (str(master), None, None)
    db.log_proposal = spy_log
    db.fetch_case_studies = lambda active_only=True: ([], None)
    assembly.personalize = spy_personalize
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=900)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "Scenario Suite"
        for key, value in state.items():
            at.session_state[key] = value
        if logo:
            at.session_state["uploaded_logo"] = {"bytes": Path(logo).read_bytes(),
                                                 "name": Path(logo).name}
        at.run()
        if at.exception:
            return None, None, str(at.exception[0].value)
        button = [b for b in at.button if b.label == "Generate proposal"]
        if not button:
            return None, None, "no Generate button"
        if button[0].disabled:
            return None, None, f"Generate is held: {at.session_state['_generate_blockers']}"
        button[0].click().run()
        if at.exception:
            return None, None, str(at.exception[0].value)
    finally:
        for name, fn in real.items():
            setattr(db, name, fn)
        assembly.personalize = real_personalize
    return captured.get("prs"), captured.get("form_json"), None


def drawn(path):
    """{shape name: {...}} for the first slide, as PowerPoint draws it:
    the box (pt) and the text's bounds and line count."""
    import pythoncom
    import win32com.client
    borrowed = deck_render.powerpoint_is_running()
    pythoncom.CoInitialize()
    out = {}
    powerpoint = presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(str(path), ReadOnly=True, WithWindow=False)
        slide = presentation.Slides(1)
        for index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(index)
            entry = {"left": shape.Left, "top": shape.Top, "width": shape.Width,
                     "height": shape.Height, "type": shape.Type}
            if shape.HasTextFrame and shape.TextFrame.HasText:
                rng = shape.TextFrame.TextRange
                entry.update({"text": rng.Text, "lines": rng.Lines().Count,
                              "paras": rng.Paragraphs().Count,
                              "b_left": rng.BoundLeft, "b_top": rng.BoundTop,
                              "b_right": rng.BoundLeft + rng.BoundWidth,
                              "b_bottom": rng.BoundTop + rng.BoundHeight})
            out[shape.Name] = entry
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        finally:
            if powerpoint is not None and not borrowed:
                powerpoint.Quit()
            pythoncom.CoUninitialize()
    return out


def check_fits(name, geo):
    """Each text block inside its own box, one line per headline paragraph,
    and no block running into the one below it."""
    tol = 1.0   # points
    order = ["COVER_EYEBROW", "COVER_HEADLINE", "COVER_SUBTITLE", "COVER_FOOTER"]
    for shape_name in order:
        g = geo.get(shape_name)
        if not g or "b_right" not in g:
            if shape_name != "COVER_SUBTITLE":
                check(f"{name}: {shape_name} is drawn", False, sorted(geo))
            continue
        check(f"{name}: {shape_name} text stays inside its box's width",
              g["b_right"] <= g["left"] + g["width"] + tol,
              f"text to {g['b_right']:.0f}pt, box to {g['left'] + g['width']:.0f}pt")
        check(f"{name}: {shape_name} text stays inside its box's height",
              g["b_bottom"] <= g["top"] + g["height"] + tol,
              f"text to {g['b_bottom']:.0f}pt, box to {g['top'] + g['height']:.0f}pt")
    head = geo.get("COVER_HEADLINE") or {}
    check(f"{name}: the headline is two lines, one per paragraph",
          head.get("lines") == 2, head.get("lines"))
    foot = geo.get("COVER_FOOTER") or {}
    check(f"{name}: the footer is two lines, one per paragraph",
          foot.get("lines") == 2, foot.get("lines"))
    drawn_blocks = [geo[n] for n in order if geo.get(n, {}).get("b_bottom") is not None]
    for upper, lower in zip(drawn_blocks, drawn_blocks[1:]):
        check(f"{name}: '{upper['text'][:24]}' clears '{lower['text'][:24]}'",
              upper["b_bottom"] <= lower["b_top"] + tol,
              f"{upper['b_bottom']:.0f}pt vs {lower['b_top']:.0f}pt")


def cover_slide_checks(name, prs):
    keys = [slide_map.notes_key(s) for s in prs.slides]
    check(f"{name}: the v2 cover is the first slide", keys[:1] == [assembly.COVER_V2_KEY],
          keys[:3])
    check(f"{name}: the old cover isn't in the deck",
          "client_title" not in keys and not any(str(k).startswith("client_title_cobrand")
                                                 for k in keys))
    text = slide_map.extract_slide_text(prs.slides[0])
    check(f"{name}: no token is left on the cover", "{{" not in text, text[:200])


BURCH_COVER = {"cover_headline_1": "Growing Ridgeway Heating",
               "cover_headline_2": "HVAC Across Southern Maryland",
               "cover_subtitle": "Targeted streaming TV and retargeting reach HVAC intenders in "
                                 "new service areas while holding the home base steady."}
MIDPOINT_COVER = {"cover_headline_1": "Northgate Systems in Market",
                  "cover_headline_2": "Executives Across Four Markets",
                  "cover_subtitle": "Streaming TV leads in Washington, DC and opens Dallas, "
                                    "Phoenix and Atlanta, with retargeting behind every view."}


def midpoint_state():
    draft = copy.deepcopy(midpoint.DRAFT)
    draft.update(MIDPOINT_COVER)
    real = app.st
    stub = midpoint._StubSt()
    stub.session_state.update({
        "targeting_groups": copy.deepcopy(midpoint.groups()), "market_choice": "DC",
        "flight_start": date(2027, 1, 1), "flight_end": date(2027, 6, 30),
        "avails_basis": app.AVAILS_BASIS_MONTHLY,
        "avails_documents": [{"name": "synthetic.pdf", "sha1": "0" * 40, "lines": 4,
                              "rfpid": "RFPID-000002", "total_impressions": 0}],
        "draft_source_notes": midpoint.NOTES, "client_name": midpoint.CLIENT,
        # The pre-Generate checks this plan raises on purpose are midpoint's
        # own subject; ticked here so Generate runs.
    })
    app.st = stub
    try:
        app.apply_draft_to_form(draft, notes=midpoint.NOTES)
    finally:
        app.st = real
    return stub.session_state


def burch_state(cover=BURCH_COVER):
    real_draft = burch.DRAFT
    burch.DRAFT = dict(copy.deepcopy(real_draft), **cover)
    try:
        return burch.apply_draft()
    finally:
        burch.DRAFT = real_draft


def check_payload():
    print("\nThe app's own cover lines")
    rows = [{"Tactic": "Premion Streaming TV", "Geo": "Growth Areas"},
            {"Tactic": "Premion Streaming TV", "Geo": "Home Base"},
            {"Tactic": "Streaming Retargeting - Display", "Geo": "Growth Areas, Home Base"}]
    cover = app.cover_payload("Growing Burch Oil", "HVAC Across Southern Maryland", "One line.",
                              "Burch Oil", [{"rows": rows}], date(2027, 1, 1), date(2027, 12, 31),
                              False, date(2026, 11, 16))
    check("eyebrow: STREAMING TV STRATEGY and the flight's first year",
          cover["eyebrow"] == "STREAMING TV STRATEGY · 2027", cover["eyebrow"])
    check("places then flight, the client's names",
          cover["markets_flight"] == "Growth Areas & Home Base · January–December 2027",
          cover["markets_flight"])
    check("Prepared for the client, in the generation month",
          cover["prepared_for"] == "Prepared for Burch Oil · November 2026", cover["prepared_for"])
    tt = app.cover_payload("", "", "", "Burch Oil", [], date(2026, 10, 1), date(2027, 3, 31), True,
                           date(2026, 9, 2), fallback_labels=["Washington, DC"])
    check("Total TV: TOTAL TV STRATEGY", tt["eyebrow"] == "TOTAL TV STRATEGY · 2026", tt["eyebrow"])
    check("a flight across years names both", tt["markets_flight"]
          == "Washington, DC · October 2026–March 2027", tt["markets_flight"])
    check("blank headlines fall back to the client / Streaming TV Strategy",
          (tt["headline_1"], tt["headline_2"]) == ("Burch Oil", "Streaming TV Strategy"))

    thin = app.drafted_cover({"cover_headline_1": "", "cover_headline_2": ""}, "Burch Oil")
    check("notes too thin: {Client} / Streaming TV Strategy, flagged as assumed",
          (thin["headline_1"], thin["headline_2"], thin["assumed"])
          == ("Burch Oil", "Streaming TV Strategy", True), thin)
    figure = app.drafted_cover({"cover_headline_1": "A", "cover_headline_2": "B",
                                "cover_subtitle": "Reach 300,000 homes for $11,850 a month."}, "X")
    check("a subtitle naming a figure is left blank", figure["subtitle"] == ""
          and figure["dropped_subtitle"], figure)


def check_thin_notes_row():
    print("\nNotes too thin for a headline: an assumption row")
    state = burch_state(cover={"cover_headline_1": "", "cover_headline_2": "", "cover_subtitle": ""})
    check("the cover fields read {Client} / Streaming TV Strategy",
          (state.get("cover_headline_1"), state.get("cover_headline_2"))
          == (burch.CLIENT, "Streaming TV Strategy"),
          (state.get("cover_headline_1"), state.get("cover_headline_2")))
    rows = [i for i in state.get("draft_review_items") or [] if i.get("kind") == "cover"]
    check("a Cover row is in the review", len(rows) == 1 and "Cover" in str(rows[0].get("row")),
          [i.get("kind") for i in state.get("draft_review_items") or []])


def check_legacy(master):
    print("\nA proposal logged before the v2 cover still builds on the old one")
    selections = copy.deepcopy(assembly.SELECTIONS)
    selections.pop("cover", None)
    prs, _, _ = assembly.build_presentation(str(master), selections)
    keys = [slide_map.notes_key(s) for s in prs.slides]
    old_cover = [k for k in keys if k == "client_title" or str(k).startswith("client_title_cobrand")]
    check("no cover selection: an old cover (plain or Total TV's), no v2",
          len(old_cover) == 1 and assembly.COVER_V2_KEY not in keys, keys[:3])
    prs, _, _ = assembly.build_presentation(str(master), dict(selections, cover=assembly.COVER_V2))
    keys = [slide_map.notes_key(s) for s in prs.slides]
    check("with it: the v2 cover only", keys[:1] == [assembly.COVER_V2_KEY]
          and "client_title" not in keys, keys[:3])


def check_scenario(name, state, master, tmp, logo=None):
    print(f"\n{name}{' (with a logo)' if logo else ' (no logo)'}")
    for key in list(state):
        if key.startswith("sanity_ok_"):
            state.pop(key)
    prs, form_json, error = generate(state, master, logo=logo)
    check(f"{name}: Generate runs", error is None, error)
    if error:
        return None
    cover_slide_checks(name, prs)
    stored = (form_json or {}).get("cover") or {}
    check(f"{name}: the cover is logged with the proposal", bool(stored.get("headline_1")), stored)
    path = Path(tmp) / f"{name.replace(' ', '_')}{'_logo' if logo else ''}.pptx"
    prs.save(path)
    geo = drawn(path)
    check_fits(name, geo)
    return prs, geo, path


def main():
    master = v2_master()
    if master is None:
        print("SKIP: no local master carries the v2 cover -- build one with "
              "build_master_cover_v2.py.")
        return
    if not deck_render.renderer_available():
        print("SKIP: PowerPoint isn't available to measure the drawn cover.")
        return
    print(f"Master: {master.name}")
    check_payload()
    check_thin_notes_row()
    check_legacy(master)

    template = Presentation(str(master))
    v2 = next(s for s in template.slides if slide_map.notes_key(s) == assembly.COVER_V2_KEY)
    box = next(s for s in v2.shapes if s.name == "CLIENT_LOGO")
    tops = {s.name: s.top for s in v2.shapes}

    with tempfile.TemporaryDirectory() as tmp:
        logo = logo_png(Path(tmp) / "logo.png")
        burch_run = check_scenario("Burch", burch_state(), master, tmp, logo=logo)
        if burch_run:
            prs, geo, _ = burch_run
            pic = next((s for s in prs.slides[0].shapes if s.name == "CLIENT_LOGO"), None)
            check("logo: a picture, inside the box's bounds", pic is not None
                  and pic.shape_type == 13
                  and pic.left == box.left and pic.top == box.top
                  and pic.width <= box.width and pic.height <= box.height,
                  pic and (pic.left, pic.top, pic.width, pic.height))
            if pic is not None:
                check("logo: aspect kept (4:1)", abs(pic.width / pic.height - 4.0) < 0.02,
                      pic.width / pic.height)
                check("logo: width-bound fills the box's width",
                      abs(pic.width - box.width) <= 2, (pic.width, box.width))
            check("logo: the dashed box is gone",
                  not any(s.name == "CLIENT_LOGO" and s.shape_type != 13
                          for s in prs.slides[0].shapes))
            texts = {s.name: s for s in prs.slides[0].shapes}
            check("logo: the text stays where the design put it",
                  all(texts[n].top == tops[n] for n in assembly.COVER_SHIFTED_SHAPES if n in texts))
        no_logo = check_scenario("Burch", burch_state(), master, tmp)
        if no_logo:
            prs, geo, _ = no_logo
            shapes = {s.name: s for s in prs.slides[0].shapes}
            check("no logo: no CLIENT_LOGO box or picture", "CLIENT_LOGO" not in shapes)
            lift = box.height + Inches(0.2)
            moved = {n: tops[n] - shapes[n].top for n in assembly.COVER_SHIFTED_SHAPES
                     if n in shapes}
            check("no logo: every text shape (and the accent bar) moved up by the box + 0.2in",
                  moved and all(v == lift for v in moved.values()), moved)
            eyebrow = geo.get("COVER_EYEBROW") or {}
            want = (tops["COVER_EYEBROW"] - lift) / EMU_PER_PT
            check("no logo: drawn, the eyebrow sits the box + 0.2in higher",
                  abs(eyebrow.get("top", 0) - want) < 1, (eyebrow.get("top"), want))
        check_scenario("MidPoint", midpoint_state(), master, tmp, logo=logo)

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED:")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("All cover checks passed.")


if __name__ == "__main__":
    main()
