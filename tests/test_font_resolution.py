"""Font resolution must never substitute silently.

    python tests/test_font_resolution.py

Offline, no PowerPoint, no API key. The same silent substitution has caused
this project's measurement to be wrong twice -- a hardcoded name->file map
with no Proxima Nova entry, then a theme resolved against the wrong slide
master -- and both times nothing said so. These assertions are about the
saying-so, which is what makes a third one findable.

The font index is faked throughout rather than read off the machine: what
must hold is the CONTRACT (exact / substituted / unavailable, each with a
reason), and a test that only passes on a box with Proxima Nova installed
tests the box.
"""

import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import assembly        # noqa: E402
import text_metrics    # noqa: E402

PASSED = FAILED = 0
_REAL_INDEX = text_metrics._font_index
_REAL_FAMILY_OF = text_metrics._family_of


def check(label, condition, detail=""):
    global PASSED, FAILED
    if condition:
        PASSED += 1
        print(f"    PASS  {label}")
    else:
        FAILED += 1
        print(f"    FAIL  {label}{('  ' + str(detail)) if detail else ''}")


def use_index(mapping):
    """Stand in for the machine's installed fonts."""
    text_metrics.reset_resolution_log()
    text_metrics._font_index = lambda: mapping
    _REAL_FAMILY_OF.cache_clear()
    # _family_of reads the file; with a fake index there is no file, so make
    # it report the key it was given instead.
    text_metrics._family_of = lambda path: path


def restore():
    text_metrics._font_index = _REAL_INDEX
    text_metrics._family_of = _REAL_FAMILY_OF
    text_metrics.reset_resolution_log()


# A family with a bold face, and one without -- which is the real shape of
# the master deck: "Proxima Nova" has a bold, "Proxima Nova Light" does not.
FULL = {
    "proxima nova": "pn-regular",
    "proxima nova bold": "pn-bold",
    "proxima nova light": "pn-light",
    "calibri": "calibri",
    "calibri bold": "calibri-bold",
}


def test_exact():
    print("\n  An installed face resolves exactly and says nothing")
    use_index(FULL)
    entry = text_metrics.resolve_font("Proxima Nova")
    check("regular resolves exact", entry["status"] == "exact", entry)
    entry = text_metrics.resolve_font("Proxima Nova", bold=True)
    check("bold resolves to the bold face", entry["path"] == "pn-bold", entry)
    check("and is exact", entry["status"] == "exact", entry)
    check("no note when everything resolved", text_metrics.measurement_note() is None,
          text_metrics.measurement_note())


def test_bold_without_a_bold_face():
    print("\n  Bold with no bold face falls back, and SAYS so")
    use_index(FULL)
    entry = text_metrics.resolve_font("Proxima Nova Light", bold=True)
    check("falls back to the regular weight", entry["path"] == "pn-light", entry)
    check("reported as a substitution, not a match",
          entry["status"] == "substituted", entry)
    check("carries a reason naming the weight", "regular" in (entry["detail"] or ""),
          entry["detail"])
    note = text_metrics.measurement_note() or ""
    check("the note names the face", "Proxima Nova Light" in note, note)
    check("the note says it was bold", "bold" in note, note)
    # Direction matters and was measured against the renderer rather than
    # assumed (PowerPoint TextRange.BoundWidth, 12pt): "Proxima Nova Light"
    # + bold DRAWS ~2% NARROWER than regular Light -- 43.2 -> 42.2pt for
    # "TACTIC", 138.1 -> 134.8 for "Full Flight Total (3 months)". So the
    # regular-weight fallback OVER-measures, which is the safe direction.
    # Do not "fix" this by promoting bold to a heavier face: Proxima Nova
    # Bold measures 45.1/148.9 on those same strings, well wider than what
    # is drawn, and over-reserving shrinks type for room nothing needs.
    check("fallback is the regular face, not a heavier one",
          entry["path"] != "pn-bold", entry)


def test_unknown_family_substitutes_loudly():
    print("\n  An unknown family substitutes, loudly")
    use_index(FULL)
    entry = text_metrics.resolve_font("Nonexistent Brand Sans")
    check("substituted", entry["status"] == "substituted", entry)
    check("to the substitute face", entry["path"] == "calibri", entry)
    check("reason names the missing font",
          "Nonexistent Brand Sans" in (entry["detail"] or ""), entry["detail"])


def test_nothing_installed_reports_the_estimate():
    print("\n  With no fonts at all, the estimate is reported")
    use_index({})
    entry = text_metrics.resolve_font("Proxima Nova Light")
    check("unavailable", entry["status"] == "unavailable", entry)
    # The estimate is only recorded once something is actually measured.
    text_metrics.wrapped_lines("Dedicated Account Management Team", 120,
                               "Proxima Nova Light", 12)
    note = text_metrics.measurement_note() or ""
    check("the note says text was estimated", "estimated" in note, note)
    check("and warns the band may compress differently",
          "compress" in note, note)
    check("estimate is flagged in the report",
          any(e["status"] == "unavailable" for e in text_metrics.resolution_report()),
          text_metrics.resolution_report())


def test_is_available_stays_quiet():
    print("\n  is_available() reports without recording")
    use_index(FULL)
    text_metrics.is_available("Aptos Black")
    text_metrics.is_available("Proxima Nova")
    check("nothing recorded", text_metrics.resolution_report() == [],
          text_metrics.resolution_report())
    check("reports a missing face", not text_metrics.is_available("Aptos Black"))
    check("reports an installed one", text_metrics.is_available("Proxima Nova"))


def test_declared_fonts_come_from_the_deck():
    print("\n  The font list comes from the deck, not from this code")
    from pptx import Presentation
    master = REPO / "TEGNA_MASTER_DECK_v1_1.pptx"
    if not master.exists():
        print("    SKIP  master deck not present")
        return
    declared = assembly.declared_fonts(Presentation(str(master)))
    check("the deck declares its fonts", len(declared) > 0, declared)
    check("including the brand font",
          any("Proxima Nova" in name for name in declared), declared)
    # Deck-driven means the list survives the brand font changing, so assert
    # against a deck rather than against this code. (An earlier version here
    # grepped assembly.py for "Proxima" and ended in `or True`, which cannot
    # fail -- a check whose expectation comes from the same source as the
    # code under test, in its purest form.)
    check("every declared name came from the deck's own list",
          set(declared) == set(assembly.declared_fonts(Presentation(str(master)))),
          declared)
    # An empty presentation declares nothing rather than raising.
    check("a deck with no embedded fonts declares none",
          assembly.declared_fonts(Presentation()) == [],
          assembly.declared_fonts(Presentation()))


def main():
    print("=" * 70)
    print("Font resolution contract")
    print("=" * 70)
    try:
        test_exact()
        test_bold_without_a_bold_face()
        test_unknown_family_substitutes_loudly()
        test_nothing_installed_reports_the_estimate()
        test_is_available_stays_quiet()
        test_declared_fonts_come_from_the_deck()
    finally:
        restore()
    print("\n" + "=" * 70)
    print(f"{PASSED} passed, {FAILED} failed")
    print("=" * 70)
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
