"""targeting_map.py + its slide-embedding half in assembly.py (roadmap §E).

    python tests/test_targeting_map.py

No COM, no PowerPoint -- Pillow, python-pptx and AppTest only. This is
deliberately NOT the same guarantee `tests/test_group_scenarios.py --render`
gives (which additionally proves PowerPoint can open the result); it's the
half that doesn't need Windows or an installed PowerPoint.

**Not actually fast, despite an earlier version of this docstring claiming
"an answer in seconds."** Most of the file's own scenarios are (pure Pillow
calls, no Streamlit). The last one -- "end to end through the real form" --
drives a real `AppTest.from_file(app.py)` through a real Generate, which
means the real assembly pipeline (font resolution, table sizing, slide
copying) on every run, and that alone runs 2-4 minutes on this machine.
Confirmed it is NOT a live-Supabase network cost (this file now stubs
`db.fetch_audiences`, the same fix `test_group_scenarios.py` needed, but
timed the AppTest scenario stubbed vs. unstubbed and found no meaningful
difference) -- it's the inherent cost of a real Generate through AppTest,
the same order of magnitude `test_group_scenarios.py`'s own state-level
checks pay for the same reason. Don't budget this file as "tier-1 speed";
budget it like anything else that drives a real Generate.
"""
import copy
import io
import math
import os
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import db                                      # noqa: E402

# Forces the local brochure-PDF fallback for the audience catalog instead of
# a live Supabase round trip -- same fix, same reason, as
# tests/test_group_scenarios.py: the "end to end through the real form"
# scenario below runs AppTest.from_file(app.py), and app.py calls
# load_audience_catalog() at its own module scope. MUST run before that
# AppTest is created.
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import assembly                                # noqa: E402
import geo_resolver                            # noqa: E402
import market_lookup                           # noqa: E402
import targeting_groups as tg                  # noqa: E402
import targeting_map as tm                     # noqa: E402
from PIL import Image                          # noqa: E402

market_lookup.install()
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def _resolved_group(term, zips, color, name=""):
    group = tg.new_group([term], geo_def={"kind": "zips", "zips": zips}, color=color, name=name)
    group["resolved_zips"] = zips
    group["resolved_markets"] = geo_resolver.zips_to_markets(zips).resolved.keys()
    group["resolved_markets"] = sorted(group["resolved_markets"])
    return group


def main():
    print("=" * 78)
    print("SCENARIO  groups_with_zips / render_map (pure, offline)")
    print("=" * 78)

    unresolved = tg.new_group(["Homeowners"])
    resolved = _resolved_group("DEMO Homeowner", ["20005", "20006", "20001"], "#4C78A8")
    check("a group with no resolved_zips is excluded",
          tm.groups_with_zips([unresolved, resolved]) == [resolved], None)
    check("None for an empty group list", tm.render_map([]) is None)
    check("None when nothing is resolved", tm.render_map([unresolved]) is None)

    png = tm.render_map([resolved], width_px=400, height_px=300)
    check("real PNG bytes for a resolved group", bool(png) and png[:8] == b"\x89PNG\r\n\x1a\n", None)
    img = Image.open(io.BytesIO(png))
    check("rendered at the requested size", img.size == (400, 300), img.size)

    print("\nSCENARIO  two groups render in two distinct colors")
    print("=" * 78)
    a = _resolved_group("A", ["20005", "20006"], "#E45756")
    b = _resolved_group("B", ["10001", "10002"], "#54A24B")
    png2 = tm.render_map([a, b], width_px=500, height_px=400)
    img2 = Image.open(io.BytesIO(png2)).convert("RGB")
    colors = {c for _count, c in img2.getcolors(maxcolors=500 * 400)}
    check("group A's own color is actually drawn on the map",
          tm._hex_to_rgb("#E45756") in colors, sorted(colors)[:10])
    check("group B's own color is actually drawn on the map",
          tm._hex_to_rgb("#54A24B") in colors, sorted(colors)[:10])

    print("\nSCENARIO  a rep's group name becomes the legend label")
    print("=" * 78)
    named = _resolved_group("A", ["20005"], "#4C78A8", name="Metro Zip Add-On")
    png3 = tm.render_map([named])
    check("named group renders a PNG (label text itself isn't pixel-checked, "
          "just that naming doesn't break rendering)", bool(png3))

    print("\nSCENARIO  color follows the audience -- two groups sharing one "
          "audience and its color collapse into ONE legend entry, labeled by "
          "the audience, not duplicated or shown as two clusters")
    print("=" * 78)
    # Two Visit Hershey & Harrisburg groups, same audience stack, different
    # named-zip clusters (52 zips vs 34 zips) -- the pre-cascade version of
    # this map gave each its OWN round-robin color and so needed geo_label
    # to tell them apart. Now that color follows the audience, both
    # following that audience's shared color is the ordinary case and they
    # read as one thing, same as the map actually draws them.
    same_audience_a = tg.new_group(
        ["AFIRST Travel Buffs and Sightseers", "DEMO Age A21-44"], op="AND",
        geo_def={"kind": "zips", "zips": ["20005", "20006"]}, color="#E45756")
    same_audience_a["resolved_zips"] = ["20005", "20006"]
    same_audience_b = tg.new_group(
        ["AFIRST Travel Buffs and Sightseers", "DEMO Age A21-44"], op="AND",
        geo_def={"kind": "zips", "zips": ["10001", "10002", "10003"]}, color="#E45756")
    same_audience_b["resolved_zips"] = ["10001", "10002", "10003"]
    check("the two groups really do share one audience (else this proves nothing)",
          tg.audience_label(same_audience_a) == tg.audience_label(same_audience_b), None)

    captured_series = {}
    real_draw_legend = tm._draw_legend

    def spy_draw_legend(draw, series, *args, **kwargs):
        captured_series["series"] = series
        return real_draw_legend(draw, series, *args, **kwargs)

    tm._draw_legend = spy_draw_legend
    try:
        tm.render_map([same_audience_a, same_audience_b])
    finally:
        tm._draw_legend = real_draw_legend
    legend_labels = [label for _color, _coords, label, *_rest in captured_series.get("series", [])]
    check("one legend entry for both groups sharing the audience's color, not two",
          legend_labels == [tg.audience_label(same_audience_a)], legend_labels)

    print("\nSCENARIO  ...but a group a rep broke out (a real, different color) still "
          "reads as its own distinguishable legend entry, not silently folded in")
    print("=" * 78)
    # The real bug this replaces: two groups sharing one audience over
    # different geography, painted in two different colors, rendered
    # IDENTICAL legend text under a naive audience_label -- no way to tell
    # which cluster was which. Bucketing by the color actually painted
    # (not just the color_locked flag) means this is caught even for a
    # group that differs from its audience's reference color without being
    # marked locked.
    broken_out = tg.new_group(
        ["AFIRST Travel Buffs and Sightseers", "DEMO Age A21-44"], op="AND",
        geo_def={"kind": "zips", "zips": ["10001", "10002", "10003"]}, color="#54A24B",
        color_locked=True)
    broken_out["resolved_zips"] = ["10001", "10002", "10003"]
    captured_series.clear()
    tm._draw_legend = spy_draw_legend
    try:
        tm.render_map([same_audience_a, broken_out])
    finally:
        tm._draw_legend = real_draw_legend
    legend_labels = [label for _color, _coords, label, *_rest in captured_series.get("series", [])]
    check("two legend entries were built, not collapsed to one", len(legend_labels) == 2,
          legend_labels)
    # Not a bare "Auto Intenders" anymore -- see the next scenario's own
    # docstring for why a bare entry above a qualified one is ambiguous.
    check("the still-shared entry names ITS OWN market too, not left bare",
          f"{tg.audience_label(same_audience_a)} ({tg.geo_label(same_audience_a)})" in legend_labels,
          legend_labels)
    check("the other names the audience AND the broken-out group's own geo_label",
          f"{tg.audience_label(same_audience_a)} ({tg.geo_label(broken_out)})" in legend_labels,
          legend_labels)

    print("\nSCENARIO  a split audience's REMAINING shared entry names its own "
          "markets too, not just the broken-out one -- a bare entry above a "
          "qualified one can't say whether it includes the broken-out market")
    print("=" * 78)
    dc = tg.new_group(["Auto Intenders"], geo_def={"kind": "text", "label": ""}, name="DC", color="#4C78A8")
    richmond = tg.new_group(["Auto Intenders"], geo_def={"kind": "text", "label": ""},
                            name="Richmond", color="#4C78A8")
    denver = tg.new_group(["Auto Intenders"], geo_def={"kind": "text", "label": ""},
                          name="Denver", color="#E45756", color_locked=True)
    entries = tm.legend_entries([dc, richmond, denver])
    check("three groups collapse to two entries (DC+Richmond shared, Denver broken out)",
          len(entries) == 2, entries)
    shared_entry = next((e for e in entries if e[1] == "#4C78A8"), None)
    check("the shared entry exists and covers both DC and Richmond",
          shared_entry is not None and len(shared_entry[2]) == 2, entries)
    if shared_entry:
        check("its label names BOTH markets, not left as a bare 'Auto Intenders'",
              shared_entry[0] == "Auto Intenders (DC, Richmond)", shared_entry[0])
    broken_entry = next((e for e in entries if e[1] == "#E45756"), None)
    check("Denver's own entry is unaffected by the new qualification rule",
          broken_entry is not None and broken_entry[0] == "Auto Intenders (Denver)",
          broken_entry)

    print("\nSCENARIO  ...but an audience with NO split stays a bare entry -- "
          "qualification only kicks in once there's something to disambiguate")
    print("=" * 78)
    entries_unsplit = tm.legend_entries([dc, richmond])
    check("one entry, still just 'Auto Intenders' -- nothing broken out, nothing to qualify",
          entries_unsplit == [("Auto Intenders", "#4C78A8", [dc, richmond])], entries_unsplit)

    print("\nSCENARIO  a qualified label too wide for the legend truncates "
          "visibly ('+N more'), measured against the real rendered glyphs, "
          "never shrunk in font size and never cut with no indication")
    print("=" * 78)
    font = tm._legend_font()
    long_label = ("Auto Intenders (DC, Richmond, Baltimore, Denver, Chicago, "
                 "Miami, Dallas, Phoenix, Boston, Seattle)")
    check("the unfit label really is wider than a realistic legend column "
          "(else this scenario proves nothing)",
          tm._text_width(font, long_label) > tm._LEGEND_TEXT_MAX_WIDTH, tm._text_width(font, long_label))
    fitted = tm._fit_legend_label(long_label, font)
    check("the fitted label fits the real legend text budget",
          tm._text_width(font, fitted) <= tm._LEGEND_TEXT_MAX_WIDTH,
          (fitted, tm._text_width(font, fitted), tm._LEGEND_TEXT_MAX_WIDTH))
    check("it says how many markets were dropped, not just where it cut off",
          re.search(r"\+\d+ more\)$", fitted) is not None, fitted)
    check("the audience name and at least one real market survive",
          fitted.startswith("Auto Intenders (DC") or fitted.startswith("Auto Intenders (DC,"),
          fitted)
    # A short label that already fits is untouched -- truncation is a fit
    # test against the real font, never a blanket rewrite.
    short = "Auto Intenders (DC, Richmond)"
    check("a label that already fits is returned unchanged",
          tm._fit_legend_label(short, font) == short, tm._fit_legend_label(short, font))
    # A label with no "(...)" list at all (nothing to drop) is left as-is --
    # there's a different kind of truncation this was never asked to invent.
    no_list = "A" * 80
    check("a label with nothing to shrink is left alone, not mangled",
          tm._fit_legend_label(no_list, font) == no_list, tm._fit_legend_label(no_list, font))
    # THE REAL BUG THIS CAUGHT: a qualified label whose parenthetical has
    # only ONE item (a single-group shared bucket, e.g. "2 zips") but is
    # still too wide because the AUDIENCE NAME itself is long ("AFIRST
    # Travel Buffs and Sightseers, DEMO Age A21-44" is 52 characters) used
    # to fall through to the "+N more" fallback with N=1 -- claiming one
    # market was hidden when the single item shown WAS the entire list, an
    # outright fabrication on a document a client signs.
    one_item_long_prefix = ("AFIRST Travel Buffs and Sightseers, DEMO Age A21-44 (2 zips)")
    check("that specific label really is over budget (else this proves nothing)",
          tm._text_width(font, one_item_long_prefix) > tm._LEGEND_TEXT_MAX_WIDTH,
          tm._text_width(font, one_item_long_prefix))
    check("a single-item parenthetical is left alone, not turned into a false '+1 more'",
          tm._fit_legend_label(one_item_long_prefix, font) == one_item_long_prefix,
          tm._fit_legend_label(one_item_long_prefix, font))

    print("\nSCENARIO  THREE OR MORE audiences overlapping the same county is a "
          "different, neutral fill -- a two-color hatch would have to silently "
          "drop whoever's third, which would assert something false")
    print("=" * 78)
    # GENUINELY overlapping, NOT identical -- three overlapping thirds of the
    # same 30-zip county, not the same 10 zips three times. Real avails
    # documents on hand never show this partial-overlap shape (every
    # multi-audience real document resolves to exact identity -- see the
    # identity-collapse scenario below), but it's fully reachable: an
    # imported avail can carry partially-overlapping geography across
    # audiences directly, and a rep can reach it by hand-building groups,
    # editing a group's markets after an import, or combining two avails
    # documents in one proposal. This is that case, built directly since no
    # real document on hand supplies one -- a permanent fixture, not a
    # one-off. (Before the identity-collapse pass existed, this scenario
    # used the SAME 10 zips for all three audiences, which was a real bug
    # waiting to happen: once exact-identity collapse landed, that fixture
    # started asserting the wrong thing -- three IDENTICAL audiences don't
    # overlap, they're the same buy, and collapse to one solid entry. This
    # is the corrected, genuinely-mixed replacement.)
    stl_zips_all = sorted(z for z, fl in geo_resolver._data()["zip_counties"].items()
                          if "29510" in fl)
    aud_a = _resolved_group("Auto Intenders", stl_zips_all[0:15], "#4C78A8")
    aud_b = _resolved_group("Travel Buffs", stl_zips_all[5:20], "#F58518")
    aud_c = _resolved_group("Home Shoppers", stl_zips_all[10:25], "#54A24B")
    check("the three audiences are genuinely NOT identical (else this scenario "
          "proves nothing once the identity-collapse pass runs upstream)",
          len({frozenset(aud_a["resolved_zips"]), frozenset(aud_b["resolved_zips"]),
               frozenset(aud_c["resolved_zips"])}) == 3, None)
    check("legend_entries does NOT collapse these three -- they're overlapping, not identical",
          len(tm.legend_entries([aud_a, aud_b, aud_c])) == 3,
          tm.legend_entries([aud_a, aud_b, aud_c]))
    fills_3way = tm._touched_counties([aud_a, aud_b, aud_c])
    multi_fills = [f for _fips, f in fills_3way if f[0] == "multi"]
    two_way_fills = [f for _fips, f in fills_3way if f[0] == "overlap"]
    check("the shared county is reported as a 3-way 'multi', not a 2-way 'overlap'",
          bool(multi_fills), fills_3way)
    check("NO county renders as a two-way hatch when three audiences actually clear "
          "the threshold -- that would silently drop the third", not two_way_fills, fills_3way)
    if multi_fills:
        _kind, labels = multi_fills[0]
        check("all three real audience names are present, none dropped",
              set(labels) == {"Auto Intenders", "Travel Buffs", "Home Shoppers"}, labels)

    png_3way = tm.render_map([aud_a, aud_b, aud_c], width_px=700, height_px=500)
    check("renders a PNG for the 3-way overlap without raising", bool(png_3way))
    colors_3way = {c for _count, c in Image.open(io.BytesIO(png_3way)).convert("RGB")
                  .getcolors(maxcolors=700 * 500)}
    check("the neutral multi-fill color is actually drawn on the map",
          tm._hex_to_rgb(tm._MULTI_FILL_COLOR) in colors_3way, sorted(colors_3way)[:10])
    # Two of the same three groups (any pair) must still resolve as an
    # ordinary two-way overlap -- the 3+ path is additive, not a regression
    # on the case this whole feature started with.
    fills_2of3 = tm._touched_counties([aud_a, aud_b])
    check("dropping back to two audiences still yields an ordinary two-way overlap",
          any(f[0] == "overlap" for _fips, f in fills_2of3), fills_2of3)

    print("\nSCENARIO  identical geography across audiences collapses to ONE entry -- "
          "a RUNTIME check on the actual resolved zips, not an assumption about what "
          "an avails document contains (see targeting_map._collapse_identical_geography)")
    print("=" * 78)
    ident_zips = stl_zips_all[:12]
    ident_a = _resolved_group("AUTO Make Subaru", ident_zips, "#4C78A8")
    ident_b = _resolved_group("AUTO Make Hyundai", ident_zips, "#FF7F0E")
    ident_c = _resolved_group("AUTO Make Volvo", ident_zips, "#54A24B")
    ident_d = _resolved_group("AUTO Make Genesis Intender", ident_zips, "#FF9DA6")
    collapsed = tm.legend_entries([ident_a, ident_b, ident_c, ident_d])
    check("four audiences with EXACTLY identical resolved zips collapse to one entry",
          len(collapsed) == 1, collapsed)
    if collapsed:
        label, color, members = collapsed[0]
        check("the merged entry keeps all four groups",
              len(members) == 4, len(members))
        check("the merged entry is colored with the FIRST audience's own color, "
              "not a new/neutral one", color == "#4C78A8", color)
        check("the label does NOT enumerate the audiences by name -- naming all four "
              "was measured (a real Annapolis case) at 565px against a 272px budget, "
              "with nothing left to trim that doesn't misrepresent who's targeted",
              not any(name in label for name in
                     ("Subaru", "Hyundai", "Volvo", "Genesis")), label)
        check("the label says it covers every targeted audience",
              "all targeted audiences" in label.lower(), label)
    check("the fill for the shared county is a plain solid, not an overlap/multi hatch -- "
          "four audiences on the SAME geography is one buy, not four competing for credit",
          all(f[0] == "solid" for _fips, f in tm._touched_counties([ident_a, ident_b, ident_c, ident_d])),
          tm._touched_counties([ident_a, ident_b, ident_c, ident_d]))

    print("\nSCENARIO  identity-collapse is selective -- two identical PLUS one genuinely "
          "different audience collapses only the identical pair, the different one stays "
          "its own entry (never merged in just because it's also present)")
    print("=" * 78)
    mixed_third = _resolved_group("Travel Buffs", stl_zips_all[15:27], "#D62728")
    partial_collapse = tm.legend_entries([ident_a, ident_b, mixed_third])
    check("exactly two entries -- the identical pair merged, the different one left alone",
          len(partial_collapse) == 2, partial_collapse)
    if len(partial_collapse) == 2:
        sizes = sorted(len(members) for _l, _c, members in partial_collapse)
        check("one entry has both identical-geography groups, the other has just the "
              "genuinely different one", sizes == [1, 2], sizes)

    print("\nSCENARIO  a single audience is a guaranteed no-op for the identity-collapse "
          "pass -- there's nothing to compare it against (Lawn & Leisure, RFPID-265521, "
          "is exactly this real shape: one group, one audience)")
    print("=" * 78)
    lone = _resolved_group("DEMO Homeowner, HH Income 200K Plus", stl_zips_all[:8], "#4C78A8")
    check("one entry in, one entry out, byte-identical",
          tm.legend_entries([lone]) == [("DEMO Homeowner, HH Income 200K Plus", "#4C78A8", [lone])],
          tm.legend_entries([lone]))

    print("\nSCENARIO  _fit_combined: a genuine two-way overlap keeps BOTH full audience "
          "names even when over budget -- naming the specific pair is the whole point of "
          "an overlap entry, unlike a 3+-way summary")
    print("=" * 78)
    long_a = "AFIRST Travel Buffs and Sightseers, DEMO Age A21-44"
    long_b = "TRAVEL Family, DEMO Age A25 Plus, LIFESTYLE Outdoor Enthusiast"
    font = tm._legend_font()
    joined_raw = f"{long_a} + {long_b}"
    check("this pair really is over budget combined (else this scenario proves nothing)",
          tm._text_width(font, joined_raw) > tm._LEGEND_TEXT_MAX_WIDTH,
          tm._text_width(font, joined_raw))
    fitted_pair = tm._fit_combined([long_a, long_b], font)
    check("both full names survive -- neither is dropped to a '+1 more', even though "
          "the pair is over budget", fitted_pair == joined_raw, fitted_pair)
    overlap_a = _resolved_group(long_a, stl_zips_all[0:20], "#4C78A8")
    overlap_b = _resolved_group(long_b, stl_zips_all[10:30], "#FF7F0E")
    overlap_fill_2 = [f for _fips, f in tm._touched_counties([overlap_a, overlap_b])
                      if f[0] == "overlap"]
    check("the two long-named audiences really do overlap (else this proves nothing)",
          bool(overlap_fill_2), overlap_fill_2)
    png_long_overlap = tm.render_map([overlap_a, overlap_b], width_px=900, height_px=560)
    check("renders without raising even with two long, over-budget names", bool(png_long_overlap))

    print("\nSCENARIO  _fit_combined: a genuine 3+-way multi (non-identical, merely "
          "overlapping) drops to whole names + '+N more' when over budget -- a legitimate "
          "summary, unlike the 2-way case above")
    print("=" * 78)
    three_names = ["AFIRST Travel Buffs and Sightseers", "TRAVEL Family and Outdoor Enthusiasts",
                   "LIFESTYLE Home and Garden Shoppers"]
    joined_three_raw = " + ".join(three_names)
    check("these three really are over budget joined in full (else this scenario proves nothing)",
          tm._text_width(font, joined_three_raw) > tm._LEGEND_TEXT_MAX_WIDTH,
          tm._text_width(font, joined_three_raw))
    fitted_three = tm._fit_combined(three_names, font)
    check("the fitted 3-way label fits the real legend budget",
          tm._text_width(font, fitted_three) <= tm._LEGEND_TEXT_MAX_WIDTH,
          (fitted_three, tm._text_width(font, fitted_three)))
    check("it says how many were dropped, not just where it cut off",
          re.search(r"\+\d+ more$", fitted_three) is not None, fitted_three)
    check("at least one WHOLE audience name survives (never a partial name cut mid-word)",
          any(fitted_three.startswith(name) for name in three_names), fitted_three)
    multi_a = _resolved_group(three_names[0], stl_zips_all[0:15], "#4C78A8")
    multi_b = _resolved_group(three_names[1], stl_zips_all[5:20], "#FF7F0E")
    multi_c = _resolved_group(three_names[2], stl_zips_all[10:25], "#54A24B")
    multi_fill_3 = [f for _fips, f in tm._touched_counties([multi_a, multi_b, multi_c])
                   if f[0] == "multi"]
    check("the three long-named audiences really do produce a 3-way multi "
          "(else this proves nothing)", bool(multi_fill_3), multi_fill_3)
    png_long_multi = tm.render_map([multi_a, multi_b, multi_c], width_px=900, height_px=560)
    check("renders without raising with three long, over-budget names", bool(png_long_multi))

    print("\nSCENARIO  dark-mode hatch fills are exempt from the solid-fill muting rule -- "
          "measured against a real dark render, not just the math, since a hatch already "
          "says 'look here' and doesn't need to sit quietly the way a solid fill does")
    print("=" * 78)
    def _dark_hatch_effective(hexc):
        """The color a hatch stripe actually reads as once composited over
        the real dark slide background -- same math render_map itself
        uses (_dark_hatch_style), applied here so the guard measures what
        a viewer would actually see, not an intermediate value."""
        rgb = tm._hex_to_rgb(hexc)
        lightened, alpha = tm._dark_hatch_style(rgb)
        bg = (15, 18, 30)   # the darker end of the deck's own background gradient
        return tuple(round(bg[i] + (lightened[i] - bg[i]) * alpha / 255) for i in range(3))

    def _lab(rgb):
        r, g, b = [c / 255 for c in rgb]
        lin = lambda c: c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
        r, g, b = lin(r), lin(g), lin(b)
        X = r * .4124 + g * .3576 + b * .1805
        Y = r * .2126 + g * .7152 + b * .0722
        Z = r * .0193 + g * .1192 + b * .9505
        X, Z = X / .95047, Z / 1.08883
        f = lambda t: t ** (1 / 3) if t > 0.008856 else 7.787 * t + 16 / 116
        fx, fy, fz = f(X), f(Y), f(Z)
        return (116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz))

    def _delta_e(rgb_a, rgb_b):
        return math.sqrt(sum((x - y) ** 2 for x, y in zip(_lab(rgb_a), _lab(rgb_b))))

    # Before this fix: both stripe colors of a real overlap were driven to
    # the SAME luminance at 23-25% alpha by _dark_fill_style (the solid-fill
    # muting rule, wrongly reused for hatches too), landing ~6.6 DeltaE apart
    # once composited over the real dark background -- texture, not a
    # legible two-color hatch. That's the real Annapolis pair (blue +
    # mauve -- Annapolis identity-collapse no longer reaches this fill kind
    # at all post-item-6, but this confirms _dark_fill_style ITSELF, still
    # exactly what a solid fill uses, really did produce that regression on
    # a real pair -- not a hypothetical one).
    def _dark_fill_effective(hexc):
        rgb = tm._hex_to_rgb(hexc)
        lightened, alpha = tm._dark_fill_style(rgb)
        bg = (15, 18, 30)
        return tuple(round(bg[i] + (lightened[i] - bg[i]) * alpha / 255) for i in range(3))

    check("confirms the OLD (pre-fix, still what _dark_fill_style itself does for a "
          "solid) muting really did collapse a real blue+mauve pair to near-"
          "indistinguishable once composited over the real dark background",
          _delta_e(_dark_fill_effective("#4C78A8"), _dark_fill_effective("#B279A2")) < 10,
          _delta_e(_dark_fill_effective("#4C78A8"), _dark_fill_effective("#B279A2")))

    # The REALISTIC worst case: among the first 5 palette slots (2-3 audience
    # overlaps are the common real shape), orange+yellow separate the least --
    # measured directly against a real dark render (see the module-level
    # comment on _DARK_HATCH_LIGHTEN/_DARK_HATCH_ALPHA for the full
    # before/after table across the realistic pairs). This is the guard's
    # own asserted number, not inherited from the solid-fill bar
    # (tests/test_targeting_groups.py's >=45): a hatch carries PATTERN
    # information (stripe direction and spacing) a flat fill has none of,
    # so it stays legible at a color separation weaker than a solid needs --
    # but it still needs an asserted floor of its own, not an unbounded
    # "better than before."
    eff_orange, eff_yellow = _dark_hatch_effective("#FF7F0E"), _dark_hatch_effective("#EECA3B")
    achieved_de = _delta_e(eff_orange, eff_yellow)
    check(f"the realistic worst-case hatch pair (orange+yellow) achieves >= 30 DeltaE "
          f"over the real dark background (measured: {achieved_de:.1f}, was 6.6 before "
          f"this fix)", achieved_de >= 30, achieved_de)

    hatch_a = _resolved_group("X", stl_zips_all[0:20], "#FF7F0E")
    hatch_b = _resolved_group("Y", stl_zips_all[10:30], "#EECA3B")
    png_hatch_dark = tm.render_map([hatch_a, hatch_b], width_px=900, height_px=560, dark=True)
    check("renders a real dark hatch without raising", bool(png_hatch_dark))
    hatch_rgba = Image.open(io.BytesIO(png_hatch_dark)).convert("RGBA")
    hatch_bg = Image.new("RGBA", hatch_rgba.size, (15, 18, 30, 255))
    hatch_composite = Image.alpha_composite(hatch_bg, hatch_rgba).convert("RGB")
    composite_colors = {rgb for _count, rgb in
                        hatch_composite.getcolors(maxcolors=hatch_composite.size[0] * hatch_composite.size[1])}
    check("the analytically-computed effective orange is really present in the "
          "composited render, not just in theory",
          eff_orange in composite_colors, eff_orange)
    check("the analytically-computed effective yellow is really present in the "
          "composited render, not just in theory",
          eff_yellow in composite_colors, eff_yellow)

    print("\nSCENARIO  the legend can widen past LEGEND_MAX_WIDTH (deliberate, for a "
          "label that genuinely needs it) but never past what the CANVAS actually "
          "affords -- fit against the real canvas-derived ceiling, not the fixed constant, "
          "which is what let a real single-audience label overflow before this fix "
          "(a real Hershey audience name, 323px, over the old fixed 272px budget, with "
          "nothing to trim since it has no market-list parenthetical)")
    print("=" * 78)
    real_long_label = "AFIRST Travel Buffs and Sightseers, DEMO Age A21-44"
    check("this label really is over the OLD fixed budget (else this proves nothing)",
          tm._text_width(font, real_long_label) > tm._LEGEND_TEXT_MAX_WIDTH,
          tm._text_width(font, real_long_label))
    for width_px in (700, 900, 1200):
        ceiling = max(tm._LEGEND_TEXT_MAX_WIDTH,
                     width_px - tm._MIN_MAP_WIDTH - tm.LEGEND_SWATCH - 3 * tm.LEGEND_PADDING)
        fitted = tm._fit_legend_label(real_long_label, font, ceiling)
        check(f"at width_px={width_px}, the fitted label never exceeds the canvas-derived "
              f"ceiling ({ceiling}px)",
              tm._text_width(font, fitted) <= ceiling, tm._text_width(font, fitted))
        legend_w = tm._legend_width([("#4C78A8", [], fitted)])
        map_w = max(tm._MIN_MAP_WIDTH, width_px - legend_w)
        check(f"at width_px={width_px}, map_w + legend_w never exceeds the canvas "
              f"(map_w={map_w}, legend_w={legend_w}, width_px={width_px})",
              map_w + legend_w <= width_px, (map_w, legend_w, width_px))
    wide_group = _resolved_group("AFIRST Travel Buffs and Sightseers, DEMO Age A21-44",
                                 stl_zips_all[:10], "#4C78A8")
    png_wide = tm.render_map([wide_group], width_px=900, height_px=560)
    check("renders without raising for the real over-budget label", bool(png_wide))

    print("\nSCENARIO  a bad zip (absent from the crosswalk) is skipped, not fatal")
    print("=" * 78)
    dirty = _resolved_group("A", ["20005", "99999"], "#4C78A8")
    png4 = tm.render_map([dirty])
    check("renders fine despite one zip with no known centroid", bool(png4))

    print("\nSCENARIO  state/county outlines give the map geographic context")
    print("=" * 78)
    boundaries = tm._boundaries()
    check("map_boundaries.json.gz is built and loads", bool(boundaries), None)
    if boundaries:
        check("all 50 states + DC + PR are present", len(boundaries["states"]) == 52,
              len(boundaries["states"]))
        check("thousands of counties are present", len(boundaries["counties"]) > 3000,
              len(boundaries["counties"]))
        check("Missouri is keyed by its 2-digit state FIPS", "29" in boundaries["states"],
              boundaries["states"].get("29", {}).get("name"))
        check("St. Louis city is keyed by its 5-digit county GEOID", "29510" in boundaries["counties"],
              boundaries["counties"].get("29510", {}).get("name"))

    stl = _resolved_group("A", ["63101", "63108", "63118"], "#4C78A8")   # St. Louis, MO
    png_stl = tm.render_map([stl], width_px=700, height_px=500)
    check("a St. Louis group still renders with the basemap wired in", bool(png_stl))
    img_stl = Image.open(io.BytesIO(png_stl)).convert("RGB")
    colors_stl = {c for _count, c in img_stl.getcolors(maxcolors=700 * 500)}
    check("a county outline color is actually drawn on the map, not just present in code",
          tm.COUNTY_OUTLINE_COLOR in colors_stl, sorted(colors_stl)[:10])
    check("a state outline color is actually drawn on the map",
          tm.STATE_OUTLINE_COLOR in colors_stl, sorted(colors_stl)[:10])

    print("\nSCENARIO  two DIFFERENT audiences targeting the same county overlap -- "
          "neither fill silently wins, and it gets its own 'A + B' legend row")
    print("=" * 78)
    # _COUNTY_FILL_MIN_COVERAGE is 20% of a county's OWN zips -- St. Louis
    # city (29510) has 30. GENUINELY overlapping, NOT identical -- two
    # overlapping 20-zip slices of the same 30, ten zips shared and five
    # unique on each side, not the same ten zips for both audiences. (This
    # used to BE the same list for both -- once the identity-collapse pass
    # landed, two audiences targeting IDENTICAL geography stopped being an
    # "overlap" at all and correctly collapsed to one solid entry instead,
    # which made the old fixture assert the wrong thing. This is the real,
    # permanent replacement -- see the 3-way scenario above for the fuller
    # reasoning on why a partial-overlap fixture has to be built rather
    # than found in a real document.)
    stl_zips_all = sorted(z for z, fips_list in geo_resolver._data()["zip_counties"].items()
                          if "29510" in fips_list)
    stl_auto = _resolved_group("Auto Intenders", stl_zips_all[0:20], "#4C78A8")
    stl_travel = _resolved_group("Travel Buffs", stl_zips_all[10:30], "#F58518")
    check("the two audiences are genuinely NOT identical (else this scenario "
          "proves nothing once the identity-collapse pass runs upstream)",
          frozenset(stl_auto["resolved_zips"]) != frozenset(stl_travel["resolved_zips"]), None)
    check("legend_entries does NOT collapse these two -- they're overlapping, not identical",
          len(tm.legend_entries([stl_auto, stl_travel])) == 2,
          tm.legend_entries([stl_auto, stl_travel]))
    overlap_fills = [f for _fips, f in tm._touched_counties([stl_auto, stl_travel]) if f[0] == "overlap"]
    check("the shared St. Louis county is reported as an overlap, not credited to just one audience",
          bool(overlap_fills), overlap_fills)
    if overlap_fills:
        _kind, (color_a, color_b), (label_a, label_b) = overlap_fills[0]
        check("the overlap names both real audience labels",
              {label_a, label_b} == {"Auto Intenders", "Travel Buffs"}, (label_a, label_b))
        check("the overlap carries both real colors",
              {color_a, color_b} == {"#4C78A8", "#F58518"}, (color_a, color_b))

    # Two same-audience groups touching the SAME county must NOT report as
    # an overlap with themselves -- they're one entry (see the color-follows-
    # the-audience scenario above), so there is only ever one side to credit.
    # Deliberately reuses stl_auto's OWN zip list, not a separately re-sliced
    # one -- this is testing the SAME-AUDIENCE dedup path (unrelated to,
    # and unaffected by, the identity-collapse pass above).
    stl_auto_2 = _resolved_group("Auto Intenders", stl_auto["resolved_zips"], "#4C78A8")
    self_overlap = [f for _fips, f in tm._touched_counties([stl_auto, stl_auto_2]) if f[0] == "overlap"]
    check("two groups sharing ONE audience never overlap with themselves",
          not self_overlap, self_overlap)

    png_overlap = tm.render_map([stl_auto, stl_travel], width_px=700, height_px=500)
    check("renders a PNG for the overlapping pair without raising", bool(png_overlap))
    png_solo = tm.render_map([stl_auto], width_px=700, height_px=500)
    colors_overlap = {c for _count, c in Image.open(io.BytesIO(png_overlap)).convert("RGB")
                      .getcolors(maxcolors=700 * 500)}
    colors_solo = {c for _count, c in Image.open(io.BytesIO(png_solo)).convert("RGB")
                  .getcolors(maxcolors=700 * 500)}
    check("the overlapping render uses colors a single-audience solid fill of the same county "
          "never would -- the hatch's stripe color adds pixels a flat fill wouldn't have",
          bool(colors_overlap - colors_solo), (len(colors_solo), len(colors_overlap)))

    print("\nSCENARIO  the projection fills the frame instead of letterboxing")
    print("=" * 78)
    # A roughly square point cluster (comparable lat and lon span, after the
    # cosine correction) dropped onto a canvas that's nowhere near square --
    # true distance-scale locks both axes to the same scale, so it fills the
    # short axis and leaves large empty margins on the long one.
    sq_lats, sq_lons = [38.0, 39.0], [-90.0, -88.7]
    width, height = 800, 300

    def x_fill_fraction(max_stretch):
        real = tm._MAX_STRETCH
        tm._MAX_STRETCH = max_stretch
        try:
            project, _frame_bounds = tm._projector(sq_lats, sq_lons, width, height)
        finally:
            tm._MAX_STRETCH = real
        x0, _ = project(sq_lats[0], sq_lons[0])
        x1, _ = project(sq_lats[1], sq_lons[1])
        return abs(x1 - x0) / width

    unstretched = x_fill_fraction(1.0)     # true distance-scale, no bound relief
    bounded = x_fill_fraction(tm._MAX_STRETCH)
    check("the bounded-stretch fit uses more of the short-mismatch axis than "
          "true distance-scale would", bounded > unstretched,
          (unstretched, bounded))
    check("but the stretch itself never exceeds the configured bound",
          bounded <= unstretched * tm._MAX_STRETCH + 1e-9,
          (unstretched, bounded, tm._MAX_STRETCH))

    print("\n" + "=" * 78)
    print("SCENARIO  the map-variant slide swap never drops the avails slide on an "
          "older deck (roadmap section E, item 4)")
    print("=" * 78)
    local_deck = str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx")
    if not Path(local_deck).exists():
        print(f"  SKIP -- no local master deck at {local_deck}")
    else:
        real_build_map = assembly.slide_map.build_slide_map_from_prs

        def strip_map_variant(prs):
            """Simulates a master deck built before this feature shipped --
            no slide carries TARGETING_AVAILS_MAP_KEY at all -- deterministic
            regardless of whether the checked-out deck happens to have the
            real variant slide or not (unlike relying on whatever deck
            db.master_deck actually returns)."""
            deck_map = real_build_map(prs)
            return {n: (k if k != assembly.TARGETING_AVAILS_MAP_KEY else "targeting_avails_template")
                   for n, k in deck_map.items()}

        selections = copy.deepcopy(assembly.SELECTIONS)
        selections["preset"] = "standard"
        selections["include_avails_template"] = True
        selections["targeting_map_present"] = True

        assembly.slide_map.build_slide_map_from_prs = strip_map_variant
        try:
            prs, _, _ = assembly.build_presentation(local_deck, selections)
        finally:
            assembly.slide_map.build_slide_map_from_prs = real_build_map
        deck_map_after = real_build_map(prs)
        keys_present = set(deck_map_after.values())
        check("map_present=True on a deck with NO map-variant slide still keeps "
              "the standard avails/targeting slide (never drops both)",
              "targeting_avails_template" in keys_present, sorted(keys_present))

        # And the real thing: this checkout's own local deck (edited for
        # this feature) actually has the variant, and the sweep picks it.
        selections2 = copy.deepcopy(selections)
        prs2, _, _ = assembly.build_presentation(local_deck, selections2)
        keys_present2 = set(assembly.slide_map.build_slide_map_from_prs(prs2).values())
        check("...and picks the real map variant when the deck actually has one",
              assembly.TARGETING_AVAILS_MAP_KEY in keys_present2, sorted(keys_present2))

    print("\n" + "=" * 78)
    print("SCENARIO  assembly.targeting_map_region / place_targeting_map (python-pptx, no COM)")
    print("=" * 78)
    master_path, _, warning = db.master_deck(str(REPO / "TEGNA_MASTER_DECK_v1_1.pptx"))
    if master_path is None:
        print(f"  SKIP -- {warning}")
    else:
        selections = copy.deepcopy(assembly.SELECTIONS)
        selections["preset"] = "standard"
        selections["market"] = "DC"
        selections["include_avails_template"] = True
        prs, _, _ = assembly.build_presentation(master_path, selections)
        avails_slide = assembly.find_slide_with_marker(prs, "{{AVAILS}}")
        check("the avails slide exists in a standard DC deck", avails_slide is not None)
        if avails_slide is not None:
            table_shape = assembly._find_table_shape(avails_slide)
            before_pics = [s for s in avails_slide.shapes if s.shape_type == 13]

            region = assembly.targeting_map_region(avails_slide)
            check("a region is derived (there's a table and a slide width to bound it)",
                  region is not None, region)
            if region is not None:
                left, top, width, height = region
                check("left edge sits right after the table's own right edge + clearance",
                      left == table_shape.left + table_shape.width + assembly._TABLE_CLEARANCE,
                      (left, table_shape.left + table_shape.width))
                check("top is aligned with the table's own top",
                      top == table_shape.top, (top, table_shape.top))
                slide_width = prs.slide_width
                check("right edge mirrors the table's own left margin",
                      left + width == slide_width - table_shape.left,
                      (left + width, slide_width - table_shape.left))

            check("place_targeting_map(None) adds no picture -- the stock image is untouched",
                  True, None)
            assembly.place_targeting_map(avails_slide, None)
            after_none = [s for s in avails_slide.shapes if s.shape_type == 13]
            check("no picture added when there's nothing to draw",
                  len(after_none) == len(before_pics), (len(before_pics), len(after_none)))

            fake_png = tm.render_map([resolved], width_px=300, height_px=200)
            assembly.place_targeting_map(avails_slide, fake_png)
            after_real = [s for s in avails_slide.shapes if s.shape_type == 13]
            check("exactly one new picture added for a real map",
                  len(after_real) == len(before_pics) + 1, (len(before_pics), len(after_real)))
            if len(after_real) == len(before_pics) + 1 and region is not None:
                new_pic = [s for s in after_real
                          if s._element not in {p._element for p in before_pics}][0]
                check("the new picture is sized and positioned to the derived region exactly",
                      (new_pic.left, new_pic.top, new_pic.width, new_pic.height) == region,
                      ((new_pic.left, new_pic.top, new_pic.width, new_pic.height), region))

    print("\n" + "=" * 78)
    print("SCENARIO  end to end through the real form: a proposal with resolved "
          "zips produces a map on the targeting slide (tier-1 speed, no PowerPoint)")
    print("=" * 78)
    os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    from streamlit.testing.v1 import AppTest

    zips_result = geo_resolver.radius_to_zips(["63101"], 10)   # St. Louis
    plaza_group = tg.new_group(
        ["AUTO Make Ford"], geo_def={"kind": "radius", "centers": ["63101"], "miles": 10},
        avails_monthly=100000, color=tg.assign_color(0))
    plaza_group["resolved_zips"] = zips_result.resolved
    plaza_group["resolved_markets"] = sorted(
        geo_resolver.zips_to_markets(zips_result.resolved).resolved.keys())
    check("this fixture's own group really has resolved zips (else the test proves nothing)",
          bool(plaza_group["resolved_zips"]), plaza_group["resolved_zips"])

    captured = {}
    real_personalize = assembly.personalize

    def spy(prs, fill_data):
        warnings = real_personalize(prs, fill_data)
        captured["prs"] = prs
        captured["fill_data"] = fill_data
        return warnings

    assembly.personalize = spy
    db.log_proposal = lambda *a, **k: ("00000000-0000-0000-0000-000000000000", None)
    try:
        at = AppTest.from_file(str(REPO / "app.py"), default_timeout=600)
        at.session_state["authed"] = True
        at.session_state["current_user"] = "T"
        # FLOW_REWORK_PLAN.md Phase 1: the setup band gates the rest of the
        # page (Generate included) on market+flight -- a truly untouched
        # fresh form no longer reaches it at all. Missed in the sweep
        # (ff0d2e5) that fixed ~20 other AppTest suites for the same gate;
        # found running the full sweep for FLOW_REWORK_PLAN.md Phase 3.
        at.session_state["market_choice"] = "DC"
        from datetime import date
        at.session_state["flight_start"] = date(2026, 9, 1)
        at.session_state["flight_end"] = date(2026, 11, 30)
        at.session_state["include_avails_template"] = True
        at.session_state["targeting_groups"] = [plaza_group]
        at.session_state["premion_streaming_tv"] = True
        at.run()
        check("the form renders", not at.exception,
              at.exception[0].message[:400] if at.exception else "")
        generate = [b for b in at.button if b.label == "Generate proposal"]
        check("Generate button present", bool(generate), [b.label for b in at.button])
        if generate:
            generate[0].click().run()
            check("Generate runs without raising", not at.exception,
                  at.exception[0].message[:600] if at.exception else "")
    finally:
        assembly.personalize = real_personalize

    prs = captured.get("prs")
    check("assembly produced a presentation", prs is not None)
    if prs is not None:
        map_png = captured["fill_data"]["avails"].get("map_png")
        check("map_png was computed for this proposal (resolved zips exist)", bool(map_png))
        avails_slide = assembly.find_slide_with_marker(prs, "{{AVAILS}}")
        # find_slide_with_marker looks for the UNFILLED token, which
        # personalize already consumed -- the slide is the one
        # place_targeting_map actually touched instead, found by locating
        # the table and asking what it measured, the same lookup the deck
        # itself uses to place the map.
        avails_slide = next(
            (s for s in prs.slides if assembly._find_table_shape(s) is not None
             and assembly.targeting_map_region(s) is not None
             and "PRECISION TARGETING" in (assembly.slide_map.extract_slide_text(s) or "").upper()),
            None)
        check("found the targeting/avails slide", avails_slide is not None)
        if avails_slide is not None:
            # A real map to draw means build_presentation swaps in the
            # map-variant slide outright (assembly.py's own comment: "the
            # map variant replaces the standard avails/targeting template
            # outright whenever a targeting map will actually be drawn") --
            # not an overlay on the standard photo template.
            #
            # This used to assert the map-variant slide carries exactly ONE
            # picture total ("no stock photo or slide-level wordmark of its
            # own"). That stopped being true the moment the active master
            # deck's own content changed -- real find, 2026-09-23: Supabase's
            # active deck version was replaced that same day with one whose
            # own notes read "Background change for avails," baking a real
            # full-slide dark-gradient PICTURE onto the map-variant slide
            # where the template used to rely on a plain PowerPoint gradient
            # fill instead. That's a legitimate template edit, not a code
            # defect -- the map-variant slide's whole DESIGN CONTRACT is "no
            # stock photo competing with the map," not "no picture at all,"
            # and a baked-in background is neither. Asserting a total picture
            # COUNT made this test re-fail the instant anyone next edits that
            # background again. Identify the map picture the only way that's
            # actually robust to that: byte-exact match against the SAME
            # map_png this proposal's own fill_data computed -- place_targeting_map
            # embeds it verbatim (python-pptx doesn't re-encode a PNG it's
            # simply handed), so this is the map, whatever else the slide
            # background carries and however many pictures that takes.
            deck_map = assembly.slide_map.build_slide_map_from_prs(prs)
            slide_number = list(prs.slides).index(avails_slide) + 1
            check("the map-variant (no-photo) slide was selected, since a real map is drawn",
                  deck_map.get(slide_number) == assembly.TARGETING_AVAILS_MAP_KEY,
                  deck_map.get(slide_number))
            pics = [s for s in avails_slide.shapes if s.shape_type == 13]
            map_pics = [p for p in pics if p.image.blob == map_png]
            check("exactly one picture on the map-variant slide is byte-identical to the "
                  "map this proposal actually computed -- however many OTHER (background) "
                  "pictures the template itself carries",
                  len(map_pics) == 1, [(p.name, len(p.image.blob)) for p in pics])

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The map draws what targeting groups already resolved, skips what hasn't been "
          "resolved without erroring, its slide footprint is derived from the avails table's "
          "own geometry rather than hand-placed, and a real proposal with resolved zips, "
          "driven through the real form and a real Generate, produces a map on the targeting "
          "slide.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
