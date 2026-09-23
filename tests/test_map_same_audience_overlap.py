"""Recoloring one group of an audience must never fabricate a spurious
"overlap" legend entry between that audience's own two halves -- found live
(Easterns, 2026-09-23): recoloring the DC group orange (breaking it out
from Baltimore's shared blue) created a 3rd, "crossover" legend entry where
the DC and Baltimore county footprints border/share territory, even though
both groups are the SAME audience ("AUTO Intenders w/ HHI $50K+"), just
drawn in two colors for legend clarity.

Root cause: `_touched_counties` groups territory by LEGEND ENTRY
(`legend_entries`'s own output), and `legend_entries` deliberately splits
one audience into several entries the moment a group's color diverges from
its audience's shared reference color -- that's correct for the legend
(a rep broke this row out, it deserves its own line), but `_touched_
counties` then treated those two entries exactly like two genuinely
different audiences sharing a county, hatching a fabricated 2-way overlap
between what is really one audience's own two halves.

Fixed: `_touched_counties` now collapses same-audience covering entries
(by `tg.audience_label`) to whichever covers the county more, BEFORE
deciding solid/overlap/multi -- so only a genuine cross-audience collision
ever reaches the hatch fills. A real different-audience overlap (the
existing Auto Intenders / Travel Buffs scenarios in test_targeting_map.py)
is untouched by this change -- confirmed there, not just here.

    python tests/test_map_same_audience_overlap.py
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import geo_resolver                            # noqa: E402
import targeting_groups as tg                  # noqa: E402
import targeting_map as tm                     # noqa: E402

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
    group["resolved_markets"] = sorted(geo_resolver.zips_to_markets(zips).resolved.keys())
    return group


def main():
    # Same real, county-clustered fixture the existing overlap scenarios in
    # test_targeting_map.py use -- two overlapping but non-identical zip
    # slices within one real county (FIPS 29510), so this proves something
    # about a GENUINE partial-overlap geography, not a contrived one.
    zips_all = sorted(z for z, fips_list in geo_resolver._data()["zip_counties"].items()
                      if "29510" in fips_list)
    check("the fixture county has enough zips to slice into an overlap", len(zips_all) >= 20,
         len(zips_all))

    print("SCENARIO  one audience, two colors (a rep recolored one group) -- no fabricated overlap")
    reference_blue = "#4C78A8"
    detached_orange = "#FF7F0E"
    baltimore = _resolved_group("AUTO Intenders", zips_all[0:20], reference_blue, name="Baltimore")
    dc = _resolved_group("AUTO Intenders", zips_all[10:30], detached_orange, name="DC")
    dc["color_locked"] = True  # a rep's own single-row override, same as the D2 grid produces

    entries = tm.legend_entries([baltimore, dc])
    check("legend_entries genuinely produces two separate entries for this audience "
         "(else this scenario proves nothing about the fix)",
         len(entries) == 2, entries)
    labels = [e[0] for e in entries]
    check("one entry is qualified 'DC', the other 'Baltimore'",
         any("DC" in l for l in labels) and any("Baltimore" in l for l in labels), labels)

    fills = tm._touched_counties([baltimore, dc])
    overlap_or_multi = [f for _fips, f in fills if f[0] in ("overlap", "multi")]
    check("NO county renders as an overlap/multi hatch between the two halves of one audience "
         "(this is the reported bug if it fails)",
         not overlap_or_multi, overlap_or_multi)
    check("the shared county still renders -- solid, in whichever half covers it more",
         all(f[0] == "solid" for _fips, f in fills), fills)

    png = tm.render_map([baltimore, dc], width_px=900, height_px=560)
    check("renders a PNG without raising", bool(png))

    print("\nSCENARIO  regression check -- a genuine two-DIFFERENT-audience overlap is untouched")
    real_second_audience = _resolved_group("Travel Buffs", zips_all[5:25], "#F58518")
    fills2 = tm._touched_counties([baltimore, real_second_audience])
    check("two genuinely different audiences sharing a county still hatch as an overlap",
         any(f[0] == "overlap" for _fips, f in fills2), fills2)

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
