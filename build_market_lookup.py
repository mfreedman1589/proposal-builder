"""Build the county -> market (DMA) lookup the geo resolver plugs in.

    python build_market_lookup.py [--out market_lookup.json.gz] [--report]

DMA definitions are Nielsen's intellectual property and there is no current,
authoritative, freely-licensed county->DMA table. What exists publicly is
secondary redistributions of unknown exact vintage. This script builds a
WORKING table out of the best of them, validates it hard, and records where
every part came from -- so replacing it with a licensed table later is a data
change (one file, one provenance block), not an archaeology exercise.

Three public sources, none of them trusted on its own:

  * BritCrit/dma_county_zip -- county FIPS + DMA code + DMA name + zip.
    The assignment source. Carries Nielsen DMA CODES, which is what makes the
    naming join exact rather than fuzzy.
  * fissehab/Nielsen-Media-Research-DMA (DMA_Names.csv) -- DMA code -> name
    for all 210, in Nielsen's own abbreviated spelling. A SECOND, independent
    naming of every code, used to confirm the code -> canonical-name bridge.
  * alex-patton/US-TVDMA-BY-COUNTY -- county name + state -> DMA name. Much
    older (it predates Broomfield County CO, created 2001), so it is not used
    for assignment at all. It is the INDEPENDENT CROSS-CHECK: two unrelated
    compilations agreeing county by county is evidence neither a bad row nor a
    mis-paired name could produce.

Everything is checked against things this script cannot move: the 210 canonical
DMA names the app already knows (market_profiles.CANONICAL_DMAS), and the 3,222
counties in geo_crosswalk.json.gz, which is built from Census files.

Counties whose geography postdates the assignment source -- Connecticut's
planning regions, the Alaska borough splits, Oglala Lakota -- are filled from
THEIR OWN ZIPS' assignments rather than by hand or by lineage guesswork, and
the vote is reported. A county with no usable zips stays unassigned and is
listed; the resolver reports an unassigned county rather than inventing a DMA.
"""
import argparse
import csv
import gzip
import io
import json
import re
import sys
import urllib.request
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import market_profiles

HERE = Path(__file__).resolve().parent
CACHE = HERE.parent / "proposal-builder-assets" / "geo"
DEFAULT_OUT = HERE / "market_lookup.json.gz"
CROSSWALK = HERE / "geo_crosswalk.json.gz"

SOURCES = {
    "county_dma_zip": {
        "name": "BritCrit/dma_county_zip",
        "url": "https://raw.githubusercontent.com/BritCrit/dma_county_zip/main/"
               "dma_county_zip_data_set.csv",
        "repo": "https://github.com/BritCrit/dma_county_zip",
        "role": "county FIPS -> DMA code/name, and zip -> DMA code",
        "file": "britcrit_dma_county_zip.csv",
    },
    "dma_names": {
        "name": "fissehab/Nielsen-Media-Research-DMA (DMA_Names.csv)",
        "url": "https://raw.githubusercontent.com/fissehab/"
               "Nielsen-Media-Research-DMA/master/DMA_Names.csv",
        "repo": "https://github.com/fissehab/Nielsen-Media-Research-DMA",
        "role": "DMA code -> Nielsen abbreviated name, second naming of all 210",
        "file": "fissehab_dma_names.csv",
    },
    "crosscheck": {
        "name": "alex-patton/US-TVDMA-BY-COUNTY (usa-tvdma-county.csv)",
        "url": "https://raw.githubusercontent.com/alex-patton/US-TVDMA-BY-COUNTY/"
               "master/usa-tvdma-county.csv",
        "repo": "https://github.com/alex-patton/US-TVDMA-BY-COUNTY",
        "role": "independent county -> DMA name, used only to cross-check",
        "file": "alex_patton_usa_tvdma_county.csv",
    },
}

# Territories have no DMA. Nielsen measures the 50 states and DC; Puerto Rico,
# the Virgin Islands, Guam, American Samoa and the Northern Marianas are simply
# outside the 210, so their counties are expected to be unassigned and are not
# reported as a gap.
TERRITORY_PREFIXES = ("60", "66", "69", "72", "78")

# The assignment source's own bucket for counties in no DMA at all.
NON_DMA_CODES = {"0", "000"}

# Markets a county-keyed table cannot express, with the reason. Palm Springs is
# carved out of Riverside County CA, which the assignment source gives whole to
# Los Angeles, and it has no rows of its own anywhere in that source -- so this
# lookup will never return it. Named here rather than left as a silent hole,
# because "the app knows 210 markets and this resolves 209" is exactly the kind
# of gap that otherwise gets discovered by a rep whose zip list comes back empty.
KNOWN_MARKET_GAPS = {
    "Palm Springs": "sub-county DMA; the assignment source has no rows for code 804",
}

# Counties the resolver can never produce, so not a gap in this table. The 2024
# county gazetteer lists Connecticut's nine planning regions while the 2020 zip
# relationship file still uses the eight legacy counties, so zips_to_counties
# only ever returns the legacy FIPS (which ARE assigned here). Fixing the
# mismatch belongs to build_geo_crosswalk.py, not to this script.
UNREACHABLE_PREFIXES = ("091",)


def fetch(key, refresh=False):
    src = SOURCES[key]
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / src["file"]
    if refresh or not path.exists():
        req = urllib.request.Request(
            src["url"], headers={"User-Agent": "proposal-builder/1.0"})
        with urllib.request.urlopen(req, timeout=180) as resp:
            path.write_bytes(resp.read())
    return path


def read_csv(path):
    return list(csv.DictReader(io.open(path, encoding="utf-8-sig")))


# ---------------------------------------------------------------------------
# Matching Nielsen's abbreviated market names to the app's canonical ones
# ---------------------------------------------------------------------------
# Nielsen ships 26-character names -- "Cedar Rapids-Wtrlo-IWC&Dub", "Minot-
# Bsmrck-Dcknsn(Wlstn)" -- and the app knows the expanded ones. 159 of the 210
# are identical; the other 51 have to be paired, and a wrong pairing would put
# a whole market's counties under another market's name. So the pairing is
# required to be a BIJECTION and every pair is confirmed by a second, unrelated
# spelling of the same code (BritCrit's), then confirmed again county by county
# against a third source. No single string comparison is trusted.

_STOPWORDS = {"amp", "dma", "the", "of"}
_STATE_ISH = re.compile(r"^[a-z]{2}$")


def _split_camel(text):
    """'SantaBarbra-SanMar' -> 'Santa Barbra-San Mar'. Nielsen runs words
    together when it needs the characters back. A one-letter fragment is put
    back where it came from: 'McA' is Mc-Allen abbreviated, not 'Mc' and 'A'."""
    parts = re.split(r"(?<=[a-z])(?=[A-Z])", text)
    merged = []
    for part in parts:
        if merged and len(part) == 1:
            merged[-1] += part
        else:
            merged.append(part)
    return " ".join(merged)


def _tokens(name):
    text = _split_camel(name)
    text = text.replace("&", " amp ")
    text = re.sub(r"[^A-Za-z0-9]+", " ", text)
    return [t for t in text.lower().split() if t and t not in _STOPWORDS]


def _skeleton(token):
    """First letter plus consonants, with doubles collapsed.

    'williston' and 'wlstn' both come out 'wlstn', which is the whole point --
    the abbreviation is the same word with its vowels taken out.
    """
    skel = token[0] + re.sub(r"[aeiou]", "", token[1:])
    return re.sub(r"(.)\1+", r"\1", skel)


def _is_subsequence(short, long):
    it = iter(long)
    return all(ch in it for ch in short)


def _token_fit(source_token, canonical_token):
    """How closely an ABBREVIATED token fits a full one, or None.

    Directional on purpose: the source is the abbreviation, so its skeleton
    must be contained in the canonical one and never the other way round.
    Reading it both ways is what made 'Boston' match 'Boise' and 'Juneau'
    match 'Jonesboro' -- two markets an unlucky bijection would have swapped.
    Containment is a SUBSEQUENCE, not a prefix, because Nielsen drops
    consonants mid-word too ('Newpt' for Newport, 'Hztn' for Hazleton).
    """
    src, can = _skeleton(source_token), _skeleton(canonical_token)
    if src[0] != can[0] or not _is_subsequence(src, can):
        return None
    return len(can) - len(src)      # 0 is an exact skeleton match


def _name_score(source_name, canonical_name):
    """How well an abbreviated name fits a canonical one. None if it can't."""
    src = _tokens(source_name)
    can = _tokens(canonical_name)
    if not src or not can:
        return None
    unused = list(can)
    matched = skipped = delta = 0
    for token in src:
        fits = [(_token_fit(token, c), c) for c in unused]
        fits = [(d, c) for d, c in fits if d is not None]
        if fits:
            d, hit = min(fits)
            unused.remove(hit)
            matched += 1
            delta += d
        elif _STATE_ISH.match(token) or token in {"and", "plus"}:
            continue          # a state qualifier or a filler word, not a place
        else:
            skipped += 1      # names something this candidate doesn't have
    if not matched or skipped > matched:
        return None
    # Most tokens explained, fewest source tokens unaccounted for, least of the
    # candidate left over, and the closest character fit -- in that order.
    return (matched, -skipped, -len(unused), -delta)


def bridge_codes_to_canonical(code_names, canonical, label):
    """{code: canonical name} for one source's naming, or raise."""
    canon = list(canonical)
    exact = {}
    pending = {}
    for code, name in code_names.items():
        if name in canon:
            exact[code] = name
        else:
            pending[code] = name

    taken = set(exact.values())
    remaining = [c for c in canon if c not in taken]
    resolved, problems = {}, []
    for code, name in pending.items():
        scored = []
        for cand in remaining:
            score = _name_score(name, cand)
            if score is not None:
                scored.append((score, cand))
        scored.sort(reverse=True)
        if not scored:
            problems.append(f"{label}: {name!r} (code {code}) matches no canonical name")
            continue
        if len(scored) > 1 and scored[0][0] == scored[1][0]:
            problems.append(
                f"{label}: {name!r} (code {code}) is ambiguous between "
                f"{scored[0][1]!r} and {scored[1][1]!r}")
            continue
        resolved[code] = scored[0][1]

    used = Counter(resolved.values())
    for cand, n in used.items():
        if n > 1:
            problems.append(f"{label}: {cand!r} was claimed by {n} different codes")
    out = dict(exact)
    out.update(resolved)
    missing = [c for c in canon if c not in set(out.values())]
    if missing:
        problems.append(f"{label}: {len(missing)} canonical name(s) unclaimed: "
                        + ", ".join(missing[:6]))
    return out, problems


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------
def load_crosswalk():
    if not CROSSWALK.exists():
        sys.exit(f"{CROSSWALK.name} is missing. Build it with "
                 f"`python build_geo_crosswalk.py` first.")
    with gzip.open(CROSSWALK, "rb") as fh:
        return json.loads(fh.read().decode("utf-8"))


def build(refresh=False, report=False):
    cw = load_crosswalk()
    counties = cw["counties"]
    zip_counties = cw["zip_counties"]
    canonical = list(market_profiles.CANONICAL_DMAS)

    assign_rows = read_csv(fetch("county_dma_zip", refresh))
    name_rows = read_csv(fetch("dma_names", refresh))

    # --- the code -> canonical name bridge, agreed by two namings ----------
    # Code 0 is the assignment source's "(NON-DMA COUNTIES)" bucket -- real,
    # and deliberately not a market: those counties are outside all 210 and
    # must come out unassigned rather than under an invented name.
    fissehab = {r["dma_code"].strip(): r["geo_dma"].strip() for r in name_rows
                if r["dma_code"].strip() not in NON_DMA_CODES}
    britcrit = {}
    for r in assign_rows:
        code = r["dma_code"].strip()
        if code not in NON_DMA_CODES:
            britcrit.setdefault(code, r["dma_name"].strip())

    # The bridge is decided by ONE naming and confirmed by the other. The
    # second naming may abstain -- an all-caps "COLUMBUS, GA" cannot separate
    # Columbus-Opelika from Columbus, OH on its own, and a vote it can't cast
    # is not evidence of anything. What it may never do is contradict.
    bridge, problems = bridge_codes_to_canonical(fissehab, canonical, "DMA_Names")
    confirm, _ = bridge_codes_to_canonical(britcrit, canonical, "dma_county_zip")
    confirmed = abstained = 0
    for code, name in bridge.items():
        second = confirm.get(code)
        if second is None:
            abstained += 1
        elif second == name:
            confirmed += 1
        else:
            problems.append(
                f"the two namings of code {code} disagree: "
                f"{fissehab.get(code)!r} -> {name!r} vs "
                f"{britcrit.get(code)!r} -> {second!r}")
    missing_codes = sorted(set(britcrit) - set(bridge), key=str)
    if missing_codes:
        problems.append(
            "the assignment source uses %d code(s) the name list doesn't have: %s"
            % (len(missing_codes), ", ".join(missing_codes[:8])))
    if problems:
        print("NAME BRIDGE FAILED -- nothing was written:")
        for p in problems:
            print("  " + p)
        sys.exit(1)
    print("name bridge: %d codes, %d confirmed by the second naming, %d abstained"
          % (len(bridge), confirmed, abstained))

    code_to_name = dict(bridge)
    key_for = {code: market_profiles.slugify(name) for code, name in code_to_name.items()}

    # --- county assignment -------------------------------------------------
    # A county the resolver can actually encounter is one the crosswalk names
    # OR one its zips point at. Connecticut needs the second half: the 2024
    # gazetteer lists the nine planning regions while the 2020 zip relationship
    # file still uses the eight legacy counties, so zips_to_counties returns
    # legacy FIPS. Assigning only the gazetteer's names would leave every
    # Connecticut zip -- and the whole Hartford-New Haven market -- unresolvable.
    zip_referenced = {f for group in zip_counties.values() for f in group}
    current = set(counties) | zip_referenced
    by_county, retired, outside = {}, [], set()
    for row in assign_rows:
        fips = row["fips"].strip().zfill(5)
        code = row["dma_code"].strip()
        if fips not in current:
            retired.append((fips, row["county"].strip(), row["st"].strip()))
            continue
        if code in NON_DMA_CODES:
            outside.add(fips)       # stated to be in no DMA -- not a gap
            continue
        key = key_for.get(code)
        if key:
            by_county[fips] = key
    retired_fips = sorted({f for f, _, _ in retired})

    # --- counties the source predates, filled from their own zips ----------
    zip_code_map = {}
    for row in assign_rows:
        zip_code_map.setdefault(row["zipcode"].strip().zfill(5), row["dma_code"].strip())
    county_zips = defaultdict(list)
    for zcta, fips_list in zip_counties.items():
        for fips in fips_list:
            county_zips[fips].append(zcta)

    expected = {f for f in current if not f.startswith(TERRITORY_PREFIXES)}
    filled, unassigned = [], []
    for fips in sorted(expected - set(by_county) - outside):
        votes, non_dma = Counter(), 0
        for zcta in county_zips.get(fips, []):
            code = zip_code_map.get(zcta)
            if code in NON_DMA_CODES:
                non_dma += 1
            elif code:
                votes[code] += 1
        if not votes:
            # Every zip it has is stated to be in no DMA -- that is an answer,
            # not a gap. Alaska outside the three metered markets is most of it.
            (outside.add(fips) if non_dma else unassigned.append(fips))
            continue
        code, n = votes.most_common(1)[0]
        by_county[fips] = key_for[code]
        filled.append((fips, counties[fips]["name"], counties[fips]["state"],
                       code_to_name[code], n, sum(votes.values()), len(votes)))

    # --- validation --------------------------------------------------------
    checks = validate(by_county, counties, code_to_name, key_for, canonical,
                      expected, unassigned, outside, refresh, report)

    payload = {
        "version": 1,
        "by_county": by_county,
        "markets": {key_for[c]: {"dma_code": int(c), "name": code_to_name[c]}
                    for c in sorted(code_to_name, key=int)},
        "provenance": provenance(retired_fips, filled, unassigned),
        "validation": checks,
    }
    return payload, filled, unassigned, retired_fips


def provenance(retired_fips, filled, unassigned):
    return {
        "retrieved": date.today().isoformat(),
        "built_by": "build_market_lookup.py",
        "definitions": (
            "DMA is a registered trademark of Nielsen and the market "
            "definitions are Nielsen's intellectual property. No current, "
            "authoritative, freely-licensed county-to-DMA table exists; the "
            "sources below are public secondary compilations. Treat this as a "
            "working table, replaceable by a licensed one without any code "
            "change."),
        "sources": [
            {k: v for k, v in src.items() if k != "file"} | {"used_for": src["role"]}
            for src in SOURCES.values()
        ],
        "assignment_vintage": (
            "The assignment source's county geography is roughly 2013-2014: it "
            "still carries Wade Hampton AK (renamed Kusilvak in 2015), Shannon "
            "SD (renamed Oglala Lakota in 2015), Valdez-Cordova AK (split in "
            "2019), Bedford city VA (reverted to a town in 2013) and the eight "
            "legacy Connecticut counties (replaced by planning regions in "
            "2022). Its DMA assignments are undated; the repository was "
            "published in December 2021. Nielsen moves a handful of counties "
            "between markets each year, so border counties are where this "
            "table is most likely to be wrong."),
        "retired_geography_ignored": len(retired_fips),
        "filled_from_zips": len(filled),
        "unassigned": len(unassigned),
    }


def validate(by_county, counties, code_to_name, key_for, canonical, expected,
             unassigned, outside, refresh, report):
    """Every check reports; the caller decides. Returns a summary dict."""
    out = {}
    print("\nValidation")
    print("-" * 70)

    names = set(code_to_name.values())
    missing_names = [n for n in canonical if n not in names]
    print("  %-58s %s" % ("every DMA name resolves against the canonical 210",
                          "PASS" if not missing_names and len(names) == 210
                          else "FAIL"))
    print("     %d codes -> %d distinct canonical names (canonical list: %d)"
          % (len(code_to_name), len(names), len(canonical)))
    if missing_names:
        print("     missing: " + ", ".join(missing_names))
    out["canonical_names_resolved"] = len(names)
    out["canonical_names_missing"] = missing_names

    assigned_keys = set(by_county.values())
    absent = sorted({code_to_name[c] for c in code_to_name
                     if key_for[c] not in assigned_keys})
    unexpected = [n for n in absent if n not in KNOWN_MARKET_GAPS]
    print("  %-58s %s" % ("all 210 DMAs appear in the county assignment",
                          "PASS" if not unexpected else "FAIL"))
    print("     %d of 210 markets have at least one county" % len(assigned_keys))
    for name in absent:
        why = KNOWN_MARKET_GAPS.get(name)
        print("     %s %s%s" % ("KNOWN GAP" if why else "MISSING", name,
                                " -- " + why if why else ""))
    out["markets_with_counties"] = len(assigned_keys)
    out["markets_without_counties"] = absent

    unreachable = [f for f in unassigned if f.startswith(UNREACHABLE_PREFIXES)]
    real_gaps = [f for f in unassigned if f not in unreachable]
    print("  %-58s %s" % ("every county the resolver can return is assigned",
                          "PASS" if not real_gaps else "FAIL"))
    print("     %d of %d counties assigned; %d stated by the source to be in "
          "no DMA; territories excluded: %d"
          % (len(by_county), len(expected), len(outside),
             len(counties) - len(expected)))
    for fips in real_gaps:
        print("     UNASSIGNED %s %s, %s"
              % (fips, counties[fips]["name"], counties[fips]["state"]))
    if unreachable:
        print("     %d county-equivalent(s) no zip points at, so the resolver "
              "never returns them (Connecticut planning regions -- see "
              "UNREACHABLE_PREFIXES)" % len(unreachable))
    out["counties_expected"] = len(expected)
    out["counties_assigned"] = len(by_county)
    out["counties_unassigned"] = real_gaps
    out["counties_unreachable"] = unreachable
    out["markets_missing_unexpectedly"] = unexpected
    out["markets_known_gaps"] = {n: KNOWN_MARKET_GAPS[n] for n in absent
                                 if n in KNOWN_MARKET_GAPS}
    out["counties_outside_all_dmas"] = sorted(outside)

    # Independent cities: a sloppy table collapses these onto their namesake
    # county. Franklin VA is the one that proves a table really distinguishes
    # them -- the county and the city are in DIFFERENT markets.
    pairs = [("24005", "24510", "Baltimore"), ("51059", "51600", "Fairfax"),
             ("51067", "51620", "Franklin"), ("51159", "51760", "Richmond"),
             ("51161", "51770", "Roanoke"), ("29189", "29510", "St. Louis")]
    print("  independent cities are their own place:")
    spot = {}
    for county_fips, city_fips, label in pairs:
        cm, im = by_county.get(county_fips), by_county.get(city_fips)
        same = "same market" if cm == im else "DIFFERENT markets"
        ok = cm is not None and im is not None
        print("     %-12s county=%-28s city=%-28s %s  %s"
              % (label, cm, im, same, "PASS" if ok else "FAIL"))
        spot[label] = {"county": cm, "city": im}
    out["independent_cities"] = spot

    return out


def crosscheck(by_county, cw, code_to_name, key_for, canonical, refresh):
    """Compare against a third compilation by COUNTY SETS, not by name.

    The obvious cross-check -- resolve their market names to canonical ones and
    compare county by county -- turned out to measure this file's own name
    matcher rather than the table: "Atlanta, GA" scored against "Albany, GA",
    and 54 Georgia counties were reported as disagreements when both sources
    had them in the same place. An expectation that comes out of the code under
    test is not a check.

    So the pairing is structural. Each of their markets is a set of counties;
    each of mine is a set of counties; two compilations of the same reality
    should partition the country almost identically. Markets are paired by
    largest overlap, and only then is the NAME of the paired market compared --
    which makes that a genuinely independent confirmation of the name bridge,
    since nothing about the pairing used a name at all.
    """
    rows = read_csv(fetch("crosscheck", refresh))
    counties, idx, amb = cw["counties"], cw["county_index"], cw["county_ambiguous"]

    theirs = defaultdict(set)
    not_comparable = 0
    for row in rows:
        key = re.sub(r"[^a-z0-9]", "", row["COUNTY"].lower()) + row["STATE_AB"].lower()
        if key in amb or key not in idx or idx[key] not in by_county:
            not_comparable += 1
            continue
        theirs[row["TVDMA"].strip()].add(idx[key])

    mine = defaultdict(set)
    for fips, market in by_county.items():
        mine[market].add(fips)

    pairings, agree, disagree, diffs = [], 0, 0, []
    for label, their_set in theirs.items():
        best, best_overlap = None, 0
        for market, my_set in mine.items():
            overlap = len(their_set & my_set)
            if overlap > best_overlap:
                best, best_overlap = market, overlap
        if best is None:
            not_comparable += len(their_set)
            continue
        union = len(their_set | mine[best])
        pairings.append((label, best, best_overlap, len(their_set), union))
        agree += best_overlap
        for fips in sorted(their_set - mine[best]):
            disagree += 1
            diffs.append((fips, counties[fips]["name"], counties[fips]["state"],
                          by_county[fips], label))

    # Now the name check, on pairings decided without looking at a name.
    name_ok = name_bad = name_unclear = 0
    name_problems = []
    by_key = {market_profiles.slugify(name): name for name in canonical}
    for label, market, overlap, size, union in pairings:
        # Only the LEAD CITY is compared -- "Atlanta, GA - AL - NC" against
        # whatever market its counties landed in. Scoring the whole label
        # against all 210 names put 'Atlanta, GA - AL - NC' closer to
        # 'Albany, GA' than to Atlanta, on the strength of the trailing state
        # list, and reported eleven mismatches on pairings that were 54/54 and
        # 27/27 counties identical. The lead city is the part of the label that
        # actually names the market; the rest is a state footnote.
        lead = _tokens(re.sub(r"\s*DMA$", "", label).split(",")[0])
        mine_tokens = _tokens(by_key.get(market, market.replace("_", " ")))
        if not lead:
            name_unclear += 1
        elif all(any(_token_fit(tok, m) is not None or _token_fit(m, tok) is not None
                     for m in mine_tokens) for tok in lead):
            name_ok += 1
        else:
            name_bad += 1
            name_problems.append((label, market, by_key.get(market, market),
                                  overlap, size))
    return {
        "agree": agree, "disagree": disagree, "not_comparable": not_comparable,
        "pairings": pairings, "diffs": diffs,
        "name_ok": name_ok, "name_bad": name_bad, "name_unclear": name_unclear,
        "name_problems": name_problems,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    ap.add_argument("--refresh", action="store_true", help="re-download the sources")
    ap.add_argument("--report", action="store_true", help="print every difference")
    args = ap.parse_args()

    payload, filled, unassigned, retired = build(refresh=args.refresh, report=args.report)
    cw = load_crosswalk()
    code_to_name = {str(m["dma_code"]): m["name"] for m in payload["markets"].values()}
    key_for = {str(m["dma_code"]): k for k, m in payload["markets"].items()}

    cc = crosscheck(payload["by_county"], cw, code_to_name, key_for,
                    list(market_profiles.CANONICAL_DMAS), args.refresh)
    total = cc["agree"] + cc["disagree"]
    print("\n  cross-check against an unrelated compilation (county sets, not names):")
    print("     %d markets paired by county overlap" % len(cc["pairings"]))
    print("     %d of %d comparable counties land in the paired market (%.2f%%), "
          "%d differ, %d not comparable"
          % (cc["agree"], total, (100.0 * cc["agree"] / total) if total else 0.0,
             cc["disagree"], cc["not_comparable"]))
    print("     name confirmation on those pairings: %d agree, %d differ, %d "
          "too ambiguous to say" % (cc["name_ok"], cc["name_bad"], cc["name_unclear"]))
    for label, market, canonical_name, overlap, size in cc["name_problems"]:
        print("       NAME MISMATCH %r paired with %r by %d/%d counties, but its "
              "name reads as %r" % (label, market, overlap, size, canonical_name))
    payload["validation"]["crosscheck"] = {
        "source": SOURCES["crosscheck"]["repo"],
        "method": "markets paired by county-set overlap; names compared afterwards",
        "counties_in_paired_market": cc["agree"],
        "counties_differing": cc["disagree"],
        "not_comparable": cc["not_comparable"],
        "markets_paired": len(cc["pairings"]),
        "name_agrees": cc["name_ok"],
        "name_differs": cc["name_bad"],
        "name_ambiguous": cc["name_unclear"],
    }
    if cc["diffs"] and args.report:
        for fips, name, state, market, label in cc["diffs"]:
            print("     %s %-30s %-3s ours=%-32s theirs=%s"
                  % (fips, name, state, market, label))

    if filled:
        print("\n  filled from their own zips (geography newer than the source):")
        for fips, name, state, market, votes, total_votes, distinct in filled:
            flag = "" if distinct == 1 else "  <- SPLIT VOTE"
            print("     %s %-40s %-3s -> %-34s %d/%d zips%s"
                  % (fips, name, state, market, votes, total_votes, flag))

    out = Path(args.out)
    blob = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    # mtime=0 so an unchanged input produces a byte-identical file, the same
    # property build_geo_crosswalk.py has.
    with io.open(out, "wb") as raw:
        with gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as fh:
            fh.write(blob)
    print("\nwrote %s (%.1f KiB, %d counties, %d markets)"
          % (out.name, out.stat().st_size / 1024.0,
             len(payload["by_county"]), len(payload["markets"])))


if __name__ == "__main__":
    main()
