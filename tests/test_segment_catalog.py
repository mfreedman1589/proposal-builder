"""segment_catalog.py + segment_recommend.py -- the ranked usage list.

Synthetic lines for every rule (always run, offline), then the same rules
against the real 10.12.26 ranked list and Matt's hand analysis of it
(`segment_catalog_draft.csv`, the answer key for parsing and the B2B tag)
when both are on disk -- they're gitignored/untracked, so those sections
SKIP without them, same as the Wide Orbit fixtures.

    python tests/test_segment_catalog.py
"""
import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pandas as pd  # noqa: E402

import segment_catalog as sc  # noqa: E402
import segment_recommend as sr  # noqa: E402

REAL_WORKBOOK = ROOT / "IMPRESSIONS BY DATA SEGMENT 10.12.26.xlsx"
ANSWER_KEY = ROOT / "segment_catalog_draft.csv"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def section(title):
    print(f"\n{title}")


def catalog_frame(derived):
    """The derived catalog as audience_catalog would serve it, with no
    `audiences` table underneath (offline)."""
    import audience_catalog
    return audience_catalog.overlay_ranked(
        pd.DataFrame(columns=audience_catalog.CATALOG_COLUMNS), derived, {}, brand_terms=set())


# Synthetic ranked list: row order is usage order. Each line exercises a rule.
SYNTHETIC = [
    "DEMO Age A25-64,HH Income 150K Plus",                                   # 1 no-space split
    "FIN Business Industry Agriculture, DEMO Age A35-64",                    # 2 standalone
    "CLT 1P Acme Widgets FIN Business Industry Agriculture",                 # 3 CLT wraps standard
    "CLT 1P Zephyrine Holdings Members Only",                                # 4 pure CLT
    "CUSTOM Personal Injury Lawyer Intender, DEMO Age A35-64",               # 5 custom recipe
    "CUSTOM Personal Injury Lawyer Intender, DEMO Age A35-64, HH Income 75K Plus",
    "Custom Personal Injury Lawyer Intender, HH Income 100K Plus",           # 7 case variant
    "CUSTOM Personal Injury Lawyer Intender, DEMO Age A35-64, DEMO Homeowner",  # 3rd pairing line
    "WEB RT - Zephyrine Holdings",                                           # 8 retargeting
    "Zephyrine Holdings_RFPID-123456_RT",                                    # 9 RFPID
    "LIFESTYLE Pets Exclusion",                                              # 10 exclusion
    "No Data Targeting",                                                     # 11
    "ShareThis - Acxiom > US > General > Music",                             # 12 taxonomy path
    "CUSTOM Zephyrine Holdings Loyal Shoppers",                              # 13 client-named custom
    "CUSTOM Homeowner AND Auto or Home Insurance",                           # 14 AND/OR kept whole
    "HLH Back Pain, DEMO Age A45 Plus",                                      # 15 typo
    "DEMO Career Entrepeneur",                                               # 16 typo
    "DEMO Career Business Owners Executives Small, DEMO Age A25-54",
    "FIN Business IT Decision Makers, HH Income 100K Plus",
    "DEMO Career Law and Government, HH Income 100K Plus",
    "DEMO Career Healthcare, DEMO Age A25-44",
    "DEMO Career Human Resources",
    "RETAIL Home Services HVAC, DEMO Homeowner",
    "HLTH Back Pain, DEMO Age A45 Plus",
    "LIFESTAGE Prospective College Students, DEMO Age A18-24",
    "AUTO Intenders, DEMO Age A25 Plus",
    "FIN Legal Services, DEMO Age A25 Plus",
    "WEB RT- Pella Carolina, Inc",                                           # comma not a split
]


def synthetic():
    lines = list(enumerate(SYNTHETIC, start=1))
    return sc.derive_segment_catalog(lines, brand_terms=set(), source_filename="synthetic")


def check_parsing():
    section("Splitting and normalizing")
    check("'DEMO Age A25-64,HH Income 150K Plus' is 2 components",
          sc.split_line("DEMO Age A25-64,HH Income 150K Plus")
          == ["DEMO Age A25-64", "HH Income 150K Plus"],
          sc.split_line("DEMO Age A25-64,HH Income 150K Plus"))
    check("a comma not followed by a prefix doesn't split",
          sc.split_line("WEB RT- Pella Carolina, Inc") == ["WEB RT- Pella Carolina, Inc"])
    check("AND/OR inside a name never splits",
          len(sc.split_line("CUSTOM Homeowner AND Auto or Home Insurance, DEMO Age A25 Plus")) == 2)
    check("a lowercase prefix after a comma splits too",
          sc.split_line("AUTO Intenders, Custom Personal Injury Lawyer Intender")
          == ["AUTO Intenders", "Custom Personal Injury Lawyer Intender"])
    check("HLH -> HLTH", sc.clean_display("HLH Back Pain") == "HLTH Back Pain")
    check("Entrepeneur -> Entrepreneur",
          sc.clean_display("DEMO Career Entrepeneur") == "DEMO Career Entrepreneur")
    check("underscores and spacing normalize to one key",
          sc.segment_key("HLTH_Veins _Fibroids") == sc.segment_key("HLTH Veins Fibroids"))

    d = synthetic()
    seg = {s["segment"]: s for s in d["segments"]}
    pi = seg.get("CUSTOM Personal Injury Lawyer Intender")
    check("case variants of one segment merge, raw spellings kept",
          pi is not None and set(pi["raw_variants"]) >= {"CUSTOM Personal Injury Lawyer Intender",
                                                          "Custom Personal Injury Lawyer Intender"},
          pi and pi["raw_variants"])
    check("HLH and HLTH spellings are one segment",
          "HLTH Back Pain" in seg and seg["HLTH Back Pain"]["line_count"] == 2)

    section("Ranking")
    agri = seg["FIN Business Industry Agriculture"]
    check("score is the sum of 1/log2(row+1) over its lines",
          abs(agri["score"] - (sc.rank_weight(2) + sc.rank_weight(3))) < 1e-6, agri["score"])
    check("line count and best row are stored", agri["line_count"] == 2 and agri["best_row"] == 2,
          (agri["line_count"], agri["best_row"]))


def check_flags():
    section("Flags")
    d = synthetic()
    names = " | ".join(s["segment"] for s in d["segments"]) + " | " + " | ".join(
        v for s in d["segments"] for v in s["raw_variants"])
    check("a CLT list wrapping a standard segment counts the standard segment",
          any(s["segment"] == "FIN Business Industry Agriculture" and 3 in range(s["best_row"], 4)
              for s in d["segments"]))
    check("the CLT client's name is dropped entirely", "Acme" not in names)
    check("a pure CLT list never becomes a segment", "Zephyrine Holdings Members" not in names)
    check("WEB RT / RFPID retargeting is excluded", "RFPID" not in names and "WEB RT" not in names)
    check("exclusions, No Data Targeting and taxonomy paths are excluded",
          "Exclusion" not in names and "No Data Targeting" not in names and ">" not in names)
    statuses = {s["segment"]: s["status"] for s in d["segments"]}
    check("a CUSTOM build is a custom recipe",
          statuses.get("CUSTOM Personal Injury Lawyer Intender") == "custom")
    check("a custom build naming a client (harvested from its own CLT/RT lines) is hidden",
          statuses.get("CUSTOM Zephyrine Holdings Loyal Shoppers") == "hidden", statuses)
    check("an AND/OR custom recipe stays one segment",
          statuses.get("CUSTOM Homeowner AND Auto or Home Insurance") == "custom")
    summary = d["summary"]
    check("excluded components are counted",
          summary["excluded_components"].get("client 1P list") == 1
          and summary["excluded_components"].get("retargeting/campaign-specific", 0) >= 3,
          summary["excluded_components"])


def check_roles_and_tags():
    section("Roles")
    for name, role in [("DEMO Age A25-64", "modifier"), ("HH Income 150K Plus", "modifier"),
                       ("DEMO Homeowner", "modifier"), ("DEMO Ethnicity African American", "modifier"),
                       ("DEMO Spanish Speakers", "modifier"), ("DEMO Education Bachelor", "modifier"),
                       ("HH Children in Household", "modifier"), ("HH Home Value 500K Plus", "modifier"),
                       ("HH Dwelling Single Family", "modifier"), ("Females Age 25-49", "modifier"),
                       ("DEMO Career Healthcare", "core"), ("HH Home Built Before 2010", "core"),
                       ("HH Net Worth 1M Plus", "core"), ("AUTO Intenders", "core")]:
        check(f"{name} is {role}", sc.classify_role(name) == role, sc.classify_role(name))

    section("Vertical tags")
    for name, tag in [("HLTH Back Pain", "Healthcare"), ("DEMO Career Healthcare", "Healthcare"),
                      ("LIFESTAGE Prospective College Students", "Education"),
                      ("AUTO Intenders", "Auto"), ("RETAIL Home Services HVAC", "Home Services"),
                      ("RETAIL Home Services HVAC", "Retail"), ("FIN Legal Services", "Legal"),
                      ("CUSTOM Personal Injury Lawyer Intender", "Legal"),
                      ("FIN Banking Intenders", "Financial")]:
        tags, _g, _f = sc.tags_for(name)
        check(f"{name} carries {tag}", tag in tags, tags)
    for name, group, function in [
            ("DEMO Career Business Owner General", "size_owner", None),
            ("FIN B2B Small Business", "size_owner", None),
            ("DEMO Career Self Employed", "size_owner", None),
            ("DEMO Career Entrepreneur", "size_owner", None),
            ("CUSTOM FIN Business Revenue 5M Plus", "size_owner", None),
            ("FIN Business IT Decision Makers", "decision_maker", "it"),
            ("DEMO Career Human Resources", "decision_maker", "hr"),
            ("DEMO Career Property Project and Facility Management", "decision_maker", "facilities"),
            ("HLTH Insurance Prospects and Decision Makers", "decision_maker", "insurance"),
            ("DEMO Career CEO Operations Management", "decision_maker", "executive"),
            ("B2B Construction", "industry", None),
            ("FIN Business Industry Agriculture", "industry", None),
            ("FIN B2B Mid Size Office", "industry", None),
            ("FIN B2B Commercial Property Owner", "industry", None),
            ("DEMO Career Law and Government", "occupation", None),
            ("TRAVEL Business Traveler", "business_travel", None),
            ("AUTO Owner Commercial Vehicle", "business_travel", None)]:
        tags, g, f = sc.tags_for(name)
        check(f"{name} is B2B / {group}" + (f" / {function}" if function else ""),
              "B2B" in tags and g == group and f == function, (tags, g, f))
    for name in ["CUSTOM New Construction Home Intender Next 12 Months", "LIFESTYLE Career Changers",
                 "DEMO Career Unemployed", "HLTH CRX DX Sports Injury"]:
        tags, g, _f = sc.tags_for(name)
        check(f"{name} is not B2B", "B2B" not in tags, (tags, g))
    check("a sports injury isn't Legal", "Legal" not in sc.tags_for("HLTH CRX DX Sports Injury")[0])

    section("Co-occurrence")
    d = synthetic()
    pi = next(s for s in d["segments"] if s["segment"] == "CUSTOM Personal Injury Lawyer Intender")
    check("a custom recipe's most common modifier is its first pairing",
          pi["pairings"] and pi["pairings"][0] == ["DEMO Age A35-64", 3], pi["pairings"])
    modifiers_paired = [p for s in d["segments"] for p in s["pairings"]]
    check("only modifiers are stored as pairings",
          all(sc.classify_role(p[0]) == "modifier" for p in modifiers_paired), modifiers_paired)


def check_overrides():
    section("Admin corrections")
    d = synthetic()
    seg = next(s for s in d["segments"] if s["segment"] == "FIN Legal Services")
    fixed = sc.apply_overrides(seg, {seg["key"]: {"tags": ["Legal"], "role": "core"}})
    check("a tag correction replaces the rule tags", fixed["tags"] == ["Legal"], fixed["tags"])
    check("the original segment is untouched", "Financial" in seg["tags"])
    hidden = sc.apply_overrides(seg, {seg["key"]: {"hidden": True}})
    check("hiding a segment takes it out of sight", hidden["status"] == "hidden")
    shown = sc.apply_overrides(hidden, {seg["key"]: {"hidden": False}})
    check("and unhiding restores it", shown["status"] == "standard")


def check_recommendations_synthetic():
    section("Recommendations (synthetic)")
    frame = catalog_frame(synthetic())
    result = sr.recommend(frame, "VoIP for small law and medical offices")
    top3 = result["core"][:3]
    names = [r["segment"] for r in result["core"]]
    check("a VoIP request reads B2B", result["primary"] == "B2B", result["tags"])
    check("slot 1 is a business-owner segment", top3 and top3[0]["b2b_group"] == "size_owner", top3)
    check("FIN Business IT Decision Makers is in the decision-maker slot",
          any(r["segment"] == "FIN Business IT Decision Makers" and r["slot"] == "Decision maker"
              for r in result["core"]), names)
    check("a law or healthcare career segment is included",
          "DEMO Career Law and Government" in names or "DEMO Career Healthcare" in names, names)
    check("no modifier in the top 3", all(r["role"] != "modifier" for r in top3), top3)
    check("a pairing on fewer than 3 lines isn't shown",
          all(r["pairing"] is None or r["pairing_count"] >= 3 for r in result["core"]),
          result["core"])

    legal = sr.recommend(frame, "personal injury law firm")
    custom = legal["custom"][0] if legal["custom"] else None
    check("a legal request surfaces the Personal Injury Lawyer recipe",
          custom and custom["segment"] == "CUSTOM Personal Injury Lawyer Intender", legal["custom"])
    check("with its most common modifier pairing (3 lines)",
          custom and custom["pairing"] == "DEMO Age A35-64" and custom["pairing_count"] == 3, custom)
    check("a modifier is never recommended alone for a request that doesn't ask for one",
          not legal["modifiers"] and all(r["role"] != "modifier" for r in legal["core"]))
    asked = sr.recommend(frame, "personal injury firm targeting household income over 100k")
    check("but is when the request asks for an income target",
          any("Income" in r["segment"] for r in asked["modifiers"]), asked["modifiers"])
    check("nothing is recommended across all segments when no tag matches",
          sr.recommend(frame, "something unrelated")["core"] == [])

    section("Near-duplicates and pairing rules")
    for a, b, same in [
            ("DEMO Career Business Owner General", "DEMO Career Business Owners Executives", True),
            ("DEMO Career Business Owners Executives", "DEMO Career Business Owners Executives Small",
             True),
            ("HH Home Built Before 2014", "HH Homes Built Before 2015", True),
            ("RETAIL Home Services HVAC", "RETAIL Home Services Home Improvement", False),
            ("DEMO Career Law and Government", "DEMO Career Healthcare", False)]:
        check(f"{a} ~ {b}: {same}", sr.near_duplicates(a, b) == same)
    capped = {"segment": "DEMO Career Law and Government", "tags": ["B2B"], "b2b_group": "occupation",
              "pairings": [["HH Income Less than 100K", 5], ["DEMO Age A25 Plus", 3]]}
    check("an income cap is never paired with a B2B segment",
          sr._pairing(capped) == ("DEMO Age A25 Plus", 3), sr._pairing(capped))
    consumer = dict(capped, tags=["Legal"], b2b_group=None)
    check("but can be with a consumer segment",
          sr._pairing(consumer)[0] == "HH Income Less than 100K")
    check("a pairing under 3 lines shows nothing",
          sr._pairing({"tags": [], "pairings": [["DEMO Age A25 Plus", 2]]}) == (None, 0))


def check_expanded_tags():
    section("Expanded vertical tags")
    for name, tag, present in [
            ("FIN Services Retirement Planning", "Entertainment/Events", False),
            ("FIN Services Retirement Planning", "Financial", True),
            ("ENT Casino Player", "Entertainment/Events", True),
            ("ENT Casino Player", "Tourism/Travel", True),
            ("TRAVEL Casino", "Entertainment/Events", True),
            ("TRAVEL Casino", "Tourism/Travel", True),
            ("AFIRST Travel Buffs and Sightseers", "Tourism/Travel", True),
            ("AFIRST Foodies", "Restaurants/Food", True),
            ("FOOD QSR Pizza", "Restaurants/Food", True),
            ("SPORT Events Visitors", "Sports", True),
            ("AFIRST Sports Super Fans", "Sports", True),
            ("LIFESTYLE Outdoors Hunting", "Outdoors/Recreation", True),
            ("AFIRST Fitness Fanatics", "Fitness/Wellness", True),
            ("HLTH Insurance Medical", "Insurance", True),
            ("HLTH Insurance Medical", "Healthcare", True),
            ("AUTO Insurance", "Auto", True),
            ("LIFESTAGE Home Buyer", "Real Estate/Movers", True),
            ("POL Premion Likely Voter", "Political/Advocacy", True),
            ("DEMO Career Military Veteran", "Military/Government", True),
            ("LIFESTYLE Charity", "Nonprofit/Charity", True),
            ("LIFESTYLE Pets", "Pets", True),
            ("ENT Cord Cutters", "Telecom/Tech", True),
            # Whole words only -- never a substring of a longer word.
            ("LIFESTYLE Rentertainment Fans", "Entertainment/Events", False),
            ("LIFESTYLE Petitions", "Pets", False)]:
        tags = sc.tags_for(name)[0]
        check(f"{name} {'has' if present else 'has no'} {tag}", (tag in tags) == present, tags)


def real_derived():
    if not REAL_WORKBOOK.exists():
        return None
    return sc.derive_from_workbook(str(REAL_WORKBOOK), source_filename=REAL_WORKBOOK.name)


def check_real_counts(derived):
    section("Real ranked list vs the answer key")
    if derived is None or not ANSWER_KEY.exists():
        print("  SKIP  the real workbook or segment_catalog_draft.csv isn't on disk")
        return
    summary = derived["summary"]
    rows = list(csv.DictReader(open(ANSWER_KEY, encoding="utf-8")))
    key_unique = len(rows)
    key_rec = sum(1 for r in rows if not r["flags"])
    print(f"  ....  unique {summary['unique_segments']} (key {key_unique}), "
          f"recommendable {summary['recommendable']} (key {key_rec})")
    # Tolerance: the key leaves seven comma+prefix lines unsplit and counts
    # each CLT wrapper as its own entry; both make it a little larger.
    check("unique segments within 5% of the key",
          abs(summary["unique_segments"] - key_unique) <= 0.05 * key_unique)
    check("recommendable within 3% of the key", abs(summary["recommendable"] - key_rec) <= 0.03 * key_rec)

    mine = {s["key"]: s for s in derived["segments"]}
    key_b2b = {sc.segment_key(r["segment"]) for r in rows if r["b2b"] == "Y"}
    real_b2b = {k for k in key_b2b if k in mine}
    tagged = {k for k in real_b2b if "B2B" in mine[k]["tags"]}
    check("every answer-key B2B row that is a real segment is tagged B2B",
          tagged == real_b2b, sorted(mine[k]["segment"] for k in real_b2b - tagged))
    key_rec_keys = {sc.segment_key(r["segment"]) for r in rows if not r["flags"]}
    mine_rec = {k for k, s in mine.items() if s["status"] == "standard"}
    check("95%+ of the key's recommendable segments are recommendable here",
          len(key_rec_keys & mine_rec) >= 0.95 * len(key_rec_keys),
          len(key_rec_keys & mine_rec))


def check_real_recommendations(derived):
    section("Real ranked list: recommendations and leaks")
    if derived is None:
        print("  SKIP  the real workbook isn't on disk")
        return
    frame = catalog_frame(derived)
    for query in ["VoIP for small law and medical offices",
                  "Netmaker sells VoIP and managed IT services to small law, medical, insurance "
                  "and financial offices"]:
        netmaker = sr.recommend(frame, query)
        core = netmaker["core"]
        names = [r["segment"] for r in core]
        top3 = core[:3]
        check("Netmaker: a business-owner segment first", top3[0]["b2b_group"] == "size_owner",
              top3[0])
        check("Netmaker: FIN Business IT Decision Makers", "FIN Business IT Decision Makers" in names,
              names)
        check("Netmaker: a law or healthcare career segment",
              "DEMO Career Law and Government" in names or "DEMO Career Healthcare" in names, names)
        check("Netmaker: no modifier in the top 3", all(r["role"] != "modifier" for r in top3), top3)
        check("Netmaker: at most 6 picks", len(core) <= 6, len(core))
        dupes = [(a, b) for i, a in enumerate(names) for b in names[i + 1:]
                 if sr.near_duplicates(a, b)]
        check("Netmaker: no near-duplicates", not dupes, dupes)
        check("Netmaker: the other owner segments are listed as similar, not picked",
              "DEMO Career Business Owners Executives" not in names
              and any("DEMO Career Business Owners Executives" in r["similar"] for r in core), core[0])
        check("Netmaker: nothing unrelated to the request (no commercial vehicle)",
              "AUTO Owner Commercial Vehicle" not in names, names)
        check("Netmaker: no income-cap pairing on a B2B pick",
              not any(r["pairing"] and "Less than" in r["pairing"] for r in core),
              [(r["segment"], r["pairing"]) for r in core])
    nearest = [r for r in core if r["nearest_match"]]
    check("financial offices fall back to the nearest available segment, labeled",
          nearest and nearest[0]["segment"] in ("DEMO Career White Collar Worker",
                                                "FIN Business Owner Small")
          and "nearest match" in nearest[0]["slot"], core)

    legal = sr.recommend(frame, "personal injury")
    pi = [r for r in legal["custom"] if "Personal Injury Lawyer Intender" in r["segment"]]
    data_pairings = next((s["pairings"] for s in derived["segments"]
                          if s["segment"] == "CUSTOM Personal Injury Lawyer Intender"), [])
    expected = next((p[0] for p in data_pairings if p[1] >= sr.MIN_PAIRING_LINES), None)
    check("Legal: the Personal Injury Lawyer recipe is surfaced", bool(pi), legal["custom"])
    check("Legal: its pairing is the data's most common one on 3+ lines, else none",
          pi and pi[0]["pairing"] == expected, (pi and pi[0]["pairing"], data_pairings[:2]))

    for query, vertical, tag in [("hospital system", "healthcare", "Healthcare"),
                                 ("HVAC company", "home_improvement", "Home Services"),
                                 ("community college enrollment", "education", "Education"),
                                 ("car dealership", "auto", "Auto"),
                                 ("CVB campaign for weekend getaways", None, "Tourism/Travel"),
                                 ("family restaurant and brewery", None, "Restaurants/Food"),
                                 ("fitness club memberships", None, "Fitness/Wellness")]:
        result = sr.recommend(frame, query, vertical_key=vertical)
        first = result["core"][:5]
        check(f"{tag}: vertical CORE segments first",
              first and all(tag in r["tags"] and r["role"] == "core" for r in first),
              [(r["segment"], r["tags"], r["role"]) for r in first])
    tourism = sr.recommend(frame, "CVB campaign for weekend getaways")["core"]
    check("a tourism request leads with TRAVEL segments, never age/income",
          tourism and tourism[0]["segment"].startswith("TRAVEL")
          and not any(r["role"] == "modifier" for r in tourism), [r["segment"] for r in tourism])
    food = sr.recommend(frame, "family restaurant and brewery")
    check("a restaurant request leads with FOOD segments",
          food["core"] and food["core"][0]["segment"].startswith("FOOD"),
          [r["segment"] for r in food["core"]])
    check("'family restaurant' doesn't read as a request for a children target",
          not food["modifiers"], food["modifiers"])
    casino = sr.recommend(frame, "casino resort")
    check("a casino request reads Entertainment, then Tourism",
          casino["tags"][:2] == ["Entertainment/Events", "Tourism/Travel"], casino["tags"])

    section("Booking evidence from the ranked list")
    import audience_evidence
    index = audience_evidence.build_index_from_stacks(derived["stacks"])
    evidence = audience_evidence.evidence_for(["AUTO Intenders", "DEMO Age A25 Plus"], index)
    lines = audience_evidence.evidence_lines(evidence)
    check("familiarity counts booked lines, not impressions",
          any("booked lines" in line for line in lines) and not any("impressions" in l for l in lines),
          lines)
    check("no CLT text in the evidence stacks",
          not any(p.upper().startswith("CLT") for stack in derived["stacks"] for p in stack))
    check("a name containing a comma stays one component",
          not any(p.startswith("General Interest - Entertainment")
                  for stack in derived["stacks"] for p in stack))

    # No client name harvested from the list's own CLT/retargeting lines
    # appears in anything a seller can see.
    harvested = []
    for _row, line in sc.read_ranked_lines(str(REAL_WORKBOOK)):
        for raw in sc.split_line(line):
            if sc._CLT_RE.match(raw):
                harvested.append(sc.unwrap_clt(raw)[1])
            elif sc._RETARGETING_RE.search(raw) and not sc._CUSTOM_RE.match(raw.strip()):
                harvested.append(sc._client_text_of_retargeting(raw))
    generic = " ".join(sc._words(s["segment"]) for s in derived["segments"]
                       if s["prefix"] not in ("", "CUSTOM"))
    terms = sc.harvest_client_terms(harvested, generic) | sc.load_brand_terms()
    visible = frame[frame["visible"]]
    outputs = list(visible["segment"])
    for query in ["B2B", "personal injury", "VoIP for small law and medical offices",
                  "credit union", "auto dealer", "home builder", "HVAC", "performing arts"]:
        outputs += [r["segment"] for r in sr.flat_list(sr.recommend(frame, query))]
        outputs += [r["pairing"] for r in sr.flat_list(sr.recommend(frame, query)) if r["pairing"]]
    leaks = sorted({(t, o) for o in outputs for t in terms if f" {t} " in sc._words(o)})
    check(f"no client name ({len(terms)} harvested + curated) in the visible catalog or "
          f"recommendations", not leaks, leaks[:10])
    check("no CLT text anywhere visible", not any(o.upper().startswith("CLT") for o in outputs))


def check_finder_page():
    section("The Finder page never shows a client name")
    from streamlit.testing.v1 import AppTest

    frame = catalog_frame(synthetic())
    script = '''
import sys
sys.path.insert(0, {root!r})
import pandas as pd
import finders
frame = pd.read_pickle({pickle!r})
finders.load_audience_catalog = lambda: frame
import catalog_shared
catalog_shared.load_audience_catalog = lambda: frame
finders.render_audience_finder(None)
'''
    import tempfile
    tmp = Path(tempfile.mkdtemp())
    frame.to_pickle(tmp / "catalog.pkl")
    page = tmp / "page.py"
    page.write_text(script.format(root=str(ROOT), pickle=str(tmp / "catalog.pkl")), encoding="utf-8")
    for query in ["B2B", "personal injury", "Zephyrine", "Acme"]:
        at = AppTest.from_file(str(page), default_timeout=60)
        at.run()
        at.text_input(key="finder_search").set_value(query).run()
        text = " ".join(str(getattr(e, "value", "")) for e in at.markdown) + " ".join(
            str(getattr(e, "value", "")) for e in at.caption)
        # The empty-result line echoes what the seller typed -- their own
        # text, not a name from the data -- so it's taken out first.
        text = text.replace(f"“{query}”", "")
        check(f"search {query!r} renders without error", not at.exception, at.exception)
        check(f"search {query!r} shows no client name",
              "Zephyrine" not in text and "Acme" not in text, text[:300])
        if query == "B2B":
            check("a B2B search shows the B2B recommendation panel",
                  "Recommended for B2B" in text and "FIN Business IT Decision Makers" in text)
        if query == "personal injury":
            check("a personal-injury search shows the custom recipe",
                  "Proven custom recipes" in text
                  and "CUSTOM Personal Injury Lawyer Intender" in text)


if __name__ == "__main__":
    check_parsing()
    check_flags()
    check_roles_and_tags()
    check_overrides()
    check_expanded_tags()
    check_recommendations_synthetic()
    derived = real_derived()
    check_real_counts(derived)
    check_real_recommendations(derived)
    check_finder_page()
    print(f"\n{'FAILED: ' + str(len(failures)) if failures else 'All checks passed.'}")
    sys.exit(1 if failures else 0)
