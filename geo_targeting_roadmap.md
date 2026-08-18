# Geo targeting & zip/map roadmap

Six pieces. A and B are small and independent. C is the foundation everything
else calls. D changes the shape of the avails table and is the invasive one.
E and F sit on top.

**Sequencing:** A and B (either order) → C (report before wiring) → D (on its
own, not alongside anything) → E → F (needs real sample PDFs in hand).

## Decisions already settled

- Markets break out as separate avails rows by default, with a toggle to
  combine them into one campaign line.
- Markets resolved from a zip list are added automatically, but easy to remove.
- Each avails row seeds its own campaign line, mergeable and splittable
  afterward.
- The map sits beside the avails table on the targeting slide, replacing the
  stock image in that position.
- No paid services anywhere — zip → county → DMA via free data.

---

## A — Selected markets seed the avails table ✅ done

One avails row per selected target market, replacing the originating-market
default rather than sitting beside it, plus a "combine markets into one row"
toggle for the case where several markets sell as a single campaign line.
Same clean/dirty discipline as everywhere else. See
`tests/test_avails_seeding.py`.

## B — Quick-add lines from the avails table ✅ done

An "Add lines" panel above the media plan grid: Product / Audience / Geo
multi-selects generating the cross product as separate lines, with a combine
option. Grid Audience (the `Targeting` column) and Geo are single-select
dropdowns sourced from the avails table. See `tests/test_quick_add_lines.py`.

## C — Geo resolver (no UI) ✅ built — market lookup pluggable, table being sourced

Shipped in `1ea9047`. `geo_resolver.py` carries the pure functions; every one of
them reports what it couldn't resolve rather than guessing.

**Working now, against the committed crosswalk** (`geo_crosswalk.json.gz`, built
by `build_geo_crosswalk.py`, which also carries the county name → FIPS index so
a county can be named the way an avails document writes it):

- `zips_to_counties` — zip → county FIPS, biggest land-area part first.
- `counties_to_zips` — every zip in a county, accepting FIPS codes and
  `NAME STATE` strings mixed, since that's how the avails documents write them.
- `radius_to_zips` — radius around a zip or a geocoded address.

**`zips_to_markets` and `market_to_zips` are live**, against the county→DMA
table described below. The lookup is still an interface rather than a file, so
a purchased or internally-licensed table replaces this one without any code
change — `register_market_lookup(TableMarketLookup(by_county=...))`. Installing
is explicit (`market_lookup.install()`): `geo_resolver` still ships with no
lookup registered, so nothing starts answering market questions by accident,
and with no table at all it reports itself unavailable exactly as before.

### The county→DMA table, and what it is worth

`market_lookup.json.gz` — 3,118 counties → 209 of the 210 markets, built by
`build_market_lookup.py`, asserted by `tests/test_market_lookup.py`.

**There is no current, authoritative, freely-licensed county→DMA table.** DMA
is a Nielsen trademark and the definitions are Nielsen's intellectual property;
everything public is a secondary redistribution of uncertain vintage. This is
therefore a *working* table, built from three public sources, cross-validated,
and documented well enough that swapping in a licensed one is a data change.

| Source | URL | Used for |
|---|---|---|
| BritCrit/dma_county_zip | https://github.com/BritCrit/dma_county_zip | The assignment: county FIPS → DMA **code** + name, and zip → code |
| fissehab/Nielsen-Media-Research-DMA (`DMA_Names.csv`) | https://github.com/fissehab/Nielsen-Media-Research-DMA | DMA code → Nielsen's own abbreviated name, for all 210 — a second, independent naming |
| alex-patton/US-TVDMA-BY-COUNTY | https://github.com/alex-patton/US-TVDMA-BY-COUNTY | Cross-check only. Too old to assign from — it predates Broomfield County CO, created 2001 |

Retrieved 18 August 2026. **Stated vintage:** the assignment source's county
geography is roughly 2013–2014 (it still carries Wade Hampton AK, Shannon SD,
Valdez-Cordova AK, Bedford city VA and the eight legacy Connecticut counties);
its DMA assignments are undated and the repository was published in December
2021. Nielsen moves a handful of counties between markets each year, so
**border counties are where this table is most likely to be wrong.**

**Why the join is exact rather than fuzzy.** The assignment source carries
Nielsen DMA *codes*, so counties attach to codes, not to spellings. Only the
210 code→name pairs need bridging to the app's canonical names, 159 of which
are already identical; the other 51 are Nielsen's 26-character abbreviations
("Cedar Rapids-Wtrlo-IWC&Dub", "Minot-Bsmrck-Dcknsn(Wlstn)") and are matched by
consonant skeleton. That bridge must be a **bijection over all 210** or the
build writes nothing, and each pairing is confirmed by the second naming —
207 of 210 confirmed, 3 abstained where an all-caps name genuinely can't
separate two markets, **zero contradictions**.

**What was validated, and against what:**

- **Names**: 210 codes → 210 distinct canonical names, every one of
  `market_profiles.CANONICAL_DMAS`. A name this table invented would match no
  market profile and no deck slide.
- **Coverage**: every county the resolver can return is assigned — 3,118, plus
  25 the source states are in no DMA at all (Alaska outside the three metered
  markets) and 70 territory counties Nielsen doesn't measure.
- **Independent cities**: Baltimore, Fairfax, Franklin, Richmond, Roanoke and
  St. Louis all carry the county and the city separately. **Franklin is the
  one that proves it** — Franklin County VA is Roanoke-Lynchburg while Franklin
  city VA is Norfolk-Portsmouth-Newport News, so a table that had collapsed the
  pair could not produce this.
- **Cross-check**: markets paired against the third source **by county set, not
  by name** — 206 markets paired, **97.35%** of comparable counties land in the
  paired market, the 2.65% scattered 1–2 counties per market as vintage drift
  would be. Pairing by name instead measured this repo's own name matcher and
  reported 54 Georgia counties as wrong when both sources agreed; the county-set
  pairing is what makes the subsequent name comparison independent evidence
  (202 of 206 lead cities agree).
- **End to end**, on the four real avails documents: Annapolis → Baltimore,
  Wilmington University → Philadelphia (+ Salisbury for lower Delmarva),
  Hershey/Harrisburg's add-on options → Philadelphia + New York, Vienna VA
  radius → Washington-Hagerstown. Every one is the market a person would name.

**Two documented limitations, both asserted by the test so they can't rot:**

1. **Palm Springs is unreachable.** It is a sub-county DMA carved out of
   Riverside County CA, which the source gives whole to Los Angeles, and it has
   no rows of its own anywhere in that source. 209 of 210 markets resolve. The
   cross-check found this independently — the third source's "Palm Springs, CA"
   market pairs with `los_angeles` on its single county.
2. **Connecticut's nine planning regions carry no assignment.** Not a defect in
   this table: `geo_crosswalk.json.gz` builds `counties` from the 2024 gazetteer
   (planning regions) but `zip_counties` from the 2020 relationship file (the
   eight legacy counties), so zips only ever resolve to legacy FIPS — which
   *are* assigned, so Connecticut zips work and Hartford-New Haven resolves.
   Fixing the mismatch belongs to `build_geo_crosswalk.py`.

Counties whose geography postdates the source are filled from **their own
zips'** assignments with the vote reported, not by lineage guesswork: today
that is Oglala Lakota County SD → Rapid City on 10 of 11 zips.

### The licensing answer, which decides what is buildable

**The two halves of the crosswalk have completely different licensing, and the
split is exactly where the roadmap predicted it.**

**zip → county is free and clean.** Both candidates are US Government works:

| Source | What it gives | Size | Notes |
|---|---|---|---|
| [Census 2020 ZCTA→County relationship file](https://www2.census.gov/geo/docs/maps-data/data/rel2020/zcta520/tab20_zcta520_county20_natl.txt) | ZCTA ↔ county, with land/area overlap | **6.5 MiB** (verified) | Decennial. Direct download, no registration. |
| [HUD–USPS ZIP crosswalk](https://www.huduser.gov/portal/datasets/usps_crosswalk.html) | ZIP ↔ county with residential/business ratios | ~5 MiB | **Quarterly**, so it tracks USPS changes. Needs a free HUD USER account. |
| [Census county gazetteer](https://www2.census.gov/geo/docs/maps-data/data/gazetteer/) | county FIPS, name, centroid lat/long | **0.1 MiB** (verified) | Supplies the centroids radius maths needs. |

ZCTAs are not ZIP codes (ZCTAs are areal approximations of USPS delivery
routes; some ZIPs are point/PO-box only and have no ZCTA). HUD is the better
primary for that reason and because it is refreshed quarterly; Census is the
no-registration fallback.

**county → DMA is Nielsen's intellectual property, and there is no clean
public source.** DMA boundaries are owned by Nielsen; the
[ZIP Code by DMA report](https://www.nielsen.com/marketplace/dma/zip-by-dma-annual-report/)
is a paid product, licensed for 12 months, "confidential and internal use
only", and explicitly not to be disclosed to a third party. The FCC's carriage
rules *reference* Nielsen's Local TV Station Information Report as the
authority rather than republishing the county list, so there is no
US-Government-work version to fall back on — the
[Federal Register notice](https://www.federalregister.gov/documents/2022/07/28/2022-16248/update-to-publication-for-television-broadcast-station-dma-determinations-for-cable-and-satellite)
points at Nielsen, not at a downloadable file.

Free county→DMA tables do circulate (Harvard Dataverse, GitHub, Tableau
Public). **They are republications of Nielsen's assignment.** Using one is a
licensing judgment, not a technical one, and it is not a judgment this project
should make silently.

### The recommendation: ask internally before sourcing externally

**TEGNA almost certainly already licenses this.** It is a broadcaster whose
business runs on Nielsen measurement, Premion's own market profile deck is
built on MRI-Simmons/Nielsen data and enumerates all 210 DMAs, and this app
already maps station call signs to DMAs. So the question is probably not
"where do we find a free copy" but **"which internal team already has the
Nielsen county-to-DMA file, and can we have it?"** — which is both lawful and
authoritative, and costs nothing new. That also satisfies "no paid services":
nothing new is bought.

**Failing that, three options, in order of preference:** (1) get a
county→DMA extract from whoever holds TEGNA's Nielsen entitlement;
(2) buy the ZIP-by-DMA report if the entitlement doesn't extend to it —
a one-off, not a per-call service, so it doesn't violate the spirit of the
constraint; (3) use a public republication, with legal sign-off, accepting
that its provenance and currency are both unverifiable.

### What can be built now, regardless

Three of the four functions need no DMA data at all, so the free layer is
buildable immediately and the DMA join is a single lookup bolted on when the
crosswalk arrives:

- `zips_to_counties(zips)` → free, works today
- `counties_to_zips(fips)` → free, works today
- `radius_to_zips(center, miles)` → free, works today
- `zips_to_markets(zips)` / `market_to_zips(market)` → **blocked on the
  county→DMA table**

**The Census Geocoder is confirmed working, free and keyless** — verified
against a real address, returning the county FIPS and coordinates a radius
search needs:

```
1100 WILSON BLVD, ARLINGTON, VA, 22209
  -> county Arlington County, FIPS 51013, (-77.0696, 38.8947)
```

### Proposed function signatures

Every one returns resolved data *and* what it could not resolve, never a
silently shortened list — the same contract as `validate_segments` and the
Wide Orbit parser.

```python
# geo_resolver.py
Resolution = namedtuple("Resolution", "resolved unresolved notes")

def zips_to_markets(zips) -> Resolution
    # resolved: {market_key: {"zips": [...], "zip_count": int}}
    # unresolved: zips with no ZCTA, or whose county maps to no DMA

def counties_to_zips(county_fips) -> Resolution
    # resolved: {fips: [zips]};  unresolved: unknown/retired FIPS

def radius_to_zips(center, miles) -> Resolution
    # center: a 5-digit zip OR a street address (Census Geocoder, keyless)
    # resolved: [zips] within the radius of the centroid
    # unresolved: an address that didn't geocode, or an ambiguous match

def market_to_zips(market_key) -> Resolution
    # resolved: [zips];  unresolved: an unknown market key
```

Radius is measured centroid-to-centroid on the county gazetteer's coordinates,
with a documented note that a zip is included when its *centroid* falls inside
the circle — the same rule freemaptools uses, so results match what account
managers see today.

### Where the data should live

**Supabase, not the repo**, same reasoning as the market profile images: 6.5
MiB of crosswalk is build artefact rather than source, it is refreshed on
someone else's schedule, and the deployed instance already knows how to read
from Supabase with a local fallback. A generated compact lookup (zip → county
FIPS → DMA key, ~46k rows) is well under a megabyte as parquet or a packed
CSV, so a checked-in fallback file is viable if offline resolution matters —
decide once the DMA half's licensing is settled, since that determines whether
the joined table can be committed at all.

### Blocked on — cleared

The county→DMA source is sourced, validated and loaded; both gates are green
and are now permanent assertions in `tests/test_market_lookup.py` rather than
one-off checks:

1. **Against `market_profiles.CANONICAL_DMAS`** — 210 of 210 names reconcile.
2. **End to end against the sample avails documents** that name their markets
   *and* carry their zip lists — the one case where the right answer is already
   written down, so the resolver is held against something it didn't produce.
   All four resolve to the market a person would name.

The remaining exposure is vintage, not correctness of the pipeline: this table
is a public compilation whose border counties may have moved since. Replacing
it with a licensed table is a data change — rebuild the artifact, keep the
provenance block honest, and the tests say whether anything broke.

## D — Targeting groups — not started

The invasive one. A targeting group is a name, an audience, a geo definition
(market / county / zip / radius), the resolved zips and markets, an avails
figure, and a map colour. Groups flow outward into the avails table, proposal
lines, market slides and geo defaults. Existing flat-avails proposals must
load, rebuild and render identically — migrate flat rows into groups on load
rather than changing what is stored for past proposals.

### The audience half of a group is BUILT, not just picked

Extend the Audience finder into a builder: search and add a segment, then add
further segments with one of three actions.

- **AND** (the default) narrows the audience — "Homeowners AND HH Income $150K
  Plus" is one audience defined more precisely. This is how combined audiences
  are almost always built.
- **OR** places two distinct audiences on the same campaign line — "College
  Planning Parents OR Education Services". Less common but real.
- **Add as a separate group** creates a new targeting group with its own avails
  row and its own campaign line.

**Each group carries one avails figure for its whole expression**, since avails
are a property of the combination — an AND stack is smaller than either segment
alone, an OR set larger. That is the same unit the Salesforce avails request
consumes, so a built stack should be directly requestable.

**Enforce the one-custom-audience rule as the stack is built, not at submit.** A
second non-RFP-selectable segment is refused or flagged the moment it is added,
while the strategist is still choosing. (Today the cap is enforced structurally
in `apply_draft_to_form` and warned about at the table; neither tells you at the
point of the decision.)

**Show booking evidence while building.** Framing settled, in this order:

0. **Match as a set.** Normalize every stack to a sorted set of components so
   part order doesn't matter — 24 stored combinations appear in more than one
   order, and string matching would miss them.
1. **Exact match — only when true, never as a warning.** "This exact audience
   has been booked before", quietly. **Never show a frequency for it** (2,745 of
   2,769 stacks were booked exactly once, so a count carries no information) and
   **never show "not booked before"** — absence is the normal case, and a flag
   that fires on ordinary audiences trains people to ignore the panel.
2. **Component familiarity — lead with this.** Per selected segment, how many
   booked stacks it appears in. This answers the question a strategist actually
   has: am I reaching for something standard or something exotic?
3. **Pairwise co-occurrence.** How often the chosen segments have appeared
   together. With three or more, surface the **weakest** pair, since that is the
   least precedented part of the stack.

   **A zero is informative, not a warning** — and what it's allowed to claim
   depends on the components, derived rather than hardcoded:

   - **Both components widely used** → *"Both segments are widely used, but
     haven't been booked together before."* That combination of facts is a real
     signal: the segments are proven, the pairing is new.
   - **Either component rare** → state the per-component familiarity and the
     absence, and claim nothing about commonality: *"Homeowners appears in 172
     booked stacks, NFCU Look Alike in 5. They haven't been booked together."*
     Calling a five-stack client segment "widely used" would be asserting
     something the data doesn't say, and it's the sentence a strategist would
     be right to stop trusting the panel over.

   **"Widely used" = 10 or more booked stacks**, which is the 90th percentile of
   the component distribution, not a round number picked by hand. The
   distribution is heavily skewed — median 2, p75 = 5, p80 = 6, p90 = 10, p95 =
   19 — so a lower bar would call most of the catalog common and say nothing.
   At 10 it selects 147 of 1,292 components (11%), and it separates the worked
   examples the way a person would: `DEMO Homeowner` 172 and `HH Income 150K
   Plus` 153 are widely used; `LIFESTYLE Military Families` 6 and `CLT 1P NFCU
   December Audiences Look Alike` 5 are not. **Recompute it from the data rather
   than freezing the constant** — the percentile is the rule, 10 is only where
   the percentile currently lands.

   With three or more components the same branch applies to the weakest pair
   alone; the other pairs don't need narrating.
4. **Suggested pairings.** Given what's selected, which components most often
   accompany it — the same data read forward instead of backward, which is what
   is useful while someone is still choosing.
5. **Nearest-neighbour stacks** as examples if cheap. Least important of the
   set; skip rather than delay on it.

**What the data says about each signal** (measured across all 3,661 rows, so
none of this is assumed):

- **Component familiarity is the strong one, as expected.** 1,302 distinct
  components with a real spread — 40 appear in 26+ stacks, 246 in 6–25, 467 in
  2–5, and **549 appear exactly once**. That 42%-of-components tail is precisely
  what makes "standard vs exotic" a meaningful thing to say. Worked example:
  `DEMO Homeowner` 172, `HH Income 150K Plus` 153, against `CLT 1P NFCU December
  Audiences Look Alike` 5.
- **Pairwise co-occurrence holds up better in practice than in aggregate.**
  82% of the 3,176 observed pairs co-occur exactly once — but that tail is pairs
  nobody would build. The pairs a strategist actually reaches for score usefully:
  `DEMO Homeowner + HH Income 150K Plus` 13, `DEMO Age A35 Plus + HH Income 100K
  Plus` 18, and even the niche `NFCU Look Alike + Military Families` 3. Still,
  a zero here should read as "these are both common but nobody has combined
  them" — informative — rather than as a warning.
- **Suggested pairings need overlap weighting, not a strict superset.** Asking
  which components accompany *all* selected segments is excellent for one
  segment (`DEMO Age A25 Plus` → Homeowner 14, HH Income 75K Plus 10, AUTO
  Intenders 6) and **collapses to a flat list of 1x at two segments**, because
  few stacks are supersets of a specific pair. Weighting each candidate by how
  many of the selected segments its stack shares restores it: for `DEMO
  Homeowner + HH Income 150K Plus` that yields `DEMO Age A35 Plus` 24, `DEMO Age
  A35-64` 20, `DEMO Age A25 Plus` 18 — real recommendations rather than noise.
  **Build it weighted from the start**; the strict version looks fine in a
  one-segment demo and dies on the second click.
- **Normalize before matching — the stored names are not clean.** 12 component
  groups differ only by case or punctuation: `Lifestyle Charity` (25x) vs
  `LIFESTYLE Charity` (6x), `Lifestyle Outdoors` (15x) vs `LIFESTYLE Outdoors`
  (2x), `LIFESTYLE Pets - Horses` vs `LIFESTYLE Pets Horses`. Matching raw
  strings would split one segment's real history in two and understate
  familiarity by up to 4x. Fold on lowercase-alphanumeric, the same
  normalization the market matcher already uses.

**The expression must round-trip:** stored on the group, restored by history,
and rendered as readable text in the avails table, the Targeting column and
Campaign Specs. A client reads "Homeowners, $150K+ households", never boolean
syntax.

### Finding: the comma in `audience_usage` is AND — confirmed, not assumed

Checked against all 3,661 rows (2,793 of them combinations). The assumption was
right, and three independent lines of evidence agree:

1. **Explicit operators are spelled out INSIDE a part; the comma never carries
   them.** 250 rows contain the word "and" and 121 contain "or", always within a
   single component — `AUTO Truck Intenders OR SUV Intenders, HH Income 75K
   Plus`, `CLT 1P Lennar Homes Buy Sell Home AND HHI 50k Plus`, `CUSTOM DEMO
   Homeowner and HH Income 200K Plus, HH Homes Built Before 2015`. When they
   mean OR they write OR. The comma is reserved for the stack. This is the
   decisive one.
2. **86.5% of combinations join DIFFERENT categories** (DEMO + HH, AUTO + DEMO)
   — narrowing across attributes. The 13.5% that share a category are AND too,
   because a category is a namespace rather than one dimension: `AUTO Intenders,
   AUTO Make Subaru` is Subaru intenders; `DEMO Age A35 Plus, DEMO Homeowner` is
   homeowners over 35.
3. **1,195 of 1,250 comparable combinations (95.6%) delivered no more than their
   smallest component did standalone** — `AUTO Intenders, DEMO Age A25-54`
   delivered 35.4M against 40.0M and 53.2M. An intersection delivers less than
   either part; a union would deliver more. Directional rather than proof, since
   delivery is budget-driven, but it points the same way.

**Two consequences for the builder:**

- **Look up booking evidence as a SET, not a string.** 24 combinations appear in
  more than one part order (`DEMO Homeowner, HH Income 100K Plus` and `HH Income
  100K Plus, DEMO Homeowner` are both stored). Exact string matching would miss
  a stack that has in fact been booked.
- **Exact-stack counts are nearly useless as a frequency signal, and the feature
  should not lean on them.** 2,745 of the 2,769 distinct stacks were booked
  exactly once; 24 twice; **nothing more than twice**. So "how often has this
  been booked" is effectively binary and will read "never" for almost anything a
  strategist builds. Component-level evidence is far richer — 972 distinct
  components, with `DEMO Age A25 Plus` in 222 stacks, `HH Income 100K Plus` in
  177, `DEMO Homeowner` in 171. Better framings: "this exact stack has been
  booked before" (yes/no), "these components appear together in N booked
  stacks", or nearest-neighbour stacks that share most components. Decide the
  framing before building, or the panel will confidently say "never booked"
  about a perfectly ordinary audience.

### Findings carried into D

**`st.column_config.MultiselectColumn` exists and is editable** (Streamlit
1.60.0, verified functionally through `st.data_editor`, not just from the
docstring — it round-trips a list per cell and takes `accept_new_options`,
and is read-only only in `st.dataframe`).

It was deliberately NOT used in B, for two reasons that both point at D:

1. **It expresses the exception, not the default.** A cell holding
   `["Denver", "Atlanta", "Phoenix"]` means *one line covering three markets*.
   B's default is *three separate lines*, which a multi-value cell cannot
   express at all — so the grid would offer only the uncommon case.
2. **It changes the row data model.** Every plan row's `Geo` is a string today,
   flowing into `resolve_row_defaults`, the totals, the preview and the deck's
   table fill. Making it a list touches all of that.

A targeting group genuinely *is* one object spanning several markets, so this
is the column type that fits D — where the row-model change belongs anyway.
`SelectboxColumn`, used in B, has **no** `accept_new_options`: a value outside
its options survives and displays fine (verified), but a rep cannot type a new
one into the cell. B works around it by seeding the options with every value
already on the grid plus the avails values; new values come from the panel
(`st.multiselect` *does* have `accept_new_options`) or the avails table.

## E — Zip/map builder page — not started

Its own page. Inputs resolve through C; outputs are a copyable/exportable zip
list and a map image placed beside the avails table on the targeting slide,
replacing the stock image. Multiple groups render in different colours with a
legend. **Recommend an approach before building**, and confirm rendering works
on Streamlit Cloud's Linux environment, not just Windows.

## F — Salesforce avails PDF import — not started

An avails PDF importer beside the avails table, same shape as the Wide Orbit
schedule importer. Parses per-audience zip lists and avails figures, populates
the table and creates a targeting group per audience with zips resolved to
markets through C. Fixtures gitignored — they contain real client avails and
pricing.

### What the four real samples show

Four real documents are in the project folder (gitignored, `*.pdf`; verified
not tracked) and are referred to here as samples A–D. Everything below was
checked against them rather than described from the format — **two things in
the original spec turned out different**.

Advertiser names and impression volumes are deliberately kept out of this
file. It is committed; the documents carry real client volumes, pricing and
contacts, and the arithmetic below is what matters rather than whose contract
it came from — the same reasoning that keeps the Wide Orbit fixtures
gitignored.

**Structure**, confirmed on all four: an `RFPID-NNNNNN` line, a validity
notice, then `Media Plan Details` (agency, advertiser, flight start/end, total
impressions, sales contact, billing calendar, frequency cap, attribution
products, 3rd-party tag, dayparting), then repeating `Product Summary` /
`Product Details` / `Zip Codes` blocks, then Premion T&Cs and signature lines.

**Totals tie exactly.** Sample C's 60 detail rows (5 audiences × 12 monthly
rows each) sum to precisely the header's Total Impressions, to the impression.
That is the assertion worth leading with, because it catches a parse that
drops or double-counts a row.

**Correction 1 — a DMA is UNPREFIXED, not "DMA Option".** The other three
forms carry a prefix; the DMA one is a bare name in Geography Included:

| Form | Verbatim example |
|---|---|
| DMA (bare name) | `Philadelphia`, `Baltimore`, `New York`, `Washington, D.C.` |
| Named zip option | `Zip Option - <name> Add-On Zips` |
| Radius | `Zip Option - 10mi radius [<origin zip>]` |
| County option | `County Option - <name>` |

So classification is *"prefix if present, otherwise treat as a DMA name"*, and
the unprefixed case has to be the fallback rather than a fourth pattern.
Note `Washington, D.C.` — periods, where our market label is `Washington, DC`;
`market_profiles.match_market` normalizes punctuation away, so it resolves,
but only because it does.

**Correction 2 — the radius origin is optional.** One sample writes
`Zip Option - 10mi radius [<zip>]` with the origin in brackets; another writes
`Zip Option - 10 Mile Radius Zips` — a radius in the name with no bracketed
origin at all, just the resolved zip list. Bracket extraction must be
optional, and the zip list is the authority either way.

**County lists are semicolon-separated `NAME STATE`**, in the shape
`SOMERSET NJ;BUCKS PA;NEW CASTLE DE`, alongside the resolved zips.
`geo_resolver.counties_to_fips` takes exactly this form.

**Parse by layout, not by line.** `extract_text()` interleaves the Zip Codes
block badly — an RFPI id, a county fragment and a zip run land on one visual
line while the columns wrap independently. This is the same trap the Wide
Orbit `.xls` sprung. Use word positions / table extraction, and assert against
the header total so an interleaving error can't pass silently.

### Monthly avails come from a daily rate — verified

Salesforce computes monthly impressions from a **daily rate**, not by
averaging or assuming 30 days. Confirmed twice:

- Sample C's twelve rows per audience all imply **one constant daily rate**
  per audience, each month's figure being that rate times the days in that
  month — a 31-day month and a 30-day month differ by exactly one day's worth.
- Sample A's flight runs 09/06–10/11, which is **36 days**, and its total
  divides to a **whole number of impressions per day**; the September (25
  days) and October (11 days) portions then split to the impression.

**Rounding is FLOOR, not round** — with the exact daily rate, 4 of sample C's
5 audiences match floor on 12/12 rows and round on only 4–5 of 12. The fifth
matches neither cleanly, which says the true daily rate carries more precision
than any single month exposes.

**That is precisely why the document's monthly rows must be used verbatim
where they exist**: the exact rate cannot be recovered from the PDF, so
re-deriving the months would disagree with the source system by an impression
or two per month. Only derive when the document gives a flight total alone:
daily = total ÷ flight days, then per-month = daily × days in that month,
respecting partial months at both ends.

**The single Max Monthly Avails figure is a 30-day equivalent** (daily × 30),
labelled so it is unambiguous. Averaging across active months would make the
same audience read differently depending on which day the flight started.
Full flight stays the grand total, and **the basis toggle converts through the
daily rate, never by multiplying the displayed monthly figure** — a 30-day
equivalent times a month count will not tie back to the real total on a flight
with partial months, which sample A (25 + 11 days) is exactly.

### What the importer populates beyond avails

- advertiser → client name
- agency → the agency toggle and gross markup. Three samples read
  `Direct - No Agency`; one names a real agency, so both paths occur
- flight start/end → the flight dates
- attribution products → the attribution toggles, **suggestively** (below)
- one targeting group per audience-geo pair, **audience-major**, which is the
  order the documents already use

**The audience target is already a boolean expression** and should be parsed
as one: parenthesized terms joined by `AND`, a single term equivalent to an
unparenthesized one — `(DEMO Homeowner) AND (HH Income 200K Plus)`. Each term
matches against the catalog and an unmatched term is reported, never dropped.
This is the same shape the audience builder produces, so **the two
representations must agree** — one parser, one renderer.

`Geography Excluded` and `Excluded Zip Codes` columns exist in the format and
are empty in all four samples. Support them; don't infer their behaviour from
these files.

### Attribution is the one carve-out: the document contributes, it doesn't decide

The general precedence rule is **rep edits > avails document > draft >
defaults**, and for avails figures, zip lists, resolved markets and audiences
the document owns the answer outright — it is the source system's own output.

**Attribution is different, and must not follow that rule.** The proposal and
the eventual order routinely carry attribution products the avail doesn't
list: the avail is requested to size inventory, not to specify the sell, so
its attribution section is typically a SUBSET of what is actually being sold.

So:

- **Pre-check what the document lists.** Those are almost certainly in the
  final order, so checking them saves the rep work and is safe.
- **Never uncheck anything.** Not on import, not on re-import. A toggle the
  rep or the notes turned on stays on.
- **An absent product is not a decision.** Silence in the avail means the
  question wasn't asked, not that the answer is no — and treating it as a no
  would quietly strip a product out of a deck that was meant to carry it.
- **Notes and the rep both extend it freely, and their additions win.**

Add a review-list item saying so — that the avail's attribution is typically a
subset and is worth confirming against what is actually being sold. It belongs
in `unresolved_internal`: it is a check the seller does, not a question for the
client.

The distinction generalises: a source system's output is authoritative about
**what it measured** and merely suggestive about **what someone intends to
sell**. Avails, zips, markets and audiences are the first kind; attribution is
the second.

### F amendment — avails upload as part of the drafting loop

The common real flow is that a rep mentions avails in their notes and has the
Salesforce document in hand, so the upload belongs in the draft review list as
well as beside the table.

- The draft schema gains a signal for whether the notes reference an avails
  pull, distinguishing **three** cases: figures actually stated, avails
  referenced without figures, and no mention at all.
- Where the notes reference avails, the review list shows an **inline
  uploader** at that item rather than a text reminder — "Your notes mention
  avails you've pulled. Upload the document to fill in the table and zip
  targeting."
- Where figures are stated but no document is uploaded, keep the figures and
  still offer the upload, noting that the document adds zip lists and market
  coverage the notes can't carry.
- Where nothing is mentioned, the existing internal review item stands
  unchanged.
- The inline uploader is the **same** uploader as the one beside the avails
  table — a second entry point, not a second implementation — and runs the
  same import and precedence rules: **rep edits > avails document > draft >
  defaults**, with conflicts flagged rather than silently resolved.

The draft-schema signal is buildable independently of the importer; the
uploader wiring is not, since there is no importer to point it at until F.

---

## To gather

Two or three real Salesforce avails PDFs (ideally one with several audiences
and zip lists), plus one real example each of a county list, a zip list and a
radius request as a client would actually send them. Those shape D and F more
than any design conversation will.
