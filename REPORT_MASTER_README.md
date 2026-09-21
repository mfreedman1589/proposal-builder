# REPORT_MASTER_v0_4.pptx — handoff notes

Eight slides, 13.333 × 7.5 in, restyled to the Mattress Warehouse rebuild look:
navy `000946` header band with periwinkle `9999FF` title / white subtitle and the
hatch triangle at the right edge, navy KPI tiles (white value, `C9CBE8` label),
`F3F4FB` cards with a soft shadow, navy-header tables with `EBEBEB` rules,
Arial throughout, TEGNA + PREMION lockup bottom-left (white, in the band, on the
recap cover). Shape names, tokens, two-run bullets and `key:` notes are
unchanged from v0_1 — v0_2 was a drop-in for the tagging/wiring work.

## What changed in v0_4

**Built by `migrate_report_master_v0_4.py` (python-pptx, not by hand) —
the ATTRIBUTION_REPORT_PLAN.md WAEPA conversions section.** Highlights
slide only:

- The three existing tiles are now NAMED, matching the convention every
  other tile row in this deck already uses: `ImpressionsTile` (was the
  `{{HEADLINE_IMPRESSIONS}}` tile), `VisitorsTile`
  (`{{HEADLINE_UNIQUE_VISITORS}}`), `RateTile`
  (`{{HEADLINE_ATTRIBUTED_RATE}}`) — each with `…Value`/`…Label` children.
  They predated the convention before v0_4; nothing about their look,
  position or token changed, only their names.
- A fourth tile, `ConversionsTile` / `ConversionsTileValue` /
  `ConversionsTileLabel`, token `{{HEADLINE_CONVERSIONS}}`, label
  "Attributed conversions" — a same-formatting deep copy of `RateTile`.
  Present only on a report with real conversions; the fill code deletes it
  and reflows the other three back to their v0_3 span (same mechanism as
  `FlightTile` on the recap slide) whenever it isn't.
- All four tiles are evenly spaced across the SAME row the original three
  occupied — same left/right edge, same ~0.3" gap between tiles, each now
  a quarter of the row instead of a third.

Everything else — slide order, every other token/named shape — is
unchanged from v0_3.

**Not done in v0_4, and not urgent:** `BreakdownTable` (4 cols → 5, add
"Conv. Rate"), `IntentSummaryTable` (3 → 4, add "Converted"), `TopUrlTable`
(3 → 4, add "Converted") — a conversions export's fill code already tries
to write these columns and gracefully drops them with a warning until the
template catches up (`_fill_named_table`'s own column-count check), so
there's no urgency; widen them whenever convenient and the extra column
starts showing up on its own, no code change needed.

## What changed in v0_3

Slide order: recap, highlights, delivery_recap, **delivery_breakdown (new)**,
attribution_breakdown, url_report, zip_analysis, takeaways.

1. **Delivery recap** — five NAMED tiles instead of four unnamed ones:
   `DeliveredTile`, `VcrTile`, `FrequencyTile`, `UniquesTile`, `CtvShareTile`,
   each with a `…Value` and a `…Label` child, so code can edit or delete a whole
   tile. New token `{{CTV_SHARE}}` (Connected TV share, from the OTT
   DISTRIBUTION tab). `CreativeTable` and `{{CREATIVE_ROWS}}` are **gone** —
   by-creative delivery moved to the new slide. `TopPublishersTable` stays and
   is sized for a top-5 cap. `ChartRegion` stays but is now the **daypart** bar
   chart; its label reads "ChartRegion — delivery by daypart".
2. **Delivery breakdown (new slide)** — `key: report:delivery_breakdown` +
   `delivery_set: true`. `{{DELIVERY_BREAKDOWN_NOTE}}` is the subtitle.
   `DeliveryByGeoTable` / `{{DELIVERY_BY_GEO_ROWS}}` (Geography | Impressions |
   VCR) and `DeliveryByCreativeTable` / `{{DELIVERY_BY_CREATIVE_ROWS}}`
   (Creative | Impressions | VCR) stack in the left column; `ChartRegion` +
   `ChartRegionLabel` (VCR by creative) fill the right; narrative is
   `DeliveryBreakdownNarrative` / `{{DELIVERY_BREAKDOWN_NARRATIVE}}`.
   **Both tables are conditional.** Each has its own named section header —
   `DeliveryByGeoHeader`, `DeliveryByCreativeHeader` — so delete the header with
   its table. The columns are independent: dropping either table leaves the
   other, the chart and the narrative in place with clean whitespace, no reflow
   required.
3. **URL report** — `IntentSummaryTable` / `{{INTENT_SUMMARY_ROWS}}`
   (Intent | Visits | Share) sits above `TopUrlTable` in the left column.
   Row budget: the intent block is laid out for 4 data rows and the URL block
   for 7, both at 0.30". Expect 1–6 intent rows — past 4, shift the
   `TOP PAGES` header and `TopUrlTable` down by the overflow, or trim the URL
   rows. The two blocks do not auto-reflow.
4. **Zip analysis** — `TopZipTable` is five columns: Zip | Area | Impression
   share | Attributed rate | Multiple vs. avg, at 9pt so all five headers fit
   one line. Token is still `{{TOP_ZIP_ROWS}}`. `MapRegion` is unchanged in
   role, narrowed to 5.8" to make room for the wider table.

Unchanged and still guaranteed: every other shape name, every other token, all
`key:` notes, the two-run head/detail bullets on highlights and takeaways, and
`TegnaLogo`. `ReportPeriodTile` / `FlightTile` / `GeographyTile` (+ `…Value`,
`…Label`) on the recap slide are carried through from v0_2.
Every slide already carries a `key:` line in speaker notes. Slide 3 also carries
`delivery_set: true` on a second line so the delivery-set membership is in the
deck, not only in code.

## Deviations from the inventory (deliberate)

- **Highlights: `KPI_STRIP_ROWS` dropped.** Three headline tiles + four
  head/detail bullets is the slide. Delivery-only stats live on slide 3.
- **Attribution breakdown: one table, one chart.** `BY_MARKET_ROWS`,
  `BY_AUDIENCE_ROWS`, `BY_CREATIVE_ROWS` replaced by a single
  `{{BREAKDOWN_ROWS}}` table whose first header cell is
  `{{BREAKDOWN_DIMENSION_LABEL}}`. Python picks the dimension.
- **`{{ATTRIBUTION_HEADLINE_NOTE}}` kept** for consistency with the URL and
  zip slides. It must be filled — an unfilled token is client-visible.
- **No budget element on the recap.** `{{FLIGHT_LABEL}}` sits in its own tile;
  if it's blank, delete that tile's shapes or leave the label with an empty value.

## Conventions the fill code should rely on

**Tables** — one header row + one template data row. The token sits in the
first cell of the data row; other cells are empty. Clone the data row per real
row and delete the template row. Named: `TopPublishersTable`, `DeliveryByGeoTable`,
`DeliveryByCreativeTable`, `BreakdownTable`, `IntentSummaryTable`,
`TopUrlTable`, `TopZipTable`.

**Head/detail bullets** — one text box, four paragraphs, each paragraph is
exactly two runs: run 1 bold navy `{{X_HEAD_n}}`, run 2 plain ` {{X_DETAIL_n}}`
(leading space is in the run). Delete unused paragraphs whole. Named:
`HIGHLIGHTBullets`, `TAKEAWAYBullets`.

**Single-token bullets** — one paragraph with the token; clone per item.
Named: `GoalsBullets`, `AudienceBullets`, `WhatsNextBullets`.

**Image regions** — a dashed rectangle named `ChartRegion` / `MapRegion` plus a
separate label text box named `ChartRegionLabel` / `MapRegionLabel` sitting on
top of it. Read position/size from the rectangle, place the picture, then
delete **both** shapes. Slides 3, 4 and 5 each have a `ChartRegion`; slide 7 has
`MapRegion`.

**Narratives** — plain text boxes named `DeliveryNarrative`,
`DeliveryBreakdownNarrative`, `AttributionNarrative`, `UrlIntentNarrative`,
`ZipNarrative`.

**KPI tiles** — named tiles come in threes: `<Name>Tile` (the rounded rect),
`<Name>TileValue` (the token) and `<Name>TileLabel` (the caption). Delete all
three to drop a tile, then reflow the survivors across the row.

## Tokens by slide

| # | key | Tokens |
|---|---|---|
| 1 | report:recap | CLIENT_NAME, REPORT_TITLE, REPORT_PERIOD_LABEL, FLIGHT_LABEL, GEOGRAPHY_LABEL, GOALS_BULLETS, AUDIENCE_BULLETS |
| 2 | report:highlights | HEADLINE_IMPRESSIONS, HEADLINE_UNIQUE_VISITORS, HEADLINE_ATTRIBUTED_RATE, HEADLINE_CONVERSIONS *(4th tile, present only when the export has conversions)*, HIGHLIGHT_HEAD_1..4, HIGHLIGHT_DETAIL_1..4 |
| 3 | report:delivery_recap *(delivery set)* | DELIVERED_IMPRESSIONS, VCR, FREQUENCY, UNIQUES, CTV_SHARE, TOP_PUBLISHERS_ROWS, DELIVERY_NARRATIVE + ChartRegion (daypart) |
| 4 | report:delivery_breakdown *(delivery set)* | DELIVERY_BREAKDOWN_NOTE, DELIVERY_BY_GEO_ROWS, DELIVERY_BY_CREATIVE_ROWS, DELIVERY_BREAKDOWN_NARRATIVE + ChartRegion (VCR by creative) |
| 5 | report:attribution_breakdown | ATTRIBUTION_HEADLINE_NOTE, BREAKDOWN_DIMENSION_LABEL, BREAKDOWN_ROWS, ATTRIBUTION_NARRATIVE + ChartRegion |
| 6 | report:url_report | URL_HEADLINE_NOTE, INTENT_SUMMARY_ROWS, TOP_URL_ROWS, URL_INTENT_NARRATIVE |
| 7 | report:zip_analysis | ZIP_HEADLINE_NOTE, TOP_ZIP_ROWS, ZIP_NARRATIVE + MapRegion |
| 8 | report:takeaways | TAKEAWAY_HEAD_1..4, TAKEAWAY_DETAIL_1..4, WHATS_NEXT_BULLETS |

## Build notes

Generated with pptxgenjs from `build_report_master_v0_3.js` (reads the four PNGs in
`assets/`), then post-processed by `postfix.py REPORT_MASTER_v0_3.pptx`, which
(1) strips the stray mid-paragraph `<a:pPr><a:buNone/>` pptxgenjs emits before
the second run of a two-run bullet, and (2) injects the TEGNA wordmark as a
vector `custGeom` shape named `TegnaLogo` (path in `tegna_path.xml`, lifted
from the MW deck — pptxgenjs can't write custom geometry). `TegnaLogo` is the
only new named shape; the fill code should leave it alone. If the deck is ever
regenerated, run both. `validate.py` passes clean.

`assets/`: `hatch_tri.png` (band corner), `hatch_L.png` (cover motif),
`premion_white.png`, `premion_black.png` — all pulled from the MW rebuild.

Tile value tokens wrap in the template because the token strings are long;
real values ("3,102,415", "1.2%", "September 2026") fit on one line.

## Corrections after wiring (2026-09-06)

- **Row caps are design rules, not "however many fit": Top 5 publishers, Top 3
  creatives** (was "Top 10" for publishers, uncapped for creatives) — a
  client reads a top-N list regardless of slide space, and 10+5 on one slide
  carrying two tables never fit to begin with. The fill code's own
  `condense_avails_table`-based shrink-to-fit pass sits underneath these caps
  as a safety net, not the other way around.
- **The Geography tile never truncates a market join.** Up to 2 real market
  names are listed; 3 or more collapses the tile to "N markets," and the full
  list moves to the recap's `AUDIENCE_BULLETS` as its own line instead. A
  tile holds one short value — a value that would run past one line belongs
  in a different container, not a smaller font.
- **The three recap tiles (`ReportPeriodTile`/`FlightTile`/`GeographyTile`,
  each with a `...Value` and `...Label` child) are now named shapes**, so the
  fill code can delete a tile outright when it has nothing to show — the
  Campaign Flight tile is deleted whole in every Phase 3 report, since
  `FLIGHT_LABEL` has no real value until a proposal is linked (Phase 5).

## Conventions the fill code should rely on

**Tables** — one header row + one template data row. The token sits in the
first cell of the data row; other cells are empty. Clone the data row per real
row and delete the template row. Named: `TopPublishersTable`, `CreativeTable`,
`BreakdownTable`, `TopUrlTable`, `TopZipTable`.

**Head/detail bullets** — one text box, four paragraphs, each paragraph is
exactly two runs: run 1 bold navy `{{X_HEAD_n}}`, run 2 plain ` {{X_DETAIL_n}}`
(leading space is in the run). Delete unused paragraphs whole. Named:
`HIGHLIGHTBullets`, `TAKEAWAYBullets`.

**Single-token bullets** — one paragraph with the token; clone per item.
Named: `GoalsBullets`, `AudienceBullets`, `WhatsNextBullets`.

**Image regions** — a dashed rectangle named `ChartRegion` / `MapRegion` plus a
separate label text box named `ChartRegionLabel` / `MapRegionLabel` sitting on
top of it. Read position/size from the rectangle, place the picture, then
delete **both** shapes. Slides 3 and 4 each have a `ChartRegion`; slide 6 has
`MapRegion`.

**Narratives** — plain text boxes named `DeliveryNarrative`,
`AttributionNarrative`, `UrlIntentNarrative`, `ZipNarrative`.

- **Row caps are design rules, not fit guesses: top 5 publishers, top 3
  creatives** (`TOP_PUBLISHERS_ROW_CAP` / `TOP_CREATIVES_ROW_CAP` in
  `report_assembly.py`). What a client reads, independent of how much room
  the slide has. `assembly.condense_avails_table` runs underneath as a
  measured shrink-to-fit safety net, not as a substitute for the cap.
- **The Geography tile never truncates a market join.** Up to 2 real market
  names are listed; 3 or more collapses to "N markets" and the full list
  moves into the recap's `AUDIENCE_BULLETS`. A tile holds one short value.

## Handoff: new slide for Phase 7 -- Automotive Registrations (Polk match-back)

**This is a human deliverable, not something Claude builds** (ATTRIBUTION_
REPORT_PLAN.md Phase 7's own explicit ruling) -- the fill code
(`report_assembly._fill_automotive_registrations`) and every other piece of
the Polk plumbing (parser, facts payload, "Project for match rate" toggle,
upload slot) are already built and tested against the real `Polk
Dashboard.xlsx` fixture; they're just waiting on this slide to exist in the
template. Until it does, the code's own drop-if-absent branch means nothing
breaks -- a Polk file can be uploaded and a report generated today, it just
won't carry this slide yet.

**Slide key** (speaker notes, same convention as every other slide in this
deck): `key: report:automotive_registrations`. This slide is OPTIONAL in the
template, the same way `report:live_sports` and `report:ott_retargeting`
were before their own templates landed -- `build_report_deck` drops it
cleanly when either the key is absent or no Polk file was uploaded, and
never treats it as a required slide.

**Placement, confirmed 2026-09-16:** put this slide LAST, after
`report:takeaways` -- immediately before wherever an appended Auto-Sales
Analyst deck would land (that deck is grafted on wholesale, after every
other slide, via `extra_deck_path`/`append_slide_deck`, so "last real slide
in the template" is what puts this slide immediately before it by
construction; no code change is needed to enforce the ordering once this
slide is placed there).

**Four KPI tiles**, following the `<Name>Tile` / `<Name>TileValue` /
`<Name>TileLabel` convention every other tile row in this deck uses (delete
all three to drop a tile; not applicable here -- all four are always shown
whenever this slide exists at all):

| Tile | Value token | Suggested label |
|---|---|---|
| `PolkHouseholdsTile` | `{{POLK_MATCHED_HOUSEHOLDS}}` | "Matched Households" |
| `PolkSalesTile` | `{{POLK_TARGET_DEALER_SALES}}` | "Target Dealer Sales" |
| `PolkBuyRateTile` | `{{POLK_BUY_RATE}}` | "Buy Rate" |
| `PolkLiftTile` | `{{POLK_CAMPAIGN_LIFT}}` | "Campaign Lift" |

Both `POLK_MATCHED_HOUSEHOLDS` and `POLK_TARGET_DEALER_SALES` may render
with a `" (projected)"` suffix baked into the value string itself when the
rep's own "Project for match rate" toggle is on (e.g. "49,461
(projected)") -- the tile doesn't need its own conditional shape for this;
the fill code decides the string. `POLK_BUY_RATE`/`POLK_CAMPAIGN_LIFT` are
never suffixed.

**One caption naming the match rate**, a plain text box named
`PolkMatchRateNote`, token `{{POLK_MATCH_RATE_NOTE}}` -- a full sentence
(not a tile), e.g. "Based on a 90.49% match rate -- matched figures are a
floor, not the campaign's full reach." Place it directly under the tile
row, matching this deck's existing pattern of a short caption line under a
KPI row (see `PlanVsActualNote` on `report:delivery_breakdown` for the same
shape: plain text box, one sentence, sits below its related content).

**One table, `PolkTargetDealersTable`**, token `{{POLK_TARGET_DEALER_ROWS}}`
-- same "one header row + one template data row, clone per real row"
convention as every table in this deck. Three columns: Dealer | Market Rank
| Campaign Rank. Sized for up to 5 data rows
(`POLK_TARGET_DEALERS_ROW_CAP` in `report_assembly.py`) -- the real fixture
on hand only has 2 target dealers, so there's no example of the cap firing
yet, but a client with more target dealers needs the row budget reserved.
Give this table its own named header shape, `PolkTargetDealersHeader`
(e.g. "TARGET DEALER PERFORMANCE"), so the fill code can delete both
together on the rare campaign with zero rows to show (mirrors
`DeliveryByGeoHeader`/`DeliveryByCreativeHeader`'s own header-plus-table
pairing on `report:delivery_breakdown`).

**One narrative**, a plain text box named `PolkNarrative`, token
`{{POLK_NARRATIVE}}` -- one short, Python-computed sentence naming the top
audience/creative/publisher share of matched impressions (e.g. "AUTO Ford
Intenders led all audiences with 65% of matched impressions; the leading
creative was TG11075TBrittChev062630 (94% of matched impressions); Pluto TV
was the leading publisher (13% of matched impressions)."). Deliberately not
a chart or a second/third table -- Phase 7's own "digest, not dump" call.

**No `ChartRegion`/`MapRegion` on this slide** -- unlike most slides in this
deck, there's no image to place. If a future round wants a visual (a bar
chart of the audience/creative/publisher shares, say), that's a separate
ask; nothing in the current fill code expects one.

**Tokens summary for this slide:**

| key | Tokens |
|---|---|
| report:automotive_registrations | POLK_MATCHED_HOUSEHOLDS, POLK_TARGET_DEALER_SALES, POLK_BUY_RATE, POLK_CAMPAIGN_LIFT, POLK_MATCH_RATE_NOTE, POLK_TARGET_DEALER_ROWS, POLK_NARRATIVE |

## Handoff: v0_13 -- Polk v2 tiles/table (Phase 7's slide, rebuilt) plus two new shapes elsewhere

ATTRIBUTION_REPORT_PLAN.md Phase 8. Every piece of fill code (`report_
assembly._fill_automotive_registrations` and friends) is already built and
tested against the real `Polk Dashboard.xlsx` fixture, and is written to
degrade safely against the CURRENT v0_12 template (detected by shape name/
column count, not a version flag) until this handoff lands -- so nothing
breaks in the meantime, it just keeps rendering Phase 7's original layout.
Names below are confirmed, not proposed -- build as specified.

**report:automotive_registrations -- four tiles, replacing the current
four** (`PolkHouseholdsTile`/`PolkSalesTile`/`PolkBuyRateTile`/
`PolkLiftTile`):

| Tile | Value token | Suggested label | Change from v0_12 |
|---|---|---|---|
| `PolkSalesTile` | `{{POLK_TARGET_DEALER_SALES}}` | "Target Dealer Sales" | Kept as-is |
| `PolkMsrpTile` | `{{POLK_MSRP_SOLD}}` | "Total MSRP Sold" | **New** -- confirmed 2026-09-21: also carries its own "(projected)" suffix when the toggle is on, same as Sales/Households, so the three never disagree side by side |
| `PolkLiftTile` | `{{POLK_CAMPAIGN_LIFT}}` | "Campaign Lift" | Kept -- now dropped and the row reflows when lift is <= 0 |
| `PolkRoiTile` | `{{POLK_ROI}}` | "ROI" | **New** -- dropped and the row reflows whenever the rep's own ROI toggle is off |

`PolkHouseholdsTile` and `PolkBuyRateTile` are **removed outright** --
both figures move into the caption below instead.

**`PolkMatchRateNote`** (kept, same token `{{POLK_MATCH_RATE_NOTE}}`) now
composes a fuller sentence: matched households (with its own "(projected)"
suffix when the toggle is on), buy rate, and match rate together, e.g.
"44,756 matched households at a 90.49% buy rate and a 90.49% match rate --
matched figures are a floor, not the campaign's full reach." No shape
change needed -- same name, same slot, just more text.

**`PolkTargetDealersTable`, widened from 3 to 4 columns.** Dealer | Sales
| MSRP Sold | a fourth marker column (blank, "Client", or "Group" -- plain
text, no special formatting needed, though a bold/highlighted CLIENT row
would read well if it's easy to add). Same `{{POLK_TARGET_DEALER_ROWS}}`
token in the first cell, same one-header-row-plus-one-template-data-row
convention as every table in this deck. The table's own content changed
too (ranked by SALES across every dealer that sold to the exposed
audience, not just the advertiser's own target-dealer group ranked by
campaign rank) -- no template implication beyond the extra column, but
worth knowing why the numbers will look different in a real render.
`PolkTargetDealersHeader` is kept; consider renaming its visible label
from "Target Dealer Performance" to something that doesn't say "target"
any more, since the table now includes every dealer, not just the
advertiser's target group -- wording is your call.

**`PolkNarrative`** (kept, same token `{{POLK_NARRATIVE}}`) -- content
changed (never names a publisher any more; may now add a competitor/halo-
group sentence), no shape change.

**Two more shapes, on OTHER slides:**

| Slide | New shape | Token | Placement |
|---|---|---|---|
| `report:recap` | `PolkSalesWindowNote` | `{{POLK_SALES_WINDOW_NOTE}}` | A plain text box naming that Polk's own sales window is a separate period from the website attribution period stated on this slide (e.g. "Polk sales data: through Jul 31, 2026 (30-day sales window) -- a separate period from the website attribution above."). Deleted outright when no Polk file is attached to the report -- a template with this shape renders identically to today's v0_12 on every non-automotive report. |
| `report:highlights` | `CostPerVisitNote` | `{{COST_PER_VISIT_NOTE}}` | A plain text box under the highlights KPI tile row, same slot style as `PlanVsActualNote` on `report:delivery_breakdown` (a short caption line under a related row of numbers). Deleted outright whenever the rep's own "Include cost per visit" toggle is off, or there's nothing to show. |

**Tokens summary for the changed/new shapes:**

| key | Tokens |
|---|---|
| report:automotive_registrations | POLK_TARGET_DEALER_SALES, POLK_MSRP_SOLD, POLK_CAMPAIGN_LIFT, POLK_ROI, POLK_MATCH_RATE_NOTE, POLK_TARGET_DEALER_ROWS, POLK_NARRATIVE |
| report:recap | POLK_SALES_WINDOW_NOTE (new, in addition to the slide's existing tokens) |
| report:highlights | COST_PER_VISIT_NOTE (new, in addition to the slide's existing tokens) |
