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
