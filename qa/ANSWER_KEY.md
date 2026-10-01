# QA fixture pack — answer key

Real seller inputs for the Proposal Builder (proposal side + Attribution Report Builder).
Copied from the developer's machine on 2026-10-01; nothing here is code, a draft fixture,
or a secret. **These are real client documents — don't share them outside the QA
engagement.**

## Where every expected figure comes from

Every number below is one of two things, and each is labelled:

- **[doc]** — printed in the file itself (a PDF's "Media Plan Details" header, a Wide
  Orbit totals row, an export's headline widget/tab, the Analyst file's own fields).
- **[test]** — a value the project's test suite already records as a fixture truth for
  that file (cited by test file).

Nothing was produced by running the app. Where a file has no external truth for an
output (drafted narrative wording, map rendering, chart shapes), it says **no answer
key**.

```
work_orders/      Wide Orbit broadcast schedules  -> Build a proposal, intake "Wide Orbit" uploader
avails/           Premion avails PDFs             -> Build a proposal, intake "avails PDF" uploader
attribution/      exports                         -> Attribution reports page
```

---

## work_orders/ — Wide Orbit schedules

**Seller use:** Build a proposal → intake area → upload the schedule. The app replaces
the rate-card "Broadcast Schedule" plan line with one imported broadcast line and adds a
"Broadcast TV | Media Plan" schedule slide.

All three: station **WUSA** → market **Washington, DC** (station→market is a fixed
lookup) [test: `tests/test_wideorbit.py`, CLAUDE.md]. Figures are the export's own
totals row [doc] as recorded in `tests/test_wideorbit.py` [test].

| File | Format | Spots | Gross cost | Impressions | CPM (export) | Reach | Freq | GRPs | Demo | Rows | Weeks | Flight |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `regency_planner.xls` | Planner .xls | 148 | $28,000 | 1,831,600 | $15.29 | 16.5 | 5.2 | 86.4 | CS-A25+ | 6 | 4 | 2025-05-05 – 2025-06-01 |
| `regency_planner.pdf` | Planner .pdf | 130 | $22,000 | 2,879,100 | $7.64 | 26.6 | 5.6 | 2879.1 | CS-A25+ | 3 | 3 | 2026-05-04 – 2026-06-01 |
| `ravens_campaign_schedule.xlsx` | Campaign Schedule Report | 14 | $79,500 | 1,785,413 | $44.53 | 65.5 | 1.8 | 114.81 | A25-64 | 12 | 21 | 2026-08-03 – 2027-01-31 |

Flight source: `regency_planner.xls` prints **"Plan Dates: 5/5/2025 - 6/1/2025"** [doc] —
the flight ends **2025-06-01**. Its "End Week: 5/26/2025" is the *start* of the last
week (May 26 – Jun 1), not the flight end (QA-008, closed By design).

Expected in the output:
- The imported broadcast line's cost ties to the gross cost above; its impressions tie to
  the impressions above. **Units trap:** the two Planner exports print impressions in
  thousands (e.g. 1,831.6 = 1,831,600); the plan must show the raw figure, never 1,832 or
  1.8 billion [test].
- The plan's CPM is derived from cost ÷ impressions, so it may differ from the export's
  printed (cent-rounded) CPM by a fraction of a cent [test/CLAUDE.md].
- Ravens crosses New Year: week columns roll from 2026 into 2027 in ascending order
  [test].
- `regency_planner.xls` contains $0.00 added-value rows that still carry spots; they are
  kept on the schedule, not dropped [test].
- With "Apply agency gross-up (×1.15)" ticked, the broadcast line's cost does **not**
  change (see RULES).
- A file that isn't a schedule gets a plain-language error, not a crash [test].

---

## avails/ — Premion avails PDFs

**Seller use:** Build a proposal → intake area → upload the avails PDF (working from an
avails document). Each audience × geography block becomes a targeting group in the D2
grid; the deck gets an avails slide and targeting map.

The core check for every document: **the groups' impressions sum to the header's own
"Total Impressions"** [doc], as recorded in `tests/test_avails_pdf_import.py` [test].

| File (RFPID) | Header Total Impressions [doc] | Groups | Geography | Flight | Other recorded truths [test] |
|---|---|---|---|---|---|
| Annapolis Cars (253813) | **2,522,716** | 8 (4 audiences × 10 mi and 5 mi) | radius, origin zip 21401 for all 8 | — | each group one period spanning the document flight |
| Visit Hershey & Harrisburg (260402) | **310,800,336** | 12 (8 DMA + 4 named-zip) | 2 audiences; Philly named zips = 52, NY named zips = 34, disjoint | — | advertiser reads "Visit Hershey & Harrisburg" (survives a line wrap); document order interleaves audiences |
| Wilmington University (253956) | **332,015,108** | 5 (one per audience) | County Option: same 10-county list, same **303** zips for all 5 | 2026-07-01 – 2027-06-30 | 12 monthly rows per group, full calendar months, month-to-month figures differ; zips 17527, 17555, 18015, 18042 are in the list |
| Capital Media / Undisclosed Advertiser (266994) | **3,321,409** | 1 | bare "County Option", **98** zips | 2026-09-21 – 2026-12-20 | 4 monthly periods: 9/21–9/30 **364,990**; Oct **1,131,469**; Nov **1,094,970**; 12/1–12/20 **729,980** |
| Lawn & Leisure (265521) | **509,868** | 1 | 10-mile radius, no origin printed | 2026-09-06 – 2026-10-11 | agency "Direct - No Agency"; audience "(DEMO Homeowner) AND (HH Income 200K Plus)" |
| Plaza Motors Group (266583) | **2,834,169** | 2 (A35-64, M35-64; same zip list) | named zips | 2026-09-17 – 2026-09-30 | each group one period spanning the flight |

Expected in the output:
- D2 shows exactly the group counts above, and the avails total ties to the header total.
- Every avails line carries the document's stated figure for its own dates until the rep
  clicks "Adjust to plan dates" (see RULES). Capital Media is the case to try this on:
  its flight starts and ends mid-month.
- Plan-line inclusion is opt-in: a freshly imported group is **not** on the media plan
  until its Plan box is ticked [CLAUDE.md].
- Targeting map dot/area rendering: **no answer key**.

The three header totals in the request (2,522,716 / 310,800,336 / 332,015,108) belong
to three different documents (Annapolis, Hershey, Wilmington), not one.

---

## attribution/nwfcu/ — Northwest Federal Credit Union

`Premion Website Attribution and Reach Extension (15).xlsx`

**Seller use:** Attribution reports → upload as the Website Attribution export → confirm
client → (no proposal, or link one) → Generate. No delivery/retargeting files exist for
this client in the pack, so the Delivery Recap and OTT Retargeting slides should be
absent, not empty.

| Figure | Expected | Source |
|---|---|---|
| Client | Northwest Federal Credit Union - Direct | [doc] advertiser tab |
| RFPID | RFPID-257030 | [doc] |
| Market | Washington, DC (Hagerstown) | [doc] |
| Impressions delivered (tile) | **902,036** | [doc] headline widget |
| Attributed impressions | 37,969 | [doc] |
| Attributed rate (tile) | **4.21%** (0.042093) | [doc] |
| Attributed unique visitors (tile) | **4,658** | [doc] |
| Unique visitor rate | 0.52% (0.005164) | [doc] |
| Homepage reach | **72%** of attributed visitors (a single page, so a true reach figure; narrative only, never a table column) | [test] `tests/test_report_assembly.py` (the real walkthrough figure) |
| Flight end | 2026-08-31 | [test] |

Drafted narrative wording: **no answer key** (check it against RULES instead).

---

## attribution/ted_britt_aug_2026/ — Ted Britt Ford & Ted Britt Chantilly, August 2026

**Seller use:** Attribution reports → upload the website export, the Delivery export
(`Premion OTT TB930.xlsx`), the OTT Retargeting export (`Audience Marketplace
TB930.xlsx`), the Auto-Sales Analyst facts JSON, and optionally the Analyst deck →
confirm client → Generate. Vertical: Automotive.

**Website export** (`Premion Website Attribution and Reach Extension TB930.xlsx`) [doc]:
client "Ted Britt Ford & Ted Britt Chantilly"; RFPID-261033; market Washington, DC
(Hagerstown); delivered **659,261**; attributed impressions 40,448; attributed rate
**6.14%** (0.061354); attributed unique visitors **2,671**; three audiences — AUTO Ford
Intenders 438,964 / AUTO Make Chevrolet 110,091 / AUTO Make Lincoln 110,206 delivered.

**Delivery export** (`Premion OTT TB930.xlsx`) [doc]: delivered **659,261**; VCR
**98.5%** (0.984853); frequency **8.7**; uniques **75,404**; top publisher Pluto TV
93,305; four 5-mile-radius geographies (4175 Auto Park 219,487; 11165 Main St 219,477;
4171 Auto Park Cir 110,206; 46990 Leesburg Pike 110,091).

**Retargeting export** (`Audience Marketplace TB930.xlsx`) [doc]: impressions
**200,204**; clicks **378**; CTR **0.19%**; ad size Video; Mobile 340 of 378 clicks.
Because retargeting is uploaded, no "add OTT Retargeting" recommendation should appear.

**Analyst facts JSON** (`auto-group-...-08-31_facts.json`) [doc]: analysis period
2026-08-01 – 2026-08-31; scanned 2026-09-30 (30 days after period end — inside the
Analyst's own 15–30 day guidance, so **no** late-scan warning); 5 sites; vehicles
shopped **1,696**; shopped vehicles since sold **918**; Look-to-Book **54.1%** (new
**41.0%**, used **73.0%**); Est. revenue sold **$38,063,252**; Pipeline Value
**$75,200,770**; unique visitors **2,671** (equals the website export — so **no**
visitor-mismatch warning); visits total 8,694; Traffic Mix New VDP 2,818 / Used VDP
2,113 / Homepage 1,072 / New Car Search 836 / Other 778 / Used Car Search 483 / General
Search 309 / Service 188 / Incentives/Offers 81 / Online Conversions 16; top model Ford
F-150 98 sold; top Missed Opportunity Ford Mustang Dark Horse SC, 58 visits.
Recorded cross-checks [test `tests/test_analyst_import.py`]: the export's vehicle pages
total 1,696; new/used shopped split 1,000 / 696 reproduces the file's own 41.0% / 73.0%.

Expected in the generated report:
- "Where Visitors Went" is filled from the Analyst: subtitle "Of the 1,696 vehicles
  attributed visitors shopped, 918 have since sold."; TRAFFIC MIX and TOP SOLD MODELS
  tables; top model row Ford F-150 / 98; traffic rows use the Analyst's names (New VDP
  2,818, Used VDP 2,113, …).
- Store comparison, if narrated: Chantilly Look-to-Book 58.6% vs Chevrolet 41.5%
  [doc: JSON by_site].
- Without the JSON (website export only, vertical Automotive), the fallback "Where
  Visitors Went" uses the same category names and counts as the Analyst's Traffic Mix
  above for every category except "Other" — the app splits Finance / Credit App and
  Trade-In out of the Analyst's Other, so Other pages + Finance + Trade-In = 778 [test]
  — and no row shows a raw path such as "Searchnew.Aspx" [test].
- Drafted narrative wording: **no answer key** beyond RULES.

**Analyst deck** (`Auto Group (5 Sites) - 12_07 PM ET_Summary.pptx`) [doc]: 12 slides;
its own cover prints Units Sold 918, Est. Revenue Sold $38,063,252, Pipeline Value
$75,200,770, Look-to-Book 54.1% (New 41.0% | Used 73.0%). When appended, all 12 slides
appear in order, **before** Takeaways on build b3b0e60 and later (see Rule 5).

---

## RULES — checkable from the output alone

1. **Agency gross-up (×1.15) moves CPM and cost, never impressions.** Compare the deck
   preview table (and the generated deck's media plan) with "Apply agency gross-up
   (×1.15)" on and off: every non-broadcast line's impressions are identical; its CPM and
   cost are ×1.15. The editable plan grid itself always shows net figures and does not
   change — that's by design, not a failure.
2. **Broadcast always stays gross.** An imported Wide Orbit line's cost equals the
   schedule's own gross cost (e.g. $28,000 for `regency_planner.xls`) with the gross-up
   on or off.
3. **Budget doesn't prorate for partial months.** A Monthly-breakout line's monthly cost
   is the same in a partial first/last month as in a full one (try a flight like Capital
   Media's 9/21–12/20); no line is scaled by a day fraction.
4. **An uploaded avail holds at its stated figure unless the rep opts in.** After import,
   each avails line shows the document's own figure for the document's own dates, even
   when the plan's flight differs; it changes only after the rep clicks "Adjust to plan
   dates" (and "Use document figure" restores it exactly).
5. **Takeaways is the last slide** of every attribution report — after Polk and any
   appended Analyst slides. **Build note:** this rule applies on build **b3b0e60** and
   later (the build stamp is on the login screen). Build **51c808a** is the last build
   where it fails: there, Polk and an appended Analyst deck still land after Takeaways.
6. **No doubled client names** — e.g. never "WAEPA WAEPA CTV Strategy" on the plan
   title, or anywhere else.
7. **No slide shows a 0-avails placeholder** — no avails table row, group or market
   reading 0 impressions or a blank "no target market" placeholder.

Also visible from the output (attribution reports):
- No share in a "% of visits" column exceeds 100%, and a Traffic Mix table's rows add to
  ~100%.
- Analyst figures are always described as vehicles our audience viewed that have since
  sold — never as sales the campaign made, never "shopped" — and never added to or
  compared with Polk sales. Every Analyst dollar figure says "estimated"; Est. total
  value viewed (formerly Pipeline Value) never shares a sentence with Est. value sold
  or a sold count. (Rule wording updated 2026-10-01 for the Analyst's schema v2; the
  Ted Britt v1 figures above are unchanged — they describe that v1 file.)
- No VIN appears on any client slide.
