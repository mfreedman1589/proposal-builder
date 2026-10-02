# QA rules — what a tester can check from the app and its output

Each rule is client- or rep-facing and observable without reading code. The source names
the section of `CLAUDE.md` (or `DECISIONS.md`) that owns it — update this file in the same
commit as any change to one of these rules.

## Build a proposal — media plan and money

| # | Rule | How to observe | Source |
|---|---|---|---|
| P1 | **Agency gross-up (×1.15) moves CPM and cost, never impressions.** It applies once, order-wide. | Deck preview table / generated plan with the box on vs off: impressions identical, CPM and cost ×1.15. The editable grid always shows net — by design. | CLAUDE.md › The media plan and the form › "The agency gross-up…" |
| P2 | **Broadcast lines are always gross** — never grossed up again. | An imported Wide Orbit line's cost equals the schedule's gross cost with the gross-up on or off. | CLAUDE.md › Broadcast and Wide Orbit › "Broadcast lines are always gross…" |
| P3 | **An imported Wide Orbit schedule replaces the rate-card broadcast line** — never a second one. Station→market is a fixed lookup (WUSA→DC, WPMT/FOX43→Harrisburg). | One broadcast line on the plan after import; its geo matches the station's market. | CLAUDE.md › Broadcast and Wide Orbit › "An imported schedule replaces…" |
| P4 | **Wide Orbit impressions are raw on the plan** — never in thousands. | Planner exports print 1,831.6; the plan shows 1,831,600. | CLAUDE.md › Broadcast and Wide Orbit › "…the units are the trap." |
| P5 | **A flat fee is never multiplied by the month count.** | A Flat Fee row's full-flight cost equals its stated amount. | CLAUDE.md › The media plan and the form › "A media-plan row can be Rate or Flat Fee…" |
| P6 | **Budget doesn't prorate for partial months.** A Monthly-breakout row's monthly cost is the same in a partial first/last month as in a full one. | Flight starting mid-month: no line scaled by a day fraction. | BACKLOG.md › Smaller / carry-over › "Day-prorating…out of scope" |
| P7 | **Changing the flight after a plan is priced never blocks Generate and needs no dismiss step.** A Monthly plan keeps its monthly budget and its total follows the months (3 months at $1K/month is $3K; shorten to 2 and it's $2K; partial months follow the month-range rules). A Full-Flight plan keeps its total spend, spread across the new dates. A short note says what happened: "Flight now 2 months — plan total $2,000 (was $3,000)" or "Total spend held at $3,000 across the new dates". (Replaces the old warn-and-block rule; resolves QA-005.) | Priced Monthly plan, shorten the flight → total drops by the removed months, note shown, Generate enabled. Same on a Full Flight plan → total unchanged, "held" note. | Matt's decision, 2026-10-02 (QA-005) |
| P8 | **A $0 media plan row never reaches the deck** (a freshly seeded rate row may sit at $0 on the grid until priced). | No $0 line in a generated plan. | CLAUDE.md › The media plan and the form › "A $0 media plan row is always wrong" |
| P9 | **One media plan option renders exactly like no options**; with 2–3, each option's name is appended to its plan title. | One option: plain title. Two: each plan slide titled with its option name. | CLAUDE.md › The media plan and the form › "A proposal carries 1–3 media plan options…" |
| P10 | **The plan title never repeats the client name** ("WAEPA WAEPA CTV Strategy"). | Plan slide title. | CLAUDE.md › The media plan and the form › "The plan slide title never echoes the client name twice" |
| P11 | **A plan row's Geo cell shows market names, not the Geography paragraph.** | Geo column reads "Washington, DC, Baltimore", not narrative text. | CLAUDE.md › The media plan and the form › "A plan row's Geo cell prefers the target markets…" |
| P12 | **One custom (non-RFP-selectable) audience per campaign** — a second is warned about, never silently dropped. | Adding a second custom segment shows a warning. | CLAUDE.md › The media plan and the form › "One custom (non-RFP-selectable) audience…" |
| P13 | **Share of voice (% of avails) is per line and only where the line has real avails** — never a deck-wide figure, never 0% or ∞%. | With "show % of avails" on: matched lines show "(X% of avails)", unmatched lines show nothing. | CLAUDE.md › The media plan and the form › "% of avails (share of voice) is per LINE…" |

## Build a proposal — avails documents and targeting

| # | Rule | How to observe | Source |
|---|---|---|---|
| A1 | **An imported avails document's groups sum to the document's own header Total Impressions.** | D2 total vs the PDF's "Total Impressions". | tests/test_avails_pdf_import.py; CLAUDE.md › Status › Avails PDF importer |
| A2 | **An uploaded avail holds at its stated figure for its own dates** until the rep clicks "Adjust to plan dates"; "Use document figure" restores it exactly. | Plan flight differs from the document's → avails unchanged until the rep opts in. | CLAUDE.md › The media plan and the form › "An avails-PDF import's figures freeze…" |
| A3 | **An import fills a header field (client name, flight) only when the field is empty or still holds what an earlier import wrote**; otherwise it leaves the rep's value and reports a conflict. | Type a client name, then import → name kept, conflict note shown. | app.py `_avails_import_field` (precedence: rep edits > document > draft > defaults) |
| A4 | **A freshly imported targeting group is not on the media plan until its Plan box is ticked** (or "Add all to plan"). | Import → no Premion Streaming TV lines until ticked. | CLAUDE.md › The media plan and the form › "Targeting groups own their Premion Streaming TV line…" |
| A5 | **Removing a hand-edited plan line asks first**, naming the numbers ("Remove them and discard those numbers" / "Keep them on the plan"). | Untick Plan on a group whose line was edited. | same section |
| A6 | **With "Working from an avails document" off, no avails slide and no 0-avails placeholder reach the deck.** | Generate with the toggle off: no avails table/slide, no blank placeholder group. | CLAUDE.md › Deck assembly › "Working from an avails document (avails_mode) is authoritative…" |
| A7 | **Whenever the personalized avails slide is in the deck, the vertical's static Precision Targeting slide is not.** | Never both in one deck. | CLAUDE.md › Deck assembly › "No part of the app depends on slide numbering…" |

## Build a proposal — deck content

| # | Rule | How to observe | Source |
|---|---|---|---|
| D1 | **No section-divider slides in a generated proposal.** | Slide list. | CLAUDE.md › Deck assembly › "No section dividers…" |
| D2 | **Total TV branding follows the market** — no competitor station's branding in a deck (DC deck never shows FOX43/WPMT, Harrisburg never shows WUSA9). | Cover, plan template, schedule slide. | CLAUDE.md › Deck assembly › "Total TV drives the deck's branding…" |
| D3 | **Case studies sit immediately before the media plan slide**, and only ones the rep can see checked are included. | Slide order; case-study picker. | CLAUDE.md › Deck assembly › "Case studies go immediately before the media plan slide…" |
| D4 | **Auto replaces the generic "Measure Sales Conversions" slide with its own attribution slide (Polk); travel adds Arrivalist without replacing it.** | Auto vs travel decks. | CLAUDE.md › Deck assembly › "Auto and travel sell their own attribution product…" |
| D5 | **Nothing overlaps on the plan slide**; a plan that can't fit warns instead. | Visual check of the plan slide. | CLAUDE.md › Layout, fitting and text measurement |
| D6 | **The "© PREMION" image slide and the reporting-metrics slide are in every proposal.** By design — not a finding. | Slide list of any generated proposal. | Matt's decision, 2026-10-02 |
| D7 | **The plan title never repeats the client name**, however the title spells it (case, punctuation, "&"/"and", a prefix like "QA-TEST-" the title lacks). | Plan slide heading. | app.py `option_plan_title` |
| D8 | **The avails slide carries one caption under the table: "Available monthly impressions by market — your plan buys N/month."** N is the plan's own monthly impressions (one figure per option, each named). | Avails slide. | Matt's decision, 2026-10-02 |

## Attribution reports

| # | Rule | How to observe | Source |
|---|---|---|---|
| R1 | **Takeaways is the last slide** — after Polk and any Auto-Sales Analyst slides (build b3b0e60 and later). | Slide order. | ATTRIBUTION_REPORT_PLAN.md › Ted Britt review, 2026-10-01 |
| R2 | **A slide with nothing to show is absent, never thin**: no Delivery slides without a delivery file, no OTT Retargeting slide without that export, no Polk slide without Polk, no Analyst slides without a usable Analyst file. | Generate with each optional file missing. | CLAUDE.md › Status › Attribution Report Builder |
| R3 | **"% of visits" columns are shares of page visits: each ≤ 100% and the Traffic Mix rows add to ~100%.** A reach figure ("of attributed visitors") appears only in sentences, never in a column. | Where Visitors Went tables. | CLAUDE.md › Status › Attribution Report Builder ("every share on Where Visitors Went is a share of page visits") |
| R4 | **No raw page paths on a client slide** (no "Searchnew.Aspx"); auto sites show vehicles and the Analyst's category names. | Where Visitors Went. | ATTRIBUTION_REPORT_PLAN.md › Fallback "Where Visitors Went" |
| R5 | **Every number in drafted text traces to the report's own figures** (a plain rounding counts). An untraceable number never ships — its sentence is removed and "Review before sending" says so. | Review panel after Generate; narrative numbers vs tables. | CLAUDE.md › Status › Attribution Report Builder (facts-only contract); QA-007 |
| R6 | **Auto-Sales Analyst figures (only) are inventory movement, never attributed sales** -- this rule covers the Analyst's sold-vehicle data and nothing else: website-attribution language ("drove X website visits", attributed visits, conversions) is standard and correct, Polk is matched-sales data and may be called sales, and case studies keep their own wording. "Of the N vehicles our audience viewed, M have since sold." The word is "viewed" -- never "shopped"/"shoppers", never "pipeline", no influence claim from visit counts. Never added to or compared with Polk; every Analyst dollar figure "estimated"; Est. total value viewed is quoted alone, never in a sentence with Est. value sold or a sold count; the two dollar tiles are never adjacent. | Inventory Movement slide, highlights, takeaways. | CLAUDE.md › Status › Attribution Report Builder; Sales Assist |
| R7 | **An Auto-Sales Analyst make/model trend at a store that sells that make is never credited to an audience** (Ford movement at a Ford store isn't proof for a Ford-intender audience). Analyst data only. | Takeaways. | ATTRIBUTION_REPORT_PLAN.md › "Same day, second round" |
| R8 | **Missed Opportunities is its own What's Next item**; Dynamic Ads, when suggested, names the specific vehicles under "Ideas to consider". Never "add OTT Retargeting" when a retargeting export was uploaded. | What's Next. | ATTRIBUTION_REPORT_PLAN.md › REPORT_MASTER_v0_15 / Third round |
| R9 | **A list is referred to by name, never by position** ("the Missed Opportunities watch list", not "…below"). | Narrative/What's Next text. | ATTRIBUTION_REPORT_PLAN.md › REPORT_MASTER_v0_15 |
| R10 | **No VIN on the report builder's own (native) slides.** The Auto-Sales Analyst's own deck, appended only when the rep ticks it, keeps its VIN detail by design — the Sales Assist uses it for spot-checks — so a VIN on those appended slides is not a finding (QA-009). | Native Analyst slides (Inventory Movement, Missed Opportunities, Store Scoreboard). | ATTRIBUTION_REPORT_PLAN.md › Handoff: native Analyst slides; Matt, 2026-10-02 |
| R11 | **An Analyst file whose period doesn't overlap the report is set aside with a warning naming both periods**; a 0-sold file is set aside; a scan more than 60 days after the period end warns and appears in Review before sending. | Upload page warnings. | ATTRIBUTION_REPORT_PLAN.md › Auto-Sales Analyst facts JSON |
| R12 | **No client-facing text describes our own inputs** ("no goals on file", "this tab was incomplete"). | All narrative text. | app.py drafting prompt › "No goals supplied" rule |

## Everywhere

| # | Rule | How to observe | Source |
|---|---|---|---|
| E1 | **No client name appears doubled** anywhere ("WAEPA WAEPA"). | Titles and headings. | as P10 |
| E2 | **Every warning on the page is something the rep can act on now**; internal/dev detail goes to the feedback export, not a banner. | Page after Generate. | CLAUDE.md › Running things › "Walk the page as a rep…" |
| E3 | **The build stamp on the login screen and sidebar names the running commit.** | Login screen. | app.py build stamp (sidebar and login screen) |
| E4 | **The draft's "Before sending" checklist never contradicts the table or plan** — no audience listed as both on the avails table and dropped, no "not on the avails table" for an audience that's there, no "has no rows" for a market that has them. | Review notes after Draft from notes, against D2 and the plan. | app.py `drop_contradicted_draft_notes` |
| E5 | **"Clear the form" clears everything except who you are and which page you're on** — client name, flight, notes, Campaign Specs, avails, plan and any warning. | New proposal → Clear the form. | QA-006 |
| E6 | **Dismissing a date picker (Escape) never changes the stored date.** | Set a flight date, press Escape on the open calendar, click elsewhere → date still set. | QA session 2026-10-02 |
| E7 | **Proposal history lists each option's own total with its name, never the sum of options.** | A 2-option proposal in Proposal history. | Matt's decision, 2026-10-02 |
