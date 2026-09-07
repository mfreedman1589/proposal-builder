# Attribution Report Builder — phase plan

**Handoff to Claude Code. Read this alongside CLAUDE.md, DECISIONS.md,
`attribution-module-framework.md` (Matt's original framework) and
`ATTRIBUTION_ROADMAP_ADDENDUM.md` (phasing/template-rule/case-study additions
on top of it). Approved 2026-09-05 with four corrections from the kickoff
draft — this file is the corrected version; nothing below should be
re-litigated without new evidence.**

## Why

A rep uploads Premion Website Attribution (+ optional delivery) Excel
exports and gets back a client-ready report deck — the opposite of the
templated deck it replaces, which buries the result under delivery detail
and internal nomenclature (see the real Cardinal PDF this replaces, and the
hand-rebuilt Mattress Warehouse deck that proved the template rules in
`ATTRIBUTION_ROADMAP_ADDENDUM.md` §2). Two modes (with/without a linked
proposal), eventually two cadences and an optimization engine — this plan
covers only report generation (framework build-order step 1 / addendum
steps 1-2). Optimization engine, proposal loop-back, case-study-from-report
and every add-on (Polk/auto, sales match-back, brand lift, Arrivalist) are
named at the bottom as deferred, not scoped here.

Real data grounding every decision below: `MW attribution excel.xlsx`
(Mattress Warehouse's attribution export) + `Mattress Warehouse - Final
Report - May-July 2026 (v5) (2).pptx` (the hand-built wrap this feature
should approximate) are the full-wrap reference; `Premion Website
Attribution Cardinal.xlsx` + `Premion OTT.xlsx` + the old templated
`Premion Streaming Campaign_Cardinal Plumbing...pdf` are the report this
replaces, and the monthly-shape sample. All are checked into the repo root
(gitignored, like every other real client file here).

**Correction (2026-09-06, Phase 3):** `Premion OTT.xlsx` is Cardinal
Plumbing's delivery export, not Mattress Warehouse's — this file
previously (and the kickoff memory) had that backwards. Confirmed by
parsing it: its RFPID (256286) and delivered impressions (917,451) match
Cardinal's own attribution file almost exactly, and its flight-detail tabs
name "Cardinal Plumbing" on every row, never Mattress Warehouse. **There is
no real Mattress Warehouse delivery export in the repo** — Correction 3's
"MW: 3.1M delivery-file delivered vs. 2.34M attribution-file delivered"
scenario below has no matching real fixture; see DECISIONS.md's Attribution
Report Builder section for the full finding. Phase 3 used MW
(attribution-only) and Cardinal (a real attribution+delivery pair) as its
two walking-skeleton fixtures instead, which between them still exercise
both the delivery-set-present and delivery-set-absent paths.

## Corrected decisions (read before Phase 3)

**1. Report master deck dimensions: 13.333 × 7.5in — matching the proposal
master exactly, NOT the MW rebuild's own 14.167×8.0in.** That size was an
accident of whatever template the rebuild started from, not a deliberate
choice. Reason to match: report slides get copied INTO proposal decks later
(create-case-study-from-report, the slide vault) — same dimensions removes
scale/clip bugs from that path before they exist. This does **not** reopen
the separate-deck decision (Q2 in the kickoff answers) — that still stands
on storage headroom (~4MiB left on the proposal master against its 50MiB
ceiling) and slide-key/token collision risk alone, independent of size.

**2. Two report shapes, decided by which files are uploaded — no toggle.**
Both upload slots (attribution export, delivery export) show by default;
the delivery export is optional. Both files → the full report, delivery
slide included. Attribution only → the delivery slide is **omitted
entirely**, not shown thin or partial. Consequences that reach the code:
- No attribution-only fallback rendering for delivery-only fields. A slide
  that needs the delivery file exists only when the delivery file does.
- Phase 1's parsed-data structure marks delivery data present-or-absent at
  the **file** level (e.g. `delivery: DeliveryExport | None`), not
  per-field `None`-checks bolted on later.
- The slide inventory below marks exactly which slide belongs to the
  delivery set (there is exactly one: Delivery Recap).

**3. Headline impressions = the delivery file's figure, when one exists.**
When the two files disagree (they will — MW: 3.1M delivery-file delivered
vs. 2.34M attribution-file delivered, because the attribution file's count
is scoped to what could be pixel-matched, not total delivery), the delivery
figure is what the client sees as "impressions delivered." The attribution
file's own delivered-impressions count is the **rate-math denominator only**
and stays internal — never shown as a second, competing total anywhere in
the deck. With no delivery file, the attribution figure is the only number
and becomes the headline — and since there's no Delivery Recap slide to
carry it, it lands on the Highlights slide instead. **Required test**
(Phase 1 or Phase 4, whichever lands the render first): building MW's
report end to end must show 3.1M as delivered impressions somewhere in the
rendered deck, and 2.34M must never appear as a delivered total anywhere in
it.

**4. Claude does not design the report master PowerPoint template — Matt
does, using the MW rebuild as visual reference.** Claude's job is the exact
spec below (slide inventory, delivery-set marking, tokens, what each token
expects) plus, once Matt hands back a built deck, tagging every slide with
a `key:` notes label (reusing `slide_map.notes_key()` — no anchors, no
carry-forward, this deck is small enough to hand-tag completely) and wiring
`assembly.py`'s existing fill primitives against it. **Phase 3 does not
start until Matt delivers the template deck built from this spec.**

## The slide inventory (Matt's build spec)

Build at **13.333 × 7.5in**, matching the proposal master. Seven slide
families below. Each gets a `key: report:<name>` line in its speaker notes
(same mechanism as the proposal master — see `SLIDE_KEYS.md`) — that's
Claude's tagging step once the deck exists, nothing you need to add
yourself. Every image-region shape (a chart or map) should be a distinctly
**named** rectangle/picture placeholder (rename it in PowerPoint's
Selection Pane — e.g. `ChartRegion`, `MapRegion`) so Python can find it by
name rather than by guessing position.

Token types: **string** = single run-level `{{TOKEN}}` replace. **bullets**
= a bulleted text box, one clone per extra item — for a two-run bullet
(bold headline + regular detail in one bullet, like the MW rebuild's own
"What You're About to See" slide), make the template's single bullet
paragraph carry two runs, one wrapping `{{X_HEAD}}` in bold and one
wrapping `{{X_DETAIL}}` plain, so both get filled together per bullet.
**table** = a table whose first data row is the fill template; extra rows
clone it. **image** = a named empty region a PNG gets placed into.

---

### 1. Campaign Recap — `report:recap` — always present

| Token | Type | Expects |
|---|---|---|
| `{{CLIENT_NAME}}` | string | Advertiser name |
| `{{REPORT_TITLE}}` | string | e.g. "Website Attribution Report" |
| `{{REPORT_PERIOD_LABEL}}` | string | The period this specific report covers, e.g. "September 2026" or "May 22 – Jul 17, 2026" |
| `{{FLIGHT_LABEL}}` | string | The full campaign flight, when it differs from the report period (a monthly report on a longer campaign) |
| `{{GOALS_BULLETS}}` | bullets | 2-4 items — from the linked proposal's goals, or drafted from notes with no proposal |
| `{{AUDIENCE_BULLETS}}` | bullets | 1-4 items — from the proposal's audience stack, or the attribution export's own Audience Name cuts with no proposal |
| `{{GEOGRAPHY_LABEL}}` | string | Target market(s) |

One design note for the build: the goals/budget area needs to read fine
when **budget is absent** (no-proposal mode never invents a dollar figure)
— either skip a dedicated budget box entirely, or make it a small element
that can just be left blank/deleted per report rather than something the
fill code has to fake a value for.

### 2. Top Highlights — `report:highlights` — always present, content varies

Per the addendum's template rule: KPIs over charts here, not a chart.
**Revised 2026-09-06: no KPI strip.** Three headline tiles plus four
head/detail bullets is the whole slide — a strip duplicated the tiles, and
the delivery-only figures (VCR, frequency, uniques) already have a home on
the Delivery Recap slide when one exists.

| Token | Type | Expects |
|---|---|---|
| `{{HEADLINE_IMPRESSIONS}}` | string | Delivery file's figure when present, else the attribution file's own (Correction 3) |
| `{{HEADLINE_UNIQUE_VISITORS}}` | string | Attributed unique visitors |
| `{{HEADLINE_ATTRIBUTED_RATE}}` | string | Percentage string — `attribution_import`'s `attributed_rate` is a raw fraction as Excel stores it (e.g. `0.01363`), formatted ×100 to 2 decimals: `"1.36%"` |
| `{{HIGHLIGHT_HEAD_1..4}}` / `{{HIGHLIGHT_DETAIL_1..4}}` | bullets (2-run) | Up to 4 items, Claude-selected/worded from the Python-computed facts payload — never a number the payload doesn't contain |

### 3. Delivery Recap — `report:delivery_recap` — **delivery set: dropped entirely when no delivery file is uploaded**

| Token | Type | Expects |
|---|---|---|
| `{{DELIVERED_IMPRESSIONS}}` | string | |
| `{{VCR}}` | string | |
| `{{FREQUENCY}}` | string | |
| `{{UNIQUES}}` | string | Delivery file's own Uniques (distinct from attributed unique visitors) |
| `{{TOP_PUBLISHERS_ROWS}}` | table | Top 5, (Channel, Delivered Impressions, VCR) -- corrected 2026-09-06 from "Top 10," which was wrong for a slide carrying two tables (see the Delivery Recap row below and DECISIONS.md) |
| `{{CREATIVE_ROWS}}` | table | Top 3, (Creative Name, Delivered Impressions, completion/VCR) -- row cap added 2026-09-06, same reasoning as TOP_PUBLISHERS_ROWS above |
| `{{DELIVERY_NARRATIVE}}` | string | One Claude-drafted sentence |
| chart region | image | Publisher delivery bar chart (`report_charts.py`) |

### 4. Website Attribution Breakdown — `report:attribution_breakdown` — always present

**Revised 2026-09-06: one table, one chart, never conditional tables.**
Building a slide that has to look right with 0, 1 or 2 of three tables
deleted is three layouts wearing one slide's clothes — the MW rule is show
what matters, so Python picks the single dimension worth telling, rather
than the template carrying three candidate tables and hiding the ones that
don't apply. Default: by market when the export has more than one market
(`result.by_market`), else by audience (`result.by_audience`). Creative
(`result.by_creative`) gets the slide only when it's genuinely the story —
one creative dramatically outperforming another — which is a judgment
call in the same facts-payload synthesis step that picks the highlight
bullets, not a fixed count rule like market's. Whichever dimension is
picked, it's the ONLY breakdown table on this slide; the others still
exist in the parsed data for the narrative/highlights to cite from, they
just don't get their own row set here.

| Token | Type | Expects |
|---|---|---|
| `{{ATTRIBUTION_HEADLINE_NOTE}}` | string | |
| `{{BREAKDOWN_DIMENSION_LABEL}}` | string | What the table is showing — "By Market" / "By Audience Segment" / "By Creative" |
| `{{BREAKDOWN_ROWS}}` | table | Dimension-agnostic 3 columns: (name, impressions, attributed rate) — same shape whichever dimension is picked, so the template needs exactly one table, not one per dimension. Give the first column a generic static header in the build (e.g. "Segment") rather than a token, since it reads fine for market, audience or creative alike |
| `{{ATTRIBUTION_NARRATIVE}}` | string | |
| chart region | image | Bar chart of `BREAKDOWN_ROWS`' own dimension |

### 5. URL Report — `report:url_report` — always present

The framework doc calls this "the richest analysis opportunity" — per the
addendum's template rule, this is the slide that earns the most narrative
space.

| Token | Type | Expects |
|---|---|---|
| `{{URL_HEADLINE_NOTE}}` | string | |
| `{{TOP_URL_ROWS}}` | table | Grouped by meaningful path segment (Python groups raw URLs into labeled buckets — "Mattress category pages," "Store locator," etc — never raw URL strings, which run up to 186 characters and are meaningless to a client) |
| `{{URL_INTENT_NARRATIVE}}` | string | Claude's "what this signals about intent" synthesis, tied back to campaign goals when a proposal is linked |

### 6. Zip-Code Analysis — `report:zip_analysis` — always present

| Token | Type | Expects |
|---|---|---|
| `{{ZIP_HEADLINE_NOTE}}` | string | |
| `{{TOP_ZIP_ROWS}}` | table | (Zip, area label, impressions/rate vs. campaign average) |
| `{{ZIP_NARRATIVE}}` | string | |
| map region | image | Visitor zips only, no proposal linked; visitor zips + targeted zips overlaid (two colors) when a proposal is linked — reuses `targeting_map.render_map` |

### 7. Takeaways / What's Next — `report:takeaways` — always present, depth varies

| Token | Type | Expects |
|---|---|---|
| `{{TAKEAWAY_HEAD_1..4}}` / `{{TAKEAWAY_DETAIL_1..4}}` | bullets (2-run) | 2-4 items for a monthly report, up to 4 for a wrap — same template slide either way, only the count/depth differs |
| `{{WHATS_NEXT_BULLETS}}` | bullets | Extension/expansion/growth items |

No separate monthly-vs-wrap slide variant is needed anywhere in this
inventory — every difference between a monthly report and a wrap is content
depth (how many bullets, how much synthesis), decided in Python/the prompt,
never a different template slide.

---

## Phase plan

Same discipline as `FLOW_REWORK_PLAN.md`: small, independently green
commits; every phase verified against **both** real datasets (MW wrap,
Cardinal monthly); a full `tests/run_all.py` sweep before any phase is
declared done. New pure modules follow the existing
`avails_pdf_import.py`/`wideorbit.py`/`notes_file_import.py` shape (no
Streamlit, no DB, never raises). Real fixture files are gitignored; tests
report `SKIP` rather than fail when absent, matching `test_wideorbit.py`.

### Phase 1 — Parse both exports, no UI, no deck (starting now)

`attribution_import.py`: `parse_attribution_export(path)` and
`parse_delivery_export(path)`, tabs identified by header row
(BOM/whitespace-stripped), the date-tab disambiguation-by-day-gap rule
(1-day = trailing window, 7-day = weekly trend, ~30-day = monthly rollup),
the fuller-vs-ranked tab preference, multi-RFPID rejection with a
plain-language error. Output is a structure that represents "no delivery
file" as a top-level present/absent state (Correction 2), never a maze of
per-field `None` checks. Covers every field the slide inventory above
needs. `tests/test_attribution_import.py` against real MW + Cardinal +
delivery fixtures (`SKIP` without them), plus synthetic-shape tests proving
the disambiguation rules without needing the GLS/Twin-Pine files checked
in.

### Phase 2 — Page shell, two entry doors, advertiser matching (no deck output yet)

**Two entry points into the same page, added together — the second is the
common case, not an alternative UI:**

- **Upload-first** (the original door): rep opens Attribution Reports cold,
  uploads the export(s), and the page fuzzy-matches an advertiser/proposal
  for them to confirm. Stays exactly as originally scoped, for a rep who
  doesn't already know which proposal this is.
- **Build report from proposal** (new, added at Matt's request): a "Build
  report" button on each Proposal History row (`_render_proposal_row`,
  alongside "Load into form"/"Rebuild as presented") opens Attribution
  Reports **pre-linked** to that proposal — client, flight, plan and
  targeting already known, via the same `goto_*`-flag-plus-`st.rerun()`
  cross-page jump every other page-to-page handoff in this app uses. The
  upload step then **verifies** against the already-linked proposal rather
  than searching for a match: "export says Cardinal Plumbing, 9/1-9/30 —
  matches" or a named disagreement (advertiser id mismatch, or a flight
  that doesn't overlap) surfaced the same way every other precedence
  conflict in this app is (a visible note, never a silent override).

**A canonical `advertisers` table, built now rather than deferred** — see
the standalone rationale above; this is where it lands. `advertisers` (id,
canonical_name, name_key, created_at) — the exact shape of `team_members`,
including its `name_key = lower(trim(name))` dedup convention.
`proposals.advertiser_id` and `attribution_reports.advertiser_id` are
nullable FKs (`alter table add column if not exists`, `on delete set
null`) — existing rows stay NULL and keep showing `client_name` as typed;
nothing breaks. A one-time backfill script (shape of `build_market_
lookup.py`) groups existing `proposals` rows by normalized `client_name`
into advertiser rows and back-fills `advertiser_id`. The SAME resolve-or-
create-with-confirmation matcher (name normalization + `difflib`, narrowed
by the `STATION_MARKETS` pixel hint, disambiguated by flight-date overlap
against candidates' `form_json`) is used in three places: the upload-first
door's advertiser match, the build-from-proposal door's verification check,
and (new, small, worth doing while this matcher already exists)
`client_name` entry on the Build-a-proposal page itself, so a future
proposal's client name resolves to the same advertiser id from the start
instead of adding to the free-text pile this table exists to fix.

New nav entry (`"Attribution reports"`), `attr_`-prefixed widgets + a
`NON_PERSISTABLE_PREFIXES` entry, both file uploaders shown by default
(delivery optional) following `render_avails_pdf_uploader`'s five-step
idiom. New `attribution_reports` table (jsonb-bag shape copying `feedback`
exactly: id, advertiser_id, proposal_id FK `on delete set null`,
report_json, status, created_by, created_at). "No proposal" is always an
explicit choice in the upload-first door. Ends with a confirmed mode and
parsed export(s) in session state — still no deck assembly.

### Phase 3 — Report master deck + walking skeleton — **landed 2026-09-06**

Matt built `REPORT_MASTER_v0_2.pptx` from the slide inventory above
(13.333×7.5in) already carrying every slide's `key:` label and the
`delivery_set: true` marker himself — the "Claude tags every slide" step
below turned out to be nothing to do; a token check against this file's own
inventory (README section below) found zero missing tokens and zero
near-misses. `report_assembly.py` fills all seven slide types (not just
Campaign Recap) end-to-end for no-proposal mode, `report_charts.py`
renders the bar charts and zip map, and the Attribution Reports page gained
a real "Generate the report deck" button + download. Verified against two
real fixtures and rendered/screenshotted per the standing UI-verification
rule. Two real, confirmed layout defects were found by rendering (the
report tables/tiles have no measured-fit pass, and the blank `FLIGHT_LABEL`
tile has no named shapes to delete) and stopgapped rather than truly fixed
— see DECISIONS.md's Attribution Report Builder section for both, plus
the `channel_vcr` parser gap and the mislabeled-fixture finding.

**Deferred out of Phase 3, unlike originally scoped:** `report_deck_versions`
+ the `report_decks` bucket + a `setup_supabase.py` step exist in
`supabase_schema.sql` (Stage 15) but haven't been run against live
Supabase — the Generate button reads the template from the repo root
directly for now. Folding report-deck management into "Update master deck"
as a second picker is still not built. The highlight/breakdown/takeaway
narrative text is deterministic Python (grounded in real computed facts,
never fabricated) rather than Claude-drafted — that's still Phase 4's
`build_attribution_prompt`/`call_claude_attribution`, which can swap in
against the same facts this phase already computes.

### Phase 4 — Full slide set, no-proposal mode

Remaining six slide types against the parsed export(s) alone. New
`build_attribution_prompt`/`call_claude_attribution`/`apply_attribution_
draft` (facts-payload-only contract — Q7 in the kickoff answers), the
no-proposal drafting schema (goals/strategy/optimizations-made/whats-next).
`report_charts.py` — one shared matplotlib module (KPI tile, bar,
horizontal bar, line), rendered at the real EMU size of each named chart
region converted to print DPI, not a screen-preview default. Needs Matt's
Auto-Sales-Analyst repo link first, to decide reuse-vs-reimplement for
`build_group_chart_images()`. Verified against both MW and Cardinal.

### Phase 5 — Proposal link (with-proposal mode)

Confirmed-match path from Phase 2 wired into real content: recap's goals/
audience/geo/budget from `form_json` instead of drafted notes; zip-analysis
gains the targeted-vs-visitor overlay via the two-pseudo-group `render_map`
call; highlights/takeaways get richer since goals are now precise.
`attribution_reports.proposal_id` populated and exercised end-to-end. Needs
Matt's Pansophic/WAEPA/Bozzuto real report+proposal pair to design the
matching UI against.

### Phase 6 — History surfacing, advertiser-as-spine, and the loop back to a follow-up proposal

Three pieces, all keyed off the `advertisers` table built in Phase 2:

- **Both grids filter by advertiser.** Proposal History already does;
  the Attribution Reports list (folded into that same page rather than a
  separate one, per the original plan) gets the identical grouped-by-client
  rendering `render_proposal_history` already uses.
- **Advertiser as the spine — designed toward, not necessarily built here.**
  A single client view showing proposals, reports and case studies
  together becomes possible once all three can join on `advertiser_id`
  (case studies would need the same nullable FK added when this phase
  starts, if they don't already have one). Scope the actual UI once
  Phases 2-5 are live and there's real multi-report-per-client data to
  design against — the table existing is what makes this cheap later, not
  a reason to build the combined view now.
- **Build follow-up proposal from the report — pulled forward from
  "deferred" to planned, per Matt's framework §7 call: "it's the half of
  the loop that makes the proposal link worth building."** A finished
  report gets a button that opens Build a proposal with client,
  originating market, the previous plan, and the report's own what's-next
  content pre-filled into the setup band and the notes — the mirror image
  of Phase 2's "Build report from proposal" button, closing the loop in
  both directions. Also the natural home for the explicitly-deferred
  **month-over-month linking** across several monthly reports for one
  advertiser ("if the user so chooses") — scoped here once real monthly
  reports exist to design against, not before.

### Deferred, named for continuity

- **Optimization engine** (framework/addendum step 3) — blocked in part on
  whether the dashboard can pull a full daily time series (today's "Day of
  Week" tab is only a trailing 7-day window, not a real weekday aggregate).
  Matt is checking.
- **Create case study from report** (addendum §3) — sequenced after Phase 5
  in the addendum itself, so it can use both the report and its proposal.
- **Polk/Auto-Sales-Analyst integration, sales match-back, brand lift,
  Arrivalist** (addendum §4-9) — each gated on a real raw deliverable from
  Matt, per the addendum.

## Open items Matt is chasing

- Auto-Sales Analyst's `build_group_chart_images()` — repo link coming
  before Phase 4.
- Whether the Premion dashboard can pull a full daily time series — not
  blocking before the optimization phase.
- Which of Pansophic / WAEPA / Bozzuto (all built in this app) has a
  matching attribution export, to design Phase 2/5's matching UI against.
