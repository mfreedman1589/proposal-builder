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

### 8. Live Sports — `report:live_sports` — **OPTIONAL, and delivery set: true** — only when the delivery file's own workbook carries a distinct live-sports block

**Not part of the original 7-slide spec — added 2026-09-08** when a real
Prince George's Community College delivery export (`Premion OTT (2).xlsx`)
turned out to carry a second, distinct block: its own KPI/detail/event/
breakdown tabs for a PREM TV live-sports package (RFPID-266713) riding
inside the same workbook as the OTT campaign (RFPID-266710). Detection is
the `IMPRESSION BY EVENT` header (`attribution_import.py`'s own
`_SPORTS_EVENT` constant) — never a sheet name, since Excel BOM-pads the
sports block's own colliding tab names ("KPI DELIVERY" appears on both
blocks; the real, load-bearing difference is that tab reading `VCR %`
where the OTT one reads bare `VCR`).

**Code-side (`report_assembly.py`) is DONE and tested against the real
file** — `_fill_live_sports`, `live_sports_applies`,
`combined_headline_impressions` all exist and are guarded so nothing
breaks against a template that doesn't have this slide yet (`build_report_
deck` only fills/drops it when `"report:live_sports"` is actually present
in `_slide_by_key`). **What's still needed is the template slide itself** —
Matt builds it once these token names are settled:

| Token | Type | Expects |
|---|---|---|
| `{{SPORTS_IMPRESSIONS}}` | string | Sports package's own delivered impressions (e.g. "55,674") |
| `{{SPORTS_VCR}}` | string | |
| `{{SPORTS_PACING}}` | string | "delivered of flight goal" — e.g. "55,674 of 625,000" |
| `{{SPORTS_PACKAGE_TYPE}}` | string | e.g. "PREM TV - LIVE NCAAF REGULAR SEASON" |
| `{{SPORTS_RFPID}}` | string | |
| `{{SPORTS_GEO}}` | string | The package's own Delivered Geo |
| `SportsEventTable` / `{{SPORTS_EVENT_ROWS}}` | table | Top 10 by impressions (Date, Event, Network, Impressions, VCR), PLUS one final rolled-up "All N events" row summing every event so the true total is always visible past the cap — "a network rollup line," read as one summary row, not a second table. **5 columns in this order; the token goes in the Date cell ONLY** (`_fill_named_table`'s own one-token-per-row convention) — Event/Network/Impressions/VCR are plain cells in that same template data row, filled positionally, no token of their own |
| `SportsByLeagueTable` / `{{SPORTS_BY_LEAGUE_ROWS}}` | table | (League, Impressions) — shown ONLY when more than one league ran; deleted (with `SportsByLeagueHeader`) otherwise, same show-what-matters rule the Delivery Breakdown slide's own two conditional tables follow. **2 columns; the token goes in the League cell ONLY**, Impressions is a plain cell filled positionally |
| `{{LIVE_SPORTS_NARRATIVE}}` | string | One to two Claude-drafted sentences, or the computed fallback (leading event/network) |

**The Highlights slide's own `{{HEADLINE_IMPRESSIONS}}` changed too**
(no template change needed there, already live): it's now OTT + live
sports **combined** whenever a sports block is present
(`report_assembly.combined_headline_impressions`) — confirmed against
this real file that the export's own totals never combine the two blocks
(`DETAILS BY FLIGHT`/`KPI DELIVERY` are OTT-only, 132,713; the sports
block is a separately-total 55,674 nowhere added to it), so the app has
to do the combining deliberately. Sports stays visibly broken out on ITS
OWN slide (above) rather than silently folded into one bigger number with
nothing to show for where the sports share went.

Guard: `tests/test_report_assembly.py`'s `check_live_sports` (facts payload
+ headline combination, fully tested today; the slide-fill assertions
themselves are gated on the key's presence and activate automatically the
moment the template lands).

---

## Phase plan

Same discipline as `FLOW_REWORK_PLAN.md`: small, independently green
commits; every phase verified against **both** real datasets (MW wrap,
Cardinal monthly); a full `tests/run_all.py` sweep before any phase is
declared done. New pure modules follow the existing
`avails_pdf_import.py`/`wideorbit.py`/`notes_file_import.py` shape (no
Streamlit, no DB, never raises). Real fixture files are gitignored; tests
report `SKIP` rather than fail when absent, matching `test_wideorbit.py`.

### Phase 1 — Parse both exports, no UI, no deck — **landed** (`40ad74e`)

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

### Phase 2 — Page shell, two entry doors, advertiser matching (no deck output yet) — **landed** (`7e0c8d7`)

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

### Phase 3 — Report master deck + walking skeleton — **landed 2026-09-06** (`45ce4af`)

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

### Phase 4 — the model-facing half — **landed 2026-09-08** (`28e66b5`, docs in `d69b870`; the WAEPA multi-RFPID/conversions/v0_4 follow-up in `9dbc781`)

**Rewritten once already (2026-09-08, after Phase 3 shipped) to correct
three wrong claims a fresh reader would have acted on** (the six other
slide types were already built; matplotlib was rejected in favor of
Pillow-only; the auto-sales-analyst chart link was never actually
blocking). That rewrite is what got built — see it above for the shape
this section confirms landed.

**What shipped, in `app.py`:** `build_attr_draft_prompt` / `call_claude_
attr_draft` / `apply_attr_draft`, the same shape as `build_draft_prompt`/
`call_claude_draft` and running through the very same `_call_claude_json`
machinery (stop_reason read before parsing, largest-balanced-JSON
extraction, one corrective retry, plain-language rep errors via
`last_claude_failure`) — no new network/parsing code, a new prompt and a
new response-to-kwargs mapping only. `report_assembly.build_facts_payload`
is the complete, pure input to the prompt: goals, rep notes, and every
computed aggregate (headline/audience/market/creative/breakdown/`intent_
facts()`'s full-precision classes/top pages/zip rows/delivery) the model
is allowed to cite a number from. `report_assembly.build_report_deck`
gained `narratives`/`breakdown_dimension_override`/`geography_label_
override` keyword arguments threading a Claude draft's content into the
five narrative-sentence tokens, the three `*_HEADLINE_NOTE` tokens, and
`pick_breakdown_dimension`'s audience-vs-creative judgment — all optional
and all defaulting to Phase 3's exact prior deterministic behavior when
omitted, so nothing about Phase 3's own callers/tests changed.

**Facts-only contract, enforced, not just asked for.** `app._attr_payload_
numbers` builds the full set of numbers the payload actually contains —
plain and comma-separated ints, 0/1/2-decimal percentages for rate/share/
vcr-shaped fractions, every list's own length (a model writing "all 10
zip codes" when the table has 10 rows is counting what it was handed, not
inventing a statistic), and every number embedded in an already-formatted
string leaf (`top_zip_rows`/`top_url_rows` hand the model pre-formatted
strings like "2.40%"/"1.47x"/a zip code, not raw floats) — plus a K/M/B-
abbreviation check by scale ("917K" against a raw 917,451 fact) rather
than by string match, since the model can round to any precision.
`app.apply_attr_draft` runs every drafted string through this and returns
non-blocking warnings for anything that doesn't trace back — a rep reviews
either way, same as every other Claude-drafted content in this app.
**Both the list-length and K/M-abbreviation handling, and the K/M one
specifically because "917K impressions" reads far better in a headline
tile than "917,451," were found missing against real MW/Cardinal live
responses, not designed in ahead of time** — this is the accepted shape of
building a facts-only checker: start strict, run it against the real
model, fix the false positives it finds.

**The checker's own limit, stated plainly rather than left implicit: it
verifies every emitted number traces to the payload; it cannot verify that
a traced number wasn't REACHED by arithmetic performed on other traced
numbers.** Found live: a Cardinal run once wrote "roughly 1 in 15
attributed visitors" by inverting a summed percentage (4.7% lead-intent +
1.9% purchase-intent ≈ 6.6% ≈ 1-in-15) — genuine new arithmetic the prompt
explicitly forbids, that the checker missed only because "15" happened to
also be a real, unrelated fact elsewhere in the payload (the lead class's
own raw visit count). A syntactic "does this number appear anywhere" check
cannot verify that a *derivation* connecting two real numbers is itself
sound — that would need actually re-deriving every ratio the model states,
which is a different, much larger checker than this one, not a bug in this
one. The fix that IS in scope: the prompt's facts-only paragraph now names
this exact pattern ("combining two SEPARATE facts into a derived one...
adding two classes' shares together, or inverting a percentage into
'roughly 1 in N'") as its own worked counter-example, not just the general
rule. Re-ran clean afterward, twice. **This lowers the odds of the whole
class recurring; it does not close it** — a differently-shaped derivation
the counter-example doesn't resemble could still slip through unflagged,
by design, not by oversight. The checker is a first-pass filter for
egregious fabrication, not a proof of narrative correctness, and was never
going to be either.

**The real backstop is a rep reading the narrative before it ships — which
is exactly why the `goal_alignment_notes` surfacing bug (below) mattered
more than it looked.** `apply_attr_draft` originally computed its kwargs
and simply discarded `draft["goal_alignment_notes"]` — the model's own
"this note disagrees with a goal" / "no data exists to quantify this"
flag, working exactly as designed, going nowhere a rep could see it. Fixed
by folding it into `apply_attr_draft`'s returned `warnings` (labeled
distinctly from a number-fabrication warning, since one is "review this
number" and the other is "review this specific stated disagreement") and
by rendering it in the UI right after drafting — persisted into
`st.session_state["attr_draft_goal_notes"]` rather than shown inline in
the same script run, since anything rendered in the same run as the
`st.rerun()` that follows a successful draft is wiped out before a rep
ever sees it (a real mistake caught before it shipped, not a hypothetical
one). Shown again alongside the facts-only warnings at Generate time, so
it's visible whether or not the rep looked right after drafting.

**The URL slide is the point of this phase, and it works.** Verified live
against both real datasets with their real goals: MW ("drive in-store
visits") names the store-visit intent class with its real 9%/667-visit
figure; Cardinal ("generate service calls") names the lead-intent class
with its real 4.7%/15-visit figure. The no-goals edge case (verified live
against MW) describes the intent mix with real numbers and does not claim
alignment to anything. A rep note describing a mid-flight optimization
(verified via the offline Tier-1 fixture, a synthetic payload — see below)
reaches the drafted content, and the model correctly declined to fabricate
a day-of-week number the facts don't contain rather than inventing one to
satisfy the note.

**`whats_next_bullets` stays rep-owned, deliberately narrower than this
section's own earlier draft implied.** `apply_attr_draft`'s returned
kwargs never include it — Claude's `whats_next_bullets` output only
*pre-fills* the rep's own text box (and only when the rep left it blank),
the same "structural replacement bumps a generation, doesn't force-
overwrite a hand-typed value" discipline this app uses everywhere a
drafted value meets an editable widget. Generate always reads the box
itself, never the draft dict directly.

**Guards:** `tests/test_attribution_draft.py` (Tier 1 — free, offline,
`tests/fixtures/attr_synthetic.draft.json` frozen against a synthetic,
non-client facts payload rather than a real MW/Cardinal one, for the same
reason MW/Cardinal's own xlsx files stay gitignored: freezing a fixture
from them would put real campaign numbers permanently into git history).
`tests/test_attribution_draft_live.py` (Tier 2 — live, a few cents, MW/
Cardinal/no-goals). `tests/test_report_assembly.py`'s `check_phase4_
override_plumbing` (offline given the real template+Cardinal export —
proves the override kwargs actually reach the rendered slide text and
REPLACE the deterministic default, and that omitting them reproduces
Phase 3's byte-for-byte prior behavior).

### Phase 5 — Proposal link (with-proposal mode) — **landed 2026-09-09** (`963d79b`, docs in `bdc02ea`/`97fc877`; UI-bug fixes and the Generate-flow redesign in `6c4a87e`)

Confirmed-match path from Phase 2 wired into real content: recap's goals/
audience/geo from `form_json` instead of drafted notes; zip-analysis gains
the targeted-vs-visitor overlay; highlights/takeaways get richer since
goals are now precise.

**Shipped shape (the field map below is what actually got built, not a
plan):** `app.linked_proposal_report_fields(form_json, target_dmas,
profiles)` is the whole field map in one pure function -- `audience_bullets`
(from `campaign_specs.audience`), `flight_label` (`flight.label`/
`.shorthand`), `geography_names` (`target_market_labels(target_dmas,
profiles)` -- never `campaign_specs.geography`, narrative prose), `budget`
(`deck_payload.media_plan_options[0].full_flight_total.cost`, parsed by
`app._parse_money_string`), and `targeted_zips` (the union of every
`targeting_groups[].resolved_zips`). `render_attribution_reports_page`
computes it once, right after `prelinked_row` resolves, and threads it
through both `build_facts_payload` calls (drafting) and the `build_report_
deck` call (Generate) -- never re-derived per call site. The linked
proposal's own `client_name` wins from the advertiser-confirm step onward
(a real gap: the export's own name was used everywhere below it before).

**Geography** is threaded as a NAMES LIST (`report_assembly.
geography_names_override`), not a pre-joined string -- `geography_label_
from_names`/`geography_overflow_bullet_from_names` were factored out of
`geography_label`/`geography_overflow_bullet` (the export-derived path) so
both sources share one collapse rule (join up to 2 names; 3+ collapses to
"N markets" with the full list as an overflow Audience-bullets line). The
existing Phase 4 `geography_label_override` (a bare string) stays exactly
as it was, for full backward compatibility with `check_phase4_override_
plumbing`.

**Budget** is a `build_facts_payload` INPUT, never a recap slot -- `facts
["budget"]` (`total`/`cost_per_attributed_visit`/`cost_per_conversion`,
None when no proposal is linked) lets the model cite cost-per-X facts,
exactly what WAEPA's own notes asked to see. **Found and fixed alongside
it:** the facts-only number checker (`app._attr_payload_numbers`) only
ever recognized whole-dollar and 0-1 rate/percentage forms -- a real
dollar-and-cents fact (a $121.11 cost-per-visit) is a genuinely new shape
this payload never carried before Phase 5, and a live WAEPA draft
correctly cited it and still failed the check. `add_int` now also
registers a non-integral value's own 2-decimal string form.

**The "band wins and flags" contradiction rule** is structural for flight/
geography, not model-mediated: FLIGHT_LABEL/GEOGRAPHY_LABEL are always
Python-filled from the linked proposal (never from `apply_attr_draft`'s
kwargs, which can't carry either key), so a note claiming a different
flight can't reach the deck regardless of what the model does with it.
`facts["proposal"]` (`flight_label`/`geography_label`, None when nothing's
linked) exists purely so the model can still NOTICE the contradiction and
name it in `goal_alignment_notes`, verified live against a real WAEPA
draft that correctly flagged a deliberately wrong flight note. Guard:
`tests/test_attribution_draft_live.py`'s `waepa_proposal` scenario.

**The zip overlay composes on top of the existing choropleth, per the
design note already in `targeting_map.py` above `render_choropleth`**:
`render_choropleth(..., targeted_zips=...)` resolves the targeted zips'
own polygons (widening the map's frame to include them even when they
have zero attributed visits), strokes them as a fixed-color outline layer
over the finished choropleth, and adds a "Targeted zip codes" legend
swatch ONLY when at least one outline actually drew -- never merely
because the parameter was passed. `None`/empty `targeted_zips` reproduces
the prior PNG byte-for-byte (verified `_parts_equal`, since two separate
`prs.save()` calls differ only in the zip container's own save-time
stamp, never in content) -- the real WAEPA fixture (`09e61e0b`,
`targeting_groups: []`, avails_mode was off) exercises exactly this
degrade path. **No real avails-linked WAEPA pair existed at build time**
(checked; none found), so the two-series draw path is verified against a
SYNTHETIC target-zip list pulled from the real crosswalk (`tests/
test_report_assembly.py`'s `check_phase5_proposal_link`, explicitly marked
synthetic in its own docstring) -- replace it the day a real pair exists.

Guards: `tests/test_report_assembly.py`'s `check_phase5_proposal_link`
(offline, the deterministic half -- geography collapse, budget facts,
`build_report_deck` threading, zip-overlay degradation and drawing),
`tests/test_attribution_reports_page.py`'s `check_phase5_proposal_link`
(through the real page via AppTest -- the goals-prefill fix, client-name
override, and the generated deck's own FLIGHT_LABEL/GEOGRAPHY_LABEL/
AUDIENCE_BULLETS tokens), `tests/test_attribution_draft_live.py`'s
`waepa_proposal` scenario (Tier 2 -- the contradiction-flagging half).
Full chunked sweep clean (81/82 files; the one failure,
`test_draft_regression.py`'s HVAC scenario, is pre-existing calendar
drift on a frozen fixture's flight date, unrelated to this phase).
`attribution_reports.proposal_id` populated and exercised end-to-end.

**Acceptance walkthrough follow-up (2026-09-10, Matt) -- two real bugs
found by clicking through the actual page, both fixed same-day.** The
field map was wired off the pre-linked door only, so linking a proposal
via this page's OWN search (section 3) prefilled nothing even though
"Proposal linked." showed -- `linked_row` now unifies both doors. The
proposal picker showed only a bare client name and match score -- each
candidate now carries title, flight dates, logged-date and logged-by.
Alongside these: "Log this report" folded into Generate (no separate
step); "Draft narrative" renamed "Preview narrative" and made optional;
Generate is now self-sufficient (drafts automatically when nothing's been
previewed, never re-drafts over an already-previewed one). Full narrative,
including a real AppTest-stubbing gotcha found writing the tests for
this, in `DECISIONS.md`'s Attribution Report Builder section. This is
also what produced CLAUDE.md's new standing rule -- walk the page as a
rep before declaring a flow change done.

**Fixture: proposal `09e61e0b-a945-4022-851d-d3c31b2acbd0` (WAEPA) + the
WAEPA attribution export.** Real `form_json` pulled and compared against
every recap-relevant token before writing the field map (2026-09-09):

| Token | Source in `form_json` | Note |
|---|---|---|
| `CLIENT_NAME` | **top-level row column** `client_name`, NOT inside `form_json` | Already the right place to read it -- `attribution_dict`'s own `client_name` already competes with this today; the linked proposal's should win when linked |
| `GOALS_BULLETS` | `form_json["campaign_specs"]["goals"]`, newline-joined string -- split with the existing `lines_to_bullets()` helper (already used for this exact transform at app.py:3829, campaign-specs-copy path) | **Real bug found along the way**, not yet fixed: the existing partial prefill at app.py's `attr_goals_input` reads `form_json.get("goals", "")` -- a bare top-level key that **does not exist** on any real proposal (confirmed against the full key list below). It has silently prefilled empty string for every proposal ever linked. Phase 5 must read `campaign_specs.goals` instead, and this latent bug should be fixed in the same change |
| `AUDIENCE_BULLETS` | `form_json["campaign_specs"]["audience"]`, same newline-join, same `lines_to_bullets()` | Straightforward |
| `GEOGRAPHY_LABEL` | **NOT** `campaign_specs.geography` (narrative prose -- WAEPA's reads "Washington DC (primary): 535,000 impressions/month\nBaltimore: 200,000...\nConcentrated two-market approach..." -- a paragraph, not a label). Resolve the **row-level** `target_dmas` (market KEYS, e.g. `['washington_hagerstown', 'baltimore']`) through the existing `target_market_labels(keys, profiles)` helper instead -- confirmed live: resolves to `['Washington, DC', 'Baltimore']`, then through `report_assembly.geography_label`'s own list-based formatting (2-or-fewer join, "N markets" collapse) via `geography_label_override`, which `build_report_deck` already accepts as a parameter | **The exact same class of bug `geo_column_default`'s 2026-09-08 fix corrected for the plan table's Geo cell** -- a report tile needs a short label, `campaign_specs.geography` is narrative, and the structured, correct source is the target-market KEYS, not the prose |
| `FLIGHT_LABEL` | `form_json["flight"]["label"]` ("Oct 2026 - Dec 2026") or `["shorthand"]` ("Oct–Nov, Dec 1–29") | Clean, already exactly shaped for this token -- `build_report_deck`'s own docstring already expected this source |
| `REPORT_PERIOD_LABEL` | **Unchanged from Phase 3/4** -- stays derived from the ATTRIBUTION EXPORT's own weekly/monthly trend span, not from `form_json` | Deliberately different from `FLIGHT_LABEL`: a report can cover one month of a longer campaign flight: the two tokens describe different things and must not collapse into one source |
| Budget | **No token exists for it anywhere in the report deck.** `REPORT_MASTER_README.md` says so explicitly: "No budget element on the recap." `campaign_specs.budget` is a real, available narrative string ("$24,990/month — $74,970 total for 90-day test\n$34 CPM; 735,000 impressions/month...") if a use is found for it, but nothing currently reads it | **Open question, not resolved here** -- does "budget" in the original Phase 5 note mean a literal tile (needs a template change), a fact fed to the drafting prompt for a cost-efficiency highlight (no template change, just a new `build_facts_payload` field), or was it aspirational/imprecise. Confirm before the field map assumes either |
| Targeted-zip overlay | `form_json["targeting_groups"]` | **Empty list `[]` on this specific fixture.** WAEPA was built with "Working from an avails document" OFF (avails_mode=False, per the very toggle-authoritative fix landed this same week) -- no avails import ever ran, so no group ever resolved `resolved_zips`. This will be the common case for any hand-typed, non-avails proposal, not a WAEPA-specific gap. **The overlay mechanism needs graceful degradation built in from the start** (matching this app's `ctv_share`-style "None, never a fabricated 0%" discipline): a linked proposal with no resolved zips falls back to today's visitor-only choropleth, never an empty overlay or a crash. **Testing the actual two-series overlay needs a second fixture** — a proposal built WITH avails_mode on and real `resolved_zips`, paired with a matching attribution export. Not identified yet; ask Matt for one, or use `group_scenario_fixtures.py`'s existing real avails documents (Hershey, Annapolis, Wilmington) if any has a matching attribution export, or accept that the overlay ships gracefully-degraded until one exists |
| `attribution_reports.proposal_id` | Already threaded end-to-end at the LOGGING step (`render_attribution_reports_page`'s "4. Confirm" → "Log this report" → `db.log_attribution_report(advertiser_id, st.session_state.get("attr_proposal_id"), ...)`) | This part of Phase 5 is effectively **already done** from Phase 2 -- what's missing is Generate (step 5) itself never reading the linked proposal's `form_json` to populate goals/audience/geo, which is the actual remaining work above |

`form_json`'s real top-level keys (this fixture): `agency_gross_up`,
`avails_basis`, `avails_label`, `avails_rows`, `campaign_specs`,
`case_studies`, `deck_payload`, `draft`, `flight`, `included_list`,
`logo_used`, `markup`, `plan_options`, `proposal_title`, `selections`,
`setup`, `targeting_groups`, `vault_slides`. `campaign_specs` itself:
`goals`, `budget`, `timing`, `audience`, `geography`, `placements` — each
a single newline-joined string, never a list.

**Decisions settled 2026-09-09, ready to implement — none of this built
yet, this section IS the field map:**

- **Budget: no recap slot, don't add one — a client already knows what
  they spent.** Instead it's a `build_facts_payload` INPUT: with a real
  number in hand, Python computes cost per attributed visit and cost per
  conversion — literally what WAEPA's own notes asked the reporting to
  show. `build_facts_payload` gains an optional `budget` kwarg; when given,
  adds a `facts["budget"]` section (`total`, `cost_per_attributed_visit`,
  `cost_per_conversion` — the latter only when conversions are in play).
  No template change; the model selects these for a highlight/narrative
  like any other fact. With no proposal linked, `budget` is absent and the
  cost-per facts simply don't exist — same "no half-states" discipline the
  conversions layer already follows.
  **Source is deliberately NOT `campaign_specs.budget`** (narrative prose —
  "$24,990/month — $74,970 total for 90-day test\n$34 CPM..." — exactly
  the kind of string this app's own rule says never to parse for a
  structured number). The clean, already-computed, deck-accurate figure is
  `form_json["deck_payload"]["media_plan_options"][i]["full_flight_total"]
  ["cost"]` (a predictable `"$74,970"`-shaped string, safe to strip `$`/`,`
  from and parse) — the SAME verbatim, never-recomputed figure the actual
  proposal deck shipped with, matching this app's "Rebuild as presented"
  discipline elsewhere. **Which option, if more than one:** a proposal
  carries 1-3 media plan OPTIONS (alternatives a client chose between, not
  sequential spend) — summing across them would be wrong. Take option 0
  (typical case: exactly one option on a real running campaign) rather
  than guessing which one the client picked; note this as a known
  simplification if a multi-option proposal is ever linked.
- **Zip overlay: build the real two-series overlay, with graceful
  degradation from the start.** No `targeting_groups`, or none with
  `resolved_zips`, on the linked proposal → visitor-only (today's
  behavior), never a crash and never an empty second series — same
  "`None`, never a fabricated 0%" discipline `ctv_share`/`geography_label`
  already follow. **Real two-series test fixture:** if a real WAEPA avails
  PDF exists from when the campaign was pulled, Matt regenerates the WAEPA
  proposal with `avails_mode` ON against it and THAT becomes the fixture
  (09e61e0b's own `targeting_groups` is `[]`, avails_mode was off — it
  cannot serve this half of Phase 5 as logged). **A quick check of Downloads
  and the project root (2026-09-09) found no file matching this project's
  `Premion Media Plan_RFPID-...` avails-PDF naming convention for WAEPA** —
  Matt is checking further. **If none turns up**, fall back to a synthetic
  `targeting_groups` entry built from a real zip list pulled from the
  crosswalk (`geo_resolver`/`market_lookup` — real zips, not invented ones,
  same standard every other geo fixture in this app already holds itself
  to) for the Washington DC / Baltimore markets, explicitly marked
  synthetic in the test's own docstring and comments, replaced the day a
  real avails-backed pair exists — never left silently passing as if it
  were real.
- **Goals prefill: fix in the same pass.** The existing `attr_goals_input`
  prefill (`(prelinked_row.get("form_json") or {}).get("goals", "")`) reads
  a top-level `form_json["goals"]` key that **does not exist on any real
  proposal** — confirmed against 09e61e0b's own full key list above. It
  has silently prefilled an empty string for every proposal ever linked,
  since this code shipped. Fix: read `form_json["campaign_specs"]["goals"]`
  and split with the existing `lines_to_bullets()` helper (app.py:7700,
  already used for this exact transform at the campaign-specs-copy call
  site, app.py:3829). **Add a test that a real linked proposal's goals
  actually reach the text box** — the fact that the current bug shipped
  and stayed shipped means the existing test coverage for this prefill
  either doesn't exist or asserts against something that agrees with
  itself; don't repeat that mistake for the fix.
- **Geo: `target_market_labels(target_dmas, profiles)` into
  `geography_label_override`.** Confirmed live: resolves 09e61e0b's real
  `target_dmas` (`['washington_hagerstown', 'baltimore']`, a top-level ROW
  column, not inside `form_json`) to `['Washington, DC', 'Baltimore']`.
  `report_assembly.geography_label`'s own "2-or-fewer join, else 'N
  markets'" formatting rule needs factoring out into a reusable pure
  function (e.g. `geography_label_from_names(names)`) that both
  `geography_label(attribution)` (derives names from the export) and this
  Phase 5 path (derives names from `target_market_labels`) call — one
  formatting rule, two name sources, not two copies of the collapse logic.
  The matching overflow case (`geography_overflow_bullet`, 3+ markets) needs
  the same treatment for consistency, even though it won't fire on the
  2-market WAEPA fixture.
- **Build target: Phase 5 against `09e61e0b-a945-4022-851d-d3c31b2acbd0`
  and the WAEPA attribution export** — for goals/audience/geo/budget. The
  zip-overlay half needs its own fixture per the decision above before it
  can be verified against something real.

### Phase 6 — History surfacing, advertiser-as-spine, and the loop back to a follow-up proposal — **landed 2026-09-15** (`8ee3165`)

All three pieces below shipped in one commit, keyed off the `advertisers`
table built in Phase 2. This supersedes the original "designed toward, not
necessarily built here" scoping note for the advertiser-as-spine piece —
Matt's own call, once Phases 2-5 were live, was to build the combined view
now rather than defer it further.

- **Both grids filter by advertiser, landed.** Proposal History already
  did; the Attribution Reports page split into New report / Report
  history tabs, the latter grouped by client the same way
  `render_proposal_history` already groups proposals.
- **Advertiser as the spine, landed.** A new Clients page joins a client's
  proposals, reports and case studies on `advertiser_id` in one view, with
  a vertical field and a trend table. `case_studies.advertiser_id` and
  `case_studies.source_report_id` (Stage 17 DDL, live) are the FKs that
  make a case study join into this view too, and mark it as generated
  from a report. `backfill_advertisers.py` linked 19 of 21 existing
  proposals (2 placeholder "Client" rows correctly left unlinked). A
  Merge/rename clients admin page handles duplicates — WAEPA TEST merged
  into WAEPA was the first real case, and (per the housekeeping above) the
  only WAEPA TEST report/case-study pair left over from testing has since
  been deleted.
- **Build follow-up proposal from the report, landed.** `app.
  build_follow_up_proposal` (called from a "Build follow-up proposal"
  button on each report-history row) opens Build a proposal with client,
  originating market, the previous plan, and the report's own what's-next
  content pre-filled into the setup band and the notes — the mirror image
  of Phase 2's "Build report from proposal" button, closing the loop in
  both directions, per Matt's framework §7 call ("it's the half of the
  loop that makes the proposal link worth building").
  **Month-over-month linking** across several monthly reports for one
  advertiser is still deferred — it wants real monthly-report chains to
  design against, which the account-trend work below (`prior_periods`)
  is starting to produce but hasn't yet at volume.

Three follow-on rounds landed on top of the spine the same day and the
next, all client-vertical/optimization work rather than new phases of
their own:

### Optimization engine (framework §4-6) — landed 2026-09-15 (`cb6276c`)

`report_assembly.optimization_candidates()`: structurally subtraction-only
(no over-performer candidate type exists at all), the material-swing floor
reused from the Highlights/Takeaways thread work (20% relative / 1.5x
multiple), an already-suppressed value (impressions far below peer
average) reported as already-limited rather than rediscovered as a fresh
cut, first report for an account gets zero candidates (Python-enforced,
never model-discretionary), ZIP gets the framework's own %-of-total cap,
other dimensions get a flat count cap. Publisher candidates default OFF
(Matt's ruling — seasonality). Verified live against real MW data twice:
correct subtraction language, correct already-limited handling, zero
facts-only violations both runs. **Still open, unchanged from the original
deferred note:** whether the Premion dashboard can pull a full daily time
series for the engine's remaining dimensions beyond weekday — Matt is
checking that separately; the engine ships and runs today without it.

### Optimization sequence: accept/edit/decline, cross-month memory, client-level setting — landed 2026-09-15 (`cad0900`)

The cross-month memory a monthly cadence needs: accepted (or edited)
candidates are stored on the report (`report_json["optimizations"]`,
including declined ones — "the declines are the interesting half for
learning later"). A later report reads the prior report's accepted list,
excludes those values from fresh candidacy, and measures their effect
against current data (share of delivery and campaign rate, before vs.
after); a wrap reads the whole chain as `facts["optimization_history"]`.
Accept/edit/decline is a real checklist on the New report tab before
Generate ever runs (default Accept) — a decline is structurally absent
from what reaches the model, not merely told not to use it. Optimization
level moved from a per-report toggle to a client-level setting
(`advertisers.optimization_level`, Stage 18 DDL, live) with a new "None"
tier (reporting only, the default for a client nobody's set a preference
for) — the engine still runs at None, since the in-effect measurement is
unconditional. The Clients page also gained vertical self-healing:
vertical now derives from a client's own linked proposals when they agree
(one-shot backfill: 9 of 14 previously-unassigned advertisers derived;
Mattress Warehouse and 4 others have no linked proposal to derive from,
still manual), left unassigned with the disagreement named when they
don't. **Confirmed live**, not just from the commit message: WAEPA's
`advertisers.vertical` is `"banking"` in production today, which
`REVERSE_VERTICALS` (`app.py`) resolves to the Clients page's displayed
"Banking & Finance" — the exact label the derivation was meant to produce.

### One-slide summary + case study from report — landed 2026-09-16 (`e8e652f`)

Closes the two oldest items on the deferred list below. `report:summary`
and `report:case_study` are two new standalone slides on
`REPORT_MASTER_v0_9.pptx` (live and active in Supabase,
`report_deck_versions` id 7) — a marker-scan drops them from every normal
multi-slide report build, and `report_assembly.build_single_slide`
isolates either on demand. Both fill from a report's own already-stored
`threads`/facts (`rehydrate_attribution`/`rehydrate_delivery` reconstruct
a logged report back into the real dataclass shape) so either can be
built later with no second drafting call. Case-study creation was
reworked mid-build, per live feedback, into "the vault is the default
action, not a destination": tags auto-suggest via the existing Claude
tagging call, "Save to case study vault" is the primary button with a
plain download as the secondary action, and both the vault browser and
the Clients page mark a case study as generated from a report, linking
back to it. Six real bugs were found this round (three by the live
browser walkthrough, not the test suite) and fixed with regression tests —
full list in the commit message (`git log -1 e8e652f`) and cross-session
memory; not yet copied into `DECISIONS.md`. Full `tests/run_all.py` gate
green (89 files, 0 failures) after.

### Phase 7 — Polk automotive match-back — **PROPOSED 2026-09-16, CLOSED 2026-09-19**

Addendum §5's build-order item 6 ("Polk automotive + Auto-Sales Analyst —
the biggest vertical, and partly built already"). The fixture is in hand
and re-verified directly against the file this session (`Polk
Dashboard.xlsx`, project root, gitignored): 19 tabs, matching the
inventory below exactly. Headline figures confirmed by reading the file,
not recalled from the addendum: 44,756 matched households, 5 Target
Dealer Sales, 0.01117% buy rate, 591,686 matched impressions, 90.49% match
rate, 1.5x campaign lift.

**The absorb/call/align fork (addendum §4) doesn't reopen here — it's
already answered.** That fork was about the Auto-Sales Analyst (a live
scanning tool), not about Polk (a flat monthly deliverable, parsed the
same way the attribution/delivery exports already are — no live system to
integrate with). The Auto-Sales Analyst deck append (`report_assembly.
append_slide_deck`, landed 2026-09-08) already settled it as **align**:
the two decks stay separate artifacts sharing one output, glued together
by `extra_deck_path`'s wholesale append. Polk slots into that same
"align" posture by construction — it becomes its own slide(s) in the
report deck, appended-deck-adjacent but built from parsed data like every
other slide, never absorbed into or calling out to the Analyst tool.

**Scope for this phase:**

- **`polk_import.py`** (new, pure — no Streamlit, no DB, mirrors
  `attribution_import.py`'s shape): `parse_polk_export(path) ->
  PolkExport`, covering the headline six tabs plus the cuts: sales by
  days-elapsed/gender/age/income, make-model-per-dealer with average MSRP,
  audience/creative/publisher share of matched impressions AND share of
  target sales (two-tab pairs — a share-of-impressions view and a
  share-of-sales view, kept distinct, never conflated into one number),
  Target Dealer(s) with market rank vs. campaign rank, All Dealers with
  campaign share. Row counts vary widely by real client (this fixture: 2
  target dealers, 2 creatives, 202 publisher rows, 132 all-dealers rows) —
  every table in the design below caps to a top-N plus "+N more," the same
  discipline `TopUrlTable`/`DeliveryByCreativeTable` already use, never a
  raw dump sized to whatever one client's file happens to contain.
- **One new optional slide, `report:automotive_registrations`** (name
  open to Matt's own convention) — delivery-set-style presence gating: no
  Polk file uploaded, the slide drops entirely, same marker-scan
  `build_report_deck` already uses for `report:live_sports` and the
  delivery-set. Proposed content, tiles-plus-one-table per this deck's
  own established density: four headline tiles (Matched Households,
  Target Dealer Sales, Buy Rate, Campaign Lift), a capped Target Dealers
  table (market rank vs. campaign rank, top 5 by campaign rank — this
  fixture only has 2, so no fallback truncation logic to design around
  yet, but a client with more needs the cap), and a short "where the
  matched impressions went" callout naming the top audience/creative/
  publisher share rather than reproducing all three full tables — the
  same "digest, not dump" call already made for `DeliveryByCreativeTable`
  (top 3 creatives) and the zip table (top 5, `ZIP_MIN_SHARE`).
- **Match-rate projection toggle** — build now, exactly as specced
  2026-09-08 above (a pure `project_for_match_rate(value, match_rate)`
  helper, default OFF, every projected figure labeled "projected" in the
  tile/table/narrative alike, never a bare number that looks like a
  fact). This is the toggle's first real consumer; sales match-back reuses
  it unchanged whenever that deliverable arrives.
- **Facts + threads, not the optimization engine.** Polk is an outcome
  measure (did the campaign drive registrations/sales), not a delivery
  dimension to cut — `optimization_candidates()` gets no new Polk-aware
  logic. `build_facts_payload` gets a new optional `polk` input; the model
  may build a new thread type (a goal thread if a stated goal mentions
  sales/registrations/conversions to a dealer, a signal thread otherwise)
  citing the match-rate-respecting language above.
- **A new upload slot** on the Attribution Reports page ("Polk automotive
  match-back (optional)"), same five-step idiom as the delivery/OTT-
  retargeting uploaders — offered unconditionally rather than gated on a
  vertical selection, since a rep uploading it is itself the signal.
- **Human-only handoff, named explicitly:** Matt builds
  `report:automotive_registrations` in PowerPoint once Claude specs the
  exact shape names and `{{TOKEN}}` placeholders (mirroring how
  `report:live_sports` and every other slide in this deck were handed
  off) — this is not a task for Claude to build directly; the template
  file is a human deliverable.
- **Guards:** `tests/test_polk_import.py` (pure parser against the real
  fixture, `SKIP` without it, same convention as every other real-file
  test in this suite), plus the usual `test_report_assembly.py`/
  `test_attribution_draft_live.py` extensions once facts/threads land.
- **Deck order, confirmed 2026-09-16:** `report:automotive_registrations`
  and the appended Auto-Sales Analyst slides are the two halves of one
  automotive argument — registrations anywhere (Polk) vs. this dealer's
  own inventory moving (the Analyst) — so the Polk slide belongs
  immediately BEFORE the appended Analyst slides, not merely somewhere
  before the deck's end. `append_slide_deck`'s call still has to run last
  (after every token-fill call, so an appended slide can never shift a
  `_slide_by_key` lookup `build_report_deck` still needs to make) — that
  constraint is about the APPEND mechanism, not slide position, so it
  doesn't conflict: `report:automotive_registrations` gets placed via the
  normal template-order slide-key scan like every other slide, and as
  long as it's the last real (non-appended) slide in the template — or
  Matt places it there when he builds it — the appended deck lands
  immediately after it by construction, with nothing else able to land
  between them.

**Open question for Matt before building:** does a Polk file arrive
per-dealer-group (one file per campaign, this fixture's shape) or could
one file ever need to be split across multiple target dealers/proposals?
The fixture on hand has exactly one target-market context, so the parser
above assumes "one Polk file = one report," matching how attribution/
delivery exports already work — flag if that assumption is wrong before
the parser is built around it.

**Status, 2026-09-18: everything except the template is built and tested.**
`polk_import.py` (new, pure) parses all 19 tabs, verified against the real
`Polk Dashboard.xlsx` (44,756 matched households, 5 target dealer sales,
0.01117% buy rate, 591,686 matched impressions, 90.49% match rate, 1.5x
lift — every figure re-confirmed directly from the file, not recalled).
`has_target_dealer_sales` is the "present-and->0" detection this phase
asked for, the same rule `attribution_import.py`'s `has_conversions`
already uses — the six headline tabs are always present in a Polk export
(a fixed dashboard shape), so only the *value* (a real paid campaign can
genuinely have zero matched sales so far) tells a real signal from a null
one. `report_assembly.project_for_match_rate` is the shared match-rate
projection helper, built to serve sales match-back too whenever that
deliverable arrives, exactly as specced above. `build_facts_payload` gained
`polk`/`polk_projected`; `facts["polk"]` carries the raw matched figures
always, and `projected_matched_households`/`projected_target_dealer_sales`
only when the toggle is on — the prompt's own new "Polk automotive
match-back" paragraph tells the model to name the match rate in the same
sentence as any raw matched figure it cites, verified live (2026-09-18,
cross-paired against MW's real attribution file, since no attribution
export exists for the same automotive client Polk's fixture describes): a
live run's own thread wrote "5 new vehicle sales tied to campaign-exposed
households... against a 90.49% household match rate" unprompted beyond the
rule itself. **A live run also surfaced a real, PRE-EXISTING model-behavior
gap, unrelated to this phase's own code**: the model twice wrote "100%" (a
number not present in the facts) while describing that intent-class
reach percentages don't sum to 100% — confirmed not a Polk regression by
re-running the completely unrelated `mw_optimization` scenario in the same
session, which failed on the identical "100%" pattern. Left as a known,
tracked gap in the general drafting prompt (out of Phase 7's scope, which
is Polk-specific) rather than patched here.

The upload slot ("Polk automotive match-back (optional)"), the "Project for
match rate" toggle (default off, appears only once a Polk file is
uploaded), and the deck-fill function (`report_assembly.
_fill_automotive_registrations`, `automotive_registrations_applies`) are
all built and wired into both Preview and Generate — independent of the
delivery set, the same way `report:ott_retargeting` is, per this phase's
own scope.

**Status, 2026-09-19: closed. The human-only handoff is done** — Matt built
`report:automotive_registrations` against `REPORT_MASTER_README.md`'s own
"Handoff: new slide for Phase 7" spec (`build_v0_12.py`, cloned from v0_11),
placed as slide 11: immediately after `report:takeaways`, before the
standalone `report:summary`/`report:case_study` slides — both are always
stripped before the ordinary multi-slide build regardless of position, so
this is the true last real slide before any Auto-Sales Analyst append, per
the deck-order ruling above. Verified shape-by-shape against the spec (all
four tiles, `PolkMatchRateNote`, `PolkTargetDealersHeader`+Table, no chart/
map region) before upload. **One real bug found by rendering, not by
reading the fill code**: `PolkMatchRateNote` was cloned from
`SportsRfpidCaption` and inherited its light `C9CBE8` color — unreadable
against this slide's white body, confirmed by a COM render before the fix
existed. Fixed in `build_v0_12.py` to match `PlanVsActualNote`'s dark
`1A1A2E` (the exact precedent the spec itself pointed at) and rebuilt
before upload; a second render confirmed the fix. The "(projected)" suffix
(18pt `PolkHouseholdsTileValue`/`PolkSalesTileValue`, Matt's own sizing
call) renders on one line against the real fixture — `"49,460 (projected)"`,
one off this doc's and the README's own illustrative "49,461" (a docs
rounding slip in the worked example, not a code issue). Uploaded and
activated as `report_deck_versions` id 10.

Guards: `tests/test_polk_import.py` (the pure parser, real fixture +
synthetic edge cases), `tests/test_report_assembly.py`'s
`check_polk_automotive_registrations` (now built against the real
`REPORT_MASTER_v0_12.pptx` instead of `rep.pending()` — exact shape-text
assertions, both toggle states), `tests/test_attribution_reports_page.py`'s
`check_polk_upload_and_projection_toggle` (extended to open the .pptx a
real Generate click just built through the real app against the live
Supabase-active template and confirm the slide is actually filled, not
just exception-free), `tests/test_attribution_draft_live.py`'s `mw_polk`
scenario.

### Phase 8 -- Polk v2 and the cost layer -- PROPOSED 2026-09-21, code built and walked as a rep same day

Matt's own interpretation rules for real Polk data, plus ROI and cost per
visit. Proven against the same `Polk Dashboard.xlsx` fixture Phase 7 used;
`Polk_Dashboard (1).pptx`/`(2).pptx` (the dashboard EXPORTS that prove the
months-stacking rule with real 414.95%/474.42% match rates and 400%/200%/
600% audience-cut shares) were named as being in the project folder but
were **not actually present** when this phase was built -- only `Polk
Dashboard.xlsx` was found. Not blocking (Matt's own division-by-months rule
is fully specified and needed no file to implement), but the multi-month
figures are unverified against a real file; add them to the project root
and re-run `tests/test_report_assembly.py`'s `check_polk_phase8` once
available.

**1. Months-covered normalization -- built.** `polk_import.normalize_
for_months(export, months)` divides match_rate, campaign_lift, and every
audience/creative/publisher cut's share (both the share-of-impressions and
share-of-sales views) by `months`; counts (matched households, target
dealer sales, matched impressions, every sales-cut count, dealer new_sales/
campaign_share, make/model rows) are untouched, since they're already
totals. Raises `PolkParseError` for `months <= 0`. Warns (in the
normalized export's own `warnings` list) when a share or the match rate is
STILL over 100% after dividing -- the sanity check Matt's own spec asked
for. **`buy_rate` is deliberately left undivided** -- Matt's own
enumeration names match rate/campaign lift/cut percentages only, and no
real multi-month buy-rate figure was on hand to check one way or the
other; flagged in the function's own docstring for the first real
multi-month Excel export to settle. Applied ONCE, in app.py, immediately
after parsing -- `report_assembly`'s fill/facts code carries no months
parameter of its own and assumes whatever `PolkExport` it's handed is
already normalized. **App wiring:** a required "Months covered by this
Polk report" number input (`value=None`, so it starts genuinely blank, not
defaulted to 1) appears the moment a Polk file is uploaded; Generate is
**hard-blocked** with a named error until it's filled in -- guessing 1
would be wrong for the common multi-month case and indistinguishable from
a rep who hasn't looked at the field yet. The 30-day-sales-window date
("Website attribution: Jul 1 - Aug 31 · Polk sales: through Jul 31 (30-day
sales window)") is a rep-typed **optional** date field
(`attr_polk_sales_through`) -- **flagged assumption**: the Polk export
itself carries no date fields anywhere in its 19 tabs, so there was no way
to derive "through" automatically; Matt's own spec didn't call out a new
input for this either, so this is this phase's own best read of what's
needed, not something explicitly confirmed. Left blank, the recap slide's
new caption degrades to a period-agnostic sentence rather than a
fabricated date.

**2. Top-line tiles -- built, template NOT yet built (v0_13 handoff
below).** Four tiles: Target Dealer Sales, Total MSRP Sold (`report_
assembly.total_msrp_sold`, sum of avg_msrp * models_sold across every
make/model row), Campaign Lift (dropped and reflowed when lift <= 0), ROI
(dropped and reflowed whenever the toggle is off or there's no cost).
Matched Households/Buy Rate move into the `PolkMatchRateNote` caption
beneath, alongside the match rate itself. **v0_12 compatibility, built in
from the start:** until Matt's v0_13 template exists, the fill code keeps
writing `POLK_MATCHED_HOUSEHOLDS`/`POLK_BUY_RATE` (the two retired
standalone tiles) so a report generated against the still-live v0_12
template doesn't show broken literal `{{...}}` tokens the moment this code
shipped -- those two tokens simply stop matching anything once v0_13
removes the shapes, a silent no-op, same convention as every other
template transition in this deck.

**MSRP IS projected, confirmed 2026-09-21.** The original build left MSRP
un-projected (a flagged assumption -- no real fixture with a stated true
total was on hand). Matt's own call: when "Project for match rate" is on,
MSRP sold moves by the SAME factor as sales/households, or the three
tiles disagree with each other side by side (5 sales projecting to 6
while the vehicles those 5 sales represent stayed un-projected reads as
internally contradictory). `total_msrp_sold` itself stays a pure,
unprojected sum; `_fill_automotive_registrations` applies `project_
for_match_rate` to it exactly like it already does for sales/households,
and carries the same "(projected)" suffix. `facts["polk"]["projected_
msrp_sold"]` mirrors `projected_matched_households`/`projected_target_
dealer_sales`. **Buy rate stays deliberately un-projected either way --
confirmed correct, not just assumed:** Matt's own real multi-month decks
show it as a ratio of two already-stacked totals (19/126,056 = 0.02%,
15/131,783 = 0.01%), not a count that itself needs floor-correcting.

**3. ROI -- built.** `report_assembly.compute_roi(sales, profit_per_
vehicle, cost)` returns `{sales, profit_per_vehicle, cost, gross_profit,
net_return, multiple}`, or None with no cost (nothing to divide against).
App wiring: "Include ROI" toggle (default off, appears only once a Polk
file is uploaded), a "Net profit per vehicle" input (default $3,000), and
a "Campaign cost for this period" input **pre-filled from the linked
proposal's own CTV/OTT plan rows for the months entered above**
(`report_assembly.estimate_plan_cost_by_channel`, see item 4 below for
what "pre-filled" means and doesn't mean) -- confirm or edit either way,
exactly as specced. Sales used for ROI honors the existing "Project for
match rate" toggle (a raw or match-rate-divided figure, the rep's own
call, not a second independent toggle).

**4. Cost per visit -- built, independent of Polk, live for every report
with a linked proposal.** `report_assembly.compute_cost_per_visit(ctv_
cost, attributed_unique_visitors, retargeting_cost, retargeting_clicks)`
returns CTV-cost-per-attributed-visitor and retargeting-cost-per-click as
two **separate** keys, never blended -- a plan running both products
produces two real, distinct answers, and blending them answers neither
question. Retargeting's own click count comes from the OTT retargeting
export's `clicks` field (the same third, separate upload `report:
ott_retargeting` already uses). App wiring: "Include cost per visit"
toggle (default off), separate CTV/OTT and retargeting cost inputs, each
pre-filled from the linked proposal. Rendered as `CostPerVisitNote` on
`report:highlights` (same "PlanVsActualNote" caption-under-a-KPI-row slot
style) and threaded into the case study's existing `CS_TILE_3` ("Cost Per
Visitor") as the preferred, CTV-only figure -- `build_case_study_slide`
gained a new `ctv_cost` parameter that takes precedence over the older,
blended `budget` parameter when both are supplied; `budget` alone (pre-
Phase-8 behavior) is the fallback, so nothing already shipping changes for
a campaign with no retargeting product in play. **`estimate_plan_cost_by_
channel(plan_rows, months=1)`** is the shared pre-fill for both ROI and
CPV -- buckets a linked proposal's own plan rows by tactic-label prefix
("Streaming Retargeting"/"Site Retargeting" vs. everything else) and sums
each bucket's own stored row cost. **Explicitly NOT exact date-range
math** -- it multiplies each row's own stored monthly (or full-flight)
cost by a month count; Matt's own spec says "confirm or edit" for both
toggles, so this is a rough, rep-verified pre-fill, the same precision the
case study's pre-Phase-8 `budget` parameter (option 0, no finer date math)
already shipped with, not a new, more exact mechanism.

**Gap closed 2026-09-21: every confirmed Phase 8 input is now persisted
onto the logged report.** `facts["phase8_inputs"]` (`app.py`, written at
Generate-time logging, right alongside `headline_facts`/`draft`/
`optimizations`) carries `polk_months`, `polk_sales_through` (ISO string
or None), `polk_roi_on`/`polk_profit_per_vehicle`/`polk_roi_cost`,
`cpv_on`/`cpv_ctv_cost`/`cpv_retargeting_cost` -- read from session_state
directly (not the widgets' own local variables, which simply don't exist
in a run where the toggle that gates them is off) so the dict is always
present, every field None/False when there was nothing to confirm, same
"always the same key" convention `facts["polk"]` already follows. **The
case study is the first, and so far only, consumer**:
`_render_case_study_review` reads `report_json["phase8_inputs"]
["cpv_ctv_cost"]` and passes it as `build_case_study_slide`'s `ctv_cost`
-- taking precedence over the proposal's own blended `budget` per that
parameter's own precedence rule, so CS_TILE_3 shows the rep's confirmed,
CTV-only figure without asking again. Verified two ways: `tests/test_
attribution_reports_page.py`'s `check_polk_phase8_persistence` (a linked
proposal with a REAL, different blended budget -- $50,000 -- specifically
so the precedence is actually exercised, not just "nothing to prefer
over"; confirms `$9,000 / attributed_unique_visitors` wins, never
`$50,000 / attributed_unique_visitors`) and the live walkthrough below
(same numbers, through the real running app against real Supabase).
Months/sales-through/ROI inputs are persisted too but have no second
consumer yet (nothing else currently re-reads a logged Polk report) --
persisted now so the day a second consumer exists, it doesn't have to
re-ask either.

**5. Polk cuts -- built.** `_polk_share_callout` now names creative and
audience only, **never publisher** (Matt's own explicit removal), and only
when there's genuinely more than one value in that dimension (a single
audience/creative is nothing to compare against, so the clause is skipped
entirely rather than saying "X led all audiences" about the only audience
that ran).

**6. Target dealers -- built.** `report_assembly.polk_dealer_table_rows`
merges the Target Dealer(s) tab (market rank vs. campaign rank, but only
the advertiser's own target-dealer group) with the All Dealers tab
(broader, campaign_share only) into ONE list by dealer name, ranked by
sales. Columns: Dealer | Sales | MSRP Sold | marker ("Client"/"Group"/
blank). **The client's own dealership(s) are ALWAYS included, even outside
the row cap** -- the whole point of the table is showing the client
against competitors, so a client row silently falling off the bottom would
defeat it; halo-group sibling rows get NO such guarantee and follow the
ordinary cap (their combined sales still reach the reader via the
aggregate "N sales landed at other group stores" narrative line even when
no individual sibling row makes the table). Dealer identity is rep-
confirmed, never fuzzy-matched: a multiselect of every dealer name in the
file, pre-guessed from a case-insensitive substring match against the
client's own name, confirm or correct. Halo grouping is a manual checkbox
+ textarea (one dealer name per line) -- **no prefill from the Auto-Sales
Analyst deck's site list yet**, since that needs the Analyst's own facts
export (item 7, still pending on Matt). **v0_12 compatibility**: the fill
code checks the live template's own `PolkTargetDealersTable` column count
at render time and only uses the new ranked-by-sales/4-column query when
the table actually has 4+ columns -- v0_12's table is still 3 columns
(Dealer/Market Rank/Campaign Rank), so it keeps getting the OLD Phase 7
query (target dealers only, ranked by campaign rank) until v0_13 widens
it. This is deliberate: naively sending the new field list at the old
column count would have positionally written Sales/MSRP numbers into
cells still headed "Market Rank"/"Campaign Rank" -- wrong content under a
real header, worse than a warning.

**7. Cross-source takeaways -- NOT built, still gated, but the schema question is now closed.** The Auto-
Sales Analyst deck is still appended wholesale (`extra_deck_path`,
unchanged since Phase 7's own predecessor feature) with nothing parsed out
of it -- no `facts["analyst"]`, no cross-source thread logic. **Gate, updated 2026-09-23**: the Analyst now ships a real facts JSON (`sample_facts.json`, a genuine export -- Ted Britt Group, 3 sites, June 2026 -- placed in the project folder, gitignored like every other real client file). What's still missing is a real WEBSITE ATTRIBUTION export for the same dealer and overlapping period -- see the Ted Britt status note below. Not built yet on purpose; this section now exists so the schema and matching rules are settled and ready the moment that second file lands, rather than being worked out from scratch under time pressure.

**Actual shipped shape, from the sample file itself -- the Analyst shipped first, so this repo adapts to its field names, not the other way around.** Top level: `schema_version` (1), `meta` (`generated_at`, `app_build`, `report_id`, `is_group`, `site_count`, `inventory_scanned_at`, `analysis_period` -- `{start, end, source}`, ISO dates -- `vdp_visit_threshold`, `group_name`, `lookback_days`), `sites` (one entry per dealer: `site_id`, `dealer_name`, `domain`, `url`, `brands_observed`), `totals`, and `by_site` (one key per `site_id`, each holding the SAME shape as `totals`). Both `totals` and each `by_site` entry carry: `vehicles_shopped`, `vehicles_sold` (+ `_new`/`_used` splits and their own `look_to_book_pct`), `est_revenue_sold`, `est_pipeline_value`, `avg_sold_price_est`, `visits_total`, `unique_visitors`, `avg_pages_per_visitor` (totals only -- not present per-site in the sample), `sold_by_make` (`[{make, count}]`), `sold_by_make_model` / `top_sellers` (`[{make, model, label, count, raw_name?}]` -- `raw_name` only present when the Analyst's own normalization changed the string, e.g. `"Ford F 150"` -> label `"FORD F-150"`), `sold_by_price_tier` (`[{tier, label, count}]`, tiers seen: `BUDGET_30K`/`CORE_30K_60K`/`PREMIUM_60K`), `shopped_vehicle_status` (`{available, sold, unverified}` -- `unverified` was `null` in the sample, so treat as optional/nullable, not guaranteed present), `traffic_mix` (`[{category, label, visits}]` -- **raw visit COUNTS, not a share/percentage** -- categories seen: `NEW_VDP`/`USED_VDP`/`NEW_CAR_SEARCH`/`SERVICE`/`HOMEPAGE`), and `missed_opportunities` (`{count, vehicles: [{make, model, label, visits, vin}]}` -- no free-text note field).

**What the Analyst's export does NOT compute -- absent from the schema, not present-as-zero: days-on-lot, body style, or a full inventory mix.** The export is demand/sales-signal-shaped (what shopped, what sold, what visited) rather than inventory-shaped -- there is no per-VIN lot-age or body-style breakdown anywhere in `totals`, `by_site`, or the per-vehicle `missed_opportunities` entries (which carry `vin` but nothing about the vehicle's own listing). Don't build a fact or a drafting rule that assumes any of these three exist; treat their absence as permanent until the Analyst's own export changes, not as a gap this repo's parsing missed.

**Matching rules, recorded now for when this is built:**
- **Match Polk vehicle-sale rows against Analyst `sold_by_make_model`/`top_sellers` on make + model FAMILY, never full trim.** The sample's own data is the worked example: `"GLE 53 4MATIC"` has to match Polk's own registration data under the `GLE` family, not fail to match because Polk records a bare `"GLE"` or a different trim string for the same nameplate. **Spot-check luxury nameplates specifically** -- trim proliferation (AMG lines, 4MATIC/xDrive/quattro badging, performance packages) is exactly where a full-trim match silently drops real matches, and it's disproportionately the luxury makes (Mercedes-Benz here) where that shows up first.
- **`analysis_period` (in `meta`) is required on the Analyst side, and must OVERLAP the report's own period, or the Analyst facts are not used for cross-source threads at all.** The sample's own `analysis_period` is `2026-06-01` to `2026-06-30` -- a report built for a non-overlapping window has nothing to cross-reference and should silently skip Analyst-sourced threads, the same "absence of data is never a finding" rule the rest of drafting already follows, never a fabricated one from a mismatched period.
- **`totals` vs. `by_site`: use `totals` before Phase 9.** Both carry the identical shape (confirmed directly against the sample: same field set, `by_site` just keyed per `site_id`), so a per-site breakdown is already there in the data the moment a multi-dealer group report (Phase 9) needs it -- but until then, a single-report cross-source thread reads `totals` only, never sums or re-derives across `by_site` entries itself.

**`sites` doubles as the halo-group prefill list (item 6) once this is built** -- a rep would still confirm/edit it via the same textarea, just pre-filled with the real `dealer_name`s instead of starting blank. `sold_by_make_model`/`top_sellers` is what lets a thread connect "F-150 interest showed up in website pages and in Analyst sales" -- the prompt rule Matt specced ("a thread that cites the same make/model or intent across two or more sources ranks first among signals") is still not added to the drafting prompt, since there's no `facts["analyst"]` for it to reference yet; adding the rule ahead of the data would teach the model to reference a key that's always absent.

**Ted Britt now has two of the three sources this feature needs: Polk (a real 2-month dashboard, see Phase 8's own verification above) and the Analyst (`sample_facts.json`, June 2026).** A Ted Britt WEBSITE ATTRIBUTION export for a period overlapping June 2026 is what unblocks building this for real, against a genuine three-source client rather than a synthetic stand-in -- the same "don't build against synthetic data" discipline Phase 9 itself is already gated on.

**Handoff to Matt: v0_13 template changes, confirmed.** Matt asked for
tile/token names to be confirmed before building; confirmed as specced,
no changes:

| Element | Change |
|---|---|
| `PolkHouseholdsTile` | **Removed** (households moves into `PolkMatchRateNote`) |
| `PolkBuyRateTile` | **Removed** (buy rate moves into `PolkMatchRateNote`) |
| `PolkSalesTile` | Kept, same token `{{POLK_TARGET_DEALER_SALES}}` |
| `PolkMsrpTile` | **New**, token `{{POLK_MSRP_SOLD}}` |
| `PolkLiftTile` | Kept, same token `{{POLK_CAMPAIGN_LIFT}}` -- now deleted+reflowed by the fill code whenever lift <= 0 |
| `PolkRoiTile` | **New**, token `{{POLK_ROI}}` -- deleted+reflowed whenever the ROI toggle is off |
| `PolkMatchRateNote` | Kept, same token `{{POLK_MATCH_RATE_NOTE}}` -- now also names households/buy rate, not just match rate |
| `PolkTargetDealersTable` | **Widen to 4 columns**: Dealer \| Sales \| MSRP Sold \| marker (a blank/"Client"/"Group" text cell) -- same `{{POLK_TARGET_DEALER_ROWS}}` token, one header row + one template data row per this deck's usual convention |
| `PolkTargetDealersHeader` | Kept -- consider renaming the visible label from "Target Dealer Performance" to something like "Dealer Performance" since the table is no longer target-dealers-only, wording is Matt's call |
| `PolkNarrative` | Kept, same token `{{POLK_NARRATIVE}}` -- now composed from creative/audience share (never publisher) plus a competitor/halo-group sentence |

Two more shapes needed elsewhere in the deck, not on `report:
automotive_registrations`:

| Slide | New shape | Token | Notes |
|---|---|---|---|
| `report:recap` | `PolkSalesWindowNote` | `{{POLK_SALES_WINDOW_NOTE}}` | Plain text box, same slot style as an existing tile-row caption; deleted outright when no Polk file is attached to the report at all |
| `report:highlights` | `CostPerVisitNote` | `{{COST_PER_VISIT_NOTE}}` | Plain text box, same "PlanVsActualNote" slot style (a short caption under the KPI tile row); deleted outright when the CPV toggle is off or has nothing to show |

**Guards:** `tests/test_report_assembly.py`'s `check_polk_phase8` (pure
functions: `normalize_for_months`, `total_msrp_sold`, `compute_roi`,
`compute_cost_per_visit`, `polk_dealer_table_rows` including the client-
always-included/sibling-no-guarantee distinction, plus a v0_12-compat
column-count check) and the existing `check_polk_automotive_registrations`
(updated for the new narrative/note text). `tests/test_attribution_
reports_page.py`'s `check_polk_phase8_wiring` (the months gate actually
blocks Generate, ROI/CPV toggles surface their own fields, a real Generate
click succeeds once the gate is satisfied). Full offline suite green:
`tests/test_report_assembly.py` 261/261, `tests/test_polk_import.py`
34/34, Tier 1 drafting regression 374/374.

### Live walkthrough -- 2026-09-21, closes this project's own standing rule for Phase 8

Driven through the real running app (`PROPOSAL_BUILDER_TEST_MODE=1`, real
Chrome automation, real Supabase -- not AppTest) with a throwaway test
advertiser ("Phase 8 Walkthrough Test Co," deleted afterward, per this
project's own "never a real advertiser" rule): the real `MW attribution
excel.xlsx` client name was swapped surgically (a direct `sharedStrings.xml`
zip-entry replace -- an `openpyxl` full load/resave round-trip was tried
first and silently dropped the weekly-trend tab's data, a real, separate
`openpyxl` fidelity gotcha worth remembering, not a Phase 8 bug) and the
real `Polk Dashboard.xlsx`.

**Every new input was driven, not just unit-tested:**
- **Months gate**: Generate blocked with the exact named error before a
  month count was entered; entering 1 against a synthetic match-rate-
  stacked fixture (414.9%, mirroring Matt's own real multi-month figures --
  no real multi-month Excel was on hand) fired the sanity warning live in
  the browser -- **a real gap found and fixed in the same pass**: the
  warning was computed by `normalize_for_months` but never actually shown
  anywhere in the UI. Fixed: surfaced unconditionally the moment a month
  count is entered, same `st.warning` convention the attribution export's
  own warnings already use.
- **Sales-through date, dealer multiselect (pre-guessed then confirmed),
  halo-group textarea, ROI (profit per vehicle + cost, edited from its
  $0 no-proposal default), Cost Per Visit (separate CTV/retargeting costs,
  both edited)**: all driven, all round-tripped correctly through Generate.
- **Post-Generate state screenshotted** (both the report-builder page and,
  separately, the case-study review it fed): the built deck's
  `report:automotive_registrations` slide (v0_12 compat path) showed the
  correct unprojected sales tile ("5"), the richer `PolkMatchRateNote`
  sentence, the OLD 3-column dealer table untouched (compat branch
  verified live, not just by column-count assertion), and a narrative
  correctly omitting publisher. The drafted highlight bullet cited MSRP
  sold as **exactly** $254,146 (2*30613 + 1*64700 + 2*64110, computed by
  the model from the facts payload, not invented) and named the 3-Ford/
  2-Lincoln sales split from `target_dealers` unprompted. The known,
  tracked "100%" model-behavior gap (Phase 7 section) recurred, confirmed
  still unrelated to Phase 8's own code.
- **Case study**: built from the logged report with NO re-asking --
  CS_TILE_3 read $9,000/4,638 = "$2," the persisted CTV cost, not the
  $50,000-if-a-proposal-had-been-linked blended figure (no proposal was
  linked in this pass, so precedence was proven in AppTest against a real
  differing budget instead -- see item 4's persistence writeup above).

**Cleanup:** the test report was hard-deleted from Report History before
ending the session (confirmed refused-if-a-case-study-was-saved -- none
was); the throwaway advertiser record was left in place (harmless, same
as every other test advertiser this project's history already
accumulated); local scratch fixtures were deleted, never committed.

Then **Phase 9 -- multi-dealer group reports**, per Matt's own outline
("multiple advertisers in a group" toggle, per-dealer upload sets, a
separate-vs-combined choice) -- not proposed yet. **Still open before
Phase 9 can be proposed against Phase 8's finished shape:** (a) Matt
builds v0_13 per the handoff table above -- **done, see below**; (b)
verify the real 414.95%/474.42%/400%/200%/600% figures against the two
dashboard PPTX exports once they're actually placed in the project folder
(today's walkthrough used a synthetic stand-in for the sanity-check UI
verification, not a real multi-month file) -- **the 166.3%/83.15%
stacking math was separately verified this same day against a genuine
2-month Polk file, see below, closing the spirit of this item even though
the original PPTX pair still isn't in the folder**; (c) Matt supplies the
Auto-Sales Analyst facts JSON per the proposed shape above, closing item
7.

### v0_13 landed and verified against a real multi-month client (Ted Britt) -- 2026-09-21

Matt built `REPORT_MASTER_v0_13.pptx` from v0_12 per the handoff table
above (`build_v0_13.py`, committed) -- structurally verified against the
spec directly via python-pptx (every renamed/new shape, the 4-column
dealer table, `PolkSalesWindowNote`/`CostPerVisitNote` present on their
named slides) before any report was built against it. Uploaded and
activated as `report_deck_versions` id 11.

**Matt supplied a genuine multi-month real client for testing**: Ted
Britt Ford & Ted Britt Chantilly, all four export types (attribution,
delivery, retargeting, Polk) for Jul-Aug 2025, plus real monthly spend
($17,500 OTT + $3,150 retargeting) -- the first real 2-month Polk file
this project has had access to. `polk_import.normalize_for_months`
against it: raw `match_rate` 166.3% over 2 months normalizes to a sane
83.15% (`normalize_for_months(export, 2)`); forcing it through as if it
were 1 month still shows >100% and correctly re-raises the sanity
warning (`normalize_for_months(export, 1)`) -- this is the real-data
verification the "still open" item above was waiting on, even though the
literal two PPTX exports named in that item are still not in the folder.

**Two real bugs found and fixed, both from this being the first real
export of its kind, not synthetic:**
1. **`report:response_profile` hard-raised** (`MissingTokenError`,
   blocking Generate outright) the first time an attribution export
   genuinely carried neither a recency nor a referral-domain tab -- every
   fixture on hand when that slide was built (2026-09-10) happened to
   have both. Fixed with `report_assembly.response_profile_applies()`,
   the same drop-the-slide-when-nothing-to-show pattern already used for
   `attribution_breakdown_applies`/`delivery_breakdown_applies` -- wired
   into both the slide-drop pass and the `_fill_response_profile` call
   site in `build_report_deck`. Verified the slide is present for WAEPA
   (which has both tabs) and correctly absent for Ted Britt (which has
   neither), via a real `build_report_deck` call against both fixtures,
   not just a unit check of the predicate.
2. **The facts-only drafting checker (`app._attr_payload_numbers`) false-
   flagged a correctly-drafted negative number.** Ted Britt's ROI
   genuinely came back negative (net_return -$23,300, later -$30,476
   projected -- see below), and the model correctly wrote "-$23,300 net
   return," but `_ATTR_NUMBER_RE` (`r"\d[\d,]*(?:\.\d+)?%?"`) starts
   matching at a digit and never captures a leading `-`, so it extracted
   the bare magnitude "23,300" and found no unsigned match registered for
   it -- `add_int()` had only ever registered the signed string forms.
   Fixed by also registering the unsigned magnitude whenever a traced
   number is negative. Guard: `check_negative_number_traced_by_magnitude`
   (`tests/test_attribution_draft.py`), covering both a citation that
   keeps the model's own minus sign and one reframed as a sign-free "net
   loss."

**MSRP's projected form fits one line, confirmed two ways** (the specific
question Matt asked before building): measured with `text_metrics.
text_width_points` against the real rendered run (`"$387,731 (projected)"`
at 16pt bold Arial measures 2.14in against a 2.69in box), and then
confirmed against the actual PowerPoint renderer (`deck_render.
render_deck`, real COM render, not python-pptx's own layout guess) --
single line, comfortably inside the tile. **One thing that did NOT fit
cleanly on first render**: the ROI tile's longest real string
(`"$-30,476 net (projected) (0.3x)"`, 32 characters) measured 3.91in
against the same 2.69in box at 20pt -- wider than the tile, landing as two
snug lines via the shape's own `auto_size=TEXT_TO_FIT_SHAPE`. **Fixed the
next day, not left as a known-tight edge case** -- see "ROI tile shows the
multiple only" below, which removes the dollar figure from the tile
entirely and closes this from a different direction than a narrower fix
would have.

**Also found while filling the ROI/CPV inputs live: "Net profit per
vehicle" defaults to $3,000 (a real per-vehicle profit figure,
`app.py`'s own `attr_polk_profit_per_vehicle` default) and is a
DIFFERENT field from "Campaign cost for this period" -- an operator
mistake during this same walkthrough (typing the $41,300 campaign total
into the profit-per-vehicle field after a mid-form UI reflow) produced a
nonsensical ROI on the first pass, caught only by sanity-checking the
number rather than by any app defect. Not a code bug -- recorded here
because the two fields sit close together and a rep moving fast after a
page reflow could make the same mistake; the field's own tooltip already
names what it's for.**

Two real Ted Britt Polk decks (projection toggle OFF and ON) and a WAEPA
deck with cost-per-visit on were built through the real running app,
verified via python-pptx content checks plus real COM renders of the key
slides (automotive_registrations, recap, highlights), and delivered to
Matt directly. `tests/test_report_assembly.py`'s
`check_v0_13_and_ted_britt` guards the v0_13 shape structure, the real
stacking-math sanity check, and `response_profile_applies` against both
real fixtures permanently; full offline suite green after (289/289 in
`test_report_assembly.py`, 34/34 in `test_attribution_draft.py`, 374/374
Tier 1).

### Follow-up round, same day -- stacking math confirmed on all fields, ROI tile redesigned, profit-per-vehicle guard added

Matt's review of the v0_13 landing above asked for four things, all
closed the same day.

**1. Stacking math confirmed across every field, not just match_rate --
the synthetic stand-in retired.** The Ted Britt real Polk file's raw
figures: match_rate 166.3%, several individual cut shares also genuinely
over 100% (audience_by_sales 163.84%, creative_by_sales 153.67%,
publisher_by_sales' Pluto TV row 157.06%, among others) -- confirming the
premise this whole feature exists for is real, not assumed. After
`normalize_for_months(export, 2)`: match_rate 83.15%, and **every single
one** of match_rate/campaign_lift/every audience/creative/publisher cut
share (impressions and sales views, both) lands at or under 100% --
`check_v0_13_and_ted_britt` now asserts this exhaustively (every share in
every cut, not a spot check), and `check_polk_phase8`'s old synthetic
`PolkExport(match_rate=4.1495, ...)` stand-in (built from Matt's own
illustrative 414.95% example, before any real multi-month file existed)
is retired -- the exact same shape of retirement as the MW-delivery
synthetic stand-in once `MW delivery.xlsx` was found. `buy_rate` reconfirmed
untouched by normalization on the real file too.

**2. ROI tile redesigned to show the multiple only, net dollars moved to
narrative/facts, sign-formatting bug fixed everywhere.** `PolkRoiTile`
now renders exactly `"0.4x"` (or `"0.4x (projected)"`) -- no dollar
figure, no "net", nothing that can run long enough to wrap. The net
return still reaches the rep: `facts["polk"]["roi"]["net_return"]` is
unchanged, and the drafting prompt (`app.py`'s Polk paragraph) was
rewritten to require the model state it explicitly, since the tile no
longer will -- verified live: Ted Britt's real negative case drafted
*"Against a campaign cost of $41,300, the 6 confirmed sales produced
$18,000 in gross profit, yielding a -$23,300 net return (0.44x) at $3,000
profit per vehicle,"* with the takeaways slide correctly framing it as an
early-campaign matched-sales floor rather than hiding or softening it.
Separately, `report_assembly._money()` had a real sign-formatting bug --
`f"${amount:,.0f}"` puts Python's own sign after the literal `$` for a
negative amount ("$-30,476"), never before it -- fixed to sign-first
("-$30,476"), the only reading a rep would write by hand. **Negative ROI
is never auto-hidden** (the rep turned the toggle on deliberately) but
now surfaces in the "Review before sending" panel --
`attr_actionable_review_items` flags any `multiple < 1.0` with *"ROI is
below 1.0x at the entered cost and profit per vehicle -- confirm before
sending, or turn ROI off,"* the same rep-actionable standard as the
existing conversion-definition and prior-period items in that panel, and
verified live to actually appear after a real Generate click on Ted
Britt's own negative case.

**3. Profit-per-vehicle bound-check guard, closing a real operator slip
from the walkthrough.** Two mitigations, both live-verified: the field is
relabeled "Net profit per vehicle sold" (help text states the $3,000
default plainly and names the OTHER field by name, "NOT the campaign's
total spend"), and it now sits in its own `st.container(border=True)`
card, visually separated from "Campaign cost for this period" in its own
card below -- the exact stacked-plain-fields layout that let a $41,300
campaign total get typed into the profit field during the prior
walkthrough. A bound check (`polk_profit_per_vehicle > 15000` OR within
10% of `polk_roi_cost`) fires a plain `st.warning` -- *"This looks like a
campaign total, not profit per vehicle..."* -- reproduced live by
deliberately re-typing $41,300 into the profit field: warning fired;
corrected back to $3,000 with the real $41,300 cost in its own field:
warning cleared, Generate proceeded clean.

**4. response_profile confirmed to drop, not render empty** -- re-verified
this round, unchanged from the prior fix: absent from both the real Ted
Britt build (no recency/referral tabs) and the live walkthrough deck's
own slide-key list; present for WAEPA (which has both tabs), so the check
isn't vacuously true.

Guards: `check_polk_phase8`/`check_v0_13_and_ted_britt`
(`tests/test_report_assembly.py`, including the new `_money` sign-fix
unit checks, the losing-ROI `compute_roi` case, and a real end-to-end
`build_report_deck` call asserting the built deck's own
`PolkRoiTileValue` text is exactly `"0.4x"`), `check_actionable_review_
items` (`tests/test_attribution_draft.py`, negative/positive/no-Polk ROI
cases). Full suite green: 299/299 `test_report_assembly.py`, 37/37
`test_attribution_draft.py`, 374/374 Tier 1, and the full
`test_attribution_reports_page.py` AppTest suite (including
`check_polk_upload_and_projection_toggle`/`check_polk_phase8_wiring`).

### Deferred, named for continuity

- **Optimization engine — LANDED 2026-09-15** (`cb6276c`, `cad0900` — see
  the two new dated sections above), moved out of this list. The one
  still-open sub-item survives unchanged: whether the Premion dashboard
  can pull a full daily time series for the engine's dimensions beyond
  weekday — Matt is checking that separately; not blocking, since the
  engine ships and runs without it.
- **Create case study from report — LANDED 2026-09-16** (`e8e652f` — see
  the new dated section above), moved out of this list.
- **One-slide report summary — LANDED 2026-09-16** (`e8e652f`, same
  commit and same template family as the item above, per the original
  "build them alongside each other" call below), moved out of this list.
- **Auto-Sales-Analyst integration — LANDED 2026-09-08** (was addendum §4-9,
  moved out of this list). See the new dated section below.
- **Match-rate projection toggle — SPECCED 2026-09-08, build with Polk.**
  Polk (`Polk_Dashboard.xlsx`) and sales match-back share the same shape: a
  match rate under 100% means every matched outcome (sales, matched
  households, matched impressions) is a FLOOR, not the true count — the
  unmatched remainder genuinely happened, it just couldn't be tied back.
  Real numbers from the Polk file on hand: 90.49% match rate, 5 Target
  Dealer Sales, 44,756 matched households, 591,686 matched impressions.

  **The toggle: "Project for match rate," default OFF.** Off (today's only
  behavior once built) shows the raw matched figures, with the match rate
  itself stated once nearby so the reader knows it's a floor, not a total
  — never silently presented as complete. On divides every matched outcome
  by the match rate (`projected = matched / match_rate`) and labels the
  result **"projected"** everywhere it appears — the tile, the table, AND
  the narrative sentence that cites it — never a bare number that looks
  like a fact. This is the exact same discipline the sports pacing tile
  and the pixel-issue-window warning both already follow: a number that
  isn't simply "what the export says" gets said out loud as such, never
  slipped in silently.

  **One toggle serves BOTH Polk and sales match-back** — build it once,
  reuse it for whichever attribution source is in play, rather than two
  near-identical toggles with the same shape. The natural home is a small,
  pure helper (`project_for_match_rate(value, match_rate)` or similar,
  mirroring the discipline `combined_headline_impressions`/`_overlaps_
  pixel_issue_window` already set — pure function, no Streamlit, easy to
  unit-test against real Polk numbers) plus a checkbox next to wherever
  Polk's own tiles render once that slide exists. **Dependency: the Polk
  phase itself has to land first** — there's no matched-outcome tile to
  attach the toggle to yet. Nothing further to build today; this paragraph
  IS the spec, ready to implement the moment Polk is scheduled. **Note:**
  this section originally called Polk "Phase 6" — Phase 6 has since
  shipped as the advertiser spine (above) instead; see Phase 7 below for
  where Polk actually landed in the sequence.
- ~~**Polk — file in hand, inventory only (2026-09-08).**~~ **Superseded
  2026-09-16 by Phase 7 above**, which re-verified the inventory directly
  against the file (the real filename is `Polk Dashboard.xlsx`, a space,
  not the underscore this bullet originally had) and gave it a scope and
  a phase number.
- **Sales match-back, brand lift, Arrivalist** (addendum §5-9) — still
  fully deferred, each gated on a real raw deliverable from Matt, per the
  addendum. Sales match-back shares the match-rate toggle's design above
  once its own deliverable arrives.
- **Phase 9 — multi-dealer group reports. NOT PROPOSED, don't build
  against synthetic data.** **Gate:** a real multi-dealer group report
  set — several dealers' website exports, Polk file(s), and Analyst decks
  for one group, one period.

  The need, as Matt described it: a "Multiple dealers in this group"
  toggle — when on, each upload section (website, delivery, Polk,
  Analyst) accepts one file **per dealer**, each named. **Separate vs
  combined**: Separate builds one deck per dealer, as today; Combined
  builds one deck — group-level stats and takeaways across all dealers,
  then per-dealer highlights. Group figures **sum counts and recompute
  rates from the sums — never average rates across dealers** (the same
  discipline `normalize_for_months`/`combined_headline_impressions`
  already follow for a different kind of aggregation). Client-dealer
  markers and halo siblings carry through; in combined mode, cross-dealer
  sales are the group's own halo story.

  Open questions for when it starts, none resolved yet:
  - Per-dealer upload state and `report_json` shape; regenerating a
    logged combined report (summaries/case studies must still work
    against it).
  - Whether one Polk file already covers several dealers (the Target
    Dealers tab already lists several) — per-dealer Polk uploads may be
    optional rather than required.
  - Combined deck slide order: group recap/highlights first, then
    per-dealer sections — which slides repeat per dealer, which stay
    group-only, and what that costs Matt in template work.
  - The optimization engine and series/MoM tracking: per dealer, per
    group, or both — not yet decided which is the primary read.
  - The Clients page: one group client with dealers under it, or several
    linked clients — `advertisers` may need a parent/child relation it
    doesn't have today.
- **Auto-Sales Analyst facts JSON → cross-source takeaways. NOT BUILT --
  the schema question is CLOSED (2026-09-23), the gate is narrower now.**
  The Analyst shipped its facts JSON (item 7 earlier in this doc has the
  full actual shape, from a real sample -- `sample_facts.json`, Ted Britt
  Group, 3 sites, June 2026 -- not the old proposed placeholder shape this
  paragraph used to point at). **Remaining gate:** a real report where
  website attribution, Polk, and Analyst data all exist for the same
  dealer and an overlapping period. Ted Britt itself is now the closest
  candidate -- it already has Polk (a real 2-month dashboard) and the
  Analyst (June 2026); a Ted Britt website attribution export for a
  period overlapping June is the one remaining piece.

  Make/model names in the real export are ALREADY normalized the way
  this was hoping for -- `label` is uppercase/canonical (`"FORD F-150"`)
  with `raw_name` carrying the pre-normalization form only when it
  differs, so matching against Polk registrations and website page names
  can read `label` directly, no second normalization pass needed here.
  **Matching is by make + model FAMILY, not full trim** (item 7's own
  matching-rules subsection has the worked example and the luxury-
  nameplate caveat) -- record that now rather than re-deriving it when
  this actually gets built. Prompt rule to add at the same time as the
  parsing: a thread citing the same make/model or intent across two or
  more sources (website + Analyst + Polk) is the strongest finding and
  ranks first among signal threads — "F-150 interest showed up in
  website pages, the inventory analysis, and Polk sales" is the sentence
  this rule exists to produce. Still not added to the drafting prompt:
  teaching the model to reference `facts["analyst"]` before that key can
  ever be present would train toward a key that's always absent.

### Auto-Sales Analyst deck append — landed 2026-09-08

A real Analyst summary export (`Auto Group (5 Sites) - 10_51 AM ET_Summary
(1).pptx`) is 13.333×7.5in, 8 slides, image charts, no native chart parts
— confirmed by inspection, which settles the addendum's own fork in favor
of **ALIGN**: it appends cleanly through the existing cross-deck copy
mechanism (`assembly.copy_slide_into`/`ImportCache`, the same primitives
case studies and the slide vault already use in the proposal builder) with
none of that mechanism's own known hazards (no chart workbook to carry, no
`get_or_add_image_part` re-encoding risk). No parsing, no facts extracted,
nothing absorbed into the report's own payload — it rides along exactly as
uploaded.

Built as `report_assembly.append_slide_deck(prs, path)` plus a new
`extra_deck_path` keyword on `build_report_deck`, called as literally the
LAST step before `prs.save` (after every token-fill call, so an appended
slide can never shift a `_slide_by_key` lookup this function still needs
to make). The upload slot ("Auto-Sales Analyst deck (optional)") sits on
the Attribution Reports page next to the attribution/delivery uploaders,
saved to a scratch path and threaded through to Generate exactly like
those two, with no parsing step of its own.

Confirmed against the real 8-slide file that the append path has no
one-slide assumption anywhere in the anchor logic: the deck's own slide
count grows by exactly 8, in order, structurally clean under
`package_check.check_package` — the same gate `tests/test_slide_vault.py`
already uses for the proposal builder's own cross-deck grafts. Guard:
`tests/test_report_assembly.py`'s `check_auto_sales_analyst_append`.

### Known pixel-issue window — landed 2026-09-08

A real Premion tracking-pixel defect under-recorded website attribution
for every campaign whose flight overlapped **June 9 – July 21, 2026**
(inclusive), fixed July 22 — confirmed via the Pansophic Learning (WUSA)
projection memo Matt supplied the same day. This CORRECTS an earlier
record (DECISIONS.md's own MW-2.15x-factor entry) that had called MW's
under-count a one-off; MW's real flight (June 1 – July 13, 2026) sits
inside the window, and Pansophic's own campaign (launched July 2) does
too, which is what surfaced the pattern as systemic rather than
MW-specific.

**Deliberately just a warning, never a correction** — Matt's own explicit
ruling ("don't build projection machinery; that decision stands") is
respected exactly: `app._overlaps_pixel_issue_window(start, end)` is a
pure date-range check against the two hardcoded constants
`PIXEL_ISSUE_WINDOW_START`/`_END`, and `render_attribution_reports_page`
fires one non-blocking, rep-facing sentence when an uploaded attribution
export's own flight overlaps it — no scaling, no projection, nothing
computed from the overlap beyond the fact of it. Guard:
`tests/test_attribution_reports_page.py`'s `check_pixel_issue_window`
(pure edge cases against the function, plus MW's real flight triggering
it and WAEPA's real flight — starting after the fix — not triggering it).
Full incident record, including the correction to the original MW entry,
is in `DECISIONS.md`.

### WAEPA testing follow-up (2026-09-08) — landed

The third real fixture (`Premion Website Attribution and Reach Extension
(13).xlsx`, no companion delivery file), found by testing live rather than
designed for in advance — exercises three shapes MW/Cardinal never touch:
a real split-IO multi-RFPID campaign, real conversions, and a real DMA
market cut (Washington DC / Baltimore).

**1. Multi-RFPID rejection was wrong for this case — replaced with a
confirm gate, never a hard refusal.** WAEPA is ONE campaign issued as TWO
RFPIDs (265618: 549,296 impressions; 263966: 60,464 — same audiences,
creatives, overlapping dates), and the old rule (built against a real
GLS/Twin Pine lifetime rollup, 180+ RFPIDs and years of history) rejected
it outright with `AttributionParseError`. `attribution_import.py` no
longer raises for this at all: `AttributionExport.rfpid_breakdown` carries
every RFPID's own delivered/attributed/conversion figures (one entry even
for an ordinary single-RFPID file, so a caller never special-cases
count==1), and `.rfpid` becomes a `" + "`-joined display string for
multiple — confirmed nothing downstream parses it structurally (unlike the
avails-PDF importer's own RFPID, which IS a dedup key elsewhere). There is
genuinely no per-RFPID date data anywhere in this export shape to judge
overlap from, so `app.render_rfpid_confirm_gate` defaults its confirm
checkbox from RFPID COUNT alone: 2-3 defaults checked (the split-IO
shape), more defaults unchecked with a stronger warning (the rollup
shape) — the rep can always override either way, informed by each RFPID's
own real impressions/attributed figures shown in a table. Unconfirmed
blocks Draft/Generate (the rest of the page doesn't render) but is never a
raised exception the rep can't get past. Guards: `tests/test_attribution_
import.py`'s `check_waepa_attribution`/`check_multi_rfpid_confirmed_not_
rejected`/`check_single_rfpid_breakdown_always_populated`, `tests/
test_attribution_reports_page.py`'s `check_rfpid_confirm_gate`.

**2. Conversions — a real, per-export OPTIONAL layer, verified against a
false-positive trap Matt named up front.** MW and Cardinal both carry the
per-dimension "Conversion Impressions"/"Conversion Impressions Rate"
columns already (they predate WAEPA) — always zero, columns present, no
data. `has_conversions` is true only when the top-line "Attributed
Conversions"/"Sales Amount" widget tab EXISTS and its value is `> 0` —
confirmed live that MW/Cardinal don't even have that tab at all (not
present-and-zero, genuinely absent), so existence alone would have been a
sufficient signal for these three files, but value-checked anyway per the
stated rule (a hypothetical export with the tab and a genuine zero must
still read as no-conversions). WAEPA: 37 attributed conversions, $0 sales
amount — a real count-without-value case, never divided into a rate.
`report_assembly.build_facts_payload`/`build_report_deck` both gained an
explicit `include_conversions` bool (never inferred from `has_conversions`
internally — always the caller's call, so a rep's own toggle can turn it
off even when the export has real conversions). Every downstream function
(`intent_facts`, `intent_summary_rows`, `top_url_rows`, `top_zip_rows`)
takes the SAME explicit opt-in param, never auto-detecting — WAEPA's own
"no half-states" rule, enforced structurally rather than by convention.

**Where conversions land, and where they deliberately don't:**
- Highlights slide gains a FOURTH tile, `{{HEADLINE_CONVERSIONS}}` (count,
  with the sales amount appended only when `> 0` — WAEPA's own $0 must
  never render as a confident "37 ($0)"). Reflowed away via the same
  `_reflow_tile_row` mechanism `FlightTile` already uses when absent/off.
  **Template change landed same-day (2026-09-08), built via python-pptx
  rather than by hand in PowerPoint** — `migrate_report_master_v0_4.py`
  (committed, one-time migration script) renames the three existing
  highlights tiles to the established `Tile`/`TileValue`/`TileLabel`
  convention this module's other tile rows already use (recap's
  `ReportPeriodTile`/`FlightTile`/`GeographyTile`, delivery recap's five)
  — `ImpressionsTile`/`ImpressionsTileValue`/`ImpressionsTileLabel`,
  `VisitorsTile`/`VisitorsTileValue`/`VisitorsTileLabel`, `RateTile`/
  `RateTileValue`/`RateTileLabel` — then deep-copies `RateTile`'s three
  shapes into a fourth group, `ConversionsTile`/`ConversionsTileValue`/
  `ConversionsTileLabel`, token `{{HEADLINE_CONVERSIONS}}`, label
  "Attributed conversions" (run-level text replace, so the source
  formatting — font, size, color — survives byte-identical), and
  repositions all four evenly across the ORIGINAL three-tile row span
  (same left/right edge, same measured ~274320 EMU gap, each shape's own
  value/label padding preserved). Shapes are found by TEXT CONTENT, not by
  the current default pptxgenjs name, so re-running it is safe if the
  names ever drift again. Verified: a real MW deck (no conversions)
  reflows cleanly back to 3 tiles at the same positions v0_3 had; a real
  WAEPA deck (conversions on) shows all 4 evenly spaced with matching
  formatting; both rendered via PowerPoint COM and eyeballed. **Uploaded
  to Supabase as `report_deck_versions` id 3 (`REPORT_MASTER_v0_4.pptx`),
  active.** Guards: `tests/test_report_assembly.py`'s tile-name assertions
  in `check_mw_attribution_only`/`check_waepa_conversions_and_multi_rfpid`
  (ConversionsTile present with conversions on, absent — reflowed — with
  them off, on two independent real files).
- Breakdown table gains a fifth column (`conv_rate`), URL-report's
  `IntentSummaryTable`/`TopUrlTable` each gain a fourth (`converted`) —
  all three via the SAME graceful-degrade path `_fill_named_table` already
  has (a field list longer than the template's real column count warns
  and drops the extra field rather than crashing) — confirmed live against
  the CURRENT (un-widened) template: three warnings fire, the deck still
  builds. No template action needed to ship this; it activates on its own
  whenever those three tables are widened by one column each, with a
  static header cell ("Conv. Rate" / "Converted") added by hand.
- **Zip table does NOT gain a sixth column** (Matt's own ruling, agreed:
  the table is already five columns wide) — `top_zip_rows`'s own
  `include_conversions` flag adds a `conversions` key to the facts payload
  only, for the narrative to cite; the deck's own zip table is untouched
  either way.
- `intent_facts()`'s conversions (full precision, one key per class) and a
  new top-level `"conversions"` fact (attributed count, sales amount when
  `> 0`, rate of attributed impressions) make conversions "a first-class
  fact the model can select," per the instruction — Phase 4's facts-only
  contract and its checker (`app._attr_payload_numbers`) needed no changes
  to cover this: conversions are just more numbers in the same payload.

Guards: `tests/test_attribution_import.py`'s `check_conversions_widget_
detection`/`check_no_conversions_on_mw_and_cardinal`, `tests/
test_report_assembly.py`'s conversions-plumbing checks (built into the
existing `check_phase4_override_plumbing` family), `tests/
test_attribution_reports_page.py`'s `check_conversions_toggle`/`check_mw_
has_no_conversions_toggle`.

**3. "Clear / new report" — landed, same deny-list-by-prefix idea as
`clear_proposal_state`.** `app.clear_attribution_report_state` sweeps every
`attr_`-prefixed session_state key — which is already, by construction,
this whole page's entire state boundary (the same prefix `NON_
PERSISTABLE_PREFIXES` already excludes wholesale on page navigation) — so
there's no allow-list of exceptions to maintain the way the Build page's
`SESSION_KEEP_ON_RESET` deny-list needs one; `standalone_groups` and
proposal-side state survive automatically, by construction, since none of
it is `attr_`-prefixed. Same inline (not `st.dialog`) confirm shape as
"New proposal," for the same reason (a dialog doesn't survive `AppTest`'s
fragment-rerun model). Guard: `tests/test_attribution_reports_page.py`'s
`check_clear_button`.

**4. The Highlights slide's own narrative quality — a genuine content bug,
not a facts-only violation.** A live MW run put 3 of 4 highlight bullets'
headline numbers on figures the slide's own tiles ALREADY show (3,104,554
impressions, 4,638 unique visitors, 1.36% attributed rate) — technically
facts-only-compliant (every number traced to the payload) but wasted
slide space: a client reads the same three facts twice. Compared against
the takeaways slide, which was already right (667 store-visit visits
against the foot-traffic goal, 63.2% product consideration, a real zip's
3.14x-baseline multiple) — findings, not restatements. Four rules added to
`build_attr_draft_prompt`'s `highlight_bullets` guidance: (1) a highlight
may not restate a tile value — the model is told the tiles' own values
explicitly, computed from the same facts payload; (2) at most ONE delivery
bullet, and only when genuinely notable; (3) prefer findings that COMPARE
(a zip's `multiple` against baseline, one market/audience/creative against
another in the same `rows` list — the payload already carries exactly
this); (4) when goals exist, at least one highlight must connect to the
goal-relevant intent class by name, not only the takeaways. Verified live,
first try, on both real datasets: MW's four highlights became the
store-visit finding (goal-connected), a zip's 3.14x-baseline outperformer,
one market beating another by their own real rates, and exactly one
delivery bullet (97.5% VCR) — zero tile restatements. Guards: `tests/
test_attribution_draft_live.py`'s `check_highlights_no_tile_restatement`/
`check_highlight_cites_fact`, run against both MW and Cardinal.

### Day of Week correction, response profile facts, OTT retargeting parser — 2026-09-10, parse+facts landed; slide fill blocked on v0_6

**The Day of Week correction.** Phase 1's own finding that a 1-day-gap date
series is "a trailing daily window" was wrong for the "ATTRIBUTED RATE BY
DAY OF WEEK" tab specifically — its 7 rows sum to the flight's own
delivered-impressions total exactly (WAEPA(14): 2,220,083), which a set of
single calendar days from a ~30K/day flight cannot do. The dates in that
tab are LABELS from one reference week, not real dates — `weekday()` on the
label date is the real key. `attribution_import._classify_date_series` now
takes the row sum and the export's own expected total and returns
`"day_of_week"` (not `"daily"`) whenever a 1-day-gap series' rows sum to
that total; `_day_of_week_rows` re-sorts Mon-Sun by `date.weekday()`.
Verified this generalizes, not a WAEPA-only fix: NO fixture on hand (WAEPA
13/14, MW, Cardinal) has ever actually produced a genuine `"daily"` classification
through this path — every 1-day-gap tab any of them carry is this same
weekday aggregate. Full incident writeup, including the "sum check would
have caught this in Phase 1" lesson, is in `DECISIONS.md`. Consequence:
the optimization engine's weekday dimension (see "Open items" below) is
unblocked — `AttributionExport.by_day_of_week` is real data now, not a
guess.

**Response profile facts** (`report_assembly.recency_facts`/
`referral_facts`/`day_of_week_facts`/`response_profile_facts`) — pure,
tested, feeding `build_facts_payload`'s new `facts["response_profile"]`
key, no template dependency:
- `recency_facts`: `AttributionExport.by_recency` (already-parsed
  "NN - NN DAYS" buckets) turned into per-bucket share plus one derived
  figure, `share_within_0_3_days` (the "00 - 03 DAYS" bucket's own share).
- `referral_facts`: `by_referral_domain` turned into per-source share plus
  `direct_share`.
- `day_of_week_facts`: gated on a named-constant threshold,
  `_DAY_OF_WEEK_UNEVEN_RATIO = 1.5` (best-day rate / worst-day rate) with a
  `_DAY_OF_WEEK_MIN_IMPRESSIONS_FLOOR = 1000` floor per day so a
  low-volume day's noisy rate can't trip "uneven" on its own — `None`
  unless at least 2 days clear the floor. Real numbers: Cardinal's week
  IS uneven (best/worst clears 1.5x); MW's is NOT — its own strongest day
  is Wednesday, the same day MW's media plan famously excludes (the
  "removed Wednesdays" cut this file already documents elsewhere), so a
  flat-looking week there is exactly what a rep already knew to expect,
  not a parsing miss.
- `build_attr_draft_prompt` (app.py) gained four new interpretation rules
  in its "Rules for each field" list, verbatim from Matt's own framing:
  0-3-day share reads as immediate response, spread beyond it reads as a
  longer consideration cycle (never assume immediate is always the better
  story); direct visits are the strongest single signal and worth naming
  by name; the other referral sources (organic/social/external) frame as
  CTV intersecting the client's other digital channels and lifting the
  funnel downstream, never as a same-fact stated as a shortfall; day-of-week
  is only worth naming when `"uneven"` is true, and a vertical-pattern claim
  (earlier-week fits home services/medical/insurance, later-week fits
  retail/travel) is only made when both the spread clears the threshold AND
  the vertical is actually known from goals/notes.

**OTT Retargeting parser** (`attribution_import.parse_ott_retargeting_export`,
new, a separate export from attribution/delivery — detected by its own
`CAMPAIGN KPIS` tab header, never a filename) — `OTTRetargetingExport`
carries campaign-level impressions/clicks/CTR, `has_actions`/`actions`
(present-and-nonzero only, same convention as conversions elsewhere),
`by_ad_size` (matched against a known five-size set, `_AD_SIZE_LABELS`
supplying the plain-English name — "320x50" -> "mobile banner" etc. —
never a bare regex that could misparse an unrelated filename fragment),
`by_creative`/`creative_groups` (grouped by base filename with the ad-size
suffix and extension stripped, so one creative CONCEPT sold in five sizes
is one group; `ott_retargeting_facts` only emits `creative_groups` in the
payload when `len(...) > 1`, matching Cardinal's real shape — one concept,
five sizes, no creative-comparison story to tell — against a hypothetical
multi-offer campaign like MW's, where it does), `by_screen` (Mobile/
Tablets/Desktop), and blended CTV+display reach (`blended_impressions`/
`blended_uniques`/`blended_frequency`). `ott_retargeting_facts` folds all
of this into `facts["ott_retargeting"]`, `None` when no OTT retargeting
export was uploaded at all. Guards: `tests/test_attribution_import.py`'s
`check_ott_retargeting_parser` (real Cardinal file, plus the
no-CAMPAIGN-KPIS-tab error-message contract), `tests/
test_report_assembly.py`'s `check_response_profile_facts`/
`check_ott_retargeting_facts` (real WAEPA/MW/Cardinal numbers for the
former; real Cardinal plus a synthetic multi-creative case using MW's own
cited real numbers, to prove creative grouping works before a real
multi-offer OTT-retargeting export exists on hand). 111/111 in both
files.

**Resolved by Matt (2026-09-10):**
- **Reach Extension: no slide, not parsed into the facts payload.** "We
  rarely if ever report it." The inventory above stays in this doc as a
  record of what the tabs hold, in case it's revisited later, but nothing
  reads them.
- **Blended frequency: the number is right, the question was period.**
  Confirmed (see `ott_retargeting_facts`'s own docstring in
  report_assembly.py for the exact evidence): the real Cardinal export's
  "PREMION + OTT RETARGETING" tab is campaign-to-date, not scoped to this
  report's own reporting period — every OTHER tab in that file sums to
  exactly 300,207 impressions for August alone, but blended impressions is
  950,329, and the file's own Pacing Report tab shows the underlying
  campaign spans March 2026 - February 2027, wide enough to account for
  the gap as accumulation since campaign start. Shipped per Matt's own
  decision tree: `ott_retargeting_facts`'s "blended" key carries impressions
  and uniques only, never frequency; `_ott_blended_stat` (report_assembly.py)
  states both figures and says "cumulative since campaign start" in the
  same sentence — never a bare number with no timeframe.
- **Thumbnails: shipped v1 without them.** `creative_previews` stays parsed
  (the image URLs are there) so a later fetch-at-import step has something
  to fetch; nothing fetches them yet.
- **One line added to the day-of-week prompt guidance** (`build_attr_draft_
  prompt`, app.py): when a day's own `delivered_impressions` sits far below
  the rest of the week, that's the media plan's own choice to limit
  delivery that day (MW's real Wednesday cut, the same cut this file's own
  correction section names), not a response pattern — with a linked
  proposal, say so plainly rather than presenting the day's rate as newly
  discovered.

**v0_6 landed (2026-09-10) — the fill code is built, the two new slides are
live end to end.** `report_assembly._fill_response_profile`/`_fill_ott_
retargeting` fill the two slides Matt's v0_6 template added (token names
exactly as he specified — `OTT_IMPRESSIONS`/`OTT_CLICKS`/`OTT_CTR` were his
own naming call where the original spec named shapes but not tokens).
`report:response_profile` is now a REQUIRED slide (added to `build_report_
deck`'s required-key check, same as recap/highlights/etc.) — every template
older than v0_6 (v0_4 and earlier) now fails to build against this code, by
design, the same way adding `report:live_sports` as required would have.
`report:ott_retargeting` is optional and driven by a NEW `ott` parameter on
`build_report_deck`/`build_facts_payload`, independent of the delivery set
per Matt's own instruction ("NOT delivery_set") — a 4th upload slot,
"OTT Retargeting export (optional)", was added to the intake row in
`render_attribution_reports_page`, parsed by its own `CAMPAIGN KPIS` tab
header the same way the parser detects it. The CreativeTable/AdSizeTable
shift-when-one-creative-concept works by reading the freed height from the
template itself (`AdSizeHeader.top - CreativeHeader.top`), never assumed —
same principle as `_fill_url_report`'s own intent-table overflow shift, run
in the opposite direction. Two new draft schema fields (`response_profile_
narrative`, `ott_retargeting_narrative`) and their own prompt rules were
added to `build_attr_draft_prompt`/`apply_attr_draft`. **v0_6 is live in
Supabase as report deck version 4, active** — confirmed directly via
`db.list_report_deck_versions()`: id 1 = v0_2 (7 slides), id 2 = v0_3,
id 3 = v0_4, id 4 = v0_6, active. **v0_5 has no row at all** — it was
never run through `setup_supabase.py report_decks` (only 3 versions had
ever actually been uploaded before this session), so it only ever existed
as a local file; version 4 is the correct next id, not a miscount. Also
confirmed directly: `db.report_master_deck(...)` serves version_id 4, and
the served template has exactly 11 slides in the order Matt specified
(recap, highlights, delivery_recap, delivery_breakdown, live_sports,
attribution_breakdown, response_profile, url_report, zip_analysis,
ott_retargeting, takeaways).

**A real, pre-existing bug found and fixed alongside this: `report:live_
sports` was left completely unfilled — every `{{SPORTS_...}}` token still
literal — on any report built with no delivery file uploaded.**
`build_report_deck`'s `drop_keys` computation took a hardcoded `if delivery
is None: drop_keys = [...]` branch that never included `report:live_sports`
(the `else` branch's `live_sports_applies(delivery)` check, which normally
adds it to `drop_keys` when a delivery file lacks a sports block, never ran
at all when there was no delivery file to check in the first place). Found
rendering a real no-delivery WAEPA report against v0_6 (the first template
old enough to even carry `report:live_sports` that this repo's own
zero-delivery fixture, WAEPA, has ever been built against) — a screenshot
showed the slide sitting there with raw `{{SPORTS_IMPRESSIONS}}` etc. text
still on it. **First patched by adding `"report:live_sports"` to that same
hardcoded list; Matt's own follow-up caught that this left the actual
shape of the bug in place** ("any list of slide keys that exists in two
places is stale the moment the first one grows") — the real fix collapses
BOTH branches into one `delivery_set_applies` dict (slide key → its own
applicability function, called uniformly whether `delivery` is None or
not, since `delivery_breakdown_applies`/`live_sports_applies` already
treat `None` as "doesn't apply" on their own), so a future delivery-set
slide needs one new dict entry, never two lists kept in sync by memory.
Full incident, both fixes, in `DECISIONS.md`. Guard: `tests/test_report_
assembly.py`'s `check_response_profile_fill`, which asserts `report:live_
sports` is gone (not merely unfilled) on every no-delivery build.

Guards for all of the above: `tests/test_report_assembly.py`'s `check_
response_profile_fill`/`check_ott_retargeting_fill` (full deck builds
against v0_6, all 3 real fixtures — day-of-week table presence, OTT slide
presence/absence, the creative-table shift measured against the raw
template's own shape positions, the blended-stat wording). 137/137 in that
file, 111/111 in `tests/test_attribution_import.py`, full green in `tests/
test_attribution_reports_page.py` (its own hardcoded template-existence
gates were also bumped from v0_4 to v0_6, where they'd gone stale --
harmless in practice since the app's own `db.report_master_deck()` was
already serving v0_6 regardless of what that gate checked, but worth fixing
so the gate means what it says). Screenshots of both new slides across all
three fixtures were sent to Matt for review before this was called done,
per the standing "look at the pictures" rule.

**Not done in this round, left for whenever it's asked for:** the
CreativeTable's row labels are the raw base filename (Cardinal's own
`MW_PremionDisplay_...` style, stripped of size suffix/extension) rather
than a rep-editable label — the original item-2 spec's "Creative LABELS
EDITABLE at Generate, same as delivery geo" turned out to describe a
pattern ("delivery geo" labels editable at Generate) that doesn't actually
exist anywhere in this codebase either, so there was nothing to mirror;
building a genuinely new editable-label UI wasn't part of what this round
asked for and would need its own scoping pass first.

## Highlights/Takeaways rework — landed 2026-09-10

Replaces the old three-list model output (`highlight_bullets`/
`takeaway_bullets`/`whats_next_bullets`, each drafted independently) with
one `threads` array the model returns and Python distributes
deterministically — `report_assembly.distribute_threads` — onto both
slides plus what's-next. Each thread carries `head`/`anchor` ("goal" or
"signal")/`goal_ref`/`finding`/`meaning`/`action`; a thread's `finding`
becomes a highlight (skipped when null — the closing-only-signal case,
e.g. a direct-visit share worth a takeaway but not a tile-adjacent
highlight), `meaning` a takeaway (every thread needs one), `action` a
what's-next item (no cap, unlike the two slides, which stay capped at 4
with signal threads dropped before any goal thread when trimming). Goal
threads are listed by the model in ITS OWN inferred priority order —
that ordering IS the priority; there's no separate schema field for it,
and no bespoke reorder widget was built, because a rep who disagrees with
the model's stated order (surfaced in `goal_alignment_notes`) already has
one: re-order the lines in the existing goals textarea and re-draft.
Guard: `tests/test_thread_distribution.py` (offline, hand-built thread
payloads — cap/trim precedence, closing-only signals, goal ordering).

**Delivery metrics (VCR, CTV share, frequency, publisher/creative mix) now
appear in a thread ONLY when a stated goal names them** — the old "at
most one delivery bullet, when genuinely notable" allowance is retired.
Verified live against MW/Cardinal (no goal names a delivery metric — none
appears in any thread) and WAEPA (a synthetic frequency goal — 3.95
appears). The dedicated Delivery Recap/Breakdown slides' own narrative
fields are untouched by this rule.

**Six new pieces of context, all optional, all wired through `report_
assembly.build_facts_payload`:**

1. **Benchmarks** (`attribution_benchmarks.py`, new file) — the real
   `PremionWebsiteAttribution-Benchmarks 2025 (1).xlsx` (found in
   Downloads, not the project folder; stays gitignored like every other
   real-data workbook here) transcribed ONCE into 8 committed vertical
   rows (not the 6 first assumed — Healthcare and Entertainment/Gambling
   were missed on a first read). `VERTICAL_TO_BENCHMARK` maps `app.
   VERTICALS`' 8 mappable verticals to a row (confirmed with Matt: Retail
   wraps Furniture Retail, Entertainment wraps Entertainment/Gambling;
   Travel and Casual Dining/QSR have no row and are never approximated to
   a neighbour). `benchmark_facts(vertical, attributed_rate,
   attributed_unique_visitor_rate)` returns None below-benchmark on both
   rates (silence, never "room to improve"); above, cites only the
   rate(s) actually cleared, phrased as a comparison ("above the Premion
   benchmark for legal campaigns") never the benchmark row's own raw
   percentage — enforced by excluding the whole `"benchmark"` subtree
   from `app._attr_payload_numbers`' traced-number set, so a model that
   DID quote the raw row percentage would be caught as a fabrication.
   Verified live: WAEPA/MW/Cardinal (all three real fixtures) correctly
   cite nothing — all three sit well below every candidate row — and a
   synthetic above-benchmark case cites the comparison exactly once
   (grouped by thread, not raw string count — a benchmark thread
   legitimately contributes one highlight AND one takeaway) without ever
   quoting the row's own percentage. Guards: `tests/test_attribution_
   benchmarks.py`, `tests/test_attribution_draft_live.py`'s
   `synthetic_benchmark` scenario.

   **The numerator/denominator settled in writing** (a real find: the
   benchmark's "Average Attributed (Impression) Rate" is confirmed
   identical to `AttributionExport.attributed_rate` — attributed
   impressions over delivered impressions — and the visitor-rate column
   is the SAME denominator with attributed unique visitors as the
   numerator, not a different one). **WAEPA now means file "(14)"**
   (single RFPID-258341, Mar-Jun 2026 flight, 0.13% attributed rate / 532
   unique visitors over 2,220,083 delivered) in `tests/test_attribution_
   draft_live.py` — this is the file a previously-generated real report
   actually used. File "(13)" (two RFPIDs, Jul-Aug 2026) is a genuinely
   DIFFERENT real WAEPA campaign, and stays the fixture `tests/test_
   report_assembly.py`'s/`tests/test_attribution_import.py`'s own
   multi-RFPID + conversions(37) regression checks are legitimately built
   around — those were correctly left untouched, not switched to "(14)".

2. **The account's own trend** (`report_assembly.report_headline_facts` —
   period dates from the export's own flight, attributed rate/unique
   visitors/conversions, top intent class — the same compact summary
   Phase 6/Polk will build on) is now logged alongside every report
   (`report_json["headline_facts"]`, additive next to the existing raw
   attribution/delivery dicts) and read back for the SAME advertiser's
   later reports as `facts["prior_periods"]` (oldest-first; a row logged
   before this shipped simply has no `headline_facts` key and is skipped
   — a known, un-backfilled limitation). The model may build a trend
   thread from it — a goal thread if a goal mentions lift, a signal
   thread otherwise. Verified live: a real Cardinal export plus a
   hand-built earlier period at half the attributed rate produces a
   trend thread.

3. **Plan vs. actual** — `report_assembly.plan_vs_actual_facts` is a
   best-effort join (normalized-label fold, substring either direction)
   between a linked proposal's own plan rows (`app.linked_proposal_
   report_fields`'s new `plan_rows` key) and a delivery export's
   `by_geo`. A "Show planned vs delivered" toggle (default off) is gated
   on THREE things: a proposal linked, a delivery export uploaded, AND
   the template's `DeliveryByGeoTable` actually having 5 columns already
   — `report_assembly.named_table_column_count` checks this at render
   time (cached by template version id), so a rep never sees a toggle
   that changes nothing on the slide. Today's v0_6 template has 3
   columns, so the toggle doesn't render yet — **v0_7 needs
   `DeliveryByGeoTable` widened to Geography | Planned | Delivered | % of
   plan | VCR, plus an optional `{{PLAN_VS_ACTUAL_NOTE}}` text run on
   report:delivery_recap** (the fallback line for unmatched markets or
   when the breakdown slide doesn't exist at all) — sent to Matt as its
   own template-change spec, same shape as every other template change
   this feature has needed. The deck-fill side (`_fill_delivery_recap`/
   `_fill_delivery_breakdown`'s new `plan_vs_actual` parameter) already
   degrades gracefully via `_fill_named_table`'s own column-count check,
   so nothing further needs to change in code once the template widens.
   **The in-app shortfall alert is independent of the toggle and ships
   now** — whenever a proposal is linked and a delivery export exists,
   every market where delivered < planned gets an `st.warning` naming the
   market and the gap, regardless of whether the toggle is even offered.
   Plan-vs-actual is never a thread on its own unless a stated goal asks
   about pacing — verified live against WAEPA (no such goal — silent).
   Guards: `tests/test_report_assembly.py`'s `check_plan_vs_actual_facts`/
   `check_named_table_column_count`, `tests/test_attribution_reports_
   page.py`'s `check_vertical_conversion_definition_and_plan_toggle`.

4. **Vertical** — a selectbox (`app.VERTICALS`), pre-filled from the
   linked proposal's own row-level `vertical` column (outside
   `form_json`, passed through `linked_proposal_report_fields`'s new
   `vertical` parameter) or, absent that, the confirmed advertiser's own
   stored vertical (the full advertiser row, not just its id, is now kept
   in `session_state["attr_advertiser_row"]` once confirmed, so this
   never needs a second fetch) — always editable, "unknown" (never
   guessed) otherwise. A changed value writes back to the advertiser
   record on Generate (`db.set_advertiser_vertical`, best-effort, needs
   the Stage 16 DDL below) so a LATER standalone report for the same
   advertiser inherits it without asking again.

5. **Conversion definition** — no existing structured field holds this
   anywhere in a proposal's `campaign_specs` (checked directly); a new
   rep-typed text input (`attr_conversion_definition_input`, session-
   scoped like notes, no persistence) feeds `facts["conversion_
   definition"]`. Present, the model names conversions using that exact
   phrase everywhere one is mentioned; absent, generic wording plus a
   `goal_alignment_notes` ask.

6. **Goal priority** — see the threads section above; no separate field,
   no new UI.

**Stage 16 DDL** (`supabase_schema.sql`, additive/idempotent, pasted by
hand per this project's standing rule): `alter table public.advertisers
add column if not exists vertical text;` — needed before vertical
persistence (item 4) actually round-trips; degrades to a harmless error
string via `db.set_advertiser_vertical` until it's run.

**Two live-drafting incidents during verification, both resolved on
re-run (single-run model slips, matching this project's own documented
"a small percentage summed into new arithmetic" failure mode), not a
prompt defect:** a WAEPA draft once wrote "37%" (20% + 17% from two
separate cited page shares, summed) and a fixture-generation draft once
wrote "21%" instead of the real 22.11% store-visit share — both caught by
the existing facts-only checker exactly as designed, and both vanished on
a second live call. Also two real TEST bugs found and fixed while writing
Tier 2's own new checks, neither a model or prompt issue: a `rep.check`
call missing its detail argument, and a broken `"plan" not in haystack or
"planned" not in haystack` boolean (an OR that's true almost regardless of
content) — replaced with an explicit plan-vs-actual phrase list.

Guards, full list: `tests/test_thread_distribution.py`, `tests/test_
attribution_benchmarks.py` (both new), `tests/test_report_assembly.py`'s
`check_report_headline_facts`/`check_plan_vs_actual_facts`/`check_named_
table_column_count`/`check_facts_payload_rework_wiring`, `tests/test_
attribution_draft.py` (rewritten for `threads`, fixture regenerated live
and re-frozen), `tests/test_attribution_reports_page.py`'s `check_
vertical_conversion_definition_and_plan_toggle`, `tests/test_attribution_
draft_live.py` (rewritten: MW/Cardinal now assert no delivery-metric
thread and no benchmark citation; a new `cardinal_trend` scenario; WAEPA
rewritten against real proposal 09e61e0b-a945-4022-851d-d3c31b2acbd0's own
stored goals plus a synthetic frequency goal and delivery object; a new
`synthetic_benchmark` scenario) — 47 passed, 0 failed on a clean full run.

## Open items Matt is chasing

- Auto-Sales Analyst's `build_group_chart_images()` — repo link coming
  before Phase 4.
- Whether the Premion dashboard can pull a full daily time series — not
  blocking before the optimization phase.
- ~~**v0_7**: widen `DeliveryByGeoTable`...~~ **Done, 2026-09-11** —
  `REPORT_MASTER_v0_7.pptx` is live in Supabase (`report_deck_versions` id
  5, active): `DeliveryByGeoTable` is 5 columns (Geography | Planned |
  Delivered | % of plan | VCR) and `report:delivery_recap` carries
  `PlanVsActualNote`. The plan-vs-actual toggle and shortfall alert both
  work against it now.
- ~~**Stage 16 DDL** needs pasting...~~ **Done** — confirmed live directly
  (`db.fetch_advertisers()` returns a populated `vertical` column; WAEPA
  reads `"banking"`), not just inferred from the migration file.
- ~~**v0_8, requested 2026-09-12**...~~ **Done, 2026-09-12** (`8f63b8d`) —
  `REPORT_MASTER_v0_8.pptx` uploaded as `report_deck_versions` id 6 and
  activated (superseded by v0_9, id 7, since); `BreakdownTable` +Conv.
  rate, `IntentSummaryTable`/`TopUrlTable` +Converted, verified against
  real WAEPA/MW renders with conversions on and off.
- **Stage 18 DDL** (`advertisers.optimization_level`, from the optimization
  sequence work above) — also confirmed live directly the same way (the
  column comes back on every `fetch_advertisers()` row).

## NWFCU review (2026-09-17) — landed, verified against WAEPA/MW/Cardinal; real NWFCU file not available on this machine for a true end-to-end re-run

Matt walked a real NWFCU report through the app and found two client-facing
errors, a gate-counting bug, and named five rounds of enrichment against
Andrea's deck. All six items landed in one round; three real bugs were
found and fixed verifying the fixes themselves (named below, each with its
own guard). **One open item:** the real NWFCU attribution export that
produced these findings isn't present in this project's fixture set (only
a same-named budget/flighting spreadsheet was found, not the Premion
Website Attribution export) — every fix below is verified against its own
literal spec and against the three real fixtures already in this repo
(WAEPA, MW, Cardinal), never against NWFCU's own file directly. Re-run
`tests/test_report_assembly.py`/`tests/run_all.py` against the real NWFCU
export the moment it's available, to close that gap.

### 1. Page/URL percentages — REACH, not share of visits

`intent_summary_rows`/`intent_facts`/`top_url_rows` (report_assembly.py)
all divided a class/page's own visits by the SUM of every class's visits
— which double-counts a visitor who reached more than one page, so the
percentage answered "what share of page VISITS was this" rather than the
question a client actually asks ("of the people you sent me, how many got
to X"). All three now divide by `attribution.attributed_unique_visitors`
instead — the campaign's own fixed unique-visitor total. Real find,
confirmed by the math alone (not the missing NWFCU file): NWFCU's own
reported 36%/3.7% roughly doubles to the stated 72%/7.4% under this
correction, consistent with visitors there averaging close to 2 pages
each. Percentages no longer sum to 100% — expected, not an error — and
the slide needs a permanent footnote saying so (see the template handoff
list below). The old visits-share metric survives as `visit_share`/
`_visit_share_raw`, distinctly named, for a mix sentence only ("most
attributed traffic also touched...") — the drafting prompt states this
explicitly and forbids citing both numbers for the same class in one
sentence. The deterministic URL narrative fallback (`_url_intent_
narrative`) moved from "X% of attributed visits landed on" to "X% of
attributed visitors reached," matching the new semantics. Guard:
`tests/test_report_assembly.py` (187 checks, green); no dedicated new
percentage-value assertion added since the real fixtures' own numbers
aren't independently known to be "correct" without NWFCU's file to check
against — flagged as the one open verification gap above.

### 2. Never recommend what's sold — "in plan, not yet live"

NWFCU's narrative called a bought Live Sports package "not yet activated
or tracked" and then recommended it in What's Next as a new idea — the
model had no fact telling it the product was already sold and scheduled.
Fixed with a new `facts["not_yet_live"]` list (`{"product", "starts"}`
per entry):
- **With a linked proposal:** `report_assembly.not_yet_live_facts_from_
  plan_rows(plan_rows, period_end)` parses each plan row's own Flight
  shorthand text for the earliest named calendar month and flags any row
  whose flight starts after the report's own period. `linked_proposal_
  report_fields`'s `plan_rows` now carries `tactic`/`flight` alongside the
  existing `geo`/`planned`.
- **Without a proposal:** a new rep field, "Tactics in plan but not yet
  live (one per line: product, start month)" (`app.parse_not_yet_live_
  lines`), sitting right after "What's next" on the New report tab. Hard
  gate at Generate: a Wrap-up report with no linked proposal and this
  field left blank is refused with an explicit error naming what to type
  ("None" is a fine answer) — a monthly report never requires it.
- **The rule, in the prompt:** a product named in `not_yet_live` is phrased
  "begins in <month>," never "not activated"/"not tracked"; it may NEVER
  appear in What's Next or a thread's action as something to add.
- **Recap slide:** a new `flight_progress_label` ("Months 1-3 of 6")
  appends to the existing FLIGHT_LABEL token whenever the campaign's own
  full flight bounds are known (a linked proposal's `flight.start`/`.end`)
  — no new template shape needed, since it rides the existing token.

**Real bug found regenerating the Harrisburg rep guide, same day**:
`flight_progress_label`'s own docstring already claimed it returns None
when the report's period is entirely outside the campaign flight, but the
implementation only ever clamped the month-index math into range —
verified against the real WAEPA fixture pair, whose linked proposal
flight (Oct-Dec 2026) and attribution export period (Mar-Jun 2026) don't
overlap at all, producing a confidently wrong "Month 1 of 3." Fixed with
an explicit no-overlap check on the raw dates, distinct from the
legitimate "overruns by a few days" case the clamp remains for.

Guard: covered by the existing `test_report_assembly.py`/`test_
attribution_reports_page.py` suites (both green, 196 checks total after
`check_nwfcu_review`'s addition); no NWFCU-specific end-to-end check yet,
pending the real file.

### 3. Optimization gate counted the wrong thing

The gate used `len(prior_periods) + 1` — PRIOR LOGGED REPORTS only — so a
multi-month wrap recap (real monthly data inside ONE export, zero prior
reports logged) read as "first report, no optimizations" while the panel
itself was already offering real candidates built from that same data.
`report_assembly.evidence_periods(attribution, prior_periods)` replaces
it: the distinct calendar months in THIS export's own `monthly_trend`,
unioned with every prior report's own logged period span, deduplicated.
`optimization_candidates`'s own gate and its `timing_note` both call this
one function now, so they can never disagree. ZIP_MIN_SHARE (the same 1%
floor `top_zip_rows` already uses) now also gates ZIP optimization
candidacy — a thin zip (0.68% share) no longer qualifies as a cut
candidate just because its rate is bad; real volume is required too, same
as the zip table's own outperformer rule.

**Two real bugs found verifying this fix, both the same shape** — "this
report" (or a prior report) silently contributing ZERO periods instead of
its guaranteed minimum of one:
1. A flight whose own date-tab classifies as WEEKLY rather than monthly
   (MW's real 6-week flight — `monthly_trend` is genuinely empty) counted
   zero months from this export, reopening the exact bug this fix closes,
   from the other direction. Caught by `test_attribution_reports_page.py`
   going red (a real optimization checklist rendered empty against real
   MW data). Fixed: falls back to the one month `attribution.flight_
   start` falls in.
2. A synthetic test export with no flight dates AT ALL, and a prior period
   using placeholder non-date strings ("x"/"y") that the OLD code never
   parsed but the new one does — both would have contributed zero.
   Caught by `test_optimization_engine.py` going red (14 failing checks,
   all "expected N candidates, got 0"). Fixed: both "this report" and any
   unparseable prior period fall back to a private sentinel key,
   guaranteeing at least one period each, never zero.

Guard: `tests/test_optimization_engine.py` (all green), `tests/test_
report_assembly.py` (187 green), `tests/test_attribution_reports_page.py`
(all green, including the real-MW optimization-checklist scenario).

### 4. The URL page, richer

- **Vertical-aware sub-classes**: `VERTICAL_URL_INTENT_PATTERNS["banking"]`
  splits "consider" into certificates/auto financing/cards/mortgage/
  membership, and adds a new `existing_member` class (login, online
  banking, loan payments) — checked before the base taxonomy when a
  vertical is known. `classify_url_intent`/`intent_summary_rows`/`intent_
  facts` all take an optional `vertical` parameter now, threaded from
  `build_facts_payload`/`build_report_deck` down.
- **Existing-member vs. prospect**: the drafting prompt now states
  explicitly that an `existing_member` class's visitors are never folded
  into "potential new members"/"prospects reached" language.
- **Goal-named pages surface regardless of reach**: `top_url_rows` gained
  `goal_keywords` (from the new `extract_goal_keywords(goals, notes)` —
  a plain word-substring match, deliberately not fuzzy/ML) — a page whose
  bucket label names something the goals talk about is force-included
  past the normal top-N cap, marked `goal_match: True`.
- **Careers never scores as a lead**: a real bug class — a careers URL
  whose path ALSO contains a generic word like "apply" ("/careers/apply-
  now") was matching the base taxonomy's own "purchase" pattern via that
  later segment. `_CAREERS_SEGMENTS` now short-circuits to "other"
  whenever any path segment names a job-listing page, checked before any
  other pattern.

Guard: `tests/test_report_assembly.py` (187 green) covers the mechanism;
no real banking-vertical fixture on hand to verify the sub-class split
against actual traffic (the same open NWFCU-file gap as items 1-2).

### 5. Within-flight trend, device, channel, place names

- **Within-flight monthly trend**: `report_assembly.within_flight_trend_
  facts(attribution, flagged_window)` exposes a wrap/full-recap export's
  own multi-month `monthly_trend` as a fact (`{"month", "attributed_rate",
  "known_tracking_issue"}` per entry) even with ZERO prior reports — "no
  prior reports" was never "no history." `flagged_window` (app.py's own
  `PIXEL_ISSUE_WINDOW_START/_END`) marks any month overlapping the known
  pixel defect, so the model names the tracking issue as the reason a
  month's own figure is unreliable rather than treating it as a real dip.
- **Weekly line chart, added as a same-day follow-up request**: `report_
  charts.render_line_chart` (Pillow-only, no new dependency, same
  convention as `render_bar_chart`) plots `weekly_trend_facts`'s own
  `attribution.weekly_trend` series. `_fill_weekly_trend_chart` is
  deck-wide and OPTIONAL — it scans every slide for a shape literally
  named `TrendChartRegion` (+ `TrendChartRegionLabel`, same two-shape
  convention as every other `ChartRegion`/`ChartRegionLabel` pair in this
  module) and is a safe no-op everywhere until a template has one. **Open
  template handoff** (see the list below): no template has this shape
  yet; `report:response_profile`, alongside the existing recency chart,
  is one reasonable home, but the mechanism doesn't require that specific
  slide. `weekly_trend` also rides into the facts payload for an optional,
  single-swing thread — the chart is the citation, so the prompt tells
  the model not to restate the whole series in prose.
- **Device sidebar**: already correctly unconditional — `device_split_
  facts` fills the one-slide summary's second sidebar block by default
  whenever the export has a device breakdown, no threshold, no toggle.
  Confirmed by reading the existing code; no change needed here, only
  verification.
- **Publisher table gains Attributed Rate**: `_fill_delivery_recap` now
  takes `attribution` and joins `TopPublishersTable`'s existing name/
  impressions/VCR rows against `attribution.by_channel`'s own per-channel
  attributed rate, adding a "rate" column via `_fill_named_table`'s
  existing graceful column-count degrade (pending a template widen, same
  shape as every other such column added this year). A publisher name
  with no `by_channel` match shows "--", never a guess.
- **Zip Area resolves to PLACE first**: `report_assembly.zip_place_name`
  (new) finds the nearest real Census Gazetteer place (`map_places.
  json.gz`, already built for map labeling) to a zip's own centroid
  (`geo_resolver`'s `zip_points`), via haversine distance, capped at 15
  miles. `top_zip_rows` tries this FIRST now; DMA/market (the old primary
  path) is the fallback, then the existing zip3 fallbacks, in the same
  order as before. Verified directly: 20171 -> "Herndon" (3.05 miles,
  confirmed as the genuinely nearest of 19,478 real places, not just the
  first candidate checked).

Guard: `tests/test_report_assembly.py` (187 green, including an updated
20001 -> "Washington" assertion reflecting the new place-first behavior).

### 6. Smaller

- **`report:attribution_breakdown` drops when nothing to break down**:
  `pick_breakdown_dimension` always returned SOME dimension, falling back
  to Market even with exactly one market — a real bug, since a single-
  market/audience/creative campaign rendered a one-row "breakdown" table
  that broke nothing down. `attribution_breakdown_applies(attribution)`
  gates the slide's existence now (deleted outright when every dimension
  tops out at one row), same shape as `delivery_breakdown_applies`'s own
  rule for its slide. `facts["breakdown"]["applies"]` tells the model when
  there's no comparison to draft a narrative or a dimension pick for.
- **Retargeting creative variants group to one concept**: `attribution_
  import._creative_base_name` already stripped ad-size suffixes; a new
  `_CREATIVE_VARIANT_RE` also strips trailing `_ALT`/`_ALT2`/`(1)`/`(2)`-
  style markers, so "ROS"/"ROS_ALT"/"ROS (1)" all reduce to one "ROS"
  concept in `_group_ott_creatives`. Verified directly against all six
  named examples plus two negative cases (a genuinely different name
  containing "ALT" mid-string, and an ad-size suffix) — all correct.
- **Blended reach says "households"**: `_ott_blended_stat`'s deterministic
  sentence and the drafting prompt's own guidance both changed from
  "unique visitors" to "households reached" for a blended CTV+display
  figure — a household-level TV reach count, not tracked individual
  visitors, borrowing the wrong vocabulary.

Guard: `tests/test_report_assembly.py` (187 green, including the blended-
reach wording assertion already in the suite).

### KEEP (explicitly unchanged, per Matt's own instruction)

Benchmark stays silent below its own threshold (4.21% < Banking's 4.81%
row is correctly NOT cited); day-of-week stays off below its own "uneven"
threshold (Fri 4.5% vs. Sun 4.0% isn't a story); the threads structure
itself is untouched. No code touched any of these three.

### Template handoffs — all three DONE, landed as v0_10 (2026-09-18, `563bf19`)

All three of the items below landed in `REPORT_MASTER_v0_10.pptx`
(`report_deck_versions` id 8, active), verified against the exact spec by
`build_v0_10.py` and live-rendered both with and without a weekly series:
1. ~~IntentSummaryTable/TopUrlTable...~~ **Done** — headers read "% of
   visitors," `UrlReachFootnote` carries the "don't sum to 100%" wording.
2. ~~TopPublishersTable...~~ **Done** — "Attributed Rate" is the fourth
   column.
3. ~~A `TrendChartRegion`...~~ **Done** — lands on `report:highlights`,
   with a `TrendChartHeader` too; `_widen_highlights_card` reflows the
   WHAT STOOD OUT card back to v0_9's full width whenever there's nothing
   to plot (see the cross-month evidence section below for how this chart
   became series-aware on top of this template work).

### Verification status

`tests/test_report_assembly.py` (187), `tests/test_attribution_import.py`
(111), `tests/test_attribution_draft.py` (30), `tests/test_attribution_
reports_page.py`, and `tests/test_optimization_engine.py` all green
individually before the sweep. Full `tests/run_all.py` chunked gate (8
chunks, 89 test files) run chunk-by-chunk end to end: **89 passed, 0
failed, 0 skipped.**

## Cross-month evidence — landed 2026-09-18, extends the accept/decline loop (`cad0900`)

**THE PRINCIPLE this round is built around:** the engine must never
optimize on one month when more exist. Consistency across periods is the
evidence, not magnitude in the latest one. Before this, "evidence" meant
*how many periods exist at all* (`evidence_periods`, still used for the
timing gate); it never asked whether a value that looked bad this month
also looked bad last month. A single wild month could out-rank a smaller
but real, persistent gap.

**1. Dimension-level `period_facts`, stored per report.**
`report_assembly.period_facts_for_report(attribution)` — additive jsonb,
`report_json["period_facts"]`, computed and stored at every Generate
alongside `headline_facts`. One list per dimension (`PERIOD_FACTS_
DIMENSIONS = ("zip", "creative", "market", "day_of_week", "publisher",
"audience")` — "publisher," not "channel," to speak the optimization
engine's own dimension names verbatim), each row `{label, delivered,
attributed, rate, share}`; plus `campaign_rate`/`campaign_delivered`/
`campaign_attributed`/`campaign_visitors`/`period_start`/`period_end` and
`intent_facts()`'s own reach-basis classes. **Backfilled for the two
already-logged reports whose real export files are still on this
machine** (`backfill_period_facts.py`, WAEPA and Mattress Warehouse) —
Marines Service Co and NWFCU's own logged reports could NOT be backfilled,
their source files aren't present locally (see the NWFCU note below).

**2. Report series — `series_id` (Stage 19 DDL, `supabase_schema.sql`).**
`report_assembly.resolve_open_series(prior_reports, new_period_start)` —
a report auto-joins the advertiser's own most recent series whose latest
period ends 0-1 calendar months before this one starts; no match starts a
fresh series (a new `uuid.uuid4()`, app-generated, never `gen_random_
uuid()`). `db.set_report_series`/`fetch_series_reports` are the manual
override — a "Edit series link" popover on every Report history row (and
on the Clients page's own Reports list, same function) shows what a
report is linked with and lets a rep unlink it (NULL — excluded from
series evidence entirely) or join it to another report's series
explicitly. **NULL means unlinked, never "no opinion yet."**

**3. The engine reads the series — `_consistency_verdict`.**
`optimization_candidates(..., series_period_facts=...)` now classifies
every material-below-baseline value across every period it appears in
(this period plus every PRIOR linked series period): a value over-average
in even one period is a WASH and is dropped outright, regardless of the
majority; otherwise a STRICT MAJORITY of its periods must be under for it
to be a real `candidates`/`watch_list` entry. Everything material this
period that hasn't cleared majority consistency yet — including a value
with real history that's only under in a minority of periods — lands in a
new `forming` list: visible, named, explicitly **not** a recommendation.
Ranking within a dimension is consistency-first (fraction of periods
under, then how many periods back it), magnitude only breaks a tie. A
value with no series history at all reduces to a single-observation,
automatically consistent — exactly the old single-period behavior,
unchanged for every standalone report.

**4. Level sets the timing gate — `_LEVEL_GATE_MIN_PERIODS`.** Replaces
the old flat "fewer than 2 periods, nothing for anyone": High opens from
month 1 (at a tightened 30% swing floor and a 2-candidate cap — one month
alone is thin evidence, so High pays for the early start with a higher
bar), Moderate needs 2, Low needs 3. Below its own level's gate, the
account still gets the FULL consistency analysis — nothing is
suppressed — every item that would have been a recommendation instead
lands in `forming`, with `timing_note` naming the gap.

**5. Month-over-month reference — the data half is live, the deck-side
caption needs a template round (below).** `report_assembly.momentum_
deltas(attribution, series_period_facts)` — this period vs. the single
most recent prior period in the series: `attributed_rate_now/_then/
_delta_points`, `attributed_unique_visitors_now/_then/_delta_percent`.
Rides into `facts["momentum"]` (a thread may already cite it in the
narrative, ahead of any slide caption existing). **`TrendChartRegion` is
now series-aware** — `report_assembly.series_trend_facts` prefers one
point per PERIOD in the linked series (a real multi-month trajectory)
over `weekly_trend_facts`'s within-this-export weekly points whenever the
series has 2+ periods; a standalone report or the first period of a fresh
series draws exactly the weekly chart it always did.

**6. The wrap reads the whole series.** `facts["series"]` (`series_
summary_facts`) — one row per prior linked period, letting a wrap's own
takeaways span the whole flight rather than only ever comparing this
month to last. `facts["series_reconciliation"]` (`series_reconciliation_
facts`) — a wrap's own fresh, whole-flight export SHOULD roughly sum to
what was already reported period-by-period; a ≥10% mismatch on delivered
or attributed impressions is flagged (`material: true`) rather than
silently absorbed — a real signal of a late month, a re-pull with
different dates, or a genuine data problem. `optimization_history`
already read each prior report's own full `attribution` dict (not just
`headline_facts`), so no change was needed there for dimension-level
detail.

**Verified:** `tests/test_optimization_engine.py` gained a dedicated
`check_cross_month_consistency` (wash exclusion, majority-vs-minority,
consistency-beats-magnitude ranking, single-observation backward
compatibility) plus a rewritten `check_timing_gate` for the new
per-level minimums; both green against real MW/WAEPA/Cardinal data
alongside every existing check. `tests/test_attribution_reports_page.py`'s
`FakeStore` gained `series_id`/`set_report_series`/`fetch_series_reports`/
`update_attribution_report_json` stubs. Full `tests/run_all.py` chunked
gate (8 chunks, 89 test files): **89 passed, 0 failed, 0 skipped.**

### Closed out (2026-09-18, `abe03df`) — Stage 19 pasted, NWFCU verified, v0_11 live

**Stage 19 DDL pasted and confirmed.** Drove a real Generate against the
real running app (Mattress Warehouse, test-mode) and confirmed `db.log_
attribution_report` writes cleanly — no "couldn't log it" caption, a real
`series_id` on the new row (a fresh uuid, since the prior MW report had
none to auto-join). Cleaned up the Test-Mode-tagged confirmation row
afterward.

**The real NWFCU export is now in the project folder** (`Premion Website
Attribution and Reach Extension (15).xlsx`, found 2026-09-18) — identified
by its own Advertiser tab ("Northwest Federal Credit Union - Direct"), not
by filename, since the same filename pattern is shared by every numbered
Reach Extension export in the repo (confirmed (13)/(14) are both WAEPA).
Backfilled its `period_facts`. **Closes the loop the whole 2026-09-17
review was written for:** Homepage reach 71.8% (rounds to the reported
72%), Lead/Contact reach 7.4% — both exact. The "in plan, begins
September" sports line isn't reproducible verbatim (this report has no
linked proposal, so the fact would have come from the no-proposal rep-
typed field, never persisted) but the RULE it exercises is verified
against the account's real dates: flight_end is 2026-08-31, genuinely
before September, so `not_yet_live`'s mandatory "begins in {month}"
phrasing (never "not activated") is the objectively correct call here,
not just a plausible one. `check_nwfcu_real_export` (`tests/test_report_
assembly.py`) guards both figures and the date logic permanently.
Marines Service Co's own real export (`Premion Website Attribution and
Reach Extension.xlsx`, no suffix) still isn't in the project folder.

**The guide was regenerated** (`build_user_guide.py --task all`) — new
build stamp, and a new Clients reference step (`ref_clients_02_series_
link`) showing a report row's own series line and its "Edit series link"
popover, inserted before the existing "Create case study" step (now
`ref_clients_03_case_study`).

**v0_11 is live** (`report_deck_versions` id 9, `build_v0_11.py` verifies
it against the exact handoff spec below) — `RateTileDelta`/
`VisitorsTileDelta` landed exactly as specified, geometry verified against
v0_10 with nothing else moved. Two real things fixed alongside the
rollout, not left as follow-ups:
1. **A real bug, caught live**: `_MOMENTUM_CAPTION_SHAPES` stored PRE-
   WRAPPED token strings (`"{{RATE_DELTA}}"`), but `_fill_tokens`/
   `assembly._replace_tokens_in_text_frame` wrap a BARE name themselves
   (`_placeholder`) — a pre-wrapped key can never match the run's own
   literal `"{{...}}"` text, so the token silently never filled. First
   full round-trip render showed the literal token text still sitting on
   the slide. Fixed to bare names (`"RATE_DELTA"`); `check_momentum_
   captions` guards the fix, the sign, and the delete-when-absent path.
2. **The rep-facing toggle the spec assumed already existed, didn't** —
   v0_11's own delete condition ("no prior linked period OR the MoM
   toggle is off") named a toggle that had never actually been built
   (item 5's original spec called for one, "toggle default ON," but the
   earlier session chose to skip it, reasoning it could just ride on
   series existence). Added "Show month-over-month comparison" (default
   ON, shown only when a series exists) and threaded `show_momentum`
   through both `build_facts_payload` call sites and `build_report_deck`
   — discovered along the way that neither `build_facts_payload` call
   site had ever been passed `series_period_facts` at all, so `facts
   ["momentum"]`/`["series"]`/`["series_reconciliation"]` had never once
   reached the model; fixed at the same time.
3. The caption wording itself changed from "vs last month" to naming the
   prior period's own real month ("vs May 2026") — "last month" stops
   meaning anything correct the moment someone reads the deck later than
   the month it was generated in; the real month name doesn't have that
   problem. Verified live against the real v0_11 template in both
   directions (a rate/visitor increase and a decrease), confirming
   correct sign, rounding, and month naming.

Stage 19's own SQL (`supabase_schema.sql`, Stage 19, right after Stage
18) — kept here for reference now that it's live:

```sql
alter table public.attribution_reports add column if not exists series_id uuid;

create index if not exists attribution_reports_series_id_idx
    on public.attribution_reports (series_id);
```
