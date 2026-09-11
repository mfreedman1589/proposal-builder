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

### Phase 4 — the model-facing half — **landed 2026-09-08**

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

### Phase 5 — Proposal link (with-proposal mode) — **landed 2026-09-09**

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

- **Optimization engine** (framework/addendum step 3) — the weekday
  dimension is UNBLOCKED (2026-09-10 correction): "Day of Week" is exactly
  what it says, a real weekday aggregate (`attribution_import.py` gotcha
  2's own corrected finding, `AttributionExport.by_day_of_week`), not a
  trailing window. Still blocked on whether the dashboard can pull a full
  daily time series for the rest of the engine's own needs — Matt is
  checking that separately.
- **Create case study from report** (addendum §3) — sequenced after Phase 5
  in the addendum itself, so it can use both the report and its proposal.
- **One-slide report summary** (2026-09-08 addition, same sequencing as
  the item above and for the same reason: better with the originating
  proposal's goals in hand, so it lands after Phase 5 too) — a toggle
  producing an abbreviated, single-slide version of the report: the
  headline tiles, two or three findings, the what's-next line. For a
  monthly check-in email, a client who won't open a deck, or an exec who
  wants the page, not the story. **Build it alongside "Create case study
  from report" — they're the same shape**: one slide, the report's
  strongest facts, a template that fits into someone else's deck. The
  difference between them is audience (the summary is for THIS client;
  the case study is for the NEXT one), and that's a PROMPT difference, not
  a layout difference — one template family, two fills, one build.
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
  Polk's own tiles render once that slide exists. **Dependency: Phase 6
  (Polk itself) has to land first** — there's no matched-outcome tile to
  attach the toggle to yet. Nothing further to build today; this paragraph
  IS the spec, ready to implement the moment Polk is scheduled.
- **Polk — file in hand, phase-6 fixture, inventory only (2026-09-08).**
  `Polk_Dashboard.xlsx`, 19 tabs. Headline figures: Matched Households,
  Target Dealer Sales, Buy Rate, Matched Impressions, Match Rate, Campaign
  Lift (1.5x in the sample on hand). Cuts available: sales by days-elapsed
  / gender / age / income; make-model per dealer with MSRP; audience /
  creative / publisher share of matched impressions AND share of target
  sales (two tabs each — a share-of-impressions view and a share-of-sales
  view, not the same thing); Target Dealers with market rank vs. campaign
  rank; All Dealers (133 of them) with campaign share. **Depends on the
  match-rate toggle above** (every matched-outcome tile on this slide
  needs it) — not scheduled, no phase number assigned yet beyond "after
  the toggle exists."
- **Sales match-back, brand lift, Arrivalist** (addendum §5-9) — still
  fully deferred, each gated on a real raw deliverable from Matt, per the
  addendum. Sales match-back shares the match-rate toggle's design above
  once its own deliverable arrives.

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
- **Stage 16 DDL** needs pasting into the Supabase SQL editor:
  `alter table public.advertisers add column if not exists vertical
  text;` (see `supabase_schema.sql`'s own Stage 16 block for the full
  comment). Unblocks vertical persisting onto the advertiser record.
- **v0_8, requested 2026-09-12 (the post-Generate warnings walkthrough):**
  widen three tables by exactly one column each, all three already
  filled and gracefully dropped by `_fill_named_table`'s own column-count
  check today (`report_assembly.py`'s `_fill_attribution_breakdown`/
  `_fill_url_report` — no code changes needed once these ship, same
  "activates on its own" shape as v0_7):
  - **`BreakdownTable`** (`report:attribution_breakdown`) — currently 4
    columns: Label | Delivered | Attributed | Rate. **Add a 5th column,
    "Conv. Rate", after Rate.** Filled from each row's own `conv_rate`
    (a percentage string, same shape as the existing Rate column).
  - **`IntentSummaryTable`** (`report:url_report`) — currently 3 columns:
    Label | Visits | Share. **Add a 4th column, "Converted", after
    Share.** Filled from each row's own `converted` (an integer count).
  - **`TopUrlTable`** (`report:url_report`, the same slide's second
    table) — currently 3 columns: Label | Visitors | Share. **Add a 4th
    column, "Converted", after Share.** Same shape as IntentSummaryTable's
    new column.
  All three columns are populated ONLY when a report has conversions data
  and the rep's "Include conversions" toggle is on (`include_conversions`,
  default on when the export has any) — a report with no conversions data
  renders these tables exactly as they do today, unaffected.
