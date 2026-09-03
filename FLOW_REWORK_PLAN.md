# Proposal Builder — Flow Rework

**Handoff to Claude Code. Read this alongside CLAUDE.md, DECISIONS.md and BACKLOG.md.**

## Why

The app currently asks the rep to make decisions in an order that doesn't match how those decisions actually cascade. Flighting, plan basis and campaign specs all sit below things that depend on them, so the app infers what it should have been told, and `call_claude_draft` carries reasoning load it shouldn't have (net vs gross, which rows an audience phrase refers to, whether a figure is monthly or full-flight).

The rework is one idea applied six times: **decisions that cascade move to the top, and the model's job shrinks to selection and percentages.** Python does the arithmetic. The rep declares the frame up front, and everything below inherits it.

The secondary goal is the BACKLOG UX sweep. A Harrisburg rep opening this for the first time sees the whole sheet at once. It should feel like being walked through a route, not handed a control panel.

### Global principles for every phase

1. **One computation feeds every surface.** Flight dates, daily rates and basis conversion each get exactly one owner. (See the DECISIONS lesson on Timing fallback / `proposal_title` / `default_geo` drift.)
2. **Order of operations must not matter.** Progressive disclosure guides the first pass; it never locks. Import-then-draft, draft-then-import and clarify-after-either still reach the same end state.
3. **The model never writes money.** In avail mode it selects rows and states a percentage. Impressions, CPM and cost are computed.
4. **Rep edits win.** Precedence is unchanged: rep edits > avails PDF > draft from notes > defaults.
5. **New `form_json` keys need migration.** Every phase that adds a field must load pre-rework proposals from Proposal History without exploding, and must pass the round-trip property test that already gates group migration.

---

## Phase 1 — Setup band

**Goal:** five decisions at the top of the page that everything below inherits.

### Changes

Add a persistent setup band above all existing sections:

| Control | Type | Default |
|---|---|---|
| Originating market | select (DC / Harrisburg) | DC |
| Flight start / end | date range | empty |
| Plan basis | radio (Monthly / Full flight) | Monthly |
| Working from an avails document | checkbox | **on** |
| Total TV | checkbox | **off** |

- Nothing below the band renders until **originating market and flight dates** are set. The three toggles have defaults, so they never block.
- Total TV **on** reveals the WO schedule upload inline in the band. Total TV **off** hides it entirely. The existing rule stands unchanged: Total TV only applies when Premion runs the broadcast schedule in our own market — the toggle must remain unavailable (not merely unchecked) when the originating market can't support it.
- The band stays visible and editable for the whole session. It is a frame, not a wizard step.
- Changing a band value after work exists downstream must **flag**, never silently rewrite. Shortening the flight after plan lines are built raises a confirm naming what will change.

### Invariants

- Originating market and target DMAs stay two distinct concepts. The band sets originating market only. Target DMAs remain where they are.
- `form_json` gains `setup.*` keys. Old proposals load with originating market inferred from existing provenance, flight dates from existing timing fields, basis defaulting to Monthly, avail mode **on** if the proposal has any targeting group carrying avails, Total TV from the existing toggle.

### Acceptance

- A fresh session shows only the band.
- Loading a pre-rework proposal from History reproduces the same generated deck as before the change (Tier 1, existing scenarios).
- Toggling Total TV off removes the WO upload and the co-brand cover, and the deck regenerates clean.

---

## Phase 2 — Flighting moves up, custom per-month ranges

**Goal:** flight dates become the single authoritative source of truth for every date in the app.

### Changes

- Start/end in the band derive the list of calendar months spanned.
- A **Custom flighting** checkbox expands the derived months into one row per month, each with its own date-range control. Each defaults to the full calendar month clipped by the flight bounds — so the default state of custom flighting is identical to the non-custom state, and ticking the box changes nothing until a rep edits a row.
- These per-month ranges feed:
  - **Avails** — the existing daily-rate reduction. An avails figure is scoped to the flight the avails *document* states; reduce to a per-day rate over that stated flight, reapply across the actual per-month days. This already exists; it now reads from one place.
  - **Plan lines** — exact dates, not month names.
  - **Campaign Specs** — the flight statement.

**Landed early, pulled forward from Phase 3 (2026-08-27):** a brand-new plan option's
Breakout now defaults to the setup band's own Plan basis (`default_breakout_for_basis`,
app.py) instead of a hardcoded Monthly — the core promise of the waterfall in two lines.
Only a genuinely new option reads the band; copying an existing option still keeps that
option's own breakout untouched. The per-section (avails vs. plan) Monthly/Full-flight
*override*, and letting a rep unlock/diverge a section from the band after the fact,
stay Phase 3 work — this piece is only the default. Verified against every
`avails_basis`/breakout-sensitive suite (group plan selection, order invariance,
breakout rescale, flight-change guard, avails reach, setup resolution, the group
scenario fixtures, drafting Tier 1, and the byte-identical backward-compat rebuild) —
all green before committing.

**Capital Media walkthrough is the standing Phase 2 proration fixture.** 12 weeks from
9/21 ends 12/14; the avails document runs to 12/20. Sep/Oct/Nov match the avails
document's own month boundaries; December is 14 plan days against 20 avails days —
a real, deliberate divergence between the plan flight and the avails flight, not a
data error. Use it to prove the per-day daily-rate reduction (the bullet above) actually
prorates December correctly rather than reading either flight's day count for the other.

### Plan-cell shorthand

A plan line spanning several months needs a compact Timing cell. Render rule:

- A month running its full length (as clipped by the flight) renders as the bare month: `Oct`
- A partial month renders with days: `Sep 8–30`
- Consecutive full months collapse: `Oct–Nov`
- Joined with commas: `Sep 8–30, Oct–Nov, Dec 1–14`

Single owner for this function — the plan slide, the on-screen grid and Campaign Specs all call it. Do not reimplement per surface.

### Invariants

- **The broadcast row is exempt.** WO dates are verbatim, exactly as the broadcast Geo column is never re-seeded. Custom flighting does not touch it.
- The existing warning when a WO schedule span doesn't cover the plan flight still fires, now comparing against the band's dates.
- Broadcast monthly allocation by each week's actual start date is unchanged.

### Acceptance

- Custom flighting ticked with no edits produces byte-identical output to unticked.
- A three-month flight with a partial first and last month produces correct per-line impressions from a single monthly avails figure.
- The WO fixture truths still hold (Regency 148 spots / $28,000 / 1,831,600; Ravens 12 programs / 21 weeks / $79,500 / 1,785,413).

---

## Phase 3 — Avail mode

**Landed.** Built across 10 commits (`git log --oneline` — "Flow rework Phase 3, commit
1" through "commit 10"). Two sections below are corrected from their original wording
by decisions made during live plan review, after this phase's design was drafted but
before it was built — **treat the corrections as what actually shipped**, not the
paragraphs they replace: "Collapsing rows" (no automatic collapse — a rep action
instead) and "Acceptance" (rewritten to match). See CLAUDE.md's Phase 3 bullet for the
full shipped shape and `DECISIONS.md`'s "Phase 3 — entity identity" section for why
each correction was made, including the earlier-audit override on the "not a new
column" line below (already marked superseded before this phase was built).

**Goal:** two well-defined paths instead of one path guessing which it's in.

### Mode A — avails document (default)

The plan is strictly downstream of the avails table.

- The model **selects rows** and **states a percentage** (`percent_of_avails` + `avails_ref`). It does not emit impressions, CPM or cost. This is already the mechanism — Phase 3 makes it the only mechanism in this mode.
- Drafting pre-selects as it does today (notes-named audiences → those rows, flagged `unresolved`; silent notes → all rows, flagged `unresolved_internal`) and never overrides a rep's checkbox.
- Groups own plan lines. Drafting contributes budget, allocation, products, flight and attribution to selected lines rather than inventing its own.

### Mode B — no avails document

Unchanged from today: the model drafts lines directly, money math still in Python, badged as AI-filled.

### The Label becomes the entity identifier

**This is the substantive change in Phase 3. Read it before touching anything else.**

An avails document exposes only **Audience Target** and **Geography Included** per row. Nothing in it records *which real-world thing* a row is for. That missing fact is what forces the model to reason instead of select, and it's the root of the crisscrossing this rework is trying to end.

The Label carries it. **Superseded, implementation-time correction:** the original design here called this "not a new column — the existing Label, given a defined job," reusing the D2 grid's existing Label column (the Geo-cell override, `group["name"]`) for entity identity too. An earlier audit already overturned this before Phase 3 was built — the case that settles it: Annapolis Cars has one entity (one make) with TWO avail rows, a 5mi and a 10mi radius tier around the same origin, needing two *different* Geo-cell overrides but ONE entity name. Reusing one field for both jobs doesn't survive that real case. What was actually built: the existing Label column is **renamed "Geo Label"**, unchanged in behavior (still `group["name"]`, still the Geo-cell override), and a genuinely new "Label" column carries the entity name (`entity_label`), joined by a separate `entity_id` — never the label text itself, so a rename can't orphan anything. See CLAUDE.md's Phase 3 entry and `targeting_groups.py`'s own module docstring for the shipped shape.

**Label = the entity the row is for.** "Toyota of Annapolis." "Undergraduate." **Many avail rows can share one Label.** Targeting keeps carrying the *how* — the audience expression and radius. A plan row then reads:

> **Toyota of Annapolis** | Subaru intenders, 10-mile radius

The Label does three jobs:

1. **Join key** — associates avail rows to plan rows.
2. **Sync point** — draft-from-notes writes it, so "split evenly across the four stores" lands on something structural rather than being translated into row-level percentages.
3. **Client-facing name, optionally** — a *Show Label in plan* toggle. A dealer group needs to see which store each budget line belongs to; a single-location advertiser doesn't.

#### Join on an id, not on the text

The link between avail rows and plan lines must run on a **stable internal id**. The Label text is a display value the rep edits freely. Renaming "Toyota" to "Toyota of Annapolis" must not orphan the avail rows or drop the SOV. Rep edits text; the link doesn't move.

#### Allocation is per-Label, not per-row

"$15K split evenly across the four stores" resolves as budget ÷ distinct Labels, then SOV derived within each store from that store's own avails. The model names the split. Python does every number.

#### Collapsing rows: never automatic — a rep action, and only then take the max

**Superseded, implementation-time correction.** The original design here said rows sharing a Label collapse into one plan line automatically, taking the largest avails figure. That was overturned during live plan review: **one ticked avail row is one plan line, always.** An entity spanning several rows — alternate radius tiers, two locations of the same brand — is a legitimate multi-line plan shape a rep may want to keep exactly as multiple lines, not something the app decides to merge. What shipped: entity grouping's job is display/allocation only (see "What was actually built," below); collapsing two rows into one stays the existing, rep-triggered `merge_plan_rows` action, unchanged by this phase.

**What max-not-sum is actually for.** When a rep *does* choose to merge two rows that share an entity, the merged line's avails takes the maximum of the two, not the sum — summing is usually wrong, and wrong in the client's favour, which is the worst direction. From the real documents:

| Case | Why summing lies |
|---|---|
| **Annapolis Cars** — 4 audiences × 2 radii each | Within an audience, the 5mi zip list is a strict subset of the 10mi list. Same households counted twice. |
| **Plaza Motors** — A35-64 and M35-64 | Identical zip list; the men are a subset of the adults. 1.71M + 1.13M = 2.83M avails that do not exist. |
| **Wilmington University** — 3 undergrad audiences | College Planning Parents, Prospective College Students and Higher Education Intender overlap partially and unknowably. |

Rows belonging to *different* entities still sum on merge, exactly as before this phase — max-not-sum is shape-driven (same entity vs. different entities), not a new universal rule.

**Per-entity allocation, without collapsing rows.** A stated split ("$15K evenly across the four stores") still divides by distinct entity count in Python, not by row count — but the result lands on exactly one row per entity (the first selected row, stable order); any other selected-but-unmerged row for that entity keeps whatever it already had. What was invisible about this — which allocation basis produced which figure, and which entities have an unpriced sibling row — is what the allocation-basis caption under the plan table exists to surface (see CLAUDE.md's Phase 3 bullet), not a new plan-grid column.

#### Single-entity proposals pay nothing

Most proposals are one entity. Blank Label is valid, the Show-in-plan toggle is off by default, and when every row shares one Label (or none) nothing downstream behaves differently. Lawn & Leisure has exactly one row and must stay a one-line plan with no new UI.

#### What the model still infers

Which rows belong to which Label, on first pass. That inference doesn't disappear — it **relocates into one visible, editable column** instead of hiding inside a selection decision. A rep sees it, fixes it in seconds, and can tell they've fixed it. Segment names are usually enough to infer from (`AUTO Make Subaru`, `LIFESTAGE Online Education`), and the Geography name sometimes carries clarifying text ("Plaza Motors Group L2T Campaign Zip List"). Where it can't tell, leave blank and flag — never guess an entity name.

### Basis override per section

The band sets the default. The avails section and the plan section each get a local Monthly/Full-flight override.

This is safe **because** Phase 2 made dates authoritative — the two views are two renderings of one underlying set of numbers, not two independent sets of dates. Conversion continues to run through `avails_to_display` / `avails_from_display` with only edited rows converting back.

The existing broadcast schedule and plan breakout toggles stay independent of all of this, as they are today.

### Invariants

- The app transcribes avails a rep states and **never fabricates them**.
- Hierarchy stays audience > geo, ordering audience-major, duplicate geos never collapse.
- Non-audience-specific products (retargeting, flat fees, broadcast) stay single lines and are never multiplied per group.
- Review-list routing rule is untouched: `unresolved` = client must answer, `unresolved_internal` = seller checks alone.

### Acceptance

**Corrected from the original wording per the no-auto-collapse decision above** — rows
never merge on their own, so row counts below reflect one line per ticked avail row,
not one per entity.

- In Mode A, no code path writes a money value that didn't come from Python.
- Import-then-draft and draft-then-import reach the same end state (existing test, must still pass).
- **Annapolis Cars** (RFPID-253813): 8 avail rows ticked → 8 plan lines by default (4 distinct entities among them, each entity's 5mi/10mi pair NOT auto-merged); a rep merging one entity's pair afterward maxes correctly. Ticking only the 4 rows on the 5mi radius → 4 lines, each pricing off its own 5mi figure (never leaking in the untouched 10mi sibling). An even four-way split of a stated budget lands correctly, by entity, without the model emitting a single dollar figure.
- **Plaza Motors** (RFPID-266583): 2 rows ticked → 2 plan lines by default (both recognized as one entity by deterministic inference); a rep merging them produces 1 line at 1,707,337 avails — not the summed 2,834,169.
- **Wilmington University** (RFPID-253956): drafting's `group_entities` ties the 3 undergraduate rows to one entity → still 5 plan lines by default (the allocation-basis caption surfaces that the undergrad entity's full share landed on its first row, the other two flagged); a rep merging those 3 rows afterward produces 3 lines total.
- **Lawn & Leisure** (RFPID-265521): single row, single line, no Label UI change visible, order byte-identical to before this phase.
- Renaming an entity's Label after the plan is built changes the displayed text and nothing else — ids, rows, avails, SOV and the caption's own claims about unrelated entities all survive untouched.

**Found after this phase shipped, fixed 2026-09-01: a shape none of the acceptance fixtures above cover.** Every real-document fixture here (Annapolis, Plaza, Wilmington, Lawn & Leisure) is a *single* plan — the acceptance criteria never exercised two *options* whose `group_selection` deliberately differ (one option selling one of several same-audience rows, another selling all of them). A real draft in that shape (LiveWell Animal Hospital, 7 locations sharing one audience) came back with every option carrying every globally-selected row regardless of that option's own selection — six phantom rows at a real, visible $0 on the option that should have had one. Root cause: `apply_draft_to_form` resolves each option's `matched_ids` correctly, but flips a group's `include_in_plan` **globally** the moment *any* option matches it; `reconcile_group_plan_lines`'s row-seeding step then reads only that global flag on the next rerun, blind to which option actually wanted the row. Fixed by having that step consult each option's own stored `matched_ids` (already computed, already stored in `draft_plan_intent`, previously read only for re-pricing) before seeding — a group excluded from an option only when a *sibling* option's intent claimed it and this one's didn't; a group no intent ever touched still reaches every option, unchanged. Full incident and the two rejected alternatives (auto-dropping $0 rows; giving `include_in_plan` real per-option scope) are in DECISIONS.md's "Drafting, the catalog and the Claude API" section. Guard: `tests/test_group_plan_selection.py`'s multi-option scenario, proven to catch the bug by reverting the fix and watching it fail with the exact reported shape before restoring it.

---

## Phase 4 — Campaign Specs relocation + agency markup

**Superseded ordering, landed as two commits (4a, 4b) under a new name — read
before touching this section.** A later design session (recorded in the
approved plan at `~/.claude/plans/idempotent-leaping-reddy.md`, and summarized
in CLAUDE.md's own "Flow rework Phase 4" status bullet) reshuffled what
belongs in this phase versus Phase 5:

- **Campaign Specs relocation did NOT land here** at the time this note was
  written. The claim below is WRONG, left for the record rather than
  deleted — an interface audit (2026-09-04) checked it against the actual
  `apply_geography_autofill` call site and found no such dependency; the
  move landed that day, matching this section's own original text almost
  exactly. See the "Suggested commit sequence" list's item 4, and
  DECISIONS.md's "A redirected item can go stale silently" for why this is
  now the worked example of a general pattern, not just a corrected note.
  ~~The rep's own flow ordering ("avails come before drafting, and before
  the market picker") means Campaign Specs' Geography autofill now depends
  on markets the AVAIL resolves, not on Section A's picker directly — a
  bigger reshuffle than this phase's original text describes. It now
  belongs to **Phase 5**, alongside moving Section A below the intake
  uploads.~~
- **Per-month flighting moving into the setup band did NOT belong to Phase 2
  or 5 — it landed HERE, as Phase 4a**, alongside a "Match avails flighting"
  button and the divergence notice's move into the band. Not in this
  document's original text at all; see CLAUDE.md's Phase 4 bullet for the
  shipped shape.
- **Agency markup landed here, as Phase 4b, matching this section's own text
  below almost exactly** — with one refinement: the "flagged suggestion in
  unresolved_internal" idea in the "How drafting handles gross/net phrasing"
  subsection below was for the DRAFTING prompt and was dropped entirely (the
  model transcribes with no flag, full stop, order-wide grossing has no
  per-line exception to flag). A flagged note DOES fire, but only from the
  avails-PDF importer's own Agency field — a document fact, not a model
  inference — never auto-ticking the checkbox.

The text below is left as originally written, for the parts that still
describe what shipped (agency markup) and as a record of what was reordered
away (Campaign Specs).

### Campaign Specs

Move the section up, directly below the setup band and above avails. It's mostly frame-level information (geography, flight, product set) and reads as an afterthought at the bottom.

Geography behaviour is unchanged: defaults to target markets, falls back to the originating label.

### Agency markup — replaces the top-level toggle

**Remove** the agency toggle from the intake area. `call_claude_draft` stops classifying net vs gross entirely.

**Add** a plain checkbox adjacent to the plan table, near Generate: *"Apply agency gross-up (×1.15)."* When checked it multiplies **CPM and cost only**. Never impressions — a grossed-up line carries a higher CPM and cost against the same impressions the client receives. Broadcast lines are excluded: broadcast is always gross, WO cost verbatim.

It is a calculator. No model involvement, no inference, no state beyond the checkbox.

### How drafting handles gross/net phrasing in notes

This is the open question from yesterday. **Recommendation: the model transcribes and does not classify.**

Whatever figure the notes state goes into the table as stated, and the toggle stays off. "Budget is $15K gross, $38 gross CPM" → those numbers land verbatim, already gross, toggle off. "Net CPM is $32, gross it up" → $32 lands, and the rep ticks the box.

The alternative — having the model pre-tick the toggle when notes say "gross it up" — reintroduces exactly the classification that caused the confusion, for one saved click. If it goes in at all it should be a flagged suggestion in `unresolved_internal`, never an automatic tick. **Default to the plain version.**

### Acceptance

- Ticking the box changes CPM and cost by exactly 1.15 and leaves impressions untouched.
- Broadcast lines are unaffected.
- No prompt in the repo mentions net vs gross classification.

---

## Phase 5 — Progressive disclosure

**Superseded — a different, narrower Phase 5 landed under this name instead;
read before touching this section.** A later design session (2026-08-30,
recorded in the approved plan at `~/.claude/plans/idempotent-leaping-reddy.md`,
summarized in CLAUDE.md's own "Flow rework Phase 5" status bullet) replaced
this section's whole design with something smaller and more targeted:
market/flight/plan-basis becoming genuine INPUTS to drafting rather than
facts the model had to infer from prose (and sometimes got wrong) — a
drafting-accuracy change, not a layout one. What shipped:

- **The band's three intake pieces (notes/draft, avails PDF, Wide Orbit)
  became toggles that reveal their own input directly beneath themselves**
  when ticked, nothing rendered when not — but this is NOT the two-row
  top/second-row split this section originally proposed, and there is no
  "Customize" expander and no reveal-in-sequence. `apply_draft_to_form` no
  longer writes `market_choice`/`flight_start`/`flight_end`/`active_months`/
  `avails_basis` at all (the schema fields for market/flight_start/flight_end
  were dropped too, so the model is never even asked); it reads them from
  the band instead. A notes date/market that disagrees with the band never
  overwrites it — the band wins, always, flagged to `unresolved_internal`
  naming both.
- **The gate is UNCHANGED** — still just originating market + flight dates,
  still a hard `return`, exactly Phase 1's own rule. The "satisfied
  declarations" gate idea floated during that design session (blocking
  Generate on an avails/WO upload actually completing) was explicitly
  withdrawn: ticking a declaration doesn't mean a file is coming — hand-typed
  avails with no PDF, and rate-card-only broadcast with no Wide Orbit
  schedule, are both real, first-class paths a hard gate would have locked
  reps out of. Two **non-blocking** warnings at Generate time replace it
  instead (avails table empty despite the toggle; Total TV on with no
  schedule imported).
- **A same-shape live bug was found and fixed alongside it**: drafting was
  unconditionally force-writing `avails_basis` to Monthly whenever it built
  its own avails rows, silently overwriting whatever plan basis the rep had
  picked. Deleted, not patched — same fix shape as the flight/market removal.
- **The drafted-months reconciliation incident (Phase 2, "The media plan and
  the form" above / DECISIONS.md) is now structurally impossible**, not
  merely guarded against — there is no longer a second, drafted flight for
  the form's own to disagree with, since drafting never proposes one.
- A ledger bug in `target_dma_choice` was fixed alongside the rest: an avails
  import's autofilled markets (monotone add-only via
  `_group_markets_applied`) used to get silently dropped by a draft naming a
  *different* market, because `apply_draft_to_form`'s own write was a
  wholesale replace. Now a union — a market the avails resolved and the
  notes simply didn't mention survives a draft naming an unrelated one.

**This section's own content below — the top-row/second-row band split,
per-section reveal, "Customize" expanders — was NOT built and stays queued.**
It's a real, still-open idea (the IKEA-route goal is still worth having), but
it's now a future phase in its own right, not something the 2026-08-30
session's "Phase 5" delivered. Read the text below as that still-open design,
not as a record of what shipped.

**Goal:** the IKEA route. Signposted path, but you can cut across.

### Decided 2026-08-27, captured now, built here: the band splits into two rows

The Phase 1 band (as shipped) does two jobs with different rules, and conflating them
is exactly the "control panel, not a route" problem this phase exists to fix. Split it:

- **Top row — what the rep HAS.** Draft from notes (default **on**), avails document,
  Total TV. Pure declarations, no typing at this row. Toggling one **on** reveals its
  input box directly below it — notes box, avails uploader, WO uploader; toggled **off**,
  the box isn't rendered at all, not just disabled.
- **Second row — what the campaign IS.** Originating market, flight dates, plan basis.
  Filled by drafting when notes are on; typed by hand in self-serve (draft off means no
  notes box, and the rep types these values directly — nothing below silently waits on
  a notes parse that isn't going to happen).

Sequence becomes: declare what you have → provide it → drafting fills the second row's
values → the gate (originating market + flight dates, per Phase 1's own rule) opens.

**Why:** keeps the one rule that actually matters — never make a rep hand-enter something
the notes already state, e.g. "9/21 start, 12 weeks" — while making the top of the page
genuinely declarative instead of a mixed bag of checkboxes and fill-in fields. It also
hands drafting real context for free: market, flight, basis and avail-presence are all
already known before `call_claude_draft` reads a word of the notes, which is a strictly
better position than today's (notes drafted blind, then reconciled against whatever the
band already held).

**Not built yet — this is a design note for when Phase 5 lands**, so the Phase 1 band's
current single-row layout (market / flight / basis / avails-toggle / Total-TV-toggle, all
in one row) stays as-is until then. Whoever picks up Phase 5 should re-read this before
touching the band.

### Changes

- **Split the setup band into the top-row/second-row shape above.**
- Sections reveal in sequence as their prerequisites are satisfied.
- **Once a section is complete it stays open and editable.** Revealing is one-way. Going back never unwinds anything.
- Everything advanced collapses behind a **Customize** expander per section. The default path is short.
- Absorb the queued BACKLOG UX sweep here rather than as a separate pass:
  - intake organised by what the rep brought
  - draft-from-notes made prominent (currently easy to miss)
  - uploads consolidated instead of scattered
  - notes uploadable, not only pasted
  - geo expanders collapsed by default after an import

### Invariants

- No hard gating beyond the Phase 1 band. A rep who knows the tool can fill things in any order.
- Nothing about disclosure changes what gets generated. Same inputs, same deck.

### Acceptance

- A first-time rep can produce a single-option CTV proposal without opening a single Customize expander.
- AppTest coverage for the reveal sequence, including revisiting a completed section and editing it.
- Watch the widget-key lifecycle here — keys must bump on row-count change, not only on other triggers. Disclosure changes what renders per run, and Streamlit garbage-collects `session_state` for keyed widgets not rendered in a run.

---

## Phase 6 — Slide Vault

Do this **after** the flow rework lands, so it doesn't tangle with sections that are moving.

**Goal:** a shared library any user can add slides to and pull from.

### Changes

- Supabase table + bucket alongside the existing `case_studies` pattern.
- Per slide: **name**, **tags**, and a **deck position** — a named anchor, not a page number. No hardcoded slide numbers anywhere, per the existing convention.
- Slides carry the hidden `key:` line in speaker notes like everything else, so `slide_map` resolves them the same way.
- Insertion needs a defined rule relative to existing ordering. Case studies insert after attribution, immediately before the first plan slide; vault slides need their own anchor set, spelled out in SLIDE_KEYS.md before building.
- Adding to the vault must preserve speaker notes — that means the Reuse Slides / copy-paste path, not the app's programmatic copy.
- Section dividers stay banned in generated proposals regardless of what a vault slide contains.

### Open before building

- Who can add — everyone, or a review step before a slide becomes available team-wide?
- Whether image-path rendering (like case studies) is acceptable, trading editability for cross-deck OOXML safety.

---

## Suggested commit sequence

1. Setup band + `form_json` migration + History round-trip test — **landed**
2. Flighting relocation, custom ranges, shorthand renderer (single owner) — **landed**
3. Avail mode split; Label-as-entity (id-based join, per-entity allocation, rep-triggered max-not-sum merge); per-section basis override — **landed** (plus a multi-option row-scoping gap found and fixed after landing, 2026-09-01 — see this phase's own section above)
4. Agency markup (as Phase 4b) — **landed**, matching this document's own text almost exactly. Campaign Specs move — **landed 2026-09-04**, matching this document's own original text (`### Campaign Specs`, above) almost exactly: directly below the setup band, above avails, Geography behaviour unchanged. The "did not land, redirected to Phase 5" note above (and Phase 5's own "Superseded" note) turned out to rest on a claim about the code that was never re-verified — an interface audit checked it against the actual `apply_geography_autofill` call site and found no such dependency; see DECISIONS.md's "A redirected item can go stale silently" for the general pattern this is now an example of.
5. The setup band becoming drafting's own input (market/flight/plan-basis, 2026-08-30) — **landed**, under this phase's name but a narrower scope than originally designed. The progressive-disclosure redesign this section actually describes (top-row/second-row band split, per-section reveal, Customize expanders) plus the BACKLOG UX sweep — **not built**, still queued (a separate, smaller UX sweep landed earlier and independently — see CLAUDE.md's own UX-sweep status bullet — but the band-split/reveal-in-sequence design below was never part of it).
6. Slide Vault — **not started**.

Commit at each phase boundary. Run the full suite before each — and remember committed `.draft.json` fixtures are frozen: if Tier 1 fails, the code is wrong, not the fixture.
