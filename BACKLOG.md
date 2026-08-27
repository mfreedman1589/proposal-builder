# BACKLOG.md

Living list of queued work. Add to it as ideas arrive; move items out when they ship.
Not ordered by priority — see "Suggested order" at the bottom.

**How to use this:** each item states what it is, why, and what's already decided, so a
prompt can be written from it without re-litigating. Open questions are listed
explicitly — resolve them before building, not during.

---

## Queued

### Slide vault
Colleagues add slides they like; saved centrally and manually selected into decks.
Separate from the master deck.

**Architecture:** this is the case-study vault again — separate library, uploaded by
anyone, tagged, selected at generate, rendered to images. That pattern exists and works,
so this should be a fast build rather than a design exercise.

### Response rates as a second audience ranker
Later addition to the usage workbook. Rank segments by performance, not just volume.
Schema is being built to take a second metric column without a migration.

---

## Smaller / carry-over

- **Per-period avails rates, deferred until a real document needs them.** FLOW_REWORK_PLAN.md
  Phase 2's avails-import freeze derives `avails_monthly` from one flat daily rate (the
  document's own total ÷ its own day count), not from `AvailsGroup.periods`' real per-row
  figures, even for a monthly-broken-out document. Measured against every real document on
  hand (Capital Media's four rows): a per-period rate model is numerically identical to the
  flat one, so building it now would add a second rate model for zero observed effect.
  **Trigger:** a real document is found whose periods genuinely disagree with their own
  implied flat rate (Wilmington's real per-audience monthly figures are the closest
  candidate on hand, but haven't been checked against this specific question).
- **Day-prorating a plan row's own cost/impressions for a partial month** — explicitly out
  of Phase 2's scope, belongs with Phase 3's per-section basis override. A Monthly-breakout
  row's stored Cost is a rate a client signs per month; scaling it by a day fraction for a
  partial first/last month would silently rewrite priced money with no row ever "rewritten,"
  the exact hazard the flight-change guard exists to catch. See FLOW_REWORK_PLAN.md.
- **A finer per-month "Adjust to plan dates" control.** The current action (D2's avails
  divergence panel) is one button per group — it reduces the document's own frozen daily
  rate across the plan's ACTIVE days as a whole, not a per-month override a rep could tune
  month by month. Revisit only if a real proposal needs the avails figure adjusted for some
  months but not others within the same group.
- **Harrisburg customization track** — stubbed and minimal; fill in when Harrisburg
  proposals actually diverge from DC.
- **AM-format media plan table** — the Deliverable/Rate column style from the Capital One
  Hall example. Only matters for AM-only proposals.
- **"Add external proposal"** — filing decks the app never generated. Open question
  whether these belong here at all versus SharePoint, since a row without a recipe loses
  the reload/revise value.
- **Media plan pagination** — deliberately unbuilt until a plan exceeds ~16 rows.
- **Five DMAs with no profile slide** (Honolulu, Palm Springs, Anchorage, Fairbanks,
  Juneau) — gap in Premion's source deck; worth asking them to author.
- **Deck storage retention** — old master deck versions are kept forever for
  rebuild-as-presented, growing ~44 MiB per update against a 1 GB tier. Needs a policy
  before it's urgent.
- **Feedback loop fast-follows.** The core loop shipped 2026-08-23 (sidebar popover,
  state capture, `Feedback reports` admin page — see CLAUDE.md). Two things deliberately
  deferred out of that first pass: an optional screenshot attachment (state capture alone
  covers most of the reproduction value); and loading a report directly into the form the
  way a History entry does (the captured state is a curated subset -- `targeting_groups`,
  `plan_options`, the scalar Section A/flight fields -- not the full `form_json` recipe a
  rebuild needs, so "load into form" needs either a richer capture or an explicit
  partial-restore UI, not a small addition to what's there now).
- **Move the avails PDF uploader next to the notes box in the intake area.** Both already
  live in the "📥 Start here" block, but as separate items a rep can do in either order.
  Drafting only asks the model to select from real, already-imported avails rows when
  those rows already exist at draft time (see CLAUDE.md's targeting-groups section, "a
  draft run before any avails work exists") -- a draft run first invents its own Premion
  Streaming TV line(s) from the notes, and nothing retroactively reshapes them once a real
  avails table shows up. That's a real, if narrow, gap: a rep who drafts before importing
  gets a plan that doesn't trace back to the avails table at all, caught today only by an
  after-the-fact note in the review list, not prevented. Putting the two side by side
  (upload first, draft second reads naturally left-to-right) makes the working order the
  natural order without anyone needing to know why it matters -- worth doing as a small,
  standalone layout change, not a rebuild-and-carry mechanism.
- **Known limitation: nested radius tiers are invisible on the targeting map.** A real
  Annapolis Cars document (RFPID-253813) sells each of 4 audiences as a 10-mile radius
  PLUS a 5-mile radius, the 5-mile zip set a strict subset of the 10-mile one. `_touched_
  counties` correctly treats one audience's several groups as one entity (never a false
  self-overlap), but the map has no way to show a nested inner tier: both radii paint in
  the audience's one color, no boundary separates them, and the dots for the 5mi ring look
  identical to the 10mi ring's. **Client-facing consequence:** a rep sells a 5-mile core
  inside a 10-mile buy and the map shows one undifferentiated shape — a client can't tell
  the two tiers apart from the picture, only from the avails table beside it. Confirmed by
  rendering the real document (2026-08-23); deliberately not designed or fixed yet —
  report-only per that session's own decision. Any fix needs a way to show relative
  containment (a lighter/darker shade per radius tier? a second, smaller dot?) without
  reopening the "no filled region, only real avails" rule `targeting_map.py`'s own
  docstring already settled.

---

## Suggested order

Andrea's live bugs, the avails PDF importer, the UX sweep, % SOV, the co-viewing
coefficient, audience usage ingestion (version 1 active in production since 2026-08-21),
and the feedback loop's core (2026-08-23) have all shipped since this was last ordered.
That leaves:

1. Slide vault — mechanical, architecture already exists
