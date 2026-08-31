# BACKLOG.md

Living list of queued work. Add to it as ideas arrive; move items out when they ship.
Not ordered by priority — see "Suggested order" at the bottom.

**How to use this:** each item states what it is, why, and what's already decided, so a
prompt can be written from it without re-litigating. Open questions are listed
explicitly — resolve them before building, not during.

---

## Queued

### The avails slide's targeting map sometimes renders with only 1 picture, not 3 — not root-caused
Found running `test_group_scenarios.py --render` on Annapolis Cars while visually
verifying FLOW_REWORK_PLAN.md Phase 3 (2026-08-28), unrelated to that phase's own
changes — confirmed by reproducing the identical failure on the commit immediately
before Phase 3's last commit landed. `place_targeting_map` reports it ran, and a map
PNG is confirmed drawn, but the avails slide's own shapes sometimes come back with
only 1 picture (a high shape id, e.g. "Picture 82") instead of the expected 3
(background, PREMION wordmark, map) — reproduced twice in a row on the same machine.
Not yet bisected against `assembly.place_targeting_map`/`copy_slide_into`'s picture
insertion, and not known whether it's state left over from a prior COM session in the
same process, an ordering issue, or a genuine picture-placement bug. **Trigger to
investigate:** a rep reports a missing map on a real generated deck, or someone wants
to root-cause it before then. Not blocking Phase 3 — `test_group_scenarios.py --render`
without `--render` still passes; the render-only check is the one that catches it.

### A separate "slow tier" for the full sweep — not now, and scoped to two files only
Measured while gating FLOW_REWORK_PLAN.md Phase 3 (2026-08-28): the full 68-file sweep
is 1431s (24 min), and it's a genuine long tail rather than evenly spread -- the top 10
files are 51% of the wall clock, the bottom 41 (60% of the suite) are 15% of it. Reviewed
file by file: most of the slow ones (`test_group_plan_selection.py`,
`test_group_builder.py`, `test_avails_grid_row_deletion.py`, `test_form_state.py`, etc.)
are real coverage spinning up the actual Streamlit form via `AppTest` -- that startup
cost is the price of testing the real app rather than mocks, not something to split
off or speed up. **Decision: leave the sweep as one command.** If a slow tier is ever
built, it should hold exactly two files, both with a real, separate reason to be slow --
`test_targeting_map.py` (the pathological CPU-bound slowdown below) and
`test_group_geo_resolution.py` (real network calls to the free Census Geocoder, ~0.15s
per address across dozens of addresses -- legitimate, not a bug). Everything else stays
in the one sweep.

### `test_targeting_map.py` runs pathologically slowly, inconsistently — not root-caused
Found by `tests/run_all.py`'s first-ever full-sweep run (2026-08-27), unrelated to
whatever prompted building the sweep. Standalone reruns stalled for 6+ minutes at
different points in the file across attempts under otherwise-identical conditions --
confirmed as real, ongoing CPU-bound work (a stalled process's own `UserModeTime` was
seen climbing), not a deadlock. One stalled process was terminated. Full write-up in
DECISIONS.md. **Trigger to investigate:** a rep hits this live in the real app, or
someone decides it's worth the time before it's needed -- `render_map`/the legend-fit
helpers (`_fit_legend_label`/`_fit_combined`) are the most likely area, since the one
run that stalled almost immediately did so on the first NAMED group's render, and
every other render in the same file (unnamed groups) was fast.

**Distinct from a real, unrelated fixture bug in the same file, found and fixed
2026-08-28 while gating Phase 3.** The file's own end-to-end AppTest scenario never
set `flight_start`/`flight_end`, so the Phase 1 setup-band gate silently returned
before Generate ever rendered -- a deterministic bug, not a symptom of the slowness
above, and unrelated to Phase 3 (it predates Phase 3 entirely; `ff0d2e5`, the commit
that fixed this same gap in ~20 other AppTest suites, missed this one file). See
CLAUDE.md's "grep, not recollection" rule, added the same day for exactly this
pattern. Fixed by adding the same three session_state lines the other ~20 files got.

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

- **Different budgets for different date ranges within one flight, stated in the notes
  ("$5K a month through October, then $8K a month after") — explicitly out of scope.**
  Surfaced deciding FLOW_REWORK_PLAN.md Phase 6's flight-ownership rule (avail fills the
  dates, rep revises in the band, drafting only ever seeds an empty flight or flags a
  disagreement — never reshapes it). Real, but rare, and it's a different kind of problem
  than flight ownership: it's about the PLAN BUILD (one line splitting into several,
  each keyed to its own sub-range of the flight) and the allocation algorithm (which
  today prices one line against the whole flight, never a date-scoped slice of it), not
  about who owns the frame. Not worth reshaping the drafting flow or
  `resolve_drafted_lines`'s waterfall around a case this narrow. **Trigger:** a real
  proposal needs it and the workaround (a rep splitting the line by hand into two rows,
  each covering its own portion of the flight via Section E's own per-row Flight text)
  turns out to be more than an occasional inconvenience.
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
