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

### In-app feedback loop
A persistent "Report an issue" button on every page: category (Avails / Proposal / Map /
Audiences / Drafting / Deck output / Other), free-text notes, optional attachment.

**Capture state, not screenshots.** Attach the current `form_json` (the same payload the
proposals table already logs), page, build stamp, selected user, active deck version,
drafted notes and review lists. That turns "the avails row disappeared" into a
reproducible case.

**Storage must stay light** — a `feedback` table plus a private bucket, but:
- `form_json` is a few KB; that's the bulk of the value
- **Never attach the generated deck** (20–40 MiB each). Store the `proposal_id` instead
  — proposals already hold the recipe and can regenerate the deck on demand
- Screenshots optional and compressed, not raw PNGs

**Admin page:** list newest-first, filter by category and status, mark open/closed, and
export one or several as a single markdown bug report — description, captured state, and
a reproduction recipe pointing at the proposal. That's what gets pasted into Claude Code.
Surface an open-report count somewhere visible.

**Open:** should a report be loadable directly into the form, the way a history entry is?
That would make reproduction one click.

### Response rates as a second audience ranker
Later addition to the usage workbook. Rank segments by performance, not just volume.
Schema is being built to take a second metric column without a migration.

---

## Smaller / carry-over

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
coefficient, and audience usage ingestion (version 1 active in production since
2026-08-21) have all shipped since this was last ordered. That leaves:

1. Feedback loop — its value compounds once other people are using the app
2. Slide vault — mechanical, architecture already exists
