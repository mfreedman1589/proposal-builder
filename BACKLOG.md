# BACKLOG.md

Living list of queued work. Add to it as ideas arrive; move items out when they ship.
Not ordered by priority — see "Suggested order" at the bottom.

**How to use this:** each item states what it is, why, and what's already decided, so a
prompt can be written from it without re-litigating. Open questions are listed
explicitly — resolve them before building, not during.

---

## In flight

### Audience usage workbook ingestion
Versioned upload of the YTD workbook as a catalog source alongside the brochure. Adds
~835 components the brochure never carries. **Code is done** — the markup is applied
(`audience_uncategorized_markup.csv`), the client-retargeting exclusion rule is
generalized (`is_client_pattern` plus `audience_component_overrides.csv` for the
structural-pattern misses), `setup_supabase.py audience_usage_bucket` exists, and both
the import and the admin page are offline-tested. **Remaining is operational, not
code**: run `setup_supabase.py audience_usage_bucket` against the real Supabase project,
then upload and activate the real workbook through the admin page — needs the live
Supabase service key this assistant can't reach, so it's a step for Matt to run, not a
prompt to write.

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

---

## Suggested order

Andrea's live bugs, the avails PDF importer, the UX sweep, % SOV, and the co-viewing
coefficient have all shipped since this was last ordered; audience usage ingestion's code
is done too, waiting only on Matt running the live Supabase upload. That leaves:

1. Audience usage ingestion — Matt's own step (`setup_supabase.py audience_usage_bucket`,
   then upload/activate the real workbook), not a coding task
2. Feedback loop — its value compounds once other people are using the app
3. Slide vault — mechanical, architecture already exists
