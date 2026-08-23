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

### Co-viewing coefficient
CTV lines only. Streaming is usually a group activity; we transact on household
impressions but sometimes want to acknowledge the likely wider reach — understated, and
transparent that it's a projection. Include a CPM variant with the factor applied.

**Prior art:** an old deck did this as a separate "Monthly Coviewing" column showing the
*additional* impressions, a dash on the retargeting line, and a footnoted study. That
shape is right — keep it.

**Multiplier — important:** the +50% (1.5) previously used is **not supported** by
TVision's published data. Their CTV VPVH runs roughly 1.29 (2022 report) to 1.44, so the
defensible bump is 30–45%. Citing TVision while quoting above their figure is an
exposure on a signable page.

**Decided:** make the multiplier *and its citation* one configurable settings record —
value, source, study date, footnote text — so the footnote can never drift from the
number it justifies, and updating the source is a data change.

**Open:** ask Premion whether the "TVision / Premion Co-viewing study" in the old
footnote is a Premion-commissioned study on their own inventory; that would be better
than any general figure. Also note Nielsen's new wearable-based co-viewing methodology
takes effect 1 Sept 2026, which may publish a better citation shortly. TVision is being
acquired by Viant — check the citation still resolves.

**Also worth knowing:** co-viewing varies ~2x by app and runs much higher in households
with children, so any flat multiplier is genuinely approximate — which supports the
understated presentation.

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

Andrea's live bugs, the avails PDF importer, the UX sweep, and % SOV have all shipped
since this was last ordered; audience usage ingestion's code is done too, waiting only on
Matt running the live Supabase upload. That leaves:

1. Audience usage ingestion — Matt's own step (`setup_supabase.py audience_usage_bucket`,
   then upload/activate the real workbook), not a coding task
2. Co-viewing — needs the multiplier and citation settled first
3. Feedback loop — its value compounds once other people are using the app
4. Slide vault — mechanical, architecture already exists
