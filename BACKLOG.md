# BACKLOG.md

Living list of queued work. Add to it as ideas arrive; move items out when they ship.
Not ordered by priority — see "Suggested order" at the bottom.

**How to use this:** each item states what it is, why, and what's already decided, so a
prompt can be written from it without re-litigating. Open questions are listed
explicitly — resolve them before building, not during.

---

## Queued

### Attribution report advertiser matching: two pixel-carrying rows silently picks only the first
Found 2026-09-08 auditing the WAEPA export's ADVERTISER tab, which has two rows — one
carrying the `WUSA WAEPA - TEGNA` pixel (4,323 impressions), one carrying no pixel at all
(605,437 impressions, ~99% of the file's total delivery). Checked and confirmed **not a bug
for that file**: `attribution_import.parse_attribution_export` already filters to
pixel-carrying rows before picking `rows[0]` for `client_name`/`advertiser_pixel`/
`market_hint`, so the no-pixel row is correctly excluded regardless of its position in the
sheet, and `advertiser_matching.find_candidates` never touches the raw tab at all — only
those two already-correct fields.

**The latent gap, not reachable by any real export seen so far:** if an export ever has TWO
rows that both carry a pixel, with different client names or different station markets
(a client running campaigns through two different TEGNA stations, each with its own pixel,
in the same export), `rows[0]` after filtering silently picks whichever one happens to sort
first — the second pixel row's name/market is dropped with no warning. Every real file this
app has parsed (MW, Cardinal, WAEPA) has exactly one pixel-carrying row.

**Trigger to investigate:** a client runs two stations and their export actually shows two
pixel rows — until then this is a real, understood gap, not a live bug. If it comes up:
either warn when `len(pixel_rows) > 1` (WAEPA's own multi-RFPID confirm-gate shape is the
precedent — never a hard reject) or, if the two rows are genuinely the same advertiser
split by station, decide whether picking rows[0] is fine as-is (client_name is usually the
same either way) versus concatenating market hints.

### Every generated deck carries every unused slide master and layout — `delete_slide` never prunes them
Found 2026-09-05 building the standalone avails-slide export, which made an existing,
previously-invisible inefficiency impossible to miss.

**Root cause, confirmed:** `assembly.delete_slide` removes a slide's own `<p:sldId>` entry
and relationship — that's its whole job, and it does it correctly (a slide's own unique
media really is dropped, per python-pptx's reachability-based part pruning on save). But
it never touches `presentation.xml`'s `<p:sldMasterIdLst>`, which references **every one of
the master deck's 21 slide masters unconditionally**, regardless of which slides survive.
Each master pulls in its own slide layouts (226 total across all 21) and its own background
images. None of that is reachable-pruned, because the presentation part's own relationship
to each master is never dropped — only a slide's relationship to ITS layout/master would
need to be, and slide deletion doesn't do that either (nor should it blindly, since several
slides can share one master).

**Measured, not estimated:**
- Master deck: 120 slides, 21 masters, 226 layouts, 47.5 MB.
- A typical "extended"-preset proposal (41 slides kept): actually resolves to only 16 of
  the 21 masters via its own slides' `slide_layout.slide_master`, but the saved file still
  carries all 21 -- 32.5 MB, only 32% smaller than the master despite dropping 84 of 125
  slides (67% of them).
- The new standalone avails-slide export (1 slide kept, this feature's own case): uses
  exactly 1 of 21 masters, but the saved file is 14.2 MB -- barely a quarter of the
  master's size for what is genuinely one slide's worth of content.

**Why this matters now and didn't before:** every proposal has always carried some of this
dead weight (a 41-slide proposal wastes ~24%, 5 masters' worth), but it read as normal file
size for a real, image-heavy proposal deck. A ONE-slide export making the same mistake at
~95% waste is what made the pattern legible -- a rep emailing a one-slide prospecting
attachment at 14MB is the kind of thing that gets bounced or ignored, which is what
surfaced this.

**Not attempted here, deliberately -- real, separate work with real risk:** a fix means a
package-level pass, after slide deletion, that determines which masters/layouts are still
reachable from the SURVIVING slides only (via each kept slide's own `slide_layout`/
`slide_master`, not presence in `prs.slide_masters`), prunes the unreachable ones' XML parts
AND removes them from `sldMasterIdLst`/each layout's own `sldLayoutIdLst`, and is re-verified
against `package_check.check_package` and a real COM render on a realistic proposal (not
just the synthetic single-slide case that found this) before trusting it on every deck this
app produces. Masters/layouts are shared, load-bearing infrastructure -- a rendering
regression here would be worse than almost any other bug this app could ship, which is
exactly why this wasn't attempted as a side effect of the standalone-slide feature that
found it.

**Trigger to investigate:** a decision to actually build the pruning pass -- likely
alongside `optimize_deck.py` and `db.prepare_deck_for_upload`, the existing size-management
code this would sit beside, and worth scoping against `assembly.py`'s own slide-deletion
tests (`package_check`, `tests/render_scenarios.py`) as the verification bar.

### Geo intelligence — nearest-location zip assignment when radii overlap
Scoped alongside the standalone map builder, 2026-09-04, also not planned yet — same
reason this file entry exists.

**The problem:** five dealerships, a radius around each, and the radii overlap.
Today a rep either accepts the duplication or hand-assigns zips. Want: assign each
zip to one dealership — nearest, or by drive time — or split the overlap
deliberately by some rule.

**Scoping question asked, not "build this":** what geo data does the app already
have, what would nearest-dealer assignment actually take, and where does it
genuinely need something beyond arithmetic?

**Investigated 2026-09-04:**
- `geo_crosswalk.json.gz` (`geo_resolver.py`, built from three public-domain Census
  files) already carries a lat/lon centroid for every US zip (`_data()["zip_points"]`)
  and `haversine_miles` is already implemented and used by `radius_to_zips`. This is
  exactly what nearest-by-straight-line-distance assignment needs — no new data, no
  new dependency.
- A dealership's own radius origin (a zip or a street address, `geo_def["centers"]`)
  resolves to a point via `geo_resolver.resolve_center` / `geocode_address` (the free,
  keyless Census geocoder already used for radius mode) — but that point isn't
  persisted on the group today, only the resolved zip list is. Re-resolving it for a
  handful of dealership centers at assignment time is cheap (the app already accepts
  ~0.15s/address for up to 20 addresses in radius mode) — not a reason to store it.
  So: **assigning by straight-line ("as the crow flies") nearest is pure arithmetic
  over data the app already holds** — for each zip in the union of overlapping
  radii, compute `haversine_miles` to every dealership's own center, assign to the
  minimum. No AI, no new dependency.
- **Where it genuinely needs something beyond arithmetic:** actual drive time (not
  straight-line distance) requires an external routing service (e.g. a distance-matrix
  API) — a new dependency, likely paid or rate-limited, unlike the free Census
  geocoder this app leans on everywhere else. That's the one real gap, and it's a
  build/cost decision, not a modeling one.
- **Splitting the overlap "by a rule"** (deliberately, rather than winner-take-all
  nearest) is also pure arithmetic once a rule is chosen (e.g. a fixed percentage
  split, or a weighting) — the open part is a business decision about which rule,
  not a technical one.
- **Not yet investigated:** how this interacts with `_touched_counties`/the map's
  overlap-hatch rendering (`targeting_map.py`) once zips are exclusively assigned
  rather than shared — the 2-vs-3-way overlap-fill logic and the "nested radius tiers
  invisible on the map" known limitation (see "Smaller / carry-over" below) both
  assume today's model where a zip can belong to more than one group's resolved set.


Found 2026-09-01 investigating why LiveWell's D2 grid showed the same string in both the
Markets and Geo Label columns (a SEPARATE bug, actually caused by `install_market_lookup()`
ordering and already fixed — see DECISIONS.md). `avails_pdf_import.classify_geography`'s
`_BRACKET_ORIGIN_RE` extracts a radius's center only from a bracketed form (Annapolis's
`"10mi radius [21401]"`); LiveWell's own real documents state the center as a plain,
unbracketed street address (`"277 S Washington St Alexandria VA 22314 5 Mile Radius"`),
which the regex doesn't recognize at all — `radius_origin` comes back `""`, so these
groups never reach real radius-mode geocoding (`resolve_group_geography(GEO_MODE_RADIUS,
...)`) and fall through to the zip-list path (`_resolve_zip_originated_geo`) instead,
using the document's own stated zip list rather than a freshly-geocoded one.

**Deliberately not fixed alongside the labeling work that touches the same regex**, per an
explicit call on this: whether to re-geocode from the real address or keep the document's
own zip list is a judgment call, not an obvious bug — the document's zips are what Premion
actually priced the avails figure against, and swapping in a geocoded list could produce a
DIFFERENT footprint than what was quoted, which is a real risk in the other direction. A
label is display text a rep can fix in one click if it's ever wrong; `resolved_zips` feeds
real reach numbers on a client-facing document, and deserves a more careful decision than
"the regex was already open, might as well."

**Trigger to investigate:** someone decides re-geocoding real-address radius origins is
worth the accuracy trade-off (or confirms the document's own zips should stay authoritative
regardless, closing this without a code change) — likely alongside `geo_targeting_roadmap.md`
if it's built, since that's where this app's other geocoding-vs-document-truth calls live.

### CLOSED (2026-09-23) — "avails slide sometimes renders with only 1 picture, not 3"
Found 2026-08-28, against the OLD pre-map-variant-slide design, where the map was
drawn as an overlay on the STANDARD (photo) template and the slide legitimately
carried three pictures (background, PREMION wordmark, map). The map-variant slide
(assembly.TARGETING_AVAILS_MAP_KEY) redesign superseded that mechanism outright --
the map-variant template carries no stock photo or slide-level wordmark of its own
by design, so "three pictures" stopped being the right expectation for any deck
built against it. Never actually re-investigated as the pre-redesign flake this
entry describes; superseded by the redesign before anyone needed to. A related but
DIFFERENT issue was found and fixed 2026-09-23 (not this one): the active master
deck changed that same day ("Background change for avails," baking a real full-
slide picture onto the map-variant slide), which made `test_targeting_map.py`/
`test_standalone_avails_builder.py` expect the wrong TOTAL picture count -- fixed
by asserting on the map picture's own identity (byte-match / size-match) rather
than a total count, so a future background edit can't cause the same miss again.

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

### `test_group_plan_selection.py` failed once in a full sweep, passed standalone immediately after — not root-caused
2026-09-01 (later same evening as the settled-clean sweep above): a second full
`run_all.py` sweep (69 files, run for a reason not captured in this note) reported
`test_group_plan_selection.py ... FAIL (43.9s)`, alongside 68 other files passing.
The captured failure tail was only the run's own trailing `use_container_width`
deprecation warnings and `missing ScriptRunContext` noise — the actual assertion
failure line was above the 4000-char tail window and wasn't preserved. A standalone
rerun immediately after, `python tests/test_group_plan_selection.py`, passed clean,
exit 0, every check green including the multi-option/LiveWell-shape scenario. Same
shape as the `test_targeting_map.py` entry above — intermittent, not reproduced on
demand — so tracking it here the same way rather than either dismissing it as noise
or treating one clean rerun as proof it's fixed.

**Run history (tracking frequency, not presence):**
- 2026-09-01 (earlier, gating the $0-phantom-row fix): full sweep — PASS, 46.8s.
- 2026-09-01 (later): full sweep — FAIL, 43.9s, assertion detail not captured;
  standalone immediate rerun — PASS, clean.

**Trigger to investigate:** a third occurrence, or a captured failure with the
actual assertion line intact (rerun a failing sweep with `run_all.py
test_group_plan_selection` alone and capture full output, not just the tail, if
this happens again).

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
- **Attribution Report Builder Phase 9 — multi-dealer group reports.** Gated on a real
  multi-dealer group report set (several dealers' website/Polk/Analyst files, one group,
  one period) — don't design against synthetic data. See ATTRIBUTION_REPORT_PLAN.md's
  "Deferred, named for continuity" section for the full spec and open questions.
- **Attribution Report Builder: Auto-Sales Analyst facts JSON → cross-source takeaways.**
  Gated on Matt adding a facts JSON export to the (separate-repo) Auto-Sales Analyst app,
  plus a real report with website/Polk/Analyst data for the same dealer and period. See
  ATTRIBUTION_REPORT_PLAN.md's "Deferred, named for continuity" section.
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
the feedback loop's core (2026-08-23) and the slide vault (2026-09-02, see CLAUDE.md) have
all shipped since this was last ordered. Nothing is currently ordered -- pick from
"Smaller / carry-over" above, or from a fresh priority conversation.
