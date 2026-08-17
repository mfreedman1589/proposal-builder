# Geo targeting & zip/map roadmap

Six pieces. A and B are small and independent. C is the foundation everything
else calls. D changes the shape of the avails table and is the invasive one.
E and F sit on top.

**Sequencing:** A and B (either order) → C (report before wiring) → D (on its
own, not alongside anything) → E → F (needs real sample PDFs in hand).

## Decisions already settled

- Markets break out as separate avails rows by default, with a toggle to
  combine them into one campaign line.
- Markets resolved from a zip list are added automatically, but easy to remove.
- Each avails row seeds its own campaign line, mergeable and splittable
  afterward.
- The map sits beside the avails table on the targeting slide, replacing the
  stock image in that position.
- No paid services anywhere — zip → county → DMA via free data.

---

## A — Selected markets seed the avails table ✅ done

One avails row per selected target market, replacing the originating-market
default rather than sitting beside it, plus a "combine markets into one row"
toggle for the case where several markets sell as a single campaign line.
Same clean/dirty discipline as everywhere else. See
`tests/test_avails_seeding.py`.

## B — Quick-add lines from the avails table ✅ done

An "Add lines" panel above the media plan grid: Product / Audience / Geo
multi-selects generating the cross product as separate lines, with a combine
option. Grid Audience (the `Targeting` column) and Geo are single-select
dropdowns sourced from the avails table. See `tests/test_quick_add_lines.py`.

## C — Geo resolver (no UI) — not started

Zip ↔ county ↔ DMA crosswalk plus pure functions: zips → markets, county list
→ zips, radius around a zip or address → zips, market → zips. Every function
reports what it couldn't resolve. **Report data sources, sizes, licensing and
function signatures before wiring anything in.**

## D — Targeting groups — not started

The invasive one. A targeting group is a name, an audience, a geo definition
(market / county / zip / radius), the resolved zips and markets, an avails
figure, and a map colour. Groups flow outward into the avails table, proposal
lines, market slides and geo defaults. Existing flat-avails proposals must
load, rebuild and render identically — migrate flat rows into groups on load
rather than changing what is stored for past proposals.

### The audience half of a group is BUILT, not just picked

Extend the Audience finder into a builder: search and add a segment, then add
further segments with one of three actions.

- **AND** (the default) narrows the audience — "Homeowners AND HH Income $150K
  Plus" is one audience defined more precisely. This is how combined audiences
  are almost always built.
- **OR** places two distinct audiences on the same campaign line — "College
  Planning Parents OR Education Services". Less common but real.
- **Add as a separate group** creates a new targeting group with its own avails
  row and its own campaign line.

**Each group carries one avails figure for its whole expression**, since avails
are a property of the combination — an AND stack is smaller than either segment
alone, an OR set larger. That is the same unit the Salesforce avails request
consumes, so a built stack should be directly requestable.

**Enforce the one-custom-audience rule as the stack is built, not at submit.** A
second non-RFP-selectable segment is refused or flagged the moment it is added,
while the strategist is still choosing. (Today the cap is enforced structurally
in `apply_draft_to_form` and warned about at the table; neither tells you at the
point of the decision.)

**Show booking evidence while building.** Framing settled, in this order:

0. **Match as a set.** Normalize every stack to a sorted set of components so
   part order doesn't matter — 24 stored combinations appear in more than one
   order, and string matching would miss them.
1. **Exact match — only when true, never as a warning.** "This exact audience
   has been booked before", quietly. **Never show a frequency for it** (2,745 of
   2,769 stacks were booked exactly once, so a count carries no information) and
   **never show "not booked before"** — absence is the normal case, and a flag
   that fires on ordinary audiences trains people to ignore the panel.
2. **Component familiarity — lead with this.** Per selected segment, how many
   booked stacks it appears in. This answers the question a strategist actually
   has: am I reaching for something standard or something exotic?
3. **Pairwise co-occurrence.** How often the chosen segments have appeared
   together. With three or more, surface the **weakest** pair, since that is the
   least precedented part of the stack.
4. **Suggested pairings.** Given what's selected, which components most often
   accompany it — the same data read forward instead of backward, which is what
   is useful while someone is still choosing.
5. **Nearest-neighbour stacks** as examples if cheap. Least important of the
   set; skip rather than delay on it.

**What the data says about each signal** (measured across all 3,661 rows, so
none of this is assumed):

- **Component familiarity is the strong one, as expected.** 1,302 distinct
  components with a real spread — 40 appear in 26+ stacks, 246 in 6–25, 467 in
  2–5, and **549 appear exactly once**. That 42%-of-components tail is precisely
  what makes "standard vs exotic" a meaningful thing to say. Worked example:
  `DEMO Homeowner` 172, `HH Income 150K Plus` 153, against `CLT 1P NFCU December
  Audiences Look Alike` 5.
- **Pairwise co-occurrence holds up better in practice than in aggregate.**
  82% of the 3,176 observed pairs co-occur exactly once — but that tail is pairs
  nobody would build. The pairs a strategist actually reaches for score usefully:
  `DEMO Homeowner + HH Income 150K Plus` 13, `DEMO Age A35 Plus + HH Income 100K
  Plus` 18, and even the niche `NFCU Look Alike + Military Families` 3. Still,
  a zero here should read as "these are both common but nobody has combined
  them" — informative — rather than as a warning.
- **Suggested pairings need overlap weighting, not a strict superset.** Asking
  which components accompany *all* selected segments is excellent for one
  segment (`DEMO Age A25 Plus` → Homeowner 14, HH Income 75K Plus 10, AUTO
  Intenders 6) and **collapses to a flat list of 1x at two segments**, because
  few stacks are supersets of a specific pair. Weighting each candidate by how
  many of the selected segments its stack shares restores it: for `DEMO
  Homeowner + HH Income 150K Plus` that yields `DEMO Age A35 Plus` 24, `DEMO Age
  A35-64` 20, `DEMO Age A25 Plus` 18 — real recommendations rather than noise.
  **Build it weighted from the start**; the strict version looks fine in a
  one-segment demo and dies on the second click.
- **Normalize before matching — the stored names are not clean.** 12 component
  groups differ only by case or punctuation: `Lifestyle Charity` (25x) vs
  `LIFESTYLE Charity` (6x), `Lifestyle Outdoors` (15x) vs `LIFESTYLE Outdoors`
  (2x), `LIFESTYLE Pets - Horses` vs `LIFESTYLE Pets Horses`. Matching raw
  strings would split one segment's real history in two and understate
  familiarity by up to 4x. Fold on lowercase-alphanumeric, the same
  normalization the market matcher already uses.

**The expression must round-trip:** stored on the group, restored by history,
and rendered as readable text in the avails table, the Targeting column and
Campaign Specs. A client reads "Homeowners, $150K+ households", never boolean
syntax.

### Finding: the comma in `audience_usage` is AND — confirmed, not assumed

Checked against all 3,661 rows (2,793 of them combinations). The assumption was
right, and three independent lines of evidence agree:

1. **Explicit operators are spelled out INSIDE a part; the comma never carries
   them.** 250 rows contain the word "and" and 121 contain "or", always within a
   single component — `AUTO Truck Intenders OR SUV Intenders, HH Income 75K
   Plus`, `CLT 1P Lennar Homes Buy Sell Home AND HHI 50k Plus`, `CUSTOM DEMO
   Homeowner and HH Income 200K Plus, HH Homes Built Before 2015`. When they
   mean OR they write OR. The comma is reserved for the stack. This is the
   decisive one.
2. **86.5% of combinations join DIFFERENT categories** (DEMO + HH, AUTO + DEMO)
   — narrowing across attributes. The 13.5% that share a category are AND too,
   because a category is a namespace rather than one dimension: `AUTO Intenders,
   AUTO Make Subaru` is Subaru intenders; `DEMO Age A35 Plus, DEMO Homeowner` is
   homeowners over 35.
3. **1,195 of 1,250 comparable combinations (95.6%) delivered no more than their
   smallest component did standalone** — `AUTO Intenders, DEMO Age A25-54`
   delivered 35.4M against 40.0M and 53.2M. An intersection delivers less than
   either part; a union would deliver more. Directional rather than proof, since
   delivery is budget-driven, but it points the same way.

**Two consequences for the builder:**

- **Look up booking evidence as a SET, not a string.** 24 combinations appear in
  more than one part order (`DEMO Homeowner, HH Income 100K Plus` and `HH Income
  100K Plus, DEMO Homeowner` are both stored). Exact string matching would miss
  a stack that has in fact been booked.
- **Exact-stack counts are nearly useless as a frequency signal, and the feature
  should not lean on them.** 2,745 of the 2,769 distinct stacks were booked
  exactly once; 24 twice; **nothing more than twice**. So "how often has this
  been booked" is effectively binary and will read "never" for almost anything a
  strategist builds. Component-level evidence is far richer — 972 distinct
  components, with `DEMO Age A25 Plus` in 222 stacks, `HH Income 100K Plus` in
  177, `DEMO Homeowner` in 171. Better framings: "this exact stack has been
  booked before" (yes/no), "these components appear together in N booked
  stacks", or nearest-neighbour stacks that share most components. Decide the
  framing before building, or the panel will confidently say "never booked"
  about a perfectly ordinary audience.

### Findings carried into D

**`st.column_config.MultiselectColumn` exists and is editable** (Streamlit
1.60.0, verified functionally through `st.data_editor`, not just from the
docstring — it round-trips a list per cell and takes `accept_new_options`,
and is read-only only in `st.dataframe`).

It was deliberately NOT used in B, for two reasons that both point at D:

1. **It expresses the exception, not the default.** A cell holding
   `["Denver", "Atlanta", "Phoenix"]` means *one line covering three markets*.
   B's default is *three separate lines*, which a multi-value cell cannot
   express at all — so the grid would offer only the uncommon case.
2. **It changes the row data model.** Every plan row's `Geo` is a string today,
   flowing into `resolve_row_defaults`, the totals, the preview and the deck's
   table fill. Making it a list touches all of that.

A targeting group genuinely *is* one object spanning several markets, so this
is the column type that fits D — where the row-model change belongs anyway.
`SelectboxColumn`, used in B, has **no** `accept_new_options`: a value outside
its options survives and displays fine (verified), but a rep cannot type a new
one into the cell. B works around it by seeding the options with every value
already on the grid plus the avails values; new values come from the panel
(`st.multiselect` *does* have `accept_new_options`) or the avails table.

## E — Zip/map builder page — not started

Its own page. Inputs resolve through C; outputs are a copyable/exportable zip
list and a map image placed beside the avails table on the targeting slide,
replacing the stock image. Multiple groups render in different colours with a
legend. **Recommend an approach before building**, and confirm rendering works
on Streamlit Cloud's Linux environment, not just Windows.

## F — Salesforce avails PDF import — not started

An avails PDF importer beside the avails table, same shape as the Wide Orbit
schedule importer. Parses per-audience zip lists and avails figures, populates
the table and creates a targeting group per audience with zips resolved to
markets through C. Fixtures gitignored — they contain real client avails and
pricing.

### F amendment — avails upload as part of the drafting loop

The common real flow is that a rep mentions avails in their notes and has the
Salesforce document in hand, so the upload belongs in the draft review list as
well as beside the table.

- The draft schema gains a signal for whether the notes reference an avails
  pull, distinguishing **three** cases: figures actually stated, avails
  referenced without figures, and no mention at all.
- Where the notes reference avails, the review list shows an **inline
  uploader** at that item rather than a text reminder — "Your notes mention
  avails you've pulled. Upload the document to fill in the table and zip
  targeting."
- Where figures are stated but no document is uploaded, keep the figures and
  still offer the upload, noting that the document adds zip lists and market
  coverage the notes can't carry.
- Where nothing is mentioned, the existing internal review item stands
  unchanged.
- The inline uploader is the **same** uploader as the one beside the avails
  table — a second entry point, not a second implementation — and runs the
  same import and precedence rules: **rep edits > avails document > draft >
  defaults**, with conflicts flagged rather than silently resolved.

The draft-schema signal is buildable independently of the importer; the
uploader wiring is not, since there is no importer to point it at until F.

---

## To gather

Two or three real Salesforce avails PDFs (ideally one with several audiences
and zip lists), plus one real example each of a county list, a zip list and a
radius request as a client would actually send them. Those shape D and F more
than any design conversation will.
