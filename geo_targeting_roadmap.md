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
