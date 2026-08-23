# PROJECT_HANDOFF.md

Matt's own handoff note between chat sessions — refreshed at the end of
every session so a cold start (new chat, or someone else) can pick up
without re-deriving what just happened. This is NOT a duplicate of
`CLAUDE.md` (the permanent rules/traps reference, loaded every session) or
`BACKLOG.md` (queued future work) — it's "what happened last, what's
pending, what to do next," and it gets overwritten each time, not
appended to. If you want durable project rules or history, they're in
`CLAUDE.md`/`DECISIONS.md`/`HISTORY.md`; this file just points at them.

**Last updated:** 2026-08-23, end of session.

---

## Where things stand right now

**Pushed.** `main` is up to date with `origin/main` at `7f11d00`. Working
tree clean.

## What happened this session

Started as "what's next" — confirmed audience usage ingestion (BACKLOG's
old #1) was already live (checked Supabase directly: version 1 active
since 2026-08-21), fixed a stale CLAUDE.md status line about it. Started
the in-app feedback loop next, but paused mid-plan when the user reported
real bugs from a real Hershey & Harrisburg avails import (RFPID-260402,
12 groups). **All of that superseded the feedback-loop work, which never
got past planning** — see "What's next" below.

The Hershey investigation turned into the whole session. In priority order,
all shipped in commit `7f11d00`:

1. **Map omitted every DMA-named group** — 95.8% of a real document's
   impressions invisible. `resolve_group_geography` returned
   `resolved_zips=[]` for markets mode by design; the map only draws
   groups with `resolved_zips`. Fixed by populating it via
   `geo_resolver.market_to_zips`.
2. **Two more real parser bugs found investigating #1**, both predating
   this session: `_block_zip_list` silently dropped a wrapped zip table's
   continuation page (Wilmington University read 48 of a real 303 zips);
   County Option groups resolved from Premion's own DERIVED county names
   instead of the rep's actual zip list, over-targeting by ~32 zips
   against a signed plan. **Checked proposal history directly — zero
   stored proposals ever used the avails importer**, so nothing already
   sent to a client was affected.
3. **Default color assignment** indexed by row/group position instead of
   audience ordinal, so two audiences could land on adjacent, hard-to-
   tell-apart palette slots (literally orange+red, on a live report).
   Palette reordered by a greedy max-min CIELAB walk.
4. **D2's "Apply a color to a whole audience"** rendered one button per
   group regardless of detachment (12 buttons on a real import); now one
   control per audience, shown only when something needs reclaiming, and
   fixed to push the audience's SHARED color (not a detached row's own
   color, which would promote the accident instead of undoing it — caught
   by an existing test, not just the new design).
5. **The geo-definition expander** rendered one per group (12 on import);
   now grouped by audience, stacked in one expander via a shared panel
   body (`_geo_panel_body`, since Streamlit disallows nested expanders).
6. **Identity-collapse**: two audiences resolving to IDENTICAL geography
   (confirmed exact identity across all three real avails documents on
   hand) now collapse to one solid legend entry instead of a hatch — a
   RUNTIME check on actual resolved zips, never an assumption about what a
   document contains. This also fixed, by construction, a dot-drawing
   last-write-wins bug the user caught from a render (dots are drawn from
   the same collapsed entries the fill uses now, not patched separately).
7. **Legend text could run past the canvas** — fitting was capped at a
   fixed 272px budget regardless of actual canvas width. Now fit against a
   canvas-derived ceiling; a genuine 2-way overlap keeps both full
   audience names, a 3+-way multi drops to "+N more".
8. **Dark-mode hatch fills were muted by the same rule solid fills use**,
   collapsing a real overlap to 6.6 ΔE apart. Hatches are now exempt;
   measured 33–80 ΔE across the realistic palette pairs.
9. **New review convention added to `CLAUDE.md`**: any UI or slide change
   comes back as a screenshot at normal width plus a rep-lens pass, not a
   written summary. RFPID-260402 is now the standard fixture for
   avails/map/D2 work (added `build_lawn_leisure()` too, for the
   single-audience no-op case).

Verified end-to-end against the real Hershey document through the live
running app (browser-driven PDF upload, not just offline fixtures) before
committing — screenshots sent to the user.

## What's pending / needs the user

Nothing blocking. One known, recorded limitation (not fixed, by explicit
decision): nested radius tiers (a real Annapolis Cars 10mi+5mi buy) are
still visually indistinguishable on the map — see `BACKLOG.md`'s "Smaller
/ carry-over" section for the client-facing consequence.

## What's next once picked back up

The in-app feedback loop (BACKLOG.md) is where this session actually
started and never returned to. Exploration is done (data-layer, app-shell,
and test-pattern conventions all mapped); one design decision was made
(reports are loadable into the form like a history entry) but the layout
question (one page vs two vs a sidebar popover) and the screenshot-
attachment question (now vs deferred) were never asked. Resume from
scratch on those two — nothing durable was written for this feature yet,
it's pure research in a prior turn's context, not in a file.

Per `BACKLOG.md`'s suggested order otherwise:
1. Feedback loop — its value compounds once other people are using the app
2. Slide vault — mechanical, the case-study vault's architecture already
   covers it

## Test suites touched/added this session

`tests/test_targeting_map.py` (heavily extended — identity collapse,
`_fit_combined` redesign, canvas-safety, hatch ΔE, all against real and
synthetic fixtures), `tests/test_avails_pdf_import.py` (continuation-page
guard), `tests/test_group_scenarios.py` (Wilmington's real-zips guard, new
Lawn & Leisure scenario), `tests/test_color_lifecycle.py` (updated for the
one-button-per-audience redesign), `tests/group_scenario_fixtures.py`
(`build_lawn_leisure()` added). Tier 1 (363 checks) and ~10 other suites
(`test_group_builder`, `test_group_geo_resolution`, `test_form_state`,
`test_avails_pdf_wiring`, `test_share_of_voice`, `test_group_merge_split`,
`test_avails_grid_row_deletion`, `test_avails_import_notes`,
`test_avails_placeholder_group`, `test_targeting_groups`) all green as of
the commit above.

## How to keep this file useful

Whoever (whatever session) is working on this project: **refresh this file
at the end of your session**, before ending the conversation — overwrite
the "Where things stand," "What happened," "What's pending," and "What's
next" sections with the current reality, not an append. If nothing
notable happened (a short Q&A session, no commits), it's fine to leave
this file untouched rather than write a near-empty update.
