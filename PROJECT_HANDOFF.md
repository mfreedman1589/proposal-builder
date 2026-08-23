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

**Pushed locally, not to origin.** `main` is at `697500e`, three commits
ahead of `origin/main` (`d64bcc0`). Working tree clean. **`! git push`
still needs to run** — nobody has done it this session.

## What happened this session

Picked up exactly where the prior session's handoff left off ("What's next
once picked back up") plus two items the user flagged as specced earlier
but never built. All three shipped, each its own commit:

1. **`85c5eac` — Intake area cleanup.** Each "Documents you have" column
   showed its label twice (an outer `st.caption` plus the file_uploader's
   own label), and the Wide Orbit column's four-line explanatory paragraph
   pushed its Upload button out of horizontal alignment with the other two.
   Fixed: one label per column (folded into the file_uploader's own,
   emoji included), explanatory copy moved into the `?` help tooltip
   (matching the pattern the avails column already used), format/size line
   comes free from Streamlit's own uploader footer. Verified via browser
   screenshot — all three Upload buttons now sit on the same row.

2. **`3fd50a5` — Avails-import report noise, aggregated.** The per-group
   calm-vs-warning collapse (an earlier fix) still repeated the same fact
   once per targeting group — "N zip(s) have no county on file," "...no
   market on file," "...counted in the larger market" turned into several
   near-identical lines on a real 12-group Hershey or 5-group/303-zip
   Wilmington import. Fixed one level up: `apply_avails_import` unions
   every zip-originated group's zips into one pool and reports each calm
   fact **at most once, aggregated**, in a new `unresolved_internal` key
   — never rep-facing. The Attribution field stopped being restated as
   its own report line too (still consumed to set a Section D toggle).
   **Verified live against both real documents through the running app**
   (real Supabase catalog, not offline fixtures) — both now show exactly
   **one** rep-facing line: the impressions ground-truth line. Also found
   and fixed a gap in the marker list (`"no market is on file"` — the
   real-world Alaska-zip case) while building the test fixtures.
   `tests/test_avails_import_notes.py` rewritten; its old scenarios 1 and
   3 were passing for the wrong reason once this landed.

3. **`697500e` — In-app feedback loop (BACKLOG.md's long-queued item).**
   A "🚩 Report an issue" sidebar popover on every page (category,
   notes, state capture — no screenshot yet, deferred per this session's
   own decision), plus a new "Feedback reports" admin page
   (list/filter/close/export-as-markdown). New `feedback` Supabase table,
   new `db.py` functions, `capture_feedback_state` pulls a curated safe
   subset of session_state (never a `data_editor` delta). Verified through
   the live app: the popover submits and degrades gracefully with a real
   "table doesn't exist yet" error, since **the schema migration hasn't
   been pasted into Supabase yet** — see "What's pending" below.
   `tests/test_feedback.py` added, stubbed against a fake in-memory store.

All three followed CLAUDE.md's screenshot+rep-lens convention for
UI/report changes — every fix was checked in the actual running app, not
just asserted in a test.

## What's pending / needs the user

1. **`! git push`** — three commits sit on local `main` only.
2. **Run the feedback-table migration in Supabase's SQL editor** (schema
   goes before code, DEPLOYMENT.md) — the `feedback` table block at the
   end of `supabase_schema.sql` (search "Stage 11: in-app feedback"). Until
   this runs, the popover and the admin page both show a real "Could not
   find the table 'public.feedback'" error — confirmed live, not a guess.
   After: `Deploy` on Streamlit Cloud, then confirm the live app is
   actually serving the new code (a push succeeding isn't proof of that).

## What's next once picked back up

**Feedback loop fast-follows, deliberately deferred out of this pass**
(see BACKLOG.md's carry-over section for the full reasoning):
- An optional screenshot attachment.
- Loading a report directly into the form like a History entry — the
  captured state is a curated subset (`targeting_groups`, `plan_options`,
  a handful of scalars), not the full `form_json` a rebuild needs, so this
  needs either a richer capture or an explicit partial-restore UI.

Per BACKLOG.md's suggested order otherwise, only one item is left:
**Slide vault** — mechanical, the case-study vault's architecture already
covers it.

## Test suites touched/added this session

`tests/test_intake_area.py` (unaffected, verified still green),
`tests/test_avails_import_notes.py` (rewritten), `tests/test_feedback.py`
(new). Tier 1 (`test_draft_regression.py`, 363 checks) run three times
this session, green every time. Also spot-checked:
`test_avails_pdf_import.py`, `test_avails_pdf_wiring.py`,
`test_avails_placeholder_group.py`, `test_group_scenarios.py`,
`test_form_state.py`, `test_entry_screens.py`, `test_no_stray_magic.py`,
`test_cross_module_refs.py` — all green.

## How to keep this file useful

Whoever (whatever session) is working on this project: **refresh this file
at the end of your session**, before ending the conversation — overwrite
the "Where things stand," "What happened," "What's pending," and "What's
next" sections with the current reality, not an append. If nothing
notable happened (a short Q&A session, no commits), it's fine to leave
this file untouched rather than write a near-empty update.
