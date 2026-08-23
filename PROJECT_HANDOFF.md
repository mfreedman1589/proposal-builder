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

**Local `main` is 6 commits ahead of `origin/main`, not yet pushed.**
Working tree is clean. The user has not yet said "push" for this batch —
check with them before running it if you're picking this up cold.

```
618830e Verify: the _placeholder marker is backward-compatible with old proposals
16a51c2 Fix: avails-import notes named the document, not the group, and dumped zips one line each
08d2d19 Fix: map colors -- orange and red were too close to tell apart
b029a65 Fix: a persistent blank "no target market picked" placeholder group
7dc857e Audit: harden test_coviewing.py against the substring-presence weakness
5d89a5f Co-viewing coefficient: an opt-in projection of person-level exposure
```

(above `origin/main`'s tip: `925b577 Correct % of avails to per-line
matching, not a deck-wide aggregate`)

## What happened this session

1. **Co-viewing coefficient shipped** (BACKLOG.md item) — an opt-in
   `show_coviewing` checkbox adds a Monthly Coviewing column to Premion
   Streaming TV / Live Sports lines, plus a short citation on the plan
   slide's Terms & Conditions box. Multiplier (1.4x) and citation are a
   Supabase `app_settings` row (`key='coviewing'`), same fallback pattern
   as the rate card. Full detail: `CLAUDE.md`'s media-plan-section bullet.

2. **Two audits, requested after co-viewing landed:**
   - Other places inferring optional-column presence by counting table
     cells (the way `add_full_flight_total_row` did before this session's
     fix) — **none found elsewhere.**
   - Test assertions vulnerable to the "X" in "X X" substring-presence
     weakness that let a real doubled-label bug through — found and fixed
     in `tests/test_coviewing.py` itself (now exact-match, not substring).

3. **Three live bugs reported by the user and fixed, all found from real
   proposal-building use:**
   - A blank "no target market picked" placeholder group that survived
     forever in the D2 avails table, even after a real avails PDF import
     or Audience-finder add. Fixed with a `_placeholder` marker; **verified
     backward-compatible** with proposals saved before the marker existed
     (see `tests/test_avails_placeholder_group.py`'s last scenario for the
     exact mechanics and the one honest limitation: an old proposal's
     pre-existing blank row is never retroactively recognized).
   - Map colors — orange/red were too close to tell apart. Swapped for a
     bolder, more saturated pair (tab10's `#FF7F0E`/`#D62728`). Verified
     with a real dark-mode render (see below) — sent to the user, awaiting
     their confirmation it looks right.
   - Avails-import notes named the whole document instead of the specific
     group, and a named-zip group's unresolved zips were dumped one line
     each instead of a calm summary — a real 60-zip document turned into
     what would have been 120+ report lines, now 19.

4. **Dark-mode render produced and sent** (3 PNGs, via `SendUserFile`) for
   the new palette: full slide at true size, a legend close-up showing
   the new orange/red plus the two-way hatch next to the three-way neutral
   fill, and a close-up of the orange/red dots side by side. **This was
   the explicit last blocker before push** — check the conversation for
   whether the user has since confirmed it.

## What's pending / needs the user

- **Confirm the render** (just sent, not yet acknowledged as of this
  writing).
- **Say "push"** when ready — nothing has been pushed this session past
  `925b577`.

## What's next once this is pushed

Per `BACKLOG.md`'s "Suggested order" (co-viewing now shipped, so it drops
off this list):
1. Audience usage ingestion — **Matt's own step**
   (`setup_supabase.py audience_usage_bucket`, then upload/activate the
   real workbook through the admin page), not a coding task.
2. Feedback loop (in-app "Report an issue") — see `BACKLOG.md` for the
   full design, already settled.
3. Slide vault — mechanical, the case-study vault's architecture already
   covers it.

## Test suites touched/added this session

`tests/test_coviewing.py` (new), `tests/test_avails_placeholder_group.py`
(new), `tests/test_avails_import_notes.py` (new). Full tier-1 suite
(`tests/test_draft_regression.py`, 363 checks), `test_form_state.py` (69),
`test_share_of_voice.py`, `test_group_merge_split.py`,
`test_color_lifecycle.py`, `test_targeting_map.py`, `test_group_scenarios.py`,
`test_group_backward_compat.py`, and `test_avails_pdf_wiring.py` are all
green as of the last commit above.

## How to keep this file useful

Whoever (whatever session) is working on this project: **refresh this file
at the end of your session**, before ending the conversation — overwrite
the "Where things stand," "What happened," "What's pending," and "What's
next" sections with the current reality, not an append. If nothing
notable happened (a short Q&A session, no commits), it's fine to leave
this file untouched rather than write a near-empty update.
