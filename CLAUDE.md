# Premion Proposal Builder

A Streamlit app that assembles polished, client-personalized Premion CTV/OTT proposal decks from a form, by deleting unselected slides from a master `.pptx` and filling the retained template slides. Phase 1 (current): form + assembly engine + local download, no database or LLM copy generation yet.

**Stack:** Streamlit (form + UI), python-pptx (deck assembly/fill), Supabase (planned, not built), Claude API (planned, not built).

**Source of truth:** `proposal_builder_build_spec.md` is the full spec (data model, form fields, slide map, phasing). Read it before making non-trivial changes. `master_v1_1_changelog.md` documents the current master deck version's token conventions and what changed from v1.

## Key conventions

- **Token fill convention:** template slides carry `{{TOKEN}}` placeholders in run text. Fill by finding the run containing the token and doing `run.text = run.text.replace(placeholder, value)`. **Never assign `text_frame.text`** — it collapses all run-level formatting (font, color, size) on that paragraph. Multi-item content (bullet lists, table rows) is filled by deep-copying the tokenized paragraph/row XML element and inserting clones as siblings, then filling each clone the same way. See `fill_bullet_list` / `fill_table_rows` in `assembly.py`.
- **Slide map is always derived by scanning, never hardcoded.** `slide_map.py` opens the master deck and classifies every slide into a `condition_key` (e.g. `vertical:healthcare`, `sport:nfl_playoffs`, `full_deck`) by matching text anchors, with a small carry-forward default for undifferentiated divider slides. If the master deck is replaced with a new version, do **not** hand-patch slide numbers — update the anchor rules in `slide_map.py` if a new slide type needs one, then let it re-scan. This was a direct, hard-learned lesson from the v1 → v1.1 deck migration (slide count and positions shifted; a stale hardcoded map would have silently mis-assembled decks).
- Slide deletion works by removing the `<p:sldId>` from `presentation.xml`'s `sldIdLst` and dropping the corresponding relationship (`prs.part.drop_rel`) — python-pptx then automatically omits any part unreachable from the package's relationship graph on save, so orphaned slide parts clean themselves up.
- `.pptx` and `.pdf` files are gitignored (master deck, generated test output, source brochure) — never commit them. Secrets live in `.streamlit/secrets.toml` (gitignored) — currently just `APP_PASSWORD`, the shared-password gate.

## Running things

- **Assembly engine standalone:** `python assembly.py` — runs the hardcoded `SELECTIONS`/`FILL_DATA` at the bottom of the file against `TEGNA_MASTER_DECK_v1_1.pptx` and writes `test.pptx`.
- **Full app:** `streamlit run app.py`, then log in with the password in `.streamlit/secrets.toml`.
- Python and `streamlit`/`pandas`/`python-pptx`/`Pillow` are installed locally (there was no system Python before this project — installed via winget). `git push` needs an interactive login (Git Credential Manager) that the sandboxed shell can't do; the user runs it via `! git push` themselves.

## Status

**Done:** git repo + GitHub remote; standalone assembly engine (`assembly.py`) with scan-based `slide_map.py`; personalization fill (client title, Campaign Specs bullets, targeting/avails table, media plan table, logo swap) all via the token convention; Streamlit form (`app.py`) covering spec sections A–E, wired to `assembly.py`, with smart defaults (vertical → healthcare/Arrivalist/Polk info, agency → gross markup, Live Sports → sports picker), password gate, and a working Generate → download flow. Verified against master deck v1.1 (118 slides).

**Not started:** Supabase (Vault slide metadata, audiences, proposal history log); Claude API copy generation (Campaign Specs bullets are manually typed for now); multiple plan options (Option A/B/C — only Option A exists); the exact sports viewership↔package slide pairing (still a coarse block, per spec §10); Vault admin page (phase 2); audience intelligence (phase 3).
