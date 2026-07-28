# Master Deck v1.1 — Changelog \& Token Reference

File: `TEGNA\_MASTER\_DECK\_v1\_1.pptx` (122 slides; replaces TEGNA MASTER DECK2.pptx as the assembly source).

## What changed from v1 (123 slides)

**Added (5 new slides):**

* **Slide 1 — Client title template.** Duplicate of the cover with `{{CLIENT\_NAME}}` (title placeholder), `{{PROPOSAL\_TITLE}}` (tagline text box inside the group), and a marker picture shape named `CLIENT\_LOGO` top-right (swap its image, keep position/size).
* **Slide 2 — Campaign Specs template** (imported from Blue Ridge deck). Tokens: `{{CLIENT\_NAME}}`, `{{VERTICAL}}` in the left-panel group; one token bullet under each right-side header: `{{GOALS\_BULLETS}}`, `{{AUDIENCE\_BULLETS}}`, `{{GEOGRAPHY\_BULLETS}}`, `{{BUDGET\_BULLETS}}`, `{{PLACEMENTS\_BULLETS}}`, `{{TIMING\_BULLETS}}`. Fill rule: replace the token run with the first bullet's text; for additional bullets, deep-copy that paragraph's `<a:p>` and insert as siblings (preserves formatting). Section headers are plain paragraphs — to drop an empty section, remove its header paragraph + token paragraph + trailing spacer.
* **Slide 13 — Personalized targeting / avails template** (imported from Blue Ridge slide 7). Header/description carry `{{VERTICAL}}` / `{{VERTICAL\_LC}}`. Table rows: \[0] `{{AUDIENCE\_GROUP}}`, \[1] column headers with `{{SHARE\_LABEL}}`, \[2] data row `{{GEO}} | {{AVAILS}} | {{SHARE}}`, \[3] `{{TOTAL\_LABEL}} | {{TOTAL\_AVAILS}} | 100.0%`. Fill rule: clone row 2's `<a:tr>` per geo; for a second audience group, clone the whole 4-row block.
* **Slide 17 — Dynamic Video Ad** (imported; no tokens; condition\_key `dynamic\_creative`, sits right after streaming retargeting).

**Modified:**

* **Slide 122 — proposal/media plan template** (was example slide 121, tokenized in place): title `{{CLIENT\_NAME}} {{PLAN\_TITLE}}`; table = header row + one token data row (`{{TACTIC}} {{FLIGHT}} {{GEO}} {{TARGETING}} {{IMPRESSIONS}} {{COST}}`) + totals row (`{{TOTALS\_LABEL}} … {{TOTAL\_IMPRESSIONS}} {{TOTAL\_COST}}`); included list collapsed to `{{INCLUDED\_LIST}}` (fill as one paragraph per item via paragraph cloning). Approval line, terms block, and MEDIA PLAN header untouched.
* "BROACAST" typo fixed on the broadcast schedule slide (now slide 116... verify by text search, not number).

**Deleted:** duplicate vertical sets (old 45–46 healthcare dup, old 49–50 education dup) and old example proposal slides 122–123. Old 121 became the template (now 122). Verified: healthcare/education stats each appear exactly once.

## Required update to assembly.py

The v1 slide numbering in spec §4 is now stale. Don't hand-patch offsets — **regenerate SLIDE\_MAP by scanning v1.1** (extract per-slide text, match section anchors) or apply: +2 after the two title-area inserts, +3 after slide 14, +4 after slide 18, then account for the four deletions in the old 45–50 range and the tail changes. The scan approach is safer and future-proofs deck updates — recommend building it as a small `slide\_map.py` utility now, since the Vault versioning flow (spec §8) needs exactly that diff logic later.

## Fill convention (applies to all templates)

Find runs whose text contains a `{{TOKEN}}`, replace via `run.text` assignment. Never assign `text\_frame.text`. Multi-item content (bullets, table rows, included list) = clone the tokenized paragraph/row element, then fill each clone.

## Verification done

Zip integrity clean; reopened with python-pptx; token audit passed (tokens on slides 2, 3, 14, 120, 122 only); LibreOffice render of all six new/modified slides visually confirmed (branding, layout, table styling intact). Still recommend one PowerPoint open-and-flip before first real use.

## Still open

* Sports viewership slides (now \~78–98) remain a coarse block — per-sport pairing still the open item from spec §10.
* AM-format media plan table (Tactic/Deliverable/…/Rate, net flight totals) has no template slide yet — future work for AM-only proposals.

