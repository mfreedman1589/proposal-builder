# Premion Proposal Builder + Vault — Build Spec v1

Handoff document for Claude Code. Source of truth for phase 1 scope, data model, form logic, and deck assembly. Master deck reference: `TEGNA_MASTER_DECK2.pptx` (123 slides, parsed 2026-07-27).

---

## 1. Overview

A Streamlit app that generates polished, client-personalized Premion proposal decks from a form. Slide library ("the Vault") and audience list live in Supabase. Claude API writes client-facing copy (What You Told Us, strategy narrative). Output is a downloadable .pptx assembled from the master deck plus filled personalization slides.

**Users:** Matt + digital strategist (later: sellers). **Access:** single shared password in Streamlit secrets (same pattern as auto-analyst).

**Stack:** Streamlit (Streamlit Cloud deploy), Supabase (Postgres + Storage), Claude API (Sonnet 4.6 for copy, Haiku 4.5 for utility), python-pptx + raw XML editing for assembly, GitHub source of truth via Claude Code.

---

## 2. Architecture summary

- **Mother-deck assembly:** the master deck lives in Supabase Storage. To build a proposal: download a working copy → delete unselected slides (edit `<p:sldIdLst>` in `ppt/presentation.xml`, then clean orphaned parts) → fill personalization placeholders on the retained template slides → save and serve. Slide *copying* between files is never needed; only deletion + text/image replacement. This preserves brand formatting perfectly.
- **Slide metadata in DB, slide content in the deck.** The Vault DB stores metadata per slide (section, tags, conditions); the actual slide content stays in the master .pptx. Updating the library = uploading a new master deck version + refreshing the slide map.
- **Two net-new template slides** must be added to the master deck before launch (see §7).

---

## 3. Supabase data model

```
slides
  id            serial PK
  deck_version  int          -- FK to deck_versions
  slide_number  int          -- position in that master deck version
  section       text         -- e.g. 'targeting', 'attribution', 'vertical', 'sports_viewership'
  title         text         -- extracted slide title
  condition_key text         -- machine key the form maps to, e.g. 'spanish', 'vertical:healthcare',
                             -- 'sport:nfl_playoffs', 'always', 'quickpitch'
  is_template   bool         -- true if slide has fill-in placeholders
  active        bool

deck_versions
  id, storage_path, uploaded_at, notes, active (only one active)

audiences                    -- canonical catalog: seed from client-facing brochure PDF (repo:
  id, name, category (HH/DEMO/POL/LIFESTAGE/AUTO/ENT/FIN/FOOD/HLTH/LIFESTYLE/RETAIL/SPORTS/TRAVEL/AFIRST),
  subcategory, salesforce_available bool, active
  -- merge in audience_segments_derived.csv (982 non-custom segments parsed from YTD usage, with times_used)

audience_usage               -- real-world planning evidence: audience_usage_ytd.csv
  id, segment_string (may be a combo, comma-separated), delivered_impressions, is_custom
  -- powers "proven combos" in the audience finder; prompt-cache a summarized top-N for Claude calls

proposal_avails              -- structured avails per proposal (renders tables on targeting slide)
  id, proposal_id, audience_label (e.g. 'RETAIL — Banking Intenders'),
  geo_label (market/region row), max_avails int, share_pct numeric (computed),
  share_col_label text ('% of Audience' | 'Budget Share' | custom)

products                     -- drives proposal line defaults
  id, name ('Premion Streaming TV', 'Streaming Retargeting', 'Audience Targeting',
  'Geofencing', 'Site Retargeting — Display', 'Site Retargeting — Pre-roll',
  'Live Sports', 'Broadcast Schedule', ...)
  default_cpm   numeric      -- prefill, editable per proposal ($35 CTV, $5.50 AM, etc.)
  default_desc  text         -- targeting-column boilerplate (e.g. retargeting description)
  line_type     text         -- 'premion' | 'am' | 'broadcast'

proposals                    -- history log
  id, client_name, vertical, market, created_by_note, form_json (full form state),
  generated_at, storage_path (output pptx)

settings
  key/value: shared password hash, gross markup rate (0.15), default market ('DC')
```

---

## 4. Slide map (master deck v1)

| Slides | Section | condition_key |
|---|---|---|
| 1–4 | Title, intro, Why Premion, 8 Pillars | `always` (1 = client title template) |
| 5–8 | Premium Content | `full_deck` |
| 9 | Bilingual-Spanish package | `spanish` |
| 10–11 | Precision Targeting core | `full_deck` |
| 12 | Healthcare targeting (Crossix) | `vertical:healthcare` |
| 13 | First-party data targeting | `first_party` |
| 14 | Streaming TV retargeting | `streaming_retargeting` (also adds proposal line) |
| 15–20 | Attribution core: overview, performance, reporting, web attribution ×2 | `full_deck` |
| 21 | Linear reach extension reporting | `linear_reach_ext` (optional attribution) |
| 22–23 | Brand lift + super case study | `brand_lift` |
| 24 | Sales conversion attribution | `sales_attribution` |
| 25 | Retail case study (BL+WA+SA) | `case_study:retail` |
| 26 | Travel creative performance | `vertical:travel` |
| 27–28 | Experience + expertise, awards | `full_deck` |
| 29 | Vertical Specialties divider | `any_vertical` |
| 30–32 | Education (stats, targeting, case study) | `vertical:education` |
| 33–35 | Healthcare (stats, reach ext, brand lift case) | `vertical:healthcare` |
| 36–37 | Retail | `vertical:retail` |
| 38–40 | Travel (incl. Arrivalist) | `vertical:travel` |
| 41–42 | Home Improvement | `vertical:home_improvement` |
| 43–44 | Banking & Finance | `vertical:banking` |
| 45–46 | Healthcare duplicate | **drop — leftover** |
| 47–48 | Entertainment | `vertical:entertainment` |
| 49–50 | Education duplicate | **drop — leftover** |
| 51–52 | Casual Dining & QSR | `vertical:dining_qsr` |
| 53–55 | Automotive (stats, Polk audiences, Polk Signals) | `vertical:auto` |
| 56–59 | TEGNA Media Group | `tegna_positioning` |
| 60–68 | Audience Marketplace (intro, what/get, retargeting, audience targeting, geofencing, formats, reporting) | `am` + sub-keys `am:retargeting`, `am:audience`, `am:geofencing` |
| 69–73 | Live Sports intro | `sports` |
| 74–94 | Sports viewership slides (per sport) | `sport:<key>` |
| 95–114 | Sports package slides (per sport) | `sport:<key>` (paired with viewership) |
| 115–118 | Total TV (divider, mindshare, WUSA variant, WPMT variant) | `total_tv` + market selects 117 vs 118 |
| 119 | Broadcast TV schedule placeholder | `total_tv` (fix header typo "BROACAST" → "BROADCAST" in v1 cleanup) |
| 120–123 | Proposal section (divider + 3 examples) | examples removed; replaced by generated proposal slide(s) |

**Sports keys (viewership ↔ package pairs):** NFL reg/playoffs/home-team, NBA reg/playoffs, WNBA reg/playoffs, NHL reg/playoffs, MLB reg/playoffs, NCAAF reg/conference/playoffs/home-team, NCAA basketball, March Madness, golf/PGA Majors, motorsports, soccer/professional soccer, World Cup, Prestige Sports, All Live Sports. Selecting a sport inserts both its viewership slide and its package slide. (Map exact pairs during build — a few package slides don't have a 1:1 viewership twin; nearest match or intro-only.)

---

## 5. Input form spec

**Section A — Client basics**
- Client name (text) → title slide, proposal header ("<Client> CTV Strategy" / "<Client> Total TV Strategy")
- Client logo upload (png/jpg) → title slide + proposal slide logo slot
- Market: DC (default) / Harrisburg → station branding (WUSA9 ↔ FOX43/WPMT) on Total TV + broadcast slides; Harrisburg customization track stubbed, empty for now
- Vertical (dropdown: the 9 verticals + None) → gates vertical slides, healthcare targeting, Arrivalist, Polk
- Ad agency involved? (toggle) → gross presentation: display CPMs/budgets × 1.15, labeled Gross
- Spanish-language campaign? (toggle) → slide 9
- Discovery notes (free text area) → Claude copy input

**Section B — Deck scope**
- Preset: Full Proposal / Quick Pitch (3–5 slides)
  - Quick Pitch default: client title → What You Told Us → Why Premion → one targeting-or-vertical slide → proposal slide
- TEGNA media positioning (toggle) → 56–59
- Vertical slides auto-on when vertical selected (overridable checkboxes)

**Section C — Products** (each selected product seeds a proposal line with default CPM + boilerplate)
- Premion Streaming TV (default on)
- Streaming Retargeting (toggle) → slide 14 + line
- Audience Marketplace (toggle) → 60–63, 67–68 + sub-checkboxes:
  - Audience Targeting → 65 + line
  - Geofencing → 66 + line
  - Site Retargeting — Display and/or Pre-roll → 64 + line(s)
- Live Sports (toggle) → intro 69–73 + sport multi-select pane → paired slides + line(s)
- Total TV (toggle) → 115–116 + market slide (117/118) + schedule placeholder 119 + broadcast line

**Section D — Targeting & attribution**
- Healthcare targeting: auto when vertical = healthcare
- First-party data (toggle)
- Standard (pre-checked, informational): dashboard, reporting, web attribution
- Linear reach extension (toggle, only if Total TV) → slide 21
- Arrivalist: auto for travel; Polk: auto for auto
- Sales attribution (toggle) → slide 24 + included-list item
- Brand lift (toggle) → 22–23 + included-list item
- Commercial production (toggle) → included-list item only

**Section D2 — Audiences & avails** (structured, replaces free-text avails)
- Add pulled audience(s): audience label + optional grouping (e.g. Retail vs Commercial tracks, per BRB slide 7)
- Per audience: geo rows (market/region label + max monthly avails); app computes share % and totals row; share column label selectable ('% of Audience' / 'Budget Share')
- Renders as table(s) on the vertical Precision Targeting slide (modeled on Blue Ridge slide 7); audience names also flow to Campaign Specs Audience section and media plan Targeting column
- Audience finder assist (phase 3): Claude recommends segments/stacks from the catalog + YTD usage evidence

**Data-flow principle (single source of truth):** the form is the master record; the Campaign Specs slide and the media plan table are both renders of the same fields. Flight dates → Timing section + Flight column; geo → Geography section + Geo column; audience → Audience section + Targeting column; budget → Budget & Allocation + cost math. Enter once, always in sync.

**Section E — Proposal / media plan** (supports 1–3 plan options per proposal — Option A/B/C tabs, per the Blue Ridge "Option 1/2/3" pattern; each option has its own line set + budget and generates its own media plan slide)

Per option, repeatable rows seeded from Section C:
- Per row: Tactic (from product) / Flight dates / Geo / Targeting text / Impressions / CPM (prefilled, editable) → Cost auto-computed
- Display mode: Monthly amounts + grand flight total (default) / Flight totals per line
- Mixed Premion + AM products: one combined Premion-style table (columns: Tactic, Flight, Geo, Targeting, Monthly Impressions, Monthly Cost)
- Gross toggle inherits from agency flag; totals row computes blended CPM
- "Included with Campaign" list auto-assembles: Dedicated Account Management Team + Monthly Reporting Calls/Optimizations + Dashboard Access (always); Web Attribution (Pixel Required) standard; Commercial Production / Sales Attribution (CRM Upload Required) / Brand Lift Study per toggles
- Terms block: Premion T&Cs (Premion/mixed) vs TEGNA terms (AM-only) — auto
- Approval + date signature line

**Smart-defaults principle:** vertical dropdown and product toggles drive suggestions; everything auto-selected stays overridable. A seller should mostly confirm, not decide.

---

## 6. Claude API integration

- **What You Told Us copy + strategy narrative:** Sonnet 4.6. Input: discovery notes + selected products/audiences + vertical. Few-shot with 1–2 sample proposals (add to repo when Matt provides). Output goes straight into the deck — no in-app review step; users edit in PowerPoint afterward.
- **Audience recommendation (phase 3):** brief + full audience table → ranked fit suggestions with one-line rationale. Prompt-cache the audience catalog.
- **Vault auto-tagging (phase 2):** Haiku 4.5 suggests section/tags for newly uploaded slides from extracted text.
- API key in Streamlit secrets (never shell env). Budget: prepaid credits; expected cost ≈ $0.03–0.05/proposal.

---

## 7. Net-new template slides (add to master deck before launch)

1. **Client title slide** — variant of slide 1 with client logo placeholder + "<Client> | <Product> Strategy" title.
2. **What You Told Us / Campaign Specs** — modeled on the existing Blue Ridge Bank "Campaign Specs" slide. Split layout: left half full-bleed image with eyebrow "PREMION + <VERTICAL>", "<Client>" + "Campaign Specs" title, © Premion footer; right half periwinkle-headed sections with dark bullets, drawn from a superset of eight: **Goals & Approach / Audience / Geography / Budget & Allocation / Placements & Creative / Timing / Measurement & Success / Situation & Challenge (optional, off by default)**. Render only sections with content, cap ~6–7 for layout. PREMION logo bottom right. Fill strategy: hybrid — Geography, Timing (flight dates), Budget, and Measurement & Success (auto-written from attribution toggles: web attribution, sales attribution, brand lift, linear reach ext) pull directly from form fields; Goals & Approach, Audience, Placements & Creative, and Situation & Challenge are Claude-written from discovery notes + selected products (structured JSON output, one array of short bullets per section, 2–3 bullets each). Keep Goals & Approach about client objectives — measurement commitments belong in Measurement & Success. Left image: per-vertical stock image stored in the Vault (image placeholder swap), fallback to a default Premion image.
3. **Generated proposal slide** — parameterized version of slide 121/122 layout (table + included list + signature + terms). Built programmatically each run; keep a styled template slide with the table shell to preserve formatting.
4. **Personalized Precision Targeting / avails slide** — vertical targeting slide variant with avails table shell(s): audience group header row, geo rows (Market | Max Avails | share %), totals row. Modeled on Blue Ridge slide 7; populated from proposal_avails.

Personalization fill rules: use python-pptx `run.text` assignment (never `text_frame.text` — collapses formatting); logo via image placeholder swap; table rows built by cloning a styled template row's XML.

---

## 8. Vault admin (phase 2)

- Upload new master deck version → parse (markitdown), diff slide list vs current map, flag new/moved slides for tagging (Haiku suggests), activate version.
- Slide browser with thumbnails (LibreOffice render on upload) for the strategist to browse the library.
- Audience CSV upload → audiences table refresh.
- Old versions kept inactive so past proposals stay reproducible.

---

## 9. Phasing

- **Phase 1 (build now):** form (A–E) → assembly engine → filled personalization slides → .pptx download; proposal history log; shared password. Master deck v1 cleanup: drop duplicate slides 45–46/49–50, fix "BROACAST" typo, add the 3 template slides, and add two slides found in real proposals but missing from the master: Dynamic Video Ad (condition_key `dynamic_creative`, product toggle in Section C) and the Discover→Recommend process slide (`full_deck` or optional). Master deck stays the single unit of versioning — no supplemental slide files in phase 1.
- **Phase 2:** Vault admin page, versioning, thumbnails, auto-tagging, sample-proposal few-shot library.
- **Phase 3:** audience intelligence — seed audiences table from brochure PDF + audience_segments_derived.csv; load audience_usage_ytd.csv; audience finder (Claude recommendations citing proven stacks); Harrisburg customization track content.
- **Acceptance test:** rebuild a recent real DC proposal side-by-side with the handmade one; gap list drives phase 1 polish before strategist rollout.

## 10. Open items

- Exact sports viewership↔package pairing during build
- Brochure PDF → audiences table extraction (phase 3 build task; PDF in repo)
- BRB slides to fold into master v1: Dynamic Video Ad, Discover→Recommend process, avails targeting template
- Harrisburg rate defaults (if different from DC) when that track activates
