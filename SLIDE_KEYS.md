# Slide keys — the tagging reference

Every slide in the master deck carries a **key** that decides *when that slide
gets included in a generated proposal*. The key lives in the slide's own
speaker notes, on its own line:

```
key: vertical:healthcare
```

That's the whole mechanism. The proposal builder reads the key, looks at what
the seller ticked on the form, and keeps the slide or drops it.

---

## How to tag a slide

1. Open the master deck in PowerPoint.
2. **View → Notes Page** (or open the notes pane under the slide).
3. Add or edit the line `key: something` — on its own line, nothing else on it.
4. Save.
5. Ask for the deck to be re-scanned (or upload it on the **Update master
   deck** page, which scans it for you and refuses to activate a deck with any
   slide it can't resolve).

Notes:

- Capitalisation and spacing don't matter — `Key:  vertical:auto` is fine.
- Anything else in the notes is left alone. The key line is the only line the
  tool touches.
- If there are two `key:` lines, **the last one wins** — so you can leave the
  old one above and add a correction below it.
- A slide with no key falls back to guessing from its text, which usually
  works but is exactly what breaks when someone rewrites a heading. **Tag
  every slide.**

`master_deck_slide_keys.md` lists what every slide is currently tagged as.

---

## The rule of thumb

> A key answers: **"what has to be true about the proposal for this slide to
> belong in it?"**

If the answer is "nothing, it's always in" → `always`.
If it's "the client is a hospital" → `vertical:healthcare`.
If it's "they're buying NFL playoffs" → `sport:nfl_playoffs`.

One key per slide. If a slide genuinely serves two purposes, it's usually two
slides.

---

## Always included

These go in every proposal, in every preset. Use them only for the handful of
slides that are genuinely universal.

| Key | What it's for |
|---|---|
| `always` | Boilerplate that belongs in every deck — the "who is Premion" opener, the advantage/pillars slide. |
| `client_title` | **The cover slide.** Carries `{{CLIENT_NAME}}` / `{{PROPOSAL_TITLE}}`. There should only ever be one. |
| `campaign_specs` | **The "Campaign Specs" slide** the form fills in with goals, audience, geography, budget, placements and timing. Only one. |
| `proposal_template` | **The media plan slide.** Carries the plan table and `{{PLAN_TITLE}}`. Only one — the builder copies it when a proposal has more than one plan option. |

⚠️ The three template slides above are *structural*. Don't tag anything else
with them, and don't remove them from the deck.

---

## Deck scope — how much detail the seller wants

The form offers three presets: **Quick Pitch**, **Standard** (the default) and
**Extended**.

| Key | When the slide appears |
|---|---|
| `full_deck` | **Extended only.** Supporting/background content — nice to have, not needed in a short pitch. This is the right key for most general Premion content slides. |
| `core:premium_content` | Standard + Extended. The single best "here's the premium content we run on" slide. |
| `core:attribution_overview` | Standard + Extended. The one-slide overview of attribution and measurement options. |
| `core:reporting` | Standard + Extended. The reporting/transparency slide. |
| `core:web_attribution` | Standard + Extended. The website-attribution explainer. |

The four `core:` keys are deliberately **one slide each** — they're the "if we
only show four content slides, show these" set. If you want to promote a
different slide into Standard, move the key onto it (and take it off the old
one), rather than tagging two slides with the same `core:` key.

---

## Verticals — client industry

Used when the seller picks a vertical on the form.

| Key | When it appears |
|---|---|
| `vertical:<name>` | Any slide specific to that industry — its stats, case studies, targeting partners. Included whenever that vertical is selected. |
| `vertical:<name>:targeting` | **Only** the vertical's own static "Precision Targeting" page. It's dropped automatically whenever the personalised avails table is used instead, so these two never both appear. |

Valid vertical names:

`education` · `healthcare` · `retail` · `travel` · `home_improvement` ·
`banking` · `entertainment` · `dining_qsr` · `auto` · `legal`

Examples: `vertical:auto`, `vertical:healthcare:targeting`, `vertical:legal`

ℹ️ **`legal` has no slides in the master deck yet.** The vertical is fully
wired up everywhere else — it's in the form dropdown, the AI drafting schema,
the case study tags and the audience finder — so a legal proposal generates
correctly today, just without any legal-specific slides in it. When legal
slides are added, tag them `key: vertical:legal` (and the vertical's own
Precision Targeting page, if it gets one, `key: vertical:legal:targeting`) and
they'll start appearing with no code change.

---

## Audiences and targeting

| Key | When it appears |
|---|---|
| `targeting_avails_template` | **The personalised targeting / avails table** — the one the form fills with real audience segments and avails numbers. Only one slide. Included when the seller turns the avails table on (and always under Standard). |
| `first_party` | The first-party / CRM data targeting slide. Appears when "First-party data targeting" is ticked. |
| `streaming_retargeting` | Streaming TV Retargeting explainer. Appears when that product is selected. |

---

## Audience Marketplace

| Key | When it appears |
|---|---|
| `am` | General Audience Marketplace content — appears whenever Audience Marketplace is selected at all. |
| `am:audience` | Only when the **Audience Targeting** format is selected. |
| `am:geofencing` | Only when **Geofencing** is selected. |
| `am:retargeting` | Only when **Site Retargeting** is selected. |

---

## Total TV / broadcast

| Key | When it appears |
|---|---|
| `total_tv` | Any Total TV content — appears when Total TV is selected. |
| `total_tv:dc` | Washington DC market only (the WUSA slide). |
| `total_tv:harrisburg` | Harrisburg market only (the WPMT slide). |
| `linear_reach_ext` | The linear reach-extension slide. Needs both Total TV **and** "Linear reach extension" ticked. |

---

## Live Sports

Two families, and it matters which one you use:

| Key | What it is |
|---|---|
| `sports` | General Live Sports content — appears whenever any sports package is selected. |
| `sports_viewership_intro` | The one generic "live sports viewers" overview slide. |
| `sport:<package>` | **A sellable package slide** — the one describing what the client actually buys. Also carries that package's rate. |
| `sport_viewership:<package>` | **The audience/viewership stats slide** for that package. Only shown when that specific package is selected. |

Valid package names (same list for both):

| | | |
|---|---|---|
| `nfl_reg` | `nfl_playoffs` | `nfl_home_team` |
| `nba_reg` | `nba_playoffs` | |
| `wnba_reg` | `wnba_playoffs` | |
| `nhl_reg` | `nhl_playoffs` | |
| `mlb_reg` | `mlb_playoffs` | |
| `ncaaf_reg` | `ncaaf_conference` | `ncaaf_playoffs` |
| `ncaaf_home_team` | `ncaa_basketball` | |
| `golf` | `pga_majors` | `soccer_pro` |
| `prestige_sports` | `all_live_sports` | |

Examples: `sport:nfl_playoffs`, `sport_viewership:mlb_reg`

⚠️ **A new sports package needs more than a tag.** Adding
`sport:motorsports` to a slide won't make it selectable — the package also has
to be added to the form's sports list and the rate card. Ask for that as a
code change.

---

## Attribution and measurement

| Key | When it appears |
|---|---|
| `sales_attribution` | Sales-attribution slides. Appears when "Sales attribution" is ticked. |
| `brand_lift` | Brand-lift slides. Appears when "Brand lift" is ticked. |

---

## Other

| Key | When it appears |
|---|---|
| `spanish` | Spanish-language / bilingual content. Appears when "Spanish-language campaign" is ticked. |
| `tegna_positioning` | TEGNA-as-a-media-company positioning. Appears when that toggle is on (**off** by default). |
| `section_divider` | **A section title card** ("Live Sports", "Proposal Slides", "Premium Content"…). These are **never** included in a generated proposal — tag dividers with this so they're dropped cleanly. |

---

## Keys that currently do nothing

Tagging a slide with one of these means **it will never appear in any
proposal**. Listed so you don't reach for one by accident, and so the two
genuine oddities are visible.

| Key | Status |
|---|---|
| `section_divider` | Intentional — dividers are always dropped. |
| `sports_viewership_unmapped` | Intentional. Three viewership slides (Motorsports, World Cup, March Madness) have no matching sellable package, so they can't be selected. If a package is ever added for one, retag it `sport_viewership:<package>`. |
| `dynamic_creative` | ⚠️ **Orphaned.** The "What is a Dynamic Video Ad?" slide. There's no form control that turns it on, so it never ships. Either it needs a toggle added, or it should be retagged (`full_deck` would put it in Extended decks). |
| `case_study:retail` | ⚠️ **Orphaned.** A built-in retail case study slide, from before the case study vault existed. Same situation — no form control. The vault is the better home for case studies now. |

## Total TV market variants

Seven keys, all driven by the **Total TV toggle alone** — there is no separate
control for any of them. With Total TV on, the market's variants replace the
standard slides; with it off, none of them appear and the deck is exactly what
it was before.

| Key | Slide | Selected when |
|---|---|---|
| `client_title_cobrand:dc` | WUSA9 + Premion cover | Total TV on, market DC — **replaces** `client_title` |
| `client_title_cobrand:harrisburg` | FOX43 + Premion cover | Total TV on, market Harrisburg — **replaces** `client_title` |
| `total_tv:harrisburg` | WPMT + Premion pitch slide | Total TV on, market Harrisburg |
| `proposal_template_total_tv:dc` | Media plan, WUSA9 branding | Total TV on, market DC — **replaces** `proposal_template` |
| `proposal_template_total_tv:harrisburg` | Media plan, FOX43 branding | Total TV on, market Harrisburg — **replaces** `proposal_template` |
| `broadcast_schedule_template:dc` | WUSA9 schedule grid | Total TV on, DC, **and** a Wide Orbit schedule imported |
| `broadcast_schedule_template:harrisburg` | FOX43 schedule grid | Total TV on, Harrisburg, **and** a schedule imported |

Three things worth knowing before editing these:

**Three of them resolve from their notes label only.** The co-brand covers and
the two Total TV plan templates are text-identical to the standard slides they
replace (and to each other across markets) — the only difference is a logo
image. There is no text anchor that can tell them apart, so if a merge loses
speaker notes they silently collapse back onto `client_title` /
`proposal_template` and the deck quietly ships the wrong cover. The broadcast
schedule templates *are* anchorable, since they name their station.

**Replacement is enforced in `build_presentation`, not by omitting keys.** The
standard cover and plan template are swept out whenever a variant is present,
so it holds however a slide got selected. Two covers in one deck would be
visible to the client, and `personalize` finds the plan slide by token — a
second one would silently take the fill.

**Both markets' variants are swept unconditionally.** Even if selection were
wrong upstream, a slide branded for the other market can't survive into the
deck. A competitor station's logo in front of a client is the worst thing this
deck can do.

There are also a few keys the app can switch on that no slide currently uses
(`any_vertical`, `proposal_divider`, and viewership keys for the three
packages that have no stats slide). Harmless — they simply match nothing.

---

## Adding a genuinely new kind of slide

If a new slide doesn't fit any key above — a new product, a new market, a new
attribution option — tagging alone isn't enough. The key has to be *switched
on* by something on the form, and that's a code change. Tag it with your
intended key, say what should turn it on, and it can be wired up.

Until then, `full_deck` is the safe parking spot: the slide ships in Extended
decks and nowhere else.

---

## How the key is read, and how the update page gates on it

Moved verbatim from `CLAUDE.md`. This is the mechanism behind the workflow above.

- **A slide's `key:` label in its speaker notes beats the text anchors.** Every slide in the tagged master carries a line like `key: vertical:healthcare` in its own speaker notes (invisible in presentation, written by `tag_deck_keys.py`). `slide_map.build_slide_map_from_prs` reads that line first and only falls back to anchors for an untagged slide, so a half-tagged deck — a new slide dropped into a tagged master — still resolves throughout. Anchors are still *run* on every slide even when a label is present, because the carry-forward section default has to keep tracking for any untagged slide further down. Matching is case-insensitive and whitespace-tolerant, and the **last** `key:` line wins, so appending a correction below an old one does the obvious thing.

  **To correct a mislabelled slide:** open the deck in PowerPoint, View → Notes Page (or the notes pane), edit that slide's `key:` line to the right condition_key, save, then run `python slide_map.py <deck>` and confirm zero `WARNING` output. `master_deck_slide_keys.md` is a generated slide → title → key reference for eyeballing the whole set — a convenience copy, **not** the source of truth; the notes are. Regenerate it (and re-tag a new deck) with `python tag_deck_keys.py <deck> --report master_deck_slide_keys.md`, which never writes in place, appends the key line as its own paragraph rather than replacing existing notes, and is idempotent (re-running replaces its own line instead of stacking duplicates).

  Why this matters: anchors infer a slide's role from its prose, which is exactly what changes when someone retitles a slide or rewrites a heading — and a wrong inference mis-assembles decks silently. Verified by wiping the text off six slides in the tagged deck: every key survived, where anchor-only resolution would have lost all six. Three slides are *deliberately* unmapped (`sports_viewership_unmapped` — Motorsports, World Cup, March Madness viewership), since they have no package counterpart and are never selectable.
- **Updating the master deck goes through the "Update master deck" page, and the gate is the point of it.** Upload a `.pptx` → it's scanned and diffed against the active version → if everything resolves you see what changed and can activate; if anything doesn't, activation is blocked and the offending slides are listed by number and heading. Activation optimizes, uploads to the `decks` bucket, inserts a `deck_versions` row, flips `active` and clears the deck cache. **Old versions are never touched** — storage keeps them and `deck_versions` keeps their rows, so rolling back is just re-activating one. The diff matches slides **by text, not position**, since position is precisely what a deck update churns; a slide that kept its text but moved is reported as "moved" (harmless — nothing depends on slide numbering any more).

  **What "unresolved" means, and why it isn't `key is None`.** Anchor classification has a carry-forward default: a slide matching no rule inherits whatever section it sits in, so a key is essentially *never* None and a naive gate would never fire. `_classify_all` therefore reports a **source** per slide — `notes` (an explicit `key:` label), `anchor` (a rule matched it directly), or `carry` (neither, it inherited). The gate blocks on `carry`, because that's a guess dressed as an answer: a newly added slide would otherwise be assembled into decks under whatever key its neighbour happened to have. Demonstrated in testing — an untagged slide inserted at position 50 would have silently inherited `any_vertical`. This is also why tagging the master mattered: the untagged deck has **11 slides that only resolve by carry-forward** and would fail its own gate, while the tagged one is 119/119 from notes.
