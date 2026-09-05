# Attribution Module — Framework for Review

The next build phase connects campaign reporting to the Proposal Builder. Below is the working framework. Flag anything that's wrong, missing, or worth changing.

---

## 1. The Loop We're Building

Proposal closes → campaign launches → monthly reports → optimizations ramp in → wrap-up report → next proposal.

The goal is one connected story. The wrap-up doesn't just report results — it feeds directly into the follow-up proposal, showing how what we learned shapes the next phase.

---

## 2. Data In

**Source:** raw Excel from the live dashboard, not the templated PowerPoint from the central team.

**Why:** the app rebuilds every chart and table from scratch. Two benefits:

- Charts land in our own template/color scheme, matching the proposal
- All internal nomenclature (RFP numbers, Salesforce codes, inside language) gets stripped out automatically, because we never carry it over

**Uploads:** attribution data and delivery data live on different dashboard tabs, so they come in as separate files. Same for other tactics.

---

## 3. What Goes In a Report

**Setup** — brief reminder of the audiences we're targeting and the goals set at proposal time. Pulls from the proposal.

**Delivery summary** — impressions delivered, delivery by creative, top publishers. Kept intentionally short. Clients care more about attribution than delivery, and we've historically over-weighted this section.

**Attribution (the crux)** — CTV, including CTV retargeting.

- Top line: unique visitors and unique visitor rate
- All data cuts: attributed rate
- 15-day attribution window throughout

**Supporting tactics** — OTT retargeting (clicks back to site, top creative — usually display, sometimes retargeted preroll). Audience marketplace tactics (geofencing, programmatic, audience targeting) integrate alongside when used.

**URL report** — unique visits by page. This is the richest analysis opportunity. Which pages people hit tells us intent: program pages vs. enrollment pages for a university, events vs. booking pages for tourism. Tied back to the original strategy.

---

## 4. Optimization Philosophy

**Optimize by subtraction.** We remove what's clearly failing so those impressions flow to what's already working. We don't reshuffle.

**Only act on outliers.** A creative at 0.45 vs. one at 0.50 isn't worth touching. 0.2 vs. 0.6 is.

**Dimensions we analyze:**

| Dimension | Signal | Notes |
|---|---|---|
| Creative | Wide gap in attributed rate | Off-limits if client mandates a fixed split |
| ZIP code | High impressions + low attributed rate | Primary optimization lever |
| Day of week | One day far below the rest | e.g. Wednesday at 0.2 vs. ~1% elsewhere |
| Market | Low attributed rate vs. impressions received | Often protected; toggle-able |
| Publisher | Wide gap | Usually avoided — too seasonal |

---

## 5. Timing

We wait for a trend, not a month.

- **Month 1** — report only. No optimizations.
- **Month 2** — most significant outliers only, and only on high-intensity campaigns.
- **Month 3** — real optimization point. Three months is a genuine trend, using campaign-start-to-date data.

The reasoning: Wednesdays can be terrible in month 1 and great in month 2. Publishers swing seasonally — a sports network looks weak until the playoffs start.

---

## 6. Controls (Per Campaign)

**Optimization level:** Low / Moderate / High

**Eligible dimensions:** checkboxes, all on by default. Uncheck anything off the table — mandated creative splits, protected markets.

**Caps** (first-pass numbers, open to adjustment):

- High → up to 20% of ZIPs removable
- Moderate → up to 10%

The cap is a ceiling, never a target. If only three ZIPs genuinely stand out under a 20% cap, we cut three.

**When more items qualify than the cap allows:** the tool cuts the worst up to the cap, then flags the rest — "these others weren't far behind, worth watching now or next month."

**Reporting period toggle:** month-only / campaign-to-date / both.

**Optional:** convert attributed rate to attributed unique visitors by variable, for clients who ask for it.

---

## 7. Back to the Proposal Builder

The follow-up proposal gets a checkbox that pulls in attribution results. Optimizations we made appear in the proposal — likely a dedicated slide plus a mention in the main proposal slide — showing how the next phase builds on what we learned.

The point is connecting what we said we'd do to what we actually delivered.

---

## 8. Build Order

1. Ingest raw Excel, rebuild clean report
2. Analysis layer — outlier detection + URL/page intent story
3. Optimization engine — levels, checkboxes, caps
4. Proposal Builder loop-back

---

## Open Questions

- Does the raw Excel export contain full granularity (attributed rate by ZIP, day, creative, market), or only summarized rollups?
- Are the 20% / 10% caps right, or should they be tuned?
- Should the ZIP cap be based on count of ZIPs, or share of impressions those ZIPs represent?
- Which parts of the current report do we actually speak to in a client presentation, versus leave in as reference?

---

**Guiding principle:** simplicity. Every section should earn its place. Shorter and clearer beats thorough.
