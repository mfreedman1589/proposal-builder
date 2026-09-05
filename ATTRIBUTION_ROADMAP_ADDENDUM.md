# Attribution Report Builder — roadmap addendum

Adds to the existing planning document rather than replacing it. Three things: a phasing decision, template rules the Mattress Warehouse rebuild established, and a new feature that closes the loop back to the Proposal Builder.

---

## 1. Phase by INPUT, not by product

Website attribution first, sales match-back and brand lift later. That's right, and the reason is stronger than "they're less common."

**Each of the three is a different ingestion problem, not a different section of one report.**

| Product | Source | Shape |
|---|---|---|
| Website attribution | Premion "Website Attribution and Reach Extension" Excel | Many tabs, one per widget, consistent schema |
| Polk automotive | S&P Global Mobility registration match-back | Registrations by exposed household |
| Auto-Sales Analyst | The existing Streamlit app's own output | Dealer inventory movement — units sold, value, look-to-book |
| Sales match-back | CRM match-back, client- or vendor-supplied | No standard format; varies per advertiser |
| Brand lift | Study vendor deliverable | Survey-shaped — lift by attribute, by segment, vs competitors |
| Arrivalist | Location-data visitation measurement | Arrivals to a destination from exposed households |

Phase 1 has **one input format** and it's already in hand. Every add-on introduces a new parser, a new set of tab-identification rules, and a new slide family. Building all three at once means three half-understood formats instead of one solid one.

This also matches how the reports actually get sold: website attribution is on nearly every campaign, the other two are added value on bigger buys.

### Phase 1 scope

- Ingest the Premion website attribution Excel
- **Identify tabs by header row, never by tab name** — the export tool emits "Untitled Widget" and "ATTRIBUTED RATE AND CONVERSIO_2"
- Rebuild every chart from raw data; never repurpose the templated PPTX graphics
- Strip internal nomenclature (RFP numbers, Salesforce IDs, hub-team language) — rebuilding from raw is what makes this free
- Proposal color scheme, not the report template's

### No projection step

The Mattress Warehouse under-reporting was an isolated system issue, not a recurring condition. The builder takes the raw export at face value — no modeling workbook, no confidence tiers, no correction pass.

One consequence to expect while building: if the MW report is rebuilt from raw as a test case, its website figures will come out lower than the deck that was sent, because that deck's numbers were projected. That's the known difference, not a bug in the rebuild.

---

## 2. Template rules the Mattress Warehouse rebuild established

These stopped being one-off fixes the moment they worked. They're the default structure for every generated report.

**Takeaways come first.** The rebuild pulled key takeaways forward into a "what you're about to see" slide. A client who reads one slide should get the whole story; everything after it is evidence.

**KPIs over charts on slide 2.** The original had a chart there. Emphasized KPI figures replaced it and read better. Charts earn their place further in, where someone is already engaged.

**Delivery is context, not content.** The tendency is to over-spend on delivery detail exactly when the client cares about attribution. Impressions delivered, delivery by creative, top publishers — enough to frame the KPIs, no more.

**Publisher story without publisher optimization.** Show the commercial running across a range of premium publishers (SVOD, FAST) with website impact traced across that range. That proves the aggregator's value without arguing for it, and without inviting publisher-level optimization the report shouldn't recommend.

**Simplicity throughout.** Concise, meeting-ready bullets. The pull toward long explanatory prose is the thing to resist.

**Ends on what's next.** Extension, evolution, growth — and that's the handoff to the Proposal Builder.

---

## 3. New feature: Create case study from report

**The idea:** a wrap-up report already contains everything a case study needs. One button turns it into a templated one-slide case study, ready for the Slide Vault's sibling — the case study vault the Proposal Builder already uses.

**Why it fits:** the case study vault's shape already exists — title, summary, verticals, products, image path, active flag, and the Claude auto-tagging pass that suggests tags a human then corrects. A report-generated case study writes into that same table through the same path. No new storage, no new picker, no new deck-insertion logic.

**What it closes:** proposal → campaign → report → case study → the next proposal. Every campaign that ran well becomes evidence for the next pitch, automatically, instead of when someone remembers to write it up. That's the part of the loop that's manual today and therefore mostly doesn't happen.

### Open question worth answering before building

**Does the case study come from the report alone, or the report plus its originating proposal?**

The report has results. The proposal has the goal — what the client wanted, what audience, what strategy. A case study that says "we drove 15K unique visitors" is weaker than one that says "the client wanted in-store intent from in-market shoppers, and we drove 15K unique visitors who engaged heavily on product and location pages."

Since the planning document already calls for linking each report to its originating proposal, the second version is available for free once that link exists. Recommend building the feature **after** the report-to-proposal link, so it can use both.

### Guards this feature needs

- **Nothing publishes unreviewed.** Same rule as the Slide Vault picker: suggestions are filled in, a human corrects before saving.
- **Client anonymity.** Some case studies can name the advertiser and some can't. That's a per-report decision the generator has to ask rather than assume.

---

## 4. The automotive phase — why it comes first among the add-ons

Automotive is the biggest vertical, so this add-on gets used most. But it also has a head start nothing else has: **half of it is already built and running in production.**

The Auto-Sales Analyst already scans dealer inventory, detects sold vehicles, and produces a PPTX framed explicitly as *complementary data showing the impact of website traffic on moving dealer inventory*. That's not a coincidence — it was designed as a companion to the Premion Website Attribution report, which is exactly what the report builder generates.

**Polk and the Auto-Sales Analyst answer the same question from opposite ends.** Polk asks whether an exposed household registered a vehicle anywhere. The Analyst asks whether this dealer's specific inventory moved. Together they say "the campaign drove buyers, and it drove them to *you*" — which is a stronger claim than either alone, and it's the one a dealer group actually wants.

### The fork to answer before building

Does the report builder **absorb**, **call**, or **align with** the Auto-Sales Analyst?

- **Absorb** — port the scanning logic in. Cleanest single report, but it duplicates a working tool and abandons a codebase that already handles the hard parts (blocked dealers, MarketCheck governance, VIN plausibility, page-ID fallbacks).
- **Call** — the report builder invokes the Analyst and folds its output in. Keeps one implementation of the hard parts. Needs an interface that doesn't exist yet.
- **Align** — both stay separate, share a visual template, and a rep assembles the two decks. Cheapest, and closest to how it works today.

Weight in the decision: TEGNA's product team is looking at integrating the Analyst into their reporting platform. Absorbing it here could work against that, or duplicate what they build. Worth knowing where that lands before committing to the first option.

---

## 5. Suggested build order

1. **Website attribution ingestion** — the Excel, tab identification by header row
2. **Report generation** — the template structure in §2, with the reporting-period toggle (month / campaign-to-date / both)
3. **Optimization engine** — the level bar, the dimension checkboxes, subtraction-only logic, the cap-vs-threshold asymmetry
4. **Report ↔ proposal link** — pull strategy and goals from the originating proposal into the report's setup section
5. **Create case study from report** — needs (4) to be worth building well

Then the add-ons, in priority order:

6. **Polk automotive + Auto-Sales Analyst** — the biggest vertical, and partly built already. Answer the fork in §4 first.
7. **Sales match-back** — new parser, new slide family. Hardest of the four, since there's no standard input format.
8. **Brand lift** — new parser, new slide family. Survey-shaped data, so the charts are unlike anything in phases 1-6.
9. **Arrivalist for tourism** — narrowest vertical, so last, but the most self-contained: visitation is a single clear metric.

Items 6-9 are independent of each other and of 3-5. Any of them can jump the queue when a real example lands and a client is waiting on it.

---

## 6. What's still needed from Matt

- One report paired with its originating proposal, to design the link in (4)
- A note on which parts of a report get spoken aloud to clients versus kept as reference
- A Polk match-back deliverable in its raw form (for phase 6)
- A real CRM sales match-back file (phase 7)
- A brand lift study deliverable, raw (phase 8)
- An Arrivalist export (phase 9)

The last four aren't needed until their phase comes up — but each one gates its phase entirely, since the whole design problem is the input format.
