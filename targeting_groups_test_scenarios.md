# Targeting Groups — Test Scenarios

Five scenarios, roughly 10 minutes each. Run them in the deployed app, not through
AppTest — the point is to see what a rep sees.

**Note on the avails documents:** the PDF importer is Prompt F and isn't built yet, so
these are entered by hand. That's not a workaround — it's the useful version of the
test, because it establishes exactly what the importer will need to produce. If the
manual path can't express a real avails document, the importer can't either.

---

## Scenario 0 — Geo builder with no avail (5 minutes)

The fastest check that all four geo modes work standalone.

Create one group with any audience, then exercise each mode in its geo expander:

| Mode | Input | Expect |
|---|---|---|
| Markets | Pick Denver, Atlanta | Resolves directly, no zips needed |
| Zips | `20120, 20147, 20148, 99999` | Resolves the three real ones; **99999 reported as unresolved, not silently dropped** |
| Counties | `NEW CASTLE DE; CHESTER PA` | Resolves to zips, then to markets |
| Radius | 10 miles around `21401` | Resolves to a zip list |

The unresolved zip is the one that matters — a bad zip vanishing quietly is the failure
mode, and the resolver was built to report rather than drop.

---

## Scenario 1 — Core group mechanics (5 minutes)

No avail, no geo resolution. Pure structure.

1. Build **audience A**: Homeowners AND HH Income $150K Plus (two AND clicks)
2. **New group**, build **audience B**: In-market for windows (or any distinct segment)
3. Assign both to markets **Denver, Atlanta, Phoenix**, separate (not combined)

**Expect: 6 avails rows and 6 plan lines, audience-major** — A/Denver, A/Atlanta,
A/Phoenix, then B/Denver, B/Atlanta, B/Phoenix. Not 3, not 6 in geo-major order.

4. **Merge** the two Denver lines. One line, both markets in its Geo.
5. **Split** it back. Two lines again.
6. Check the avails table is unchanged throughout — merge and split act on plan lines,
   never on groups.

Also confirm: the split control says the division is even and original amounts aren't
recovered.

---

## Scenario 2 — Annapolis Cars (radius mode, agency, 8 groups)

Reproduces a real avails document. **Agency: SR&B Advertising** — so gross markup applies.
Flight 03/01/2026 – 03/31/2026.

Four audiences, each at two radii from **21401**:

| Audience | Geo | Avails |
|---|---|---|
| AUTO Make Subaru | 10mi radius 21401 | 602,647 |
| AUTO Make Subaru | 5mi radius 21401 | 165,687 |
| AUTO Make Hyundai | 10mi radius 21401 | 669,089 |
| AUTO Make Hyundai | 5mi radius 21401 | 184,798 |
| AUTO Make Volvo | 10mi radius 21401 | 279,310 |
| AUTO Make Volvo | 5mi radius 21401 | 84,637 |
| AUTO Make Genesis Intender | 10mi radius 21401 | 418,144 |
| AUTO Make Genesis Intender | 5mi radius 21401 | 118,404 |

**Checks:**
- 8 rows in exactly that order — audience-major, matching the document's own ordering
- Avails total **2,522,716**
- 10mi resolves to ~18 zips, 5mi to ~7 — the document lists both, so compare
- Markets resolve to **Baltimore** (21401 is Annapolis, Anne Arundel County) — confirm
  it's the market you'd name
- Vertical detects as auto; the **Polk New Car Sales Attribution** checkbox appears
  pre-checked
- Agency toggle on → every line marked Gross

This is also the scenario where the same audience appears twice with different geos, and
the same geo appears under four audiences — the case that must never collapse.

---

## Scenario 3 — Visit Hershey (multi-market + zip add-ons, 12 groups)

The richest one. **Direct, no agency.** Flight 05/25/2026 – 07/05/2026.

Two compound audiences:
- **A:** AFIRST Travel Buffs and Sightseers **AND** DEMO Age A21-44
- **B:** TRAVEL Family **AND** DEMO Age A25 Plus

| Audience | Geo | Avails |
|---|---|---|
| A | Philadelphia | 38,426,724 |
| A | Baltimore | 14,635,026 |
| A | Washington, D.C. | 25,264,764 |
| A | New York | 66,349,836 |
| A | Philly zip add-on | 5,829,978 |
| A | NY zip add-on | 1,011,108 |
| B | Philadelphia | 40,732,230 |
| B | Baltimore | 15,489,096 |
| B | Washington, D.C. | 26,736,864 |
| B | New York | 70,003,248 |
| B | Philly zip add-on | 5,397,126 |
| B | NY zip add-on | 924,336 |

Zip add-on lists are in the source document — paste them into the two zip-mode groups.

**Checks:**
- 12 rows, audience-major
- Avails total **310,800,336**
- The four market slides auto-add (Philadelphia, Baltimore, Washington DC, New York) and
  are removable
- The **zip add-on** groups resolve to markets independently — the Philly list should
  land in Philadelphia's orbit, the NY list in New York's. This is the check that the
  resolver agrees with what the document already says.
- Both audiences render as plain language on the slides ("Travel Buffs and Sightseers,
  Adults 21-44"), never as boolean syntax
- Remove one auto-added market, then add a group that resolves to it again — **it must
  stay removed**

---

## Scenario 4 — Wilmington University (county mode, 5 audiences, 12-month flight)

**Direct, no agency.** Flight 07/01/2026 – 06/30/2027.

One county list, five audiences:

`NEW CASTLE DE; KENT DE; SUSSEX DE; CHESTER PA; PHILADELPHIA PA; DELAWARE PA; BUCKS PA; SALEM NJ; GLOUCESTER NJ; CAMDEN NJ`

| Audience | Avails (full flight) |
|---|---|
| LIFESTAGE College Planning Parents | 54,005,249 |
| LIFESTAGE Prospective College Students | 77,173,481 |
| LIFESTYLE Online Education | 73,412,846 |
| DEMO Career Employed **AND** LIFESTAGE Education Services | 52,725,960 |
| LIFESTAGE Higher Education Intender | 74,697,572 |

**Checks:**
- 5 rows, one per audience, all sharing the same county-derived geo
- Total **332,015,108**
- The semicolon-separated `NAME STATE` form resolves — no suffixes given, which is how
  the documents write it
- Markets resolve to more than one DMA (Philadelphia for most; the Delaware counties may
  reach further) — confirm the set is one a person would accept
- Avails basis toggle: these are full-flight figures over a 12-month flight, so check the
  monthly view derives sensibly through the daily rate rather than dividing by month count
- Four single-term audiences and one AND compound in the same proposal

---

## Scenario 5 — Backward compatibility (the one that matters most)

1. Open **Proposal History** and pick a proposal generated **before** targeting groups
2. **Rebuild as presented**
3. Compare the deck to the one produced originally — it must be **byte-identical**

If you don't have the original file, generate the same proposal twice: once via Rebuild
as presented, then Load into form and generate without touching anything. Those two
should match.

This is the promise the whole feature was built around, and the round-trip property
proves it at the data layer. This proves it where it counts.

---

## What to watch for beyond pass/fail

- Does the three-action audience builder (AND / OR / New group) read clearly, or do you
  have to think about which one you want?
- Does the booking evidence panel say anything useful while you build, or is it noise?
- With 8 or 12 rows, does the media plan table still render legibly, and does the deck's
  plan slide hold up?
- Is anything in the grouped avails table harder to do than it was before?

That last question is the one no test covers: the feature is only worth it if a rep
finds it easier than the flat table it replaced.
