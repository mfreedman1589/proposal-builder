"""
test_vertical_hints.py -- _detect_vertical_hint coverage.

Discovery notes name a business the way its owner would ("an HVAC company",
"a credit union"), never by vertical key. This checks that every vertical is
reachable from its trade names, that short names don't misfire on ordinary
English, and that the cases with no good answer return None rather than a
wrong one.

Plain script, no test framework -- run it directly:

    python test_vertical_hints.py
"""
import app

failures = []


def check(label, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {label}{(' -- ' + detail) if detail else ''}")
    if not ok:
        failures.append(label)


def expect(notes, want):
    got = app._detect_vertical_hint(notes)
    check(f"{notes[:56]:56s} -> {str(got)}", got == want,
          f"expected {want}")


print("== Trade names resolve to their vertical ==")
CASES = [
    # home improvement -- the trades
    ("Regional HVAC company, wants more service calls", "home_improvement"),
    ("A roofing contractor in three markets", "home_improvement"),
    ("Local plumbing business", "home_improvement"),
    ("Pest control franchise, spring campaign", "home_improvement"),
    ("Landscaping and lawn care company", "home_improvement"),
    ("Replacement windows and siding installer", "home_improvement"),
    ("Kitchen and bath remodeling showroom", "home_improvement"),
    ("Solar installer targeting homeowners", "home_improvement"),
    ("Water damage restoration company", "home_improvement"),
    # automotive
    ("Ford dealership group, Q4 push", "auto"),
    ("Independent auto repair chain", "auto"),
    ("Tire shop with four locations", "auto"),
    ("Collision center near the highway", "auto"),
    # healthcare
    ("A med spa opening a second location", "healthcare"),
    ("Dental practice, new patient acquisition", "healthcare"),
    ("Urgent care network across two DMAs", "healthcare"),
    ("Physical therapy clinic", "healthcare"),
    ("Senior living community", "healthcare"),
    ("Veterinary hospital group", "healthcare"),
    # banking
    ("Local credit union promoting auto loans", "banking"),
    ("Wealth management firm", "banking"),
    ("Mortgage lender in a rising-rate market", "banking"),
    # dining
    ("Regional pizzeria chain", "dining_qsr"),
    ("Fast casual restaurant group", "dining_qsr"),
    ("Local brewery and taproom", "dining_qsr"),
    ("Coffee shop with six locations", "dining_qsr"),
    # retail
    ("Furniture store clearance event", "retail"),
    ("Fine jewelry retailer, holiday campaign", "retail"),
    ("Mattress chain", "retail"),
    ("Garden center spring season", "retail"),
    # travel
    ("Beachfront resort, summer bookings", "travel"),
    ("Regional casino promoting a concert series", "travel"),
    ("Convention and visitors bureau", "travel"),
    # entertainment
    ("The county fair, ticket sales", "entertainment"),
    ("Independent movie theater", "entertainment"),
    ("Children's museum membership drive", "entertainment"),
    # education
    ("Community college enrollment campaign", "education"),
    ("Trade school recruiting welders", "education"),
    ("Career training program", "education"),
    # legal
    ("Personal injury law firm", "legal"),
    ("Criminal defense attorney", "legal"),
    ("Estate planning practice", "legal"),
]
for notes, want in CASES:
    expect(notes, want)

print("\n== Short names must not misfire on ordinary words ==")
GUARDS = [
    ("The entire campaign runs September through November", None),
    ("Revenue goals for the year are aggressive", None),
    ("Automatic bid adjustments are enabled", None),   # 'auto' inside 'automatic'
    ("Send the deck to the client on Friday", None),
    ("Spanish-language creative for both markets", None),
    ("Barbershop quartet sponsorship", None),
    ("A conduit for brand awareness", None),           # 'dui'
    ("Their marketing budget is unclear", None),
]
for notes, want in GUARDS:
    expect(notes, want)

print("\n== The vertical's own name beats an incidental trade name ==")
expect("Our automotive client also needs a bank partner", "auto")
# Longest phrase wins over one combined table, so a genuinely mixed sentence
# resolves on its most specific term rather than on term *type*. Realistic
# notes name one business; contrived ones are a coin toss either way.
expect("Regional healthcare system, three hospitals", "healthcare")

print("\n== Longest phrase wins ==")
expect("A med spa, not a day spa", "healthcare")
expect("Auto dealer group", "auto")

print("\n== Deliberately unmapped: no vertical fits property/real estate ==")
for notes in ["Apartment complex leasing campaign", "Realtor group in two markets",
              "Property management company"]:
    got = app._detect_vertical_hint(notes)
    check(f"{notes[:52]:52s} -> {got}", got is None,
          f"got {got}; a wrong hint is worse than none")

print("\n== Every synonym maps to a real vertical ==")
valid = set(app.VERTICALS.values())
bad = {p: v for p, v in app.VERTICAL_HINT_SYNONYMS.items() if v not in valid}
check("all synonym targets exist in VERTICALS", not bad, str(bad))

print("\n== Every vertical has at least one trade synonym ==")
covered = set(app.VERTICAL_HINT_SYNONYMS.values())
missing = {v for v in valid if v != "none"} - covered
check("no vertical left without trade names", not missing, str(missing))

print(f"\n{len(app.VERTICAL_HINT_SYNONYMS)} synonyms across {len(covered)} verticals")
print("\n" + ("HINTS VERIFIED" if not failures else f"{len(failures)} FAILURE(S)"))
raise SystemExit(1 if failures else 0)
