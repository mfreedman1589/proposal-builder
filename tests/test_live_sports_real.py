"""Live sports against the real 2026-11 exports (local only: client data, in
the gitignored tests/fixtures/live_sports/). The synthetic versions CI runs
are tests/test_live_sports_delivery.py.

    python tests/test_live_sports_real.py

Expected values are the ones the QA answer key gives for this fixture.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import attribution_import as ai  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures" / "live_sports"
DELIVERY = FIXTURES / "delivery.xlsx"

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {detail}" if detail else ""))


def check_delivery():
    print("Delivery export, 2026-11 format")
    if not DELIVERY.exists():
        print(f"  SKIP  {DELIVERY.name} not present")
        return
    d = ai.parse_delivery_export(str(DELIVERY))
    ls = d.live_sports
    check("the sports block is detected, with no warning", ls is not None and not d.warnings,
          d.warnings)
    if ls is None:
        return
    check("543,301 sports impressions delivered", ls.delivered_impressions == 543_301,
          ls.delivered_impressions)
    check("98.5% sports VCR", round(ls.vcr * 100, 1) == 98.5, ls.vcr)
    check("two RFPIDs", ls.rfpids == ["RFPID-268085", "RFPID-257034"], ls.rfpids)
    check("OTT delivered is the OTT figure (601,561), not the sports one",
          d.delivered_impressions == 601_561, d.delivered_impressions)
    top = ls.games()[0]
    check("49ers v. Commanders is the top game after aggregation",
          top.event == "San Francisco 49ers - Washington Commanders"
          and top.delivered_impressions == 87_307, (top.event, top.delivered_impressions))
    check("...across ESPN and ABC", top.network == "ESPN / ABC", top.network)
    check("the 30 event rows become 18 games", (len(ls.events), len(ls.games())) == (30, 18),
          (len(ls.events), len(ls.games())))


def main():
    check_delivery()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
