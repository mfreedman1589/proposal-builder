"""The stitched heating fixture (2026-11) -- the real stitched workbook
(local only: client data in the gitignored tests/fixtures/stitched/). The
synthetic versions CI runs are tests/test_stitched_attribution.py.

    python tests/test_stitched_attribution_real.py

Expected values are the QA answer key's for "the stitched heating fixture
(2026-11)": the stitched workbook must read exactly like the plain export of
the same figures, and June and July (projected) carry the estimated note.
"""
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import attribution_import as ai  # noqa: E402
import report_assembly as ra  # noqa: E402

FIXTURES = REPO / "tests" / "fixtures" / "stitched"
failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:400]}" if detail else ""))


def main():
    stitched_path, plain_path = FIXTURES / "stitched.xlsx", FIXTURES / "plain.xlsx"
    if not (stitched_path.exists() and plain_path.exists()):
        print("SKIP -- the stitched fixtures aren't present")
        return 0
    stitched = ai.parse_attribution_export(str(stitched_path))
    plain = ai.parse_attribution_export(str(plain_path))
    check("the stitched workbook reads 3,264,153 delivered / 25,997 attributed / 3,696 visitors",
          (stitched.delivered_impressions, stitched.attributed_impressions,
           stitched.attributed_unique_visitors) == (3264153, 25997, 3696),
          (stitched.delivered_impressions, stitched.attributed_impressions))
    check("...with its client, 2 audiences, 55 URLs and 125 ZIPs",
          bool(stitched.client_name) and len(stitched.by_audience) == 2
          and len(stitched.by_url) == 55 and len(stitched.by_zip) == 125,
          (stitched.client_name, len(stitched.by_audience), len(stitched.by_url), len(stitched.by_zip)))
    def norm(value):
        """Rates differ in the 17th digit between the two files' own cells."""
        if isinstance(value, list):
            return [norm(v) for v in value]
        if isinstance(value, dict):
            return {k: norm(v) for k, v in value.items()}
        if hasattr(value, "__dataclass_fields__"):
            return norm(vars(value))
        return round(value, 9) if isinstance(value, float) else value
    same = all(norm(getattr(stitched, f)) == norm(getattr(plain, f)) for f in
               ("client_name", "delivered_impressions", "attributed_impressions",
                "attributed_unique_visitors", "by_audience", "by_creative", "by_market", "by_zip",
                "by_url", "by_recency", "by_referral_domain", "by_device", "weekly_trend",
                "flight_start", "flight_end"))
    check("every figure matches the plain export of the same data", same)
    check("June and July are estimated", ra.estimated_periods(stitched) == ["June", "July"],
          stitched.period_summary)
    check("the note", ra.estimated_periods_note(stitched)
          == "June–July figures are estimated (tracking gap).")
    check("the plain export has no estimated periods", ra.estimated_periods_note(plain) is None)
    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
