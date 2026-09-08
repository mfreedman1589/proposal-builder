"""option_plan_title strips a leading client-name echo before it reaches
PLAN_TITLE -- pure, offline, no framework.

    python tests/test_plan_title_dedup.py

The bug: the master deck's own media plan slide composes its heading as the
literal template text "{{CLIENT_NAME}} {{PLAN_TITLE}}" (TEGNA_MASTER_
DECK_v1_1.pptx, slides 122-124, Title 10 -- one text run, not something this
app assembles). A real WAEPA proposal had "Proposal title" (a field labelled
"appears on cover + media plan") typed as "WAEPA CTV Strategy" -- a natural
thing to type, not realizing the slide already prepends the client name on
its own -- and the plan slide read "WAEPA WAEPA CTV Strategy". This pins the
fix: PLAN_TITLE composition is the one place that owns deduping against
CLIENT_NAME, whatever put the echo there.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    print("A title echoing the client name is stripped")
    check("plain 'ClientName Title' -> 'Title'",
         app.option_plan_title("WAEPA CTV Strategy", "Good", False, client_name="WAEPA")
         == "CTV Strategy")
    check("case-insensitive match",
         app.option_plan_title("waepa CTV Strategy", "Good", False, client_name="WAEPA")
         == "CTV Strategy")
    check("a dash separator is stripped too",
         app.option_plan_title("WAEPA - CTV Strategy", "Good", False, client_name="WAEPA")
         == "CTV Strategy")
    check("a colon separator is stripped too",
         app.option_plan_title("WAEPA: CTV Strategy", "Good", False, client_name="WAEPA")
         == "CTV Strategy")

    print("\nA title with no echo is untouched")
    check("no client name in the title -- unchanged",
         app.option_plan_title("CTV Strategy", "Good", False, client_name="WAEPA")
         == "CTV Strategy")
    check("no client_name given at all -- unchanged (default None)",
         app.option_plan_title("WAEPA CTV Strategy", "Good", False) == "WAEPA CTV Strategy")
    check("client name appearing mid-title (not a LEADING echo) is left alone",
         app.option_plan_title("Growth Plan for WAEPA", "Good", False, client_name="WAEPA")
         == "Growth Plan for WAEPA")

    print("\nStripping to nothing falls back to the original rather than an empty title")
    check("the title IS just the client name -- keep it, don't blank the slide",
         app.option_plan_title("WAEPA", "Good", False, client_name="WAEPA") == "WAEPA")

    print("\nThe multi-option suffix still appends after stripping")
    check("multi-option suffix applies to the STRIPPED title",
         app.option_plan_title("WAEPA CTV Strategy", "Good", True, client_name="WAEPA")
         == "CTV Strategy — Good")

    total = len(failures)
    print(f"\n{'ALL PASSED' if not failures else f'{total} FAILED'}")
    if failures:
        for f in failures:
            print(f"  FAILED  {f}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
