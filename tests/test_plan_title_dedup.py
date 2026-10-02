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
    # 2026-10-02 rule (Matt): the title must not CONTAIN the client name
    # anywhere -- the slide prints the name before it, so a trailing echo
    # doubles it just as a leading one does. A connector left hanging goes too.
    check("client name at the END of the title is stripped, dangling 'for' with it",
         app.option_plan_title("Growth Plan for WAEPA", "Good", False, client_name="WAEPA")
         == "Growth Plan", app.option_plan_title("Growth Plan for WAEPA", "Good", False, client_name="WAEPA"))

    print("\nA client name the title spells differently still matches (Visit Hershey QA, 2026-10-02)")
    cases = [
        ("a prefix on the client name the title lacks",
         "Visit Hershey & Harrisburg CTV Plan", "QA-TEST-Visit Hershey & Harrisburg", "CTV Plan"),
        ("'&' vs 'and'", "Visit Hershey and Harrisburg CTV Plan",
         "Visit Hershey & Harrisburg", "CTV Plan"),
        ("punctuation and case", "visit hershey, harrisburg: Summer Streaming",
         "Visit Hershey Harrisburg", "Summer Streaming"),
        ("a suffix on the client name the title lacks", "Ridgeline Heating Fall Push",
         "Ridgeline Heating & Air, Inc.", "Fall Push"),
        ("possessive", "WAEPA's Spring Strategy", "WAEPA", "Spring Strategy"),
        ("one shared common word is NOT a match", "Visit Planning Guide",
         "Visit Hershey & Harrisburg", "Visit Planning Guide"),
        ("a word starting with s after the name keeps its s", "WAEPA spring plan", "WAEPA",
         "spring plan"),
    ]
    for label, title, client, expected in cases:
        got = app.option_plan_title(title, "Good", False, client_name=client)
        check(label, got == expected, got)

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
