"""The avails PDF importer's unresolved-notes list -- two live bugs in one
report, both from `apply_avails_import` (app.py):

1. Every note named the WHOLE DOCUMENT ("Premion Media Plan_RFPID-...pdf"),
   never the specific group it was about. Harmless with one group; with a
   real 12-group document, two unrelated groups hitting the same failure
   read as one duplicated line, because neither said which of the 12 it
   was. Fixed: every note is prefixed by that group's own geo/audience
   label instead.

2. `apply_avails_import` never applied the calm-vs-warning split
   `geo_resolver`'s `Resolution.notes`/`.unresolved` shape exists for (see
   DECISIONS.md, the 2026-08-20 zip-messaging fix already applied to the
   interactive geo-definition expander) -- it dumped EVERY entry of a
   named-zip group's `.unresolved` as its own report line, including every
   well-formed zip that simply has no county/market on file, a case
   `.notes` already summarizes in one sentence. A real Hershey zip-add-on
   group has 60 such zips; this turned one calm, expected fact into 60
   near-identical lines. Fixed: `.notes` is surfaced (it used to be
   silently discarded entirely) and `.unresolved` is filtered to genuinely
   malformed entries only for the named-zip branch.

    python tests/test_avails_import_notes.py
"""
import os
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

import app  # noqa: E402
import avails_pdf_import as api  # noqa: E402

REAL_PDF = REPO / ("Premion Media Plan_RFPID-260402_Direct - No Agency_"
                    "Visit Hershey & Harrisburg_4-30-2026--ver0.pdf")

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


class _StStub:
    def __init__(self, session_state):
        self.session_state = session_state


def run_import(document):
    real, stub = app.st, _StStub({})
    app.st = stub
    try:
        return app.apply_avails_import(document)
    finally:
        app.st = real


def main():
    print("=" * 78)
    print("SCENARIO  a bulk of well-formed, unmapped zips gets ONE calm summary, "
          "not one line each")
    print("=" * 78)
    # 99999 is the project's own established "well-formed, deliberately not
    # on file" test zip (tests/test_group_geo_resolution.py); a real PDF
    # doesn't repeat one zip 60 times, so this uses 60 distinct fake-but-
    # well-formed ones (99900-99959) to reproduce the same shape honestly.
    fake_zips = [str(99900 + i) for i in range(60)]
    document = api.AvailsDocument(
        rfpid="RFPID-BULK", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=1_000_000, attribution_text="",
        groups=[api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                                geo_name="Bulk Zip Add-On", zips=fake_zips,
                                impressions=1_000_000, row_count=1)],
        source_name="test.pdf")

    new_groups, report = run_import(document)
    zip_notes = [n for n in report["unresolved"] if "Bulk Zip Add-On" in n]
    check("60 unmapped zips produce only a couple of notes, not 60",
          0 < len(zip_notes) <= 3, zip_notes)
    check("no single zip code from the fake list appears verbatim as its own "
          "report line (that would mean the bulk dump survived)",
          not any(any(z in n for z in fake_zips) for n in zip_notes), zip_notes)
    check("the resolved group still carries every one of the 60 zips -- "
          "unmapped-to-a-market was never a drop, just quieter messaging",
          new_groups[0]["resolved_zips"] == sorted(fake_zips), new_groups[0]["resolved_zips"])

    print("\n" + "=" * 78)
    print("SCENARIO  a genuinely malformed zip still gets its own line")
    print("=" * 78)
    document2 = api.AvailsDocument(
        rfpid="RFPID-BAD", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=100_000, attribution_text="",
        groups=[api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                                geo_name="Bad Zip Option", zips=["20005", "NOT-A-ZIP"],
                                impressions=100_000, row_count=1)],
        source_name="test.pdf")
    _, report2 = run_import(document2)
    malformed_notes = [n for n in report2["unresolved"] if "Bad Zip Option" in n]
    check("the malformed entry gets its own, specific line",
          any("NOT-A-ZIP" in n for n in malformed_notes), malformed_notes)
    check("a well-formed zip (20005) is never reported as invalid",
          not any("20005" in n and "isn't a valid zip" in n for n in malformed_notes),
          malformed_notes)

    print("\n" + "=" * 78)
    print("SCENARIO  notes are labeled by GROUP, not by the whole document")
    print("=" * 78)
    document3 = api.AvailsDocument(
        rfpid="RFPID-TWO", advertiser="Test Advertiser", agency="No Agency",
        flight_start=date(2026, 9, 1), flight_end=date(2026, 11, 30),
        total_impressions=200_000, attribution_text="",
        groups=[
            api.AvailsGroup(audience_text="Homeowners", geo_kind=api.GEO_KIND_NAMED_ZIP,
                            geo_name="Option A", zips=["99900"], impressions=100_000, row_count=1),
            api.AvailsGroup(audience_text="Auto Intenders", geo_kind=api.GEO_KIND_NAMED_ZIP,
                            geo_name="Option B", zips=["99900"], impressions=100_000, row_count=1),
        ],
        source_name="a_very_long_filename_that_says_nothing_useful_twice.pdf")
    _, report3 = run_import(document3)
    check("Option A's note names Option A, not the filename",
          any("Option A" in n for n in report3["unresolved"]), report3["unresolved"])
    check("Option B's note names Option B, not the filename",
          any("Option B" in n for n in report3["unresolved"]), report3["unresolved"])
    check("neither note is bare document-filename text with no group identity",
          not any(n.startswith(document3.source_name) for n in report3["unresolved"]),
          report3["unresolved"])
    # The two groups hit the identical underlying fact (99900 has no
    # county on file) -- with group labels, that's two DISTINGUISHABLE
    # notes, not one that reads like a duplicate.
    check("the two groups' notes are distinguishable from each other "
          "even though they hit the identical underlying zip",
          len({n for n in report3["unresolved"]}) == len(report3["unresolved"]),
          report3["unresolved"])

    if REAL_PDF.exists():
        print("\n" + "=" * 78)
        print("SCENARIO  the real Visit Hershey & Harrisburg document stays readable")
        print("=" * 78)
        import market_lookup
        market_lookup.install()
        document4 = api.parse_avails_pdf(str(REAL_PDF), REAL_PDF.name)
        _, report4 = run_import(document4)
        # Before this fix: 9 catalog notes + ~120 individual zip lines (one
        # of the document's two zip-add-on options alone lists 60 zips) + 1
        # attribution note. After: 9 + a handful of summaries + 1.
        check(f"the real document's report stays well under the old "
              f"one-line-per-zip volume (got {len(report4['unresolved'])})",
              len(report4["unresolved"]) < 25, len(report4["unresolved"]))
        # The one deliberate exception: the document-level Attribution note
        # is genuinely about the WHOLE document, not one group, and keeps
        # naming it -- every GROUP note (everything else) should not.
        group_notes = [n for n in report4["unresolved"] if "lists Attribution:" not in n]
        check("no group-level note is prefixed by the raw document filename anymore",
              not any(n.startswith(REAL_PDF.name) for n in group_notes), group_notes)
    else:
        print(f"\nSKIP -- real document not found at {REAL_PDF.name}")

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("Unresolved notes are labeled by the group they're about, and a well-formed zip "
          "with no county/market on file is summarized once per group instead of listed "
          "individually -- a real 60-zip add-on option now reads as one calm sentence "
          "instead of 60 near-identical lines.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
