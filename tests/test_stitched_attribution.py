"""Stitched attribution workbooks and the failures around them (2026-11).

    python tests/test_stitched_attribution.py

Synthetic workbooks only (CI); tests/test_stitched_attribution_real.py
checks the real stitched fixture locally.

  1. A stitched workbook -- the combined block left of a "Source period"
     column, per-period blocks to its right, a leading "Period Summary" tab
     -- parses exactly like the plain export of the same figures; a
     workbook whose core tabs come out empty raises an error naming them,
     never a silent zero.
  2. Periods the Period Summary marks as projected (a factor other than 1,
     or reported != used) are footnoted "June–July figures are estimated
     (tracking gap)", on the slides and the one-sheet, and the model is told.
  3. A build failure names the missing field -- or the failed upload.
  4. An existing client can always be picked; a different advertiser with
     the same name gets its own record without renaming.
"""
import os
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"

import openpyxl  # noqa: E402
from pptx import Presentation  # noqa: E402

import app  # noqa: E402
import attribution_import as ai  # noqa: E402
import db  # noqa: E402
import report_assembly as ra  # noqa: E402
import slide_map  # noqa: E402

failures = []
TMP = Path(tempfile.mkdtemp())
CLIENT = "Example Heating & Cooling"


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        failures.append(label)
        print(f"  FAIL  {label}" + (f"  -- {str(detail)[:400]}" if detail else ""))


PERIODS = [("January–May", 1_600_000, 10_000), ("June", 330_000, 1_900),
           ("July", 330_000, 300), ("August–October", 980_000, 13_600)]
DELIVERED = sum(p[1] for p in PERIODS)
ATTRIBUTED = sum(p[2] for p in PERIODS)
AUDIENCES = [("DEMO Homeowners", 1_620_000, 12_800), ("LIFESTYLE New Movers", 1_620_000, 13_000)]
URLS = [("https://www.exampleheating.com/", 900), ("https://www.exampleheating.com/furnace", 300)]


def _sheet(wb, title, header, rows, right_header=None, right_rows=()):
    """A tab with the combined block, and -- stitched -- a gap of empty
    columns, then "Source period" and the per-period block."""
    ws = wb.create_sheet(title)
    stitched = right_header is not None
    pad = [None] * 3
    ws.append(list(header) + (pad + ["Source period"] + list(right_header) if stitched else []))
    for i in range(max(len(rows), len(right_rows) if stitched else 0)):
        left = list(rows[i]) if i < len(rows) else [None] * len(header)
        right = list(right_rows[i]) if stitched and i < len(right_rows) else []
        ws.append(left + (pad + right if right else []))


def workbook(path, stitched=True, summary=True, empty=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if summary:
        ws = wb.create_sheet("Period Summary")
        ws.append(["Period", "Delivered Impressions", "Reported Unique Visitors", "Unique Visitors",
                   "Reported Attributed Impressions", "Attributed Impressions",
                   "Unique Visitor Rate", "Attributed Rate", "Impression Factor", "Visitor Factor"])
        for name, delivered, attributed in PERIODS:
            projected = name in ("June", "July")
            ws.append([name, delivered, 100, 180 if projected else 100,
                       attributed // 2 if projected else attributed, attributed, 0.001, 0.006,
                       1.7 if projected else 1, 1.8 if projected else 1])
        ws.append(["January–October total", DELIVERED, 400, 560, ATTRIBUTED - 1_100, ATTRIBUTED])
        ws.append([])
        ws.append(["June and July are projected using the supplied corrected totals."])
    totals = [] if empty else [(DELIVERED, ATTRIBUTED, ATTRIBUTED / DELIVERED)]
    right = (["Delivered Impressions", "Attributed Impressions", "Attributed Rate"], [
        (n, d, a, a / d) for n, d, a in PERIODS]) if stitched else (None, ())
    _sheet(wb, "Untitled Widget", ["Delivered Impressions", "Attributed Impressions", "Attributed Rate"],
           totals, *right)
    _sheet(wb, "Untitled Widget_1", ["Attributed Unique Visitors", "Attributed Unique Visitor Rate"],
           [] if empty else [(1_200, 1_200 / DELIVERED)])
    dim = ["Audience Name", "Delivered Impressions", "Attributed Impressions", "Attributed Rate",
           "Conversion Impressions", "Conversion Impressions Rate"]
    rows = [] if empty else [(n, d, a, a / d, 0, 0) for n, d, a in AUDIENCES]
    right = (dim, [(p[0], n, d // 4, a // 4, a / d, 0, 0) for p in PERIODS for n, d, a in AUDIENCES]) \
        if stitched else (None, ())
    _sheet(wb, "ATTRIBUTED RATE BY AUDIENCE", dim, rows, *right)
    _sheet(wb, "ATTRIBUTED UNIQUE PAGE VISITS B", ["Page Url", "Attributed Unique Visitors"],
           [] if empty else URLS)
    _sheet(wb, "WEBSITE ATTRIBUTION ADVERTISER",
           ["Client", "Advertiser Pixel", "Delivered Impressions", "Attributed Impressions",
            "Attributed Unique Visitors"],
           [] if empty else [(CLIENT, "TEGNA - WUSA - Example Heating Pixels", ATTRIBUTED, ATTRIBUTED,
                              1_200)])
    wb.save(path)
    return path


def summary(a):
    return (a.client_name, a.delivered_impressions, a.attributed_impressions,
            a.attributed_unique_visitors, [(r.label, r.delivered_impressions, r.attributed_impressions)
                                           for r in a.by_audience], dict(a.by_url))


def check_stitched_layout():
    print("1. Stitched layout; never a silent zero")
    stitched = ai.parse_attribution_export(str(workbook(TMP / "stitched.xlsx")))
    plain = ai.parse_attribution_export(str(workbook(TMP / "plain.xlsx", stitched=False,
                                                     summary=False)))
    check("the stitched workbook parses exactly like the plain export of the same figures",
          summary(stitched) == summary(plain), (summary(stitched), summary(plain)))
    check("...the combined figures, not a period's (3,240,000 / 25,800)",
          (stitched.delivered_impressions, stitched.attributed_impressions) == (DELIVERED, ATTRIBUTED))
    check("...client, audiences and URLs present", stitched.client_name == CLIENT
          and len(stitched.by_audience) == 2 and len(stitched.by_url) == 2)
    check("the Period Summary feeds no total, and its notes and total row aren't periods",
          [p["period"] for p in stitched.period_summary] == [p[0] for p in PERIODS],
          stitched.period_summary)
    try:
        ai.parse_attribution_export(str(workbook(TMP / "empty.xlsx", empty=True)), "Empty.xlsx")
        raised = None
    except ai.AttributionParseError as exc:
        raised = str(exc)
    check("core tabs that parse empty raise a rep-facing error naming them",
          raised is not None and '"Empty.xlsx"' in raised and '"Untitled Widget"' in raised
          and '"ATTRIBUTED RATE BY AUDIENCE"' in raised, raised)


def check_estimated_note():
    print("\n2. Projected periods are footnoted")
    a = ai.parse_attribution_export(str(workbook(TMP / "stitched.xlsx")))
    check("June and July are the estimated periods", ra.estimated_periods(a) == ["June", "July"])
    check("the note", ra.estimated_periods_note(a) == "June–July figures are estimated "
                                                       "(tracking gap).", ra.estimated_periods_note(a))
    plain = ai.parse_attribution_export(str(workbook(TMP / "plain.xlsx", stitched=False, summary=False)))
    check("an ordinary export: no note", ra.estimated_periods_note(plain) is None)
    check("non-consecutive periods are listed",
          ra.estimated_periods_note(ai.AttributionExport(period_summary=[
              {"period": "March", "estimated": True}, {"period": "July", "estimated": True}]))
          == "March and July figures are estimated (tracking gap).")
    facts = ra.build_facts_payload(a, None, client_name=CLIENT)
    check("the model is told which periods are estimates",
          facts["estimated_periods"]["periods"] == ["June", "July"])
    template = REPO / "REPORT_MASTER_v0_18.pptx"
    if not template.exists():
        print("  SKIP  the deck check -- no report template")
        return
    a.flight_start, a.flight_end = date(2026, 1, 1), date(2026, 10, 31)
    a.weekly_trend = [ai.DateSeriesPoint(date(2026, 1, 5) + timedelta(weeks=i), 75_000, 600, 0.008)
                      for i in range(43)]
    a.by_market = [ai.AttributionRow("WASHINGTON, DC (HAGRSTWN)", DELIVERED, ATTRIBUTED,
                                     ATTRIBUTED / DELIVERED)]
    a.by_zip = [ai.AttributionRow("20001", 400_000, 4_000, 0.01),
                ai.AttributionRow("20002", 300_000, 1_500, 0.005)]
    a.zip_heat = {"20001": 4_000, "20002": 1_500}
    out = TMP / "estimated.pptx"
    ra.build_report_deck(str(template), a, None, str(out), client_name=CLIENT,
                         goals_bullets=["More calls"], whats_next_bullets=["Keep going"])
    prs = Presentation(str(out))
    notes = {slide_map.notes_key(s): next((sh.text_frame.text for sh in s.shapes
                                           if sh.name == "ImpressionBasisNote"), "")
             for s in prs.slides}
    check("the recap, highlights and takeaways carry the note",
          all("June–July figures are estimated (tracking gap)." in notes.get(k, "")
              for k in ("report:recap", "report:highlights", "report:takeaways")), notes)
    out = TMP / "estimated_summary.pptx"
    ra.build_summary_slide(str(template), str(out), attribution=a, delivery=None, client_name=CLIENT,
                           threads=[], accepted_optimizations=[])
    text = " ".join(sh.text_frame.text for sh in Presentation(str(out)).slides[0].shapes
                    if sh.has_text_frame)
    check("...and the one-sheet", "June–July figures are estimated" in text, text[-300:])


def check_messages():
    print("\n3. Build failures name the cause")
    empty = ai.AttributionExport()
    msg = app.missing_info_message(ra.MissingTokenError("report:recap/CLIENT_NAME: x"), "the report",
                                   empty)
    check("a failed upload is named, not the form", "came through empty" in msg
          and "client name" not in msg, msg)
    ok = ai.AttributionExport(delivered_impressions=10)
    msg = app.missing_info_message(ra.MissingTokenError(
        "report:recap/GEOGRAPHY_LABEL: no by-market breakdown"), "the report", ok)
    check("otherwise the exact field", "missing: the geography" in msg
          and "report period" not in msg, msg)
    two = app.missing_info_message(ra.MissingTokenError(
        "report:recap/CLIENT_NAME and REPORT_PERIOD_LABEL"), "the report", ok)
    check("several fields, each named", "the client name" in two and "the report period" in two, two)


class _Query:
    def __init__(self, store, table):
        self.store, self.table, self.filters, self.row = store, table, [], None

    def select(self, *_a):
        return self

    def eq(self, col, val):
        self.filters.append((col, val))
        return self

    def limit(self, _n):
        return self

    def insert(self, row):
        self.row = row
        return self

    def execute(self):
        rows = self.store.setdefault(self.table, [])
        if self.row is not None:
            row = dict(self.row, id=f"id{len(rows) + 1}")
            assert all(r["name_key"] != row["name_key"] for r in rows), "unique index"
            rows.append(row)
            return type("R", (), {"data": [row]})
        return type("R", (), {"data": [r for r in rows if all(r.get(c) == v for c, v in self.filters)]})


class _Client:
    def __init__(self):
        self.store = {}

    def table(self, name):
        return _Query(self.store, name)


def check_clients():
    print("\n4. Existing clients: pick, or a separate record -- no renames")
    fake = _Client()
    real = db.get_client
    db.get_client = lambda: fake
    try:
        first, _ = db.create_advertiser(CLIENT)
        again, _ = db.create_advertiser(CLIENT.upper())
        second, err = db.create_advertiser(CLIENT, distinct=True)
        third, _ = db.create_advertiser(CLIENT, distinct=True)
    finally:
        db.get_client = real
    check("the same name again is the same client", again["id"] == first["id"])
    check("a confirmed different advertiser gets its own record, name unchanged",
          err is None and second["id"] != first["id"] and second["canonical_name"] == CLIENT
          and third["canonical_name"] == CLIENT and third["id"] not in (first["id"], second["id"]),
          fake.store)
    roster = [{"id": "a1", "canonical_name": CLIENT, "name_key": db.advertiser_name_key(CLIENT)},
              {"id": "a2", "canonical_name": "Someone Else", "name_key": "someone else"}]
    options = app.existing_client_options(CLIENT, roster, [])
    check("an exact-name client is offered even when the fuzzy matcher missed it",
          options and options[0]["id"] == "a1" and options[0]["score"] == 1.0, options)
    check("...and not twice when it was already a candidate",
          len(app.existing_client_options(CLIENT, roster, [{"id": "a1", "name": CLIENT, "score": 1.0}]))
          == 1)
    source = (REPO / "app.py").read_text(encoding="utf-8")
    check("the clash offers 'Use existing client', never 'give it a different name'",
          "Use existing client" in source and "give the new client a different name" not in source)


def main():
    check_stitched_layout()
    check_estimated_note()
    check_messages()
    check_clients()
    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
