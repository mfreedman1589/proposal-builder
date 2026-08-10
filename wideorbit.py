"""wideorbit.py -- read a Wide Orbit broadcast schedule export into one shape.

Three formats come out of Wide Orbit and none of them agree on anything:

    Campaign Schedule Report (.xlsx)  week columns are M/D dates, impressions raw
    Planner (.xls)                    week columns are day-of-month, impressions thousands
    Planner (.pdf)                    week columns only in a summary table, impressions thousands

`parse_schedule(path)` returns a `Schedule` regardless, or raises
`ScheduleParseError` with something a seller can act on. Callers are expected
to catch it and fall back to typing the summary numbers in by hand -- a
broadcast schedule that can't be read must never be a broken page.

Two traps worth knowing before changing anything here:

**Impressions arrive in different units per format** -- 1,831.6 meaning
1,831,600 in the Planner exports, 1,785,413 meaning itself in the Campaign
Schedule Report. Everything is normalized to raw impressions, and where cost
and CPM are both present that relationship (imps = cost / cpm * 1000) is used
to decide which unit a number is in rather than trusting the format. Guessing
wrong here is a 1000x error in a client-facing media plan.

**The demo varies per schedule** -- A25-64, CS-A25+, A35+ -- and is carried
through to the deck. Never hardcode it; a schedule sold on one demo and
reported on another is a credibility problem, not a cosmetic one.

$0.00 rows are legitimate (added value / bonus weight) and are kept.
"""

import re
from dataclasses import dataclass, field
from datetime import date, datetime

# The Planner exports quote impressions in thousands. Both are sanity-checked
# against cost/CPM where possible, so this is a starting assumption rather
# than a rule the parser stakes anything on.
THOUSANDS_FORMATS = {"planner_xls", "planner_pdf"}

# How far off cost/cpm*1000 a figure may be and still be believed. Wide Orbit
# rounds CPM to cents, which at a $44 CPM is worth a few thousand impressions.
UNIT_TOLERANCE = 0.08


class ScheduleParseError(Exception):
    """Raised with a message intended for a seller, not a developer."""


@dataclass
class ScheduleRow:
    station: str = ""
    time: str = ""
    days: str = ""
    program: str = ""
    length: str = ""
    rate: float = 0.0
    spots_per_week: dict = field(default_factory=dict)   # {week_start(date): spots}
    total_spots: int = 0
    impressions: float = 0.0
    demo: str = ""

    @property
    def cost(self):
        return self.rate * self.total_spots


@dataclass
class ScheduleSummary:
    flight_start: date = None
    flight_end: date = None
    total_spots: int = 0
    gross_cost: float = 0.0
    grps: float = 0.0
    impressions: float = 0.0
    reach: float = 0.0
    frequency: float = 0.0
    cpm: float = 0.0
    demo_label: str = ""
    station: str = ""


@dataclass
class Schedule:
    rows: list = field(default_factory=list)
    summary: ScheduleSummary = field(default_factory=ScheduleSummary)
    source_format: str = ""
    source_name: str = ""
    notes: list = field(default_factory=list)
    # Every week column the export declared, including ones nobody bought.
    # Distinct from `weeks` on purpose: a schedule grid should show the
    # flight's shape, gaps included, while a totals calculation shouldn't
    # invent activity in an empty week.
    flight_weeks: list = field(default_factory=list)
    # How a schedule slide should label its week columns. A Campaign
    # Schedule Report headers its columns with real dates, so the deck
    # shows them; a Planner headers by position, so the deck says "Wk N"
    # rather than implying a precision the export didn't carry.
    week_header_style: str = "index"

    @property
    def weeks(self):
        """Week starts that actually carry spots, ascending."""
        seen = set()
        for row in self.rows:
            seen.update(row.spots_per_week)
        return sorted(seen)

    @property
    def grid_weeks(self):
        """The columns a schedule slide should show."""
        return self.flight_weeks or self.weeks


# ---------------------------------------------------------------------------
# shared helpers
# ---------------------------------------------------------------------------

def _clean(value):
    """Wide Orbit pads its header labels with non-breaking spaces."""
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").split())


def _money(value):
    """A dollar figure from any of the shapes WO writes it in."""
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = _clean(value).replace("$", "").replace(",", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except ValueError:
        return 0.0


def _number(value):
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = _clean(value).replace(",", "").replace("%", "").replace("$", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _label_value(text, label):
    """`Campaign:<nbsp padding>Ravens 2026-2027` -> `Ravens 2026-2027`."""
    cleaned = _clean(text)
    if not cleaned.lower().startswith(label.lower()):
        return ""
    return cleaned[len(label):].lstrip(":").strip()


def _parse_date_range(text):
    """(start, end) from `08/03/26-01/31/27` or `5/5/2025 - 6/1/2025`."""
    found = re.findall(r"(\d{1,2}/\d{1,2}/\d{2,4})", _clean(text))
    if len(found) < 2:
        return None, None
    return _parse_date(found[0]), _parse_date(found[1])


def _parse_date(text):
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(_clean(text), fmt).date()
        except ValueError:
            continue
    return None


def normalize_impressions(value, cost=None, cpm=None, assume_thousands=False):
    """Raw impressions from a figure that may be quoted in thousands.

    Where cost and CPM are both present they settle it: impressions must be
    cost / cpm * 1000, so whichever of `value` and `value * 1000` lands closer
    to that is the right reading. That check is what makes this safe across
    formats rather than a per-format guess -- and it catches the case where a
    format changes its own units between exports, which is not a hypothetical
    with Wide Orbit.
    """
    value = _number(value)
    if value <= 0:
        return 0.0

    # A $0.00 CPM is common (added value rows) and is truthy as a string, so
    # the guard has to be on the parsed number, not the raw cell.
    cost_value, cpm_value = _money(cost), _number(cpm)
    expected = (cost_value / cpm_value) * 1000 if cost_value > 0 and cpm_value > 0 else None

    if expected and expected > 0:
        raw_error = abs(value - expected) / expected
        thousands_error = abs(value * 1000 - expected) / expected
        if min(raw_error, thousands_error) <= UNIT_TOLERANCE:
            return value if raw_error <= thousands_error else value * 1000

    # Nothing to check against: fall back to what the format usually does.
    return value * 1000 if assume_thousands else value


def _week_key(month, day, year):
    try:
        return date(year, month, day)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Campaign Schedule Report (.xlsx)
# ---------------------------------------------------------------------------

def _parse_campaign_schedule_xlsx(path):
    import openpyxl

    book = openpyxl.load_workbook(path, data_only=True)
    sheet = book[book.sheetnames[0]]

    header = {}
    header_row = None
    for r in range(1, min(sheet.max_row, 40) + 1):
        for c in range(1, min(sheet.max_column, 6) + 1):
            text = _clean(sheet.cell(r, c).value)
            if not text:
                continue
            for label in ("Campaign", "Market(s)", "Spot Length(s)", "Campaign Dates",
                          "Book(s)", "AE/SR", "Exported"):
                value = _label_value(text, label)
                if value:
                    header.setdefault(label, value)
            if text == "Property":
                header_row = r
        if header_row:
            break

    if header_row is None:
        raise ScheduleParseError(
            "This looks like a Campaign Schedule Report but has no 'Property' header row, "
            "so the programs couldn't be located.")

    flight_start, flight_end = _parse_date_range(header.get("Campaign Dates", ""))

    # Week columns are M/D with no year; the campaign can straddle New Year
    # (a football schedule always does), so the year advances the first time
    # the month goes backwards rather than being taken from the start date.
    columns, trailing = {}, {}
    year = flight_start.year if flight_start else date.today().year
    previous_month = None
    for c in range(1, sheet.max_column + 1):
        text = _clean(sheet.cell(header_row, c).value)
        if not text:
            continue
        matched = re.fullmatch(r"(\d{1,2})/(\d{1,2})", text)
        if matched:
            month, day = int(matched.group(1)), int(matched.group(2))
            if previous_month is not None and month < previous_month:
                year += 1
            previous_month = month
            key = _week_key(month, day, year)
            if key:
                columns[c] = key
        elif text in ("Property", "Day/Time", "Program Name", "Book", "Length", "Rate",
                      "Spots", "Totals $", "Rtg", "CPP", "IMPs", "CPM"):
            trailing[text] = c

    rows, totals_row = [], None
    for r in range(header_row + 1, sheet.max_row + 1):
        first = _clean(sheet.cell(r, trailing.get("Property", 2)).value)
        marker = _clean(sheet.cell(r, trailing.get("Rate", 7)).value)
        if marker.lower() == "totals" or first.lower() == "totals":
            # Everything after the totals row is footer -- Comscore
            # attribution, the WideOrbit copyright line -- and those land in
            # the Property column, so they were being read as programs with
            # no spots. They tied to the totals (zero is zero) but showed up
            # as junk rows on the schedule slide.
            totals_row = r
            break
        if not first:
            continue
        row = ScheduleRow(
            station=first,
            time=_clean(sheet.cell(r, trailing.get("Day/Time", 3)).value),
            program=_clean(sheet.cell(r, trailing.get("Program Name", 4)).value),
            length=_clean(sheet.cell(r, trailing.get("Length", 6)).value),
            rate=_money(sheet.cell(r, trailing.get("Rate", 7)).value),
        )
        for c, week in columns.items():
            spots = _number(sheet.cell(r, c).value)
            if spots:
                row.spots_per_week[week] = int(spots)
        row.total_spots = int(_number(sheet.cell(r, trailing["Spots"]).value)) if "Spots" in trailing \
            else sum(row.spots_per_week.values())
        if "IMPs" in trailing:
            row.impressions = normalize_impressions(
                sheet.cell(r, trailing["IMPs"]).value,
                cost=sheet.cell(r, trailing["Totals $"]).value if "Totals $" in trailing else None,
                cpm=sheet.cell(r, trailing["CPM"]).value if "CPM" in trailing else None)
        rows.append(row)

    if not rows:
        raise ScheduleParseError("No program rows were found under the schedule's header row.")

    summary = ScheduleSummary(flight_start=flight_start, flight_end=flight_end,
                              station=rows[0].station if rows else "")
    if totals_row:
        summary.total_spots = int(_number(sheet.cell(totals_row, trailing["Spots"]).value)) \
            if "Spots" in trailing else sum(r.total_spots for r in rows)
        summary.gross_cost = _money(sheet.cell(totals_row, trailing["Totals $"]).value) \
            if "Totals $" in trailing else sum(r.cost for r in rows)
        summary.grps = _number(sheet.cell(totals_row, trailing["Rtg"]).value) if "Rtg" in trailing else 0.0
        summary.cpm = _money(sheet.cell(totals_row, trailing["CPM"]).value) if "CPM" in trailing else 0.0
        summary.impressions = normalize_impressions(
            sheet.cell(totals_row, trailing["IMPs"]).value if "IMPs" in trailing else 0,
            cost=summary.gross_cost, cpm=summary.cpm)
    else:
        summary.total_spots = sum(r.total_spots for r in rows)
        summary.gross_cost = sum(r.cost for r in rows)
        summary.impressions = sum(r.impressions for r in rows)

    _apply_general_summary_sheet(book, summary)
    for row in rows:
        row.demo = summary.demo_label
    return Schedule(rows=rows, summary=summary, source_format="campaign_schedule_xlsx",
                    flight_weeks=sorted(columns.values()),
                    week_header_style="date")


def _apply_general_summary_sheet(book, summary):
    """Reach/Freq/demo live on a second sheet in the Campaign Schedule Report."""
    name = next((n for n in book.sheetnames if "summary" in n.lower()), None)
    if not name:
        return
    sheet = book[name]
    header_row = None
    columns = {}
    for r in range(1, sheet.max_row + 1):
        labels = {_clean(sheet.cell(r, c).value): c for c in range(1, sheet.max_column + 1)}
        if "Spots" in labels and ("Imps" in labels or "Impressions" in labels):
            header_row, columns = r, labels
            break
    if header_row is None:
        return
    for r in range(header_row + 1, sheet.max_row + 1):
        name_cell = _clean(sheet.cell(r, columns.get("General", 2)).value)
        if not name_cell:
            continue
        # The `Total (A25-64)` row carries both the totals and the demo label.
        matched = re.search(r"total\s*\(([^)]+)\)", name_cell, re.I)
        if matched:
            summary.demo_label = matched.group(1).strip()
            if "Reach" in columns:
                summary.reach = _number(sheet.cell(r, columns["Reach"]).value)
            if "Freq" in columns:
                summary.frequency = _number(sheet.cell(r, columns["Freq"]).value)
            return
        if "Reach" in columns and not summary.reach:
            summary.reach = _number(sheet.cell(r, columns["Reach"]).value)
            summary.frequency = _number(sheet.cell(r, columns.get("Freq", 0)).value) \
                if "Freq" in columns else summary.frequency


# ---------------------------------------------------------------------------
# Planner (.xls)
# ---------------------------------------------------------------------------

def _parse_planner_xls(path):
    try:
        import xlrd
    except ImportError as exc:                                    # pragma: no cover
        raise ScheduleParseError(
            "Reading a .xls Planner export needs the `xlrd` package, which isn't "
            "installed. Re-export as .xlsx, or install xlrd.") from exc

    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)

    plan_start = plan_end = None
    station = demo = ""
    header_row = None
    for r in range(min(sheet.nrows, 30)):
        for c in range(sheet.ncols):
            text = _clean(sheet.cell_value(r, c))
            if not text:
                continue
            # The header row's own "Station" cell would otherwise be eaten by
            # the "Station:" label branch below, since one is a prefix of the
            # other -- so the grid check comes first.
            if text == "Station" and _clean(sheet.cell_value(r, c + 1)) == "Time":
                header_row = r
            elif text.lower().startswith("plan dates"):
                plan_start, plan_end = _parse_date_range(text)
            elif text.lower().startswith("station"):
                station = _label_value(text, "Station") or station
            elif text.lower().startswith("demo"):
                demo = (_label_value(text, "Demo") or "").split(",")[0].split()[0] if _label_value(text, "Demo") else demo
        if header_row is not None:
            break

    if header_row is None:
        raise ScheduleParseError(
            "This looks like a Planner export but has no Station/Time header row, "
            "so the programs couldn't be located.")

    labels = {_clean(sheet.cell_value(header_row, c)): c for c in range(sheet.ncols)}

    # Week columns are bare day-of-month numbers. The month advances whenever
    # the day goes backwards, which is what makes a flight spanning a month
    # boundary come out right.
    week_columns = {}
    if plan_start:
        year, month, previous_day = plan_start.year, plan_start.month, None
        for c in range(sheet.ncols):
            raw = sheet.cell_value(header_row, c)
            if not isinstance(raw, (int, float)) or raw <= 0 or raw > 31:
                continue
            day = int(raw)
            if previous_day is not None and day < previous_day:
                month += 1
                if month > 12:
                    month, year = 1, year + 1
            previous_day = day
            key = _week_key(month, day, year)
            if key:
                week_columns[c] = key

    rows, totals_row = [], None
    for r in range(header_row + 1, sheet.nrows):
        first = _clean(sheet.cell_value(r, labels.get("Station", 0)))
        if not first:
            continue
        if "total" in first.lower():
            totals_row = r
            continue
        row = ScheduleRow(
            station=first,
            time=_clean(sheet.cell_value(r, labels.get("Time", 1))),
            days=_clean(sheet.cell_value(r, labels.get("Days", 2))),
            program=_clean(sheet.cell_value(r, labels.get("Program Name", 3))),
            length=_clean(sheet.cell_value(r, labels.get("Length", 4))),
            rate=_money(sheet.cell_value(r, labels.get("Rate", 5))),
        )
        for c, week in week_columns.items():
            spots = _number(sheet.cell_value(r, c))
            if spots:
                row.spots_per_week[week] = int(spots)
        row.total_spots = int(_number(sheet.cell_value(r, labels["#"]))) if "#" in labels \
            else sum(row.spots_per_week.values())
        if "(000)" in labels:
            # The labelled column holds the per-spot audience; the unlabelled
            # column immediately right of it holds that row's extended total
            # (12.7 -> 304.8 across 24 spots). Summing the labelled one would
            # under-report a schedule by the spot count, which is a large and
            # silent error -- the extended column is what ties to FLIGHT
            # TOTALS. Same pairing applies to RTG and CPP.
            extended = labels["(000)"] + 1
            raw = sheet.cell_value(r, extended) if extended < sheet.ncols else None
            if not _number(raw):
                raw = sheet.cell_value(r, labels["(000)"])
            row.impressions = normalize_impressions(
                raw,
                cost=sheet.cell_value(r, labels["Total"]) if "Total" in labels else None,
                cpm=sheet.cell_value(r, labels["CPM"]) if "CPM" in labels else None,
                assume_thousands=True)
        rows.append(row)

    if not rows:
        raise ScheduleParseError("No program rows were found in this Planner export.")

    summary = ScheduleSummary(flight_start=plan_start, flight_end=plan_end,
                              station=station or rows[0].station, demo_label=demo)
    if totals_row is not None:
        summary.total_spots = int(_number(sheet.cell_value(totals_row, labels["#"]))) if "#" in labels \
            else sum(r.total_spots for r in rows)
        summary.gross_cost = _money(sheet.cell_value(totals_row, labels["Total"])) if "Total" in labels \
            else sum(r.cost for r in rows)
        summary.grps = _number(sheet.cell_value(totals_row, labels["RTG"])) if "RTG" in labels else 0.0
        summary.cpm = _number(sheet.cell_value(totals_row, labels["CPM"])) if "CPM" in labels else 0.0
        summary.impressions = normalize_impressions(
            sheet.cell_value(totals_row, labels["(000)"]) if "(000)" in labels else 0,
            cost=summary.gross_cost, cpm=summary.cpm, assume_thousands=True)
    else:
        summary.total_spots = sum(r.total_spots for r in rows)
        summary.gross_cost = sum(r.cost for r in rows)
        summary.impressions = sum(r.impressions for r in rows)

    _apply_planner_summaries_sheet(book, summary)
    for row in rows:
        row.demo = summary.demo_label
    return Schedule(rows=rows, summary=summary, source_format="planner_xls",
                    flight_weeks=sorted(week_columns.values()))


def _apply_planner_summaries_sheet(book, summary):
    """Reach/Frequency live on the Summaries sheet, not the flight sheet."""
    name = next((n for n in book.sheet_names() if "summar" in n.lower()), None)
    if not name:
        return
    sheet = book.sheet_by_name(name)
    for r in range(sheet.nrows):
        labels = {_clean(sheet.cell_value(r, c)): c for c in range(sheet.ncols)}
        if "Reach" in labels and "Freq" in labels:
            for rr in range(r + 1, sheet.nrows):
                if not _clean(sheet.cell_value(rr, 0)):
                    continue
                summary.reach = _number(sheet.cell_value(rr, labels["Reach"]))
                summary.frequency = _number(sheet.cell_value(rr, labels["Freq"]))
                break
            break
    if not summary.demo_label:
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                matched = re.search(r"summary\s*\(\s*([^)]+?)\s*\)",
                                    _clean(sheet.cell_value(r, c)), re.I)
                if matched:
                    summary.demo_label = matched.group(1).split()[0]
                    return


# ---------------------------------------------------------------------------
# Planner (.pdf)
# ---------------------------------------------------------------------------

def _parse_planner_pdf(path):
    import pdfplumber

    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    if not text.strip():
        raise ScheduleParseError("No text could be extracted from this PDF -- if it's a scan "
                                 "rather than a Wide Orbit export, enter the numbers by hand.")

    lines = [line.rstrip() for line in text.splitlines()]
    summary = ScheduleSummary()

    for line in lines:
        cleaned = _clean(line)
        if cleaned.lower().startswith("station:"):
            summary.station = summary.station or _label_value(cleaned.split("Phone")[0], "Station")
        if cleaned.lower().startswith("flight dates") or cleaned.lower().startswith("schedule date"):
            start, end = _parse_date_range(cleaned)
            summary.flight_start = summary.flight_start or start
            summary.flight_end = summary.flight_end or end
        matched = re.search(r"week summary\s*\(\s*([^)]+?)\s*\)", cleaned, re.I)
        if matched:
            summary.demo_label = matched.group(1).split()[0]
        matched = re.search(r"reach:\s*([\d.]+)", cleaned, re.I)
        if matched:
            summary.reach = float(matched.group(1))
        matched = re.search(r"frequency:\s*([\d.]+)", cleaned, re.I)
        if matched:
            summary.frequency = float(matched.group(1))
        matched = re.search(r"spts:\s*([\d,]+)", cleaned, re.I)
        if matched:
            summary.total_spots = int(matched.group(1).replace(",", ""))
        matched = re.search(r"total cost:\s*\$?([\d,.]+)", cleaned, re.I)
        if matched:
            summary.gross_cost = _money(matched.group(1))
        matched = re.search(r"cpp/cpm:\s*\$?([\d,.]+)", cleaned, re.I)
        if matched:
            summary.cpm = _money(matched.group(1))

    # The Week Summary's Total row carries spots, cost and impressions, and is
    # the most reliable thing on the page -- the program blocks come out of
    # pdfplumber with their columns interleaved.
    weeks = []
    for line in lines:
        cleaned = _clean(line)
        # Week Summary row: date, spots, %, cost, %, avg rate, imps(000), CPM.
        matched = re.match(r"^(\d{1,2}/\d{1,2}/\d{2,4})\s+(\d+)\s+\d+%\s+\$([\d,.]+)\s+\d+%"
                           r"(?:\s+\$[\d,.]+\s+([\d,.]+)\s+\$([\d,.]+))?", cleaned)
        if matched:
            week = _parse_date(matched.group(1))
            if week:
                weeks.append({"week": week, "spots": int(matched.group(2)),
                              "cost": _money(matched.group(3)),
                              "imps": matched.group(4), "cpm": matched.group(5)})
        matched = re.match(r"^total\s*\(([^)]+)\)\s+([\d,]+)\s+\d+%\s+\$([\d,.]+)\s+\d+%\s+"
                           r"\$[\d,.]+\s+([\d,.]+)\s+\$([\d,.]+)", cleaned, re.I)
        if matched:
            summary.demo_label = summary.demo_label or matched.group(1).strip()
            summary.total_spots = summary.total_spots or int(matched.group(2).replace(",", ""))
            summary.gross_cost = summary.gross_cost or _money(matched.group(3))
            summary.cpm = summary.cpm or _money(matched.group(5))
            summary.impressions = normalize_impressions(
                matched.group(4), cost=summary.gross_cost, cpm=summary.cpm,
                assume_thousands=True)
        matched = re.search(r"grp/\(000\)\s*([\d,.]+)", cleaned, re.I)
        if matched:
            summary.grps = _number(matched.group(1))

    if not summary.total_spots and not summary.gross_cost:
        raise ScheduleParseError(
            "This PDF doesn't look like a Wide Orbit Planner export -- no totals block was "
            "found. Enter the summary numbers by hand.")

    if not summary.impressions and summary.gross_cost and summary.cpm:
        summary.impressions = (summary.gross_cost / summary.cpm) * 1000

    # One synthetic row per week. The per-program blocks are recoverable but
    # pdfplumber interleaves their columns badly enough that anything derived
    # from them would be a guess; the Week Summary is exact, and the totals it
    # feeds are what the deck and the media plan line actually use.
    rows = []
    if weeks:
        for entry in weeks:
            week = entry["week"]
            rows.append(ScheduleRow(
                station=summary.station,
                # Not %-m/%-d: that's a platform-specific strftime extension
                # and raises on Windows.
                program=f"Week of {week.month}/{week.day}",
                length="", rate=(entry["cost"] / entry["spots"]) if entry["spots"] else 0.0,
                spots_per_week={entry["week"]: entry["spots"]},
                total_spots=entry["spots"], demo=summary.demo_label,
                impressions=normalize_impressions(
                    entry.get("imps"), cost=entry["cost"], cpm=entry.get("cpm"),
                    assume_thousands=True)))
        if not summary.flight_start:
            summary.flight_start = min(e["week"] for e in weeks)
    schedule = Schedule(rows=rows, summary=summary, source_format="planner_pdf",
                        flight_weeks=sorted(e["week"] for e in weeks))
    schedule.notes.append(
        "Read from a PDF, so the schedule grid is summarized by week rather than by program. "
        "The totals are exact; re-import the .xls or .xlsx export if you need program detail.")
    return schedule


# ---------------------------------------------------------------------------

def parse_schedule(path, filename=None):
    """Parse any supported Wide Orbit export. Raises ScheduleParseError."""
    name = (filename or str(path)).lower()
    try:
        if name.endswith(".xlsx"):
            schedule = _parse_campaign_schedule_xlsx(path)
        elif name.endswith(".xls"):
            schedule = _parse_planner_xls(path)
        elif name.endswith(".pdf"):
            schedule = _parse_planner_pdf(path)
        else:
            raise ScheduleParseError(
                f"'{filename or path}' isn't a file type this reads -- Wide Orbit exports "
                f"come as .xlsx, .xls or .pdf.")
    except ScheduleParseError:
        raise
    except Exception as exc:
        raise ScheduleParseError(
            f"Couldn't read this schedule ({type(exc).__name__}: {exc}). It may be an export "
            f"format this doesn't recognize yet -- enter the summary numbers by hand.") from exc

    schedule.source_name = filename or ""
    if not schedule.summary.cpm and schedule.summary.impressions:
        schedule.summary.cpm = schedule.summary.gross_cost / schedule.summary.impressions * 1000
    return schedule
