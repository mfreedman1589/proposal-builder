"""avails_pdf_import.py -- read a Salesforce avails PDF (Premion Media Plan
export) into one shape: a header (advertiser, agency, flight, attribution)
plus one row per audience-geography pair, summed from whatever detail rows
the document gives.

Pure -- no Streamlit, no DB, no network geocoding. `parse_avails_pdf(path)`
returns an `AvailsDocument` or raises `AvailsParseError` with a message aimed
at a seller, the same convention as `wideorbit.py`. Geography is
CLASSIFIED and its raw values (a market name, a zip list, a radius
origin/miles, a county list) are extracted here; RESOLVING that geography
into `resolved_zips`/`resolved_markets` is deliberately left to the caller
(`app.resolve_group_geography`, the one place that already owns that job) --
this module only ever reports what the document says, never what a zip
resolves to.

**Structure, confirmed against six real documents**: an RFPID line, a
validity notice, `Media Plan Details` (agency, advertiser, flight dates,
total impressions, attribution), then one or more `Product Summary` /
`Product Details` / `Zip Codes` blocks -- one block per audience, a new block
starting wherever `Product Summary` next appears -- then T&Cs and signatures.
A block's `Product Details` table carries one row per REPORTING PERIOD for
that audience's one geography: one row when the document gives only a
flight total, twelve when it breaks the flight into calendar months. Every
row in one block shares the same Geography Included and Audience Target
text -- asserted, not assumed.

**Each Product Details row carries its own Start Date/End Date, not only
Impressions** -- kept as `AvailsGroup.periods` (see that class), one
`AvailsPeriod` per row, alongside the already-existing summed `impressions`/
`row_count`. A single-row block's one period spans the document's own
overall flight exactly; a monthly-broken-out block's periods are real
calendar months except the first/last, clipped wherever the flight starts
or ends mid-month. This is real, per-period data the document already
states -- never apportioned, never inferred from a calendar-month split of
the total.

**Geography classifies into four forms, prefix-first, DMA as the fallback**:
a bare name ("Philadelphia", "Washington, D.C.") is a DMA -- the only
unprefixed form, so it has to be tried last, not matched as a fourth
pattern. `Zip Option - <name> Add-On Zips` is a named zip option. Radius is
also a `Zip Option -`, distinguished by the word "radius"/"Mile Radius" in
its name; the bracketed origin `[<zip>]` is OPTIONAL -- one real sample
carries it, another doesn't, so its absence is not an error. `County Option
- <name>` is a county option; `<name>` is a human label ("New Jersey PA and
Delaware Counties"), never the real semicolon-separated `NAME STATE` list,
which lives in the Zip Codes block and has to be extracted separately.

**Parse by layout, not by line.** `extract_text()` interleaves a `Zip Codes`
block badly once it spans more than one page: the zip list, the RFPI
repeats and (for a county option) the semicolon county list all wrap
independently and land jumbled on the same visual line. Two techniques,
neither of them line-based:
  - The zip list is extracted with a bare `\\d{5}` search over the WHOLE
    block's text. Safe against the RFPI numbers next to it -- those are
    7-8 digits, and `(?<!\\d)\\d{5}(?!\\d)` never matches a 5-digit
    substring of a longer run -- so no column logic is needed for this one.
  - The county list can't use that trick (it's words, not digits), so it's
    read from WORD POSITIONS: words are grouped into visual lines, lines
    are clustered by their OWN leading word's x-position (the real column
    each line's wrapped text belongs to, not any single word's x), and the
    narrowest/leftmost cluster -- which also carries the repeated RFPI
    tokens -- has those RFPI tokens stripped back out, leaving the
    semicolon county fragments in reading order.

**Totals tie exactly**, and that is the assertion this module leads with:
every group's impressions, summed, must equal the header's own stated Total
Impressions -- external ground truth, never a total this module computes for
itself. A parse that drops or double-counts a row fails this immediately.
"""
import re
from dataclasses import dataclass, field
from datetime import date, datetime

import pdfplumber

import targeting_groups as tg

GEO_KIND_DMA = "dma"
GEO_KIND_NAMED_ZIP = "named_zip"
GEO_KIND_RADIUS = "radius"
GEO_KIND_COUNTY = "county"

_ZIP_RE = re.compile(r"(?<!\d)\d{5}(?!\d)")
_RFPI_TOKEN_RE = re.compile(r"^RFPI$|^-?\d{6,9},?$", re.IGNORECASE)
_BRACKET_ORIGIN_RE = re.compile(r"\[([^\]]+)\]")
_RADIUS_MILES_RE = re.compile(r"(\d+(?:\.\d+)?)\s*[Mm](?:i\b|ile)")


class AvailsParseError(Exception):
    """Raised with a message intended for a seller, not a developer."""


@dataclass
class AvailsPeriod:
    """One Product Details row's own reporting period -- a REAL date range
    the document itself states, not a calendar month inferred from
    anything. A single-RFPI product's one period spans the document's own
    overall flight exactly (confirmed against Annapolis Cars, Lawn &
    Leisure, Plaza Motors Group and Hershey's twelve single-row blocks --
    every one of them carries a Start Date/End Date identical to the
    document header's own Flight Start/End Date); a monthly-broken-out
    product's periods are real calendar months except (usually) the first
    and last, which are clipped to wherever the flight actually starts/ends
    mid-month (Capital Media: 09/21-09/30, 10/01-10/31, 11/01-11/30,
    12/01-12/20). `impressions` is THIS period's own stated figure, never
    apportioned from the group total -- real months in the same document
    can and do carry different impressions (Wilmington's five audiences
    each vary month to month), so summing/dividing would silently discard
    real data the document already gives."""
    start: date = None
    end: date = None
    impressions: int = 0


@dataclass
class AvailsGroup:
    """One audience-geography pair -- the unit `targeting_groups.new_group`
    is built from, one-to-one. `impressions` is the FULL-FLIGHT total,
    summed verbatim from the document's own detail rows (one row if the
    document gives only a flight total, twelve if it breaks the flight into
    calendar months -- summed either way, never re-derived from a rate).
    `periods` is the same detail rows kept individually rather than only
    summed -- see `AvailsPeriod`. Empty only when a future document's
    Product Details table drops the Start Date/End Date columns this was
    built against; every real document parses at least one period per
    group."""
    audience_text: str = ""
    geo_kind: str = ""
    geo_raw: str = ""              # the "Geography Included" cell, verbatim
    geo_name: str = ""             # the extracted display value (a market name, an option's <name>)
    county_list: str = ""          # GEO_KIND_COUNTY only -- "NAME ST;NAME ST;..."
    zips: list = field(default_factory=list)   # GEO_KIND_NAMED_ZIP only -- the document's own list
    radius_miles: float = None     # GEO_KIND_RADIUS only
    radius_origin: str = ""        # GEO_KIND_RADIUS only -- "" when the bracket is absent
    impressions: int = 0
    row_count: int = 0
    periods: list = field(default_factory=list)   # AvailsPeriod, document row order


@dataclass
class AvailsDocument:
    rfpid: str = ""
    advertiser: str = ""
    agency: str = ""
    flight_start: date = None
    flight_end: date = None
    total_impressions: int = 0
    attribution_text: str = ""     # raw blob -- app.py keyword-matches it against known products
    groups: list = field(default_factory=list)   # AvailsGroup, audience-major, document order
    source_name: str = ""


def _clean_int(text):
    try:
        return int(str(text).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0


def _clean_date(text):
    try:
        return datetime.strptime(str(text).strip(), "%m/%d/%Y").date()
    except (TypeError, ValueError):
        return None


def _table_field(page0, label):
    """One "Label / Value" cell pair from page 0's own table extraction --
    used for Advertiser and Agency, whose values can wrap onto a second
    physical line the flat extract_text() regex approach can't recover (the
    wrap lands between two DIFFERENT table cells -- "Advertiser Website" or
    "Total Impressions" among them -- not after it). A table cell preserves
    a wrap as an embedded newline rather than losing it to the next field
    entirely, which is exactly what's needed here. Returns "" when the
    label isn't found as a table cell -- not every field is reliably a
    clean table cell (Total Impressions' own value drops out of table
    extraction in every real sample checked), which is why this is used
    narrowly, not as the header parser's general strategy."""
    for table in page0.extract_tables():
        for row in table:
            if row and str(row[0] or "").strip() == label and len(row) > 1:
                return str(row[1] or "").replace("\n", " ").strip()
    return ""


def _field(text, label, stop_labels):
    """The value after `label` up to whichever of `stop_labels` comes next,
    all on the assumption the document prints "Label Value Label2 Value2"
    with no line breaks the parser can rely on. Returns "" rather than
    raising when the label itself isn't found -- a genuinely absent field
    (blank Advertiser Website, no CoSeller) is reported, not fatal."""
    m = re.search(re.escape(label) + r"\s+(.*)", text)
    if not m:
        return ""
    rest = m.group(1)
    cut = len(rest)
    for stop in stop_labels:
        idx = rest.find(stop)
        if idx != -1:
            cut = min(cut, idx)
    return rest[:cut].strip()


def _parse_header(page0, source_name):
    text = page0.extract_text() or ""
    if "Media Plan Details" not in text:
        raise AvailsParseError(
            "This doesn't look like a Premion avails PDF -- no 'Media Plan Details' "
            "section on the first page.")
    rfpid_m = re.search(r"RFPID-\d+", text)
    advertiser = _table_field(page0, "Advertiser") or _field(text, "Advertiser", ["Advertiser Website"])
    # _table_field, not the flat-text _field -- an agency name can wrap onto
    # a second physical line the same way Advertiser's can (real find,
    # 2026-09-23: "TBC, Inc - Trahan, Burden & Charles" wraps after the
    # comma in both real documents that carry it, RFPID-266583 and
    # RFPID-268846), and the flat regex silently truncates at the wrap same
    # as it would for Advertiser.
    agency = _table_field(page0, "Agency") or _field(text, "Agency", ["Total Impressions"])
    # Tolerates an optional dash between the label and the number -- most
    # real documents print "Total Impressions 2,522,716" with plain
    # whitespace, but RFPID-268846 (2026-09-23) prints "Total Impressions -
    # 763,712,864", and the old whitespace-only pattern didn't match either
    # a hyphen or an en dash, which aborted header parsing before ANY field
    # (including Advertiser, which itself extracts fine) ever got returned.
    total_m = re.search(r"Total Impressions\s*[-–]?\s*([\d,]+)", text)
    start_m = re.search(r"Flight Start Date\s+(\d{2}/\d{2}/\d{4})", text)
    end_m = re.search(r"Flight End Date\s+(\d{2}/\d{2}/\d{4})", text)
    if not (start_m and end_m and total_m):
        raise AvailsParseError(
            "Couldn't find Flight Start/End Date or Total Impressions on the "
            "Media Plan Details page -- this may not be a standard Premion avails export.")

    # The Attribution / Attribution Notes block is the one part of this page
    # whose line-wrapping isn't consistent across real samples (it wraps
    # differently depending on how many products are listed), so it's read
    # by LINE position rather than a single regex: everything between the
    # CoSeller line and the line that starts "Notes" (which is also where
    # a present-but-empty Attribution Notes field ends up).
    lines = text.splitlines()
    attribution_lines = []
    in_block = False
    for line in lines:
        if line.startswith("CoSeller"):
            in_block = True
            continue
        if in_block and line.startswith("Notes"):
            break
        if in_block:
            attribution_lines.append(line)
    attribution_text = " ".join(attribution_lines).replace("Attribution", "").strip()

    return AvailsDocument(
        rfpid=rfpid_m.group(0) if rfpid_m else "",
        advertiser=advertiser,
        agency=agency,
        flight_start=_clean_date(start_m.group(1)),
        flight_end=_clean_date(end_m.group(1)),
        total_impressions=_clean_int(total_m.group(1)),
        attribution_text=attribution_text,
        source_name=source_name,
    )


def classify_geography(geo_included):
    """(kind, name, radius_miles, radius_origin) from one "Geography
    Included" cell. Prefix first, DMA (bare) as the fallback -- see module
    docstring for why the order matters."""
    text = str(geo_included or "").strip()
    if text.lower().startswith("zip option -"):
        name = text.split("-", 1)[1].strip()
        if "radius" in name.lower():
            origin_m = _BRACKET_ORIGIN_RE.search(name)
            miles_m = _RADIUS_MILES_RE.search(name)
            miles = float(miles_m.group(1)) if miles_m else None
            origin = origin_m.group(1).strip() if origin_m else ""
            return GEO_KIND_RADIUS, name, miles, origin
        name = re.sub(r"\s*Add-?On Zips\s*$", "", name, flags=re.IGNORECASE).strip()
        return GEO_KIND_NAMED_ZIP, name, None, ""
    if text.lower().startswith("county option"):
        # The dash-and-name suffix is OPTIONAL, confirmed against a real
        # document (Capital Media RFPID-266994): its "Geography Included"
        # cell is the bare "County Option", no "- <name>" at all -- the
        # Zip Codes table's own GEO Target Name for that same block is a
        # literal "County Option - null", so this document's own export
        # never had a real label to give it. Requiring the dash misread
        # this as a bare DMA named "County Option" (falling through to the
        # fallback below), which fails to match any real market AND skips
        # `_parse_block`'s COUNTY branch entirely -- silently dropping the
        # block's own zip list (never populated because neither the
        # NAMED_ZIP/RADIUS nor COUNTY branch fired), which is why the
        # targeting map drew nothing for a document that plainly has real
        # zip codes in it.
        name = text.split("-", 1)[1].strip() if "-" in text else ""
        if name.lower() == "null":
            name = ""
        return GEO_KIND_COUNTY, name, None, ""
    return GEO_KIND_DMA, text, None, ""


def _find_product_details_table(tables):
    """The one table (of however many `extract_tables()` returned for a
    page) whose header row names Geography Included, Audience Target and
    Impressions -- column COUNT and exact header wording vary between real
    samples (one carries a "Package" column the others don't, and the
    header's own first cell -- nominally "RFPI" -- comes back None on at
    least one real sample's table-line detection), so this locates columns
    by NAME, never by position, and never requires the RFPI header cell
    specifically to be present."""
    for table in tables:
        if not table or not table[0]:
            continue
        header = [str(c or "").strip() for c in table[0]]
        if ("Geography Included" in header and "Audience Target" in header
                and "Impressions" in header):
            return header, table[1:]
    return None, None


def _column(header, *names):
    for name in names:
        for i, h in enumerate(header):
            if h == name:
                return i
    return None


def _split_into_blocks(pdf):
    """Page indexes grouped into one list per audience-geography block, a
    new block starting on every page whose text contains "Product Summary".
    Pages before the first "Product Summary" (the Media Plan Details page)
    and from "Premion Terms & Conditions" onward belong to no block."""
    blocks = []
    current = None
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        if "Product Summary" in text:
            current = [i]
            blocks.append(current)
        elif "Premion Terms & Conditions" in text:
            current = None
        elif current is not None:
            current.append(i)
    return blocks


def _county_fragment(pdf, page_indexes):
    """The semicolon-separated county list for one block, read by word
    position across however many pages the block's Zip Codes section spans.

    Grouping every word sharing a physical row into one "line" is NOT
    enough: the RFPI/county-fragment column and the zip column sometimes
    wrap to the same physical row (their independent wraps land on the same
    y-coordinate by coincidence), so a row can genuinely mix both columns'
    text -- confirmed on a real sample, where trusting the row's own
    leading word alone pulled a run of zip codes into the county string.
    So each row is split into RUNS first, wherever the gap between two
    consecutive words' x-position exceeds a threshold comfortably bigger
    than the space between two words in the SAME run (~10-20pt here) and
    comfortably smaller than the gap between two genuinely different
    columns (~100pt+ here) -- then only the runs whose own leading edge is
    closest to the leftmost edge seen anywhere in the block are kept. That
    leftmost run is also where the repeated "RFPI -NNNNNNN," tokens live
    (the county fragments fill the space left over on those same lines once
    a block's RFPIs run out), so those tokens are stripped back out before
    joining what's left.
    """
    RUN_GAP = 40
    runs = []   # (lead_x, [word texts in reading order])
    for position, page_i in enumerate(page_indexes):
        page = pdf.pages[page_i]
        # Only the FIRST page of a block carries the "Zip Codes" section
        # title -- every later page in the block is pure continuation, with
        # no title to find, so it has to be included in full rather than
        # skipped for lacking one.
        y0 = -1
        if position == 0:
            header_word = None
            for w in page.extract_words():
                if w["text"] == "Zip" and header_word is None and w["x0"] < 60:
                    # The section TITLE, not the table's own "Included Zip
                    # Codes" column header -- the title is a lone word at
                    # the left margin; the column header is not.
                    header_word = w
            if header_word is None:
                continue  # no Zip Codes section on this page at all
            y0 = header_word["top"]
        by_top = {}
        for w in page.extract_words():
            if w["top"] <= y0:
                continue
            by_top.setdefault(round(w["top"]), []).append(w)
        for top in sorted(by_top):
            ws = sorted(by_top[top], key=lambda w: w["x0"])
            run = [ws[0]]
            for w in ws[1:]:
                if w["x0"] - run[-1]["x1"] > RUN_GAP:
                    runs.append((run[0]["x0"], [x["text"] for x in run]))
                    run = [w]
                else:
                    run.append(w)
            runs.append((run[0]["x0"], [x["text"] for x in run]))
    if not runs:
        return ""

    leftmost = min(lead_x for lead_x, _ in runs)
    fragment_words = []
    for lead_x, texts in runs:
        if abs(lead_x - leftmost) > 8:
            continue
        for text in texts:
            if _RFPI_TOKEN_RE.match(text):
                continue
            fragment_words.append(text)
    return " ".join(fragment_words).strip()


def _block_zip_list(pdf, page_indexes):
    """Every 5-digit zip in a block's Zip Codes section -- see module
    docstring for why a bare digit search is safe here.

    A wrapped zip table's CONTINUATION page has no repeated "Zip Codes"
    header of its own -- gating every page on that literal string (the
    original code) silently dropped every page after the first one that
    actually carries it, on a real document long enough to wrap: Wilmington
    University's 5 county blocks (real geography, live client data) read 48
    of a real 303 zips this way, discarding the very entries a real gap
    report flagged as "out of county" (17527, 17555, 18015, 18042 -- all
    genuinely present, just on the dropped page). Once the header is seen
    on ANY page in the block, every remaining page in that same block is
    assumed to be its continuation and contributes too -- true for every
    real document seen so far (a block never resumes a DIFFERENT section
    after its Zip Codes table; `_split_into_blocks` would have started a
    new block at the next "Product Summary" if it did). Latent everywhere
    else only because no other real block, in any of the 4 real documents
    on hand, happens to span more than one page -- this is a general
    parser correctness fix, not scoped to County Option.
    """
    zips = []
    started = False
    for page_i in page_indexes:
        text = pdf.pages[page_i].extract_text() or ""
        if "Zip Codes" in text:
            text = text[text.index("Zip Codes"):]
            started = True
        elif not started:
            continue
        zips.extend(_ZIP_RE.findall(text))
    # Order-preserving de-dup -- a zip can legitimately repeat across a
    # block's wrapped lines if the export itself repeats it, but the
    # document never intends a duplicate count.
    seen = set()
    out = []
    for z in zips:
        if z not in seen:
            seen.add(z)
            out.append(z)
    return out


def _parse_block(pdf, page_indexes):
    tables = []
    for i in page_indexes:
        tables.extend(pdf.pages[i].extract_tables())
    header, rows = _find_product_details_table(tables)
    if header is None:
        raise AvailsParseError(
            f"Couldn't find a Product Details table on page {page_indexes[0] + 1} -- "
            f"this document's layout doesn't match what this importer expects.")

    geo_i = _column(header, "Geography Included")
    aud_i = _column(header, "Audience Target")
    imp_i = _column(header, "Impressions")
    if geo_i is None or aud_i is None or imp_i is None:
        raise AvailsParseError(
            f"The Product Details table on page {page_indexes[0] + 1} is missing a "
            f"column this importer needs (Geography Included / Audience Target / Impressions).")

    geo_values = {str(r[geo_i] or "").strip() for r in rows if r}
    aud_values = {str(r[aud_i] or "").strip() for r in rows if r}
    if len(geo_values) > 1:
        raise AvailsParseError(
            f"The Product Details table on page {page_indexes[0] + 1} names more than one "
            f"geography for what should be a single audience-geography block: {geo_values}.")
    if len(aud_values) > 1:
        raise AvailsParseError(
            f"The Product Details table on page {page_indexes[0] + 1} names more than one "
            f"audience for what should be a single audience-geography block: {aud_values}.")

    geo_raw = next(iter(geo_values), "")
    audience_text = next(iter(aud_values), "")
    impressions_total = sum(_clean_int(r[imp_i]) for r in rows if r)
    kind, name, miles, origin = classify_geography(geo_raw)

    # Start Date/End Date: the SAME Product Details rows already read for
    # geography/audience/impressions, kept individually rather than only
    # summed. Confirmed present, same two names, in every real document on
    # hand (Capital Media, Wilmington, Hershey, Annapolis, Lawn & Leisure,
    # Plaza Motors) -- start_i/end_i is None only for a hypothetical future
    # export that drops them, in which case periods is left empty rather
    # than raising: nothing downstream requires it (yet), the same
    # "reported, not fatal" stance _field() already takes on an absent
    # field. Sorted by start date defensively -- real documents already
    # print rows chronologically, but nothing guarantees a future one does.
    start_i = _column(header, "Start Date")
    end_i = _column(header, "End Date")
    periods = []
    if start_i is not None and end_i is not None:
        periods = sorted(
            (AvailsPeriod(start=_clean_date(r[start_i]), end=_clean_date(r[end_i]),
                          impressions=_clean_int(r[imp_i]))
             for r in rows if r),
            key=lambda p: (p.start is None, p.start))

    group = AvailsGroup(
        audience_text=audience_text, geo_kind=kind, geo_raw=geo_raw, geo_name=name,
        radius_miles=miles, radius_origin=origin,
        impressions=impressions_total, row_count=len(rows), periods=periods,
    )
    if kind in (GEO_KIND_NAMED_ZIP, GEO_KIND_RADIUS):
        # Radius gets the raw zip list too, not just a named option -- when
        # the bracketed origin is absent (a real sample does this; see
        # classify_geography), there is no center left to re-resolve a
        # radius from, so the caller falls back to the document's own zip
        # list directly, the same thing a rep would be forced to do by hand
        # with no address to type into the Radius mode's own expander.
        group.zips = _block_zip_list(pdf, page_indexes)
    elif kind == GEO_KIND_COUNTY:
        # A "County Option" block is zip-originated too, same as Zip
        # Option -- a rep enters zips, Premion's system resolves them to
        # county names and reports BOTH back (confirmed directly against a
        # real document: every County Option page carries a genuine Zip
        # Codes table alongside the county summary). `county_list` is
        # Premion's own DERIVED, lossy summary -- real, useful as a label,
        # but never the resolution input; `zips` is the rep's actual,
        # authoritative geography and is what the caller resolves from
        # (see app.apply_avails_import's GEO_KIND_COUNTY branch). Both are
        # captured; only one is load-bearing.
        group.county_list = _county_fragment(pdf, page_indexes)
        group.zips = _block_zip_list(pdf, page_indexes)
    return group


def parse_avails_pdf(path, source_name=None):
    """The whole document: header fields plus one AvailsGroup per
    audience-geography block, in document order (audience-major, since
    that's the order the documents already use).

    Raises `AvailsParseError` on anything that doesn't match the expected
    Premion avails layout, always with a message a seller could act on --
    never a bare traceback. Asserts nothing itself; the header-total check
    belongs to the caller (see tests/test_avails_pdf_import.py), because a
    caller may legitimately want the partial result even when the totals
    don't tie, to show the seller what WAS read.
    """
    source_name = source_name or str(path)
    try:
        with pdfplumber.open(path) as pdf:
            document = _parse_header(pdf.pages[0], source_name)
            blocks = _split_into_blocks(pdf)
            if not blocks:
                raise AvailsParseError(
                    "No 'Product Summary' blocks found -- this document doesn't have "
                    "any audience/geography detail to import.")
            document.groups = [_parse_block(pdf, block) for block in blocks]
            return document
    except AvailsParseError:
        raise
    except Exception as exc:  # pdfplumber/table-extraction failures, corrupt file, etc.
        raise AvailsParseError(
            f"Couldn't read this PDF as a Premion avails export: {exc}") from exc


# ---------------------------------------------------------------------------
# FLOW_REWORK_PLAN.md Phase 3: deterministic entity inference. Provable
# containment ONLY -- offline, no LLM, no network -- so this is one more
# pure fact about the document, exactly like geography classification
# above. It never fabricates an entity NAME (that's `entity_label`, always
# blank here, left for a rep or a draft's `group_entities` to fill in) and
# never merges rows/lines -- it only assigns `entity_id`, which changes how
# a stated budget divides and how a later merge aggregates avails (see
# targeting_groups.py's module docstring). Two rules, both verified against
# the real documents on hand, not guessed:
#
#   R1 -- nested geography under one audience: the SAME audience term set,
#   the SAME radius origin, a DIFFERENT radius (Annapolis Cars: 8 rows, 4
#   audiences x 2 radii each around one origin zip -> 4 entities), or one
#   zip-originated group's raw zip list a strict subset of another's under
#   the same audience.
#
#   R2 -- nested audience under one geography: identical geography, term
#   sets differing by exactly one term, where that one term is an "All"
#   DEMO Age bracket on one side and the SAME range's Male or Female
#   bracket on the other (Plaza Motors Group: "(DEMO Age A35-64) AND (AUTO
#   Type Luxury)" / "(AUTO Type Luxury) AND (DEMO Age M35-64)" -> 1
#   entity). **Honesty check done while building this, not skipped**: the
#   plan's own assumption was that this bracket table would be read off the
#   real audience catalog. It can't be -- the local brochure catalog (369
#   segments) has NO "DEMO Age" family at all; Plaza's own two segments
#   aren't in it. So this rule is scoped to exactly the one real pattern
#   actually observed (A/M/W over an identical numeric range), not a wider
#   convention this project has any catalog evidence for. Extending it
#   needs a new real document to verify against, not a guess.
#
# Anything else -- two rows sharing geography that satisfy neither rule
# (Wilmington's three undergraduate audiences, whose overlap the plan doc
# itself calls "unknowable" from the document alone) -- gets ONE aggregate
# note for the whole document, never one per row, and stays ungrouped:
# commit 9's drafting-based `group_entities` is the only mechanism that can
# tie those together, because it's the only one allowed to use meaning
# rather than a provable structural fact.
# ---------------------------------------------------------------------------
_AGE_BRACKET_RE = re.compile(r"^DEMO Age ([AMW])(\d+-\d+)$")


def _age_bracket_subset(term_a, term_b):
    """True when `term_a` is the "All" DEMO Age bracket and `term_b`
    narrows it to Male or Female over the SAME numeric range -- e.g.
    "DEMO Age A35-64" subsumes "DEMO Age M35-64". See this section's own
    module-level comment for why this is scoped to exactly this one
    verified pattern, not a broader table."""
    ma, mb = _AGE_BRACKET_RE.match(term_a), _AGE_BRACKET_RE.match(term_b)
    if not ma or not mb:
        return False
    return ma.group(1) == "A" and mb.group(1) in ("M", "W") and ma.group(2) == mb.group(2)


def _r1_nested_geography(g1, g2, terms1, terms2):
    """R1: same audience, one geography nested inside the other."""
    if frozenset(terms1) != frozenset(terms2):
        return False
    if (g1.geo_kind == GEO_KIND_RADIUS and g2.geo_kind == GEO_KIND_RADIUS
            and g1.radius_origin and g1.radius_origin == g2.radius_origin
            and g1.radius_miles != g2.radius_miles):
        return True
    if (g1.geo_kind == g2.geo_kind and g1.geo_kind in (GEO_KIND_NAMED_ZIP, GEO_KIND_COUNTY)
            and g1.zips and g2.zips):
        set1, set2 = set(g1.zips), set(g2.zips)
        if set1 != set2 and (set1 <= set2 or set2 <= set1):
            return True
    return False


def _r2_nested_audience(g1, g2, terms1, terms2):
    """R2: same geography, one audience an age-bracket narrowing of the
    other (see this section's own module-level comment)."""
    geo_key = lambda g: (g.geo_kind, g.geo_name, g.radius_origin, tuple(sorted(g.zips)))
    if geo_key(g1) != geo_key(g2):
        return False
    s1, s2 = set(terms1), set(terms2)
    if s1 == s2:
        return False
    shared = s1 & s2
    diff1, diff2 = list(s1 - shared), list(s2 - shared)
    if len(diff1) != 1 or len(diff2) != 1:
        return False
    return _age_bracket_subset(diff1[0], diff2[0]) or _age_bracket_subset(diff2[0], diff1[0])


def infer_entities(document):
    """One entity key per group in `document.groups` (parallel list; `None`
    means "no grouping inferred -- stays its own entity"), plus a list of
    plain-language notes for whatever this pass genuinely couldn't resolve
    -- for the caller to route to `unresolved_internal` (seller checks
    alone, never a client-facing claim).

    Pure and deterministic: the same document always infers the same
    entities, no LLM call, no randomness. Groups a component together only
    on R1/R2 above; connects entities transitively (if A pairs with B under
    R1 and B pairs with C under R2, all three share one entity) via a
    simple union-find, since a document's own group count is always small.
    """
    groups = document.groups
    terms_per_group = [tg.terms_from_audience_text(g.audience_text)[0] for g in groups]

    parent = list(range(len(groups)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[max(ri, rj)] = min(ri, rj)

    for i in range(len(groups)):
        for j in range(i + 1, len(groups)):
            if (_r1_nested_geography(groups[i], groups[j], terms_per_group[i], terms_per_group[j])
                    or _r2_nested_audience(groups[i], groups[j], terms_per_group[i], terms_per_group[j])):
                union(i, j)

    components = {}
    for i in range(len(groups)):
        components.setdefault(find(i), []).append(i)

    entity_keys = [None] * len(groups)
    for root, members in components.items():
        if len(members) > 1:
            key = f"entity_{root}"
            for i in members:
                entity_keys[i] = key

    # "Otherwise" case: rows sharing geography that satisfy neither rule --
    # ONE aggregate note for the whole document, never one per row (a real
    # prior bug this project already fixed once for a different note class
    # -- see tests/test_avails_import_notes.py).
    geo_key = lambda g: (g.geo_kind, g.geo_name, g.radius_origin, tuple(sorted(g.zips)))
    by_geo = {}
    for i, g in enumerate(groups):
        if entity_keys[i] is None:
            by_geo.setdefault(geo_key(g), []).append(i)
    ambiguous = sum(len(members) for members in by_geo.values() if len(members) > 1)
    notes = []
    if ambiguous:
        notes.append(
            f"{ambiguous} row(s) share geography with another row on this document but "
            f"couldn't be tied to one entity automatically -- group them by hand below the "
            f"avails table if they're really the same real-world thing.")
    return entity_keys, notes


# LiveWell's real shape: several radius rows sharing one audience, told
# apart only by a full street address in geo_name ("277 S Washington St
# Alexandria VA 22314 5 Mile Radius"). Labeling-only -- this is a DIFFERENT,
# narrower job than classify_geography's own radius_origin extraction
# (_BRACKET_ORIGIN_RE above), which feeds real geocoding/resolved_zips and
# only recognizes a bracketed origin like Annapolis's "[21401]". That gap is
# real and deliberately NOT touched here (BACKLOG.md) -- a label is display
# text a rep can fix in one click if it's ever wrong; resolved_zips feeds
# real reach numbers, and deserves its own, more careful fix. The two
# happen to read from the same geo_name string, which is coincidence, not
# shared machinery.
_STREET_SUFFIX_RE = re.compile(
    r"\b(?:St|Ave|Blvd|Dr|Rd|Ln|Way|Pl|Ct|Pkwy|Hwy|Cir|Ter|Trl|Sq)\b\.?", re.IGNORECASE)
_DIRECTIONAL_RE = re.compile(r"^(?:NE|NW|SE|SW|N|S|E|W)\b\.?\s*", re.IGNORECASE)
_STATE_ZIP_RE = re.compile(r"\b([A-Z]{2})\s+(\d{5})\b")


def _infer_entity_label(geo_kind, geo_name):
    """A best-effort, deterministic entity label from one row's own
    geo_name -- or None when nothing reliable extracts. See the module-
    level comment above for what this is (and, on purpose, is not).

    DMA: the geo_name IS already a clean, short label ("Philadelphia",
    "Baltimore") -- used verbatim, no parsing at all.

    radius/named_zip with a real street address: the city is the text
    between the LAST recognized street suffix (St/Ave/Blvd/...) and the
    trailing STATE+ZIP, with a leading directional token (NW/NE/SW/SE/
    N/S/E/W) stripped -- "277 S Washington St Alexandria VA 22314 5 Mile
    Radius" -> "Alexandria"; "945 Florida Ave NW Washington DC 20001 5
    Mile Radius" -> "Washington" (the directional sits between the suffix
    and the city here, not before the street name, which is why it's
    stripped from the FRONT of the extracted span, not the whole string).

    Everything else has no city text in it to extract at all, and
    correctly returns None: a bracketed-zip radius (Annapolis's "10mi
    radius [21401]" -- no address, just a zip), a county option's own
    human label, or a name Salesforce repeats identically across every
    row in the document (Wilmington, Plaza) -- none of these carry a
    place name this function could find even in principle. The audience
    already tells rows like that apart; that's not this function's job.

    Deliberately does NOT fall back to the raw address when the city
    can't be cleanly isolated -- a street address is exactly the content
    a plan-table label must not carry, the same rule geo_label() itself
    already follows for the Geo cell.
    """
    if geo_kind == GEO_KIND_DMA:
        name = str(geo_name or "").strip()
        return name or None
    if geo_kind not in (GEO_KIND_RADIUS, GEO_KIND_NAMED_ZIP):
        return None
    text = str(geo_name or "").strip()
    state_zip_matches = list(_STATE_ZIP_RE.finditer(text))
    if not state_zip_matches:
        return None
    state_zip = state_zip_matches[-1]
    suffix_matches = list(_STREET_SUFFIX_RE.finditer(text[:state_zip.start()]))
    if not suffix_matches:
        return None
    city_text = text[suffix_matches[-1].end():state_zip.start()].strip(" .,-")
    city_text = _DIRECTIONAL_RE.sub("", city_text).strip(" .,-")
    if not city_text or any(ch.isdigit() for ch in city_text):
        return None
    return city_text


def infer_entity_labels(document):
    """One label per group in `document.groups` (parallel list; `None`
    where nothing reliable extracts) -- pure and deterministic, same shape
    as `infer_entities`.

    A label that TWO OR MORE rows would extract identically is dropped
    from ALL of them, not assigned to any -- three different DC-area
    radii that each reduce to "Washington" is exactly this (found on the
    real LiveWell document). A collision means the extraction genuinely
    cannot tell those rows apart, and confidently mislabeling three
    different locations with the one name a client would read as "these
    are the same place" is worse than leaving them blank for a rep to
    name from what they actually know. Blank is always safe; guessed-but-
    wrong is not.
    """
    candidates = [_infer_entity_label(g.geo_kind, g.geo_name) for g in document.groups]
    counts = {}
    for label in candidates:
        if label:
            counts[label] = counts.get(label, 0) + 1
    return [label if label and counts[label] == 1 else None for label in candidates]
