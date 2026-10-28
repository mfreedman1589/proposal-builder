"""The QA ledger: qa_findings (Supabase) rendered as markdown.

    python qa_ledger.py          regenerate qa/ledger.md from the table
    python qa_ledger.py --gate   the nightly merge's QA gate: print any
                                 Critical/High finding still Open or Triaged
                                 and exit 1; exit 0 when nothing blocks

qa/ledger.md is a generated export -- never edited by hand. The table is the
source of truth: the QA agent files and verifies findings through the dev
app's QA page; development writes triage and fix status back.
"""
import sys
from pathlib import Path

LEDGER_PATH = Path(__file__).resolve().parent / "qa" / "ledger.md"
_COLUMNS = ("qa_id", "title", "severity", "status", "also_on_main", "triage_verdict", "fix_commit")


def _cell(value):
    return str(value if value not in (None, "") else "—").replace("|", "\\|").replace("\n", " ")


def ledger_markdown(rows, build_stamp=None):
    """The whole ledger: a summary table, then one card per finding."""
    lines = ["# QA ledger", "",
             "Generated from the `qa_findings` table -- do not edit by hand."
             + (f" Exported at build {build_stamp}." if build_stamp else ""), "",
             "| ID | Title | Severity | Status | Also on main | Verdict | Fix commit |",
             "|---|---|---|---|---|---|---|"]
    lines += ["| " + " | ".join(_cell(r.get(c)) for c in _COLUMNS) + " |" for r in rows]
    for r in rows:
        lines += ["", f"## {r.get('qa_id')} — {r.get('title') or ''}", ""]
        for label, key in (("Status", "status"), ("Severity", "severity"), ("Area", "area"),
                           ("Type", "finding_type"), ("Client-facing", "client_facing"),
                           ("Also on main", "also_on_main"),
                           ("Reproducibility", "reproducibility"), ("Rule cited", "rule_cited"),
                           ("Filed by", "created_by"), ("Build", "build_stamp"),
                           ("Filed", "created_at"), ("Verified", "verified_at")):
            lines.append(f"- **{label}:** {_cell(r.get(key))}")
        for label, key in (("Repro steps", "repro_steps"), ("Expected", "expected"),
                           ("Actual", "actual"), ("Triage", "triage_note")):
            if r.get(key):
                lines += ["", f"**{label}:**", "", str(r[key])]
        if r.get("triage_verdict"):
            lines.append(f"\n**Verdict:** {r['triage_verdict']}")
        if r.get("fix_commit"):
            lines.append(f"**Fix commit:** {r['fix_commit']}")
        history = history_lines(r)
        if history:
            lines += ["", "**History:**", ""] + [f"- {h}" for h in history]
    return "\n".join(lines) + "\n"


_FIELD_LABELS = {"finding_type": "Type", "severity": "Severity", "status": "Status",
                 "triage_verdict": "Verdict", "fix_commit": "Fix commit",
                 "also_on_main": "Also on main"}


def _cell(value):
    return ("" if value is None else str(value)).replace("|", "\\|").replace("\n", " ").strip()


def ledger_table(rows):
    """The QA page's summary table as markdown, every cell in full."""
    lines = ["| ID | Title | Severity | Status | Verdict | Triage note | Fix commit |",
             "|---|---|---|---|---|---|---|"]
    for r in rows:
        lines.append("| " + " | ".join(_cell(r.get(k)) for k in (
            "qa_id", "title", "severity", "status", "triage_verdict", "triage_note",
            "fix_commit")) + " |")
    return "\n".join(lines)


def history_lines(row):
    """A finding's history, oldest first, one line per entry:
    "2026-10-23 14:05 · Matt · Ruling: ..." / "... · Status: Open -> Fixed"."""
    out = []
    for entry in row.get("history") or []:
        when = str(entry.get("at") or "")[:16].replace("T", " ")
        who = entry.get("by") or "unknown"
        if entry.get("kind") == "note":
            what = f"Ruling / note: {entry.get('note')}"
        else:
            field = _FIELD_LABELS.get(entry.get("field"), entry.get("field"))
            what = f"{field}: {entry.get('from') or '—'} -> {entry.get('to') or '—'}"
        out.append(f"{when} · {who} · {what}")
    return out


def main(argv):
    import db
    rows, warning = db.fetch_qa_findings()
    if rows is None:
        print(f"Couldn't read qa_findings: {warning}")
        return 2
    if "--gate" in argv:
        blocking = db.merge_blocking_findings(rows)
        if blocking:
            print("QA gate: BLOCKED by " + ", ".join(blocking))
            return 1
        print("QA gate: clear (no Critical/High finding Open or Triaged)")
        return 0
    LEDGER_PATH.write_text(ledger_markdown(rows), encoding="utf-8")
    print(f"Wrote {LEDGER_PATH} ({len(rows)} findings)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
