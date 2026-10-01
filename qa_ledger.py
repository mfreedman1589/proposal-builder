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
_COLUMNS = ("qa_id", "title", "severity", "status", "triage_verdict", "fix_commit")


def _cell(value):
    return str(value if value not in (None, "") else "—").replace("|", "\\|").replace("\n", " ")


def ledger_markdown(rows, build_stamp=None):
    """The whole ledger: a summary table, then one card per finding."""
    lines = ["# QA ledger", "",
             "Generated from the `qa_findings` table -- do not edit by hand."
             + (f" Exported at build {build_stamp}." if build_stamp else ""), "",
             "| ID | Title | Severity | Status | Verdict | Fix commit |",
             "|---|---|---|---|---|---|"]
    lines += ["| " + " | ".join(_cell(r.get(c)) for c in _COLUMNS) + " |" for r in rows]
    for r in rows:
        lines += ["", f"## {r.get('qa_id')} — {r.get('title') or ''}", ""]
        for label, key in (("Status", "status"), ("Severity", "severity"), ("Area", "area"),
                           ("Type", "finding_type"), ("Client-facing", "client_facing"),
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
    return "\n".join(lines) + "\n"


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
