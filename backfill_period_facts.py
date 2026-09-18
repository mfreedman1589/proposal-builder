"""backfill_period_facts.py -- one-time backfill of `report_json["period_
facts"]` for every attribution report logged before the cross-month
evidence work shipped (2026-09-18).

Re-parses each report's ORIGINAL export from a local file matching its own
stored `attribution.source_name` -- the raw AttributionRow data was never
stored before this shipped, so there is no DB-only reconstruction (see
`report_assembly.period_facts_for_report`'s own docstring). A report whose
source file isn't present locally is skipped and named at the end, not
guessed at -- the standing project rule that a fixture's identity comes
from the file itself, never a name in a doc or a report_json blob that
might not still match what's on disk.

    python backfill_period_facts.py            -- dry run, prints a plan
    python backfill_period_facts.py --write     -- actually updates Supabase

Filename resolution: the export's own `source_name` first (an exact match
in the project root); a small manual map below for the handful of real
files that were renamed/re-exported on disk since a report was logged.
"""
import sys
from pathlib import Path

import attribution_import as ai
import db
import report_assembly as ra

REPO = Path(__file__).resolve().parent

# source_name (as stored in a report's own report_json) -> the real file on
# disk today, when they've drifted (a re-export, a rename). Update this
# when a backfill run reports "no local file" for a report whose data you
# know still exists under a different name.
FILENAME_OVERRIDES = {
}


def _find_export_file(source_name):
    if not source_name:
        return None
    exact = REPO / source_name
    if exact.exists():
        return exact
    override = FILENAME_OVERRIDES.get(source_name)
    if override:
        path = REPO / override
        if path.exists():
            return path
    return None


def main():
    write = "--write" in sys.argv
    rows, warning = db.fetch_attribution_reports()
    if warning:
        print(f"Couldn't load attribution reports: {warning}")
        return 1
    rows = rows or []
    print(f"{len(rows)} report(s) logged.\n")

    done, skipped, already = [], [], []
    for row in rows:
        report_json = row.get("report_json") or {}
        if report_json.get("period_facts"):
            already.append(row)
            continue
        attribution_dict = report_json.get("attribution") or {}
        source_name = attribution_dict.get("source_name")
        client_name = attribution_dict.get("client_name") or "(unknown client)"
        path = _find_export_file(source_name)
        if path is None:
            skipped.append((row, client_name, source_name))
            continue
        attribution_obj = ai.parse_attribution_export(str(path))
        period_facts = ra.period_facts_for_report(attribution_obj)
        report_json["period_facts"] = period_facts
        done.append((row, client_name, source_name))
        if write:
            ok, error = db.update_attribution_report_json(row["id"], report_json)
            if not ok:
                print(f"  FAILED to write {row['id']} ({client_name}): {error}")

    print(f"Already had period_facts: {len(already)}")
    print(f"\n{'Backfilled' if write else 'Would backfill'} ({len(done)}):")
    for row, client_name, source_name in done:
        print(f"  {row['id']}  {client_name}  <-  {source_name}")

    print(f"\nSkipped, no local file found ({len(skipped)}):")
    for row, client_name, source_name in skipped:
        print(f"  {row['id']}  {client_name}  wants '{source_name}'")

    if not write and done:
        print("\nDry run only -- re-run with --write to commit these updates.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
