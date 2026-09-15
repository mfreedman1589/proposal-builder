"""backfill_advertiser_verticals.py -- one-time migration (ATTRIBUTION_
REPORT_PLAN.md Phase 6, "category should come from the proposal"): every
advertiser whose own `vertical` column is still null gets it derived from
its linked proposals, via `db.derive_advertiser_vertical_from_proposals`
(the same function `link_proposal_advertiser` now calls automatically at
link time -- this script is only for advertisers whose proposals were all
linked BEFORE that hook existed).

--dry-run (the default): prints every advertiser with no vertical, what it
would derive to (or why it can't -- no linked proposals, or the linked
proposals disagree), and writes NOTHING.

--write: actually calls the derivation for each one.

Idempotent: an advertiser that already has a vertical is skipped by `derive_
advertiser_vertical_from_proposals` itself (never overwritten), so running
--write twice is a no-op the second time for everything the first run
reached.
"""
import argparse
import sys

import db


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--write", action="store_true",
                        help="Actually derive and write verticals. Default is a dry run.")
    args = parser.parse_args(argv)

    if not db.is_configured():
        print("SUPABASE_URL / SUPABASE_SERVICE_KEY not found -- check .streamlit/secrets.toml.")
        return 1

    advertisers, warning = db.fetch_advertisers(limit=10000, active_only=False)
    if warning:
        print(f"Couldn't load advertisers: {warning}")
        return 1
    unset = [a for a in (advertisers or []) if not a.get("vertical")]
    print(f"{len(advertisers or [])} advertiser(s) total -- {len(unset)} with no vertical set.")

    proposals, p_warning = db.fetch_proposals(limit=10000)
    if p_warning:
        print(f"Couldn't load proposals: {p_warning}")
        return 1
    by_advertiser = {}
    for row in proposals or []:
        aid = row.get("advertiser_id")
        if aid:
            by_advertiser.setdefault(aid, []).append(row.get("vertical"))

    print()
    print("=== Writing ===" if args.write else "=== Plan (dry run -- nothing written yet) ===")
    set_count, disagree_count, no_proposals_count = 0, 0, 0
    for advertiser in unset:
        aid = advertiser["id"]
        name = advertiser.get("canonical_name") or aid
        # "none" (a rep explicitly picking "no vertical" on the Build form)
        # is not signal to derive FROM -- same exclusion db.py's own
        # derive_advertiser_vertical_from_proposals applies.
        verticals = {v for v in by_advertiser.get(aid, []) if v and v != "none"}
        if not verticals:
            no_proposals_count += 1
            print(f'  "{name}" -- no linked proposals with a vertical, skipped')
            continue
        if len(verticals) > 1:
            disagree_count += 1
            print(f'  "{name}" -- proposals disagree ({", ".join(sorted(verticals))}), left unassigned')
            continue
        vertical = next(iter(verticals))
        set_count += 1
        print(f'  "{name}" -> {vertical}')
        if args.write:
            status, error = db.derive_advertiser_vertical_from_proposals(aid)
            if status != "set" or error:
                print(f"    !! didn't write as expected (status={status}): {error}")

    print()
    print(f"{set_count} would be set, {disagree_count} disagree (left unassigned), "
         f"{no_proposals_count} have no linked proposals to derive from.")
    if not args.write:
        print("Dry run only -- nothing written. Re-run with --write to apply.")
    else:
        print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
