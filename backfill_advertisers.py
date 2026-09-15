"""backfill_advertisers.py -- one-time migration (ATTRIBUTION_REPORT_PLAN.md
Phase 6, item 5): groups every proposal that has no `advertiser_id` yet by
`advertiser_matching.normalize_name(client_name)` -- the same fuzzy key the
live "Confirm the client" step on the Attribution Reports page scores
candidates against -- and resolves each group through `db.create_advertiser`
(which dedupes on the exact `name_key`, same as the live app) and
`db.link_proposal_advertiser`.

--dry-run (the default): prints the plan -- every proposal, which advertiser
it would join (existing or new), and a separate "can't place" list for a
blank or placeholder client_name (normalizes to nothing) -- and writes
NOTHING. Read this output before ever passing --write.

--write: actually creates advertisers and links proposals.

Idempotent: an already-linked proposal (`advertiser_id` already set) is
skipped outright on every run, so running --write twice is a no-op the
second time for everything the first run reached -- it only ever picks up a
proposal logged (or unlinked) since.

Never touches attribution_reports or case_studies -- those get their own
advertiser_id at the point they're created or tagged (the live app's
"Confirm the client" step, or the Clients page's "Tag an existing case
study"), not by this backfill.
"""
import argparse
import collections
import sys

import advertiser_matching
import db


# Generic placeholder text, not a real client name -- reported alongside the
# genuinely-blank case rather than silently becoming an advertiser literally
# named "Client". Real, current case: "Client" appears twice in the live
# table (a proposal built before a rep typed the real name in, or a demo
# run). Deliberately a short, exact-match denylist rather than a heuristic
# -- a real company that happens to BE named one of these is vanishingly
# unlikely, and a denylist that's wrong in the other direction (missing a
# placeholder) just leaves that proposal correctly linkable by hand from the
# Clients page, same as a genuinely blank name.
_PLACEHOLDER_KEYS = {"client", "clientname", "test", "testclient", "testing",
                     "na", "tbd", "todo", "sample", "demo"}


def group_proposals(rows):
    """(groups, unplaceable). `groups` is an insertion-ordered
    {normalized_key: [proposal_row, ...]}; `unplaceable` is every row whose
    client_name normalizes to nothing at all (blank, or punctuation-only) OR
    to a known placeholder (`_PLACEHOLDER_KEYS`) -- neither is a real
    advertiser name to create, so both are reported for a rep to place by
    hand rather than guessed at."""
    groups = collections.OrderedDict()
    unplaceable = []
    for row in rows:
        name = (row.get("client_name") or "").strip()
        key = advertiser_matching.normalize_name(name)
        if not key or key in _PLACEHOLDER_KEYS:
            unplaceable.append(row)
            continue
        groups.setdefault(key, []).append(row)
    return groups, unplaceable


def _row_line(row):
    return (f"    {row['id']}  {row.get('client_name')!r}  "
            f"(generated {str(row.get('generated_at') or '')[:10]})")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--write", action="store_true",
                        help="Actually create advertisers and link proposals. "
                             "Default is a dry run that writes nothing.")
    args = parser.parse_args(argv)

    if not db.is_configured():
        print("SUPABASE_URL / SUPABASE_SERVICE_KEY not found -- check .streamlit/secrets.toml.")
        return 1

    rows, warning = db.fetch_proposals(limit=10000)
    if warning:
        print(f"Couldn't load proposals: {warning}")
        return 1
    rows = rows or []

    already_linked = [r for r in rows if r.get("advertiser_id")]
    to_place = [r for r in rows if not r.get("advertiser_id")]
    print(f"{len(rows)} proposal(s) total -- {len(already_linked)} already linked (skipped), "
          f"{len(to_place)} to place.")

    groups, unplaceable = group_proposals(to_place)

    existing_advertisers, adv_warning = db.fetch_advertisers(limit=10000, active_only=False)
    if adv_warning:
        print(f"Couldn't load existing advertisers: {adv_warning}")
        return 1
    existing_by_key = {advertiser_matching.normalize_name(a.get("canonical_name") or ""): a
                       for a in (existing_advertisers or [])}

    print()
    print("=== Writing ===" if args.write else "=== Plan (dry run -- nothing written yet) ===")
    for key, group_rows in groups.items():
        sample_name = group_rows[0].get("client_name") or ""
        existing = existing_by_key.get(key)
        verb = "join existing advertiser" if existing else "create new advertiser"
        target_name = existing.get("canonical_name") if existing else sample_name
        print(f"\n\"{target_name}\" ({verb}) -- {len(group_rows)} proposal(s):")
        for row in group_rows:
            print(_row_line(row))
        if args.write:
            advertiser_row, error = db.create_advertiser(target_name)
            if error or not advertiser_row:
                print(f"    !! couldn't create/find advertiser: {error}")
                continue
            for row in group_rows:
                ok, link_error = db.link_proposal_advertiser(row["id"], advertiser_row["id"])
                if not ok:
                    print(f"    !! couldn't link {row['id']}: {link_error}")

    if unplaceable:
        print(f"\n=== Can't place ({len(unplaceable)}) ===")
        print("    Blank, punctuation-only, or a placeholder (e.g. \"Client\", \"Test\") -- "
              "link these by hand from the Clients page once a real advertiser exists for them.")
        for row in unplaceable:
            print(_row_line(row))

    print()
    if not args.write:
        print(f"Dry run only -- {sum(len(g) for g in groups.values())} proposal(s) across "
              f"{len(groups)} advertiser group(s) would be linked, "
              f"{len(unplaceable)} left unplaced. Re-run with --write to apply.")
    else:
        print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
