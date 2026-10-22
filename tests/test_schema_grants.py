"""Every table in supabase_schema.sql grants Data API access to service_role.

From 2026-10-30 Supabase gives a NEW table in the public schema no Data API
access by default; without an explicit GRANT, every read and write through
the API fails with 42501 permission denied. Existing tables keep their
access, so the failure would only show up on the next new table -- after
its schema is pasted in, on first use. This reads supabase_schema.sql
itself and fails when a table is missing its grants:

- every CREATE TABLE has `grant select, insert, update, delete on table
  public.<t> to service_role` (any order of those four);
- every identity/serial id also grants usage + select on its sequence --
  either through pg_get_serial_sequence('public.<t>', 'id') or by naming
  public.<t>_id_seq;
- nothing is granted to anon or authenticated: the app uses the service
  key only and no table has a policy for them;
- RLS stays on for every table.

    python tests/test_schema_grants.py
"""
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SCHEMA = (REPO / "supabase_schema.sql").read_text(encoding="utf-8")

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def strip_comments(sql):
    return re.sub(r"--[^\n]*", "", sql)


SQL = strip_comments(SCHEMA)
TABLES = re.findall(r"create\s+table\s+(?:if\s+not\s+exists\s+)?public\.(\w+)\s*\((.*?)\n\);",
                    SQL, re.S | re.I)
PRIVILEGES = {"select", "insert", "update", "delete"}


def table_grant(name):
    for privs in re.findall(rf"grant\s+([\w\s,]+?)\s+on\s+table\s+public\.{name}\s+to\s+"
                            rf"service_role\s*;", SQL, re.I):
        if PRIVILEGES <= {p.strip().lower() for p in privs.split(",")}:
            return True
    return False


def sequence_grant(name):
    by_lookup = re.search(rf"grant\s+usage\s*,\s*select\s+on\s+sequence.*?pg_get_serial_sequence"
                          rf"\(\s*'public\.{name}'\s*,\s*'id'\s*\)", SQL, re.I | re.S)
    by_name = re.search(rf"grant\s+usage\s*,\s*select\s+on\s+sequence\s+public\.{name}_id_seq\s+"
                        rf"to\s+service_role", SQL, re.I)
    return bool(by_lookup or by_name)


def main():
    print(f"{len(TABLES)} tables in supabase_schema.sql")
    check("the schema declares tables at all (the parser found them)", len(TABLES) >= 17,
          [t for t, _ in TABLES])
    for name, body in TABLES:
        check(f"{name}: select/insert/update/delete granted to service_role", table_grant(name))
        if re.search(r"^\s*id\s+(?:big)?(?:serial|int\w*\s+generated\s+\w+\s+(?:by\s+default\s+)?"
                     r"as\s+identity)", body, re.I | re.M):
            check(f"{name}: identity id -> usage/select on its sequence", sequence_grant(name))
        check(f"{name}: RLS enabled",
              re.search(rf"alter\s+table\s+public\.{name}\s+enable\s+row\s+level\s+security",
                        SQL, re.I) is not None)
    # The grantee list follows the last "to"; "public." is the schema, not the role.
    grantees = re.findall(r"\bgrant\b[^;]*?\bto\s+([\w\s,]+?)\s*;", SQL, re.I)
    loose = [g for g in grantees if re.search(r"\b(anon|authenticated|public)\b", g, re.I)]
    check("nothing is granted to anon, authenticated or public", not loose, loose)
    print(f"\n{'ALL PASS' if not failures else f'{len(failures)} FAILED'}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
