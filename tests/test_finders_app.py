"""The finders-only app (finders_app.py) -- Premion Finder for DC sellers.

Enforced here, not by hidden buttons:
1. It imports no builder module. A fresh interpreter imports finders_app and
   the only project modules it loads must be the shared, non-builder ones
   below -- app.py, assembly, slide_map, the report builder and the rest
   never load at all.
2. No write path is reachable. Every `db.` call in the modules it runs is
   checked against an allowlist: reads, plus the only writes this app may
   cause -- its own usage log, a "Report an issue" submission, and adding a
   name to the team list. No case-study, tag, audience or vault write.
3. The real app renders two pages and nothing else, with no edit control
   anywhere, behind its OWN password (FINDERS_PASSWORD -- the builder's
   APP_PASSWORD doesn't open it), and logs searches with the seller's name.
4. A case study with rendered slides downloads as a real PDF.

    python tests/test_finders_app.py
"""
import ast
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

failures = []

# The only project modules the finders app may load. Adding one is a
# decision: check it isn't builder code, then add it here.
ALLOWED_MODULES = {"app_shell", "audience_catalog", "catalog_shared", "claude_client", "db",
                   "finders", "finders_app", "market_profiles"}
BUILDER_MODULES = {"app", "assembly", "slide_map", "report_assembly", "attribution_import",
                   "analyst_import", "targeting_groups", "targeting_map", "avails_pdf_import",
                   "wideorbit", "deck_render", "optimize_deck", "text_metrics", "geo_resolver",
                   "slide_inheritance", "audience_usage_import"}
# The modules finders_app runs, and every db function they may call.
FINDERS_SOURCES = ("finders_app.py", "finders.py", "catalog_shared.py", "claude_client.py",
                   "app_shell.py")
DB_READS = {"fetch_case_studies", "case_study_cached_path", "case_study_file", "case_study_images",
            "case_study_render_path", "describe_error", "fetch_attribution_reports",
            "fetch_advertisers", "fetch_team_members", "deck_channel"}
DB_ALLOWED_WRITES = {"log_finder_usage", "submit_feedback", "add_team_member"}


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def check_imports():
    print("It imports no builder module")
    probe = (
        "import os, sys\n"
        "before = set(sys.modules)\n"
        "import finders_app\n"
        "here = os.getcwd()\n"
        "def local(m):\n"
        "    f = getattr(sys.modules[m], '__file__', None)\n"
        "    return bool(f) and os.path.dirname(os.path.abspath(f)) == here\n"
        "print(sorted(m for m in set(sys.modules) - before if local(m)))\n")
    result = subprocess.run([sys.executable, "-c", probe], cwd=REPO, capture_output=True,
                            text=True, env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    loaded = set(ast.literal_eval(result.stdout.strip().splitlines()[-1])) if result.stdout.strip() else None
    check("importing finders_app works", loaded is not None, result.stderr[-400:])
    if loaded is None:
        return
    check("no builder module loads (app, assembly, slide_map, the report builder...)",
          not (loaded & BUILDER_MODULES), sorted(loaded & BUILDER_MODULES))
    check("only the shared, reviewed modules load", loaded <= ALLOWED_MODULES,
          sorted(loaded - ALLOWED_MODULES))
    for name in FINDERS_SOURCES:
        tree = ast.parse((REPO / name).read_text(encoding="utf-8"))
        imported = {a.name.split(".")[0] for node in ast.walk(tree)
                    if isinstance(node, ast.Import) for a in node.names}
        imported |= {node.module.split(".")[0] for node in ast.walk(tree)
                     if isinstance(node, ast.ImportFrom) and node.module}
        check(f"{name} has no import of a builder module", not (imported & BUILDER_MODULES),
              sorted(imported & BUILDER_MODULES))


def check_write_paths():
    print("\nNo write path is reachable")
    for name in FINDERS_SOURCES:
        tree = ast.parse((REPO / name).read_text(encoding="utf-8"))
        calls = {node.attr for node in ast.walk(tree)
                 if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name)
                 and node.value.id == "db"}
        check(f"{name}: every db call is a read or an allowed write (usage log, feedback, "
              f"team name)", calls <= DB_READS | DB_ALLOWED_WRITES,
              sorted(calls - DB_READS - DB_ALLOWED_WRITES))
    finders_src = (REPO / "finders.py").read_text(encoding="utf-8")
    check("the shared finder code carries no case-study write at all",
          "update_case_study" not in finders_src and "upload_case_study" not in finders_src
          and "delete" not in finders_src.lower().replace("deleted", ""))


ROWS = [
    {"id": "cs1", "title": "HVAC & Plumbing - DC Sales Conversions", "filename": "hvac.pptx",
     "storage_path": "hvac.pptx", "summary": "Service calls up 30%.", "verticals": ["home_improvement"],
     "products": ["premion_streaming_tv"], "active": True, "slide_images": [], "date_added": "2026-10-05"},
    {"id": "cs2", "title": "Regional Bank - Lending Lift", "filename": "bank.pptx",
     "storage_path": "bank.pptx", "summary": "20% lift.", "verticals": ["banking"],
     "products": [], "active": True, "slide_images": [], "date_added": "2026-09-01"},
]


def new_app(test_mode=True):
    from streamlit.testing.v1 import AppTest
    if test_mode:
        os.environ["PROPOSAL_BUILDER_TEST_MODE"] = "1"
    else:
        os.environ.pop("PROPOSAL_BUILDER_TEST_MODE", None)
    return AppTest.from_file(str(REPO / "finders_app.py"), default_timeout=300)


def check_app():
    print("\nThe app itself: two pages, no edit controls, its own password, usage logged")
    import db
    logged = []
    real = (db.fetch_case_studies, db.case_study_cached_path, db.log_finder_usage,
            db.fetch_attribution_reports)
    db.fetch_case_studies = lambda active_only=True: (list(ROWS), None)
    db.case_study_cached_path = lambda *a, **k: None
    db.log_finder_usage = lambda event, user=None, detail=None, app="finders": logged.append(
        (event, user, detail, app)) or (True, None)
    db.fetch_attribution_reports = lambda *a, **k: ([], None)
    forbidden = {"Save changes", "Deactivate", "Reactivate", "➕ Add case study", "AND", "OR",
                 "New group"}
    try:
        at = new_app()
        at.run()
        check("renders without error", not at.exception, at.exception)
        nav = next((r for r in at.radio if r.key == "finders_page"), None)
        check("the sidebar offers exactly the two finders",
              nav is not None and list(nav.options) == ["Audience finder", "Case study finder"],
              nav and nav.options)
        check("titled Premion Finder, with a how-to line on the page",
              any("How to use" in str(c.value) for c in at.caption), [c.value for c in at.caption][:3])
        check("no builder or edit control on the audience page",
              not ({b.label for b in at.button} & forbidden), [b.label for b in at.button])
        at.text_input(key="finder_search").input("homeowner").run()
        check("an audience search is logged with the seller's name and the query",
              any(e == "audience_search" and u == "Test Mode" and d.get("query") == "homeowner"
                  and a == "finders" for e, u, d, a in logged), logged)

        nav.set_value("Case study finder").run()
        check("the case study page renders", not at.exception, at.exception)
        labels = {b.label for b in at.button}
        check("no Add / Save / Deactivate anywhere on the case study page",
              not (labels & forbidden), sorted(labels & forbidden))
        check("no 'Include deactivated' toggle -- sellers see active case studies only",
              not any(c.key == "vault_show_inactive" for c in at.checkbox))
        editor_prefixes = tuple(f"vault_{r['id']}_" for r in ROWS)
        editor_keys = [w.key for w in list(at.text_input) + list(at.text_area) + list(at.multiselect)
                       if (w.key or "").startswith(editor_prefixes)]
        check("no title/summary/tag editor fields for any case study", not editor_keys, editor_keys)
        check("each case study offers a download", any(b.label == "Get the case study" for b in at.button),
              sorted(labels))
        at.text_input(key="vault_search").input("hvac").run()
        check("a case study search is logged",
              any(e == "case_study_search" and d.get("query") == "hvac" for e, _u, d, _a in logged),
              logged)

        print("\nIts own password")
        gate = new_app(test_mode=False)
        gate.secrets["FINDERS_PASSWORD"] = "finders-pass"
        gate.secrets["APP_PASSWORD"] = "builder-pass"
        gate.run()
        check("the login screen is Premion Finder's", any(t.value == "Premion Finder" for t in gate.title),
              [t.value for t in gate.title])
        gate.text_input[0].input("builder-pass")
        gate.button[0].click().run()
        check("the builder's APP_PASSWORD does NOT open it",
              not gate.session_state["authed"] if "authed" in gate.session_state else True)
        gate.text_input[0].input("finders-pass")
        gate.button[0].click().run()
        check("FINDERS_PASSWORD does", "authed" in gate.session_state and gate.session_state["authed"])
    finally:
        (db.fetch_case_studies, db.case_study_cached_path, db.log_finder_usage,
         db.fetch_attribution_reports) = real
        os.environ.pop("PROPOSAL_BUILDER_TEST_MODE", None)


def check_pdf():
    print("\nA rendered case study downloads as a PDF")
    import db
    import finders
    from PIL import Image
    tmp = Path(tempfile.mkdtemp())
    paths = []
    for i, color in enumerate(("navy", "white")):
        path = tmp / f"slide_{i}.png"
        Image.new("RGB", (320, 180), color).save(path)
        paths.append(str(path))
    real = db.case_study_images
    db.case_study_images = lambda case_study_id, storage_paths: paths
    try:
        pdf = finders.case_study_pdf_bytes(dict(ROWS[0], slide_images=["a.png", "b.png"]))
    finally:
        db.case_study_images = real
    check("a real PDF, one page per slide", bool(pdf) and pdf.startswith(b"%PDF")
          and pdf.count(b"/Type /Page") - pdf.count(b"/Type /Pages") == 2, pdf[:8] if pdf else pdf)
    check("no PDF for a case study without rendered slides (it would be empty)",
          finders.case_study_pdf_bytes(ROWS[0]) is None)
    check("the PDF is named like the .pptx, from the title (not the source file's name)",
          finders.case_study_pdf_filename(ROWS[0]) == "HVAC_Plumbing_DC_Sales_Conversions.pdf",
          finders.case_study_pdf_filename(ROWS[0]))


def main():
    check_imports()
    check_write_paths()
    check_app()
    check_pdf()
    print()
    print(f"{len(failures)} failure(s)" if failures else "All checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
