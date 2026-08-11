"""Catch text that would be rendered into the app by Streamlit's magic.

    python tests/test_no_stray_magic.py

Streamlit rewrites any bare expression in the script into `st.write(...)`.
A string is exempt only when it is the FIRST statement of a module, class or
function -- `magic._is_docstring_node` tests `node_index == 0` and nothing
else. So a docstring stops being a docstring the moment anything is inserted
above it, and its text starts rendering to every user.

That is not hypothetical: `check_identity`'s design rationale -- "Not
authentication, the password is the gate..." -- shipped to the top of the
login screen because a test-mode early return was added above it. Nothing
failed, no test noticed, and the only symptom was several paragraphs of
internal prose on screen.

This applies the same rule Streamlit does, to every module that imports it.
Free, offline, no rendering required.
"""

import ast
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SCOPE = ("app.py", "db.py", "assembly.py", "audience_catalog.py", "wideorbit.py",
         "deck_render.py", "package_check.py", "text_metrics.py", "slide_map.py")

# Streamlit's own rule, from runtime/scriptrunner/magic.py.
DOCSTRING_PARENTS = (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)


def stray_expressions(path):
    """Bare expressions Streamlit's magic would turn into st.write()."""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    found = []
    for parent in ast.walk(tree):
        body = getattr(parent, "body", None)
        if not isinstance(body, list):
            continue
        for index, node in enumerate(body):
            if not isinstance(node, ast.Expr):
                continue
            # A call is a statement doing work (st.write, a mutation); only
            # values sitting on their own are what magic displays.
            if isinstance(node.value, (ast.Call, ast.Await, ast.Yield, ast.YieldFrom)):
                continue
            is_docstring = (index == 0 and isinstance(parent, DOCSTRING_PARENTS)
                            and isinstance(node.value, ast.Constant)
                            and isinstance(node.value.value, str))
            if is_docstring:
                continue
            found.append((node.lineno, ast.dump(node.value)[:60]))
    return found


def main():
    failures = []
    for name in SCOPE:
        path = REPO / name
        if not path.exists():
            print(f"  SKIP  {name} (not present)")
            continue
        stray = stray_expressions(path)
        if stray:
            failures.append(name)
            print(f"  FAIL  {name} would render {len(stray)} expression(s) into the app")
            for lineno, dump in stray:
                print(f"          line {lineno}: {dump}")
        else:
            print(f"  PASS  {name}")
    print()
    if failures:
        print(f"{len(failures)} file(s) carry text that renders into the UI: {failures}")
        print("A docstring that isn't the first statement is not a docstring. Move the "
              "text into # comments,\nor put it back at the top of its function.")
        return 1
    print("No stray expressions -- nothing renders into the app by accident.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
