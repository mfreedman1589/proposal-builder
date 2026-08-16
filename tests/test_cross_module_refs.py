"""Every `module.NAME` between this project's own modules must resolve.

    python tests/test_cross_module_refs.py

The failure this covers: app.py referring to something assembly.py doesn't
define. `app.py` and `assembly.py` change together constantly, and the
reference is usually a bare constant -- `assembly.VERTICAL_ATTRIBUTION`,
`assembly.SCHEDULE_PLACEHOLDER_MARKER` -- that no test exercises directly.
Push one half without the other and the deployed app dies at import, because
several of these references sit at module scope: `DRAFT_KEY_SECTIONS.update(
{...for v in assembly.VERTICAL_ATTRIBUTION})` runs when app.py is imported,
so a missing name takes down every page rather than one feature.

Two deliberate choices about what counts:

  * Call-time references count as much as import-time ones. A name used
    inside a function is not "safer" -- it fails the moment that page
    renders, in front of whoever rendered it.

  * Modules are DISCOVERED, not listed. A hardcoded roster is the thing that
    silently stops covering a module somebody adds later, which is the same
    failure mode this file exists to prevent (and the same reasoning behind
    the font index being discovered rather than mapped).

Definitions are collected from module scope only -- recursing through
module-level try/except and if/else, which is how `create_client` and
`ClientOptions` are conditionally imported in db.py, but NOT into function
bodies. Counting locals would over-approximate and quietly stop catching
anything; missing a genuinely exotic definition would cry wolf. Neither is
free, and a guard that cries wolf is one people learn to skip.
"""
import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Not first-party application modules: entry points and one-off scripts are
# still scanned as *callers*, they just aren't expected to be imported by
# anything. Nothing is excluded from being checked.
SKIP_FILES = {"setup.py", "conftest.py"}


def module_scope_names(tree):
    """Names bound at module scope, following module-level control flow."""
    names = set()

    def visit(body):
        for node in body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                names.add(node.name)
            elif isinstance(node, ast.Assign):
                for target in node.targets:
                    for sub in ast.walk(target):
                        if isinstance(sub, ast.Name):
                            names.add(sub.id)
            elif isinstance(node, (ast.AnnAssign, ast.AugAssign)):
                if isinstance(node.target, ast.Name):
                    names.add(node.target.id)
            elif isinstance(node, (ast.Import, ast.ImportFrom)):
                for alias in node.names:
                    names.add(alias.asname or alias.name.split(".")[0])
            # Module-level control flow still binds module-level names.
            elif isinstance(node, ast.Try):
                visit(node.body)
                for handler in node.handlers:
                    visit(handler.body)
                visit(node.orelse)
                visit(node.finalbody)
            elif isinstance(node, (ast.If, ast.For, ast.While, ast.With)):
                visit(node.body)
                visit(getattr(node, "orelse", []))

    visit(tree.body)
    return names


def first_party_modules():
    """Every .py at the repo root, by module name."""
    return {p.stem: p for p in sorted(ROOT.glob("*.py"))
            if p.name not in SKIP_FILES}


def main():
    modules = first_party_modules()
    trees, defined = {}, {}
    for name, path in modules.items():
        try:
            trees[name] = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except SyntaxError as exc:
            print(f"  FAIL  {path.name} does not parse: {exc}")
            return 1
        defined[name] = module_scope_names(trees[name])

    print(f"{len(modules)} first-party module(s): {', '.join(sorted(modules))}\n")

    problems, checked = [], 0
    for caller, tree in sorted(trees.items()):
        # local name -> real module name. It has to be a MAPPING, not a set:
        # db.py does `import optimize_deck as optimizer`, and resolving
        # `optimizer.optimize_deck` against a module called "optimizer"
        # reports a missing name on code that is perfectly correct. A guard
        # that cries wolf is one people learn to skip, so the alias is
        # followed rather than assumed away.
        imported = {}
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    base = alias.name.split(".")[0]
                    if base in modules:
                        imported[alias.asname or base] = base

        for node in ast.walk(tree):
            if not (isinstance(node, ast.Attribute)
                    and isinstance(node.value, ast.Name)
                    and node.value.id in imported):
                continue
            target = imported[node.value.id]
            if target == caller:
                continue
            checked += 1
            if node.attr.startswith("__"):
                continue
            if node.attr not in defined.get(target, set()):
                problems.append((caller, node.lineno, target, node.attr))

    if problems:
        print(f"{len(problems)} UNRESOLVED cross-module reference(s):\n")
        for caller, line, target, attr in sorted(set(problems)):
            print(f"  FAIL  {caller}.py:{line}  ->  {target}.{attr}  "
                  f"(not defined in {target}.py)")
        print("\nOne half of a change is missing. Both files have to ship in "
              "the same commit, or the deployed app breaks on import.")
        return 1

    print(f"  PASS  {checked} cross-module reference(s) all resolve")
    return 0


if __name__ == "__main__":
    sys.exit(main())
