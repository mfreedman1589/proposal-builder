"""Run every tests/test_*.py file and report one pass/fail/skip summary --
the full-sweep gate this project's own history shows a curated list can't
replace: a hand-picked "suites I thought were relevant" run has already
missed real regressions (the "Avail dates" column breaking three OTHER
suites' `st.data_editor` monkeypatches at once is the concrete example
that prompted this script) that only a run of literally everything would
have caught before a commit, not after.

    python tests/run_all.py                 # every tests/test_*.py (minus EXCLUDED)
    python tests/run_all.py avails           # only files whose name contains "avails"
    python tests/run_all.py --timeout 600    # per-file timeout in seconds (default 600)
    python tests/run_all.py --include-live   # also runs Tier 2 -- see EXCLUDED below

This is a SEPARATE command from Tier 1 (`test_draft_regression.py`) on
purpose, not a widening of it: Tier 1's whole value is being free, offline,
and (documented) ~1 minute, cheap enough to run before every push touching
drafting/assembly/the rate card. This sweep is neither fast nor uniformly
runnable -- several suites SKIP without gitignored real client PDFs, and
the whole run takes much longer than Tier 1 -- so it's a coarser, less
frequent gate: run it before starting a phase (a baseline) and before
declaring one done (the actual gate), not on every edit.

**EXCLUDED BY DEFAULT: `test_draft_live.py` (Tier 2).** It makes real,
paid Anthropic API calls and asserts only probabilistic model-layer
behaviour (its own docstring says to re-run once before concluding
anything) -- a routine, unattended sweep must never spend real money or
fail the gate on model drift that isn't a code regression. Tier 2 stays
its own explicit, occasional command per CLAUDE.md's own rule: run it
before pushing a change to the prompt, the draft schema, the audience
catalog, or the model version. `--include-live` overrides this for the
rare case someone wants it folded into one run anyway.

Each test file already runs standalone as `python tests/test_X.py` and
follows one convention this script leans on: exit 0 means "nothing this
file checked came back wrong" (which includes a graceful, printed
"SKIP -- reason" when a gitignored fixture is missing), and a non-zero
exit or an uncaught exception means a real failure. This script does not
import any test module -- each one runs in its own subprocess, exactly as
a human would invoke it, so one file's `sys.modules`/monkeypatches/session
state can never bleed into another's the way importing them all into one
process could (and, per DECISIONS.md's own AppTest findings, demonstrably
does for some of these suites).

A TIMEOUT is reported like any other failure, never silently swallowed --
found live on this script's first real run: Tier 1 alone took 4m27s
standalone on a loaded machine (this project's own docs say "~1 minute"),
so the default here is deliberately generous rather than tuned to the
fastest observed run. A file that times out even at a generous budget is
real information (`test_targeting_map.py` did, inconsistently, across
runs -- see DECISIONS.md/BACKLOG.md), not a reason to raise the number
again without looking at why.
"""
import argparse
import subprocess
import sys
import time
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
REPO = TESTS_DIR.parent

# name -> why it's excluded from the default sweep. Checked via `in`, so this
# is an exact filename match, never a substring guess.
EXCLUDED_BY_DEFAULT = {
    "test_draft_live.py": "Tier 2 -- real, paid Anthropic API calls; probabilistic "
                          "model-layer assertions, not a code regression gate. Run "
                          "explicitly (python tests/test_draft_live.py) before pushing "
                          "a prompt/schema/catalog/model change, per CLAUDE.md.",
}


def discover(pattern=None, include_live=False):
    files = sorted(TESTS_DIR.glob("test_*.py"))
    if not include_live:
        files = [f for f in files if f.name not in EXCLUDED_BY_DEFAULT]
    if pattern:
        pattern = pattern.lower()
        files = [f for f in files if pattern in f.name.lower()]
    return files


def run_one(path, timeout):
    start = time.time()
    try:
        proc = subprocess.run(
            [sys.executable, str(path)], cwd=str(REPO),
            capture_output=True, text=True, timeout=timeout)
        elapsed = time.time() - start
        output = proc.stdout + proc.stderr
        if proc.returncode != 0:
            return "FAIL", elapsed, output
        if any(line.strip().startswith("SKIP") for line in output.splitlines()):
            return "SKIP", elapsed, output
        return "PASS", elapsed, output
    except subprocess.TimeoutExpired as exc:
        elapsed = time.time() - start
        stdout = exc.stdout or ""
        stderr = exc.stderr or ""
        if isinstance(stdout, bytes):
            stdout = stdout.decode("utf-8", "replace")
        if isinstance(stderr, bytes):
            stderr = stderr.decode("utf-8", "replace")
        return "TIMEOUT", elapsed, stdout + stderr


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("pattern", nargs="?", default=None,
                        help="Only run files whose name contains this substring.")
    parser.add_argument("--timeout", type=float, default=600,
                        help="Per-file timeout in seconds (default 600).")
    parser.add_argument("--include-live", action="store_true",
                        help="Also run Tier 2 (test_draft_live.py) -- real API cost, "
                             "probabilistic assertions. Off by default; see this "
                             "file's own docstring.")
    args = parser.parse_args()

    files = discover(args.pattern, include_live=args.include_live)
    if not files:
        print(f"No test_*.py files matched {args.pattern!r}.")
        return 1

    if not args.include_live and not args.pattern:
        excluded_present = [f for f in EXCLUDED_BY_DEFAULT if (TESTS_DIR / f).exists()]
        if excluded_present:
            print("Excluded by default (see this script's docstring): "
                  + ", ".join(excluded_present))
            print("Pass --include-live to include them.\n")

    print(f"Running {len(files)} test file(s)...\n")
    results = []
    for i, path in enumerate(files, 1):
        print(f"[{i:>3}/{len(files)}] {path.name} ...", end=" ", flush=True)
        status, elapsed, output = run_one(path, args.timeout)
        print(f"{status} ({elapsed:.1f}s)")
        results.append((path.name, status, elapsed, output))

    passed = [r for r in results if r[1] == "PASS"]
    skipped = [r for r in results if r[1] == "SKIP"]
    failed = [r for r in results if r[1] in ("FAIL", "TIMEOUT")]

    print("\n" + "=" * 78)

    if skipped:
        print(f"\nSKIPPED ({len(skipped)}) -- gitignored fixtures, PowerPoint, or similar not present:")
        for name, _status, _elapsed, _output in skipped:
            print(f"  {name}")

    if failed:
        print(f"\nFAILED ({len(failed)}):")
        for name, status, _elapsed, _output in failed:
            print(f"  {name} [{status}]")
        for name, status, _elapsed, output in failed:
            print(f"\n--- {name} [{status}] (last 4000 chars) ---")
            print(output[-4000:])

    total_time = sum(r[2] for r in results)
    print(f"\n{len(passed)} passed, {len(failed)} failed, {len(skipped)} skipped "
          f"out of {len(results)} test files, {total_time:.0f}s total")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
