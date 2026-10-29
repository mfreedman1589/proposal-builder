"""The nightly merge's CI gate: has GitHub Actions passed on dev's current commit?

    python ci_status.py      exit 0 only when the latest CI run for dev
                             is on dev's HEAD and succeeded; else print why, exit 1

Reads the public GitHub API without a token (the repo is public), so it
needs no login. A run that is still going, or one on an older commit, is
not a pass: wait for it, or push, and run this again.
"""
import json
import subprocess
import sys
import urllib.error
import urllib.request

REPO = "mfreedman1589/proposal-builder"
RUNS_URL = (f"https://api.github.com/repos/{REPO}/actions/workflows/ci.yml/runs"
            "?branch=dev&event=push&per_page=1")


def main():
    head = subprocess.run(["git", "rev-parse", "dev"], capture_output=True,
                          text=True, check=True).stdout.strip()
    request = urllib.request.Request(RUNS_URL, headers={"Accept": "application/vnd.github+json"})
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            runs = json.load(response).get("workflow_runs", [])
    except urllib.error.HTTPError as exc:
        print(f"CI: GitHub answered {exc.code} -- a 404 means ci.yml isn't on GitHub yet.")
        return 1
    except urllib.error.URLError as exc:
        print(f"CI: couldn't reach GitHub ({exc.reason}).")
        return 1
    if not runs:
        print("CI: no run found for dev.")
        return 1
    run = runs[0]
    print(f"CI: {run['status']}/{run['conclusion']} on {run['head_sha'][:7]} "
          f"(dev is {head[:7]}) -- {run['html_url']}")
    if run["head_sha"] != head:
        print("CI: the latest run is not on dev's HEAD -- push dev, or wait for its run.")
        return 1
    if run["status"] != "completed":
        print("CI: still running -- wait and run this again.")
        return 1
    if run["conclusion"] == "success":
        return 0
    print_failures(run)
    return 1


def _get(url):
    request = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def print_failures(run):
    """Each failed job's error annotations -- run_all.py writes one per failed
    test file, and unlike the job log they're readable without a login."""
    for job in _get(run["jobs_url"]).get("jobs", []):
        if job["conclusion"] in ("success", "skipped"):
            continue
        print(f"\n== {job['name']}: {job['conclusion']}")
        for note in _get(f"https://api.github.com/repos/{REPO}/check-runs/{job['id']}/annotations"):
            if note.get("annotation_level") == "failure":
                print(f"-- {note.get('title') or ''}\n{note.get('message', '')}")


if __name__ == "__main__":
    sys.exit(main())
