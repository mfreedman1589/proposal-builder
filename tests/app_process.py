"""Own the local app process, so "restart and try again" is something the
assistant does, not something it has to ask a human to do.

    python tests/app_process.py start|stop|restart|status [--port N]

Launches `streamlit run app.py` in test mode
(`PROPOSAL_BUILDER_TEST_MODE=1`, see app.py's own guard -- read only from the
OS environment, refused outright when the process looks like Streamlit
Cloud) on a fixed, non-default port so it can never collide with a real
instance the user has open at Streamlit's own default (8501). PID and logs
live under the OS temp dir, the same place deck_render.py keeps rendered
output -- outside the synced project directory, and gitignored either way.

This exists because of an incident (twice, now -- see CLAUDE.md's "When a
feature is verified working in tests but reported broken in the running
app" note): a feature that was correct in every offline test and in a fresh
AppTest process was reported broken in the live app, and the actual cause
both times was a stale warm process that predated the fix. Restarting to
rule that out first used to mean asking the user to do it and wait for their
answer; this makes it something the reproduce-and-verify loop (run_scenario.py)
does on every run, for free.

Deliberately NOT for the deployed app -- there is no "deployed app" concept
here at all, only ever a local subprocess this script starts and later kills
by the exact PID it recorded. Nothing here touches Streamlit Cloud, sets any
secret, or reads .streamlit/secrets.toml (Streamlit reads that itself; this
script never opens it).
"""
import argparse
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

DEFAULT_PORT = 8517  # distinct from Streamlit's own default (8501) on purpose

TEST_MODE_ENV = "PROPOSAL_BUILDER_TEST_MODE"


def _state_dir():
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or "/tmp"
    d = Path(base) / "proposal-builder-app"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _pid_file(port):
    return _state_dir() / f"app_{port}.pid"


def _log_file(port):
    return _state_dir() / f"app_{port}.log"


def _read_pid(port):
    pf = _pid_file(port)
    if not pf.exists():
        return None
    try:
        return int(pf.read_text().strip())
    except (ValueError, OSError):
        return None


def _alive(pid):
    import psutil
    if pid is None:
        return False
    try:
        proc = psutil.Process(pid)
        # A recycled PID belongs to a different process -- the same identity
        # trap this project has already been bitten by once with id()-keyed
        # caches (ImportCache, _THEME_CACHE). Confirm it's actually still
        # ours, not just that the number is in use by something.
        return proc.is_running() and "streamlit" in " ".join(proc.cmdline()).lower()
    except psutil.NoSuchProcess:
        return False


def _health_ok(port, timeout=2):
    for path in ("/_stcore/health", "/healthz"):
        try:
            with urllib.request.urlopen(f"http://localhost:{port}{path}", timeout=timeout) as resp:
                if resp.status == 200:
                    return True
        except (urllib.error.URLError, ConnectionError, TimeoutError):
            continue
    return False


def status(port=DEFAULT_PORT):
    pid = _read_pid(port)
    running = _alive(pid)
    healthy = _health_ok(port) if running else False
    return {"port": port, "pid": pid, "process_alive": running, "http_healthy": healthy}


def start(port=DEFAULT_PORT, wait_seconds=45):
    """Launch app.py in test mode on `port` if it isn't already running
    there; wait until it answers its health check. Returns (ok, message).
    """
    import psutil
    st = status(port)
    if st["process_alive"] and st["http_healthy"]:
        return True, f"already running (pid {st['pid']}, port {port})"
    if st["process_alive"] and not st["http_healthy"]:
        # A process is there but never came up cleanly -- don't pile a
        # second one on top of it.
        stop(port)

    env = dict(os.environ)
    env[TEST_MODE_ENV] = "1"
    log_path = _log_file(port)
    log = open(log_path, "w", encoding="utf-8")
    creationflags = 0
    if sys.platform == "win32":
        creationflags = subprocess_no_window()
    proc = _popen(
        [sys.executable, "-m", "streamlit", "run", "app.py",
         "--server.headless=true", f"--server.port={port}",
         "--server.address=localhost", "--browser.gatherUsageStats=false"],
        cwd=str(REPO), env=env, stdout=log, stderr=log, creationflags=creationflags)
    _pid_file(port).write_text(str(proc.pid))

    deadline = time.time() + wait_seconds
    while time.time() < deadline:
        if _health_ok(port):
            return True, f"started (pid {proc.pid}, port {port}), healthy after " \
                         f"{wait_seconds - (deadline - time.time()):.1f}s"
        if proc.poll() is not None:
            return False, (f"process exited during startup (code {proc.returncode}) -- "
                           f"see {log_path}")
        time.sleep(0.5)
    return False, f"did not become healthy within {wait_seconds}s -- see {log_path}"


def _popen(args, **kwargs):
    import subprocess
    return subprocess.Popen(args, **kwargs)


def subprocess_no_window():
    import subprocess
    return subprocess.CREATE_NO_WINDOW


def stop(port=DEFAULT_PORT, timeout=10):
    """Terminate the process THIS tool started, by the PID it recorded --
    never a blind taskkill by name, which could hit an unrelated Streamlit
    instance. Returns (ok, message)."""
    import psutil
    pid = _read_pid(port)
    if pid is None:
        return True, "not running (no pid file)"
    if not _alive(pid):
        _pid_file(port).unlink(missing_ok=True)
        return True, "not running (stale pid file removed)"
    try:
        proc = psutil.Process(pid)
        children = proc.children(recursive=True)
        for child in children:
            child.terminate()
        proc.terminate()
        gone, alive = psutil.wait_procs([proc] + children, timeout=timeout)
        for p in alive:
            p.kill()
    except psutil.NoSuchProcess:
        pass
    _pid_file(port).unlink(missing_ok=True)
    return True, f"stopped (was pid {pid})"


def restart(port=DEFAULT_PORT, wait_seconds=45):
    stop(port)
    return start(port, wait_seconds=wait_seconds)


def ensure_running(port=DEFAULT_PORT, wait_seconds=45):
    """Start it only if it isn't already healthy -- the common case for a
    tool that runs this before every scenario."""
    st = status(port)
    if st["process_alive"] and st["http_healthy"]:
        return True, f"already running (pid {st['pid']}, port {port})"
    return start(port, wait_seconds=wait_seconds)


def base_url(port=DEFAULT_PORT):
    return f"http://localhost:{port}"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["start", "stop", "restart", "status", "ensure"])
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--wait", type=float, default=45)
    args = parser.parse_args()

    if args.action == "status":
        print(status(args.port))
        return 0
    if args.action == "start":
        ok, msg = start(args.port, wait_seconds=args.wait)
    elif args.action == "stop":
        ok, msg = stop(args.port)
    elif args.action == "ensure":
        ok, msg = ensure_running(args.port, wait_seconds=args.wait)
    else:
        ok, msg = restart(args.port, wait_seconds=args.wait)
    print(("OK  " if ok else "FAIL") + " " + msg)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
