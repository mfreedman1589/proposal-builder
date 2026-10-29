"""Loaded into every Python process the CI workflow starts (PYTHONPATH=tests/ci).

The repo is public, and dev and prod share one Supabase database, so a CI run
must never reach a backend, however a future change might try to. Two locks:
the Supabase and Anthropic environment variables are removed before anything
reads them, and every hostname lookup other than loopback fails. A test that
needs the network therefore fails in CI, and belongs in run_all.LOCAL_ONLY.
"""
import os
import socket
import sys

for _name in ("SUPABASE_URL", "SUPABASE_SERVICE_KEY", "ANTHROPIC_API_KEY"):
    os.environ.pop(_name, None)

_real_getaddrinfo = socket.getaddrinfo


def _loopback_only(host, *args, **kwargs):
    name = host.decode() if isinstance(host, bytes) else host
    if name and name != "localhost" and not name.startswith("127.") and name != "::1":
        sys.stderr.write(f"CI: network blocked ({name}) -- see tests/ci/sitecustomize.py\n")
        raise socket.gaierror(f"network is blocked in CI ({name})")
    return _real_getaddrinfo(host, *args, **kwargs)


socket.getaddrinfo = _loopback_only
