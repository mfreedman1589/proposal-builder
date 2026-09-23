"""DEV_MODE banner and DECK_CHANNEL -- the dev/main split (CLAUDE.md's git
workflow section). Off by default (the public app, unchanged); on only via
an explicit env var or st.secrets flag (the dev deployment), which is what
keeps a public deploy from ever accidentally showing the banner or serving
an unactivated deck/template.

    python tests/test_dev_mode.py
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import db  # noqa: E402
db.fetch_audiences = lambda: (None, "stubbed for test isolation")

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    for key in ("PROPOSAL_BUILDER_DEV_MODE", "PROPOSAL_BUILDER_DECK_CHANNEL"):
        os.environ.pop(key, None)

    import app  # noqa: E402  (after env is clean -- these read it at call time, not import time)

    print("Off by default")
    check("dev_mode_active() is False with nothing set", app.dev_mode_active() is False)
    check("deck_channel() is 'active' with nothing set", db.deck_channel() == "active")

    print("\nOn via the OS environment (a local dev process)")
    os.environ["PROPOSAL_BUILDER_DEV_MODE"] = "1"
    os.environ["PROPOSAL_BUILDER_DECK_CHANNEL"] = "latest"
    check("dev_mode_active() is True", app.dev_mode_active() is True)
    check("deck_channel() is 'latest'", db.deck_channel() == "latest")
    del os.environ["PROPOSAL_BUILDER_DEV_MODE"]
    del os.environ["PROPOSAL_BUILDER_DECK_CHANNEL"]
    check("clearing the env turns both back off",
         app.dev_mode_active() is False and db.deck_channel() == "active")

    print("\nAn unrecognised DECK_CHANNEL value is the safe default, not a typo left running")
    os.environ["PROPOSAL_BUILDER_DECK_CHANNEL"] = "prod"
    check("a typo'd channel value falls back to 'active'", db.deck_channel() == "active")
    del os.environ["PROPOSAL_BUILDER_DECK_CHANNEL"]

    print("\nrender_dev_banner() only calls st.warning when dev mode is on")
    calls = []
    real_st = app.st

    class _StubSt:
        session_state = {}

        @staticmethod
        def warning(msg):
            calls.append(msg)

    app.st = _StubSt()
    try:
        app.render_dev_banner()
        check("no banner call with dev mode off", calls == [], calls)
        os.environ["PROPOSAL_BUILDER_DEV_MODE"] = "1"
        app.render_dev_banner()
        check("exactly one banner call with dev mode on", len(calls) == 1, calls)
        check("the banner names itself DEV", "DEV" in calls[0], calls)
    finally:
        app.st = real_st
        os.environ.pop("PROPOSAL_BUILDER_DEV_MODE", None)

    print(f"\n{len(failures)} failure(s)" if failures else "\nAll checks passed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
