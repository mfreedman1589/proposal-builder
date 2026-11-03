"""The second half of tests/test_attribution_reports_page.py, split
2026-11-03 so each file finishes well inside a 10-minute foreground call:
Clear, the pre-linked door, the proposal link, Report history, the follow-up
proposal, goals prefill, the proposal picker, the pixel-issue window, the
vertical/conversion fields, optimization, the one-slide summary and case
study, the vault browser label, and dev warnings reaching the feedback
export. Same fixtures, fake store and helpers, imported from the first half.

    python tests/test_attribution_reports_page_2.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import test_attribution_reports_page as suite  # noqa: E402


if __name__ == "__main__":
    sys.exit(suite.run_checks(suite.SECOND_HALF, with_feedback_export=True))
