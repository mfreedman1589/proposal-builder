"""call_claude_categorize's validation logic -- offline, no API key needed.
Stubs app._call_claude_json so the network is never touched; the live
comparison against the real workbook is tests/test_categorize_live.py
(Tier 2, costs money, user-run).

    python tests/test_categorize_prompt.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

failures = []


def check(label, condition, detail=""):
    if condition:
        print(f"  PASS  {label}")
    else:
        print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
        failures.append(label)


def main():
    components = [
        {"name": "CUSTOM Blue Bell Ice Cream Shoppers", "impressions": 100},
        {"name": "CUSTOM City of Mesa", "impressions": 200},
        {"name": "Something Ambiguous", "impressions": 50},
    ]

    print("build_categorize_prompt")
    prompt = app.build_categorize_prompt(components)
    check("every real category name appears in the prompt",
          all(cat in prompt for cat in app.CATEGORY_DESCRIPTIONS), None)
    check("every component name and its impressions appear",
          all(c["name"] in prompt and f"{c['impressions']:,}" in prompt for c in components), None)

    print("\ncall_claude_categorize: valid response passes through")
    real_call = app._call_claude_json

    def stub_valid(prompt, label="draft", attempts=2):
        return {"components": [
            {"index": 0, "category": "FOOD", "category2": None,
             "is_client_specific": False, "confidence": "high", "reason": "Food shopper audience."},
            {"index": 1, "category": None, "category2": None,
             "is_client_specific": True, "confidence": "high", "reason": "A specific city's own campaign."},
            {"index": 2, "category": "LIFESTYLE", "category2": "TRAVEL",
             "is_client_specific": False, "confidence": "low", "reason": "Could go either way."},
        ]}, None

    app._call_claude_json = stub_valid
    try:
        suggestions, error = app.call_claude_categorize(components)
    finally:
        app._call_claude_json = real_call
    check("no error", error is None, error)
    by_name = {s["segment"]: s for s in suggestions}
    check("FOOD category passed through", by_name["CUSTOM Blue Bell Ice Cream Shoppers"]["category"] == "FOOD",
          by_name.get("CUSTOM Blue Bell Ice Cream Shoppers"))
    check("is_client_specific passed through, category stays blank",
          by_name["CUSTOM City of Mesa"]["is_client_specific"] is True
          and by_name["CUSTOM City of Mesa"]["category"] == "", by_name.get("CUSTOM City of Mesa"))
    check("a genuine second category passes through",
          by_name["Something Ambiguous"]["category"] == "LIFESTYLE"
          and by_name["Something Ambiguous"]["category2"] == "TRAVEL", by_name.get("Something Ambiguous"))

    print("\nAssertion: an invented category is never trusted")

    def stub_invented(prompt, label="draft", attempts=2):
        return {"components": [
            {"index": 0, "category": "GROCERY", "category2": None,  # not a real category
             "is_client_specific": False, "confidence": "high", "reason": "Made up."},
            {"index": 1, "category": "AUTO", "category2": "NOT_REAL_EITHER",
             "is_client_specific": False, "confidence": "high", "reason": "Second one invented."},
        ]}, None

    app._call_claude_json = stub_invented
    try:
        suggestions, error = app.call_claude_categorize(components[:2])
    finally:
        app._call_claude_json = real_call
    by_name = {s["segment"]: s for s in suggestions}
    check("an invented category is dropped to blank, not trusted",
          by_name["CUSTOM Blue Bell Ice Cream Shoppers"]["category"] == "", by_name)
    check("a valid category1 survives even when category2 is invented",
          by_name["CUSTOM City of Mesa"]["category"] == "AUTO"
          and by_name["CUSTOM City of Mesa"]["category2"] == "", by_name)

    print("\nAssertion: an index Claude invents or skips is ignored, not guessed at")

    def stub_bad_index(prompt, label="draft", attempts=2):
        return {"components": [
            {"index": 99, "category": "FOOD", "is_client_specific": False,
             "confidence": "high", "reason": "Wrong index."},
        ]}, None

    app._call_claude_json = stub_bad_index
    try:
        suggestions, error = app.call_claude_categorize(components)
    finally:
        app._call_claude_json = real_call
    check("no suggestion is fabricated for an out-of-range index",
          suggestions == [], suggestions)

    print("\nempty input never calls the network")
    calls = []
    app._call_claude_json = lambda *a, **k: calls.append(1) or ({}, None)
    try:
        suggestions, error = app.call_claude_categorize([])
    finally:
        app._call_claude_json = real_call
    check("no components in, no suggestions out, no call made",
          suggestions == [] and error is None and not calls, (suggestions, error, calls))

    print()
    if failures:
        print(f"{len(failures)} FAILED: {failures}")
        return 1
    print("The prompt names every real category and every component; the response parser "
          "trusts nothing it can't validate -- an invented category or an out-of-range index "
          "is dropped rather than passed through.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
