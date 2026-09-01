"""A failed Claude call must say what actually went wrong.

    python tests/test_claude_call.py

Offline, no API key, no network. Every case here is a response-handling
branch, fed a stand-in response object rather than a live call -- which is
the point: the truncation path cost a real drafting run and a long wait to
discover, and re-checking it should not cost another one.

The failure that prompted this: a large scenario came back as "Claude's
response wasn't valid JSON even after stripping markdown fences: Expecting
value: line 1 column 1 (char 0)". Char 0 means the text was EMPTY, so the
fence-stripping the message blamed was never involved -- the message sent
the reader to the wrong place entirely.
"""

import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)
os.environ.setdefault("PROPOSAL_BUILDER_TEST_MODE", "1")

import app  # noqa: E402

# Captured at import time, before any test's captured_log() has a chance to
# swap app.log_claude_call for a stub -- that swap is never undone until
# main()'s own final `finally`, so any test needing the REAL function (not
# a test's own capturing stub) mid-run has to reach for this, not
# app.log_claude_call directly.
REAL_LOG_CLAUDE_CALL = app.log_claude_call

PASSED = FAILED = 0


def check(label, condition, detail=""):
    global PASSED, FAILED
    if condition:
        PASSED += 1
        print(f"    PASS  {label}")
    else:
        FAILED += 1
        print(f"    FAIL  {label}{('  ' + str(detail)) if detail else ''}")


class Block:
    def __init__(self, text=None, type="text"):
        self.type = type
        if text is not None:
            self.text = text


class Usage:
    input_tokens = 12345
    output_tokens = 16000


class Response:
    """Shaped like an anthropic Message, only as far as this code reads it."""

    def __init__(self, content, stop_reason="end_turn", stop_details=None):
        self.content = content
        self.stop_reason = stop_reason
        self.stop_details = stop_details
        self.usage = Usage()
        self._request_id = "req_test_0001"


def captured_log():
    """Collect log records instead of writing them, and return the list."""
    records = []
    app.log_claude_call = records.append
    return records


def restore_log(real):
    app.log_claude_call = real


def test_truncation_is_not_reported_as_a_parse_error():
    print("\n  A truncated response says so, and says what to do")
    records = captured_log()
    # Real shape of a cut-off draft: valid JSON prefix, no closing braces.
    cut_off = '{\n  "client_name": "Ridgeline Dermatology",\n  "options": [\n    {"name": "Opt'
    parsed, error, retryable = app.interpret_claude_response(
        Response([Block(cut_off)], stop_reason="max_tokens"), "draft")
    check("nothing parsed", parsed is None, parsed)
    check("retryable", retryable is True)
    check("says the draft was too long", "too long" in error, error)
    check("names the token limit", f"{app.ANTHROPIC_MAX_TOKENS:,}" in error, error)
    check("suggests something actionable",
          "fewer plan options" in error and "shorter notes" in error, error)
    check("does NOT blame JSON parsing",
          "valid JSON" not in error and "markdown fences" not in error, error)
    check("logged as truncated",
          [r for r in records if r.get("outcome") == "truncated"], records)
    check("log carries stop_reason",
          records and records[-1]["stop_reason"] == "max_tokens", records[-1])
    check("log carries token counts",
          records[-1]["output_tokens"] == 16000 and records[-1]["input_tokens"] == 12345,
          records[-1])
    check("log carries the first 200 characters",
          records[-1]["head"].startswith('{\n  "client_name"'), records[-1]["head"])
    check("log carries the last 200 characters",
          records[-1]["tail"].endswith('{"name": "Opt'), records[-1]["tail"])
    check("log carries the request id",
          records[-1]["request_id"] == "req_test_0001", records[-1])


def test_empty_response():
    print("\n  An empty response tells a rep to retry, not a JSON/log-path detail")
    records = captured_log()
    parsed, error, retryable = app.interpret_claude_response(
        Response([Block("")], stop_reason="end_turn"), "draft")
    check("nothing parsed", parsed is None)
    check("retryable", retryable is True)
    check("plain, rep-facing wording", "didn't come back in a usable form" in error, error)
    check("tells the rep to retry", "try again" in error.lower(), error)
    check("escalates via Report an issue", "Report an issue" in error, error)
    # The incident that made this a rule: this message used to point a rep
    # at a local temp-file path that doesn't exist on Streamlit Cloud, and
    # named the raw stop_reason -- neither is something a rep can act on.
    # That detail still goes to the log (checked below); it just isn't in
    # the string she reads on screen.
    check("no dev-facing log path on screen", str(app.CLAUDE_LOG_PATH) not in error, error)
    check("no raw stop_reason jargon on screen", "end_turn" not in error, error)
    check("logged as empty", records[-1]["outcome"] == "empty", records[-1])
    check("the log itself still carries the real stop_reason",
          records[-1]["stop_reason"] == "end_turn", records[-1])


def test_no_content_blocks_at_all():
    print("\n  A response with no blocks doesn't crash the handler")
    # This is what a pre-output refusal returns. Reading content[0].text
    # here would be an IndexError reported as "Claude API call failed".
    captured_log()
    parsed, error, _ = app.interpret_claude_response(
        Response([], stop_reason="end_turn"), "draft")
    check("handled, not raised", parsed is None and "didn't come back in a usable form" in error, error)


def test_non_text_first_block():
    print("\n  A non-text block first doesn't hide the text after it")
    captured_log()
    payload = '{"client_name": "Acme"}'
    parsed, error, _ = app.interpret_claude_response(
        Response([Block(type="thinking"), Block(payload)]), "draft")
    check("parsed from the text block", parsed == {"client_name": "Acme"}, (parsed, error))


def test_refusal():
    print("\n  A refusal is named as a refusal and is not retried")
    records = captured_log()

    class Details:
        category = "cyber"

    parsed, error, retryable = app.interpret_claude_response(
        Response([], stop_reason="refusal", stop_details=Details()), "draft")
    check("nothing parsed", parsed is None)
    check("not retryable", retryable is False)
    check("says Claude declined", "declined" in error, error)
    check("category recorded", records[-1].get("refusal_category") == "cyber", records[-1])


def test_success_and_fences():
    print("\n  A good response parses, fenced or not")
    records = captured_log()
    parsed, error, _ = app.interpret_claude_response(
        Response([Block('{"client_name": "Acme"}')]), "draft")
    check("plain JSON parses", parsed == {"client_name": "Acme"}, (parsed, error))
    parsed, error, _ = app.interpret_claude_response(
        Response([Block('```json\n{"client_name": "Acme"}\n```')]), "draft")
    check("fenced JSON parses", parsed == {"client_name": "Acme"}, (parsed, error))
    check("logged as ok", records[-1]["outcome"] == "ok", records[-1])


def test_retry_then_succeed():
    print("\n  A retryable failure is retried once before anything is surfaced")
    captured_log()
    calls = {"n": 0}
    good = '{"client_name": "Acme"}'

    class Client:
        class messages:
            @staticmethod
            def create(**kwargs):
                calls["n"] += 1
                if calls["n"] == 1:
                    return Response([Block("")], stop_reason="end_turn")
                return Response([Block(good)])

    real_anthropic, real_st = app.anthropic, app.st

    class FakeAnthropic:
        Anthropic = staticmethod(lambda api_key=None: Client())

    class FakeSt:
        secrets = {"ANTHROPIC_API_KEY": "sk-test"}

    app.anthropic, app.st = FakeAnthropic, FakeSt
    try:
        parsed, error = app._call_claude_json("prompt", label="draft")
    finally:
        app.anthropic, app.st = real_anthropic, real_st
    check("two attempts made", calls["n"] == 2, calls["n"])
    check("second attempt's result returned", parsed == {"client_name": "Acme"}, (parsed, error))
    check("no error surfaced", error is None, error)


def test_retry_gives_up_with_the_real_message():
    print("\n  Two failures surface the specific message, not a generic one")
    captured_log()
    calls = {"n": 0}

    class Client:
        class messages:
            @staticmethod
            def create(**kwargs):
                calls["n"] += 1
                return Response([Block("{partial")], stop_reason="max_tokens")

    real_anthropic, real_st = app.anthropic, app.st

    class FakeAnthropic:
        Anthropic = staticmethod(lambda api_key=None: Client())

    class FakeSt:
        secrets = {"ANTHROPIC_API_KEY": "sk-test"}

    app.anthropic, app.st = FakeAnthropic, FakeSt
    try:
        parsed, error = app._call_claude_json("prompt", label="draft")
    finally:
        app.anthropic, app.st = real_anthropic, real_st
    check("retried once, then stopped", calls["n"] == 2, calls["n"])
    check("surfaces the truncation message", "too long" in error, error)


def test_refusal_is_not_retried():
    print("\n  A refusal is not retried -- it would just refuse again")
    captured_log()
    calls = {"n": 0}

    class Details:
        category = "cyber"

    class Client:
        class messages:
            @staticmethod
            def create(**kwargs):
                calls["n"] += 1
                return Response([], stop_reason="refusal", stop_details=Details())

    real_anthropic, real_st = app.anthropic, app.st

    class FakeAnthropic:
        Anthropic = staticmethod(lambda api_key=None: Client())

    class FakeSt:
        secrets = {"ANTHROPIC_API_KEY": "sk-test"}

    app.anthropic, app.st = FakeAnthropic, FakeSt
    try:
        parsed, error = app._call_claude_json("prompt", label="draft")
    finally:
        app.anthropic, app.st = real_anthropic, real_st
    check("called exactly once", calls["n"] == 1, calls["n"])
    check("surfaces the refusal message", "declined" in error, error)


def test_extract_first_json_object():
    print("\n  A JSON object embedded in a preamble or followed by commentary is salvaged")
    preamble = 'I need to think about this carefully.\n\n{"client_name": "Acme", "n": {"a": 1}}'
    check("extracts the object past a preamble",
          app._extract_first_json_object(preamble) == '{"client_name": "Acme", "n": {"a": 1}}',
          app._extract_first_json_object(preamble))

    trailing = '{"client_name": "Acme"} Let me know if you would like changes.'
    check("extracts the object before trailing commentary",
          app._extract_first_json_object(trailing) == '{"client_name": "Acme"}',
          app._extract_first_json_object(trailing))

    quoted_brace = '{"note": "use { and }"}'
    check("a brace inside a quoted string doesn't miscount",
          app._extract_first_json_object(quoted_brace) == quoted_brace,
          app._extract_first_json_object(quoted_brace))

    check("no object at all returns None", app._extract_first_json_object("no json here") is None)


def test_preamble_salvaged_by_extraction_fallback():
    print("\n  A preamble before the JSON is salvaged, not reported as a parse failure")
    records = captured_log()
    # The actual shape of the LiveWell incident: the model narrates instead
    # of continuing straight into JSON, despite the prompt's own "no
    # preamble" instruction. An assistant-turn prefill was tried as a
    # structural fix and abandoned (ANTHROPIC_MODEL rejects it -- see the
    # comment above ANTHROPIC_MAX_TOKENS), so this extraction fallback is
    # the ACTUAL defense, not a backstop for one.
    preamble_response = ('I need to analyze these notes carefully before building the JSON.\n\n'
                         '{"client_name": "Acme"}')
    parsed, error, _ = app.interpret_claude_response(
        Response([Block(preamble_response)]), "draft")
    check("salvaged despite the preamble", parsed == {"client_name": "Acme"}, (parsed, error))
    check("logged as ok, not unparseable", records[-1]["outcome"] == "ok", records[-1])


def test_retry_prompt_is_corrected_not_identical():
    print("\n  The retry appends a corrective instruction, not an identical resend")
    captured_log()
    messages_seen = []

    class Client:
        class messages:
            @staticmethod
            def create(**kwargs):
                messages_seen.append(kwargs["messages"])
                if len(messages_seen) == 1:
                    return Response([Block("")], stop_reason="end_turn")
                return Response([Block('{"client_name": "Acme"}')])

    real_anthropic, real_st = app.anthropic, app.st

    class FakeAnthropic:
        Anthropic = staticmethod(lambda api_key=None: Client())

    class FakeSt:
        secrets = {"ANTHROPIC_API_KEY": "sk-test"}

    app.anthropic, app.st = FakeAnthropic, FakeSt
    try:
        parsed, error = app._call_claude_json("ORIGINAL PROMPT", label="draft")
    finally:
        app.anthropic, app.st = real_anthropic, real_st
    prompts_seen = [m[0]["content"] for m in messages_seen]
    check("two attempts made", len(prompts_seen) == 2, len(prompts_seen))
    check("second prompt differs from the first", prompts_seen[1] != prompts_seen[0])
    check("second prompt still contains the original notes/prompt",
          "ORIGINAL PROMPT" in prompts_seen[1], prompts_seen[1])
    # Deliberately generic, not the previous attempt's own rep-facing error
    # text -- see _call_claude_json's docstring for why those two are kept
    # separate now.
    check("second prompt asks for JSON only",
          "return ONLY the JSON" in prompts_seen[1], prompts_seen[1])
    check("every attempt is a single plain user message (no prefill -- ANTHROPIC_MODEL rejects it)",
          all(m == [{"role": "user", "content": prompts_seen[i]}]
              for i, m in enumerate(messages_seen)),
          messages_seen)
    check("parsed on the corrected retry", parsed == {"client_name": "Acme"}, (parsed, error))


def test_on_attempt_callback_fires_per_attempt():
    print("\n  on_attempt fires before each network call, with the right attempt numbers")
    captured_log()
    seen = []

    class Client:
        class messages:
            @staticmethod
            def create(**kwargs):
                if len(seen) == 1:
                    return Response([Block("")], stop_reason="end_turn")
                return Response([Block('{"client_name": "Acme"}')])

    real_anthropic, real_st = app.anthropic, app.st

    class FakeAnthropic:
        Anthropic = staticmethod(lambda api_key=None: Client())

    class FakeSt:
        secrets = {"ANTHROPIC_API_KEY": "sk-test"}

    app.anthropic, app.st = FakeAnthropic, FakeSt
    try:
        app._call_claude_json("prompt", label="draft", on_attempt=lambda a, n: seen.append((a, n)))
    finally:
        app.anthropic, app.st = real_anthropic, real_st
    check("called once per attempt, with correct (attempt, attempts)", seen == [(1, 2), (2, 2)], seen)


def test_last_claude_failure_captured_for_feedback():
    print("\n  A failure is captured into session_state for the feedback export -- and cleared on success")
    # Uses the REAL log_claude_call (not captured_log()'s stub, and not
    # whatever an EARLIER test's captured_log() left installed -- that
    # stub is never undone until main()'s own final `finally`) -- this
    # test is specifically about the session_state side effect the real
    # function has.
    real_st, real_log = app.st, app.log_claude_call
    app.log_claude_call = REAL_LOG_CLAUDE_CALL

    class FakeSt:
        secrets = {}
        session_state = {}

    app.st = FakeSt
    try:
        app.interpret_claude_response(
            Response([Block("{partial")], stop_reason="max_tokens"), "draft")
        failure = FakeSt.session_state.get("last_claude_failure")
        check("failure captured", failure is not None, failure)
        check("carries outcome", failure and failure["outcome"] == "truncated", failure)
        check("carries stop_reason", failure and failure["stop_reason"] == "max_tokens", failure)
        check("carries a head snippet",
              failure and failure["head"] == "{partial", failure)
        check("carries which call this was", failure and failure["label"] == "draft", failure)

        app.interpret_claude_response(Response([Block('{"client_name": "Acme"}')]), "draft")
        check("cleared by the NEXT call's success -- doesn't linger from an earlier failure",
              "last_claude_failure" not in FakeSt.session_state, dict(FakeSt.session_state))
    finally:
        app.st, app.log_claude_call = real_st, real_log


def test_ceiling_covers_the_largest_realistic_draft():
    print("\n  The token ceiling clears the largest realistic draft")
    # Asserted against a draft built here, not against a number copied from
    # the constant -- a check whose expectation comes from the thing it is
    # checking proves only that the code agrees with itself.
    import json
    note = ("The notes did not say whether the production fee sits inside the stated budget "
            "or on top of it. This draft puts it on top; confirm before sending.")
    line = {"product": "premion_streaming_tv", "label": "Premion Streaming TV - Track",
            "audience_track": "Adults 35+, homeowners, higher household income (100K+), "
                              "in-market for cosmetic procedures",
            "allocation": {"percent_of_avails": 40, "avails_ref": "RETAIL Home Services"},
            "cpm": 30}
    option = {"name": "Option 1 - 40% Reach", "total_budget": 150000,
              "breakout": "monthly", "media_plan_lines": [line] * 6}
    draft = {
        "client_name": "Ridgeline Dermatology Associates of Northern Virginia",
        "vertical": "healthcare", "market": "DC", "geo": "Washington, DC DMA",
        "flight_start": "2026-09-01", "flight_end": "2026-11-30", "breakout": "monthly",
        "total_budget": 150000,
        "campaign_specs": {k: ["A reasonably long campaign specification bullet of the kind "
                               "the model writes from detailed discovery notes."] * 4
                           for k in ("goals", "audience", "geography", "budget",
                                     "placements", "timing")},
        "audiences": [{"segment": "RETAIL Home Services Home Improvement",
                       "geo": "Washington, DC DMA", "max_avails": 1000000,
                       "avails_basis": "monthly"}] * 3,
        "options": [option] * 2, "sports": ["nfl_reg"],
        "attribution": {"sales_attribution": True, "brand_lift": True},
        "unresolved": [note] * 8, "unresolved_internal": [note] * 8,
    }
    chars = len(json.dumps(draft, indent=2))
    # 3.5 characters per token is a deliberately pessimistic rate for JSON
    # (punctuation-dense text tokenizes closer to 4).
    tokens = chars / 3.5
    print(f"    ....  largest realistic draft: {chars:,} chars, ~{tokens:,.0f} tokens")
    check(f"ceiling ({app.ANTHROPIC_MAX_TOKENS:,}) clears it with 2x headroom",
          app.ANTHROPIC_MAX_TOKENS >= tokens * 2, app.ANTHROPIC_MAX_TOKENS)
    # Above roughly this size a non-streaming request risks an HTTP timeout;
    # raising the ceiling further means switching to streaming, so the two
    # decisions are pinned together here.
    check("ceiling stays inside the non-streaming safe range",
          app.ANTHROPIC_MAX_TOKENS <= 16000, app.ANTHROPIC_MAX_TOKENS)


def main():
    print("=" * 70)
    print("Claude call failure handling")
    print("=" * 70)
    real_log = app.log_claude_call
    try:
        test_truncation_is_not_reported_as_a_parse_error()
        test_empty_response()
        test_no_content_blocks_at_all()
        test_non_text_first_block()
        test_refusal()
        test_success_and_fences()
        test_retry_then_succeed()
        test_retry_gives_up_with_the_real_message()
        test_refusal_is_not_retried()
        test_extract_first_json_object()
        test_preamble_salvaged_by_extraction_fallback()
        test_retry_prompt_is_corrected_not_identical()
        test_on_attempt_callback_fires_per_attempt()
        test_last_claude_failure_captured_for_feedback()
        test_ceiling_covers_the_largest_realistic_draft()
    finally:
        restore_log(real_log)
    print("\n" + "=" * 70)
    print(f"{PASSED} passed, {FAILED} failed")
    print("=" * 70)
    return 1 if FAILED else 0


if __name__ == "__main__":
    sys.exit(main())
