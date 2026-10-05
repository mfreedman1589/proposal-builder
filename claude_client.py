"""The one Claude caller every page shares -- the proposal builder, the
attribution report builder, and the finders-only app (finders_app.py). Moved
out of app.py (2026-10-05) so the finders app can call Claude without
importing the builder. app.py re-imports every name here, so `app.X` keeps
working; a test that needs to patch the caller's internals (its `st`,
`anthropic` or `log_claude_call`) patches THIS module.
"""
import json
import tempfile
from pathlib import Path

import anthropic
import streamlit as st

ANTHROPIC_MODEL = "claude-sonnet-4-6"


# Sized against the LARGEST realistic draft, not the typical one. The old
# ceiling was 2000, set when a draft was a handful of lines and one review
# list -- below what a real proposal now needs, so the model was being cut
# off mid-JSON. Measured worst case: two options of six lines each, three
# audiences carrying avails, full Campaign Specs, and both review lists at
# their 8-item cap comes to ~12,400 characters of pretty-printed JSON, or
# roughly 3,500 output tokens. 16000 leaves ~4.5x headroom on that.
#
# It is also the ceiling for a NON-STREAMING request: past roughly this
# size the SDK starts refusing non-streaming calls it estimates will exceed
# the HTTP timeout. Raising this further means switching these calls to
# client.messages.stream() + get_final_message(), not just editing the
# number. Sonnet 4.6 itself allows up to 128K output.
ANTHROPIC_MAX_TOKENS = 16000


# Where a record of every Claude call goes. The draft path failed live with
# "Claude's response wasn't valid JSON even after stripping markdown fences:
# Expecting value: line 1 column 1 (char 0)" -- char 0 means the text was
# EMPTY, so the fence-stripping the message blamed was never the problem and
# the message pointed at the wrong thing. Nothing was recorded, so a failure
# that cost a live API call and a long wait told us nothing at all. Every
# call now leaves a line behind whether it worked or not.
CLAUDE_LOG_PATH = Path(tempfile.gettempdir()) / "proposal_builder_claude_calls.log"


def _strip_markdown_fences(text):
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines)
    return text.strip()


def _extract_largest_json_object(text):
    """The LARGEST balanced {...} object anywhere in `text`, respecting
    string literals so a brace inside a quoted value doesn't miscount --
    or None if no balanced object is found.

    Not the FIRST one -- that was tried and it silently corrupted a real
    draft. A preamble routinely quotes a small schema fragment verbatim
    while reasoning about a choice ("I'll use group_selection: {"mode":
    "all"}, since the notes describe the whole buy"), and that fragment is
    itself complete, valid JSON. Taking the first balanced object grabbed
    exactly that two-word fragment and returned it as the entire draft --
    no error, no crash, just 41,000 characters of a real LiveWell plan
    silently replaced by `{"mode": "all"}`, discovered only by reading the
    output rather than checking that it parsed. The real answer is
    overwhelmingly the largest object in the response, whether the noise
    around it is a preamble, trailing commentary, or both, so this scans
    every top-level object in the text and returns the longest."""
    candidates = []
    n = len(text)
    i = 0
    while i < n:
        if text[i] != "{":
            i += 1
            continue
        start = i
        depth = 0
        in_string = False
        escape = False
        matched_end = None
        j = i
        while j < n:
            ch = text[j]
            if in_string:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == '"':
                    in_string = False
            else:
                if ch == '"':
                    in_string = True
                elif ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        matched_end = j
                        break
            j += 1
        if matched_end is not None:
            candidates.append(text[start:matched_end + 1])
            i = matched_end + 1
        else:
            # This "{" never closes -- don't skip past it to j (== n),
            # which would abandon the scan; a later "{" might still start
            # a real, well-formed object.
            i = start + 1
    return max(candidates, key=len) if candidates else None


def _parse_draft_json(raw_text):
    try:
        return json.loads(raw_text), None
    except json.JSONDecodeError as exc:
        first_error = exc
    try:
        return json.loads(_strip_markdown_fences(raw_text)), None
    except json.JSONDecodeError:
        pass
    extracted = _extract_largest_json_object(raw_text)
    if extracted is not None:
        try:
            return json.loads(extracted), None
        except json.JSONDecodeError:
            pass
    return None, f"Claude's response wasn't valid JSON even after stripping markdown fences: {first_error}"


_CLAUDE_FAILURE_OUTCOMES = {"truncated", "empty", "unparseable", "refusal", "api_error"}


def log_claude_call(record):
    """Record one call, to the console, to a file, and (on a failure only)
    into session_state. Never raises.

    Console + file, deliberately: the console is what's visible in
    Streamlit Cloud's log viewer, the file is what survives locally after
    the tab is closed. A logging failure must never be what takes a draft
    down, so every error here is swallowed -- the draft is the point, the
    log is the evidence.

    Neither of those is reachable from where a rep actually is, though --
    Streamlit Cloud's log viewer is a developer tool, and the local file
    doesn't exist on the machine a rep is using at all (the LiveWell
    incident: nobody could see what Claude actually returned, and the
    on-screen message pointed at a path that was never going to be there).
    session_state["last_claude_failure"] is the fix -- `stop_reason` plus
    the first 200 characters of what Claude returned, exactly enough to
    diagnose without either an unbounded transcript or anything
    client-facing, which `capture_feedback_state` folds into a rep's
    "Report an issue" so it's readable from the admin page. Overwritten by
    every call (success clears it) rather than accumulated, so it always
    describes the failure that JUST happened, never a stale one from
    earlier in the session.
    """
    line = json.dumps(record, default=str)
    print(f"[claude] {line}")
    try:
        with open(CLAUDE_LOG_PATH, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass
    try:
        outcome = record.get("outcome")
        if outcome in _CLAUDE_FAILURE_OUTCOMES:
            head = (record.get("head") or record.get("error") or "")[:200]
            st.session_state["last_claude_failure"] = {
                "label": record.get("label"),
                "outcome": outcome,
                "stop_reason": record.get("stop_reason"),
                "head": head,
            }
        elif outcome == "ok":
            st.session_state.pop("last_claude_failure", None)
    except Exception:
        pass


def _response_text_and_facts(response):
    """(raw_text, facts) for a Claude response.

    Joins EVERY text block rather than reading content[0].text: a response
    whose first block isn't text -- or which has no blocks at all, which is
    what a pre-output refusal returns -- would otherwise raise IndexError or
    AttributeError inside the try block and be reported as "Claude API call
    failed", hiding what actually happened.
    """
    blocks = list(getattr(response, "content", None) or [])
    raw_text = "".join(getattr(b, "text", "") for b in blocks
                       if getattr(b, "type", None) == "text")
    usage = getattr(response, "usage", None)
    stop_reason = getattr(response, "stop_reason", None)
    facts = {
        "stop_reason": stop_reason,
        "blocks": [getattr(b, "type", "?") for b in blocks],
        "chars": len(raw_text),
        "input_tokens": getattr(usage, "input_tokens", None),
        "output_tokens": getattr(usage, "output_tokens", None),
        "max_tokens": ANTHROPIC_MAX_TOKENS,
        "request_id": getattr(response, "_request_id", None),
        "head": raw_text[:200],
        "tail": raw_text[-200:],
    }
    # stop_details is populated only on a refusal and is None otherwise --
    # read it without guarding and it's an AttributeError on every ordinary
    # response.
    if stop_reason == "refusal":
        details = getattr(response, "stop_details", None)
        facts["refusal_category"] = getattr(details, "category", None)
    return raw_text, facts


def interpret_claude_response(response, label):
    """(parsed, error, retryable) for one raw response.

    Split out from the network call so the failure paths are testable
    without an API key: the truncation and empty-response branches are
    exactly the ones that cost a live run to discover, and they should not
    need another one to re-check.

    stop_reason is consulted BEFORE the JSON is parsed. Truncated JSON fails
    to parse, so a max_tokens cut-off used to be reported as a parse error --
    which points the reader at the response's formatting instead of at its
    length, and is why the original failure was mis-diagnosed as a
    fence-stripping problem.
    """
    # Every message returned below is what a REP sees on screen, so it says
    # what happened in plain terms and what to do next -- retry, then
    # escalate via "Report an issue" -- and nothing dev-facing: no file
    # paths, no session keys, no function/constant names. The one incident
    # that made this the rule: a message once pointed a rep at
    # "%TEMP%/proposal_builder_claude_calls.log", a path that doesn't exist
    # on the machine she was actually using (Streamlit Cloud), and named
    # "ANTHROPIC_MAX_TOKENS in app.py" as if she could edit it. The real
    # diagnostic detail (stop_reason, the response's own head/tail) still
    # goes to `log_claude_call` on every branch below, same as always --
    # it's just not printed into the string a rep reads. See DECISIONS.md,
    # the LiveWell incident, and `capture_feedback_state`'s
    # `last_claude_failure`, which is where that detail actually surfaces.
    raw_text, facts = _response_text_and_facts(response)
    facts["label"] = label

    if facts["stop_reason"] == "max_tokens":
        log_claude_call({**facts, "outcome": "truncated"})
        return None, (
            f"The draft got too long to finish -- it hit Claude's {ANTHROPIC_MAX_TOKENS:,}-token "
            f"response limit and was cut off partway through. Try fewer plan options, fewer "
            f"lines per option, or shorter notes, then draft again. If a plan this size keeps "
            f"happening, use Report an issue and I'll take a look."), True

    if facts["stop_reason"] == "refusal":
        log_claude_call({**facts, "outcome": "refusal"})
        return None, (
            "Claude declined to answer this request. Re-word the notes and try again -- if "
            "they contain nothing unusual, use Report an issue and I'll take a look."), False

    if not raw_text.strip():
        log_claude_call({**facts, "outcome": "empty"})
        return None, (
            "The draft didn't come back in a usable form. Try again -- if it happens twice, "
            "use Report an issue and I'll take a look."), True

    parsed, parse_error = _parse_draft_json(raw_text)
    log_claude_call({**facts, "outcome": "ok" if parsed is not None else "unparseable"})
    if parsed is None:
        return None, (
            "The draft didn't come back in a usable form. Try again -- if it happens twice, "
            "use Report an issue and I'll take a look."), True
    return parsed, None, False


def _call_claude_json(prompt, label="draft", attempts=2, on_attempt=None):
    """Sends one prompt to Claude and parses the response as JSON. Returns
    (parsed_dict, error_message) -- exactly one is None.

    The original bug here (see DECISIONS.md, the LiveWell incident): a
    genuinely open-ended, judgment-heavy notes set made the model narrate
    ("I need to analyze these notes carefully...") instead of returning raw
    JSON, on BOTH attempts, because the retry resent a byte-identical
    prompt and had no reason to behave differently the second time. An
    assistant-turn prefill was tried as the structural fix and abandoned --
    ANTHROPIC_MODEL rejects it with a 400 (see the comment above
    `ANTHROPIC_MAX_TOKENS`) -- so the real defense is
    `_extract_largest_json_object` (`_parse_draft_json`'s fallback, which
    salvages the JSON even when a preamble gets through) plus the corrective
    retry below, not anything at the network-call level.

    Retries once on an empty, truncated or unparseable response before
    surfacing anything -- but the retry is no longer that identical resend.
    It appends a corrective instruction to the prompt ("return ONLY the
    JSON object...") so attempt 2 is actually a different request, not a
    re-roll of the same one. That correction is deliberately generic, not
    the previous attempt's own error text -- `interpret_claude_response`'s
    return value is rep-facing now (plain language, no dev detail), and
    feeding a rep-facing sentence back into the model as if it were
    technical guidance would be both useless to the model and a way for
    on-screen wording to leak into the next request's prompt. An API-level
    failure (auth, network, rate limit) is not retried here -- the SDK
    already retries those itself.

    `on_attempt(attempt, attempts)`, if given, fires before each attempt's
    network call -- the caller's hook for telling a rep what's happening
    instead of a spinner that just sits there (attempt 1 needs no comment;
    attempt 2 means the first pass didn't come back clean, which is worth
    saying).
    """
    api_key = st.secrets.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None, "ANTHROPIC_API_KEY is not set in .streamlit/secrets.toml."

    client = anthropic.Anthropic(api_key=api_key)
    error = "Claude was not called."
    attempt_prompt = prompt
    for attempt in range(1, attempts + 1):
        if on_attempt:
            on_attempt(attempt, attempts)
        try:
            response = client.messages.create(
                model=ANTHROPIC_MODEL,
                max_tokens=ANTHROPIC_MAX_TOKENS,
                messages=[{"role": "user", "content": attempt_prompt}],
            )
        except Exception as exc:                                 # noqa: BLE001
            log_claude_call({"label": label, "attempt": attempt, "outcome": "api_error",
                             "error": f"{type(exc).__name__}: {exc}",
                             "prompt_chars": len(attempt_prompt)})
            return None, ("Claude couldn't be reached. Try again in a moment -- if it keeps "
                          "failing, use Report an issue and I'll take a look.")

        parsed, error, retryable = interpret_claude_response(
            response, f"{label} (attempt {attempt} of {attempts})")
        if parsed is not None:
            return parsed, None
        if not retryable or attempt == attempts:
            break
        attempt_prompt = prompt + (
            "\n\nYour previous response could not be used. This time, return ONLY the JSON "
            "object -- nothing before the opening brace, nothing after the closing one, no "
            "reasoning or commentary.")
    return None, error
