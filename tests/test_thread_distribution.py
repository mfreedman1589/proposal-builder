"""Offline, free -- report_assembly.distribute_threads, the Highlights/
Takeaways rework's own Python-side fan-out from one `threads` array into
(highlight_bullets, takeaway_bullets, whats_next_bullets). Hand-built thread
payloads only -- no model call, no export, per the rework's own instruction
that this logic is tested offline against a hand-built payload.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import report_assembly as ra  # noqa: E402


class Report:
    def __init__(self):
        self.passed, self.failed = 0, []

    def check(self, label, ok, detail=""):
        if ok:
            print(f"  PASS  {label}")
            self.passed += 1
        else:
            print(f"  FAIL  {label}{'  -- ' + str(detail) if detail else ''}")
            self.failed.append(label)


def _thread(head, anchor="goal", goal_ref=None, finding=None, meaning=None, action=None):
    return {"head": head, "anchor": anchor, "goal_ref": goal_ref,
           "finding": finding, "meaning": meaning, "action": action}


def check_basic_shape(rep):
    print("\nBasic shape -- one goal thread, full fields")
    threads = [_thread("Frequency on target", finding="3.95 average frequency, inside the 3-5 goal.",
                       meaning="The campaign held the client's target frequency band all flight.",
                       action="Keep frequency in the 3-5 band for the renewal.")]
    hi, ta, wn = ra.distribute_threads(threads)
    rep.check("one highlight", hi == [("Frequency on target",
             "3.95 average frequency, inside the 3-5 goal.")], hi)
    rep.check("one takeaway", ta == [("Frequency on target",
             "The campaign held the client's target frequency band all flight.")], ta)
    rep.check("one what's-next item", wn == ["Keep frequency in the 3-5 band for the renewal."], wn)


def check_closing_only_signal(rep):
    print("\nA signal thread with no finding contributes a takeaway only")
    threads = [
        _thread("Direct visits lead", anchor="signal", finding=None,
               meaning="31% of visits arrived direct -- strong ad recall.", action=None),
    ]
    hi, ta, wn = ra.distribute_threads(threads)
    rep.check("no highlight (finding is null)", hi == [], hi)
    rep.check("takeaway still appears", ta == [("Direct visits lead",
             "31% of visits arrived direct -- strong ad recall.")], ta)
    rep.check("no what's-next item", wn == [], wn)


def check_no_action_no_whats_next_item(rep):
    print("\nA thread with no action contributes nothing to what's-next")
    threads = [_thread("Store visits", finding="f", meaning="m", action=None)]
    _hi, _ta, wn = ra.distribute_threads(threads)
    rep.check("what's-next stays empty", wn == [], wn)


def check_goal_order_is_priority(rep):
    print("\nGoal threads keep the model's own order; signals always come after")
    threads = [
        _thread("Signal A", anchor="signal", finding="sf", meaning="sm"),
        _thread("Goal 1", anchor="goal", finding="g1f", meaning="g1m"),
        _thread("Goal 2", anchor="goal", finding="g2f", meaning="g2m"),
    ]
    hi, _ta, _wn = ra.distribute_threads(threads)
    rep.check("goal threads first, in listed order, signal last",
             [h for h, _d in hi] == ["Goal 1", "Goal 2", "Signal A"], hi)


def check_cap_drops_signals_first(rep):
    print("\nOver the cap of 4 -- signals drop before any goal thread")
    goals = [_thread(f"Goal {i}", finding=f"gf{i}", meaning=f"gm{i}") for i in range(4)]
    signals = [_thread(f"Signal {i}", anchor="signal", finding=f"sf{i}", meaning=f"sm{i}")
              for i in range(2)]
    hi, ta, _wn = ra.distribute_threads(goals + signals)
    rep.check("all 4 goal highlights survive, both signals dropped",
             [h for h, _d in hi] == [f"Goal {i}" for i in range(4)], hi)
    rep.check("same for takeaways", [h for h, _d in ta] == [f"Goal {i}" for i in range(4)], ta)


def check_cap_drops_last_signal_first(rep):
    print("\nOver the cap with 3 goals + 2 signals -- only the LAST signal drops")
    goals = [_thread(f"Goal {i}", finding=f"gf{i}", meaning=f"gm{i}") for i in range(3)]
    signals = [_thread(f"Signal {i}", anchor="signal", finding=f"sf{i}", meaning=f"sm{i}")
              for i in range(2)]
    hi, _ta, _wn = ra.distribute_threads(goals + signals)
    rep.check("3 goals + Signal 0 survive (5 total trimmed to 4, Signal 1 dropped)",
             [h for h, _d in hi] == ["Goal 0", "Goal 1", "Goal 2", "Signal 0"], hi)


def check_whats_next_has_no_cap_of_its_own(rep):
    print("\nWhat's-next has no cap TIGHTER than the takeaway cap it inherits")
    threads = [_thread(f"Goal {i}", finding=f"f{i}", meaning=f"m{i}", action=f"a{i}")
              for i in range(4)]
    _hi, _ta, wn = ra.distribute_threads(threads)
    rep.check("all 4 actions present -- nothing trims what's-next below the "
             "takeaway cap itself", wn == [f"a{i}" for i in range(4)], wn)


def check_no_orphan_actions(rep):
    """2026-09-11 real find (WAEPA): what's-next used to include EVERY
    thread's action regardless of whether that thread's meaning survived
    onto the takeaways slide -- 6 goal threads (all anchor="goal", so the
    cap has no signal to drop and falls back to trimming from the end)
    produced 4 takeaways but 6 actions, arguing points for two threads the
    client's Takeaways slide never showed."""
    print("\nNo orphan actions -- an action only survives if its OWN thread's "
         "meaning made the takeaways slide")
    threads = [_thread(f"Goal {i}", finding=f"f{i}", meaning=f"m{i}", action=f"a{i}")
              for i in range(6)]
    _hi, ta, wn = ra.distribute_threads(threads)
    rep.check("takeaways capped at 4", len(ta) == 4, ta)
    rep.check("what's-next has exactly as many actions as surviving takeaways, "
             "not one per original thread", len(wn) == 4, wn)
    kept_heads = {head for head, _meaning in ta}
    rep.check("every surviving action's own head is among the kept takeaway heads",
             all(f"a{i}" in wn for i in range(6) if f"Goal {i}" in kept_heads), (kept_heads, wn))
    dropped_indices = [i for i in range(6) if f"Goal {i}" not in kept_heads]
    rep.check("a dropped thread's action does NOT appear in what's-next",
             all(f"a{i}" not in wn for i in dropped_indices), (dropped_indices, wn))


def check_one_selection_both_slides_well_formed(rep):
    """2026-09-12 real find (WAEPA): the set of survivors must be chosen
    ONCE and shared by both slides -- not capped independently per surface.
    With every thread fully formed (finding AND meaning both present), the
    same 4 survivors must appear on both Highlights and Takeaways, and the
    one dropped signal must not leak onto either."""
    print("\nOne selection, both slides -- well-formed threads: highlights "
         "and takeaways show the identical surviving set")
    threads = [
        _thread("Market", finding="f-market", meaning="m-market", action="a-market"),
        _thread("Site", finding="f-site", meaning="m-site", action="a-site"),
        _thread("Frequency", finding="f-freq", meaning="m-freq", action=None),
        _thread("Zip", anchor="signal", finding="f-zip", meaning="m-zip", action="a-zip"),
        _thread("Direct", anchor="signal", finding="f-direct", meaning="m-direct", action=None),
    ]
    hi, ta, _wn = ra.distribute_threads(threads)
    hi_heads = {h for h, _d in hi}
    ta_heads = {h for h, _d in ta}
    rep.check("highlight heads equal takeaway heads when every survivor is fully formed",
             hi_heads == ta_heads, (hi_heads, ta_heads))
    goal_heads = {"Market", "Site", "Frequency"}
    rep.check("every goal thread survives on both surfaces",
             goal_heads <= hi_heads and goal_heads <= ta_heads, (hi_heads, ta_heads))
    rep.check("the least-prioritized signal (Direct) is dropped from BOTH, not just one",
             "Direct" not in hi_heads and "Direct" not in ta_heads, (hi_heads, ta_heads))


def check_one_selection_both_slides_asymmetric_fields(rep):
    """The exact real WAEPA shape: a goal thread missing its own finding
    (Frequency) and a signal thread missing its own meaning (Direct) used to
    survive independently per surface -- Direct landed on Highlights only
    (its pool never saw Frequency's absence and had room to spare), Frequency
    landed on Takeaways only (symmetric problem the other way). Now the cap
    runs ONCE on the ranked thread list itself, so it's the SAME four threads
    underneath both surfaces -- Direct, the less-prioritized signal, is
    dropped from both; Frequency survives and correctly appears on Takeaways
    only (a goal thread degrading to closing-only, never fabricated a
    finding it doesn't have) rather than one surface disagreeing with the
    other about which four threads survived at all."""
    print("\nOne selection, both slides -- the real WAEPA asymmetric-fields shape")
    threads = [
        _thread("Market", finding="f-market", meaning="m-market", action="a-market"),
        _thread("Site", finding="f-site", meaning="m-site", action="a-site"),
        _thread("Frequency", finding=None, meaning="m-freq", action=None),
        _thread("Zip", anchor="signal", finding="f-zip", meaning="m-zip", action="a-zip"),
        _thread("Direct", anchor="signal", finding="f-direct", meaning=None, action=None),
    ]
    hi, ta, _wn = ra.distribute_threads(threads)
    hi_heads = {h for h, _d in hi}
    ta_heads = {h for h, _d in ta}
    rep.check("highlight heads is a subset of takeaway heads", hi_heads <= ta_heads, (hi_heads, ta_heads))
    rep.check("Direct (the less-prioritized signal) is dropped from BOTH surfaces, "
             "not stranded on one", "Direct" not in hi_heads and "Direct" not in ta_heads,
             (hi_heads, ta_heads))
    rep.check("Frequency (goal, no finding) survives on takeaways only -- never "
             "fabricates a highlight it has no finding for",
             "Frequency" in ta_heads and "Frequency" not in hi_heads, (hi_heads, ta_heads))
    rep.check("Market and Site (fully formed goal threads) are on both",
             {"Market", "Site"} <= hi_heads and {"Market", "Site"} <= ta_heads, (hi_heads, ta_heads))


def check_malformed_input_degrades(rep):
    print("\nMalformed input degrades rather than raising")
    hi, ta, wn = ra.distribute_threads(None)
    rep.check("None -> all empty", (hi, ta, wn) == ([], [], []))
    hi, ta, wn = ra.distribute_threads([{"not": "a proper thread"}, "a bare string", 42])
    rep.check("junk entries -> all empty, no crash", (hi, ta, wn) == ([], [], []))


def check_every_highlight_has_a_takeaway(rep):
    print("\nEvery thread that produces a highlight also produces a takeaway "
         "(well-formed threads always carry meaning)")
    threads = [_thread("Goal 1", finding="f1", meaning="m1")]
    hi, ta, _wn = ra.distribute_threads(threads)
    rep.check("highlight head appears among takeaway heads too",
             hi[0][0] in [h for h, _d in ta], (hi, ta))


if __name__ == "__main__":
    rep = Report()
    check_basic_shape(rep)
    check_closing_only_signal(rep)
    check_no_action_no_whats_next_item(rep)
    check_goal_order_is_priority(rep)
    check_cap_drops_signals_first(rep)
    check_cap_drops_last_signal_first(rep)
    check_whats_next_has_no_cap_of_its_own(rep)
    check_no_orphan_actions(rep)
    check_one_selection_both_slides_well_formed(rep)
    check_one_selection_both_slides_asymmetric_fields(rep)
    check_malformed_input_degrades(rep)
    check_every_highlight_has_a_takeaway(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed out of {total}")
    sys.exit(1 if rep.failed else 0)
