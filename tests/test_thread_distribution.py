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


def check_whats_next_has_no_cap(rep):
    print("\nWhat's-next takes every action, no cap even past 4 threads")
    threads = [_thread(f"Goal {i}", finding=f"f{i}", meaning=f"m{i}", action=f"a{i}")
              for i in range(6)]
    _hi, _ta, wn = ra.distribute_threads(threads)
    rep.check("all 6 actions present", wn == [f"a{i}" for i in range(6)], wn)


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
    check_whats_next_has_no_cap(rep)
    check_malformed_input_degrades(rep)
    check_every_highlight_has_a_takeaway(rep)
    total = rep.passed + len(rep.failed)
    print(f"\n{rep.passed} passed, {len(rep.failed)} failed out of {total}")
    sys.exit(1 if rep.failed else 0)
