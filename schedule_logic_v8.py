#!/usr/bin/env python3
"""
schedule_logic_v8.py
Two-phase scheduler (PRIMARY roles first, then BU backfill) with "BU is scalable" behavior.

Why v8:
- Your primary roles are non-negotiable.
- BU roles are *extra coverage* and, per your push-week description, are the easiest hours to scale back.
- In real life, if hard time-off + max-consecutive-days makes it impossible to staff BU2 (or BU3/BU4),
  we should simply run short on BU that day rather than failing the entire month.

Change vs v7:
- Phase 1 unchanged (solve PN/AN/W for the entire month).
- Phase 2: BU filler now fills BU roles day-by-day and will STOP adding BU roles for that day
  when it hits an infeasible BU slot. It records a "shortage" entry.
  (So Feb 6 in your example becomes a BU shortage day rather than a full failure.)

You still get:
- Progress lines
- A shortage report at end (dates where BU roles could not be fully staffed)
- Hard constraints still enforced for any roles that ARE assigned:
  * hard time off, one role per day, max consecutive days, weekly cap (with push-week waiver),
    and no same specific BU role back-to-back.

No external packages required.
"""

from __future__ import annotations
from dataclasses import dataclass
from datetime import date, timedelta
import calendar
import random
import sys
import time
from collections import defaultdict, Counter
from typing import Dict, List, Optional, Set, Iterable, Any, Tuple


@dataclass(frozen=True)
class Inputs:
    year: int
    month: int
    fsas: List[str]
    sales_volume: Dict[str, float]
    seniority_order: List[str]
    time_off: Dict[str, Set[date]]
    sunday_rotation_order: List[str]
    sunday_first_assignee: str
    observed_holidays: Set[date]
    push_week_enabled: bool = True
    hours_per_day: int = 8
    sunday_hours: int = 5
    weekly_cap_hours: int = 40
    max_consecutive_days: int = 5
    monday_pn_max_per_month: int = 1


# -----------------------------
# Utilities
# -----------------------------

def daterange(d1: date, d2: date) -> Iterable[date]:
    cur = d1
    while cur <= d2:
        yield cur
        cur += timedelta(days=1)

def month_days(year: int, month: int) -> List[date]:
    _, last = calendar.monthrange(year, month)
    return [date(year, month, d) for d in range(1, last + 1)]

def is_saturday(d: date) -> bool:
    return d.weekday() == 5

def is_sunday(d: date) -> bool:
    return d.weekday() == 6

def week_start_sun(d: date) -> date:
    return d - timedelta(days=(d.weekday() + 1) % 7)

def last_weekdays_block(year: int, month: int) -> List[date]:
    days = month_days(year, month)
    last_day = days[-1]
    last_fri = last_day
    while last_fri.weekday() != 4:
        last_fri -= timedelta(days=1)
    last_mon = last_fri - timedelta(days=4)
    return [last_mon + timedelta(days=i) for i in range(5)]

def hours_for(d: date, role: str, inp: Inputs) -> int:
    if d in inp.observed_holidays:
        return 0
    if is_sunday(d):
        return inp.sunday_hours if role == "PN" else 0
    return inp.hours_per_day

def sort_by_hierarchy(fsas: List[str], sales: Dict[str, float], seniority: List[str]) -> List[str]:
    seniority_rank = {name: i for i, name in enumerate(seniority)}
    return sorted(fsas, key=lambda n: (-sales.get(n, 0.0), seniority_rank.get(n, 10_000)))


# -----------------------------
# Push days & cap waiver
# -----------------------------

def compute_push_days(inp: Inputs) -> Set[date]:
    if not inp.push_week_enabled:
        return set()
    days = month_days(inp.year, inp.month)
    push = set(last_weekdays_block(inp.year, inp.month))
    push.add(max(d for d in days if is_saturday(d)))
    return push

def week_contains_push(wk_start: date, push_days: Set[date]) -> bool:
    return any((wk_start + timedelta(days=i)) in push_days for i in range(7))


# -----------------------------
# Role model
# -----------------------------

PRIMARY_ROLES_ORDER = ["PN", "AN", "W"]
BU_POOL_WEEKDAY = ["BU1", "BU2", "BU3"]
BU_POOL_SAT_PUSH = ["BU1", "BU2", "BU3", "BU4"]

def required_primary_roles(d: date, inp: Inputs) -> List[str]:
    if d in inp.observed_holidays:
        return []
    if is_sunday(d):
        return ["PN"]
    if is_saturday(d):
        return ["PN", "AN"]
    return ["PN", "AN", "W"]

def desired_bu_roles(d: date, inp: Inputs, push_days: Set[date]) -> List[str]:
    if d in inp.observed_holidays:
        return []
    if is_sunday(d):
        return []

    available = [p for p in inp.fsas if d not in inp.time_off.get(p, set())]
    cap = len(available)
    prim_count = len(required_primary_roles(d, inp))

    if is_saturday(d):
        if d in push_days:
            desired = BU_POOL_SAT_PUSH[:]
            allow = max(0, min(len(desired), cap - prim_count))
            return desired[:allow]
        return []

    if d.weekday() == 0 or d in push_days:
        desired = BU_POOL_WEEKDAY[:3]
    else:
        desired = BU_POOL_WEEKDAY[:2]

    allow = max(0, min(len(desired), cap - prim_count))
    return desired[:allow]


def compute_primary_targets(inp: Inputs) -> Dict[str, Dict[str, int]]:
    days = month_days(inp.year, inp.month)
    fsas = inp.fsas
    hier = sort_by_hierarchy(fsas, inp.sales_volume, inp.seniority_order)

    totals = {r: 0 for r in PRIMARY_ROLES_ORDER}
    for d in days:
        for r in required_primary_roles(d, inp):
            if r in totals:
                totals[r] += 1

    def split(total: int) -> Dict[str, int]:
        base = total // len(fsas)
        rem = total % len(fsas)
        out = {p: base for p in fsas}
        for p in hier[:rem]:
            out[p] += 1
        return out

    return {r: split(totals[r]) for r in PRIMARY_ROLES_ORDER}

def worked_today(roles_for_day: Dict[str, str], person: str) -> bool:
    return person in roles_for_day.values()

def consecutive_streak(sched: Dict[date, Dict[str, str]], days_sorted: List[date], person: str, up_to_index: int) -> int:
    streak = 0
    i = up_to_index
    while i >= 0:
        d = days_sorted[i]
        if worked_today(sched[d], person):
            streak += 1
            i -= 1
        else:
            break
    return streak


# -----------------------------
# Phase 1: Primary search
# -----------------------------

class PrimarySolver:
    def __init__(self, inp: Inputs, seed: int = 11):
        self.inp = inp
        self.rng = random.Random(seed)
        self.days = month_days(inp.year, inp.month)
        self.push_days = compute_push_days(inp)

        self._validate_inputs()

        self.sched: Dict[date, Dict[str, str]] = {d: {} for d in self.days}
        self.week_hours: Dict[date, Counter] = defaultdict(Counter)
        self.mon_pn = Counter()
        self.primary_counts = {r: Counter() for r in PRIMARY_ROLES_ORDER}

        self.hierarchy = sort_by_hierarchy(inp.fsas, inp.sales_volume, inp.seniority_order)
        self.targets = compute_primary_targets(inp)

        self.first_dead_end: Optional[Dict[str, Any]] = None

    def _validate_inputs(self):
        if not self.inp.fsas:
            raise ValueError("Inputs.fsas is empty.")
        if len(set(self.inp.fsas)) != len(self.inp.fsas):
            raise ValueError("Inputs.fsas contains duplicates.")
        rot = self.inp.sunday_rotation_order
        if len(set(rot)) != len(rot):
            raise ValueError("sunday_rotation_order contains duplicates.")
        if set(rot) != set(self.inp.fsas):
            raise ValueError("sunday_rotation_order must match fsas exactly.")
        if self.inp.sunday_first_assignee not in rot:
            raise ValueError("sunday_first_assignee must be in sunday_rotation_order.")

    def can_assign_primary(self, p: str, d: date, role: str) -> Tuple[bool, str]:
        inp = self.inp

        if d in inp.time_off.get(p, set()):
            return False, "time_off"
        if p in self.sched[d].values():
            return False, "already_today"

        prev = d - timedelta(days=1)
        if prev.month == d.month and self.sched.get(prev, {}).get(role) == p:
            return False, "same_role_back_to_back"

        if role == "PN" and d.weekday() == 0 and self.mon_pn[p] >= inp.monday_pn_max_per_month:
            return False, "monday_pn_limit"

        # weekly cap (waived if week contains push day)
        wk = week_start_sun(d)
        waive = week_contains_push(wk, self.push_days)
        add_h = hours_for(d, role, inp)
        if (not waive) and (self.week_hours[wk][p] + add_h > inp.weekly_cap_hours):
            return False, "weekly_cap"

        # Saturday AN must match Sunday PN if Sunday PN already assigned
        if role == "AN" and is_saturday(d):
            nxt = d + timedelta(days=1)
            if nxt.month == inp.month and is_sunday(nxt) and nxt not in inp.observed_holidays:
                sunday_pn = self.sched[nxt].get("PN")
                if sunday_pn is not None and p != sunday_pn:
                    return False, "sat_an_must_match_sunday_pn"

        return True, "ok"

    def assign(self, d: date, role: str, p: str):
        self.sched[d][role] = p
        wk = week_start_sun(d)
        self.week_hours[wk][p] += hours_for(d, role, self.inp)
        if role == "PN" and d.weekday() == 0:
            self.mon_pn[p] += 1
        self.primary_counts[role][p] += 1

    def unassign(self, d: date, role: str):
        p = self.sched[d].pop(role, None)
        if p is None:
            return
        wk = week_start_sun(d)
        self.week_hours[wk][p] -= hours_for(d, role, self.inp)
        if role == "PN" and d.weekday() == 0:
            self.mon_pn[p] -= 1
        self.primary_counts[role][p] -= 1

    def set_sundays_by_rotation(self, first_assignee: Optional[str] = None):
        order = self.inp.sunday_rotation_order
        n = len(order)
        start = first_assignee or self.inp.sunday_first_assignee
        idx = order.index(start)

        sundays = [d for d in self.days if is_sunday(d) and d not in self.inp.observed_holidays]
        for d in sundays:
            chosen = None
            for k in range(n):
                cand = order[(idx + k) % n]
                ok, _ = self.can_assign_primary(cand, d, "PN")
                if ok:
                    chosen = cand
                    break
            if chosen is None:
                raise RuntimeError(f"Could not assign Sunday PN on {d} (rotation candidates unavailable).")
            self.assign(d, "PN", chosen)
            idx = (order.index(chosen) + 1) % n

    def _need_score(self, role: str, p: str) -> int:
        return self.targets[role][p] - self.primary_counts[role][p]

    def candidate_list(self, d: date, role: str) -> List[str]:
        cands = [p for p in self.inp.fsas if self.can_assign_primary(p, d, role)[0]]
        if not cands:
            return []
        hier_rank = {p: i for i, p in enumerate(self.hierarchy)}
        cands.sort(key=lambda p: (self._need_score(role, p), -hier_rank.get(p, 999)), reverse=True)
        band = max(2, min(4, len(cands)))
        top = cands[:band]
        self.rng.shuffle(top)
        return top + cands[band:]

    def candidate_count(self, d: date, role: str) -> int:
        return sum(1 for p in self.inp.fsas if self.can_assign_primary(p, d, role)[0])

    def next_unfilled_primary_slot(self) -> Optional[Tuple[date, str]]:
        best = None
        best_len = 10**9
        for d in self.days:
            req = required_primary_roles(d, self.inp)
            if not req:
                continue
            for role in PRIMARY_ROLES_ORDER:
                if role not in req:
                    continue
                if role in self.sched[d]:
                    continue
                n = self.candidate_count(d, role)
                if n == 0:
                    return (d, role)
                if n < best_len:
                    best_len = n
                    best = (d, role)
        return best

    def explain_no_candidates(self, d: date, role: str) -> Dict[str, Any]:
        reasons = {}
        for p in self.inp.fsas:
            ok, why = self.can_assign_primary(p, d, role)
            reasons[p] = "OK" if ok else why
        return {"phase": "PRIMARY", "date": d, "role": role, "reasons": reasons, "assigned_today": dict(self.sched[d])}

    def solve(self, max_nodes: int = 2_000_000, report_every: int = 25_000, spinner: bool = True) -> Dict[date, Dict[str, str]]:
        # Sunday fallback trials
        order = self.inp.sunday_rotation_order
        start_idx = order.index(self.inp.sunday_first_assignee)
        seed_trials = 3

        last_err = None
        for s_try in range(seed_trials):
            for shift in range(len(order)):
                # reset state
                self.sched = {d: {} for d in self.days}
                self.week_hours = defaultdict(Counter)
                self.mon_pn = Counter()
                self.primary_counts = {r: Counter() for r in PRIMARY_ROLES_ORDER}
                self.first_dead_end = None
                self.rng = random.Random(11 + s_try * 101 + shift * 17)

                first = order[(start_idx + shift) % len(order)]
                try:
                    self.set_sundays_by_rotation(first_assignee=first)
                except RuntimeError as e:
                    last_err = e
                    continue

                try:
                    return self._backtrack(max_nodes=max_nodes, report_every=report_every, spinner=spinner)
                except RuntimeError as e:
                    last_err = e
                    continue

        raise RuntimeError(f"Primary phase failed after fallback attempts. Last error: {last_err}") from last_err

    def _backtrack(self, max_nodes: int, report_every: int, spinner: bool) -> Dict[date, Dict[str, str]]:
        t0 = time.time()
        last_report_t = t0
        nodes = 0
        spin_chars = "|/-\\"
        spin_i = 0

        def progress_line(depth: int, slot: Optional[Tuple[date, str]]):
            nonlocal spin_i
            now = time.time()
            elapsed = now - t0
            rate = nodes / elapsed if elapsed > 0 else 0.0
            s = spin_chars[spin_i % len(spin_chars)] if spinner else ""
            spin_i += 1
            if slot:
                dd, rr = slot
                msg = f"{s} [PRIMARY] nodes={nodes:,} rate={rate:,.0f}/s depth={depth} slot={dd} {rr}"
            else:
                msg = f"{s} [PRIMARY] nodes={nodes:,} rate={rate:,.0f}/s depth={depth} slot=done"
            sys.stdout.write("\r" + msg[:140].ljust(140))
            sys.stdout.flush()

        def backtrack(depth: int) -> bool:
            nonlocal nodes, last_report_t
            nodes += 1
            if nodes > max_nodes:
                return False

            slot = self.next_unfilled_primary_slot()
            if slot is None:
                progress_line(depth, None)
                return True

            dd, rr = slot
            cands = self.candidate_list(dd, rr)
            if not cands:
                if self.first_dead_end is None:
                    self.first_dead_end = self.explain_no_candidates(dd, rr)
                return False

            now = time.time()
            if nodes % report_every == 0 or (now - last_report_t) > 2.0:
                progress_line(depth, slot)
                last_report_t = now

            for p in cands:
                ok, _ = self.can_assign_primary(p, dd, rr)
                if not ok:
                    continue
                # tentative assign
                self.assign(dd, rr, p)
                # max consecutive WORK days (primaries only) exact check
                idx = self.days.index(dd)
                if consecutive_streak(self.sched, self.days, p, idx) > self.inp.max_consecutive_days:
                    self.unassign(dd, rr)
                    continue
                if backtrack(depth + 1):
                    return True
                self.unassign(dd, rr)
            return False

        ok = backtrack(0)
        sys.stdout.write("\n")
        sys.stdout.flush()

        if not ok:
            if self.first_dead_end:
                de = self.first_dead_end
                print("❌ PRIMARY phase failed. First dead-end:")
                print(f"  Date: {de['date']}  Role: {de['role']}")
                print(f"  Already assigned that day: {de['assigned_today']}")
                print("  Block reasons:")
                for p, why in de["reasons"].items():
                    print(f"    - {p}: {why}")
            raise RuntimeError("Failed to assign all primary roles.")
        return self.sched


# -----------------------------
# Phase 2: BU backfill (scalable)
# -----------------------------

class BUBackfill:
    def __init__(self, inp: Inputs, primary_sched: Dict[date, Dict[str, str]], seed: int = 11):
        self.inp = inp
        self.rng = random.Random(seed)
        self.days = month_days(inp.year, inp.month)
        self.push_days = compute_push_days(inp)

        # start from primaries
        self.sched: Dict[date, Dict[str, str]] = {d: dict(primary_sched.get(d, {})) for d in self.days}

        # weekly hours from all assigned roles so far
        self.week_hours: Dict[date, Counter] = defaultdict(Counter)
        for d in self.days:
            wk = week_start_sun(d)
            for role, p in self.sched[d].items():
                self.week_hours[wk][p] += hours_for(d, role, inp)

        # record days where BU demand couldn't be fully met
        self.shortages: List[Dict[str, Any]] = []

    def _max_consecutive_if_assign(self, p: str, d: date) -> int:
        """Consecutive workday streak ending at d if p works on d (including existing assignments)."""
        idx = self.days.index(d)
        streak = 1
        i = idx - 1
        while i >= 0:
            dd = self.days[i]
            if p in self.sched[dd].values():
                streak += 1
                i -= 1
            else:
                break
        return streak

    def can_assign_bu(self, p: str, d: date, bu_role: str) -> Tuple[bool, str]:
        inp = self.inp

        if d in inp.time_off.get(p, set()):
            return False, "time_off"
        if p in self.sched[d].values():
            return False, "already_today"

        prev = d - timedelta(days=1)
        if prev.month == d.month and self.sched.get(prev, {}).get(bu_role) == p:
            return False, "same_bu_role_back_to_back"

        if self._max_consecutive_if_assign(p, d) > inp.max_consecutive_days:
            return False, "max_consecutive_days"

        wk = week_start_sun(d)
        waive = week_contains_push(wk, self.push_days)
        add_h = hours_for(d, bu_role, inp)
        if (not waive) and (self.week_hours[wk][p] + add_h > inp.weekly_cap_hours):
            return False, "weekly_cap"

        return True, "ok"

    def assign_bu(self, d: date, bu_role: str, p: str):
        self.sched[d][bu_role] = p
        wk = week_start_sun(d)
        self.week_hours[wk][p] += hours_for(d, bu_role, self.inp)

    def candidates(self, d: date, bu_role: str) -> List[str]:
        cands = [p for p in self.inp.fsas if self.can_assign_bu(p, d, bu_role)[0]]
        if not cands:
            return []
        # prefer lower week-hours to keep caps safe, then randomize among top band
        wk = week_start_sun(d)
        cands.sort(key=lambda p: (self.week_hours[wk][p], p))
        band = max(2, min(4, len(cands)))
        top = cands[:band]
        self.rng.shuffle(top)
        return top + cands[band:]

    def fill_day(self, d: date):
        desired = desired_bu_roles(d, self.inp, self.push_days)
        if not desired:
            return

        unfilled = []
        for bu_role in desired:
            if bu_role in self.sched[d]:
                continue
            cands = self.candidates(d, bu_role)
            if not cands:
                # scalable: stop trying to add further BU roles for this day
                unfilled.append(bu_role)
                # record why nobody could take it
                reasons = {p: ("OK" if self.can_assign_bu(p, d, bu_role)[0] else self.can_assign_bu(p, d, bu_role)[1])
                           for p in self.inp.fsas}
                self.shortages.append({
                    "date": d,
                    "missing_from": bu_role,
                    "missing_roles": [bu_role] + [r for r in desired if r not in self.sched[d] and r != bu_role],
                    "assigned_today": dict(self.sched[d]),
                    "reasons_for_first_missing": reasons
                })
                break

            # choose best candidate (already sorted)
            chosen = cands[0]
            self.assign_bu(d, bu_role, chosen)

    def fill_month(self, report_every_days: int = 3):
        t0 = time.time()
        for i, d in enumerate(self.days):
            self.fill_day(d)
            if (i % report_every_days) == 0:
                elapsed = time.time() - t0
                sys.stdout.write(f"\r[BU] day {i+1}/{len(self.days)} elapsed {elapsed:0.1f}s")
                sys.stdout.flush()
        sys.stdout.write("\n")
        sys.stdout.flush()
        return self.sched


# -----------------------------
# High-level API
# -----------------------------

def build_schedule(inp: Inputs, seed: int = 11) -> Tuple[Dict[date, Dict[str, str]], List[Dict[str, Any]]]:
    prim = PrimarySolver(inp, seed=seed)
    primary_sched = prim.solve(max_nodes=2_000_000, report_every=25_000, spinner=True)

    bu = BUBackfill(inp, primary_sched, seed=seed)
    full = bu.fill_month(report_every_days=2)

    return full, bu.shortages


def print_schedule(sched: Dict[date, Dict[str, str]], inp: Inputs):
    days = month_days(inp.year, inp.month)
    push_days = compute_push_days(inp)
    for d in days:
        prim = required_primary_roles(d, inp)
        bus = desired_bu_roles(d, inp, push_days)
        req = prim + bus
        if not req:
            print(f"{d} (HOLIDAY)")
            continue
        a = sched.get(d, {})
        parts = []
        for r in ["PN", "AN", "W", "BU1", "BU2", "BU3", "BU4"]:
            if r in req:
                parts.append(f"{r}:{a.get(r,'-')}")
        print(f"{d}  " + "  ".join(parts))


def print_shortages(shortages: List[Dict[str, Any]]):
    if not shortages:
        print("\n✅ No BU shortages (all desired BU roles were staffed).")
        return
    print(f"\n⚠️  BU SHORTAGES: {len(shortages)} day(s) ran short on BU coverage.\n")
    for s in shortages:
        d = s["date"]
        miss = s["missing_roles"]
        print(f"- {d}: missing {miss}  assigned={s['assigned_today']}")
        # show condensed reasons for the first missing role
        reasons = s["reasons_for_first_missing"]
        why_counts = Counter(reasons.values())
        why_summary = ", ".join(f"{k}:{v}" for k, v in why_counts.items())
        print(f"  blockers summary: {why_summary}")


# -----------------------------
# Example run (Feb 2026)
# -----------------------------
if __name__ == "__main__":
    FSAS = ["Mark", "Dawn", "Will", "Kyle", "CJ", "Greg"]

    sales = {"Mark": 1000000, "Dawn": 900000, "Will": 850000, "Kyle": 800000, "CJ": 750000, "Greg": 700000}
    seniority = ["Mark", "Dawn", "Will", "Kyle", "CJ", "Greg"]

    holidays: Set[date] = set()

    toff = {p: set() for p in FSAS}
    for dd in daterange(date(2026, 2, 7), date(2026, 2, 12)):
        toff["Will"].add(dd)
    toff["CJ"].update({date(2026, 2, 5), date(2026, 2, 6)})
    toff["Dawn"].add(date(2026, 2, 25))

    inp = Inputs(
        year=2026,
        month=2,
        fsas=FSAS,
        sales_volume=sales,
        seniority_order=seniority,
        time_off=toff,
        sunday_rotation_order=["Greg", "Mark", "Dawn", "Will", "CJ", "Kyle"],
        sunday_first_assignee="Greg",
        observed_holidays=holidays,
        push_week_enabled=True,
    )

    sched, shortages = build_schedule(inp, seed=11)
    print_schedule(sched, inp)
    print_shortages(shortages)
