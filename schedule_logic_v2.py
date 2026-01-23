#!/usr/bin/env python3
"""
schedule_logic.py (v2 - backtracking solver)
Logic-first scheduler for Harpeth Hills FSAs (NO Excel formatting yet).

This version replaces the "multi-attempt greedy fill" with a true backtracking solver
with forward-checking. This avoids dead-ends like Feb 2026 (CJ off Feb 5–6) where
a greedy approach can accidentally push someone to a 5-day streak and then run out
of available people on a tight day.

Hard rules implemented (same as before):
- Observed holidays: no roles at all
- Roles per day:
  * Sun: PN only (5 hours)
  * Mon: PN, AN, W, BU1, BU2, BU3 (all 6)
  * Tue–Fri (normal): PN, AN, W, BU1, BU2 (5 people)
  * Push week (last Mon–Fri): PN, AN, W, BU1, BU2, BU3 (all 6)
  * Sat (normal): PN, AN
  * Push-week Sat (last Sat): PN, AN, BU1, BU2, BU3, BU4 (all 6; no W)
- Sunday PN rotation, and Sunday PN must be AN on Saturday prior
- One role per person per day
- No same specific role back-to-back days (BU1->BU2 allowed, BU1->BU1 not)
- Monday PN max 1 per month
- Max 5 consecutive work days (within month)
- Weekly hours cap (Sun–Sat) <= 40 OUTSIDE push week.
  NOTE: push week relaxation is implemented as: any day in push week ignores weekly cap checks.

Inputs each run:
1) time_off requests (hard) as dates per FSA
2) sales volumes per FSA (Total Volume = PAF + Cemetery)
3) Sunday rotation inputs (order + who is first Sunday assignee)

Output:
- Prints day-by-day schedule to console
- Prints validation results

Next phase later:
- Pay period (80) with cross-month carryover
- Preferences (e.g., Dawn "no PN/AN on Feb 2") as SOFT constraints
- Excel template rendering
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
import calendar
import random
from collections import defaultdict, Counter
from typing import Dict, List, Optional, Set, Iterable, Any, Tuple


@dataclass(frozen=True)
class Inputs:
    year: int
    month: int
    fsas: List[str]  # must be 6

    sales_volume: Dict[str, float]       # Total Volume = PAF + Cemetery
    seniority_order: List[str]           # tie-break for sales ties

    time_off: Dict[str, Set[date]]       # hard: cannot work these dates

    sunday_rotation_order: List[str]     # length 6
    sunday_first_assignee: str           # PN on first Sunday of month

    observed_holidays: Set[date]         # no roles at all

    push_week_enabled: bool = True

    hours_per_day: int = 8
    sunday_hours: int = 5

    weekly_cap_hours: int = 40
    max_consecutive_days: int = 5
    monday_pn_max_per_month: int = 1


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

def sort_by_hierarchy(fsas: List[str], sales: Dict[str, float], seniority: List[str]) -> List[str]:
    seniority_rank = {name: i for i, name in enumerate(seniority)}
    return sorted(fsas, key=lambda n: (-sales.get(n, 0.0), seniority_rank.get(n, 10_000)))

def last_weekdays_block(year: int, month: int) -> List[date]:
    days = month_days(year, month)
    last_day = days[-1]
    last_fri = last_day
    while last_fri.weekday() != 4:
        last_fri -= timedelta(days=1)
    last_mon = last_fri - timedelta(days=4)
    return [last_mon + timedelta(days=i) for i in range(5)]


BU_ROLES_TUE_FRI = ["BU1", "BU2"]
BU_ROLES_MON_ALL = ["BU1", "BU2", "BU3"]
BU_ROLES_PUSH_SAT = ["BU1", "BU2", "BU3", "BU4"]

def roles_for_day(d: date, inp: Inputs, push_days: Set[date]) -> List[str]:
    if d in inp.observed_holidays:
        return []
    if is_sunday(d):
        return ["PN"]
    if is_saturday(d):
        if inp.push_week_enabled and d in push_days:
            return ["PN", "AN"] + BU_ROLES_PUSH_SAT
        return ["PN", "AN"]
    if d.weekday() == 0:
        return ["PN", "AN", "W"] + BU_ROLES_MON_ALL
    if inp.push_week_enabled and d in push_days:
        return ["PN", "AN", "W"] + BU_ROLES_MON_ALL
    return ["PN", "AN", "W"] + BU_ROLES_TUE_FRI

def hours_for(d: date, role: str, inp: Inputs) -> int:
    if d in inp.observed_holidays:
        return 0
    if is_sunday(d):
        return inp.sunday_hours if role == "PN" else 0
    return inp.hours_per_day


def validate_schedule(sched: Dict[date, Dict[str, str]], inp: Inputs) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    days = month_days(inp.year, inp.month)

    push_days = set()
    if inp.push_week_enabled:
        push_days.update(last_weekdays_block(inp.year, inp.month))
        push_days.add(max(d for d in days if is_saturday(d)))

    def add(sev: str, msg: str, d: Optional[date] = None, who: Optional[str] = None, role: Optional[str] = None):
        issues.append({"severity": sev, "date": d, "person": who, "role": role, "message": msg})

    for d in days:
        required = roles_for_day(d, inp, push_days)
        assigned = set(sched.get(d, {}).keys())
        if not required:
            if assigned:
                add("HARD", "Holiday should have no roles assigned.", d=d)
            continue
        missing = set(required) - assigned
        extra = assigned - set(required)
        if missing:
            add("HARD", f"Missing required roles: {sorted(missing)}", d=d)
        if extra:
            add("HARD", f"Unexpected roles: {sorted(extra)}", d=d)

    for d in days:
        vals = list(sched.get(d, {}).values())
        for p, c in Counter(vals).items():
            if c > 1:
                add("HARD", "Person scheduled multiple times on same day.", d=d, who=p)

    for d in days:
        for role, p in sched.get(d, {}).items():
            if d in inp.time_off.get(p, set()):
                add("HARD", "Scheduled on a requested day off.", d=d, who=p, role=role)

    for i in range(1, len(days)):
        d, prev = days[i], days[i - 1]
        for role in set(sched.get(d, {})).intersection(sched.get(prev, {})):
            if sched[prev].get(role) == sched[d].get(role):
                add("HARD", "Same role assigned back-to-back days.", d=d, who=sched[d][role], role=role)

    mon_pn = Counter()
    for d in days:
        if d.weekday() == 0 and sched.get(d, {}).get("PN"):
            mon_pn[sched[d]["PN"]] += 1
    for p, c in mon_pn.items():
        if c > inp.monday_pn_max_per_month:
            add("HARD", f"Monday PN exceeds max ({c}).", who=p)

    for d in days:
        if is_sunday(d) and sched.get(d, {}).get("PN"):
            p = sched[d]["PN"]
            sat = d - timedelta(days=1)
            if sat.month == inp.month and "AN" in roles_for_day(sat, inp, push_days):
                if sched.get(sat, {}).get("AN") != p:
                    add("HARD", "Sunday PN must be AN on prior Saturday.", d=d, who=p)

    worked = defaultdict(set)
    for d in days:
        for _, p in sched.get(d, {}).items():
            worked[p].add(d)
    for p in inp.fsas:
        streak = 0
        for d in days:
            if d in worked[p]:
                streak += 1
                if streak > inp.max_consecutive_days:
                    add("HARD", f"Exceeded max consecutive days ({streak}).", d=d, who=p)
            else:
                streak = 0

    week_hours = defaultdict(Counter)
    for d in days:
        wk = week_start_sun(d)
        for role, p in sched.get(d, {}).items():
            week_hours[wk][p] += hours_for(d, role, inp)

    for wk, ctr in week_hours.items():
        week_contains_push = any((wk + timedelta(days=i)) in push_days for i in range(7))
        for p, h in ctr.items():
            if (not week_contains_push) and h > inp.weekly_cap_hours:
                add("HARD", f"Weekly hours exceed cap ({h}).", d=wk, who=p)

    return issues


class Solver:
    def __init__(self, inp: Inputs, seed: int = 7):
        self.inp = inp
        self.rng = random.Random(seed)
        self.days = month_days(inp.year, inp.month)

        self.push_days: Set[date] = set()
        if inp.push_week_enabled:
            self.push_days.update(last_weekdays_block(inp.year, inp.month))
            self.push_days.add(max(d for d in self.days if is_saturday(d)))

        self.sched: Dict[date, Dict[str, str]] = {d: {} for d in self.days}
        self.week_hours: Dict[date, Counter] = defaultdict(Counter)
        self.mon_pn = Counter()
        self.role_counts = {r: Counter() for r in ["PN", "AN", "W", "BU"]}

        self.targets = self.compute_targets()
        self.hierarchy = sort_by_hierarchy(inp.fsas, inp.sales_volume, inp.seniority_order)

    def compute_targets(self) -> Dict[str, Dict[str, int]]:
        fsas = self.inp.fsas
        hier = sort_by_hierarchy(fsas, self.inp.sales_volume, self.inp.seniority_order)

        pn_days = [d for d in self.days if d not in self.inp.observed_holidays]
        pn_slots = len(pn_days)
        an_days = [d for d in pn_days if not is_sunday(d)]
        an_slots = len(an_days)
        w_days = [d for d in self.days if "W" in roles_for_day(d, self.inp, self.push_days)]
        w_slots = len(w_days)

        def split(total: int) -> Dict[str, int]:
            base = total // len(fsas)
            rem = total % len(fsas)
            out = {p: base for p in fsas}
            for p in hier[:rem]:
                out[p] += 1
            return out

        return {"PN": split(pn_slots), "AN": split(an_slots), "W": split(w_slots)}

    def is_available(self, p: str, d: date) -> bool:
        return d not in self.inp.time_off.get(p, set())

    def consecutive_streak_if_work(self, p: str, d: date) -> int:
        streak = 1
        cur = d - timedelta(days=1)
        while cur.month == d.month:
            if p in self.sched[cur].values():
                streak += 1
                cur -= timedelta(days=1)
            else:
                break
        return streak

    def can_assign(self, p: str, d: date, role: str) -> Tuple[bool, str]:
        inp = self.inp
        if not self.is_available(p, d):
            return False, "time_off"
        if p in self.sched[d].values():
            return False, "already_today"
        prev = d - timedelta(days=1)
        if prev.month == d.month and self.sched.get(prev, {}).get(role) == p:
            return False, "same_role_back_to_back"
        if role == "PN" and d.weekday() == 0 and self.mon_pn[p] >= inp.monday_pn_max_per_month:
            return False, "monday_pn_limit"
        if self.consecutive_streak_if_work(p, d) > inp.max_consecutive_days:
            return False, "max_consecutive_days"

        wk = week_start_sun(d)
        add_h = hours_for(d, role, inp)
        if d not in self.push_days:
            if self.week_hours[wk][p] + add_h > inp.weekly_cap_hours:
                return False, "weekly_cap"
        return True, "ok"

    def assign(self, d: date, role: str, p: str):
        self.sched[d][role] = p
        wk = week_start_sun(d)
        h = hours_for(d, role, self.inp)
        self.week_hours[wk][p] += h

        if role in ["PN", "AN", "W"]:
            self.role_counts[role][p] += 1
        else:
            self.role_counts["BU"][p] += 1

        if role == "PN" and d.weekday() == 0:
            self.mon_pn[p] += 1

    def unassign(self, d: date, role: str):
        p = self.sched[d].pop(role, None)
        if p is None:
            return
        wk = week_start_sun(d)
        h = hours_for(d, role, self.inp)
        self.week_hours[wk][p] -= h

        if role in ["PN", "AN", "W"]:
            self.role_counts[role][p] -= 1
        else:
            self.role_counts["BU"][p] -= 1

        if role == "PN" and d.weekday() == 0:
            self.mon_pn[p] -= 1

    def set_sundays_by_rotation(self):
        order = self.inp.sunday_rotation_order
        if len(order) != 6:
            raise ValueError("sunday_rotation_order must have 6 names.")
        if self.inp.sunday_first_assignee not in order:
            raise ValueError("sunday_first_assignee must be in sunday_rotation_order.")

        idx = order.index(self.inp.sunday_first_assignee)
        sundays = [d for d in self.days if is_sunday(d) and d not in self.inp.observed_holidays]

        for d in sundays:
            chosen = None
            for k in range(6):
                cand = order[(idx + k) % 6]
                ok, _ = self.can_assign(cand, d, "PN")
                if ok:
                    chosen = cand
                    break
            if chosen is None:
                raise RuntimeError(f"Could not assign Sunday PN on {d}.")
            self.assign(d, "PN", chosen)
            idx = (order.index(chosen) + 1) % 6

    def enforce_sat_an_for_sun(self):
        for d in self.days:
            if is_sunday(d) and self.sched[d].get("PN"):
                p = self.sched[d]["PN"]
                sat = d - timedelta(days=1)
                if sat.month != self.inp.month:
                    continue
                if sat in self.inp.observed_holidays:
                    continue
                req = roles_for_day(sat, self.inp, self.push_days)
                if "AN" not in req:
                    continue
                if self.sched[sat].get("AN") and self.sched[sat]["AN"] != p:
                    raise RuntimeError(f"Saturday AN conflict on {sat}: needs {p}, has {self.sched[sat]['AN']}.")
                if "AN" not in self.sched[sat]:
                    ok, reason = self.can_assign(p, sat, "AN")
                    if not ok:
                        raise RuntimeError(f"Cannot enforce Saturday AN for Sunday PN: {p} on {sat} ({reason}).")
                    self.assign(sat, "AN", p)

    def candidate_list(self, d: date, role: str) -> List[str]:
        cands = []
        for p in self.inp.fsas:
            ok, _ = self.can_assign(p, d, role)
            if ok:
                cands.append(p)
        if not cands:
            return []

        hier_rank = {p: i for i, p in enumerate(self.hierarchy)}

        def need(p: str) -> int:
            if role in self.targets:
                return self.targets[role][p] - self.role_counts[role][p]
            return 0

        def key(p: str):
            return (need(p), -hier_rank.get(p, 999), self.rng.random())

        cands.sort(key=key, reverse=True)
        band = max(2, min(4, len(cands)))
        top = cands[:band]
        self.rng.shuffle(top)
        return top + cands[band:]

    def next_unfilled_slot(self) -> Optional[Tuple[date, str]]:
        best = None
        best_len = 10**9
        for d in self.days:
            req = roles_for_day(d, self.inp, self.push_days)
            if not req:
                continue
            # Fill primaries first within each day
            prim_first = [r for r in ["PN", "AN", "W"] if r in req] + [r for r in req if r not in ["PN", "AN", "W"]]
            for role in prim_first:
                if role in self.sched[d]:
                    continue
                cands = self.candidate_list(d, role)
                n = len(cands)
                if n == 0:
                    return (d, role)
                if n < best_len:
                    best_len = n
                    best = (d, role)
        return best

    def solve(self, max_nodes: int = 400000) -> Dict[date, Dict[str, str]]:
        self.set_sundays_by_rotation()
        self.enforce_sat_an_for_sun()

        nodes = 0

        def backtrack() -> bool:
            nonlocal nodes
            nodes += 1
            if nodes > max_nodes:
                return False

            slot = self.next_unfilled_slot()
            if slot is None:
                return True

            d, role = slot
            cands = self.candidate_list(d, role)
            if not cands:
                return False

            for p in cands:
                ok, _ = self.can_assign(p, d, role)
                if not ok:
                    continue
                self.assign(d, role, p)
                if backtrack():
                    return True
                self.unassign(d, role)
            return False

        if not backtrack():
            raise RuntimeError("Failed to build a valid schedule under current constraints (backtracking exhausted).")
        return self.sched


def print_schedule(sched: Dict[date, Dict[str, str]], inp: Inputs):
    days = month_days(inp.year, inp.month)
    push_days = set()
    if inp.push_week_enabled:
        push_days.update(last_weekdays_block(inp.year, inp.month))
        push_days.add(max(d for d in days if is_saturday(d)))

    for d in days:
        req = roles_for_day(d, inp, push_days)
        if not req:
            print(f"{d} (HOLIDAY)")
            continue
        a = sched.get(d, {})
        ordered = []
        for r in ["PN", "AN", "W", "BU1", "BU2", "BU3", "BU4"]:
            if r in req:
                ordered.append(f"{r}:{a.get(r,'-')}")
        print(f"{d}  " + "  ".join(ordered))


if __name__ == "__main__":
    FSAS = ["Mark", "Dawn", "Will", "Kyle", "CJ", "Greg"]

    # Replace with your real Total Volume (PAF + Cemetery) each run
    sales = {
        "Mark": 1000000,
        "Dawn": 900000,
        "Will": 850000,
        "Kyle": 800000,
        "CJ": 750000,
        "Greg": 700000,
    }

    seniority = ["Mark", "Dawn", "Will", "Kyle", "CJ", "Greg"]

    holidays: Set[date] = set()  # none in Feb 2026

    # Hard time off
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

    solver = Solver(inp, seed=11)
    sched = solver.solve(max_nodes=400000)

    print_schedule(sched, inp)

    violations = validate_schedule(sched, inp)
    hard = [v for v in violations if v["severity"] == "HARD"]
    print("\nVALIDATION")
    print(f"Hard violations: {len(hard)}")
    if hard:
        for v in hard[:60]:
            print(v)
    else:
        print("✅ No hard violations detected.")
