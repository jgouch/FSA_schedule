#!/usr/bin/env python3
"""
schedule_logic_v4.py
Logic-first scheduler for Harpeth Hills FSAs (NO Excel formatting yet).

Fixes vs v3 (per Codex review):
1) Weekly cap rule aligned with validation:
   - Cap is waived for ANY week (Sun–Sat) that contains a push day.
2) Sunday rotation generalized:
   - Uses len(rotation_order) not hard-coded 6; validates against fsas set.
3) Saturday AN-before-Sunday-PN constraint integrated into backtracking:
   - No pre-assignment hard-fail; enforced via can_assign restriction for Saturday AN.
4) Input validation (fsas non-empty; rotation members; duplicates).
5) MRV selection does not consume RNG:
   - Candidate counting is deterministic; randomness only used when ordering actual tries.
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
    fsas: List[str]  # typically 6

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

def hours_for(d: date, role: str, inp: Inputs) -> int:
    if d in inp.observed_holidays:
        return 0
    if is_sunday(d):
        return inp.sunday_hours if role == "PN" else 0
    return inp.hours_per_day


def compute_push_days(inp: Inputs) -> Set[date]:
    days = month_days(inp.year, inp.month)
    if not inp.push_week_enabled:
        return set()
    push = set(last_weekdays_block(inp.year, inp.month))
    push.add(max(d for d in days if is_saturday(d)))
    return push

def week_contains_push(wk_start: date, push_days: Set[date]) -> bool:
    return any((wk_start + timedelta(days=i)) in push_days for i in range(7))

def available_fsas_on(d: date, inp: Inputs) -> List[str]:
    return [p for p in inp.fsas if d not in inp.time_off.get(p, set())]


BU_POOL_WEEKDAY = ["BU1", "BU2", "BU3"]
BU_POOL_SAT_PUSH = ["BU1", "BU2", "BU3", "BU4"]

def roles_for_day(d: date, inp: Inputs, push_days: Set[date]) -> List[str]:
    if d in inp.observed_holidays:
        return []
    cap = len(available_fsas_on(d, inp))
    if cap <= 0:
        return []

    if is_sunday(d):
        return ["PN"]

    if is_saturday(d):
        prim = ["PN", "AN"]
        if cap < len(prim):
            return prim[:cap]
        if d in push_days:
            desired_total = min(6, cap)
            bu_needed = max(0, desired_total - len(prim))
            return prim + BU_POOL_SAT_PUSH[:bu_needed]
        return prim

    prim = ["PN", "AN", "W"]
    if cap < len(prim):
        return prim[:cap]

    if d.weekday() == 0 or d in push_days:
        desired_total = min(6, cap)
    else:
        desired_total = min(5, cap)

    bu_needed = max(0, desired_total - len(prim))
    return prim + BU_POOL_WEEKDAY[:bu_needed]


def validate_schedule(sched: Dict[date, Dict[str, str]], inp: Inputs) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    days = month_days(inp.year, inp.month)
    push_days = compute_push_days(inp)

    def add(sev: str, msg: str, d: Optional[date] = None, who: Optional[str] = None, role: Optional[str] = None):
        issues.append({"severity": sev, "date": d, "person": who, "role": role, "message": msg})

    for d in days:
        required = roles_for_day(d, inp, push_days)
        assigned = set(sched.get(d, {}).keys())
        if not required:
            if assigned:
                add("HARD", "Holiday/no-staff day should have no roles assigned.", d=d)
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
            if sat.month == inp.month:
                req_sat = roles_for_day(sat, inp, push_days)
                if "AN" in req_sat and sched.get(sat, {}).get("AN") != p:
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
        waive = week_contains_push(wk, push_days)
        for p, h in ctr.items():
            if (not waive) and h > inp.weekly_cap_hours:
                add("HARD", f"Weekly hours exceed cap ({h}).", d=wk, who=p)

    return issues


class Solver:
    def __init__(self, inp: Inputs, seed: int = 7):
        self.inp = inp
        self.rng = random.Random(seed)
        self.days = month_days(inp.year, inp.month)
        self.push_days = compute_push_days(inp)

        self._validate_inputs()

        self.sched: Dict[date, Dict[str, str]] = {d: {} for d in self.days}
        self.week_hours: Dict[date, Counter] = defaultdict(Counter)
        self.mon_pn = Counter()
        self.role_counts = {r: Counter() for r in ["PN", "AN", "W", "BU"]}

        self.hierarchy = sort_by_hierarchy(inp.fsas, inp.sales_volume, inp.seniority_order)
        self.targets = self.compute_targets()

    def _validate_inputs(self):
        if not self.inp.fsas:
            raise ValueError("Inputs.fsas is empty. Provide at least 1 FSA (typically 6).")
        if len(set(self.inp.fsas)) != len(self.inp.fsas):
            raise ValueError("Inputs.fsas contains duplicates. Names must be unique.")
        rot = self.inp.sunday_rotation_order
        if len(set(rot)) != len(rot):
            raise ValueError("sunday_rotation_order contains duplicates.")
        if self.inp.sunday_first_assignee not in rot:
            raise ValueError("sunday_first_assignee must be present in sunday_rotation_order.")
        if set(rot) != set(self.inp.fsas):
            raise ValueError("sunday_rotation_order must contain exactly the same names as fsas.")

    def compute_targets(self) -> Dict[str, Dict[str, int]]:
        fsas = self.inp.fsas
        hier = sort_by_hierarchy(fsas, self.inp.sales_volume, self.inp.seniority_order)

        pn_slots = an_slots = w_slots = 0
        for d in self.days:
            req = roles_for_day(d, self.inp, self.push_days)
            pn_slots += 1 if "PN" in req else 0
            an_slots += 1 if "AN" in req else 0
            w_slots += 1 if "W" in req else 0

        def split(total: int) -> Dict[str, int]:
            base = total // len(fsas)
            rem = total % len(fsas)
            out = {p: base for p in fsas}
            for p in hier[:rem]:
                out[p] += 1
            return out

        return {"PN": split(pn_slots), "AN": split(an_slots), "W": split(w_slots)}

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

    def _required_sat_an_person(self, sat: date) -> Optional[str]:
        nxt = sat + timedelta(days=1)
        if nxt.month != self.inp.month or not is_sunday(nxt):
            return None
        return self.sched.get(nxt, {}).get("PN")

    def can_assign(self, p: str, d: date, role: str) -> Tuple[bool, str]:
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
        if self.consecutive_streak_if_work(p, d) > inp.max_consecutive_days:
            return False, "max_consecutive_days"

        if role == "AN" and is_saturday(d):
            req = roles_for_day(d, inp, self.push_days)
            if "AN" in req:
                must_be = self._required_sat_an_person(d)
                if must_be is not None and p != must_be:
                    return False, "sat_an_must_match_sunday_pn"

        wk = week_start_sun(d)
        waive = week_contains_push(wk, self.push_days)
        add_h = hours_for(d, role, inp)
        if (not waive) and (self.week_hours[wk][p] + add_h > inp.weekly_cap_hours):
            return False, "weekly_cap"

        return True, "ok"

    def assign(self, d: date, role: str, p: str):
        self.sched[d][role] = p
        wk = week_start_sun(d)
        self.week_hours[wk][p] += hours_for(d, role, self.inp)
        if role == "PN" and d.weekday() == 0:
            self.mon_pn[p] += 1
        if role in ["PN", "AN", "W"]:
            self.role_counts[role][p] += 1
        else:
            self.role_counts["BU"][p] += 1

    def unassign(self, d: date, role: str):
        p = self.sched[d].pop(role, None)
        if p is None:
            return
        wk = week_start_sun(d)
        self.week_hours[wk][p] -= hours_for(d, role, self.inp)
        if role == "PN" and d.weekday() == 0:
            self.mon_pn[p] -= 1
        if role in ["PN", "AN", "W"]:
            self.role_counts[role][p] -= 1
        else:
            self.role_counts["BU"][p] -= 1

    def set_sundays_by_rotation(self):
        order = self.inp.sunday_rotation_order
        n = len(order)
        idx = order.index(self.inp.sunday_first_assignee)
        sundays = [d for d in self.days if is_sunday(d) and d not in self.inp.observed_holidays]
        for d in sundays:
            chosen = None
            for k in range(n):
                cand = order[(idx + k) % n]
                ok, _ = self.can_assign(cand, d, "PN")
                if ok:
                    chosen = cand
                    break
            if chosen is None:
                raise RuntimeError(f"Could not assign Sunday PN on {d} (rotation candidates unavailable).")
            self.assign(d, "PN", chosen)
            idx = (order.index(chosen) + 1) % n

    def candidate_count(self, d: date, role: str) -> int:
        return sum(1 for p in self.inp.fsas if self.can_assign(p, d, role)[0])

    def candidate_list(self, d: date, role: str) -> List[str]:
        cands = [p for p in self.inp.fsas if self.can_assign(p, d, role)[0]]
        if not cands:
            return []
        hier_rank = {p: i for i, p in enumerate(self.hierarchy)}
        def need(p: str) -> int:
            if role in self.targets:
                return self.targets[role][p] - self.role_counts[role][p]
            return 0
        cands.sort(key=lambda p: (need(p), -hier_rank.get(p, 999)), reverse=True)
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
            ordered = [r for r in ["PN", "AN", "W"] if r in req] + [r for r in req if r not in ["PN", "AN", "W"]]
            for role in ordered:
                if role in self.sched[d]:
                    continue
                n = self.candidate_count(d, role)
                if n == 0:
                    return (d, role)
                if n < best_len:
                    best_len = n
                    best = (d, role)
        return best

    def solve(self, max_nodes: int = 1_500_000) -> Dict[date, Dict[str, str]]:
        self.set_sundays_by_rotation()
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
                self.assign(d, role, p)
                if backtrack():
                    return True
                self.unassign(d, role)
            return False

        if not backtrack():
            raise RuntimeError("Failed to build a valid schedule under current constraints (search exhausted).")
        return self.sched


def print_schedule(sched: Dict[date, Dict[str, str]], inp: Inputs):
    days = month_days(inp.year, inp.month)
    push_days = compute_push_days(inp)
    for d in days:
        req = roles_for_day(d, inp, push_days)
        if not req:
            print(f"{d} (HOLIDAY/NO STAFF)")
            continue
        a = sched.get(d, {})
        ordered = []
        for r in ["PN", "AN", "W", "BU1", "BU2", "BU3", "BU4"]:
            if r in req:
                ordered.append(f"{r}:{a.get(r,'-')}")
        print(f"{d}  " + "  ".join(ordered))


if __name__ == "__main__":
    FSAS = ["Mark", "Dawn", "Will", "Kyle", "CJ", "Greg"]

    sales = {
        "Mark": 1000000,
        "Dawn": 900000,
        "Will": 850000,
        "Kyle": 800000,
        "CJ": 750000,
        "Greg": 700000,
    }

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

    solver = Solver(inp, seed=11)
    sched = solver.solve()

    print_schedule(sched, inp)

    violations = validate_schedule(sched, inp)
    hard = [v for v in violations if v["severity"] == "HARD"]
    print("\\nVALIDATION")
    print(f"Hard violations: {len(hard)}")
    if hard:
        for v in hard[:80]:
            print(v)
    else:
        print("✅ No hard violations detected.")
