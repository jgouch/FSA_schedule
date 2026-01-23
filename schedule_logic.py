#!/usr/bin/env python3
"""
schedule_logic.py
Logic-first scheduler for Harpeth Hills FSAs (NO Excel formatting yet).

Inputs each run:
1) time_off requests (hard): dates each FSA cannot work
2) sales volumes: Total Volume = Cemetery + PAF (used for PN/AN remainder distribution)
3) Sunday rotation: fixed order + who is first Sunday assignee in this month

Hard rules implemented:
- Observed holidays: no roles at all
- Roles per day:
  * Sun: PN only (5 hours)
  * Mon: PN, AN, W, BU1, BU2, BU3 (all 6)
  * Tue–Fri (normal): PN, AN, W, BU1, BU2 (5 people)
  * Push week (last Mon–Fri): PN, AN, W, BU1, BU2, BU3 (all 6)
  * Sat (normal): PN, AN
  * Push-week Sat (last Sat): PN, AN, BU1, BU2, BU3, BU4 (all 6; no W)
- Sunday PN rotation, and Sunday PN must be AN on Saturday prior
- No person scheduled more than once per day
- No one holds the SAME specific role on back-to-back days (BU1->BU2 allowed, BU1->BU1 not allowed)
- Monday PN max 1 per month
- Max 5 consecutive work days (within month)
- Weekly hours cap (Sun–Sat) <= 40 OUTSIDE push week (push week relaxes cap)

Output:
- Prints day-by-day assignments
- Prints validation (0 HARD violations = schedule complies with hard rules)

Next phase (after you approve logic):
- Add pay period caps (80) with cross-month carryover
- Add “preferences” (e.g., Dawn no PN/AN on Feb 2) as SOFT constraints
- Add Excel template rendering
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
import calendar
import random
from collections import defaultdict, Counter
from typing import Dict, List, Optional, Set, Iterable, Any


# -----------------------------
# Inputs
# -----------------------------

@dataclass(frozen=True)
class Inputs:
    year: int
    month: int
    fsas: List[str]  # must be 6 FSAs

    # Used for PN/AN remainder distribution
    sales_volume: Dict[str, float]
    # Tie-break for sales ties (earlier = more senior)
    seniority_order: List[str]

    # Hard day-off requests
    time_off: Dict[str, Set[date]]

    # Sunday rotation
    sunday_rotation_order: List[str]  # length 6
    sunday_first_assignee: str        # who works PN on first Sunday of this month

    # Observed holidays (no roles)
    observed_holidays: Set[date]

    # Push week settings
    push_week_enabled: bool = True

    # Shift hours
    hours_per_day: int = 8
    sunday_hours: int = 5

    # Hard constraints
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
    # Sunday-start week
    return d - timedelta(days=(d.weekday() + 1) % 7)

def sort_by_hierarchy(fsas: List[str], sales: Dict[str, float], seniority: List[str]) -> List[str]:
    seniority_rank = {name: i for i, name in enumerate(seniority)}
    # higher sales first; tie => more senior first
    return sorted(fsas, key=lambda n: (-sales.get(n, 0.0), seniority_rank.get(n, 10_000)))

def last_weekdays_block(year: int, month: int) -> List[date]:
    """Return the last Mon–Fri dates of the month (the final workweek ending in the month)."""
    days = month_days(year, month)
    last_day = days[-1]
    last_fri = last_day
    while last_fri.weekday() != 4:
        last_fri -= timedelta(days=1)
    last_mon = last_fri - timedelta(days=4)
    return [last_mon + timedelta(days=i) for i in range(5)]


# -----------------------------
# Role model
# -----------------------------

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

    # Mon–Fri
    if d.weekday() == 0:  # Monday
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


# -----------------------------
# Validation
# -----------------------------

def validate_schedule(sched: Dict[date, Dict[str, str]], inp: Inputs) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    days = month_days(inp.year, inp.month)

    push_days = set()
    if inp.push_week_enabled:
        push_days.update(last_weekdays_block(inp.year, inp.month))
        push_days.add(max(d for d in days if is_saturday(d)))

    def add(sev: str, msg: str, d: Optional[date] = None, who: Optional[str] = None, role: Optional[str] = None):
        issues.append({"severity": sev, "date": d, "person": who, "role": role, "message": msg})

    # Required roles
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

    # One role per person per day
    for d in days:
        vals = list(sched.get(d, {}).values())
        for p, c in Counter(vals).items():
            if c > 1:
                add("HARD", "Person scheduled multiple times on same day.", d=d, who=p)

    # Time off
    for d in days:
        for role, p in sched.get(d, {}).items():
            if d in inp.time_off.get(p, set()):
                add("HARD", "Scheduled on a requested day off.", d=d, who=p, role=role)

    # Same role back-to-back (within month)
    for i in range(1, len(days)):
        d, prev = days[i], days[i - 1]
        for role in set(sched.get(d, {})).intersection(sched.get(prev, {})):
            if sched[prev].get(role) == sched[d].get(role):
                add("HARD", "Same role assigned back-to-back days.", d=d, who=sched[d][role], role=role)

    # Monday PN max
    mon_pn = Counter()
    for d in days:
        if d.weekday() == 0 and sched.get(d, {}).get("PN"):
            mon_pn[sched[d]["PN"]] += 1
    for p, c in mon_pn.items():
        if c > inp.monday_pn_max_per_month:
            add("HARD", f"Monday PN exceeds max ({c}).", who=p)

    # Sunday PN must be Saturday AN
    for d in days:
        if is_sunday(d) and sched.get(d, {}).get("PN"):
            p = sched[d]["PN"]
            sat = d - timedelta(days=1)
            if sat.month == inp.month and "AN" in roles_for_day(sat, inp, push_days):
                if sched.get(sat, {}).get("AN") != p:
                    add("HARD", "Sunday PN must be AN on prior Saturday.", d=d, who=p)

    # Max consecutive days (within month)
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

    # Weekly cap (Sun–Sat) outside push week
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


# -----------------------------
# Scheduler
# -----------------------------

class Scheduler:
    def __init__(self, inp: Inputs, seed: int = 7):
        self.inp = inp
        self.rng = random.Random(seed)
        self.days = month_days(inp.year, inp.month)

        self.push_days = set()
        if inp.push_week_enabled:
            self.push_days.update(last_weekdays_block(inp.year, inp.month))
            self.push_days.add(max(d for d in self.days if is_saturday(d)))

        self.reset()

    def reset(self):
        self.sched: Dict[date, Dict[str, str]] = {d: {} for d in self.days}
        self.week_hours: Dict[date, Counter] = defaultdict(Counter)
        self.role_counts = {r: Counter() for r in ["PN", "AN", "W", "BU"]}
        self.mon_pn = Counter()

    def is_available(self, person: str, d: date) -> bool:
        return d not in self.inp.time_off.get(person, set())

    def can_assign(self, person: str, d: date, role: str) -> tuple[bool, str]:
        inp = self.inp

        if not self.is_available(person, d):
            return False, "time_off"

        if person in self.sched[d].values():
            return False, "already_scheduled_today"

        prev = d - timedelta(days=1)
        if prev.month == d.month and self.sched.get(prev, {}).get(role) == person:
            return False, "same_role_back_to_back"

        if role == "PN" and d.weekday() == 0 and self.mon_pn[person] >= inp.monday_pn_max_per_month:
            return False, "monday_pn_limit"

        # consecutive day streak (within month)
        streak = 1
        cur = d - timedelta(days=1)
        while cur.month == d.month:
            if person in self.sched.get(cur, {}).values():
                streak += 1
                cur -= timedelta(days=1)
            else:
                break
        if streak > inp.max_consecutive_days:
            return False, "max_consecutive_days"

        # weekly cap outside push week
        wk = week_start_sun(d)
        add_h = hours_for(d, role, inp)
        if d not in self.push_days:
            if self.week_hours[wk][person] + add_h > inp.weekly_cap_hours:
                return False, "weekly_cap"

        return True, "ok"

    def assign(self, d: date, role: str, person: str):
        self.sched[d][role] = person
        h = hours_for(d, role, self.inp)
        wk = week_start_sun(d)
        self.week_hours[wk][person] += h

        if role in ["PN", "AN", "W"]:
            self.role_counts[role][person] += 1
        else:
            self.role_counts["BU"][person] += 1

        if role == "PN" and d.weekday() == 0:
            self.mon_pn[person] += 1

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

    def set_sundays(self):
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
                raise RuntimeError(f"Could not assign Sunday PN on {d} due to time off/constraints.")
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
                if "AN" not in roles_for_day(sat, self.inp, self.push_days):
                    continue
                ok, reason = self.can_assign(p, sat, "AN")
                if not ok:
                    raise RuntimeError(f"Cannot enforce Saturday AN for Sunday PN: {p} on {sat} ({reason}).")
                # If AN already filled (e.g., earlier forced), it should match p; otherwise assign.
                if self.sched[sat].get("AN") and self.sched[sat]["AN"] != p:
                    raise RuntimeError(f"Saturday AN conflict on {sat}: needs {p}, has {self.sched[sat]['AN']}.")
                if "AN" not in self.sched[sat]:
                    self.assign(sat, "AN", p)

    def pick_candidate(self, d: date, role: str, targets: Dict[str, Dict[str, int]]) -> Optional[str]:
        fsas = self.inp.fsas
        cands = []
        for p in fsas:
            ok, _ = self.can_assign(p, d, role)
            if ok:
                cands.append(p)
        if not cands:
            return None

        def score(p: str) -> float:
            s = 0.0

            # fairness pressure for PN/AN/W
            if role in targets:
                need = targets[role][p] - self.role_counts[role][p]
                s += 6.0 * need

            # avoid repeated weekday PN (soft)
            if role == "PN":
                wd = d.weekday()
                already = 0
                for dd in self.days:
                    if dd.weekday() == wd and self.sched[dd].get("PN") == p:
                        already += 1
                # Monday already hard-limited; still discourage repeats
                s -= 3.0 * max(0, already - (0 if wd != 0 else 0))

            # lower weekly hours preferred (soft)
            wk = week_start_sun(d)
            s -= 0.25 * self.week_hours[wk][p]

            # small randomness
            s += self.rng.uniform(-0.4, 0.4)
            return s

        cands.sort(key=score, reverse=True)
        topk = min(3, len(cands))
        return self.rng.choice(cands[:topk])

    def fill_day(self, d: date, targets: Dict[str, Dict[str, int]]) -> bool:
        required = roles_for_day(d, self.inp, self.push_days)
        if not required:
            return True

        if is_sunday(d):
            return "PN" in self.sched[d]

        # primary roles
        for r in ["PN", "AN", "W"]:
            if r in required and r not in self.sched[d]:
                cand = self.pick_candidate(d, r, targets)
                if cand is None:
                    return False
                self.assign(d, r, cand)

        # BU roles
        for r in required:
            if r in self.sched[d]:
                continue
            cand = self.pick_candidate(d, "BU", {"BU": {p: 9999 for p in self.inp.fsas}})
            if cand is None:
                return False
            self.assign(d, r, cand)

        return True

    def build(self, attempts: int = 800, seed_start: int = 1) -> Dict[date, Dict[str, str]]:
        targets = self.compute_targets()

        best = None
        best_soft = float("-inf")

        for a in range(attempts):
            self.rng.seed(seed_start + a)  # vary randomness
            self.reset()

            try:
                self.set_sundays()
                self.enforce_sat_an_for_sun()

                ok = True
                for d in self.days:
                    if not self.fill_day(d, targets):
                        ok = False
                        break
                if not ok:
                    continue

                violations = validate_schedule(self.sched, self.inp)
                if any(v["severity"] == "HARD" for v in violations):
                    continue

                soft = self._soft_score(targets)
                if soft > best_soft:
                    best_soft = soft
                    best = {d: dict(self.sched[d]) for d in self.days}

            except Exception:
                continue

        if best is None:
            raise RuntimeError("Failed to build a valid schedule under current constraints.")
        self.sched = best
        return best

    def _soft_score(self, targets: Dict[str, Dict[str, int]]) -> float:
        # penalize deviation from targets; penalize duplicate day groups
        score = 0.0
        for r in ["PN", "AN"]:
            for p in self.inp.fsas:
                score -= 25.0 * abs(self.role_counts[r][p] - targets[r][p])
        for p in self.inp.fsas:
            score -= 8.0 * abs(self.role_counts["W"][p] - targets["W"][p])

        groups = Counter()
        for d in self.days:
            if roles_for_day(d, self.inp, self.push_days):
                groups[frozenset(self.sched[d].values())] += 1
        for _, c in groups.items():
            score -= 1.0 * max(0, c - 1)
        return score


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


# -----------------------------
# Example run (edit per month)
# -----------------------------
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

    # Observed holidays (no roles) for THIS month only
    holidays: Set[date] = set()  # none in Feb 2026

    # Hard time off requests
    toff = {p: set() for p in FSAS}
    # Will: Feb 7–12
    for dd in daterange(date(2026, 2, 7), date(2026, 2, 12)):
        toff["Will"].add(dd)
    # CJ: Feb 5–6
    toff["CJ"].update({date(2026, 2, 5), date(2026, 2, 6)})
    # Dawn: Feb 25
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

    sch = Scheduler(inp).build(attempts=1200, seed_start=11)
    print_schedule(sch, inp)

    violations = validate_schedule(sch, inp)
    hard = [v for v in violations if v["severity"] == "HARD"]
    print("\nVALIDATION")
    print(f"Hard violations: {len(hard)}")
    if hard:
        for v in hard[:50]:
            print(v)
    else:
        print("✅ No hard violations detected.")
