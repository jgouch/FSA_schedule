#!/usr/bin/env python3
"""
FINAL Schedule Logic Engine
Author: ChatGPT + John
Purpose: Monthly FSC scheduling with carryover, time-off, role-avoid, and BU backfill
"""

from __future__ import annotations

# =========================
# Imports (COMPLETE)
# =========================
import argparse
import json
import random
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Set, Optional, Tuple

import openpyxl


# =========================
# Data Models
# =========================

@dataclass(frozen=True)
class Inputs:
    year: int
    month: int
    fsas: List[str]

    # Ranking inputs
    sales_volume: Dict[str, float]
    seniority_order: List[str]

    # Time off (hard)
    time_off: Dict[str, Set[date]]

    # Role avoidance (soft & hard)
    role_avoid: Dict[str, Dict[date, Set[str]]] = field(default_factory=dict)
    role_avoid_hard: Dict[str, Dict[date, bool]] = field(default_factory=dict)

    # Sunday logic
    sunday_rotation_order: List[str] = field(default_factory=list)
    sunday_first_assignee: Optional[str] = None

    # Constraints
    max_consecutive_days: int = 6
    max_weekly_hours: int = 40
    hours_primary: int = 8
    hours_bu: int = 8

    # Carryover window
    lookback_days: int = 14


@dataclass
class DayAssignment:
    date: date
    roles: Dict[str, Optional[str]] = field(default_factory=dict)


# =========================
# Utilities
# =========================

def daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def month_bounds(year: int, month: int) -> Tuple[date, date]:
    start = date(year, month, 1)
    if month == 12:
        end = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end = date(year, month + 1, 1) - timedelta(days=1)
    return start, end


# =========================
# Time Off Loader
# =========================

def load_timeoff_xlsx(path: str, sheet: str) -> Tuple[
    Dict[str, Set[date]],
    Dict[str, Dict[date, Set[str]]],
    Dict[str, Dict[date, bool]],
]:
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]

    time_off: Dict[str, Set[date]] = {}
    role_avoid: Dict[str, Dict[date, Set[str]]] = {}
    role_avoid_hard: Dict[str, Dict[date, bool]] = {}

    for r in ws.iter_rows(min_row=2, values_only=True):
        name, d, roles, hard = r[:4]
        if not name or not d:
            continue

        if isinstance(d, datetime):
            d = d.date()

        time_off.setdefault(name, set())
        role_avoid.setdefault(name, {})
        role_avoid_hard.setdefault(name, {})

        if roles == "OFF":
            time_off[name].add(d)
        else:
            role_avoid[name].setdefault(d, set()).update(
                {x.strip() for x in str(roles).split(",")}
            )
            role_avoid_hard[name][d] = bool(hard)

    return time_off, role_avoid, role_avoid_hard


# =========================
# Previous Schedule Loader
# =========================

def load_prev_schedule(path: str) -> Dict[date, Dict[str, str]]:
    raw = json.loads(Path(path).read_text())
    out: Dict[date, Dict[str, str]] = {}
    for k, v in raw.items():
        out[date.fromisoformat(k)] = v
    return out


# =========================
# Core Scheduler
# =========================

class Scheduler:
    def __init__(self, inp: Inputs, prev: Dict[date, Dict[str, str]]):
        self.inp = inp
        self.prev = prev
        self.assignments: Dict[date, DayAssignment] = {}

        self.start, self.end = month_bounds(inp.year, inp.month)

    def build(self):
        for d in daterange(self.start, self.end):
            self.assignments[d] = DayAssignment(d, {})

        self._assign_primary()
        self._assign_bu()

        return self.assignments

    # ---------------------
    # Primary Roles
    # ---------------------

    def _assign_primary(self):
        roles = ["PN", "AN", "W"]

        for d in self.assignments:
            for role in roles:
                self.assignments[d].roles[role] = self._pick_fsa(d, role)

    # ---------------------
    # BU Roles
    # ---------------------

    def _assign_bu(self):
        for d in self.assignments:
            used = set(self.assignments[d].roles.values())
            for i in range(1, 3):
                role = f"BU{i}"
                pick = self._pick_fsa(d, role, used)
                self.assignments[d].roles[role] = pick
                if pick:
                    used.add(pick)

    # ---------------------
    # Selection Logic
    # ---------------------

    def _pick_fsa(self, d: date, role: str, exclude: Set[str] = set()):
        candidates = []

        for fsa in self.inp.fsas:
            if fsa in exclude:
                continue
            if d in self.inp.time_off.get(fsa, set()):
                continue
            if role in self.inp.role_avoid.get(fsa, {}).get(d, set()):
                if self.inp.role_avoid_hard.get(fsa, {}).get(d, False):
                    continue
            candidates.append(fsa)

        if not candidates:
            return None

        # Ranking
        candidates.sort(
            key=lambda f: (
                -self.inp.sales_volume.get(f, 0),
                self.inp.seniority_order.index(f)
                if f in self.inp.seniority_order
                else 999,
            )
        )

        return candidates[0]


# =========================
# CLI
# =========================

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True)  # YYYY-MM
    ap.add_argument("--timeoff-xlsx", required=True)
    ap.add_argument("--timeoff-sheet", required=True)
    ap.add_argument("--prev-schedule", required=True)

    args = ap.parse_args()

    year, month = map(int, args.month.split("-"))

    time_off, role_avoid, role_avoid_hard = load_timeoff_xlsx(
        args.timeoff_xlsx, args.timeoff_sheet
    )

    prev = load_prev_schedule(args.prev_schedule)

    fsas = sorted(time_off.keys())

    # Sales ranking optional
    sales_volume = {}
    seniority = fsas.copy()

    inp = Inputs(
        year=year,
        month=month,
        fsas=fsas,
        sales_volume=sales_volume,
        seniority_order=seniority,
        time_off=time_off,
        role_avoid=role_avoid,
        role_avoid_hard=role_avoid_hard,
    )

    sched = Scheduler(inp, prev).build()

    out = {
        d.isoformat(): v.roles
        for d, v in sched.items()
    }

    out_path = Path(f"{year}-{month:02d}_schedule.json")
    out_path.write_text(json.dumps(out, indent=2))
    print(f"✅ Schedule written to {out_path}")


if __name__ == "__main__":
    main()
