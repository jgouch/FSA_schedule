#!/usr/bin/env python3
import argparse
import json
import re
from datetime import date
from pathlib import Path
from typing import Dict, Tuple, List
import openpyxl

def parse_month_year(s: str) -> Tuple[int, int]:
    s = str(s).strip()
    m = re.search(r"([A-Za-z]+)\s+(\d{4})", s)
    if not m:
        raise ValueError(f"Cannot parse month/year from '{s}'")
    month_name = m.group(1).lower()
    year = int(m.group(2))
    months = {
        "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
        "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12
    }
    return year, months[month_name]

def export_template(path_in: str):
    wb = openpyxl.load_workbook(path_in, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header = ws.cell(1, 2).value
    year, month = parse_month_year(header)

    date_cells = []
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, int) and 1 <= cell.value <= 31:
                date_cells.append((cell.row, cell.column, cell.value))

    schedule = {}
    for r, c, daynum in date_cells:
        d = date(year, month, daynum).isoformat()
        label_col = c + 1
        name_col = c + 2
        roles = {}

        last_primary_row = r
        for rr in range(r + 1, r + 8):
            label = ws.cell(rr, label_col).value
            name = ws.cell(rr, name_col).value
            if not label or not name:
                continue
            label = str(label).strip().upper()
            name = str(name).strip()
            if label.startswith("PN"):
                roles["PN"] = name
                last_primary_row = rr
            elif label.startswith("AN"):
                roles["AN"] = name
                last_primary_row = rr
            elif label.startswith("W"):
                roles["W"] = name
                last_primary_row = rr

        bu_idx = 1
        for rr in range(last_primary_row + 1, r + 12):
            name = ws.cell(rr, name_col).value
            if name:
                roles[f"BU{bu_idx}"] = str(name).strip()
                bu_idx += 1

        if roles:
            schedule[d] = roles

    return year, month, schedule

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True)
    ap.add_argument("--out", dest="out", required=True)
    args = ap.parse_args()

    year, month, sched = export_template(args.inp)
    Path(args.out).write_text(json.dumps(sched, indent=2), encoding="utf-8")
    print(f"✅ Exported {len(sched)} days → {args.out}")

if __name__ == "__main__":
    main()
