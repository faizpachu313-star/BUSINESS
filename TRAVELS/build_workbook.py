#!/usr/bin/env python3
"""Build the Faiz Travels monthly workbook in the Faiz Medical Records house style.

A "Summary" sheet followed by one sheet per month (``Jan-2026`` ...). Each month
sheet has Sl No / Date / Name / Service / Charged, the month's entries from
data/<year>.json, blank rows ready to fill, and totals at the bottom (number of
entries and the month's total Charged). Sl No, totals and the Summary are live
formulas. The layout lives in ../ledger.py, shared with PATIENTS and RADIOLOGY.

Usage:
    python build_workbook.py                 # 2026
    python build_workbook.py --year 2027 [--rows 60] [--out FILE.xlsx]

Daily entries are added with ``python daily.py add travels ...`` from the repo root.
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import ledger  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--rows", type=int, default=50, help="blank rows per month sheet")
    ap.add_argument("--out")
    args = ap.parse_args()
    spec = ledger.SECTIONS["travels"]
    out = args.out or os.path.join(HERE, "output", spec["file"].format(year=args.year))
    data = os.path.join(HERE, "data", f"{args.year}.json")
    entries = []
    if os.path.exists(data):
        with open(data, encoding="utf-8") as fh:
            entries = json.load(fh).get("entries", [])
    print(ledger.build(spec, args.year, entries, out, args.rows))


if __name__ == "__main__":
    main()
