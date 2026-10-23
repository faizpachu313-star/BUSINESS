#!/usr/bin/env python3
"""File each day's records into the business workbooks.

    python daily.py add patients  --file-no TVIP01027300 --name "Ali Khan" --consultation GP --services Lab --bill 250
    python daily.py add radiology --file-no TVIP01027301 --patient "Sara Ali" --consultation "MRI LUMBR"
    python daily.py add undi      --rate 25.10 --name Ali --received 100 --given 2400
    python daily.py add travels   --name Ali --service "Visa renewal" --charged 350
    python daily.py add travels   --json '[{"name": "Ali", "service": "Visa", "charged": 350}, ...]'

``add`` takes ``--date YYYY-MM-DD`` (default: today in the UAE), stores the entries in
the section's data file, then rebuilds that section's workbook so they land on that
day's record:

    PATIENTS   PATIENTS/output/Faiz_Patient_Data_Jan<year>-Dec<year>.xlsx   (month sheet, in date order)
    RADIOLOGY  RADIOLOGY/output/Faiz_Radiology_Data_Jan<year>-Dec<year>.xlsx
    UNDI       UNDI/output/AED-INR Exchange Tracker - <Mon> <year>.xlsx     (the day's own sheet)
    TRAVELS    TRAVELS/output/Faiz_Travels_Jan<year>-Dec<year>.xlsx (+ its PDF copy)

Undi profit (AED received x rate - INR given) and every total are Excel formulas.

    python daily.py show [--date D]                    # a day's records in every section
    python daily.py remove SECTION --date D --number N # N = the entry's number that day (1 = first)
    python daily.py build [SECTION ...]                # rebuild workbooks from the data files
    python daily.py export DIR                         # rows for the dashboard: one CSV per section
"""
import argparse
import csv
import glob
import json
import os
import subprocess
import sys
from datetime import date, datetime
from zoneinfo import ZoneInfo

import ledger

ROOT = os.path.dirname(os.path.abspath(__file__))
TZ = ZoneInfo("Asia/Dubai")

# section -> (folder, fields in column order, required fields, number fields)
FIELDS = {
    "patients": ("PATIENTS", ["file_no", "name", "consultation", "services", "bill"], ["name"], ["bill"]),
    "radiology": ("RADIOLOGY", ["file_no", "patient", "consultation"], ["patient"], []),
    "undi": ("UNDI", ["rate", "name", "aed", "inr"], ["name", "aed"], ["rate", "aed", "inr"]),
    "travels": ("TRAVELS", ["name", "service", "charged"], ["name"], ["charged"]),
}
# what people call a field -> its key
ALIASES = {"received": "aed", "aed_received": "aed", "given": "inr", "inr_given": "inr",
           "file": "file_no", "file_number": "file_no", "exchange_rate": "rate"}
# ``add`` flags; each goes through the same names as a --json key
FLAGS = ("file-no", "name", "patient", "consultation", "services", "service",
         "rate", "received", "given", "bill", "charged")
# dashboard column names, where they differ from the stored key (the dashboard page reads these names)
EXPORT_NAMES = {"undi": {"aed": "aed_received", "inr": "inr_given"}}


def today():
    return datetime.now(TZ).date().isoformat()


def number(value, field):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        n = float(str(value).replace(",", "").strip())
    except ValueError:
        sys.exit(f"{field} must be a number, got {value!r}")
    return int(n) if n.is_integer() else n


def read_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def write_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1)
        fh.write("\n")


# ---- where each section keeps its entries ----------------------------------

def ledger_path(section, year):
    return os.path.join(ROOT, FIELDS[section][0], "data", f"{year}.json")


def undi_path(year, month):
    return os.path.join(ROOT, "UNDI", "data", f"{year}-{month:02d}.json")


def day_entries(section, day):
    """The day's entries, in the order they were added."""
    d = date.fromisoformat(day)
    if section == "undi":
        data = read_json(undi_path(d.year, d.month), {})
        return data.get("days", {}).get(str(d.day), {}).get("transactions", [])
    return [e for e in read_json(ledger_path(section, d.year), {}).get("entries", []) if e["date"] == day]


def store(section, day, entries):
    d = date.fromisoformat(day)
    if section == "undi":
        path = undi_path(d.year, d.month)
        data = read_json(path, {"month": d.month, "year": d.year, "days": {}})
        rec = data["days"].setdefault(str(d.day), {"rate": None, "transactions": []})
        for e in entries:
            if e.get("rate") is None:
                e["rate"] = rec.get("rate")
            if e["rate"] is None:
                sys.exit(f"undi: no exchange rate for {day} yet - pass --rate")
            if rec.get("rate") is None:
                rec["rate"] = e["rate"]            # the day's first rate is the sheet's rate for the day
            rec["transactions"].append(e)
        write_json(path, data)
    else:
        path = ledger_path(section, d.year)
        data = read_json(path, {"year": d.year, "entries": []})
        data["entries"].extend(dict(date=day, **e) for e in entries)
        write_json(path, data)
    return d


def remove(section, day, n):
    d = date.fromisoformat(day)
    if section == "undi":
        path = undi_path(d.year, d.month)
        data = read_json(path, {})
        txns = data.get("days", {}).get(str(d.day), {}).get("transactions", [])
        if not 1 <= n <= len(txns):
            sys.exit(f"undi has {len(txns)} entries on {day}")
        gone = txns.pop(n - 1)
    else:
        path = ledger_path(section, d.year)
        data = read_json(path, {})
        idx = [i for i, e in enumerate(data.get("entries", [])) if e["date"] == day]
        if not 1 <= n <= len(idx):
            sys.exit(f"{section} has {len(idx)} entries on {day}")
        gone = data["entries"].pop(idx[n - 1])
    write_json(path, data)
    return gone


# ---- workbooks --------------------------------------------------------------

def build(section, year, month=None):
    folder = os.path.join(ROOT, FIELDS[section][0])
    if section == "undi":
        months = [month] if month else sorted(
            int(os.path.basename(p)[5:7]) for p in glob.glob(os.path.join(folder, "data", f"{year}-??.json")))
        for m in months:
            subprocess.run([sys.executable, os.path.join(folder, "build_workbook.py"),
                            "--data", undi_path(year, m)], check=True)
        return
    spec = ledger.SECTIONS[section]
    out = os.path.join(folder, "output", spec["file"].format(year=year))
    ledger.build(spec, year, read_json(ledger_path(section, year), {}).get("entries", []), out)
    print(out)
    if section == "travels":                       # the folder keeps a PDF copy next to the workbook
        pdf = ledger.libreoffice(out, os.path.dirname(out), "pdf")
        print(pdf or "warning: LibreOffice not found - PDF copy not refreshed")


def years_with_data(section):
    pattern = "????-??.json" if section == "undi" else "????.json"
    found = glob.glob(os.path.join(ROOT, FIELDS[section][0], "data", pattern))
    return sorted({int(os.path.basename(p)[:4]) for p in found})


# ---- commands ---------------------------------------------------------------

def entries_from_args(section, args):
    _, fields, required, numeric = FIELDS[section]
    if args.json:
        raw = json.loads(args.json)
        raw = raw if isinstance(raw, list) else [raw]
    else:
        raw = [{f: v for f in FLAGS if (v := getattr(args, f.replace("-", "_"))) is not None}]

    def key(k):
        k = k.replace("-", "_").replace(" ", "_").lower()
        k = ALIASES.get(k, k)
        if k not in fields and k + "s" in fields:            # service / services, either way round
            k += "s"
        elif k not in fields and k[:-1] in fields:
            k = k[:-1]
        elif k in ("name", "patient") and k not in fields:   # a patient's name, whichever word is used
            k = "patient" if k == "name" else "name"
        return k

    entries = []
    for item in raw:
        item = {key(k): v for k, v in item.items()}
        unknown = set(item) - set(fields) - {"date"}
        if unknown:
            sys.exit(f"{section}: unknown field(s) {', '.join(sorted(unknown))}; fields are {', '.join(fields)}")
        e = {f: item.get(f) for f in fields}
        for f in numeric:
            e[f] = number(e[f], f)
        for f in fields:
            if isinstance(e[f], str):
                e[f] = e[f].strip() or None
        missing = [f for f in required if e[f] in (None, "")]
        if missing:
            sys.exit(f"{section}: missing {', '.join(missing)}")
        entries.append((item.get("date") or args.date, e))
    return entries


def cmd_add(args):
    entries = entries_from_args(args.section, args)
    days = sorted({d for d, _ in entries})
    for day in days:
        date.fromisoformat(day)                     # reject a malformed date before writing anything
    touched = set()
    for day in days:                                # one write per day, so a day is stored whole or not at all
        d = store(args.section, day, [e for d_, e in entries if d_ == day])
        touched.add((d.year, d.month if args.section == "undi" else None))
    for year, month in sorted(touched):
        build(args.section, year, month)
    for day in days:
        show(day, [args.section])


def show(day, sections):
    print(f"\n{date.fromisoformat(day):%A %d %b %Y}")
    for section in sections:
        rows = day_entries(section, day)
        fields = FIELDS[section][1]
        print(f"  {section.upper()} ({len(rows)})")
        for i, e in enumerate(rows, start=1):
            print(f"    {i}. " + " | ".join("" if e.get(f) is None else str(e.get(f)) for f in fields))


def cmd_export(args):
    """Every entry as dashboard rows: id (date#n), date, then the section's fields.

    CSV rather than JSON: a CSV's header names the columns even before the first entry,
    and the dashboard reads a file with no rows only when it knows its columns.
    """
    os.makedirs(args.dir, exist_ok=True)
    for section, (_, fields, _, _) in FIELDS.items():
        names = EXPORT_NAMES.get(section, {})
        rows = []
        for year in years_with_data(section):
            if section == "undi":
                for path in sorted(glob.glob(os.path.join(ROOT, "UNDI", "data", f"{year}-??.json"))):
                    data = read_json(path, {})
                    for dnum, rec in sorted(data.get("days", {}).items(), key=lambda kv: int(kv[0])):
                        day = date(data["year"], data["month"], int(dnum)).isoformat()
                        rows += [(day, t) for t in rec.get("transactions", [])]
            else:
                entries = read_json(ledger_path(section, year), {}).get("entries", [])
                rows += [(e["date"], e) for e in sorted(entries, key=lambda e: e["date"])]
        columns = ["id", "date"] + [names.get(f, f) for f in fields]
        path = os.path.join(args.dir, f"{section}.csv")
        seen = {}
        with open(path, "w", encoding="utf-8", newline="") as fh:
            out = csv.writer(fh)
            out.writerow(columns)
            for day, e in rows:
                seen[day] = seen.get(day, 0) + 1
                out.writerow([f"{day}#{seen[day]}", day] + ["" if e.get(f) is None else e.get(f) for f in fields])
        print(path, len(rows), "rows")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    add = sub.add_parser("add", help="add entries and rebuild the workbook")
    add.add_argument("section", choices=FIELDS)
    add.add_argument("--date", default=today(), help="YYYY-MM-DD (default: today, UAE time)")
    add.add_argument("--json", help="one entry or a list of entries as JSON (keys as the flags below)")
    for flag in FLAGS:
        add.add_argument(f"--{flag}")

    sh = sub.add_parser("show", help="a day's records")
    sh.add_argument("--date", default=today())

    rm = sub.add_parser("remove", help="delete one entry, then rebuild")
    rm.add_argument("section", choices=FIELDS)
    rm.add_argument("--date", required=True)
    rm.add_argument("--number", type=int, required=True, help="the entry's number on that day (1 = first)")

    bd = sub.add_parser("build", help="rebuild workbooks from the data files")
    bd.add_argument("sections", nargs="*", metavar="SECTION", help=", ".join(FIELDS) + " (default: all)")

    ex = sub.add_parser("export", help="write dashboard rows")
    ex.add_argument("dir")

    args = ap.parse_args()
    if args.cmd == "add":
        cmd_add(args)
    elif args.cmd == "show":
        show(args.date, list(FIELDS))
    elif args.cmd == "remove":
        gone = remove(args.section, args.date, args.number)
        print("removed:", gone)
        d = date.fromisoformat(args.date)
        build(args.section, d.year, d.month if args.section == "undi" else None)
        show(args.date, [args.section])
    elif args.cmd == "build":
        unknown = set(args.sections) - set(FIELDS)
        if unknown:
            ap.error(f"unknown section(s): {', '.join(sorted(unknown))}")
        for section in args.sections or FIELDS:
            for year in years_with_data(section) or [datetime.now(TZ).year]:
                build(section, year)
    elif args.cmd == "export":
        cmd_export(args)


if __name__ == "__main__":
    main()
