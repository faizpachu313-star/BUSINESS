# AED ⇄ INR daily exchange tracker

Generates `output/AED-INR Exchange Tracker - <Mon>.xlsx`: one sheet per day (`01-Jan` …), each with its own
editable rate (`1 AED = ₹ [rate]`), a Name / AED Received / INR Given / Profit-Loss table, daily totals, and a
`Monthly Summary` sheet that consolidates every day with live formulas.

Only inputs are written as values (rate, name, AED, INR); Profit / Loss, totals and the summary are formulas.

    python build_workbook.py                       # reads data/transactions.json
    python build_workbook.py --data other.json --out some.xlsx

`data/transactions.json`:

    {"month": 1, "year": 2027,
     "days": {"1": {"rate": 25.0,
                    "transactions": [{"name": "Ali", "aed": 100, "inr": 2400}]}}}

* `inr: null` = INR not handed over yet (row counts as a transaction, Profit / Loss stays blank until filled in).
* A day with no rate keeps Profit / Loss blank and shows "enter today's rate".
* Each day sheet has 40 formula rows; a busier day is generated with `transactions + 10` rows automatically.
* `compact_xlsx.py` shrinks the file (shared formulas / strings) and embeds cached results (needs LibreOffice;
  without it the file is still correct, Excel just calculates on open).
