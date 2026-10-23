# TRAVELS

`output/Faiz_Travels_Jan2026-Dec2026.xlsx` (plus a PDF copy) in the Faiz Medical Records house style:
Times New Roman, navy title band, slate header row, striped rows, teal divider, totals at the bottom.

* **Summary** sheet first: each month (linked to its sheet), total entries, total charged, and a grand total for the year.
* **One sheet per month** (`Jan-2026` … `Dec-2026`) with columns **Sl No, Date, Name, Service, Charged**
  and 50 blank rows. At the bottom: `TOTAL ENTRIES OF <MONTH>` and `TOTAL CHARGED OF <MONTH>` (the month's sum).
* Fill in Date (as text, `01 Jan 2026`), Name, Service and Charged. Sl No, the totals and the Summary are formulas.

Rebuild for another year or more rows:

    python build_workbook.py --year 2027 --rows 80
