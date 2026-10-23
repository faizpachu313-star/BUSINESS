# BUSINESS — daily records

The owner sends the day's records in chat. File every entry with `daily.py`, which stores it under
the section's `data/` folder and rebuilds that section's Excel workbook, so it lands on that day's record:

| Section | Owner's columns | Command |
|---|---|---|
| Patient data | File no., Name, Consultation, Services, Bill | `python daily.py add patients --file-no … --name … --consultation … --services … --bill …` |
| Radiology | File no., Patient, Consultation | `python daily.py add radiology --file-no … --patient … --consultation …` |
| Undi | Exchange rate, Name, Received (AED), Given (INR) | `python daily.py add undi --rate … --name … --received … --given …` |
| Travels | Name, Services, Charged | `python daily.py add travels --name … --service … --charged …` |

* Several entries at once: `--json '[{…}, {…}]'` with the same keys (`file_no`, `received`, …).
* `--date YYYY-MM-DD` for any day but today (today = UAE time). Undi profit and all totals are formulas;
  never type them in. Undi `--given` can be left out until the INR is handed over.
* Undi `--rate` can be left out after the day's first entry (the day's rate is reused).
* Check with `python daily.py show [--date D]`; fix a mistake with
  `python daily.py remove SECTION --date D --number N`, then add the corrected entry.
* If anything is unclear (which section, an amount, the date), ask before filing.

## After filing

1. `python daily.py export <scratch dir>` writes `patients.csv`, `radiology.csv`, `undi.csv`,
   `travels.csv`.
2. Refresh the dashboard https://claude.ai/artifact/U3PwzpXXUvBbEEikf2xNAs (Dashboard type): upload each
   file as an asset (Artifact publish, `asset: true`), then `update` `datasets/<section>` in its store
   with the new `source.url`, `source.name` and `updated: {at, by: "Claude"}`.
3. Commit the `data/` and `output/` changes and push.

## Privacy

This repository is public. Patient and radiology entries hold patients' names, file numbers and
treatment: do not commit them while the repository is public. Ask the owner to make it private first
(GitHub → Settings → General → Danger Zone → Change visibility).
