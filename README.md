# BUSINESS

| Section | Folder | Contents |
|---|---|---|
| PATIENT DATA | [`PATIENTS/`](PATIENTS/) | Yearly workbook, a sheet per month: Sl No, Date, File No., Name, Consultation, Services, Bill |
| RADIOLOGY | [`RADIOLOGY/`](RADIOLOGY/) | Yearly workbook, a sheet per month: Sl No, Date of Visit, File No., Patient, Consultation |
| UNDI | [`UNDI/`](UNDI/) | AED ⇄ INR daily exchange tracker: a workbook per month, a sheet per day, profit calculated |
| TRAVELS | [`TRAVELS/`](TRAVELS/) | Monthly travels workbook: Sl No, Date, Name, Service, Charged, monthly totals |
| FLAT-FM | [`FLAT-FM/`](FLAT-FM/) | FLAT-FM section (empty for now) |
| FLAT-ML | [`FLAT-ML/`](FLAT-ML/) | FLAT-ML section (empty for now) |

## Daily records

`daily.py` files each day's entries on that day's record and rebuilds the section's workbook:

    python daily.py add patients  --file-no TVIP01027300 --name "Ali Khan" --consultation GP --services Lab --bill 250
    python daily.py add radiology --file-no TVIP01027301 --patient "Sara Ali" --consultation "MRI LUMBR"
    python daily.py add undi      --rate 25.10 --name Ali --received 100 --given 2400
    python daily.py add travels   --name Ali --service "Visa renewal" --charged 350
    python daily.py show                     # today's records in every section

Entries are kept in each section's `data/` folder; the workbooks in `output/` are rebuilt from them
(`python daily.py build`). Sl No, every total and the Undi profit are Excel formulas. `ledger.py` holds
the Faiz Medical Records layout shared by PATIENTS, RADIOLOGY and TRAVELS. Workbooks are passed
through LibreOffice when it is installed, so they carry calculated results for phone and Drive previews.
