# Assam Pathology Project – Daily Dashboard

A daily performance dashboard for the Assam Pathology PPP project (Krsnaa Diagnostics Ltd, NHM Assam).
Upload the daily Excel file and the dashboard updates automatically.

## What it tracks

For the selected **day** (with % change vs the previous day) and for the selected **period** (month to date, last 7/30/90 days, all, or custom):

| Metric | Formula |
|---|---|
| Patients | Unique Patient IDs per day (line list) or the patient count column (summary file) |
| Tests | Number of test rows (or Qty column) |
| Revenue | Sum of the Amount / Rate column |
| Avg revenue per patient | Revenue ÷ Patients |
| Avg revenue per test | Revenue ÷ Tests |
| Avg tests per patient | Tests ÷ Patients |

Also included: daily trend charts, district-wise, facility-wise and top-tests tables, a daily summary table (export to Excel) and an upload history.

## Daily workflow

1. Open the dashboard and click **Upload Excel** (or drag the file onto the page).
2. Check the columns and totals shown in the **Check & import** window, then click **Import**.
3. The dashboard jumps to the uploaded date.

Notes:

- **Re-uploading a date replaces it** – no double counting if the same file is uploaded twice or a corrected file is sent.
- Column names are detected automatically (e.g. *Reg No, UHID, Barcode, Test Name, Investigation, Amount, Rate, Hospital, District*). If something is wrong, choose the correct column once – it is remembered for the next upload of the same format.
- Title rows above the header and *Total / Grand Total* rows are ignored.
- If the file has no date column, the date is taken from the file name (e.g. `Report 24-10-2026.xlsx`) or can be selected.
- Dates are read in Indian format (DD-MM-YYYY).
- A wrong day can be deleted from the **Daily summary** table (✕).
- **More → Download Excel template** gives both supported formats:
  - *Line list* – one row per test: Date, Patient ID, District, Facility, Test Name, Amount
  - *Daily summary* – Date, District, Facility, No of Patients, No of Tests, Revenue

## Where the data is stored

Data is stored **in the browser on the computer where it is uploaded** (nothing is sent to any server).
So:

- Use the same computer and browser for daily uploads.
- Take a weekly **More → Backup data** file. Use **Restore data from backup** to move the data to another computer or browser.
- Clearing browser data/history for this site deletes the dashboard data – restore from the backup if this happens.

## Running it

No installation needed – it is a static website.

- **Locally:** open `index.html` in Chrome or Edge.
- **Online (free):** GitHub repo → *Settings → Pages* → Source: *Deploy from a branch*, select the branch and `/ (root)`. The dashboard will be available at the GitHub Pages link.

## Files

```
index.html          Page layout
css/styles.css      Styling (light and dark mode)
js/app.js           Excel reading, calculations, charts and tables
js/vendor/          SheetJS (Excel reader) and Chart.js – bundled so it works offline
```
