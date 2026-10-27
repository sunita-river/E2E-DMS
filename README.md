# E2E-DMS

End-to-end Playwright automation for **River DMS** (`rivermobility.gaindms.com`). It checks reports on the live system and produces Excel outputs and email summaries.

| Script | What it does | Output |
|---|---|---|
| `tests/enquiryListReport.spec.ts` | Dealer login → Enquiry List (full list, every row), then **Product-wise Sales Details** by Invoice date and by Booking date in the same session. Checks each report's columns, blanks and dates, counts dealers, and builds an Enquiry / Invoice / Booking summary by date and dealer (bookings counted once by Booking No) whose totals must match the reports | `Output/Enquiry_List_<from>_<to>.xlsx`, `Output/Enquiry_Summary_<from>_<to>.xlsx` (By Date, By Dealer, By Date and Dealer); the per-day table is pasted into the run email |
| `tests/GSTR Report.spec.ts` | For every dealer in `credentials.json`: GSTR ProductWise Sales Details, plus the **Post GST Discount** read from each Product Sale invoice PDF | `Output/<dealer>_SalesGSTNReport_<from>_<to>.xlsx` and `MasterData_…xlsx`; anything not found in DMS (report gone from the menu, missing columns, bookings not invoiced, discounts not read, dealers not run) goes to `Output/GSTR_Missing_<from>_<to>.xlsx` (highlighted red) and a red "not found in DMS" section in the run email |
| `pdi/PDI.spec.ts` | For every PDI dealer (`PIDDealerCodes` / `PIDUsers`): finds **WORKSHOP → PDI** rows with a blank field and fills Engine / Battery / Charger from the VIN Details sheet in **Settings → Admin → Modify vehicle details** | one copy per day, `Output/PDI_Daily_<date>.xlsx` + dashboard, refreshed by every run (each run's own file goes to `Output/PDI runs/`); no email per run |
| `tests/Automation for Testing.spec.ts` | **Stage DMS only** (`credentials.stageUser`): River User login → create enquiry → follow-up → booking → ₹5,000 receipt, checking each step's fields and saved records. Each run leaves new records on stage | Playwright report only |
| `asm-report/asmwise_report.py` | River User login → Enquiry List and Booking Master Tracker Details for Yesterday, MTD and YTD. Counts enquiries, test rides (Home / Store / not taken), Booked-Active and Retail bookings per Dealer Code and ASM, in total and by lead source (Online, Walk-in, Telephone, Activation), with conversion %s and an All Dealers total. Runs daily from 08:07 IST on GitHub Actions (`.github/workflows/asmwise-report.yml`), which then sends the ASM emails (retried hourly to 11:07 IST and started from Google too, by [resources/apps-script/triggerAsmReport.gs](resources/apps-script/triggerAsmReport.gs), in case GitHub skips a run; sent once a day); `asm-report/run_asmwise_report.bat` does the same locally. Everything for this report lives in `asm-report/` (see its README) | `Output/ASM/ASMwise_AllSources_Summary_<yesterday>.xlsx`; log in `Output/ASM/logs/asmwise_report.log` |
| `asm-report/asm_email.py` | Builds one email per ASM from the latest summary, matched by ASM name (regular expressions) to `asm-report/Mapping/ASM_Name_Mapping.xlsx`. Sent once a day from 08:07 IST by the same workflow, to Santhosh, Narendra Thota, Dinith and Akram with ramkishor@, hemanthkr@ and ashishg@ on Cc. A preview by default; `-- --send` mails them through Gmail and `-- --only <address>` limits it to those ASMs | `Output/ASM/email_preview/<yesterday>/` |
| `mis-check/run_mis_check.js` | Opens every report under every **MIS** page (~274), classifies each one (OK / server error / timeout / blank / needs input), records its definition and available columns, and compares with the previous run | `Output/MIS/` (formatted Excel, CSV, summary, snapshot) and an email summary |

Reports with more than 1,000 rows are read from the DMS's Excel download, so large dealers are fully covered.

---

## Setup

Requires Windows with PowerShell, or any machine with **Node.js 18+**.

```powershell
git clone https://github.com/sunita-river/E2E-DMS.git
cd E2E-DMS
.\setup.bat            # or: npm run setup
```

`setup.ps1` installs Node.js and Google Chrome if they're missing, runs `npm ci`, installs the Playwright browser, and creates the two private config files from their templates:

| File (never committed) | Created from | Fill in |
|---|---|---|
| `resources/credentials.json` | `resources/credentials.example.json` | `dealervalidUser` (Enquiry List, MIS check), `dealerUsers` (one entry per dealer for GSTR) and, for PDI, `PIDLogin` (the shared login) with `PIDDealerCodes` (the dealer codes), plus `PIDUsers` for any dealer with its own login; `stageUser` for the Stage DMS spec; `riverUserValidUser` for the ASM-wise report |
| `.env` | `.env.example` | `GMAIL_USER`, `GMAIL_APP_PASSWORD` (a [Gmail App Password](https://myaccount.google.com/apppasswords)), `GMAIL_TO`; optional `PDI_CC` |
| `resources/VIN_Details.csv` | Fetched automatically before every `npm run test:pdi*` (or by hand: `npm run pdi:vin`) from the **VIN BIN Update (Sunitha)** Google Sheet's copy on the `vin-data` branch, sent by its Apps Script every 15 minutes (see CI/CD). Only this file is used; other `VIN*` files are ignored | VIN sheet for PDI: columns found by their header in row 1 (VIN, Battery No, Motor No, Charger), in any order |

The ASM-wise report is Python (3.10+). Install its packages once with `pip install -r asm-report/requirements.txt`.

---

## Running

```powershell
npm run test:enquiry-list      # Enquiry List + Product-wise Sales Invoice / Booking + summary (~3 min)
npm run test:gstr              # GSTR Report, all dealers (~15-60 min)
npm run test:stage             # Stage DMS: login, enquiry, follow-up, booking, receipt (creates records on stage)
npm run test:pdi:dry           # PDI: fill the form for every blank chassis, don't Save
npm run test:pdi               # PDI: fill and Save (changes live DMS records; ~25 s per chassis)
npm run report:pdi-daily       # PDI: the one email of the day, for all of today's runs (run it after the last PDI run)
npm run report:asmwise         # ASM-wise summary (Python; ~6 min). Add -- --skip-download to rebuild from the last exports
npm run report:asmwise:email   # preview the per-ASM emails for the latest summary (add -- --send to mail them)
npm start                      # MIS report check (~10 min)
npm test                       # every spec in tests/: Enquiry List, GSTR, Stage DMS and the legacy Finance Dashboard
                               # (PDI is skipped unless run with the commands above; Stage needs credentials.stageUser)
npm run report:html            # open the last Playwright HTML report
```

To watch the browser, add `-- --headed` to the Playwright commands (for example `npm run test:gstr -- --headed`). By default the tests run on Google Chrome. To use Microsoft Edge instead:

```powershell
$env:BROWSER_CHANNEL='msedge'; npm run test:gstr
```

After each Playwright run, `tests/reporters/gmail-reporter.ts` emails the pass/fail status to `GMAIL_TO`. The MIS check sends its own HTML summary with the Excel file attached.

### GSTR Report details
- **Blocked since 08-Oct-2026:** the DMS removed *ProductWise Sales Details* from MIS → Accounts Reports → GSTR Reports. The spec stops at once with that message until the report is back or its replacement is known.
- **Date range:** 01-09-2026 to yesterday by default. Override with `GSTR_FROM` / `GSTR_TO` (dd-mm-yyyy).
- **Data checks** (after every file is saved, per dealer): required columns present, every InvoiceDate inside the date range, TaxableAmount and TotalAmount are numbers, TotalAmount = Taxable + CGST + SGST + IGST + CESS (within ₹1), Product Sale rows have an InvoiceNo, ReferenceNo, Chassis No and a numeric DiscountAmount, and MasterData has exactly the dealer files' rows. Any failure lists the dealer and the first few rows.
- **Resumable:** a dealer whose file for the same date range already exists in `Output/` is reused, not re-run. Set `GSTR_RERUN_ALL=1` to redo every dealer.
- **One dealer failing doesn't stop the run.** It's retried once at the end, and MasterData is always written. The test fails (red) only after everything that could be saved is saved.
- **An invoice that can't be read is skipped** with a `[WARN] Reference …` line; check those by hand.

### PDI details
- **Run it** with `npm run test:pdi:dry` (fill only) or `npm run test:pdi` (fill and Save). In PowerShell use `npm.cmd` if scripts are blocked.
- **Excel:** a formatted workbook — a Summary sheet (headline figures and a table by dealer), then one row per chassis with the result colour-coded, then every dealer checked. Header rows are frozen and filterable.
- **Checked:** every vehicle on the PDI page — the page lists 25 by default, so when its "Displaying : N" count is higher the spec sets the "Top" box to N and clicks Show again. If the list says "Displaying : 0" it clicks Show up to `PDI_REFRESHES` more times; if it shows no count and the grid is full at 25, it warns. After Save, each chassis is opened again to confirm the DMS kept the values; a timeout triggers a fresh login and one retry.
- **Speed:** `PDI_PARALLEL` dealers (default 3) run at once, each in its own browser window. After Show / Save the spec waits only until the page has reloaded (up to `PDI_LOAD_TIMEOUT_S`, 30s) instead of a fixed 10s. If the DMS allows only one session per login, set `PDI_PARALLEL=1`.
- **Pass / fail:** the run passes only when every blank chassis is updated and no dealer fails. A failed run is still red. Runs send **no email**: rerun `npm run test:pdi` as often as needed during the day, and each run refreshes the day's single copy (`Output/PDI_Daily_<date>.xlsx` + `.html`, latest result per chassis). After the last run, `npm run report:pdi-daily` emails that copy once (to `GMAIL_TO`, CC `PDI_CC`), pass or fail, with problems under "Needs attention" and flagged in the subject. Running it again the same day rebuilds the files but doesn't resend (`PDI_FORCE_EMAIL=1` to resend). The generic "DMS Automation Report" email is not sent for PDI. Don't delete the files in `Output/PDI runs/` during the day: they are what the daily copy is built from (a rebuild refuses to replace a daily copy that covers more runs than are left).
- **Master and clean-up:** every rebuild also rewrites `Output/PDI_Master.xlsx`: a **Tracking** sheet first (all-time totals of blank PDI and In-Transit chassis: found, updated, still no usable VIN data, plus one row per day with running totals; a blank chassis counts as updated when it drops off the list only if that day's PDI run checked its dealer), then one sheet per date for the last 3 days (today included, newest first) copied from that day's daily file. A day whose daily file is missing keeps the sheet already in the master. Alongside it, `Output/PDI_Master.html` is an interactive dashboard over those days: chassis updated in total, a day-wise outcome chart, updated chassis per dealer per day, and every chassis row with day/outcome filters and search. The all-time history behind the Tracking sheet is kept in `Output/PDI_History.json`, which is never deleted. PDI files older than those 3 days are deleted automatically: run files and email markers in `Output/PDI runs/` and the `PDI_Daily_<date>` files. GSTR and other outputs are not touched. If the master is open in Excel it is left as it is and updated on the next run.
- **Same numbers everywhere:** this PC and the GitHub workflow share their PDI results through the `pdi-runs` branch (`pdi/scripts/pdiSync.js`). Before every run or report (`npm run test:pdi*`, `npm run report:pdi-daily`) each side takes the other's run files of the last 3 days, `PDI_History.json`, `PDI_InTransit_History.json` and the "emailed today" marker; afterwards it sends its own back. So the daily file, the "till now" totals and the master are the same whichever side builds them, and the day's email goes out once, from whichever side sends it first. `npm run pdi:sync` syncs by hand.
- **Report:** the email body has the day's totals, the all-time "till now" KPIs, a day-wise tracking table (last 14 days) and the per-dealer summary. Only Excel files are attached: the day's `PDI_Daily_<date>.xlsx` and `PDI_Master.xlsx`. The HTML dashboards (`PDI_Daily_<date>.html`, `PDI_Master.html`) are saved in `Output/` for local use.

### MIS check details
- **Date range:** the 1st of the month up to yesterday. Set `fromDate` / `toDate` in the `CONFIG` block at the top of `mis-check/run_mis_check.js` to use fixed dates.
- **Extra inputs** for reports that need more than dates (chassis number, spare part code, Dead Stock days, month/year) are also in `CONFIG.extraInputs`.
- **Comparison:** each run is compared with the newest earlier run in `Output/MIS/`, and added, removed or moved reports and changed columns are listed.
- **Colours in the Excel file:** red = the report failed, or its name or definition is missing; orange = no columns could be read.

---

## Settings (environment variables)

All optional. See `.env.example`.

| Variable | Used by | Purpose |
|---|---|---|
| `BROWSER_CHANNEL` | all | `chrome` (default), `msedge`, or `chromium` |
| `CREDENTIALS_FILE` | all | Use a credentials file other than `resources/credentials.json` |
| `GMAIL_CC` | Playwright email | Comma-separated CC list |
| `GSTR_FROM`, `GSTR_TO` | GSTR | Report dates (dd-mm-yyyy) |
| `GSTR_OUTPUT_DIR` | GSTR | Output folder (default `Output/`) |
| `GSTR_RERUN_ALL` | GSTR | `1` = don't reuse saved dealers |
| `ENQ_FROM`, `ENQ_TO` | Enquiry List | Report dates (dd-mm-yyyy); default 07-09-2026 to yesterday |
| `PDI_SHOW_WAIT_S` | PDI | Seconds to wait after clicking Show on the PDI Due List (default 30) |
| `PDI_REFRESHES` | PDI | Extra Show clicks when the PDI Due List says "Displaying : 0" (default 2) |
| `PDI_ACCEPT_VIN_COLUMNS` | PDI | `1` = skip the swapped-columns check once and take this VIN sheet as the new baseline |
| `PDI_FIX_FILE` | PDI | Correction run: a CSV of `dealer,chassis` lines. Re-saves Engine No, Battery and Charger from the VIN sheet for exactly those chassis, even though they are no longer blank |
| `PDI_PARALLEL` | PDI | Dealers processed at once, each in its own browser window (default 3; `1` = one at a time) |
| `PDI_LOAD_TIMEOUT_S` | PDI | Longest wait for Show / Save to reload the page (default 30) |
| `PDI_DEALER` | PDI | Only these dealers, comma-separated (default: all PDI dealers) |
| `PDI_DRY_RUN` | PDI | `1` = fill the form but don't Save (set by `test:pdi:dry`) |
| `PDI_VIN_FILE` | PDI | Use another VIN Details workbook or CSV (default: `resources/VIN_Details.csv`, the Google Sheet's copy) |
| `PDI_CC` | PDI daily report | CC list for the daily PDI email |
| `PDI_SEND_EMAIL` | PDI daily report | `0` = build the report but don't email it |
| `PDI_FORCE_EMAIL` | PDI daily report | `1` = send again although today's email already went out |
| `PDI_EMAIL_REVIEW` | PDI daily report | `1` = send a [REVIEW] copy to `GMAIL_TO` only (no CC); it doesn't count as the day's email (add `PDI_FORCE_EMAIL=1` once the day's email has gone) |
| `PDI_REBUILD_DAILY` | PDI | `1` = rebuild the day's copy even though it covers more runs than the run files left |
| `PDI_REPORT_DATE` | PDI daily report | Day to report on, `yyyy-mm-dd` (default today) |
| `MIS_PAGES` | MIS | Only these MIS pages, e.g. `CRM Reports,Workshop Reports` |
| `MIS_SEND_EMAIL` | MIS | `0` = don't send the summary email |
| `MIS_MAX_MINUTES` | MIS | Stop the whole run after this many minutes (default 60) |
| `MIS_PROFILE_DIR` | MIS | Browser profile folder (default `Output/MIS/dms-profile/`) |
| `HEADLESS`, `CI` | MIS | `1` = no visible browser; a failed login exits instead of waiting for Enter |
| `PW_GLOBAL_TIMEOUT_MIN` | Playwright | Cap for the whole test run (default 180) |

---

## CI/CD

`.github/workflows/playwright.yml` runs on GitHub Actions **on demand** (Actions → *DMS Automation* → *Run workflow*, then choose `all`, `enquiry-list`, `gstr` or `mis-check`) and **daily at 07:00 IST**. It runs headless on Google Chrome in Indian time, and keeps the reports as run downloads for 7 days.

Add these under **Settings → Secrets and variables → Actions → Secrets tab → New repository secret** (one secret per row; no variables are needed):

| Secret | Required | Value |
|---|---|---|
| `CREDENTIALS_JSON` | Yes | The whole of `resources/credentials.json`, pasted as-is (starts with `{`). CI uses `dealervalidUser` (Enquiry List, MIS check), `dealerUsers` (GSTR), `riverUserValidUser` (ASM report, scheduled PDI run) and `PIDLogin` / `PIDDealerCodes` / `PIDUsers` (`PDI.spec.ts`) |
| `GMAIL_USER` | For email | The Gmail address the reports are sent from |
| `GMAIL_APP_PASSWORD` | For email | That account's 16-character [Gmail App Password](https://myaccount.google.com/apppasswords) (not its login password) |
| `GMAIL_TO` | For email | Who receives the reports |
| `GMAIL_CC` | No | Comma-separated CC list for the Playwright run report |
| `PDI_CC` | No | CC list for the daily PDI email |

Then run it once: **Actions → DMS Automation → Run workflow**. A missing or broken `CREDENTIALS_JSON` stops the run at the first step with a clear error.

**PDI** has its own workflow, `.github/workflows/pdi.yml`: fill and **Save** hourly from 07:47 to 17:47 IST, and the day's one PDI email at 19:17 IST (off :00 / :30, when GitHub skips the most scheduled runs; it can still skip or delay some, so start a missed one from Actions). For now the hourly run is `pdi/PDI_Update_InTransit.spec.ts` only: one River User login (`riverUserValidUser`) downloads Inventory Stock Details and fills every vehicle with a blank Engine / Battery / Charger, whatever its status (the In-Transit filter is commented out). The per-dealer `pdi/PDI.spec.ts` step is commented out in the workflow; `npm run test:pdi` still runs it by hand. The VIN details come from the **VIN BIN Update (Sunitha)** Google Sheet, which itself stays private (the `VIN_Details.csv` copy on the `vin-data` branch is visible to anyone who can see this repository): an Apps Script ([resources/apps-script/pushVinToGitHub.gs](resources/apps-script/pushVinToGitHub.gs), setup steps at its top) opens it by ID every 15 minutes and, when the sheet has changed, sends it to `VIN_Details.csv` on the `vin-data` branch, which each run reads. So the team only keeps that sheet up to date. Only one copy of the script may be scheduled, or two sheets would overwrite each other. The day's run files, history and VIN-column baseline are carried between runs in the Actions cache. On demand: **Actions → PDI → Run workflow**, then choose `pdi`, `intransit` or `daily-email` (and optionally some dealer codes).

The GitHub-hosted runner must be able to reach `rivermobility.gaindms.com`. If it can't, use a self-hosted runner inside the company network.

**Exit codes of `mis-check/run_mis_check.js`:** `0` = the run finished (report failures are in the summary), `1` = the run itself failed (browser, login, or site down), `2` = stopped by `MIS_MAX_MINUTES`.

---

## Project structure

```
tests/
  enquiryListReport.spec.ts     Enquiry List
  GSTR Report.spec.ts           GSTR ProductWise Sales + invoice-PDF discounts
  models/                       Page objects (login, report viewer, report pages)
  reporters/gmail-reporter.ts   Emails the Playwright run result
pdi/
  PDI.spec.ts                   PDI blank Engine / Battery / Charger fill-up
  PDIDailyReport.spec.ts        Combined daily PDI report
  PDI_Update_InTransit.spec.ts  Blank vehicle details from Inventory Stock Details, as River User
  utils/                        PDI report, daily merge, history, Excel workbook
  scripts/pdiSync.js            Syncs PDI runs, histories and the emailed marker with the pdi-runs branch (PC and GitHub)
  scripts/pdiMergeHistory.js    Merges two PDI history files (used by pdiSync.js)
asm-report/                     ASM-wise report (Python) and its Mapping/ sheet (see its README)
mis-check/run_mis_check.js      MIS report checker
resources/                      credentials.example.json, apps-script/, local VIN sheet copy
docs/                           Confluence.md (team overview), test case workbook, vendor ticket
playwright.config.ts            Browser, timeouts, reporters
setup.ps1 / setup.bat           One-click setup (double-click setup.bat on a new PC)
Output/                         Everything the scripts generate (not in Git):
  Downloads/                    Report downloads (ASM/ and PDI/ have their own subfolders)
  ASM/                          ASM-wise summaries, email previews, logs
  MIS/                          MIS check results and its saved DMS login (dms-profile/)
  PDI runs/, PDI_*              PDI run files, daily report, history, master workbook
```

Kept out of Git (`.gitignore`): `.env`, `resources/credentials.json`, `resources/VIN*.xlsx` / `VIN*.csv`, `Output/` (contains customer data and the MIS check's saved DMS login), `node_modules/`, and the test run results.

---

## Troubleshooting

| Message | Fix |
|---|---|
| `Chromium distribution 'chrome' is not found` | Install Google Chrome, or set `BROWSER_CHANNEL=msedge` |
| `npx : ... running scripts is disabled on this system` | Use `npx.cmd`, or run once: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |
| `Login failed: still on the login page` | Check that user's details in `resources/credentials.json`. If it says `DMS says "You do not have permission for the selected branch"`, the password is right but that user has no access to the dealer: ask the DMS admin to grant it |
| `[WARN] Could not write … is it open in Excel?` | The file was saved under a new name; close Excel before the next run |
| MIS: `Could not start the browser … wait for the other run` | Another MIS check is using `Output/MIS/dms-profile/`; wait for it to finish |
| MIS: Crystal reports show `crv.js not served (HTTP 404)` | A DMS server issue: the Crystal Reports viewer files are missing on the server |
