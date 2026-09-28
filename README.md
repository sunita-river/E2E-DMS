# E2E-DMS

End-to-end Playwright automation for **River DMS** (`rivermobility.gaindms.com`). It checks reports on the live system and produces Excel outputs and email summaries.

| Script | What it does | Output |
|---|---|---|
| `tests/enquiryListReport.spec.ts` | Dealer login → Enquiry List, then Product-wise Sales Details (Invoice and Booking date types) → combined Enquiry / Invoice / Booking summary by date and dealer | Summary tables in the log, `combined-summary.csv` attached to the test result |
| `tests/GSTR Report.spec.ts` | For every dealer in `credentials.json`: GSTR ProductWise Sales Details, plus the **Post GST Discount** read from each Product Sale invoice PDF | `Output/<dealer>_SalesGSTNReport_<from>_<to>.xlsx` and `MasterData_…xlsx` |
| `tests/PDI.spec.ts` | For every dealer in `PIDUsers`: finds **WORKSHOP → PDI** rows with a blank field and fills Engine / Battery / Charger from the VIN Details sheet in **Settings → Admin → Modify vehicle details** | `Output/PDI_BlankChassis_<time>.xlsx` + dashboard, emailed per dealer when the run passes |
| `run_mis_check.js` | Opens every report under every **MIS** page (~274), classifies each one (OK / server error / timeout / blank / needs input), records its definition and available columns, and compares with the previous run | `MIS Reports/` (formatted Excel, CSV, summary, snapshot) and an email summary |

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
| `resources/credentials.json` | `resources/credentials.example.json` | `dealervalidUser` (Enquiry List, MIS check), `dealerUsers` (one entry per dealer for GSTR) and `PIDUsers` (one entry per dealer for PDI) |
| `.env` | `.env.example` | `GMAIL_USER`, `GMAIL_APP_PASSWORD` (a [Gmail App Password](https://myaccount.google.com/apppasswords)), `GMAIL_TO`; optional `PDI_CC` |
| `resources/VIN Details (1).xlsx` | copy it in by hand | VIN sheet for PDI: A = VIN, B = Battery No, C = Motor No, D = Charger |

---

## Running

```powershell
npm run test:enquiry-list      # Enquiry List (~1-2 min)
npm run test:gstr              # GSTR Report, all dealers (~15-60 min)
npm run test:pdi:dry           # PDI: fill the form for every blank chassis, don't Save
npm run test:pdi               # PDI: fill and Save (changes live DMS records; ~25 s per chassis)
npm run report:pdi-daily       # PDI: one combined report for all of today's runs
npm start                      # MIS report check (~10 min)
npm test                       # Enquiry List + GSTR (PDI is skipped unless run with the commands above)
npm run report:html            # open the last Playwright HTML report
```

To watch the browser, add `-- --headed` to the Playwright commands (for example `npm run test:gstr -- --headed`). By default the tests run on Google Chrome. To use Microsoft Edge instead:

```powershell
$env:BROWSER_CHANNEL='msedge'; npm run test:gstr
```

After each Playwright run, `tests/reporters/gmail-reporter.ts` emails the pass/fail status to `GMAIL_TO`. The MIS check sends its own HTML summary with the Excel file attached.

### GSTR Report details
- **Date range:** 01-09-2026 to 24-09-2026 by default. Override with `GSTR_FROM` / `GSTR_TO` (dd-mm-yyyy).
- **Resumable:** a dealer whose file for the same date range already exists in `Output/` is reused, not re-run. Set `GSTR_RERUN_ALL=1` to redo every dealer.
- **One dealer failing doesn't stop the run.** It's retried once at the end, and MasterData is always written. The test fails (red) only after everything that could be saved is saved.
- **An invoice that can't be read is skipped** with a `[WARN] Reference …` line; check those by hand.

### PDI details
- **Double-click `run-pdi.bat`** to choose Dry run or Save; the npm commands above do the same. In PowerShell use `npm.cmd` if scripts are blocked.
- **Checked:** every vehicle on the PDI page — the spec raises the page's "Top" box (25 by default) to `PDI_TOP` (500) and clicks Show first, and warns if the list is full. After Save, each chassis is opened again to confirm the DMS kept the values; a timeout triggers a fresh login and one retry.
- **Pass / fail:** the run passes only when every blank chassis is updated and no dealer fails. Only a passing run emails the report (to `GMAIL_TO`, CC `PDI_CC`); a failed run is red and sends no report — see the dashboard and Excel in `Output/`. Dry-run reports go to `GMAIL_TO` only.
- **Report:** per-dealer summary in the email body; `PDI_Dashboard.html` attached (open in a browser to filter by dealer or outcome, search, sort) with the Excel.

### MIS check details
- **Date range:** the 1st of the month up to yesterday. Set `fromDate` / `toDate` in the `CONFIG` block at the top of `run_mis_check.js` to use fixed dates.
- **Extra inputs** for reports that need more than dates (chassis number, spare part code, Dead Stock days, month/year) are also in `CONFIG.extraInputs`.
- **Comparison:** each run is compared with the newest earlier run in `MIS Reports/`, and added, removed or moved reports and changed columns are listed.
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
| `ENQ_FROM`, `ENQ_TO` | Enquiry List | Report dates (dd-mm-yyyy); default 07-09-2026 to today |
| `PDI_TOP` | PDI | Vehicles to list on the PDI page (default 500) |
| `PDI_DEALER` | PDI | Only these dealers, comma-separated (default: all `PIDUsers`) |
| `PDI_DRY_RUN` | PDI | `1` = fill the form but don't Save (set by `test:pdi:dry`) |
| `PDI_VIN_FILE` | PDI | VIN Details workbook (default `resources/VIN Details (1).xlsx`) |
| `PDI_CC` | PDI | CC list for the PDI report emails (passing, non-dry runs only) |
| `PDI_SEND_EMAIL` | PDI | `0` = build the report but don't email it |
| `PDI_REPORT_DATE` | PDI daily report | Day to report on, `yyyy-mm-dd` (default today) |
| `MIS_PAGES` | MIS | Only these MIS pages, e.g. `CRM Reports,Workshop Reports` |
| `MIS_SEND_EMAIL` | MIS | `0` = don't send the summary email |
| `MIS_MAX_MINUTES` | MIS | Stop the whole run after this many minutes (default 60) |
| `MIS_PROFILE_DIR` | MIS | Browser profile folder (default `dms-profile/`) |
| `HEADLESS`, `CI` | MIS | `1` = no visible browser; a failed login exits instead of waiting for Enter |
| `PW_GLOBAL_TIMEOUT_MIN` | Playwright | Cap for the whole test run (default 180) |

---

## CI/CD

`.github/workflows/playwright.yml` runs on GitHub Actions **on demand** (Actions → *DMS Automation* → *Run workflow*, then choose `all`, `enquiry-list`, `gstr` or `mis-check`) and **daily at 07:00 IST**. It runs headless on Google Chrome in Indian time, and keeps the reports as run downloads for 7 days.

Add these under **Settings → Secrets and variables → Actions**:

| Secret | Value |
|---|---|
| `CREDENTIALS_JSON` | The full contents of your `resources/credentials.json` |
| `GMAIL_USER`, `GMAIL_APP_PASSWORD`, `GMAIL_TO` | As in `.env` |
| `GMAIL_CC` | Optional |

The GitHub-hosted runner must be able to reach `rivermobility.gaindms.com`. If it can't, use a self-hosted runner inside the company network.

**Exit codes of `run_mis_check.js`:** `0` = the run finished (report failures are in the summary), `1` = the run itself failed (browser, login, or site down), `2` = stopped by `MIS_MAX_MINUTES`.

---

## Project structure

```
tests/
  enquiryListReport.spec.ts     Enquiry List + Product-wise Sales Details
  GSTR Report.spec.ts           GSTR ProductWise Sales + invoice-PDF discounts
  PDI.spec.ts                   PDI blank Engine / Battery / Charger fill-up
  PDIDailyReport.spec.ts        Combined daily PDI report
  utils/pdiReport.ts            PDI report email + interactive dashboard
  models/                       Page objects (login, report viewer, report pages)
  reporters/gmail-reporter.ts   Emails the Playwright run result
run_mis_check.js                MIS report checker
playwright.config.ts            Browser, timeouts, reporters
resources/credentials.example.json
Test Cases/                     Test case workbook (executed 25-Sep-2026; PDI 28-Sep-2026)
setup.ps1 / setup.bat           One-click setup;  run-tests.bat = setup + run all tests
run-pdi.bat                     Double-click PDI run (asks Dry run or Save)
```

Kept out of Git (`.gitignore`): `.env`, `resources/credentials.json`, `resources/VIN Details*.xlsx`, `dms-profile/` (saved DMS login), `Output/` and `MIS Reports/` (contain customer data), `node_modules/`, and the test run results.

---

## Troubleshooting

| Message | Fix |
|---|---|
| `Chromium distribution 'chrome' is not found` | Install Google Chrome, or set `BROWSER_CHANNEL=msedge` |
| `npx : ... running scripts is disabled on this system` | Use `npx.cmd`, or run once: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |
| `Login failed: still on the login page` | Check that user's details in `resources/credentials.json` |
| `[WARN] Could not write … is it open in Excel?` | The file was saved under a new name; close Excel before the next run |
| MIS: `Could not start the browser … wait for the other run` | Another MIS check is using `dms-profile/`; wait for it to finish |
| MIS: Crystal reports show `crv.js not served (HTTP 404)` | A DMS server issue: the Crystal Reports viewer files are missing on the server |
