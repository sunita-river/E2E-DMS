# E2E-DMS

End-to-end Playwright automation for **River DMS** (`rivermobility.gaindms.com`). It checks reports on the live system and produces Excel outputs and email summaries.

| Script | What it does | Output |
|---|---|---|
| `tests/enquiryListReport.spec.ts` | Dealer login → Enquiry List (full list, every row) | `Output/Enquiry_List_<from>_<to>.xlsx` |
| `tests/GSTR Report.spec.ts` | For every dealer in `credentials.json`: GSTR ProductWise Sales Details, plus the **Post GST Discount** read from each Product Sale invoice PDF | `Output/<dealer>_SalesGSTNReport_<from>_<to>.xlsx` and `MasterData_…xlsx` |
| `tests/PDI.spec.ts` | For every PDI dealer (`PIDDealerCodes` / `PIDUsers`): finds **WORKSHOP → PDI** rows with a blank field and fills Engine / Battery / Charger from the VIN Details sheet in **Settings → Admin → Modify vehicle details** | one copy per day, `Output/PDI_Daily_<date>.xlsx` + dashboard, refreshed by every run (each run's own file goes to `Output/PDI runs/`); no email per run |
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
| `resources/credentials.json` | `resources/credentials.example.json` | `dealervalidUser` (Enquiry List, MIS check), `dealerUsers` (one entry per dealer for GSTR) and, for PDI, `PIDLogin` (the shared login) with `PIDDealerCodes` (the dealer codes), plus `PIDUsers` for any dealer with its own login |
| `.env` | `.env.example` | `GMAIL_USER`, `GMAIL_APP_PASSWORD` (a [Gmail App Password](https://myaccount.google.com/apppasswords)), `GMAIL_TO`; optional `PDI_CC` |
| `resources/VIN Details.xlsx` | copy it in by hand (any `VIN Details….xlsx` name works; the newest is used) | VIN sheet for PDI: A = VIN, B = Battery No, C = Motor No, D = Charger |

---

## Running

```powershell
npm run test:enquiry-list      # Enquiry List (~1-2 min)
npm run test:gstr              # GSTR Report, all dealers (~15-60 min)
npm run test:pdi:dry           # PDI: fill the form for every blank chassis, don't Save
npm run test:pdi               # PDI: fill and Save (changes live DMS records; ~25 s per chassis)
npm run report:pdi-daily       # PDI: the one email of the day, for all of today's runs (run it after the last PDI run)
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
- **Run it** with `npm run test:pdi:dry` (fill only) or `npm run test:pdi` (fill and Save). In PowerShell use `npm.cmd` if scripts are blocked.
- **Excel:** a formatted workbook — a Summary sheet (headline figures and a table by dealer), then one row per chassis with the result colour-coded, then every dealer checked. Header rows are frozen and filterable.
- **Checked:** every vehicle on the PDI page — the spec raises the page's "Top" box (25 by default) to `PDI_TOP` (500) and clicks Show first, and warns if the list is full. After Save, each chassis is opened again to confirm the DMS kept the values; a timeout triggers a fresh login and one retry.
- **Speed:** `PDI_PARALLEL` dealers (default 3) run at once, each in its own browser window. After Show / Save the spec waits only until the page has reloaded (up to `PDI_LOAD_TIMEOUT_S`, 30s) instead of a fixed 10s. If the DMS allows only one session per login, set `PDI_PARALLEL=1`.
- **Pass / fail:** the run passes only when every blank chassis is updated and no dealer fails. A failed run is still red. Runs send **no email**: rerun `npm run test:pdi` as often as needed during the day, and each run refreshes the day's single copy (`Output/PDI_Daily_<date>.xlsx` + `.html`, latest result per chassis). After the last run, `npm run report:pdi-daily` emails that copy once (to `GMAIL_TO`, CC `PDI_CC`), pass or fail, with problems under "Needs attention" and flagged in the subject. Running it again the same day rebuilds the files but doesn't resend (`PDI_FORCE_EMAIL=1` to resend). The generic "DMS Automation Report" email is not sent for PDI. Don't delete the files in `Output/PDI runs/` during the day: they are what the daily copy is built from (a rebuild refuses to replace a daily copy that covers more runs than are left).
- **Master and clean-up:** every rebuild also rewrites `Output/PDI_Master.xlsx`: a **Tracking** sheet first (all-time totals: chassis found, updated, still no usable VIN data, plus one row per day with running totals), then one sheet per date for the last 3 days (today included, newest first) copied from that day's daily file. A day whose daily file is missing keeps the sheet already in the master. Alongside it, `Output/PDI_Master.html` is an interactive dashboard over those days: chassis updated in total, a day-wise outcome chart, updated chassis per dealer per day, and every chassis row with day/outcome filters and search. The all-time history behind the Tracking sheet is kept in `Output/PDI_History.json`, which is never deleted. PDI files older than those 3 days are deleted automatically: run files and email markers in `Output/PDI runs/` and the `PDI_Daily_<date>` files. GSTR and other outputs are not touched. If the master is open in Excel it is left as it is and updated on the next run.
- **Report:** the email body has the day's totals, the all-time "till now" KPIs, a day-wise tracking table (last 14 days) and the per-dealer summary. Only Excel files are attached: the day's `PDI_Daily_<date>.xlsx` and `PDI_Master.xlsx`. The HTML dashboards (`PDI_Daily_<date>.html`, `PDI_Master.html`) are saved in `Output/` for local use.

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
| `ENQ_FROM`, `ENQ_TO` | Enquiry List | Report dates (dd-mm-yyyy); default 07-09-2026 to yesterday |
| `PDI_TOP` | PDI | Vehicles to list on the PDI page (default 500) |
| `PDI_PARALLEL` | PDI | Dealers processed at once, each in its own browser window (default 3; `1` = one at a time) |
| `PDI_LOAD_TIMEOUT_S` | PDI | Longest wait for Show / Save to reload the page (default 30) |
| `PDI_DEALER` | PDI | Only these dealers, comma-separated (default: all PDI dealers) |
| `PDI_DRY_RUN` | PDI | `1` = fill the form but don't Save (set by `test:pdi:dry`) |
| `PDI_VIN_FILE` | PDI | VIN Details workbook (default: the newest `resources/VIN Details*.xlsx`) |
| `PDI_CC` | PDI daily report | CC list for the daily PDI email |
| `PDI_SEND_EMAIL` | PDI daily report | `0` = build the report but don't email it |
| `PDI_FORCE_EMAIL` | PDI daily report | `1` = send again although today's email already went out |
| `PDI_EMAIL_REVIEW` | PDI daily report | `1` = send a [REVIEW] copy to `GMAIL_TO` only (no CC); it doesn't count as the day's email (add `PDI_FORCE_EMAIL=1` once the day's email has gone) |
| `PDI_REBUILD_DAILY` | PDI | `1` = rebuild the day's copy even though it covers more runs than the run files left |
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

Add these under **Settings → Secrets and variables → Actions → Secrets tab → New repository secret** (one secret per row; no variables are needed):

| Secret | Required | Value |
|---|---|---|
| `CREDENTIALS_JSON` | Yes | The whole of `resources/credentials.json`, pasted as-is (starts with `{`). CI uses `dealervalidUser` (Enquiry List, MIS check) and `dealerUsers` (GSTR); the PDI entries can stay in or be left out, as PDI doesn't run on CI |
| `GMAIL_USER` | For email | The Gmail address the reports are sent from |
| `GMAIL_APP_PASSWORD` | For email | That account's 16-character [Gmail App Password](https://myaccount.google.com/apppasswords) (not its login password) |
| `GMAIL_TO` | For email | Who receives the reports |
| `GMAIL_CC` | No | Comma-separated CC list for the Playwright run report |

Then run it once: **Actions → DMS Automation → Run workflow**. A missing or broken `CREDENTIALS_JSON` stops the run at the first step with a clear error. PDI isn't part of CI on purpose: it changes live DMS records and needs the VIN sheet, so it's run by hand.

The GitHub-hosted runner must be able to reach `rivermobility.gaindms.com`. If it can't, use a self-hosted runner inside the company network.

**Exit codes of `run_mis_check.js`:** `0` = the run finished (report failures are in the summary), `1` = the run itself failed (browser, login, or site down), `2` = stopped by `MIS_MAX_MINUTES`.

---

## Project structure

```
tests/
  enquiryListReport.spec.ts     Enquiry List
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
docs/Confluence.md              Team-facing overview, ready to paste into Confluence
setup.ps1 / setup.bat           One-click setup (double-click setup.bat on a new PC)
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
