# River DMS – E2E Automation

Automated checks and data fixes for **River DMS** (rivermobility.gaindms.com), built with Playwright. Each automation logs in to the live DMS, does its job, saves the results to Excel and emails a summary.

- **Repository:** https://github.com/sunita-river/E2E-DMS
- **Test cases:** `Test Cases/DMS_Automation_Test_Cases.xlsx` in the repository (83 cases, all passed)
- **Owner:** Sunita Sahu

---

## What the automation covers

| Automation | What it does | Changes DMS data? | Output |
|---|---|---|---|
| Enquiry List | Dealer login → Enquiry List, then Product-wise Sales Details (Invoice and Booking date types) → combined Enquiry / Invoice / Booking summary by date and dealer | No | Summary tables and a CSV |
| GSTR Report | For every dealer: GSTR ProductWise Sales Details plus the Post GST Discount read from each Product Sale invoice PDF | No | One Excel per dealer + a MasterData Excel |
| MIS Report Check | Opens every report under every MIS page (~274), classifies each (OK / server error / timeout / blank / needs input) and compares with the previous run | No | Formatted Excel + email summary |
| PDI Vehicle Details | For every PDI dealer: finds vehicles in WORKSHOP → PDI with a blank field and fills Engine / Battery / Charger from the VIN Details sheet in Settings → Admin → Modify vehicle details | **Yes** | Excel + interactive dashboard, emailed per dealer |

---

## PDI Vehicle Details

### Why
Vehicles waiting for PDI can have a blank Engine No (Motor No), Battery or Charger number. The correct values are in the manufacturer's VIN Details sheet. Filling them in by hand is slow, so the automation does it for every dealer.

### What it does, step by step
1. Logs in as each dealer listed under `PIDUsers` in the credentials file.
2. Opens **WORKSHOP → PDI**, raises the page's "Top" box from 25 to 500 and clicks Show, so every vehicle is listed.
3. Finds every vehicle with a blank field and saves its chassis number.
4. Opens **Settings → Admin → Modify vehicle details**. For each saved chassis:
   - enters the chassis number and clicks **Show**, then waits 10 seconds;
   - looks the chassis up in the VIN Details sheet;
   - fills **Engine No** (sheet column C, Motor No), **Battery Details** (column B) and **Vehicle Charger No** (column D);
   - adds the comment **"Details updated"** and clicks **Save**;
   - opens the chassis again to confirm the DMS kept the new values.
5. Logs out and moves on to the next dealer.

### Safety
- **Dry run first:** a dry run fills the form but never clicks Save, so you can check everything before changing live data.
- **Never runs by accident:** the PDI automation is skipped by the normal "run all tests" command and on CI. It only runs through its own commands.
- **Confirmed saves:** a chassis counts as updated only after it has been re-opened and shows the new values.
- **Timeouts:** if the DMS drops the session mid-run, the automation logs in again and retries that chassis once.
- **One dealer failing doesn't stop the others.**

### Pass or fail
A run **passes** only when every blank chassis was updated and no dealer failed. Anything else fails the run: a timeout that couldn't be recovered, a chassis missing from the VIN sheet, a failed login, or values the DMS didn't keep.

| Result | Report email |
|---|---|
| Passed | Sent to the report owner, with the PDI team on CC |
| Failed | **Not sent to anyone.** The results are still saved on the PC for the owner to check and re-run |
| Dry run | Sent to the report owner only |

### The report
- **Email body:** headline numbers (blank chassis found, updated, not in the VIN sheet, errors), a bar per dealer and a list of anything that needs attention.
- **Attached dashboard (PDI_Dashboard.html):** open it in a browser. Click a dealer to filter, filter by outcome, search chassis / engine / battery / charger numbers, sort any column. Works on a phone and in dark mode.
- **Attached Excel:** one row per chassis with the values used and the result, plus a sheet listing every dealer checked.
- **Daily report:** one combined email for all of the day's runs, showing each chassis's latest result.

### Results so far (28-Sep-2026)
**160 chassis updated across 12 dealers, 0 errors left.** 15 dealers were checked; 3 had nothing blank.

| Dealer | Updated | Dealer | Updated |
|---|---|---|---|
| 332008 | 20 | 372010 | 12 |
| 372002 | 17 | 292001 | 12 |
| 372009 | 17 | 292007 | 12 |
| 372007 | 14 | 292013 | 12 |
| 332009 | 13 | 292003 | 12 |
| 292006 | 11 | R81001 | 8 |

---

## How to run (Windows PC)

**First time:** clone the repository and run `setup.bat`. It installs Node.js, Google Chrome and everything else needed, then creates the two private settings files.

**Private files (never committed to Git):**

| File | What goes in it |
|---|---|
| `resources/credentials.json` | Dealer logins: `dealervalidUser` (Enquiry List, MIS), `dealerUsers` (GSTR), `PIDUsers` (PDI) |
| `.env` | Gmail sender, App Password, recipient, and the PDI CC list |
| `resources/VIN Details (1).xlsx` | The manufacturer's VIN sheet (A = VIN, B = Battery No, C = Motor No, D = Charger) |

**Commands** (in PowerShell use `npm.cmd` if scripts are blocked):

```
npm run test:enquiry-list      Enquiry List (~1–2 min)
npm run test:gstr              GSTR Report, all dealers (~15–60 min)
npm start                      MIS report check (~10 min)
npm run test:pdi:dry           PDI: fill the form, don't Save
npm run test:pdi               PDI: fill and Save (changes live data; ~25 s per chassis)
npm run report:pdi-daily       PDI: one combined report for today's runs
```

For PDI you can also double-click **run-pdi.bat** and choose D (dry run) or S (save).

**Useful settings** (set before the command, or in `.env`):

| Setting | What it does |
|---|---|
| `PDI_DEALER` | Only these PDI dealers, comma-separated, e.g. `332009,332008` |
| `PDI_TOP` | How many vehicles to list on the PDI page (default 500) |
| `PDI_VIN_FILE` | Use a different VIN sheet |
| `PDI_REPORT_DATE` | Daily report for another day (yyyy-mm-dd) |
| `ENQ_FROM`, `ENQ_TO` | Enquiry List dates (dd-mm-yyyy); default 07-09-2026 to today |
| `GSTR_FROM`, `GSTR_TO` | GSTR dates (dd-mm-yyyy) |
| `BROWSER_CHANNEL` | `msedge` to use Microsoft Edge instead of Chrome |

**Adding a PDI dealer:** add an entry to `PIDUsers` in `resources/credentials.json`, with the dealer code, username and password. No code change is needed.

---

## Scheduled runs (GitHub Actions)

Enquiry List, GSTR and the MIS check can run on GitHub every day at 07:00 IST, or on demand. PDI is **not** scheduled, because it changes live data and needs the VIN sheet.

**One-time setup:** in the repository go to **Settings → Secrets and variables → Actions → Secrets → New repository secret** and add:

| Secret | Required | Value |
|---|---|---|
| `CREDENTIALS_JSON` | Yes | The whole contents of `resources/credentials.json`, pasted as-is |
| `GMAIL_USER` | For email | The Gmail address the reports are sent from |
| `GMAIL_APP_PASSWORD` | For email | That account's 16-character Gmail App Password (not the login password) |
| `GMAIL_TO` | For email | Who receives the reports |
| `GMAIL_CC` | No | Comma-separated CC list |

No variables are needed. Then go to **Actions → DMS Automation → Run workflow** and run it once to confirm.

**Note:** GitHub's servers must be able to reach rivermobility.gaindms.com. If they can't, a self-hosted runner inside the company network is needed.

---

## Test coverage

| Area | Cases | Positive | Negative | Edge | Passed |
|---|---|---|---|---|---|
| Enquiry List | 16 | 14 | 1 | 1 | 16 |
| GSTR Report | 20 | 10 | 5 | 5 | 20 |
| MIS Report Check | 23 | 12 | 7 | 4 | 23 |
| PDI Vehicle Details | 24 | 12 | 8 | 4 | 24 |
| **Total** | **83** | **48** | **21** | **14** | **83** |

PDI negative cases tested include: a missing or placeholder login, an unknown dealer code, a missing VIN sheet, a wrong login, a chassis not in the VIN sheet, an invalid Top setting, a results file left open in Excel, and a day with no runs. The full list with steps and results is in the test case workbook.

---

## Troubleshooting

| Message | What to do |
|---|---|
| `running scripts is disabled on this system` | Use `npm.cmd` / `npx.cmd`, or run once: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |
| `still has the <user>/<password> placeholders` | Put the dealer's real login in `resources/credentials.json` |
| `Login failed: still on the login page` | Check that dealer's login in `resources/credentials.json` |
| `Not in VIN Details sheet` | The chassis isn't in the VIN sheet; get an updated sheet and run that dealer again |
| `Could not write … is it open in Excel?` | The file was saved under a new name; close Excel before the next run |
| `The grid is full at Top 500` | Set `PDI_TOP` higher and run again |
| A run seems stuck | Don't start two runs at once — each run clears the shared results folder |
