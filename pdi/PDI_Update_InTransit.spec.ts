import { test, Page, Locator } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from '../tests/models/LoginPage';
import { ReportViewerPage } from '../tests/models/ReportViewerPage';
import { vinFilePath, FailedDealer, writePdiDashboard, outcomeOf, problemsOf, splitProblems } from './utils/pdiReport';
import { writePdiWorkbook } from './utils/pdiWorkbook';
import { IN_TRANSIT_DIR, OUTPUT_DIR, buildPdiDaily, localDate } from './utils/pdiDaily';
import { loadVinDetails, updateVehicles, ResultRow, Log } from './utils/pdiVehicle';

// Vehicle details from Inventory Stock Details, all in one River User session (riverUserValidUser in
// credentials.json), no dealer logins:
//   1. Log in as River User, then MIS -> MFR Vehicles Reports -> Stock -> Inventory Stock Details -> View Report,
//      saved as Excel in "Output/Downloads/PDI/InventoryStockDetails_RiverUser_<date>.xlsx". It lists every
//      dealer's stock.
//   2. Rows with a blank Engine No. / Battery No / Vehicle ChargerNo are taken, whatever their Status (the
//      In-Transit filter is commented out in readStock for now). VINs missing from the VIN sheet are reported
//      without being opened.
//   3. In the same session, SETTINGS -> Admin -> Modify vehicle details: for each VIN, Show it, look it up in
//      the VIN Details sheet and fill Engine No (Motor No column), Battery and Charger, then Save and re-check
//      (same code as PDI.spec.ts). The values the DMS held before go in the "Was blank" column of the result file.
//
// Settings (all optional):
//   PDI_DEALER          only VINs under these Dealer Codes, comma-separated (default: every dealer in the report)
//   PDI_VINS            only these VINs: comma-separated, or the path of a text file with one per line (to
//                       rerun the ones that failed); still only those the report shows with a blank detail
//   PDI_FIX_VINS        VINs to correct from the VIN sheet even though they are already filled (same format):
//                       e.g. a vehicle holding a charger the sheet has since moved to another VIN, which makes
//                       the DMS refuse that VIN ("Vehicle Charger No must be unique"). They run first, so the
//                       numbers they free can be saved on the blank VINs after them. Old values are recorded.
//   PDI_DRY_RUN         1 = fill the form but don't click Save (for a first check)
//   PDI_VIN_FILE        VIN Details workbook or CSV (default: same as PDI.spec.ts)
//   PDI_LOAD_TIMEOUT_S  longest wait for Show / Save to reload the page (default 30)
//   STOCK_DOWNLOAD_DIR  where the Inventory Stock Details files go (default: "Output/Downloads/PDI")
//   STOCK_PAGE_WAIT_S   seconds to wait on Inventory Stock Details before View Report (default 5)
//
// The result file is Output/PDI runs/PDI_InTransit_<date-time>.xlsx (_DRYRUN added for a dry run). A real
// run then refreshes the day's PDI report (Output/PDI_Daily_<date>.xlsx + .html: In-Transit KPIs and sheet)
// and PDI_Master.xlsx, whose In-Transit sheet lists every chassis handled so far
// (Output/PDI_InTransit_History.json); npm run report:pdi-daily emails it as usual.
//
// Speed (measured 19-Oct-2026): login ~4s, Inventory Stock Details download ~100s (once per run), then ~2s
// per vehicle in a dry run; a real Save adds a Save and a re-check, so expect a few seconds more per vehicle.
const CREDENTIALS_FILE = process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
const VIN_FILE = vinFilePath();
const DRY_RUN = process.env.PDI_DRY_RUN === '1';
const LOAD_TIMEOUT = Number(process.env.PDI_LOAD_TIMEOUT_S || 30) * 1000;
const DOWNLOAD_DIR = process.env.STOCK_DOWNLOAD_DIR || path.join(OUTPUT_DIR, 'Downloads', 'PDI');
const PAGE_WAIT_MS = Number(process.env.STOCK_PAGE_WAIT_S || 5) * 1000;
const RIVER_USER = 'River User';

// MIS is a toggle: clicking it while its menu is open collapses it. Only click the parent when the
// item isn't already showing (see ProductWiseSalesDetailsReportPage).
async function clickMenuItem(parent: Locator, item: Locator) {
  if (!(await item.isVisible().catch(() => false))) {
    await parent.click();
    await item.waitFor({ state: 'visible', timeout: 10000 }).catch(async () => {
      await parent.click(); // the click closed an open menu; open it again
    });
  }
  await item.click();
}

async function downloadInventoryStock(page: Page, log: Log): Promise<string> {
  log('[STEP] Navigating to MIS > MFR Vehicles Reports > Stock > Inventory Stock Details...');
  await clickMenuItem(page.getByRole('link', { name: ' MIS ' }), page.getByRole('link', { name: ' MFR Vehicles Reports' }));
  await page.locator('#TableContent').getByText('Stock', { exact: true }).click();
  await page.locator('#TableContent').getByText('Inventory Stock Details').first().click();

  // Let the report's filter page finish loading before View Report.
  const viewer = new ReportViewerPage(page);
  await viewer.viewReportLink.waitFor({ state: 'visible', timeout: 60000 });
  await page.waitForTimeout(PAGE_WAIT_MS);

  const reportPage = await viewer.viewReport();
  const stamp = new Date().toISOString().slice(0, 10);
  const file = await viewer.saveReportExcel(reportPage, path.join(DOWNLOAD_DIR, `InventoryStockDetails_RiverUser_${stamp}.xlsx`));
  // The Excel is saved by now. Closing the viewer tab takes ~9s, so it closes in the background
  // while the updates start (the main window stays open, so Chrome keeps running).
  viewer.closeReport(reportPage, 'Inventory Stock Details').catch(() => {});
  return file;
}

// Vehicles from the report with a blank Engine No., Battery No or Vehicle ChargerNo (as in PDI.spec.ts), so
// hourly runs don't re-save vehicles that are already filled. Columns are found by header text. A VIN listed
// more than once is taken once, under the first Dealer Code it appears with.
function readStock(file: string): { work: { dealer: string; vin: string }[]; upToDate: number } {
  const workbook = XLSX.readFile(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
  const table = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: '' });
  const norm = (s: unknown) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const headerIndex = table.findIndex((r) => r.some((c) => norm(c) === 'status') && r.some((c) => norm(c).startsWith('vin')));
  if (headerIndex === -1) throw new Error(`No "Status" / "VIN No." header row in ${path.basename(file)}.`);
  const headers = table[headerIndex]!.map(norm);
  // const statusCol = headers.indexOf('status');
  const vinCol = headers.findIndex((h) => h.startsWith('vin'));
  const dealerCol = headers.indexOf('dealercode');
  // Report columns "Engine No.", "Battery No" and "Vehicle ChargerNo".
  const detailCols = [
    headers.findIndex((h) => h.startsWith('engine') || h.startsWith('motor')),
    headers.findIndex((h) => h.startsWith('battery')),
    headers.findIndex((h) => h.includes('charger')),
  ];
  if (detailCols.includes(-1)) throw new Error(`No Engine No. / Battery No / Vehicle ChargerNo column in ${path.basename(file)}.`);
  const rows = table.slice(headerIndex + 1)
    // In-Transit only, off for now: every vehicle with a blank detail is taken.
    // .filter((r) => norm(r[statusCol]) === 'intransit')
    .map((r) => ({
      dealer: dealerCol === -1 ? '' : String(r[dealerCol]).trim().toUpperCase(),
      vin: String(r[vinCol]).trim().toUpperCase(),
      blank: detailCols.some((c) => !String(r[c] ?? '').trim()),
    }))
    .filter((r) => r.vin);

  const seen = new Set<string>();
  const work: { dealer: string; vin: string }[] = [];
  let upToDate = 0;
  for (const r of rows) {
    if (!r.blank) upToDate++;
    else if (!seen.has(r.vin)) {
      seen.add(r.vin);
      work.push({ dealer: r.dealer, vin: r.vin });
    }
  }
  return { work, upToDate };
}

test.describe('PDI In-Transit', () => {
  test.use({ storageState: { cookies: [], origins: [] } });
  // This writes to live DMS records, so it never runs as part of a plain `npm test`:
  // only via npm run test:pdi:intransit / test:pdi:intransit:dry, which set PDI_RUN=1.
  test.skip(process.env.PDI_RUN !== '1', 'PDI In-Transit update runs only via npm run test:pdi:intransit (sets PDI_RUN=1).');

  test('Should fill blank vehicle details from the VIN Details sheet', async ({ page }) => {
    const riverUser: UserCredentials = credentials.riverUserValidUser;
    if (!riverUser?.username) throw new Error(`No riverUserValidUser in ${path.basename(CREDENTIALS_FILE)}.`);
    const vinDetails = loadVinDetails(VIN_FILE);
    console.log(`[STEP] Logging in as River User ${riverUser.username}; ${vinDetails.size} VINs loaded from ${path.basename(VIN_FILE)}${DRY_RUN ? ' (DRY RUN, no Save)' : ''}`);

    const start = Date.now();
    const log: Log = (message) => console.log(`[${RIVER_USER} +${Math.round((Date.now() - start) / 1000)}s] ${message}`);
    const loginPage = new LoginPage(page, log);
    const login = async () => {
      await loginPage.goto();
      await loginPage.loginAsRiverUser(riverUser);
      await loginPage.verifyDashboard();
    };

    // 1. One login, one download.
    test.setTimeout(15 * 60 * 1000);
    await login();
    const stockFile = await downloadInventoryStock(page, log);
    log(`[STEP] Inventory Stock Details saved: ${stockFile}`);
    const { work: all, upToDate } = readStock(stockFile);

    // 2. Work list, narrowed by PDI_DEALER and PDI_VINS. A VIN that isn't in the VIN sheet has nothing to save,
    //    so it is reported here without being opened.
    const list = (s: string) => new Set(s.split(/[,\r\n]+/).map((v) => v.trim().toUpperCase()).filter(Boolean));
    const vinList = (setting: string) => list(setting && fs.existsSync(setting) ? fs.readFileSync(setting, 'utf8') : setting);
    const wanted = list(process.env.PDI_DEALER || '');
    const wantedVins = vinList(process.env.PDI_VINS || '');
    const fixVins = vinList(process.env.PDI_FIX_VINS || '');
    const work = all.filter((w) => (!wanted.size || wanted.has(w.dealer)) && (!wantedVins.size || wantedVins.has(w.vin)) && !fixVins.has(w.vin));
    if (wantedVins.size) {
      const listed = new Set(all.map((w) => w.vin));
      const filled = [...wantedVins].filter((v) => !listed.has(v));
      console.log(`[STEP] PDI_VINS: ${wantedVins.size} VIN(s) asked, ${work.length} still blank in the report` +
        `${filled.length ? `; not blank any more (nothing to do): ${filled.join(', ')}` : ''}.`);
    }
    const results: ResultRow[] = [];
    const toUpdate: ResultRow[] = [];
    // Corrections (PDI_FIX_VINS) are kept apart, on the result file's Corrections sheet, so they don't count
    // as blank vehicles in the day's numbers. Their dealer comes from the report when they are in stock.
    const dealerOf = new Map(all.map((w) => [w.vin, w.dealer]));
    const corrections: ResultRow[] = [...fixVins].map((vin) => ({ Dealer: dealerOf.get(vin) ?? '', ChassisNo: vin, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: '' }));
    if (corrections.length) console.log(`[STEP] PDI_FIX_VINS: ${corrections.length} filled VIN(s) to correct from the VIN sheet first.`);
    for (const { dealer, vin } of work) {
      const row: ResultRow = { Dealer: dealer, ChassisNo: vin, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: '' };
      if (vinDetails.has(vin)) toUpdate.push(row);
      else results.push({ ...row, Status: 'Not in VIN Details sheet' });
    }
    const notInSheet = results.length;
    if (notInSheet) console.log(`[WARN] ${notInSheet} vehicle(s) with blank details are not in the VIN Details sheet; reported, nothing to save.`);
    const dealers = [...new Set(work.map((w) => w.dealer))];
    console.log(`[STEP] ${work.length} vehicle(s) with a blank Engine / Battery / Charger across ${dealers.length} dealer(s) ` +
      `(${upToDate} already filled); ${toUpdate.length} to update as River User.`);

    // About 10s per vehicle (Save + re-check), with headroom.
    test.setTimeout(15 * 60 * 1000 + (corrections.length + toUpdate.length) * 10 * 1000 * 2);

    // 3. Same session: every vehicle, one after another, corrections first.
    const failures: FailedDealer[] = [];
    if (corrections.length || toUpdate.length) {
      try {
        if (corrections.length) {
          await updateVehicles(page, loginPage, riverUser, corrections, vinDetails, log,
            { dryRun: DRY_RUN, loadTimeoutMs: LOAD_TIMEOUT, recordBefore: 'Correction, before', relogin: login });
        }
        if (toUpdate.length) {
          results.push(...await updateVehicles(page, loginPage, riverUser, toUpdate, vinDetails, log,
            { dryRun: DRY_RUN, loadTimeoutMs: LOAD_TIMEOUT, recordBefore: 'Before', relogin: login }));
        }
      } catch (err) {
        const message = String((err as Error).message).split('\n')[0]!.slice(0, 200);
        failures.push({ dealer: RIVER_USER, error: message });
        log(`[ERROR] Update failed: ${message}`);
        corrections.forEach((r) => { r.Status ||= `Not attempted: ${message}`; });
        results.push(...toUpdate.filter((r) => !results.includes(r)).map((r) => ({ ...r, Status: r.Status || `Not attempted: ${message}` })));
        await test.info().attach('river-user-failure.png', {
          body: await page.screenshot({ fullPage: true, timeout: 15000 }).catch(() => Buffer.from('')),
          contentType: 'image/png',
        }).catch(() => {});
      }
    }
    results.sort((a, b) => a.Dealer!.localeCompare(b.Dealer!));

    const stamp = new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-');
    fs.mkdirSync(IN_TRANSIT_DIR, { recursive: true });
    const file = await writePdiWorkbook(path.join(IN_TRANSIT_DIR, `PDI_InTransit_${stamp}${DRY_RUN ? '_DRYRUN' : ''}.xlsx`), {
      title: 'PDI vehicle details (Inventory Stock Details)',
      subtitle: `${new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })} · River User ${riverUser.username} · ${dealers.length} dealer(s) · VIN sheet: ${path.basename(VIN_FILE)}`,
      dryRun: DRY_RUN, dealers, failed: failures, rows: results, rowsSheet: 'In-Transit',
      ...(corrections.length ? { corrections } : {}),
    });
    console.log(`\n[STEP] Results saved to ${file}`);

    // A real run refreshes the day's PDI report so its "In-Transit updated" KPI is current; no email here.
    if (!DRY_RUN) {
      try {
        const daily = await buildPdiDaily(localDate(new Date()));
        if (daily) writePdiDashboard(daily.results, daily.failures, daily.meta, daily.file);
      } catch (err) {
        console.log(`[WARN] Day's PDI report not refreshed: ${(err as Error).message}`);
      }
    }

    const isDone = (r: ResultRow) => /^(Updated|Dry run)/.test(r.Status || '');
    if (corrections.length) {
      console.log(`[RESULT] Corrections: ${corrections.filter(isDone).length} of ${corrections.length} ${DRY_RUN ? 'filled (dry run, nothing saved)' : 'corrected'}`);
      corrections.forEach((r) => console.log(`   ${isDone(r) ? 'OK ' : 'NOT'} ${r.ChassisNo}  ${r.BlankFields} -> Engine ${r.EngineNo}, Battery ${r.BatteryDetails}, Charger ${r.VehicleChargerNo}: ${r.Status}`));
    }
    const done = results.filter(isDone);
    console.log(`[RESULT] ${results.length} vehicle(s), ${done.length} ${DRY_RUN ? 'filled (dry run, nothing saved)' : 'updated'}`);
    results.filter((r) => !done.includes(r)).forEach((r) => console.log(`   NOT ${DRY_RUN ? 'filled' : 'updated'}: ${r.Dealer} ${r.ChassisNo}  ${r.Status}`));
    // As in PDI.spec.ts, data problems only warn (they are in the report; a rerun can't fix them): not in the
    // VIN sheet, a Save the DMS refuses (e.g. "Battery No must be unique"). Timeouts and page errors fail the run.
    const problems = splitProblems(problemsOf([...corrections, ...results].filter((r) => outcomeOf(r) !== 'missing'), failures, DRY_RUN));
    if (problems.length) {
      throw new Error(`${problems.length} problem(s).\n  ${problems.slice(0, 20).join('\n  ')}`);
    }
  });
});
