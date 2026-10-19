import { test, Page, Locator } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from '../tests/models/LoginPage';
import { ReportViewerPage } from '../tests/models/ReportViewerPage';
import { vinFilePath, FailedDealer, writePdiDashboard, outcomeOf, problemsOf, splitProblems } from './utils/pdiReport';
import { writePdiWorkbook } from './utils/pdiWorkbook';
import { IN_TRANSIT_DIR, buildPdiDaily, localDate } from './utils/pdiDaily';
import { loadVinDetails, updateVehicles, ResultRow, Log } from './utils/pdiVehicle';

// In-Transit vehicle details, for every PDI dealer (PIDDealerCodes / PIDUsers in credentials.json):
//   1. Once, with one login: MIS -> MFR Vehicles Reports -> Stock -> Inventory Stock Details -> View Report,
//      saved as Excel in "PDI Update/Downloads/InventoryStockDetails_<dealer>_<date>.xls(x)". It lists every
//      dealer's stock.
//   2. Rows with Status "In-Transit" and a blank Engine No. / Battery No / Vehicle ChargerNo go to the
//      dealer in their Dealer Code column. VINs missing from the VIN sheet are reported straight away;
//      dealers with no VIN left to save are not logged into.
//   3. INTRANSIT_PARALLEL dealers at a time, each in its own browser window and session: log in as the dealer,
//      then SETTINGS -> Admin -> Modify vehicle details: for each of its VINs, Show it, look it up in the VIN
//      Details sheet and fill Engine No (Motor No column), Battery and Charger, then Save and re-check (same
//      code as PDI.spec.ts). The values the DMS held before go in the "Was blank" column of the result file.
//   Each VIN belongs to exactly one dealer, so no two windows touch the same vehicle; result files are
//   written once, after every window is done.
//
// Settings (all optional):
//   PDI_DEALER          only these dealer codes, comma-separated (default: all PDI dealers)
//   PDI_DRY_RUN         1 = fill the form but don't click Save (for a first check)
//   INTRANSIT_PARALLEL  dealers updated at once, each in its own browser window (default 3; 1 = one at a time)
//   PDI_VIN_FILE        VIN Details workbook or CSV (default: same as PDI.spec.ts)
//   PDI_LOAD_TIMEOUT_S  longest wait for Show / Save to reload the page (default 30)
//   STOCK_DOWNLOAD_DIR  where the Inventory Stock Details files go (default: "PDI Update/Downloads")
//   STOCK_PAGE_WAIT_S   seconds to wait on Inventory Stock Details before View Report (default 5)
//
// The result file is Output/PDI runs/PDI_InTransit_<date-time>.xlsx (_DRYRUN added for a dry run). A real
// run then refreshes the day's PDI report (Output/PDI_Daily_<date>.xlsx + .html: In-Transit KPIs and sheet)
// and PDI_Master.xlsx, whose In-Transit sheet lists every In-Transit chassis handled so far
// (Output/PDI_InTransit_History.json); npm run report:pdi-daily emails it as usual.
//
// Speed (measured 19-Oct-2026): login ~4s, Inventory Stock Details download ~100s (once per run), then ~2s
// per vehicle in a dry run; a real Save adds a Save and a re-check, so expect a few seconds more per vehicle.
const CREDENTIALS_FILE = process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
const VIN_FILE = vinFilePath();
const DRY_RUN = process.env.PDI_DRY_RUN === '1';
const LOAD_TIMEOUT = Number(process.env.PDI_LOAD_TIMEOUT_S || 30) * 1000;
const DOWNLOAD_DIR = process.env.STOCK_DOWNLOAD_DIR || path.join(__dirname, 'Downloads');
const PAGE_WAIT_MS = Number(process.env.STOCK_PAGE_WAIT_S || 5) * 1000;
const PARALLEL = Number(process.env.INTRANSIT_PARALLEL || 3);
if (!Number.isInteger(PARALLEL) || PARALLEL < 1) throw new Error(`INTRANSIT_PARALLEL must be a whole number above 0 (got "${process.env.INTRANSIT_PARALLEL}")`);

// Same dealer list as PDI.spec.ts: PIDUsers (own login) plus each PIDDealerCodes code with the shared
// PIDLogin, narrowed by PDI_DEALER (which may also name dealerUsers / dealervalidUser).
function pickDealers(): UserCredentials[] {
  const ownLogin: UserCredentials[] = credentials.PIDUsers || [];
  const listed = new Set(ownLogin.map((d) => String(d.dealercode).toLowerCase()));
  const codes: string[] = (credentials.PIDDealerCodes || []).map((c: unknown) => String(c).trim()).filter(Boolean);
  if (codes.length && !credentials.PIDLogin) throw new Error(`PIDDealerCodes needs a PIDLogin (username, password) in ${path.basename(CREDENTIALS_FILE)}.`);
  const pidUsers: UserCredentials[] = [
    ...ownLogin,
    ...codes.filter((c) => !listed.has(c.toLowerCase())).map((dealercode) => ({ ...credentials.PIDLogin, dealercode })),
  ];
  const wanted = [...new Set((process.env.PDI_DEALER || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean))];
  const all: UserCredentials[] = [...pidUsers, credentials.dealervalidUser, ...(credentials.dealerUsers || [])].filter(Boolean);
  const dealers = wanted.length
    ? wanted.map((code) => {
        const match = all.find((d) => String(d.dealercode).toLowerCase() === code);
        if (!match) throw new Error(`PDI_DEALER: dealer ${code} is not in ${path.basename(CREDENTIALS_FILE)}.`);
        return match;
      })
    : pidUsers.filter((d, i) => pidUsers.findIndex((o) => String(o.dealercode).toLowerCase() === String(d.dealercode).toLowerCase()) === i);
  if (!dealers.length) throw new Error(`No PIDDealerCodes or PIDUsers in ${path.basename(CREDENTIALS_FILE)}. Add some or set PDI_DEALER.`);
  return dealers;
}

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

async function downloadInventoryStock(page: Page, code: string, log: Log): Promise<string> {
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
  const file = await viewer.saveReportExcel(reportPage, path.join(DOWNLOAD_DIR, `InventoryStockDetails_${code}_${stamp}.xlsx`));
  // The Excel is saved by now. Closing the viewer tab takes ~9s, so it closes in the background
  // while the dealers start (the main window stays open, so Chrome keeps running).
  viewer.closeReport(reportPage, 'Inventory Stock Details').catch(() => {});
  return file;
}

interface DealerWork { vins: string[]; upToDate: number }

// Every dealer's In-Transit work from the report, keyed by lower-case dealer code. The report lists every
// dealer's stock, so each row goes to the dealer in its Dealer Code column. Columns are found by header text.
//
// Only vehicles with a blank Engine No., Battery No or Vehicle ChargerNo in the report are taken (as in
// PDI.spec.ts). In-Transit vehicles stay In-Transit after an update, so this keeps hourly runs from
// re-saving vehicles that are already filled.
//
// Dealers run side by side, so a VIN must belong to exactly one dealer: a VIN listed under more than one
// Dealer Code is given to none of them and returned in `conflicts` instead (reported, never saved).
function readInTransit(file: string): { byDealer: Map<string, DealerWork>; conflicts: Map<string, string[]> } {
  const workbook = XLSX.readFile(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
  const table = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: '' });
  const norm = (s: unknown) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const headerIndex = table.findIndex((r) => r.some((c) => norm(c) === 'status') && r.some((c) => norm(c).startsWith('vin')));
  if (headerIndex === -1) throw new Error(`No "Status" / "VIN No." header row in ${path.basename(file)}.`);
  const headers = table[headerIndex]!.map(norm);
  const statusCol = headers.indexOf('status');
  const vinCol = headers.findIndex((h) => h.startsWith('vin'));
  const dealerCol = headers.indexOf('dealercode');
  if (dealerCol === -1) throw new Error(`No "Dealer Code" column in ${path.basename(file)}; can't tell whose vehicles are whose.`);
  // Report columns "Engine No.", "Battery No" and "Vehicle ChargerNo".
  const detailCols = [
    headers.findIndex((h) => h.startsWith('engine') || h.startsWith('motor')),
    headers.findIndex((h) => h.startsWith('battery')),
    headers.findIndex((h) => h.includes('charger')),
  ];
  if (detailCols.includes(-1)) throw new Error(`No Engine No. / Battery No / Vehicle ChargerNo column in ${path.basename(file)}.`);
  const rows = table.slice(headerIndex + 1)
    .filter((r) => norm(r[statusCol]) === 'intransit')
    .map((r) => ({ dealer: String(r[dealerCol]).trim().toLowerCase(), vin: String(r[vinCol]).trim().toUpperCase(), blank: detailCols.some((c) => !String(r[c] ?? '').trim()) }))
    .filter((r) => r.dealer && r.vin);

  const dealersOfVin = new Map<string, Set<string>>();
  rows.forEach((r) => dealersOfVin.set(r.vin, (dealersOfVin.get(r.vin) ?? new Set()).add(r.dealer)));
  const conflicts = new Map([...dealersOfVin].filter(([, d]) => d.size > 1).map(([vin, d]) => [vin, [...d]]));

  const byDealer = new Map<string, DealerWork>();
  for (const r of rows) {
    const work = byDealer.get(r.dealer) ?? { vins: [], upToDate: 0 };
    if (!r.blank) work.upToDate++;
    else if (!conflicts.has(r.vin) && !work.vins.includes(r.vin)) work.vins.push(r.vin);
    byDealer.set(r.dealer, work);
  }
  return { byDealer, conflicts };
}

// Logs out (or at least clears the session) so the next dealer's login starts clean.
async function resetSession(page: Page, username: string, log: Log) {
  try {
    await page.locator(`a:has-text("${username}")`).click({ timeout: 10000 });
    await page.getByRole('link', { name: /Sign out/i }).click({ timeout: 10000 });
  } catch {
    log('[STEP] Logout link not available, clearing the session instead.');
  }
  await page.context().clearCookies().catch(() => {});
}

// Step 1, before any parallel work: one login (the first dealer that works) downloads the report, which
// lists every dealer's stock. Nothing is saved in the DMS here.
async function downloadOnce(page: Page, dealers: UserCredentials[]): Promise<string> {
  let lastError: unknown;
  for (const dealer of dealers.slice(0, 3)) {
    const code = String(dealer.dealercode);
    const log: Log = (message) => console.log(`[report via ${code}] ${message}`);
    try {
      const loginPage = new LoginPage(page, log);
      await loginPage.goto();
      await loginPage.login(dealer);
      await loginPage.verifyDashboard();
      const file = await downloadInventoryStock(page, code, log);
      log(`[STEP] Inventory Stock Details saved: ${file}`);
      return file;
    } catch (err) {
      lastError = err;
      log(`[WARN] Could not download the report with this login: ${String((err as Error).message).split('\n')[0]!.slice(0, 150)}`);
    } finally {
      await resetSession(page, String(dealer.username), log);
    }
  }
  throw new Error(`Inventory Stock Details could not be downloaded: ${String((lastError as Error)?.message).split('\n')[0]}`);
}

// Step 2, one dealer in one lane: log in as the dealer and save its own VINs only.
async function updateDealer(page: Page, dealer: UserCredentials, vins: string[], vinDetails: ReturnType<typeof loadVinDetails>, log: Log): Promise<ResultRow[]> {
  const code = String(dealer.dealercode);
  const results: ResultRow[] = vins.map((chassisNo) => ({
    Dealer: code, ChassisNo: chassisNo, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: '',
  }));
  try {
    const loginPage = new LoginPage(page, log);
    await loginPage.goto();
    await loginPage.login(dealer);
    await loginPage.verifyDashboard();
    return await updateVehicles(page, loginPage, dealer, results, vinDetails, log,
      { dryRun: DRY_RUN, loadTimeoutMs: LOAD_TIMEOUT, recordBefore: 'Before' });
  } finally {
    await resetSession(page, String(dealer.username), log);
  }
}

test.describe('PDI In-Transit', () => {
  test.use({ storageState: { cookies: [], origins: [] } });
  // This writes to live DMS records, so it never runs as part of a plain `npm test`:
  // only via npm run test:pdi:intransit / test:pdi:intransit:dry, which set PDI_RUN=1.
  test.skip(process.env.PDI_RUN !== '1', 'PDI In-Transit update runs only via npm run test:pdi:intransit (sets PDI_RUN=1).');

  // All dealers in one test: real Chrome quits when its last window closes (see PDI.spec.ts). PARALLEL
  // dealers are updated at once, each lane in its own browser context (own cookies, so its own DMS session);
  // the first lane uses the test's own page, which keeps Chrome open throughout.
  test('Should fill In-Transit vehicle details from the VIN Details sheet', async ({ page, browser }) => {
    const dealers = pickDealers();
    const vinDetails = loadVinDetails(VIN_FILE);
    console.log(`[STEP] ${dealers.length} dealer(s) to check; ${vinDetails.size} VINs loaded from ${path.basename(VIN_FILE)}${DRY_RUN ? ' (DRY RUN, no Save)' : ''}`);

    // 1. One download, before any lane starts, so no two windows ever write the same file.
    test.setTimeout(15 * 60 * 1000);
    const stockFile = await downloadOnce(page, dealers);
    const { byDealer, conflicts } = readInTransit(stockFile);

    // 2. Work list: each dealer gets only the VINs under its own Dealer Code; a VIN under two dealers goes to
    //    no one (reported below). So no two lanes can ever open or save the same vehicle.
    //    A VIN that isn't in the VIN sheet has nothing to save, so it is reported here without logging in;
    //    a dealer is only logged into when at least one of its VINs is in the sheet.
    const results: ResultRow[] = [];
    const queue: { dealer: UserCredentials; vins: string[] }[] = [];
    let notInSheet = 0;
    for (const dealer of dealers) {
      const code = String(dealer.dealercode);
      const work = byDealer.get(code.toLowerCase());
      if (!work?.vins.length) continue;
      const missing = work.vins.filter((vin) => !vinDetails.has(vin));
      missing.forEach((vin) => results.push({ Dealer: code, ChassisNo: vin, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: 'Not in VIN Details sheet' }));
      notInSheet += missing.length;
      const toUpdate = work.vins.filter((vin) => vinDetails.has(vin));
      if (toUpdate.length) queue.push({ dealer, vins: toUpdate });
    }
    if (notInSheet) console.log(`[WARN] ${notInSheet} In-Transit vehicle(s) with blank details are not in the VIN Details sheet; reported, nothing to save.`);

    // In-Transit vehicles of a dealer that isn't a PDI dealer in credentials.json: never logged into (it isn't
    // approved for automation), but reported so they show under Needs attention instead of being missed.
    // Only on a full run: with PDI_DEALER set, the other dealers are left out on purpose.
    if (!process.env.PDI_DEALER) {
      const known = new Set(dealers.map((d) => String(d.dealercode).toLowerCase()));
      const unknown = [...byDealer].filter(([code, work]) => !known.has(code) && work.vins.length);
      for (const [code, work] of unknown) {
        work.vins.forEach((vin) => results.push({ Dealer: code.toUpperCase(), ChassisNo: vin, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '',
          Status: 'Error: dealer is not in the PDI dealer list (credentials.json), not updated' }));
      }
      if (unknown.length) console.log(`[WARN] In-Transit vehicles with blank details for dealer(s) not in ${path.basename(CREDENTIALS_FILE)}: ${unknown.map(([c, w]) => `${c.toUpperCase()} (${w.vins.length})`).join(', ')}; reported, not updated.`);
    }
    for (const [vin, owners] of conflicts) {
      console.log(`[WARN] ${vin} is In-Transit under more than one dealer (${owners.join(', ')}); not updated.`);
      results.push({ Dealer: owners.join('/'), ChassisNo: vin, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '',
        Status: `Error: listed under more than one dealer (${owners.join(', ')}) in Inventory Stock Details, not updated` });
    }
    const totalVins = queue.reduce((n, q) => n + q.vins.length, 0);
    const skipped = dealers.length - queue.length;
    console.log(`[STEP] ${totalVins} In-Transit vehicle(s) to update across ${queue.length} dealer(s)` +
      `${queue.length ? `: ${queue.map((q) => `${q.dealer.dealercode} (${q.vins.length})`).join(', ')}` : ''}; ${skipped} dealer(s) have nothing to update and are not logged into.`);

    // About 10s per vehicle (Save + re-check) plus logins, with headroom, spread over the lanes.
    const lanes = Math.min(PARALLEL, queue.length);
    test.setTimeout(15 * 60 * 1000 + Math.ceil((totalVins * 10 + queue.length * 60) / Math.max(1, lanes)) * 1000 * 2);

    // 3. Lanes: each takes the next dealer from the shared queue (shift() hands a dealer to exactly one lane)
    //    and only that dealer's VINs. Results are collected per dealer; files are written once, after all lanes.
    const failures: FailedDealer[] = [];
    let started = 0;
    const runLane = async (firstPage: Page) => {
      let lanePage = firstPage;
      for (let item = queue.shift(); item; item = queue.shift()) {
        const { dealer, vins } = item;
        const code = String(dealer.dealercode);
        const position = `${++started} of ${started + queue.length}`;
        // Seconds since the dealer started, so slow steps show up in the log.
        const dealerStart = Date.now();
        const log: Log = (message) => console.log(`[${code} +${Math.round((Date.now() - dealerStart) / 1000)}s] ${message}`);
        console.log(`\n[STEP] ===== Dealer ${code} (${position}): ${vins.length} In-Transit vehicle(s): ${vins.join(', ')} =====`);
        // If the tab was closed (but Chrome is still open), carry on in a new one in the same lane.
        if (lanePage.isClosed()) lanePage = await lanePage.context().newPage();
        try {
          const mine = await updateDealer(lanePage, dealer, vins, vinDetails, log);
          results.push(...mine);
          const done = mine.filter((r) => /^(Updated|Dry run)/.test(r.Status || '')).length;
          console.log(`[STEP] ===== Dealer ${code} done: ${mine.length} In-Transit, ${done} ${DRY_RUN ? 'filled' : 'updated'} =====`);
        } catch (err) {
          const message = String((err as Error).message).split('\n')[0]!.slice(0, 200);
          failures.push({ dealer: code, error: message });
          log(`[ERROR] Dealer failed: ${message}`);
          await test.info().attach(`${code}-failure.png`, {
            body: await lanePage.screenshot({ fullPage: true, timeout: 15000 }).catch(() => Buffer.from('')),
            contentType: 'image/png',
          }).catch(() => {});
        }
      }
    };
    if (lanes) console.log(`[STEP] Updating ${queue.length} dealer(s), ${lanes} at a time...`);
    const extraContexts = await Promise.all(Array.from({ length: Math.max(0, lanes - 1) }, () =>
      browser.newContext({ baseURL: test.info().project.use.baseURL ?? 'https://rivermobility.gaindms.com/', viewport: null })));
    try {
      await Promise.all([...(lanes ? [runLane(page)] : []), ...extraContexts.map(async (c) => runLane(await c.newPage()))]);
    } finally {
      await Promise.all(extraContexts.map((c) => c.close().catch(() => {})));
    }
    // Lanes finish in any order; report in the dealer list's order.
    const order = dealers.map((d) => String(d.dealercode));
    results.sort((a, b) => order.indexOf(a.Dealer!) - order.indexOf(b.Dealer!));

    const stamp = new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-');
    fs.mkdirSync(IN_TRANSIT_DIR, { recursive: true });
    const file = await writePdiWorkbook(path.join(IN_TRANSIT_DIR, `PDI_InTransit_${stamp}${DRY_RUN ? '_DRYRUN' : ''}.xlsx`), {
      title: 'PDI In-Transit vehicle details',
      subtitle: `${new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })} · ${dealers.length} dealer(s) · VIN sheet: ${path.basename(VIN_FILE)}`,
      dryRun: DRY_RUN, dealers: dealers.map((d) => String(d.dealercode)), failed: failures, rows: results, rowsSheet: 'In-Transit',
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

    const done = results.filter((r) => /^(Updated|Dry run)/.test(r.Status || ''));
    console.log(`[RESULT] ${results.length} In-Transit vehicle(s), ${done.length} ${DRY_RUN ? 'filled (dry run, nothing saved)' : 'updated'}`);
    results.filter((r) => !done.includes(r)).forEach((r) => console.log(`   NOT ${DRY_RUN ? 'filled' : 'updated'}: ${r.Dealer} ${r.ChassisNo}  ${r.Status}`));
    // As in PDI.spec.ts, data problems only warn (they are in the report; a rerun can't fix them, and failing
    // would turn every hourly run red until someone fixes the data): not in the VIN sheet, a Save the DMS
    // refuses (e.g. "Battery No must be unique"), a wrong login, and here also a dealer missing from
    // credentials.json or a VIN under two dealers. Timeouts and page errors still fail the run.
    const IN_TRANSIT_DATA_ISSUE = /not in the PDI dealer list|listed under more than one dealer/;
    const all = problemsOf(results.filter((r) => outcomeOf(r) !== 'missing'), failures, DRY_RUN);
    const itData = all.filter((p) => IN_TRANSIT_DATA_ISSUE.test(p));
    if (itData.length) console.log(`[WARN] ${itData.length} In-Transit vehicle(s) that need fixing in credentials.json or the DMS (listed in the report):\n  ${itData.join('\n  ')}`);
    const problems = splitProblems(all.filter((p) => !IN_TRANSIT_DATA_ISSUE.test(p)));
    if (problems.length) {
      const rerun = [...new Set(problems.map((p) => p.replace(/^Dealer /, '').split(/[\s:]/)[0]!))];
      throw new Error(`${problems.length} problem(s). Rerun with PDI_DEALER=${rerun.join(',')}\n  ${problems.slice(0, 20).join('\n  ')}`);
    }
  });
});
