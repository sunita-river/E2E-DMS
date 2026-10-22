import { test, Page } from '@playwright/test';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from '../tests/models/LoginPage';
import { writePdiDashboard, problemsOf, splitProblems, outcomeOf, FailedDealer, vinFilePath } from './utils/pdiReport';
import { writePdiWorkbook } from './utils/pdiWorkbook';
import { buildPdiDaily, localDate, RUNS_DIR } from './utils/pdiDaily';
import { loadVinDetails as loadVinSheet, updateVehicles, VinDetails, ResultRow, Log } from './utils/pdiVehicle';

// PDI blank-detail fix-up, for every dealer in PIDDealerCodes / PIDUsers (credentials.json), PDI_PARALLEL at a time:
//   1. WORKSHOP -> PDI: find grid rows with a blank field (usually EngineNo) and save their chassis numbers.
//   2. SETTINGS -> Admin -> Modify vehicle details: for each saved chassis, Show it, look it up in the
//      VIN Details sheet and fill Engine No (Motor No column), Battery and Charger, then Save.
//      The sheet's columns are found by header name, so their order doesn't matter.
//
// Settings (all optional):
//   PDI_DEALER    only these dealer codes, comma-separated, e.g. 332009,332008 (default: all PDI dealers)
//   PDI_VIN_FILE  VIN Details workbook or CSV (default: the newest "PDI Update"/VIN*.xlsx / VIN*.csv, e.g. the
//                 VIN_Details.csv that npm run pdi:vin fetches from the vin-data branch)
//   PDI_DRY_RUN   1 = fill the form but don't click Save (for a first check)
//   PDI_PARALLEL  dealers processed at once, each in its own browser window (default 3; 1 = one at a time)
//   PDI_LOAD_TIMEOUT_S  longest wait for Show / Save to reload the page (default 30)
//   PDI_SHOW_WAIT_S seconds to wait after clicking Show on the PDI Due List (default 30)
//   PDI_REFRESHES extra Show clicks when the PDI Due List says "Displaying : 0" (default 2)
//   PDI_ACCEPT_VIN_COLUMNS 1 = skip the swapped-columns check once and take this sheet as the new baseline
//   PDI_FIX_FILE  correction run: a CSV of "dealer,chassis" lines. Skips the PDI Due List and re-saves Engine
//                 No, Battery and Charger from the VIN sheet for exactly these chassis, even though they are no
//                 longer blank (their old values go in the "Was blank" column). Only the dealers in the file
//                 are run (PDI_DEALER narrows it further).
//
// A run sends no email. It saves its own result file in Output/PDI runs/ and refreshes the day's single
// copy, Output/PDI_Daily_<date>.xlsx + .html, plus Output/PDI_Master.xlsx + .html (Tracking sheet and the
// last 3 days) and the all-time history Output/PDI_History.json; npm run report:pdi-daily emails it once a day.
const CREDENTIALS_FILE = process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
const VIN_FILE = vinFilePath();
const DRY_RUN = process.env.PDI_DRY_RUN === '1';
const PDI_PARALLEL = Number(process.env.PDI_PARALLEL || 3);
if (!Number.isInteger(PDI_PARALLEL) || PDI_PARALLEL < 1) throw new Error(`PDI_PARALLEL must be a whole number above 0 (got "${process.env.PDI_PARALLEL}")`);
const SHOW_WAIT_MS = Number(process.env.PDI_SHOW_WAIT_S || 30) * 1000;
const PDI_REFRESHES =Number(process.env.PDI_REFRESHES || 2);
const LOAD_TIMEOUT =Number(process.env.PDI_LOAD_TIMEOUT_S || 30) * 1000;
const loadVinDetails = () => loadVinSheet(VIN_FILE);

// Correction run (PDI_FIX_FILE): dealer code -> chassis numbers to re-save. Lines that aren't
// "<dealer>,<17-character chassis>" (a header, blank lines) are skipped.
const FIX_FILE = process.env.PDI_FIX_FILE ? path.resolve(process.env.PDI_FIX_FILE) : '';
function loadFixList(): Map<string, string[]> | null {
  if (!FIX_FILE) return null;
  if (!fs.existsSync(FIX_FILE)) throw new Error(`PDI_FIX_FILE not found: ${FIX_FILE}`);
  const byDealer = new Map<string, string[]>();
  for (const line of fs.readFileSync(FIX_FILE, 'utf8').split(/\r?\n/)) {
    const [dealer, chassis] = line.split(/[,;\t]/).map((s) => s.trim().replace(/^"|"$/g, ''));
    if (!dealer || !chassis || !/^[A-Z0-9]{17}$/i.test(chassis)) continue;
    const list = byDealer.get(dealer) || [];
    if (!list.includes(chassis.toUpperCase())) list.push(chassis.toUpperCase());
    byDealer.set(dealer, list);
  }
  if (!byDealer.size) throw new Error(`PDI_FIX_FILE ${path.basename(FIX_FILE)} has no "dealer,chassis" lines.`);
  return byDealer;
}
const FIX_LIST = loadFixList();

const PDI_URL = 'Product/PDIStockCheckingMain.aspx';

// Every PDI dealer, or just the PDI_DEALER codes (which may also come from dealerUsers /
// dealervalidUser). PDI dealers are the PIDUsers entries (own login) plus each PIDDealerCodes code
// signed in with the shared PIDLogin; a code in both uses its PIDUsers login.
// Fails up front if a dealer is unknown or still has placeholder credentials.
function pickDealers(): UserCredentials[] {
  const ownLogin: UserCredentials[] = credentials.PIDUsers || [];
  const listed = new Set(ownLogin.map((d) => String(d.dealercode).toLowerCase()));
  const codes: string[] = (credentials.PIDDealerCodes || []).map((c: unknown) => String(c).trim()).filter(Boolean);
  if (codes.length && !credentials.PIDLogin) throw new Error(`PIDDealerCodes needs a PIDLogin (username, password) in ${path.basename(CREDENTIALS_FILE)}.`);
  const pidUsers: UserCredentials[] = [
    ...ownLogin,
    ...codes.filter((c) => !listed.has(c.toLowerCase())).map((dealercode) => ({ ...credentials.PIDLogin, dealercode })),
  ];
  let wanted = (process.env.PDI_DEALER || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean);
  // A correction run covers only the dealers in its fix file (narrowed by PDI_DEALER when set).
  if (FIX_LIST) {
    const inFile = [...FIX_LIST.keys()].map((c) => c.toLowerCase());
    wanted = wanted.length ? inFile.filter((c) => wanted.includes(c)) : inFile;
    if (!wanted.length) throw new Error(`None of the PDI_DEALER codes are in ${path.basename(FIX_FILE)}.`);
  }
  const all: UserCredentials[] = [...pidUsers, credentials.dealervalidUser, ...(credentials.dealerUsers || [])];
  let dealers = wanted.length
    ? wanted.map((code) => {
        const match = all.find((d) => String(d.dealercode).toLowerCase() === code);
        if (!match) throw new Error(`PDI_DEALER: dealer ${code} is not in ${path.basename(CREDENTIALS_FILE)}.`);
        return match;
      })
    : pidUsers;
  // A dealer listed twice is only run once.
  const seen = new Set<string>();
  dealers = dealers.filter((d) => {
    const code = String(d.dealercode).toLowerCase();
    if (seen.has(code)) { console.log(`[WARN] Dealer ${d.dealercode} is listed more than once, running it once.`); return false; }
    seen.add(code);
    return true;
  });
  if (!dealers.length) throw new Error(`No PIDDealerCodes or PIDUsers in ${path.basename(CREDENTIALS_FILE)}. Add some or set PDI_DEALER.`);
  const incomplete = dealers.filter((d) => !d.username || !d.password || /^<.*>$/.test(String(d.username)) || /^<.*>$/.test(String(d.password)));
  if (incomplete.length) {
    throw new Error(`Username/password missing or still <user>/<password> in ${path.basename(CREDENTIALS_FILE)} for: ${incomplete.map((d) => d.dealercode).join(', ')}`);
  }
  return dealers;
}


interface BlankRow { chassisNo: string; blankFields: string[] }

// Grid inputs are named ctl00_cpMain_grdMain_ctlNN_<Field>; group them by row (ctlNN) and keep
// the chassis number of every row where any field is empty.
async function readBlankPdiRows(page: Page): Promise<BlankRow[]> {
  const read = () => page.locator("input[id^='ctl00_cpMain_grdMain_']").evaluateAll((els) =>
    els.map((e) => ({ id: e.id, value: ((e as HTMLInputElement).value || '').trim() })));
  // If the page is still navigating, wait for it to settle and read again (up to 3 tries).
  let cells: { id: string; value: string }[] = [];
  for (let attempt = 1; ; attempt++) {
    try { cells = await read(); break; } catch (err) {
      if (attempt >= 3 || !/context was destroyed|navigat/i.test(String((err as Error).message))) throw err;
      await page.waitForLoadState('load').catch(() => {});
      await page.waitForTimeout(2000);
    }
  }
  const rows = new Map<string, Record<string, string>>();
  for (const { id, value } of cells) {
    const m = id.match(/^ctl00_cpMain_grdMain_(ctl\d+)_(\w+)$/);
    if (!m) continue;
    rows.set(m[1]!, { ...rows.get(m[1]!), [m[2]!]: value });
  }
  const blank: BlankRow[] = [];
  for (const fields of rows.values()) {
    const blankFields = Object.keys(fields).filter((f) => fields[f] === '');
    if (blankFields.length && fields.ChassisNo) blank.push({ chassisNo: fields.ChassisNo, blankFields });
  }
  return blank;
}

// One workbook for all dealers: Output/PDI runs/PDI_BlankChassis_<date-time>.xlsx. The "Dealers" sheet
// lists every dealer checked, even with no blank rows, so the day's merge (pdiDaily.ts) can clear
// earlier errors that a later run fixed.
async function saveResults(results: ResultRow[], dealersChecked: ResultRow[], dealers: string[], failures: FailedDealer[]): Promise<string> {
  const stamp = new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-');
  fs.mkdirSync(RUNS_DIR, { recursive: true });
  return writePdiWorkbook(path.join(RUNS_DIR, `PDI_BlankChassis_${stamp}.xlsx`), {
    title: 'PDI vehicle details',
    subtitle: `${new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })} · ${dealers.length} dealer(s) · VIN sheet: ${path.basename(VIN_FILE)}${FIX_FILE ? ` · correction run (${path.basename(FIX_FILE)})` : ''}`,
    dryRun: DRY_RUN, dealers, failed: failures, rows: results, rowsSheet: 'PDI Blank Chassis', dealersSheet: dealersChecked,
  });
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

// WORKSHOP -> PDI: list every vehicle on the PDI Due List and return the rows with a blank field.
async function findBlankPdiRows(page: Page, log: Log): Promise<BlankRow[]> {
  log('[STEP] Opening WORKSHOP -> PDI...');
  // Menu link names carry an icon glyph (e.g. " PDI"), so match loosely / by href.
  await page.getByRole('link', { name: /^\W*WORKSHOP\W*$/ }).click();
  await page.locator('a[href$="PDIStockCheckingMain.aspx"]').first().click();
  await page.waitForURL(new RegExp(PDI_URL.replace('.', '\\.'), 'i'), { timeout: 60000 })
    .catch(() => page.goto(PDI_URL)); // menu hidden/collapsed for some users, so go direct
  const chassisCells = page.locator("input[id^='ctl00_cpMain_grdMain_'][id$='_ChassisNo']");

  // "Displaying : N" is the PDI Due List's own count; it appears once the list has loaded. A count of 0,
  // or no count yet, can mean the records haven't come through, so click Show (the page's refresh;
  // Top is left at its default) and read it again, up to PDI_REFRESHES times, before trusting it.
  const displayedCount = async (timeout: number) => {
    const text = await page.getByText(/Displaying\s*:\s*\d+/i).first().textContent({ timeout }).catch(() => null);
    const m = text?.match(/Displaying\s*:\s*(\d+)/i);
    return m ? Number(m[1]) : null;
  };
  const refreshList = async () => {
    // Tag the current page so the old list isn't mistaken for the refreshed one.
    await page.locator("#ctl00_cpMain_Topn, input[id^='ctl00_cpMain_grdMain_']")
      .evaluateAll((els) => els.forEach((e) => e.setAttribute('data-pw-stale', '1'))).catch(() => {});
    // Click the Show button's label; pressing Enter in Top posts the form without a command.
    await page.locator('#ctl00_cpMain_cmdShow > span').click();
    await page.waitForTimeout(SHOW_WAIT_MS);
    await page.waitForLoadState('load').catch(() => {});
    const reloaded = await page.locator('#ctl00_cpMain_Topn:not([data-pw-stale])').count().catch(() => 0);
    if (!reloaded) log(`[WARN] PDI Due List had not reloaded ${SHOW_WAIT_MS / 1000}s after Show, reading what is on screen.`);
  };
  let displayed = await displayedCount(60000);
  for (let i = 1; i <= PDI_REFRESHES && !displayed; i++) {
    log(`[STEP] PDI Due List shows ${displayed === null ? 'no "Displaying" count' : '"Displaying : 0"'}, refreshing to fetch the records (${i} of ${PDI_REFRESHES})...`);
    await refreshList();
    displayed = await displayedCount(30000);
  }
  if (displayed) await chassisCells.first().waitFor({ timeout: 30000 }).catch(() => {});

  // The list shows 25 vehicles by default. When "Displaying" is higher, put that exact count in Top
  // (never a fixed 500) and Show again so every vehicle is checked; otherwise Top is left alone.
  if (displayed && displayed > (await chassisCells.count())) {
    log(`[STEP] PDI Due List says Displaying : ${displayed}, setting Top to ${displayed} and showing all of them...`);
    await page.locator('#ctl00_cpMain_Topn').fill(String(displayed));
    await refreshList();
    await chassisCells.first().waitFor({ timeout: 30000 }).catch(() => {});
    displayed = (await displayedCount(30000)) ?? displayed;
  }

  // Old fixed "Top = 500" step, kept for reference. Top is now set from the "Displaying" count above.
  // // The page lists only the top 25 vehicles by default; raise "Top" and click Show so every
  // // vehicle waiting for PDI is checked. (Show is the page's own refresh; press no Enter here,
  // // it posts the form without a command and empties the grid.)
  // await page.locator('#ctl00_cpMain_Topn').fill(String(PDI_TOP));
  // // Tag the current page so the Top-25 grid still on screen isn't mistaken for the reloaded one
  // // (reading it mid-reload fails with "Execution context was destroyed").
  // await page.locator("#ctl00_cpMain_Topn, input[id^='ctl00_cpMain_grdMain_'][id$='_ChassisNo']")
  //   .evaluateAll((els) => els.forEach((e) => e.setAttribute('data-pw-stale', '1'))).catch(() => {});
  // await page.locator('#ctl00_cpMain_cmdShow').click();
  // const reloaded = await page.waitForFunction(() => {
  //   if (document.readyState !== 'complete') return false;
  //   const top = document.querySelector('#ctl00_cpMain_Topn');
  //   const cells = [...document.querySelectorAll("input[id^='ctl00_cpMain_grdMain_'][id$='_ChassisNo']")];
  //   return (!!top && !top.hasAttribute('data-pw-stale')) || (cells.length > 0 && cells.every((c) => !c.hasAttribute('data-pw-stale')));
  // }, undefined, { timeout: 90000 }).then(() => true, () => false);
  // if (!reloaded) log('[WARN] PDI page did not reload within 90s after Show, reading what is on screen.');
  // await page.waitForLoadState('load').catch(() => {});
  // await chassisCells.first().waitFor({ timeout: 60000 }).catch(() => log('[STEP] PDI grid is empty.'));
  const gridRows = await chassisCells.count();
  log(`[STEP] PDI Due List: Displaying : ${displayed ?? '?'}, ${gridRows} vehicle(s) in the grid`);
  if (displayed === null) log('[WARN] No "Displaying" count on the PDI Due List page, reading what is on screen.');
  else if (gridRows < displayed) log(`[WARN] Page says Displaying : ${displayed} but only ${gridRows} rows are in the grid; some vehicles may not be checked.`);
  if (displayed === null && gridRows >= 25) log('[WARN] The grid is full at 25 and the page shows no "Displaying" count; there may be more vehicles that were not checked.');

  return readBlankPdiRows(page);
}

async function processDealer(page: Page, dealer: UserCredentials, vinDetails: Map<string, VinDetails>, log: Log): Promise<ResultRow[]> {
  const loginPage = new LoginPage(page, log);
  await loginPage.goto();
  await loginPage.login(dealer);
  await loginPage.verifyDashboard();

  // A correction run (PDI_FIX_FILE) skips the PDI Due List and re-saves exactly the listed chassis.
  const fixChassis = FIX_LIST?.get(String(dealer.dealercode));
  const blankRows: BlankRow[] = fixChassis
    ? fixChassis.map((chassisNo) => ({ chassisNo, blankFields: [] }))
    : await findBlankPdiRows(page, log);
  if (fixChassis) {
    log(`[STEP] Correction run: ${fixChassis.length} chassis from ${path.basename(FIX_FILE)}: ${fixChassis.join(', ')}`);
  } else {
    log(`[STEP] ${blankRows.length} PDI row(s) with blank fields:`);
    blankRows.forEach((r) => log(`   ${r.chassisNo}  (blank: ${r.blankFields.join(', ')})`));
  }

  const results: ResultRow[] = blankRows.map((r) => ({
    Dealer: String(dealer.dealercode), ChassisNo: r.chassisNo, BlankFields: r.blankFields.join(', '),
    EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: '',
  }));
  if (!blankRows.length) {
    log('[STEP] Nothing to update.');
    return results;
  }

  return updateVehicles(page, loginPage, dealer, results, vinDetails, log,
    { dryRun: DRY_RUN, loadTimeoutMs: LOAD_TIMEOUT, recordBefore: fixChassis ? 'Correction, before' : undefined });
}

test.describe('PDI', () => {
  test.use({ storageState: { cookies: [], origins: [] } });
  // This writes to live DMS records, so it never runs as part of a plain `npm test`:
  // only via npm run test:pdi / test:pdi:dry, which set PDI_RUN=1.
  test.skip(process.env.PDI_RUN !== '1', 'PDI update runs only via npm run test:pdi (sets PDI_RUN=1).');

  // All dealers run inside one test: real Chrome quits when its last window closes, so one test()
  // per dealer would kill the browser between dealers (see GSTR Report.spec.ts). PDI_PARALLEL
  // dealers run at once, each lane in its own browser context (own cookies, so its own DMS session);
  // the first lane uses the test's own page, which keeps Chrome open throughout.
  test('Should fill blank PDI vehicle details from the VIN Details sheet', async ({ page, browser }) => {
    const dealers = pickDealers();
    const vinDetails = loadVinDetails();
    const lanes = Math.min(PDI_PARALLEL, dealers.length);
    console.log(`[STEP] ${dealers.length} dealer(s), ${lanes} at a time: ${dealers.map((d) => d.dealercode).join(', ')}; ${vinDetails.size} VINs loaded from ${path.basename(VIN_FILE)}${DRY_RUN ? ' (DRY RUN, no Save)' : ''}${FIX_FILE ? `, CORRECTION RUN from ${path.basename(FIX_FILE)}` : ''}`);

    // About 25 blank rows x 25s each per dealer, with headroom.
    test.setTimeout(Math.ceil(dealers.length / lanes) * 20 * 60 * 1000);

    const startedAt = new Date();
    const results: ResultRow[] = [];
    const failedDealers: string[] = [];
    const failures: FailedDealer[] = [];
    const queue = [...dealers];
    const runLane = async (lanePage: Page) => {
      for (let dealer = queue.shift(); dealer; dealer = queue.shift()) {
        const code = String(dealer.dealercode);
        const log: Log = (message) => console.log(`[${code}] ${message}`);
        const position = `${dealers.length - queue.length} of ${dealers.length}`;
        console.log(`\n[STEP] ===== Dealer ${code} (${position}) started =====`);
        try {
          const mine = await processDealer(lanePage, dealer, vinDetails, log);
          results.push(...mine);
          const done = mine.filter((r) => /^(Updated|Dry run)/.test(r.Status || '')).length;
          console.log(`[STEP] ===== Dealer ${code} (${position}) done: ${mine.length} blank, ${done} ${DRY_RUN ? 'filled' : 'updated'} =====`);
        } catch (err) {
          // Keep Playwright's "waiting for <locator>" line so a timeout says which element it was.
          const lines = String((err as Error).message).split('\n');
          const waitingFor = lines.find((l) => /^\s*- waiting for /.test(l))?.trim().replace(/^- /, '');
          const message = (lines[0]! + (waitingFor ? ` (${waitingFor})` : '')).slice(0, 300);
          failedDealers.push(code);
          failures.push({ dealer: code, error: message });
          log(`[ERROR] Dealer failed: ${message}`);
          console.log(`[STEP] ===== Dealer ${code} (${position}) FAILED =====`);
          await test.info().attach(`${code}-failure.png`, {
            body: await lanePage.screenshot({ fullPage: true, timeout: 15000 }).catch(() => Buffer.from('')),
            contentType: 'image/png',
          }).catch(() => {});
        } finally {
          await resetSession(lanePage, String(dealer.username), log);
        }
      }
    };
    const extraContexts = await Promise.all(Array.from({ length: lanes - 1 }, () =>
      browser.newContext({ baseURL: test.info().project.use.baseURL ?? 'https://rivermobility.gaindms.com/', viewport: null })));
    try {
      await Promise.all([runLane(page), ...extraContexts.map(async (c) => runLane(await c.newPage()))]);
    } finally {
      await Promise.all(extraContexts.map((c) => c.close().catch(() => {})));
    }
    // Lanes finish in any order; report in the dealer list's order.
    const order = dealers.map((d) => String(d.dealercode));
    results.sort((a, b) => order.indexOf(a.Dealer!) - order.indexOf(b.Dealer!));

    const dealersChecked = dealers.map((d) => {
      const code = String(d.dealercode);
      const failure = failures.find((f) => f.dealer === code);
      return { Dealer: code, DryRun: DRY_RUN ? 'Yes' : 'No', BlankFound: String(results.filter((r) => r.Dealer === code).length), Result: failure ? `Failed: ${failure.error}` : 'Checked' };
    });
    const file = await saveResults(results, dealersChecked, order, failures);
    console.log(`\n[STEP] Results saved to ${file}`);

    const clean: string[] = [];
    for (const code of order) {
      if (failedDealers.includes(code)) {
        console.log(`\n[RESULT] Dealer ${code}: FAILED before finishing, see the error above.`);
        continue;
      }
      const mine = results.filter((r) => r.Dealer === code);
      if (!mine.length) { clean.push(code); continue; }
      const done = (r: ResultRow) => r.Status!.startsWith(DRY_RUN ? 'Dry run' : 'Updated');
      const updated = mine.filter(done);
      console.log(`\n[RESULT] Dealer ${code}: ${updated.length} of ${mine.length} chassis ${DRY_RUN ? 'filled (dry run, nothing saved)' : 'updated'}`);
      updated.forEach((r) => console.log(`   ${r.ChassisNo}  Engine ${r.EngineNo}, Battery ${r.BatteryDetails}, Charger ${r.VehicleChargerNo}`));
      const notUpdated = mine.filter((r) => !done(r));
      if (notUpdated.length) {
        console.log(`   NOT ${DRY_RUN ? 'filled' : 'updated'}:`);
        notUpdated.forEach((r) => console.log(`   ${r.ChassisNo}  ${r.Status}`));
      }
    }
    if (clean.length) console.log(`\n[RESULT] No blank PDI rows (${clean.length} dealers): ${clean.join(', ')}`);

    // No email here. A real run refreshes the day's single copy (every run so far today, latest result
    // per chassis); a dry run saved nothing, so it only gets its own dashboard next to its run file.
    if (DRY_RUN) {
      writePdiDashboard(results, failures, { dryRun: true, startedAt, finishedAt: new Date(), dealers: order, vinFile: VIN_FILE }, file);
    } else {
      try {
        const daily = await buildPdiDaily(localDate(startedAt));
        if (daily) writePdiDashboard(daily.results, daily.failures, daily.meta, daily.file);
      } catch (err) {
        console.log(`[WARN] Day's PDI copy not refreshed: ${(err as Error).message}`);
      }
      console.log('[STEP] No email sent. Run npm run report:pdi-daily after the day\'s last PDI run to email the day\'s report once.');
    }

    // A chassis missing from the VIN Details sheet is a data gap, not a test failure; only real errors fail.
    const noData = results.filter((r) => outcomeOf(r) === 'missing');
    if (noData.length) {
      const dealersMissing = [...new Set(noData.map((r) => r.Dealer!))];
      console.log(`\n[RESULT] No data found in the VIN Details sheet (${path.basename(VIN_FILE)}) for ${noData.length} chassis, not updated. ` +
        `Add these VINs to the sheet, then rerun with PDI_DEALER=${dealersMissing.join(',')}`);
      noData.forEach((r) => console.log(`   ${r.Dealer} ${r.ChassisNo}: ${r.Status}`));
    }
    const problems = splitProblems(problemsOf(results.filter((r) => outcomeOf(r) !== 'missing'), failures, DRY_RUN));
    if (problems.length) {
      const rerun = [...new Set([...failedDealers, ...results.filter((r) => problems.some((p) => p.startsWith(`${r.Dealer} ${r.ChassisNo}:`))).map((r) => r.Dealer!)])];
      throw new Error(`${problems.length} problem(s). Rerun with PDI_DEALER=${rerun.join(',')}\n  ${problems.slice(0, 20).join('\n  ')}`);
    }
  });
});
