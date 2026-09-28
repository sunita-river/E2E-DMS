import { test, Page, Dialog } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from './models/LoginPage';
import { sendPdiReport, problemsOf, FailedDealer } from './utils/pdiReport';

// PDI blank-detail fix-up, for every dealer in PIDUsers (credentials.json), one after another:
//   1. WORKSHOP -> PDI: find grid rows with a blank field (usually EngineNo) and save their chassis numbers.
//   2. SETTINGS -> Admin -> Modify vehicle details: for each saved chassis, Show it, look it up in the
//      VIN Details sheet and fill Engine No (col C), Battery (col B) and Charger (col D), then Save.
//
// Settings (all optional):
//   PDI_DEALER    only these dealer codes, comma-separated, e.g. 332009,332008 (default: all PIDUsers)
//   PDI_VIN_FILE  VIN Details workbook (default: resources/VIN Details (1).xlsx)
//   PDI_DRY_RUN   1 = fill the form but don't click Save (for a first check)
//   PDI_TOP       vehicles to list on the PDI page (default 500; the page itself shows 25)
//   PDI_SEND_EMAIL 0 = don't email the per-dealer report (it uses the GMAIL_* settings in .env)
const CREDENTIALS_FILE = process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
const VIN_FILE = process.env.PDI_VIN_FILE || path.join(__dirname, '..', 'resources', 'VIN Details (1).xlsx');
const OUTPUT_DIR = process.env.GSTR_OUTPUT_DIR || path.join(process.cwd(), 'Output');
const DRY_RUN = process.env.PDI_DRY_RUN === '1';
const EDIT_REASON = 'Details updated';
const PDI_TOP = Number(process.env.PDI_TOP || 500); // rows to list on the PDI page (the page's own default is 25)
if (!Number.isInteger(PDI_TOP) || PDI_TOP < 1) throw new Error(`PDI_TOP must be a whole number above 0 (got "${process.env.PDI_TOP}")`);

const PDI_URL = 'Product/PDIStockCheckingMain.aspx';
const MODIFY_VEHICLE_URL = 'Common/ChangeChassisNo.aspx';

type ResultRow = Record<string, string>;

// Every PIDUsers dealer, or just the PDI_DEALER codes (which may also come from dealerUsers /
// dealervalidUser). Fails up front if a dealer is unknown or still has placeholder credentials.
function pickDealers(): UserCredentials[] {
  const pidUsers: UserCredentials[] = credentials.PIDUsers || [];
  const wanted = (process.env.PDI_DEALER || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean);
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
    if (seen.has(code)) { console.log(`[WARN] Dealer ${d.dealercode} is listed more than once — running it once.`); return false; }
    seen.add(code);
    return true;
  });
  if (!dealers.length) throw new Error(`No PIDUsers in ${path.basename(CREDENTIALS_FILE)} — add some or set PDI_DEALER.`);
  const incomplete = dealers.filter((d) => !d.username || !d.password || /^<.*>$/.test(String(d.username)) || /^<.*>$/.test(String(d.password)));
  if (incomplete.length) {
    throw new Error(`Username/password missing or still <user>/<password> in ${path.basename(CREDENTIALS_FILE)} for: ${incomplete.map((d) => d.dealercode).join(', ')}`);
  }
  return dealers;
}

interface VinDetails { battery: string; motor: string; charger: string }

// VIN Details sheet: A = VIN, B = Battery No, C = Motor No, D = CHARGER (header in row 1).
function loadVinDetails(): Map<string, VinDetails> {
  if (!fs.existsSync(VIN_FILE)) throw new Error(`VIN Details workbook not found: ${VIN_FILE} (set PDI_VIN_FILE).`);
  const workbook = XLSX.readFile(VIN_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
  const rows = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: '' });
  const byVin = new Map<string, VinDetails>();
  for (const [vin, battery, motor, charger] of rows.slice(1)) {
    const key = String(vin || '').trim().toUpperCase();
    if (key) byVin.set(key, { battery: String(battery).trim(), motor: String(motor).trim(), charger: String(charger).trim() });
  }
  return byVin;
}

interface BlankRow { chassisNo: string; blankFields: string[] }

// Grid inputs are named ctl00_cpMain_grdMain_ctlNN_<Field>; group them by row (ctlNN) and keep
// the chassis number of every row where any field is empty.
async function readBlankPdiRows(page: Page): Promise<BlankRow[]> {
  const cells = await page.locator("input[id^='ctl00_cpMain_grdMain_']").evaluateAll((els) =>
    els.map((e) => ({ id: e.id, value: ((e as HTMLInputElement).value || '').trim() })));
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

// All dealers' rows go into one workbook: Output/PDI_BlankChassis_<date-time>.xlsx. The "Dealers"
// sheet lists every dealer checked (including ones with no blank rows) so the daily report
// (tests/PDIDailyReport.spec.ts) can show them and spot earlier errors that a later run cleared.
function saveResults(results: ResultRow[], dealersChecked: ResultRow[]): string {
  const stamp = new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-');
  const filePath = path.join(OUTPUT_DIR, `PDI_BlankChassis_${stamp}.xlsx`);
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(results), 'PDI Blank Chassis');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dealersChecked), 'Dealers');
  try {
    XLSX.writeFile(workbook, filePath);
  } catch (err) {
    // Usually the file is open in Excel (Windows locks it). Save beside it rather than lose the run's results.
    const fallback = filePath.replace(/\.xlsx$/, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${String((err as Error).message).slice(0, 80)}) — saved as ${path.basename(fallback)} instead.`);
    XLSX.writeFile(workbook, fallback);
    return fallback;
  }
  return filePath;
}

// Logs out (or at least clears the session) so the next dealer's login starts clean.
async function resetSession(page: Page, username: string) {
  try {
    await page.locator(`a:has-text("${username}")`).click({ timeout: 10000 });
    await page.getByRole('link', { name: /Sign out/i }).click({ timeout: 10000 });
  } catch {
    console.log('[STEP] Logout link not available — clearing the session instead.');
  }
  await page.context().clearCookies().catch(() => {});
}

async function processDealer(page: Page, dealer: UserCredentials, vinDetails: Map<string, VinDetails>): Promise<ResultRow[]> {
  const loginPage = new LoginPage(page);
  await loginPage.goto();
  await loginPage.login(dealer);
  await loginPage.verifyDashboard();

  // WORKSHOP -> PDI
  console.log('[STEP] Opening WORKSHOP -> PDI...');
  // Menu link names carry an icon glyph (e.g. " PDI"), so match loosely / by href.
  await page.getByRole('link', { name: /^\W*WORKSHOP\W*$/ }).click();
  await page.locator('a[href$="PDIStockCheckingMain.aspx"]').first().click();
  await page.waitForURL(new RegExp(PDI_URL.replace('.', '\\.'), 'i'), { timeout: 60000 })
    .catch(() => page.goto(PDI_URL)); // menu hidden/collapsed for some users — go direct
  const chassisCells = page.locator("input[id^='ctl00_cpMain_grdMain_'][id$='_ChassisNo']");
  await chassisCells.first().waitFor({ timeout: 60000 }).catch(() => {});

  // The page lists only the top 25 vehicles by default; raise "Top" and click Show so every
  // vehicle waiting for PDI is checked. (Show is the page's own refresh; press no Enter here —
  // it posts the form without a command and empties the grid.)
  await page.locator('#ctl00_cpMain_Topn').fill(String(PDI_TOP));
  await page.locator('#ctl00_cpMain_cmdShow').click();
  await page.waitForLoadState('domcontentloaded');
  await chassisCells.first().waitFor({ timeout: 60000 }).catch(() => console.log('[STEP] PDI grid is empty.'));
  const gridRows = await chassisCells.count();
  console.log(`[STEP] PDI grid: ${gridRows} vehicle(s) (Top ${PDI_TOP})`);
  if (gridRows >= PDI_TOP) console.log(`[WARN] The grid is full at Top ${PDI_TOP} — there may be more vehicles; raise PDI_TOP.`);

  const blankRows = await readBlankPdiRows(page);
  console.log(`[STEP] ${blankRows.length} PDI row(s) with blank fields:`);
  blankRows.forEach((r) => console.log(`   ${r.chassisNo}  (blank: ${r.blankFields.join(', ')})`));

  const results: ResultRow[] = blankRows.map((r) => ({
    Dealer: String(dealer.dealercode), ChassisNo: r.chassisNo, BlankFields: r.blankFields.join(', '),
    EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: '',
  }));
  if (!blankRows.length) {
    console.log('[STEP] Nothing to update.');
    return results;
  }

  // SETTINGS -> Admin -> Modify vehicle details
  console.log('[STEP] Opening SETTINGS -> Admin -> Modify vehicle details...');
  await page.getByRole('link', { name: /^\W*SETTINGS\W*$/ }).click();
  await page.getByRole('link', { name: /^\W*Admin\W*$/ }).first().click();
  await page.locator('a[href$="ChangeChassisNo.aspx"]').first().click();
  await page.waitForURL(/ChangeChassisNo\.aspx/i, { timeout: 60000 }).catch(() => page.goto(MODIFY_VEHICLE_URL));

  // Save may raise a confirm()/alert(); accept it and keep the text for the result sheet.
  let lastDialog = '';
  const onDialog = async (d: Dialog) => { lastDialog = d.message(); await d.accept().catch(() => {}); };
  page.on('dialog', onDialog);

  const showVehicle = async (chassis: string) => {
    await page.locator('[name="ctl00$cpMain$IdentifierValue"]').fill(chassis);
    await page.locator('#ctl00_cpMain_cmdShow').click();
    await page.waitForTimeout(10000);
  };
  const fieldValue = async (name: string) => (await page.locator(`[name="ctl00$cpMain$${name}"]`).inputValue()).trim();
  const identifierInput = page.locator('[name="ctl00$cpMain$IdentifierValue"]');

  // After an error, get back to a usable Modify vehicle details page. The DMS sometimes drops the
  // session mid-dealer (a click times out, then every later field is missing), so reloading the page
  // alone isn't enough — log in again when the form doesn't come back.
  const recover = async () => {
    await page.goto(MODIFY_VEHICLE_URL).catch(() => {});
    if (await identifierInput.isVisible({ timeout: 15000 }).catch(() => false)) return;
    console.log('[STEP] Modify vehicle details not available (session lost?) — logging in again...');
    await loginPage.goto();
    await loginPage.login(dealer);
    await loginPage.verifyDashboard();
    await page.goto(MODIFY_VEHICLE_URL);
    await identifierInput.waitFor({ timeout: 60000 });
  };

  // One attempt at a chassis; returns its Status. Throws on a page error (timeouts etc.).
  const updateChassis = async (chassis: string, details: VinDetails): Promise<string> => {
    await showVehicle(chassis);

    const loaded = await fieldValue('ChassisNo');
    if (loaded.toUpperCase() !== chassis.toUpperCase()) return `Show did not load this vehicle (got "${loaded}")`;

    await page.locator('[name="ctl00$cpMain$EngineNo"]').fill(details.motor);
    await page.locator('[name="ctl00$cpMain$BatteryDetails"]').fill(details.battery);
    await page.locator('[name="ctl00$cpMain$VehicleChargerNo"]').fill(details.charger);
    await page.locator('[name="ctl00$cpMain$EditChassisReason"]').fill(EDIT_REASON);
    if (DRY_RUN) return 'Dry run — filled, not saved';

    lastDialog = '';
    await page.locator('#ctl00_cpMain_cmdSave').click();
    await page.waitForLoadState('domcontentloaded');
    await page.waitForTimeout(3000);
    const saveMessage = lastDialog;

    // Re-open the vehicle to confirm the DMS actually kept the new values.
    await showVehicle(chassis);
    const kept = (await fieldValue('EngineNo')) === details.motor
      && (await fieldValue('BatteryDetails')) === details.battery
      && (await fieldValue('VehicleChargerNo')) === details.charger;
    return (kept ? 'Updated' : 'Save clicked but values not kept') + (saveMessage ? ` (${saveMessage})` : '');
  };

  try {
    for (const row of results) {
      const chassis = row.ChassisNo!;
      const details = vinDetails.get(chassis.toUpperCase());
      if (!details) {
        row.Status = 'Not in VIN Details sheet';
        console.log(`[WARN] ${chassis}: not found in VIN Details — skipped.`);
        continue;
      }
      row.EngineNo = details.motor;
      row.BatteryDetails = details.battery;
      row.VehicleChargerNo = details.charger;

      // A page error gets one retry after recovering (re-saving the same values is harmless).
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          row.Status = await updateChassis(chassis, details);
          const ok = /^(Updated|Dry run)/.test(row.Status);
          console.log(ok
            ? `[OK] ${chassis}: Engine ${details.motor}, Battery ${details.battery}, Charger ${details.charger} — ${row.Status}${attempt > 1 ? ' (on retry)' : ''}`
            : `[WARN] ${chassis}: ${row.Status}`);
          break;
        } catch (err) {
          row.Status = `Error: ${String((err as Error).message).split('\n')[0]!.slice(0, 150)}`;
          console.log(`[WARN] ${chassis}: ${row.Status}${attempt === 1 ? ' — recovering and retrying once' : ' — giving up on this chassis'}`);
          try {
            await recover();
          } catch (recoverErr) {
            // Can't get the form back at all: stop this dealer (keeping what's done) instead of
            // failing every remaining chassis the same way.
            const reason = `Not attempted: page could not be recovered after ${chassis} (${String((recoverErr as Error).message).split('\n')[0]!.slice(0, 100)})`;
            console.log(`[ERROR] ${reason}`);
            results.filter((r) => !r.Status).forEach((r) => { r.Status = reason; });
            return results;
          }
        }
      }
    }
  } finally {
    page.off('dialog', onDialog);
  }
  return results;
}

test.describe('PDI', () => {
  test.use({ storageState: { cookies: [], origins: [] } });
  // This writes to live DMS records, so it never runs as part of a plain `npm test` / run-tests.bat:
  // only via npm run test:pdi / test:pdi:dry / run-pdi.bat, which set PDI_RUN=1.
  test.skip(process.env.PDI_RUN !== '1', 'PDI update runs only via npm run test:pdi (sets PDI_RUN=1).');

  // All dealers run inside one test on the same page: real Chrome quits when its last window
  // closes, so one test() per dealer would kill the browser between dealers (see GSTR Report.spec.ts).
  test('Should fill blank PDI vehicle details from the VIN Details sheet', async ({ page }) => {
    const dealers = pickDealers();
    const vinDetails = loadVinDetails();
    console.log(`[STEP] Dealers: ${dealers.map((d) => d.dealercode).join(', ')}; ${vinDetails.size} VINs loaded from ${path.basename(VIN_FILE)}${DRY_RUN ? ' (DRY RUN — no Save)' : ''}`);

    // ~25 blank rows x ~25s each per dealer, with headroom.
    test.setTimeout(dealers.length * 20 * 60 * 1000);

    const startedAt = new Date();
    const results: ResultRow[] = [];
    const failedDealers: string[] = [];
    const failures: FailedDealer[] = [];
    for (const dealer of dealers) {
      console.log(`\n[STEP] ===== Dealer ${dealer.dealercode} =====`);
      try {
        results.push(...await processDealer(page, dealer, vinDetails));
      } catch (err) {
        const message = String((err as Error).message).split('\n')[0]!.slice(0, 200);
        failedDealers.push(String(dealer.dealercode));
        failures.push({ dealer: String(dealer.dealercode), error: message });
        console.log(`[ERROR] Dealer ${dealer.dealercode} failed: ${message}`);
        await test.info().attach(`${dealer.dealercode}-failure.png`, {
          body: await page.screenshot({ fullPage: true, timeout: 15000 }).catch(() => Buffer.from('')),
          contentType: 'image/png',
        }).catch(() => {});
      } finally {
        await resetSession(page, String(dealer.username));
      }
    }

    const dealersChecked = dealers.map((d) => {
      const code = String(d.dealercode);
      const failure = failures.find((f) => f.dealer === code);
      return { Dealer: code, DryRun: DRY_RUN ? 'Yes' : 'No', BlankFound: String(results.filter((r) => r.Dealer === code).length), Result: failure ? `Failed: ${failure.error}` : 'Checked' };
    });
    const file = saveResults(results, dealersChecked);
    console.log(`\n[STEP] Results saved to ${file}`);

    for (const dealer of dealers) {
      const code = String(dealer.dealercode);
      if (failedDealers.includes(code)) {
        console.log(`\n[RESULT] Dealer ${code}: FAILED before finishing — see the error above.`);
        continue;
      }
      const mine = results.filter((r) => r.Dealer === code);
      // In a dry run "done" means filled (nothing is saved).
      const done = (r: ResultRow) => r.Status!.startsWith(DRY_RUN ? 'Dry run' : 'Updated');
      const updated = mine.filter(done);
      console.log(`\n[RESULT] Dealer ${code}: ${updated.length} of ${mine.length} chassis ${DRY_RUN ? 'filled (dry run — nothing saved)' : 'updated'}`);
      updated.forEach((r) => console.log(`   ${r.ChassisNo}  Engine ${r.EngineNo}, Battery ${r.BatteryDetails}, Charger ${r.VehicleChargerNo}`));
      const notUpdated = mine.filter((r) => !done(r));
      if (notUpdated.length) {
        console.log(`   NOT ${DRY_RUN ? 'filled' : 'updated'}:`);
        notUpdated.forEach((r) => console.log(`   ${r.ChassisNo}  ${r.Status}`));
      }
    }

    await sendPdiReport(results, failures, {
      dryRun: DRY_RUN, startedAt, finishedAt: new Date(), dealers: dealers.map((d) => String(d.dealercode)), vinFile: VIN_FILE,
    }, file);

    // Fail (red) unless every blank chassis was updated — the same rule that decides whether the
    // report is emailed, so a red run never sends the report.
    const problems = problemsOf(results, failures, DRY_RUN);
    if (problems.length) {
      const rerun = [...new Set([...failedDealers, ...results.filter((r) => problems.some((p) => p.startsWith(`${r.Dealer} ${r.ChassisNo}:`))).map((r) => r.Dealer!)])];
      throw new Error(`${problems.length} problem(s) — report not emailed. Rerun with PDI_DEALER=${rerun.join(',')}\n  ${problems.slice(0, 20).join('\n  ')}`);
    }
  });
});
