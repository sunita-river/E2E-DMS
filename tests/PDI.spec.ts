import { test, Page, Dialog } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from './models/LoginPage';
import { writePdiDashboard, problemsOf, outcomeOf, FailedDealer, vinFilePath } from './utils/pdiReport';
import { writePdiWorkbook } from './utils/pdiWorkbook';
import { buildPdiDaily, localDate, RUNS_DIR, OUTPUT_DIR } from './utils/pdiDaily';

// PDI blank-detail fix-up, for every dealer in PIDDealerCodes / PIDUsers (credentials.json), PDI_PARALLEL at a time:
//   1. WORKSHOP -> PDI: find grid rows with a blank field (usually EngineNo) and save their chassis numbers.
//   2. SETTINGS -> Admin -> Modify vehicle details: for each saved chassis, Show it, look it up in the
//      VIN Details sheet and fill Engine No (Motor No column), Battery and Charger, then Save.
//      The sheet's columns are found by header name, so their order doesn't matter.
//
// Settings (all optional):
//   PDI_DEALER    only these dealer codes, comma-separated, e.g. 332009,332008 (default: all PDI dealers)
//   PDI_VIN_FILE  VIN Details workbook (default: the newest resources/VIN*.xlsx, e.g. "VIN Details.xlsx")
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
const EDIT_REASON = 'Details updated';
const PDI_PARALLEL = Number(process.env.PDI_PARALLEL || 3);
if (!Number.isInteger(PDI_PARALLEL) || PDI_PARALLEL < 1) throw new Error(`PDI_PARALLEL must be a whole number above 0 (got "${process.env.PDI_PARALLEL}")`);
const SHOW_WAIT_MS = Number(process.env.PDI_SHOW_WAIT_S || 30) * 1000;
const PDI_REFRESHES =Number(process.env.PDI_REFRESHES || 2);
const LOAD_TIMEOUT =Number(process.env.PDI_LOAD_TIMEOUT_S || 30) * 1000;
const SETTLE_MS = 1000;

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
const MODIFY_VEHICLE_URL = 'Common/ChangeChassisNo.aspx';

type ResultRow = Record<string, string>;
// Dealers run side by side, so each dealer's log lines carry its code.
type Log = (message: string) => void;

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

interface VinDetails { battery: string; motor: string; charger: string }

// Modify vehicle details fields and where their value comes from in the VIN sheet.
const VIN_FIELDS: { field: string; key: keyof VinDetails; label: string }[] = [
  { field: 'EngineNo', key: 'motor', label: 'Engine No' },
  { field: 'BatteryDetails', key: 'battery', label: 'Battery' },
  { field: 'VehicleChargerNo', key: 'charger', label: 'Charger' },
];
// A real serial has digits and no spaces. About 7% of VIN-sheet cells hold a note instead
// ("CHARGER NOT ASSIGNED YET", "BIN SERIAL DUPLICATION\nCONTACT FACTORY TEAM", "BEFIS3000395 : 480 W")
// or are empty; those must never be typed into the DMS.
const isSerial = (v: string) => !!v && !/\s/.test(v) && /\d/.test(v);

// VIN Details sheet columns are found by their header text in row 1, never by position. On 1-Oct-2026 the
// sheet's Battery and Motor columns were swapped and 41 chassis got Engine No and Battery the wrong way round.
const VIN_COLUMNS: { key: keyof VinDetails | 'vin'; header: RegExp; label: string }[] = [
  { key: 'vin', header: /^(vin|chassis)/, label: 'VIN' },
  { key: 'motor', header: /^(motor|engine)/, label: 'Motor No' },
  { key: 'battery', header: /^battery/, label: 'Battery No' },
  { key: 'charger', header: /^charger/, label: 'Charger' },
];

function loadVinDetails(): Map<string, VinDetails> {
  if (!fs.existsSync(VIN_FILE)) throw new Error(`VIN Details workbook not found: ${VIN_FILE}. Put a "VIN*.xlsx" in resources/ or set PDI_VIN_FILE.`);
  const workbook = XLSX.readFile(VIN_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
  const rows = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: '' });
  const headers = (rows[0] || []).map((h) => String(h).trim().toLowerCase());
  // Each column must match exactly one header, or the run stops before touching the DMS.
  const col = {} as Record<keyof VinDetails | 'vin', number>;
  for (const c of VIN_COLUMNS) {
    const found = headers.flatMap((h, i) => (c.header.test(h) ? [i] : []));
    if (found.length !== 1) {
      throw new Error(`VIN Details sheet (${path.basename(VIN_FILE)}): ${found.length ? 'more than one' : 'no'} "${c.label}" column in row 1 ` +
        `(headers: ${headers.map((h) => `"${h}"`).join(', ')}). Fix the header row and run again.`);
    }
    col[c.key] = found[0]!;
  }
  console.log(`[STEP] VIN sheet columns: ${VIN_COLUMNS.map((c) => `${c.label} = ${XLSX.utils.encode_col(col[c.key])} ("${rows[0]![col[c.key]]}")`).join(', ')}`);
  const byVin = new Map<string, VinDetails>();
  for (const row of rows.slice(1)) {
    const key = String(row[col.vin] || '').trim().toUpperCase();
    const cell = (k: keyof VinDetails) => String(row[col[k]] ?? '').trim();
    if (key) byVin.set(key, { battery: cell('battery'), motor: cell('motor'), charger: cell('charger') });
  }
  checkColumnProfile(byVin);
  return byVin;
}

// Headers can be right while the data under them is swapped (as on 1-Oct-2026). There's no fixed rule for
// what each serial looks like; instead each run saves how each column's serials start (first 3 characters)
// and the next run compares: if this sheet's Motor column looks like last time's Battery column, the run
// stops before touching the DMS. A column whose format is simply new (a new supplier) only gets a warning.
// PDI_ACCEPT_VIN_COLUMNS=1 skips the check once and saves the sheet as the new baseline.
const PROFILE_FILE = path.join(OUTPUT_DIR, 'PDI_VIN_Column_Profile.json');
type ColumnProfile = Record<keyof VinDetails, Record<string, number>>; // prefix -> share of the column's serials

function profileOf(byVin: Map<string, VinDetails>): ColumnProfile {
  const profile = {} as ColumnProfile;
  for (const k of ['motor', 'battery', 'charger'] as const) {
    const prefixes = [...byVin.values()].map((d) => d[k]).filter(isSerial).map((v) => v.slice(0, 3).toUpperCase());
    const counts: Record<string, number> = {};
    for (const p of prefixes) counts[p] = (counts[p] || 0) + 1;
    profile[k] = Object.fromEntries(Object.entries(counts).map(([p, n]) => [p, n / prefixes.length]));
  }
  return profile;
}

// How much two columns' prefixes overlap: 1 = same mix of prefixes, 0 = nothing in common.
const overlap = (a: Record<string, number>, b: Record<string, number>) =>
  Object.entries(a).reduce((sum, [p, share]) => sum + Math.min(share, b[p] || 0), 0);

function checkColumnProfile(byVin: Map<string, VinDetails>) {
  const now = profileOf(byVin);
  const label = (k: keyof VinDetails) => VIN_COLUMNS.find((c) => c.key === k)!.label;
  const prev: ColumnProfile | null = fs.existsSync(PROFILE_FILE) ? JSON.parse(fs.readFileSync(PROFILE_FILE, 'utf8')) : null;
  if (prev && process.env.PDI_ACCEPT_VIN_COLUMNS !== '1') {
    const keys = ['motor', 'battery', 'charger'] as const;
    const mixedUp: string[] = [];
    for (const k of keys) {
      if (!Object.keys(now[k]).length || !prev[k]) continue;
      const own = overlap(now[k], prev[k]);
      const [other, best] = keys.filter((j) => j !== k && prev[j])
        .map((j) => [j, overlap(now[k], prev[j])] as const).sort((a, b) => b[1] - a[1])[0] || [];
      if (other && best! >= 0.5 && best! > own) {
        mixedUp.push(`the "${label(k)}" column now holds values like last run's "${label(other)}" column`);
      } else if (own < 0.2) {
        console.log(`[WARN] VIN sheet: the "${label(k)}" column's values start differently from last run ` +
          `(now ${Object.keys(now[k]).slice(0, 3).join(', ')}...). New format? Carrying on.`);
      }
    }
    if (mixedUp.length) {
      throw new Error(`VIN Details sheet (${path.basename(VIN_FILE)}): ${mixedUp.join('; ')}. The columns look swapped. ` +
        'Nothing was typed into the DMS. Fix the sheet and run again, or if it is right, run once with PDI_ACCEPT_VIN_COLUMNS=1.');
    }
  }
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(PROFILE_FILE, JSON.stringify(now, null, 2));
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

  log('[STEP] Opening SETTINGS -> Admin -> Modify vehicle details...');
  await page.getByRole('link', { name: /^\W*SETTINGS\W*$/ }).click();
  await page.getByRole('link', { name: /^\W*Admin\W*$/ }).first().click();
  await page.locator('a[href$="ChangeChassisNo.aspx"]').first().click();
  await page.waitForURL(/ChangeChassisNo\.aspx/i, { timeout: 60000 }).catch(() => page.goto(MODIFY_VEHICLE_URL));

  // Save may raise a confirm()/alert(); accept it and keep the text for the result sheet.
  let lastDialog = '';
  const onDialog = async (d: Dialog) => { lastDialog = d.message(); await d.accept().catch(() => {}); };
  page.on('dialog', onDialog);

  // Show and Save post the page back. Tag the current Chassis No field and wait for an untagged one
  // (holding `chassis`, if given), so a vehicle already on screen doesn't count as reloaded. On timeout
  // carry on; the caller's checks catch a vehicle that didn't load.
  const chassisField = '[name="ctl00$cpMain$ChassisNo"]';
  const postBack = async (button: string, chassis?: string) => {
    await page.locator(chassisField).evaluateAll((els) => els.forEach((e) => e.setAttribute('data-pw-stale', '1'))).catch(() => {});
    await page.locator(button).click();
    const reloaded = await page.waitForFunction(({ sel, want }) => {
      const el = document.querySelector(sel) as HTMLInputElement | null;
      return !!el && !el.hasAttribute('data-pw-stale') && (!want || el.value.trim().toUpperCase() === want);
    }, { sel: chassisField, want: chassis?.toUpperCase() }, { timeout: LOAD_TIMEOUT }).then(() => true, () => false);
    if (!reloaded) log(`[WARN] Page did not reload within ${LOAD_TIMEOUT / 1000}s after ${button}${chassis ? ` for ${chassis}` : ''}, continuing.`);
    await page.waitForLoadState('load').catch(() => {});
    await page.waitForTimeout(SETTLE_MS); // let any alert() raised on load fire
  };
  const showVehicle = async (chassis: string) => {
    await page.locator('[name="ctl00$cpMain$IdentifierValue"]').fill(chassis);
    await postBack('#ctl00_cpMain_cmdShow', chassis);
  };
  const fieldValue = async (name: string) => (await page.locator(`[name="ctl00$cpMain$${name}"]`).inputValue()).trim();
  const identifierInput = page.locator('[name="ctl00$cpMain$IdentifierValue"]');

  // The DMS sometimes drops the session mid-dealer (a click times out, then every later field is
  // missing), so if reloading Modify vehicle details doesn't bring the form back, log in again.
  const recover = async () => {
    await page.goto(MODIFY_VEHICLE_URL).catch(() => {});
    if (await identifierInput.isVisible({ timeout: 15000 }).catch(() => false)) return;
    log('[STEP] Modify vehicle details not available (session lost?), logging in again...');
    await loginPage.goto();
    await loginPage.login(dealer);
    await loginPage.verifyDashboard();
    await page.goto(MODIFY_VEHICLE_URL);
    await identifierInput.waitFor({ timeout: 60000 });
  };

  // One attempt at a chassis; returns its Status. Throws on a page error (timeouts etc.).
  const updateChassis = async (row: ResultRow, details: VinDetails): Promise<string> => {
    const chassis = row.ChassisNo!;
    await showVehicle(chassis);

    const loaded = await fieldValue('ChassisNo');
    if (loaded.toUpperCase() !== chassis.toUpperCase()) return `Show did not load this vehicle (got "${loaded}")`;

    // A correction run keeps what the DMS held before, so the report shows what was overwritten.
    if (fixChassis) {
      const before = await Promise.all(VIN_FIELDS.map(async (f) => `${f.label} ${(await fieldValue(f.field)) || '(empty)'}`));
      row.BlankFields = `Correction, before: ${before.join(', ')}`;
    }

    // Only real serials are typed in. A VIN-sheet note leaves the DMS field as it is; it is reported when
    // the field is still empty, or always in a correction run (the value left there may be the wrong one).
    const filled: [field: string, value: string][] = [];
    const notes: string[] = [];
    for (const f of VIN_FIELDS) {
      const value = details[f.key];
      const current = await fieldValue(f.field);
      if (isSerial(value)) {
        await page.locator(`[name="ctl00$cpMain$${f.field}"]`).fill(value);
        filled.push([f.field, value]);
      } else if (!current || fixChassis) {
        notes.push(`${f.label} "${value.replace(/\s+/g, ' ') || 'empty'}"${current ? ` (DMS still has "${current}")` : ''}`);
      }
    }
    const noteText = notes.length ? `VIN sheet has no usable ${notes.join(', ')}` : '';
    if (!filled.length) return `${noteText || 'VIN sheet has no usable values'}, nothing saved`;
    await page.locator('[name="ctl00$cpMain$EditChassisReason"]').fill(EDIT_REASON);
    if (DRY_RUN) return noteText ? `${noteText}; the rest filled (dry run)` : 'Dry run: filled, not saved';

    lastDialog = '';
    await postBack('#ctl00_cpMain_cmdSave');
    const saveMessage = lastDialog;

    // Re-open the vehicle to check the DMS actually kept the new values.
    await showVehicle(chassis);
    const notKept: string[] = [];
    for (const [field, value] of filled) {
      const now = await fieldValue(field);
      if (now !== value) notKept.push(`${VIN_FIELDS.find((f) => f.field === field)!.label}: DMS has "${now}"`);
    }
    const suffix = saveMessage ? ` (${saveMessage})` : '';
    if (notKept.length) return `Save clicked but values not kept: ${notKept.join('; ')}${suffix}`;
    return noteText ? `${noteText}; the rest saved${suffix}` : `Updated${suffix}`;
  };

  try {
    for (const row of results) {
      const chassis = row.ChassisNo!;
      const details = vinDetails.get(chassis.toUpperCase());
      if (!details) {
        row.Status = 'Not in VIN Details sheet';
        log(`[WARN] ${chassis}: not found in VIN Details, skipped.`);
        continue;
      }
      row.EngineNo = isSerial(details.motor) ? details.motor : '';
      row.BatteryDetails = isSerial(details.battery) ? details.battery : '';
      row.VehicleChargerNo = isSerial(details.charger) ? details.charger : '';

      // A page error gets one retry after recovering (re-saving the same values is harmless).
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          row.Status = await updateChassis(row, details);
          const ok = /^(Updated|Dry run)/.test(row.Status);
          log(ok
            ? `[OK] ${chassis}: Engine ${details.motor}, Battery ${details.battery}, Charger ${details.charger}: ${row.Status}${attempt > 1 ? ' (on retry)' : ''}`
            : `[WARN] ${chassis}: ${row.Status}`);
          break;
        } catch (err) {
          row.Status = `Error: ${String((err as Error).message).split('\n')[0]!.slice(0, 150)}`;
          log(`[WARN] ${chassis}: ${row.Status}${attempt === 1 ? ', recovering and retrying once' : ', giving up on this chassis'}`);
          try {
            await recover();
          } catch (recoverErr) {
            // Can't get the form back at all: stop this dealer (keeping what's done) instead of
            // failing every remaining chassis the same way.
            const reason = `Not attempted: page could not be recovered after ${chassis} (${String((recoverErr as Error).message).split('\n')[0]!.slice(0, 100)})`;
            log(`[ERROR] ${reason}`);
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
          const message = String((err as Error).message).split('\n')[0]!.slice(0, 200);
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
    const problems = problemsOf(results.filter((r) => outcomeOf(r) !== 'missing'), failures, DRY_RUN);
    if (problems.length) {
      const rerun = [...new Set([...failedDealers, ...results.filter((r) => problems.some((p) => p.startsWith(`${r.Dealer} ${r.ChassisNo}:`))).map((r) => r.Dealer!)])];
      throw new Error(`${problems.length} problem(s). Rerun with PDI_DEALER=${rerun.join(',')}\n  ${problems.slice(0, 20).join('\n  ')}`);
    }
  });
});
