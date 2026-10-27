import { Page, Dialog } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from '../../tests/models/LoginPage';
import { OUTPUT_DIR } from './pdiDaily';

// VIN Details sheet + SETTINGS -> Admin -> Modify vehicle details, shared by PDI.spec.ts (blank PDI rows)
// and PDI_Update_InTransit.spec.ts (In-Transit vehicles).

export type ResultRow = Record<string, string>;
export type Log = (message: string) => void;

// Reason the DMS records with each Save, dated (local date, dd-mm-yyyy) so automated edits can be told apart.
export const editReason = (d = new Date()) =>
  `Added via Automation ${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
export const MODIFY_VEHICLE_URL = 'Common/ChangeChassisNo.aspx';
const SETTLE_MS = 1000;

export interface VinDetails { battery: string; motor: string; charger: string }

// Modify vehicle details fields and where their value comes from in the VIN sheet.
export const VIN_FIELDS: { field: string; key: keyof VinDetails; label: string }[] = [
  { field: 'EngineNo', key: 'motor', label: 'Engine No' },
  { field: 'BatteryDetails', key: 'battery', label: 'Battery' },
  { field: 'VehicleChargerNo', key: 'charger', label: 'Charger' },
];
// A real serial has digits and no spaces. About 7% of VIN-sheet cells hold a note instead
// ("CHARGER NOT ASSIGNED YET", "BIN SERIAL DUPLICATION\nCONTACT FACTORY TEAM", "BEFIS3000395 : 480 W")
// or are empty; those must never be typed into the DMS.
export const isSerial = (v: string) => !!v && !/\s/.test(v) && /\d/.test(v);

// VIN Details sheet columns are found by their header text in row 1, never by position. On 1-Oct-2026 the
// sheet's Battery and Motor columns were swapped and 41 chassis got Engine No and Battery the wrong way round.
const VIN_COLUMNS: { key: keyof VinDetails | 'vin'; header: RegExp; label: string }[] = [
  { key: 'vin', header: /^(vin|chassis)/, label: 'VIN' },
  { key: 'motor', header: /^(motor|engine|min\b)/, label: 'Motor No' },   // MIN = Motor Identification No
  { key: 'battery', header: /^(battery|bin\b)/, label: 'Battery No' },   // BIN = Battery Identification No
  { key: 'charger', header: /^charger/, label: 'Charger' },
];

export function loadVinDetails(vinFile: string): Map<string, VinDetails> {
  if (!fs.existsSync(vinFile)) throw new Error(`VIN Details workbook not found: ${vinFile}. Run npm run pdi:vin to fetch the VIN Google Sheet, or set PDI_VIN_FILE.`);
  // raw: a CSV's cells stay text, so a long serial isn't turned into a number like 8.9E+18 (no effect on .xlsx).
  const workbook = XLSX.readFile(vinFile, { raw: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
  const rows = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: '' });
  const headers = (rows[0] || []).map((h) => String(h).trim().toLowerCase());
  // Each column must match exactly one header, or the run stops before touching the DMS.
  const col = {} as Record<keyof VinDetails | 'vin', number>;
  for (const c of VIN_COLUMNS) {
    const found = headers.flatMap((h, i) => (c.header.test(h) ? [i] : []));
    if (found.length !== 1) {
      throw new Error(`VIN Details sheet (${path.basename(vinFile)}): ${found.length ? 'more than one' : 'no'} "${c.label}" column in row 1 ` +
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
  checkColumnProfile(byVin, vinFile);
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

function checkColumnProfile(byVin: Map<string, VinDetails>, vinFile: string) {
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
      throw new Error(`VIN Details sheet (${path.basename(vinFile)}): ${mixedUp.join('; ')}. The columns look swapped. ` +
        'Nothing was typed into the DMS. Fix the sheet and run again, or if it is right, run once with PDI_ACCEPT_VIN_COLUMNS=1.');
    }
  }
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(PROFILE_FILE, JSON.stringify(now, null, 2));
}

export interface UpdateOptions {
  dryRun: boolean;
  loadTimeoutMs: number;
  // Set when the fields may already hold values: their old values are written to row.BlankFields
  // after this label (e.g. "Correction, before"), so the report shows what was overwritten.
  recordBefore?: string;
  // Only type the fields whose DMS value differs from the VIN sheet; a vehicle that already matches is not
  // saved ("Already correct"). Also sets row.Before<Field> to the DMS values and row.Changed to the fields saved.
  onlyChanged?: boolean;
  // Called as soon as each row has its final Status (e.g. to keep progress of a long run).
  onRow?: (row: ResultRow) => void;
  // How to log in again if the session drops (default: the dealer login with `dealer`).
  relogin?: () => Promise<void>;
}

// SETTINGS -> Admin -> Modify vehicle details: for each row's ChassisNo, Show it, fill Engine No, Battery and
// Charger from the VIN sheet, Save and re-open it to check the DMS kept them. Sets each row's EngineNo,
// BatteryDetails, VehicleChargerNo and Status in place.
export async function updateVehicles(page: Page, loginPage: LoginPage, dealer: UserCredentials,
  results: ResultRow[], vinDetails: Map<string, VinDetails>, log: Log, opts: UpdateOptions): Promise<ResultRow[]> {
  const { dryRun, loadTimeoutMs, recordBefore, onlyChanged } = opts;
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
    }, { sel: chassisField, want: chassis?.toUpperCase() }, { timeout: loadTimeoutMs }).then(() => true, () => false);
    if (!reloaded) log(`[WARN] Page did not reload within ${loadTimeoutMs / 1000}s after ${button}${chassis ? ` for ${chassis}` : ''}, continuing.`);
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
    if (opts.relogin) {
      await opts.relogin();
    } else {
      await loginPage.goto();
      await loginPage.login(dealer);
      await loginPage.verifyDashboard();
    }
    await page.goto(MODIFY_VEHICLE_URL);
    await identifierInput.waitFor({ timeout: 60000 });
  };

  // One attempt at a chassis; returns its Status. Throws on a page error (timeouts etc.).
  const updateChassis = async (row: ResultRow, details: VinDetails): Promise<string> => {
    const chassis = row.ChassisNo!;
    await showVehicle(chassis);

    const loaded = await fieldValue('ChassisNo');
    if (loaded.toUpperCase() !== chassis.toUpperCase()) return `Show did not load this vehicle (got "${loaded}")`;

    const before: Record<string, string> = {};
    for (const f of VIN_FIELDS) before[f.field] = await fieldValue(f.field);
    if (recordBefore) row.BlankFields = `${recordBefore}: ${VIN_FIELDS.map((f) => `${f.label} ${before[f.field] || '(empty)'}`).join(', ')}`;
    if (onlyChanged) VIN_FIELDS.forEach((f) => { row[`Before${f.field}`] = before[f.field]!; });

    // Only real serials are typed in. A VIN-sheet note leaves the DMS field as it is; it is reported when
    // the field is still empty, or always when recording old values (the value left there may be the wrong one).
    const filled: [field: string, value: string][] = [];
    const notes: string[] = [];
    let same = 0;
    for (const f of VIN_FIELDS) {
      const value = details[f.key];
      const current = before[f.field]!;
      if (isSerial(value)) {
        if (onlyChanged && current.toUpperCase() === value.toUpperCase()) { same++; continue; }
        await page.locator(`[name="ctl00$cpMain$${f.field}"]`).fill(value);
        filled.push([f.field, value]);
      } else if (!current || recordBefore) {
        notes.push(`${f.label} "${value.replace(/\s+/g, ' ') || 'empty'}"${current ? ` (DMS still has "${current}")` : ''}`);
      }
    }
    if (onlyChanged) row.Changed = filled.map(([field]) => VIN_FIELDS.find((f) => f.field === field)!.label).join(', ');
    const noteText = notes.length ? `VIN sheet has no usable ${notes.join(', ')}` : '';
    if (!filled.length) {
      if (onlyChanged && same) return noteText ? `${noteText}; the rest already correct` : 'Already correct';
      return `${noteText || 'VIN sheet has no usable values'}, nothing saved`;
    }
    await page.locator('[name="ctl00$cpMain$EditChassisReason"]').fill(editReason());
    if (dryRun) return noteText ? `${noteText}; the rest filled (dry run)` : 'Dry run: filled, not saved';

    lastDialog = '';
    await postBack('#ctl00_cpMain_cmdSave');
    // The DMS can refuse a Save without an alert(), only showing a message on the page (e.g. a serial
    // already used on another vehicle), so pick up any visible message/error text as well.
    const pageMessage = await page.evaluate(() => {
      const texts = [...document.querySelectorAll<HTMLElement>('[id*="lbl" i], [id*="msg" i], [id*="message" i], [id*="error" i], [id*="valid" i], [class*="error" i], [class*="alert" i]')]
        .filter((el) => el.offsetParent !== null && !el.querySelector('input, select, textarea'))
        .map((el) => (el.innerText || '').replace(/\s+/g, ' ').trim())
        .filter((t) => t.length > 3 && t.length < 300 && /(exist|already|duplicate|invalid|error|not|fail|must|required|success|saved|updated)/i.test(t));
      return [...new Set(texts)].join(' | ');
    }).catch(() => '');
    const saveMessage = [lastDialog, pageMessage].filter(Boolean).join(' | ');

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
        opts.onRow?.(row);
        continue;
      }
      row.EngineNo = isSerial(details.motor) ? details.motor : '';
      row.BatteryDetails = isSerial(details.battery) ? details.battery : '';
      row.VehicleChargerNo = isSerial(details.charger) ? details.charger : '';

      // A page error gets one retry after recovering (re-saving the same values is harmless).
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          row.Status = await updateChassis(row, details);
          const ok = /^(Updated|Dry run|Already correct)/.test(row.Status);
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
      opts.onRow?.(row);
    }
  } finally {
    page.off('dialog', onDialog);
  }
  return results;
}
