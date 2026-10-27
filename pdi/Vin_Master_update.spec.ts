import { test, Page } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage, UserCredentials } from '../tests/models/LoginPage';
import { vinFilePath } from './utils/pdiReport';
import { OUTPUT_DIR } from './utils/pdiDaily';
import { loadVinDetails, updateVehicles, isSerial, ResultRow, Log } from './utils/pdiVehicle';

// VIN master correction: every VIN in the VIN Details sheet ("VIN BIN Update (Sunitha)" Google Sheet, stock and
// sold vehicles alike) is checked against the DMS and corrected where it differs, as River User
// (riverUserValidUser in credentials.json), VIN_MASTER_PARALLEL sessions at once:
//   SETTINGS -> Admin -> Modify vehicle details: Show the VIN and read its Engine No, Battery and Charger.
//   - all three already match the sheet: nothing is saved ("Already correct")
//   - any differs or is blank: only those fields are typed from the sheet, saved, and the VIN re-opened to
//     check the DMS kept them; the old values are kept in the result
//   - a sheet note ("CHARGER NOT ASSIGNED YET ...") or empty cell is never typed; that DMS field is left as it is
// A Save the DMS refuses because another VIN still holds the number ("must be unique", e.g. two vehicles with
// each other's charger) is tried once more after the full pass, when the first pass has usually freed it.
//
// Every finished VIN goes to a progress file straight away, so a run that stops (PC asleep, DMS down) carries on
// where it stopped when started again; VINs that ended in an error are tried again. The result workbook is
// built from that file at the end, so it covers every run so far:
//   Output/VIN Master/VIN_Master_Update_<date-time>.xlsx  Summary, Changed (every chassis whose DMS values were
//   changed, before and after), Problems (not changed and why), All
//
// Settings (all optional):
//   PDI_DRY_RUN          1 = fill the form but don't click Save (own progress file, the real one is not touched)
//   VIN_LIMIT            only the first N VINs still to do (e.g. 50 for a first check)
//   VIN_LIST             only these VINs: comma-separated, or the path of a text file with one per line
//   VIN_MASTER_PARALLEL  sessions at once, each in its own browser window (default 4)
//   VIN_MASTER_RESTART   1 = start again from the first VIN (the old progress file is kept, renamed)
//   VIN_MASTER_REPORT    1 = only build the result workbook from the progress file, no login
//   PDI_VIN_FILE, PDI_LOAD_TIMEOUT_S, PDI_ACCEPT_VIN_COLUMNS  as for PDI.spec.ts
const CREDENTIALS_FILE = process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
const VIN_FILE = vinFilePath();
const DRY_RUN = process.env.PDI_DRY_RUN === '1';
const LOAD_TIMEOUT = Number(process.env.PDI_LOAD_TIMEOUT_S || 30) * 1000;
const PARALLEL = Number(process.env.VIN_MASTER_PARALLEL || 4);
if (!Number.isInteger(PARALLEL) || PARALLEL < 1) throw new Error(`VIN_MASTER_PARALLEL must be a whole number above 0 (got "${process.env.VIN_MASTER_PARALLEL}")`);
const LIMIT = Number(process.env.VIN_LIMIT || 0);
const BATCH = 25; // VINs a session takes from the shared queue at a time

const MASTER_DIR = path.join(OUTPUT_DIR, 'VIN Master');
const PROGRESS_FILE = path.join(MASTER_DIR, `VIN_Master_progress${DRY_RUN ? '_DRYRUN' : ''}.jsonl`);

// Results worth another try in a later pass or run: page errors, sessions that couldn't be recovered,
// a vehicle that didn't load, and Saves the DMS refused because another VIN holds the number.
const RETRY = /^(Error|Not attempted)|Show did not load|must be unique/i;
const UNIQUE = /must be unique/i;

type Outcome = 'Changed' | 'Partly changed' | 'Already correct' | 'Dry run' | 'Sheet has no usable value' | 'Not found in DMS' | 'Refused by DMS' | 'Error';
function outcomeOf(status: string): Outcome {
  if (/^Already correct/.test(status)) return 'Already correct';
  if (/^Updated/.test(status)) return 'Changed';
  if (/^Dry run|\(dry run\)/.test(status)) return 'Dry run';
  if (/; the rest saved/.test(status)) return 'Partly changed';
  if (/^VIN sheet has no usable/.test(status)) return 'Sheet has no usable value';
  if (/Show did not load/.test(status)) return 'Not found in DMS';
  if (/^Save clicked but values not kept/.test(status)) return 'Refused by DMS';
  return 'Error';
}

// Latest record per VIN from the progress file.
function readProgress(): Map<string, ResultRow> {
  const done = new Map<string, ResultRow>();
  if (!fs.existsSync(PROGRESS_FILE)) return done;
  for (const line of fs.readFileSync(PROGRESS_FILE, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    try { const r = JSON.parse(line) as ResultRow; done.set(r.ChassisNo!, r); } catch { /* a line cut off by a stop */ }
  }
  return done;
}

const vinList = (s: string) => {
  const text = fs.existsSync(s) ? fs.readFileSync(s, 'utf8') : s;
  return new Set(text.split(/[,\s]+/).map((v) => v.trim().toUpperCase()).filter(Boolean));
};

function writeReport(): string {
  const rows = [...readProgress().values()];
  const sheetRow = (r: ResultRow) => ({
    'Chassis No': r.ChassisNo, Result: outcomeOf(r.Status || ''), 'Fields changed': r.Changed || '',
    'Engine No before': r.BeforeEngineNo ?? '', 'Engine No sheet': r.SheetMotor ?? '',
    'Battery before': r.BeforeBatteryDetails ?? '', 'Battery sheet': r.SheetBattery ?? '',
    'Charger before': r.BeforeVehicleChargerNo ?? '', 'Charger sheet': r.SheetCharger ?? '',
    'DMS message': r.Status, Session: r.Lane, 'Checked at': r.At,
  });
  const widths = [{ wch: 20 }, { wch: 22 }, { wch: 26 }, { wch: 16 }, { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 16 }, { wch: 16 }, { wch: 90 }, { wch: 8 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  const counts = new Map<Outcome, number>();
  rows.forEach((r) => counts.set(outcomeOf(r.Status || ''), (counts.get(outcomeOf(r.Status || '')) || 0) + 1));
  const summary = XLSX.utils.aoa_to_sheet([
    [`River DMS · VIN master update${DRY_RUN ? ' (DRY RUN, nothing saved)' : ''}`],
    [`${new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })} · VIN sheet: ${path.basename(VIN_FILE)} · ${rows.length} VIN(s) checked so far`],
    [],
    ['Result', 'VINs'],
    ...[...counts.entries()].sort((a, b) => b[1] - a[1]),
  ]);
  summary['!cols'] = [{ wch: 30 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, summary, 'Summary');
  const add = (name: string, list: ResultRow[]) => {
    const ws = XLSX.utils.json_to_sheet(list.map(sheetRow));
    ws['!cols'] = widths;
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    XLSX.utils.book_append_sheet(wb, ws, name);
  };
  const byOutcome = (...o: Outcome[]) => rows.filter((r) => o.includes(outcomeOf(r.Status || '')));
  add('Changed', byOutcome('Changed', 'Partly changed', 'Dry run'));
  add('Problems', byOutcome('Sheet has no usable value', 'Not found in DMS', 'Refused by DMS', 'Error'));
  add('All', rows);
  const stamp = new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-');
  const file = path.join(MASTER_DIR, `VIN_Master_Update_${stamp}${DRY_RUN ? '_DRYRUN' : ''}.xlsx`);
  fs.mkdirSync(MASTER_DIR, { recursive: true });
  XLSX.writeFile(wb, file);
  console.log(`[STEP] Result workbook: ${file}\n[RESULT] ${[...counts.entries()].map(([o, n]) => `${o}: ${n}`).join(', ')}`);
  return file;
}

test.describe('VIN master update', () => {
  test.use({ storageState: { cookies: [], origins: [] } });
  // This writes to live DMS records, so it never runs as part of a plain `npm test`:
  // only via npm run vin:master / vin:master:dry, which set PDI_RUN=1.
  test.skip(process.env.PDI_RUN !== '1', 'VIN master update runs only via npm run vin:master (sets PDI_RUN=1).');

  // All sessions run inside one test: real Chrome quits when its last window closes (see PDI.spec.ts).
  test('Should correct Engine No, Battery and Charger of every VIN from the VIN Details sheet', async ({ page, browser }) => {
    test.setTimeout(0); // tens of thousands of VINs; PW_GLOBAL_TIMEOUT_MIN in the npm script is the cap
    if (process.env.VIN_MASTER_REPORT === '1') { writeReport(); return; }

    const riverUser: UserCredentials = credentials.riverUserValidUser;
    if (!riverUser?.username) throw new Error(`No riverUserValidUser in ${path.basename(CREDENTIALS_FILE)}.`);
    const vinDetails = loadVinDetails(VIN_FILE);

    fs.mkdirSync(MASTER_DIR, { recursive: true });
    if (process.env.VIN_MASTER_RESTART === '1' && fs.existsSync(PROGRESS_FILE)) {
      const old = PROGRESS_FILE.replace(/\.jsonl$/, `_${new Date().toISOString().slice(0, 16).replace(/[T:]/g, '-')}.jsonl`);
      fs.renameSync(PROGRESS_FILE, old);
      console.log(`[STEP] Starting again from the first VIN; the old progress file is now ${path.basename(old)}`);
    }

    // A VIN is done when its last result isn't one to retry and the sheet still has the values it was checked with.
    const progress = readProgress();
    const sheetOf = (vin: string) => vinDetails.get(vin)!;
    const isDone = (vin: string) => {
      const r = progress.get(vin);
      if (!r || RETRY.test(r.Status || '')) return false;
      const d = sheetOf(vin);
      return r.SheetMotor === d.motor && r.SheetBattery === d.battery && r.SheetCharger === d.charger;
    };
    let vins = [...vinDetails.keys()].filter((v) => /^[A-Z0-9]{17}$/.test(v));
    if (process.env.VIN_LIST) {
      const wanted = vinList(process.env.VIN_LIST);
      const unknown = [...wanted].filter((v) => !vinDetails.has(v));
      if (unknown.length) console.log(`[WARN] Not in the VIN sheet, skipped: ${unknown.join(', ')}`);
      vins = vins.filter((v) => wanted.has(v));
    } else {
      vins = vins.filter((v) => !isDone(v));
    }
    if (LIMIT > 0) vins = vins.slice(0, LIMIT);
    const nothingUsable = vins.filter((v) => !Object.values(sheetOf(v)).some(isSerial)).length;
    console.log(`[STEP] ${vinDetails.size} VINs in ${path.basename(VIN_FILE)}, ${progress.size} checked before, ${vins.length} to do now ` +
      `(${nothingUsable} with no usable value in the sheet), ${PARALLEL} session(s) as River User ${riverUser.username}${DRY_RUN ? ' (DRY RUN, no Save)' : ''}`);
    if (!vins.length) { writeReport(); return; }

    const start = Date.now();
    let finished = 0;
    let total = vins.length;
    const record = (lane: number) => (row: ResultRow) => {
      const d = sheetOf(row.ChassisNo!);
      const rec = { ...row, SheetMotor: d.motor, SheetBattery: d.battery, SheetCharger: d.charger, Lane: `S${lane}`, At: new Date().toLocaleString('en-IN') };
      fs.appendFileSync(PROGRESS_FILE, `${JSON.stringify(rec)}\n`);
      finished++;
      if (finished % 100 === 0 || finished === total) {
        const perVin = (Date.now() - start) / finished;
        const left = Math.round(((total - finished) * perVin) / 60000);
        console.log(`[PROGRESS] ${finished} of ${total} VINs (${Math.round((finished / total) * 100)}%), about ${Math.floor(left / 60)}h ${left % 60}m left`);
      }
    };

    const runLane = async (lanePage: Page, lane: number, queue: string[], results: ResultRow[]) => {
      const t0 = Date.now();
      const log: Log = (message) => console.log(`[S${lane} +${Math.round((Date.now() - t0) / 1000)}s] ${message}`);
      const loginPage = new LoginPage(lanePage, log);
      const login = async () => {
        await loginPage.goto();
        await loginPage.loginAsRiverUser(riverUser);
        await loginPage.verifyDashboard();
      };
      await login();
      for (let batch = queue.splice(0, BATCH); batch.length; batch = queue.splice(0, BATCH)) {
        const rows: ResultRow[] = batch.map((vin) => ({ Dealer: '', ChassisNo: vin, BlankFields: '', EngineNo: '', BatteryDetails: '', VehicleChargerNo: '', Status: '' }));
        try {
          results.push(...await updateVehicles(lanePage, loginPage, riverUser, rows, vinDetails, log,
            { dryRun: DRY_RUN, loadTimeoutMs: LOAD_TIMEOUT, onlyChanged: true, relogin: login, onRow: record(lane) }));
        } catch (err) {
          // Opening Modify vehicle details failed (session gone?): these VINs stay to do, log in again for the next batch.
          log(`[ERROR] Batch of ${batch.length} not finished: ${String((err as Error).message).split('\n')[0]}; logging in again`);
          await login().catch((e) => log(`[ERROR] Login failed: ${String((e as Error).message).split('\n')[0]}`));
        }
      }
    };

    const contexts = await Promise.all(Array.from({ length: Math.min(PARALLEL, vins.length) - 1 }, () =>
      browser.newContext({ baseURL: test.info().project.use.baseURL ?? 'https://rivermobility.gaindms.com/', viewport: null })));
    const pages = [page, ...await Promise.all(contexts.map((c) => c.newPage()))];
    const results: ResultRow[] = [];
    try {
      const queue = [...vins];
      await Promise.all(pages.map((p, i) => runLane(p, i + 1, queue, results)));

      // Second pass: Saves refused because another VIN held the number, which the first pass may have changed since.
      const again = results.filter((r) => UNIQUE.test(r.Status || '')).map((r) => r.ChassisNo!);
      if (again.length && !DRY_RUN) {
        console.log(`\n[STEP] Second pass: ${again.length} VIN(s) the DMS refused as "must be unique", trying them again...`);
        total = finished + again.length;
        const queue2 = [...again];
        await Promise.all(pages.map((p, i) => runLane(p, i + 1, queue2, [])));
      }
    } finally {
      await Promise.all(contexts.map((c) => c.close().catch(() => {})));
      writeReport();
    }
    const left = vins.filter((v) => !readProgress().has(v)).length;
    if (left) console.log(`[WARN] ${left} VIN(s) were not finished; run again to carry on with them.`);
  });
});
