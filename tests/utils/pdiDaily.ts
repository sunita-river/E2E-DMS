import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { ResultRow, FailedDealer, ReportMeta, vinFilePath, writePdiMasterDashboard } from './pdiReport';
import { writePdiWorkbook, writePdiMaster, readPdiRow } from './pdiWorkbook';
import { recordPdiHistory, PdiHistory } from './pdiHistory';

// The day's single PDI copy: Output/PDI_Daily_<yyyy-mm-dd>.xlsx (+ .html dashboard), rebuilt by every
// PDI run (tests/PDI.spec.ts) from that day's run files and emailed once by tests/PDIDailyReport.spec.ts.
//
// Each run's own result file (PDI_BlankChassis_<time>.xlsx) goes to Output/PDI runs/, which is only
// the input for this merge. Merging rules:
//   - A chassis seen in several runs keeps its latest result.
//   - An earlier error/skip is shown as "Updated (no longer blank at a later check)" when a later run
//     checked that dealer again and the chassis was no longer blank.
//   - Dry-run runs are left out (nothing was saved).
export const OUTPUT_DIR = process.env.GSTR_OUTPUT_DIR || path.join(process.cwd(), 'Output');
export const RUNS_DIR = path.join(OUTPUT_DIR, 'PDI runs');

export const MASTER_FILE = path.join(OUTPUT_DIR, 'PDI_Master.xlsx');
export const HISTORY_FILE = path.join(OUTPUT_DIR, 'PDI_History.json');

export const localDate = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

// PDI files and the master keep the last KEEP_DAYS days, today included (by this PC's date); older
// PDI files are deleted and older days drop out of the master.
const KEEP_DAYS = 3;
function keptDays(): string[] {
  return Array.from({ length: KEEP_DAYS }, (_, i) => {
    const d = new Date();
    d.setDate(d.getDate() - i);
    return localDate(d);
  });
}

interface RunFile { file: string; savedAt: Date; rows: ResultRow[]; dealers: ResultRow[] | null }

// Run time comes from the file name (PDI_BlankChassis_<UTC yyyy-mm-dd-hh-mm>[suffix].xlsx, written by
// PDI.spec.ts), not the file's modified time, so opening/re-saving a file doesn't reorder the runs.
function runTimeOf(file: string): Date | null {
  // Older single-dealer runs were named <dealer>_PDI_BlankChassis_<stamp>.xlsx.
  const m = path.basename(file).match(/PDI_BlankChassis_(\d{4})-(\d{2})-(\d{2})-(\d{2})-(\d{2})/i);
  return m ? new Date(Date.UTC(+m[1]!, +m[2]! - 1, +m[3]!, +m[4]!, +m[5]!)) : null;
}

// Run files from Output/PDI runs/, plus any older ones still saved straight in Output/.
function readRunFiles(day: string): RunFile[] {
  return [RUNS_DIR, OUTPUT_DIR]
    .filter((dir) => fs.existsSync(dir))
    .flatMap((dir) => fs.readdirSync(dir)
      .filter((f) => /^(\w+_)?PDI_BlankChassis_.*\.xlsx$/i.test(f) && !f.startsWith('~$'))
      .map((f) => path.join(dir, f)))
    .map((file) => ({ file, savedAt: runTimeOf(file) }))
    .filter((r): r is { file: string; savedAt: Date } => !!r.savedAt && localDate(r.savedAt) === day)
    .sort((a, b) => a.savedAt.getTime() - b.savedAt.getTime() || a.file.localeCompare(b.file))
    .map((r) => {
      const wb = XLSX.readFile(r.file);
      const sheet = (name: string) => wb.Sheets[name] ? XLSX.utils.sheet_to_json<ResultRow>(wb.Sheets[name]!, { raw: false, defval: '' }).map(readPdiRow) : null;
      return { ...r, rows: sheet('PDI Blank Chassis') ?? sheet(wb.SheetNames[0]!) ?? [], dealers: sheet('Dealers') };
    })
    // A dry run saved nothing, so the whole run is left out — including any errors it hit.
    .filter((r) => {
      const dry = r.dealers ? r.dealers.some((d) => d.DryRun === 'Yes') : r.rows.some((row) => String(row.Status).startsWith('Dry run'));
      if (dry) console.log(`   (skipping dry run ${path.basename(r.file)})`);
      return !dry;
    });
}

export interface PdiDay { file: string; results: ResultRow[]; failures: FailedDealer[]; meta: ReportMeta }

// Merges the day's runs into Output/PDI_Daily_<day>.xlsx and returns what the report/email needs,
// or null when there is no real (non-dry) run that day.
export async function buildPdiDaily(day: string): Promise<PdiDay | null> {
  const runs = readRunFiles(day);
  if (!runs.length) return null;
  console.log(`[STEP] ${runs.length} PDI run(s) on ${day}:`);
  runs.forEach((r) => console.log(`   ${path.basename(r.file)}  (${r.rows.length} chassis${r.dealers ? `, ${r.dealers.length} dealer(s) checked` : ''})`));

  // Latest result per dealer+chassis, remembering which run it came from.
  const latest = new Map<string, { row: ResultRow; runIndex: number }>();
  const dealerOrder: string[] = [];
  const addDealer = (code: string) => { if (code && !dealerOrder.includes(code)) dealerOrder.push(code); };
  runs.forEach((run, runIndex) => {
    run.dealers?.filter((d) => d.DryRun !== 'Yes').forEach((d) => addDealer(d.Dealer!));
    run.rows.filter((r) => !String(r.Status).startsWith('Dry run')).forEach((row) => {
      addDealer(row.Dealer!);
      latest.set(`${row.Dealer}|${row.ChassisNo}`, { row: { ...row }, runIndex });
    });
  });

  // An error from an earlier run is cleared if a later real run checked that dealer and the chassis
  // wasn't blank any more (e.g. the Save went through even though the page then timed out).
  for (const { row, runIndex } of latest.values()) {
    if (String(row.Status).startsWith('Updated')) continue;
    const laterCheck = runs.slice(runIndex + 1).find((run) =>
      run.dealers?.some((d) => d.Dealer === row.Dealer && d.DryRun !== 'Yes' && d.Result === 'Checked')
      && !run.rows.some((r) => r.Dealer === row.Dealer && r.ChassisNo === row.ChassisNo));
    if (laterCheck) row.Status = `Updated (no longer blank at a later check; earlier: ${row.Status})`;
  }

  // A dealer counts as failed only if its most recent check failed.
  const failures: FailedDealer[] = [];
  for (const code of dealerOrder) {
    const lastCheck = [...runs].reverse().flatMap((r) => r.dealers ?? []).find((d) => d.Dealer === code && d.DryRun !== 'Yes');
    if (lastCheck?.Result?.startsWith('Failed')) failures.push({ dealer: code, error: lastCheck.Result.replace(/^Failed:\s*/, '') });
  }

  const results = [...latest.values()].map((v) => v.row)
    .sort((a, b) => dealerOrder.indexOf(a.Dealer!) - dealerOrder.indexOf(b.Dealer!));

  // The run files are the only input, so if some were deleted the rebuild would lose their rows. Never
  // replace a daily copy that covers more runs than are left (PDI_REBUILD_DAILY=1 to do it anyway).
  const dailyPath = path.join(OUTPUT_DIR, `PDI_Daily_${day}.xlsx`);
  if (fs.existsSync(dailyPath) && process.env.PDI_REBUILD_DAILY !== '1') {
    const wb = XLSX.readFile(dailyPath);
    const heading = String(wb.Sheets[wb.SheetNames[0]!]?.A1?.v ?? '');
    const before = Number(heading.match(/·\s*(\d+)\s*run\(s\)/)?.[1] ?? 0);
    if (before > runs.length) {
      throw new Error(`${path.basename(dailyPath)} covers ${before} run(s) but only ${runs.length} run file(s) from ${day} are left in ${RUNS_DIR} — not replacing it. Put the missing PDI_BlankChassis_*.xlsx files back, or set PDI_REBUILD_DAILY=1 to rebuild from what is left.`);
    }
  }

  const title = `PDI daily summary · ${day} · ${runs.length} run(s)`;
  const file = await writePdiWorkbook(dailyPath, {
    title, subtitle: `${dealerOrder.length} dealer(s) · latest result per chassis across the day's runs`,
    dryRun: false, dealers: dealerOrder, failed: failures, rows: results, rowsSheet: 'PDI Daily',
  });
  console.log(`[STEP] Combined ${results.length} chassis across ${dealerOrder.length} dealer(s) into ${file}`);

  const withoutDealerList = runs.filter((r) => !r.dealers).length;
  if (withoutDealerList) {
    console.log(`[WARN] ${withoutDealerList} earlier run file(s) predate the dealer list — dealers with no blank rows in those runs aren't shown.`);
  }

  deleteOldPdiFiles();
  const history = await refreshPdiMaster();

  return {
    file, results, failures,
    meta: {
      dryRun: false,
      startedAt: runs[0]!.savedAt,
      finishedAt: runs[runs.length - 1]!.savedAt,
      dealers: dealerOrder,
      vinFile: vinFilePath(),
      title,
      history,
    },
  };
}

// Deletes PDI files dated before the oldest kept day: run files and email markers in Output/PDI runs/, and the
// PDI_Daily_<date> copies in Output/. Only PDI files are touched; GSTR and other outputs stay.
function deleteOldPdiFiles() {
  const oldest = keptDays()[KEEP_DAYS - 1]!;
  const dayOf = (f: string) => runTimeOf(f) ? localDate(runTimeOf(f)!) : f.match(/^(?:PDI_Daily|emailed)_(\d{4}-\d{2}-\d{2})/i)?.[1];
  const old = [RUNS_DIR, OUTPUT_DIR]
    .filter((dir) => fs.existsSync(dir))
    .flatMap((dir) => fs.readdirSync(dir)
      .filter((f) => /^((\w+_)?PDI_BlankChassis_|PDI_Daily_|emailed_)/i.test(f) && /\.(xlsx|html|txt)$/i.test(f))
      .filter((f) => { const d = dayOf(f); return !!d && d < oldest; })
      .map((f) => path.join(dir, f)));
  let deleted = 0;
  for (const file of old) {
    try { fs.unlinkSync(file); deleted++; } catch (err) {
      console.log(`[WARN] Could not delete old PDI file ${path.basename(file)}: ${String((err as Error).message).slice(0, 80)}`);
    }
  }
  if (deleted) console.log(`[STEP] Deleted ${deleted} PDI file(s) from before ${oldest}.`);
}

// Output/PDI_Master.xlsx (+ PDI_Master.html dashboard): one sheet per kept day, copied from that day's
// PDI_Daily file. A kept day whose PDI_Daily file is gone keeps the sheet already in the master, so a
// day is never lost from the master while it is still inside the window.
//
// Every day that has a PDI_Daily file is also saved into Output/PDI_History.json, which keeps all days
// (see pdiHistory.ts); the master's Tracking sheet and the email show totals from it.
async function refreshPdiMaster(): Promise<PdiHistory> {
  const sheetRows = (file: string, name: string) => {
    const sheet = XLSX.readFile(file).Sheets[name];
    return sheet ? XLSX.utils.sheet_to_json<ResultRow>(sheet, { raw: false, defval: '' }).map(readPdiRow) : null;
  };
  // Dealers whose check failed, from the daily file's Summary sheet (Dealer in A, "Failed: …" in F).
  const failedDealers = (file: string) => {
    const sheet = XLSX.readFile(file).Sheets['Summary'];
    const rows = sheet ? XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: '' }) : [];
    return rows.filter((r) => String(r[5] ?? '').startsWith('Failed')).map((r) => String(r[0]));
  };
  const fromDaily: { day: string; rows: ResultRow[]; failed: string[] }[] = [];
  const days = keptDays()
    .map((day) => {
      const daily = path.join(OUTPUT_DIR, `PDI_Daily_${day}.xlsx`);
      if (fs.existsSync(daily)) {
        const entry = { day, rows: sheetRows(daily, 'PDI Daily') ?? [], failed: failedDealers(daily) };
        fromDaily.push(entry);
        return entry;
      }
      const kept = fs.existsSync(MASTER_FILE) ? sheetRows(MASTER_FILE, day) : null;
      return kept ? { day, rows: kept } : null;
    })
    .filter((d): d is { day: string; rows: ResultRow[] } => !!d);
  const history = recordPdiHistory(HISTORY_FILE, fromDaily);
  if (days.length && await writePdiMaster(MASTER_FILE, days, history)) {
    console.log(`[STEP] PDI master updated: ${MASTER_FILE} (${days.map((d) => `${d.day}: ${d.rows.length} chassis`).join(', ')})`);
    writePdiMasterDashboard(days, MASTER_FILE);
  }
  if (history.latest) {
    console.log(`[STEP] PDI history since ${history.since}: ${history.latest.totalFound} chassis found, ${history.latest.totalUpdated} updated, ${history.latest.openMissing} still with no usable VIN data.`);
  }
  return history;
}

// Marks the day's email as sent so it goes out once, not after every rerun of the daily report.
export const emailedMarker = (day: string) => path.join(RUNS_DIR, `emailed_${day}.txt`);
