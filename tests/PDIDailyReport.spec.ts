import { test } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { sendPdiReport, problemsOf, ResultRow, FailedDealer } from './utils/pdiReport';

// Daily PDI summary: merges every PDI run saved in Output today (PDI_BlankChassis_*.xlsx written by
// tests/PDI.spec.ts) into one per-dealer report and emails it, with the interactive dashboard and a
// combined Excel attached. No browser, no DMS login — it only reads the saved result files.
//
// Merging rules:
//   - A chassis seen in several runs keeps its latest result.
//   - An earlier error/skip is shown as "Updated (no longer blank at a later check)" when a later run
//     checked that dealer again and the chassis was no longer blank.
//   - Dry-run rows are left out (nothing was saved).
//
// Settings (optional):
//   PDI_REPORT_DATE  yyyy-mm-dd to report on (default: today, by the files' saved time)
//   PDI_SEND_EMAIL   0 = build the report files but don't email them
const OUTPUT_DIR = process.env.GSTR_OUTPUT_DIR || path.join(process.cwd(), 'Output');
const VIN_FILE = process.env.PDI_VIN_FILE || path.join(__dirname, '..', 'resources', 'VIN Details (1).xlsx');

const localDate = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

interface RunFile { file: string; savedAt: Date; rows: ResultRow[]; dealers: ResultRow[] | null }

// Run time comes from the file name (PDI_BlankChassis_<UTC yyyy-mm-dd-hh-mm>[suffix].xlsx, written by
// PDI.spec.ts), not the file's modified time, so opening/re-saving a file doesn't reorder the runs.
function runTimeOf(file: string): Date | null {
  // Older single-dealer runs were named <dealer>_PDI_BlankChassis_<stamp>.xlsx.
  const m = path.basename(file).match(/PDI_BlankChassis_(\d{4})-(\d{2})-(\d{2})-(\d{2})-(\d{2})/i);
  return m ? new Date(Date.UTC(+m[1]!, +m[2]! - 1, +m[3]!, +m[4]!, +m[5]!)) : null;
}

function readRunFiles(day: string): RunFile[] {
  if (!fs.existsSync(OUTPUT_DIR)) return [];
  return fs.readdirSync(OUTPUT_DIR)
    .filter((f) => /^(\w+_)?PDI_BlankChassis_.*\.xlsx$/i.test(f))
    .map((f) => path.join(OUTPUT_DIR, f))
    .map((file) => ({ file, savedAt: runTimeOf(file) }))
    .filter((r): r is { file: string; savedAt: Date } => !!r.savedAt && localDate(r.savedAt) === day)
    .sort((a, b) => a.savedAt.getTime() - b.savedAt.getTime() || a.file.localeCompare(b.file))
    .map((r) => {
      const wb = XLSX.readFile(r.file);
      const sheet = (name: string) => wb.Sheets[name] ? XLSX.utils.sheet_to_json<ResultRow>(wb.Sheets[name]!, { raw: false, defval: '' }) : null;
      return { ...r, rows: sheet('PDI Blank Chassis') ?? sheet(wb.SheetNames[0]!) ?? [], dealers: sheet('Dealers') };
    })
    // A dry run saved nothing, so the whole run is left out — including any errors it hit.
    .filter((r) => {
      const dry = r.dealers ? r.dealers.some((d) => d.DryRun === 'Yes') : r.rows.some((row) => String(row.Status).startsWith('Dry run'));
      if (dry) console.log(`   (skipping dry run ${path.basename(r.file)})`);
      return !dry;
    });
}

// Sends an email, so only when asked for: npm run report:pdi-daily sets PDI_RUN=1.
test.skip(process.env.PDI_RUN !== '1', 'Daily PDI report runs only via npm run report:pdi-daily (sets PDI_RUN=1).');

test('Should email the combined PDI report for the day', async () => {
  const day = process.env.PDI_REPORT_DATE || localDate(new Date());
  const runs = readRunFiles(day);
  if (!runs.length) throw new Error(`No PDI result files from ${day} in ${OUTPUT_DIR}.`);
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

  let dailyFile = path.join(OUTPUT_DIR, `PDI_Daily_${day}.xlsx`);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(results), 'PDI Daily');
  try {
    XLSX.writeFile(workbook, dailyFile);
  } catch (err) {
    // Usually yesterday's/earlier copy is open in Excel (Windows locks it) — save beside it instead.
    const fallback = dailyFile.replace(/\.xlsx$/, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(dailyFile)} (${String((err as Error).message).slice(0, 80)}) — is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    XLSX.writeFile(workbook, fallback);
    dailyFile = fallback;
  }
  console.log(`[STEP] Combined ${results.length} chassis across ${dealerOrder.length} dealer(s) into ${dailyFile}`);

  const withoutDealerList = runs.filter((r) => !r.dealers).length;
  if (withoutDealerList) {
    console.log(`[WARN] ${withoutDealerList} earlier run file(s) predate the dealer list — dealers with no blank rows in those runs aren't shown.`);
  }

  await sendPdiReport(results, failures, {
    dryRun: false,
    startedAt: runs[0]!.savedAt,
    finishedAt: runs[runs.length - 1]!.savedAt,
    dealers: dealerOrder,
    vinFile: VIN_FILE,
    title: `PDI daily summary · ${day} · ${runs.length} run(s)`,
  }, dailyFile);

  const problems = problemsOf(results, failures, false);
  if (problems.length) throw new Error(`${problems.length} problem(s) today — daily report not emailed:\n  ${problems.slice(0, 20).join('\n  ')}`);
});
