import * as path from 'node:path';
import ExcelJS from 'exceljs';
import { outcomeOf, ResultRow, FailedDealer } from './pdiReport';

// Formatted Excel for the PDI run (tests/PDI.spec.ts) and the daily summary (PDIDailyReport.spec.ts):
// a Summary sheet first, then one row per chassis, then (for a run) every dealer checked.
// Columns carry friendly headers; readPdiRow() maps them back so saved files can be read again.

const HEADERS: Record<string, string> = {
  Dealer: 'Dealer', ChassisNo: 'Chassis No', BlankFields: 'Was blank', EngineNo: 'Engine No',
  BatteryDetails: 'Battery', VehicleChargerNo: 'Charger', Status: 'Status',
  DryRun: 'Dry run', BlankFound: 'Blank found', Result: 'Result',
};
const KEY_OF = Object.fromEntries(Object.entries(HEADERS).map(([k, v]) => [v, k]));
const CHASSIS_KEYS = ['Dealer', 'ChassisNo', 'BlankFields', 'EngineNo', 'BatteryDetails', 'VehicleChargerNo', 'Status'];
const MONO = new Set(['ChassisNo', 'EngineNo', 'BatteryDetails', 'VehicleChargerNo']);

// Rows read back from a saved file (new friendly headers or older raw ones) use the raw keys.
export function readPdiRow(row: ResultRow): ResultRow {
  return Object.fromEntries(Object.entries(row).map(([k, v]) => [KEY_OF[k] ?? k, v]));
}

const BRAND = 'FF16325C';
const INK2 = 'FF52514E';
const LINE = 'FFE4E3DE';
const TONES = {
  good: { fill: 'FFE6F4EA', font: 'FF137333' },
  info: { fill: 'FFE8F0FE', font: 'FF1A56B8' },
  warn: { fill: 'FFFEF4D8', font: 'FF8A5A00' },
  crit: { fill: 'FFFCE8E6', font: 'FFB3261E' },
};
const TONE_OF = { updated: TONES.good, dryrun: TONES.info, missing: TONES.warn, error: TONES.crit } as const;
const fill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const thin = { style: 'thin' as const, color: { argb: LINE } };
const BORDER: Partial<ExcelJS.Borders> = { top: thin, bottom: thin, left: thin, right: thin };

export interface PdiWorkbook {
  title: string;            // e.g. "PDI vehicle details"
  subtitle: string;         // e.g. "29 Sep 2026, 14:53 · 85 dealer(s) · VIN sheet: VIN Details.xlsx"
  dryRun: boolean;
  dealers: string[];        // every dealer checked, in report order
  failed: FailedDealer[];
  rows: ResultRow[];        // one per chassis
  rowsSheet: string;        // name of the chassis sheet (the daily report reads it by name)
  dealersSheet?: ResultRow[]; // run files only: one row per dealer checked
}

// A plain table: branded header, frozen, filterable, fitted column widths, light banding.
function addTable(wb: ExcelJS.Workbook, name: string, keys: string[], rows: ResultRow[], statusKey: string) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }], properties: { defaultRowHeight: 18 } });
  ws.columns = keys.map((k) => ({
    header: HEADERS[k] ?? k, key: k,
    width: Math.min(60, Math.max(10, (HEADERS[k] ?? k).length + 2, ...rows.map((r) => String(r[k] ?? '').length + 2))),
  }));
  rows.forEach((r) => ws.addRow(Object.fromEntries(keys.map((k) => [k, r[k] ?? '']))));

  const header = ws.getRow(1);
  header.height = 22;
  header.eachCell((c) => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = fill(BRAND);
    c.alignment = { vertical: 'middle' };
    c.border = BORDER;
  });
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: keys.length } };

  ws.eachRow((row, n) => {
    if (n === 1) return;
    row.eachCell({ includeEmpty: true }, (c, col) => {
      const key = keys[col - 1]!;
      c.border = BORDER;
      c.alignment = { vertical: 'middle', wrapText: key === statusKey };
      if (MONO.has(key)) c.font = { name: 'Consolas' };
      if (n % 2 === 0) c.fill = fill('FFF7F7F4');
    });
    // Colour the status cell by outcome.
    const cell = row.getCell(keys.indexOf(statusKey) + 1);
    const text = String(cell.value ?? '');
    const tone = statusKey === 'Result'
      ? (text.startsWith('Failed') ? TONES.crit : TONES.good)
      : TONE_OF[outcomeOf({ Status: text })];
    cell.fill = fill(tone.fill);
    cell.font = { bold: true, color: { argb: tone.font } };
  });
  return ws;
}

function addSummary(wb: ExcelJS.Workbook, book: PdiWorkbook) {
  const ws = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 14 }, { width: 12 }, { width: 14 }, { width: 16 }, { width: 12 }, { width: 50 }];

  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = `River DMS · ${book.title}${book.dryRun ? ' · DRY RUN (nothing saved)' : ''}`;
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  ws.getCell('A1').fill = fill(BRAND);
  ws.getCell('A1').alignment = { vertical: 'middle', indent: 1 };
  ws.getRow(1).height = 32;
  ws.mergeCells('A2:F2');
  ws.getCell('A2').value = book.subtitle;
  ws.getCell('A2').font = { color: { argb: INK2 } };
  ws.getCell('A2').alignment = { indent: 1 };

  // Headline figures.
  const count = (o: string) => book.rows.filter((r) => outcomeOf(r) === o).length;
  const done = book.dryRun ? count('dryrun') : count('updated');
  const tiles: [string, number, typeof TONES.good][] = [
    ['Blank chassis found', book.rows.length, { fill: 'FFE9EEF6', font: BRAND }],
    [book.dryRun ? 'Filled (dry run)' : 'Updated', done, book.dryRun ? TONES.info : TONES.good],
    ['No usable VIN data', count('missing'), TONES.warn],
    ['Errors', count('error') + book.failed.length, TONES.crit],
  ];
  tiles.forEach(([label, value, tone], i) => {
    const col = [1, 2, 4, 5][i]!; // A, B, D, E — C is wide enough for its own label
    const v = ws.getCell(4, col);
    const l = ws.getCell(5, col);
    v.value = value;
    v.font = { bold: true, size: 18, color: { argb: tone.font } };
    l.value = label;
    l.font = { size: 9, color: { argb: INK2 } };
    [v, l].forEach((c) => { c.fill = fill(tone.fill); c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
  });
  ws.getRow(4).height = 30;
  ws.getRow(5).height = 26;

  // By dealer: only dealers with something to show; the rest share one line.
  const head = ws.getRow(7);
  head.values = ['Dealer', 'Blank', book.dryRun ? 'Filled' : 'Updated', 'No usable VIN data', 'Errors', 'Result'];
  head.height = 22;
  head.eachCell((c) => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = fill(BRAND); c.border = BORDER; c.alignment = { vertical: 'middle' }; });

  let n = 8;
  const clean: string[] = [];
  for (const dealer of book.dealers) {
    const mine = book.rows.filter((r) => r.Dealer === dealer);
    const failure = book.failed.find((f) => f.dealer === dealer);
    if (!mine.length && !failure) { clean.push(dealer); continue; }
    const c = (o: string) => mine.filter((r) => outcomeOf(r) === o).length;
    const d = book.dryRun ? c('dryrun') : c('updated');
    const problems = c('missing') + c('error');
    const result = failure ? `Failed: ${failure.error}` : problems ? `${problems} need attention` : `All ${d} done`;
    const row = ws.getRow(n++);
    row.values = [dealer, mine.length, d, c('missing'), c('error') + (failure ? 1 : 0), result];
    const tone = failure || problems ? TONES.crit : TONES.good;
    row.eachCell({ includeEmpty: true }, (cell, col) => {
      cell.border = BORDER;
      cell.alignment = { vertical: 'middle', wrapText: col === 6 };
      if (col === 6) { cell.fill = fill(tone.fill); cell.font = { bold: true, color: { argb: tone.font } }; }
    });
  }
  if (clean.length) {
    ws.mergeCells(n, 1, n, 6);
    const cell = ws.getCell(n, 1);
    cell.value = `No blank PDI rows (${clean.length} dealers): ${clean.join(', ')}`;
    cell.font = { color: { argb: INK2 } };
    cell.alignment = { wrapText: true, vertical: 'top' };
    ws.getRow(n).height = Math.max(18, Math.ceil(cell.value.length / 95) * 15);
    n++;
  }
  ws.getCell(n + 1, 1).value = 'Full details are on the next sheet(s). Use the filter arrows in the header row to narrow them down.';
  ws.getCell(n + 1, 1).font = { italic: true, size: 9, color: { argb: INK2 } };
}

// Writes the workbook; if the file is locked (usually open in Excel), saves beside it instead.
// Returns the path actually written.
export async function writePdiWorkbook(filePath: string, book: PdiWorkbook): Promise<string> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'E2E DMS automation';
  addSummary(wb, book);
  addTable(wb, book.rowsSheet, CHASSIS_KEYS, book.rows, 'Status');
  if (book.dealersSheet) addTable(wb, 'Dealers', ['Dealer', 'DryRun', 'BlankFound', 'Result'], book.dealersSheet, 'Result');
  try {
    await wb.xlsx.writeFile(filePath);
    return filePath;
  } catch (err) {
    const fallback = filePath.replace(/\.xlsx$/i, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${String((err as Error).message).slice(0, 80)}) — is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    await wb.xlsx.writeFile(fallback);
    return fallback;
  }
}
