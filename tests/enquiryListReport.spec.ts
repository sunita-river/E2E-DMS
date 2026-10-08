import { test, expect } from '@playwright/test';
import { LoginPage } from './models/LoginPage';
import { EnquiryListReportPage } from './models/EnquiryListReportPage';
import { ProductWiseSalesDetailsReportPage } from './models/ProductWiseSalesDetailsReportPage';
import { missingColumns, blankValues, datesOutsideRange, parseReportDate } from './utils/reportChecks';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';

// CREDENTIALS_FILE: optional path to a credentials.json (e.g. a CI secrets file)
const credentials = JSON.parse(fs.readFileSync(
  process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json'), 'utf8'));

// Report range: 7 September 2026 to yesterday, unless ENQ_FROM / ENQ_TO (dd-mm-yyyy) are set.
// The same range is used for Enquiry List and both Product-wise Sales Details reports.
function parseDateEnv(name: string, fallback: Date): Date {
  const v = process.env[name];
  if (!v) return fallback;
  const m = v.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!m) throw new Error(`${name} must be dd-mm-yyyy (got "${v}")`);
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}
const REPORT_FROM = parseDateEnv('ENQ_FROM', new Date(2026, 8, 7)); // months are 0-based: 8 = September
const yesterday = (() => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate() - 1); })();
const REPORT_TO = parseDateEnv('ENQ_TO', yesterday);
if (REPORT_FROM > REPORT_TO) throw new Error('ENQ_FROM is after ENQ_TO.');
const REQUIRED_COLUMNS = ['Dealer Code', 'Dealer name', 'Lead No', 'Created Date', 'Lead Source', 'Lead Status', 'Product name'];
const OUTPUT_DIR = process.env.ENQ_OUTPUT_DIR || path.join(process.cwd(), 'Output');
const dateLabel = (d: Date) =>
  `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
// Date-picker arguments: day, full month name, year.
const pickerArgs = (d: Date): [string, string, number] => [String(d.getDate()), d.toLocaleString('en-US', { month: 'long' }), d.getFullYear()];

// Writes a workbook, saving beside the target if it is locked (usually open in Excel).
function writeWorkbook(workbook: XLSX.WorkBook, filePath: string): string {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  try {
    XLSX.writeFile(workbook, filePath);
    return filePath;
  } catch (err) {
    const fallback = filePath.replace(/\.xlsx$/, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${String((err as Error).message).slice(0, 80)}) — is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    XLSX.writeFile(workbook, fallback);
    return fallback;
  }
}

// Saves the full Enquiry List (every row, every column, as shown in the report)
// to Output/Enquiry_List_<from>_<to>.xlsx.
function saveEnquiryList(rows: Record<string, string>[]): string {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'Enquiry List');
  return writeWorkbook(workbook, path.join(OUTPUT_DIR, `Enquiry_List_${dateLabel(REPORT_FROM)}_${dateLabel(REPORT_TO)}.xlsx`));
}

// Reports don't use consistent header names for the same field (Enquiry List has "Dealer Code" /
// "Dealer name", Product-wise Sales Details has "DealerCode" / "DealerName"), so a field is looked up
// by trying each known name, ignoring whitespace/case. Blank values count as not found.
function getField(row: Record<string, string>, ...names: string[]): string {
  const normalize = (s: string) => s.replace(/\s+/g, '').toLowerCase();
  const normalizedRow = new Map(Object.entries(row).map(([k, v]) => [normalize(k), v]));
  for (const name of names) {
    const value = normalizedRow.get(normalize(name));
    if (value !== undefined && value.trim() !== '') return value.trim();
  }
  return '';
}

// Column check for reports whose header names vary: each entry is one field and the names it may have.
function missingFields(rows: Record<string, string>[], fields: string[][]): string[] {
  if (!rows.length) return [];
  const normalize = (s: string) => s.replace(/\s+/g, '').toLowerCase();
  const present = new Set(Object.keys(rows[0]!).map(normalize));
  const missing = fields.filter((names) => !names.some((n) => present.has(normalize(n))));
  return missing.length ? [`Missing column(s): ${missing.map((n) => n[0]).join(', ')}`] : [];
}

// One report feeding the summary. dateField is the date the report was run on (Invoice Date for the
// Invoice date type, Booking Date for Booking), so a booking that is already invoiced is still counted
// on its booking date. countBy counts distinct values (a booking can span several product rows).
type Source = { label: string; rows: Record<string, string>[]; dateField: string[]; countBy?: string[] };

const DEALER_CODE = ['Dealer Code', 'DealerCode'];
const DEALER_NAME = ['Dealer name', 'DealerName'];
const BOOKING_NO = ['Booking No', 'BookingNo'];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const ordinal = (d: Date) => {
  const day = d.getDate();
  const suffix = day >= 11 && day <= 13 ? 'th' : (['th', 'st', 'nd', 'rd'][day % 10] ?? 'th');
  return `${day}${suffix} ${MONTHS[d.getMonth()]}`;
};

type Group = { date: Date; dealerCode: string; dealerName: string; counts: Map<string, Set<string> | number> };

// Groups every source by Date + Dealer Code, one count column per source, oldest date first.
function buildSummary(sources: Source[]) {
  const groups = new Map<string, Group>();
  for (const { label, rows, dateField, countBy } of sources) {
    for (const row of rows) {
      const date = parseReportDate(getField(row, ...dateField));
      if (!date) continue; // reported by the data checks
      const dealerCode = getField(row, ...DEALER_CODE);
      const key = `${date.getTime()}|${dealerCode}`;
      let g = groups.get(key);
      if (!g) groups.set(key, (g = { date, dealerCode, dealerName: getField(row, ...DEALER_NAME), counts: new Map() }));
      if (!g.dealerName) g.dealerName = getField(row, ...DEALER_NAME);
      if (countBy) {
        const set = (g.counts.get(label) as Set<string> | undefined) ?? new Set<string>();
        set.add(getField(row, ...countBy));
        g.counts.set(label, set);
      } else {
        g.counts.set(label, ((g.counts.get(label) as number | undefined) ?? 0) + 1);
      }
    }
  }
  const count = (g: Group, label: string) => {
    const c = g.counts.get(label);
    return c === undefined ? 0 : typeof c === 'number' ? c : c.size;
  };
  return [...groups.values()]
    .sort((a, b) => a.date.getTime() - b.date.getTime() || a.dealerCode.localeCompare(b.dealerCode))
    .map((g) => ({ date: g.date, dealerCode: g.dealerCode, dealerName: g.dealerName, counts: sources.map((s) => count(g, s.label)) }));
}

type SummaryRow = ReturnType<typeof buildSummary>[number];

function withTotal(rows: (string | number)[][], firstCount: number): (string | number)[][] {
  if (!rows.length) return rows;
  const total = rows[0]!.map((_, i) => (i === 0 ? 'Total' : i < firstCount ? '' : rows.reduce((s, r) => s + Number(r[i]), 0)));
  return [...rows, total];
}

// Totals per day (summed over dealers) and per dealer (summed over days).
function rollUp(summary: SummaryRow[], by: (r: SummaryRow) => string) {
  const out = new Map<string, { first: SummaryRow; counts: number[] }>();
  for (const r of summary) {
    const e = out.get(by(r));
    if (e) r.counts.forEach((c, i) => (e.counts[i]! += c));
    else out.set(by(r), { first: r, counts: [...r.counts] });
  }
  return [...out.values()];
}

function toCsv(rows: (string | number)[][]): string {
  const escape = (v: string | number) => (/[",\r\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v));
  return rows.map((r) => r.map(escape).join(',')).join('\r\n');
}

function printTable(rows: (string | number)[][]) {
  const widths = rows[0]!.map((_, i) => Math.max(...rows.map((r) => String(r[i]).length)));
  const line = (cells: (string | number)[]) => '│ ' + cells.map((c, i) => String(c).padEnd(widths[i]!)).join(' │ ') + ' │';
  const border = (l: string, m: string, r: string) => l + widths.map((w) => '─'.repeat(w + 2)).join(m) + r;
  console.log(border('┌', '┬', '┐'));
  rows.forEach((r, i) => { console.log(line(r)); if (i === 0) console.log(border('├', '┼', '┤')); });
  console.log(border('└', '┴', '┘'));
}

test.describe('MFR Vehicles Reports - Enquiry List and Product-wise Sales Details', () => {
  test.use({ storageState: { cookies: [], origins: [] } });

  test('Should view Enquiry List, Invoice and Booking reports and summarise them by date and dealer', async ({ page }) => {
    const loginPage = new LoginPage(page);
    const enquiryListReportPage = new EnquiryListReportPage(page);
    const salesPage = new ProductWiseSalesDetailsReportPage(page);

    await loginPage.goto();
    await loginPage.login(credentials.dealervalidUser);
    await loginPage.verifyDashboard();

    // --- Enquiry List ---
    const enquiryRows = await test.step('Enquiry List', async () => {
      await enquiryListReportPage.navigateToEnquiryList();
      await enquiryListReportPage.setFromDate(...pickerArgs(REPORT_FROM));
      await enquiryListReportPage.setToDate(...pickerArgs(REPORT_TO));
      const reportPage = await enquiryListReportPage.viewReport();
      await enquiryListReportPage.openColumnSelector(reportPage);
      const rows = await enquiryListReportPage.getReportRows(reportPage);
      console.log(`[STEP] Enquiry List record count (excluding header): ${rows.length}`);
      await enquiryListReportPage.showRecordCountPopup(reportPage, rows.length);
      await enquiryListReportPage.closeReport(reportPage, 'Enquiry List');
      return rows;
    });
    if (enquiryRows.length === 0) {
      test.info().annotations.push({ type: 'warning', description: 'No Enquiry Found for the selected date range.' });
    }
    expect(enquiryRows.length).toBeGreaterThan(0);
    const savedPath = saveEnquiryList(enquiryRows);
    console.log(`[STEP] Enquiry List (${enquiryRows.length} rows) saved to ${savedPath}`);

    // Data checks run after saving, so the file is there to inspect when one fails. Soft, so the
    // sales reports and the summary are still produced. Lead No is not unique: a lead can appear
    // on more than one row.
    const enquiryProblems = [
      ...missingColumns(enquiryRows, REQUIRED_COLUMNS),
      ...blankValues(enquiryRows, ['Dealer Code', 'Lead No', 'Created Date']),
      ...datesOutsideRange(enquiryRows, 'Created Date', REPORT_FROM, REPORT_TO),
    ];
    expect.soft(enquiryProblems, `Enquiry List data problems:\n${enquiryProblems.join('\n')}`).toEqual([]);

    // --- Product-wise Sales Details, Invoice date type (the report's default), same session ---
    const invoiceRows = await test.step('Product-wise Sales Details (Invoice)', async () => {
      await salesPage.navigateToProductWiseSalesDetails();
      await salesPage.setFromDate(...pickerArgs(REPORT_FROM));
      await salesPage.setToDate(...pickerArgs(REPORT_TO));
      const reportPage = await salesPage.viewReport();
      await salesPage.openColumnSelector(reportPage);
      const rows = await salesPage.getReportRows(reportPage);
      console.log(`[STEP] Product-wise Sales Details, Invoice date type: ${rows.length} rows`);
      await salesPage.closeReport(reportPage, 'Product-wise Sales Details (Invoice)');
      return rows;
    });

    // --- Product-wise Sales Details, Booking date type ---
    // The date type can't be switched in place: the main form has to be reopened from the dashboard
    // before the date type dropdown and View Report are usable again.
    const bookingRows = await test.step('Product-wise Sales Details (Booking)', async () => {
      await page.getByRole('banner').locator('a[href*="UserDashboard.aspx"]').first().click();
      await page.waitForLoadState('load');
      await salesPage.navigateToProductWiseSalesDetails();
      await salesPage.setFromDate(...pickerArgs(REPORT_FROM));
      await salesPage.setToDate(...pickerArgs(REPORT_TO));
      await salesPage.setDateType('B');
      const reportPage = await salesPage.viewReport();
      await salesPage.openColumnSelector(reportPage);
      const rows = await salesPage.getReportRows(reportPage);
      console.log(`[STEP] Product-wise Sales Details, Booking date type: ${rows.length} rows`);
      await salesPage.closeReport(reportPage, 'Product-wise Sales Details (Booking)');
      return rows;
    });
    if (bookingRows.length === 0) {
      test.info().annotations.push({ type: 'warning', description: 'No bookings found for the selected date range.' });
    }
    expect.soft(bookingRows.length, 'Booking report should have rows').toBeGreaterThan(0);

    const INVOICE_DATE = ['Invoice Date', 'InvoiceDate'];
    const BOOKING_DATE = ['Booking Date', 'BookingDate'];
    for (const [label, rows, dateField] of [['Invoice', invoiceRows, INVOICE_DATE], ['Booking', bookingRows, BOOKING_DATE]] as const) {
      const problems = missingFields(rows, [DEALER_CODE, DEALER_NAME, BOOKING_NO, [...dateField]]);
      const dated = rows.map((r) => ({ ...r, __date: getField(r, ...dateField) }));
      if (!problems.length) {
        problems.push(...blankValues(dated.map((r) => ({ 'Booking No': getField(r, ...BOOKING_NO), 'Dealer Code': getField(r, ...DEALER_CODE) })), ['Booking No', 'Dealer Code']));
        problems.push(...datesOutsideRange(dated, '__date', REPORT_FROM, REPORT_TO).map((p) => p.replace('__date', dateField[0])));
      }
      expect.soft(problems, `${label} report data problems:\n${problems.join('\n')}`).toEqual([]);
    }

    // --- Dealers, and the combined Enquiry / Invoice / Booking summary ---
    const sources: Source[] = [
      { label: 'Enquiry', rows: enquiryRows, dateField: ['Created Date'] },
      { label: 'Invoice', rows: invoiceRows, dateField: INVOICE_DATE, countBy: BOOKING_NO },
      { label: 'Booking', rows: bookingRows, dateField: BOOKING_DATE, countBy: BOOKING_NO },
    ];
    for (const s of sources) {
      const dealers = new Set(s.rows.map((r) => getField(r, ...DEALER_CODE)).filter(Boolean));
      console.log(`[STEP] ${s.label}: ${s.rows.length} rows from ${dealers.size} dealers`);
    }

    const labels = sources.map((s) => s.label);
    const summary = buildSummary(sources);
    const byDateDealer = withTotal(summary.map((r, i) => [i + 1, ordinal(r.date), r.dealerCode, r.dealerName, ...r.counts]), 4);
    const byDate = withTotal(rollUp(summary, (r) => String(r.date.getTime())).map((e) => [ordinal(e.first.date), ...e.counts]), 1);
    const byDealer = withTotal(
      rollUp(summary, (r) => r.dealerCode).sort((a, b) => a.first.dealerCode.localeCompare(b.first.dealerCode))
        .map((e) => [e.first.dealerCode, e.first.dealerName, ...e.counts]), 2);

    const dateDealerHeader = ['Serial No', 'Date', 'Dealer Code', 'Dealer Name', ...labels];
    const dateHeader = ['Date', ...labels];
    const dealerHeader = ['Dealer Code', 'Dealer Name', ...labels];

    console.log(`[STEP] ${labels.join(' / ')} by date:`);
    printTable([dateHeader, ...byDate]);

    // The summary's totals must match the reports: every enquiry row, every distinct booking.
    const totals = byDateDealer[byDateDealer.length - 1]!.slice(4).map(Number);
    const expected = sources.map((s) => (s.countBy
      ? new Set(s.rows.map((r) => `${getField(r, ...DEALER_CODE)}|${getField(r, ...s.countBy!)}|${parseReportDate(getField(r, ...s.dateField))?.getTime()}`)).size
      : s.rows.length));
    expect.soft(totals, `Summary totals (${labels.join(', ')}) should match the reports`).toEqual(expected);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([dateHeader, ...byDate]), 'By Date');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([dealerHeader, ...byDealer]), 'By Dealer');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([dateDealerHeader, ...byDateDealer]), 'By Date and Dealer');
    const summaryPath = writeWorkbook(workbook, path.join(OUTPUT_DIR, `Enquiry_Summary_${dateLabel(REPORT_FROM)}_${dateLabel(REPORT_TO)}.xlsx`));
    console.log(`[STEP] Summary (${summary.length} date/dealer rows) saved to ${summaryPath}`);

    // The gmail reporter pastes .csv attachments into the email body. Only the per-day table goes
    // there: the date/dealer table runs to ~2,000 rows, which Gmail would clip. It is in the Excel file.
    await test.info().attach('enquiry-invoice-booking-by-date.csv', { body: toCsv([dateHeader, ...byDate]), contentType: 'text/csv' });
  });
});
