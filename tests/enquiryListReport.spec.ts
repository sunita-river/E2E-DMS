import { test, expect } from '@playwright/test';
import { LoginPage } from './models/LoginPage';
import { EnquiryListReportPage } from './models/EnquiryListReportPage';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';

// CREDENTIALS_FILE: optional path to a credentials.json (e.g. a CI secrets file)
const credentials = JSON.parse(fs.readFileSync(
  process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json'), 'utf8'));

// Report range: 7 to 29 September 2026, unless ENQ_FROM / ENQ_TO (dd-mm-yyyy) are set.
function parseDateEnv(name: string, fallback: Date): Date {
  const v = process.env[name];
  if (!v) return fallback;
  const m = v.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!m) throw new Error(`${name} must be dd-mm-yyyy (got "${v}")`);
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}
const REPORT_FROM = parseDateEnv('ENQ_FROM', new Date(2026, 8, 7)); // months are 0-based: 8 = September
const REPORT_TO = parseDateEnv('ENQ_TO', new Date(2026, 8, 29));
if (REPORT_FROM > REPORT_TO) throw new Error('ENQ_FROM is after ENQ_TO.');
const OUTPUT_DIR = process.env.ENQ_OUTPUT_DIR || path.join(process.cwd(), 'Output');
const dateLabel = (d: Date) =>
  `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
// Date-picker arguments: day, full month name, year.
const pickerArgs = (d: Date): [string, string, number] => [String(d.getDate()), d.toLocaleString('en-US', { month: 'long' }), d.getFullYear()];

// Saves the full Enquiry List (every row, every column, as shown in the report)
// to Output/Enquiry_List_<from>_<to>.xlsx.
function saveEnquiryList(rows: Record<string, string>[]): string {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const filePath = path.join(OUTPUT_DIR, `Enquiry_List_${dateLabel(REPORT_FROM)}_${dateLabel(REPORT_TO)}.xlsx`);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'Enquiry List');
  try {
    XLSX.writeFile(workbook, filePath);
    return filePath;
  } catch (err) {
    // Usually the file is open in Excel (Windows locks it). Save beside it instead of losing the run.
    const fallback = filePath.replace(/\.xlsx$/, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${String((err as Error).message).slice(0, 80)}) — is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    XLSX.writeFile(workbook, fallback);
    return fallback;
  }
}

test.describe('MFR Vehicles Reports - Enquiry List', () => {
  test.use({ storageState: { cookies: [], origins: [] } });

  test('Should view Enquiry List report', async ({ page }) => {
    const loginPage = new LoginPage(page);
    const enquiryListReportPage = new EnquiryListReportPage(page);

    await loginPage.goto();
    await loginPage.login(credentials.dealervalidUser);
    await loginPage.verifyDashboard();

    await enquiryListReportPage.navigateToEnquiryList();
    await enquiryListReportPage.setFromDate(...pickerArgs(REPORT_FROM));
    await enquiryListReportPage.setToDate(...pickerArgs(REPORT_TO));

    const enquiryReportPage = await enquiryListReportPage.viewReport();
    await enquiryListReportPage.openColumnSelector(enquiryReportPage);

    const enquiryRows = await enquiryListReportPage.getReportRows(enquiryReportPage);
    console.log(`[STEP] Enquiry List record count (excluding header): ${enquiryRows.length}`);
    await enquiryListReportPage.showRecordCountPopup(enquiryReportPage, enquiryRows.length);

    if (enquiryRows.length === 0) {
      test.info().annotations.push({ type: 'warning', description: 'No Enquiry Found for the selected date range.' });
    }
    expect(enquiryRows.length).toBeGreaterThan(0);

    await enquiryListReportPage.closeReport(enquiryReportPage, 'Enquiry List');

    const savedPath = saveEnquiryList(enquiryRows);
    console.log(`[STEP] Enquiry List (${enquiryRows.length} rows) saved to ${savedPath}`);
  });
});
