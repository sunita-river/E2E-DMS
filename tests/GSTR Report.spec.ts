import { test, expect, Page } from '@playwright/test';
import { PDFParse } from 'pdf-parse';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage } from './models/LoginPage';
import { ReportViewerPage } from './models/ReportViewerPage';
import { missingColumns, blankValues, datesOutsideRange, nonNumeric, gstrTotalsMismatch } from './utils/reportChecks';
// Optional overrides (handy on CI/CD, where secrets and folders live elsewhere):
//   CREDENTIALS_FILE   path to a credentials.json (default: resources/credentials.json)
//   GSTR_OUTPUT_DIR    where the Excel files go (default: <project>/Output)
//   GSTR_FROM/GSTR_TO  report dates as dd-mm-yyyy (default: 01-09-2026 to yesterday)
//   REPORT_DOWNLOAD_DIR where 1000+ row report Excels are saved and read from (default: <project>/Output/Downloads)
const CREDENTIALS_FILE = process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
const OUTPUT_DIR = process.env.GSTR_OUTPUT_DIR || path.join(process.cwd(), 'Output');

function parseDateEnv(name: string, fallback: Date): Date {
  const v = process.env[name];
  if (!v) return fallback;
  const m = v.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!m) throw new Error(`${name} must be dd-mm-yyyy (got "${v}")`);
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

// Report range: 1st September 2026 to yesterday (unless GSTR_FROM / GSTR_TO are set).
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const yesterday = (() => {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() - 1);
})();
const REPORT_FROM = parseDateEnv('GSTR_FROM', new Date(2026, 8, 1));  // months are 0-based: 8 = September
const REPORT_TO = parseDateEnv('GSTR_TO', yesterday);
if (REPORT_TO < REPORT_FROM) {
  throw new Error(`GSTR To date ${REPORT_TO.toDateString()} is before From date ${REPORT_FROM.toDateString()} - set GSTR_FROM/GSTR_TO.`);
}

const GSTR_REQUIRED_COLUMNS = ['InvoiceNo', 'InvoiceDate', 'ReferenceNo', 'ReportOn', 'TaxableAmount', 'TotalAmount', 'DiscountAmount', 'Chassis No'];

const dateLabel = (d: Date) =>
  `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
const REPORT_FROM_DATE_LABEL = dateLabel(REPORT_FROM);
function todayDateLabel(): string {
  return dateLabel(REPORT_TO);
}

// DiscountAmount is either the report's own value ("0.00") or one read from the invoice PDF
// ("7500"). Write both as plain numbers ("0.00" -> "0").
function normalizeDiscountAmount(value: string): string {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? String(parsed) : value;
}

// <OUTPUT_DIR>/<dealercode>_SalesGSTNReport_<fromDate>_<toDate>.xlsx
function outputFilePathFor(dealerCode: string): string {
  return path.join(OUTPUT_DIR, `${dealerCode}_SalesGSTNReport_${REPORT_FROM_DATE_LABEL}_${todayDateLabel()}.xlsx`);
}

// A dealer whose file for this exact date range is already in Output is not re-run (so a
// rerun after a failure picks up where it stopped); its saved rows still go into MasterData.
// Set GSTR_RERUN_ALL=1 to force every dealer to run again.
function readSavedDealerRows(dealerCode: string): (Record<string, string> & { ReferenceNo: string })[] | null {
  const filePath = outputFilePathFor(dealerCode);
  if (process.env.GSTR_RERUN_ALL || !fs.existsSync(filePath)) return null;
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
  return XLSX.utils.sheet_to_json<Record<string, string>>(sheet, { raw: false, defval: '' })
    .map((row) => ({ ...row, ReferenceNo: row.ReferenceNo ?? '' }));
}

// Post GST Discounts already read from invoice PDFs, keyed "<dealercode>|<ReferenceNo>". An
// invoiced booking's discount doesn't change, so a daily run only opens new invoices.
// GSTR_NO_CACHE=1 re-reads every invoice (e.g. after one was cancelled and re-issued).
const DISCOUNT_CACHE_FILE = path.join(OUTPUT_DIR, 'gstr-discount-cache.json');

function loadDiscountCache(): Record<string, string> {
  try {
    return JSON.parse(fs.readFileSync(DISCOUNT_CACHE_FILE, 'utf8'));
  } catch {
    return {}; // first run, or unreadable: look everything up
  }
}

function saveDiscountCache(cache: Record<string, string>) {
  try {
    fs.mkdirSync(path.dirname(DISCOUNT_CACHE_FILE), { recursive: true });
    fs.writeFileSync(DISCOUNT_CACHE_FILE, JSON.stringify(cache, null, 1));
  } catch (err) {
    console.log(`[WARN] Could not save the discount cache (${String((err as Error).message).slice(0, 80)}) - next run will re-read these invoices.`);
  }
}

const discountCache = loadDiscountCache();

function saveRowsToExcel(dealerCode: string, rows: Record<string, string>[]): string {
  const filePath = outputFilePathFor(dealerCode);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });

  // "No." is always written as a fresh 1..N serial in the first column: the on-screen grid has
  // one, but the Excel export used for 1000+ row reports doesn't, and in MasterData each
  // dealer's own numbering would restart at 1 anyway.
  const normalizedRows = rows.map((row, i) => {
    const { ['No.']: _oldNo, ...rest } = row;
    const fixed = rest.DiscountAmount === undefined ? rest : { ...rest, DiscountAmount: normalizeDiscountAmount(rest.DiscountAmount) };
    return { 'No.': String(i + 1), ...fixed };
  });

  const worksheet = XLSX.utils.json_to_sheet(normalizedRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sales GSTN Report');
  try {
    XLSX.writeFile(workbook, filePath);
  } catch (err) {
    // Usually the file is open in Excel (Windows locks it). Save beside it instead of losing the run.
    const fallback = filePath.replace(/\.xlsx$/, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${String((err as Error).message).slice(0, 80)}) - is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    XLSX.writeFile(workbook, fallback);
    return fallback;
  }

  return filePath;
}

// Everything the run expected in DMS but did not find (a report gone from the menu, missing
// columns, bookings without an invoice, discounts that could not be read, dealers not finished).
// Written to its own red-highlighted Excel and attached as GSTR-missing.json, which
// tests/reporters/gmail-reporter.ts shows as a highlighted "Not found in DMS" section in the email.
type MissingItem = { Dealer: string; Missing: string; Where: string; Details: string };
const missingItems: MissingItem[] = [];

function missingFilePath(): string {
  return path.join(OUTPUT_DIR, `GSTR_Missing_${REPORT_FROM_DATE_LABEL}_${todayDateLabel()}.xlsx`);
}

async function saveMissingToExcel(items: MissingItem[]): Promise<string> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Not found in DMS', { views: [{ state: 'frozen', ySplit: 1 }] });
  sheet.columns = [
    { header: 'No.', key: 'no', width: 6 },
    { header: 'Dealer', key: 'Dealer', width: 12 },
    { header: 'Missing', key: 'Missing', width: 34 },
    { header: 'Where', key: 'Where', width: 44 },
    { header: 'Details', key: 'Details', width: 90 },
  ];
  const fill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  sheet.getRow(1).eachCell((c) => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = fill('FFB3261E');
  });
  if (!items.length) sheet.addRow({ no: '', Dealer: '', Missing: 'Nothing missing', Where: '', Details: 'Everything the run needed was found in DMS.' });
  items.forEach((item, i) => {
    const row = sheet.addRow({ no: i + 1, ...item });
    row.eachCell({ includeEmpty: true }, (c) => {
      c.fill = fill('FFFCE8E6');
      c.font = { color: { argb: 'FFB3261E' } };
      c.alignment = { vertical: 'top', wrapText: true };
    });
  });
  const filePath = missingFilePath();
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  try {
    await workbook.xlsx.writeFile(filePath);
    return filePath;
  } catch (err) {
    const fallback = filePath.replace(/\.xlsx$/, `_${Date.now()}.xlsx`);
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${oneLine(err)}) - is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    await workbook.xlsx.writeFile(fallback);
    return fallback;
  }
}

function printTable(headers: string[], rows: (string | number)[][]) {
  const widths = headers.map((h, i) => Math.max(h.length, ...rows.map((r) => String(r[i]).length)));
  const border = (l: string, m: string, r: string) => l + widths.map((w) => '─'.repeat(w + 2)).join(m) + r;
  const formatRow = (cells: (string | number)[]) =>
    '│ ' + cells.map((c, i) => String(c).padEnd(widths[i]!)).join(' │ ') + ' │';

  console.log(border('┌', '┬', '┐'));
  console.log(formatRow(headers));
  console.log(border('├', '┼', '┤'));
  rows.forEach((r) => console.log(formatRow(r)));
  console.log(border('└', '┴', '┘'));
}

// Reports don't use consistent header names for the same field, so look a field
// up by trying each known name, ignoring whitespace/case differences.
function getField(row: Record<string, string>, ...names: string[]): string {
  const normalize = (s: string) => s.replace(/\s+/g, '').toLowerCase();
  const normalizedRow = new Map(Object.entries(row).map(([k, v]) => [normalize(k), v]));
  for (const name of names) {
    const value = normalizedRow.get(normalize(name));
    if (value !== undefined && value.trim() !== '') return value;
  }
  return '';
}

// Print Invoice opens a popup showing the PDF in Chrome's viewer (ReportMs.aspx?outputtype=pdf),
// which the DOM can't read, so the same URL is fetched again with the session's cookies.
// Returns null when no popup opens: a booking not yet at "Invoice & Insurance" still shows the
// dropdown but Print Invoice does nothing (seen on 00184074). That means not invoiced, not an error.
async function getPrintInvoicePdfBuffer(page: Page): Promise<Buffer | null> {
  const popupPromise = page.waitForEvent('popup', { timeout: 8000 }).catch(() => null);
  await page.locator('#ctl00_cpMain_cmdPrintInvoice').click({ timeout: 15000 });
  const popup = await popupPromise;
  if (!popup) return null;

  // Take the URL as soon as the popup navigates; letting it load meant DMS generated every
  // invoice twice. Wait for the load only if it lands somewhere other than the report page.
  await popup.waitForURL((u) => u.href !== 'about:blank', { waitUntil: 'commit', timeout: 30000 }).catch(() => {});
  if (!/ReportMs\.aspx/i.test(popup.url())) await popup.waitForLoadState('load', { timeout: 30000 }).catch(() => {});
  const pdfUrl = popup.url();
  await popup.close().catch(() => {});
  if (pdfUrl === 'about:blank') throw new Error('Print Invoice popup never got an address');

  const response = await page.context().request.get(pdfUrl, { timeout: 60000 });
  if (!response.ok()) throw new Error(`Invoice PDF request returned HTTP ${response.status()}`);
  return response.body();
}

// For steps with no timeout of their own (e.g. PDF parsing), so nothing can hang a CI run.
function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: NodeJS.Timeout;
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => { timer = setTimeout(() => reject(new Error(`${what} took longer than ${ms / 1000}s`)), ms); }),
  ]).finally(() => clearTimeout(timer));
}

const oneLine = (err: unknown) => String(err instanceof Error ? err.message : err).replace(/\s+/g, ' ').slice(0, 160);

// The report itself is gone from the menu: retrying, or moving on to the next dealer, can't help.
class ReportMissingError extends Error {}

type DealerCredentials = { dealercode: string; username: string; password: string };
type ReportRow = Record<string, string> & { ReferenceNo: string };

// Searches one booking on the Product Sale Work Sheet and reads the Post GST Discount from its
// invoice PDF. Returns the discount, or null when the booking has no invoice yet. Throws on
// anything unexpected so the caller can retry once and then warn.
async function lookUpPostGstDiscount(page: Page, referenceNo: string): Promise<string | null> {
  await page.fill('#ctl00_cpMain_code', referenceNo, { timeout: 15000 });
  // Wait for the postback to reload the page, or the previous booking's Print Invoice button can
  // be read by mistake. Waits on navigation rather than a URL: the search moved from
  // ProductSaleWorkSheet.aspx to CustomerCenter.aspx once already. The shorter click timeout
  // gives a clearer error when a leftover dropdown covers the Show button.
  await Promise.all([
    page.waitForEvent('framenavigated', { predicate: (f) => f === page.mainFrame(), timeout: 60000 }),
    page.click('#ctl00_cpMain_cmdShow', { timeout: 15000 }),
  ]);
  await page.waitForLoadState('load', { timeout: 60000 });

  // Not rendered for bookings still at Booking/Payment/Pre-Invoice (no invoice yet). It can
  // appear a moment after load, hence the short wait.
  const printDropdownToggle = page.locator('#ctl00_cpMain_PrintInvoiceFormat');
  if (!(await printDropdownToggle.waitFor({ state: 'visible', timeout: 5000 }).then(() => true, () => false))) return null;

  await printDropdownToggle.scrollIntoViewIfNeeded({ timeout: 15000 });
  await printDropdownToggle.click({ timeout: 15000 });
  const pdfBuffer = await getPrintInvoicePdfBuffer(page);
  if (!pdfBuffer) return null;
  return (await withTimeout(findValueInInvoicePdf(pdfBuffer, 'Post GST Discount'), 60000, 'Reading the invoice PDF')) || '0';
}

// Closes invoice popups and the open Print Invoice dropdown so they can't cover the Show
// button for the next booking.
async function closeLeftovers(page: Page) {
  for (const other of page.context().pages()) {
    if (other !== page) await other.close().catch(() => {});
  }
  await page.keyboard.press('Escape').catch(() => {});
}

// Opens the Product Sale Work Sheet search fresh. DMS can log the user out partway through a
// long run (the search page then redirects to the login page), so sign in again when that happens.
async function openProductSaleWorkSheet(page: Page, dealer: DealerCredentials) {
  await closeLeftovers(page);
  const url = 'https://rivermobility.gaindms.com/Product/ProductSaleWorkSheet.aspx';
  await page.goto(url, { timeout: 60000 });
  if (await page.getByRole('button', { name: 'Sign In' }).isVisible().catch(() => false)) {
    console.log(`[STEP] Dealer ${dealer.dealercode}: DMS session expired - signing in again...`);
    const loginPage = new LoginPage(page);
    await loginPage.login(dealer);
    await loginPage.verifyDashboard();
    await page.goto(url, { timeout: 60000 });
  }
  await page.locator('table#ctl00_cpMain_grdMain > tbody > tr:nth-of-type(2) > td:nth-of-type(2) > a').click({ timeout: 30000 });
  await page.locator('#ctl00_cpMain_code').waitFor({ state: 'visible', timeout: 30000 });
}

// After this many bookings in a row fail even with a fresh page, the dealer is failed (and
// retried at the end of the run) instead of working through hundreds of doomed lookups.
const MAX_CONSECUTIVE_LOOKUP_FAILURES = 5;

// Leaves the browser ready for the next dealer whatever state this one ended in: closes any
// report/invoice tabs left open, logs out if possible, and clears the session cookies so the
// next login starts clean even when the logout link couldn't be clicked.
async function resetSession(page: Page, username: string) {
  for (const other of page.context().pages()) {
    if (other !== page) await other.close().catch(() => {});
  }
  try {
    await page.locator(`a:has-text("${username}")`).click({ timeout: 10000 });
    await page.getByRole('link', { name: /Sign out/i }).click({ timeout: 10000 });
  } catch {
    console.log('[STEP] Logout link not available - clearing the session instead.');
  }
  await page.context().clearCookies().catch(() => {});
}

// Returns the text after labelText on the same line ("Post GST Discount 0" -> "0"). The
// discount isn't in a bordered table on the invoice, so plain text is the reliable way in.
async function findValueInInvoicePdf(pdfBuffer: Buffer, labelText: string): Promise<string> {
  // An error page (HTML) instead of a PDF is the usual cause of "Invalid PDF structure".
  if (!pdfBuffer.subarray(0, 1024).toString('latin1').includes('%PDF')) {
    throw new Error(`Print Invoice did not return a PDF (got: ${pdfBuffer.subarray(0, 80).toString('utf8').replace(/\s+/g, ' ')})`);
  }
  const parser = new PDFParse({ data: pdfBuffer });
  try {
    const { text } = await parser.getText();
    const lines = text.split('\n').map((l: string) => l.trim()).filter(Boolean);
    for (const line of lines) {
      const idx = line.indexOf(labelText);
      if (idx === -1) continue;
      const after = line.slice(idx + labelText.length).trim();
      if (after) return after;
    }
    return '';
  } finally {
    await parser.destroy();
  }
}

// Runs one dealer end to end: login -> ProductWise Sales Details -> Post GST Discount lookups ->
// save that dealer's Excel. Returns the dealer's rows. Throws if the dealer can't be completed;
// the caller retries once and then records the failure without stopping the other dealers.
async function processDealer(page: Page, dealer: DealerCredentials): Promise<ReportRow[]> {
  const loginPage = new LoginPage(page);

  await loginPage.goto();
  await loginPage.login(dealer);
  await loginPage.verifyDashboard();

  console.log('[STEP] Navigating to MIS > Accounts Reports > GSTR Reports > ProductWise Sales Details...');
  await page.getByRole('link', { name: ' MIS ' }).click();
  await page.getByRole('link', { name: 'Accounts Reports' }).click();
  await page.locator('[href="#GSTRReports"]:visible').click();
  const productWiseSalesHeading = page.getByRole('heading', { name: 'ProductWise Sales Details', level: 5 });
  await productWiseSalesHeading.waitFor({ state: 'visible', timeout: 30000 }).catch(async () => {
    const shown = await page.locator('h5:visible').allInnerTexts().catch(() => [] as string[]);
    missingItems.push({
      Dealer: dealer.dealercode,
      Missing: 'Report: ProductWise Sales Details',
      Where: 'MIS > Accounts Reports > GSTR Reports',
      Details: `Not in the menu any more (first seen 08-Oct-2026). The page now lists: ${shown.join(', ') || 'nothing'}. ` +
        'Sales Tax Details is not a replacement (no ReferenceNo / Chassis No). Raised with the DMS vendor.',
    });
    throw new ReportMissingError(`"ProductWise Sales Details" was not found at menu path ` +
      `MIS > Accounts Reports > GSTR Reports (heading h5 "ProductWise Sales Details") for dealer ` +
      `${dealer.dealercode}. That GSTR Reports page lists: ${shown.join(', ') || 'nothing'}. The DMS removed it from ` +
      `the menu (first seen 08-Oct-2026) and the 08-Oct MIS check found no renamed copy in any other MIS menu ` +
      `(Sales Tax Details lacks ReferenceNo and Chassis No). Ask the DMS vendor where it went; see ` +
      `docs/GainDMS-vendor-ticket-2026-10-08.md.`);
  });
  await productWiseSalesHeading.click();

  const gstrReportViewer = new ReportViewerPage(page, {
    from: '#ctl00_cpMain_FromDat',
    to: '#ctl00_cpMain_ToDat',
  });
  // The calendar clicks are only a convenience: if they fail (e.g. the popup renders
  // differently on a server), the check below types the dates in directly.
  await gstrReportViewer.setFromDate(String(REPORT_FROM.getDate()), MONTH_NAMES[REPORT_FROM.getMonth()], REPORT_FROM.getFullYear())
    .catch((err) => console.log(`[WARN] From Date calendar click failed (${oneLine(err)}) - typing it instead.`));
  await gstrReportViewer.setToDate(String(REPORT_TO.getDate()), MONTH_NAMES[REPORT_TO.getMonth()], REPORT_TO.getFullYear())
    .catch((err) => console.log(`[WARN] To Date calendar click failed (${oneLine(err)}) - typing it instead.`));

  // In a headed window the From calendar can stay open over the To box, so the To click lands
  // on it and To ends up as the From date. Type the date in when a box shows the wrong one.
  for (const [input, date, label] of [
    [gstrReportViewer.fromDateInput, REPORT_FROM, 'From'],
    [gstrReportViewer.toDateInput, REPORT_TO, 'To'],
  ] as const) {
    const expected = dateLabel(date).replace(/-/g, '/'); // the box uses dd/mm/yyyy
    const actual = await input.inputValue();
    if (actual !== expected) {
      console.log(`[WARN] ${label} Date showed "${actual}" after the calendar click; setting it to ${expected}.`);
      await input.fill(expected);
      await input.press('Tab'); // blur so the calendar popup closes
    }
  }
  const shownFrom = await gstrReportViewer.fromDateInput.inputValue();
  const shownTo = await gstrReportViewer.toDateInput.inputValue();
  console.log(`[STEP] Report dates: ${shownFrom} to ${shownTo}`);
  if (shownFrom !== dateLabel(REPORT_FROM).replace(/-/g, '/') || shownTo !== dateLabel(REPORT_TO).replace(/-/g, '/')) {
    throw new Error(`Report dates could not be set (showing ${shownFrom} to ${shownTo})`);
  }

  const reportPage = await gstrReportViewer.viewReport();
  await gstrReportViewer.openColumnSelector(reportPage).catch(() => {}); // cosmetic only

  const reportRows = await gstrReportViewer.getReportRows(reportPage);
  console.log(`[STEP] ProductWise Sales Details record count (excluding header): ${reportRows.length}`);
  await gstrReportViewer.showRecordCountPopup(reportPage, reportRows.length).catch(() => {}); // cosmetic only
  await gstrReportViewer.closeReport(reportPage, 'ProductWise Sales Details').catch(() => {});

  // A dealer with no sales in the date range is saved as an empty file (so a rerun won't
  // redo it) and doesn't fail the run.
  if (reportRows.length === 0) {
    console.log(`[WARN] Dealer ${dealer.dealercode}: no records found for the selected date range.`);
    missingItems.push({
      Dealer: dealer.dealercode,
      Missing: 'Records',
      Where: 'ProductWise Sales Details',
      Details: `No records for ${REPORT_FROM_DATE_LABEL} to ${todayDateLabel()}.`,
    });
    test.info().annotations.push({ type: 'warning', description: `Dealer ${dealer.dealercode}: no records for the selected date range.` });
    saveRowsToExcel(dealer.dealercode, []);
    return [];
  }

  // productSaleRows holds the same objects, so discounts set in the loop below land in
  // allRowsWithDiscount. Other rows keep the report's own DiscountAmount.
  const allRowsWithDiscount: ReportRow[] = reportRows.map((row) => ({
    ...row,
    ReferenceNo: getField(row, 'ReferenceNo'),
  }));
  const productSaleRows = allRowsWithDiscount.filter((row) => getField(row, 'ReportOn') === 'Product Sale' && row.ReferenceNo);
  console.log(`[STEP] Reference numbers for ReportOn = Product Sale (${productSaleRows.length}):`, productSaleRows.map((r) => r.ReferenceNo));

  // Discounts read on an earlier run are filled in straight away; only the rest need an invoice.
  const cacheKey = (row: ReportRow) => `${dealer.dealercode}|${row.ReferenceNo}`;
  const rowsToLookUp = productSaleRows.filter((row) => {
    const cached = process.env.GSTR_NO_CACHE ? undefined : discountCache[cacheKey(row)];
    if (cached === undefined) return true;
    row.DiscountAmount = cached;
    return false;
  });
  console.log(`[STEP] Post GST Discount: ${productSaleRows.length - rowsToLookUp.length} reused from earlier runs, ${rowsToLookUp.length} to look up.`);

  if (rowsToLookUp.length) {
    console.log('[STEP] Navigating to Product Sale Work Sheet...');
    await openProductSaleWorkSheet(page, dealer);
  }

  const skipped: string[] = [];
  const notInvoiced: string[] = [];
  let consecutiveFailures = 0;
  for (const row of rowsToLookUp) {
    console.log(`[STEP] Looking up Post GST Discount for reference ${row.ReferenceNo}...`);
    let discount: string | null = null;
    let lastError: unknown = null;
    // Failures so far (error page instead of PDF, slow postback) cleared on a second try, but
    // only from a freshly opened search page; retrying on the stuck page fails again.
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        if (attempt === 2) await openProductSaleWorkSheet(page, dealer);
        discount = await lookUpPostGstDiscount(page, row.ReferenceNo);
        lastError = null;
        break;
      } catch (err) {
        lastError = err;
        if (attempt === 1) console.log(`[STEP] Reference ${row.ReferenceNo}: attempt 1 failed (${oneLine(err)}) - reopening the search page and retrying once...`);
      }
    }
    await closeLeftovers(page);
    if (lastError) {
      // Leave the report's own DiscountAmount for this row and carry on with the next one.
      console.log(`[WARN] Reference ${row.ReferenceNo}: discount could not be read after 2 attempts - left as in the report (${oneLine(lastError)}).`);
      skipped.push(row.ReferenceNo);
      if (++consecutiveFailures >= MAX_CONSECUTIVE_LOOKUP_FAILURES) {
        saveDiscountCache(discountCache); // keep what this dealer did read before giving up
        throw new Error(`Post GST Discount lookup failed for ${consecutiveFailures} bookings in a row (last: ${oneLine(lastError)}). ` +
          `DMS looks stuck, stopping this dealer so it is retried at the end.`);
      }
      // Don't carry a possibly stuck page into the next booking.
      await openProductSaleWorkSheet(page, dealer).catch((err) => console.log(`[WARN] Could not reopen the Product Sale Work Sheet (${oneLine(err)}).`));
      continue;
    }
    consecutiveFailures = 0;
    if (discount === null) {
      console.log(`[STEP] Reference ${row.ReferenceNo}: no Print Invoice option yet (not invoiced) - skipping.`);
      notInvoiced.push(row.ReferenceNo);
      continue;
    }
    row.DiscountAmount = discount;
    discountCache[cacheKey(row)] = discount;
    console.log(`[STEP] Reference ${row.ReferenceNo}: Post GST Discount = ${row.DiscountAmount}`);
  }
  // Saved once per dealer rather than after every booking.
  if (rowsToLookUp.length) saveDiscountCache(discountCache);
  if (skipped.length) {
    test.info().annotations.push({
      type: 'warning',
      description: `Dealer ${dealer.dealercode}: discount not read for ${skipped.length} reference(s): ${skipped.join(', ')}`,
    });
    missingItems.push({
      Dealer: dealer.dealercode,
      Missing: 'Post GST Discount (could not be read)',
      Where: 'Product Sale Work Sheet > Print Invoice',
      Details: `${skipped.length} booking(s), DiscountAmount left as in the report: ${skipped.join(', ')}`,
    });
  }
  if (notInvoiced.length) {
    missingItems.push({
      Dealer: dealer.dealercode,
      Missing: 'Invoice (not invoiced yet)',
      Where: 'Product Sale Work Sheet > Print Invoice',
      Details: `${notInvoiced.length} booking(s) have no invoice, so no Post GST Discount: ${notInvoiced.join(', ')}`,
    });
  }

  console.log('[STEP] Product Sale rows with Post GST Discount:');
  printTable(
    ['Serial No', 'Reference No', 'Discount Amount'],
    productSaleRows.map((row, i) => [i + 1, row.ReferenceNo, row.DiscountAmount ?? ''])
  );

  const outputFilePath = saveRowsToExcel(dealer.dealercode, allRowsWithDiscount);
  console.log(`[STEP] Saved full report (${allRowsWithDiscount.length} rows) to ${outputFilePath}`);
  return allRowsWithDiscount;
}

test.describe('Vehicle Sales - Sales > Vehicle Sales', () => {
  test.use({ storageState: { cookies: [], origins: [] } });

  // One test for all dealers: real Chrome (channel: 'chrome') exits when its last window
  // closes, so a test per dealer would kill the browser at the first teardown
  // (same issue as downloadExcelDirect() in ReportViewerPage.ts).
  test('Should login as each dealer and look up Product Sale discounts', async ({ page }) => {
    const dealers = (credentials.dealerUsers ?? []) as DealerCredentials[];
    if (!dealers.length) throw new Error('resources/credentials.json has no "dealerUsers" - nothing to run.');
    const incomplete = dealers.filter((d) => !d.dealercode || !d.username || !d.password);
    if (incomplete.length) {
      throw new Error(`credentials.json: dealercode/username/password missing for ${incomplete.map((d) => d.dealercode || '(no code)').join(', ')}`);
    }

    // Budget ~15 min per dealer (big dealers take ~6-8 min, plus one retry) so a large run is
    // never cut off just for being long, but a truly stuck run still ends.
    test.setTimeout(dealers.length * 15 * 60 * 1000 + 5 * 60 * 1000);

    const rowsByDealer = new Map<string, ReportRow[]>();
    const failed: { dealer: DealerCredentials; error: unknown }[] = [];

    const runDealer = async (dealer: DealerCredentials, attemptLabel: string) => {
      console.log(`\n[STEP] ===== Dealer ${dealer.dealercode} (${attemptLabel}) =====`);
      try {
        rowsByDealer.set(dealer.dealercode, await processDealer(page, dealer));
        return true;
      } catch (err) {
        if (err instanceof ReportMissingError) throw err;
        console.log(`[ERROR] Dealer ${dealer.dealercode} failed (${attemptLabel}): ${oneLine(err)}`);
        await test.info().attach(`${dealer.dealercode}-failure.png`, {
          body: await page.screenshot({ fullPage: true, timeout: 15000 }).catch(() => Buffer.from('')),
          contentType: 'image/png',
        }).catch(() => {});
        return false;
      } finally {
        await resetSession(page, dealer.username);
      }
    };

    // Runs every dealer and the data checks. The finally below saves and emails the
    // missing-items list whichever way this ends.
    const runAllDealers = async () => {
      for (const dealer of dealers) {
        const savedRows = readSavedDealerRows(dealer.dealercode);
        if (savedRows) {
          console.log(`[STEP] Dealer ${dealer.dealercode}: already saved for this date range (${savedRows.length} rows) - reusing ${outputFilePathFor(dealer.dealercode)}`);
          rowsByDealer.set(dealer.dealercode, savedRows);
          continue;
        }
        if (!(await runDealer(dealer, 'attempt 1'))) failed.push({ dealer, error: null });
      }

      // Retry failed dealers once at the end (DMS hiccups are usually gone by then).
      for (const f of [...failed]) {
        if (await runDealer(f.dealer, 'retry')) failed.splice(failed.indexOf(f), 1);
      }

      // MasterData is always written from whatever finished, in credentials.json order.
      const masterRows = dealers.flatMap((d) => rowsByDealer.get(d.dealercode) ?? []);
      const masterFilePath = saveRowsToExcel('MasterData', masterRows);
      console.log(`[STEP] Saved combined master report (${masterRows.length} rows across ${rowsByDealer.size}/${dealers.length} dealers) to ${masterFilePath}`);

      expect(masterRows.length, 'MasterData row count should equal the sum of the dealer files')
        .toBe([...rowsByDealer.values()].reduce((n, rows) => n + rows.length, 0));

      // Data checks per dealer, after every file is saved. Soft, so a data problem is reported
      // alongside (not instead of) any dealers that could not be completed.
      for (const dealer of dealers) {
        const rows = rowsByDealer.get(dealer.dealercode);
        if (!rows?.length) continue;
        const productSaleRows = rows.filter((r) => getField(r, 'ReportOn') === 'Product Sale');
        const absentColumns = GSTR_REQUIRED_COLUMNS.filter((c) => !(c in rows[0]!));
        if (absentColumns.length) {
          missingItems.push({
            Dealer: dealer.dealercode,
            Missing: `Column(s): ${absentColumns.join(', ')}`,
            Where: 'MIS > Accounts Reports > GSTR Reports > ProductWise Sales Details',
            Details: 'The report no longer has these columns, so they are blank in the output.',
          });
        }
        const problems = [
          ...missingColumns(rows, GSTR_REQUIRED_COLUMNS),
          // InvoiceNo is only required on Product Sale: RSA Service rows come without one (33 on 28-Sep).
          ...blankValues(rows, ['ReportOn']),
          ...datesOutsideRange(rows, 'InvoiceDate', REPORT_FROM, REPORT_TO),
          ...nonNumeric(rows, ['TaxableAmount', 'TotalAmount']),
          ...gstrTotalsMismatch(rows),
          ...blankValues(productSaleRows, ['InvoiceNo', 'ReferenceNo', 'Chassis No']).map((p) => `Product Sale: ${p}`),
          ...nonNumeric(productSaleRows, ['DiscountAmount']).map((p) => `Product Sale: ${p}`),
        ];
        expect.soft(problems, `Dealer ${dealer.dealercode} data problems:\n${problems.join('\n')}`).toEqual([]);
      }

      // Fail the run (so CI/CD shows it red) only after everything that could be saved is saved.
      if (failed.length) {
        throw new Error(`${failed.length} dealer(s) could not be completed after a retry: ${failed.map((f) => f.dealer.dealercode).join(', ')}. ` +
          `MasterData is missing their rows; rerun the test to fill them in (finished dealers are reused).`);
      }
    };

    try {
      await runAllDealers();
    } finally {
      // Dealers that never got a result: still failing after the retry, or never reached
      // because the run stopped (e.g. the report is gone from the menu).
      for (const dealer of dealers) {
        if (rowsByDealer.has(dealer.dealercode)) continue;
        if (missingItems.some((m) => m.Dealer === dealer.dealercode && m.Missing.startsWith('Report:'))) continue; // already listed
        const didFail = failed.some((f) => f.dealer.dealercode === dealer.dealercode);
        missingItems.push({
          Dealer: dealer.dealercode,
          Missing: didFail ? 'Dealer data (failed after a retry)' : 'Dealer data (not run)',
          Where: 'GSTR run',
          Details: didFail ? 'Could not be completed after a retry. Its rows are not in MasterData.'
            : 'The run stopped before this dealer. Its rows are not in MasterData.',
        });
      }
      const missingFile = await saveMissingToExcel(missingItems).catch((err) => {
        console.log(`[WARN] Could not save the missing-items Excel (${oneLine(err)}).`);
        return '';
      });
      if (missingItems.length) {
        console.log(`[WARN] ${missingItems.length} item(s) not found in DMS${missingFile ? ` - listed in ${missingFile}` : ''}:`);
        printTable(['Dealer', 'Missing', 'Where'], missingItems.map((m) => [m.Dealer, m.Missing, m.Where]));
      }
      await test.info().attach('GSTR-missing.json', {
        body: JSON.stringify({ title: 'GSTR Report - not found in DMS', file: missingFile, items: missingItems }),
        contentType: 'application/json',
      });
    }
  });
});
