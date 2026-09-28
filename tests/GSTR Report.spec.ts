import { test, expect, Page } from '@playwright/test';
import { PDFParse } from 'pdf-parse';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { LoginPage } from './models/LoginPage';
import { ReportViewerPage } from './models/ReportViewerPage';
// Optional overrides (handy on CI/CD, where secrets and folders live elsewhere):
//   CREDENTIALS_FILE   path to a credentials.json (default: resources/credentials.json)
//   GSTR_OUTPUT_DIR    where the Excel files go (default: <project>/Output)
//   GSTR_FROM/GSTR_TO  report dates as dd-mm-yyyy (default: 01-09-2026 to 24-09-2026)
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

// Report range: 1st September 2026 to 24th September 2026 (unless GSTR_FROM / GSTR_TO are set).
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const REPORT_FROM = parseDateEnv('GSTR_FROM', new Date(2026, 8, 1));  // months are 0-based: 8 = September
const REPORT_TO = parseDateEnv('GSTR_TO', new Date(2026, 8, 24));

const dateLabel = (d: Date) =>
  `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
const REPORT_FROM_DATE_LABEL = dateLabel(REPORT_FROM);
function todayDateLabel(): string {
  return dateLabel(REPORT_TO);
}

// DiscountAmount mixes two sources: the report's own raw value (e.g. "0.00") for rows
// left untouched, and our PDF lookup's plain "0"/"7500" for Product Sale rows — normalize
// both to the same bare-integer style (e.g. "0.00" -> "0") for consistency in the output.
function normalizeDiscountAmount(value: string): string {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? String(parsed) : value;
}

// Saves rows to <project root>/Output/<dealercode>_SalesGSTNReport_<fromDate>_<toDate>.xlsx,
// creating the Output folder if it doesn't exist yet. Per-dealer files are meant to be
// clubbed into one master Excel by hand after every dealer has been run.
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

// Post GST Discounts already read from invoice PDFs, keyed "<dealercode>|<ReferenceNo>". Once a
// booking is invoiced its discount doesn't change, so later runs reuse it instead of opening the
// invoice again — only new and not-yet-invoiced bookings are looked up, which is what keeps a
// daily run short. Set GSTR_NO_CACHE=1 to ignore it and re-read every invoice (e.g. after an
// invoice was cancelled and re-issued); fresh reads are still saved to it.
const DISCOUNT_CACHE_FILE = path.join(OUTPUT_DIR, 'gstr-discount-cache.json');

function loadDiscountCache(): Record<string, string> {
  try {
    return JSON.parse(fs.readFileSync(DISCOUNT_CACHE_FILE, 'utf8'));
  } catch {
    return {}; // no cache yet (first run) or unreadable — just look everything up
  }
}

function saveDiscountCache(cache: Record<string, string>) {
  try {
    fs.mkdirSync(path.dirname(DISCOUNT_CACHE_FILE), { recursive: true });
    fs.writeFileSync(DISCOUNT_CACHE_FILE, JSON.stringify(cache, null, 1));
  } catch (err) {
    console.log(`[WARN] Could not save the discount cache (${String((err as Error).message).slice(0, 80)}) — next run will re-read these invoices.`);
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
    console.log(`[WARN] Could not write ${path.basename(filePath)} (${String((err as Error).message).slice(0, 80)}) — is it open in Excel? Saved as ${path.basename(fallback)} instead.`);
    XLSX.writeFile(workbook, fallback);
    return fallback;
  }

  return filePath;
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

// "Print Invoice" opens a popup whose URL renders a PDF directly
// (ReportMs.aspx?outputtype=pdf) via the browser's native PDF viewer, so its
// text isn't reachable through the DOM. Re-fetch that same URL through the
// browser context's request API (shares the session's auth cookies) to get
// the raw PDF bytes instead.
//
// A booking that hasn't reached the "Invoice & Insurance" stage yet has no
// invoice, and clicking Print Invoice for it does nothing visible at all (no
// popup, no dialog) — even though the Print Invoice dropdown itself is shown
// (confirmed on 00184074: no popup on either attempt, even with 20s and a fresh
// page). So a null return here means "not invoiced yet", not an error, and
// callers should skip that reference rather than fail the run.
async function getPrintInvoicePdfBuffer(page: Page): Promise<Buffer | null> {
  const popupPromise = page.waitForEvent('popup', { timeout: 8000 }).catch(() => null);
  await page.locator('#ctl00_cpMain_cmdPrintInvoice').click({ timeout: 15000 });
  const popup = await popupPromise;
  if (!popup) return null;

  // Only the popup's address is needed — the PDF itself is fetched once below. So take the URL
  // as soon as the popup navigates to it instead of letting the popup load the whole PDF first
  // (that generated every invoice twice). If it lands somewhere other than the report page,
  // fall back to waiting for it to finish loading, as before.
  await popup.waitForURL((u) => u.href !== 'about:blank', { waitUntil: 'commit', timeout: 30000 }).catch(() => {});
  if (!/ReportMs\.aspx/i.test(popup.url())) await popup.waitForLoadState('load', { timeout: 30000 }).catch(() => {});
  const pdfUrl = popup.url();
  await popup.close().catch(() => {});
  if (pdfUrl === 'about:blank') throw new Error('Print Invoice popup never got an address');

  const response = await page.context().request.get(pdfUrl, { timeout: 60000 });
  if (!response.ok()) throw new Error(`Invoice PDF request returned HTTP ${response.status()}`);
  return response.body();
}

// Rejects if the promise hasn't settled in time — for steps that have no timeout of their own
// (e.g. PDF parsing), so nothing can hang a CI run indefinitely.
function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: NodeJS.Timeout;
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => { timer = setTimeout(() => reject(new Error(`${what} took longer than ${ms / 1000}s`)), ms); }),
  ]).finally(() => clearTimeout(timer));
}

const oneLine = (err: unknown) => String(err instanceof Error ? err.message : err).replace(/\s+/g, ' ').slice(0, 160);

type DealerCredentials = { dealercode: string; username: string; password: string };
type ReportRow = Record<string, string> & { ReferenceNo: string };

// Searches one booking on the Product Sale Work Sheet and reads the Post GST Discount from its
// invoice PDF. Returns the discount, or null when the booking has no invoice yet. Throws on
// anything unexpected so the caller can retry once and then warn.
async function lookUpPostGstDiscount(page: Page, referenceNo: string): Promise<string | null> {
  await page.fill('#ctl00_cpMain_code', referenceNo, { timeout: 15000 });
  // Wait for the search postback to reload the page before looking at it — otherwise the
  // previous booking's Print Invoice button can still be on screen and get read by mistake.
  // This waits for the page itself to change, not a response from a particular URL: the search
  // used to post to ProductSaleWorkSheet.aspx but now posts to CustomerCenter.aspx (then
  // redirects), and matching on the URL made every lookup time out when that changed.
  // The click has its own shorter limit so a Show button covered by a leftover dropdown/popup
  // fails with Playwright's "element intercepts pointer events" message, not a bare 60s timeout.
  await Promise.all([
    page.waitForEvent('framenavigated', { predicate: (f) => f === page.mainFrame(), timeout: 60000 }),
    page.click('#ctl00_cpMain_cmdShow', { timeout: 15000 }),
  ]);
  await page.waitForLoadState('load', { timeout: 60000 });

  // A booking that hasn't reached the "Invoice & Insurance" stage yet (still at
  // Booking/Payment/Pre-Invoice) has no invoice, so this control isn't rendered at all.
  // Give it a few seconds to appear rather than checking the instant the page loads.
  const printDropdownToggle = page.locator('#ctl00_cpMain_PrintInvoiceFormat');
  if (!(await printDropdownToggle.waitFor({ state: 'visible', timeout: 5000 }).then(() => true, () => false))) return null;

  await printDropdownToggle.scrollIntoViewIfNeeded({ timeout: 15000 });
  await printDropdownToggle.click({ timeout: 15000 });
  const pdfBuffer = await getPrintInvoicePdfBuffer(page);
  if (!pdfBuffer) return null;
  return (await withTimeout(findValueInInvoicePdf(pdfBuffer, 'Post GST Discount'), 60000, 'Reading the invoice PDF')) || '0';
}

// Closes whatever one booking's lookup can leave behind — invoice popups/tabs and the open
// Print Invoice dropdown — so it can't cover the Show button for the next booking.
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
    console.log(`[STEP] Dealer ${dealer.dealercode}: DMS session expired — signing in again...`);
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
    console.log('[STEP] Logout link not available — clearing the session instead.');
  }
  await page.context().clearCookies().catch(() => {});
}

// Scans the PDF's plain text for a line containing labelText and returns whatever
// follows it on that same line (e.g. "Post GST Discount 0" -> "0"). The invoice's
// discount line isn't drawn as a bordered table, so plain text scanning is what
// actually finds it — confirmed across every reference number tested.
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
  await productWiseSalesHeading.waitFor({ state: 'visible', timeout: 30000 });
  await productWiseSalesHeading.click();

  const gstrReportViewer = new ReportViewerPage(page, {
    from: '#ctl00_cpMain_FromDat',
    to: '#ctl00_cpMain_ToDat',
  });
  // The calendar clicks are only a convenience: if they fail (e.g. the popup renders
  // differently on a server), the check below types the dates in directly.
  await gstrReportViewer.setFromDate(String(REPORT_FROM.getDate()), MONTH_NAMES[REPORT_FROM.getMonth()], REPORT_FROM.getFullYear())
    .catch((err) => console.log(`[WARN] From Date calendar click failed (${oneLine(err)}) — typing it instead.`));
  await gstrReportViewer.setToDate(String(REPORT_TO.getDate()), MONTH_NAMES[REPORT_TO.getMonth()], REPORT_TO.getFullYear())
    .catch((err) => console.log(`[WARN] To Date calendar click failed (${oneLine(err)}) — typing it instead.`));

  // The From calendar popup stays open after picking a day and, in a headed window, can sit
  // over the To Date box — the To click then lands on the From calendar and To ends up as the
  // From date (01/09/2026). Check both boxes and type the right date in if a click went astray.
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
    test.info().annotations.push({ type: 'warning', description: `Dealer ${dealer.dealercode}: no records for the selected date range.` });
    saveRowsToExcel(dealer.dealercode, []);
    return [];
  }

  // Only ReferenceNo is added here — any Discount Amount column the report already
  // has for the other rows is left completely untouched. productSaleRows is a
  // *filter* (not a copy) of the same objects, so filling in DiscountAmount during
  // the lookup loop below updates only those rows within allRowsWithDiscount.
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
  let consecutiveFailures = 0;
  for (const row of rowsToLookUp) {
    console.log(`[STEP] Looking up Post GST Discount for reference ${row.ReferenceNo}...`);
    let discount: string | null = null;
    let lastError: unknown = null;
    // One retry: the failures seen so far (error page instead of PDF, slow postback) were
    // temporary and worked on the next attempt. The retry starts from a freshly opened search
    // page (signing in again if needed) — retrying on the same stuck page just fails again.
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        if (attempt === 2) await openProductSaleWorkSheet(page, dealer);
        discount = await lookUpPostGstDiscount(page, row.ReferenceNo);
        lastError = null;
        break;
      } catch (err) {
        lastError = err;
        if (attempt === 1) console.log(`[STEP] Reference ${row.ReferenceNo}: attempt 1 failed (${oneLine(err)}) — reopening the search page and retrying once...`);
      }
    }
    await closeLeftovers(page);
    if (lastError) {
      // Leave the report's own DiscountAmount for this row and carry on with the next one.
      console.log(`[WARN] Reference ${row.ReferenceNo}: discount could not be read after 2 attempts — left as in the report (${oneLine(lastError)}).`);
      skipped.push(row.ReferenceNo);
      if (++consecutiveFailures >= MAX_CONSECUTIVE_LOOKUP_FAILURES) {
        throw new Error(`Post GST Discount lookup failed for ${consecutiveFailures} bookings in a row (last: ${oneLine(lastError)}) — ` +
          `DMS looks stuck, stopping this dealer so it is retried at the end.`);
      }
      // Don't carry a possibly stuck page into the next booking.
      await openProductSaleWorkSheet(page, dealer).catch((err) => console.log(`[WARN] Could not reopen the Product Sale Work Sheet (${oneLine(err)}).`));
      continue;
    }
    consecutiveFailures = 0;
    if (discount === null) {
      console.log(`[STEP] Reference ${row.ReferenceNo}: no Print Invoice option yet (not invoiced) — skipping.`);
      continue;
    }
    row.DiscountAmount = discount;
    discountCache[cacheKey(row)] = discount;
    saveDiscountCache(discountCache); // after every read, so a run that dies midway still keeps them
    console.log(`[STEP] Reference ${row.ReferenceNo}: Post GST Discount = ${row.DiscountAmount}`);
  }
  if (skipped.length) {
    test.info().annotations.push({
      type: 'warning',
      description: `Dealer ${dealer.dealercode}: discount not read for ${skipped.length} reference(s): ${skipped.join(', ')}`,
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

  // All dealers run inside one test, reusing the same page/browser context, rather than
  // one test() per dealer. Real Chrome (channel: 'chrome') quits its whole process when
  // its last open window closes — with a separate test per dealer, each dealer's teardown
  // would close the only remaining window and kill the browser before the next dealer's
  // test could start (see the same note on downloadExcelDirect() in ReportViewerPage.ts).
  test('Should login as each dealer and look up Product Sale discounts', async ({ page }) => {
    const dealers = (credentials.dealerUsers ?? []) as DealerCredentials[];
    if (!dealers.length) throw new Error('resources/credentials.json has no "dealerUsers" — nothing to run.');
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

    for (const dealer of dealers) {
      const savedRows = readSavedDealerRows(dealer.dealercode);
      if (savedRows) {
        console.log(`[STEP] Dealer ${dealer.dealercode}: already saved for this date range (${savedRows.length} rows) — reusing ${outputFilePathFor(dealer.dealercode)}`);
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

    // Fail the run (so CI/CD shows it red) only after everything that could be saved is saved.
    if (failed.length) {
      throw new Error(`${failed.length} dealer(s) could not be completed after a retry: ${failed.map((f) => f.dealer.dealercode).join(', ')}. ` +
        `MasterData is missing their rows — rerun the test to fill them in (finished dealers are reused).`);
    }
  });
});
