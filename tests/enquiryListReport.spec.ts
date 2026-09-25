import { test, expect } from '@playwright/test';
import { LoginPage } from './models/LoginPage';
import { EnquiryListReportPage } from './models/EnquiryListReportPage';
import { ProductWiseSalesDetailsReportPage } from './models/ProductWiseSalesDetailsReportPage';
import * as fs from 'node:fs';
import * as path from 'node:path';

// CREDENTIALS_FILE: optional path to a credentials.json (e.g. a CI secrets file)
const credentials = JSON.parse(fs.readFileSync(
  process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json'), 'utf8'));

const todayDay = String(new Date().getDate());

// Formats a "DD-MM-YYYY ..." grid date as e.g. "10th Aug".
function formatDateOrdinal(raw: string): string {
  const datePart = raw.split(' ')[0] ?? raw;
  const [dayStr, monthStr] = datePart.split('-');
  const day = Number(dayStr);
  const month = Number(monthStr);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const suffix = day >= 11 && day <= 13 ? 'th' : (['th', 'st', 'nd', 'rd'][day % 10] ?? 'th');
  return `${day}${suffix} ${months[month - 1]}`;
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

// Reports don't use consistent header names for the same field (e.g. Enquiry
// List has "Invoice Date" / "Dealer Code" / "Dealer name", while Product-wise
// Sales Details has "Booking Date" / "DealerCode" / "DealerName"), so look a
// field up by trying each known name, ignoring whitespace/case differences.
function getField(row: Record<string, string>, ...names: string[]): string {
  const normalize = (s: string) => s.replace(/\s+/g, '').toLowerCase();
  const normalizedRow = new Map(Object.entries(row).map(([k, v]) => [normalize(k), v]));
  for (const name of names) {
    const value = normalizedRow.get(normalize(name));
    // A row can carry a known field as an empty string rather than omit it
    // entirely — e.g. a booking with no invoice yet still has an 'InvoiceDate'
    // column, just blank. Treat that as "not found" so later, more specific
    // names (like 'Booking Date') get a chance instead of the blank winning.
    if (value !== undefined && value.trim() !== '') return value;
  }
  return '';
}

// Last-resort fallback when none of the known date header names match: scan
// the row for any column whose header contains "date" and use the first one
// that actually has a value.
function findAnyDateField(row: Record<string, string>): string {
  for (const [header, value] of Object.entries(row)) {
    if (/date/i.test(header) && value.trim() !== '') {
      console.log(`[DEBUG] None of the known date fields had a value; falling back to column "${header}": "${value}"`);
      return value;
    }
  }
  return '';
}

function dateDealerKeyParts(row: Record<string, string>) {
  const date = getField(row, 'Invoice Date', 'Booking Date', 'Created Date') || findAnyDateField(row);
  return {
    date,
    dealerCode: getField(row, 'Dealer Code', 'DealerCode'),
    dealerName: getField(row, 'Dealer name', 'DealerName'),
  };
}

// Groups report rows by Date + Dealer Code + Dealer name. Counts rows by
// default, or — when countBy is given (e.g. 'Booking No') — counts distinct
// values of that field instead, since a booking can span multiple product
// rows in the Product-wise Sales Details grid and a plain row count would
// overcount bookings per dealer.
function summarizeByDateAndDealer(rows: Record<string, string>[], countBy?: string) {
  const summaryByKey = new Map<
    string,
    { date: string; dealerCode: string; dealerName: string; rowCount: number; distinctValues: Set<string> }
  >();
  for (const row of rows) {
    const { date, dealerCode, dealerName } = dateDealerKeyParts(row);
    const key = `${date}|${dealerCode}|${dealerName}`;
    let entry = summaryByKey.get(key);
    if (!entry) {
      entry = { date, dealerCode, dealerName, rowCount: 0, distinctValues: new Set() };
      summaryByKey.set(key, entry);
    }
    entry.rowCount++;
    if (countBy) {
      entry.distinctValues.add(getField(row, countBy));
    }
  }
  return [...summaryByKey.values()]
    .map((e) => ({
      date: e.date,
      dealerCode: e.dealerCode,
      dealerName: e.dealerName,
      count: countBy ? e.distinctValues.size : e.rowCount,
    }))
    .sort((a, b) => a.date.localeCompare(b.date) || a.dealerCode.localeCompare(b.dealerCode));
}

// Appends a "Total" row summing every count column (index 4 onward — the
// columns after Serial No / Date / Dealer Code / Dealer Name). Used by both
// the single-source and combined summary tables, whose count columns always
// start at that same index.
function appendTotalRow(tableRows: (string | number)[][]): (string | number)[][] {
  if (tableRows.length === 0) return tableRows;
  const columnCount = tableRows[0]!.length;
  const totalRow: (string | number)[] = Array.from({ length: columnCount }, (_, colIndex) => {
    if (colIndex === 0) return 'Total';
    if (colIndex <= 3) return '';
    return tableRows.reduce((sum, row) => sum + Number(row[colIndex]), 0);
  });
  return [...tableRows, totalRow];
}

function printDateDealerSummary(rows: Record<string, string>[], countLabel: string, countBy?: string) {
  const summary = summarizeByDateAndDealer(rows, countBy);
  console.log(`[STEP] ${countLabel} count by Date / Dealer Code / Dealer Name:`);
  const tableRows = summary.map((s, i) => [i + 1, formatDateOrdinal(s.date), s.dealerCode, s.dealerName, s.count]);
  printTable(['Serial No', 'Date', 'Dealer Code', 'Dealer Name', countLabel], appendTotalRow(tableRows));
}

// Merges several row sets (e.g. Enquiry, Invoice, Booking) into a single table
// keyed by Date + Dealer Code + Dealer Name, with one count column per source.
// Each source can set countBy (e.g. 'Booking No') to count distinct values of
// that field per group instead of raw rows.
function buildCombinedSummaryTable(sources: { label: string; rows: Record<string, string>[]; countBy?: string }[]) {
  const summaryByKey = new Map<
    string,
    { date: string; dealerCode: string; dealerName: string; counts: Record<string, number>; distinctValues: Record<string, Set<string>> }
  >();

  for (const { label, rows, countBy } of sources) {
    for (const row of rows) {
      const { date, dealerCode, dealerName } = dateDealerKeyParts(row);
      const key = `${date}|${dealerCode}|${dealerName}`;
      let entry = summaryByKey.get(key);
      if (!entry) {
        entry = { date, dealerCode, dealerName, counts: {}, distinctValues: {} };
        summaryByKey.set(key, entry);
      }
      if (countBy) {
        const set = entry.distinctValues[label] ?? (entry.distinctValues[label] = new Set());
        set.add(getField(row, countBy));
        entry.counts[label] = set.size;
      } else {
        entry.counts[label] = (entry.counts[label] ?? 0) + 1;
      }
    }
  }

  const summary = [...summaryByKey.values()].sort(
    (a, b) => a.date.localeCompare(b.date) || a.dealerCode.localeCompare(b.dealerCode)
  );

  const labels = sources.map((s) => s.label);
  const headers = ['Serial No', 'Date', 'Dealer Code', 'Dealer Name', ...labels];
  const tableRows = summary.map((s, i) => [
    i + 1,
    formatDateOrdinal(s.date),
    s.dealerCode,
    s.dealerName,
    ...labels.map((label) => s.counts[label] ?? 0),
  ]);
  return { headers, tableRows: appendTotalRow(tableRows), labels };
}

function printCombinedSummary(sources: { label: string; rows: Record<string, string>[]; countBy?: string }[]) {
  const { headers, tableRows, labels } = buildCombinedSummaryTable(sources);
  console.log(`[STEP] Combined ${labels.join(' / ')} count by Date / Dealer Code / Dealer Name:`);
  printTable(headers, tableRows);
}

function toCsv(headers: string[], rows: (string | number)[][]): string {
  const escape = (value: string | number) => {
    const s = String(value);
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.map(escape).join(','), ...rows.map((r) => r.map(escape).join(','))];
  return lines.join('\r\n');
}

test.describe('MFR Vehicles Reports - Enquiry List', () => {
  test.use({ storageState: { cookies: [], origins: [] } });

  test('Should view Enquiry List then Product-wise Sales Details reports', async ({ page }) => {
    const loginPage = new LoginPage(page);
    const enquiryListReportPage = new EnquiryListReportPage(page);
    const salesDetailsReportPage = new ProductWiseSalesDetailsReportPage(page);

    await loginPage.goto();
    await loginPage.login(credentials.dealervalidUser);
    await loginPage.verifyDashboard();

    // --- Enquiry List ---
    await enquiryListReportPage.navigateToEnquiryList();
    await enquiryListReportPage.setFromDate('7', 'September', 2026);
    await enquiryListReportPage.setToDate(todayDay);

    const enquiryReportPage = await enquiryListReportPage.viewReport();
    await enquiryListReportPage.openColumnSelector(enquiryReportPage);

    const enquiryRecordCount = await enquiryListReportPage.getRecordCount(enquiryReportPage);
    console.log(`[STEP] Enquiry List record count (excluding header): ${enquiryRecordCount}`);
    await enquiryListReportPage.showRecordCountPopup(enquiryReportPage, enquiryRecordCount);

    if (enquiryRecordCount === 0) {
      test.info().annotations.push({ type: 'warning', description: 'No Enquiry Found for the selected date range.' });
    }
    expect(enquiryRecordCount).toBeGreaterThan(0);

    const enquiryDealerCodes = await enquiryListReportPage.getDealerCodes(enquiryReportPage);
    console.log(`[STEP] Dealer codes: ${enquiryDealerCodes.length} rows, ${new Set(enquiryDealerCodes).size} distinct dealers`);

    const enquiryRows = await enquiryListReportPage.getReportRows(enquiryReportPage);
    printDateDealerSummary(enquiryRows, 'Enquiries');

    await enquiryListReportPage.closeReport(enquiryReportPage, 'Enquiry List');

    // --- Product-wise Sales Details (continues in the same session, no re-login) ---
    await salesDetailsReportPage.navigateToProductWiseSalesDetails();
    await salesDetailsReportPage.setFromDate('7', 'September', 2026);
    await salesDetailsReportPage.setToDate(todayDay);

    const salesReportPage = await salesDetailsReportPage.viewReport();
    await salesDetailsReportPage.openColumnSelector(salesReportPage);

    // Default date type (Invoice) — captured before switching to Booking below.
    const invoiceRecordCount = await salesDetailsReportPage.getRecordCount(salesReportPage);
    console.log(`[STEP] Product-wise Sales Details record count, Invoice date type (excluding header): ${invoiceRecordCount}`);
    const invoiceRows = await salesDetailsReportPage.getReportRows(salesReportPage);
    console.log('[DEBUG] Invoice row sample (first row):', invoiceRows[0]);
    const invoiceDealerCodes = await salesDetailsReportPage.getDealerCodes(salesReportPage);
    console.log(`[STEP] Dealer codes: ${invoiceDealerCodes.length} rows, ${new Set(invoiceDealerCodes).size} distinct dealers`);
    printDateDealerSummary(invoiceRows, 'Invoice', 'Booking No');

    // The Booking date type report can't be generated by switching date type
    // in place — the invoice report tab has to be closed and the main form
    // navigated back to Product-wise Sales Details from the home link before
    // the date type dropdown and "View Report" link are actionable again.
    await salesDetailsReportPage.closeReport(salesReportPage, 'Product-wise Sales Details (Invoice)');
    console.log('[DEBUG] Returning to home before re-navigating to Product-wise Sales Details for Booking...');
    // The old #ctl00_homelink is gone; the header logo now links to the dashboard.
    await salesDetailsReportPage.page.getByRole('banner').locator('a[href*="UserDashboard.aspx"]').first().click();
    await salesDetailsReportPage.page.waitForLoadState('load');

    await salesDetailsReportPage.navigateToProductWiseSalesDetails();
    await salesDetailsReportPage.setFromDate('7', 'September', 2026);
    await salesDetailsReportPage.setToDate(todayDay);
    await salesDetailsReportPage.setDateType('B');
    const bookingReportPage = await salesDetailsReportPage.viewReport();
    await salesDetailsReportPage.openColumnSelector(bookingReportPage);

    const salesRecordCount = await salesDetailsReportPage.getRecordCount(bookingReportPage);
    console.log(`[STEP] Product-wise Sales Details record count, Booking date type (excluding header): ${salesRecordCount}`);
    await salesDetailsReportPage.showRecordCountPopup(bookingReportPage, salesRecordCount);

    if (salesRecordCount === 0) {
      test.info().annotations.push({ type: 'warning', description: 'No sales records found for the selected date range.' });
    }
    expect(salesRecordCount).toBeGreaterThan(0);

    const salesDealerCodes = await salesDetailsReportPage.getDealerCodes(bookingReportPage);
    console.log(`[STEP] Dealer codes: ${salesDealerCodes.length} rows, ${new Set(salesDealerCodes).size} distinct dealers`);

    const salesRows = await salesDetailsReportPage.getReportRows(bookingReportPage);
    printDateDealerSummary(salesRows, 'Booking', 'Booking No');

    await salesDetailsReportPage.closeReport(bookingReportPage, 'Product-wise Sales Details (Booking)');

    // --- Combined Enquiry / Invoice / Booking summary ---
    const combinedSources = [
      { label: 'Enquiry', rows: enquiryRows },
      { label: 'Invoice', rows: invoiceRows, countBy: 'Booking No' },
      { label: 'Booking', rows: salesRows, countBy: 'Booking No' },
    ];
    printCombinedSummary(combinedSources);

    const { headers: combinedHeaders, tableRows: combinedRows } = buildCombinedSummaryTable(combinedSources);
    await test.info().attach('combined-summary.csv', {
      body: toCsv(combinedHeaders, combinedRows),
      contentType: 'text/csv',
    });
  });
});
