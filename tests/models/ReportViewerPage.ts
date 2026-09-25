import { Page, Locator, Download } from '@playwright/test';
import * as XLSX from 'xlsx';

// Reports over the viewer's display limit (1000 rows) don't render #GrdMain at all —
// the viewer shows a "more than the 1000-record display limit" banner and downloads
// the full report as Excel instead. Rows are cached per report tab so the several
// grid reads a test makes (count, dealer codes, rows) share one download/parse.
const autoDownloads = new WeakMap<Page, Promise<Download | null>>();
const rowCache = new WeakMap<Page, Promise<Record<string, string>[]>>();

// Shared behavior for report screens reached via the "View Report" link:
// date range pickers, opening the report in its own tab, and reading/
// interacting with the resulting grid. Enquiry List and Product-wise Sales
// Details report pages both extend this.
export class ReportViewerPage {
  readonly page: Page;
  readonly fromDateInput: Locator;
  readonly toDateInput: Locator;
  readonly viewReportLink: Locator;
  readonly excelDownloadLink: Locator;

  // Most report pages use #ctl00_cpMain_FromDate / ctl00$cpMain$ToDate for their date
  // pickers, but some (e.g. the Accounts > GSTR Reports > ProductWise Sales Details
  // page) use ...FromDat / ...ToDat instead — dateInputSelectors lets callers override
  // the defaults for those without duplicating the calendar navigation logic below.
  constructor(page: Page, dateInputSelectors?: { from: string; to: string }) {
    this.page = page;
    this.fromDateInput = page.locator(dateInputSelectors?.from ?? '#ctl00_cpMain_FromDate');
    this.toDateInput = page.locator(dateInputSelectors?.to ?? 'input[name="ctl00$cpMain$ToDate"]');
    this.viewReportLink = page.getByRole('link', { name: ' View Report' });
    this.excelDownloadLink = page.getByRole('link', { name: ' Excel Download' });
  }

  // Both the From and To date pickers keep their calendar popup markup in the DOM at
  // once (only one is actually visible at a time), so an unscoped day-text locator
  // matches both and trips Playwright's strict mode. Scope to the visible calendar.
  //
  // The visible calendar itself also shows leading/trailing days from adjacent
  // months (grayed out, but still :visible) to fill out the grid, so e.g. "26"
  // can match both a real August 26 cell and a leftover July 26 cell in the same
  // popup. Each cell's title attribute carries the full date ("Wednesday, August
  // 26, 2026"), so require the target month name there too to disambiguate.
  private calendarDay(day: string, monthName: string): Locator {
    return this.page
      .locator(`.ajax__calendar_day:visible[title*="${monthName}"]`)
      .filter({ hasText: new RegExp(`^${day}$`) });
  }

  // The calendar always opens on the current month, so a fixed target date earlier in
  // the year (e.g. 10/08/2026, opened in September or later) needs stepping back via
  // the prev-month arrow until the header title (e.g. "August, 2026") matches. Capped
  // at 24 steps so a bad month/year combination fails loudly instead of looping.
  private async navigateCalendarToMonth(monthName: string, year: number) {
    const title = this.page.locator('.ajax__calendar_title:visible');
    const prevArrow = this.page.locator('.ajax__calendar_prev:visible');
    for (let attempt = 0; attempt < 24; attempt++) {
      const titleText = await title.innerText();
      if (titleText.includes(monthName) && titleText.includes(String(year))) return;
      await prevArrow.click();
    }
    throw new Error(`Calendar did not reach ${monthName} ${year} after 24 prev-month clicks.`);
  }

  // day/month/year pin the picker to one exact date (e.g. setFromDate('10', 'August', 2026)
  // for a fixed 10/08/2026 start date). Omitting month/year keeps the old behavior of
  // picking that day in whichever month the calendar already has open (i.e. "today").
  async setFromDate(day: string, monthName?: string, year?: number) {
    await this.fromDateInput.click();
    const targetMonth = monthName ?? new Date().toLocaleString('en-US', { month: 'long' });
    if (monthName) await this.navigateCalendarToMonth(monthName, year ?? new Date().getFullYear());
    await this.calendarDay(day, targetMonth).click();
  }

  async setToDate(day: string, monthName?: string, year?: number) {
    await this.toDateInput.click();
    const targetMonth = monthName ?? new Date().toLocaleString('en-US', { month: 'long' });
    if (monthName) await this.navigateCalendarToMonth(monthName, year ?? new Date().getFullYear());
    await this.calendarDay(day, targetMonth).click();
  }

  /** Sets both From and To date pickers to the same day (e.g. yesterday). */
  async setDateRange(date: Date) {
    const day = String(date.getDate());
    await this.setFromDate(day);
    await this.setToDate(day);
  }

  // Clicking "Excel Download" on the report list page (before opening the interactive
  // viewer) fires both a file download and an incidental blank popup window. This runs
  // against real Google Chrome (channel: 'chrome'), where closing the last open window
  // quits the whole browser process — so the popup is deliberately left open rather than
  // closed, to avoid killing the main page (and the in-flight download) if it happens to
  // be the only other window at that moment.
  async downloadExcelDirect(): Promise<Download> {
    console.log('[STEP] Downloading report directly via Excel Download...');
    const downloadPromise = this.page.waitForEvent('download');
    await this.excelDownloadLink.click();
    const download = await downloadPromise;

    // "Excel Download" is an ASP.NET postback on the main page (not a plain anchor), so
    // the page is still re-rendering right after it fires. Clicking "View Report" too
    // soon after races that re-render: the click lands but the postback never completes,
    // so no popup ever opens and the caller hangs waiting for one indefinitely. Give the
    // postback time to settle first, mirroring the same wait LoginPage uses after Sign In.
    await this.page.waitForTimeout(2000);

    return download;
  }

  // Exports the grid from within an opened report viewer tab (as returned by viewReport()).
  async exportToExcel(reportPage: Page): Promise<Download> {
    console.log('[STEP] Exporting report to Excel from the report viewer...');
    const downloadPromise = reportPage.waitForEvent('download');
    await reportPage.getByRole('button', { name: 'Export to Excel' }).click();
    return downloadPromise;
  }

  async viewReport(): Promise<Page> {
    console.log('[STEP] Opening the report in a new tab...');
    // Big reports can take a while to build server-side before the viewer tab opens.
    const popupPromise = this.page.waitForEvent('popup', { timeout: 180000 });
    await this.viewReportLink.click();
    const reportPage = await popupPromise;
    // Listen straight away: an over-limit report starts its Excel download on load.
    autoDownloads.set(reportPage, reportPage.waitForEvent('download', { timeout: 60000 }).catch(() => null));
    await reportPage.bringToFront();

    // The main window runs maximized (--start-maximized, needed for the app itself to render),
    // and new windows opened from it inherit that maximized state. The report popup doesn't
    // need to be maximized, so put it back to a normal, non-maximized size. Cosmetic only —
    // skipped where there is no real window (headless CI) or no CDP (non-Chromium).
    try {
      const session = await reportPage.context().newCDPSession(reportPage);
      const { windowId } = await session.send('Browser.getWindowForTarget');
      await session.send('Browser.setWindowBounds', {
        windowId,
        bounds: { windowState: 'normal', width: 1280, height: 800 },
      });
    } catch {
      // no window to resize
    }

    return reportPage;
  }

  async openColumnSelector(reportPage: Page) {
    await reportPage.locator('section').filter({ hasText: 'Select the columns to be' }).click();
  }

  async getRecordCount(reportPage: Page): Promise<number> {
    return (await this.getReportRows(reportPage)).length;
  }

  // Playwright auto-dismisses unhandled dialogs, so the accept handler must be
  // registered before triggering the alert or the popup never actually shows.
  async showRecordCountPopup(reportPage: Page, count: number) {
    const message = count === 0
      ? 'Warning: No Enquiry Found'
      : `Enquiry List record count: ${count}`;
    reportPage.once('dialog', (dialog) => dialog.accept());
    await reportPage.evaluate((msg) => window.alert(msg), message);
  }

  // Closing the report tab (e.g. after the 2nd Product-wise Sales Details
  // view, once the date type has been switched and the report refreshed) can
  // trigger a native browser dialog from the page's own unload handling. With
  // no listener registered, Playwright auto-dismisses it, which leaves the
  // close() call blocked waiting on a dialog that never gets accepted.
  // Registering an accept handler first — and logging the dialog's message —
  // lets the tab close cleanly and gives visibility into what triggered it.
  async closeReport(reportPage: Page, label: string) {
    reportPage.on('dialog', async (dialog) => {
      console.log(`[DEBUG] Dialog appeared while closing ${label} report viewer: type=${dialog.type()}, message="${dialog.message()}"`);
      await dialog.accept();
    });
    console.log(`[STEP] Closing the ${label} report viewer...`);
    await reportPage.close();
  }

  // Looks up the column by its header text rather than a fixed index, since the
  // selected report columns (and therefore their order) can change per run.
  async getDealerCodes(reportPage: Page): Promise<string[]> {
    const rows = await this.getReportRows(reportPage);
    if (rows.length === 0) return [];
    const dealerCodeKey = Object.keys(rows[0]!).find((h) => h.replace(/\s+/g, '') === 'DealerCode');
    if (!dealerCodeKey) {
      throw new Error('Dealer Code column not found in the report.');
    }
    return rows.map((r) => (r[dealerCodeKey] ?? '').trim());
  }

  // Keyed by header text (rather than a fixed column layout) so callers can pull
  // whichever fields they need regardless of the report's configured column order.
  // Comes from the on-screen grid, or from the Excel export when the report is over
  // the viewer's display limit.
  async getReportRows(reportPage: Page): Promise<Record<string, string>[]> {
    let rows = rowCache.get(reportPage);
    if (!rows) {
      rows = this.loadReportRows(reportPage);
      rowCache.set(reportPage, rows);
    }
    return rows;
  }

  private async loadReportRows(reportPage: Page): Promise<Record<string, string>[]> {
    const overLimitBanner = reportPage.getByText(/record display limit/i);
    if (await overLimitBanner.isVisible()) {
      console.log(`[STEP] ${(await overLimitBanner.innerText()).trim()}`);
      return this.readOverLimitExcel(reportPage);
    }
    return this.readGridRows(reportPage);
  }

  // Uses the automatic download if it fired; otherwise clicks the banner's "click here".
  private async readOverLimitExcel(reportPage: Page): Promise<Record<string, string>[]> {
    let download = await Promise.race([
      autoDownloads.get(reportPage) ?? Promise.resolve(null),
      reportPage.waitForTimeout(15000).then(() => null),
    ]);
    if (!download) {
      console.log('[STEP] Automatic download did not start; clicking "click here"...');
      const downloadPromise = reportPage.waitForEvent('download', { timeout: 120000 });
      await reportPage.getByRole('link', { name: 'click here' }).click();
      download = await downloadPromise;
    }
    const filePath = await download.path();
    if (!filePath) {
      throw new Error('Over-limit report Excel download did not produce a file (the download may have failed).');
    }
    console.log(`[STEP] Reading over-limit report from Excel: ${download.suggestedFilename()}`);

    // cellDates + dateNF keep dates in the grid's "DD-MM-YYYY hh:mm" shape, which the
    // specs' date formatting expects.
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
    const table = XLSX.utils.sheet_to_json<string[]>(sheet, {
      header: 1,
      raw: false,
      defval: '',
      dateNF: 'dd-mm-yyyy hh:mm',
    });
    // The export may carry title rows above the column headers; the header row is the
    // one with a Dealer Code column. Single-dealer reports (e.g. GSTR ProductWise Sales
    // Details) have no Dealer Code column, so fall back to the first full-width row —
    // title rows above the headers only fill one or two cells.
    const filled = (r: string[]) => r.filter((c) => String(c).trim() !== '').length;
    let headerIndex = table.findIndex((r) => r.some((c) => String(c).replace(/\s+/g, '').toLowerCase() === 'dealercode'));
    if (headerIndex === -1) {
      const top = table.slice(0, 20);
      const widest = Math.max(0, ...top.map(filled));
      headerIndex = widest >= 3 ? top.findIndex((r) => filled(r) === widest) : -1;
    }
    if (headerIndex === -1) {
      throw new Error(`Header row not found in ${download.suggestedFilename()}.`);
    }
    const headers = table[headerIndex]!.map((h) => String(h).trim());
    return table
      .slice(headerIndex + 1)
      .filter((r) => r.some((c) => String(c).trim() !== ''))
      .map((r) => {
        const record: Record<string, string> = {};
        headers.forEach((header, idx) => {
          if (header) record[header] = String(r[idx] ?? '').trim();
        });
        return record;
      });
  }

  // Reads the whole grid in one page.evaluate() round trip instead of issuing a
  // separate Playwright call per row/cell — with ~1000 rows, per-row locator calls
  // are slow enough to blow the test timeout.
  private async readGridRows(reportPage: Page): Promise<Record<string, string>[]> {
    return reportPage.evaluate(() => {
      const headers = Array.from(document.querySelectorAll('#GrdMain tr:first-child th')).map((th) =>
        (th as HTMLElement).innerText.trim()
      );
      const rows = Array.from(document.querySelectorAll('#GrdMain tbody tr')).filter((tr) => tr.querySelector('td'));
      return rows.map((tr) => {
        const cells = Array.from(tr.querySelectorAll('td')).map((td) => (td as HTMLElement).innerText.trim());
        const record: Record<string, string> = {};
        headers.forEach((header, idx) => {
          record[header] = cells[idx] ?? '';
        });
        return record;
      });
    });
  }
}
