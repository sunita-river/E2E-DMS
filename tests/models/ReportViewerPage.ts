import { Page, Locator, Download } from '@playwright/test';
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

// Reports over the viewer's 1000-row display limit don't render #GrdMain; the viewer shows a
// "record display limit" banner and downloads the report as Excel instead. Rows are cached per
// report tab so repeated reads share one download.
const autoDownloads = new WeakMap<Page, Promise<Download | null>>();
const rowCache = new WeakMap<Page, Promise<Record<string, string>[]>>();
// When each report tab was opened; only Excel files newer than this count as its download.
const reportOpenedAt = new WeakMap<Page, number>();

// Over-limit Excel files are kept here (Playwright's own copy is deleted when the run ends),
// and the rows are read from the kept file. Override with REPORT_DOWNLOAD_DIR.
const DOWNLOAD_DIR = process.env.REPORT_DOWNLOAD_DIR || path.join(process.cwd(), 'Output', 'Downloads');

// Newest .xls/.xlsx in the download folders modified since `since`, or null. Used when the
// download event never reached Playwright but the browser still saved the file.
function newestExcelSince(since: number): string | null {
  const dirs = [DOWNLOAD_DIR, path.join(os.homedir(), 'Downloads')];
  const candidates = dirs.flatMap((dir) => {
    try {
      return fs.readdirSync(dir)
        .filter((f) => /\.xlsx?$/i.test(f) && !f.startsWith('~$'))
        .map((f) => path.join(dir, f))
        .map((file) => ({ file, mtime: fs.statSync(file).mtimeMs }));
    } catch {
      return []; // folder doesn't exist
    }
  });
  const recent = candidates.filter((c) => c.mtime >= since - 5000).sort((a, b) => b.mtime - a.mtime);
  return recent[0]?.file ?? null;
}

// Shared parts of the report screens opened with "View Report": date pickers, the report tab
// and reading its grid.
export class ReportViewerPage {
  readonly page: Page;
  readonly fromDateInput: Locator;
  readonly toDateInput: Locator;
  readonly viewReportLink: Locator;
  readonly excelDownloadLink: Locator;

  // Most report pages use #ctl00_cpMain_FromDate / ctl00$cpMain$ToDate; some (e.g. GSTR
  // ProductWise Sales Details) use ...FromDat / ...ToDat, so callers can pass their own.
  constructor(page: Page, dateInputSelectors?: { from: string; to: string }) {
    this.page = page;
    this.fromDateInput = page.locator(dateInputSelectors?.from ?? '#ctl00_cpMain_FromDate');
    this.toDateInput = page.locator(dateInputSelectors?.to ?? 'input[name="ctl00$cpMain$ToDate"]');
    this.viewReportLink = page.getByRole('link', { name: ' View Report' });
    this.excelDownloadLink = page.getByRole('link', { name: ' Excel Download' });
  }

  // Both calendars stay in the DOM, so scope to the visible one. It also shows days from the
  // adjacent months (still :visible), so match the month name in the cell's title
  // ("Wednesday, August 26, 2026") as well.
  private calendarDay(day: string, monthName: string): Locator {
    return this.page
      .locator(`.ajax__calendar_day:visible[title*="${monthName}"]`)
      .filter({ hasText: new RegExp(`^${day}$`) });
  }

  // The calendar opens on the current month; step back until the title (e.g. "August, 2026")
  // matches. Capped at 24 steps so a bad month/year fails instead of looping.
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

  // Without month/year, picks the day in the month the calendar opens on (the current month).
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

  // Sets From and To to the same day in the current month.
  async setDateRange(date: Date) {
    const day = String(date.getDate());
    await this.setFromDate(day);
    await this.setToDate(day);
  }

  // "Excel Download" also opens a blank popup. It is left open: in real Chrome, closing the
  // last other window can quit the browser and kill the download.
  async downloadExcelDirect(): Promise<Download> {
    console.log('[STEP] Downloading report directly via Excel Download...');
    const downloadPromise = this.page.waitForEvent('download');
    await this.excelDownloadLink.click();
    const download = await downloadPromise;

    // "Excel Download" is a postback; clicking "View Report" before it settles gets lost and
    // no report tab ever opens.
    await this.page.waitForTimeout(2000);

    return download;
  }

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
    const clickedAt = Date.now();
    await this.viewReportLink.click();
    const reportPage = await popupPromise;
    reportOpenedAt.set(reportPage, clickedAt);
    // Listen straight away: an over-limit report starts its Excel download on load.
    autoDownloads.set(reportPage, reportPage.waitForEvent('download', { timeout: 60000 }).catch(() => null));
    await reportPage.bringToFront();

    // The popup inherits the main window's maximized state; give it a normal size. Cosmetic,
    // skipped where there is no window or no CDP.
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

  // Saves the open report as Excel at `filePath` (extension taken from the DMS's file name).
  // Over the display limit the viewer downloads it by itself (or via "click here"); otherwise
  // "Export to Excel" is clicked. Returns the saved path.
  async saveReportExcel(reportPage: Page, filePath: string): Promise<string> {
    await reportPage.waitForLoadState('load').catch(() => {});
    const overLimit = await reportPage.getByText(/record display limit/i).isVisible().catch(() => false);
    let download: Download | null = null;
    if (overLimit) {
      download = await Promise.race([
        autoDownloads.get(reportPage) ?? Promise.resolve(null),
        reportPage.waitForTimeout(15000).then(() => null),
      ]);
      if (!download) {
        console.log('[STEP] Automatic download did not start; clicking "click here"...');
        const downloadPromise = reportPage.waitForEvent('download', { timeout: 120000 });
        await reportPage.getByRole('link', { name: 'click here' }).click();
        download = await downloadPromise;
      }
    } else {
      download = await this.exportToExcel(reportPage);
    }
    const ext = path.extname(download.suggestedFilename()) || '.xls';
    const target = filePath.replace(/\.xlsx?$/i, '') + ext;
    fs.mkdirSync(path.dirname(target), { recursive: true });
    await download.saveAs(target);
    return target;
  }

  async openColumnSelector(reportPage: Page) {
    await reportPage.locator('section').filter({ hasText: 'Select the columns to be' }).click();
  }

  async getRecordCount(reportPage: Page): Promise<number> {
    return (await this.getReportRows(reportPage)).length;
  }

  // Register the handler before the alert, or Playwright dismisses it unseen.
  async showRecordCountPopup(reportPage: Page, count: number) {
    const message = count === 0
      ? 'Warning: No Enquiry Found'
      : `Enquiry List record count: ${count}`;
    reportPage.once('dialog', (dialog) => dialog.accept());
    await reportPage.evaluate((msg) => window.alert(msg), message);
  }

  // Closing a report tab can raise a dialog from the page's unload handler, and close()
  // hangs unless it is accepted. Accept and log it.
  async closeReport(reportPage: Page, label: string) {
    reportPage.on('dialog', async (dialog) => {
      console.log(`[DEBUG] Dialog appeared while closing ${label} report viewer: type=${dialog.type()}, message="${dialog.message()}"`);
      await dialog.accept();
    });
    console.log(`[STEP] Closing the ${label} report viewer...`);
    await reportPage.close();
  }

  // By header text, since the selected columns and their order can change.
  async getDealerCodes(reportPage: Page): Promise<string[]> {
    const rows = await this.getReportRows(reportPage);
    if (rows.length === 0) return [];
    const dealerCodeKey = Object.keys(rows[0]!).find((h) => h.replace(/\s+/g, '') === 'DealerCode');
    if (!dealerCodeKey) {
      throw new Error('Dealer Code column not found in the report.');
    }
    return rows.map((r) => (r[dealerCodeKey] ?? '').trim());
  }

  // Rows keyed by header text, from the on-screen grid or, over the display limit, the Excel export.
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

  // Uses the automatic download if it fired, otherwise the banner's "click here". If Playwright
  // never gets the download, falls back to the newest Excel saved since the report opened.
  private async readOverLimitExcel(reportPage: Page): Promise<Record<string, string>[]> {
    const openedAt = reportOpenedAt.get(reportPage) ?? Date.now() - 5 * 60 * 1000;
    let download = await Promise.race([
      autoDownloads.get(reportPage) ?? Promise.resolve(null),
      reportPage.waitForTimeout(15000).then(() => null),
    ]);
    if (!download) {
      console.log('[STEP] Automatic download did not start; clicking "click here"...');
      const downloadPromise = reportPage.waitForEvent('download', { timeout: 120000 }).catch(() => null);
      await reportPage.getByRole('link', { name: 'click here' }).click().catch(() => {});
      download = await downloadPromise;
    }

    let filePath: string | null = null;
    if (download) {
      const savedPath = path.join(DOWNLOAD_DIR, `${Date.now()}_${download.suggestedFilename()}`);
      try {
        fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
        await download.saveAs(savedPath);
        filePath = savedPath;
      } catch (err) {
        console.log(`[WARN] Could not save the report Excel to ${DOWNLOAD_DIR} (${String((err as Error).message).slice(0, 80)}).`);
        filePath = await download.path().catch(() => null);
      }
    }
    if (!filePath) {
      filePath = newestExcelSince(openedAt);
      if (filePath) console.log('[WARN] The report download did not reach Playwright; using the newest Excel in the Downloads folder.');
    }
    if (!filePath) {
      throw new Error(`Over-limit report Excel download did not produce a file (nothing new in ${DOWNLOAD_DIR} or ~/Downloads either).`);
    }
    console.log(`[STEP] Reading over-limit report from Excel: ${filePath}`);

    // cellDates + dateNF keep dates in the grid's "DD-MM-YYYY hh:mm" format.
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]!]!;
    const table = XLSX.utils.sheet_to_json<string[]>(sheet, {
      header: 1,
      raw: false,
      defval: '',
      dateNF: 'dd-mm-yyyy hh:mm',
    });
    // The export can have title rows above the headers. The header row is the one with Dealer
    // Code, or for single-dealer reports (GSTR) the first full-width row.
    const filled = (r: string[]) => r.filter((c) => String(c).trim() !== '').length;
    let headerIndex = table.findIndex((r) => r.some((c) => String(c).replace(/\s+/g, '').toLowerCase() === 'dealercode'));
    if (headerIndex === -1) {
      const top = table.slice(0, 20);
      const widest = Math.max(0, ...top.map(filled));
      headerIndex = widest >= 3 ? top.findIndex((r) => filled(r) === widest) : -1;
    }
    if (headerIndex === -1) {
      throw new Error(`Header row not found in ${path.basename(filePath)}.`);
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

  // One evaluate() for the whole grid: per-row locator calls on ~1000 rows hit the test timeout.
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
