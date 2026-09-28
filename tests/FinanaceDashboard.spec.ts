import path from 'node:path';
import fs from 'node:fs/promises';
import { test } from '@playwright/test';
import { LoginPage } from './models/LoginPage';
import { ProductWiseSalesDetailsReportPage } from './models/ProductWiseSalesDetailsReportPage';
import { readExcelDownload } from './utils/excelExport';
import { pushRowsToSheet } from './utils/appsScriptSheet';
import credentialsRaw from '../resources/credentials.json' with { type: 'json' };

const credentials = credentialsRaw as any;

const todayDay = String(new Date().getDate());

const DOWNLOAD_DIR = 'C:\\Users\\sunita_rideriver\\Downloads\\Vehicle Tax Invoice';

test.describe('Finance Dashboard - Product-wise Sales Details report', () => {
  test.use({ storageState: { cookies: [], origins: [] } });

  test('Should view Product-wise Sales Details report by Invoice date type', async ({ page }) => {
    const loginPage = new LoginPage(page);
    const salesDetailsReportPage = new ProductWiseSalesDetailsReportPage(page);

    await loginPage.goto();
    await loginPage.loginAsRiverUser(credentials.riverUserValidUser);
    await loginPage.verifyDashboard();

    // --- Product-wise Sales Details ---
    await salesDetailsReportPage.navigateToProductWiseSalesDetails();
    await salesDetailsReportPage.setFromDate('1', 'September', 2026);
    await salesDetailsReportPage.setToDate(todayDay);

    const invoiceExcelDownload = await salesDetailsReportPage.downloadExcelDirect();

    await fs.mkdir(DOWNLOAD_DIR, { recursive: true });
    const savedPath = path.join(DOWNLOAD_DIR, invoiceExcelDownload.suggestedFilename());
    await invoiceExcelDownload.saveAs(savedPath);
    console.log(`[STEP] Saved report to ${savedPath}`);

    const { headers: excelHeaders, rows: excelRows } = await readExcelDownload(invoiceExcelDownload);
    await pushRowsToSheet(excelHeaders, excelRows);
  });
});
