import { Page, Locator } from '@playwright/test';
import { ReportViewerPage } from './ReportViewerPage';

export class ProductWiseSalesDetailsReportPage extends ReportViewerPage {
  readonly salesLink: Locator;
  readonly productWiseSalesDetailsLink: Locator;
  readonly dateTypeSelect: Locator;
  readonly misLink: Locator;
  readonly mfrVehiclesReportsLink: Locator;
  readonly leadReportsLink: Locator;

  constructor(page: Page) {
    super(page);
    this.misLink = page.getByRole('link', { name: ' MIS ' });
    this.mfrVehiclesReportsLink = page.getByRole('link', { name: ' MFR Vehicles Reports' });
    this.leadReportsLink = page.getByRole('link', { name: ' Enquiry Reports' });
    this.salesLink = page.locator('#TableContent').getByText('Sales', { exact: true });
    this.productWiseSalesDetailsLink = page.locator('#TableContent').getByText('Product-wise sales details');
    this.dateTypeSelect = page.locator('#ctl00_cpMain_datetype2');
  }

  async navigateToProductWiseSalesDetails() {
    console.log('[STEP] Navigating to Sales > Product-wise sales details...');

    // First pass: MIS > Lead Reports just to clear whatever submenu the
    // previous screen left expanded, so the second pass below reliably finds
    // Sales under MFR Vehicles Reports.
    await this.misLink.click();
    await this.leadReportsLink.click();
    console.log('[STEP] Clicked MIS > Enquiry Reports to reset the menu.');

    // Second pass: the actual path to Product-wise sales details.
    await this.misLink.click();
    await this.mfrVehiclesReportsLink.click();
    await this.salesLink.click();
    await this.productWiseSalesDetailsLink.click();
    console.log('[STEP] Clicked MIS > MFR Vehicles Reports > Sales > Product-wise sales details.');
  }

  async setDateType(option: string) {
    await this.dateTypeSelect.selectOption(option);
  }
}
