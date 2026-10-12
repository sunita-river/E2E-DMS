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

    // Open MIS > Enquiry Reports first to collapse whatever submenu the previous screen left
    // expanded; otherwise Sales under MFR Vehicles Reports isn't reliably found.
    await this.clickMenuItem(this.misLink, this.leadReportsLink);
    console.log('[STEP] Clicked MIS > Enquiry Reports to reset the menu.');

    await this.clickMenuItem(this.misLink, this.mfrVehiclesReportsLink);
    await this.salesLink.click();
    await this.productWiseSalesDetailsLink.click();
    console.log('[STEP] Clicked MIS > MFR Vehicles Reports > Sales > Product-wise sales details.');
  }

  // MIS is a toggle: clicking it while its menu is open collapses it (this broke the second
  // navigation in CI). Only click the parent when the item isn't already showing.
  private async clickMenuItem(parent: Locator, item: Locator) {
    if (!(await item.isVisible().catch(() => false))) {
      await parent.click();
      await item.waitFor({ state: 'visible', timeout: 10000 }).catch(async () => {
        await parent.click(); // the click closed an open menu; open it again
      });
    }
    await item.click();
  }

  async setDateType(option: string) {
    await this.dateTypeSelect.selectOption(option);
  }
}
