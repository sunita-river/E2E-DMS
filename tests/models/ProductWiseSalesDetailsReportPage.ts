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
    await this.clickMenuItem(this.misLink, this.leadReportsLink);
    console.log('[STEP] Clicked MIS > Enquiry Reports to reset the menu.');

    // Second pass: the actual path to Product-wise sales details.
    await this.clickMenuItem(this.misLink, this.mfrVehiclesReportsLink);
    await this.salesLink.click();
    await this.productWiseSalesDetailsLink.click();
    console.log('[STEP] Clicked MIS > MFR Vehicles Reports > Sales > Product-wise sales details.');
  }

  // MIS is a toggle: clicking it while its menu is already open collapses it, hiding the item
  // we want. The likely cause of a CI failure where "MFR Vehicles Reports" was never found on the
  // second navigation: MIS was still open, so the second MIS click closed it. Only click the
  // parent when the item isn't already showing.
  private async clickMenuItem(parent: Locator, item: Locator) {
    if (!(await item.isVisible().catch(() => false))) {
      await parent.click();
      await item.waitFor({ state: 'visible', timeout: 10000 }).catch(async () => {
        await parent.click(); // the click closed an open menu — open it again
      });
    }
    await item.click();
  }

  async setDateType(option: string) {
    await this.dateTypeSelect.selectOption(option);
  }
}
