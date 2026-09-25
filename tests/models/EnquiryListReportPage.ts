import { Page, Locator } from '@playwright/test';
import { ReportViewerPage } from './ReportViewerPage';

export class EnquiryListReportPage extends ReportViewerPage {
  readonly misLink: Locator;
  readonly mfrVehiclesReportsLink: Locator;
  readonly preSalesLink: Locator;
  readonly enquiryListLink: Locator;

  constructor(page: Page) {
    super(page);
    this.misLink = page.getByRole('link', { name: ' MIS ' });
    this.mfrVehiclesReportsLink = page.getByRole('link', { name: ' MFR Vehicles Reports' });
    this.preSalesLink = page.locator('#TableContent').getByText('Pre-sales');
    this.enquiryListLink = page.locator('#TableContent').getByText('Enquiry ListThe list of all');
  }

  async navigateToEnquiryList() {
    console.log('[STEP] Navigating to MIS > MFR Vehicles Reports > Pre-sales > Enquiry List...');
    await this.misLink.click();
    await this.mfrVehiclesReportsLink.click();
    await this.preSalesLink.click();
    await this.enquiryListLink.click();
  }
}
