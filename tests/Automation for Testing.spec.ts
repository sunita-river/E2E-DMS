import { test, expect, Page, BrowserContext, Locator } from '@playwright/test';
import { LoginPage } from './models/LoginPage';
import * as fs from 'node:fs';
import * as path from 'node:path';

// Stage DMS scenarios. The login is the River User sign-in with credentials.stageUser
// (resources/credentials.json, kept out of Git).
// These scenarios save real records (enquiry, follow-up, booking, receipt), so they only ever run against stage.
const STAGE_URL = 'https://stage.rivermobility.gaindms.com/';
const STAGE_HOST = new URL(STAGE_URL).hostname;
if (!STAGE_HOST.startsWith('stage.')) throw new Error(`STAGE_URL must be the stage DMS, not ${STAGE_HOST}.`);

const credentials = JSON.parse(fs.readFileSync(
  process.env.CREDENTIALS_FILE || path.join(__dirname, '..', 'resources', 'credentials.json'), 'utf8'));
const stageUser: { username: string; password: string } = credentials.stageUser;
if (!stageUser?.username || !stageUser?.password) {
  throw new Error('credentials.json needs a "stageUser": { "username", "password" } entry.');
}

const ddmmyyyy = (d: Date) =>
  `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
// Lists and the Activity Log show dates as dd/mm/yyyy.
const ddmmyyyySlash = (d: Date) => ddmmyyyy(d).replace(/-/g, '/');
const daysFromToday = (n: number) => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); };
const today = daysFromToday(0);
const tomorrow = daysFromToday(1);
const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// A unique name / mobile per run, so the saved enquiry can be found again and never clashes with an earlier one.
const runId = Date.now().toString().slice(-8);
// Only the fields the form marks mandatory (*); Date and Executive are pre-filled, Status is set from Enquiry Category.
const ENQUIRY = {
  primarySource: 'Walk-in',
  secondarySource: 'Walk-in',
  title: 'Mr.',
  name: `Auto Test ${runId}`,
  phone: `9${runId}0`,   // 10 digits, starts with 9
  state: /Karnataka/i,
  city: 'Bengaluru',
  pinCode: '560001',
  model: 'Gen',        // autosuggest: the first match is picked (stage models: Gen-1 / Gen-2 / Gen-3)
  product: '',         // '' = first product of the chosen model
  color: '',           // '' = first colour of the chosen product
  category: 'Hot – 30 days',
  status: 'Hot',       // follows the category
  stepToSale: 'Enquiry assigned',
  followUpTime: '10:00',
  loanRequired: 'No',
  summary: 'Created by Playwright automation (stage).',
};

// Scenario 3: the follow-up added to that enquiry (mandatory fields only; Follow up date/time,
// Followup by and Enquiry Category are pre-filled, Status follows the category).
const FOLLOW_UP = {
  callMode: 'Telephone',
  primaryReason: 'Needs More Time',
  secondaryReason: 'Customer asked to follow up later',
  callNote: `Follow-up by Playwright automation (stage) ${runId}`,
  stepToSale: 'Test Ride Scheduled',
  nextCallDate: daysFromToday(2),
  nextCallTime: '11:00',
};

// Scenario 4: booking made from that enquiry. Booking from an enquiry pre-fills 8 of the 11 mandatory
// fields (Sales exec, Customer, Phone, Place Of Supply, Pin Code, State, Product, Color); these are the rest.
const BOOKING = {
  financeType: 'Full cash',
  rtoType: 'Permanent',
  insuranceType: 'Insurance By Dealer',
};

// Scenario 5: receipt against that booking (Amount and Narration; everything else on the receipt is pre-filled).
const RECEIPT = {
  amount: 5000,
  narration: `Booking advance by Playwright automation (stage) ${runId}`,
};

// Progress through the scenarios: each one checks the step before it finished, so running one on its own
// (e.g. -g TC03) fails with a clear message instead of a timeout.
const done = { login: false, enquiry: false, followUp: false };
let bookingNo = '';     // saved by TC04, used by TC05
const receipt = { no: '', pageUrl: '', customerCenterUrl: '', total: 0 }; // saved by TC05, cancelled by TC06

// Amount boxes may show thousands separators ("1,23,456.00").
async function amountOf(field: Locator, label: string) {
  const value = Number((await field.inputValue()).replace(/,/g, ''));
  expect(Number.isFinite(value), `${label} is not a number`).toBe(true);
  return value;
}

async function loginToStage(page: Page) {
  const loginPage = new LoginPage(page);
  await loginPage.goto(STAGE_URL);
  await loginPage.loginAsRiverUser(stageUser);
  // LoginPage.verifyDashboard() also expects a "Sales/Service Dashboard" link, which the stage
  // River User dashboard doesn't have, so check for the dashboard page, the user menu and the main menu instead.
  await expect(page, 'Login failed: still on the login page (check stageUser in credentials.json).')
    .toHaveURL(/userdashboard\.aspx/i, { timeout: 30000 });
  await expect(page.getByRole('link', { name: stageUser.username }).first()).toBeVisible();
  await expect(page.getByRole('link', { name: /PRESALES/ }).first()).toBeVisible();
}

// Presales > My Enquiries, filtered by mobile number. Returns the results grid.
async function searchMyEnquiries(page: Page, phone: string) {
  await page.goto(`${STAGE_URL}Presale/EnquiryWorksheet.aspx`);
  await page.locator('#ctl00_cpMain_SearchPhone').fill(phone);
  await expect(page.locator('#ctl00_cpMain_SearchPhone')).toHaveValue(phone);
  await page.locator('#ctl00_cpMain_ApplyFilterTop').click();
  await page.waitForLoadState('domcontentloaded');
  const results = page.locator('#ctl00_cpMain_DLMain');
  await expect(results).toContainText(/Displaying \d+ records/, { timeout: 30000 });
  return results;
}

// The run's enquiry row in a My Enquiries search result.
const enquiryRow = (page: Page, results: Locator) =>
  results.locator('tr').filter({ has: page.locator('a#hSelect', { hasText: ENQUIRY.name }) });

// GainDMS autosuggest fields: typing fires a lookup that fills #ctl00_MP_DivAutosuggest; Down + Enter
// picks the first row and writes its id into the hidden field named by vtb= in the field's onkeyup.
async function pickAutosuggest(page: Page, fieldId: string, text: string, label: string) {
  const field = page.locator(`#${fieldId}`);
  await expect(field, `${label} field`).toBeEditable();
  const hiddenId = ((await field.getAttribute('onkeyup')) || '').match(/vtb=([\w$]+)/)?.[1];
  const suggestions = page.locator('#ctl00_MP_DivAutosuggest');

  await field.click();
  await field.fill('');
  // pressSequentially, not fill(): the lookup only runs on keyup. An empty search text still needs one
  // keystroke, so type a space and let the field's own filter (model / product) narrow the list.
  await field.pressSequentially(text || ' ', { delay: 150 });
  await expect(suggestions.locator('tr, li').first(), `No ${label} suggestions for "${text}"; does this user have a dealer/branch with ${label}s set up?`)
    .toBeVisible({ timeout: 15000 });
  await field.press('ArrowDown');
  await field.press('Enter');

  if (hiddenId) {
    await expect(page.locator(`#${hiddenId}`), `${label} was typed but not selected from the list`).not.toHaveValue(/^0?$/);
  }
  const picked = (await field.inputValue()).trim();
  expect(picked, `${label} picked from the list`).not.toBe('');
  if (text) expect(picked.toLowerCase(), `${label} matches "${text}"`).toContain(text.toLowerCase());
  console.log(`[STEP] ${label}: ${picked}`);
  return picked;
}

// Presales > My Enquiries > New, on a fresh page (after Apply Filter the filter side panel can sit over the buttons).
async function openNewEnquiry(page: Page) {
  console.log('[STEP] Opening Presales > My Enquiries > New');
  // A search that finds nothing can send the page on to the New Enquiry form by itself, which
  // interrupts this goto ("interrupted by another navigation"). Let that finish, then go again.
  await page.waitForLoadState('load');
  await page.goto(`${STAGE_URL}Presale/EnquiryWorksheet.aspx`).catch(async (err) => {
    if (!/interrupted by another navigation/.test(String(err))) throw err;
    await page.waitForLoadState('load');
    await page.goto(`${STAGE_URL}Presale/EnquiryWorksheet.aspx`);
  });
  await page.locator('#ctl00_cpMain_cmdNew').click();
  await expect(page).toHaveURL(/Presale\/Enquiry\.aspx/i);
  await expect(page.locator('#ctl00_cpMain_EnquiryName')).toHaveValue('');
  await expect(page.locator('#ctl00_cpMain_CmdEnqSave')).toBeVisible();
}

// Fills every mandatory field of the New Enquiry form from `data` (the validation tests pass a copy
// of ENQUIRY with one field made wrong).
async function fillEnquiryForm(page: Page, data: typeof ENQUIRY) {
  // Primary DataSource posts back to load the Secondary DataSource list.
  await page.locator('#ctl00_cpMain_DataSourceCategorySlno').selectOption({ label: data.primarySource });
  await page.waitForLoadState('domcontentloaded');
  await expect(page.locator('#ctl00_cpMain_DatasourceSlno option', { hasText: data.secondarySource }).first()).toBeAttached({ timeout: 30000 });
  await page.locator('#ctl00_cpMain_DatasourceSlno').selectOption({ label: data.secondarySource });
  await page.waitForLoadState('domcontentloaded');

  console.log(`[STEP] Customer: ${data.name}, ${data.phone}`);
  await page.locator('#ctl00_cpMain_TitleSlno').selectOption({ label: data.title });
  await page.locator('#ctl00_cpMain_EnquiryName').fill(data.name);
  await page.locator('#ctl00_cpMain_PhoneM').fill(data.phone);

  // State labels look like "29(Karnataka)", so match on the name.
  const stateSelect = page.locator('#ctl00_cpMain_StateSlno');
  const stateOption = stateSelect.locator('option').filter({ hasText: data.state }).first();
  await expect(stateOption, `State ${data.state} in the list`).toBeAttached();
  await stateSelect.selectOption((await stateOption.getAttribute('value'))!);
  await page.waitForLoadState('domcontentloaded');
  await pickAutosuggest(page, 'ctl00_cpMain_CityName', data.city, 'City');
  await page.locator('#ctl00_cpMain_PinCode').fill(data.pinCode);

  await pickAutosuggest(page, 'ctl00_cpMain_Model', data.model, 'Model');
  await pickAutosuggest(page, 'ctl00_cpMain_Product', data.product, 'Product');
  await pickAutosuggest(page, 'ctl00_cpMain_Color', data.color, 'Color');

  await page.locator('#ctl00_cpMain_ExpPurchaseTime').selectOption({ label: data.category });
  await page.locator('#ctl00_cpMain_EnquirySteptosaleSlno').selectOption({ label: data.stepToSale });
  await page.locator('#ctl00_cpMain_CallBackOnDate').fill(ddmmyyyy(tomorrow));
  await page.locator('#ctl00_cpMain_CallBackOnDate').press('Tab');
  await page.locator('#ctl00_cpMain_CallBackOnTime').fill(data.followUpTime);
  await page.locator('#ctl00_cpMain_CallSummary').fill(data.summary);
  await page.locator('#ctl00_cpMain_LoanRequired').selectOption({ label: data.loanRequired });
}

test.describe('Stage DMS - Automation for Testing', () => {
  // Each scenario works on what the one before it saved, so run in order and skip the rest once one fails.
  test.describe.configure({ mode: 'serial' });

  // One browser session for the whole group: TC01 logs in and the later scenarios carry on as that user.
  let context: BrowserContext;
  let page: Page;
  // Save-time validation messages come as alert()s, from the main page or a tab it opens (the booking form);
  // each scenario clears the list and keeps them for its failure messages.
  const alerts: string[] = [];
  const alertText = () => alerts.join(' | ') || 'none';
  const collectAlerts = (p: Page) => p.on('dialog', async (dialog) => {
    alerts.push(dialog.message());
    await dialog.accept().catch(() => {}); // already closed by the page
  });

  test.beforeAll(async ({ browser }) => {
    // Same window handling as the config's projects (the LoginPage maximizes it).
    context = await browser.newContext({ storageState: { cookies: [], origins: [] }, viewport: null });
    context.on('page', collectAlerts);
    page = await context.newPage();
  });
  test.beforeEach(() => { alerts.length = 0; });
  // The shared page isn't Playwright's own fixture, so its screenshot-on-failure doesn't cover it; attach one here.
  test.afterEach(async ({}, testInfo) => {
    if (testInfo.status === testInfo.expectedStatus) return;
    for (const [i, p] of context.pages().entries()) {
      const shot = await p.screenshot({ fullPage: true }).catch(() => undefined);
      if (shot) await testInfo.attach(`failure-tab-${i + 1}`, { body: shot, contentType: 'image/png' });
    }
    if (alerts.length) await testInfo.attach('alerts', { body: alertText(), contentType: 'text/plain' });
  });
  test.afterAll(async () => { await context?.close(); });

  // Scenario 1
  test('TC01 - Login to stage DMS as River User', async () => {
    await loginToStage(page);

    // Signed in on stage (never production), with a session and no login form left on the page.
    expect(new URL(page.url()).hostname, 'Logged in on the stage DMS').toBe(STAGE_HOST);
    await expect(page.locator('#UserName')).toBeHidden();
    await expect(page.locator('#CmdLogin')).toBeHidden();
    expect((await context.cookies(STAGE_URL)).length, 'Session cookie after login').toBeGreaterThan(0);
    expect(alerts, 'Alerts during login').toEqual([]);

    done.login = true;
    console.log('[DONE] Logged in to stage DMS');
  });

  // Scenario 2
  test('TC02 - Create a new Enquiry', async () => {
    expect(done.login, 'TC01 must log in first').toBe(true);
    expect(ENQUIRY.phone, 'Generated mobile number').toMatch(/^9\d{9}$/);

    // The mobile number is new, so whatever is found after the save is this run's enquiry.
    const before = await searchMyEnquiries(page, ENQUIRY.phone);
    await expect(before, `An enquiry with ${ENQUIRY.phone} already exists`).toContainText('Displaying 0 records');

    await openNewEnquiry(page);
    await fillEnquiryForm(page, ENQUIRY);

    // Everything entered is still on the form (a postback can wipe fields filled before it).
    const f = (id: string) => page.locator(`#ctl00_cpMain_${id}`);
    await expect(f('DataSourceCategorySlno').locator('option:checked')).toHaveText(ENQUIRY.primarySource);
    await expect(f('DatasourceSlno').locator('option:checked')).toHaveText(ENQUIRY.secondarySource);
    await expect(f('TitleSlno').locator('option:checked')).toHaveText(ENQUIRY.title);
    await expect(f('EnquiryName')).toHaveValue(ENQUIRY.name);
    await expect(f('PhoneM')).toHaveValue(ENQUIRY.phone);
    await expect(f('StateSlno').locator('option:checked')).toHaveText(ENQUIRY.state);
    await expect(f('CityName')).toHaveValue(ENQUIRY.city);
    await expect(f('PinCode')).toHaveValue(ENQUIRY.pinCode);
    await expect(f('Model')).toHaveValue(new RegExp(`^${escapeRegExp(ENQUIRY.model)}`, 'i'));
    await expect(f('Product')).not.toHaveValue('');
    await expect(f('Color')).not.toHaveValue('');
    await expect(f('ExpPurchaseTime').locator('option:checked')).toHaveText(ENQUIRY.category);
    await expect(f('EnquirySteptosaleSlno').locator('option:checked')).toHaveText(ENQUIRY.stepToSale);
    await expect(f('CallBackOnDate')).toHaveValue(ddmmyyyy(tomorrow));
    await expect(f('CallBackOnTime')).toHaveValue(ENQUIRY.followUpTime);
    await expect(f('CallSummary')).toHaveValue(ENQUIRY.summary);
    await expect(f('LoanRequired').locator('option:checked')).toHaveText(ENQUIRY.loanRequired);

    // Wait for the save postback itself, rather than a fixed pause: leaving the page earlier could cancel the save.
    // No postback at all means a client-side check stopped it (with an alert).
    console.log('[STEP] Saving enquiry...');
    const savePostback = page.waitForResponse(
      (r) => r.request().method() === 'POST' && (r.request().postData() || '').includes('CmdEnqSave'), { timeout: 60000 });
    await f('CmdEnqSave').click();
    const saveResponse = await savePostback.catch(() => {
      throw new Error(`Enquiry save was not sent. Alerts: ${alertText()}`);
    });
    expect(saveResponse.status(), 'Enquiry save response').toBeLessThan(400);
    await page.waitForLoadState('domcontentloaded');
    if (alerts.length) console.log(`[INFO] Alerts on save: ${alertText()}`);

    // Verify: exactly one enquiry with this mobile, with the details entered above.
    console.log('[STEP] Verifying the enquiry in My Enquiries...');
    const results = await searchMyEnquiries(page, ENQUIRY.phone);
    await expect(results, `Enquiry not found after save. Alerts: ${alertText()}`).toContainText(ENQUIRY.name);
    await expect(results).toContainText('Displaying 1 records');
    const row = enquiryRow(page, results);
    await expect(row).toHaveCount(1);
    await expect(row).toContainText(ENQUIRY.phone);
    await expect(row).toContainText(ddmmyyyySlash(today));    // enquiry date
    await expect(row).toContainText(ddmmyyyySlash(tomorrow)); // next follow-up date
    await expect(row).toContainText(ENQUIRY.status);

    done.enquiry = true;
    console.log(`[DONE] Enquiry created for ${ENQUIRY.name} (${ENQUIRY.phone})`);
  });

  // Scenario 3
  test('TC03 - Add a follow-up to the enquiry created in TC02', async () => {
    expect(done.enquiry, 'TC02 must create the enquiry first').toBe(true);

    console.log(`[STEP] Finding ${ENQUIRY.name} in My Enquiries...`);
    const results = await searchMyEnquiries(page, ENQUIRY.phone);
    await expect(results).toContainText('Displaying 1 records');
    const row = enquiryRow(page, results);
    await expect(row).toHaveCount(1);
    await expect(row).toContainText(ENQUIRY.phone);
    await expect(row).toContainText(ddmmyyyySlash(today));    // enquiry date
    await expect(row).toContainText(ddmmyyyySlash(tomorrow)); // next follow-up date set in TC02
    await expect(row).toContainText(ENQUIRY.status);

    console.log('[STEP] Opening the enquiry...');
    await row.locator('a#hSelect').click();
    await expect(page).toHaveURL(/Presale\/Enquiry\.aspx\?keyid=\d+/i);
    const enquiryUrl = page.url();

    // The saved enquiry details come back as entered in TC02.
    const f = (id: string) => page.locator(`#ctl00_cpMain_${id}`);
    await expect(f('EnquiryName')).toHaveValue(ENQUIRY.name);
    await expect(f('PhoneM')).toHaveValue(ENQUIRY.phone);
    await expect(f('TitleSlno').locator('option:checked')).toHaveText(ENQUIRY.title);
    await expect(f('CityName')).toHaveValue(ENQUIRY.city);
    await expect(f('PinCode')).toHaveValue(ENQUIRY.pinCode);
    await expect(f('Model')).toHaveValue(new RegExp(`^${escapeRegExp(ENQUIRY.model)}`, 'i'));
    await expect(f('Product')).not.toHaveValue('');
    await expect(f('Color')).not.toHaveValue('');
    await expect(f('DataSourceCategorySlno')).toBeDisabled(); // fixed once saved
    await expect(f('LoanRequired').locator('option:checked')).toHaveText(ENQUIRY.loanRequired);

    // Follow-up form: pre-filled fields.
    console.log('[STEP] Filling the follow-up...');
    await page.getByRole('link', { name: 'Follow Up', exact: true }).click();
    const updateFollowUp = f('CmdSaveFollowUp');
    await expect(updateFollowUp).toBeVisible();
    await expect(updateFollowUp).toHaveValue('Update Followup');
    await expect(f('FollowUpDate')).toHaveValue(ddmmyyyy(today));
    await expect(f('FollowUpTime')).toHaveValue(/^\d{1,2}:\d{2}$/);
    await expect(f('FollowUpHandledBy')).toHaveValue(stageUser.username);
    await expect(f('FollowupExpPurchasetime').locator('option:checked')).toHaveText(ENQUIRY.category);
    await expect(f('FollowUpStatusSlno')).toBeDisabled();
    await expect(f('FollowUpStatusSlno').locator('option:checked')).toHaveText(ENQUIRY.status);
    await expect(f('FollowupSummary')).toHaveValue('');

    // Mandatory fields.
    await f('FollowUpCallMode').selectOption({ label: FOLLOW_UP.callMode });
    await expect(f('FollowUpCallMode').locator('option:checked')).toHaveText(FOLLOW_UP.callMode);

    // Primary reason loads the Secondary reason list.
    await f('PrimaryFolowUpReasonSlno').selectOption({ label: FOLLOW_UP.primaryReason });
    await page.waitForLoadState('domcontentloaded');
    const secondaryReason = f('SecondaryReasonSlno');
    await expect(secondaryReason.locator('option', { hasText: FOLLOW_UP.secondaryReason })).toBeAttached({ timeout: 30000 });
    await secondaryReason.selectOption({ label: FOLLOW_UP.secondaryReason });
    await expect(f('PrimaryFolowUpReasonSlno').locator('option:checked')).toHaveText(FOLLOW_UP.primaryReason);
    await expect(secondaryReason.locator('option:checked')).toHaveText(FOLLOW_UP.secondaryReason);

    await f('FollowupSummary').fill(FOLLOW_UP.callNote);
    await f('StepTosaleSlno').selectOption({ label: FOLLOW_UP.stepToSale });
    await f('NextCallDate').fill(ddmmyyyy(FOLLOW_UP.nextCallDate));
    await f('NextCallDate').press('Tab');
    await f('NextCallTime').fill(FOLLOW_UP.nextCallTime);
    await expect(f('FollowUpCallMode').locator('option:checked')).toHaveText(FOLLOW_UP.callMode);
    await expect(f('FollowupSummary')).toHaveValue(FOLLOW_UP.callNote);
    await expect(f('StepTosaleSlno').locator('option:checked')).toHaveText(FOLLOW_UP.stepToSale);
    await expect(f('NextCallDate')).toHaveValue(ddmmyyyy(FOLLOW_UP.nextCallDate));
    await expect(f('NextCallTime')).toHaveValue(FOLLOW_UP.nextCallTime);

    console.log('[STEP] Clicking Update Followup...');
    await updateFollowUp.click();
    await page.waitForLoadState('domcontentloaded');

    // Saved: the "Followup updated" message, no validation alert, the form is cleared, and the Activity Log has the new entry.
    await expect(f('LblFMsg'), `Follow-up not saved. Alerts: ${alertText()}`)
      .toHaveText('Followup updated', { timeout: 30000 });
    await expect(f('LblFMsg')).toBeVisible();
    await expect(f('FollowupSummary')).toHaveValue('');
    expect(alerts, 'Validation alerts on Update Followup').toEqual([]);
    expect(page.url(), 'Still on the same enquiry').toBe(enquiryUrl);

    const activityLog = f('GrdMain');
    await expect(activityLog).toBeVisible();
    const entry = activityLog.locator('table').filter({ hasText: FOLLOW_UP.callNote });
    await expect(entry, 'One Activity Log entry for this follow-up').toHaveCount(1);
    await expect(entry).toContainText(ddmmyyyySlash(today));                   // call date
    await expect(entry).toContainText(ddmmyyyySlash(FOLLOW_UP.nextCallDate));  // call on
    await expect(entry).toContainText(FOLLOW_UP.stepToSale);

    // My Enquiries now shows the new next follow-up date for this enquiry.
    console.log('[STEP] Verifying My Enquiries shows the new follow-up date...');
    const after = await searchMyEnquiries(page, ENQUIRY.phone);
    await expect(after).toContainText('Displaying 1 records');
    const updatedRow = enquiryRow(page, after);
    await expect(updatedRow).toHaveCount(1);
    await expect(updatedRow).toContainText(ddmmyyyySlash(FOLLOW_UP.nextCallDate));
    await expect(updatedRow).not.toContainText(ddmmyyyySlash(tomorrow));

    done.followUp = true;
    console.log(`[DONE] Follow-up added to ${ENQUIRY.name}: next call ${ddmmyyyySlash(FOLLOW_UP.nextCallDate)} ${FOLLOW_UP.nextCallTime}`);
  });

  // Scenario 4
  test('TC04 - Create a booking from the enquiry created in TC02', async () => {
    expect(done.enquiry, 'TC02 must create the enquiry first').toBe(true);

    console.log(`[STEP] Selecting ${ENQUIRY.name} in My Enquiries...`);
    const results = await searchMyEnquiries(page, ENQUIRY.phone);
    await expect(results).toContainText('Displaying 1 records');
    const row = enquiryRow(page, results);
    await expect(row).toHaveCount(1);
    const selectBox = row.locator('input[type=checkbox]');
    await selectBox.check();
    await expect(selectBox).toBeChecked();

    const bookingButton = row.locator('a#hBooking');
    await expect(bookingButton).toBeVisible();
    await expect(bookingButton).toHaveText('Booking');

    console.log('[STEP] Clicking Booking...');
    // Opens Product/ProductSaleBooking.aspx in a new tab, unless the dealer requires a quotation first,
    // in which case it only shows an alert.
    // The filter side panel overlaps the right-hand row buttons after Apply Filter (and its Hide link
    // doesn't close it under automation), so fire the button's click event instead of a mouse click.
    const bookingTabPromise = context.waitForEvent('page', { timeout: 30000 });
    await bookingButton.dispatchEvent('click');
    const booking = await bookingTabPromise.catch(() => {
      throw new Error(`Booking form did not open. Alerts: ${alertText()}`);
    });

    let product = '';
    let total = 0;
    try {
      await booking.waitForLoadState('domcontentloaded');
      await expect(booking).toHaveURL(/Product\/ProductSaleBooking\.aspx\?OppurtunitySlno=\d+&CustomerSlno=\d+/i);
      await expect(booking).toHaveTitle('Booking Info');

      // Pre-filled from the enquiry.
      const field = (id: string) => booking.locator(`#ctl00_cpMain_${id}`);
      await expect(field('BookingDate')).toHaveValue(ddmmyyyy(today));
      await expect(field('EmployeeName')).not.toHaveValue('');
      await expect(field('TitleSlno').locator('option:checked')).toHaveText(ENQUIRY.title);
      await expect(field('CustomerName')).toHaveValue(ENQUIRY.name);
      await expect(field('PhoneM')).toHaveValue(ENQUIRY.phone);
      await expect(field('PlaceOfSupplySlno').locator('option:checked')).toHaveText(ENQUIRY.state);
      await expect(field('StateSlno').locator('option:checked')).toHaveText(ENQUIRY.state);
      await expect(field('CityName')).toHaveValue(ENQUIRY.city);
      await expect(field('PinCode')).toHaveValue(ENQUIRY.pinCode);
      await expect(field('Product')).not.toHaveValue('');
      await expect(field('Color')).not.toHaveValue('');
      await expect(field('DatasourceSlno').locator('option:checked')).toHaveText(ENQUIRY.secondarySource);
      await expect(field('DatasourceSlno')).toBeDisabled(); // comes from the enquiry
      product = await field('Product').inputValue();
      console.log(`[STEP] Booking for ${ENQUIRY.name}: ${product} / ${await field('Color').inputValue()}`);

      // The remaining mandatory fields.
      console.log('[STEP] Filling Finance Type, RTO Type and Insurance Type...');
      for (const [id, label] of [['FinanceType', BOOKING.financeType], ['CustomerRTo', BOOKING.rtoType], ['CustomerInsurance', BOOKING.insuranceType]] as const) {
        await field(id).selectOption({ label });
        await booking.waitForLoadState('domcontentloaded');
        await expect(field(id).locator('option:checked')).toHaveText(label);
      }
      // A later postback must not have reset an earlier choice or the enquiry details.
      await expect(field('FinanceType').locator('option:checked')).toHaveText(BOOKING.financeType);
      await expect(field('CustomerRTo').locator('option:checked')).toHaveText(BOOKING.rtoType);
      await expect(field('CustomerName')).toHaveValue(ENQUIRY.name);
      await expect(field('Product')).toHaveValue(product);

      console.log('[STEP] Saving booking...');
      const saveButton = field('cmdSave');
      await expect(saveButton).toBeEnabled();
      await expect(saveButton).toHaveText(/Save Booking/i);
      await saveButton.click();

      // Saved: the app moves straight on to the booking's Receipt page.
      await expect(booking, `Booking not saved. Alerts: ${alertText()}`)
        .toHaveURL(/Accounts\/Receipt\.aspx\?transactionslno=\d+/i, { timeout: 60000 });
      expect(alerts, 'Validation alerts on Save Booking').toEqual([]);
      await expect(booking).toHaveTitle('Receipt');
      await expect(field('CustomerName')).toHaveValue(ENQUIRY.name);
      await expect(field('NameInReceipt')).toHaveValue(ENQUIRY.name);
      await expect(field('CustomerStateCode')).toHaveValue(ENQUIRY.state);
      await expect(field('ReceiptDate')).toHaveValue(ddmmyyyy(today));
      await expect(field('ReceiptMode').locator('option:checked')).toHaveText('Complete Product Sale Value');
      const savedNo = await field('ReferenceNo').inputValue();
      expect(savedNo, 'Booking number on the receipt').toMatch(/\S+\/\d+\/\S+/);
      total = await amountOf(field('TotalAmount'), 'Booking total');
      expect(total, 'Booking total amount').toBeGreaterThan(0);
      expect(await amountOf(field('BalanceAmount'), 'Balance'), 'Balance = total (nothing paid yet)').toBe(total);
      bookingNo = savedNo;
    } finally {
      await booking.close();
    }

    // The booked enquiry is no longer an open enquiry in My Enquiries.
    console.log('[STEP] Verifying the enquiry has left the open list...');
    const after = await searchMyEnquiries(page, ENQUIRY.phone);
    await expect(after).toContainText('Displaying 0 records');
    await expect(after).not.toContainText(ENQUIRY.name);
    console.log(`[DONE] Booking ${bookingNo} created for ${ENQUIRY.name}: ${product}, total ₹${total}`);
  });

  // Scenario 5
  test('TC05 - Make a receipt for the booking created in TC04', async () => {
    expect(bookingNo, 'TC04 must save a booking first').not.toBe('');

    console.log(`[STEP] Sales > Vehicle sales: finding booking ${bookingNo}...`);
    await page.goto(`${STAGE_URL}Product/ProductSaleWorkSheet.aspx`);
    await expect(page).toHaveTitle('Product Sale Worksheet');
    await page.locator('#ctl00_cpMain_Phone').fill(ENQUIRY.phone);
    await page.locator('#ctl00_cpMain_ApplyFilterTop').click();
    await page.waitForLoadState('domcontentloaded');
    const bookingLink = page.locator('a[href*="CustomerCenter.aspx?productsaleslno="]', { hasText: bookingNo });
    await expect(bookingLink, `Booking ${bookingNo} in Vehicle sales`).toHaveCount(1);
    await bookingLink.dispatchEvent('click'); // same filter side panel as My Enquiries can overlap the list

    // Customer Center for the booking, before the receipt.
    await expect(page).toHaveURL(/Product\/CustomerCenter\.aspx\?productsaleslno=\d+/i);
    await expect(page).toHaveTitle('Customer Center');
    const customerCenterUrl = page.url();
    const f = (id: string) => page.locator(`#ctl00_cpMain_${id}`);
    await expect(f('code')).toHaveValue(bookingNo);
    await expect(f('CustomerName')).toHaveValue(ENQUIRY.name);
    await expect(f('ContactNo')).toHaveValue(ENQUIRY.phone);
    await expect(f('BookingDate')).toHaveValue(ddmmyyyySlash(today));
    const total = await amountOf(f('TotalPayable'), 'Product Total');
    expect(total, 'Product Total').toBeGreaterThan(0);
    expect(await amountOf(f('PaidAmount'), 'Total Paid'), 'Total Paid before the receipt').toBe(0);
    expect(await amountOf(f('BalanceAmount'), 'Balance'), 'Balance before the receipt').toBe(total);
    expect(RECEIPT.amount, 'Receipt amount must be part of the total').toBeLessThanOrEqual(total);

    console.log('[STEP] Clicking Make Receipt...');
    const makeReceipt = f('cmdStatus');
    await expect(makeReceipt).toBeVisible();
    await expect(makeReceipt).toHaveText(/Make Receipt/);
    await makeReceipt.click();
    await expect(page).toHaveURL(/Accounts\/Receipt\.aspx\?transactionslno=\d+/i, { timeout: 60000 });
    await expect(page).toHaveTitle('Receipt');

    // Pre-filled receipt details.
    await expect(f('ReferenceNo')).toHaveValue(bookingNo);
    await expect(f('CustomerName')).toHaveValue(ENQUIRY.name);
    await expect(f('NameInReceipt')).toHaveValue(ENQUIRY.name);
    await expect(f('CustomerStateCode')).toHaveValue(ENQUIRY.state);
    await expect(f('ReceiptMode').locator('option:checked')).toHaveText('Complete Product Sale Value');
    await expect(f('ReceiptDate')).toHaveValue(ddmmyyyy(today));
    await expect(f('TotalAmount')).toHaveValue(total.toFixed(2));
    await expect(f('PaidAmount')).toHaveValue('0.00');
    await expect(f('BalanceAmount')).toHaveValue(total.toFixed(2));
    for (const id of ['ForBranchSlno', 'ForFranchiseSlno', 'CompanySlno']) {
      await expect(f(id).locator('option:checked'), `${id} pre-selected`).not.toHaveText(/^\s*\.?\s*$/);
    }
    await expect(f('AccountCode')).not.toHaveValue('');
    await expect(f('PaymentModeSlno_0')).toBeChecked(); // Cash
    await expect(f('ReceiptAmount')).toHaveValue('');
    await expect(f('Narration')).toHaveValue('');

    console.log(`[STEP] Receipt: ₹${RECEIPT.amount}, "${RECEIPT.narration}"`);
    await f('ReceiptAmount').fill(String(RECEIPT.amount));
    await f('Narration').fill(RECEIPT.narration);
    await expect(f('ReceiptAmount')).toHaveValue(String(RECEIPT.amount));
    await expect(f('Narration')).toHaveValue(RECEIPT.narration);

    console.log('[STEP] Saving receipt...');
    await expect(f('cmdSave')).toBeEnabled();
    await expect(f('cmdSave')).toHaveText('SAVE');
    await f('cmdSave').click();

    // Saved: "Last receipt saved was <receipt no>", and the receipt is listed against the booking.
    const savedMessage = page.getByText(/Last receipt saved was \S+/);
    await expect(savedMessage, `Receipt not saved. Alerts: ${alertText()}`).toBeVisible({ timeout: 60000 });
    expect(alerts, 'Validation alerts on receipt Save').toEqual([]);
    const receiptNo = (await savedMessage.innerText()).match(/Last receipt saved was (\S+)/)?.[1] ?? '';
    expect(receiptNo, 'Receipt number').toMatch(/^\w+$/);
    Object.assign(receipt, { no: receiptNo, pageUrl: page.url(), customerCenterUrl, total });
    await expect(page).toHaveURL(/Accounts\/Receipt\.aspx/i);
    await expect(f('ReceiptAmount')).toHaveValue(''); // form cleared for the next receipt
    await expect(f('Narration')).toHaveValue('');

    const receiptRow = page.locator('tr').filter({ hasText: receiptNo }).filter({ hasText: bookingNo }).last();
    await expect(receiptRow).toBeVisible();
    await expect(receiptRow).toContainText(ddmmyyyySlash(today));
    await expect(receiptRow).toContainText(RECEIPT.amount.toFixed(2));
    await expect(receiptRow).toContainText('Complete Product Sale Value');
    await expect(receiptRow.getByText('Cancel')).toBeVisible();
    await expect(receiptRow.getByText('Print')).toBeVisible();

    // Customer Center now shows the payment against the booking.
    console.log('[STEP] Verifying Total Paid / Balance in Customer Center...');
    await page.goto(customerCenterUrl);
    await expect(f('code')).toHaveValue(bookingNo);
    expect(await amountOf(f('TotalPayable'), 'Product Total'), 'Product Total unchanged').toBe(total);
    await expect(f('PaidAmount')).toHaveValue(RECEIPT.amount.toFixed(2));
    await expect(f('BalanceAmount')).toHaveValue((total - RECEIPT.amount).toFixed(2));
    console.log(`[DONE] Receipt ${receiptNo} for ₹${RECEIPT.amount} saved against booking ${bookingNo}; balance ₹${(total - RECEIPT.amount).toFixed(2)}`);
  });

  // Scenario 6: clean-up, so each run doesn't leave a paid booking behind on stage. The stage River User
  // can't cancel receipts or bookings yet (both open "You do not have permission"), so this is skipped
  // with that reason until a DMS admin grants it; then the cancel form steps go here.
  test('TC06 - Cancel the receipt made in TC05', async () => {
    expect(receipt.no, 'TC05 must save a receipt first').not.toBe('');

    console.log(`[STEP] Cancelling receipt ${receipt.no}...`);
    await page.goto(receipt.pageUrl);
    const row = page.locator('tr').filter({ hasText: receipt.no }).filter({ hasText: bookingNo }).last();
    await expect(row, `Receipt ${receipt.no} listed on the receipt page`).toBeVisible({ timeout: 30000 });
    const cancel = row.locator('a[id$="_hReceiptCancel"]');
    await expect(cancel).toHaveAttribute('href', /ReceiptCancel\.aspx\?receiptslno=\d+/);
    await cancel.click();
    await page.waitForLoadState('load');

    const noPermission = /ErrorPage\.aspx\?error=np/i.test(page.url());
    test.skip(noPermission, `Stage user ${stageUser.username} has no permission to cancel receipts (ReceiptCancel.aspx). ` +
      `Booking ${bookingNo} and receipt ${receipt.no} stay on stage. Ask the DMS admin for receipt and booking cancel rights.`);
    throw new Error(`Receipt cancel page opened (${page.url()}): the cancel form steps still need to be written.`);
  });
});

// Validation on the New Enquiry form. Nothing here is saved: each case checks the DMS refuses it.
test.describe('Stage DMS - Enquiry form validation', () => {
  test.describe.configure({ mode: 'serial' });
  let page: Page;
  const alerts: string[] = [];
  const enquirySavePosts: string[] = [];

  test.beforeAll(async ({ browser }) => {
    const context = await browser.newContext({ storageState: { cookies: [], origins: [] }, viewport: null });
    page = await context.newPage();
    page.on('dialog', async (dialog) => { alerts.push(dialog.message()); await dialog.accept().catch(() => {}); });
    page.on('request', (r) => {
      if (r.method() === 'POST' && (r.postData() || '').includes('CmdEnqSave')) enquirySavePosts.push(r.url());
    });
    await loginToStage(page);
  });
  test.beforeEach(async () => {
    await openNewEnquiry(page);
    alerts.length = 0;
    enquirySavePosts.length = 0;
  });
  test.afterAll(async () => { await page?.context().close(); });

  // Save, then give the alert a moment to appear. Returns the alerts shown.
  async function saveAndCollectAlerts() {
    await page.locator('#ctl00_cpMain_CmdEnqSave').click();
    await expect.poll(() => alerts.length, { timeout: 10000 }).toBeGreaterThan(0);
    await page.waitForLoadState('load');
    return alerts.join(' | ');
  }

  test('TC07 - Enquiry with nothing filled in is refused', async () => {
    const shown = await saveAndCollectAlerts();
    expect(shown).toContain('Please select Source of enquiry');
    await expect(page, 'No enquiry was opened (no keyid)').not.toHaveURL(/keyid=/i);
    await expect(page.locator('#ctl00_cpMain_EnquiryName')).toHaveValue('');
  });

  test('TC08 - Enquiry with a 5-digit mobile number is refused', async () => {
    const phone = '98765';
    await fillEnquiryForm(page, { ...ENQUIRY, name: `Auto Neg ${runId}`, phone });
    const shown = await saveAndCollectAlerts();
    expect(shown).toMatch(/validations are not matching|mobile|phone/i);
    expect(enquirySavePosts, 'Save must not reach the server').toEqual([]);
    await expect(page).not.toHaveURL(/keyid=/i);
    await expect(page.locator('#ctl00_cpMain_PhoneM')).toHaveValue(phone);
  });

  test('TC09 - Mobile number field does not take letters', async () => {
    const mobile = page.locator('#ctl00_cpMain_PhoneM');
    await mobile.fill('abcdefghij');
    await mobile.press('Tab');
    await expect(mobile).toHaveValue('');
  });

  test('TC10 - Mobile number and Pin Code stop at their maximum length', async () => {
    const mobile = page.locator('#ctl00_cpMain_PhoneM');
    const pin = page.locator('#ctl00_cpMain_PinCode');
    await expect(mobile).toHaveAttribute('maxlength', '11');
    await expect(pin).toHaveAttribute('maxlength', '6');
    await mobile.pressSequentially('987654321012');
    await pin.pressSequentially('5600011');
    await expect(mobile).toHaveValue('98765432101');
    await expect(pin).toHaveValue('560001');
  });

});
