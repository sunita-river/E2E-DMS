import { Page, Locator, expect } from '@playwright/test';

export interface UserCredentials {
  dealercode?: string;
  username?: string;
  password?: string;
  [key: string]: any;
}

export class LoginPage {
  readonly page: Page;
  readonly dealerCodeInput: Locator;
  readonly branchDropdown: Locator;
  readonly usernameInput: Locator;
  readonly passwordInput: Locator;
  readonly loginButton: Locator;
  readonly riverUserLink: Locator;
  readonly riverUserUsernameInput: Locator;
  readonly riverUserPasswordInput: Locator;
  readonly riverUserSignInButton: Locator;

  // log: where step messages go (default console). PDI.spec.ts passes one that prefixes the dealer
  // code, since several dealers log at once there.
  constructor(page: Page, private readonly log: (message: string) => void = console.log) {
    this.page = page;
    this.dealerCodeInput = page.locator('input[name="DealerCode"]');
    this.branchDropdown = page.locator('select[name="BranchSlno"]');
    this.usernameInput = page.getByRole('textbox', { name: 'User Name' });
    this.passwordInput = page.getByRole('textbox', { name: 'Password' });
    this.loginButton = page.getByRole('button', { name: 'Sign In' });
    this.riverUserLink = page.getByRole('link', { name: 'River User' });

    // The dealer-code form stays in the DOM (hidden) while the River User form is shown, so
    // role/name lookups match both and hit strict mode. Use the River User form's own IDs.
    this.riverUserUsernameInput = page.locator('#UserName');
    this.riverUserPasswordInput = page.locator('#Password');
    this.riverUserSignInButton = page.locator('#CmdLogin');
  }

  // url: defaults to production; the stage specs pass https://stage.rivermobility.gaindms.com/
  async goto(url = 'https://rivermobility.gaindms.com/') {
    this.log('[STEP 1] Navigating to login page...');
    await this.page.goto(url, { waitUntil: 'domcontentloaded' });

    // The ASP.NET report viewer renders blank in Playwright's fixed default viewport, so the
    // context runs with viewport: null and the window is maximized here over CDP
    // (--start-maximized raced Chrome's window manager). Skipped where there is no window.
    try {
      const session = await this.page.context().newCDPSession(this.page);
      const { windowId } = await session.send('Browser.getWindowForTarget');
      // CDP won't go straight from minimized/fullscreen to maximized (seen on repeated
      // logins in one window), so set 'normal' first.
      await session.send('Browser.setWindowBounds', { windowId, bounds: { windowState: 'normal' } });
      await session.send('Browser.setWindowBounds', { windowId, bounds: { windowState: 'maximized' } });
    } catch {
      this.log('[STEP] Window maximize not available (headless?), continuing.');
    }
  }

  async login(creds: UserCredentials) {
    await this.dealerCodeInput.fill(creds.dealercode || '');

    // The branch lookup only fires on blur, and fill() doesn't blur; without the Tab the
    // User Name field stays disabled.
    await this.dealerCodeInput.press('Tab');

    // The branch lookup re-renders the form and wipes anything typed before it finishes,
    // so wait for the Branch dropdown to load.
    if (creds.dealercode) {
      await this.page
        .waitForFunction(
          () => {
            const el = document.querySelector('select[name="BranchSlno"]') as HTMLSelectElement | null;
            return !!el && el.options.length > 0 && !!el.value;
          },
          { timeout: 10000 }
        )
        .catch(() => {});
    }

    await this.usernameInput.fill(creds.username || '');
    await this.passwordInput.fill(creds.password || '');

    this.log('[STEP 2] Clicking Sign In button...');
    await this.loginButton.click();
    await this.page.waitForTimeout(2000);
  }

  // River User login has no Dealer Code / Branch step: just username and password.
  async loginAsRiverUser(creds: UserCredentials) {
    this.log('[STEP] Switching to River User login...');
    await this.riverUserLink.click();

    await this.riverUserUsernameInput.fill(creds.username || '');
    await this.riverUserUsernameInput.press('Tab');
    await this.riverUserPasswordInput.fill(creds.password || '');

    this.log('[STEP] Clicking Sign In button...');
    await this.riverUserSignInButton.click();
    await this.page.waitForTimeout(2000);
  }

  async verifyDashboard() {
    this.log('[STEP 3] Waiting for dashboard navigation...');

    // A wrong password or locked account stays on the login page; say so instead of a bare timeout.
    try {
      await this.page.locator('#ctl00_cpMain_DashboardCaption').waitFor({ state: 'visible', timeout: 30000 });
    } catch {
      const onLogin = await this.page.getByRole('button', { name: 'Sign In' }).isVisible().catch(() => false);
      // The DMS shows why under the Sign In button, e.g. "You do not have permission for the selected
      // branch" when the user has no access to that dealer, which no credentials.json change can fix.
      const dmsMessage = onLogin
        ? ((await this.page.getByText(/permission|invalid|incorrect|locked|not active|expired/i).first()
            .innerText({ timeout: 2000 }).catch(() => '')).replace(/\s+/g, ' ').trim())
        : '';
      throw new Error(onLogin
        ? `Login failed: still on the login page after Sign In (${dmsMessage ? `DMS says "${dmsMessage}"` : 'check the credentials in credentials.json'}).`
        : `Dashboard did not load within 30s after login (page: ${this.page.url()}).`);
    }

    await this.page.locator('.spinner').waitFor({ state: 'detached', timeout: 10000 }).catch(() => {});

    this.log('[STEP 3] Verifying Dashboard visibility...');

    const dashboardLink = this.page.getByRole('link', { name: /(Sales|Service) Dashboard/i });
    await expect(dashboardLink.first()).toBeVisible({ timeout: 10000 });
    this.log('[STEP 3] Dashboard verified.');
  }

 
  


}
