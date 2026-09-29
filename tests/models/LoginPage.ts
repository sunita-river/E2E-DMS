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

    // The dealer-code login form stays in the DOM (just hidden) when the River User form is
    // shown, so a role/name lookup like the dealer flow's usernameInput can match both and
    // trip Playwright's strict mode. Scope these to the River User form's own element IDs.
    this.riverUserUsernameInput = page.locator('#UserName');
    this.riverUserPasswordInput = page.locator('#Password');
    this.riverUserSignInButton = page.locator('#CmdLogin');
  }

  async goto() {
    this.log('[STEP 1] Navigating to login page...');
    await this.page.goto('https://rivermobility.gaindms.com/', { waitUntil: 'domcontentloaded' });

    // The app's legacy ASP.NET report viewer renders blank under Playwright's fixed default
    // viewport (fine in a manually-opened, maximized browser), so the context runs with
    // viewport: null and this maximizes the main window explicitly via CDP instead of the
    // riskier --start-maximized launch flag, which raced Chrome's window manager on startup.
    // Skipped quietly where there's no real window to maximize (headless CI, non-Chromium).
    try {
      const session = await this.page.context().newCDPSession(this.page);
      const { windowId } = await session.send('Browser.getWindowForTarget');
      // Chrome's CDP refuses to go straight from a minimized/fullscreen state to
      // maximized (hit when re-navigating an already-open window across repeated
      // logins, e.g. looping over dealers) — reset to 'normal' first, a no-op if
      // it's already normal, then maximize.
      await session.send('Browser.setWindowBounds', { windowId, bounds: { windowState: 'normal' } });
      await session.send('Browser.setWindowBounds', { windowId, bounds: { windowState: 'maximized' } });
    } catch {
      this.log('[STEP] Window maximize not available (headless?) — continuing.');
    }
  }

  async login(creds: UserCredentials) {
    

    // Safely clear and fill inputs (handles missing/undefined credentials gracefully)
    await this.dealerCodeInput.fill(creds.dealercode || '');

    // fill() sets the value but never blurs the field, and the app only fires its
    // branch lookup on blur/change — so without an explicit Tab, the User Name
    // field stays disabled and every following action hangs waiting for it.
    await this.dealerCodeInput.press('Tab');

    // Entering the dealer code triggers an async lookup that populates the Branch
    // dropdown; filling username/password before it resolves gets them wiped out
    // by the re-render, so wait for the branch to actually load before continuing.
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

    // Allow server response processing
    await this.page.waitForTimeout(2000);
  }

  // "River User" login skips the Dealer Code / Branch lookup entirely — it's a plain
  // username + password sign-in, used for the salesforceValidUser-style credential.
  async loginAsRiverUser(creds: UserCredentials) {
    this.log('[STEP] Switching to River User login...');
    await this.riverUserLink.click();

    await this.riverUserUsernameInput.fill(creds.username || '');
    await this.riverUserUsernameInput.press('Tab');
    await this.riverUserPasswordInput.fill(creds.password || '');

    this.log('[STEP] Clicking Sign In button...');
    await this.riverUserSignInButton.click();

    // Allow server response processing
    await this.page.waitForTimeout(2000);
  }

  async verifyDashboard() {
    this.log('[STEP 3] Waiting for dashboard navigation...');

    // Wait for main dashboard container or caption
    // A wrong password / locked account stays on the login page — say so instead of a bare timeout.
    try {
      await this.page.locator('#ctl00_cpMain_DashboardCaption').waitFor({ state: 'visible', timeout: 30000 });
    } catch {
      const onLogin = await this.page.getByRole('button', { name: 'Sign In' }).isVisible().catch(() => false);
      throw new Error(onLogin
        ? 'Login failed: still on the login page after Sign In (check the credentials in credentials.json).'
        : `Dashboard did not load within 30s after login (page: ${this.page.url()}).`);
    }

    // Safely handle spinner if active
    await this.page.locator('.spinner').waitFor({ state: 'detached', timeout: 10000 }).catch(() => {});

    this.log('[STEP 3] Verifying Dashboard visibility...');

    // Unified Regex to catch either Sales or Service Dashboard links safely
    const dashboardLink = this.page.getByRole('link', { name: /(Sales|Service) Dashboard/i });
    await expect(dashboardLink.first()).toBeVisible({ timeout: 10000 });
    this.log('✅ Dashboard verified successfully!');
  }

 
  


}
