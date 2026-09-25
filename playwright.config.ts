import 'dotenv/config';
import { defineConfig, devices } from '@playwright/test';

// Browser channel for all projects. Defaults to real Google Chrome; set BROWSER_CHANNEL=msedge
// to run on Microsoft Edge instead (Chromium-based, so the launch flags below still apply).
// On a CI server without Chrome: BROWSER_CHANNEL=chromium after `npx playwright install chromium`.
const browserChannel = process.env.BROWSER_CHANNEL || 'chrome';

export default defineConfig({
  testDir: './tests',
    fullyParallel: false,
  timeout:500000,
  // Hard cap for the whole run so a CI job can never hang forever (the GSTR test alone can
  // legitimately take ~1h for many dealers). Override with PW_GLOBAL_TIMEOUT_MIN.
  globalTimeout: Number(process.env.PW_GLOBAL_TIMEOUT_MIN || 180) * 60 * 1000,
  // Fail CI if someone leaves a test.only in the code.
  forbidOnly: !!process.env.CI,

  //retries: process.env.CI ? 2 : 0,
 // workers: process.env.CI ? 1 : undefined,
 workers: 1,
 reporter: [
    ['line'],
    ['html', { outputFolder: 'playwright-report', open: 'never' }],
    ['junit', { outputFile: 'test-results/junit.xml' }], // Consumed by CI/CD test-result publishers
    ['allure-playwright', { outputFolder: 'allure-results' }], // Raw results land here
    ['./tests/reporters/gmail-reporter.ts'],
  ],

  use: {
    /* Base URL to use in actions like `await page.goto('')`. */
     baseURL: 'https://rivermobility.gaindms.com/',

    /* Collect trace when retrying the failed test. */
    trace: 'on-first-retry',
    screenshot:'only-on-failure',
    video:'retain-on-failure',
    /* Without these, a click/fill/waitForEvent that never completes waits until the whole test
       times out (Playwright's default is no limit). The DMS is slow, so they're generous. */
    actionTimeout: 30000,
    navigationTimeout: 90000,
    /* Slow down Playwright operations by N milliseconds (e.g., 500ms delay per action) */


    /* Optional: Set headless to false if you want to visually watch the browser window */
    // headless: false,

    /* Real Chrome (channel: 'chrome' below) applies Safe Browsing "download protection" even
       under CDP automation, which blocks report .xls downloads with a native "blocked" bar that
       Playwright can't dismiss (it's outside the page). These flags disable that check for the
       test browser only.

       --disable-popup-blocking: some report exports (e.g. "Export to Excel" in the report
       viewer tab) trigger their window.open() asynchronously after an AJAX/postback rather
       than synchronously inside the click, so Chrome's popup blocker doesn't recognize it as
       a user gesture and opens "about:blank#blocked" instead of the export, silently starving
       the test's waitForEvent('download'). */
    launchOptions: {
      args: [
        '--safebrowsing-disable-download-protection',
        '--disable-client-side-phishing-detection',
        '--disable-popup-blocking',
        // "Print Invoice" (VehicleSales.spec.ts) opens a popup that calls window.print()
        // automatically, which otherwise pops Chrome's native OS print-preview dialog —
        // a real system window page.close() can't dismiss via CDP, leaving it stuck open
        // across every iteration of that test's reference-number loop. This flag silently
        // bypasses that dialog instead of showing it.
        '--kiosk-printing',
      ],
    },
  },

  /* Configure projects for Google Chrome only */
  projects: [
    {
      name: 'setup',
      testMatch: /.*\.setup\.ts/,
      use: {
        ...devices['Desktop Chrome'],
        channel: browserChannel,
        viewport: null,
        deviceScaleFactor: undefined,
      },
    },
    {
      // Manual, human-driven passkey enrollment (see tests/enrollPasskey.manual.ts).
      // Deliberately not a dependency of any other project — it pauses for
      // interactive input and must only run when explicitly requested via
      // `--project=enroll-passkey`, never as part of a normal/CI test run.
      name: 'enroll-passkey',
      testMatch: /enrollPasskey\.manual\.ts/,
      use: {
        ...devices['Desktop Chrome'],
        channel: browserChannel,
        viewport: null,
        deviceScaleFactor: undefined,
      },
    },
    {
      name: 'Google Chrome',
      dependencies: ['setup'],
      use: {
        ...devices['Desktop Chrome'],
        channel: browserChannel, // Installed Google Chrome (or Edge via BROWSER_CHANNEL) instead of bundled Chromium
        storageState: 'playwright/.auth/salesforce.json',
        viewport: null,
        deviceScaleFactor: undefined,
      },
    },
  ],

});