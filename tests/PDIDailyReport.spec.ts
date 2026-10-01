import { test } from '@playwright/test';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { sendPdiReport, problemsOf } from './utils/pdiReport';
import { buildPdiDaily, localDate, emailedMarker, OUTPUT_DIR, RUNS_DIR } from './utils/pdiDaily';

// Daily PDI email: the one email of the day for PDI. It rebuilds Output/PDI_Daily_<day>.xlsx from every
// PDI run saved that day (see tests/utils/pdiDaily.ts for the merging rules), refreshes PDI_Master.xlsx
// (Tracking sheet + last 3 days) and the all-time history, and emails the day's Excel and the master to
// GMAIL_TO with PDI_CC. The HTML dashboards are saved in Output/, not emailed. No browser, no DMS login.
// The PDI runs themselves send no email; run this once after the day's last PDI run.
//
// It sends only once a day: a second run rebuilds the files but skips the email (it says so).
//
// Settings (optional):
//   PDI_REPORT_DATE  yyyy-mm-dd to report on (default: today, by the files' saved time)
//   PDI_SEND_EMAIL   0 = build the report files but don't email them
//   PDI_FORCE_EMAIL  1 = send even though today's email has already gone out
//   PDI_EMAIL_REVIEW 1 = send a [REVIEW] copy to GMAIL_TO only (no CC); it doesn't count as the day's email

// Sends an email, so only when asked for: npm run report:pdi-daily sets PDI_RUN=1.
test.skip(process.env.PDI_RUN !== '1', 'Daily PDI report runs only via npm run report:pdi-daily (sets PDI_RUN=1).');

test('Should email the combined PDI report for the day', async () => {
  const day = process.env.PDI_REPORT_DATE || localDate(new Date());
  const daily = await buildPdiDaily(day);
  if (!daily) throw new Error(`No PDI result files from ${day} in ${RUNS_DIR} or ${OUTPUT_DIR}.`);

  const marker = emailedMarker(day);
  if (fs.existsSync(marker) && process.env.PDI_FORCE_EMAIL !== '1') {
    console.log(`[STEP] Today's PDI email already went out (${fs.readFileSync(marker, 'utf8').trim()}) — not sending it again. Set PDI_FORCE_EMAIL=1 to resend.`);
  } else if (await sendPdiReport(daily.results, daily.failures, daily.meta, daily.file) && process.env.PDI_EMAIL_REVIEW !== '1') {
    // A review copy (PDI_EMAIL_REVIEW=1, to GMAIL_TO only) doesn't count as the day's email.
    fs.mkdirSync(path.dirname(marker), { recursive: true });
    fs.writeFileSync(marker, `sent ${new Date().toLocaleString('en-IN')}\n`);
  }

  const problems = problemsOf(daily.results, daily.failures, false);
  if (problems.length) throw new Error(`${problems.length} problem(s) today (listed in the report):\n  ${problems.slice(0, 20).join('\n  ')}`);
});
