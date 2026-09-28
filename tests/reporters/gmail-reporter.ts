import type { Reporter, TestCase, TestResult, FullResult } from '@playwright/test/reporter';
import fs from 'fs';
import os from 'os';
import nodemailer from 'nodemailer';

type CollectedCsvReport = { filename: string; content: string };

// Splits a single CSV row into fields, honoring double-quoted fields that may
// contain commas, embedded quotes ("" escapes), or newlines — matching the
// escaping toCsv() in enquiryListReport.spec.ts produces.
function parseCsv(csv: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < csv.length; i++) {
    const char = csv[i];
    if (inQuotes) {
      if (char === '"') {
        if (csv[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
    } else if (char === '"') {
      inQuotes = true;
    } else if (char === ',') {
      row.push(field);
      field = '';
    } else if (char === '\r') {
      // skip — the trailing \n in the \r\n pair ends the row
    } else if (char === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Renders a CSV report as an HTML table so it reads directly in the email
// body instead of requiring the recipient to open a separate attachment.
function csvToHtmlTable(title: string, csv: string): string {
  const rows = parseCsv(csv.trim());
  if (rows.length === 0) return '';
  const headerRow = rows[0]!;
  const bodyRows = rows.slice(1);

  const thead = `<tr>${headerRow
    .map((h) => `<th style="border:1px solid #ccc;padding:4px 8px;background:#f2f2f2;text-align:left;">${escapeHtml(h)}</th>`)
    .join('')}</tr>`;
  const tbody = bodyRows
    .map((r) => `<tr>${r.map((c) => `<td style="border:1px solid #ccc;padding:4px 8px;">${escapeHtml(c)}</td>`).join('')}</tr>`)
    .join('');

  return `
    <h3 style="font-family:sans-serif;">${escapeHtml(title)}</h3>
    <table style="border-collapse:collapse;font-family:sans-serif;font-size:13px;">${thead}${tbody}</table>
  `;
}

// Emails the run summary and any CSV reports (e.g. combined-summary.csv from
// enquiryListReport.spec.ts) to a Gmail address once the whole run finishes,
// with the CSV pasted directly into the email body as a table rather than
// sent as a file attachment. Configure via env vars — nothing is hardcoded here:
//   GMAIL_USER          the Gmail account to send from (also used as SMTP login)
//   GMAIL_APP_PASSWORD  a Gmail App Password for that account (not its login password —
//                       generate one at https://myaccount.google.com/apppasswords)
//   GMAIL_TO            recipient address (defaults to GMAIL_USER, i.e. mail yourself)
// If any of these aren't set, the reporter logs why it's skipping instead of failing the run.
export default class GmailReporter implements Reporter {
  private csvReports: CollectedCsvReport[] = [];
  private passed = 0;
  private failed = 0;
  private skipped = 0;

  onTestEnd(test: TestCase, result: TestResult) {
    if (result.status === 'passed') this.passed++;
    else if (result.status === 'skipped') this.skipped++;
    else this.failed++;

    for (const attachment of result.attachments) {
      if (!attachment.name.endsWith('.csv')) continue;
      const content = attachment.body ?? (attachment.path ? fs.readFileSync(attachment.path) : undefined);
      if (!content) continue;
      this.csvReports.push({ filename: attachment.name, content: content.toString('utf-8') });
    }
  }

  async onEnd(result: FullResult) {
    if (this.passed + this.failed + this.skipped === 0) {
      console.log('[gmail-reporter] Skipping email — no tests ran (e.g. --list).');
      return;
    }

    const user = process.env.GMAIL_USER;
    const appPassword = process.env.GMAIL_APP_PASSWORD;
    const to = process.env.GMAIL_TO || user;

    if (!user || !appPassword || !to) {
      console.log(
        '[gmail-reporter] Skipping email — set GMAIL_USER and GMAIL_APP_PASSWORD env vars (and optionally GMAIL_TO) to enable it.'
      );
      return;
    }

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user, pass: appPassword },
    });

    const runDate = new Date().toISOString().slice(0, 10);
    const summary = `Passed: ${this.passed}, Failed: ${this.failed}, Skipped: ${this.skipped}`;
    // Says where the run came from, so GitHub and local-PC emails can be told apart in the inbox.
    const runUrl = process.env.GITHUB_ACTIONS
      ? `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`
      : '';
    const sourceTag = runUrl ? 'GitHub' : 'Local';
    const sourceText = runUrl ? `Run from: GitHub Actions (${runUrl})` : `Run from: local PC (${os.hostname()})`;
    const sourceHtml = runUrl
      ? `Run from: <b>GitHub Actions</b> - <a href="${runUrl}">open this run</a>`
      : `Run from: <b>local PC</b> (${escapeHtml(os.hostname())})`;
    const tablesHtml = this.csvReports.map((r) => csvToHtmlTable(r.filename, r.content)).join('<br/>');

    try {
      await transporter.sendMail({
        from: user,
        to,
        // Optional comma-separated CC list from GMAIL_CC (kept out of the code/Git).
        ...(process.env.GMAIL_CC ? { cc: process.env.GMAIL_CC } : {}),
        subject: `DMS Automation Report [${sourceTag}] - ${result.status} - ${runDate}`,
        text: `Test run finished with status: ${result.status}\n${summary}\n${sourceText}`,
        html: `
          <p style="font-family:sans-serif;">Test run finished with status: <b>${escapeHtml(result.status)}</b></p>
          <p style="font-family:sans-serif;">${escapeHtml(summary)}</p>
          <p style="font-family:sans-serif;">${sourceHtml}</p>
          ${tablesHtml}
        `,
      });
      console.log(`[gmail-reporter] Report emailed to ${to} (${this.csvReports.length} report(s) pasted in body, no attachments).`);
    } catch (err) {
      console.error('[gmail-reporter] Failed to send email:', err);
    }
  }
}
