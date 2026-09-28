/*
 * MIS Report Checker for River DMS - run from VS Code
 * ---------------------------------------------------
 * Setup (one time, in the VS Code terminal, inside this folder):
 *     npm install
 * Run:
 *     npm start
 *
 * A browser window opens and logs in automatically with the "dealervalidUser"
 * account from resources/credentials.json (same login as the Playwright tests).
 * Your session is kept in the "dms-profile" folder, so next time it goes straight in.
 * Uses Google Chrome by default; set BROWSER_CHANNEL=msedge to use Microsoft Edge.
 *
 * Output ("MIS Reports" folder):
 *     MIS_Report_Objects_<date>.xlsx  - formatted sheet (report name, definition, objects) + "Changes vs Previous" sheet
 *     MIS_Report_Check_<date>.csv     - every report with its full path and result
 *     MIS_Report_Summary_<date>.txt   - totals, list of failures, and changes since the previous run
 *     MIS_Report_Snapshot_<date>.json - used to compare the next run with this one
 *
 * It only opens/views reports. It does not save, edit or delete anything in the DMS.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
// CREDENTIALS_FILE / MIS_PROFILE_DIR: optional overrides (e.g. CI secrets file, separate browser profile)
const credentials = JSON.parse(fs.readFileSync(process.env.CREDENTIALS_FILE || path.join(__dirname, 'resources', 'credentials.json'), 'utf8'));
const PROFILE_DIR = process.env.MIS_PROFILE_DIR || path.join(__dirname, 'dms-profile');
require('dotenv').config({ path: path.join(__dirname, '.env') }); // GMAIL_USER / GMAIL_APP_PASSWORD / GMAIL_TO

// ======================= SETTINGS =======================
const CONFIG = {
  baseUrl: 'https://rivermobility.gaindms.com',
  startPage: '/Accounts/AccountsReports.aspx', // any page that shows the MIS menu
  fromDate: null,        // 'dd/mm/yyyy'  - null = 1st of the month (of yesterday)
  toDate: null,          // 'dd/mm/yyyy'  - null = yesterday (today - 1)
  concurrency: 4,        // how many MIS pages to test at the same time
  reportTimeoutSec: 30,  // a report slower than this is marked TIMEOUT
  // No visible browser on servers/CI (CI=true or HEADLESS=1); visible window when run by hand.
  headless: !!(process.env.CI || process.env.HEADLESS === '1'),
  maxRunMinutes: Number(process.env.MIS_MAX_MINUTES || 60), // whole run is stopped (exit code 2) after this
  maxPageMinutes: 25,    // one MIS page (all its reports) is given up after this
  // e.g. ['Workshop Reports', 'CRM Reports'] ; empty = all MIS pages. Env: MIS_PAGES="CRM Reports,Workshop Reports"
  onlyPages: (process.env.MIS_PAGES || '').split(',').map(s => s.trim()).filter(Boolean),
  sendEmail: process.env.MIS_SEND_EMAIL !== '0', // email the summary (Excel attached) to GMAIL_TO; MIS_SEND_EMAIL=0 to skip
  // Values for reports that need more than dates (filled only when the report shows that field)
  extraInputs: {
    chassisNo: 'P7BAAAXX1KS016088',  // Vehicle History
    sparePartCode: 'R106020383',     // Movement Analysis (only filled when the DMS asks for a spare part code)
    // Dead Stock Analysis (Spares, MFR Spares, Accessories) - any sensible values work
    purchaseDateDaysAgo: 1,          // Purchase Date = today - 1
    workshopConsideredDays: '6',
    counterSaleConsideredDays: '9',
    // Month + Year fields (Leakage And Redemption Summary, Spares FSN Analysis) use the current month and year
  },
};
// ========================================================

const OUT_DIR = path.join(__dirname, 'MIS Reports'); // every output file is saved here

const pad = n => String(n).padStart(2, '0');
const fmt = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const now = new Date();
const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
// Default range: 1st of yesterday's month up to yesterday (on the 1st this is the whole previous month)
const FROM = CONFIG.fromDate || fmt(new Date(yesterday.getFullYear(), yesterday.getMonth(), 1));
const TO = CONFIG.toDate || fmt(yesterday);
const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;

// Waits for Enter - only possible when a person is at the terminal. On a server/CI (no TTY)
// it fails straight away instead of waiting forever for input that will never come.
function ask(q) {
  if (!process.stdin.isTTY || CONFIG.headless) {
    return Promise.reject(new Error('Manual login needed but no one is at the terminal (CI / headless run).'));
  }
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(res => rl.question(q, a => { rl.close(); res(a); }));
}

// Rejects if the promise hasn't settled in time.
function withTimeout(promise, ms, what) {
  let t;
  return Promise.race([promise, new Promise((_, rej) => { t = setTimeout(() => rej(new Error(`${what} took longer than ${Math.round(ms / 60000)} min`)), ms); })])
    .finally(() => clearTimeout(t));
}

// Last line of defence: nothing may keep the process alive forever on a server.
process.on('unhandledRejection', e => { console.error('Unexpected error:', e && e.message || e); process.exitCode = 1; });
setTimeout(() => {
  console.error(`\nMIS check stopped: still running after ${CONFIG.maxRunMinutes} min (MIS_MAX_MINUTES).`);
  process.exit(2);
}, CONFIG.maxRunMinutes * 60000).unref();

// ---------- This function runs INSIDE the DMS page (browser side) ----------
async function checkPageInBrowser(cfg) {
  const ERR_RE = /(Server Error|Runtime Error|Exception|error occurred|Invalid (object|column)|Timeout expired|Could not find|Incorrect syntax|Conversion failed|Procedure or function|Something went wrong|Object reference|Gateway Time-out|Bad Gateway|Service Unavailable)/i;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const scripts = t => [...t.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1].trim());
  const fetchT = async (url, opts = {}, timeoutMs = cfg.timeoutMs) => {
    const c = new AbortController(); const t = setTimeout(() => c.abort(), timeoutMs);
    try { return await fetch(url, { credentials: 'include', ...opts, signal: c.signal }); }
    finally { clearTimeout(t); }
  };
  const log = m => { try { window.reportProgress && window.reportProgress(m); } catch (e) { } };

  // The page's own highlight function crashes when reports are clicked by a script; make it safe.
  if (window.HighlightSelected) {
    const o = window.HighlightSelected;
    window.HighlightSelected = function () { try { return o.apply(this, arguments); } catch (e) { } };
  }
  // Some reports open their own page in a new tab straight from SetTab() (e.g. Appointment Retention
  // Metrics -> AppointmentRetentionMetrics.aspx). Catch that URL instead of letting stray tabs open;
  // the Node side then opens it in a real tab and reads the table headings.
  let openedUrl = null;
  window.open = function (u) {
    if (u) openedUrl = new URL(u, location.href).href;
    return { moveTo() { }, focus() { }, close() { }, document: { write() { }, close() { } } };
  };
  const f = document.forms[0];
  const base = new Set(scripts(await (await fetchT(location.href, {}, 90000)).text()));
  const seen = new Set();
  const anchors = [...document.querySelectorAll('a[onclick*="SetTab"]')].filter(a => {
    const id = (a.getAttribute('onclick').match(/SetTab\(this\.id,(\d+)/) || [])[1];
    if (!id || seen.has(id)) return false; seen.add(id); return true;
  });

  const post = async (timeoutMs) => {
    const fd = new FormData(f);
    fd.set('__EVENTTARGET', 'ctl00$cpMain$btnViewReport'); fd.set('__EVENTARGUMENT', '');
    const r = await fetchT(new URL(f.getAttribute('action'), location.href), { method: 'POST', body: new URLSearchParams(fd) }, timeoutMs);
    const t = await r.text(); const d = new DOMParser().parseFromString(t, 'text/html');
    const msg = ((d.getElementById('ctl00_cpMain_lblMsg')?.textContent || '') + ' ' +
                 (d.getElementById('ctl00_cpMain_LblMsg')?.textContent || '')).trim();
    return { r, t, d, msg };
  };

  // Extra mandatory fields some reports ask for (only touched when visible for the selected report)
  const fillExtraInputs = (vis, msgs) => {
    const x = cfg.extra;
    const changed = []; // [element, old value, old checked] so the caller can undo
    const keep = e => changed.push([e, e.value, e.checked]);
    const set = (idRe, v) => vis.filter(e => idRe.test(e.id) && !e.value.trim()).forEach(e => { keep(e); e.value = v; });
    const pick = (idRe, v) => vis.filter(e => e.tagName === 'SELECT' && idRe.test(e.id)).forEach(e => {
      const o = [...e.options].find(o => o.value === v || Number(o.value) === Number(v)); if (o) { keep(e); e.value = o.value; }
    });
    // Chassis / spare part only when the DMS message asks for them: on other reports (e.g. Dead
    // Stock's optional Item Code) they would narrow the report to a single vehicle or part.
    const asked = (msgs || []).join(' ');
    if (/chassis|registr/i.test(asked)) set(/_ChassisNo$/, x.chassisNo);
    if (/spare ?part|item ?code/i.test(asked)) set(/_SparePartCode$/, x.sparePartCode);
    // Field ids on the MFR Spares page end in "Mfr" (e.g. WorkshopConsideredDaysMfr).
    set(/_PurchaseDate(Mfr)?$/, cfg.purchaseDate);
    set(/_WorkshopConsideredDays(Mfr)?$/, x.workshopConsideredDays);
    set(/_CounterSaleConsideredDays(Mfr)?$/, x.counterSaleConsideredDays);
    pick(/_(Month|FromMonthNo|ToMonthNo)$/, cfg.month);
    set(/_(Year|FromYearNo|ToYearNo)$/, cfg.year);
    // "Group by*": tick the first option so the page fills its group-by box
    for (const box of vis.filter(e => e.tagName === 'TEXTAREA' && /group|Textpurchase/i.test(e.id) && !e.value.trim())) {
      const first = vis.find(e => e.type === 'checkbox' && /CheckBoxList\d_0$/.test(e.id));
      keep(box);
      if (first && !first.checked) { keep(first); first.click(); }
      if (!box.value.trim() && first) box.value = first.value.replace(/&nbsp\s*/g, '').trim();
    }
    return changed;
  };

  // Objects = the viewer's "Customize Columns" list; if a report has none, the grid's header row.
  const readObjects = d2 => {
    const fromList = [...d2.querySelectorAll('#CheckBoxList1 label')].map(l => l.textContent.trim()).filter(Boolean);
    if (fromList.length) return fromList;
    const head = d2.getElementById('GrdMain')?.rows[0];
    return head ? [...head.cells].map(c => c.textContent.trim().replace(/\s+/g, ' ')).filter(Boolean) : [];
  };
  const viewerUrl = t => {
    const m = scripts(t).filter(s => !base.has(s)).join('\n').match(/window\.open\s*\(\s*'([^']+)'/);
    return m && new URL(m[1], location.href);
  };
  // Crystal reports (Report.aspx) are drawn by crv.js; if the server can't serve it the page stays white.
  const crystalScriptStatus = {};
  const crystalBroken = async (tt, u) => {
    const src = (tt.match(/<script[^>]+src="([^"]*crv\.js[^"]*)"/i) || [])[1];
    if (!src) return false;
    const abs = new URL(src.replace(/&amp;/g, '&'), u).href;
    if (!(abs in crystalScriptStatus)) {
      try { crystalScriptStatus[abs] = (await fetchT(abs, {}, 30000)).status; } catch (e) { crystalScriptStatus[abs] = 0; }
    }
    return crystalScriptStatus[abs] !== 200 && `Crystal viewer script crv.js not served (HTTP ${crystalScriptStatus[abs]}) - page shows blank`;
  };

  const results = [];
  let k = 0;
  for (const a of anchors) {
    k++;
    const oc = a.getAttribute('onclick');
    const rid = oc.match(/SetTab\(this\.id,(\d+)/)[1];
    const name = (oc.match(/SetTab\(this\.id,\d+,'([^']*)'/) || [, a.textContent.trim()])[1].trim();
    let section = '';
    for (const c of document.querySelectorAll(`a[onclick*="SetTab(this.id,${rid},"]`)) {
      const p = c.closest('.panel,.card,.accordion-item,.box');
      const h = p && p.querySelector('.panel-heading,.card-header,.accordion-header,.panel-title,.box-header');
      if (h && h.textContent.trim()) { section = h.textContent.trim().replace(/\s+/g, ' '); if (!/^All Reports$/i.test(section)) break; }
    }
    // Report definition = the italic description under the report name in the list
    const definition = (a.querySelector('h6')?.textContent || '').trim().replace(/\s+/g, ' ');
    const rec = { section, report: name, definition, objects: [], result: '', detail: '', rows: '', seconds: '' };
    const t0 = Date.now();
    try {
      // 1. select the report (same as clicking it)
      const rs = document.getElementById('ctl00_cpMain_ReportSlno');
      let ok = false;
      openedUrl = null;
      for (const c of [...document.querySelectorAll(`a[onclick*="SetTab(this.id,${rid},"]`)].reverse()) {
        const tv = document.getElementById('ctl00_cpMain_tabval'); if (tv) tv.value = '';
        try { c.click(); } catch (e) { }
        await sleep(300);
        if (!rs || rs.value == rid) { ok = true; break; }
      }
      if (!ok) { rec.result = 'COULD NOT SELECT'; rec.detail = 'Report link click failed'; }
      else if (openedUrl) { rec.result = 'NEW TAB'; rec.openUrl = openedUrl; }
      else {
        // 2. fill dates
        const vis = [...f.querySelectorAll('input,select,textarea')]
          .filter(e => e.offsetParent && e.type !== 'hidden' && !/ddlSearchOption|txtSearch/i.test(e.id));
        vis.filter(e => e.tagName === 'INPUT' && (e.type === 'text' || e.type === '') && /Dat/i.test(e.id + e.name))
          .forEach(e => { e.value = /PurchaseDate/i.test(e.id) ? cfg.purchaseDate : /To|End/i.test(e.id) ? cfg.to : cfg.from; });
        // 3. View Report. If the DMS asks for more input: first fill the extra values from CONFIG
        //    (chassis no, spare part, month/year...), then also pick the first branch/franchise.
        const readReply = ({ r, t, d, msg }) => {
          const nw = scripts(t).filter(s => !base.has(s)).join('\n');
          const msgs = [...nw.matchAll(/(?:alert|Messagebox|MessageBox|ShowMessage)\s*\(\s*['"]([^'"]{0,200})/g)].map(x => x[1]);
          if (msg) msgs.push(msg);
          return { r, t, d, nw, msgs, m: nw.match(/window\.open\s*\(\s*'([^']+)'/) };
        };
        let reply = readReply(await post());
        const restore = [];
        for (let i = 0; i < 2 && !reply.m && reply.msgs.some(s => /select|enter/i.test(s)); i++) {
          if (i === 0) restore.push(...fillExtraInputs(vis, reply.msgs));
          else vis.filter(e => e.tagName === 'SELECT' && (e.value === '' || e.value === '0') && e.options.length > 1)
            .forEach(e => { restore.push([e, e.value, e.checked]); e.selectedIndex = 1; });
          reply = readReply(await post());
        }
        if (restore.length) rec.detail = 'Used extra inputs';
        // put the shared form back so the next report doesn't inherit these values
        restore.reverse().forEach(([e, v, c]) => { e.value = v; if (e.type === 'checkbox') e.checked = c; });
        const { r, d, m, msgs } = reply;

        if (r.status !== 200 || ERR_RE.test(d.title + ' ' + (d.body?.innerText || '').slice(0, 1500))) {
          rec.result = 'SERVER ERROR'; rec.detail = (d.title || `HTTP ${r.status}`) + ' (on View Report)';
        } else if (!m) {
          rec.result = msgs.length ? 'NEEDS EXTRA INPUT' : 'NO REPORT OPENED'; rec.detail = msgs.join(' | ');
        } else {
          // 4. open the report viewer and check it
          const u = new URL(m[1], location.href);
          const r2 = await fetchT(u); const tt = await r2.text();
          const d2 = new DOMParser().parseFromString(tt, 'text/html');
          const body = (d2.body?.innerText || '').replace(/\s+/g, ' ');
          const i = body.search(ERR_RE);
          const grid = d2.getElementById('GrdMain');
          rec.rows = grid ? Math.max(0, grid.rows.length - 1) : 0;
          rec.objects = readObjects(d2);
          // Reports over the 1000-row display limit show no grid, just a banner + Excel download
          const overLimit = /more than the \d+-record display limit/i.test(d2.getElementById('ErroeMsg')?.textContent || '');
          const crystal = u.pathname.includes('Report.aspx') && await crystalBroken(tt, u);
          if (r2.status !== 200 || i >= 0) {
            rec.result = 'SERVER ERROR';
            rec.detail = ((d2.title || `HTTP ${r2.status}`) + ' ' + (i >= 0 ? body.slice(i, i + 120) : '')).trim();
          } else if (crystal) {
            rec.result = 'SERVER ERROR'; rec.detail = crystal;
          } else if (overLimit) {
            rec.result = 'OK'; rec.rows = '1000+'; rec.detail = 'Over display limit (downloads as Excel)';
          } else if (rec.rows > 0) rec.result = 'OK';
          else {
            rec.result = 'BLANK / NO DATA';
            rec.detail = u.pathname.includes('Report.aspx') ? 'Crystal viewer shows blank page' : 'Viewer window empty';
          }
        }
      }
    } catch (e) {
      if (e.name === 'AbortError') {
        rec.result = 'TIMEOUT'; rec.detail = `Slower than ${cfg.timeoutMs / 1000}s`;
        // Retry with a one-day range and a longer wait, only to read the report's columns.
        try {
          [...f.querySelectorAll('input')].filter(e => e.offsetParent && (e.type === 'text' || e.type === '') && /Dat/i.test(e.id + e.name))
            .forEach(e => { e.value = cfg.today; });
          const u = viewerUrl((await post(cfg.retryTimeoutMs)).t);
          if (u) {
            const d2 = new DOMParser().parseFromString(await (await fetchT(u, {}, cfg.retryTimeoutMs)).text(), 'text/html');
            rec.objects = readObjects(d2);
            if (rec.objects.length) rec.detail += ' (columns read from a 1-day run)';
          }
        } catch (e2) { }
      }
      else { rec.result = 'SERVER ERROR'; rec.detail = String(e).slice(0, 150); }
    }
    rec.seconds = Math.round((Date.now() - t0) / 1000);
    results.push(rec);
    log(`${k}/${anchors.length} ${name} -> ${rec.result}`);
  }
  return results;
}
// ----------------------------------------------------------------------------

// ---------- Compare with the previous run ----------
// Newest earlier MIS_Report_Snapshot_*.json in the folder; falls back to the newest CSV that has the
// report definitions/objects columns (older runs, before snapshots existed).
function loadPreviousRun(dir) {
  if (!fs.existsSync(dir)) return null;
  const files = fs.readdirSync(dir).sort().reverse(); // names end in a sortable timestamp
  const snap = files.find(f => /^MIS_Report_Snapshot_.*\.json$/.test(f));
  if (snap) return { name: snap, rows: JSON.parse(fs.readFileSync(path.join(dir, snap), 'utf8')) };
  for (const f of files.filter(f => /^MIS_Report_Check_.*\.csv$/.test(f))) {
    const [head, ...data] = parseCsv(fs.readFileSync(path.join(dir, f), 'utf8').replace(/^﻿/, ''));
    if (!head || !head.includes('objectsText')) continue;
    const col = n => head.indexOf(n);
    return {
      name: f,
      rows: data.filter(r => r.length === head.length).map(r => ({
        page: r[col('page')], section: r[col('section')], report: r[col('report')], definition: r[col('definition')],
        objects: r[col('objectsText')] ? r[col('objectsText')].split(', ') : [], result: r[col('result')],
      })),
    };
  }
  return null;
}

function parseCsv(text) {
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n') { row.push(cell.replace(/\r$/, '')); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

// A report is identified by MIS page + report name. Column changes are only reported when both runs
// managed to read the columns (a timeout in one run is not a real column change).
function compareRuns(prevRows, curRows) {
  // Only compare MIS pages checked in both runs, so a partial run (MIS_PAGES / onlyPages) or a page
  // that failed to load doesn't show up as hundreds of added/removed reports.
  const pagesIn = rows => new Set(rows.filter(r => !/^\((page|menu link)\)$/.test(r.report)).map(r => r.page));
  const prevPages = pagesIn(prevRows), curPages = pagesIn(curRows);
  prevRows = prevRows.filter(r => curPages.has(r.page));
  curRows = curRows.filter(r => prevPages.has(r.page));
  const key = r => `${r.page}|${String(r.report).trim().toLowerCase()}`;
  const prev = new Map(prevRows.map(r => [key(r), r]));
  const cur = new Map(curRows.map(r => [key(r), r]));
  const added = curRows.filter(r => !prev.has(key(r)));
  const removed = prevRows.filter(r => !cur.has(key(r)));
  const moved = [], columns = [];
  for (const r of curRows) {
    const p = prev.get(key(r));
    if (!p) continue;
    if ((p.section || '') !== (r.section || '')) moved.push({ page: r.page, report: r.report, from: p.section, to: r.section });
    const po = p.objects || [], co = r.objects || [];
    if (po.length && co.length) {
      const addedCols = co.filter(o => !po.includes(o)), removedCols = po.filter(o => !co.includes(o));
      if (addedCols.length || removedCols.length) columns.push({ page: r.page, section: r.section, report: r.report, addedCols, removedCols });
    }
  }
  return { added, removed, moved, columns };
}

// ---------- Email summary ----------
// HTML email (inline styles + tables so it renders the same in Gmail/Outlook), Excel report attached.
async function sendSummaryEmail({ rows, changes, previousName, from, to, minutes, attachments }) {
  const user = process.env.GMAIL_USER, pass = process.env.GMAIL_APP_PASSWORD, rcpt = process.env.GMAIL_TO || user;
  if (!user || !pass) { console.log('Email skipped: set GMAIL_USER and GMAIL_APP_PASSWORD in .env'); return; }
  const nodemailer = require('nodemailer');
  const h = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const count = k => rows.filter(r => r.result === k).length;
  const total = rows.length, ok = count('OK');
  const failed = count('SERVER ERROR') + count('TIMEOUT') + count('NO REPORT OPENED') + count('COULD NOT SELECT');
  const pct = total ? Math.round(ok * 100 / total) : 0;
  const FONT = "font-family:'Segoe UI',Arial,sans-serif;";

  const card = (label, value, color) => `
    <td style="padding:6px;" width="25%"><table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;border-top:4px solid ${color};">
      <tr><td style="${FONT}padding:14px 12px 4px;font-size:26px;font-weight:700;color:${color};text-align:center;">${value}</td></tr>
      <tr><td style="${FONT}padding:0 12px 14px;font-size:12px;color:#5f6368;text-align:center;text-transform:uppercase;letter-spacing:.5px;">${label}</td></tr>
    </table></td>`;

  const table = (title, color, list, cols) => !list.length ? '' : `
    <tr><td style="${FONT}padding:22px 24px 8px;font-size:16px;font-weight:600;color:#202124;">
      <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};margin-right:8px;"></span>${h(title)} (${list.length})</td></tr>
    <tr><td style="padding:0 24px;"><table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;${FONT}font-size:13px;">
      <tr>${cols.map(c => `<th align="left" style="background:#f1f3f4;color:#3c4043;padding:8px 10px;border-bottom:2px solid #dadce0;">${h(c[0])}</th>`).join('')}</tr>
      ${list.map((r, i) => `<tr style="background:${i % 2 ? '#fafafa' : '#fff'};">${cols.map(c => `<td style="padding:7px 10px;border-bottom:1px solid #eee;color:#3c4043;vertical-align:top;">${h(c[1](r))}</td>`).join('')}</tr>`).join('')}
    </table></td></tr>`;

  const place = [['MIS Page', r => r.page], ['Section', r => r.section || 'All Reports'], ['Report', r => r.report]];
  const withDetail = [...place, ['Details', r => r.detail]];
  let changesHtml;
  if (!changes) changesHtml = `<tr><td style="${FONT}padding:16px 24px;color:#5f6368;font-size:13px;">No earlier run to compare with - this run is the baseline.</td></tr>`;
  else {
    const n = changes.added.length + changes.removed.length + changes.moved.length + changes.columns.length;
    changesHtml = `<tr><td style="${FONT}padding:16px 24px 0;font-size:13px;color:#5f6368;">Compared with <b>${h(previousName)}</b>: ` +
      (n ? `<b style="color:#d93025;">${n} change(s)</b>` : `<b style="color:#188038;">no changes</b> - same reports, sections and columns.`) + `</td></tr>` +
      table('Reports added', '#188038', changes.added, place) +
      table('Reports removed', '#d93025', changes.removed, place) +
      table('Moved to another section', '#f9ab00', changes.moved, [['MIS Page', c => c.page], ['Report', c => c.report], ['From', c => c.from || 'All Reports'], ['To', c => c.to || 'All Reports']]) +
      table('Columns changed', '#e8710a', changes.columns, [['MIS Page', c => c.page], ['Report', c => c.report],
        ['Added', c => c.addedCols.join(', ')], ['Removed', c => c.removedCols.join(', ')]]);
  }

  // Where the run came from, so GitHub and local-PC emails can be told apart in the inbox.
  const runUrl = process.env.GITHUB_ACTIONS
    ? `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`
    : '';
  const sourceTag = runUrl ? 'GitHub' : 'Local';
  const sourceHtml = runUrl
    ? `Run from GitHub Actions - <a href="${h(runUrl)}" style="color:#fff;">open this run</a>`
    : `Run from local PC (${h(require('os').hostname())})`;

  const html = `
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:24px 0;"><tr><td align="center">
  <table width="760" cellpadding="0" cellspacing="0" style="max-width:760px;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08);">
    <tr><td style="background:linear-gradient(90deg,#0b5394,#1a73e8);background-color:#0b5394;padding:26px 24px;${FONT}color:#fff;">
      <div style="font-size:22px;font-weight:700;">River DMS - MIS Report Check</div>
      <div style="font-size:13px;opacity:.9;margin-top:6px;">${h(new Date().toLocaleString('en-IN'))} &nbsp;|&nbsp; Report dates ${h(from)} to ${h(to)} &nbsp;|&nbsp; ${minutes} min</div>
      <div style="font-size:13px;opacity:.9;margin-top:4px;">${sourceHtml}</div>
    </td></tr>
    <tr><td style="padding:18px 18px 4px;background:#f8f9fa;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
      ${card('Reports checked', total, '#1a73e8')}${card('Passed', ok, '#188038')}${card('Failed', failed, '#d93025')}${card('Blank / no data', count('BLANK / NO DATA'), '#f9ab00')}
    </tr></table></td></tr>
    <tr><td style="padding:10px 24px 18px;background:#f8f9fa;">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>
        <td style="background:#e6e6e6;border-radius:6px;height:10px;"><div style="width:${pct}%;background:#188038;height:10px;border-radius:6px;"></div></td>
        <td width="70" style="${FONT}font-size:12px;color:#5f6368;padding-left:10px;">${pct}% pass</td></tr></table>
      <div style="${FONT}font-size:12px;color:#5f6368;margin-top:8px;">Needs extra input: ${count('NEEDS EXTRA INPUT')} &nbsp;&middot;&nbsp; Timeout: ${count('TIMEOUT')} &nbsp;&middot;&nbsp; No report opened: ${count('NO REPORT OPENED')}</div>
    </td></tr>
    <tr><td style="${FONT}padding:22px 24px 0;font-size:18px;font-weight:700;color:#202124;border-top:1px solid #eee;">Changes since last run</td></tr>
    ${changesHtml}
    <tr><td style="${FONT}padding:26px 24px 0;font-size:18px;font-weight:700;color:#202124;">Reports that need attention</td></tr>
    ${table('Server error', '#d93025', rows.filter(r => r.result === 'SERVER ERROR'), withDetail)}
    ${table('Timeout', '#e8710a', rows.filter(r => r.result === 'TIMEOUT'), withDetail)}
    ${table('No report opened', '#9334e6', rows.filter(r => r.result === 'NO REPORT OPENED' || r.result === 'COULD NOT SELECT'), withDetail)}
    ${table('Needs extra input', '#f9ab00', rows.filter(r => r.result === 'NEEDS EXTRA INPUT'), withDetail)}
    <tr><td style="${FONT}padding:24px;font-size:12px;color:#80868b;">
      The full list (every report with its definition and available columns) is in the attached Excel file.
      ${count('BLANK / NO DATA')} reports opened but returned no data for the selected dates - see the "Check Result" column.
    </td></tr>
  </table></td></tr></table>`;

  const transporter = nodemailer.createTransport({ service: 'gmail', auth: { user, pass } });
  await transporter.sendMail({
    from: user, to: rcpt,
    subject: `MIS Report Check [${sourceTag}] - ${ok}/${total} passed, ${failed} failed${changes ? ` - ${changes.added.length} added / ${changes.removed.length} removed` : ''} - ${new Date().toLocaleDateString('en-IN')}`,
    text: `MIS Report Check: ${total} reports, ${ok} passed, ${failed} failed, ${count('BLANK / NO DATA')} blank. Full details in the attached Excel file.\n${runUrl ? `Run from: GitHub Actions (${runUrl})` : `Run from: local PC (${require('os').hostname()})`}`,
    html,
    attachments: attachments.map(p => ({ filename: path.basename(p), path: p })),
  });
  console.log(`Summary emailed to ${rcpt}`);
}

// ---------- Reports that open their own page in a new tab ----------
// Opens the page for real (it builds its table with its own scripts), then uses the headings of the
// main data table as the report's objects.
async function checkNewTabReport(context, rec) {
  const t0 = Date.now();
  const tab = await context.newPage();
  const errors = [];
  tab.on('pageerror', e => errors.push(String(e.message || e)));
  tab.on('dialog', d => d.dismiss().catch(() => { }));
  try {
    const resp = await tab.goto(rec.openUrl, { waitUntil: 'load', timeout: CONFIG.reportTimeoutSec * 2000 });
    await tab.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => { });
    await tab.waitForTimeout(1500);
    const info = await tab.evaluate(() => {
      const clean = s => s.textContent.trim().replace(/\s+/g, ' ');
      // Pick the table the user actually sees: visible, with data rows first, then the most headings.
      let best = { heads: [], rows: 0, score: -1 };
      for (const t of document.querySelectorAll('table')) {
        if (t.querySelector('input:not([type=hidden]), select, textarea, table')) continue; // filter/layout tables
        if (!t.getClientRects().length) continue; // hidden drill-down tables
        const headRow = t.querySelector('thead tr') || [...t.rows].find(r => r.querySelector('th')) || t.rows[0];
        if (!headRow) continue;
        const heads = [...headRow.cells].map(clean).filter(h => h && h.length <= 60);
        const rows = Math.max(0, t.rows.length - (t.tHead ? t.tHead.rows.length : 1));
        const score = (rows > 0 ? 1000 : 0) + heads.length;
        if (heads.length > 1 && score > best.score) best = { heads, rows, score };
      }
      return { heads: best.heads, rows: best.rows, title: document.title, text: document.body.innerText.slice(0, 1500) };
    });
    rec.objects = info.heads;
    rec.rows = info.rows;
    if (!resp || resp.status() !== 200 || /Server Error|Runtime Error|Exception/i.test(info.title + ' ' + info.text)) {
      rec.result = 'SERVER ERROR'; rec.detail = `Own page did not load (${resp ? 'HTTP ' + resp.status() : 'no response'})`;
    } else if (info.rows > 0) {
      rec.result = 'OK'; rec.detail = 'Opens in new tab';
    } else {
      rec.result = 'BLANK / NO DATA'; rec.detail = 'Opens in new tab, table empty' + (errors.length ? ` (script error: ${errors[0].slice(0, 80)})` : '');
    }
  } catch (e) {
    rec.result = /Timeout/i.test(e.message) ? 'TIMEOUT' : 'SERVER ERROR';
    rec.detail = `Opens in new tab: ${String(e.message || e).split('\n')[0].slice(0, 120)}`;
  }
  rec.seconds = Math.round((Date.now() - t0) / 1000);
  await tab.close();
}

// ---------- Formatted Excel: one row per report, grouped by MIS page and section ----------
// Layout: S.No | Master Report Name | Report Type | Report Name | Report Definition | Objects Available | Check Result
// Master Report Name / Report Type cells are merged over their group.
// Red    = Report Name / Definition of a report that could not be viewed (error, timeout, nothing opened),
//          or a Report Name / Definition cell that is empty
// Orange = Objects Available could not be read (no Customize Columns list and no grid header)
async function writeObjectsWorkbook(rows, file, changes, previousName) {
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('MIS Reports', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [
    { header: 'S.No', width: 6 },
    { header: 'Master Report Name', width: 22 },
    { header: 'Report Type', width: 22 },
    { header: 'Report Name', width: 40 },
    { header: 'Report Definition', width: 55 },
    { header: 'Objects Available', width: 110 },
    { header: 'Check Result', width: 30 },
  ];
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const FILL = { header: fill('FFC9DAF8'), master: fill('FFD9EAD3'), type: fill('FFFFF2CC'), red: fill('FFFF0000'), orange: fill('FFFF9900') };
  const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const FAILED = ['SERVER ERROR', 'TIMEOUT', 'NO REPORT OPENED', 'COULD NOT SELECT'];
  const PER_LINE = 6; // objects per line inside the cell, like the sample sheet

  ws.getRow(1).eachCell(c => {
    c.fill = FILL.header; c.font = { bold: true }; c.border = border;
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Keep each section's reports together (stable: pages and sections stay in menu order)
  const firstSeen = new Map();
  rows.forEach((r, i) => { const k = `${r.page}|${r.section || 'All Reports'}`; if (!firstSeen.has(k)) firstSeen.set(k, i); });
  rows = rows.map((r, i) => ({ r, i, g: firstSeen.get(`${r.page}|${r.section || 'All Reports'}`) }))
    .sort((a, b) => a.g - b.g || a.i - b.i).map(x => x.r);

  const merges = []; // [col, startRow, endRow]
  let masterStart = 2, typeStart = 2;
  rows.forEach((r, i) => {
    const objs = r.objects || [];
    const lines = [];
    for (let j = 0; j < objs.length; j += PER_LINE) lines.push(objs.slice(j, j + PER_LINE).join('      '));
    const result = r.result + (r.detail ? ` - ${r.detail}` : '');
    const row = ws.addRow([i + 1, r.page, r.section || 'All Reports', r.report, r.definition || '', lines.join('\n'), result]);
    const rn = row.number;
    row.eachCell({ includeEmpty: true }, c => { c.border = border; c.alignment = { vertical: 'middle', wrapText: true }; });
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).fill = FILL.master; row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    row.getCell(3).fill = FILL.type; row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    row.getCell(6).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    if (FAILED.includes(r.result)) { row.getCell(4).fill = FILL.red; row.getCell(5).fill = FILL.red; }
    if (!String(r.report || '').trim()) row.getCell(4).fill = FILL.red;     // report name missing
    if (!String(r.definition || '').trim()) row.getCell(5).fill = FILL.red; // definition missing
    if (!objs.length) row.getCell(6).fill = FILL.orange;
    row.height = Math.max(15, 15 * Math.max(1, lines.length));

    // close merge groups when the page / section changes
    const nxt = rows[i + 1];
    const typeKey = x => `${x.page}|${x.section || 'All Reports'}`;
    if (!nxt || typeKey(nxt) !== typeKey(r)) { if (rn > typeStart) merges.push([3, typeStart, rn]); typeStart = rn + 1; }
    if (!nxt || nxt.page !== r.page) { if (rn > masterStart) merges.push([2, masterStart, rn]); masterStart = rn + 1; }
  });
  merges.forEach(([col, s, e]) => ws.mergeCells(s, col, e, col));
  // (written at the end of this function; the caller catches a failure so CSV/TXT are kept)

  // Second sheet: what changed since the previous run
  const cs = wb.addWorksheet('Changes vs Previous', { views: [{ state: 'frozen', ySplit: 2 }] });
  cs.columns = [{ width: 12 }, { width: 22 }, { width: 22 }, { width: 40 }, { width: 70 }];
  cs.getCell('A1').value = changes ? `Compared with previous run: ${previousName}` : 'No earlier run found - this run is the baseline for the next comparison.';
  cs.getCell('A1').font = { bold: true };
  const hdr = cs.addRow(['Change', 'Master Report Name', 'Report Type', 'Report Name', 'Details']);
  hdr.eachCell(c => { c.fill = FILL.header; c.font = { bold: true }; c.border = border; });
  const CHANGE_FILL = { Added: fill('FFB6D7A8'), Removed: FILL.red, Moved: fill('FFFFE599'), 'Columns changed': FILL.orange };
  const addChange = (kind, page, section, report, details) => {
    const r = cs.addRow([kind, page, section || 'All Reports', report, details]);
    r.eachCell({ includeEmpty: true }, c => { c.border = border; c.alignment = { vertical: 'middle', wrapText: true }; });
    r.getCell(1).fill = CHANGE_FILL[kind];
  };
  if (changes) {
    changes.added.forEach(r => addChange('Added', r.page, r.section, r.report, 'New report in this run'));
    changes.removed.forEach(r => addChange('Removed', r.page, r.section, r.report, 'Not in the menu any more'));
    changes.moved.forEach(c => addChange('Moved', c.page, c.to, c.report, `Section changed: ${c.from || 'All Reports'} -> ${c.to || 'All Reports'}`));
    changes.columns.forEach(c => addChange('Columns changed', c.page, c.section, c.report,
      [c.addedCols.length && `Added: ${c.addedCols.join(', ')}`, c.removedCols.length && `Removed: ${c.removedCols.join(', ')}`].filter(Boolean).join('\n')));
    if (cs.rowCount === 2) cs.addRow(['No changes', '', '', '', 'Same reports, sections and columns as the previous run']);
  }
  await wb.xlsx.writeFile(file);
}

// Exit codes: 0 = run finished (see the summary for report failures), 1 = the run itself failed
// (browser/login/site down), 2 = stopped by the MIS_MAX_MINUTES watchdog.
runCheck()
  .catch(e => { console.error(`\nMIS check failed: ${e && e.message || e}`); process.exitCode = 1; })
  .finally(() => process.exit(process.exitCode || 0)); // don't let a stray handle keep CI waiting

async function runCheck() {
  const channel = process.env.BROWSER_CHANNEL || 'chrome'; // installed Google Chrome, or msedge / chromium
  let context;
  try {
    context = await chromium.launchPersistentContext(PROFILE_DIR, {
      channel, headless: CONFIG.headless, viewport: { width: 1300, height: 800 },
    });
  } catch (e) {
    throw new Error(`Could not start the browser (channel "${channel}"): ${String(e.message).split('\n')[0]}. ` +
      `Install it, or set BROWSER_CHANNEL=msedge / chromium (after "npx playwright install chromium"). ` +
      `If another MIS check is still running, wait for it - the dms-profile folder can only be used by one at a time.`);
  }
  try {
    await runWithContext(context);
  } finally {
    await withTimeout(context.close(), 30000, 'Closing the browser').catch(() => { });
  }
}

async function runWithContext(context) {
  // No step may wait forever: every click/fill/wait and every page load has a limit.
  context.setDefaultTimeout(30000);
  context.setDefaultNavigationTimeout(90000);
  const main = context.pages()[0] || await context.newPage();
  await main.goto(CONFIG.baseUrl + CONFIG.startPage, { waitUntil: 'load' });

  const readMenu = () => main.evaluate(() => {
    const li = [...document.querySelectorAll('li')].find(l => {
      const a = l.querySelector(':scope > a'); return a && /^\s*MIS\s*$/i.test(a.textContent);
    });
    return li ? [...li.querySelectorAll('ul a')].map(a => ({ name: a.textContent.trim(), url: a.href })) : null;
  });

  // Same dealer-code login as tests/models/LoginPage.ts
  async function login(creds) {
    console.log('Logging in to the DMS...');
    await main.goto(CONFIG.baseUrl + '/', { waitUntil: 'domcontentloaded' });
    const dealerCode = main.locator('input[name="DealerCode"]');
    await dealerCode.fill(creds.dealercode || '');
    // The app only looks up branches on blur, so Tab out of the field.
    await dealerCode.press('Tab');
    if (creds.dealercode) {
      await main.waitForFunction(() => {
        const el = document.querySelector('select[name="BranchSlno"]');
        return !!el && el.options.length > 0 && !!el.value;
      }, null, { timeout: 10000 }).catch(() => {});
    }
    await main.getByRole('textbox', { name: 'User Name' }).fill(creds.username || '');
    await main.getByRole('textbox', { name: 'Password' }).fill(creds.password || '');
    await main.getByRole('button', { name: 'Sign In' }).click();
    await main.locator('#ctl00_cpMain_DashboardCaption').waitFor({ state: 'visible', timeout: 30000 });
    console.log('Logged in.');
  }

  let pages = await readMenu();
  if (!pages) {
    if (!credentials.dealervalidUser) throw new Error('resources/credentials.json has no "dealervalidUser" to log in with.');
    try {
      await login(credentials.dealervalidUser);
    } catch (e) {
      console.log(`Automatic login failed (${e.message.split('\n')[0]}).`);
      await ask('\nPlease log in to the DMS in the browser window, then press Enter here... '); // throws on CI
    }
    await main.goto(CONFIG.baseUrl + CONFIG.startPage, { waitUntil: 'load' });
    pages = await readMenu();
    if (!pages) throw new Error('Could not find the MIS menu after logging in - check the dealervalidUser credentials.');
  }
  // The menu's own "MIS" entry just points back at Accounts Reports (url ends in "#") - skip the duplicate.
  pages = pages.filter(p => !/^MIS$/i.test(p.name));
  if (CONFIG.onlyPages.length) pages = pages.filter(p => CONFIG.onlyPages.includes(p.name));
  console.log(`\nTesting ${pages.length} MIS pages, dates ${FROM} to ${TO}, ${CONFIG.concurrency} at a time\n`);

  const cfg = { from: FROM, to: TO, today: fmt(yesterday), extra: CONFIG.extraInputs,
    purchaseDate: fmt(new Date(now.getFullYear(), now.getMonth(), now.getDate() - CONFIG.extraInputs.purchaseDateDaysAgo)),
    month: pad(now.getMonth() + 1), year: String(now.getFullYear()), timeoutMs: CONFIG.reportTimeoutSec * 1000, retryTimeoutMs: 90000 };
  const byPage = new Array(pages.length);
  let next = 0;
  const started = Date.now();

  async function worker() {
    while (next < pages.length) {
      const idx = next++;
      const pg = pages[idx];
      let tab = null;
      try {
        tab = await context.newPage();
        await tab.exposeFunction('reportProgress', m => console.log(`[${pg.name}] ${m}`));
        const resp = await tab.goto(pg.url, { waitUntil: 'load', timeout: 90000 });
        const text = await tab.content();
        if (!resp || resp.status() !== 200 || /Server Error|Runtime Error/i.test(text.slice(0, 5000))) {
          byPage[idx] = [{ section: '', report: '(menu link)', result: 'SERVER ERROR', detail: 'MIS page itself did not open', rows: '', seconds: '' }];
        } else {
          await tab.waitForTimeout(2500); // let the page's scripts finish
          // Each report inside has its own time limit; this caps the page as a whole.
          byPage[idx] = await withTimeout(tab.evaluate(checkPageInBrowser, cfg), CONFIG.maxPageMinutes * 60000, `MIS page "${pg.name}"`);
          for (const rec of byPage[idx].filter(r => r.result === 'NEW TAB')) {
            await withTimeout(checkNewTabReport(context, rec), 5 * 60000, `${rec.report} (own page)`)
              .catch(e => { rec.result = 'TIMEOUT'; rec.detail = e.message; });
            console.log(`[${pg.name}] ${rec.report} (own page) -> ${rec.result}`);
          }
        }
      } catch (e) {
        const timedOut = /took longer than|Timeout/i.test(String(e.message || e));
        byPage[idx] = [{ section: '', report: '(page)', result: timedOut ? 'TIMEOUT' : 'SERVER ERROR', detail: String(e.message || e).slice(0, 150), rows: '', seconds: '' }];
      } finally {
        if (tab) await withTimeout(tab.close(), 15000, 'Closing tab').catch(() => { });
      }
      console.log(`=== Finished ${pg.name} (${byPage[idx].length} reports) ===`);
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONFIG.concurrency, pages.length) }, worker));

  // ---------- Build outputs ----------
  const rows = [];
  pages.forEach((pg, i) => (byPage[i] || []).forEach(r => rows.push({
    path: `MIS > ${pg.name}${r.section ? ' > ' + r.section : ''} > ${r.report}`,
    page: pg.name, ...r,
  })));

  rows.forEach(r => { r.objects = r.objects || []; r.objectsText = r.objects.join(', '); });

  // All outputs go to the "MIS Reports" folder; compare with the newest earlier run found there.
  fs.mkdirSync(OUT_DIR, { recursive: true });
  let previous = null;
  try { previous = loadPreviousRun(OUT_DIR); }
  catch (e) { console.log(`Previous run could not be read (${e.message}) - skipping the comparison this time.`); }
  const changes = previous ? compareRuns(previous.rows, rows) : null;
  fs.writeFileSync(path.join(OUT_DIR, `MIS_Report_Snapshot_${stamp}.json`), JSON.stringify(
    rows.map(r => ({ page: r.page, section: r.section, report: r.report, definition: r.definition, objects: r.objects, result: r.result })), null, 1));

  const cols = ['path', 'page', 'section', 'report', 'definition', 'objectsText', 'result', 'detail', 'rows', 'seconds'];
  const esc = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csvFile = path.join(OUT_DIR, `MIS_Report_Check_${stamp}.csv`);
  fs.writeFileSync(csvFile, '﻿' + [cols.join(','), ...rows.map(r => cols.map(c => esc(r[c])).join(','))].join('\r\n'));

  const count = k => rows.filter(r => r.result === k).length;
  const lines = [];
  lines.push(`MIS Report Check - ${new Date().toLocaleString()}   (dates ${FROM} to ${TO})`);
  lines.push('='.repeat(70));
  lines.push(`Total reports tested : ${rows.length}`);
  lines.push(`Passed (OK)          : ${count('OK')}`);
  lines.push(`Server error         : ${count('SERVER ERROR')}`);
  lines.push(`Timeout              : ${count('TIMEOUT')}`);
  lines.push(`Blank / no data      : ${count('BLANK / NO DATA')}`);
  lines.push(`Needs extra input    : ${count('NEEDS EXTRA INPUT')}`);
  lines.push(`No report opened     : ${count('NO REPORT OPENED')}`);
  lines.push(`Could not select     : ${count('COULD NOT SELECT')}`);
  lines.push(`Time taken           : ${Math.round((Date.now() - started) / 60000)} min`);
  for (const cat of ['SERVER ERROR', 'TIMEOUT', 'BLANK / NO DATA', 'NO REPORT OPENED', 'COULD NOT SELECT', 'NEEDS EXTRA INPUT']) {
    const list = rows.filter(r => r.result === cat);
    if (!list.length) continue;
    lines.push('', `--- ${cat} (${list.length}) ---`);
    list.forEach(r => lines.push(`  ${r.path}${r.detail ? '   [' + r.detail + ']' : ''}`));
  }

  lines.push('', '='.repeat(70));
  if (!changes) lines.push('CHANGES: no earlier run found in "MIS Reports" - this run is the baseline for next time.');
  else {
    lines.push(`CHANGES since previous run (${previous.name})`);
    lines.push(`  Reports added   : ${changes.added.length}`);
    lines.push(`  Reports removed : ${changes.removed.length}`);
    lines.push(`  Moved section   : ${changes.moved.length}`);
    lines.push(`  Columns changed : ${changes.columns.length}`);
    const show = (title, list, fmt) => { if (list.length) { lines.push('', `--- ${title} (${list.length}) ---`); list.forEach(x => lines.push('  ' + fmt(x))); } };
    show('ADDED', changes.added, r => `${r.page} > ${r.section} > ${r.report}`);
    show('REMOVED', changes.removed, r => `${r.page} > ${r.section} > ${r.report}`);
    show('MOVED', changes.moved, c => `${c.page} > ${c.report}: ${c.from} -> ${c.to}`);
    show('COLUMNS CHANGED', changes.columns, c => `${c.page} > ${c.report}: ${c.addedCols.length ? '+ ' + c.addedCols.join(', ') : ''}${c.addedCols.length && c.removedCols.length ? '  ' : ''}${c.removedCols.length ? '- ' + c.removedCols.join(', ') : ''}`);
  }

  const txtFile = path.join(OUT_DIR, `MIS_Report_Summary_${stamp}.txt`);
  fs.writeFileSync(txtFile, lines.join('\r\n'));

  // The CSV/TXT above are already saved, so a problem with the Excel file doesn't lose the run.
  let xlsxFile = path.join(OUT_DIR, `MIS_Report_Objects_${stamp}.xlsx`);
  try {
    await writeObjectsWorkbook(rows, xlsxFile, changes, previous && previous.name);
  } catch (e) {
    console.error(`Excel file could not be written (${e.message}).`);
    xlsxFile = null;
  }

  console.log('\n' + lines.join('\n'));
  console.log(`\nSaved:\n  ${csvFile}\n  ${txtFile}${xlsxFile ? '\n  ' + xlsxFile : ''}`);

  if (CONFIG.sendEmail) {
    try {
      await withTimeout(sendSummaryEmail({ rows, changes, previousName: previous && previous.name, from: FROM, to: TO,
        minutes: Math.round((Date.now() - started) / 60000), attachments: xlsxFile ? [xlsxFile] : [csvFile] }), 2 * 60000, 'Sending the email');
    } catch (e) { console.error('Email failed: ' + (e.message || e)); }
  }
  if (!rows.length) throw new Error('No reports were checked - the MIS pages did not load.');
}