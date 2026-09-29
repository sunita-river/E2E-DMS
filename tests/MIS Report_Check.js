/*
 * MIS Report Checker for River DMS - run from VS Code
 * ---------------------------------------------------
 * Setup (one time, in the VS Code terminal, inside this folder):
 *     npm install
 * Run:
 *     npm start
 *
 * A Chrome window opens. The first time, log in to the DMS in that window,
 * then come back to the VS Code terminal and press Enter. Your login is kept
 * in the "dms-profile" folder, so next time it goes straight in.
 *
 * Output (same folder):
 *     MIS_Report_Check_<date>.csv   - every report with its full path and result
 *     MIS_Report_Summary_<date>.txt - totals + list of failures
 *
 * It only opens/views reports. It does not save, edit or delete anything in the DMS.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// ======================= SETTINGS =======================
const CONFIG = {
  baseUrl: 'https://rivermobility.gaindms.com',
  startPage: '/Accounts/AccountsReports.aspx', // any page that shows the MIS menu
  fromDate: null,        // 'dd/mm/yyyy'  - null = 1st of current month
  toDate: null,          // 'dd/mm/yyyy'  - null = today
  concurrency: 4,        // how many MIS pages to test at the same time
  reportTimeoutSec: 30,  // a report slower than this is marked TIMEOUT
  headless: false,       // true = no visible browser (only after first login)
  onlyPages: [],         // e.g. ['Workshop Reports', 'CRM Reports'] ; empty = all MIS pages
};
// ========================================================

const pad = n => String(n).padStart(2, '0');
const fmt = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const now = new Date();
const FROM = CONFIG.fromDate || fmt(new Date(now.getFullYear(), now.getMonth(), 1));
const TO = CONFIG.toDate || fmt(now);
const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;

function ask(q) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(res => rl.question(q, a => { rl.close(); res(a); }));
}

// ---------- This function runs INSIDE the DMS page (browser side) ----------
async function checkPageInBrowser(cfg) {
  const ERR_RE = /(Server Error|Runtime Error|Exception|error occurred|Invalid (object|column)|Timeout expired|Could not find|Incorrect syntax|Conversion failed|Procedure or function|Something went wrong|Object reference|Gateway Time-out|Bad Gateway|Service Unavailable)/i;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const scripts = t => [...t.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1].trim());
  const fetchT = async (url, opts = {}) => {
    const c = new AbortController(); const t = setTimeout(() => c.abort(), cfg.timeoutMs);
    try { return await fetch(url, { credentials: 'include', ...opts, signal: c.signal }); }
    finally { clearTimeout(t); }
  };
  const log = m => { try { window.reportProgress && window.reportProgress(m); } catch (e) { } };

  // The page's own highlight function crashes when reports are clicked by a script; make it safe.
  if (window.HighlightSelected) {
    const o = window.HighlightSelected;
    window.HighlightSelected = function () { try { return o.apply(this, arguments); } catch (e) { } };
  }
  const f = document.forms[0];
  const base = new Set(scripts(await (await fetch(location.href, { credentials: 'include' })).text()));
  const seen = new Set();
  const anchors = [...document.querySelectorAll('a[onclick*="SetTab"]')].filter(a => {
    const id = (a.getAttribute('onclick').match(/SetTab\(this\.id,(\d+)/) || [])[1];
    if (!id || seen.has(id)) return false; seen.add(id); return true;
  });

  const post = async () => {
    const fd = new FormData(f);
    fd.set('__EVENTTARGET', 'ctl00$cpMain$btnViewReport'); fd.set('__EVENTARGUMENT', '');
    const r = await fetchT(new URL(f.getAttribute('action'), location.href), { method: 'POST', body: new URLSearchParams(fd) });
    const t = await r.text(); const d = new DOMParser().parseFromString(t, 'text/html');
    const msg = ((d.getElementById('ctl00_cpMain_lblMsg')?.textContent || '') + ' ' +
                 (d.getElementById('ctl00_cpMain_LblMsg')?.textContent || '')).trim();
    return { r, t, d, msg };
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
    const rec = { section, report: name, result: '', detail: '', rows: '', seconds: '' };
    const t0 = Date.now();
    try {
      // 1. select the report (same as clicking it)
      const rs = document.getElementById('ctl00_cpMain_ReportSlno');
      let ok = false;
      for (const c of [...document.querySelectorAll(`a[onclick*="SetTab(this.id,${rid},"]`)].reverse()) {
        const tv = document.getElementById('ctl00_cpMain_tabval'); if (tv) tv.value = '';
        try { c.click(); } catch (e) { }
        await sleep(300);
        if (!rs || rs.value == rid) { ok = true; break; }
      }
      if (!ok) { rec.result = 'COULD NOT SELECT'; rec.detail = 'Report link click failed'; }
      else {
        // 2. fill dates
        const vis = [...f.querySelectorAll('input,select,textarea')]
          .filter(e => e.offsetParent && e.type !== 'hidden' && !/ddlSearchOption|txtSearch/i.test(e.id));
        vis.filter(e => e.tagName === 'INPUT' && (e.type === 'text' || e.type === '') && /Dat/i.test(e.id + e.name))
          .forEach(e => { e.value = /To|End/i.test(e.id) ? cfg.to : cfg.from; });
        // 3. View Report (retry with first branch/franchise if the DMS asks)
        let { r, t, d, msg } = await post();
        for (let i = 0; i < 2 && /select|enter/i.test(msg); i++) {
          vis.filter(e => e.tagName === 'SELECT' && (e.value === '' || e.value === '0') && e.options.length > 1)
            .forEach(e => e.selectedIndex = 1);
          ({ r, t, d, msg } = await post());
        }
        const nw = scripts(t).filter(s => !base.has(s)).join('\n');
        const m = nw.match(/window\.open\s*\(\s*'([^']+)'/);
        const msgs = [...nw.matchAll(/(?:alert|Messagebox|MessageBox|ShowMessage)\s*\(\s*['"]([^'"]{0,200})/g)].map(x => x[1]);
        if (msg) msgs.push(msg);

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
          if (r2.status !== 200 || i >= 0) {
            rec.result = 'SERVER ERROR';
            rec.detail = ((d2.title || `HTTP ${r2.status}`) + ' ' + (i >= 0 ? body.slice(i, i + 120) : '')).trim();
          } else if (rec.rows > 0) rec.result = 'OK';
          else {
            rec.result = 'BLANK / NO DATA';
            rec.detail = u.pathname.includes('Report.aspx') ? 'Crystal viewer shows blank page' : 'Viewer window empty';
          }
        }
      }
    } catch (e) {
      if (e.name === 'AbortError') { rec.result = 'TIMEOUT'; rec.detail = `Slower than ${cfg.timeoutMs / 1000}s`; }
      else { rec.result = 'SERVER ERROR'; rec.detail = String(e).slice(0, 150); }
    }
    rec.seconds = Math.round((Date.now() - t0) / 1000);
    results.push(rec);
    log(`${k}/${anchors.length} ${name} -> ${rec.result}`);
  }
  return results;
}
// ----------------------------------------------------------------------------

(async () => {
  const context = await chromium.launchPersistentContext(path.join(__dirname, 'dms-profile'), {
    channel: 'chrome',            // uses your installed Google Chrome
    headless: CONFIG.headless,
    viewport: { width: 1300, height: 800 },
  });
  const main = context.pages()[0] || await context.newPage();
  await main.goto(CONFIG.baseUrl + CONFIG.startPage, { waitUntil: 'load' });

  const readMenu = () => main.evaluate(() => {
    const li = [...document.querySelectorAll('li')].find(l => {
      const a = l.querySelector(':scope > a'); return a && /^\s*MIS\s*$/i.test(a.textContent);
    });
    return li ? [...li.querySelectorAll('ul a')].map(a => ({ name: a.textContent.trim(), url: a.href })) : null;
  });

  let pages = await readMenu();
  if (!pages) {
    await ask('\nPlease log in to the DMS in the Chrome window, then press Enter here... ');
    await main.goto(CONFIG.baseUrl + CONFIG.startPage, { waitUntil: 'load' });
    pages = await readMenu();
    if (!pages) { console.error('Could not find the MIS menu. Are you logged in?'); await context.close(); return; }
  }
  if (CONFIG.onlyPages.length) pages = pages.filter(p => CONFIG.onlyPages.includes(p.name));
  console.log(`\nTesting ${pages.length} MIS pages, dates ${FROM} to ${TO}, ${CONFIG.concurrency} at a time\n`);

  const cfg = { from: FROM, to: TO, timeoutMs: CONFIG.reportTimeoutSec * 1000 };
  const byPage = new Array(pages.length);
  let next = 0;
  const started = Date.now();

  async function worker() {
    while (next < pages.length) {
      const idx = next++;
      const pg = pages[idx];
      const tab = await context.newPage();
      await tab.exposeFunction('reportProgress', m => console.log(`[${pg.name}] ${m}`));
      try {
        const resp = await tab.goto(pg.url, { waitUntil: 'load', timeout: 90000 });
        const text = await tab.content();
        if (!resp || resp.status() !== 200 || /Server Error|Runtime Error/i.test(text.slice(0, 5000))) {
          byPage[idx] = [{ section: '', report: '(menu link)', result: 'SERVER ERROR', detail: 'MIS page itself did not open', rows: '', seconds: '' }];
        } else {
          await tab.waitForTimeout(2500); // let the page's scripts finish
          byPage[idx] = await tab.evaluate(checkPageInBrowser, cfg);
        }
      } catch (e) {
        byPage[idx] = [{ section: '', report: '(page)', result: 'SERVER ERROR', detail: String(e.message || e).slice(0, 150), rows: '', seconds: '' }];
      }
      console.log(`=== Finished ${pg.name} (${byPage[idx].length} reports) ===`);
      await tab.close();
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONFIG.concurrency, pages.length) }, worker));

  // ---------- Build outputs ----------
  const rows = [];
  pages.forEach((pg, i) => (byPage[i] || []).forEach(r => rows.push({
    path: `MIS > ${pg.name}${r.section ? ' > ' + r.section : ''} > ${r.report}`,
    page: pg.name, ...r,
  })));

  const cols = ['path', 'page', 'section', 'report', 'result', 'detail', 'rows', 'seconds'];
  const esc = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csvFile = path.join(__dirname, `MIS_Report_Check_${stamp}.csv`);
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
  const txtFile = path.join(__dirname, `MIS_Report_Summary_${stamp}.txt`);
  fs.writeFileSync(txtFile, lines.join('\r\n'));

  console.log('\n' + lines.join('\n'));
  console.log(`\nSaved:\n  ${csvFile}\n  ${txtFile}`);
  await context.close();
})();