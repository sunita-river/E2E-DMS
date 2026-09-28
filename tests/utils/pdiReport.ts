import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import nodemailer from 'nodemailer';

// Per-dealer summary email + interactive HTML dashboard for tests/PDI.spec.ts.
// Mail apps strip JavaScript, so the email body is a static (table-based, inline-styled)
// summary and the interactive dashboard travels as an attached .html file to open in a browser.

export type ResultRow = Record<string, string>;
export interface FailedDealer { dealer: string; error: string }
// title labels the report, e.g. "PDI vehicle details · Daily summary" (default: "PDI vehicle details").
export interface ReportMeta { dryRun: boolean; startedAt: Date; finishedAt: Date; dealers: string[]; vinFile: string; title?: string }
const titleOf = (m: ReportMeta) => m.title || 'PDI vehicle details';

type Outcome = 'updated' | 'dryrun' | 'missing' | 'error';
const OUTCOMES: { key: Outcome; label: string; icon: string; color: string }[] = [
  { key: 'updated', label: 'Updated', icon: '✔', color: '#0ca30c' },
  { key: 'dryrun', label: 'Filled (dry run)', icon: '◌', color: '#2a78d6' },
  { key: 'missing', label: 'Not in VIN sheet', icon: '▲', color: '#fab219' },
  { key: 'error', label: 'Error', icon: '✖', color: '#d03b3b' },
];

export function outcomeOf(row: ResultRow): Outcome {
  const s = row.Status || '';
  if (s.startsWith('Updated')) return 'updated';
  if (s.startsWith('Dry run')) return 'dryrun';
  if (s.startsWith('Not in VIN')) return 'missing';
  return 'error';
}

interface DealerSummary { dealer: string; failed?: string; total: number; counts: Record<Outcome, number> }

function summarize(results: ResultRow[], meta: ReportMeta, failed: FailedDealer[]): DealerSummary[] {
  return meta.dealers.map((dealer) => {
    const mine = results.filter((r) => r.Dealer === dealer);
    const counts = { updated: 0, dryrun: 0, missing: 0, error: 0 } as Record<Outcome, number>;
    mine.forEach((r) => counts[outcomeOf(r)]++);
    return { dealer, failed: failed.find((f) => f.dealer === dealer)?.error, total: mine.length, counts };
  });
}

const esc = (s: unknown) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const fmtTime = (d: Date) => d.toLocaleString('en-IN', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const minutes = (m: ReportMeta) => Math.max(1, Math.round((m.finishedAt.getTime() - m.startedAt.getTime()) / 60000));

function totals(summary: DealerSummary[]) {
  const t = { dealers: summary.length, failedDealers: summary.filter((s) => s.failed).length, blank: 0, updated: 0, dryrun: 0, missing: 0, error: 0 };
  summary.forEach((s) => { t.blank += s.total; t.updated += s.counts.updated; t.dryrun += s.counts.dryrun; t.missing += s.counts.missing; t.error += s.counts.error; });
  return t;
}

// ---------------------------------------------------------------- email body (static)

function buildEmailHtml(summary: DealerSummary[], results: ResultRow[], meta: ReportMeta): string {
  const t = totals(summary);
  const font = "font-family:'Segoe UI',Roboto,Arial,sans-serif;";
  const tile = (value: number | string, label: string, accent: string) => `
    <td style="padding:6px;" width="25%">
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #e4e3de;border-radius:10px;border-top:4px solid ${accent};">
        <tr><td style="${font}padding:12px 14px 2px;font-size:26px;font-weight:700;color:#0b0b0b;">${esc(value)}</td></tr>
        <tr><td style="${font}padding:0 14px 12px;font-size:12px;color:#52514e;">${esc(label)}</td></tr>
      </table>
    </td>`;

  const maxTotal = Math.max(1, ...summary.map((s) => s.total));
  const dealerRows = summary.map((s) => {
    const segs = OUTCOMES.filter((o) => s.counts[o.key] > 0).map((o) =>
      `<td style="background:${o.color};height:14px;border-right:2px solid #ffffff;" width="${(s.counts[o.key] / maxTotal) * 100}%" title="${esc(o.label)}: ${s.counts[o.key]}"></td>`).join('');
    const rest = s.total < maxTotal ? `<td width="${((maxTotal - s.total) / maxTotal) * 100}%"></td>` : '';
    const bar = s.failed
      ? `<span style="color:#d03b3b;font-weight:600;">✖ Dealer run failed</span>`
      : s.total === 0
        ? `<span style="color:#52514e;">No blank PDI rows</span>`
        : `<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>${segs}${rest}</tr></table>`;
    const counts = OUTCOMES.filter((o) => s.counts[o.key] > 0).map((o) =>
      `<span style="white-space:nowrap;margin-right:10px;"><span style="color:${o.color};">${o.icon}</span> ${esc(o.label)} <b>${s.counts[o.key]}</b></span>`).join('');
    return `
      <tr>
        <td style="${font}padding:10px 12px;border-bottom:1px solid #eeede8;font-weight:600;font-size:14px;color:#0b0b0b;" width="90">${esc(s.dealer)}</td>
        <td style="${font}padding:10px 12px;border-bottom:1px solid #eeede8;font-size:13px;color:#0b0b0b;" width="60" align="right">${s.total}</td>
        <td style="${font}padding:10px 12px;border-bottom:1px solid #eeede8;font-size:12px;color:#52514e;">${bar}<div style="margin-top:6px;">${counts}</div></td>
      </tr>`;
  }).join('');

  const problems = results.filter((r) => ['missing', 'error'].includes(outcomeOf(r)));
  const failedDealers = summary.filter((s) => s.failed);
  const problemRows = [
    ...failedDealers.map((s) => `<tr><td style="${font}padding:6px 10px;border-bottom:1px solid #eeede8;font-size:12px;">${esc(s.dealer)}</td><td style="${font}padding:6px 10px;border-bottom:1px solid #eeede8;font-size:12px;">—</td><td style="${font}padding:6px 10px;border-bottom:1px solid #eeede8;font-size:12px;color:#d03b3b;">✖ Dealer failed: ${esc(s.failed)}</td></tr>`),
    ...problems.slice(0, 50).map((r) => {
      const o = OUTCOMES.find((x) => x.key === outcomeOf(r))!;
      return `<tr><td style="${font}padding:6px 10px;border-bottom:1px solid #eeede8;font-size:12px;">${esc(r.Dealer)}</td><td style="${font}padding:6px 10px;border-bottom:1px solid #eeede8;font-size:12px;font-family:Consolas,monospace;">${esc(r.ChassisNo)}</td><td style="${font}padding:6px 10px;border-bottom:1px solid #eeede8;font-size:12px;"><span style="color:${o.color};">${o.icon}</span> ${esc(r.Status)}</td></tr>`;
    }),
  ].join('');
  const problemsHtml = problemRows
    ? `<h3 style="${font}font-size:15px;color:#0b0b0b;margin:24px 0 8px;">Needs attention</h3>
       <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #e4e3de;border-radius:10px;border-collapse:separate;">
         <tr><th align="left" style="${font}padding:8px 10px;font-size:11px;color:#52514e;text-transform:uppercase;">Dealer</th><th align="left" style="${font}padding:8px 10px;font-size:11px;color:#52514e;text-transform:uppercase;">Chassis</th><th align="left" style="${font}padding:8px 10px;font-size:11px;color:#52514e;text-transform:uppercase;">Problem</th></tr>
         ${problemRows}
       </table>
       ${problems.length > 50 ? `<p style="${font}font-size:12px;color:#52514e;">…and ${problems.length - 50} more — see the attached dashboard / Excel.</p>` : ''}`
    : `<p style="${font}font-size:14px;color:#0ca30c;margin:24px 0 0;">✔ No errors — every blank chassis was ${meta.dryRun ? 'filled' : 'updated'}.</p>`;

  const doneLabel = meta.dryRun ? 'Filled (dry run)' : 'Updated';
  const doneValue = meta.dryRun ? t.dryrun : t.updated;
  return `<!doctype html><html><body style="margin:0;background:#f4f3ef;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f3ef;"><tr><td align="center" style="padding:20px 10px;">
  <table width="680" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;">
    <tr><td style="${font}background:#16325c;color:#ffffff;border-radius:12px 12px 0 0;padding:20px 22px;">
      <div style="font-size:12px;letter-spacing:1px;text-transform:uppercase;opacity:.75;">River DMS · ${esc(titleOf(meta))}</div>
      <div style="font-size:22px;font-weight:700;margin-top:4px;">${doneValue} of ${t.blank} blank chassis ${meta.dryRun ? 'filled' : 'updated'}</div>
      <div style="font-size:13px;opacity:.85;margin-top:6px;">${meta.dryRun ? '<b style="background:#2a78d6;padding:2px 8px;border-radius:10px;">DRY RUN — nothing saved</b> · ' : ''}${esc(fmtTime(meta.finishedAt))} · ${minutes(meta)} min · ${t.dealers} dealer(s)</div>
    </td></tr>
    <tr><td style="background:#fcfcfb;padding:12px 16px 22px;border-radius:0 0 12px 12px;">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>
        ${tile(t.blank, 'Blank chassis found', '#16325c')}
        ${tile(doneValue, doneLabel, meta.dryRun ? '#2a78d6' : '#0ca30c')}
        ${tile(t.missing, 'Not in VIN sheet', '#fab219')}
        ${tile(t.error + t.failedDealers, 'Errors', '#d03b3b')}
      </tr></table>
      <h3 style="${font}font-size:15px;color:#0b0b0b;margin:20px 0 8px;">By dealer</h3>
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #e4e3de;border-radius:10px;border-collapse:separate;">
        <tr><th align="left" style="${font}padding:8px 12px;font-size:11px;color:#52514e;text-transform:uppercase;">Dealer</th><th align="right" style="${font}padding:8px 12px;font-size:11px;color:#52514e;text-transform:uppercase;">Blank</th><th align="left" style="${font}padding:8px 12px;font-size:11px;color:#52514e;text-transform:uppercase;">Outcome</th></tr>
        ${dealerRows}
      </table>
      ${problemsHtml}
      <p style="${font}font-size:12px;color:#52514e;margin:22px 0 0;line-height:1.5;">
        📊 <b>Open the attached PDI_Dashboard.html</b> in a browser for the interactive view (filter by dealer or outcome, search chassis numbers, sort).<br/>
        The full list is also in the attached Excel. Run from ${esc(process.env.GITHUB_ACTIONS ? 'GitHub Actions' : `local PC (${os.hostname()})`)} · VIN sheet: ${esc(path.basename(meta.vinFile))}
      </p>
    </td></tr>
  </table></td></tr></table></body></html>`;
}

// ---------------------------------------------------------------- interactive dashboard (attachment)

function buildDashboardHtml(summary: DealerSummary[], results: ResultRow[], meta: ReportMeta): string {
  const data = {
    meta: { dryRun: meta.dryRun, finished: fmtTime(meta.finishedAt), minutes: minutes(meta), vinFile: path.basename(meta.vinFile) },
    outcomes: OUTCOMES,
    dealers: summary.map((s) => ({ dealer: s.dealer, failed: s.failed || '', total: s.total, counts: s.counts })),
    rows: results.map((r) => ({ dealer: r.Dealer, chassis: r.ChassisNo, blank: r.BlankFields, engine: r.EngineNo, battery: r.BatteryDetails, charger: r.VehicleChargerNo, status: r.Status, outcome: outcomeOf(r) })),
  };
  // Keep "</script>" and friends from breaking out of the inline JSON.
  const json = JSON.stringify(data).replace(/</g, '\\u003c');

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PDI Update Dashboard</title>
<style>
:root{color-scheme:light;--bg:#f4f3ef;--surface:#fcfcfb;--card:#ffffff;--line:#e4e3de;--grid:#eeede8;--ink:#0b0b0b;--ink2:#52514e;--ink3:#8a8984;--brand:#16325c;--chip:#efeee9;
--good:#0ca30c;--info:#2a78d6;--warn:#fab219;--crit:#d03b3b}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--bg:#111110;--surface:#1a1a19;--card:#222220;--line:#34332f;--grid:#2b2a27;--ink:#ffffff;--ink2:#c3c2b7;--ink3:#8f8e86;--brand:#3987e5;--chip:#2b2a27;--info:#3987e5}}
:root[data-theme="dark"]{color-scheme:dark;--bg:#111110;--surface:#1a1a19;--card:#222220;--line:#34332f;--grid:#2b2a27;--ink:#ffffff;--ink2:#c3c2b7;--ink3:#8f8e86;--brand:#3987e5;--chip:#2b2a27;--info:#3987e5}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 'Segoe UI',Roboto,Arial,sans-serif}
.wrap{max-width:1100px;margin:0 auto;padding:20px 16px 48px}
header{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;justify-content:space-between;margin-bottom:18px}
.eyebrow{font-size:12px;letter-spacing:1px;text-transform:uppercase;color:var(--ink2)}
h1{margin:2px 0 4px;font-size:24px}
.sub{color:var(--ink2);font-size:13px}
.badge{display:inline-block;background:var(--info);color:#fff;border-radius:10px;padding:1px 9px;font-size:12px;font-weight:600;margin-right:6px}
button.theme{background:var(--card);border:1px solid var(--line);color:var(--ink);border-radius:8px;padding:6px 10px;cursor:pointer}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:18px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;border-top:4px solid var(--accent)}
.tile .v{font-size:30px;font-weight:700}
.tile .l{color:var(--ink2);font-size:12px}
.tile .p{color:var(--ink3);font-size:12px;margin-top:2px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:18px}
.card h2{font-size:15px;margin:0 0 4px}
.card .hint{color:var(--ink3);font-size:12px;margin-bottom:12px}
.legend{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;color:var(--ink2);margin-bottom:10px}
.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:5px;vertical-align:-1px}
.drow{display:grid;grid-template-columns:90px 1fr 80px;gap:12px;align-items:center;padding:7px 6px;border-radius:8px;cursor:pointer}
.drow:hover,.drow.sel{background:var(--chip)}
.drow .name{font-weight:600}
.drow .num{text-align:right;color:var(--ink2);font-variant-numeric:tabular-nums}
.bar{display:flex;gap:2px;height:18px;align-items:stretch}
.bar span{border-radius:4px;min-width:3px;transition:opacity .15s}
.bar span:hover{opacity:.8}
.failed{color:var(--crit);font-weight:600;font-size:13px}
.none{color:var(--ink3);font-size:13px}
.filters{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
.chip{border:1px solid var(--line);background:var(--chip);color:var(--ink);border-radius:999px;padding:4px 11px;font-size:12px;cursor:pointer}
.chip[aria-pressed="true"]{background:var(--ink);color:var(--surface);border-color:var(--ink)}
input[type=search]{flex:1;min-width:180px;background:var(--surface);color:var(--ink);border:1px solid var(--line);border-radius:8px;padding:7px 10px;font:inherit}
.tablewrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:var(--ink2);padding:8px;border-bottom:1px solid var(--line);cursor:pointer;white-space:nowrap;user-select:none}
th:hover{color:var(--ink)}
td{padding:8px;border-bottom:1px solid var(--grid);vertical-align:top}
td.mono{font-family:Consolas,'Cascadia Mono',monospace;font-size:12px}
.st{white-space:nowrap}
.st b{font-weight:600}
.count{color:var(--ink3);font-size:12px;margin-left:auto}
#tip{position:fixed;pointer-events:none;background:var(--ink);color:var(--surface);padding:6px 9px;border-radius:6px;font-size:12px;opacity:0;transition:opacity .1s;z-index:10}
.errbox{border-left:4px solid var(--crit)}
.errbox li{margin:4px 0}
@media (max-width:560px){.drow{grid-template-columns:70px 1fr 50px}h1{font-size:20px}}
</style></head>
<body><div class="wrap">
<header>
  <div><div class="eyebrow">River DMS · ${esc(titleOf(meta))}</div><h1 id="headline"></h1><div class="sub" id="sub"></div></div>
  <button class="theme" id="theme" type="button" aria-label="Toggle dark mode">◐ Theme</button>
</header>
<section class="tiles" id="tiles"></section>
<section class="card errbox" id="dealerErrors" hidden><h2>✖ Dealers that failed</h2><ul id="dealerErrList"></ul></section>
<section class="card">
  <h2>Outcome by dealer</h2><div class="hint">Click a dealer to filter the table below. Hover a segment for its count.</div>
  <div class="legend" id="legend"></div>
  <div id="dealers"></div>
</section>
<section class="card">
  <h2>Chassis details</h2>
  <div class="filters" id="outcomeFilters"></div>
  <div class="filters"><input type="search" id="q" placeholder="Search chassis, engine, battery, charger or status…"><span class="count" id="count"></span></div>
  <div class="tablewrap"><table><thead><tr id="thead"></tr></thead><tbody id="tbody"></tbody></table></div>
</section>
<div class="sub">Generated ${esc(fmtTime(meta.finishedAt))} · VIN sheet: ${esc(path.basename(meta.vinFile))}</div>
</div>
<div id="tip" role="tooltip"></div>
<script>
const D = ${json};
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const O = Object.fromEntries(D.outcomes.map((o) => [o.key, o]));
const cssVar = { updated: 'var(--good)', dryrun: 'var(--info)', missing: 'var(--warn)', error: 'var(--crit)' };
const state = { dealer: null, outcomes: new Set(), q: '', sort: { key: 'dealer', dir: 1 } };

// theme toggle (remembered per browser when storage is available)
try { const t = localStorage.getItem('pdi-theme'); if (t) document.documentElement.dataset.theme = t; } catch {}
$('theme').onclick = () => {
  const dark = getComputedStyle(document.documentElement).colorScheme === 'dark';
  document.documentElement.dataset.theme = dark ? 'light' : 'dark';
  try { localStorage.setItem('pdi-theme', document.documentElement.dataset.theme); } catch {}
};

const sum = (k) => D.dealers.reduce((a, d) => a + d.counts[k], 0);
const total = D.dealers.reduce((a, d) => a + d.total, 0);
const failedDealers = D.dealers.filter((d) => d.failed);
const done = D.meta.dryRun ? sum('dryrun') : sum('updated');
const pct = (n) => total ? Math.round((n / total) * 100) + '%' : '—';

$('headline').textContent = done + ' of ' + total + ' blank chassis ' + (D.meta.dryRun ? 'filled' : 'updated');
$('sub').innerHTML = (D.meta.dryRun ? '<span class="badge">DRY RUN — nothing saved</span>' : '') +
  esc(D.meta.finished) + ' · ' + D.meta.minutes + ' min · ' + D.dealers.length + ' dealer(s)';

const tiles = [
  { v: total, l: 'Blank chassis found', a: 'var(--brand)', p: D.dealers.length + ' dealer(s) checked' },
  { v: done, l: D.meta.dryRun ? 'Filled (dry run)' : 'Updated', a: D.meta.dryRun ? 'var(--info)' : 'var(--good)', p: pct(done) + ' of blank' },
  { v: sum('missing'), l: 'Not in VIN sheet', a: 'var(--warn)', p: pct(sum('missing')) + ' of blank' },
  { v: sum('error') + failedDealers.length, l: 'Errors', a: 'var(--crit)', p: sum('error') + ' chassis · ' + failedDealers.length + ' dealer(s)' },
];
$('tiles').innerHTML = tiles.map((t) => '<div class="tile" style="--accent:' + t.a + '"><div class="v">' + t.v + '</div><div class="l">' + esc(t.l) + '</div><div class="p">' + esc(t.p) + '</div></div>').join('');

if (failedDealers.length) {
  $('dealerErrors').hidden = false;
  $('dealerErrList').innerHTML = failedDealers.map((d) => '<li><b>' + esc(d.dealer) + '</b> — ' + esc(d.failed) + '</li>').join('');
}

const shown = D.outcomes.filter((o) => D.dealers.some((d) => d.counts[o.key] > 0));
$('legend').innerHTML = shown.map((o) => '<span><i style="background:' + cssVar[o.key] + '"></i>' + esc(o.icon + ' ' + o.label) + '</span>').join('');

const tip = $('tip');
const showTip = (e, text) => { tip.textContent = text; tip.style.opacity = 1; tip.style.left = (e.clientX + 12) + 'px'; tip.style.top = (e.clientY + 12) + 'px'; };
const hideTip = () => { tip.style.opacity = 0; };

function renderDealers() {
  const max = Math.max(1, ...D.dealers.map((d) => d.total));
  $('dealers').innerHTML = D.dealers.map((d) => {
    let bar;
    if (d.failed) bar = '<div class="failed">✖ Run failed</div>';
    else if (!d.total) bar = '<div class="none">No blank PDI rows</div>';
    else bar = '<div class="bar" style="width:' + (d.total / max) * 100 + '%">' + D.outcomes.filter((o) => d.counts[o.key])
      .map((o) => '<span data-tip="' + esc(d.dealer + ' · ' + o.label + ': ' + d.counts[o.key]) + '" style="flex:' + d.counts[o.key] + ';background:' + cssVar[o.key] + '"></span>').join('') + '</div>';
    return '<div class="drow' + (state.dealer === d.dealer ? ' sel' : '') + '" data-dealer="' + esc(d.dealer) + '" role="button" tabindex="0" aria-pressed="' + (state.dealer === d.dealer) + '">' +
      '<div class="name">' + esc(d.dealer) + '</div><div>' + bar + '</div><div class="num">' + d.total + ' blank</div></div>';
  }).join('');
  document.querySelectorAll('.drow').forEach((el) => {
    const pick = () => { state.dealer = state.dealer === el.dataset.dealer ? null : el.dataset.dealer; renderDealers(); renderTable(); };
    el.onclick = pick;
    el.onkeydown = (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } };
  });
  document.querySelectorAll('.bar span').forEach((s) => { s.onmousemove = (e) => showTip(e, s.dataset.tip); s.onmouseleave = hideTip; });
}

function renderOutcomeFilters() {
  $('outcomeFilters').innerHTML = '<button class="chip" data-o="" aria-pressed="' + (state.outcomes.size === 0) + '">All</button>' +
    shown.map((o) => '<button class="chip" data-o="' + o.key + '" aria-pressed="' + state.outcomes.has(o.key) + '">' + esc(o.icon + ' ' + o.label) + '</button>').join('') +
    (state.dealer ? '<button class="chip" data-clear-dealer aria-pressed="true">Dealer ' + esc(state.dealer) + ' ✕</button>' : '');
  document.querySelectorAll('#outcomeFilters .chip').forEach((b) => b.onclick = () => {
    if (b.hasAttribute('data-clear-dealer')) { state.dealer = null; renderDealers(); }
    else if (!b.dataset.o) state.outcomes.clear();
    else state.outcomes.has(b.dataset.o) ? state.outcomes.delete(b.dataset.o) : state.outcomes.add(b.dataset.o);
    renderTable();
  });
}

const COLS = [
  { key: 'dealer', label: 'Dealer' }, { key: 'chassis', label: 'Chassis No', mono: true }, { key: 'blank', label: 'Was blank' },
  { key: 'engine', label: 'Engine No', mono: true }, { key: 'battery', label: 'Battery', mono: true }, { key: 'charger', label: 'Charger', mono: true },
  { key: 'status', label: 'Result' },
];
function renderTable() {
  renderOutcomeFilters();
  const q = state.q.toLowerCase();
  const rows = D.rows.filter((r) => (!state.dealer || r.dealer === state.dealer) && (!state.outcomes.size || state.outcomes.has(r.outcome)) &&
    (!q || Object.values(r).some((v) => String(v).toLowerCase().includes(q))))
    .sort((a, b) => String(a[state.sort.key]).localeCompare(String(b[state.sort.key]), undefined, { numeric: true }) * state.sort.dir);
  $('thead').innerHTML = COLS.map((c) => '<th data-k="' + c.key + '">' + c.label + (state.sort.key === c.key ? (state.sort.dir > 0 ? ' ▲' : ' ▼') : '') + '</th>').join('');
  document.querySelectorAll('#thead th').forEach((th) => th.onclick = () => {
    state.sort = { key: th.dataset.k, dir: state.sort.key === th.dataset.k ? -state.sort.dir : 1 }; renderTable();
  });
  $('tbody').innerHTML = rows.length ? rows.map((r) => '<tr>' + COLS.map((c) => {
    if (c.key === 'status') return '<td class="st"><span style="color:' + cssVar[r.outcome] + '">' + O[r.outcome].icon + '</span> <b>' + esc(O[r.outcome].label) + '</b>' +
      (r.outcome === 'error' || /\\(/.test(r.status) ? '<div style="color:var(--ink2);font-size:12px;white-space:normal">' + esc(r.status) + '</div>' : '') + '</td>';
    return '<td' + (c.mono ? ' class="mono"' : '') + '>' + esc(r[c.key] || '—') + '</td>';
  }).join('') + '</tr>').join('') : '<tr><td colspan="' + COLS.length + '" style="color:var(--ink3);text-align:center;padding:24px">No chassis match these filters.</td></tr>';
  $('count').textContent = rows.length + ' of ' + D.rows.length + ' chassis';
}

$('q').oninput = (e) => { state.q = e.target.value; renderTable(); };
renderDealers();
renderTable();
</script></body></html>`;
}

// ---------------------------------------------------------------- send

// A run passes only when every blank chassis was updated (or filled, in a dry run) and no dealer
// failed. Returns one line per problem; empty = passed. Used for both the email and the test result.
export function problemsOf(results: ResultRow[], failed: FailedDealer[], dryRun: boolean): string[] {
  const done: Outcome = dryRun ? 'dryrun' : 'updated';
  return [
    ...failed.map((f) => `Dealer ${f.dealer} failed: ${f.error}`),
    ...results.filter((r) => outcomeOf(r) !== done).map((r) => `${r.Dealer} ${r.ChassisNo}: ${r.Status}`),
  ];
}

// Uses the same GMAIL_USER / GMAIL_APP_PASSWORD / GMAIL_TO / GMAIL_CC settings as the other reports.
// Set PDI_SEND_EMAIL=0 to skip. Never throws: a mail problem must not fail the update run.
export async function sendPdiReport(results: ResultRow[], failed: FailedDealer[], meta: ReportMeta, excelFile: string): Promise<void> {
  const summary = summarize(results, meta, failed);
  const dashboard = buildDashboardHtml(summary, results, meta);
  const dashboardFile = excelFile.replace(/\.xlsx$/i, '.html');
  try {
    fs.writeFileSync(dashboardFile, dashboard);
    console.log(`[STEP] Dashboard saved to ${dashboardFile}`);
  } catch (err) {
    console.log(`[WARN] Could not save the dashboard: ${String((err as Error).message).slice(0, 120)}`);
  }
  if (process.env.PDI_SEND_EMAIL === '0') {
    console.log('[STEP] PDI_SEND_EMAIL=0 — report email skipped.');
    return;
  }
  // A failed run's report isn't emailed to anyone — only a clean run goes out.
  const problems = problemsOf(results, failed, meta.dryRun);
  if (problems.length) {
    console.log(`[STEP] Report email NOT sent — the run did not fully pass (${problems.length} problem(s)). Dashboard and Excel are saved locally.`);
    return;
  }
  const user = process.env.GMAIL_USER;
  const pass = process.env.GMAIL_APP_PASSWORD;
  const to = process.env.GMAIL_TO || user;
  if (!user || !pass || !to) {
    console.log('[STEP] Report email skipped — GMAIL_USER / GMAIL_APP_PASSWORD not set in .env.');
    return;
  }

  const t = totals(summary);
  const done = meta.dryRun ? t.dryrun : t.updated;
  const subject = `${meta.title ? meta.title + ' — ' : 'PDI details '}${meta.dryRun ? '[DRY RUN] filled' : 'updated'}: ${done}/${t.blank} across ${t.dealers} dealer(s)` +
    ` — ${meta.finishedAt.toISOString().slice(0, 10)}`;
  // PDI_CC (the PDI audience) only on real runs; a dry run goes to GMAIL_TO alone.
  const cc = meta.dryRun ? '' : [process.env.PDI_CC, process.env.GMAIL_CC].filter(Boolean).join(', ');
  try {
    await nodemailer.createTransport({ service: 'gmail', auth: { user, pass } }).sendMail({
      from: user,
      to,
      ...(cc ? { cc } : {}),
      subject,
      text: summary.map((s) => `${s.dealer}: ${s.failed ? 'FAILED — ' + s.failed : `${s.counts.updated + s.counts.dryrun}/${s.total} done, ${s.counts.missing} not in VIN sheet, ${s.counts.error} error(s)`}`).join('\n'),
      html: buildEmailHtml(summary, results, meta),
      attachments: [
        { filename: 'PDI_Dashboard.html', content: dashboard, contentType: 'text/html' },
        { filename: path.basename(excelFile), path: excelFile },
      ],
    });
    console.log(`[STEP] Report emailed to ${to}${cc ? ` (cc ${cc})` : ''}.`);
  } catch (err) {
    console.log(`[WARN] Could not send the report email: ${String((err as Error).message).slice(0, 150)}`);
  }
}

// For a local preview without sending anything.
export function renderPdiReport(results: ResultRow[], failed: FailedDealer[], meta: ReportMeta) {
  const summary = summarize(results, meta, failed);
  return { email: buildEmailHtml(summary, results, meta), dashboard: buildDashboardHtml(summary, results, meta) };
}
