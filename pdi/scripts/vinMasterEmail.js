// VIN master result email (pdi/Vin_Master_update.spec.ts): from the progress file builds
//   Output/VIN Master/VIN_Master_Dashboard_<date-time>.html   interactive: outcome chart (click a bar to filter),
//                                                              search, sort, download the rows shown as CSV
//   Output/VIN Master/VIN_Master_Not_Saved_<date-time>.xlsx   one sheet per reason a VIN was not saved
// and emails both, with a summary dashboard in the body (email apps don't run scripts, so the body is static).
//
//   node "pdi/scripts/vinMasterEmail.js"                  build and send to GMAIL_TO (or GMAIL_USER)
//   node "pdi/scripts/vinMasterEmail.js" --to a@b.com     send to that address instead
//   node "pdi/scripts/vinMasterEmail.js" --no-send        build the files only
// Uses GMAIL_USER / GMAIL_APP_PASSWORD from .env, as the other reports do.
const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');

const ROOT = path.join(__dirname, '..', '..');
require('dotenv').config({ path: path.join(ROOT, '.env') });
const MASTER_DIR = path.join(ROOT, 'Output', 'VIN Master');
const PROGRESS_FILE = path.join(MASTER_DIR, 'VIN_Master_progress.jsonl');
const VIN_FILE = process.env.PDI_VIN_FILE || path.join(ROOT, 'resources', 'VIN_Details.csv');
const arg = (name) => { const i = process.argv.indexOf(name); return i > 0 ? process.argv[i + 1] : undefined; };

// Same as outcomeOf in Vin_Master_update.spec.ts, plus the order, colour and meaning used here.
function outcomeOf(status) {
  if (/^Already correct/.test(status)) return 'Already correct';
  if (/^Updated/.test(status)) return 'Changed';
  if (/^Dry run|\(dry run\)/.test(status)) return 'Dry run';
  if (/; the rest saved/.test(status)) return 'Corrected, sheet value missing';
  if (/^VIN sheet has no usable/.test(status)) return 'Sheet has no usable value';
  if (/Show did not load/.test(status)) return 'Not found in DMS';
  if (/^Save clicked but values not kept/.test(status)) return 'Refused by DMS';
  return 'Error';
}
const OUTCOMES = [
  { name: 'Already correct', color: '#1baf7a', saved: true, meaning: 'DMS already matched the VIN sheet, nothing to save' },
  { name: 'Changed', color: '#2a78d6', saved: true, meaning: 'Corrected from the VIN sheet and saved' },
  { name: 'Corrected, sheet value missing', color: '#6250d6', saved: true, meaning: 'Some fields corrected; the others have a note in the VIN sheet instead of a serial' },
  { name: 'Sheet has no usable value', color: '#eda100', saved: false, meaning: 'The VIN sheet has a note (CONTACTFACTORYTEAM, CHARGER NOT ASSIGNED YET...), nothing to save' },
  { name: 'Refused by DMS', color: '#d03b3b', saved: false, meaning: 'DMS refused the Save: the number is already on another vehicle ("must be unique")' },
  { name: 'Not found in DMS', color: '#eb6834', saved: false, meaning: 'DMS did not open this VIN: probably not in DMS yet' },
  { name: 'Error', color: '#898781', saved: false, meaning: 'The page failed for this VIN' },
];
// A VIN whose old charger/battery got a trailing "." to free the number and that still isn't corrected.
const dotLeft = (r) => !['Changed', 'Already correct'].includes(outcomeOf(r.Status || '')) &&
  (/freed: set to/.test(r.Status || '') || /DMS has "[^"]*\."/.test(r.Status || '') ||
   [r.BeforeBatteryDetails, r.BeforeVehicleChargerNo].some((v) => String(v || '').endsWith('.')));

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const fmt = (n) => n.toLocaleString('en-IN');

function load() {
  const latest = new Map();
  for (const line of fs.readFileSync(PROGRESS_FILE, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    try { const r = JSON.parse(line); latest.set(r.ChassisNo, r); } catch { /* a line cut off by a stop */ }
  }
  const sheetVins = fs.readFileSync(VIN_FILE, 'utf8').split(/\r?\n/).slice(1).filter((l) => l.trim()).length;
  const rows = [...latest.values()].map((r) => ({
    vin: r.ChassisNo, outcome: outcomeOf(r.Status || ''), dot: dotLeft(r), changed: r.Changed || '',
    engineBefore: r.BeforeEngineNo ?? '', engineSheet: r.SheetMotor ?? '',
    batteryBefore: r.BeforeBatteryDetails ?? '', batterySheet: r.SheetBattery ?? '',
    chargerBefore: r.BeforeVehicleChargerNo ?? '', chargerSheet: r.SheetCharger ?? '',
    message: r.Status || '', at: r.At || '',
  }));
  const counts = Object.fromEntries(OUTCOMES.map((o) => [o.name, rows.filter((r) => r.outcome === o.name).length]));
  return { rows, counts, sheetVins, dots: rows.filter((r) => r.dot).length, when: new Date() };
}

// ---- Excel: VINs not saved, one sheet per reason ----
function writeExcel(d, file) {
  const groups = [
    ['Refused by DMS', 'Refused by DMS'], ['Not found in DMS', 'Not found in DMS'],
    ['No usable sheet value', 'Sheet has no usable value'], ['Sheet value missing', 'Corrected, sheet value missing'], ['Error', 'Error'],
  ];
  const row = (r) => ({
    'Chassis No': r.vin, Reason: r.outcome, 'Engine No in DMS': r.engineBefore, 'Engine No in sheet': r.engineSheet,
    'Battery in DMS': r.batteryBefore, 'Battery in sheet': r.batterySheet, 'Charger in DMS': r.chargerBefore, 'Charger in sheet': r.chargerSheet,
    'DMS message': r.message, 'Checked at': r.at,
  });
  const wb = XLSX.utils.book_new();
  const summary = [['VIN master: VINs not saved', '', ''], [`${d.when.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })} · ${fmt(d.rows.length)} VINs checked`, '', ''], [], ['Sheet', 'VINs', 'What it means']];
  for (const [sheet, outcome] of groups) summary.push([sheet, d.counts[outcome], OUTCOMES.find((o) => o.name === outcome).meaning]);
  summary.push(['Dot left in DMS', d.dots, 'Old charger/battery given a trailing "." to free the number; this VIN still could not get its own']);
  const ws0 = XLSX.utils.aoa_to_sheet(summary);
  ws0['!cols'] = [{ wch: 24 }, { wch: 8 }, { wch: 110 }];
  XLSX.utils.book_append_sheet(wb, ws0, 'Summary');
  const widths = [{ wch: 20 }, { wch: 28 }, { wch: 16 }, { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 16 }, { wch: 16 }, { wch: 100 }, { wch: 22 }];
  const add = (name, list) => {
    const ws = XLSX.utils.json_to_sheet(list.map(row));
    ws['!cols'] = widths;
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    XLSX.utils.book_append_sheet(wb, ws, name);
  };
  for (const [sheet, outcome] of groups) add(sheet, d.rows.filter((r) => r.outcome === outcome));
  add('Dot left in DMS', d.rows.filter((r) => r.dot));
  XLSX.writeFile(wb, file);
}

// ---- Interactive dashboard (opened in a browser) ----
function dashboardHtml(d) {
  const data = JSON.stringify(d.rows).replace(/</g, '\\u003c');
  const outcomes = JSON.stringify(OUTCOMES);
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>VIN Master Dashboard</title>
<style>
:root { --page:#f9f9f7; --surface:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e; --muted:#898781; --border:#e1e0d9; --grid:#e1e0d9; --accent:#2a78d6; --sel:#eef4fc; }
@media (prefers-color-scheme: dark) { :root { --page:#0d0d0d; --surface:#1a1a19; --ink:#f0efec; --ink2:#c3c2b7; --border:#2c2c2a; --grid:#2c2c2a; --accent:#3987e5; --sel:#1c2a3d; } }
* { box-sizing:border-box; }
body { margin:0; background:var(--page); color:var(--ink); font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif; }
main { max-width:1200px; margin:0 auto; padding:24px 16px 48px; }
h1 { font-size:21px; margin:0; } h2 { font-size:15px; margin:26px 0 10px; }
.sub { color:var(--muted); }
.tiles { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:10px; margin-top:16px; }
.tile { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:12px 14px; cursor:pointer; text-align:left; font:inherit; color:inherit; }
.tile:hover, .tile.on { border-color:var(--accent); background:var(--sel); }
.tile .k { color:var(--ink2); font-size:12px; display:flex; align-items:center; gap:6px; } .tile .v { font-size:24px; font-weight:650; font-variant-numeric:tabular-nums; }
.sw { width:10px; height:10px; border-radius:3px; display:inline-block; }
.card { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:16px; }
.bars { display:grid; grid-template-columns:minmax(150px,240px) 1fr 70px; gap:6px 12px; align-items:center; }
.bars .lab { font-size:13px; color:var(--ink2); cursor:pointer; } .bars .num { text-align:right; font-variant-numeric:tabular-nums; }
.track { height:18px; cursor:pointer; } .bar { height:100%; border-radius:0 4px 4px 0; min-width:2px; }
.row.on .lab { color:var(--ink); font-weight:600; } .row:hover .bar { opacity:.8; }
.controls { display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-bottom:10px; }
input[type=search], select, button.btn { font:inherit; padding:7px 10px; border:1px solid var(--border); border-radius:8px; background:var(--page); color:var(--ink); }
input[type=search] { flex:1; min-width:220px; } button.btn { cursor:pointer; } button.btn:hover { border-color:var(--accent); }
.scroll { overflow:auto; max-height:70vh; border:1px solid var(--border); border-radius:8px; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th, td { text-align:left; padding:6px 8px; border-bottom:1px solid var(--border); vertical-align:top; white-space:nowrap; }
th { position:sticky; top:0; background:var(--surface); color:var(--ink2); cursor:pointer; user-select:none; }
th:hover { color:var(--ink); } td.msg { white-space:normal; min-width:320px; color:var(--ink2); }
.mono { font-family:ui-monospace,Consolas,monospace; } .diff { font-weight:650; } s { color:var(--muted); }
.pill { display:inline-flex; align-items:center; gap:6px; } .dot { color:#d03b3b; font-weight:600; }
.pager { display:flex; justify-content:space-between; align-items:center; margin-top:8px; color:var(--muted); }
</style></head><body><main>
<h1>VIN master update</h1>
<div class="sub">${esc(d.when.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }))} · ${fmt(d.rows.length)} of ${fmt(d.sheetVins)} VINs in the VIN sheet checked against DMS</div>

<div class="tiles" id="tiles"></div>

<h2>Results by outcome <span class="sub">(click a bar or tile to filter the table)</span></h2>
<div class="card"><div class="bars" id="bars"></div></div>

<h2>VINs</h2>
<div class="card">
  <div class="controls">
    <input type="search" id="q" placeholder="Search chassis, serial or message…" aria-label="Search">
    <select id="f" aria-label="Outcome"></select>
    <label class="pill"><input type="checkbox" id="dots"> Dot left in DMS only</label>
    <button class="btn" id="csv">Download CSV</button>
  </div>
  <div class="scroll"><table><thead><tr id="head"></tr></thead><tbody id="body"></tbody></table></div>
  <div class="pager"><span id="count"></span><span><button class="btn" id="prev">‹ Prev</button> <button class="btn" id="next">Next ›</button></span></div>
</div>
</main>
<script>
const ROWS = ${data};
const OUTCOMES = ${outcomes};
const COLS = [
  ['vin','Chassis No'],['outcome','Outcome'],['changed','Changed'],
  ['engineBefore','Engine No in DMS'],['engineSheet','Engine No in sheet'],
  ['batteryBefore','Battery in DMS'],['batterySheet','Battery in sheet'],
  ['chargerBefore','Charger in DMS'],['chargerSheet','Charger in sheet'],['message','DMS message'],['at','Checked at']];
const PAGE = 200;
let filter = 'Problems', sortKey = 'outcome', sortDir = 1, page = 0;
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' })[c]);
const fmt = (n) => n.toLocaleString('en-IN');
const color = (o) => (OUTCOMES.find((x) => x.name === o) || {}).color || '#898781';
const count = (o) => ROWS.filter((r) => r.outcome === o).length;
const notSaved = (r) => !(OUTCOMES.find((x) => x.name === r.outcome) || {}).saved;
const VIEWS = { 'All VINs': () => true, 'Problems': notSaved, 'Corrected': (r) => r.outcome === 'Changed' || r.outcome === 'Corrected, sheet value missing' };
OUTCOMES.forEach((o) => { if (count(o.name)) VIEWS[o.name] = (r) => r.outcome === o.name; });

function tiles() {
  const t = [['All VINs', ROWS.length, ''], ['Corrected', VIEWS.Corrected && ROWS.filter(VIEWS.Corrected).length, '#2a78d6'],
    ['Problems', ROWS.filter(notSaved).length, '#d03b3b'], ['Dot left in DMS', ROWS.filter((r) => r.dot).length, '#d03b3b']];
  $('tiles').innerHTML = t.map(([k, v, c]) => '<button class="tile' + ((k === filter && !$('dots').checked) || (k === 'Dot left in DMS' && $('dots').checked) ? ' on' : '') + '" data-k="' + esc(k) + '">' +
    '<div class="k">' + (c ? '<span class="sw" style="background:' + c + '"></span>' : '') + esc(k) + '</div><div class="v">' + fmt(v) + '</div></button>').join('');
  $('tiles').querySelectorAll('.tile').forEach((b) => b.onclick = () => {
    if (b.dataset.k === 'Dot left in DMS') { $('dots').checked = true; filter = 'All VINs'; } else { $('dots').checked = false; filter = b.dataset.k; }
    page = 0; render();
  });
}
function bars() {
  const max = Math.max(...OUTCOMES.map((o) => count(o.name)), 1);
  $('bars').innerHTML = OUTCOMES.filter((o) => count(o.name)).map((o) => {
    const n = count(o.name);
    return '<div class="row' + (filter === o.name ? ' on' : '') + '" style="display:contents" data-o="' + esc(o.name) + '" title="' + esc(o.name + ': ' + fmt(n) + ' VINs. ' + o.meaning) + '">' +
      '<div class="lab">' + esc(o.name) + '</div><div class="track"><div class="bar" style="width:' + (100 * n / max).toFixed(2) + '%;background:' + o.color + '"></div></div>' +
      '<div class="num">' + fmt(n) + '</div></div>';
  }).join('');
  $('bars').querySelectorAll('.row').forEach((r) => r.onclick = () => { filter = r.dataset.o; $('dots').checked = false; page = 0; render(); });
}
function visible() {
  const q = $('q').value.trim().toUpperCase();
  return ROWS.filter((r) => VIEWS[filter](r) && (!$('dots').checked || r.dot) &&
    (!q || COLS.some(([k]) => String(r[k]).toUpperCase().includes(q))))
    .sort((a, b) => String(a[sortKey]).localeCompare(String(b[sortKey]), 'en', { numeric: true }) * sortDir);
}
function cell(r, k) {
  if (k === 'outcome') return '<span class="pill"><span class="sw" style="background:' + color(r.outcome) + '"></span>' + esc(r.outcome) + (r.dot ? ' <span class="dot">· dot left</span>' : '') + '</span>';
  if (k === 'message') return esc(r.message);
  const pair = { engineBefore: 'engineSheet', batteryBefore: 'batterySheet', chargerBefore: 'chargerSheet' }[k];
  const differs = pair && r[k].toUpperCase() !== r[pair].toUpperCase();
  return '<span class="mono' + (differs ? ' diff' : '') + '">' + esc(r[k]) + '</span>';
}
function render() {
  $('f').innerHTML = Object.keys(VIEWS).map((k) => '<option' + (k === filter ? ' selected' : '') + '>' + esc(k) + '</option>').join('');
  tiles(); bars();
  const list = visible();
  const pages = Math.max(1, Math.ceil(list.length / PAGE));
  page = Math.min(page, pages - 1);
  $('head').innerHTML = COLS.map(([k, l]) => '<th data-k="' + k + '">' + esc(l) + (k === sortKey ? (sortDir > 0 ? ' ▲' : ' ▼') : '') + '</th>').join('');
  $('head').querySelectorAll('th').forEach((th) => th.onclick = () => { sortDir = th.dataset.k === sortKey ? -sortDir : 1; sortKey = th.dataset.k; render(); });
  $('body').innerHTML = list.slice(page * PAGE, (page + 1) * PAGE).map((r) => '<tr>' + COLS.map(([k]) => '<td' + (k === 'message' ? ' class="msg"' : '') + '>' + cell(r, k) + '</td>').join('') + '</tr>').join('')
    || '<tr><td colspan="' + COLS.length + '" class="msg">No VINs match.</td></tr>';
  $('count').textContent = fmt(list.length) + ' VIN(s)' + (pages > 1 ? ' · page ' + (page + 1) + ' of ' + pages : '');
  $('prev').disabled = page === 0; $('next').disabled = page >= pages - 1;
}
$('q').oninput = () => { page = 0; render(); };
$('f').onchange = () => { filter = $('f').value; page = 0; render(); };
$('dots').onchange = () => { page = 0; render(); };
$('prev').onclick = () => { page--; render(); };
$('next').onclick = () => { page++; render(); };
$('csv').onclick = () => {
  const q = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const csv = [COLS.map(([, l]) => q(l)).join(','), ...visible().map((r) => COLS.map(([k]) => q(r[k])).join(','))].join('\\r\\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['\\ufeff' + csv], { type: 'text/csv' }));
  a.download = 'VIN_Master_' + filter.replace(/\\W+/g, '_') + '.csv';
  a.click();
};
render();
</script></body></html>`;
}

// ---- Email body: tables and inline styles only ----
function emailHtml(d, files) {
  const td = 'padding:6px 8px;border-bottom:1px solid #e1e0d9;font-size:13px;vertical-align:top;';
  const th = 'padding:6px 8px;border-bottom:1px solid #c3c2b7;font-size:12px;color:#52514e;text-align:left;';
  const mono = 'font-family:Consolas,monospace;';
  const corrected = d.rows.filter((r) => r.outcome === 'Changed' || r.outcome === 'Corrected, sheet value missing');
  const problems = d.counts['Refused by DMS'] + d.counts['Not found in DMS'] + d.counts['Sheet has no usable value'] + d.counts.Error;
  const tile = (label, value, color) => `<td style="padding:4px;width:25%;"><div style="border:1px solid #e1e0d9;border-radius:8px;padding:10px 12px;background:#fcfcfb;">
    <div style="font-size:12px;color:#52514e;">${color ? `<span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:${color};margin-right:5px;"></span>` : ''}${label}</div>
    <div style="font-size:22px;font-weight:700;color:#0b0b0b;">${fmt(value)}</div></div></td>`;
  const max = Math.max(...OUTCOMES.map((o) => d.counts[o.name]), 1);
  const bars = OUTCOMES.filter((o) => d.counts[o.name]).map((o) => `<tr>
    <td style="${td}width:200px;color:#52514e;">${esc(o.name)}</td>
    <td style="${td}"><div style="background:${o.color};height:14px;width:${Math.max(0.5, (100 * d.counts[o.name]) / max).toFixed(1)}%;border-radius:0 3px 3px 0;"></div></td>
    <td style="${td}text-align:right;width:70px;">${fmt(d.counts[o.name])}</td></tr>`).join('');
  // Only the fields in "Changed" were saved; a field with a sheet note was left as DMS had it.
  const change = (r, label, before, after) => !r.changed.split(', ').includes(label)
    ? `<span style="${mono}color:#898781;white-space:nowrap;">${esc(before)}</span>`
    : `<span style="${mono}white-space:nowrap;"><s style="color:#898781;">${esc(before) || '(blank)'}</s> → <b>${esc(after)}</b></span>`;
  const correctedRows = corrected.map((r) => `<tr><td style="${td}${mono}">${esc(r.vin)}</td><td style="${td}">${esc(r.changed)}</td>
    <td style="${td}">${change(r, 'Engine No', r.engineBefore, r.engineSheet)}</td><td style="${td}">${change(r, 'Battery', r.batteryBefore, r.batterySheet)}</td>
    <td style="${td}">${change(r, 'Charger', r.chargerBefore, r.chargerSheet)}</td></tr>`).join('');
  const refused = d.rows.filter((r) => r.outcome === 'Refused by DMS').map((r) => `<tr><td style="${td}${mono}">${esc(r.vin)}</td>
    <td style="${td}">${r.dot ? '<b style="color:#d03b3b;">dot left</b>' : ''}</td><td style="${td}color:#52514e;">${esc(r.message)}</td></tr>`).join('');
  const section = (title, body) => `<h3 style="font-size:15px;margin:24px 0 8px;color:#0b0b0b;">${title}</h3>${body}`;
  return `<div style="font-family:Segoe UI,Arial,sans-serif;color:#0b0b0b;max-width:900px;">
  <h2 style="margin:0 0 4px;font-size:20px;">VIN master update</h2>
  <div style="color:#898781;font-size:13px;">${esc(d.when.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }))} · ${fmt(d.rows.length)} of ${fmt(d.sheetVins)} VINs checked against DMS</div>
  <table style="width:100%;border-collapse:collapse;margin-top:12px;"><tr>
    ${tile('VINs checked', d.rows.length)}${tile('Corrected', corrected.length, '#2a78d6')}${tile('Not saved', problems, '#d03b3b')}${tile('Dot left in DMS', d.dots, '#d03b3b')}
  </tr></table>
  <p style="font-size:13px;color:#52514e;">For search, filters and sorting open the attached <b>${esc(files.dashboard)}</b> in a browser. Every VIN not saved, by reason, is in <b>${esc(files.excel)}</b>.</p>
  ${section('Results by outcome', `<table style="width:100%;border-collapse:collapse;">${bars}</table>`)}
  ${section(`Corrected and saved (${fmt(corrected.length)})`, `<table style="width:100%;border-collapse:collapse;"><tr><th style="${th}">Chassis No</th><th style="${th}">Changed</th><th style="${th}">Engine No</th><th style="${th}">Battery</th><th style="${th}">Charger</th></tr>${correctedRows}</table>`)}
  ${section(`Refused by DMS (${fmt(d.counts['Refused by DMS'])})`, `<p style="font-size:13px;color:#52514e;margin:0 0 6px;">The number is already on another vehicle. "dot left": the old number now ends in "." in DMS.</p>
    <table style="width:100%;border-collapse:collapse;"><tr><th style="${th}">Chassis No</th><th style="${th}"></th><th style="${th}">DMS message</th></tr>${refused}</table>`)}
  ${section(`Not found in DMS (${fmt(d.counts['Not found in DMS'])})`, `<p style="font-size:13px;color:#52514e;margin:0;">DMS did not open these VINs, probably not in DMS yet. Full list in the attachments.</p>`)}
  ${section(`Sheet has no usable value (${fmt(d.counts['Sheet has no usable value'])})`, `<p style="font-size:13px;color:#52514e;margin:0;">The VIN sheet has a note instead of a serial; nothing to save. Full list in the attachments.</p>`)}
</div>`;
}

async function main() {
  const d = load();
  fs.mkdirSync(MASTER_DIR, { recursive: true });
  // Local date and time in the names: a copy left open in Excel can't block the next one.
  const p2 = (n) => String(n).padStart(2, '0');
  const w = d.when;
  const stamp = `${w.getFullYear()}-${p2(w.getMonth() + 1)}-${p2(w.getDate())}`;
  const fileStamp = `${stamp}-${p2(w.getHours())}${p2(w.getMinutes())}`;
  const dashboard = path.join(MASTER_DIR, `VIN_Master_Dashboard_${fileStamp}.html`);
  const excel = path.join(MASTER_DIR, `VIN_Master_Not_Saved_${fileStamp}.xlsx`);
  fs.writeFileSync(dashboard, dashboardHtml(d));
  writeExcel(d, excel);
  const html = emailHtml(d, { dashboard: path.basename(dashboard), excel: path.basename(excel) });
  fs.writeFileSync(path.join(MASTER_DIR, 'VIN_Master_Email_preview.html'), html);
  console.log(`[STEP] ${dashboard}\n[STEP] ${excel}`);
  if (process.argv.includes('--no-send')) return;

  const user = process.env.GMAIL_USER;
  const pass = process.env.GMAIL_APP_PASSWORD;
  const to = arg('--to') || process.env.GMAIL_TO || user;
  if (!user || !pass) throw new Error('GMAIL_USER / GMAIL_APP_PASSWORD not set in .env');
  const notSaved = d.counts['Refused by DMS'] + d.counts['Not found in DMS'] + d.counts['Sheet has no usable value'] + d.counts.Error;
  const corrected = d.counts.Changed + d.counts['Corrected, sheet value missing'];
  await require('nodemailer').createTransport({ service: 'gmail', auth: { user, pass } }).sendMail({
    from: user, to,
    subject: `VIN master update: ${fmt(d.rows.length)} VINs checked, ${fmt(corrected)} corrected, ${fmt(notSaved)} not saved - ${stamp}`,
    html,
    attachments: [{ filename: path.basename(dashboard), path: dashboard }, { filename: path.basename(excel), path: excel }],
  });
  console.log(`[STEP] Email sent to ${to}`);
}

main().catch((err) => { console.error(err.message); process.exit(1); });
