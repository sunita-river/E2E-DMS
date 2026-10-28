// Progress page for a running VIN master update (pdi/Vin_Master_update.spec.ts): reads the progress file and
// the VIN sheet and writes Output/VIN Master/VIN_Master_Progress.html, which reloads itself every minute.
//
//   node "pdi/scripts/vinMasterProgress.js"           write the page once
//   node "pdi/scripts/vinMasterProgress.js" --watch   rewrite it every minute until stopped (Ctrl+C)
//   add --merge to first take the GitHub jobs' results (vinMasterGitHub.js --merge) each time
//
// Each VIN counts once, by its latest result, as the spec's own report does; DMS refusals and pages that
// did not load are retried by the next run, so they count as problems, not done.
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..', '..');
const MASTER_DIR = path.join(ROOT, 'Output', 'VIN Master');
const PROGRESS_FILE = path.join(MASTER_DIR, 'VIN_Master_progress.jsonl');
const VIN_FILE = process.env.PDI_VIN_FILE || path.join(ROOT, 'resources', 'VIN_Details.csv');
const PAGE = path.join(MASTER_DIR, 'VIN_Master_Progress.html');
const BUCKET_MIN = 10;

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const fmt = (n) => n.toLocaleString('en-IN');
const hm = (d) => d.toLocaleTimeString('en-IN', { hour: 'numeric', minute: '2-digit' });

// "28/10/2026, 5:26:17 pm"
function parseAt(at) {
  const m = /^(\d+)\/(\d+)\/(\d+), (\d+):(\d+):(\d+) (am|pm)$/i.exec(at || '');
  if (!m) return null;
  const h = (+m[4] % 12) + (/pm/i.test(m[7]) ? 12 : 0);
  return new Date(+m[3], +m[2] - 1, +m[1], h, +m[5], +m[6]);
}

// Same as outcomeOf in Vin_Master_update.spec.ts.
function outcomeOf(status) {
  if (/^Already correct/.test(status)) return 'Already correct';
  if (/^Updated/.test(status)) return 'Changed';
  if (/^Dry run|\(dry run\)/.test(status)) return 'Dry run';
  if (/; the rest saved/.test(status)) return 'Partly changed';
  if (/^VIN sheet has no usable/.test(status)) return 'Sheet has no usable value';
  if (/Show did not load/.test(status)) return 'Not found in DMS';
  if (/^Save clicked but values not kept/.test(status)) return 'Refused by DMS';
  return 'Error';
}

function build() {
  const sheetVins = fs.readFileSync(VIN_FILE, 'utf8').split(/\r?\n/).slice(1).filter((l) => l.trim()).length;
  const latest = new Map();
  const times = [];
  for (const line of fs.existsSync(PROGRESS_FILE) ? fs.readFileSync(PROGRESS_FILE, 'utf8').split('\n') : []) {
    if (!line.trim()) continue;
    try {
      const r = JSON.parse(line);
      latest.set(r.ChassisNo, r);
      const t = parseAt(r.At);
      if (t) times.push(t);
    } catch { /* a line cut off by a stop */ }
  }
  const rows = [...latest.values()];
  const counts = {};
  for (const r of rows) counts[outcomeOf(r.Status || '')] = (counts[outcomeOf(r.Status || '')] || 0) + 1;
  const problemsN = (counts['Refused by DMS'] || 0) + (counts['Not found in DMS'] || 0) + (counts.Error || 0);
  const checked = rows.length;
  const done = checked - problemsN;
  const remaining = Math.max(sheetVins - done, 0);
  const pct = sheetVins ? (done / sheetVins) * 100 : 0;

  // Speed over the last 15 minutes of activity; the ETA assumes the run keeps that pace.
  times.sort((a, b) => a - b);
  const last = times[times.length - 1];
  const now = new Date();
  const recent = last ? times.filter((t) => t >= last - 15 * 60e3) : [];
  const span = recent.length > 1 ? (recent[recent.length - 1] - recent[0]) / 1000 : 0;
  const perSec = span ? recent.length / span : 0;
  const running = last && now - last < 3 * 60e3;
  const eta = running && perSec ? new Date(now.getTime() + (remaining / perSec) * 1000) : null;

  // VINs checked per 10 minutes, today only.
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const buckets = new Map();
  for (const t of times) {
    if (t < today) continue;
    const k = Math.floor(t.getTime() / (BUCKET_MIN * 60e3)) * BUCKET_MIN * 60e3;
    buckets.set(k, (buckets.get(k) || 0) + 1);
  }
  const bars = [...buckets.entries()].sort((a, b) => a[0] - b[0]);

  const changed = rows.filter((r) => ['Changed', 'Partly changed'].includes(outcomeOf(r.Status || ''))).sort((a, b) => parseAt(b.At) - parseAt(a.At));
  const problems = rows.filter((r) => ['Refused by DMS', 'Not found in DMS', 'Error'].includes(outcomeOf(r.Status || '')));

  return { sheetVins, checked, done, remaining, pct, counts, perSec, running, eta, last, bars, changed, problems, now };
}

function barChart(bars) {
  if (!bars.length) return '<p class="muted">Nothing checked yet today.</p>';
  const W = 760, H = 220, padL = 44, padB = 28, padT = 12;
  const max = Math.max(...bars.map((b) => b[1]));
  const niceMax = Math.ceil(max / 500) * 500 || 500;
  const slot = (W - padL) / bars.length;
  const bw = Math.max(Math.min(slot - 2, 36), 2);
  const y = (v) => padT + (H - padT - padB) * (1 - v / niceMax);
  const grid = [0, 0.5, 1].map((f) => {
    const v = niceMax * f;
    return `<line x1="${padL}" x2="${W}" y1="${y(v)}" y2="${y(v)}" class="grid"/><text x="${padL - 6}" y="${y(v) + 4}" class="axis" text-anchor="end">${fmt(v)}</text>`;
  }).join('');
  const every = Math.ceil(bars.length / 10);
  const marks = bars.map(([k, v], i) => {
    const x = padL + i * slot + (slot - bw) / 2;
    const top = y(v), h = H - padB - top;
    const r = Math.min(4, bw / 2, h);
    const from = new Date(k), to = new Date(k + BUCKET_MIN * 60e3);
    // Rounded top only, square at the baseline.
    const d = `M${x},${H - padB} V${top + r} Q${x},${top} ${x + r},${top} H${x + bw - r} Q${x + bw},${top} ${x + bw},${top + r} V${H - padB} Z`;
    const label = i % every === 0 ? `<text x="${x + bw / 2}" y="${H - 8}" class="axis" text-anchor="middle">${hm(from)}</text>` : '';
    return `<g class="bar"><rect x="${padL + i * slot}" y="${padT}" width="${slot}" height="${H - padT - padB}" fill="transparent"/>` +
      `<path d="${d}"/><title>${hm(from)}–${hm(to)}: ${fmt(v)} VINs checked</title></g>${label}`;
  }).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="VINs checked per ${BUCKET_MIN} minutes today">${grid}` +
    `<line x1="${padL}" x2="${W}" y1="${H - padB}" y2="${H - padB}" class="base"/>${marks}</svg>`;
}

function render(s) {
  const tile = (label, value, sub = '') => `<div class="tile"><div class="label">${label}</div><div class="value">${value}</div><div class="sub">${sub}</div></div>`;
  const status = s.running ? `<span class="dot run"></span>Running` : `<span class="dot stop"></span>Not running${s.last ? ` · last VIN checked ${hm(s.last)}` : ''}`;
  const breakdown = Object.entries(s.counts).sort((a, b) => b[1] - a[1])
    .map(([k, v]) => `<tr><td>${esc(k)}</td><td class="num">${fmt(v)}</td></tr>`).join('');
  const changedRows = s.changed.map((r) => `<tr><td>${esc(r.At)}</td><td class="mono">${esc(r.ChassisNo)}</td><td>${esc(r.Changed)}</td>` +
    ['EngineNo', 'BatteryDetails', 'VehicleChargerNo'].map((f) => {
      const before = r[`Before${f}`] || '', after = r[f] || '';
      return before.toUpperCase() === after.toUpperCase() ? `<td class="mono muted">${esc(after)}</td>` : `<td class="mono"><s>${esc(before) || '(blank)'}</s> → <b>${esc(after)}</b></td>`;
    }).join('') + '</tr>').join('');
  const problemRows = s.problems.map((r) => `<tr><td class="mono">${esc(r.ChassisNo)}</td><td>${esc(outcomeOf(r.Status))}</td><td class="msg">${esc(r.Status)}</td></tr>`).join('');

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="60"><title>VIN Master Progress</title>
<style>
:root { --page:#f9f9f7; --surface:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e; --muted:#898781; --grid:#e1e0d9; --base:#c3c2b7;
  --fill:#2a78d6; --track:#e1e0d9; --good:#0ca30c; --critical:#d03b3b; --border:#e1e0d9; }
@media (prefers-color-scheme: dark) { :root { --page:#0d0d0d; --surface:#1a1a19; --ink:#f0efec; --ink2:#c3c2b7; --grid:#2c2c2a;
  --base:#383835; --fill:#3987e5; --track:#2c2c2a; --border:#2c2c2a; } }
* { box-sizing:border-box; }
body { margin:0; background:var(--page); color:var(--ink); font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif; }
main { max-width:1000px; margin:0 auto; padding:24px 16px 48px; }
h1 { font-size:20px; margin:0 0 4px; } h2 { font-size:15px; margin:28px 0 10px; }
.muted, .sub, .axis { color:var(--muted); }
.head { display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px; align-items:baseline; }
.dot { display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:6px; }
.dot.run { background:var(--good); } .dot.stop { background:var(--critical); }
.card { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:16px; }
.progress { height:22px; background:var(--track); border-radius:6px; overflow:hidden; margin:10px 0 6px; }
.progress div { height:100%; background:var(--fill); border-radius:6px 0 0 6px; }
.big { font-size:32px; font-weight:650; font-variant-numeric:tabular-nums; }
.tiles { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:10px; margin-top:12px; }
.tile { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:12px; }
.tile .label { color:var(--ink2); font-size:12px; } .tile .value { font-size:22px; font-weight:600; font-variant-numeric:tabular-nums; }
.tile .sub { font-size:12px; }
svg { width:100%; height:auto; display:block; }
svg .grid { stroke:var(--grid); stroke-width:1; } svg .base { stroke:var(--base); stroke-width:1; }
svg .axis { fill:var(--muted); font-size:11px; } svg .bar path { fill:var(--fill); } svg .bar:hover path { opacity:.75; }
.scroll { overflow-x:auto; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th, td { text-align:left; padding:6px 8px; border-bottom:1px solid var(--border); vertical-align:top; }
th { color:var(--ink2); font-weight:600; } .num { text-align:right; font-variant-numeric:tabular-nums; }
.mono { font-family:ui-monospace,Consolas,monospace; white-space:nowrap; } .msg { color:var(--ink2); }
s { color:var(--muted); }
</style></head><body><main>
<div class="head"><h1>VIN master update</h1><div>${status} · page updated ${hm(s.now)}, reloads every minute</div></div>

<div class="card" style="margin-top:14px">
  <div class="big">${s.pct.toFixed(1)}%</div>
  <div class="progress" role="progressbar" aria-valuenow="${s.pct.toFixed(1)}" aria-valuemin="0" aria-valuemax="100"><div style="width:${s.pct.toFixed(2)}%"></div></div>
  <div class="muted">${fmt(s.done)} of ${fmt(s.sheetVins)} VINs done · ${fmt(s.remaining)} left</div>
</div>

<div class="tiles">
  ${tile('Speed', s.perSec ? `${s.perSec.toFixed(1)}/s` : '–', 'VINs per second, last 15 min')}
  ${tile('Expected finish', s.eta ? hm(s.eta) : '–', s.eta ? s.eta.toLocaleDateString('en-IN', { day: 'numeric', month: 'short' }) + ' at this speed' : 'when running')}
  ${tile('Corrected & saved', fmt(s.changed.length), 'listed below, before → after')}
  ${tile('Problems', fmt(s.problems.length), 'refused by DMS or not loaded')}
</div>

<h2>VINs checked per ${BUCKET_MIN} minutes, today</h2>
<div class="card">${barChart(s.bars)}</div>

<h2>Results so far (each VIN once, latest result)</h2>
<div class="card scroll"><table><tr><th>Result</th><th class="num">VINs</th></tr>${breakdown}</table></div>

<h2>Corrected and saved (${fmt(s.changed.length)})</h2>
<div class="card scroll">${s.changed.length ? `<table><tr><th>Saved at</th><th>Chassis No</th><th>Changed</th><th>Engine No</th><th>Battery</th><th>Charger</th></tr>${changedRows}</table>` : '<p class="muted">None yet.</p>'}</div>

<h2>Problems to fix in DMS (${fmt(s.problems.length)})</h2>
<div class="card scroll">${s.problems.length ? `<table><tr><th>Chassis No</th><th>Result</th><th>DMS message</th></tr>${problemRows}</table>` : '<p class="muted">None.</p>'}</div>
</main></body></html>`;
}

function write() {
  if (process.argv.includes('--merge')) {
    try { require('./vinMasterGitHub').merge(); } catch (err) { console.log(`[WARN] Merge failed: ${String(err.message).split('\n')[0]}`); }
  }
  fs.mkdirSync(MASTER_DIR, { recursive: true });
  const s = build();
  fs.writeFileSync(PAGE, render(s));
  return s;
}

const s = write();
console.log(`[STEP] ${PAGE}: ${s.done} of ${s.sheetVins} VINs done (${s.pct.toFixed(1)}%), ${s.changed.length} corrected, ${s.problems.length} problem(s)`);
if (process.argv.includes('--watch')) setInterval(write, 60e3);
