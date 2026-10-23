// Merges another PDI history file (from a PC, via the pdi-runs branch) into this one, so the GitHub
// workflow's "till now" totals include days it never ran (it started on 16-Oct-2026; the PC has 29-Sep on).
//   node "pdi/scripts/pdiMergeHistory.js" <target PDI_History.json> <incoming PDI_History.json>
// Days only in the incoming file are copied. For a day in both, chassis are combined (dealer + chassis
// no) and "updated" wins over any other outcome; the target's list of failed dealers is kept.
const fs = require('node:fs');

const [target, incoming] = process.argv.slice(2);
if (!target || !incoming) {
  console.log('Usage: node "pdi/scripts/pdiMergeHistory.js" <target> <incoming>');
  process.exit(1);
}
const load = (f) => {
  try {
    const s = JSON.parse(fs.readFileSync(f, 'utf8'));
    if (s && s.days) return s;
  } catch { /* missing or unreadable */ }
  return { version: 1, days: {} };
};

const base = load(target);
const other = load(incoming);
let addedDays = 0, addedRows = 0;
for (const [day, entry] of Object.entries(other.days)) {
  const mine = base.days[day];
  if (!mine) {
    base.days[day] = entry;
    addedDays++;
    continue;
  }
  const rows = new Map(mine.rows.map((r) => [`${r[0]}|${r[1]}`, r]));
  for (const r of entry.rows) {
    const key = `${r[0]}|${r[1]}`;
    const have = rows.get(key);
    if (!have) { rows.set(key, r); addedRows++; }
    else if (r[2] === 'updated' && have[2] !== 'updated') rows.set(key, r);
  }
  mine.rows = [...rows.values()];
}
fs.mkdirSync(require('node:path').dirname(target), { recursive: true });
fs.writeFileSync(target, JSON.stringify(base));
console.log(`[STEP] PDI history merged: ${addedDays} day(s) and ${addedRows} chassis added; ${Object.keys(base.days).length} day(s) in total.`);
