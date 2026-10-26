// Merges one PDI history file into another, so the PC and the GitHub workflow end up with the same
// "till now" totals whichever of them builds the report (pdi/scripts/pdiSync.js calls it both ways).
//   node "pdi/scripts/pdiMergeHistory.js" <target PDI_History.json> <incoming PDI_History.json>
// Days only in the incoming file are copied. For a day in both, chassis are combined (dealer + chassis
// no) and "updated" wins over any other outcome, for the blank PDI chassis and the In-Transit vehicles
// alike; the dealers checked are combined; the target's list of failed dealers is kept.
const fs = require('node:fs');
const path = require('node:path');

const load = (f) => {
  try {
    const s = JSON.parse(fs.readFileSync(f, 'utf8'));
    if (s && s.days) return s;
  } catch { /* missing or unreadable */ }
  return { version: 1, days: {} };
};

function mergeHistory(target, incoming) {
  const base = load(target);
  const other = load(incoming);
  let addedDays = 0, addedRows = 0;

  // Chassis lists combined by dealer + chassis no; "updated" wins over any other outcome.
  const combine = (mineRows, otherRows) => {
    const rows = new Map(mineRows.map((r) => [`${r[0]}|${r[1]}`, r]));
    for (const r of otherRows) {
      const key = `${r[0]}|${r[1]}`;
      const have = rows.get(key);
      if (!have) { rows.set(key, r); addedRows++; }
      else if (r[2] === 'updated' && have[2] !== 'updated') rows.set(key, r);
    }
    return [...rows.values()];
  };

  for (const [day, entry] of Object.entries(other.days)) {
    const mine = base.days[day];
    if (!mine) {
      base.days[day] = entry;
      addedDays++;
      continue;
    }
    // Dealers checked: combined. A day saved before they were kept, with blank rows, means a full PDI run;
    // it stays without the list so every dealer but the failed ones still counts as checked.
    const legacyFullRun = !mine.checked && mine.rows.length > 0;
    mine.rows = combine(mine.rows, entry.rows);
    if (entry.inTransit || mine.inTransit) mine.inTransit = combine(mine.inTransit || [], entry.inTransit || []);
    if (!legacyFullRun && (mine.checked || entry.checked)) mine.checked = [...new Set([...(mine.checked || []), ...(entry.checked || [])])];
  }
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, JSON.stringify(base));
  console.log(`[STEP] PDI history merged: ${addedDays} day(s) and ${addedRows} chassis added; ${Object.keys(base.days).length} day(s) in total.`);
}

module.exports = { mergeHistory };

if (require.main === module) {
  const [target, incoming] = process.argv.slice(2);
  if (!target || !incoming) {
    console.log('Usage: node "pdi/scripts/pdiMergeHistory.js" <target> <incoming>');
    process.exit(1);
  }
  mergeHistory(target, incoming);
}
