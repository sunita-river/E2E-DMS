import * as fs from 'node:fs';
import { outcomeOf, ResultRow } from './pdiReport';

// Long-term PDI history: Output/PDI_History.json keeps every day's chassis results for good (the
// 3-day clean-up never touches it), so the report can show totals "till now" and a day-wise trend.
//
// Each day stores its chassis (dealer, chassis no, outcome) and the dealers whose check failed.
// Totals walk the days in order, keeping each chassis's latest result:
//   - A chassis that is updated counts as updated from then on.
//   - A chassis that was not updated on one day but is missing from a later day's list (its dealer was
//     checked and the row was no longer blank) counts as updated too — the same rule as within a day.
//     Dealers that failed that day are skipped, since their chassis weren't checked.

type Outcome = ReturnType<typeof outcomeOf>;
interface StoredDay { rows: [dealer: string, chassis: string, outcome: Outcome][]; failed: string[] }
interface Store { version: 1; days: Record<string, StoredDay> }

export interface HistoryDay {
  day: string;
  found: number; updated: number; missing: number; error: number;   // that day's blank chassis
  totalFound: number; totalUpdated: number; openMissing: number; openError: number; // till that day
}
export interface PdiHistory { since: string | null; days: HistoryDay[]; latest: HistoryDay | null }

function load(file: string): Store {
  try {
    const s = JSON.parse(fs.readFileSync(file, 'utf8')) as Store;
    if (s && s.days) return s;
  } catch { /* missing or unreadable: start fresh */ }
  return { version: 1, days: {} };
}

function summarize(store: Store): PdiHistory {
  const latest = new Map<string, Outcome>();
  const days = Object.keys(store.days).sort().map((day) => {
    const { rows, failed } = store.days[day]!;
    const seen = new Set(rows.map(([d, c]) => `${d}|${c}`));
    for (const [key, outcome] of latest) {
      if (outcome !== 'updated' && !seen.has(key) && !failed.includes(key.split('|')[0]!)) latest.set(key, 'updated');
    }
    const count = { updated: 0, dryrun: 0, missing: 0, error: 0 } as Record<Outcome, number>;
    rows.forEach(([d, c, o]) => { count[o]++; latest.set(`${d}|${c}`, o); });
    const open = (o: Outcome) => [...latest.values()].filter((v) => v === o).length;
    return {
      day, found: rows.length, updated: count.updated, missing: count.missing, error: count.error,
      totalFound: latest.size, totalUpdated: open('updated'), openMissing: open('missing'), openError: open('error'),
    };
  });
  return { since: days[0]?.day ?? null, days, latest: days[days.length - 1] ?? null };
}

// Saves (or replaces) the given days in the history file and returns the summary over all days.
export function recordPdiHistory(file: string, entries: { day: string; rows: ResultRow[]; failed: string[] }[]): PdiHistory {
  const store = load(file);
  for (const { day, rows, failed } of entries) {
    store.days[day] = { rows: rows.map((r) => [r.Dealer ?? '', r.ChassisNo ?? '', outcomeOf(r)]), failed };
  }
  try {
    fs.writeFileSync(file, JSON.stringify(store));
  } catch (err) {
    console.log(`[WARN] Could not save the PDI history: ${String((err as Error).message).slice(0, 120)}`);
  }
  return summarize(store);
}
