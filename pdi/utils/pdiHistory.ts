import * as fs from 'node:fs';
import { outcomeOf, ResultRow } from './pdiReport';

// Long-term PDI history: Output/PDI_History.json keeps every day's chassis results for good (the
// 3-day clean-up never touches it), so the report can show totals "till now" and a day-wise trend.
//
// Each day stores its blank PDI chassis (dealer, chassis no, outcome), the dealers checked and failed,
// and its In-Transit vehicles. Totals walk the days in order, keeping each chassis's latest result:
//   - A chassis that is updated counts as updated from then on.
//   - A blank PDI chassis that was not updated on one day but is missing from a later day's list counts
//     as updated too, as within a day, but only when that later day's PDI run checked its dealer. A day
//     with no PDI run (only In-Transit) checks nobody, so nothing is assumed updated.
//   - In-Transit vehicles count only by their own results: one dropping off the In-Transit list may
//     have left stock rather than been filled, so it stays open until a run updates it.

type Outcome = ReturnType<typeof outcomeOf>;
type StoredRow = [dealer: string, chassis: string, outcome: Outcome];
// checked: dealers the day's PDI runs checked (missing on days saved before it was kept).
interface StoredDay { rows: StoredRow[]; failed: string[]; checked?: string[]; inTransit?: StoredRow[] }
interface Store { version: 1; days: Record<string, StoredDay> }

export interface HistoryDay {
  day: string;
  found: number; updated: number; missing: number; error: number;   // that day's blank PDI + In-Transit chassis
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
  const inTransitKeys = new Set<string>();
  const days = Object.keys(store.days).sort().map((day) => {
    const { rows, failed, checked, inTransit = [] } = store.days[day]!;
    const seen = new Set(rows.map(([d, c]) => `${d}|${c}`));
    // Days saved before "checked" was kept: a day with blank rows came from a full PDI run (all dealers
    // but the failed ones checked); a day without any had no PDI run.
    const wasChecked = (dealer: string) => checked ? checked.includes(dealer) : rows.length > 0 && !failed.includes(dealer);
    for (const [key, outcome] of latest) {
      if (outcome !== 'updated' && !inTransitKeys.has(key) && !seen.has(key) && wasChecked(key.split('|')[0]!)) latest.set(key, 'updated');
    }
    const count = { updated: 0, dryrun: 0, missing: 0, error: 0 } as Record<Outcome, number>;
    inTransit.forEach(([d, c, o]) => { count[o]++; latest.set(`${d}|${c}`, o); inTransitKeys.add(`${d}|${c}`); });
    // The blank PDI result wins over the day's In-Transit result for the same chassis.
    rows.forEach(([d, c, o]) => { count[o]++; latest.set(`${d}|${c}`, o); inTransitKeys.delete(`${d}|${c}`); });
    const open = (o: Outcome) => [...latest.values()].filter((v) => v === o).length;
    return {
      day, found: rows.length + inTransit.length, updated: count.updated, missing: count.missing, error: count.error + count.dryrun,
      totalFound: latest.size, totalUpdated: open('updated'), openMissing: open('missing'), openError: open('error'),
    };
  });
  return { since: days[0]?.day ?? null, days, latest: days[days.length - 1] ?? null };
}

const toStored = (rows: ResultRow[]): StoredRow[] => rows.map((r) => [r.Dealer ?? '', r.ChassisNo ?? '', outcomeOf(r)]);

export interface HistoryEntry { day: string; rows: ResultRow[]; failed: string[]; checked: string[]; inTransit: ResultRow[] }

// Saves (or replaces) the given days in the history file and returns the summary over all days.
// inTransitAll (every In-Transit vehicle's latest result, with its Day) fills in the In-Transit vehicles of
// days stored without them, i.e. days whose PDI_Daily file is already gone.
export function recordPdiHistory(file: string, entries: HistoryEntry[], inTransitAll: ResultRow[] = []): PdiHistory {
  const store = load(file);
  for (const { day, rows, failed, checked, inTransit } of entries) {
    store.days[day] = { rows: toStored(rows), failed, checked, inTransit: toStored(inTransit) };
  }
  const byDay = new Map<string, ResultRow[]>();
  inTransitAll.forEach((r) => { const d = String(r.Day ?? ''); if (d) byDay.set(d, [...(byDay.get(d) ?? []), r]); });
  for (const [day, rows] of byDay) {
    const stored = store.days[day];
    if (!stored) store.days[day] = { rows: [], failed: [], checked: [], inTransit: toStored(rows) };
    else if (!stored.inTransit) stored.inTransit = toStored(rows);
  }
  try {
    fs.writeFileSync(file, JSON.stringify(store));
  } catch (err) {
    console.log(`[WARN] Could not save the PDI history: ${String((err as Error).message).slice(0, 120)}`);
  }
  return summarize(store);
}
