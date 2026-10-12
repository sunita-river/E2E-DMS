// Data checks for downloaded report rows. Each check returns a list of problems (empty = OK)
// instead of throwing, so a spec can save its output first and then fail with every problem listed.

type Row = Record<string, string>;

// Report dates come as "07-09-2026 00:00:00" (grid), "07-09-2026 00:00" (1000+ row Excel path),
// "02/09/2026" (GSTR grid) or, since early Oct 2026, "07-Sep-2026" (Enquiry List Excel export).
// Returns null if the value isn't one of those.
const MONTH_ABBR = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
export function parseReportDate(value: string): Date | null {
  const m = String(value).trim().match(/^(\d{1,2})[-/ ](\d{1,2}|[A-Za-z]{3})[-/ ](\d{4})/);
  if (!m) return null;
  const month = /^\d+$/.test(m[2]!) ? Number(m[2]) - 1 : MONTH_ABBR.indexOf(m[2]!.toLowerCase());
  return month < 0 || month > 11 ? null : new Date(Number(m[3]), month, Number(m[1]));
}

const toNumber = (value: string) => Number(String(value).replace(/,/g, '').trim());

// Lists only the first few offending rows per check so a bad run doesn't produce a 20,000-line error.
function summarize(label: string, badRows: number[], detail: (i: number) => string): string[] {
  if (!badRows.length) return [];
  const sample = badRows.slice(0, 5).map((i) => `row ${i + 1}: ${detail(i)}`).join('; ');
  return [`${label}: ${badRows.length} row(s): ${sample}${badRows.length > 5 ? '; ...' : ''}`];
}

export function missingColumns(rows: Row[], required: string[]): string[] {
  if (!rows.length) return [];
  const present = new Set(Object.keys(rows[0]!));
  const missing = required.filter((c) => !present.has(c));
  return missing.length ? [`Missing column(s): ${missing.join(', ')}`] : [];
}

export function blankValues(rows: Row[], columns: string[]): string[] {
  return columns.flatMap((c) => {
    const bad = rows.flatMap((r, i) => (String(r[c] ?? '').trim() === '' ? [i] : []));
    return summarize(`Blank ${c}`, bad, () => 'blank');
  });
}

// Every date in `column` must parse and fall between from and to (whole days, inclusive).
export function datesOutsideRange(rows: Row[], column: string, from: Date, to: Date): string[] {
  const end = new Date(to.getFullYear(), to.getMonth(), to.getDate() + 1);
  const bad = rows.flatMap((r, i) => {
    const d = parseReportDate(r[column] ?? '');
    return d && d >= from && d < end ? [] : [i];
  });
  return summarize(`${column} missing or outside ${from.toDateString()} to ${to.toDateString()}`, bad, (i) => `"${rows[i]![column]}"`);
}

export function nonNumeric(rows: Row[], columns: string[]): string[] {
  return columns.flatMap((c) => {
    const bad = rows.flatMap((r, i) => (Number.isFinite(toNumber(r[c] ?? '')) && String(r[c] ?? '').trim() !== '' ? [] : [i]));
    return summarize(`${c} not a number`, bad, (i) => `"${rows[i]![c]}"`);
  });
}

// GSTR: TotalAmount must equal TaxableAmount + CGST + SGST + IGST + CESS (to within ₹1 for rounding).
export function gstrTotalsMismatch(rows: Row[]): string[] {
  const parts = ['TaxableAmount', 'CGSTInAmount', 'SGSTInAmount', 'IGSTInAmount', 'CESSInAmount'];
  const bad = rows.flatMap((r, i) => {
    const sum = parts.reduce((s, c) => s + toNumber(r[c] ?? '0'), 0);
    return Math.abs(sum - toNumber(r.TotalAmount ?? '')) > 1 ? [i] : [];
  });
  return summarize('TotalAmount does not equal Taxable + taxes', bad, (i) => `${rows[i]!.InvoiceNo}: total ${rows[i]!.TotalAmount}`);
}
