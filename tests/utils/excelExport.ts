import { Download } from '@playwright/test';
import * as XLSX from 'xlsx';

// Reads a Playwright Download of an exported .xlsx report into a header row + data rows,
// ready to hand to pushRowsToSheet.
export function readExcelDownload(download: Download): Promise<{ headers: string[]; rows: (string | number)[][] }> {
  return download.path().then((filePath) => {
    if (!filePath) {
      throw new Error('Excel download did not produce a file path (the download may have failed).');
    }

    const workbook = XLSX.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      throw new Error(`Exported workbook at "${filePath}" has no sheets.`);
    }

    const sheet = workbook.Sheets[firstSheetName]!;
    const [headers, ...dataRows] = XLSX.utils.sheet_to_json<(string | number)[]>(sheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    return { headers: headers ?? [], rows: dataRows };
  });
}
