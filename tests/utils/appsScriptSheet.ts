const WEB_APP_URL = process.env.SHEETS_WEB_APP_URL;

// Posts rows to the Apps Script Web App bound to the Productwise_Invoice sheet. The Apps
// Script (see resources/apps-script/pushToSheet.gs) clears everything below the header and
// re-writes header + rows, so stale rows from a previous run never linger.
export async function pushRowsToSheet(headers: (string | number)[], rows: (string | number)[][]) {
  if (!WEB_APP_URL) {
    throw new Error(
      'SHEETS_WEB_APP_URL is not set in .env. Deploy the Apps Script Web App bound to the ' +
        'Productwise_Invoice sheet (see resources/apps-script/pushToSheet.gs) and set ' +
        'SHEETS_WEB_APP_URL to its deployed /exec URL.'
    );
  }

  if (rows.length === 0) {
    console.log('[STEP] No rows to push to Google Sheets — skipping.');
    return;
  }

  console.log(`[STEP] Pushing ${rows.length} rows to the Productwise_Invoice sheet via Apps Script...`);

  const response = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sheetName: 'Productwise_Invoice', headers, rows }),
  });

  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`Apps Script Web App responded with ${response.status}: ${responseText}`);
  }

  let result: { ok?: boolean; error?: string };
  try {
    result = JSON.parse(responseText);
  } catch {
    throw new Error(`Apps Script Web App returned a non-JSON response: ${responseText}`);
  }
  if (!result.ok) {
    throw new Error(`Apps Script Web App reported failure: ${result.error ?? responseText}`);
  }

  console.log('[STEP] Finished pushing rows to the sheet.');
}
