// Deploy this bound to the target spreadsheet: Extensions > Apps Script, paste this in,
// then Deploy > New deployment > type "Web app", execute as "Me", access "Anyone" (or
// "Anyone with Google account" if that's acceptable — "Anyone" is required if the test
// machine won't be logged into a Google account when it calls this).
//
// Copy the resulting /exec URL into this project's .env as SHEETS_WEB_APP_URL.
//
// Expects a POST body of: { sheetName: string, headers: any[], rows: any[][] }
// Clears everything in the named sheet, then writes headers as row 1 and rows below it.
function doPost(e) {
  var result = { ok: true };
  try {
    var payload = JSON.parse(e.postData.contents);
    var sheetName = payload.sheetName;
    var headers = payload.headers;
    var rows = payload.rows;

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('Sheet "' + sheetName + '" not found in this spreadsheet.');
    }

    sheet.clearContents();

    var values = [headers].concat(rows);
    sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
