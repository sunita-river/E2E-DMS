// Sends the VIN details sheet to GitHub so the PDI workflow (.github/workflows/pdi.yml) can read it
// while the sheet itself stays private. Runs every 30 minutes as the person who set up the trigger;
// it only commits when the sheet's contents have changed.
//
// Only one copy of this script may push to the vin-data branch: two copies reading different sheets
// would keep overwriting each other's VIN_Details.csv. Moving to another sheet? Delete the old copy's
// trigger first (Apps Script > Triggers).
//
// Setup (once):
//   1. Open the VIN sheet > Extensions > Apps Script, paste this file in, save. Or, for a sheet you can
//      only view (e.g. "VIN BIN Update (Sunitha)", owned by someone else): script.google.com > New project,
//      paste this file in, and set SPREADSHEET_ID below; it then runs as you and needs only view access.
//   2. Project Settings (gear) > Script properties > Add: GITHUB_TOKEN = a fine-grained GitHub token
//      with access to only the E2E-DMS repository and "Contents: Read and write" permission.
//   3. Run setUpTrigger once (Run button) and allow the permissions it asks for.
//   4. Run pushVinToGitHub once to send the sheet now; its log shows what happened.
//
// The sheet goes to VIN_Details.csv on the vin-data branch (created on the first run), as the
// values shown in the sheet, so long serial numbers stay exactly as typed.
var GITHUB_REPO = 'sunita-river/E2E-DMS';
var BRANCH = 'vin-data';
var FILE_PATH = 'VIN_Details.csv';
var SHEET_NAME = ''; // blank = the first tab
// "VIN BIN Update (Sunitha)"; blank = the spreadsheet this script is attached to.
var SPREADSHEET_ID = '13GyvyCiYFchfEemtICD8ez9jx3EGADK764pAEQMdEb0';

function pushVinToGitHub() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Add GITHUB_TOKEN under Project Settings > Script properties.');

  var spreadsheet = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SHEET_NAME ? spreadsheet.getSheetByName(SHEET_NAME) : spreadsheet.getSheets()[0];
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found in this spreadsheet.');
  var values = sheet.getDataRange().getDisplayValues();
  var csv = values.map(function (row) { return row.map(csvCell).join(','); }).join('\r\n') + '\r\n';

  var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, csv, Utilities.Charset.UTF_8));
  if (props.getProperty('LAST_HASH') === hash) {
    Logger.log('Sheet unchanged since the last push, nothing sent.');
    return;
  }

  ensureBranch(token);
  var filePath = '/contents/' + encodeURIComponent(FILE_PATH);
  // The "object" media type still returns the sha for files over 1 MB (the sheet's CSV is ~2 MB).
  var existing = github('GET', filePath + '?ref=' + BRANCH, token, null, true, 'application/vnd.github.object');
  var body = {
    message: 'VIN sheet: ' + (values.length - 1) + ' rows (' + sheet.getName() + ')',
    content: Utilities.base64Encode(csv, Utilities.Charset.UTF_8),
    branch: BRANCH,
  };
  if (existing) body.sha = existing.sha;
  github('PUT', filePath, token, body);
  props.setProperty('LAST_HASH', hash);
  Logger.log('Sent ' + (values.length - 1) + ' rows to ' + GITHUB_REPO + ' ' + BRANCH + ':' + FILE_PATH);
}

// Every 30 minutes; replaces any earlier trigger for pushVinToGitHub.
function setUpTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'pushVinToGitHub'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pushVinToGitHub').timeBased().everyMinutes(30).create();
  Logger.log('pushVinToGitHub will run every 30 minutes.');
}

// The data branch starts from the default branch; after that only VIN_Details.csv changes on it.
function ensureBranch(token) {
  if (github('GET', '/git/ref/heads/' + BRANCH, token, null, true)) return;
  var repo = github('GET', '', token);
  var base = github('GET', '/git/ref/heads/' + repo.default_branch, token);
  github('POST', '/git/refs', token, { ref: 'refs/heads/' + BRANCH, sha: base.object.sha });
}

function github(method, path, token, body, allow404, accept) {
  var options = {
    method: method,
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, Accept: accept || 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
  };
  if (body) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + GITHUB_REPO + path, options);
  var code = res.getResponseCode();
  if (code === 404 && allow404) return null;
  if (code >= 300) throw new Error(method + ' ' + (path || '/') + ': HTTP ' + code + ' ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText());
}

function csvCell(value) {
  var s = String(value);
  return /[",\r\n]|^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
