// Starts the ASM-wise Report workflow (.github/workflows/asmwise-report.yml) from Google, every day between
// 07:00 and 08:00 IST, in case GitHub drops its own scheduled runs (it skipped all of them on 27-Oct-2026).
// It starts the workflow in "auto" mode, which mails the ASMs only between 08:00 and 09:00 IST (waiting for
// 08:00 if the report is ready early) and only if no run has sent today's email yet, so it is never sent twice.
//
// Setup (once):
//   1. script.google.com > New project (or the project that already holds pushVinToGitHub.gs),
//      paste this file in, save.
//   2. Project Settings (gear) > Script properties > GITHUB_TOKEN = a fine-grained GitHub token with access
//      to only the E2E-DMS repository and "Actions: Read and write" permission. The pushVinToGitHub token
//      works once "Actions: Read and write" is added to it (github.com > Settings > Developer settings).
//   3. Run setUpAsmTrigger once (Run button) and allow the permissions it asks for.
//   4. Run triggerAsmReport once; its log should say "Started", and Actions shows a new ASM-wise Report run
//      (it sends nothing if today's email has already gone).
var ASM_REPO = 'sunita-river/E2E-DMS';
var ASM_WORKFLOW = 'asmwise-report.yml';
var ASM_HOURS = [7]; // IST; Google runs it within that hour, so the report is ready by 08:00

function triggerAsmReport() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Add GITHUB_TOKEN under Project Settings > Script properties.');
  var response = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + ASM_REPO + '/actions/workflows/' + ASM_WORKFLOW + '/dispatches', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ ref: 'main', inputs: { mode: 'auto' } }),
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
    });
  if (response.getResponseCode() !== 204) {
    throw new Error('GitHub answered ' + response.getResponseCode() + ': ' + response.getContentText());
  }
  Logger.log('Started ' + ASM_WORKFLOW + ' (auto) on ' + ASM_REPO);
}

// Replaces any earlier trigger for triggerAsmReport.
function setUpAsmTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'triggerAsmReport'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ASM_HOURS.forEach(function (hour) {
    ScriptApp.newTrigger('triggerAsmReport').timeBased().everyDays(1).atHour(hour).inTimezone('Asia/Kolkata').create();
  });
  Logger.log('triggerAsmReport will run daily at ' + ASM_HOURS.join(':00 and ') + ':00 IST.');
}
