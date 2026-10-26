// Sends this PC's PDI run files (Output/PDI runs/PDI_BlankChassis_*.xlsx and PDI_InTransit_*.xlsx) to the pdi-runs branch on
// GitHub, so the GitHub PDI workflow (.github/workflows/pdi.yml) merges them into its daily file,
// history and 19:17 email along with its own hourly runs.
//
// Only the last 3 days' run files are sent (the daily merge keeps 3 days). The working tree and the
// current branch are not touched: the commit is built with a temporary index and pushed straight to
// refs/heads/pdi-runs. Run with: npm run pdi:push-runs
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const BRANCH = 'pdi-runs';
const RUNS_DIR = path.join(process.env.GSTR_OUTPUT_DIR || path.join(process.cwd(), 'Output'), 'PDI runs');
const HISTORY_FILE = path.join(path.dirname(RUNS_DIR), 'PDI_History.json');
const KEEP_DAYS = 3;

const git = (args, env) => execFileSync('git', args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, ...env } }).trim();

function runFiles() {
  if (!fs.existsSync(RUNS_DIR)) return [];
  const oldest = new Date(Date.now() - KEEP_DAYS * 24 * 3600 * 1000).toISOString().slice(0, 10);
  return fs.readdirSync(RUNS_DIR)
    // In-Transit runs too (dry runs stay on the PC: nothing was saved).
    .filter((f) => /^PDI_(BlankChassis|InTransit)_(\d{4}-\d{2}-\d{2})-.*\.xlsx$/i.test(f) && !/_DRYRUN/i.test(f) && !f.startsWith('~$'))
    .filter((f) => f.match(/_(\d{4}-\d{2}-\d{2})-/)[1] >= oldest);
}

function main() {
  const files = runFiles();
  if (!files.length) {
    console.log(`[STEP] No PDI run files from the last ${KEEP_DAYS} days in ${RUNS_DIR}, nothing to send.`);
    return;
  }

  let parent = null;
  try {
    git(['fetch', '--quiet', 'origin', `${BRANCH}:refs/remotes/origin/${BRANCH}`]);
    parent = git(['rev-parse', `refs/remotes/origin/${BRANCH}`]);
  } catch { /* first push: the branch doesn't exist yet */ }

  const index = path.join(os.tmpdir(), `pdi-runs-index-${process.pid}`);
  const env = { GIT_INDEX_FILE: index };
  try {
    if (parent) git(['read-tree', parent], env);
    else git(['read-tree', '--empty'], env);
    for (const f of files) {
      const blob = git(['hash-object', '-w', path.join(RUNS_DIR, f)]);
      git(['update-index', '--add', '--cacheinfo', `100644,${blob},runs/${f}`], env);
    }
    // The PC's all-time history too, so GitHub's totals include the days before it started running PDI.
    if (fs.existsSync(HISTORY_FILE)) {
      const blob = git(['hash-object', '-w', HISTORY_FILE]);
      git(['update-index', '--add', '--cacheinfo', `100644,${blob},history/PDI_History.json`], env);
    }
    const tree = git(['write-tree'], env);
    if (parent && tree === git(['rev-parse', `${parent}^{tree}`])) {
      console.log(`[STEP] ${files.length} PDI run file(s) already on ${BRANCH}, nothing new to send.`);
      return;
    }
    const commit = git(['commit-tree', tree, ...(parent ? ['-p', parent] : []), '-m', `PDI run files from ${os.hostname()}`]);
    git(['push', '--quiet', 'origin', `${commit}:refs/heads/${BRANCH}`]);
    console.log(`[STEP] Sent ${files.length} PDI run file(s) to the ${BRANCH} branch; the next GitHub PDI run merges them.`);
  } finally {
    fs.rmSync(index, { force: true });
  }
}

try {
  main();
} catch (err) {
  console.log(`[WARN] Could not send the PDI run files to GitHub: ${String(err.message).slice(0, 200)}`);
  process.exitCode = 1;
}
