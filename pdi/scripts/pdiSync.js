// Keeps this copy of the PDI results in step with the shared copy on the pdi-runs branch, so the PC and
// the GitHub PDI workflow (.github/workflows/pdi.yml) always build the same daily file, "till now"
// totals and master, and send the day's email only once, whichever of them sends it.
//
//   node "pdi/scripts/pdiSync.js" --pull   take what the other side added (before a run or a report)
//   node "pdi/scripts/pdiSync.js"          take it, then send ours (after a run or a report)
//
// Shared on the branch:
//   runs/                               real PDI and In-Transit run files and emailed_<day>.txt markers, last 3 days
//   history/PDI_History.json            all-time history, merged both ways (pdiMergeHistory.js)
//   history/PDI_InTransit_History.json  every In-Transit vehicle's latest result, the newer Day wins
// The working tree and the current branch are not touched: the commit is built with a temporary index
// and pushed straight to refs/heads/pdi-runs, retried if someone else pushed in between.
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { mergeHistory } = require('./pdiMergeHistory');

const BRANCH = 'pdi-runs';
const REMOTE_REF = `refs/remotes/origin/${BRANCH}`;
const OUTPUT_DIR = process.env.GSTR_OUTPUT_DIR || path.join(process.cwd(), 'Output');
const RUNS_DIR = path.join(OUTPUT_DIR, 'PDI runs');
const HISTORIES = ['PDI_History.json', 'PDI_InTransit_History.json'];
const KEEP_DAYS = 3;
const PULL_ONLY = process.argv.includes('--pull');

const git = (args, env) => execFileSync('git', args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, ...env } }).trim();
// Same days as the daily clean-up (pdiDaily.ts keptDays): today and the 2 before, by local date.
const oldest = (() => {
  const d = new Date();
  d.setDate(d.getDate() - (KEEP_DAYS - 1));
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
})();

// Run files (dry runs stay where they were made: nothing was saved) and the "emailed" markers.
function shared(name) {
  if (name.startsWith('~$')) return false;
  const run = name.match(/^PDI_(?:BlankChassis|InTransit)_(\d{4}-\d{2}-\d{2})-.*\.xlsx$/i);
  const mark = name.match(/^emailed_(\d{4}-\d{2}-\d{2})\.txt$/i);
  const day = run && !/_DRYRUN/i.test(name) ? run[1] : mark?.[1];
  return !!day && day >= oldest;
}

function fetchBranch() {
  try {
    git(['fetch', '--quiet', 'origin', `+${BRANCH}:${REMOTE_REF}`]);
    return git(['rev-parse', REMOTE_REF]);
  } catch { return null; } // first sync: the branch doesn't exist yet
}

function pull(parent) {
  if (!parent) return;
  fs.mkdirSync(RUNS_DIR, { recursive: true });
  let added = 0;
  for (const f of git(['ls-tree', '--name-only', parent, 'runs/']).split('\n').filter(Boolean)) {
    const name = path.basename(f);
    if (!shared(name) || fs.existsSync(path.join(RUNS_DIR, name))) continue;
    fs.writeFileSync(path.join(RUNS_DIR, name), execFileSync('git', ['show', `${parent}:${f}`]));
    added++;
  }
  if (added) console.log(`[STEP] Took ${added} PDI file(s) from the ${BRANCH} branch.`);

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pdi-sync-'));
  try {
    for (const h of HISTORIES) {
      let incoming;
      try { incoming = execFileSync('git', ['show', `${parent}:history/${h}`], { stdio: ['ignore', 'pipe', 'ignore'] }); } catch { continue; }
      const file = path.join(tmp, h);
      fs.writeFileSync(file, incoming);
      if (h === 'PDI_History.json') mergeHistory(path.join(OUTPUT_DIR, h), file);
      else mergeInTransit(path.join(OUTPUT_DIR, h), file);
    }
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

// Latest result per dealer + chassis: the incoming row wins when its Day is the same or newer.
function mergeInTransit(target, incoming) {
  const load = (f) => { try { return JSON.parse(fs.readFileSync(f, 'utf8')) || {}; } catch { return {}; } };
  const base = load(target);
  for (const [key, row] of Object.entries(load(incoming))) {
    if (!base[key] || String(base[key].Day) <= String(row.Day)) base[key] = row;
  }
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, JSON.stringify(base));
}

// One commit on top of parent: our shared files added, files older than KEEP_DAYS dropped.
function push(parent) {
  const index = path.join(os.tmpdir(), `pdi-sync-index-${process.pid}`);
  const env = { GIT_INDEX_FILE: index };
  try {
    if (parent) git(['read-tree', parent], env);
    else git(['read-tree', '--empty'], env);
    for (const f of git(['ls-files'], env).split('\n').filter((f) => f.startsWith('runs/'))) {
      if (!shared(path.basename(f))) git(['update-index', '--force-remove', f], env);
    }
    const local = fs.existsSync(RUNS_DIR) ? fs.readdirSync(RUNS_DIR).filter(shared) : [];
    for (const f of local) {
      const blob = git(['hash-object', '-w', path.join(RUNS_DIR, f)]);
      git(['update-index', '--add', '--cacheinfo', `100644,${blob},runs/${f}`], env);
    }
    for (const h of HISTORIES) {
      if (!fs.existsSync(path.join(OUTPUT_DIR, h))) continue;
      const blob = git(['hash-object', '-w', path.join(OUTPUT_DIR, h)]);
      git(['update-index', '--add', '--cacheinfo', `100644,${blob},history/${h}`], env);
    }
    const tree = git(['write-tree'], env);
    if (parent && tree === git(['rev-parse', `${parent}^{tree}`])) {
      console.log(`[STEP] The ${BRANCH} branch already has everything, nothing to send.`);
      return true;
    }
    const who = { GIT_AUTHOR_NAME: 'PDI sync', GIT_AUTHOR_EMAIL: 'pdi-sync@users.noreply.github.com' };
    const commit = git(['commit-tree', tree, ...(parent ? ['-p', parent] : []), '-m', `PDI sync from ${process.env.GITHUB_ACTIONS ? 'GitHub' : os.hostname()}`],
      { ...who, GIT_COMMITTER_NAME: who.GIT_AUTHOR_NAME, GIT_COMMITTER_EMAIL: who.GIT_AUTHOR_EMAIL });
    try {
      git(['push', '--quiet', 'origin', `${commit}:refs/heads/${BRANCH}`]);
    } catch (err) {
      if (/rejected|non-fast-forward|fetch first/i.test(String(err.stderr || err.message))) return false; // someone pushed meanwhile
      throw err;
    }
    console.log(`[STEP] Sent ${local.length} PDI file(s) and the histories to the ${BRANCH} branch.`);
    return true;
  } finally {
    fs.rmSync(index, { force: true });
  }
}

function main() {
  for (let attempt = 1; attempt <= 3; attempt++) {
    const parent = fetchBranch();
    pull(parent);
    if (PULL_ONLY || push(parent)) return;
    console.log(`[STEP] The ${BRANCH} branch changed while syncing, trying again.`);
  }
  throw new Error(`the ${BRANCH} branch kept changing; run npm run pdi:sync again`);
}

try {
  main();
} catch (err) {
  console.log(`[WARN] Could not sync the PDI results with GitHub: ${String(err.message).slice(0, 200)}`);
  // A failed pull warns but doesn't stop the run or report that follows it.
  if (!PULL_ONLY) process.exitCode = 1;
}
