// Shares a VIN master update (pdi/Vin_Master_update.spec.ts) between this PC and GitHub jobs
// (.github/workflows/vin-master.yml) through the vin-master-runs branch, to finish sooner.
//
//   node "pdi/scripts/vinMasterGitHub.js" --split <jobs> --keep <n>
//       The VINs still to do (same rule as the spec: not checked yet, ended in a retry status, or the sheet has
//       changed since) in sheet order: the first <n> stay on this PC (Output/VIN Master/VIN_Master_todo_pc.txt,
//       run them with VIN_LIST set to that file), the rest go to the branch as todo/chunk-<i>.txt, one per job.
//       Replaces the branch's todo lists and clears its old results.
//   node "pdi/scripts/vinMasterGitHub.js" --merge
//       Adds the results the GitHub jobs sent (results/chunk-<i>.jsonl) to this PC's progress file. Lines already
//       there are skipped, so merging again is harmless; the spec reads the latest line per VIN.
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const BRANCH = 'vin-master-runs';
const REMOTE_REF = `refs/remotes/origin/${BRANCH}`;
const ROOT = path.join(__dirname, '..', '..');
const MASTER_DIR = path.join(ROOT, 'Output', 'VIN Master');
const PROGRESS_FILE = path.join(MASTER_DIR, 'VIN_Master_progress.jsonl');
const PC_TODO = path.join(MASTER_DIR, 'VIN_Master_todo_pc.txt');
const VIN_FILE = process.env.PDI_VIN_FILE || path.join(ROOT, 'resources', 'VIN_Details.csv');
// Same as RETRY in Vin_Master_update.spec.ts.
const RETRY = /^(Error|Not attempted)|Show did not load|must be unique/i;

const git = (args, env) => execFileSync('git', args, { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, ...env }, maxBuffer: 64 << 20 }).trim();
const arg = (name) => { const i = process.argv.indexOf(name); return i > 0 ? process.argv[i + 1] : undefined; };

function fetchBranch() {
  try {
    git(['fetch', '--quiet', 'origin', `+${BRANCH}:${REMOTE_REF}`]);
    return git(['rev-parse', REMOTE_REF]);
  } catch { return null; }
}

function readProgress() {
  const done = new Map();
  if (!fs.existsSync(PROGRESS_FILE)) return done;
  for (const line of fs.readFileSync(PROGRESS_FILE, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    try { const r = JSON.parse(line); done.set(r.ChassisNo, r); } catch { /* a line cut off by a stop */ }
  }
  return done;
}

// VIN, BIN (battery), Motor No, Charger Number, as pushVinToGitHub.gs writes them; quoted cells are rare
// (notes with commas), so a small CSV reader is enough.
function readSheet() {
  const rows = [];
  const text = fs.readFileSync(VIN_FILE, 'utf8');
  let row = [], cell = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (c === '"') quoted = false; else cell += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\n') { row.push(cell.replace(/\r$/, '')); rows.push(row); row = []; cell = ''; }
    else cell += c;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  const head = rows.shift().map((h) => h.trim().toLowerCase());
  const col = (re) => head.findIndex((h) => re.test(h));
  const c = { vin: col(/^vin$/), battery: col(/^(bin|battery)/), motor: col(/^motor/), charger: col(/^charger/) };
  if (Object.values(c).some((i) => i < 0)) throw new Error(`${path.basename(VIN_FILE)}: VIN / BIN / Motor No / Charger columns not found in row 1`);
  return rows.map((r) => ({ vin: (r[c.vin] || '').trim().toUpperCase(), motor: (r[c.motor] || '').trim(), battery: (r[c.battery] || '').trim(), charger: (r[c.charger] || '').trim() }))
    .filter((r) => /^[A-Z0-9]{17}$/.test(r.vin));
}

// One commit on top of parent with the given files (path -> content, null = remove), pushed; false if someone pushed first.
function commit(parent, files, message) {
  const index = path.join(os.tmpdir(), `vin-master-index-${process.pid}`);
  const env = { GIT_INDEX_FILE: index };
  try {
    if (parent) git(['read-tree', parent], env); else git(['read-tree', '--empty'], env);
    for (const [file, content] of Object.entries(files)) {
      if (content === null) { git(['rm', '--cached', '--quiet', '-r', '--ignore-unmatch', file], env); continue; }
      const tmp = path.join(os.tmpdir(), `vin-master-blob-${process.pid}`);
      fs.writeFileSync(tmp, content);
      const blob = git(['hash-object', '-w', tmp]);
      fs.rmSync(tmp, { force: true });
      git(['update-index', '--add', '--cacheinfo', `100644,${blob},${file}`], env);
    }
    const who = { GIT_AUTHOR_NAME: 'VIN master', GIT_AUTHOR_EMAIL: 'vin-master@users.noreply.github.com', GIT_COMMITTER_NAME: 'VIN master', GIT_COMMITTER_EMAIL: 'vin-master@users.noreply.github.com' };
    const sha = git(['commit-tree', git(['write-tree'], env), ...(parent ? ['-p', parent] : []), '-m', message], who);
    try {
      git(['push', '--quiet', 'origin', `${sha}:refs/heads/${BRANCH}`]);
    } catch (err) {
      if (/rejected|non-fast-forward|fetch first/i.test(String(err.stderr || err.message))) return false;
      throw err;
    }
    return true;
  } finally {
    fs.rmSync(index, { force: true });
  }
}

function split(jobs, keep) {
  const progress = readProgress();
  const todo = readSheet().filter((d) => {
    const r = progress.get(d.vin);
    return !r || RETRY.test(r.Status || '') || r.SheetMotor !== d.motor || r.SheetBattery !== d.battery || r.SheetCharger !== d.charger;
  }).map((d) => d.vin);
  const pc = todo.slice(0, keep);
  const rest = todo.slice(keep);
  const size = Math.ceil(rest.length / jobs);
  const files = { todo: null, results: null };
  for (let i = 0; i < jobs; i++) {
    const chunk = rest.slice(i * size, (i + 1) * size);
    if (chunk.length) files[`todo/chunk-${i + 1}.txt`] = `${chunk.join('\n')}\n`;
  }
  for (let attempt = 1; attempt <= 3; attempt++) {
    if (commit(fetchBranch(), files, `VIN master split: ${pc.length} on the PC, ${rest.length} in ${Object.keys(files).length - 2} GitHub job(s)`)) break;
    if (attempt === 3) throw new Error(`Could not push to ${BRANCH}, try again.`);
  }
  fs.mkdirSync(MASTER_DIR, { recursive: true });
  fs.writeFileSync(PC_TODO, `${pc.join('\n')}\n`);
  console.log(`[STEP] ${todo.length} VINs to do: ${pc.length} on this PC (${path.relative(ROOT, PC_TODO)}), ` +
    `${rest.length} in ${Object.keys(files).length - 2} GitHub chunk(s) of up to ${size} on the ${BRANCH} branch.`);
}

function merge() {
  const parent = fetchBranch();
  if (!parent) { console.log(`[STEP] No ${BRANCH} branch yet, nothing to merge.`); return; }
  const files = git(['ls-tree', '--name-only', parent, 'results/']).split('\n').filter(Boolean);
  const have = new Set(fs.existsSync(PROGRESS_FILE) ? fs.readFileSync(PROGRESS_FILE, 'utf8').split('\n').filter((l) => l.trim()) : []);
  let added = 0;
  const lines = [];
  for (const f of files) {
    for (const line of git(['show', `${parent}:${f}`]).split('\n')) {
      if (!line.trim() || have.has(line)) continue;
      try { JSON.parse(line); } catch { continue; } // a line cut off mid-send
      have.add(line);
      lines.push(line);
      added++;
    }
  }
  if (lines.length) {
    fs.mkdirSync(MASTER_DIR, { recursive: true });
    fs.appendFileSync(PROGRESS_FILE, `${lines.join('\n')}\n`);
  }
  console.log(`[STEP] Merged ${added} new result(s) from ${files.length} GitHub chunk(s) into ${path.basename(PROGRESS_FILE)}.`);
}

module.exports = { merge };

if (require.main === module) {
  if (process.argv.includes('--split')) {
    const jobs = Number(arg('--split'));
    const keep = Number(arg('--keep') || 0);
    if (!Number.isInteger(jobs) || jobs < 1 || !Number.isInteger(keep) || keep < 0) throw new Error('Usage: --split <jobs> --keep <VINs for this PC>');
    split(jobs, keep);
  } else if (process.argv.includes('--merge')) {
    merge();
  } else {
    console.log('Usage: node "pdi/scripts/vinMasterGitHub.js" --split <jobs> --keep <n> | --merge');
  }
}
