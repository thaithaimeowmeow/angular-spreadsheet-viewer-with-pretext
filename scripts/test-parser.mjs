// Runs the real parser worker in Node against spreadsheet files and checks the output.
// Usage: node scripts/test-parser.mjs <file-or-dir>... [--ref <path-to-sheetjs-module>]
//   - If "<file>.expected.json" exists, sheet 1 must match it exactly.
//   - Otherwise, with --ref, every sheet is compared to SheetJS (used only as a test oracle).
import { build } from 'esbuild';
import { existsSync, readdirSync, statSync, readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const args = process.argv.slice(2);
const refIdx = args.indexOf('--ref');
const refPath = refIdx >= 0 ? args.splice(refIdx, 2)[1] : null;
const XLSX = refPath ? createRequire(import.meta.url)(refPath) : null;

const files = args.flatMap((p) =>
  statSync(p).isDirectory()
    ? readdirSync(p).filter((f) => !f.endsWith('.json')).map((f) => join(p, f))
    : [p]);

const out = join(process.env.TEMP ?? '.', `parser-worker-${process.pid}.mjs`);
await build({
  entryPoints: [fileURLToPath(new URL('../src/app/parser/parser.worker.ts', import.meta.url))],
  bundle: true, format: 'esm', platform: 'neutral', outfile: out, logLevel: 'error',
});

let handler;
let sink = () => {};
globalThis.addEventListener = (_type, fn) => { handler = fn; };
globalThis.postMessage = (msg) => sink(msg);
await import(pathToFileURL(out).href);

let job = 0;
const logs = new Map(); // job -> messages
let wake = () => {};
sink = (m) => {
  if (!logs.has(m.job)) logs.set(m.job, []);
  logs.get(m.job).push(m);
  wake();
};
/** Resolves with the job's log once a message matching `pred` has arrived (or rejects after 20 s). */
function waitFor(id, pred) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('timeout waiting on job ' + id)), 20000);
    const check = () => {
      const log = logs.get(id) ?? [];
      if (log.some(pred)) {
        clearTimeout(timer);
        wake = () => {};
        resolve(log);
      } else {
        wake = check;
      }
    };
    check();
  });
}
function run(msg, until) {
  logs.set(msg.job, []);
  handler({ data: msg });
  return waitFor(msg.job, (m) => until.includes(m.type));
}
const query = (msg, type) => {
  handler({ data: msg });
  return waitFor(msg.job, (m) => m.type === type && m.reqId === msg.reqId).then((log) => log.find((m) => m.type === type && m.reqId === msg.reqId));
};

let failures = 0;
for (const path of files) {
  const bytes = readFileSync(path);
  const file = new File([bytes], basename(path));
  const t0 = performance.now();
  const meta = (await run({ type: 'open', file, job: ++job }, ['meta', 'error'])).at(-1);
  if (meta.type === 'error') {
    failures++;
    console.log(`FAIL  ${basename(path)}: ${meta.message}`);
    continue;
  }
  const expectedPath = path + '.expected.json';
  const expected = existsSync(expectedPath) ? JSON.parse(readFileSync(expectedPath, 'utf8')) : null;
  const ref = !expected && XLSX ? XLSX.read(bytes, { type: 'buffer', cellDates: false }) : null;
  const parts = [];
  let fileOk = true;
  for (let s = 0; s < meta.sheets.length; s++) {
    const sid = ++job;
    const msgs = await run({ type: 'sheet', index: s, job: sid, wrap: true }, ['done', 'error']);
    const end = msgs.find((m) => m.type === 'done' || m.type === 'error');
    if (end.type === 'error') { fileOk = false; parts.push(`${meta.sheets[s]}: ERROR ${end.message}`); continue; }
    // Rows stay in the worker: fetch them like the UI does.
    const win = await query({ type: 'window', job: sid, reqId: 1, start: 0, end: end.rowCount }, 'window');
    const rows = win.rows;
    // Protocol checks: columns before the first grow, estimates cover every row, layoutDone.
    const log = await waitFor(sid, (m) => m.type === 'layoutDone');
    const firstGrow = log.findIndex((m) => m.type === 'grow');
    const columnsAt = log.findIndex((m) => m.type === 'columns');
    const covered = new Uint8Array(end.rowCount);
    for (const m of log) if (m.type === 'heights' && m.version === 0) m.heights.forEach((h, i) => { if (h >= 25) covered[m.start + i] = 1; });
    const protocolErrors = [];
    if (columnsAt < 0 || (firstGrow >= 0 && columnsAt > firstGrow)) protocolErrors.push('columns not before first grow');
    if (covered.some((v) => !v)) protocolErrors.push(`${covered.filter((v) => !v).length} rows without height estimate`);
    if (rows.length !== end.rowCount) protocolErrors.push(`window returned ${rows.length} of ${end.rowCount} rows`);
    // Copy round-trip of the first cell; search for its first word.
    const first = rows.flat().find((v) => typeof v === 'string' && /\p{L}{3,}/u.test(v));
    if (first) {
      const c0 = rows.findIndex((r) => r?.includes(first));
      const copy = await query({ type: 'copy', job: sid, reqId: 2, r0: c0, r1: c0, c0: rows[c0].indexOf(first), c1: rows[c0].indexOf(first) }, 'copy');
      const expectText = /[\t\n\r"]/.test(first) ? `"${first.replace(/"/g, '""')}"` : first;
      if (copy.text !== expectText) protocolErrors.push('copy text mismatch');
      const word = first.match(/\p{L}{3,}/u)[0];
      handler({ data: { type: 'search', job: sid, searchId: 7, query: word } });
      const hits = await waitFor(sid, (m) => m.type === 'searchHits' && m.searchId === 7 && m.done);
      const n = hits.filter((m) => m.type === 'searchHits' && m.searchId === 7).reduce((a, m) => a + m.keys.length, 0);
      if (!n) protocolErrors.push(`search "${word}" found nothing`);
    }
    if (protocolErrors.length) { fileOk = false; parts.push(`${meta.sheets[s]}: PROTOCOL ${protocolErrors.join('; ')}`); }
    const filled = rows.reduce((n, r) => n + (r ? r.filter((v) => v !== undefined && v !== '').length : 0), 0);
    let note = '';
    if (expected && s === 0) note = compareExpected(expected, rows);
    else if (ref) note = compareRef(ref.Sheets[ref.SheetNames[s]], rows);
    if (note.startsWith('MISMATCH')) fileOk = false;
    parts.push(`${meta.sheets[s]}: ${end.rowCount}×${end.colCount}, ${filled} cells ${note}`);
  }
  if (!fileOk) failures++;
  console.log(`${fileOk ? 'OK  ' : 'FAIL'}  ${basename(path)} [${meta.format}] (${(performance.now() - t0).toFixed(0)} ms)\n        ${parts.join('\n        ')}`);
}
console.log(`\n${files.length - failures}/${files.length} files OK`);
process.exit(failures ? 1 : 0);

/** Compares against SheetJS raw values: strings exactly, numbers numerically, dates by day. */
function compareRef(ws, rows) {
  if (!ws || !ws['!ref']) return rows.some((r) => r?.some((v) => v)) ? 'MISMATCH ref empty' : '[ref ✓]';
  const range = XLSX.utils.decode_range(ws['!ref']);
  const diffs = [];
  let checked = 0;
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const ours = rows[r]?.[c] ?? '';
      if (!cell || cell.v === undefined || cell.v === '') {
        if (ours !== '') diffs.push(`${addr}: ref empty, ours ${JSON.stringify(ours)}`);
        continue;
      }
      checked++;
      let ok;
      if (cell.t === 'n') {
        if (ours.endsWith('%')) ok = Math.abs(parseFloat(ours) / 100 - cell.v) < 1e-9;
        else if (/^\d{4}-\d\d-\d\d/.test(ours)) ok = ours.slice(0, 10) === new Date(Math.round((cell.v - 25569) * 86400) * 1000).toISOString().slice(0, 10);
        else if (/^\d\d:\d\d/.test(ours)) ok = true;
        else ok = Math.abs(Number(ours) - cell.v) <= Math.abs(cell.v) * 1e-14;
      } else if (cell.t === 'b') ok = ours === (cell.v ? 'TRUE' : 'FALSE');
      else if (cell.t === 'e') ok = ours === cell.w;
      else if (cell.t === 'd') ok = ours.slice(0, 10) === cell.v.toISOString().slice(0, 10);
      else ok = ours === String(cell.v).replace(/\r\n?/g, '\n');
      if (!ok && cell.w !== undefined && ours === cell.w) ok = true; // formatted text (e.g. ODS display)
      if (!ok) diffs.push(`${addr}: ref ${JSON.stringify(cell.v)}, ours ${JSON.stringify(ours)}`);
    }
  }
  return diffs.length ? `MISMATCH ${diffs.length}/${checked}: ${diffs.slice(0, 3).join(' | ')}` : `[ref ✓ ${checked} cells]`;
}

/** Exact comparison against hand-written expected rows (hole / null / undefined = ''). */
function compareExpected(expected, rows) {
  const diffs = [];
  const n = Math.max(expected.length, rows.length);
  for (let r = 0; r < n; r++) {
    const a = expected[r] ?? [];
    const b = rows[r] ?? [];
    for (let c = 0; c < Math.max(a.length, b.length); c++) {
      const x = a[c] ?? '';
      const y = b[c] ?? '';
      if (x !== y) diffs.push(`R${r + 1}C${c + 1}: expected ${JSON.stringify(x)}, ours ${JSON.stringify(y)}`);
    }
  }
  return diffs.length ? `MISMATCH ${diffs.length}: ${diffs.slice(0, 4).join(' | ')}` : '[expected ✓]';
}
