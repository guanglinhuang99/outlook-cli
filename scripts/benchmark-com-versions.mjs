// Read-only real-Outlook comparison. Logs timing/counts, never mail data or IDs.
// node scripts/benchmark-com-versions.mjs --baseline <checkout> [--legacy <checkout>]
//   --out <new output directory> [--before <ISO timestamp>] [--rounds 7]
import assert from 'node:assert/strict';
import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';
import { parseArgs } from 'node:util';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';
import { createHash } from 'node:crypto';

const { values } = parseArgs({ options: {
  comparison: { type: 'string' }, baseline: { type: 'string' }, legacy: { type: 'string' }, out: { type: 'string' },
  before: { type: 'string' }, rounds: { type: 'string', default: '7' }, 'include-indexed': { type: 'boolean', default: false }
} });
if (!values.baseline || !values.out) throw Error('--baseline and --out are required');
const rounds = Number(values.rounds);
if (!Number.isSafeInteger(rounds) || rounds < 1 || rounds > 100) throw Error('rounds must be 1..100');
const output = path.resolve(values.out);
// Require a fresh directory so older reports/indexes are never overwritten.
await mkdir(output);
const current = fileURLToPath(new URL('..', import.meta.url));
const roots = { baseline: path.resolve(values.baseline), ...(values.legacy ? { legacy: path.resolve(values.legacy) } : {}), current,
  ...(values['include-indexed'] ? { indexed: current } : {}) };
if (values.comparison) roots.comparison = path.resolve(values.comparison);
const versions = Object.keys(roots);
const loaded = {};
const report = { started_at: new Date().toISOString(), node: process.version, platform: process.platform,
  rounds, before: values.before ?? new Date().toISOString(), versions: {}, scenarios: [],
  method: 'Same live COM mailbox and fixed before; scan_limit=50, limit=200; warmed workers; alternating/reversed and rotating version order; medians; bridge+validation timing, excluding CLI/HTTP/MCP startup.' };
const save = () => writeFile(path.join(output, 'results.json'), JSON.stringify(report, null, 2));
const median = numbers => [...numbers].sort((a, b) => a - b)[Math.floor(numbers.length / 2)];
const brief = result => ({ count: result.messages?.length, scanned: result.scanned,
  next_offset: result.next_offset, fulltext: result.fulltext,
  incomplete: result.body_search_incomplete ?? false, retry: result.retry_required ?? false });
const comparable = message => Object.fromEntries(['id', 'subject', 'sender', 'received_at', 'unread',
  'size_bytes', 'attachment_count', 'download_state'].map(key => [key, message[key]]));
function equal(a, b) {
  if (a.messages) assert.deepEqual(a.messages.map(comparable), b.messages.map(comparable));
  if (a.body) assert.deepEqual(a.body, b.body);
  for (const key of ['scanned', 'next_offset', 'complete']) assert.equal(a[key], b[key]);
  assert.ok(!a.retry_required && !b.retry_required && !a.body_search_incomplete && !b.body_search_incomplete);
}
const call = (version, method, params) => loaded[version].bridge.call(method, loaded[version].validateCall(method,
  version === 'indexed' && method === 'search' ? { ...params, com_body_mode: 'indexed' } : params));
async function scenario(name, method, params, count = rounds, warmup = true) {
  if (warmup) for (const version of versions) await call(version, method, params);
  const samples = Object.fromEntries(versions.map(version => [version, []]));
  let last;
  for (let i = 0; i < count; i++) {
    let order = versions.slice(i % versions.length).concat(versions.slice(0, i % versions.length));
    if (i % 2) order.reverse();
    const results = {};
    for (const version of order) {
      loaded[version].stages = {};
      const begin = performance.now();
      results[version] = await call(version, method, params);
      samples[version].push({ ms: performance.now() - begin, stages_ms: structuredClone(loaded[version].stages) });
    }
    for (const version of versions) equal(results.baseline, results[version]);
    last = Object.fromEntries(versions.map(version => [version, brief(results[version])]));
    console.log(JSON.stringify({ scenario: name, round: i + 1, rounds: count, results_equal: true }));
  }
  const medians = Object.fromEntries(versions.map(version => [version, median(samples[version].map(sample => sample.ms))]));
  const comparisons = Object.fromEntries(versions.filter(version => version !== 'current').map(version => [version, {
    latency_reduction_percent: (medians[version] - medians.current) / medians[version] * 100,
    speedup: medians[version] / medians.current
  }]));
  const result = { name, rounds: count, median_ms: medians, comparisons, results_equal: true, last, samples };
  report.scenarios.push(result);
  await save();
  console.log(JSON.stringify({ ...result, samples: undefined }));
}
try {
  for (const [name, root] of Object.entries(roots)) {
    const moduleAt = relative => import(pathToFileURL(path.join(root, relative)).href);
    const { OutlookBridge } = await moduleAt('src/bridge.mjs');
    const { validateCall } = await moduleAt('src/tools.mjs');
    const packageInfo = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
    let index;
    if (name !== 'legacy') {
      const { WindowsFulltextIndex } = await moduleAt('src/windows-fulltext-index.mjs');
      index = new WindowsFulltextIndex({ directory: path.join(output, `${name}-index`) });
    }
    const bridge = new OutlookBridge({ fulltextIndex: index });
    const entry = { bridge, index, validateCall, stages: {} };
    loaded[name] = entry;
    const direct = bridge.call.bind(bridge);
    bridge.call = async (method, ...args) => {
      const start = performance.now();
      try { return await direct(method, ...args); }
      finally { (entry.stages[method] ??= []).push(performance.now() - start); }
    };
    if (index) {
      const evaluate = index.evaluate.bind(index);
      index.evaluate = async (...args) => {
        const start = performance.now();
        try { return await evaluate(...args); }
        finally { (entry.stages.sqlite_evaluate ??= []).push(performance.now() - start); }
      };
      if (index.prepareColdRead) {
        const prepare = index.prepareColdRead.bind(index);
        index.prepareColdRead = async (...args) => {
          const start = performance.now();
          try { return await prepare(...args); }
          finally { (entry.stages.sqlite_prepare_cold ??= []).push(performance.now() - start); }
        };
      }
      if (index.beginColdRead) {
        const begin = index.beginColdRead.bind(index);
        index.beginColdRead = async (...args) => {
          const start = performance.now();
          const cold = await begin(...args);
          if (cold) cold.ready = cold.ready.finally(() => {
            (entry.stages.cache_preparation_overlapped ??= []).push(performance.now() - start);
          });
          return cold;
        };
      }
    }
    report.versions[name] = { root, version: packageInfo.version, search_mode: name === 'indexed' ? 'indexed' : 'version-default',
      worker_sha256: createHash('sha256').update(await readFile(path.join(root, 'powershell/worker.ps1'))).digest('hex') };
    report.versions[name].source_sha256 = {};
    for (const relative of ['powershell/worker.ps1', 'src/bridge.mjs', 'src/windows-fulltext.mjs', 'src/windows-fulltext-index.mjs']) {
      try { report.versions[name].source_sha256[relative] = createHash('sha256').update(await readFile(path.join(root, relative))).digest('hex'); }
      catch (error) { if (error.code !== 'ENOENT') throw error; }
    }
    const health = await call(name, 'health', {});
    assert.equal(health.connected, true);
    report.versions[name].outlook_version = health.outlook_version;
  }
  const scope = { folder: 'inbox', before: report.before, scan_limit: 50, limit: 200 };
  const messages = await call('baseline', 'messages', scope);
  for (const version of versions) equal(messages, await call(version, 'messages', scope));
  assert.ok(messages.messages.length);
  report.mail_count = messages.messages.length;
  report.unread_before = (await call('current', 'folders', { folder: 'inbox', limit: 1 })).parent.unread_count;
  await scenario('body_first_build_50', 'search', { ...scope, query: '__benchmark_no_match_720d4e__', search_body: true }, 1, false);
  // Charge deferred persistence to the immediately following query, without
  // a warmup that would hide the cost. Legacy simply scans again.
  await scenario('body_second_query_50', 'search', { ...scope, query: '__benchmark_no_match_720d4e__', search_body: true }, 1, false);
  await scenario('messages_25', 'messages', { ...scope, limit: 25 });
  await scenario('read_20000_chars', 'read', { id: messages.messages[0].id, max_chars: 20000 });
  await scenario('metadata_search_50', 'search', { ...scope, query: '__benchmark_no_match_720d4e__' });
  for (const [name, query] of [['no_match', '__benchmark_no_match_720d4e__'], ['english', 'the'], ['chinese', '项目']]) {
    await scenario(`body_warm_${name}_50`, 'search', { ...scope, query, search_body: true });
  }
  report.unread_after = (await call('current', 'folders', { folder: 'inbox', limit: 1 })).parent.unread_count;
} catch (error) {
  // Assertions contain actual mailbox values: never log their message/stack.
  report.error = { code: error.code ?? 'BENCHMARK_FAILED', message: 'Benchmark did not finish; no performance claim for incomplete scenarios.' };
  process.exitCode = 1;
  console.log(JSON.stringify({ error: report.error }));
} finally {
  for (const { bridge, index } of Object.values(loaded)) { bridge.close(); index?.close(); }
  report.finished_at = new Date().toISOString();
  await save();
}
