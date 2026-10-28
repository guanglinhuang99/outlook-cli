// Cold-only supplement to benchmark-com-versions.mjs. Every sample gets a
// fresh worker and private index, while health/messages warmup stays identical.
// Logs only timings, counts and source hashes; never mailbox content or IDs.
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';
import { parseArgs } from 'node:util';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import path from 'node:path';

const { values } = parseArgs({ options: { comparison: { type: 'string' }, baseline: { type: 'string' }, legacy: { type: 'string' },
  out: { type: 'string' }, before: { type: 'string' }, rounds: { type: 'string', default: '7' }, 'include-indexed': { type: 'boolean', default: false } } });
if (!values.baseline || !values.legacy || !values.out || !values.before) throw Error('baseline, legacy, out and before are required');
const rounds = Number(values.rounds);
if (!Number.isSafeInteger(rounds) || rounds < 1 || rounds > 30) throw Error('rounds must be 1..30');
const output = path.resolve(values.out);
await mkdir(output);
const roots = { baseline: path.resolve(values.baseline), legacy: path.resolve(values.legacy),
  current: fileURLToPath(new URL('..', import.meta.url)) };
if (values.comparison) roots.comparison = path.resolve(values.comparison);
if (values['include-indexed']) roots.indexed = roots.current;
const versions = Object.keys(roots);
const loaded = {};
const scope = { folder: 'inbox', before: values.before, scan_limit: 50, limit: 200 };
const params = { ...scope, query: '__benchmark_no_match_720d4e__', search_body: true };
const searchParams = version => loaded[version].validateCall('search', version === 'indexed' ? { ...params, com_body_mode: 'indexed' } : params);
const report = { started_at: new Date().toISOString(), node: process.version, platform: process.platform,
  rounds, before: values.before, versions: {}, samples: [],
  method: 'Fresh worker and index for every sample; health/messages warmup outside timer; fixed before, scan_limit=50, limit=200; rotating/reversed version order; all cache preparation inside search timer. First query timed unchanged; second query timed separately without warmup to expose deferred persistence.' };
const save = () => writeFile(path.join(output, 'results.json'), JSON.stringify(report, null, 2));
const comparable = message => Object.fromEntries(['id', 'subject', 'sender', 'received_at', 'unread', 'size_bytes', 'attachment_count', 'download_state'].map(key => [key, message[key]]));
try {
  for (const [version, root] of Object.entries(roots)) {
    const moduleAt = relative => import(pathToFileURL(path.join(root, relative)).href);
    loaded[version] = { ...(await moduleAt('src/bridge.mjs')), ...(await moduleAt('src/tools.mjs')),
      ...(version === 'legacy' ? {} : await moduleAt('src/windows-fulltext-index.mjs')) };
    const hashes = {};
    for (const relative of ['powershell/worker.ps1', 'src/bridge.mjs', 'src/windows-fulltext.mjs', 'src/windows-fulltext-index.mjs']) {
      try { hashes[relative] = createHash('sha256').update(await readFile(path.join(root, relative))).digest('hex'); }
      catch (error) { if (error.code !== 'ENOENT') throw error; }
    }
    report.versions[version] = { root, search_mode: version === 'indexed' ? 'indexed' : 'version-default', source_sha256: hashes };
  }
  for (let round = 0; round < rounds; round++) {
    let order = versions.slice(round % versions.length).concat(versions.slice(0, round % versions.length));
    if (round % 2) order.reverse();
    const entries = {};
    const results = {};
    const sample = { round: round + 1, order, versions: {} };
    try {
      let reference;
      for (const version of order) {
        report.active_operation = { round: round + 1, version, phase: 'warmup' };
        const api = loaded[version];
        const index = version === 'legacy' ? undefined : new api.WindowsFulltextIndex({ directory: path.join(output, `${round + 1}-${version}-index`) });
        const bridge = new api.OutlookBridge({ fulltextIndex: index });
        const entry = { bridge, index, stages: {} };
        entries[version] = entry;
        const direct = bridge.call.bind(bridge);
        bridge.call = async (method, ...args) => {
          const start = performance.now();
          try { return await direct(method, ...args); }
          finally { (entry.stages[method] ??= []).push(performance.now() - start); }
        };
        for (const method of ['evaluate', 'prepareColdRead']) if (index?.[method]) {
          const original = index[method].bind(index);
          index[method] = async (...args) => {
            const start = performance.now();
            try { return await original(...args); }
            finally { (entry.stages[`sqlite_${method}`] ??= []).push(performance.now() - start); }
          };
        }
        if (index?.evaluateDirect) {
          const evaluate = index.evaluateDirect.bind(index);
          index.evaluateDirect = (...args) => {
            const start = performance.now();
            try { return evaluate(...args); }
            finally { (entry.stages.direct_evaluate ??= []).push(performance.now() - start); }
          };
        }
        if (index?.beginColdRead) {
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
        const health = await bridge.call('health', api.validateCall('health', {}));
        assert.equal(health.connected, true);
        const messages = await bridge.call('messages', api.validateCall('messages', scope));
        if (reference) assert.deepEqual(messages.messages.map(comparable), reference);
        else reference = messages.messages.map(comparable);
        report.versions[version].outlook_version = health.outlook_version;
      }
      for (const version of order) {
        report.active_operation = { round: round + 1, version, phase: 'first-search' };
        const entry = entries[version];
        entry.stages = {};
        const start = performance.now();
        const result = await entry.bridge.call('search', searchParams(version));
        const ms = performance.now() - start;
        assert.ok(!result.retry_required && !result.body_search_incomplete);
        if (version !== 'legacy') {
          assert.equal(result.fulltext.reused, 0);
          if (result.fulltext.engine === 'direct-scan') {
            assert.equal(result.fulltext.indexed, 0);
            assert.equal(result.fulltext.updated, 0);
            assert.equal(result.fulltext.evaluated, result.fulltext.candidates);
            assert.ok(['deferred', 'disabled'].includes(result.fulltext.index_state));
          } else assert.equal(result.fulltext.updated, result.fulltext.candidates);
          assert.equal(result.fulltext.complete, true);
        }
        results[version] = result;
        sample.versions[version] = { ms, stages_ms: entry.stages, count: result.messages.length, scanned: result.scanned, fulltext: result.fulltext };
      }
      for (const result of Object.values(results)) {
        assert.deepEqual(result.messages.map(comparable), results.legacy.messages.map(comparable));
        for (const key of ['scanned', 'next_offset', 'complete']) assert.equal(result[key], results.legacy[key]);
      }
      sample.results_equal = true;
      sample.second_versions = {};
      for (const version of order) {
        report.active_operation = { round: round + 1, version, phase: 'second-search' };
        const entry = entries[version];
        entry.stages = {};
        const start = performance.now();
        const result = await entry.bridge.call('search', searchParams(version));
        const ms = performance.now() - start;
        assert.ok(!result.retry_required && !result.body_search_incomplete);
        assert.deepEqual(result.messages.map(comparable), results.legacy.messages.map(comparable));
        for (const key of ['scanned', 'next_offset', 'complete']) assert.equal(result[key], results.legacy[key]);
        if (results[version].fulltext?.index_state === 'deferred') {
          assert.equal(result.fulltext.updated, result.fulltext.candidates);
          assert.equal(result.fulltext.staged_reused, result.fulltext.candidates);
        } else if (results[version].fulltext?.index_state === 'disabled') {
          assert.equal(result.fulltext.engine, 'direct-scan');
          assert.equal(result.fulltext.indexed, 0);
          assert.equal(result.fulltext.evaluated, result.fulltext.candidates);
          assert.equal(result.fulltext.complete, true);
        }
        sample.second_versions[version] = { ms, stages_ms: entry.stages, fulltext: result.fulltext };
      }
      report.samples.push(sample);
      delete report.active_operation;
      await save();
      console.log(JSON.stringify({ round: round + 1, results_equal: true, ms: Object.fromEntries(versions.map(version => [version, sample.versions[version].ms])) }));
    } finally {
      for (const { bridge, index } of Object.values(entries)) { bridge.close(); index?.close(); }
    }
  }
  const median = numbers => [...numbers].sort((a, b) => a - b)[Math.floor(numbers.length / 2)];
  report.median_ms = Object.fromEntries(versions.map(version => [version, median(report.samples.map(sample => sample.versions[version].ms))]));
  report.second_median_ms = Object.fromEntries(versions.map(version => [version, median(report.samples.map(sample => sample.second_versions[version].ms))]));
  report.first_two_total_median_ms = Object.fromEntries(versions.map(version => [version, median(report.samples.map(sample => sample.versions[version].ms + sample.second_versions[version].ms))]));
  report.comparisons = Object.fromEntries(versions.filter(version => version !== 'current').map(version => [version, {
    latency_reduction_percent: (report.median_ms[version] - report.median_ms.current) / report.median_ms[version] * 100,
    speedup: report.median_ms[version] / report.median_ms.current
  }]));
  console.log(JSON.stringify({ median_ms: report.median_ms, second_median_ms: report.second_median_ms, first_two_total_median_ms: report.first_two_total_median_ms, comparisons: report.comparisons }));
} catch (error) {
  report.error = { code: error.code ?? 'BENCHMARK_FAILED', message: 'Cold benchmark failed; do not use incomplete samples as performance evidence.',
    ...(error.code === 'FULLTEXT_SOURCE_CHANGED' ? { source_check: error.message } : {}) };
  process.exitCode = 1;
  console.log(JSON.stringify(report.error));
} finally { report.finished_at = new Date().toISOString(); await save(); }
