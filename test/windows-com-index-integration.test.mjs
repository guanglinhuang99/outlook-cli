import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { EventEmitter } from 'node:events';
import { PassThrough, Writable } from 'node:stream';
import { OutlookBridge } from '../src/bridge.mjs';
import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';
import { validateCall } from '../src/tools.mjs';
import { searchWindowsFulltext } from '../src/windows-fulltext.mjs';

test('explicit COM direct mode repeatedly scans even with an unusable cache path', async t => {
  const root = await mkdtemp(path.join(tmpdir(), 'direct-default-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const directory = path.join(root, 'blocked');
  await writeFile(directory, 'not a cache directory');
  const index = new WindowsFulltextIndex({ directory });
  let scans = 0, proofs = 0;
  const bridge = new OutlookBridge({ fulltextIndex: index, spawnWorker: () => {
    const worker = new EventEmitter(); worker.stdout = new PassThrough(); worker.stderr = new PassThrough();
    worker.stdin = new Writable({ write(chunk, encoding, callback) {
      const request = JSON.parse(chunk.toString());
      let result;
      if (request.method === 'fulltext_candidates') {
        scans++;
        if (request.params.com_body_mode !== 'indexed') {
          assert.equal(request.params.direct_scan, true);
          assert.equal(request.params.direct_match_query, 'hit');
        }
        const signature = 'v' + scans;
        result = { messages: [{ id: 'a', subject: 'other', _fulltext_signature: signature }], next_offset: null,
          validation: { candidates: [{ id: 'a', signature }] },
          inline_bodies: [{ id: 'a', signature, status: request.params.direct_scan ? 'unverified' : 'ready', text: scans === 1 ? 'hit' : 'changed' }] };
      } else { assert.equal(request.method, 'fulltext_validate'); proofs++; result = { valid: true }; }
      worker.stdout.write(JSON.stringify({ id: request.id, result }) + '\n'); callback();
    } });
    worker.kill = () => { worker.stdout.end(); worker.stderr.end(); worker.emit('exit', 0); }; return worker;
  } });
  t.after(() => bridge.close());
  for (const mode of ['direct', 'direct', 'direct']) {
    const result = await bridge.call('search', validateCall('search', { query: 'HIT', search_body: true, ...(mode ? { com_body_mode: mode } : {}) }));
    assert.equal(result.fulltext.engine, 'direct-scan'); assert.equal(result.fulltext.index_state, 'disabled');
    assert.equal(result.messages.length, scans === 1 ? 1 : 0);
  }
  assert.equal(scans, 3); assert.equal(proofs, 3);
  assert.ok((await stat(directory)).isFile());
  await assert.rejects(bridge.call('search', validateCall('search', { query: 'hit', search_body: true, com_body_mode: 'indexed' })));
  assert.equal(scans, 4); assert.equal(proofs, 3);
});

test('direct scans publish nothing without the complete final source proof', async t => {
  const directory = await mkdtemp(path.join(tmpdir(), 'direct-proof-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  for (const failure of ['missing-proof', 'mismatched-proof', 'source-changed', 'false-valid', 'missing-metadata']) {
    const index = new WindowsFulltextIndex({ directory: path.join(directory, failure) });
    t.after(() => index.close());
    const bridge = { call: async method => {
      if (method === 'fulltext_candidates') return {
        messages: [{ id: 'a', _fulltext_signature: 'v1', subject: 'other' }], next_offset: null,
        inline_bodies: [{ id: 'a', signature: 'v1', status: 'unverified', text: 'hit' }],
        fulltext_capabilities: { deferred_metadata: true },
        validation: failure === 'missing-proof' ? undefined : { candidates: [{ id: 'a', signature: failure === 'mismatched-proof' ? 'v2' : 'v1' }] }
      };
      assert.equal(method, 'fulltext_validate');
      if (failure === 'source-changed') throw Object.assign(Error('changed'), { code: 'FULLTEXT_SOURCE_CHANGED' });
      return { valid: failure !== 'false-valid', metadata: failure === 'missing-metadata' ? [] : [{ id: 'a', attachment_count: 0 }] };
    } };
    await assert.rejects(searchWindowsFulltext(bridge, { query: 'hit' }, index, { directScan: true }), { code: 'FULLTEXT_SOURCE_CHANGED' });
    await assert.rejects(stat(index.directory), { code: 'ENOENT' });
  }
});

test('direct queries always reread changed bodies; explicitly indexed queries retain cache protections', async t => {
  const directory = await mkdtemp(path.join(tmpdir(), 'direct-promotion-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  for (const unsafe of [false, true]) {
    const index = new WindowsFulltextIndex({ directory: path.join(directory, String(unsafe)) });
    t.after(() => index.close());
    let version = 'v1', reads = 0;
    const bridge = { call: async (method, params) => {
      if (method === 'fulltext_candidates') return {
        messages: [{ id: 'a', _fulltext_signature: version, subject: 'other' }], next_offset: null,
        ...(params.direct_scan ? { inline_bodies: [{ id: 'a', signature: version, status: 'unverified', text: version === 'v1' ? 'old hit' : 'changed body' }] } : {}),
        fulltext_capabilities: { batch_read: true, deferred_metadata: true },
        validation: { candidates: [{ id: 'a', signature: version }] }
      };
      if (method === 'fulltext_read_batch') {
        reads++;
        return { bodies: [{ id: 'a', signature: version, status: 'ready', text: 'changed body' }] };
      }
      assert.equal(method, 'fulltext_validate');
      return { valid: true, metadata: params.metadata_ids.map(id => ({ id, attachment_count: 0 })) };
    } };
    assert.equal((await searchWindowsFulltext(bridge, { query: 'hit' }, index, { directScan: true })).messages.length, 1);
    version = 'v2';
    if (unsafe) {
      index._ensureDirectory = async () => { throw Object.assign(Error('unsafe'), { code: 'FULLTEXT_UNSAFE_PATH' }); };
      await assert.rejects(searchWindowsFulltext(bridge, { query: 'hit' }, index, { inlineColdRead: true }), { code: 'FULLTEXT_UNSAFE_PATH' });
      assert.equal(reads, 0);
      await assert.rejects(stat(index.directory), { code: 'ENOENT' });
    } else {
      index._ensureDirectory = () => { throw Error('direct mode must not prepare any cache'); };
      const second = await searchWindowsFulltext(bridge, { query: 'hit' }, index, { directScan: true });
      assert.equal(second.messages.length, 0);
      assert.equal(reads, 0);
      assert.equal(second.fulltext.updated, 0);
      assert.equal(second.fulltext.evaluated, 1);
      assert.equal(second.fulltext.index_state, 'disabled');
      await assert.rejects(stat(index.directory), { code: 'ENOENT' });
    }
  }
});

test('cold prefetch overlaps preparation but never evaluates before private cache readiness', async () => {
  let finishPreparation;
  let candidateRead;
  const readStarted = new Promise(resolve => { candidateRead = resolve; });
  const ready = new Promise(resolve => { finishPreparation = resolve; });
  let evaluated = false;
  const index = {
    beginColdRead: async () => ({ budget: { max_bytes: 8, max_body_bytes: 8 }, ready }),
    evaluate: async () => { evaluated = true; return { hits: new Map(), retry_required: false }; }
  };
  const bridge = { call: async () => { candidateRead(); return { messages: [], inline_bodies: [], next_offset: null }; } };
  const pending = searchWindowsFulltext(bridge, { query: 'hit' }, index, { inlineColdRead: true });
  await readStarted;
  assert.equal(evaluated, false);
  finishPreparation(true);
  await pending;
  assert.equal(evaluated, true);
});

test('a failed concurrent cache preparation cannot persist prefetched bodies or publish results', async () => {
  let rejectPreparation;
  const ready = new Promise((resolve, reject) => { rejectPreparation = reject; });
  const index = {
    beginColdRead: async () => ({ budget: { max_bytes: 8, max_body_bytes: 8 }, ready }),
    evaluate: () => { throw Error('must not evaluate after preparation failure'); }
  };
  const bridge = { call: async () => {
    rejectPreparation(Object.assign(Error('private cache unavailable'), { code: 'FULLTEXT_UNSAFE_PATH' }));
    return { messages: [{ id: 'a' }], inline_bodies: [{ id: 'a', text: 'body', status: 'ready' }] };
  } };
  await assert.rejects(searchWindowsFulltext(bridge, { query: 'hit' }, index, { inlineColdRead: true }), { code: 'FULLTEXT_UNSAFE_PATH' });
});

test('inline batches reject duplicate or missing bodies instead of silently losing results', async () => {
  const candidates = [{ id: 'a', _fulltext_signature: 'v1' }, { id: 'b', _fulltext_signature: 'v1' }];
  const index = { prepareColdRead: async () => ({ max_bytes: 8, max_body_bytes: 8 }),
    evaluate: () => { throw Error('invalid inline batch must not reach SQLite'); } };
  for (const ids of [['a'], ['a', 'a'], ['a', 'wrong']]) {
    const bridge = { call: async (method, params) => {
      assert.equal(method, 'fulltext_candidates');
      assert.ok(params.inline_body_budget.budget_ms > 0);
      return { messages: candidates, inline_bodies: ids.map(id => ({ id, signature: 'v1', status: 'ready', text: 'hit' })) };
    } };
    await assert.rejects(searchWindowsFulltext(bridge, { query: 'hit' }, index, { inlineColdRead: true }), { code: 'FULLTEXT_SOURCE_CHANGED' });
  }
});

test('deferred metadata validates non-hits and fills only returned hits without skipping pagination', async () => {
  const rows = ['a', 'b', 'c'].map((id, i) => ({ id, subject: 'other', _fulltext_signature: 'v1', _fulltext_offset: i + 1 }));
  const validation = { candidates: rows.map(item => ({ id: item.id, signature: 'v1' })) };
  let omitMetadata = false;
  const bridge = { call: async (method, params) => {
    if (method === 'fulltext_candidates') return { messages: rows, scanned: 3, next_offset: null, folder_id: 'inbox', validation,
      fulltext_capabilities: { deferred_metadata: true } };
    assert.equal(method, 'fulltext_validate');
    assert.deepEqual(params.validation, validation);
    assert.deepEqual(params.metadata_ids, ['b']);
    return { valid: true, metadata: omitMetadata ? [] : [{ id: 'b', attachment_count: 3 }] };
  } };
  const index = { evaluate: async () => ({ hits: new Map([['b', 'hit'], ['c', 'hit']]), retry_required: false, body_search_incomplete: false }) };
  const result = await searchWindowsFulltext(bridge, { folder: 'inbox', query: 'hit', limit: 1 }, index);
  assert.deepEqual(result.messages.map(item => item.id), ['b']);
  assert.equal(result.messages[0].attachment_count, 3);
  assert.equal(result.next_offset, 2); assert.equal(result.complete, false);
  omitMetadata = true;
  await assert.rejects(searchWindowsFulltext(bridge, { folder: 'inbox', query: 'hit', limit: 1 }, index), { code: 'FULLTEXT_SOURCE_CHANGED' });
});

for (const mode of ['single', 'batch', 'inline']) test(`COM protocol (${mode}) and real SQLite reuse unchanged bodies and update changed or new mail`, async t => {
  const batch = mode !== 'single';
  const directory = await mkdtemp(path.join(tmpdir(), 'windows-com-index-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const rows = [
    { id: 'a', subject: '通知', _fulltext_signature: 'change-a|1', _fulltext_offset: 1 },
    { id: 'b', subject: '通知', _fulltext_signature: 'change-b|1', _fulltext_offset: 2 }
  ];
  const bodies = new Map([['a', '预算'], ['b', '预算审核']]);
  const reads = [];
  const methods = [];
  const makeBridge = () => new OutlookBridge({ fulltextIndex: new WindowsFulltextIndex({ directory: path.join(directory, 'index') }), spawnWorker: () => {
    const worker = new EventEmitter();
    worker.stdout = new PassThrough(); worker.stderr = new PassThrough();
    worker.stdin = new Writable({ write(chunk, encoding, callback) {
      const request = JSON.parse(chunk.toString());
      methods.push(request.method);
      let result;
      if (request.method === 'fulltext_candidates') {
        assert.equal(request.params.search_body, true);
        assert.equal(request.params.after, '2026-10-22T00:00:00Z');
        result = { messages: rows, scanned: rows.length, next_offset: null, folder_id: 'resolved-folder', person_search_incomplete: false,
          fulltext_capabilities: batch ? { batch_read: true, deferred_metadata: true } : undefined,
          validation: { candidates: rows.map(row => ({ id: row.id, signature: row._fulltext_signature, guard: null })) } };
        if (mode === 'inline' && request.params.inline_body_budget) {
          result.inline_bodies = rows.map(row => {
            reads.push(row.id);
            return { id: row.id, signature: row._fulltext_signature, status: request.params.direct_scan ? 'unverified' : 'ready', text: bodies.get(row.id) };
          });
        }
      } else if (request.method === 'fulltext_validate') {
        const changed = request.params.validation.candidates.some(candidate =>
          rows.find(row => row.id === candidate.id)?._fulltext_signature !== candidate.signature);
        if (changed) {
          worker.stdout.write(JSON.stringify({ id: request.id, error: { code: 'FULLTEXT_SOURCE_CHANGED', message: 'Source changed during indexed search' } }) + '\n');
          callback(); return;
        }
        result = { valid: true, metadata: request.params.metadata_ids?.map(id => ({ id, attachment_count: 2 })) };
      } else if (request.method === 'fulltext_read_batch') {
        assert.equal(batch, true);
        assert.equal(request.params.folder, 'resolved-folder');
        assert.ok(request.params.budget_ms > 0 && request.params.budget_ms <= 30000);
        result = { bodies: request.params.candidates.map(candidate => {
          reads.push(candidate.id);
          return { ...candidate, status: 'ready', text: bodies.get(candidate.id) };
        }) };
      } else {
        assert.equal(request.method, 'fulltext_read');
        assert.equal(request.params.folder, 'resolved-folder');
        reads.push(request.params.id);
        const row = rows.find(item => item.id === request.params.id);
        result = { status: 'ready', text: bodies.get(row.id), signature: row._fulltext_signature };
      }
      worker.stdout.write(JSON.stringify({ id: request.id, result }) + '\n');
      callback();
    } });
    worker.kill = () => { worker.stdout.end(); worker.stderr.end(); worker.emit('exit', 0); };
    return worker;
  } });
  const bridge = makeBridge(); t.after(() => bridge.close());
  const args = validateCall('search', { query: '预算', search_body: true, com_body_mode: 'indexed', after: '2026-10-22T00:00:00Z' });
  const initial = await bridge.call('search', args);
  assert.equal(initial.messages.length, 2);
  assert.equal(initial.fulltext_capabilities, undefined);
  assert.equal(initial.inline_bodies, undefined);
  if (mode === 'inline') assert.deepEqual(methods, ['fulltext_candidates', 'fulltext_validate']);
  if (batch) assert.ok(initial.messages.every(item => item.attachment_count === 2));
  assert.deepEqual(reads, ['a', 'b']);
  assert.equal(initial.fulltext.indexed, 2);
  assert.equal((await bridge.call('search', args)).fulltext.reused, 2);
  rows[0]._fulltext_signature = 'change-a-new|1'; bodies.set('a', '决算');
  const updated = await bridge.call('search', args);
  assert.deepEqual(updated.messages.map(item => item.id), ['b']);
  assert.equal(updated.fulltext.updated, 1);
  assert.equal(updated.fulltext.reused, 1);
  rows.push({ id: 'c', subject: '迟到邮件', _fulltext_signature: 'change-c|1', _fulltext_offset: 3 }); bodies.set('c', '预算');
  const restarted = makeBridge(); t.after(() => restarted.close());
  assert.equal((await restarted.call('search', args)).fulltext.reused, 2);
  assert.deepEqual(reads, ['a', 'b', 'a', 'c']);
  rows.splice(1, 1); // A moved/deleted item is absent from the fresh candidate list.
  assert.deepEqual((await bridge.call('search', args)).messages.map(item => item.id), ['c']);
  const evaluate = bridge.fulltextIndex.evaluate.bind(bridge.fulltextIndex);
  bridge.fulltextIndex.evaluate = async (...values) => {
    const result = await evaluate(...values);
    rows[0]._fulltext_signature = 'changed-after-candidate-check';
    return result;
  };
  await assert.rejects(bridge.call('search', args), { code: 'FULLTEXT_SOURCE_CHANGED' });
  assert.ok(methods.every(method => ['fulltext_candidates', batch ? 'fulltext_read_batch' : 'fulltext_read', 'fulltext_validate'].includes(method)));
});

test('index budget exhaustion still validates the source and returns a retry page', async t => {
  let clock = 1_800_000_000_000;
  t.mock.method(Date, 'now', () => clock);
  const methods = [];
  const bridge = new OutlookBridge({
    fulltextIndex: { evaluate: async (candidates, { deadline }) => {
      clock = deadline;
      return { hits: new Map(), retry_required: true, body_search_incomplete: true,
        fulltext: { candidates: 1, pending: 1, indexed: 0, complete: false } };
    } },
    spawnWorker: () => {
      const worker = new EventEmitter();
      worker.stdout = new PassThrough(); worker.stderr = new PassThrough();
      worker.stdin = new Writable({ write(chunk, encoding, callback) {
        const request = JSON.parse(chunk.toString()); methods.push(request.method);
        const result = request.method === 'fulltext_candidates'
          ? { messages: [{ id: 'a', subject: 'other', _fulltext_signature: 'v1', _fulltext_offset: 1 }],
            scanned: 1, next_offset: null, folder_id: 'folder', validation: { candidates: [{ id: 'a', signature: 'v1' }] } }
          : { valid: true };
        if (request.method === 'fulltext_validate') clock += 10;
        worker.stdout.write(JSON.stringify({ id: request.id, result }) + '\n'); callback();
      } });
      worker.kill = () => { worker.stdout.end(); worker.stderr.end(); worker.emit('exit', 0); };
      return worker;
    }
  });
  t.after(() => bridge.close());
  const result = await bridge.call('search', validateCall('search', { query: '预算', search_body: true, com_body_mode: 'indexed' }));
  assert.deepEqual(methods, ['fulltext_candidates', 'fulltext_validate']);
  assert.equal(result.retry_required, true);
  assert.equal(result.next_offset, 0);
  assert.equal(result.complete, false);
  assert.equal(result.fulltext.pending, 1);
});

test('COM person-only search bypasses body indexing even when search_body is enabled', async t => {
  const methods = [];
  const bridge = new OutlookBridge({ fulltextIndex: { evaluate: () => { throw Error('body index must not run'); } },
    spawnWorker: () => {
      const worker = new EventEmitter(); worker.stdout = new PassThrough(); worker.stderr = new PassThrough();
      worker.stdin = new Writable({ write(chunk, encoding, callback) {
        const request = JSON.parse(chunk.toString()); methods.push(request.method);
        assert.equal(request.params.search_body, false);
        worker.stdout.write(JSON.stringify({ id: request.id, result: { messages: [{ id: 'a' }], next_offset: null, complete: true } }) + '\n'); callback();
      } });
      worker.kill = () => { worker.stdout.end(); worker.stderr.end(); worker.emit('exit', 0); }; return worker;
    }
  });
  t.after(() => bridge.close());
  const result = await bridge.call('search', validateCall('search', { person: '张三', search_body: true }));
  assert.deepEqual(methods, ['search']);
  assert.equal(result.complete, true);
  assert.equal(result.next_offset, null);
});
