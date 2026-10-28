import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { EventEmitter } from 'node:events';
import { PassThrough, Writable } from 'node:stream';
import { OutlookBridge } from '../src/bridge.mjs';
import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';
import { searchWindowsFulltext } from '../src/windows-fulltext.mjs';

async function fixture(t) {
  const root = await mkdtemp(path.join(tmpdir(), 'com-deferred-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  return path.join(root, 'index');
}

test('COM default scans once, promotes complete non-hit bodies, reuses and invalidates versions', async t => {
  const directory = await fixture(t);
  let signature = 'v1', text = 'later keyword', bodyReads = 0, validations = 0;
  const makeBridge = () => new OutlookBridge({ fulltextIndex: new WindowsFulltextIndex({ directory }), spawnWorker: () => {
    const worker = new EventEmitter(); worker.stdout = new PassThrough(); worker.stderr = new PassThrough();
    worker.stdin = new Writable({ write(chunk, encoding, callback) {
      const request = JSON.parse(chunk.toString());
      const body = { id: 'a', signature, status: 'ready', text };
      let result;
      if (request.method === 'fulltext_candidates') {
        assert.equal(request.params.direct_match_query, undefined);
        if (request.params.direct_scan) bodyReads++;
        result = { messages: [{ id: 'a', subject: 'other', _fulltext_signature: signature }], next_offset: null,
          validation: { candidates: [{ id: 'a', signature }] }, fulltext_capabilities: { batch_read: true, deferred_metadata: true },
          ...(request.params.direct_scan ? { inline_bodies: [{ ...body, status: 'unverified' }] } : {}) };
      } else if (request.method === 'fulltext_read_batch') { bodyReads++; result = { bodies: [body] }; }
      else { assert.equal(request.method, 'fulltext_validate'); validations++; result = { valid: true, metadata: request.params.metadata_ids.map(id => ({ id, attachment_count: 0 })) }; }
      worker.stdout.write(JSON.stringify({ id: request.id, result }) + '\n'); callback();
    } });
    worker.kill = () => { worker.stdout.end(); worker.stderr.end(); worker.emit('exit', 0); }; return worker;
  } });
  const bridge = makeBridge(); t.after(() => bridge.close());
  const search = (query, extra = {}) => bridge.call('search', { query, search_body: true, ...extra });
  const first = await search('absent');
  assert.equal(first.messages.length, 0); assert.equal(first.fulltext.index_state, 'deferred');
  await assert.rejects(stat(directory), { code: 'ENOENT' });
  const second = await search('keyword', { com_body_mode: 'auto' });
  assert.equal(second.messages.length, 1); assert.equal(second.fulltext.staged_reused, 1);
  assert.equal(second.fulltext.updated, 1); assert.equal(bodyReads, 1);
  assert.equal((await search('later')).fulltext.reused, 1); assert.equal(bodyReads, 1);
  signature = 'v2'; text = 'replacement';
  assert.equal((await search('keyword')).messages.length, 0); assert.equal(bodyReads, 2);
  bridge.close();
  const restarted = makeBridge(); t.after(() => restarted.close());
  const afterRestart = await restarted.call('search', { query: 'replacement', search_body: true });
  assert.equal(afterRestart.fulltext.reused, 1); assert.equal(afterRestart.messages.length, 1);
  assert.equal(bodyReads, 2); assert.equal(validations, 5);
});

test('deferred bodies expire, reject weak/changed versions, exclude compact omissions, and clear on close', async t => {
  let now = 100000;
  const index = new WindowsFulltextIndex({ directory: await fixture(t), now: () => now });
  t.after(() => index.close());
  const candidates = [{ id: 'a', signature: 'v1' }, { id: 'weak', signature: null }, { id: 'omitted', signature: 'v1' }];
  const bodies = new Map(candidates.map(c => [c.id, c.id === 'omitted'
    ? { ...c, status: 'unverified_no_match', query: 'absent', body_bytes: 4, body_code_units: 4 }
    : { ...c, status: 'unverified', text: 'later keyword' }]));
  const result = index.evaluateDirect(candidates, bodies, { query: 'absent', retainBodies: true });
  assert.equal(result.accepted.has('omitted'), false);
  const stage = () => index.stageValidatedBodies(result.accepted);
  stage(); assert.deepEqual([...index.takeDeferredBodies(candidates).keys()], ['a']);
  assert.equal(index.takeDeferredBodies(candidates).size, 0);
  stage(); assert.equal(index.takeDeferredBodies([{ id: 'a', signature: 'v2' }]).size, 0);
  stage(); now += 30000; assert.equal(index.prepareDeferredCache(), null); assert.equal(index.takeDeferredBodies(candidates).size, 0);
  stage(); index.close(); assert.equal(index.takeDeferredBodies(candidates).size, 0);
});

test('failed hybrid proofs never stage bodies and changed versions are reread before promotion', async t => {
  const directory = await fixture(t);
  for (const fail of [true, false]) {
    const index = new WindowsFulltextIndex({ directory: path.join(directory, String(fail)) });
    t.after(() => index.close());
    let signature = 'v1', reads = 0;
    const bridge = { call: async (method, params) => {
      if (method === 'fulltext_candidates') return { messages: [{ id: 'a', _fulltext_signature: signature }], next_offset: null,
        validation: { candidates: [{ id: 'a', signature }] }, fulltext_capabilities: { batch_read: true },
        ...(params.direct_scan ? { inline_bodies: [{ id: 'a', signature, status: 'unverified', text: 'old' }] } : {}) };
      if (method === 'fulltext_read_batch') { reads++; return { bodies: [{ id: 'a', signature, status: 'ready', text: 'new' }] }; }
      assert.equal(method, 'fulltext_validate'); return { valid: !fail };
    } };
    const search = query => searchWindowsFulltext(bridge, { query }, index, { inlineColdRead: true, deferredIndex: true });
    if (fail) {
      await assert.rejects(search('old'), { code: 'FULLTEXT_SOURCE_CHANGED' });
      assert.equal(index.takeDeferredBodies([{ id: 'a', signature }]).size, 0);
      await assert.rejects(stat(index.directory), { code: 'ENOENT' });
    } else {
      await search('old'); signature = 'v2';
      const result = await search('new');
      assert.equal(result.messages.length, 1); assert.equal(reads, 1); assert.equal(result.fulltext.updated, 1);
    }
  }
});
