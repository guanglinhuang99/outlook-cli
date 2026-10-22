import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { EventEmitter } from 'node:events';
import { PassThrough, Writable } from 'node:stream';
import { OutlookBridge } from '../src/bridge.mjs';
import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';
import { validateCall } from '../src/tools.mjs';

test('COM protocol and real SQLite reuse unchanged bodies and independently update changed or new mail', async t => {
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
          validation: { candidates: rows.map(row => ({ id: row.id, signature: row._fulltext_signature, guard: null })) } };
      } else if (request.method === 'fulltext_validate') {
        const changed = request.params.validation.candidates.some(candidate =>
          rows.find(row => row.id === candidate.id)?._fulltext_signature !== candidate.signature);
        if (changed) {
          worker.stdout.write(JSON.stringify({ id: request.id, error: { code: 'FULLTEXT_SOURCE_CHANGED', message: 'Source changed during indexed search' } }) + '\n');
          callback(); return;
        }
        result = { valid: true };
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
  const args = validateCall('search', { query: '预算', search_body: true, after: '2026-10-22T00:00:00Z' });
  assert.equal((await bridge.call('search', args)).messages.length, 2);
  assert.deepEqual(reads, ['a', 'b']);
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
  assert.ok(methods.every(method => ['fulltext_candidates', 'fulltext_read', 'fulltext_validate'].includes(method)));
});
