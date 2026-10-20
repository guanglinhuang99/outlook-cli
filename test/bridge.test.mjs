import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { PassThrough, Writable } from 'node:stream';
import { OutlookBridge } from '../src/bridge.mjs';

function fakeWorker(handler) {
  const worker = new EventEmitter();
  worker.stdout = new PassThrough();
  worker.stderr = new PassThrough();
  worker.stdin = new Writable({ write(chunk, encoding, callback) { handler(JSON.parse(chunk.toString()), worker); callback(); } });
  worker.kill = () => { worker.stdout.end(); worker.stderr.end(); worker.emit('exit', 0); };
  return worker;
}
test('COM calls serialize and preserve request-response correlation', async t => {
  let inflight = 0;
  let max = 0;
  const bridge = new OutlookBridge({ spawnWorker: () => fakeWorker((request, worker) => {
    max = Math.max(max, ++inflight);
    setTimeout(() => { inflight--; worker.stdout.write(JSON.stringify({ id: request.id, result: { value: request.params.value } }) + '\n'); }, 10);
  }) });
  t.after(() => bridge.close());
  const results = await Promise.all([1, 2, 3].map(value => bridge.call('health', { value })));
  assert.deepEqual(results.map(r => r.value), [1, 2, 3]);
  assert.equal(max, 1);
});
test('timeout kills worker, rejects queued calls, next call starts a fresh worker', async t => {
  let starts = 0;
  const bridge = new OutlookBridge({ timeoutMs: 25, spawnWorker: () => {
    starts++;
    return fakeWorker((request, worker) => {
      if (starts > 1) worker.stdout.write(JSON.stringify({ id: request.id, result: { recovered: true } }) + '\n');
    });
  } });
  t.after(() => bridge.close());
  const results = await Promise.allSettled([bridge.call('health', {}), bridge.call('health', {})]);
  assert.ok(results.every(r => r.status === 'rejected' && r.reason.code === 'OUTLOOK_TIMEOUT'));
  assert.equal((await bridge.call('health', {})).recovered, true);
  assert.equal(starts, 2);
});
