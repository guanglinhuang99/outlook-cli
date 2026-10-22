import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import path from 'node:path';
import { PassThrough, Writable } from 'node:stream';
import { BackendRouter } from '../src/backend.mjs';
import { OutlookBridge } from '../src/bridge.mjs';
import { MacCacheBridge, macWorkerPath } from '../src/mac.mjs';

const stub = call => ({ call, queue: [], active: false, close() {} });

function fakeWorker(handler) {
  const worker = new EventEmitter();
  worker.stdout = new PassThrough();
  worker.stderr = new PassThrough();
  worker.stdin = new Writable({ write(chunk, encoding, callback) { handler(JSON.parse(chunk.toString()), worker); callback(); } });
  worker.kill = () => {
    if (worker.killed) return;
    worker.killed = true;
    worker.stdout.end();
    worker.stderr.end();
    worker.emit('exit', 0);
  };
  return worker;
}

function encoded(value) {
  return Buffer.from(JSON.stringify(value)).toString('base64url');
}

const macId = encoded({ v: 3, backend: 'mac-cache', kind: 'message', key: 'message-key', snapshot: 'a'.repeat(24) });
const localId = encoded({ v: 2, backend: 'local', kind: 'message', store: 'store-key', nid: 42 });
const comId = encoded({ v: 1, kind: 'message', entry: 'AABB', store: 'CCDD' });

test('Mac worker uses the shared JSON-line lifecycle and serializes calls', async t => {
  assert.equal(path.isAbsolute(macWorkerPath), true);
  assert.equal(path.basename(macWorkerPath), 'common_worker.py');
  assert.equal(path.basename(path.dirname(macWorkerPath)), 'macos');
  let inflight = 0;
  let maxInflight = 0;
  let starts = 0;
  const bridge = new MacCacheBridge({
    timeoutMs: 100,
    spawnWorker: () => {
      starts++;
      return fakeWorker((request, worker) => {
        maxInflight = Math.max(maxInflight, ++inflight);
        setTimeout(() => {
          inflight--;
          worker.stdout.write(JSON.stringify({ id: request.id, result: { backend: 'mac-cache', value: request.params.value } }) + '\n');
        }, 5);
      });
    }
  });
  t.after(() => bridge.close());
  const results = await Promise.all([1, 2, 3].map(value => bridge.call('health', { value })));
  assert.deepEqual(results.map(result => result.value), [1, 2, 3]);
  assert.equal(maxInflight, 1);
  assert.equal(starts, 1);
});

test('Mac search with body indexing stays on the Python worker and preserves the request/result', async t => {
  const requests = [];
  const pythonResult = {
    backend: 'mac-cache',
    date_basis: 'received_at',
    messages: [{ id: macId, subject: '项目', body_snippet: '正文命中' }],
    fulltext: { engine: 'sqlite-fts5-trigram', candidates: 1, indexed: 1 },
    body_search_incomplete: false
  };
  const bridge = new MacCacheBridge({
    spawnWorker: () => fakeWorker((request, worker) => {
      requests.push(request);
      worker.stdout.write(JSON.stringify({ id: request.id, result: pythonResult }) + '\n');
    })
  });
  const router = new BackendRouter({ platform: 'darwin', mac: bridge });
  t.after(() => router.close());
  const params = {
    backend: 'mac-cache', folder: 'inbox', query: '正文', person: '张三', person_role: 'to',
    search_body: true, after: null, before: null, unread: null, limit: 20, offset: 0, scan_limit: 50
  };
  const result = await router.call('search', params);
  assert.equal(requests.length, 1);
  assert.equal(requests[0].method, 'search');
  const { backend: _backend, ...workerParams } = params;
  assert.deepEqual(requests[0].params, workerParams);
  assert.equal(bridge.fulltextIndex, undefined);
  assert.deepEqual(result, pythonResult);
});

test('Windows person-only search with the body flag uses the original metadata worker call', async t => {
  const requests = [];
  const fulltextIndex = { evaluate: async () => { throw new Error('person-only search must not invoke the full-text index'); } };
  const workerResult = {
    messages: [{ id: comId, subject: 'metadata result', sender: { name: '张三' }, matched_person_roles: ['to'] }],
    scanned: 1, next_offset: null, complete: true, person_search_incomplete: false
  };
  const bridge = new OutlookBridge({
    fulltextIndex,
    spawnWorker: () => fakeWorker((request, worker) => {
      requests.push(request);
      worker.stdout.write(JSON.stringify({ id: request.id, result: workerResult }) + '\n');
    })
  });
  t.after(() => bridge.close());
  const params = {
    folder: 'inbox', query: '', person: '张三', person_role: 'to', search_body: true,
    after: null, before: null, unread: null, limit: 20, offset: 0, scan_limit: 50
  };
  const result = await bridge.call('search', params);
  assert.equal(requests.length, 1);
  assert.equal(requests[0].method, 'search');
  assert.deepEqual(requests[0].params, { ...params, search_body: false });
  assert.equal(bridge.fulltextIndex, fulltextIndex);
  assert.deepEqual(result, workerResult);
});

test('Mac worker close rejects the active request and missing worker text is generic', async t => {
  const bridge = new MacCacheBridge({ spawnWorker: () => fakeWorker(() => {}) });
  const pending = bridge.call('health', {});
  bridge.close();
  await assert.rejects(pending, { code: 'STOPPING' });

  const missing = new MacCacheBridge({ spawnWorker: () => { throw new Error('python missing'); } });
  t.after(() => missing.close());
  await assert.rejects(missing.call('health', {}), { code: 'WORKER_START', message: 'Could not start reader worker' });
});

test('platform construction and foreign IDs never cross native backends', async () => {
  const defaultMac = new BackendRouter({ platform: 'darwin' });
  assert.ok(defaultMac.mac instanceof MacCacheBridge);
  assert.equal(defaultMac.mac.enableWindowsFulltext, false);
  assert.equal(defaultMac.com, undefined);
  assert.equal(defaultMac.local, undefined);
  defaultMac.close();
  const defaultWin = new BackendRouter({ platform: 'win32' });
  assert.ok(defaultWin.com instanceof OutlookBridge);
  assert.equal(defaultWin.com.enableWindowsFulltext, true);
  assert.equal(defaultWin.mac, undefined);
  defaultWin.close();

  const macCalls = [];
  const mac = stub(async (method, params) => { macCalls.push({ method, params }); return { backend: 'mac-cache', date_basis: 'received_at' }; });
  const macRouter = new BackendRouter({ platform: 'darwin', mac });
  assert.equal(macRouter.com, undefined);
  assert.equal(macRouter.local, undefined);
  const macResult = await macRouter.call('read', { id: macId });
  assert.equal(macResult.backend, 'mac-cache');
  assert.equal(macResult.date_basis, 'received_at');
  await assert.rejects(macRouter.call('read', { id: localId }), { code: 'BACKEND_ID_MISMATCH' });
  await assert.rejects(macRouter.call('read', { id: comId }), { code: 'BACKEND_ID_MISMATCH' });
  await assert.rejects(macRouter.call('health', { backend: 'com' }), { code: 'UNSUPPORTED_BACKEND' });
  assert.equal(macCalls.length, 1);

  const com = stub(async () => ({ connected: true }));
  const local = stub(async () => ({ backend: 'local' }));
  const winRouter = new BackendRouter({ platform: 'win32', com, local });
  assert.equal(winRouter.mac, undefined);
  await assert.rejects(winRouter.call('read', { id: macId }), { code: 'BACKEND_ID_MISMATCH' });
  await assert.rejects(winRouter.call('health', { backend: 'mac-cache' }), { code: 'UNSUPPORTED_BACKEND' });
  const winResult = await winRouter.call('health');
  assert.equal(winResult.backend, 'com');
  assert.equal(winResult.date_basis, 'received_at');
  const localResult = await winRouter.call('health', { backend: 'local' });
  assert.equal(localResult.backend, 'local');
  assert.equal(localResult.date_basis, 'received_at');

  const unsupported = new BackendRouter({ platform: 'linux', com, local, mac });
  await assert.rejects(unsupported.call('health'), { code: 'UNSUPPORTED_PLATFORM' });
});

test('today fixes the Singapore date across UTC+8 midnight and reports the backend date basis', async () => {
  let now = Date.parse('2026-10-19T16:30:00Z');
  const seen = [];
  const com = stub(async (method, params) => { seen.push({ method, params }); return { messages: [], backend: 'com' }; });
  const router = new BackendRouter({ platform: 'win32', com, local: stub(async () => ({ backend: 'local' })), now: () => now });

  const first = await router.call('today', { backend: 'com', folder: 'inbox', limit: 2 });
  assert.equal(first.date, '2026-10-20');
  assert.equal(first.timezone, 'Asia/Singapore');
  assert.equal(first.date_basis, 'received_at');
  assert.deepEqual(seen[0], {
    method: 'messages',
    params: { folder: 'inbox', limit: 2, after: '2026-10-20T00:00:00+08:00', before: '2026-10-21T00:00:00+08:00' }
  });

  now = Date.parse('2026-10-20T16:30:00Z');
  const repeated = await router.call('today', { backend: 'com', date: first.date, limit: 2 });
  assert.equal(repeated.date, first.date);
  assert.equal(seen[1].params.after, '2026-10-20T00:00:00+08:00');
  assert.equal(seen[1].params.before, '2026-10-21T00:00:00+08:00');

  const mac = new BackendRouter({ platform: 'darwin', mac: stub(async () => ({ messages: [], backend: 'mac-cache' })), now: () => now });
  const macToday = await mac.call('today', { backend: 'mac-cache' });
  assert.equal(macToday.date, '2026-10-21');
  assert.equal(macToday.date_basis, 'received_at');
});

test('today rejects malformed calendar dates before touching a backend', async () => {
  let calls = 0;
  const router = new BackendRouter({ platform: 'win32', com: stub(async () => { calls++; return {}; }), local: stub(async () => ({})) });
  await assert.rejects(router.call('today', { date: '2026-02-29' }), { code: 'INVALID_ARGUMENT' });
  await assert.rejects(router.call('today', { date: '2026-1-1' }), { code: 'INVALID_ARGUMENT' });
  assert.equal(calls, 0);
});

test('today rejects unsupported year boundaries', async () => {
  const seen = [];
  const com = stub(async (method, params) => {
    seen.push({ method, params });
    return { messages: [], backend: 'com' };
  });
  const router = new BackendRouter({ platform: 'win32', com, local: stub(async () => ({})) });
  for (const date of ['0000-01-01', '0001-01-01', '1969-12-31', '9999-12-31']) {
    await assert.rejects(router.call('today', { date }), { code: 'INVALID_ARGUMENT' });
  }
  assert.equal(seen.length, 0);
});

test('today uses historical Singapore local-midnight offsets', async () => {
  const seen = [];
  const com = stub(async (method, params) => {
    seen.push({ method, params });
    return { messages: [], backend: 'com' };
  });
  const router = new BackendRouter({ platform: 'win32', com, local: stub(async () => ({})) });

  await router.call('today', { date: '1970-01-01' });
  assert.deepEqual(seen[0], {
    method: 'messages',
    params: {
      after: '1970-01-01T00:00:00+07:30',
      before: '1970-01-02T00:00:00+07:30'
    }
  });

  await router.call('today', { date: '1981-12-31' });
  assert.equal(seen[1].params.after, '1981-12-31T00:00:00+07:30');
  assert.equal(seen[1].params.before, '1982-01-01T00:00:00+08:00');

  await router.call('today', { date: '2026-10-20' });
  assert.equal(seen[2].params.after, '2026-10-20T00:00:00+08:00');
  assert.equal(seen[2].params.before, '2026-10-21T00:00:00+08:00');
});
