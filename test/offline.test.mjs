import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { BackendRouter } from '../src/backend.mjs';
import { LocalFileBridge, encodeLocalId, localError, asText } from '../src/offline.mjs';
import { validateCall } from '../src/tools.mjs';

const stub = call => ({ call, queue: [], active: false, close() {} });
test('auto falls back only for unavailable Outlook, explicit COM never falls back', async () => {
  let calls = 0;
  const local = stub(async () => { calls++; return { backend: 'local' }; });
  const router = new BackendRouter({ com: stub(async () => { throw localError('OUTLOOK_UNAVAILABLE', 'closed'); }), local });
  assert.equal((await router.call('health')).fallback_reason, 'OUTLOOK_UNAVAILABLE');
  await assert.rejects(router.call('health', { backend: 'com' }), { code: 'OUTLOOK_UNAVAILABLE' });
  router.com = stub(async () => { throw localError('OUTLOOK_TIMEOUT', 'busy'); });
  await assert.rejects(router.call('health'), { code: 'OUTLOOK_TIMEOUT' });
  assert.equal(calls, 1);
});
test('local IDs retain local routing; COM IDs fail clearly after Outlook closes', async () => {
  const router = new BackendRouter({ com: stub(async () => { throw localError('OUTLOOK_UNAVAILABLE', 'closed'); }), local: stub(async () => ({ backend: 'local' })) });
  const id = encodeLocalId('message', 'test', 42);
  assert.equal((await router.call('read', { id })).backend, 'local');
  await assert.rejects(router.call('read', { id, backend: 'com' }), { code: 'BACKEND_ID_MISMATCH' });
  const oldId = Buffer.from(JSON.stringify({ v: 1, kind: 'message', entry: 'AA', store: 'BB' })).toString('base64url');
  await assert.rejects(router.call('read', { id: oldId }), { code: 'BACKEND_ID_MISMATCH' });
});
test('HTML-only cached body converts without scripts, remote image loading or entity loss', () => {
  const text = asText({ body_format: 'html', body: '<style>.x{color:red}</style><script>alert(1)</script><p>中文 &amp; test</p><img src="https://example.com/tracker"><p>第二行</p>' });
  assert.ok(text.includes('中文 & test') && text.includes('第二行'));
  assert.ok(!text.includes('alert') && !text.includes('tracker') && !text.includes('color:red'));
});

test('local scans are bounded, query text is not SQL, missing fields stay null', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'outlook-offline-test-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const filename = path.join(dir, 'fixture.ost');
  await writeFile(filename, 'fixture');
  const calls = [];
  const bridge = new LocalFileBridge({ files: [filename], run: async args => {
    calls.push(args);
    if (args.includes('--message')) return JSON.stringify({ sender_name: 'A', sender_email: 'a@test.invalid', body: '', body_format: 'plain', body_truncated: false, attachments: [] });
    if (args[2].includes('FROM folders')) return JSON.stringify([{ nid: 1, parent_nid: null, name: 'Root', path: '' }, { nid: 2, parent_nid: 1, name: 'Inbox', path: '/IPM_SUBTREE/Inbox', is_search_folder: false }]);
    return JSON.stringify([1, 2, 3].map(n => ({ nid: n + 10, subject: 'hello', sender_name: 'A', sender_email: null, has_attachments: true, total: 3, delivered: '2026-10-20 01:00:00' })));
  } });
  t.after(() => bridge.close());
  const result = await bridge.call('search', validateCall('search', { query: "'; DELETE FROM messages; --", scan_limit: 2 }));
  assert.equal(result.scanned, 2);
  assert.equal(result.next_offset, 2);
  assert.equal(result.messages.length, 0);
  assert.ok(calls.filter(a => a.includes('--sql')).every(a => !a[2].includes('DELETE')));
  const messages = await bridge.call('messages', validateCall('messages', { limit: 1 }));
  assert.equal(messages.messages[0].attachment_count, null);
  assert.equal(messages.messages[0].sender.smtp_address, null);
});
test('local store mutation is reported instead of silently returning mixed data', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'outlook-offline-change-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const filename = path.join(dir, 'fixture.ost');
  await writeFile(filename, 'before');
  const bridge = new LocalFileBridge({ files: [filename], run: async () => { await writeFile(filename, 'changed-length'); return JSON.stringify({ kind: 'ost-v36', folders: 1 }); } });
  t.after(() => bridge.close());
  await assert.rejects(bridge.call('health', {}), { code: 'OFFLINE_STORE_CHANGED' });
});
