import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { BackendRouter } from '../src/backend.mjs';
import { LocalFileBridge, encodeLocalId, decodeId, localError, asText } from '../src/offline.mjs';
import { validateCall } from '../src/tools.mjs';
import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';

const stub = call => ({ call, queue: [], active: false, close() {} });
async function fixtureBridge(t, rows, messages) {
  const dir = await mkdtemp(path.join(tmpdir(), 'outlook-offline-person-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const filename = path.join(dir, 'fixture.ost');
  await writeFile(filename, 'fixture');
  const calls = [];
  const bridge = new LocalFileBridge({ files: [filename], fulltextIndex: new WindowsFulltextIndex({ directory: path.join(dir, 'index') }), run: async args => {
    calls.push(args);
    if (args.includes('--message')) {
      const nid = Number(args[args.indexOf('--message') + 1]);
      return JSON.stringify(messages[nid] ?? { sender_name: null, sender_email: null, recipients: null, body: null, body_format: 'plain', body_truncated: false, attachments: [] });
    }
    if (args[2].includes('FROM folders')) return JSON.stringify([{ nid: 1, parent_nid: null, name: 'Root', path: '' }, { nid: 2, parent_nid: 1, name: 'Inbox', path: '/IPM_SUBTREE/Inbox', is_search_folder: false }]);
    return JSON.stringify(rows.map(row => ({ ...row, total: rows.length })));
  } });
  t.after(() => bridge.close());
  return { bridge, calls };
}
const searchParams = extra => ({ folder: 'inbox', query: '', search_body: false, after: null, before: null, unread: null, limit: 20, offset: 0, scan_limit: 50, ...extra });
test('auto falls back only for unavailable Outlook, explicit COM never falls back', async () => {
  let calls = 0;
  const local = stub(async () => { calls++; return { backend: 'local' }; });
  const router = new BackendRouter({ platform: 'win32', com: stub(async () => { throw localError('OUTLOOK_UNAVAILABLE', 'closed'); }), local });
  const fallback = await router.call('health');
  assert.equal(fallback.fallback_reason, 'OUTLOOK_UNAVAILABLE');
  assert.equal(fallback.date_basis, 'received_at');
  await assert.rejects(router.call('health', { backend: 'com' }), { code: 'OUTLOOK_UNAVAILABLE' });
  router.com = stub(async () => { throw localError('OUTLOOK_TIMEOUT', 'busy'); });
  await assert.rejects(router.call('health'), { code: 'OUTLOOK_TIMEOUT' });
  assert.equal(calls, 1);
});
test('local IDs retain local routing; COM IDs fail clearly after Outlook closes', async () => {
  const router = new BackendRouter({ platform: 'win32', com: stub(async () => { throw localError('OUTLOOK_UNAVAILABLE', 'closed'); }), local: stub(async () => ({ backend: 'local' })) });
  const id = encodeLocalId('message', 'test', 42);
  const localResult = await router.call('read', { id });
  assert.equal(localResult.backend, 'local');
  assert.equal(localResult.date_basis, 'received_at');
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
test('offline person search supports person-only, combined AND, roles, Chinese text, and incomplete recipients', async t => {
  const rows = [
    { nid: 11, subject: '项目通知', sender_name: '发件人', sender_email: 'from@example.invalid', has_attachments: false, delivered: '2026-10-20 04:00:00' },
    { nid: 12, subject: 'Project cc', sender_name: '发件人', sender_email: 'from@example.invalid', has_attachments: false, delivered: '2026-10-20 03:00:00' },
    { nid: 13, subject: 'Sender project', sender_name: 'Alice Sender', sender_email: 'alice@example.invalid', has_attachments: false, delivered: '2026-10-20 02:00:00' },
    { nid: 14, subject: 'Other', sender_name: 'Other', sender_email: 'other@example.invalid', has_attachments: false, delivered: '2026-10-20 01:00:00' }
  ];
  const messages = {
    11: { sender_name: '发件人', sender_email: 'from@example.invalid', recipients: [{ kind: 'to', name: '张伟', email: 'zhang@example.invalid' }], body: '', body_format: 'plain', body_truncated: false, attachments: [] },
    12: { sender_name: '发件人', sender_email: 'from@example.invalid', recipients: [{ kind: 'cc', name: 'Alice Cc', email: 'alice-cc@example.invalid' }], body: '', body_format: 'plain', body_truncated: false, attachments: [] },
    13: { sender_name: 'Alice Sender', sender_email: 'alice@example.invalid', recipients: [], body: '', body_format: 'plain', body_truncated: false, attachments: [] },
    14: { sender_name: 'Other', sender_email: 'other@example.invalid', recipients: null, body: '', body_format: 'plain', body_truncated: false, attachments: [] }
  };
  const { bridge, calls } = await fixtureBridge(t, rows, messages);

  const personOnly = await bridge.call('search', searchParams({ person: 'alice', person_role: 'any' }));
  assert.deepEqual(personOnly.messages.map(message => messages[decodeId(message.id).nid].sender_name), ['发件人', 'Alice Sender']);
  assert.deepEqual(personOnly.messages.map(message => message.matched_person_roles), [['cc'], ['sender']]);
  assert.equal(personOnly.person_search_incomplete, true);
  const beforeSenderOnly = calls.filter(args => args.includes('--message')).length;
  const senderOnly = await bridge.call('search', searchParams({ person: 'alice', person_role: 'sender' }));
  assert.equal(senderOnly.messages.length, 1);
  assert.equal(calls.filter(args => args.includes('--message')).length, beforeSenderOnly);

  const combined = await bridge.call('search', searchParams({ query: '项目', person: '张伟', person_role: 'to' }));
  assert.equal(combined.messages.length, 1);
  assert.deepEqual(combined.messages[0].matched_person_roles, ['to']);
  const roleIsolated = await bridge.call('search', searchParams({ query: '项目', person: '张伟', person_role: 'cc' }));
  assert.equal(roleIsolated.messages.length, 0);
  assert.equal(roleIsolated.person_search_incomplete, false);

  const missing = await bridge.call('search', searchParams({ person: 'other', person_role: 'to' }));
  assert.equal(missing.messages.length, 0);
  assert.equal(missing.person_search_incomplete, true);
  await assert.rejects(bridge.call('search', searchParams({ person: 'alice', person_role: 'bogus' })), { code: 'INVALID_ARGUMENT' });
});
test('offline partial typed contacts remain searchable but report incomplete evidence', async t => {
  const rows = [{ nid: 21, subject: 'partial', sender_name: 'Sender', sender_email: 'sender@example.invalid', has_attachments: false, delivered: '2026-10-20 01:00:00' }];
  const messages = { 21: { sender_name: 'Sender', sender_email: 'sender@example.invalid', recipients: [{ kind: 'to', name: '张三', email: null }], body: '', body_format: 'plain', body_truncated: false, attachments: [] } };
  const { bridge } = await fixtureBridge(t, rows, messages);
  const result = await bridge.call('search', searchParams({ person: '张', person_role: 'to' }));
  assert.equal(result.messages.length, 1);
  assert.deepEqual(result.messages[0].matched_person_roles, ['to']);
  assert.equal(result.person_search_incomplete, true);
});
test('offline person search fetches each candidate at most once and keeps the 50-message search bound', async t => {
  const rows = Array.from({ length: 60 }, (_, index) => ({ nid: index + 10, subject: 'other', sender_name: 'Other', sender_email: 'other@example.invalid', has_attachments: false, delivered: `2026-10-20 01:${String(index % 60).padStart(2, '0')}:00` }));
  const messages = Object.fromEntries(rows.map(row => [row.nid, { sender_name: 'Other', sender_email: 'other@example.invalid', recipients: [], body: '', body_format: 'plain', body_truncated: false, attachments: [] }]));
  const { bridge, calls } = await fixtureBridge(t, rows, messages);
  const result = await bridge.call('search', searchParams({ person: 'nobody', person_role: 'to', scan_limit: 60 }));
  assert.equal(result.scanned, 50);
  assert.equal(result.scan_limit_applied, 50);
  assert.equal(result.next_offset, 50);
  assert.equal(calls.filter(args => args.includes('--message')).length, 50);
});
test('offline combined body and person search filters recipients before fetching full body', async t => {
  const rows = [{ nid: 71, subject: 'metadata miss', sender_name: 'Sender', sender_email: 'sender@example.invalid', has_attachments: false, delivered: '2026-10-20 01:00:00' }];
  const messages = { 71: { sender_name: 'Sender', sender_email: 'sender@example.invalid', recipients: [{ kind: 'to', name: '张三', email: 'zhang@example.invalid' }], body: '正文命中', body_format: 'plain', body_truncated: false, attachments: [] } };
  const { bridge, calls } = await fixtureBridge(t, rows, messages);
  const result = await bridge.call('search', searchParams({ query: '正文命中', search_body: true, person: '张三', person_role: 'to' }));
  assert.equal(result.messages.length, 1);
  assert.deepEqual(calls.filter(args => args.includes('--message')).map(args => args.at(-1)), ['1', '16000000']);
  assert.equal(result.body_search_incomplete, false);
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
