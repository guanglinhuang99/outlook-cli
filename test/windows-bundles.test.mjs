import test from 'node:test';
import assert from 'node:assert/strict';
import { BackendRouter } from '../src/backend.mjs';
import { validateCall } from '../src/tools.mjs';
import { runWindowsBundle } from '../src/windows-bundles.mjs';

const folder = (name, child_count = 0) => ({ id: `folder_${name}`, name, path: `/${name}`, child_count });
const message = id => ({ id, subject: id });
const page = (ids = [], next_offset = null, extra = {}) => ({
  messages: ids.map(message), next_offset, complete: next_offset === null, ...extra
});
const body = (id, text) => ({
  ...message(id), to: 'recipient', cc: '', attachments: [],
  body: { text, total_chars: text.length, offset: 0, next_offset: null, truncated: false }
});
const decode = cursor => JSON.parse(Buffer.from(cursor, 'base64url'));
const run = (method, params, call) => runWindowsBundle(method, validateCall(method, params), call);
const root = folder('inbox');
const sent = folder('sent');

test('expands paged folder trees breadth-first and groups messages by folder', async () => {
  const a = folder('a', 1);
  const b = folder('b');
  const c = folder('c');
  const tree = new Map([
    ['inbox:0', { parent: root, folders: [a], next_offset: 401 }],
    [`${root.id}:401`, { parent: root, folders: [b], next_offset: null }],
    [`${a.id}:0`, { parent: a, folders: [c], next_offset: null }],
    ['sent:0', { parent: sent, folders: [], next_offset: null }]
  ]);
  const visited = [];
  const result = await run('context_bundle', {
    folder: 'all', from_date: '2024-02-28', to_date: '2024-02-29', unread: false
  }, async (method, params) => {
    if (method === 'folders') {
      assert.ok(params.limit <= 200);
      visited.push(`${params.folder}:${params.offset}`);
      return tree.get(visited.at(-1));
    }
    assert.equal(method, 'messages');
    assert.equal(params.after, '2024-02-28T00:00:00+08:00');
    assert.equal(params.before, '2024-03-01T00:00:00+08:00');
    assert.equal(params.unread, false);
    return page([`${params.folder}_new`, `${params.folder}_old`]);
  });
  assert.deepEqual(visited, ['inbox:0', `${root.id}:401`, `${a.id}:0`, 'sent:0']);
  assert.deepEqual(result.items, [root, a, b, c, sent].flatMap(({ id, name, path }) => ['new', 'old'].map(age => ({
    ...message(`${id}_${age}`), folder: { id, name, path }
  }))));
  assert.equal(result.folders_scanned, 5);
  assert.equal(result.folders_truncated, false);
  assert.equal(result.body_complete, null);
  assert.equal(result.next_cursor, null);
  assert.equal(result.complete, true);
});

test('uses returned live offsets and resumes at the next folder after exhaustion', async () => {
  const child = folder('child');
  let children = [child];
  const visited = [];
  const call = async (method, params) => {
    if (method === 'folders') return { parent: root, folders: children, next_offset: null };
    visited.push([params.folder, params.offset, params.limit]);
    if (params.folder === root.id && params.offset === 0) return page(['first'], 39);
    return page([params.folder === root.id ? 'second' : 'third']);
  };
  const options = { limit: 1 };
  const first = await run('context_bundle', options, call);
  assert.deepEqual({ ...decode(first.next_cursor), q: 'hash' }, {
    v: 1, m: 'context_bundle', d: null, f: null, t: null, q: 'hash', fid: root.id, o: 39
  });
  const second = await run('context_bundle', { ...options, cursor: first.next_cursor }, call);
  assert.equal(decode(second.next_cursor).fid, child.id);
  assert.equal(decode(second.next_cursor).o, 0);
  const third = await run('context_bundle', { ...options, cursor: second.next_cursor }, call);
  assert.equal(third.next_cursor, null);
  assert.deepEqual(visited, [[root.id, 0, 1], [root.id, 39, 1], [child.id, 0, 1]]);
  const mismatch = {
    code: 'INVALID_ARGUMENT',
    message: 'cursor belongs to another folder or query, or the folder list changed; start again without cursor'
  };
  await assert.rejects(run('context_bundle', { ...options, cursor: second.next_cursor, query: 'changed' }, call), mismatch);
  children = [];
  await assert.rejects(run('context_bundle', { ...options, cursor: second.next_cursor }, call), mismatch);
});

test('returns an empty continuation page after eight calls and caps expansion at 300 folders', async () => {
  const children = Array.from({ length: 310 }, (_, index) => folder(String(index)));
  const scanned = [];
  const call = async (method, params) => {
    if (method === 'folders') {
      const end = params.offset + params.limit;
      return { parent: root, folders: children.slice(params.offset, end), next_offset: end < children.length ? end : null };
    }
    scanned.push(params.folder);
    return page();
  };
  const first = await run('context_bundle', {}, call);
  assert.deepEqual(first.items, []);
  assert.equal(first.folders_scanned, 8);
  assert.equal(first.folders_truncated, true);
  assert.equal(first.complete, false);
  assert.equal(decode(first.next_cursor).fid, children[7].id);
  let cursor = first.next_cursor;
  while (cursor) {
    const next = await run('context_bundle', { cursor }, call);
    cursor = next.next_cursor;
  }
  assert.deepEqual(scanned, [root, ...children.slice(0, 299)].map(entry => entry.id));
});

test('today cursor pins the Singapore date across midnight', async () => {
  let now = Date.parse('2026-11-28T15:59:00Z');
  const bridge = { async call(method, params) {
    if (method === 'folders') return { parent: root, folders: [folder('ignored')], next_offset: 20 };
    assert.equal(method, 'messages');
    assert.equal(params.after, '2026-11-28T00:00:00+08:00');
    assert.equal(params.before, '2026-11-29T00:00:00+08:00');
    return params.offset === 0 ? page([], 3) : page(['found']);
  } };
  const router = new BackendRouter({ platform: 'win32', com: bridge, local: bridge, mode: 'com', now: () => now });
  const options = { mode: 'metadata', include_subfolders: false };
  const first = await router.call('today_bundle', validateCall('today_bundle', options));
  assert.equal(first.date, '2026-11-28');
  now += 120000;
  const second = await router.call('today_bundle', validateCall('today_bundle', { ...options, cursor: first.next_cursor }));
  assert.equal(second.date, first.date);
  assert.equal(second.next_cursor, null);
  await assert.rejects(router.call('today_bundle', validateCall('today_bundle', {
    ...options, cursor: first.next_cursor, date: '2026-11-29'
  })), { code: 'INVALID_ARGUMENT', message: 'date does not match the cursor date' });
});

test('context mode retains metadata after read failure and budget exhaustion', async () => {
  const reads = [];
  const result = await run('context_bundle', { mode: 'context', max_chars: 600, text_budget_chars: 1000 }, async (method, params) => {
    if (method === 'folders') return { parent: root, folders: [], next_offset: null };
    if (method === 'messages') return page(['broken', 'first', 'second', 'skipped']);
    reads.push([params.id, params.max_chars]);
    if (params.id === 'broken') throw Object.assign(new Error('read failed'), { code: 'FAKE_READ' });
    return body(params.id, '𐐀'.repeat(params.max_chars / 2));
  });
  assert.deepEqual(reads, [['broken', 600], ['first', 600], ['second', 400]]);
  assert.deepEqual(result.errors, [
    { id: 'broken', code: 'FAKE_READ', message: 'read failed' },
    { id: 'skipped', code: 'TEXT_BUDGET_EXHAUSTED', message: 'Bundle text budget exhausted' }
  ]);
  assert.deepEqual(result.items.map(item => [item.id, item.body?.text.length]), [
    ['broken', undefined], ['first', 600], ['second', 400], ['skipped', undefined]
  ]);
  assert.deepEqual(result.items[1].attachments, []);
  assert.equal(result.items[1].to, 'recipient');
  assert.equal(result.items[1].cc, '');
  assert.equal(result.body_text_chars, 1000);
  assert.equal(result.body_complete, false);
  assert.equal(result.complete, true);
});

test('search retry keeps the same folder and offset and reports incomplete flags', async () => {
  const inputs = [];
  const call = async (method, params) => {
    if (method === 'folders') return { parent: root, folders: [], next_offset: null };
    assert.equal(method, 'search');
    inputs.push(params);
    if (inputs.length === 1) return page([], 17);
    if (inputs.length === 2) return page(['partial'], 99, {
      retry_required: true, body_search_incomplete: true, person_search_incomplete: true
    });
    return page(['found']);
  };
  const options = { person: 'Kim', person_role: 'to', search_body: true };
  const first = await run('context_bundle', options, call);
  const retry = await run('context_bundle', { ...options, cursor: first.next_cursor }, call);
  assert.equal(retry.next_cursor, first.next_cursor);
  assert.equal(retry.retry_required, true);
  assert.deepEqual(retry.items, [], 'partial hits of a page that must be repeated are not returned twice');
  assert.equal(retry.body_search_incomplete, true);
  assert.equal(retry.person_search_incomplete, true);
  assert.equal(retry.complete, false);
  const done = await run('context_bundle', { ...options, cursor: retry.next_cursor }, call);
  assert.equal(done.complete, true);
  assert.deepEqual(inputs.map(input => input.offset), [0, 17, 17]);
  assert.equal(inputs[2].person, 'Kim');
  assert.equal(inputs[2].person_role, 'to');
});

test('rejects unsupported filters before calling the backend', async () => {
  const call = () => assert.fail('unexpected backend call');
  for (const params of [{ field: 'subject' }, { has_cached_attachments: false }]) {
    await assert.rejects(run('context_bundle', params, call), {
      code: 'INVALID_ARGUMENT', message: new RegExp(Object.keys(params)[0])
    });
  }
});

test('messages bundle deduplicates IDs and shares the text budget', async () => {
  const args = [];
  const ids = ['message_1', 'message_1', 'message_2', 'message_3', 'message_4'];
  const result = await run('messages_bundle', { ids, max_chars: 700, text_budget_chars: 1000 }, async (method, params) => {
    assert.equal(method, 'read');
    args.push([params.id, params.body_offset, params.max_chars]);
    if (params.id === 'message_1') throw Object.assign(new Error('item removed'), { code: 'INVALID_ID' });
    return body(params.id, 'x'.repeat(params.max_chars));
  });
  assert.deepEqual(args, [['message_1', 0, 700], ['message_2', 0, 700], ['message_3', 0, 300]]);
  assert.deepEqual(result.items.map(item => item.id), ['message_2', 'message_3']);
  assert.deepEqual(result.errors.map(({ id, code }) => [id, code]), [
    ['message_1', 'INVALID_ID'], ['message_4', 'TEXT_BUDGET_EXHAUSTED']
  ]);
  assert.equal(result.body_text_chars, 1000);
  assert.equal(result.body_complete, false);
});

test('router exposes all three bundles on Windows and rejects them on Linux', async () => {
  const bridge = { async call(method, params) {
    if (method === 'folders') return {
      parent: root, folders: [], next_offset: null, backend: 'local', data_scope: 'cached_only',
      warnings: [{ code: 'CACHE', message: 'First warning' }]
    };
    if (method === 'messages') return page(['message_1'], null, {
      backend: 'local', data_scope: 'cached_only', warnings: [{ code: 'CACHE', message: 'Duplicate warning' }]
    });
    assert.equal(method, 'read');
    return { ...body(params.id, 'hello'), backend: 'com', data_scope: 'outlook_session' };
  } };
  const com = { call() { throw Object.assign(new Error('Outlook unavailable'), { code: 'OUTLOOK_UNAVAILABLE' }); } };
  const windows = new BackendRouter({ platform: 'win32', com, local: bridge, mode: 'auto' });
  const linux = new BackendRouter({ platform: 'linux', com: bridge, local: bridge, mode: 'auto' });
  for (const [method, input] of [
    ['context_bundle', { mode: 'context' }], ['today_bundle', { date: '2026-11-28' }],
    ['messages_bundle', { ids: ['message_1'] }]
  ]) {
    const params = validateCall(method, { ...input, backend: 'local' });
    const result = await windows.call(method, params);
    assert.equal(result.items[0].id, 'message_1');
    assert.equal(result.body_complete, true);
    assert.equal(result.backend, 'com');
    assert.equal(result.data_scope, 'outlook_session');
    if (method !== 'messages_bundle') assert.deepEqual(result.warnings, [
      { code: 'CACHE', message: 'First warning' },
      { code: 'MIXED_BACKENDS', message: 'Bundle operations returned different backends or data scopes.' }
    ]);
    await assert.rejects(linux.call(method, params), { code: 'UNSUPPORTED_PLATFORM' });
  }
  const fallback = await windows.call('context_bundle', validateCall('context_bundle'));
  assert.equal(fallback.backend, 'local');
  assert.ok(fallback.warnings.some(({ code }) => code === 'LOCAL_FALLBACK'));
});
