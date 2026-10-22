import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { LocalFileBridge, encodeLocalId, decodeId } from '../src/offline.mjs';
import { WindowsFulltextIndex } from '../src/windows-fulltext-index.mjs';
import { validateCall } from '../src/tools.mjs';
import { searchWindowsFulltext } from '../src/windows-fulltext.mjs';

async function fixture(t) {
  const directory = await mkdtemp(path.join(tmpdir(), 'windows-fulltext-offline-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const filename = path.join(directory, 'mail.ost');
  await writeFile(filename, 'original fixture bytes');
  const folders = [
    { nid: 1, parent_nid: null, name: 'Root', path: '' },
    { nid: 2, parent_nid: 1, name: 'Inbox', path: '/IPM_SUBTREE/Inbox' },
    { nid: 3, parent_nid: 2, name: 'Projects', path: '/IPM_SUBTREE/Inbox/Projects' },
    { nid: 4, parent_nid: 1, name: 'Sent Items', path: '/IPM_SUBTREE/Sent Items' },
    { nid: 5, parent_nid: 1, name: 'Deleted Items', path: '/IPM_SUBTREE/Deleted Items' }
  ];
  const rows = [];
  const messages = new Map();
  const calls = [];
  const run = async args => {
    calls.push(args);
    if (args.includes('--message')) {
      const message = messages.get(Number(args[2]));
      return JSON.stringify({ ...message, attachments: [] });
    }
    const sql = args[2];
    if (sql.includes('FROM folders')) return JSON.stringify(folders);
    let selected = rows.filter(row => row.folder_nid === Number(/folder_nid = (\d+)/.exec(sql)[1]));
    const id = /WHERE nid = (\d+)/.exec(sql);
    if (id) return JSON.stringify(selected.filter(row => row.nid === Number(id[1])));
    const after = /delivered >= TIMESTAMP '([^']+)'/.exec(sql)?.[1];
    const before = /delivered < TIMESTAMP '([^']+)'/.exec(sql)?.[1];
    selected = selected.filter(row => (!after || row.delivered >= after) && (!before || row.delivered < before))
      .sort((a, b) => b.delivered.localeCompare(a.delivered) || b.nid - a.nid);
    const total = selected.length;
    const [, limit, offset] = /LIMIT (\d+) OFFSET (\d+)/.exec(sql);
    return JSON.stringify(selected.slice(Number(offset), Number(offset) + Number(limit)).map(row => ({ ...row, total })));
  };
  const makeBridge = () => new LocalFileBridge({ files: [filename], run,
    fulltextIndex: new WindowsFulltextIndex({ directory: path.join(directory, 'index') }) });
  const bridge = makeBridge();
  t.after(() => bridge.close());
  const add = (nid, body, extra = {}) => {
    const row = { nid, folder_nid: 2, subject: '通知', sender_name: '王五', sender_email: 'wang@example.invalid',
      delivered: '2026-10-22 01:00:00.000', submitted: '2026-01-01 01:00:00.000', modified: '2026-01-01 01:00:00.000', size: 100,
      recipients: [{ kind: 'to', name: '张三', email: 'zhang@example.invalid' }], ...extra };
    rows.push(row);
    messages.set(nid, { ...row, body, body_format: 'plain', body_truncated: false });
  };
  const search = (extra = {}, target = bridge) => target.call('search', validateCall('search', {
    folder: 'inbox', query: '预算', search_body: true, ...extra
  }));
  const bodyReads = () => calls.filter(args => args.includes('--max-body-chars') && args.at(-1) === '16000000').length;
  return { bridge, makeBridge, filename, calls, rows, messages, add, search, bodyReads };
}

test('offline persistent index finds Chinese short words and full-body tail, filters receipt/person before bodies', async t => {
  const f = await fixture(t);
  f.add(10, '前'.repeat(2100000) + '预算审批');
  f.add(11, '预算', { delivered: '2026-10-21 23:59:00.000' });
  f.add(12, '预算', { recipients: [{ kind: 'to', name: '李四', email: 'li@example.invalid' }] });
  const options = { after: '2026-10-22T00:00:00Z', before: '2026-10-23T00:00:00Z', person: '张三', person_role: 'to' };
  const cold = await f.search(options);
  assert.deepEqual(cold.messages.map(item => decodeId(item.id).nid), [10]);
  assert.match(cold.messages[0].body_snippet, /预算审批/);
  assert.equal(cold.date_basis, 'received_at');
  assert.equal(cold.body_search_incomplete, false);
  assert.equal(f.bodyReads(), 1);
  for (const query of ['预', '预算', '预算审批']) assert.equal((await f.search({ ...options, query })).messages.length, 1);
  assert.equal(f.bodyReads(), 1);
  const restarted = f.makeBridge();
  t.after(() => restarted.close());
  assert.equal((await f.search(options, restarted)).fulltext.reused, 1);
  assert.equal(f.bodyReads(), 1);
  assert.equal(await readFile(f.filename, 'utf8'), 'original fixture bytes');
});

test('offline store changes invalidate old bodies even with unchanged message modified time', async t => {
  const f = await fixture(t);
  f.add(10, '旧预算');
  assert.equal((await f.search()).messages.length, 1);
  f.messages.get(10).body = '新决算';
  await writeFile(f.filename, 'new generation of the source fixture');
  assert.equal((await f.search()).messages.length, 0);
  assert.equal((await f.search({ query: '决算' })).messages.length, 1);
  assert.equal(f.bodyReads(), 2);
  // A newly delivered mail can have an old sent/modified timestamp.
  f.add(11, '预算');
  await writeFile(f.filename, 'next generation with delayed delivery');
  assert.deepEqual((await f.search({ after: '2026-10-22T00:00:00Z' })).messages.map(item => decodeId(item.id).nid), [11]);
});

test('missing, truncated and unsupported bodies stay explicit and retry when filled', async t => {
  const f = await fixture(t);
  f.add(10, null);
  f.add(11, '预算'); f.messages.get(11).body_truncated = true;
  f.add(12, '预算'); f.messages.get(12).body_format = 'rtf';
  const partial = await f.search();
  assert.equal(partial.messages.length, 0);
  assert.equal(partial.body_search_incomplete, true);
  assert.equal(partial.retry_required, false);
  f.messages.get(10).body = '预算';
  f.messages.get(11).body_truncated = false;
  f.messages.get(12).body_format = 'plain';
  // Even without a detectable file change, unavailable bodies are never frozen.
  assert.equal((await f.search()).messages.length, 3);
});

test('indexed results cannot leak from other folders or Deleted, and live pagination keeps hits', async t => {
  const f = await fixture(t);
  f.add(10, '预算'); f.add(11, '预算'); f.add(12, '预算');
  const first = await f.search({ limit: 1 });
  const second = await f.search({ limit: 1, offset: first.next_offset });
  const third = await f.search({ limit: 1, offset: second.next_offset });
  assert.deepEqual([first, second, third].flatMap(page => page.messages.map(item => decodeId(item.id).nid)), [12, 11, 10]);
  assert.equal(third.next_offset, null);
  const store = decodeId(first.messages[0].id).store;
  f.rows[0].folder_nid = 3;
  await writeFile(f.filename, 'moved into inbox subfolder');
  assert.equal((await f.search()).messages.length, 2);
  assert.equal((await f.search({ folder: encodeLocalId('folder', store, 3) })).messages.length, 1);
  f.add(13, '预算', { folder_nid: 4 });
  assert.equal((await f.search({ folder: 'sent' })).messages.length, 1);
  await assert.rejects(f.search({ folder: 'deleted' }), { code: 'FULLTEXT_FOLDER_SCOPE' });
  f.rows[0].folder_nid = 5;
  assert.equal((await f.search({ folder: encodeLocalId('folder', store, 3) })).messages.length, 0);
});

test('offline validates source after a warm index hit, not just on body reads', async t => {
  const f = await fixture(t);
  f.add(10, '预算');
  await f.search();
  const evaluate = f.bridge.fulltextIndex.evaluate.bind(f.bridge.fulltextIndex);
  f.bridge.fulltextIndex.evaluate = async (...args) => {
    const result = await evaluate(...args);
    await writeFile(f.filename, 'changed after cached search');
    return result;
  };
  await assert.rejects(f.search(), { code: 'OFFLINE_STORE_CHANGED' });
});

test('shared orchestration keeps current page when body work is pending and combines metadata OR body', async () => {
  const page = { messages: [
    { id: 'a', subject: '预算', _fulltext_signature: '1', _fulltext_offset: 1 },
    { id: 'b', subject: '通知', _fulltext_signature: '2', _fulltext_offset: 3 }
  ], next_offset: 4, scanned: 4 };
  const calls = [];
  const bridge = { call: async (method, params) => { calls.push({ method, params }); return page; } };
  const index = { evaluate: async () => ({ hits: new Map([['b', '预算正文']]), fulltext: { pending: 1 }, retry_required: true, body_search_incomplete: true }) };
  const result = await searchWindowsFulltext(bridge, { query: '预算', offset: 0, limit: 10, scan_limit: 500 }, index);
  assert.deepEqual(result.messages.map(item => item.id), ['a', 'b']);
  assert.equal(result.next_offset, 0);
  assert.equal(result.complete, false);
  assert.equal(calls[0].params.scan_limit, 50);
  assert.ok(result.messages.every(item => !Object.hasOwn(item, '_fulltext_signature')));
});
