import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, readdir, rm, stat, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { runReadParity } from '../src/read-parity.mjs';

const SNAP1 = 'a'.repeat(24);
const SNAP2 = 'b'.repeat(24);
const macCacheState = { skip: process.platform === 'win32' };

function id(kind, key, snapshot) {
  return Buffer.from(JSON.stringify({ v: 3, backend: 'mac-cache', kind, key, snapshot })).toString('base64url');
}

function item(key, snapshot, fields = {}) {
  return {
    id: id('message', key, snapshot),
    internet_message_id: `<${key}@example.test>`,
    related_internet_message_ids: [],
    subject: 'Project update',
    folder_id: id('folder', 'inbox-root', snapshot),
    folder_path: 'Inbox',
    unread: null,
    received_at: '2026-11-24T20:00:00Z',
    ...fields
  };
}

function fixture(rows, options = {}) {
  const calls = [];
  const snapshot = options.snapshot ?? SNAP1;
  const context = { backend: 'mac-cache', snapshot, date_basis: 'received_at',
    scope_from_date: options.horizon ?? '2026-01-01', folder_scope_complete: false };
  const call = async (method, params) => {
    calls.push({ method, params });
    if (method === 'stores') return { ...context, snapshot: options.head ?? snapshot,
      stores: [{ root_id: id('folder', 'test-store', options.head ?? snapshot) }] };
    if (method === 'folders') return { ...context, snapshot: options.head ?? snapshot,
      parent: { id: id('folder', options.aliasRoot ?? 'inbox-root', options.head ?? snapshot) } };
    if (method === 'messages_bundle') {
      return {
        ...context,
        items: params.ids.map(id => ({ id, body: { available: true, text: 'Selected text' } })),
        errors: [],
        body_complete: true
      };
    }
    if (method === 'read' && options.anchor) return options.anchor;
    assert.equal(method, 'context_bundle');
    assert.equal(params.mode, 'metadata');
    const offset = Number(params.cursor ?? 0);
    const snapshotAtOffset = offset && options.laterSnapshot ? options.laterSnapshot : snapshot;
    const items = rows.slice(offset, offset + params.limit);
    return {
      ...context, snapshot: snapshotAtOffset, items, total_matched: rows.length,
      received_time_unknown: options.datedIncomplete ? 1 : 0, date_filter_incomplete: options.datedIncomplete ?? false,
      include_subfolders: params.include_subfolders,
      complete: offset + items.length === rows.length,
      next_cursor: offset + items.length < rows.length ? String(offset + items.length) : null
    };
  };
  return { calls, call };
}

async function state(t) {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'outlook-read-workflow-test-'));
  t.after(() => rm(directory, { force: true, recursive: true }));
  return path.join(directory, 'checkpoint-state');
}

test('conversation orders explicit reference and ID matches before labeled subject fallback', async () => {
  const anchor = item('anchor', SNAP1);
  const rows = [
    item('guess', SNAP1, { subject: '回复： RE: Project   update' }),
    item('other-chain', SNAP1, { related_internet_message_ids: ['<other@example.test>'] }),
    item('left-aside', SNAP1, { related_internet_message_ids: ['<ANCHOR@example.test>'] }),
    item('follow-up', SNAP1, { subject: 'Changed subject', related_internet_message_ids: [anchor.internet_message_id] }),
    item('follow-up-copy', SNAP1, { internet_message_id: '<follow-up@example.test>', subject: 'Different title' }),
    anchor
  ];
  const { call, calls } = fixture(rows);
  const result = await runReadParity('conversation', { id: anchor.id, limit: 3, include_body: true }, { call });
  assert.deepEqual(result.items.map(({ id, match_basis, body }) => [id, match_basis, body?.text]), [
    [anchor.id, 'anchor', 'Selected text'],
    [rows[3].id, 'exact_message_id_reference', 'Selected text'],
    [rows[4].id, 'exact_internet_message_id', 'Selected text']
  ]);
  assert.equal(result.matched_in_scan, 4);
  assert.equal(result.mailbox_complete, false);
  assert.equal(result.body.complete, true);
  assert.equal(result.truncated, true);
  assert.deepEqual(calls[1].params.ids, result.items.map(row => row.id));

  const fallback = await runReadParity('conversation', { id: anchor.id }, { call });
  assert.equal(fallback.items.at(-1).id, rows[0].id);
  assert.equal(fallback.items.at(-1).match_basis, 'normalized_subject_fallback');
  const strict = await runReadParity('conversation', { id: anchor.id, subject_fallback: false }, { call });
  assert.equal(strict.items.length, 3);
});

test('conversation can use explicit shared references to an unavailable parent, without treating a different-case ID as equal', async () => {
  const missingParent = '<missing@example.test>';
  const anchor = item('anchor', SNAP1, { subject: 'Anchor', related_internet_message_ids: [missingParent] });
  const sibling = item('sibling', SNAP1, { subject: 'Sibling', related_internet_message_ids: [missingParent] });
  const other = item('other', SNAP1, { subject: 'Other chain', related_internet_message_ids: ['<Missing@example.test>'] });
  const { call } = fixture([other, sibling], { anchor: { ...anchor, body: { text: 'private' }, attachments: [] } });
  const output = await runReadParity('conversation', { id: anchor.id }, { call });
  assert.equal(output.anchor_in_scan, false);
  assert.equal(output.matched_in_scan, 1);
  assert.deepEqual(output.items.map(({ id, match_basis }) => [id, match_basis]),
    [[anchor.id, 'anchor'], [sibling.id, 'exact_shared_message_id_reference']]);
  assert.equal(output.items[0].body, undefined);
});

test('delta keeps stable logical keys across opaque IDs, snapshots and same-key folder IDs', macCacheState, async t => {
  const stateDir = await state(t);
  const rows = [item('same', SNAP1), item('changed', SNAP1), item('gone', SNAP1)];
  const before = fixture(rows);
  const folder1 = rows[0].folder_id;
  const folder2 = id('folder', 'inbox-root', SNAP2);
  const baseline = await runReadParity('delta-bundle', { folder: folder1 }, { call: before.call, stateDir });
  assert.match(baseline.checkpoint, /^[a-f0-9]{64}$/);
  assert.equal(baseline.baseline, true);
  assert.equal(baseline.added.length, 3);
  const file = path.join(stateDir, `checkpoint-${baseline.checkpoint}.json`);
  const content = await readFile(file, 'utf8');
  assert.doesNotMatch(content, /example\.test|Project update|inbox-root/);
  assert.equal((await stat(file)).mode & 0o777, 0o600);
  assert.equal((await stat(stateDir)).mode & 0o777, 0o700);

  const nextRows = [item('same', SNAP2), item('changed', SNAP2, { subject: 'Changed subject' }), item('new', SNAP2)];
  const next = fixture(nextRows, { snapshot: SNAP2 });
  const result = await runReadParity('delta-bundle', { folder: folder2, since: baseline.checkpoint },
    { call: next.call, stateDir });
  assert.deepEqual(result.added.map(row => row.id), [nextRows[2].id]);
  assert.deepEqual(result.changed.map(row => row.id), [nextRows[1].id]);
  assert.deepEqual(result.removed_from_local_scope, [{ stable_id: baseline.added[2].stable_id }]);
  assert.equal(result.changed[0].stable_id, baseline.added[1].stable_id);
  assert.equal(result.removed_from_local_scope_complete, true);
  assert.equal(result.server_deletions_verified, false);
});

test('delta never saves checkpoints or implies removal after a cap, mixed snapshot, dated omission, or outdated head',
  macCacheState, async t => {
  const stateDir = await state(t);
  const baseline = await runReadParity('delta-bundle', { from_date: '2026-01-01' },
    { call: fixture([item('gone', SNAP1)]).call, stateDir });
  const current = Array.from({ length: 101 }, (_, i) => item(String(i), SNAP2));
  for (const [fixtures, params, reason] of [
    [{ snapshot: SNAP2 }, { scan_limit: 1 }, 'scan_limit_reached'],
    [{ snapshot: SNAP2, laterSnapshot: SNAP1 }, {}, 'snapshot_changed_during_scan'],
    [{ snapshot: SNAP2, datedIncomplete: true }, {}, 'unknown_received_time_excluded'],
    [{ snapshot: SNAP2, head: SNAP1 }, {}, 'snapshot_no_longer_current']
  ]) {
    const mocked = fixture(current, fixtures);
    const result = await runReadParity('delta-bundle', { ...params, from_date: '2026-01-01', since: baseline.checkpoint },
      { call: mocked.call, stateDir });
    assert.equal(result.checkpoint, null, reason);
    assert.deepEqual(result.removed_from_local_scope, [], reason);
    assert.equal(result.complete, false, reason);
    assert.ok(result.incomplete_reasons.includes(reason));
    const pages = mocked.calls.filter(({ method }) => method === 'context_bundle');
    assert.ok(pages.every(({ params }) => params.limit === pages[0].params.limit));
  }
  assert.equal((await readdir(stateDir)).filter(name => name.endsWith('.json')).length, 1);
});

test('delta uses complete pages at a fixed page size and distinguishes horizon changes and scope mismatches',
  macCacheState, async t => {
  const stateDir = await state(t);
  const rows = Array.from({ length: 101 }, (_, i) => item(String(i), SNAP1));
  const first = fixture(rows);
  const result = await runReadParity('delta-bundle', {}, { call: first.call, stateDir });
  assert.equal(result.scope_traversal_complete, true);
  assert.equal(result.added.length, 101);
  assert.deepEqual(first.calls.map(({ method, params }) => [method, params.limit, params.cursor]), [
    ['context_bundle', 100, undefined], ['context_bundle', 100, '100'], ['folders', 1, undefined]
  ]);
  const next = fixture([], { snapshot: SNAP2, horizon: '2026-10-01' });
  const mismatch = await runReadParity('delta-bundle', { since: result.checkpoint }, { call: next.call, stateDir });
  assert.equal(mismatch.removal_status, 'cache_horizon_changed');
  assert.deepEqual(mismatch.removed_from_local_scope, []);
  assert.match(mismatch.checkpoint, /^[a-f0-9]{64}$/);
  await assert.rejects(runReadParity('delta-bundle', { folder: 'sent', since: result.checkpoint },
    { call: next.call, stateDir }), { code: 'CHECKPOINT_SCOPE_MISMATCH' });
  const differentAccount = fixture([], { snapshot: SNAP2, aliasRoot: 'other-inbox-root' });
  await assert.rejects(runReadParity('delta-bundle', { since: result.checkpoint },
    { call: differentAccount.call, stateDir }), { code: 'CHECKPOINT_SCOPE_MISMATCH' });
});

test('delta refuses symlink checkpoint directories, symlink files, and altered checkpoint contents', macCacheState,
  async t => {
  const stateDir = await state(t);
  const { call } = fixture([]);
  const baseline = await runReadParity('delta-bundle', {}, { call, stateDir });
  const file = path.join(stateDir, `checkpoint-${baseline.checkpoint}.json`);
  await writeFile(file, '{}');
  await assert.rejects(runReadParity('delta-bundle', { since: baseline.checkpoint }, { call, stateDir }),
    { code: 'INVALID_CHECKPOINT' });

  const target = path.join(path.dirname(stateDir), 'target');
  const alias = path.join(path.dirname(stateDir), 'alias');
  await mkdir(target, { mode: 0o700 });
  await symlink(target, alias);
  await assert.rejects(runReadParity('delta-bundle', {}, { call, stateDir: alias }), { code: 'UNSAFE_CHECKPOINT_STATE' });
  const candidate = 'c'.repeat(64);
  await symlink(file, path.join(stateDir, `checkpoint-${candidate}.json`));
  await assert.rejects(runReadParity('delta-bundle', { since: candidate }, { call, stateDir }),
    { code: 'UNSAFE_CHECKPOINT_STATE' });
});

test('task presets classify explicit evidence, distinguish unknown unread, and summarize cached folders', async () => {
  const rows = [
    item('needs-review', SNAP1, { subject: 'Action required, please review', unread: true, attachment_count: 1 }),
    item('routine', SNAP1, { subject: 'Market risk weekly review', unread: false }),
    item('alert', SNAP1, { subject: '例行通知', preview: '集中度比例超标', unread: null,
      folder_id: id('folder', 'other', SNAP1), folder_path: 'Inbox/Monitoring' })
  ];
  const { call } = fixture(rows);
  const task = await runReadParity('task-bundle', {}, { call });
  assert.equal(task.preset, 'actionable');
  assert.equal(task.unmatched, 2);
  assert.deepEqual(task.items.map(row => [row.id, row.reasons]),
    [[rows[0].id, [{ rule: 'action_requested', field: 'subject' }]]]);
  const risk = await runReadParity('task-bundle', { preset: 'risk' }, { call });
  assert.deepEqual(risk.items.map(row => [row.id, row.reasons]),
    [[rows[2].id, [{ rule: 'limit_breach', field: 'preview' }]]]);
  const unread = await runReadParity('task-bundle', { preset: 'unread' }, { call });
  assert.equal(unread.matched, 1);
  assert.equal(unread.unmatched, 1);
  assert.equal(unread.unknown, 1);
  const summary = await runReadParity('task-bundle', { preset: 'folder-summary', limit: 1 }, { call });
  assert.equal(summary.total_folders_in_scan, 2);
  assert.equal(summary.truncated, true);
  assert.equal(summary.folders[0].known_cached_attachments, 1);
  assert.equal(summary.folders[0].count, 2);
});
