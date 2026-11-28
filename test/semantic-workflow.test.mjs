import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { runSemanticWorkflow } from '../src/semantic-workflow.mjs';

const originalSnapshot = 'a'.repeat(24);
const nextSnapshot = 'b'.repeat(24);
const originalModel = '1'.repeat(64);
const nextModel = '2'.repeat(64);

async function fixture(t) {
  const base = await mkdtemp(path.join(os.tmpdir(), 'outlook-semantic-'));
  t.after(() => rm(base, { recursive: true, force: true }));
  const stateDir = path.join(base, 'state');
  const env = { OUTLOOK_EMBEDDING_MODEL: 'test-embedding' };
  const mailbox = {
    snapshot: originalSnapshot,
    rows: [
      { id: 'invoice_id', subject: '账单', content: '境外差旅报销凭据' },
      { id: 'plan_id', subject: '应急预案', content: '水灾后转运受灾群众', truncated: true },
      { id: 'missing_id', subject: '缺失正文', content: null },
      { id: 'empty_id', subject: '空正文', content: '' },
      { id: 'unavailable_id', subject: '读取失败', unavailable: true }
    ]
  };
  const model = { digest: originalModel, cloudDisabled: true };
  const requests = [], serviceCalls = [], embedInputs = [];
  async function call(method, args) {
    serviceCalls.push({ method, args });
    const { snapshot, rows } = mailbox;
    if (method === 'context_bundle') {
      const offset = args.cursor ? Number(args.cursor) : 0;
      const end = offset + Math.min(args.limit, 2);
      return {
        backend: 'mac-cache', snapshot, total_matched: rows.length, complete: end >= rows.length,
        items: rows.slice(offset, end).map(({ id, subject }) => ({ id, subject })),
        next_cursor: end >= rows.length ? null : String(end),
        date_filter_incomplete: true, received_time_unknown: 2
      };
    }
    assert.equal(method, 'messages_bundle');
    assert.equal(args.max_chars, 4000);
    assert.ok(args.ids.length <= 8);
    return {
      snapshot,
      items: rows.filter(row => args.ids.includes(row.id) && !row.unavailable).map(({ id, subject, content, truncated }) => ({
        id, subject, sender: { name: 'Example Sender' }, received_at: '2026-11-18T10:00:00+08:00',
        folder_path: 'Inbox',
        body: { available: content !== null, text: content, status: content === null ? 'not_cached' : 'cached_html',
          truncated: Boolean(truncated), total_chars: content?.length }
      }))
    };
  }
  async function fetchImpl(url, { method, body, redirect }) {
    assert.equal(redirect, 'error');
    const request = new URL(url);
    assert.equal(request.origin, 'http://127.0.0.1:11434');
    requests.push(request.pathname);
    if (request.pathname === '/api/status') {
      assert.equal(method, 'GET');
      return Response.json({ cloud: { disabled: model.cloudDisabled } });
    }
    if (request.pathname === '/api/tags') {
      assert.equal(method, 'GET');
      return Response.json({ models: [{ name: 'test-embedding:latest', digest: model.digest }] });
    }
    assert.equal(request.pathname, '/api/embed');
    assert.equal(method, 'POST');
    const args = JSON.parse(body);
    assert.equal(args.model, 'test-embedding:latest');
    assert.equal(args.truncate, false);
    embedInputs.push(...args.input);
    return Response.json({
      embeddings: args.input.map(text => /水灾|emergency/i.test(text) ? [0, 1] : [1, 0])
    });
  }
  return { base, stateDir, env, mailbox, model, requests, serviceCalls, embedInputs,
    options: { stateDir, env, call, fetchImpl } };
}

test('sync indexes only verified body text, pages within the requested scope, and reports limits', async t => {
  const f = await fixture(t);
  assert.deepEqual(await runSemanticWorkflow('status', {}, f.options), { status: 'not_indexed' });
  assert.deepEqual(f.requests, []);
  const result = await runSemanticWorkflow('sync', {
    folder: 'all', include_subfolders: false, from_date: '2026-11-18', scan_limit: 4
  }, f.options);
  assert.deepEqual(f.requests, ['/api/status', '/api/tags', '/api/embed']);
  assert.equal(result.snapshot_verified, true);
  assert.equal(result.coverage.scanned, 4);
  assert.equal(result.coverage.matched_in_local_metadata, 5);
  assert.equal(result.coverage.scope_traversal_complete, false);
  assert.deepEqual(
    [result.coverage.indexed, result.coverage.truncated, result.coverage.missing, result.coverage.empty],
    [1, 1, 1, 1]
  );
  assert.equal(result.coverage.body_coverage_complete, false);
  assert.equal(result.coverage.mailbox_complete, false);
  assert.equal(result.coverage.date_filter_incomplete, true);
  assert.ok(f.serviceCalls.some(({ method, args }) => method === 'context_bundle' && args.cursor));
  for (const entry of f.serviceCalls.filter(row => row.method === 'context_bundle')) {
    assert.equal(entry.args.folder, 'all');
    assert.equal(entry.args.include_subfolders, false);
    assert.equal(entry.args.from_date, '2026-11-18');
    assert.equal(entry.args.mode, 'metadata');
  }
  const index = path.join(f.stateDir, 'index.json');
  if (process.platform !== 'win32') {
    assert.equal((await stat(index)).mode & 0o777, 0o600);
    assert.equal((await stat(f.stateDir)).mode & 0o777, 0o700);
  }
  const raw = await readFile(index, 'utf8');
  assert.ok(!raw.includes('境外差旅'));
  assert.ok(!raw.includes('应急预案'));
  assert.ok(!raw.includes('Example Sender'));
  const status = await runSemanticWorkflow('status', {}, f.options);
  assert.equal(status.snapshot_verified, false);
  assert.equal(status.coverage.truncated, 1);
});

test('semantic ranking verifies each returned body against its pinned cache ID', async t => {
  const f = await fixture(t);
  await runSemanticWorkflow('sync', {}, f.options);
  const semantic = await runSemanticWorkflow('search', { query: 'emergency transport', limit: 2 }, f.options);
  assert.deepEqual(semantic.items.map(item => item.id), ['plan_id', 'invoice_id']);
  assert.equal(semantic.items[0].body_coverage, 'truncated');
  assert.equal(semantic.items[0].verified_current, true);
  assert.ok(semantic.items[0].score > semantic.items[1].score);
  f.mailbox.rows[1].content = 'changed after the prior embedding';
  const changed = await runSemanticWorkflow('search', { query: 'emergency transport', limit: 1 }, f.options);
  assert.deepEqual(changed.items.map(item => item.id), ['invoice_id']);
  assert.equal(changed.skipped_unverifiable, 1);
});

test('a new snapshot requires sync, which reuses embeddings by exact verified content', async t => {
  const f = await fixture(t);
  await runSemanticWorkflow('sync', {}, f.options);
  assert.equal(f.embedInputs.length, 2);
  f.mailbox.snapshot = nextSnapshot;
  f.mailbox.rows.forEach(row => { row.id = `current_${row.id}`; });
  await assert.rejects(runSemanticWorkflow('search', { query: 'emergency' }, f.options), { code: 'SEMANTIC_INDEX_STALE' });
  assert.equal(f.embedInputs.length, 2);
  const refresh = await runSemanticWorkflow('sync', {}, f.options);
  assert.equal(refresh.reused_embeddings, 2);
  assert.equal(refresh.snapshot, nextSnapshot);
  assert.equal(f.embedInputs.length, 2);
  const current = await runSemanticWorkflow('search', { query: 'emergency' }, f.options);
  assert.equal(current.items[0].id, 'current_plan_id');
  f.model.digest = nextModel;
  await assert.rejects(runSemanticWorkflow('search', { query: 'emergency' }, f.options), { code: 'SEMANTIC_MODEL_CHANGED' });
  await assert.rejects(runSemanticWorkflow('search', { query: 'emergency', folder: 'sent' }, f.options), { code: 'SEMANTIC_SCOPE_CHANGED' });
});

test('refuses non-loopback origins and cloud-enabled servers before accessing a mailbox', async t => {
  const f = await fixture(t);
  const options = { ...f.options, env: { ...f.env, OUTLOOK_OLLAMA_URL: 'http://external.example' } };
  await assert.rejects(runSemanticWorkflow('sync', {}, options), { code: 'INVALID_EMBEDDING_ORIGIN' });
  assert.deepEqual(f.requests, []);
  f.model.cloudDisabled = false;
  await assert.rejects(runSemanticWorkflow('sync', {}, f.options), { code: 'EMBEDDING_CLOUD_ENABLED' });
  assert.deepEqual(f.requests, ['/api/status']);
  assert.deepEqual(f.serviceCalls, []);
});

test('supports the Apple local provider without contacting Ollama', async t => {
  const f = await fixture(t);
  const appleInputs = [];
  const options = { ...f.options, env: { OUTLOOK_EMBEDDING_PROVIDER: 'apple' },
    createAppleProvider: async ({ stateDir }) => {
      assert.equal(stateDir, f.stateDir);
      return { model: 'apple-test-en', digest: originalModel, async embed(inputs) {
        appleInputs.push(...inputs);
        return inputs.map(text => /水灾|emergency/i.test(text) ? [0, 1] : [1, 0]);
      } };
    } };
  const result = await runSemanticWorkflow('sync', {}, options);
  assert.equal(result.provider, 'apple-natural-language');
  const search = await runSemanticWorkflow('search', { query: 'emergency transport' }, options);
  assert.equal(search.items[0].id, 'plan_id');
  assert.equal(appleInputs.length, 3);
  assert.deepEqual(f.requests, []);
});

test('a local embedding model can decline one message without discarding valid vectors', async t => {
  const f = await fixture(t);
  const options = { ...f.options, env: { OUTLOOK_EMBEDDING_PROVIDER: 'apple' },
    createAppleProvider: async () => ({
      model: 'apple-test-en', digest: originalModel,
      embed: async inputs => inputs.map(text => text.includes('水灾') ? null : [1, 0])
    }) };
  const result = await runSemanticWorkflow('sync', {}, options);
  assert.equal(result.coverage.indexed, 1);
  assert.equal(result.coverage.embedding_unavailable, 1);
  const search = await runSemanticWorkflow('search', { query: 'reimburse expenses' }, options);
  assert.deepEqual(search.items.map(item => item.id), ['invoice_id']);
  await assert.rejects(runSemanticWorkflow('search', { query: '水灾' }, options), { code: 'EMBEDDING_INPUT_UNSUPPORTED' });
});

test('semantic state refuses symlinked indexes without modifying their target', async t => {
  const f = await fixture(t);
  await runSemanticWorkflow('sync', {}, f.options);
  const target = path.join(f.base, 'other.json');
  await writeFile(target, 'untouched', { mode: 0o600 });
  const index = path.join(f.stateDir, 'index.json');
  await rm(index);
  await symlink(target, index);
  await assert.rejects(runSemanticWorkflow('status', {}, f.options));
  assert.equal(await readFile(target, 'utf8'), 'untouched');
});
