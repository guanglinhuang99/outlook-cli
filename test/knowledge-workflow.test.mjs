import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { link, mkdir, mkdtemp, readFile, readdir, realpath, rm, stat, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { runKnowledgeWorkflow } from '../src/knowledge-workflow.mjs';

const SNAPSHOT = 'a'.repeat(24);
const SCOPE = 'verified_inbox_sent_cache';
const ACCOUNT_SCOPE = 'verified_account_cache';
const BODY = 'Meeting notes: quarterly exposure and market risk.';
const ATTACHMENT = 'Spreadsheet: inflation ladder and cashflows.';
const sha = value => createHash('sha256').update(value).digest('hex');

function id(number, snapshot = SNAPSHOT) {
  return Buffer.from(JSON.stringify({ v: 3, backend: 'mac-cache', kind: 'message',
    key: 'hx_' + number, snapshot })).toString('base64url');
}

function message(id, body = BODY) {
  return {
    id, subject: 'Quarterly report', internet_message_id: '<synthetic@example.test>',
    sender: { name: 'Synthetic sender', smtp_address: 'example@example.test' },
    received_at: '2026-11-01T00:00:00Z', attachments_complete: false,
    attachments: [{ index: 1, filename: 'cashflows.csv', size_bytes: 100, status: 'cached_at_snapshot' },
      { index: 2, filename: 'uncached.pdf', size_bytes: null, status: 'unavailable' }],
    body: { available: body !== null, text: body, total_chars: body?.length ?? null,
      derived_from: body === null ? null : 'html', status: body === null ? 'unavailable' : 'available', truncated: false }
  };
}

function bundle(ids, { body = BODY, snapshot = SNAPSHOT, scope = SCOPE, errors = [] } = {}) {
  return {
    backend: 'mac-cache', data_scope: scope, snapshot, captured_at: '2026-11-01T00:01:00Z',
    items: ids.map(value => message(value, body)), errors
  };
}

async function fixture(t) {
  const tmp = await mkdtemp(path.join(os.tmpdir(), 'outlook-knowledge-test-'));
  t.after(() => rm(tmp, { recursive: true, force: true }));
  const stateDir = path.join(tmp, 'knowledge');
  const outDir = path.join(tmp, 'export');
  await mkdir(outDir);
  return { stateDir, outDir, tmp };
}

const cacheCall = async (method, params) => {
  assert.equal(method, 'messages_bundle');
  assert.equal(params.backend, 'mac-cache');
  return bundle(params.ids);
};

const fixtureAttachment = async ({ id, index, max_chars }) => ({
  id, index, backend: 'mac-cache', data_scope: SCOPE, sha256: sha('fixture attachment bytes'),
  size_bytes: 100, filename: 'cashflows.csv', format: 'text',
  text: ATTACHMENT.slice(0, max_chars), truncated: ATTACHMENT.length > max_chars
});

test('bundle limits each request to 25 IDs and clips to one shared UTF-16 text budget', async () => {
  const ids = Array.from({ length: 50 }, (_, index) => id(index));
  const calls = [];
  const result = await runKnowledgeWorkflow('bundle', {
    ids, text_budget_chars: 26, body_max_chars: 1, attachment_requests: [{ id: ids[49], index: 1 }]
  }, {
    call: async (method, params) => {
      calls.push(params.ids.length);
      assert.equal(method, 'messages_bundle');
      assert.equal(params.max_chars, 1);
      return bundle(params.ids, { body: '点X' });
    },
    readAttachmentText: () => assert.fail('attachments cannot read after the shared budget is spent')
  });
  assert.deepEqual(calls, [25, 25]);
  assert.equal(result.items.length, 50);
  assert.equal(result.items[0].body.text, '点');
  assert.equal(result.items[0].body.truncated, true);
  assert.equal(result.items[26].body.available, false);
  assert.equal(result.items[26].body.status, 'text_budget_exhausted');
  assert.equal(result.attachments[0].code, 'TEXT_BUDGET_EXHAUSTED');
  assert.equal(result.text_chars, 26);
  assert.equal(result.complete, false);
});

test('bundle rejects mixed snapshots, mixed scopes and unverified sources', async () => {
  const notCalled = () => assert.fail('mixed snapshots should fail before reading the cache');
  const ids = Array.from({ length: 26 }, (_, index) => id(index));
  await assert.rejects(runKnowledgeWorkflow('bundle', { ids: [ids[0], id(99, 'c'.repeat(24))] },
    { call: notCalled }), { code: 'SNAPSHOT_MISMATCH' });
  await assert.rejects(runKnowledgeWorkflow('bundle', { ids }, {
    call: async (method, { ids: selected }) => bundle(selected,
      { snapshot: selected.length === 25 ? SNAPSHOT : 'b'.repeat(24) })
  }), { code: 'CACHE_INTEGRITY_ERROR' });
  await assert.rejects(runKnowledgeWorkflow('bundle', { ids }, {
    call: async (method, { ids: selected }) => bundle(selected,
      { scope: selected.length === 25 ? SCOPE : ACCOUNT_SCOPE })
  }), { code: 'CACHE_INTEGRITY_ERROR' });
  await assert.rejects(runKnowledgeWorkflow('bundle', { ids: [ids[0]] }, {
    call: async (method, { ids: selected }) => bundle(selected, { scope: 'unverified_cache' })
  }), { code: 'CACHE_INTEGRITY_ERROR' });
});

test('bundle sends only explicit, verified attachment requests and uses the remaining text budget', async () => {
  const mid = id(4);
  const budget = BODY.length + 8;
  const readRequests = [];
  const result = await runKnowledgeWorkflow('bundle', {
    ids: [mid], text_budget_chars: budget,
    attachment_requests: [{ id: mid, index: 99 }, { id: mid, index: 1 }, { id: mid, index: 2 }]
  }, {
    call: cacheCall,
    readAttachmentText: async params => {
      readRequests.push(params);
      return fixtureAttachment(params);
    }
  });
  assert.deepEqual(readRequests, [{ id: mid, index: 1, max_chars: 8, backend: 'mac-cache' }]);
  assert.equal(result.attachments[0].code, 'ATTACHMENT_NOT_VERIFIED');
  assert.equal(result.attachments[1].text, ATTACHMENT.slice(0, 8));
  assert.equal(result.attachments[1].text_sha256, sha(ATTACHMENT.slice(0, 8)));
  assert.equal(result.attachments[1].sha256, sha('fixture attachment bytes'));
  assert.equal(result.attachments[1].truncated, true);
  assert.equal(result.attachments[2].code, 'TEXT_BUDGET_EXHAUSTED');
  assert.equal(result.text_chars, budget);
});

test('account bundles require exact attachment scope before exposing text', async () => {
  const mid = id(7);
  const params = { ids: [mid], attachment_requests: [{ id: mid, index: 1 }] };
  const call = async (method, input) => bundle(input.ids, { scope: ACCOUNT_SCOPE });
  const mismatch = await runKnowledgeWorkflow('bundle', params, { call, readAttachmentText: fixtureAttachment });
  assert.equal(mismatch.data_scope, ACCOUNT_SCOPE);
  assert.equal(mismatch.attachments[0].code, 'ATTACHMENT_INTEGRITY_ERROR');
  assert.equal(mismatch.attachments[0].text, null);
  assert.equal(mismatch.attachments[0].sha256, null);
  assert.equal(mismatch.complete, false);
  const accepted = await runKnowledgeWorkflow('bundle', params, { call, readAttachmentText: async request =>
    ({ ...await fixtureAttachment(request), data_scope: ACCOUNT_SCOPE }) });
  assert.equal(accepted.attachments[0].text, ATTACHMENT);
  assert.equal(accepted.complete, true);
});

test('account bundles enforce the negotiated scope through the real attachment text workflow', async () => {
  const mid = id(9);
  const text = ATTACHMENT.padEnd(100, ' ');
  const bytes = Buffer.from(text);
  const state = { port: 19317, token: 'fixture', pid: 9, startedAt: 'fixture' };
  const listing = { backend: 'mac-cache', data_scope: ACCOUNT_SCOPE, snapshot: SNAPSHOT };
  const createCall = chunkScope => async (method, input, options) => {
    if (method === 'messages_bundle') return bundle(input.ids, { scope: ACCOUNT_SCOPE });
    assert.equal(input.id, mid);
    assert.equal(options.state, state);
    if (method === 'attachments') return { ...listing, attachments: [{ ...message(mid).attachments[0], type: 1 }] };
    assert.equal(method, 'read_attachment');
    return { ...listing, data_scope: chunkScope, filename: 'cashflows.csv', offset: 0,
      size_bytes: bytes.length, chunk_size: bytes.length, next_offset: null,
      sha256: sha(bytes), encoding: 'base64', data: bytes.toString('base64') };
  };
  const params = { ids: [mid], attachment_requests: [{ id: mid, index: 1 }] };
  const readState = async () => state;
  const accepted = await runKnowledgeWorkflow('bundle', params, { call: createCall(ACCOUNT_SCOPE), readState });
  assert.equal(accepted.attachments[0].text, text);
  const mismatch = await runKnowledgeWorkflow('bundle', params, { call: createCall(SCOPE), readState });
  assert.equal(mismatch.attachments[0].code, 'ATTACHMENT_INTEGRITY_ERROR');
  assert.equal(mismatch.attachments[0].text, null);
});

test('mixed archive search, source and export preserve each record scope', async t => {
  const { stateDir, outDir } = await fixture(t);
  const mid = id(17);
  const options = { stateDir, call: cacheCall, platform: 'linux' };
  const empty = await runKnowledgeWorkflow('status', {}, options);
  assert.equal(empty.data_scope, null);
  assert.deepEqual(empty.data_scopes, []);
  const inbox = await runKnowledgeWorkflow('sync', { ids: [mid] }, options);
  const account = await runKnowledgeWorkflow('sync', { ids: [mid] }, {
    ...options, call: async (method, input) => bundle(input.ids, { scope: ACCOUNT_SCOPE })
  });
  const archiveIds = [inbox.evidence[0].archive_id, account.evidence[0].archive_id];
  assert.notEqual(archiveIds[0], archiveIds[1]);
  assert.equal(account.data_scope, ACCOUNT_SCOPE);
  const status = await runKnowledgeWorkflow('status', {}, options);
  assert.equal(status.data_scope, 'mixed_verified_cache');
  assert.deepEqual(status.data_scopes, [ACCOUNT_SCOPE, SCOPE]);
  const hits = await runKnowledgeWorkflow('search', { query: 'risk' }, options);
  assert.deepEqual(hits.data_scopes, status.data_scopes);
  assert.deepEqual(new Set(hits.items.map(item => item.source.data_scope)), new Set(status.data_scopes));
  const source = await runKnowledgeWorkflow('source', { archive_id: archiveIds[1] }, options);
  assert.equal(source.source.data_scope, ACCOUNT_SCOPE);
  const exported = await runKnowledgeWorkflow('export', { ids: archiveIds, out_dir: outDir }, options);
  assert.equal(exported.data_scope, 'mixed_verified_cache');
  const payload = JSON.parse(await readFile(exported.path, 'utf8'));
  assert.deepEqual(payload.data_scopes, status.data_scopes);
  assert.deepEqual(payload.manifest.map(item => item.source.data_scope), [SCOPE, ACCOUNT_SCOPE]);
});

test('index keeps cache provenance, literal lookups, explicit unknowns and checked private exports', async t => {
  const { stateDir, outDir } = await fixture(t);
  const a = id(1);
  const options = {
    stateDir, call: cacheCall, readAttachmentText: async params => {
      if (params.index === 2) throw Object.assign(new Error('do not store internal extractor errors'),
        { code: 'ATTACHMENT_NOT_CACHED' });
      return fixtureAttachment(params);
    },
    platform: 'linux'
  };
  assert.equal((await runKnowledgeWorkflow('status', {}, options)).count, 0);
  const ingest = { ids: [a], attachment_requests: [{ id: a, index: 1 }, { id: a, index: 2 }, { id: a, index: 99 }] };
  const synced = await runKnowledgeWorkflow('index', ingest, options);
  assert.equal(synced.stored, 3);
  assert.equal(synced.complete, false);
  assert.deepEqual(synced.attachment_errors.map(item => item.code),
    ['ATTACHMENT_NOT_CACHED', 'ATTACHMENT_NOT_VERIFIED']);
  assert.equal((await runKnowledgeWorkflow('sync', ingest, options)).stored, 0);
  const status = await runKnowledgeWorkflow('status', {}, options);
  assert.deepEqual(status.snapshots, [SNAPSHOT]);
  assert.equal(status.text_available, 2);

  const hits = await runKnowledgeWorkflow('search', { query: 'INFLATION', kind: 'attachment' }, options);
  assert.equal(hits.matches, 1);
  assert.equal(hits.items[0].filename, 'cashflows.csv');
  const archiveId = hits.items[0].archive_id;
  const detail = await runKnowledgeWorkflow('read', { archive_id: archiveId }, options);
  assert.equal(detail.content.text, ATTACHMENT);
  const source = await runKnowledgeWorkflow('source', { archive_id: archiveId }, options);
  assert.equal(source.source.snapshot, SNAPSHOT);
  assert.equal(source.source.date_basis, 'received_at');
  assert.equal(source.source.attachment_bytes_sha256, sha('fixture attachment bytes'));
  assert.equal(source.source.original_rfc822_available, false);
  assert.equal(source.text_sha256, sha(ATTACHMENT));
  assert.equal('text' in source, false);
  const unavailable = await runKnowledgeWorkflow('read', { archive_id: synced.evidence[2].archive_id }, options);
  assert.equal(unavailable.content.text, null);
  assert.equal(unavailable.content.code, 'ATTACHMENT_NOT_CACHED');
  assert.equal(unavailable.content.available, false);

  const args = { ids: [archiveId, synced.evidence[0].archive_id], out_dir: outDir };
  const exported = await runKnowledgeWorkflow('export', args, options);
  const raw = await readFile(exported.path);
  const { payload_sha256, ...payload } = JSON.parse(raw);
  assert.equal(exported.sha256, sha(raw));
  assert.equal(payload_sha256, sha(JSON.stringify(payload)));
  assert.equal(payload.original_rfc822_available, false);
  assert.equal(payload.records[0].content.text_sha256, sha(ATTACHMENT));
  assert.equal(payload.manifest.length, 2);
  assert.equal(exported.format, 'outlook-cache-derived-evidence-v1');
  if (process.platform !== 'win32') {
    assert.equal((await stat(exported.path)).mode & 0o777, 0o600);
  }
  await assert.rejects(runKnowledgeWorkflow('export', args, options), { code: 'KNOWLEDGE_EXPORT_EXISTS' });
  assert.equal(sha(await readFile(exported.path)), exported.sha256);
});

test('semantic and hybrid queries report model compatibility and literal coverage independently', async t => {
  const { stateDir } = await fixture(t);
  const apple = {
    model: 'apple:fixture', digest: 'a'.repeat(64),
    async embed(texts) {
      return texts.map(text => text.includes('cashflow') || text === 'liquidity' ? [1, 0] : [0, 4]);
    }
  };
  const options = { stateDir, call: cacheCall, readAttachmentText: fixtureAttachment,
    createAppleProvider: async () => apple };
  const mid = id(10);
  await runKnowledgeWorkflow('sync', { ids: [mid], semantic: true,
    attachment_requests: [{ id: mid, index: 1 }] }, options);
  await runKnowledgeWorkflow('sync', { ids: [id(11)] }, options);
  const semantic = await runKnowledgeWorkflow('search', { query: 'liquidity', mode: 'semantic' }, options);
  assert.equal(semantic.items[0].kind, 'attachment');
  assert.equal(semantic.items[0].semantic_score, 1);
  assert.deepEqual(semantic.semantic_coverage, { eligible: 3, indexed: 2, complete: false });
  const otherModel = { ...options, createAppleProvider: async () => ({ ...apple, digest: 'b'.repeat(64) }) };
  const hybrid = await runKnowledgeWorkflow('search', { query: 'exposure', mode: 'hybrid' }, otherModel);
  assert.equal(hybrid.items.length, 2);
  assert.ok(hybrid.items.every(hit => hit.literal_match && hit.semantic_score === null));
  assert.deepEqual(hybrid.semantic_coverage, { eligible: 3, indexed: 0, complete: false });
  const unavailable = { ...options, createAppleProvider: async () => {
    throw Object.assign(new Error('fixture model unavailable'), { code: 'EMBEDDING_PROVIDER_UNAVAILABLE' });
  } };
  const fallback = await runKnowledgeWorkflow('search', { query: 'exposure', mode: 'hybrid' }, unavailable);
  assert.equal(fallback.items.length, 2);
  assert.equal(fallback.semantic_coverage.error, 'EMBEDDING_PROVIDER_UNAVAILABLE');
  await assert.rejects(runKnowledgeWorkflow('search', { query: 'exposure', mode: 'semantic' }, unavailable),
    { code: 'EMBEDDING_PROVIDER_UNAVAILABLE' });
});

test('concurrent different syncs both publish and tampered evidence cannot be read or exported', async t => {
  const { stateDir } = await fixture(t);
  const options = { stateDir, call: cacheCall };
  const [first] = await Promise.all([0, 1, 2, 3, 4].map(index =>
    runKnowledgeWorkflow('sync', { ids: [id(index)] }, options)));
  assert.equal((await runKnowledgeWorkflow('status', {}, options)).count, 5);
  const segment = path.join(stateDir, 'segment-000.json');
  const data = JSON.parse(await readFile(segment, 'utf8'));
  // A crash after publishing may retain exactly the staged name referenced by this segment.
  const staging = path.join(stateDir, data.pending);
  await link(segment, staging);
  assert.equal((await runKnowledgeWorkflow('status', {}, options)).count, 5);
  await rm(staging);
  data.entries[0].record.content.text = 'tampering';
  await writeFile(segment, JSON.stringify(data));
  await assert.rejects(runKnowledgeWorkflow('read', { archive_id: first.evidence[0].archive_id }, options),
    { code: 'KNOWLEDGE_INTEGRITY_ERROR' });
});

test('export checks the recorded custom protected mail roots before publishing', async t => {
  const { stateDir, outDir } = await fixture(t);
  const options = { stateDir, call: cacheCall };
  const result = await runKnowledgeWorkflow('sync', { ids: [id(8)] }, options);
  const output = await realpath(outDir);
  await assert.rejects(runKnowledgeWorkflow('export', { ids: [result.evidence[0].archive_id], out_dir: outDir }, {
    ...options, platform: 'darwin',
    readState: async ({ requireMacProtectedRoots }) => {
      assert.equal(requireMacProtectedRoots, true);
      return { protected_mail_roots: [output] };
    },
    assertSafePath: async (input, { protectedRoots = [] } = {}) => {
      const canonical = await realpath(input);
      if (protectedRoots.includes(canonical)) throw Object.assign(new Error('Protected'), { code: 'MAC_PATH_FORBIDDEN' });
      return canonical;
    }
  }), { code: 'MAC_PATH_FORBIDDEN' });
  assert.deepEqual(await readdir(outDir), []);
});
