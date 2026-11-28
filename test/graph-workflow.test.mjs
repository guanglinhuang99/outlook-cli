import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { GRAPH_ACTIONS, runGraphWorkflow } from '../src/graph-workflow.mjs';

const graph = 'https://graph.microsoft.com/v1.0';

function response(value, status = 200) {
  return value === undefined ? new Response(null, { status }) : Response.json(value, { status });
}

async function graphFixture(t, handler) {
  const parent = await mkdtemp(path.join(os.tmpdir(), 'outlook-graph-workflow-'));
  t.after(() => rm(parent, { recursive: true, force: true }));
  const stateDirectory = path.join(parent, 'graph-state');
  const requests = [];
  const fetch = async (url, options) => {
    assert.equal(url.origin, 'https://graph.microsoft.com');
    assert.equal(options.headers.Prefer, 'IdType="ImmutableId", outlook.body-content-type="text"');
    assert.equal(options.redirect, 'error');
    assert.ok(options.headers.Authorization.startsWith('Bearer '));
    const request = {
      pathname: url.pathname.replace(/^\/v1\.0/, ''),
      query: url.searchParams,
      method: options.method,
      data: options.body ? JSON.parse(options.body) : undefined
    };
    requests.push(request);
    if (request.pathname === '/me/mailFolders/msgfolderroot') {
      return response({ id: options.headers.Authorization === 'Bearer another-account' ? 'other-root' : 'first-root' });
    }
    return handler(request);
  };
  const call = (action, input, token = 'fixture-token') =>
    runGraphWorkflow(action, input, { token, fetch, stateDirectory });
  return { call, requests, stateDirectory };
}

test('Graph discovery, status, token checking, and configuration do not expose a token', async t => {
  assert.ok(GRAPH_ACTIONS.includes('draft-send'));
  assert.deepEqual(await runGraphWorkflow('status', {}, { token: '' }),
    { backend: 'graph', configured: false, verified: false });
  assert.deepEqual(await runGraphWorkflow('status', {}, { token: 'test secret' }),
    { backend: 'graph', configured: false, verified: false });
  await assert.rejects(runGraphWorkflow('status', { verify: true }, { token: '' }),
    { code: 'GRAPH_AUTH_REQUIRED' });
  const { call, requests } = await graphFixture(t, req => {
    assert.equal(req.pathname, '/me/mailFolders');
    assert.equal(req.query.get('$select'), 'id');
    return response({ value: [{ id: 'inbox' }] });
  });
  assert.deepEqual(await call('status', {}), { backend: 'graph', configured: true, verified: false });
  assert.equal(requests.length, 0);
  assert.deepEqual(await call('status', { verify: true }), { backend: 'graph', configured: true, verified: true });
});

test('Graph list retains date and unread semantics and refuses any cursor outside the selected resource', async t => {
  const next = graph + '/me/mailFolders/inbox/messages?$skiptoken=abc';
  const { call, requests } = await graphFixture(t, req => {
    assert.equal(req.pathname, '/me/mailFolders/inbox/messages');
    if (req.query.has('$skiptoken')) return response({ value: [{ id: 'second' }] });
    assert.equal(req.query.get('$top'), '2');
    assert.equal(req.query.get('$orderby'), 'receivedDateTime desc');
    assert.equal(req.query.get('$filter'),
      'receivedDateTime ge 2026-11-24T16:00:00.000Z and receivedDateTime lt 2026-11-25T16:00:00.000Z and isRead eq false');
    return response({ value: [{ id: 'first' }], '@odata.nextLink': next });
  });
  const page = await call('list', {
    folder: 'inbox', limit: 2, unread: true, after: '2026-11-25T00:00:00+08:00', before: '2026-11-26T00:00:00+08:00'
  });
  assert.equal(page.next_cursor, next);
  assert.deepEqual(await call('list', { folder: 'inbox', cursor: next }),
    { backend: 'graph', items: [{ id: 'second' }], next_cursor: null });
  await assert.rejects(call('list', { folder: 'sent', cursor: next }), { code: 'GRAPH_INVALID_CURSOR' });
  await assert.rejects(call('list', { folder: 'inbox', cursor: next.replace('graph.microsoft.com', 'example.test') }),
    { code: 'GRAPH_INVALID_CURSOR' });
  await assert.rejects(call('list', { after: '2026-11-25T01:00:00Z', before: '2026-11-25T01:00:00Z' }),
    { code: 'INVALID_ARGUMENT' });
  const oldId = Buffer.from(JSON.stringify({ v: 3, backend: 'mac-cache', kind: 'message', key: 'x' })).toString('base64url');
  await assert.rejects(call('read', { id: oldId }), { code: 'INVALID_ARGUMENT' });
  assert.equal(requests.length, 2);
});

test('Graph resolution escapes the query, matches the exact Internet Message-ID, and fails on ambiguity', async t => {
  const mid = "<name's@example.test>";
  const next = graph + '/me/messages?$skiptoken=next';
  let ambiguous = false;
  const { call } = await graphFixture(t, req => {
    assert.equal(req.pathname, '/me/messages');
    if (req.query.has('$skiptoken')) {
      const matches = ambiguous ? [{ id: 'two', internetMessageId: mid }] : [];
      return response({ value: matches });
    }
    assert.equal(req.query.get('$filter'), "internetMessageId eq '<name''s@example.test>'");
    return response({ value: [{ id: 'near', internetMessageId: mid.toUpperCase() },
      { id: 'one', internetMessageId: mid }], '@odata.nextLink': next });
  });
  const found = await call('resolve', { internet_message_id: mid });
  assert.equal(found.id, 'one');
  ambiguous = true;
  await assert.rejects(call('resolve', { internet_message_id: mid }), { code: 'GRAPH_ID_AMBIGUOUS' });
});

test('draft creation, replies, forwarding, editing and discard use Graph draft semantics', async t => {
  const actions = new Map([
    ['POST /me/messages', { subject: 'Subject', body: { contentType: 'Text', content: 'Private draft body' },
      toRecipients: [{ emailAddress: { address: 'recipient@example.test' } }] }],
    ['POST /me/messages/source/createReply', { comment: 'Draft a reply' }],
    ['POST /me/messages/source/createReplyAll', {}],
    ['POST /me/messages/source/createForward', { comment: 'Draft a forward',
      toRecipients: [{ emailAddress: { address: 'another@example.test' } }] }],
    ['PATCH /me/mailFolders/drafts/messages/draft', { subject: 'Changed', ccRecipients: [] }],
    ['DELETE /me/mailFolders/drafts/messages/draft', undefined]
  ]);
  const { call, requests, stateDirectory } = await graphFixture(t, req => {
    if (req.method === 'GET') return response({ id: 'draft', isDraft: true });
    const key = req.method + ' ' + req.pathname;
    assert.ok(actions.has(key), key);
    assert.deepEqual(req.data, actions.get(key));
    actions.delete(key);
    return req.method === 'DELETE' ? response(undefined, 204)
      : response({ id: 'draft', body: { content: 'Private draft body' } }, req.method === 'POST' ? 201 : 200);
  });
  assert.equal((await call('compose', { request_id: 'compose', subject: 'Subject',
    body: 'Private draft body', to: ['recipient@example.test'] })).id, 'draft');
  await call('reply', { request_id: 'reply', id: 'source', comment: 'Draft a reply' });
  await call('reply-all', { request_id: 'replyall', id: 'source' });
  await call('forward', { request_id: 'forward', id: 'source', comment: 'Draft a forward', to: ['another@example.test'] });
  await call('draft-update', { request_id: 'edit', id: 'draft', subject: 'Changed', cc: [] });
  await assert.rejects(call('draft-discard', { request_id: 'discard', id: 'draft' }), { code: 'INVALID_ARGUMENT' });
  assert.equal((await call('draft-discard', { request_id: 'discard', id: 'draft', confirm: true })).status, 204);
  assert.equal(actions.size, 0);
  assert.equal(requests.filter(r => r.pathname.endsWith('/send')).length, 0);
  const journal = await Promise.all((await readdir(stateDirectory)).map(name => readFile(path.join(stateDirectory, name), 'utf8')));
  assert.ok(journal.every(entry => !entry.includes('Private draft body') && !entry.includes('@example.test')));
});

test('Graph mark, flag, move and batch use explicit operations and report partial results', async t => {
  const { call, requests } = await graphFixture(t, req => {
    if (req.pathname.endsWith('/move')) {
      assert.deepEqual(req.data, { destinationId: 'sentitems' });
      return response({ id: 'source' }, 201);
    }
    if (req.pathname === '/$batch') {
      assert.deepEqual(req.data.requests.map(({ id, method, url, dependsOn }) => [id, method, url, dependsOn]), [
        ['1', 'PATCH', '/me/messages/first', undefined], ['2', 'POST', '/me/messages/first/move', ['1']],
        ['3', 'PATCH', '/me/messages/second', undefined], ['4', 'POST', '/me/messages/second/move', ['3']]
      ]);
      return response({ responses: [
        { id: '4', status: 424 }, { id: '1', status: 200, body: { id: 'first' } },
        { id: '3', status: 403 }, { id: '2', status: 201, body: { id: 'first' } }
      ] });
    }
    assert.equal(req.method, 'PATCH');
    assert.deepEqual(req.data, req.pathname.endsWith('/read') ? { isRead: false } : { flag: { flagStatus: 'complete' } });
    return response({ id: req.pathname.split('/').at(-1) });
  });
  await call('mark-read', { request_id: 'mark-read', id: 'read', is_read: false });
  await call('flag', { request_id: 'flag', id: 'flag', flag: 'complete' });
  await assert.rejects(call('move', { request_id: 'move', id: 'source', folder: 'sent' }), { code: 'INVALID_ARGUMENT' });
  await call('move', { request_id: 'move', id: 'source', folder: 'sent', confirm: true });
  const batchParams = { request_id: 'batch', ids: ['first', 'second'], is_read: true, move_to: 'archive', confirm: true };
  const batch = await call('batch-update', batchParams);
  assert.equal(batch.complete, false);
  assert.deepEqual(batch.items.map(item => [item.id, item.action, item.status]), [
    ['first', 'update', 200], ['first', 'move', 201], ['second', 'update', 403], ['second', 'move', 424]
  ]);
  assert.equal((await call('batch-update', batchParams)).replayed, true);
  assert.equal(requests.filter(r => r.pathname === '/$batch').length, 1);
  await assert.rejects(call('mark-read', { request_id: 'mark-read', id: 'read', is_read: true }),
    { code: 'GRAPH_REQUEST_CONFLICT' });
});

test('Graph send only accepts a confirmed existing draft and never repeats a successful send under another ID', async t => {
  let sends = 0;
  let permitSend;
  const gate = new Promise(resolve => { permitSend = resolve; });
  let started;
  const dispatched = new Promise(resolve => { started = resolve; });
  const { call } = await graphFixture(t, async req => {
    if (req.method === 'GET') {
      const id = req.pathname.split('/').at(-1);
      return response({ id, isDraft: id !== 'nondraft' });
    }
    assert.equal(req.pathname, '/me/messages/draft/send');
    assert.equal(req.data, undefined);
    sends += 1;
    started();
    await gate;
    return response(undefined, 202);
  });
  await assert.rejects(call('draft-send', { id: 'draft', request_id: 'missing-confirm' }), { code: 'INVALID_ARGUMENT' });
  await assert.rejects(call('draft-send', { id: 'nondraft', request_id: 'nondraft-send', confirm: true }), { code: 'GRAPH_NOT_DRAFT' });
  const first = call('draft-send', { id: 'draft', request_id: 'first-send', confirm: true });
  await dispatched;
  try {
    await assert.rejects(call('draft-send', { id: 'draft', request_id: 'concurrent-send', confirm: true }),
      { code: 'GRAPH_REQUEST_IN_DOUBT' });
  } finally {
    permitSend();
  }
  assert.equal((await first).status, 202);
  assert.equal((await call('draft-send', { id: 'draft', request_id: 'first-send', confirm: true })).replayed, true);
  assert.equal((await call('draft-send', { id: 'draft', request_id: 'different-send', confirm: true })).replayed, true);
  assert.equal(sends, 1);
});

test('an indeterminate Graph write stays reserved, and a separate account has its own ledger', async t => {
  let sends = 0;
  const { call } = await graphFixture(t, req => {
    if (req.method === 'GET') return response({ id: 'draft', isDraft: true });
    sends += 1;
    if (sends === 1) throw new Error('transport disconnected');
    return response(undefined, 202);
  });
  const input = { id: 'draft', request_id: 'send-1', confirm: true };
  await assert.rejects(call('draft-send', input), { code: 'GRAPH_REQUEST_IN_DOUBT' });
  await assert.rejects(call('draft-send', input), { code: 'GRAPH_REQUEST_IN_DOUBT' });
  await assert.rejects(call('draft-send', { ...input, request_id: 'send-2' }), { code: 'GRAPH_REQUEST_IN_DOUBT' });
  assert.equal(sends, 1);
  assert.equal((await call('draft-send', input, 'another-account')).status, 202);
  assert.equal(sends, 2);
});

test('a definite Graph rejection permits a new explicit request after the mailbox is checked', async t => {
  let attempts = 0;
  const { call } = await graphFixture(t, req => {
    if (req.method === 'GET') return response({ id: 'draft', isDraft: true });
    attempts += 1;
    if (attempts === 1) return response({ error: { message: 'Contains private data' } }, 403);
    return response(undefined, 202);
  });
  const input = { id: 'draft', request_id: 'first', confirm: true };
  await assert.rejects(call('draft-send', input), error => error.code === 'GRAPH_FORBIDDEN'
    && !error.message.includes('Contains private data'));
  await assert.rejects(call('draft-send', input), { code: 'GRAPH_FORBIDDEN' });
  assert.equal((await call('draft-send', { ...input, request_id: 'second' })).status, 202);
  assert.equal(attempts, 2);
});

test('Graph attachment reads validate metadata, bound bytes, and verify the returned digest', async t => {
  const bytes = Buffer.from('Attachment byte content');
  let claimedSize = bytes.length;
  let type = '#microsoft.graph.fileAttachment';
  const { call, requests } = await graphFixture(t, req => {
    const attachment = { id: 'file', name: 'sample.txt', size: claimedSize, contentType: 'text/plain', '@odata.type': type };
    if (req.pathname === '/me/messages/source/attachments') return response({ value: [attachment] });
    if (req.pathname.endsWith('/$value')) return new Response(bytes);
    return response(attachment);
  });
  assert.equal((await call('attachments', { id: 'source' })).items[0].id, 'file');
  const args = { id: 'source', attachment_id: 'file' };
  const result = await call('attachment-read', args);
  assert.deepEqual(Buffer.from(result.data, 'base64'), bytes);
  assert.equal(result.size_bytes, bytes.length);
  assert.equal(result.sha256, await crypto.subtle.digest('SHA-256', bytes).then(sum => Buffer.from(sum).toString('hex')));
  await assert.rejects(call('attachment-read', { ...args, max_bytes: 3 }), { code: 'ATTACHMENT_TOO_LARGE' });
  claimedSize = 3;
  await assert.rejects(call('attachment-read', { ...args, max_bytes: 5 }), { code: 'ATTACHMENT_TOO_LARGE' });
  await assert.rejects(call('attachment-read', args), { code: 'GRAPH_ATTACHMENT_CHANGED' });
  type = '#microsoft.graph.referenceAttachment';
  await assert.rejects(call('attachment-read', args), { code: 'ATTACHMENT_UNSUPPORTED' });
  assert.equal(requests.filter(req => req.pathname.endsWith('/$value')).length, 3);
});

test('Graph MIME returns bounded bytes from the read-only message representation', async t => {
  const bytes = Buffer.from('Subject: Sample\r\n\r\nMail body\r\n');
  const { call, requests } = await graphFixture(t, () => new Response(bytes));
  const result = await call('mime', { id: 'source' });
  assert.equal(result.source, 'graph_mime_representation');
  assert.equal(result.format, 'rfc822');
  assert.equal(result.size_bytes, bytes.length);
  assert.deepEqual(Buffer.from(result.data, 'base64'), bytes);
  assert.equal(result.sha256, await crypto.subtle.digest('SHA-256', bytes).then(sum => Buffer.from(sum).toString('hex')));
  await assert.rejects(call('mime', { id: 'source', max_bytes: 3 }), { code: 'MIME_TOO_LARGE' });
  assert.deepEqual(requests.map(({ method, pathname }) => [method, pathname]),
    [['GET', '/me/messages/source/$value'], ['GET', '/me/messages/source/$value']]);
});

test('adding an explicit small file stays in the draft and never sends', async t => {
  const { call, requests } = await graphFixture(t, req => {
    if (req.method === 'GET') return response({ id: 'draft', isDraft: true });
    assert.equal(req.pathname, '/me/mailFolders/drafts/messages/draft/attachments');
    assert.deepEqual(req.data, { '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'sample.txt', contentType: 'text/plain', contentBytes: 'aGVsbG8=' });
    return response({ id: 'file' }, 201);
  });
  const args = { id: 'draft', request_id: 'attach', name: 'sample.txt',
    content_type: 'text/plain', content_base64: 'aGVsbG8=' };
  await assert.rejects(call('draft-attach', args), { code: 'INVALID_ARGUMENT' });
  await assert.rejects(call('draft-attach', { ...args, confirm: true, content_base64: 'not a base64 string' }),
    { code: 'INVALID_ARGUMENT' });
  assert.equal((await call('draft-attach', { ...args, confirm: true })).id, 'file');
  assert.equal(requests.filter(r => r.method === 'POST').length, 1);
});
