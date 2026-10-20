// Read-only real-cache smoke check. Emit counts/booleans, never mail content.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { call, readState, root, runtimeDir } from '../src/runtime.mjs';

let stage = 'platform';
async function run() {
  assert.equal(process.platform, 'darwin');
  const report = {};
  stage = 'health';
  const health = await call('health');
  assert.equal(health.backend, 'mac-cache');
  assert.equal(health.date_basis, 'sent_at');
  report.health = true;
  stage = 'folder_tree';
  const stores = await call('stores');
  assert.ok(stores.stores.length > 0);
  const queue = stores.stores.map(store => ({ id: store.root_id, role: null }));
  const seen = new Set();
  let folderCount = 0;
  for (let i = 0; i < queue.length; i++) {
    const { id, role } = queue[i];
    assert.ok(!seen.has(id), 'Folder tree must not contain cycles');
    seen.add(id);
    let offset = 0;
    do {
      const page = await call('folders', { folder: id, limit: 10, offset });
      queue.push(...page.folders.map(folder => ({ ...folder,
        role: folder.is_inbox ? 'inbox' : folder.is_sent ? 'sent' : role })));
      folderCount += page.folders.length;
      offset = page.next_offset;
    } while (offset !== null);
  }
  report.stores = stores.stores.length;
  report.folders = folderCount;
  const inbox = queue.find(folder => folder.role === 'inbox' && folder.item_count > 0);
  const sent = queue.find(folder => folder.role === 'sent' && folder.item_count > 0)
    ?? queue.find(folder => folder.role === 'sent');
  assert.ok(inbox, 'An Inbox folder must have a cached test message');
  const selection = { folder: inbox.id };
  const list = await call('messages', { ...selection, limit: 200 });
  assert.ok(list.messages.length > 0, 'Inbox must have a cached test message');
  assert.ok(list.messages.every(row => row.received_at === null && row.unread === null));
  const sample = list.messages[0];
  stage = 'today';
  for (const [label, folder] of [['inbox', inbox], ['sent', sent]]) {
    if (!folder) { report[`today_${label}_sample`] = null; continue; }
    const today = await call('today', { folder: folder.id, limit: 5 });
    assert.equal(today.date_basis, 'sent_at');
    assert.equal(today.timezone, 'Asia/Singapore');
    const start = Date.parse(`${today.date}T00:00:00+08:00`);
    assert.ok(today.messages.every(row => Date.parse(row.sent_at) >= start && Date.parse(row.sent_at) < start + 86400000));
    const repeated = await call('today', { folder: folder.id, date: today.date, limit: 5 });
    assert.deepEqual(repeated.messages.map(row => row.id), today.messages.map(row => row.id));
    report[`today_${label}_sample`] = today.messages.length;
  }
  stage = 'search_and_dates';
  const ranged = await call('messages', { ...selection, after: sample.sent_at, limit: 1 });
  assert.equal(ranged.messages[0]?.id, sample.id);
  const excluded = await call('messages', { ...selection, before: sample.sent_at, limit: 5 });
  assert.ok(excluded.messages.every(row => Date.parse(row.sent_at) < Date.parse(sample.sent_at)));
  if (sample.subject) {
    const match = await call('search', { ...selection, query: sample.subject.slice(0, 500), limit: 1 });
    assert.equal(match.messages[0]?.id, sample.id);
  }
  const missing = await call('search', { ...selection, query: '__outlook_probe_no_match_6aa92e70__', scan_limit: 2 });
  assert.equal(missing.messages.length, 0);
  assert.equal(missing.scanned, Math.min(2, list.filtered_item_count));
  report.search_and_dates = true;
  stage = 'body';
  let available = false;
  let absent = false;
  for (const row of list.messages.slice(0, 30)) {
    const first = await call('read', { id: row.id, max_chars: 100 });
    if (first.body.available) {
      const second = await call('read', { id: row.id, body_offset: 100, max_chars: 100 });
      const both = await call('read', { id: row.id, max_chars: 200 });
      assert.equal(first.body.text + second.body.text, both.body.text);
      available = true;
    } else {
      assert.equal(first.body.text, null);
      assert.equal(first.body.total_chars, null);
      assert.ok(first.body.status);
      absent = true;
    }
    if (available && absent) break;
  }
  report.body_pagination = available;
  report.missing_body_reported = absent;
  stage = 'scope_and_capabilities';
  await assert.rejects(call('messages', { folder: 'deleted' }), { code: 'SCOPE_DENIED' });
  await assert.rejects(call('messages', { unread: false }), { code: 'UNSUPPORTED_FILTER' });
  await assert.rejects(call('search', { query: 'test', search_body: true }), { code: 'UNSUPPORTED_FILTER' });
  await assert.rejects(call('read', { id: sample.id, format: 'html' }), { code: 'UNSUPPORTED_FORMAT' });
  report.scope_and_capabilities = true;
  stage = 'attachments';
  report.attachment_verified = false;
  for (const row of list.messages.filter(row => row.attachment_count > 0)) {
    const metadata = await call('attachments', { id: row.id });
    const candidate = metadata.attachments.find(a => a.type === 1 && a.size_bytes <= 10485760);
    if (!candidate) continue;
    const chunks = [];
    let offset = 0;
    let hash;
    let size;
    do {
      const part = await call('read_attachment', { id: row.id, index: candidate.index, byte_offset: offset, chunk_bytes: 1048576 });
      hash ??= part.sha256;
      size ??= part.size_bytes;
      assert.equal(part.sha256, hash);
      assert.equal(part.size_bytes, size);
      assert.equal(part.offset, offset);
      assert.ok(part.next_offset === null || part.next_offset > offset);
      chunks.push(Buffer.from(part.data, 'base64'));
      offset = part.next_offset;
    } while (offset !== null);
    const bytes = Buffer.concat(chunks);
    assert.equal(bytes.length, size);
    assert.equal(createHash('sha256').update(bytes).digest('hex'), hash);
    const dir = await mkdtemp(path.join(runtimeDir, 'smoke-download-'));
    try {
      const destination = path.join(dir, 'attachment.bin');
      const args = [path.join(root, 'bin/outlook.mjs'), 'download', '--id', row.id, '--index', String(candidate.index), '--out', destination];
      await promisify(execFile)(process.execPath, args);
      assert.ok((await readFile(destination)).equals(bytes));
      await assert.rejects(promisify(execFile)(process.execPath, args));
      report.attachment_verified = true;
      report.cli_download_no_overwrite = true;
    } finally { await rm(dir, { recursive: true, force: true }); }
    break;
  }
  stage = 'mcp';
  const state = await readState();
  for (const type of ['http', 'stdio']) {
    const client = new Client({ name: 'outlook-mac-smoke', version: '1.0.0' });
    const transport = type === 'http'
      ? new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${state.port}/mcp`), { requestInit: { headers: { Authorization: `Bearer ${state.token}` } } })
      : new StdioClientTransport({ command: process.execPath, args: [path.join(root, 'bin/outlook.mjs'), 'mcp'], env: { ...process.env, OUTLOOK_LOCAL_STATE_DIR: runtimeDir }, stderr: 'pipe' });
    try {
      await client.connect(transport);
      const tools = await client.listTools();
      assert.equal(tools.tools.length, 9);
      const mail = await client.callTool({ name: 'outlook_today', arguments: { ...selection, limit: 1 } });
      assert.equal(mail.isError, undefined);
      assert.equal(mail.structuredContent.date_basis, 'sent_at');
      const body = await client.callTool({ name: 'outlook_read', arguments: { id: sample.id, max_chars: 100 } });
      assert.equal(body.isError, undefined);
      assert.ok(Object.hasOwn(body.structuredContent.body, 'available'));
      report[`mcp_${type}`] = true;
    } finally { await client.close(); }
  }
  console.log(JSON.stringify(report, null, 2));
}

run().catch(() => {
  // SDK/assertion diagnostics can include mail values. Keep failure output private.
  console.error(JSON.stringify({ error: 'MAC_SMOKE_FAILED', stage }));
  process.exitCode = 1;
});
