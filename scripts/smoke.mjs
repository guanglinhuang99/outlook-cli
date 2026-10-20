// Read-only real-mailbox integration test. Emits only counters/booleans, no mail content.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, readFile, unlink, rmdir } from 'node:fs/promises';
import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { call, readState, root, runtimeDir } from '../src/runtime.mjs';
import path from 'node:path';

const report = {};
report.health = await call('health');
const stores = await call('stores');
assert.ok(stores.stores.length > 0);
report.store_count = stores.stores.length;
const folders = await call('folders', { folder: stores.stores.at(-1).root_id, limit: 2 });
assert.equal(folders.folders.length, Math.min(2, folders.total));
report.folder_pagination = folders.next_offset !== null;
const before = await call('folders', { folder: 'inbox', limit: 1 });
const list = await call('messages', { limit: 5 });
assert.ok(list.messages.length);
const sample = list.messages[0];
const read = await call('read', { id: sample.id, max_chars: 100 });
assert.ok(read.body.text.length <= 100);
assert.equal(read.unread, sample.unread);
if (read.body.next_offset !== null) {
  const second = await call('read', { id: sample.id, max_chars: 100, body_offset: read.body.next_offset });
  const both = await call('read', { id: sample.id, max_chars: 200 });
  assert.ok(read.body.text + second.body.text === both.body.text, 'Body segments must reconstruct the larger segment');
}
report.body_pagination = true;
report.smtp_resolved = Boolean(read.sender.smtp_address);
const match = await call('search', { query: sample.subject, limit: 1, scan_limit: 5 });
assert.equal(match.messages[0].id, sample.id);
const missing = await call('search', { query: '__outlook_probe_no_match_6aa92e70__', scan_limit: 2 });
assert.equal(missing.messages.length, 0);
assert.equal(missing.scanned, 2);
assert.equal(missing.next_offset, 2);
report.search_and_scan_limit = true;
const unread = await call('messages', { unread: true, limit: 3 });
assert.ok(unread.messages.every(m => m.unread));
const range = await call('messages', { after: '2000-01-01T00:00:00Z', before: '2100-01-01T00:00:00Z', limit: 1 });
assert.equal(range.messages[0].id, sample.id);
const exact = await call('messages', { after: sample.received_at, limit: 1 });
assert.equal(exact.messages[0].id, sample.id);
const excluded = await call('messages', { before: sample.received_at, limit: 1 });
assert.ok(excluded.messages.every(m => Date.parse(m.received_at) < Date.parse(sample.received_at)));
report.filters = true;
const attached = list.messages.find(m => m.attachment_count > 0);
if (attached) {
  const metadata = await call('attachments', { id: attached.id });
  const candidate = metadata.attachments.find(a => a.type === 1 && a.size_bytes <= 10485760);
  if (candidate) {
    const data = await call('read_attachment', { id: attached.id, index: candidate.index });
    const bytes = Buffer.from(data.data, 'base64');
    assert.equal(bytes.length, data.size_bytes);
    assert.equal(createHash('sha256').update(bytes).digest('hex'), data.sha256);
    report.attachment_bytes_verified = bytes.length;
    const firstChunk = await call('read_attachment', { id: attached.id, index: candidate.index, chunk_bytes: 100 });
    assert.equal(firstChunk.chunk_size, Math.min(100, bytes.length));
    if (firstChunk.next_offset !== null) {
      const nextChunk = await call('read_attachment', { id: attached.id, index: candidate.index, byte_offset: firstChunk.next_offset, chunk_bytes: 100 });
      assert.ok(Buffer.concat([Buffer.from(firstChunk.data, 'base64'), Buffer.from(nextChunk.data, 'base64')]).equals(bytes.subarray(0, 200)));
    }
    report.attachment_pagination = true;
    const dir = await mkdtemp(path.join(runtimeDir, 'smoke-download-'));
    const destination = path.join(dir, 'attachment.bin');
    try {
      const args = [path.join(root, 'bin/outlook.mjs'), 'download', '--id', attached.id, '--index', String(candidate.index), '--out', destination];
      await promisify(execFile)(process.execPath, args, { windowsHide: true });
      assert.ok((await readFile(destination)).equals(bytes));
      await assert.rejects(promisify(execFile)(process.execPath, args, { windowsHide: true }));
      report.cli_download_and_no_overwrite = true;
    } finally {
      await unlink(destination).catch(() => {});
      await rmdir(dir);
    }
  }
}
report.concurrent_calls = (await Promise.all([call('health'), call('health'), call('health')])).length;
const state = await readState();
for (const type of ['http', 'stdio']) {
  const client = new Client({ name: 'outlook-smoke', version: '1.0.0' });
  const transport = type === 'http'
    ? new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${state.port}/mcp`), { requestInit: { headers: { Authorization: `Bearer ${state.token}` } } })
    : new StdioClientTransport({ command: process.execPath, args: [path.join(root, 'bin/outlook.mjs'), 'mcp'], env: { ...process.env, OUTLOOK_LOCAL_STATE_DIR: runtimeDir }, stderr: 'pipe' });
  try {
    await client.connect(transport);
    const tools = await client.listTools();
    assert.equal(tools.tools.length, 9);
    const result = await client.callTool({ name: 'outlook_health', arguments: {} });
    assert.equal(result.structuredContent.connected, true);
    const mail = await client.callTool({ name: 'outlook_messages', arguments: { limit: 1 } });
    assert.equal(mail.structuredContent.messages.length, 1);
    report[`mcp_${type}`] = true;
  } finally { await client.close(); }
}
const after = await call('folders', { folder: 'inbox', limit: 1 });
report.unread_count_before = before.parent.unread_count;
report.unread_count_after = after.parent.unread_count;
// Count can legitimately change if mail arrives or the user reads a message during this test.
report.sample_unread_unchanged = (await call('read', { id: sample.id, max_chars: 1 })).unread === sample.unread;
assert.ok(report.sample_unread_unchanged);
console.log(JSON.stringify(report, null, 2));
