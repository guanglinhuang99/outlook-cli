// Read-only real-cache verification. No Outlook close/launch and no mail content output.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { Client, StreamableHTTPClientTransport } from '@modelcontextprotocol/client';
import { call, readState } from '../src/runtime.mjs';
import { BackendRouter } from '../src/backend.mjs';
import { LocalFileBridge, localError } from '../src/offline.mjs';

const health = await call('health', { backend: 'local' });
assert.equal(health.backend, 'local');
const report = { backend: health.backend, readable_stores: health.readable_store_count, folders: health.stores.filter(s => s.readable).map(s => s.folders) };
const stores = await call('stores', { backend: 'local' });
const root = stores.stores.find(s => s.readable).root_id;
const folders = await call('folders', { backend: 'local', folder: root });
assert.ok(folders.folders.length > 0);
const inboxBefore = await call('folders', { backend: 'local', folder: 'inbox', limit: 1 });
const list = await call('messages', { backend: 'local', limit: 5 });
assert.ok(list.messages.length > 0);
report.inbox_items = inboxBefore.parent.item_count;
const mail = list.messages[0];
// Backend remains local for a local ID even if Outlook is currently running.
const first = await call('read', { id: mail.id, max_chars: 100 });
assert.equal(first.backend, 'local');
assert.ok(first.body.available);
const second = await call('read', { id: mail.id, body_offset: 100, max_chars: 100 });
const both = await call('read', { id: mail.id, max_chars: 200 });
assert.ok(first.body.text + second.body.text === both.body.text, 'Local body pagination must be stable');
report.body_pagination = true;
report.body_derived_from = first.body.derived_from;
const matches = await call('search', { backend: 'local', query: mail.subject, limit: 1, scan_limit: 5 });
assert.equal(matches.messages[0].id, mail.id);
const unread = await call('messages', { backend: 'local', unread: true, limit: 3 });
assert.ok(unread.messages.every(m => m.unread));
const range = await call('messages', { backend: 'local', after: mail.received_at, limit: 1 });
assert.equal(range.messages[0].id, mail.id);
report.search_and_filters = true;
const attached = list.messages.find(m => m.has_attachments);
if (attached) {
  const attachments = await call('attachments', { id: attached.id });
  const candidate = attachments.attachments.find(a => a.type === 1 && a.size_bytes < 10485760);
  if (candidate) {
    const result = await call('read_attachment', { id: attached.id, index: candidate.index, chunk_bytes: 1048576 });
    assert.equal(result.next_offset, null);
    const bytes = Buffer.from(result.data, 'base64');
    assert.equal(createHash('sha256').update(bytes).digest('hex'), result.sha256);
    report.attachment_bytes_verified = bytes.length;
    const com = await call('messages', { backend: 'com', limit: 5 }).catch(() => null);
    const same = com?.messages.find(m => m.subject === attached.subject);
    if (same) {
      const comAttachments = await call('attachments', { id: same.id, backend: 'com' });
      const matching = comAttachments.attachments.find(a => a.filename === candidate.filename);
      if (matching) {
        const comData = await call('read_attachment', { id: same.id, backend: 'com', index: matching.index, chunk_bytes: 1048576 });
        assert.equal(comData.sha256, result.sha256);
        report.attachment_matches_com = true;
      }
    }
  }
}
// Exercise automatic transition with a simulated unavailable COM and the REAL reader.
const router = new BackendRouter({ com: { queue: [], active: false, close() {}, async call() { throw localError('OUTLOOK_UNAVAILABLE', 'simulated closed session'); } }, local: new LocalFileBridge() });
try {
  const automatic = await router.call('health');
  assert.equal(automatic.backend, 'local');
  assert.equal(automatic.fallback_reason, 'OUTLOOK_UNAVAILABLE');
  report.auto_fallback_with_simulated_com_unavailable = true;
} finally { router.close(); }
const state = await readState();
const client = new Client({ name: 'offline-smoke', version: '1.0.0' });
try {
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${state.port}/mcp`), { requestInit: { headers: { Authorization: `Bearer ${state.token}` } } }));
  const result = await client.callTool({ name: 'outlook_messages', arguments: { backend: 'local', limit: 1 } });
  assert.equal(result.structuredContent.backend, 'local');
  assert.equal(result.structuredContent.messages.length, 1);
  report.mcp_local_backend = true;
} finally { await client.close(); }
const inboxAfter = await call('folders', { backend: 'local', folder: 'inbox', limit: 1 });
report.unread_before = inboxBefore.parent.unread_count;
report.unread_after = inboxAfter.parent.unread_count;
console.log(JSON.stringify(report, null, 2));
