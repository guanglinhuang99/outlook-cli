import assert from 'node:assert/strict';
import test from 'node:test';
import { compactMessage, compactResult } from '../src/compact.mjs';

const macMessage = {
  id: 'mac-message-id', internet_message_id: '<third@example.test>', related_internet_message_ids: [],
  subject: '中文主题丙', subject_status: 'paired',
  sender: { name: '赵六', address: 'third@example.com', address_type: 'SMTP', smtp_address: 'third@example.com' },
  received_at: '2000-01-01T08:01:30+08:00', sent_at: '2000-01-01T08:00:30+08:00',
  received_time_status: 'verified', received_time_method: 'osa_last_delivery_time', unread: null,
  unread_evidence: { observed_at: '2000-01-01T00:00:00Z', source: 'osa' },
  participants: { to: [{ name: '甲', address: 'a@example.com' }] }, participants_status: { to: 'verified' },
  participants_evidence: { to: { source: 'hx' } }, size_bytes: null, download_state: null, preview: '请审批',
  body_status: 'deferred', attachment_count: 0, attachments_complete: false,
  folder_id: 'mac-folder-id', folder_path: '收件箱', folder_status: 'verified_latest_observation'
};

test('compact mail lists keep what an agent needs and drop diagnostics', () => {
  const full = {
    backend: 'mac-cache', data_scope: 'verified_inbox_sent_cache', date_basis: 'received_at',
    snapshot: 'abc', captured_at: 'now', coverage: 'Locally recoverable cache only; long description',
    messages: [macMessage], scanned: 1, filtered_item_count: 1, next_offset: null, complete: true,
    scope: 'single_folder', pagination: 'live_offset_not_snapshot', body_search_incomplete: false,
    person_search_incomplete: true, received_time_unknown: 0, unread_unknown_count: 1, date_filter_incomplete: false
  };
  const compact = compactResult('messages', full);
  assert.deepEqual(compact, {
    backend: 'mac-cache', data_scope: 'verified_inbox_sent_cache', next_offset: null, complete: true,
    person_search_incomplete: true, unread_unknown_count: 1,
    messages: [{ id: 'mac-message-id', subject: '中文主题丙', sender: { name: '赵六', address: 'third@example.com' },
      received_at: '2000-01-01T08:01:30+08:00', unread: null, folder: '收件箱', preview: '请审批' }]
  });
  assert.ok(JSON.stringify(compact).length < JSON.stringify(full).length / 3);
});

test('abnormal states stay visible in the compact view', () => {
  const message = compactMessage({ ...macMessage, received_time_status: 'conflict', folder_status: 'unresolved',
    subject_status: 'missing', download_state: 0, attachment_count: 2,
    attachments: [{ index: 1, filename: 'a.pdf', size_bytes: 10, type: 1, status: 'not_cached' }] });
  assert.equal(message.received_time_status, 'conflict');
  assert.equal(message.folder_status, 'unresolved');
  assert.equal(message.subject_status, 'missing');
  assert.equal(message.header_only, true);
  assert.deepEqual(message.attachments, [{ index: 1, filename: 'a.pdf', size_bytes: 10, status: 'not_cached' }]);
});

test('read keeps the body page, recipients and paging; bundles keep errors and cursors', () => {
  const read = compactResult('read', {
    ...macMessage, backend: 'com', data_scope: 'outlook_session', to: '甲 <a@example.com>', cc: '',
    body: { format: 'text', text: '正文', total_chars: 9, offset: 0, next_offset: 2, truncated: true, available: true, status: 'cached' }
  });
  assert.deepEqual(read.body, { text: '正文', total_chars: 9, next_offset: 2, truncated: true });
  assert.equal(read.to, '甲 <a@example.com>');
  assert.equal('cc' in read, false);
  const missing = compactMessage({ id: 'x', body: { text: null, available: false, status: 'missing', diagnosis: 'not cached' } });
  assert.deepEqual(missing.body, { text: null, total_chars: undefined, available: false, status: 'missing', diagnosis: 'not cached' });
  const bundle = compactResult('context_bundle', { items: [{ ...macMessage, folder: { id: 'f', name: '项目', path: '\\\\收件箱\\项目' } }],
    errors: [{ id: 'y', code: 'TEXT_BUDGET_EXHAUSTED' }], next_cursor: 'c', complete: false, mode: 'context',
    folder: 'all', body_complete: false, retry_required: true, fulltext: { pending: 3, complete: false } });
  assert.equal(bundle.items[0].folder, '\\\\收件箱\\项目');
  assert.deepEqual(bundle.errors, [{ id: 'y', code: 'TEXT_BUDGET_EXHAUSTED' }]);
  assert.equal(bundle.next_cursor, 'c');
  assert.equal(bundle.body_complete, false);
  assert.equal(bundle.retry_required, true);
  assert.deepEqual(bundle.fulltext, { pending: 3, complete: false });
});
