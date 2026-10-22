import base64
import gzip
import hashlib
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent))
import common_worker as common
import test_outlook_service as fixtures


class CommonWorkerTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.ServiceTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.fixture.refresh()
        self.reader = common.CommonReader(self.fixture.client)

    def message_id(self, raw_id):
        _, index = self.fixture.client._load()
        return common.encode_id('message', raw_id, index['snapshot'])

    def test_shared_reader_defaults_to_state_separate_from_the_legacy_native_mcp(self):
        with patch.dict(common.os.environ), patch.object(common, 'MailService') as service:
            common.os.environ.pop('OUTLOOK_MAC_STATE_DIR', None)
            common.CommonReader()
        self.assertEqual(service.call_args.kwargs['state'], str(Path.home() / 'Library/Caches/outlook-cli/mac-cache'))

    def test_read_formats_verified_people_and_keeps_partial_structured(self):
        identifier = self.fixture.ids[0]
        membership = self.fixture.scope['messages'][identifier]
        membership.update(participants={'to': [{'name': '中文姓名', 'address': 'person@example.test'}], 'cc': []},
                          participants_status={'to': 'verified', 'cc': 'verified'})
        self.fixture.refresh()
        result = self.reader.perform('read', {'id': self.message_id(identifier)})
        self.assertEqual(result['to'], '中文姓名 <person@example.test>')
        self.assertEqual(result['cc'], '')
        membership['participants_status']['to'] = 'partial'
        self.fixture.refresh()
        result = self.reader.perform('read', {'id': self.message_id(identifier)})
        self.assertIsNone(result['to'])
        self.assertEqual(result['participants']['to'][0]['name'], '中文姓名')

    def test_public_results_do_not_mutate_shared_snapshot(self):
        identifier = self.fixture.ids[0]
        self.fixture.scope['messages'][identifier].update(
            participants={'to': [{'name': '中文姓名', 'address': 'person@example.test'}]},
            participants_status={'to': 'verified'})
        self.fixture.refresh()
        health = self.reader.perform('health', {})
        expected_count = health['cache_summary']['messages']
        health['cache_summary']['messages'] = -1
        self.assertEqual(self.reader.perform('health', {})['cache_summary']['messages'], expected_count)
        with patch.object(common, 'deepcopy', wraps=common.deepcopy) as copied:
            result = self.reader.perform('messages', {'folder': 'inbox', 'limit': 1})
        self.assertEqual(copied.call_count, 1)
        self.assertEqual(len(copied.call_args.args[0]['messages']), 1)
        result['messages'][0]['participants']['to'][0]['name'] = 'changed'
        again = self.reader.perform('messages', {'folder': 'inbox', 'limit': 1})
        self.assertEqual(again['messages'][0]['participants']['to'][0]['name'], '中文姓名')

    def test_common_contract_discovers_stores_and_restricts_folder_tree(self):
        health = self.reader.perform('health', {})
        self.assertEqual(health['backend'], 'mac-cache')
        self.assertEqual(health['date_basis'], 'received_at')
        self.assertTrue(health['capabilities']['received_time'])
        stores = self.reader.perform('stores', {})['stores']
        self.assertEqual(len(stores), 1)
        root_listing = self.reader.perform('folders', {'folder': stores[0]['root_id']})
        self.assertFalse(root_listing['parent']['is_inbox'])
        self.assertFalse(root_listing['parent']['is_sent'])
        roots = root_listing['folders']
        self.assertEqual({row['name'] for row in roots}, {'收件箱', '已发送'})
        self.assertTrue(next(row for row in roots if row['name'] == '收件箱')['is_inbox'])
        self.assertTrue(next(row for row in roots if row['name'] == '已发送')['is_sent'])
        children = self.reader.perform('folders', {'folder': 'inbox'})['folders']
        self.assertEqual(children[0]['name'], '研究')
        self.assertIsNone(children[0]['unread_count'])
        with self.assertRaisesRegex(common.ReadError, 'Inbox'):
            self.reader.perform('messages', {'folder': 'deleted'})

    def test_default_folder_aliases_count_cached_descendants_per_account(self):
        index = {'folders': [
            {'id': 'inbox-a', 'parent_id': 'account-a', 'is_inbox': True,
             'scope_root_id': 'inbox-a', 'cached_messages': 1},
            {'id': 'sent-a', 'parent_id': 'account-a', 'is_sent': True,
             'scope_root_id': 'sent-a', 'cached_messages': 0},
            {'id': 'inbox-b', 'parent_id': 'account-b', 'is_inbox': True,
             'scope_root_id': 'inbox-b', 'cached_messages': 0},
            {'id': 'inbox-b-child', 'parent_id': 'inbox-b',
             'scope_root_id': 'inbox-b', 'cached_messages': 1},
            {'id': 'sent-b', 'parent_id': 'account-b', 'is_sent': True,
             'scope_root_id': 'sent-b', 'cached_messages': 0},
            {'id': 'sent-b-child', 'parent_id': 'sent-b',
             'scope_root_id': 'sent-b', 'cached_messages': 1},
        ]}
        for alias in ('inbox', 'sent'):
            with self.assertRaises(common.ReadError) as error:
                self.reader.folder_key(alias, index)
            self.assertEqual(error.exception.code, 'FOLDER_AMBIGUOUS')
        self.assertEqual(self.reader.folder_key(common.encode_id('folder', 'inbox-b-child', 'a' * 24), index),
                         'inbox-b-child')
        self.assertEqual(self.reader.folder_key(common.encode_id('folder', 'sent-b', 'a' * 24), index),
                         'sent-b')

    def test_received_time_filters_preserve_separate_sent_time(self):
        result = self.reader.perform('messages', {})
        self.assertEqual(len(result['messages']), 1)
        mail = result['messages'][0]
        self.assertIsNotNone(mail['received_at'])
        self.assertNotEqual(mail['received_at'], mail['sent_at'])
        self.assertIsNone(mail['unread'])
        self.assertIsNone(mail['size_bytes'])
        self.assertEqual(mail['folder_path'], '收件箱')
        self.assertEqual(self.reader.perform('messages', {'after': '2000-01-02T00:00:00+08:00'})['messages'], [])
        self.assertEqual(len(self.reader.perform('search', {'query': '中文'})['messages']), 1)
        self.assertEqual(self.reader.perform('search', {'query': "' OR 1=1"})['messages'], [])

    def test_person_search_combines_with_keyword_and_never_reads_bodies(self):
        directory, index = self.fixture.client._current()
        row = next(item for item in index['items'] if item['folder_id'] == 'inbox')
        row['participants'] = {'sender': None, 'to': [{'name': '李四', 'address': 'LI@example.test'}], 'cc': []}
        row['preview'] = '预算讨论'
        with patch.object(self.fixture.client, '_current', return_value=(directory, index)), \
                patch.object(common.core, 'resolve_message_body', side_effect=AssertionError('search must not read bodies')):
            result = self.reader.perform('search', {'query': '预算', 'person': 'li@EXAMPLE', 'person_role': 'to'})
            self.assertEqual(result['messages'][0]['matched_person_roles'], ['to'])
            self.assertFalse(result['person_search_incomplete'])
            self.assertEqual(self.reader.perform('search', {'person': '李四', 'person_role': 'cc'})['messages'], [])
            self.assertEqual(self.reader.perform('search', {'query': '不存在', 'person': '李四'})['messages'], [])
            row['participants']['to'] = None
            missing = self.reader.perform('search', {'person': '李四', 'person_role': 'to'})
            self.assertEqual(missing['messages'], [])
            self.assertTrue(missing['person_search_incomplete'])

    def test_body_search_matches_real_chinese_body_only_text(self):
        body_id = next(row['id'] for row in self.fixture.base['items']
                       if row['message_id'] == '<first@example.test>')
        for identifier, membership in self.fixture.scope['messages'].items():
            if identifier == body_id:
                membership['folder_id'] = 'inbox'
            elif membership['folder_id'] == 'inbox':
                membership['folder_id'] = 'deep'
        self.fixture.scope['fingerprint'] = 'shared-body-only'
        source = self.fixture.fixture.body_source
        source.write_text(source.read_text().replace(
            '</body>', '<p>正文独有预算复核</p></body>'))
        self.fixture.refresh()

        metadata_only = self.reader.perform('search', {
            'folder': 'inbox', 'query': '正文独有预算复核', 'limit': 1})
        self.assertEqual(metadata_only['messages'], [])
        result = self.reader.perform('search', {
            'folder': 'inbox', 'query': '正文独有预算复核', 'search_body': True,
            'limit': 1})
        self.assertEqual(len(result['messages']), 1)
        self.assertIn('正文独有预算复核', result['messages'][0]['body_snippet'])
        self.assertFalse(result['body_search_incomplete'])
        self.assertEqual(result['fulltext']['candidates'], 1)
        self.assertEqual(result['fulltext']['pending'], 0)
        self.assertEqual(result['fulltext']['errors'], 0)
        self.assertTrue(self.reader.perform('health', {})['capabilities']['body_search'])

    def test_body_search_filters_person_before_index_and_retries_pending_window(self):
        directory, index = self.fixture.client._current()
        inbox = next(item for item in index['items'] if item['folder_id'] == 'inbox')
        nested = next(item for item in index['items'] if item['folder_id'] == 'deep')
        inbox['participants'] = {
            'sender': None, 'to': [{'name': '李四', 'address': 'li@example.test'}], 'cc': []}
        inbox['participants_status'] = {'sender': 'redacted', 'to': 'verified', 'cc': 'verified'}
        nested['folder_id'] = 'inbox'
        stats = dict(candidates=1, indexed=0, missing=0, errors=0, pending=1,
                     updated=0, reused=0, complete=False, revision='partial')
        with patch.object(self.fixture.client, '_current', return_value=(directory, index)), \
                patch.object(self.fixture.client._fulltext_index, 'search',
                             return_value=({inbox['id']: '正文片段'}, stats)) as search:
            result = self.reader.perform('search', {
                'folder': 'inbox', 'query': '预算', 'search_body': True,
                'person': '李四', 'person_role': 'to', 'scan_limit': 2, 'limit': 1})
        self.assertEqual([row['id'] for row in search.call_args.args[2]], [inbox['id']])
        self.assertEqual(result['messages'][0]['body_snippet'], '正文片段')
        self.assertEqual(result['next_offset'], 0)
        self.assertTrue(result['retry_required'])
        self.assertTrue(result['body_search_incomplete'])
        self.assertFalse(result['complete'])

    def test_body_search_paginates_each_real_body_once(self):
        body_ids = {row['id'] for row in self.fixture.base['items']
                    if row['message_id'] in {'<first@example.test>', '<second@example.test>'}}
        for identifier, membership in self.fixture.scope['messages'].items():
            membership['folder_id'] = 'inbox' if identifier in body_ids else 'deep'
        self.fixture.scope['fingerprint'] = 'shared-body-pagination'
        first_source = self.fixture.fixture.body_source
        first_source.write_text(first_source.read_text().replace(
            '</body>', '<p>预算分页</p></body>'))
        second_source = self.fixture.fixture.gzip_body_source
        second_html = gzip.decompress(second_source.read_bytes()).decode().replace(
            '</body>', '<p>预算分页</p></body>')
        second_source.write_bytes(gzip.compress(second_html.encode()))
        self.fixture.refresh()

        params = {'folder': 'inbox', 'query': '预算分页', 'search_body': True,
                  'scan_limit': 2, 'limit': 1}
        first = self.reader.perform('search', params)
        self.assertEqual(len(first['messages']), 1)
        self.assertIsNotNone(first['next_offset'])
        self.assertFalse(first['body_search_incomplete'])
        second = self.reader.perform('search', {**params, 'offset': first['next_offset']})
        self.assertEqual(len(second['messages']), 1)
        self.assertNotEqual(first['messages'][0]['id'], second['messages'][0]['id'])
        self.assertIsNone(second['next_offset'])
        self.assertTrue(second['complete'])

    def test_invalid_person_search_is_rejected_before_loading_index(self):
        with patch.object(self.fixture.client, '_current', side_effect=AssertionError('invalid search must fail early')):
            for params in ({}, {'person': ' '}, {'person': 'x', 'person_role': 'bcc'}, {'person_role': 'to'},
                           {'query': '预算', 'search_body': 'true'}):
                with self.assertRaises(common.ReadError):
                    self.reader.perform('search', params)

    def test_late_delivery_listing_defers_body_until_shared_read(self):
        # The source header was sent in 2000; delivery in 2001 must still match.
        raw_id = self.fixture.ids[1]
        self.fixture.scope['messages'][raw_id].update(
            folder_id='inbox', received_unix=common.timestamp('2001-01-02T10:00:00+08:00'))
        client = common.MailService(self.fixture.fixture.state.parent / 'lazy-common-state',
                                    self.fixture.fixture.profile, from_date='2001-01-02')
        reader = common.CommonReader(client)
        with patch.object(common.core, 'freeze_body', wraps=common.core.freeze_body) as freeze:
            result = reader.perform('messages', {
                'after': '2001-01-02T00:00:00+08:00', 'before': '2001-01-03T00:00:00+08:00'})
            self.assertEqual(len(result['messages']), 1)
            mail = result['messages'][0]
            self.assertTrue(mail['sent_at'].startswith('2000-'))
            self.assertTrue(mail['received_at'].startswith('2001-01-02'))
            self.assertEqual(freeze.call_count, 0)
            first = reader.perform('read', {'id': mail['id'], 'max_chars': 2})
            self.assertTrue(first['body']['available'])
            self.assertEqual(first['body']['status'], 'cached_html')
            self.assertEqual(freeze.call_count, 1)
            reader.perform('read', {'id': mail['id'], 'body_offset': first['body']['next_offset']})
            self.assertEqual(freeze.call_count, 1)

    def test_millisecond_received_time_keeps_filter_and_order_boundaries(self):
        directory, index = self.fixture.client._current()
        inbox = next(row for row in index['items'] if row['folder_id'] == 'inbox')
        nested = next(row for row in index['items'] if row['folder_id'] == 'deep')
        inbox['received_unix'] = 946_684_800.500
        nested['folder_id'] = 'inbox'
        nested['received_unix'] = 946_684_800.250
        with patch.object(self.fixture.client, '_current', return_value=(directory, index)):
            result = self.reader.perform('messages', {
                'after': '2000-01-01T00:00:00.250Z',
                'before': '2000-01-01T00:00:00.750Z',
            })
            self.assertEqual([row['id'] for row in result['messages']],
                             [self.message_id(inbox['id']), self.message_id(nested['id'])])
            self.assertEqual(result['messages'][0]['received_at'], '2000-01-01T08:00:00.500000+08:00')
            self.assertEqual(result['messages'][1]['received_at'], '2000-01-01T08:00:00.250000+08:00')
            excluded = self.reader.perform('messages', {
                'after': '2000-01-01T00:00:00Z',
                'before': '2000-01-01T00:00:00.250Z',
            })
        self.assertEqual(excluded['messages'], [])

    def test_received_day_wins_over_sent_day_and_missing_times_are_reported(self):
        directory, index = self.fixture.client._current()
        inbox = next(row for row in index['items'] if row['folder_id'] == 'inbox')
        nested = next(row for row in index['items'] if row['folder_id'] == 'deep')
        inbox['received_unix'] = common.timestamp('2000-01-02T00:00:00+08:00')
        # Sent timestamps remain Jan 1; no fallback for the unknown row.
        nested.update(folder_id='inbox', received_unix=None, received_time_status='missing')
        with patch.object(self.fixture.client, '_current', return_value=(directory, index)):
            filters = {'after': '2000-01-02T00:00:00+08:00', 'before': '2000-01-03T00:00:00+08:00'}
            for method, params in [('messages', filters), ('search', {**filters, 'query': '中文'})]:
                result = self.reader.perform(method, params)
                self.assertEqual(len(result['messages']), 1)
                self.assertEqual(result['messages'][0]['received_at'], '2000-01-02T00:00:00+08:00')
                self.assertEqual(result['received_time_unknown'], 1)
                self.assertTrue(result['date_filter_incomplete'])
            undated = self.reader.perform('messages', {})
            self.assertEqual(len(undated['messages']), 2)
            self.assertIsNone(undated['messages'][-1]['received_at'])
            self.assertFalse(undated['date_filter_incomplete'])
            previous_day = self.reader.perform('messages', {
                'after': '2000-01-01T00:00:00+08:00', 'before': '2000-01-02T00:00:00+08:00'})
            self.assertEqual(previous_day['messages'], [])

    def test_unsupported_filters_and_foreign_ids_fail_explicitly(self):
        for method, params in [('messages', {'unread': False}),
                               ('read', {'id': self.message_id(self.fixture.ids[0]), 'format': 'html'})]:
            with self.assertRaises(common.ReadError):
                self.reader.perform(method, params)
        foreign = base64.urlsafe_b64encode(json.dumps({'v': 2, 'backend': 'local', 'kind': 'message'}).encode()).decode().rstrip('=')
        with self.assertRaises(common.ReadError) as error:
            self.reader.perform('read', {'id': foreign})
        self.assertEqual(error.exception.code, 'BACKEND_ID_MISMATCH')
        with self.assertRaises(common.ReadError) as error:
            self.reader.perform('read', {'id': self.message_id(self.fixture.ids[2])})
        self.assertEqual(error.exception.code, 'SCOPE_DENIED')
        with self.assertRaises(common.ReadError):
            self.reader.perform('delete', {'id': self.message_id(self.fixture.ids[0])})

    def test_utf16_body_pagination_preserves_chinese_and_surrogate_pairs(self):
        source = self.fixture.fixture.gzip_body_source
        original = gzip.decompress(source.read_bytes()).decode()
        source.write_bytes(gzip.compress(original.replace('</body>', '<p>🧪跨平台</p></body>').encode()))
        self.fixture.refresh()
        identifier = self.message_id(self.fixture.ids[1])
        with patch.object(common.core, 'read_message', side_effect=AssertionError('HTML read path used')):
            full = self.reader.perform('read', {'id': identifier})['body']
        self.assertTrue(full['available'])
        self.assertIn('🧪', full['text'])
        offset, chunks = 0, []
        while True:
            body = self.reader.perform('read', {'id': identifier, 'body_offset': offset, 'max_chars': 1})['body']
            self.assertEqual(body['offset'], offset)
            chunks.append(body['text'].encode('utf-16-le', 'surrogatepass'))
            if body['next_offset'] is None:
                break
            self.assertGreater(body['next_offset'], offset)
            offset = body['next_offset']
        self.assertEqual(b''.join(chunks).decode('utf-16-le'), full['text'])
        self.assertEqual(full['total_chars'], len(full['text'].encode('utf-16-le')) // 2)

    def test_uncached_body_remains_null_not_an_empty_success(self):
        result = self.reader.perform('read', {'id': self.message_id(self.fixture.ids[0])})
        self.assertFalse(result['body']['available'])
        self.assertIsNone(result['body']['text'])
        self.assertIsNone(result['body']['total_chars'])
        self.assertIsNotNone(result['received_at'])

    def test_chunked_attachments_match_original_without_source_writes(self):
        row = next(row for row in self.fixture.base['items'] if row['attachments'])
        self.fixture.scope['messages'][row['id']]['folder_id'] = 'sent'
        self.fixture.scope['fingerprint'] = 'sent-attachment'
        self.fixture.refresh()
        identifier = self.message_id(row['id'])
        before = self.fixture.fixture.profile_tree()
        listing = self.reader.perform('attachments', {'id': identifier})['attachments']
        self.assertEqual(listing[0]['index'], 1)
        for invalid in ({'chunk_bytes': 0}, {'byte_offset': 10485761}):
            with patch.object(self.fixture.client, '_attachment', side_effect=AssertionError('export must not start')):
                with self.assertRaises(common.ReadError) as error:
                    self.reader.perform('read_attachment', {'id': identifier, 'index': 1, **invalid})
            self.assertEqual(error.exception.code, 'INVALID_ARGUMENT')
        source_bytes = self.fixture.fixture.attachment_source.read_bytes()
        attachment_hashes = []
        original_sha = common.core.sha

        def count_attachment_hash(data):
            if data == source_bytes:
                attachment_hashes.append(True)
            return original_sha(data)

        offset, chunks, digest = 0, [], None
        with patch.object(common.core, 'read_attachment', wraps=common.core.read_attachment) as export, \
             patch.object(common.core, 'sha', side_effect=count_attachment_hash):
            while True:
                chunk = self.reader.perform('read_attachment', {'id': identifier, 'index': 1, 'byte_offset': offset, 'chunk_bytes': 3})
                chunks.append(base64.b64decode(chunk['data']))
                digest = chunk['sha256']
                if chunk['next_offset'] is None:
                    break
                offset = chunk['next_offset']
        self.assertEqual(export.call_count, 1)
        self.assertEqual(len(attachment_hashes), 1)
        data = b''.join(chunks)
        self.assertEqual(data, source_bytes)
        self.assertEqual(hashlib.sha256(data).hexdigest(), digest)
        self.assertEqual(before, self.fixture.fixture.profile_tree())
        with self.assertRaises(common.ReadError) as error:
            self.reader.perform('read_attachment', {'id': identifier, 'index': 1, 'max_bytes': 1})
        self.assertEqual(error.exception.code, 'ATTACHMENT_TOO_LARGE')

    def test_changed_export_is_rejected_before_chunk_read(self):
        row = next(row for row in self.fixture.base['items'] if row['attachments'])
        self.fixture.scope['messages'][row['id']]['folder_id'] = 'sent'
        self.fixture.scope['fingerprint'] = 'changed-export'
        self.fixture.refresh()
        identifier = self.message_id(row['id'])
        directory, index = self.fixture.client._current()
        attachment = row['attachments'][0]
        cached = self.fixture.client._attachment(directory, index, {'id': row['id'], 'attachment_id': attachment['id']})
        output = Path(cached['output'])
        output.chmod(0o600)
        output.write_bytes(b'changed-export')
        with patch.object(self.fixture.client, '_attachment', return_value=cached):
            with self.assertRaises(common.ReadError) as error:
                self.reader.perform('read_attachment', {'id': identifier, 'index': 1, 'chunk_bytes': 3})
        self.assertEqual(error.exception.code, 'ATTACHMENT_CHANGED')

    def test_old_message_id_cannot_bypass_new_deleted_membership(self):
        identifier = self.message_id(self.fixture.ids[1])
        self.fixture.scope['messages'][self.fixture.ids[1]]['folder_id'] = 'deleted'
        self.fixture.scope['fingerprint'] = 'moved-to-deleted'
        self.fixture.refresh()
        with self.assertRaisesRegex(ValueError, 'folder access changed'):
            self.reader.perform('read', {'id': identifier})


if __name__ == '__main__':
    unittest.main()
