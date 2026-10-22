"""Full-text search through the native API, including real synthetic source files."""
import gzip
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent))
import outlook_service as service
import test_outlook_service as fixtures
from mcp_server import validate_tool_arguments, ValidationError


class FullTextServiceTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.ServiceTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.client = self.fixture.client
        self.fixture.ids = [next(row['id'] for row in self.fixture.base['items']
                                 if row['message_id'] == message_id)
                            for message_id in ('<first@example.test>', '<second@example.test>', '<third@example.test>')]
        for identifier, folder in zip(self.fixture.ids, ('inbox', 'deep', 'deleted')):
            self.fixture.scope['messages'][identifier]['folder_id'] = folder
        self.source = self.fixture.fixture.body_source
        self.source.write_text(self.source.read_text().replace(
            '</body>', '<p>' + '常规文字' * 6000 + '尾部独有预算复核</p></body>'))
        self.fixture.refresh()

    def search(self, **kwargs):
        return self.client.call('outlook_local_search', {'query': '预算', 'search_body': True, **kwargs})

    def test_full_body_short_chinese_and_metadata_opt_in_are_read_only(self):
        original = self.fixture.fixture.profile_tree()
        self.assertEqual(self.search(search_body=False)['total_matched'], 0)
        for query in ('算', '预算', '预算复核'):
            result = self.search(query=query)
            self.assertEqual([r['id'] for r in result['items']], [self.fixture.ids[0]])
            self.assertEqual(result['items'][0]['matched_fields'], ['body'])
            self.assertIn(query, result['items'][0]['body_snippet'])
            self.assertFalse(result['body_search_incomplete'])
        with patch.object(service.core, '_resolve_deferred_body', side_effect=AssertionError('unchanged body reread')):
            restarted = service.MailService(self.client.state, self.client.profile)
            result = restarted.call('outlook_local_search', {'query': '预算', 'search_body': True})
        self.assertEqual(result['fulltext']['updated'], 0)
        self.assertEqual(result['fulltext']['reused'], 2)
        self.assertEqual(self.fixture.fixture.profile_tree(), original)

    def test_received_date_folder_and_person_filter_before_body_reads(self):
        identifier = self.fixture.ids[0]
        self.fixture.scope['messages'][identifier].update(
            received_unix=service.core.day_start('2000-01-05'),
            participants={'to': [{'name': '张三', 'address': 'zhang@example.test'}], 'cc': []},
            participants_status={'to': 'verified', 'cc': 'verified'})
        self.fixture.refresh()
        with patch.object(service.core, '_resolve_deferred_body', wraps=service.core._resolve_deferred_body) as reader:
            result = self.search(from_date='2000-01-05', to_date='2000-01-05',
                                 folder_id='inbox', person='张三', person_role='to')
        self.assertEqual(result['fulltext']['candidates'], 1)
        self.assertEqual([r['id'] for r in result['items']], [identifier])
        self.assertEqual(reader.call_count, 1)
        self.assertTrue(result['items'][0]['sent_at'].startswith('2000-01-01'))
        with patch.object(service.core, '_resolve_deferred_body', side_effect=AssertionError('excluded body read')):
            self.assertEqual(self.search(folder_id='inbox', person='别的人')['total_matched'], 0)
            self.assertEqual(self.search(from_date='2000-01-04', to_date='2000-01-04')['total_matched'], 0)

    def test_missing_body_is_retried_after_new_cache_and_refresh(self):
        data = self.source.read_bytes()
        self.source.unlink()
        self.fixture.refresh()
        missing = self.search()
        self.assertEqual(missing['total_matched'], 0)
        self.assertTrue(missing['body_search_incomplete'])
        self.assertEqual(missing['fulltext']['missing'], 1)
        self.source.write_bytes(data)
        self.fixture.refresh()
        found = self.search()
        self.assertEqual(found['total_matched'], 1)
        self.assertFalse(found['body_search_incomplete'])
        self.assertEqual(found['fulltext']['updated'], 1)

    def test_changed_body_and_deleted_folder_do_not_leave_stale_hits(self):
        self.assertEqual(self.search()['total_matched'], 1)
        self.source.write_text(self.source.read_text().replace('预算复核', '结算确认'))
        self.fixture.refresh()
        self.assertEqual(self.search()['total_matched'], 0)
        self.assertEqual(self.search(query='结算确认')['total_matched'], 1)
        self.fixture.scope['messages'][self.fixture.ids[0]]['folder_id'] = 'deleted'
        self.fixture.refresh()
        self.assertEqual(self.search(query='结算确认')['total_matched'], 0)

    def test_pagination_is_bound_to_fulltext_mode_and_searches_subfolders_sent(self):
        source = self.fixture.fixture.gzip_body_source
        html = gzip.decompress(source.read_bytes()).decode().replace('</body>', '<p>预算复核</p></body>')
        source.write_bytes(gzip.compress(html.encode()))
        self.fixture.refresh()
        first = self.search(limit=1)
        self.assertEqual(first['total_matched'], 2)
        second = self.search(limit=1, cursor=first['next_cursor'])
        self.assertNotEqual(first['items'][0]['id'], second['items'][0]['id'])
        self.assertTrue(second['complete'])
        with self.assertRaisesRegex(ValueError, 'another snapshot or query'):
            self.search(search_body=False, limit=1, cursor=first['next_cursor'])
        self.assertEqual(self.search(folder_id='inbox', include_subfolders=False)['total_matched'], 1)
        self.assertEqual(self.search(folder_id='inbox')['total_matched'], 2)
        self.fixture.scope['messages'][self.fixture.ids[1]]['folder_id'] = 'sent'
        self.fixture.refresh()
        self.assertEqual(self.search(folder_id='sent')['total_matched'], 1)

    def test_budget_result_requires_retry_before_pagination(self):
        stats = dict(candidates=2, indexed=0, missing=0, errors=0, pending=2,
                     updated=0, reused=0, complete=False, revision='partial')
        with patch.object(self.client._fulltext_index, 'search', return_value=({}, stats)):
            result = self.search(query='中文', limit=1)
        self.assertFalse(result['complete'])
        self.assertTrue(result['retry_required'])
        self.assertTrue(result['body_search_incomplete'])
        self.assertIsNone(result['next_cursor'])

    def test_mcp_accepts_only_boolean_fulltext_option(self):
        self.assertTrue(validate_tool_arguments('outlook_local_search', {'search_body': True})['search_body'])
        with self.assertRaises(ValidationError):
            validate_tool_arguments('outlook_local_search', {'search_body': 'true'})


if __name__ == '__main__':
    unittest.main()
