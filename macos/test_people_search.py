"""Keyword/person semantics and metadata-only pagination regressions."""
import copy
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent))
import outlook_local as core
import test_outlook_local as fixtures


class PeopleSearchTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.OutlookLocalTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        _, _, self.index = self.fixture.snapshot()
        row = copy.deepcopy(self.index['items'][0])
        row.update(subject='预算方案', preview='讨论张三但不是收件人', sender='owner@example.test', sender_name='Owner',
                   participants={'sender': [{'name': 'Owner', 'address': 'owner@example.test'}],
                                 'to': [{'name': '李四', 'address': 'LI@example.test'}],
                                 'cc': [{'name': '张三', 'address': 'zhang@example.test'}]})
        self.index['items'] = [row]

    def search(self, **values):
        options = dict(query='', field='all', person=None, person_role='any',
                       from_date=None, to_date=None, limit=20, cursor=None)
        options.update(values)
        return core.search(self.index, SimpleNamespace(**options))

    def test_combined_keyword_and_people_roles_are_literal_and_case_insensitive(self):
        result = self.search(query='预算', person='li@EXAMPLE', person_role='to')
        self.assertEqual(result['total_matched'], 1)
        self.assertEqual(result['items'][0]['matched_person_roles'], ['to'])
        self.assertFalse(result['person_search_incomplete'])
        self.assertEqual(self.search(query='不匹配', person='李四')['total_matched'], 0)
        self.assertEqual(self.search(person='张三', person_role='to')['total_matched'], 0)
        self.assertEqual(self.search(person='张三', person_role='cc')['total_matched'], 1)
        self.assertEqual(self.search(person="' OR 1=1")['total_matched'], 0)

    def test_missing_roles_are_unknown_and_never_guessed_from_preview(self):
        row = self.index['items'][0]
        row['participants'] = {'sender': None, 'to': None, 'cc': []}
        result = self.search(person='张三')
        self.assertEqual(result['total_matched'], 0)
        self.assertTrue(result['person_search_incomplete'])
        self.assertFalse(self.search(person='张三', person_role='cc')['person_search_incomplete'])
        # Explicitly unresolved sender must not fall back to a conflicting header.
        self.assertEqual(self.search(person='Owner')['total_matched'], 0)
        row['participants_status'] = {'sender': 'redacted'}
        self.assertEqual(self.search(person='owner', person_role='sender')['total_matched'], 1)
        row['participants_status'] = {'sender': 'conflict'}
        self.assertEqual(self.search(person='owner', person_role='sender')['total_matched'], 0)
        del row['participants']
        self.assertEqual(self.search(person='owner', person_role='sender')['total_matched'], 1)

    def test_only_page_metadata_is_formatted_and_cursor_is_bound_to_person(self):
        row = self.index['items'][0]
        self.index['items'] = [{**row, 'id': str(i)} for i in range(1000)]
        with patch.object(core, 'metadata', wraps=core.metadata) as metadata:
            first = self.search(person='李四', limit=3)
            self.assertEqual(metadata.call_count, 3)
        self.assertEqual(first['total_matched'], 1000)
        second = self.search(person='李四', limit=3, cursor=first['next_cursor'])
        self.assertTrue(set(r['id'] for r in first['items']).isdisjoint(r['id'] for r in second['items']))
        with self.assertRaisesRegex(ValueError, 'another snapshot or query'):
            self.search(person='张三', limit=3, cursor=first['next_cursor'])


if __name__ == '__main__':
    unittest.main()
