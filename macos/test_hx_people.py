"""Person versions, exact mail ownership and incomplete-search behavior."""
import copy
import struct
import unittest
from pathlib import Path
from types import SimpleNamespace
import sys

sys.path.insert(0, str(Path(__file__).parent))
import outlook_local as core
import test_outlook_local as fixtures
from test_hxrecords import person_frame, serialized_frame


class HxPeopleTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.OutlookLocalTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)

    def snapshot_with(self, people, *, targets=(0xABC,), second_owner=False):
        original = list(core.iter_frames(self.fixture.block))
        header = next(f for f in original if f.kind == 201 and f.key == 0xABC)
        refs = b''.join(struct.pack('<H', 202) + b'\0' * 8 + struct.pack('<Q', k) for k in targets)
        appended = serialized_frame(201, header.key, header.data[38:] + refs)
        if second_owner:
            other = next(f for f in original if f.kind == 201 and f.key == 0xABD)
            appended += serialized_frame(201, other.key, other.data[38:] + refs)
        self.fixture.block.data += appended + b''.join(people)
        _, _, index = self.fixture.snapshot()
        return next(r for r in index['items'] if r['key'] == 0xABC)

    def test_exact_parent_roles_survive_external_body_precedence(self):
        row = self.snapshot_with([
            person_frame(11, owner=0xABC, role='to', name='张三🧪', address='zhang@example.test'),
            person_frame(12, owner=0xABC, role='cc', name='李四', address='li@example.test'),
            person_frame(13, owner=0xABD, role='to', name='其他人', address='other@example.test'),
        ])
        self.assertIsNone(row['body'].get('body_reference'))
        merged = core.merge_participants(row, self.membership(to=1, cc=1))
        self.assertEqual(merged['participants']['to'], [{'name': '张三🧪', 'address': 'zhang@example.test'}])
        self.assertEqual(merged['participants']['cc'][0]['name'], '李四')
        self.assertEqual(merged['participants_status']['to'], 'verified')

    @staticmethod
    def membership(to=1, cc=0):
        return {'participants': {'to': None, 'cc': None},
                'participants_status': {'to': 'redacted', 'cc': 'redacted'},
                'participant_counts': {role: {'count': n, 'status': 'verified'} for role, n in [('to', to), ('cc', cc)]}}

    def test_ambiguous_header_target_is_not_attached(self):
        row = self.snapshot_with([person_frame(11, owner=0xABC)], targets=(0xABC, 0xABD))
        self.assertNotIn('hx_participants', row)

    def test_shared_message_data_target_is_not_attached(self):
        row = self.snapshot_with([person_frame(11, owner=0xABC)], second_owner=True)
        self.assertNotIn('hx_participants', row)

    def test_matching_versions_deduplicate_but_distinct_people_do_not(self):
        first = person_frame(11, owner=0xABC)
        row = self.snapshot_with([first, first, person_frame(12, owner=0xABC)])
        self.assertEqual(len(row['hx_participants']['to']['contacts']), 2)
        self.assertEqual(core.merge_participants(row, self.membership(to=2))['participants_status']['to'], 'verified')

    def test_conflicting_person_versions_fail_closed(self):
        row = self.snapshot_with([person_frame(11, owner=0xABC, address='a@example.test'),
                                  person_frame(11, owner=0xABC, address='b@example.test')])
        merged = core.merge_participants(row, self.membership())
        self.assertIsNone(merged['participants']['to'])
        self.assertEqual(merged['participants_status']['to'], 'conflict')

    def test_parent_change_poisoning_and_invalid_version_are_not_ignored(self):
        records = {}
        raws = [person_frame(11, owner=0xABC), person_frame(11, owner=0xABD)]
        for raw in raws:
            core._record_person(records, next(core.iter_frames(SimpleNamespace(offset=0, data=raw))))
        groups = core._person_collections(records)
        self.assertTrue(groups[0xABC]['to']['conflict'])
        self.assertTrue(groups[0xABD]['to']['conflict'])
        broken = bytearray(person_frame(12, owner=0xABC))
        struct.pack_into('<I', broken, 310, 0xFFFFFF)
        row = self.snapshot_with([person_frame(12, owner=0xABC), bytes(broken), person_frame(13, owner=0xABC)])
        merged = core.merge_participants(row, self.membership(to=2))
        self.assertEqual(len(merged['participants']['to']), 1)
        self.assertEqual(merged['participants_status']['to'], 'partial')
        self.assertEqual(merged['participants_evidence']['to']['rejected_count'], 1)

    def test_count_mismatch_missing_and_conflict_stay_incomplete(self):
        row = self.snapshot_with([person_frame(11, owner=0xABC, name='张三', address='zhang@example.test')])
        for membership in (self.membership(to=2), {'participants_status': {'to': 'partial'}}):
            merged = core.merge_participants(row, membership)
            self.assertEqual(merged['participants_status']['to'], 'partial')
            roles, incomplete = core.person_matches({**row, **merged}, '张三', 'to')
            self.assertEqual(roles, ['to'])
            self.assertTrue(incomplete)
        membership = self.membership()
        membership['participant_counts']['to']['status'] = 'conflict'
        merged = core.merge_participants(row, membership)
        self.assertEqual(merged['participants_status']['to'], 'conflict')
        self.assertIsNone(merged['participants']['to'])

    def test_verified_osa_and_explicit_conflict_are_preserved(self):
        row = self.snapshot_with([person_frame(11, owner=0xABC)])
        for status, contacts in [('verified', []), ('conflict', None)]:
            membership = {'participants': {'to': contacts}, 'participants_status': {'to': status}}
            before = copy.deepcopy(membership)
            merged = core.merge_participants(row, membership)
            self.assertEqual(merged['participants']['to'], contacts)
            self.assertEqual(merged['participants_status']['to'], status)
            self.assertEqual(membership, before)

    def test_all_rejected_people_report_partial_not_redacted(self):
        broken = bytearray(person_frame(11, owner=0xABC))
        struct.pack_into('<I', broken, 310, 0xFFFFFF)
        row = self.snapshot_with([bytes(broken)])
        merged = core.merge_participants(row, self.membership())
        self.assertIsNone(merged['participants']['to'])
        self.assertEqual(merged['participants_status']['to'], 'partial')
        self.assertEqual(merged['participants_evidence']['to']['rejected_count'], 1)


if __name__ == '__main__':
    unittest.main()
