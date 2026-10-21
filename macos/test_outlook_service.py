import copy
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import gzip
import hashlib
import json
import os
import subprocess
import shutil
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent))
import outlook_service as service
import test_outlook_local as fixtures


class FrozenDatetime(datetime):
    instant = datetime(1999, 12, 31, 16, 30, tzinfo=timezone.utc)

    @classmethod
    def now(cls, tz=None):
        return cls.instant.astimezone(tz)


class ServiceTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.OutlookLocalTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        _, _, self.base = self.fixture.snapshot()
        self.ids = [r['id'] for r in self.base['items']]
        self.scope = {
            'folders': [
                dict(id='inbox', parent_id='root', name='收件箱', is_inbox=True),
                dict(id='child', parent_id='inbox', name='研究', is_inbox=False),
                dict(id='deep', parent_id='child', name='二级', is_inbox=False),
                dict(id='sent', parent_id='root', name='已发送', is_inbox=False, is_sent=True),
                dict(id='deleted', parent_id='root', name='已删除', is_inbox=False, is_deleted=True),
                dict(id='trashchild', parent_id='deleted', name='旧邮件', is_inbox=False),
            ],
            'messages': {
                self.ids[0]: dict(folder_id='inbox', folder_status='verified_latest_observation'),
                self.ids[1]: dict(folder_id='deep', folder_status='verified_latest_observation'),
                self.ids[2]: dict(folder_id='deleted', folder_status='verified_latest_observation'),
            },
            'fingerprint': 'first-catalog', 'evidence': {'source': 'synthetic'},
        }
        for row in self.base['items']:
            self.scope['messages'][row['id']].update(
                received_unix=row['sent_unix'] + 60,
                received_time_status='verified', received_time_method='osa_last_delivery_time')
        module = types.ModuleType('osa_scope')
        module.read_scope = lambda profile, items: copy.deepcopy(self.scope)
        self.addCleanup(patch.stopall)
        patch.dict(sys.modules, {'osa_scope': module}).start()
        self.client = service.MailService(self.fixture.state, self.fixture.profile)

    def refresh(self):
        return self.client.call('outlook_local_refresh', {})

    def test_receipt_scope_does_not_drop_mail_sent_before_cutoff(self):
        # These headers were sent in 2000 but delivered in 2001.
        received = datetime(2001, 1, 1, tzinfo=timezone.utc).timestamp()
        for identifier in self.ids[:2]:
            self.scope['messages'][identifier]['received_unix'] = received
        result = self.client.call('outlook_local_refresh', {'from_date': '2001-01-01'})
        self.assertEqual(result['messages'], 2)
        page = self.client.call('outlook_local_search', {'from_date': '2001-01-01', 'to_date': '2001-01-01'})
        self.assertEqual(page['total_matched'], 2)
        self.assertTrue(all(row['sent_at'].startswith('2000-') for row in page['items']))
        self.assertTrue(all(row['received_at'].startswith('2001-01-01') for row in page['items']))
        self.assertEqual(page['date_basis'], 'received_at')

    def test_delayed_received_mail_can_read_body_after_received_filter(self):
        received = datetime(2001, 1, 1, tzinfo=timezone.utc).timestamp()
        for identifier in self.ids[:2]:
            self.scope['messages'][identifier]['received_unix'] = received
        self.client.call('outlook_local_refresh', {'from_date': '2001-01-01'})
        page = self.client.call('outlook_local_search', {
            'from_date': '2001-01-01', 'to_date': '2001-01-01'
        })
        message = self.client.call('outlook_local_read', {'id': page['items'][0]['id']})
        self.assertEqual(message['body']['status'], 'cached_html')
        self.assertTrue(message['body'].get('text'))

    def test_refresh_defers_body_freeze_and_reuses_resolved_sidecar(self):
        original = self.fixture.profile_tree()
        with patch.object(service.core, 'freeze_body', wraps=service.core.freeze_body) as freeze:
            result = self.refresh()
            self.assertEqual(result['deferred_bodies'], result['messages'])
            self.assertEqual(freeze.call_count, 0)
            listing = self.client.call('outlook_local_search', {})
            self.assertEqual(freeze.call_count, 0)
            target = next(item for item in listing['items'] if item['subject'] == '中文主题乙')
            message = self.client.call('outlook_local_read', {'id': target['id']})
            self.assertEqual(message['body']['status'], 'cached_html')
            self.assertEqual(freeze.call_count, 1)
            self.client.call('outlook_local_read', {'id': target['id']})
            self.assertEqual(freeze.call_count, 1)
        self.assertEqual(self.fixture.profile_tree(), original)
        _, index = self.client._load()
        self.assertTrue((self.fixture.state / 'snapshots' / index['source_snapshot'] / 'deferred').is_dir())

    def test_deferred_cache_is_bound_to_message_identity(self):
        self.refresh()
        directory, index = self.client._current()
        original = next(row for row in index['items'] if row['subject'] == '中文主题乙')
        twin = copy.deepcopy(original)
        twin['id'] = f"{original['id']}-different"
        twin['message_id'] = '<different-message@example.test>'
        index['items'].append(twin)

        with patch.object(service.core, 'freeze_body', wraps=service.core.freeze_body) as freeze:
            service.core.read_message(directory, index, original['id'], 20_000)
            service.core.read_message(directory, index, twin['id'], 20_000)
        self.assertEqual(freeze.call_count, 2)

    def test_concurrent_deferred_reads_materialize_once(self):
        self.refresh()
        directory, index = self.client._current()
        with patch.object(service.core, 'freeze_body', wraps=service.core.freeze_body) as freeze:
            with ThreadPoolExecutor(max_workers=4) as pool:
                results = list(pool.map(lambda _: service.core.read_message(
                    directory, index, self.ids[1], 20000), range(8)))
        self.assertEqual(freeze.call_count, 1)
        self.assertTrue(all(r['body']['status'] == 'cached_html' for r in results))
        self.assertEqual(len({r['body']['sha256'] for r in results}), 1)

    def test_cached_body_survives_source_change_after_freeze(self):
        self.refresh()
        listing = self.client.call('outlook_local_search', {})
        identifier = next(item['id'] for item in listing['items'] if item['subject'] == '中文主题乙')
        first = self.client.call('outlook_local_read', {'id': identifier, 'body_max_chars': 2})
        self.assertEqual(first['body']['status'], 'cached_html')
        self.fixture.gzip_body_source.write_bytes(b"changed after freeze")
        second = self.client.call('outlook_local_read', {'id': identifier})
        self.assertEqual(second['body']['status'], 'cached_html')
        self.assertIn('gzip 冻结正文', second['body']['text'])

    def test_deferred_read_repairs_legacy_partial_payload_with_cached_metadata(self):
        self.refresh()
        first = self.client.call('outlook_local_read', {'id': self.ids[1]})
        directory, _ = self.client._load()
        for root, key in (('payloads', 'sha256'), ('textpayloads', 'text_sha256')):
            with self.subTest(root=root):
                payload = directory / root / first['body'][key]
                payload.chmod(0o600)
                payload.write_bytes(b'partial')
                recovered = self.client.call('outlook_local_read', {'id': self.ids[1]})
                self.assertEqual(recovered['body']['text'], first['body']['text'])
                self.assertEqual(recovered['body'][key], first['body'][key])

    def test_deferred_body_rejects_changed_source_before_read(self):
        self.refresh()
        listing = self.client.call('outlook_local_search', {})
        identifier = next(item['id'] for item in listing['items'] if item['subject'] == '中文主题乙')
        self.fixture.gzip_body_source.write_bytes(self.fixture.gzip_body_source.read_bytes() + b' changed')
        with self.assertRaisesRegex(ValueError, 'body source changed since snapshot'):
            self.client.call('outlook_local_read', {'id': identifier})

    def test_deferred_read_uses_private_container_after_outlook_updates(self):
        source_bytes = self.fixture.source_cache.read_bytes()
        self.refresh()
        directory, index = self.client._load()
        self.assertEqual((directory / 'source.hxd').read_bytes(), source_bytes)
        self.fixture.source_cache.write_bytes(b'Outlook synchronised a new container')
        message = self.client.call('outlook_local_read', {'id': self.ids[1]})
        self.assertEqual(message['body']['status'], 'cached_html')
        self.assertEqual((directory / 'source.hxd').read_bytes(), source_bytes)

    def test_deferred_read_rejects_modified_private_container(self):
        self.refresh()
        directory, index = self.client._load()
        source = directory / 'source.hxd'
        source.chmod(0o600)
        source.write_bytes(source.read_bytes() + b' corrupt')
        with self.assertRaisesRegex(ValueError, 'frozen HxStore source changed'):
            self.client.call('outlook_local_read', {'id': self.ids[1]})

    def test_refresh_repairs_private_container_without_changing_original_mail(self):
        original = self.fixture.profile_tree()
        self.refresh()
        for damage in ('modified', 'missing'):
            with self.subTest(damage=damage):
                directory, old_index = self.client._load()
                source = directory / 'source.hxd'
                if damage == 'modified':
                    source.chmod(0o600)
                    source.write_bytes(source.read_bytes() + b' corrupt')
                else:
                    source.unlink()
                repaired = self.refresh()
                new_directory, new_index = self.client._load()
                self.assertNotEqual(new_index['source_snapshot'], old_index['source_snapshot'])
                self.assertEqual((new_directory / 'source.hxd').read_bytes(), self.fixture.source_cache.read_bytes())
                body = self.client.call('outlook_local_read', {'id': self.ids[1]})['body']
                self.assertEqual(body['status'], 'cached_html')
                self.assertEqual(self.refresh()['snapshot'], repaired['snapshot'])
        self.assertEqual(self.fixture.profile_tree(), original)

    def test_deferred_bundle_checks_source_budget_before_freezing(self):
        self.refresh()
        with patch.object(service, 'BUNDLE_SOURCE_BYTES', 1), patch.object(
                service.core, 'freeze_body', wraps=service.core.freeze_body) as freeze:
            result = self.client.call('outlook_local_messages_bundle', {'ids': [self.ids[1]]})
        self.assertEqual(freeze.call_count, 0)
        self.assertEqual(result['items'], [])
        self.assertIn('source byte budget', result['errors'][0]['error'])

    def test_failed_deferred_bundle_read_still_consumes_source_budget(self):
        self.refresh()
        directory, index = self.client._load()
        row = copy.deepcopy(next(r for r in index['items'] if r['id'] == self.ids[1]))
        twin = {**row, 'id': row['id'] + '-second'}
        index['items'] = [row, twin]
        budget = sum((fp or {}).get('size', 0) for fp in
                     row['body']['deferred']['external_fingerprints'].values())
        budget += sum(index['source_block_bytes'].values())
        self.assertGreater(budget, 0)
        with patch.object(service, 'BUNDLE_SOURCE_BYTES', budget), patch.object(
                self.client, '_current', return_value=(directory, index)), patch.object(
                service.core, 'stable_read', side_effect=OSError('body read failed')) as read:
            result = self.client.call('outlook_local_messages_bundle', {'ids': [row['id'], twin['id']]})
        self.assertEqual(read.call_count, 1)
        self.assertEqual(len(result['items']), 1)
        self.assertEqual(len(result['errors']), 1)
        self.assertIn('source byte budget', result['errors'][0]['error'])

    def test_bundle_bounds_external_gzip_inflation_and_can_retry(self):
        html = ('<html><body>' + fixtures.GZIP_PREVIEW + 'x' * 1024**2 + '</body></html>').encode()
        compressed = gzip.compress(html)
        self.assertLess(len(compressed), len(html) // 100)
        self.fixture.gzip_body_source.write_bytes(compressed)
        self.refresh()
        directory, index = self.client._load()
        row = next(row for row in index['items'] if row['subject'] == '中文主题乙')
        twin = {**row, 'id': row['id'] + '-second'}
        index['items'] = [row, twin]
        budget = (len(html) + len(compressed) + sum(index['source_block_bytes'].values())) * 3 // 2
        with patch.object(self.client, '_current', return_value=(directory, index)), \
                patch.object(service, 'BUNDLE_SOURCE_BYTES', budget):
            result = self.client.call('outlook_local_messages_bundle', {'ids': [row['id'], twin['id']]})
            self.assertEqual(len(result['items']), 1)
            self.assertIn('source byte budget', result['errors'][0]['error'])
            self.assertIsNone(service.core._read_deferred_cache(directory, twin))
            retried = self.client.call('outlook_local_read', {'id': twin['id'], 'body_max_chars': 100})
            self.assertEqual(retried['body']['status'], 'cached_html')

    def test_bundle_skips_oversized_source_but_reads_later_small_message(self):
        self.fixture.body_source.write_bytes(
            ('<html><body>' + fixtures.PREVIEW + 'x' * 100_000 + '</body></html>').encode())
        large_id = next(row['id'] for row in self.base['items']
                        if row['message_id'] == '<first@example.test>')
        self.scope['messages'][large_id]['folder_id'] = 'inbox'
        self.refresh()
        _, index = self.client._load()
        large = next(row for row in index['items'] if row['message_id'] == '<first@example.test>')
        small = next(row for row in index['items'] if row['message_id'] == '<second@example.test>')
        with patch.object(service, 'BUNDLE_SOURCE_BYTES', 48 * 1024), \
                patch.object(service.core, 'stable_read', wraps=service.core.stable_read) as read:
            result = self.client.call('outlook_local_messages_bundle', {'ids': [large['id'], small['id']]})
        self.assertEqual(read.call_count, 1)
        self.assertEqual([item['id'] for item in result['items']], [small['id']])
        self.assertEqual(result['errors'][0]['id'], large['id'])
        self.assertIn('source byte budget', result['errors'][0]['error'])

    def test_excluded_folder_read_does_not_freeze_body(self):
        self.refresh()
        with patch.object(service.core, 'freeze_body', wraps=service.core.freeze_body) as freeze:
            with self.assertRaises(ValueError):
                self.client.call('outlook_local_read', {'id': self.ids[2]})
        self.assertEqual(freeze.call_count, 0)

    def test_bundle_counts_decoded_blocks_for_tiny_inline_bodies(self):
        self.refresh()
        directory, index = self.client._load()
        row = index['items'][0]
        index['items'] = [{**row, 'id': f'block-{n}', 'body': {
            'status': 'deferred', 'deferred': {'external_fingerprints': {},
                                             'inline_frames': [{'full_length': 1}]}}}
                          for n in range(3)]
        index['source_block_bytes'] = {str(n): 24 * 1024**2 for n in range(3)}

        def read(directory, index, identifier, maximum):
            offset = int(identifier.rsplit('-', 1)[1])
            service.core._read_source_block(directory, index, offset, {})
            return {'id': identifier, 'body': {'text': 'x'}}

        with patch.object(self.client, '_current', return_value=(directory, index)), \
                patch.object(self.client, '_read', side_effect=read), \
                patch('hxcontainer.read_block_at', return_value=self.fixture.block) as decode:
            result = self.client.call('outlook_local_messages_bundle', {
                'ids': [row['id'] for row in index['items']]})
        self.assertEqual(decode.call_count, 2)
        self.assertEqual(len(result['items']), 2)
        self.assertIn('source byte budget', result['errors'][0]['error'])

    def test_native_today_does_not_use_sent_time_when_receipt_is_missing(self):
        self.scope['messages'][self.ids[0]].update(received_unix=None, received_time_status='missing',
                                                  received_time_method=None)
        self.refresh()
        listing = self.client.call('outlook_local_search', {})
        self.assertEqual(listing['total_matched'], 2)
        self.assertEqual(listing['received_time_unknown'], 1)
        with patch.object(service, 'datetime', FrozenDatetime):
            FrozenDatetime.instant = datetime(1999, 12, 31, 16, 30, tzinfo=timezone.utc)
            page = self.client.call('outlook_local_today', {})
        self.assertEqual(page['total_matched'], 1)
        self.assertTrue(page['date_filter_incomplete'])
        self.assertEqual(page['received_time_unknown'], 1)
        undated = self.client.call('outlook_local_read', {'id': self.ids[0]})
        self.assertIsNone(undated['received_at'])
        self.assertIsNotNone(undated['sent_at'])

    def test_status_does_not_scan_or_create_state(self):
        fresh = self.fixture.state.parent / 'fresh'
        client = service.MailService(fresh, self.fixture.profile)
        self.assertFalse(client.call('outlook_local_status', {})['ready'])
        self.assertFalse(fresh.exists())

    def test_scoped_view_and_source_roots_require_private_permissions(self):
        self.refresh()
        directory, index = self.client._load()
        view = self.client.state / 'views' / index['snapshot']
        view.chmod(0o755)
        with self.assertRaisesRegex(ValueError, 'private permissions'):
            self.client._load()
        view.chmod(0o700)

        source_snapshot = directory
        source_snapshot.chmod(0o755)
        with self.assertRaisesRegex(ValueError, 'private permissions'):
            self.client._load()
        source_snapshot.chmod(0o700)

        for name in ('payloads', 'textpayloads'):
            root = source_snapshot / name
            root.chmod(0o755)
            with self.assertRaisesRegex(ValueError, 'private permissions'):
                self.client._load()
            root.chmod(0o700)

        exports = self.client.state / 'exports'
        exports.mkdir(mode=0o700)
        exports.chmod(0o755)
        with self.assertRaisesRegex(ValueError, 'private permissions'):
            self.client._load()
        exports.chmod(0o700)

    def test_restarted_reader_cannot_reuse_another_profiles_fresh_view(self):
        before = self.refresh()['snapshot']
        profile = self.fixture.profile.parent / 'profile-b'
        shutil.copytree(self.fixture.profile, profile)
        other = service.MailService(self.fixture.state, profile)
        self.assertFalse(other._fresh())
        self.assertFalse(other.call('outlook_local_status', {})['ready'])
        with patch.object(service.core, 'make_snapshot', wraps=service.core.make_snapshot) as make:
            result = other.call('outlook_local_search', {})
        self.assertEqual(make.call_count, 1)
        _, index = other._load()
        self.assertEqual(index['profile'], str(profile.resolve()))
        self.assertNotEqual(result['snapshot'], before)
        with self.assertRaises((ValueError, FileNotFoundError)):
            other._current(before)

    def test_retry_cleans_builds_left_by_terminated_worker(self):
        script = '''
import os, pathlib, sys, tempfile, time
state = pathlib.Path(sys.argv[1])
for kind in ('building', 'view'):
    root = pathlib.Path(tempfile.mkdtemp(prefix=f'.{kind}-{os.getpid()}-', dir=state))
    (root / 'partial-body').write_bytes(b'private partial cache')
print('READY', flush=True)
time.sleep(30)
'''
        child = subprocess.Popen([sys.executable, '-c', script, str(self.fixture.state)],
                                 stdout=subprocess.PIPE, text=True)
        try:
            self.assertEqual(child.stdout.readline().strip(), 'READY')
            child.kill()
            child.wait(timeout=5)
            self.refresh()
            self.assertFalse(any(self.fixture.state.glob('.building-*')))
            self.assertFalse(any(self.fixture.state.glob('.view-*')))
        finally:
            if child.poll() is None:
                child.kill()
            child.wait(timeout=5)
            child.stdout.close()

    def test_build_cleanup_preserves_live_owners_and_never_follows_links(self):
        live = self.fixture.state / f'.building-{os.getpid()}-active'
        live.mkdir()
        outside = self.fixture.state.parent / 'keep-build-source'
        outside.mkdir()
        (outside / 'sentinel').write_bytes(b'original')
        old_link = self.fixture.state / '.building-old-version'
        old_link.symlink_to(outside, target_is_directory=True)
        os.utime(old_link, (0, 0), follow_symlinks=False)
        self.refresh()
        self.assertTrue(live.exists())
        self.assertFalse(old_link.is_symlink())
        self.assertEqual((outside / 'sentinel').read_bytes(), b'original')

    def test_explicit_refresh_date_survives_the_next_read(self):
        self.refresh()
        result = self.client.call('outlook_local_refresh', {'from_date': '2001-01-01'})
        self.assertEqual(self.client.from_date, '2001-01-01')
        with patch.object(service.core, 'make_snapshot', side_effect=AssertionError('unexpected rebuild')):
            read = self.client.call('outlook_local_search', {})
        self.assertEqual(read['snapshot'], result['snapshot'])
        self.assertEqual(read['items'], [])

    def test_failed_explicit_refresh_does_not_change_requested_date(self):
        self.refresh()
        self.scope['folders'] = []
        with self.assertRaisesRegex(ValueError, 'no verified'):
            self.client.call('outlook_local_refresh', {'from_date': '2001-01-01'})
        self.assertEqual(self.client.from_date, '2000-01-01')

    def test_scope_includes_recursive_inbox_and_sent_excludes_deleted_and_unknown(self):
        original = self.fixture.profile_tree()
        result = self.refresh()
        self.assertEqual(result['messages'], 2)
        self.assertEqual(result['folders'], 4)
        listing = self.client.call('outlook_local_folders', {})
        self.assertEqual({f['id'] for f in listing['items']}, {'inbox', 'child', 'deep', 'sent'})
        self.assertEqual(
            {f['id']: f['cached_messages'] for f in listing['items']},
            {'inbox': 1, 'child': 0, 'deep': 1, 'sent': 0},
        )
        self.assertEqual(next(f['path'] for f in listing['items'] if f['id'] == 'deep'), '收件箱/研究/二级')
        self.assertEqual(self.fixture.profile_tree(), original)
        for method in ('outlook_local_read', 'outlook_local_attachments', 'outlook_local_attachment'):
            with self.assertRaises(ValueError):
                self.client.call(method, {'id': self.ids[2], 'attachment_id': 'fake'})
        self.scope['messages'][self.ids[2]]['folder_id'] = 'sent'
        self.scope['fingerprint'] = 'sent-catalog'
        self.assertEqual(self.refresh()['messages'], 3)
        self.scope['messages'].pop(self.ids[2])
        self.scope['fingerprint'] = 'unknown-catalog'
        self.assertEqual(self.refresh()['unresolved_folder_excluded'], 1)

    def test_search_recursion_and_filter_cursor_mismatch(self):
        self.refresh()
        parent = self.client.call('outlook_local_search', {'folder_id': 'inbox'})
        self.assertEqual(len(parent['items']), 2)
        self.assertEqual(parent['items'][1]['folder_path'], '收件箱/研究/二级')
        single = self.client.call('outlook_local_search', {'folder_id': 'inbox', 'include_subfolders': False})
        self.assertEqual(len(single['items']), 1)
        first = self.client.call('outlook_local_search', {'limit': 1})
        with self.assertRaises(ValueError):
            self.client.call('outlook_local_search', {'folder_id': 'inbox', 'limit': 1, 'cursor': first['next_cursor']})
        with self.assertRaises(ValueError):
            self.client.call('outlook_local_search', {'folder_id': 'deleted'})

    def test_today_uses_singapore_received_date_and_existing_read_only_scope(self):
        self.refresh()
        original = self.fixture.profile_tree()
        FrozenDatetime.instant = datetime(1999, 12, 31, 16, 30, tzinfo=timezone.utc)
        with patch.object(service, 'datetime', FrozenDatetime):
            result = self.client.call('outlook_local_today', {})
            self.assertEqual(result['date'], '2000-01-01')
            self.assertEqual(result['timezone'], 'Asia/Singapore')
            self.assertEqual(result['date_basis'], 'received_at')
            self.assertEqual(result['total_matched'], 2)
            self.assertTrue(result['complete'])
            self.assertEqual({row['id'] for row in result['items']}, set(self.ids[:2]))
            nested = self.client.call('outlook_local_today', {'folder_id': 'inbox', 'include_subfolders': False})
            self.assertEqual(nested['total_matched'], 1)
            with self.assertRaises(ValueError):
                self.client.call('outlook_local_today', {'folder_id': 'deleted'})
        self.assertEqual(self.fixture.profile_tree(), original)

    def test_today_keeps_first_page_date_across_midnight(self):
        self.refresh()
        FrozenDatetime.instant = datetime(2000, 1, 1, 15, 59, tzinfo=timezone.utc)
        with patch.object(service, 'datetime', FrozenDatetime):
            first = self.client.call('outlook_local_today', {'limit': 1})
            FrozenDatetime.instant = datetime(2000, 1, 1, 16, 1, tzinfo=timezone.utc)
            second = self.client.call('outlook_local_today', {'limit': 1, 'cursor': first['next_cursor']})
            self.assertEqual(second['date'], '2000-01-01')
            self.assertEqual(second['snapshot'], first['snapshot'])
            self.assertTrue(second['complete'])
            self.assertNotEqual(first['items'][0]['id'], second['items'][0]['id'])
            new_day = self.client.call('outlook_local_today', {})
            self.assertEqual(new_day['date'], '2000-01-02')
            self.assertEqual(new_day['items'], [])

    def test_today_rejects_foreign_or_modified_cursors(self):
        self.refresh()
        FrozenDatetime.instant = datetime(2000, 1, 1, tzinfo=timezone.utc)
        with patch.object(service, 'datetime', FrozenDatetime):
            search = self.client.call('outlook_local_search', {'limit': 1})
            with self.assertRaisesRegex(ValueError, 'invalid today cursor'):
                self.client.call('outlook_local_today', {'limit': 1, 'cursor': search['next_cursor']})
            with self.assertRaisesRegex(ValueError, 'invalid today cursor'):
                self.client.call('outlook_local_today', {'cursor': 'today-v1|2000-01-01|'})
            first = self.client.call('outlook_local_today', {'limit': 1})
            modified = first['next_cursor'].replace('2000-01-01', '2000-01-02')
            with self.assertRaises(ValueError):
                self.client.call('outlook_local_today', {'limit': 1, 'cursor': modified})

    def test_cursor_and_reads_stay_on_old_snapshot_after_other_client_refresh(self):
        before = self.refresh()['snapshot']
        first = self.client.call('outlook_local_search', {'limit': 1})
        self.scope['messages'][self.ids[1]]['folder_id'] = 'sent'
        self.scope['fingerprint'] = 'moved-catalog'
        other = service.MailService(self.fixture.state, self.fixture.profile)
        after = other.call('outlook_local_refresh', {})['snapshot']
        self.assertNotEqual(before, after)
        second = self.client.call('outlook_local_search', {'limit': 1, 'cursor': first['next_cursor']})
        self.assertEqual(second['snapshot'], before)
        self.assertEqual(len(second['items']), 1)
        pinned = self.client.call('outlook_local_read', {'id': self.ids[1], 'snapshot': before})
        self.assertEqual(pinned['snapshot'], before)
        self.assertEqual(self.client.call('outlook_local_read', {'id': self.ids[1], 'snapshot': after})['folder_path'], '已发送')
        with self.assertRaises(ValueError):
            self.client.call('outlook_local_read', {'id': self.ids[1], 'snapshot': '../escape'})

    def test_latest_deleted_membership_revokes_old_snapshot_access(self):
        before = self.refresh()['snapshot']
        first = self.client.call('outlook_local_search', {'limit': 1})
        self.scope['messages'][self.ids[1]]['folder_id'] = 'deleted'
        self.scope['fingerprint'] = 'deleted-catalog'
        self.refresh()
        with self.assertRaisesRegex(ValueError, 'folder access changed'):
            self.client.call('outlook_local_search', {'limit': 1, 'cursor': first['next_cursor']})
        for method in ('outlook_local_read', 'outlook_local_attachments', 'outlook_local_attachment'):
            with self.assertRaisesRegex(ValueError, 'folder access changed'):
                self.client.call(method, {'id': self.ids[1], 'snapshot': before, 'attachment_id': 'anything'})

    def test_empty_folder_scope_revocation_invalidates_old_snapshot(self):
        self.scope['folders'].append(dict(id='empty', parent_id='inbox', name='空文件夹'))
        before = self.refresh()['snapshot']
        # The child is empty: only comparing message IDs misses its revocation.
        self.scope['folders'] = [row for row in self.scope['folders'] if row['id'] != 'empty']
        self.scope['fingerprint'] = 'empty-folder-revoked'
        self.refresh()
        with self.assertRaisesRegex(ValueError, 'folder access changed'):
            self.client._current(before)

    def test_folder_reparenting_invalidates_old_snapshot(self):
        before = self.refresh()['snapshot']
        next(row for row in self.scope['folders'] if row['id'] == 'child')['parent_id'] = 'sent'
        self.scope['fingerprint'] = 'folder-reparented'
        self.refresh()
        with self.assertRaisesRegex(ValueError, 'folder access changed'):
            self.client._current(before)

    def test_cached_bundle_needs_no_source_budget_or_inflation(self):
        self.refresh()
        for identifier in self.ids[:2]:
            self.client.call('outlook_local_read', {'id': identifier})
        with patch.object(service, 'BUNDLE_SOURCE_BYTES', 0), \
                patch.object(service.core, 'stable_read', side_effect=AssertionError('source read')), \
                patch.object(service.core, 'decode_body', side_effect=AssertionError('body inflation')):
            result = self.client.call('outlook_local_messages_bundle', {'ids': self.ids[:2]})
        self.assertEqual(len(result['items']), 2)
        self.assertEqual(result['errors'], [])

    def test_cached_bundle_bounds_cold_verification_and_reuses_verified_payloads(self):
        self.fixture.body_source.write_bytes(
            ('<html><body>' + fixtures.PREVIEW + 'x' * 2048 + '</body></html>').encode())
        self.fixture.gzip_body_source.write_bytes(
            ('<html><body>' + fixtures.GZIP_PREVIEW + 'y' * 2048 + '</body></html>').encode())
        first_id = next(row['id'] for row in self.base['items']
                        if row['message_id'] == '<first@example.test>')
        second_id = next(row['id'] for row in self.base['items']
                         if row['message_id'] == '<second@example.test>')
        self.scope['messages'][first_id]['folder_id'] = 'inbox'
        self.refresh()
        for identifier in (first_id, second_id):
            self.client.call('outlook_local_read', {'id': identifier})
        directory, index = self.client._load()
        first = service.core.find_message(index, first_id)
        cached = service.core._read_deferred_cache(directory, first)
        verification_bytes = cached['bytes'] + cached['text_bytes']
        # A new process starts with neither verification memo populated.
        service.core._verified_raw_payloads.clear()
        service.core._verified_text_chunks.clear()
        with patch.object(service, 'BUNDLE_SOURCE_BYTES', 0), \
                patch.object(service, 'BUNDLE_VERIFY_BYTES', verification_bytes), \
                patch.object(service.core, 'freeze_body', side_effect=AssertionError('cache rebuilt')):
            result = self.client.call('outlook_local_messages_bundle', {'ids': [first_id, second_id]})
        self.assertEqual([item['id'] for item in result['items']], [first_id])
        self.assertEqual(result['errors'][0]['id'], second_id)
        self.assertIn('cache verification byte budget', result['errors'][0]['error'])
        with patch.object(service, 'BUNDLE_SOURCE_BYTES', 0), \
                patch.object(service, 'BUNDLE_VERIFY_BYTES', 0):
            warm = self.client.call('outlook_local_messages_bundle', {'ids': [first_id]})
        self.assertEqual(len(warm['items']), 1)
        self.assertEqual(warm['errors'], [])
        self.assertEqual(self.client.call('outlook_local_read', {'id': second_id})['body']['status'],
                         'cached_html')

    def test_expired_pinned_snapshot_refreshes_before_deleted_access(self):
        before = self.refresh()['snapshot']
        first = self.client.call('outlook_local_search', {'limit': 1})
        self.scope['messages'][self.ids[1]]['folder_id'] = 'deleted'
        self.scope['fingerprint'] = 'deleted-without-explicit-refresh'
        self.client.refresh_seconds = 0
        for method in ('outlook_local_read', 'outlook_local_attachments', 'outlook_local_attachment'):
            with self.assertRaisesRegex(ValueError, 'folder access changed'):
                self.client.call(method, {'id': self.ids[1], 'snapshot': before, 'attachment_id': 'anything'})
        with self.assertRaisesRegex(ValueError, 'folder access changed'):
            self.client.call('outlook_local_search', {'limit': 1, 'cursor': first['next_cursor']})
        self.assertNotEqual(self.client._status()['snapshot'], before)

    def test_expired_pinned_snapshot_fails_closed_when_refresh_fails(self):
        before = self.refresh()['snapshot']
        self.scope['folders'] = []
        self.client.refresh_seconds = 0
        with self.assertRaisesRegex(ValueError, 'no verified'):
            self.client.call('outlook_local_read', {'id': self.ids[1], 'snapshot': before})
        self.assertEqual(self.client._status()['snapshot'], before)

    def test_body_budget_and_attachment_association_owned_export(self):
        self.refresh()
        message = self.client.call('outlook_local_read', {'id': self.ids[0]})
        self.assertNotIn('text_chunk_sha256', message['body'])
        bundle = self.client.call('outlook_local_messages_bundle', {'ids': self.ids + [self.ids[1]], 'text_budget_chars': 8})
        self.assertLessEqual(bundle['body_text_chars'], 8)
        self.assertTrue(bundle['errors'])
        attachment_row = next(r for r in self.base['items'] if r['attachments'])
        # This fixture's attachment is on the originally excluded row; explicitly classify it in Sent.
        self.scope['messages'][attachment_row['id']]['folder_id'] = 'sent'
        self.scope['fingerprint'] = 'attachment-catalog'
        self.refresh()
        arguments = {'id': attachment_row['id'], 'attachment_id': attachment_row['attachments'][0]['id']}
        original = self.fixture.profile_tree()
        with patch.object(service.core, 'read_attachment', wraps=service.core.read_attachment) as export:
            result = self.client.call('outlook_local_attachment', arguments)
            self.assertEqual(self.client.call('outlook_local_attachment', arguments)['output'], result['output'])
            self.assertEqual(export.call_count, 1)
        self.assertNotIn('_export_fingerprint', result)
        target = Path(result['output'])
        self.assertTrue(target.is_relative_to(self.fixture.state.resolve() / 'exports'))
        self.assertEqual(hashlib.sha256(target.read_bytes()).hexdigest(), result['sha256'])
        self.assertEqual(self.client.call('outlook_local_attachment', arguments)['output'], str(target))
        target.unlink()
        rebuilt = self.client.call('outlook_local_attachment', arguments)
        self.assertEqual(rebuilt['sha256'], result['sha256'])
        self.assertTrue(target.is_file())
        self.assertEqual(self.fixture.profile_tree(), original)
        self.fixture.attachment_source.write_bytes(b'updated by Outlook')
        with self.assertRaises(ValueError):
            self.client.call('outlook_local_attachment', arguments)

    def test_cached_export_change_is_rejected(self):
        row = next(r for r in self.base['items'] if r['attachments'])
        self.scope['messages'][row['id']]['folder_id'] = 'sent'
        self.refresh()
        arguments = {'id': row['id'], 'attachment_id': row['attachments'][0]['id']}
        result = self.client.call('outlook_local_attachment', arguments)
        target = Path(result['output'])
        target.chmod(0o600)
        target.write_bytes(b'changed export')
        with self.assertRaisesRegex(ValueError, 'export changed'):
            self.client.call('outlook_local_attachment', arguments)

    def test_export_eviction_removes_unreferenced_files(self):
        row = next(row for row in self.base['items'] if row['attachments'])
        self.scope['messages'][row['id']]['folder_id'] = 'sent'
        snapshot = self.refresh()['snapshot']
        directory, index = self.client._load(snapshot)
        message = service.core.find_message(index, row['id'])
        attachment = message['attachments'][0]
        message['attachments'] = [dict(attachment, id=f'attachment-{i}') for i in range(3)]
        def export(directory, index, message_id, attachment_id, output):
            data = attachment_id.encode()
            Path(output).write_bytes(data)
            return dict(snapshot=snapshot, message_id=message_id, attachment_id=attachment_id,
                        name=None, bytes=len(data), sha256=service.core.sha(data), output=str(output))
        with patch.object(service, 'MAX_EXPORT_FILES', 2), patch.object(service.core, 'read_attachment', side_effect=export):
            results = [self.client._attachment(directory, index, {'id': row['id'], 'attachment_id': f'attachment-{i}'})
                       for i in range(3)]
        self.assertFalse(Path(results[0]['output']).exists())
        self.assertTrue(Path(results[1]['output']).exists())
        self.assertTrue(Path(results[2]['output']).exists())
        self.assertEqual(len(self.client._attachment_exports), 2)
        self.assertEqual(len(list((self.fixture.state / 'exports' / snapshot).iterdir())), 2)

    def test_export_byte_budget_applies_across_process_cache_maps(self):
        snapshot = self.refresh()['snapshot']
        (self.fixture.state / 'exports').mkdir(mode=0o700)
        directory = self.fixture.state / 'exports' / snapshot
        directory.mkdir(mode=0o700)
        paths = [directory / (str(i) * 64) for i in range(3)]
        for path in paths:
            path.write_bytes(b'123456')
        other = service.MailService(self.fixture.state, self.fixture.profile)
        with patch.object(service, 'MAX_EXPORT_BYTES', 12), other._lock():
            other._prune_exports(protected=paths[0])
        self.assertTrue(paths[0].exists())
        self.assertEqual(sum(path.stat().st_size for path in directory.iterdir()), 12)

    def test_terminated_export_is_cleaned_before_the_next_copy(self):
        row = next(row for row in self.base['items'] if row['attachments'])
        self.scope['messages'][row['id']]['folder_id'] = 'sent'
        snapshot = self.refresh()['snapshot']
        (self.fixture.state / 'exports').mkdir(mode=0o700)
        exports = self.fixture.state / 'exports' / snapshot
        exports.mkdir(mode=0o700)
        script = '''
import os, pathlib, sys, time
path = pathlib.Path(sys.argv[1]) / f'.export-{os.getpid()}-interrupted'
path.write_bytes(b'partial export')
print('READY', flush=True)
time.sleep(30)
'''
        child = subprocess.Popen([sys.executable, '-c', script, str(exports)],
                                 stdout=subprocess.PIPE, text=True)
        try:
            self.assertEqual(child.stdout.readline().strip(), 'READY')
            child.kill()
            child.wait(timeout=5)
            self.client.call('outlook_local_attachment', {
                'id': row['id'], 'attachment_id': row['attachments'][0]['id']})
            self.assertEqual(list(exports.glob('.export-*')), [])
        finally:
            if child.poll() is None:
                child.kill()
            child.wait(timeout=5)
            child.stdout.close()

    def test_active_export_bytes_reserve_quota_before_new_copy(self):
        row = next(row for row in self.base['items'] if row['attachments'])
        self.scope['messages'][row['id']]['folder_id'] = 'sent'
        snapshot = self.refresh()['snapshot']
        (self.fixture.state / 'exports').mkdir(mode=0o700)
        exports = self.fixture.state / 'exports' / snapshot
        exports.mkdir(mode=0o700)
        active = exports / f'.export-{os.getpid()}-active'
        active.write_bytes(b'12345678')
        size = row['attachments'][0]['fingerprint']['size']
        with patch.object(service, 'MAX_EXPORT_BYTES', size + 7), \
                patch.object(service.core, 'read_attachment') as read:
            with self.assertRaisesRegex(ValueError, 'quota exhausted'):
                self.client.call('outlook_local_attachment', {
                    'id': row['id'], 'attachment_id': row['attachments'][0]['id']})
        read.assert_not_called()
        self.assertTrue(active.exists())

    def test_status_reports_unready_for_bounded_cache_state_errors(self):
        self.refresh()
        for error in (FileNotFoundError('missing view'), KeyError('profile'), TypeError('invalid view')):
            with self.subTest(error=type(error).__name__), patch.object(self.client, '_load', side_effect=error):
                result = self.client.call('outlook_local_status', {})
            self.assertFalse(result['ready'])
            self.assertIn('next_step', result)

    def test_unverified_membership_and_index_corruption_rejected(self):
        self.scope['messages'][self.ids[0]]['folder_status'] = 'unresolved'
        self.assertEqual(self.refresh()['unresolved_folder_excluded'], 1)
        path = self.client.state / 'views' / self.client._status()['snapshot'] / 'index.json'
        path.chmod(0o600)
        path.write_text(path.read_text().replace('收件箱', 'changed'))
        with self.assertRaisesRegex(ValueError, 'checksum'):
            self.client._load()

    def test_failed_scope_refresh_retains_previous_view(self):
        before = self.refresh()['snapshot']
        self.scope['folders'] = []
        with self.assertRaises(ValueError):
            self.refresh()
        self.assertEqual(self.client.call('outlook_local_status', {})['snapshot'], before)
        self.assertFalse(self.client._fresh())
        with self.assertRaisesRegex(ValueError, 'no verified'):
            self.client.call('outlook_local_read', {'id': self.ids[0], 'snapshot': before})

    def test_failed_osa_read_invalidates_freshness_for_other_clients(self):
        before = self.refresh()['snapshot']
        self.assertTrue(self.client._fresh())
        with patch.object(sys.modules['osa_scope'], 'read_scope', side_effect=ValueError('unsafe Osa log')):
            with self.assertRaisesRegex(ValueError, 'unsafe Osa'):
                self.refresh()
            other = service.MailService(self.fixture.state, self.fixture.profile)
            self.assertFalse(other._fresh())
            with self.assertRaisesRegex(ValueError, 'unsafe Osa'):
                other.call('outlook_local_read', {'id': self.ids[0], 'snapshot': before})

    def test_refresh_bounds_views_bases_and_exports_without_touching_source(self):
        row = next(row for row in self.base['items'] if row['attachments'])
        self.scope['messages'][row['id']]['folder_id'] = 'sent'
        original = self.fixture.profile_tree()
        orphan = self.fixture.state / 'snapshots' / ('e' * 24)
        orphan.mkdir()
        (orphan / 'payloads').mkdir()
        (orphan / 'payloads' / 'old-body').write_bytes(b'old frozen body')
        snapshots, outputs = [], []
        for position in range(5):
            self.scope['fingerprint'] = f'catalog-{position}'
            snapshots.append(self.refresh()['snapshot'])
            result = self.client.call('outlook_local_attachment', {
                'id': row['id'], 'attachment_id': row['attachments'][0]['id']})
            outputs.append(Path(result['output']))
        retained = set(snapshots[-service.RETAINED_VIEWS:])
        self.assertEqual({path.name for path in (self.fixture.state / 'views').iterdir()}, retained)
        self.assertEqual({path.name for path in (self.fixture.state / 'exports').iterdir()}, retained)
        self.assertFalse(outputs[0].exists())
        self.assertTrue(outputs[-1].exists())
        self.assertEqual(len(list((self.fixture.state / 'snapshots').iterdir())), 1)
        self.assertFalse(orphan.exists())
        self.assertEqual({key[0] for key in self.client._attachment_exports}, retained)
        with self.assertRaises(FileNotFoundError):
            self.client._load(snapshots[0])
        self.assertEqual(self.client._load(snapshots[-2])[1]['snapshot'], snapshots[-2])
        self.assertEqual(self.fixture.profile_tree(), original)

    def test_pruning_does_not_follow_state_symlinks(self):
        self.refresh()
        outside = self.fixture.state.parent / 'outside-pruning'
        outside.mkdir()
        sentinel = outside / 'keep'
        sentinel.write_text('preserve')
        link = self.fixture.state / 'views' / ('f' * 24)
        link.symlink_to(outside, target_is_directory=True)
        self.scope['fingerprint'] = 'prune-symlink'
        self.refresh()
        self.assertFalse(link.is_symlink())
        self.assertEqual(sentinel.read_text(), 'preserve')

    def test_schema_upgrade_prunes_unreadable_history_after_publishing(self):
        self.refresh()
        _, old = self.client._load()
        old = {key: value for key, value in old.items() if key != 'snapshot'}
        old['schema_version'] = service.core.VERSION - 1
        identifier = service.core.sha(json.dumps(old, sort_keys=True, ensure_ascii=False).encode())[:24]
        old['snapshot'] = identifier
        directory = self.fixture.state / 'views' / identifier
        directory.mkdir()
        (directory / 'index.json').write_text(json.dumps(old, ensure_ascii=False))
        self.scope['fingerprint'] = 'upgraded-schema'
        self.refresh()
        self.assertTrue(self.client._fresh())
        self.assertFalse(directory.exists())

    def test_repeated_failed_scope_refreshes_do_not_accumulate_orphan_bases(self):
        before = self.refresh()['snapshot']
        self.scope['folders'] = []
        for position in range(4):
            self.fixture.attachment_source.write_bytes(f'cache-change-{position}'.encode())
            with self.assertRaisesRegex(ValueError, 'no verified'):
                self.refresh()
            self.assertEqual(len(list((self.fixture.state / 'snapshots').iterdir())), 1)
            self.assertEqual(self.client._status()['snapshot'], before)
            self.assertFalse(self.client._fresh())

    def test_changed_source_retains_only_one_compressed_copy(self):
        original = self.fixture.profile_tree()
        snapshots = []
        for position in range(4):
            # A changed asset manifest also creates a new source snapshot.
            self.fixture.attachment_source.write_bytes(f'cache-change-{position}'.encode())
            snapshots.append(self.refresh()['snapshot'])
            copies = list((self.fixture.state / 'snapshots').glob('*/source.hxd'))
            self.assertEqual(len(copies), 1)
            self.assertEqual({p.name for p in (self.fixture.state / 'views').iterdir()},
                             {snapshots[-1]})
        self.assertEqual(self.fixture.source_cache.read_bytes(),
                         (copies[0]).read_bytes())
        self.assertEqual(set(original), set(self.fixture.profile_tree()))

    def test_failed_upgrade_refresh_restores_prior_base_without_loading_old_schema(self):
        self.refresh()
        _, index = self.client._load()
        previous_base = (self.fixture.state / 'CURRENT').read_text()
        view_index = self.fixture.state / 'views' / index['snapshot'] / 'index.json'
        old_index = json.loads(view_index.read_text())
        old_index['schema_version'] = service.core.VERSION - 2
        view_index.chmod(0o600)
        view_index.write_text(json.dumps(old_index))
        view_index.chmod(0o400)
        self.fixture.attachment_source.write_bytes(b'changed-source-before-failed-upgrade')
        with patch.object(sys.modules['osa_scope'], 'read_scope',
                          side_effect=ValueError('scope unavailable during upgrade')):
            with self.assertRaisesRegex(ValueError, '^scope unavailable during upgrade$'):
                self.refresh()
        self.assertEqual((self.fixture.state / 'CURRENT').read_text(), previous_base)
        self.assertEqual({p.name for p in (self.fixture.state / 'snapshots').iterdir()},
                         {previous_base.strip()})

    def test_cycles_conflicts_and_deleted_descendants_not_allowed(self):
        folders = self.scope['folders'] + [dict(id='cycle', parent_id='cycle', name='循环'),
                                          dict(id='badroot', parent_id='deleted', name='删除下', is_sent=True)]
        tree = service.folder_tree(folders)
        self.assertNotIn('cycle', tree)
        self.assertNotIn('trashchild', tree)
        self.assertNotIn('badroot', tree)
        folders.append(dict(id='deep', parent_id='sent', name='冲突'))
        self.assertNotIn('deep', service.folder_tree(folders))


if __name__ == '__main__':
    unittest.main()
