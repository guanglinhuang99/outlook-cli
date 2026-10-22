import copy
import os
import shutil
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).parent))
import outlook_service as service
import test_outlook_service as service_tests


class IndexCacheTests(unittest.TestCase):
    """Focused checks for the bounded, security-validated index cache."""

    def setUp(self):
        helper = service_tests.ServiceTests("runTest")
        helper.setUp()
        self.helper = helper
        self.addCleanup(helper.doCleanups)
        self.client = helper.client
        helper.refresh()
        _, index = self.client._load()
        self.snapshot = index["snapshot"]
        self.view = self.client.state / "views" / self.snapshot
        self.index_path = self.view / "index.json"

    def test_cache_reuses_parsed_index_and_checksum(self):
        self.client._index_cache.clear()
        with patch.object(service.json, "loads", wraps=service.json.loads) as loads, \
                patch.object(service.core, "sha", wraps=service.core.sha) as digest:
            self.client._load()
            self.client._load()
        self.assertEqual(loads.call_count, 1)
        self.assertEqual(digest.call_count, 1)

    def test_cache_fingerprint_rejects_same_size_tamper_with_restored_mtime(self):
        original = self.index_path.read_bytes()
        old_marker = b"first-catalog"
        new_marker = b"other-catalog"
        self.assertEqual(len(old_marker), len(new_marker))
        self.assertIn(old_marker, original)
        original_stat = self.index_path.stat()
        self.index_path.chmod(0o600)
        self.index_path.write_bytes(original.replace(old_marker, new_marker, 1))
        os.utime(self.index_path, ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns))
        self.index_path.chmod(0o400)

        with self.assertRaisesRegex(ValueError, "checksum"):
            self.client._load()

    def test_cache_fingerprint_reloads_replaced_inode(self):
        original = self.index_path.read_bytes()
        replacement = self.index_path.with_name(".replacement-index")
        replacement.write_bytes(original)
        replacement.chmod(0o400)
        replacement.replace(self.index_path)

        with patch.object(service.json, "loads", wraps=service.json.loads) as loads:
            _, index = self.client._load()
        self.assertEqual(index["snapshot"], self.snapshot)
        self.assertEqual(loads.call_count, 1)

    def test_default_load_returns_mutation_isolated_index(self):
        _, first = self.client._load()
        item_count = len(first["items"])
        first["items"].clear()
        first["summary"]["messages"] = -1

        _, second = self.client._load()
        self.assertEqual(len(second["items"]), item_count)
        self.assertGreaterEqual(second["summary"]["messages"], 0)

    def test_search_copies_only_the_public_result_not_the_whole_index(self):
        with patch.object(service, 'deepcopy', wraps=service.deepcopy) as copied:
            result = self.client.call('outlook_local_search', {'limit': 1})
        self.assertEqual(copied.call_count, 1)
        value = copied.call_args.args[0]
        self.assertEqual(len(value['items']), 1)
        self.assertIn('total_matched', value)
        self.assertNotIn('source_block_bytes', value)
        self.assertEqual(len(result['items']), 1)

    def test_all_public_responses_are_isolated_from_the_borrowed_index(self):
        identifier = self.helper.ids[0]
        self.helper.scope['messages'][identifier].update(
            participants={'to': [{'name': '测试姓名', 'address': 'person@example.test'}]},
            participants_status={'to': 'verified'})
        self.helper.refresh()
        _, borrowed = self.client._load(copy_index=False)
        expected = copy.deepcopy(borrowed)

        def corrupt(value):
            if isinstance(value, dict):
                for child in list(value.values()):
                    corrupt(child)
                value['client_mutation'] = True
            elif isinstance(value, list):
                for child in list(value):
                    corrupt(child)
                value.append('client_mutation')

        for name, args in [('status', {}), ('folders', {}), ('search', {'limit': 1}),
                           ('read', {'id': identifier}), ('attachments', {'id': identifier}),
                           ('messages_bundle', {'ids': [identifier]})]:
            with self.subTest(name=name):
                corrupt(self.client.call('outlook_local_' + name, args))
                self.assertEqual(self.client._load(copy_index=False)[1], expected)

    def test_cache_hit_rechecks_view_permissions_and_index_symlink(self):
        self.view.chmod(0o755)
        try:
            with self.assertRaisesRegex(ValueError, "private permissions"):
                self.client._load()
        finally:
            self.view.chmod(0o700)

        backup = self.index_path.with_name("index.json.backup")
        self.index_path.rename(backup)
        try:
            self.index_path.symlink_to(backup.name)
            with self.assertRaisesRegex(ValueError, "symbolic links"):
                self.client._load()
        finally:
            self.index_path.unlink(missing_ok=True)
            backup.rename(self.index_path)
            self.index_path.chmod(0o400)

    def test_cache_hit_rechecks_source_permissions_and_profile(self):
        directory, _ = self.client._load()
        directory.chmod(0o755)
        try:
            with self.assertRaisesRegex(ValueError, "private permissions"):
                self.client._load()
        finally:
            directory.chmod(0o700)

        other_profile = self.client.profile.parent / "profile-other"
        shutil.copytree(self.client.profile, other_profile)
        self.client.profile = other_profile
        with self.assertRaisesRegex(ValueError, "different Outlook profile"):
            self.client._load()

    def test_bounded_cache_keeps_pinned_source_scope_rules(self):
        old_snapshot = self.snapshot
        self.helper.scope["fingerprint"] = "second-catalog"
        new_snapshot = self.helper.refresh()["snapshot"]
        self.assertNotEqual(new_snapshot, old_snapshot)

        _, old_index = self.client._load(old_snapshot)
        _, new_index = self.client._load(new_snapshot)
        self.assertEqual(old_index["snapshot"], old_snapshot)
        self.assertEqual(new_index["snapshot"], new_snapshot)
        self.assertLessEqual(len(self.client._index_cache), service.INDEX_CACHE_LIMIT)

        _, pinned = self.client._current(old_snapshot)
        self.assertEqual(pinned["snapshot"], old_snapshot)


if __name__ == "__main__":
    unittest.main()
