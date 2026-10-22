import copy
from datetime import datetime, timezone
import hashlib
import sys
import types
import unittest
from unittest.mock import patch


sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
import mcp_server  # noqa: E402
import outlook_service as service  # noqa: E402
import test_outlook_local as fixtures  # noqa: E402


class DigestTests(unittest.TestCase):
    """Regression coverage for the compact, incremental digest contract."""

    def setUp(self):
        self.fixture = fixtures.OutlookLocalTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        _, _, self.base = self.fixture.snapshot()
        self.ids = [row["id"] for row in self.base["items"]]
        self.scope = {
            "folders": [
                dict(id="inbox", parent_id="root", name="收件箱", is_inbox=True),
                dict(id="child", parent_id="inbox", name="研究", is_inbox=False),
                dict(id="deep", parent_id="child", name="二级", is_inbox=False),
                dict(id="sent", parent_id="root", name="已发送", is_inbox=False, is_sent=True),
                dict(id="deleted", parent_id="root", name="已删除", is_inbox=False, is_deleted=True),
            ],
            "messages": {
                self.ids[0]: dict(folder_id="inbox", folder_status="verified_latest_observation"),
                self.ids[1]: dict(folder_id="deep", folder_status="verified_latest_observation"),
                self.ids[2]: dict(folder_id="deleted", folder_status="verified_latest_observation"),
            },
            "fingerprint": "digest-fixture",
            "evidence": {"source": "synthetic"},
        }
        for row in self.base["items"]:
            self.scope["messages"][row["id"]].update(
                received_unix=row["sent_unix"] + 60,
                received_time_status="verified",
                received_time_method="osa_last_delivery_time",
            )
        module = types.ModuleType("osa_scope")
        module.read_scope = lambda profile, items, **kwargs: copy.deepcopy(self.scope)
        self.scope_patch = patch.dict(sys.modules, {"osa_scope": module})
        self.scope_patch.start()
        self.addCleanup(self.scope_patch.stop)
        self.client = service.MailService(self.fixture.state, self.fixture.profile)

    def refresh(self, **arguments):
        return self.client.call("outlook_local_refresh", arguments)

    @staticmethod
    def _digest(text):
        return hashlib.sha256(text.encode("utf-8")).hexdigest()

    def _body_message(self, row, *, text=None, status="cached_html", full_hash=None,
                      text_chars=None, truncated=False, diagnosis=None):
        if text_chars is None and text is not None:
            text_chars = len(text)
        body = {
            "status": status,
            "diagnosis": diagnosis or {"code": "synthetic"},
            "text_chars": text_chars,
            "truncated": truncated,
            "text_sha256": full_hash,
        }
        if text is not None:
            body["text"] = text
        return {**row, "attachments": [], "body": body}

    def _real_page(self, **arguments):
        self.refresh()
        return self.client.call("outlook_local_search", arguments)

    def test_exact_duplicate_body_is_emitted_once_while_mail_rows_stay_distinct(self):
        page = self._real_page(limit=25)
        rows = {row["id"]: row for row in page["items"]}
        same_text = "同一正文"
        same_hash = self._digest(same_text)

        def read(_directory, _index, identifier, _maximum):
            return self._body_message(rows[identifier], text=same_text, full_hash=same_hash)

        before = self.fixture.profile_tree()
        with patch.object(self.client, "_read", side_effect=read):
            result = self.client.call("outlook_local_digest", {"limit": 25})
        self.assertEqual(self.fixture.profile_tree(), before)
        self.assertEqual(set(result["page_ids"]), set(rows))
        self.assertEqual(result["snapshot"], page["snapshot"])
        self.assertEqual(result["total_matched"], 2)
        self.assertIsNone(result["next_cursor"])
        self.assertEqual(result["date_basis"], "received_at")
        self.assertEqual(len(result["items"]), 2)
        self.assertEqual(len(result["bodies"]), 1)
        content_ids = {item["body"]["content_id"] for item in result["items"]}
        self.assertEqual(len(content_ids), 1)
        self.assertEqual(list(result["bodies"].values())[0]["text"], same_text)
        self.assertNotEqual(result["items"][0]["id"], result["items"][1]["id"])
        self.assertEqual({item["folder_id"] for item in result["items"]}, {"inbox", "deep"})

        known = next(iter(content_ids))
        with patch.object(self.client, "_read", side_effect=read):
            suppressed = self.client.call("outlook_local_digest", {
                "known_bodies": [known], "limit": 25,
            })
        self.assertEqual(suppressed["bodies"], {})
        self.assertTrue(all(item["body"]["content_id"] == known
                            for item in suppressed["items"]))

    def test_same_returned_prefix_with_different_verified_full_hashes_is_not_merged(self):
        page = self._real_page(limit=25)
        rows = {row["id"]: row for row in page["items"]}
        prefix = "相同前缀"
        hashes = {identifier: self._digest(f"完整正文-{identifier}") for identifier in rows}

        def read(_directory, _index, identifier, _maximum):
            return self._body_message(rows[identifier], text=prefix, full_hash=hashes[identifier],
                                      text_chars=99, truncated=True)

        with patch.object(self.client, "_read", side_effect=read):
            result = self.client.call("outlook_local_digest", {"limit": 25, "body_max_chars": 4})
        self.assertEqual(len(result["bodies"]), 2)
        self.assertEqual({body["text"] for body in result["bodies"].values()}, {prefix})
        self.assertEqual(len({item["body"]["content_id"] for item in result["items"]}), 2)

    def test_known_versions_suppress_rows_and_body_changes_invalidate_versions(self):
        page = self._real_page(limit=25)
        rows = {row["id"]: row for row in page["items"]}
        texts = {identifier: f"正文-{position}" for position, identifier in enumerate(rows)}

        def read_with_text(_directory, _index, identifier, _maximum):
            text = texts[identifier]
            return self._body_message(rows[identifier], text=text, full_hash=self._digest(text))

        with patch.object(self.client, "_read", side_effect=read_with_text):
            first = self.client.call("outlook_local_digest", {"limit": 25})
        versions = [item["version"] for item in first["items"]]
        body_ids = list(first["bodies"])
        with patch.object(self.client, "_read", side_effect=read_with_text):
            unchanged = self.client.call("outlook_local_digest", {
                "limit": 25, "known_versions": versions, "known_bodies": body_ids,
            })
        self.assertEqual(unchanged["items"], [])
        self.assertEqual(
            {row["id"] for row in unchanged["unchanged"]},
            set(rows),
        )
        self.assertEqual(unchanged["bodies"], {})

        changed_id = next(iter(rows))
        texts[changed_id] = "正文-已修改"
        with patch.object(self.client, "_read", side_effect=read_with_text):
            changed = self.client.call("outlook_local_digest", {
                "limit": 25, "known_versions": versions, "known_bodies": body_ids,
            })
        self.assertEqual([item["id"] for item in changed["items"]], [changed_id])
        self.assertEqual(changed["unchanged"], [
            {"id": identifier, "version": version}
            for identifier, version in zip(rows, versions)
            if identifier != changed_id
        ])
        self.assertNotEqual(changed["items"][0]["version"], versions[0])
        self.assertEqual(len(changed["bodies"]), 1)

    def test_body_cap_changes_version_and_content_id(self):
        page = self._real_page(limit=25)
        rows = {row["id"]: row for row in page["items"]}
        full_text = "abcdefghijklmnopqrstuvwxyz"
        full_hash = self._digest(full_text)

        def read(_directory, _index, identifier, maximum):
            text = full_text[:maximum]
            return self._body_message(rows[identifier], text=text, full_hash=full_hash,
                                      text_chars=len(full_text), truncated=len(text) < len(full_text))

        with patch.object(self.client, "_read", side_effect=read):
            short = self.client.call("outlook_local_digest", {"limit": 25, "body_max_chars": 5})
        known_versions = [item["version"] for item in short["items"]]
        known_bodies = list(short["bodies"])
        with patch.object(self.client, "_read", side_effect=read):
            longer = self.client.call("outlook_local_digest", {
                "limit": 25, "body_max_chars": 8,
                "known_versions": known_versions, "known_bodies": known_bodies,
            })
        self.assertEqual(len(longer["items"]), 2)
        self.assertEqual(longer["unchanged"], [])
        self.assertEqual(len(longer["bodies"]), 1)
        self.assertNotEqual(
            {item["body"]["content_id"] for item in short["items"]},
            {item["body"]["content_id"] for item in longer["items"]},
        )

    def test_missing_body_becoming_cached_invalidates_incremental_version(self):
        page = self._real_page(folder_id="inbox", include_subfolders=False, limit=25)
        row = page["items"][0]
        state = {"cached": False}
        text = "正文后来进入本地缓存"

        def read(_directory, _index, identifier, _maximum):
            if not state["cached"]:
                return self._body_message(
                    row, text=None, status="not_cached",
                    diagnosis={"code": "no_local_cache"},
                )
            return self._body_message(row, text=text, full_hash=self._digest(text))

        arguments = {"folder_id": "inbox", "include_subfolders": False, "limit": 25}
        with patch.object(self.client, "_read", side_effect=read):
            missing = self.client.call("outlook_local_digest", arguments)
        old_version = missing["items"][0]["version"]
        state["cached"] = True
        with patch.object(self.client, "_read", side_effect=read):
            cached = self.client.call("outlook_local_digest", {
                **arguments, "known_versions": [old_version],
            })
        self.assertEqual([item["id"] for item in cached["items"]], [row["id"]])
        self.assertEqual(cached["unchanged"], [])
        self.assertNotEqual(cached["items"][0]["version"], old_version)
        self.assertEqual(next(iter(cached["bodies"].values()))["text"], text)

    def test_recipient_and_folder_metadata_changes_invalidate_versions(self):
        page = self._real_page(folder_id="inbox", include_subfolders=False, limit=25)
        row = page["items"][0]
        state = {
            "participants": {"to": [{"name": "甲", "address": "a@example.test"}]},
            "participants_status": {"to": "verified"},
            "folder_id": "inbox",
            "folder_path": "收件箱",
        }
        text = "稳定正文"

        def read(_directory, _index, identifier, _maximum):
            message = self._body_message(row, text=text, full_hash=self._digest(text))
            return {**message, **state}

        arguments = {"folder_id": "inbox", "include_subfolders": False, "limit": 25}
        with patch.object(self.client, "_read", side_effect=read):
            first = self.client.call("outlook_local_digest", arguments)
        first_item = first["items"][0]

        state["participants"] = {"to": [{"name": "乙", "address": "b@example.test"}]}
        with patch.object(self.client, "_read", side_effect=read):
            recipient_changed = self.client.call("outlook_local_digest", {
                **arguments, "known_versions": [first_item["version"]],
            })
        self.assertEqual([item["id"] for item in recipient_changed["items"]], [row["id"]])
        self.assertNotEqual(recipient_changed["items"][0]["version"], first_item["version"])
        self.assertEqual(recipient_changed["items"][0]["participants"], state["participants"])

        state["folder_id"] = "sent"
        state["folder_path"] = "已发送"
        with patch.object(self.client, "_read", side_effect=read):
            folder_changed = self.client.call("outlook_local_digest", {
                **arguments, "known_versions": [recipient_changed["items"][0]["version"]],
            })
        self.assertEqual(folder_changed["items"][0]["folder_id"], "sent")
        self.assertNotEqual(
            folder_changed["items"][0]["version"], recipient_changed["items"][0]["version"]
        )

    def test_snapshot_only_change_does_not_invalidate_compact_version(self):
        page = self._real_page(limit=25)
        rows = {row["id"]: row for row in page["items"]}
        text = "同样的正文和元数据"

        def read(_directory, _index, identifier, _maximum):
            return self._body_message(rows[identifier], text=text, full_hash=self._digest(text))

        with patch.object(self.client, "_read", side_effect=read):
            first = self.client.call("outlook_local_digest", {"limit": 25})
        first_snapshot = first["snapshot"]
        versions = [item["version"] for item in first["items"]]
        self.scope["fingerprint"] = "digest-snapshot-only-change"
        refreshed = self.refresh()
        self.assertNotEqual(refreshed["snapshot"], first_snapshot)
        with patch.object(self.client, "_read", side_effect=read):
            second = self.client.call("outlook_local_digest", {
                "limit": 25, "known_versions": versions,
            })
        self.assertEqual(second["items"], [])
        self.assertEqual({row["version"] for row in second["unchanged"]}, set(versions))

    def test_moved_to_deleted_is_excluded_and_old_cursor_is_revoked(self):
        first = self.client.call("outlook_local_digest", {"limit": 1})
        self.assertIsNotNone(first["next_cursor"])
        moved_id = first["page_ids"][0]
        old_cursor = first["next_cursor"]
        self.scope["messages"][moved_id]["folder_id"] = "deleted"
        self.scope["fingerprint"] = "digest-moved-to-deleted"
        self.refresh()

        current = self.client.call("outlook_local_digest", {"limit": 25})
        self.assertNotIn(moved_id, current["page_ids"])
        self.assertEqual(current["total_matched"], 1)
        with self.assertRaisesRegex(ValueError, "folder access changed"):
            self.client.call("outlook_local_digest", {"limit": 1, "cursor": old_cursor})

    def test_missing_empty_and_row_read_failures_remain_distinguishable(self):
        page = self._real_page(limit=25)
        rows = {row["id"]: row for row in page["items"]}
        empty_id, missing_id = rows
        error_row = {**rows[empty_id], "id": "synthetic-error"}
        page = {**page, "items": [rows[empty_id], rows[missing_id], error_row], "total_matched": 3}

        def read(_directory, _index, identifier, _maximum):
            if identifier == empty_id:
                return self._body_message(rows[identifier], text="", full_hash=self._digest(""),
                                          text_chars=0, truncated=False)
            if identifier == missing_id:
                return self._body_message(rows[identifier], text=None, status="not_cached",
                                          diagnosis={"code": "no_local_cache"})
            raise OSError("body read failed")

        with patch.object(self.client, "_search", return_value=page), \
                patch.object(self.client, "_read", side_effect=read):
            result = self.client.call("outlook_local_digest", {"limit": 25})
        self.assertEqual(result["page_ids"], [empty_id, missing_id, "synthetic-error"])
        self.assertEqual(len(result["items"]), 2)
        empty = next(item for item in result["items"] if item["id"] == empty_id)
        missing = next(item for item in result["items"] if item["id"] == missing_id)
        self.assertIn("content_id", empty["body"])
        self.assertNotIn("content_id", missing["body"])
        self.assertEqual(next(iter(result["bodies"].values()))["text"], "")
        self.assertEqual(result["errors"], [{"id": "synthetic-error", "error": "body read failed"}])

    def test_received_time_folder_scope_and_cursor_are_preserved_read_only(self):
        received = datetime(2001, 1, 1, tzinfo=timezone.utc).timestamp()
        for identifier in self.ids[:2]:
            self.scope["messages"][identifier]["received_unix"] = received
        self.scope["fingerprint"] = "delayed-receipt"
        before = self.fixture.profile_tree()
        self.refresh(from_date="2001-01-01")
        first = self.client.call("outlook_local_digest", {"limit": 1})
        second = self.client.call("outlook_local_digest", {
            "limit": 1, "cursor": first["next_cursor"],
        })
        self.assertEqual(first["date_basis"], "received_at")
        self.assertEqual(first["snapshot"], second["snapshot"])
        self.assertNotEqual(first["page_ids"], second["page_ids"])
        self.assertEqual(first["items"][0]["sent_at"][:4], "2000")
        self.assertTrue(first["items"][0]["received_at"].startswith("2001-01-01"))
        self.assertIsNone(second["next_cursor"])

        inbox = self.client.call("outlook_local_digest", {
            "folder_id": "inbox", "include_subfolders": False, "limit": 25,
        })
        self.assertEqual(inbox["total_matched"], 1)
        self.assertEqual(inbox["items"][0]["folder_id"], "inbox")
        with self.assertRaises(ValueError):
            self.client.call("outlook_local_digest", {"folder_id": "deleted"})
        self.assertEqual(self.fixture.profile_tree(), before)

    def test_digest_schema_bounds_known_state_and_compact_limits(self):
        defaults = mcp_server.validate_tool_arguments("outlook_local_digest", {})
        self.assertEqual(defaults["limit"], 25)
        self.assertEqual(defaults["body_max_chars"], 12_000)
        for arguments in (
            {"limit": 26}, {"body_max_chars": 20_001},
            {"known_versions": ["a" * 64] * 5001},
            {"known_bodies": ["a" * 64] * 5001},
        ):
            with self.subTest(arguments=arguments), self.assertRaises(mcp_server.ValidationError):
                mcp_server.validate_tool_arguments("outlook_local_digest", arguments)


if __name__ == "__main__":
    unittest.main()
