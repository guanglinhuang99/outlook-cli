import gzip
import io
from xml.sax.saxutils import escape
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).parent))
import osa_scope  # noqa: E402


def folder_view(identifier, parent, display, type_value):
    return f"""
      <View>
        <ImmId>{identifier}</ImmId>
        <ImmParentId>{parent}</ImmParentId>
        <DisplayName>{display}</DisplayName>
        <Type>{type_value}</Type>
      </View>
    """


def message_header(entry, message_id, sent, folder):
    return f"""
      <FlattenedMessage>
        <MessageHeader>
          <ImmId>{entry}</ImmId>
          <InternetMessageId>{escape(message_id)}</InternetMessageId>
          <SentTime>{sent}</SentTime>
          <ViewId><OutlookId></OutlookId><ImmId>{folder}</ImmId><Type>1</Type></ViewId>
          <Body><Html>this body must not be retained</Html></Body>
        </MessageHeader>
        <Body><Html>large body placeholder</Html></Body>
      </FlattenedMessage>
    """


class OsaScopeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        root = Path(self.temporary.name)
        self.profile = root / "profile"
        self.log_root = root / "Osa"
        self.log_dir = self.log_root / "OutlookServiceApiLogs_test"
        self.log_dir.mkdir(parents=True)
        self.profile.mkdir()

    def tearDown(self):
        self.temporary.cleanup()

    def write_log(self, name, xml=None, raw=None):
        path = self.log_dir / name
        if raw is None:
            raw = xml.encode("utf-8")
        with gzip.open(path, "wb") as handle:
            handle.write(raw)
        return path

    def folder_log(self, log_time="2026-10-20T02:00:00Z"):
        return (
            f"<Response><LogTime>{log_time}</LogTime>"
            + folder_view("INBOX", "ROOT", "pii:inbox", "Inbox (1)")
            + folder_view("CHILD", "INBOX", "客户规则", "NotSpecial (0)")
            + folder_view("SENT", "ROOT", "Sent", "SentItems (3)")
            + folder_view("DELETED", "ROOT", "Deleted", "DeletedItems (4)")
            + "</Response>"
        )

    def test_exact_entry_id_maps_to_child_and_preserves_folder_flags(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-1", "<one@example.test>", "2026-10-20T01:59:00Z", "CHILD")
            + "</Response>",
        )

        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_one", "message_id": "<one@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["entry-1"]}],
            log_root=self.log_root,
        )
        folders = {row["id"]: row for row in result["folders"]}
        self.assertTrue(folders["INBOX"]["is_inbox"])
        self.assertFalse(folders["CHILD"]["is_inbox"])
        self.assertTrue(folders["CHILD"]["evidence"]["inbox_subtree"])
        self.assertTrue(folders["SENT"]["is_sent"])
        self.assertTrue(folders["DELETED"]["is_deleted"])
        self.assertEqual(folders["INBOX"]["name"], "收件箱")
        self.assertEqual(folders["SENT"]["name"], "已发送")
        self.assertEqual(folders["DELETED"]["name"], "已删除")
        self.assertEqual(folders["INBOX"]["evidence"]["observation_count"], 1)
        self.assertEqual(folders["CHILD"]["evidence"]["name_status"], "observed")
        self.assertEqual(result["messages"]["hx_one"]["folder_id"], "CHILD")
        self.assertEqual(
            result["messages"]["hx_one"]["folder_status"],
            "verified_latest_observation",
        )
        self.assertEqual(
            result["messages"]["hx_one"]["evidence"]["match_method"], "exact_entry_id"
        )
        self.assertFalse(result["messages"]["hx_one"]["evidence"]["live"])
        self.assertFalse(result["evidence"]["body_read"])

    def test_pre_epoch_timestamps_are_valid_and_map_membership(self):
        self.assertIsNone(osa_scope._parse_time("not-a-date"))
        self.assertEqual(
            osa_scope._parse_time("1969-12-31T23:59:00Z"), -60.0
        )
        self.assertEqual(osa_scope._item_time({"sent_at": "1969-12-31T23:58:00Z"}), -120.0)
        self.write_log(
            "x.GetFolder.1.res.xmlgz",
            self.folder_log("1969-12-31T23:59:00Z"),
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>1969-12-31T23:59:30Z</LogTime>"
            + message_header(
                "ENTRY-PRE-EPOCH",
                "<pre-epoch@example.test>",
                "1969-12-31T23:58:00Z",
                "CHILD",
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{
                "id": "hx_pre_epoch",
                "message_id": "<pre-epoch@example.test>",
                "sent_unix": -120,
                "entry_ids": ["ENTRY-PRE-EPOCH"],
            }],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_pre_epoch"]["folder_id"], "CHILD")

    def test_unix_epoch_zero_is_a_valid_timestamp(self):
        self.assertEqual(osa_scope._parse_time("1970-01-01T00:00:00Z"), 0.0)
        self.write_log(
            "x.GetFolder.1.res.xmlgz",
            self.folder_log("1970-01-01T00:00:00Z"),
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>1970-01-01T00:00:01Z</LogTime>"
            + message_header(
                "ENTRY-EPOCH",
                "<epoch@example.test>",
                "1970-01-01T00:00:00Z",
                "CHILD",
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{
                "id": "hx_epoch",
                "message_id": "<epoch@example.test>",
                "sent_unix": 0,
                "entry_ids": ["ENTRY-EPOCH"],
            }],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_epoch"]["folder_id"], "CHILD")

    def test_exact_entry_id_conflicts_are_excluded_without_mid_fallback(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-CONFLICT", "<different@example.test>", "2026-10-20T01:59:00Z", "CHILD")
            + message_header("", "<expected@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_conflict", "message_id": "<expected@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-CONFLICT"]}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_conflict", result["messages"])
        self.assertEqual(result["evidence"]["messages_exact_identity_conflict"], 1)

    def test_exact_entry_id_time_conflict_is_excluded(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-TIME", "<time@example.test>", "2026-10-20T02:01:00Z", "CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_time", "message_id": "<time@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-TIME"]}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_time", result["messages"])
        self.assertEqual(result["evidence"]["messages_exact_time_conflict"], 1)

    def test_partial_exact_entry_id_hit_is_excluded_without_mid_fallback(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-FOUND", "<partial@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_partial", "message_id": "<partial@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-FOUND", "ENTRY-MISSING"]}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_partial", result["messages"])
        self.assertEqual(result["evidence"]["messages_exact_identity_conflict"], 1)

    def test_known_entry_id_does_not_use_mid_fallback_and_duplicate_mid_is_rejected(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-OTHER", "<known@example.test>", "2026-10-20T01:59:00Z", "CHILD")
            + message_header("", "<duplicate@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [
                {"id": "hx_known", "message_id": "<known@example.test>",
                 "sent_unix": 1792461540, "entry_ids": ["ENTRY-MISSING"]},
                {"id": "hx_duplicate-a", "message_id": "<duplicate@example.test>",
                 "sent_unix": 1792461540},
                {"id": "hx_duplicate-b", "message_id": "<duplicate@example.test>",
                 "sent_unix": 1792461540},
            ],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_known", result["messages"])
        self.assertNotIn("hx_duplicate-a", result["messages"])
        self.assertNotIn("hx_duplicate-b", result["messages"])
        self.assertEqual(result["evidence"]["messages_fallback_input_duplicate"], 2)

    def test_missing_cache_row_id_does_not_use_internet_message_id_as_key(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("", "<no-row-id@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"message_id": "<no-row-id@example.test>", "sent_unix": 1792461540}],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"], {})
        self.assertEqual(result["evidence"]["messages_unknown"], 1)

    def test_ambiguous_and_untimestamped_catalog_nodes_are_not_returned(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetFolder.2.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:00:00Z</LogTime>"
            + folder_view("CHILD", "SENT", "other", "NotSpecial (0)")
            + "</Response>",
        )
        self.write_log(
            "x.GetFolder.3.res.xmlgz",
            "<Response><LogTime>invalid</LogTime>"
            + folder_view("NO-TIME", "INBOX", "no time", "NotSpecial (0)")
            + "</Response>",
        )
        result = osa_scope.read_scope(self.profile, [], log_root=self.log_root)
        ids = {row["id"] for row in result["folders"]}
        self.assertNotIn("CHILD", ids)
        self.assertNotIn("NO-TIME", ids)
        self.assertEqual(result["evidence"]["ambiguous_folder_nodes"], 1)
        self.assertEqual(result["evidence"]["no_valid_folder_time_nodes"], 1)

    def test_fallback_requires_time_and_uses_latest_observed_folder(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        message = "<Response><LogTime>{}</LogTime>{}</Response>"
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            message.format(
                "2026-10-20T02:01:00Z",
                message_header("", "<two@example.test>", "2026-10-20T01:59:00Z", "INBOX"),
            ),
        )
        self.write_log(
            "x.GetMessage.2.res.xmlgz",
            message.format(
                "2026-10-20T02:02:00Z",
                message_header("", "<two@example.test>", "2026-10-20T01:59:00.500000Z", "SENT"),
            ),
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_two", "message_id": "<two@example.test>", "sent_unix": 1792461540}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_two"]
        self.assertEqual(row["folder_id"], "SENT")
        self.assertEqual(row["folder_status"], "verified_latest_observation")
        self.assertEqual(row["evidence"]["match_method"], "internet_message_id_sent_time")
        self.assertEqual(row["evidence"]["scope"], "sent")
        self.assertTrue(row["evidence"]["latest_observed"])
        self.assertFalse(row["evidence"]["live"])

    def test_sent_descendant_scope_evidence_is_sent(self):
        self.write_log(
            "x.GetFolder.1.res.xmlgz",
            self.folder_log().replace(
                "</Response>",
                folder_view("SENT-CHILD", "SENT", "项目", "NotSpecial (0)") + "</Response>",
            ),
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-SENT-CHILD", "<sent-child@example.test>",
                             "2026-10-20T01:59:00Z", "SENT-CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_sent_child", "message_id": "<sent-child@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-SENT-CHILD"]}],
            log_root=self.log_root,
        )
        folders = {row["id"]: row for row in result["folders"]}
        self.assertTrue(folders["SENT-CHILD"]["evidence"]["sent_subtree"])
        self.assertFalse(folders["SENT-CHILD"]["evidence"]["deleted_subtree"])
        self.assertEqual(result["messages"]["hx_sent_child"]["evidence"]["scope"], "sent")

    def test_deleted_descendant_scope_evidence_is_deleted(self):
        self.write_log(
            "x.GetFolder.1.res.xmlgz",
            self.folder_log().replace(
                "</Response>",
                folder_view("DELETED-CHILD", "DELETED", "旧邮件", "NotSpecial (0)") + "</Response>",
            ),
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-DELETED-CHILD", "<deleted-child@example.test>",
                             "2026-10-20T01:59:00Z", "DELETED-CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_deleted_child", "message_id": "<deleted-child@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-DELETED-CHILD"]}],
            log_root=self.log_root,
        )
        folders = {row["id"]: row for row in result["folders"]}
        self.assertTrue(folders["DELETED-CHILD"]["evidence"]["deleted_subtree"])
        self.assertFalse(folders["DELETED-CHILD"]["evidence"]["sent_subtree"])
        self.assertEqual(result["messages"]["hx_deleted_child"]["evidence"]["scope"], "deleted")

    def test_fallback_rejects_nonunique_internet_message_id(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-A", "<same@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + message_header("ENTRY-B", "<same@example.test>", "2026-10-20T01:59:00Z", "CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_same", "message_id": "<same@example.test>", "sent_unix": 1792461540}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_same", result["messages"])
        self.assertEqual(result["evidence"]["messages_fallback_identity_conflict"], 1)

    def test_fallback_rejects_mixed_identified_and_unidentified_log_rows(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-A", "<mixed@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + message_header("", "<mixed@example.test>", "2026-10-20T01:59:00Z", "CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_mixed", "message_id": "<mixed@example.test>",
              "sent_unix": 1792461540}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_mixed", result["messages"])
        self.assertEqual(result["evidence"]["messages_fallback_identity_conflict"], 1)

    def test_same_latest_log_time_with_different_folders_is_ambiguous(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-TIE", "<tie@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + "</Response>",
        )
        self.write_log(
            "x.GetMessage.2.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-TIE", "<tie@example.test>", "2026-10-20T01:59:00Z", "CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_tie", "message_id": "<tie@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-TIE"]}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_tie", result["messages"])
        self.assertEqual(result["evidence"]["messages_ambiguous"], 1)

    def test_bad_xml_fails_closed_and_missing_root_stays_empty(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log("x.GetFolder.2.res.xmlgz", raw=b"not gzip")
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<!DOCTYPE Response SYSTEM 'https://example.invalid/evil.dtd'>"
            "<Response></Response>",
        )
        with self.assertRaisesRegex(osa_scope.OsaScopeError, "unreadable or invalid"):
            osa_scope.read_scope(self.profile, [], log_root=self.log_root)

        missing = osa_scope.read_scope(self.profile, [], log_root=self.profile / "missing")
        self.assertEqual(missing["folders"], [])
        self.assertEqual(missing["messages"], {})
        self.assertTrue(missing["evidence"]["severe_missing"])

        symlink_root = self.profile / "Osa-link"
        symlink_root.symlink_to(self.log_root, target_is_directory=True)
        with self.assertRaisesRegex(osa_scope.OsaScopeError, "symbolic link"):
            osa_scope.read_scope(self.profile, [], log_root=symlink_root)

    def test_corrupt_deflate_with_valid_gzip_header_fails_closed(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        corrupted = bytearray(gzip.compress(
            b"<Response><LogTime>2026-10-20T02:00:00Z</LogTime></Response>"
        ))
        corrupted[10] ^= 0xFF
        (self.log_dir / "x.GetFolder.2.res.xmlgz").write_bytes(corrupted)
        with self.assertRaisesRegex(osa_scope.OsaScopeError, "unreadable or invalid"):
            osa_scope.read_scope(self.profile, [], log_root=self.log_root)

    def test_existing_empty_osa_root_is_missing(self):
        result = osa_scope.read_scope(self.profile, [], log_root=self.log_root)
        self.assertEqual(result["folders"], [])
        self.assertEqual(result["messages"], {})
        self.assertEqual(result["evidence"]["status"], "missing")
        self.assertTrue(result["evidence"]["severe_missing"])

    def test_utf16_doctype_and_entity_are_rejected(self):
        xml = (
            "<?xml version='1.0' encoding='UTF-16'?>"
            "<!DOCTYPE Response [<!ENTITY injected 'unsafe'>]>"
            "<Response><LogTime>2026-10-20T02:00:00Z</LogTime></Response>"
        )
        for suffix, encoded in (
            ("le", b"\xff\xfe" + xml.encode("utf-16-le")),
            ("be", b"\xfe\xff" + xml.encode("utf-16-be")),
        ):
            with self.subTest(encoding=suffix):
                self.write_log(f"x.GetFolder.utf16-{suffix}.res.xmlgz", raw=encoded)
                with self.assertRaisesRegex(osa_scope.OsaScopeError, "unsafe XML"):
                    osa_scope.read_scope(self.profile, [], log_root=self.log_root)

    def test_utf16_guard_rejects_declaration_split_across_raw_chunks(self):
        xml = "<!DOCTYPE Response [<!ENTITY injected 'unsafe'>]><Response/>"
        for encoding, prefix in (("utf-16-le", b"\xff\xfe"), ("utf-16-be", b"\xfe\xff")):
            with self.subTest(encoding=encoding):
                source = io.BytesIO(prefix + xml.encode(encoding))
                guarded = osa_scope._GuardedReader(source, 4096)
                with self.assertRaises(osa_scope._UnsafeXML):
                    for _ in range(1024):
                        if not guarded.read(3):
                            self.fail("guard reached EOF without rejecting the declaration")

    def test_newest_unreadable_log_does_not_reuse_old_inbox_membership(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-STALE", "<stale@example.test>", "2026-10-20T01:59:00Z", "INBOX")
            + "</Response>",
        )
        self.write_log("x.GetMessage.2.res.xmlgz", raw=b"truncated newest Osa response")
        with self.assertRaisesRegex(osa_scope.OsaScopeError, "unreadable or invalid"):
            osa_scope.read_scope(
                self.profile,
                [{"id": "hx_stale", "message_id": "<stale@example.test>",
                  "sent_unix": 1792461540}],
                log_root=self.log_root,
            )

    def test_changing_log_fails_closed_instead_of_publishing_partial_membership(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        with patch.object(osa_scope, "_same_stat", return_value=False):
            with self.assertRaisesRegex(osa_scope.OsaScopeError, "changed during read"):
                osa_scope.read_scope(self.profile, [], log_root=self.log_root)

    def test_log_set_change_at_scan_end_fails_closed(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        original_log_files = osa_scope._log_files
        calls = 0

        def relist_with_new_log(root, evidence):
            nonlocal calls
            calls += 1
            if calls == 2:
                self.write_log("x.GetMessage.2.res.xmlgz", "<Response></Response>")
            return original_log_files(root, evidence)

        with patch.object(osa_scope, "_log_files", side_effect=relist_with_new_log):
            with self.assertRaisesRegex(osa_scope.OsaScopeError, "changed during scan"):
                osa_scope.read_scope(self.profile, [], log_root=self.log_root)
        self.assertEqual(calls, 2)


if __name__ == "__main__":
    unittest.main()
