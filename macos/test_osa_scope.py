import gzip
import hashlib
import io
import os
import sqlite3
import json
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


def participant(role, name, address):
    return (
        f"<{role}><DisplayName>{escape(name)}</DisplayName>"
        f"<EmailAddress d=\"raw\">{escape(address)}</EmailAddress>"
        f"<EmailAddressType>SMTP</EmailAddressType></{role}>"
    )


def message_data(entry, sender=(), to=(), cc=(), partial="0", truncated="0"):
    return (
        f"<MessageData><ImmId>{escape(entry)}</ImmId>"
        f"<IsPartial>{partial}</IsPartial><IsTruncated>{truncated}</IsTruncated>"
        + "".join(participant("Sender", name, address) for name, address in sender)
        + "".join(participant("To", name, address) for name, address in to)
        + "".join(participant("Cc", name, address) for name, address in cc)
        + "</MessageData>"
    )


def message_header(entry, message_id, sent, folder, received=None, received_ticks=None,
                   data=None, from_name=None, from_address=None):
    receipt = ""
    if received is not None:
        receipt += f"<LastDeliveryTime>{escape(received)}</LastDeliveryTime>"
    if received_ticks is not None:
        receipt += f"<ReceivedOrRenewTime>{escape(str(received_ticks))}</ReceivedOrRenewTime>"
    typed_from = ""
    if from_name is not None or from_address is not None:
        typed_from = (
            "<From><EmailAddress>"
            f"<Name>{escape(from_name or '')}</Name>"
            f"<Address>{escape(from_address or '')}</Address>"
            "</EmailAddress></From>"
        )
    return f"""
      <FlattenedMessage>
        <MessageHeader>
          <ImmId>{entry}</ImmId>
          <InternetMessageId>{escape(message_id)}</InternetMessageId>
          <SentTime>{sent}</SentTime>
          {receipt}
          {typed_from}
          <ViewId><OutlookId></OutlookId><ImmId>{folder}</ImmId><Type>1</Type></ViewId>
          <Body><Html>this body must not be retained</Html></Body>
        </MessageHeader>
        {data or ''}
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

    def test_typed_message_data_binds_roles_and_ignores_data_view_id(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        data = message_data(
            "ENTRY-PARTICIPANTS",
            sender=[("Sender", "sender@example.test")],
            to=[("To One", "to-one@example.test"), ("To Two", "to-two@example.test")],
            cc=[("Cc One", "cc-one@example.test")],
        ).replace(
            "</MessageData>",
            "<ViewId><ImmId>WRONG-VIEW-ID</ImmId></ViewId></MessageData>",
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-PARTICIPANTS", "<participants@example.test>",
                "2026-10-20T01:59:00Z", "CHILD", data=data,
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_participants", "message_id": "<participants@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-PARTICIPANTS"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_participants"]
        self.assertEqual(row["folder_id"], "CHILD")
        self.assertEqual(row["participants"], {
            "sender": [{"name": "Sender", "address": "sender@example.test"}],
            "to": [{"name": "To One", "address": "to-one@example.test"},
                   {"name": "To Two", "address": "to-two@example.test"}],
            "cc": [{"name": "Cc One", "address": "cc-one@example.test"}],
        })
        self.assertEqual(row["participants_status"], {
            "sender": "verified", "to": "verified", "cc": "verified",
        })
        self.assertEqual(row["participant_counts"], {
            "sender": {"count": 1, "status": "verified"},
            "to": {"count": 2, "status": "verified"},
            "cc": {"count": 1, "status": "verified"},
        })

    def test_participant_counts_include_redacted_elements_without_exposing_values(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        data = message_data(
            "ENTRY-REDACTED-COUNTS",
            to=[("pii:one", "pii:one"), ("<redacted>", "<redacted>")],
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-REDACTED-COUNTS", "<redacted-counts@example.test>",
                "2026-10-20T01:59:00Z", "CHILD", data=data,
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_redacted_counts", "message_id": "<redacted-counts@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-REDACTED-COUNTS"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_redacted_counts"]
        self.assertEqual(row["participant_counts"]["to"], {"count": 2, "status": "verified"})
        self.assertNotIn("pii:", repr(row["participant_counts"]))
        self.assertNotIn("redacted", repr(row["participant_counts"]).casefold())

    def test_participant_counts_explicit_zero_is_verified(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-COUNT-ZERO", "<count-zero@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                data=message_data("ENTRY-COUNT-ZERO"),
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_count_zero", "message_id": "<count-zero@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-COUNT-ZERO"]}],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_count_zero"]["participant_counts"], {
            "sender": {"count": 0, "status": "verified"},
            "to": {"count": 0, "status": "verified"},
            "cc": {"count": 0, "status": "verified"},
        })

    def test_participant_counts_missing_flags_fail_closed(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        data = message_data(
            "ENTRY-COUNT-MISSING-FLAGS", to=[("Observed", "observed@example.test")],
        ).replace("<IsPartial>0</IsPartial>", "").replace(
            "<IsTruncated>0</IsTruncated>", ""
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-COUNT-MISSING-FLAGS", "<count-missing-flags@example.test>",
                "2026-10-20T01:59:00Z", "CHILD", data=data,
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_count_missing_flags", "message_id": "<count-missing-flags@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-COUNT-MISSING-FLAGS"]}],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_count_missing_flags"]["participant_counts"], {
            role: {"count": None, "status": "partial"}
            for role in ("sender", "to", "cc")
        })

    def test_data_imm_id_mismatch_makes_all_participant_roles_unknown(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-MISMATCH", "<mismatch@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                data=message_data("OTHER-ENTRY", to=[("Typed", "to@example.test")]),
                from_name="Header Sender", from_address="header@example.test",
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_mismatch", "message_id": "<mismatch@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-MISMATCH"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_mismatch"]
        self.assertEqual(row["participants"], {"sender": None, "to": None, "cc": None})
        self.assertEqual(row["participants_status"], {
            "sender": "conflict", "to": "conflict", "cc": "conflict",
        })
        self.assertEqual(row["participant_counts"], {
            role: {"count": None, "status": "conflict"}
            for role in ("sender", "to", "cc")
        })

    def test_newer_partial_observation_does_not_backfill_older_complete_roles(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-PARTIAL", "<partial@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                data=message_data("ENTRY-PARTIAL", to=[("Older", "older@example.test")]),
            )
            + "</Response>",
        )
        self.write_log(
            "x.GetMessage.2.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:02:00Z</LogTime>"
            + message_header(
                "ENTRY-PARTIAL", "<partial@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                data=message_data("ENTRY-PARTIAL", to=[("New", "new@example.test")], partial="1"),
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_partial", "message_id": "<partial@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-PARTIAL"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_partial"]
        self.assertEqual(row["participants"], {
            "sender": None, "to": None, "cc": None,
        })
        self.assertEqual(row["participants_status"], {
            "sender": "partial", "to": "partial", "cc": "partial",
        })
        self.assertEqual(row["participant_counts"], {
            role: {"count": None, "status": "partial"}
            for role in ("sender", "to", "cc")
        })

    def test_newer_missing_observation_does_not_backfill_older_counts(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-COUNT-LATEST-MISSING", "<count-latest-missing@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                data=message_data("ENTRY-COUNT-LATEST-MISSING", to=[("Old", "old@example.test")]),
            )
            + "</Response>",
        )
        self.write_log(
            "x.GetMessage.2.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:02:00Z</LogTime>"
            + message_header(
                "ENTRY-COUNT-LATEST-MISSING", "<count-latest-missing@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_count_latest_missing", "message_id": "<count-latest-missing@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-COUNT-LATEST-MISSING"]}],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_count_latest_missing"]["participant_counts"], {
            role: {"count": None, "status": "missing"}
            for role in ("sender", "to", "cc")
        })

    def test_redacted_and_malformed_roles_are_never_indexed(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        data = (
            "<MessageData><ImmId>ENTRY-REDACTED</ImmId>"
            "<IsPartial>0</IsPartial><IsTruncated>0</IsTruncated>"
            + participant("Sender", "pii:sender", "pii:address")
            + "<To><DisplayName>Missing Address</DisplayName></To>"
            + participant("Cc", "Cc One", "cc@example.test")
            + "</MessageData>"
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-REDACTED", "<redacted@example.test>",
                "2026-10-20T01:59:00Z", "CHILD", data=data,
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_redacted", "message_id": "<redacted@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-REDACTED"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_redacted"]
        self.assertNotIn("sender", row["participants"])
        self.assertIsNone(row["participants"]["to"])
        self.assertEqual(row["participants"]["cc"], [
            {"name": "Cc One", "address": "cc@example.test"},
        ])
        self.assertEqual(row["participants_status"], {
            "sender": "redacted", "to": "malformed", "cc": "verified",
        })
        self.assertEqual(row["participant_counts"], {
            "sender": {"count": 1, "status": "verified"},
            "to": {"count": 1, "status": "malformed"},
            "cc": {"count": 1, "status": "verified"},
        })

    def test_invalid_recipient_is_not_cleared_by_a_later_valid_recipient(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        data = (
            "<MessageData><ImmId>ENTRY-INVALID-FIRST</ImmId>"
            "<IsPartial>0</IsPartial><IsTruncated>0</IsTruncated>"
            "<To><DisplayName>Missing Address</DisplayName></To>"
            + participant("To", "Valid", "valid@example.test")
            + "</MessageData>"
        )
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-INVALID-FIRST", "<invalid-first@example.test>",
                "2026-10-20T01:59:00Z", "CHILD", data=data,
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_invalid_first", "message_id": "<invalid-first@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-INVALID-FIRST"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_invalid_first"]
        self.assertIsNone(row["participants"]["to"])
        self.assertEqual(row["participants_status"]["to"], "malformed")

    def test_duplicate_sibling_headers_are_preserved_and_fail_identity_matching(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        first = message_header(
            "ENTRY-DUP-HEADER", "<first@example.test>",
            "2026-10-20T01:59:00Z", "CHILD",
        ).split("<FlattenedMessage>", 1)[1].split("</FlattenedMessage>", 1)[0]
        second = message_header(
            "ENTRY-DUP-HEADER", "<second@example.test>",
            "2026-10-20T01:59:00Z", "CHILD",
        ).split("<FlattenedMessage>", 1)[1].split("</FlattenedMessage>", 1)[0]
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            f"<FlattenedMessage>{first}{second}</FlattenedMessage></Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_dup_header", "message_id": "<first@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-DUP-HEADER"]}],
            log_root=self.log_root,
        )
        self.assertNotIn("hx_dup_header", result["messages"])
        self.assertEqual(result["evidence"]["messages_exact_identity_conflict"], 1)

    def test_duplicate_sibling_data_does_not_bind_recipient_roles(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        inner = message_header(
            "ENTRY-DUP-DATA", "<duplicate-data@example.test>",
            "2026-10-20T01:59:00Z", "CHILD",
        ).split("<FlattenedMessage>", 1)[1].split("</FlattenedMessage>", 1)[0]
        data = message_data("ENTRY-DUP-DATA", to=[("One", "one@example.test")])
        data += message_data("ENTRY-DUP-DATA", to=[("Two", "two@example.test")])
        inner = inner.replace("</MessageHeader>", "</MessageHeader>" + data, 1)
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            f"<FlattenedMessage>{inner}</FlattenedMessage></Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_dup_data", "message_id": "<duplicate-data@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-DUP-DATA"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_dup_data"]
        self.assertEqual(row["participants"], {
            "sender": None, "to": None, "cc": None,
        })
        self.assertEqual(row["participants_status"], {
            "sender": "conflict", "to": "conflict", "cc": "conflict",
        })

    def test_header_from_is_allowed_as_sender_when_message_data_is_absent(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-HEADER-FROM", "<header-from@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                from_name="Header Sender", from_address="header@example.test",
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_header_from", "message_id": "<header-from@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-HEADER-FROM"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_header_from"]
        self.assertEqual(row["participants"]["sender"], [
            {"name": "Header Sender", "address": "header@example.test"},
        ])
        self.assertIsNone(row["participants"]["to"])
        self.assertIsNone(row["participants"]["cc"])
        self.assertEqual(row["participants_status"]["sender"], "verified")
        self.assertEqual(row["participants_status"]["to"], "missing")
        self.assertEqual(row["participants_status"]["cc"], "missing")

    def test_standalone_message_header_exposes_typed_sender(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        header = message_header(
            "ENTRY-STANDALONE-HEADER", "<standalone@example.test>",
            "2026-10-20T01:59:00Z", "CHILD",
            from_name="Header Sender", from_address="header@example.test",
        ).split("<MessageHeader>", 1)[1].split("</MessageHeader>", 1)[0]
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            f"<MessageHeader>{header}</MessageHeader></Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_standalone_header", "message_id": "<standalone@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-STANDALONE-HEADER"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_standalone_header"]
        self.assertEqual(row["participants"]["sender"], [
            {"name": "Header Sender", "address": "header@example.test"},
        ])
        self.assertEqual(row["participants_status"]["sender"], "verified")

    def test_same_latest_log_time_conflicts_participants_even_with_one_folder(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        for suffix, name in (("1", "First"), ("2", "Second")):
            self.write_log(
                f"x.GetMessage.{suffix}.res.xmlgz",
                "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
                + message_header(
                    "ENTRY-TIE-PARTICIPANTS", "<tie-participants@example.test>",
                    "2026-10-20T01:59:00Z", "CHILD",
                    data=message_data("ENTRY-TIE-PARTICIPANTS", to=[(name, f"{name.casefold()}@example.test")]),
                )
                + "</Response>",
            )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_tie_participants", "message_id": "<tie-participants@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-TIE-PARTICIPANTS"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_tie_participants"]
        self.assertEqual(row["participants"], {
            "sender": [], "to": None, "cc": [],
        })
        self.assertEqual(row["participants_status"], {
            "sender": "verified", "to": "conflict", "cc": "verified",
        })

    def test_same_latest_log_time_conflicts_counts_when_status_differs(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        complete = message_data(
            "ENTRY-TIE-COUNTS", to=[("Same", "same@example.test")],
        )
        malformed = (
            "<MessageData><ImmId>ENTRY-TIE-COUNTS</ImmId>"
            "<IsPartial>0</IsPartial><IsTruncated>0</IsTruncated>"
            "<To><DisplayName>Missing Address</DisplayName></To>"
            "</MessageData>"
        )
        for suffix, data in (("1", complete), ("2", malformed)):
            self.write_log(
                f"x.GetMessage.{suffix}.res.xmlgz",
                "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
                + message_header(
                    "ENTRY-TIE-COUNTS", "<tie-counts@example.test>",
                    "2026-10-20T01:59:00Z", "CHILD", data=data,
                )
                + "</Response>",
            )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_tie_counts", "message_id": "<tie-counts@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-TIE-COUNTS"]}],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_tie_counts"]["participant_counts"]["to"], {
            "count": None, "status": "conflict",
        })

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

    def test_receipt_pair_maps_across_sent_date_boundary(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-RECEIVED",
                "<received@example.test>",
                "2026-10-19T23:59:00Z",
                "CHILD",
                "2026-10-20T00:01:00Z",
                "639280512600000000",
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_received", "message_id": "<received@example.test>",
              "sent_unix": 1792454340, "entry_ids": ["ENTRY-RECEIVED"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_received"]
        self.assertEqual(row["folder_id"], "CHILD")
        self.assertEqual(row["received_unix"], 1792454460.0)
        self.assertEqual(row["received_time_status"], "verified")
        self.assertEqual(
            row["received_time_method"],
            "osa_last_delivery_time",
        )

    def test_missing_receipt_does_not_fallback_to_root_receive_time(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><ReceiveTime>2026-10-20T03:00:00Z</ReceiveTime>"
            + message_header("ENTRY-MISSING-RECEIPT", "<missing@example.test>",
                             "2026-10-20T01:59:00Z", "CHILD")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_missing_receipt", "message_id": "<missing@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-MISSING-RECEIPT"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_missing_receipt"]
        self.assertEqual(row["folder_id"], "CHILD")
        self.assertIsNone(row["received_unix"])
        self.assertEqual(row["received_time_status"], "missing")
        self.assertEqual(row["received_time_method"], "osa_receipt_fields_missing")

    def test_valid_receipt_ignores_candidate_missing_both_fields(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        message = "<Response><LogTime>{}</LogTime>{}</Response>"
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            message.format(
                "2026-10-20T02:01:00Z",
                message_header("ENTRY-PARTIAL-RECEIPT", "<partial@example.test>",
                               "2026-10-20T01:59:00Z", "CHILD"),
            ),
        )
        self.write_log(
            "x.GetMessage.2.res.xmlgz",
            message.format(
                "2026-10-20T02:02:00Z",
                message_header("ENTRY-PARTIAL-RECEIPT", "<partial@example.test>",
                               "2026-10-20T01:59:00Z", "CHILD",
                               "2026-10-20T02:00:00Z", "639280584000000000"),
            ),
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_partial_receipt", "message_id": "<partial@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-PARTIAL-RECEIPT"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_partial_receipt"]
        self.assertEqual(row["received_unix"], 1792461600.0)
        self.assertEqual(row["received_time_status"], "verified")

    def test_invalid_or_naive_receipt_is_exposed_without_hiding_membership(self):
        self.assertIsNone(osa_scope._parse_explicit_time_ticks("2026-10-20T02:00:00"))
        self.assertIsNone(osa_scope._parse_dotnet_ticks("nan"))
        self.assertIsNone(osa_scope._parse_dotnet_ticks(float("nan")))
        self.assertIsNone(osa_scope._parse_dotnet_ticks(True))
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-BAD-RECEIPT", "<bad@example.test>",
                             "2026-10-20T01:59:00Z", "CHILD",
                             "2026-10-20T02:00:00", "nan")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_bad_receipt", "message_id": "<bad@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-BAD-RECEIPT"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_bad_receipt"]
        self.assertEqual(row["folder_id"], "CHILD")
        self.assertIsNone(row["received_unix"])
        self.assertEqual(row["received_time_status"], "invalid")

    def test_receipt_beyond_metadata_safe_upper_bound_is_invalid(self):
        safe = osa_scope._DOTNET_MAX_SUPPORTED_TICKS_EXCLUSIVE
        self.assertEqual(
            osa_scope._parse_explicit_time_ticks(
                "9999-12-31T15:59:59.9999999Z"
            ),
            safe - 1,
        )
        self.assertIsNone(
            osa_scope._parse_explicit_time_ticks("9999-12-31T16:00:00Z")
        )
        self.assertIsNone(osa_scope._parse_dotnet_ticks(str(safe)))
        rounded = osa_scope._received_time_evidence([{
            'received_time': '9999-12-31T15:59:59.9999999Z',
            'received_ticks': str(safe - 1),
        }])
        self.assertEqual(rounded['received_time_status'], 'invalid')
        self.assertIsNone(rounded['received_unix'])

        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-RECEIPT-OVERFLOW",
                "<overflow-receipt@example.test>",
                "2026-10-20T01:59:00Z",
                "CHILD",
                "9999-12-31T16:00:00Z",
                safe,
            )
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_receipt_overflow", "message_id": "<overflow-receipt@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-RECEIPT-OVERFLOW"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_receipt_overflow"]
        self.assertEqual(row["folder_id"], "CHILD")
        self.assertIsNone(row["received_unix"])
        self.assertEqual(row["received_time_status"], "invalid")

    def test_receipt_pair_mismatch_is_conflict(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-TICK-MISMATCH", "<mismatch@example.test>",
                             "2026-10-20T01:59:00Z", "CHILD",
                             "2026-10-20T02:00:00Z", "639280584001230000")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_tick_mismatch", "message_id": "<mismatch@example.test>",
              "sent_unix": 1792461540, "entry_ids": ["ENTRY-TICK-MISMATCH"]}],
            log_root=self.log_root,
        )
        row = result["messages"]["hx_tick_mismatch"]
        self.assertEqual(row["folder_id"], "CHILD")
        self.assertIsNone(row["received_unix"])
        self.assertEqual(row["received_time_status"], "conflict")
        self.assertEqual(
            row["received_time_method"],
            "osa_last_delivery_time_and_received_or_renew_time_mismatch",
        )

    def test_pre_epoch_and_epoch_zero_receipt_ticks_are_valid(self):
        self.assertEqual(
            osa_scope._parse_explicit_time_ticks("1970-01-01T00:00:00Z"),
            osa_scope._DOTNET_EPOCH_TICKS,
        )
        self.assertEqual(
            osa_scope._parse_explicit_time_ticks("1970-01-01T00:00:00.0000001Z"),
            osa_scope._DOTNET_EPOCH_TICKS + 1,
        )
        self.assertEqual(
            osa_scope._parse_explicit_time_ticks("1969-12-31T23:59:00Z"),
            621355967400000000,
        )
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header("ENTRY-PRE-EPOCH-RECEIPT", "<pre-receipt@example.test>",
                             "1969-12-31T23:58:00Z", "CHILD",
                             "1969-12-31T23:59:00Z", "621355967400000000")
            + "</Response>",
        )
        result = osa_scope.read_scope(
            self.profile,
            [{"id": "hx_pre_receipt", "message_id": "<pre-receipt@example.test>",
              "sent_unix": -120, "entry_ids": ["ENTRY-PRE-EPOCH-RECEIPT"]}],
            log_root=self.log_root,
        )
        self.assertEqual(result["messages"]["hx_pre_receipt"]["received_unix"], -60.0)

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

    def test_parse_cache_reuses_unchanged_logs_and_returns_fresh_metadata(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-CACHE", "<cache@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
            )
            + "</Response>",
        )
        items = [{
            "id": "hx_cache", "message_id": "<cache@example.test>",
            "sent_unix": 1792461540, "entry_ids": ["ENTRY-CACHE"],
        }]
        cache = osa_scope.LogParseCache()
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            first = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=cache
            )
            first["folders"][0]["name"] = "caller mutation"
            first["messages"]["hx_cache"]["folder_id"] = "caller mutation"
            second = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=cache
            )
        self.assertEqual(parse.call_count, 2)
        self.assertEqual(cache.hits, 2)
        self.assertEqual(cache.misses, 2)
        self.assertEqual(
            next(row for row in second["folders"] if row["id"] == "CHILD")["name"],
            "客户规则",
        )
        self.assertEqual(second["messages"]["hx_cache"]["folder_id"], "CHILD")
        self.assertEqual(first["fingerprint"], second["fingerprint"])

    def test_parse_cache_reparses_modified_log_and_drops_deleted_log(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        message_path = self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-REFRESH", "<refresh@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
            )
            + "</Response>",
        )
        items = [{
            "id": "hx_refresh", "message_id": "<refresh@example.test>",
            "sent_unix": 1792461540, "entry_ids": ["ENTRY-REFRESH"],
        }]
        cache = osa_scope.LogParseCache()
        osa_scope.read_scope(self.profile, items, log_root=self.log_root, parse_cache=cache)
        message_path.write_bytes(gzip.compress(
            (
                "<Response><LogTime>2026-10-20T02:02:00Z</LogTime>"
                + message_header(
                    "ENTRY-REFRESH", "<refresh@example.test>",
                    "2026-10-20T01:59:00Z", "INBOX",
                    received="2026-10-20T02:00:00Z",
                    received_ticks="639280584000000000",
                )
                + "</Response>"
            ).encode("utf-8"),
            mtime=0,
        ))
        refreshed = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(refreshed["messages"]["hx_refresh"]["folder_id"], "INBOX")
        self.assertEqual(refreshed["messages"]["hx_refresh"]["received_time_status"], "verified")
        self.assertEqual(cache.hits, 1)
        self.assertEqual(cache.misses, 3)
        message_path.unlink()
        removed = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(removed["messages"], {})
        self.assertEqual(removed["evidence"]["status"], "ok")

    def test_parse_cache_picks_up_added_log_and_drops_newer_received_value_when_removed(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-ADDED", "<added@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
            )
            + "</Response>",
        )
        items = [{
            "id": "hx_added", "message_id": "<added@example.test>",
            "sent_unix": 1792461540, "entry_ids": ["ENTRY-ADDED"],
        }]
        cache = osa_scope.LogParseCache()
        initial = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(initial["messages"]["hx_added"]["received_time_status"], "missing")
        added_path = self.write_log(
            "x.GetMessage.2.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:02:00Z</LogTime>"
            + message_header(
                "ENTRY-ADDED", "<added@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                received="2026-10-20T02:00:00Z",
                received_ticks="639280584000000000",
            )
            + "</Response>",
        )
        refreshed = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(refreshed["messages"]["hx_added"]["received_time_status"], "verified")
        self.assertEqual(cache.hits, 2)
        self.assertEqual(cache.misses, 3)
        added_path.unlink()
        restored = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(restored["messages"]["hx_added"]["received_time_status"], "missing")

    def test_parse_cache_invalidates_same_size_mtime_restoration_by_ctime(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        message_path = self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-SAME-STAT", "<same-stat@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
            )
            + "</Response>",
        )
        items = [{
            "id": "hx_same_stat", "message_id": "<same-stat@example.test>",
            "sent_unix": 1792461540, "entry_ids": ["ENTRY-SAME-STAT"],
        }]
        cache = osa_scope.LogParseCache()
        osa_scope.read_scope(self.profile, items, log_root=self.log_root, parse_cache=cache)
        before = message_path.stat()
        replacement = (
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-SAME-STAT", "<same-stat@example.test>",
                "2026-10-20T01:59:00Z", "INBOX",
            )
            + "</Response>"
        ).encode("utf-8")
        with gzip.open(message_path, "wb") as handle:
            handle.write(replacement)
        replacement_bytes = message_path.read_bytes()
        self.assertEqual(len(replacement_bytes), before.st_size)
        os.utime(message_path, ns=(before.st_atime_ns, before.st_mtime_ns))
        self.assertEqual(message_path.stat().st_mtime_ns, before.st_mtime_ns)
        refreshed = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(refreshed["messages"]["hx_same_stat"]["folder_id"], "INBOX")
        self.assertEqual(cache.hits, 1)
        self.assertEqual(cache.misses, 3)

    def test_parse_cache_validates_symlink_and_unreadable_sources(self):
        path = self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        cache = osa_scope.LogParseCache()
        first = osa_scope._file_stat(path)
        parsed = osa_scope._parse_log(path, "folder", self.log_root)
        cache.put(self.profile, self.log_root, "folder", path, first, parsed)
        replacement = self.log_dir / "replacement.res.xmlgz"
        replacement.write_bytes(path.read_bytes())
        path.unlink()
        path.symlink_to(replacement)
        with self.assertRaises(OSError):
            cache.get(self.profile, self.log_root, "folder", path, first)
        path.unlink()
        path.write_bytes(replacement.read_bytes())
        fresh_stat = osa_scope._file_stat(path)
        cache.put(
            self.profile, self.log_root, "folder", path, fresh_stat,
            osa_scope._parse_log(path, "folder", self.log_root),
        )
        with patch.object(osa_scope, "_readable_file_stat", side_effect=PermissionError("denied")):
            with self.assertRaises(PermissionError):
                cache.get(self.profile, self.log_root, "folder", path, fresh_stat)

    def test_parse_cache_bounds_entries_and_serialized_metadata(self):
        cache = osa_scope.LogParseCache(max_files=1, max_bytes=64 * 1024)
        for index in range(2):
            path = self.write_log(
                f"x.GetFolder.{index}.res.xmlgz",
                self.folder_log(log_time=f"2026-10-20T02:0{index}:00Z"),
            )
            source_stat = osa_scope._file_stat(path)
            parsed = osa_scope._parse_log(path, "folder", self.log_root)
            cache.put(self.profile, self.log_root, "folder", path, source_stat, parsed)
        self.assertLessEqual(cache.entry_count, 1)
        self.assertLessEqual(cache.serialized_bytes, 64 * 1024)
        self.assertLessEqual(osa_scope.LogParseCache(max_files=99999).max_files, 4096)

    def test_parse_cache_does_not_store_failed_parse_or_old_received_value(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        path = self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(
                "ENTRY-BAD-CACHE", "<bad-cache@example.test>",
                "2026-10-20T01:59:00Z", "CHILD",
                received="2026-10-20T02:00:00Z",
                received_ticks="639280584000000000",
            )
            + "</Response>",
        )
        items = [{
            "id": "hx_bad_cache", "message_id": "<bad-cache@example.test>",
            "sent_unix": 1792461540, "entry_ids": ["ENTRY-BAD-CACHE"],
        }]
        cache = osa_scope.LogParseCache()
        first = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        self.assertEqual(first["messages"]["hx_bad_cache"]["received_time_status"], "verified")
        path.write_bytes(gzip.compress(b"malformed xml", mtime=0))
        with self.assertRaisesRegex(osa_scope.OsaScopeError, "unreadable or invalid"):
            osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=cache
            )
        self.assertEqual(cache.entry_count, 1)

    def _persistent_cache_path(self, name="osa-metadata.sqlite3"):
        state = self.profile / "state"
        state.mkdir(mode=0o700)
        os.chmod(state, 0o700)
        return state / name

    def _write_cache_fixture(self, entry="ENTRY-PERSIST", message_id="<persist@example.test>"):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            + message_header(entry, message_id, "2026-10-20T01:59:00Z", "CHILD")
            + "</Response>",
        )
        return [{
            "id": "hx_persist", "message_id": message_id,
            "sent_unix": 1792461540, "entry_ids": [entry],
        }]

    def test_sqlite_cache_survives_restart_and_preserves_exact_scope(self):
        items = self._write_cache_fixture()
        cache_path = self._persistent_cache_path()
        cache = osa_scope.LogParseCache(cache_path)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            first = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=cache
            )
            self.assertEqual(parse.call_count, 2)
        self.assertTrue(cache_path.exists())
        self.assertEqual(cache_path.stat().st_mode & 0o077, 0)

        restarted = osa_scope.LogParseCache(cache_path)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            second = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=restarted
            )
        self.assertEqual(parse.call_count, 0)
        self.assertEqual(first, second)
        connection = sqlite3.connect(cache_path)
        try:
            count, payload_bytes = connection.execute(
                "SELECT COUNT(*), COALESCE(SUM(length(CAST(payload AS BLOB))), 0) "
                "FROM osa_metadata_cache"
            ).fetchone()
        finally:
            connection.close()
        self.assertLessEqual(count, 4096)
        self.assertLessEqual(payload_bytes, 64 * 1024 * 1024)

    def test_sqlite_cache_accepts_source_header_only_rows_without_counts(self):
        self.write_log("x.GetFolder.1.res.xmlgz", self.folder_log())
        wrapped = message_header(
            "ENTRY-HEADER-ONLY", "<header-only@example.test>",
            "2026-10-20T01:59:00Z", "CHILD",
        )
        header = wrapped.split("<MessageHeader>", 1)[1].split(
            "</MessageHeader>", 1
        )[0]
        self.write_log(
            "x.GetMessage.1.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:01:00Z</LogTime>"
            f"<MessageHeader>{header}</MessageHeader></Response>",
        )
        message_path = self.log_dir / "x.GetMessage.1.res.xmlgz"
        parsed = osa_scope._parse_log(message_path, "message", self.log_root)[1]
        self.assertEqual(len(parsed), 1)
        self.assertNotIn("_participant_count_values", parsed[0])
        self.assertNotIn("_participant_count_statuses", parsed[0])
        items = [{
            "id": "hx_header_only", "message_id": "<header-only@example.test>",
            "sent_unix": 1792461540, "entry_ids": ["ENTRY-HEADER-ONLY"],
        }]
        cache_path = self._persistent_cache_path()
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            first = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root,
                parse_cache=osa_scope.LogParseCache(cache_path),
            )
        self.assertEqual(parse.call_count, 2)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            second = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root,
                parse_cache=osa_scope.LogParseCache(cache_path),
            )
        self.assertEqual(parse.call_count, 0)
        self.assertEqual(first, second)

    def test_sqlite_cache_picks_up_late_receipt_and_drops_deleted_log_after_restart(self):
        items = self._write_cache_fixture("ENTRY-LATE", "<late@example.test>")
        cache_path = self._persistent_cache_path()
        osa_scope.read_scope(
            self.profile, items, log_root=self.log_root,
            parse_cache=osa_scope.LogParseCache(cache_path),
        )
        late_path = self.write_log(
            "x.GetMessage.2.res.xmlgz",
            "<Response><LogTime>2026-10-20T02:02:00Z</LogTime>"
            + message_header(
                "ENTRY-LATE", "<late@example.test>", "2026-10-20T01:59:00Z", "CHILD",
                received="2026-10-20T02:00:00Z",
                received_ticks="639280584000000000",
            )
            + "</Response>",
        )
        restarted = osa_scope.LogParseCache(cache_path)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            refreshed = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=restarted
            )
        self.assertEqual(parse.call_count, 1)
        self.assertEqual(
            refreshed["messages"]["hx_persist"]["received_time_status"], "verified"
        )
        late_path.unlink()
        final = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root,
            parse_cache=osa_scope.LogParseCache(cache_path),
        )
        self.assertEqual(final["messages"]["hx_persist"]["received_time_status"], "missing")

    def test_sqlite_cache_tamper_and_corruption_fall_back_to_source_parse(self):
        items = self._write_cache_fixture("ENTRY-TAMPER", "<tamper@example.test>")
        cache_path = self._persistent_cache_path()
        osa_scope.read_scope(
            self.profile, items, log_root=self.log_root,
            parse_cache=osa_scope.LogParseCache(cache_path),
        )
        connection = sqlite3.connect(cache_path)
        try:
            malformed_payload = b'[[],[{}],""]'
            connection.execute(
                "UPDATE osa_metadata_cache SET payload=?, payload_sha256=?, payload_bytes=?",
                (
                    malformed_payload.decode("utf-8"),
                    hashlib.sha256(malformed_payload).hexdigest(),
                    len(malformed_payload),
                ),
            )
            connection.commit()
        finally:
            connection.close()
        restarted = osa_scope.LogParseCache(cache_path)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            repaired = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=restarted
            )
        self.assertEqual(parse.call_count, 2)
        self.assertEqual(repaired["messages"]["hx_persist"]["folder_id"], "CHILD")

        cache_path.write_bytes(b"not a sqlite database")
        restarted = osa_scope.LogParseCache(cache_path)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            recovered = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=restarted
            )
        self.assertEqual(parse.call_count, 2)
        self.assertEqual(recovered["messages"]["hx_persist"]["folder_id"], "CHILD")

    def test_sqlite_cache_overflowing_timestamp_reparses_source(self):
        items = self._write_cache_fixture("ENTRY-OVERFLOW", "<overflow@example.test>")
        cache_path = self._persistent_cache_path()
        expected = osa_scope.read_scope(
            self.profile, items, log_root=self.log_root,
            parse_cache=osa_scope.LogParseCache(cache_path),
        )
        connection = sqlite3.connect(cache_path)
        try:
            for key, payload in connection.execute(
                "SELECT cache_key, payload FROM osa_metadata_cache"
            ).fetchall():
                parsed = json.loads(payload)
                if not parsed[1]:
                    continue
                parsed[1][0]["sent_unix"] = 10 ** 400
                raw = json.dumps(parsed).encode("utf-8")
                connection.execute(
                    "UPDATE osa_metadata_cache SET payload=?, payload_sha256=?, "
                    "payload_bytes=? WHERE cache_key=?",
                    (raw.decode("utf-8"), hashlib.sha256(raw).hexdigest(), len(raw), key),
                )
            connection.commit()
        finally:
            connection.close()
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            result = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root,
                parse_cache=osa_scope.LogParseCache(cache_path),
            )
        self.assertEqual(parse.call_count, 1)
        self.assertEqual(result, expected)

    def test_sqlite_cache_rejects_symlink_database_and_sidecars(self):
        items = self._write_cache_fixture("ENTRY-SAFE", "<safe@example.test>")
        state = self.profile / "state"
        state.mkdir(mode=0o700)
        os.chmod(state, 0o700)
        target = state / "target.sqlite3"
        target.write_bytes(b"do not touch")
        db_path = state / "osa-metadata.sqlite3"
        db_path.symlink_to(target)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            result = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root,
                parse_cache=osa_scope.LogParseCache(db_path),
            )
        self.assertEqual(parse.call_count, 2)
        self.assertEqual(result["messages"]["hx_persist"]["folder_id"], "CHILD")
        db_path.unlink()
        osa_scope.read_scope(
            self.profile, items, log_root=self.log_root,
            parse_cache=osa_scope.LogParseCache(db_path),
        )
        wal = Path(f"{db_path}-wal")
        wal.symlink_to(target)
        restarted = osa_scope.LogParseCache(db_path)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            recovered = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root, parse_cache=restarted
            )
        self.assertEqual(parse.call_count, 2)
        self.assertEqual(recovered["messages"]["hx_persist"]["folder_id"], "CHILD")

    def test_sqlite_cache_enforces_small_row_and_serialized_quotas(self):
        items = self._write_cache_fixture("ENTRY-QUOTA", "<quota@example.test>")
        cache_path = self._persistent_cache_path()
        cache = osa_scope.LogParseCache(cache_path, max_files=1, max_bytes=64 * 1024)
        osa_scope.read_scope(
            self.profile, items, log_root=self.log_root, parse_cache=cache
        )
        connection = sqlite3.connect(cache_path)
        try:
            count, payload_bytes = connection.execute(
                "SELECT COUNT(*), COALESCE(SUM(length(CAST(payload AS BLOB))), 0) "
                "FROM osa_metadata_cache"
            ).fetchone()
        finally:
            connection.close()
        self.assertLessEqual(count, 1)
        self.assertLessEqual(payload_bytes, 64 * 1024)
        self.assertLessEqual(cache.entry_count, 1)
        self.assertLessEqual(cache.serialized_bytes, 64 * 1024)

    def test_sqlite_cache_counts_utf8_bytes_across_restart(self):
        items = self._write_cache_fixture("ENTRY-UTF8", "<utf8@example.test>")
        folder_path = self.log_dir / "x.GetFolder.1.res.xmlgz"
        message_path = self.log_dir / "x.GetMessage.1.res.xmlgz"
        parsed_folder = osa_scope._parse_log(folder_path, "folder", self.log_root)
        parsed_message = osa_scope._parse_log(message_path, "message", self.log_root)
        seed = osa_scope.LogParseCache()
        folder_payload = seed._serialized(parsed_folder)
        message_payload = seed._serialized(parsed_message)
        self.assertIsNotNone(folder_payload)
        self.assertIsNotNone(message_payload)
        character_budget = sum(
            len(payload.decode("utf-8"))
            for payload in (folder_payload, message_payload)
        )
        byte_total = len(folder_payload) + len(message_payload)
        self.assertGreater(byte_total, character_budget)
        cache_path = self._persistent_cache_path()
        cache = osa_scope.LogParseCache(cache_path, max_bytes=character_budget)
        cache.put(
            self.profile, self.log_root, "folder", folder_path,
            osa_scope._file_stat(folder_path), parsed_folder,
        )
        cache.put(
            self.profile, self.log_root, "message", message_path,
            osa_scope._file_stat(message_path), parsed_message,
        )
        cache.flush()
        restarted = osa_scope.LogParseCache(cache_path, max_bytes=character_budget)
        restarted.begin()
        self.assertEqual(restarted.entry_count, 1)
        connection = sqlite3.connect(cache_path)
        try:
            count, byte_count = connection.execute(
                "SELECT COUNT(*), COALESCE(SUM(length(CAST(payload AS BLOB))), 0) "
                "FROM osa_metadata_cache"
            ).fetchone()
        finally:
            connection.close()
        self.assertEqual(count, 1)
        self.assertLessEqual(byte_count, character_budget)

    def test_sqlite_cache_eviction_marks_dirty_persistent_row_deleted(self):
        cache_path = self._persistent_cache_path()
        cache = osa_scope.LogParseCache(cache_path, max_files=1, max_bytes=1024)
        first_path = self.log_dir / "x.GetFolder.1.res.xmlgz"
        second_path = self.log_dir / "x.GetMessage.1.res.xmlgz"
        first_key = cache._key(self.profile, self.log_root, "folder", first_path)
        second_key = cache._key(self.profile, self.log_root, "message", second_path)
        cache._entries[first_key] = ({}, b"first")
        cache._entries[second_key] = ({}, b"second")
        cache._serialized_bytes = 11
        first_cache_key = cache._cache_key(cache._key_json(first_key))
        cache._dirty_keys.add(first_cache_key)
        cache._trim_memory()
        self.assertNotIn(first_key, cache._entries)
        self.assertNotIn(first_cache_key, cache._dirty_keys)
        self.assertIn(first_cache_key, cache._deleted_keys)

    def test_memory_fallback_does_not_accumulate_persistent_bookkeeping(self):
        path = self.write_log("x.GetFolder.seed.res.xmlgz", self.folder_log())
        source_stat = osa_scope._file_stat(path)
        parsed = osa_scope._parse_log(path, "folder", self.log_root)
        for database_failure in (False, True):
            with self.subTest(database_failure=database_failure):
                cache = osa_scope.LogParseCache(
                    self._persistent_cache_path() if database_failure else None,
                    max_files=1, max_bytes=64 * 1024,
                )
                cache.put(self.profile, self.log_root, "folder", path, source_stat, parsed)
                if database_failure:
                    with patch.object(osa_scope.sqlite3, "connect",
                                      side_effect=sqlite3.OperationalError("database unavailable")):
                        cache.flush()
                    self.assertTrue(cache._db_disabled)
                    self.assertEqual(len(cache._dirty_keys), 0)
                    self.assertEqual(len(cache._deleted_keys), 0)
                # The file safety checks have separate coverage; use rotating
                # identities here without creating 5,000 synthetic log files.
                with patch.object(osa_scope, "_readable_file_stat", return_value=source_stat):
                    for number in range(5000):
                        rotated = self.log_dir / f"x.GetFolder.{number}.res.xmlgz"
                        cache.put(self.profile, self.log_root, "folder", rotated, source_stat, parsed)
                        cache.flush()
                    self.assertEqual(cache.get(
                        self.profile, self.log_root, "folder", rotated, source_stat), parsed)
                    self.assertEqual(cache.entry_count, 1)
                    self.assertLessEqual(cache.serialized_bytes, cache.max_bytes)
                    self.assertEqual(len(cache._dirty_keys), 0)
                    self.assertEqual(len(cache._deleted_keys), 0)
                    cache.invalidate(self.profile, self.log_root, "folder", rotated)
                    self.assertEqual(cache.entry_count, 0)
                    self.assertEqual(len(cache._deleted_keys), 0)

    def test_sqlite_cache_busy_or_nonprivate_file_falls_back_to_source(self):
        items = self._write_cache_fixture("ENTRY-BUSY", "<busy@example.test>")
        cache_path = self._persistent_cache_path()
        osa_scope.read_scope(
            self.profile, items, log_root=self.log_root,
            parse_cache=osa_scope.LogParseCache(cache_path),
        )
        os.chmod(cache_path, 0o644)
        with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
            result = osa_scope.read_scope(
                self.profile, items, log_root=self.log_root,
                parse_cache=osa_scope.LogParseCache(cache_path),
            )
        self.assertEqual(parse.call_count, 2)
        self.assertEqual(result["messages"]["hx_persist"]["folder_id"], "CHILD")
        os.chmod(cache_path, 0o600)
        locker = sqlite3.connect(cache_path)
        try:
            locker.execute("BEGIN EXCLUSIVE")
            with patch.object(osa_scope, "_parse_log", wraps=osa_scope._parse_log) as parse:
                result = osa_scope.read_scope(
                    self.profile, items, log_root=self.log_root,
                    parse_cache=osa_scope.LogParseCache(cache_path),
                )
            self.assertEqual(parse.call_count, 2)
            self.assertEqual(result["messages"]["hx_persist"]["folder_id"], "CHILD")
        finally:
            locker.rollback()
            locker.close()


if __name__ == "__main__":
    unittest.main()
