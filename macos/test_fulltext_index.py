"""Focused coverage for the private incremental full-text cache."""

from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).parent))
import fulltext_index


class FullTextIndexTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.root.chmod(0o700)
        self.db = self.root / "fulltext.sqlite3"
        self.texts = {
            "one": '会议预算复核 Straße a*b "OR" 中文关键字',
            "two": "没有目标词",
            "three": "x" * 20_000 + "尾部独有预算复核",
        }
        self.rows = [self.row(identifier, "cached_html") for identifier in self.texts]
        self.index = {"profile": "/synthetic/profile", "items": self.rows}
        self.reader = patch.object(
            fulltext_index.core, "_resolve_deferred_body", side_effect=self.resolve
        )
        self.raw = patch.object(fulltext_index.core, "_verified_raw_payload", return_value=None)
        self.units = patch.object(
            fulltext_index.core, "_read_text_units", side_effect=self.read_units
        )
        self.reader_mock = self.reader.start()
        self.raw.start()
        self.units.start()
        self.addCleanup(self.units.stop)
        self.addCleanup(self.raw.stop)
        self.addCleanup(self.reader.stop)
        self.addCleanup(self.temporary.cleanup)

    def row(self, identifier, status="cached_html"):
        text = self.texts.get(identifier, "")
        return {
            "id": identifier,
            "message_id": f"<{identifier}@example.test>",
            "preview": text[:12],
            "body": {
                "status": status,
                "source": identifier,
                "text_utf16_units": len(text.encode("utf-16-le")) // 2,
            },
            "evidence": {"frame": identifier},
            "folder_id": "inbox",
            "received_unix": 1,
        }

    def resolve(self, _directory, _index, row):
        return row["body"]

    def read_units(self, _directory, body, _start, _end):
        return self.texts[body["source"]]

    def test_literal_queries_short_chinese_punctuation_casefold_and_long_tail(self):
        index = fulltext_index.FullTextIndex(self.db)
        for query, expected in (
            ("算", {"one", "three"}),
            ("预算", {"one", "three"}),
            ("预算复核", {"one", "three"}),
            ("STRASSE", {"one"}),
            ("a*b", {"one"}),
            ('"OR"', {"one"}),
            ("' OR 1=1", set()),
            ("独有", {"three"}),
            ("x" * 1000, {"three"}),
        ):
            with self.subTest(query=query):
                hits, stats = index.search(self.root, self.index, self.rows, query)
                self.assertEqual(set(hits), expected)
                self.assertEqual(stats["errors"], 0)
                self.assertTrue(all(value for value in hits.values()))
                self.assertTrue(all(len(value) <= 242 for value in hits.values()))
        self.assertIn("Straße", index.search(self.root, self.index, self.rows, "STRASSE")[0]["one"])
        nul_hits, nul_stats = index.search(self.root, self.index, self.rows, "尾\x00部")
        self.assertEqual(nul_hits, {})
        self.assertTrue(nul_stats["complete"])

    def test_incremental_reopen_changed_body_and_scope_pruning(self):
        index = fulltext_index.FullTextIndex(self.db)
        first_hits, first = index.search(self.root, self.index, self.rows, "预算")
        self.assertEqual(set(first_hits), {"one", "three"})
        self.assertEqual(first["updated"], 3)
        self.reader_mock.reset_mock()

        reopened = fulltext_index.FullTextIndex(self.db)
        _, reused = reopened.search(self.root, self.index, self.rows, "预算")
        self.assertEqual(reused["updated"], 0)
        self.assertEqual(reused["reused"], 3)
        self.reader_mock.assert_not_called()

        # Folder/received metadata is excluded from the body signature.
        self.rows[0]["folder_id"] = "sent"
        _, metadata_only = reopened.search(self.root, self.index, [self.rows[0]], "预算")
        self.assertEqual(metadata_only["reused"], 1)
        # Pruning uses the full scoped index, so the two unfiltered rows remain.
        _, other = reopened.search(self.root, self.index, [self.rows[1]], "目标")
        self.assertEqual(other["reused"], 1)

        old_signature = self.rows[0]["body"]
        self.rows[0]["body"] = {**old_signature, "source_fingerprint": "changed"}
        self.texts["one"] = "已改正文结算确认"
        self.rows[0]["body"]["text_utf16_units"] = len(self.texts["one"].encode("utf-16-le")) // 2
        changed_hits, changed = reopened.search(self.root, self.index, [self.rows[0]], "结算确认")
        self.assertEqual(set(changed_hits), {"one"})
        self.assertEqual(changed["updated"], 1)
        self.assertEqual(reopened.search(self.root, self.index, [self.rows[0]], "预算")[0], {})

        # The body is removed only after it leaves index.items entirely.
        self.index["items"] = [self.rows[0]]
        _, pruned = reopened.search(self.root, self.index, [self.rows[0]], "结算确认")
        self.assertEqual(pruned["candidates"], 1)
        self.assertEqual(reopened.search(self.root, self.index, [self.rows[0]], "目标")[0], {})

    def test_missing_is_cached_but_errors_retry_and_budget_leaves_pending(self):
        missing = self.row("missing", "not_cached")
        self.index["items"] = [missing]
        first_index = fulltext_index.FullTextIndex(self.db)
        _, first = first_index.search(self.root, self.index, [missing], "目标")
        self.assertEqual((first["missing"], first["updated"], first["complete"]), (1, 1, False))
        self.reader_mock.reset_mock()
        _, reused = first_index.search(self.root, self.index, [missing], "目标")
        self.assertEqual((reused["missing"], reused["reused"]), (1, 1))
        self.reader_mock.assert_not_called()

        self.texts["missing"] = "补回来的目标正文"
        missing["body"] = {"status": "cached_html", "source": "missing",
                            "text_utf16_units": len(self.texts["missing"].encode("utf-16-le")) // 2,
                            "source_fingerprint": "new"}
        hits, refreshed = first_index.search(self.root, self.index, [missing], "目标")
        self.assertEqual(set(hits), {"missing"})
        self.assertEqual(refreshed["updated"], 1)

        self.texts["broken"] = "预算 body"
        broken = self.row("broken")
        self.index["items"] = [broken]
        with patch.object(fulltext_index.core, "_resolve_deferred_body",
                          side_effect=ValueError("transient body error")):
            _, error_stats = first_index.search(self.root, self.index, [broken], "预算")
        self.assertEqual((error_stats["errors"], error_stats["pending"]), (1, 0))
        hits, retry = first_index.search(self.root, self.index, [broken], "预算")
        self.assertEqual(set(hits), {"broken"})
        self.assertEqual(retry["updated"], 1)

        self.texts["pending-a"] = "预算 pending a"
        self.texts["pending-b"] = "预算 pending b"
        pending_rows = [self.row("pending-a"), self.row("pending-b")]
        self.index["items"] = pending_rows
        calls = {"count": 0}

        def budgeted(_directory, _index, row):
            calls["count"] += 1
            if calls["count"] == 2:
                raise fulltext_index.core.SourceBudgetExceeded("test budget")
            return row["body"]

        with patch.object(fulltext_index.core, "_resolve_deferred_body", side_effect=budgeted):
            _, bounded = first_index.search(self.root, self.index, pending_rows, "预算")
        self.assertEqual((bounded["indexed"], bounded["pending"], bounded["complete"]), (1, 1, False))
        hits, resumed = first_index.search(self.root, self.index, pending_rows, "预算")
        self.assertEqual(set(hits), {"pending-a", "pending-b"})
        self.assertEqual((resumed["reused"], resumed["updated"], resumed["pending"]), (1, 1, 0))

    def test_external_fts_stays_consistent_for_missing_scope_prune(self):
        missing = self.row("missing", "not_cached")
        indexed = self.rows[0]
        index = fulltext_index.FullTextIndex(self.db)
        scoped = {"profile": "/synthetic/profile", "items": [missing, indexed]}
        index.search(self.root, scoped, [missing, indexed], "预算")

        connection = sqlite3.connect(str(self.db))
        try:
            self.assertEqual(
                connection.execute(
                    "INSERT INTO body_fts(body_fts,rank) VALUES('integrity-check',1)"
                ).fetchall(),
                [],
            )
        finally:
            connection.close()

        # Removing the missing row from the full scope must not remove the
        # indexed row or leave an external-content FTS row orphaned.
        narrowed = {"profile": "/synthetic/profile", "items": [indexed]}
        hits, stats = index.search(self.root, narrowed, [indexed], "预算")
        self.assertEqual(set(hits), {"one"})
        self.assertEqual(stats["reused"], 1)
        connection = sqlite3.connect(str(self.db))
        try:
            self.assertEqual(
                connection.execute(
                    "INSERT INTO body_fts(body_fts,rank) VALUES('integrity-check',1)"
                ).fetchall(),
                [],
            )
            self.assertEqual(
                connection.execute("SELECT id FROM documents").fetchall(), [("one",)]
            )
        finally:
            connection.close()

    def test_private_paths_corruption_and_source_are_rejected_or_untouched(self):
        source = self.root / "original-outlook-cache"
        source.write_bytes(b"unchanged")
        index = fulltext_index.FullTextIndex(self.db)
        index.search(self.root, self.index, self.rows, "预算")
        self.assertEqual(source.read_bytes(), b"unchanged")

        # An external-content FTS mismatch must fail explicitly instead of
        # silently returning no hits.
        connection = sqlite3.connect(str(self.db))
        try:
            connection.execute(
                "DELETE FROM body_fts WHERE rowid=("
                "SELECT doc_key FROM documents WHERE id='one')"
            )
            connection.commit()
        finally:
            connection.close()
        with self.assertRaisesRegex(fulltext_index.FullTextIndexError, "corrupt or unsupported"):
            fulltext_index.FullTextIndex(self.db).search(self.root, self.index, self.rows, "预算")

        unsafe_parent = self.root / "unsafe"
        unsafe_parent.mkdir()
        unsafe_parent.chmod(0o755)
        with self.assertRaises(fulltext_index.FullTextIndexError):
            fulltext_index.FullTextIndex(unsafe_parent / "fulltext.sqlite3")

        corrupt = self.root / "corrupt.sqlite3"
        corrupt.write_bytes(b"not sqlite")
        corrupt.chmod(0o600)
        corrupt_index = fulltext_index.FullTextIndex(corrupt)
        with self.assertRaisesRegex(fulltext_index.FullTextIndexError, "corrupt or unsupported"):
            corrupt_index.search(self.root, self.index, [], "x")

        target = self.root / "db-target"
        target.write_bytes(b"target")
        link = self.root / "link.sqlite3"
        link.symlink_to(target)
        with self.assertRaises(fulltext_index.FullTextIndexError):
            fulltext_index.FullTextIndex(link)


if __name__ == "__main__":
    unittest.main()
