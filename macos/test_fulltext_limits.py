"""Regression coverage for bounded SQLite full-text indexing failures."""

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import re
import sqlite3
import sys
import threading
import time
import unittest
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).parent))
import fulltext_index
import test_fulltext_index as fulltext_fixtures


class FullTextLimitTests(unittest.TestCase):
    """Exercise failure paths without duplicating the body fixture."""

    def setUp(self):
        self.fixture = fulltext_fixtures.FullTextIndexTests(methodName="runTest")
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)

    def _row_with_text(self, identifier, text):
        self.fixture.texts[identifier] = text
        return self.fixture.row(identifier)

    def _run_contended_search(self, wait_seconds):
        fixture = self.fixture
        first = fulltext_index.FullTextIndex(fixture.db)
        second = fulltext_index.FullTextIndex(fixture.db)
        identifier = f"contended-{wait_seconds}"
        row = self._row_with_text(identifier, "预算 contention")
        index = {"profile": "/synthetic/profile", "items": [row]}
        started = threading.Event()
        release = threading.Event()

        def slow_resolve(directory, source_index, candidate):
            started.set()
            if not release.wait(3):
                raise AssertionError("contended body read was not released")
            return fixture.resolve(directory, source_index, candidate)

        def run(reader):
            return reader.search(fixture.root, index, [row], "预算")

        with patch.object(fulltext_index, "DB_WAIT_SECONDS", wait_seconds), \
             patch.object(fulltext_index.core, "_resolve_deferred_body", side_effect=slow_resolve):
            with ThreadPoolExecutor(max_workers=2) as pool:
                first_future = pool.submit(run, first)
                self.assertTrue(started.wait(2), "first search did not reach body resolution")
                second_future = pool.submit(run, second)
                if wait_seconds:
                    time.sleep(0.1)
                    release.set()
                    first_result = first_future.result(timeout=6)
                    second_result = second_future.result(timeout=6)
                    return first_result, second_result
                time.sleep(0.1)
                try:
                    second_future.result(timeout=2)
                except Exception as error:
                    release.set()
                    first_future.result(timeout=6)
                    return error
                finally:
                    release.set()
        self.fail("contended search unexpectedly completed")

    def test_sqlite_full_preserves_original_error_and_committed_rows(self):
        index = fulltext_index.FullTextIndex(self.fixture.db)
        first_row = self.fixture.rows[0]
        first_scope = {"profile": "/synthetic/profile", "items": [first_row]}
        index.search(self.fixture.root, first_scope, [first_row], "预算")

        huge = self._row_with_text("huge", "z" * 2_000_000)
        scope = {"profile": "/synthetic/profile", "items": [first_row, huge]}
        original_connect = fulltext_index.FullTextIndex._connect

        def low_page_connect(reader, *args, **kwargs):
            connection = original_connect(reader, *args, **kwargs)
            connection.execute("PRAGMA max_page_count=32")
            connection.commit()
            return connection

        with patch.object(fulltext_index.FullTextIndex, "_connect", low_page_connect):
            with self.assertRaises(fulltext_index.FullTextIndexError) as raised:
                index.search(self.fixture.root, scope, [first_row, huge], "预算")

        message = str(raised.exception)
        self.assertRegex(message, re.compile(r"capacity|disk|full", re.IGNORECASE))
        self.assertNotRegex(message, re.compile(r"no such savepoint|corrupt or unsupported", re.IGNORECASE))
        connection = sqlite3.connect(str(self.fixture.db))
        try:
            self.assertEqual(connection.execute("SELECT id FROM documents").fetchall(), [("one",)])
        finally:
            connection.close()

    def test_concurrent_instances_wait_for_writer_and_busy_is_retryable(self):
        first_result, second_result = self._run_contended_search(fulltext_index.DB_WAIT_SECONDS)
        self.assertEqual(first_result[1]["errors"], 0)
        self.assertEqual(second_result[1]["reused"], 1)

        busy = self._run_contended_search(0.0)
        self.assertIsInstance(busy, fulltext_index.FullTextIndexBusy)
        self.assertTrue(getattr(busy, "retryable", False))
        self.assertRegex(str(busy), re.compile(r"another.*search|retry", re.IGNORECASE))
        self.assertNotRegex(str(busy), re.compile(r"corrupt or unsupported", re.IGNORECASE))

    def test_query_timeout_commits_index_and_explains_narrow_retry(self):
        index = fulltext_index.FullTextIndex(self.fixture.db)
        with patch.object(fulltext_index, "QUERY_SECONDS", 0.0):
            with self.assertRaises(fulltext_index.FullTextQueryTimeout) as raised:
                index.search(self.fixture.root, self.fixture.index, self.fixture.rows, "预算")

        self.assertRegex(
            str(raised.exception),
            re.compile(r"narrow.*(date|folder|person)|(date|folder|person).*narrow", re.IGNORECASE),
        )
        hits, stats = index.search(self.fixture.root, self.fixture.index, self.fixture.rows, "预算")
        self.assertEqual(set(hits), {"one", "three"})
        self.assertEqual(stats["updated"], 0)
        self.assertEqual(stats["reused"], 3)


if __name__ == "__main__":
    unittest.main()
