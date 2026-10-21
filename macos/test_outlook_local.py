import copy
import gzip
import hashlib
import io
import mmap
import struct
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


sys.path.insert(0, str(Path(__file__).parent))
import hxcontainer  # noqa: E402
import outlook_local  # noqa: E402
from test_hxcontainer import make_block, store_header  # noqa: E402
from test_hxrecords import (  # noqa: E402
    attachment_frame,
    exchange_entry,
    inline_body_frame,
    mail_frame,
    registry_frame,
    serialized_frame,
    ticks,
    verified_ticks,
)


PREVIEW = "中文正文预览内容以及搜索关键词"
GZIP_PREVIEW = "压缩正文预览内容以及另一组搜索词"
BODY_REFERENCE = "~/Files/EFMData/message-body.html"
GZIP_BODY_REFERENCE = "~/Files/EFMData/compressed-body.html.gz"
ATTACHMENT_REFERENCE = "~/Files/Attachments/owner-a/123456.bin"


class OutlookLocalTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        root = Path(self.temporary.name)
        self.profile = root / "profile"
        self.state = root / "state"
        self.output = root / "exported.bin"
        self.profile.mkdir()

        self.source_cache = self.profile / "HxStore.hxd"
        self.source_cache.write_bytes(b"synthetic HxStore snapshot A")
        self.body_source = self.profile / "Files/EFMData/message-body.html"
        self.body_source.parent.mkdir(parents=True)
        self.body_source.write_bytes(
            (
                "<html><body><p>"
                + PREVIEW
                + "。</p><p>冻结后仍可读取的正文。</p></body></html>"
            ).encode("utf-8")
        )
        self.gzip_body_source = self.profile / "Files/EFMData/compressed-body.html.gz"
        self.gzip_body_source.write_bytes(
            gzip.compress(
                (
                    "<html><body><p>"
                    + GZIP_PREVIEW
                    + "。</p><p>gzip 冻结正文。</p></body></html>"
                ).encode("utf-8")
            )
        )
        self.attachment_source = self.profile / "Files/Attachments/owner-a/123456.bin"
        self.attachment_source.parent.mkdir(parents=True)
        self.attachment_source.write_bytes(b"attachment-bytes-\x00\xff")

        self.entry_first = exchange_entry(1)
        self.entry_second = exchange_entry(2)
        first = mail_frame(
            0xABC,
            before=("sender@example.com", "张伟🧪"),
            after=("中文主题甲", "<first@example.test>", PREVIEW, "中文主题甲"),
            tail=verified_ticks(0xABC, 946_684_810) + self.entry_first,
        )
        first_registry = registry_frame(
            0xF1,
            self.entry_first,
            (BODY_REFERENCE, ATTACHMENT_REFERENCE),
        )
        second = mail_frame(
            0xABD,
            before=("other@example.com", "王小明😀"),
            after=("中文主题乙", "<second@example.test>", GZIP_PREVIEW, "中文主题乙"),
            tail=verified_ticks(0xABD, 946_684_820) + self.entry_second,
        )
        second_registry = registry_frame(
            0xF2,
            self.entry_second,
            (GZIP_BODY_REFERENCE,),
        )
        third = mail_frame(
            0xABE,
            before=("third@example.com", "赵六"),
            after=("中文主题丙", "<third@example.test>", "中文主题丙"),
            tail=verified_ticks(0xABE, 946_684_830),
        )
        self.block = SimpleNamespace(
            offset=0xE3000,
            data=first + first_registry + second + second_registry + third,
        )

        def iter_blocks_stub(_data, stats=None):
            if stats is not None:
                stats.update(candidates=1, verified=1, rejected=0)
            return [self.block]

        self.iter_blocks_patch = patch.object(
            outlook_local, "iter_blocks", side_effect=iter_blocks_stub
        )
        self.iter_blocks_mock = self.iter_blocks_patch.start()

        def read_block_at_stub(_source, offset):
            if offset != self.block.offset:
                raise ValueError("unknown synthetic block")
            return hxcontainer.Block(self.block.offset, 8, self.block.data)

        self.read_block_patch = patch.object(
            hxcontainer, "read_block_at", side_effect=read_block_at_stub
        )
        self.read_block_patch.start()

    def tearDown(self):
        self.iter_blocks_patch.stop()
        self.read_block_patch.stop()
        self.temporary.cleanup()

    def snapshot(self):
        args = SimpleNamespace(
            profile=str(self.profile), state=str(self.state), from_date="2000-01-01"
        )
        result = outlook_local.make_snapshot(args)
        directory, index = outlook_local.load_index(str(self.state))
        return result, directory, index

    def deferred_snapshot(self):
        args = SimpleNamespace(
            profile=str(self.profile), state=str(self.state), from_date="2000-01-01",
            freeze_bodies=False,
        )
        result = outlook_local.make_snapshot(args)
        directory, index = outlook_local.load_index(str(self.state))
        return result, directory, index

    @staticmethod
    def flaky_payload_temporary(fail_number):
        """Patch payload temp writes once, after leaving a partial file."""
        original = outlook_local.tempfile.NamedTemporaryFile
        attempts = {"payload": 0}

        def factory(*args, **kwargs):
            handle = original(*args, **kwargs)
            if kwargs.get("prefix") != ".payload-":
                return handle
            attempts["payload"] += 1
            if attempts["payload"] != fail_number:
                return handle

            class PartialWriter:
                def __init__(self, wrapped):
                    self.wrapped = wrapped

                def __enter__(self):
                    return self

                def __exit__(self, *exit_args):
                    return self.wrapped.__exit__(*exit_args)

                def write(self, data):
                    partial = data[:max(1, len(data) // 2)]
                    self.wrapped.write(partial)
                    raise OSError("simulated interrupted payload write")

                def __getattr__(self, name):
                    return getattr(self.wrapped, name)

            return PartialWriter(handle)

        return patch.object(
            outlook_local.tempfile, "NamedTemporaryFile", side_effect=factory
        ), attempts

    def profile_tree(self):
        result = {}
        for path in self.profile.rglob("*"):
            if path.is_file():
                result[str(path.relative_to(self.profile))] = (
                    hashlib.sha256(path.read_bytes()).hexdigest(),
                    path.stat().st_mtime_ns,
                )
        return result

    @staticmethod
    def row(index, message_id):
        return next(row for row in index["items"] if row["message_id"] == message_id)

    def set_block_data(self, *frames):
        self.block.data = b"".join(frames)

    def test_metadata_preserves_milliseconds_even_near_supported_maximum(self):
        _, _, index = self.snapshot()
        row = dict(index['items'][0])
        for timestamp, suffix in [(1893456000.5, '00.500000+08:00'),
                                  (253402243199.123, '59.123000+08:00')]:
            row['sent_unix'] = timestamp
            self.assertTrue(outlook_local.metadata(row)['sent_at'].endswith(suffix))

    def test_last_supported_to_date_has_no_unrepresentable_next_day(self):
        _, _, index = self.snapshot()
        for row in index['items']:
            row['received_unix'] = row['sent_unix'] + 60
        result = outlook_local.search(index, SimpleNamespace(
            from_date=None, to_date='9999-12-31', query='', field='all', limit=100, cursor=None))
        self.assertEqual(result['total_matched'], len(index['items']))

    def test_snapshot_verification_streams_in_bounded_chunks(self):
        original_open = Path.open
        source = self.source_cache.resolve()
        reads = []
        source_opens = 0

        class BoundedReader:
            def __init__(self, handle):
                self.handle = handle

            def __enter__(self):
                return self

            def __exit__(self, *args):
                self.handle.close()

            def read(self, size):
                if not 0 <= size <= 1024 * 1024:
                    raise AssertionError('verification must not read a full cache at once')
                reads.append(size)
                return self.handle.read(size)

        def open_source(path, *args, **kwargs):
            nonlocal source_opens
            handle = original_open(path, *args, **kwargs)
            if path == source and args and args[0] == 'rb':
                source_opens += 1
                if source_opens == 2:
                    return BoundedReader(handle)
            return handle

        with patch.object(Path, 'open', open_source):
            self.snapshot()
        self.assertEqual(source_opens, 2)
        self.assertTrue(reads)

    def test_snapshot_rejects_rewritten_bytes_in_streaming_verification(self):
        source = self.source_cache.resolve()
        source_bytes = source.read_bytes()
        original_open = Path.open
        source_opens = 0

        def inconsistent_source(path, *args, **kwargs):
            nonlocal source_opens
            if path == source and args and args[0] == 'rb':
                source_opens += 1
                if source_opens % 2 == 0:
                    return io.BytesIO(b'X' * len(source_bytes))
            return original_open(path, *args, **kwargs)

        with patch.object(Path, 'open', inconsistent_source):
            with self.assertRaisesRegex(ValueError, 'cache kept changing'):
                self.snapshot()
        self.assertFalse((self.state / 'CURRENT').exists())

    def test_snapshot_parses_private_mmap_when_memory_bound_is_small(self):
        fake_stat = {"size": 2 * 1024**3 + 1, "mtime_ns": 1, "inode": 1}

        def fake_copy(_source, directory):
            temporary = Path(directory) / ".source-fake"
            temporary.write_bytes(b"private mmap source")
            return temporary, fake_stat, outlook_local.sha(b"private mmap source")

        with patch.object(outlook_local, "MAX_IN_MEMORY_SOURCE", 1), patch.object(
            outlook_local, "copy_stable_source", side_effect=fake_copy
        ) as copy:
            self.snapshot()
        copy.assert_called_once()
        source_argument = self.iter_blocks_mock.call_args.args[0]
        self.assertIsInstance(source_argument, mmap.mmap)
        self.assertEqual(list(self.state.glob(".building-*")), [])

    def test_copy_stable_source_uses_bounded_reads(self):
        source_bytes = b"source-data" * (2 * outlook_local.SOURCE_COPY_CHUNK // 11) + b"tail"
        self.source_cache.write_bytes(source_bytes)
        build = self.state / "build"
        build.mkdir(parents=True)
        original_open = Path.open
        source = self.source_cache.resolve()
        reads = []

        class BoundedReader:
            def __init__(self, handle):
                self.handle = handle

            def __enter__(self):
                return self

            def __exit__(self, *args):
                self.handle.close()

            def read(self, size):
                reads.append(size)
                if size > outlook_local.SOURCE_COPY_CHUNK:
                    raise AssertionError("source copy read exceeded bounded chunk")
                return self.handle.read(size)

        def bounded_open(path, *args, **kwargs):
            handle = original_open(path, *args, **kwargs)
            if path.resolve() == source and args and args[0] == "rb":
                return BoundedReader(handle)
            return handle

        with patch.object(Path, "open", bounded_open):
            temporary, _, digest = outlook_local.copy_stable_source(self.source_cache, build)
        try:
            self.assertEqual(temporary.read_bytes(), source_bytes)
            self.assertEqual(digest, outlook_local.sha(source_bytes))
        finally:
            temporary.unlink(missing_ok=True)
        self.assertTrue(reads)
        self.assertTrue(all(size <= outlook_local.SOURCE_COPY_CHUNK for size in reads))

    def test_copy_stable_source_rejects_mutation_and_removes_private_copy(self):
        source_bytes = b"stable-source"
        self.source_cache.write_bytes(source_bytes)
        build = self.state / "build"
        build.mkdir(parents=True)
        original_open = Path.open
        source = self.source_cache.resolve()
        source_opens = 0

        def changing_open(path, *args, **kwargs):
            nonlocal source_opens
            if path.resolve() == source and args and args[0] == "rb":
                source_opens += 1
                if source_opens == 2:
                    return io.BytesIO(b"X" * len(source_bytes))
            return original_open(path, *args, **kwargs)

        with patch.object(Path, "open", changing_open):
            with self.assertRaisesRegex(ValueError, "cache changed"):
                outlook_local.copy_stable_source(self.source_cache, build)
        self.assertEqual(list(build.glob(".source-*")), [])

    def test_copy_stable_source_reports_insufficient_scratch_space(self):
        build = self.state / "build"
        build.mkdir(parents=True)
        with patch.object(outlook_local.shutil, "disk_usage") as disk_usage:
            disk_usage.return_value.free = 0
            with self.assertRaisesRegex(ValueError, "insufficient scratch space"):
                outlook_local.copy_stable_source(self.source_cache, build)
        self.assertEqual(list(build.glob(".source-*")), [])

    def test_unchanged_snapshot_reuses_without_reading_hxstore(self):
        first_result, _, _ = self.snapshot()

        with patch.object(
            outlook_local, "stable_read", side_effect=AssertionError("unexpected full read")
        ):
            second_result, _, _ = self.snapshot()

        self.assertEqual(first_result["snapshot"], second_result["snapshot"])
        self.assertTrue(second_result["reused"])

    def test_source_change_forces_stable_hxstore_read_and_rebuilds(self):
        first_result, _, _ = self.snapshot()
        self.source_cache.write_bytes(self.source_cache.read_bytes() + b" source change")

        with patch.object(outlook_local, "stable_read", wraps=outlook_local.stable_read) as stable:
            second_result, _, _ = self.snapshot()

        self.assertNotEqual(first_result["snapshot"], second_result["snapshot"])
        self.assertGreaterEqual(stable.call_count, 1)

    def test_snapshot_search_read_attachment_preserve_original_profile_tree(self):
        before = self.profile_tree()
        summary, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")

        search_args = SimpleNamespace(
            query="搜索关键词",
            field="all",
            from_date=None,
            to_date=None,
            limit=20,
            cursor=None,
        )
        search_result = outlook_local.search(index, search_args)
        message = outlook_local.read_message(directory, index, first["id"], 20_000)
        attachment_result = outlook_local.read_attachment(
            directory, index, first["id"], first["attachments"][0]["id"], self.output
        )

        self.assertEqual(summary["messages"], 3)
        self.assertEqual([item["id"] for item in search_result["items"]], [first["id"]])
        self.assertFalse(search_result["items"][0]["attachments_complete"])
        self.assertIn(PREVIEW, message["body"]["text"])
        self.assertFalse(message["attachments_complete"])
        self.assertIsNone(first["attachments"][0]["name"])
        self.assertEqual(first["attachments"][0]["cache_name"], "123456.bin")
        self.assertEqual(attachment_result["sha256"], hashlib.sha256(self.output.read_bytes()).hexdigest())
        self.assertEqual(self.output.read_bytes(), b"attachment-bytes-\x00\xff")
        self.assertEqual(before, self.profile_tree())

    def test_registry_raw_and_gzip_bodies_are_frozen_and_readable(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        second = self.row(index, "<second@example.test>")

        raw = outlook_local.read_message(directory, index, first["id"], 20_000)
        compressed = outlook_local.read_message(directory, index, second["id"], 20_000)
        self.assertEqual(raw["body"]["format"], "html")
        self.assertEqual(compressed["body"]["format"], "gzip-html")
        self.assertIn(PREVIEW, raw["body"]["text"])
        self.assertIn(GZIP_PREVIEW, compressed["body"]["text"])

    def test_corrupt_gzip_external_body_is_unavailable_without_aborting_snapshot(self):
        corrupted = bytearray(self.gzip_body_source.read_bytes())
        corrupted[10] ^= 0xFF
        self.gzip_body_source.write_bytes(corrupted)

        _, _, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        second = self.row(index, "<second@example.test>")
        self.assertEqual(first["body"]["status"], "cached_html")
        self.assertEqual(second["body"]["status"], "unavailable")
        self.assertEqual(second["body"]["diagnosis"]["code"], "body_candidates_rejected")
        self.assertEqual(second["body"]["checks"][0]["status"], "unavailable")

    def test_type_202_external_body_follows_header_target(self):
        wrong_reference = "~/Files/EFMData/wrong-target.html"
        target_reference = "~/Files/EFMData/declared-target.html"
        self.profile.joinpath("Files/EFMData/wrong-target.html").write_text(
            f"<html><body>{PREVIEW}。错误正文。</body></html>"
        )
        self.profile.joinpath("Files/EFMData/declared-target.html").write_text(
            f"<html><body>{PREVIEW}。正确正文。</body></html>"
        )
        header_key, body_key = 0xA01, 0xA02
        target = struct.pack("<H", 202) + b"\0" * 8 + struct.pack("<Q", body_key)
        header = mail_frame(
            header_key,
            before=("target@example.com", "Target sender"),
            after=("Target subject", "<target@example.test>", PREVIEW, "Target subject"),
            tail=verified_ticks(header_key, 946_684_900) + target,
        )
        self.set_block_data(
            header,
            serialized_frame(202, header_key, wrong_reference.encode("utf-16le") + b"\0\0"),
            serialized_frame(202, body_key, target_reference.encode("utf-16le") + b"\0\0"),
        )

        _, directory, index = self.snapshot()
        row = self.row(index, "<target@example.test>")
        self.assertEqual(row["body"]["status"], "cached_html")
        self.assertEqual(row["body"]["source_path"], target_reference)
        result = outlook_local.read_message(directory, index, row["id"], 20_000)
        self.assertIn("正确正文", result["body"]["text"])
        self.assertNotIn("错误正文", result["body"]["text"])

    def test_exact_entry_id_body_wins_over_frame_key_candidate(self):
        weak_reference = "~/Files/EFMData/weak-frame-key.html"
        exact_reference = "~/Files/EFMData/exact-entry-id.html"
        self.profile.joinpath("Files/EFMData/weak-frame-key.html").write_text(
            f"<html><body>{PREVIEW}。错误正文，较长的候选。</body></html>"
        )
        self.profile.joinpath("Files/EFMData/exact-entry-id.html").write_text(
            f"<html><body>{PREVIEW}。精确正文。</body></html>"
        )
        key = 0xA03
        self.set_block_data(
            mail_frame(
                key,
                before=("exact@example.com", "Exact sender"),
                after=("Exact subject", "<exact@example.test>", PREVIEW, "Exact subject"),
                tail=verified_ticks(key, 946_684_910) + self.entry_first,
            ),
            serialized_frame(202, key, weak_reference.encode("utf-16le") + b"\0\0"),
            registry_frame(0xFA03, self.entry_first, (exact_reference,)),
        )

        _, directory, index = self.snapshot()
        row = self.row(index, "<exact@example.test>")
        self.assertEqual(row["body"]["status"], "cached_html")
        self.assertEqual(row["body"]["source_path"], exact_reference)
        result = outlook_local.read_message(directory, index, row["id"], 20_000)
        self.assertIn("精确正文", result["body"]["text"])
        self.assertNotIn("错误正文", result["body"]["text"])

    def test_read_uses_frozen_text_without_reparsing_html(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        body = first["body"]
        text_payload = directory / "textpayloads" / body["text_sha256"]
        text = text_payload.read_bytes().decode("utf-16-le")
        self.assertEqual(body["text_chars"], len(text))
        self.assertEqual(body["text_utf16_units"], len(text_payload.read_bytes()) // 2)
        with patch.object(outlook_local, "decode_body", side_effect=AssertionError("HTML was reparsed")):
            result = outlook_local.read_message(directory, index, first["id"], len(PREVIEW) + 1)
        self.assertEqual(result["body"]["text"], PREVIEW + "。")
        self.assertTrue(result["body"]["truncated"])

    def _assert_deferred_payload_retry_after_interruption(self, fail_number):
        _, directory, index = self.deferred_snapshot()
        first = self.row(index, "<first@example.test>")
        before = self.profile_tree()
        flaky, attempts = self.flaky_payload_temporary(fail_number)
        with flaky:
            try:
                first_read = outlook_local.read_message(directory, index, first["id"], 20_000)
            except OSError:
                first_read = None
        self.assertEqual(attempts["payload"], fail_number)
        if first_read is not None:
            self.assertNotEqual(first_read["body"]["status"], "cached_html")
        self.assertEqual(list((directory / "payloads").glob(".payload-*")), [])
        self.assertEqual(list((directory / "textpayloads").glob(".payload-*")), [])

        result = outlook_local.read_message(directory, index, first["id"], 20_000)
        raw = self.body_source.read_bytes()
        text, _ = outlook_local.decode_body(raw)
        self.assertEqual(result["body"]["status"], "cached_html")
        self.assertEqual(result["body"]["sha256"], outlook_local.sha(raw))
        self.assertEqual(result["body"]["text_sha256"], outlook_local.sha(
            text.encode("utf-16-le", errors="surrogatepass")
        ))
        self.assertIn(PREVIEW, result["body"]["text"])
        self.assertEqual(self.profile_tree(), before)

    def test_deferred_raw_payload_write_interruption_recovers_on_retry(self):
        self._assert_deferred_payload_retry_after_interruption(fail_number=1)

    def test_deferred_text_payload_write_interruption_recovers_on_retry(self):
        self._assert_deferred_payload_retry_after_interruption(fail_number=2)

    def test_deferred_partial_hash_named_payloads_are_repaired(self):
        _, directory, index = self.deferred_snapshot()
        first = self.row(index, "<first@example.test>")
        raw = self.body_source.read_bytes()
        text, _ = outlook_local.decode_body(raw)
        text_bytes = text.encode("utf-16-le", errors="surrogatepass")
        raw_path = directory / "payloads" / outlook_local.sha(raw)
        text_path = directory / "textpayloads" / outlook_local.sha(text_bytes)
        raw_path.write_bytes(raw[:max(1, len(raw) // 3)])
        text_path.write_bytes(text_bytes[:max(2, len(text_bytes) // 3)])
        raw_path.chmod(0o600)
        text_path.chmod(0o600)
        before = self.profile_tree()

        result = outlook_local.read_message(directory, index, first["id"], 20_000)

        self.assertEqual(result["body"]["status"], "cached_html")
        self.assertEqual(raw_path.read_bytes(), raw)
        self.assertEqual(text_path.read_bytes(), text_bytes)
        self.assertEqual(raw_path.stat().st_mode & 0o777, 0o400)
        self.assertEqual(text_path.stat().st_mode & 0o777, 0o400)
        self.assertEqual(result["body"]["sha256"], outlook_local.sha(raw))
        self.assertEqual(result["body"]["text_sha256"], outlook_local.sha(text_bytes))
        self.assertIn(PREVIEW, result["body"]["text"])
        self.assertEqual(self.profile_tree(), before)

    def test_deferred_unavailable_external_marker_returns_metadata_without_source_read(self):
        unsafe_reference = "~/Files/EFMData/../unsafe-body.html"
        header_key, body_key = 0xA21, 0xA22
        self.set_block_data(
            self.inline_header(header_key, "<unsafe-body@example.test>", PREVIEW, target=body_key),
            serialized_frame(202, body_key, unsafe_reference.encode("utf-16le") + b"\0\0"),
        )
        _, directory, index = self.deferred_snapshot()
        row = self.row(index, "<unsafe-body@example.test>")
        marker = row["body"]["deferred"]["external_fingerprints"][unsafe_reference]
        self.assertTrue(marker["unavailable"])
        self.assertIn("unsafe cache path", marker["error"])
        before = self.profile_tree()

        with patch.object(outlook_local, "stable_read",
                          side_effect=AssertionError("unsafe external source was read")):
            result = outlook_local.read_message(directory, index, row["id"], 20_000)

        self.assertEqual(result["subject"], row["subject"])
        self.assertEqual(result["body"]["status"], "unavailable")
        self.assertEqual(result["body"]["checks"][0]["path"], unsafe_reference)
        self.assertEqual(result["body"]["checks"][0]["status"], "unavailable")
        self.assertIn("unsafe cache path", result["body"]["checks"][0]["error"])
        self.assertEqual(self.profile_tree(), before)

    def test_deferred_small_source_copy_checks_scratch_space_before_write(self):
        before = self.profile_tree()
        required = self.source_cache.stat().st_size
        with patch.object(outlook_local.shutil, "disk_usage") as disk_usage:
            disk_usage.return_value.free = max(0, required - 1)
            with self.assertRaisesRegex(ValueError, "insufficient scratch space"):
                self.deferred_snapshot()
        self.assertFalse(any(self.state.glob(".building-*")))
        self.assertEqual(self.profile_tree(), before)

    def test_source_read_budget_debits_block_once_and_consumes_failed_read(self):
        _, directory, index = self.deferred_snapshot()
        offset = str(self.block.offset)
        size = index["source_block_bytes"][offset]
        cache = {}
        budget = {"remaining": size}
        with outlook_local.source_read_budget(budget):
            outlook_local._read_source_block(directory, index, self.block.offset, cache)
            self.assertEqual(budget["remaining"], 0)
            outlook_local._read_source_block(directory, index, self.block.offset, cache)
            self.assertEqual(budget["remaining"], 0)

        failed_budget = {"remaining": size}
        with patch.object(hxcontainer, "read_block_at", side_effect=ValueError("synthetic read failure")):
            with outlook_local.source_read_budget(failed_budget):
                with self.assertRaisesRegex(ValueError, "synthetic read failure"):
                    outlook_local._read_source_block(directory, index, self.block.offset, {})
        self.assertEqual(failed_budget["remaining"], 0)

    def test_decode_body_charges_non_gzip_decoded_bytes(self):
        raw = f"<html><body>{PREVIEW}</body></html>".encode()
        budget = {"remaining": len(raw) - 1}
        with outlook_local.source_read_budget(budget):
            with self.assertRaises(outlook_local.SourceBudgetExceeded):
                outlook_local.decode_body(raw)
        self.assertEqual(budget["remaining"], len(raw) - 1)

        budget = {"remaining": len(raw)}
        with outlook_local.source_read_budget(budget):
            text, encoding = outlook_local.decode_body(raw)
        self.assertEqual(encoding, "html")
        self.assertIn(PREVIEW, text)
        self.assertEqual(budget["remaining"], 0)

    def test_oversized_gzip_without_bundle_budget_keeps_prototype_error(self):
        raw = b"<html><body>" + b"x" * 64 + b"</body></html>"
        with patch.object(outlook_local, "MAX_BODY", 32):
            with self.assertRaisesRegex(ValueError, "inflated body exceeds prototype limit"):
                outlook_local.decode_body(gzip.compress(raw))

    def test_deferred_external_gzip_budget_bounds_inflation_without_sidecar(self):
        html = (
            f"<html><body>{GZIP_PREVIEW}"
            + "x" * (256 * 1024)
            + "</body></html>"
        ).encode()
        compressed = gzip.compress(html)
        self.assertLess(len(compressed), len(html))
        self.gzip_body_source.write_bytes(compressed)

        _, directory, index = self.deferred_snapshot()
        row = self.row(index, "<second@example.test>")
        block_size = index["source_block_bytes"][str(self.block.offset)]
        budget = {"remaining": block_size + len(compressed) + 1024}
        with outlook_local.source_read_budget(budget):
            with self.assertRaises(outlook_local.SourceBudgetExceeded):
                outlook_local.read_message(directory, index, row["id"], 20_000)
        self.assertEqual(budget["remaining"], 0)
        self.assertIsNone(outlook_local._read_deferred_cache(directory, row))

        budget = {"remaining": block_size + len(compressed) + len(html)}
        with outlook_local.source_read_budget(budget):
            result = outlook_local.read_message(directory, index, row["id"], 20_000)
        self.assertEqual(result["body"]["status"], "cached_html")
        self.assertIn(GZIP_PREVIEW, result["body"]["text"])

    def test_asset_changes_refresh_snapshot_without_hxstore_change(self):
        first_result, _, first_index = self.snapshot()
        first_row = self.row(first_index, "<first@example.test>")
        original_hxstore = self.source_cache.read_bytes()

        self.body_source.write_bytes(
            (
                "<html><body><p>"
                + PREVIEW
                + "。</p><p>资产清单变化后的正文。</p></body></html>"
            ).encode("utf-8")
        )
        with patch.object(outlook_local, "stable_read", wraps=outlook_local.stable_read) as stable:
            second_result, _, second_index = self.snapshot()
        self.assertGreaterEqual(stable.call_count, 1)
        second_row = self.row(second_index, "<first@example.test>")

        self.assertEqual(self.source_cache.read_bytes(), original_hxstore)
        self.assertNotEqual(first_result["snapshot"], second_result["snapshot"])
        self.assertEqual(first_index["source_sha256"], second_index["source_sha256"])
        self.assertNotEqual(
            first_index["asset_manifest_sha256"], second_index["asset_manifest_sha256"]
        )
        self.assertNotEqual(first_row["body"]["sha256"], second_row["body"]["sha256"])

        self.attachment_source.write_bytes(b"attachment-after-asset-refresh")
        with patch.object(outlook_local, "stable_read", wraps=outlook_local.stable_read) as stable:
            third_result, _, third_index = self.snapshot()
        self.assertGreaterEqual(stable.call_count, 1)
        third_row = self.row(third_index, "<first@example.test>")

        self.assertEqual(self.source_cache.read_bytes(), original_hxstore)
        self.assertNotEqual(second_result["snapshot"], third_result["snapshot"])
        self.assertEqual(second_index["source_sha256"], third_index["source_sha256"])
        self.assertNotEqual(
            second_index["asset_manifest_sha256"], third_index["asset_manifest_sha256"]
        )
        self.assertNotEqual(
            second_row["attachments"][0]["fingerprint"],
            third_row["attachments"][0]["fingerprint"],
        )

    def test_html_list_counter_difference_is_preview_verified(self):
        preview = "Dear all， 1. AIMS要求资产配置。"
        self.body_source.write_bytes(
            "<html><body><p>Dear all，</p><p>AIMS要求资产配置。</p></body></html>".encode(
                "utf-8"
            )
        )
        self.set_block_data(
            mail_frame(
                0x811,
                before=("list@example.com", "列表收件人"),
                after=("列表主题", "<list@example.test>", preview, "列表主题"),
                tail=verified_ticks(0x811, 946_684_880) + self.entry_first,
            ),
            registry_frame(0xF811, self.entry_first, (BODY_REFERENCE,)),
        )

        _, _, index = self.snapshot()
        row = self.row(index, "<list@example.test>")
        self.assertEqual(row["body"]["status"], "cached_html")
        self.assertTrue(row["body"]["preview_verified"])
        self.assertNotIn("preview_mismatch", {check["status"] for check in row["body"]["checks"]})

    def test_unrelated_html_body_remains_preview_mismatch(self):
        preview = "Dear all， 1. AIMS要求资产配置。"
        self.body_source.write_bytes(
            "<html><body><p>完全无关的正文。</p></body></html>".encode("utf-8")
        )
        self.set_block_data(
            mail_frame(
                0x812,
                before=("mismatch@example.com", "正文收件人"),
                after=("不匹配主题", "<mismatch@example.test>", preview, "不匹配主题"),
                tail=verified_ticks(0x812, 946_684_890) + self.entry_first,
            ),
            registry_frame(0xF812, self.entry_first, (BODY_REFERENCE,)),
        )

        _, _, index = self.snapshot()
        row = self.row(index, "<mismatch@example.test>")
        self.assertEqual(row["body"]["status"], "unavailable")
        self.assertIn(
            {"path": BODY_REFERENCE, "status": "preview_mismatch"},
            row["body"]["checks"],
        )

    def test_short_nonempty_preview_mismatch_is_not_cached(self):
        short_preview = "短预览"
        self.body_source.write_bytes(
            "<html><body><p>完全不包含该预览的正文。</p></body></html>".encode("utf-8")
        )
        self.set_block_data(
            mail_frame(
                0x813,
                before=("short@example.com", "短预览收件人"),
                after=("短预览主题", "<short-preview@example.test>", short_preview, "短预览主题"),
                tail=verified_ticks(0x813, 946_684_892) + self.entry_first,
            ),
            registry_frame(0xF813, self.entry_first, (BODY_REFERENCE,)),
        )

        _, _, index = self.snapshot()
        row = self.row(index, "<short-preview@example.test>")
        self.assertNotEqual(row["body"]["status"], "cached_html")
        self.assertIn(
            {"path": BODY_REFERENCE, "status": "preview_mismatch"},
            row["body"]["checks"],
        )

    def test_state_inside_profile_and_attachment_output_inside_profile_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "write destination"):
            outlook_local.make_snapshot(
                SimpleNamespace(
                    profile=str(self.profile),
                    state=str(self.profile / "private-state"),
                    from_date="2000-01-01",
                )
            )

        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        with self.assertRaisesRegex(ValueError, "write destination"):
            outlook_local.read_attachment(
                directory,
                index,
                first["id"],
                first["attachments"][0]["id"],
                self.profile / "exported.bin",
            )

    def test_state_marker_symlink_to_profile_is_rejected_without_profile_mutation(self):
        before = self.profile_tree()
        self.state.mkdir()
        (self.state / outlook_local.MARKER).symlink_to(self.source_cache)

        with self.assertRaisesRegex(ValueError, "symbolic links"):
            outlook_local.make_snapshot(
                SimpleNamespace(
                    profile=str(self.profile), state=str(self.state), from_date="2000-01-01"
                )
            )

        self.assertEqual(before, self.profile_tree())

    def test_state_root_symlink_is_rejected_without_profile_mutation(self):
        before = self.profile_tree()
        linked_state = self.state.parent / "state-link"
        linked_state.symlink_to(self.state, target_is_directory=True)

        with self.assertRaisesRegex(ValueError, "state root.*symbolic links"):
            outlook_local.make_snapshot(
                SimpleNamespace(
                    profile=str(self.profile), state=str(linked_state), from_date="2000-01-01"
                )
            )

        self.assertEqual(before, self.profile_tree())

    def test_hardlinked_state_marker_is_rejected(self):
        self.state.mkdir(mode=0o700)
        marker_source = self.state.parent / "marker-source"
        marker_source.write_text(outlook_local.MARKER_CONTENT)
        (self.state / outlook_local.MARKER).hardlink_to(marker_source)

        with self.assertRaisesRegex(ValueError, "hard link"):
            outlook_local.state_root(self.state, self.profile, create=True)

    def test_unsafe_state_directories_are_rejected_without_repair(self):
        _, directory, _ = self.snapshot()
        snapshot_root = self.state / "snapshots"
        payload_root = directory / "payloads"
        text_root = directory / "textpayloads"
        for path in (self.state, snapshot_root, directory, payload_root, text_root):
            path.chmod(0o755)

        with self.assertRaisesRegex(ValueError, "private permissions"):
            outlook_local.state_root(self.state, self.profile, create=True)
        self.assertEqual(self.state.stat().st_mode & 0o777, 0o755)
        self.assertEqual(snapshot_root.stat().st_mode & 0o777, 0o755)

        self.state.chmod(0o700)
        snapshot_root.chmod(0o700)
        with self.assertRaisesRegex(ValueError, "private permissions"):
            outlook_local.load_index(self.state)
        self.assertEqual(directory.stat().st_mode & 0o777, 0o755)

        directory.chmod(0o700)
        with self.assertRaisesRegex(ValueError, "private permissions"):
            outlook_local.load_index(self.state)
        self.assertEqual(payload_root.stat().st_mode & 0o777, 0o755)
        self.assertEqual(text_root.stat().st_mode & 0o777, 0o755)

    def test_foreign_state_owner_is_rejected(self):
        self.snapshot()
        self.state.chmod(0o755)
        current_uid = outlook_local.os.getuid()
        with patch.object(outlook_local.os, "getuid", return_value=current_uid + 1):
            with self.assertRaisesRegex(ValueError, "owner"):
                outlook_local.state_root(self.state, self.profile, create=True)
        self.assertEqual(self.state.stat().st_mode & 0o777, 0o755)

    def test_state_snapshots_symlink_to_profile_is_rejected_without_profile_mutation(self):
        before = self.profile_tree()
        self.state.mkdir(mode=0o700)
        (self.state / outlook_local.MARKER).write_text("outlook-local-prototype")
        (self.state / "snapshots").symlink_to(self.profile, target_is_directory=True)

        with self.assertRaisesRegex(ValueError, "symbolic links"):
            outlook_local.make_snapshot(
                SimpleNamespace(
                    profile=str(self.profile), state=str(self.state), from_date="2000-01-01"
                )
            )

        self.assertEqual(before, self.profile_tree())

    def test_current_symlink_to_profile_is_replaced_as_pointer_without_profile_mutation(self):
        before = self.profile_tree()
        self.state.mkdir(mode=0o700)
        (self.state / outlook_local.MARKER).write_text("outlook-local-prototype")
        (self.state / "CURRENT").symlink_to(self.source_cache)

        summary, _, _ = self.snapshot()

        current = self.state / "CURRENT"
        self.assertFalse(current.is_symlink())
        self.assertEqual(current.read_text().strip(), summary["snapshot"])
        self.assertEqual(before, self.profile_tree())

    def test_parent_traversal_and_symlink_escapes_are_rejected(self):
        outside = Path(self.temporary.name) / "outside"
        outside.mkdir()
        link = self.profile / "Files/Attachments/link"
        link.symlink_to(outside, target_is_directory=True)
        before = self.profile_tree()

        with self.assertRaisesRegex(ValueError, "unsafe cache path"):
            outlook_local.source_path(self.profile, "~/Files/../../outside/secret")
        with self.assertRaisesRegex(ValueError, "escapes profile"):
            outlook_local.source_path(self.profile, "~/Files/Attachments/link/secret")

        with self.assertRaisesRegex(ValueError, "symbolic link"):
            self.snapshot()
        self.assertEqual(before, self.profile_tree())
        link.unlink()

        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        output_link = Path(self.temporary.name) / "output-link"
        output_link.symlink_to(self.profile / "Files", target_is_directory=True)
        with self.assertRaisesRegex(ValueError, "write destination"):
            outlook_local.read_attachment(
                directory,
                index,
                first["id"],
                first["attachments"][0]["id"],
                output_link / "exported.bin",
            )

    def test_wrong_message_attachment_id_is_rejected(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        second = self.row(index, "<second@example.test>")

        with self.assertRaisesRegex(ValueError, "does not belong"):
            outlook_local.read_attachment(
                directory,
                index,
                second["id"],
                first["attachments"][0]["id"],
                self.output,
            )

    def test_changed_source_attachment_is_rejected(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        self.attachment_source.write_bytes(b"changed-after-snapshot")

        with self.assertRaisesRegex(ValueError, "missing or changed"):
            outlook_local.read_attachment(
                directory,
                index,
                first["id"],
                first["attachments"][0]["id"],
                self.output,
            )
        self.assertFalse(self.output.exists())

    def test_frozen_body_survives_source_removal_then_checksum_failure_is_reported(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        payload = directory / "payloads" / first["body"]["sha256"]
        self.body_source.unlink()

        message = outlook_local.read_message(directory, index, first["id"], 20_000)
        self.assertEqual(message["body"]["status"], "cached_html")
        self.assertIn("冻结后仍可读取", message["body"]["text"])

        payload.chmod(0o600)
        payload.write_bytes(b"corrupted-frozen-body")
        with self.assertRaisesRegex(ValueError, "checksum mismatch"):
            outlook_local.read_message(directory, index, first["id"], 20_000)

    def test_frozen_body_remains_paginable_after_source_changes(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        first_page = outlook_local.read_body_page(directory, index, first["id"], 0, 2)
        self.assertIsNotNone(first_page["next_offset"])

        self.body_source.write_bytes(b"<html><body>changed after freeze</body></html>")
        second_page = outlook_local.read_body_page(
            directory, index, first["id"], first_page["next_offset"], 20_000
        )
        self.assertIn("冻结后仍可读取", first_page["text"] + second_page["text"])

    def test_frozen_text_payload_checksum_failure_is_reported(self):
        _, directory, index = self.snapshot()
        first = self.row(index, "<first@example.test>")
        payload = directory / "textpayloads" / first["body"]["text_sha256"]
        payload.chmod(0o600)
        original = payload.read_bytes()
        payload.write_bytes(b"X" * len(original))
        with self.assertRaisesRegex(ValueError, "text payload checksum mismatch"):
            outlook_local.read_message(directory, index, first["id"], 20_000)

    def test_missing_body_without_reference_is_explicitly_not_linked(self):
        _, directory, index = self.snapshot()
        second = self.row(index, "<third@example.test>")

        self.assertEqual(second["body"]["status"], "not_linked")
        self.assertEqual(second["body"]["checks"], [])
        message = outlook_local.read_message(directory, index, second["id"], 20_000)
        self.assertEqual(message["body"]["status"], "not_linked")
        self.assertNotIn("text", message["body"])

    def test_missing_referenced_body_is_explicitly_not_cached(self):
        missing_reference = "~/Files/EFMData/missing-body.html"
        self.set_block_data(
            mail_frame(
                0xAEE,
                before=("missing@example.com", "缺失正文收件人"),
                after=("缺失正文主题", "<missing-body@example.test>", "缺失正文主题"),
                tail=verified_ticks(0xAEE, 946_684_895) + self.entry_first,
            ),
            registry_frame(0xFEE, self.entry_first, (missing_reference,)),
        )

        _, directory, index = self.snapshot()
        row = self.row(index, "<missing-body@example.test>")
        self.assertEqual(row["body"]["status"], "not_cached")
        self.assertEqual(
            row["body"]["checks"],
            [{"path": missing_reference, "status": "not_cached"}],
        )
        message = outlook_local.read_message(directory, index, row["id"], 20_000)
        self.assertEqual(message["body"]["status"], "not_cached")
        self.assertNotIn("text", message["body"])

    def test_cursor_cannot_cross_query_or_snapshot(self):
        _, _, index = self.snapshot()
        page_args = SimpleNamespace(
            query="",
            field="all",
            from_date=None,
            to_date=None,
            limit=1,
            cursor=None,
        )
        first_page = outlook_local.search(index, page_args)
        self.assertIsNotNone(first_page["next_cursor"])

        with self.assertRaisesRegex(ValueError, "another snapshot or query"):
            outlook_local.search(
                index,
                SimpleNamespace(
                    query="中文",
                    field="all",
                    from_date=None,
                    to_date=None,
                    limit=1,
                    cursor=first_page["next_cursor"],
                ),
            )

        other_snapshot = copy.copy(index)
        other_snapshot["snapshot"] = "a" * 24
        with self.assertRaisesRegex(ValueError, "another snapshot or query"):
            outlook_local.search(
                other_snapshot,
                SimpleNamespace(
                    query="",
                    field="all",
                    from_date=None,
                    to_date=None,
                    limit=1,
                    cursor=first_page["next_cursor"],
                ),
            )

    def test_to_date_before_snapshot_scope_is_rejected(self):
        _, _, index = self.snapshot()

        with self.assertRaisesRegex(ValueError, "scope"):
            outlook_local.search(
                index,
                SimpleNamespace(
                    query="",
                    field="all",
                    from_date=None,
                    to_date="1999-12-31",
                    limit=20,
                    cursor=None,
                ),
            )

    def test_latest_same_key_header_keeps_latest_tick_and_entry_but_supplements_subject(self):
        key = 0xB01
        stable_id = "<versioned@example.test>"
        old_entry = exchange_entry(801)
        latest_entry = exchange_entry(803)
        old = mail_frame(
            key,
            before=("versioned@example.com", "版本收件人"),
            after=("旧主题", stable_id, "旧主题"),
            tail=verified_ticks(key, 946_684_800) + old_entry,
        )
        latest = mail_frame(
            key,
            before=("versioned@example.com", "版本收件人"),
            after=(stable_id, "最新预览"),
            tail=verified_ticks(key, 946_684_800) + ticks(946_684_810) + latest_entry,
        )
        self.set_block_data(old, latest)

        summary, _, index = self.snapshot()
        row = self.row(index, stable_id)
        self.assertEqual(summary["messages"], 1)
        self.assertEqual(row["message_id"], stable_id)
        self.assertEqual(row["subject"], "旧主题")
        self.assertEqual(row["subject_status"], "paired_older_same_identity")
        self.assertEqual(row["sent_unix"], 946_684_800)
        self.assertEqual(row["latest_tick_unix"], 946_684_810)
        self.assertEqual(row["entry_ids"], [latest_entry.hex()])
        self.assertEqual(row["preview"], "最新预览")

    def test_same_key_different_message_ids_are_conservatively_excluded(self):
        key = 0xB02
        old = mail_frame(
            key,
            before=("versioned@example.com", "版本收件人"),
            after=("旧主题", "<versioned@example.test>", "旧主题"),
            tail=verified_ticks(key, 946_684_800) + exchange_entry(804),
        )
        latest_other_id = mail_frame(
            key,
            before=("versioned@example.com", "版本收件人"),
            after=("不应跨 MID 借用", "<stale@example.test>", "不应跨 MID 借用"),
            tail=verified_ticks(key, 946_684_800) + ticks(946_684_810) + exchange_entry(805),
        )
        self.set_block_data(old, latest_other_id)

        summary, _, index = self.snapshot()
        self.assertEqual(summary["messages"], 0)
        self.assertEqual(summary["identity_conflicts_excluded"], 1)
        self.assertEqual(index["items"], [])

    def test_same_identity_conflicting_verified_send_times_are_excluded(self):
        key = 0xB03
        message_id = "<time-conflict@example.test>"
        first = mail_frame(
            key,
            before=("time@example.com", "时间收件人"),
            after=("时间主题", message_id, "时间主题"),
            tail=verified_ticks(key, 946_684_800) + exchange_entry(806),
        )
        second = mail_frame(
            key,
            before=("time@example.com", "时间收件人"),
            after=("时间主题", message_id, "时间主题"),
            tail=verified_ticks(key, 946_684_810) + exchange_entry(807),
        )
        self.set_block_data(first, second)

        summary, _, index = self.snapshot()
        self.assertEqual(summary["messages"], 0)
        self.assertEqual(summary["time_conflicts_excluded"], 1)
        self.assertEqual(index["items"], [])

    def test_snapshot_without_verified_block_fails_explicitly(self):
        def no_verified_blocks(_data, stats=None):
            if stats is not None:
                stats.update(candidates=1, verified=0, rejected=1)
            return []

        self.iter_blocks_mock.side_effect = no_verified_blocks
        with self.assertRaisesRegex(ValueError, "verified"):
            self.snapshot()

    def test_snapshot_without_recognizable_header_fails_explicitly(self):
        self.set_block_data(serialized_frame(202, 0xB02, b"state-only"))

        with self.assertRaisesRegex(ValueError, "recognizable"):
            self.snapshot()

    def test_registry_entry_id_with_multiple_header_keys_is_not_associated(self):
        entry = exchange_entry(301)
        first = mail_frame(
            0x601,
            before=("one@example.com", "一号"),
            after=("冲突主题甲", "<conflict-one@example.test>", "冲突主题甲"),
            tail=verified_ticks(0x601, 946_684_840) + entry,
        )
        second = mail_frame(
            0x602,
            before=("two@example.com", "二号"),
            after=("冲突主题乙", "<conflict-two@example.test>", "冲突主题乙"),
            tail=verified_ticks(0x602, 946_684_850) + entry,
        )
        registry = registry_frame(0xF301, entry, (BODY_REFERENCE, ATTACHMENT_REFERENCE))
        self.set_block_data(first, second, registry)

        _, _, index = self.snapshot()
        first_row = self.row(index, "<conflict-one@example.test>")
        second_row = self.row(index, "<conflict-two@example.test>")
        self.assertEqual(index["summary"]["messages"], 2)
        self.assertEqual(index["summary"]["linked_attachments"], 0)
        self.assertEqual(first_row["body"]["status"], "not_linked")
        self.assertEqual(second_row["body"]["status"], "not_linked")
        self.assertEqual(first_row["attachments"], [])
        self.assertEqual(second_row["attachments"], [])

    def test_different_entry_ids_keep_same_numeric_attachment_names_with_their_owner(self):
        first_entry, second_entry = exchange_entry(401), exchange_entry(402)
        first_reference = "~/Files/Attachments/owner-one/00042.dat"
        second_reference = "~/Files/Attachments/owner-two/00042.dat"
        first_path = self.profile / "Files/Attachments/owner-one/00042.dat"
        second_path = self.profile / "Files/Attachments/owner-two/00042.dat"
        first_path.parent.mkdir(parents=True)
        second_path.parent.mkdir(parents=True)
        first_path.write_bytes(b"first-entry-attachment")
        second_path.write_bytes(b"second-entry-attachment")

        first = mail_frame(
            0x701,
            before=("one@example.com", "一号"),
            after=("数字文件甲", "<numeric-one@example.test>", "数字文件甲"),
            tail=verified_ticks(0x701, 946_684_860) + first_entry,
        )
        second = mail_frame(
            0x702,
            before=("two@example.com", "二号"),
            after=("数字文件乙", "<numeric-two@example.test>", "数字文件乙"),
            tail=verified_ticks(0x702, 946_684_870) + second_entry,
        )
        self.set_block_data(
            first,
            second,
            registry_frame(0xF401, first_entry, (first_reference,)),
            registry_frame(0xF402, second_entry, (second_reference,)),
        )

        _, _, index = self.snapshot()
        first_row = self.row(index, "<numeric-one@example.test>")
        second_row = self.row(index, "<numeric-two@example.test>")
        first_attachment, second_attachment = first_row["attachments"][0], second_row["attachments"][0]
        self.assertEqual(first_attachment["owner"], 0x701)
        self.assertEqual(second_attachment["owner"], 0x702)
        self.assertEqual(first_attachment["path"], first_reference)
        self.assertEqual(second_attachment["path"], second_reference)
        self.assertIsNone(first_attachment["name"])
        self.assertIsNone(second_attachment["name"])
        self.assertEqual(first_attachment["cache_name"], "00042.dat")
        self.assertEqual(second_attachment["cache_name"], "00042.dat")

    def test_registry_frame_with_multiple_attachment_paths_reads_each_path(self):
        entry = exchange_entry(501)
        references = (
            "~/Files/Attachments/registry-one/first.bin",
            "~/Files/Attachments/registry-two/second.bin",
        )
        expected = (b"first-registry-attachment", b"second-registry-attachment")
        for reference, data in zip(references, expected):
            source = self.profile / Path(reference[2:])
            source.parent.mkdir(parents=True, exist_ok=True)
            source.write_bytes(data)

        key = 0x751
        self.set_block_data(
            mail_frame(
                key,
                before=("registry@example.com", "Registry sender"),
                after=("Registry subject", "<registry@example.test>", "Registry subject"),
                tail=verified_ticks(key, 946_684_875) + entry,
            ),
            registry_frame(0xF751, entry, references),
        )

        _, directory, index = self.snapshot()
        row = self.row(index, "<registry@example.test>")
        attachments = row["attachments"]
        self.assertEqual([attachment["path"] for attachment in attachments], list(references))
        self.assertEqual(len({attachment["id"] for attachment in attachments}), 2)

        for attachment, data, name in zip(attachments, expected, ("first.bin", "second.bin")):
            output = self.profile.parent / f"exported-{name}"
            result = outlook_local.read_attachment(directory, index, row["id"], attachment["id"], output)
            self.assertEqual(output.read_bytes(), data)
            self.assertEqual(result["name"], None)

    def test_entry_only_attachment_is_resolved_by_unique_entry_id(self):
        entry = exchange_entry(901)
        reference = "~/Files/Attachments/entry-only/entry.bin"
        source = self.profile / "Files/Attachments/entry-only/entry.bin"
        source.parent.mkdir(parents=True)
        source.write_bytes(b"entry-only-attachment")
        key = 0xC01
        self.set_block_data(
            mail_frame(
                key,
                before=("entry-only@example.com", "Entry Only"),
                after=("Entry Only Subject", "<entry-only@example.test>", "Entry Only Subject"),
                tail=verified_ticks(key, 946_684_900) + entry,
            ),
            attachment_frame(0xC02, None, reference, entry=entry, name="ignored-display-name.bin"),
        )

        _, _, index = self.snapshot()
        row = self.row(index, "<entry-only@example.test>")
        self.assertEqual(len(row["attachments"]), 1)
        attachment = row["attachments"][0]
        self.assertEqual(attachment["owner"], key)
        self.assertEqual(attachment["path"], reference)
        self.assertEqual(attachment["status"], "cached_at_snapshot")

    def test_typed_attachment_owner_conflicting_with_entry_id_is_rejected(self):
        first_entry, second_entry = exchange_entry(911), exchange_entry(912)
        first_key, second_key = 0xC11, 0xC12
        reference = "~/Files/Attachments/conflict/typed-owner.bin"
        source = self.profile / "Files/Attachments/conflict/typed-owner.bin"
        source.parent.mkdir(parents=True)
        source.write_bytes(b"conflicting-owner")
        unknown_reference = "~/Files/Attachments/conflict/unknown-owner.bin"
        unknown_source = self.profile / "Files/Attachments/conflict/unknown-owner.bin"
        unknown_source.write_bytes(b"conflicting-unknown-owner")
        self.set_block_data(
            mail_frame(
                first_key,
                before=("first-owner@example.com", "甲"),
                after=("第一封", "<first-owner@example.test>", "第一封"),
                tail=verified_ticks(first_key, 946_684_910) + first_entry,
            ),
            mail_frame(
                second_key,
                before=("second-owner@example.com", "乙"),
                after=("第二封", "<second-owner@example.test>", "第二封"),
                tail=verified_ticks(second_key, 946_684_920) + second_entry,
            ),
            attachment_frame(0xC13, first_key, reference, entry=second_entry),
            attachment_frame(0xC14, 999, unknown_reference, entry=second_entry),
        )

        summary, _, index = self.snapshot()
        first_row = self.row(index, "<first-owner@example.test>")
        second_row = self.row(index, "<second-owner@example.test>")
        self.assertEqual(summary["attachment_conflicts_excluded"], 2)
        self.assertEqual(first_row["attachments"], [])
        self.assertEqual(second_row["attachments"], [])

    def test_type_202_only_attachment_owner_is_used_without_entry_owner(self):
        reference = "~/Files/Attachments/type-202-only/type-202-only.bin"
        source = self.profile / "Files/Attachments/type-202-only/type-202-only.bin"
        source.parent.mkdir(parents=True)
        source.write_bytes(b"type-202-only")
        key = 0xC21
        self.set_block_data(
            mail_frame(
                key,
                before=("type-202-only@example.com", "202 only"),
                after=("Type 202 only", "<type-202-only@example.test>", "Type 202 only"),
                tail=verified_ticks(key, 946_684_925),
            ),
            attachment_frame(0xC22, None, reference, owner_202=key),
        )

        _, _, index = self.snapshot()
        row = self.row(index, "<type-202-only@example.test>")
        self.assertEqual(len(row["attachments"]), 1)
        self.assertEqual(row["attachments"][0]["owner"], key)
        self.assertEqual(row["attachments"][0]["path"], reference)

    def test_type_202_attachment_owner_conflicting_with_entry_id_is_rejected(self):
        first_entry, second_entry = exchange_entry(921), exchange_entry(922)
        first_key, second_key = 0xC31, 0xC32
        reference = "~/Files/Attachments/conflict/type-202-owner.bin"
        source = self.profile / "Files/Attachments/conflict/type-202-owner.bin"
        source.parent.mkdir(parents=True)
        source.write_bytes(b"conflicting-type-202-owner")
        self.set_block_data(
            mail_frame(
                first_key,
                before=("first-202@example.com", "甲"),
                after=("第一封", "<first-202@example.test>", "第一封"),
                tail=verified_ticks(first_key, 946_684_930) + first_entry,
            ),
            mail_frame(
                second_key,
                before=("second-202@example.com", "乙"),
                after=("第二封", "<second-202@example.test>", "第二封"),
                tail=verified_ticks(second_key, 946_684_940) + second_entry,
            ),
            attachment_frame(0xC33, None, reference, entry=first_entry, owner_202=second_key),
        )

        summary, _, index = self.snapshot()
        first_row = self.row(index, "<first-202@example.test>")
        second_row = self.row(index, "<second-202@example.test>")
        self.assertEqual(summary["attachment_conflicts_excluded"], 1)
        self.assertEqual(first_row["attachments"], [])
        self.assertEqual(second_row["attachments"], [])

    @staticmethod
    def inline_header(key, message_id, preview, target=None):
        reference = struct.pack("<H", 202) + b"\0" * 8 + struct.pack("<Q", key if target is None else target)
        return mail_frame(key, before=("inline@example.com", "测试发送人"),
                          after=("内嵌中文主题", message_id, preview, "内嵌中文主题"),
                          tail=verified_ticks(key, 946_684_930) + reference)

    def test_inline_full_body_uses_declared_field_and_preserves_source(self):
        full = f"<html><body>{PREVIEW}。完整历史正文🧪。</body></html>".encode()
        unique = "<html><body>只有本次回复。</body></html>".encode()
        self.set_block_data(self.inline_header(0x901, "<inline@example.test>", PREVIEW),
                            inline_body_frame(0x901, full, unique_html=unique))
        before = self.profile_tree()
        summary, directory, index = self.snapshot()
        row = self.row(index, "<inline@example.test>")
        self.assertEqual(row["body"]["status"], "cached_html")
        self.assertEqual(row["body"]["storage"], "hxstore_inline")
        self.assertEqual(row["body"]["sha256"], hashlib.sha256(full).hexdigest())
        self.assertTrue(row["body"]["preview_verified"])
        self.assertEqual(summary["body_diagnoses"], {"hxstore_inline_recovered": 1})
        result = outlook_local.read_message(directory, index, row["id"], 20000)
        self.assertIn("完整历史正文🧪", result["body"]["text"])
        self.assertNotIn("只有本次回复", result["body"]["text"])
        self.assertEqual(self.profile_tree(), before)

    def test_deferred_adjacent_cross_key_inline_bodies_use_real_frozen_source(self):
        # The normal fixture uses a synthetic block reader for fast parser
        # tests. Stop both patches here so deferred recovery exercises the
        # bounded HxStore reader against the private source.hxd copy.
        self.iter_blocks_patch.stop()
        self.read_block_patch.stop()
        try:
            first_key, first_body_key = 0xA11, 0xA12
            second_key, second_body_key = 0xA13, 0xA14
            first_preview = "甲邮件正文预览"
            second_preview = "乙邮件正文预览"
            first_html = f"<html><body>{first_preview}。甲邮件完整正文。</body></html>"
            second_html = f"<html><body>{second_preview}。乙邮件完整正文。</body></html>"
            block_data = b"".join((
                self.inline_header(first_key, "<deferred-first@example.test>",
                                   first_preview, target=first_body_key),
                inline_body_frame(first_body_key, first_html),
                self.inline_header(second_key, "<deferred-second@example.test>",
                                   second_preview, target=second_body_key),
                inline_body_frame(second_body_key, second_html),
            ))
            self.source_cache.write_bytes(store_header() + make_block(block_data))
            before = self.profile_tree()
            summary = outlook_local.make_snapshot(SimpleNamespace(
                profile=str(self.profile), state=str(self.state), from_date="2000-01-01",
                freeze_bodies=False,
            ))
            directory, index = outlook_local.load_index(str(self.state))
            self.assertEqual(summary["body_mode"], "deferred")
            first = self.row(index, "<deferred-first@example.test>")
            second = self.row(index, "<deferred-second@example.test>")
            self.assertEqual(first["body"]["status"], "deferred")
            self.assertEqual(second["body"]["status"], "deferred")
            self.assertEqual(first["body"]["diagnosis"]["inline_candidates"], 1)
            self.assertEqual(second["body"]["diagnosis"]["inline_candidates"], 1)
            self.assertEqual((directory / "source.hxd").read_bytes(), self.source_cache.read_bytes())

            with patch.object(outlook_local, "freeze_body", wraps=outlook_local.freeze_body) as freeze:
                first_read = outlook_local.read_message(directory, index, first["id"], 20_000)
                first_again = outlook_local.read_message(directory, index, first["id"], 20_000)
                second_read = outlook_local.read_message(directory, index, second["id"], 20_000)
            self.assertEqual(freeze.call_count, 2)
            self.assertIn("甲邮件完整正文", first_read["body"]["text"])
            self.assertNotIn("乙邮件完整正文", first_read["body"]["text"])
            self.assertEqual(first_again["body"]["text"], first_read["body"]["text"])
            self.assertIn("乙邮件完整正文", second_read["body"]["text"])
            self.assertNotIn("甲邮件完整正文", second_read["body"]["text"])
            self.assertEqual(self.profile_tree(), before)
        finally:
            self.iter_blocks_patch.start()
            self.read_block_patch.start()

    def test_corrupt_gzip_inline_body_is_unavailable_without_aborting_snapshot(self):
        corrupted = bytearray(gzip.compress(f"<html><body>{PREVIEW}</body></html>".encode()))
        corrupted[10] ^= 0xFF
        key = 0x909
        self.set_block_data(
            self.inline_header(key, "<corrupt-gzip-inline@example.test>", PREVIEW),
            inline_body_frame(key, bytes(corrupted)),
        )

        _, _, index = self.snapshot()
        row = self.row(index, "<corrupt-gzip-inline@example.test>")
        self.assertEqual(row["body"]["status"], "unavailable")
        self.assertEqual(row["body"]["diagnosis"]["code"], "body_candidates_rejected")
        self.assertEqual(row["body"]["checks"][0]["status"], "unavailable")

    def test_cross_key_inline_body_follows_declared_target(self):
        body_key = 0x90B
        full = f"<html><body>{PREVIEW}。跨对象的完整正文。</body></html>".encode()
        self.set_block_data(
            self.inline_header(0x90A, "<cross-key-inline@example.test>", PREVIEW, target=body_key),
            inline_body_frame(body_key, full),
        )

        _, directory, index = self.snapshot()
        row = self.row(index, "<cross-key-inline@example.test>")
        self.assertEqual(row["body"]["status"], "cached_html")
        self.assertEqual(row["body"]["storage"], "hxstore_inline")
        result = outlook_local.read_message(directory, index, row["id"], 20_000)
        self.assertIn("跨对象的完整正文", result["body"]["text"])

    def test_missing_body_object_and_unparsed_object_have_distinct_diagnoses(self):
        header = self.inline_header(0x902, "<missing-object@example.test>", PREVIEW)
        self.set_block_data(header)
        _, _, index = self.snapshot()
        row = self.row(index, "<missing-object@example.test>")
        self.assertEqual(row["body"]["status"], "not_cached")
        self.assertEqual(row["body"]["diagnosis"]["code"], "referenced_body_object_absent")
        self.assertEqual(row["body"]["body_reference"]["target_key"], 0x902)

        self.source_cache.write_bytes(b"synthetic snapshot with unparsed body object")
        self.set_block_data(header, serialized_frame(202, 0x902, b"unknown body layout"))
        _, _, index = self.snapshot()
        row = self.row(index, "<missing-object@example.test>")
        self.assertEqual(row["body"]["status"], "unavailable")
        self.assertEqual(row["body"]["diagnosis"]["code"], "body_object_layout_or_owner_unresolved")
        self.assertEqual(row["body"]["diagnosis"]["body_object_frames"], 1)

    def test_inline_body_needs_both_header_target_and_body_owner(self):
        body = f"<html><body>{PREVIEW}</body></html>".encode()
        self.set_block_data(self.inline_header(0x903, "<wrong-target@example.test>", PREVIEW, target=0x904),
                            inline_body_frame(0x903, body),
                            self.inline_header(0x904, "<wrong-owner@example.test>", PREVIEW),
                            inline_body_frame(0x904, body, owner=0x903))
        _, directory, index = self.snapshot()
        for row in index["items"]:
            self.assertEqual(row["body"]["status"], "unavailable")
            self.assertEqual(row["body"]["diagnosis"]["code"], "body_object_layout_or_owner_unresolved")
        self.assertEqual(list(directory.glob(".inline-*")), [])

    def test_inline_preview_mismatch_is_rejected_even_with_matching_owner(self):
        self.set_block_data(self.inline_header(0x905, "<inline-mismatch@example.test>", PREVIEW),
                            inline_body_frame(0x905, b"<html><body>Different message</body></html>"))
        _, _, index = self.snapshot()
        row = self.row(index, "<inline-mismatch@example.test>")
        self.assertEqual(row["body"]["status"], "unavailable")
        self.assertEqual(row["body"]["diagnosis"]["code"], "body_candidates_rejected")
        self.assertEqual(row["body"]["checks"][0]["status"], "preview_mismatch")

    def test_adjacent_inline_bodies_keep_each_messages_full_text(self):
        first_preview, second_preview = "甲邮件的正文开头", "乙邮件的正文开头"
        self.set_block_data(
            self.inline_header(0x906, "<inline-first@example.test>", first_preview),
            inline_body_frame(0x907, f"<html><body>{second_preview}。乙全文。</body></html>".encode()),
            self.inline_header(0x907, "<inline-second@example.test>", second_preview),
            inline_body_frame(0x906, f"<html><body>{first_preview}。甲全文。</body></html>".encode()),
        )
        _, directory, index = self.snapshot()
        for message_id, own, other in (("<inline-first@example.test>", "甲全文", "乙全文"),
                                       ("<inline-second@example.test>", "乙全文", "甲全文")):
            row = self.row(index, message_id)
            result = outlook_local.read_message(directory, index, row["id"], 20000)
            self.assertIn(own, result["body"]["text"])
            self.assertNotIn(other, result["body"]["text"])

    def test_files_root_symbolic_link_is_rejected(self):
        files = self.profile / "Files"
        saved = self.profile / "SavedFiles"
        files.rename(saved)
        files.symlink_to(saved, target_is_directory=True)
        before = self.profile_tree()
        with self.assertRaisesRegex(ValueError, "symbolic link"):
            self.snapshot()
        self.assertEqual(before, self.profile_tree())

    def test_missing_file_does_not_hide_present_unparsed_body_object(self):
        key = 0x908
        self.body_source.unlink()
        header = self.inline_header(key, "<unparsed-and-missing@example.test>", PREVIEW)
        # Retain an unparsed object version alongside a known file reference.
        file_object = serialized_frame(202, key, BODY_REFERENCE.encode("utf-16le") + b"\0\0")
        self.set_block_data(header, file_object, serialized_frame(202, key, b"unknown layout"))
        _, _, index = self.snapshot()
        row = self.row(index, "<unparsed-and-missing@example.test>")
        self.assertEqual(row["body"]["status"], "unavailable")
        self.assertEqual(row["body"]["diagnosis"]["code"], "body_object_layout_or_owner_unresolved")

        self.source_cache.write_bytes(b"only recognized external body reference remains")
        self.set_block_data(header, file_object)
        _, _, index = self.snapshot()
        row = self.row(index, "<unparsed-and-missing@example.test>")
        self.assertEqual(row["body"]["status"], "not_cached")
        self.assertEqual(row["body"]["diagnosis"]["code"], "referenced_file_absent")

    def test_html_with_binary_metadata_prefix_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "NUL"):
            outlook_local.decode_body(b"j\0o\0i\0<html><body>plausible body</body></html>")


if __name__ == "__main__":
    unittest.main()
