import mmap
import random
import struct
import sys
import tempfile
import unittest
import zlib
from pathlib import Path


sys.path.insert(0, str(Path(__file__).parent))
from hxcontainer import (  # noqa: E402
    BLOCK_HEADER_SIZE,
    MAX_INFLATED,
    Block,
    decode_lz4,
    iter_blocks,
    read_block_at,
)


def literal_block(data: bytes) -> bytes:
    length = len(data)
    if length < 15:
        return bytes((length << 4,)) + data
    extension = bytearray((0xF0,))
    length -= 15
    while length >= 255:
        extension.append(255)
        length -= 255
    extension.append(length)
    return bytes(extension) + data


def store_header(version: int = ord("i"), page_size: int = 4096) -> bytes:
    result = bytearray(0x40)
    result[:8] = b"Nostromo"
    result[8] = version
    struct.pack_into("<Q", result, 0x38, page_size)
    return bytes(result)


def make_block(
    data: bytes,
    *,
    kind: int = 8,
    payload: bytes | None = None,
    inflated_length: int | None = None,
    corrupt_header: bool = False,
    corrupt_body: bool = False,
) -> bytes:
    if payload is None:
        payload = literal_block(data)
    if inflated_length is None:
        inflated_length = len(data)
    result = bytearray(BLOCK_HEADER_SIZE + len(payload))
    struct.pack_into("<Q", result, 0x08, 0x5D0245643B706A05)
    struct.pack_into("<III", result, 0x10, kind, len(payload), inflated_length)
    struct.pack_into("<I", result, 0x1C, 4)
    result[BLOCK_HEADER_SIZE:] = payload
    body_crc = zlib.crc32(result[8:]) & 0xFFFFFFFF
    struct.pack_into("<I", result, 0x04, body_crc ^ (1 if corrupt_body else 0))
    header_crc = zlib.crc32(result[4:0x20]) & 0xFFFFFFFF
    struct.pack_into("<I", result, 0x00, header_crc ^ (1 if corrupt_header else 0))
    return bytes(result)


def reference_decode_lz4(src: bytes, expected: int) -> bytes:
    """The previous byte-at-a-time match decoder used for differential tests."""
    source = memoryview(src)
    output = bytearray()
    index = 0
    while len(output) < expected:
        token = source[index]
        index += 1
        literal_length = token >> 4
        if literal_length == 15:
            while True:
                extension = source[index]
                index += 1
                literal_length += extension
                if extension != 255:
                    break
        output.extend(source[index : index + literal_length])
        index += literal_length
        if len(output) == expected:
            return bytes(output)
        distance = source[index] | (source[index + 1] << 8)
        index += 2
        match_length = (token & 15) + 4
        if (token & 15) == 15:
            while True:
                extension = source[index]
                index += 1
                match_length += extension
                if extension != 255:
                    break
        match_start = len(output) - distance
        for copy_index in range(match_length):
            output.append(output[match_start + copy_index])
    return bytes(output)


def encoded_length(length: int) -> bytes:
    """Encode the continuation bytes after a 15-valued LZ4 length nibble."""
    remaining = length - 15
    result = bytearray()
    while remaining >= 255:
        result.append(255)
        remaining -= 255
    result.append(remaining)
    return bytes(result)


def literal_match_block(prefix: bytes, distance: int, match_length: int) -> tuple[bytes, int]:
    literal_nibble = min(len(prefix), 15)
    match_nibble = min(match_length - 4, 15)
    payload = bytearray(((literal_nibble << 4) | match_nibble,))
    if len(prefix) >= 15:
        payload.extend(encoded_length(len(prefix)))
    payload.extend(prefix)
    payload.extend((distance & 255, distance >> 8))
    if match_length >= 19:
        payload.extend(encoded_length(match_length - 4))
    return bytes(payload), len(prefix) + match_length


class HxContainerTests(unittest.TestCase):
    def test_targeted_block_read_skips_unrelated_corrupt_block(self):
        first = make_block(b'broken', corrupt_body=True)
        second = make_block('中文正文'.encode())
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / 'HxStore.hxd'
            source.write_bytes(store_header() + first + second)
            before = source.read_bytes()
            offset = 0x40 + len(first)
            self.assertEqual(read_block_at(source, offset), Block(offset, 8, '中文正文'.encode()))
            self.assertEqual(source.read_bytes(), before)
            with self.assertRaises(ValueError):
                read_block_at(source, 0x40)
            source.write_bytes(before[:-1])
            with self.assertRaisesRegex(ValueError, 'truncated'):
                read_block_at(source, offset)

    def test_targeted_block_read_rejects_symlink_and_invalid_offset(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / 'HxStore.hxd'
            source.write_bytes(store_header() + make_block(b'body'))
            link = source.with_name('link')
            link.symlink_to(source)
            with self.assertRaises(OSError):
                read_block_at(link, 0x40)
            with self.assertRaises(ValueError):
                read_block_at(source, -1)

    def test_utf8_and_binary_bytes_round_trip(self):
        value = "主题：你好，世界 🌏".encode("utf-8") + b"\x00\xff\x10"
        self.assertEqual(decode_lz4(literal_block(value), len(value)), value)

    def test_overlapping_match_is_copied_with_repeated_pattern(self):
        self.assertEqual(decode_lz4(bytes((0x10, ord("x"), 1, 0)), 5), b"xxxxx")

    def test_non_overlapping_match_is_copied_as_one_slice(self):
        payload, expected = literal_match_block(b"abcdefgh", 8, 8)
        self.assertEqual(decode_lz4(payload, expected), b"abcdefghabcdefgh")

    def test_randomized_matches_match_the_previous_decoder(self):
        generator = random.Random(0x4858)
        for _ in range(400):
            prefix = bytes(generator.randrange(256) for _ in range(generator.randrange(1, 96)))
            distance = generator.randrange(1, len(prefix) + 1)
            match_length = generator.randrange(4, 512)
            payload, expected = literal_match_block(prefix, distance, match_length)
            self.assertEqual(decode_lz4(payload, expected),
                             reference_decode_lz4(payload, expected))

    def test_corrupt_distance_and_short_input_raise(self):
        with self.assertRaises(ValueError):
            decode_lz4(bytes((0x10, ord("x"), 0, 0)), 5)
        with self.assertRaises(ValueError):
            decode_lz4(bytes((0x10, ord("x"), 1)), 5)

    def test_trailing_input_is_rejected(self):
        value = b"hello"
        with self.assertRaises(ValueError):
            decode_lz4(literal_block(value) + b"\x00", len(value))

    def test_output_overflow_and_oversized_declaration_are_rejected(self):
        # One literal followed by a six-byte match would produce seven bytes.
        with self.assertRaises(ValueError):
            decode_lz4(bytes((0x10, ord("x"), 1, 0)), 4)
        with self.assertRaises(ValueError):
            decode_lz4(b"", MAX_INFLATED + 1)

    def test_double_crc_and_stats(self):
        good = make_block(b"ok")
        bad_header = make_block(b"bad", corrupt_header=True)
        bad_body = make_block(b"bad", corrupt_body=True)
        stats = {}
        blocks = list(iter_blocks(store_header() + good + bad_header + bad_body, stats))
        self.assertEqual(blocks, [Block(0x40, 8, b"ok")])
        self.assertEqual(stats, {"candidates": 3, "verified": 1, "rejected": 2})

    def test_unsupported_kind_is_rejected(self):
        stats = {}
        self.assertEqual(list(iter_blocks(store_header() + make_block(b"x", kind=99), stats)), [])
        self.assertEqual(stats, {"candidates": 1, "verified": 0, "rejected": 1})

    def test_rejected_candidate_recovers_following_valid_block(self):
        bad = make_block(b"bad", corrupt_body=True)
        good = make_block(b"good")
        stats = {}
        blocks = list(iter_blocks(store_header() + bad + good, stats))
        self.assertEqual(blocks, [Block(0x40 + len(bad), 8, b"good")])
        self.assertEqual(stats, {"candidates": 2, "verified": 1, "rejected": 1})

    def test_nested_valid_magic_inside_outer_payload_is_not_top_level(self):
        inner = make_block(b"inner")
        outer = make_block(inner)
        stats = {}
        blocks = list(iter_blocks(store_header() + outer, stats))
        self.assertEqual(blocks, [Block(0x40, 8, inner)])
        self.assertEqual(stats, {"candidates": 1, "verified": 1, "rejected": 0})

    def test_unknown_version_and_page_size_fail_before_scan(self):
        with self.assertRaises(ValueError):
            iter_blocks(store_header(version=ord("h")))
        with self.assertRaises(ValueError):
            iter_blocks(store_header(page_size=8192))

    def test_corrupt_block_is_skipped(self):
        stats = {}
        value = "中文记录".encode("utf-8")
        blocks = list(iter_blocks(store_header() + make_block(value, corrupt_body=True), stats))
        self.assertEqual(blocks, [])
        self.assertEqual(stats["rejected"], 1)

    def test_mmap_source_is_scanned_without_materializing_hxstore(self):
        value = store_header() + make_block(b"mapped")
        with tempfile.TemporaryFile() as handle:
            handle.write(value)
            handle.flush()
            mapped = mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ)
            try:
                self.assertEqual(list(iter_blocks(mapped)), [Block(0x40, 8, b"mapped")])
            finally:
                mapped.close()

    def test_sparse_source_over_two_gib_maps_and_parses(self):
        header = store_header()
        block = make_block(b"sparse")
        sparse_size = 2 * 1024**3 + 1
        with tempfile.TemporaryFile() as handle:
            handle.write(header)
            handle.truncate(sparse_size)
            offset = sparse_size - len(block)
            handle.seek(offset)
            handle.write(block)
            handle.flush()
            mapped = mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ)
            blocks = iter_blocks(mapped)
            try:
                self.assertEqual(next(blocks), Block(offset, 8, b"sparse"))
            finally:
                blocks.close()
                mapped.close()


if __name__ == "__main__":
    unittest.main()
