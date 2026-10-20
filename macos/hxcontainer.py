"""Read-only parser for the HxStore Nostromo i block container.

The module accepts caller-provided bytes-like snapshots, including a private
read-only mmap.  It does not open or write any Outlook cache file; callers
provide a store snapshot and receive verified decompressed blocks.
"""

from __future__ import annotations

from dataclasses import dataclass
import mmap
import zlib


MAGIC = b"Nostromo"
BLOCK_MAGIC = (0x5D0245643B706A05).to_bytes(8, "little")
PAGE_SIZE = 4096
MAX_INFLATED = 32 * 1024 * 1024
BLOCK_HEADER_SIZE = 0x28
SUPPORTED_KINDS = frozenset((8, 16))


@dataclass(frozen=True)
class Block:
    offset: int
    kind: int
    data: bytes


def _bad(message: str) -> ValueError:
    return ValueError(message)


def _read_u32(data: bytes, offset: int) -> int:
    return int.from_bytes(data[offset : offset + 4], "little")


def _read_u64(data: bytes, offset: int) -> int:
    return int.from_bytes(data[offset : offset + 8], "little")


def _read_length(src: bytes, index: int, base: int, label: str) -> tuple[int, int]:
    """Read an LZ4 255-continuation length without crossing ``src``."""

    length = base
    while True:
        if index >= len(src):
            raise _bad(f"truncated {label} length")
        extension = src[index]
        index += 1
        length += extension
        if extension != 0xFF:
            return length, index


def decode_lz4(src: bytes, expected: int) -> bytes:
    """Decode one raw LZ4 block and require exact input/output consumption.

    ``src`` has no LZ4 frame header or checksum.  A final sequence may contain
    literals only; once the declared output length is reached every source
    byte must have been consumed.  Invalid input raises ``ValueError``.
    """

    if not isinstance(src, (bytes, bytearray, memoryview)):
        raise _bad("src must be a bytes-like object")
    if not isinstance(expected, int) or isinstance(expected, bool):
        raise _bad("expected must be an integer")
    if expected < 0 or expected > MAX_INFLATED:
        raise _bad(f"inflated length outside 0..{MAX_INFLATED}")

    # A memoryview avoids a copy when iter_blocks passes a payload slice, while
    # keeping all indexing and bounds checks explicit.
    source = memoryview(src)
    if expected == 0:
        if len(source) != 0:
            raise _bad("trailing input after empty output")
        return b""

    out = bytearray()
    index = 0
    while len(out) < expected:
        if index >= len(source):
            raise _bad("truncated LZ4 sequence")

        token = source[index]
        index += 1
        literal_length = token >> 4
        if literal_length == 15:
            literal_length, index = _read_length(source, index, literal_length, "literal")
        if literal_length > len(source) - index:
            raise _bad("literal run exceeds input")
        if literal_length > expected - len(out):
            raise _bad("literal run exceeds expected output")

        out.extend(source[index : index + literal_length])
        index += literal_length

        # The final sequence is literals-only.  A nonzero low token nibble is
        # ignored here, as it has no match bytes associated with it.
        if len(out) == expected:
            if index != len(source):
                raise _bad("trailing input after final literals")
            return bytes(out)

        if len(source) - index < 2:
            raise _bad("truncated match distance")
        distance = source[index] | (source[index + 1] << 8)
        index += 2
        if distance == 0 or distance > len(out):
            raise _bad("invalid LZ4 match distance")

        match_length = (token & 0x0F) + 4
        if (token & 0x0F) == 0x0F:
            match_length, index = _read_length(source, index, match_length, "match")
        if match_length > expected - len(out):
            raise _bad("match exceeds expected output")

        # Copy one byte at a time so legal overlapping matches can expand runs.
        match_start = len(out) - distance
        for copy_index in range(match_length):
            out.append(out[match_start + copy_index])
        if len(out) == expected:
            if index != len(source):
                raise _bad("trailing input after final match")
            return bytes(out)

    # The loop returns as soon as output reaches ``expected``.  Keep this guard
    # for clarity if its control flow is changed later.
    raise _bad("LZ4 output length mismatch")


def _check_store_header(data: bytes) -> None:
    if not isinstance(data, (bytes, bytearray, memoryview, mmap.mmap)):
        raise _bad("data must be a bytes-like object")
    if len(data) < 0x40:
        raise _bad("file is too small to hold a Nostromo header")
    if bytes(data[:8]) != MAGIC:
        raise _bad("not an HxStore: missing Nostromo magic")
    version = data[8]
    if version != ord("i"):
        raise _bad(f"unsupported Nostromo version 0x{version:02x}; expected 0x69 ('i')")
    page_size = _read_u64(data, 0x38)
    if page_size != PAGE_SIZE:
        raise _bad(f"unsupported HxStore page size {page_size}; expected {PAGE_SIZE}")


def iter_blocks(data: bytes, stats: dict | None = None):
    """Yield verified HxStore blocks from a bytes-like or mmap snapshot.

    Candidate headers are located by their block magic.  Every candidate is
    independently bounds checked, double-CRC checked, kind checked, and
    strictly decompressed.  A bad or unsupported candidate increments
    ``stats['rejected']`` and is skipped.
    """

    _check_store_header(data)
    # mmap.find and memoryview slices let a private frozen source be scanned
    # without materialising another copy of the entire HxStore.
    searchable = data if isinstance(data, (bytes, bytearray, mmap.mmap)) else bytes(data)
    source = memoryview(searchable)
    if stats is not None:
        stats["candidates"] = 0
        stats["verified"] = 0
        stats["rejected"] = 0

    def walk():
        position = 0
        while True:
            marker = searchable.find(BLOCK_MAGIC, position)
            if marker < 0:
                return
            position = marker + 1
            if marker < 8:
                continue
            offset = marker - 8
            if stats is not None:
                stats["candidates"] += 1

            if offset > len(source) - BLOCK_HEADER_SIZE:
                if stats is not None:
                    stats["rejected"] += 1
                continue

            crc_header = _read_u32(source, offset)
            crc_body = _read_u32(source, offset + 4)
            kind = _read_u32(source, offset + 0x10)
            payload_length = _read_u32(source, offset + 0x14)
            inflated_length = _read_u32(source, offset + 0x18)
            remaining = len(source) - (offset + BLOCK_HEADER_SIZE)
            if inflated_length == 0 or inflated_length > MAX_INFLATED:
                if stats is not None:
                    stats["rejected"] += 1
                continue
            if payload_length > remaining:
                if stats is not None:
                    stats["rejected"] += 1
                continue

            end = offset + BLOCK_HEADER_SIZE + payload_length
            header_ok = (zlib.crc32(source[offset + 4 : offset + 0x20]) & 0xFFFFFFFF) == crc_header
            body_ok = (zlib.crc32(source[offset + 8 : end]) & 0xFFFFFFFF) == crc_body
            if not header_ok or not body_ok or kind not in SUPPORTED_KINDS:
                if stats is not None:
                    stats["rejected"] += 1
                continue

            try:
                block_data = decode_lz4(source[offset + BLOCK_HEADER_SIZE : end], inflated_length)
            except ValueError:
                if stats is not None:
                    stats["rejected"] += 1
                continue

            if stats is not None:
                stats["verified"] += 1
            # A verified block owns its header and payload.  Skip magic bytes
            # embedded in that payload; rejected candidates keep the
            # one-byte advance above so a later valid block can still recover.
            position = end
            yield Block(offset=offset, kind=kind, data=block_data)

    return walk()
