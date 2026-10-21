"""Experimental Nostromo object records, bounded by their serialized lengths.

Types/offsets below were verified on Outlook 16.114. Unknown layouts and
ambiguous identities are deliberately left out instead of guessed.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
import hashlib
import re
import struct
import unicodedata


FRAME_PREFIX = b"\x01\x00\x00\x00\x00\x01"
ANCHOR = "IPM.Note\0".encode("utf-16le")
EMAIL = re.compile(r"[^<>\s@]+@[^<>\s@]+\.[^<>\s@]+\Z")
MESSAGE_ID = re.compile(r"<[^<>\s]+@[^<>\s]+>\Z")
PREFIXES = re.compile(r"^(?:(?:re|fw|fwd|aw|回复|答复|转发)\s*[:：]\s*)+", re.I)
EPOCH_TICKS = 621355968000000000
DOTNET_TICKS_PER_SECOND = 10_000_000
# .NET DateTime ticks cover 0001-01-01 through 9999-12-31.  The local
# metadata path converts Unix seconds to Asia/Singapore datetime values, so
# cap UTC at 9999-12-31 16:00:00 exclusive; +08:00 then stays in year 9999.
MIN_SUPPORTED_TICKS = 0
MAX_SUPPORTED_TICKS_EXCLUSIVE = (
    3155378976000000000 - 8 * 60 * 60 * DOTNET_TICKS_PER_SECOND
)
ENTRY_PREFIX = bytes.fromhex("0009002e00000000001d840311aa6611cd9bc800aa002fc45a")

# Verified type-202 inline HTML layout (Outlook 16.114).  These offsets are
# relative to the serialized frame, not to its payload after the 38-byte
# frame header.  Keep them explicit so a nearby descriptor cannot silently
# become a body association.
INLINE_BODY_CODEPAGE_OFFSET = 1658
INLINE_BODY_FULL_OFFSET = 1670
INLINE_BODY_FULL_DESCRIPTOR = 1674
INLINE_BODY_UNIQUE_OFFSET = 1678
INLINE_BODY_UNIQUE_DESCRIPTOR = 1682
INLINE_BODY_VARIANT = 1714
INLINE_BODY_EXTRA_POOL_OFFSET = 1730
INLINE_BODY_EXTRA_POOL_LENGTH = 1734
INLINE_BODY_CLASS_LENGTH = 1690
INLINE_BODY_CLASS_OFFSET = 1880
INLINE_BODY_PAYLOAD_BASE = 1898
INLINE_BODY_OWNER_OFFSET = 68
INLINE_BODY_CODEPAGE = 65001
INLINE_BODY_CLASS = "IPM.Note"
INLINE_BODY_STANDARD_VARIANT = 4
INLINE_BODY_EXTRA_VARIANT = 1
INLINE_BODY_EXTRA_OFFSET = 18
INLINE_BODY_EXTRA_LENGTH = 16


@dataclass(frozen=True)
class Frame:
    block: int
    offset: int
    kind: int
    key: int
    data: bytes


def iter_frames(block):
    data, position = block.data, 0
    while (start := data.find(FRAME_PREFIX, position)) >= 0:
        if start + 38 > len(data) or data[start + 10:start + 12] != b"\x05\0":
            position = start + 1
            continue
        size = struct.unpack_from("<I", data, start + 6)[0]
        if size < 24 or size != struct.unpack_from("<I", data, start + 14)[0]:
            position = start + 1
            continue
        end = start + 15 + size
        # A block-final serialized object omits the separator byte that is
        # present between objects. Never borrow bytes from another block.
        if end == len(data) + 1:
            end -= 1
        if end > len(data):
            position = start + 1
            continue
        kind = struct.unpack_from("<H", data, start + 20)[0]
        key = struct.unpack_from("<Q", data, start + 30)[0]
        # A structurally valid frame owns its complete payload, even when its
        # kind is unsupported or its key is unusable. Only malformed frame
        # candidates may be rescanned for a nested prefix.
        position = end
        if key and kind in (201, 202, 247, 362):
            yield Frame(block.offset, start, kind, key, data[start:end])


def entry_ids(data):
    """Exchange message EntryIds; attachment suffixes follow this 51-byte id."""
    found, position = set(), 0
    while (start := data.find(ENTRY_PREFIX, position)) >= 0:
        position = start + len(ENTRY_PREFIX)
        value = data[start:start + 51]
        if len(value) == 51 and value[-2:] == b"\0\0":
            found.add(value.hex())
    return found


def _typed_reference_values(frame, kind):
    """Return every non-zero u64 target following a typed-reference marker."""
    marker = struct.pack("<H", kind) + b"\0" * 8
    values = []
    position = 0
    while (start := frame.data.find(marker, position)) >= 0:
        position = start + len(marker)
        if position + 8 <= len(frame.data):
            value = struct.unpack_from("<Q", frame.data, position)[0]
            if value:
                values.append((start, value))
    return values


def typed_references(frame, kind):
    """Return non-zero u64 targets following every typed-reference marker."""
    return {value for _, value in _typed_reference_values(frame, kind)}


def parse_inline_body(frame, include_data=True):
    """Parse the verified fixed-layout inline HTML carried by a type-202 frame.

    The parser validates only structural evidence.  It deliberately leaves
    UTF-8 and HTML decoding to the caller that freezes the body.
    """
    if frame.kind != 202:
        return None
    data = frame.data

    owner_references = [
        (position, target)
        for position, target in _typed_reference_values(frame, 201)
        # The fixed header fields before the class are UTF-16/u32/u64 aligned.
        # An odd-position marker there can be a cross-field byte window (the
        # real 48513898 object has one at 1641); keep aligned markers strict.
        if not (position < INLINE_BODY_CLASS_OFFSET and position % 2)
    ]
    if (len(owner_references) != 1
            or owner_references[0][0] != INLINE_BODY_OWNER_OFFSET
            or owner_references[0][1] != frame.key):
        return None
    # A shared object graph may carry another message's EntryId even when its
    # typed owner marker happens to be singular.  Do not guess which body it
    # belongs to in that case.
    if len(entry_ids(data)) > 1:
        return None

    class_positions = []
    position = 0
    while (start := data.find(ANCHOR, position)) >= 0:
        class_positions.append(start)
        position = start + len(ANCHOR)
    if class_positions != [INLINE_BODY_CLASS_OFFSET]:
        return None

    fields_end = max(
        INLINE_BODY_CODEPAGE_OFFSET + 4,
        INLINE_BODY_FULL_DESCRIPTOR + 4,
        INLINE_BODY_UNIQUE_DESCRIPTOR + 4,
        INLINE_BODY_EXTRA_POOL_LENGTH + 4,
        INLINE_BODY_CLASS_LENGTH + 4,
        INLINE_BODY_CLASS_OFFSET + len(ANCHOR),
    )
    if len(data) < fields_end:
        return None

    codepage = struct.unpack_from("<I", data, INLINE_BODY_CODEPAGE_OFFSET)[0]
    class_length = struct.unpack_from("<I", data, INLINE_BODY_CLASS_LENGTH)[0]
    if codepage != INLINE_BODY_CODEPAGE or class_length != len(ANCHOR):
        return None
    variant = struct.unpack_from("<I", data, INLINE_BODY_VARIANT)[0]
    extra_offset = struct.unpack_from("<I", data, INLINE_BODY_EXTRA_POOL_OFFSET)[0]
    extra_length = struct.unpack_from("<I", data, INLINE_BODY_EXTRA_POOL_LENGTH)[0]
    if (variant in (INLINE_BODY_STANDARD_VARIANT, INLINE_BODY_EXTRA_VARIANT)
            and extra_offset == 0 and extra_length == 0):
        payload_base = INLINE_BODY_PAYLOAD_BASE
    elif (variant, extra_offset, extra_length) == (
            INLINE_BODY_EXTRA_VARIANT, INLINE_BODY_EXTRA_OFFSET, INLINE_BODY_EXTRA_LENGTH):
        extra_start = INLINE_BODY_CLASS_OFFSET + extra_offset
        payload_base = extra_start + extra_length
        if len(data) < payload_base:
            return None
    else:
        return None

    full_offset = struct.unpack_from("<I", data, INLINE_BODY_FULL_OFFSET)[0]
    full_descriptor = struct.unpack_from("<I", data, INLINE_BODY_FULL_DESCRIPTOR)[0]
    if not full_descriptor & 0x80000000:
        return None
    full_length = full_descriptor & 0x7FFFFFFF
    if not full_length:
        return None
    start = payload_base + full_offset
    end = start + full_length
    if start < payload_base or end > len(data) or end <= start:
        return None
    body = data[start:end]
    if not body or not any(body):
        return None

    unique_offset = struct.unpack_from("<I", data, INLINE_BODY_UNIQUE_OFFSET)[0]
    unique_descriptor = struct.unpack_from("<I", data, INLINE_BODY_UNIQUE_DESCRIPTOR)[0]
    evidence = {
            "block": frame.block,
            "frame_offset": frame.offset,
            "frame_bytes": len(data),
            "frame_sha256": hashlib.sha256(data).hexdigest(),
            "owner_kind": 201,
            "owner_offset": owner_references[0][0],
            "owner_key": frame.key,
            "class": INLINE_BODY_CLASS,
            "class_offset": INLINE_BODY_CLASS_OFFSET,
            "class_length": class_length,
            "codepage": codepage,
            "codepage_offset": INLINE_BODY_CODEPAGE_OFFSET,
            "payload_base": payload_base,
            "extra_pool_field": {
                "variant": variant,
                "offset": extra_offset,
                "length": extra_length,
                "offset_field": INLINE_BODY_EXTRA_POOL_OFFSET,
                "length_field": INLINE_BODY_EXTRA_POOL_LENGTH,
            },
            "full_offset": full_offset,
            "full_descriptor": full_descriptor,
            "full_length": full_length,
            "full_data_offset": start,
            # Retain the alternate descriptor as evidence only.  It is never
            # used as a fallback for the full representation above.
            "unique_offset": unique_offset,
            "unique_descriptor": unique_descriptor,
    }
    # Metadata-first callers need the verified layout and source locator, but
    # must not copy every inline body into temporary files during indexing.
    return {**({"data": body} if include_data else {}), "evidence": evidence}


def clean_text(value):
    return all(c in "\t\r\n\ufeff\u200b\u200c\u200d" or
               unicodedata.category(c)[0] != "C" for c in value)


def read_wide(data, start):
    end = start
    while end + 1 < len(data) and data[end:end + 2] != b"\0\0":
        end += 2
    if end + 1 >= len(data):
        raise ValueError("unterminated UTF-16 field")
    value = data[start:end].decode("utf-16le", errors="strict")
    if not clean_text(value):
        raise ValueError("non-text UTF-16 field")
    return value, end + 2


def preceding_fields(data, position, limit=24):
    fields = []
    for _ in range(limit):
        if position < 2 or data[position - 2:position] != b"\0\0":
            break
        end, start = position - 2, position - 4
        while start >= 0 and data[start:start + 2] != b"\0\0":
            start -= 2
        start += 2
        try:
            value = data[start:end].decode("utf-16le", errors="strict")
        except UnicodeError:
            break
        if not clean_text(value):
            break
        if value:
            fields.insert(0, value)
        position = start
    return fields


def normalize_subject(value):
    return PREFIXES.sub("", value.strip())


def participant_pair_preview(after, subject):
    """Return the preview from the verified participant-pair after layout."""
    prefix = 0
    while prefix < len(after) and MESSAGE_ID.fullmatch(after[prefix]):
        prefix += 1
    if prefix not in (1, 2) or len(after) - prefix != 5:
        return None
    name, email, preview, display_subject, normalized_subject = after[prefix:]
    if not EMAIL.fullmatch(email):
        return None
    if (EMAIL.fullmatch(name) or MESSAGE_ID.fullmatch(name)
            or EMAIL.fullmatch(preview) or MESSAGE_ID.fullmatch(preview)):
        return None
    if not subject:
        return None
    subject_key = normalize_subject(subject)
    if (normalize_subject(display_subject) != subject_key
            or normalize_subject(normalized_subject) != subject_key):
        return None
    return preview


def record_times(data, key):
    times = set()
    for position in range(len(data) - 7):
        value = struct.unpack_from("<Q", data, position)[0]
        if _supported_ticks(value):
            times.add((value - EPOCH_TICKS) / DOTNET_TICKS_PER_SECOND)
    marker = struct.pack("<H", 271) + b"\0" * 8 + struct.pack("<Q", key)
    explicit, position = set(), 0
    while (start := data.find(marker, position)) >= 0:
        position = start + len(marker)
        if start >= 10:
            value = struct.unpack_from("<Q", data, start - 10)[0]
            if _supported_ticks(value):
                explicit.add((value - EPOCH_TICKS) / DOTNET_TICKS_PER_SECOND)
    if len(explicit) == 1:
        return next(iter(explicit)), max(times) if times else None, "header_field_271"
    return None, max(times) if times else None, "unresolved"


def _supported_ticks(value):
    """Return whether ``value`` is a trusted, millisecond-aligned DateTime tick."""
    return (MIN_SUPPORTED_TICKS <= value < MAX_SUPPORTED_TICKS_EXCLUSIVE
            and value % 10_000 == 0)


def parse_header(frame):
    if frame.kind != 201:
        return None
    data = frame.data
    entries = entry_ids(data)
    if len(entries) > 1:
        # Shared object graphs can contain another mail's identity and fields.
        # Skip until their property ownership has been independently decoded.
        return None
    anchor = data.find(ANCHOR)
    if anchor < 0 or data.find(ANCHOR, anchor + len(ANCHOR)) >= 0:
        return None
    before = preceding_fields(data, anchor)
    after, position = [], anchor + len(ANCHOR)
    for _ in range(16):
        try:
            value, position = read_wide(data, position)
        except (ValueError, UnicodeError):
            break
        if value:
            after.append(value)
    sender_indices = [i for i, value in enumerate(before) if EMAIL.fullmatch(value)]
    sender_index = sender_indices[0] if len(sender_indices) == 1 else None
    sender = before[sender_index] if sender_index is not None else None
    sender_name = None
    if sender_index is not None and sender_index + 1 < len(before):
        candidate = before[sender_index + 1]
        if not EMAIL.fullmatch(candidate):
            sender_name = candidate
    ids = [value for value in after if MESSAGE_ID.fullmatch(value)]
    if not ids:
        return None
    # Before the anchor, repeated display names are not a typed subject pair.
    # Use that layout only when its sender-address slot is unambiguous.
    strings = (before if len(sender_indices) == 1 else []) + after
    candidates = [s for s in strings if 2 <= len(s) <= 1024
                  and not EMAIL.fullmatch(s) and not s.startswith(("<", "~/", "IPM."))
                  and s != sender_name]
    repeated = [s for s, count in Counter(map(normalize_subject, candidates)).items()
                if s and count > 1]
    subject = None
    if len(repeated) == 1:
        subject = max((s for s in candidates if normalize_subject(s) == repeated[0]), key=len)
    preview = participant_pair_preview(after, subject)
    preview_method = "participant_pair_layout" if preview is not None else "after_class_candidate"
    if preview is None:
        preview = next((s for s in after if not MESSAGE_ID.fullmatch(s)
                        and (not subject or normalize_subject(s) != normalize_subject(subject))), None)
    sent, modified, time_method = record_times(data, frame.key)
    # Osa response/header cross-checks locate InternetMessageId last in the
    # string heap. Earlier ids are related-message references, not this mail.
    message_id = ids[-1]
    identity = hashlib.sha256(message_id.encode()).hexdigest()[:12]
    return {
        "id": f"hx_{frame.key:x}_{identity}", "key": frame.key,
        "message_id": message_id, "subject": subject,
        "related_message_ids": ids[:-1],
        "subject_status": "paired" if subject else "unresolved",
        "sender": sender, "sender_name": sender_name, "preview": preview,
        "preview_method": preview_method,
        "sent_unix": sent, "latest_tick_unix": modified,
        "time_method": time_method,
        "sender_status": "single_header_address" if sender else "unresolved",
        "entry_ids": sorted(entries),
        "evidence": {"block": frame.block, "frame_offset": frame.offset,
                     "frame_bytes": len(data), "frame_sha256": hashlib.sha256(data).hexdigest()},
    }


def file_references(frame):
    needle, position, found = "~/Files/".encode("utf-16le"), 0, []
    while (start := frame.data.find(needle, position)) >= 0:
        position = start + len(needle)
        try:
            value, _ = read_wide(frame.data, start)
        except (ValueError, UnicodeError):
            continue
        previous = preceding_fields(frame.data, start, 2)
        found.append({"path": value, "name": previous[-1] if previous else None})
    return found


def parse_attachment(frame):
    if frame.kind != 362:
        return None
    owners = {}
    for kind in (201, 202):
        marker = struct.pack("<H", kind) + b"\0" * 8
        values, position = set(), 0
        while (start := frame.data.find(marker, position)) >= 0:
            position = start + len(marker)
            if position + 8 <= len(frame.data):
                values.add(struct.unpack_from("<Q", frame.data, position)[0])
        owners[kind] = values - {0}
    typed_owners = owners[201] | owners[202]
    references = [ref for ref in file_references(frame) if "/Attachments/" in ref["path"]]
    entries = entry_ids(frame.data)
    if len(typed_owners) > 1 or len(references) != 1:
        return None
    if not typed_owners and len(entries) != 1:
        return None
    return {"key": frame.key, "owner": next(iter(typed_owners), None),
            "entry_ids": sorted(entries), **references[0],
            "evidence": {"block": frame.block, "frame_offset": frame.offset,
                         "owner_reference_types": [k for k in owners if owners[k]],
                         "owner_references": sorted(typed_owners)}}
