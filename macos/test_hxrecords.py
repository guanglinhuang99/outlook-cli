import struct
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace


sys.path.insert(0, str(Path(__file__).parent))
from hxrecords import (  # noqa: E402
    ANCHOR,
    DOTNET_TICKS_PER_SECOND,
    EPOCH_TICKS,
    ENTRY_PREFIX,
    FRAME_PREFIX,
    Frame,
    INLINE_BODY_CLASS_OFFSET,
    INLINE_BODY_EXTRA_VARIANT,
    INLINE_BODY_EXTRA_POOL_LENGTH,
    INLINE_BODY_EXTRA_POOL_OFFSET,
    INLINE_BODY_PAYLOAD_BASE,
    INLINE_BODY_STANDARD_VARIANT,
    INLINE_BODY_VARIANT_5,
    INLINE_BODY_VARIANT,
    MAX_SUPPORTED_TICKS_EXCLUSIVE,
    MIN_SUPPORTED_TICKS,
    entry_ids,
    iter_frames,
    normalize_subject,
    parse_attachment,
    parse_header,
    parse_inline_body,
    parse_person,
    person_parent,
    typed_references,
)


def wide(value: str) -> bytes:
    return value.encode("utf-16le") + b"\0\0"


def serialized_frame(
    kind: int,
    key: int,
    payload: bytes = b"x",
    *,
    declared_extra: int = 0,
) -> bytes:
    """Build the length-bounded frame shape consumed by ``iter_frames``."""

    size = 23 + len(payload) + declared_extra
    result = bytearray(38 + len(payload))
    result[:6] = FRAME_PREFIX
    struct.pack_into("<I", result, 6, size)
    result[10:12] = b"\x05\0"
    struct.pack_into("<I", result, 14, size)
    struct.pack_into("<H", result, 20, kind)
    struct.pack_into("<Q", result, 30, key)
    result[38:] = payload
    return bytes(result)


def exchange_entry(tag: int) -> bytes:
    """Build one synthetic 51-byte Exchange message EntryId."""

    return ENTRY_PREFIX + tag.to_bytes(24, "little") + b"\0\0"


def registry_frame(frame_key: int, entry: bytes, references=()) -> bytes:
    payload = bytearray(entry)
    for reference in references:
        path = str(reference)
        payload += b"\0\0" + wide(path.rsplit("/", 1)[-1]) + wide(path)
    return serialized_frame(247, frame_key, bytes(payload))


def mail_payload(before=(), after=(), tail=b"") -> bytes:
    return b"".join(wide(value) for value in before) + ANCHOR + b"".join(
        wide(value) for value in after
    ) + tail


def mail_frame(
    key: int,
    *,
    before=(),
    after=(),
    tail=b"",
    kind: int = 201,
) -> bytes:
    return serialized_frame(kind, key, mail_payload(before, after, tail))


def as_frame(raw: bytes, *, kind: int = 201, key: int = 1) -> Frame:
    return Frame(block=7, offset=0, kind=kind, key=key, data=raw)


def attachment_frame(
    frame_key: int,
    owner: int | None,
    path: str | None,
    *,
    entry: bytes | None = None,
    owner_202: int | None = None,
    extra_owners_201=(),
    extra_paths=(),
    name: str = "同名附件.txt",
) -> bytes:
    payload = bytearray()
    if entry is not None:
        payload += entry
    if owner is not None:
        payload += struct.pack("<H", 201) + b"\0" * 8 + struct.pack("<Q", owner)
    for extra in extra_owners_201:
        payload += struct.pack("<H", 201) + b"\0" * 8 + struct.pack("<Q", extra)
    if owner_202 is not None:
        payload += struct.pack("<H", 202) + b"\0" * 8 + struct.pack("<Q", owner_202)
    if path is not None:
        # Keep a UTF-16 field boundary after the binary owner references so
        # preceding_fields can stop before the filename.
        payload += b"\0\0" + wide(name) + wide(path)
    for extra_path in extra_paths:
        payload += b"\0\0" + wide("另一个附件.txt") + wide(extra_path)
    return serialized_frame(362, frame_key, bytes(payload))


def ticks(seconds: int | float) -> bytes:
    tick_value = (EPOCH_TICKS + seconds * DOTNET_TICKS_PER_SECOND
                  if isinstance(seconds, int)
                  else EPOCH_TICKS + round(seconds * DOTNET_TICKS_PER_SECOND))
    return struct.pack("<Q", tick_value)


def verified_ticks(key: int, seconds: int | float) -> bytes:
    return ticks(seconds) + b"\0\0" + struct.pack("<H", 271) + b"\0" * 8 + struct.pack("<Q", key)


def inline_body_frame(key, full_html, unique_html=None, owner=None, **overrides):
    """Build a synthetic type-202 inline-body frame at the verified offsets."""

    def body_bytes(value):
        return value.encode("utf-8") if isinstance(value, str) else bytes(value)

    full = body_bytes(full_html)
    unique = full if unique_html is None else body_bytes(unique_html)
    codepage = overrides.pop("codepage", 65001)
    class_name = overrides.pop("class_name", "IPM.Note")
    class_bytes = (class_name + "\0").encode("utf-16le")
    class_offset = overrides.pop("class_offset", 1880)
    class_length = overrides.pop("class_length", len(class_bytes))
    owner_offset = overrides.pop("owner_offset", 68)
    full_offset = overrides.pop("full_offset", 0)
    unique_offset = overrides.pop(
        "unique_offset", full_offset + len(full) if unique_html is not None else 0
    )
    full_length = overrides.pop("full_length", len(full))
    unique_length = overrides.pop("unique_length", len(unique))
    full_descriptor = overrides.pop("full_descriptor", 0x80000000 | full_length)
    unique_descriptor = overrides.pop("unique_descriptor", 0x80000000 | unique_length)
    owner_refs = overrides.pop("owner_refs", None)
    if owner_refs is None:
        owner_refs = (key if owner is None else owner,)
    extra_classes = overrides.pop("extra_classes", ())
    extra_pool = overrides.pop("extra_pool", None)
    if extra_pool is not None:
        extra_pool = bytes(extra_pool)
        extra_variant = overrides.pop("extra_pool_variant", 1)
        extra_offset = overrides.pop("extra_pool_offset", 18)
        extra_length = overrides.pop("extra_pool_length", 16)
    else:
        extra_variant = overrides.pop("extra_pool_variant", 4)
        extra_offset = overrides.pop("extra_pool_offset", 0)
        extra_length = overrides.pop("extra_pool_length", 0)
    if overrides:
        raise TypeError(f"unknown inline body override(s): {sorted(overrides)}")

    # The offsets are absolute within the serialized frame.  The helper keeps
    # the frame compact while allowing malformed descriptors to be tested by
    # pointing beyond the bytes that were actually written.
    payload = bytearray(max(0, INLINE_BODY_PAYLOAD_BASE - 38))

    def put(offset, value):
        if offset < 38:
            raise ValueError("inline body offsets are serialized-frame offsets")
        index = offset - 38
        end = index + len(value)
        if end > len(payload):
            payload.extend(b"\0" * (end - len(payload)))
        payload[index:end] = value

    def put_u32(offset, value):
        put(offset, struct.pack("<I", value))

    put_u32(1658, codepage)
    put_u32(1670, full_offset)
    put_u32(1674, full_descriptor)
    put_u32(1678, unique_offset)
    put_u32(1682, unique_descriptor)
    put_u32(INLINE_BODY_VARIANT, extra_variant)
    put_u32(INLINE_BODY_EXTRA_POOL_OFFSET, extra_offset)
    put_u32(INLINE_BODY_EXTRA_POOL_LENGTH, extra_length)
    put_u32(1690, class_length)
    put(class_offset, class_bytes)

    if extra_pool is not None and 0 <= extra_offset <= 1_048_576:
        put(INLINE_BODY_CLASS_OFFSET + extra_offset, extra_pool)

    owner_cursor = owner_offset
    marker_prefix = struct.pack("<H", 201) + b"\0" * 8
    for owner_ref in owner_refs:
        put(owner_cursor, marker_prefix + struct.pack("<Q", owner_ref))
        owner_cursor += len(marker_prefix) + 8

    payload_base = INLINE_BODY_PAYLOAD_BASE
    if (extra_variant in (INLINE_BODY_EXTRA_VARIANT,
                          INLINE_BODY_STANDARD_VARIANT,
                          INLINE_BODY_VARIANT_5)
            and extra_offset == 18 and extra_length == 16):
        payload_base = INLINE_BODY_CLASS_OFFSET + extra_offset + extra_length

    # Avoid allocating unbounded memory for a deliberately invalid offset.
    if 0 <= full_offset <= 1_048_576:
        put(payload_base + full_offset, full)
    if 0 <= unique_offset <= 1_048_576:
        put(payload_base + unique_offset, unique)
    extra_class_cursor = 38 + len(payload) + 16
    for extra_class in extra_classes:
        value = (extra_class + "\0").encode("utf-16le")
        put(extra_class_cursor, value)
        extra_class_cursor += len(value)
    return serialized_frame(202, key, bytes(payload))


def person_frame(key, owner, role="to", name="张伟", address="zhang@example.com",
                 *, variant=1):
    """Build a verified HxPerson frame at the researched fixed offsets."""
    parent_type = {"to": 204, "cc": 205}.get(role, role)
    email = wide(address)
    display_name = wide(name)
    email_offset = 53
    name_offset = email_offset + len(email)
    frame_length = 360 + name_offset + len(display_name) + 5
    payload = bytearray(frame_length - 38)

    def put(offset, value):
        index = offset - 38
        payload[index:index + len(value)] = value

    def put_u16(offset, value):
        put(offset, struct.pack("<H", value))

    def put_u32(offset, value):
        put(offset, struct.pack("<I", value))

    def put_u64(offset, value):
        put(offset, struct.pack("<Q", value))

    email_length, name_length = len(email), len(display_name)
    put_u32(38, parent_type)
    put_u64(42, owner)
    put_u64(50, key)
    put_u64(58, owner)
    put_u16(68, parent_type)
    put_u64(78, owner)
    put_u32(302, name_offset)
    put_u32(306, 0x80000000 | name_length)
    put_u32(310, email_offset)
    put_u32(314, 0x80000000 | email_length)
    put_u32(322, 0x80000000 | email_offset)
    put_u32(326, variant)
    put(360 + email_offset - 5, b"\x01" * 4)
    put(360 + email_offset, email)
    put(360 + name_offset, display_name)
    return serialized_frame(85, key, bytes(payload))


class HxRecordTests(unittest.TestCase):
    def test_entry_ids_are_exact_51_byte_values(self):
        first, second = exchange_entry(1), exchange_entry(2)

        self.assertEqual(len(first), 51)
        self.assertEqual(entry_ids(b"prefix" + first + b"gap" + second), {first.hex(), second.hex()})
        header = parse_header(
            as_frame(
                mail_frame(
                    9,
                    before=("sender@example.com", "张伟"),
                    after=("主题", "<entry@example.test>", "主题"),
                    tail=first,
                )
            )
        )
        self.assertEqual(header["entry_ids"], [first.hex()])

    def test_last_message_id_is_actual_and_earlier_ids_are_related(self):
        first = mail_frame(
            0x901,
            before=("first@example.com", "甲"),
            after=(
                "相同主题",
                "<parent@example.test>",
                "<actual-one@example.test>",
                "相同主题",
            ),
        )
        second = mail_frame(
            0x902,
            before=("second@example.com", "乙"),
            after=(
                "相同主题",
                "<other-parent@example.test>",
                "<actual-two@example.test>",
                "相同主题",
            ),
        )
        frames = list(iter_frames(SimpleNamespace(offset=12, data=first + second)))
        first_header, second_header = (parse_header(frame) for frame in frames)

        self.assertEqual(first_header["message_id"], "<actual-one@example.test>")
        self.assertEqual(second_header["message_id"], "<actual-two@example.test>")
        self.assertEqual(first_header["related_message_ids"], ["<parent@example.test>"])
        self.assertEqual(second_header["related_message_ids"], ["<other-parent@example.test>"])
        self.assertNotIn("<parent@example.test>", {first_header["message_id"], second_header["message_id"]})
        self.assertNotEqual(first_header["id"], second_header["id"])

    def test_header_with_two_exchange_entry_ids_is_rejected(self):
        first, second = exchange_entry(31), exchange_entry(32)
        frame = as_frame(
            mail_frame(
                0x931,
                before=("sender@example.com", "张伟"),
                after=("主题", "<shared-object@example.test>", "主题"),
                tail=first + second,
            )
        )

        self.assertIsNone(parse_header(frame))

    def test_entry_only_attachment_is_a_candidate_without_typed_owner(self):
        entry = exchange_entry(33)
        frame = next(
            iter_frames(
                SimpleNamespace(
                    offset=13,
                    data=attachment_frame(
                        0x933,
                        None,
                        "~/Files/Attachments/entry-only.bin",
                        entry=entry,
                    ),
                )
            )
        )

        parsed = parse_attachment(frame)
        self.assertIsNone(parsed["owner"])
        self.assertEqual(parsed["entry_ids"], [entry.hex()])
        self.assertEqual(parsed["path"], "~/Files/Attachments/entry-only.bin")

    def test_adjacent_chinese_mail_frames_keep_their_subjects_separate(self):
        first = mail_frame(
            0x101,
            before=("sender1@example.com", "李小龙🧪"),
            after=("中文主题甲", "<first@example.test>", "中文主题甲"),
        )
        second = mail_frame(
            0x202,
            before=("sender2@example.com", "王小明😀"),
            after=("中文主题乙", "<second@example.test>", "中文主题乙"),
        )
        frames = list(iter_frames(SimpleNamespace(offset=91, data=first + second)))

        self.assertEqual([frame.key for frame in frames], [0x101, 0x202])
        self.assertEqual([frame.offset for frame in frames], [0, len(first)])
        self.assertEqual(parse_header(frames[0])["subject"], "中文主题甲")
        self.assertEqual(parse_header(frames[1])["subject"], "中文主题乙")
        self.assertEqual(parse_header(frames[0])["sender_name"], "李小龙🧪")
        self.assertEqual(parse_header(frames[1])["sender_name"], "王小明😀")

    def test_subject_can_be_after_or_before_anchor(self):
        after_layout = as_frame(
            mail_frame(
                1,
                before=("sender@example.com", "张伟🧪"),
                after=("回复：中文主题甲", "<after@example.test>", "Re:中文主题甲"),
            )
        )
        before_layout = as_frame(
            mail_frame(
                2,
                before=("中文主题乙", "中文主题乙", "sender@example.com", "张伟🧪"),
                after=("<before@example.test>",),
            )
        )

        after_header = parse_header(after_layout)
        before_header = parse_header(before_layout)
        self.assertEqual(after_header["subject"], "回复：中文主题甲")
        self.assertEqual(before_header["subject"], "中文主题乙")
        self.assertEqual(normalize_subject(after_header["subject"]), "中文主题甲")

    def test_participant_pair_layout_uses_true_preview_after_two_message_ids(self):
        header = parse_header(
            as_frame(
                mail_frame(
                    0x303,
                    before=("sender@example.com", "张伟🧪"),
                    after=(
                        "<parent@example.test>",
                        "<actual@example.test>",
                        "李小龙😀",
                        "sender@example.com",
                        "真正的中文预览内容🧪",
                        "回复：严格布局主题",
                        "严格布局主题",
                    ),
                ),
                key=0x303,
            )
        )

        self.assertEqual(header["preview"], "真正的中文预览内容🧪")
        self.assertEqual(header["preview_method"], "participant_pair_layout")
        self.assertEqual(header["subject"], "回复：严格布局主题")

    def test_participant_pair_layout_uses_true_preview_after_one_message_id(self):
        header = parse_header(
            as_frame(
                mail_frame(
                    0x304,
                    before=("sender@example.com", "王小明😀"),
                    after=(
                        "<actual-one@example.test>",
                        "赵六🧪",
                        "sender@example.com",
                        "另一段中文正文预览😀",
                        "单一严格主题",
                        "单一严格主题",
                    ),
                ),
                key=0x304,
            )
        )

        self.assertEqual(header["preview"], "另一段中文正文预览😀")
        self.assertEqual(header["preview_method"], "participant_pair_layout")
        self.assertEqual(header["subject"], "单一严格主题")

    def test_participant_pair_layout_rejects_missing_extra_or_unpaired_tail(self):
        cases = (
            ("missing", ("<missing@example.test>", "姓名", "sender@example.com", "预览", "主题")),
            ("extra", ("<extra@example.test>", "姓名", "sender@example.com", "预览", "主题", "主题", "多余")),
            ("unpaired", ("<unpaired@example.test>", "姓名", "sender@example.com", "预览", "主题甲", "主题乙")),
        )
        for _, after in cases:
            header = parse_header(
                as_frame(
                    mail_frame(
                        0x305,
                        before=("sender@example.com", "测试发送人"),
                        after=after,
                    ),
                    key=0x305,
                )
            )
            self.assertEqual(header["preview_method"], "after_class_candidate")

    def test_reply_prefixes_compare_normalized_but_keep_display_subject(self):
        frame = as_frame(
            mail_frame(
                3,
                before=("回复：原始主题", "sender@example.com", "张伟"),
                after=("<prefix@example.test>", "Re:原始主题"),
            )
        )
        header = parse_header(frame)

        self.assertEqual(header["subject"], "回复：原始主题")
        self.assertEqual(header["subject_status"], "paired")
        self.assertEqual(normalize_subject(header["subject"]), "原始主题")

    def test_201_and_202_share_key_and_attachment_uses_unique_internal_owner(self):
        owner = 0x1122334455667788
        mail_201 = mail_frame(0xABC, before=("sender@example.com", "张伟"), after=("<m@example.test>",))
        mail_202 = serialized_frame(202, 0xABC, b"state")
        attachment = attachment_frame(
            0xDEF,
            owner,
            "~/Files/Attachments/owner-a/合同.pdf",
            owner_202=owner,
            name="合同.pdf",
        )
        frames = list(iter_frames(SimpleNamespace(offset=3, data=mail_201 + mail_202 + attachment)))

        self.assertEqual([(frame.kind, frame.key) for frame in frames], [(201, 0xABC), (202, 0xABC), (362, 0xDEF)])
        parsed = parse_attachment(frames[2])
        self.assertEqual(parsed["key"], 0xDEF)
        self.assertEqual(parsed["owner"], owner)
        self.assertEqual(parsed["name"], "合同.pdf")
        self.assertEqual(parsed["path"], "~/Files/Attachments/owner-a/合同.pdf")
        self.assertEqual(parsed["evidence"]["owner_reference_types"], [201, 202])
        self.assertEqual(parsed["evidence"]["owner_references"], [owner])

    def test_type_202_only_owner_is_preserved(self):
        owner = 0x2233445566778899
        frame = next(
            iter_frames(
                SimpleNamespace(
                    offset=3,
                    data=attachment_frame(
                        0xD01,
                        None,
                        "~/Files/Attachments/type-202-only/合同.pdf",
                        owner_202=owner,
                        name="合同.pdf",
                    ),
                )
            )
        )

        parsed = parse_attachment(frame)
        self.assertEqual(parsed["owner"], owner)
        self.assertEqual(parsed["entry_ids"], [])
        self.assertEqual(parsed["evidence"]["owner_reference_types"], [202])
        self.assertEqual(parsed["evidence"]["owner_references"], [owner])

    def test_same_filename_with_different_owners_stays_frame_local(self):
        first = attachment_frame(
            0xF1,
            0x1111111111111111,
            "~/Files/Attachments/owner-a/同名附件.txt",
        )
        second = attachment_frame(
            0xF2,
            0x2222222222222222,
            "~/Files/Attachments/owner-b/同名附件.txt",
        )
        frames = list(iter_frames(SimpleNamespace(offset=4, data=first + second)))
        first_ref, second_ref = parse_attachment(frames[0]), parse_attachment(frames[1])

        self.assertEqual((first_ref["owner"], first_ref["path"]), (0x1111111111111111, "~/Files/Attachments/owner-a/同名附件.txt"))
        self.assertEqual((second_ref["owner"], second_ref["path"]), (0x2222222222222222, "~/Files/Attachments/owner-b/同名附件.txt"))
        self.assertEqual(first_ref["name"], second_ref["name"])

    def test_owner_conflicts_and_multiple_paths_are_rejected(self):
        conflicts = (
            attachment_frame(
                1,
                101,
                "~/Files/Attachments/owner-101/a.txt",
                extra_owners_201=(202,),
            ),
            attachment_frame(
                2,
                101,
                "~/Files/Attachments/owner-101/a.txt",
                owner_202=202,
            ),
            attachment_frame(
                3,
                101,
                "~/Files/Attachments/owner-101/a.txt",
                extra_paths=("~/Files/Attachments/owner-101/b.txt",),
            ),
        )

        for raw in conflicts:
            frame = next(iter_frames(SimpleNamespace(offset=5, data=raw)))
            self.assertIsNone(parse_attachment(frame))

    def test_duplicate_size_and_out_of_bounds_frames_are_skipped(self):
        duplicate = bytearray(serialized_frame(201, 1, b"duplicate"))
        size = struct.unpack_from("<I", duplicate, 6)[0]
        struct.pack_into("<I", duplicate, 14, size + 1)
        out_of_bounds = bytearray(serialized_frame(201, 2, b"overflow"))
        size = struct.unpack_from("<I", out_of_bounds, 6)[0]
        struct.pack_into("<I", out_of_bounds, 6, size + 100)
        struct.pack_into("<I", out_of_bounds, 14, size + 100)
        good = serialized_frame(202, 3, b"valid")

        frames = list(
            iter_frames(
                SimpleNamespace(offset=6, data=bytes(duplicate) + bytes(out_of_bounds) + good)
            )
        )
        self.assertEqual([(frame.kind, frame.key) for frame in frames], [(202, 3)])

    def test_structural_outer_frames_hide_nested_prefixes_but_invalid_candidates_recover(self):
        nested = serialized_frame(202, 0xB02, b"nested")
        supported_outer = serialized_frame(201, 0xB01, b"prefix" + nested + b"suffix")
        unsupported_outer = serialized_frame(999, 0xB03, b"prefix" + nested + b"suffix")
        trailing = serialized_frame(247, 0xB04, b"trailing")
        frames = list(
            iter_frames(SimpleNamespace(
                offset=6, data=supported_outer + unsupported_outer + trailing
            ))
        )
        self.assertEqual(
            [(frame.kind, frame.key) for frame in frames],
            [(201, 0xB01), (247, 0xB04)],
        )

        invalid_outer = bytearray(supported_outer)
        size = struct.unpack_from("<I", invalid_outer, 6)[0]
        struct.pack_into("<I", invalid_outer, 14, size + 1)
        recovered = list(iter_frames(SimpleNamespace(offset=6, data=bytes(invalid_outer))))
        self.assertEqual([(frame.kind, frame.key) for frame in recovered], [(202, 0xB02)])

    def test_block_final_missing_one_separator_is_legal_but_two_is_rejected(self):
        one_missing = serialized_frame(247, 1, b"registry", declared_extra=1)
        two_missing = serialized_frame(247, 2, b"registry", declared_extra=2)

        accepted = list(iter_frames(SimpleNamespace(offset=6, data=one_missing)))
        rejected = list(iter_frames(SimpleNamespace(offset=6, data=two_missing)))
        self.assertEqual(len(accepted), 1)
        self.assertEqual((accepted[0].kind, accepted[0].key), (247, 1))
        self.assertEqual(accepted[0].data, one_missing)
        self.assertEqual(rejected, [])

    def test_unpaired_subject_is_unresolved_and_missing_message_id_is_not_a_header(self):
        unresolved = as_frame(
            mail_frame(
                11,
                before=("sender@example.com", "张伟"),
                after=("孤立主题", "<unresolved@example.test>"),
            )
        )
        no_message_id = as_frame(
            mail_frame(
                12,
                before=("成对主题", "成对主题", "sender@example.com", "张伟"),
                after=("预览文本",),
            )
        )

        unresolved_header = parse_header(unresolved)
        self.assertIsNone(unresolved_header["subject"])
        self.assertEqual(unresolved_header["subject_status"], "unresolved")
        self.assertIsNone(parse_header(no_message_id))

    def test_unknown_ticks_are_not_send_time_and_stay_frame_local(self):
        first = as_frame(
            mail_frame(
                21,
                before=("sender@example.com", "张伟"),
                after=("主题甲", "<first@example.test>"),
                tail=ticks(946_684_810),
            )
        )
        second = as_frame(
            mail_frame(
                22,
                before=("sender@example.com", "张伟"),
                after=("主题乙", "<second@example.test>"),
                tail=ticks(946_684_820),
            )
        )

        first_header, second_header = parse_header(first), parse_header(second)
        self.assertIsNone(first_header["sent_unix"])
        self.assertEqual(first_header["latest_tick_unix"], 946_684_810)
        self.assertEqual(first_header["time_method"], "unresolved")
        self.assertIsNone(second_header["sent_unix"])
        self.assertEqual(second_header["latest_tick_unix"], 946_684_820)
        self.assertEqual(second_header["time_method"], "unresolved")

    def test_verified_271_marker_is_the_only_send_time_source(self):
        key = 23
        header = parse_header(
            as_frame(
                mail_frame(
                    key,
                    before=("sender@example.com", "张伟"),
                    after=("主题", "<verified@example.test>", "主题"),
                    tail=verified_ticks(key, 946_684_830),
                ),
                key=key,
            )
        )

        self.assertEqual(header["sent_unix"], 946_684_830)
        self.assertEqual(header["latest_tick_unix"], 946_684_830)
        self.assertEqual(header["time_method"], "header_field_271")

    def test_verified_271_preserves_millisecond_send_time(self):
        key = 36
        sent = 946_684_800.5
        header = parse_header(
            as_frame(
                mail_frame(
                    key,
                    before=("sender@example.com", "张伟"),
                    after=("主题", "<millisecond@example.test>"),
                    tail=verified_ticks(key, sent),
                ),
                key=key,
            )
        )

        self.assertEqual(header["sent_unix"], sent)
        self.assertEqual(header["latest_tick_unix"], sent)
        self.assertEqual(header["time_method"], "header_field_271")

    def test_verified_ticks_accept_2030_and_safe_datetime_boundaries(self):
        lower_seconds = (MIN_SUPPORTED_TICKS - EPOCH_TICKS) // DOTNET_TICKS_PER_SECOND
        upper_seconds = (
            MAX_SUPPORTED_TICKS_EXCLUSIVE - 10_000 - EPOCH_TICKS
        ) // DOTNET_TICKS_PER_SECOND
        cases = (
            (31, 1_893_456_000),  # 2030-01-01T00:00:00Z
            (32, lower_seconds),
            (33, upper_seconds),
        )
        for key, seconds in cases:
            with self.subTest(key=key, seconds=seconds):
                header = parse_header(
                    as_frame(
                        mail_frame(
                            key,
                            before=("sender@example.com", "张伟"),
                            after=("主题", f"<{key}@example.test>", "主题"),
                            tail=verified_ticks(key, seconds),
                        ),
                        key=key,
                    )
                )
                self.assertEqual(header["sent_unix"], seconds)
                self.assertEqual(header["time_method"], "header_field_271")

    def test_verified_tick_at_local_year_10000_boundary_is_rejected(self):
        key = 34
        upper_seconds = (MAX_SUPPORTED_TICKS_EXCLUSIVE - EPOCH_TICKS) // DOTNET_TICKS_PER_SECOND
        header = parse_header(
            as_frame(
                mail_frame(
                    key,
                    before=("sender@example.com", "张伟"),
                    after=("主题", "<overflow@example.test>", "主题"),
                    tail=verified_ticks(key, upper_seconds),
                ),
                key=key,
            )
        )
        self.assertIsNone(header["sent_unix"])
        self.assertEqual(header["time_method"], "unresolved")

    def test_conflicting_verified_271_times_remain_unresolved(self):
        key = 35
        header = parse_header(
            as_frame(
                mail_frame(
                    key,
                    before=("sender@example.com", "张伟"),
                    after=("主题", "<conflict@example.test>", "主题"),
                    tail=verified_ticks(key, 1_893_456_000)
                    + verified_ticks(key, 1_893_456_001),
                ),
                key=key,
            )
        )
        self.assertIsNone(header["sent_unix"])
        self.assertEqual(header["latest_tick_unix"], 1_893_456_001)
        self.assertEqual(header["time_method"], "unresolved")

    def test_multiple_sender_addresses_are_unresolved(self):
        header = parse_header(
            as_frame(
                mail_frame(
                    24,
                    before=("张伟", "张伟", "first@example.com", "second@example.com"),
                    after=("<multiple-senders@example.test>",),
                ),
                key=24,
            )
        )

        self.assertIsNone(header["sender"])
        self.assertIsNone(header["sender_name"])
        self.assertIsNone(header["subject"])
        self.assertEqual(header["sender_status"], "unresolved")

    def test_typed_references_returns_unique_nonzero_targets(self):
        key, other = 0xA01, 0xA02
        raw = inline_body_frame(
            key,
            "<html><body>中文正文😀</body></html>",
            owner_refs=(key, 0, other, key),
        )
        frame = as_frame(raw, kind=202, key=key)

        self.assertEqual(typed_references(frame, 201), {key, other})
        self.assertEqual(typed_references(frame, 202), {key})

    def test_inline_body_uses_declared_full_representation(self):
        key = 0xA11
        full = "<html><body>完整正文中文😀。</body></html>".encode()
        unique = "<html><body>唯一版本。</body></html>".encode()
        frame = as_frame(
            inline_body_frame(key, full, unique_html=unique), kind=202, key=key
        )

        parsed = parse_inline_body(frame)

        self.assertEqual(parsed["data"], full)
        self.assertEqual(parsed["evidence"]["owner_offset"], 68)
        self.assertEqual(parsed["evidence"]["full_data_offset"], 1898)
        self.assertEqual(parsed["evidence"]["full_length"], len(full))
        self.assertNotEqual(parsed["evidence"]["unique_offset"], 0)

    def test_inline_body_accepts_variant_one_standard_payload_base(self):
        key = 0xA111
        full = "<html><body>旧布局完整正文中文😀。</body></html>".encode()
        unique = "<html><body>旧布局唯一版本。</body></html>".encode()
        raw = inline_body_frame(
            key,
            full,
            unique_html=unique,
            extra_pool_variant=1,
        )

        parsed = parse_inline_body(as_frame(raw, kind=202, key=key))

        self.assertEqual(parsed["data"], full)
        self.assertEqual(parsed["evidence"]["payload_base"], 1898)
        self.assertEqual(parsed["evidence"]["extra_pool_field"]["variant"], 1)
        self.assertEqual(parsed["evidence"]["extra_pool_field"]["offset"], 0)
        self.assertEqual(parsed["evidence"]["extra_pool_field"]["length"], 0)
        self.assertNotIn("requires_preview_match", parsed["evidence"])

    def test_inline_body_accepts_variant_five_standard_with_preview_guard(self):
        key = 0xA112
        full = "<html><body>新布局完整正文中文😀。</body></html>".encode()
        unique = "<html><body>新布局唯一版本。</body></html>".encode()
        raw = inline_body_frame(
            key,
            full,
            unique_html=unique,
            extra_pool_variant=INLINE_BODY_VARIANT_5,
        )

        parsed = parse_inline_body(as_frame(raw, kind=202, key=key))

        self.assertEqual(parsed["data"], full)
        self.assertEqual(parsed["evidence"]["payload_base"], 1898)
        self.assertTrue(parsed["evidence"]["requires_preview_match"])

        metadata = parse_inline_body(as_frame(raw, kind=202, key=key), include_data=False)
        self.assertNotIn("data", metadata)
        self.assertTrue(metadata["evidence"]["requires_preview_match"])

    def test_inline_body_uses_verified_extra_pool_payload_base(self):
        key = 0xA12
        prefix = "<html><body>特殊完整正文中文😀"
        suffix = "</body></html>"
        full = (
            prefix.encode()
            + b"x" * (421 - len(prefix.encode()) - len(suffix.encode()))
            + suffix.encode()
        )
        unique = "<html><body>唯一版本中文😀。</body></html>".encode()
        extra_pool = b"\x00\xff\x01\xfe\x80\x81\x00\xfe\xff\x10\x11\x00\x90\xa0\xb0\xc0"
        self.assertEqual(len(full), 421)
        self.assertEqual(len(extra_pool), 16)

        raw = inline_body_frame(
            key,
            full,
            unique_html=unique,
            extra_pool=extra_pool,
            full_offset=863,
        )
        parsed = parse_inline_body(as_frame(raw, kind=202, key=key))

        self.assertEqual(raw[1898:1914], extra_pool)
        self.assertNotEqual(raw[1898 + 863:1898 + 863 + 16], full[:16])
        self.assertEqual(parsed["data"], full)
        self.assertEqual(parsed["evidence"]["payload_base"], 1914)
        self.assertEqual(parsed["evidence"]["full_data_offset"], 2777)
        self.assertEqual(parsed["evidence"]["full_length"], 421)
        self.assertEqual(parsed["evidence"]["unique_offset"], 863 + 421)
        self.assertEqual(
            parsed["evidence"]["extra_pool_field"],
            {
                "variant": 1,
                "offset": 18,
                "length": 16,
                "offset_field": INLINE_BODY_EXTRA_POOL_OFFSET,
                "length_field": INLINE_BODY_EXTRA_POOL_LENGTH,
            },
        )
        self.assertNotIn("requires_preview_match", parsed["evidence"])

    def test_inline_body_uses_new_variant_extra_pool_payload_bases(self):
        pool = b"\x00\xff\x01\xfe\x80\x81\x00\xfe\xff\x10\x11\x00\x90\xa0\xb0\xc0"
        for offset, variant in enumerate((INLINE_BODY_STANDARD_VARIANT,
                                           INLINE_BODY_VARIANT_5), 1):
            key = 0xA120 + offset
            full = f"<html><body>变体{variant}完整中文正文😀。</body></html>".encode()
            unique = f"<html><body>变体{variant}唯一版本。</body></html>".encode()
            raw = inline_body_frame(
                key,
                full,
                unique_html=unique,
                extra_pool=pool,
                extra_pool_variant=variant,
                full_offset=23,
            )

            self.assertEqual(raw[1898:1914], pool)
            parsed = parse_inline_body(as_frame(raw, kind=202, key=key))

            self.assertEqual(parsed["data"], full)
            self.assertEqual(parsed["evidence"]["payload_base"], 1914)
            self.assertTrue(parsed["evidence"]["requires_preview_match"])

    def test_inline_body_rejects_unknown_extra_pool_layouts(self):
        key = 0xA13
        body = b"<html><body>unknown layout</body></html>"
        pool = b"\x00\xff\x01\xfe\x80\x81\x00\xfe\xff\x10\x11\x00\x90\xa0\xb0\xc0"
        cases = (
            inline_body_frame(key, body, extra_pool=pool, extra_pool_variant=6),
            inline_body_frame(key, body, extra_pool=pool, extra_pool_offset=19),
            inline_body_frame(key, body, extra_pool=pool, extra_pool_length=15),
            inline_body_frame(key, body, extra_pool_variant=2),
        )

        for raw in cases:
            self.assertIsNone(parse_inline_body(as_frame(raw, kind=202, key=key)))

    def test_inline_body_new_variants_retain_owner_and_bounds_guards(self):
        key = 0xA14
        body = "<html><body>新变体中文正文</body></html>"
        pool = b"\x00\xff\x01\xfe\x80\x81\x00\xfe\xff\x10\x11\x00\x90\xa0\xb0\xc0"
        cases = (
            inline_body_frame(
                key,
                body,
                extra_pool_variant=INLINE_BODY_VARIANT_5,
                owner=key + 1,
            ),
            inline_body_frame(
                key,
                body,
                extra_pool=pool,
                extra_pool_variant=INLINE_BODY_STANDARD_VARIANT,
                full_offset=0x100001,
            ),
            inline_body_frame(
                key,
                body,
                extra_pool=pool,
                extra_pool_variant=INLINE_BODY_VARIANT_5,
                extra_pool_offset=17,
            ),
            inline_body_frame(
                key,
                body,
                extra_pool=pool,
                extra_pool_variant=INLINE_BODY_VARIANT_5,
                extra_pool_length=15,
            ),
        )

        for raw in cases:
            self.assertIsNone(parse_inline_body(as_frame(raw, kind=202, key=key)))

    def test_inline_body_rejects_wrong_kind_owner_or_shared_class(self):
        key = 0xA21
        valid = inline_body_frame(key, "<html><body>正文</body></html>")
        cases = (
            as_frame(valid, kind=201, key=key),
            as_frame(inline_body_frame(key, b"body", owner=0), kind=202, key=key),
            as_frame(inline_body_frame(key, b"body", owner=key + 1), kind=202, key=key),
            as_frame(
                inline_body_frame(key, b"body", owner_refs=(key, key)),
                kind=202,
                key=key,
            ),
            as_frame(
                inline_body_frame(key, b"body", extra_classes=("IPM.Note",)),
                kind=202,
                key=key,
            ),
            as_frame(
                inline_body_frame(key, b"body", owner_offset=69),
                kind=202,
                key=key,
            ),
        )

        for frame in cases:
            self.assertIsNone(parse_inline_body(frame))

    def test_inline_body_rejects_shared_multiple_entry_ids(self):
        key = 0xA22
        raw = bytearray(inline_body_frame(key, "<html><body>正文</body></html>"))
        raw.extend(exchange_entry(1) + exchange_entry(2))

        self.assertIsNone(parse_inline_body(as_frame(bytes(raw), kind=202, key=key)))

    def test_inline_owner_ignores_odd_cross_field_marker_but_rejects_real_conflicts(self):
        key = 0xA23
        target = 0xE900000000000000

        def add_marker(raw, position, value):
            struct.pack_into("<H", raw, position, 201)
            raw[position + 2:position + 10] = b"\0" * 8
            struct.pack_into("<Q", raw, position + 10, value)

        odd_window = bytearray(inline_body_frame(key, "<html><body>正文</body></html>"))
        add_marker(odd_window, 1641, target)
        odd_frame = as_frame(bytes(odd_window), kind=202, key=key)
        self.assertIn(target, typed_references(odd_frame, 201))
        self.assertIsNotNone(parse_inline_body(odd_frame))

        even_conflict = bytearray(inline_body_frame(key, "<html><body>正文</body></html>"))
        add_marker(even_conflict, 1700, key + 1)
        self.assertIsNone(parse_inline_body(as_frame(bytes(even_conflict), kind=202, key=key)))

        variable_conflict = bytearray(
            inline_body_frame(key, b"<html><body>" + b"x" * 64 + b"</body></html>")
        )
        add_marker(variable_conflict, 1900, key + 2)
        self.assertIsNone(parse_inline_body(as_frame(bytes(variable_conflict), kind=202, key=key)))

    def test_inline_body_rejects_unverified_descriptor_fields(self):
        key = 0xA31
        body = b"<html><body>descriptor</body></html>"
        mutations = []

        no_flag = bytearray(inline_body_frame(key, body))
        struct.pack_into("<I", no_flag, 1674, len(body))
        mutations.append(no_flag)

        zero_length = bytearray(inline_body_frame(key, body))
        struct.pack_into("<I", zero_length, 1674, 0x80000000)
        mutations.append(zero_length)

        outside = bytearray(inline_body_frame(key, body))
        struct.pack_into("<I", outside, 1674, 0x80000000 | 0x100000)
        mutations.append(outside)

        wrong_codepage = bytearray(inline_body_frame(key, body))
        struct.pack_into("<I", wrong_codepage, 1658, 1200)
        mutations.append(wrong_codepage)

        wrong_class_length = bytearray(inline_body_frame(key, body))
        struct.pack_into("<I", wrong_class_length, 1690, 16)
        mutations.append(wrong_class_length)

        wrong_class_offset = bytearray(inline_body_frame(key, body, class_offset=1882))
        mutations.append(wrong_class_offset)

        for raw in mutations:
            self.assertIsNone(
                parse_inline_body(as_frame(bytes(raw), kind=202, key=key))
            )

    def test_inline_body_does_not_fallback_to_unique_descriptor(self):
        key = 0xA41
        full = b"<html><body>declared-full</body></html>"
        raw = bytearray(inline_body_frame(key, full, unique_html=b"unique"))
        # The unique descriptor is deliberately unusable.  The full field is
        # the only accepted body representation.
        struct.pack_into("<II", raw, 1678, 0xFFFFFFFF, 0)

        parsed = parse_inline_body(as_frame(bytes(raw), kind=202, key=key))

        self.assertEqual(parsed["data"], full)

    def test_hxperson_parses_utf16le_name_and_to_cc_roles(self):
        for role, owner in (("to", 0xC001), ("cc", 0xC002)):
            key = owner + 1
            raw = person_frame(
                key,
                owner,
                role=role,
                name="张伟😀",
                address=f"person-{role}@example.com",
                variant=2,
            )
            frame = next(iter_frames(SimpleNamespace(offset=19, data=raw)))

            self.assertEqual(person_parent(frame), (role, owner))
            parsed = parse_person(frame)
            self.assertEqual(
                parsed["key"], key,
            )
            self.assertEqual(
                {parsed[field] for field in ("owner", "role", "name", "address")},
                {owner, role, "张伟😀", f"person-{role}@example.com"},
            )
            self.assertEqual(parsed["evidence"]["kind"], 85)
            self.assertEqual(parsed["evidence"]["frame_bytes"], len(raw))
            self.assertEqual(parsed["evidence"]["block"], 19)
            self.assertEqual(parsed["evidence"]["frame_offset"], 0)
            self.assertEqual(len(parsed["evidence"]["frame_sha256"]), 64)
            self.assertEqual(
                parsed["evidence"]["method"], "verified_hxperson_layout_utf16le"
            )

    def test_hxperson_rejects_parent_type_and_field_contamination(self):
        valid = person_frame(0xD01, 0xD10)
        mutations = []

        wrong_parent = bytearray(valid)
        struct.pack_into("<I", wrong_parent, 38, 413)
        struct.pack_into("<H", wrong_parent, 68, 413)
        mutations.append(wrong_parent)

        bcc_parent = bytearray(valid)
        struct.pack_into("<I", bcc_parent, 38, 206)
        struct.pack_into("<H", bcc_parent, 68, 206)
        mutations.append(bcc_parent)

        for offset in (22, 66, 70):
            polluted = bytearray(valid)
            polluted[offset] = 1
            mutations.append(polluted)

        for offset, value in ((42, 0xD11), (50, 0xD11), (58, 0xD11), (78, 0xD11)):
            polluted = bytearray(valid)
            struct.pack_into("<Q", polluted, offset, value)
            mutations.append(polluted)

        for raw in mutations:
            frame = next(iter_frames(SimpleNamespace(offset=23, data=bytes(raw))))
            self.assertIsNone(person_parent(frame))
            self.assertIsNone(parse_person(frame))

    def test_hxperson_rejects_descriptor_boundary_and_content_failures(self):
        valid = person_frame(0xD21, 0xD22)
        mutations = []

        for offset, value in (
            (306, 0x80000000),
            (314, 0x80000000),
            (322, 0x80000000 | 54),
            (326, 4),
        ):
            polluted = bytearray(valid)
            struct.pack_into("<I", polluted, offset, value)
            mutations.append(polluted)

        bad_separator = bytearray(valid)
        bad_separator[360 + 53 - 5] = 0
        mutations.append(bad_separator)

        bad_email_terminator = bytearray(valid)
        bad_email_terminator[360 + 53 + len(wide("zhang@example.com")) - 2] = 1
        mutations.append(bad_email_terminator)

        bad_offset = bytearray(valid)
        struct.pack_into("<I", bad_offset, 310, 0x7FFFFFF0)
        mutations.append(bad_offset)

        for raw in mutations:
            frame = next(iter_frames(SimpleNamespace(offset=29, data=bytes(raw))))
            self.assertIsNone(parse_person(frame))

        for name, address in (
            ("pii:display", "valid@example.com"),
            ("显示\x01名", "valid@example.com"),
            ("显示名", "pii:address@example.com"),
            ("显示名", "invalid@example"),
            ("显示名", "invalid@@example.com"),
        ):
            raw = person_frame(0xD30, 0xD31, name=name, address=address)
            frame = next(iter_frames(SimpleNamespace(offset=29, data=raw)))
            self.assertIsNone(parse_person(frame))

    def test_hxperson_kind_support_and_nested_frame_boundaries(self):
        key, owner = 0xD41, 0xD42
        raw = person_frame(key, owner)
        frames = list(iter_frames(SimpleNamespace(offset=31, data=raw)))
        self.assertEqual([(frame.kind, frame.key) for frame in frames], [(85, key)])

        wrong_kind = bytearray(raw)
        struct.pack_into("<H", wrong_kind, 20, 999)
        self.assertEqual(
            list(iter_frames(SimpleNamespace(offset=31, data=bytes(wrong_kind)))), []
        )
        self.assertIsNone(parse_person(as_frame(raw, kind=201, key=key)))

        nested = serialized_frame(201, 0xD43, b"prefix" + raw + b"suffix")
        frames = list(iter_frames(SimpleNamespace(offset=31, data=nested)))
        self.assertEqual([(frame.kind, frame.key) for frame in frames], [(201, 0xD43)])


if __name__ == "__main__":
    unittest.main()
