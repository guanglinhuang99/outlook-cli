"""Windows-compatible read-only RPC over the macOS cache service."""
from __future__ import annotations

import base64
from datetime import datetime
import json
import os
from pathlib import Path
import sys

import outlook_local as core
from outlook_service import MailService
from mcp_server import _readline_limited


class ReadError(ValueError):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def fail(code, message):
    raise ReadError(code, message)


def encode_id(kind, key, snapshot):
    raw = json.dumps({"v": 3, "backend": "mac-cache", "kind": kind,
                      "key": key, "snapshot": snapshot}, separators=(",", ":")).encode()
    return base64.urlsafe_b64encode(raw).decode().rstrip("=")


def decode_id(value, kind):
    try:
        if not isinstance(value, str) or len(value) > 16000:
            raise ValueError()
        raw = base64.b64decode(value + "=" * (-len(value) % 4), altchars=b"-_", validate=True)
        payload = json.loads(raw)
        if (payload.get("v") != 3 or payload.get("backend") != "mac-cache"
                or payload.get("kind") != kind or not isinstance(payload.get("key"), str)
                or not isinstance(payload.get("snapshot"), str)):
            raise ValueError()
        return payload
    except (ValueError, TypeError, AttributeError, UnicodeError):
        fail("BACKEND_ID_MISMATCH", "Expected a mac-cache ID; list stores or messages again on this platform")


def number(params, name, default, minimum, maximum):
    value = params.get(name, default)
    if type(value) is not int or not minimum <= value <= maximum:
        fail("INVALID_ARGUMENT", f"{name} must be an integer in {minimum}..{maximum}")
    return value


def timestamp(value):
    if value is None:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if parsed.utcoffset() is None:
            raise ValueError()
        return parsed.timestamp()
    except (TypeError, ValueError, AttributeError):
        fail("INVALID_ARGUMENT", "after/before must be ISO timestamps with timezone")


FIELDS = {
    "health": set(), "stores": set(),
    "folders": {"folder", "offset", "limit"},
    "messages": {"folder", "offset", "limit", "unread", "after", "before", "scan_limit"},
    "search": {"folder", "offset", "limit", "unread", "after", "before", "scan_limit", "query", "search_body"},
    "read": {"id", "format", "body_offset", "max_chars"},
    "attachments": {"id"},
    "read_attachment": {"id", "index", "max_bytes", "byte_offset", "chunk_bytes"},
}


class CommonReader:
    def __init__(self, service=None):
        self.service = service or MailService(
            state=os.environ.get("OUTLOOK_MAC_STATE_DIR", str(Path.home() / "Library/Caches/outlook-cli/mac-cache")),
            profile=os.environ.get("OUTLOOK_MAC_PROFILE"),
        )

    @staticmethod
    def groups(index):
        result = {}
        for folder in index["folders"]:
            if folder.get("is_inbox") or folder.get("is_sent"):
                key = "store:" + (folder.get("parent_id") or folder["id"])
                result.setdefault(key, []).append(folder)
        return result

    @staticmethod
    def active_groups(index, groups):
        """Return account groups with cached messages anywhere in their trees."""
        folders = index["folders"]
        active = []
        for rows in groups.values():
            roots = {row["id"] for row in rows}
            descendants = [folder for folder in folders
                           if folder.get("scope_root_id") in roots]
            if any(folder.get("cached_messages") for folder in descendants):
                active.append(rows)
        return active

    def folder_key(self, value, index):
        folders = {row["id"]: row for row in index["folders"]}
        if value in ("deleted", "drafts", "junk", "outbox"):
            fail("SCOPE_DENIED", "macOS exposes only verified Inbox, Sent Items and their known descendants")
        if value in ("inbox", "sent"):
            groups = self.groups(index)
            active = self.active_groups(index, groups)
            choices = active or list(groups.values())
            if len(choices) != 1:
                fail("FOLDER_AMBIGUOUS", "Select a store and explicit folder ID; no unique cached default account")
            roots = [row for row in choices[0] if row.get("is_inbox" if value == "inbox" else "is_sent")]
            if len(roots) != 1:
                fail("FOLDER_UNAVAILABLE", "The requested special folder is not verified in this cache")
            return roots[0]["id"]
        payload = decode_id(value, "folder")
        key = payload["key"]
        if key not in folders and key not in self.groups(index):
            fail("SCOPE_DENIED", "Folder is outside the verified Inbox/Sent scope")
        return key

    def folder_info(self, key, index):
        groups = self.groups(index)
        row = next((f for f in index["folders"] if f["id"] == key), None)
        children = groups[key] if key in groups else [f for f in index["folders"] if f.get("parent_id") == key]
        return {"id": encode_id("folder", key, index["snapshot"]),
                "name": row["name"] if row else "Outlook cache " + core.sha(key.encode())[:8],
                "path": row["path"] if row else "/", "item_count": row["cached_messages"] if row else 0,
                "unread_count": None, "child_count": len(children), "counts_scope": "local_index",
                "is_inbox": bool(row.get("is_inbox")) if row else False,
                "is_sent": bool(row.get("is_sent")) if row else False,
                "name_status": row.get("evidence", {}).get("name_status", "unavailable") if row else "unavailable"}

    @staticmethod
    def metadata(row, index):
        native = core.metadata(row)
        address = row.get("sender") or None
        return {"id": encode_id("message", row["id"], index["snapshot"]),
                "subject": row["subject"] or None, "subject_status": row["subject_status"],
                "sender": {"name": row.get("sender_name") or None, "address": address,
                           "address_type": "SMTP" if address and "@" in address else None,
                           "smtp_address": address if address and "@" in address else None},
                "received_at": native["received_at"], "sent_at": native["sent_at"],
                "received_time_status": native["received_time_status"],
                "received_time_method": native["received_time_method"], "unread": None,
                "size_bytes": None, "download_state": None, "preview": row.get("preview"),
                "body_status": row["body"]["status"], "attachment_count": len(row["attachments"]),
                "attachments_complete": False,
                "folder_id": encode_id("folder", row["folder_id"], index["snapshot"]),
                "folder_path": row["folder_path"], "folder_status": row["folder_status"]}

    @staticmethod
    def attachment_metadata(row):
        return [{"index": position, "filename": a.get("name") or a.get("cache_name"),
                 "size_bytes": a.get("fingerprint", {}).get("size"),
                 "type": 1 if a["status"] == "cached_at_snapshot" else None,
                 "status": a["status"]} for position, a in enumerate(row["attachments"], 1)]

    def perform(self, method, params):
        if method not in FIELDS or not isinstance(params, dict) or set(params) - FIELDS[method]:
            fail("INVALID_ARGUMENT", "Unknown operation or parameters")
        if "unread" in params:
            fail("UNSUPPORTED_FILTER", "Unread state is not verified by the macOS cache parser")
        if params.get("search_body"):
            fail("UNSUPPORTED_FILTER", "Common macOS search supports subject/sender; cached body text is available via read")
        if params.get("format", "text") != "text":
            fail("UNSUPPORTED_FORMAT", "Common macOS reads currently return text only")
        value = params.get("id", params.get("folder", "inbox"))
        kind = "message" if "id" in params else "folder"
        identity = None if value in ("inbox", "sent", "deleted", "drafts", "junk", "outbox") else decode_id(value, kind)
        directory, index = self.service._current(identity["snapshot"] if identity else None)
        groups = self.groups(index)
        context = {"backend": "mac-cache", "data_scope": "verified_inbox_sent_cache", "date_basis": "received_at",
                   "snapshot": index["snapshot"], "captured_at": index["created_at"],
                   "coverage": index["summary"]["coverage"]}
        if method == "health":
            return {**context, "connected": True, "offline": True, "outlook_required": False,
                    "store_count": len(groups), "readable_store_count": len(groups),
                    "outlook_version": None, "cache_summary": index["summary"],
                    "capabilities": {"unread_filter": False, "body_search": False, "html_body": False,
                                     "received_time": True, "sent_time": True},
                    "limitations": ["Only recoverable cache in verified Inbox/Sent folders is exposed.",
                                    "Folder names may be redacted; membership is a local observation, not live server state."]}
        if method == "stores":
            return {**context, "stores": [{"name": self.folder_info(key, index)["name"],
                                           "root_id": encode_id("folder", key, index["snapshot"]),
                                           "readable": True, "is_data_file": True, "exchange_store_type": None}
                                          for key in sorted(groups)]}
        if method == "folders":
            key = self.folder_key(params.get("folder", "inbox"), index)
            rows = groups[key] if key in groups else [f for f in index["folders"] if f.get("parent_id") == key]
            offset = number(params, "offset", 0, 0, 10000000)
            end = min(len(rows), offset + number(params, "limit", 25, 1, 200))
            return {**context, "parent": self.folder_info(key, index),
                    "folders": [self.folder_info(row["id"], index) for row in rows[offset:end]],
                    "total": len(rows), "next_offset": end if end < len(rows) else None}
        if method in ("messages", "search"):
            key = self.folder_key(params.get("folder", "inbox"), index)
            after, before = timestamp(params.get("after")), timestamp(params.get("before"))
            if after is not None and before is not None and after >= before:
                fail("INVALID_ARGUMENT", "after must be earlier than before")
            folder_rows = [row for row in index["items"] if row["folder_id"] == key]
            unknown_times = sum(row.get("received_unix") is None for row in folder_rows)
            dated = after is not None or before is not None
            rows = sorted((row for row in folder_rows
                           if not dated or (row.get("received_unix") is not None
                               and (after is None or row["received_unix"] >= after)
                               and (before is None or row["received_unix"] < before))),
                          key=core.received_sort_key)
            offset = number(params, "offset", 0, 0, 10000000)
            limit = number(params, "limit", 25, 1, 200)
            scan = number(params, "scan_limit", 500, 1, 2000)
            query = params.get("query", "")
            if method == "search" and (not isinstance(query, str) or not 1 <= len(query) <= 500):
                fail("INVALID_ARGUMENT", "query must be 1..500 characters")
            matched, position = [], min(offset, len(rows))
            while position < len(rows) and position - offset < scan and len(matched) < limit:
                row = rows[position]
                position += 1
                if method != "search" or any(query.casefold() in (row.get(field) or "").casefold()
                                             for field in ("subject", "sender", "sender_name")):
                    matched.append(self.metadata(row, index))
            return {**context, "messages": matched, "scanned": max(0, position - offset),
                    "filtered_item_count": len(rows), "next_offset": position if position < len(rows) else None,
                    "complete": position >= len(rows), "scope": "single_folder",
                    "pagination": "live_offset_not_snapshot", "body_search_incomplete": False,
                    "received_time_unknown": unknown_times,
                    "date_filter_incomplete": bool(dated and unknown_times)}
        if not identity or kind != "message":
            fail("INVALID_ID", "Expected a message ID")
        row = next((row for row in index["items"] if row["id"] == identity["key"]), None)
        if row is None:
            fail("SCOPE_DENIED", "Message is outside the verified Inbox/Sent scope")
        attachments = self.attachment_metadata(row)
        if method == "attachments":
            return {**context, "attachments": attachments, "attachments_complete": False}
        if method == "read":
            offset = number(params, "body_offset", 0, 0, 100000000)
            maximum = number(params, "max_chars", 20000, 1, 200000)
            page = core.read_body_page(directory, index, row["id"], offset, maximum)
            text = None if page is None else page["text"]
            total = None if page is None else page["total_chars"]
            start = 0 if page is None else page["offset"]
            next_offset = None if page is None else page["next_offset"]
            return {**context, **self.metadata(row, index), "to": None, "cc": None,
                    "attachments": attachments, "body": {"format": "text", "derived_from": "html" if text is not None else None,
                    "text": text, "available": text is not None, "total_chars": total, "offset": start,
                    "next_offset": next_offset, "truncated": next_offset is not None, "source_truncated": False,
                    "status": row["body"]["status"], "diagnosis": row["body"]["diagnosis"]}}
        position = number(params, "index", 0, 1, 10000)
        if position > len(attachments):
            fail("INVALID_ARGUMENT", "Attachment index is out of range")
        maximum = number(params, "max_bytes", 10485760, 1, 10485760)
        attachment = row["attachments"][position - 1]
        if attachment["status"] != "cached_at_snapshot":
            fail("ATTACHMENT_NOT_CACHED", "Attachment bytes are not cached")
        if attachment["fingerprint"]["size"] > maximum:
            fail("ATTACHMENT_TOO_LARGE", "Attachment exceeds max_bytes")
        offset = number(params, "byte_offset", 0, 0, 10485760)
        chunk = number(params, "chunk_bytes", 65536, 1, 1048576)
        result = self.service._attachment(directory, index, {"id": row["id"], "attachment_id": attachment["id"]})
        output = Path(result["output"])
        expected_bytes = result.get("bytes")
        expected_fingerprint = result.get("_export_fingerprint")
        if (type(expected_bytes) is not int or expected_bytes < 0
                or not isinstance(expected_fingerprint, dict)
                or expected_fingerprint.get("size") != expected_bytes):
            fail("ATTACHMENT_CHANGED", "Attachment integrity changed")
        if expected_bytes > maximum:
            fail("ATTACHMENT_TOO_LARGE", "Attachment exceeds max_bytes")
        try:
            before = core.fingerprint(output)
        except OSError:
            fail("ATTACHMENT_CHANGED", "Attachment integrity changed")
        if before != expected_fingerprint:
            fail("ATTACHMENT_CHANGED", "Attachment integrity changed")
        offset = min(offset, expected_bytes)
        end = min(offset + chunk, expected_bytes)
        try:
            with output.open("rb") as handle:
                handle.seek(offset)
                data = handle.read(end - offset)
            after = core.fingerprint(output)
        except OSError:
            fail("ATTACHMENT_CHANGED", "Attachment integrity changed")
        if after != expected_fingerprint or len(data) != end - offset:
            fail("ATTACHMENT_CHANGED", "Attachment integrity changed")
        return {**context, "filename": attachments[position - 1]["filename"], "size_bytes": expected_bytes,
                "sha256": result["sha256"], "encoding": "base64", "offset": offset,
                "chunk_size": end - offset, "next_offset": end if end < expected_bytes else None,
                "data": base64.b64encode(data).decode()}


def main():
    reader = CommonReader()
    while True:
        line, oversized = _readline_limited(sys.stdin.buffer)
        if line is None:
            return
        request = None
        try:
            if oversized:
                fail("INVALID_ARGUMENT", "Request exceeds 1 MiB")
            request = json.loads(line)
            if not isinstance(request, dict) or type(request.get("id")) is not int:
                fail("INVALID_ARGUMENT", "Expected an integer request ID")
            result = reader.perform(request.get("method"), request.get("params", {}))
            response = {"id": request["id"], "result": result}
        except Exception as error:
            code = getattr(error, "code", "MAC_CACHE_ERROR")
            # Source OS errors can include private message cache paths.
            message = str(error) if isinstance(error, (ValueError, ReadError)) else "Could not read the local Outlook cache"
            response = {"id": request.get("id") if isinstance(request, dict) else None,
                        "error": {"code": code, "message": message}}
        print(json.dumps(response, ensure_ascii=True), flush=True)


if __name__ == "__main__":
    main()
