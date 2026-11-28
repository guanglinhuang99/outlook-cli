"""Windows-compatible read-only RPC over the macOS cache service."""
from __future__ import annotations

import base64
from copy import deepcopy
from datetime import date, datetime
import json
import os
from pathlib import Path
import sys

from python_version import answer_unsupported, unsupported_message

if __name__ == "__main__" and unsupported_message():
    answer_unsupported(sys.stdin, sys.stdout, unsupported_message())
    raise SystemExit(0)

import outlook_local as core
from outlook_service import BUNDLE_SOURCE_BYTES, BUNDLE_VERIFY_BYTES, TOKEN, MailService
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


BUNDLE_OPTIONS = {"folder", "include_subfolders", "mode", "limit", "cursor", "query", "field",
                  "person", "person_role", "search_body", "has_cached_attachments", "unread", "max_chars",
                  "text_budget_chars"}
BUNDLE_CURSOR = "workflow-v1"


def bundle_day(value, name):
    try:
        parsed = date.fromisoformat(value).isoformat()
        if value != parsed or not "1970-01-01" <= parsed <= "9999-12-30":
            raise ValueError()
        return parsed
    except (TypeError, ValueError):
        fail("INVALID_ARGUMENT", f"{name} must be YYYY-MM-DD in 1970-01-01..9999-12-30")


def unpack_bundle_cursor(value, method):
    if not isinstance(value, str) or not 1 <= len(value) <= 4096:
        fail("INVALID_ARGUMENT", "cursor must be a nonempty workflow cursor")
    pieces = value.split("|", 3)
    if len(pieces) != 4 or pieces[:2] != [BUNDLE_CURSOR, method]:
        fail("INVALID_ARGUMENT", "cursor belongs to another workflow or is invalid")
    _, _, day, native = pieces
    if day and method != "today_bundle":
        fail("INVALID_ARGUMENT", "cursor belongs to another workflow or is invalid")
    parsed = native.split(".", 3)
    if len(parsed) != 4 or parsed[0] != "v1" or TOKEN.fullmatch(parsed[1]) is None:
        fail("INVALID_ARGUMENT", "invalid workflow cursor")
    _, separator, offset = parsed[3].partition(":")
    if not separator or not offset.isascii() or not offset.isdecimal():
        fail("INVALID_ARGUMENT", "invalid workflow cursor")
    return day, native, parsed[1]


FIELDS = {
    "health": set(), "stores": set(),
    "folders": {"folder", "offset", "limit"},
    "messages": {"folder", "offset", "limit", "unread", "after", "before", "scan_limit"},
    "search": {"folder", "offset", "limit", "unread", "after", "before", "scan_limit", "query", "search_body", "person", "person_role"},
    "read": {"id", "format", "body_offset", "max_chars"},
    "attachments": {"id"},
    "read_attachment": {"id", "index", "max_bytes", "byte_offset", "chunk_bytes"},
    "context_bundle": BUNDLE_OPTIONS | {"from_date", "to_date"},
    "today_bundle": BUNDLE_OPTIONS | {"date"},
    "messages_bundle": {"ids", "max_chars", "text_budget_chars"},
}


class CommonReader:
    def __init__(self, service=None):
        self.service = service or MailService(
            state=os.environ.get("OUTLOOK_MAC_STATE_DIR", str(Path.home() / "Library/Caches/outlook-cli/mac-cache")),
            profile=os.environ.get("OUTLOOK_MAC_PROFILE"),
            folder_scope=os.environ.get("OUTLOOK_MAC_FOLDER_SCOPE", "inbox-sent"),
        )

    @staticmethod
    def groups(index):
        result = {}
        account = index.get("scope_policy") == "verified-account-cache"
        for folder in index["folders"]:
            is_root = (folder["id"] == folder.get("scope_root_id") if account
                       else folder.get("is_inbox") or folder.get("is_sent"))
            if is_root:
                key = "store:" + (folder.get("account_root_id") or folder.get("parent_id") or folder["id"])
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
        specials = {"inbox", "sent", "deleted", "drafts", "junk", "outbox"}
        if value in specials - {"inbox", "sent"} and index.get("scope_policy") != "verified-account-cache":
            fail("SCOPE_DENIED", "macOS exposes only verified Inbox, Sent Items and their known descendants")
        if value in specials:
            groups = self.groups(index)
            flag = "is_" + value
            def matching_aliases(rows):
                if index.get("scope_policy") == "verified-account-cache":
                    root_ids = {row["id"] for row in rows}
                    return [row for row in folders.values()
                            if row.get(flag) and row.get("scope_root_id") in root_ids]
                return [row for row in rows if row.get(flag)]
            if index.get("scope_policy") == "verified-account-cache":
                groups = {key: rows for key, rows in groups.items() if matching_aliases(rows)}
            choices = self.active_groups(index, groups) or list(groups.values())
            if len(choices) > 1:
                fail("FOLDER_AMBIGUOUS", "Select a store and explicit folder ID; no unique cached default account")
            roots = matching_aliases(choices[0]) if choices else []
            if len(roots) != 1:
                fail("FOLDER_UNAVAILABLE", "The requested special folder is not verified in this cache")
            return roots[0]["id"]
        payload = decode_id(value, "folder")
        key = payload["key"]
        if key not in folders and key not in self.groups(index):
            fail("SCOPE_DENIED", "Folder is outside the verified cache scope")
        return key

    def folder_info(self, key, index):
        groups = self.groups(index)
        row = next((f for f in index["folders"] if f["id"] == key), None)
        children = groups[key] if key in groups else [f for f in index["folders"] if f.get("parent_id") == key]
        return {"id": encode_id("folder", key, index["snapshot"]),
                "name": row["name"] if row else "Outlook cache " + core.sha(key.encode())[:8],
                "path": row["path"] if row else "/", "item_count": row["cached_messages"] if row else 0,
                "unread_count": row.get("known_unread_count") if row and row.get("unread_unknown_count") == 0 else None,
                "known_unread_count": row.get("known_unread_count") if row else None,
                "unread_unknown_count": row.get("unread_unknown_count") if row else None,
                "child_count": len(children), "counts_scope": "local_index",
                "is_inbox": bool(row.get("is_inbox")) if row else False,
                "is_sent": bool(row.get("is_sent")) if row else False,
                "folder_type": row.get("evidence", {}).get("folder_type") if row else None,
                "name_status": row.get("evidence", {}).get("name_status", "unavailable") if row else "unavailable"}

    @staticmethod
    def metadata(row, index):
        native = core.metadata(row)
        address = row.get("sender") or None
        related = dict.fromkeys(mid for mid in row.get("related_message_ids", []) if mid != native["message_id"])
        return {"id": encode_id("message", row["id"], index["snapshot"]),
                "internet_message_id": native["message_id"],
                "related_internet_message_ids": list(related),
                "subject": row["subject"] or None, "subject_status": row["subject_status"],
                "sender": {"name": row.get("sender_name") or None, "address": address,
                           "address_type": "SMTP" if address and "@" in address else None,
                           "smtp_address": address if address and "@" in address else None},
                "received_at": native["received_at"], "sent_at": native["sent_at"],
                "received_time_status": native["received_time_status"],
                "received_time_method": native["received_time_method"], "unread": row.get("unread"),
                "unread_evidence": row.get("unread_evidence"),
                "participants": row.get("participants"),
                "participants_status": row.get("participants_status"),
                "participants_evidence": row.get("participants_evidence"),
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

    @staticmethod
    def cache_context(index):
        account = index and index.get("scope_policy") == "verified-account-cache"
        return {"backend": "mac-cache",
                "data_scope": "verified_account_cache" if account else "verified_inbox_sent_cache",
                "date_basis": "received_at",
                "snapshot": index["snapshot"] if index else None,
                "captured_at": index["created_at"] if index else None,
                "coverage": index["summary"]["coverage"] if index else None}

    def _read(self, directory, index, row, offset, maximum):
        row = {**row, "body": core.resolve_message_body(directory, index, row["id"])}
        page = core.read_body_page(directory, {**index, "items": [row]}, row["id"], offset, maximum)
        text = None if page is None else page["text"]
        total = None if page is None else page["total_chars"]
        start = 0 if page is None else page["offset"]
        next_offset = None if page is None else page["next_offset"]
        display = {role: ("; ".join(f"{p['name']} <{p['address']}>" for p in row['participants'][role])
                          if (row.get('participants_status') or {}).get(role) == 'verified'
                          else None) for role in ('to', 'cc')}
        return {**self.metadata(row, index), **display, "attachments": self.attachment_metadata(row),
                "body": {"format": "text", "derived_from": "html" if text is not None else None,
                         "text": text, "available": text is not None, "total_chars": total,
                         "offset": start, "next_offset": next_offset, "truncated": next_offset is not None,
                         "source_truncated": False, "status": row["body"]["status"],
                         "diagnosis": row["body"]["diagnosis"]}}

    def _bundle_rows(self, directory, index, selections, maximum, budget):
        items, errors, used, complete = [], [], 0, True
        source_budget = {"remaining": BUNDLE_SOURCE_BYTES, "verification_remaining": BUNDLE_VERIFY_BYTES}
        for row, item in selections:
            if used >= budget:
                errors.append({"id": item["id"], "code": "TEXT_BUDGET_EXHAUSTED",
                               "message": "Bundle text budget exhausted"})
                complete = False
            else:
                try:
                    with core.source_read_budget(source_budget):
                        item = {**item, **self._read(directory, index, row, 0, min(maximum, budget - used))}
                    body = item["body"]
                    if body["available"]:
                        end = body["next_offset"] if body["truncated"] else body["total_chars"]
                        used += end - body["offset"]
                    complete = complete and body["available"] and not body["truncated"]
                except (ValueError, OSError):
                    errors.append({"id": item["id"], "code": "BODY_READ_FAILED",
                                   "message": "Could not read the local Outlook body"})
                    complete = False
            items.append(item)
        return {"items": items, "errors": errors, "body_text_chars": used, "body_complete": complete}

    def _messages_bundle(self, params):
        maximum = number(params, "max_chars", 12000, 1, 100000)
        budget = number(params, "text_budget_chars", 60000, 1000, 200000)
        ids = params.get("ids")
        if (not isinstance(ids, list) or not 1 <= len(ids) <= 25
                or any(not isinstance(v, str) or len(v) > 16000 for v in ids)):
            fail("INVALID_ARGUMENT", "ids must contain 1..25 message IDs")
        snapshot, identities, errors = None, [], []
        for identifier in dict.fromkeys(ids):
            try:
                identity = decode_id(identifier, "message")
                if TOKEN.fullmatch(identity["snapshot"]) is None:
                    fail("INVALID_ID", "Invalid message snapshot")
                snapshot = snapshot or identity["snapshot"]
                if identity["snapshot"] != snapshot:
                    fail("SNAPSHOT_MISMATCH", "Message IDs in a bundle must use the same snapshot")
                identities.append((identifier, identity["key"]))
            except ReadError as error:
                errors.append({"id": identifier, "code": error.code, "message": str(error)})
        if snapshot is None:
            return {**self.cache_context(None), "items": [], "errors": errors,
                    "body_text_chars": 0, "body_complete": False}
        directory, index = self.service._current(snapshot)
        rows = {row["id"]: row for row in index["items"]}
        selected = []
        for identifier, key in identities:
            if key not in rows:
                errors.append({"id": identifier, "code": "SCOPE_DENIED",
                               "message": "Message is outside the verified cache scope"})
            else:
                selected.append((rows[key], self.metadata(rows[key], index)))
        result = self._bundle_rows(directory, index, selected, maximum, budget)
        return {**self.cache_context(index), **result, "errors": errors + result["errors"],
                "body_complete": not errors and result["body_complete"]}

    def _context_bundle(self, method, params):
        today = method == "today_bundle"
        mode = params.get("mode", "context" if today else "metadata")
        if mode not in ("metadata", "context"):
            fail("INVALID_ARGUMENT", "mode must be metadata or context")
        limit = number(params, "limit", 25, 1, 25 if mode == "context" else 100)
        maximum = number(params, "max_chars", 4000, 1, 100000)
        budget = number(params, "text_budget_chars", 60000, 1000, 200000)
        for flag in ("include_subfolders", "search_body", "has_cached_attachments", "unread"):
            if flag in params and type(params[flag]) is not bool:
                fail("INVALID_ARGUMENT", f"{flag} must be boolean")
        descendants = params.get("include_subfolders", True)
        field, query = params.get("field", "all"), params.get("query", "")
        if field not in ("all", "subject", "sender", "preview"):
            fail("INVALID_ARGUMENT", "field must be all, subject, sender or preview")
        if not isinstance(query, str) or len(query.strip()) > 500:
            fail("INVALID_ARGUMENT", "query must be a string of at most 500 characters")
        person, role = params.get("person"), params.get("person_role", "any")
        if isinstance(person, str) and person.strip():
            person = person.strip()
        try:
            core.person_matches({}, person, role)
        except ValueError as error:
            fail("INVALID_ARGUMENT", str(error))

        day, native_cursor, snapshot = "", None, None
        if "cursor" in params:
            day, native_cursor, snapshot = unpack_bundle_cursor(params["cursor"], method)
        if today:
            requested = bundle_day(params["date"], "date") if "date" in params else None
            if native_cursor:
                day = bundle_day(day, "cursor date")
                if requested and requested != day:
                    fail("INVALID_ARGUMENT", "date does not match the cursor date")
            else:
                day = requested or datetime.now(core.ZONE).date().isoformat()
            first = last = day
        else:
            first = bundle_day(params["from_date"], "from_date") if "from_date" in params else None
            last = bundle_day(params["to_date"], "to_date") if "to_date" in params else None
            if first and last and first > last:
                fail("INVALID_ARGUMENT", "from_date must not exceed to_date")

        folder = params.get("folder", "inbox")
        if not isinstance(folder, str):
            fail("INVALID_ARGUMENT", "folder must be a folder name or ID")
        if folder not in ("inbox", "sent", "all", "deleted", "drafts", "junk", "outbox"):
            folder_snapshot = decode_id(folder, "folder")["snapshot"]
            if snapshot and folder_snapshot != snapshot:
                fail("INVALID_ARGUMENT", "folder ID and cursor refer to different snapshots")
            snapshot = folder_snapshot
        directory, index = self.service._current(snapshot)
        folder_id, roots = None, None
        if folder != "all":
            folder_id = self.folder_key(folder, index)
            groups = self.groups(index)
            if folder_id in groups:
                roots = tuple(sorted(row["id"] for row in groups[folder_id]))
                folder_id = None
        elif not descendants:
            roots = tuple(sorted(row["id"] for row in index["folders"]
                                 if row["id"] == row.get("scope_root_id")))

        options = {"query": query.strip(), "field": field, "person": person, "person_role": role,
                   "from_date": first, "to_date": last, "cursor": native_cursor, "limit": limit,
                   "folder_id": folder_id, "include_subfolders": descendants,
                   "search_body": params.get("search_body", False)}
        if "has_cached_attachments" in params:
            options["has_cached_attachments"] = params["has_cached_attachments"]
        if "unread" in params:
            options["unread"] = params["unread"]
        try:
            page = self.service._search(options, snapshot=index["snapshot"], root_folders=roots)
        except ValueError as error:
            if native_cursor and "cursor" in str(error):
                fail("INVALID_ARGUMENT", "cursor belongs to another folder or query, or is invalid")
            raise
        rows = {row["id"]: row for row in index["items"]}
        selected = []
        for match in page["items"]:
            row = rows[match["id"]]
            details = {key: match[key] for key in ("matched_fields", "matched_person_roles", "body_snippet") if key in match}
            selected.append((row, {**self.metadata(row, index), **details}))
        if mode == "context" and not page.get("retry_required"):
            result = self._bundle_rows(directory, index, selected, maximum, budget)
        else:
            result = {"items": [item for _, item in selected], "errors": [],
                      "body_text_chars": 0, "body_complete": None if mode == "metadata" else False}
        next_cursor = page["next_cursor"]
        if next_cursor:
            next_cursor = f"{BUNDLE_CURSOR}|{method}|{day}|{next_cursor}"
        page.pop("folder_id", None)
        return {**self.cache_context(index), **page, **result, "next_cursor": next_cursor,
                "folder": folder, "mode": mode, **({"date": day} if today else {})}

    def perform(self, method, params):
        return deepcopy(self._perform(method, params))

    def _perform(self, method, params):
        if method not in FIELDS or not isinstance(params, dict) or set(params) - FIELDS[method]:
            fail("INVALID_ARGUMENT", "Unknown operation or parameters")
        if method in ("context_bundle", "today_bundle"):
            return self._context_bundle(method, params)
        if method == "messages_bundle":
            return self._messages_bundle(params)
        if "unread" in params and type(params["unread"]) is not bool:
            fail("INVALID_ARGUMENT", "unread must be boolean")
        if params.get("format", "text") != "text":
            fail("UNSUPPORTED_FORMAT", "Common macOS reads currently return text only")
        person, role = params.get("person"), params.get("person_role", "any")
        if method == "search":
            if "search_body" in params and type(params["search_body"]) is not bool:
                fail("INVALID_ARGUMENT", "search_body must be boolean")
            query = params.get("query", "")
            if not isinstance(query, str) or len(query) > 500 or (not query.strip() and not person):
                fail("INVALID_ARGUMENT", "search requires query or person (1..500 characters)")
            try:
                core.person_matches({}, person, role)
            except ValueError as error:
                fail("INVALID_ARGUMENT", str(error))
        value = params.get("id", params.get("folder", "inbox"))
        kind = "message" if "id" in params else "folder"
        identity = None if value in ("inbox", "sent", "deleted", "drafts", "junk", "outbox") else decode_id(value, kind)
        directory, index = self.service._current(identity["snapshot"] if identity else None)
        groups = self.groups(index)
        context = self.cache_context(index)
        if method == "health":
            return {**context, "connected": True, "offline": True, "outlook_required": False,
                    "store_count": len(groups), "readable_store_count": len(groups),
                    "outlook_version": None, "cache_summary": index["summary"],
                    "capabilities": {"unread_filter": True, "body_search": True, "html_body": False,
                                     "received_time": True, "sent_time": True},
                    "limitations": [index["summary"]["coverage"],
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
            unread = params.get("unread")
            unread_unknown = sum(row.get("unread") is None for row in rows)
            if unread is not None:
                rows = [row for row in rows if row.get("unread") is unread]
            offset = number(params, "offset", 0, 0, 10000000)
            limit = number(params, "limit", 25, 1, 200)
            scan = number(params, "scan_limit", 500, 1, 2000)
            query_text = params.get("query", "")
            query = query_text.casefold()
            person_incomplete = False
            matched, position = [], min(offset, len(rows))
            window = rows[position:min(len(rows), position + scan)]
            candidate_rows, person_roles, person_incomplete_by_id = [], {}, {}
            for row in window:
                roles, incomplete = core.person_matches(row, person, role)
                person_incomplete_by_id[row["id"]] = incomplete
                if person and not roles:
                    continue
                candidate_rows.append(row)
                person_roles[row["id"]] = roles

            body_hits, fulltext = {}, None
            if method == "search" and params.get("search_body") and query_text:
                body_hits, fulltext = self.service._fulltext_index.search(
                    directory, index, candidate_rows, query_text)

            examined = 0
            for row in window:
                examined += 1
                person_incomplete |= person_incomplete_by_id[row["id"]]
                roles = person_roles.get(row["id"])
                if roles is None:
                    continue
                metadata_match = not query or any(
                    query in (row.get(field) or "").casefold()
                    for field in ("subject", "sender", "sender_name", "preview"))
                body_match = row["id"] in body_hits
                if method == "search" and query and not (metadata_match or body_match):
                    continue
                result = {**self.metadata(row, index),
                          **({"matched_person_roles": roles} if person else {})}
                if body_match:
                    result["body_snippet"] = body_hits[row["id"]]
                matched.append(result)
                if len(matched) >= limit:
                    break

            position += examined
            pending = bool(fulltext and (fulltext.get("pending") or fulltext.get("errors")))
            next_offset = position if position < len(rows) else None
            if pending:
                # Keep the window anchored until every candidate has been searched.
                next_offset = offset
            result = {**context, "messages": matched, "scanned": max(0, examined),
                      "filtered_item_count": len(rows), "next_offset": next_offset,
                      "complete": not pending and position >= len(rows),
                      "scope": "single_folder", "pagination": "live_offset_not_snapshot",
                      "body_search_incomplete": bool(fulltext and not fulltext.get("complete", False)),
                      "person_search_incomplete": person_incomplete,
                      "received_time_unknown": unknown_times,
                      "unread_unknown_count": unread_unknown,
                      "unread_filter_incomplete": unread is not None and unread_unknown > 0,
                      "date_filter_incomplete": bool(dated and unknown_times)}
            if fulltext is not None:
                result["fulltext"] = fulltext
                if pending:
                    result["retry_required"] = True
            return result
        if not identity or kind != "message":
            fail("INVALID_ID", "Expected a message ID")
        row = next((row for row in index["items"] if row["id"] == identity["key"]), None)
        if row is None:
            fail("SCOPE_DENIED", "Message is outside the verified cache scope")
        if method == "read":
            offset = number(params, "body_offset", 0, 0, 100000000)
            maximum = number(params, "max_chars", 20000, 1, 200000)
            return {**context, **self._read(directory, index, row, offset, maximum)}
        attachments = self.attachment_metadata(row)
        if method == "attachments":
            return {**context, "attachments": attachments, "attachments_complete": False}
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
