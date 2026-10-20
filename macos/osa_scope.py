"""Read-only Inbox scope evidence from Outlook's Osa response logs.

The cache index contains message identities, but not a trustworthy current
folder membership.  This module reads only bounded folder/header metadata from
the local Osa logs.  It never follows Outlook paths, reads message bodies, or
calls Outlook automation.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime, timezone
import gzip
import hashlib
import json
import math
import os
from pathlib import Path
import re
import stat
import zlib
import xml.etree.ElementTree as ET
from typing import Any, Iterable, Mapping


# The production profile currently has far fewer relevant files than these
# bounds.  A complete catalog is required; hitting a bound fails closed.
MAX_LOG_FILES = 4096
MAX_LOG_DIRECTORIES = 64
MAX_COMPRESSED_FILE_BYTES = 64 * 1024 * 1024
MAX_DECOMPRESSED_XML_BYTES = 128 * 1024 * 1024

_FOLDER_NAME = re.compile(r"\.(?:getfolder|getfolderlist)\.", re.I)
_MESSAGE_NAME = re.compile(
    r"\.(?:getmessage|getmessageheaderbyid|getmessageheaderinfolder)\.", re.I
)
_FORBIDDEN_XML = (b"<!doctype", b"<!entity")
_FORBIDDEN_TAIL_BYTES = 128


class _BoundError(ValueError):
    pass


class _UnsafeXML(ValueError):
    pass


class _UnstableLog(OSError):
    pass


class OsaScopeError(ValueError):
    """The Osa evidence set cannot be used safely for folder membership."""

    pass


def _local(tag: Any) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _child(node: ET.Element, name: str) -> ET.Element | None:
    for child in list(node):
        if _local(child.tag) == name:
            return child
    return None


def _text(node: ET.Element | None, name: str) -> str:
    child = _child(node, name) if node is not None else None
    return (child.text or "").strip() if child is not None else ""


def _view_folder_id(header: ET.Element) -> str:
    view = _child(header, "ViewId")
    return _text(view, "ImmId")


def _canonical(value: Any) -> str:
    return str(value or "").strip().casefold()


def _file_stat(path: Path) -> dict[str, int]:
    value = path.stat()
    return {"size": value.st_size, "mtime_ns": value.st_mtime_ns, "inode": value.st_ino}


def _same_stat(path: Path, before: Mapping[str, int]) -> bool:
    try:
        return _file_stat(path) == dict(before)
    except OSError:
        return False


def _finite_time(value: Any) -> float | None:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return None
    try:
        value = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return value if math.isfinite(value) else None


def _parse_time(value: str | None) -> float | None:
    if not isinstance(value, str) or not value.strip():
        return None
    try:
        parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        timestamp = parsed.timestamp()
        return timestamp if math.isfinite(timestamp) else None
    except (TypeError, ValueError, OverflowError, OSError):
        return None


def _item_time(item: Mapping[str, Any]) -> float | None:
    for field in ("sent_unix", "sent_time_unix"):
        value = _finite_time(item.get(field))
        if value is not None and math.isfinite(value):
            return value
    for field in ("sent_at", "sent_time"):
        value = item.get(field)
        parsed = _parse_time(value if isinstance(value, str) else None)
        if parsed is not None and math.isfinite(parsed):
            return parsed
    return None


class _GuardedReader:
    """Bound decompression and reject DTD/entity declarations before parsing."""

    def __init__(self, source: Any, maximum: int):
        self.source = source
        self.maximum = maximum
        self.read_bytes = 0
        self.tail = b""

    def read(self, size: int = -1) -> bytes:
        if size is None or size < 0:
            size = 64 * 1024
        if self.read_bytes >= self.maximum:
            raise _BoundError("decompressed Osa XML exceeds limit")
        chunk = self.source.read(min(size, self.maximum - self.read_bytes + 1))
        self.read_bytes += len(chunk)
        if self.read_bytes > self.maximum:
            raise _BoundError("decompressed Osa XML exceeds limit")
        raw_window = self.tail + chunk
        # Osa responses may be UTF-16/32.  Removing NUL bytes from the
        # inspection copy catches their ASCII declaration markers while the
        # original bytes are still returned unchanged to ElementTree.
        window = raw_window.lower().replace(b"\x00", b"")
        if any(marker in window for marker in _FORBIDDEN_XML):
            raise _UnsafeXML("DTD and external entities are not accepted")
        self.tail = raw_window[-_FORBIDDEN_TAIL_BYTES:]
        return chunk

    def readable(self) -> bool:
        return True


def _file_kind(path: Path) -> str | None:
    name = path.name
    if _FOLDER_NAME.search(name):
        return "folder"
    if _MESSAGE_NAME.search(name):
        return "message"
    return None


def _relative_name(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return path.name


def _log_files(root: Path, evidence: dict[str, Any]) -> list[tuple[str, Path]]:
    try:
        root_stat = root.lstat()
    except FileNotFoundError:
        evidence["status"] = "missing"
        evidence["missing_log_root"] = True
        return []
    except OSError as error:
        raise OsaScopeError("Osa log root is unreadable") from error
    if stat.S_ISLNK(root_stat.st_mode):
        raise OsaScopeError("Osa log root must not be a symbolic link")
    if not stat.S_ISDIR(root_stat.st_mode):
        raise OsaScopeError("Osa log root is not a directory")
    result: list[tuple[str, Path]] = []
    directories = []
    try:
        for child in sorted(root.iterdir(), key=lambda item: item.name):
            if child.name.startswith("OutlookServiceApiLogs_"):
                try:
                    child_stat = child.lstat()
                except OSError as error:
                    raise OsaScopeError("Osa log directory is unreadable") from error
                if stat.S_ISLNK(child_stat.st_mode):
                    raise OsaScopeError("Osa log directory must not be a symbolic link")
                if not stat.S_ISDIR(child_stat.st_mode):
                    raise OsaScopeError("Osa log directory is not a directory")
                directories.append(child)
    except OSError:
        raise OsaScopeError("Osa log root is unreadable")
    if len(directories) > MAX_LOG_DIRECTORIES:
        evidence["status"] = "limit_exceeded"
        evidence["directory_limit"] = MAX_LOG_DIRECTORIES
        return []
    for directory in directories:
        try:
            children = sorted(directory.iterdir(), key=lambda item: item.name)
        except OSError as error:
            raise OsaScopeError("Osa log directory is unreadable") from error
        for path in children:
            if not path.name.endswith(".res.xmlgz"):
                continue
            kind = _file_kind(path)
            if kind is None:
                continue
            try:
                path_stat = path.lstat()
            except OSError as error:
                raise OsaScopeError("Osa log file is unreadable") from error
            if stat.S_ISLNK(path_stat.st_mode):
                raise OsaScopeError("Osa log file must not be a symbolic link")
            if not stat.S_ISREG(path_stat.st_mode):
                raise OsaScopeError("Osa log file is not a regular file")
            result.append((kind, path))
            if len(result) > MAX_LOG_FILES:
                evidence["status"] = "limit_exceeded"
                evidence["file_limit"] = MAX_LOG_FILES
                return []
    return result


def _parse_log(path: Path, kind: str, root: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str]:
    before = _file_stat(path)
    if before["size"] > MAX_COMPRESSED_FILE_BYTES:
        raise _BoundError("compressed Osa XML exceeds limit")
    folder_rows: list[dict[str, Any]] = []
    message_rows: list[dict[str, Any]] = []
    log_time = ""
    stack: list[str] = []
    descriptor = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0))
    with os.fdopen(descriptor, "rb") as raw:
        with gzip.GzipFile(fileobj=raw) as compressed:
            guarded = _GuardedReader(compressed, MAX_DECOMPRESSED_XML_BYTES)
            for event, element in ET.iterparse(guarded, events=("start", "end")):
                tag = _local(element.tag)
                if event == "start":
                    stack.append(tag)
                    continue
                parent = stack[-2] if len(stack) > 1 else ""
                if tag in ("LogTime", "ReceiveTime") and len(stack) <= 2:
                    value = (element.text or "").strip()
                    if value and not log_time:
                        log_time = value
                elif kind == "folder" and tag == "View" and parent != "MessageHeader":
                    identifier = _text(element, "ImmId")
                    if identifier:
                        folder_rows.append({
                            "id": identifier,
                            "parent_id": _text(element, "ImmParentId"),
                            "display_name": _text(element, "DisplayName"),
                            "type": _text(element, "Type"),
                            "log_time": log_time,
                            "source_file": _relative_name(path, root),
                        })
                    element.clear()
                elif kind == "message" and tag == "MessageHeader":
                    message_rows.append({
                        "entry_id": _text(element, "ImmId"),
                        "message_id": _text(element, "InternetMessageId"),
                        "sent_time": _text(element, "SentTime"),
                        "sent_unix": _parse_time(_text(element, "SentTime")),
                        "folder_id": _view_folder_id(element),
                        "log_time": log_time,
                        "log_unix": _parse_time(log_time),
                        "source_file": _relative_name(path, root),
                    })
                    element.clear()
                elif tag in ("FlattenedMessage", "NewMessage"):
                    # Do not retain body/attachment payloads in the parse tree.
                    element.clear()
                else:
                    inside_view = any(value == "View" for value in stack[:-1])
                    inside_view_id = any(value == "ViewId" for value in stack[:-1])
                    inside_header = any(value == "MessageHeader" for value in stack[:-1])
                    keep_header_field = (
                        inside_header
                        and (inside_view_id or tag in {"ImmId", "InternetMessageId", "SentTime", "ViewId"})
                    )
                    # Keep only fields needed by a View or MessageHeader;
                    # clear unrelated response data, including real body data.
                    if not inside_view and not keep_header_field:
                        element.clear()
                stack.pop()
    if not _same_stat(path, before):
        raise _UnstableLog("Osa log changed during read")
    for row in folder_rows:
        if not row["log_time"]:
            row["log_time"] = log_time
    for row in message_rows:
        if not row["log_time"]:
            row["log_time"] = log_time
            row["log_unix"] = _parse_time(log_time)
    return folder_rows, message_rows, log_time


def _is_inbox(type_value: str) -> bool:
    normalized = type_value.strip().casefold()
    return normalized in {"1", "inbox (1)", "inbox"}


def _is_type(type_value: str, code: int, label: str) -> bool:
    normalized = type_value.strip().casefold()
    return normalized in {str(code), f"{label.casefold()} ({code})", label.casefold()}


def _folder_name(display_name: str, type_value: str) -> tuple[str, str]:
    if _is_inbox(type_value):
        return "收件箱", "inbox_type"
    if _is_type(type_value, 3, "SentItems"):
        return "已发送", "sent_type"
    if _is_type(type_value, 4, "DeletedItems"):
        return "已删除", "deleted_type"
    if display_name:
        if display_name.casefold().startswith("pii:"):
            return display_name, "redacted"
        return display_name, "observed"
    return "[未命名文件夹]", "unavailable"


def _latest_folder_rows(
    rows: Iterable[dict[str, Any]],
) -> tuple[dict[str, dict[str, Any]], set[str], set[str]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[_canonical(row["id"])].append(row)
    latest: dict[str, dict[str, Any]] = {}
    ambiguous: set[str] = set()
    no_valid_time: set[str] = set()
    for key, values in grouped.items():
        parsed_values = [
            (row, _parse_time(row.get("log_time")))
            for row in values
        ]
        valid_values = [
            (row, parsed)
            for row, parsed in parsed_values
            if parsed is not None and math.isfinite(parsed)
        ]
        if not valid_values:
            no_valid_time.add(key)
            continue
        maximum = max(parsed for _, parsed in valid_values)
        candidates = [row for row, parsed in valid_values if parsed == maximum]
        candidates.sort(key=lambda row: (row.get("source_file", ""), row.get("parent_id", "")))
        chosen = dict(candidates[-1])
        shapes = {
            (_canonical(row.get("parent_id")), row.get("type", ""), row.get("display_name", ""))
            for row in candidates
        }
        if len(shapes) > 1:
            ambiguous.add(key)
        latest[key] = chosen
    return latest, ambiguous, no_valid_time


def _choose_latest(rows: list[dict[str, Any]]) -> tuple[dict[str, Any] | None, bool]:
    if not rows:
        return None, False
    valid_rows = [
        (row, value)
        for row in rows
        if (value := _finite_time(row.get("log_unix"))) is not None
        and math.isfinite(value)
    ]
    if not valid_rows:
        return None, False
    maximum = max(value for _, value in valid_rows)
    candidates = [row for row, value in valid_rows if value == maximum]
    folder_ids = {_canonical(row.get("folder_id")) for row in candidates if row.get("folder_id")}
    has_missing_folder = any(not row.get("folder_id") for row in candidates)
    if len(folder_ids) > 1 or (folder_ids and has_missing_folder):
        return None, True
    candidates.sort(key=lambda row: (row.get("source_file", ""), row.get("folder_id", "")))
    return dict(candidates[-1]), False


def _item_entry_ids(item: Mapping[str, Any]) -> list[str]:
    values: Any = item.get("entry_ids", [])
    if isinstance(values, str):
        values = [values]
    elif not isinstance(values, (list, tuple, set)):
        values = []
    values = list(values)
    for field in ("entry_id", "imm_id"):
        if item.get(field):
            values.append(item[field])
    return sorted({_canonical(value) for value in values if isinstance(value, str) and value.strip()})


def _input_rows(items: Iterable[Mapping[str, Any]] | Mapping[str, Any] | None) -> list[Mapping[str, Any]]:
    if items is None:
        return []
    if isinstance(items, Mapping):
        items = items.get("items", [])
    if isinstance(items, (str, bytes)):
        return []
    try:
        return [item for item in items if isinstance(item, Mapping)]
    except TypeError:
        return []


def _descendant_keys(roots: set[str], children: Mapping[str, set[str]]) -> set[str]:
    result = set(roots)
    pending = list(roots)
    while pending:
        parent = pending.pop()
        for child in sorted(children.get(parent, ())):
            if child not in result:
                result.add(child)
                pending.append(child)
    return result


def _scope_manifest(root: Path, files: list[tuple[str, Path]]) -> list[tuple[str, str, dict[str, int]]]:
    manifest: list[tuple[str, str, dict[str, int]]] = []
    for kind, path in files:
        try:
            stat_value = _file_stat(path)
        except OSError as error:
            raise OsaScopeError("Osa log file became unreadable") from error
        manifest.append((kind, _relative_name(path, root), stat_value))
    return manifest


def _scope_fingerprint(
    profile: Path,
    root: Path,
    files: list[tuple[str, Path]],
    evidence: Mapping[str, Any],
    manifest: list[tuple[str, str, dict[str, int]]] | None = None,
) -> str:
    if manifest is None:
        manifest = _scope_manifest(root, files)
    material = {
        "profile": str(profile),
        "log_root": str(root),
        "files": manifest,
        "limits": {
            "max_log_files": MAX_LOG_FILES,
            "max_compressed_file_bytes": MAX_COMPRESSED_FILE_BYTES,
            "max_decompressed_xml_bytes": MAX_DECOMPRESSED_XML_BYTES,
        },
        "status": evidence.get("status"),
        "bad_files": evidence.get("bad_files", 0),
        "unstable_files": evidence.get("unstable_files", 0),
    }
    return hashlib.sha256(json.dumps(material, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def read_scope(
    profile: str | os.PathLike[str],
    items: Iterable[Mapping[str, Any]] | Mapping[str, Any] | None,
    *,
    log_root: str | os.PathLike[str] | None = None,
) -> dict[str, Any]:
    """Return local Inbox-folder evidence for the supplied cache-index rows.

    ``log_root`` exists only to make synthetic tests possible.  Callers use
    the two-argument form, which reads ``profile/Osa`` and never an arbitrary
    directory.  Folder membership is an observation from the latest retained
    log response, never a live Outlook assertion.
    """

    profile_path = Path(profile).expanduser().resolve()
    root = (Path(log_root).expanduser() if log_root is not None
            else profile_path / "Osa")
    evidence: dict[str, Any] = {
        "source": "Osa",
        "log_root": str(root),
        "status": "ok",
        "latest_observed": True,
        "live": False,
        "body_read": False,
        "folder_files": 0,
        "message_files": 0,
        "files_scanned": 0,
        "bad_files": 0,
        "unstable_files": 0,
        "skipped_entries": 0,
        "oversized_files": 0,
        "unsafe_xml_files": 0,
        "limit_failures": 0,
        "items": 0,
        "messages_mapped": 0,
        "messages_unknown": 0,
        "messages_ambiguous": 0,
        "messages_unmatched": 0,
        "messages_unverified_folder": 0,
        "messages_exact_identity_conflict": 0,
        "messages_exact_time_conflict": 0,
        "messages_fallback_input_duplicate": 0,
        "messages_fallback_identity_conflict": 0,
        "messages_no_valid_log_time": 0,
    }
    files = _log_files(root, evidence)
    rows = _input_rows(items)
    evidence["items"] = len(rows)
    if not files or evidence["status"] != "ok":
        if not files and evidence["status"] == "ok":
            evidence["status"] = "missing"
        fingerprint = _scope_fingerprint(profile_path, root, files, evidence)
        evidence["severe_missing"] = True
        return {"folders": [], "messages": {}, "evidence": evidence, "fingerprint": fingerprint}
    initial_manifest = _scope_manifest(root, files)
    initial_fingerprint = _scope_fingerprint(profile_path, root, files, evidence, initial_manifest)

    folder_observations: list[dict[str, Any]] = []
    message_observations: list[dict[str, Any]] = []
    for kind, path in files:
        try:
            folder_rows, message_rows, _ = _parse_log(path, kind, root)
        except _BoundError as error:
            evidence["oversized_files"] += 1
            evidence["bad_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} exceeds the safety limit") from error
        except _UnsafeXML as error:
            evidence["unsafe_xml_files"] += 1
            evidence["bad_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} contains unsafe XML") from error
        except _UnstableLog as error:
            evidence["bad_files"] += 1
            evidence["unstable_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} changed during read") from error
        except (OSError, EOFError, gzip.BadGzipFile, zlib.error, ET.ParseError, ValueError) as error:
            evidence["bad_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} is unreadable or invalid") from error
        evidence["files_scanned"] += 1
        if kind == "folder":
            evidence["folder_files"] += 1
            folder_observations.extend(folder_rows)
        else:
            evidence["message_files"] += 1
            message_observations.extend(message_rows)

    evidence["files_seen"] = len(files)
    evidence["folder_observations"] = len(folder_observations)
    evidence["message_observations"] = len(message_observations)
    end_evidence: dict[str, Any] = {"status": "ok"}
    end_files = _log_files(root, end_evidence)
    end_manifest = _scope_manifest(root, end_files)
    end_fingerprint = _scope_fingerprint(profile_path, root, end_files, evidence, end_manifest)
    if end_manifest != initial_manifest or end_fingerprint != initial_fingerprint:
        raise OsaScopeError("Osa log evidence changed during scan")
    if not folder_observations:
        evidence["status"] = "missing"
        evidence["severe_missing"] = True
        return {"folders": [], "messages": {}, "evidence": evidence,
                "fingerprint": _scope_fingerprint(profile_path, root, end_files, evidence, end_manifest)}

    folder_map, ambiguous_folders, no_valid_folder_times = _latest_folder_rows(folder_observations)
    for key in ambiguous_folders | no_valid_folder_times:
        folder_map.pop(key, None)
    evidence["ambiguous_folder_nodes"] = len(ambiguous_folders)
    evidence["no_valid_folder_time_nodes"] = len(no_valid_folder_times)
    inbox_roots = {key for key, row in folder_map.items() if _is_inbox(row.get("type", ""))}
    children: dict[str, set[str]] = defaultdict(set)
    for key, row in folder_map.items():
        parent = _canonical(row.get("parent_id"))
        if parent:
            children[parent].add(key)
    sent_roots = {key for key, row in folder_map.items()
                  if _is_type(row.get("type", ""), 3, "SentItems")}
    deleted_roots = {key for key, row in folder_map.items()
                     if _is_type(row.get("type", ""), 4, "DeletedItems")}
    inbox_subtree = _descendant_keys(inbox_roots, children)
    sent_subtree = _descendant_keys(sent_roots, children)
    deleted_subtree = _descendant_keys(deleted_roots, children)

    folders: list[dict[str, Any]] = []
    observation_counts = Counter(
        _canonical(observation.get("id")) for observation in folder_observations
    )
    public_id_by_key: dict[str, str] = {
        key: row["id"].strip() for key, row in folder_map.items()
    }
    for key, row in sorted(folder_map.items(), key=lambda pair: pair[1].get("id", "")):
        is_inbox = key in inbox_roots
        folder_type = row.get("type", "")
        name, name_status = _folder_name(row.get("display_name", ""), folder_type)
        public_id = row["id"].strip()
        parent_key = _canonical(row.get("parent_id"))
        parent_id = public_id_by_key.get(parent_key, row.get("parent_id", "").strip())
        folders.append({
            "id": public_id,
            "parent_id": parent_id,
            "name": name,
            "is_inbox": is_inbox,
            "is_sent": _is_type(folder_type, 3, "SentItems"),
            "is_deleted": _is_type(folder_type, 4, "DeletedItems"),
            "evidence": {
                "folder_type": folder_type,
                "name_status": name_status,
                "display_name": row.get("display_name", "") or None,
                "inbox_subtree": key in inbox_subtree,
                "sent_subtree": key in sent_subtree,
                "deleted_subtree": key in deleted_subtree,
                "latest_observed": True,
                "live": False,
                "log_time": row.get("log_time", ""),
                "source_file": row.get("source_file", ""),
                "observation_count": observation_counts[key],
                "catalog_status": "ambiguous_same_log_time" if key in ambiguous_folders else "latest",
            },
        })

    by_entry: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_message_id: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for observation in message_observations:
        if observation.get("entry_id"):
            by_entry[_canonical(observation["entry_id"])].append(observation)
        if observation.get("message_id"):
            by_message_id[observation["message_id"]].append(observation)

    input_mid_counts = Counter(
        item["message_id"].strip()
        for item in rows
        if isinstance(item.get("message_id"), str) and item["message_id"].strip()
    )
    messages: dict[str, dict[str, Any]] = {}
    for item in rows:
        # The caller's stable cache-row id (for example ``hx_*``) is the
        # public mapping key.  Never expose an InternetMessageId as a key.
        identifier = item.get("id")
        message_id = item.get("message_id")
        if not isinstance(identifier, str) or not identifier.strip() or not isinstance(message_id, str) or not message_id.strip():
            evidence["messages_unknown"] += 1
            evidence["messages_unmatched"] += 1
            continue
        identifier = identifier.strip()
        message_id = message_id.strip()
        exact_ids = _item_entry_ids(item)
        exact_matches: list[dict[str, Any]] = []
        matched_entry_keys: set[str] = set()
        for entry_key in exact_ids:
            if by_entry.get(entry_key):
                matched_entry_keys.add(entry_key)
                exact_matches.extend(by_entry[entry_key])
        method = "exact_entry_id"
        candidates: list[dict[str, Any]] = []
        if exact_ids:
            # A known EntryID is authoritative only after the header's own
            # InternetMessageId and SentTime agree with the index row.  Never
            # route a conflicting EntryID through the weaker MID fallback.
            if len(matched_entry_keys) > 1:
                evidence["messages_unknown"] += 1
                evidence["messages_exact_identity_conflict"] += 1
                continue
            if not matched_entry_keys:
                evidence["messages_unknown"] += 1
                evidence["messages_unmatched"] += 1
                continue
            if len(matched_entry_keys) != len(exact_ids):
                # Every supplied exact identity must be observed.  A partial
                # hit is still a conflict; it must not be rescued by MID.
                evidence["messages_unknown"] += 1
                evidence["messages_exact_identity_conflict"] += 1
                continue
            item_sent = _item_time(item)
            identity_conflict = any(
                row.get("message_id", "") != message_id for row in exact_matches
            )
            time_conflict = (
                item_sent is None
                or any(
                    (row_sent := _finite_time(row.get("sent_unix"))) is None
                    or not math.isfinite(row_sent)
                    or abs(row_sent - item_sent) >= 1.0
                    for row in exact_matches
                )
            )
            if identity_conflict:
                evidence["messages_unknown"] += 1
                evidence["messages_exact_identity_conflict"] += 1
                continue
            if time_conflict:
                evidence["messages_unknown"] += 1
                evidence["messages_exact_time_conflict"] += 1
                continue
            candidates = exact_matches
        else:
            # MID fallback is deliberately unavailable when the index already
            # has an EntryID, even if that EntryID was absent from Osa logs.
            method = "internet_message_id_sent_time"
            if input_mid_counts[message_id] != 1:
                evidence["messages_unknown"] += 1
                evidence["messages_fallback_input_duplicate"] += 1
                continue
            possible = by_message_id.get(message_id, [])
            identities = {_canonical(row.get("entry_id")) for row in possible if row.get("entry_id")}
            # Repeated observations of one Osa identity are allowed; distinct
            # identities sharing an InternetMessageId are not.  A mixture of
            # identified and unidentified observations cannot prove one
            # identity either.
            if len(identities) > 1 or (identities and any(not row.get("entry_id") for row in possible)):
                evidence["messages_unknown"] += 1
                evidence["messages_fallback_identity_conflict"] += 1
                continue
            item_sent = _item_time(item)
            candidates = [
                row for row in possible
                if item_sent is not None
                and (row_sent := _finite_time(row.get("sent_unix"))) is not None
                and math.isfinite(row_sent)
                and abs(row_sent - item_sent) < 1.0
            ]

        if not candidates:
            evidence["messages_unknown"] += 1
            evidence["messages_unmatched"] += 1
            continue
        if not any(
            (log_time := _finite_time(row.get("log_unix"))) is not None
            and math.isfinite(log_time)
            for row in candidates
        ):
            evidence["messages_unknown"] += 1
            evidence["messages_no_valid_log_time"] += 1
            continue
        chosen, ambiguous = _choose_latest(candidates)

        folder_id = chosen.get("folder_id", "") if chosen is not None else ""
        folder_key = _canonical(folder_id)
        if ambiguous:
            evidence["messages_unknown"] += 1
            evidence["messages_ambiguous"] += 1
            continue
        elif not chosen or not folder_id:
            evidence["messages_unknown"] += 1
            evidence["messages_unmatched"] += 1
            continue
        elif folder_key not in folder_map:
            evidence["messages_unknown"] += 1
            evidence["messages_unverified_folder"] += 1
            continue

        # Deleted descendants remain excluded even if malformed catalog data
        # also places them below another scoped root.
        if folder_key in deleted_subtree:
            scope = "deleted"
        elif folder_key in inbox_subtree:
            scope = "inbox"
        elif folder_key in sent_subtree:
            scope = "sent"
        else:
            scope = "outside"
        folder_status = "verified_latest_observation"
        messages[identifier] = {
            "folder_id": public_id_by_key[folder_key],
            "folder_status": folder_status,
            "evidence": {
                "match_method": method if candidates else "none",
                "scope": scope,
                "candidate_count": len(candidates),
                "observed_log_times": sorted({row.get("log_time", "") for row in candidates if row.get("log_time")}),
                "latest_log_time": chosen.get("log_time", "") if chosen else "",
                "latest_observed": bool(chosen),
                "live": False,
                "source_files": sorted({row.get("source_file", "") for row in candidates if row.get("source_file")}),
                "entry_id_keys": sorted(matched_entry_keys),
            },
        }
        evidence["messages_mapped"] += 1

    if evidence["bad_files"] or evidence["unsafe_xml_files"] or evidence["oversized_files"]:
        evidence["status"] = "partial"
    evidence["inbox_roots"] = len(inbox_roots)
    evidence["inbox_subtree_folders"] = len(inbox_subtree)
    fingerprint = _scope_fingerprint(profile_path, root, end_files, evidence, end_manifest)
    return {"folders": folders, "messages": messages, "evidence": evidence, "fingerprint": fingerprint}
