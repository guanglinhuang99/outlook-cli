"""Read-only Inbox scope evidence from Outlook's Osa response logs.

The cache index contains message identities, but not a trustworthy current
folder membership.  This module reads only bounded folder/header metadata from
the local Osa logs.  It never follows Outlook paths, reads message bodies, or
calls Outlook automation.
"""

from __future__ import annotations

from collections import Counter, OrderedDict, defaultdict
from datetime import datetime, timezone
import gzip
import hashlib
import json
import math
import os
from pathlib import Path
import re
import sqlite3
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
_INTEGER_TEXT = re.compile(r"\A\d+\Z")
_ISO_EXPLICIT_TIME = re.compile(
    r"\A(?P<date>\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})"
    r"(?:[.,](?P<fraction>\d+))?(?P<offset>Z|[+-]\d{2}:?\d{2})\Z"
)
_DOTNET_TICKS_PER_SECOND = 10_000_000
_DOTNET_EPOCH_TICKS = 621355968000000000
_DOTNET_MAX_TICKS_EXCLUSIVE = 3155378976000000000
# The metadata path converts Unix seconds to an Asia/Singapore datetime.  Keep
# the UTC receipt below 9999-12-31 16:00:00 so that conversion cannot roll into
# year 10000.
_DOTNET_MAX_SUPPORTED_TICKS_EXCLUSIVE = (
    _DOTNET_MAX_TICKS_EXCLUSIVE
    - 8 * 60 * 60 * _DOTNET_TICKS_PER_SECOND
)
_PARSER_VERSION = "osa-metadata-v1"
_CACHE_DB_VERSION = 1
_CACHE_MAX_FILES = 4096
_CACHE_MAX_BYTES = 64 * 1024 * 1024
# The rows are capped at 64 MiB.  Keep a small, explicit SQLite page/schema
# allowance so repeated refreshes cannot grow the private cache without bound.
_CACHE_DB_MAX_BYTES = _CACHE_MAX_BYTES + 4 * 1024 * 1024
_PARTICIPANT_ROLES = ("sender", "to", "cc")
_PARTICIPANT_ROLE_TAGS = {role.title(): role for role in _PARTICIPANT_ROLES}
_PARTICIPANT_FIELDS = {"DisplayName", "EmailAddress", "EmailAddressType",
                       "ServerId_Substrate"}
_REDACTED_PARTICIPANT = re.compile(
    r"\A(?:pii:|\[redacted\]|<redacted>|\(redacted\)|redacted\Z)",
    re.I,
)


class _BoundError(ValueError):
    pass


class _UnsafeXML(ValueError):
    pass


class _UnstableLog(OSError):
    pass


def _stat_fields(value: os.stat_result) -> dict[str, int]:
    """Return the source identity fields used for cache validation."""
    return {
        "dev": int(value.st_dev),
        "inode": int(value.st_ino),
        "size": int(value.st_size),
        "mtime_ns": int(value.st_mtime_ns),
        "ctime_ns": int(value.st_ctime_ns),
        "mode": int(value.st_mode),
        "nlink": int(value.st_nlink),
        "uid": int(getattr(value, "st_uid", 0)),
        "gid": int(getattr(value, "st_gid", 0)),
    }


def _readable_file_stat(path: Path) -> dict[str, int]:
    """Validate and stat a regular, non-symlink file through an open fd."""
    path_stat = path.lstat()
    if stat.S_ISLNK(path_stat.st_mode):
        raise OSError("Osa log file must not be a symbolic link")
    if not stat.S_ISREG(path_stat.st_mode):
        raise OSError("Osa log file is not a regular file")
    descriptor = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0))
    try:
        descriptor_stat = os.fstat(descriptor)
        if not stat.S_ISREG(descriptor_stat.st_mode):
            raise OSError("Osa log file is not a regular file")
        return _stat_fields(descriptor_stat)
    finally:
        os.close(descriptor)


def _private_file(path: Path, *, allow_missing: bool = True) -> bool:
    """Accept only a private, single-link regular file owned by this user."""
    try:
        value = path.lstat()
    except FileNotFoundError:
        return allow_missing
    except OSError:
        return False
    if stat.S_ISLNK(value.st_mode) or not stat.S_ISREG(value.st_mode):
        return False
    if int(getattr(value, "st_nlink", 0)) != 1:
        return False
    uid = getattr(os, "getuid", lambda: None)()
    if uid is not None and int(getattr(value, "st_uid", -1)) != int(uid):
        return False
    return not (stat.S_IMODE(value.st_mode) & 0o077)


def _private_cache_parent(path: Path) -> bool:
    """Validate the already-created state directory before SQLite opens."""
    try:
        value = path.parent.lstat()
    except OSError:
        return False
    if stat.S_ISLNK(value.st_mode) or not stat.S_ISDIR(value.st_mode):
        return False
    uid = getattr(os, "getuid", lambda: None)()
    if uid is not None and int(getattr(value, "st_uid", -1)) != int(uid):
        return False
    return not (stat.S_IMODE(value.st_mode) & 0o077)


class LogParseCache:
    """Bounded cache for successful Osa metadata parses.

    ``path`` is an optional private, state-owned SQLite file.  It is opened
    lazily by :meth:`begin` after the caller has validated the current Osa
    source and is closed again after :meth:`flush`; the default remains an
    in-process cache.  SQLite failures disable only persistence and never
    replace a source parsing error with cached data.
    """

    _TABLE = "osa_metadata_cache"
    _COLUMNS = {
        "cache_key", "parser_version", "key_json", "fingerprint_json",
        "payload", "payload_sha256", "payload_bytes",
    }

    def __init__(
        self,
        path: str | os.PathLike[str] | None = None,
        *,
        max_files: int = _CACHE_MAX_FILES,
        max_bytes: int = _CACHE_MAX_BYTES,
    ) -> None:
        self.max_files = max(0, min(int(max_files), _CACHE_MAX_FILES))
        self.max_bytes = max(0, min(int(max_bytes), _CACHE_MAX_BYTES))
        self.path = (
            Path(path).expanduser() if path is not None else None
        )
        self._entries: OrderedDict[tuple[Any, ...], tuple[dict[str, int], bytes]] = OrderedDict()
        self._serialized_bytes = 0
        self._dirty_keys: set[str] = set()
        self._deleted_keys: set[str] = set()
        self._persistent_loaded = self.path is None
        self._db_disabled = False
        self._maintenance_needed = False
        self.hits = 0
        self.misses = 0
        self.stores = 0

    @property
    def entry_count(self) -> int:
        return len(self._entries)

    @property
    def serialized_bytes(self) -> int:
        return self._serialized_bytes

    def __len__(self) -> int:
        return len(self._entries)

    @staticmethod
    def _limits() -> tuple[int, int, int]:
        return (
            MAX_LOG_FILES,
            MAX_COMPRESSED_FILE_BYTES,
            MAX_DECOMPRESSED_XML_BYTES,
        )

    def _key(
        self,
        profile: Path,
        root: Path,
        kind: str,
        path: Path,
    ) -> tuple[Any, ...]:
        return (
            _PARSER_VERSION,
            str(profile.resolve()),
            str(root.resolve()),
            self._limits(),
            self.max_files,
            self.max_bytes,
            kind,
            _relative_name(path, root),
        )

    @staticmethod
    def _key_json(key: tuple[Any, ...]) -> str:
        return json.dumps(key, ensure_ascii=False, separators=(",", ":"))

    @classmethod
    def _cache_key(cls, key_json: str) -> str:
        return hashlib.sha256(key_json.encode("utf-8")).hexdigest()

    @staticmethod
    def _fingerprint_json(source_stat: Mapping[str, int]) -> str:
        return json.dumps(dict(source_stat), sort_keys=True, separators=(",", ":"))

    @staticmethod
    def _stat_from_json(value: Any) -> dict[str, int] | None:
        if not isinstance(value, dict):
            return None
        if set(value) != {
            "dev", "inode", "size", "mtime_ns", "ctime_ns", "mode",
            "nlink", "uid", "gid",
        }:
            return None
        result: dict[str, int] = {}
        for key, item in value.items():
            if isinstance(item, bool) or not isinstance(item, int):
                return None
            result[key] = item
        return result

    @classmethod
    def _key_from_json(cls, value: Any) -> tuple[Any, ...] | None:
        if not isinstance(value, list) or len(value) != 8:
            return None
        if not isinstance(value[0], str) or value[0] != _PARSER_VERSION:
            return None
        if not isinstance(value[1], str) or not isinstance(value[2], str):
            return None
        if not isinstance(value[3], list) or len(value[3]) != 3:
            return None
        if any(isinstance(item, bool) or not isinstance(item, int) for item in value[3]):
            return None
        if any(isinstance(value[index], bool) or not isinstance(value[index], int)
               for index in (4, 5)):
            return None
        if not isinstance(value[6], str) or not isinstance(value[7], str):
            return None
        return (
            value[0], value[1], value[2], tuple(value[3]), value[4], value[5],
            value[6], value[7],
        )

    def _discard(self, key: tuple[Any, ...], *, remove_persistent: bool = True) -> None:
        entry = self._entries.pop(key, None)
        if entry is not None:
            self._serialized_bytes -= len(entry[1])
        if remove_persistent and self.path is not None and not self._db_disabled:
            key_json = self._key_json(key)
            self._deleted_keys.add(self._cache_key(key_json))

    def _serialized(self, parsed: tuple[Any, Any, Any]) -> bytes | None:
        try:
            payload = json.dumps(
                parsed,
                ensure_ascii=False,
                allow_nan=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        except (TypeError, ValueError, OverflowError):
            return None
        return payload if len(payload) <= self.max_bytes else None

    @staticmethod
    def _valid_contact(value: Any) -> bool:
        return (
            isinstance(value, dict)
            and set(value) == {"name", "address"}
            and isinstance(value["name"], str)
            and isinstance(value["address"], str)
        )

    @classmethod
    def _valid_participants(cls, value: Any) -> bool:
        if not isinstance(value, dict) or not set(value).issubset(_PARTICIPANT_ROLES):
            return False
        for contacts in value.values():
            if contacts is not None and (
                not isinstance(contacts, list)
                or not all(cls._valid_contact(contact) for contact in contacts)
            ):
                return False
        return True

    @staticmethod
    def _valid_participant_statuses(value: Any) -> bool:
        allowed = {"verified", "missing", "partial", "conflict", "redacted", "malformed"}
        return (
            isinstance(value, dict)
            and set(value).issubset(_PARTICIPANT_ROLES)
            and all(isinstance(status, str) and status in allowed for status in value.values())
        )

    @staticmethod
    def _valid_participant_counts(value: Any) -> bool:
        return (
            isinstance(value, dict)
            and set(value).issubset(_PARTICIPANT_ROLES)
            and all(
                count is None
                or (isinstance(count, int) and not isinstance(count, bool) and count >= 0)
                for count in value.values()
            )
        )

    @classmethod
    def _valid_folder_row(cls, row: dict[str, Any]) -> bool:
        fields = {"id", "parent_id", "display_name", "type", "log_time", "source_file"}
        return set(row) == fields and all(isinstance(row[field], str) for field in fields)

    @classmethod
    def _valid_message_row(cls, row: dict[str, Any]) -> bool:
        fields = {
            "entry_id", "message_id", "sent_time", "sent_unix", "received_time",
            "received_ticks", "folder_id", "log_time", "log_unix", "source_file",
            "_header_sender", "_header_sender_status", "participants",
            "participants_status",
        }
        count_fields = {"_participant_count_values", "_participant_count_statuses"}
        if set(row) not in (fields, fields | count_fields):
            return False
        for field in (
            "entry_id", "message_id", "sent_time", "received_time", "received_ticks",
            "folder_id", "log_time", "source_file",
        ):
            if not isinstance(row[field], str):
                return False
        if row["_header_sender_status"] is not None and not isinstance(
            row["_header_sender_status"], str
        ):
            return False
        for field in ("sent_unix", "log_unix"):
            value = row[field]
            if value is not None and _finite_time(value) is None:
                return False
        sender = row["_header_sender"]
        if sender is not None and not cls._valid_contact(sender):
            return False
        return (
            cls._valid_participants(row["participants"])
            and cls._valid_participant_statuses(row["participants_status"])
            and (
                not count_fields.issubset(row)
                or (
                    cls._valid_participant_counts(row["_participant_count_values"])
                    and cls._valid_participant_statuses(row["_participant_count_statuses"])
                )
            )
        )

    @classmethod
    def _decoded(cls, payload: bytes) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str] | None:
        try:
            value = json.loads(payload.decode("utf-8"))
        except (UnicodeDecodeError, TypeError, ValueError):
            return None
        if (
            not isinstance(value, list)
            or len(value) != 3
            or not isinstance(value[0], list)
            or not isinstance(value[1], list)
            or not isinstance(value[2], str)
            or not all(isinstance(row, dict) for row in value[0])
            or not all(isinstance(row, dict) for row in value[1])
            or not all(cls._valid_folder_row(row) for row in value[0])
            or not all(cls._valid_message_row(row) for row in value[1])
        ):
            return None
        return value[0], value[1], value[2]

    def _safe_paths(self) -> bool:
        if self.path is None or not _private_cache_parent(self.path):
            return False
        if not _private_file(self.path):
            return False
        try:
            if self.path.exists() and self.path.stat().st_size > _CACHE_DB_MAX_BYTES:
                return False
        except OSError:
            return False
        journal = Path(f"{self.path}-journal")
        wal = Path(f"{self.path}-wal")
        shm = Path(f"{self.path}-shm")
        # The implementation uses DELETE journaling.  A pre-existing WAL or
        # shared-memory sidecar could contain state SQLite would interpret;
        # refuse it, while allowing a private crash-recovery journal.
        if not _private_file(journal):
            return False
        try:
            if journal.exists() and journal.stat().st_size > _CACHE_DB_MAX_BYTES:
                return False
        except OSError:
            return False
        for sidecar in (wal, shm):
            try:
                sidecar.lstat()
            except FileNotFoundError:
                continue
            except OSError:
                return False
            return False
        return True

    def _disable_db(self) -> None:
        self._db_disabled = True
        self._persistent_loaded = True
        self._dirty_keys.clear()
        self._deleted_keys.clear()

    def _connect_db(self) -> sqlite3.Connection | None:
        if self.path is None or self._db_disabled or not self._safe_paths():
            self._disable_db()
            return None
        connection: sqlite3.Connection | None = None
        try:
            connection = sqlite3.connect(str(self.path), timeout=0.05)
            connection.execute("PRAGMA trusted_schema=OFF")
            connection.execute("PRAGMA foreign_keys=ON")
            connection.execute("PRAGMA busy_timeout=50")
            journal_mode = connection.execute("PRAGMA journal_mode=DELETE").fetchone()
            if not journal_mode or str(journal_mode[0]).casefold() != "delete":
                raise sqlite3.DatabaseError("unexpected SQLite journal mode")
            page_size_row = connection.execute("PRAGMA page_size").fetchone()
            page_count_row = connection.execute("PRAGMA page_count").fetchone()
            if not page_size_row or not page_count_row:
                raise sqlite3.DatabaseError("SQLite page metadata unavailable")
            page_size = int(page_size_row[0])
            page_count = int(page_count_row[0])
            if page_size <= 0 or page_count < 0:
                raise sqlite3.DatabaseError("invalid SQLite page metadata")
            max_pages = max(1, _CACHE_DB_MAX_BYTES // page_size)
            configured_max_pages = connection.execute(
                f"PRAGMA max_page_count={max_pages}"
            ).fetchone()
            if (
                not configured_max_pages
                or int(configured_max_pages[0]) > max_pages
                or page_count > max_pages
            ):
                raise sqlite3.DatabaseError("SQLite cache exceeds page limit")
            quick_check = connection.execute("PRAGMA quick_check").fetchone()
            if quick_check != ("ok",):
                raise sqlite3.DatabaseError("SQLite quick_check failed")
            table = connection.execute(
                f"SELECT name FROM sqlite_master WHERE type='table' AND name=?",
                (self._TABLE,),
            ).fetchone()
            if table is None:
                connection.execute("PRAGMA auto_vacuum=INCREMENTAL")
                connection.execute(
                    f"CREATE TABLE {self._TABLE} ("
                    "cache_key TEXT PRIMARY KEY,"
                    "parser_version TEXT NOT NULL,"
                    "key_json TEXT NOT NULL,"
                    "fingerprint_json TEXT NOT NULL,"
                    "payload TEXT NOT NULL,"
                    "payload_sha256 TEXT NOT NULL,"
                    "payload_bytes INTEGER NOT NULL"
                    ")"
                )
                connection.execute(f"PRAGMA user_version={_CACHE_DB_VERSION}")
            else:
                columns = {
                    str(row[1])
                    for row in connection.execute(f"PRAGMA table_info({self._TABLE})")
                }
                if columns != self._COLUMNS:
                    raise sqlite3.DatabaseError("SQLite cache schema mismatch")
                version = connection.execute("PRAGMA user_version").fetchone()
                if version and int(version[0]) not in (0, _CACHE_DB_VERSION):
                    raise sqlite3.DatabaseError("SQLite cache version mismatch")
                if version == (0,):
                    connection.execute(f"PRAGMA user_version={_CACHE_DB_VERSION}")
            os.chmod(self.path, 0o600)
            if not _private_file(self.path, allow_missing=False):
                raise sqlite3.DatabaseError("SQLite cache file is not private")
            connection.commit()
            return connection
        except (OSError, sqlite3.Error, ValueError, TypeError):
            if connection is not None:
                try:
                    connection.close()
                except sqlite3.Error:
                    pass
            self._disable_db()
            return None

    def begin(self) -> None:
        """Load the private SQLite cache once for the current scope scan."""
        if self.path is None or self._persistent_loaded or self._db_disabled:
            return
        connection = self._connect_db()
        if connection is None:
            return
        try:
            cursor = connection.execute(
                f"SELECT cache_key, parser_version, key_json, fingerprint_json, "
                f"payload_sha256, payload_bytes, "
                f"substr(CAST(payload AS BLOB), 1, ?) "
                f"FROM {self._TABLE} ORDER BY rowid DESC LIMIT ?",
                (self.max_bytes + 1, self.max_files + 1),
            )
            for row in cursor:
                self._load_row(row)
            row_count = connection.execute(
                f"SELECT COUNT(*), COALESCE(SUM(length(CAST(payload AS BLOB))), 0) "
                f"FROM {self._TABLE}"
            ).fetchone()
            if row_count and (
                int(row_count[0]) > self.max_files
                or int(row_count[1]) > self.max_bytes
            ):
                self._maintenance_needed = True
        except (sqlite3.Error, OSError, ValueError, TypeError):
            self._disable_db()
        finally:
            try:
                connection.close()
            except sqlite3.Error:
                self._disable_db()
            self._persistent_loaded = True

    def _load_row(self, row: tuple[Any, ...]) -> None:
        if len(row) != 7:
            self._maintenance_needed = True
            return
        cache_key, parser_version, key_json, fingerprint_json, payload_sha256, payload_bytes, payload = row
        if (
            not isinstance(cache_key, str)
            or not isinstance(parser_version, str)
            or parser_version != _PARSER_VERSION
            or not isinstance(key_json, str)
            or not isinstance(fingerprint_json, str)
            or not isinstance(payload_sha256, str)
            or not isinstance(payload_bytes, int)
            or isinstance(payload_bytes, bool)
            or not isinstance(payload, (str, bytes))
        ):
            self._maintenance_needed = True
            return
        key_value: Any
        fingerprint_value: Any
        try:
            key_value = json.loads(key_json)
            fingerprint_value = json.loads(fingerprint_json)
        except (TypeError, ValueError, json.JSONDecodeError):
            self._maintenance_needed = True
            return
        key = self._key_from_json(key_value)
        source_stat = self._stat_from_json(fingerprint_value)
        if key is None or source_stat is None:
            self._maintenance_needed = True
            return
        if self._cache_key(key_json) != cache_key:
            self._maintenance_needed = True
            return
        raw_payload = payload if isinstance(payload, bytes) else payload.encode("utf-8")
        if payload_bytes != len(raw_payload) or payload_bytes > self.max_bytes:
            self._maintenance_needed = True
            return
        if hashlib.sha256(raw_payload).hexdigest() != payload_sha256:
            self._maintenance_needed = True
            return
        self._entries[key] = (source_stat, raw_payload)
        self._serialized_bytes += len(raw_payload)
        self._trim_memory()

    def _trim_memory(self) -> None:
        while (
            len(self._entries) > self.max_files
            or self._serialized_bytes > self.max_bytes
        ):
            key, entry = self._entries.popitem(last=False)
            self._serialized_bytes -= len(entry[1])
            if self.path is not None and not self._db_disabled:
                cache_key = self._cache_key(self._key_json(key))
                # A modified row which is evicted before ``flush`` must not
                # leave its old persistent value behind.  An eviction of an
                # unchanged row keeps the durable copy available for a later
                # process.
                if cache_key in self._dirty_keys:
                    self._dirty_keys.discard(cache_key)
                    self._deleted_keys.add(cache_key)

    def flush(self) -> None:
        """Persist only entries from a successful, stable scope scan."""
        if self.path is None or self._db_disabled:
            return
        if not self._dirty_keys and not self._deleted_keys and not self._maintenance_needed:
            return
        connection = self._connect_db()
        if connection is None:
            return
        try:
            connection.execute("BEGIN IMMEDIATE")
            for cache_key in sorted(self._deleted_keys):
                connection.execute(
                    f"DELETE FROM {self._TABLE} WHERE cache_key=?", (cache_key,)
                )
            for key, (source_stat, payload) in self._entries.items():
                key_json = self._key_json(key)
                cache_key = self._cache_key(key_json)
                if cache_key not in self._dirty_keys:
                    continue
                connection.execute(
                    f"INSERT OR REPLACE INTO {self._TABLE} "
                    "(cache_key, parser_version, key_json, fingerprint_json, "
                    "payload, payload_sha256, payload_bytes) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (
                        cache_key,
                        _PARSER_VERSION,
                        key_json,
                        self._fingerprint_json(source_stat),
                        payload.decode("utf-8"),
                        hashlib.sha256(payload).hexdigest(),
                        len(payload),
                    ),
                )
            # Enforce both quotas against actual serialized SQLite values,
            # including rows left by a previous process.
            connection.execute(
                f"DELETE FROM {self._TABLE} WHERE rowid NOT IN "
                f"(SELECT rowid FROM {self._TABLE} ORDER BY rowid DESC LIMIT ?)",
                (self.max_files,),
            )
            total = 0
            for rowid, payload_bytes in connection.execute(
                f"SELECT rowid, length(CAST(payload AS BLOB)) "
                f"FROM {self._TABLE} ORDER BY rowid DESC"
            ):
                if (
                    not isinstance(payload_bytes, int)
                    or payload_bytes > self.max_bytes
                    or total + payload_bytes > self.max_bytes
                ):
                    connection.execute(
                        f"DELETE FROM {self._TABLE} WHERE rowid=?", (rowid,)
                    )
                else:
                    total += payload_bytes
            if not _private_file(self.path, allow_missing=False):
                raise sqlite3.DatabaseError("SQLite cache file is not private")
            if not self._safe_paths():
                raise sqlite3.DatabaseError("SQLite cache sidecar is unsafe")
            connection.commit()
            try:
                db_size = self.path.stat().st_size
            except OSError as error:
                raise sqlite3.DatabaseError("SQLite cache file cannot be stat'ed") from error
            if db_size > _CACHE_DB_MAX_BYTES:
                connection.execute("VACUUM")
                connection.commit()
                if self.path.stat().st_size > _CACHE_DB_MAX_BYTES:
                    raise sqlite3.DatabaseError("SQLite cache file exceeds size limit")
            self._dirty_keys.clear()
            self._deleted_keys.clear()
            self._maintenance_needed = False
        except (OSError, sqlite3.Error, ValueError, TypeError):
            try:
                connection.rollback()
            except sqlite3.Error:
                pass
            self._disable_db()
        finally:
            try:
                connection.close()
            except sqlite3.Error:
                self._disable_db()

    def get(
        self,
        profile: Path,
        root: Path,
        kind: str,
        path: Path,
        source_stat: Mapping[str, int],
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str] | None:
        self.begin()
        key = self._key(profile, root, kind, path)
        entry = self._entries.get(key)
        if entry is None:
            self.misses += 1
            return None
        expected = dict(source_stat)
        if entry[0] != expected:
            self._discard(key)
            self.misses += 1
            return None
        try:
            current = _readable_file_stat(path)
        except OSError:
            self._discard(key)
            self.misses += 1
            raise
        if current != expected or current != entry[0]:
            self._discard(key)
            self.misses += 1
            return None
        decoded = self._decoded(entry[1])
        if decoded is None:
            self._discard(key)
            self.misses += 1
            return None
        self._entries.move_to_end(key)
        self.hits += 1
        return decoded

    def invalidate(self, profile: Path, root: Path, kind: str, path: Path) -> None:
        self._discard(self._key(profile, root, kind, path))

    def put(
        self,
        profile: Path,
        root: Path,
        kind: str,
        path: Path,
        source_stat: Mapping[str, int],
        parsed: tuple[Any, Any, Any],
    ) -> bool:
        key = self._key(profile, root, kind, path)
        payload = self._serialized(parsed)
        self._discard(key)
        if self.max_files == 0 or payload is None:
            return False
        current = _readable_file_stat(path)
        if current != dict(source_stat):
            raise _UnstableLog("Osa log changed during read")
        self._entries[key] = (dict(source_stat), payload)
        self._serialized_bytes += len(payload)
        if self.path is not None and not self._db_disabled:
            cache_key = self._cache_key(self._key_json(key))
            self._dirty_keys.add(cache_key)
            self._deleted_keys.discard(cache_key)
        self.stores += 1
        self._trim_memory()
        return key in self._entries


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


def _redacted_participant_text(value: Any) -> tuple[str | None, str | None]:
    """Return non-PII participant text, keeping redaction explicit."""
    text = value.strip() if isinstance(value, str) else ""
    if not text:
        return None, "missing"
    if _REDACTED_PARTICIPANT.match(text):
        return None, "redacted"
    if re.fullmatch(r"[x*#…]{2,}", text):
        return None, "redacted"
    return text, None


def _participant_contact(name: Any, address: Any) -> tuple[dict[str, str] | None, str | None]:
    """Build one typed contact; never expose a redacted or partial pair."""
    clean_name, name_status = _redacted_participant_text(name)
    clean_address, address_status = _redacted_participant_text(address)
    if name_status == "redacted" or address_status == "redacted":
        return None, "redacted"
    if name_status or address_status:
        return None, "malformed"
    return {"name": clean_name, "address": clean_address}, None


def _new_data_context() -> dict[str, Any]:
    return {
        "imm_id": "",
        "is_partial": "",
        "is_truncated": "",
        "roles": {role: [] for role in _PARTICIPANT_ROLES},
        "role_count": {role: 0 for role in _PARTICIPANT_ROLES},
        "role_seen": {role: False for role in _PARTICIPANT_ROLES},
        "role_invalid": {role: None for role in _PARTICIPANT_ROLES},
        "role_malformed": {role: False for role in _PARTICIPANT_ROLES},
    }


def _record_role_invalid(data: dict[str, Any], role: str, status: str | None) -> None:
    if not status:
        return
    if status == "malformed":
        data["role_malformed"][role] = True
    existing = data["role_invalid"].get(role)
    if status == "redacted" or existing == "redacted":
        data["role_invalid"][role] = "redacted"
    elif status == "malformed" or existing == "malformed":
        data["role_invalid"][role] = "malformed"
    else:
        data["role_invalid"][role] = status


def _false_flag(value: Any) -> bool:
    return isinstance(value, str) and value.strip().casefold() in {"0", "false", "no"}


def _finalize_data_context(
    data: Mapping[str, Any],
) -> tuple[dict[str, list[dict[str, str]] | None], dict[str, str], bool]:
    """Freeze only complete typed roles from a MessageData element."""
    roles: dict[str, list[dict[str, str]] | None] = {}
    statuses: dict[str, str] = {}
    complete = _false_flag(data.get("is_partial")) and _false_flag(data.get("is_truncated"))
    for role in _PARTICIPANT_ROLES:
        invalid = data.get("role_invalid", {}).get(role)
        if not complete:
            roles[role] = None
            statuses[role] = "partial"
        elif invalid:
            roles[role] = None
            statuses[role] = str(invalid)
        elif data.get("role_seen", {}).get(role):
            roles[role] = list(data.get("roles", {}).get(role, ()))
            statuses[role] = "verified"
        else:
            # An explicit zero participant count is known-empty only when the
            # MessageData itself is complete and untruncated.
            roles[role] = []
            statuses[role] = "verified"
    return roles, statuses, complete


def _finalize_count_context(
    data: Mapping[str, Any],
) -> tuple[dict[str, int | None], dict[str, str]]:
    """Freeze direct role-element counts only for complete MessageData."""
    complete = _false_flag(data.get("is_partial")) and _false_flag(data.get("is_truncated"))
    counts: dict[str, int | None] = {}
    statuses: dict[str, str] = {}
    for role in _PARTICIPANT_ROLES:
        if not complete:
            counts[role] = None
            statuses[role] = "partial"
        else:
            counts[role] = int(data.get("role_count", {}).get(role, 0))
            statuses[role] = (
                "malformed" if data.get("role_malformed", {}).get(role)
                else "verified"
            )
    return counts, statuses


def _latest_participants(
    candidates: list[Mapping[str, Any]],
) -> tuple[dict[str, list[dict[str, str]] | None], dict[str, str]]:
    """Use only the sole latest observation; never backfill an older one."""
    if not candidates:
        return {}, {role: "missing" for role in _PARTICIPANT_ROLES}
    valid = [
        (row, value)
        for row in candidates
        if (value := _finite_time(row.get("log_unix"))) is not None
        and math.isfinite(value)
    ]
    if not valid:
        return {}, {role: "missing" for role in _PARTICIPANT_ROLES}
    maximum = max(value for _, value in valid)
    latest = [row for row, value in valid if value == maximum]

    def role_value(row: Mapping[str, Any], role: str) -> tuple[Any, str, bool]:
        participants = row.get("participants")
        if not isinstance(participants, Mapping):
            return None, "missing", False
        statuses = row.get("participants_status")
        if not isinstance(statuses, Mapping):
            status = "verified" if role in participants else "missing"
        else:
            status = str(statuses.get(role, "missing"))
        value = participants.get(role)
        if not (isinstance(value, list) or value is None):
            return None, "malformed", True
        return value, status, role in participants or status not in {"missing", "redacted"}

    values: dict[str, list[dict[str, str]] | None] = {}
    result_statuses: dict[str, str] = {}
    for role in _PARTICIPANT_ROLES:
        snapshots = [role_value(row, role) for row in latest]
        first_snapshot = snapshots[0]
        if any(snapshot != first_snapshot for snapshot in snapshots[1:]):
            values[role] = None
            result_statuses[role] = "conflict"
            continue
        value, status, present = snapshots[0]
        if present:
            values[role] = value
        result_statuses[role] = status
    return values, result_statuses


def _latest_participant_counts(
    candidates: list[Mapping[str, Any]],
) -> dict[str, dict[str, int | None | str]]:
    """Use only the latest count observation, without older-value backfill."""
    missing = {
        role: {"count": None, "status": "missing"}
        for role in _PARTICIPANT_ROLES
    }
    if not candidates:
        return missing
    valid = [
        (row, value)
        for row in candidates
        if (value := _finite_time(row.get("log_unix"))) is not None
        and math.isfinite(value)
    ]
    if not valid:
        return missing
    maximum = max(value for _, value in valid)
    latest = [row for row, value in valid if value == maximum]

    def role_value(row: Mapping[str, Any], role: str) -> tuple[Any, str, bool]:
        counts = row.get("_participant_count_values")
        if not isinstance(counts, Mapping):
            return None, "missing", False
        statuses = row.get("_participant_count_statuses")
        if not isinstance(statuses, Mapping):
            status = "verified" if role in counts else "missing"
        else:
            status = str(statuses.get(role, "missing"))
        count = counts.get(role)
        if not (isinstance(count, int) and not isinstance(count, bool) and count >= 0) and count is not None:
            return None, "malformed", True
        return count, status, role in counts or status != "missing"

    result: dict[str, dict[str, int | None | str]] = {}
    for role in _PARTICIPANT_ROLES:
        snapshots = [role_value(row, role) for row in latest]
        first_snapshot = snapshots[0]
        if any(snapshot != first_snapshot for snapshot in snapshots[1:]):
            result[role] = {"count": None, "status": "conflict"}
            continue
        count, status, present = first_snapshot
        result[role] = {
            "count": count if present else None,
            "status": status,
        }
    return result


def _bound_participants(
    header: Mapping[str, Any],
    data: Mapping[str, Any] | None,
) -> tuple[dict[str, list[dict[str, str]] | None], dict[str, str]]:
    """Bind MessageData contacts to its sibling header by the exact ImmId."""
    header_contact = header.get("_header_sender")
    if data is None:
        values: dict[str, list[dict[str, str]] | None] = {"to": None, "cc": None}
        statuses = {role: "missing" for role in _PARTICIPANT_ROLES}
        if isinstance(header_contact, Mapping):
            values["sender"] = [dict(header_contact)]
            statuses["sender"] = "verified"
        return values, statuses

    header_id = _canonical(header.get("entry_id"))
    data_id = _canonical(data.get("imm_id"))
    if not header_id or not data_id or header_id != data_id:
        return ({role: None for role in _PARTICIPANT_ROLES},
                {role: "conflict" for role in _PARTICIPANT_ROLES})

    roles = data.get("roles", {})
    statuses = data.get("statuses", {})
    values: dict[str, list[dict[str, str]] | None] = {}
    result_statuses: dict[str, str] = {}
    for role in _PARTICIPANT_ROLES:
        status = str(statuses.get(role, "missing"))
        value = roles.get(role)
        if isinstance(value, list) or value is None:
            # A typed From in MessageHeader is a valid sender source when the
            # corresponding MessageData sender is absent or redacted.  Keep an
            # explicit incomplete/conflict result authoritative.
            if role == "sender" and status in {"missing", "redacted"} and isinstance(header_contact, Mapping):
                values[role] = [dict(header_contact)]
                result_statuses[role] = "verified"
            elif role == "sender" and status in {"missing", "redacted"}:
                # Omit these two cases so Hx's independently verified sender
                # remains available to the common person matcher.
                result_statuses[role] = status
            else:
                values[role] = value
                result_statuses[role] = status
        else:
            values[role] = None
            result_statuses[role] = "malformed"
    return values, result_statuses


def _bound_participant_counts(
    header: Mapping[str, Any],
    data: Mapping[str, Any] | None,
) -> dict[str, dict[str, int | None | str]]:
    """Bind only exact, unique Header/Data count evidence."""
    missing = {
        role: {"count": None, "status": "missing"}
        for role in _PARTICIPANT_ROLES
    }
    if data is None:
        return missing
    header_id = _canonical(header.get("entry_id"))
    data_id = _canonical(data.get("imm_id"))
    if not header_id or not data_id or header_id != data_id:
        return {
            role: {"count": None, "status": "conflict"}
            for role in _PARTICIPANT_ROLES
        }
    counts = data.get("counts")
    statuses = data.get("count_statuses")
    if not isinstance(counts, Mapping) or not isinstance(statuses, Mapping):
        return {
            role: {"count": None, "status": "malformed"}
            for role in _PARTICIPANT_ROLES
        }
    result: dict[str, dict[str, int | None | str]] = {}
    for role in _PARTICIPANT_ROLES:
        count = counts.get(role)
        status = str(statuses.get(role, "missing"))
        if count is not None and not (
            isinstance(count, int) and not isinstance(count, bool) and count >= 0
        ):
            result[role] = {"count": None, "status": "malformed"}
        elif status not in {"verified", "missing", "partial", "conflict", "malformed"}:
            result[role] = {"count": None, "status": "malformed"}
        else:
            result[role] = {"count": count, "status": status}
    return result


def _canonical(value: Any) -> str:
    return str(value or "").strip().casefold()


def _file_stat(path: Path) -> dict[str, int]:
    value = path.lstat()
    if stat.S_ISLNK(value.st_mode):
        raise OSError("Osa log file must not be a symbolic link")
    if not stat.S_ISREG(value.st_mode):
        raise OSError("Osa log file is not a regular file")
    return _stat_fields(value)


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


def _parse_dotnet_ticks(value: Any) -> int | None:
    """Parse a nonnegative .NET DateTime tick value without float conversion."""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        ticks = value
    elif isinstance(value, str) and _INTEGER_TEXT.fullmatch(value.strip()):
        try:
            ticks = int(value.strip())
        except (TypeError, ValueError, OverflowError):
            return None
    else:
        return None
    if not 0 <= ticks < _DOTNET_MAX_SUPPORTED_TICKS_EXCLUSIVE:
        return None
    return ticks


def _parse_explicit_time_ticks(value: Any) -> int | None:
    """Parse an ISO timestamp with an explicit offset into .NET ticks."""
    if not isinstance(value, str) or not value.strip():
        return None
    match = _ISO_EXPLICIT_TIME.fullmatch(value.strip())
    if match is None:
        return None
    fraction = match.group("fraction") or ""
    if len(fraction) > 7:
        return None
    microseconds = fraction[:6].ljust(6, "0")
    offset = match.group("offset")
    if offset == "Z":
        offset = "+00:00"
    elif len(offset) == 5:
        offset = f"{offset[:3]}:{offset[3:]}"
    try:
        parsed = datetime.fromisoformat(
            f"{match.group('date')}.{microseconds}{offset}"
        )
        if parsed.tzinfo is None or parsed.utcoffset() is None:
            return None
        parsed = parsed.astimezone(timezone.utc)
        epoch = datetime(1, 1, 1, tzinfo=timezone.utc)
        delta = parsed - epoch
        ticks = (
            delta.days * 86_400 * _DOTNET_TICKS_PER_SECOND
            + delta.seconds * _DOTNET_TICKS_PER_SECOND
            + delta.microseconds * 10
            + (int(fraction[6]) if len(fraction) == 7 else 0)
        )
    except (TypeError, ValueError, OverflowError, OSError):
        return None
    return (
        ticks
        if 0 <= ticks < _DOTNET_MAX_SUPPORTED_TICKS_EXCLUSIVE
        else None
    )


def _received_time_evidence(rows: Iterable[Mapping[str, Any]]) -> dict[str, Any]:
    """Validate direct delivery fields across all identity-matched observations."""
    values: list[int] = []
    has_nonempty = False
    has_invalid = False
    has_pair = False
    has_mismatch = False
    for row in rows:
        iso_value = row.get("received_time")
        raw_value = row.get("received_ticks")
        iso_present = isinstance(iso_value, str) and bool(iso_value.strip())
        raw_present = isinstance(raw_value, str) and bool(raw_value.strip())
        if not iso_present and not raw_present:
            continue
        has_nonempty = True
        iso_ticks = _parse_explicit_time_ticks(iso_value) if iso_present else None
        raw_ticks = _parse_dotnet_ticks(raw_value) if raw_present else None
        if iso_present and iso_ticks is None:
            has_invalid = True
        if raw_present and raw_ticks is None:
            has_invalid = True
        if iso_present and raw_present:
            if iso_ticks is None or raw_ticks is None:
                continue
            has_pair = True
            if iso_ticks != raw_ticks:
                has_mismatch = True
                continue
        values.extend(value for value in (iso_ticks, raw_ticks) if value is not None)

    if has_invalid:
        return {"received_unix": None, "received_time_status": "invalid",
                "received_time_method": "osa_last_delivery_or_renew_invalid"}
    if has_mismatch:
        return {"received_unix": None, "received_time_status": "conflict",
                "received_time_method": "osa_last_delivery_time_and_received_or_renew_time_mismatch"}
    if values and len(set(values)) > 1:
        return {"received_unix": None, "received_time_status": "conflict",
                "received_time_method": "osa_receipt_candidates_conflict"}
    if has_pair and values:
        received = (values[0] - _DOTNET_EPOCH_TICKS) / _DOTNET_TICKS_PER_SECOND
        # Public metadata rounds to milliseconds; float conversion must not
        # carry a near-maximum timestamp into year 10000 in the display zone.
        upper_ms = (_DOTNET_MAX_SUPPORTED_TICKS_EXCLUSIVE - _DOTNET_EPOCH_TICKS) // 10_000
        if round(received * 1000) >= upper_ms:
            return {"received_unix": None, "received_time_status": "invalid",
                    "received_time_method": "osa_receipt_out_of_display_range"}
        return {
            "received_unix": received,
            "received_time_status": "verified",
            "received_time_method": "osa_last_delivery_time",
        }
    if has_nonempty:
        return {"received_unix": None, "received_time_status": "incomplete",
                "received_time_method": "osa_receipt_pair_incomplete"}
    return {"received_unix": None, "received_time_status": "missing",
            "received_time_method": "osa_receipt_fields_missing"}


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
    message_stack: list[dict[str, Any]] = []
    header_stack: list[dict[str, Any]] = []
    data_stack: list[dict[str, Any]] = []
    role_stack: list[dict[str, Any]] = []
    descriptor = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0))
    with os.fdopen(descriptor, "rb") as raw:
        with gzip.GzipFile(fileobj=raw) as compressed:
            guarded = _GuardedReader(compressed, MAX_DECOMPRESSED_XML_BYTES)
            for event, element in ET.iterparse(guarded, events=("start", "end")):
                tag = _local(element.tag)
                if event == "start":
                    stack.append(tag)
                    parent = stack[-2] if len(stack) > 1 else ""
                    if tag in ("FlattenedMessage", "NewMessage"):
                        message_stack.append({"headers": [], "data": []})
                    elif tag == "MessageHeader":
                        header_stack.append({
                            "from_name": "",
                            "from_address": "",
                            "from_email": None,
                            "from_contact": None,
                            "from_status": "missing",
                            "row": None,
                        })
                    elif tag == "MessageData":
                        data_stack.append(_new_data_context())
                    elif tag == "EmailAddress" and parent == "From" and header_stack:
                        header_stack[-1]["from_email"] = True
                    elif tag in _PARTICIPANT_ROLE_TAGS and parent == "MessageData":
                        role_stack.append({
                            "role": _PARTICIPANT_ROLE_TAGS[tag],
                            "name": "",
                            "address": "",
                            "address_nested": False,
                        })
                    continue
                parent = stack[-2] if len(stack) > 1 else ""
                current_message = message_stack[-1] if message_stack else None
                current_header = header_stack[-1] if header_stack else None
                current_data = data_stack[-1] if data_stack else None
                current_role = role_stack[-1] if role_stack else None
                handled = False

                if (current_header and current_header.get("from_email") is not None
                        and parent == "EmailAddress" and tag in {"Name", "Address"}):
                    value = (element.text or "").strip()
                    if tag == "Name":
                        current_header["from_name"] = value
                    else:
                        current_header["from_address"] = value
                    element.clear()
                    handled = True
                elif (current_header and tag == "EmailAddress" and parent == "From"
                      and current_header.get("from_email") is not None):
                    contact, status = _participant_contact(
                        current_header.get("from_name"), current_header.get("from_address")
                    )
                    current_header["from_contact"] = contact
                    current_header["from_status"] = status or "verified"
                    current_header["from_email"] = None
                    element.clear()
                    handled = True
                elif (current_role and parent == current_role.get("role", "").title()
                      and tag in _PARTICIPANT_FIELDS):
                    value = (element.text or "").strip()
                    if tag == "DisplayName":
                        current_role["name"] = value
                    elif tag == "EmailAddress":
                        current_role["address"] = value
                        if list(element):
                            current_role["address_nested"] = True
                    element.clear()
                    handled = True
                elif (current_data and parent == "MessageData"
                      and tag in {"ImmId", "IsPartial", "IsTruncated"}):
                    current_data[{"ImmId": "imm_id", "IsPartial": "is_partial",
                                  "IsTruncated": "is_truncated"}[tag]] = (element.text or "").strip()
                    element.clear()
                    handled = True
                elif current_role and tag == current_role.get("role", "").title() and parent == "MessageData":
                    role = current_role["role"]
                    current_data["role_count"][role] += 1
                    if current_role.get("address_nested"):
                        _record_role_invalid(current_data, role, "malformed")
                    else:
                        contact, status = _participant_contact(
                            current_role.get("name"), current_role.get("address")
                        )
                        _record_role_invalid(current_data, role, status)
                        if contact is not None:
                            current_data["roles"][role].append(contact)
                    current_data["role_seen"][role] = True
                    role_stack.pop()
                    element.clear()
                    handled = True
                elif tag == "MessageHeader" and current_header is not None:
                    row = {
                        "entry_id": _text(element, "ImmId"),
                        "message_id": _text(element, "InternetMessageId"),
                        "sent_time": _text(element, "SentTime"),
                        "sent_unix": _parse_time(_text(element, "SentTime")),
                        "received_time": _text(element, "LastDeliveryTime"),
                        "received_ticks": _text(element, "ReceivedOrRenewTime"),
                        "folder_id": _view_folder_id(element),
                        "log_time": log_time,
                        "log_unix": _parse_time(log_time),
                        "source_file": _relative_name(path, root),
                        "_header_sender": current_header.get("from_contact"),
                        "_header_sender_status": current_header.get("from_status", "missing"),
                    }
                    current_header["row"] = row
                    if current_message is not None:
                        current_message["headers"].append(row)
                    else:
                        participants, statuses = _bound_participants(row, None)
                        row["participants"] = participants
                        row["participants_status"] = statuses
                        message_rows.append(row)
                    header_stack.pop()
                    element.clear()
                    handled = True
                elif tag == "MessageData" and current_data is not None:
                    roles, statuses, complete = _finalize_data_context(current_data)
                    counts, count_statuses = _finalize_count_context(current_data)
                    data = {
                        "imm_id": current_data.get("imm_id", ""),
                        "roles": roles,
                        "statuses": statuses,
                        "counts": counts,
                        "count_statuses": count_statuses,
                        "complete": complete,
                    }
                    if current_message is not None:
                        current_message["data"].append(data)
                    data_stack.pop()
                    element.clear()
                    handled = True
                elif tag in ("FlattenedMessage", "NewMessage") and current_message is not None:
                    headers = current_message.get("headers", [])
                    data_rows = current_message.get("data", [])
                    if len(headers) == 1 and len(data_rows) == 1:
                        participant_source = data_rows[0]
                    elif len(headers) == 1 and not data_rows:
                        participant_source = None
                    else:
                        participant_source = False
                    for header in headers:
                        if participant_source is False:
                            participants = {role: None for role in _PARTICIPANT_ROLES}
                            statuses = {role: "conflict" for role in _PARTICIPANT_ROLES}
                            participant_counts = {
                                role: {"count": None, "status": "conflict"}
                                for role in _PARTICIPANT_ROLES
                            }
                        else:
                            participants, statuses = _bound_participants(
                                header, participant_source
                            )
                            participant_counts = _bound_participant_counts(
                                header, participant_source
                            )
                        row = dict(header)
                        row["participants"] = participants
                        row["participants_status"] = statuses
                        row["_participant_count_values"] = {
                            role: value["count"] for role, value in participant_counts.items()
                        }
                        row["_participant_count_statuses"] = {
                            role: value["status"] for role, value in participant_counts.items()
                        }
                        message_rows.append(row)
                    element.clear()
                    message_stack.pop()
                    handled = True
                elif tag in ("LogTime", "ReceiveTime") and len(stack) <= 2:
                    value = (element.text or "").strip()
                    if value and not log_time:
                        log_time = value
                    handled = True
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
                if not handled and tag in ("FlattenedMessage", "NewMessage"):
                    # Do not retain body/attachment payloads in the parse tree.
                    element.clear()
                    handled = True
                elif not handled:
                    inside_view = any(value == "View" for value in stack[:-1])
                    inside_view_id = any(value == "ViewId" for value in stack[:-1])
                    inside_header = any(value == "MessageHeader" for value in stack[:-1])
                    keep_header_field = (
                        inside_header
                        and (inside_view_id or tag in {
                            "ImmId", "InternetMessageId", "SentTime", "LastDeliveryTime",
                            "ReceivedOrRenewTime", "ViewId",
                        })
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
    parse_cache: LogParseCache | None = None,
) -> dict[str, Any]:
    """Return local Inbox-folder evidence for the supplied cache-index rows.

    ``log_root`` exists only to make synthetic tests possible.  Callers use
    the two-argument form, which reads ``profile/Osa`` and never an arbitrary
    directory.  ``parse_cache`` is an optional process-local metadata cache;
    folder membership is still recomputed from the complete current log set.
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
        "messages_received_verified": 0,
        "messages_received_missing": 0,
        "messages_received_incomplete": 0,
        "messages_received_invalid": 0,
        "messages_received_conflict": 0,
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
    initial_stats = {
        (kind, relative): stat_value
        for kind, relative, stat_value in initial_manifest
    }
    # Open a state-owned persistent cache only after the complete current Osa
    # file set has passed the source safety checks.  ``begin`` is a no-op for
    # the ordinary in-process cache.
    if parse_cache is not None:
        parse_cache.begin()

    folder_observations: list[dict[str, Any]] = []
    message_observations: list[dict[str, Any]] = []
    for kind, path in files:
        source_stat = initial_stats[(kind, _relative_name(path, root))]
        parsed: tuple[list[dict[str, Any]], list[dict[str, Any]], str] | None = None
        if parse_cache is not None:
            try:
                parsed = parse_cache.get(
                    profile_path, root, kind, path, source_stat
                )
            except OSError as error:
                evidence["bad_files"] += 1
                raise OsaScopeError(
                    f"Osa log {_relative_name(path, root)} is unreadable or invalid"
                ) from error
        try:
            if parsed is None:
                parsed = _parse_log(path, kind, root)
                if parse_cache is not None:
                    parse_cache.put(
                        profile_path, root, kind, path, source_stat, parsed
                    )
            folder_rows, message_rows, _ = parsed
        except _BoundError as error:
            if parse_cache is not None:
                parse_cache.invalidate(profile_path, root, kind, path)
            evidence["oversized_files"] += 1
            evidence["bad_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} exceeds the safety limit") from error
        except _UnsafeXML as error:
            if parse_cache is not None:
                parse_cache.invalidate(profile_path, root, kind, path)
            evidence["unsafe_xml_files"] += 1
            evidence["bad_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} contains unsafe XML") from error
        except _UnstableLog as error:
            if parse_cache is not None:
                parse_cache.invalidate(profile_path, root, kind, path)
            evidence["bad_files"] += 1
            evidence["unstable_files"] += 1
            raise OsaScopeError(f"Osa log {_relative_name(path, root)} changed during read") from error
        except (OSError, EOFError, gzip.BadGzipFile, zlib.error, ET.ParseError, ValueError) as error:
            if parse_cache is not None:
                parse_cache.invalidate(profile_path, root, kind, path)
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
        if parse_cache is not None:
            parse_cache.flush()
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
        received = _received_time_evidence(candidates)
        participants, participant_status = _latest_participants(candidates)
        participant_counts = _latest_participant_counts(candidates)
        evidence[f"messages_received_{received['received_time_status']}"] += 1
        messages[identifier] = {
            "folder_id": public_id_by_key[folder_key],
            "folder_status": folder_status,
            **received,
            "participants": participants,
            "participants_status": participant_status,
            "participant_counts": participant_counts,
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
    if parse_cache is not None:
        parse_cache.flush()
    return {"folders": folders, "messages": messages, "evidence": evidence, "fingerprint": fingerprint}
