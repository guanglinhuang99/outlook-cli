"""Inbox and Sent scoped, immutable views over the read-only Outlook cache parser."""
from __future__ import annotations

from collections import Counter, OrderedDict
from contextlib import contextmanager
from copy import deepcopy
from datetime import date, datetime, timezone
import fcntl
import json
import os
from pathlib import Path
import re
import shutil
import stat
import tempfile
import time
from types import SimpleNamespace
import uuid

import outlook_local as core
from osa_scope import LogParseCache

VIEW_VERSION = 5
RETAINED_VIEWS = 3
MAX_EXPORT_FILES = 64
MAX_EXPORT_BYTES = 512 * 1024**2
BUNDLE_SOURCE_BYTES = 64 * 1024**2
BUNDLE_VERIFY_BYTES = 64 * 1024**2
INDEX_CACHE_LIMIT = 3
INDEX_FINGERPRINT_FIELDS = (
    "st_dev", "st_ino", "st_size", "st_mtime_ns", "st_ctime_ns",
    "st_mode", "st_nlink", "st_uid", "st_gid",
)
VERIFIED_MEMBERSHIP = {"verified_latest_observation", "verified_hx_folder_membership"}
TOKEN = re.compile(r"[a-f0-9]{24}\Z")
COVERAGE = ("Locally recoverable cache only; Inbox and Sent Items with all known descendants. "
            "Deleted Items and its descendants are excluded. "
            "Unresolved folder membership is excluded. Folder membership is the latest "
            "reliable local observation, not a live server guarantee. Dates use verified last delivery time "
            "in Asia/Singapore; missing bodies and attachments remain explicit.")


def folder_tree(folders):
    """Keep only unambiguous nodes with a proven parent chain to an Inbox or Sent root."""
    nodes, conflicts = {}, set()
    for folder in folders:
        identifier = folder["id"]
        if identifier in nodes and any(nodes[identifier].get(k) != folder.get(k)
                                       for k in ("parent_id", "is_inbox", "is_sent", "is_deleted", "name")):
            conflicts.add(identifier)
        nodes[identifier] = dict(folder)
    for identifier in conflicts:
        nodes.pop(identifier, None)
    blocked = {key for key, node in nodes.items() if node.get("is_deleted") is True}
    while True:
        children = {key for key, node in nodes.items() if node.get("parent_id") in blocked}
        if children.issubset(blocked):
            break
        blocked.update(children)
    nodes = {key: node for key, node in nodes.items() if key not in blocked}
    allowed = {}
    for identifier, folder in nodes.items():
        chain, seen, cursor = [], set(), identifier
        while cursor in nodes and cursor not in seen:
            seen.add(cursor)
            node = nodes[cursor]
            chain.append(node)
            if node.get("is_deleted") is True:
                break
            if node.get("is_inbox") is True or node.get("is_sent") is True:
                allowed[identifier] = {**folder, "path": "/".join(n["name"] for n in reversed(chain)),
                                       "depth": len(chain) - 1, "scope_root_id": cursor}
                break
            cursor = node.get("parent_id")
    return allowed


class MailService:
    def __init__(self, state, profile=None, from_date="2000-01-01", refresh_seconds=300):
        date.fromisoformat(from_date)
        self.state = Path(state).expanduser().resolve()
        self.profile = profile
        self.from_date = from_date
        self.refresh_seconds = refresh_seconds
        self._attachment_exports = OrderedDict()
        self._index_cache = OrderedDict()
        # The cache opens lazily during refresh, after the private state guard.
        self._osa_parse_cache = LogParseCache(path=self.state / "osa-metadata.sqlite3")

    @contextmanager
    def _lock(self):
        profile = core.discover_profile(self.profile)
        core.state_root(self.state, profile, create=True)
        descriptor = os.open(self.state / ".refresh.lock", os.O_CREAT | os.O_RDWR | os.O_NOFOLLOW, 0o600)
        try:
            deadline = time.monotonic() + 150
            while True:
                try:
                    fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
                    break
                except BlockingIOError:
                    if time.monotonic() >= deadline:
                        raise ValueError("another client is refreshing; retry shortly")
                    time.sleep(0.1)
            yield profile
        finally:
            os.close(descriptor)

    @staticmethod
    def _index_fingerprint(status):
        return tuple(getattr(status, field) for field in INDEX_FINGERPRINT_FIELDS)

    @classmethod
    def _read_index(cls, path):
        """Read one immutable index, retrying once if its file changes."""
        for attempt in range(2):
            core._private_file(path, "scoped snapshot index")
            descriptor = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0))
            owned = True
            try:
                before_status = os.fstat(descriptor)
                if not stat.S_ISREG(before_status.st_mode):
                    raise ValueError("scoped snapshot index must be a regular file")
                before = cls._index_fingerprint(before_status)
                with os.fdopen(descriptor, "rb") as handle:
                    owned = False
                    raw = handle.read()
                    after_status = os.fstat(handle.fileno())
                after = cls._index_fingerprint(after_status)
            finally:
                if owned:
                    os.close(descriptor)
            core._private_file(path, "scoped snapshot index")
            path_fingerprint = cls._index_fingerprint(path.stat())
            if before != after or after != path_fingerprint:
                if attempt == 0:
                    continue
                raise ValueError("scoped snapshot index changed during read")
            return json.loads(raw), after
        raise ValueError("scoped snapshot index changed during read")

    def _validate_index(self, snapshot, view, index, profile, *, verify_checksum):
        if (index.get("view_version") != VIEW_VERSION or index.get("snapshot") != snapshot
                or index.get("schema_version") != core.VERSION):
            raise ValueError("unsupported or mismatched scoped snapshot")
        if verify_checksum:
            content = {k: v for k, v in index.items() if k != "snapshot"}
            if core.sha(json.dumps(content, sort_keys=True, ensure_ascii=False).encode())[:24] != snapshot:
                raise ValueError("scoped snapshot checksum mismatch")
        if Path(index["profile"]).resolve() != profile:
            raise ValueError("snapshot belongs to a different Outlook profile; refresh before reading")
        base = index["source_snapshot"]
        if not TOKEN.fullmatch(base):
            raise ValueError("invalid source snapshot identifier")
        directory = self.state / "snapshots" / base
        if directory.exists() or directory.is_symlink():
            core._private_directory(directory, "source snapshot directory")
            for name in ("payloads", "textpayloads", "deferred"):
                root = directory / name
                if root.exists() or root.is_symlink():
                    core._private_directory(root, f"source snapshot {name} directory")
        if directory.is_symlink() or (directory / "payloads").is_symlink():
            raise ValueError("source snapshot entries must not be symbolic links")
        core.outside_source(self.state, index["profile"])
        return directory

    def _load(self, snapshot=None, *, copy_index=True):
        core.state_root(self.state)
        if any((self.state / name).is_symlink() for name in ("CURRENT_SCOPE", "views", "snapshots")):
            raise ValueError("state roots and pointer must not be symbolic links")
        if snapshot is None:
            pointer = self.state / "CURRENT_SCOPE"
            core._private_file(pointer, "scoped snapshot pointer")
            snapshot = pointer.read_text().strip()
        if not TOKEN.fullmatch(snapshot):
            raise ValueError("invalid snapshot identifier")
        view = self.state / "views" / snapshot
        if view.exists() or view.is_symlink():
            core._private_directory(view, "scoped view directory")
        index_path = view / "index.json"
        if view.is_symlink() or index_path.is_symlink():
            raise ValueError("snapshot entries must not be symbolic links")
        if not index_path.exists():
            raise FileNotFoundError(index_path)
        core._private_file(index_path, "scoped snapshot index")
        profile = core.discover_profile(self.profile)
        fingerprint = self._index_fingerprint(index_path.stat())
        cached = self._index_cache.get(snapshot)
        if cached is not None and cached["fingerprint"] == fingerprint:
            index = cached["index"]
            directory = self._validate_index(snapshot, view, index, profile, verify_checksum=False)
            self._index_cache.move_to_end(snapshot)
            return directory, deepcopy(index) if copy_index else index

        index, fingerprint = self._read_index(index_path)
        directory = self._validate_index(snapshot, view, index, profile, verify_checksum=True)
        self._index_cache[snapshot] = {"fingerprint": fingerprint, "index": index}
        self._index_cache.move_to_end(snapshot)
        while len(self._index_cache) > INDEX_CACHE_LIMIT:
            self._index_cache.popitem(last=False)
        return directory, deepcopy(index) if copy_index else index

    def _fresh(self):
        try:
            check = json.loads((self.state / "REFRESH.json").read_text())
            _, index = self._load(copy_index=False)
            return (check["snapshot"] == index["snapshot"] and
                    0 <= time.time() - check["checked_unix"] < self.refresh_seconds and
                    core.day_start(index["summary"]["scope_from_date"]) <= core.day_start(self.from_date))
        except (OSError, ValueError, KeyError, TypeError):
            return False

    def _write_json(self, target, value):
        if target.is_symlink():
            raise ValueError("state entry must not be a symbolic link")
        with tempfile.NamedTemporaryFile(mode="w", dir=self.state, prefix=".publish-", delete=False) as handle:
            json.dump(value, handle, ensure_ascii=False)
            temporary = Path(handle.name)
        try:
            temporary.replace(target)
        finally:
            temporary.unlink(missing_ok=True)

    def _refresh(self, from_date=None, automatic=False):
        requested = from_date or self.from_date
        date.fromisoformat(requested)
        with self._lock() as profile:
            if automatic and self._fresh():
                return self._status()
            self._clean_abandoned_builds()
            previous_base = None
            source_pointer = self.state / "CURRENT"
            if source_pointer.exists() or source_pointer.is_symlink():
                core._private_file(source_pointer, "source snapshot pointer")
                previous_base = source_pointer.read_text().strip()
                if not TOKEN.fullmatch(previous_base):
                    raise ValueError("invalid source snapshot pointer")
            # A failed refresh must not leave the previous scope marked fresh.
            (self.state / "REFRESH.json").unlink(missing_ok=True)
            # The raw parser knows sent time only. Do not trim by sent date before
            # associating delivery evidence: a much older mail may arrive today.
            core.make_snapshot(SimpleNamespace(
                state=str(self.state), profile=str(profile), from_date="0001-01-01",
                freeze_bodies=False,
            ))
            _, base = core.load_index(self.state)
            from osa_scope import read_scope
            try:
                scope = read_scope(profile, base["items"], parse_cache=self._osa_parse_cache)
                folders = folder_tree(scope["folders"])
                if not folders:
                    raise ValueError("no verified Inbox or Sent folder tree is available in the local cache; previous view retained")
            except (OSError, ValueError):
                self._prune_failed_sources(previous_base)
                raise
            items, outside, unknown = [], 0, 0
            cutoff = core.day_start(requested)
            for row in base["items"]:
                membership = scope["messages"].get(row["id"])
                if (membership is None or not membership.get("folder_id") or
                        membership.get("folder_status") not in VERIFIED_MEMBERSHIP):
                    unknown += 1
                    continue
                identifier = membership["folder_id"]
                if identifier not in folders:
                    outside += 1
                    continue
                received = membership.get("received_unix")
                if received is not None and received < cutoff:
                    continue
                items.append({**row, "received_unix": received,
                              **core.merge_participants(row, membership),
                              "received_time_status": membership.get("received_time_status", "missing"),
                              "received_time_method": membership.get("received_time_method"),
                              "folder_id": identifier, "folder_path": folders[identifier]["path"],
                              "folder_status": membership["folder_status"],
                              "folder_evidence": membership.get("evidence", {}),
                              "mailbox_state": "inbox_or_sent_latest_local_observation"})
                # Keep object-key/version diagnostics in the source snapshot,
                # without copying the full contact collection twice per view.
                items[-1].pop("hx_participants", None)
            items.sort(key=core.received_sort_key)
            folder_rows = sorted(folders.values(), key=lambda f: (f["scope_root_id"], f["path"], f["id"]))
            cached_message_counts = Counter(row["folder_id"] for row in items)
            for folder in folder_rows:
                folder["cached_messages"] = cached_message_counts[folder["id"]]
            summary = {"scope_from_date": requested, "timezone": "Asia/Singapore", "date_basis": "received_at",
                       "messages": len(items), "folders": len(folder_rows),
                       "received_time_known": sum(r["received_unix"] is not None for r in items),
                       "received_time_unknown": sum(r["received_unix"] is None for r in items),
                       "inbox_roots": sum(bool(f.get("is_inbox")) for f in folder_rows),
                       "sent_roots": sum(bool(f.get("is_sent")) for f in folder_rows),
                       "outside_scope_excluded": outside, "unresolved_folder_excluded": unknown,
                       "source_messages": len(base["items"]),
                       "cached_bodies": sum(r["body"]["status"] == "cached_html" for r in items),
                       "deferred_bodies": sum(r["body"]["status"] == "deferred" for r in items),
                       "linked_attachments": sum(len(r["attachments"]) for r in items),
                       "body_statuses": {s: sum(r["body"]["status"] == s for r in items)
                                         for s in ("cached_html", "deferred", "not_cached", "not_linked", "unavailable")},
                       "body_diagnoses": {code: sum(r["body"]["diagnosis"]["code"] == code for r in items)
                                          for code in sorted({r["body"]["diagnosis"]["code"] for r in items})},
                       "source_exclusions": {key: base["summary"].get(key) for key in (
                           "identity_conflicts_excluded", "time_conflicts_excluded", "unresolved_times_excluded",
                           "unparsed_header_frames", "attachment_conflicts_excluded")},
                       "container": base["summary"].get("container"),
                       "coverage": COVERAGE}
            index = {**base, "view_version": VIEW_VERSION, "source_snapshot": base["snapshot"],
                     "items": items, "folders": folder_rows,
                     "scope_policy": "inbox-sent-exclude-deleted", "scope_fingerprint": scope["fingerprint"],
                     "scope_evidence": scope["evidence"], "summary": summary}
            index.pop("snapshot")
            snapshot = core.sha(json.dumps(index, sort_keys=True, ensure_ascii=False).encode())[:24]
            index["snapshot"] = snapshot
            views = self.state / "views"
            if views.is_symlink():
                raise ValueError("views directory must not be a symbolic link")
            views.mkdir(exist_ok=True, mode=0o700)
            core._private_directory(views, "views root")
            final = views / snapshot
            if final.is_symlink():
                raise ValueError("view directory must not be a symbolic link")
            if not final.exists():
                build = Path(tempfile.mkdtemp(prefix=f".view-{os.getpid()}-", dir=self.state))
                try:
                    (build / "index.json").write_text(json.dumps(index, ensure_ascii=False))
                    (build / "index.json").chmod(0o400)
                    build.rename(final)
                finally:
                    if build.exists():
                        import shutil
                        shutil.rmtree(build)
            # Publish the complete immutable view before its pointer, under the shared lock.
            with tempfile.NamedTemporaryFile(mode="w", dir=self.state, prefix=".scope-", delete=False) as handle:
                handle.write(snapshot)
                temporary = Path(handle.name)
            try:
                temporary.replace(self.state / "CURRENT_SCOPE")
            finally:
                temporary.unlink(missing_ok=True)
            self._prune(snapshot)
            self._write_json(self.state / "REFRESH.json", {"snapshot": snapshot, "checked_unix": time.time()})
            if from_date is not None:
                self.from_date = requested
            return self._status()

    def _prune(self, current):
        """Keep up to three views sharing the current compressed source copy."""
        views = self.state / "views"
        candidates = [path for path in views.iterdir()
                      if TOKEN.fullmatch(path.name) and path.is_dir() and not path.is_symlink()]
        candidates.sort(key=lambda path: path.stat().st_mtime_ns, reverse=True)
        retained = {current}
        _, current_index = self._load(current)
        bases = {current_index["source_snapshot"]}
        for path in candidates:
            if len(retained) >= RETAINED_VIEWS:
                break
            if path.name == current:
                continue
            try:
                _, index = self._load(path.name)
            except (OSError, ValueError, KeyError, TypeError):
                # Old schemas and broken historical views cannot serve reads.
                continue
            if index["source_snapshot"] != current_index["source_snapshot"]:
                continue
            retained.add(path.name)
            bases.add(index["source_snapshot"])
        # The standalone parser may still use its latest unscoped base snapshot.
        latest_base = (self.state / "CURRENT").read_text().strip()
        if not TOKEN.fullmatch(latest_base):
            raise ValueError("invalid source snapshot pointer")
        bases.add(latest_base)
        for root, keep in ((views, retained), (self.state / "snapshots", bases),
                           (self.state / "exports", retained)):
            if root.is_symlink():
                raise ValueError("state roots must not be symbolic links")
            if not root.exists():
                continue
            for path in root.iterdir():
                # Delete only recognized service entries, never follow links.
                if TOKEN.fullmatch(path.name) and path.name not in keep:
                    if path.is_symlink() or not path.is_dir():
                        path.unlink()
                    else:
                        shutil.rmtree(path)
                elif root.name == "exports" and re.fullmatch(r"[a-f0-9]{64}", path.name):
                    # Older versions stored individual verified exports at the root.
                    if path.is_file() or path.is_symlink():
                        path.unlink()
        self._attachment_exports = OrderedDict(
            (key, value) for key, value in self._attachment_exports.items() if key[0] in retained)
        self._prune_exports()

    def _clean_abandoned_builds(self):
        """The refresh lock excludes another service build; preserve live standalone owners."""
        for path in self.state.iterdir():
            if not path.name.startswith((".building-", ".view-")):
                continue
            if not self._abandoned_temporary(path):
                continue
            if path.is_symlink() or not path.is_dir():
                path.unlink(missing_ok=True)
            else:
                shutil.rmtree(path)

    @staticmethod
    def _abandoned_temporary(path):
        owner = re.match(r"\.(?:building|view|export)-(\d+)-", path.name)
        if owner:
            pid = int(owner[1])
            if pid <= 0:
                return False
            try:
                os.kill(pid, 0)
            except ProcessLookupError:
                return True
            except PermissionError:
                pass
            return False
        # Prior versions lacked owners; let their worker deadline expire.
        return time.time() - path.lstat().st_mtime >= 185

    def _prune_failed_sources(self, previous_base):
        """Failed scope refreshes may replace the base but must not accumulate it."""
        if previous_base is not None:
            # A base without a published scope cannot serve this service. Roll
            # back without interpreting a previous schema during an upgrade.
            core.publish_current(self.state, previous_base)
        else:
            (self.state / "CURRENT").unlink(missing_ok=True)
        keep = {previous_base} if previous_base is not None else set()
        views = self.state / "views"
        if views.is_symlink():
            raise ValueError("views directory must not be a symbolic link")
        if views.exists():
            for view in views.iterdir():
                index_path = view / "index.json"
                if not TOKEN.fullmatch(view.name) or view.is_symlink() or index_path.is_symlink():
                    continue
                # Preserve bases referenced even by an older schema until successful upgrade.
                index = json.loads(index_path.read_text())
                if TOKEN.fullmatch(index.get("source_snapshot", "")):
                    keep.add(index["source_snapshot"])
        snapshots = self.state / "snapshots"
        if snapshots.is_symlink():
            raise ValueError("snapshots directory must not be a symbolic link")
        for path in snapshots.iterdir():
            if TOKEN.fullmatch(path.name) and path.name not in keep:
                if path.is_symlink() or not path.is_dir():
                    path.unlink()
                else:
                    shutil.rmtree(path)

    def _prune_exports(self, protected=None, reserve_bytes=0, reserve_files=0):
        """Bound private export files across retained views, under the shared lock."""
        if protected is not None:
            protected = Path(protected).resolve()
        root = self.state / "exports"
        if root.is_symlink():
            raise ValueError("exports directory must not be a symbolic link")
        if not root.exists():
            return
        files, directories = [], [root]
        for directory in root.iterdir():
            if not TOKEN.fullmatch(directory.name):
                continue
            if directory.is_symlink():
                raise ValueError("export directory must not be a symbolic link")
            if not directory.is_dir():
                continue
            directories.append(directory)
        count, size = reserve_files, reserve_bytes
        for directory in directories:
            for path in directory.iterdir():
                if path.name.startswith(".export-"):
                    if self._abandoned_temporary(path):
                        if path.is_symlink() or path.is_file():
                            path.unlink()
                        continue
                    if path.is_symlink() or not path.is_file():
                        raise ValueError("temporary export must be a regular file")
                    count += 1
                    size += path.stat().st_size
                    continue
                if not re.fullmatch(r"[a-f0-9]{64}", path.name):
                    continue
                if path.is_symlink() or not path.is_file():
                    raise ValueError("export must be a regular file")
                files.append((path, path.stat()))
        if count > MAX_EXPORT_FILES or size > MAX_EXPORT_BYTES:
            raise ValueError("attachment cache quota exhausted by active exports; retry later")
        files.sort(key=lambda row: (row[0] == protected, row[1].st_mtime_ns), reverse=True)
        for path, metadata in files:
            if count < MAX_EXPORT_FILES and size + metadata.st_size <= MAX_EXPORT_BYTES:
                count += 1
                size += metadata.st_size
            else:
                if path == protected:
                    raise ValueError("attachment cache quota exhausted by active exports; retry later")
                path.unlink()
        self._attachment_exports = OrderedDict(
            (key, value) for key, value in self._attachment_exports.items()
            if Path(value["output"]).exists())

    def _status(self):
        if not (self.state / "CURRENT_SCOPE").exists():
            return {"ready": False, "scope": "Inbox and Sent Items, including descendants; excluding Deleted Items",
                    "state": str(self.state), "next_step": "call outlook_local_refresh or outlook_local_search"}
        try:
            _, index = self._load(copy_index=False)
        except (OSError, ValueError, KeyError, TypeError) as error:
            return {"ready": False, "state": str(self.state), "error": str(error),
                    "next_step": "call outlook_local_refresh or outlook_local_search"}
        checked_at = None
        try:
            check = json.loads((self.state / "REFRESH.json").read_text())
            if check["snapshot"] == index["snapshot"]:
                checked_at = datetime.fromtimestamp(check["checked_unix"], timezone.utc).isoformat()
        except (OSError, ValueError, KeyError, TypeError):
            pass
        return {"ready": True, "snapshot": index["snapshot"], "captured_at": index["created_at"],
                "last_checked_at": checked_at, "state": str(self.state), **index["summary"],
                "scope_evidence": index["scope_evidence"]}

    def _current(self, snapshot=None):
        if not self._fresh():
            self._refresh(automatic=True)
        # Internal readers borrow the validated snapshot; public results are
        # copied at the call boundary so clients cannot mutate cached records.
        directory, index = self._load(snapshot, copy_index=False)
        if snapshot is not None:
            _, latest = self._load(copy_index=False)
            allowed = {row["id"] for row in latest["items"]}
            folders = {row["id"]: row for row in latest["folders"]}
            if (any(row["id"] not in allowed for row in index["items"])
                    or any(row["id"] not in folders or any(
                        row.get(key) != folders[row["id"]].get(key)
                        for key in ("parent_id", "scope_root_id")) for row in index["folders"])):
                raise ValueError("folder access changed since this snapshot; repeat search against current scope")
        return directory, index

    @staticmethod
    def _metadata(result, index):
        rows = {row["id"]: row for row in index["items"]}
        values = result["items"] if "items" in result else [result]
        for value in values:
            row = rows[value["id"]]
            value.update({k: row[k] for k in ("folder_id", "folder_path", "folder_status")})
        return result

    def _read(self, directory, index, identifier, maximum):
        result = core.read_message(directory, index, identifier, maximum)
        result["body"].pop("text_chunk_sha256", None)
        self._metadata(result, index)
        result["folder_evidence"] = core.find_message(index, identifier)["folder_evidence"]
        return result

    def _search(self, arguments):
        cursor = arguments.get("cursor")
        snapshot, raw_cursor = None, None
        folder_id = arguments.get("folder_id")
        descendants = arguments.get("include_subfolders", True)
        folder_signature = core.sha(json.dumps([folder_id, descendants]).encode())[:16]
        if cursor:
            parts = cursor.split(".", 3)
            if len(parts) != 4 or parts[0] != "v1" or parts[2] != folder_signature:
                raise ValueError("cursor belongs to another folder filter or is invalid")
            _, snapshot, _, raw_cursor = parts
        _, index = self._current(snapshot)
        folders = {folder["id"]: folder for folder in index["folders"]}
        filtered = index
        if folder_id is not None:
            if folder_id not in folders:
                raise ValueError("folder is outside the verified Inbox and Sent scope")
            allowed = {folder_id}
            if descendants:
                changed = True
                while changed:
                    children = {f["id"] for f in folders.values() if f.get("parent_id") in allowed}
                    changed = not children.issubset(allowed)
                    allowed.update(children)
            filtered = {**index, "items": [r for r in index["items"] if r["folder_id"] in allowed]}
        options = SimpleNamespace(query=arguments.get("query", ""), field=arguments.get("field", "subject"),
                                  person=arguments.get("person"), person_role=arguments.get("person_role", "any"),
                                  from_date=arguments.get("from_date"), to_date=arguments.get("to_date"),
                                  limit=arguments.get("limit", 20), cursor=raw_cursor)
        result = core.search(filtered, options)
        self._metadata(result, index)
        if result["next_cursor"]:
            result["next_cursor"] = f"v1.{index['snapshot']}.{folder_signature}.{result['next_cursor']}"
        result.update(date_basis="received_at", timezone="Asia/Singapore", folder_id=folder_id,
                      include_subfolders=descendants, folder_scope_complete=False)
        return result

    def _attachment(self, directory, index, arguments):
        with self._lock():
            return self._export_attachment(directory, index, arguments)

    def _digest(self, arguments):
        page = self._search({**arguments, "limit": arguments.get("limit", 25)})
        directory, index = self._current(page["snapshot"])
        maximum = arguments.get("body_max_chars", 12000)
        known_versions = set(arguments.get("known_versions", []))
        known_bodies = set(arguments.get("known_bodies", []))
        items, unchanged, bodies, errors = [], [], {}, []
        source_budget = {"remaining": BUNDLE_SOURCE_BYTES,
                         "verification_remaining": BUNDLE_VERIFY_BYTES}
        fields = (
            "id", "subject", "subject_status", "sender", "sender_name", "sender_status",
            "participants", "participants_status", "preview", "sent_at", "received_at",
            "received_time_status", "folder_id", "folder_path", "folder_status",
            "attachment_count", "attachments_complete",
        )
        for row in page["items"]:
            try:
                with core.source_read_budget(source_budget):
                    message = self._read(directory, index, row["id"], maximum)
                item = {key: message.get(key) for key in fields}
                item["attachments"] = [
                    {key: attachment.get(key) for key in ("id", "name", "status")}
                    for attachment in message.get("attachments", [])
                ]
                body = message["body"]
                item["body"] = {key: body.get(key) for key in (
                    "status", "diagnosis", "text_chars", "truncated", "text_sha256")}
                content = None
                if "text" in body:
                    content = {key: body.get(key) for key in (
                        "text", "truncated", "text_sha256", "text_chars")}
                    # A matching prefix alone cannot prove identical full bodies.
                    identity = [content, None if body.get("text_sha256") else row["id"]]
                    content_id = core.sha(json.dumps(identity, sort_keys=True, ensure_ascii=False).encode())
                    item["body"]["content_id"] = content_id
                version = core.sha(json.dumps(["digest-v1", item], sort_keys=True,
                                              ensure_ascii=False).encode())
                item["version"] = version
                if version in known_versions:
                    unchanged.append({"id": row["id"], "version": version})
                else:
                    items.append(item)
                    if content is not None and content_id not in known_bodies:
                        bodies[content_id] = content
            except (ValueError, OSError) as error:
                errors.append({"id": row["id"], "error": str(error)})
        return {**page, "page_ids": [row["id"] for row in page["items"]],
                "items": items, "unchanged": unchanged, "bodies": bodies, "errors": errors,
                "body_text_chars": sum(len(body["text"]) for body in bodies.values()),
                "coverage": COVERAGE}

    def _export_attachment(self, directory, index, arguments):
        export_root = self.state / "exports"
        exports = export_root / index["snapshot"]
        if export_root.is_symlink() or exports.is_symlink():
            raise ValueError("exports directory must not be a symbolic link")
        export_root.mkdir(exist_ok=True, mode=0o700)
        core._private_directory(export_root, "exports root")
        exports.mkdir(exist_ok=True, mode=0o700)
        core._private_directory(exports, "export directory")
        message = core.find_message(index, arguments["id"])
        attachment = next((row for row in message["attachments"]
                           if row["id"] == arguments["attachment_id"]), None)
        if attachment is None:
            raise ValueError("attachment does not belong to this message")
        if attachment["status"] != "cached_at_snapshot":
            raise ValueError("attachment is not cached")
        source = core.source_path(Path(index["profile"]), attachment["path"])
        if core.fingerprint(source) != attachment["fingerprint"]:
            raise ValueError("attachment is missing or changed since snapshot; build a fresh snapshot")
        key = (index["snapshot"], arguments["id"], arguments["attachment_id"])
        cached = self._attachment_exports.get(key)
        if cached is not None:
            target = Path(cached["output"])
            if target.is_symlink():
                raise ValueError("attachment export changed since verification")
            try:
                current_fingerprint = core.fingerprint(target)
            except FileNotFoundError:
                del self._attachment_exports[key]
            else:
                if current_fingerprint != cached["_export_fingerprint"]:
                    raise ValueError("attachment export changed since verification")
                self._attachment_exports.move_to_end(key)
                self._prune_exports(protected=target)
                return dict(cached)
        self._prune_exports(reserve_bytes=attachment["fingerprint"]["size"], reserve_files=1)
        temporary = exports / f".export-{os.getpid()}-{uuid.uuid4().hex}"
        try:
            result = core.read_attachment(directory, index, arguments["id"], arguments["attachment_id"], temporary)
            target = exports / result["sha256"]
            if target.is_symlink():
                raise ValueError("attachment export must not be a symbolic link")
            if target.exists():
                if core.sha(target.read_bytes()) != result["sha256"]:
                    raise ValueError("existing attachment export checksum mismatch")
            else:
                temporary.chmod(0o400)
                try:
                    os.link(temporary, target)
                except FileExistsError:
                    if target.is_symlink() or core.sha(target.read_bytes()) != result["sha256"]:
                        raise ValueError("concurrent attachment export checksum mismatch")
            temporary.unlink(missing_ok=True)
            result = {**result, "output": str(target), "source_read_only": True,
                      "_export_fingerprint": core.fingerprint(target)}
            self._attachment_exports[key] = result
            if len(self._attachment_exports) > MAX_EXPORT_FILES:
                _, evicted = self._attachment_exports.popitem(last=False)
                if not any(row["output"] == evicted["output"] for row in self._attachment_exports.values()):
                    Path(evicted["output"]).unlink(missing_ok=True)
            self._prune_exports(protected=target)
            return dict(result)
        finally:
            temporary.unlink(missing_ok=True)

    def _today(self, arguments):
        day = datetime.now(core.ZONE).date().isoformat()
        options = {"limit": 100, **arguments}
        if options.get("cursor"):
            parts = options["cursor"].split("|", 2)
            if len(parts) != 3 or parts[0] != "today-v1" or not parts[2]:
                raise ValueError("invalid today cursor; use a cursor from outlook_local_today")
            _, day, options["cursor"] = parts
            if date.fromisoformat(day).isoformat() != day:
                raise ValueError("invalid date in today cursor")
        result = self._search({**options, "from_date": day, "to_date": day})
        result["date"] = day
        if result["next_cursor"]:
            result["next_cursor"] = f"today-v1|{day}|{result['next_cursor']}"
        return result

    def call(self, name, arguments):
        return deepcopy(self._call(name, arguments))

    def _call(self, name, arguments):
        if name == "outlook_local_status":
            return self._status()
        if name == "outlook_local_refresh":
            return self._refresh(arguments.get("from_date"))
        if name == "outlook_local_search":
            return self._search(arguments)
        if name == "outlook_local_today":
            return self._today(arguments)
        if name == "outlook_local_digest":
            return self._digest(arguments)
        directory, index = self._current(arguments.get("snapshot"))
        if name == "outlook_local_folders":
            return {"snapshot": index["snapshot"], "items": index["folders"], "coverage": COVERAGE}
        if name == "outlook_local_read":
            return self._read(directory, index, arguments["id"], arguments.get("body_max_chars", 20000))
        if name == "outlook_local_attachments":
            message = core.find_message(index, arguments["id"])
            return {"snapshot": index["snapshot"], "message_id": message["id"],
                    "items": message["attachments"], "attachments_complete": False}
        if name == "outlook_local_attachment":
            result = self._attachment(directory, index, arguments)
            return {key: value for key, value in result.items() if key != "_export_fingerprint"}
        if name == "outlook_local_messages_bundle":
            budget = arguments.get("text_budget_chars", 60000)
            items, errors, used = [], [], 0
            source_budget = {"remaining": BUNDLE_SOURCE_BYTES,
                             "verification_remaining": BUNDLE_VERIFY_BYTES}
            for identifier in dict.fromkeys(arguments["ids"]):
                if budget <= 0:
                    errors.append({"id": identifier, "error": "bundle text budget exhausted"})
                    continue
                try:
                    # Charge only actual source reads and body decoding. A
                    # resolved sidecar can serve a body without using this budget.
                    with core.source_read_budget(source_budget):
                        item = self._read(directory, index, identifier, min(arguments.get("body_max_chars", 12000), budget))
                    length = len(item["body"].get("text", ""))
                    budget -= length
                    used += length
                    items.append(item)
                except (ValueError, OSError) as error:
                    errors.append({"id": identifier, "error": str(error)})
            return {"snapshot": index["snapshot"], "items": items, "errors": errors,
                    "body_text_chars": used, "coverage": COVERAGE}
        raise ValueError("unknown Outlook tool")
