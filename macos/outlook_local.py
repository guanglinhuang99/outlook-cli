#!/usr/bin/env python3
"""Standalone, read-only Outlook cache prototype. All writes target private state."""

from __future__ import annotations

import argparse
from collections import OrderedDict, defaultdict
from datetime import date, datetime, time, timedelta, timezone
import errno
import gzip
import hashlib
from html.parser import HTMLParser
import io
import json
import mmap
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import sys
import tempfile
import time as clock
import zlib
from zoneinfo import ZoneInfo

from hxcontainer import iter_blocks
from hxrecords import (entry_ids, file_references, iter_frames, normalize_subject,
                       parse_attachment, parse_header, parse_inline_body, typed_references)


VERSION = 17
MAX_BODY = 64 * 1024 * 1024
MAX_ATTACHMENT = 256 * 1024 * 1024
MAX_IN_MEMORY_SOURCE = 64 * 1024**2
SOURCE_COPY_CHUNK = 1024 * 1024
TEXT_CHUNK_BYTES = 64 * 1024
VERIFIED_CACHE_LIMIT = 64
MARKER = ".outlook-local-prototype"
MARKER_CONTENT = "outlook-local-prototype"
DEFAULT_STATE = Path.home() / "Library/Caches/outlook-local-prototype"
ZONE = ZoneInfo("Asia/Singapore")

_verified_raw_payloads = OrderedDict()
_verified_text_chunks = OrderedDict()


def sha(data):
    return hashlib.sha256(data).hexdigest()


def emit(value):
    print(json.dumps(value, ensure_ascii=False, indent=2))


def outside_source(path, profile):
    destination = Path(path).expanduser().resolve()
    protected = [Path(profile).resolve(), Path.home() / "Library/Mail",
                 Path.home() / "Library/Group Containers/UBF8T346G9.Office/Outlook"]
    if any(destination.is_relative_to(root.resolve()) for root in protected):
        raise ValueError("write destination must be outside all original mail/cache directories")
    return destination


def _private_directory(path, label):
    if path.is_symlink():
        raise ValueError(f"{label} must not be symbolic links")
    try:
        status = path.stat()
    except FileNotFoundError as error:
        raise ValueError(f"{label} is missing") from error
    if not path.is_dir():
        raise ValueError(f"{label} must be a directory")
    if status.st_uid != os.getuid():
        raise ValueError(f"{label} owner is not the current user")
    if status.st_mode & 0o7777 != 0o700:
        raise ValueError(f"{label} does not have private permissions")
    return path


def _private_file(path, label):
    if path.is_symlink():
        raise ValueError(f"{label} must not be symbolic links")
    try:
        status = path.stat()
    except FileNotFoundError as error:
        raise ValueError(f"{label} is missing") from error
    if not path.is_file():
        raise ValueError(f"{label} must be a regular file")
    if status.st_uid != os.getuid():
        raise ValueError(f"{label} owner is not the current user")
    if status.st_nlink != 1:
        raise ValueError(f"{label} must not be a hard link")
    return path


def _validate_marker(path):
    _private_file(path, "prototype marker")
    try:
        with path.open("r") as handle:
            content = handle.read(len(MARKER_CONTENT) + 1)
    except OSError as error:
        raise ValueError("prototype marker is unreadable") from error
    if content != MARKER_CONTENT:
        raise ValueError("prototype marker content is invalid")


def state_root(path, profile=None, create=False):
    state = Path(path).expanduser()
    if state.is_symlink():
        raise ValueError("state root must not be symbolic links")
    state = state.resolve()
    if profile is not None:
        outside_source(state, profile)
    if create:
        marker = state / MARKER
        if marker.is_symlink() or (state / "snapshots").is_symlink():
            raise ValueError("prototype state entries must not be symbolic links")
        if state.exists():
            if not state.is_dir():
                raise ValueError("state root must be a directory")
            if any(state.iterdir()) and not marker.is_file():
                raise ValueError("state directory is not empty and is not owned by this prototype")
            _private_directory(state, "state root")
            if marker.is_file():
                _validate_marker(marker)
        state.mkdir(parents=True, exist_ok=True, mode=0o700)
        _private_directory(state, "state root")
        if not marker.exists():
            with marker.open("x") as handle:
                handle.write(MARKER_CONTENT)
        _validate_marker(marker)
        snapshots = state / "snapshots"
        snapshots.mkdir(exist_ok=True, mode=0o700)
        _private_directory(snapshots, "snapshot root")
    else:
        if not state.exists():
            raise ValueError("no prototype state; run snapshot first")
        marker = state / MARKER
        if marker.is_symlink():
            raise ValueError("prototype marker must not be symbolic links")
        if not marker.is_file():
            raise ValueError("no prototype state; run snapshot first")
        _private_directory(state, "state root")
        _validate_marker(marker)
        snapshots = state / "snapshots"
        if snapshots.exists() or snapshots.is_symlink():
            _private_directory(snapshots, "snapshot root")
    for name in ("views", "exports"):
        managed = state / name
        if managed.exists() or managed.is_symlink():
            _private_directory(managed, f"{name} root")
    return state


def publish_current(state, snapshot_id):
    # Replace the pointer itself; never follow a pre-existing CURRENT symlink.
    with tempfile.NamedTemporaryFile(mode="w", prefix=".CURRENT-", dir=state, delete=False) as handle:
        handle.write(snapshot_id)
        temporary = Path(handle.name)
    try:
        temporary.replace(state / "CURRENT")
    finally:
        temporary.unlink(missing_ok=True)


def fingerprint(path):
    st = path.stat()
    return {"size": st.st_size, "mtime_ns": st.st_mtime_ns, "inode": st.st_ino}


def stable_read(path, maximum):
    before = fingerprint(path)
    if before["size"] > maximum:
        raise ValueError(f"cache file exceeds {maximum} byte prototype limit")
    with path.open("rb") as handle:
        data = handle.read(maximum + 1)
    if len(data) > maximum or before != fingerprint(path) or len(data) != before["size"]:
        raise ValueError("cache changed during read; retry with a fresh snapshot")
    return data, before


def copy_stable_source(path, directory):
    """Copy a large source into private state without retaining it in RAM."""
    before = fingerprint(path)
    required = before["size"]
    try:
        available = shutil.disk_usage(directory).free
    except OSError as error:
        raise ValueError("insufficient scratch space for HxStore snapshot") from error
    if available < required:
        raise ValueError(
            f"insufficient scratch space for HxStore snapshot: need {required} bytes, "
            f"have {available}"
        )
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(prefix=".source-", dir=directory, delete=False) as handle:
            temporary = Path(handle.name)
            digest = hashlib.sha256()
            copied = 0
            with path.open("rb") as source:
                while copied < required:
                    chunk = source.read(min(SOURCE_COPY_CHUNK, required - copied))
                    if not chunk:
                        break
                    copied += len(chunk)
                    handle.write(chunk)
                    digest.update(chunk)
                extra = source.read(1)
            handle.flush()
        if copied != required or extra or before != fingerprint(path):
            raise ValueError("cache changed during read; retry with a fresh snapshot")
        verification = hashlib.sha256()
        remaining = required
        with path.open("rb") as source:
            while remaining:
                chunk = source.read(min(SOURCE_COPY_CHUNK, remaining))
                if not chunk:
                    break
                verification.update(chunk)
                remaining -= len(chunk)
            extra = source.read(1)
        if (remaining or extra or before != fingerprint(path)
                or digest.hexdigest() != verification.hexdigest()):
            raise ValueError("cache changed during read; retry with a fresh snapshot")
        return temporary, before, digest.hexdigest()
    except OSError as error:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        if error.errno == errno.ENOSPC:
            raise ValueError("insufficient scratch space for HxStore snapshot") from error
        raise
    except Exception:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        raise


def source_path(profile, reference):
    if not reference.startswith("~/Files/"):
        raise ValueError("unsupported cache path")
    relative = PurePosixPath(reference[2:])
    if any(part in ("..", ".") for part in relative.parts) or relative.is_absolute():
        raise ValueError("unsafe cache path")
    root = Path(profile).resolve()
    candidate = (root / str(relative)).resolve()
    if not candidate.is_relative_to(root):
        raise ValueError("cache reference escapes profile")
    return candidate


class BodyText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.hidden = 0
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.hidden += 1
        elif not self.hidden and tag in ("p", "div", "br", "tr", "li"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.hidden = max(0, self.hidden - 1)

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)


def decode_body(data):
    if data.startswith(b"\x1f\x8b"):
        with gzip.GzipFile(fileobj=io.BytesIO(data)) as handle:
            raw = handle.read(MAX_BODY + 1)
        encoding = "gzip-html"
    else:
        raw, encoding = data, "html"
    if len(raw) > MAX_BODY:
        raise ValueError("inflated body exceeds prototype limit")
    html = raw.decode("utf-8-sig", errors="strict")
    if "\0" in html:
        raise ValueError("HTML body contains NUL; unsupported payload boundary")
    if not any(tag in html.lower() for tag in ("<html", "<div", "<body", "<!doctype")):
        raise ValueError("unsupported body format")
    parser = BodyText()
    parser.feed(html)
    return "".join(parser.parts).strip(), encoding


def folded(value):
    # Outlook's preview renders automatic list counters; HTMLParser does not.
    value = re.sub(r"(?<![A-Za-z0-9])(?:\d{1,2}[.、)]|[①-⑳])\s*", "", value or "")
    return "".join(c for c in value.casefold() if not c.isspace() and c not in "\ufeff\u200b")


def frozen_payload(directory, data):
    name = sha(data)
    target = directory / "payloads" / name
    target.parent.mkdir(exist_ok=True, mode=0o700)
    _private_directory(target.parent, "frozen body payload root")
    if not target.exists():
        with target.open("xb") as handle:
            handle.write(data)
        target.chmod(0o400)
    return name


def frozen_text_payload(directory, text):
    """Persist the validated plain text in the unit used by common pagination."""
    data = text.encode("utf-16-le", errors="surrogatepass")
    if len(data) % 2 or TEXT_CHUNK_BYTES % 2:
        raise ValueError("invalid UTF-16 text payload size")
    name = sha(data)
    root = directory / "textpayloads"
    if root.is_symlink():
        raise ValueError("frozen text payload directory must not be a symbolic link")
    root.mkdir(exist_ok=True, mode=0o700)
    _private_directory(root, "frozen text payload root")
    target = root / name
    if target.exists():
        if target.is_symlink() or not target.is_file():
            raise ValueError("frozen text payload must be a regular file")
    else:
        with target.open("xb") as handle:
            handle.write(data)
        target.chmod(0o400)
    return {"text_sha256": name, "text_bytes": len(data),
            "text_utf16_units": len(data) // 2, "text_chars": len(text),
            "text_chunk_bytes": TEXT_CHUNK_BYTES,
            "text_chunk_sha256": [sha(data[offset:offset + TEXT_CHUNK_BYTES])
                                  for offset in range(0, len(data), TEXT_CHUNK_BYTES)]}


def asset_digest(profile):
    """Fingerprint dependent cache files so a body-only update can refresh."""
    manifest = []
    files_root = profile / "Files"
    if files_root.is_symlink():
        raise ValueError("cache asset directory must not be a symbolic link")
    for directory, dirs, files in os.walk(files_root, followlinks=False):
        if any((Path(directory) / name).is_symlink() for name in dirs):
            raise ValueError("cache asset directory must not be a symbolic link")
        for name in files:
            path = Path(directory) / name
            if path.is_symlink():
                raise ValueError("cache asset must not be a symbolic link")
            try:
                manifest.append((str(path.relative_to(profile)), fingerprint(path)))
            except FileNotFoundError:
                continue
    return sha(json.dumps(sorted(manifest), sort_keys=True).encode())


def reusable_snapshot(state, profile, source, from_date):
    try:
        directory, index = load_index(state)
        if (index.get("profile") != str(profile)
                or index.get("summary", {}).get("scope_from_date") != from_date):
            return None
        source_stat = fingerprint(source)
        if index.get("source_fingerprint") != source_stat:
            return None
        assets = asset_digest(profile)
        if index.get("asset_manifest_sha256") != assets:
            return None
        # Keep the source stability check around the metadata-only fast path.
        if fingerprint(source) != source_stat:
            return None
    except (OSError, ValueError, KeyError, TypeError):
        return None
    publish_current(state, index["snapshot"])
    return {"snapshot": index["snapshot"], "reused": True, **index["summary"]}


def freeze_body(directory, profile, header, references, inline=(), body_objects=0,
                body_reference=None, unparsed_objects=0):
    best_candidate, errors = None, []

    def candidate(data, metadata):
        nonlocal best_candidate
        text, encoding = decode_body(data)
        preview = folded(header.get("preview"))
        verified = bool(preview) and preview[:60] in folded(text[:12000])
        if preview and not verified:
            source = {"path": metadata["source_path"]} if "source_path" in metadata else metadata
            errors.append({**source, "status": "preview_mismatch"})
            return
        value = (len(text), data, text,
                 {**metadata, "status": "cached_html", "format": encoding,
                  "text_chars": len(text), "preview_verified": verified})
        if best_candidate is None or value[0] > best_candidate[0]:
            best_candidate = value

    for reference in sorted(references):
        try:
            data, source_stat = stable_read(source_path(profile, reference), MAX_BODY)
            candidate(data, {"storage": "external_file", "source_path": reference,
                            "source_fingerprint": source_stat, "association": references[reference]})
        except FileNotFoundError:
            errors.append({"path": reference, "status": "not_cached"})
        except (ValueError, UnicodeError, OSError, EOFError, zlib.error) as error:
            errors.append({"path": reference, "status": "unavailable", "error": str(error)})
    for body in inline:
        metadata = {"storage": "hxstore_inline", "association": body["evidence"]}
        temporary = body.get("data_path")
        try:
            if temporary is not None:
                temporary = Path(temporary)
                source_stat = fingerprint(temporary)
                if source_stat["size"] > MAX_BODY:
                    raise ValueError("inline body exceeds prototype limit")
                with temporary.open("rb") as handle:
                    data = handle.read(source_stat["size"])
                if len(data) != source_stat["size"] or fingerprint(temporary) != source_stat:
                    raise ValueError("inline body changed during read")
            else:
                data = body.get("data")
            if not isinstance(data, (bytes, bytearray, memoryview)):
                raise ValueError("inline body payload is missing")
            expected_digest = body.get("data_sha256")
            if expected_digest is not None and sha(data) != expected_digest:
                raise ValueError("inline body checksum mismatch")
            if len(data) > MAX_BODY:
                raise ValueError("inline body exceeds prototype limit")
            candidate(data, metadata)
        except (ValueError, UnicodeError, OSError, EOFError, zlib.error) as error:
            errors.append({**metadata, "status": "unavailable", "error": str(error)})
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
    diagnosis = {"body_object_frames": body_objects, "external_references": len(references),
                 "inline_candidates": len(inline), "unparsed_body_objects": unparsed_objects}
    if best_candidate is None:
        if errors:
            status = "unavailable" if any(e["status"] != "not_cached" for e in errors) else "not_cached"
            code = "body_candidates_rejected" if status == "unavailable" else "referenced_file_absent"
            if status == "not_cached" and unparsed_objects:
                status, code = "unavailable", "body_object_layout_or_owner_unresolved"
        elif body_objects:
            status, code = "unavailable", "body_object_layout_or_owner_unresolved"
        elif body_reference:
            status, code = "not_cached", "referenced_body_object_absent"
            diagnosis["absence_scope"] = "verified_hxstore_frames_and_recognized_file_registry"
        else:
            status, code = "not_linked", "no_reliable_body_reference"
        return {"status": status, "diagnosis": {**diagnosis, "code": code}, "checks": errors,
                "body_reference": body_reference}
    # Multiple paths can be full/unique body variants of the same typed owner.
    # Preserve the longest cached representation only after preview validation.
    _, data, text, metadata = best_candidate
    metadata.update(sha256=frozen_payload(directory, data), bytes=len(data),
                    checks=errors, diagnosis={**diagnosis, "code": f"{metadata['storage']}_recovered"})
    metadata.update(frozen_text_payload(directory, text))
    return metadata


def day_start(value):
    return int(datetime.combine(date.fromisoformat(value), time(), ZONE).timestamp())


def discover_profile(value):
    if value:
        profile = Path(value).expanduser().resolve()
    else:
        root = Path.home() / "Library/Group Containers/UBF8T346G9.Office/Outlook"
        matches = list(root.glob("* Profiles/*/HxStore.hxd"))
        if len(matches) != 1:
            raise ValueError("specify --profile: expected exactly one HxStore profile")
        profile = matches[0].parent.resolve()
    if not (profile / "HxStore.hxd").is_file():
        raise ValueError("profile does not contain HxStore.hxd")
    return profile


def make_snapshot(args):
    started = clock.monotonic()
    profile = discover_profile(args.profile)
    state = state_root(args.state, profile, create=True)
    source = profile / "HxStore.hxd"
    reused = reusable_snapshot(state, profile, source, args.from_date)
    if reused is not None:
        return reused
    build = Path(tempfile.mkdtemp(prefix=f".building-{os.getpid()}-", dir=state))
    data = None
    source_map = None
    source_copy = None
    try:
        for _ in range(3):
            try:
                source_stat = fingerprint(source)
                if source_stat["size"] <= MAX_IN_MEMORY_SOURCE:
                    data, source_stat = stable_read(source, MAX_IN_MEMORY_SOURCE)
                    digest = sha(data)
                    # Verify rewrites in bounded chunks; keep only one small cache in memory.
                    verification = hashlib.sha256()
                    remaining = source_stat["size"]
                    with source.open("rb") as handle:
                        while remaining:
                            chunk = handle.read(min(SOURCE_COPY_CHUNK, remaining))
                            if not chunk:
                                break
                            verification.update(chunk)
                            remaining -= len(chunk)
                        extra = handle.read(1)
                    if (not remaining and not extra and source_stat == fingerprint(source)
                            and digest == verification.hexdigest()):
                        break
                else:
                    source_copy, source_stat, digest = copy_stable_source(source, build)
                    with source_copy.open("rb") as handle:
                        source_map = mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ)
                    data = source_map
                    break
            except ValueError as error:
                if "insufficient scratch space" in str(error):
                    raise
            except OSError:
                pass
            if source_map is not None:
                source_map.close()
                source_map = None
            if source_copy is not None:
                source_copy.unlink(missing_ok=True)
                source_copy = None
            data = None
        if data is None:
            raise ValueError("Outlook cache kept changing; no snapshot published")
        assets = asset_digest(profile)
        snapshot_id = sha(f"{VERSION}|{profile}|{digest}|{assets}|{args.from_date}".encode())[:24]
        final = state / "snapshots" / snapshot_id
        if final.is_symlink():
            raise ValueError("snapshot directory must not be a symbolic link")
        if final.exists():
            publish_current(state, snapshot_id)
            return {"snapshot": snapshot_id, "reused": True, **json.loads((final / "index.json").read_text())["summary"]}
        headers, body_paths, exact_body_paths, attachments = (
            defaultdict(list), defaultdict(dict), defaultdict(dict), defaultdict(list)
        )
        body_objects, inline_bodies, header_body_targets = defaultdict(int), defaultdict(list), defaultdict(set)
        unparsed_bodies = defaultdict(int)
        header_body_evidence = {}
        file_records = []
        stats, frame_counts = {}, defaultdict(int)
        blocks = iter_blocks(data, stats)
        try:
            for block in blocks:
                for frame in iter_frames(block):
                    frame_counts[str(frame.kind)] += 1
                    if frame.kind == 201:
                        header = parse_header(frame)
                        if header:
                            headers[frame.key].append(header)
                            targets = typed_references(frame, 202)
                            header_body_targets[frame.key].update(targets)
                            if len(targets) == 1:
                                target_key = next(iter(targets))
                                header_body_evidence[frame.key] = {**header["evidence"], "kind": 202,
                                                                   "target_key": target_key}
                    elif frame.kind == 202:
                        body_objects[frame.key] += 1
                        inline_body = parse_inline_body(frame)
                        if inline_body:
                            inline_data = inline_body.pop("data")
                            with tempfile.NamedTemporaryFile(
                                    prefix=".inline-", dir=build, delete=False) as handle:
                                handle.write(inline_data)
                                inline_body["data_path"] = handle.name
                            inline_body["data_sha256"] = sha(inline_data)
                            inline_bodies[frame.key].append(inline_body)
                            del inline_data
                        references = [ref for ref in file_references(frame) if "/EFMData/" in ref["path"]]
                        if not inline_body and not references:
                            unparsed_bodies[frame.key] += 1
                        for ref in references:
                            body_paths[frame.key][ref["path"]] = {
                                "method": "same_typed_message_key", "key": frame.key,
                                "block": frame.block, "frame_offset": frame.offset, "kind": 202}
                    elif frame.kind == 362:
                        attachment = parse_attachment(frame)
                        if attachment:
                            attachments[frame.key].append(attachment)
                    elif frame.kind == 247:
                        entries = entry_ids(frame.data)
                        references = file_references(frame)
                        if len(entries) == 1 and references:
                            file_records.append((frame.key, next(iter(entries)), references,
                                                 {"block": frame.block, "frame_offset": frame.offset,
                                                  "association": "exact_exchange_entry_id"}))
        finally:
            close_blocks = getattr(blocks, "close", None)
            if close_blocks is not None:
                close_blocks()
            if source_map is not None:
                source_map.close()
                source_map = None
            if source_copy is not None:
                source_copy.unlink(missing_ok=True)
                source_copy = None
        entry_owners = defaultdict(set)
        for key, versions in headers.items():
            for header in versions:
                for entry in header["entry_ids"]:
                    entry_owners[entry].add(key)
        for file_key, entry, references, evidence in file_records:
            if len(entry_owners[entry]) != 1:
                continue
            owner = next(iter(entry_owners[entry]))
            for reference in references:
                if "/EFMData/" in reference["path"]:
                    exact_body_paths[owner].setdefault(reference["path"], {
                        **evidence, "kind": 247, "file_key": file_key,
                        "entry_id_sha256": sha(bytes.fromhex(entry)), "owner_key": owner})
        if not stats.get("verified"):
            raise ValueError("no verified HxStore blocks; cache is unsupported or damaged")
        if not headers:
            raise ValueError("no recognizable mail headers; cannot infer an empty mailbox")
        cutoff, chosen, conflicts, unresolved_times, time_conflicts = day_start(args.from_date), {}, 0, 0, 0
        for key, versions in headers.items():
            identities = {row["message_id"] for row in versions}
            subjects = {normalize_subject(row["subject"]) for row in versions if row["subject"]}
            senders = {row["sender"].casefold() for row in versions if row["sender"]}
            if len(identities) != 1 or len(subjects) > 1 or len(senders) > 1:
                conflicts += 1
                continue
            sent_times = {version["sent_unix"] for version in versions if version["sent_unix"] is not None}
            if len(sent_times) > 1:
                time_conflicts += 1
                continue
            row = dict(max(versions, key=lambda row: row["latest_tick_unix"] or 0))
            if not row["subject"]:
                paired = [version for version in versions if version["subject"]]
                if paired:
                    source_version = max(paired, key=lambda version: version["latest_tick_unix"] or 0)
                    row.update(subject=source_version["subject"], subject_status="paired_older_same_identity",
                               subject_evidence=source_version["evidence"])
            if row["sent_unix"] is None:
                if len(sent_times) == 1:
                    row.update(sent_unix=next(iter(sent_times)), time_method="header_field_271_same_identity")
                else:
                    unresolved_times += 1
                    continue
            if row["sent_unix"] is None or row["sent_unix"] < cutoff:
                continue
            row["attachments"] = []
            row["mailbox_state"] = "unverified"
            targets = header_body_targets[key]
            body_key = key
            reference = None
            if exact_body_paths[key]:
                # EntryId ownership is stronger than a frame-key association;
                # do not let a weaker candidate win by being longer.
                references = exact_body_paths[key]
            elif len(targets) == 1:
                body_key = next(iter(targets))
                references = body_paths[body_key]
                reference = header_body_evidence[key]
            else:
                references = {}
            row["body"] = freeze_body(
                build, profile, row, references,
                inline_bodies[body_key] if reference else (),
                body_objects[body_key], reference, unparsed_bodies[body_key]
            )
            chosen[key] = row
        attachment_conflicts = 0
        for key, versions in attachments.items():
            resolved = []
            for attachment in versions:
                explicit_owners = set().union(*(entry_owners[entry] for entry in attachment["entry_ids"]))
                if len(explicit_owners) > 1:
                    attachment_conflicts += 1
                    continue
                if len(explicit_owners) == 1:
                    owner = next(iter(explicit_owners))
                    if attachment["owner"] is not None and attachment["owner"] != owner:
                        attachment_conflicts += 1
                        continue
                    attachment = {**attachment, "owner": owner,
                                  "evidence": {**attachment["evidence"],
                                               "association": "exact_exchange_entry_id"}}
                if attachment["owner"] is not None:
                    resolved.append(attachment)
            versions = resolved
            if not versions:
                continue
            links = {(row["owner"], row["path"]) for row in versions}
            if len(links) != 1:
                attachment_conflicts += 1
                continue
            row = versions[-1]
            if row["owner"] not in chosen:
                continue
            row["id"] = f"att_{key:x}"
            row["cache_name"] = PurePosixPath(row["path"]).name
            try:
                path = source_path(profile, row["path"])
                row["fingerprint"] = fingerprint(path)
                row["status"] = "cached_at_snapshot"
            except FileNotFoundError:
                row["status"] = "not_cached"
            chosen[row["owner"]]["attachments"].append(row)
        # Registry-only attachment references retain their on-disk cache name;
        # do not invent the original display name when no type-362 row exists.
        known_paths = {a["path"] for row in chosen.values() for a in row["attachments"]}
        for file_key, entry, references, evidence in file_records:
            owners = entry_owners[entry]
            if len(owners) != 1 or next(iter(owners)) not in chosen:
                continue
            owner = next(iter(owners))
            for reference in references:
                if "/Attachments/" not in reference["path"] or reference["path"] in known_paths:
                    continue
                path_id = sha(reference["path"].encode())[:12]
                row = {"id": f"file_{file_key:x}_{path_id}", "key": file_key, "owner": owner,
                       "path": reference["path"], "name": None,
                       "cache_name": PurePosixPath(reference["path"]).name, "evidence": evidence}
                try:
                    row["fingerprint"] = fingerprint(source_path(profile, row["path"]))
                    row["status"] = "cached_at_snapshot"
                except FileNotFoundError:
                    row["status"] = "not_cached"
                chosen[owner]["attachments"].append(row)
                known_paths.add(row["path"])
        for versions in inline_bodies.values():
            for body in versions:
                temporary = body.get("data_path")
                if temporary is not None:
                    Path(temporary).unlink(missing_ok=True)
        items = sorted(chosen.values(), key=lambda row: (-row["sent_unix"], row["id"]))
        summary = {"scope_from_date": args.from_date, "timezone": "Asia/Singapore",
                   "messages": len(items), "paired_subjects": sum(bool(r["subject"]) for r in items),
                   "unresolved_times_excluded": unresolved_times,
                   "time_conflicts_excluded": time_conflicts,
                   "unparsed_header_frames": frame_counts["201"] - sum(len(v) for v in headers.values()),
                   "cached_bodies": sum(r["body"]["status"] == "cached_html" for r in items),
                   "body_statuses": {status: sum(r["body"]["status"] == status for r in items)
                                     for status in ("cached_html", "not_linked", "not_cached", "unavailable")},
                   "body_diagnoses": {code: sum(r["body"]["diagnosis"]["code"] == code for r in items)
                                      for code in sorted({r["body"]["diagnosis"]["code"] for r in items})},
                   "linked_attachments": sum(len(r["attachments"]) for r in items),
                   "identity_conflicts_excluded": conflicts, "attachment_conflicts_excluded": attachment_conflicts,
                   "container": stats, "frames": dict(frame_counts),
                   "coverage": "recoverable_cached_records; mailbox membership and completeness unverified"}
        index = {"schema_version": VERSION, "snapshot": snapshot_id, "profile": str(profile),
                 "source_sha256": digest, "source_fingerprint": source_stat, "asset_manifest_sha256": assets,
                 "created_at": datetime.now(timezone.utc).isoformat(), "summary": summary, "items": items}
        (build / "index.json").write_text(json.dumps(index, ensure_ascii=False))
        build.rename(final)
        publish_current(state, snapshot_id)
        return {"snapshot": snapshot_id, "elapsed_seconds": round(clock.monotonic() - started, 3), **summary}
    finally:
        if source_map is not None:
            source_map.close()
        if source_copy is not None:
            source_copy.unlink(missing_ok=True)
        if build.exists():
            shutil.rmtree(build)


def load_index(state_arg):
    state = state_root(state_arg)
    snapshot = (state / "CURRENT").read_text().strip()
    if len(snapshot) != 24 or any(c not in "0123456789abcdef" for c in snapshot):
        raise ValueError("invalid snapshot pointer")
    directory = state / "snapshots" / snapshot
    _private_directory(directory, "snapshot directory")
    for name in ("payloads", "textpayloads"):
        entry = directory / name
        if entry.exists() or entry.is_symlink():
            _private_directory(entry, f"snapshot {name} directory")
    index = json.loads((directory / "index.json").read_text())
    if index["schema_version"] != VERSION or index["snapshot"] != snapshot:
        raise ValueError("unsupported or mismatched prototype index")
    outside_source(state, index["profile"])
    return directory, index


def find_message(index, message_id):
    row = next((item for item in index["items"] if item["id"] == message_id), None)
    if row is None:
        raise ValueError("message id is absent from this snapshot")
    return row


def received_sort_key(row):
    value = row.get("received_unix")
    return (value is None, -value if value is not None else 0, row["id"])


def metadata(row):
    sent = (datetime(1970, 1, 1, tzinfo=timezone.utc)
            + timedelta(milliseconds=round(row["sent_unix"] * 1000))).astimezone(ZONE)
    received = row.get("received_unix")
    received_at = ((datetime(1970, 1, 1, tzinfo=timezone.utc)
                    + timedelta(milliseconds=round(received * 1000))).astimezone(ZONE).isoformat()
                   if received is not None else None)
    return {**{k: row[k] for k in ("id", "message_id", "subject", "subject_status", "sender", "sender_status", "sender_name", "preview", "mailbox_state")},
            "preview_method": row.get("preview_method", "after_class_candidate"),
            "sent_at": sent.isoformat(), "received_at": received_at,
            "received_time_status": row.get("received_time_status", "missing"),
            "received_time_method": row.get("received_time_method"),
            "time_method": row["time_method"],
            "body_status": row["body"]["status"], "body_diagnosis": row["body"]["diagnosis"]["code"],
            "attachment_count": len(row["attachments"]),
            "attachments_complete": False}


def search(index, args):
    if args.from_date and day_start(args.from_date) < day_start(index["summary"]["scope_from_date"]):
        raise ValueError("requested date predates snapshot scope; build a wider snapshot")
    start = day_start(args.from_date) if args.from_date else None
    end = (day_start((date.fromisoformat(args.to_date) + timedelta(days=1)).isoformat())
           if args.to_date and args.to_date != "9999-12-31" else None)
    if end is not None and end <= day_start(index["summary"]["scope_from_date"]):
        raise ValueError("requested end date predates snapshot scope; build a wider snapshot")
    if start is not None and end is not None and start >= end:
        raise ValueError("from-date must not exceed to-date")
    query, matched = args.query.casefold(), []
    unknown_times = sum(row.get("received_unix") is None for row in index["items"])
    for row in sorted(index["items"], key=received_sort_key):
        received = row.get("received_unix")
        if start is not None or end is not None:
            if received is None or (start is not None and received < start) or (end is not None and received >= end):
                continue
        fields = [args.field] if args.field != "all" else ["subject", "sender", "sender_name", "preview"]
        hits = [field for field in fields if query in (row.get(field) or "").casefold()]
        if query and not hits:
            continue
        matched.append({**metadata(row), "matched_fields": hits if query else []})
    signature = sha(json.dumps([index["snapshot"], args.query, args.field, args.from_date, args.to_date, args.limit]).encode())[:16]
    offset = 0
    if args.cursor:
        cursor_signature, offset_value = args.cursor.split(":")
        if cursor_signature != signature:
            raise ValueError("cursor belongs to another snapshot or query")
        offset = int(offset_value)
        if offset < 0 or offset > len(matched):
            raise ValueError("invalid cursor offset")
    items = matched[offset:offset + args.limit]
    next_offset = offset + len(items)
    return {"snapshot": index["snapshot"], "items": items, "total_matched": len(matched),
            "scope_from_date": index["summary"]["scope_from_date"], "captured_at": index["created_at"],
            "date_basis": "received_at", "received_time_unknown": unknown_times,
            "date_filter_incomplete": bool(unknown_times and (start is not None or end is not None)),
            "complete": next_offset >= len(matched),
            "next_cursor": f"{signature}:{next_offset}" if next_offset < len(matched) else None,
            "coverage": index["summary"]["coverage"]}


def _remember(cache, key):
    cache[key] = True
    cache.move_to_end(key)
    while len(cache) > VERIFIED_CACHE_LIMIT:
        cache.popitem(last=False)


def _verified_raw_payload(directory, body):
    expected = body.get("sha256")
    expected_bytes = body.get("bytes")
    if (not isinstance(expected, str) or not re.fullmatch(r"[a-f0-9]{64}", expected)
            or type(expected_bytes) is not int or expected_bytes < 0):
        raise ValueError("frozen body checksum metadata is invalid")
    root = directory / "payloads"
    if root.is_symlink():
        raise ValueError("frozen body payload directory must not be a symbolic link")
    path = root / expected
    if path.is_symlink() or not path.is_file():
        raise ValueError("frozen body payload is missing")
    current = fingerprint(path)
    if current["size"] != expected_bytes:
        raise ValueError("frozen body checksum mismatch")
    key = (str(path), expected, tuple(current.items()))
    if key not in _verified_raw_payloads:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            while True:
                chunk = handle.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
        if digest.hexdigest() != expected:
            raise ValueError("frozen body checksum mismatch")
        _remember(_verified_raw_payloads, key)
    else:
        _verified_raw_payloads.move_to_end(key)
    return path


def _text_payload_info(directory, body):
    expected = body.get("text_sha256")
    text_bytes = body.get("text_bytes")
    text_units = body.get("text_utf16_units")
    chunk_bytes = body.get("text_chunk_bytes")
    chunks = body.get("text_chunk_sha256")
    text_chars = body.get("text_chars")
    if (not isinstance(expected, str) or not re.fullmatch(r"[a-f0-9]{64}", expected)
            or type(text_bytes) is not int or text_bytes < 0
            or type(text_units) is not int or text_units < 0 or text_bytes != text_units * 2
            or type(text_chars) is not int or text_chars < 0
            or type(chunk_bytes) is not int or chunk_bytes <= 0 or chunk_bytes % 2
            or not isinstance(chunks, list)
            or len(chunks) != (text_bytes + chunk_bytes - 1) // chunk_bytes
            or any(not isinstance(value, str) or not re.fullmatch(r"[a-f0-9]{64}", value)
                   for value in chunks)):
        raise ValueError("frozen text payload checksum metadata is invalid")
    root = directory / "textpayloads"
    if root.is_symlink():
        raise ValueError("frozen text payload directory must not be a symbolic link")
    path = root / expected
    if path.is_symlink() or not path.is_file():
        raise ValueError("frozen text payload is missing")
    current = fingerprint(path)
    if current["size"] != text_bytes:
        raise ValueError("frozen text payload checksum mismatch")
    return path, current, text_bytes, text_units, text_chars, chunk_bytes, chunks


def _verified_text_range(directory, body, start, end):
    path, current, text_bytes, text_units, _, chunk_bytes, chunks = _text_payload_info(directory, body)
    if type(start) is not int or type(end) is not int or not 0 <= start <= end <= text_units:
        raise ValueError("invalid frozen text payload range")
    first = (2 * start) // chunk_bytes
    last = (2 * end - 1) // chunk_bytes if start < end else first - 1
    if first <= last:
        with path.open("rb") as handle:
            for number in range(first, last + 1):
                expected_size = min(chunk_bytes, text_bytes - number * chunk_bytes)
                key = (str(path), tuple(current.items()), number, chunks[number])
                if key in _verified_text_chunks:
                    _verified_text_chunks.move_to_end(key)
                    continue
                handle.seek(number * chunk_bytes)
                data = handle.read(expected_size)
                if len(data) != expected_size or sha(data) != chunks[number]:
                    raise ValueError("frozen text payload checksum mismatch")
                _remember(_verified_text_chunks, key)
    if fingerprint(path) != current:
        raise ValueError("frozen text payload changed during read")
    return path, current, text_bytes, text_units, chunk_bytes


def _read_text_units(directory, body, start, end):
    path, current, text_bytes, _, chunk_bytes = _verified_text_range(directory, body, start, end)
    byte_start, byte_end = start * 2, end * 2
    with path.open("rb") as handle:
        handle.seek(byte_start)
        data = handle.read(byte_end - byte_start)
    if len(data) != byte_end - byte_start or fingerprint(path) != current:
        raise ValueError("frozen text payload changed during read")
    return data.decode("utf-16-le", errors="surrogatepass")


def _read_python_prefix(directory, body, maximum):
    _, _, _, total_units, total_chars, _, _ = _text_payload_info(directory, body)
    if type(maximum) is not int or maximum < 1:
        raise ValueError("body character budget must be positive")
    end = total_units if total_chars <= maximum else min(total_units, maximum * 2)
    text = _read_text_units(directory, body, 0, end)
    return (text if total_chars <= maximum else text[:maximum], total_chars > maximum)


def read_body_page(directory, index, message_id, offset, maximum):
    """Read a UTF-16-unit page for the common worker without rebuilding HTML."""
    row = find_message(index, message_id)
    body = row["body"]
    if body["status"] != "cached_html":
        return None
    _verified_raw_payload(directory, body)
    _, _, _, total_units, _, _, _ = _text_payload_info(directory, body)
    start = min(offset, total_units)
    end = min(start + maximum, total_units)
    return {"text": _read_text_units(directory, body, start, end),
            "total_chars": total_units, "offset": start,
            "next_offset": end if end < total_units else None,
            "truncated": end < total_units}


def read_message(directory, index, message_id, maximum):
    row = find_message(index, message_id)
    result = {**metadata(row), "snapshot": index["snapshot"], "body": {**row["body"]},
              "attachments": row["attachments"], "evidence": row["evidence"],
              "subject_evidence": row.get("subject_evidence", row["evidence"])}
    if row["body"]["status"] == "cached_html":
        _verified_raw_payload(directory, row["body"])
        text, truncated = _read_python_prefix(directory, row["body"], maximum)
        result["body"].update(text=text, truncated=truncated)
    return result


def read_attachment(directory, index, message_id, attachment_id, output):
    message = find_message(index, message_id)
    attachment = next((row for row in message["attachments"] if row["id"] == attachment_id), None)
    if attachment is None:
        raise ValueError("attachment does not belong to this message")
    profile = Path(index["profile"])
    destination = outside_source(output, profile)
    if destination.exists():
        raise ValueError("output already exists; refusing to overwrite")
    path = source_path(profile, attachment["path"])
    if attachment["status"] != "cached_at_snapshot" or fingerprint(path) != attachment["fingerprint"]:
        raise ValueError("attachment is missing or changed since snapshot; build a fresh snapshot")
    data, _ = stable_read(path, MAX_ATTACHMENT)
    if fingerprint(path) != attachment["fingerprint"]:
        raise ValueError("attachment changed since snapshot")
    with destination.open("xb") as handle:
        handle.write(data)
    return {"snapshot": index["snapshot"], "message_id": message_id, "attachment_id": attachment_id,
            "name": attachment["name"], "bytes": len(data), "sha256": sha(data), "output": str(destination)}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", default=str(DEFAULT_STATE))
    commands = parser.add_subparsers(dest="command", required=True)
    snapshot = commands.add_parser("snapshot", help="capture metadata and recoverable recent bodies")
    snapshot.add_argument("--profile")
    snapshot.add_argument("--from-date", default=(datetime.now(ZONE).date() - timedelta(days=7)).isoformat())
    commands.add_parser("status")
    query = commands.add_parser("search")
    query.add_argument("--query", default="")
    query.add_argument("--field", choices=("all", "subject", "sender", "preview"), default="subject")
    query.add_argument("--from-date")
    query.add_argument("--to-date")
    query.add_argument("--limit", type=int, default=20)
    query.add_argument("--cursor")
    read = commands.add_parser("read")
    read.add_argument("id")
    read.add_argument("--body-max-chars", type=int, default=20000)
    attachments = commands.add_parser("attachments")
    attachments.add_argument("id")
    attachment = commands.add_parser("attachment")
    attachment.add_argument("id")
    attachment.add_argument("attachment_id")
    attachment.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "snapshot":
            result = make_snapshot(args)
        else:
            directory, index = load_index(args.state)
            if args.command == "status":
                result = {"snapshot": index["snapshot"], "created_at": index["created_at"], **index["summary"]}
            elif args.command == "search":
                if not 1 <= args.limit <= 100:
                    raise ValueError("limit must be 1..100")
                result = search(index, args)
            elif args.command == "read":
                if not 1 <= args.body_max_chars <= 2000000:
                    raise ValueError("body-max-chars must be 1..2000000")
                result = read_message(directory, index, args.id, args.body_max_chars)
            elif args.command == "attachments":
                result = {"snapshot": index["snapshot"], "items": find_message(index, args.id)["attachments"]}
            else:
                result = read_attachment(directory, index, args.id, args.attachment_id, args.out)
        emit({"ok": True, "data": result})
        return 0
    except (ValueError, OSError, KeyError, EOFError) as error:
        emit({"ok": False, "error": {"type": type(error).__name__, "message": str(error)}})
        return 1


if __name__ == "__main__":
    sys.exit(main())
