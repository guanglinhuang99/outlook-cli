#!/usr/bin/env python3
"""Install the local Outlook MCP server without touching client settings."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shlex
import stat
import sys
import tempfile


INSTALLER_MARKER = "# outlook-local-mcp managed wrapper v1"
SERVER_NAME = "outlook-local"
DEFAULT_PREFIX = Path.home() / ".local"
REQUIRED_FILES = (
    "hxcontainer.py",
    "hxrecords.py",
    "outlook_local.py",
    "osa_scope.py",
    "outlook_service.py",
    "fulltext_index.py",
    "mcp_server.py",
)
OPTIONAL_FILES = ("hxfolders.py",)


class InstallError(ValueError):
    """Raised when installation would follow or overwrite an unsafe path."""


def _reject_symlink(path: Path, label: str) -> None:
    if path.is_symlink():
        raise InstallError(f"{label} must not be a symbolic link: {path}")


def _verify_trusted_directory(path: Path, label: str, *, allow_sticky: bool) -> None:
    try:
        info = path.lstat()
    except FileNotFoundError:
        return
    if not stat.S_ISDIR(info.st_mode):
        raise InstallError(f"{label} is not a directory: {path}")
    uid = os.getuid() if hasattr(os, "getuid") else None
    if uid is not None and info.st_uid not in (0, uid):
        raise InstallError(f"{label} is not owned by the current user or root: {path}")
    if uid is not None and info.st_mode & 0o022:
        if not (allow_sticky and info.st_mode & stat.S_ISVTX and info.st_uid in (0, uid)):
            raise InstallError(f"{label} must not be group/world writable: {path}")


def _verify_trusted_path(path: Path, label: str) -> None:
    current = path
    first = True
    while True:
        _verify_trusted_directory(current, label if first else "ancestor directory", allow_sticky=not first)
        parent = current.parent
        if parent == current:
            return
        current = parent
        first = False


def _ensure_directory(path: Path, label: str) -> None:
    _verify_trusted_path(path, label)
    _reject_symlink(path, label)
    if path.exists():
        if not path.is_dir():
            raise InstallError(f"{label} is not a directory: {path}")
        _verify_trusted_path(path, label)
        return
    path.mkdir(mode=0o700)
    _verify_trusted_path(path, label)


def _source_paths(source_root: Path) -> list[tuple[str, Path]]:
    paths: list[tuple[str, Path]] = []
    for name in (*REQUIRED_FILES, *OPTIONAL_FILES):
        path = source_root / name
        if not path.exists() and not path.is_symlink():
            if name in OPTIONAL_FILES:
                continue
            raise InstallError(f"required source file is missing: {path}")
        _reject_symlink(path, "source file")
        if not path.is_file():
            raise InstallError(f"source file is not a regular file: {path}")
        paths.append((name, path))
    return paths


def _read_sources(source_root: Path) -> list[tuple[str, bytes]]:
    values = []
    for name, path in _source_paths(source_root):
        try:
            data = path.read_bytes()
        except OSError as error:
            raise InstallError(f"cannot read source file {path}: {error}") from error
        values.append((name, data))
    return values


def source_digest(source_root: Path) -> tuple[str, list[tuple[str, bytes]]]:
    """Return a deterministic digest and the exact bounded source payload."""

    values = _read_sources(source_root)
    digest = hashlib.sha256()
    for name, data in values:
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(len(data).to_bytes(8, "big"))
        digest.update(data)
    return digest.hexdigest(), values


def _safe_prefix(value: str | Path) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    _reject_symlink(path, "prefix")
    return path.resolve()


def _write_new_or_verify(path: Path, data: bytes, label: str) -> None:
    _reject_symlink(path, label)
    if path.exists():
        if not path.is_file() or path.read_bytes() != data:
            raise InstallError(f"refusing to overwrite existing {label}: {path}")
        path.chmod(0o600)
        return
    try:
        with path.open("xb") as handle:
            handle.write(data)
        path.chmod(0o600)
    except FileExistsError:
        _reject_symlink(path, label)
        if not path.is_file() or path.read_bytes() != data:
            raise InstallError(f"refusing to overwrite existing {label}: {path}")
        path.chmod(0o600)


def _wrapper_text(python_executable: Path, entrypoint: Path) -> bytes:
    return (
        "#!/bin/sh\n"
        f"{INSTALLER_MARKER}\n"
        f"exec {shlex.quote(str(python_executable))} {shlex.quote(str(entrypoint))} \"$@\"\n"
    ).encode("utf-8")


def _install_wrapper(path: Path, data: bytes) -> None:
    _reject_symlink(path, "wrapper")
    if path.exists():
        if not path.is_file():
            raise InstallError(f"wrapper is not a regular file: {path}")
        current = path.read_bytes()
        if INSTALLER_MARKER.encode("utf-8") not in current.splitlines():
            raise InstallError(f"refusing to overwrite unowned wrapper: {path}")
        if current == data:
            path.chmod(0o700)
            return
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=".outlook-local-mcp-", dir=path.parent
        )
        os.close(descriptor)
        temporary = Path(temporary_name)
        try:
            temporary.write_bytes(data)
            temporary.chmod(0o700)
            temporary.replace(path)
        finally:
            temporary.unlink(missing_ok=True)
        return
    try:
        with path.open("xb") as handle:
            handle.write(data)
        path.chmod(0o700)
    except FileExistsError:
        _reject_symlink(path, "wrapper")
        if not path.is_file() or INSTALLER_MARKER.encode("utf-8") not in path.read_bytes().splitlines():
            raise InstallError(f"refusing to overwrite unowned wrapper: {path}")
        _install_wrapper(path, data)


def install(source_root: Path | None = None, prefix: Path | str = DEFAULT_PREFIX,
            python_executable: str | Path | None = None) -> dict[str, object]:
    """Install source files into a content-addressed private version directory."""

    source_root = (Path(__file__).parent if source_root is None
                   else Path(source_root).expanduser())
    if not source_root.is_absolute():
        source_root = Path.cwd() / source_root
    _reject_symlink(source_root, "source directory")
    source_root = source_root.resolve()
    if not source_root.is_dir():
        raise InstallError(f"source directory is not a directory: {source_root}")
    digest, values = source_digest(source_root)
    prefix = _safe_prefix(prefix)
    _ensure_directory(prefix, "prefix")
    share = prefix / "share"
    package_root = share / "outlook-local-mcp"
    bin_root = prefix / "bin"
    _ensure_directory(share, "share directory")
    _ensure_directory(package_root, "installation directory")
    _ensure_directory(bin_root, "bin directory")

    version_root = package_root / digest
    _ensure_directory(version_root, "version directory")
    for name, data in values:
        _write_new_or_verify(version_root / name, data, "installed source file")

    interpreter = Path(
        sys.executable if python_executable is None else python_executable
    ).expanduser()
    if not interpreter.is_absolute():
        interpreter = Path.cwd() / interpreter
    if not interpreter.is_file():
        raise InstallError(f"interpreter is not a regular file: {interpreter}")
    if not os.access(interpreter, os.X_OK):
        raise InstallError(f"interpreter is not executable: {interpreter}")
    wrapper = bin_root / "outlook-local-mcp"
    wrapper_data = _wrapper_text(interpreter, version_root / "mcp_server.py")
    _install_wrapper(wrapper, wrapper_data)
    return {
        "digest": digest,
        "version_dir": str(version_root),
        "wrapper": str(wrapper),
        "files": [name for name, _ in values],
    }


def mcp_servers_config(wrapper: str | Path, profile: str | Path | None = None) -> dict[str, object]:
    """Return a client-neutral MCP server configuration object."""

    args = [] if profile is None else ["--profile", str(Path(profile).expanduser().resolve())]
    return {"mcpServers": {SERVER_NAME: {"command": str(wrapper), "args": args}}}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--prefix", default=str(DEFAULT_PREFIX),
                        help="installation prefix; defaults to ~/.local")
    parser.add_argument("--profile", help="Outlook profile for the generated MCP configuration")
    args = parser.parse_args(argv)
    try:
        result = install(prefix=args.prefix)
    except (InstallError, OSError) as error:
        print(f"install_mcp: {error}", file=sys.stderr)
        return 1
    print(json.dumps(mcp_servers_config(result["wrapper"], args.profile), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
