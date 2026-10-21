import contextlib
import io
import json
import os
from pathlib import Path
import stat
import tempfile
import unittest


MODULE_DIR = Path(__file__).parent
import sys

sys.path.insert(0, str(MODULE_DIR))
import install_mcp  # noqa: E402


class InstallerTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        self.root = Path(self.tempdir.name)
        self.source = self.root / "source"
        self.source.mkdir()
        self._write_source()

    def _write_source(self, *, optional=False):
        for name in install_mcp.REQUIRED_FILES:
            (self.source / name).write_bytes((f"# fixture {name}\n").encode())
        if optional:
            (self.source / install_mcp.OPTIONAL_FILES[0]).write_bytes(b"# fixture optional\n")

    def test_install_copies_versioned_sources_and_managed_wrapper(self):
        prefix = self.root / "prefix"
        (prefix / "keep.txt").parent.mkdir(parents=True)
        (prefix / "keep.txt").write_text("leave me", encoding="utf-8")
        self._write_source(optional=True)

        result = install_mcp.install(
            source_root=self.source,
            prefix=prefix,
            python_executable=sys.executable,
        )

        digest, values = install_mcp.source_digest(self.source)
        prefix = prefix.resolve()
        version_dir = prefix / "share" / "outlook-local-mcp" / digest
        wrapper = prefix / "bin" / "outlook-local-mcp"
        self.assertEqual(result["digest"], digest)
        self.assertEqual(Path(result["version_dir"]), version_dir)
        self.assertEqual(Path(result["wrapper"]), wrapper)
        self.assertEqual(set(result["files"]), {name for name, _ in values})
        for name, data in values:
            self.assertEqual((version_dir / name).read_bytes(), data)
            self.assertEqual(stat.S_IMODE((version_dir / name).stat().st_mode), 0o600)
        wrapper_text = wrapper.read_text(encoding="utf-8")
        self.assertIn(install_mcp.INSTALLER_MARKER, wrapper_text)
        self.assertIn(str(Path(sys.executable)), wrapper_text)
        self.assertIn(str(version_dir / "mcp_server.py"), wrapper_text)
        self.assertEqual(stat.S_IMODE(wrapper.stat().st_mode), 0o700)
        self.assertEqual((prefix / "keep.txt").read_text(encoding="utf-8"), "leave me")

    def test_install_is_idempotent_and_updates_owned_wrapper(self):
        prefix = self.root / "prefix"
        first = install_mcp.install(source_root=self.source, prefix=prefix)
        wrapper = Path(first["wrapper"])
        first_wrapper = wrapper.read_bytes()
        first_mtime = wrapper.stat().st_mtime_ns

        second = install_mcp.install(source_root=self.source, prefix=prefix)

        self.assertEqual(first, second)
        self.assertEqual(wrapper.read_bytes(), first_wrapper)
        self.assertEqual(wrapper.stat().st_mtime_ns, first_mtime)

        (self.source / "mcp_server.py").write_bytes(b"# changed\n")
        updated = install_mcp.install(source_root=self.source, prefix=prefix)
        updated_wrapper = Path(updated["wrapper"])
        self.assertNotEqual(updated["digest"], first["digest"])
        self.assertIn(str(Path(updated["version_dir"]) / "mcp_server.py"),
                      updated_wrapper.read_text(encoding="utf-8"))
        self.assertIn(install_mcp.INSTALLER_MARKER,
                      updated_wrapper.read_text(encoding="utf-8"))

    def test_idempotent_install_repairs_modes(self):
        prefix = self.root / "prefix"
        first = install_mcp.install(source_root=self.source, prefix=prefix)
        version_dir = Path(first["version_dir"])
        wrapper = Path(first["wrapper"])
        for path in version_dir.iterdir():
            path.chmod(0o644)
        wrapper.chmod(0o644)

        second = install_mcp.install(source_root=self.source, prefix=prefix)

        self.assertEqual(second, first)
        for name in install_mcp.REQUIRED_FILES:
            self.assertEqual(stat.S_IMODE((version_dir / name).stat().st_mode), 0o600)
        self.assertEqual(stat.S_IMODE(wrapper.stat().st_mode), 0o700)

    def test_symlinked_interpreter_path_is_preserved_and_validated(self):
        interpreter_link = self.root / "python-link"
        interpreter_link.symlink_to(sys.executable)

        result = install_mcp.install(
            source_root=self.source,
            prefix=self.root / "prefix",
            python_executable=interpreter_link,
        )

        wrapper_text = Path(result["wrapper"]).read_text(encoding="utf-8")
        self.assertIn(str(interpreter_link), wrapper_text)

    def test_main_prints_only_generic_mcp_configuration(self):
        prefix = self.root / "prefix"
        stdout = io.StringIO()
        stderr = io.StringIO()
        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
            code = install_mcp.main(["--prefix", str(prefix)])

        self.assertEqual(code, 0)
        self.assertEqual(stderr.getvalue(), "")
        config = json.loads(stdout.getvalue())
        self.assertEqual(set(config), {"mcpServers"})
        server = config["mcpServers"][install_mcp.SERVER_NAME]
        self.assertEqual(server["args"], [])
        self.assertEqual(server["command"],
                         str(prefix.resolve() / "bin" / "outlook-local-mcp"))

    def test_main_selects_explicit_profile_in_multi_profile_configuration(self):
        profiles = [self.root / "First Profile", self.root / "中文 Profile"]
        for profile in profiles:
            profile.mkdir()
            (profile / "HxStore.hxd").write_bytes(b"synthetic cache")
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            code = install_mcp.main([
                "--prefix", str(self.root / "prefix"), "--profile", str(profiles[1])
            ])
        self.assertEqual(code, 0)
        server = json.loads(stdout.getvalue())["mcpServers"][install_mcp.SERVER_NAME]
        self.assertEqual(server["args"], ["--profile", str(profiles[1].resolve())])
        self.assertTrue(Path(server["command"]).is_file())

    def test_unowned_wrapper_is_never_overwritten(self):
        prefix = self.root / "prefix"
        wrapper = prefix / "bin" / "outlook-local-mcp"
        wrapper.parent.mkdir(parents=True)
        original = b"#!/bin/sh\necho user-owned\n"
        wrapper.write_bytes(original)

        with self.assertRaises(install_mcp.InstallError):
            install_mcp.install(source_root=self.source, prefix=prefix)
        self.assertEqual(wrapper.read_bytes(), original)

    def test_existing_installed_file_with_different_content_is_never_overwritten(self):
        prefix = self.root / "prefix"
        digest, _ = install_mcp.source_digest(self.source)
        version_dir = prefix / "share" / "outlook-local-mcp" / digest
        version_dir.mkdir(parents=True)
        target = version_dir / install_mcp.REQUIRED_FILES[0]
        target.write_bytes(b"unrelated content")

        with self.assertRaises(install_mcp.InstallError):
            install_mcp.install(source_root=self.source, prefix=prefix)
        self.assertEqual(target.read_bytes(), b"unrelated content")

    def test_symlinked_source_directory_and_destination_paths_are_rejected(self):
        source_link = self.root / "source-link"
        source_link.symlink_to(self.source, target_is_directory=True)
        with self.assertRaises(install_mcp.InstallError):
            install_mcp.install(source_root=source_link, prefix=self.root / "prefix")

        prefix = self.root / "prefix"
        prefix.mkdir()
        external_share = self.root / "external-share"
        external_share.mkdir()
        (prefix / "share").symlink_to(external_share, target_is_directory=True)
        with self.assertRaises(install_mcp.InstallError):
            install_mcp.install(source_root=self.source, prefix=prefix)

        safe_prefix = self.root / "safe-prefix"
        wrapper = safe_prefix / "bin" / "outlook-local-mcp"
        wrapper.parent.mkdir(parents=True)
        external_wrapper = self.root / "external-wrapper"
        external_wrapper.write_bytes(b"keep")
        wrapper.symlink_to(external_wrapper)
        with self.assertRaises(install_mcp.InstallError):
            install_mcp.install(source_root=self.source, prefix=safe_prefix)
        self.assertEqual(external_wrapper.read_bytes(), b"keep")

    def test_shared_prefix_is_rejected_without_writing_under_it(self):
        if not hasattr(os, "getuid"):
            self.skipTest("POSIX ownership check")
        shared = self.root / "shared"
        shared.mkdir()
        shared.chmod(0o777)
        prefix = shared / "prefix"

        with self.assertRaises(install_mcp.InstallError):
            install_mcp.install(source_root=self.source, prefix=prefix)

        self.assertFalse(prefix.exists())
        self.assertEqual(list(shared.iterdir()), [])
        self.assertEqual(stat.S_IMODE(shared.stat().st_mode), 0o777)


if __name__ == "__main__":
    unittest.main()
