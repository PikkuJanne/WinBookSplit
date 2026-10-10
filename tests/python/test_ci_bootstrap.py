"""Offline authored controls for isolated developer-package extraction."""
from hashlib import sha256
import importlib.util
import io
from pathlib import Path
import stat
import tempfile
import unittest
from unittest.mock import patch
import zipfile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_ci_bootstrap", ROOT / "tools/ci/bootstrap.py")
bootstrap = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bootstrap)


def archive(members):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as target:
        for name, data in members:
            if isinstance(name, str):
                # ZipInfo otherwise normalizes backslashes while authoring on Windows.
                info = zipfile.ZipInfo()
                info.filename = name
                name = info
            target.writestr(name, data)
    return stream.getvalue()


class BootstrapTests(unittest.TestCase):
    def test_matching_bytes_extract_only_declared_ordinary_members(self):
        data = archive([("Pester.psd1", b"Authored manifest"), ("bin/helper.txt", b"Owned helper")])
        with tempfile.TemporaryDirectory() as work:
            target = Path(work) / "tools/Pester/6.2.0"
            observed = bootstrap.extract_package(data, target, sha256(data).hexdigest())
            self.assertEqual(observed, {"Pester.psd1": sha256(b"Authored manifest").hexdigest(),
                                        "bin/helper.txt": sha256(b"Owned helper").hexdigest()})

    def test_wrong_hash_rejects_before_any_destination_creation(self):
        data = archive([("Pester.psd1", b"Authored manifest")])
        with tempfile.TemporaryDirectory() as work:
            target = Path(work) / "untouched"
            with self.assertRaisesRegex(ValueError, "SHA256"):
                bootstrap.extract_package(data, target, "0" * 64)
            self.assertFalse(target.exists())

    def test_traversal_absolute_alias_and_duplicate_members_reject_before_write(self):
        invalid = [[("../outside.txt", b"x")], [("/outside.txt", b"x")], [("C:/outside.txt", b"x")],
                   [("bin\\outside.txt", b"x")], [("bin/./outside.txt", b"x")],
                   [("file.txt", b"a"), ("FILE.txt", b"b")]]
        with tempfile.TemporaryDirectory() as work:
            for number, members in enumerate(invalid):
                with self.subTest(members=members):
                    target = Path(work) / str(number)
                    data = archive(members)
                    with self.assertRaisesRegex(ValueError, "unsafe/duplicate"):
                        bootstrap.extract_package(data, target, sha256(data).hexdigest())
                    self.assertFalse(target.exists())

    def test_symlink_and_expanded_bounds_reject_without_adopting_files(self):
        link = zipfile.ZipInfo("owned-link")
        link.create_system = 3
        link.external_attr = (stat.S_IFLNK | 0o777) << 16
        data = archive([(link, b"../neighbor")])
        with tempfile.TemporaryDirectory() as work:
            target = Path(work) / "tools"
            with self.assertRaisesRegex(ValueError, "unsafe/duplicate"):
                bootstrap.extract_package(data, target, sha256(data).hexdigest())
            ordinary = archive([("large.txt", b"Authored bytes")])
            with patch.object(bootstrap, "MAX_EXPANDED", 1):
                with self.assertRaisesRegex(ValueError, "expanded"):
                    bootstrap.extract_package(ordinary, target, sha256(ordinary).hexdigest())
            self.assertFalse(target.exists())

    def test_windows_device_trailing_and_special_aliases_reject_before_write(self):
        invalid = ['CON', 'nul.txt', 'aux/log.txt', 'COM9.bin', 'LPT1', 'authored.txt.', 'authored.txt ',
                   'a//b', 'a//', 'a/\x00b', 'a/\x1fb', 'a/\x7fb', 'a/q?b', 'a/s*b', 'a/quote"b',
                   'a/pipe|b', 'a/left<b', 'a/right>b', 'a/stream:ads']
        with tempfile.TemporaryDirectory() as work:
            for number, name in enumerate(invalid):
                with self.subTest(member=name):
                    target = Path(work)/str(number)
                    data = archive([(name, b'Authored alias bytes')])
                    with self.assertRaisesRegex(ValueError, 'unsafe/duplicate'):
                        bootstrap.extract_package(data, target, sha256(data).hexdigest())
                    self.assertFalse(target.exists())

    def test_existing_tool_root_rejects_without_network_or_overwrite(self):
        with tempfile.TemporaryDirectory() as work:
            root = Path(work) / "tools"
            root.mkdir()
            prior = root / "prior.txt"
            prior.write_bytes(b"Immutable prior tools")
            with patch.object(bootstrap, "urlopen", side_effect=AssertionError("No network permitted")):
                with self.assertRaisesRegex(ValueError, "new absolute external"):
                    bootstrap.bootstrap(root, Path(work) / "report.json")
            self.assertEqual(prior.read_bytes(), b"Immutable prior tools")


if __name__ == "__main__":
    unittest.main(verbosity=2)
