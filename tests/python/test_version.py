"""Canonical VERSION parsing and actual Python early-failure controls; no PDFs."""
import ast
from hashlib import sha256
import io
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
ENGINE = ROOT / "engine/winbooksplit_engine.py"
tree = ast.parse(ENGINE.read_text(encoding="utf-8"))
function = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "_read_application_version")
namespace = {"Path": Path, "re": re}
exec(compile(ast.Module(body=[function], type_ignores=[]), str(ENGINE), "exec"), namespace)
read_version = namespace["_read_application_version"]


class ApplicationVersionTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-version-unit-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.version = self.work / "VERSION"

    def test_shipped_candidate_has_one_version_source_and_unchanged_outcome_schema(self):
        self.assertEqual((ROOT / "VERSION").read_bytes(), b"1.0.0\n")
        self.assertEqual(read_version(ROOT / "VERSION"), "1.0.0")
        contract = json.loads((ROOT / "engine/WinBookSplit.Outcomes.json").read_text(encoding="utf-8"))
        self.assertEqual(set(contract), {"schema_version", "codes"})
        self.assertEqual(contract["schema_version"], 1)
        self.assertEqual(contract["codes"]["processor_protocol_failed"], 6)

    def test_bounded_ascii_numeric_versions_accept_one_optional_line_ending(self):
        for data, expected in ((b"1.0.0", "1.0.0"), (b"1.0.0\n", "1.0.0"),
                (b"1.0.0\r\n", "1.0.0"), (b"0.0.0", "0.0.0"),
                (b"1234567890.1234567890.1234567890\r\n", "1234567890.1234567890.1234567890")):
            with self.subTest(data=data):
                self.version.write_bytes(data)
                self.assertEqual(read_version(self.version), expected)

    def test_missing_malformed_nonascii_and_oversized_version_never_fall_back(self):
        with self.assertRaisesRegex(ValueError, "Missing or malformed shipped application VERSION"):
            read_version(self.version)
        for data in (b"", b"1.0.0-dev\n", b"v1.0.0", b"01.0.0", b"1.00.0", b"1.0.00", b"-1.0.0",
                b"1.0", b"1.0.0.0", b"1.0.0\r", b"1.0.0\n\n", b"1.0.0\r\n\r\n", b" 1.0.0",
                b"1.0.0 ", b"\xef\xbb\xbf1.0.0", "１.0.0".encode(), b"1.0.0\x00", b"9" * 100000):
            with self.subTest(data=data[:40]):
                self.version.write_bytes(data)
                with self.assertRaisesRegex(ValueError, "Missing or malformed shipped application VERSION"):
                    read_version(self.version)

    def test_reader_consumes_at_most_35_bytes_before_oversize_refusal(self):
        class ObservedStream(io.BytesIO):
            sizes = []
            def read(self, size=-1):
                self.sizes.append(size)
                return super().read(size)
        stream = ObservedStream(b"x" * 100000)
        with patch.object(Path, "open", return_value=stream):
            with self.assertRaises(ValueError):
                read_version(self.version)
        self.assertEqual(stream.sizes, [35])
        self.assertTrue(stream.closed)

    def test_actual_engine_missing_or_bad_version_fails_before_dependency_and_source_access(self):
        app = self.work / "app"
        (app / "engine").mkdir(parents=True)
        for name in ("winbooksplit_engine.py", "WinBookSplit.Outcomes.json"):
            shutil.copyfile(ROOT / "engine" / name, app / "engine" / name)
        source, output = self.work / "absent-input.pdf", self.work / "absent-output"
        sentinel = self.work / "neighbor.txt"
        sentinel.write_bytes(b"Owned immutable neighbor\n")
        command = [sys.executable, "-I", "-B", "-S", str(app / "engine/winbooksplit_engine.py"),
                   str(source), str(output), "manual", "1"]
        for data in (None, b"1.0.0-dev\n", b"x" * 100000):
            if data is not None:
                (app / "VERSION").write_bytes(data)
            before = {p.relative_to(self.work).as_posix(): sha256(p.read_bytes()).hexdigest()
                      for p in self.work.rglob("*") if p.is_file()}
            process = subprocess.run(command, stdin=subprocess.DEVNULL, capture_output=True, timeout=30, check=False)
            self.assertEqual(process.returncode, 6, process.stderr.decode(errors="replace"))
            self.assertEqual(process.stderr, b"")
            lines = process.stdout.decode("utf-8").splitlines()
            self.assertEqual(len(lines), 1)
            result = json.loads(lines[0])
            self.assertEqual((result["protocol"], result["version"], result["code"], result["exit_code"]),
                             ("winbooksplit.result", 1, "processor_protocol_failed", 6))
            self.assertEqual(result["message"], "Missing or malformed shipped application VERSION.")
            self.assertEqual(result["written_count"], 0)
            self.assertIsNone(result["execution"])
            self.assertIsNone(result["diagnostic"])
            self.assertFalse(source.exists())
            self.assertFalse(output.exists())
            self.assertEqual(before, {p.relative_to(self.work).as_posix(): sha256(p.read_bytes()).hexdigest()
                                     for p in self.work.rglob("*") if p.is_file()})
        (app / "VERSION").write_bytes(b"1.0.0\n")
        process = subprocess.run(command, stdin=subprocess.DEVNULL, capture_output=True, timeout=30, check=False)
        self.assertEqual(process.returncode, 3)
        self.assertEqual(json.loads(process.stdout)["code"], "dependency_missing")


if __name__ == "__main__":
    unittest.main()
