"""Native kit tooling controls; cleanup owns this test directory."""
from contextlib import ExitStack
from hashlib import sha256
import importlib.util
import json
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import patch
import tempfile

ROOT = Path(__file__).resolve().parents[2]
WINDOWS_PATH = ROOT / "engine/winbooksplit_windows.py"


@unittest.skipUnless(os.name == "nt", "Native Windows directory guard controls")
class Controls(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        global kit, windows, WORK
        source = ROOT / "tools/compatibility/kit_builder.py"
        spec = importlib.util.spec_from_file_location("wbs_native_kit_controls", source)
        kit = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = kit
        spec.loader.exec_module(kit)
        data = WINDOWS_PATH.read_bytes()
        windows = kit.load_authenticated(WINDOWS_PATH,
            {"bytes": len(data), "sha256": sha256(data).hexdigest()}, "wbs_native_kit_windows")
        temporary = tempfile.TemporaryDirectory(prefix="wbs-native-kit-controls-")
        cls.addClassCleanup(temporary.cleanup)
        WORK = Path(temporary.name).resolve()

    def test_duplicate_case_json(self):
        with self.assertRaises(kit.KitError):
            kit.unique(b'{"sha256":"a","SHA256":"b"}')

    def test_nonfinite_json(self):
        with self.assertRaises(kit.KitError):
            kit.unique(b'{"bytes":NaN}')

    def test_wrong_code_hash_does_not_execute(self):
        source = WORK / "untrusted.py"
        source.write_bytes(b'raise RuntimeError("UNTRUSTED_EXECUTED")\n')
        with self.assertRaisesRegex(kit.KitError, "helper_identity_mismatch"):
            kit.load_authenticated(source, {"bytes": source.stat().st_size, "sha256": "0" * 64}, "untrusted_control")

    def test_bool_code_size_does_not_execute(self):
        source = WORK / "bool.py"
        source.write_bytes(b'raise RuntimeError("UNTRUSTED_EXECUTED")\n')
        with self.assertRaises(kit.KitError):
            kit.load_authenticated(source, {"bytes": True, "sha256": sha256(source.read_bytes()).hexdigest()}, "bool_control")

    def test_positive_native_guarded_tree(self):
        output = WORK / "positive"
        kit.guarded_write_tree(output, {"engine/code.py": b"authored code\n", "VERSION": b"1.0.0\n"}, windows)
        self.assertEqual((output / "engine/code.py").read_bytes(), b"authored code\n")

    def test_existing_target_preserved(self):
        output = WORK / "existing"
        output.mkdir()
        (output / "sentinel").write_bytes(b"AUTHORED ORIGINAL")
        with self.assertRaises(FileExistsError):
            kit.guarded_write_tree(output, {"new": b"new"}, windows)
        self.assertEqual(list(output.iterdir()), [output / "sentinel"])
        self.assertEqual((output / "sentinel").read_bytes(), b"AUTHORED ORIGINAL")

    def test_unsafe_paths_rejected_before_target_creation(self):
        for number, relative in enumerate(("../outside", "engine/../outside", "/outside", "engine\\outside", "x:stream", "CON", "x. ")):
            target = WORK / ("unsafe-" + str(number))
            with self.assertRaises(kit.KitError):
                kit.guarded_write_tree(target, {relative: b"authored"}, windows)
            self.assertFalse(target.exists())

    def test_case_and_ancestor_collisions_rejected_before_create(self):
        for number, payload in enumerate(({"engine": b"f", "engine/x": b"x"},
                {"ENGINE/x": b"x", "engine/y": b"y"}, {"a": b"x", "A": b"y"})):
            target = WORK / ("collision-" + str(number))
            with self.assertRaises(kit.KitError):
                kit.guarded_write_tree(target, payload, windows)
            self.assertFalse(target.exists())

    def test_developer_git_identity_ignores_repository_environment(self):
        root = WINDOWS_PATH.parent.parent
        source = root / "tools/ci/check_package_inputs.py"
        data = source.read_bytes()
        checker = kit.load_authenticated(source, {"bytes": len(data), "sha256": sha256(data).hexdigest()}, "kit_control_checker")
        expected = kit.git_identity(root, checker)
        hostile = {name: str(WORK / "unrelated-git") for name in checker.REPOSITORY_ENV_KEYS}
        hostile["GIT_NO_REPLACE_OBJECTS"] = "0"
        inherited = dict(os.environ)
        real_run = checker.subprocess.run
        calls = []
        def observed_run(*arguments, **options):
            calls.append(options)
            return real_run(*arguments, **options)
        with patch.dict(os.environ, hostile), patch.object(checker.subprocess, "run", side_effect=observed_run):
            self.assertEqual(kit.git_identity(root, checker), expected)
            self.assertEqual(os.environ["GIT_NO_REPLACE_OBJECTS"], "0")
        self.assertEqual(dict(os.environ), inherited)
        self.assertEqual(len(calls), 2)
        for options in calls:
            self.assertEqual(options["timeout"], 30)
            self.assertEqual(options["env"]["GIT_NO_REPLACE_OBJECTS"], "1")
            self.assertFalse(set(checker.REPOSITORY_ENV_KEYS).intersection(options["env"]))

    def test_output_identities_held_through_archive_and_receipt(self):
        output = WORK / "held-publication"
        bindings = {"production": {"authored": "pin"}, "kit": {"mode": "SYNTHETIC_CONTROL"}}
        payload = {"VERSION": b"1.0.0\n"}
        real_zip = kit.zipfile.ZipFile
        real_bounded = kit.bounded
        observations = []
        def held_identity(phase):
            with self.assertRaises(OSError):
                output.rename(WORK / ("replacement-" + phase))
            with self.assertRaises(OSError):
                WORK.rename(WORK.with_name(WORK.name + "-replaced"))
            observations.append(phase)
        def observed_zip(*arguments, **options):
            held_identity("archive_write")
            return real_zip(*arguments, **options)
        def observed_bounded(path, maximum):
            if Path(path) == output.with_suffix(".zip"):
                held_identity("archive_read")
            elif Path(path) == output.with_suffix(".build.json"):
                held_identity("receipt_read")
            return real_bounded(path, maximum)
        with patch.object(kit.zipfile, "ZipFile", side_effect=observed_zip), patch.object(kit, "bounded", side_effect=observed_bounded):
            result = kit.write_build_outputs(output, payload, windows, bindings)
        self.assertEqual(observations, ["archive_write", "archive_read", "receipt_read"])
        self.assertTrue(result["success"])
        self.assertEqual((output / "VERSION").read_bytes(), payload["VERSION"])
        self.assertEqual(kit.unique(output.with_suffix(".build.json").read_bytes()), result)

    def test_recreated_archive_is_not_adopted(self):
        output = WORK / "archive-substitution"
        archive = output.with_suffix(".zip")
        bindings = {"production": {"authored": "pin"}, "kit": {"mode": "SYNTHETIC_CONTROL"}}
        real_guard = windows.FileGuard
        replaced = []
        def substitute_before_guard(path, *arguments, **options):
            if Path(path) == archive:
                archive.unlink()
                archive.write_bytes(b"AUTHORED SUBSTITUTION CONTROL")
                replaced.append(True)
            return real_guard(path, *arguments, **options)
        with patch.object(windows, "FileGuard", side_effect=substitute_before_guard):
            with self.assertRaisesRegex(kit.KitError, "kit_archive_bytes_differ"):
                kit.write_build_outputs(output, {"VERSION": b"1.0.0\n"}, windows, bindings)
        self.assertEqual(replaced, [True])
        self.assertEqual(archive.read_bytes(), b"AUTHORED SUBSTITUTION CONTROL")
        self.assertFalse(output.with_suffix(".build.json").exists())


if __name__ == "__main__":
    unittest.main()
