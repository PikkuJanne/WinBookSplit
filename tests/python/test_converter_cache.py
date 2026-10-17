"""Owned cache/provenance controls; synthetic metadata is never native evidence."""
from copy import deepcopy
import importlib.util
import os
import subprocess
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_converter_cache_controls", ROOT / "tests/runtime/converter_cache.py")
cache = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(cache)


def synthetic_reference(marker):
    package = {"Name": "PythonSoftwareFoundation.PythonManager", "Version": "26.4.240.0",
               "PackageFullName": "PythonSoftwareFoundation.PythonManager_26.4.240.0_x64__qbz5n2kfra8p0",
               "PackageFamilyName": "PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0",
               "Publisher": "CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B", "SignatureKind": "Store",
               "Architecture": "X64", "Status": "Ok",
               **dict.fromkeys(("IsDevelopmentMode", "IsPartiallyStaged", "IsFramework", "IsResourcePackage"), False)}
    probe = {"Pid": 12345, "ExitCode": 0, "ElapsedSeconds": 0.1, "PackageIdentityVerified": True,
             "PackageFullName": package["PackageFullName"], "Stdout": "", "Stderr": "",
             "StdoutTotalBytes": 0, "StderrTotalBytes": 0}
    probe.update(dict.fromkeys(("JobAssigned", "ParentStopped", "DescendantsStopped", "StreamsComplete", "InputWriterStopped"), True))
    probe.update(dict.fromkeys(("TimedOut", "Cancelled", "StdoutTruncated", "StderrTruncated"), False))
    probe.update(dict.fromkeys(("StartError", "StopError", "StreamError", "ResultError", "InteractionError", "InputError"), None))
    row = {"registration": {"Package": package}, "listing_command": ["list", "--format=exe"], "listing": probe,
           "child_localappdata_only": True, "parent_localappdata_restored": True, "marker_absent_before": True,
           "manager_sha256_before": "a" * 64, "manager_sha256_after": "a" * 64, "manager_marker": deepcopy(marker)}
    return {"kind": "VERIFIED_INSTALLED_ONLY_MANAGER_CACHE_REFERENCE", "parent_environment_unchanged": True,
            "runtime_acquisition_attempted": False, "supplied_to_application": False, "probes": [row],
            "fixture_kind": "SYNTHETIC_UNIT_ONLY"}


class ConverterCacheControls(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-cache-controls-")
        self.addCleanup(self.owned.cleanup)
        self.root = Path(self.owned.name).resolve() / "cache"
        (self.root / "Microsoft/Windows/Caches").mkdir(parents=True)
        self.before = cache.observe(self.root)

    def welcome(self):
        path = self.root / cache.MARKER
        path.parent.mkdir(parents=True)
        path.write_bytes(cache.welcome_bytes("26.4"))
        return path

    def complete(self):
        self.welcome()
        after = cache.observe(self.root)
        return cache.environment_record(self.before, after, synthetic_reference(after["manager_marker"]))

    def test_original_empty_windows_chain_remains_supported(self):
        record = cache.environment_record(self.before, cache.observe(self.root), None)
        self.assertTrue(record["windows_cache_inventory_unchanged"])
        self.assertEqual(record["manager_cache_observation"], {"before": None, "after": None, "reference": None})
        self.assertNotIn("owned_empty_cache_inventory_unchanged", record)

    def test_only_known_marker_matches_separate_synthetic_reference(self):
        record = self.complete()
        cache.validate_environment(record)
        self.assertEqual(record["manager_cache_observation"]["after"]["relative"], cache.MARKER)

    def test_marker_without_matching_verified_reference_is_rejected(self):
        record = self.complete()
        for defect in (lambda r: r.update(reference=None),
                       lambda r: r["reference"].update(probes=[]),
                       lambda r: r["reference"].update(runtime_acquisition_attempted=True),
                       lambda r: r["reference"].update(supplied_to_application=True),
                       lambda r: r["reference"]["probes"][0].update(listing_command=["install", "3.15"]),
                       lambda r: r["reference"]["probes"][0].update(manager_sha256_after="b" * 64),
                       lambda r: r["reference"]["probes"][0].update(manager_marker=None),
                       lambda r: r["reference"]["probes"][0]["listing"].update(PackageIdentityVerified=False),
                       lambda r: r["reference"]["probes"][0]["listing"].update(ExitCode=23),
                       lambda r: r["reference"]["probes"][0]["listing"].update(StreamsComplete=False),
                       lambda r: r["reference"]["probes"][0]["listing"].update(ElapsedSeconds=99),
                       lambda r: r["reference"]["probes"][0]["listing"].update(StderrTotalBytes=1)):
            changed = deepcopy(record)
            defect(changed["manager_cache_observation"])
            with self.subTest(defect=defect), self.assertRaises(ValueError):
                cache.validate_environment(changed)

    def test_missing_or_cross_paired_full_manager_identity_is_rejected(self):
        record = self.complete()
        for fields in (None, "Foreign_26.4.240.0_x64__qbz5n2kfra8p0"):
            changed = deepcopy(record)
            row = changed["manager_cache_observation"]["reference"]["probes"][0]
            row["registration"]["Package"]["PackageFullName"] = fields
            row["listing"]["PackageFullName"] = fields
            with self.subTest(identity=fields), self.assertRaises(ValueError):
                cache.validate_environment(changed)

    def test_windows_cache_identity_and_new_proof_cannot_be_omitted(self):
        record = self.complete()
        for defect in (lambda r: r.pop("manager_cache_observation"),
                       lambda r: r.update(schema_version=1),
                       lambda r: r.update(schema_version=True),
                       lambda r: r.update(owned_empty_cache_inventory_unchanged=True),
                       lambda r: r["cache_directory_identities_after"]["Microsoft"].update(inode=0),
                       lambda r: r["manager_cache_observation"].update(before=r["manager_cache_observation"]["after"]),
                       lambda r: r["manager_cache_observation"]["after"].update(links=2),
                       lambda r: r["manager_cache_observation"]["after"].update(sha256="f" * 64)):
            changed = deepcopy(record)
            defect(changed)
            with self.subTest(defect=defect), self.assertRaises(ValueError):
                cache.validate_environment(changed)

    def test_unknown_entries_still_fail_each_exact_directory_closure(self):
        for relative in ("Programs", "Other", "Microsoft/Windows/Caches/x", "Microsoft/Windows/Calibre",
                         "Python/other.txt", "Python/_cache/log.txt", "Python/_cache/sub"):
            with self.subTest(relative=relative), tempfile.TemporaryDirectory(prefix="wbs-cache-extra-") as owned:
                root = Path(owned)
                (root / "Microsoft/Windows/Caches").mkdir(parents=True)
                if relative.startswith("Python/"):
                    (root / "Python/_cache").mkdir(parents=True)
                    (root / cache.MARKER).write_bytes(cache.welcome_bytes("26.4"))
                target = root / relative
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(b"AUTHORED_UNEXPECTED_ENTRY")
                with self.assertRaises(ValueError):
                    cache.observe(root)

    def test_unknown_empty_oversized_or_malformed_marker_fails(self):
        path = self.welcome()
        for raw in (b"", b"X" * (cache.MAX_MARKER_BYTES + 1), b"26.4\nunknown\n",
                    cache.welcome_bytes("26.4").replace(b"github.com", b"example.co"),
                    b"\xff\r\n"):
            with self.subTest(size=len(raw)):
                path.write_bytes(raw)
                with self.assertRaises(ValueError):
                    cache.observe(self.root)

    def test_authored_hardlink_marker_is_rejected(self):
        path = self.welcome()
        os.link(path, Path(self.owned.name) / "owned-hardlink.txt")
        with self.assertRaises(ValueError):
            cache.observe(self.root)

    def test_reparse_attribute_and_symbolic_kind_are_rejected(self):
        path = self.welcome()
        original = Path.lstat
        for attributes, mode in ((0x400, 0o100644), (0, 0o120777)):
            def linked(target):
                if target == path:
                    return SimpleNamespace(st_mode=mode, st_file_attributes=attributes)
                return original(target)
            with self.subTest(attributes=attributes, mode=mode), patch.object(Path, "lstat", linked), self.assertRaises(ValueError):
                cache.observe(self.root)

    @unittest.skipUnless(os.name == "nt", "Actual Windows junction control")
    def test_actual_authored_junction_directory_is_rejected(self):
        target = Path(self.owned.name) / "owned-junction-target"
        (target / "_cache").mkdir(parents=True)
        (target / "_cache/last_welcome.txt").write_bytes(cache.welcome_bytes("26.4"))
        link = self.root / "Python"
        self.assertTrue(target.is_relative_to(Path(self.owned.name)) and link.is_relative_to(Path(self.owned.name)))
        command = [str(Path(os.environ["SystemRoot"]) / "System32/cmd.exe"), "/d", "/c", "mklink", "/J", str(link), str(target)]
        result = subprocess.run(command, capture_output=True, timeout=10, check=False)
        self.assertEqual(result.returncode, 0, "Authored junction positive control failed")
        try:
            with self.assertRaises(ValueError):
                cache.observe(self.root)
        finally:
            os.rmdir(link)  # This run's junction only; its owned target remains.
        self.assertEqual((target / "_cache/last_welcome.txt").read_bytes(), cache.welcome_bytes("26.4"))

    def test_directory_enumeration_refuses_early_at_tiny_fixed_bound(self):
        consumed = []
        class Entries:
            def __enter__(self):
                for name in ("Microsoft", "Python", "THIRD", "MUST_NOT_READ"):
                    consumed.append(name)
                    yield SimpleNamespace(name=name)
            def __exit__(self, *args):
                return False
        with patch.object(cache.os, "scandir", return_value=Entries()), self.assertRaises(ValueError):
            cache.observe(self.root)
        self.assertEqual(consumed, ["Microsoft", "Python", "THIRD"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
