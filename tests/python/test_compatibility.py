"""Synthetic policy boundary tests; these are not native version-matrix passes."""

import builtins
import copy
import importlib.util
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
MODULE = ROOT / "engine" / "winbooksplit_compatibility.py"
spec = importlib.util.spec_from_file_location("wbs_test_compatibility", MODULE)
compatibility = importlib.util.module_from_spec(spec)
spec.loader.exec_module(compatibility)
_DEFAULT_POLICY = object()


def synthetic_policy():
    return {
        "schema_version": 1, "policy_version": 1, "reviewed_on": "2026-10-11",
        "python": {"implementation": "cpython", "platform": "win32", "machine": "AMD64", "bits": 64,
                   "stable_only": True, "gil_disabled": False,
                   "branches": [{"major": 3, "minor": minor, "compatibility_min_patch": 0,
                                 "security_min_patch": minimum, "recommended_patch": latest,
                                 "blocked_patches": [], "security_sources": ["SyntheticSecurity"]}
                                for minor, minimum, latest in ((14, 7, 8), (13, 15, 16))]},
        "pypdf": {"version": "6.20.0", "wheel_sha256": "a" * 64},
        "calibre": {"version": "9.15.0"},
        "excluded_python_branches": [{"major": 3, "minor": minor,
                                      "reason": "Synthetic evidence-based distribution exclusion"}
                                     for minor in (11, 12)],
        "sources": [{"id": "SyntheticSecurity", "url": "https://www.python.org/synthetic-policy-test",
                     "relevance": "Synthetic fixture reference; not an actual advisory claim"}],
    }


def synthetic_runtime(minor=14, patch_number=8):
    return {"version": "3." + str(minor) + "." + str(patch_number), "version_info": [3, minor, patch_number],
            "releaselevel": "final", "serial": 0, "implementation": "cpython", "platform": "win32",
            "machine": "AMD64", "bits": 64, "gil_disabled": False, "isolated": True,
            "dont_write_bytecode": True, "executable": r"C:\Synthetic Trusted Python\python.exe"}


class InterpreterPolicyTests(unittest.TestCase):
    def setUp(self):
        self.policy = synthetic_policy()

    def assert_reason(self, record, reason, policy=_DEFAULT_POLICY):
        result = compatibility.validate_interpreter(record, self.policy if policy is _DEFAULT_POLICY else policy)
        self.assertEqual(result["reason"], reason)
        self.assertIs(result["ok"], reason == "supported")
        self.assertTrue(result["message"])
        return result

    def test_both_reviewed_branches_accept_floor_and_later_stable_patches(self):
        for minor, patches in ((13, (15, 16, 17, 99)), (14, (7, 8, 9, 10, 99))):
            for patch_number in patches:
                with self.subTest(minor=minor, patch=patch_number):
                    self.assert_reason(synthetic_runtime(minor, patch_number), "supported")

    def test_security_floor_below_boundary_explains_the_reported_3_14_2(self):
        for minor, patches in ((13, (0, 14)), (14, (0, 2, 6))):
            for patch_number in patches:
                with self.subTest(minor=minor, patch=patch_number):
                    result = self.assert_reason(synthetic_runtime(minor, patch_number), "python_security_floor")
                    self.assertIn("SyntheticSecurity", result["message"])
                    self.assertNotIn("exactly", result["message"])

    def test_compatibility_floor_is_independent_of_security_floor(self):
        self.policy["python"]["branches"][0]["compatibility_min_patch"] = 8
        self.assert_reason(synthetic_runtime(14, 7), "python_compatibility_floor")
        self.assert_reason(synthetic_runtime(14, 8), "supported")
        self.assert_reason(synthetic_runtime(14, 10), "supported")

    def test_precise_blocked_patch_does_not_become_equality_gate(self):
        self.policy["python"]["branches"][0]["blocked_patches"] = [9]
        self.assert_reason(synthetic_runtime(14, 8), "supported")
        self.assert_reason(synthetic_runtime(14, 9), "python_security_floor")
        self.assert_reason(synthetic_runtime(14, 10), "supported")

    def test_excluded_and_future_minor_are_rejected(self):
        for minor in (11, 12, 16):
            with self.subTest(minor=minor):
                result = self.assert_reason(synthetic_runtime(minor, 99), "unsupported_python")
                self.assertIn("distribution exclusion" if minor in (11, 12) else "outside", result["message"])

    def test_prerelease_and_development_records_are_rejected(self):
        for releaselevel in ("alpha", "beta", "candidate"):
            record = synthetic_runtime()
            record["releaselevel"] = releaselevel
            record["serial"] = 1
            self.assert_reason(record, "unstable_python")
        for version in ("3.14.8a1", "3.14.8b2", "3.14.8rc1", "3.14.8.dev1", "3.14.8+"):
            record = synthetic_runtime()
            record["version"] = version
            self.assert_reason(record, "unstable_python")
        record = synthetic_runtime()
        record["serial"] = 1
        self.assert_reason(record, "unstable_python")

    def test_non_cpython_and_free_threaded_are_rejected(self):
        record = synthetic_runtime()
        record["implementation"] = "pypy"
        self.assert_reason(record, "unsupported_flavor")
        record = synthetic_runtime()
        record["gil_disabled"] = True
        self.assert_reason(record, "unsupported_flavor")

    def test_arm_32bit_and_other_platform_are_rejected(self):
        for field, value in (("machine", "ARM64"), ("bits", 32), ("platform", "linux")):
            record = synthetic_runtime()
            record[field] = value
            self.assert_reason(record, "unsupported_architecture")

    def test_scalar_types_and_boolean_numbers_fail_closed(self):
        for field, value in (("bits", True), ("serial", False), ("gil_disabled", 0), ("isolated", 1),
                             ("dont_write_bytecode", "true"), ("implementation", ["cpython"]),
                             ("releaselevel", "unknown"), ("executable", "python.exe")):
            record = synthetic_runtime()
            record[field] = value
            self.assert_reason(record, "malformed_runtime")
        for version in ([3, 14, True], [3, 14], "3.14", "03.14.8", "3.14.-1", "3.14.1000001", "three"):
            record = synthetic_runtime()
            record["version"] = version
            self.assert_reason(record, "malformed_runtime")
        for info in ([True, 14, 8], [3, 14, False], [3, 14], "3.14.8", [3, 14, -1]):
            record = synthetic_runtime()
            record["version_info"] = info
            self.assert_reason(record, "malformed_runtime")

    def test_missing_fields_and_version_disagreement_fail_closed(self):
        for field in synthetic_runtime():
            record = synthetic_runtime()
            del record[field]
            self.assert_reason(record, "malformed_runtime")
        record = synthetic_runtime()
        record["version_info"] = [3, 14, 7]
        self.assert_reason(record, "malformed_runtime")
        for record in (None, [], True):
            self.assert_reason(record, "malformed_runtime")

    def test_runtime_version_string_is_required_beside_numeric_version_info(self):
        for version in ((3, 14, 10), [3, 14, 10]):
            record = synthetic_runtime(14, 10)
            record["version"] = version
            self.assert_reason(record, "malformed_runtime")

    def test_candidate_ok_flag_and_package_fields_do_not_override_identity(self):
        record = synthetic_runtime(14, 2)
        record.update(ok=True, pypdf_version="6.20.0", protocol="winbooksplit.runtime", schema_version=1)
        self.assert_reason(record, "python_security_floor")

    def test_required_isolation_flags_are_enforced(self):
        for field in ("isolated", "dont_write_bytecode"):
            record = synthetic_runtime()
            record[field] = False
            self.assert_reason(record, "malformed_runtime")

    def test_direct_malformed_policy_cannot_enable_unsupported_interpreter(self):
        for policy in (None, {}, {"schema_version": True}):
            self.assert_reason(synthetic_runtime(15, 99), "malformed_policy", policy)
        self.policy["python"]["stable_only"] = False
        self.assert_reason(synthetic_runtime(), "malformed_policy")


class PolicyLoaderTests(unittest.TestCase):
    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-compatibility-unit-")
        self.addCleanup(self.owned.cleanup)
        self.path = Path(self.owned.name) / "policy.json"

    def write(self, policy):
        self.path.write_text(json.dumps(policy), encoding="utf-8")

    def assert_rejected(self, policy):
        self.write(policy)
        with self.assertRaises(compatibility.PolicyError):
            compatibility.load_policy(self.path)

    def test_valid_policy_loads_without_mutation(self):
        original = synthetic_policy()
        self.write(original)
        self.assertEqual(compatibility.load_policy(self.path), original)

    def test_shipped_policy_is_valid_and_does_not_pin_python_to_recommendation(self):
        policy = compatibility.load_policy()
        self.assertEqual(policy["schema_version"], 1)
        for branch in policy["python"]["branches"]:
            minimum = max(branch["security_min_patch"], branch["compatibility_min_patch"])
            for patch_number in (minimum, branch["recommended_patch"] + 1, 99):
                if patch_number not in branch["blocked_patches"]:
                    record = synthetic_runtime(branch["minor"], patch_number)
                    self.assertTrue(compatibility.validate_interpreter(record, policy)["ok"])

    def test_shipped_policy_three_adds_315_without_weakening_existing_floors(self):
        policy = compatibility.load_policy()
        self.assertEqual(policy["policy_version"], 3)
        self.assertEqual(policy["reviewed_on"], "2026-10-12")
        self.assertEqual([(b["minor"], b["security_min_patch"]) for b in policy["python"]["branches"]],
                         [(15, 0), (14, 8), (13, 16)])
        for patch_number in (0, 1, 9, 10, 99):
            with self.subTest(minor=15, patch=patch_number):
                self.assertTrue(compatibility.validate_interpreter(synthetic_runtime(15, patch_number), policy)["ok"])
        for minor, old_floor, accepted in ((14, 7, (8, 9, 10, 99)), (13, 15, (16, 17, 99))):
            with self.subTest(minor=minor, old_floor=old_floor):
                self.assertEqual(compatibility.validate_interpreter(synthetic_runtime(minor, old_floor), policy)["reason"],
                                 "python_security_floor")
            for patch_number in accepted:
                with self.subTest(minor=minor, patch=patch_number):
                    self.assertTrue(compatibility.validate_interpreter(synthetic_runtime(minor, patch_number), policy)["ok"])
        self.assertEqual(policy["pypdf"]["version"], "6.20.0")
        self.assertEqual(policy["calibre"]["version"], "9.15.0")

    def test_supplied_v102_version_fixtures_use_the_shipped_shared_evaluator(self):
        policy = compatibility.load_policy()
        fixtures = json.loads((ROOT / "docs/codex-v1.0.2/RUNTIME_POLICY_CASES.json").read_bytes())
        self.assertTrue(fixtures["not_runtime_configuration"])
        for case in fixtures["cases"]:
            record = synthetic_runtime()
            record["version"] = case["input"]
            if isinstance(case["input"], str) and re.fullmatch(r"(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)", case["input"]):
                record["version_info"] = list(map(int, case["input"].split(".")))
            with self.subTest(case=case["id"]):
                result = compatibility.validate_interpreter(record, policy)
                self.assertIs(result["ok"], case["expected_ok"])
                self.assertEqual(result["reason"], case["expected_reason"])
                if result["reason"] == "unsupported_python":
                    self.assertNotIn("insecure", result["message"].lower())
                    self.assertNotIn("security floor", result["message"].lower())

    def test_shipped_policy_version_float_and_boolean_tokens_fail_closed(self):
        policy = compatibility.load_policy()
        token = '"policy_version": ' + str(policy["policy_version"])
        original = json.dumps(policy)
        for value in (str(policy["policy_version"]) + ".0", "true", "false"):
            text = original.replace(token, '"policy_version": ' + value, 1)
            self.assertNotEqual(text, original)
            self.path.write_text(text, encoding="utf-8")
            with self.subTest(value=value), self.assertRaises(compatibility.PolicyError):
                compatibility.load_policy(self.path)

    def test_missing_empty_oversized_invalid_utf8_and_recursive_json_are_rejected(self):
        with self.assertRaises(compatibility.PolicyError):
            compatibility.load_policy(self.path)
        for data in (b"", b" " * (compatibility.POLICY_MAX_BYTES + 1), b"\xff", b"[" * 1100 + b"]" * 1100):
            self.path.write_bytes(data)
            with self.subTest(data_length=len(data)), self.assertRaises(compatibility.PolicyError):
                compatibility.load_policy(self.path)
        with self.assertRaises(compatibility.PolicyError):
            compatibility.load_policy(42)

    def test_duplicate_top_and_nested_keys_fail_closed(self):
        text = json.dumps(synthetic_policy())
        for damaged in (text.replace('"schema_version": 1', '"schema_version": 1, "schema_version": 1', 1),
                        text.replace('"major": 3', '"major": 3, "major": 3', 1)):
            self.path.write_text(damaged, encoding="utf-8")
            with self.assertRaises(compatibility.PolicyError):
                compatibility.load_policy(self.path)

    def test_nonfinite_json_numbers_are_rejected(self):
        text = json.dumps(synthetic_policy())
        for value in ("NaN", "Infinity", "-Infinity"):
            self.path.write_text(text.replace('"policy_version": 1', '"policy_version": ' + value), encoding="utf-8")
            with self.assertRaises(compatibility.PolicyError):
                compatibility.load_policy(self.path)

    def test_schema_unknown_keys_missing_keys_and_bad_dates_fail_closed(self):
        for schema in (True, 1.0, 2, "1"):
            policy = synthetic_policy()
            policy["schema_version"] = schema
            self.assert_rejected(policy)
        for field in synthetic_policy():
            policy = synthetic_policy()
            del policy[field]
            self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["accept_any_python"] = True
        self.assert_rejected(policy)
        for reviewed in ("2026-02-30", "20261011", True):
            policy = synthetic_policy()
            policy["reviewed_on"] = reviewed
            self.assert_rejected(policy)

    def test_policy_integer_boolean_and_identity_mutations_fail_closed(self):
        for field in ("major", "minor", "compatibility_min_patch", "security_min_patch", "recommended_patch"):
            for value in (True, -1, 8.0, "8"):
                policy = synthetic_policy()
                policy["python"]["branches"][0][field] = value
                self.assert_rejected(policy)
        for field, value in (("implementation", "pypy"), ("platform", "linux"), ("machine", "ARM64"),
                             ("bits", True), ("stable_only", 1), ("gil_disabled", 0)):
            policy = synthetic_policy()
            policy["python"][field] = value
            self.assert_rejected(policy)

    def test_duplicate_or_contradictory_branches_fail_closed(self):
        policy = synthetic_policy()
        policy["python"]["branches"].append(copy.deepcopy(policy["python"]["branches"][0]))
        self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["excluded_python_branches"].append({"major": 3, "minor": 14, "reason": "contradiction"})
        self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["excluded_python_branches"].append(copy.deepcopy(policy["excluded_python_branches"][0]))
        self.assert_rejected(policy)

    def test_invalid_blocked_patches_and_recommendation_fail_closed(self):
        for blocked in ([True], [9, 9], [-1], list(range(33))):
            policy = synthetic_policy()
            policy["python"]["branches"][0]["blocked_patches"] = blocked
            self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["python"]["branches"][0]["recommended_patch"] = 6
        self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["python"]["branches"][0]["blocked_patches"] = [8]
        self.assert_rejected(policy)

    def test_missing_duplicate_and_unknown_source_references_fail_closed(self):
        for references in ([], ["Missing"], ["SyntheticSecurity", "SyntheticSecurity"], [True]):
            policy = synthetic_policy()
            policy["python"]["branches"][0]["security_sources"] = references
            self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["sources"].append(copy.deepcopy(policy["sources"][0]))
        self.assert_rejected(policy)
        for url in ("http://www.python.org/", "https://user:secret@www.python.org/", "https://", "https://example.org/a b"):
            policy = synthetic_policy()
            policy["sources"][0]["url"] = url
            self.assert_rejected(policy)

    def test_empty_credentials_and_invalid_source_ports_fail_closed(self):
        for url in ("https://@www.python.org/", "https://:@www.python.org/",
                    "https://www.python.org:bad/", "https://www.python.org:99999/",
                    "https://www.python.org:65536/", "https://www.python.org:-1/"):
            with self.subTest(url=url):
                policy = synthetic_policy()
                policy["sources"][0]["url"] = url
                self.assert_rejected(policy)
        for port in (0, 443, 65535):
            with self.subTest(valid_port=port):
                policy = synthetic_policy()
                policy["sources"][0]["url"] = "https://www.python.org:" + str(port) + "/"
                self.write(policy)
                self.assertEqual(compatibility.load_policy(self.path), policy)

    def test_package_lock_and_converter_versions_remain_controlled(self):
        for field, value in (("version", "6.20"), ("wheel_sha256", "a" * 63), ("wheel_sha256", "A" * 64)):
            policy = synthetic_policy()
            policy["pypdf"][field] = value
            self.assert_rejected(policy)
        policy = synthetic_policy()
        policy["calibre"]["version"] = "9.15"
        self.assert_rejected(policy)

    def test_import_and_validation_do_not_require_pypdf(self):
        imports = []
        original_import = builtins.__import__

        def guarded_import(name, *args, **kwargs):
            if name == "pypdf" or name.startswith("pypdf."):
                imports.append(name)
                raise AssertionError("Policy validation attempted a package import")
            return original_import(name, *args, **kwargs)

        with patch("builtins.__import__", guarded_import):
            isolated_spec = importlib.util.spec_from_file_location("wbs_policy_no_pypdf", MODULE)
            module = importlib.util.module_from_spec(isolated_spec)
            isolated_spec.loader.exec_module(module)
            self.assertTrue(module.validate_interpreter(synthetic_runtime(), synthetic_policy())["ok"])
        self.assertEqual(imports, [])


class PackageProbeTests(unittest.TestCase):
    """Package-boundary failures use explicitly synthetic identities/modules."""

    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-package-probe-unit-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)
        self.policy_path = self.work / "policy.json"
        self.policy_path.write_text(json.dumps(synthetic_policy()), encoding="utf-8")
        self.packages = self.work / "site-packages"
        package = self.packages / "pypdf"
        package.mkdir(parents=True)
        (package / "__init__.py").write_text("# Synthetic package-origin fixture", encoding="utf-8")
        dist_info = self.packages / "pypdf-6.20.0.dist-info"
        dist_info.mkdir()
        (dist_info / "METADATA").write_text("Name: pypdf\nVersion: 6.20.0\n", encoding="utf-8")
        self.package = SimpleNamespace(__file__=str(package / "__init__.py"), __version__="6.20.0")
        self.distribution = SimpleNamespace(version="6.20.0", metadata={"Name": "pypdf"},
            files=[Path("pypdf-6.20.0.dist-info/METADATA")], locate_file=lambda item: self.packages / str(item))
        self.identity = synthetic_runtime()
        self.identity.update(protocol="winbooksplit.runtime", schema_version=1, ok=False,
                             policy_version=0, policy_sha256="", reason="malformed_runtime", message="Synthetic identity")

    def probe(self, require_pypdf=True):
        with patch.object(compatibility, "_runtime_identity", return_value=copy.deepcopy(self.identity)), \
                patch.object(compatibility.importlib, "import_module", return_value=self.package), \
                patch.object(compatibility.metadata, "distribution", return_value=self.distribution), \
                patch.object(compatibility.sysconfig, "get_paths", return_value={"purelib": str(self.packages), "platlib": str(self.packages)}), \
                patch.dict(sys.modules, {"pypdf": self.package}):
            # A full-suite invocation may already have real pypdf children.
            # Keep only this test's deliberate synthetic child during its
            # mocked package probe; patch.dict restores the original modules.
            for name in tuple(sys.modules):
                if name.startswith("pypdf.") and name != "pypdf.synthetic_child":
                    del sys.modules[name]
            return compatibility.probe_interpreter(require_pypdf, self.policy_path)

    def test_interpreter_only_has_no_package_import_or_package_fields(self):
        with patch.object(compatibility, "_runtime_identity", return_value=copy.deepcopy(self.identity)), \
                patch.object(compatibility.importlib, "import_module", side_effect=AssertionError("pypdf import forbidden")) as importer:
            result = compatibility.probe_interpreter(False, self.policy_path)
        self.assertTrue(result["ok"])
        self.assertEqual(result["reason"], "supported")
        self.assertEqual(result["policy_version"], 1)
        self.assertRegex(result["policy_sha256"], r"\A[0-9a-f]{64}\Z")
        self.assertNotIn("pypdf_path", result)
        self.assertNotIn("pypdf_version", result)
        importer.assert_not_called()

    def test_unsupported_interpreter_precedes_any_package_import(self):
        self.identity.update(version="3.14.2", version_info=[3, 14, 2])
        with patch.object(compatibility, "_runtime_identity", return_value=copy.deepcopy(self.identity)), \
                patch.object(compatibility.importlib, "import_module", side_effect=AssertionError("pypdf import forbidden")) as importer:
            result = compatibility.probe_interpreter(True, self.policy_path)
        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "python_security_floor")
        importer.assert_not_called()

    def test_missing_locked_package_is_separate_from_interpreter_acceptance(self):
        with patch.object(compatibility, "_runtime_identity", return_value=copy.deepcopy(self.identity)), \
                patch.object(compatibility.importlib, "import_module", side_effect=ModuleNotFoundError("Synthetic missing pypdf")):
            result = compatibility.probe_interpreter(True, self.policy_path)
        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "missing_pypdf")

    def test_full_package_probe_verifies_import_and_distribution_versions(self):
        result = self.probe()
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["pypdf_version"], "6.20.0")
        self.assertEqual(Path(result["pypdf_path"]), Path(self.package.__file__))
        self.package.__version__ = "6.19.0"
        self.assertEqual(self.probe()["reason"], "pypdf_version_mismatch")
        self.package.__version__ = "6.20.0"
        self.distribution.version = "6.19.0"
        self.assertEqual(self.probe()["reason"], "pypdf_version_mismatch")

    def test_package_file_outside_interpreter_roots_is_rejected(self):
        outside = self.work / "outside.py"
        outside.write_text("# Synthetic wrong origin", encoding="utf-8")
        self.package.__file__ = str(outside)
        self.assertEqual(self.probe()["reason"], "pypdf_origin_invalid")

    def test_distribution_root_and_metadata_origin_are_checked(self):
        self.distribution.locate_file = lambda item: self.work / str(item)
        self.assertEqual(self.probe()["reason"], "pypdf_origin_invalid")
        self.distribution.locate_file = lambda item: self.packages / str(item)
        self.distribution.files = None
        self.assertEqual(self.probe()["reason"], "pypdf_origin_invalid")
        self.distribution.files = [Path("pypdf-6.20.0.dist-info/METADATA")]
        self.distribution.metadata = {"Name": "some-other-project"}
        self.assertEqual(self.probe()["reason"], "pypdf_origin_invalid")

    def test_every_imported_package_submodule_origin_is_checked(self):
        outside = self.work / "outside.py"
        outside.write_text("# Synthetic shadow child", encoding="utf-8")
        with patch.dict(sys.modules, {"pypdf.synthetic_child": SimpleNamespace(__file__=str(outside))}):
            self.assertEqual(self.probe()["reason"], "pypdf_origin_invalid")
        with patch.dict(sys.modules, {"pypdf.synthetic_child": SimpleNamespace()}):
            self.assertEqual(self.probe()["reason"], "pypdf_origin_invalid")

    def test_bad_or_missing_policy_precedes_package_import(self):
        self.policy_path.write_text('{"schema_version": true}', encoding="utf-8")
        with patch.object(compatibility, "_runtime_identity", return_value=copy.deepcopy(self.identity)), \
                patch.object(compatibility.importlib, "import_module", side_effect=AssertionError("pypdf import forbidden")) as importer:
            result = compatibility.probe_interpreter(True, self.policy_path)
            self.assertEqual(result["reason"], "malformed_policy")
            result = compatibility.probe_interpreter(True, self.work / "missing.json")
            self.assertEqual(result["reason"], "malformed_policy")
        importer.assert_not_called()

    def test_nonboolean_probe_option_does_not_skip_package_guard(self):
        result = self.probe(require_pypdf=0)
        self.assertFalse(result["ok"])
        self.assertEqual(result["reason"], "malformed_runtime")


class NativeProbeEntrypointTests(unittest.TestCase):
    """Actual current-interpreter subprocess checks; no synthetic native pass."""

    def setUp(self):
        self.owned = tempfile.TemporaryDirectory(prefix="wbs-probe-cli-unit-")
        self.addCleanup(self.owned.cleanup)
        self.work = Path(self.owned.name)

    def run_cli(self, path, arguments, no_site=False):
        return subprocess.run([sys.executable, "-I", "-B", *( ["-S"] if no_site else []),
                               str(path), *arguments], cwd=self.work,
                              capture_output=True, text=True, encoding="utf-8", timeout=30)

    def one_record(self, child):
        self.assertEqual(child.stderr, "", child.stderr)
        lines = child.stdout.splitlines()
        self.assertEqual(len(lines), 1, child.stdout)
        return json.loads(lines[0])

    def test_actual_interpreter_only_probe_has_one_record_with_site_disabled(self):
        child = self.run_cli(MODULE, ["--probe", "--interpreter-only", "--policy",
                                    str(ROOT / "engine/WinBookSplit.Compatibility.json")], no_site=True)
        record = self.one_record(child)
        self.assertEqual(record["version_info"], list(sys.version_info[:3]))
        self.assertEqual(record["protocol"], "winbooksplit.runtime")
        self.assertNotIn("pypdf_version", record)
        self.assertNotIn("pypdf_path", record)
        decision = compatibility.validate_interpreter(record, compatibility.load_policy())
        self.assertEqual(record["ok"], decision["ok"])
        self.assertEqual(child.returncode, 0 if decision["ok"] else 1)

    def test_actual_processing_probe_cannot_pass_without_package(self):
        child = self.run_cli(MODULE, ["--probe"], no_site=True)
        record = self.one_record(child)
        decision = compatibility.validate_interpreter(record, compatibility.load_policy())
        self.assertFalse(record["ok"])
        self.assertEqual(record["reason"], "missing_pypdf" if decision["ok"] else decision["reason"])
        self.assertEqual(child.returncode, 1)

    def prepare_direct_app(self, policy_bytes):
        app = self.work / "app"
        app.mkdir()
        engine_dir = app / "engine"
        engine_dir.mkdir()
        for name in ("winbooksplit_engine.py", "winbooksplit_compatibility.py", "WinBookSplit.Outcomes.json"):
            shutil.copyfile(ROOT / "engine" / name, engine_dir / name)
        shutil.copyfile(ROOT / "VERSION", app / "VERSION")
        if policy_bytes is not None:
            (engine_dir / "WinBookSplit.Compatibility.json").write_bytes(policy_bytes)
        source = self.work / "authored-input.pdf"
        source.write_bytes(b"Authored preservation fixture; not a private document")
        output = self.work / "output"
        output.mkdir()
        neighbor = output / "neighbor.txt"
        neighbor.write_bytes(b"Preserve authored neighbor")
        return engine_dir / "winbooksplit_engine.py", source, output, neighbor

    def assert_direct_rejection_preserves_inputs(self, policy_bytes, reason, code):
        engine, source, output, neighbor = self.prepare_direct_app(policy_bytes)
        source_before, neighbor_before = source.read_bytes(), neighbor.read_bytes()
        child = self.run_cli(engine, [str(source), str(output), "manual", "1"], no_site=True)
        record = self.one_record(child)
        self.assertEqual(child.returncode, 3)
        self.assertEqual(record["protocol"], "winbooksplit.result")
        self.assertEqual(record["code"], code)
        self.assertEqual(record["diagnostic"]["reason"], reason)
        self.assertEqual(record["written_count"], 0)
        self.assertIsNone(record["execution"])
        self.assertEqual(source.read_bytes(), source_before)
        self.assertEqual(neighbor.read_bytes(), neighbor_before)
        self.assertEqual(list(output.iterdir()), [neighbor])

    def test_direct_engine_rejects_malformed_policy_before_package_or_document(self):
        self.assert_direct_rejection_preserves_inputs(b'{"schema_version": true}', "malformed_policy", "runtime_invalid")

    def test_direct_engine_rejects_missing_policy_before_package_or_document(self):
        self.assert_direct_rejection_preserves_inputs(None, "malformed_policy", "runtime_invalid")

    def test_direct_engine_missing_package_preserves_existing_outcome_contract(self):
        identity = compatibility._runtime_identity()
        decision = compatibility.validate_interpreter(identity, compatibility.load_policy())
        self.assert_direct_rejection_preserves_inputs(
            (ROOT / "engine/WinBookSplit.Compatibility.json").read_bytes(),
            "missing_pypdf" if decision["ok"] else decision["reason"],
            "dependency_missing" if decision["ok"] else "runtime_invalid")

    def assert_direct_shadow_package_rejection(self, package_version):
        engine, source, output, neighbor = self.prepare_direct_app(
            (ROOT / "engine/WinBookSplit.Compatibility.json").read_bytes())
        source_before, neighbor_before = source.read_bytes(), neighbor.read_bytes()
        shadow = self.work / "controlled-package-shadow"
        (shadow / "pypdf").mkdir(parents=True)
        (shadow / "pypdf/__init__.py").write_text(
            "# Authored package-origin fixture\n__version__=" + repr(package_version) + "\n", encoding="utf-8")
        # This deliberate subprocess path injection exercises the guard. Normal
        # processing retains -I and never adds this synthetic shadow directory.
        wrapper = "import runpy,sys; sys.path.insert(0,sys.argv[1]); target=sys.argv[2]; sys.argv=sys.argv[2:]; runpy.run_path(target,run_name='__main__')"
        child = subprocess.run([sys.executable, "-I", "-B", "-c", wrapper, str(shadow), str(engine),
                                str(source), str(output), "manual", "1"], cwd=self.work,
                               capture_output=True, text=True, encoding="utf-8", timeout=30)
        record = self.one_record(child)
        decision = compatibility.validate_interpreter(compatibility._runtime_identity(), compatibility.load_policy())
        try:
            compatibility.metadata.distribution("pypdf")
            metadata_present = True
        except compatibility.metadata.PackageNotFoundError:
            metadata_present = False
        if not decision["ok"]:
            reason, code = decision["reason"], "runtime_invalid"
        elif not metadata_present:
            reason, code = "missing_pypdf", "dependency_missing"
        else:
            reason = "pypdf_origin_invalid" if package_version == "6.20.0" else "pypdf_version_mismatch"
            code = "runtime_invalid"
        self.assertEqual(child.returncode, 3)
        self.assertEqual(record["diagnostic"]["reason"], reason)
        self.assertEqual(record["code"], code)
        self.assertEqual(record["written_count"], 0)
        self.assertIsNone(record["execution"])
        self.assertEqual(source.read_bytes(), source_before)
        self.assertEqual(neighbor.read_bytes(), neighbor_before)
        self.assertEqual(list(output.iterdir()), [neighbor])

    def test_direct_engine_rejects_controlled_wrong_package_origin_before_document(self):
        self.assert_direct_shadow_package_rejection("6.20.0")

    def test_direct_engine_rejects_controlled_wrong_package_version_before_document(self):
        self.assert_direct_shadow_package_rejection("6.19.0")


if __name__ == "__main__":
    unittest.main()
