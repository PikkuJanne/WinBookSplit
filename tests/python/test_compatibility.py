"""Synthetic policy boundary tests; these are not native version-matrix passes."""

import builtins
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
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
        for minor in (11, 12, 15):
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

    def test_tuple_and_list_version_values_are_compared_numerically(self):
        for version in ((3, 14, 10), [3, 14, 10]):
            record = synthetic_runtime(14, 10)
            record["version"] = version
            self.assert_reason(record, "supported")

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


if __name__ == "__main__":
    unittest.main()
