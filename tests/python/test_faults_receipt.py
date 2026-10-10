"""Synthetic M4-T03 envelope mutations; these do not certify native fault cases."""

from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_fault_envelope_test", ROOT / "tests/faults/validate_faults_report.py")
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)


def envelope():
    source = {name: "a" * 64 for name in (
        "engine/winbooksplit_engine.py", "engine/winbooksplit_job.py", "tests/run_tests.py",
        "tests/faults/characterize_faults.py", "tests/faults/validate_faults_report.py")}
    return {"schema_version": 1, "task_id": "M4-T03", "result": "FAULT_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": validator.ACCEPTANCE,
            "source_unchanged": True, "cleanup_safe": True, "owned_temp_removed": True,
            "requested_shell_paths": ["C:/synthetic/powershell.exe", "C:/synthetic/pwsh.exe"],
            "cross_host_error_scope": "Requires the same-source paths, process and outcomes parent stages",
            "tested_path_sha256": source,
            "sections": {name: {"source_unchanged": True, "tested_path_sha256": deepcopy(source)}
                         for name in validator.SECTIONS}}


class FaultEnvelopeTests(unittest.TestCase):
    def setUp(self):
        self.complete = envelope()
        self.sections = SimpleNamespace(validate_isolation_report=Mock(), validate_invariant_report=Mock(),
                                        validate_cross_host_report=Mock())
        self.loader = patch.object(validator, "load", return_value=self.sections)
        self.loader.start()
        self.addCleanup(self.loader.stop)

    def check(self, report, **kwargs):
        validator.validate_faults_report(report, self.complete["requested_shell_paths"], **kwargs)

    def test_valid_envelope_delegates_to_all_three_strict_section_validators(self):
        self.check(self.complete, expected_source_sha256=self.complete["tested_path_sha256"])
        self.sections.validate_isolation_report.assert_called_once_with(self.complete["sections"]["isolation"])
        self.sections.validate_invariant_report.assert_called_once_with(self.complete["sections"]["invariants"],
            expected_source_sha256=self.complete["tested_path_sha256"])
        self.sections.validate_cross_host_report.assert_called_once_with(self.complete["sections"]["cross_host"],
            self.complete["requested_shell_paths"])

    def test_failed_or_boolean_status_does_not_become_acceptance(self):
        for field, value in (("schema_version", True), ("task_id", "M4-T02"), ("result", "FAULT_REGRESSION_FAILED"),
                             ("success", False), ("exit_code", False), ("exit_code", 1), ("acceptance_ids", ["AC-073"]),
                             ("source_unchanged", False), ("cleanup_safe", False)):
            with self.subTest(field=field, value=value), self.assertRaises(ValueError):
                self.check({**self.complete, field: value})

    def test_both_requested_hosts_and_exact_source_scope_are_required(self):
        for field, value in (("requested_shell_paths", self.complete["requested_shell_paths"][:1]),
                             ("cross_host_error_scope", "all cross-host errors passed")):
            with self.subTest(field=field), self.assertRaises(ValueError):
                self.check({**self.complete, field: value})

    def test_missing_duplicate_or_unproved_section_fails(self):
        for name in validator.SECTIONS:
            changed = deepcopy(self.complete)
            del changed["sections"][name]
            with self.subTest(section=name), self.assertRaises(ValueError):
                self.check(changed)
            changed = deepcopy(self.complete)
            changed["sections"][name]["source_unchanged"] = False
            with self.subTest(section=name, failure="changed source"), self.assertRaises(ValueError):
                self.check(changed)
        with self.assertRaises(ValueError):
            self.check({**self.complete, "sections": {**self.complete["sections"], "other": {}}})

    def test_cross_checkpoint_section_map_is_rejected(self):
        for name in validator.SECTIONS:
            changed = deepcopy(self.complete)
            changed["sections"][name]["tested_path_sha256"]["tests/run_tests.py"] = "b" * 64
            with self.subTest(section=name), self.assertRaises(ValueError):
                self.check(changed)

    def test_unbound_parent_map_or_missing_runtime_path_is_rejected(self):
        different = {**self.complete["tested_path_sha256"], "tests/run_tests.py": "b" * 64}
        with self.assertRaises(ValueError):
            self.check(self.complete, expected_source_sha256=different)
        changed = deepcopy(self.complete)
        del changed["tested_path_sha256"]["engine/winbooksplit_engine.py"]
        with self.assertRaises(ValueError):
            self.check(changed)
        with self.assertRaises(ValueError):
            self.check({**self.complete, "tested_path_sha256": {"invalid": "g" * 64}})

    def test_final_cleanup_is_required_but_in_progress_validation_is_explicit(self):
        changed = {**self.complete, "owned_temp_removed": False}
        with self.assertRaises(ValueError):
            self.check(changed)
        self.check(changed, cleanup_complete=False)

    def test_failed_section_validator_cannot_be_hidden_by_passing_envelope(self):
        self.sections.validate_cross_host_report.side_effect = ValueError("actual native status disagrees")
        with self.assertRaisesRegex(ValueError, "native status"):
            self.check(self.complete)


class FaultFailureRetentionTests(unittest.TestCase):
    """Use the real receipt/ancestor path, with no application or child process."""

    def assert_failure_retained(self, exception, expected_exit):
        spec = importlib.util.spec_from_file_location(
            "wbs_fault_retention_driver", ROOT / "tests/faults/characterize_faults.py")
        driver = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(driver)
        runner = driver.runner
        real_temporary = tempfile.TemporaryDirectory
        last_case = {"id": "synthetic-last-attempt", "scope": "mocked-no-child-record", "complete": False}
        completed_cases = [{"id": "synthetic-earlier-case", "scope": "mocked-no-child-record"}]
        completed_sections = {"isolation": {"scope": "synthetic completed-section placeholder"}}
        with real_temporary(prefix="wbs-fault-retention-unit-") as unit_scope:
            root = Path(unit_scope).resolve()
            target = root / "failed.json"
            state, created, steps = {}, [], []

            def owned_temporary(*args, **kwargs):
                kwargs["dir"] = state.get("ancestor", root)
                temporary = real_temporary(*args, **kwargs)
                created.append(Path(temporary.name).resolve())
                return temporary

            def authored_failure(*args):
                driver.LAST_SECTION = "cross_host"
                driver.ACTIVE_SECTION = SimpleNamespace(LAST_CASE=last_case, COMPLETED_CASES=completed_cases)
                driver.COMPLETED_SECTIONS = completed_sections
                raise exception

            with patch.object(tempfile, "TemporaryDirectory", side_effect=owned_temporary), \
                    patch.object(runner, "source_manifest", return_value={"synthetic-source": "a" * 64}), \
                    patch.object(Path, "is_file", return_value=True), \
                    patch.object(driver, "characterize", side_effect=authored_failure):
                with runner.owned_evidence_workspace(steps) as ancestor:
                    state["ancestor"] = Path(ancestor)
                    argv = ["characterize_faults.py", "--report", str(target),
                            "--shell-path", "C:/synthetic/powershell.exe",
                            "--shell-path", "C:/synthetic/pwsh.exe"]
                    with patch.object(sys, "argv", argv), patch.object(sys, "stdout", Mock()), \
                            patch.object(sys, "stderr", Mock()):
                        actual_exit = driver.main()
                    failed = json.loads(target.read_text(encoding="utf-8"))
                    inner = created[-1]
                    self.assertEqual(actual_exit, expected_exit)
                    self.assertEqual(failed["exit_code"], expected_exit)
                    self.assertIs(type(failed["exit_code"]), int)
                    self.assertEqual(failed["error_type"], type(exception).__name__)
                    self.assertEqual(failed["last_section"], "cross_host")
                    self.assertEqual(failed["failed_section_last_case"], last_case)
                    self.assertEqual(failed["failed_section_completed_cases"], completed_cases)
                    self.assertEqual(failed["completed_sections"], completed_sections)
                    self.assertIs(failed["success"], False)
                    self.assertIs(failed["cleanup_safe"], False)
                    self.assertIs(failed["owned_temp_removed"], False)
                    self.assertEqual(Path(failed["workspace_retained"]), inner)
                    self.assertTrue(inner.is_dir())
                    step = {"exit_code": actual_exit, "stdout": "", "stderr": "Authored failure; no child launched"}
                    runner.attach_child_report(step, target, "faults")
                    steps.append(step)
                    self.assertEqual(step["failed_evidence"], failed)
                    self.assertIs(step["cleanup_safe"], False)
                    self.assertNotIn("evidence", step)
                self.assertTrue(Path(ancestor).is_dir())
                self.assertTrue(inner.is_dir())
                self.assertTrue(target.is_file())
            # The enclosing unit scope owns only empty directories and synthetic
            # JSON. No child was launched, so its ordinary cleanup is proved.

    def test_assertion_failure_preserves_partial_receipt_and_ancestor(self):
        self.assert_failure_retained(AssertionError("authored assertion"), 1)

    def test_key_failure_preserves_partial_receipt_and_ancestor(self):
        self.assert_failure_retained(KeyError("authored missing field"), 1)

    def test_keyboard_interrupt_preserves_partial_receipt_and_ancestor(self):
        self.assert_failure_retained(KeyboardInterrupt(), 130)

    def test_system_exit_preserves_native_code_partial_receipt_and_ancestor(self):
        self.assert_failure_retained(SystemExit(23), 23)

    def test_missing_and_malformed_fault_receipts_fail_closed_and_retain_ancestor(self):
        spec = importlib.util.spec_from_file_location(
            "wbs_fault_missing_driver", ROOT / "tests/faults/characterize_faults.py")
        driver = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(driver)
        runner = driver.runner
        for malformed in (False, True):
            steps = []
            with self.subTest(malformed=malformed), runner.owned_evidence_workspace(steps) as ancestor:
                path = Path(ancestor) / "child.json"
                if malformed:
                    path.write_bytes(b"{incomplete authored JSON")
                step = {"exit_code": 0, "stdout": "", "stderr": ""}
                runner.attach_child_report(step, path, "faults")
                steps.append(step)
                self.assertEqual(step["exit_code"], 126)
                self.assertIs(step["cleanup_safe"], False)
                self.assertNotIn("evidence", step)
            self.assertTrue(Path(ancestor).is_dir())
            # This unit authored the sole optional file and launched no child.
            if path.exists():
                path.unlink()
            Path(ancestor).rmdir()

    def test_unexpected_validator_key_error_retains_raw_child_and_ancestor(self):
        spec = importlib.util.spec_from_file_location(
            "wbs_fault_validator_failure_driver", ROOT / "tests/faults/characterize_faults.py")
        driver = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(driver)
        runner = driver.runner
        steps = []
        child = envelope()
        loader = SimpleNamespace(exec_module=Mock())
        broken_validator = SimpleNamespace(validate_faults_report=Mock(side_effect=KeyError("authored validator failure")))
        with runner.owned_evidence_workspace(steps) as ancestor:
            path = Path(ancestor) / "child.json"
            path.write_text(json.dumps(child), encoding="utf-8")
            step = {"exit_code": 0, "stdout": "", "stderr": ""}
            with patch.object(importlib.util, "spec_from_file_location", return_value=SimpleNamespace(loader=loader)), \
                    patch.object(importlib.util, "module_from_spec", return_value=broken_validator):
                runner.attach_child_report(step, path, "faults")
            steps.append(step)
            broken_validator.validate_faults_report.assert_called_once()
            self.assertEqual(step["exit_code"], 126)
            self.assertIs(step["cleanup_safe"], False)
            self.assertEqual(step["failed_evidence"], child)
            self.assertIn("authored validator failure", step["evidence_error"])
            self.assertNotIn("evidence", step)
        self.assertTrue(Path(ancestor).is_dir())
        # The authored sole JSON and empty ancestor are known; no child existed.
        path.unlink()
        Path(ancestor).rmdir()


if __name__ == "__main__":
    unittest.main()
