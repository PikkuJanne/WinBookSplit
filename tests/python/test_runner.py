"""Harness regressions only; these are not corrected-engine acceptance tests."""

import importlib.util
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "tests/run_tests.py"
SPEC = importlib.util.spec_from_file_location("wbs_test_runner", SCRIPT)
runner = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(runner)


class RunnerTests(unittest.TestCase):
    def test_ac009_python_failure_stays_failed_after_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-runner-probe-") as directory:
            work = Path(directory)
            neighbor = work / "neighbor.txt"
            neighbor.write_bytes(b"unrelated synthetic neighbor")
            report_path = work / "python.json"
            result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--layer", "python", "--failure-probe", "python",
                                         "--report", str(report_path)], work)
            self.assertEqual(result["exit_code"], 1, result)
            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertFalse(report["success"])
            self.assertEqual(report["steps"][0]["exit_code"], 1)
            self.assertEqual(report["steps"][-1]["exit_code"], 0)
            self.assertIn("Intentional synthetic Python failure", report["steps"][0]["stderr"])
            self.assertTrue(report["owned_work_directory_removed"])
            self.assertTrue(report["source"]["source_unchanged"])
            self.assertEqual(neighbor.read_bytes(), b"unrelated synthetic neighbor")

    @unittest.skipUnless(sys.platform == "win32", "Native cmd.exe probe requires Windows")
    def test_ac009_native_failure_stays_failed_after_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-native-probe-") as directory:
            path = Path(directory) / "native.json"
            result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--layer", "python", "--failure-probe", "native",
                                         "--report", str(path)], Path(directory))
            self.assertEqual(result["exit_code"], 1, result)
            report = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(report["steps"][0]["exit_code"], 23)
            self.assertEqual(report["steps"][-1]["exit_code"], 0)
            self.assertFalse(report["success"])

    def test_ac010_repeat_probe_from_unrelated_directory(self):
        before = runner.source_manifest()
        with tempfile.TemporaryDirectory(prefix="wbs-repeat-probe-") as directory:
            work = Path(directory)
            # A hostile same-name module cannot be imported by -I or our explicit loads.
            (work / "json.py").write_text("raise RuntimeError('CWD module imported')", encoding="utf-8")
            observations = []
            for index in range(2):
                path = work / f"repeat-{index}.json"
                result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                             "--layer", "python", "--failure-probe", "python",
                                             "--report", str(path)], work)
                self.assertEqual(result["exit_code"], 1, result)
                report = json.loads(path.read_text(encoding="utf-8"))
                self.assertTrue(report["owned_work_directory_removed"])
                self.assertFalse(Path(report["owned_work_directory"]).exists())
                self.assertTrue(report["synthetic_input_unchanged"])
                observations.append(([step["exit_code"] for step in report["steps"]],
                                     report["source"]["tested_paths_digest"]))
            self.assertEqual(observations[0], observations[1])
            self.assertTrue((work / "json.py").exists())
        self.assertEqual(before, runner.source_manifest())

    def test_drain_both_large_process_streams(self):
        with tempfile.TemporaryDirectory(prefix="wbs-pipes-") as directory:
            result = runner.run_command([sys.executable, "-I", "-B", "-c",
                                         "import sys; sys.stdout.write('o'*200000); "
                                         "sys.stderr.write('e'*200000); sys.exit(19)"],
                                        Path(directory), timeout=15)
            self.assertEqual(result["exit_code"], 19)
            self.assertEqual(len(result["stdout"]), 200000)
            self.assertEqual(len(result["stderr"]), 200000)
            self.assertFalse(result["timed_out"])

    def test_timeout_is_nonzero_and_keeps_partial_streams(self):
        with tempfile.TemporaryDirectory(prefix="wbs-timeout-") as directory:
            result = runner.run_command([sys.executable, "-I", "-B", "-c",
                                         "import time; print('before timeout',flush=True); time.sleep(3)"],
                                        Path(directory), timeout=0.5)
            self.assertEqual(result["exit_code"], 124)
            self.assertTrue(result["timed_out"])
            self.assertIn("before timeout", result["stdout"])

    def test_missing_command_is_nonzero(self):
        with tempfile.TemporaryDirectory(prefix="wbs-missing-") as directory:
            result = runner.run_command([str(Path(directory) / "missing.exe")], Path(directory))
            self.assertEqual(result["exit_code"], 125)
            self.assertTrue(result["stderr"])

    def test_evidence_missing_or_invalid_cannot_mask_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-evidence-") as directory:
            path = Path(directory) / "missing.json"
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 126)
            path.write_text("[]", encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 126)
            step = {"exit_code": 23}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 23)
            path.write_text('{"schema_version":1,"success":false,"exit_code":1}', encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "shell")
            self.assertEqual(step["exit_code"], 126)
            for payload in ({}, {"schema_version": 1, "success": True, "exit_code": 0}):
                path.write_text(json.dumps(payload), encoding="utf-8")
                step = {"exit_code": 0}
                runner.attach_child_report(step, path, "shell")
                self.assertEqual(step["exit_code"], 126)
            for kind, payload in (
                ("shell", {"schema_version": 1, "success": True, "exit_code": 0,
                           "pester": {"result": "Passed", "total": 7}}),
                ("baseline", {"schema_version": 1, "result": "ORIGINAL_BEHAVIOR_REPRODUCED",
                              "engine_case_count": 14}),
            ):
                path.write_text(json.dumps(payload), encoding="utf-8")
                step = {"exit_code": 0}
                runner.attach_child_report(step, path, kind)
                self.assertEqual(step["exit_code"], 0)
                self.assertIn("evidence_sha256", step)

    def test_extraction_count_without_actual_cases_cannot_mask_success(self):
        with tempfile.TemporaryDirectory(prefix="wbs-extraction-evidence-") as directory:
            path = Path(directory) / "incomplete.json"
            path.write_text(json.dumps({"schema_version": 1, "success": True, "exit_code": 0,
                                        "result": "EXTRACTION_EQUIVALENCE_REPRODUCED", "engine_case_count": 14,
                                        "source_unchanged": True, "import_observation": {"import_safe": True}}),
                            encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "extraction")
            self.assertEqual(step["exit_code"], 126)

    def test_existing_report_is_preserved(self):
        with tempfile.TemporaryDirectory(prefix="wbs-no-overwrite-") as directory:
            path = Path(directory) / "existing.json"
            path.write_bytes(b"preexisting evidence")
            result = runner.run_command([sys.executable, "-I", "-B", str(SCRIPT),
                                         "--layer", "python", "--failure-probe", "python",
                                         "--report", str(path)], Path(directory))
            self.assertEqual(result["exit_code"], 1)
            self.assertEqual(path.read_bytes(), b"preexisting evidence")

    def test_reports_require_absolute_external_paths(self):
        with self.assertRaises(ValueError):
            runner.new_external_path(Path("relative.json"))
        with self.assertRaises(ValueError):
            runner.new_external_path(ROOT / "must-not-create.json")
        self.assertFalse((ROOT / "must-not-create.json").exists())

    def test_shell_bootstrap_passes_paths_as_data(self):
        with tempfile.TemporaryDirectory(prefix="wbs-shell-data-") as directory:
            work = Path(directory) / "quote' [space] å"
            work.mkdir()
            with patch.dict(os.environ, {"PSModulePath": "untrusted foreign modules"}):
                command, environment = runner.shell_command(Path("C:/trusted/pwsh.exe"),
                                                           work, work / "evidence.json", work, False)
            self.assertEqual(environment["WBS_TEST_WORK"], str(work))
            self.assertEqual(environment["PSModulePath"], str(Path("C:/trusted/Modules")))
            self.assertNotIn(str(work), " ".join(command))
            self.assertIn("-EncodedCommand", command)
            self.assertEqual(command[command.index("-ExecutionPolicy") + 1], "RemoteSigned")
            self.assertNotIn("Bypass", command)

    def test_git_provenance_drops_inherited_repository_selectors(self):
        with patch.dict(os.environ, {name: "hostile-unrelated-value" for name in runner.GIT_SELECTORS}):
            environment = runner.child_environment()
            self.assertFalse(any(name in environment for name in runner.GIT_SELECTORS))
            self.assertTrue(all(name in os.environ for name in runner.GIT_SELECTORS))


if __name__ == "__main__":
    unittest.main(verbosity=2)
