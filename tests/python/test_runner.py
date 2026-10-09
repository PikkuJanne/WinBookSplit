"""Harness regressions only; these are not corrected-engine acceptance tests."""

import importlib.util
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from types import SimpleNamespace
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

    def test_manual_evidence_requires_complete_target_and_preservation_results(self):
        complete = {
            "schema_version": 1, "task_id": "M1-T02", "result": "MANUAL_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": [f"AC-{number:03}" for number in range(13, 19)],
            "engine_case_count": 22,
            "engine_cases": [{"oracle_id": f"MAN-{number:02}", "passed": True} for number in range(1, 23)],
            "extra_cli_cases": [{"id": name, "passed": True} for name in sorted(runner.MANUAL_EXTRA_IDS)],
            "sampled_writer_cases": [{"index": number, "outputs": [{"filename": "synthetic.pdf"}]}
                                     for number in range(25)],
            "historical_bookmark_cases": [{"oracle_id": "BM-03", "equivalent": True}],
            "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
            "import_observation": {"import_safe": True},
            "seeded_property_cases": {"count": 250, "passed": True, "seed": 20261009},
        }
        with tempfile.TemporaryDirectory(prefix="wbs-manual-evidence-") as directory:
            path = Path(directory) / "manual.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "manual")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            step = {"exit_code": 0, "requested_shell_paths": ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]}
            runner.attach_child_report(step, path, "manual")
            self.assertEqual(step["exit_code"], 126, "Core-only report cannot pass requested host probes")
            self.assertNotIn("evidence", step)

            defects = {
                "boolean schema version": {"schema_version": True},
                "wrong task": {"task_id": "M1-T01"},
                "wrong result": {"result": "EXTRACTION_EQUIVALENCE_REPRODUCED"},
                "wrong case count": {"engine_case_count": 21},
                "missing cases despite correct count": {"engine_cases": []},
                "duplicate oracle IDs": {"engine_cases": complete["engine_cases"][:-1]
                                         + [{"oracle_id": "MAN-01", "passed": True}]},
                "unknown oracle IDs": {"engine_cases": complete["engine_cases"][:-1]
                                       + [{"oracle_id": "MAN-99", "passed": True}]},
                "invalid oracle ID type": {"engine_cases": complete["engine_cases"][:-1]
                                           + [{"oracle_id": [], "passed": True}]},
                "failed target case": {"engine_cases": complete["engine_cases"][:-1]
                                      + [{"oracle_id": "MAN-22", "passed": False}]},
                "missing extra CLI cases": {"extra_cli_cases": []},
                "duplicate extra CLI IDs": {"extra_cli_cases": complete["extra_cli_cases"][:-1]
                                             + [complete["extra_cli_cases"][0]]},
                "failed extra CLI case": {"extra_cli_cases": complete["extra_cli_cases"][:-1]
                                           + [{**complete["extra_cli_cases"][-1], "passed": False}]},
                "missing writer samples": {"sampled_writer_cases": []},
                "duplicate writer indexes": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                               + [complete["sampled_writer_cases"][0]]},
                "boolean writer index": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                          + [{"index": True, "outputs": [{"filename": "synthetic.pdf"}]}]},
                "zero writer files": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                       + [{"index": 24, "outputs": []}]},
                "boolean exit code": {"exit_code": False},
                "incomplete acceptance": {"acceptance_ids": ["AC-013"]},
                "missing bookmark observations": {"historical_bookmark_cases": []},
                "obsolete Level 1 known-bad comparisons": {"historical_bookmark_cases": [
                    {"oracle_id": f"BM-{number:02}", "equivalent": True} for number in range(1, 4)]},
                "wrong unchanged comparison": {"historical_bookmark_cases": [{"oracle_id": "BM-01", "equivalent": True}]},
                "changed bookmark behavior": {"historical_bookmark_cases": complete["historical_bookmark_cases"][:-1]
                                             + [{"oracle_id": "BM-03", "equivalent": False}]},
                "changed source": {"source_unchanged": False},
                "changed original guards": {"baseline_guards_preserved": False},
                "changed inputs": {"input_and_neighbor_unchanged": False},
                "missing temp cleanup": {"owned_temp_removed": False},
                "wrong immutable reference": {"immutable_original_commit": "88c2149b3b3034fbd0d7ef23c2382f4b01648e4c"},
                "unsafe import": {"import_observation": {"import_safe": False}},
                "invalid import observation": {"import_observation": []},
                "unrun properties": {"seeded_property_cases": {"count": 0, "passed": True, "seed": 20261009}},
                "failed properties": {"seeded_property_cases": {"count": 250, "passed": False, "seed": 20261009}},
                "missing reproducible seed": {"seeded_property_cases": {"count": 250, "passed": True}},
            }
            for description, changes in defects.items():
                with self.subTest(defect=description):
                    path.write_text(json.dumps({**complete, **changes}), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "manual")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    # Missing evidence must preserve a real nonzero child exit.
                    step = {"exit_code": 19}
                    runner.attach_child_report(step, path, "manual")
                    self.assertEqual(step["exit_code"], 19)

    def test_requested_manual_entrypoints_require_records_preservation_and_requested_hosts(self):
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        probes = [{"id": name, "exit_code": 0, "input_unchanged": True,
                   "cwd_engine_untouched": True, "owned_neighbor_unchanged": True,
                   "shell_executable": hosts[1] if "PS7" in name else hosts[0],
                   "outputs": [{"filename": "synthetic.pdf", "page_ids": [1, 2]}]}
                  for name in sorted(runner.MANUAL_ENTRYPOINT_IDS)]
        complete = {"probes": probes, "probe_count": 6, "parallel_launch_count": 3,
                    "shared_temp_engine_sentinel_unchanged": True, "owned_document_outputs_removed": True}
        runner.validate_manual_entrypoints({"entrypoints": complete}, hosts)
        malformed = {
            "count without records": {"probes": []},
            "duplicate IDs": {"probes": probes[:-1] + [probes[0]]},
            "failed child": {"probes": probes[:-1] + [{**probes[-1], "exit_code": 19}]},
            "boolean child exit": {"probes": probes[:-1] + [{**probes[-1], "exit_code": False}]},
            "changed input": {"probes": probes[:-1] + [{**probes[-1], "input_unchanged": False}]},
            "unrequested host": {"probes": probes[:-1] + [{**probes[-1], "shell_executable": "C:/other/pwsh.exe"}]},
            "only one host used": {"probes": [{**probe, "shell_executable": hosts[0]} for probe in probes]},
            "zero files": {"probes": probes[:-1] + [{**probes[-1], "outputs": []}]},
            "missing parallel coverage": {"parallel_launch_count": 0},
            "changed shared sentinel": {"shared_temp_engine_sentinel_unchanged": False},
            "missing cleanup": {"owned_document_outputs_removed": False},
        }
        for description, changes in malformed.items():
            with self.subTest(defect=description):
                with self.assertRaises(ValueError):
                    runner.validate_manual_entrypoints({"entrypoints": {**complete, **changes}}, hosts)
        with self.assertRaises(ValueError):
            runner.validate_manual_entrypoints({}, hosts)

    @unittest.skipUnless(sys.platform == "win32", "Windows host paths use case-insensitive comparison")
    def test_requested_manual_entrypoints_accept_windows_path_case_variants(self):
        hosts = [r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
                 r"C:\Program Files\PowerShell\7\pwsh.exe"]
        receipts = [r"C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe",
                    "c:/PROGRAM FILES/PowerShell/7/pwsh.exe"]
        probes = [{"id": name, "exit_code": 0, "input_unchanged": True,
                   "cwd_engine_untouched": True, "owned_neighbor_unchanged": True,
                   "shell_executable": receipts[1] if "PS7" in name else receipts[0],
                   "outputs": [{"filename": "synthetic.pdf", "page_ids": [1, 2]}]}
                  for name in sorted(runner.MANUAL_ENTRYPOINT_IDS)]
        complete = {"probes": probes, "probe_count": 6, "parallel_launch_count": 3,
                    "shared_temp_engine_sentinel_unchanged": True, "owned_document_outputs_removed": True}
        runner.validate_manual_entrypoints({"entrypoints": complete}, hosts)
        # Case normalization must not accept a different executable/location.
        unrequested = [{**probe, "shell_executable": r"C:\WINDOWS\System32\other.exe"}
                       if probe["id"] == "PS51-unrelated" else probe for probe in probes]
        with self.assertRaises(ValueError):
            runner.validate_manual_entrypoints({"entrypoints": {**complete, "probes": unrequested}}, hosts)
        # Repeating the same path with different casing is still only one host.
        with self.assertRaises(ValueError):
            runner.validate_manual_entrypoints({"entrypoints": complete}, [hosts[0], receipts[0]])

    def test_full_selects_manual_and_level1_regressions_and_forwards_hosts_only_to_manual(self):
        hosts = [Path("C:/trusted/ps51.exe"), Path("C:/trusted/pwsh.exe")]
        args = SimpleNamespace(layer="full", failure_probe=None, shell_path=hosts,
                               tool_root=Path("C:/trusted/tool-root"))
        commands = []

        def successful_command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git-identity\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic-test-source": "hash"}), \
                patch.object(runner, "run_command", side_effect=successful_command), \
                patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        manual = [command for command in commands
                  if str(ROOT / "tests/manual/characterize_manual.py") in command]
        self.assertEqual(len(manual), 1)
        self.assertEqual(manual[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertFalse(any(str(ROOT / "tests/extraction/characterize_extraction.py") in command
                             for command in commands))
        bookmarks = [command for command in commands
                     if str(ROOT / "tests/bookmarks/characterize_level1.py") in command]
        self.assertEqual(len(bookmarks), 1)
        self.assertNotIn("--shell-path", bookmarks[0])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["shell", "shell", "manual", "bookmarks"])
        self.assertEqual(report["steps"][-2]["name"], "manual-regression")
        self.assertEqual(report["steps"][-1]["name"], "level1-regression")
        self.assertEqual(len(report["steps"]), 5)
        self.assertTrue(report["success"])

    def test_bookmark_evidence_requires_complete_targets_normalization_and_preservation(self):
        complete = {
            "schema_version": 1, "task_id": "M1-T03", "result": "LEVEL1_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": ["AC-019", "AC-020", "AC-021", "AC-022"],
            "engine_case_count": 4,
            "engine_cases": [{"oracle_id": name, "passed": True} for name in ("BM-01", "BM-02", "BM-05", "BM-08")],
            "normalization_cases": [{"id": name, "passed": True} for name in (
                "invalid-destinations", "deep-lineage", "cyclic-outline", "malformed-outline", "traversal-limits")],
            "seeded_level1_cases": {"seed": 20261009, "count": 150, "passed": True},
            "sampled_writer_cases": [{"index": number, "outputs": [{"filename": "synthetic.pdf"}]}
                                     for number in range(10)],
            "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
            "import_observation": {"import_safe": True},
        }
        with tempfile.TemporaryDirectory(prefix="wbs-bookmark-evidence-") as directory:
            path = Path(directory) / "bookmarks.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "bookmarks")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            defects = {
                "boolean schema": {"schema_version": True},
                "wrong task": {"task_id": "M1-T02"},
                "wrong result": {"result": "MANUAL_REGRESSION_PASSED"},
                "claimed failure": {"success": False},
                "boolean exit": {"exit_code": False},
                "wrong case count": {"engine_case_count": 3},
                "count without records": {"engine_cases": []},
                "duplicate oracle IDs": {"engine_cases": complete["engine_cases"][:-1] + [complete["engine_cases"][0]]},
                "unknown oracle ID": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-03", "passed": True}]},
                "invalid oracle ID type": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": [], "passed": True}]},
                "failed target": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-08", "passed": False}]},
                "missing normalization": {"normalization_cases": []},
                "duplicate normalization ID": {"normalization_cases": complete["normalization_cases"][:-1]
                                               + [complete["normalization_cases"][0]]},
                "unknown normalization ID": {"normalization_cases": complete["normalization_cases"][:-1]
                                             + [{"id": "unexpected", "passed": True}]},
                "invalid normalization ID type": {"normalization_cases": complete["normalization_cases"][:-1]
                                                  + [{"id": [], "passed": True}]},
                "failed normalization": {"normalization_cases": complete["normalization_cases"][:-1]
                                         + [{**complete["normalization_cases"][-1], "passed": False}]},
                "wrong seed": {"seeded_level1_cases": {"seed": 0, "count": 150, "passed": True}},
                "wrong property count": {"seeded_level1_cases": {"seed": 20261009, "count": 0, "passed": True}},
                "failed properties": {"seeded_level1_cases": {"seed": 20261009, "count": 150, "passed": False}},
                "missing writer samples": {"sampled_writer_cases": []},
                "duplicate writer indexes": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                             + [complete["sampled_writer_cases"][0]]},
                "boolean writer index": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                         + [{"index": True, "outputs": [{"filename": "synthetic.pdf"}]}]},
                "zero output files": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                      + [{"index": 9, "outputs": []}]},
                "incomplete acceptance": {"acceptance_ids": ["AC-019"]},
                "changed source": {"source_unchanged": False},
                "changed guards": {"baseline_guards_preserved": False},
                "changed inputs or neighbors": {"input_and_neighbor_unchanged": False},
                "missing cleanup": {"owned_temp_removed": False},
                "wrong immutable reference": {"immutable_original_commit": "unknown"},
                "unsafe import": {"import_observation": {"import_safe": False}},
                "invalid import observation": {"import_observation": []},
            }
            for description, changes in defects.items():
                with self.subTest(defect=description):
                    path.write_text(json.dumps({**complete, **changes}), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "bookmarks")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    step = {"exit_code": 19}
                    runner.attach_child_report(step, path, "bookmarks")
                    self.assertEqual(step["exit_code"], 19)

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
