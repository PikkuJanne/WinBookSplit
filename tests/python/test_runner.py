"""Harness regressions only; these are not corrected-engine acceptance tests."""

import importlib.util
from copy import deepcopy
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


def corrected_level2_reference():
    ranges = [[0, 2], [2, 3], [3, 6], [6, 8], [8, 10], [10, 12]]
    titles = ["Front matter", "A - Opening pages", "A1", "A2", "B - Opening pages", "B1"]
    return {"oracle_id": "BM-03", "passed": True, "exit_code": 0,
            "expected_ranges": ranges, "titles": titles,
            "outputs": [{"filename": f"{index:02d} - {title}.pdf", "range": pair,
                         "page_ids": list(range(pair[0] + 1, pair[1] + 1))}
                        for index, (title, pair) in enumerate(zip(titles, ranges), 1)]}


def complete_plan_report():
    entries = [{"sequence": 1, "start": 0, "end": 1, "filename": "01 - Opening.pdf"},
               {"sequence": 2, "start": 1, "end": 3, "filename": "02 - Chapter.pdf"}]
    outputs = [{"filename": "01 - Opening.pdf", "range": [0, 1], "page_ids": [1]},
               {"filename": "02 - Chapter.pdf", "range": [1, 3], "page_ids": [2, 3]}]
    parity = {"passed": True, "total_pages": 3, "preview_entries": entries, "outputs": outputs,
              "coverage": {"complete": True, "covered_pages": 3, "section_count": 2}, "written_count": 2,
              "source_identity": {"path": "C:/synthetic/original.pdf", "sha256": "a" * 64,
                                  "size_bytes": 100, "binding": "reader_snapshot"},
              "neighbor_unchanged": True, "no_replanning_or_reopening": True,
              "original_page_content_sha256": ["c" * 64, "d" * 64, "e" * 64],
              "page_content_sha256": ["c" * 64, "d" * 64, "e" * 64]}
    report = {
        "schema_version": 1, "task_id": "M1-T05", "result": "SHARED_PLAN_REGRESSION_PASSED",
        "success": True, "exit_code": 0, "acceptance_ids": ["AC-027", "AC-028", "AC-029"],
        "structural_cases": [{"id": name, "passed": True, "rejected": True, "error_code": "invalid_plan"}
                             for name in ("gap", "overlap", "empty", "negative", "reversed", "overflow", "zero-pages", "noninteger")]
                            + [{"id": "valid-whole-document", "passed": True, "accepted": True}],
        "preview_cases": [{"id": "read-only-preview", "passed": True, "mutation_rejected": True, "chapter_files_written": 0},
                          {"id": "isolated-preview-data", "passed": True, "detached_from_input": True,
                           "unbound_execution_rejected": True, "chapter_files_written": 0}],
        "parity_cases": [{"mode": mode, **deepcopy(parity)} for mode in ("manual", "1", "2")],
        "source_cases": [{"id": name, **deepcopy(parity), "outcome": "bound_original_preserved",
                          "captured_sha256": "a" * 64, "replacement_sha256": "b" * 64, "replacement_unchanged": True,
                          "original_page_content_sha256": ["c" * 64, "d" * 64, "e" * 64],
                          "page_content_sha256": ["c" * 64, "d" * 64, "e" * 64]}
                         for name in ("changed-same-path", "repointed-source")],
        "seeded_plan_cases": {"seed": 20261009, "count": 300, "per_mode": {"manual": 100, "1": 100, "2": 100}, "passed": True},
        "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
        "owned_temp_removed": True, "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
        "import_observation": {"import_safe": True},
    }
    for case in report["parity_cases"] + report["source_cases"]:
        case["writer_result"] = {"mode": case.get("mode", "manual"), "total_pages": 3,
                                 "written_count": 2, "coverage": deepcopy(case["coverage"]),
                                 "source_identity": deepcopy(case["source_identity"]),
                                 "outputs": [{**entry, "page_count": entry["end"] - entry["start"]}
                                             for entry in case["preview_entries"]]}
    repointed = report["source_cases"][1]
    repointed["deleted_path_snapshot"] = {field: deepcopy(repointed[field])
                                         for field in (*parity, "writer_result", "original_page_content_sha256", "page_content_sha256")}
    return report


def complete_diagnostic_report(hosts):
    parities = {case["mode"]: case for case in complete_plan_report()["parity_cases"]}
    unicode_parity = parities["1"]
    unicode_ranges = [(0, 3), (3, 6), (6, 10)]
    names = ["01 - Front matter.pdf", "02 - 章节 å.pdf", "03 - 次章 é.pdf"]
    entries = [{"sequence": index, "start": start, "end": end, "filename": filename}
               for index, ((start, end), filename) in enumerate(zip(unicode_ranges, names), 1)]
    unicode_parity.update(total_pages=10, preview_entries=entries, written_count=3,
                          coverage={"complete": True, "covered_pages": 10, "section_count": 3},
                          outputs=[{"filename": filename, "range": [start, end], "page_ids": list(range(start + 1, end + 1))}
                                   for (start, end), filename in zip(unicode_ranges, names)])
    unicode_parity["writer_result"].update(total_pages=10, written_count=3, coverage=deepcopy(unicode_parity["coverage"]),
        outputs=[{**entry, "page_count": entry["end"] - entry["start"]} for entry in entries])
    cases = []
    for identifier, (mode, status, code, exit_code, fallback) in runner.DIAGNOSTIC_EXPECTATIONS.items():
        result = {"protocol": "winbooksplit.result", "version": 1, "mode": mode, "status": status, "code": code,
                  "message": "Synthetic diagnostic message", "exit_code": exit_code, "warnings": [],
                  "fallback_modes": fallback, "written_count": 0, "execution": None}
        case = {"id": identifier, "passed": True, "exit_code": exit_code, "outputs": [], "input_unchanged": True,
                "neighbor_unchanged": True, "no_automatic_fallback": True, "stderr": ""}
        if status == "success":
            parity = deepcopy(parities[mode])
            result.update(written_count=parity["written_count"], execution=parity["writer_result"])
            case.update(total_pages=parity["total_pages"], preview_entries=parity["preview_entries"], outputs=parity["outputs"])
        if identifier == "existing-output":
            case["existing_outputs_preserved"] = True
        if identifier == "success-level1":
            case["unicode_preserved"] = True
        case.update(diagnostic=result, stdout=json.dumps(result) + "\n")
        cases.append(case)
    shells = []
    for index, host in enumerate(hosts):
        categories = []
        for case in cases:
            if case["id"] in {"invalid-mode", "missing-arguments"}:
                continue
            decisions = []
            for choice in runner.DIAGNOSTIC_CHOICES:
                decision, retry = runner.diagnostic_decision(case["diagnostic"]["fallback_modes"], choice)
                decisions.append({"choice": choice, "decision": decision, "retry_mode": retry,
                                  "message": "No usable Level 2 bookmarks" if case["id"] == "parents-no-level2" else "Synthetic handler message",
                                  "fallback_modes": case["diagnostic"]["fallback_modes"]})
            categories.append({"id": case["id"], "passed": True, "parsed_result": deepcopy(case["diagnostic"]), "decisions": decisions})
        shells.append({"shell_executable": host, "passed": True, "exit_code": 0, "host_major": 5 if index == 0 else 7,
                       "host_version": "5.1.1234.1" if index == 0 else "7.6.5", "category_cases": categories,
                       "protocol_cases": [{"id": name, "passed": True, "rejected": True} for name in sorted(runner.DIAGNOSTIC_PROTOCOL_IDS)],
                       "native_argument_probe": {"passed": True, "exit_code": 0, "actual_arguments": runner.DIAGNOSTIC_NATIVE_ARGUMENTS},
                       "stream_probe": {"passed": True, "exit_code": 1, "stdout_length": 200000, "stderr_length": 200000,
                                        "actual_run_function": True, "arguments_preserved": True},
                       "no_automatic_execution": True, "owned_neighbor_unchanged": True})
    return {"schema_version": 1, "task_id": "M1-T06", "result": "DIAGNOSTIC_REGRESSION_PASSED", "success": True, "exit_code": 0,
            "acceptance_ids": ["AC-030", "AC-031"], "engine_case_count": 23, "engine_cases": cases, "shell_cases": shells,
            "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f", "import_observation": {"import_safe": True}}


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
            "historical_bookmark_cases": [], "level2_launcher_reference": corrected_level2_reference(),
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
                "missing current historical declaration": {"historical_bookmark_cases": None},
                "obsolete Level 1 known-bad comparisons": {"historical_bookmark_cases": [
                    {"oracle_id": f"BM-{number:02}", "equivalent": True} for number in range(1, 4)]},
                "obsolete Level 2 known-bad comparison": {"historical_bookmark_cases": [{"oracle_id": "BM-03", "equivalent": True}]},
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
                "missing corrected launcher reference": {"level2_launcher_reference": None},
                "wrong corrected title": {"level2_launcher_reference": {**corrected_level2_reference(), "titles": ["A1", "A2", "B1"]}},
                "old crossing reference ranges": {"level2_launcher_reference": {**corrected_level2_reference(),
                                                                               "expected_ranges": [[3, 6], [6, 10], [10, 12]]}},
                "incomplete corrected outputs": {"level2_launcher_reference": {**corrected_level2_reference(),
                                                                                "outputs": corrected_level2_reference()["outputs"][:-1]}},
                "wrong corrected page identity": {"level2_launcher_reference": {**corrected_level2_reference(),
                    "outputs": [{**corrected_level2_reference()["outputs"][0], "page_ids": [2, 1]},
                                *corrected_level2_reference()["outputs"][1:]]}},
                "wrong corrected output title": {"level2_launcher_reference": {**corrected_level2_reference(),
                    "outputs": [{**corrected_level2_reference()["outputs"][0], "filename": "01 - A1.pdf"},
                                *corrected_level2_reference()["outputs"][1:]]}},
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

    def test_full_selects_all_five_regressions_and_forwards_hosts_to_manual_and_diagnostics(self):
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
        level2 = [command for command in commands if str(ROOT / "tests/bookmarks/characterize_level2.py") in command]
        self.assertEqual(len(level2), 1)
        self.assertNotIn("--shell-path", level2[0])
        plans = [command for command in commands if str(ROOT / "tests/plans/characterize_plan.py") in command]
        self.assertEqual(len(plans), 1)
        self.assertNotIn("--shell-path", plans[0])
        diagnostics = [command for command in commands if str(ROOT / "tests/diagnostics/characterize_diagnostics.py") in command]
        self.assertEqual(len(diagnostics), 1)
        self.assertEqual(diagnostics[0][-4:], ["--shell-path", str(hosts[0]), "--shell-path", str(hosts[1])])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["shell", "shell", "manual", "bookmarks", "level2", "plan", "diagnostics"])
        self.assertEqual([step["name"] for step in report["steps"][-5:]],
                         ["manual-regression", "level1-regression", "level2-regression", "shared-plan-regression", "diagnostic-regression"])
        self.assertEqual(len(report["steps"]), 8)
        self.assertTrue(report["success"])

    def test_diagnostic_evidence_requires_categories_zero_outputs_both_hosts_and_all_decisions(self):
        hosts = ["C:/trusted/ps51.exe", "C:/trusted/pwsh.exe"]
        complete = complete_diagnostic_report(hosts)
        with tempfile.TemporaryDirectory(prefix="wbs-diagnostic-evidence-") as directory:
            path = Path(directory) / "diagnostics.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            valid = {"exit_code": 0, "requested_shell_paths": hosts}
            runner.attach_child_report(valid, path, "diagnostics")
            self.assertEqual(valid["exit_code"], 0)
            self.assertEqual(valid["evidence"], complete)
            self.assertIn("evidence_sha256", valid)
            defects = []

            def changed(description, mutate):
                report = deepcopy(complete)
                mutate(report)
                defects.append((description, report))

            for field, value in (("schema_version", True), ("task_id", "M1-T05"), ("result", "SHARED_PLAN_REGRESSION_PASSED"),
                                 ("success", False), ("exit_code", False), ("acceptance_ids", ["AC-030"]),
                                 ("engine_case_count", 0), ("source_unchanged", False), ("baseline_guards_preserved", False),
                                 ("input_and_neighbor_unchanged", False), ("owned_temp_removed", False),
                                 ("immutable_original_commit", "unknown"), ("import_observation", []), ("engine_cases", []), ("shell_cases", [])):
                defects.append((field, {**complete, field: value}))
            for field, value in (("id", []), ("id", "unknown"), ("passed", False), ("exit_code", 0),
                                 ("outputs", [{"filename": "unwanted.pdf"}]), ("stdout", "no result"), ("input_unchanged", False),
                                 ("neighbor_unchanged", False), ("no_automatic_fallback", False)):
                changed("engine " + field, lambda report, field=field, value=value: report["engine_cases"][0].update({field: value}))
            for field, value in (("protocol", "wrong"), ("version", True), ("status", "success"), ("code", "no_bookmarks_at_level"),
                                 ("written_count", 1), ("execution", {}), ("fallback_modes", ["1", "manual"]), ("warnings", {})):
                changed("category " + field, lambda report, field=field, value=value: report["engine_cases"][0]["diagnostic"].update({field: value}))
            changed("duplicate engine", lambda report: report["engine_cases"].__setitem__(-1, deepcopy(report["engine_cases"][0])))
            changed("existing output changed", lambda report: next(case for case in report["engine_cases"] if case["id"] == "existing-output").update(existing_outputs_preserved=False))
            changed("unicode not verified", lambda report: next(case for case in report["engine_cases"] if case["id"] == "success-level1").update(unicode_preserved=False))
            for field, value in (("shell_executable", "C:/other/pwsh.exe"), ("host_major", []), ("host_version", "6.0"),
                                 ("passed", False), ("exit_code", 19), ("category_cases", []), ("protocol_cases", []),
                                 ("no_automatic_execution", False), ("owned_neighbor_unchanged", False),
                                 ("native_argument_probe", {}), ("stream_probe", {})):
                changed("host " + field, lambda report, field=field, value=value: report["shell_cases"][0].update({field: value}))
            changed("duplicate host", lambda report: report["shell_cases"].__setitem__(1, deepcopy(report["shell_cases"][0])))
            changed("duplicate category", lambda report: report["shell_cases"][0]["category_cases"].__setitem__(-1, deepcopy(report["shell_cases"][0]["category_cases"][0])))
            changed("unparsed category", lambda report: report["shell_cases"][0]["category_cases"][0].update(parsed_result={}))
            changed("missing decisions", lambda report: report["shell_cases"][0]["category_cases"][0].update(decisions=[]))
            changed("automatic retry", lambda report: report["shell_cases"][0]["category_cases"][0]["decisions"][0].update(decision="retry", retry_mode="manual"))
            changed("wrong level choice", lambda report: report["shell_cases"][0]["category_cases"][0]["decisions"][3].update(decision="retry", retry_mode="1"))
            changed("substring choice accepted", lambda report: report["shell_cases"][0]["category_cases"][0]["decisions"][6].update(decision="retry", retry_mode="manual"))
            changed("misleading Level2 message", lambda report: next(case for case in report["shell_cases"][0]["category_cases"] if case["id"] == "parents-no-level2")["decisions"][0].update(message="No outline bookmarks in PDF"))
            changed("accepted invalid protocol", lambda report: report["shell_cases"][0]["protocol_cases"][0].update(rejected=False))
            changed("changed literal argv", lambda report: report["shell_cases"][0]["native_argument_probe"].update(actual_arguments=["wrong"]))
            changed("lost stderr", lambda report: report["shell_cases"][0]["stream_probe"].update(stderr_length=0))
            changed("synthetic stream replacement", lambda report: report["shell_cases"][0]["stream_probe"].update(actual_run_function=False))
            for description, payload in defects:
                with self.subTest(defect=description):
                    path.write_text(json.dumps(payload), encoding="utf-8")
                    step = {"exit_code": 0, "requested_shell_paths": hosts}
                    runner.attach_child_report(step, path, "diagnostics")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    failed = {"exit_code": 19, "requested_shell_paths": hosts}
                    runner.attach_child_report(failed, path, "diagnostics")
                    self.assertEqual(failed["exit_code"], 19)
            path.write_text(json.dumps(complete), encoding="utf-8")
            for requested in ([], [hosts[0]], [hosts[0], hosts[0]]):
                with self.subTest(requested=requested):
                    step = {"exit_code": 0, "requested_shell_paths": requested}
                    runner.attach_child_report(step, path, "diagnostics")
                    self.assertEqual(step["exit_code"], 126)

    def test_plan_targeted_route_runs_only_plan_without_shell_arguments(self):
        args = SimpleNamespace(layer="plan", failure_probe=None, shell_path=[Path("C:/unused/pwsh.exe")], tool_root=None)
        commands = []

        def successful_command(argv, cwd, **kwargs):
            commands.append(argv)
            return {"exit_code": 0, "stdout": "synthetic-git-identity\n", "stderr": ""}

        with patch.object(runner, "source_manifest", return_value={"synthetic-source": "hash"}), \
                patch.object(runner, "run_command", side_effect=successful_command), \
                patch.object(runner, "attach_child_report") as attach:
            report = runner.execute(args)
        self.assertEqual([step["name"] for step in report["steps"]], ["shared-plan-regression"])
        self.assertEqual([call.args[2] for call in attach.call_args_list], ["plan"])
        self.assertEqual(len(commands), 3)  # Git commit, tree, then the single targeted child.
        self.assertEqual(commands[-1][3], str(ROOT / "tests/plans/characterize_plan.py"))
        self.assertNotIn("--shell-path", commands[-1])
        self.assertTrue(report["success"])

    def test_plan_evidence_requires_all_promises_and_exact_preview_writer_identities(self):
        complete = complete_plan_report()
        with tempfile.TemporaryDirectory(prefix="wbs-plan-evidence-") as directory:
            path = Path(directory) / "plan.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "plan")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            defects = []

            def defect(description, group, index, changes):
                report = deepcopy(complete)
                report[group][index].update(changes)
                defects.append((description, report))

            for field, value in (("schema_version", True), ("task_id", "M1-T04"), ("result", "LEVEL2_REGRESSION_PASSED"),
                                 ("success", False), ("exit_code", False), ("acceptance_ids", ["AC-027"]),
                                 ("source_unchanged", False), ("baseline_guards_preserved", False),
                                 ("input_and_neighbor_unchanged", False), ("owned_temp_removed", False),
                                 ("immutable_original_commit", "unknown"), ("import_observation", [])):
                defects.append((field, {**complete, field: value}))
            for group in ("structural_cases", "preview_cases", "parity_cases", "source_cases"):
                defects.append(("missing " + group, {**complete, group: []}))
                defects.append(("duplicate " + group, {**complete, group: complete[group][:-1] + [complete[group][0]]}))
                key = "mode" if group == "parity_cases" else "id"
                defect("wrong " + group + " ID", group, 0, {key: "unexpected"})
                defect("unhashable " + group + " ID", group, 0, {key: []})
                defect("failed " + group, group, 0, {"passed": False})
            defect("no actual structural rejection", "structural_cases", 0, {"rejected": False})
            defect("missing rejection diagnostic", "structural_cases", 0, {"error_code": ""})
            defect("whole document rejected", "structural_cases", 8, {"accepted": False})
            defect("mutable preview", "preview_cases", 0, {"mutation_rejected": False})
            defect("preview wrote files", "preview_cases", 0, {"chapter_files_written": 1})
            defect("boolean preview count", "preview_cases", 0, {"chapter_files_written": False})
            defect("aliased preview metadata", "preview_cases", 1, {"detached_from_input": False})
            defect("unbound mapping executed", "preview_cases", 1, {"unbound_execution_rejected": False})
            for group in ("parity_cases", "source_cases"):
                missing_result = deepcopy(complete)
                del missing_result[group][0]["writer_result"]
                defects.append(("missing writer result " + group, missing_result))
                defect("zero files " + group, group, 0, {"outputs": [], "written_count": 0})
                defect("missing preview " + group, group, 0, {"preview_entries": []})
                defect("false output count " + group, group, 0, {"written_count": True})
                defect("wrong coverage " + group, group, 0, {"coverage": {"complete": True, "covered_pages": 2, "section_count": 2}})
                defect("integer complete flag " + group, group, 0, {"coverage": {"complete": 1, "covered_pages": 3, "section_count": 2}})
                defect("changed neighbor " + group, group, 0, {"neighbor_unchanged": False})
                defect("replanned " + group, group, 0, {"no_replanning_or_reopening": False})
                output_changes = [{"page_ids": [3, 2]}, {"page_ids": [True]}, {"filename": "different.pdf"}, {"range": [0, 3]}]
                for index, changes in enumerate(output_changes):
                    report = deepcopy(complete)
                    report[group][0]["outputs"][0].update(changes)
                    defects.append((f"wrong physical output {group} {index}", report))
                report = deepcopy(complete)
                report[group][0]["preview_entries"][1]["start"] = 2
                defects.append(("preview gap " + group, report))
                report = deepcopy(complete)
                report[group][0]["preview_entries"][0]["sequence"] = True
                defects.append(("boolean sequence " + group, report))
            defect("executed replacement", "source_cases", 0, {"outcome": "changed_source_used"})
            defect("replacement changed", "source_cases", 0, {"replacement_unchanged": False})
            defect("wrong bound digest", "source_cases", 0, {"captured_sha256": "c" * 64})
            defect("unchanged replacement", "source_cases", 0, {"replacement_sha256": "a" * 64})
            defect("invalid replacement digest", "source_cases", 0, {"replacement_sha256": "not-a-digest"})
            defect("nonhex replacement digest", "source_cases", 0, {"replacement_sha256": "g" * 64})
            defect("replacement page content used", "source_cases", 0, {"page_content_sha256": ["f" * 64] * 3})
            defect("missing original page content", "source_cases", 0, {"original_page_content_sha256": []})
            defect("ordinary writer page content mismatch", "parity_cases", 0, {"page_content_sha256": ["f" * 64] * 3})
            defect("ordinary original content count mismatch", "parity_cases", 0, {"original_page_content_sha256": []})
            defect("ordinary original content invalid hash", "parity_cases", 0, {"original_page_content_sha256": ["not-a-hash"] * 3})
            for group in ("parity_cases", "source_cases"):
                for field in ("original_page_content_sha256", "page_content_sha256"):
                    missing_content = deepcopy(complete)
                    del missing_content[group][0][field]
                    defects.append(("omitted " + group + " " + field, missing_content))
            defect("uncaptured source", "source_cases", 0, {"source_identity": {"binding": "path"}})
            for group in ("parity_cases", "source_cases"):
                report = deepcopy(complete)
                case = report[group][0]
                case["writer_result"] = {"mode": case.get("mode", "manual"), "total_pages": 3,
                                         "written_count": 2, "coverage": deepcopy(case["coverage"]),
                                         "source_identity": deepcopy(case["source_identity"]),
                                         "outputs": [{**entry, "page_count": entry["end"] - entry["start"]}
                                                     for entry in case["preview_entries"]]}
                path.write_text(json.dumps(report), encoding="utf-8")
                valid_result_step = {"exit_code": 0}
                runner.attach_child_report(valid_result_step, path, "plan")
                self.assertEqual(valid_result_step["exit_code"], 0)
                for field in ("mode", "total_pages", "written_count", "coverage", "source_identity", "outputs"):
                    changed = deepcopy(report)
                    del changed[group][0]["writer_result"][field]
                    defects.append(("omitted writer result " + group + " " + field, changed))
                for field, value in (("written_count", 1), ("total_pages", True), ("outputs", []),
                                     ("coverage", {}), ("source_identity", {}), ("mode", "invalid"), ("mode", [])):
                    changed = deepcopy(report)
                    changed[group][0]["writer_result"][field] = value
                    defects.append(("writer result " + group + " " + field, changed))
                changed = deepcopy(report)
                changed[group][0]["writer_result"]["outputs"][0]["page_count"] = True
                defects.append(("boolean result page count " + group, changed))
            missing_deleted = deepcopy(complete)
            del missing_deleted["source_cases"][1]["deleted_path_snapshot"]
            defects.append(("omitted deleted-path execution", missing_deleted))
            defect("invalid deleted-path object", "source_cases", 1, {"deleted_path_snapshot": []})
            for field in ("writer_result", "original_page_content_sha256", "page_content_sha256", "source_identity"):
                missing_nested = deepcopy(complete)
                del missing_nested["source_cases"][1]["deleted_path_snapshot"][field]
                defects.append(("omitted deleted-path " + field, missing_nested))
            for field, value in (("written_count", 0), ("coverage", {}), ("outputs", []),
                                 ("page_content_sha256", ["f" * 64] * 3)):
                changed = deepcopy(complete)
                changed["source_cases"][1]["deleted_path_snapshot"][field] = value
                defects.append(("deleted-path mismatch " + field, changed))
            changed = deepcopy(complete)
            nested = changed["source_cases"][1]["deleted_path_snapshot"]
            nested["source_identity"]["sha256"] = "f" * 64
            nested["writer_result"]["source_identity"]["sha256"] = "f" * 64
            defects.append(("deleted-path switched bound source", changed))
            changed = deepcopy(complete)
            nested = changed["source_cases"][1]["deleted_path_snapshot"]
            nested["original_page_content_sha256"] = nested["page_content_sha256"] = ["f" * 64] * 3
            defects.append(("deleted-path switched original content", changed))
            changed = deepcopy(complete)
            changed["source_cases"][1]["deleted_path_snapshot"]["writer_result"]["outputs"] = []
            defects.append(("deleted-path writer result mismatch", changed))
            for changes in ({"seed": 0}, {"count": 299}, {"count": True}, {"passed": False},
                            {"per_mode": {"manual": 100, "1": 100}}, {"per_mode": {"manual": 99, "1": 101, "2": 100}}):
                defects.append(("seeded modes " + repr(changes), {**complete, "seeded_plan_cases": {**complete["seeded_plan_cases"], **changes}}))
            for description, payload in defects:
                with self.subTest(defect=description):
                    path.write_text(json.dumps(payload), encoding="utf-8")
                    step = {"exit_code": 0}
                    runner.attach_child_report(step, path, "plan")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    failed_child = {"exit_code": 19}
                    runner.attach_child_report(failed_child, path, "plan")
                    self.assertEqual(failed_child["exit_code"], 19)

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

    def test_level2_evidence_requires_complete_targets_hierarchy_and_preservation(self):
        complete = {
            "schema_version": 1, "task_id": "M1-T04", "result": "LEVEL2_REGRESSION_PASSED",
            "success": True, "exit_code": 0, "acceptance_ids": ["AC-023", "AC-024", "AC-025", "AC-026"],
            "engine_case_count": 4,
            "engine_cases": [{"oracle_id": name, "passed": True} for name in ("BM-03", "BM-04", "BM-06", "BM-07")],
            "hierarchy_cases": [{"id": name, "passed": True} for name in (
                "duplicate-parent-subtrees", "invalid-parent-subtrees", "child-order-and-aliases",
                "invalid-child-destinations", "deep-and-malformed-outlines", "no-usable-level2")],
            "seeded_level2_cases": {"seed": 20261009, "count": 150, "passed": True},
            "sampled_writer_cases": [{"index": number, "outputs": [{"filename": "synthetic.pdf"}]}
                                     for number in range(10)],
            "source_unchanged": True, "baseline_guards_preserved": True,
            "input_and_neighbor_unchanged": True, "owned_temp_removed": True,
            "immutable_original_commit": "0de84f367f9bd5ddfa3f408a9c29505d7a39633f",
            "import_observation": {"import_safe": True},
        }
        with tempfile.TemporaryDirectory(prefix="wbs-level2-evidence-") as directory:
            path = Path(directory) / "level2.json"
            path.write_text(json.dumps(complete), encoding="utf-8")
            step = {"exit_code": 0}
            runner.attach_child_report(step, path, "level2")
            self.assertEqual(step["exit_code"], 0)
            self.assertEqual(step["evidence"], complete)
            self.assertIn("evidence_sha256", step)
            defects = {
                "boolean schema": {"schema_version": True},
                "wrong task": {"task_id": "M1-T03"},
                "wrong result": {"result": "LEVEL1_REGRESSION_PASSED"},
                "claimed failure": {"success": False},
                "boolean exit": {"exit_code": False},
                "incomplete acceptance": {"acceptance_ids": ["AC-023"]},
                "wrong case count": {"engine_case_count": 3},
                "count without actual cases": {"engine_cases": []},
                "duplicate target ID": {"engine_cases": complete["engine_cases"][:-1] + [complete["engine_cases"][0]]},
                "wrong target ID": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-08", "passed": True}]},
                "invalid target ID type": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": [], "passed": True}]},
                "failed target": {"engine_cases": complete["engine_cases"][:-1] + [{"oracle_id": "BM-07", "passed": False}]},
                "missing hierarchy cases": {"hierarchy_cases": []},
                "duplicate hierarchy ID": {"hierarchy_cases": complete["hierarchy_cases"][:-1] + [complete["hierarchy_cases"][0]]},
                "wrong hierarchy ID": {"hierarchy_cases": complete["hierarchy_cases"][:-1] + [{"id": "unexpected", "passed": True}]},
                "invalid hierarchy ID type": {"hierarchy_cases": complete["hierarchy_cases"][:-1] + [{"id": [], "passed": True}]},
                "failed hierarchy case": {"hierarchy_cases": complete["hierarchy_cases"][:-1]
                                          + [{**complete["hierarchy_cases"][-1], "passed": False}]},
                "wrong seed": {"seeded_level2_cases": {"seed": 0, "count": 150, "passed": True}},
                "wrong property count": {"seeded_level2_cases": {"seed": 20261009, "count": 0, "passed": True}},
                "failed properties": {"seeded_level2_cases": {"seed": 20261009, "count": 150, "passed": False}},
                "missing writer samples": {"sampled_writer_cases": []},
                "duplicate writer indexes": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                             + [complete["sampled_writer_cases"][0]]},
                "boolean writer index": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                         + [{"index": True, "outputs": [{"filename": "synthetic.pdf"}]}]},
                "zero writer files": {"sampled_writer_cases": complete["sampled_writer_cases"][:-1]
                                      + [{"index": 9, "outputs": []}]},
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
                    runner.attach_child_report(step, path, "level2")
                    self.assertEqual(step["exit_code"], 126)
                    self.assertIn("evidence_error", step)
                    self.assertNotIn("evidence", step)
                    step = {"exit_code": 19}
                    runner.attach_child_report(step, path, "level2")
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
