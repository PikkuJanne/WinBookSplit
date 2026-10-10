"""Synthetic structural units; no actual Windows acceptance is claimed here."""

from copy import deepcopy
import importlib.util
import json
import io
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
from unittest.mock import patch
from contextlib import redirect_stdout, redirect_stderr
import unittest

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_process_receipt_validator", ROOT / "tests/process/validate_process_report.py")
validator = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(validator)


def complete_fast_case():
    stdout = "\n\nstdout blank\n\nno-newline-output"
    stderr = "\n\nFINAL_STDERR_NO_NEWLINE"
    return {"id": "PS51-fast-tail", "kind": "fast-tail", "passed": True,
        "actual_process": True, "unrelated_process_alive": True,
        "process": {"Pid": 1234, "ExitCode": 23, "TimedOut": False, "Cancelled": False,
            "ParentStopped": True, "DescendantsStopped": True, "StreamsComplete": True, "JobAssigned": True,
            "StartError": None, "StreamError": None, "StopError": None, "ElapsedSeconds": .1,
            "Stdout": stdout, "Stderr": stderr, "StdoutTotalBytes": len(stdout.encode()),
            "StderrTotalBytes": len(stderr.encode()), "StdoutTruncated": False, "StderrTruncated": False,
            "ResultRecords": [], "ResultError": None}}


def complete_application_case():
    engine = validator.expected_frame()
    outcome = {"protocol": "winbooksplit.outcome", "version": 1, "status": "invalid_input", "code": "invalid_start_pages",
        "exit_code": 2, "mode": "manual", "written_count": 0, "final_directory": None, "engine_result": deepcopy(engine)}
    return {"kind": "fast-tail", "exit_code": 2, "engine_record": engine,
        "provisional_outcome": deepcopy(outcome), "final_outcome": deepcopy(outcome),
        "stdout": "[OUTCOME] " + json.dumps(outcome) + "\n"}


class ProcessReceiptTests(unittest.TestCase):
    def test_failed_readiness_preserves_raw_native_case_and_missing_pid_observation(self):
        if os.name != "nt":
            self.skipTest("The existing runner validates actual Windows executable paths")
        spec = importlib.util.spec_from_file_location("wbs_process_failure_unit", ROOT / "tests/process/characterize_process.py")
        harness = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = harness
        spec.loader.exec_module(harness)
        shell51 = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
        with tempfile.TemporaryDirectory(prefix="wbs-pure-failure-") as outside:
            root = Path(outside)
            second = root / "synthetic-host.exe"
            second.write_bytes(b"No child is executed by this pure unit\n")
            report_path = root / "failure.json"
            retained = []
            host = {"id": "PS51", "shell_executable": str(shell51), "host_version": "5.1.26100.9444",
                "stored_policies": [{"scope": "CurrentUser", "policy": "Undefined"}]}
            raw_case = complete_fast_case()
            raw_case.update(id="PS51-detached-pipe", kind="detached-pipe", arguments=["authored"])
            raw_case.pop("passed")
            raw_case.pop("unrelated_process_alive")
            raw_case["process"].update(ExitCode=0, TimedOut=True, Stdout="", Stderr="", StdoutTotalBytes=0, StderrTotalBytes=0)

            def fake_host(command, cwd, **options):
                (cwd / "native-report.json").write_text(json.dumps({"host_version": host["host_version"],
                    "stored_policies": host["stored_policies"], "cases": [raw_case]}), encoding="utf-8")
                (cwd / "detached-pipe-parent.json").write_text(json.dumps({"pid": 1234}), encoding="utf-8")
                return {"exit_code": 0, "stdout": "authored host tail", "stderr": "authored host stderr"}

            def fail(directory, shells):
                retained.append(directory)
                harness.OBSERVABILITY.update(tested_path_sha256={"authored.py": "a" * 64}, host_cases=[host])
                harness.native_cases(directory, [host], SimpleNamespace(pid=99, poll=lambda: None))

            class Capture(io.StringIO):
                def reconfigure(self, **options):
                    pass

            argv = ["characterize_process.py", "--report", str(report_path), "--shell-path", str(shell51),
                "--shell-path", str(second)]
            with patch.object(sys, "argv", argv), patch.object(harness, "characterize", fail), \
                    patch.object(harness.history, "run_entrypoint", fake_host), \
                    patch.object(harness.runner, "source_manifest", return_value={"authored.py": "a" * 64}), \
                    patch.object(harness, "running", side_effect=lambda pid: pid == 99), \
                    redirect_stdout(Capture()), redirect_stderr(Capture()):
                self.assertEqual(harness.main(), 1)
            report = json.loads(report_path.read_text(encoding="utf-8"))
            try:
                self.assertTrue(retained[0].is_dir())
                self.assertTrue((retained[0] / "PS51-native/native-report.json").is_file())
                self.assertEqual(report["last_case"]["process"], raw_case["process"])
                self.assertEqual(report["last_case"]["command"], report["native_host_evidence"][0]["command"])
                self.assertFalse(report["last_case"]["pid_observations"][-1]["receipt"]["available"])
                validator.validate_failed_process_report(report)
                for mutate in (lambda r: r.update(success=True), lambda r: r.update(owned_temp_removed=True),
                        lambda r: r.update(native_host_evidence=[]), lambda r: r["last_case"]["process"].update(ExitCode=23),
                        lambda r: r["last_case"]["pid_observations"][-1]["receipt"].update(available=True),
                        lambda r: r["native_host_evidence"][0]["native_report"].update(sha256="b" * 64)):
                    changed = deepcopy(report)
                    mutate(changed)
                    with self.subTest(mutate=mutate), self.assertRaises(ValueError):
                        validator.validate_failed_process_report(changed)
            finally:
                # Only exact authored ordinary JSON members exist; no process was started.
                for member in (retained[0] / "PS51-native").iterdir():
                    self.assertTrue(member.is_file())
                    member.unlink()
                (retained[0] / "PS51-native").rmdir()
                retained[0].rmdir()

    def test_raw_artifact_rejects_resealed_length_and_absent_file_claims(self):
        raw = '{"pid":1234}'
        record = {"path": "C:/authored/child.json", "available": True, "size_bytes": len(raw.encode()),
            "created_ns": 1, "modified_ns": 1, "raw_text": raw, "sha256": validator.sha256(raw.encode()).hexdigest()}
        self.assertEqual(validator.artifact(record), raw)
        for mutate in (lambda r: r.update(size_bytes=True), lambda r: r.update(size_bytes=1),
                lambda r: r.update(raw_text=raw + " "), lambda r: r.update(sha256="c" * 64)):
            changed = deepcopy(record)
            mutate(changed)
            with self.subTest(mutate=mutate), self.assertRaises(ValueError):
                validator.artifact(changed)
        with self.assertRaises(ValueError):
            validator.artifact({"path": record["path"], "available": False, "raw_text": raw}, available=False)

    def test_failed_main_control_exit_codes_match_written_receipt(self):
        spec = importlib.util.spec_from_file_location("wbs_process_control_exit_unit", ROOT / "tests/process/characterize_process.py")
        harness = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = harness
        spec.loader.exec_module(harness)

        class Capture(io.StringIO):
            def reconfigure(self, **options):
                pass

        with tempfile.TemporaryDirectory(prefix="wbs-pure-control-") as outside:
            root = Path(outside)
            shells = [root / "first-host.exe", root / "second-host.exe"]
            for shell in shells:
                shell.write_bytes(b"Authored parser precondition, never executed\n")
            for number, (error, code) in enumerate(((KeyboardInterrupt(), 130), (SystemExit(19), 19), (TypeError("authored failure"), 1))):
                target = root / (str(number) + ".json")
                retained = []

                def fail(directory, actual_shells):
                    retained.append(directory)
                    harness.OBSERVABILITY["tested_path_sha256"] = {"authored.py": "a" * 64}
                    raise error

                argv = ["characterize_process.py", "--report", str(target), "--shell-path", str(shells[0]), "--shell-path", str(shells[1])]
                with patch.object(sys, "argv", argv), patch.object(harness, "characterize", fail), \
                        patch.object(harness.runner, "source_manifest", return_value={"authored.py": "a" * 64}), \
                        redirect_stdout(Capture()), redirect_stderr(Capture()):
                    self.assertEqual(harness.main(), code)
                record = json.loads(target.read_text(encoding="utf-8"))
                self.assertEqual(record["exit_code"], code)
                self.assertEqual(record["error_type"], type(error).__name__)
                validator.validate_failed_process_report(record)
                self.assertTrue(retained[0].is_dir())
                retained[0].rmdir()  # Exact empty authored directory; no child was started.

    def test_finalizer_footer_is_separate_from_exact_unterminated_stderr(self):
        case = complete_application_case()
        expected_stderr = b"E" * (2 * 1024 * 1024) + b"\nFINAL_STDERR_" + validator.UNICODE.encode("utf-8")
        retained_stderr = validator.tail(expected_stderr)
        logged_stdout = "\n\nstdout blank\n\nno-newline-output"
        log = "[STDOUT]\r\n" + logged_stdout + "\r\n[STDERR]\r\n" + retained_stderr \
            + "\r\n[OPERATION-OUTCOME] " + json.dumps(case["provisional_outcome"]) + "\r\n"
        actual_stdout, actual_stderr, provisional, final = validator.application_log_streams(log, case["stdout"])
        self.assertEqual(actual_stdout, logged_stdout)
        self.assertEqual(actual_stderr, retained_stderr)
        self.assertEqual(provisional, final)
        self.assertEqual(len(actual_stderr.encode("utf-8")), 65536)
        validator.stream({"Stderr": actual_stderr, "StderrTotalBytes": len(expected_stderr), "StderrTruncated": True},
            "Stderr", expected_stderr)
        validator.application_outcome(case)

    def test_finalizer_footer_missing_duplicate_malformed_or_nonterminal_is_rejected(self):
        case = complete_application_case()
        footer = "\r\n[OPERATION-OUTCOME] " + json.dumps(case["provisional_outcome"]) + "\r\n"
        transport = "[STDOUT]\r\n\n\nstdout\r\n[STDERR]\r\n\n\nstderr no newline"
        for log, stdout in ((transport, case["stdout"]), (transport + footer + footer, case["stdout"]),
                (transport + "\r\n[OPERATION-OUTCOME] invalid\r\n", case["stdout"]),
                (transport + footer + "trailing data", case["stdout"]), (transport + footer + "\r\n", case["stdout"]),
                (transport + footer, "missing final outcome"), (transport + footer, case["stdout"] * 2),
                (transport + footer, "[OUTCOME] {}\n")):
            with self.subTest(log=log[-30:], stdout=stdout[-30:]), self.assertRaises(ValueError):
                validator.application_log_streams(log, stdout)

    def test_application_finalizer_outcome_native_and_engine_binding_is_required(self):
        valid = complete_application_case()
        validator.application_outcome(valid)
        for field, replacement in (("protocol", "wrong"), ("version", True), ("status", "success"),
                ("code", "invalid_mode"), ("exit_code", 0), ("mode", "1"), ("written_count", 3),
                ("final_directory", "C:/invented"), ("engine_result", {})):
            case = deepcopy(valid)
            case["provisional_outcome"][field] = deepcopy(replacement)
            case["final_outcome"][field] = deepcopy(replacement)
            case["stdout"] = "[OUTCOME] " + json.dumps(case["final_outcome"]) + "\n"
            with self.subTest(field=field), self.assertRaises(ValueError):
                validator.application_outcome(case)
        for change in (lambda case: case.update(exit_code=0), lambda case: case.update(provisional_outcome={}),
                lambda case: case.update(final_outcome={}), lambda case: case.update(stdout="[OUTCOME] {}\n"),
                lambda case: case.update(stdout=None),
                lambda case: case["engine_record"].update(code="invalid_mode")):
            case = deepcopy(valid)
            change(case)
            with self.subTest(change=change), self.assertRaises(ValueError):
                validator.application_outcome(case)
    def test_complete_fast_exit_blank_and_tail_evidence_is_accepted(self):
        validator.native(complete_fast_case(), "PS51")

    def test_native_receipt_rejects_missing_and_contradictory_proofs(self):
        case = complete_fast_case()
        mutations = [
            ("case", "actual_process", False), ("case", "unrelated_process_alive", False),
            ("case", "id", "PS7-fast-tail"), ("case", "kind", "flood"),
            ("process", "Pid", 0), ("process", "Pid", True), ("process", "ExitCode", 0),
            ("process", "TimedOut", True), ("process", "Cancelled", True),
            ("process", "ParentStopped", False), ("process", "DescendantsStopped", None),
            ("process", "StreamsComplete", False), ("process", "JobAssigned", False),
            ("process", "StartError", "authored start error"), ("process", "StreamError", "decode error"),
            ("process", "StopError", "unproved stop"), ("process", "ElapsedSeconds", float("inf")),
            ("process", "ElapsedSeconds", 20), ("process", "Stdout", "stdout blank\nno-newline-output"),
            ("process", "Stderr", "\n\nFINAL_STDERR_NO_NEWLIN"), ("process", "StdoutTotalBytes", 1),
            ("process", "StderrTotalBytes", True), ("process", "StdoutTruncated", True),
            ("process", "ResultRecords", "[]"), ("process", "Stderr", "\ufffd"),
        ]
        for scope, field, value in mutations:
            with self.subTest(scope=scope, field=field):
                changed = deepcopy(case)
                target = changed if scope == "case" else changed["process"]
                target[field] = value
                with self.assertRaises(ValueError):
                    validator.native(changed, "PS51")
        for field in ("Pid", "ExitCode", "TimedOut", "Cancelled", "ParentStopped", "DescendantsStopped", "StreamsComplete",
                      "JobAssigned", "StartError", "StreamError", "StopError", "StdoutTotalBytes", "StderrTotalBytes", "ResultRecords"):
            with self.subTest(missing=field):
                changed = deepcopy(case)
                del changed["process"][field]
                with self.assertRaises(ValueError):
                    validator.native(changed, "PS51")

    def test_valid_large_result_is_independent_of_bounded_human_tail(self):
        case = complete_fast_case()
        case.update(id="PS51-large-frame", kind="large-frame")
        message = "L" * 100000 + validator.UNICODE
        data = validator.frame_bytes(message)
        case["process"].update(ExitCode=2, Stdout=validator.tail(data), Stderr="", StdoutTotalBytes=len(data),
            StderrTotalBytes=0, StdoutTruncated=True, ResultRecords=[data.decode()], ResultError=None)
        case.update(parsed_result=validator.expected_frame(message), protocol_error=None)
        validator.native(case, "PS51")
        for field, value in (("ResultRecords", []), ("ResultError", "over budget")):
            changed = deepcopy(case)
            changed["process"][field] = value
            with self.assertRaises(ValueError):
                validator.native(changed, "PS51")

    def test_exact_cases_rejects_missing_duplicate_and_extra_rows(self):
        cases = [{"id": "PS51-fast-tail"}, {"id": "PS7-fast-tail"}]
        expected = [item["id"] for item in cases]
        self.assertEqual(set(validator.exact_cases(cases, expected)), set(expected))
        for changed in (cases[:1], [cases[0], cases[0]], [*cases, {"id": "invented"}], None):
            with self.subTest(changed=changed), self.assertRaises(ValueError):
                validator.exact_cases(changed, expected)

    def test_unicode_boundary_retains_valid_whole_codepoints(self):
        data = ("日" * (65536 // 3 + 1) + validator.UNICODE).encode()
        text = validator.tail(data)
        self.assertLessEqual(len(text.encode()), 65536)
        self.assertNotIn("\ufffd", text)
        self.assertTrue(text.endswith(validator.UNICODE))
        validator.stream({"Stdout": text, "StdoutTotalBytes": len(data), "StdoutTruncated": True}, "Stdout", data)

    def test_top_level_rejects_a_claim_without_promised_host_case_evidence(self):
        report = {"schema_version": 1, "task_id": "M2-T05", "result": "PROCESS_REGRESSION_PASSED", "exit_code": 0,
            "success": True, "source_unchanged": True, "baseline_guards_preserved": True, "input_and_neighbor_unchanged": True,
            "machine_settings_unchanged": True, "owned_temp_removed": True, "unrelated_process_removed": True,
            "acceptance_ids": validator.ACCEPTANCE, "tested_path_sha256": {"authored.py": "a" * 64}}
        with self.assertRaisesRegex(ValueError, "Promised supervisor/application/harness source hashes missing"):
            validator.validate_process_report(report, [r"C:\PS51\powershell.exe", r"C:\PS7\pwsh.exe"])


if __name__ == "__main__":
    unittest.main()
