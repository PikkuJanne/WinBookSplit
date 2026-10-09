"""Synthetic structural units; no actual Windows acceptance is claimed here."""

from copy import deepcopy
import importlib.util
import json
from pathlib import Path
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


class ProcessReceiptTests(unittest.TestCase):
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
        case["process"].update(ExitCode=1, Stdout=validator.tail(data), Stderr="", StdoutTotalBytes=len(data),
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
