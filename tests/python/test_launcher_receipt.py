"""Synthetic receipt mutations only; these are not Windows/Explorer passes."""

from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_launcher_unit_validator", ROOT / "tests/launcher/validate_launcher_report.py")
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)
spec = importlib.util.spec_from_file_location("wbs_launcher_unit_runner", ROOT / "tests/run_tests.py")
runner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(runner)


def seal(row):
    row["stdout"] = "Enter the literal path to one PDF, EPUB or AZW3 (blank or C cancels)\n" + (
        "[OUTCOME] " + json.dumps(row["outcome"]) + "\n" if row["outcome"] else "Please select one PDF, EPUB or AZW3 file at a time.\n")
    row["stdout_sha256"] = sha256(row["stdout"].encode()).hexdigest()
    return row


def cancelled_case(kind="source-c"):
    source = str(Path("C:/synthetic/b/Å [1] & %literal%!literal!.pdf"))
    identities = {name: {"sha256": "a" * 64, "size_bytes": 10, "device": 1, "inode": index, "attributes": 1}
                  for index, name in enumerate(("source", "neighbor", "prior"), 1)}
    answers, _ = validator.recipe(kind, source)
    params = ([] if kind.startswith("source-") else ["-InputFile", source]) + ["-NoPause", "-PythonPath", "C:/synthetic/python.exe", "-OutputDirectory", "C:/synthetic/o"]
    row = {"id": "PS51-" + kind, "kind": kind, "host_id": "PS51", "host_version": "5.1.26100.9444",
        "shell_executable": "C:/synthetic/powershell.exe", "exit_code": 130, "timed_out": False, "pid": 1234,
        "stdin": "PIPE_UTF8", "stdin_utf8": answers, "stdin_sha256": sha256(answers.encode()).hexdigest(),
        "elapsed_seconds": .1, "timeout_seconds": 90, "actual_process": True, "passed": True,
        "source_read_only_observed": True, "source_neighbor_prior_unchanged": True, "application_unchanged": True,
        "known_output_cleanup_complete": True, "stderr": "", "stderr_sha256": sha256(b"").hexdigest(),
        "source_observations_before": deepcopy(identities), "source_observations_after": deepcopy(identities),
        "source_observations_after_cleanup": deepcopy(identities), "additional_inputs_before": {"second": deepcopy(identities["source"]), "third": deepcopy(identities["neighbor"])},
        "output_members_before": ["prior-output.pdf"], "output_members_after": ["prior-output.pdf"], "output_members_after_cleanup": ["prior-output.pdf"],
        "application_sha256": {name: "b" * 64 for name in validator.APPLICATION}, "modifications": [],
        "command": ["C:/synthetic/powershell.exe", "-NoProfile", "-File", "C:/synthetic/a/WinBookSplit.ps1", *params],
        "parameters": params, "cwd": "C:/synthetic/c", "input_path": source, "output_base": "C:/synthetic/o",
        "outcome": {"protocol": "winbooksplit.outcome", "version": 1, "status": "cancelled", "code": "cancelled", "exit_code": 130,
                    "written_count": 0, "engine_result": None, "final_directory": None}, "engine_records": [], "process_summaries": [], "log_sha256": None,
        "final_publication": None, "console_operation_outcomes": []}
    row["additional_inputs_after"] = deepcopy(row["additional_inputs_before"])
    row["actual_application_sha256"] = deepcopy(row["application_sha256"])
    if kind in {"menu-cancel", "menu-eof"}:
        row["log_sha256"] = "c" * 64
        row["console_operation_outcomes"] = [deepcopy(row["outcome"])]
        row["console_log_scope"] = {"path": str(Path("C:/synthetic/o/.WinBookSplit-console-" + "f" * 32 + "/console.log")),
            "owner": {"run_id": "f" * 32, "kind": "console"}, "members": [".WinBookSplit-console-owner.json", "console.log"], "sha256": "c" * 64}
    return seal(row)


def extra_case():
    row = cancelled_case()
    kind = "empty-second-third"
    row.update(id="BAT-" + kind, kind=kind, host_id="BAT", exit_code=2, stdin_utf8="", stdin_sha256=sha256(b"").hexdigest(),
        parameters=[row["input_path"], "", "C:/synthetic/b/third.pdf"], command=["C:/Windows/System32/cmd.exe", "/d", "/v:off", "/c", "C:/synthetic/invoke.cmd"],
        modifications=["copied PS launch sentinel only"], powershell_launch_sentinel_absent=True, outcome=None)
    row["actual_application_sha256"]["WinBookSplit.ps1"] = sha256(validator.SENTINEL_SOURCE.encode("utf-8-sig")).hexdigest()
    row["wrapper_text"] = '@echo off\r\nsetlocal DisableDelayedExpansion\r\n"%WBS_LAUNCHER_BAT%" "%WBS_LAUNCHER_INPUT1%" "%WBS_LAUNCHER_INPUT2%" "%WBS_LAUNCHER_INPUT3%"\r\n'
    row["wrapper_sha256"] = sha256(row["wrapper_text"].encode("ascii")).hexdigest()
    return seal(row)


class LauncherReceiptTests(unittest.TestCase):
    def test_source_cancel_requires_native_status_no_processing_and_immutable_identity(self):
        valid = cancelled_case()
        validator.validate_case(valid)
        defects = [("exit_code", 0), ("pid", True), ("timed_out", True), ("stdin", "DEVNULL"), ("stdin_utf8", "yes\n"),
                   ("stdin_sha256", "c" * 64), ("elapsed_seconds", float("inf")), ("source_read_only_observed", False),
                   ("source_observations_after", {}), ("additional_inputs_after", {}), ("output_members_after", ["new-output"]),
                   ("output_members_after_cleanup", []), ("engine_records", [{}]), ("process_summaries", [{}]), ("log_sha256", "c" * 64),
                   ("command", ["different.exe"]), ("application_sha256", {}), ("modifications", ["edited app"])]
        for key, value in defects:
            row = deepcopy(valid)
            row[key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                validator.validate_case(row)
        for key, value in (("exit_code", 6), ("version", True), ("written_count", 1), ("status", "success"), ("code", "invalid_arguments"), ("engine_result", {})):
            row = deepcopy(valid)
            row["outcome"][key] = value
            seal(row)
            with self.subTest(outcome=key), self.assertRaises(ValueError):
                validator.validate_case(row)

    def test_extra_arguments_require_exact_bat_bytes_count_and_no_ps_launch(self):
        valid = extra_case()
        validator.validate_case(valid)
        for key, value in (("powershell_launch_sentinel_absent", False), ("parameters", [valid["input_path"]]),
                           ("wrapper_sha256", "c" * 64), ("modifications", []), ("outcome", {}), ("engine_records", [{}])):
            row = deepcopy(valid)
            row[key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                validator.validate_case(row)
        for filename in ("WinBookSplit.bat", "WinBookSplit.ps1", "engine/winbooksplit_engine.py"):
            row = deepcopy(valid)
            row["actual_application_sha256"][filename] = "d" * 64
            with self.subTest(file=filename), self.assertRaises(ValueError):
                validator.validate_case(row)

    def test_method_eof_cannot_claim_a_split_or_execution_result(self):
        valid = cancelled_case("menu-eof")
        validator.validate_case(valid)
        for key, value in (("engine_records", [{}]), ("process_summaries", [{}]), ("final_publication", [{}])):
            row = deepcopy(valid)
            row[key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                validator.validate_case(row)
        row = deepcopy(valid)
        row["outcome"]["engine_result"] = {"status": "success", "written_count": 3}
        seal(row)
        with self.assertRaises(ValueError):
            validator.validate_case(row)

    def test_stream_receipt_must_have_one_exact_authoritative_outcome(self):
        valid = cancelled_case()
        for text in (valid["stdout"] * 2, "prompt but no final record\n", valid["stdout"] + "Done.\n",
                     valid["stdout"] + "[!] Error: selection cancelled\n", valid["stdout"] + "[!] Processor failure: cancelled\n"):
            row = deepcopy(valid)
            row["stdout"] = text
            row["stdout_sha256"] = sha256(text.encode()).hexdigest()
            with self.assertRaises(ValueError):
                validator.validate_case(row)

    def test_launcher_runner_fails_closed_and_keeps_failed_retained_workspace_receipt(self):
        with tempfile.TemporaryDirectory(prefix="wbs-launcher-unit-") as directory:
            path = Path(directory) / "child.json"
            payloads = [{"schema_version": 1, "success": True, "case_count": 70},
                        {"schema_version": 1, "task_id": "M3-T03", "success": False, "cleanup_safe": False, "retained_workspace": "C:/authored/work"}]
            for payload in payloads:
                path.write_text(json.dumps(payload), encoding="utf-8")
                step = {"exit_code": 0}
                runner.attach_child_report(step, path, "launcher")
                self.assertEqual(step["exit_code"], 126)
                self.assertNotIn("evidence", step)
                self.assertEqual(step["failed_evidence"], payload)
                if payload.get("cleanup_safe") is False:
                    self.assertIs(step["cleanup_safe"], False)
                step = {"exit_code": 19}
                runner.attach_child_report(step, path, "launcher")
                self.assertEqual(step["exit_code"], 19)
            absent = Path(directory) / "absent.json"
            step = {"exit_code": 0}
            runner.attach_child_report(step, absent, "launcher")
            self.assertEqual(step["exit_code"], 126)


if __name__ == "__main__":
    unittest.main(verbosity=2)
