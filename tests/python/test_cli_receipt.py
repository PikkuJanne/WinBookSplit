"""Synthetic strict-receipt units only; these never claim actual Windows passes."""

from copy import deepcopy
from hashlib import sha256
import importlib.util
import json
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_cli_unit_validator", ROOT / "tests/cli/validate_cli_report.py")
validator = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(validator)


def seal_stdout(case, text=None):
    if text is None:
        text = "[OUTCOME] " + json.dumps(case["outcome"]) + "\n"
    case["stdout"] = text
    case["stdout_sha256"] = sha256(text.encode()).hexdigest()
    return case


def synthetic_case(kind="bad-mode"):
    observed = {name: {"sha256": "a" * 64, "size_bytes": 10, "device": 1, "inode": number, "attributes": 1}
                for number, name in enumerate(("source", "neighbor", "prior"), 1)}
    app = {name: "b" * 64 for name in validator.APPLICATION}
    status = {"protocol": "winbooksplit.outcome", "version": 1, "status": "failed", "code": "invalid_arguments",
        "exit_code": 2, "written_count": 0, "final_directory": None, "engine_result": None}
    case = {"id": "PS51-" + kind, "kind": kind, "host_version": "5.1.26100.9444", "shell_executable": "C:/synthetic/powershell.exe",
        "command": ["C:/synthetic/powershell.exe", "-NoProfile", "-NonInteractive", "-File", "C:/synthetic/a/WinBookSplit.ps1", "-Mode", "Unknown"],
        "parameters": ["-Mode", "Unknown"], "passed": True, "actual_process": True, "source_read_only_observed": True,
        "source_neighbor_prior_unchanged": True, "application_unchanged": True, "known_output_cleanup_complete": True,
        "exit_code": 2, "timed_out": False, "stdin": "DEVNULL", "pid": 1234, "elapsed_seconds": .1, "timeout_seconds": 90,
        "stderr": "", "stderr_sha256": sha256(b"").hexdigest(), "application_sha256": app, "application_members": sorted(app),
        "output_members_before": ["prior-output.pdf"], "output_members_after": ["prior-output.pdf"], "output_members_after_cleanup": ["prior-output.pdf"],
        "source_observations_before": deepcopy(observed), "source_observations_after": deepcopy(observed), "source_observations_after_cleanup": deepcopy(observed),
        "outcome": status, "engine_records": [], "process_summaries": [], "log_sha256": None, "output_base": "C:/synthetic/o"}
    return seal_stdout(case)


def synthetic_preview():
    case = synthetic_case("preview-auto1")
    entries = [{"sequence": number, "start": start, "end": end, "title": "Title " + str(number), "filename": f"{number:02d} - Title.pdf"}
               for number, (start, end) in enumerate(validator.RANGES["1"], 1)]
    plan = {"mode": "1", "total_pages": 6, "entries": entries, "ranges": deepcopy(validator.RANGES["1"]),
            "coverage": {"complete": True, "covered_pages": 6, "section_count": 3}, "output_naming": {"resolved_base": case["output_base"]}}
    frame = {"protocol": "winbooksplit.result", "version": 1, "mode": "1", "status": "preview", "code": "preview_complete",
             "exit_code": 0, "written_count": 0, "execution": None, "fallback_modes": [], "plan": plan}
    case.update(exit_code=0, plan=deepcopy(plan), expected_entries=deepcopy(entries), expected_ranges=deepcopy(validator.RANGES["1"]))
    case["outcome"].update(status="preview", code="preview_complete", exit_code=0, mode="1", engine_result=frame)
    return seal_stdout(case, "[OUTCOME] " + json.dumps(case["outcome"]) + "\nPreview only: no chapter PDFs were written.\n")


class CliReceiptTests(unittest.TestCase):
    def test_early_native_refusal_requires_concrete_no_processing_no_write_and_identity_proofs(self):
        valid = synthetic_case()
        validator.validate_case(valid)
        mutations = [("exit_code", 0), ("timed_out", True), ("stdin", "PIPE"), ("pid", True),
            ("elapsed_seconds", float("inf")), ("command", ["other.exe"]), ("parameters", []),
            ("source_read_only_observed", False), ("known_output_cleanup_complete", False), ("application_unchanged", False),
            ("application_sha256", {}), ("application_members", []), ("output_members_after", ["prior-output.pdf", "new"]),
            ("source_observations_after", {}), ("engine_records", [{}]), ("process_summaries", [{}]), ("log_sha256", "a" * 64)]
        for field, value in mutations:
            with self.subTest(field=field):
                case = deepcopy(valid)
                case[field] = value
                with self.assertRaises(ValueError):
                    validator.validate_case(case)
        for field, value in (("exit_code", 6), ("written_count", 1), ("engine_result", {}), ("code", "runtime_invalid"), ("version", True)):
            case = deepcopy(valid)
            case["outcome"][field] = value
            seal_stdout(case)
            with self.subTest(outcome=field), self.assertRaises(ValueError):
                validator.validate_case(case)
        for text in (valid["stdout"] * 2, valid["stdout"] + "[DEPENDENCY-ERROR] {}\n", valid["stdout"] + "Read-Host\n"):
            case = deepcopy(valid)
            seal_stdout(case, text)
            with self.assertRaises(ValueError):
                validator.validate_case(case)

    def test_preview_requires_zero_execution_bound_plan_and_exact_complete_ranges(self):
        valid = synthetic_preview()
        validator.validate_case(valid)
        for field, value in (("status", "success"), ("code", "split_complete"), ("written_count", 3), ("execution", {}), ("fallback_modes", ["manual"])):
            case = deepcopy(valid)
            case["outcome"]["engine_result"][field] = value
            seal_stdout(case, "[OUTCOME] " + json.dumps(case["outcome"]) + "\nPreview only:\n")
            with self.subTest(engine=field), self.assertRaises(ValueError):
                validator.validate_case(case)
        for change in (lambda case: case["plan"]["entries"][1].update(start=1),
                lambda case: case["plan"]["coverage"].update(covered_pages=5),
                lambda case: case["plan"]["output_naming"].update(resolved_base="C:/other"),
                lambda case: case.update(output_members_after=["prior-output.pdf", "chapter.pdf"]),
                lambda case: case.update(expected_ranges=[[0, 6]])):
            case = deepcopy(valid)
            change(case)
            case["outcome"]["engine_result"]["plan"] = deepcopy(case["plan"])
            seal_stdout(case, "[OUTCOME] " + json.dumps(case["outcome"]) + "\nPreview only:\n")
            with self.assertRaises(ValueError):
                validator.validate_case(case)

    def test_version_requires_helper_free_application_and_exact_canonical_output(self):
        case = synthetic_case("version")
        app = {name: "b" * 64 for name in ("WinBookSplit.ps1", "VERSION")}
        case.update(application_sha256=app, application_members=sorted(app), outcome=None, exit_code=0, parameters=["-Version"])
        case["command"] = case["command"][:5] + ["-Version"]
        seal_stdout(case, "WinBookSplit 1.0.0\n")
        validator.validate_case(case)
        for field, value in (("outcome", {}), ("log_sha256", "a" * 64), ("application_sha256", {name: "b" * 64 for name in validator.APPLICATION})):
            changed = deepcopy(case)
            changed[field] = value
            with self.subTest(field=field), self.assertRaises(ValueError):
                validator.validate_case(changed)
        changed = deepcopy(case)
        seal_stdout(changed, "WinBookSplit v2.1\n")
        with self.assertRaises(ValueError):
            validator.validate_case(changed)

    def test_case_sets_reject_missing_duplicate_extra_receipts(self):
        rows = [{"id": "one"}, {"id": "two"}]
        validator.exact_cases(rows, {"one", "two"})
        for changed in (None, rows[:1], [rows[0], rows[0]], [*rows, {"id": "three"}]):
            with self.assertRaises(ValueError):
                validator.exact_cases(changed, {"one", "two"})

    def test_top_level_cannot_pass_without_promised_actual_source_host_and_case_receipts(self):
        report = {"schema_version": 1, "task_id": "M3-T02", "result": "CLI_REGRESSION_PASSED", "success": True,
            "exit_code": 0, "acceptance_ids": validator.ACCEPTANCE, "source_unchanged": True, "baseline_guards_preserved": True,
            "machine_settings_unchanged": True, "owned_temp_removed": True, "tested_path_sha256": {"synthetic.py": "a" * 64}}
        with self.assertRaisesRegex(ValueError, "Promised CLI"):
            validator.validate_cli_report(report, ["C:/synthetic/ps51.exe", "C:/synthetic/pwsh.exe"], "C:/synthetic/ebook-convert.exe")


if __name__ == "__main__":
    unittest.main()
