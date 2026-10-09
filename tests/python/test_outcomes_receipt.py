"""Synthetic strict-receipt guards; these tests do not certify Windows behavior."""

from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("wbs_unit_outcomes_validator", ROOT / "tests/outcomes/validate_outcomes_report.py")
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)
DIGEST = "a" * 64


def fixture():
    """A transparent invented receipt for validator tests, never runtime evidence."""
    root = Path(tempfile.gettempdir()).resolve() / "synthetic-outcome-receipt"
    shells = [str(root / "powershell.exe"), str(root / "pwsh.exe")]
    versions = {"PS51": "5.1.26100.9444", "PS7": "7.6.5"}
    hosts = [{"id": host, "shell_executable": shells[index], "host_version": versions[host],
        "stored_policies": [], "policies_after": [], "syntax_error_count": 0}
        for index, host in enumerate(("PS51", "PS7"))]
    source_hashes = {name: DIGEST for name in validator.APPLICATION}
    cases = []
    for prefix in ("PS51", "PS7", "BAT"):
        for kind, exit_code in validator.EXPECTED.items():
            host = hosts[1 if prefix == "PS7" else 0]
            base = root / (prefix + "-" + kind) / "o"
            status, code = validator.OUTCOMES[kind]
            outcome = {"protocol": "winbooksplit.outcome", "version": 1, "status": status, "code": code,
                "message": "Synthetic receipt only", "exit_code": exit_code, "mode": "manual", "written_count": 0,
                "final_directory": None, "engine_result": None}
            engine = {"protocol": "winbooksplit.result", "version": 1, "mode": "manual", "status": status,
                "code": code, "exit_code": exit_code, "written_count": 0, "execution": None, "diagnostic": None}
            publication = content = manifest = manifest_hash = None
            if kind in {"fallback-success", "log-finalize"}:
                publication = [{"filename": f"{page:02d} - Chapter {page}.pdf", "page_ids": [page],
                    "range": [page - 1, page], "sha256": DIGEST, "size_bytes": 100, "page_count": 1} for page in (1, 2, 3)]
                outputs = [{"filename": item["filename"], "sha256": DIGEST, "size_bytes": 100, "page_count": 1} for item in publication]
                manifest = {"status": "complete", "written_count": 3, "outputs": outputs}
                execution = {"final_directory": str(base / ("Book_20000101-000000_" + "0" * 32)),
                    "written_count": 3, "outputs": deepcopy(outputs), "manifest": deepcopy(manifest)}
                engine.update(status="success", code="split_complete", exit_code=0, written_count=3, execution=execution)
                outcome.update(engine_result=engine, final_directory=execution["final_directory"], written_count=3 if kind == "fallback-success" else 0)
                content, manifest_hash = [DIGEST] * 3, DIGEST
            elif kind != "startup-dependency" and kind not in {"engine-cancel", "engine-timeout", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}:
                outcome["engine_result"] = engine
            records = [engine] if outcome["engine_result"] else []
            if kind.startswith("fallback-"):
                initial = {"protocol": "winbooksplit.result", "version": 1, "status": "no_plan", "code": "no_bookmarks",
                    "exit_code": 5, "mode": "1", "written_count": 0, "execution": None, "fallback_modes": ["manual"]}
                records = [initial] + ([] if kind == "fallback-cancel" else [engine])
                if kind == "fallback-cancel":
                    outcome["engine_result"] = initial
            if kind == "conversion-timeout":
                engine.update(status="timeout", code="conversion_timeout", exit_code=130)
            transport = {"ParentStopped": True, "DescendantsStopped": True, "StreamsComplete": True,
                "JobAssigned": True, "Pid": 101, "Cancelled": kind in {"engine-cancel", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"},
                "TimedOut": kind == "engine-timeout", "ExitCode": exit_code, "ResultRecordCount": len(records)}
            if kind == "fallback-cancel":
                transport.update(ExitCode=5, ResultRecordCount=1)
            elif kind in {"engine-cancel", "engine-timeout", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}:
                transport.update(ExitCode=1, ResultRecordCount=0)
            elif kind == "log-finalize":
                transport.update(ExitCode=0, ResultRecordCount=1)
            transports = [] if kind == "startup-dependency" else [deepcopy(transport)]
            if kind in {"fallback-success", "fallback-failure"}:
                transports = [deepcopy(transport), deepcopy(transport)]
                transports[0]["ExitCode"] = 5
                transports[0]["ResultRecordCount"] = transports[1]["ResultRecordCount"] = 1
            retained = kind in {"cleanup-refusal", "engine-cancel", "engine-timeout", "conversion-cancel", "engine-ctrlc", "conversion-ctrlc"}
            stage = str(base / (".WinBookSplit-stage-" + "0" * 32)) if retained else None
            fault = {"pid": 101, "stage": stage, "members": {}} if kind in validator.ENGINE_FAULTS or kind in validator.INTERRUPTED else None
            if kind == "cleanup-refusal":
                fault["members"]["authored-unrelated.txt"] = {"sha256": DIGEST, "identity": [1, 2]}
                engine["diagnostic"] = {"cleanup_complete": False, "cleanup_error": "Synthetic foreign member", "retained_staging": stage}
            failure_record = failure_hash = converter_path = converter_hash = None
            if kind in {"output-write", "manifest-finalize", "publish-rename", "cleanup-refusal", "converter-failure", "conversion-timeout"}:
                engine["diagnostic"] = engine.get("diagnostic") or {"cleanup_complete": True}
                failure_record = {"code": code, "status": "timeout" if kind == "conversion-timeout" else "failed",
                    "cleanup_complete": kind != "cleanup-refusal"}
                failure_hash = DIGEST
            if kind in {"converter-failure", "conversion-cancel", "conversion-timeout", "conversion-ctrlc"}:
                converter_path, converter_hash = str(root / "OwnedConverter.exe"), DIGEST
                fault = fault or {"pid": 101}
                fault["output"] = str(base / (".WinBookSplit-stage-" + "0" * 32) / "converted.pdf")
                if kind in {"converter-failure", "conversion-timeout"}:
                    engine["diagnostic"]["conversion"] = {"stderr_tail": "AUTHORED_CONVERTER_FINAL_STDERR",
                        "argv": [converter_path, str(root / "source.epub"), fault["output"], "--output-profile", "tablet"],
                        "exit_code": 17 if kind == "converter-failure" else 1}
            identities = {str(root / (name + ".txt")): {"sha256": DIGEST, "size_bytes": 10,
                "device": 1, "inode": index + 1, "attributes": 0} for index, name in enumerate(("source", "neighbor", "prior"))}
            controlled = deepcopy(source_hashes)
            if kind in validator.ENGINE_FAULTS:
                controlled["engine/saved_engine.py"] = DIGEST
            row = {"id": prefix + "-" + kind, "kind": kind, "batch": prefix == "BAT",
                "host_version": host["host_version"], "shell_executable": host["shell_executable"],
                "passed": True, "actual_process": True, "batch_bytes_unchanged": True,
                "source_neighbor_prior_unchanged": True, "unrelated_process_alive": True,
                "final_completion_truthful": True, "owned_application_outputs_removed": True,
                "post_cleanup_source_neighbor_prior_unchanged": True, "exit_code": exit_code, "timed_out": False,
                "outcome": outcome, "stdout": ("Done.\nOutput: " + str(base) if kind == "fallback-success" else "failure")
                    + "\n[OUTCOME] " + json.dumps(outcome) + "\n",
                "engine_records": records, "process_summaries": transports, "application_source_sha256": deepcopy(source_hashes),
                "controlled_application_sha256": controlled, "fault_receipt": fault,
                "final_publication": publication, "final_publication_validated": publication is not None,
                "source_page_content_sha256": content, "output_page_content_sha256": deepcopy(content),
                "publication_manifest": manifest, "publication_manifest_sha256": manifest_hash,
                "failure_diagnostic_record": failure_record, "failure_diagnostic_sha256": failure_hash,
                "converter_path": converter_path, "converter_sha256": converter_hash,
                "log_finalizer_failure_visible": kind == "log-finalize", "log_sha256": None if kind == "startup-dependency" else DIGEST,
                "pid_observations": [{"pid": 101, "running": False}, {"pid": 102, "running": False}] if kind in validator.INTERRUPTED else [],
                "unrelated_pid": 999, "retained_staging": stage, "known_stage_cleanup_complete": not retained,
                "source_observations": identities, "after_operation_source_observations": deepcopy(identities),
                "post_cleanup_source_observations": deepcopy(identities),
                "output_members_before_cleanup": ["prior-output.pdf"], "output_members_after_cleanup": ["prior-output.pdf"],
                "command": [host["shell_executable"], "authored.ps1"], "cwd": str(base.parent), "output_base": str(base)}
            if kind.endswith("ctrlc"):
                row["console_signal"] = {"actual_native_console_signal": True, "hidden_private_console": True,
                    "signal_sent": True, "forced_target_stop": False, "controller_ignore_restored": True}
            cases.append(row)
    report = {"schema_version": 1, "task_id": "M3-T01", "result": "OUTCOME_REGRESSION_PASSED", "success": True,
        "exit_code": 0, "acceptance_ids": ["AC-050", "AC-051", "AC-052"], "source_unchanged": True,
        "machine_settings_unchanged": True, "baseline_guards_preserved": True, "owned_temp_removed": True,
        "host_cases": hosts, "cases": cases, "case_count": len(cases),
        "tested_path_sha256": {**source_hashes, "tests/outcomes/OwnedConverter.cs": DIGEST,
            "tests/outcomes/characterize_outcomes.py": DIGEST, "tests/outcomes/validate_outcomes_report.py": DIGEST,
            "tests/outcomes/fault_engine.py": DIGEST},
        "converter_compilation": {"exit_code": 0, "source_sha256": DIGEST, "executable_sha256": DIGEST},
        "unrelated_process_observation": {"alive_after_all_operations": True, "running_after_retained_handle_stop": False,
            "retained_pid": 999, "actual_identity": {"pid": 999}}}
    return report, shells


class OutcomeReceiptGuards(unittest.TestCase):
    def setUp(self):
        self.report, self.shells = fixture()

    def reject(self, change):
        candidate = deepcopy(self.report)
        change(candidate)
        with self.assertRaises(ValueError):
            validator.validate_outcomes_report(candidate, self.shells)

    def test_complete_synthetic_receipt_is_only_structural(self):
        validator.validate_outcomes_report(self.report, self.shells)

    def test_missing_duplicate_host_case_and_provenance_are_rejected(self):
        changes = [lambda row: row["cases"].pop(), lambda row: row["cases"].append(deepcopy(row["cases"][0])),
            lambda row: row.update(case_count=0), lambda row: row.update(host_cases=[]),
            lambda row: row.update(tested_path_sha256={}), lambda row: row.update(owned_temp_removed=False),
            lambda row: row["host_cases"][0].update(policies_after=[{"changed": True}]),
            lambda row: row["cases"][0]["application_source_sha256"].update({"WinBookSplit.ps1": "b" * 64}),
            lambda row: row["cases"][0]["controlled_application_sha256"].update({"engine/WinBookSplit.Process.ps1": "b" * 64}),
            lambda row: row["converter_compilation"].update(source_sha256="b" * 64)]
        for index, change in enumerate(changes):
            with self.subTest(index=index):
                self.reject(change)

    def test_outcome_protocol_code_status_native_and_cleanup_contradictions_are_rejected(self):
        mutations = [("outcome", "protocol", "wrong"), ("outcome", "version", True),
            ("outcome", "code", "split_complete"), ("outcome", "status", "success"),
            ("outcome", "exit_code", 0), ("outcome", "written_count", 1),
            (None, "exit_code", 0), (None, "timed_out", True), (None, "unrelated_process_alive", False),
            (None, "owned_application_outputs_removed", False), (None, "output_members_after_cleanup", ["unrelated"]),
            (None, "source_neighbor_prior_unchanged", False), (None, "source_observations", {}),
            (None, "post_cleanup_source_observations", {}), (None, "after_operation_source_observations", {}),
            (None, "log_sha256", None), (None, "process_summaries", []),
            (None, "final_publication", []), (None, "known_stage_cleanup_complete", False)]
        for group, key, value in mutations:
            with self.subTest(group=group, key=key):
                self.reject(lambda report, g=group, k=key, v=value: (report["cases"][0][g] if g else report["cases"][0]).update({k: v}))

    def test_publication_fallback_interrupt_and_survival_contradictions_are_rejected(self):
        def selected(report, kind):
            return next(row for row in report["cases"] if row["id"] == "PS51-" + kind)
        changes = [lambda row: selected(row, "fallback-success")["final_publication"][1].update(page_ids=[1]),
            lambda row: selected(row, "log-finalize")["final_publication"][0].update(page_count=0),
            lambda row: selected(row, "log-finalize")["final_publication"][0].update(sha256="b" * 64),
            lambda row: selected(row, "log-finalize").update(output_page_content_sha256=["b" * 64] * 3),
            lambda row: selected(row, "log-finalize").update(publication_manifest={}),
            lambda row: selected(row, "log-finalize").update(publication_manifest_sha256=None),
            lambda row: selected(row, "fallback-success")["engine_records"].pop(0),
            lambda row: selected(row, "fallback-failure")["outcome"].update(engine_result={"status": "no_plan"}),
            lambda row: selected(row, "engine-cancel").update(pid_observations=[]),
            lambda row: selected(row, "engine-cancel")["pid_observations"][1].update(running=True),
            lambda row: selected(row, "engine-cancel")["process_summaries"][0].update(ParentStopped=False),
            lambda row: selected(row, "engine-timeout")["process_summaries"][0].update(TimedOut=False),
            lambda row: selected(row, "engine-ctrlc")["console_signal"].update(signal_sent=False),
            lambda row: selected(row, "conversion-ctrlc")["console_signal"].update(forced_target_stop=True),
            lambda row: selected(row, "cleanup-refusal")["outcome"]["engine_result"]["diagnostic"].update(cleanup_complete=True),
            lambda row: selected(row, "cleanup-refusal").update(retained_staging=None),
            lambda row: selected(row, "output-write")["controlled_application_sha256"].update({"engine/saved_engine.py": "b" * 64}),
            lambda row: selected(row, "output-write")["controlled_application_sha256"].update({"engine/winbooksplit_engine.py": "b" * 64}),
            lambda row: row["tested_path_sha256"].pop("tests/outcomes/fault_engine.py"),
            lambda row: row["tested_path_sha256"].pop("tests/outcomes/characterize_outcomes.py"),
            lambda row: row["unrelated_process_observation"].update(alive_after_all_operations=False),
            lambda row: row["unrelated_process_observation"].update(running_after_retained_handle_stop=True),
            lambda row: row["unrelated_process_observation"]["actual_identity"].update(pid=123)]
        for index, change in enumerate(changes):
            with self.subTest(index=index):
                self.reject(change)

    def test_outer_workspace_honors_unsafe_child_retention(self):
        runner_spec = importlib.util.spec_from_file_location("wbs_unit_outcomes_runner", ROOT / "tests/run_tests.py")
        runner = importlib.util.module_from_spec(runner_spec)
        runner_spec.loader.exec_module(runner)
        steps = []
        with runner.owned_evidence_workspace(steps) as directory:
            owned = Path(directory).resolve()
            sentinel = owned / "authored-retained-marker.txt"
            sentinel.write_bytes(b"retain when owned tree stop is unknown")
            steps.append({"cleanup_safe": False})
        try:
            self.assertTrue(owned.is_dir())
            self.assertEqual(sentinel.read_bytes(), b"retain when owned tree stop is unknown")
            self.assertEqual({path.name for path in owned.iterdir()}, {sentinel.name})
        finally:
            # This test alone authored the exact ordinary file and empty parent.
            self.assertEqual(sentinel.parent, owned)
            sentinel.unlink()
            owned.rmdir()

    def test_nonfallback_logged_frames_and_native_binding_are_required(self):
        ordinary = ("original-failure", "converter-failure", "output-write", "manifest-finalize",
            "publish-rename", "cleanup-refusal", "conversion-timeout", "log-finalize")
        for kind in ordinary:
            for mutation in ("missing", "duplicate", "contradicting", "native-exit", "native-count"):
                def change(report, selected_kind=kind, change_kind=mutation):
                    case = next(value for value in report["cases"] if value["id"] == "PS51-" + selected_kind)
                    if change_kind == "missing":
                        case["engine_records"] = []
                    elif change_kind == "duplicate":
                        case["engine_records"] *= 2
                    elif change_kind == "contradicting":
                        case["engine_records"] = [{"status": "success", "exit_code": 0}]
                    elif change_kind == "native-exit":
                        case["process_summaries"][0]["ExitCode"] = 123
                    else:
                        case["process_summaries"][0]["ResultRecordCount"] = 0
                with self.subTest(kind=kind, mutation=mutation):
                    self.reject(change)

    def test_stdout_fallback_readiness_and_retained_path_bindings_are_required(self):
        def selected(report, kind):
            return next(row for row in report["cases"] if row["id"] == "PS51-" + kind)
        changes = [lambda row: selected(row, "original-failure").update(stdout="failure without final record"),
            lambda row: selected(row, "original-failure").update(stdout="[OUTCOME] malformed json"),
            lambda row: selected(row, "original-failure").update(stdout="[OUTCOME] {}"),
            lambda row: selected(row, "original-failure").update(stdout=selected(row, "original-failure")["stdout"] * 2),
            lambda row: selected(row, "engine-cancel")["fault_receipt"].update(pid=123),
            lambda row: selected(row, "conversion-timeout")["fault_receipt"].update(pid=123),
            lambda row: selected(row, "engine-cancel").update(retained_staging=str(ROOT / (".WinBookSplit-stage-" + "0" * 32))),
            lambda row: selected(row, "engine-cancel").update(retained_staging=str(Path(selected(row, "engine-cancel")["output_base"]) / "foreign-stage")),
            lambda row: selected(row, "engine-cancel")["process_summaries"][0].update(ExitCode=130)]
        for index, change in enumerate(changes):
            with self.subTest(index=index):
                self.reject(change)
        for field, value in (("protocol", "wrong"), ("version", True), ("mode", "manual"),
                ("status", "success"), ("code", "wrong"), ("exit_code", 1), ("written_count", 1),
                ("execution", {}), ("fallback_modes", [])):
            def change(report, key=field, replacement=value):
                # Preserve stdout equality so this checks the independent initial fallback contract.
                case = selected(report, "fallback-cancel")
                case["engine_records"][0][key] = replacement
                case["stdout"] = "failure\n[OUTCOME] " + json.dumps(case["outcome"]) + "\n"
            with self.subTest(fallback_field=field):
                self.reject(change)

    def test_ordinary_outcome_inherits_last_engine_semantics(self):
        ordinary = ("original-failure", "converter-failure", "output-write", "manifest-finalize",
            "publish-rename", "cleanup-refusal", "conversion-timeout")
        for kind in ordinary:
            for field, value in (("status", "invented"), ("code", "invented"), ("exit_code", 17), ("mode", "1")):
                def change(report, selected_kind=kind, key=field, replacement=value):
                    case = next(row for row in report["cases"] if row["id"] == "PS51-" + selected_kind)
                    # Keep frame/native/stdout aligned to isolate the final-attempt semantic binding.
                    case["outcome"]["engine_result"][key] = replacement
                    case["process_summaries"][0]["ExitCode"] = case["outcome"]["engine_result"]["exit_code"]
                    case["stdout"] = "failure\n[OUTCOME] " + json.dumps(case["outcome"]) + "\n"
                with self.subTest(kind=kind, field=field):
                    self.reject(change)

    def test_final_fallback_frame_semantics_match_outcome(self):
        for field, value in (("status", "invented"), ("code", "invalid_mode"), ("exit_code", 17), ("mode", "1")):
            def change(report, key=field, replacement=value):
                case = next(row for row in report["cases"] if row["id"] == "PS51-fallback-failure")
                case["engine_records"][1][key] = replacement
                case["process_summaries"][1]["ExitCode"] = case["engine_records"][1]["exit_code"]
                case["stdout"] = "failure\n[OUTCOME] " + json.dumps(case["outcome"]) + "\n"
            with self.subTest(field=field):
                self.reject(change)


if __name__ == "__main__":
    unittest.main()
